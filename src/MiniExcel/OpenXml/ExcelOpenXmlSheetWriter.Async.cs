using System.ComponentModel;
using MiniExcelLibs.OpenXml.Constants;
using MiniExcelLibs.OpenXml.Models;
using MiniExcelLibs.OpenXml.Styles;
using MiniExcelLibs.Utils;
using MiniExcelLibs.WriteAdapter;
using MiniExcelLibs.Zip;
using System.IO.Compression;
using System.Text;
using System.Xml.Linq;

namespace MiniExcelLibs.OpenXml;

internal partial class ExcelOpenXmlSheetWriter
{
    public async Task<int[]> SaveAsAsync(CancellationToken cancellationToken = default)
    {
        try
        {
            await CreateZipEntryAsync(ExcelFileNames.Rels, ExcelContentTypes.Relationships, ExcelXml.DefaultRels, cancellationToken);
            await CreateZipEntryAsync(ExcelFileNames.SharedStrings, ExcelContentTypes.SharedStrings, ExcelXml.DefaultSharedString, cancellationToken);

            using var sbc = _sheetStyleBuildContext;
            var styleBuilder = await GetSheetStyleBuilderAsync(cancellationToken);

            var sheets = GetSheets();
            var rowsWritten = new List<int>();

            foreach (var sheet in sheets)
            {
                cancellationToken.ThrowIfCancellationRequested();
                    
                _sheets.Add(sheet.Item1); //TODO:remove
                _currentSheetIndex = sheet.Item1.SheetIdx;
                var rows = await CreateSheetXmlAsync(sheet.Item2, sheet.Item1.Path, cancellationToken);
                rowsWritten.Add(rows);
            }

            await styleBuilder.BuildAsync(cancellationToken);
            await GenerateEndXmlAsync(cancellationToken);
            return rowsWritten.ToArray();
        }
        finally
        {
            _archive.Dispose();
        }
    }

    public async Task<int> InsertAsync(bool overwriteSheet = false, CancellationToken cancellationToken = default)
    {
        try
        {
            if (!_configuration.FastMode)
                throw new InvalidOperationException("Insert requires fast mode to be enabled");

            var sheetRecords = new ExcelOpenXmlSheetReader(_stream, _configuration).GetWorkbookRels(_archive.Entries).ToArray();
            foreach (var sheetRecord in sheetRecords.OrderBy(o => o.Id))
            {
                cancellationToken.ThrowIfCancellationRequested();
                _sheets.Add(new SheetDto { Name = sheetRecord.Name, SheetIdx = (int)sheetRecord.Id, State = sheetRecord.State });
            }
            var existSheetDto = _sheets.SingleOrDefault(s => s.Name == _defaultSheetName);
            if (existSheetDto != null && !overwriteSheet)
                throw new Exception($"Sheet “{_defaultSheetName}” already exist");

            var styleBuilder = await GetSheetStyleBuilderAsync(cancellationToken).ConfigureAwait(false);

            int rowsWritten;
            if (existSheetDto == null)
            {
                _currentSheetIndex = (int)sheetRecords.Max(m => m.Id) + 1;
                var insertSheetInfo = GetSheetInfos(_defaultSheetName);
                var insertSheetDto = insertSheetInfo.ToDto(_currentSheetIndex);
                _sheets.Add(insertSheetDto);
                rowsWritten = await CreateSheetXmlAsync(_value, insertSheetDto.Path, cancellationToken);
            }
            else
            {
                _currentSheetIndex = existSheetDto.SheetIdx;
                _archive.Entries.Single(s => s.FullName == existSheetDto.Path).Delete();
                rowsWritten = await CreateSheetXmlAsync(_value, existSheetDto.Path, cancellationToken);
            }

            await styleBuilder.BuildAsync(cancellationToken);
            await AddFilesToZipAsync(cancellationToken);

            _archive.Entries.SingleOrDefault(s => s.FullName == ExcelFileNames.DrawingRels(_currentSheetIndex - 1))?.Delete();
            await GenerateDrawinRelXmlAsync(_currentSheetIndex - 1, cancellationToken);

            _archive.Entries.SingleOrDefault(s => s.FullName == ExcelFileNames.Drawing(_currentSheetIndex - 1))?.Delete();
            await GenerateDrawingXmlAsync(_currentSheetIndex - 1, cancellationToken);

            GenerateWorkBookXmls(out var workbookXml, out var workbookRelsXml, out var sheetsRelsXml);
            foreach (var sheetRelsXml in sheetsRelsXml)
            {
                var sheetRelsXmlPath = ExcelFileNames.SheetRels(sheetRelsXml.Key);
                _archive.Entries.SingleOrDefault(s => s.FullName == sheetRelsXmlPath)?.Delete();
                await CreateZipEntryAsync(sheetRelsXmlPath, null, ExcelXml.DefaultSheetRelXml.Replace("{{format}}", sheetRelsXml.Value), cancellationToken);
            }

            _archive.Entries.SingleOrDefault(s => s.FullName == ExcelFileNames.Workbook)?.Delete();
            await CreateZipEntryAsync(ExcelFileNames.Workbook, ExcelContentTypes.Workbook, ExcelXml.DefaultWorkbookXml.Replace("{{sheets}}", workbookXml.ToString()), cancellationToken);

            _archive.Entries.SingleOrDefault(s => s.FullName == ExcelFileNames.WorkbookRels)?.Delete();
            await CreateZipEntryAsync(ExcelFileNames.WorkbookRels, null, ExcelXml.DefaultWorkbookXmlRels.Replace("{{sheets}}", workbookRelsXml.ToString()), cancellationToken);

            await InsertContentTypesXmlAsync(cancellationToken);

            return rowsWritten;
        }
        finally
        {
#if NET10_0_OR_GREATER
            await _archive.DisposeAsync();
#else
            _archive.Dispose();
#endif
            _sheetStyleBuildContext.Dispose();
        }
    }

    private async Task<int> CreateSheetXmlAsync(object values, string sheetPath, CancellationToken cancellationToken)
    {
        cancellationToken.ThrowIfCancellationRequested();
            
        var entry = _archive.CreateEntry(sheetPath, CompressionLevel.Fastest);
        var rowsWritten = 0;
            
        using (var zipStream = entry.Open())
        using (var writer = new MiniExcelAsyncStreamWriter(zipStream, Utf8WithBom, _configuration.BufferSize, cancellationToken))
        {
            if (values == null)
            {
                await WriteEmptySheetAsync(writer);
            }
            else
            {
                rowsWritten = await WriteValuesAsync(writer, values, cancellationToken);
            }
        }
        _zipDictionary.Add(sheetPath, new ZipPackageInfo(entry, ExcelContentTypes.Worksheet));
        return rowsWritten;
    }

    private static async Task WriteEmptySheetAsync(MiniExcelAsyncStreamWriter writer)
    {
        await writer.WriteAsync(ExcelXml.EmptySheetXml);
    }

    private static async Task<long> WriteDimensionPlaceholderAsync(MiniExcelAsyncStreamWriter writer)
    {
        var dimensionPlaceholderPostition = await writer.WriteAndFlushAsync(WorksheetXml.StartDimension);
        await writer.WriteAsync(WorksheetXml.DimensionPlaceholder); // end of code will be replaced

        return dimensionPlaceholderPostition;
    }

    private static async Task WriteDimensionAsync(MiniExcelAsyncStreamWriter writer, int maxRowIndex, int maxColumnIndex, long placeholderPosition)
    {
        // Flush and save position so that we can get back again.
        var position = await writer.FlushAsync();

        writer.SetPosition(placeholderPosition);
        await writer.WriteAndFlushAsync($@"{GetDimensionRef(maxRowIndex, maxColumnIndex)}""");

        writer.SetPosition(position);
    }

    private async Task<int> WriteValuesAsync(MiniExcelAsyncStreamWriter writer, object values, CancellationToken cancellationToken)
    {
        cancellationToken.ThrowIfCancellationRequested();

        IMiniExcelWriteAdapter writeAdapter = null;
#if NETSTANDARD2_0_OR_GREATER || NET
        IAsyncMiniExcelWriteAdapter asyncWriteAdapter = null;
#endif
        try
        {
#if NETSTANDARD2_0_OR_GREATER || NET
            if (!MiniExcelWriteAdapterFactory.TryGetAsyncWriteAdapter(values, _configuration, out asyncWriteAdapter))
            {
                writeAdapter = MiniExcelWriteAdapterFactory.GetWriteAdapter(values, _configuration);
            }

            var count = 0;
            var isKnownCount = writeAdapter != null && writeAdapter.TryGetKnownCount(out count);
            var props = writeAdapter != null
                ? writeAdapter?.GetColumns()
                : await asyncWriteAdapter.GetColumnsAsync();
#else
            writeAdapter = MiniExcelWriteAdapterFactory.GetWriteAdapter(values, _configuration);

            var isKnownCount = writeAdapter.TryGetKnownCount(out var count);
            var props = writeAdapter.GetColumns();
#endif

            if (props is null or { Count: 0 })
            {
                await WriteEmptySheetAsync(writer);
                return 0;
            }

            _sheetStyleBuildContext.UpdateFormatIds(props);

            int maxRowIndex;
            var maxColumnIndex = props.Count;
            long dimensionPlaceholderPostition = 0;

            await writer.WriteAsync(WorksheetXml.StartWorksheetWithRelationship);

            // We can write the dimensions directly if the row count is known
            if (_configuration.FastMode && !isKnownCount)
            {
                dimensionPlaceholderPostition = await WriteDimensionPlaceholderAsync(writer);
            }
            else if (isKnownCount)
            {
                maxRowIndex = count + (_printHeader ? 1 : 0);
                await writer.WriteAsync(WorksheetXml.Dimension(GetDimensionRef(maxRowIndex, props.Count)));
            }

            //sheet view
            await writer.WriteAsync(GetSheetViews());

            //cols:width
            ExcelWidthCollection widths = null;
            long columnWidthsPlaceholderPosition = 0;
            if (_configuration.EnableAutoWidth)
            {
                columnWidthsPlaceholderPosition = await WriteColumnWidthPlaceholdersAsync(writer, props);
                widths = ExcelWidthCollection.FromProps(props, _configuration.MinWidth, _configuration.MaxWidth);
            }
            else
            {
                await WriteColumnsWidthsAsync(writer, ExcelWidthCollection.FromProps(props), cancellationToken);
            }

            //header
            await writer.WriteAsync(WorksheetXml.StartSheetData);
            var currentRowIndex = 0;
            if (_printHeader)
            {
                await PrintHeaderAsync(writer, props, cancellationToken);
                currentRowIndex++;
            }

            if (writeAdapter != null)
            {
                foreach (var row in writeAdapter.GetRows(props, cancellationToken))
                {
                    cancellationToken.ThrowIfCancellationRequested();

                    await writer.WriteAsync(WorksheetXml.StartRow(++currentRowIndex));
                    foreach (var cellValue in row)
                    {
                        cancellationToken.ThrowIfCancellationRequested();
                        await WriteCellAsync(writer, currentRowIndex, cellValue.CellIndex, cellValue.Value,
                            cellValue.Prop, widths);
                    }

                    await writer.WriteAsync(WorksheetXml.EndRow);
                }
            }
#if NETSTANDARD2_0_OR_GREATER || NET
                else
                {
                    await foreach (var row in asyncWriteAdapter.GetRowsAsync(props, cancellationToken))
                    {
                        cancellationToken.ThrowIfCancellationRequested();
                        await writer.WriteAsync(WorksheetXml.StartRow(++currentRowIndex));

                        await foreach (var cellValue in row)
                        {
                            cancellationToken.ThrowIfCancellationRequested();
                            await WriteCellAsync(writer, currentRowIndex, cellValue.CellIndex, cellValue.Value,
                                cellValue.Prop, widths);
                        }

                        await writer.WriteAsync(WorksheetXml.EndRow);
                    }
                }
#endif

            maxRowIndex = currentRowIndex;

            await writer.WriteAsync(WorksheetXml.Drawing(_currentSheetIndex));
            await writer.WriteAsync(WorksheetXml.EndSheetData);

            if (_configuration.AutoFilter)
            {
                await writer.WriteAsync(WorksheetXml.Autofilter(GetDimensionRef(maxRowIndex, maxColumnIndex)));
            }

            await writer.WriteAsync(WorksheetXml.EndWorksheet);

            if (_configuration.FastMode && dimensionPlaceholderPostition != 0)
            {
                await WriteDimensionAsync(writer, maxRowIndex, maxColumnIndex, dimensionPlaceholderPostition);
            }

            if (_configuration.EnableAutoWidth)
            {
                await OverWriteColumnWidthPlaceholdersAsync(writer, columnWidthsPlaceholderPosition, widths.Columns,
                    cancellationToken);
            }

            var toSubtract = _printHeader ? 1 : 0;
            return maxRowIndex - toSubtract;
        }
        finally
        {
#if NETSTANDARD2_0_OR_GREATER || NET
                if (asyncWriteAdapter is IAsyncDisposable asyncDisposable)
                {
                    await asyncDisposable.DisposeAsync().ConfigureAwait(false);
                }
#endif
        }
    }

    private static async Task<long> WriteColumnWidthPlaceholdersAsync(MiniExcelAsyncStreamWriter writer, ICollection<ExcelColumnInfo> props)
    {
        var placeholderPosition = await writer.FlushAsync();
        await writer.WriteWhitespaceAsync(WorksheetXml.GetColumnPlaceholderLength(props.Count));
        return placeholderPosition;
    }

    private static async Task OverWriteColumnWidthPlaceholdersAsync(MiniExcelAsyncStreamWriter writer, long placeholderPosition, IEnumerable<ExcelColumnWidth> columnWidths, CancellationToken cancellationToken = default)
    {
        cancellationToken.ThrowIfCancellationRequested();
            
        var position = await writer.FlushAsync();

        writer.SetPosition(placeholderPosition);
        await WriteColumnsWidthsAsync(writer, columnWidths, cancellationToken);

        await writer.FlushAsync();
        writer.SetPosition(position);
    }

    private static async Task WriteColumnsWidthsAsync(MiniExcelAsyncStreamWriter writer, IEnumerable<ExcelColumnWidth> columnWidths, CancellationToken cancellationToken = default)
    {
        var hasWrittenStart = false;
        foreach (var column in columnWidths)
        {
            cancellationToken.ThrowIfCancellationRequested();
                
            if (!hasWrittenStart)
            {
                await writer.WriteAsync(WorksheetXml.StartCols);
                hasWrittenStart = true;
            }
            await writer.WriteAsync(WorksheetXml.Column(column.Index, column.Width, column.Hidden));
        }
            
        if (!hasWrittenStart)
            return;
            
        await writer.WriteAsync(WorksheetXml.EndCols);
    }

    private async Task PrintHeaderAsync(MiniExcelAsyncStreamWriter writer, List<ExcelColumnInfo> props, CancellationToken cancellationToken = default)
    {
        const int yIndex = 1;
        await writer.WriteAsync(WorksheetXml.StartRow(yIndex));

        var xIndex = 1;
        foreach (var p in props)
        {
            //reason : https://github.com/mini-software/MiniExcel/issues/142
            if (p != null)
            {
                if (p.ExcelIgnore)
                    continue;

                var r = ExcelOpenXmlUtils.ConvertXyToCell(xIndex, yIndex);
                await writer.WriteAsync(WorksheetXml.Cell(r, StringDataType, HeaderCellStyleIndex, ExcelOpenXmlUtils.EncodeXML(p.ExcelColumnName))).ConfigureAwait(false);
            }
            xIndex++;
        }

        await writer.WriteAsync(WorksheetXml.EndRow);
    }

    private async Task WriteCellAsync(MiniExcelAsyncStreamWriter writer, int rowIndex, int cellIndex, object value, ExcelColumnInfo p, ExcelWidthCollection widthCollection)
    {
        if (p?.CustomFormatter != null)
        {
            try
            {
                value = p.CustomFormatter(value);
            }
            catch
            {
                //ignored
            }
        }

        var columnReference = ExcelOpenXmlUtils.ConvertXyToCell(cellIndex, rowIndex);
        var valueIsNull = value is null or DBNull;

        if (_configuration.EnableWriteNullValueCell && valueIsNull)
        {
            await writer.WriteAsync(WorksheetXml.EmptyCell(columnReference, DefaultCellStyleIndex));
            return;
        }

        var tuple = GetCellValue(rowIndex, cellIndex, value, p, valueIsNull);

        var styleIndex = tuple.Item1;
        var dataType = tuple.Item2;
        var cellValue = tuple.Item3;
        var columnType = p.ExcelColumnType;

        /*Prefix and suffix blank space will lost after SaveAs #294*/
        var preserveSpace = cellValue != null && (cellValue.StartsWith(" ") || cellValue.EndsWith(" "));

        await writer.WriteAsync(WorksheetXml.Cell(columnReference, dataType, styleIndex, cellValue, preserveSpace: preserveSpace, columnType: columnType));
        widthCollection?.AdjustWidth(cellIndex, cellValue);
    }

    private async Task GenerateEndXmlAsync(CancellationToken cancellationToken)
    {
        await AddFilesToZipAsync(cancellationToken);
        await GenerateDrawinRelXmlAsync(cancellationToken);
        await GenerateDrawingXmlAsync(cancellationToken);
        await GenerateWorkbookXmlAsync(cancellationToken);
        await GenerateContentTypesXmlAsync(cancellationToken);
    }

    private async Task AddFilesToZipAsync(CancellationToken cancellationToken)
    {
        foreach (var item in _files)
        {
            cancellationToken.ThrowIfCancellationRequested();
            await CreateZipEntryAsync(item.Path, item.Byte, cancellationToken);
        }
    }

    private async Task<ISheetStyleBuilder> GetSheetStyleBuilderAsync(CancellationToken cancellationToken)
    {
        SheetStyleBuilderBase builder = _configuration.TableStyles switch
        {
            TableStyles.None => new MinimalSheetStyleBuilder(_sheetStyleBuildContext),
            TableStyles.Default => new DefaultSheetStyleBuilder(_sheetStyleBuildContext, _configuration.StyleOptions),
            _ => throw new InvalidEnumArgumentException(nameof(_configuration.TableStyles), (int)_configuration.TableStyles, typeof(TableStyles))
        };

        var newInfos = builder.GetGenerateElementInfos();
        await _sheetStyleBuildContext.CreateAsync(newInfos, cancellationToken).ConfigureAwait(false);

        return builder;
    }

    private async Task GenerateDrawinRelXmlAsync(CancellationToken cancellationToken)
    {
        for (int sheetIndex = 0; sheetIndex < _sheets.Count; sheetIndex++)
        {
            cancellationToken.ThrowIfCancellationRequested();
            await GenerateDrawinRelXmlAsync(sheetIndex, cancellationToken);
        }
    }

    private async Task GenerateDrawinRelXmlAsync(int sheetIndex, CancellationToken cancellationToken)
    {
        var drawing = GetDrawingRelationshipXml(sheetIndex);
        await CreateZipEntryAsync(
            ExcelFileNames.DrawingRels(sheetIndex),
            string.Empty,
            ExcelXml.DefaultDrawingXmlRels.Replace("{{format}}", drawing),
            cancellationToken);
    }

    private async Task GenerateDrawingXmlAsync(CancellationToken cancellationToken)
    {
        for (int sheetIndex = 0; sheetIndex < _sheets.Count; sheetIndex++)
        {
            cancellationToken.ThrowIfCancellationRequested();
            await GenerateDrawingXmlAsync(sheetIndex, cancellationToken);
        }
    }

    private async Task GenerateDrawingXmlAsync(int sheetIndex, CancellationToken cancellationToken)
    {
        var drawing = GetDrawingXml(sheetIndex);
        await CreateZipEntryAsync(
            ExcelFileNames.Drawing(sheetIndex),
            ExcelContentTypes.Drawing,
            ExcelXml.DefaultDrawing.Replace("{{format}}", drawing),
            cancellationToken);
    }

    private async Task GenerateWorkbookXmlAsync(CancellationToken cancellationToken)
    {
        cancellationToken.ThrowIfCancellationRequested();

        GenerateWorkBookXmls(
            out StringBuilder workbookXml,
            out StringBuilder workbookRelsXml,
            out Dictionary<int, string> sheetsRelsXml);

        foreach (var sheetRelsXml in sheetsRelsXml)
        {
            await CreateZipEntryAsync(
                ExcelFileNames.SheetRels(sheetRelsXml.Key),
                null,
                ExcelXml.DefaultSheetRelXml.Replace("{{format}}", sheetRelsXml.Value),
                cancellationToken);
        }

        await CreateZipEntryAsync(
            ExcelFileNames.Workbook,
            ExcelContentTypes.Workbook,
            ExcelXml.DefaultWorkbookXml.Replace("{{sheets}}", workbookXml.ToString()),
            cancellationToken);

        await CreateZipEntryAsync(
            ExcelFileNames.WorkbookRels,
            null,
            ExcelXml.DefaultWorkbookXmlRels.Replace("{{sheets}}", workbookRelsXml.ToString()),
            cancellationToken);
    }

    private async Task GenerateContentTypesXmlAsync(CancellationToken cancellationToken)
    {
        var contentTypes = GetContentTypesXml();
        await CreateZipEntryAsync(ExcelFileNames.ContentTypes, null, contentTypes, cancellationToken);
    }

    private async Task InsertContentTypesXmlAsync(CancellationToken cancellationToken)
    {
        cancellationToken.ThrowIfCancellationRequested();
            
        var contentTypesZipEntry = _archive.Entries.SingleOrDefault(s => s.FullName == ExcelFileNames.ContentTypes);
        if (contentTypesZipEntry == null)
        {
            await GenerateContentTypesXmlAsync(cancellationToken);
            return;
        }
#if NET5_0_OR_GREATER
            await using (var stream = contentTypesZipEntry.Open())
#else
        using (var stream = contentTypesZipEntry.Open())
#endif
        {
            var doc = XDocument.Load(stream);
            var ns = doc.Root?.GetDefaultNamespace();
            var typesElement = doc.Descendants(ns + "Types").Single();
                
            var partNames = new HashSet<string>(StringComparer.InvariantCultureIgnoreCase);
            foreach (var partName in typesElement.Elements(ns + "Override").Select(s => s.Attribute("PartName").Value))
            {
                partNames.Add(partName);
            }
                
            foreach (var p in _zipDictionary)
            {
                cancellationToken.ThrowIfCancellationRequested();
                    
                var partName = $"/{p.Key}";
                if (!partNames.Contains(partName))
                {
                    var newElement = new XElement(ns + "Override", new XAttribute("ContentType", p.Value.ContentType), new XAttribute("PartName", partName));
                    typesElement.Add(newElement);
                }
            }
                
            stream.Position = 0;
            doc.Save(stream);
        }
    }

    private async Task CreateZipEntryAsync(string path, string contentType, string content, CancellationToken cancellationToken)
    {
        cancellationToken.ThrowIfCancellationRequested();
            
        var entry = _archive.CreateEntry(path, CompressionLevel.Fastest);

#if NET5_0_OR_GREATER
            await using (var zipStream = entry.Open())
#else
        using (var zipStream = entry.Open())
#endif
        using (var writer = new MiniExcelAsyncStreamWriter(zipStream, Utf8WithBom, _configuration.BufferSize, cancellationToken))
            await writer.WriteAsync(content);
            
        if (!string.IsNullOrEmpty(contentType))
            _zipDictionary.Add(path, new ZipPackageInfo(entry, contentType));
    }

    private async Task CreateZipEntryAsync(string path, byte[] content, CancellationToken cancellationToken)
    {
        cancellationToken.ThrowIfCancellationRequested();
            
        var entry = _archive.CreateEntry(path, CompressionLevel.Fastest);
            
#if NET5_0_OR_GREATER
            await using (var zipStream = entry.Open())
#else
        using (var zipStream = entry.Open())
#endif
            await zipStream.WriteAsync(content, 0, content.Length, cancellationToken);
    }
}