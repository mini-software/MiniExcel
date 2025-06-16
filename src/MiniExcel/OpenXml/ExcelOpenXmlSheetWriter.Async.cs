using MiniExcelLibs.OpenXml.Constants;
using MiniExcelLibs.OpenXml.Models;
using MiniExcelLibs.OpenXml.Styles;
using MiniExcelLibs.Utils;
using MiniExcelLibs.WriteAdapter;
using MiniExcelLibs.Zip;
using System;
using System.Collections.Generic;
using System.IO.Compression;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Xml.Linq;

namespace MiniExcelLibs.OpenXml
{
    internal partial class ExcelOpenXmlSheetWriter : IExcelWriter
    {
        [Zomp.SyncMethodGenerator.CreateSyncVersion]
        public async Task<int[]> SaveAsAsync(CancellationToken cancellationToken = default)
        {
            try
            {
                cancellationToken.ThrowIfCancellationRequested();

                await GenerateDefaultOpenXmlAsync(cancellationToken).ConfigureAwait(false);

                var sheets = GetSheets();
                var rowsWritten = new List<int>();

                foreach (var sheet in sheets)
                {
                    cancellationToken.ThrowIfCancellationRequested();

                    _sheets.Add(sheet.Item1); //TODO:remove
                    _currentSheetIndex = sheet.Item1.SheetIdx;
                    var rows = await CreateSheetXmlAsync(sheet.Item2, sheet.Item1.Path, cancellationToken).ConfigureAwait(false);
                    rowsWritten.Add(rows);
                }

                await GenerateEndXmlAsync(cancellationToken).ConfigureAwait(false);
                return rowsWritten.ToArray();
            }
            finally
            {
                _archive.Dispose();
            }
        }

        [Zomp.SyncMethodGenerator.CreateSyncVersion]
        public async Task<int> InsertAsync(bool overwriteSheet = false, CancellationToken cancellationToken = default)
        {
            try
            {
                cancellationToken.ThrowIfCancellationRequested();

                if (!_configuration.FastMode)
                    throw new InvalidOperationException("Insert requires fast mode to be enabled");

                cancellationToken.ThrowIfCancellationRequested();

                using var reader = await ExcelOpenXmlSheetReader.CreateAsync(_stream, _configuration, cancellationToken: cancellationToken).ConfigureAwait(false);
                var sheetRecords = (await reader.GetWorkbookRelsAsync(_archive.Entries, cancellationToken).ConfigureAwait(false)).ToArray();
                foreach (var sheetRecord in sheetRecords.OrderBy(o => o.Id))
                {
                    cancellationToken.ThrowIfCancellationRequested();
                    _sheets.Add(new SheetDto { Name = sheetRecord.Name, SheetIdx = (int)sheetRecord.Id, State = sheetRecord.State });
                }
                var existSheetDto = _sheets.SingleOrDefault(s => s.Name == _defaultSheetName);
                if (existSheetDto != null && !overwriteSheet)
                    throw new Exception($"Sheet “{_defaultSheetName}” already exist");

                await GenerateStylesXmlAsync(cancellationToken).ConfigureAwait(false);//GenerateStylesXml必须在校验overwriteSheet之后，避免不必要的样式更改

                int rowsWritten;
                if (existSheetDto == null)
                {
                    _currentSheetIndex = (int)sheetRecords.Max(m => m.Id) + 1;
                    var insertSheetInfo = GetSheetInfos(_defaultSheetName);
                    var insertSheetDto = insertSheetInfo.ToDto(_currentSheetIndex);
                    _sheets.Add(insertSheetDto);
                    rowsWritten = await CreateSheetXmlAsync(_value, insertSheetDto.Path, cancellationToken).ConfigureAwait(false);
                }
                else
                {
                    _currentSheetIndex = existSheetDto.SheetIdx;
                    _archive.Entries.Single(s => s.FullName == existSheetDto.Path).Delete();
                    rowsWritten = await CreateSheetXmlAsync(_value, existSheetDto.Path, cancellationToken).ConfigureAwait(false);
                }

                await AddFilesToZipAsync(cancellationToken).ConfigureAwait(false);

                _archive.Entries.SingleOrDefault(s => s.FullName == ExcelFileNames.DrawingRels(_currentSheetIndex - 1))?.Delete();
                await GenerateDrawinRelXmlAsync(_currentSheetIndex - 1, cancellationToken).ConfigureAwait(false);

                _archive.Entries.SingleOrDefault(s => s.FullName == ExcelFileNames.Drawing(_currentSheetIndex - 1))?.Delete();
                await GenerateDrawingXmlAsync(_currentSheetIndex - 1, cancellationToken).ConfigureAwait(false);

                GenerateWorkBookXmls(out StringBuilder workbookXml, out StringBuilder workbookRelsXml, out Dictionary<int, string> sheetsRelsXml);
                foreach (var sheetRelsXml in sheetsRelsXml)
                {
                    var sheetRelsXmlPath = ExcelFileNames.SheetRels(sheetRelsXml.Key);
                    _archive.Entries.SingleOrDefault(s => s.FullName == sheetRelsXmlPath)?.Delete();
                    await CreateZipEntryAsync(sheetRelsXmlPath, null, ExcelXml.DefaultSheetRelXml.Replace("{{format}}", sheetRelsXml.Value), cancellationToken).ConfigureAwait(false);
                }

                _archive.Entries.SingleOrDefault(s => s.FullName == ExcelFileNames.Workbook)?.Delete();
                await CreateZipEntryAsync(ExcelFileNames.Workbook, ExcelContentTypes.Workbook, ExcelXml.DefaultWorkbookXml.Replace("{{sheets}}", workbookXml.ToString()), cancellationToken).ConfigureAwait(false);

                _archive.Entries.SingleOrDefault(s => s.FullName == ExcelFileNames.WorkbookRels)?.Delete();
                await CreateZipEntryAsync(ExcelFileNames.WorkbookRels, null, ExcelXml.DefaultWorkbookXmlRels.Replace("{{sheets}}", workbookRelsXml.ToString()), cancellationToken).ConfigureAwait(false);

                await InsertContentTypesXmlAsync(cancellationToken).ConfigureAwait(false);

                return rowsWritten;
            }
            finally
            {
                _archive.Dispose();
            }
        }

        [Zomp.SyncMethodGenerator.CreateSyncVersion]
        internal async Task GenerateDefaultOpenXmlAsync(CancellationToken cancellationToken)
        {
            await CreateZipEntryAsync(ExcelFileNames.Rels, ExcelContentTypes.Relationships, ExcelXml.DefaultRels, cancellationToken).ConfigureAwait(false);
            await CreateZipEntryAsync(ExcelFileNames.SharedStrings, ExcelContentTypes.SharedStrings, ExcelXml.DefaultSharedString, cancellationToken).ConfigureAwait(false);
            await GenerateStylesXmlAsync(cancellationToken).ConfigureAwait(false);
        }

        private async Task<int> CreateSheetXmlAsync(object values, string sheetPath, CancellationToken cancellationToken)
        {
            cancellationToken.ThrowIfCancellationRequested();

            var entry = _archive.CreateEntry(sheetPath, CompressionLevel.Fastest);
            var rowsWritten = 0;

            using (var zipStream = entry.Open())
            using (var writer = new MiniExcelAsyncStreamWriter(zipStream, _utf8WithBom, _configuration.BufferSize, cancellationToken))
            {
                if (values == null)
                {
                    await WriteEmptySheetAsync(writer).ConfigureAwait(false);
                }
                else
                {
                    rowsWritten = await WriteValuesAsync(writer, values, cancellationToken).ConfigureAwait(false);
                }
            }
            _zipDictionary.Add(sheetPath, new ZipPackageInfo(entry, ExcelContentTypes.Worksheet));
            return rowsWritten;
        }

        private static async Task WriteEmptySheetAsync(MiniExcelAsyncStreamWriter writer)
        {
            await writer.WriteAsync(ExcelXml.EmptySheetXml).ConfigureAwait(false);
        }

        [Zomp.SyncMethodGenerator.CreateSyncVersion]
        private static async Task<long> WriteDimensionPlaceholderAsync(
#if SYNC_ONLY
            global::MiniExcelLibs.OpenXml.MiniExcelStreamWriter writer
#else
            MiniExcelAsyncStreamWriter writer
#endif
            )
        {
            var dimensionPlaceholderPostition = await writer.WriteAndFlushAsync(WorksheetXml.StartDimension).ConfigureAwait(false);
            await writer.WriteAsync(WorksheetXml.DimensionPlaceholder).ConfigureAwait(false); // end of code will be replaced

            return dimensionPlaceholderPostition;
        }

        [Zomp.SyncMethodGenerator.CreateSyncVersion]
        private static async Task WriteDimensionAsync(
#if SYNC_ONLY
            global::MiniExcelLibs.OpenXml.MiniExcelStreamWriter writer,
#else
            MiniExcelAsyncStreamWriter writer,
#endif
            int maxRowIndex, int maxColumnIndex, long placeholderPosition)
        {
            // Flush and save position so that we can get back again.
            var position = await writer.FlushAsync().ConfigureAwait(false);

            writer.SetPosition(placeholderPosition);
            await writer.WriteAndFlushAsync($@"{GetDimensionRef(maxRowIndex, maxColumnIndex)}""").ConfigureAwait(false);

            writer.SetPosition(position);
        }

        [Zomp.SyncMethodGenerator.CreateSyncVersion]
        private async Task<int> WriteValuesAsync(
#if SYNC_ONLY
            global::MiniExcelLibs.OpenXml.MiniExcelStreamWriter writer,
#else
            MiniExcelAsyncStreamWriter writer,
#endif
            object values, CancellationToken cancellationToken)
        {
            cancellationToken.ThrowIfCancellationRequested();

            IMiniExcelWriteAdapter writeAdapter = null;
            if (!MiniExcelWriteAdapterFactory.TryGetAsyncWriteAdapter(values, _configuration, out var asyncWriteAdapter))
            {
                writeAdapter = MiniExcelWriteAdapterFactory.GetWriteAdapter(values, _configuration);
            }

            var count = 0;
            var isKnownCount = writeAdapter != null && writeAdapter.TryGetKnownCount(out count);
            List<ExcelColumnInfo> props;
#if SYNC_ONLY
            props = writeAdapter?.GetColumns();
#else
            props = writeAdapter != null ? writeAdapter?.GetColumns() : await asyncWriteAdapter.GetColumnsAsync().ConfigureAwait(false);
#endif
            if (props == null)
            {
                await WriteEmptySheetAsync(writer).ConfigureAwait(false);
                return 0;
            }
            int maxRowIndex;
            var maxColumnIndex = props.Count(x => x != null && !x.ExcelIgnore);

            await writer.WriteAsync(WorksheetXml.StartWorksheetWithRelationship).ConfigureAwait(false);

            long dimensionPlaceholderPostition = 0;

            // We can write the dimensions directly if the row count is known
            if (isKnownCount)
            {
                maxRowIndex = _printHeader ? count + 1 : count;
                await writer.WriteAsync(WorksheetXml.Dimension(GetDimensionRef(maxRowIndex, props.Count))).ConfigureAwait(false);
            }
            else if (_configuration.FastMode)
            {
                dimensionPlaceholderPostition = await WriteDimensionPlaceholderAsync(writer).ConfigureAwait(false);
            }

            //sheet view
            await writer.WriteAsync(GetSheetViews()).ConfigureAwait(false);

            //cols:width
            ExcelWidthCollection widths = null;
            long columnWidthsPlaceholderPosition = 0;
            if (_configuration.EnableAutoWidth)
            {
                columnWidthsPlaceholderPosition = await WriteColumnWidthPlaceholdersAsync(writer, maxColumnIndex, cancellationToken).ConfigureAwait(false);
                widths = new ExcelWidthCollection(_configuration.MinWidth, _configuration.MaxWidth, props);
            }
            else
            {
                await WriteColumnsWidthsAsync(writer, ExcelColumnWidth.FromProps(props), cancellationToken).ConfigureAwait(false);
            }

            //header
            await writer.WriteAsync(WorksheetXml.StartSheetData).ConfigureAwait(false);
            var currentRowIndex = 0;
            if (_printHeader)
            {
                await PrintHeaderAsync(writer, props, cancellationToken).ConfigureAwait(false);
                currentRowIndex++;
            }

            if (writeAdapter != null)
            {
                foreach (var row in writeAdapter.GetRows(props, cancellationToken))
                {
                    cancellationToken.ThrowIfCancellationRequested();

                    await writer.WriteAsync(WorksheetXml.StartRow(++currentRowIndex)).ConfigureAwait(false);
                    foreach (var cellValue in row)
                    {
                        cancellationToken.ThrowIfCancellationRequested();
                        await WriteCellAsync(writer, currentRowIndex, cellValue.CellIndex, cellValue.Value, cellValue.Prop, widths).ConfigureAwait(false);
                    }
                    await writer.WriteAsync(WorksheetXml.EndRow).ConfigureAwait(false);
                }
            }
            else
            {
#if !SYNC_ONLY
                await foreach (var row in asyncWriteAdapter.GetRowsAsync(props, cancellationToken).ConfigureAwait(false))
                {
                    cancellationToken.ThrowIfCancellationRequested();
                    await writer.WriteAsync(WorksheetXml.StartRow(++currentRowIndex)).ConfigureAwait(false);

                    await foreach (var cellValue in row.ConfigureAwait(false))
                    {
                        cancellationToken.ThrowIfCancellationRequested();
                        await WriteCellAsync(writer, currentRowIndex, cellValue.CellIndex, cellValue.Value, cellValue.Prop, widths).ConfigureAwait(false);
                    }
                    await writer.WriteAsync(WorksheetXml.EndRow).ConfigureAwait(false);
                }
#endif
            }
            maxRowIndex = currentRowIndex;

            await writer.WriteAsync(WorksheetXml.EndSheetData).ConfigureAwait(false);

            if (_configuration.AutoFilter)
            {
                await writer.WriteAsync(WorksheetXml.Autofilter(GetDimensionRef(maxRowIndex, maxColumnIndex))).ConfigureAwait(false);
            }

            await writer.WriteAsync(WorksheetXml.Drawing(_currentSheetIndex)).ConfigureAwait(false);
            await writer.WriteAsync(WorksheetXml.EndWorksheet).ConfigureAwait(false);

            if (_configuration.FastMode && dimensionPlaceholderPostition != 0)
            {
                await WriteDimensionAsync(writer, maxRowIndex, maxColumnIndex, dimensionPlaceholderPostition).ConfigureAwait(false);
            }
            if (_configuration.EnableAutoWidth)
            {
                await OverwriteColumnWidthPlaceholdersAsync(writer, columnWidthsPlaceholderPosition, widths?.Columns, cancellationToken).ConfigureAwait(false);
            }

            if (_printHeader)
                maxRowIndex--;

            return maxRowIndex;
        }

        [Zomp.SyncMethodGenerator.CreateSyncVersion]
        private static async Task<long> WriteColumnWidthPlaceholdersAsync(
#if SYNC_ONLY
            global::MiniExcelLibs.OpenXml.MiniExcelStreamWriter writer,
#else
            MiniExcelAsyncStreamWriter writer,
#endif
            int count, CancellationToken cancellationToken = default)
        {
            cancellationToken.ThrowIfCancellationRequested();

            var placeholderPosition = await writer.FlushAsync().ConfigureAwait(false);
            await writer.WriteWhitespaceAsync(WorksheetXml.GetColumnPlaceholderLength(count)).ConfigureAwait(false);
            return placeholderPosition;
        }

        [Zomp.SyncMethodGenerator.CreateSyncVersion]
        private static async Task OverwriteColumnWidthPlaceholdersAsync(
#if SYNC_ONLY
            global::MiniExcelLibs.OpenXml.MiniExcelStreamWriter writer,
#else
            MiniExcelAsyncStreamWriter writer,
#endif
            long placeholderPosition, IEnumerable<ExcelColumnWidth> columnWidths, CancellationToken cancellationToken = default)
        {
            cancellationToken.ThrowIfCancellationRequested();

            var position = await writer.FlushAsync().ConfigureAwait(false);

            writer.SetPosition(placeholderPosition);
            await WriteColumnsWidthsAsync(writer, columnWidths, cancellationToken).ConfigureAwait(false);

            await writer.FlushAsync().ConfigureAwait(false);
            writer.SetPosition(position);
        }

        [Zomp.SyncMethodGenerator.CreateSyncVersion]
        private static async Task WriteColumnsWidthsAsync(
#if SYNC_ONLY
            global::MiniExcelLibs.OpenXml.MiniExcelStreamWriter writer,
#else
            MiniExcelAsyncStreamWriter writer,
#endif
            IEnumerable<ExcelColumnWidth> columnWidths, CancellationToken cancellationToken = default)
        {
            var hasWrittenStart = false;
            foreach (var column in columnWidths)
            {
                cancellationToken.ThrowIfCancellationRequested();

                if (!hasWrittenStart)
                {
                    await writer.WriteAsync(WorksheetXml.StartCols).ConfigureAwait(false);
                    hasWrittenStart = true;
                }
                await writer.WriteAsync(WorksheetXml.Column(column.Index, column.Width)).ConfigureAwait(false);
            }

            if (!hasWrittenStart)
                return;

            await writer.WriteAsync(WorksheetXml.EndCols).ConfigureAwait(false);
        }

        [Zomp.SyncMethodGenerator.CreateSyncVersion]
        private async Task PrintHeaderAsync(
#if SYNC_ONLY
            global::MiniExcelLibs.OpenXml.MiniExcelStreamWriter writer,
#else
            MiniExcelAsyncStreamWriter writer,
#endif
            List<ExcelColumnInfo> props, CancellationToken cancellationToken = default)
        {
            const int yIndex = 1;
            await writer.WriteAsync(WorksheetXml.StartRow(yIndex)).ConfigureAwait(false);

            var xIndex = 1;
            foreach (var p in props)
            {
                //reason : https://github.com/mini-software/MiniExcel/issues/142
                if (p != null)
                {
                    if (p.ExcelIgnore)
                        continue;

                    var r = ExcelOpenXmlUtils.ConvertXyToCell(xIndex, yIndex);
                    await WriteCellAsync(writer, r, columnName: p.ExcelColumnName).ConfigureAwait(false);
                }
                xIndex++;
            }

            await writer.WriteAsync(WorksheetXml.EndRow).ConfigureAwait(false);
        }

        [Zomp.SyncMethodGenerator.CreateSyncVersion]
        private async Task WriteCellAsync(
#if SYNC_ONLY
            global::MiniExcelLibs.OpenXml.MiniExcelStreamWriter writer,
#else
            MiniExcelAsyncStreamWriter writer,
#endif
            string cellReference, string columnName)
        {
            await writer.WriteAsync(WorksheetXml.Cell(cellReference, "str", GetCellXfId("1"), ExcelOpenXmlUtils.EncodeXML(columnName))).ConfigureAwait(false);
        }

        [Zomp.SyncMethodGenerator.CreateSyncVersion]
        private async Task WriteCellAsync(
#if SYNC_ONLY
            global::MiniExcelLibs.OpenXml.MiniExcelStreamWriter writer,
#else
            MiniExcelAsyncStreamWriter writer,
#endif
            int rowIndex, int cellIndex, object value, ExcelColumnInfo columnInfo, ExcelWidthCollection widthCollection)
        {
            if (columnInfo?.CustomFormatter != null)
            {
                try
                {
                    value = columnInfo.CustomFormatter(value);
                }
                catch
                {
                    //ignored
                }
            }

            var columnReference = ExcelOpenXmlUtils.ConvertXyToCell(cellIndex, rowIndex);
            var valueIsNull = value is null ||
                              value is DBNull ||
                              (_configuration.WriteEmptyStringAsNull && value is string vs && vs == string.Empty);

            if (_configuration.EnableWriteNullValueCell && valueIsNull)
            {
                await writer.WriteAsync(WorksheetXml.EmptyCell(columnReference, GetCellXfId("2"))).ConfigureAwait(false);
                return;
            }

            var tuple = GetCellValue(rowIndex, cellIndex, value, columnInfo, valueIsNull);

            var styleIndex = tuple.Item1;
            var dataType = tuple.Item2;
            var cellValue = tuple.Item3;
            var columnType = columnInfo.ExcelColumnType;

            /*Prefix and suffix blank space will lost after SaveAs #294*/
            var preserveSpace = cellValue != null && (cellValue.StartsWith(" ", StringComparison.Ordinal) ||
                                                      cellValue.EndsWith(" ", StringComparison.Ordinal));

            await writer.WriteAsync(WorksheetXml.Cell(columnReference, dataType, GetCellXfId(styleIndex), cellValue, preserveSpace: preserveSpace, columnType: columnType)).ConfigureAwait(false);
            widthCollection?.AdjustWidth(cellIndex, cellValue);
        }

        [Zomp.SyncMethodGenerator.CreateSyncVersion]
        private async Task GenerateEndXmlAsync(CancellationToken cancellationToken)
        {
            await AddFilesToZipAsync(cancellationToken).ConfigureAwait(false);
            await GenerateDrawinRelXmlAsync(cancellationToken).ConfigureAwait(false);
            await GenerateDrawingXmlAsync(cancellationToken).ConfigureAwait(false);
            await GenerateWorkbookXmlAsync(cancellationToken).ConfigureAwait(false);
            await GenerateContentTypesXmlAsync(cancellationToken).ConfigureAwait(false);
        }

        [Zomp.SyncMethodGenerator.CreateSyncVersion]
        private async Task AddFilesToZipAsync(CancellationToken cancellationToken)
        {
            foreach (var item in _files)
            {
                cancellationToken.ThrowIfCancellationRequested();
                await CreateZipEntryAsync(item.Path, item.Byte, cancellationToken).ConfigureAwait(false);
            }
        }

        [Zomp.SyncMethodGenerator.CreateSyncVersion]
        private async Task GenerateStylesXmlAsync(CancellationToken cancellationToken)
        {
            cancellationToken.ThrowIfCancellationRequested();

            using (var context = new SheetStyleBuildContext(_zipDictionary, _archive, _utf8WithBom, _configuration.DynamicColumns))
            {
                ISheetStyleBuilder builder = null;
                switch (_configuration.TableStyles)
                {
                    case TableStyles.None:
                        builder = new MinimalSheetStyleBuilder(context);
                        break;
                    case TableStyles.Default:
                        builder = new DefaultSheetStyleBuilder(context, _configuration.StyleOptions);
                        break;
                }
                var result = await builder.BuildAsync(cancellationToken).ConfigureAwait(false);
                _cellXfIdMap = result.CellXfIdMap;
            }
        }

        [Zomp.SyncMethodGenerator.CreateSyncVersion]
        private async Task GenerateDrawinRelXmlAsync(CancellationToken cancellationToken)
        {
            for (int sheetIndex = 0; sheetIndex < _sheets.Count; sheetIndex++)
            {
                cancellationToken.ThrowIfCancellationRequested();
                await GenerateDrawinRelXmlAsync(sheetIndex, cancellationToken).ConfigureAwait(false);
            }
        }

        [Zomp.SyncMethodGenerator.CreateSyncVersion]
        private async Task GenerateDrawinRelXmlAsync(int sheetIndex, CancellationToken cancellationToken)
        {
            var drawing = GetDrawingRelationshipXml(sheetIndex);
            await CreateZipEntryAsync(
                ExcelFileNames.DrawingRels(sheetIndex),
                string.Empty,
                ExcelXml.DefaultDrawingXmlRels.Replace("{{format}}", drawing),
                cancellationToken).ConfigureAwait(false);
        }

        [Zomp.SyncMethodGenerator.CreateSyncVersion]
        private async Task GenerateDrawingXmlAsync(CancellationToken cancellationToken)
        {
            for (int sheetIndex = 0; sheetIndex < _sheets.Count; sheetIndex++)
            {
                cancellationToken.ThrowIfCancellationRequested();
                await GenerateDrawingXmlAsync(sheetIndex, cancellationToken).ConfigureAwait(false);
            }
        }

        [Zomp.SyncMethodGenerator.CreateSyncVersion]
        private async Task GenerateDrawingXmlAsync(int sheetIndex, CancellationToken cancellationToken)
        {
            var drawing = GetDrawingXml(sheetIndex);
            await CreateZipEntryAsync(
                ExcelFileNames.Drawing(sheetIndex),
                ExcelContentTypes.Drawing,
                ExcelXml.DefaultDrawing.Replace("{{format}}", drawing),
                cancellationToken).ConfigureAwait(false);
        }

        [Zomp.SyncMethodGenerator.CreateSyncVersion]
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
                    cancellationToken).ConfigureAwait(false);
            }

            await CreateZipEntryAsync(
                ExcelFileNames.Workbook,
                ExcelContentTypes.Workbook,
                ExcelXml.DefaultWorkbookXml.Replace("{{sheets}}", workbookXml.ToString()),
                    cancellationToken).ConfigureAwait(false);

            await CreateZipEntryAsync(
                ExcelFileNames.WorkbookRels,
                null,
                ExcelXml.DefaultWorkbookXmlRels.Replace("{{sheets}}", workbookRelsXml.ToString()),
                    cancellationToken).ConfigureAwait(false);
        }

        [Zomp.SyncMethodGenerator.CreateSyncVersion]
        private async Task GenerateContentTypesXmlAsync(CancellationToken cancellationToken)
        {
            var contentTypes = GetContentTypesXml();
            await CreateZipEntryAsync(ExcelFileNames.ContentTypes, null, contentTypes, cancellationToken).ConfigureAwait(false);
        }

        [Zomp.SyncMethodGenerator.CreateSyncVersion]
        private async Task InsertContentTypesXmlAsync(CancellationToken cancellationToken)
        {
            cancellationToken.ThrowIfCancellationRequested();

            var contentTypesZipEntry = _archive.Entries.SingleOrDefault(s => s.FullName == ExcelFileNames.ContentTypes);
            if (contentTypesZipEntry == null)
            {
                await GenerateContentTypesXmlAsync(cancellationToken).ConfigureAwait(false);
                return;
            }
#if NET5_0_OR_GREATER
#pragma warning disable CA2007 // Consider calling ConfigureAwait on the awaited task
            await using (var stream = contentTypesZipEntry.Open())
#pragma warning restore CA2007 // Consider calling ConfigureAwait on the awaited task
#else
            using (var stream = contentTypesZipEntry.Open())
#endif
            {
#if NETCOREAPP2_0_OR_GREATER
                var doc = await XDocument.LoadAsync(stream, LoadOptions.None, cancellationToken).ConfigureAwait(false);
#else
                var doc = XDocument.Load(stream);
#endif
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
#if NETCOREAPP2_0_OR_GREATER
                await doc.SaveAsync(stream, SaveOptions.None, cancellationToken).ConfigureAwait(false);
#else
                doc.Save(stream);
#endif
            }
        }

        private async Task CreateZipEntryAsync(string path, string contentType, string content, CancellationToken cancellationToken)
        {
            cancellationToken.ThrowIfCancellationRequested();

            var entry = _archive.CreateEntry(path, CompressionLevel.Fastest);

#if NET5_0_OR_GREATER
#pragma warning disable CA2007 // Consider calling ConfigureAwait on the awaited task
            await using (var zipStream = entry.Open())
#pragma warning restore CA2007 // Consider calling ConfigureAwait on the awaited task
#else
            using (var zipStream = entry.Open())
#endif
            using (var writer = new MiniExcelAsyncStreamWriter(zipStream, _utf8WithBom, _configuration.BufferSize, cancellationToken))
                await writer.WriteAsync(content).ConfigureAwait(false);

            if (!string.IsNullOrEmpty(contentType))
                _zipDictionary.Add(path, new ZipPackageInfo(entry, contentType));
        }

        private async Task CreateZipEntryAsync(string path, byte[] content, CancellationToken cancellationToken)
        {
            cancellationToken.ThrowIfCancellationRequested();

            var entry = _archive.CreateEntry(path, CompressionLevel.Fastest);

#if NET5_0_OR_GREATER
#pragma warning disable CA2007 // Consider calling ConfigureAwait on the awaited task
            await using (var zipStream = entry.Open())
                await zipStream.WriteAsync(content, cancellationToken).ConfigureAwait(false);
#pragma warning restore CA2007 // Consider calling ConfigureAwait on the awaited task
#else
            using (var zipStream = entry.Open())
                await zipStream.WriteAsync(content, 0, content.Length, cancellationToken).ConfigureAwait(false);
#endif
        }
    }
}