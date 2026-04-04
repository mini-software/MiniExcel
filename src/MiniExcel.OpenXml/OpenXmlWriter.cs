using System.ComponentModel;
using System.Xml.Linq;
using MiniExcelLib.Core;
using MiniExcelLib.Core.WriteAdapters;
using MiniExcelLib.OpenXml.Constants;
using MiniExcelLib.OpenXml.Styles.Builder;

namespace MiniExcelLib.OpenXml;

internal partial class OpenXmlWriter : IMiniExcelWriter
{
    private static readonly UTF8Encoding Utf8WithBom = new(true);

    private readonly ZipArchive _archive;
    private readonly OpenXmlConfiguration _configuration;
    private readonly Stream _stream;
    private readonly List<SheetDto> _sheets = [];
    private readonly List<FileDto> _files = [];

    private readonly string? _defaultSheetName;
    private readonly bool _printHeader;
    private readonly object? _value;

    private int _currentSheetIndex = 0;

    private OpenXmlWriter(Stream stream, ZipArchive archive, object? value, string? sheetName, OpenXmlConfiguration configuration, bool printHeader)
    {
        _stream = stream;

        _configuration = configuration;
        _archive = archive;

        _value = value;
        _printHeader = printHeader;
        _defaultSheetName = sheetName;
    }

    [CreateSyncVersion]
    internal static async ValueTask<OpenXmlWriter> CreateAsync(Stream stream, object? value, string? sheetName, bool printHeader, IMiniExcelConfiguration? configuration, CancellationToken cancellationToken = default)
    {
        ThrowHelper.ThrowIfInvalidSheetName(sheetName);

        var conf = configuration as OpenXmlConfiguration ?? OpenXmlConfiguration.Default;
        if (conf is { EnableAutoWidth: true, FastMode: false })
            throw new InvalidOperationException("Auto width requires fast mode to be enabled");

        // A. Why ZipArchiveMode.Update and not ZipArchiveMode.Create?
        // R. ZipArchiveEntry does not support seeking when Mode is Create.
        var archiveMode = conf.FastMode ? ZipArchiveMode.Update : ZipArchiveMode.Create;

#if NET10_0_OR_GREATER
        var archive = await ZipArchive.CreateAsync(stream, archiveMode, true, Utf8WithBom, cancellationToken).ConfigureAwait(false);
#else
        var archive = new ZipArchive(stream, archiveMode, true, Utf8WithBom);
#endif
        return new OpenXmlWriter(stream, archive, value, sheetName, conf, printHeader);
    }

    [CreateSyncVersion]
    public async Task<int[]> SaveAsAsync(IProgress<int>? progress = null, CancellationToken cancellationToken = default)
    {
        try
        {
            await GenerateDefaultOpenXmlAsync(cancellationToken).ConfigureAwait(false);

            var sheets = GetSheets();
            var rowsWritten = new List<int>();

            foreach (var sheet in sheets)
            {
                cancellationToken.ThrowIfCancellationRequested();

                _sheets.Add(sheet.Item1); //TODO:remove
                _currentSheetIndex = sheet.Item1.SheetIdx;
                var rows = await CreateSheetXmlAsync(sheet.Item2, sheet.Item1.Path, progress, cancellationToken).ConfigureAwait(false);
                rowsWritten.Add(rows);
            }

            await GenerateEndXmlAsync(cancellationToken).ConfigureAwait(false);
            return rowsWritten.ToArray();
        }
        finally
        {
#if NET10_0_OR_GREATER
            await _archive.DisposeAsync().ConfigureAwait(false);
#else
            _archive.Dispose();
#endif
        }
    }

    [CreateSyncVersion]
    public async Task<int> InsertAsync(bool overwriteSheet = false, IProgress<int>? progress = null, CancellationToken cancellationToken = default)
    {
        try
        {
            if (!_configuration.FastMode)
                throw new InvalidOperationException("Insert requires fast mode to be enabled");

            using var reader = await OpenXmlReader.CreateAsync(_stream, _configuration, cancellationToken: cancellationToken).ConfigureAwait(false);
            var sheetRecords = (await reader.GetWorkbookRelsAsync(_archive.Entries, cancellationToken).ConfigureAwait(false))?.ToArray() ?? [];
            foreach (var sheetRecord in sheetRecords.OrderBy(o => o.Id))
            {
                cancellationToken.ThrowIfCancellationRequested();
                _sheets.Add(new SheetDto
                {
                    Name = sheetRecord.Name,
                    SheetIdx = (int)sheetRecord.Id,
                    State = sheetRecord.State
                });
            }

            var existSheetDto = _sheets.SingleOrDefault(s => s.Name == _defaultSheetName);
            if (existSheetDto is not null && !overwriteSheet)
                throw new Exception($"Sheet “{_defaultSheetName}” already exist");

            // GenerateStylesXml must be invoked after validating the overwritesheet parameter to avoid unnecessary style changes.
            await GenerateStylesXmlAsync(cancellationToken).ConfigureAwait(false); 

            int rowsWritten;
            if (existSheetDto is null)
            {
                _currentSheetIndex = (int)sheetRecords.Max(m => m.Id) + 1;
                var insertSheetInfo = GetSheetInfos(_defaultSheetName);
                var insertSheetDto = insertSheetInfo.ToDto(_currentSheetIndex);
                _sheets.Add(insertSheetDto);
                rowsWritten = await CreateSheetXmlAsync(_value, insertSheetDto.Path, progress, cancellationToken).ConfigureAwait(false);
            }
            else
            {
                _currentSheetIndex = existSheetDto.SheetIdx;
                _archive.Entries.Single(s => s.FullName == existSheetDto.Path).Delete();
                rowsWritten = await CreateSheetXmlAsync(_value, existSheetDto.Path, progress, cancellationToken).ConfigureAwait(false);
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
#if NET10_0_OR_GREATER
            await _archive.DisposeAsync().ConfigureAwait(false);
#else
            _archive.Dispose();
#endif
        }
    }

    [CreateSyncVersion]
    private async Task GenerateDefaultOpenXmlAsync(CancellationToken cancellationToken)
    {
        await CreateZipEntryAsync(ExcelFileNames.Rels, ExcelContentTypes.Relationships, ExcelXml.DefaultRels, cancellationToken).ConfigureAwait(false);
        await CreateZipEntryAsync(ExcelFileNames.SharedStrings, ExcelContentTypes.SharedStrings, ExcelXml.DefaultSharedString, cancellationToken).ConfigureAwait(false);
        await GenerateStylesXmlAsync(cancellationToken).ConfigureAwait(false);
    }

    [CreateSyncVersion]
    private async Task<int> CreateSheetXmlAsync(object? values, string sheetPath, IProgress<int>? progress, CancellationToken cancellationToken)
    {
        var entry = _archive.CreateEntry(sheetPath, CompressionLevel.Fastest);
        var rowsWritten = 0;

#if NET8_0_OR_GREATER
#if NET10_0_OR_GREATER
        var zipStream = await entry.OpenAsync(cancellationToken).ConfigureAwait(false);
#else
        var zipStream = entry.Open();
#endif
        await using var disposableZipStream = zipStream.ConfigureAwait(false);

        var writer = new MiniExcelStreamWriter(zipStream, Utf8WithBom, _configuration.BufferSize);
        await using var disposableWriter = writer.ConfigureAwait(false);
#else
        using var zipStream = entry.Open();
        using var writer = new MiniExcelStreamWriter(zipStream, Utf8WithBom, _configuration.BufferSize);
#endif
        if (values is null)
        {
            await WriteEmptySheetAsync(writer).ConfigureAwait(false);
        }
        else
        {
            rowsWritten = await WriteValuesAsync(writer, values, cancellationToken, progress).ConfigureAwait(false);
        }

        _zipDictionary.Add(sheetPath, new ZipPackageInfo(entry, ExcelContentTypes.Worksheet));
        return rowsWritten;
    }

    [CreateSyncVersion]
    private static async Task WriteEmptySheetAsync(MiniExcelStreamWriter writer)
    {
        await writer.WriteAsync(ExcelXml.EmptySheetXml).ConfigureAwait(false);
    }

    [CreateSyncVersion]
    private static async Task<long> WriteDimensionPlaceholderAsync(MiniExcelStreamWriter writer)
    {
        var dimensionPlaceholderPostition = await writer.WriteAndFlushAsync(WorksheetXml.StartDimension).ConfigureAwait(false);
        await writer.WriteAsync(WorksheetXml.DimensionPlaceholder).ConfigureAwait(false); // end of code will be replaced

        return dimensionPlaceholderPostition;
    }

    [CreateSyncVersion]
    private static async Task WriteDimensionAsync(MiniExcelStreamWriter writer, int maxRowIndex, int maxColumnIndex, long placeholderPosition)
    {
        // Flush and save position so that we can get back again.
        var position = await writer.FlushAndGetPositionAsync().ConfigureAwait(false);

        writer.SetPosition(placeholderPosition);
        await writer.WriteAndFlushAsync($@"{GetDimensionRef(maxRowIndex, maxColumnIndex)}""").ConfigureAwait(false);

        writer.SetPosition(position);
    }

    [CreateSyncVersion]
    private async Task<int> WriteValuesAsync(MiniExcelStreamWriter writer, object values, CancellationToken cancellationToken, IProgress<int>? progress = null)
    {
        cancellationToken.ThrowIfCancellationRequested();

        IMiniExcelWriteAdapter? writeAdapter = null;
        if (!MiniExcelWriteAdapterFactory.TryGetAsyncWriteAdapter(values, _configuration, out var asyncWriteAdapter))
        {
            writeAdapter = MiniExcelWriteAdapterFactory.GetWriteAdapter(values, _configuration);
        }
        try
        {
            var count = 0;
            var isKnownCount = writeAdapter is not null && writeAdapter.TryGetKnownCount(out count);

#if SYNC_ONLY
            var mappings = writeAdapter?.GetColumns();
#else
            var mappings = writeAdapter is not null
                ? writeAdapter.GetColumns()
                : await (asyncWriteAdapter?.GetColumnsAsync() ?? Task.FromResult<List<MiniExcelColumnMapping>?>(null)).ConfigureAwait(false);
#endif

            if (mappings is null)
            {
                await WriteEmptySheetAsync(writer).ConfigureAwait(false);
                return 0;
            }

            int maxRowIndex;
            var maxColumnIndex = mappings.Count(x => x is { ExcelIgnoreColumn: false });
            long dimensionPlaceholderPostition = 0;

            await writer.WriteAsync(WorksheetXml.StartWorksheetWithRelationship, cancellationToken).ConfigureAwait(false);

            // We can write the dimensions directly if the row count is known
            if (isKnownCount)
            {
                maxRowIndex = _printHeader ? count + 1 : count;
                await writer.WriteAsync(WorksheetXml.Dimension(GetDimensionRef(maxRowIndex, mappings.Count)), cancellationToken).ConfigureAwait(false);
            }
            else if (_configuration.FastMode)
            {
                dimensionPlaceholderPostition = await WriteDimensionPlaceholderAsync(writer).ConfigureAwait(false);
            }

            //sheet view
            await writer.WriteAsync(GetSheetViews(), cancellationToken).ConfigureAwait(false);

            //cols:width
            ExcelColumnWidthCollection? widths = null;
            long columnWidthsPlaceholderPosition = 0;
            if (_configuration.EnableAutoWidth)
            {
                columnWidthsPlaceholderPosition = await WriteColumnWidthPlaceholdersAsync(writer, maxColumnIndex, cancellationToken).ConfigureAwait(false);
                widths = ExcelColumnWidthCollection.GetFromMappings(mappings!, _configuration.MinWidth, _configuration.MaxWidth);
            }
            else
            {
                var colWidths = ExcelColumnWidthCollection.GetFromMappings(mappings!);
                await WriteColumnsWidthsAsync(writer, colWidths.Columns, cancellationToken).ConfigureAwait(false);
            }

            //header
            await writer.WriteAsync(WorksheetXml.StartSheetData, cancellationToken).ConfigureAwait(false);
            var currentRowIndex = 0;
            if (_printHeader)
            {
                await PrintHeaderAsync(writer, mappings!, cancellationToken).ConfigureAwait(false);
                currentRowIndex++;
            }

            if (writeAdapter is not null)
            {
                foreach (var row in writeAdapter.GetRows(mappings, cancellationToken))
                {
                    cancellationToken.ThrowIfCancellationRequested();
                    await writer.WriteAsync(WorksheetXml.StartRow(++currentRowIndex), cancellationToken).ConfigureAwait(false);

                    foreach (var cellValue in row)
                    {
                        cancellationToken.ThrowIfCancellationRequested();
                        await WriteCellAsync(writer, currentRowIndex, cellValue.CellIndex, cellValue.Value, cellValue.Mapping, widths, cancellationToken).ConfigureAwait(false);
                        progress?.Report(1);
                    }
                    await writer.WriteAsync(WorksheetXml.EndRow, cancellationToken).ConfigureAwait(false);
                }
            }
            else
            {
#if !SYNC_ONLY
                await foreach (var row in asyncWriteAdapter!.GetRowsAsync(mappings, cancellationToken).ConfigureAwait(false))
                {
                    cancellationToken.ThrowIfCancellationRequested();
                    await writer.WriteAsync(WorksheetXml.StartRow(++currentRowIndex), cancellationToken).ConfigureAwait(false);

                    foreach (var cellValue in row)
                    {
                        await WriteCellAsync(writer, currentRowIndex, cellValue.CellIndex, cellValue.Value, cellValue.Mapping, widths, cancellationToken).ConfigureAwait(false);
                        progress?.Report(1);
                    }
                    await writer.WriteAsync(WorksheetXml.EndRow, cancellationToken).ConfigureAwait(false);
                }
#endif
            }
            maxRowIndex = currentRowIndex;

            await writer.WriteAsync(WorksheetXml.EndSheetData, cancellationToken).ConfigureAwait(false);

            if (_configuration.AutoFilter)
            {
                await writer.WriteAsync(WorksheetXml.Autofilter(GetDimensionRef(maxRowIndex, maxColumnIndex)), cancellationToken).ConfigureAwait(false);
            }

            await writer.WriteAsync(WorksheetXml.Drawing(_currentSheetIndex), cancellationToken).ConfigureAwait(false);
            await writer.WriteAsync(WorksheetXml.EndWorksheet, cancellationToken).ConfigureAwait(false);

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
        finally
        {
#if !SYNC_ONLY
            if (asyncWriteAdapter is IAsyncDisposable asyncDisposable)
            {
                await asyncDisposable.DisposeAsync().ConfigureAwait(false);
            }
#endif
        }
    }

    [CreateSyncVersion]
    private static async Task<long> WriteColumnWidthPlaceholdersAsync(MiniExcelStreamWriter writer, int count, CancellationToken cancellationToken = default)
    {
        cancellationToken.ThrowIfCancellationRequested();

        var placeholderPosition = await writer.FlushAndGetPositionAsync(cancellationToken).ConfigureAwait(false);
        var placeholderLength = WorksheetXml.GetColumnPlaceholderLength(count);
        await writer.WriteAsync(new string(' ', placeholderLength), cancellationToken).ConfigureAwait(false);
        return placeholderPosition;
    }

    [CreateSyncVersion]
    private static async Task OverwriteColumnWidthPlaceholdersAsync(MiniExcelStreamWriter writer, long placeholderPosition, IEnumerable<ExcelColumnWidth>? columnWidths, CancellationToken cancellationToken = default)
    {
        cancellationToken.ThrowIfCancellationRequested();

        var position = await writer.FlushAndGetPositionAsync(cancellationToken).ConfigureAwait(false);

        writer.SetPosition(placeholderPosition);
        await WriteColumnsWidthsAsync(writer, columnWidths, cancellationToken).ConfigureAwait(false);

        await writer.FlushAndGetPositionAsync(cancellationToken).ConfigureAwait(false);
        writer.SetPosition(position);
    }

    [CreateSyncVersion]
    private static async Task WriteColumnsWidthsAsync(MiniExcelStreamWriter writer, IEnumerable<ExcelColumnWidth>? columnWidths, CancellationToken cancellationToken = default)
    {
        var hasWrittenStart = false;

        columnWidths ??= [];
        foreach (var column in columnWidths)
        {
            cancellationToken.ThrowIfCancellationRequested();

            if (!hasWrittenStart)
            {
                await writer.WriteAsync(WorksheetXml.StartCols, cancellationToken).ConfigureAwait(false);
                hasWrittenStart = true;
            }
            await writer.WriteAsync(WorksheetXml.Column(column.Index, column.Width, column.Hidden), cancellationToken).ConfigureAwait(false);
        }

        if (!hasWrittenStart)
            return;

        await writer.WriteAsync(WorksheetXml.EndCols, cancellationToken).ConfigureAwait(false);
    }

    [CreateSyncVersion]
    private async Task PrintHeaderAsync(MiniExcelStreamWriter writer, List<MiniExcelColumnMapping?> mappings, CancellationToken cancellationToken = default)
    {
        const int yIndex = 1;
        await writer.WriteAsync(WorksheetXml.StartRow(yIndex), cancellationToken).ConfigureAwait(false);

        var xIndex = 1;
        foreach (var map in mappings)
        {
            //reason : https://github.com/mini-software/MiniExcel/issues/142
            if (map is not null)
            {
                if (map.ExcelIgnoreColumn)
                    continue;

                var r = CellReferenceConverter.GetCellFromCoordinates(xIndex, yIndex);
                await WriteCellAsync(writer, r, columnName: map.ExcelColumnName).ConfigureAwait(false);
            }
            xIndex++;
        }

        await writer.WriteAsync(WorksheetXml.EndRow, cancellationToken).ConfigureAwait(false);
    }

    [CreateSyncVersion]
    private async Task WriteCellAsync(MiniExcelStreamWriter writer, string cellReference, string? columnName)
    {
        await writer.WriteAsync(WorksheetXml.Cell(cellReference, "str", GetCellXfId("1"), XmlHelper.EncodeXml(columnName))).ConfigureAwait(false);
    }

    [CreateSyncVersion]
    private async Task WriteCellAsync(MiniExcelStreamWriter writer, int rowIndex, int cellIndex, object? value, MiniExcelColumnMapping? columnInfo, ExcelColumnWidthCollection? widthCollection, CancellationToken cancellationToken = default)
    {
        if (columnInfo?.CustomFormatter is { } formatter)
        {
            try
            {
                value = formatter.Invoke(value);
            }
            catch
            {
                //ignored
            }
        }

        var columnReference = CellReferenceConverter.GetCellFromCoordinates(cellIndex, rowIndex);
        var valueIsNull = value is null or DBNull || (_configuration.WriteEmptyStringAsNull && value is "");

        if (_configuration.EnableWriteNullValueCell && valueIsNull)
        {
            await writer.WriteAsync(WorksheetXml.EmptyCell(columnReference, GetCellXfId("2")), cancellationToken).ConfigureAwait(false);
            return;
        }

        var tuple = GetCellValue(rowIndex, cellIndex, value, columnInfo, valueIsNull);

        var styleIndex = tuple.Item1;
        var dataType = tuple.Item2;
        var cellValue = tuple.Item3;
        var columnType = columnInfo.ExcelColumnType;

        /*Prefix and suffix blank space will lost after SaveAs #294*/
        var preserveSpace = cellValue.StartsWith(" ") || cellValue.EndsWith(" ");

        await writer.WriteAsync(WorksheetXml.Cell(columnReference, dataType, GetCellXfId(styleIndex), cellValue, preserveSpace: preserveSpace, columnType: columnType), cancellationToken).ConfigureAwait(false);
        widthCollection?.AdjustWidth(cellIndex, cellValue);
    }

    [CreateSyncVersion]
    private async Task GenerateEndXmlAsync(CancellationToken cancellationToken)
    {
        await AddFilesToZipAsync(cancellationToken).ConfigureAwait(false);
        await GenerateDrawinRelXmlAsync(cancellationToken).ConfigureAwait(false);
        await GenerateDrawingXmlAsync(cancellationToken).ConfigureAwait(false);
        await GenerateWorkbookXmlAsync(cancellationToken).ConfigureAwait(false);
        await GenerateContentTypesXmlAsync(cancellationToken).ConfigureAwait(false);
    }

    [CreateSyncVersion]
    private async Task AddFilesToZipAsync(CancellationToken cancellationToken)
    {
        foreach (var item in _files)
        {
            cancellationToken.ThrowIfCancellationRequested();
            await CreateZipEntryAsync(item.Path, item.Byte, cancellationToken).ConfigureAwait(false);
        }
    }

    [CreateSyncVersion]
    private async Task GenerateStylesXmlAsync(CancellationToken cancellationToken)
    {
#if NET8_0_OR_GREATER        
        var context = new SheetStyleBuildContext(_zipDictionary, _archive, Utf8WithBom, _configuration.DynamicColumns ?? []);
        await using var disposableContext = context.ConfigureAwait(false); 
#else
        using var context = new SheetStyleBuildContext(_zipDictionary, _archive, Utf8WithBom, _configuration.DynamicColumns ?? []);
#endif
        ISheetStyleBuilder builder = _configuration.TableStyles switch
        {
            TableStyles.None => new MinimalSheetStyleBuilder(context),
            TableStyles.Default => new DefaultSheetStyleBuilder(context, _configuration.StyleOptions),
            _ => throw new InvalidEnumArgumentException(nameof(_configuration.TableStyles), (int)_configuration.TableStyles, typeof(TableStyles))
        };

        var result = await builder.BuildAsync(cancellationToken).ConfigureAwait(false);
        _cellXfIdMap = result.CellXfIdMap;
    }

    [CreateSyncVersion]
    private async Task GenerateDrawinRelXmlAsync(CancellationToken cancellationToken)
    {
        for (int sheetIndex = 0; sheetIndex < _sheets.Count; sheetIndex++)
        {
            cancellationToken.ThrowIfCancellationRequested();
            await GenerateDrawinRelXmlAsync(sheetIndex, cancellationToken).ConfigureAwait(false);
        }
    }

    [CreateSyncVersion]
    private async Task GenerateDrawinRelXmlAsync(int sheetIndex, CancellationToken cancellationToken)
    {
        var drawing = GetDrawingRelationshipXml(sheetIndex);
        await CreateZipEntryAsync(
            ExcelFileNames.DrawingRels(sheetIndex),
            string.Empty,
            ExcelXml.DefaultDrawingXmlRels.Replace("{{format}}", drawing),
            cancellationToken).ConfigureAwait(false);
    }

    [CreateSyncVersion]
    private async Task GenerateDrawingXmlAsync(CancellationToken cancellationToken)
    {
        for (int sheetIndex = 0; sheetIndex < _sheets.Count; sheetIndex++)
        {
            cancellationToken.ThrowIfCancellationRequested();
            await GenerateDrawingXmlAsync(sheetIndex, cancellationToken).ConfigureAwait(false);
        }
    }

    [CreateSyncVersion]
    private async Task GenerateDrawingXmlAsync(int sheetIndex, CancellationToken cancellationToken)
    {
        var drawing = GetDrawingXml(sheetIndex);
        await CreateZipEntryAsync(
            ExcelFileNames.Drawing(sheetIndex),
            ExcelContentTypes.Drawing,
            ExcelXml.DefaultDrawing.Replace("{{format}}", drawing),
            cancellationToken).ConfigureAwait(false);
    }

    [CreateSyncVersion]
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

    [CreateSyncVersion]
    private async Task GenerateContentTypesXmlAsync(CancellationToken cancellationToken)
    {
        var contentTypes = GetContentTypesXml();
        await CreateZipEntryAsync(ExcelFileNames.ContentTypes, null, contentTypes, cancellationToken).ConfigureAwait(false);
    }

    [CreateSyncVersion]
    private async Task InsertContentTypesXmlAsync(CancellationToken cancellationToken)
    {
        var contentTypesZipEntry = _archive.Entries.SingleOrDefault(s => s.FullName == ExcelFileNames.ContentTypes);
        if (contentTypesZipEntry is null)
        {
            await GenerateContentTypesXmlAsync(cancellationToken).ConfigureAwait(false);
            return;
        }

#if NET8_0_OR_GREATER
#if NET10_0_OR_GREATER
        var stream = await contentTypesZipEntry.OpenAsync(cancellationToken).ConfigureAwait(false);
#else
        var stream = contentTypesZipEntry.Open();
#endif
        await using var disposableStream = stream.ConfigureAwait(false);
#else
        using var stream = contentTypesZipEntry.Open();
#endif

#if NETCOREAPP2_0_OR_GREATER
        var doc = await XDocument.LoadAsync(stream, LoadOptions.None, cancellationToken).ConfigureAwait(false);
#else
        var doc = XDocument.Load(stream);
#endif
        var ns = doc.Root?.GetDefaultNamespace();
        var typesElement = doc.Descendants(ns + "Types").Single();

        var partNames = new HashSet<string>(StringComparer.InvariantCultureIgnoreCase);
        var attrNames = typesElement.Elements(ns + "Override").Select(s => s.Attribute("PartName")?.Value);
        foreach (var partName in attrNames.OfType<string>())
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
#if NET8_0_OR_GREATER
        await doc.SaveAsync(stream, SaveOptions.None, cancellationToken).ConfigureAwait(false);
#else
        doc.Save(stream);
#endif
    }

    [CreateSyncVersion]
    private async Task CreateZipEntryAsync(string path, string? contentType, string content, CancellationToken cancellationToken)
    {
        cancellationToken.ThrowIfCancellationRequested();

        var entry = _archive.CreateEntry(path, CompressionLevel.Fastest);

#if NET8_0_OR_GREATER
#if NET10_0_OR_GREATER
        var zipStream = await entry.OpenAsync(cancellationToken).ConfigureAwait(false);
#else
        var zipStream = entry.Open();
#endif
        await using var disposableZipStream = zipStream.ConfigureAwait(false);
        var writer = new MiniExcelStreamWriter(zipStream, Utf8WithBom, _configuration.BufferSize);
        await using var disposableWriter = writer.ConfigureAwait(false);
#else
        using var zipStream = entry.Open();
        using var writer = new MiniExcelStreamWriter(zipStream, Utf8WithBom, _configuration.BufferSize);
#endif
        await writer.WriteAsync(content, cancellationToken).ConfigureAwait(false);

        if (contentType is not (null or ""))
            _zipDictionary.Add(path, new ZipPackageInfo(entry, contentType));
    }

    [CreateSyncVersion]
    private async Task CreateZipEntryAsync(string path, byte[] content, CancellationToken cancellationToken)
    {
        cancellationToken.ThrowIfCancellationRequested();

        var entry = _archive.CreateEntry(path, CompressionLevel.Fastest);

#if NET8_0_OR_GREATER
#if NET10_0_OR_GREATER
        var zipStream = await entry.OpenAsync(cancellationToken).ConfigureAwait(false);
#else
        var zipStream = entry.Open();
#endif
        await using var  disposableZipStream = zipStream.ConfigureAwait(false);
        await zipStream.WriteAsync(content, cancellationToken).ConfigureAwait(false);
#else
        using var zipStream = entry.Open();
        await zipStream.WriteAsync(content, 0, content.Length, cancellationToken).ConfigureAwait(false);
#endif
    }
}
