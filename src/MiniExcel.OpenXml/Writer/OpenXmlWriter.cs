using MiniExcelLib.Core.WriteAdapters;
using MiniExcelLib.OpenXml.Reader;
using MiniExcelLib.OpenXml.Styles.Builder;

namespace MiniExcelLib.OpenXml.Writer;

internal sealed partial class OpenXmlWriter : IMiniExcelWriter
{
    private static readonly UTF8Encoding Utf8WithBom = new(true);

    private readonly Stream _stream;
    private readonly ZipArchive _archive;
    
    private readonly OpenXmlConfiguration _configuration;
    private readonly List<SheetDto> _sheets = [];
    private readonly List<FileDto> _files = [];
    
    private readonly string _sheetName;
    private readonly bool _printHeader;
    private readonly object? _value;

    // the index is 1-based to match how the rels work in OpenXml and to make it more intuitive to reason about them 
    private int _currentSheetIndex;
    private SheetStyleBuilderContext _sheetStyleBuilderContext;


    private OpenXmlWriter(Stream stream, ZipArchive archive, object? value, string sheetName, OpenXmlConfiguration configuration, bool printHeader)
    {
        _stream = stream;

        _configuration = configuration;
        _archive = archive;

        _value = value;
        _printHeader = printHeader;
        _sheetName = sheetName;

        _sheetStyleBuilderContext = new SheetStyleBuilderContext(_zipContentsMap, _archive, Utf8WithBom);
    }

    [CreateSyncVersion]
    internal static async ValueTask<OpenXmlWriter> CreateAsync(Stream stream, object? value, string sheetName, bool printHeader, IMiniExcelConfiguration? configuration, CancellationToken cancellationToken = default)
    {
        ThrowHelper.ThrowIfInvalidSheetName(sheetName);

        var conf = configuration as OpenXmlConfiguration ?? OpenXmlConfiguration.Default;
        if (conf is { EnableAutoWidth: true, FastMode: false })
            throw new InvalidOperationException("Auto width requires fast mode to be enabled");

        // A. Why ZipArchiveMode.Update and not ZipArchiveMode.Create?
        // R. ZipArchiveEntry does not support seeking when Mode is Create.
        var archiveMode = conf.FastMode ? ZipArchiveMode.Update : ZipArchiveMode.Create;
        var archive = await ZipArchive.CreateAsync(stream, archiveMode, true, Utf8WithBom, cancellationToken).ConfigureAwait(false);

        return new OpenXmlWriter(stream, archive, value, sheetName, conf, printHeader);
    }

    [CreateSyncVersion]
    public async Task<int[]> SaveAsAsync(IProgress<int>? progress = null, CancellationToken cancellationToken = default)
    {
#if NET10_0_OR_GREATER
        await using var disposableArchive = _archive.ConfigureAwait(false);
#else
        using var disposableArchive = _archive;
#endif
        await CreateZipEntryAsync(ExcelFileNames.Rels, ExcelContentTypes.Relationships, ExcelXml.DefaultRels, cancellationToken).ConfigureAwait(false);

        await using var sbc = _sheetStyleBuilderContext.ConfigureAwait(false);
        var styleBuilder = await GetSheetStyleBuilderAsync(cancellationToken).ConfigureAwait(false);

        var sheets = GetSheets();
        var rowsWritten = new List<int>();

        foreach (var sheet in sheets)
        {
            _sheets.Add(sheet.Sheet); //TODO:remove
            _currentSheetIndex = sheet.Sheet.SheetIdx;

            var rows = await CreateSheetXmlAsync(sheet.Data, sheet.Sheet.Path, progress, cancellationToken).ConfigureAwait(false);
            rowsWritten.Add(rows);
        }

        await styleBuilder.BuildAsync(cancellationToken).ConfigureAwait(false);

        await AddFilesToZipAsync(cancellationToken).ConfigureAwait(false);
        await GenerateSharedStringsAsync(cancellationToken).ConfigureAwait(false);
        await GenerateDrawingRelXmlAsync(cancellationToken).ConfigureAwait(false);
        await GenerateDrawingXmlAsync(cancellationToken).ConfigureAwait(false);
        await GenerateWorkbookXmlAsync(false, cancellationToken).ConfigureAwait(false);
        await GenerateContentTypesXmlAsync(cancellationToken).ConfigureAwait(false);

        return rowsWritten.ToArray();
    }

    [CreateSyncVersion]
    public async Task<int> InsertAsync(bool overwriteSheet = false, IProgress<int>? progress = null, CancellationToken cancellationToken = default)
    {
        if (!_configuration.FastMode)
            throw new InvalidOperationException("Insert requires fast mode to be enabled");

#if NET10_0_OR_GREATER
        await using var disposableArchive = _archive.ConfigureAwait(false);
#else
        using var disposableArchive = _archive;
#endif
        await using var sbc = _sheetStyleBuilderContext.ConfigureAwait(false);

        using var reader = await OpenXmlReader.CreateAsync(_stream, _configuration, cancellationToken: cancellationToken).ConfigureAwait(false);
        var rels = await OpenXmlReader.GetWorkbookRelsAsync(_archive.Entries, cancellationToken).ConfigureAwait(false) ?? [];

        _sheets.AddRange(rels
            .OrderBy(sheet => sheet.Id)
            .Select(sheet => new SheetDto
            {
                Name = sheet.Name,
                SheetIdx = (int)sheet.Id,
                State = sheet.State
            })
        );

        var existingSheetDto = _sheets.SingleOrDefault(s => s.Name == _sheetName);
        if (existingSheetDto is not null && !overwriteSheet)
            throw new InvalidOperationException($"Sheet \"{_sheetName}\" already exists");

        // GenerateStylesXml must be invoked after validating the overwritesheet parameter to avoid unnecessary style changes.
        var styleBuilder = await GetSheetStyleBuilderAsync(cancellationToken).ConfigureAwait(false);

         var sharedStringsEntry = _archive.GetEntry(ExcelFileNames.SharedStrings);
         if (sharedStringsEntry is not null)
         {
             foreach (var (key, value) in reader.SharedStrings)
             {
                 _sharedStrings[value] = key;
             }
         }

        int rowsWritten;
        if (existingSheetDto is null)
        {
            _currentSheetIndex = (int)rels.Max(m => m.Id) + 1;
            var insertSheetInfo = GetSheetInfos(_sheetName);
            var insertSheetDto = insertSheetInfo.ToDto(_currentSheetIndex);
            _sheets.Add(insertSheetDto);
            rowsWritten = await CreateSheetXmlAsync(_value, insertSheetDto.Path, progress, cancellationToken).ConfigureAwait(false);
        }
        else
        {
            _currentSheetIndex = existingSheetDto.SheetIdx;
            _archive.Entries.Single(s => s.FullName == existingSheetDto.Path).Delete();
            rowsWritten = await CreateSheetXmlAsync(_value, existingSheetDto.Path, progress, cancellationToken).ConfigureAwait(false);
        }

        await styleBuilder.BuildAsync(cancellationToken).ConfigureAwait(false);
        await AddFilesToZipAsync(cancellationToken).ConfigureAwait(false);

        sharedStringsEntry?.Delete();
        await GenerateSharedStringsAsync(cancellationToken).ConfigureAwait(false);

        _archive.Entries.SingleOrDefault(s => s.FullName == ExcelFileNames.DrawingRels(_currentSheetIndex))?.Delete();
        await GenerateDrawingRelXmlAsync(_currentSheetIndex, cancellationToken).ConfigureAwait(false);

        _archive.Entries.SingleOrDefault(s => s.FullName == ExcelFileNames.Drawing(_currentSheetIndex))?.Delete();
        await GenerateDrawingXmlAsync(_currentSheetIndex, cancellationToken).ConfigureAwait(false);
        await GenerateWorkbookXmlAsync(true, cancellationToken).ConfigureAwait(false);
        await InsertContentTypesXmlAsync(cancellationToken).ConfigureAwait(false);

        return rowsWritten;
    }

    [CreateSyncVersion]
    private async Task<int> CreateSheetXmlAsync(object? values, string sheetPath, IProgress<int>? progress, CancellationToken cancellationToken)
    {
        var entry = _archive.CreateEntry(sheetPath, CompressionLevel.Fastest);
        var rowsWritten = 0;

        var zipStream = await entry.OpenAsync(cancellationToken).ConfigureAwait(false);
        await using var disposableZipStream = zipStream.ConfigureAwait(false);

        var writer = new MiniExcelStreamWriter(zipStream, Utf8WithBom, _configuration.BufferSize);
        await using var disposableWriter = writer.ConfigureAwait(false);

        if (values is null)
        {
            await WriteEmptySheetAsync(writer).ConfigureAwait(false);
        }
        else
        {
            rowsWritten = await WriteValuesAsync(writer, values, cancellationToken, progress).ConfigureAwait(false);
        }

        _zipContentsMap.Add(sheetPath, ExcelContentTypes.Worksheet);
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
        IMiniExcelWriteAdapter? writeAdapter = null;
        if (!MiniExcelWriteAdapterFactory.TryGetAsyncWriteAdapter(values, _configuration, out var asyncWriteAdapter))
        {
            writeAdapter = MiniExcelWriteAdapterFactory.GetWriteAdapter(values, _configuration);
        }

        try
        {
            var count = 0;
            var isKnownCount = writeAdapter?.TryGetKnownCount(out count) is true;

#if SYNC_ONLY
            var mappings = writeAdapter?.GetColumns();
#else
            var mappings = asyncWriteAdapter is not null
                ? await asyncWriteAdapter.GetColumnsAsync().ConfigureAwait(false)
                : writeAdapter?.GetColumns() ?? [];
#endif

            if (mappings is null or [])
            {
                await WriteEmptySheetAsync(writer).ConfigureAwait(false);
                return 0;
            }

            _sheetStyleBuilderContext.UpdateFormatIds(mappings);
            
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
            else if (_archive.Mode == ZipArchiveMode.Update)
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

                    foreach (var cell in row)
                    {
                        cancellationToken.ThrowIfCancellationRequested();
                        await WriteCellAsync(writer, currentRowIndex, cell.Index, cell.Value, cell.Mapping, widths, cancellationToken).ConfigureAwait(false);
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

                    foreach (var cell in row)
                    {
                        await WriteCellAsync(writer, currentRowIndex, cell.Index, cell.Value, cell.Mapping, widths, cancellationToken).ConfigureAwait(false);
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

            if (_archive.Mode == ZipArchiveMode.Update && dimensionPlaceholderPostition != 0)
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
    private static async Task PrintHeaderAsync(MiniExcelStreamWriter writer, List<MiniExcelColumnMapping?> mappings, CancellationToken cancellationToken = default)
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
                await writer.WriteAsync(WorksheetXml.Cell(r, ExcelDataTypes.InlineString, HeaderCellStyleIndex, map.ExcelColumnName), cancellationToken).ConfigureAwait(false);
            }
            xIndex++;
        }

        await writer.WriteAsync(WorksheetXml.EndRow, cancellationToken).ConfigureAwait(false);
    }

    [CreateSyncVersion]
    private async Task WriteCellAsync(MiniExcelStreamWriter writer, int rowIndex, int cellIndex, object? value, MiniExcelColumnMapping? columnMapping, ExcelColumnWidthCollection? widthCollection, CancellationToken cancellationToken = default)
    {
        if (columnMapping?.CustomFormatter is { } formatter)
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
            await writer.WriteAsync(WorksheetXml.EmptyCell(columnReference, DefaultCellStyleIndex), cancellationToken).ConfigureAwait(false);
            return;
        }

        var columnType = columnMapping.ExcelColumnType;
        var (styleIndex, dataType, cellValue) = GetCellValue(rowIndex, cellIndex, value, columnMapping, valueIsNull);
        widthCollection?.AdjustWidth(cellIndex, cellValue);

        if (_configuration.StringStorageMode == StringStorageMode.Shared && 
            dataType == ExcelDataTypes.SharedString 
            && cellValue is not null)
        {
            if (_sharedStrings.TryGetValue(cellValue, out var sharedStringIndex))
            {
                cellValue = sharedStringIndex.ToString();
            }
            else
            {
                var count = _sharedStrings.Count;
                _sharedStrings[cellValue] = count;
                cellValue = count.ToString();
            }
        }

        /*Prefix and suffix blank space will lost after SaveAs #294*/
        var preserveSpace = cellValue.StartsWith(" ") || cellValue.EndsWith(" ");
        await writer.WriteAsync(WorksheetXml.Cell(columnReference, dataType, styleIndex, cellValue, preserveSpace: preserveSpace, columnType: columnType), cancellationToken).ConfigureAwait(false);
    }

    [CreateSyncVersion]
    private async Task AddFilesToZipAsync(CancellationToken cancellationToken)
    {
        foreach (var item in _files)
        {
            cancellationToken.ThrowIfCancellationRequested();

            var entry = _archive.CreateEntry(item.Path, CompressionLevel.Fastest);

            var zipStream = await entry.OpenAsync(cancellationToken).ConfigureAwait(false);
            await using var disposableZipStream = zipStream.ConfigureAwait(false);

#if NET
            await zipStream.WriteAsync(item.Contents, cancellationToken).ConfigureAwait(false);
#else
            await zipStream.WriteAsync(item.Contents, 0, item.Contents.Length, cancellationToken).ConfigureAwait(false);
#endif
        }
    }

    [CreateSyncVersion]
    private async Task<ISheetStyleBuilder> GetSheetStyleBuilderAsync(CancellationToken cancellationToken = default)
    {
        SheetStyleBuilderBase builder = _configuration.TableStyles switch
        {
            TableStyles.None => new MinimalSheetStyleBuilder(_sheetStyleBuilderContext),
            TableStyles.Default => new DefaultSheetStyleBuilder(_sheetStyleBuilderContext, _configuration.StyleOptions),
            _ => throw new InvalidEnumArgumentException(nameof(_configuration.TableStyles), (int)_configuration.TableStyles, typeof(TableStyles))
        };

        var newInfos = builder.GetGeneratedElementInfos();
        await _sheetStyleBuilderContext.CreateAsync(newInfos, cancellationToken).ConfigureAwait(false);

        return builder;
    }

    [CreateSyncVersion]
    private async Task GenerateDrawingRelXmlAsync(CancellationToken cancellationToken)
    {
        for (int sheetIndex = 1; sheetIndex <= _sheets.Count; sheetIndex++)
        {
            cancellationToken.ThrowIfCancellationRequested();
            await GenerateDrawingRelXmlAsync(sheetIndex, cancellationToken).ConfigureAwait(false);
        }
    }

    [CreateSyncVersion]
    private async Task GenerateDrawingRelXmlAsync(int sheetIndex, CancellationToken cancellationToken)
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
        for (int sheetIndex = 1; sheetIndex <= _sheets.Count; sheetIndex++)
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
    private async Task GenerateWorkbookXmlAsync(bool removeOriginalEntry = false, CancellationToken cancellationToken = default)
    {
        var (workbookXml, workbookRelsXml, sheetsRelsXml) = GenerateWorkbookXmls();
        foreach (var (key, value) in sheetsRelsXml)
        {
            var sheetRelsXmlPath = ExcelFileNames.SheetRels(key);
            if (removeOriginalEntry)
                _archive.Entries.SingleOrDefault(s => s.FullName == sheetRelsXmlPath)?.Delete();

            await CreateZipEntryAsync(
                sheetRelsXmlPath,
                null,
                ExcelXml.DefaultSheetRelXml.Replace("{{format}}", value),
                cancellationToken).ConfigureAwait(false);
        }

        if(removeOriginalEntry)
            _archive.Entries.SingleOrDefault(s => s.FullName == ExcelFileNames.Workbook)?.Delete();

        await CreateZipEntryAsync(
            ExcelFileNames.Workbook,
            ExcelContentTypes.Workbook,
            ExcelXml.DefaultWorkbookXml.Replace("{{sheets}}", workbookXml),
            cancellationToken).ConfigureAwait(false);

        if(removeOriginalEntry)
            _archive.Entries.SingleOrDefault(s => s.FullName == ExcelFileNames.WorkbookRels)?.Delete();

        await CreateZipEntryAsync(
            ExcelFileNames.WorkbookRels,
            null,
            ExcelXml.DefaultWorkbookXmlRels.Replace("{{sheets}}", workbookRelsXml),
            cancellationToken).ConfigureAwait(false);
    }

    [CreateSyncVersion]
    private async Task GenerateSharedStringsAsync(CancellationToken cancellationToken)
    {
        await CreateZipEntryAsync(
            ExcelFileNames.SharedStrings, 
            ExcelContentTypes.SharedStrings, 
            ExcelXml.SharedStrings(_sharedStrings), 
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

        var stream = await contentTypesZipEntry.OpenAsync(cancellationToken).ConfigureAwait(false);
        await using var disposableStream = stream.ConfigureAwait(false);

        var doc = await XDocument.LoadAsync(stream, LoadOptions.None, cancellationToken).ConfigureAwait(false);
        var ns = doc.Root!.GetDefaultNamespace();
        var typesElement = doc.Descendants(ns + "Types").Single();

        var partNames = new HashSet<string>(StringComparer.InvariantCultureIgnoreCase);
        var attrNames = typesElement.Elements(ns + "Override").Select(s => s.Attribute("PartName")?.Value);
        foreach (var partName in attrNames.OfType<string>())
        {
            partNames.Add(partName);
        }

        foreach (var (entry, contentType) in _zipContentsMap)
        {
            cancellationToken.ThrowIfCancellationRequested();

            var entryPath = $"/{entry}";
            if (!partNames.Contains(entryPath))
            {
                var newElement = new XElement(ns + "Override", new XAttribute("ContentType", contentType), new XAttribute("PartName", entryPath));
                typesElement.Add(newElement);
            }
        }

        stream.Seek(0, SeekOrigin.Begin);
        await doc.SaveAsync(stream, SaveOptions.None, cancellationToken).ConfigureAwait(false);
    }

    [CreateSyncVersion]
    private async Task CreateZipEntryAsync(string path, string? contentType, string content, CancellationToken cancellationToken)
    {
        cancellationToken.ThrowIfCancellationRequested();

        var entry = _archive.CreateEntry(path, CompressionLevel.Fastest);

        var zipStream = await entry.OpenAsync(cancellationToken).ConfigureAwait(false);
        await using var disposableZipStream = zipStream.ConfigureAwait(false);

        var writer = new MiniExcelStreamWriter(zipStream, Utf8WithBom, _configuration.BufferSize);
        await using var disposableWriter = writer.ConfigureAwait(false);
        await writer.WriteAsync(content, cancellationToken).ConfigureAwait(false);

        if (!string.IsNullOrEmpty(contentType))
            _zipContentsMap.Add(path, contentType);
    }

    [CreateSyncVersion]
    /* Todo: this method is not very efficient, but workbook.xml is generally a very small file so at the moment it's not worth over-optimizing it.
     Also, consider adding active sheet as one of the editable properties.*/
    internal async Task AlterWorksheetAsync(string sheetName, string? newSheetName, int? newSheetIndex, SheetState? newSheetState, CancellationToken cancellationToken = default)
    {
        if (newSheetName is null && newSheetIndex is null && newSheetState is null)
            return;

        var oldWorkbookEntry = _archive.GetEntry(ExcelFileNames.Workbook)!;
 
        try
        {
            var xmlDoc = await LoadWorkbook().ConfigureAwait(false);

            oldWorkbookEntry.Delete();
            var newWorkbookEntry = _archive.CreateEntry(ExcelFileNames.Workbook, CompressionLevel.Fastest);

            var newZipStream = await newWorkbookEntry.OpenAsync(cancellationToken).ConfigureAwait(false);
            await using var newDisposableZipStream = newZipStream.ConfigureAwait(false);
#if NET
            var writer = XmlWriter.Create(newZipStream, new XmlWriterSettings
            {
#if !SYNC_ONLY
                Async = true
#endif
            });
            await using var disposableWriter = writer.ConfigureAwait(false);
            await xmlDoc.WriteToAsync(writer, CancellationToken.None).ConfigureAwait(false);
#else
            using var writer = XmlWriter.Create(newZipStream, new XmlWriterSettings { Async = false });
            xmlDoc.WriteTo(writer);
#endif
        }
        finally
        {
#if NET10_0_OR_GREATER
            await _archive.DisposeAsync().ConfigureAwait(false);
#else
            _archive.Dispose();
#endif
        }
        return;

        async Task<XDocument> LoadWorkbook()
        {
            var zipStream = await oldWorkbookEntry.OpenAsync(cancellationToken).ConfigureAwait(false);
            await using var disposableZipStream = zipStream.ConfigureAwait(false);

            var workbookDoc = await XDocument.LoadAsync(zipStream, LoadOptions.None, cancellationToken).ConfigureAwait(false);
            var sheetsContainer = workbookDoc.Root?.Element((XNamespace)Schemas.SpreadsheetmlXmlMain + "sheets")!;
            var sheets = sheetsContainer.Elements().ToList();

            if (sheets.Find(s => s.Attribute("name")?.Value.Equals(sheetName, StringComparison.OrdinalIgnoreCase) is true) is not { } sheet)
                throw new InvalidDataException($"Sheet {sheetName} not found");

            if (newSheetName is not null)
            {
                ThrowHelper.ThrowIfInvalidSheetName(newSheetName);
                sheet.SetAttributeValue("name", newSheetName);
            }

            if (newSheetIndex is not null)
            {
                var newIndex = Math.Clamp(newSheetIndex.Value, 0, sheets.Count - 1);
                sheets.Remove(sheet);
                sheets.Insert(newIndex, sheet);

                sheetsContainer.RemoveAll();
                sheetsContainer.Add(sheets);
            }

            if (newSheetState is not null)
            {
                sheet.SetAttributeValue("state", newSheetState switch
                {
                    SheetState.Visible => "visible",
                    SheetState.Hidden => "hidden",
                    SheetState.VeryHidden => "veryHidden",
                    _ => "visible"
                });
            }

            return workbookDoc;
        }
    }
}
