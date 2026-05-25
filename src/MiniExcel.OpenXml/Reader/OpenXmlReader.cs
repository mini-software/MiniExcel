using System.Collections.ObjectModel;
using MiniExcelLib.OpenXml.Styles;
using MiniExcelMapper = MiniExcelLib.Core.Reflection.MiniExcelMapper;
using XmlReaderHelper = MiniExcelLib.OpenXml.Utils.XmlReaderHelper;

namespace MiniExcelLib.OpenXml.Reader;

internal sealed partial class OpenXmlReader : IMiniExcelReader
{
    private static readonly string[] Ns = [Schemas.SpreadsheetmlXmlMain, Schemas.SpreadsheetmlXmlStrictNs];
    private static readonly string[] RelationshiopNs = [Schemas.SpreadsheetmlXmlRelationships, Schemas.SpreadsheetmlXmlStrictRelationships];
    private readonly OpenXmlConfiguration _config;
    
    private List<SheetRecord>? _sheetRecords;
    private OpenXmlStyles? _style;
    private bool _disposed;
    
    internal readonly OpenXmlZip Archive;
    internal IDictionary<int, string> SharedStrings = new Dictionary<int, string>();
    
    private OpenXmlReader(OpenXmlZip archive, IMiniExcelConfiguration? configuration)
    {
        Archive = archive;
        _config = (OpenXmlConfiguration?)configuration ?? OpenXmlConfiguration.Default;
    }

    [CreateSyncVersion]
    internal static async Task<OpenXmlReader> CreateAsync(Stream stream, IMiniExcelConfiguration? configuration, bool leaveOpen = false, CancellationToken cancellationToken = default)
    {
        ThrowHelper.ThrowIfInvalidOpenXml(stream);

        var archive = await OpenXmlZip.CreateAsync(stream, leaveOpen: leaveOpen, cancellationToken: cancellationToken).ConfigureAwait(false);
        var reader = new OpenXmlReader(archive, configuration);
        await reader.SetSharedStringsAsync(cancellationToken).ConfigureAwait(false);

        return reader;
    }
    
    [CreateSyncVersion]
    public IAsyncEnumerable<IDictionary<string, object?>> QueryAsync(bool hasHeaderRow, string? sheetName, string startCell, CancellationToken cancellationToken = default)
    {
        return QueryRangeAsync(hasHeaderRow, sheetName, startCell, "", cancellationToken);
    }

    [CreateSyncVersion]
    public IAsyncEnumerable<T> QueryAsync<T>(string? sheetName, string startCell, bool mapHeaderAsData, CancellationToken cancellationToken = default) where T : class, new()
    {
        sheetName ??= MiniExcelPropertyHelper.GetExcelSheetInfo(typeof(T), _config)?.ExcelSheetName;
        var query = QueryAsync(false, sheetName, startCell, cancellationToken);

        if (!CellReferenceConverter.TryParseCellReference(startCell, out _, out var rowOffset))
            throw new InvalidDataException($"Value {startCell} is not a valid cell reference.");
        
        //Todo: Find a way if possible to remove the 'hasHeader' parameter to check whether or not to include
        // the first row in the result set in favor of modifying the already present 'useHeaderRow' to do the same job          
        return MiniExcelMapper.MapQueryAsync<T>(query, rowOffset, mapHeaderAsData, _config.TrimColumnNames, _config, XmlHelper.DecodeString, cancellationToken);    
    }

    [CreateSyncVersion]
    public IAsyncEnumerable<IDictionary<string, object?>> QueryRangeAsync(bool hasHeaderRow, string? sheetName, string? startCell, string? endCell, CancellationToken cancellationToken = default)
    {
        cancellationToken.ThrowIfCancellationRequested();

        if (!CellReferenceConverter.TryParseCellReference(startCell, out var startColumnIndex, out var startRowIndex))
            throw new InvalidDataException($"Value {startCell} is not a valid cell reference.");
        
        // convert to 0-based
        startColumnIndex--;
        startRowIndex--;

        // endCell is allowed to be empty to query for all rows and columns
        int? endColumnIndex = null;
        int? endRowIndex = null;
        if (!string.IsNullOrWhiteSpace(endCell))
        {
            if (!CellReferenceConverter.TryParseCellReference(endCell, out int cIndex, out int rIndex))
                throw new InvalidDataException($"Value {endCell} is not a valid cell reference.");

            // convert to 0-based
            endColumnIndex = cIndex - 1;
            endRowIndex = rIndex - 1;
        }

        return InternalQueryRangeAsync(hasHeaderRow, sheetName, startRowIndex, startColumnIndex, endRowIndex, endColumnIndex, cancellationToken);
    }

    [CreateSyncVersion]
    public IAsyncEnumerable<T> QueryRangeAsync<T>(string? sheetName, string startCell, string endCell, bool treatHeaderAsData, CancellationToken cancellationToken = default) where T : class, new()
    {
        var query = QueryRangeAsync(false, sheetName, startCell, endCell, cancellationToken);
        if (!CellReferenceConverter.TryParseCellReference(startCell, out _, out var rowOffset))
            throw new InvalidDataException($"Value {startCell} is not a valid cell reference.");
        
        return MiniExcelMapper.MapQueryAsync<T>(query, rowOffset, treatHeaderAsData, _config.TrimColumnNames, _config, XmlHelper.DecodeString, cancellationToken);
    }

    [CreateSyncVersion]
    public IAsyncEnumerable<IDictionary<string, object?>> QueryRangeAsync(bool hasHeaderRow, string? sheetName, int startRowIndex, int startColumnIndex, int? endRowIndex, int? endColumnIndex, CancellationToken cancellationToken = default)
    {
        cancellationToken.ThrowIfCancellationRequested();

        if (startRowIndex <= 0)
            throw new ArgumentOutOfRangeException(nameof(startRowIndex), "Start row index is 1-based and must be greater than 0.");
        if (startColumnIndex <= 0)
            throw new ArgumentOutOfRangeException(nameof(startColumnIndex), "Start column index is 1-based and must be greater than 0.");
        
        // convert to 0-based
        startColumnIndex--;
        startRowIndex--;

        if (endRowIndex.HasValue)
        {
            if (endRowIndex.Value <= 0)
                throw new ArgumentOutOfRangeException(nameof(endRowIndex), "End row index is 1-based and must be greater than 0.");
            
            // convert to 0-based
            endRowIndex--;
        }
        if (endColumnIndex.HasValue)
        {
            if (endColumnIndex.Value > 0)
            {
                // convert to 0-based
                endColumnIndex--;
            }
            else
            {
                throw new ArgumentOutOfRangeException(nameof(endColumnIndex), "End column index is 1-based and must be greater than 0.");
            }
        }

        return InternalQueryRangeAsync(hasHeaderRow, sheetName, startRowIndex, startColumnIndex, endRowIndex, endColumnIndex, cancellationToken);
    }

    [CreateSyncVersion]
    public IAsyncEnumerable<T> QueryRangeAsync<T>(string? sheetName, int startRowIndex, int startColumnIndex, int? endRowIndex, int? endColumnIndex, bool treatHeaderAsData, CancellationToken cancellationToken = default) where T : class, new()
    {
        var query = QueryRangeAsync(false, sheetName, startRowIndex, startColumnIndex, endRowIndex, endColumnIndex, cancellationToken);
        return MiniExcelMapper.MapQueryAsync<T>(query, startRowIndex, treatHeaderAsData, _config.TrimColumnNames, _config, XmlHelper.DecodeString, cancellationToken);
    }

    [CreateSyncVersion]
    private async IAsyncEnumerable<IDictionary<string, object?>> InternalQueryRangeAsync(bool useHeaderRow, string? sheetName, int startRowIndex, int startColumnIndex, int? endRowIndex, int? endColumnIndex, [EnumeratorCancellation] CancellationToken cancellationToken = default)
    {
        cancellationToken.ThrowIfCancellationRequested();

        var xmlSettings = XmlReaderHelper.GetXmlReaderSettings();
        var sheetEntry = GetSheetEntry(sheetName);

        // TODO: need to optimize performance
        // Q. why need 3 times openstream merge one open read? A. no, zipstream can't use position = 0

        MergeCells? mergeCells = null;
        if (_config.FillMergedCells)
        {
            var mergeCellsResult = await TryGetMergeCellsAsync(sheetEntry, cancellationToken).ConfigureAwait(false);
            if (mergeCellsResult.Success)
            {
                mergeCells = mergeCellsResult.MergeCells;
            }
            else
            {
                yield break;
            }
        }

        var maxRowColumnIndexResult = await TryGetMaxRowColumnIndexAsync(sheetEntry, cancellationToken).ConfigureAwait(false);
        if (!maxRowColumnIndexResult.IsSuccess)
            yield break;

        var maxRowIndex = maxRowColumnIndexResult.MaxRowIndex;
        var maxColumnIndex = maxRowColumnIndexResult.MaxColumnIndex;
        var withoutCr = maxRowColumnIndexResult.WithoutCr;

        if (endColumnIndex.HasValue)
        {
            maxColumnIndex = endColumnIndex.Value;
        }

        var sheetStream = await sheetEntry.OpenAsync(cancellationToken).ConfigureAwait(false);
        await using var disposableSheetStream = sheetStream.ConfigureAwait(false);

        using var reader = XmlReader.Create(sheetStream, xmlSettings);

        if (!reader.IsStartElement("worksheet", Ns))
            yield break;
        if (!await reader.ReadFirstContentAsync(cancellationToken).ConfigureAwait(false))
            yield break;

        while (!reader.EOF)
        {
            if (reader.IsStartElement("sheetData", Ns))
            {
                if (!await reader.ReadFirstContentAsync(cancellationToken).ConfigureAwait(false))
                    continue;

                int rowIndex = -1;
                bool isFirstRow = true;
                var headRows = new Dictionary<int, string>();
                while (!reader.EOF)
                {
                    if (reader.IsStartElement("row", Ns))
                    {
                        var nextRowIndex = rowIndex + 1;
                        if (int.TryParse(reader.GetAttribute("r"), out int arValue))
                            rowIndex = arValue - 1; // The row attribute is 1-based
                        else
                            rowIndex++;

                        if (rowIndex < startRowIndex)
                        {
                            if (await reader.ReadFirstContentAsync(cancellationToken).ConfigureAwait(false))
                            {
                                await reader.SkipToNextSiblingAsync(cancellationToken).ConfigureAwait(false);
                            }
                            continue;
                        }

                        if (rowIndex > endRowIndex)
                        {
                            break;
                        }

                        var query = QueryRowAsync(reader, isFirstRow, startRowIndex, nextRowIndex, rowIndex, 
                            startColumnIndex, endColumnIndex, maxColumnIndex, withoutCr, useHeaderRow, headRows, 
                            mergeCells, cancellationToken);

                        await foreach (var row in query.ConfigureAwait(false))
                        {
                            if (isFirstRow)
                            {
                                isFirstRow = false; // for startcell logic
                                if (useHeaderRow)
                                    continue;
                            }

                            yield return row;
                        }
                    }
                    else if (!await reader.SkipContentAsync(cancellationToken).ConfigureAwait(false))
                    {
                        break;
                    }
                }
            }
            else if (!await reader.SkipContentAsync(cancellationToken).ConfigureAwait(false))
            {
                break;
            }
        }
    }

    [CreateSyncVersion]
    private async IAsyncEnumerable<IDictionary<string, object?>> QueryRowAsync(
        XmlReader reader,
        bool isFirstRow,
        int startRowIndex,
        int nextRowIndex,
        int rowIndex,
        int startColumnIndex,
        int? endColumnIndex,
        int maxColumnIndex,
        bool withoutCr,
        bool hasHeaderRow,
        Dictionary<int, string> headRows,
        MergeCells? mergeCells,
        [EnumeratorCancellation] CancellationToken cancellationToken = default)
    {
        // fill empty rows
        if (!_config.IgnoreEmptyRows)
        {
            var expectedRowIndex = isFirstRow ? startRowIndex : nextRowIndex;
            if (startRowIndex <= expectedRowIndex && expectedRowIndex < rowIndex)
            {
                for (int i = expectedRowIndex; i < rowIndex; i++)
                {
                    yield return GetCell(hasHeaderRow, maxColumnIndex, headRows, startColumnIndex);
                }
            }
        }

        // row -> c, must after `if (nextRowIndex < rowIndex)` condition code, eg. The first empty row has no xml element,and the second row xml element is <row r="2"/>
        if (!await reader.ReadFirstContentAsync(cancellationToken).ConfigureAwait(false) && !_config.IgnoreEmptyRows)
        {
            //Fill in case of self closed empty row tag eg. <row r="1"/>
            yield return GetCell(hasHeaderRow, maxColumnIndex, headRows, startColumnIndex);
            yield break;
        }

        var cell = GetCell(hasHeaderRow, maxColumnIndex, headRows, startColumnIndex);
        var columnIndex = withoutCr ? -1 : 0;
        while (!reader.EOF)
        {
            if (reader.IsStartElement("c", Ns))
            {
                var aS = reader.GetAttribute("s");
                var aR = reader.GetAttribute("r");
                var aT = reader.GetAttribute("t");
                var cellAndColumn = await ReadCellAndSetColumnIndexAsync(reader, columnIndex, withoutCr, startColumnIndex, aR, aT, cancellationToken).ConfigureAwait(false);

                var cellValue = cellAndColumn.CellValue;
                columnIndex = cellAndColumn.ColumnIndex;

                if (_config.FillMergedCells)
                {
                    if (mergeCells?.MergesValues.ContainsKey(aR) ?? false)
                    {
                        mergeCells.MergesValues[aR] = cellValue;
                    }
                    else if (mergeCells?.MergesMap.TryGetValue(aR, out var mergeKey) ?? false)
                    {
                        mergeCells.MergesValues.TryGetValue(mergeKey, out cellValue);
                    }
                }

                if (columnIndex < startColumnIndex || columnIndex > endColumnIndex)
                    continue;

                if (!string.IsNullOrEmpty(aS)) // if c with s meaning is custom style need to check type by xl/style.xml
                {
                    int xfIndex = -1;
                    if (int.TryParse(aS, NumberStyles.Any, CultureInfo.InvariantCulture, out var styleIndex))
                        xfIndex = styleIndex;

                    // only when have s attribute then load styles xml data
                    _style ??= await OpenXmlStyles.CreateAsync(Archive, cancellationToken).ConfigureAwait(false);
                    cellValue = _style.ConvertValueByStyleFormat(xfIndex, cellValue);
                }

                SetCellsValueAndHeaders(cellValue, hasHeaderRow, headRows, isFirstRow, cell, columnIndex);
            }
            else if (!await reader.SkipContentAsync(cancellationToken).ConfigureAwait(false))
            {
                break;
            }
        }
        yield return cell;
    }
    
    private ZipArchiveEntry GetSheetEntry(string? sheetName)
    {
        // if sheets count > 1 need to read xl/_rels/workbook.xml.rels
        var sheets = Archive.EntryCollection
            .Where(w => w.FullName.TrimStart('/').StartsWith(ExcelFileNames.WorksheetBase, StringComparison.OrdinalIgnoreCase))
            .ToArray();

        ZipArchiveEntry sheetEntry;
        if (sheetName is not null)
        {
            SetWorkbookRels(Archive.EntryCollection);
            var sheetRecord = _sheetRecords.SingleOrDefault(s => s.Name == sheetName);
            if (sheetRecord is null)
            {
                if (_config.DynamicSheets is null)
                    throw new InvalidOperationException("Please check that parameters sheetName/Index are correct");

                if (_config.DynamicSheets.FirstOrDefault(ds => ds.Key == sheetName) is { } sheetConfig)
                {
                    sheetRecord = _sheetRecords.SingleOrDefault(s => s.Name == sheetConfig.Name);
                }
            }
            sheetEntry = sheets.Single(w => w.FullName.TrimStart('/') == $"xl/{sheetRecord?.Path}" || w.FullName == sheetRecord?.Path?.TrimStart('/'));
        }
        else if (sheets.Length > 1)
        {
            SetWorkbookRels(Archive.EntryCollection);
            var s = _sheetRecords[0];
            sheetEntry = sheets.Single(w => w.FullName.TrimStart('/') == $"xl/{s.Path}" || w.FullName.TrimStart('/') == s.Path?.TrimStart('/'));
        }
        else
        {
            sheetEntry = sheets.Single();
        }

        return sheetEntry;
    }

    private static IDictionary<string, object?> GetCell(bool useHeaderRow, int maxColumnIndex, Dictionary<int, string> headRows, int startColumnIndex)
    {
        return useHeaderRow 
            ? ExpandoHelper.CreateEmptyByHeaders(headRows) 
            : ExpandoHelper.CreateEmptyByIndices(maxColumnIndex, startColumnIndex);
    }

    private static void SetCellsValueAndHeaders(object? cellValue, bool useHeaderRow, Dictionary<int, string> headRows, bool isFirstRow, IDictionary<string, object?> cell, int columnIndex)
    {
        if (!useHeaderRow)
        {
            //if not using First Head then using A,B,C as index
            cell[CellReferenceConverter.GetAlphabeticalIndex(columnIndex)] = cellValue;
            return;
        }

        if (isFirstRow) // for startcell logic
        {
            var cellValueString = cellValue?.ToString();
            if (!string.IsNullOrWhiteSpace(cellValueString))
                headRows.Add(columnIndex, cellValueString);
        }
        else if (headRows.TryGetValue(columnIndex, out var key))
        {
            cell[key] = cellValue;
        }
    }

    [CreateSyncVersion]
    private async Task SetSharedStringsAsync(CancellationToken cancellationToken = default)
    {
        cancellationToken.ThrowIfCancellationRequested();

        if (SharedStrings is { Count: > 0 })
            return;

        if (Archive.GetEntry(ExcelFileNames.SharedStrings) is not { } sharedStringsEntry)
            return;

        var stream = await sharedStringsEntry.OpenAsync(cancellationToken).ConfigureAwait(false);
        await using var disposableStream = stream.ConfigureAwait(false);

        var idx = 0;
        if (_config.EnableSharedStringCache && sharedStringsEntry.Length >= _config.SharedStringCacheSize)
        {
            SharedStrings = new SharedStringsDiskCache(_config.SharedStringCachePath);
            await foreach (var sharedString in XmlReaderHelper.GetSharedStringsAsync(stream, cancellationToken, Ns).ConfigureAwait(false))
            {
                SharedStrings[idx++] = sharedString;
            }
        }
        else if (SharedStrings is { Count: 0 })
        {
            var list = await XmlReaderHelper.GetSharedStringsAsync(stream, cancellationToken, Ns)
                .CreateListAsync(cancellationToken)
                .ConfigureAwait(false);

            SharedStrings = list.ToDictionary(_ => idx++, x => x);
        }
    }

    private void SetWorkbookRels(ReadOnlyCollection<ZipArchiveEntry> entries)
    {
        _sheetRecords ??= GetWorkbookRels(entries);
    }

    [CreateSyncVersion]
    private static async IAsyncEnumerable<SheetRecord> ReadWorkbookAsync(ReadOnlyCollection<ZipArchiveEntry> entries, [EnumeratorCancellation] CancellationToken cancellationToken = default)
    {
        var xmlSettings = XmlReaderHelper.GetXmlReaderSettings();

        var entry = entries.Single(w => w.FullName == ExcelFileNames.Workbook);
        var stream = await entry.OpenAsync(cancellationToken).ConfigureAwait(false);
        await using var disposableStream = stream.ConfigureAwait(false);

        using var reader = XmlReader.Create(stream, xmlSettings);
        
        if (!reader.IsStartElement("workbook", Ns))
            yield break;
        if (!await reader.ReadFirstContentAsync(cancellationToken).ConfigureAwait(false))
            yield break;

        var activeSheetIndex = 0;
        while (!reader.EOF)
        {
            if (reader.IsStartElement("bookViews", Ns))
            {
                if (!await reader.ReadFirstContentAsync(cancellationToken).ConfigureAwait(false))
                    continue;

                while (!reader.EOF)
                {
                    if (reader.IsStartElement("workbookView", Ns))
                    {
                        var activeSheet = reader.GetAttribute("activeTab");
                        if (int.TryParse(activeSheet, out var index))
                        {
                            activeSheetIndex = index;
                        }

                        await reader.SkipAsync()
#if NET
                            .WaitAsync(cancellationToken)
#endif
                            .ConfigureAwait(false);
                    }
                    else if (!await reader.SkipContentAsync(cancellationToken).ConfigureAwait(false))
                    {
                        break;
                    }
                }
            }
            else if (reader.IsStartElement("sheets", Ns))
            {
                if (!await reader.ReadFirstContentAsync(cancellationToken).ConfigureAwait(false))
                    continue;

                var sheetCount = 0;
                while (!reader.EOF)
                {
                    if (reader.IsStartElement("sheet", Ns))
                    {
                        yield return new SheetRecord(
                            reader.GetAttribute("name"),
                            reader.GetAttribute("state"),
                            uint.Parse(reader.GetAttribute("sheetId")),
                            reader.GetAttribute("id", RelationshiopNs),
                            sheetCount == activeSheetIndex
                        );
                        sheetCount++;
                        await reader.SkipAsync()
#if NET
                            .WaitAsync(cancellationToken)
#endif
                            .ConfigureAwait(false);
                    }
                    else if (!await reader.SkipContentAsync(cancellationToken).ConfigureAwait(false))
                    {
                        break;
                    }
                }
            }
            else if (!await reader.SkipContentAsync(cancellationToken).ConfigureAwait(false))
            {
                yield break;
            }
        }
    }

    [CreateSyncVersion]
    internal static async Task<List<SheetRecord>?> GetWorkbookRelsAsync(ReadOnlyCollection<ZipArchiveEntry> entries, CancellationToken cancellationToken = default)
    {
        var sheetRecords = await ReadWorkbookAsync(entries, cancellationToken)
            .CreateListAsync(cancellationToken)
            .ConfigureAwait(false);

        var entry = entries.Single(w => w.FullName == ExcelFileNames.WorkbookRels);
        var stream = await entry.OpenAsync(cancellationToken).ConfigureAwait(false);
        await using var disposableStream = stream.ConfigureAwait(false);

        var readerSettings = XmlReaderHelper.GetXmlReaderSettings();
        using var reader = XmlReader.Create(stream, readerSettings);
        
        if (!XmlReaderHelper.IsStartElement(reader, "Relationships", Schemas.OpenXmlPackageRelationships))
            return null;

        if (!await reader.ReadFirstContentAsync(cancellationToken).ConfigureAwait(false))
            return null;

        while (!reader.EOF)
        {
            if (reader.IsStartElement("Relationship", Schemas.OpenXmlPackageRelationships))
            {
                var rid = reader.GetAttribute("Id");
                foreach (var sheet in sheetRecords.Where(sh => sh.Rid == rid))
                {
                    sheet.Path = reader.GetAttribute("Target");
                    break;
                }

                await reader.SkipAsync()
#if NET
                    .WaitAsync(cancellationToken)
#endif
                    .ConfigureAwait(false);
            }
            else if (!await reader.SkipContentAsync(cancellationToken).ConfigureAwait(false))
            {
                break;
            }
        }

        return sheetRecords;
    }

    internal class CellAndColumn(object? cellValue, int columnIndex)
    {
        public object? CellValue { get; } = cellValue;
        public int ColumnIndex { get; } = columnIndex;
    }

    [CreateSyncVersion]
    private async Task<CellAndColumn> ReadCellAndSetColumnIndexAsync(XmlReader reader, int columnIndex, bool withoutCr, int startColumnIndex, string aR, string aT, CancellationToken cancellationToken = default)
    {
        cancellationToken.ThrowIfCancellationRequested();

        const int xfIndex = -1;
        int newColumnIndex;

        if (withoutCr)
        {
            newColumnIndex = columnIndex + 1;
        }
        else if (CellReferenceConverter.TryParseCellReference(aR, out int referenceColumn, out _))
        {
            //TODO:need to check only need nextColumnIndex or columnIndex
            newColumnIndex = referenceColumn - 1; // ParseReference is 1-based
        }
        else
        {
            newColumnIndex = columnIndex;
        }

        columnIndex = newColumnIndex;

        if (columnIndex < startColumnIndex)
        {
            if (!await reader.ReadFirstContentAsync(cancellationToken).ConfigureAwait(false))
                return new CellAndColumn(null, columnIndex);

            while (!reader.EOF)
            {
                if (!await reader.SkipContentAsync(cancellationToken).ConfigureAwait(false))
                    break;
            }

            return new CellAndColumn(null, columnIndex);
        }

        if (!await reader.ReadFirstContentAsync(cancellationToken).ConfigureAwait(false))
            return new CellAndColumn(null, columnIndex);

        object? value = null;
        while (!reader.EOF)
        {
            if (reader.IsStartElement("v", Ns))
            {
                var rawValue = await reader.ReadElementContentAsStringAsync()
#if NET
                    .WaitAsync(cancellationToken)
#endif
                    .ConfigureAwait(false);
                
                if (!string.IsNullOrEmpty(rawValue))
                    ConvertCellValue(rawValue, aT, xfIndex, out value);
            }
            else if (reader.IsStartElement("is", Ns))
            {
                var rawValue = await reader.ReadStringItemAsync(cancellationToken).ConfigureAwait(false);
                if (!string.IsNullOrEmpty(rawValue))
                    ConvertCellValue(rawValue, aT, xfIndex, out value);
            }
            else if (!await reader.SkipContentAsync(cancellationToken).ConfigureAwait(false))
            {
                break;
            }
        }

        return new CellAndColumn(value, columnIndex);
    }

    private void ConvertCellValue(string rawValue, string aT, int xfIndex, out object? value)
    {
        const NumberStyles style = NumberStyles.Any;
        var invariantCulture = CultureInfo.InvariantCulture;
        value = null;
        
        switch (aT)
        {
            case ExcelDataTypes.SharedString:
                if (int.TryParse(rawValue, style, invariantCulture, out var sstIndex))
                {
                    if (sstIndex >= 0 && sstIndex < SharedStrings?.Count)
                    {
                        value = XmlHelper.DecodeString(SharedStrings[sstIndex]);
                    }
                }
                break;

            case ExcelDataTypes.InlineString or ExcelDataTypes.CalculatedString:
                //TODO: it will unbox,box
                var v = XmlHelper.DecodeString(rawValue);
                value = v;
                if (_config.EnableConvertByteArray)
                {
                    //if str start with "data:image/png;base64," then convert to byte[] https://github.com/mini-software/MiniExcel/issues/318
                    if (v?.StartsWith("@@@fileid@@@,", StringComparison.Ordinal) ?? false)
                    {
                        var path = v[13..];
                        var entry = Archive.GetEntry(path)!;

                        using var stream = entry.Open();
                        using var ms = new MemoryStream((int)entry.Length);

                        stream.CopyTo(ms);
                        value = ms.ToArray();
                    }
                }
                break;

            case ExcelDataTypes.Boolean:
                value = rawValue == "1";
                return;

            case ExcelDataTypes.DateTime:
                value = DateTime.TryParseExact(rawValue, "yyyy-MM-dd", invariantCulture, DateTimeStyles.AllowLeadingWhite | DateTimeStyles.AllowTrailingWhite, out var date)
                    ? date
                    : rawValue;
                return;

            case ExcelDataTypes.Error:
                value = rawValue;
                return;

            default:
                value = double.TryParse(rawValue, style, invariantCulture, out var n) ? n : rawValue;
                return;
        }
    }

    [CreateSyncVersion]
    internal async Task<IList<ExcelRange>> GetDimensionsAsync(CancellationToken cancellationToken = default)
    {
        cancellationToken.ThrowIfCancellationRequested();

        var xmlSettings = XmlReaderHelper.GetXmlReaderSettings();
        var ranges = new List<ExcelRange>();

        var sheets = Archive.EntryCollection.Where(e =>
            e.FullName.TrimStart('/').StartsWith(ExcelFileNames.WorksheetBase, StringComparison.OrdinalIgnoreCase));

        foreach (var sheet in sheets)
        {
            var maxRowIndex = -1;
            var maxColumnIndex = -1;

            string? startCell = null;
            string? endCell = null;

            var withoutCr = false;

            var crSheetStream = await sheet.OpenAsync(cancellationToken).ConfigureAwait(false);
            await using var disposableCrSheetStream = crSheetStream.ConfigureAwait(false);
            using (var reader = XmlReader.Create(crSheetStream, xmlSettings))
            {
                while (await reader.ReadAsync().ConfigureAwait(false))
                {
                    if (reader.IsStartElement("c", Ns))
                    {
                        var r = reader.GetAttribute("r");
                        if (r is not null)
                        {
                            if (CellReferenceConverter.TryParseCellReference(r, out var column, out var row))
                            {
                                column--;
                                row--;
                                maxRowIndex = Math.Max(maxRowIndex, row);
                                maxColumnIndex = Math.Max(maxColumnIndex, column);
                            }
                        }
                        else
                        {
                            withoutCr = true;
                            break;
                        }
                    }

                    else if (reader.IsStartElement("dimension", Ns))
                    {
                        var refAttr = reader.GetAttribute("ref");
                        if (string.IsNullOrEmpty(refAttr))
                            throw new InvalidDataException("No dimension data found for the sheet");

                        var rs = refAttr.Split(':');
                        if (!CellReferenceConverter.TryParseCellReference(rs.Length == 2 ? rs[1] : rs[0], out var col, out var row))
                            throw new InvalidDataException("The dimensions of the sheet are invalid");

                        maxColumnIndex = col;
                        maxRowIndex = row;

                        startCell = rs[0];
                        endCell = rs[1];

                        break;
                    }
                }
            }

            if (withoutCr)
            {
                var sheetStream = await sheet.OpenAsync(cancellationToken).ConfigureAwait(false);
                await using var disposableSheetStream = sheetStream.ConfigureAwait(false);

                using var reader = XmlReader.Create(sheetStream, xmlSettings);
                
                if (!reader.IsStartElement("worksheet", Ns))
                    throw new InvalidDataException("No worksheet data found for the sheet");

                if (!await reader.ReadFirstContentAsync(cancellationToken).ConfigureAwait(false))
                    throw new InvalidOperationException("Excel sheet does not contain any data");

                while (!reader.EOF)
                {
                    if (reader.IsStartElement("sheetData", Ns))
                    {
                        if (!await reader.ReadFirstContentAsync(cancellationToken).ConfigureAwait(false))
                            continue;

                        while (!reader.EOF)
                        {
                            if (reader.IsStartElement("row", Ns))
                            {
                                maxRowIndex++;

                                if (!await reader.ReadFirstContentAsync(cancellationToken).ConfigureAwait(false))
                                    continue;

                                var cellIndex = -1;
                                while (!reader.EOF)
                                {
                                    if (reader.IsStartElement("c", Ns))
                                    {
                                        cellIndex++;
                                        maxColumnIndex = Math.Max(maxColumnIndex, cellIndex);
                                    }

                                    if (!await reader.SkipContentAsync(cancellationToken).ConfigureAwait(false))
                                        break;
                                }
                            }
                            else if (!await reader.SkipContentAsync(cancellationToken).ConfigureAwait(false))
                            {
                                break;
                            }
                        }
                    }
                    else if (!await reader.SkipContentAsync(cancellationToken).ConfigureAwait(false))
                    {
                        break;
                    }
                }
            }

            var range = new ExcelRange(maxRowIndex, maxColumnIndex)
            {
                StartCell = startCell,
                EndCell = endCell
            };
            ranges.Add(range);
        }

        return ranges;
    }

    internal class GetMaxRowColumnIndexResult(bool isSuccess)
    {
        public bool IsSuccess { get; } = isSuccess;
        public bool WithoutCr { get; }
        public int MaxRowIndex { get; } = -1;
        public int MaxColumnIndex { get; } = -1;
        
        public GetMaxRowColumnIndexResult(bool isSuccess, bool withoutCr, int maxRowIndex, int maxColumnIndex)
            : this(isSuccess)
        {
            WithoutCr = withoutCr;
            MaxRowIndex = maxRowIndex;
            MaxColumnIndex = maxColumnIndex;
        }
    }
    
    [CreateSyncVersion]
    internal static async Task<GetMaxRowColumnIndexResult> TryGetMaxRowColumnIndexAsync(ZipArchiveEntry sheetEntry, CancellationToken cancellationToken = default)
    {
        cancellationToken.ThrowIfCancellationRequested();

        var xmlSettings = XmlReaderHelper.GetXmlReaderSettings();

        bool withoutCr = false;
        int maxRowIndex = -1;
        int maxColumnIndex = -1;

        var crSheetStream = await sheetEntry.OpenAsync(cancellationToken).ConfigureAwait(false);
        await using var disposableCrSheetStream  = crSheetStream.ConfigureAwait(false);
        using (var reader = XmlReader.Create(crSheetStream, xmlSettings))
        {
            while (await reader.ReadAsync()
#if NET
                .WaitAsync(cancellationToken)
#endif
                .ConfigureAwait(false))
            {
                if (reader.IsStartElement("c", Ns))
                {
                    var r = reader.GetAttribute("r");
                    if (r is not null)
                    {
                        if (CellReferenceConverter.TryParseCellReference(r, out var column, out var row))
                        {
                            column--;
                            row--;
                            maxRowIndex = Math.Max(maxRowIndex, row);
                            maxColumnIndex = Math.Max(maxColumnIndex, column);
                        }
                    }
                    else
                    {
                        withoutCr = true;
                        break;
                    }
                }
                //this method logic depends on dimension to get maxcolumnIndex, if without dimension then it need to foreach all rows first time to get maxColumn and maxRowColumn
                else if (reader.IsStartElement("dimension", Ns))
                {
                    var refAttr = reader.GetAttribute("ref");
                    if (string.IsNullOrEmpty(refAttr))
                        throw new InvalidDataException("No dimension data found for the sheet");

                    var rs = refAttr.Split(':');

                    // issue : https://github.com/mini-software/MiniExcel/issues/102
                    if (!CellReferenceConverter.TryParseCellReference(rs.Length == 2 ? rs[1] : rs[0], out int cIndex, out int rIndex))
                        throw new InvalidDataException("The dimensions of the sheet are invalid");

                    maxRowIndex = rIndex - 1;
                    maxColumnIndex = cIndex - 1;
                    break;
                }
            }
        }

        if (withoutCr)
        {
            var sheetStream = await sheetEntry.OpenAsync(cancellationToken).ConfigureAwait(false);
            await using var disposableSheetStream = sheetStream.ConfigureAwait(false);

            using var reader = XmlReader.Create(sheetStream, xmlSettings);
            
            if (!reader.IsStartElement("worksheet", Ns))
                return new GetMaxRowColumnIndexResult(false);

            if (!await reader.ReadFirstContentAsync(cancellationToken).ConfigureAwait(false))
                return new GetMaxRowColumnIndexResult(false);

            while (!reader.EOF)
            {
                if (reader.IsStartElement("sheetData", Ns))
                {
                    if (!await reader.ReadFirstContentAsync(cancellationToken).ConfigureAwait(false))
                        continue;

                    while (!reader.EOF)
                    {
                        if (reader.IsStartElement("row", Ns))
                        {
                            maxRowIndex++;

                            if (!await reader.ReadFirstContentAsync(cancellationToken).ConfigureAwait(false))
                                continue;

                            // Cells
                            var cellIndex = -1;
                            while (!reader.EOF)
                            {
                                if (reader.IsStartElement("c", Ns))
                                {
                                    cellIndex++;
                                    maxColumnIndex = Math.Max(maxColumnIndex, cellIndex);
                                }

                                if (!await reader.SkipContentAsync(cancellationToken).ConfigureAwait(false))
                                    break;
                            }
                        }
                        else if (!await reader.SkipContentAsync(cancellationToken).ConfigureAwait(false))
                        {
                            break;
                        }
                    }
                }
                else if (!await reader.SkipContentAsync(cancellationToken).ConfigureAwait(false))
                {
                    break;
                }
            }
        }

        return new GetMaxRowColumnIndexResult(true, withoutCr, maxRowIndex, maxColumnIndex);
    }

    public void Dispose()
    {
        if (_disposed)
            return;

        if (SharedStrings is SharedStringsDiskCache cache)
            cache.Dispose();

        _disposed = true;
    }
}
