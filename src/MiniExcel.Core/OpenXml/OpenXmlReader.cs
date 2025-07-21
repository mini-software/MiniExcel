using System.Collections.ObjectModel;
using MiniExcelLib.Core.Helpers;
using MiniExcelLib.Core.OpenXml.Constants;
using MiniExcelLib.Core.OpenXml.Models;
using MiniExcelLib.Core.OpenXml.Styles;
using MiniExcelLib.Core.OpenXml.Utils;
using MiniExcelLib.Core.OpenXml.Zip;
using MiniExcelLib.Core.Reflection;
using MiniExcelMapper = MiniExcelLib.Core.Reflection.MiniExcelMapper;
using XmlReaderHelper = MiniExcelLib.Core.OpenXml.Utils.XmlReaderHelper;

namespace MiniExcelLib.Core.OpenXml;

internal partial class OpenXmlReader : Abstractions.IMiniExcelReader
{
    private static readonly string[] Ns = [Schemas.SpreadsheetmlXmlns, Schemas.SpreadsheetmlXmlStrictns];
    private static readonly string[] RelationshiopNs = [Schemas.SpreadsheetmlXmlRelationshipns, Schemas.SpreadsheetmlXmlStrictRelationshipns];
    private readonly OpenXmlConfiguration _config;
    
    private List<SheetRecord>? _sheetRecords;
    private OpenXmlStyles? _style;
    private bool _disposed;
    
    internal readonly OpenXmlZip Archive;
    internal IDictionary<int, string>? SharedStrings;
    
    private OpenXmlReader(Stream stream, IMiniExcelConfiguration? configuration)
    {
        Archive = new OpenXmlZip(stream);
        _config = (OpenXmlConfiguration?)configuration ?? OpenXmlConfiguration.Default;
    }

    [CreateSyncVersion]
    internal static async Task<OpenXmlReader> CreateAsync(Stream stream, IMiniExcelConfiguration? configuration, CancellationToken cancellationToken = default)
    {
        ThrowHelper.ThrowIfInvalidOpenXml(stream);
        
        var reader = new OpenXmlReader(stream, configuration);
        await reader.SetSharedStringsAsync(cancellationToken).ConfigureAwait(false);
        return reader;
    }
    
    [CreateSyncVersion]
    public IAsyncEnumerable<IDictionary<string, object?>> QueryAsync(bool useHeaderRow, string? sheetName, string startCell, CancellationToken cancellationToken = default)
    {
        return QueryRangeAsync(useHeaderRow, sheetName, startCell, "", cancellationToken);
    }

    [CreateSyncVersion]
    public IAsyncEnumerable<T> QueryAsync<T>(string? sheetName, string startCell, bool mapHeaderAsData, CancellationToken cancellationToken = default) where T : class, new()
    {
        sheetName ??= ExcelPropertyHelper.GetExcellSheetInfo(typeof(T), _config)?.ExcelSheetName;

        //Todo: Find a way if possible to remove the 'hasHeader' parameter to check whether or not to include
        // the first row in the result set in favor of modifying the already present 'useHeaderRow' to do the same job          
        return MiniExcelMapper.MapQueryAsync<T>(QueryAsync(false, sheetName, startCell, cancellationToken), startCell, mapHeaderAsData, _config.TrimColumnNames, _config, cancellationToken);    
    }

    [CreateSyncVersion]
    public IAsyncEnumerable<IDictionary<string, object?>> QueryRangeAsync(bool useHeaderRow, string? sheetName, string startCell, string endCell, CancellationToken cancellationToken = default)
    {
        cancellationToken.ThrowIfCancellationRequested();

        if (!ReferenceHelper.ParseReference(startCell, out var startColumnIndex, out var startRowIndex))
            throw new InvalidDataException($"Value {startCell} is not a valid cell reference.");
        
        // convert to 0-based
        startColumnIndex--;
        startRowIndex--;

        // endCell is allowed to be empty to query for all rows and columns
        int? endColumnIndex = null;
        int? endRowIndex = null;
        if (!string.IsNullOrWhiteSpace(endCell))
        {
            if (!ReferenceHelper.ParseReference(endCell, out int cIndex, out int rIndex))
                throw new InvalidDataException($"Value {endCell} is not a valid cell reference.");

            // convert to 0-based
            endColumnIndex = cIndex - 1;
            endRowIndex = rIndex - 1;
        }

        return InternalQueryRangeAsync(useHeaderRow, sheetName, startRowIndex, startColumnIndex, endRowIndex, endColumnIndex, cancellationToken);
    }

    [CreateSyncVersion]
    public IAsyncEnumerable<T> QueryRangeAsync<T>(string? sheetName, string startCell, string endCell, bool treatHeaderAsData, CancellationToken cancellationToken = default) where T : class, new()
    {
        return MiniExcelMapper.MapQueryAsync<T>(QueryRangeAsync(false, sheetName, startCell, endCell, cancellationToken), startCell, treatHeaderAsData, _config.TrimColumnNames, _config, cancellationToken);
    }

    [CreateSyncVersion]
    public IAsyncEnumerable<IDictionary<string, object?>> QueryRangeAsync(bool useHeaderRow, string? sheetName, int startRowIndex, int startColumnIndex, int? endRowIndex, int? endColumnIndex, CancellationToken cancellationToken = default)
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

        return InternalQueryRangeAsync(useHeaderRow, sheetName, startRowIndex, startColumnIndex, endRowIndex, endColumnIndex, cancellationToken);
    }

    [CreateSyncVersion]
    public IAsyncEnumerable<T> QueryRangeAsync<T>(string? sheetName, int startRowIndex, int startColumnIndex, int? endRowIndex, int? endColumnIndex, bool treatHeaderAsData, CancellationToken cancellationToken = default) where T : class, new()
    {
        return MiniExcelMapper.MapQueryAsync<T>(QueryRangeAsync(false, sheetName, startRowIndex, startColumnIndex, endRowIndex, endColumnIndex, cancellationToken), ReferenceHelper.ConvertCoordinatesToCell(startColumnIndex, startRowIndex), treatHeaderAsData, _config.TrimColumnNames, _config, cancellationToken);
    }

    [CreateSyncVersion]
    internal async IAsyncEnumerable<IDictionary<string, object?>> InternalQueryRangeAsync(bool useHeaderRow, string? sheetName, int startRowIndex, int startColumnIndex, int? endRowIndex, int? endColumnIndex, [EnumeratorCancellation] CancellationToken cancellationToken = default)
    {
        cancellationToken.ThrowIfCancellationRequested();

        var xmlSettings = XmlReaderHelper.GetXmlReaderSettings(
#if SYNC_ONLY
            false
#else
            true
#endif
        );

        var sheetEntry = GetSheetEntry(sheetName);

        // TODO: need to optimize performance
        // Q. why need 3 times openstream merge one open read? A. no, zipstream can't use position = 0

        var mergeCellsContext = new MergeCellsContext();
        if (_config.FillMergedCells && !await TryGetMergeCellsAsync(sheetEntry, mergeCellsContext, cancellationToken).ConfigureAwait(false))
            yield break;

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

#if NET10_0_OR_GREATER
        using var sheetStream = await sheetEntry.OpenAsync(cancellationToken).ConfigureAwait(false);
#else
        using var sheetStream = sheetEntry.Open();
#endif
        using var reader = XmlReader.Create(sheetStream, xmlSettings);

        if (!XmlReaderHelper.IsStartElement(reader, "worksheet", Ns))
            yield break;
        if (!await XmlReaderHelper.ReadFirstContentAsync(reader, cancellationToken).ConfigureAwait(false))
            yield break;

        while (!reader.EOF)
        {
            if (XmlReaderHelper.IsStartElement(reader, "sheetData", Ns))
            {
                if (!await XmlReaderHelper.ReadFirstContentAsync(reader, cancellationToken).ConfigureAwait(false))
                    continue;

                int rowIndex = -1;
                bool isFirstRow = true;
                var headRows = new Dictionary<int, string>();
                while (!reader.EOF)
                {
                    if (XmlReaderHelper.IsStartElement(reader, "row", Ns))
                    {
                        var nextRowIndex = rowIndex + 1;
                        if (int.TryParse(reader.GetAttribute("r"), out int arValue))
                            rowIndex = arValue - 1; // The row attribute is 1-based
                        else
                            rowIndex++;

                        if (rowIndex < startRowIndex)
                        {
                            await XmlReaderHelper.ReadFirstContentAsync(reader, cancellationToken)
                                .ConfigureAwait(false);
                            await XmlReaderHelper.SkipToNextSameLevelDomAsync(reader, cancellationToken)
                                .ConfigureAwait(false);
                            continue;
                        }

                        if (rowIndex > endRowIndex)
                        {
                            break;
                        }

                        await foreach (var row in QueryRowAsync(reader, isFirstRow, startRowIndex, nextRowIndex,
                                               rowIndex, startColumnIndex, endColumnIndex, maxColumnIndex,
                                               withoutCr,
                                               useHeaderRow, headRows, mergeCellsContext.MergeCells,
                                               cancellationToken)
                                           .ConfigureAwait(false))
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
                    else if (!await XmlReaderHelper.SkipContentAsync(reader, cancellationToken)
                                 .ConfigureAwait(false))
                    {
                        break;
                    }
                }
            }
            else if (!await XmlReaderHelper.SkipContentAsync(reader, cancellationToken).ConfigureAwait(false))
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
        bool useHeaderRow,
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
                    yield return GetCell(useHeaderRow, maxColumnIndex, headRows, startColumnIndex);
                }
            }
        }

        // row -> c, must after `if (nextRowIndex < rowIndex)` condition code, eg. The first empty row has no xml element,and the second row xml element is <row r="2"/>
        if (!await XmlReaderHelper.ReadFirstContentAsync(reader, cancellationToken).ConfigureAwait(false) && !_config.IgnoreEmptyRows)
        {
            //Fill in case of self closed empty row tag eg. <row r="1"/>
            yield return GetCell(useHeaderRow, maxColumnIndex, headRows, startColumnIndex);
            yield break;
        }

        var cell = GetCell(useHeaderRow, maxColumnIndex, headRows, startColumnIndex);
        var columnIndex = withoutCr ? -1 : 0;
        while (!reader.EOF)
        {
            if (XmlReaderHelper.IsStartElement(reader, "c", Ns))
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
                    _style ??= new OpenXmlStyles(Archive);
                    cellValue = _style.ConvertValueByStyleFormat(xfIndex, cellValue);
                }

                SetCellsValueAndHeaders(cellValue, useHeaderRow, headRows, isFirstRow, cell, columnIndex);
            }
            else if (!await XmlReaderHelper.SkipContentAsync(reader, cancellationToken).ConfigureAwait(false))
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
            .Where(w => w.FullName.StartsWith("xl/worksheets/sheet", StringComparison.OrdinalIgnoreCase) || 
                        w.FullName.StartsWith("/xl/worksheets/sheet", StringComparison.OrdinalIgnoreCase))
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

                var sheetConfig = _config.DynamicSheets.FirstOrDefault(ds => ds.Key == sheetName);
                if (sheetConfig is not null)
                {
                    sheetRecord = _sheetRecords.SingleOrDefault(s => s.Name == sheetConfig.Name);
                }
            }
            sheetEntry = sheets.Single(w => w.FullName == $"xl/{sheetRecord.Path}" || 
                                            w.FullName == $"/xl/{sheetRecord.Path}" || 
                                            w.FullName == sheetRecord.Path || 
                                            $"/{w.FullName}" == sheetRecord.Path);
        }
        else if (sheets.Length > 1)
        {
            SetWorkbookRels(Archive.EntryCollection);
            var s = _sheetRecords[0];
            sheetEntry = sheets.Single(w => w.FullName == $"xl/{s.Path}" || 
                                            w.FullName == $"/xl/{s.Path}" || 
                                            w.FullName.TrimStart('/') == s.Path.TrimStart('/'));
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
            ? CustomPropertyHelper.GetEmptyExpandoObject(headRows) 
            : CustomPropertyHelper.GetEmptyExpandoObject(maxColumnIndex, startColumnIndex);
    }

    private static void SetCellsValueAndHeaders(object? cellValue, bool useHeaderRow, Dictionary<int, string?> headRows, bool isFirstRow, IDictionary<string, object?> cell, int columnIndex)
    {
        if (!useHeaderRow)
        {
            //if not using First Head then using A,B,C as index
            cell[ColumnHelper.GetAlphabetColumnName(columnIndex)] = cellValue;
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

        if (SharedStrings is not null)
            return;
        
        var sharedStringsEntry = Archive.GetEntry("xl/sharedStrings.xml");
        if (sharedStringsEntry is null)
            return;
        
        var idx = 0;
#if NET10_0_OR_GREATER
        using var stream = await sharedStringsEntry.OpenAsync(cancellationToken).ConfigureAwait(false);
#else
        using var stream = sharedStringsEntry.Open();
#endif
        if (_config.EnableSharedStringCache && sharedStringsEntry.Length >= _config.SharedStringCacheSize)
        {
            SharedStrings = new SharedStringsDiskCache();
            await foreach (var sharedString in XmlReaderHelper.GetSharedStringsAsync(stream, cancellationToken, Ns).ConfigureAwait(false))
            {
                SharedStrings[idx++] = sharedString;
            }
        }
        else if (SharedStrings is null)
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
    internal static async IAsyncEnumerable<SheetRecord> ReadWorkbookAsync(ReadOnlyCollection<ZipArchiveEntry> entries, [EnumeratorCancellation] CancellationToken cancellationToken = default)
    {
        var xmlSettings = XmlReaderHelper.GetXmlReaderSettings(
#if SYNC_ONLY
            false
#else
            true
#endif
        );

        var entry = entries.Single(w => w.FullName == "xl/workbook.xml");
#if NET10_0_OR_GREATER
        using var stream = await entry.OpenAsync(cancellationToken).ConfigureAwait(false);
#else
        using var stream = entry.Open();
#endif
        using var reader = XmlReader.Create(stream, xmlSettings);
        
        if (!XmlReaderHelper.IsStartElement(reader, "workbook", Ns))
            yield break;
        if (!await XmlReaderHelper.ReadFirstContentAsync(reader, cancellationToken).ConfigureAwait(false))
            yield break;

        var activeSheetIndex = 0;
        while (!reader.EOF)
        {
            if (XmlReaderHelper.IsStartElement(reader, "bookViews", Ns))
            {
                if (!await XmlReaderHelper.ReadFirstContentAsync(reader, cancellationToken).ConfigureAwait(false))
                    continue;

                while (!reader.EOF)
                {
                    if (XmlReaderHelper.IsStartElement(reader, "workbookView", Ns))
                    {
                        var activeSheet = reader.GetAttribute("activeTab");
                        if (int.TryParse(activeSheet, out var index))
                        {
                            activeSheetIndex = index;
                        }

                        await reader.SkipAsync()
#if NET6_0_OR_GREATER
                            .WaitAsync(cancellationToken)
#endif
                            .ConfigureAwait(false);
                    }
                    else if (!await XmlReaderHelper.SkipContentAsync(reader, cancellationToken).ConfigureAwait(false))
                    {
                        break;
                    }
                }
            }
            else if (XmlReaderHelper.IsStartElement(reader, "sheets", Ns))
            {
                if (!await XmlReaderHelper.ReadFirstContentAsync(reader, cancellationToken).ConfigureAwait(false))
                    continue;

                var sheetCount = 0;
                while (!reader.EOF)
                {
                    if (XmlReaderHelper.IsStartElement(reader, "sheet", Ns))
                    {
                        yield return new SheetRecord(
                            reader.GetAttribute("name"),
                            reader.GetAttribute("state"),
                            uint.Parse(reader.GetAttribute("sheetId")),
                            XmlReaderHelper.GetAttribute(reader, "id", RelationshiopNs),
                            sheetCount == activeSheetIndex
                        );
                        sheetCount++;
                        await reader.SkipAsync()
#if NET6_0_OR_GREATER
                                .WaitAsync(cancellationToken)
#endif
                            .ConfigureAwait(false);
                    }
                    else if (!await XmlReaderHelper.SkipContentAsync(reader, cancellationToken).ConfigureAwait(false))
                    {
                        break;
                    }
                }
            }
            else if (!await XmlReaderHelper.SkipContentAsync(reader, cancellationToken).ConfigureAwait(false))
            {
                yield break;
            }
        }
    }

    [CreateSyncVersion]
    internal async Task<List<SheetRecord>?> GetWorkbookRelsAsync(ReadOnlyCollection<ZipArchiveEntry> entries, CancellationToken cancellationToken = default)
    {
        var xmlSettings = XmlReaderHelper.GetXmlReaderSettings(
#if SYNC_ONLY
            false
#else
            true
#endif
        );

        var sheetRecords = await ReadWorkbookAsync(entries, cancellationToken)
            .CreateListAsync(cancellationToken)
            .ConfigureAwait(false);

        var entry = entries.Single(w => w.FullName == "xl/_rels/workbook.xml.rels");
#if NET10_0_OR_GREATER
        using var stream = await entry.OpenAsync(cancellationToken).ConfigureAwait(false);
#else
        using var stream = entry.Open();
#endif
        using var reader = XmlReader.Create(stream, xmlSettings);
        
        if (!XmlReaderHelper.IsStartElement(reader, "Relationships", "http://schemas.openxmlformats.org/package/2006/relationships"))
            return null;
        if (!await XmlReaderHelper.ReadFirstContentAsync(reader, cancellationToken).ConfigureAwait(false))
            return null;

        while (!reader.EOF)
        {
            if (XmlReaderHelper.IsStartElement(reader, "Relationship", "http://schemas.openxmlformats.org/package/2006/relationships"))
            {
                var rid = reader.GetAttribute("Id");
                foreach (var sheet in sheetRecords.Where(sh => sh.Rid == rid))
                {
                    sheet.Path = reader.GetAttribute("Target");
                    break;
                }

                await reader.SkipAsync()
#if NET6_0_OR_GREATER
                    .WaitAsync(cancellationToken)
#endif
                    .ConfigureAwait(false);
            }
            else if (!await XmlReaderHelper.SkipContentAsync(reader, cancellationToken).ConfigureAwait(false))
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
        else if (ReferenceHelper.ParseReference(aR, out int referenceColumn, out _))
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
            if (!await XmlReaderHelper.ReadFirstContentAsync(reader, cancellationToken).ConfigureAwait(false))
                return new CellAndColumn(null, columnIndex);

            while (!reader.EOF)
            {
                if (!await XmlReaderHelper.SkipContentAsync(reader, cancellationToken).ConfigureAwait(false))
                    break;
            }

            return new CellAndColumn(null, columnIndex);
        }

        if (!await XmlReaderHelper.ReadFirstContentAsync(reader, cancellationToken).ConfigureAwait(false))
            return new CellAndColumn(null, columnIndex);

        object? value = null;
        while (!reader.EOF)
        {
            if (XmlReaderHelper.IsStartElement(reader, "v", Ns))
            {
                var rawValue = await reader.ReadElementContentAsStringAsync()
#if NET6_0_OR_GREATER
                    .WaitAsync(cancellationToken)
#endif
                    .ConfigureAwait(false);
                
                if (!string.IsNullOrEmpty(rawValue))
                    ConvertCellValue(rawValue, aT, xfIndex, out value);
            }
            else if (XmlReaderHelper.IsStartElement(reader, "is", Ns))
            {
                var rawValue = await XmlReaderHelper.ReadStringItemAsync(reader, cancellationToken).ConfigureAwait(false);
                if (!string.IsNullOrEmpty(rawValue))
                    ConvertCellValue(rawValue, aT, xfIndex, out value);
            }
            else if (!await XmlReaderHelper.SkipContentAsync(reader, cancellationToken).ConfigureAwait(false))
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

        switch (aT)
        {
            case "s":
                if (int.TryParse(rawValue, style, invariantCulture, out var sstIndex))
                {
                    if (sstIndex >= 0 && sstIndex < SharedStrings?.Count)
                    {
                        //value = Helpers.ConvertEscapeChars(_SharedStrings[sstIndex]);
                        value = XmlHelper.DecodeString(SharedStrings[sstIndex]);
                        return;
                    }
                }
                value = null;
                return;

            case "inlineStr":
            case "str":
                //TODO: it will unbox,box
                var v = XmlHelper.DecodeString(rawValue);
                if (_config.EnableConvertByteArray)
                {
                    //if str start with "data:image/png;base64," then convert to byte[] https://github.com/mini-software/MiniExcel/issues/318
                    if (v is not null && v.StartsWith("@@@fileid@@@,", StringComparison.Ordinal))
                    {
                        var path = v[13..];
                        var entry = Archive.GetEntry(path);
                        var bytes = new byte[entry.Length];

                        using (var stream = entry.Open())
                        using (var ms = new MemoryStream(bytes))
                        {
                            stream.CopyTo(ms);
                        }
                        value = bytes;
                    }
                    else
                    {
                        value = v;
                    }
                }
                else
                {
                    value = v;
                }
                return;

            case "b":
                value = rawValue == "1";
                return;

            case "d":
                if (DateTime.TryParseExact(rawValue, "yyyy-MM-dd", invariantCulture, DateTimeStyles.AllowLeadingWhite | DateTimeStyles.AllowTrailingWhite, out var date))
                {
                    value = date;
                    return;
                }

                value = rawValue;
                return;

            case "e":
                value = rawValue;
                return;

            default:
                if (double.TryParse(rawValue, style, invariantCulture, out var n))
                {
                    value = n;
                    return;
                }

                value = rawValue;
                return;
        }
    }

    [CreateSyncVersion]
    internal async Task<IList<ExcelRange>> GetDimensionsAsync(CancellationToken cancellationToken = default)
    {
        cancellationToken.ThrowIfCancellationRequested();

        var xmlSettings = XmlReaderHelper.GetXmlReaderSettings(
#if SYNC_ONLY
            false
#else
            true
#endif
        );

        var ranges = new List<ExcelRange>();

        var sheets = Archive.EntryCollection.Where(e =>
            e.FullName.StartsWith("xl/worksheets/sheet", StringComparison.OrdinalIgnoreCase) ||
            e.FullName.StartsWith("/xl/worksheets/sheet", StringComparison.OrdinalIgnoreCase));

        foreach (var sheet in sheets)
        {
            var maxRowIndex = -1;
            var maxColumnIndex = -1;

            string? startCell = null;
            string? endCell = null;

            var withoutCr = false;

#if NET10_0_OR_GREATER
            using (var sheetStream = await sheet.OpenAsync(cancellationToken).ConfigureAwait(false))
#else
            using (var sheetStream = sheet.Open())
#endif
            using (var reader = XmlReader.Create(sheetStream, xmlSettings))
            {
                while (await reader.ReadAsync().ConfigureAwait(false))
                {
                    if (XmlReaderHelper.IsStartElement(reader, "c", Ns))
                    {
                        var r = reader.GetAttribute("r");
                        if (r is not null)
                        {
                            if (ReferenceHelper.ParseReference(r, out var column, out var row))
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

                    else if (XmlReaderHelper.IsStartElement(reader, "dimension", Ns))
                    {
                        var refAttr = reader.GetAttribute("ref");
                        if (string.IsNullOrEmpty(refAttr))
                            throw new InvalidDataException("No dimension data found for the sheet");

                        var rs = refAttr.Split(':');
                        if (!ReferenceHelper.ParseReference(rs.Length == 2 ? rs[1] : rs[0], out var col, out var row))
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
#if NET10_0_OR_GREATER
                using var sheetStream = await sheet.OpenAsync(cancellationToken).ConfigureAwait(false);
#else
                using var sheetStream = sheet.Open();
#endif
                using var reader = XmlReader.Create(sheetStream, xmlSettings);
                
                if (!XmlReaderHelper.IsStartElement(reader, "worksheet", Ns))
                    throw new InvalidDataException("No worksheet data found for the sheet");

                if (!await XmlReaderHelper.ReadFirstContentAsync(reader, cancellationToken).ConfigureAwait(false))
                    throw new InvalidOperationException("Excel sheet does not contain any data");

                while (!reader.EOF)
                {
                    if (XmlReaderHelper.IsStartElement(reader, "sheetData", Ns))
                    {
                        if (!await XmlReaderHelper.ReadFirstContentAsync(reader, cancellationToken).ConfigureAwait(false))
                            continue;

                        while (!reader.EOF)
                        {
                            if (XmlReaderHelper.IsStartElement(reader, "row", Ns))
                            {
                                maxRowIndex++;

                                if (!await XmlReaderHelper.ReadFirstContentAsync(reader, cancellationToken).ConfigureAwait(false))
                                    continue;

                                var cellIndex = -1;
                                while (!reader.EOF)
                                {
                                    if (XmlReaderHelper.IsStartElement(reader, "c", Ns))
                                    {
                                        cellIndex++;
                                        maxColumnIndex = Math.Max(maxColumnIndex, cellIndex);
                                    }

                                    if (!await XmlReaderHelper.SkipContentAsync(reader, cancellationToken).ConfigureAwait(false))
                                        break;
                                }
                            }
                            else if (!await XmlReaderHelper.SkipContentAsync(reader, cancellationToken).ConfigureAwait(false))
                            {
                                break;
                            }
                        }
                    }
                    else if (!await XmlReaderHelper.SkipContentAsync(reader, cancellationToken).ConfigureAwait(false))
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

        var xmlSettings = XmlReaderHelper.GetXmlReaderSettings(
#if SYNC_ONLY
            false
#else
            true
#endif
        );

        bool withoutCr = false;
        int maxRowIndex = -1;
        int maxColumnIndex = -1;
#if NET10_0_OR_GREATER
        using (var sheetStream = await sheetEntry.OpenAsync(cancellationToken).ConfigureAwait(false))
#else
        using (var sheetStream = sheetEntry.Open())
#endif
        using (var reader = XmlReader.Create(sheetStream, xmlSettings))
        {
            while (await reader.ReadAsync()
#if NET6_0_OR_GREATER
                .WaitAsync(cancellationToken)
#endif
                .ConfigureAwait(false))
            {
                if (XmlReaderHelper.IsStartElement(reader, "c", Ns))
                {
                    var r = reader.GetAttribute("r");
                    if (r is not null)
                    {
                        if (ReferenceHelper.ParseReference(r, out var column, out var row))
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
                else if (XmlReaderHelper.IsStartElement(reader, "dimension", Ns))
                {
                    var refAttr = reader.GetAttribute("ref");
                    if (string.IsNullOrEmpty(refAttr))
                        throw new InvalidDataException("No dimension data found for the sheet");

                    var rs = refAttr.Split(':');

                    // issue : https://github.com/mini-software/MiniExcel/issues/102
                    if (!ReferenceHelper.ParseReference(rs.Length == 2 ? rs[1] : rs[0], out int cIndex, out int rIndex))
                        throw new InvalidDataException("The dimensions of the sheet are invalid");

                    maxRowIndex = rIndex - 1;
                    maxColumnIndex = cIndex - 1;
                    break;
                }
            }
        }

        if (withoutCr)
        {
#if NET10_0_OR_GREATER
            using var sheetStream = await sheetEntry.OpenAsync(cancellationToken).ConfigureAwait(false);
#else
            using var sheetStream = sheetEntry.Open();
#endif
            using var reader = XmlReader.Create(sheetStream, xmlSettings);
            
            if (!XmlReaderHelper.IsStartElement(reader, "worksheet", Ns))
                return new GetMaxRowColumnIndexResult(false);
            if (!await XmlReaderHelper.ReadFirstContentAsync(reader, cancellationToken).ConfigureAwait(false))
                return new GetMaxRowColumnIndexResult(false);

            while (!reader.EOF)
            {
                if (XmlReaderHelper.IsStartElement(reader, "sheetData", Ns))
                {
                    if (!await XmlReaderHelper.ReadFirstContentAsync(reader, cancellationToken).ConfigureAwait(false))
                        continue;

                    while (!reader.EOF)
                    {
                        if (XmlReaderHelper.IsStartElement(reader, "row", Ns))
                        {
                            maxRowIndex++;

                            if (!await XmlReaderHelper.ReadFirstContentAsync(reader, cancellationToken).ConfigureAwait(false))
                                continue;

                            // Cells
                            var cellIndex = -1;
                            while (!reader.EOF)
                            {
                                if (XmlReaderHelper.IsStartElement(reader, "c", Ns))
                                {
                                    cellIndex++;
                                    maxColumnIndex = Math.Max(maxColumnIndex, cellIndex);
                                }

                                if (!await XmlReaderHelper.SkipContentAsync(reader, cancellationToken).ConfigureAwait(false))
                                    break;
                            }
                        }
                        else if (!await XmlReaderHelper.SkipContentAsync(reader, cancellationToken).ConfigureAwait(false))
                        {
                            break;
                        }
                    }
                }
                else if (!await XmlReaderHelper.SkipContentAsync(reader, cancellationToken).ConfigureAwait(false))
                {
                    break;
                }
            }
        }

        return new GetMaxRowColumnIndexResult(true, withoutCr, maxRowIndex, maxColumnIndex);
    }

    internal class MergeCellsContext
    {
        public MergeCells? MergeCells { get; set; }
    }
    
    
    [CreateSyncVersion]
    internal static async Task<bool> TryGetMergeCellsAsync(ZipArchiveEntry sheetEntry, MergeCellsContext mergeCellsContext, CancellationToken cancellationToken = default)
    {
        cancellationToken.ThrowIfCancellationRequested();

        var xmlSettings = XmlReaderHelper.GetXmlReaderSettings(
#if SYNC_ONLY
            false
#else
            true
#endif
        );
        var mergeCells = new MergeCells();

#if NET10_0_OR_GREATER
        using var sheetStream = await sheetEntry.OpenAsync(cancellationToken).ConfigureAwait(false);
#else
        using var sheetStream = sheetEntry.Open();
#endif
        using var reader = XmlReader.Create(sheetStream, xmlSettings);
        
        if (!XmlReaderHelper.IsStartElement(reader, "worksheet", Ns))
            return false;
        
        while (await reader.ReadAsync().ConfigureAwait(false))
        {
            if (!XmlReaderHelper.IsStartElement(reader, "mergeCells", Ns))
                continue;

            if (!await XmlReaderHelper.ReadFirstContentAsync(reader, cancellationToken).ConfigureAwait(false))
                return false;

            while (!reader.EOF)
            {
                if (XmlReaderHelper.IsStartElement(reader, "mergeCell", Ns))
                {
                    var refAttr = reader.GetAttribute("ref");
                    var refs = refAttr.Split(':');
                    if (refs.Length == 1)
                        continue;

                    ReferenceHelper.ParseReference(refs[0], out var x1, out var y1);
                    ReferenceHelper.ParseReference(refs[1], out var x2, out var y2);

                    mergeCells.MergesValues.Add(refs[0], null);

                    // foreach range
                    var isFirst = true;
                    for (int x = x1; x <= x2; x++)
                    {
                        for (int y = y1; y <= y2; y++)
                        {
                            if (!isFirst)
                                mergeCells.MergesMap.Add(ReferenceHelper.ConvertCoordinatesToCell(x, y), refs[0]);
                            isFirst = false;
                        }
                    }

                    await XmlReaderHelper.SkipContentAsync(reader, cancellationToken).ConfigureAwait(false);
                }
                else if (!await XmlReaderHelper.SkipContentAsync(reader, cancellationToken).ConfigureAwait(false))
                {
                    break;
                }
            }
        }

        mergeCellsContext.MergeCells = mergeCells;
        return true;
    }

    ~OpenXmlReader()
    {
        Dispose(false);
    }

    public void Dispose()
    {
        Dispose(true);
        GC.SuppressFinalize(this);
    }

    protected virtual void Dispose(bool disposing)
    {
        if (!_disposed)
        {
            if (disposing)
            {
                if (SharedStrings is SharedStringsDiskCache cache)
                {
                    cache.Dispose();
                }
            }

            _disposed = true;
        }
    }
}