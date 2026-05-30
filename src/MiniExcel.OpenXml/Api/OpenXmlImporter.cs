// ReSharper disable once CheckNamespace
namespace MiniExcelLib.OpenXml;

public sealed partial class OpenXmlImporter
{
    internal OpenXmlImporter() { }

    #region Query

    /// <summary>
    /// Queries an Excel document using a strongly-typed class model.
    /// </summary>
    /// <typeparam name="T">The class type to map each row to. Must have a parameterless constructor.</typeparam>
    /// <param name="path">The path to the Excel document.</param>
    /// <param name="sheetName">The name of the worsksheet to query. If not specified, the first sheet is used.</param>
    /// <param name="startCell">The starting cell reference (e.g., "C2"). Default is "A1".</param>
    /// <param name="treatHeaderAsData">If true, the first row is treated as data. If false (default), the first row is used as headers.</param>
    /// <param name="configuration">Optional configuration settings.</param>
    /// <param name="cancellationToken">A token to cancel the asynchronous operation.</param>
    [CreateSyncVersion]
    public async IAsyncEnumerable<T> QueryAsync<T>(string path, string? sheetName = null,
        string startCell = "A1", bool treatHeaderAsData = false, OpenXmlConfiguration? configuration = null,
        [EnumeratorCancellation] CancellationToken cancellationToken = default) where T : class, new()
    {
        var stream = FileHelper.OpenSharedRead(path);
        await using var disposableStream = stream.ConfigureAwait(false); 

        var query = QueryAsync<T>(stream, sheetName, startCell, treatHeaderAsData, configuration, false, cancellationToken);
        await foreach (var item in query.ConfigureAwait(false))
            yield return item; 
    }

    /// <summary>
    /// Queries an Excel document using a strongly-typed class model.
    /// </summary>
    /// <typeparam name="T">The class type to map each row to. Must have a parameterless constructor.</typeparam>
    /// <param name="stream">The stream containing the Excel file data. The stream position is not reset after reading.</param>
    /// <param name="sheetName">The name of the worsksheet to query. If not specified, the first sheet is used.</param>
    /// <param name="startCell">The starting cell reference (e.g., "C2"). Default is "A1".</param>
    /// <param name="treatHeaderAsData">If true, the first row is treated as data. If false (default), the first row is used as headers.</param>
    /// <param name="configuration">Optional configuration settings.</param>
    /// <param name="leaveOpen">True to leave the stream open after the query is completed, otherwise false.</param>
    /// <param name="cancellationToken">A token to cancel the asynchronous operation.</param>
    [CreateSyncVersion]
    public async IAsyncEnumerable<T> QueryAsync<T>(Stream stream, string? sheetName = null,
        string startCell = "A1", bool treatHeaderAsData = false, OpenXmlConfiguration? configuration = null,
        bool leaveOpen = false, [EnumeratorCancellation] CancellationToken cancellationToken = default) where T : class, new()
    {
        using var reader = await OpenXmlReader.CreateAsync(stream, configuration, leaveOpen, cancellationToken).ConfigureAwait(false);
        await foreach (var item in reader.QueryAsync<T>(sheetName, startCell, treatHeaderAsData, cancellationToken).ConfigureAwait(false))
            yield return item;
    }

    /// <summary>
    /// Queries an Excel document and returns dynamic objects representing each row.
    /// </summary>
    /// <param name="path">The path to the OpenXml document.</param>
    /// <param name="hasHeaderRow">If true, the first row is used as column headers for the dynamic object properties. Default is false.</param>
    /// <param name="sheetName">The name of the sheet to query. If null, the first sheet is used.</param>
    /// <param name="startCell">The starting cell reference (e.g., "C2"). Default is "A1".</param>
    /// <param name="configuration">Optional configuration settings.</param>
    /// <param name="cancellationToken">A token to cancel the asynchronous operation.</param>
    /// <remarks>
    /// When <paramref name="hasHeaderRow"/> is true, column names from the first row become dynamic property names, otherwise they will be assigned alphabetically (A, B, C, etc.).
    /// </remarks>
    [CreateSyncVersion]
    public async IAsyncEnumerable<dynamic> QueryAsync(string path, bool hasHeaderRow = false,
        string? sheetName = null, string startCell = "A1", OpenXmlConfiguration? configuration = null,
        [EnumeratorCancellation] CancellationToken cancellationToken = default)
    {
        var stream = FileHelper.OpenSharedRead(path);
        await using var disposableStream = stream.ConfigureAwait(false); 

        await foreach (var item in QueryAsync(stream, hasHeaderRow, sheetName, startCell, configuration, false, cancellationToken).ConfigureAwait(false))
            yield return item;
    }

    /// <summary>
    /// Queries an Excel document and returns dynamic objects representing each row.
    /// </summary>
    /// <param name="stream">The stream containing the Excel file data. The stream position is not reset after reading.</param>
    /// <param name="hasHeaderRow">If true, the first row is used as column headers for the dynamic object properties. Default is false.</param>
    /// <param name="sheetName">The name of the sheet to query. If null, the first sheet is used.</param>
    /// <param name="startCell">The starting cell reference (e.g., "C2"). Default is "A1".</param>
    /// <param name="configuration">Optional configuration settings.</param>
    /// <param name="leaveOpen">True to leave the stream open after the query is completed, otherwise false.</param>
    /// <param name="cancellationToken">A token to cancel the asynchronous operation.</param>
    /// <remarks>
    /// When <paramref name="hasHeaderRow"/> is true, column names from the first row become dynamic property names, otherwise they will be assigned alphabetically (A, B, C, etc.).
    /// </remarks>
    [CreateSyncVersion]
    public async IAsyncEnumerable<dynamic> QueryAsync(Stream stream, bool hasHeaderRow = false,
        string? sheetName = null, string startCell = "A1", OpenXmlConfiguration? configuration = null,
        bool leaveOpen = false, [EnumeratorCancellation] CancellationToken cancellationToken = default)
    {
        using var reader = await OpenXmlReader.CreateAsync(stream, configuration, leaveOpen, cancellationToken).ConfigureAwait(false);
        await foreach (var item in reader.QueryAsync(hasHeaderRow, sheetName, startCell, cancellationToken).ConfigureAwait(false))
            yield return item;
    }

    #endregion

    #region Query Range

    /// <summary>
    /// Queries a specific rectangular region within an worksheet using index-based coordinates.
    /// </summary>
    /// <param name="path">The path to the Excel document.</param>
    /// <param name="hasHeaderRow">If true, the first row within the range is used as column headers for dynamic object properties. Default is false.</param>
    /// <param name="sheetName">The name of the sheet to query. If null, the first sheet is used.</param>
    /// <param name="startCell">The starting cell reference. Default is "A1".</param>
    /// <param name="endCell">The ending cell reference. If left empty, the last cell containing data will be used.</param>
    /// <param name="configuration">Optional configuration settings.</param>
    /// <param name="cancellationToken">A token to cancel the asynchronous operation.</param>
    [CreateSyncVersion]
    public async IAsyncEnumerable<dynamic> QueryRangeAsync(string path, bool hasHeaderRow = false,
        string? sheetName = null, string startCell = "A1", string? endCell = null, OpenXmlConfiguration? configuration = null,
        [EnumeratorCancellation] CancellationToken cancellationToken = default)
    {
        var stream = FileHelper.OpenSharedRead(path);
        await using var disposableStream = stream.ConfigureAwait(false); 

        await foreach (var item in QueryRangeAsync(stream, hasHeaderRow, sheetName, startCell, endCell, configuration, false, cancellationToken).ConfigureAwait(false))
            yield return item;
    }

    /// <summary>
    /// Queries a specific rectangular region within an worksheet using index-based coordinates.
    /// </summary>
    /// <param name="stream">The stream containing the Excel file data. The stream position is not reset after reading.</param>
    /// <param name="hasHeaderRow">If true, the first row within the range is used as column headers for dynamic object properties. Default is false.</param>
    /// <param name="sheetName">The name of the sheet to query. If null, the first sheet is used.</param>
    /// <param name="startCell">The starting cell reference. Default is "A1".</param>
    /// <param name="endCell">The ending cell reference. If left empty, the last cell containing data will be used.</param>
    /// <param name="configuration">Optional configuration settings.</param>
    /// <param name="leaveOpen">True to leave the stream open after the query is completed, otherwise false.</param>
    /// <param name="cancellationToken">A token to cancel the asynchronous operation.</param>
    [CreateSyncVersion]
    public async IAsyncEnumerable<dynamic> QueryRangeAsync(Stream stream, bool hasHeaderRow = false,
        string? sheetName = null, string startCell = "A1", string? endCell = null, OpenXmlConfiguration? configuration = null,
        bool leaveOpen = false, [EnumeratorCancellation] CancellationToken cancellationToken = default)
    {
        using var reader = await OpenXmlReader.CreateAsync(stream, configuration, leaveOpen, cancellationToken).ConfigureAwait(false);
        await foreach (var item in reader.QueryRangeAsync(hasHeaderRow, sheetName, startCell, endCell, cancellationToken).ConfigureAwait(false))
            yield return item;
    }

    /// <summary>
    /// Queries a specific rectangular region within an worksheet using index-based coordinates.
    /// </summary>
    /// <param name="path">The path to Excel document.</param>
    /// <param name="hasHeaderRow">If true, the first row within the range is used as column headers for dynamic object properties. Default is false.</param>
    /// <param name="sheetName">The name of the sheet to query. If null, the first sheet is used.</param>
    /// <param name="startRowIndex">The 1-based index of the starting row.</param>
    /// <param name="startColumnIndex">The 1-based index of the starting column.</param>
    /// <param name="endRowIndex">The 1-based index of the ending row (inclusive). If null, reads to the last row containing data.</param>
    /// <param name="endColumnIndex">The 1-based index of the ending column (inclusive). If null, reads to the last column containing data.</param>
    /// <param name="configuration">Optional configuration settings.</param>
    /// <param name="cancellationToken">A token to cancel the asynchronous operation.</param>
    [CreateSyncVersion]
    public async IAsyncEnumerable<dynamic> QueryRangeAsync(string path, bool hasHeaderRow = false,
        string? sheetName = null, int startRowIndex = 1, int startColumnIndex = 1, int? endRowIndex = null,
        int? endColumnIndex = null, OpenXmlConfiguration? configuration = null,
        [EnumeratorCancellation] CancellationToken cancellationToken = default)
    {
        var stream = FileHelper.OpenSharedRead(path);
        await using var disposableStream = stream.ConfigureAwait(false); 

        await foreach (var item in QueryRangeAsync(stream, hasHeaderRow, sheetName, startRowIndex, startColumnIndex, endRowIndex, endColumnIndex, configuration, false, cancellationToken).ConfigureAwait(false))
            yield return item;
    }

    /// <summary>
    /// Queries a specific rectangular region within an worksheet using index-based coordinates.
    /// </summary>
    /// <param name="stream">The stream containing the Excel file data. The stream position is not reset after reading.</param>
    /// <param name="hasHeaderRow">If true, the first row within the range is used as column headers for dynamic object properties. Default is false.</param>
    /// <param name="sheetName">The name of the sheet to query. If null, the first sheet is used.</param>
    /// <param name="startRowIndex">The 1-based index of the starting row.</param>
    /// <param name="startColumnIndex">The 1-based index of the starting column.</param>
    /// <param name="endRowIndex">The 1-based index of the ending row (inclusive). If null, reads to the last row containing data.</param>
    /// <param name="endColumnIndex">The 1-based index of the ending column (inclusive). If null, reads to the last column containing data.</param>
    /// <param name="configuration">Optional configuration settings.</param>
    /// <param name="leaveOpen">True to leave the stream open after the query is completed, otherwise false.</param>
    /// <param name="cancellationToken">A token to cancel the asynchronous operation.</param>
    [CreateSyncVersion]
    public async IAsyncEnumerable<dynamic> QueryRangeAsync(Stream stream, bool hasHeaderRow = false,
        string? sheetName = null, int startRowIndex = 1, int startColumnIndex = 1, int? endRowIndex = null,
        int? endColumnIndex = null, OpenXmlConfiguration? configuration = null, bool leaveOpen = false, 
        [EnumeratorCancellation] CancellationToken cancellationToken = default)
    {
        using var reader = await OpenXmlReader.CreateAsync(stream, configuration, leaveOpen, cancellationToken).ConfigureAwait(false);
        await foreach (var item in reader.QueryRangeAsync(hasHeaderRow, sheetName, startRowIndex, startColumnIndex, endRowIndex, endColumnIndex, cancellationToken).ConfigureAwait(false))
            yield return item;
    }

    #endregion

    #region Query As DataTable

    /// <summary>
    /// Queries an Excel sheet and returns the results as a <see cref="DataTable"/>.
    /// </summary>
    /// <param name="path">The path to the Excel file data.</param>
    /// <param name="hasHeaderRow">If true, the first row is used as column headers.</param>
    /// <param name="sheetName">The name of the sheet to query. If not specified, the first sheet is used.</param>
    /// <param name="startCell">The starting cell reference (e.g., "C2"). Default is "A1".</param>
    /// <param name="configuration">Optional configuration settings.</param>
    /// <param name="cancellationToken">A token to cancel the asynchronous operation.</param>
    /// <remarks>
    /// Empty column names are skipped.
    /// This method loads the entire worksheet into memory, so its usage is recommended only for datasets of moderate size.
    /// </remarks>
    [CreateSyncVersion]
    public async Task<DataTable> QueryAsDataTableAsync(string path, bool hasHeaderRow = true,
        string? sheetName = null, string startCell = "A1", OpenXmlConfiguration? configuration = null,
        CancellationToken cancellationToken = default)
    {
        var stream = FileHelper.OpenSharedRead(path);
        await using var disposableStream = stream.ConfigureAwait(false); 

        return await QueryAsDataTableAsync(stream, hasHeaderRow, sheetName, startCell, configuration, false, cancellationToken).ConfigureAwait(false);
    }

    /// <summary>
    /// Queries an Excel sheet and returns the results as a <see cref="DataTable"/>.
    /// </summary>
    /// <param name="stream">The stream containing the Excel file data. The stream position is not reset after reading.</param>
    /// <param name="hasHeaderRow">If true, the first row is used as column headers.</param>
    /// <param name="sheetName">The name of the sheet to query. If not specified, the first sheet is used.</param>
    /// <param name="startCell">The starting cell reference (e.g., "C2"). Default is "A1".</param>
    /// <param name="configuration">Optional configuration settings.</param>
    /// <param name="leaveOpen">True to leave the stream open after the query is completed, otherwise false.</param>
    /// <param name="cancellationToken">A token to cancel the asynchronous operation.</param>
    /// <remarks>
    /// Empty column names are skipped.
    /// This method loads the entire worksheet into memory, so its usage is recommended only for datasets of moderate size.
    /// </remarks>
    [CreateSyncVersion]
    public async Task<DataTable> QueryAsDataTableAsync(Stream stream, bool hasHeaderRow = true,
        string? sheetName = null, string startCell = "A1", OpenXmlConfiguration? configuration = null,
        bool leaveOpen = false, CancellationToken cancellationToken = default)
    {
        sheetName ??= (await GetSheetNamesAsync(stream, false, cancellationToken).ConfigureAwait(false)).First();

        var dt = new DataTable(sheetName);
        var first = true;
        using var reader = await OpenXmlReader.CreateAsync(stream, configuration, leaveOpen, cancellationToken).ConfigureAwait(false);
        var rows = reader.QueryAsync(false, sheetName, startCell, cancellationToken);

        var columnDict = new Dictionary<string, string>();
        await foreach (var row in rows.ConfigureAwait(false))
        {
            if (first)
            {
                foreach (var entry in row)
                {
                    cancellationToken.ThrowIfCancellationRequested();

                    var columnName = hasHeaderRow ? entry.Value?.ToString() : entry.Key;
                    if (!string.IsNullOrWhiteSpace(columnName)) // avoid #298 : Column '' does not belong to table
                    {
                        var column = new DataColumn(columnName, typeof(object)) { Caption = columnName };
                        dt.Columns.Add(column);
                        columnDict.Add(entry.Key, columnName); //same column name throw exception???
                    }
                }

                dt.BeginLoadData();
                first = false;
                if (hasHeaderRow)
                {
                    continue;
                }
            }

            var newRow = dt.NewRow();
            foreach (var entry in columnDict)
            {
                newRow[entry.Value] = row[entry.Key]; //TODO: optimize not using string key
            }

            dt.Rows.Add(newRow);
        }

        dt.EndLoadData();
        return dt;
    }

    #endregion

    #region Sheet Info

    /// <summary>
    /// Retrieves the names of all sheets in an Excel workbook.
    /// </summary>
    /// <param name="path">The path to the Excel file.</param>
    /// <param name="cancellationToken">A token to cancel the asynchronous operation.</param>
    /// <returns>A list of sheet names in the workbook, or an empty list if no sheets are found.</returns>
    /// <remarks>
    /// Sheet names are returned in the order they appear in the workbook.
    /// </remarks>
    [CreateSyncVersion]
    public async Task<List<string>> GetSheetNamesAsync(string path, CancellationToken cancellationToken = default)
    {
        var stream = FileHelper.OpenSharedRead(path);
        await using var disposableStream = stream.ConfigureAwait(false); 

        return await GetSheetNamesAsync(stream, false, cancellationToken).ConfigureAwait(false);
    }

    /// <summary>
    /// Retrieves the names of all sheets in an Excel workbook.
    /// </summary>
    /// <param name="stream">The stream containing the Excel file data. The stream position is not reset after reading.</param>
    /// <param name="leaveOpen">True to leave the stream open after the operation is completed, otherwise false.</param>
    /// <param name="cancellationToken">A token to cancel the asynchronous operation.</param>
    /// <returns>A list of sheet names in the workbook, or an empty list if no sheets are found.</returns>
    /// <remarks>
    /// Sheet names are returned in the order they appear in the workbook.
    /// </remarks>
    [CreateSyncVersion]
    public async Task<List<string>> GetSheetNamesAsync(Stream stream, bool leaveOpen = false, CancellationToken cancellationToken = default)
    {
        var archive = await OpenXmlZip.CreateAsync(stream, leaveOpen: true, cancellationToken: cancellationToken).ConfigureAwait(false);
        await using var disposableArchive = archive.ConfigureAwait(false);
        using var reader = await OpenXmlReader.CreateAsync(stream, null, leaveOpen, cancellationToken).ConfigureAwait(false);

        var rels = await reader.GetWorkbookRelsAsync(archive.EntryCollection, cancellationToken).ConfigureAwait(false);
        return rels?.Select(s => s.Name).ToList() ?? [];
    }

    /// <summary>
    /// Retrieves detailed information about all sheets in an Excel workbook.
    /// </summary>
    /// <param name="path">The path to the Excel file.</param>
    /// <param name="cancellationToken">A token to cancel the asynchronous operation.</param>
    /// <returns>A list of <see cref="SheetInfo"/> objects containing metadata for each sheet, including name, dimensions, and sheet index.</returns>
    /// <remarks>
    /// Sheet information is returned in the order sheets appear in the workbook.
    /// </remarks>
    [CreateSyncVersion]
    public async Task<List<SheetInfo>> GetSheetInformationsAsync(string path, CancellationToken cancellationToken = default)
    {
        var stream = FileHelper.OpenSharedRead(path);
        await using var disposableStream = stream.ConfigureAwait(false); 

        return await GetSheetInformationsAsync(stream, false, cancellationToken).ConfigureAwait(false);
    }

    /// <summary>
    /// Retrieves detailed information about all sheets in an Excel workbook.
    /// </summary>
    /// <param name="stream">The stream containing the Excel file data. The stream position is not reset after reading.</param>
    /// <param name="leaveOpen">True to leave the stream open after the operation is completed, otherwise false.</param>
    /// <param name="cancellationToken">A token to cancel the asynchronous operation.</param>
    /// <returns>A list of <see cref="SheetInfo"/> objects containing metadata for each sheet, including name, dimensions, and sheet index.</returns>
    /// <remarks>
    /// Sheet information is returned in the order sheets appear in the workbook.
    /// </remarks>
    [CreateSyncVersion]
    public async Task<List<SheetInfo>> GetSheetInformationsAsync(Stream stream, bool leaveOpen = false, CancellationToken cancellationToken = default)
    {
        var archive = await OpenXmlZip.CreateAsync(stream, cancellationToken: cancellationToken).ConfigureAwait(false);

        await using var disposableArchve = archive.ConfigureAwait(false);
        using var reader = await OpenXmlReader.CreateAsync(stream, null, leaveOpen, cancellationToken).ConfigureAwait(false);

        var rels = await reader.GetWorkbookRelsAsync(archive.EntryCollection, cancellationToken).ConfigureAwait(false);
        return rels?.Select((s, i) => s.ToSheetInfo((uint)i)).ToList() ?? [];
    }

    /// <summary>
    /// Retrieves the dimensions (used cell range) for all sheets in an Excel workbook.
    /// </summary>
    /// <param name="path">The path to the Excel file.</param>
    /// <param name="cancellationToken">A token to cancel the asynchronous operation.</param>
    /// <returns>A list of <see cref="ExcelRange"/> objects representing the used dimensions for each sheet in the workbook.</returns>
    /// <remarks>
    /// The dimension of a sheet represents the rectangular range of cells that contain data.
    /// Each <see cref="ExcelRange"/> in the returned list corresponds to a sheet, in the order sheets appear in the workbook.
    /// Empty sheets will have dimensions that reflect no used cells.
    /// A synchronous version of this method is automatically generated via the [CreateSyncVersion] attribute.
    /// </remarks>
    [CreateSyncVersion]
    public async Task<IList<ExcelRange>> GetSheetDimensionsAsync(string path, CancellationToken cancellationToken = default)
    {
        var stream = FileHelper.OpenSharedRead(path);
        await using var disposableStream = stream.ConfigureAwait(false); 

        return await GetSheetDimensionsAsync(stream, false, cancellationToken).ConfigureAwait(false);
    }

    /// <summary>
    /// Retrieves the dimensions (used cell range) for all sheets in an Excel workbook.
    /// </summary>
    /// <param name="stream">The stream containing the Excel file data. The stream position is not reset after reading.</param>
    /// <param name="leaveOpen">True to leave the stream open after the operation is completed, otherwise false.</param>
    /// <param name="cancellationToken">A token to cancel the asynchronous operation.</param>
    /// <returns>A list of <see cref="ExcelRange"/> objects representing the used dimensions for each sheet in the workbook.</returns>
    /// <remarks>
    /// The dimension of a sheet represents the rectangular range of cells that contain data.
    /// Each <see cref="ExcelRange"/> in the returned list corresponds to a sheet, in the order sheets appear in the workbook.
    /// Empty sheets will have dimensions that reflect no used cells.
    /// A synchronous version of this method is automatically generated via the [CreateSyncVersion] attribute.
    /// </remarks>
    [CreateSyncVersion]
    public async Task<IList<ExcelRange>> GetSheetDimensionsAsync(Stream stream, bool leaveOpen = false, CancellationToken cancellationToken = default)
    {
        using var reader = await OpenXmlReader.CreateAsync(stream, null, leaveOpen, cancellationToken).ConfigureAwait(false);
        return await reader.GetDimensionsAsync(cancellationToken).ConfigureAwait(false);
    }

    /// <summary>
    /// Retrieves the column names from the first row (header row) of an Excel sheet.
    /// </summary>
    /// <param name="path">The path to the Excel document.</param>
    /// <param name="hasHeaderRow">If true, the first row values are used as column names. If false, column letters (A, B, C, etc.) are used. Default is false.</param>
    /// <param name="sheetName">The name of the worksheet to query. If not provided, the first sheet is used.</param>
    /// <param name="startCell">The starting cell reference (e.g., "C2"). Default is "A1".</param>
    /// <param name="cancellationToken">A token to cancel the asynchronous operation.</param>
    /// <returns>A collection of column names from the specified location, or an empty collection if the sheet is empty.</returns>
    /// <remarks>
    /// Returns an empty collection if the sheet has no rows starting from <paramref name="startCell"/>.
    /// </remarks>
    [CreateSyncVersion]
    public async Task<ICollection<string>> GetColumnNamesAsync(string path, bool hasHeaderRow = false,
        string? sheetName = null, string startCell = "A1", CancellationToken cancellationToken = default)
    {
        var stream = FileHelper.OpenSharedRead(path);
        await using var disposableStream = stream.ConfigureAwait(false); 

        return await GetColumnNamesAsync(stream, hasHeaderRow, sheetName, startCell, false, cancellationToken).ConfigureAwait(false);
    }


    /// <summary>
    /// Retrieves the column names from the first row (header row) of an Excel sheet.
    /// </summary>
    /// <param name="stream">The stream containing the Excel file data.</param>
    /// <param name="hasHeaderRow">If true, the first row values are used as column names. If false, column letters (A, B, C, etc.) are used. Default is false.</param>
    /// <param name="sheetName">The name of the worksheet to query. If not provided, the first sheet is used.</param>
    /// <param name="startCell">The starting cell reference (e.g., "C2"). Default is "A1".</param>
    /// <param name="leaveOpen">True to leave the stream open after the query is completed, otherwise false.</param>
    /// <param name="cancellationToken">A token to cancel the asynchronous operation.</param>
    /// <returns>A collection of column names from the specified location, or an empty collection if the sheet is empty.</returns>
    /// <remarks>
    /// Returns an empty collection if the sheet has no rows starting from <paramref name="startCell"/>.
    /// </remarks>
    [CreateSyncVersion]
    public async Task<ICollection<string>> GetColumnNamesAsync(Stream stream, bool hasHeaderRow = false,
        string? sheetName = null, string startCell = "A1", bool leaveOpen = false, CancellationToken cancellationToken = default)
    {
        var enumerator = QueryAsync(stream, hasHeaderRow, sheetName, startCell, null, leaveOpen, cancellationToken).GetAsyncEnumerator(cancellationToken);
        await using var disposableEnumerator = enumerator.ConfigureAwait(false);

        if (await enumerator.MoveNextAsync().ConfigureAwait(false))
            return (enumerator.Current as IDictionary<string, object?>)?.Keys ?? [];

        return [];
    }

    /// <summary>
    /// Retrieves all threaded comments and notes from a specific sheet in an Excel workbook.
    /// </summary>
    /// <param name="path">The path to the Excel document.</param>
    /// <param name="sheetName">The name of the worksheet from which to retrieve comments. If not provided, comments from the first sheet are returned.</param>
    /// <param name="cancellationToken">A token to cancel the asynchronous operation.</param>
    /// <remarks>
    /// Comments are cell-level annotations in Excel files that are stored separately from the cell data.
    /// The returned <see cref="CommentResultSet"/> provides access to both threaded comments and legacy note comments, along with the associated metadata.
    /// </remarks>
    [CreateSyncVersion]
    public async Task<CommentResultSet> RetrieveCommentsAsync(string path, string? sheetName, CancellationToken cancellationToken = default)
    {
        var stream = FileHelper.OpenSharedRead(path);
        await using var disposableStream = stream.ConfigureAwait(false); 

        return await RetrieveCommentsAsync(stream, sheetName, false, cancellationToken).ConfigureAwait(false);
    }

    /// <summary>
    /// Retrieves all threaded comments and notes from a specific sheet in an Excel workbook.
    /// </summary>
    /// <param name="stream">The stream containing the Excel file data. The stream position is not reset after reading.</param>
    /// <param name="sheetName">The name of the worksheet from which to retrieve comments. If not provided, comments from the first sheet are retrieved.</param>
    /// <param name="leaveOpen">True to leave the stream open after the operation is completed, otherwise false.</param>
    /// <param name="cancellationToken">A token to cancel the asynchronous operation.</param>
    /// <remarks>
    /// Comments are cell-level annotations in Excel files that are stored separately from the cell data.
    /// The returned <see cref="CommentResultSet"/> provides access to both threaded comments and legacy note comments, along with the associated metadata.
    /// </remarks>
    [CreateSyncVersion]
    public async Task<CommentResultSet> RetrieveCommentsAsync(Stream stream, string? sheetName, bool leaveOpen = false, CancellationToken cancellationToken = default)
    {
        using var reader = await OpenXmlReader.CreateAsync(stream, null, leaveOpen, cancellationToken).ConfigureAwait(false);
        return await reader.ReadCommentsAsync(sheetName, cancellationToken).ConfigureAwait(false);
    }

    #endregion

    #region DataReader

    /// <summary>
    /// Gets an <see cref="IDataReader" /> for the Excel document provided for synchronous reading.
    /// </summary>
    /// <param name="path">The path to the Excel document.</param>
    /// <param name="hasHeaderRow">If true, the first row is used as column headers. Default is false.</param>
    /// <param name="sheetName">The name of the worksheet to read. If not provided, the first sheet is used.</param>
    /// <param name="startCell">The starting cell reference (e.g."C2"). Default is "A1".</param>
    /// <param name="configuration">Optional configuration settings.</param>
    /// <remarks>
    /// The returned <see cref="MiniExcelDataReader"/> implements <see cref="IDataReader"/> and supports its standard reading patterns.
    /// The data reader returned by this method is designed to perform synchronous, blocking reads, and will throw <exception cref="InvalidOperationException" /> if an asynchronous operation is called from it.
    /// For asynchronous reading scenarios, use <see cref="GetAsyncDataReader(string, bool, string?, string, OpenXmlConfiguration?, CancellationToken)"/> instead.
    /// </remarks>
    public MiniExcelDataReader GetDataReader(string path, bool hasHeaderRow = false,
        string? sheetName = null, string startCell = "A1", OpenXmlConfiguration? configuration = null)
    {
        var stream = FileHelper.OpenSharedRead(path);
        var values = Query(stream, hasHeaderRow, sheetName, startCell, configuration, leaveOpen: false).Cast<IDictionary<string, object?>>();

        return MiniExcelDataReader.Create(stream, values, leaveOpen: false);
    }

    /// <summary>
    /// Gets an <see cref="IDataReader" /> for the Excel document from an underlying stream for synchronous reading.
    /// </summary>
    /// <param name="stream">The stream containing the Excel file data.</param>
    /// <param name="hasHeaderRow">If true, the first row is used as column headers. Default is false.</param>
    /// <param name="sheetName">The name of the worksheet to read. If not provided, the first sheet is used.</param>
    /// <param name="startCell">The starting cell reference (e.g."C2"). Default is "A1".</param>
    /// <param name="configuration">Optional configuration settings.</param>
    /// <param name="leaveOpen">True to leave the stream open after the data reader is disposed, otherwise false.</param>
    /// <remarks>
    /// The returned <see cref="MiniExcelDataReader"/> implements <see cref="IDataReader"/> and supports its standard reading patterns.
    /// The data reader returned by this method is designed to perform synchronous, blocking reads, and will throw <exception cref="InvalidOperationException" /> if an asynchronous operation is called from it.
    /// For asynchronous reading scenarios, use <see cref="GetAsyncDataReader(Stream, bool, string?, string, OpenXmlConfiguration?, bool, CancellationToken)"/> instead.
    /// </remarks>
    public MiniExcelDataReader GetDataReader(Stream stream, bool hasHeaderRow = false,
        string? sheetName = null, string startCell = "A1", OpenXmlConfiguration? configuration = null, bool leaveOpen = false)
    {
        var values = Query(stream, hasHeaderRow, sheetName, startCell, configuration, leaveOpen).Cast<IDictionary<string, object?>>();
        return MiniExcelDataReader.Create(stream, values, leaveOpen);
    }

    /// <summary>
    /// Gets an <see cref="IDataReader" /> for the Excel document from an underlying stream for asynchronous reading.
    /// </summary>
    /// <param name="path">The path to the Excel document.</param>
    /// <param name="hasHeaderRow">If true, the first row is used as column headers. Default is false.</param>
    /// <param name="sheetName">The name of the worksheet to read. If null, the first sheet is used.</param>
    /// <param name="startCell">The starting cell reference (e.g."C2"). Default is "A1".</param>
    /// <param name="configuration">Optional configuration settings.</param>
    /// <param name="cancellationToken">A token to cancel the asynchronous operation.</param>
    /// <remarks>
    /// The returned <see cref="MiniExcelDataReader"/> implements <see cref="IDataReader"/> and supports its standard reading patterns.
    /// The data reader returned by this method is designed to supports asynchronous reads, but will not throw an exception if a synchronous operation is performed.
    /// Still, it's advised to use <see cref="GetDataReader(Stream, bool, string?, string, OpenXmlConfiguration?, bool)"/> for synchronous reads instead.
    /// </remarks>
    public async Task<MiniExcelDataReader> GetAsyncDataReader(string path, bool hasHeaderRow = false,
        string? sheetName = null, string startCell = "A1", OpenXmlConfiguration? configuration = null, CancellationToken cancellationToken = default)
    {
        var stream = FileHelper.OpenSharedRead(path);
        var values = QueryAsync(stream, hasHeaderRow, sheetName, startCell, configuration, leaveOpen: false, cancellationToken);
        
        return await MiniExcelDataReader.CreateAsync(stream, values.CastToDictionary(cancellationToken), leaveOpen: false).ConfigureAwait(false);
    }

    /// <summary>
    /// Gets an <see cref="IDataReader" /> for the Excel document from an underlying stream for asynchronous reading.
    /// </summary>
    /// <param name="stream">The stream containing the Excel file data.</param>
    /// <param name="hasHeaderRow">If true, the first row is used as column headers. Default is false.</param>
    /// <param name="sheetName">The name of the worksheet to read. If null, the first sheet is used.</param>
    /// <param name="startCell">The starting cell reference (e.g."C2"). Default is "A1".</param>
    /// <param name="configuration">Optional configuration settings.</param>
    /// <param name="leaveOpen">True to leave the stream open after the data reader is disposed, otherwise false.</param>
    /// <param name="cancellationToken">A token to cancel the asynchronous operation.</param>
    /// <remarks>
    /// The returned <see cref="MiniExcelDataReader"/> implements <see cref="IDataReader"/> and supports its standard reading patterns.
    /// The data reader returned by this method is designed to supports asynchronous reads, but will not throw an exception if a synchronous operation is performed.
    /// Still, it's advised to use <see cref="GetDataReader(Stream, bool, string?, string, OpenXmlConfiguration?, bool)"/> for synchronous reads instead.
    /// </remarks>
    public async Task<MiniExcelDataReader> GetAsyncDataReader(Stream stream, bool hasHeaderRow = false,
        string? sheetName = null, string startCell = "A1", OpenXmlConfiguration? configuration = null, bool leaveOpen = false,
        CancellationToken cancellationToken = default)
    {
        var values = QueryAsync(stream, hasHeaderRow, sheetName, startCell, configuration, leaveOpen, cancellationToken);
        return await MiniExcelDataReader.CreateAsync(stream, values.CastToDictionary(cancellationToken), leaveOpen).ConfigureAwait(false);
    }

    #endregion
}
