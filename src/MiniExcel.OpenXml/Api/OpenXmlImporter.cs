

// ReSharper disable once CheckNamespace
namespace MiniExcelLib.OpenXml;

public sealed partial class OpenXmlImporter
{
    internal OpenXmlImporter() { }
    
    #region Query

    [CreateSyncVersion]
    public async IAsyncEnumerable<T> QueryAsync<T>(string path, string? sheetName = null,
        string startCell = "A1", bool treatHeaderAsData = false, OpenXmlConfiguration? configuration = null,
        [EnumeratorCancellation] CancellationToken cancellationToken = default) where T : class, new()
    {
        var stream = FileHelper.OpenSharedRead(path);
        await using var disposableStream = stream.ConfigureAwait(false); 

        var query = QueryAsync<T>(stream, sheetName, startCell, treatHeaderAsData, configuration, cancellationToken);
        
        await foreach (var item in query.ConfigureAwait(false))
            yield return item; 
    }

    [CreateSyncVersion]
    public async IAsyncEnumerable<T> QueryAsync<T>(Stream stream, string? sheetName = null,
        string startCell = "A1", bool treatHeaderAsData = false, OpenXmlConfiguration? configuration = null,
        [EnumeratorCancellation] CancellationToken cancellationToken = default) where T : class, new()
    {
        using var reader = await OpenXmlReader.CreateAsync(stream, configuration, cancellationToken).ConfigureAwait(false);
        await foreach (var item in reader.QueryAsync<T>(sheetName, startCell, treatHeaderAsData, cancellationToken).ConfigureAwait(false))
            yield return item;
    }

    [CreateSyncVersion]
    public async IAsyncEnumerable<dynamic> QueryAsync(string path, bool hasHeaderRow = false,
        string? sheetName = null, string startCell = "A1", OpenXmlConfiguration? configuration = null,
        [EnumeratorCancellation] CancellationToken cancellationToken = default)
    {
        var stream = FileHelper.OpenSharedRead(path);
        await using var disposableStream = stream.ConfigureAwait(false); 

        await foreach (var item in QueryAsync(stream, hasHeaderRow, sheetName, startCell, configuration, cancellationToken).ConfigureAwait(false))
            yield return item;
    }

    [CreateSyncVersion]
    public async IAsyncEnumerable<dynamic> QueryAsync(Stream stream, bool hasHeaderRow = false,
        string? sheetName = null, string startCell = "A1", OpenXmlConfiguration? configuration = null,
        [EnumeratorCancellation] CancellationToken cancellationToken = default)
    {
        using var excelReader = await OpenXmlReader.CreateAsync(stream, configuration, cancellationToken).ConfigureAwait(false);
        await foreach (var item in excelReader.QueryAsync(hasHeaderRow, sheetName, startCell, cancellationToken).ConfigureAwait(false))
            yield return item;
    }

    #endregion

    #region Query Range

    /// <summary>
    /// Extract the given range。 Only uppercase letters are effective。
    /// e.g.
    ///     MiniExcel.QueryRange(path, startCell: "A2", endCell: "C3")
    ///     A2 represents the second row of column A, C3 represents the third row of column C
    ///     If you don't want to restrict rows, just don't include numbers
    /// </summary>
    /// <returns></returns>
    [CreateSyncVersion]
    public async IAsyncEnumerable<dynamic> QueryRangeAsync(string path, bool hasHeaderRow = false,
        string? sheetName = null, string startCell = "A1", string? endCell = null, OpenXmlConfiguration? configuration = null,
        [EnumeratorCancellation] CancellationToken cancellationToken = default)
    {
        var stream = FileHelper.OpenSharedRead(path);
        await using var disposableStream = stream.ConfigureAwait(false); 

        await foreach (var item in QueryRangeAsync(stream, hasHeaderRow, sheetName, startCell, endCell, configuration, cancellationToken).ConfigureAwait(false))
            yield return item;
    }

    [CreateSyncVersion]
    public async IAsyncEnumerable<dynamic> QueryRangeAsync(Stream stream, bool hasHeaderRow = false,
        string? sheetName = null, string startCell = "A1", string? endCell = null, OpenXmlConfiguration? configuration = null,
        [EnumeratorCancellation] CancellationToken cancellationToken = default)
    {
        using var excelReader = await OpenXmlReader.CreateAsync(stream, configuration, cancellationToken).ConfigureAwait(false);
        await foreach (var item in excelReader.QueryRangeAsync(hasHeaderRow, sheetName, startCell, endCell, cancellationToken).ConfigureAwait(false))
            yield return item;
    }

    [CreateSyncVersion]
    public async IAsyncEnumerable<dynamic> QueryRangeAsync(string path, bool hasHeaderRow = false,
        string? sheetName = null, int startRowIndex = 1, int startColumnIndex = 1, int? endRowIndex = null,
        int? endColumnIndex = null, OpenXmlConfiguration? configuration = null,
        [EnumeratorCancellation] CancellationToken cancellationToken = default)
    {
        var stream = FileHelper.OpenSharedRead(path);
        await using var disposableStream = stream.ConfigureAwait(false); 

        await foreach (var item in QueryRangeAsync(stream, hasHeaderRow, sheetName, startRowIndex, startColumnIndex, endRowIndex, endColumnIndex, configuration, cancellationToken).ConfigureAwait(false))
            yield return item;
    }

    [CreateSyncVersion]
    public async IAsyncEnumerable<dynamic> QueryRangeAsync(Stream stream, bool hasHeaderRow = false,
        string? sheetName = null, int startRowIndex = 1, int startColumnIndex = 1, int? endRowIndex = null,
        int? endColumnIndex = null, OpenXmlConfiguration? configuration = null,
        [EnumeratorCancellation] CancellationToken cancellationToken = default)
    {
        using var excelReader = await OpenXmlReader.CreateAsync(stream, configuration, cancellationToken).ConfigureAwait(false);
        await foreach (var item in excelReader.QueryRangeAsync(hasHeaderRow, sheetName, startRowIndex, startColumnIndex, endRowIndex, endColumnIndex, cancellationToken).ConfigureAwait(false))
            yield return item;
    }

    #endregion

    #region Query As DataTable

    /// <summary>
    /// QueryAsDataTable is not recommended, because it'll load all data into memory.
    /// </summary>
    [CreateSyncVersion]
    public async Task<DataTable> QueryAsDataTableAsync(string path, bool hasHeaderRow = true,
        string? sheetName = null, string startCell = "A1", OpenXmlConfiguration? configuration = null,
        CancellationToken cancellationToken = default)
    {
        var stream = FileHelper.OpenSharedRead(path);
        await using var disposableStream = stream.ConfigureAwait(false); 

        return await QueryAsDataTableAsync(stream, hasHeaderRow, sheetName, startCell, configuration, cancellationToken).ConfigureAwait(false);
    }

    /// <summary>
    /// QueryAsDataTable is not recommended, because it'll load all data into memory.
    /// </summary>
    [CreateSyncVersion]
    public async Task<DataTable> QueryAsDataTableAsync(Stream stream, bool hasHeaderRow = true,
        string? sheetName = null, string startCell = "A1", OpenXmlConfiguration? configuration = null,
        CancellationToken cancellationToken = default)
    {
        sheetName ??= (await GetSheetNamesAsync(stream, cancellationToken).ConfigureAwait(false)).First();

        var dt = new DataTable(sheetName);
        var first = true;
        using var reader = await OpenXmlReader.CreateAsync(stream, configuration, cancellationToken).ConfigureAwait(false);
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

    [CreateSyncVersion]
    public async Task<List<string>> GetSheetNamesAsync(string path, CancellationToken cancellationToken = default)
    {
        var stream = FileHelper.OpenSharedRead(path);
        await using var disposableStream = stream.ConfigureAwait(false); 

        return await GetSheetNamesAsync(stream, cancellationToken).ConfigureAwait(false);
    }

    [CreateSyncVersion]
    public async Task<List<string>> GetSheetNamesAsync(Stream stream, CancellationToken cancellationToken = default)
    {
        var archive = await OpenXmlZip.CreateAsync(stream, leaveOpen: true, cancellationToken: cancellationToken).ConfigureAwait(false);
        await using var disposableArchive = archive.ConfigureAwait(false);
        using var reader = await OpenXmlReader.CreateAsync(stream, null, cancellationToken).ConfigureAwait(false);

        var rels = await reader.GetWorkbookRelsAsync(archive.EntryCollection, cancellationToken).ConfigureAwait(false);
        return rels?.Select(s => s.Name).ToList() ?? [];
    }

    /// Retrieves detailed information about all sheets in an Excel workbook.
    /// </summary>
    [CreateSyncVersion]
    public async Task<List<SheetInfo>> GetSheetInformationsAsync(string path, CancellationToken cancellationToken = default)
    {
        var stream = FileHelper.OpenSharedRead(path);
        await using var disposableStream = stream.ConfigureAwait(false); 

        return await GetSheetInformationsAsync(stream, cancellationToken).ConfigureAwait(false);
    }

    /// <param name="stream">The stream containing the Excel file data. The stream position is not reset after reading.</param>
    /// <param name="cancellationToken">A token to cancel the asynchronous operation.</param>
    [CreateSyncVersion]
    public async Task<List<SheetInfo>> GetSheetInformationsAsync(Stream stream, CancellationToken cancellationToken = default)
    {
        var archive = await OpenXmlZip.CreateAsync(stream, cancellationToken: cancellationToken).ConfigureAwait(false);

        await using var disposableArchve = archive.ConfigureAwait(false);
        using var reader = await OpenXmlReader.CreateAsync(stream, null, cancellationToken: cancellationToken).ConfigureAwait(false);

        var rels = await reader.GetWorkbookRelsAsync(archive.EntryCollection, cancellationToken).ConfigureAwait(false);
        return rels?.Select((s, i) => s.ToSheetInfo((uint)i)).ToList() ?? [];
    }

    [CreateSyncVersion]
    public async Task<IList<ExcelRange>> GetSheetDimensionsAsync(string path, CancellationToken cancellationToken = default)
    {
        var stream = FileHelper.OpenSharedRead(path);
        await using var disposableStream = stream.ConfigureAwait(false); 

        return await GetSheetDimensionsAsync(stream, cancellationToken).ConfigureAwait(false);
    }

    [CreateSyncVersion]
    public async Task<IList<ExcelRange>> GetSheetDimensionsAsync(Stream stream, CancellationToken cancellationToken = default)
    {
        using var reader = await OpenXmlReader.CreateAsync(stream, null, cancellationToken: cancellationToken).ConfigureAwait(false);
        return await reader.GetDimensionsAsync(cancellationToken).ConfigureAwait(false);
    }

    [CreateSyncVersion]
    public async Task<ICollection<string>> GetColumnNamesAsync(string path, bool hasHeaderRow = false,
        string? sheetName = null, string startCell = "A1", CancellationToken cancellationToken = default)
    {
        var stream = FileHelper.OpenSharedRead(path);
        await using var disposableStream = stream.ConfigureAwait(false); 

        return await GetColumnNamesAsync(stream, hasHeaderRow, sheetName, startCell, cancellationToken).ConfigureAwait(false);
    }

    /// <param name="hasHeaderRow">If true, the first row values are used as column names. If false, column letters (A, B, C, etc.) are used. Default is false.</param>
    /// <param name="sheetName">The name of the worksheet to query. If not provided, the first sheet is used.</param>
    [CreateSyncVersion]
    public async Task<ICollection<string>> GetColumnNamesAsync(Stream stream, bool hasHeaderRow = false, 
        string? sheetName = null, string startCell = "A1", CancellationToken cancellationToken = default)
    {
        var enumerator = QueryAsync(stream, hasHeaderRow, sheetName, startCell, null, cancellationToken).GetAsyncEnumerator(cancellationToken);
        await using var disposableEnumerator = enumerator.ConfigureAwait(false);

        if (await enumerator.MoveNextAsync().ConfigureAwait(false))
            return (enumerator.Current as IDictionary<string, object?>)?.Keys ?? [];

        return [];
    }

    [CreateSyncVersion]
    public async Task<CommentResultSet> RetrieveCommentsAsync(string path, string? sheetName, CancellationToken cancellationToken = default)
    {
        var stream = FileHelper.OpenSharedRead(path);
        await using var disposableStream = stream.ConfigureAwait(false); 

        return await RetrieveCommentsAsync(stream, sheetName, cancellationToken).ConfigureAwait(false);
    }

    [CreateSyncVersion]
    public async Task<CommentResultSet> RetrieveCommentsAsync(Stream stream, string? sheetName, CancellationToken cancellationToken = default)
    {
        using var reader = await OpenXmlReader.CreateAsync(stream, null, cancellationToken).ConfigureAwait(false);
        return await reader.ReadCommentsAsync(sheetName, cancellationToken).ConfigureAwait(false);
    }

    #endregion

    #region DataReader

    /// <summary>
    /// Gets an <see cref="IDataReader" /> for the Excel document at the specified path.
    /// </summary>
    /// <exception cref="InvalidOperationException">
    /// Asynchronous reads are not allowed when creating the data reader from this overload and will result in an exception.
    /// </exception>
    public MiniExcelDataReader GetDataReader(string path, bool useHeaderRow = false,
    public MiniExcelDataReader GetDataReader(string path, bool hasHeaderRow = false,
        string? sheetName = null, string startCell = "A1", OpenXmlConfiguration? configuration = null)
    {
        var stream = FileHelper.OpenSharedRead(path);
        var values = Query(stream, hasHeaderRow, sheetName, startCell, configuration).Cast<IDictionary<string, object?>>();

        return MiniExcelDataReader.Create(stream, values, leaveOpen: false);
    }

    /// <summary>
    /// Gets an <see cref="IDataReader" /> for the Excel document from an underlying stream.
    /// </summary>
    /// <exception cref="InvalidOperationException">
    /// Asynchronous reads are not allowed when creating the data reader from this overload and will result in an exception.
    /// </exception>
    public MiniExcelDataReader GetDataReader(Stream stream, bool useHeaderRow = false,
        string? sheetName = null, string startCell = "A1", OpenXmlConfiguration? configuration = null)
    public MiniExcelDataReader GetDataReader(Stream stream, bool hasHeaderRow = false,
        string? sheetName = null, string startCell = "A1", OpenXmlConfiguration? configuration = null, bool leaveOpen = false)
    {
        var values = Query(stream, hasHeaderRow, sheetName, startCell, configuration).Cast<IDictionary<string, object?>>();
        return MiniExcelDataReader.Create(stream, values, leaveOpen);
    }

    /// <summary>
    /// Gets an <see cref="IDataReader" /> for the Excel document at the specific path.
    /// When created from this overload, the resulting data reader is supposed to be advanced asynchronously.
    /// </summary>
    public async Task<MiniExcelDataReader> GetAsyncDataReader(string path, bool useHeaderRow = false,
        string? sheetName = null, string startCell = "A1", OpenXmlConfiguration? configuration = null, 
        CancellationToken cancellationToken = default)
    public async Task<MiniExcelDataReader> GetAsyncDataReader(string path, bool hasHeaderRow = false,
        string? sheetName = null, string startCell = "A1", OpenXmlConfiguration? configuration = null, CancellationToken cancellationToken = default)
    {
        var stream = FileHelper.OpenSharedRead(path);
        var values = QueryAsync(stream, useHeaderRow, sheetName, startCell, configuration, cancellationToken);
        var values = QueryAsync(stream, hasHeaderRow, sheetName, startCell, configuration, cancellationToken);
        
        return await MiniExcelDataReader.CreateAsync(stream, values.CastToDictionary(cancellationToken), leaveOpen: false).ConfigureAwait(false);
    }

    /// <summary>
    /// Gets an <see cref="IDataReader" /> for the Excel document from an underlying stream.
    /// When created from this overload, the resulting data reader is supposed to be advanced asynchronously.
    /// </summary>
    public async Task<MiniExcelDataReader> GetAsyncDataReader(Stream stream, bool useHeaderRow = false,
        string? sheetName = null, string startCell = "A1", OpenXmlConfiguration? configuration = null,
    public async Task<MiniExcelDataReader> GetAsyncDataReader(Stream stream, bool hasHeaderRow = false,
        string? sheetName = null, string startCell = "A1", OpenXmlConfiguration? configuration = null, bool leaveOpen = false,
        CancellationToken cancellationToken = default)
    {
        var values = QueryAsync(stream, useHeaderRow, sheetName, startCell, configuration, cancellationToken);
        return await MiniExcelDataReader.CreateAsync(stream, values.CastToDictionary(cancellationToken)).ConfigureAwait(false);
        var values = QueryAsync(stream, hasHeaderRow, sheetName, startCell, configuration, cancellationToken);
        return await MiniExcelDataReader.CreateAsync(stream, values.CastToDictionary(cancellationToken), leaveOpen).ConfigureAwait(false);
    }

    #endregion
}
