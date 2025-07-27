using System.Dynamic;
using MiniExcelLib.Core.DataReader;
using MiniExcelLib.Core.OpenXml.Models;
using MiniExcelLib.Core.OpenXml.Zip;

// ReSharper disable once CheckNamespace
namespace MiniExcelLib.Core;

public sealed partial class OpenXmlImporter
{
    internal OpenXmlImporter() { }
    
    #region Query

    [CreateSyncVersion]
    public async IAsyncEnumerable<T> QueryAsync<T>(string path, string? sheetName = null,
        string startCell = "A1", bool treatHeaderAsData = false, OpenXmlConfiguration? configuration = null,
        [EnumeratorCancellation] CancellationToken cancellationToken = default) where T : class, new()
    {
        using var stream = FileHelper.OpenSharedRead(path);

        var query = QueryAsync<T>(stream, sheetName, startCell, treatHeaderAsData, configuration, cancellationToken);
        
        //Foreach yield return twice reason : https://stackoverflow.com/questions/66791982/ienumerable-extract-code-lazy-loading-show-stream-was-not-readable
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
    public async IAsyncEnumerable<dynamic> QueryAsync(string path, bool useHeaderRow = false,
        string? sheetName = null, string startCell = "A1", OpenXmlConfiguration? configuration = null,
        [EnumeratorCancellation] CancellationToken cancellationToken = default)
    {
        using var stream = FileHelper.OpenSharedRead(path);
        await foreach (var item in QueryAsync(stream, useHeaderRow, sheetName, startCell, configuration, cancellationToken).ConfigureAwait(false))
            yield return item;
    }

    [CreateSyncVersion]
    public async IAsyncEnumerable<dynamic> QueryAsync(Stream stream, bool useHeaderRow = false,
        string? sheetName = null, string startCell = "A1", OpenXmlConfiguration? configuration = null,
        [EnumeratorCancellation] CancellationToken cancellationToken = default)
    {
        using var excelReader = await OpenXmlReader.CreateAsync(stream, configuration, cancellationToken).ConfigureAwait(false);
        await foreach (var item in excelReader.QueryAsync(useHeaderRow, sheetName, startCell, cancellationToken).ConfigureAwait(false))
            yield return item.Aggregate(seed: GetNewExpandoObject(), func: AddPairToDict);
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
    public async IAsyncEnumerable<dynamic> QueryRangeAsync(string path, bool useHeaderRow = false,
        string? sheetName = null, string startCell = "A1", string endCell = "", OpenXmlConfiguration? configuration = null,
        [EnumeratorCancellation] CancellationToken cancellationToken = default)
    {
        using var stream = FileHelper.OpenSharedRead(path);
        await foreach (var item in QueryRangeAsync(stream, useHeaderRow, sheetName, startCell, endCell, configuration, cancellationToken).ConfigureAwait(false))
            yield return item;
    }

    [CreateSyncVersion]
    public async IAsyncEnumerable<dynamic> QueryRangeAsync(Stream stream, bool useHeaderRow = false,
        string? sheetName = null, string startCell = "A1", string endCell = "", OpenXmlConfiguration? configuration = null,
        [EnumeratorCancellation] CancellationToken cancellationToken = default)
    {
        using var excelReader = await OpenXmlReader.CreateAsync(stream, configuration, cancellationToken).ConfigureAwait(false);
        await foreach (var item in excelReader.QueryRangeAsync(useHeaderRow, sheetName, startCell, endCell, cancellationToken).ConfigureAwait(false))
            yield return item.Aggregate(seed: GetNewExpandoObject(), func: AddPairToDict);
    }

    [CreateSyncVersion]
    public async IAsyncEnumerable<dynamic> QueryRangeAsync(string path, bool useHeaderRow = false,
        string? sheetName = null, int startRowIndex = 1, int startColumnIndex = 1, int? endRowIndex = null,
        int? endColumnIndex = null, OpenXmlConfiguration? configuration = null,
        [EnumeratorCancellation] CancellationToken cancellationToken = default)
    {
        using var stream = FileHelper.OpenSharedRead(path);
        await foreach (var item in QueryRangeAsync(stream, useHeaderRow, sheetName, startRowIndex, startColumnIndex, endRowIndex, endColumnIndex, configuration, cancellationToken).ConfigureAwait(false))
            yield return item;
    }

    [CreateSyncVersion]
    public async IAsyncEnumerable<dynamic> QueryRangeAsync(Stream stream, bool useHeaderRow = false,
        string? sheetName = null, int startRowIndex = 1, int startColumnIndex = 1, int? endRowIndex = null,
        int? endColumnIndex = null, OpenXmlConfiguration? configuration = null,
        [EnumeratorCancellation] CancellationToken cancellationToken = default)
    {
        using var excelReader = await OpenXmlReader.CreateAsync(stream, configuration, cancellationToken).ConfigureAwait(false);
        await foreach (var item in excelReader.QueryRangeAsync(useHeaderRow, sheetName, startRowIndex, startColumnIndex, endRowIndex, endColumnIndex, cancellationToken).ConfigureAwait(false))
            yield return item.Aggregate(seed: GetNewExpandoObject(), func: AddPairToDict);
    }

    #endregion

    #region Query As DataTable

    /// <summary>
    /// QueryAsDataTable is not recommended, because it'll load all data into memory.
    /// </summary>
    [CreateSyncVersion]
    public async Task<DataTable> QueryAsDataTableAsync(string path, bool useHeaderRow = true,
        string? sheetName = null, string startCell = "A1", OpenXmlConfiguration? configuration = null,
        CancellationToken cancellationToken = default)
    {
        using var stream = FileHelper.OpenSharedRead(path);
        return await QueryAsDataTableAsync(stream, useHeaderRow, sheetName, startCell, configuration, cancellationToken).ConfigureAwait(false);
    }

    /// <summary>
    /// QueryAsDataTable is not recommended, because it'll load all data into memory.
    /// </summary>
    [CreateSyncVersion]
    public async Task<DataTable> QueryAsDataTableAsync(Stream stream, bool useHeaderRow = true,
        string? sheetName = null, string startCell = "A1", OpenXmlConfiguration? configuration = null,
        CancellationToken cancellationToken = default)
    {
        /*Issue #279*/
        sheetName ??= (await GetSheetNamesAsync(stream, configuration, cancellationToken).ConfigureAwait(false)).First();

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

                    var columnName = useHeaderRow ? entry.Value?.ToString() : entry.Key;
                    if (!string.IsNullOrWhiteSpace(columnName)) // avoid #298 : Column '' does not belong to table
                    {
                        var column = new DataColumn(columnName, typeof(object)) { Caption = columnName };
                        dt.Columns.Add(column);
                        columnDict.Add(entry.Key, columnName); //same column name throw exception???
                    }
                }

                dt.BeginLoadData();
                first = false;
                if (useHeaderRow)
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
    public async Task<List<string>> GetSheetNamesAsync(string path, OpenXmlConfiguration? config = null, CancellationToken cancellationToken = default)
    {
        using var stream = FileHelper.OpenSharedRead(path);
        return await GetSheetNamesAsync(stream, config, cancellationToken).ConfigureAwait(false);
    }

    [CreateSyncVersion]
    public async Task<List<string>> GetSheetNamesAsync(Stream stream, OpenXmlConfiguration? config = null, CancellationToken cancellationToken = default)
    {
        config ??= OpenXmlConfiguration.Default;

        using var archive = new OpenXmlZip(stream, leaveOpen: true);

        using var reader = await OpenXmlReader.CreateAsync(stream, config, cancellationToken: cancellationToken).ConfigureAwait(false);
        var rels = await reader.GetWorkbookRelsAsync(archive.EntryCollection, cancellationToken).ConfigureAwait(false);

        return rels?.Select(s => s.Name).ToList() ?? [];
    }

    [CreateSyncVersion]
    public async Task<List<SheetInfo>> GetSheetInformationsAsync(string path, OpenXmlConfiguration? config = null, CancellationToken cancellationToken = default)
    {
        using var stream = FileHelper.OpenSharedRead(path);
        return await GetSheetInformationsAsync(stream, config, cancellationToken).ConfigureAwait(false);
    }

    [CreateSyncVersion]
    public async Task<List<SheetInfo>> GetSheetInformationsAsync(Stream stream, OpenXmlConfiguration? config = null, CancellationToken cancellationToken = default)
    {
        config ??= OpenXmlConfiguration.Default;

        using var archive = new OpenXmlZip(stream);
        using var reader = await OpenXmlReader.CreateAsync(stream, config, cancellationToken: cancellationToken).ConfigureAwait(false);
        var rels = await reader.GetWorkbookRelsAsync(archive.EntryCollection, cancellationToken).ConfigureAwait(false);

        return rels?.Select((s, i) => s.ToSheetInfo((uint)i)).ToList() ?? [];
    }

    [CreateSyncVersion]
    public async Task<IList<ExcelRange>> GetSheetDimensionsAsync(string path, CancellationToken cancellationToken = default)
    {
        using var stream = FileHelper.OpenSharedRead(path);
        return await GetSheetDimensionsAsync(stream, cancellationToken).ConfigureAwait(false);
    }

    [CreateSyncVersion]
    public async Task<IList<ExcelRange>> GetSheetDimensionsAsync(Stream stream, CancellationToken cancellationToken = default)
    {
        using var reader = await OpenXmlReader.CreateAsync(stream, null, cancellationToken: cancellationToken).ConfigureAwait(false);
        return await reader.GetDimensionsAsync(cancellationToken).ConfigureAwait(false);
    }

    [CreateSyncVersion]
    public async Task<ICollection<string>> GetColumnNamesAsync(string path, bool useHeaderRow = false,
        string? sheetName = null, string startCell = "A1", OpenXmlConfiguration? configuration = null,
        CancellationToken cancellationToken = default)
    {
        using var stream = FileHelper.OpenSharedRead(path);
        return await GetColumnNamesAsync(stream, useHeaderRow, sheetName, startCell, configuration, cancellationToken).ConfigureAwait(false);
    }

    [CreateSyncVersion]
    public async Task<ICollection<string>> GetColumnNamesAsync(Stream stream, bool useHeaderRow = false,
        string? sheetName = null, string startCell = "A1", OpenXmlConfiguration? configuration = null,
        CancellationToken cancellationToken = default)
    {
#pragma warning disable CA2007 // We need to assign the AsyncEnumerator before we can call ConfigureAwait on it
        await using var enumerator = QueryAsync(stream, useHeaderRow, sheetName, startCell, configuration, cancellationToken).GetAsyncEnumerator(cancellationToken);
#pragma warning restore CA2007

        _ = enumerator.ConfigureAwait(false);
        if (await enumerator.MoveNextAsync().ConfigureAwait(false))
        {
            return (enumerator.Current as IDictionary<string, object?>)?.Keys ?? [];
        }

        return [];
    }

    #endregion

    #region DataReader

    public MiniExcelDataReader GetDataReader(string path, bool useHeaderRow = false,
        string? sheetName = null, string startCell = "A1", OpenXmlConfiguration? configuration = null)
    {
        var stream = FileHelper.OpenSharedRead(path);
        var values = Query(stream, useHeaderRow, sheetName, startCell, configuration).Cast<IDictionary<string, object?>>();

        return MiniExcelDataReader.Create(stream, values);
    }

    public MiniExcelDataReader GetDataReader(Stream stream, bool useHeaderRow = false,
        string? sheetName = null, string startCell = "A1", OpenXmlConfiguration? configuration = null)
    {
        var values = Query(stream, useHeaderRow, sheetName, startCell, configuration).Cast<IDictionary<string, object?>>();
        return MiniExcelDataReader.Create(stream, values);
    }
    
    public async Task<MiniExcelAsyncDataReader> GetAsyncDataReader(string path, bool useHeaderRow = false,
        string? sheetName = null, string startCell = "A1", OpenXmlConfiguration? configuration = null, 
        CancellationToken cancellationToken = default)
    {
        var stream = FileHelper.OpenSharedRead(path);
        var values = QueryAsync(stream, useHeaderRow, sheetName, startCell, configuration, cancellationToken);
        
        return await MiniExcelAsyncDataReader.CreateAsync(stream, CastAsync(values, cancellationToken)).ConfigureAwait(false);
    }

    public async Task<MiniExcelAsyncDataReader> GetAsyncDataReader(Stream stream, bool useHeaderRow = false,
        string? sheetName = null, string startCell = "A1", OpenXmlConfiguration? configuration = null,
        CancellationToken cancellationToken = default)
    {
        var values = QueryAsync(stream, useHeaderRow, sheetName, startCell, configuration, cancellationToken);
        return await MiniExcelAsyncDataReader.CreateAsync(stream, CastAsync(values, cancellationToken)).ConfigureAwait(false);
    }

    #endregion
    
    private static IDictionary<string, object?> GetNewExpandoObject() => new ExpandoObject();
    private static IDictionary<string, object?> AddPairToDict(IDictionary<string, object?> dict, KeyValuePair<string, object?> pair)
    {
        dict.Add(pair);
        return dict; 
    }

    private static async IAsyncEnumerable<IDictionary<string, object?>> CastAsync(IAsyncEnumerable<dynamic> enumerable, CancellationToken cancellationToken = default)
    {
        await foreach (var item in enumerable.WithCancellation(cancellationToken).ConfigureAwait(false))
        {
            if (item is IDictionary<string, object?> dict)
                yield return dict;
        }
    }
}