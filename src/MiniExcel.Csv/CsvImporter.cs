using MiniExcelLib.DataReader;

namespace MiniExcelLib.Csv;

public partial class CsvImporter
{
    internal CsvImporter() { }
    
    
    #region Query

    [CreateSyncVersion]
    public async IAsyncEnumerable<T> QueryCsvAsync<T>(string path, bool treatHeaderAsData = false,
        CsvConfiguration? configuration = null, [EnumeratorCancellation] CancellationToken cancellationToken = default)
        where T : class, new()
    {
        using var stream = FileHelper.OpenSharedRead(path);

        var query = QueryCsvAsync<T>(stream, treatHeaderAsData, configuration, cancellationToken);
        
        //Foreach yield return twice reason : https://stackoverflow.com/questions/66791982/ienumerable-extract-code-lazy-loading-show-stream-was-not-readable
        await foreach (var item in query.ConfigureAwait(false))
            yield return item;
    }

    [CreateSyncVersion]
    public async IAsyncEnumerable<T> QueryCsvAsync<T>(Stream stream, bool treatHeaderAsData = false, 
        CsvConfiguration? configuration = null, [EnumeratorCancellation] CancellationToken cancellationToken = default)
        where T : class, new()
    {
        using var csv = new CsvReader(stream, configuration);
        await foreach (var item in csv.QueryAsync<T>(null, "A1", treatHeaderAsData, cancellationToken).ConfigureAwait(false))
            yield return item;
    }

    [CreateSyncVersion]
    public async IAsyncEnumerable<dynamic> QueryCsvAsync(string path, bool useHeaderRow = false, 
        CsvConfiguration? configuration = null, [EnumeratorCancellation] CancellationToken cancellationToken = default)
    {
        using var stream = FileHelper.OpenSharedRead(path);
        await foreach (var item in QueryCsvAsync(stream, useHeaderRow, configuration, cancellationToken).ConfigureAwait(false))
            yield return item;
    }

    [CreateSyncVersion]
    public async IAsyncEnumerable<dynamic> QueryCsvAsync(Stream stream, bool useHeaderRow = false,
        CsvConfiguration? configuration = null, [EnumeratorCancellation] CancellationToken cancellationToken = default)
    {
#pragma warning disable CA2007
        using var excelReader = new CsvReader(stream, configuration);
        await foreach (var item in excelReader.QueryAsync(useHeaderRow, null, "A1", cancellationToken).ConfigureAwait(false))
            yield return item.Aggregate(seed: GetNewExpandoObject(), func: AddPairToDict);
#pragma warning restore CA2007
    }

    #endregion

    #region Query As DataTable

    /// <summary>
    /// QueryAsDataTable is not recommended, because it'll load all data into memory.
    /// </summary>
    [CreateSyncVersion]
    public async Task<DataTable> QueryCsvAsDataTableAsync(string path, bool useHeaderRow = true,
        CsvConfiguration? configuration = null, CancellationToken cancellationToken = default)
    {
        using var stream = FileHelper.OpenSharedRead(path);
        return await QueryCsvAsDataTableAsync(stream, useHeaderRow, configuration, cancellationToken).ConfigureAwait(false);
    }

    /// <summary>
    /// QueryAsDataTable is not recommended, because it'll load all data into memory.
    /// </summary>
    [CreateSyncVersion]
    public async Task<DataTable> QueryCsvAsDataTableAsync(Stream stream, bool useHeaderRow = true,
        CsvConfiguration? configuration = null, CancellationToken cancellationToken = default)
    {
        var dt = new DataTable();
        var first = true;
        using var reader = new CsvReader(stream, configuration);
        var rows = reader.QueryAsync(false, null, "A1", cancellationToken);

        var columnDict = new Dictionary<string, string>();
#pragma warning disable CA2007
        await foreach (var row in rows.ConfigureAwait(false))
#pragma warning restore CA2007
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
                        columnDict.Add(entry.Key, columnName!); //same column name throw exception???
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

    #region Info

    [CreateSyncVersion]
    public async Task<ICollection<string>> GetCsvColumnsAsync(string path, bool useHeaderRow = false,
        CsvConfiguration? configuration = null, CancellationToken cancellationToken = default)
    {
        using var stream = FileHelper.OpenSharedRead(path);
        return await GetCsvColumnsAsync(stream, useHeaderRow, configuration, cancellationToken).ConfigureAwait(false);
    }

    [CreateSyncVersion]
    public async Task<ICollection<string>> GetCsvColumnsAsync(Stream stream, bool useHeaderRow = false,
        CsvConfiguration? configuration = null, CancellationToken cancellationToken = default)
    {
#pragma warning disable CA2007 // Consider calling ConfigureAwait on the awaited task
        await using var enumerator = QueryCsvAsync(stream, useHeaderRow, configuration, cancellationToken).GetAsyncEnumerator(cancellationToken);
#pragma warning restore CA2007 // Consider calling ConfigureAwait on the awaited task

        _ = enumerator.ConfigureAwait(false);
        if (await enumerator.MoveNextAsync().ConfigureAwait(false))
        {
            return (enumerator.Current as IDictionary<string, object>)?.Keys ?? [];
        }

        return [];
    }

    #endregion

    #region DataReader

    public MiniExcelDataReader GetCsvDataReader(string path, bool useHeaderRow = false, CsvConfiguration? configuration = null)
    {
        var stream = FileHelper.OpenSharedRead(path);
        var values = QueryCsv(stream, useHeaderRow, configuration).Cast<IDictionary<string, object?>>();

        return MiniExcelDataReader.Create(stream, values);
    }

    public MiniExcelDataReader GetCsvDataReader(Stream stream, bool useHeaderRow = false, CsvConfiguration? configuration = null)
    {
        var values = QueryCsv(stream, useHeaderRow, configuration).Cast<IDictionary<string, object?>>();
        return MiniExcelDataReader.Create(stream, values);
    }

    public async Task<MiniExcelAsyncDataReader> GetAsyncCsvDataReader(string path, bool useHeaderRow = false,
        string? sheetName = null, string startCell = "A1", CsvConfiguration? configuration = null, 
        CancellationToken cancellationToken = default)
    {
        var stream = FileHelper.OpenSharedRead(path);
        var values = QueryCsvAsync(stream, useHeaderRow, configuration, cancellationToken);
        
        return await MiniExcelAsyncDataReader.CreateAsync(stream, CastAsync(values, cancellationToken)).ConfigureAwait(false);
    }

    public async Task<MiniExcelAsyncDataReader> GetAsyncCsvDataReader(Stream stream, bool useHeaderRow = false,
        string? sheetName = null, string startCell = "A1", CsvConfiguration? configuration = null,
        CancellationToken cancellationToken = default)
    {
        var values = QueryCsvAsync(stream, useHeaderRow, configuration, cancellationToken);
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