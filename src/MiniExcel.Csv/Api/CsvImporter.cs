using MiniExcelLib.Core;

// ReSharper disable once CheckNamespace
namespace MiniExcelLib.Csv;

public partial class CsvImporter
{
    internal CsvImporter() { }
    
    
    #region Query

    /// <summary>
    /// Queries a CSV document using a strongly-typed class model.
    /// </summary>
    /// <typeparam name="T">The class type to map each row to. Must have a parameterless constructor.</typeparam>
    /// <param name="path">The path to the CSV document.</param>
    /// <param name="treatHeaderAsData">If true, the first row is treated as data. If false (default), the first row is used as headers.</param>
    /// <param name="configuration">Optional configuration settings (delimiters, encoding, etc.).</param>
    /// <param name="cancellationToken">A token to cancel the asynchronous operation.</param>
    [CreateSyncVersion]
    public async IAsyncEnumerable<T> QueryAsync<T>(string path, bool treatHeaderAsData = false,
        CsvConfiguration? configuration = null, [EnumeratorCancellation] CancellationToken cancellationToken = default)
        where T : class, new()
    {
        var stream = FileHelper.OpenSharedRead(path);
        await using var disposableStream = stream.ConfigureAwait(false);

        var query = QueryAsync<T>(stream, treatHeaderAsData, configuration, leaveOpen: false, cancellationToken);
        await foreach (var item in query.ConfigureAwait(false))
            yield return item;
    }

    /// <summary>
    /// Queries a CSV document using a strongly-typed class model.
    /// </summary>
    /// <typeparam name="T">The class type to map each row to. Must have a parameterless constructor.</typeparam>
    /// <param name="stream">The stream containing the CSV data.</param>
    /// <param name="treatHeaderAsData">If true, the first row is treated as data. If false (default), the first row is used as headers.</param>
    /// <param name="configuration">Optional configuration settings (delimiters, encoding, etc.).</param>
    /// <param name="leaveOpen">True to leave the stream open after the query is completed, otherwise false.</param>
    /// <param name="cancellationToken">A token to cancel the asynchronous operation.</param>
    [CreateSyncVersion]
    public async IAsyncEnumerable<T> QueryAsync<T>(Stream stream, bool treatHeaderAsData = false, 
        CsvConfiguration? configuration = null, bool leaveOpen = false, [EnumeratorCancellation] CancellationToken cancellationToken = default)
        where T : class, new()
    {
        var reader = new CsvReader(stream, configuration, leaveOpen);
        await using var disposableReader = reader.ConfigureAwait(false);

        await foreach (var item in reader.QueryAsync<T>(null, "A1", treatHeaderAsData, cancellationToken).ConfigureAwait(false))
            yield return item;
    }

    /// <summary>
    /// Queries a CSV document and returns dynamic objects representing each row.
    /// </summary>
    /// <param name="path">The path to the CSV document.</param>
    /// <param name="hasHeaderRow">If true, the first row is used as column headers for the dynamic object properties. Default is false.</param>
    /// <param name="configuration">Optional configuration settings (delimiters, encoding, etc.).</param>
    /// <param name="cancellationToken">A token to cancel the asynchronous operation.</param>
    /// <remarks>
    /// When <paramref name="hasHeaderRow"/> is true, column names from the first row become dynamic property names, otherwise they will be assigned alphabetically (A, B, C, etc.).
    /// </remarks>
    [CreateSyncVersion]
    public async IAsyncEnumerable<dynamic> QueryAsync(string path, bool hasHeaderRow = false, 
        CsvConfiguration? configuration = null, [EnumeratorCancellation] CancellationToken cancellationToken = default)
    {
        var stream = FileHelper.OpenSharedRead(path);
        await using var disposableStream = stream.ConfigureAwait(false);

        await foreach (var item in QueryAsync(stream, hasHeaderRow, configuration, leaveOpen: false, cancellationToken).ConfigureAwait(false))
            yield return item;
    }

    /// <summary>
    /// Queries a CSV document and returns dynamic objects representing each row.
    /// </summary>
    /// <param name="stream">The stream containing the CSV data.</param>
    /// <param name="hasHeaderRow">If true, the first row is used as column headers for the dynamic object properties. Default is false.</param>
    /// <param name="configuration">Optional configuration settings (delimiters, encoding, etc.).</param>
    /// <param name="leaveOpen">True to leave the stream open after the query is completed, otherwise false.</param>
    /// <param name="cancellationToken">A token to cancel the asynchronous operation.</param>
    /// <remarks>
    /// When <paramref name="hasHeaderRow"/> is true, column names from the first row become dynamic property names, otherwise they will be assigned alphabetically (A, B, C, etc.).
    /// </remarks>
    [CreateSyncVersion]
    public async IAsyncEnumerable<dynamic> QueryAsync(Stream stream, bool hasHeaderRow = false,
        CsvConfiguration? configuration = null, bool leaveOpen = false, [EnumeratorCancellation] CancellationToken cancellationToken = default)
    {
        var reader = new CsvReader(stream, configuration, leaveOpen);
        await using var disposableReader = reader.ConfigureAwait(false);

        await foreach (var item in reader.QueryAsync(hasHeaderRow, null, "A1", cancellationToken).ConfigureAwait(false))
            yield return item;
    }

    #endregion

    #region Query As DataTable

    /// <summary>
    /// Queries a CSV file and returns the results as a <see cref="DataTable"/>.
    /// </summary>
    /// <param name="path">The path to the CSV document.</param>
    /// <param name="hasHeaderRow">If true, the first row is used as column headers.</param>
    /// <param name="configuration">Optional configuration settings.</param>
    /// <param name="cancellationToken">A token to cancel the asynchronous operation.</param>
    /// <remarks>
    /// Empty column names are skipped.
    /// This method loads the entire file into memory, so its usage is recommended only for datasets of moderate size.
    /// </remarks>
    [CreateSyncVersion]
    public async Task<DataTable> QueryAsDataTableAsync(string path, bool hasHeaderRow = true,
        CsvConfiguration? configuration = null, CancellationToken cancellationToken = default)
    {
        var stream = FileHelper.OpenSharedRead(path);
        await using var disposableStream = stream.ConfigureAwait(false);

        return await QueryAsDataTableAsync(stream, hasHeaderRow, configuration, leaveOpen: false, cancellationToken).ConfigureAwait(false);
    }

    /// <summary>
    /// Queries a CSV stream and returns the results as a <see cref="DataTable"/>.
    /// </summary>
    /// <param name="stream">The stream containing the CSV data.</param>
    /// <param name="hasHeaderRow">If true, the first row is used as column headers.</param>
    /// <param name="configuration">Optional configuration settings.</param>
    /// <param name="leaveOpen">True to leave the stream open after the query is completed, otherwise false.</param>
    /// <param name="cancellationToken">A token to cancel the asynchronous operation.</param>
    /// <remarks>
    /// Empty column names are skipped.
    /// This method loads the entire file into memory, so its usage is recommended only for datasets of moderate size.
    /// </remarks>
    [CreateSyncVersion]
    public async Task<DataTable> QueryAsDataTableAsync(Stream stream, bool hasHeaderRow = true,
        CsvConfiguration? configuration = null, bool leaveOpen = false, CancellationToken cancellationToken = default)
    {
        var dt = new DataTable();
        var first = true;
        var reader = new CsvReader(stream, configuration, leaveOpen);
        await using var disposableReader = reader.ConfigureAwait(false);

        var rows = reader.QueryAsync(false, null, "A1", cancellationToken);

        var columnDict = new Dictionary<string, string>();
        await foreach (var row in rows.ConfigureAwait(false))
        {
            if (first)
            {
                foreach (var entry in row)
                {
                    cancellationToken.ThrowIfCancellationRequested();

                    var columnName = hasHeaderRow ? entry.Value?.ToString() : entry.Key;
                    if (columnName is { Length: > 0 }) // avoid #298 : Column '' does not belong to table
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

    #region Info

    /// <summary>
    /// Retrieves the column names from the first row (header row) of a CSV document.
    /// </summary>
    /// <param name="path">The path to the CSV document.</param>
    /// <param name="hasHeaderRow">If true, the first row values are used as column names. If false, column letters (A, B, C, etc.) are used. Default is false.</param>
    /// <param name="configuration">Optional configuration settings (delimiters, encoding, etc.).</param>
    /// <param name="cancellationToken">A token to cancel the asynchronous operation.</param>
    /// <returns>A collection of column names from the specified location, or an empty collection if the sheet is empty.</returns>
    [CreateSyncVersion]
    public async Task<ICollection<string>> GetColumnNamesAsync(string path, bool hasHeaderRow = false,
        CsvConfiguration? configuration = null, CancellationToken cancellationToken = default)
    {
        var stream = FileHelper.OpenSharedRead(path);
        await using var disposableStream = stream.ConfigureAwait(false);
        return await GetColumnNamesAsync(stream, hasHeaderRow, configuration, false, cancellationToken).ConfigureAwait(false);
    }

    /// <summary>
    /// Retrieves the column names from the first row (header row) of a CSV document.
    /// </summary>
    /// <param name="stream">The stream containing the CSV data.</param>
    /// <param name="hasHeaderRow">If true, the first row values are used as column names. If false, column letters (A, B, C, etc.) are used. Default is false.</param>
    /// <param name="configuration">Optional configuration settings (delimiters, encoding, etc.).</param>
    /// <param name="leaveOpen">True to leave the stream open after the operation is completed, otherwise false.</param>
    /// <param name="cancellationToken">A token to cancel the asynchronous operation.</param>
    /// <returns>A collection of column names from the specified location, or an empty collection if the sheet is empty.</returns>
    [CreateSyncVersion]
    public async Task<ICollection<string>> GetColumnNamesAsync(Stream stream, bool hasHeaderRow = false,
        CsvConfiguration? configuration = null, bool leaveOpen = false, CancellationToken cancellationToken = default)
    {
        var enumerator = QueryAsync(stream, hasHeaderRow, configuration, leaveOpen: leaveOpen, cancellationToken).GetAsyncEnumerator(cancellationToken);
        await using var disposableEnumerator = enumerator.ConfigureAwait(false);

        if (await enumerator.MoveNextAsync().ConfigureAwait(false))
            return (enumerator.Current as IDictionary<string, object>)?.Keys ?? [];

        return [];
    }

    #endregion

    #region DataReader

    /// <summary>
    /// Gets an <see cref="IDataReader" /> for the CSV document provided for synchronous reading.
    /// </summary>
    /// <param name="path">The path to the CSV document.</param>
    /// <param name="hasHeaderRow">If true, the first row is used as column headers. Default is false.</param>
    /// <param name="configuration">Optional configuration settings (delimiters, encoding, etc.).</param>    /// <remarks>
    /// The returned <see cref="MiniExcelDataReader"/> implements <see cref="IDataReader"/> and supports its standard reading patterns.
    /// The data reader returned by this method is designed to perform synchronous, blocking reads, and will throw <exception cref="InvalidOperationException" /> if an asynchronous operation is called from it.
    /// For asynchronous reading scenarios, use <see cref="GetAsyncDataReader(string, bool, CsvConfiguration?, CancellationToken)"/> instead.
    /// </remarks>
    public MiniExcelDataReader GetDataReader(string path, bool hasHeaderRow = false, CsvConfiguration? configuration = null)
    {
        var stream = FileHelper.OpenSharedRead(path);
        var values = Query(stream, hasHeaderRow, configuration, leaveOpen: false).Cast<IDictionary<string, object?>>();

        return MiniExcelDataReader.Create(stream, values);
    }

    /// <summary>
    /// Gets an <see cref="IDataReader" /> for the CSV document provided for synchronous reading.
    /// </summary>
    /// <param name="stream">The stream containing the CSV data.</param>
    /// <param name="hasHeaderRow">If true, the first row is used as column headers. Default is false.</param>
    /// <param name="configuration">Optional configuration settings (delimiters, encoding, etc.).</param>
    /// <param name="leaveOpen">True to leave the stream open after the data reader is disposed, otherwise false.</param>
    /// <remarks>
    /// The returned <see cref="MiniExcelDataReader"/> implements <see cref="IDataReader"/> and supports its standard reading patterns.
    /// The data reader returned by this method is designed to perform synchronous, blocking reads, and will throw <exception cref="InvalidOperationException" /> if an asynchronous operation is called from it.
    /// For asynchronous reading scenarios, use <see cref="GetAsyncDataReader(Stream, bool, CsvConfiguration?, bool, CancellationToken)"/> instead.
    /// </remarks>
    public MiniExcelDataReader GetDataReader(Stream stream, bool hasHeaderRow = false, CsvConfiguration ? configuration = null, bool leaveOpen = false)
    {
        var values = Query(stream, hasHeaderRow, configuration, leaveOpen).Cast<IDictionary<string, object?>>();
        return MiniExcelDataReader.Create(stream, values, leaveOpen);
    }

    /// <summary>
    /// Gets an <see cref="IDataReader" /> for the CSV document provided for synchronous reading.
    /// </summary>
    /// <param name="path">The path to the CSV document.</param>
    /// <param name="hasHeaderRow">If true, the first row is used as column headers. Default is false.</param>
    /// <param name="configuration">Optional configuration settings (delimiters, encoding, etc.).</param>    /// <param name="cancellationToken">A token to cancel the asynchronous operation.</param>
    /// <remarks>
    /// The returned <see cref="MiniExcelDataReader"/> implements <see cref="IDataReader"/> and supports its standard reading patterns.
    /// The data reader returned by this method is designed to supports asynchronous reads, but will not throw an exception if a synchronous operation is performed.
    /// Still, it's advised to use <see cref="GetDataReader(string, bool, CsvConfiguration?)"/> for synchronous reads instead.
    /// </remarks>
    public async Task<MiniExcelDataReader> GetAsyncDataReader(string path, bool hasHeaderRow = false, 
        CsvConfiguration? configuration = null, CancellationToken cancellationToken = default)
    {
        var stream = FileHelper.OpenSharedRead(path);
        var values = QueryAsync(stream, hasHeaderRow, configuration, leaveOpen: false, cancellationToken);
        
        return await MiniExcelDataReader.CreateAsync(stream, values.CastToDictionary(cancellationToken)).ConfigureAwait(false);
    }

    /// <summary>
    /// Gets an <see cref="IDataReader" /> for the CSV document provided for synchronous reading.
    /// </summary>
    /// <param name="stream">The stream containing the CSV data.</param>
    /// <param name="hasHeaderRow">If true, the first row is used as column headers. Default is false.</param>
    /// <param name="configuration">Optional configuration settings (delimiters, encoding, etc.).</param>
    /// <param name="leaveOpen">True to leave the stream open after the data reader is disposed, otherwise false.</param>
    /// <param name="cancellationToken">A token to cancel the asynchronous operation.</param>
    /// <remarks>
    /// The returned <see cref="MiniExcelDataReader"/> implements <see cref="IDataReader"/> and supports its standard reading patterns.
    /// The data reader returned by this method is designed to supports asynchronous reads, but will not throw an exception if a synchronous operation is performed.
    /// Still, it's advised to use <see cref="GetDataReader(Stream, bool, CsvConfiguration?, bool)"/> for synchronous reads instead.
    /// </remarks>
    public async Task<MiniExcelDataReader> GetAsyncDataReader(Stream stream, bool hasHeaderRow = false,
        CsvConfiguration? configuration = null, bool leaveOpen = false, CancellationToken cancellationToken = default)
    {
        var values = QueryAsync(stream, hasHeaderRow, configuration, leaveOpen, cancellationToken);
        return await MiniExcelDataReader.CreateAsync(stream, values.CastToDictionary(cancellationToken), leaveOpen).ConfigureAwait(false);
    }
    
    #endregion
}
