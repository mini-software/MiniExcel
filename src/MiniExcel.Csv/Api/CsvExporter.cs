// ReSharper disable once CheckNamespace
namespace MiniExcelLib.Csv;

public partial class CsvExporter
{
    internal CsvExporter() { }

    /// <summary>
    /// Appends data rows to an existing CSV file without overwriting existing content.
    /// </summary>
    /// <param name="path">The path to the CSV file to append to.</param>
    /// <param name="value">The data to append. Can be an object, <see cref="IEnumerable"/>, <see cref="DataTable"/>, or <see cref="IDataReader"/>.</param>
    /// <param name="printHeader">If true, when the file does not exist already the header row is added to the output. Default is true</param>
    /// <param name="configuration">Optional configuration settings (delimiters, encoding, etc.).</param>
    /// <param name="progress">An optional <see cref="IProgress{T}"/> to report progress as values are written.</param>
    /// <param name="cancellationToken">A token to cancel the asynchronous operation.</param>
    /// <returns>The number of rows appended.</returns>
    [CreateSyncVersion]
    public async Task<int> AppendAsync(string path, object value, bool printHeader = true, 
        CsvConfiguration? configuration = null, IProgress<int>? progress = null, CancellationToken cancellationToken = default)
    {
        if (!File.Exists(path))
        {
            return await ExportAsync(path, value, printHeader, false, configuration, progress, cancellationToken).ConfigureAwait(false);
        }

        var stream = new FileStream(path, FileMode.Append, FileAccess.Write, FileShare.Read, 4096, FileOptions.SequentialScan);
        await using var disposableStream = stream.ConfigureAwait(false);

        return await AppendAsync(stream, value, configuration, progress, cancellationToken).ConfigureAwait(false);
    }

    /// <summary>
    /// Appends data rows to an existing CSV stream without overwriting existing content.
    /// </summary>
    /// <param name="stream">The stream containing the CSV file to append to. The stream will be positioned at the end before appending.</param>
    /// <param name="value">The data to append. Can be an object, <see cref="IEnumerable"/>, <see cref="DataTable"/>, or <see cref="IDataReader"/>.</param>
    /// <param name="configuration">Optional configuration settings (delimiters, encoding, etc.).</param>
    /// <param name="progress">An optional <see cref="IProgress{T}"/> to report progress as values are written.</param>
    /// <param name="cancellationToken">A token to cancel the asynchronous operation.</param>
    /// <returns>The number of rows appended.</returns>
    [CreateSyncVersion]
    public async Task<int> AppendAsync(Stream stream, object value, CsvConfiguration? configuration = null, IProgress<int>? progress = null, CancellationToken cancellationToken = default)
    {
        stream.Seek(0, SeekOrigin.End);

        var newValue = value is IEnumerable or IDataReader ? value : new[] { value };

        using var writer = new CsvWriter(stream, newValue, false, configuration);
        return await writer.InsertAsync(false, progress, cancellationToken).ConfigureAwait(false);
    }

    /// <summary>
    /// Exports data to a CSV file.
    /// </summary>
    /// <param name="path">The path to write CSV data to.</param>
    /// <param name="value">The data to export. Can be an object, <see cref="IEnumerable"/>, <see cref="DataTable"/>, or <see cref="IDataReader"/>.</param>
    /// <param name="printHeader">If true, the first row will contain column headers derived from property names or DataTable column names. Default is true.</param>
    /// <param name="overwriteFile">If true, when a file at the specified path already exists it will be overwritten, otherwise an <see cref="IOException" /> will be thrown. Default is false.</param>
    /// <param name="configuration">Optional configuration settings (delimiters, encoding, etc.).</param>
    /// <param name="progress">An optional <see cref="IProgress{T}"/> to report progress as values are written.</param>
    /// <param name="cancellationToken">A token to cancel the asynchronous operation.</param>
    /// <returns>The number of rows written.</returns>
    [CreateSyncVersion]
    public async Task<int> ExportAsync(string path, object value, bool printHeader = true, bool overwriteFile = false,
        CsvConfiguration? configuration = null, IProgress<int>? progress = null, CancellationToken cancellationToken = default)
    {
        var stream = overwriteFile ? File.Create(path) : new FileStream(path, FileMode.CreateNew);
        await using var disposableStream = stream.ConfigureAwait(false);

        return await ExportAsync(stream, value, printHeader, configuration, progress, cancellationToken).ConfigureAwait(false);
    }

    /// <summary>
    /// Exports data to a CSV stream.
    /// </summary>
    /// <param name="stream">The stream to write CSV data to. Existing content will be overwritten.</param>
    /// <param name="value">The data to export. Can be an object, <see cref="IEnumerable"/>, <see cref="DataTable"/>, or <see cref="IDataReader"/>.</param>
    /// <param name="printHeader">If true, the first row will contain column headers derived from property names or DataTable column names. Default is true.</param>
    /// <param name="configuration">Optional configuration settings (delimiters, encoding, etc.).</param>
    /// <param name="progress">An optional <see cref="IProgress{T}"/> to report progress as values are written.</param>
    /// <param name="cancellationToken">A token to cancel the asynchronous operation.</param>
    /// <returns>The number of rows written.</returns>
    [CreateSyncVersion]
    public async Task<int> ExportAsync(Stream stream, object value, bool printHeader = true,
        CsvConfiguration? configuration = null, IProgress<int>? progress = null, CancellationToken cancellationToken = default)
    {
        using var writer = new CsvWriter(stream, value, printHeader, configuration);
        var result = await writer.SaveAsAsync(progress, cancellationToken).ConfigureAwait(false);

        return result.FirstOrDefault();
    }
}
