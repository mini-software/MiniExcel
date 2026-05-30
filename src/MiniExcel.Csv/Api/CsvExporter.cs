// ReSharper disable once CheckNamespace
namespace MiniExcelLib.Csv;

public partial class CsvExporter
{
    internal CsvExporter() { }
    
    
    #region Append / Export
    
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

    [CreateSyncVersion]
    public async Task<int> AppendAsync(Stream stream, object value, CsvConfiguration? configuration = null, IProgress<int>? progress = null, CancellationToken cancellationToken = default)
    {
        stream.Seek(0, SeekOrigin.End);

        var newValue = value is IEnumerable or IDataReader ? value : new[] { value };

        using var writer = new CsvWriter(stream, newValue, false, configuration);
        return await writer.InsertAsync(false, progress, cancellationToken).ConfigureAwait(false);
    }

    [CreateSyncVersion]
    public async Task<int> ExportAsync(string path, object value, bool printHeader = true, bool overwriteFile = false,
        CsvConfiguration? configuration = null, IProgress<int>? progress = null, CancellationToken cancellationToken = default)
    {
        var stream = overwriteFile ? File.Create(path) : new FileStream(path, FileMode.CreateNew);
        await using var disposableStream = stream.ConfigureAwait(false);

        return await ExportAsync(stream, value, printHeader, configuration, progress, cancellationToken).ConfigureAwait(false);
    }

    [CreateSyncVersion]
    public async Task<int> ExportAsync(Stream stream, object value, bool printHeader = true,
        CsvConfiguration? configuration = null, IProgress<int>? progress = null, CancellationToken cancellationToken = default)
    {
        using var writer = new CsvWriter(stream, value, printHeader, configuration);
        var result = await writer.SaveAsAsync(progress, cancellationToken).ConfigureAwait(false);

        return result.FirstOrDefault();
    }
}
