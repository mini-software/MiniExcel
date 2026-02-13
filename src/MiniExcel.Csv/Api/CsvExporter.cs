using MiniExcelLib.Core;

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
            var rowsWritten = await ExportAsync(path, value, printHeader, false, configuration, progress, cancellationToken).ConfigureAwait(false);
            return rowsWritten.FirstOrDefault();
        }

        using var stream = new FileStream(path, FileMode.Append, FileAccess.Write, FileShare.Read, 4096, FileOptions.SequentialScan);
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
    public async Task<int[]> ExportAsync(string path, object value, bool printHeader = true, bool overwriteFile = false,
        CsvConfiguration? configuration = null, IProgress<int>? progress = null, CancellationToken cancellationToken = default)
    {
        using var stream = overwriteFile ? File.Create(path) : new FileStream(path, FileMode.CreateNew);
        return await ExportAsync(stream, value, printHeader, configuration, progress, cancellationToken).ConfigureAwait(false);
    }

    [CreateSyncVersion]
    public async Task<int[]> ExportAsync(Stream stream, object value, bool printHeader = true,
        CsvConfiguration? configuration = null, IProgress<int>? progress = null, CancellationToken cancellationToken = default)
    {
        using var writer = new CsvWriter(stream, value, printHeader, configuration);
        return await writer.SaveAsAsync(progress, cancellationToken).ConfigureAwait(false);
    }

    #endregion
}