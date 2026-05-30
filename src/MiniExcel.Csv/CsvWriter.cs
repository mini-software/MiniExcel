using MiniExcelLib.Core;
using MiniExcelLib.Core.WriteAdapters;
using IMiniExcelWriter = MiniExcelLib.Core.Abstractions.IMiniExcelWriter;

namespace MiniExcelLib.Csv;

internal sealed partial class CsvWriter : IMiniExcelWriter, IDisposable
{
    private readonly StreamWriter _writer;
    private readonly CsvConfiguration _configuration;
    private readonly object? _value;
    private readonly bool _printHeader;

    private bool _disposed;

    // todo: should we add an explicit parameter to leave the stream open instead of the convoluted way to do it through a Func?
    internal CsvWriter(Stream stream, object? value, bool printHeader, IMiniExcelConfiguration? configuration)
    {
        _configuration = configuration as CsvConfiguration ?? CsvConfiguration.Default;
        _writer = _configuration.StreamWriterFunc(stream);
        _printHeader = printHeader;
        _value = value;
    }
    
    private void AppendColumn(StringBuilder rowBuilder, CellWriteInfo column)
    {
        rowBuilder.Append(CsvSanitizer.SanitizeCsvField(ToCsvString(column.Value, column.Mapping), _configuration));
        rowBuilder.Append(_configuration.Seperator);
    }

    private static void RemoveTrailingSeparator(StringBuilder rowBuilder)
    {
        if (rowBuilder.Length is var len and > 0)
        {
            rowBuilder.Remove(len - 1, 1);
        }
    }

    [CreateSyncVersion]
    private async Task<int> WriteValuesAsync(StreamWriter writer, object values, string separator, string newLine,
        IProgress<int>? progress = null, CancellationToken cancellationToken = default)
    {
        cancellationToken.ThrowIfCancellationRequested();

        IMiniExcelWriteAdapter? writeAdapter = null;
        if (!MiniExcelWriteAdapterFactory.TryGetAsyncWriteAdapter(values, _configuration, out var asyncWriteAdapter))
        {
            writeAdapter = MiniExcelWriteAdapterFactory.GetWriteAdapter(values, _configuration);
        }

        try
        {
#if SYNC_ONLY
            var mappings = writeAdapter?.GetColumns();
#else
            var mappings = writeAdapter is not null 
                ? writeAdapter.GetColumns() 
                : await asyncWriteAdapter!.GetColumnsAsync().ConfigureAwait(false);
#endif
    
            if (mappings is null)
            {
#if NET
                await _writer.WriteAsync(_configuration.NewLine.AsMemory(), cancellationToken).ConfigureAwait(false);
                await _writer.FlushAsync(cancellationToken).ConfigureAwait(false);
#else
                await _writer.WriteAsync(_configuration.NewLine).ConfigureAwait(false);
                await _writer.FlushAsync().ConfigureAwait(false);
#endif

                return 0;
            }
    
            if (_printHeader)
            {
#if NET
                await _writer.WriteAsync(GetHeader(mappings).AsMemory(), cancellationToken).ConfigureAwait(false);
                await _writer.WriteAsync(newLine.AsMemory(), cancellationToken).ConfigureAwait(false);
#else
                await _writer.WriteAsync(GetHeader(mappings)).ConfigureAwait(false);
                await _writer.WriteAsync(newLine).ConfigureAwait(false);
#endif
            }
            
            var rowBuilder = new StringBuilder();
            var rowsWritten = 0;
    
            if (writeAdapter is not null)
            {
                foreach (var row in writeAdapter.GetRows(mappings, cancellationToken))
                {
                    rowBuilder.Clear();
                    foreach (var column in row)
                    {
                        cancellationToken.ThrowIfCancellationRequested();
                        AppendColumn(rowBuilder, column);
                        progress?.Report(1);
                    }
    
                    RemoveTrailingSeparator(rowBuilder);
#if NET
                    await _writer.WriteAsync(rowBuilder.ToString().AsMemory(), cancellationToken).ConfigureAwait(false);
                    await _writer.WriteAsync(newLine.AsMemory(), cancellationToken).ConfigureAwait(false);
#else
                    await _writer.WriteAsync(rowBuilder.ToString()).ConfigureAwait(false);
                    await _writer.WriteAsync(newLine).ConfigureAwait(false);
#endif
                    rowsWritten++;
                }
            }
            else
            {
#if !SYNC_ONLY
                await foreach (var row in asyncWriteAdapter!.GetRowsAsync(mappings, cancellationToken).ConfigureAwait(false))
                {
                    cancellationToken.ThrowIfCancellationRequested();
                    rowBuilder.Clear();
    
                    foreach (var column in row)
                    {
                        AppendColumn(rowBuilder, column);
                        progress?.Report(1);
                    }
    
                    RemoveTrailingSeparator(rowBuilder);
#if NET
                    await _writer.WriteAsync(rowBuilder.ToString().AsMemory(), cancellationToken).ConfigureAwait(false);
                    await _writer.WriteAsync(newLine.AsMemory(), cancellationToken).ConfigureAwait(false);
#else
                    await _writer.WriteAsync(rowBuilder.ToString()).ConfigureAwait(false);
                    await _writer.WriteAsync(newLine).ConfigureAwait(false);
#endif
                    rowsWritten++;
                }
#endif
            }
            return rowsWritten;
        }
        finally
        {
#if !SYNC_ONLY
            if (asyncWriteAdapter is not null)
                await asyncWriteAdapter.DisposeAsync().ConfigureAwait(false);
#endif
        }
    }

    [CreateSyncVersion]
    public async Task<int[]> SaveAsAsync(IProgress<int>? progress = null, CancellationToken cancellationToken = default)
    {
        cancellationToken.ThrowIfCancellationRequested();

        var seperator = _configuration.Seperator.ToString();
        var newLine = _configuration.NewLine;

        if (_value is null)
        {
#if NET
            await _writer.WriteAsync("".AsMemory(), cancellationToken).ConfigureAwait(false);
            await _writer.FlushAsync(cancellationToken).ConfigureAwait(false);
#else
            await _writer.WriteAsync("").ConfigureAwait(false);
            await _writer.FlushAsync().ConfigureAwait(false);
#endif
            return [];
        }

        var rowsWritten = await WriteValuesAsync(_writer, _value, seperator, newLine, progress, cancellationToken).ConfigureAwait(false);
        await _writer.FlushAsync(
#if NET
            cancellationToken
#endif
        ).ConfigureAwait(false);

        return [rowsWritten];
    }

    [CreateSyncVersion]
    public async Task<int> InsertAsync(bool overwriteSheet = false, IProgress<int>? progress = null, CancellationToken cancellationToken = default)
    {
        var rowsWritten = await SaveAsAsync(progress, cancellationToken).ConfigureAwait(false);
        return rowsWritten.FirstOrDefault();
    }

    private string ToCsvString(object? value, MiniExcelColumnMapping? mapping)
    {
        if (value is null)
            return "";

        if (value is DateTime dateTime)
        {
            if (mapping?.ExcelFormat is not null)
                return dateTime.ToString(mapping.ExcelFormat, _configuration.Culture);
            
            return _configuration.Culture.Equals(CultureInfo.InvariantCulture)
                ? dateTime.ToString("yyyy-MM-dd HH:mm:ss", _configuration.Culture)
                : dateTime.ToString(_configuration.Culture);
        }

        if (mapping?.ExcelFormat is not null && value is IFormattable formattableValue)
            return formattableValue.ToString(mapping.ExcelFormat, _configuration.Culture);

        return Convert.ToString(value, _configuration.Culture) ?? "";
    }
    
    private string GetHeader(List<MiniExcelColumnMapping> mappings) => string.Join(
        _configuration.Seperator.ToString(),
        mappings.Select(s => CsvSanitizer.SanitizeCsvField(s?.ExcelColumnName, _configuration)));
    
    public void Dispose()
    {
        if (!_disposed)
        {
            _writer.Dispose();
            _disposed = true;
        }
    }
}
