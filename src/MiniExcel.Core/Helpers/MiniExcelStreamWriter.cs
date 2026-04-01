namespace MiniExcelLib.Core.Helpers;

public sealed partial class MiniExcelStreamWriter(Stream stream, Encoding encoding, int bufferSize) : IDisposable
#if NET8_0_OR_GREATER
    , IAsyncDisposable
#endif
{
    // if leaveOpen is set to false, the StreamWriter closes the underlying stream synchronously in a finally block.
    // Since we want to avoid all synchronous operations when dealing with streams we leave it open here, as it will disposed from the caller anyways 
    private readonly StreamWriter _streamWriter = new(stream, encoding, bufferSize, true);
    private bool _disposed;

    [CreateSyncVersion]
    public async Task WriteAsync(string content, CancellationToken cancellationToken = default)
    {
        if (!string.IsNullOrEmpty(content))
        {
#if NET8_0_OR_GREATER
            await _streamWriter.WriteAsync(content.AsMemory(), cancellationToken)
#else
            cancellationToken.ThrowIfCancellationRequested();
            await _streamWriter.WriteAsync(content)
#endif
                .ConfigureAwait(false);
        }
    }

    [CreateSyncVersion]
    public async Task<long> WriteAndFlushAsync(string content, CancellationToken cancellationToken = default)
    {
        await WriteAsync(content, cancellationToken).ConfigureAwait(false);
        return await FlushAndGetPositionAsync(cancellationToken).ConfigureAwait(false);
    }

    [CreateSyncVersion]
    public async Task<long> FlushAndGetPositionAsync(CancellationToken cancellationToken = default)
    {
        await _streamWriter.FlushAsync(
#if NET8_0_OR_GREATER
            cancellationToken
#endif
        ).ConfigureAwait(false);
        return _streamWriter.BaseStream.Position;
    }

    public void SetPosition(long position)
    {
        _streamWriter.BaseStream.Position = position;
    }

    public void Dispose()
    {
        if (!_disposed)
        {
            _streamWriter.Dispose();
            _disposed = true;
        }
    }

#if NET8_0_OR_GREATER
    public async ValueTask DisposeAsync()
    {
        if (!_disposed)
        {
            await _streamWriter.DisposeAsync().ConfigureAwait(false);
            _disposed = true;
        }
    }
#endif
}
