﻿using System;
using System.IO;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace MiniExcelLibs.OpenXml;

internal partial class MiniExcelStreamWriter(Stream stream, Encoding encoding, int bufferSize) : IDisposable
{
    private readonly StreamWriter _streamWriter = new(stream, encoding, bufferSize);
    private bool _disposedValue;

    [Zomp.SyncMethodGenerator.CreateSyncVersion]
    public async Task WriteAsync(string content, CancellationToken cancellationToken = default)
    {
        cancellationToken.ThrowIfCancellationRequested();

        if (!string.IsNullOrEmpty(content))
        {
            await _streamWriter.WriteAsync(content).ConfigureAwait(false);
        }
    }

    [Zomp.SyncMethodGenerator.CreateSyncVersion]
    public async Task<long> WriteAndFlushAsync(string content, CancellationToken cancellationToken = default)
    {
        cancellationToken.ThrowIfCancellationRequested();

        await WriteAsync(content, cancellationToken).ConfigureAwait(false);
        return await FlushAsync(cancellationToken).ConfigureAwait(false);
    }

    [Zomp.SyncMethodGenerator.CreateSyncVersion]
    public async Task WriteWhitespaceAsync(int length)
    {
        await _streamWriter.WriteAsync(new string(' ', length)).ConfigureAwait(false);
    }

    [Zomp.SyncMethodGenerator.CreateSyncVersion]
    public async Task<long> FlushAsync(CancellationToken cancellationToken = default)
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

    protected virtual void Dispose(bool disposing)
    {
        if (!_disposedValue)
        {
            _streamWriter?.Dispose();
            _disposedValue = true;
        }
    }

    public void Dispose()
    {
        // Do not change this code. Put cleanup code in 'Dispose(bool disposing)' method
        Dispose(disposing: true);
        GC.SuppressFinalize(this);
    }
}