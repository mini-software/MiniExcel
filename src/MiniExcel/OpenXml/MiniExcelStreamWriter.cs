using System;
using System.IO;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace MiniExcelLibs.OpenXml
{
    internal partial class MiniExcelStreamWriter : IDisposable
    {
        private readonly StreamWriter _streamWriter;
        private bool _disposedValue;

        public MiniExcelStreamWriter(Stream stream, Encoding encoding, int bufferSize)
        {
            _streamWriter = new StreamWriter(stream, encoding, bufferSize);
        }

        [Zomp.SyncMethodGenerator.CreateSyncVersion]
        public async Task WriteAsync(string content, CancellationToken cancellationToken = default)
        {
            cancellationToken.ThrowIfCancellationRequested();

            if (string.IsNullOrEmpty(content))
                return;
            await _streamWriter.WriteAsync(content).ConfigureAwait(false);
        }

        [Zomp.SyncMethodGenerator.CreateSyncVersion]
        public async Task<long> WriteAndFlushAsync(string content, CancellationToken cancellationToken = default)
        {
            cancellationToken.ThrowIfCancellationRequested();

            await WriteAsync(content, cancellationToken).ConfigureAwait(false);
            return await FlushAsync().ConfigureAwait(false);
        }

        [Zomp.SyncMethodGenerator.CreateSyncVersion]
        public async Task WriteWhitespaceAsync(int length)
        {
            await _streamWriter.WriteAsync(new string(' ', length)).ConfigureAwait(false);
        }

        [Zomp.SyncMethodGenerator.CreateSyncVersion]
        public async Task<long> FlushAsync()
        {
            await _streamWriter.FlushAsync().ConfigureAwait(false);
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
}