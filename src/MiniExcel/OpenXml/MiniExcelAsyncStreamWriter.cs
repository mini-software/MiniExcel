// <copyright file="MiniExcelAsyncStreamWriter.cs" company="Rolls-Royce plc">
// Copyright (c) 2024 Rolls-Royce plc
// </copyright>

namespace MiniExcelLibs.OpenXml
{
    using System;
    using System.IO;
    using System.Text;
    using System.Threading;
    using System.Threading.Tasks;

    internal class MiniExcelAsyncStreamWriter : IDisposable
    {
        private readonly Stream _stream;
        private readonly Encoding _encoding;
        private readonly CancellationToken _cancellationToken;
        private readonly StreamWriter _streamWriter;
        private bool _disposedValue;
        
        public MiniExcelAsyncStreamWriter(Stream stream, Encoding encoding, int bufferSize, CancellationToken cancellationToken)
        {
            _stream = stream;
            _encoding = encoding;
            _cancellationToken = cancellationToken;
            _streamWriter = new StreamWriter(stream, _encoding, bufferSize);
        }
        public async Task WriteAsync(string content)
        {
            _cancellationToken.ThrowIfCancellationRequested();

            if (string.IsNullOrEmpty(content))
                return;
            await _streamWriter.WriteAsync(content).ConfigureAwait(false);
        }

        public async Task<long> WriteAndFlushAsync(string content)
        {
            await WriteAsync(content).ConfigureAwait(false);
            return await FlushAsync().ConfigureAwait(false);
        }

        public async Task WriteWhitespaceAsync(int length)
        {
            await _streamWriter.WriteAsync(new string(' ', length)).ConfigureAwait(false);
        }

        public async Task<long> FlushAsync()
        {
            _cancellationToken.ThrowIfCancellationRequested();

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