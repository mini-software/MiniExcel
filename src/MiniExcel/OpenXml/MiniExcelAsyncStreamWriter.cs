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
        private bool disposedValue;
        public MiniExcelAsyncStreamWriter(Stream stream, Encoding encoding, int bufferSize, System.Threading.CancellationToken cancellationToken)
        {
            this._stream = stream;
            this._encoding = encoding;
            this._cancellationToken = cancellationToken;
            this._streamWriter = new StreamWriter(stream, this._encoding, bufferSize);
        }

        private int writeTimes = 0;
        public async Task WriteAsync(string content)
        {
            this._cancellationToken.ThrowIfCancellationRequested();

            if (string.IsNullOrEmpty(content))
                return;
            await this._streamWriter.WriteAsync(content);
            if (++writeTimes % 1000 == 0) await this.FlushAsync();
        }

        public async Task<long> WriteAndFlushAsync(string content)
        {
            await this.WriteAsync(content);
            return await this.FlushAsync();
        }

        public async Task<long> FlushAsync()
        {
            this._cancellationToken.ThrowIfCancellationRequested();

            await this._streamWriter.FlushAsync();
            return this._streamWriter.BaseStream.Position;
        }

        protected virtual void Dispose(bool disposing)
        {
            if (!disposedValue)
            {
                this._streamWriter?.Dispose();
                disposedValue = true;
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