using System;
using System.IO;
using System.Text;

namespace MiniExcelLibs.OpenXml
{
    internal class MiniExcelStreamWriter : IDisposable
    {
        private readonly Stream _stream;
        private readonly Encoding _encoding;
        private readonly StreamWriter _streamWriter;
        private bool disposedValue;
        public MiniExcelStreamWriter(Stream stream, Encoding encoding, int bufferSize)
        {
            this._stream = stream;
            this._encoding = encoding;
            this._streamWriter = new StreamWriter(stream, this._encoding, bufferSize);
        }
        private int writeTimes = 0;
        public void Write(string content)
        {
            if (string.IsNullOrEmpty(content))
                return;
            this._streamWriter.Write(content);
            if (++writeTimes % 1000 == 0) this.Flush();
        }

        public long WriteAndFlush(string content)
        {
            this.Write(content);
            this._streamWriter.Flush();
            return this._streamWriter.BaseStream.Position;
        }

        public long Flush()
        {
            this._streamWriter.Flush();
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
