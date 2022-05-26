using System;
using System.IO;
using System.Text;

namespace MiniExcelLibs.OpenXml
{
    internal class MiniExcelStreamWriter : IDisposable
    {
        private readonly Stream _stream;
        private readonly Encoding _encoding;
        private bool disposedValue;

        public MiniExcelStreamWriter(Stream stream,Encoding encoding)
        {
            this._stream = stream;
            this._encoding = encoding;
        }
        public void Write(string content)
        {
            var bytes = this._encoding.GetBytes(content);
            this._stream.Write(bytes, 0, bytes.Length);
        }

        protected virtual void Dispose(bool disposing)
        {
            if (!disposedValue)
            {
                if (disposing)
                {
                    // TODO: dispose managed state (managed objects)
                }

                // free unmanaged resources (unmanaged objects) and override finalizer
                this._stream?.Dispose();
                // TODO: set large fields to null
                disposedValue = true;
            }
        }

        // // TODO: override finalizer only if 'Dispose(bool disposing)' has code to free unmanaged resources
        // ~MiniExcelStreamWriter()
        // {
        //     // Do not change this code. Put cleanup code in 'Dispose(bool disposing)' method
        //     Dispose(disposing: false);
        // }

        public void Dispose()
        {
            // Do not change this code. Put cleanup code in 'Dispose(bool disposing)' method
            Dispose(disposing: true);
            GC.SuppressFinalize(this);
        }
    }
}
