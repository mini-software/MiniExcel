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
        //private byte[] _cacheValueBytes;

        public MiniExcelStreamWriter(Stream stream,Encoding encoding, int bufferSize)
        {
            this._stream = stream;
            this._encoding = encoding;
            this._streamWriter = new StreamWriter(stream, this._encoding, bufferSize);
        }
        public void Write(string content,bool flushImmediately=false)
        {
            if (string.IsNullOrEmpty(content))
                return;
            //if (flushImmediately)
            //else
                //_cacheValueBytes.CopyTo
            //TODO:
            //var bytes = this._encoding.GetBytes(content);
            //this._stream.Write(bytes, 0, bytes.Length);

            this._streamWriter.Write(content);
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
                this._streamWriter?.Dispose();
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
