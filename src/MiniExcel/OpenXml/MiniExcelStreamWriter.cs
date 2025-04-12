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
            _stream = stream;
            _encoding = encoding;
            _streamWriter = new StreamWriter(stream, _encoding, bufferSize);
        }
        public void Write(string content)
        {
            if (string.IsNullOrEmpty(content))
                return;
            
            _streamWriter.Write(content);
        }

        public long WriteAndFlush(string content)
        {
            Write(content);
            _streamWriter.Flush();
            return _streamWriter.BaseStream.Position;
        }

        public void WriteWhitespace(int length)
        {
            _streamWriter.Write(new string(' ', length));
        }

        public long Flush()
        {
            _streamWriter.Flush();
            return _streamWriter.BaseStream.Position;
        }

        public void SetPosition(long position)
        {
            _streamWriter.BaseStream.Position = position;
        }

        protected virtual void Dispose(bool disposing)
        {
            if (!disposedValue)
            {
                _streamWriter?.Dispose();
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
