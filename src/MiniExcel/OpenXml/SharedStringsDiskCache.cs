using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Runtime.CompilerServices;
using System.Text;

namespace MiniExcelLibs.OpenXml
{
    internal class SharedStringsDiskCache : IDictionary<int, string>, IDisposable
    {
        private static readonly Encoding _encoding = new UTF8Encoding(true);
        
        private readonly FileStream _positionFs;
        private readonly FileStream _lengthFs;
        private readonly FileStream _valueFs;
        private bool _disposedValue;
        
        private long _maxIndx = -1;
        public int Count => checked((int)(_maxIndx + 1));
        public string this[int key] { get => GetValue(key); set => Add(key, value); }
        public bool ContainsKey(int key)
        {
            return key <= _maxIndx;
        }

        public SharedStringsDiskCache()
        {
            var path = $"{Guid.NewGuid().ToString()}_miniexcelcache";
            _positionFs = new FileStream($"{path}_position", FileMode.OpenOrCreate);
            _lengthFs = new FileStream($"{path}_length", FileMode.OpenOrCreate);
            _valueFs = new FileStream($"{path}_data", FileMode.OpenOrCreate);
        }

        // index must start with 0-N
        internal void Add(int index, string value)
        {
            if (index > _maxIndx)
                _maxIndx = index;
            byte[] valueBs = _encoding.GetBytes(value);
            if (value.Length > 32767) //check info length, becasue cell string max length is 47483647
                throw new ArgumentOutOfRangeException("Excel one cell max length is 32,767 characters");
            _positionFs.Write(BitConverter.GetBytes(_valueFs.Position), 0, 4);
            _lengthFs.Write(BitConverter.GetBytes(valueBs.Length), 0, 4);
            _valueFs.Write(valueBs, 0, valueBs.Length);
        }

        private string GetValue(int index)
        {
            _positionFs.Position = index * 4;
            var bytes = new byte[4];
            _positionFs.Read(bytes, 0, 4);
            var position = BitConverter.ToInt32(bytes, 0);
            _lengthFs.Position = index * 4;
            _lengthFs.Read(bytes, 0, 4);
            var length = BitConverter.ToInt32(bytes, 0);
            _valueFs.Position = position;
            bytes = new byte[length];
            _valueFs.Read(bytes, 0, length);
            var v = _encoding.GetString(bytes);
            return v;
        }

        protected virtual void Dispose(bool disposing)
        {
            if (!_disposedValue)
            {
                if (disposing)
                {
                    // TODO: dispose managed state (managed objects)
                }
                _positionFs.Dispose();
                if (File.Exists(_positionFs.Name))
                    File.Delete(_positionFs.Name);
                _lengthFs.Dispose();
                if (File.Exists(_lengthFs.Name))
                    File.Delete(_lengthFs.Name);
                _valueFs.Dispose();
                if (File.Exists(_valueFs.Name))
                    File.Delete(_valueFs.Name);
                _disposedValue = true;
            }
        }

        ~SharedStringsDiskCache()
        {
            Dispose(disposing: false);
        }

        public void Dispose()
        {
            Dispose(disposing: true);
            GC.SuppressFinalize(this);
        }
        public ICollection<int> Keys => throw new NotImplementedException();
        public ICollection<string> Values => throw new NotImplementedException();
        public bool IsReadOnly => throw new NotImplementedException();
        public bool Remove(int key)
        {
            throw new NotImplementedException();
        }

        public bool TryGetValue(int key, out string value)
        {
            throw new NotImplementedException();
        }

        public void Add(KeyValuePair<int, string> item)
        {
            throw new NotImplementedException();
        }

        public void Clear()
        {
            throw new NotImplementedException();
        }

        public bool Contains(KeyValuePair<int, string> item)
        {
            throw new NotImplementedException();
        }

        public void CopyTo(KeyValuePair<int, string>[] array, int arrayIndex)
        {
            throw new NotImplementedException();
        }

        public bool Remove(KeyValuePair<int, string> item)
        {
            throw new NotImplementedException();
        }

        public IEnumerator<KeyValuePair<int, string>> GetEnumerator()
        {
            for (int i = 0; i < _maxIndx; i++)
                yield return new KeyValuePair<int, string>(i, this[i]);
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            for (int i = 0; i < _maxIndx; i++)
                yield return this[i];
        }

        void IDictionary<int, string>.Add(int key, string value)
        {
            throw new NotImplementedException();
        }
    }
}
