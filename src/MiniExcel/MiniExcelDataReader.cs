namespace MiniExcelLibs
{
    using MiniExcelLibs.Utils;
    using System;
    using System.Collections.Generic;
    using System.Data;
    using System.IO;
    using System.Linq;

    public class MiniExcelDataReader : IDataReader
    {
        private readonly IEnumerator<IDictionary<string, object>> _source;
        private readonly int _fieldCount;
        private readonly List<string> _keys;
        private readonly Stream _stream;
        private bool _isFirst = true;

        internal MiniExcelDataReader(Stream stream, bool useHeaderRow = false, string sheetName = null, ExcelType excelType = ExcelType.UNKNOWN, string startCell = "A1", IConfiguration configuration = null)
        {
            _stream = stream;
            _source = MiniExcel.Query(_stream, useHeaderRow, sheetName, excelType, startCell, configuration).Cast<IDictionary<string, object>>().GetEnumerator();
            var isNext = _source.MoveNext();
            if (isNext)
            {
                _keys = _source.Current.Keys.ToList();
                _fieldCount = _keys.Count;
            }
        }

        public void Dispose()
        {
            _stream.Dispose();
        }

        public object GetValue(int i)
        {
            return _source.Current[_keys[i]];
        }

        public int FieldCount
        {
            get { return _fieldCount; }
        }

        public bool Read()
        {
            if (_isFirst)
            {
                _isFirst = false;
                return true;
            }
            return _source.MoveNext();
        }

        public string GetName(int i)
        {
            return _keys[i];
        }

        public int GetOrdinal(string name)
        {
            var i = _keys.IndexOf(name);
            return _keys.IndexOf(name);
        }

        public void Close()
        {
              return;
        }

        public int Depth => throw new NotImplementedException();

        public bool IsClosed => throw new NotImplementedException();

        public int RecordsAffected => throw new NotImplementedException();

        public object this[string name] => throw new NotImplementedException();

        public object this[int i] => throw new NotImplementedException();

        public DataTable GetSchemaTable()
        {
            throw new NotImplementedException();
        }

        public bool NextResult()
        {
            throw new NotImplementedException();
        }

        public bool GetBoolean(int i)
        {
            throw new NotImplementedException();
        }

        public byte GetByte(int i)
        {
            throw new NotImplementedException();
        }

        public long GetBytes(int i, long fieldOffset, byte[] buffer, int bufferoffset, int length)
        {
            throw new NotImplementedException();
        }

        public char GetChar(int i)
        {
            throw new NotImplementedException();
        }

        public long GetChars(int i, long fieldoffset, char[] buffer, int bufferoffset, int length)
        {
            throw new NotImplementedException();
        }

        public IDataReader GetData(int i)
        {
            throw new NotImplementedException();
        }

        public string GetDataTypeName(int i)
        {
            throw new NotImplementedException();
        }

        public DateTime GetDateTime(int i)
        {
            throw new NotImplementedException();
        }

        public decimal GetDecimal(int i)
        {
            throw new NotImplementedException();
        }

        public double GetDouble(int i)
        {
            throw new NotImplementedException();
        }

        public Type GetFieldType(int i)
        {
            throw new NotImplementedException();
        }

        public float GetFloat(int i)
        {
            throw new NotImplementedException();
        }

        public Guid GetGuid(int i)
        {
            throw new NotImplementedException();
        }

        public short GetInt16(int i)
        {
            throw new NotImplementedException();
        }

        public int GetInt32(int i)
        {
            throw new NotImplementedException();
        }

        public long GetInt64(int i)
        {
            throw new NotImplementedException();
        }

        public string GetString(int i)
        {
            throw new NotImplementedException();
        }

        public int GetValues(object[] values)
        {
            throw new NotImplementedException();
        }

        public bool IsDBNull(int i)
        {
            throw new NotImplementedException();
        }
    }
}
