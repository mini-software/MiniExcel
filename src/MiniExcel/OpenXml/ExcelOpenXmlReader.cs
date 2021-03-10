namespace MiniExcelLibs.OpenXml
{
    using System;
    using System.Collections.Generic;
    using System.Data;
    using System.IO;
    using System.Linq;

    internal class ExcelOpenXmlReader : IDataReader
    {
        private static Worksheet _Sheet ;
        private Dictionary<int, Dictionary<int, object>> _Rows { get { return _Sheet.Rows; } }

        public ExcelOpenXmlReader(Stream stream)
        {
            _Sheet = MiniExcel.GetFirstSheet(stream);
        }

        public int RowCount { get { return _Sheet.RowCount; } }
        public int FieldCount { get { return _Sheet.FieldCount; } }
        public int Depth { get; private set; }
        public int CurrentRowIndex { get { return Depth - 1; } }

        public object this[int i] => GetValue(i);
        public object this[string name] => GetValue(GetOrdinal(name));

        public bool Read()
        {
            if (Depth == RowCount)
                return false;
            Depth++;
            return true;
        }

        public string GetName(int i) => ExcelOpenXmlUtils.ConvertColumnName(i + 1);


        public int GetOrdinal(string name) => ExcelOpenXmlUtils.GetCellColumnIndex(name);

        public object GetValue(int i)
        {
            //if (CurrentRowIndex < 0)
            //	throw new InvalidOperationException("Invalid attempt to read when no data is present.");
            if (!_Rows.Keys.Contains(CurrentRowIndex))
                return null;
            if (_Rows[this.CurrentRowIndex].TryGetValue(i, out var v))
                return v;
            return null;
        }

        public int GetValues(object[] values)
        {
            return this.Depth;
        }

        //TODO: multiple sheets
        public bool NextResult() => false;

        public void Dispose() { }

        public void Close() { }

        public int RecordsAffected => throw new NotImplementedException();

        bool IDataReader.IsClosed => this.RowCount - 1 == this.Depth;

        public string GetString(int i) => (string)GetValue(i);

        public bool GetBoolean(int i) => (bool)GetValue(i);

        public byte GetByte(int i) => (byte)GetValue(i);

        public long GetBytes(int i, long fieldOffset, byte[] buffer, int bufferoffset, int length) => throw new NotImplementedException();

        public char GetChar(int i) => (char)GetValue(i);

        public long GetChars(int i, long fieldoffset, char[] buffer, int bufferoffset, int length) => throw new NotImplementedException();

        public IDataReader GetData(int i) => throw new NotImplementedException();

        public string GetDataTypeName(int i) => throw new NotImplementedException();

        public DateTime GetDateTime(int i) => (DateTime)GetValue(i);

        public decimal GetDecimal(int i) => (decimal)GetValue(i);

        public double GetDouble(int i) => (double)GetValue(i);

        public Type GetFieldType(int i)
        {
            var v = GetValue(i);
            return v == null ? typeof(string) : v.GetType();
        }

        public float GetFloat(int i) => (float)GetValue(i);

        public Guid GetGuid(int i) => (Guid)GetValue(i);

        public short GetInt16(int i) => (short)GetValue(i);

        public int GetInt32(int i) => (int)GetValue(i);

        public long GetInt64(int i) => (long)GetValue(i);

        public DataTable GetSchemaTable()
        {
            var dataTable = new DataTable("SchemaTable");
            dataTable.Locale = System.Globalization.CultureInfo.InvariantCulture;
            dataTable.Columns.Add("ColumnName", typeof(string));
            dataTable.Columns.Add("ColumnOrdinal", typeof(int));
            for (int i = 0; i < this.FieldCount; i++)
            {
                dataTable.Rows.Add(this.GetName(i), i);
            }
            DataColumnCollection columns = dataTable.Columns;
            foreach (DataColumn item in columns)
            {
                item.ReadOnly = true;
            }
            return dataTable;
        }

        public bool IsDBNull(int i) => GetValue(i) == null;
    }
}
