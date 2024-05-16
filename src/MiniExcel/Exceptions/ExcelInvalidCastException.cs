using System;

namespace MiniExcelLibs.Exceptions
{
    public class ExcelInvalidCastException : InvalidCastException
    {
        public string ColumnName { get; set; }
        public int Row { get; set; }
        public object Value { get; set; }
        public Type InvalidCastType { get; set; }
        public ExcelInvalidCastException(string columnName, int row, object value, Type invalidCastType, string message) : base(message)
        {
            ColumnName = columnName;
            Row = row;
            Value = value;
            InvalidCastType = invalidCastType;
        }
    }
}
