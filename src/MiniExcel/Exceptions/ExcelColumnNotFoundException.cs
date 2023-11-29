using System.Collections.Generic;

namespace MiniExcelLibs.Exceptions
{
    public class ExcelColumnNotFoundException : KeyNotFoundException
    {
        public string ColumnName { get; set; }
        public string[] ColumnAliases { get; }
        public string ColumnIndex { get; set; }
        public int RowIndex { get; set; }
        public IDictionary<string, int> Headers { get; }
        public object RowValues { get; set; }

        public ExcelColumnNotFoundException(string columnIndex, string columnName, string[] columnAliases, int rowIndex, IDictionary<string, int> headers, object value, string message) : base(message)
        {
            ColumnIndex = columnIndex;
            ColumnName = columnName;
            ColumnAliases = columnAliases;
            RowIndex = rowIndex;
            Headers = headers;
            RowValues = value;
        }
    }
}
