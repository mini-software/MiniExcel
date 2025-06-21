using System.Collections.Generic;

namespace MiniExcelLibs.Exceptions;

public class ExcelColumnNotFoundException(
    string? columnIndex,
    string? columnName,
    string[] columnAliases,
    int rowIndex,
    IDictionary<string, int> headers,
    object value,
    string message) : KeyNotFoundException(message)
{
    public string? ColumnName { get; set; } = columnName;
    public string? ColumnIndex { get; set; } = columnIndex;
    public string[] ColumnAliases { get; } = columnAliases;
    public int RowIndex { get; set; } = rowIndex;
    public IDictionary<string, int> Headers { get; } = headers;
    public object RowValues { get; set; } = value;
}