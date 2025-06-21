using System;

namespace MiniExcelLibs.Exceptions;

public class ExcelInvalidCastException(string columnName, int row, object value, Type invalidCastType, string message)
    : InvalidCastException(message)
{
    public string ColumnName { get; set; } = columnName;
    public int Row { get; set; } = row;
    public object Value { get; set; } = value;
    public Type InvalidCastType { get; set; } = invalidCastType;
}