namespace MiniExcelLib.Core.Exceptions;

public class ValueNotAssignableException(string columnName, int row, object value, Type columnType, string message)
    : InvalidCastException(message)
{
    public string ColumnName { get; set; } = columnName;
    public int Row { get; set; } = row;
    public object Value { get; set; } = value;
    public Type ColumnType { get; set; } = columnType;
}