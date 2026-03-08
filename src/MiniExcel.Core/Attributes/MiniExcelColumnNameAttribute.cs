namespace MiniExcelLib.Core.Attributes;

public class MiniExcelColumnNameAttribute(string columnName, string[]? aliases = null) : MiniExcelAttributeBase
{
    public string ExcelColumnName { get; set; } = columnName;
    public string[] Aliases { get; set; } = aliases ?? [];
}