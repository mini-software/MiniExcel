namespace MiniExcelLib.Core.Attributes;

[AttributeUsage(AttributeTargets.Property | AttributeTargets.Field)]
public class MiniExcelColumnNameAttribute(string columnName, string[]? aliases = null) : Attribute
{
    public string ExcelColumnName { get; set; } = columnName;
    public string[] Aliases { get; set; } = aliases ?? [];
}