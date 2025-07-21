namespace MiniExcelLib.Core.Attributes;

[AttributeUsage(AttributeTargets.Property | AttributeTargets.Field)]
public class MiniExcelColumnNameAttribute(string excelColumnName, string[]? aliases = null) : Attribute
{
    public string ExcelColumnName { get; set; } = excelColumnName;
    public string[] Aliases { get; set; } = aliases ?? [];
}