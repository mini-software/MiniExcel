using System;

namespace MiniExcelLibs.Attributes;

[AttributeUsage(AttributeTargets.Property | AttributeTargets.Field)]
public class ExcelColumnNameAttribute(string excelColumnName, string[]? aliases = null) : Attribute
{
    public string ExcelColumnName { get; set; } = excelColumnName;
    public string[] Aliases { get; set; } = aliases ?? [];
}