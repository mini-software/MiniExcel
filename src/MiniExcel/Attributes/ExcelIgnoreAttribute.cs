using System;

namespace MiniExcelLibs.Attributes;

[AttributeUsage(AttributeTargets.Property | AttributeTargets.Field)]
public class ExcelIgnoreAttribute(bool excelIgnore = true) : Attribute
{
    public bool ExcelIgnore { get; set; } = excelIgnore;
}