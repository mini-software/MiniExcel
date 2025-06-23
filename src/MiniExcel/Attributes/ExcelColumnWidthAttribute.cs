using System;

namespace MiniExcelLibs.Attributes;

[AttributeUsage(AttributeTargets.Property | AttributeTargets.Field)]
public class ExcelColumnWidthAttribute(double excelColumnWidth) : Attribute
{
    public double ExcelColumnWidth { get; set; } = excelColumnWidth;
}