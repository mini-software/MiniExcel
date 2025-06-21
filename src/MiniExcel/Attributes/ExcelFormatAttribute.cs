using System;

namespace MiniExcelLibs.Attributes;

[AttributeUsage(AttributeTargets.Property | AttributeTargets.Field)]
public class ExcelFormatAttribute(string format) : Attribute
{
    public string Format { get; set; } = format;
}