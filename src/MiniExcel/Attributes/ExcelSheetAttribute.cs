using System;
using MiniExcelLibs.OpenXml;

namespace MiniExcelLibs.Attributes;

[AttributeUsage(AttributeTargets.Class)]
public class ExcelSheetAttribute : Attribute
{
    public string? Name { get; set; }
    public SheetState State { get; set; } = SheetState.Visible;
}

public class DynamicExcelSheet(string key) : ExcelSheetAttribute
{
    public string Key { get; set; } = key;
}