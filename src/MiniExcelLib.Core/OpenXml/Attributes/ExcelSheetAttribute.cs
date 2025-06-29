using MiniExcelLib.Core.OpenXml.Models;

namespace MiniExcelLib.Core.OpenXml.Attributes;

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