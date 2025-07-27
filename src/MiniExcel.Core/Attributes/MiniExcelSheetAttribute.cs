using MiniExcelLib.Core.OpenXml.Models;

namespace MiniExcelLib.Core.Attributes;

[AttributeUsage(AttributeTargets.Class)]
public class MiniExcelSheetAttribute : Attribute
{
    public string? Name { get; set; }
    public SheetState State { get; set; } = SheetState.Visible;
}

public class DynamicExcelSheetAttribute(string key) : MiniExcelSheetAttribute
{
    public string Key { get; set; } = key;
}