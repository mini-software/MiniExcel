namespace MiniExcelLib.Core.Attributes;

[AttributeUsage(AttributeTargets.Property | AttributeTargets.Field)]
public class MiniExcelColumnWidthAttribute(double columnWidth) : Attribute
{
    public double ExcelColumnWidth { get; set; } = columnWidth;
}