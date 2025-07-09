namespace MiniExcelLib.Attributes;

[AttributeUsage(AttributeTargets.Property | AttributeTargets.Field)]
public class MiniExcelColumnWidthAttribute(double excelColumnWidth) : Attribute
{
    public double ExcelColumnWidth { get; set; } = excelColumnWidth;
}