namespace MiniExcelLib.Core.Attributes;

public class MiniExcelColumnWidthAttribute(double columnWidth) : MiniExcelAttributeBase
{
    public double ExcelColumnWidth { get; set; } = columnWidth;
}