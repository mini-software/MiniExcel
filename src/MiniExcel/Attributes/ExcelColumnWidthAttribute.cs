namespace MiniExcelLibs.Attributes;

[AttributeUsage(AttributeTargets.Property | AttributeTargets.Field, AllowMultiple = false)]
public class ExcelColumnWidthAttribute : Attribute
{
    public double ExcelColumnWidth { get; set; }
    public ExcelColumnWidthAttribute(double excelColumnWidth) => ExcelColumnWidth = excelColumnWidth;
}