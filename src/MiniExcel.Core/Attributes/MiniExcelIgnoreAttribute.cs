namespace MiniExcelLib.Core.Attributes;

[AttributeUsage(AttributeTargets.Property | AttributeTargets.Field)]
public class MiniExcelIgnoreAttribute(bool excelIgnore = true) : Attribute
{
    public bool ExcelIgnore { get; set; } = excelIgnore;
}