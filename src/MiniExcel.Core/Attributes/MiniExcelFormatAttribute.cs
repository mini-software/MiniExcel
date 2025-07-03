namespace MiniExcelLib.Core.Attributes;

[AttributeUsage(AttributeTargets.Property | AttributeTargets.Field)]
public class MiniExcelFormatAttribute(string format) : Attribute
{
    public string Format { get; set; } = format;
}