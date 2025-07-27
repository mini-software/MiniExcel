namespace MiniExcelLib.Core.Attributes;

[AttributeUsage(AttributeTargets.Property | AttributeTargets.Field)]
public class MiniExcelIgnoreAttribute(bool ignore = true) : Attribute
{
    public bool Ignore { get; set; } = ignore;
}