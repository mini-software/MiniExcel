namespace MiniExcelLib.Core.Attributes;

public class MiniExcelIgnoreAttribute(bool ignore = true) : MiniExcelAttributeBase
{
    public bool Ignore { get; set; } = ignore;
}