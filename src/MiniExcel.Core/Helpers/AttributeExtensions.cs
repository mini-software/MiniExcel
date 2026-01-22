namespace MiniExcelLib.Core.Helpers;

public static class AttributeExtensions
{
    private static TValue? GetValueOrDefault<TAttribute, TValue>(
        Func<TAttribute, TValue> selector, 
        TAttribute? attr) where TAttribute : Attribute
    {
        return attr is not null ? selector(attr) : default;
    }

    public static TAttribute? GetAttribute<TAttribute>(this MemberInfo prop, bool isInherit = true) 
        where TAttribute : Attribute
    {
        return prop.GetAttributeValue((TAttribute attr) => attr, isInherit);
    }

    public static TValue? GetAttributeValue<TAttribute, TValue>(
        this MemberInfo prop,
        Func<TAttribute, TValue> selector,
        bool isInherit = true ) where TAttribute : Attribute
    {
        var attr = Attribute.GetCustomAttribute(prop, typeof(TAttribute), isInherit) as TAttribute;
        return GetValueOrDefault(selector, attr);
    }
}
