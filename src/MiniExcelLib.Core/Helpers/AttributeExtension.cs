using System.Reflection;

namespace MiniExcelLib.Core.Helpers;

internal static class AttributeExtension
{
    internal static TValue? GetAttributeValue<TAttribute, TValue>(
        this Type targetType,
        Func<TAttribute?, TValue> selector) where TAttribute : Attribute
    {
        var attributeType = targetType
            .GetCustomAttributes(typeof(TAttribute), true)
            .FirstOrDefault() as TAttribute;
            
        return GetValueOrDefault(selector, attributeType);
    }

    private static TValue? GetValueOrDefault<TAttribute, TValue>(
        Func<TAttribute, TValue> selector, 
        TAttribute? attr) where TAttribute : Attribute
    {
        return attr is not null ? selector(attr) : default;
    }

    internal static TAttribute? GetAttribute<TAttribute>(
        this MemberInfo prop,
        bool isInherit = true) where TAttribute : Attribute
    {
        return GetAttributeValue(prop, (TAttribute attr) => attr, isInherit);
    }

    internal static TValue? GetAttributeValue<TAttribute, TValue>(
        this MemberInfo prop,
        Func<TAttribute, TValue> selector,
        bool isInherit = true ) where TAttribute : Attribute
    {
        var attr = Attribute.GetCustomAttribute(prop, typeof(TAttribute), isInherit) as TAttribute;
        return GetValueOrDefault(selector, attr);
    }
}