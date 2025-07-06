using System.Linq.Expressions;

namespace MiniExcelLib.Reflection;

public class MemberGetter(PropertyInfo property)
{
    private readonly Func<object, object?> _mGetFunc = CreateGetterDelegate(property);

    public object? Invoke(object instance)
    {
        return _mGetFunc.Invoke(instance);
    }

    private static Func<object, object?> CreateGetterDelegate(PropertyInfo property)
    {
        var paramInstance = Expression.Parameter(typeof(object));
        var bodyInstance = Expression.Convert(paramInstance, property.DeclaringType!);
        var bodyProperty = Expression.Property(bodyInstance, property);
        var bodyReturn = Expression.Convert(bodyProperty, typeof(object));

        return Expression.Lambda<Func<object, object?>>(bodyReturn, paramInstance).Compile();
    }
}