using System.Linq.Expressions;
using System.Reflection;

namespace MiniExcelLib.Core.Reflection;

public class MemberSetter
{
    private readonly Action<object, object?> _setFunc;

    public MemberSetter(PropertyInfo property)
    {
        if (property is null)
            throw new ArgumentNullException(nameof(property));
        
        _setFunc = CreateSetterDelegate(property);
    }

    public void Invoke(object instance, object? value)
    {
        _setFunc.Invoke(instance, value);
    }

    private static Action<object, object?> CreateSetterDelegate(PropertyInfo property)
    {
        var paramInstance = Expression.Parameter(typeof(object));
        var paramValue = Expression.Parameter(typeof(object));

        var bodyInstance = Expression.Convert(paramInstance, property.DeclaringType!);
        var bodyValue = Expression.Convert(paramValue, property.PropertyType);
        var bodyCall = Expression.Call(bodyInstance, property.GetSetMethod(true)!, bodyValue);

        return Expression.Lambda<Action<object, object?>>(bodyCall, paramInstance, paramValue).Compile();
    }
}