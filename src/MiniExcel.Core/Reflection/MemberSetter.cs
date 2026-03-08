using System.Linq.Expressions;

namespace MiniExcelLib.Core.Reflection;

public class MemberSetter
{
    private readonly Action<object, object?> _setFunc;

    public MemberSetter(MemberInfo property)
    {
        if (property is null)
            throw new ArgumentNullException(nameof(property));
        
        _setFunc = CreateSetterDelegate(property);
    }

    public void Invoke(object instance, object? value)
    {
        _setFunc.Invoke(instance, value);
    }

    private static Action<object, object?> CreateSetterDelegate(MemberInfo member)
    {
        var paramInstance = Expression.Parameter(typeof(object));
        var paramValue = Expression.Parameter(typeof(object));

        var bodyInstance = Expression.Convert(paramInstance, member.DeclaringType!);
        
        var memberAccess = Expression.MakeMemberAccess(bodyInstance, member);
        var bodyValue = Expression.Convert(paramValue, memberAccess.Type);
        var assignExp = Expression.Assign(memberAccess, bodyValue);

        return Expression.Lambda<Action<object, object?>>(assignExp, paramInstance, paramValue).Compile();
    }
}
