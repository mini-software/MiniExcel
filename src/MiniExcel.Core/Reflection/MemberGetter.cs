using System.Linq.Expressions;

namespace MiniExcelLib.Core.Reflection;

public class MemberGetter(MemberInfo member)
{
    private readonly Func<object, object?> _mGetFunc = CreateGetterDelegate(member);

    public object? Invoke(object instance)
        => _mGetFunc.Invoke(instance);

    private static Func<object, object?> CreateGetterDelegate(MemberInfo member)
    {
        var paramInstance = Expression.Parameter(typeof(object));
        var bodyInstance = Expression.Convert(paramInstance, member.DeclaringType!);
        
        var bodyProperty = Expression.MakeMemberAccess(bodyInstance, member);
        var bodyReturn = Expression.Convert(bodyProperty, typeof(object));

        return Expression.Lambda<Func<object, object?>>(bodyReturn, paramInstance).Compile();
    }
}
