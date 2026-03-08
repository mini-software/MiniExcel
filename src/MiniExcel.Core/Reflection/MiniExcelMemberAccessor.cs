using MiniExcelLib.Core.Exceptions;

namespace MiniExcelLib.Core.Reflection;

public class MiniExcelMemberAccessor
{
    private readonly MemberGetter? _getter;
    private readonly MemberSetter? _setter;

    public MiniExcelMemberAccessor(MemberInfo member)
    {
        Name = member.Name;
        MemberInfo = member;
        
        var type = member switch
        {
            PropertyInfo p => p.PropertyType,
            FieldInfo f => f.FieldType,
            _ => throw new InvalidMappingException("Only properties and fields can be mapped", member.DeclaringType!, member) // unreachable exception
        };

        var nullableType = Nullable.GetUnderlyingType(type);
        IsNullable = nullableType is not null;
        Type = nullableType ?? type;

        if (member is PropertyInfo property && property.GetIndexParameters().Length != 0)
        {
            const string msg = "Types containing indexers cannot be serialized. Please remove them or decorate them with MiniExcelIgnoreAttribute.";
            throw new MemberNotSerializableException(msg, member);
        }

        if (member is FieldInfo or PropertyInfo { CanRead: true })
        {
            CanRead = true;
            _getter = new MemberGetter(member);
        }

        if (member is FieldInfo { IsInitOnly: false } || (member is PropertyInfo prop && prop.GetSetMethod() is not null))
        {
            CanWrite = true;
            _setter = new MemberSetter(member);
        }
    }

    public string Name { get; private set; }
    public Type Type { get; private set; }
    public MemberInfo MemberInfo { get; private set; }
    public bool IsNullable { get; private set; }
    public bool CanRead { get; private set; }
    public bool CanWrite { get; private set; }

    public object? GetValue(object instance) => _getter is not null
        ? _getter.Invoke(instance)
        : throw new InvalidOperationException($"The value of member \"{Name}\" cannot be retrieved");

    public void SetValue(object instance, object? value)
    {
        if (_setter is null)
            throw new InvalidOperationException($"The value of member \"{Name}\" cannot be set");
            
        _setter.Invoke(instance, value);
    }
}
