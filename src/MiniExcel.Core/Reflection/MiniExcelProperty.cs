using MiniExcelLib.Core.Exceptions;

namespace MiniExcelLib.Core.Reflection;

public abstract class Member;

public class MiniExcelProperty : Member
{
    private readonly MemberGetter? _getter;
    private readonly MemberSetter? _setter;

    public MiniExcelProperty(PropertyInfo property)
    {
        Name = property.Name;
        Info = property;

        if (property.GetIndexParameters().Length != 0)
        {
            const string msg = "Types containing indexers cannot be serialized. Please remove them or decorate them with MiniExcelIgnoreAttribute.";
            throw new MiniExcelNotSerializableException(msg, property);
        }

        if (property.CanRead)
        {
            CanRead = true;
            _getter = new MemberGetter(property);
        }
        
        if (property.CanWrite)
        {
            CanWrite = true;
            _setter = new MemberSetter(property);
        }
    }

    public string Name { get; protected set; }
    public bool CanRead { get; private set; }
    public bool CanWrite { get; private set; }
    public PropertyInfo Info { get; private set; }

    public object? GetValue(object instance) => _getter is not null 
        ? _getter.Invoke(instance) 
        : throw new NotSupportedException();

    public void SetValue(object instance, object? value)
    {
        if (_setter is null)
            throw new NotSupportedException($"{Name} can't set value");
            
        _setter.Invoke(instance, value);
    }
}