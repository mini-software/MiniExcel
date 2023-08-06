
using System;
using System.Collections.Concurrent;
using System.Linq;
using System.Reflection;

namespace MiniExcelLibs
{
    public abstract class Member
    {
    }

    public class Property: Member
    {
        private static readonly ConcurrentDictionary<Type, Property[]> m_cached = new ConcurrentDictionary<Type, Property[]>();

        private readonly MemberGetter m_geter;

        private readonly MemberSetter m_seter;

        public Property(PropertyInfo property)
        {
            Name = property.Name;
            Info = property;

            if (property.CanRead == true)
            {
                CanRead = true;
                m_geter = new MemberGetter(property);
            }
            if (property.CanWrite == true)
            {
                CanWrite = true;
                m_seter = new MemberSetter(property);
            }
        }

        public bool CanRead { get; private set; }

        public bool CanWrite { get; private set; }
        public PropertyInfo Info { get; private set; }

        public string Name { get; protected set; }

        public static Property[] GetProperties(Type type)
        {
            return m_cached.GetOrAdd(type, t => t.GetProperties().Select(p => new Property(p)).ToArray());
        }

        public object GetValue(object instance)
        {
            if (m_geter == null)
            {
                throw new NotSupportedException();
            }
            return m_geter.Invoke(instance);
        }

        public void SetValue(object instance, object value)
        {
            if (m_seter == null)
            {
                throw new NotSupportedException($"{Name} can't set value");
            }
            m_seter.Invoke(instance, value);
        }
    }
}