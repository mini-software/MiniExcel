
using System;
using System.Collections.Concurrent;
using System.Linq;
using System.Reflection;

namespace MiniExcelLibs
{
    /// <summary>
    /// 表示属性
    /// </summary>
    public class Property: Member
    {
        /// <summary>
        /// 类型属性的Setter缓存
        /// </summary>
        private static readonly ConcurrentDictionary<Type, Property[]> m_cached = new ConcurrentDictionary<Type, Property[]>();

        /// <summary>
        /// 获取器
        /// </summary>
        private readonly MemberGetter m_geter;

        /// <summary>
        /// 设置器
        /// </summary>
        private readonly MemberSetter m_seter;

        /// <summary>
        /// 属性
        /// </summary>
        /// <param name="property">属性信息</param>
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

        /// <summary>
        /// 是否可以读取
        /// </summary>
        public bool CanRead { get; private set; }

        /// <summary>
        /// 是否可以写入
        /// </summary>
        public bool CanWrite { get; private set; }

        /// <summary>
        /// 获取属性信息
        /// </summary>
        public PropertyInfo Info { get; private set; }

        /// <summary>
        /// 获取属性名称
        /// </summary>
        public string Name { get; protected set; }

        /// <summary>
        /// 从类型的属性获取属性
        /// </summary>
        /// <param name="type">类型</param>
        /// <returns></returns>
        public static Property[] GetProperties(Type type)
        {
            return m_cached.GetOrAdd(type, t => t.GetProperties().Select(p => new Property(p)).ToArray());
        }

        /// <summary>
        /// 获取属性的值
        /// </summary>
        /// <param name="instance">实例</param>
        /// <exception cref="NotSupportedException"></exception>
        /// <returns></returns>
        public object GetValue(object instance)
        {
            if (m_geter == null)
            {
                throw new NotSupportedException();
            }
            return m_geter.Invoke(instance);
        }

        /// <summary>
        /// 设置属性的值
        /// </summary>
        /// <param name="instance">实例</param>
        /// <param name="value">值</param>
        /// <exception cref="NotSupportedException"></exception>
        public void SetValue(object instance, object value)
        {
            if (m_seter == null)
            {
                throw new NotSupportedException($"{Name}不允许赋值");
            }
            m_seter.Invoke(instance, value);
        }
    }
}