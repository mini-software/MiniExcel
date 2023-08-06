using System;
using System.Linq.Expressions;
using System.Reflection;

namespace MiniExcelLibs
{
    public class MemberSetter
    {
        private readonly Action<object, object> setFunc;

        public MemberSetter(PropertyInfo property)
        {
            if (property == null)
            {
                throw new ArgumentNullException(nameof(property));
            }
            setFunc = CreateSetterDelegate(property);
        }

        public void Invoke(object instance, object value)
        {
            setFunc.Invoke(instance, value);
        }

        private static Action<object, object> CreateSetterDelegate(PropertyInfo property)
        {
            var param_instance = Expression.Parameter(typeof(object));
            var param_value = Expression.Parameter(typeof(object));

            var body_instance = Expression.Convert(param_instance, property.DeclaringType);
            var body_value = Expression.Convert(param_value, property.PropertyType);
            var body_call = Expression.Call(body_instance, property.GetSetMethod(true), body_value);

            return Expression.Lambda<Action<object, object>>(body_call, param_instance, param_value).Compile();
        }
    }
}