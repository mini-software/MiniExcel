using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Reflection;
using MiniExcelLibs.Utils;

namespace MiniExcelLibs.OpenXml.SaveByTemplate
{
    public class InputValueExtractor : IInputValueExtractor
    {
        public IDictionary<string, object> ToValueDictionary(object valueObject)
            => valueObject is Dictionary<string, object> valueDictionary
                ? GetValuesFromDictionary(valueDictionary)
                : GetValuesFromObject(valueObject);

        private static IDictionary<string, object> GetValuesFromDictionary(Dictionary<string, object> valueDictionary)
        {
            return valueDictionary
                .ToDictionary(
                    x => x.Key,
                    x => x.Value is IDataReader dataReader
                        ? TypeHelper.ConvertToEnumerableDictionary(dataReader).ToList()
                        : x.Value);
        }

        private static IDictionary<string, object> GetValuesFromObject(object valueObject)
        {
            var type = valueObject.GetType();

            var propertyValues = type
                .GetProperties(BindingFlags.Public | BindingFlags.Instance)
                .Select(property => new { property.Name, Value = property.GetValue(valueObject) });

            var fieldValues = type
                .GetFields(BindingFlags.Public | BindingFlags.Instance)
                .Select(field => new { field.Name, Value = field.GetValue(valueObject) });

            return propertyValues
                .Concat(fieldValues)
                .ToDictionary(x => x.Name, x => x.Value);
        }
    }
}