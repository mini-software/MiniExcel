namespace MiniExcelLib.OpenXml.Templates;

public class OpenXmlValueExtractor : IInputValueExtractor
{
    public IDictionary<string, object?> ToValueDictionary(object valueObject)
        => valueObject is Dictionary<string, object> valueDictionary
            ? GetValuesFromDictionary(valueDictionary)
            : GetValuesFromObject(valueObject);

    private static Dictionary<string, object?> GetValuesFromDictionary(Dictionary<string, object> valueDictionary)
    {
        return valueDictionary.ToDictionary(
            x => x.Key,
            x => x.Value is IDataReader dataReader
                ? dataReader.ToEnumerableDictionaries().ToList()
                : x.Value)!;
    }

    private static Dictionary<string, object?> GetValuesFromObject(object valueObject)
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