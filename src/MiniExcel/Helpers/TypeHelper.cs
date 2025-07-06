namespace MiniExcelLib.Helpers;

internal static class TypeHelper
{
    public static IEnumerable<IDictionary<string, object>> ConvertToEnumerableDictionary(IDataReader reader)
    {
        while (reader.Read())
        {
            yield return Enumerable
                .Range(0, reader.FieldCount)
                .ToDictionary(reader.GetName, reader.GetValue);
        }
    }

    /// <summary>
    /// From : https://stackoverflow.com/questions/906499/getting-type-t-from-ienumerablet
    /// </summary>
    public static IEnumerable<Type?> GetGenericIEnumerables(object o)
    {
        return o.GetType()
            .GetInterfaces()
            .Where(t => t.IsGenericType && t.GetGenericTypeDefinition() == typeof(IEnumerable<>))
            .Select(t => t.GetGenericArguments()[0]);
    }

    public static bool IsNumericType(Type type, bool isNullableUnderlyingType = false)
    {
        if (isNullableUnderlyingType)
            type = Nullable.GetUnderlyingType(type) ?? type;

        // True for all numeric types except bool, sbyte and byte
        return Type.GetTypeCode(type) is >= TypeCode.Int16 and <= TypeCode.Decimal;
    }

    public static bool IsAsyncEnumerable(this Type type, out Type? genericArgument)
    {
        var asyncEnumrableInterfaceType = type
            .GetInterfaces()
            .FirstOrDefault(i => i.IsGenericType && i.GetGenericTypeDefinition() == typeof(IAsyncEnumerable<>));
        
        genericArgument = asyncEnumrableInterfaceType?.GetGenericArguments().FirstOrDefault();
        return genericArgument is not null;
    }
}