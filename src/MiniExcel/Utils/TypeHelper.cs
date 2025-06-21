using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Globalization;
using System.Linq;
using System.Reflection;
using MiniExcelLibs.Exceptions;

namespace MiniExcelLibs.Utils;

internal static class TypeHelper
{
    public static IEnumerable<IDictionary<string, object>> ConvertToEnumerableDictionary(IDataReader reader)
    {
        while (reader.Read())
        {
            yield return Enumerable.Range(0, reader.FieldCount)
                .ToDictionary(
                    reader.GetName, 
                    reader.GetValue);
        }
    }

    /// <summary>
    /// From : https://stackoverflow.com/questions/906499/getting-type-t-from-ienumerablet
    /// </summary>
    public static IEnumerable<Type> GetGenericIEnumerables(object o)
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

    public static object? TypeMapping<T>(T v, ExcelColumnInfo pInfo, object itemValue, int rowIndex, string startCell, MiniExcelConfiguration config) where T : class, new()
    {
        try
        {
            return TypeMappingImpl(v, pInfo, itemValue, config);
        }
        catch (Exception ex) when (ex is InvalidCastException or FormatException)
        {
            var columnName = pInfo.ExcelColumnName ?? pInfo.Property.Name;
            var startRowIndex = ReferenceHelper.ConvertCellToXY(startCell).Item2;
            var errorRow = startRowIndex + rowIndex + 1;
                
            var msg = $"ColumnName: {columnName}, CellRow: {errorRow}, Value: {itemValue}. The value cannot be cast to type {pInfo.Property.Info.PropertyType.Name}.";
            throw new ExcelInvalidCastException(columnName, errorRow, itemValue, pInfo.Property.Info.PropertyType, msg);
        }
    }

    private static object? TypeMappingImpl<T>(T v, ExcelColumnInfo pInfo, object? itemValue, MiniExcelConfiguration config) where T : class, new()
    {
        object? newValue = null;
        if (pInfo.Nullable && string.IsNullOrWhiteSpace(itemValue?.ToString()))
        {
        }
        else if (pInfo.ExcludeNullableType == typeof(Guid))
        {
            newValue = Guid.Parse(itemValue?.ToString() ?? Guid.Empty.ToString());
        }
        else if (pInfo.ExcludeNullableType == typeof(DateTimeOffset))
        {
            var vs = itemValue?.ToString();
            if (pInfo.ExcelFormat is not null)
            {
                if (DateTimeOffset.TryParseExact(vs, pInfo.ExcelFormat, CultureInfo.InvariantCulture, DateTimeStyles.None, out var value))
                {
                    newValue = value;
                }
            }
            else if (DateTimeOffset.TryParse(vs, config.Culture, DateTimeStyles.None, out var value))
            {
                newValue = value;
            }
            else
            {
                throw new InvalidCastException($"{vs} cannot be cast to DateTime");
            }
        }
        else if (pInfo.ExcludeNullableType == typeof(DateTime))
        {
            // fix issue 257 https://github.com/mini-software/MiniExcel/issues/257
            if (itemValue is DateTime)
            {
                newValue = itemValue;
                pInfo.Property.SetValue(v, newValue);
                return newValue;
            }

            var vs = itemValue?.ToString();
            if (pInfo.ExcelFormat is not null)
            {
                if (pInfo.Property.Info.PropertyType == typeof(DateTimeOffset) && DateTimeOffset.TryParseExact(vs, pInfo.ExcelFormat, CultureInfo.InvariantCulture, DateTimeStyles.None, out var offsetValue))
                {
                    newValue = offsetValue;
                }
                else if (DateTime.TryParseExact(vs, pInfo.ExcelFormat, CultureInfo.InvariantCulture, DateTimeStyles.None, out var value))
                {
                    newValue = value;
                }
            }
            else if (DateTime.TryParse(vs, config.Culture, DateTimeStyles.None, out var dtValue))
                newValue = dtValue;
            else if (DateTime.TryParseExact(vs, "dd/MM/yyyy", CultureInfo.InvariantCulture, DateTimeStyles.None, out var dtExactValue))
                newValue = dtExactValue;
            else if (double.TryParse(vs, NumberStyles.None, CultureInfo.InvariantCulture, out var doubleValue))
                newValue = DateTime.FromOADate(doubleValue);
            else
                throw new InvalidCastException($"{vs} cannot be cast to DateTime");
        }
#if NET6_0_OR_GREATER
        else if (pInfo.ExcludeNullableType == typeof(DateOnly))
        {
            if (itemValue is DateOnly)
            {
                newValue = itemValue;
                pInfo.Property.SetValue(v, newValue);
                return newValue;
            }

            var vs = itemValue?.ToString();
            if (pInfo.ExcelFormat is not null)
            {
                if (DateOnly.TryParseExact(vs, pInfo.ExcelFormat, CultureInfo.InvariantCulture, DateTimeStyles.None, out var dateOnlyCustom))
                {
                    newValue = dateOnlyCustom;
                }
            }
            else if (DateOnly.TryParse(vs, config.Culture, DateTimeStyles.None, out var dateOnly))
                newValue = dateOnly;
            else if (DateOnly.TryParseExact(vs, "dd/MM/yyyy", CultureInfo.InvariantCulture, DateTimeStyles.None, out var dateOnlyExact))
                newValue = dateOnlyExact;
            else if (double.TryParse(vs, NumberStyles.None, CultureInfo.InvariantCulture, out var dateOnlyDouble))
                newValue = DateOnly.FromDateTime(DateTime.FromOADate(dateOnlyDouble));
            else
                throw new InvalidCastException($"{vs} cannot be cast to DateOnly");                
        }
#endif
        else if (pInfo.ExcludeNullableType == typeof(TimeSpan))
        {
            if (itemValue is TimeSpan)
            {
                newValue = itemValue;
                pInfo.Property.SetValue(v, newValue);
                return newValue;
            }

            var vs = itemValue?.ToString();
            if (pInfo.ExcelFormat is not null)
            {
                if (TimeSpan.TryParseExact(vs, pInfo.ExcelFormat, CultureInfo.InvariantCulture, out var value))
                {
                    newValue = value;
                }
            }
            else if (TimeSpan.TryParse(vs, config.Culture, out var tsValue))
                newValue = tsValue;
            else if (TimeSpan.TryParseExact(vs, @"hh\:mm\:ss\.fff", CultureInfo.InvariantCulture, out var tsExactValue))
                newValue = tsExactValue;
            else if (double.TryParse(vs, NumberStyles.None, CultureInfo.InvariantCulture, out var msValue))
                newValue = TimeSpan.FromMilliseconds(msValue);
            else
                throw new InvalidCastException($"{vs} cannot be cast to TimeSpan");
        }
        else if (pInfo.ExcludeNullableType == typeof(double)) // && (!Regex.IsMatch(itemValue.ToString(), @"^-?\d+(\.\d+)?([eE][-+]?\d+)?$") || itemValue.ToString().Trim().Equals("NaN")))
        {
            var invariantString = Convert.ToString(itemValue, CultureInfo.InvariantCulture);
            newValue = double.TryParse(invariantString, NumberStyles.Any, CultureInfo.InvariantCulture, out var value) ? value : double.NaN;
        }
        else if (pInfo.ExcludeNullableType == typeof(bool))
        {
            var vs = itemValue?.ToString();
            newValue = vs switch
            {
                "1" => true,
                "0" => false,
                _ => bool.TryParse(vs, out var parsed) ? parsed : null
            };
        }
        else if (pInfo.Property.Info.PropertyType == typeof(string))
        {
            newValue = XmlEncoder.DecodeString(itemValue?.ToString());
        }
        else if (pInfo.ExcludeNullableType.IsEnum)
        {
            var fieldInfo = pInfo.ExcludeNullableType.GetFields().FirstOrDefault(e => e.GetCustomAttribute<DescriptionAttribute>(false)?.Description == itemValue?.ToString());
            var value = fieldInfo?.Name ?? itemValue?.ToString() ?? "";
            newValue = Enum.Parse(pInfo.ExcludeNullableType, value, true);
        }
        else if (pInfo.ExcludeNullableType == typeof(Uri))
        {
            var rawValue = itemValue?.ToString();
            if (!Uri.TryCreate(rawValue, UriKind.RelativeOrAbsolute, out var uri))
                throw new InvalidCastException($"Value \"{rawValue}\" cannot be converted to Uri");
            newValue = uri;
        }
        else
        {
            // Use pInfo.ExcludeNullableType to resolve : https://github.com/mini-software/MiniExcel/issues/138
            newValue = Convert.ChangeType(itemValue, pInfo.ExcludeNullableType, config.Culture);
        }

        pInfo.Property.SetValue(v, newValue);
        return newValue;
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