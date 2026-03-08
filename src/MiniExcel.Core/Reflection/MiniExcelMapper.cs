using System.ComponentModel;
using MiniExcelLib.Core.Enums;
using MiniExcelLib.Core.Exceptions;

namespace MiniExcelLib.Core.Reflection;

public static partial class MiniExcelMapper
{
    [CreateSyncVersion]
    public static async IAsyncEnumerable<T> MapQueryAsync<T>(IAsyncEnumerable<IDictionary<string, object?>> values, int rowOffset, bool mapHeaderAsData, bool trimColumnNames, MiniExcelBaseConfiguration configuration, Func<string?, string?>? stringDecoderFunc = null, [EnumeratorCancellation] CancellationToken cancellationToken = default) where T : class, new()
    {
        cancellationToken.ThrowIfCancellationRequested();

        var type = typeof(T);

        //TODO:need to optimize
        List<MiniExcelColumnMapping> mappings = [];
        Dictionary<string, int> headersDic = [];
        string[] keys = [];
        var first = true;
        var rowIndex = 0;

        await foreach (var item in values.WithCancellation(cancellationToken).ConfigureAwait(false))
        {
            if (first)
            {
                keys = item.Keys.ToArray();
                headersDic = ColumnMappingsProvider.GetHeaders(item, trimColumnNames);
                mappings = ColumnMappingsProvider.GetMappingsForImport(type, keys, configuration);
                first = false;

                // if we treat the header as data we move forwards with the mapping otherwise we jump to the next iteration
                if (!mapHeaderAsData) 
                    continue;
            }

            var v = new T();
            foreach (var map in mappings)
            {
                if (map.ExcelColumnAliases is not null)
                {
                    foreach (var alias in map.ExcelColumnAliases)
                    {
                        if (headersDic?.TryGetValue(alias, out var columnId) is true)
                        {
                            var columnName = keys[columnId];
                            item.TryGetValue(columnName, out var aliasItemValue);

                            if (aliasItemValue is not null)
                            {
                                var newAliasValue = MapValue(v, map, aliasItemValue, rowIndex + rowOffset, configuration, stringDecoderFunc);
                            }
                        }
                    }
                }

                //Q: Why need to check every time? A: it needs to check everytime, because it's dictionary
                object? itemValue = null;
                if (map.ExcelIndexName is not null && (keys?.Contains(map.ExcelIndexName) is true))
                {
                    item.TryGetValue(map.ExcelIndexName, out itemValue);
                }
                else if (map.ExcelColumnName is not null && (headersDic?.TryGetValue(map.ExcelColumnName, out var columnId) is true))
                {
                    var columnName = keys[columnId];
                    item.TryGetValue(columnName, out itemValue);
                }

                if (itemValue is not null)
                {
                    var newValue = MapValue(v, map, itemValue, rowIndex + rowOffset, configuration, stringDecoderFunc);
                }
            }
            
            rowIndex++;
            yield return v;
        }
    }
    
    public static object? MapValue<T>(T v, MiniExcelColumnMapping map, object itemValue, int rowIndex, MiniExcelBaseConfiguration config, Func<string?, string?>? stringDecoderFunc = null) where T : class, new()
    {
        try
        {
            object? newValue = null;
            
            if (map.Nullable && string.IsNullOrWhiteSpace(itemValue?.ToString()))
            {
                // value is null, no transformation required
            }
            
            else if (map.ExcludeNullableType == typeof(Guid))
            {
                newValue = itemValue switch
                {
                    Guid g => g, 
                    string str => Guid.Parse(str),
                    _ => Guid.Empty.ToString()
                };
            }
            
            else if (map.ExcludeNullableType == typeof(DateTimeOffset))
            {
                var vs = itemValue?.ToString();
                if (map.ExcelFormat is not null)
                {
                    if (DateTimeOffset.TryParseExact(vs, map.ExcelFormat, CultureInfo.InvariantCulture, DateTimeStyles.None, out var value))
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
            
            else if (map.ExcludeNullableType == typeof(DateTime))
            {
                // fix issue 257 https://github.com/mini-software/MiniExcel/issues/257
                if (itemValue is DateTime)
                {
                    newValue = itemValue;
                    map.MemberAccessor.SetValue(v, newValue);
                    return newValue;
                }

                var vs = itemValue?.ToString();
                if (map.ExcelFormat is not null)
                {
                    if (DateTime.TryParseExact(vs, map.ExcelFormat, CultureInfo.InvariantCulture, DateTimeStyles.None, out var value))
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
            else if (map.ExcludeNullableType == typeof(DateOnly))
            {
                if (itemValue is DateOnly)
                {
                    newValue = itemValue;
                    map.MemberAccessor.SetValue(v, newValue);
                    return newValue;
                }

                if (itemValue is DateTime dateTimeValue && config.DateOnlyConversionMode is not DateOnlyConversionMode.None)
                {
                    if (config.DateOnlyConversionMode == DateOnlyConversionMode.RequireMidnight && dateTimeValue.TimeOfDay != TimeSpan.Zero)
                        throw new InvalidCastException($"Could not convert cell of type DateTime to DateOnly, because DateTime was not at midnight, but at {dateTimeValue:HH:mm:ss}.");

                    newValue = DateOnly.FromDateTime(dateTimeValue);
                    map.MemberAccessor.SetValue(v, newValue);
                    return newValue;
                }

                var vs = itemValue?.ToString();
                if (map.ExcelFormat is not null)
                {
                    if (DateOnly.TryParseExact(vs, map.ExcelFormat, CultureInfo.InvariantCulture, DateTimeStyles.None, out var dateOnlyCustom))
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

            else if (map.ExcludeNullableType == typeof(TimeSpan))
            {
                if (itemValue is TimeSpan)
                {
                    newValue = itemValue;
                    map.MemberAccessor.SetValue(v, newValue);
                    return newValue;
                }

                var vs = itemValue?.ToString();
                if (map.ExcelFormat is not null)
                {
                    if (TimeSpan.TryParseExact(vs, map.ExcelFormat, CultureInfo.InvariantCulture, out var value))
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
            
            else if (map.ExcludeNullableType == typeof(double))
            {
                if (double.TryParse(Convert.ToString(itemValue, config.Culture), NumberStyles.Any, config.Culture, out var doubleValue))
                {
                    newValue = doubleValue;
                }
                else
                {
                    var invariantString = Convert.ToString(itemValue, CultureInfo.InvariantCulture);
                    newValue = double.TryParse(invariantString, NumberStyles.Any, CultureInfo.InvariantCulture, out var value) 
                        ? value 
                        : throw new InvalidCastException();
                }
            }
            
            else if (map.ExcludeNullableType == typeof(bool))
            {
                var vs = itemValue?.ToString();
                newValue = vs switch
                {
                    "1" => true,
                    "0" => false,
                    _ => bool.TryParse(vs, out var parsed) ? parsed : null
                };
            }
            
            else if (map.ExcludeNullableType == typeof(string))
            {
                var strValue = itemValue?.ToString();
                newValue = stringDecoderFunc?.Invoke(strValue) ?? strValue;
            }
            
            else if (map.ExcludeNullableType.IsEnum)
            {
                var fieldInfo = map.ExcludeNullableType.GetFields().FirstOrDefault(e => e.GetCustomAttribute<DescriptionAttribute>(false)?.Description == itemValue?.ToString());
                var value = fieldInfo?.Name ?? itemValue?.ToString() ?? "";
                newValue = Enum.Parse(map.ExcludeNullableType, value, true);
            }

            else if (map.ExcludeNullableType == typeof(Uri))
            {
                var rawValue = itemValue?.ToString();
                if (!Uri.TryCreate(rawValue, UriKind.RelativeOrAbsolute, out var uri))
                    throw new InvalidCastException($"Value \"{rawValue}\" cannot be converted to Uri");
                newValue = uri;
            }

            else
            {
                // Use map.ExcludeNullableType to resolve : https://github.com/mini-software/MiniExcel/issues/138
                newValue = Convert.ChangeType(itemValue, map.ExcludeNullableType, config.Culture);
            }

            map.MemberAccessor.SetValue(v, newValue);
            return newValue;
        }
        catch (Exception ex) when (ex is InvalidCastException or FormatException)
        {
            var columnName = map.ExcelColumnName ?? map.MemberAccessor.Name;
            var errorRow = rowIndex + 1;
            
            throw new ValueNotAssignableException(
                columnName: columnName, 
                row: errorRow, 
                value: itemValue, 
                columnType: map.ExcludeNullableType,
                message: $"The value {itemValue} cannot be assigned to type {map.ExcludeNullableType.Name}."
            );
        }
    }
}
