using System.ComponentModel;
using MiniExcelLib.Core.Exceptions;

namespace MiniExcelLib.Core.Reflection;

public static partial class MiniExcelMapper
{
    [CreateSyncVersion]
    public static async IAsyncEnumerable<T> MapQueryAsync<T>(IAsyncEnumerable<IDictionary<string, object?>> values, string startCell, bool mapHeaderAsData, bool trimColumnNames, MiniExcelBaseConfiguration configuration, [EnumeratorCancellation] CancellationToken cancellationToken = default) where T : class, new()
    {
        cancellationToken.ThrowIfCancellationRequested();

        var type = typeof(T);

        //TODO:need to optimize
        List<MiniExcelColumnInfo> props = [];
        Dictionary<string, int> headersDic = [];
        string[] keys = [];
        var first = true;
        var rowIndex = 0;

        await foreach (var item in values.WithCancellation(cancellationToken).ConfigureAwait(false))
        {
            if (first)
            {
                keys = item.Keys.ToArray();
                headersDic = CustomPropertyHelper.GetHeaders(item, trimColumnNames);

                //TODO: alert don't duplicate column name
                props = CustomPropertyHelper.GetExcelCustomPropertyInfos(type, keys, configuration);
                first = false;

                // if we treat the header as data we move forwards with the mapping otherwise we jump to the next iteration
                if (!mapHeaderAsData) 
                    continue;
            }

            var v = new T();
            foreach (var pInfo in props)
            {
                if (pInfo.ExcelColumnAliases is not null)
                {
                    foreach (var alias in pInfo.ExcelColumnAliases)
                    {
                        if (headersDic?.TryGetValue(alias, out var columnId) ?? false)
                        {
                            var columnName = keys[columnId];
                            item.TryGetValue(columnName, out var aliasItemValue);

                            if (aliasItemValue is not null)
                            {
                                var newAliasValue = MapValue(v, pInfo, aliasItemValue, rowIndex, startCell, configuration);
                            }
                        }
                    }
                }

                //Q: Why need to check every time? A: it needs to check everytime, because it's dictionary
                object? itemValue = null;
                if (pInfo.ExcelIndexName is not null && (keys?.Contains(pInfo.ExcelIndexName) ?? false))
                {
                    item.TryGetValue(pInfo.ExcelIndexName, out itemValue);
                }
                else if (pInfo.ExcelColumnName is not null && (headersDic?.TryGetValue(pInfo.ExcelColumnName, out var columnId) ?? false))
                {
                    var columnName = keys[columnId];
                    item.TryGetValue(columnName, out itemValue);
                }

                if (itemValue is not null)
                {
                    var newValue = MapValue(v, pInfo, itemValue, rowIndex, startCell, configuration);
                }
            }
            
            rowIndex++;
            yield return v;
        }
    }
    
    internal static object? MapValue<T>(T v, MiniExcelColumnInfo pInfo, object itemValue, int rowIndex, string startCell, MiniExcelBaseConfiguration config) where T : class, new()
    {
        try
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

                if (itemValue is DateTime dateTimeValue &&
                    config.DateOnlyConversionMode is DateOnlyConversionMode.EnforceMidnight
                        or DateOnlyConversionMode.IgnoreTimePart)
                {
                    if (config.DateOnlyConversionMode == DateOnlyConversionMode.EnforceMidnight && dateTimeValue.TimeOfDay != TimeSpan.Zero)
                    {
                        throw new InvalidCastException(
                            $"Could not convert cell of type DateTime to DateOnly, because DateTime was not at midnight, but at {dateTimeValue:HH:mm:ss}.");
                    }
                    return DateOnly.FromDateTime(dateTimeValue);
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
                newValue = XmlHelper.DecodeString(itemValue?.ToString());
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
        catch (Exception ex) when (ex is InvalidCastException or FormatException)
        {
            var columnName = pInfo.ExcelColumnName ?? pInfo.Property.Name;
            var startRowIndex = ReferenceHelper.ConvertCellToCoordinates(startCell).Item2;
            var errorRow = startRowIndex + rowIndex + 1;
                
            var msg = $"ColumnName: {columnName}, CellRow: {errorRow}, Value: {itemValue}. The value cannot be cast to type {pInfo.Property.Info.PropertyType.Name}.";
            throw new MiniExcelInvalidCastException(columnName, errorRow, itemValue, pInfo.Property.Info.PropertyType, msg);
        }
    }
}