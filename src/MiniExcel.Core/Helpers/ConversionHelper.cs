using System.Collections.Concurrent;
using System.Linq.Expressions;

namespace MiniExcelLib.Core.Helpers;

/// <summary>
/// Optimized value conversion with caching
/// </summary>
internal static class ConversionHelper
{
    // Cache compiled conversion delegates
    private static readonly ConcurrentDictionary<(Type Source, Type Target), Func<object, object?>> ConversionCache = new();
    
    public static object? ConvertValue(object value, Type targetType, string? format = null)
    {
        var sourceType = value.GetType();
        
        // Fast path: no conversion needed
        if (targetType.IsAssignableFrom(sourceType))
            return value;
        
        // Get or create cached converter
        var key = (sourceType, targetType);
        var converter = ConversionCache.GetOrAdd(key, CreateConverter);
        
        try
        {
            var result = converter(value);
            
            // Note: Format is for writing/display, not for reading
            // When reading, we return the typed value, not formatted string
            
            return result;
        }
        catch
        {
            // Fallback to basic conversion
            return ConvertValueFallback(value, targetType);
        }
    }
    
    private static Func<object, object?> CreateConverter((Type Source, Type Target) types)
    {
        var (sourceType, targetType) = types;
        
        // Handle nullable types
        var underlyingType = Nullable.GetUnderlyingType(targetType) ?? targetType;
        
        // Special case for string source (most common in Excel)
        if (sourceType == typeof(string))
        {
            return CreateStringConverter(underlyingType, targetType != underlyingType);
        }
        
        // Try to create expression-based converter
        try
        {
            var parameter = Expression.Parameter(typeof(object), "value");
            var convert = Expression.Convert(
                Expression.Convert(parameter, sourceType),
                targetType
            );
            var lambda = Expression.Lambda<Func<object, object?>>(
                Expression.Convert(convert, typeof(object)),
                parameter
            );
            return lambda.Compile();
        }
        catch
        {
            // Fallback to runtime conversion
            return value => ConvertValueFallback(value, targetType);
        }
    }
    
    private static Func<object, object?> CreateStringConverter(Type targetType, bool isNullable)
    {
        // Optimized converters for common types from string
        if (targetType == typeof(int))
        {
            return value => 
            {
                var str = (string)value;
                if (string.IsNullOrWhiteSpace(str))
                    return isNullable ? null : 0;
                return int.TryParse(str, out var result) ? result : (isNullable ? null : 0);
            };
        }
        
        if (targetType == typeof(long))
        {
            return value =>
            {
                var str = (string)value;
                if (string.IsNullOrWhiteSpace(str))
                    return isNullable ? null : 0L;
                return long.TryParse(str, out var result) ? result : (isNullable ? null : 0L);
            };
        }
        
        if (targetType == typeof(double))
        {
            return value =>
            {
                var str = (string)value;
                if (string.IsNullOrWhiteSpace(str))
                    return isNullable ? null : 0.0;
                return double.TryParse(str, NumberStyles.Any, CultureInfo.InvariantCulture, out var result) 
                    ? result : (isNullable ? null : 0.0);
            };
        }
        
        if (targetType == typeof(decimal))
        {
            return value =>
            {
                var str = (string)value;
                if (string.IsNullOrWhiteSpace(str))
                    return isNullable ? null : 0m;
                return decimal.TryParse(str, NumberStyles.Any, CultureInfo.InvariantCulture, out var result) 
                    ? result : (isNullable ? null : 0m);
            };
        }
        
        if (targetType == typeof(bool))
        {
            return value =>
            {
                var str = (string)value;
                if (string.IsNullOrWhiteSpace(str))
                    return isNullable ? null : false;
                return bool.TryParse(str, out var result) ? result : (isNullable ? null : false);
            };
        }
        
        if (targetType == typeof(DateTime))
        {
            return value =>
            {
                var str = (string)value;
                if (string.IsNullOrWhiteSpace(str))
                    return isNullable ? null : DateTime.MinValue;
                return DateTime.TryParse(str, CultureInfo.InvariantCulture, DateTimeStyles.None, out var result) 
                    ? result : (isNullable ? null : DateTime.MinValue);
            };
        }
        
        if (targetType == typeof(Guid))
        {
            return value =>
            {
                var str = (string)value;
                if (string.IsNullOrWhiteSpace(str))
                    return isNullable ? null : Guid.Empty;
                return Guid.TryParse(str, out var result) ? result : (isNullable ? null : Guid.Empty);
            };
        }
        
        // Default converter using Convert.ChangeType
        return value => ConvertValueFallback(value, isNullable ? typeof(Nullable<>).MakeGenericType(targetType) : targetType);
    }
    
    private static object? ConvertValueFallback(object value, Type targetType)
    {
        try
        {
            var underlyingType = Nullable.GetUnderlyingType(targetType);
            if (underlyingType != null)
            {
                if (value is string str && string.IsNullOrWhiteSpace(str))
                    return null;
                    
                return Convert.ChangeType(value, underlyingType, CultureInfo.InvariantCulture);
            }
            
            return Convert.ChangeType(value, targetType, CultureInfo.InvariantCulture);
        }
        catch
        {
            // Last resort: return default value
            return targetType.IsValueType ? Activator.CreateInstance(targetType) : null;
        }
    }
    
    /// <summary>
    /// Clear the conversion cache (useful for testing or memory management)
    /// </summary>
    public static void ClearCache()
    {
        ConversionCache.Clear();
    }
}