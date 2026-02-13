using System.Collections.Concurrent;
using System.Globalization;
using System.Linq.Expressions;
using System.Reflection;

namespace MiniExcelLib.OpenXml.FluentMapping.Helpers;

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
            return CreateStringConverter(underlyingType, targetType != underlyingType);
        
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
                var str = value as string;
                if (string.IsNullOrWhiteSpace(str))
                    return isNullable ? null : 0;

                return int.TryParse(str, out var result) 
                    ? result 
                    : isNullable ? null : 0;
            };
        }
        
        if (targetType == typeof(long))
        {
            return value =>
            {
                var str = value as string;
                if (string.IsNullOrWhiteSpace(str))
                    return isNullable ? null : 0L;
                
                return long.TryParse(str, out var result) 
                    ? result 
                    : isNullable ? null : 0L;
            };
        }
        
        if (targetType == typeof(double))
        {
            return value =>
            {
                var str = value as string;
                if (string.IsNullOrWhiteSpace(str))
                    return isNullable ? null : 0D;

                return double.TryParse(str, NumberStyles.Any, CultureInfo.InvariantCulture, out var result) 
                    ? result 
                    : isNullable ? null : 0D;
            };
        }
        
        if (targetType == typeof(decimal))
        {
            return value =>
            {
                var str = value as string;
                if (string.IsNullOrWhiteSpace(str))
                    return isNullable ? null : 0M;
                
                return decimal.TryParse(str, NumberStyles.Any, CultureInfo.InvariantCulture, out var result) 
                    ? result 
                    : isNullable ? null : 0M;
            };
        }
        
        if (targetType == typeof(bool))
        {
            return value =>
            {
                var str = value as string;
                if (string.IsNullOrWhiteSpace(str))
                    return isNullable ? null : false;
                
                return bool.TryParse(str, out var result) 
                    ? result 
                    : isNullable ? null : false;
            };
        }
        
        if (targetType == typeof(DateTime))
        {
            return value =>
            {
                var str = value as string;
                if (string.IsNullOrWhiteSpace(str))
                    return isNullable ? null : DateTime.MinValue;
                
                return DateTime.TryParse(str, CultureInfo.InvariantCulture, DateTimeStyles.None, out var result) 
                    ? result 
                    : isNullable ? null : DateTime.MinValue;
            };
        }
        
        if (targetType == typeof(TimeSpan))
        {
            return value =>
            {
                var str = value as string;
                if (string.IsNullOrWhiteSpace(str))
                    return isNullable ? null : TimeSpan.MinValue;
                
                return TimeSpan.TryParse(str, CultureInfo.InvariantCulture, out var result) 
                    ? result 
                    : isNullable ? null : TimeSpan.MinValue;
            };
        }
        
        if (targetType == typeof(Guid))
        {
            return value =>
            {
                var str = value as string;
                if (string.IsNullOrWhiteSpace(str))
                    return isNullable ? null : Guid.Empty;
                
                return Guid.TryParse(str, out var result) 
                    ? result 
                    : isNullable ? null : Guid.Empty;
            };
        }
        
        // Default converter using Convert.ChangeType
        var newType = isNullable ? typeof(Nullable<>).MakeGenericType(targetType) : targetType;
        return value => ConvertValueFallback(value, newType);
    }
    
    private static object? ConvertValueFallback(object? value, Type targetType)
    {
        try
        {
            if (Nullable.GetUnderlyingType(targetType) is { } underlyingType)
            {
                return value is not (null or "" or " ") 
                    ? Convert.ChangeType(value, underlyingType, CultureInfo.InvariantCulture) 
                    : null;
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
    /// Creates a compiled setter expression for the specified target type with proper conversion handling.
    /// This consolidates type conversion logic from various parts of the codebase.
    /// </summary>
    /// <param name="targetType">The target property type</param>
    /// <param name="valueParameter">The parameter expression for the input value</param>
    /// <returns>A compiled expression that converts and assigns values</returns>
    private static Expression CreateTypedConversionExpression(Type targetType, ParameterExpression valueParameter)
    {
        // Handle nullable types
        var underlyingType = Nullable.GetUnderlyingType(targetType);
        var isNullable = underlyingType is not null;
        var effectiveType = underlyingType ?? targetType;
        
        Expression convertExpression;
        
        // Create conversion expression based on effective type
        if (effectiveType == typeof(int))
        {
            var convertMethod = typeof(Convert).GetMethod("ToInt32", [typeof(object)]);
            convertExpression = Expression.Call(convertMethod!, valueParameter);
        }
        else if (effectiveType == typeof(decimal))
        {
            var convertMethod = typeof(Convert).GetMethod("ToDecimal", [typeof(object)]);
            convertExpression = Expression.Call(convertMethod!, valueParameter);
        }
        else if (effectiveType == typeof(long))
        {
            var convertMethod = typeof(Convert).GetMethod("ToInt64", [typeof(object)]);
            convertExpression = Expression.Call(convertMethod!, valueParameter);
        }
        else if (effectiveType == typeof(float))
        {
            var convertMethod = typeof(Convert).GetMethod("ToSingle", [typeof(object)]);
            convertExpression = Expression.Call(convertMethod!, valueParameter);
        }
        else if (effectiveType == typeof(double))
        {
            var convertMethod = typeof(Convert).GetMethod("ToDouble", [typeof(object)]);
            convertExpression = Expression.Call(convertMethod!, valueParameter);
        }
        else if (effectiveType == typeof(DateTime))
        {
            var convertMethod = typeof(Convert).GetMethod("ToDateTime", [typeof(object)]);
            convertExpression = Expression.Call(convertMethod!, valueParameter);
        }
        else if (effectiveType == typeof(bool))
        {
            var convertMethod = typeof(Convert).GetMethod("ToBoolean", [typeof(object)]);
            convertExpression = Expression.Call(convertMethod!, valueParameter);
        }
        else if (effectiveType == typeof(string))
        {
            var convertMethod = typeof(Convert).GetMethod("ToString", [typeof(object)]);
            convertExpression = Expression.Call(convertMethod!, valueParameter);
        }
        else
        {
            // Default: direct cast for other types
            convertExpression = Expression.Convert(valueParameter, effectiveType);
        }
        
        // If the target type is nullable, convert the result to nullable
        if (isNullable)
        {
            convertExpression = Expression.Convert(convertExpression, targetType);
        }
        
        return convertExpression;
    }

    /// <summary>
    /// Creates a compiled property setter with type conversion for the specified property.
    /// </summary>
    /// <typeparam name="T">The containing type</typeparam>
    /// <param name="propertyInfo">The property to create a setter for</param>
    /// <returns>A compiled setter action or null if the property is not settable</returns>
    public static Action<object, object?>? CreateTypedPropertySetter<T>(PropertyInfo propertyInfo)
    {
        if (!propertyInfo.CanWrite)
            return null;
            
        var setterParam = Expression.Parameter(typeof(object), "obj");
        var valueParam = Expression.Parameter(typeof(object), "value");
        var castObj = Expression.Convert(setterParam, typeof(T));
        
        // Use the centralized conversion logic
        var convertedValue = CreateTypedConversionExpression(propertyInfo.PropertyType, valueParam);
        
        var assign = Expression.Assign(Expression.Property(castObj, propertyInfo), convertedValue);
        var setterLambda = Expression.Lambda<Action<object, object?>>(assign, setterParam, valueParam);
        return setterLambda.Compile();
    }

    /// <summary>
    /// Clear the conversion cache (useful for testing or memory management)
    /// </summary>
    public static void ClearCache()
    {
        ConversionCache.Clear();
    }
}