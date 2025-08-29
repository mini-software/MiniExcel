namespace MiniExcelLib.Core.Helpers;

/// <summary>
/// Helper class for extracting mapping metadata using reflection.
/// Consolidates reflection-based property extraction logic to reduce duplication and improve performance.
/// </summary>
internal static class MappingMetadataExtractor
{
    /// <summary>
    /// Extracts nested mapping information from a compiled mapping object.
    /// This method minimizes reflection by extracting properties once at compile time.
    /// </summary>
    /// <param name="nestedMapping">The nested mapping object to extract information from</param>
    /// <param name="itemType">The type of items in the nested mapping</param>
    /// <returns>Nested mapping information or null if extraction fails</returns>
    public static NestedMappingInfo? ExtractNestedMappingInfo(object nestedMapping, Type itemType)
    {
        // Use reflection minimally to extract properties from the nested mapping
        // This is done once at compile time, not at runtime
        var nestedMappingType = nestedMapping.GetType();
        var propsProperty = nestedMappingType.GetProperty("Properties");
        if (propsProperty == null) return null;
        
        var properties = propsProperty.GetValue(nestedMapping) as IEnumerable;
        if (properties == null) return null;
        
        var nestedInfo = new NestedMappingInfo
        {
            ItemType = itemType,
            ItemFactory = CollectionAccessor.CreateItemFactory(itemType)
        };
        
        var propertyList = ExtractPropertyList(properties);
        nestedInfo.Properties = propertyList;
        
        return nestedInfo;
    }

    /// <summary>
    /// Extracts a list of property information from a collection of property mapping objects.
    /// </summary>
    /// <param name="properties">The collection of property mappings</param>
    /// <returns>A list of nested property information</returns>
    private static List<NestedPropertyInfo> ExtractPropertyList(IEnumerable properties)
    {
        var propertyList = new List<NestedPropertyInfo>();
        
        foreach (var prop in properties)
        {
            var propType = prop.GetType();
            var nameProperty = propType.GetProperty("PropertyName");
            var columnProperty = propType.GetProperty("CellColumn");
            var getterProperty = propType.GetProperty("Getter");
            var setterProperty = propType.GetProperty("Setter");
            var typeProperty = propType.GetProperty("PropertyType");

            if (nameProperty == null || columnProperty == null || getterProperty == null) continue;
            
            var name = nameProperty.GetValue(prop) as string;
            var column = (int)columnProperty.GetValue(prop)!;
            var getter = getterProperty.GetValue(prop) as Func<object, object?>;
            var setter = setterProperty?.GetValue(prop) as Action<object, object?>;
            var propTypeValue = typeProperty?.GetValue(prop) as Type;
                
            if (name != null && getter != null)
            {
                propertyList.Add(new NestedPropertyInfo
                {
                    PropertyName = name,
                    ColumnIndex = column,
                    Getter = getter,
                    Setter = setter ?? ((_, _) => { }),
                    PropertyType = propTypeValue ?? typeof(object)
                });
            }
        }
        
        return propertyList;
    }

    /// <summary>
    /// Gets a specific property by name from a type.
    /// </summary>
    /// <param name="type">The type to search</param>
    /// <param name="propertyName">The name of the property</param>
    /// <returns>PropertyInfo if found, otherwise null</returns>
    public static PropertyInfo? GetPropertyByName(Type type, string propertyName)
    {
        return type.GetProperty(propertyName, BindingFlags.Public | BindingFlags.Instance);
    }

    private static bool IsSimpleType(Type type)
    {
        return type == typeof(string) || type.IsValueType || type.IsPrimitive;
    }

    /// <summary>
    /// Determines if a type is a complex type that likely has nested properties.
    /// </summary>
    /// <param name="type">The type to check</param>
    /// <returns>True if the type is considered complex</returns>
    public static bool IsComplexType(Type type)
    {
        return !IsSimpleType(type) && type != typeof(object);
    }
}