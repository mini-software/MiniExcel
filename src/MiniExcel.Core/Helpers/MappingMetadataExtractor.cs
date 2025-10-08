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

        if (propsProperty?.GetValue(nestedMapping) is not IEnumerable properties)
            return null;
        
        var nestedInfo = new NestedMappingInfo
        {
            ItemType = itemType,
            ItemFactory = CollectionAccessor.CreateItemFactory(itemType)
        };
        
        var propertyList = ExtractPropertyList(properties, itemType);
        nestedInfo.Properties = propertyList;

        var collectionsProperty = nestedMappingType.GetProperty("Collections");
        if (collectionsProperty?.GetValue(nestedMapping) is IEnumerable collectionMappings)
        {
            var nestedCollections = new Dictionary<string, NestedCollectionInfo>(StringComparer.Ordinal);

            foreach (var collection in collectionMappings)
            {
                if (collection is not CompiledCollectionMapping compiledCollection)
                    continue;

                var nestedItemType = compiledCollection.ItemType ?? typeof(object);
                var collectionInfo = new NestedCollectionInfo
                {
                    PropertyName = compiledCollection.PropertyName,
                    StartColumn = compiledCollection.StartCellColumn,
                    StartRow = compiledCollection.StartCellRow,
                    Layout = compiledCollection.Layout,
                    RowSpacing = compiledCollection.RowSpacing,
                    ItemType = nestedItemType,
                    Getter = compiledCollection.Getter,
                    Setter = compiledCollection.Setter,
                    ListFactory = () => CollectionAccessor.CreateTypedList(nestedItemType),
                    ItemFactory = CollectionAccessor.CreateItemFactory(nestedItemType)
                };

                if (compiledCollection.Registry is not null && nestedItemType != typeof(object))
                {
                    var childMapping = compiledCollection.Registry.GetCompiledMapping(nestedItemType);
                    if (childMapping is not null)
                    {
                        collectionInfo.NestedMapping = ExtractNestedMappingInfo(childMapping, nestedItemType);
                    }
                }

                nestedCollections[collectionInfo.PropertyName] = collectionInfo;
            }

            if (nestedCollections.Count > 0)
            {
                nestedInfo.Collections = nestedCollections;
            }
        }
        
        return nestedInfo;
    }

    /// <summary>
    /// Extracts a list of property information from a collection of property mapping objects.
    /// </summary>
    /// <param name="properties">The collection of property mappings</param>
    /// <returns>A list of nested property information</returns>
    private static readonly MethodInfo? CreateTypedSetterMethod = typeof(ConversionHelper)
        .GetMethods(BindingFlags.Public | BindingFlags.Static)
        .FirstOrDefault(m => m.Name == nameof(ConversionHelper.CreateTypedPropertySetter) && m.IsGenericMethodDefinition);

    private static List<NestedPropertyInfo> ExtractPropertyList(IEnumerable properties, Type itemType)
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

            if (nameProperty is null || columnProperty is null || getterProperty is null) 
                continue;
            
            var name = nameProperty.GetValue(prop) as string;
            var column = (int)columnProperty.GetValue(prop)!;
            var getter = getterProperty.GetValue(prop) as Func<object, object?>;
            var setter = setterProperty?.GetValue(prop) as Action<object, object?>;
            var propTypeValue = typeProperty?.GetValue(prop) as Type;

            if (setter is null && name is not null)
            {
                var propertyInfo = itemType.GetProperty(name, BindingFlags.Public | BindingFlags.Instance);
                if (propertyInfo?.CanWrite == true)
                {
                    setter = CreateSetterWithConversion(itemType, propertyInfo)
                             ?? CreateFallbackSetter(propertyInfo);
                }
            }

            setter ??= (_, _) => { };
                
            if (name is not null && getter is not null)
            {
                propertyList.Add(new NestedPropertyInfo
                {
                    PropertyName = name,
                    ColumnIndex = column,
                    Getter = getter,
                    Setter = setter,
                    PropertyType = propTypeValue ?? typeof(object)
                });
            }
        }
        
        return propertyList;
    }

    private static Action<object, object?>? CreateSetterWithConversion(Type itemType, PropertyInfo propertyInfo)
    {
        if (CreateTypedSetterMethod is null)
            return null;

        try
        {
            var generic = CreateTypedSetterMethod.MakeGenericMethod(itemType);
            return generic.Invoke(null, new object[] { propertyInfo }) as Action<object, object?>;
        }
        catch
        {
            return null;
        }
    }

    private static Action<object, object?>? CreateFallbackSetter(PropertyInfo propertyInfo)
    {
        try
        {
            var memberSetter = new MemberSetter(propertyInfo);
            return memberSetter.Invoke;
        }
        catch
        {
            return null;
        }
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