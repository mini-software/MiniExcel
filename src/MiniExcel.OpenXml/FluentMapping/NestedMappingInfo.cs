using System.Collections;

namespace MiniExcelLib.OpenXml.FluentMapping;

/// <summary>
/// Stores pre-compiled information about nested properties in collection mappings.
/// This eliminates the need for runtime reflection when processing complex collection types.
/// </summary>
internal class NestedMappingInfo
{
    /// <summary>
    /// Pre-compiled property accessors for the nested type.
    /// </summary>
    public IReadOnlyList<NestedPropertyInfo> Properties { get; set; } = new List<NestedPropertyInfo>();
    
    /// <summary>
    /// Pre-compiled nested collection accessors keyed by property name.
    /// </summary>
    public IReadOnlyDictionary<string, NestedCollectionInfo> Collections { get; set; } = new Dictionary<string, NestedCollectionInfo>();

    /// <summary>
    /// The type of items in the collection.
    /// </summary>
    public Type ItemType { get; set; } = null!;
    
    /// <summary>
    /// Pre-compiled factory for creating instances of the item type.
    /// </summary>
    public Func<object?> ItemFactory { get; set; } = null!;
}

/// <summary>
/// Pre-compiled information about a single property in a nested type.
/// </summary>
internal class NestedPropertyInfo
{
    /// <summary>
    /// The name of the property.
    /// </summary>
    public string PropertyName { get; set; } = null!;
    
    /// <summary>
    /// The Excel column index (1-based) where this property is mapped.
    /// </summary>
    public int ColumnIndex { get; set; }
    
    /// <summary>
    /// Pre-compiled getter for extracting the property value from an object.
    /// </summary>
    public Func<object, object?> Getter { get; set; } = null!;
    
    /// <summary>
    /// Pre-compiled setter for setting the property value on an object.
    /// </summary>
    public Action<object, object?> Setter { get; set; } = null!;
    
    /// <summary>
    /// The type of the property.
    /// </summary>
    public Type PropertyType { get; set; } = null!;
}

/// <summary>
/// Pre-compiled information about a nested collection within a complex type.
/// </summary>
internal class NestedCollectionInfo
{
    public string PropertyName { get; set; } = null!;
    public int StartColumn { get; set; }
    public int StartRow { get; set; }
    public CollectionLayout Layout { get; set; }
    public int RowSpacing { get; set; }
    public Type ItemType { get; set; } = typeof(object);
    public Func<object, IEnumerable> Getter { get; set; } = null!;
    public Action<object, object?>? Setter { get; set; }
    public Func<IList> ListFactory { get; set; } = null!;
    public Func<object?> ItemFactory { get; set; } = null!;
    public NestedMappingInfo? NestedMapping { get; set; }
}