using System.Linq.Expressions;

namespace MiniExcelLib.Core.Mapping.Configuration;

internal class MappingConfiguration<T> : IMappingConfiguration<T>
{
    internal readonly List<PropertyMapping> PropertyMappings = [];
    internal readonly List<CollectionMapping> CollectionMappings = [];
    internal string? WorksheetName { get; private set; }
    
    public IPropertyMappingBuilder<T, TProperty> Property<TProperty>(
        Expression<Func<T, TProperty>> property)
    {
        if (property is null)
            throw new ArgumentNullException(nameof(property));
            
        var mapping = new PropertyMapping
        {
            Expression = property,
            PropertyType = typeof(TProperty)
        };
        PropertyMappings.Add(mapping);

        return new PropertyMappingBuilder<T, TProperty>(mapping);
    }
    
    public ICollectionMappingBuilder<T, TCollection> Collection<TCollection>(
        Expression<Func<T, TCollection>> collection) where TCollection : IEnumerable
    {
        if (collection is null)
            throw new ArgumentNullException(nameof(collection));
            
        var mapping = new CollectionMapping
        {
            Expression = collection,
            PropertyType = typeof(TCollection)
        };
        CollectionMappings.Add(mapping);
        
        return new CollectionMappingBuilder<T, TCollection>(mapping);
    }
    
    public IMappingConfiguration<T> ToWorksheet(string worksheetName)
    {
        if (string.IsNullOrEmpty(worksheetName))
            throw new ArgumentException("Sheet names cannot be empty or null", nameof(worksheetName));
        if (worksheetName.Length > 31)
            throw new ArgumentException("Sheet names must be less than 31 characters", nameof(worksheetName));
            
        WorksheetName = worksheetName;
        return this;
    }
}

internal class PropertyMapping
{
    public LambdaExpression Expression { get; set; } = null!;
    public Type PropertyType { get; set; } = null!;
    public string? CellAddress { get; set; }
    public string? Format { get; set; }
    public string? Formula { get; set; }
}

internal class CollectionMapping : PropertyMapping
{
    public string? StartCell { get; set; }
    public CollectionLayout Layout { get; set; } = CollectionLayout.Vertical;
    public int RowSpacing { get; set; }
    public object? ItemConfiguration { get; set; }
    public Type? ItemType { get; set; }
}