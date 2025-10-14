namespace MiniExcelLib.Core.FluentMapping.Configuration;

internal partial class CollectionMappingBuilder<T, TCollection> : ICollectionMappingBuilder<T, TCollection> where TCollection : IEnumerable
{
#if NET7_0_OR_GREATER
    [GeneratedRegex("^[A-Z]+[0-9]+$")] private static partial Regex CellAddressRegexImpl();
    private static readonly Regex CellAddressRegex = CellAddressRegexImpl();
#else
    private static readonly Regex CellAddressRegex = new("^[A-Z]+[0-9]+$",  RegexOptions.Compiled);
#endif

    private readonly CollectionMapping _mapping;
    
    internal CollectionMappingBuilder(CollectionMapping mapping)
    {
        _mapping = mapping;
        // Collections are always vertical (rows) by default
        _mapping.Layout = CollectionLayout.Vertical;
    }
    
    public ICollectionMappingBuilder<T, TCollection> StartAt(string cellAddress)
    {
        if (string.IsNullOrEmpty(cellAddress))
            throw new ArgumentException("Cell address cannot be null or empty", nameof(cellAddress));
        
        // Basic validation for cell address format
        if (!CellAddressRegex.IsMatch(cellAddress))
            throw new ArgumentException($"Invalid cell address format: {cellAddress}. Expected format like A1, B2, AA10, etc.", nameof(cellAddress));
            
        _mapping.StartCell = cellAddress;
        return this;
    }
    
    public ICollectionMappingBuilder<T, TCollection> WithSpacing(int spacing)
    {
        if (spacing < 0)
            throw new ArgumentException("Spacing cannot be negative", nameof(spacing));
            
        _mapping.RowSpacing = spacing;
        return this;
    }
    
    public ICollectionMappingBuilder<T, TCollection> WithItemMapping<TItem>(Action<IMappingConfiguration<TItem>> configure)
    {
        if (configure is null)
            throw new ArgumentNullException(nameof(configure));
            
        var itemConfig = new MappingConfiguration<TItem>();
        configure(itemConfig);
        
        _mapping.ItemConfiguration = itemConfig;
        _mapping.ItemType = typeof(TItem);
 
        return this;
    }
}