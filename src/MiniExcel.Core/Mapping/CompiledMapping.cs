namespace MiniExcelLib.Core.Mapping;

internal class CompiledMapping<T>
{
    public string WorksheetName { get; set; } = "Sheet1";
    public IReadOnlyList<CompiledPropertyMapping> Properties { get; set; } = new List<CompiledPropertyMapping>();
    public IReadOnlyList<CompiledCollectionMapping> Collections { get; set; } = new List<CompiledCollectionMapping>();
    
    // Optimization structures
    /// <summary>
    /// Pre-calculated cell grid for fast mapping. 
    /// Indexed as [row-relative][column-relative] where indices are relative to MinRow/MinColumn
    /// </summary>
    public OptimizedCellHandler[,]? OptimizedCellGrid { get; set; }
    
    /// <summary>Mapping boundaries and optimization metadata</summary>
    public OptimizedMappingBoundaries? OptimizedBoundaries { get; set; }
    
    /// <summary>
    /// For reading: array of column handlers indexed by (column - MinColumn).
    /// Provides O(1) lookup from column index to property setter.
    /// </summary>
    public OptimizedCellHandler[]? OptimizedColumnHandlers { get; set; }
    
    /// <summary>
    /// Pre-compiled collection helpers for fast collection handling
    /// </summary>
    public IReadOnlyList<OptimizedCollectionHelper>? OptimizedCollectionHelpers { get; set; }
    
    /// <summary>
    /// Pre-compiled nested mapping information for complex collection types.
    /// Indexed by collection index to provide fast access to nested property info.
    /// </summary>
    public IReadOnlyDictionary<int, NestedMappingInfo>? NestedMappings { get; set; }
    
    /// <summary>
    /// Tries to get the cell handler at the specified absolute row and column position.
    /// </summary>
    /// <param name="absoluteRow">The absolute row number (1-based)</param>
    /// <param name="absoluteCol">The absolute column number (1-based)</param>
    /// <param name="handler">The handler if found, or default if not</param>
    /// <returns>True if a non-empty handler was found at the position</returns>
    public bool TryGetHandler(int absoluteRow, int absoluteCol, out OptimizedCellHandler handler)
    {
        handler = default!;
        
        if (OptimizedCellGrid == null || OptimizedBoundaries == null)
            return false;
        
        var relRow = absoluteRow - OptimizedBoundaries.MinRow;
        var relCol = absoluteCol - OptimizedBoundaries.MinColumn;
        
        if (relRow < 0 || relRow >= OptimizedBoundaries.GridHeight ||
            relCol < 0 || relCol >= OptimizedBoundaries.GridWidth)
            return false;
        
        handler = OptimizedCellGrid[relRow, relCol];
        return handler.Type != CellHandlerType.Empty;
    }
    
    /// <summary>
    /// Tries to extract a value from an item using the specified handler.
    /// </summary>
    /// <typeparam name="TItem">The type of the item</typeparam>
    /// <param name="handler">The cell handler containing the value extractor</param>
    /// <param name="item">The item to extract the value from</param>
    /// <param name="value">The extracted value, or null if extraction failed</param>
    /// <returns>True if the value was successfully extracted</returns>
    public bool TryGetValue<TItem>(OptimizedCellHandler handler, TItem? item, out object? value) where TItem : class
    {
        value = null;
        
        if (item == null || handler.ValueExtractor == null)
            return false;
        
        value = handler.ValueExtractor(item, 0);
        return true;
    }
    
    /// <summary>
    /// Tries to set a value on an item using the specified handler.
    /// </summary>
    /// <typeparam name="TItem">The type of the item</typeparam>
    /// <param name="handler">The cell handler containing the value setter</param>
    /// <param name="item">The item to set the value on</param>
    /// <param name="value">The value to set</param>
    /// <returns>True if the value was successfully set</returns>
    public bool TrySetValue<TItem>(OptimizedCellHandler handler, TItem? item, object? value) where TItem : class
    {
        if (item == null || handler.ValueSetter == null)
            return false;
        
        handler.ValueSetter(item, value);
        return true;
    }
    
    /// <summary>
    /// Convenience method that combines TryGetHandler and TryGetValue in a single call.
    /// </summary>
    /// <typeparam name="TItem">The type of the item</typeparam>
    /// <param name="row">The absolute row number (1-based)</param>
    /// <param name="col">The absolute column number (1-based)</param>
    /// <param name="item">The item to extract the value from</param>
    /// <param name="value">The extracted value, or null if not found</param>
    /// <returns>True if a value was successfully extracted</returns>
    public bool TryGetCellValue<TItem>(int row, int col, TItem item, out object? value) where TItem : class
    {
        value = null;
        
        if (!TryGetHandler(row, col, out var handler))
            return false;
        
        return TryGetValue(handler, item, out value);
    }
    
    /// <summary>
    /// Convenience method that combines TryGetHandler and TrySetValue in a single call.
    /// </summary>
    /// <typeparam name="TItem">The type of the item</typeparam>
    /// <param name="row">The absolute row number (1-based)</param>
    /// <param name="col">The absolute column number (1-based)</param>
    /// <param name="item">The item to set the value on</param>
    /// <param name="value">The value to set</param>
    /// <returns>True if the value was successfully set</returns>
    public bool TrySetCellValue<TItem>(int row, int col, TItem item, object? value) where TItem : class
    {
        if (!TryGetHandler(row, col, out var handler))
            return false;
        
        return TrySetValue(handler, item, value);
    }
    
    /// <summary>
    /// Tries to set a property value on an item using the compiled property mapping.
    /// </summary>
    /// <typeparam name="TItem">The type of the item</typeparam>
    /// <param name="property">The compiled property mapping</param>
    /// <param name="item">The item to set the value on</param>
    /// <param name="value">The value to set</param>
    /// <returns>True if the value was successfully set</returns>
    public bool TrySetPropertyValue<TItem>(CompiledPropertyMapping property, TItem item, object? value) where TItem : class
    {
        if (property.Setter == null)
            return false;
        
        property.Setter(item, value);
        return true;
    }
    
    /// <summary>
    /// Tries to set a collection value on an item using the compiled collection mapping.
    /// </summary>
    /// <typeparam name="TItem">The type of the item</typeparam>
    /// <param name="collection">The compiled collection mapping</param>
    /// <param name="item">The item to set the collection on</param>
    /// <param name="value">The collection value to set</param>
    /// <returns>True if the collection was successfully set</returns>
    public bool TrySetCollectionValue<TItem>(CompiledCollectionMapping collection, TItem item, object? value) where TItem : class
    {
        if (collection.Setter == null)
            return false;
        
        collection.Setter(item, value);
        return true;
    }
}

/// <summary>
/// Pre-compiled helpers for collection handling
/// </summary>
internal class OptimizedCollectionHelper
{
    public Func<IList> Factory { get; set; } = null!;
    public Func<IList, object> Finalizer { get; set; } = null!;
    public Action<object, object?>? Setter { get; set; }
    public bool IsArray { get; set; }
    public Func<object?> DefaultItemFactory { get; set; } = null!;
    public Type ItemType { get; set; } = null!;
    public bool IsItemValueType { get; set; }
    public bool IsItemPrimitive { get; set; }
    public object? DefaultValue { get; set; }
}

internal class CompiledPropertyMapping
{
    public Func<object, object> Getter { get; set; } = null!;
    public string CellAddress { get; set; } = null!;
    public int CellColumn { get; set; }  // Pre-parsed column index
    public int CellRow { get; set; }     // Pre-parsed row index
    public string? Format { get; set; }
    public string? Formula { get; set; }
    public Type PropertyType { get; set; } = null!;
    public string PropertyName { get; set; } = null!;
    public Action<object, object?>? Setter { get; set; }
}

internal class CompiledCollectionMapping
{
    public Func<object, IEnumerable> Getter { get; set; } = null!;
    public int StartCellColumn { get; set; }  // Pre-parsed column index
    public int StartCellRow { get; set; }     // Pre-parsed row index
    public CollectionLayout Layout { get; set; }
    public int RowSpacing { get; set; }
    public Type? ItemType { get; set; }
    public string PropertyName { get; set; } = null!;
    public Action<object, object?>? Setter { get; set; }
    public MappingRegistry? Registry { get; set; } // For looking up nested type mappings
}

/// <summary>
/// Defines the layout direction for collections in Excel mappings.
/// </summary>
internal enum CollectionLayout
{
    /// <summary>Collections expand vertically (downward in rows)</summary>
    Vertical = 0,
}


/// <summary>
/// Represents the type of data a cell contains in the mapping
/// </summary>
internal enum CellHandlerType
{
    /// <summary>Cell is empty/unused</summary>
    Empty,
    /// <summary>Cell contains a simple property value</summary>
    Property,
    /// <summary>Cell contains an item from a collection</summary>
    CollectionItem,
    /// <summary>Cell contains a formula</summary>
    Formula
}

/// <summary>
/// Pre-compiled handler for a specific cell in the mapping grid.
/// Contains all information needed to extract/set values for that cell without runtime parsing.
/// </summary>
internal class OptimizedCellHandler
{
    /// <summary>Type of data this cell contains</summary>
    public CellHandlerType Type { get; set; } = CellHandlerType.Empty;
    
    /// <summary>For Property/Formula: direct property getter. For CollectionItem: collection getter + indexer</summary>
    public Func<object, int, object?>? ValueExtractor { get; set; }
    
    /// <summary>For reading: direct property setter with conversion built-in</summary>
    public Action<object, object?>? ValueSetter { get; set; }
    
    /// <summary>Property name for debugging/error reporting</summary>
    public string? PropertyName { get; set; }
    
    /// <summary>For collection items: which collection this belongs to</summary>
    public int CollectionIndex { get; set; } = -1;
    
    /// <summary>For collection items: offset within collection</summary>
    public int CollectionItemOffset { get; set; }
    
    /// <summary>For formulas: the formula text</summary>
    public string? Formula { get; set; }
    
    /// <summary>For formatted values: the format string</summary>
    public string? Format { get; set; }
    
    /// <summary>For collection items: reference to the collection mapping</summary>
    public CompiledCollectionMapping? CollectionMapping { get; set; }
    
    /// <summary>For collection items: pre-compiled converter from cell value to collection item type</summary>
    public Func<object?, object?>? CollectionItemConverter { get; set; }
    
    /// <summary>
    /// For multiple items scenario: which item (0, 1, 2...) this handler belongs to.
    /// -1 means this handler applies to all items (unbounded collection).
    /// </summary>
    public int ItemIndex { get; set; }
    
    /// <summary>
    /// For collection handlers: the row where this collection stops reading (exclusive).
    /// -1 means unbounded (continue until no more data).
    /// </summary>
    public int BoundaryRow { get; set; } = -1;
    
    /// <summary>
    /// For collection handlers: the column where this collection stops reading (exclusive).
    /// -1 means unbounded (continue until no more data).
    /// </summary>
    public int BoundaryColumn { get; set; } = -1;
}

/// <summary>
/// Optimized mapping boundaries and metadata
/// </summary>
internal class OptimizedMappingBoundaries
{
    /// <summary>Minimum row used by any mapping (1-based)</summary>
    public int MinRow { get; set; } = int.MaxValue;
    
    /// <summary>Maximum row used by any mapping (1-based)</summary>
    public int MaxRow { get; set; }
    
    /// <summary>Minimum column used by any mapping (1-based)</summary>
    public int MinColumn { get; set; } = int.MaxValue;
    
    /// <summary>Maximum column used by any mapping (1-based)</summary>
    public int MaxColumn { get; set; }
    
    /// <summary>Width of the cell grid (MaxColumn - MinColumn + 1)</summary>
    public int GridWidth => MaxColumn > 0 ? MaxColumn - MinColumn + 1 : 0;
    
    /// <summary>Height of the cell grid (MaxRow - MinRow + 1)</summary>
    public int GridHeight => MaxRow > 0 ? MaxRow - MinRow + 1 : 0;
    
    /// <summary>Whether this mapping has collections that can expand dynamically</summary>
    public bool HasDynamicCollections { get; set; }
    
    /// <summary>
    /// For multiple items with collections: the height of the repeating pattern.
    /// This is the distance from one item's properties to the next item's properties.
    /// 0 means no repeating pattern (single item or no collections).
    /// </summary>
    public int PatternHeight { get; set; }
    
    /// <summary>
    /// For multiple items: whether this mapping supports multiple items with collections.
    /// When true, the grid pattern repeats every PatternHeight rows.
    /// </summary>
    public bool IsMultiItemPattern { get; set; }
}