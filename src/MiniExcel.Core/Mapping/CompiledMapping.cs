namespace MiniExcelLib.Core.Mapping;

public class CompiledMapping<T>
{
    public string WorksheetName { get; set; } = "Sheet1";
    public IReadOnlyList<CompiledPropertyMapping> Properties { get; set; } = new List<CompiledPropertyMapping>();
    public IReadOnlyList<CompiledCollectionMapping> Collections { get; set; } = new List<CompiledCollectionMapping>();
    
    // Universal optimization structures
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
    /// Collection expansion strategies for handling dynamic collections
    /// </summary>
    public IReadOnlyList<CollectionExpansionInfo>? CollectionExpansions { get; set; }
    
    /// <summary>
    /// Whether this mapping has been optimized with the universal optimization system
    /// </summary>
    public bool IsUniversallyOptimized => OptimizedCellGrid != null && OptimizedBoundaries != null;
    
    /// <summary>
    /// Pre-compiled collection helpers for fast collection handling
    /// </summary>
    public IReadOnlyList<OptimizedCollectionHelper>? OptimizedCollectionHelpers { get; set; }
}

/// <summary>
/// Pre-compiled helpers for collection handling
/// </summary>
public class OptimizedCollectionHelper
{
    public Func<System.Collections.IList> Factory { get; set; } = null!;
    public Func<System.Collections.IList, object> Finalizer { get; set; } = null!;
    public Action<object, object?>? Setter { get; set; }
    public bool IsArray { get; set; }
}

public class CompiledPropertyMapping
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

public class CompiledCollectionMapping
{
    public Func<object, IEnumerable> Getter { get; set; } = null!;
    public string StartCell { get; set; } = null!;
    public int StartCellColumn { get; set; }  // Pre-parsed column index
    public int StartCellRow { get; set; }     // Pre-parsed row index
    public CollectionLayout Layout { get; set; }
    public int RowSpacing { get; set; } = 0;
    public object? ItemMapping { get; set; } // CompiledMapping<TItem>
    public Type? ItemType { get; set; }
    public string PropertyName { get; set; } = null!;
    public Action<object, object?>? Setter { get; set; }
    public MappingRegistry? Registry { get; set; } // For looking up nested type mappings
}

/// <summary>
/// Defines the layout direction for collections in Excel mappings.
/// </summary>
public enum CollectionLayout
{
    /// <summary>Collections expand vertically (downward in rows)</summary>
    Vertical = 0,
    
    /// <summary>Collections expand horizontally (rightward in columns) - DEPRECATED</summary>
    [Obsolete("Horizontal collections are no longer supported. Use Vertical layout instead.")]
    Horizontal = 1,
    
    /// <summary>Collections expand in a grid pattern - DEPRECATED</summary>
    [Obsolete("Grid collections are no longer supported. Use Vertical layout instead.")]
    Grid = 2
}


/// <summary>
/// Represents the type of data a cell contains in the mapping
/// </summary>
public enum CellHandlerType
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
public class OptimizedCellHandler
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
    public int CollectionItemOffset { get; set; } = 0;
    
    /// <summary>For formulas: the formula text</summary>
    public string? Formula { get; set; }
    
    /// <summary>For formatted values: the format string</summary>
    public string? Format { get; set; }
    
    /// <summary>For collection items: reference to the collection mapping</summary>
    public CompiledCollectionMapping? CollectionMapping { get; set; }
    
    /// <summary>For collection items: pre-compiled converter from cell value to collection item type</summary>
    public Func<object?, object?>? CollectionItemConverter { get; set; }
    
    /// <summary>For collections: pre-compiled factory to create the collection instance</summary>
    public Func<System.Collections.IList>? CollectionFactory { get; set; }
    
    /// <summary>For collections: pre-compiled converter from list to final type (e.g., array)</summary>
    public Func<System.Collections.IList, object>? CollectionFinalizer { get; set; }
    
    /// <summary>For collections: whether the target type is an array (vs list)</summary>
    public bool IsArrayTarget { get; set; }
    
    /// <summary>
    /// For multiple items scenario: which item (0, 1, 2...) this handler belongs to.
    /// -1 means this handler applies to all items (unbounded collection).
    /// </summary>
    public int ItemIndex { get; set; } = 0;
    
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
public class OptimizedMappingBoundaries
{
    /// <summary>Minimum row used by any mapping (1-based)</summary>
    public int MinRow { get; set; } = int.MaxValue;
    
    /// <summary>Maximum row used by any mapping (1-based)</summary>
    public int MaxRow { get; set; } = 0;
    
    /// <summary>Minimum column used by any mapping (1-based)</summary>
    public int MinColumn { get; set; } = int.MaxValue;
    
    /// <summary>Maximum column used by any mapping (1-based)</summary>
    public int MaxColumn { get; set; } = 0;
    
    /// <summary>Width of the cell grid (MaxColumn - MinColumn + 1)</summary>
    public int GridWidth => MaxColumn > 0 ? MaxColumn - MinColumn + 1 : 0;
    
    /// <summary>Height of the cell grid (MaxRow - MinRow + 1)</summary>
    public int GridHeight => MaxRow > 0 ? MaxRow - MinRow + 1 : 0;
    
    /// <summary>Total number of items this mapping can handle (based on collection layouts)</summary>
    public int MaxItemCapacity { get; set; } = 1;
    
    /// <summary>Whether this mapping has collections that can expand dynamically</summary>
    public bool HasDynamicCollections { get; set; }
    
    /// <summary>
    /// For multiple items with collections: the height of the repeating pattern.
    /// This is the distance from one item's properties to the next item's properties.
    /// 0 means no repeating pattern (single item or no collections).
    /// </summary>
    public int PatternHeight { get; set; } = 0;
    
    /// <summary>
    /// For multiple items: whether this mapping supports multiple items with collections.
    /// When true, the grid pattern repeats every PatternHeight rows.
    /// </summary>
    public bool IsMultiItemPattern { get; set; } = false;
}

/// <summary>
/// Collection expansion strategy - how to handle collections with unknown sizes
/// </summary>
public class CollectionExpansionInfo
{
    /// <summary>Starting row for expansion</summary>
    public int StartRow { get; set; }
    
    /// <summary>Starting column for expansion</summary>
    public int StartColumn { get; set; }
    
    /// <summary>Layout direction for expansion</summary>
    public CollectionLayout Layout { get; set; }
    
    /// <summary>Row spacing between items</summary>
    public int RowSpacing { get; set; }
    
    /// <summary>Collection mapping this expansion belongs to</summary>
    public CompiledCollectionMapping CollectionMapping { get; set; } = null!;
}