using System.Linq.Expressions;
using MiniExcelLib.Core.Mapping.Configuration;

namespace MiniExcelLib.Core.Mapping;

/// <summary>
/// Compiles mapping configurations into optimized runtime representations for efficient Excel read/write operations.
/// Uses a universal optimization strategy with pre-compiled property accessors and cell grids.
/// </summary>
internal static class MappingCompiler
{
    // Conservative estimates for collection bounds when actual size is unknown
    private const int DefaultCollectionHeight = 100;
    private const int DefaultGridSize = 10;
    private const int MaxPatternHeight = 20;
    private const int MinItemsForPatternCalc = 2;
    
    /// <summary>
    /// Compiles a mapping configuration into an optimized runtime representation.
    /// </summary>
    public static CompiledMapping<T> Compile<T>(MappingConfiguration<T>? configuration, MappingRegistry? registry = null)
    {
        if (configuration is null)
            throw new ArgumentNullException(nameof(configuration));

        var properties = new List<CompiledPropertyMapping>();
        var collections = new List<CompiledCollectionMapping>();
        
        // Compile property mappings
        foreach (var prop in configuration.PropertyMappings)
        {
            if (string.IsNullOrEmpty(prop.CellAddress))
                throw new InvalidOperationException($"Property mapping must specify a cell address using ToCell()");

            var propertyName = GetPropertyName(prop.Expression);
            
            // Build getter expression
            var parameter = Expression.Parameter(typeof(object), "obj");
            var cast = Expression.Convert(parameter, typeof(T));
            var propertyAccess = Expression.Invoke(prop.Expression, cast);
            var convertToObject = Expression.Convert(propertyAccess, typeof(object));
            var lambda = Expression.Lambda<Func<object, object>>(convertToObject, parameter);
            var compiled = lambda.Compile();
            
            // Create setter with proper type conversion using centralized logic
            Action<object, object?>? setter = null;
            if (prop.Expression.Body is MemberExpression { Member: PropertyInfo propInfo })
            {
                setter = ConversionHelper.CreateTypedPropertySetter<T>(propInfo);
            }
            
            // Pre-parse cell coordinates for runtime performance
            if (prop.CellAddress is null) 
                continue;
            
            ReferenceHelper.ParseReference(prop.CellAddress, out int cellCol, out int cellRow);

            properties.Add(new CompiledPropertyMapping
            {
                Getter = compiled,
                CellAddress = prop.CellAddress ?? string.Empty,
                CellColumn = cellCol,
                CellRow = cellRow,
                Format = prop.Format,
                Formula = prop.Formula,
                PropertyType = prop.PropertyType,
                PropertyName = propertyName,
                Setter = setter
            });
        }
        
        // Compile collection mappings
        foreach (var coll in configuration.CollectionMappings)
        {
            if (string.IsNullOrEmpty(coll.StartCell))
                throw new InvalidOperationException($"Collection mapping must specify a start cell using StartAt()");

            var parameter = Expression.Parameter(typeof(object), "obj");
            var cast = Expression.Convert(parameter, typeof(T));
            var collectionAccess = Expression.Invoke(coll.Expression, cast);
            var convertToEnumerable = Expression.Convert(collectionAccess, typeof(IEnumerable));
            var lambda = Expression.Lambda<Func<object, IEnumerable>>(convertToEnumerable, parameter);
            var compiled = lambda.Compile();

            // Extract property name from expression
            var collectionPropertyName = GetPropertyName(coll.Expression);
            
            // Determine the item type using centralized logic
            var collectionType = coll.PropertyType;
            Type? itemType = CollectionAccessor.GetItemType(collectionType);
            
            // Create setter for collection
            Action<object, object?>? collectionSetter = null;
            if (coll.Expression.Body is MemberExpression { Member: PropertyInfo collPropInfo })
            {
                var memberSetter = new MemberSetter(collPropInfo);
                collectionSetter = memberSetter.Invoke;
            }
            
            // Pre-parse start cell coordinates
            if (coll.StartCell is null) 
                continue;

            ReferenceHelper.ParseReference(coll.StartCell, out int startCol, out int startRow);

            var compiledCollection = new CompiledCollectionMapping
            {
                Getter = compiled,
                StartCellColumn = startCol,
                StartCellRow = startRow,
                Layout = coll.Layout,
                RowSpacing = coll.RowSpacing,
                ItemType = itemType ?? coll.ItemType,
                PropertyName = collectionPropertyName,
                Setter = collectionSetter,
                Registry = registry
            };
            
            collections.Add(compiledCollection);
        }

        var compiledMapping = new CompiledMapping<T>
        {
            WorksheetName = configuration.WorksheetName ?? "Sheet1",
            Properties = properties,
            Collections = collections
        };
        
        OptimizeMapping(compiledMapping);
        return compiledMapping;
    }
    
    private static string GetPropertyName(LambdaExpression expression)
    {
        return expression.Body switch
        {
            MemberExpression memberExpr => memberExpr.Member.Name,
            UnaryExpression { Operand: MemberExpression unaryMemberExpr } => unaryMemberExpr.Member.Name,
            _ => throw new InvalidOperationException($"Cannot extract property name from expression: {expression}")
        };
    }
    
    /// <summary>
    /// Optimizes a compiled mapping for runtime performance by pre-calculating cell positions
    /// and building optimized data structures for fast lookup and processing.
    /// </summary>
    private static void OptimizeMapping<T>(CompiledMapping<T>? mapping)
    {
        if (mapping is null)
            throw new ArgumentNullException(nameof(mapping));

        // If already optimized, skip
        if (mapping is { OptimizedCellGrid: not null, OptimizedBoundaries: not null })
            return;

        // Step 1: Calculate mapping boundaries
        var boundaries = CalculateMappingBoundaries(mapping);
        mapping.OptimizedBoundaries = boundaries;

        // Step 3: Build the optimized cell grid
        var cellGrid = BuildOptimizedCellGrid(mapping, boundaries);
        mapping.OptimizedCellGrid = cellGrid;

        // Step 4: Build optimized column handlers for reading
        var columnHandlers = BuildOptimizedColumnHandlers(mapping, boundaries);
        mapping.OptimizedColumnHandlers = columnHandlers;
        
        // Step 5: Pre-compile collection factories and finalizers
        PreCompileCollectionHelpers(mapping);
    }

    private static OptimizedMappingBoundaries CalculateMappingBoundaries<T>(CompiledMapping<T> mapping)
    {
        var boundaries = new OptimizedMappingBoundaries();

        // Process simple properties
        foreach (var prop in mapping.Properties)
        {
            UpdateBoundaries(boundaries, prop.CellColumn, prop.CellRow);
        }

        // Process collections - calculate their potential extent
        foreach (var coll in mapping.Collections)
        {
            var (minRow, maxRow, minCol, maxCol) = CalculateCollectionBounds(coll);
            
            UpdateBoundaries(boundaries, minCol, minRow);
            UpdateBoundaries(boundaries, maxCol, maxRow);
            
            boundaries.HasDynamicCollections = true; // Collections can expand dynamically
        }

        // Set reasonable defaults if no mappings found
        if (boundaries.MinRow == int.MaxValue)
        {
            boundaries.MinRow = 1;
            boundaries.MaxRow = 1;
            boundaries.MinColumn = 1;  
            boundaries.MaxColumn = 1;
        }

        // Detect multiple item pattern
        // NOTE: Multi-item pattern should only be detected when we have simple collections
        // that belong directly to the root item. Nested collections (like Departments in a Company)
        // should NOT trigger multi-item pattern detection.
        // For now, we'll be conservative and only enable multi-item pattern for specific scenarios
        if (mapping is { Collections.Count: > 0, Properties.Count: > 0 })
        {
            // Check if any collection has nested mapping (complex types)
            bool hasNestedCollections = false;
            foreach (var coll in mapping.Collections)
            {
                // Check if the collection's item type has a mapping (complex type)
                if (coll is { ItemType: not null, Registry: not null})
                {
                    // Try to get the nested mapping - if it exists, it's a complex type
                    var nestedMapping = coll.Registry.GetCompiledMapping(coll.ItemType);
                    if (nestedMapping is not null && 
                        coll.ItemType != typeof(string) && 
                        coll.ItemType is { IsValueType: false, IsPrimitive: false }) 
                    {
                        hasNestedCollections = true;
                        break;
                    }
                }
            }
            
            // Only enable multi-item pattern for simple collections (not nested)
            // This is a heuristic - nested collections typically mean a single root item
            // with complex child items, not multiple root items
            if (!hasNestedCollections)
            {
                // Calculate pattern height for multiple items with collections
                var firstPropRow = mapping.Properties.Min(p => p.CellRow);
                
                // Find the actual last row of mapped elements (not the conservative bounds)
                var lastMappedRow = firstPropRow;
                
                // Check actual collection start positions
                foreach (var coll in mapping.Collections)
                {
                    // For vertical collections, we need a reasonable estimate
                    // Use startRow + a small number of items (not the full 100 conservative limit)
                    var estimatedEndRow = coll.StartCellRow + MinItemsForPatternCalc;
                    lastMappedRow = Math.Max(lastMappedRow, estimatedEndRow);
                }
                
                // The pattern height is the total height needed for one complete item
                // including its properties and collections
                boundaries.PatternHeight = lastMappedRow - firstPropRow + 1;
                
                // If we have a reasonable pattern height, mark this as a multi-item pattern
                // This allows the grid to repeat for multiple items
                if (boundaries.PatternHeight is > 0 and < MaxPatternHeight)
                {
                    boundaries.IsMultiItemPattern = true;
                }
            }
        }

        return boundaries;
    }

    private static void UpdateBoundaries(OptimizedMappingBoundaries boundaries, int column, int row)
    {
        if (row < boundaries.MinRow) boundaries.MinRow = row;
        if (row > boundaries.MaxRow) boundaries.MaxRow = row;
        if (column < boundaries.MinColumn) boundaries.MinColumn = column;
        if (column > boundaries.MaxColumn) boundaries.MaxColumn = column;
    }

    private static (int minRow, int maxRow, int minCol, int maxCol) CalculateCollectionBounds(CompiledCollectionMapping collection)
    {
        var startRow = collection.StartCellRow;
        var startCol = collection.StartCellColumn;
        
        // Calculate bounds based on layout
        switch (collection.Layout)
        {
            case CollectionLayout.Vertical:
                // Vertical collections: grow downward
                // Use conservative estimate for initial bounds  
                var verticalHeight = DefaultCollectionHeight;
                
                // Check if this is a complex type with nested mapping
                var maxCol = startCol;
                if (collection.ItemType is null || collection.Registry is null)
                    return (startRow, startRow + verticalHeight, startCol, maxCol);
                
                var nestedMapping = collection.Registry.GetCompiledMapping(collection.ItemType);
                if (nestedMapping is null || !MappingMetadataExtractor.IsComplexType(collection.ItemType)) 
                    return (startRow, startRow + verticalHeight, startCol, maxCol);
                
                // Extract nested mapping info to get max column
                var nestedInfo = MappingMetadataExtractor.ExtractNestedMappingInfo(nestedMapping, collection.ItemType);
                if (nestedInfo is { Properties.Count: > 0 })
                {
                    maxCol = GetMaxColumnIndex(nestedInfo, maxCol);
                }

                return (startRow, startRow + verticalHeight, startCol, maxCol);
        }

        // Default fallback
        return (startRow, startRow + DefaultGridSize, startCol, startCol + DefaultGridSize);
    }

    private static OptimizedCellHandler[,] BuildOptimizedCellGrid<T>(CompiledMapping<T> mapping, OptimizedMappingBoundaries boundaries)
    {
        var height = boundaries.GridHeight;
        var width = boundaries.GridWidth;
        
        var grid = new OptimizedCellHandler[height, width];

        // Initialize all cells as empty
        for (int r = 0; r < height; r++)
        {
            for (int c = 0; c < width; c++)
            {
                grid[r, c] = new OptimizedCellHandler { Type = CellHandlerType.Empty };
            }
        }

        // Process simple properties
        foreach (var prop in mapping.Properties)
        {
            var relativeRow = prop.CellRow - boundaries.MinRow;
            var relativeCol = prop.CellColumn - boundaries.MinColumn;

            if (relativeRow >= 0 && relativeRow < height && relativeCol >= 0 && relativeCol < width)
            {
                grid[relativeRow, relativeCol] = new OptimizedCellHandler
                {
                    Type = string.IsNullOrEmpty(prop.Formula) ? CellHandlerType.Property : CellHandlerType.Formula,
                    ValueExtractor = CreatePropertyValueExtractor(prop),
                    ValueSetter = prop.Setter,
                    PropertyName = prop.PropertyName,
                    Format = prop.Format,
                    Formula = prop.Formula,
                    ItemIndex = 0,  // Properties belong to the first item in the pattern
                    BoundaryRow = -1,  // Properties don't have boundaries
                    BoundaryColumn = -1
                };
            }
        }

        // Process collections - mark their cell ranges
        // Sort collections by start position to process them in order
        var sortedCollections = mapping.Collections
            .Select((c, i) => new { Collection = c, Index = i })
            .OrderBy(x => x.Collection.StartCellRow)
            .ThenBy(x => x.Collection.StartCellColumn)
            .ToList();
            
        for (int i = 0; i < sortedCollections.Count; i++)
        {
            var item = sortedCollections[i];
            // Find the next collection's start row to use as boundary
            int? nextCollectionStartRow = null;
            if (i + 1 < sortedCollections.Count)
            {
                nextCollectionStartRow = sortedCollections[i + 1].Collection.StartCellRow;
            }
            MarkCollectionCells(grid, item.Collection, item.Index, boundaries, nextCollectionStartRow);
        }

        return grid;
    }

    private static void MarkCollectionCells(OptimizedCellHandler[,] grid, CompiledCollectionMapping collection, 
        int collectionIndex, OptimizedMappingBoundaries boundaries, int? nextCollectionStartRow = null)
    {
        var startRow = collection.StartCellRow;
        var startCol = collection.StartCellColumn;
        
        // Mark collection cells based on layout
        // Only support vertical collections
        if (collection.Layout == CollectionLayout.Vertical)
        {
            // Mark vertical range - we'll handle dynamic expansion during runtime
            MarkVerticalCollectionCells(grid, collection, collectionIndex, boundaries, startRow, startCol, nextCollectionStartRow);
        }
    }


    private static void MarkVerticalCollectionCells(OptimizedCellHandler[,] grid, CompiledCollectionMapping collection,
        int collectionIndex, OptimizedMappingBoundaries boundaries, int startRow, int startCol, int? nextCollectionStartRow = null)
    {
        var relativeCol = startCol - boundaries.MinColumn;
        if (relativeCol < 0 || relativeCol >= grid.GetLength(1)) 
            return;

        // Check if the collection's item type has a mapping (complex type)
        var itemType = collection.ItemType ?? typeof(object);
        var nestedMapping = collection.Registry?.GetCompiledMapping(itemType);
        
        if (nestedMapping is not null && MappingMetadataExtractor.IsComplexType(itemType))
        {
            // Complex type with mapping - expand each item across multiple columns
            MarkVerticalComplexCollectionCells(grid, collection, collectionIndex, boundaries, startRow, nestedMapping, nextCollectionStartRow);
        }
        else
        {
            // Simple type - single column
            var maxRows = Math.Min(DefaultCollectionHeight, grid.GetLength(0));
            var startRelativeRow = startRow - boundaries.MinRow;

            // Pre-compile the item converter for this collection
            var itemConverter = CreatePreCompiledItemConverter(itemType);

            for (int r = startRelativeRow; r >= 0 && r < maxRows && r < grid.GetLength(0); r++)
            {
                // Skip rows with spacing
                var itemIndex = (r - startRelativeRow) / (1 + collection.RowSpacing);
                var isDataRow = (r - startRelativeRow) % (1 + collection.RowSpacing) == 0;
                
                if (isDataRow && grid[r, relativeCol].Type == CellHandlerType.Empty)
                {
                    grid[r, relativeCol] = new OptimizedCellHandler
                    {
                        Type = CellHandlerType.CollectionItem,
                        ValueExtractor = CreateCollectionValueExtractor(collection, itemIndex),
                        CollectionIndex = collectionIndex,
                        CollectionItemOffset = itemIndex,
                        CollectionMapping = collection,
                        CollectionItemConverter = itemConverter,
                        ItemIndex = 0,  // Collections belong to the first item in the pattern
                        BoundaryRow = boundaries.IsMultiItemPattern ? boundaries.MinRow + boundaries.PatternHeight : -1,
                        BoundaryColumn = -1  // Vertical collections don't have column boundaries
                    };
                }
            }
        }
    }


    private static OptimizedCellHandler[] BuildOptimizedColumnHandlers<T>(CompiledMapping<T> mapping, OptimizedMappingBoundaries boundaries)
    {
        var columnHandlers = new OptimizedCellHandler[boundaries.GridWidth];

        // Initialize all columns as empty
        for (int i = 0; i < columnHandlers.Length; i++)
        {
            columnHandlers[i] = new OptimizedCellHandler { Type = CellHandlerType.Empty };
        }

        // For reading, we primarily care about the first row where headers/properties are typically defined
        // Build column handlers from the first row that has properties
        foreach (var prop in mapping.Properties)
        {
            var relativeCol = prop.CellColumn - boundaries.MinColumn;
            if (relativeCol >= 0 && relativeCol < columnHandlers.Length)
            {
                columnHandlers[relativeCol] = new OptimizedCellHandler
                {
                    Type = CellHandlerType.Property,
                    ValueSetter = prop.Setter,
                    PropertyName = prop.PropertyName
                };
            }
        }

        return columnHandlers;
    }

    private static Func<object, int, object?> CreatePropertyValueExtractor(CompiledPropertyMapping property)
    {
        // The property getter is already compiled, just wrap it to match our signature
        var getter = property.Getter;
        return (obj, _) => getter(obj);
    }

    private static Func<object, int, object?> CreateCollectionValueExtractor(CompiledCollectionMapping collection, int offset)
    {
        var getter = collection.Getter;
        return (obj, _) =>
        {
            var enumerable = getter(obj);
            return CollectionAccessor.GetItemAt(enumerable, offset);
        };
    }
    
    private static void PreCompileCollectionHelpers<T>(CompiledMapping<T> mapping)
    {
        if (!mapping.Collections.Any()) 
            return;
        
        // Store pre-compiled helpers for each collection
        var helpers = new List<OptimizedCollectionHelper>();
        var nestedMappings = new Dictionary<int, NestedMappingInfo>();
        
        for (int i = 0; i < mapping.Collections.Count; i++)
        {
            var collection = mapping.Collections[i];
            var helper = new OptimizedCollectionHelper();
            
            // Get the actual property info using centralized helper
            var propInfo = MappingMetadataExtractor.GetPropertyByName(typeof(T), collection.PropertyName);
            if (propInfo is null)
                continue;
            
            var propertyType = propInfo.PropertyType;
            var itemType = collection.ItemType ?? typeof(object);
            helper.ItemType = itemType;
            
            // Create simple factory functions using centralized logic
            helper.Factory = () => CollectionAccessor.CreateTypedList(itemType);
            helper.DefaultItemFactory = CollectionAccessor.CreateItemFactory(itemType);
            helper.Finalizer = list => CollectionAccessor.FinalizeCollection(list, propertyType, itemType);
            helper.IsArray = propertyType.IsArray;
            helper.Setter = collection.Setter;
            
            // Pre-compute type metadata to avoid runtime reflection
            helper.IsItemValueType = itemType.IsValueType;
            helper.IsItemPrimitive = itemType.IsPrimitive;
            helper.DefaultValue = itemType.IsValueType ? helper.DefaultItemFactory() : null;
            
            helpers.Add(helper);
            
            // Pre-compile nested mapping info if it's a complex type
            if (collection.Registry is not null && MappingMetadataExtractor.IsComplexType(itemType))
            {
                var nestedMapping = collection.Registry.GetCompiledMapping(itemType);
                if (nestedMapping is not null)
                {
                    var nestedInfo = MappingMetadataExtractor.ExtractNestedMappingInfo(nestedMapping, itemType);
                    if (nestedInfo is not null)
                    {
                        nestedMappings[i] = nestedInfo;
                    }
                }
            }
        }
        
        mapping.OptimizedCollectionHelpers = helpers;
        if (nestedMappings.Count > 0)
        {
            mapping.NestedMappings = nestedMappings;
        }
    }
    
    private static void MarkVerticalComplexCollectionCells(OptimizedCellHandler[,] grid, CompiledCollectionMapping collection,
        int collectionIndex, OptimizedMappingBoundaries boundaries, int startRow, object nestedMapping, int? nextCollectionStartRow = null)
    {
        // Extract pre-compiled nested mapping info without reflection
        var nestedInfo = MappingMetadataExtractor.ExtractNestedMappingInfo(nestedMapping, collection.ItemType ?? typeof(object));
        if (nestedInfo is null) 
            return;
        
        // Now mark cells for each property of each collection item
        var maxRows = Math.Min(100, grid.GetLength(0)); // Conservative range
        var startRelativeRow = startRow - boundaries.MinRow;
        var rowSpacing = collection.RowSpacing;
        
        // Calculate the maximum number of items we can mark
        var maxItems = 20; // Conservative default
        if (nextCollectionStartRow.HasValue)
        {
            // Limit to the rows before the next collection starts
            var availableRows = nextCollectionStartRow.Value - startRow;
            maxItems = Math.Min(maxItems, Math.Max(0, availableRows / (1 + rowSpacing)));
        }
        
        for (int itemIndex = 0; itemIndex < maxItems; itemIndex++)
        {
            var r = startRelativeRow + itemIndex * (1 + rowSpacing);
            if (r < 0 || r >= maxRows || r >= grid.GetLength(0))
                continue;
            
            // Additional check: don't go past the next collection's start
            var absoluteRow = r + boundaries.MinRow;
            if (absoluteRow >= nextCollectionStartRow)
                break;
            
            foreach (var prop in nestedInfo.Properties)
            {
                var c = prop.ColumnIndex - boundaries.MinColumn;
                if (c >= 0 && c < grid.GetLength(1))
                {
                    if (prop.Setter is null)
                        throw new InvalidOperationException($"Nested property '{prop.PropertyName}' is missing a setter. Ensure the mapping for '{collection.ItemType?.Name}' is configured correctly.");
                    // Only mark if not already occupied
                    if (grid[r, c].Type == CellHandlerType.Empty)
                    {
                        grid[r, c] = new OptimizedCellHandler
                        {
                            Type = CellHandlerType.CollectionItem,
                            ValueExtractor = CreateNestedPropertyExtractor(collection, itemIndex, prop.Getter),
                            ValueSetter = prop.Setter,
                            CollectionIndex = collectionIndex,
                            CollectionItemOffset = itemIndex,
                            PropertyName = prop.PropertyName,
                            CollectionMapping = collection,
                            CollectionItemConverter = null // No conversion needed, property getter handles it
                        };
                    }
                }
            }

            if (nestedInfo.Collections.Count <= 0) 
                continue;
            
            foreach (var nestedCollection in nestedInfo.Collections.Values)
            {
                if (nestedCollection.Layout != CollectionLayout.Vertical)
                    continue;

                var nestedMappingInfo = nestedCollection.NestedMapping;
                if (nestedMappingInfo is null || nestedMappingInfo.Properties.Count == 0)
                    continue;

                var nestedMaxItems = 20;
                for (int nestedIndex = 0; nestedIndex < nestedMaxItems; nestedIndex++)
                {
                    var nestedAbsoluteRow = nestedCollection.StartRow + nestedIndex * (1 + nestedCollection.RowSpacing);
                    // Offset by the parent item index so nested items follow the parent row pattern
                    nestedAbsoluteRow += itemIndex * (1 + rowSpacing);
                    if (nextCollectionStartRow.HasValue && nestedAbsoluteRow >= nextCollectionStartRow.Value)
                    {
                        break;
                    }

                    var nestedRelativeRow = nestedAbsoluteRow - boundaries.MinRow;
                    if (nestedRelativeRow < 0 || nestedRelativeRow >= maxRows || nestedRelativeRow >= grid.GetLength(0))
                        continue;

                    foreach (var nestedProp in nestedMappingInfo.Properties)
                    {
                        if (nestedProp.Setter is null)
                            throw new InvalidOperationException($"Nested property '{nestedProp.PropertyName}' is missing a setter. Ensure the mapping for '{nestedCollection.ItemType.Name}' is configured correctly.");

                        var columnIndex = nestedProp.ColumnIndex - boundaries.MinColumn;
                        if (columnIndex < 0 || columnIndex >= grid.GetLength(1))
                            continue;

                        if (grid[nestedRelativeRow, columnIndex].Type != CellHandlerType.Empty)
                        {
                            continue;
                        }

                        grid[nestedRelativeRow, columnIndex] = new OptimizedCellHandler
                        {
                            Type = CellHandlerType.CollectionItem,
                            ValueExtractor = CreateNestedCollectionPropertyExtractor(collection, itemIndex, nestedCollection, nestedIndex, nestedProp.Getter),
                            ValueSetter = CreateNestedCollectionPropertySetter(nestedCollection, nestedIndex, nestedProp.Setter),
                            CollectionIndex = collectionIndex,
                            CollectionItemOffset = itemIndex,
                            PropertyName = nestedProp.PropertyName,
                            CollectionMapping = collection,
                            CollectionItemConverter = null
                        };
                    }
                }
            }
        }
    }
    
    private static Func<object, int, object?> CreateNestedPropertyExtractor(CompiledCollectionMapping collection, int offset, Func<object, object?> propertyGetter)
    {
        var collectionGetter = collection.Getter;
        return (obj, _) =>
        {
            var enumerable = collectionGetter(obj);
            var item = CollectionAccessor.GetItemAt(enumerable, offset);
            
            return item is not null ? propertyGetter(item) : null;
        };
    }

    private static Func<object, int, object?> CreateNestedCollectionPropertyExtractor(
        CompiledCollectionMapping parentCollection,
        int parentOffset,
        NestedCollectionInfo nestedCollection,
        int nestedOffset,
        Func<object, object?> propertyGetter)
    {
        var parentGetter = parentCollection.Getter;
        return (obj, _) =>
        {
            var parents = parentGetter(obj);
            var parentItem = CollectionAccessor.GetItemAt(parents, parentOffset);
            if (parentItem is null)
                return null;

            var nestedEnumerable = nestedCollection.Getter(parentItem);
            var nestedItem = CollectionAccessor.GetItemAt(nestedEnumerable, nestedOffset);

            return nestedItem is not null ? propertyGetter(nestedItem) : null;
        };
    }

    private static Action<object, object?> CreateNestedCollectionPropertySetter(
        NestedCollectionInfo collectionInfo,
        int nestedOffset,
        Action<object, object?> setter)
    {
        return (parent, value) =>
        {
            if (parent is null)
                return;

            var collection = collectionInfo.Getter(parent);
            IList list;

            if (collection is IList existingList)
            {
                list = existingList;
            }
            else if (collection is IEnumerable enumerable)
            {
                list = collectionInfo.ListFactory();
                foreach (var item in enumerable)
                {
                    list.Add(item);
                }

                if (collectionInfo.Setter is null)
                    throw new InvalidOperationException($"Collection property '{collectionInfo.PropertyName}' must be writable to capture nested values.");

                collectionInfo.Setter(parent, list);
            }
            else
            {
                if (collectionInfo.Setter is null)
                    throw new InvalidOperationException($"Collection property '{collectionInfo.PropertyName}' must be writable to capture nested values.");

                list = collectionInfo.ListFactory();
                collectionInfo.Setter(parent, list);
            }

            while (list.Count <= nestedOffset)
            {
                var newItem = collectionInfo.ItemFactory();
                if (newItem is null)
                    throw new InvalidOperationException($"Collection item factory returned null for type '{collectionInfo.ItemType}'. Ensure it has an accessible parameterless constructor.");

                list.Add(newItem);
            }

            var nestedItem = list[nestedOffset];
            if (nestedItem is null)
            {
                nestedItem = collectionInfo.ItemFactory();
                if (nestedItem is null)
                    throw new InvalidOperationException($"Collection item factory returned null for type '{collectionInfo.ItemType}'. Ensure it has an accessible parameterless constructor.");

                list[nestedOffset] = nestedItem;
            }

            setter(nestedItem, value);
        };
    }
    
    private static Func<object?, object?> CreatePreCompiledItemConverter(Type targetType)
    {
        return value => ConversionHelper.ConvertValue(value, targetType);
    }

    private static int GetMaxColumnIndex(NestedMappingInfo nestedInfo, int currentMax)
    {
        if (nestedInfo.Properties.Count > 0)
        {
            currentMax = Math.Max(currentMax, nestedInfo.Properties.Max(p => p.ColumnIndex));
        }

        if (nestedInfo.Collections.Count > 0)
        {
            foreach (var collectionInfo in nestedInfo.Collections.Values)
            {
                currentMax = Math.Max(currentMax, collectionInfo.StartColumn);
                if (collectionInfo.NestedMapping is not null)
                {
                    currentMax = GetMaxColumnIndex(collectionInfo.NestedMapping, currentMax);
                }
            }
        }

        return currentMax;
    }
}