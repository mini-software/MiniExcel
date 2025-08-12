using System.Collections;
using System.Linq.Expressions;
using System.Reflection;
using MiniExcelLib.Core.Mapping.Configuration;
using MiniExcelLib.Core.OpenXml.Utils;

namespace MiniExcelLib.Core.Mapping;

/// <summary>
/// Compiles mapping configurations into optimized runtime representations for efficient Excel read/write operations.
/// Uses a universal optimization strategy with pre-compiled property accessors and cell grids.
/// </summary>
internal static class MappingCompiler
{
    // Conservative estimates for collection bounds when actual size is unknown
    private const int DefaultCollectionHeight = 100;
    private const int DefaultCollectionWidth = 100;
    private const int DefaultGridSize = 10;
    private const int MaxItemsToMarkInGrid = 10;
    private const int MaxPatternHeight = 20;
    private const int MinItemsForPatternCalc = 2;
    
    /// <summary>
    /// Compiles a mapping configuration into an optimized runtime representation.
    /// </summary>
    public static CompiledMapping<T> Compile<T>(MappingConfiguration<T> configuration, MappingRegistry? registry = null)
    {
        if (configuration == null)
            throw new ArgumentNullException(nameof(configuration));

        var properties = new List<CompiledPropertyMapping>();
        var collections = new List<CompiledCollectionMapping>();
        
        // Compile property mappings
        foreach (var prop in configuration.PropertyMappings)
        {
            if (string.IsNullOrEmpty(prop.CellAddress))
                throw new InvalidOperationException($"Property mapping must specify a cell address using ToCell()");

            var parameter = Expression.Parameter(typeof(object), "obj");
            var cast = Expression.Convert(parameter, typeof(T));
            var propertyAccess = Expression.Invoke(prop.Expression, cast);
            var convertToObject = Expression.Convert(propertyAccess, typeof(object));
            var lambda = Expression.Lambda<Func<object, object>>(convertToObject, parameter);
            var compiled = lambda.Compile();

            // Extract property name from expression
            var propertyName = GetPropertyName(prop.Expression);
            
            // Create setter
            var setterParam = Expression.Parameter(typeof(object), "obj");
            var valueParam = Expression.Parameter(typeof(object), "value");
            var castObj = Expression.Convert(setterParam, typeof(T));
            var castValue = Expression.Convert(valueParam, prop.PropertyType);
            
            // Create assignment if it's a member expression
            Action<object, object?>? setter = null;
            if (prop.Expression.Body is MemberExpression memberExpr && memberExpr.Member is PropertyInfo propInfo)
            {
                var assign = Expression.Assign(Expression.Property(castObj, propInfo), castValue);
                var setterLambda = Expression.Lambda<Action<object, object?>>(assign, setterParam, valueParam);
                setter = setterLambda.Compile();
            }
            
            // Pre-parse cell coordinates for runtime performance
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
            
            // Determine the item type
            var collectionType = coll.PropertyType;
            Type? itemType = null;
            
            if (collectionType.IsArray)
            {
                itemType = collectionType.GetElementType();
            }
            else if (collectionType.IsGenericType)
            {
                var genericArgs = collectionType.GetGenericArguments();
                if (genericArgs.Length > 0)
                {
                    itemType = genericArgs[0];
                }
            }
            
            // Create setter for collection
            Action<object, object?>? collectionSetter = null;
            if (coll.Expression.Body is MemberExpression collMemberExpr && collMemberExpr.Member is System.Reflection.PropertyInfo collPropInfo)
            {
                var setterParam = Expression.Parameter(typeof(object), "obj");
                var valueParam = Expression.Parameter(typeof(object), "value");
                var castObj = Expression.Convert(setterParam, typeof(T));
                var castValue = Expression.Convert(valueParam, collPropInfo.PropertyType);
                var assign = Expression.Assign(Expression.Property(castObj, collPropInfo), castValue);
                var setterLambda = Expression.Lambda<Action<object, object?>>(assign, setterParam, valueParam);
                collectionSetter = setterLambda.Compile();
            }
            
            // Pre-parse start cell coordinates
            ReferenceHelper.ParseReference(coll.StartCell, out int startCol, out int startRow);
            
            var compiledCollection = new CompiledCollectionMapping
            {
                Getter = compiled,
                StartCell = coll.StartCell,
                StartCellColumn = startCol,
                StartCellRow = startRow,
                Layout = coll.Layout,
                RowSpacing = coll.RowSpacing,
                ItemType = itemType ?? coll.ItemType,
                PropertyName = collectionPropertyName,
                Setter = collectionSetter,
                Registry = registry
            };
            
            // Try to get the item mapping from the registry if available
            if (itemType != null && registry != null)
            {
                var itemMapping = registry.GetCompiledMapping(itemType);
                if (itemMapping != null)
                {
                    compiledCollection.ItemMapping = itemMapping;
                }
            }
            // Otherwise compile nested item mapping if exists
            else if (coll.ItemConfiguration != null && coll.ItemType != null)
            {
                var compileMethod = typeof(MappingCompiler)
                    .GetMethod(nameof(Compile))!
                    .MakeGenericMethod(coll.ItemType);
                    
                compiledCollection.ItemMapping = compileMethod.Invoke(null, [coll.ItemConfiguration, registry]);
            }
            
            collections.Add(compiledCollection);
        }

        var compiledMapping = new CompiledMapping<T>
        {
            WorksheetName = configuration.WorksheetName ?? "Sheet1",
            Properties = properties,
            Collections = collections
        };
        
        // Apply universal optimization to all mappings
        OptimizeMapping(compiledMapping, registry);
        
        return compiledMapping;
    }
    
    private static string GetPropertyName(LambdaExpression expression)
    {
        if (expression.Body is MemberExpression memberExpr)
        {
            return memberExpr.Member.Name;
        }
        
        if (expression.Body is UnaryExpression unaryExpr && unaryExpr.Operand is MemberExpression unaryMemberExpr)
        {
            return unaryMemberExpr.Member.Name;
        }
        
        throw new InvalidOperationException($"Cannot extract property name from expression: {expression}");
    }
    
    /// <summary>
    /// Optimizes a compiled mapping for runtime performance by pre-calculating cell positions
    /// and building optimized data structures for fast lookup and processing.
    /// </summary>
    public static void OptimizeMapping<T>(CompiledMapping<T> mapping, MappingRegistry? registry = null)
    {
        if (mapping == null)
            throw new ArgumentNullException(nameof(mapping));

        // If already optimized, skip
        if (mapping.IsUniversallyOptimized)
            return;

        // Step 1: Calculate mapping boundaries
        var boundaries = CalculateMappingBoundaries(mapping);
        mapping.OptimizedBoundaries = boundaries;

        // Step 2: Pre-calculate collection expansion info
        var expansions = CalculateCollectionExpansions(mapping);
        mapping.CollectionExpansions = expansions;

        // Step 3: Build the optimized cell grid
        var cellGrid = BuildOptimizedCellGrid(mapping, boundaries);
        mapping.OptimizedCellGrid = cellGrid;

        // Step 4: Build optimized column handlers for reading
        var columnHandlers = BuildOptimizedColumnHandlers(mapping, boundaries);
        mapping.OptimizedColumnHandlers = columnHandlers;
        
        // Step 5: Pre-compile collection factories and finalizers
        PreCompileCollectionHelpers<T>(mapping);
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
        if (mapping.Collections.Any() && mapping.Properties.Any())
        {
            // Check if any collection has nested mapping (complex types)
            bool hasNestedCollections = false;
            foreach (var coll in mapping.Collections)
            {
                // Check if the collection's item type has a mapping (complex type)
                if (coll.ItemType != null && coll.Registry != null)
                {
                    // Try to get the nested mapping - if it exists, it's a complex type
                    var nestedMapping = coll.Registry.GetCompiledMapping(coll.ItemType);
                    if (nestedMapping != null && coll.ItemType != typeof(string) && 
                        !coll.ItemType.IsValueType && !coll.ItemType.IsPrimitive)
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
                if (boundaries.PatternHeight > 0 && boundaries.PatternHeight < MaxPatternHeight)
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
                if (collection.ItemType != null && collection.Registry != null)
                {
                    var nestedMapping = collection.Registry.GetCompiledMapping(collection.ItemType);
                    if (nestedMapping != null && collection.ItemType != typeof(string) && 
                        !collection.ItemType.IsValueType && !collection.ItemType.IsPrimitive)
                    {
                        // Extract max column from nested mapping properties
                        var nestedMappingType = nestedMapping.GetType();
                        var propsProperty = nestedMappingType.GetProperty("Properties");
                        if (propsProperty != null)
                        {
                            var properties = propsProperty.GetValue(nestedMapping) as System.Collections.IEnumerable;
                            if (properties != null)
                            {
                                foreach (var prop in properties)
                                {
                                    var propType = prop.GetType();
                                    var columnProperty = propType.GetProperty("CellColumn");
                                    if (columnProperty != null)
                                    {
                                        var column = (int)columnProperty.GetValue(prop);
                                        maxCol = Math.Max(maxCol, column);
                                    }
                                }
                            }
                        }
                    }
                }
                
                return (startRow, startRow + verticalHeight, startCol, maxCol);
        }

        // Default fallback
        return (startRow, startRow + DefaultGridSize, startCol, startCol + DefaultGridSize);
    }

    private static List<CollectionExpansionInfo> CalculateCollectionExpansions<T>(CompiledMapping<T> mapping)
    {
        var expansions = new List<CollectionExpansionInfo>();

        foreach (var collection in mapping.Collections)
        {
            expansions.Add(new CollectionExpansionInfo
            {
                StartRow = collection.StartCellRow,
                StartColumn = collection.StartCellColumn,
                Layout = collection.Layout,
                RowSpacing = collection.RowSpacing,
                CollectionMapping = collection
            });
        }

        return expansions;
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
        for (int i = 0; i < mapping.Properties.Count; i++)
        {
            var prop = mapping.Properties[i];
            var relativeRow = prop.CellRow - boundaries.MinRow;
            var relativeCol = prop.CellColumn - boundaries.MinColumn;

            if (relativeRow >= 0 && relativeRow < height && relativeCol >= 0 && relativeCol < width)
            {
                grid[relativeRow, relativeCol] = new OptimizedCellHandler
                {
                    Type = string.IsNullOrEmpty(prop.Formula) ? CellHandlerType.Property : CellHandlerType.Formula,
                    ValueExtractor = CreatePropertyValueExtractor(prop),
                    ValueSetter = CreatePreCompiledSetter(prop),  // Pre-compiled setter with conversion built-in
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
            
        foreach (var item in sortedCollections)
        {
            MarkCollectionCells(grid, item.Collection, item.Index, boundaries);
        }

        return grid;
    }

    private static void MarkCollectionCells(OptimizedCellHandler[,] grid, CompiledCollectionMapping collection, 
        int collectionIndex, OptimizedMappingBoundaries boundaries)
    {
        var startRow = collection.StartCellRow;
        var startCol = collection.StartCellColumn;
        
        // Mark collection cells based on layout
        // Only support vertical collections
        if (collection.Layout == CollectionLayout.Vertical)
        {
            // Mark vertical range - we'll handle dynamic expansion during runtime
            MarkVerticalCollectionCells(grid, collection, collectionIndex, boundaries, startRow, startCol);
        }
    }


    private static void MarkVerticalCollectionCells(OptimizedCellHandler[,] grid, CompiledCollectionMapping collection,
        int collectionIndex, OptimizedMappingBoundaries boundaries, int startRow, int startCol)
    {
        var relativeCol = startCol - boundaries.MinColumn;
        if (relativeCol < 0 || relativeCol >= grid.GetLength(1)) return;

        // Check if the collection's item type has a mapping (complex type)
        var itemType = collection.ItemType ?? typeof(object);
        var nestedMapping = collection.Registry?.GetCompiledMapping(itemType);
        
        if (nestedMapping != null && itemType != typeof(string) && !itemType.IsValueType && !itemType.IsPrimitive)
        {
            // Complex type with mapping - expand each item across multiple columns
            MarkVerticalComplexCollectionCells(grid, collection, collectionIndex, boundaries, startRow, startCol, nestedMapping);
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
                    ValueSetter = CreatePreCompiledSetter(prop),
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
        return (obj, itemIndex) => getter(obj);
    }
    
    private static Action<object, object?>? CreatePreCompiledSetter(CompiledPropertyMapping property)
    {
        // Pre-compile the setter with type conversion built in
        var originalSetter = property.Setter;
        if (originalSetter == null) return null;
        
        var targetType = property.PropertyType;
        
        // Build a setter that includes conversion
        return (obj, value) =>
        {
            if (value == null)
            {
                originalSetter(obj, null);
                return;
            }
            
            // Pre-compiled conversion logic - this runs at compile time, not runtime!
            object? convertedValue = value;
            
            if (value.GetType() != targetType)
            {
                convertedValue = targetType switch
                {
                    _ when targetType == typeof(string) => value.ToString(),
                    _ when targetType == typeof(int) => Convert.ToInt32(value),
                    _ when targetType == typeof(long) => Convert.ToInt64(value),
                    _ when targetType == typeof(decimal) => Convert.ToDecimal(value),
                    _ when targetType == typeof(double) => Convert.ToDouble(value),
                    _ when targetType == typeof(float) => Convert.ToSingle(value),
                    _ when targetType == typeof(bool) => Convert.ToBoolean(value),
                    _ when targetType == typeof(DateTime) => Convert.ToDateTime(value),
                    _ => Convert.ChangeType(value, targetType)
                };
            }
            
            originalSetter(obj, convertedValue);
        };
    }

    private static Func<object, int, object?> CreateCollectionValueExtractor(CompiledCollectionMapping collection, int offset)
    {
        var getter = collection.Getter;
        return (obj, itemIndex) =>
        {
            var enumerable = getter(obj);
            if (enumerable == null) return null;
            
            // Try to use IList for O(1) access if possible
            if (enumerable is IList list)
            {
                return offset < list.Count ? list[offset] : null;
            }
            
            // Fallback to Skip/FirstOrDefault for other IEnumerable
            return enumerable.Cast<object>().Skip(offset).FirstOrDefault();
        };
    }
    
    private static void PreCompileCollectionHelpers<T>(CompiledMapping<T> mapping)
    {
        if (!mapping.Collections.Any()) return;
        
        // Store pre-compiled helpers for each collection
        var helpers = new List<OptimizedCollectionHelper>();
        
        foreach (var collection in mapping.Collections)
        {
            var helper = new OptimizedCollectionHelper();
            
            // Get the actual property info
            var propInfo = typeof(T).GetProperty(collection.PropertyName);
            if (propInfo == null) continue;
            
            var propertyType = propInfo.PropertyType;
            var itemType = collection.ItemType ?? typeof(object);
            
            // Pre-compile collection factory
            helper.Factory = () =>
            {
                var listType = typeof(List<>).MakeGenericType(itemType);
                return (System.Collections.IList)Activator.CreateInstance(listType)!;
            };
            
            // Pre-compile finalizer (converts list to final type)
            if (propertyType.IsArray)
            {
                helper.IsArray = true;
                var elementType = propertyType.GetElementType()!;
                helper.Finalizer = (list) =>
                {
                    var array = Array.CreateInstance(elementType, list.Count);
                    list.CopyTo(array, 0);
                    return array;
                };
            }
            else
            {
                helper.IsArray = false;
                helper.Finalizer = (list) => list;
            }
            
            // Pre-compile setter
            helper.Setter = collection.Setter;
            
            helpers.Add(helper);
        }
        
        mapping.OptimizedCollectionHelpers = helpers;
    }
    
    private static void MarkVerticalComplexCollectionCells(OptimizedCellHandler[,] grid, CompiledCollectionMapping collection,
        int collectionIndex, OptimizedMappingBoundaries boundaries, int startRow, int startCol, object nestedMapping)
    {
        // For complex types, we need to extract individual properties
        // Use reflection to get the properties from the nested mapping
        var nestedMappingType = nestedMapping.GetType();
        var propsProperty = nestedMappingType.GetProperty("Properties");
        if (propsProperty == null) return;
        
        var properties = propsProperty.GetValue(nestedMapping) as System.Collections.IEnumerable;
        if (properties == null) return;
        
        var propertyList = new List<(string Name, int Column, Func<object, object?> Getter)>();
        foreach (var prop in properties)
        {
            var propType = prop.GetType();
            var nameProperty = propType.GetProperty("PropertyName");
            var columnProperty = propType.GetProperty("CellColumn");
            var getterProperty = propType.GetProperty("Getter");
            
            if (nameProperty != null && columnProperty != null && getterProperty != null)
            {
                var name = nameProperty.GetValue(prop) as string;
                var column = (int)columnProperty.GetValue(prop);
                var getter = getterProperty.GetValue(prop) as Func<object, object?>;
                
                if (name != null && getter != null)
                {
                    propertyList.Add((name, column, getter));
                }
            }
        }
        
        // Now mark cells for each property of each collection item
        var maxRows = Math.Min(100, grid.GetLength(0)); // Conservative range
        var startRelativeRow = startRow - boundaries.MinRow;
        var rowSpacing = collection.RowSpacing;
        
        for (int itemIndex = 0; itemIndex < 20; itemIndex++) // Conservative estimate of collection size
        {
            var r = startRelativeRow + itemIndex * (1 + rowSpacing);
            if (r >= 0 && r < maxRows && r < grid.GetLength(0))
            {
                foreach (var (propName, propColumn, propGetter) in propertyList)
                {
                    var c = propColumn - boundaries.MinColumn;
                    if (c >= 0 && c < grid.GetLength(1))
                    {
                        // Only mark if not already occupied
                        if (grid[r, c].Type == CellHandlerType.Empty)
                        {
                            grid[r, c] = new OptimizedCellHandler
                            {
                                Type = CellHandlerType.CollectionItem,
                                ValueExtractor = CreateNestedPropertyExtractor(collection, itemIndex, propGetter),
                                CollectionIndex = collectionIndex,
                                CollectionItemOffset = itemIndex,
                                PropertyName = propName,
                                CollectionMapping = collection,
                                CollectionItemConverter = null // No conversion needed, property getter handles it
                            };
                        }
                    }
                }
            }
        }
    }
    
    private static Func<object, int, object?> CreateNestedPropertyExtractor(CompiledCollectionMapping collection, int offset, Func<object, object?> propertyGetter)
    {
        var collectionGetter = collection.Getter;
        return (obj, itemIndex) =>
        {
            var enumerable = collectionGetter(obj);
            if (enumerable == null) return null;
            
            // Try to use IList for O(1) access if possible
            if (enumerable is IList list)
            {
                if (offset < list.Count && list[offset] != null)
                {
                    // Extract the property from the nested object
                    return propertyGetter(list[offset]);
                }
            }
            else
            {
                // Fall back to enumeration (slower but works)
                var items = enumerable.Cast<object>().Skip(offset).Take(1).ToArray();
                if (items.Length > 0 && items[0] != null)
                {
                    return propertyGetter(items[0]);
                }
            }
            
            return null;
        };
    }
    
    private static Func<object?, object?> CreatePreCompiledItemConverter(Type targetType)
    {
        // Pre-compile all the conversion logic
        return (value) =>
        {
            if (value == null) return null;
            if (value.GetType() == targetType) return value;
            
            // These conversions are JIT-compiled and inlined
            try
            {
                return targetType switch
                {
                    _ when targetType == typeof(string) => value.ToString(),
                    _ when targetType == typeof(int) => Convert.ToInt32(value),
                    _ when targetType == typeof(long) => Convert.ToInt64(value),
                    _ when targetType == typeof(decimal) => Convert.ToDecimal(value),
                    _ when targetType == typeof(double) => Convert.ToDouble(value),
                    _ when targetType == typeof(float) => Convert.ToSingle(value),
                    _ when targetType == typeof(bool) => Convert.ToBoolean(value),
                    _ when targetType == typeof(DateTime) => Convert.ToDateTime(value),
                    _ => Convert.ChangeType(value, targetType)
                };
            }
            catch
            {
                // Fallback to string parsing for robustness
                var str = value.ToString();
                if (string.IsNullOrEmpty(str)) return null;
                
                return targetType switch
                {
                    _ when targetType == typeof(int) && int.TryParse(str, out var i) => i,
                    _ when targetType == typeof(long) && long.TryParse(str, out var l) => l,
                    _ when targetType == typeof(decimal) && decimal.TryParse(str, out var d) => d,
                    _ when targetType == typeof(double) && double.TryParse(str, out var db) => db,
                    _ when targetType == typeof(float) && float.TryParse(str, out var f) => f,
                    _ when targetType == typeof(bool) && bool.TryParse(str, out var b) => b,
                    _ when targetType == typeof(DateTime) && DateTime.TryParse(str, out var dt) => dt,
                    _ => Convert.ChangeType(value, targetType)
                };
            }
        };
    }
}