namespace MiniExcelLib.Core.Mapping;

internal static partial class MappingReader<T> where T : class, new()
{
    [CreateSyncVersion]
    public static async IAsyncEnumerable<T> QueryAsync(Stream stream, CompiledMapping<T> mapping, [EnumeratorCancellation] CancellationToken cancellationToken = default)
    {
        if (stream is null)
            throw new ArgumentNullException(nameof(stream));
        if (mapping is null)
            throw new ArgumentNullException(nameof(mapping));

        await foreach (var item in QueryOptimizedAsync(stream, mapping, cancellationToken).ConfigureAwait(false))
            yield return item;
    }
    
    [CreateSyncVersion]
    private static async IAsyncEnumerable<T> QueryOptimizedAsync(Stream stream, CompiledMapping<T> mapping, [EnumeratorCancellation] CancellationToken cancellationToken = default)
    {
        if (mapping.OptimizedCellGrid is null || mapping.OptimizedBoundaries is null)
            throw new InvalidOperationException("QueryOptimizedAsync requires an optimized mapping");

        var boundaries = mapping.OptimizedBoundaries!;
        var cellGrid = mapping.OptimizedCellGrid!;
        
        // Read the Excel file using OpenXmlReader's direct mapping path
        using var reader = await OpenXmlReader.CreateAsync(stream, new OpenXmlConfiguration
        {
            FillMergedCells = false,
            FastMode = false
        }, cancellationToken).ConfigureAwait(false);
        
        // If we have collections, we need to handle multiple items with collections
        if (mapping.Collections.Any())
        {
            // Check if this is a multi-item pattern
            bool isMultiItemPattern = boundaries is { IsMultiItemPattern: true, PatternHeight: > 0 };
            
            T? currentItem = null;
            Dictionary<int, IList>? currentCollections = null;
            var currentItemIndex = -1;

            await foreach (var mappedRow in reader.QueryMappedAsync(mapping.WorksheetName, cancellationToken).ConfigureAwait(false))
            {
                var currentRowIndex = mappedRow.RowIndex + 1;

                // Use our own row counter since OpenXmlReader doesn't provide row numbers
                int rowNumber = currentRowIndex;
                if (rowNumber < boundaries.MinRow)
                    continue;
                
                // Calculate which item this row belongs to based on the pattern
                var relativeRow = rowNumber - boundaries.MinRow;
                int itemIndex = 0;
                int gridRow = relativeRow;
                
                if (isMultiItemPattern && boundaries.PatternHeight > 0)
                {
                    // Pre-calculated: which item does this row belong to?
                    itemIndex = relativeRow / boundaries.PatternHeight;
                    gridRow = relativeRow % boundaries.PatternHeight;
                }
                
                // Check if we're starting a new item
                if (itemIndex != currentItemIndex)
                {
                    // Save the previous item if we have one
                    if (currentItem is not null && currentCollections is not null)
                    {
                        FinalizeCollections(currentItem, mapping, currentCollections);
                        if (HasAnyData(currentItem, mapping))
                        {
                            yield return currentItem;
                        }
                    }
                    
                    // Start the new item
                    currentItem = new T();
                    currentCollections = InitializeCollections(mapping);
                    currentItemIndex = itemIndex;
                }
                
                // If we don't have a current item yet, skip this row
                if (currentItem is null) 
                    continue;
                
                if (gridRow < 0 || gridRow >= cellGrid.GetLength(0)) 
                    continue;
                
                // Process each cell in the row using the pre-calculated grid
                for (int col = boundaries.MinColumn; col <= boundaries.MaxColumn; col++)
                {
                    var cellValue = mappedRow.GetCell(col - 1); // Convert to 0-based for MappedRow
                    
                    if (mapping.TryGetHandler(rowNumber, col, out var handler))
                    {
                        ProcessCellValue(handler, cellValue, currentItem, currentCollections, mapping);
                    }
                }
            }
            
            // Finalize the last item if we have one
            if (currentItem is null || currentCollections is null) 
                yield break;
            
            FinalizeCollections(currentItem, mapping, currentCollections);
            if (HasAnyData(currentItem, mapping))
            {
                yield return currentItem;
            }
        }
        else
        {
            // Check if this is a column layout (properties in same column, different rows)
            // Column layout has GridHeight > 1 and all properties in same column
            bool isColumnLayout = boundaries.GridHeight > 1;
            
            if (isColumnLayout)
            {
                // Column layout mode - all rows form a single object
                var item = new T();

                await foreach (var mappedRow in reader.QueryMappedAsync(mapping.WorksheetName, cancellationToken).ConfigureAwait(false))
                {
                    var currentRowIndex = mappedRow.RowIndex + 1;

                    int rowNumber = currentRowIndex;
                    
                    // Process properties for this row
                    foreach (var prop in mapping.Properties)
                    {
                        if (prop.CellRow == rowNumber)
                        {
                            var cellValue = mappedRow.GetCell(prop.CellColumn - 1); // Convert to 0-based
                            
                            if (cellValue is not null)
                            {
                                // Trust the precompiled setter to handle conversion
                                mapping.TrySetPropertyValue(prop, item, cellValue);
                            }
                        }
                    }
                }
                
                if (HasAnyData(item, mapping))
                {
                    yield return item;
                }
            }
            else
            {
                // Row layout mode - each row is a separate item
                await foreach (var mappedRow in reader.QueryMappedAsync(mapping.WorksheetName, cancellationToken).ConfigureAwait(false))
                {
                    // Use our own row counter since OpenXmlReader doesn't provide row numbers
                    var currentRowIndex = mappedRow.RowIndex + 1;
                    if (currentRowIndex < boundaries.MinRow) 
                        continue;
                    
                    var item = new T();
                    
                    // Process properties for this row
                    // Check if this is a table pattern (all properties on row 1)
                    var allOnRow1 = mapping.Properties.All(p => p.CellRow == 1);
                    
                    foreach (var prop in mapping.Properties)
                    {
                        // For table pattern (all on row 1), properties define columns
                        // For cell-specific mapping, only read from the specific row
                        if (!allOnRow1 && prop.CellRow != currentRowIndex) 
                            continue;
                        
                        var cellValue = mappedRow.GetCell(prop.CellColumn - 1); // Convert to 0-based
                        if (cellValue is not null)
                        {
                            // Trust the precompiled setter to handle conversion
                            prop.Setter?.Invoke(item, cellValue);
                        }
                    }
                    
                    if (HasAnyData(item, mapping))
                    {
                        yield return item;
                    }
                }
            }
        }
    }
    
    private static Dictionary<int, IList> InitializeCollections(CompiledMapping<T> mapping)
    {
        var collections = new Dictionary<int, IList>();
        
        // Use precompiled collection helpers if available
        if (mapping.OptimizedCollectionHelpers is not null)
        {
            for (int i = 0; i < mapping.OptimizedCollectionHelpers.Count && i < mapping.Collections.Count; i++)
            {
                var helper = mapping.OptimizedCollectionHelpers[i];
                collections[i] = helper.Factory();
            }
        }
        else
        {
            // This should never happen with properly optimized mappings
            throw new InvalidOperationException(
                "OptimizedCollectionHelpers is null. Ensure the mapping was properly compiled and optimized.");
        }
        
        return collections;
    }
    
    private static void ProcessCellValue(OptimizedCellHandler handler, object? value, T item, 
        Dictionary<int, IList>? collections, CompiledMapping<T> mapping)
    {
        // Skip empty handlers
        if (handler.Type == CellHandlerType.Empty)
            return;
            
        switch (handler.Type)
        {
            case CellHandlerType.Property:
                // Direct property - use pre-compiled setter
                mapping.TrySetValue(handler, item, value);
                break;
                
            case CellHandlerType.CollectionItem:
                if (handler.CollectionIndex >= 0 
                    && collections is not null 
                    && collections.TryGetValue(handler.CollectionIndex, out var collection))
                {
                    var collectionMapping = handler.CollectionMapping!;
                    var itemType = collectionMapping.ItemType ?? typeof(object);
                    
                    // Check if this is a complex type with nested properties
                    var nestedMapping = collectionMapping.Registry?.GetCompiledMapping(itemType);
                    
                    // Use pre-compiled type metadata from the helper instead of runtime reflection
                    var typeHelper = mapping.OptimizedCollectionHelpers?[handler.CollectionIndex];
                    
                    if (nestedMapping is not null && 
                        itemType != typeof(string) && 
                        typeHelper is { IsItemValueType: false, IsItemPrimitive: false })
                    {
                        // Complex type - we need to build/update the object
                        ProcessComplexCollectionItem(collection, handler, value, mapping);
                    }
                    else
                    {
                        // Simple type - add directly
                        while (collection.Count <= handler.CollectionItemOffset)
                        {
                            // Use precompiled default factory if available
                            object? defaultValue;
                            if (mapping.OptimizedCollectionHelpers is not null && 
                                handler.CollectionIndex >= 0 && 
                                handler.CollectionIndex < mapping.OptimizedCollectionHelpers.Count)
                            {
                                var helper = mapping.OptimizedCollectionHelpers[handler.CollectionIndex];
                                defaultValue = helper.DefaultItemFactory.Invoke();
                            }
                            else
                            {
                                // This should never happen with properly optimized mappings
                                throw new InvalidOperationException(
                                    $"No OptimizedCollectionHelper found for collection at index {handler.CollectionIndex}. " +
                                    "Ensure the mapping was properly compiled and optimized.");
                            }
                            collection.Add(defaultValue);
                        }
                        
                        // Skip empty values for value type collections
                        if (value is string str && string.IsNullOrEmpty(str))
                        {
                            // Don't add empty values to value type collections
                            // Use pre-compiled type metadata from the helper
                            var itemHelper = mapping.OptimizedCollectionHelpers?[handler.CollectionIndex];
                            if (itemHelper is { IsItemValueType: false })
                            {
                                // Only set null if the collection has the item already
                                if (handler.CollectionItemOffset < collection.Count)
                                {
                                    collection[handler.CollectionItemOffset] = null;
                                }
                            }
                            // For value types, we just skip - the default value is already there
                        }
                        else
                        {
                            // Use pre-compiled converter if available
                            var convertedValue = handler.CollectionItemConverter is not null
                                ? handler.CollectionItemConverter(value)
                                : value;

                            collection[handler.CollectionItemOffset] = convertedValue;
                        }
                    }
                }
                break;
        }
    }
    
    private static void ProcessComplexCollectionItem(IList collection, OptimizedCellHandler handler, object? value, CompiledMapping<T> mapping)
    {
        if (collection.Count <= handler.CollectionItemOffset && !HasMeaningfulValue(value))
            return;

        // Ensure the collection has enough items
        while (collection.Count <= handler.CollectionItemOffset)
        {
            // Use precompiled default factory
            if (mapping.OptimizedCollectionHelpers is null || 
                handler.CollectionIndex < 0 || 
                handler.CollectionIndex >= mapping.OptimizedCollectionHelpers.Count)
            {
                throw new InvalidOperationException(
                    $"No OptimizedCollectionHelper found for collection at index {handler.CollectionIndex}. " +
                    "Ensure the mapping was properly compiled and optimized.");
            }
            
            var helper = mapping.OptimizedCollectionHelpers[handler.CollectionIndex];
            var newItem = helper.DefaultItemFactory.Invoke();
            collection.Add(newItem);
        }
        
        var item = collection[handler.CollectionItemOffset];
        if (item is null)
        {
            if (mapping.OptimizedCollectionHelpers is null ||
                handler.CollectionIndex < 0 ||
                handler.CollectionIndex >= mapping.OptimizedCollectionHelpers.Count)
            {
                throw new InvalidOperationException(
                    $"No OptimizedCollectionHelper found for collection at index {handler.CollectionIndex}. " +
                    "Ensure the mapping was properly compiled and optimized.");
            }

            var helper = mapping.OptimizedCollectionHelpers[handler.CollectionIndex];
            item = helper.DefaultItemFactory.Invoke();

            collection[handler.CollectionItemOffset] = item ?? throw new InvalidOperationException(
                $"Collection item factory returned null for type '{helper.ItemType}'. Ensure it has an accessible parameterless constructor.");
        }
        
        // Try to set the value using the handler
        if (!mapping.TrySetValue(handler, item, value))
        {
            // For nested mappings, we need to look up the pre-compiled setter
            if (mapping.NestedMappings?.TryGetValue(handler.CollectionIndex, out var nestedInfo) is true)
            {
                // Find the matching property setter in the nested mapping
                var nestedProp = nestedInfo.Properties.FirstOrDefault(p => p.PropertyName == handler.PropertyName);
                if (nestedProp?.Setter is not null)
                {
                    handler.ValueSetter = nestedProp.Setter;
                    nestedProp.Setter(item, value);
                    return;
                }
            }
            
            throw new InvalidOperationException(
                $"ValueSetter is null for complex collection item handler at property '{handler.PropertyName}'. " +
                "This indicates the mapping was not properly optimized. Ensure the type was mapped in the MappingRegistry.");
        }
    }

    private static bool HasMeaningfulValue(object? value) => value switch
    {
        null => false,
        string str => !string.IsNullOrWhiteSpace(str),
        _ => true
    };

    private static void FinalizeCollections(T item, CompiledMapping<T> mapping, Dictionary<int, IList> collections)
    {
        for (int i = 0; i < mapping.Collections.Count; i++)
        {
            var collectionMapping = mapping.Collections[i];
            if (!collections.TryGetValue(i, out var list)) 
                continue;
            
            // Get the default value using precompiled factory if available
            object? defaultValue = null;
            if (mapping.OptimizedCollectionHelpers is not null && i < mapping.OptimizedCollectionHelpers.Count)
            {
                var helper = mapping.OptimizedCollectionHelpers[i];
                // Use pre-compiled type metadata instead of runtime check
                if (helper.IsItemValueType)
                {
                    defaultValue = helper.DefaultValue ?? helper.DefaultItemFactory.Invoke();
                }
            }
            else
            {
                // This should never happen with properly optimized mappings  
                throw new InvalidOperationException(
                    $"No OptimizedCollectionHelper found for collection at index {i}. " +
                    "Ensure the mapping was properly compiled and optimized.");
            }

            while (list.Count > 0)
            {
                var lastItem = list[^1];
                // Use pre-compiled type metadata from helper
                var listHelper = mapping.OptimizedCollectionHelpers?[i];
                bool isDefault = lastItem is null ||
                                 (lastItem.Equals(defaultValue) && listHelper is { IsItemValueType: true });
                if (isDefault)
                {
                    list.RemoveAt(list.Count - 1);
                }
                else
                {
                    break; // Stop when we find a non-default value
                }
            }

            // Convert to final type if needed
            object finalValue = list;

            if (collectionMapping.Setter is null) 
                continue;

            // Use precompiled collection helper to convert to final type
            if (mapping.OptimizedCollectionHelpers is not null && i < mapping.OptimizedCollectionHelpers.Count)
            {
                var helper = mapping.OptimizedCollectionHelpers[i];
                finalValue = helper.Finalizer(list);
            }

            mapping.TrySetCollectionValue(collectionMapping, item, finalValue);
        }
    }
    
    
    private static bool HasAnyData(T item, CompiledMapping<T> mapping)
    {
        // Check if any properties have non-default values
        var values = mapping.Properties.Select(prop => prop.Getter(item));
        if (values.Any(v => !IsDefaultValue(v)))
        {
            return true;
        }

        // Check if any collections have items
        foreach (var collMap in mapping.Collections)
        {
            var collection = collMap.Getter(item);
            var enumerator = collection.GetEnumerator();
            using var disposableEnumerator = enumerator as IDisposable;
            if (enumerator.MoveNext())
            {
                return true;
            }
        }

        return false;
    }
    
    private static bool IsDefaultValue(object value) => value switch
    {
        string s => string.IsNullOrEmpty(s),
        DateTime dt => dt == default,
        int i => i == 0,
        long l => l == 0L,
        decimal m => m == 0M,
        double d => d == 0D,
        float f => f == 0F,
        bool b => !b,
        _ => false
    };
}