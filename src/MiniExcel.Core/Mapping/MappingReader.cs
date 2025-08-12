namespace MiniExcelLib.Core.Mapping;

internal static partial class MappingReader<T> where T : class, new()
{
    [CreateSyncVersion]
    public static async IAsyncEnumerable<T> QueryAsync(Stream stream, CompiledMapping<T> mapping, [EnumeratorCancellation] CancellationToken cancellationToken = default)
    {
        if (stream == null)
            throw new ArgumentNullException(nameof(stream));
        if (mapping == null)
            throw new ArgumentNullException(nameof(mapping));

        await foreach (var item in QueryOptimizedAsync(stream, mapping, cancellationToken).ConfigureAwait(false))
            yield return item;
    }
    
    [CreateSyncVersion]
    private static async IAsyncEnumerable<T> QueryOptimizedAsync(Stream stream, CompiledMapping<T> mapping, [EnumeratorCancellation] CancellationToken cancellationToken = default)
    {
        if (mapping.OptimizedCellGrid == null || mapping.OptimizedBoundaries == null)
            throw new InvalidOperationException("QueryOptimizedAsync requires an optimized mapping");

        var boundaries = mapping.OptimizedBoundaries!;
        var cellGrid = mapping.OptimizedCellGrid!;
        
        // Read the Excel file using OpenXmlReader
        using var reader = await OpenXmlReader.CreateAsync(stream, new OpenXmlConfiguration
        {
            FillMergedCells = false,
            FastMode = true
        }, cancellationToken).ConfigureAwait(false);
        
        // If we have collections, we need to handle multiple items with collections
        if (mapping.Collections.Any())
        {
            // Check if this is a multi-item pattern
            bool isMultiItemPattern = boundaries.IsMultiItemPattern && boundaries.PatternHeight > 0;
            
            T? currentItem = null;
            Dictionary<int, IList>? currentCollections = null;
            int currentItemIndex = -1;
            
            int currentRowIndex = 0;
            await foreach (var row in reader.QueryAsync(false, mapping.WorksheetName, "A1", cancellationToken).ConfigureAwait(false))
            {
                currentRowIndex++;
                if (row == null) continue;
                var rowDict = row as IDictionary<string, object>;


                // Use our own row counter since OpenXmlReader doesn't provide row numbers
                int rowNumber = currentRowIndex;
                if (rowNumber < boundaries.MinRow) continue;
                
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
                    if (currentItem != null && currentCollections != null)
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
                if (currentItem == null) continue;
                
                if (gridRow < 0 || gridRow >= cellGrid.GetLength(0)) continue;
                
                // Process each cell in the row using the pre-calculated grid
                foreach (var kvp in rowDict)
                {
                    if (kvp.Key.StartsWith("__")) continue; // Skip metadata keys
                    
                    // Convert column letter to index
                    if (!TryParseColumnIndex(kvp.Key, out int columnIndex))
                        continue;
                    
                    var relativeCol = columnIndex - boundaries.MinColumn;
                    if (relativeCol < 0 || relativeCol >= cellGrid.GetLength(1))
                        continue;
                    
                    var handler = cellGrid[gridRow, relativeCol];
                    ProcessCellValue(handler, kvp.Value, currentItem, currentCollections, mapping);
                }
            }
            
            // Finalize the last item if we have one
            if (currentItem != null && currentCollections != null)
            {
                FinalizeCollections(currentItem, mapping, currentCollections);
                
                if (HasAnyData(currentItem, mapping))
                {
                    yield return currentItem;
                }
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
                int currentRowIndex = 0;
                
                await foreach (var row in reader.QueryAsync(false, mapping.WorksheetName, "A1", cancellationToken).ConfigureAwait(false))
                {
                    currentRowIndex++;
                    if (row == null) continue;
                    var rowDict = row as IDictionary<string, object>;

                    int rowNumber = currentRowIndex;
                    
                    // Process properties for this row
                    foreach (var prop in mapping.Properties)
                    {
                        if (prop.CellRow == rowNumber)
                        {
                            var columnLetter = ReferenceHelper.GetCellLetter(
                                ReferenceHelper.ConvertCoordinatesToCell(prop.CellColumn, 1));
                            
                            if (rowDict.TryGetValue(columnLetter, out var value) && value != null)
                            {
                                // Trust the precompiled setter to handle conversion
                                prop.Setter?.Invoke(item, value);
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
                int currentRowIndex = 0;
                await foreach (var row in reader.QueryAsync(false, mapping.WorksheetName, "A1", cancellationToken).ConfigureAwait(false))
                {
                    currentRowIndex++;
                    if (row == null) continue;
                    var rowDict = row as IDictionary<string, object>;

                    // Use our own row counter since OpenXmlReader doesn't provide row numbers
                    int rowNumber = currentRowIndex;
                    if (rowNumber < boundaries.MinRow) continue;
                    
                    var item = new T();
                    
                    // Process properties for this row
                    // Check if this is a table pattern (all properties on row 1)
                    var allOnRow1 = mapping.Properties.All(p => p.CellRow == 1);
                    
                    foreach (var prop in mapping.Properties)
                    {
                        // For table pattern (all on row 1), properties define columns
                        // For cell-specific mapping, only read from the specific row
                        if (!allOnRow1 && prop.CellRow != rowNumber) continue;
                        
                        var columnLetter = ReferenceHelper.GetCellLetter(
                            ReferenceHelper.ConvertCoordinatesToCell(prop.CellColumn, 1));

                        if (!rowDict.TryGetValue(columnLetter, out var value) || value == null) continue;
                        
                        // Trust the precompiled setter to handle conversion
                        if (prop.Setter == null) continue;
                        prop.Setter.Invoke(item, value);
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
        if (mapping.OptimizedCollectionHelpers != null)
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
    
    private static void ProcessCellValue(OptimizedCellHandler handler, object value, T item, 
        Dictionary<int, IList>? collections, CompiledMapping<T> mapping)
    {
        switch (handler.Type)
        {
            case CellHandlerType.Property:
                // Direct property - use pre-compiled setter
                handler.ValueSetter?.Invoke(item, value);
                break;
                
            case CellHandlerType.CollectionItem:
                if (handler.CollectionIndex >= 0 
                    && collections != null 
                    && collections.TryGetValue(handler.CollectionIndex, out var collection))
                {
                    var collectionMapping = handler.CollectionMapping!;
                    var itemType = collectionMapping.ItemType ?? typeof(object);
                    
                    // Check if this is a complex type with nested properties
                    var nestedMapping = collectionMapping.Registry?.GetCompiledMapping(itemType);
                    // Use pre-compiled type metadata from the helper instead of runtime reflection
                    var typeHelper = mapping.OptimizedCollectionHelpers?[handler.CollectionIndex];
                    if (nestedMapping != null && itemType != typeof(string) && typeHelper != null && !typeHelper.IsItemValueType && !typeHelper.IsItemPrimitive)
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
                            if (mapping.OptimizedCollectionHelpers != null && 
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
                            if (itemHelper != null && !itemHelper.IsItemValueType)
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
                            if (handler.CollectionItemConverter != null)
                            {
                                collection[handler.CollectionItemOffset] = handler.CollectionItemConverter(value);
                            }
                            else
                            {
                                collection[handler.CollectionItemOffset] = value;
                            }
                        }
                    }
                }
                break;
        }
    }
    
    private static void ProcessComplexCollectionItem(IList collection, OptimizedCellHandler handler, 
        object value, CompiledMapping<T> mapping)
    {
        // Ensure the collection has enough items
        while (collection.Count <= handler.CollectionItemOffset)
        {
            // Use precompiled default factory
            if (mapping.OptimizedCollectionHelpers == null || 
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
        if (item == null)
        {
            // Use precompiled factory for creating the item
            if (mapping.OptimizedCollectionHelpers == null || 
                handler.CollectionIndex < 0 || 
                handler.CollectionIndex >= mapping.OptimizedCollectionHelpers.Count)
            {
                throw new InvalidOperationException(
                    $"No OptimizedCollectionHelper found for collection at index {handler.CollectionIndex}. " +
                    "Ensure the mapping was properly compiled and optimized.");
            }
            
            var helper = mapping.OptimizedCollectionHelpers[handler.CollectionIndex];
            item = helper.DefaultItemFactory.Invoke();
            collection[handler.CollectionItemOffset] = item;
        }
        
        // The ValueSetter must be pre-compiled during optimization
        if (handler.ValueSetter == null)
        {
            // For nested mappings, we need to look up the pre-compiled setter
            if (mapping.NestedMappings != null && 
                mapping.NestedMappings.TryGetValue(handler.CollectionIndex, out var nestedInfo))
            {
                // Find the matching property setter in the nested mapping
                var nestedProp = nestedInfo.Properties.FirstOrDefault(p => p.PropertyName == handler.PropertyName);
                if (nestedProp?.Setter != null && item != null)
                {
                    nestedProp.Setter(item, value);
                    return;
                }
            }
            
            throw new InvalidOperationException(
                $"ValueSetter is null for complex collection item handler at property '{handler.PropertyName}'. " +
                "This indicates the mapping was not properly optimized. Ensure the type was mapped in the MappingRegistry.");
        }
        
        // Use the pre-compiled setter with built-in type conversion
        if (item != null) 
            handler.ValueSetter(item, value);
    }
    
    private static void FinalizeCollections(T item, CompiledMapping<T> mapping, Dictionary<int, IList> collections)
    {
        for (int i = 0; i < mapping.Collections.Count; i++)
        {
            var collectionMapping = mapping.Collections[i];
            if (collections.TryGetValue(i, out var list))
            {
                // Get the default value using precompiled factory if available
                object? defaultValue = null;
                if (mapping.OptimizedCollectionHelpers != null && i < mapping.OptimizedCollectionHelpers.Count)
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
                    bool isDefault = lastItem == null || 
                                   (listHelper != null && listHelper.IsItemValueType && lastItem.Equals(defaultValue));
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
                
                if (collectionMapping.Setter != null)
                {
                    // Use precompiled collection helper to convert to final type
                    if (mapping.OptimizedCollectionHelpers != null && i < mapping.OptimizedCollectionHelpers.Count)
                    {
                        var helper = mapping.OptimizedCollectionHelpers[i];
                        finalValue = helper.Finalizer(list);
                    }
                    
                    collectionMapping.Setter(item, finalValue);
                }
            }
        }
    }
    
    private static bool HasAnyData(T item, CompiledMapping<T> mapping)
    {
        // Check if any properties have non-default values
        foreach (var prop in mapping.Properties)
        {
            var value = prop.Getter(item);
            if (value != null && !IsDefaultValue(value))
            {
                return true;
            }
        }
        
        // Check if any collections have items
        foreach (var coll in mapping.Collections)
        {
            var collection = coll.Getter(item);
            if (collection != null && collection.Cast<object>().Any())
            {
                return true;
            }
        }
        
        return false;
    }
    
    private static bool IsDefaultValue(object value)
    {
        return value switch
        {
            string s => string.IsNullOrEmpty(s),
            int i => i == 0,
            long l => l == 0,
            decimal d => d == 0,
            double d => d == 0,
            float f => f == 0,
            bool b => !b,
            DateTime dt => dt == default,
            _ => false
        };
    }
    
    private static bool TryParseColumnIndex(string columnLetter, out int columnIndex)
    {
        columnIndex = 0;
        if (string.IsNullOrEmpty(columnLetter)) return false;
        
        // Convert column letter (A, B, AA, etc.) to index
        columnLetter = columnLetter.ToUpperInvariant();
        for (int i = 0; i < columnLetter.Length; i++)
        {
            char c = columnLetter[i];
            if (c < 'A' || c > 'Z') return false;
            columnIndex = columnIndex * 26 + (c - 'A' + 1);
        }
        
        return columnIndex > 0;
    }
}