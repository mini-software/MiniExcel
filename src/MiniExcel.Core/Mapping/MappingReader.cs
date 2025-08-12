using MiniExcelLib.Core.Helpers;
using MiniExcelLib.Core.OpenXml.Utils;

namespace MiniExcelLib.Core.Mapping;

internal static partial class MappingReader<T> where T : class, new()
{
    [CreateSyncVersion]
    public static async Task<IEnumerable<T>> QueryAsync(Stream stream, CompiledMapping<T> mapping, CancellationToken cancellationToken = default)
    {
        if (stream == null)
            throw new ArgumentNullException(nameof(stream));
        if (mapping == null)
            throw new ArgumentNullException(nameof(mapping));

        // Use optimized universal reader if mapping is optimized
        if (mapping.IsUniversallyOptimized)
        {
            return await QueryUniversalAsync(stream, mapping, cancellationToken).ConfigureAwait(false);
        }

        // Legacy path for non-optimized mappings
        var importer = new OpenXmlImporter();
        
        var dataList = new List<IDictionary<string, object>>();
        
        await foreach (var row in importer.QueryAsync(stream, useHeaderRow: false, sheetName: mapping.WorksheetName, startCell: "A1", cancellationToken: cancellationToken).ConfigureAwait(false))
        {
            if (row is IDictionary<string, object> dict)
            {
                // Include all rows, even if they appear empty
                dataList.Add(dict);
            }
        }
        
        if (!dataList.Any())
        {
            return [];
        }

        // Build a cell lookup dictionary for efficient access
        var cellLookup = BuildCellLookup(dataList);
        
        // Read the mapped data
        var results = ReadMappedData(cellLookup, mapping);
        return results;
    }
    
    [CreateSyncVersion]
    public static async Task<T> QuerySingleAsync(Stream stream, CompiledMapping<T> mapping, CancellationToken cancellationToken = default)
    {
        var results = await QueryAsync(stream, mapping, cancellationToken).ConfigureAwait(false);
        return results.FirstOrDefault() ?? new T();
    }
    
    private static Dictionary<string, object> BuildCellLookup(List<IDictionary<string, object>> data)
    {
        var lookup = new Dictionary<string, object>();
        
        for (int rowIndex = 0; rowIndex < data.Count; rowIndex++)
        {
            var row = data[rowIndex];
            var rowNumber = rowIndex + 1; // Row is 1-based
            
            foreach (var kvp in row)
            {
                var columnLetter = kvp.Key;
                var cellAddress = $"{columnLetter}{rowNumber}";
                lookup[cellAddress] = kvp.Value;
            }
        }
        
        return lookup;
    }
    
    private static IEnumerable<T> ReadMappedData(Dictionary<string, object> cellLookup, CompiledMapping<T> mapping)
    {
        // Calculate the expected spacing between items based on mapping configuration
        var maxPropertyRow = 0;
        foreach (var prop in mapping.Properties)
        {
            if (prop.CellRow > maxPropertyRow)
                maxPropertyRow = prop.CellRow;
        }
        
        // Also check collection start rows as they may be part of the item definition
        foreach (var coll in mapping.Collections)
        {
            if (coll.StartCellRow > maxPropertyRow)
                maxPropertyRow = coll.StartCellRow;
        }
        
        // Determine item spacing based on the mapping pattern
        // If all properties are on row 1 (A1, B1, C1...), it's likely a table pattern where each row is an item
        // Otherwise, use the writer's spacing pattern (maxPropertyRow + 2)
        var allOnRow1 = mapping.Properties.All(p => p.CellRow == 1);
        var itemSpacing = allOnRow1 ? 1 : maxPropertyRow + 2;
        
        #if DEBUG
        System.Diagnostics.Debug.WriteLine($"ReadMappedData: allOnRow1={allOnRow1}, itemSpacing={itemSpacing}");
        #endif
        
        // Find the base row where properties start
        var baseRow = int.MaxValue;
        foreach (var prop in mapping.Properties)
        {
            if (prop.CellRow < baseRow)
                baseRow = prop.CellRow;
        }
        
        if (baseRow == int.MaxValue)
            baseRow = 1;
            
        // Debug logging
        #if DEBUG
        System.Diagnostics.Debug.WriteLine($"ReadMappedData: maxPropertyRow={maxPropertyRow}, itemSpacing={itemSpacing}, baseRow={baseRow}");
        #endif
            
        // Read items at expected intervals
        var currentRow = baseRow;
        var itemsFound = 0;
        var maxItems = 1_000_000; // Safety limit - 1 million items should be enough
        
        while (itemsFound < maxItems)
        {
            var result = new T();
            var hasData = false;
            
            #if DEBUG
            System.Diagnostics.Debug.WriteLine($"Reading item at currentRow={currentRow}");
            #endif
            
            // Read simple properties at current offset
            foreach (var prop in mapping.Properties)
            {
                var offsetRow = currentRow + (prop.CellRow - baseRow);
                var cellAddress = ReferenceHelper.ConvertCoordinatesToCell(prop.CellColumn, offsetRow);
                    
                    if (cellLookup.TryGetValue(cellAddress, out var value))
                    {
                        SetPropertyValue(result, prop, value);
                        hasData = true;
                        #if DEBUG
                        System.Diagnostics.Debug.WriteLine($"  Found property {prop.PropertyName} at {cellAddress}: {value}");
                        #endif
                    }
                    else
                    {
                        // Try column compression fallback
                        var fallbackAddress = $"A{offsetRow}";
                        if (cellLookup.TryGetValue(fallbackAddress, out var fallbackValue))
                        {
                            SetPropertyValue(result, prop, fallbackValue);
                            hasData = true;
                            #if DEBUG
                            System.Diagnostics.Debug.WriteLine($"  Found property {prop.PropertyName} at fallback {fallbackAddress}: {fallbackValue}");
                            #endif
                        }
                    }
                }
            
            if (!hasData)
            {
                // No more items found
                #if DEBUG
                System.Diagnostics.Debug.WriteLine($"No data found at row {currentRow}, stopping");
                #endif
                break;
            }
            
            // Read collections at current offset
            for (int collIndex = 0; collIndex < mapping.Collections.Count; collIndex++)
            {
                var coll = mapping.Collections[collIndex];
                var offsetStartRow = currentRow + (coll.StartCellRow - baseRow);
                var offsetStartCell = ReferenceHelper.ConvertCoordinatesToCell(coll.StartCellColumn, offsetStartRow);
                    
                    // Determine collection boundaries
                    int? maxRow = null;
                    int? maxCol = null;
                    
                    // Check if there's another collection after this one on the same item
                    for (int nextCollIndex = collIndex + 1; nextCollIndex < mapping.Collections.Count; nextCollIndex++)
                    {
                        var nextColl = mapping.Collections[nextCollIndex];
                        var nextOffsetStartRow = currentRow + (nextColl.StartCellRow - baseRow);
                        
                        // Only vertical collections are supported
                        if (coll.Layout == CollectionLayout.Vertical && nextColl.Layout == CollectionLayout.Vertical && nextColl.StartCellColumn == coll.StartCellColumn)
                        {
                            maxRow = nextOffsetStartRow - 1;
                            break;
                        }
                    }
                    
                    // Check if there's definitely another item to limit collection boundaries
                    // This prevents reading collection data from the next item
                    if (maxRow == null && coll.Layout == CollectionLayout.Vertical)
                    {
                        // Check if there's a next item by looking for ALL properties (not just one)
                        // Only consider it a next item if we find MULTIPLE property values
                        var nextItemPropertyCount = 0;
                        if (mapping.Properties.Any())
                        {
                            // Check all properties to see if any exist at the next item position
                            foreach (var prop in mapping.Properties)
                            {
                                if (ReferenceHelper.ParseReference(prop.CellAddress, out int propCol, out int propRow))
                                {
                                    var nextItemRow = currentRow + itemSpacing + (propRow - baseRow);
                                    var nextItemCell = ReferenceHelper.ConvertCoordinatesToCell(propCol, nextItemRow);
                                    if (cellLookup.TryGetValue(nextItemCell, out var value) && value != null)
                                    {
                                        nextItemPropertyCount++;
                                    }
                                }
                            }
                        }
                        
                        // Only limit if we find at least 2 properties or the majority of properties
                        var minPropsForNextItem = Math.Max(2, mapping.Properties.Count / 2);
                        if (nextItemPropertyCount >= minPropsForNextItem)
                        {
                            maxRow = currentRow + itemSpacing - 1;
                        }
                    }
                    
                    var collectionData = ReadCollectionDataWithOffset(cellLookup, coll, offsetStartCell, maxRow, maxCol);
                    
                    #if DEBUG
                    System.Diagnostics.Debug.WriteLine($"  Collection {coll.PropertyName} at {offsetStartCell}: {collectionData.Count} items");
                    #endif
                    
                    SetCollectionValue(result, coll, collectionData);
            }
            
            yield return result;
            itemsFound++;
            currentRow += itemSpacing;
            
            #if DEBUG
            System.Diagnostics.Debug.WriteLine($"Item {itemsFound} read, moving to row {currentRow}");
            #endif
        }
    }
    
    private static void SetPropertyValue(T instance, CompiledPropertyMapping prop, object value)
    {
        if (prop.Setter != null)
        {
            var convertedValue = ConversionHelper.ConvertValue(value, prop.PropertyType, prop.Format);
            prop.Setter(instance, convertedValue);
        }
    }
    
    private static List<object> ReadCollectionDataWithOffset(Dictionary<string, object> cellLookup, CompiledCollectionMapping coll, string offsetStartCell, int? maxRow = null, int? maxCol = null)
    {
        var results = new List<object>();
        
        if (!ReferenceHelper.ParseReference(offsetStartCell, out int startColumn, out int startRow))
            return results;
        
        var currentRow = startRow;
        var currentCol = startColumn;
        var itemIndex = 0;
        var emptyCellCount = 0;
        const int maxEmptyCells = 10;
        const int maxIterations = 1000;
        var iterations = 0;
        
        while (emptyCellCount < maxEmptyCells && iterations < maxIterations && (!maxRow.HasValue || currentRow <= maxRow.Value))
        {
            if (coll.ItemMapping != null && coll.ItemType != null)
            {
                var item = ReadComplexItem(cellLookup, coll, currentRow, currentCol, itemIndex);
                if (item != null)
                {
                    results.Add(item);
                    emptyCellCount = 0;
                }
                else
                {
                    emptyCellCount++;
                }
            }
            else
            {
                var cellAddress = CalculateCellPosition(offsetStartCell, currentRow, currentCol, itemIndex, coll);
                if (cellLookup.TryGetValue(cellAddress, out var value) && value != null && !string.IsNullOrEmpty(value.ToString()))
                {
                    results.Add(value);
                    emptyCellCount = 0;
                }
                else
                {
                    emptyCellCount++;
                }
            }
            
            UpdatePosition(ref currentRow, ref currentCol, ref itemIndex, coll);
            iterations++;
        }
        
        return results;
    }
    
    
    private static object? ReadComplexItem(Dictionary<string, object> cellLookup, CompiledCollectionMapping coll, int currentRow, int currentCol, int itemIndex)
    {
        if (coll.ItemType == null || coll.ItemMapping == null)
            return null;
            
        var item = Activator.CreateInstance(coll.ItemType);
        if (item == null)
            return null;
            
        var itemMapping = coll.ItemMapping;
        var itemMappingType = itemMapping.GetType();
        var propsProperty = itemMappingType.GetProperty("Properties");
        var properties = propsProperty?.GetValue(itemMapping) as IEnumerable<CompiledPropertyMapping>;
        
        var hasAnyValue = false;
        
        if (properties != null)
        {
            foreach (var prop in properties)
            {
                // For nested mappings, we need to adjust the property's cell address relative to the collection item's position
                if (!ReferenceHelper.ParseReference(prop.CellAddress, out int propCol, out int propRow))
                    continue;
                    
                // Only vertical layout is supported
                var cellAddress = coll.Layout == CollectionLayout.Vertical
                    ? ReferenceHelper.ConvertCoordinatesToCell(currentCol + propCol - 1, currentRow + propRow - 1)
                    : prop.CellAddress;
                
                if (cellLookup.TryGetValue(cellAddress, out var value) && value != null && !string.IsNullOrEmpty(value.ToString()))
                {
                    SetItemPropertyValue(item, prop, value);
                    hasAnyValue = true;
                }
            }
        }
        
        return hasAnyValue ? item : null;
    }
    
    private static void SetItemPropertyValue(object instance, CompiledPropertyMapping prop, object value)
    {
        if (prop.Setter == null) return;
        
        var convertedValue = ConversionHelper.ConvertValue(value, prop.PropertyType, prop.Format);
        prop.Setter(instance, convertedValue);
    }
    
    private static void SetCollectionValue(T instance, CompiledCollectionMapping coll, List<object> items)
    {
        if (coll.Setter != null)
        {
            var targetType = typeof(T);
            var propertyInfo = targetType.GetProperty(coll.PropertyName);
            
            if (propertyInfo != null)
            {
                var collectionType = propertyInfo.PropertyType;
                var convertedCollection = ConvertToTypedCollection(items, collectionType, coll.ItemType);
                coll.Setter(instance, convertedCollection);
            }
        }
    }
    
    private static string CalculateCellPosition(string baseCellAddress, int currentRow, int currentCol, int itemIndex, CompiledCollectionMapping mapping)
    {
        if (!ReferenceHelper.ParseReference(baseCellAddress, out int baseColumn, out int baseRow))
            return baseCellAddress;
        
        // Only vertical layout is supported
        return ReferenceHelper.ConvertCoordinatesToCell(baseColumn, currentRow);
    }
    
    private static void UpdatePosition(ref int currentRow, ref int currentCol, ref int itemIndex, CompiledCollectionMapping mapping)
    {
        itemIndex++;
        
        // Only vertical layout is supported
        if (mapping.Layout == CollectionLayout.Vertical)
        {
            currentRow += 1 + mapping.RowSpacing;
        }
    }
    private static object? ConvertToTypedCollection(List<object> items, Type collectionType, Type? itemType)
    {
        if (items.Count == 0)
        {
            // For arrays, return empty array instead of null
            if (collectionType.IsArray)
            {
                var elementType = collectionType.GetElementType() ?? typeof(object);
                return Array.CreateInstance(elementType, 0);
            }
            return null;
        }
            
        // Handle arrays
        if (collectionType.IsArray)
        {
            var elementType = collectionType.GetElementType() ?? typeof(object);
            var array = Array.CreateInstance(elementType, items.Count);
            for (int i = 0; i < items.Count; i++)
            {
                array.SetValue(ConversionHelper.ConvertValue(items[i], elementType, null), i);
            }
            return array;
        }
        
        // Handle List<T>
        if (collectionType.IsGenericType && collectionType.GetGenericTypeDefinition() == typeof(List<>))
        {
            var elementType = collectionType.GetGenericArguments()[0];
            var listType = typeof(List<>).MakeGenericType(elementType);
            var list = Activator.CreateInstance(listType) as IList;
            
            foreach (var item in items)
            {
                var convertedValue = ConversionHelper.ConvertValue(item, elementType, null);
                if (convertedValue != null)
                {
                    list?.Add(convertedValue);
                }
            }
            return list;
        }
        
        // Handle IEnumerable<T>
        if (collectionType.IsGenericType && 
            (collectionType.GetGenericTypeDefinition() == typeof(IEnumerable<>) ||
             collectionType.GetInterface(typeof(IEnumerable<>).Name) != null))
        {
            var elementType = itemType ?? collectionType.GetGenericArguments()[0];
            var listType = typeof(List<>).MakeGenericType(elementType);
            var list = Activator.CreateInstance(listType) as IList;
            
            foreach (var item in items)
            {
                list?.Add(item); // Items are already converted
            }
            return list;
        }
        
        return items;
    }
    
    // Universal optimized reader implementation
    [CreateSyncVersion]
    private static async Task<IEnumerable<T>> QueryUniversalAsync(Stream stream, CompiledMapping<T> mapping, CancellationToken cancellationToken = default)
    {
        if (!mapping.IsUniversallyOptimized)
            throw new InvalidOperationException("QueryUniversalAsync requires a universally optimized mapping");

        var boundaries = mapping.OptimizedBoundaries!;
        var cellGrid = mapping.OptimizedCellGrid!;
        var columnHandlers = mapping.OptimizedColumnHandlers!;
        
        var results = new List<T>();
        
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
                if (rowDict == null) continue;
                
                
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
                            results.Add(currentItem);
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
                    ProcessCellValue(handler, kvp.Value, currentItem, currentCollections, gridRow);
                }
            }
            
            // Finalize the last item if we have one
            if (currentItem != null && currentCollections != null)
            {
                FinalizeCollections(currentItem, mapping, currentCollections);
                
                
                if (HasAnyData(currentItem, mapping))
                {
                    results.Add(currentItem);
                }
                else
                {
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
                    if (rowDict == null) continue;
                    
                    int rowNumber = currentRowIndex;
                    
                    // Process properties for this row
                    foreach (var prop in mapping.Properties)
                    {
                        if (prop.CellRow == rowNumber)
                        {
                            var columnLetter = OpenXml.Utils.ReferenceHelper.GetCellLetter(
                                OpenXml.Utils.ReferenceHelper.ConvertCoordinatesToCell(prop.CellColumn, 1));
                            
                            if (rowDict.TryGetValue(columnLetter, out var value) && value != null)
                            {
                                // Apply type conversion if needed
                                if (prop.Setter != null)
                                {
                                    var targetType = prop.PropertyType;
                                    if (value.GetType() != targetType)
                                    {
                                        // Pre-compiled conversion logic
                                        try
                                        {
                                            value = targetType switch
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
                                            // Fallback to string parsing
                                            var str = value.ToString();
                                            if (!string.IsNullOrEmpty(str))
                                            {
                                                value = targetType switch
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
                                        }
                                    }
                                    prop.Setter.Invoke(item, value);
                                }
                            }
                        }
                    }
                }
                
                if (HasAnyData(item, mapping))
                {
                    results.Add(item);
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
                    if (rowDict == null) continue;
                    
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
                        if (allOnRow1 || prop.CellRow == rowNumber)
                        {
                            var columnLetter = OpenXml.Utils.ReferenceHelper.GetCellLetter(
                                OpenXml.Utils.ReferenceHelper.ConvertCoordinatesToCell(prop.CellColumn, 1));
                            
                            if (rowDict.TryGetValue(columnLetter, out var value) && value != null)
                            {
                                // Apply type conversion if needed
                                if (prop.Setter != null)
                                {
                                    var targetType = prop.PropertyType;
                                    if (value.GetType() != targetType)
                                    {
                                        // Pre-compiled conversion logic
                                        try
                                        {
                                            value = targetType switch
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
                                            // Fallback to string parsing
                                            var str = value.ToString();
                                            if (!string.IsNullOrEmpty(str))
                                            {
                                                value = targetType switch
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
                                        }
                                    }
                                    prop.Setter.Invoke(item, value);
                                }
                            }
                        }
                    }
                    
                    if (HasAnyData(item, mapping))
                    {
                        results.Add(item);
                    }
                }
            }
        }
        
        return results;
    }
    
    private static Dictionary<int, IList> InitializeCollections(CompiledMapping<T> mapping)
    {
        var collections = new Dictionary<int, IList>();
        
        for (int i = 0; i < mapping.Collections.Count; i++)
        {
            var collection = mapping.Collections[i];
            var itemType = collection.ItemType ?? typeof(object);
            
            // Check if this is a complex type with nested mapping
            var nestedMapping = collection.Registry?.GetCompiledMapping(itemType);
            if (nestedMapping != null && itemType != typeof(string) && !itemType.IsValueType && !itemType.IsPrimitive)
            {
                // Complex type - we'll build objects as we go
                var listType = typeof(List<>).MakeGenericType(itemType);
                collections[i] = (IList)Activator.CreateInstance(listType)!;
            }
            else
            {
                // Simple type collection
                var listType = typeof(List<>).MakeGenericType(itemType);
                collections[i] = (IList)Activator.CreateInstance(listType)!;
            }
        }
        
        return collections;
    }
    
    private static void ProcessCellValue(OptimizedCellHandler handler, object value, T item, 
        Dictionary<int, IList> collections, int relativeRow)
    {
        switch (handler.Type)
        {
            case CellHandlerType.Property:
                // Direct property - use pre-compiled setter
                handler.ValueSetter?.Invoke(item, value);
                break;
                
            case CellHandlerType.CollectionItem:
                if (handler.CollectionIndex >= 0 && collections.ContainsKey(handler.CollectionIndex))
                {
                    var collection = collections[handler.CollectionIndex];
                    var collectionMapping = handler.CollectionMapping!;
                    var itemType = collectionMapping.ItemType ?? typeof(object);
                    
                    // Check if this is a complex type with nested properties
                    var nestedMapping = collectionMapping.Registry?.GetCompiledMapping(itemType);
                    if (nestedMapping != null && itemType != typeof(string) && !itemType.IsValueType && !itemType.IsPrimitive)
                    {
                        // Complex type - we need to build/update the object
                        ProcessComplexCollectionItem(collection, handler, value, itemType, nestedMapping);
                    }
                    else
                    {
                        // Simple type - add directly
                        while (collection.Count <= handler.CollectionItemOffset)
                        {
                            // For value types, we need to add default value not null
                            var defaultValue = itemType.IsValueType ? Activator.CreateInstance(itemType) : null;
                            collection.Add(defaultValue);
                        }
                        
                        // Skip empty values for value type collections
                        if (value == null || (value is string str && string.IsNullOrEmpty(str)))
                        {
                            // Don't add empty values to value type collections
                            if (!itemType.IsValueType)
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
        object value, Type itemType, object nestedMapping)
    {
        // Ensure the collection has enough items
        while (collection.Count <= handler.CollectionItemOffset)
        {
            collection.Add(Activator.CreateInstance(itemType));
        }
        
        var item = collection[handler.CollectionItemOffset];
        if (item == null)
        {
            item = Activator.CreateInstance(itemType)!;
            collection[handler.CollectionItemOffset] = item;
        }
        
        // The ValueSetter must be pre-compiled during optimization
        if (handler.ValueSetter == null)
        {
            throw new InvalidOperationException(
                $"ValueSetter is null for complex collection item handler at property '{handler.PropertyName}'. " +
                "This indicates the mapping was not properly optimized. Ensure the type was mapped in the MappingRegistry.");
        }
        
        // Use the pre-compiled setter with built-in type conversion
        handler.ValueSetter(item, value);
    }
    
    private static void FinalizeCollections(T item, CompiledMapping<T> mapping, Dictionary<int, IList> collections)
    {
        for (int i = 0; i < mapping.Collections.Count; i++)
        {
            var collectionMapping = mapping.Collections[i];
            if (collections.TryGetValue(i, out var list))
            {
                // Remove any trailing null or default values
                var itemType = collectionMapping.ItemType ?? typeof(object);
                while (list.Count > 0)
                {
                    var lastItem = list[list.Count - 1];
                    bool isDefault = lastItem == null || 
                                   (itemType.IsValueType && lastItem.Equals(Activator.CreateInstance(itemType)));
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
                    // Get the property type to determine if we need array conversion
                    var propInfo = typeof(T).GetProperty(collectionMapping.PropertyName);
                    if (propInfo != null && propInfo.PropertyType.IsArray)
                    {
                        var elementType = propInfo.PropertyType.GetElementType()!;
                        var array = Array.CreateInstance(elementType, list.Count);
                        list.CopyTo(array, 0);
                        finalValue = array;
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
    
    private static int ExtractRowNumber(IDictionary<string, object> row)
    {
        // Try to get row number from metadata
        if (row.TryGetValue("__rowIndex", out var rowIndex) && rowIndex is int index)
        {
            return index;
        }
        
        // Fallback: parse from first cell reference
        foreach (var key in row.Keys)
        {
            if (!key.StartsWith("__") && TryParseRowFromCellReference(key, out int rowNum))
            {
                return rowNum;
            }
        }
        
        return 1; // Default to row 1
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
    
    private static bool TryParseRowFromCellReference(string cellRef, out int row)
    {
        row = 0;
        if (string.IsNullOrEmpty(cellRef)) return false;
        
        // Find where letters end and numbers begin
        int i = 0;
        while (i < cellRef.Length && char.IsLetter(cellRef[i])) i++;
        
        if (i < cellRef.Length && int.TryParse(cellRef.Substring(i), out row))
        {
            return row > 0;
        }
        
        return false;
    }
}