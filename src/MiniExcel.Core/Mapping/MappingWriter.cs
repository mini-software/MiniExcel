using MiniExcelLib.Core.Helpers;
using MiniExcelLib.Core.OpenXml.Utils;

namespace MiniExcelLib.Core.Mapping;

internal class WriterBoundaries
{
    public int MaxPropertyRow { get; set; }
    public int MaxPropertyColumn { get; set; }
}

internal class ItemOffset
{
    public int RowOffset { get; set; }
    public int ColumnOffset { get; set; }
}


internal static partial class MappingWriter<T>
{
    [CreateSyncVersion]
    public static async Task<int[]> SaveAsAsync(Stream stream, IEnumerable<T> value, CompiledMapping<T> mapping, CancellationToken cancellationToken = default)
    {
        if (stream == null)
            throw new ArgumentNullException(nameof(stream));
        if (value == null)
            throw new ArgumentNullException(nameof(value));
        if (mapping == null)
            throw new ArgumentNullException(nameof(mapping));

        // Use optimized universal writer if mapping is optimized
        if (mapping.IsUniversallyOptimized)
        {
            return await SaveAsUniversalAsync(stream, value, mapping, cancellationToken).ConfigureAwait(false);
        }

        // Legacy path for non-optimized mappings
        // Convert mapped data to row-based format that OpenXmlWriter expects.
        // This uses an optimized streaming approach that only buffers rows for the current item.
        var mappedData = ConvertToMappedData(value, mapping);
        
        var configuration = new OpenXmlConfiguration { FastMode = true };
        var writer = await OpenXmlWriter
            .CreateAsync(stream, mappedData, mapping.WorksheetName, false, configuration, cancellationToken)
            .ConfigureAwait(false);
        
        return await writer.SaveAsAsync(cancellationToken).ConfigureAwait(false);
    }

    private static IEnumerable<IDictionary<string, object>> ConvertToMappedData(IEnumerable<T> value, CompiledMapping<T> mapping)
    {
        // Analyze mapping configuration to determine overlapping areas and boundaries
        var boundaries = CalculateMappingBoundaries(mapping);
        var allColumns = DetermineAllColumns(mapping);
        var rowBuffer = new SortedDictionary<int, Dictionary<string, object>>();
        
        // Process all items with proper offsets based on boundaries
        var itemIndex = 0;
        foreach (var item in value)
        {
            ProcessItemData(item, mapping, rowBuffer, itemIndex, boundaries);
            itemIndex++;
        }
        
        // Yield all rows from 1 to maxRow to ensure proper row positioning
        if (rowBuffer.Count > 0)
        {
            var maxRow = rowBuffer.Keys.Max();
            
            for (int rowNum = 1; rowNum <= maxRow; rowNum++)
            {
                var rowDict = CreateRowDict(rowBuffer, rowNum, allColumns);
                yield return rowDict;
                
                // Return dictionary to pool after yielding
                if (rowBuffer.TryGetValue(rowNum, out var sourceDict))
                {
                    DictionaryPool.Return(sourceDict);
                    rowBuffer.Remove(rowNum);
                }
            }
        }
    }
    
    private static void ProcessItemData(T item, CompiledMapping<T> mapping, SortedDictionary<int, Dictionary<string, object>> rowBuffer, int itemIndex, WriterBoundaries boundaries)
    {
        if(item == null)
            throw new ArgumentNullException(nameof(item));
        
        // Calculate offset for this item based on boundaries and mapping type
        var itemOffset = CalculateItemOffset(itemIndex, boundaries, mapping);
        
        // Add simple properties with offset
        foreach (var prop in mapping.Properties)
        {
            var propValue = prop.Getter(item);
            var cellValue = ConvertValue(propValue, prop);
            var offsetCellAddress = ApplyOffset(prop.CellAddress, itemOffset.RowOffset, itemOffset.ColumnOffset);
            AddCellToBuffer(rowBuffer, offsetCellAddress, cellValue);
        }
        
        // Add collections with offset
        foreach (var coll in mapping.Collections)
        {
            var collection = coll.Getter(item);
            if (collection != null)
            {
                var offsetStartCell = ApplyOffset(coll.StartCell, itemOffset.RowOffset, itemOffset.ColumnOffset);
                foreach (var cellInfo in StreamCollectionCellsWithOffset(collection, coll, offsetStartCell))
                {
                    AddCellToBuffer(rowBuffer, cellInfo.Key, cellInfo.Value);
                }
            }
        }
    }
    
    private static void AddCellToBuffer(SortedDictionary<int, Dictionary<string, object>> rowBuffer, string cellAddress, object value)
    {
        if (!ReferenceHelper.ParseReference(cellAddress, out _, out int row))
            return;
            
        if (!rowBuffer.ContainsKey(row))
            rowBuffer[row] = DictionaryPool.Rent();
            
        rowBuffer[row][cellAddress] = value;
    }
    
    private static WriterBoundaries CalculateMappingBoundaries<TItem>(CompiledMapping<TItem> mapping)
    {
        var boundaries = new WriterBoundaries();
        
        // Find the maximum row used by properties
        foreach (var prop in mapping.Properties)
        {
            if (ReferenceHelper.ParseReference(prop.CellAddress, out int col, out int row))
            {
                if (row > boundaries.MaxPropertyRow)
                    boundaries.MaxPropertyRow = row;
                if (col > boundaries.MaxPropertyColumn)
                    boundaries.MaxPropertyColumn = col;
            }
        }
        
        // Calculate boundaries for collections
        foreach (var coll in mapping.Collections)
        {
            if (ReferenceHelper.ParseReference(coll.StartCell, out _, out var startRow))
            {
                // Also update MaxPropertyRow to include collection start rows
                if (startRow > boundaries.MaxPropertyRow)
                    boundaries.MaxPropertyRow = startRow;
                    
                // Collection layout information is handled during actual writing
            }
        }
        
        return boundaries;
    }
    
    private static ItemOffset CalculateItemOffset<TItem>(int itemIndex, WriterBoundaries boundaries, CompiledMapping<TItem> mapping)
    {
        if (itemIndex == 0)
        {
            // First item uses original positions
            return new ItemOffset { RowOffset = 0, ColumnOffset = 0 };
        }
        
        // For subsequent items, we need to offset based on the mapping layout
        
        // If all properties are on row 1 (A1, B1, C1...), it's a table pattern - items go in consecutive rows
        // Otherwise, offset by the max property row + spacing for each item
        var allOnRow1 = mapping.Properties.All(p => p.CellRow == 1);
        var spacing = allOnRow1 ? 1 : (boundaries.MaxPropertyRow + 2);
        var rowOffset = itemIndex * spacing;
        
        return new ItemOffset { RowOffset = rowOffset, ColumnOffset = 0 };
    }
    
    private static string ApplyOffset(string cellAddress, int rowOffset, int columnOffset)
    {
        if (!ReferenceHelper.ParseReference(cellAddress, out int col, out int row))
            return cellAddress;
            
        var newRow = row + rowOffset;
        var newCol = col + columnOffset;
        
        return ReferenceHelper.ConvertCoordinatesToCell(newCol, newRow);
    }
    
    private static Dictionary<string, object> CreateRowDict(SortedDictionary<int, Dictionary<string, object>> rowBuffer, int rowNum, List<string> allColumns)
    {
        var rowDict = DictionaryPool.Rent();
        
        if (rowBuffer.TryGetValue(rowNum, out var sourceRow))
        {
            foreach (var column in allColumns)
            {
                object? cellValue = null;
                foreach (var kvp in sourceRow)
                {
                    if (ReferenceHelper.GetCellLetter(kvp.Key) == column)
                    {
                        cellValue = kvp.Value;
                        break;
                    }
                }
                rowDict[column] = cellValue ?? string.Empty;
            }
        }
        else
        {
            // Empty row
            foreach (var column in allColumns)
            {
                rowDict[column] = string.Empty;
            }
        }
        
        return rowDict;
    }
    
    private static List<string> DetermineAllColumns<TItem>(CompiledMapping<TItem> mapping)
    {
        var columns = new HashSet<string>();
        
        // Add columns from properties
        foreach (var prop in mapping.Properties)
        {
            if (ReferenceHelper.ParseReference(prop.CellAddress, out int col, out _))
            {
                var column = ReferenceHelper.GetCellLetter(ReferenceHelper.ConvertCoordinatesToCell(col, 1));
                if (!string.IsNullOrEmpty(column))
                    columns.Add(column);
            }
        }
        
        // For collections, determine columns from mapping configuration without iterating data
        foreach (var coll in mapping.Collections)
        {
            if (ReferenceHelper.ParseReference(coll.StartCell, out int startCol, out _))
            {
                // Only support vertical collections
                if (coll.Layout == CollectionLayout.Vertical)
                {
                    // Vertical layout or nested collection
                    if (coll.ItemMapping != null)
                    {
                        // For nested mappings, get columns from the item mapping
                            var itemMappingType = coll.ItemMapping.GetType();
                            var propsProperty = itemMappingType.GetProperty("Properties");

                            if (propsProperty?.GetValue(coll.ItemMapping) is IEnumerable<CompiledPropertyMapping> properties)
                            {
                                foreach (var prop in properties)
                                {
                                    if (ReferenceHelper.ParseReference(prop.CellAddress, out int propCol, out _))
                                    {
                                        var propColumn = ReferenceHelper.GetCellLetter(ReferenceHelper.ConvertCoordinatesToCell(propCol, 1));
                                        if (!string.IsNullOrEmpty(propColumn))
                                            columns.Add(propColumn);
                                    }
                                }
                            }
                        }
                        else
                        {
                            // Simple vertical collection
                            var col = ReferenceHelper.GetCellLetter(ReferenceHelper.ConvertCoordinatesToCell(startCol, 1));
                            if (!string.IsNullOrEmpty(col))
                                columns.Add(col);
                        }
                }
            }
        }
        
        // Ensure all columns between min and max are included to prevent compression
        if (columns.Count > 0)
        {
            var sortedColumns = columns.OrderBy(c => c).ToList();
            var minColumn = sortedColumns.First();
            var maxColumn = sortedColumns.Last();
            
            // Convert column letters to numbers for easier range calculation
            var minColNum = ReferenceHelper.ParseReference(minColumn + "1", out int minCol, out _) ? minCol : 1;
            var maxColNum = ReferenceHelper.ParseReference(maxColumn + "1", out int maxCol, out _) ? maxCol : 1;
            
            // Add all columns in the range
            var allColumnsInRange = new List<string>();
            for (int col = minColNum; col <= maxColNum; col++)
            {
                var columnLetter = ReferenceHelper.GetCellLetter(ReferenceHelper.ConvertCoordinatesToCell(col, 1));
                if (!string.IsNullOrEmpty(columnLetter))
                {
                    allColumnsInRange.Add(columnLetter);
                }
            }
            
            return allColumnsInRange;
        }
        
        return columns.OrderBy(c => c).ToList();
    }
    
    private static IEnumerable<KeyValuePair<string, object>> StreamCollectionCellsWithOffset(IEnumerable collection, CompiledCollectionMapping mapping, string offsetStartCell)
    {
        if (!ReferenceHelper.ParseReference(offsetStartCell, out int startColumn, out int startRow))
            throw new InvalidOperationException($"Invalid start cell address: {offsetStartCell}");
            
        var currentRow = startRow;
        var currentCol = startColumn;
        var itemIndex = 0;
        
        // Process collection items one at a time without buffering
        foreach (var item in collection)
        {
            if (mapping.ItemMapping != null && mapping.ItemType != null)
            {
                // Complex item with nested mapping
                var itemMapping = mapping.ItemMapping;
                var itemMappingType = itemMapping.GetType();
                var propsProperty = itemMappingType.GetProperty("Properties");

                if (propsProperty?.GetValue(itemMapping) is IEnumerable<CompiledPropertyMapping> properties)
                {
                    foreach (var prop in properties)
                    {
                        var propValue = prop.Getter(item);
                        var cellValue = ConvertValue(propValue, prop);
                        
                        // For nested mappings, we need to adjust the property's cell address relative to the collection item's position
                        if (!ReferenceHelper.ParseReference(prop.CellAddress, out int propCol, out int propRow))
                            continue;
                            
                        // Only vertical layout is supported
                        var cellAddress = mapping.Layout == CollectionLayout.Vertical 
                            ? ReferenceHelper.ConvertCoordinatesToCell(startColumn + propCol - 1, currentRow + propRow - 1)
                            : prop.CellAddress;
                        
                        yield return new KeyValuePair<string, object>(cellAddress, cellValue);
                    }
                }
            }
            else
            {
                // Simple item - just write the value
                var cellAddress = CalculateCellPositionWithBase(offsetStartCell, currentRow, currentCol, itemIndex, mapping, startColumn, startRow);
                yield return new KeyValuePair<string, object>(cellAddress, ConvertValue(item, null));
            }
            
            // Update position for next item
            UpdatePosition(ref currentRow, ref currentCol, ref itemIndex, mapping);
        }
    }
    
    private static string CalculateCellPositionWithBase(string baseCellAddress, int currentRow, int currentCol, int itemIndex, CompiledCollectionMapping mapping, int startColumn, int startRow)
    {
        // Only vertical layout is supported
        return ReferenceHelper.ConvertCoordinatesToCell(startColumn, currentRow);
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

    private static object ConvertValue(object? value, CompiledPropertyMapping? prop)
    {
        if (value == null) return string.Empty;
        
        if (prop != null && !string.IsNullOrEmpty(prop.Format))
        {
            if (value is IFormattable formattable)
                return formattable.ToString(prop.Format, null);
        }

        return value;
    }
    
    // Universal optimized writer implementation
    [CreateSyncVersion]
    private static async Task<int[]> SaveAsUniversalAsync(Stream stream, IEnumerable<T> value, CompiledMapping<T> mapping, CancellationToken cancellationToken = default)
    {
        if (!mapping.IsUniversallyOptimized)
            throw new InvalidOperationException("SaveAsUniversalAsync requires a universally optimized mapping");

        // Use optimized direct row streaming based on pre-calculated cell grid
        var rowEnumerable = CreateOptimizedRows(value, mapping);
        
        var configuration = new OpenXmlConfiguration { FastMode = true };
        var writer = await OpenXmlWriter
            .CreateAsync(stream, rowEnumerable, mapping.WorksheetName, false, configuration, cancellationToken)
            .ConfigureAwait(false);
        
        return await writer.SaveAsAsync(cancellationToken).ConfigureAwait(false);
    }

    private static IEnumerable<IDictionary<string, object>> CreateOptimizedRows(IEnumerable<T> items, CompiledMapping<T> mapping)
    {
        var boundaries = mapping.OptimizedBoundaries!;
        
        // For simple mappings without collections, handle row positioning
        if (!mapping.Collections.Any())
        {
            // If data starts at row > 1, we need to write placeholder rows
            if (boundaries.MinRow > 1)
            {
                // Write a single placeholder row with all columns
                var placeholderRow = new Dictionary<string, object>();
                for (int col = boundaries.MinColumn; col <= boundaries.MaxColumn; col++)
                {
                    var columnLetter = OpenXml.Utils.ReferenceHelper.GetCellLetter(
                        OpenXml.Utils.ReferenceHelper.ConvertCoordinatesToCell(col, 1));
                    placeholderRow[columnLetter] = "";
                }
                
                // Write placeholder rows until we reach the data start row
                for (int emptyRow = 1; emptyRow < boundaries.MinRow; emptyRow++)
                {
                    yield return new Dictionary<string, object>(placeholderRow);
                }
            }
            
            // Now write the actual data rows
            var currentRow = boundaries.MinRow;
            foreach (var item in items)
            {
                if (item == null) continue;
                
                // For column layouts (GridHeight > 1), we need to write multiple rows for each item
                if (boundaries.GridHeight > 1)
                {
                    // Write multiple rows for this item - one for each row in the grid
                    for (int gridRow = 0; gridRow < boundaries.GridHeight; gridRow++)
                    {
                        var absoluteRow = boundaries.MinRow + gridRow;
                        var row = CreateColumnLayoutRowForItem(item, absoluteRow, gridRow, mapping, boundaries);
                        if (row.Count > 0)
                        {
                            yield return row;
                        }
                    }
                }
                else
                {
                    // Regular single-row layout
                    var row = CreateSimpleRowForItem(item, currentRow, mapping, boundaries);
                    if (row.Count > 0)
                    {
                        yield return row;
                    }
                    currentRow++;
                }
            }
        }
        else
        {
            // Stream complex mappings with collections without buffering
            var cellGrid = mapping.OptimizedCellGrid!;
            
            // Stream rows without buffering the entire collection
            foreach (var row in StreamOptimizedRowsWithCollections(items, mapping, cellGrid, boundaries))
            {
                yield return row;
            }
        }
    }
    
    private static IEnumerable<IDictionary<string, object>> StreamOptimizedRowsWithCollections(
        IEnumerable<T> items, CompiledMapping<T> mapping, OptimizedCellHandler[,] cellGrid, 
        OptimizedMappingBoundaries boundaries)
    {
        // Write placeholder rows if needed
        if (boundaries.MinRow > 1)
        {
            var placeholderRow = new Dictionary<string, object>();
            
            // Find the maximum column that will have data
            int maxDataCol = 0;
            for (int relativeCol = 0; relativeCol < cellGrid.GetLength(1); relativeCol++)
            {
                for (int relativeRow = 0; relativeRow < cellGrid.GetLength(0); relativeRow++)
                {
                    var handler = cellGrid[relativeRow, relativeCol];
                    if (handler.Type != CellHandlerType.Empty)
                    {
                        maxDataCol = Math.Max(maxDataCol, relativeCol + boundaries.MinColumn);
                    }
                }
            }
            
            // Initialize all columns that will be used
            for (int col = 1; col <= maxDataCol; col++)
            {
                var columnLetter = OpenXml.Utils.ReferenceHelper.GetCellLetter(
                    OpenXml.Utils.ReferenceHelper.ConvertCoordinatesToCell(col, 1));
                placeholderRow[columnLetter] = "";
            }
            
            for (int emptyRow = 1; emptyRow < boundaries.MinRow; emptyRow++)
            {
                yield return new Dictionary<string, object>(placeholderRow);
            }
        }
        
        // Now stream the actual data using pre-calculated boundaries
        var itemEnumerator = items.GetEnumerator();
        if (!itemEnumerator.MoveNext()) yield break;
        
        var currentItem = itemEnumerator.Current;
        var currentItemIndex = 0;
        var currentRow = boundaries.MinRow;
        var hasMoreItems = true;
        
        // Track active collection enumerators
        var collectionEnumerators = new Dictionary<int, IEnumerator>();
        var collectionItems = new Dictionary<int, object?>();
        
        while (hasMoreItems || collectionEnumerators.Count > 0)
        {
            var row = new Dictionary<string, object>();
            
            // Initialize all columns with empty values to ensure proper column structure
            for (int col = boundaries.MinColumn; col <= boundaries.MaxColumn; col++)
            {
                var columnLetter = OpenXml.Utils.ReferenceHelper.GetCellLetter(
                    OpenXml.Utils.ReferenceHelper.ConvertCoordinatesToCell(col, 1));
                row[columnLetter] = "";
            }
            
            // Process each column in the current row
            for (int col = boundaries.MinColumn; col <= boundaries.MaxColumn; col++)
            {
                var relativeRow = currentRow - boundaries.MinRow;
                var relativeCol = col - boundaries.MinColumn;
                
                if (relativeRow >= 0 && relativeRow < cellGrid.GetLength(0) && 
                    relativeCol >= 0 && relativeCol < cellGrid.GetLength(1))
                {
                    var handler = cellGrid[relativeRow, relativeCol];
                    
                    if (handler.Type == CellHandlerType.Property && currentItem != null)
                    {
                        // Simple property - extract value
                        if (handler.ValueExtractor != null)
                        {
                            var value = handler.ValueExtractor(currentItem, 0);
                            var columnLetter = OpenXml.Utils.ReferenceHelper.GetCellLetter(
                                OpenXml.Utils.ReferenceHelper.ConvertCoordinatesToCell(col, 1));
                            row[columnLetter] = value ?? "";
                        }
                    }
                    else if (handler.Type == CellHandlerType.CollectionItem && currentItem != null)
                    {
                        // Collection item - check if we need to start/continue enumeration
                        var collIndex = handler.CollectionIndex;
                        
                        // Check if we're within collection boundaries
                        if (handler.BoundaryRow == -1 || currentRow < handler.BoundaryRow)
                        {
                            // Initialize enumerator if needed
                            if (!collectionEnumerators.ContainsKey(collIndex))
                            {
                                var collection = handler.CollectionMapping?.Getter(currentItem);
                                if (collection != null)
                                {
                                    var enumerator = collection.GetEnumerator();
                                    if (enumerator.MoveNext())
                                    {
                                        collectionEnumerators[collIndex] = enumerator;
                                        collectionItems[collIndex] = enumerator.Current;
                                    }
                                }
                            }
                            
                            // Get current collection item
                            if (collectionItems.TryGetValue(collIndex, out var collItem))
                            {
                                var columnLetter = OpenXml.Utils.ReferenceHelper.GetCellLetter(
                                    OpenXml.Utils.ReferenceHelper.ConvertCoordinatesToCell(col, 1));
                                row[columnLetter] = collItem ?? "";
                            }
                        }
                    }
                }
            }
            
            // Always yield rows to maintain proper spacing
            yield return row;
            
            currentRow++;
            
            // Check if we need to advance collection enumerators
            bool advancedAnyCollection = false;
            foreach (var kvp in collectionEnumerators.ToList())
            {
                var collIndex = kvp.Key;
                var enumerator = kvp.Value;
                
                // Check if this collection should advance based on row spacing
                // This is simplified - real logic would check the actual collection mapping
                if (enumerator.MoveNext())
                {
                    collectionItems[collIndex] = enumerator.Current;
                    advancedAnyCollection = true;
                }
                else
                {
                    // Collection exhausted
                    collectionEnumerators.Remove(collIndex);
                    collectionItems.Remove(collIndex);
                }
            }
            
            // If no collections advanced and we're past the pattern height, move to next item
            if (!advancedAnyCollection && boundaries.PatternHeight > 0 && 
                (currentRow - boundaries.MinRow) >= boundaries.PatternHeight)
            {
                if (itemEnumerator.MoveNext())
                {
                    currentItem = itemEnumerator.Current;
                    currentItemIndex++;
                    // Clear collection enumerators for new item
                    collectionEnumerators.Clear();
                    collectionItems.Clear();
                }
                else
                {
                    hasMoreItems = false;
                    currentItem = default(T);
                }
            }
        }
    }

    private static Dictionary<string, object> CreateSimpleRowForItem(T item, int currentRow, CompiledMapping<T> mapping, OptimizedMappingBoundaries boundaries)
    {
        var row = new Dictionary<string, object>();
        
        // Initialize all columns with empty values
        for (int col = boundaries.MinColumn; col <= boundaries.MaxColumn; col++)
        {
            var columnLetter = OpenXml.Utils.ReferenceHelper.GetCellLetter(
                OpenXml.Utils.ReferenceHelper.ConvertCoordinatesToCell(col, 1));
            row[columnLetter] = string.Empty;
        }
        
        // Fill in property values
        foreach (var prop in mapping.Properties)
        {
            var value = prop.Getter(item);
            if (value != null)
            {
                var columnLetter = OpenXml.Utils.ReferenceHelper.GetCellLetter(
                    OpenXml.Utils.ReferenceHelper.ConvertCoordinatesToCell(prop.CellColumn, 1));
                
                
                // Apply formatting if specified
                if (!string.IsNullOrEmpty(prop.Format) && value is IFormattable formattable)
                {
                    value = formattable.ToString(prop.Format, null);
                }
                
                row[columnLetter] = value;
            }
        }
        
        return row;
    }

    private static Dictionary<string, object> CreateColumnLayoutRowForItem(T item, int absoluteRow, int gridRow, CompiledMapping<T> mapping, OptimizedMappingBoundaries boundaries)
    {
        var row = new Dictionary<string, object>();
        
        // Initialize all columns with empty values - start from column A to ensure proper column positioning
        int startCol = Math.Min(1, boundaries.MinColumn);  // Always include column A (column 1)
        for (int col = startCol; col <= boundaries.MaxColumn; col++)
        {
            var columnLetter = OpenXml.Utils.ReferenceHelper.GetCellLetter(
                OpenXml.Utils.ReferenceHelper.ConvertCoordinatesToCell(col, 1));
            row[columnLetter] = string.Empty;
        }
        
        // Only fill in the property value that belongs to this specific row
        foreach (var prop in mapping.Properties)
        {
            // Check if this property belongs to the current row
            if (prop.CellRow == absoluteRow)
            {
                var value = prop.Getter(item);
                if (value != null)
                {
                    var columnLetter = OpenXml.Utils.ReferenceHelper.GetCellLetter(
                        OpenXml.Utils.ReferenceHelper.ConvertCoordinatesToCell(prop.CellColumn, 1));
                    
                    
                    // Apply formatting if specified
                    if (!string.IsNullOrEmpty(prop.Format) && value is IFormattable formattable)
                    {
                        value = formattable.ToString(prop.Format, null);
                    }
                    
                    row[columnLetter] = value;
                }
            }
        }
        
        return row;
    }

    private static object? ExtractCollectionItemValueForItem(T item, OptimizedCellHandler cellHandler,
        int absoluteRow, int absoluteCol, OptimizedMappingBoundaries boundaries)
    {
        if (cellHandler.CollectionMapping == null || item == null)
            return null;

        var collectionMapping = cellHandler.CollectionMapping;
        var collection = collectionMapping.Getter(item);
        if (collection == null) return null;

        // Calculate the actual item index based on the absolute row
        // This handles rows beyond our pre-calculated grid
        int actualItemIndex = cellHandler.CollectionItemOffset;
        
        // For vertical collections with row spacing, calculate the actual index
        if (collectionMapping.Layout == CollectionLayout.Vertical)
        {
            var rowsSinceStart = absoluteRow - collectionMapping.StartCellRow;
            if (rowsSinceStart >= 0)
            {
                // Calculate which item this row belongs to based on row spacing
                // This is O(1) - just arithmetic, no iteration
                actualItemIndex = rowsSinceStart / (1 + collectionMapping.RowSpacing);
            }
        }
        
        // If we have a pre-compiled value extractor for nested properties, use it
        if (cellHandler.ValueExtractor != null)
        {
            // The ValueExtractor was pre-compiled to extract the specific property from the nested object
            return cellHandler.ValueExtractor(item, actualItemIndex);
        }

        // Otherwise fall back to simple collection item extraction
        var collectionItems = collection.Cast<object>().ToArray();
        if (collectionItems.Length == 0) return null;

        // Return the collection item if index is valid
        return actualItemIndex >= 0 && actualItemIndex < collectionItems.Length ? collectionItems[actualItemIndex] : null;
    }
}