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
            // For complex mappings with collections, use the existing grid approach
            var cellGrid = mapping.OptimizedCellGrid!;
            var itemList = items.ToList();
            if (itemList.Count == 0) yield break;

            // If data starts at row > 1, we need to write placeholder rows first
            if (boundaries.MinRow > 1)
            {
                // Write empty placeholder rows to position data correctly
                // IMPORTANT: Must include ALL columns that will be used in data rows
                // Otherwise OpenXmlWriter will only write columns present in first row
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

            var totalRowsNeeded = CalculateTotalRowsNeeded(itemList, mapping, boundaries);
            
            for (int absoluteRow = boundaries.MinRow; absoluteRow <= totalRowsNeeded; absoluteRow++)
            {
                var row = CreateRowForAbsoluteRow(absoluteRow, itemList, mapping, cellGrid, boundaries);
                
                // Always yield rows, even if empty, to maintain proper row spacing
                // If row is empty, ensure it has at least column A for OpenXmlWriter
                if (row.Count == 0)
                {
                    row["A"] = "";
                }
                
                yield return row;
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
    
    private static int CalculateTotalRowsNeeded(List<T> items, CompiledMapping<T> mapping, OptimizedMappingBoundaries boundaries)
    {
        // Use pre-calculated pattern height if available
        if (boundaries.IsMultiItemPattern && boundaries.PatternHeight > 0 && items.Count > 1)
        {
            // Use the pre-calculated pattern height for efficiency
            var totalRows = boundaries.MinRow - 1; // Start counting from before first data row
            
            foreach (var item in items)
            {
                if (item == null) continue;
                
                // Each item starts where the previous one ended
                var itemStartRow = totalRows + 1;
                var itemMaxRow = itemStartRow;
                
                // Use pre-calculated pattern height, but also check actual collection sizes
                // to ensure we have enough space
                foreach (var expansion in mapping.CollectionExpansions ?? [])
                {
                    var collection = expansion.CollectionMapping.Getter(item);
                    if (collection != null)
                    {
                        var collectionSize = collection.Cast<object>().Count();
                        if (collectionSize > 0)
                        {
                            // Adjust collection start row relative to this item's start
                            var collStartRow = itemStartRow + (expansion.StartRow - boundaries.MinRow);
                            var collectionMaxRow = collStartRow + (collectionSize - 1) * (1 + expansion.RowSpacing);
                            itemMaxRow = Math.Max(itemMaxRow, collectionMaxRow);
                        }
                    }
                }
                
                // Use at least the pattern height to maintain consistent spacing
                itemMaxRow = Math.Max(itemMaxRow, itemStartRow + boundaries.PatternHeight - 1);
                totalRows = itemMaxRow;
            }
            
            return totalRows;
        }
        // For scenarios with multiple items and collections without pre-calculated pattern
        else if (items.Count > 1 && mapping.Collections.Any())
        {
            // Multiple items with collections - each item needs its own space
            var totalRows = boundaries.MinRow - 1; // Start counting from before first data row
            
            foreach (var item in items)
            {
                if (item == null) continue;
                
                // Each item starts where the previous one ended
                var itemStartRow = totalRows + 1;
                var itemMaxRow = itemStartRow;
                
                // Account for properties
                foreach (var prop in mapping.Properties)
                {
                    // Adjust property row relative to this item's start
                    var propRow = itemStartRow + (prop.CellRow - boundaries.MinRow);
                    itemMaxRow = Math.Max(itemMaxRow, propRow);
                }
                
                // Account for collections
                foreach (var expansion in mapping.CollectionExpansions ?? [])
                {
                    var collection = expansion.CollectionMapping.Getter(item);
                    if (collection != null)
                    {
                        var collectionSize = collection.Cast<object>().Count();
                        if (collectionSize > 0)
                        {
                            // Adjust collection start row relative to this item's start
                            var collStartRow = itemStartRow + (expansion.StartRow - boundaries.MinRow);
                            var collectionMaxRow = collStartRow + (collectionSize - 1) * (1 + expansion.RowSpacing);
                            itemMaxRow = Math.Max(itemMaxRow, collectionMaxRow);
                        }
                    }
                }
                
                totalRows = itemMaxRow;
            }
            
            return totalRows;
        }
        else
        {
            // Single item or no collections - use original logic
            var maxRow = 0;
            
            // Consider fixed properties
            foreach (var prop in mapping.Properties)
            {
                maxRow = Math.Max(maxRow, prop.CellRow);
            }

            // Calculate expanded rows based on actual collection sizes
            foreach (var item in items)
            {
                if (item == null) continue;

                foreach (var expansion in mapping.CollectionExpansions ?? [])
                {
                    var collection = expansion.CollectionMapping.Getter(item);
                    if (collection != null)
                    {
                        var collectionSize = collection.Cast<object>().Count();
                        if (collectionSize > 0)
                        {
                            var collectionMaxRow = CalculateCollectionMaxRow(expansion, collectionSize);
                            maxRow = Math.Max(maxRow, collectionMaxRow);
                        }
                    }
                }
            }

            return maxRow;
        }
    }

    private static int CalculateCollectionMaxRow(CollectionExpansionInfo expansion, int collectionSize)
    {
        // Only support vertical collections
        if (expansion.Layout == CollectionLayout.Vertical)
        {
            return expansion.StartRow + (collectionSize - 1) * (1 + expansion.RowSpacing);
        }

        return expansion.StartRow;
    }

    private static Dictionary<string, object> CreateRowForAbsoluteRow(int absoluteRow, List<T> items, 
        CompiledMapping<T> mapping, OptimizedCellHandler[,] cellGrid, OptimizedMappingBoundaries boundaries)
    {
        var row = new Dictionary<string, object>();
        
        // For multiple items with collections, determine which item this row belongs to
        int itemIndex = 0;
        int itemStartRow = boundaries.MinRow;
        
        if (boundaries.IsMultiItemPattern && boundaries.PatternHeight > 0 && items.Count > 1)
        {
            // Use pre-calculated pattern to determine item - ZERO runtime calculation!
            var relativeRowForPattern = absoluteRow - boundaries.MinRow;
            
            // Pre-calculated: which item does this belong to?
            // But we need to account for actual collection sizes which may vary
            var currentRow = boundaries.MinRow;
            
            for (int i = 0; i < items.Count; i++)
            {
                if (items[i] == null) continue;
                
                // Calculate this item's actual row span (may differ from pattern if collections vary)
                var itemEndRow = currentRow + boundaries.PatternHeight - 1;
                
                // Check actual collection sizes to ensure we have the right boundaries
                foreach (var expansion in mapping.CollectionExpansions ?? [])
                {
                    var collection = expansion.CollectionMapping.Getter(items[i]);
                    if (collection != null)
                    {
                        var collectionSize = collection.Cast<object>().Count();
                        if (collectionSize > 0)
                        {
                            var collStartRow = currentRow + (expansion.StartRow - boundaries.MinRow);
                            var collectionMaxRow = collStartRow + (collectionSize - 1) * (1 + expansion.RowSpacing);
                            itemEndRow = Math.Max(itemEndRow, collectionMaxRow);
                        }
                    }
                }
                
                if (absoluteRow >= currentRow && absoluteRow <= itemEndRow)
                {
                    itemIndex = i;
                    itemStartRow = currentRow;
                    break;
                }
                
                currentRow = itemEndRow + 1;
            }
        }
        else if (items.Count > 1 && mapping.Collections.Any())
        {
            // Fallback for non-pattern scenarios
            var currentRow = boundaries.MinRow;
            
            for (int i = 0; i < items.Count; i++)
            {
                if (items[i] == null) continue;
                
                // Calculate this item's row span
                var itemEndRow = currentRow;
                
                // Account for properties
                foreach (var prop in mapping.Properties)
                {
                    var propRow = currentRow + (prop.CellRow - boundaries.MinRow);
                    itemEndRow = Math.Max(itemEndRow, propRow);
                }
                
                // Account for collections
                foreach (var expansion in mapping.CollectionExpansions ?? [])
                {
                    var collection = expansion.CollectionMapping.Getter(items[i]);
                    if (collection != null)
                    {
                        var collectionSize = collection.Cast<object>().Count();
                        if (collectionSize > 0)
                        {
                            var collStartRow = currentRow + (expansion.StartRow - boundaries.MinRow);
                            var collectionMaxRow = collStartRow + (collectionSize - 1) * (1 + expansion.RowSpacing);
                            itemEndRow = Math.Max(itemEndRow, collectionMaxRow);
                        }
                    }
                }
                
                if (absoluteRow >= currentRow && absoluteRow <= itemEndRow)
                {
                    itemIndex = i;
                    itemStartRow = currentRow;
                    break;
                }
                
                currentRow = itemEndRow + 1;
            }
        }
        
        // Calculate relative row within the item's space
        var relativeRow = absoluteRow - itemStartRow;
        
        // Map to grid row (original mapping space)
        var gridRow = relativeRow;
        
        // If row is beyond our pre-calculated grid, use the pattern from the last grid row
        // This allows unlimited data without runtime overhead
        if (gridRow >= cellGrid.GetLength(0))
        {
            // For collections that extend beyond our grid, repeat the last row's pattern
            // This is pre-calculated with zero runtime parsing
            gridRow = cellGrid.GetLength(0) - 1;
        }
        
        if (gridRow < 0)
        {
            return row;
        }

        // Initialize columns up to the last actual data column
        // This ensures proper column spacing
        int maxDataCol = 0;
        for (int relativeCol = 0; relativeCol < cellGrid.GetLength(1); relativeCol++)
        {
            var handler = cellGrid[gridRow, relativeCol];
            if (handler.Type != CellHandlerType.Empty)
            {
                maxDataCol = relativeCol + boundaries.MinColumn;
            }
        }
        
        // Initialize ALL columns from A to the maximum column in boundaries
        // This ensures all rows have the same columns, which is required by OpenXmlWriter
        for (int col = 1; col <= boundaries.MaxColumn; col++)
        {
            var columnLetter = OpenXml.Utils.ReferenceHelper.GetCellLetter(
                OpenXml.Utils.ReferenceHelper.ConvertCoordinatesToCell(col, 1));
            row[columnLetter] = "";
        }

        // Process each column in the row using the pre-calculated cell grid
        for (int relativeCol = 0; relativeCol < cellGrid.GetLength(1); relativeCol++)
        {
            var cellHandler = cellGrid[gridRow, relativeCol];
            if (cellHandler.Type == CellHandlerType.Empty)
                continue;

            var absoluteCol = relativeCol + boundaries.MinColumn;
            
            // Get just the column letter for the dictionary key
            var columnLetter = OpenXml.Utils.ReferenceHelper.GetCellLetter(
                OpenXml.Utils.ReferenceHelper.ConvertCoordinatesToCell(absoluteCol, 1));

            object? cellValue = null;

            switch (cellHandler.Type)
            {
                case CellHandlerType.Property:
                    // Use the item that this row belongs to
                    if (itemIndex >= 0 && itemIndex < items.Count && items[itemIndex] != null)
                    {
                        if (cellHandler.ValueExtractor != null)
                        {
                            cellValue = cellHandler.ValueExtractor(items[itemIndex], itemIndex);
                        }
                    }
                    break;

                case CellHandlerType.CollectionItem:
                    // Collection item - extract from the correct item
                    if (itemIndex >= 0 && itemIndex < items.Count && items[itemIndex] != null)
                    {
                        cellValue = ExtractCollectionItemValueForItem(items[itemIndex], cellHandler, absoluteRow - itemStartRow + boundaries.MinRow, absoluteCol, boundaries);
                    }
                    break;

                case CellHandlerType.Formula:
                    // Formula - set the formula directly
                    cellValue = cellHandler.Formula;
                    break;
            }

            if (cellValue != null)
            {
                // Type conversion is built into the ValueExtractor
                
                // Apply formatting if specified
                if (!string.IsNullOrEmpty(cellHandler.Format) && cellValue is IFormattable formattable)
                {
                    cellValue = formattable.ToString(cellHandler.Format, null);
                }

                row[columnLetter] = cellValue;
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