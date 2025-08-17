namespace MiniExcelLib.Core.Mapping;

internal partial struct MappingTemplateProcessor<T>(CompiledMapping<T> mapping)
    where T : class
{
    [CreateSyncVersion]
    public async Task ProcessSheetAsync(
        Stream sourceStream,
        Stream targetStream,
        IEnumerator<T> dataEnumerator,
        CancellationToken cancellationToken)
    {
        var readerSettings = new XmlReaderSettings
        {
            Async = true,
            IgnoreWhitespace = false,
            IgnoreComments = false,
            CheckCharacters = false
        };
        
        var writerSettings = new XmlWriterSettings
        {
            Async = true,
            Indent = false,
            OmitXmlDeclaration = false,
            Encoding = Encoding.UTF8
        };
        
        using var reader = XmlReader.Create(sourceStream, readerSettings);
        using var writer = XmlWriter.Create(targetStream, writerSettings);
        
        // Get first data item
        var currentItem = dataEnumerator.MoveNext() ? dataEnumerator.Current : null;
        var currentItemIndex = currentItem != null ? 0 : -1;
        
        
        // Track which rows have been written from the template
        var writtenRows = new HashSet<int>();
        
        // Process the XML stream
        while (await reader.ReadAsync().ConfigureAwait(false))
        {
            cancellationToken.ThrowIfCancellationRequested();
            
            switch (reader.NodeType)
            {
                case XmlNodeType.XmlDeclaration:
                    await writer.WriteStartDocumentAsync().ConfigureAwait(false);
                    break;
                    
                case XmlNodeType.Element:
                    if (reader.LocalName == "row")
                    {
                        var rowNumber = GetRowNumber(reader);
                        writtenRows.Add(rowNumber);
                        
                        // Check if we need to advance to next item
                        if (mapping.OptimizedBoundaries != null && 
                            mapping.OptimizedBoundaries.IsMultiItemPattern &&
                            mapping.OptimizedBoundaries.PatternHeight > 0)
                        {
                            var relativeRow = rowNumber - mapping.OptimizedBoundaries.MinRow;
                            var itemIndex = relativeRow / mapping.OptimizedBoundaries.PatternHeight;
                            
                            if (itemIndex > currentItemIndex)
                            {
                                // Advance to next item
                                currentItem = dataEnumerator.MoveNext() ? dataEnumerator.Current : null;
                                currentItemIndex = itemIndex;
                            }
                        }
                        
                        // Process the row
                        await ProcessRowAsync(
                            reader, writer, rowNumber, 
                            currentItem).ConfigureAwait(false);
                    }
                    else if (reader.LocalName == "worksheet" || reader.LocalName == "sheetData")
                    {
                        // For worksheet and sheetData elements, we need to process their content manually
                        // Copy start tag with attributes
                        await writer.WriteStartElementAsync(reader.Prefix, reader.LocalName, reader.NamespaceURI).ConfigureAwait(false);
                        
                        if (reader.HasAttributes)
                        {
                            while (reader.MoveToNextAttribute())
                            {
                                await writer.WriteAttributeStringAsync(
                                    reader.Prefix,
                                    reader.LocalName,
                                    reader.NamespaceURI,
                                    reader.Value).ConfigureAwait(false);
                            }
                            reader.MoveToElement();
                        }
                        
                        // Don't call CopyElementAsync as it will consume all content
                        // Just continue processing in the main loop
                    }
                    else
                    {
                        // Copy element as-is
                        await CopyElementAsync(reader, writer).ConfigureAwait(false);
                    }
                    break;
                    
                case XmlNodeType.EndElement:
                    if (reader.LocalName == "sheetData")
                    {
                        // Before closing sheetData, write any missing rows that have mappings
                        await WriteMissingRowsAsync(writer, currentItem, writtenRows).ConfigureAwait(false);
                    }
                    await writer.WriteEndElementAsync().ConfigureAwait(false);
                    break;
                    
                default:
                    // Copy node as-is
                    await CopyNodeAsync(reader, writer).ConfigureAwait(false);
                    break;
            }
        }
        
        await writer.FlushAsync().ConfigureAwait(false);
    }
    
    private static int GetRowNumber(XmlReader reader)
    {
        var rowAttr = reader.GetAttribute("r");
        if (!string.IsNullOrEmpty(rowAttr) && int.TryParse(rowAttr, out var rowNum))
        {
            return rowNum;
        }
        return 0;
    }
    
    [CreateSyncVersion]
    private async Task ProcessRowAsync(
        XmlReader reader,
        XmlWriter writer,
        int rowNumber,
        T? currentItem)
    {
        // Write row start tag with all attributes
        await writer.WriteStartElementAsync(reader.Prefix, "row", reader.NamespaceURI).ConfigureAwait(false);
        
        // Copy all row attributes
        if (reader.HasAttributes)
        {
            while (reader.MoveToNextAttribute())
            {
                await writer.WriteAttributeStringAsync(
                    reader.Prefix,
                    reader.LocalName,
                    reader.NamespaceURI,
                    reader.Value).ConfigureAwait(false);
            }
            reader.MoveToElement();
        }
        
        // Track which columns have been written
        var writtenColumns = new HashSet<int>();
        
        // Read row content
        var isEmpty = reader.IsEmptyElement;
        if (!isEmpty)
        {
            // Process cells in the row
            while (await reader.ReadAsync().ConfigureAwait(false))
            {
                if (reader.NodeType == XmlNodeType.Element && reader.LocalName == "c")
                {
                    // Get cell reference
                    var cellRef = reader.GetAttribute("r");
                    
                    
                    if (!string.IsNullOrEmpty(cellRef))
                    {
                        // Parse cell reference to get column and row
                        if (ReferenceHelper.TryParseCellReference(cellRef, out var col, out var row))
                        {
                            // Track that we've written this column
                            writtenColumns.Add(col);
                            
                            bool cellHandled = false;
                            
                            // Check if we have an optimized grid and this cell is within bounds
                            if (mapping.OptimizedCellGrid != null && mapping.OptimizedBoundaries != null)
                            {
                                var relRow = row - mapping.OptimizedBoundaries.MinRow;
                                var relCol = col - mapping.OptimizedBoundaries.MinColumn;
                                
                                
                                if (relRow >= 0 && relRow < mapping.OptimizedBoundaries.GridHeight &&
                                    relCol >= 0 && relCol < mapping.OptimizedBoundaries.GridWidth)
                                {
                                    var handler = mapping.OptimizedCellGrid[relRow, relCol];
                                    
                                    
                                    if (handler.Type != CellHandlerType.Empty)
                                    {
                                        // Use the pre-calculated handler to extract the value
                                        object? value = null;
                                        bool skipCell = false;
                                        
                                        if (handler.Type == CellHandlerType.Property && handler.ValueExtractor != null)
                                        {
                                            value = currentItem != null ? handler.ValueExtractor(currentItem, 0) : null;
                                        }
                                        else if (handler.Type == CellHandlerType.CollectionItem && handler.ValueExtractor != null)
                                        {
                                            // For collections, the ValueExtractor is pre-configured with the right offset
                                            // Just pass the parent object that contains the collection
                                            value = currentItem != null ? handler.ValueExtractor(currentItem, 0) : null;
                                            
                                            // IMPORTANT: If collection item is null (beyond collection bounds),
                                            // preserve template content instead of overwriting with null
                                            if (value == null)
                                            {
                                                skipCell = true;
                                            }
                                        }
                                        
                                        
                                        // Only write if we have a value to write
                                        if (!skipCell)
                                        {
                                            await WriteMappedCellAsync(reader, writer, value).ConfigureAwait(false);
                                            cellHandled = true;
                                        }
                                    }
                                }
                            }
                            
                            if (!cellHandled)
                            {
                                // Cell not in grid - just copy as-is from template
                                await CopyElementAsync(reader, writer).ConfigureAwait(false);
                            }
                        }
                        else
                        {
                            // Copy cell as-is if we can't parse the reference
                            await CopyElementAsync(reader, writer).ConfigureAwait(false);
                        }
                    }
                    else
                    {
                        // No cell reference, copy as-is
                        await CopyElementAsync(reader, writer).ConfigureAwait(false);
                    }
                }
                else if (reader.NodeType == XmlNodeType.EndElement && reader.LocalName == "row")
                {
                    break;
                }
                else
                {
                    await CopyNodeAsync(reader, writer).ConfigureAwait(false);
                }
            }
        }
        
        // After processing existing cells, check for missing mapped cells in this row
        await WriteMissingCellsAsync(writer, rowNumber, writtenColumns, currentItem).ConfigureAwait(false);
        
        await writer.WriteEndElementAsync().ConfigureAwait(false);
    }
    
    [CreateSyncVersion]
    private async Task WriteMissingRowsAsync(
        XmlWriter writer,
        T? currentItem,
        HashSet<int> writtenRows)
    {
        // Check if we have an optimized grid with mappings
        if (mapping.OptimizedCellGrid == null || mapping.OptimizedBoundaries == null)
            return;
        
        
        // Check each row in the grid to see if it has mappings but wasn't written
        for (int relRow = 0; relRow < mapping.OptimizedBoundaries.GridHeight; relRow++)
        {
            var actualRow = relRow + mapping.OptimizedBoundaries.MinRow;
            
            // Skip if this row was already written from the template
            if (writtenRows.Contains(actualRow))
                continue;
            
            // Check if this row has any mapped cells with actual values
            bool hasMapping = false;
            bool hasValue = false;
            for (int relCol = 0; relCol < mapping.OptimizedBoundaries.GridWidth; relCol++)
            {
                var handler = mapping.OptimizedCellGrid[relRow, relCol];
                if (handler.Type != CellHandlerType.Empty)
                {
                    hasMapping = true;
                    // Check if there's an actual value to write
                    if (handler.ValueExtractor != null && currentItem != null)
                    {
                        var value = handler.ValueExtractor(currentItem, 0);
                        if (value != null)
                        {
                            hasValue = true;
                            break;
                        }
                    }
                }
            }
            
            if (hasMapping && hasValue)
            {
                // Write this missing row
                await WriteNewRowAsync(writer, actualRow, currentItem).ConfigureAwait(false);
            }
        }
    }
    
    [CreateSyncVersion]
    private async Task WriteNewRowAsync(
        XmlWriter writer,
        int rowNumber,
        T? currentItem)
    {
        // Write row element
        await writer.WriteStartElementAsync("", "row", "").ConfigureAwait(false);
        await writer.WriteAttributeStringAsync("", "r", "", rowNumber.ToString()).ConfigureAwait(false);
        
        // Check each column in this row for mapped cells
        if (mapping.OptimizedCellGrid != null && mapping.OptimizedBoundaries != null)
        {
            var relRow = rowNumber - mapping.OptimizedBoundaries.MinRow;
            
            if (relRow >= 0 && relRow < mapping.OptimizedBoundaries.GridHeight)
            {
                for (int relCol = 0; relCol < mapping.OptimizedBoundaries.GridWidth; relCol++)
                {
                    var handler = mapping.OptimizedCellGrid[relRow, relCol];
                    if (handler.Type == CellHandlerType.Empty) continue;
                    
                    // Extract the value
                    object? value = null;
                    
                    if (handler.Type == CellHandlerType.Property && handler.ValueExtractor != null)
                    {
                        value = currentItem != null ? handler.ValueExtractor(currentItem, 0) : null;
                    }
                    else if (handler.Type == CellHandlerType.CollectionItem && handler.ValueExtractor != null)
                    {
                        value = currentItem != null ? handler.ValueExtractor(currentItem, 0) : null;
                    }

                    if (value == null) continue;
                    
                    var actualCol = relCol + mapping.OptimizedBoundaries.MinColumn;
                    var cellRef = ReferenceHelper.ConvertCoordinatesToCell(actualCol, rowNumber);
                    await WriteNewCellAsync(writer, cellRef, value).ConfigureAwait(false);
                }
            }
        }
        
        await writer.WriteEndElementAsync().ConfigureAwait(false); // </row>
    }
    
    [CreateSyncVersion]
    private async Task WriteMissingCellsAsync(
        XmlWriter writer,
        int rowNumber,
        HashSet<int> writtenColumns,
        T? currentItem)
    {
        
        // Check if we have an optimized grid with mappings for this row
        if (mapping.OptimizedCellGrid != null && mapping.OptimizedBoundaries != null)
        {
            var relRow = rowNumber - mapping.OptimizedBoundaries.MinRow;
            
            if (relRow >= 0 && relRow < mapping.OptimizedBoundaries.GridHeight)
            {
                // Check each column in the grid for this row
                for (int relCol = 0; relCol < mapping.OptimizedBoundaries.GridWidth; relCol++)
                {
                    var actualCol = relCol + mapping.OptimizedBoundaries.MinColumn;
                    
                    // Skip if we already wrote this column
                    if (writtenColumns.Contains(actualCol))
                        continue;
                    
                    var handler = mapping.OptimizedCellGrid[relRow, relCol];
                    if (handler.Type == CellHandlerType.Empty) continue;
                    
                    // We have a mapping for this cell but it wasn't in the template
                    // Create a new cell for it
                    object? value = null;
                        
                    if (handler.Type == CellHandlerType.Property && handler.ValueExtractor != null)
                    {
                        value = currentItem != null ? handler.ValueExtractor(currentItem, 0) : null;
                    }
                    else if (handler.Type == CellHandlerType.CollectionItem && handler.ValueExtractor != null)
                    {
                        value = currentItem != null ? handler.ValueExtractor(currentItem, 0) : null;
                    }
                        
                    if (value != null)
                    {
                        // Create cell reference
                        var cellRef = ReferenceHelper.ConvertCoordinatesToCell(actualCol, rowNumber);
                            
                            
                        // Write the cell
                        await WriteNewCellAsync(writer, cellRef, value).ConfigureAwait(false);
                    }
                }
            }
        }
    }
    
    [CreateSyncVersion]
    private async Task WriteNewCellAsync(
        XmlWriter writer,
        string cellRef,
        object? value)
    {
        // Determine cell type and formatted value
        var (cellValue, cellType) = FormatCellValue(value);
        
        if (string.IsNullOrEmpty(cellValue) && string.IsNullOrEmpty(cellType))
            return; // Don't write empty cells
        
        // Write cell element
        await writer.WriteStartElementAsync("", "c", "").ConfigureAwait(false);
        await writer.WriteAttributeStringAsync("", "r", "", cellRef).ConfigureAwait(false);
        
        if (!string.IsNullOrEmpty(cellType))
        {
            await writer.WriteAttributeStringAsync("", "t", "", cellType).ConfigureAwait(false);
        }
        
        // Write the value
        if (cellType == "inlineStr" && !string.IsNullOrEmpty(cellValue))
        {
            // Write inline string
            await writer.WriteStartElementAsync("", "is", "").ConfigureAwait(false);
            await writer.WriteStartElementAsync("", "t", "").ConfigureAwait(false);
            await writer.WriteStringAsync(cellValue).ConfigureAwait(false);
            await writer.WriteEndElementAsync().ConfigureAwait(false); // </t>
            await writer.WriteEndElementAsync().ConfigureAwait(false); // </is>
        }
        else if (!string.IsNullOrEmpty(cellValue))
        {
            // Write value element
            await writer.WriteStartElementAsync("", "v", "").ConfigureAwait(false);
            await writer.WriteStringAsync(cellValue).ConfigureAwait(false);
            await writer.WriteEndElementAsync().ConfigureAwait(false); // </v>
        }
        
        await writer.WriteEndElementAsync().ConfigureAwait(false); // </c>
    }
    
    [CreateSyncVersion]
    private static async Task WriteMappedCellAsync(
        XmlReader reader,
        XmlWriter writer,
        object? value)
    {
        // Determine cell type and formatted value
        var (cellValue, cellType) = FormatCellValue(value);
        
        // Write cell start tag
        await writer.WriteStartElementAsync(reader.Prefix, "c", reader.NamespaceURI).ConfigureAwait(false);
        
        // Copy attributes, potentially updating type
        if (reader.HasAttributes)
        {
            while (reader.MoveToNextAttribute())
            {
                if (reader.LocalName == "t")
                {
                    // Write our type instead
                    if (!string.IsNullOrEmpty(cellType))
                    {
                        await writer.WriteAttributeStringAsync("", "t", "", cellType).ConfigureAwait(false);
                    }
                }
                else if (reader.LocalName == "s")
                {
                    // Skip style if we're writing inline string
                    if (cellType != "inlineStr")
                    {
                        await writer.WriteAttributeStringAsync(
                            reader.Prefix,
                            reader.LocalName,
                            reader.NamespaceURI,
                            reader.Value).ConfigureAwait(false);
                    }
                }
                else
                {
                    // Copy other attributes
                    await writer.WriteAttributeStringAsync(
                        reader.Prefix,
                        reader.LocalName,
                        reader.NamespaceURI,
                        reader.Value).ConfigureAwait(false);
                }
            }
            reader.MoveToElement();
        }
        
        // If we didn't have a type attribute but need one, add it
        if (!string.IsNullOrEmpty(cellType) && reader.GetAttribute("t") == null)
        {
            await writer.WriteAttributeStringAsync("", "t", "", cellType).ConfigureAwait(false);
        }
        
        // Write the value
        if (cellType == "inlineStr" && !string.IsNullOrEmpty(cellValue))
        {
            // Write inline string
            await writer.WriteStartElementAsync("", "is", reader.NamespaceURI).ConfigureAwait(false);
            await writer.WriteStartElementAsync("", "t", reader.NamespaceURI).ConfigureAwait(false);
            await writer.WriteStringAsync(cellValue).ConfigureAwait(false);
            await writer.WriteEndElementAsync().ConfigureAwait(false); // </t>
            await writer.WriteEndElementAsync().ConfigureAwait(false); // </is>
        }
        else if (!string.IsNullOrEmpty(cellValue))
        {
            // Write value element
            await writer.WriteStartElementAsync("", "v", reader.NamespaceURI).ConfigureAwait(false);
            await writer.WriteStringAsync(cellValue).ConfigureAwait(false);
            await writer.WriteEndElementAsync().ConfigureAwait(false); // </v>
        }
        
        // Skip original cell content
        var isEmpty = reader.IsEmptyElement;
        if (!isEmpty)
        {
            var depth = reader.Depth;
            while (await reader.ReadAsync().ConfigureAwait(false))
            {
                if (reader.NodeType == XmlNodeType.EndElement && reader.Depth == depth)
                {
                    break;
                }
            }
        }
        
        await writer.WriteEndElementAsync().ConfigureAwait(false); // </c>
    }
    
    private static (string? value, string? type) FormatCellValue(object? value)
    {
        if (value == null)
            return (null, null);
        
        switch (value)
        {
            case string s:
                // Use inline string to avoid shared string table
                return (s, "inlineStr");
                
            case DateTime dt:
                // Excel stores dates as numbers
                var excelDate = (dt - new DateTime(1899, 12, 30)).TotalDays;
                return (excelDate.ToString(CultureInfo.InvariantCulture), null);
                
            case DateTimeOffset dto:
                var excelDateOffset = (dto.DateTime - new DateTime(1899, 12, 30)).TotalDays;
                return (excelDateOffset.ToString(CultureInfo.InvariantCulture), null);
                
            case bool b:
                return (b ? "1" : "0", "b");
                
            case byte:
            case sbyte:
            case short:
            case ushort:
            case int:
            case uint:
            case long:
            case ulong:
            case float:
            case double:
            case decimal:
                return (Convert.ToString(value, CultureInfo.InvariantCulture), null);
                
            default:
                // Convert to string
                return (value.ToString(), "inlineStr");
        }
    }
    
    [CreateSyncVersion]
    private static async Task CopyElementAsync(XmlReader reader, XmlWriter writer)
    {
        // Write start element
        await writer.WriteStartElementAsync(reader.Prefix, reader.LocalName, reader.NamespaceURI).ConfigureAwait(false);
        
        // Copy attributes
        if (reader.HasAttributes)
        {
            while (reader.MoveToNextAttribute())
            {
                await writer.WriteAttributeStringAsync(
                    reader.Prefix,
                    reader.LocalName,
                    reader.NamespaceURI,
                    reader.Value).ConfigureAwait(false);
            }
            reader.MoveToElement();
        }
        
        // If empty element, we're done
        if (reader.IsEmptyElement)
        {
            await writer.WriteEndElementAsync().ConfigureAwait(false);
            return;
        }
        
        // Copy content
        var depth = reader.Depth;
        while (await reader.ReadAsync().ConfigureAwait(false))
        {
            if (reader.NodeType == XmlNodeType.EndElement && reader.Depth == depth)
            {
                await writer.WriteEndElementAsync().ConfigureAwait(false);
                break;
            }
            
            await CopyNodeAsync(reader, writer).ConfigureAwait(false);
        }
    }
    
    [CreateSyncVersion]
    private static async Task CopyNodeAsync(XmlReader reader, XmlWriter writer)
    {
        switch (reader.NodeType)
        {
            case XmlNodeType.Element:
                await CopyElementAsync(reader, writer).ConfigureAwait(false);
                break;
                
            case XmlNodeType.Text:
                await writer.WriteStringAsync(reader.Value).ConfigureAwait(false);
                break;
                
            case XmlNodeType.Whitespace:
            case XmlNodeType.SignificantWhitespace:
                await writer.WriteWhitespaceAsync(reader.Value).ConfigureAwait(false);
                break;
                
            case XmlNodeType.CDATA:
                await writer.WriteCDataAsync(reader.Value).ConfigureAwait(false);
                break;
                
            case XmlNodeType.Comment:
                await writer.WriteCommentAsync(reader.Value).ConfigureAwait(false);
                break;
                
            case XmlNodeType.ProcessingInstruction:
                await writer.WriteProcessingInstructionAsync(reader.Name, reader.Value).ConfigureAwait(false);
                break;
                
            case XmlNodeType.EntityReference:
                await writer.WriteEntityRefAsync(reader.Name).ConfigureAwait(false);
                break;
                
            case XmlNodeType.XmlDeclaration:
                // Write the XML declaration properly
                await writer.WriteStartDocumentAsync().ConfigureAwait(false);
                break;
                
            case XmlNodeType.DocumentType:
                await writer.WriteRawAsync(reader.Value).ConfigureAwait(false);
                break;
                
            case XmlNodeType.EndElement:
                await writer.WriteEndElementAsync().ConfigureAwait(false);
                break;
        }
    }
    
}