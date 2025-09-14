namespace MiniExcelLib.Core.Mapping;

internal partial struct MappingTemplateProcessor<T>(CompiledMapping<T> mapping) where T : class
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
        var currentItemIndex = currentItem is not null ? 0 : -1;
        
        
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
                        if (mapping.OptimizedBoundaries is { IsMultiItemPattern: true, PatternHeight: > 0 })
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
                        await ProcessRowAsync(reader, writer, rowNumber, currentItem).ConfigureAwait(false);
                    }
                    else if (reader.LocalName is "worksheet" or "sheetData")
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
                if (reader is { NodeType: XmlNodeType.Element, LocalName: "c" })
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
                            
                            // Check if we have a handler for this cell
                            if (mapping.TryGetHandler(row, col, out var handler))
                            {
                                // Use the pre-calculated handler to extract the value
                                if (mapping.TryGetValue(handler, currentItem, out var value))
                                {
                                    // Special handling for collection items
                                    if (handler.Type == CellHandlerType.CollectionItem && value is null)
                                    {
                                        // IMPORTANT: If collection item is null (beyond collection bounds),
                                        // preserve template content instead of overwriting with null
                                        // Skip this cell to preserve template content
                                    }
                                    else
                                    {
                                        // Write the mapped value using centralized helper
                                        await XmlCellWriter.WriteMappedCellAsync(reader, writer, value).ConfigureAwait(false);
                                        cellHandled = true;
                                    }
                                }
                                else if (handler.Type == CellHandlerType.Property)
                                {
                                    // Property with no value - write null using centralized helper
                                    await XmlCellWriter.WriteMappedCellAsync(reader, writer, null).ConfigureAwait(false);
                                    cellHandled = true;
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
                else if (reader is { NodeType: XmlNodeType.EndElement, LocalName: "row" })
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
        if (mapping.OptimizedCellGrid is null || mapping.OptimizedBoundaries is null)
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
                var actualCol = relCol + mapping.OptimizedBoundaries.MinColumn;
                if (mapping.TryGetHandler(actualRow, actualCol, out var handler))
                {
                    hasMapping = true;
                    // Check if there's an actual value to write
                    if (mapping.TryGetValue(handler, currentItem, out var value) && value is not null)
                    {
                        hasValue = true;
                        break;
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
        if (mapping.OptimizedBoundaries is not null)
        {
            for (int col = mapping.OptimizedBoundaries.MinColumn; col <= mapping.OptimizedBoundaries.MaxColumn; col++)
            {
                // Check if we have a handler for this cell
                if (mapping.TryGetHandler(rowNumber, col, out var handler))
                {
                    // Try to get the value
                    if (mapping.TryGetValue(handler, currentItem, out var value) && value is not null)
                    {
                        var cellRef = ReferenceHelper.ConvertCoordinatesToCell(col, rowNumber);
                        await XmlCellWriter.WriteNewCellAsync(writer, cellRef, value).ConfigureAwait(false);
                    }
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
        if (mapping.OptimizedBoundaries is not null)
        {
            // Check each column in the grid for this row
            for (int col = mapping.OptimizedBoundaries.MinColumn; col <= mapping.OptimizedBoundaries.MaxColumn; col++)
            {
                // Skip if we already wrote this column
                if (writtenColumns.Contains(col))
                    continue;
                
                // Check if we have a handler for this cell
                if (mapping.TryGetHandler(rowNumber, col, out var handler))
                {
                    // We have a mapping for this cell but it wasn't in the template
                    // Try to get the value
                    if (mapping.TryGetValue(handler, currentItem, out var value) && value is not null)
                    {
                        // Create cell reference
                        var cellRef = ReferenceHelper.ConvertCoordinatesToCell(col, rowNumber);
                        
                        // Write the cell using centralized helper
                        await XmlCellWriter.WriteNewCellAsync(writer, cellRef, value).ConfigureAwait(false);
                    }
                }
            }
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