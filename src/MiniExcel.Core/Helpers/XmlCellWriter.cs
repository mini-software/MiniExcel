namespace MiniExcelLib.Core.Helpers;

/// <summary>
/// Helper class for writing Excel cell XML with consistent formatting.
/// Consolidates XML cell writing patterns to reduce duplication.
/// </summary>
internal static partial class XmlCellWriter
{
    /// <summary>
    /// Writes a new cell element with the specified reference and value.
    /// </summary>
    /// <param name="writer">The XML writer</param>
    /// <param name="cellRef">The cell reference (e.g., "A1")</param>
    /// <param name="value">The cell value</param>
    /// <param name="cancellationToken">Cancellation token</param>
    [CreateSyncVersion]
    public static async Task WriteNewCellAsync(
        XmlWriter writer,
        string cellRef,
        object? value,
        CancellationToken cancellationToken = default)
    {
        // Use centralized formatting
        var (cellValue, cellType) = CellFormatter.FormatCellValue(value);
        
        if (string.IsNullOrEmpty(cellValue) && string.IsNullOrEmpty(cellType))
            return; // Don't write empty cells
        
        // Write cell element
        await writer.WriteStartElementAsync("", "c", "").ConfigureAwait(false);
        await writer.WriteAttributeStringAsync("", "r", "", cellRef).ConfigureAwait(false);
        
        if (!string.IsNullOrEmpty(cellType))
        {
            await writer.WriteAttributeStringAsync("", "t", "", cellType).ConfigureAwait(false);
        }
        
        // Write the value content
        await WriteCellValueContentAsync(writer, cellValue, cellType).ConfigureAwait(false);
        
        await writer.WriteEndElementAsync().ConfigureAwait(false); // </c>
    }

    /// <summary>
    /// Writes a cell element replacing template content with new value.
    /// </summary>
    /// <param name="reader">The XML reader positioned on the cell element</param>
    /// <param name="writer">The XML writer</param>
    /// <param name="value">The new cell value</param>
    /// <param name="cancellationToken">Cancellation token</param>
    [CreateSyncVersion]
    public static async Task WriteMappedCellAsync(
        XmlReader reader,
        XmlWriter writer,
        object? value,
        CancellationToken cancellationToken = default)
    {
        // Use centralized formatting
        var (cellValue, cellType) = CellFormatter.FormatCellValue(value);
        
        // Write cell start tag
        await writer.WriteStartElementAsync(reader.Prefix, "c", reader.NamespaceURI).ConfigureAwait(false);
        
        // Copy attributes, potentially updating type
        await CopyAndUpdateCellAttributesAsync(reader, writer, cellType).ConfigureAwait(false);
        
        // Write the value content
        await WriteCellValueContentAsync(writer, cellValue, cellType, reader.NamespaceURI).ConfigureAwait(false);
        
        // Skip original cell content
        await SkipOriginalCellContentAsync(reader).ConfigureAwait(false);
        
        await writer.WriteEndElementAsync().ConfigureAwait(false); // </c>
    }

    /// <summary>
    /// Writes the value content (v or is elements) for a cell.
    /// </summary>
    [CreateSyncVersion]
    private static async Task WriteCellValueContentAsync(
        XmlWriter writer, 
        string? cellValue, 
        string? cellType, 
        string namespaceUri = "")
    {
        if (cellType == "inlineStr" && !string.IsNullOrEmpty(cellValue))
        {
            // Write inline string
            await writer.WriteStartElementAsync("", "is", namespaceUri).ConfigureAwait(false);
            await writer.WriteStartElementAsync("", "t", namespaceUri).ConfigureAwait(false);
            await writer.WriteStringAsync(cellValue).ConfigureAwait(false);
            await writer.WriteEndElementAsync().ConfigureAwait(false); // </t>
            await writer.WriteEndElementAsync().ConfigureAwait(false); // </is>
        }
        else if (!string.IsNullOrEmpty(cellValue))
        {
            // Write value element
            await writer.WriteStartElementAsync("", "v", namespaceUri).ConfigureAwait(false);
            await writer.WriteStringAsync(cellValue).ConfigureAwait(false);
            await writer.WriteEndElementAsync().ConfigureAwait(false); // </v>
        }
    }

    /// <summary>
    /// Copies cell attributes from reader to writer, updating the type attribute if needed.
    /// </summary>
    [CreateSyncVersion]
    private static async Task CopyAndUpdateCellAttributesAsync(
        XmlReader reader, 
        XmlWriter writer, 
        string? newCellType)
    {
        if (reader.HasAttributes)
        {
            while (reader.MoveToNextAttribute())
            {
                if (reader.LocalName == "t")
                {
                    // Write our type instead
                    if (!string.IsNullOrEmpty(newCellType))
                    {
                        await writer.WriteAttributeStringAsync("", "t", "", newCellType).ConfigureAwait(false);
                    }
                }
                else if (reader.LocalName == "s")
                {
                    // Skip style if we're writing inline string
                    if (newCellType != "inlineStr")
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
        if (!string.IsNullOrEmpty(newCellType) && reader.GetAttribute("t") is null)
        {
            await writer.WriteAttributeStringAsync("", "t", "", newCellType).ConfigureAwait(false);
        }
    }

    /// <summary>
    /// Skips the original cell content when replacing with new content.
    /// </summary>
    [CreateSyncVersion]
    private static async Task SkipOriginalCellContentAsync(XmlReader reader)
    {
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
    }
}