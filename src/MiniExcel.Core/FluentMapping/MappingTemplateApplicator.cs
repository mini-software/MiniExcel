namespace MiniExcelLib.Core.FluentMapping;

internal static partial class MappingTemplateApplicator<T> where T : class
{
    [CreateSyncVersion]
    public static async Task ApplyTemplateAsync(
        Stream outputStream,
        Stream templateStream,
        IEnumerable<T> values,
        CompiledMapping<T> mapping,
        CancellationToken cancellationToken = default)
    {
        if (outputStream is null)
            throw new ArgumentNullException(nameof(outputStream));
        if (templateStream is null)
            throw new ArgumentNullException(nameof(templateStream));
        if (values is null)
            throw new ArgumentNullException(nameof(values));
        if (mapping is null)
            throw new ArgumentNullException(nameof(mapping));
        
        // Ensure we can seek the template stream
        if (!templateStream.CanSeek)
        {
            // Copy to memory stream if not seekable
            var memStream = new MemoryStream();
#if NETCOREAPP2_1_OR_GREATER
            await templateStream.CopyToAsync(memStream, cancellationToken).ConfigureAwait(false);
#else
            await templateStream.CopyToAsync(memStream).ConfigureAwait(false);
#endif
            memStream.Position = 0;
            templateStream = memStream;
        }
        
        templateStream.Position = 0;
        
        // Open template archive for reading
        using var templateArchive = new ZipArchive(templateStream, ZipArchiveMode.Read, leaveOpen: true);
        
        // Create output archive
        using var outputArchive = new ZipArchive(outputStream, ZipArchiveMode.Create, leaveOpen: true);
        
        // Process each entry
        foreach (var entry in templateArchive.Entries)
        {
            cancellationToken.ThrowIfCancellationRequested();
            
            if (IsWorksheetEntry(entry.FullName))
            {
                // Get worksheet name from path (e.g., "xl/worksheets/sheet1.xml" -> "sheet1")
                var worksheetName = GetWorksheetName(entry.FullName);
                
                // Check if this worksheet matches the mapping's worksheet
                if (mapping.WorksheetName is null || 
                    string.Equals(worksheetName, mapping.WorksheetName, StringComparison.OrdinalIgnoreCase) ||
                    (mapping.WorksheetName == "Sheet1" && worksheetName == "sheet1"))
                {
                    // Process this worksheet with mapping
                    await ProcessWorksheetAsync(
                        entry, 
                        outputArchive, 
                        values, 
                        mapping, 
                        cancellationToken).ConfigureAwait(false);
                }
                else
                {
                    // Copy worksheet as-is
                    await CopyEntryAsync(entry, outputArchive, cancellationToken).ConfigureAwait(false);
                }
            }
            else
            {
                // Copy non-worksheet files as-is
                await CopyEntryAsync(entry, outputArchive, cancellationToken).ConfigureAwait(false);
            }
        }
    }
    
    private static bool IsWorksheetEntry(string fullName)
    {
        return fullName.StartsWith("xl/worksheets/sheet", StringComparison.OrdinalIgnoreCase) &&
               fullName.EndsWith(".xml", StringComparison.OrdinalIgnoreCase);
    }
    
    private static string GetWorksheetName(string fullName)
    {
        // Extract "sheet1" from "xl/worksheets/sheet1.xml"
        var fileName = Path.GetFileNameWithoutExtension(fullName);
        return fileName;
    }
    
    [CreateSyncVersion]
    private static async Task CopyEntryAsync(
        ZipArchiveEntry sourceEntry,
        ZipArchive targetArchive,
        CancellationToken cancellationToken)
    {
        var targetEntry = targetArchive.CreateEntry(sourceEntry.FullName, CompressionLevel.Fastest);
        
        // Copy metadata
        targetEntry.LastWriteTime = sourceEntry.LastWriteTime;
        
        // Copy content
#if NET10_0_OR_GREATER
        using var sourceStream = await sourceEntry.OpenAsync(cancellationToken).ConfigureAwait(false);
        using var targetStream = await targetEntry.OpenAsync(cancellationToken).ConfigureAwait(false);
#else
        using var sourceStream = sourceEntry.Open();
        using var targetStream = targetEntry.Open();
#endif
        
#if NETCOREAPP2_1_OR_GREATER
        await sourceStream.CopyToAsync(targetStream, cancellationToken).ConfigureAwait(false);
#else
        await sourceStream.CopyToAsync(targetStream).ConfigureAwait(false);
#endif
    }
    
    [CreateSyncVersion]
    private static async Task ProcessWorksheetAsync(
        ZipArchiveEntry sourceEntry,
        ZipArchive targetArchive,
        IEnumerable<T> values,
        CompiledMapping<T> mapping,
        CancellationToken cancellationToken)
    {
        var targetEntry = targetArchive.CreateEntry(sourceEntry.FullName, CompressionLevel.Fastest);
        
        // Copy metadata
        targetEntry.LastWriteTime = sourceEntry.LastWriteTime;
        
        // Open streams
#if NET10_0_OR_GREATER
        using var sourceStream = await sourceEntry.OpenAsync(cancellationToken).ConfigureAwait(false);
        using var targetStream = await targetEntry.OpenAsync(cancellationToken).ConfigureAwait(false);
#else
        using var sourceStream = sourceEntry.Open();
        using var targetStream = targetEntry.Open();
#endif
        
        // Create processor for this worksheet
        var processor = new MappingTemplateProcessor<T>(mapping);
        
        // Use enumerator for values
        using var enumerator = values.GetEnumerator();
        
        // Process the worksheet
        await processor.ProcessSheetAsync(
            sourceStream,
            targetStream,
            enumerator,
            cancellationToken).ConfigureAwait(false);
    }
}