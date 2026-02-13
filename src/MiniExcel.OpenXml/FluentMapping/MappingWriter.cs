namespace MiniExcelLib.OpenXml.FluentMapping;

internal static partial class MappingWriter<T> where T : class
{
    [CreateSyncVersion]
    public static async Task<int[]> SaveAsAsync(Stream stream, IEnumerable<T> value, CompiledMapping<T> mapping, CancellationToken cancellationToken = default)
    {
        if (stream is null)
            throw new ArgumentNullException(nameof(stream));
        if (value is null)
            throw new ArgumentNullException(nameof(value));
        if (mapping is null)
            throw new ArgumentNullException(nameof(mapping));

        return await SaveAsOptimizedAsync(stream, value, mapping, cancellationToken).ConfigureAwait(false);
    }

    [CreateSyncVersion]
    private static async Task<int[]> SaveAsOptimizedAsync(Stream stream, IEnumerable<T> value, CompiledMapping<T> mapping, CancellationToken cancellationToken = default)
    {
        if (mapping.OptimizedCellGrid is null || mapping.OptimizedBoundaries is null)
            throw new InvalidOperationException("SaveAsOptimizedAsync requires an optimized mapping");

        var configuration = new OpenXmlConfiguration { FastMode = false };
        
        // Pre-calculate column letters once for all cells
        var boundaries = mapping.OptimizedBoundaries;
        var columnLetters = new string[boundaries.MaxColumn - boundaries.MinColumn + 1];
        for (int i = 0; i < columnLetters.Length; i++)
        {
            columnLetters[i] = CellReferenceConverter.GetAlphabeticalIndex(boundaries.MinColumn + i);
        }
        
        // Create cell stream instead of dictionary rows
        var cellStream = new MappingCellStream<T>(value, mapping, columnLetters);
        
        // Use the cell stream directly - it will be handled by the adapter
        var writer = await OpenXmlWriter
            .CreateAsync(stream, cellStream, mapping.WorksheetName, false, configuration, cancellationToken)
            .ConfigureAwait(false);
        
        return await writer.SaveAsAsync(null, cancellationToken).ConfigureAwait(false);
    }
}
