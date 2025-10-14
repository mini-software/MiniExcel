namespace MiniExcelLib.Core.Mapping;

public partial class MappingImporter
{
    private readonly MappingRegistry _registry;

    public MappingImporter() 
    {
        _registry = new MappingRegistry();
    }

    public MappingImporter(MappingRegistry registry)
    {
        _registry = registry ?? throw new ArgumentNullException(nameof(registry));
    }

    [CreateSyncVersion]
    public async IAsyncEnumerable<T> QueryAsync<T>(string path, [EnumeratorCancellation] CancellationToken cancellationToken = default) where T : class, new()
    {
        using var stream = File.OpenRead(path);
        await foreach (var item in QueryAsync<T>(stream, cancellationToken).ConfigureAwait(false))
            yield return item;
    }

    [CreateSyncVersion]
    public async IAsyncEnumerable<T> QueryAsync<T>(Stream stream, [EnumeratorCancellation] CancellationToken cancellationToken = default) where T : class, new()
    {
        if (stream == null)
            throw new ArgumentNullException(nameof(stream));

        var mapping = _registry.GetCompiledMapping<T>();
        if (mapping == null)
            throw new InvalidOperationException($"No mapping configuration found for type {typeof(T).Name}. Configure the mapping using MappingRegistry.Configure<{typeof(T).Name}>().");

        await foreach (var item in MappingReader<T>.QueryAsync(stream, mapping, cancellationToken).ConfigureAwait(false))
            yield return item;
    }
    
    [CreateSyncVersion]
    public async Task<T> QuerySingleAsync<T>(string path, CancellationToken cancellationToken = default) where T : class, new()
    {
        using var stream = File.OpenRead(path);
        return await QuerySingleAsync<T>(stream, cancellationToken).ConfigureAwait(false);
    }

    [CreateSyncVersion]
    private async Task<T> QuerySingleAsync<T>(Stream stream, CancellationToken cancellationToken = default) where T : class, new()
    {
        if (stream == null)
            throw new ArgumentNullException(nameof(stream));

        var mapping = _registry.GetCompiledMapping<T>();
        if (mapping == null)
            throw new InvalidOperationException($"No mapping configuration found for type {typeof(T).Name}. Configure the mapping using MappingRegistry.Configure<{typeof(T).Name}>().");

        await foreach (var item in MappingReader<T>.QueryAsync(stream, mapping, cancellationToken).ConfigureAwait(false))
        {
            return item; // Return the first item
        }
        
        throw new InvalidOperationException("No data found.");
    }
}