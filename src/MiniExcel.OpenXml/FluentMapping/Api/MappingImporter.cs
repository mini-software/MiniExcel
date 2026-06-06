// ReSharper disable once CheckNamespace
namespace MiniExcelLib.OpenXml.FluentMapping;

public sealed partial class MappingImporter()
{
    private readonly MappingRegistry _registry = new();

    public MappingImporter(MappingRegistry registry) : this()
    {
        _registry = registry ?? throw new ArgumentNullException(nameof(registry));
    }

    [CreateSyncVersion]
    public async IAsyncEnumerable<T> QueryAsync<T>(string path, [EnumeratorCancellation] CancellationToken cancellationToken = default) where T : class, new()
    {
        var stream = File.OpenRead(path);
        await using var disposableStream = stream.ConfigureAwait(false);

        await foreach (var item in QueryAsync<T>(stream, false, cancellationToken).ConfigureAwait(false))
            yield return item;
    }

    [CreateSyncVersion]
    public async IAsyncEnumerable<T> QueryAsync<T>(Stream? stream, bool leaveOpen = false, [EnumeratorCancellation] CancellationToken cancellationToken = default) where T : class, new()
    {
        if (stream is null)
            throw new ArgumentNullException(nameof(stream));

        if (_registry.GetCompiledMapping<T>() is not { } mapping)
            throw new InvalidOperationException($"No mapping configuration found for type {typeof(T).Name}. Configure the mapping using MappingRegistry.Configure<{typeof(T).Name}>().");

        await foreach (var item in MappingReader<T>.QueryAsync(stream, mapping, leaveOpen, cancellationToken).ConfigureAwait(false))
            yield return item;
    }
    
    [CreateSyncVersion]
    public async Task<T> QuerySingleAsync<T>(string path, CancellationToken cancellationToken = default) where T : class, new()
    {
        var stream = File.OpenRead(path);
        await using var disposableStream = stream.ConfigureAwait(false);

        return await QuerySingleAsync<T>(stream, false, cancellationToken).ConfigureAwait(false);
    }

    [CreateSyncVersion]
    private async Task<T> QuerySingleAsync<T>(Stream? stream, bool leaveOpen = false, CancellationToken cancellationToken = default) where T : class, new()
    {
        if (stream is null)
            throw new ArgumentNullException(nameof(stream));

        if (_registry.GetCompiledMapping<T>() is not { }  mapping)
            throw new InvalidOperationException($"No mapping configuration found for type {typeof(T).Name}. Configure the mapping using MappingRegistry.Configure<{typeof(T).Name}>().");

        await foreach (var item in MappingReader<T>.QueryAsync(stream, mapping, leaveOpen, cancellationToken).ConfigureAwait(false))
        {
            return item; // Return the first item
        }
        
        throw new InvalidOperationException("No data found.");
    }
}
