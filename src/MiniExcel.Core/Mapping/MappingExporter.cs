namespace MiniExcelLib.Core.Mapping;

public sealed partial class MappingExporter()
{
    private readonly MappingRegistry _registry = new();

    public MappingExporter(MappingRegistry registry) : this()
    {
        _registry = registry ?? throw new ArgumentNullException(nameof(registry));
    }

    [CreateSyncVersion]
    public async Task ExportAsync<T>(Stream? stream, IEnumerable<T>? values, CancellationToken cancellationToken = default) where T : class
    {
        if (stream is null)
            throw new ArgumentNullException(nameof(stream));
        if (values is null)
            throw new ArgumentNullException(nameof(values));

        if (!_registry.HasMapping<T>())
            throw new InvalidOperationException($"No mapping configured for type {typeof(T).Name}. Call Configure<{typeof(T).Name}>() first.");

        var mapping = _registry.GetMapping<T>();
        
        await MappingWriter<T>.SaveAsAsync(stream, values, mapping, cancellationToken).ConfigureAwait(false);
    }
}