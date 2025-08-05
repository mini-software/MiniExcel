namespace MiniExcelLib.Core.Mapping;

public partial class MappingExporter
{
    private readonly MappingRegistry _registry;

    public MappingExporter() 
    {
        _registry = new MappingRegistry();
    }

    public MappingExporter(MappingRegistry registry)
    {
        _registry = registry ?? throw new ArgumentNullException(nameof(registry));
    }

    [CreateSyncVersion]
    public async Task SaveAsAsync<T>(Stream stream, IEnumerable<T> values, CancellationToken cancellationToken = default)
    {
        if (stream == null)
            throw new ArgumentNullException(nameof(stream));
        if (values == null)
            throw new ArgumentNullException(nameof(values));

        if (!_registry.HasMapping<T>())
            throw new InvalidOperationException($"No mapping configured for type {typeof(T).Name}. Call Configure<{typeof(T).Name}>() first.");

        var mapping = _registry.GetMapping<T>();
        
        await MappingWriter<T>.SaveAsAsync(stream, values, mapping, cancellationToken).ConfigureAwait(false);
    }
    

    public MappingRegistry Registry => _registry;
}