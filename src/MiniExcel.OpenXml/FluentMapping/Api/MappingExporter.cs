using Zomp.SyncMethodGenerator;

namespace MiniExcelLib.OpenXml.FluentMapping.Api;

public sealed partial class MappingExporter
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
    public async Task ExportAsync<T>(string path, IEnumerable<T>? values, bool overwriteFile = false, CancellationToken cancellationToken = default) where T : class
    {
        var filePath = path.EndsWith(".xlsx",  StringComparison.InvariantCultureIgnoreCase) ? path : $"{path}.xlsx" ;
        
        using var stream = overwriteFile ? File.Create(filePath) : new FileStream(filePath, FileMode.CreateNew);
        await ExportAsync(stream, values, cancellationToken).ConfigureAwait(false);
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