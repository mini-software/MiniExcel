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
    public async Task SaveAsAsync<T>(Stream? stream, IEnumerable<T>? values, CancellationToken cancellationToken = default)
        where T : class
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
    
    [CreateSyncVersion]
    public async Task ApplyTemplateAsync<T>(
        string? outputPath,
        string? templatePath,
        IEnumerable<T>? values,
        CancellationToken cancellationToken = default) where T : class
    {
        if (string.IsNullOrEmpty(outputPath))
            throw new ArgumentException("Output path cannot be null or empty", nameof(outputPath));
        if (string.IsNullOrEmpty(templatePath))
            throw new ArgumentException("Template path cannot be null or empty", nameof(templatePath));
        if (values is null)
            throw new ArgumentNullException(nameof(values));
        
        using var outputStream = File.Create(outputPath);
        using var templateStream = File.OpenRead(templatePath);
        await ApplyTemplateAsync(outputStream, templateStream, values, cancellationToken).ConfigureAwait(false);
    }
    
    [CreateSyncVersion]
    public async Task ApplyTemplateAsync<T>(
        Stream? outputStream,
        Stream? templateStream,
        IEnumerable<T>? values,
        CancellationToken cancellationToken = default) where T : class
    {
        if (outputStream is null)
            throw new ArgumentNullException(nameof(outputStream));
        if (templateStream is null)
            throw new ArgumentNullException(nameof(templateStream));
        if (values is null)
            throw new ArgumentNullException(nameof(values));
        
        if (!_registry.HasMapping<T>())
            throw new InvalidOperationException($"No mapping configured for type {typeof(T).Name}. Call Configure<{typeof(T).Name}>() first.");
        
        var mapping = _registry.GetMapping<T>();
        await MappingTemplateApplicator<T>.ApplyTemplateAsync(
            outputStream, templateStream, values, mapping, cancellationToken).ConfigureAwait(false);
    }
    
    [CreateSyncVersion]
    public async Task ApplyTemplateAsync<T>(
        Stream? outputStream,
        byte[]? templateBytes,
        IEnumerable<T>? values,
        CancellationToken cancellationToken = default) where T : class
    {
        if (outputStream is null)
            throw new ArgumentNullException(nameof(outputStream));
        if (templateBytes is null)
            throw new ArgumentNullException(nameof(templateBytes));
        if (values is null)
            throw new ArgumentNullException(nameof(values));
        
        using var templateStream = new MemoryStream(templateBytes);
        await ApplyTemplateAsync(outputStream, templateStream, values, cancellationToken).ConfigureAwait(false);
    }
}