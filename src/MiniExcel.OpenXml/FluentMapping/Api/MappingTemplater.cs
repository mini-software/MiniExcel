namespace MiniExcelLib.OpenXml.FluentMapping.Api;

public sealed partial class MappingTemplater()
{
    private readonly MappingRegistry _registry = new();

    public MappingTemplater(MappingRegistry registry) : this()
    {
        _registry = registry ?? throw new ArgumentNullException(nameof(registry));
    }
    
    [CreateSyncVersion]
    public async Task FillTemplateAsync<T>(
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

#if NET8_0_OR_GREATER
        var outputStream = File.Create(outputPath);
        await using var disposableOutputStream = outputStream.ConfigureAwait(false);
        
        var templateStream = File.OpenRead(templatePath);
        await using var disposableTemplateStream = templateStream.ConfigureAwait(false);
#else
        using var outputStream = File.Create(outputPath);
        using var templateStream = File.OpenRead(templatePath);
#endif
        await FillTemplateAsync(outputStream, templateStream, values, cancellationToken).ConfigureAwait(false);
    }

    [CreateSyncVersion]
    public async Task FillTemplateAsync<T>(
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
            throw new InvalidOperationException(
                $"No mapping configured for type {typeof(T).Name}. Call Configure<{typeof(T).Name}>() first.");

        var mapping = _registry.GetMapping<T>();
        await MappingTemplateApplicator<T>.ApplyTemplateAsync(
            outputStream, templateStream, values, mapping, cancellationToken).ConfigureAwait(false);
    }

    [CreateSyncVersion]
    public async Task FillTemplateAsync<T>(
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

#if NET8_0_OR_GREATER
        var templateStream = new MemoryStream(templateBytes);
        await using var disposableTemplateStream = templateStream.ConfigureAwait(false);
#else
        using var templateStream = new MemoryStream(templateBytes);
#endif
        await FillTemplateAsync(outputStream, templateStream, values, cancellationToken).ConfigureAwait(false);
    }

#region Obsolete
[CreateSyncVersion, Obsolete("Please use FillTemplate or FillTemplateAsync instead.")]
public Task ApplyTemplateAsync<T>(
    string? outputPath,
    string? templatePath,
    IEnumerable<T>? values,
    CancellationToken cancellationToken = default) where T : class
{
    return FillTemplateAsync(outputPath, templatePath, values, cancellationToken);
}

[CreateSyncVersion, Obsolete("Please use FillTemplate or FillTemplateAsync instead.")]
public Task ApplyTemplateAsync<T>(
    Stream? outputStream,
    Stream? templateStream,
    IEnumerable<T>? values,
    CancellationToken cancellationToken = default) where T : class
{
    return FillTemplateAsync(outputStream, templateStream, values, cancellationToken);
}


[CreateSyncVersion, Obsolete("Please use FillTemplate or FillTemplateAsync instead.")]
public Task ApplyTemplateAsync<T>(
    Stream? outputStream,
    byte[]? templateBytes,
    IEnumerable<T>? values,
    CancellationToken cancellationToken = default) where T : class
{
    return FillTemplateAsync(outputStream, templateBytes, values, cancellationToken);
}
#endregion
}
