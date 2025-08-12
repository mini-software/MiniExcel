using MiniExcelLib.Core.Mapping.Configuration;

namespace MiniExcelLib.Core.Mapping;

public sealed class MappingRegistry
{
    private readonly Dictionary<Type, object> _compiledMappings = new();

#if NET9_0_OR_GREATER
    private readonly Lock _lock = new();
#else
    private readonly object _lock = new();
#endif

    public void Configure<T>(Action<IMappingConfiguration<T>>? configure)
    {
        if (configure is null)
            throw new ArgumentNullException(nameof(configure));
            
        lock (_lock)
        {
            var config = new MappingConfiguration<T>();
            configure(config);
            
            var compiledMapping = MappingCompiler.Compile(config, this);
            _compiledMappings[typeof(T)] = compiledMapping;
        }
    }
    
    internal CompiledMapping<T> GetMapping<T>()
    {
        lock (_lock)
        {
            return _compiledMappings.TryGetValue(typeof(T), out var mapping)
                ? (CompiledMapping<T>)mapping
                : throw new InvalidOperationException($"No mapping configured for type {typeof(T).Name}. Call Configure<{typeof(T).Name}>() first.");
        }
    }
    
    public bool HasMapping<T>()
    {
        lock (_lock)
        {
            return _compiledMappings.ContainsKey(typeof(T));
        }
    }
    
    internal CompiledMapping<T>? GetCompiledMapping<T>()
    {
        lock (_lock)
        {
            return _compiledMappings.TryGetValue(typeof(T), out var mapping) 
                ? (CompiledMapping<T>)mapping 
                : null;
        }
    }
    
    internal object? GetCompiledMapping(Type type)
    {
        lock (_lock)
        {
            return _compiledMappings.TryGetValue(type, out var mapping) 
                ? mapping 
                : null;
        }
    }
}