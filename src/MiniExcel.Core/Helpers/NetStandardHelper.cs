namespace MiniExcelLib.Core.Helpers;

#if NETSTANDARD2_0

/// <summary>
/// Provides .NET Standard 2.0 polyfills for utility methods found in later framework versions.
/// This enables a unified API surface across the codebase without the need for conditional compilation directives.
/// </summary>
public static class NetStandardExtensions
{
    public static TValue? GetValueOrDefault<TKey, TValue>(this IReadOnlyDictionary<TKey, TValue> dictionary, TKey key, TValue? defaultValue = default)
    {
        return dictionary.TryGetValue(key, out var value) ? value : defaultValue;
    }
}
#endif
