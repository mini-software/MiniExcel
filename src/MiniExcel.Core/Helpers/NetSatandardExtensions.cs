namespace MiniExcelLib.Core.Helpers;

public static class NetStandardExtensions
{
#if NETSTANDARD2_0
    public static TValue? GetValueOrDefault<TKey, TValue>(this IReadOnlyDictionary<TKey, TValue> dictionary, TKey key, TValue? defaultValue = default)
    {
        return dictionary.TryGetValue(key, out var value) ? value : defaultValue;
    }
#endif
}
