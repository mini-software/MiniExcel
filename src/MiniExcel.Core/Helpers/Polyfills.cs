namespace MiniExcelLib.Core.Helpers;

public static class Polyfills
{
#if NETSTANDARD2_0
    public static TValue? GetValueOrDefault<TKey, TValue>(this IDictionary<TKey, TValue> dictionary, TKey key, TValue? defaultValue = default)
    {
        return dictionary.TryGetValue(key, out var value) ? value : defaultValue;
    }

    extension(Math)
    {
        public static TNumber Clamp<TNumber>(TNumber value, TNumber min, TNumber max) where TNumber : unmanaged, IComparable<TNumber>
        {
            if (value.CompareTo(min) < 0) return min;
            if (value.CompareTo(max) > 0) return max;
            return value;
        } 
    }
#endif
}
