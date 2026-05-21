using System.ComponentModel;
using System.IO.Compression;

namespace MiniExcelLib.Core.Helpers;

/* todo: instead of using the EditorBrowsableAttribute consider making this class internal and link it for compilation
 in the other projects that require it so as to prevent the consumers' IDEs to be polluted with these extension methods */
[EditorBrowsable(EditorBrowsableState.Advanced)]
public static class Polyfills
{
#if NETSTANDARD2_0
    [EditorBrowsable(EditorBrowsableState.Advanced)]
    public static TValue? GetValueOrDefault<TKey, TValue>(this IDictionary<TKey, TValue> dictionary, TKey key, TValue? defaultValue = default)
    {
        return dictionary.TryGetValue(key, out var value) ? value : defaultValue;
    }

    [EditorBrowsable(EditorBrowsableState.Advanced)]
    public static void Deconstruct<TKey, TValue>(this KeyValuePair<TKey, TValue> kvp, out TKey key, out TValue value)
    {
        key = kvp.Key;
        value = kvp.Value;
    }

    extension(Math)
    {
        [EditorBrowsable(EditorBrowsableState.Advanced)]
        public static TNumber Clamp<TNumber>(TNumber value, TNumber min, TNumber max) where TNumber : IComparable<TNumber>
        {
            if (value.CompareTo(min) < 0) return min;
            if (value.CompareTo(max) > 0) return max;
            return value;
        }
    }

    [EditorBrowsable(EditorBrowsableState.Advanced)]
    public static IEnumerable<TSource> ExceptBy<TSource, TKey>(this IEnumerable<TSource> first, IEnumerable<TKey> second, Func<TSource, TKey> keySelector, IEqualityComparer<TKey>? comparer)
    {
        var set = new HashSet<TKey>(second, comparer);
        foreach (var element in first)
        {
            if (set.Add(keySelector(element)))
            {
                yield return element;
            }
        }
    }
#endif

#if !NET10_0_OR_GREATER
    extension(ZipArchiveEntry entry)
    {
        [EditorBrowsable(EditorBrowsableState.Advanced)]
        public ValueTask<Stream> OpenAsync(CancellationToken cancellationToken = default)
        {
            var stream = entry.Open();
            return new ValueTask<Stream>(stream);
        }
    }

    extension(ZipArchive)
    {
        [EditorBrowsable(EditorBrowsableState.Advanced)]
        public static ValueTask<ZipArchive> CreateAsync(Stream stream, ZipArchiveMode mode, bool leaveOpen, Encoding? entryNameEncoding = null, CancellationToken cancellationToken = default)
        {
            ZipArchive? archive = null;

            try
            {
                archive = new ZipArchive(stream, mode, leaveOpen, entryNameEncoding);
                var result = new ValueTask<ZipArchive>(archive);

                archive = null;
                return result;
            }
            finally
            {
                archive?.Dispose();
            }
        }
    }
#endif
}
