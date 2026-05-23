using System.ComponentModel;
using System.IO.Compression;
using System.Runtime.InteropServices;
using System.Xml.Linq;

namespace MiniExcelLib.Core.Helpers;

/* todo: instead of using the EditorBrowsableAttribute consider making this class internal and link it for compilation
 in the other projects that require it so as to prevent the consumers' IDEs to be polluted with these extension methods */
[EditorBrowsable(EditorBrowsableState.Advanced)]
public static class Polyfills
{
#if NETSTANDARD2_0
    extension<TKey, TValue>(IDictionary<TKey, TValue> dictionary)
    {
        [EditorBrowsable(EditorBrowsableState.Advanced)]
        public TValue? GetValueOrDefault(TKey key, TValue? defaultValue = default)
        {
            return dictionary.TryGetValue(key, out var value) ? value : defaultValue;
        }

        [EditorBrowsable(EditorBrowsableState.Advanced)]
        public bool TryAdd(TKey key, TValue value)
        {
            if (dictionary.ContainsKey(key))
                return false;
 
            dictionary.Add(key, value);
            return true;
        }
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

    extension(Stream? stream)
    {
        [EditorBrowsable(EditorBrowsableState.Advanced)]
        public ValueTask DisposeAsync()
        {
            if (stream is IAsyncDisposable asyncDisposable)
                return asyncDisposable.DisposeAsync();
    
            stream?.Dispose();
            return default;
        }
    
        [EditorBrowsable(EditorBrowsableState.Advanced)]
        public StreamConfiguredAsyncDisposable ConfigureAwait(bool continueOnCapturedContext) 
            => new(stream, continueOnCapturedContext);
    }

    /// <summary>
    /// This is a copy of the runtime's <see cref="ConfiguredAsyncDisposable" />, which we cannot instantiate directly for our needs
    /// due to the constructor that initializes the object to eventually dispose being internal. 
    /// </summary>
    [StructLayout(LayoutKind.Auto)]
    [EditorBrowsable(EditorBrowsableState.Advanced)]
    public readonly struct StreamConfiguredAsyncDisposable : IDisposable
    {
        private readonly Stream? _source;
        private readonly bool _continueOnCapturedContext;
    
        internal StreamConfiguredAsyncDisposable(Stream? source, bool continueOnCapturedContext)
        {
            _source = source;
            _continueOnCapturedContext = continueOnCapturedContext;
        }
    
        public ConfiguredValueTaskAwaitable DisposeAsync() 
            => _source.DisposeAsync().ConfigureAwait(_continueOnCapturedContext);

        public void Dispose() 
            => _source?.Dispose();
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

    extension(XDocument doc)
    {
        [EditorBrowsable(EditorBrowsableState.Advanced)]
        public static ValueTask<XDocument> LoadAsync(Stream stream, LoadOptions loadOptions, CancellationToken cancellationToken = default)
        {
            return new ValueTask<XDocument>(XDocument.Load(stream, loadOptions));
        }

        [EditorBrowsable(EditorBrowsableState.Advanced)]
        public ValueTask SaveAsync(Stream stream, SaveOptions saveOptions, CancellationToken cancellationToken = default)
        {
            doc.Save(stream, saveOptions);
            return default;
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
