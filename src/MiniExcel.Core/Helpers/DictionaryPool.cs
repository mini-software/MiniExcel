using System.Collections.Concurrent;

namespace MiniExcelLib.Core.Helpers;

/// <summary>
/// Simple object pool for dictionaries to reduce allocations
/// </summary>
internal static class DictionaryPool
{
    // Simple ConcurrentBag-based pool for .NET Standard
    private static readonly ConcurrentBag<Dictionary<string, object>> Pool = new();
    private static int _poolSize;
    private const int MaxPoolSize = 100;
    
    public static Dictionary<string, object> Rent()
    {
        if (Pool.TryTake(out var dictionary))
        {
            Interlocked.Decrement(ref _poolSize);
            return dictionary;
        }
        
        return new Dictionary<string, object>(16); // Pre-size for typical row
    }
    
    public static void Return(Dictionary<string, object> dictionary)
    {
        // Don't pool huge dictionaries
        if (dictionary.Count > 1000)
            return;
            
        dictionary.Clear();
        
        // Limit pool size to prevent unbounded growth
        if (_poolSize < MaxPoolSize)
        {
            Pool.Add(dictionary);
            Interlocked.Increment(ref _poolSize);
        }
    }
}