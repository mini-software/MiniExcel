namespace MiniExcelLib.Core.Helpers;

public static class AsyncEnumerableExtensions
{
    public static async Task<List<T>> CreateListAsync<T>(this IAsyncEnumerable<T> enumerable, CancellationToken cancellationToken = default)
    {
        List<T> list = [];
        await foreach (var item in enumerable.WithCancellation(cancellationToken).ConfigureAwait(false))
        {
            list.Add(item);
        }

        return list;
    }

    // needed by the SyncGenerator
    public static List<T> CreateList<T>(this IEnumerable<T> enumerable) => [..enumerable];
}
