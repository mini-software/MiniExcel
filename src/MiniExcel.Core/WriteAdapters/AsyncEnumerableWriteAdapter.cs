namespace MiniExcelLib.Core.WriteAdapters;

internal sealed class AsyncEnumerableWriteAdapter<T>(IAsyncEnumerable<T> values, MiniExcelBaseConfiguration configuration) : IMiniExcelWriteAdapterAsync, IAsyncDisposable
{
    private readonly IAsyncEnumerable<T> _values = values;
    private readonly MiniExcelBaseConfiguration _configuration = configuration;
    
    private IAsyncEnumerator<T>? _enumerator;
    private bool _empty;
    private bool _disposed = false;

    
    public async Task<List<MiniExcelColumnInfo>?> GetColumnsAsync()
    {
        if (CustomPropertyHelper.TryGetTypeColumnInfo(typeof(T), _configuration, out var props))
        {
            return props;
        }

        _enumerator = _values.GetAsyncEnumerator();
        if (!await _enumerator.MoveNextAsync().ConfigureAwait(false))
        {
            _empty = true;
            return null;
        }

        return CustomPropertyHelper.GetColumnInfoFromValue(_enumerator.Current, _configuration);
    }

    public async IAsyncEnumerable<IAsyncEnumerable<CellWriteInfo>> GetRowsAsync(List<MiniExcelColumnInfo> props, [EnumeratorCancellation] CancellationToken cancellationToken)
    {
        if (_empty)
        {
            yield break;
        }

        if (_enumerator is null)
        {
            _enumerator = _values.GetAsyncEnumerator(cancellationToken);
            if (!await _enumerator.MoveNextAsync().ConfigureAwait(false))
            {
                yield break;
            }
        }

        do
        {
            cancellationToken.ThrowIfCancellationRequested();
            yield return GetRowValuesAsync(_enumerator.Current, props);
        }
        while (await _enumerator.MoveNextAsync().ConfigureAwait(false));
    }

#pragma warning disable CS1998 // Async method lacks 'await' operators and will run synchronously
    private static async IAsyncEnumerable<CellWriteInfo> GetRowValuesAsync(T currentValue, List<MiniExcelColumnInfo> props)
#pragma warning restore CS1998
    {
        var column = 0;
        foreach (var prop in props)
        {
            column++;
            
            if (prop is null)
                continue;

            yield return currentValue switch
            {
                IDictionary<string, object> genericDictionary => new CellWriteInfo(genericDictionary[prop.Key.ToString()], column, prop),
                IDictionary dictionary => new CellWriteInfo(dictionary[prop.Key], column, prop),
                _ => new CellWriteInfo(prop.Property.GetValue(currentValue), column, prop)
            };
        }
    }

    public async ValueTask DisposeAsync()
    {
        if (!_disposed)
        {
            if (_enumerator is not null)
            {
                await _enumerator.DisposeAsync().ConfigureAwait(false);
            }
            _disposed = true;
        }
    }
}