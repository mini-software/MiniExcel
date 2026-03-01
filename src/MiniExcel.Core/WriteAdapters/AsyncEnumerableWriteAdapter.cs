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

    public async IAsyncEnumerable<CellWriteInfo[]> GetRowsAsync(List<MiniExcelColumnInfo> props, [EnumeratorCancellation] CancellationToken cancellationToken)
    {
        if (_empty)
            yield break;

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
            yield return GetRowValues(_enumerator.Current, props);
        }
        while (await _enumerator.MoveNextAsync().ConfigureAwait(false));
    }

    private static CellWriteInfo[] GetRowValues(T currentValue, List<MiniExcelColumnInfo> props)
    {
        var column = 0;
        var result = new List<CellWriteInfo>();
        
        foreach (var prop in props)
        {
            column++;
            
            if (prop is null)
                continue;

            var info = currentValue switch
            {
                IDictionary<string, object> genericDictionary => new CellWriteInfo(genericDictionary[prop.Key.ToString()], column, prop),
                IDictionary dictionary => new CellWriteInfo(dictionary[prop.Key], column, prop),
                _ => new CellWriteInfo(prop.Property.GetValue(currentValue), column, prop)
            };
            result.Add(info);
        }
        
        return result.ToArray();
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