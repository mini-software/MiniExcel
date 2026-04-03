namespace MiniExcelLib.Core.WriteAdapters;

internal sealed class AsyncEnumerableWriteAdapter<T>(IAsyncEnumerable<T> values, MiniExcelBaseConfiguration configuration) : IMiniExcelWriteAdapterAsync, IAsyncDisposable
{
    private readonly IAsyncEnumerable<T> _values = values;
    private readonly MiniExcelBaseConfiguration _configuration = configuration;
    
    private IAsyncEnumerator<T>? _enumerator;
    private bool _empty;
    private bool _disposed = false;

    
    public async Task<List<MiniExcelColumnMapping>?> GetColumnsAsync()
    {
        if (ColumnMappingsProvider.TryGetColumnMappings(typeof(T), _configuration, out var mappings))
        {
            return mappings;
        }

        _enumerator = _values.GetAsyncEnumerator();
        if (!await _enumerator.MoveNextAsync().ConfigureAwait(false))
        {
            _empty = true;
            return null;
        }

        return ColumnMappingsProvider.GetColumnMappingFromValue(_enumerator.Current, _configuration);
    }

    public async IAsyncEnumerable<CellWriteInfo[]> GetRowsAsync(List<MiniExcelColumnMapping> mappings, [EnumeratorCancellation] CancellationToken cancellationToken)
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
            yield return GetRowValues(_enumerator.Current, mappings);
        }
        while (await _enumerator.MoveNextAsync().ConfigureAwait(false));
    }

    private static CellWriteInfo[] GetRowValues(T currentValue, List<MiniExcelColumnMapping?> mappings)
    {
        var column = 0;
        var result = new List<CellWriteInfo>(mappings.Count);
        
        foreach (var map in mappings)
        {
            column++;
            var cellValue = currentValue switch
            {
                _ when map is null => null,
                IDictionary<string, object> genericDictionary => genericDictionary[map.Key.ToString()],
                IDictionary dictionary => dictionary[map.Key],
                _ => map.MemberAccessor.GetValue(currentValue)
            };
            result.Add(new CellWriteInfo(cellValue, column, map));
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
