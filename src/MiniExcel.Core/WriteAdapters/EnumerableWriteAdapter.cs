namespace MiniExcelLib.Core.WriteAdapters;

internal class EnumerableWriteAdapter(IEnumerable values, MiniExcelBaseConfiguration configuration) : IMiniExcelWriteAdapter
{
    private readonly IEnumerable _values = values;
    private readonly MiniExcelBaseConfiguration _configuration = configuration;
    private readonly Type? _genericType = TypeHelper.GetGenericIEnumerables(values).FirstOrDefault();

    private IEnumerator? _enumerator;
    private bool _empty;

    public bool TryGetKnownCount(out int count)
    {
        count = 0;
        if (_values is ICollection collection)
        {
            count = collection.Count;
            return true;
        }

        return false;
    }

    public List<MiniExcelColumnMapping>? GetColumns()
    {
        if (ColumnMappingsProvider.TryGetColumnMappings(_genericType, _configuration, out var mappings))
            return mappings;

        _enumerator = _values.GetEnumerator();
        if (_enumerator.MoveNext())
            return ColumnMappingsProvider.GetColumnMappingFromValue(_enumerator.Current, _configuration);
            
        try
        {
            _empty = true;
            return null;
        }
        finally
        {
            (_enumerator as IDisposable)?.Dispose();
            _enumerator = null;
        }
    }

    public IEnumerable<CellWriteInfo[]> GetRows(List<MiniExcelColumnMapping> mappings, CancellationToken cancellationToken = default)
    {
        if (_empty)
            yield break;

        try
        {
            if (_enumerator is null)
            {
                _enumerator = _values.GetEnumerator();
                if (!_enumerator.MoveNext())
                    yield break;
            }

            do
            {
                cancellationToken.ThrowIfCancellationRequested();
                yield return GetRowValues(_enumerator.Current, mappings);
            } 
            while (_enumerator.MoveNext());
        }
        finally
        {
            (_enumerator as IDisposable)?.Dispose();
            _enumerator = null;
        }
    }

    private static CellWriteInfo[] GetRowValues(object currentValue, List<MiniExcelColumnMapping?> mappings)
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
}
