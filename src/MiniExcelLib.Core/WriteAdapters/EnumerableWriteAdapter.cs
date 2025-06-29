using System.Collections;
using MiniExcelLib.Core.Abstractions;
using MiniExcelLib.Core.Helpers;
using MiniExcelLib.Core.Reflection;

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

    public List<MiniExcelColumnInfo>? GetColumns()
    {
        if (CustomPropertyHelper.TryGetTypeColumnInfo(_genericType, _configuration, out var props))
            return props;

        _enumerator = _values.GetEnumerator();
        if (_enumerator.MoveNext())
            return CustomPropertyHelper.GetColumnInfoFromValue(_enumerator.Current, _configuration);
            
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

    public IEnumerable<IEnumerable<CellWriteInfo>> GetRows(List<MiniExcelColumnInfo> props, CancellationToken cancellationToken = default)
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
                yield return GetRowValues(_enumerator.Current, props);
            } 
            while (_enumerator.MoveNext());
        }
        finally
        {
            (_enumerator as IDisposable)?.Dispose();
            _enumerator = null;
        }
    }
        
    public static IEnumerable<CellWriteInfo> GetRowValues(object currentValue, List<MiniExcelColumnInfo> props)
    {
        var column = 1;
        foreach (var prop in props)
        {
            object? cellValue;
            if (prop is null)
            {
                cellValue = null;
            }
            else if (currentValue is IDictionary<string, object> genericDictionary)
            {
                cellValue = genericDictionary[prop.Key.ToString()];
            }
            else if (currentValue is IDictionary dictionary)
            {
                cellValue = dictionary[prop.Key];
            }
            else 
            {
                cellValue = prop.Property.GetValue(currentValue);
            }
                
            yield return new CellWriteInfo(cellValue, column, prop);
            column++;
        }
    }
}