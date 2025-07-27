using MiniExcelLib.Core.Abstractions;
using MiniExcelLib.Core.Reflection;

namespace MiniExcelLib.Core.WriteAdapters;

internal class DataReaderWriteAdapter(IDataReader reader, MiniExcelBaseConfiguration configuration) : IMiniExcelWriteAdapter
{
    private readonly IDataReader _reader = reader;
    private readonly MiniExcelBaseConfiguration _configuration = configuration;

    public bool TryGetKnownCount(out int count)
    {
        count = 0;
        return false;
    }

    public List<MiniExcelColumnInfo> GetColumns()
    {
        var props = new List<MiniExcelColumnInfo>();
        for (var i = 0; i < _reader.FieldCount; i++)
        {
            var columnName = _reader.GetName(i);
            if (!_configuration.DynamicColumnFirst || 
                _configuration.DynamicColumns.Any(d => string.Equals(d.Key, columnName, StringComparison.OrdinalIgnoreCase)))
            {
                var prop = CustomPropertyHelper.GetColumnInfosFromDynamicConfiguration(columnName, _configuration);
                props.Add(prop);
            }
        }
        return props;
    }

    public IEnumerable<IEnumerable<CellWriteInfo>> GetRows(List<MiniExcelColumnInfo> props, CancellationToken cancellationToken = default)
    {
        while (_reader.Read())
        {
            cancellationToken.ThrowIfCancellationRequested();
            yield return GetRowValues(props);
        }
    }

    private IEnumerable<CellWriteInfo> GetRowValues(List<MiniExcelColumnInfo> props)
    {
        var column = 1;
        for (int i = 0; i < _reader.FieldCount; i++)
        {
            var prop = props[i];
            if (prop is { ExcelIgnore: false })
            {
                var columnIndex = _configuration.DynamicColumnFirst 
                    ? _reader.GetOrdinal(prop.Key.ToString())
                    : i;
                
                yield return new CellWriteInfo(_reader.GetValue(columnIndex), column, prop);
                column++;
            }
        }
    }
}