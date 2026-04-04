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

    public List<MiniExcelColumnMapping> GetColumns()
    {
        var mappings = new List<MiniExcelColumnMapping>();
        for (var i = 0; i < _reader.FieldCount; i++)
        {
            var columnName = _reader.GetName(i);
            if (!_configuration.DynamicColumnFirst || 
                _configuration.DynamicColumns?.Any(d => string.Equals(d.Key, columnName, StringComparison.OrdinalIgnoreCase)) is true)
            {
                var map = ColumnMappingsProvider.GetColumnMappingFromDynamicConfiguration(columnName, _configuration);
                mappings.Add(map);
            }
        }
        return mappings;
    }

    public IEnumerable<CellWriteInfo[]> GetRows(List<MiniExcelColumnMapping> mappings, CancellationToken cancellationToken = default)
    {
        while (_reader.Read())
        {
            cancellationToken.ThrowIfCancellationRequested();
            yield return GetRowValues(mappings);
        }
    }

    private CellWriteInfo[] GetRowValues(List<MiniExcelColumnMapping> mappings)
    {
        var column = 1;
        var result = new List<CellWriteInfo>(mappings.Count);
    
        for (int i = 0; i < _reader.FieldCount; i++)
        {
            var map = mappings[i];
            if (map is not { ExcelIgnoreColumn: true })
            {
                var columnIndex = _configuration.DynamicColumnFirst 
                    ? _reader.GetOrdinal(map.Key.ToString())
                    : i;

                result.Add(new CellWriteInfo(_reader.GetValue(columnIndex), column, map));
                column++;
            }
        }

        return result.ToArray();
    }
}
