namespace MiniExcelLib.Core.WriteAdapters;

internal class DataTableWriteAdapter(DataTable dataTable, MiniExcelBaseConfiguration configuration) : IMiniExcelWriteAdapter
{
    private readonly DataTable _dataTable = dataTable;
    private readonly MiniExcelBaseConfiguration _configuration = configuration;

    public bool TryGetKnownCount(out int count)
    {
        count = _dataTable.Rows.Count;
        return true;
    }

    public List<MiniExcelColumnMapping> GetColumns()
    {
        var mappings = new List<MiniExcelColumnMapping>();
        for (var i = 0; i < _dataTable.Columns.Count; i++)
        {
            var columnName = _dataTable.Columns[i].Caption ?? _dataTable.Columns[i].ColumnName;
            var map = ColumnMappingsProvider.GetColumnMappingFromDynamicConfiguration(columnName, _configuration);
            mappings.Add(map);
        }
        return mappings;
    }

    public IEnumerable<CellWriteInfo[]> GetRows(List<MiniExcelColumnMapping> mappings, CancellationToken cancellationToken = default)
    {
        for (int row = 0; row < _dataTable.Rows.Count; row++)
        {
            cancellationToken.ThrowIfCancellationRequested();
            yield return GetRowValues(row, mappings);
        }
    }

    private CellWriteInfo[] GetRowValues(int row, List<MiniExcelColumnMapping> mappings)
    {
        var result = new List<CellWriteInfo>(mappings.Count);
        for (int i = 0, column = 1; i < _dataTable.Columns.Count; i++, column++)
        {
            result.Add(new CellWriteInfo(_dataTable.Rows[row][i], column, mappings[i]));
        }
        return result.ToArray();
    }
}
