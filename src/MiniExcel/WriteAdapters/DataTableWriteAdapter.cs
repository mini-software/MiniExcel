namespace MiniExcelLib.WriteAdapters;

internal class DataTableWriteAdapter(DataTable dataTable, MiniExcelBaseConfiguration configuration) : IMiniExcelWriteAdapter
{
    private readonly DataTable _dataTable = dataTable;
    private readonly MiniExcelBaseConfiguration _configuration = configuration;

    public bool TryGetKnownCount(out int count)
    {
        count = _dataTable.Rows.Count;
        return true;
    }

    public List<MiniExcelColumnInfo> GetColumns()
    {
        var props = new List<MiniExcelColumnInfo>();
        for (var i = 0; i < _dataTable.Columns.Count; i++)
        {
            var columnName = _dataTable.Columns[i].Caption ?? _dataTable.Columns[i].ColumnName;
            var prop = CustomPropertyHelper.GetColumnInfosFromDynamicConfiguration(columnName, _configuration);
            props.Add(prop);
        }
        return props;
    }

    public IEnumerable<IEnumerable<CellWriteInfo>> GetRows(List<MiniExcelColumnInfo> props, CancellationToken cancellationToken = default)
    {
        for (int row = 0; row < _dataTable.Rows.Count; row++)
        {
            cancellationToken.ThrowIfCancellationRequested();
            yield return GetRowValues(row, props);
        }
    }

    private IEnumerable<CellWriteInfo> GetRowValues(int row, List<MiniExcelColumnInfo> props)
    {
        for (int i = 0, column = 1; i < _dataTable.Columns.Count; i++, column++)
        {
            yield return new CellWriteInfo(_dataTable.Rows[row][i], column, props[i]);
        }
    }
}