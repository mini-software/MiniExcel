using MiniExcelLib.Core.FluentMapping;

namespace MiniExcelLib.Core.WriteAdapters;

internal class MappingCellStreamAdapter<T>(MappingCellStream<T> cellStream, string[] columnLetters)
    : IMiniExcelWriteAdapter
    where T : class
{
    private readonly MappingCellStream<T> _cellStream = cellStream;
    private readonly string[] _columnLetters = columnLetters;

    public bool TryGetKnownCount(out int count)
    {
        // We don't know the exact row count without iterating
        count = 0;
        return false;
    }

    public List<MiniExcelColumnInfo> GetColumns()
    {
        var props = new List<MiniExcelColumnInfo>();
        
        for (int i = 0; i < _columnLetters.Length; i++)
        {
            props.Add(new MiniExcelColumnInfo
            {
                Key = _columnLetters[i],
                ExcelColumnName = _columnLetters[i],
                ExcelColumnIndex = i
            });
        }
        
        return props;
    }

    public IEnumerable<IEnumerable<CellWriteInfo>> GetRows(List<MiniExcelColumnInfo> props, CancellationToken cancellationToken = default)
    {
        var currentRow = new Dictionary<string, object?>();
        var currentRowIndex = 0;
        
        foreach (var cell in _cellStream)
        {
            cancellationToken.ThrowIfCancellationRequested();
            
            // Check if we've moved to a new row
            if (cell.RowIndex != currentRowIndex)
            {
                // Yield the completed row if we have one
                if (currentRowIndex > 0 && currentRow.Count > 0)
                {
                    yield return ConvertRowToCellWriteInfos(currentRow, props);
                }
                
                // Start new row
                currentRow.Clear();
                currentRowIndex = cell.RowIndex;
            }
            
            // Add cell to current row
            currentRow[cell.ColumnLetter] = cell.Value;
        }
        
        // Yield the final row
        if (currentRow.Count > 0)
        {
            yield return ConvertRowToCellWriteInfos(currentRow, props);
        }
    }

    private static IEnumerable<CellWriteInfo> ConvertRowToCellWriteInfos(Dictionary<string, object?> row, List<MiniExcelColumnInfo> props)
    {
        var columnIndex = 1;
        foreach (var prop in props)
        {
            object? cellValue = null;
            if (row.TryGetValue(prop.Key.ToString(), out var value))
            {
                cellValue = value;
            }
            
            yield return new CellWriteInfo(cellValue, columnIndex, prop);
            columnIndex++;
        }
    }
}