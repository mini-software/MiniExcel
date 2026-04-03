using MiniExcelLib.Core.Abstractions;
using MiniExcelLib.Core.Reflection;

namespace MiniExcelLib.OpenXml.FluentMapping;

internal class MappingCellStreamAdapter<T>(MappingCellStream<T> cellStream, string[] columnLetters)
    : IMiniExcelWriteAdapter where T : class
{
    private readonly MappingCellStream<T> _cellStream = cellStream;
    private readonly string[] _columnLetters = columnLetters;

    public bool TryGetKnownCount(out int count)
    {
        // We don't know the exact row count without iterating
        count = 0;
        return false;
    }

    public List<MiniExcelColumnMapping> GetColumns()
    {
        var mappings = new List<MiniExcelColumnMapping>();
        for (int i = 0; i < _columnLetters.Length; i++)
        {
            mappings.Add(new MiniExcelColumnMapping
            {
                Key = _columnLetters[i],
                ExcelColumnName = _columnLetters[i],
                ExcelColumnIndex = i
            });
        }
        
        return mappings;
    }

    public IEnumerable<CellWriteInfo[]> GetRows(List<MiniExcelColumnMapping> mappings, CancellationToken cancellationToken = default)
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
                    yield return ConvertRowToCellWriteInfos(currentRow, mappings);
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
            yield return ConvertRowToCellWriteInfos(currentRow, mappings);
        }
    }

    private static CellWriteInfo[] ConvertRowToCellWriteInfos(Dictionary<string, object?> row, List<MiniExcelColumnMapping> mappings)
    {
        var columnIndex = 0;
        var result = new List<CellWriteInfo>(mappings.Count);

        foreach (var map in mappings)
        {
            object? cellValue = null;
            if (row.TryGetValue(prop.Key.ToString(), out var value))
            {
                cellValue = value;
            }
            
            columnIndex++;
            result.Add(new CellWriteInfo(cellValue, columnIndex, map));
        }

        return result.ToArray();
    }
}
