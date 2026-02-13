namespace MiniExcelLib.OpenXml;

public struct MappedRow(int rowIndex)
{
    private const int MaxColumns = 100;
    private object?[]? _cells = null;
    
    public int RowIndex { get; } = rowIndex;

    public void SetCell(int columnIndex, object? value)
    {
        if (value is null)
            return;
            
        // Lazy initialize cells array
        _cells ??= new object?[MaxColumns];
        
        if (columnIndex is >= 0 and < MaxColumns)
        {
            _cells[columnIndex] = value;
        }
    }
    
    public object? GetCell(int columnIndex)
    {
        if (_cells is null || (columnIndex is < 0 or >= MaxColumns))
            return null;
            
        return _cells[columnIndex];
    }
    
    public bool HasData => _cells is not null;
}