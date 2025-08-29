namespace MiniExcelLib.Core.OpenXml;

internal struct MappedRow(int rowIndex)
{
    private const int MaxColumns = 100;
    private object?[]? _cells = null;
    
    public int RowIndex { get; } = rowIndex;

    public void SetCell(int columnIndex, object? value)
    {
        if (value == null)
            return;
            
        // Lazy initialize cells array
        if (_cells == null)
        {
            _cells = new object?[MaxColumns];
        }
        
        if (columnIndex >= 0 && columnIndex < MaxColumns)
        {
            _cells[columnIndex] = value;
        }
    }
    
    public object? GetCell(int columnIndex)
    {
        if (_cells == null || columnIndex < 0 || columnIndex >= MaxColumns)
            return null;
            
        return _cells[columnIndex];
    }
    
    public bool HasData => _cells != null;
}