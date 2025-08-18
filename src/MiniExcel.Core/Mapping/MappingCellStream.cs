using MiniExcelLib.Core.WriteAdapters;

namespace MiniExcelLib.Core.Mapping;

internal readonly struct MappingCellStream<T>(IEnumerable<T> items, CompiledMapping<T> mapping, string[] columnLetters) : IMappingCellStream
    where T : class
{
    public MappingCellEnumerator<T> GetEnumerator() 
        => new(items.GetEnumerator(), mapping, columnLetters);
    
    public IMiniExcelWriteAdapter CreateAdapter()
        => new MappingCellStreamAdapter<T>(this, columnLetters);
}

internal struct MappingCellEnumerator<T>
    where T : class
{
    private readonly IEnumerator<T> _itemEnumerator;
    private readonly CompiledMapping<T> _mapping;
    private readonly string[] _columnLetters;
    private readonly OptimizedMappingBoundaries _boundaries;
    private readonly int _columnCount;
    
    private T? _currentItem;
    private int _currentRowIndex;
    private int _currentColumnIndex;
    private bool _hasStartedData;
    private bool _isComplete;
    private readonly object _emptyCell;
    private int _maxCollectionRows;
    private int _currentCollectionRow;
    private object?[][]? _currentCollectionArrays;

    public MappingCellEnumerator(IEnumerator<T> itemEnumerator, CompiledMapping<T> mapping, string[] columnLetters)
    {
        _itemEnumerator = itemEnumerator;
        _mapping = mapping;
        _columnLetters = columnLetters;
        _boundaries = mapping.OptimizedBoundaries!;
        _columnCount = _boundaries.MaxColumn - _boundaries.MinColumn + 1;
        
        _currentItem = default;
        _currentRowIndex = 0;
        _currentColumnIndex = 0;
        _hasStartedData = false;
        _isComplete = false;
        _emptyCell = string.Empty;
        _maxCollectionRows = 0;
        _currentCollectionRow = 0;
        _currentCollectionArrays = null;
        
        Current = default;
    }

    public MappingCell Current { get; private set; }

    public bool MoveNext()
    {
        if (_isComplete)
            return false;

        // Handle rows before data starts (if MinRow > 1)
        if (!_hasStartedData)
        {
            if (_currentRowIndex == 0)
            {
                _currentRowIndex = 1;
                _currentColumnIndex = 0;
            }

            // Emit empty cells for rows before MinRow
            if (_currentRowIndex < _boundaries.MinRow)
            {
                if (_currentColumnIndex < _columnCount)
                {
                    var columnLetter = _columnLetters[_currentColumnIndex];
                    Current = new MappingCell(columnLetter, _currentRowIndex, _emptyCell);
                    _currentColumnIndex++;
                    return true;
                }

                // Move to next empty row
                _currentRowIndex++;
                _currentColumnIndex = 0;
                    
                if (_currentRowIndex < _boundaries.MinRow)
                {
                    return MoveNext(); // Continue with next empty row
                }
            }

            // Start processing actual data
            _hasStartedData = true;
            if (!_itemEnumerator.MoveNext())
            {
                _isComplete = true;
                return false;
            }
            _currentItem = _itemEnumerator.Current;
            _currentColumnIndex = 0;
        }

        // Process current item's cells
        if (_currentItem != null)
        {
            // Cache collections as arrays when we start processing an item
            if (_currentColumnIndex == 0 && _currentCollectionRow == 0 && _mapping.Collections.Count > 0)
            {
                _maxCollectionRows = 0;
                _currentCollectionArrays = new object?[_mapping.Collections.Count][];
                
                for (var i = 0; i < _mapping.Collections.Count; i++)
                {
                    var coll = _mapping.Collections[i];
                    var collectionData = coll.Getter(_currentItem);
                    if (collectionData != null)
                    {
                        // Convert to array once - this is the only enumeration
                        var items = new List<object?>();
                        foreach (var item in collectionData)
                        {
                            items.Add(item);
                        }
                        _currentCollectionArrays[i] = items.ToArray();
                        
                        // For vertical collections, we need rows from StartCellRow
                        var neededRows = coll.StartCellRow + items.Count - _currentRowIndex;
                        if (neededRows > _maxCollectionRows)
                            _maxCollectionRows = neededRows;
                    }
                    else
                    {
                        _currentCollectionArrays[i] = [];
                    }
                }
            }
            
            // Emit cells for current row
            if (_currentColumnIndex < _columnCount)
            {
                var columnLetter = _columnLetters[_currentColumnIndex];
                var columnNumber = _boundaries.MinColumn + _currentColumnIndex;
                
                object? cellValue = _emptyCell;

                // Use the optimized grid for fast lookup
                if (_mapping.TryGetCellValue(_currentRowIndex, columnNumber, _currentItem, out var value))
                {
                    cellValue = value ?? _emptyCell;
                    
                    // Apply formatting if needed
                    if (value is IFormattable formattable && 
                        _mapping.TryGetHandler(_currentRowIndex, columnNumber, out var handler) &&
                        !string.IsNullOrEmpty(handler.Format))
                    {
                        cellValue = formattable.ToString(handler.Format, null);
                    }
                }
                else if (_currentCollectionArrays != null)
                {
                    // Fallback for collections that might not be in the grid yet
                    // This handles dynamic collection expansion
                    for (var collIndex = 0; collIndex < _mapping.Collections.Count; collIndex++)
                    {
                        var coll = _mapping.Collections[collIndex];
                        if (coll.StartCellColumn == columnNumber)
                        {
                            // This is a collection column - check if this row has a collection item
                            var collectionRowOffset = _currentRowIndex - coll.StartCellRow;
                            if (collectionRowOffset >= 0)
                            {
                                var collectionArray = _currentCollectionArrays[collIndex];
                                if (collectionRowOffset < collectionArray.Length)
                                {
                                    cellValue = collectionArray[collectionRowOffset] ?? _emptyCell;
                                }
                            }
                            break;
                        }
                    }
                }

                Current = new MappingCell(columnLetter, _currentRowIndex, cellValue);
                _currentColumnIndex++;
                return true;
            }

            // Check if we need to emit more rows for collections
            _currentCollectionRow++;
            if (_currentCollectionRow < _maxCollectionRows)
            {
                _currentRowIndex++;
                _currentColumnIndex = 0;
                return MoveNext();
            }

            // Reset for next item
            _currentCollectionRow = 0;
            
            // Move to next item
            if (_itemEnumerator.MoveNext())
            {
                _currentItem = _itemEnumerator.Current;
                _currentRowIndex++;
                _currentColumnIndex = 0;
                return MoveNext();
            }
        }

        _isComplete = true;
        return false;
    }

    public void Reset()
    {
        throw new NotSupportedException();
    }

    public void Dispose()
    {
        _itemEnumerator?.Dispose();
    }
}

internal readonly struct MappingCell(string columnLetter, int rowIndex, object? value)
{
    public readonly string ColumnLetter = columnLetter;
    public readonly int RowIndex = rowIndex;
    public readonly object? Value = value;
}