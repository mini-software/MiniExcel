using MiniExcelLib.Core.WriteAdapters;

namespace MiniExcelLib.Core.FluentMapping;

internal interface IMappingCellStream
{
    IMiniExcelWriteAdapter CreateAdapter();
}

internal readonly struct MappingCellStream<T>(IEnumerable<T> items, CompiledMapping<T> mapping, string[] columnLetters) : IMappingCellStream
    where T : class
{
    public MappingCellEnumerator<T> GetEnumerator() 
        => new(items.GetEnumerator(), mapping, columnLetters);
    
    public IMiniExcelWriteAdapter CreateAdapter()
        => new MappingCellStreamAdapter<T>(this, columnLetters);
}

internal struct MappingCellEnumerator<T> : IEnumerator<MappingCell>
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
        
        Current = default;
    }

    public MappingCell Current { get; private set; }
    object IEnumerator.Current => Current;

    public bool MoveNext()
    {
        while (true)
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
                        continue;
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
            if (_currentItem is not null)
            {
                // Cache collection metrics when we start processing an item
                if (_currentColumnIndex == 0 && _currentCollectionRow == 0 && _mapping.Collections.Count > 0)
                {
                    _maxCollectionRows = 0;

                    for (var i = 0; i < _mapping.Collections.Count; i++)
                    {
                        var collection = _mapping.Collections[i];
                        if (collection.Getter(_currentItem) is not { } collectionData)
                            continue;

                        // Convert to a list once - this is the only enumeration
                        var items = collectionData.Cast<object?>().ToList();

                        // Resolve nested mapping info if available
                        NestedMappingInfo? nestedInfo = null;
                        if (_mapping.NestedMappings?.TryGetValue(i, out var precompiledNested) is true)
                        {
                            nestedInfo = precompiledNested;
                        }
                        
                        // Calculate the furthest row this collection (including nested collections) needs
                        var collectionMaxRow = collection.StartCellRow - 1;

                        for (var itemIndex = 0; itemIndex < items.Count; itemIndex++)
                        {
                            if (items[itemIndex] is not { } item)
                                continue;
                            
                            var itemRow = collection.StartCellRow + itemIndex * (1 + collection.RowSpacing);
                            if (itemRow > collectionMaxRow)
                            {
                                collectionMaxRow = itemRow;
                            }

                            if (nestedInfo?.Collections is { Count: > 0 } collections)
                            {
                                foreach (var nested in collections.Values)
                                {
                                    if (nested.Getter(item) is { } nestedData)
                                    {
                                        var nestedIndex = 0;
                                        foreach (var _ in nestedData)
                                        {
                                            var nestedRow = nested.StartRow + 
                                                            itemIndex * (1 + collection.RowSpacing) + 
                                                            nestedIndex * (1 + nested.RowSpacing);

                                            if (nestedRow > collectionMaxRow)
                                            {
                                                collectionMaxRow = nestedRow;
                                            }

                                            nestedIndex++;
                                        }
                                    }
                                }
                            }
                        }

                        var neededRows = collectionMaxRow - _currentRowIndex + 1;
                        if (neededRows > _maxCollectionRows)
                        {
                            _maxCollectionRows = neededRows;
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
                    if (_mapping.TryGetHandler(_currentRowIndex, columnNumber, out var handler))
                    {
                        if (_mapping.TryGetValue(handler, _currentItem, out var value))
                        {
                            cellValue = value ?? _emptyCell;

                            if (value is IFormattable formattable && !string.IsNullOrEmpty(handler.Format))
                            {
                                cellValue = formattable.ToString(handler.Format, null);
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
                    continue;
                }

                // Reset for next item
                _currentCollectionRow = 0;

                // Move to next item
                if (_itemEnumerator.MoveNext())
                {
                    _currentItem = _itemEnumerator.Current;
                    _currentRowIndex++;
                    _currentColumnIndex = 0;
                    continue;
                }
            }

            _isComplete = true;
            return false;
        }
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