namespace MiniExcelLib.Core.Abstractions;

public interface IMiniExcelWriteAdapter
{
    bool TryGetKnownCount(out int count);
    List<MiniExcelColumnMapping>? GetColumns();
    IEnumerable<CellWriteInfo[]> GetRows(List<MiniExcelColumnMapping> mappings, CancellationToken cancellationToken = default);
}

public readonly struct CellWriteInfo(object? value, int cellIndex, MiniExcelColumnMapping? mapping)
{
    public object? Value { get; } = value;
    public int CellIndex { get; } = cellIndex;
    public MiniExcelColumnMapping? Mapping { get; } = mapping;
}
