namespace MiniExcelLib.Core.Abstractions;

public interface IMiniExcelWriteAdapter
{
    bool TryGetKnownCount(out int count);
    List<MiniExcelColumnMapping>? GetColumns();
    IEnumerable<CellWriteInfo[]> GetRows(List<MiniExcelColumnMapping> mappings, CancellationToken cancellationToken = default);
}

public readonly struct CellWriteInfo(object? value, int index, MiniExcelColumnMapping? mapping)
{
    public object? Value { get; } = value;
    public int Index { get; } = index;
    public MiniExcelColumnMapping? Mapping { get; } = mapping;
}
