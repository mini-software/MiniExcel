namespace MiniExcelLib.Abstractions;

public interface IMiniExcelWriteAdapter
{
    bool TryGetKnownCount(out int count);
    List<MiniExcelColumnInfo>? GetColumns();
    IEnumerable<IEnumerable<CellWriteInfo>> GetRows(List<MiniExcelColumnInfo> props, CancellationToken cancellationToken = default);
}

public readonly struct CellWriteInfo(object? value, int cellIndex, MiniExcelColumnInfo prop)
{
    public object? Value { get; } = value;
    public int CellIndex { get; } = cellIndex;
    public MiniExcelColumnInfo Prop { get; } = prop;
}