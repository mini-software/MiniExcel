namespace MiniExcelLib.Core.Abstractions;

public interface IMiniExcelWriteAdapterAsync 
{
    Task<List<MiniExcelColumnInfo>?> GetColumnsAsync();
    IAsyncEnumerable<CellWriteInfo[]> GetRowsAsync(List<MiniExcelColumnInfo> props, CancellationToken cancellationToken);
}