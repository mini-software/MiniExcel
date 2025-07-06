using MiniExcelLib.Reflection;

namespace MiniExcelLib.Abstractions;

public interface IMiniExcelWriteAdapterAsync 
{
    Task<List<MiniExcelColumnInfo>?> GetColumnsAsync();
    IAsyncEnumerable<IAsyncEnumerable<CellWriteInfo>> GetRowsAsync(List<MiniExcelColumnInfo> props, CancellationToken cancellationToken);
}