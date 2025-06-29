using MiniExcelLib.Core.Reflection;

namespace MiniExcelLib.Core.Abstractions;

public interface IMiniExcelWriteAdapterAsync 
{
    Task<List<MiniExcelColumnInfo>?> GetColumnsAsync();
    IAsyncEnumerable<IAsyncEnumerable<CellWriteInfo>> GetRowsAsync(List<MiniExcelColumnInfo> props, CancellationToken cancellationToken);
}