using MiniExcelLibs.Utils;

#if NETSTANDARD2_0_OR_GREATER || NET
namespace MiniExcelLibs.WriteAdapter
{
    internal interface IAsyncMiniExcelWriteAdapter 
    {
        Task<List<ExcelColumnInfo>> GetColumnsAsync();

        IAsyncEnumerable<IAsyncEnumerable<CellWriteInfo>> GetRowsAsync(List<ExcelColumnInfo> props, CancellationToken cancellationToken);
    }
}
#endif
