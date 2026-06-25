using MiniExcelLibs.Utils;

#if !NET45
namespace MiniExcelLibs.WriteAdapter
{
    internal interface IAsyncMiniExcelWriteAdapter 
    {
        Task<List<ExcelColumnInfo>> GetColumnsAsync();

        IAsyncEnumerable<IAsyncEnumerable<CellWriteInfo>> GetRowsAsync(List<ExcelColumnInfo> props, CancellationToken cancellationToken);
    }
}
#endif
