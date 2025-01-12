using MiniExcelLibs.Utils;
using System.Collections.Generic;
using System.Threading.Tasks;

#if NETSTANDARD2_0_OR_GREATER || NET
namespace MiniExcelLibs.WriteAdapter
{
    internal interface IAsyncMiniExcelWriteAdapter 
    {
        Task<List<ExcelColumnInfo>> GetColumnsAsync();

        IAsyncEnumerable<IAsyncEnumerable<CellWriteInfo>> GetRowsAsync(List<ExcelColumnInfo> props);
    }
}
#endif
