using MiniExcelLibs.Utils;
using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;

namespace MiniExcelLibs.WriteAdapter
{
    internal interface IAsyncMiniExcelWriteAdapter 
    {
        Task<List<ExcelColumnInfo>> GetColumnsAsync();

        IAsyncEnumerable<IAsyncEnumerable<CellWriteInfo>> GetRowsAsync(List<ExcelColumnInfo> props, CancellationToken cancellationToken);
    }
}
