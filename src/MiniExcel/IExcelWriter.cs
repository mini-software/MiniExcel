using System.Threading;
using System.Threading.Tasks;

namespace MiniExcelLibs
{
    internal interface IExcelWriter
    {
        int[] SaveAs();
        Task<int[]> SaveAsAsync(CancellationToken cancellationToken = default);
        int Insert(bool overwriteSheet = false);
        Task<int> InsertAsync(bool overwriteSheet = false, CancellationToken cancellationToken = default);
    }
}
