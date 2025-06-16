using System.Threading;
using System.Threading.Tasks;

namespace MiniExcelLibs
{
    internal partial interface IExcelWriter
    {
        [Zomp.SyncMethodGenerator.CreateSyncVersion]
        Task<int[]> SaveAsAsync(CancellationToken cancellationToken = default);
        [Zomp.SyncMethodGenerator.CreateSyncVersion]
        Task<int> InsertAsync(bool overwriteSheet = false, CancellationToken cancellationToken = default);
    }
}
