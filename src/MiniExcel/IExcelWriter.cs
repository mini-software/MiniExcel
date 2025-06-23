using System.Threading;
using System.Threading.Tasks;
using Zomp.SyncMethodGenerator;

namespace MiniExcelLibs;

internal partial interface IExcelWriter
{
    [CreateSyncVersion]
    Task<int[]> SaveAsAsync(CancellationToken cancellationToken = default);

    [CreateSyncVersion]
    Task<int> InsertAsync(bool overwriteSheet = false, CancellationToken cancellationToken = default);
}