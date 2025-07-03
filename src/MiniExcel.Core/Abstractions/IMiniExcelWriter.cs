using Zomp.SyncMethodGenerator;

namespace MiniExcelLib.Core.Abstractions;

public partial interface IMiniExcelWriter
{
    [CreateSyncVersion]
    Task<int[]> SaveAsAsync(CancellationToken cancellationToken = default);

    [CreateSyncVersion]
    Task<int> InsertAsync(bool overwriteSheet = false, CancellationToken cancellationToken = default);
}