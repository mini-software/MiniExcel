namespace MiniExcelLib.Core.Abstractions;

public partial interface IMiniExcelWriter
{
    [CreateSyncVersion]
    Task<int[]> SaveAsAsync(CancellationToken cancellationToken = default, IProgress<int>? progress = null);

    [CreateSyncVersion]
    Task<int> InsertAsync(bool overwriteSheet = false, CancellationToken cancellationToken = default);
}