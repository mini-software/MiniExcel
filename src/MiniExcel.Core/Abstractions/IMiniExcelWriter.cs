namespace MiniExcelLib.Core.Abstractions;

public partial interface IMiniExcelWriter
{
    [CreateSyncVersion]
    Task<int[]> SaveAsAsync(IProgress<int>? progress = null, CancellationToken cancellationToken = default);

    [CreateSyncVersion]
    Task<int> InsertAsync(bool overwriteSheet = false, IProgress<int>? progress = null, CancellationToken cancellationToken = default);
}