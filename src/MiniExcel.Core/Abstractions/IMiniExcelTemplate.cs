using Zomp.SyncMethodGenerator;

namespace MiniExcelLib.Core.Abstractions;

public partial interface IMiniExcelTemplate
{
    [CreateSyncVersion]
    Task SaveAsByTemplateAsync(string templatePath, object value, CancellationToken cancellationToken = default);
    
    [CreateSyncVersion]
    Task SaveAsByTemplateAsync(byte[] templateBytes, object value, CancellationToken cancellationToken = default);
    
    [CreateSyncVersion]
    Task MergeSameCellsAsync(string path, CancellationToken cancellationToken = default);
    
    [CreateSyncVersion]
    Task MergeSameCellsAsync(byte[] fileInBytes, CancellationToken cancellationToken = default);
}