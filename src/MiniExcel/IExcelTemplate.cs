using System.Threading;
using System.Threading.Tasks;
using Zomp.SyncMethodGenerator;

namespace MiniExcelLibs;

internal partial interface IExcelTemplate
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