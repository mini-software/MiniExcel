using System.Threading;
using System.Threading.Tasks;

namespace MiniExcelLibs
{
    internal partial interface IExcelTemplateAsync
    {
        [Zomp.SyncMethodGenerator.CreateSyncVersion]
        Task SaveAsByTemplateAsync(string templatePath, object value, CancellationToken cancellationToken = default(CancellationToken));
        [Zomp.SyncMethodGenerator.CreateSyncVersion]
        Task SaveAsByTemplateAsync(byte[] templateBytes, object value, CancellationToken cancellationToken = default(CancellationToken));
        [Zomp.SyncMethodGenerator.CreateSyncVersion]
        Task MergeSameCellsAsync(string path, CancellationToken cancellationToken = default(CancellationToken));
        [Zomp.SyncMethodGenerator.CreateSyncVersion]
        Task MergeSameCellsAsync(byte[] fileInBytes, CancellationToken cancellationToken = default(CancellationToken));
    }
}
