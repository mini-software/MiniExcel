using System.Threading;
using System.Threading.Tasks;

namespace MiniExcelLibs.OpenXml.Styles;

internal partial interface ISheetStyleBuilder
{
    [Zomp.SyncMethodGenerator.CreateSyncVersion]
    Task<SheetStyleBuildResult> BuildAsync(CancellationToken cancellationToken = default);
}