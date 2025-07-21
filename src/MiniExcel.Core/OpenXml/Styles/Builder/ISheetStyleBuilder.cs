namespace MiniExcelLib.Core.OpenXml.Styles.Builder;

internal partial interface ISheetStyleBuilder
{
    [CreateSyncVersion]
    Task<SheetStyleBuildResult> BuildAsync(CancellationToken cancellationToken = default);
}