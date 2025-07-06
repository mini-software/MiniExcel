namespace MiniExcelLib.OpenXml.Styles;

internal partial interface ISheetStyleBuilder
{
    [Zomp.SyncMethodGenerator.CreateSyncVersion]
    Task<SheetStyleBuildResult> BuildAsync(CancellationToken cancellationToken = default);
}