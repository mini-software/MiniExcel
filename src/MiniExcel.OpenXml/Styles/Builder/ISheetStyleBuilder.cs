namespace MiniExcelLib.OpenXml.Styles.Builder;

internal partial interface ISheetStyleBuilder
{
    [CreateSyncVersion]
    Task BuildAsync(CancellationToken cancellationToken = default);
}
