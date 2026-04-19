namespace MiniExcelLibs.OpenXml.Styles;

internal interface ISheetStyleBuilder
{
    SheetStyleBuildResult Build();

    Task<SheetStyleBuildResult> BuildAsync(CancellationToken cancellationToken = default);
}