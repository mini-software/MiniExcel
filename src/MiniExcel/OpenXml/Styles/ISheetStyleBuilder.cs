namespace MiniExcelLibs.OpenXml.Styles;

internal interface ISheetStyleBuilder
{
    void Build();

    Task BuildAsync(CancellationToken cancellationToken = default);
}