namespace MiniExcelLib.OpenXml.Styles;

internal class SheetStyleBuildResult
{
    public SheetStyleBuildResult(Dictionary<string, string> cellXfIdMap)
    {
        CellXfIdMap = cellXfIdMap;
    }

    public Dictionary<string, string> CellXfIdMap { get; set; }
}