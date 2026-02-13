namespace MiniExcelLib.OpenXml.Styles.Builder;

internal class SheetStyleBuildResult(Dictionary<string, string> cellXfIdMap)
{
    public Dictionary<string, string> CellXfIdMap { get; set; } = cellXfIdMap;
}