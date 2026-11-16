namespace MiniExcelLib.OpenXml.Models;

internal class SheetDto
{
    internal int SheetIdx { get; set; }
    internal string Id => $"rSheetId{SheetIdx}";
    internal string? Name { get; set; }
    internal bool IsChartSheet { get; set; }
    internal string Path => IsChartSheet ? ExcelFileNames.ChartSheet(SheetIdx) : ExcelFileNames.Worksheet(SheetIdx);
    internal string RelsPath => IsChartSheet ? ExcelFileNames.ChartSheetRels(SheetIdx) : ExcelFileNames.SheetRels(SheetIdx);
    internal string State { get; set; }
}
