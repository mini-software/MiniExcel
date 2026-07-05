namespace MiniExcelLib.OpenXml.Models;

internal class SheetDto
{
    internal int SheetIdx { get; set; }
    internal string Id => $"rSheetId{SheetIdx}";
    internal string? Name { get; set; }
    internal string Path => ExcelFileNames.Worksheet(SheetIdx);
    internal string State { get; set; }
}
