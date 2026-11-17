namespace MiniExcelLib.OpenXml.Models;

internal class ExcelSheetInfo
{
    public object Key { get; set; }
    public string? ExcelSheetName { get; set; }
    public SheetState ExcelSheetState { get; set; }

    public SheetDto ToDto(int sheetIndex) => new()
    {
        Name = ExcelSheetName,
        SheetIdx = sheetIndex, 
        State = ExcelSheetState switch
        {
            SheetState.Hidden => "hidden",
            SheetState.VeryHidden => "veryHidden",
            _ => "visible"
        }
    };
}