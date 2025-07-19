namespace MiniExcelLib.OpenXml.Models;

internal class ExcellSheetInfo
{
    public object Key { get; set; }
    public string? ExcelSheetName { get; set; }
    public SheetState ExcelSheetState { get; set; }
        
    private string ExcelSheetStateAsString => ExcelSheetState.ToString().ToLower();

    public SheetDto ToDto(int sheetIndex)
    {
        return new SheetDto { Name = ExcelSheetName, SheetIdx = sheetIndex, State = ExcelSheetStateAsString };
    }
}