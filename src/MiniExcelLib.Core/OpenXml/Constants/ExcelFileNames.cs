namespace MiniExcelLib.Core.OpenXml.Constants;

internal static class ExcelFileNames
{
    internal const string Rels = "_rels/.rels";
    internal const string SharedStrings = "xl/sharedStrings.xml";

    internal const string ContentTypes = "[Content_Types].xml";
    internal const string Styles = "xl/styles.xml";
    internal const string Workbook = "xl/workbook.xml";
    internal const string WorkbookRels = "xl/_rels/workbook.xml.rels";
        
    internal static string SheetRels(int sheetId) => $"xl/worksheets/_rels/sheet{sheetId}.xml.rels";
    internal static string Drawing(int sheetIndex) => $"xl/drawings/drawing{sheetIndex + 1}.xml";
    internal static string DrawingRels(int sheetIndex) => $"xl/drawings/_rels/drawing{sheetIndex + 1}.xml.rels";
}