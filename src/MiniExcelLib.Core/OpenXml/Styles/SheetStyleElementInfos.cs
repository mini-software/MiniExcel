namespace MiniExcelLib.Core.OpenXml.Styles;

public class SheetStyleElementInfos
{
    public bool ExistsNumFmts { get; set; }
    public int NumFmtCount { get; set; }
    public bool ExistsFonts { get; set; }
    public int FontCount { get; set; }
    public bool ExistsFills { get; set; }
    public int FillCount { get; set; }
    public bool ExistsBorders { get; set; }
    public int BorderCount { get; set; }
    public bool ExistsCellStyleXfs { get; set; }
    public int CellStyleXfCount { get; set; }
    public bool ExistsCellXfs { get; set; }
    public int CellXfCount { get; set; }
}