namespace MiniExcelLibs.OpenXml
{
    using System.Xml.Linq;
    internal static class ExcelOpenXmlXName
    {
        internal readonly static XNamespace ExcelNamespace = XNamespace.Get("http://schemas.openxmlformats.org/spreadsheetml/2006/main");
        internal readonly static XNamespace ExcelRelationshipsNamepace = XNamespace.Get("http://schemas.openxmlformats.org/officeDocument/2006/relationships");
        internal readonly static XName Row;
        internal readonly static XName R;
        internal readonly static XName V;
        internal readonly static XName T;
        internal readonly static XName C;
        internal readonly static XName Dimension;
        internal readonly static XName Sheet;
        static ExcelOpenXmlXName()
        {
            Row = XNamespace.Get("http://schemas.openxmlformats.org/spreadsheetml/2006/main") + "row";
            R = XNamespace.Get("http://schemas.openxmlformats.org/spreadsheetml/2006/main") + "r";
            V = XNamespace.Get("http://schemas.openxmlformats.org/spreadsheetml/2006/main") + "v";
            T = XNamespace.Get("http://schemas.openxmlformats.org/spreadsheetml/2006/main") + "t";
            C = XNamespace.Get("http://schemas.openxmlformats.org/spreadsheetml/2006/main") + "c";
            Dimension = XNamespace.Get("http://schemas.openxmlformats.org/spreadsheetml/2006/main") + "dimension";
            Sheet = XNamespace.Get("http://schemas.openxmlformats.org/spreadsheetml/2006/main") + "sheet";
        }
    }


}