namespace MiniExcel
{
    using System.Xml.Linq;

    internal static class ExcelNamespaces
    {
        internal static XNamespace excelNamespace = XNamespace.Get("http://schemas.openxmlformats.org/spreadsheetml/2006/main");
        internal static XNamespace excelRelationshipsNamepace = XNamespace.Get("http://schemas.openxmlformats.org/officeDocument/2006/relationships");
    }
}
