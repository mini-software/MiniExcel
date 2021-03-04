namespace MiniExcel
{
    using System.Xml.Linq;
    public static partial class MiniExcelHelper
    {
        internal static class ExcelNamespaces
        {
            internal static XNamespace excelNamespace = XNamespace.Get("http://schemas.openxmlformats.org/spreadsheetml/2006/main");
            internal static XNamespace excelRelationshipsNamepace = XNamespace.Get("http://schemas.openxmlformats.org/officeDocument/2006/relationships");
        }
    }
}
