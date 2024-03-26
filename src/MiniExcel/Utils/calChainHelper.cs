using System.IO;
using System.Text;
using System.Xml;

namespace MiniExcelLibs.Utils
{
    internal static class CalcChainHelper
    {

        // The calcChain.xml file in an Excel file (in the xl folder) is an XML file that stores the calculation chain for the workbook.
        // The calculation chain specifies the order in which cells should be recalculated in order to update all formulas in the workbook correctly.
        // It should include a series of <c> elements, each of which represents a cell in the workbook that contains a formula.
        //      Each <c> element should have a r attribute that specifies the cell's address (e.g., "A1" or "B2").
        //      The  <c> element should also have a i attribute that specifies the index of the formula in the formulas collection (in the workbook's sheet data file).
        // https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.calculationchain?view=openxml-2.8.1
        public static string GetCalcChainContentFromSheet(in XmlNode sheetData, XmlNamespaceManager ns, int sheetIndex)
        {

            StringBuilder calcChainContent = new StringBuilder();

            // each c having f nodes
            var cs = sheetData.SelectNodes($"x:row/x:c[./x:f]", ns);
            foreach (XmlElement c in cs)
            {
                calcChainContent.Append($@"<c r=""{c.GetAttribute("r")}"" i=""{sheetIndex}""/>");
            }

            return calcChainContent.ToString();
        }

        public static void GenerateCalcChainSheet(Stream calcChainStream, string calcChainContent)
        {
            using (var writer = new StreamWriter(calcChainStream, Encoding.UTF8))
            {
                writer.Write($"<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><calcChain xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">{calcChainContent}</calcChain>");
            }
        }
    }
}
