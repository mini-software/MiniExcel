namespace MiniExcel
{
    using System.Collections.Generic;
    using System.Linq;
    public class XlsxWorkbook
    {
        public IEnumerable<XlsxWorksheet> Worksheets { get; set; }
        public XlsxWorksheet GetWorksheet(int index)
        {
            return (Worksheets as IList<XlsxWorksheet>)[index];
        }

        public XlsxWorksheet GetWorksheet(string sheetName)
        {
            return Worksheets.Single(w=>w.Name== sheetName);
        }
    }
}
