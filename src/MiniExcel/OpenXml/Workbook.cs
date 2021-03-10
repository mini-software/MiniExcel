using System.Collections.Generic;
using System.Linq;

namespace MiniExcelLibs.OpenXml
{
    public class Workbook
    {
        public IEnumerable<Worksheet> OpenXmlWorksheet { get; set; }
        public Worksheet GetWorksheet(int index)
        {
            return (OpenXmlWorksheet as IList<Worksheet>)[index];
        }

        public Worksheet GetWorksheet(string sheetName)
        {
            return OpenXmlWorksheet.Single(w => w.Name == sheetName);
        }
    }
}
