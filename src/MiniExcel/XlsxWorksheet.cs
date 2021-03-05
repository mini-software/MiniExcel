namespace MiniExcelLibs
{
    using System.Collections.Generic;
    public class XlsxWorksheet
    {
        public string ID { get; set; }
        public string SheetID { get; set; }
        public string Name { get; set; }
        public IEnumerable<XlsxRow> Rows { get; set; }
    }
}
