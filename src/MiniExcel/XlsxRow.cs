namespace MiniExcelLibs
{
    using System.Collections.Generic;
    public class XlsxRow
    {
        public string RowNumber { get; set; }
        public IEnumerable<XlsxCell> Cells { get; set; }
    }
}
