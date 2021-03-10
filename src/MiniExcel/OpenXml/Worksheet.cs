namespace MiniExcelLibs.OpenXml
{
    using System.Collections.Generic;
    public class Worksheet
    {
        public string ID { get; set; }
        public string SheetID { get; set; }
        public string Name { get; set; }
        public int RowCount { get; set; }
        public int FieldCount { get; set; }
        public Dictionary<int, Dictionary<int, object>> Rows { get; set; }
    }
}
