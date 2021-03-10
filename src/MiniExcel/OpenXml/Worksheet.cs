namespace MiniExcelLibs.OpenXml
{
    using System.Collections.Generic;
    internal class Worksheet
    {
        public int RowCount { get; set; }
        public int FieldCount { get; set; }
        public Dictionary<int, Dictionary<int, object>> Rows { get; set; }
    }
}
