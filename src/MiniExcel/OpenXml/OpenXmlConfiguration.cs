
using System.ComponentModel;

namespace MiniExcelLibs.OpenXml
{
    public class OpenXmlConfiguration : Configuration
    {
        internal static readonly OpenXmlConfiguration DefaultConfig = new OpenXmlConfiguration();
        public bool FillMergedCells { get; set; }
        public TableStyles TableStyles { get; set; } = TableStyles.Default;
        public bool AutoFilter { get; set; } = true;
        public bool EnableConvertByteArray { get; set; } = true;
    }
}