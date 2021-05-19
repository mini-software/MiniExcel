
using System.ComponentModel;

namespace MiniExcelLibs.OpenXml
{
    public class OpenXmlConfiguration : IConfiguration
    {
        internal static readonly OpenXmlConfiguration DefaultConfig = new OpenXmlConfiguration();
        public bool FillMergedCells { get; set; }
        public TableStyles TableStyles { get; set; } = TableStyles.Default;
    }
}