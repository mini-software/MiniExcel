
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
        public bool IgnoreTemplateParameterMissing { get; set; } = true;

        // currently, this is a preview functional
        public bool EnableSharedStringCache { get; set; } = false;
        public int SharedStringCacheSize { get; set; } = 5 * 1024 * 1024;
    }
}