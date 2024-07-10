using MiniExcelLibs.Attributes;

namespace MiniExcelLibs.OpenXml
{
    public class OpenXmlConfiguration : Configuration
    {
        internal static readonly OpenXmlConfiguration DefaultConfig = new OpenXmlConfiguration();
        public bool FillMergedCells { get; set; }
        public TableStyles TableStyles { get; set; } = TableStyles.Default;
        public bool AutoFilter { get; set; } = true;
        public bool FreezeTopRow { get; set; } = true;
        public bool EnableConvertByteArray { get; set; } = true;
        public bool IgnoreTemplateParameterMissing { get; set; } = true;
        public bool EnableWriteNullValueCell { get; set; } = true;
        public bool EnableSharedStringCache { get; set; } = true;
        public long SharedStringCacheSize { get; set; } = 5 * 1024 * 1024;
        public DynamicExcelSheet[] DynamicSheets { get; set; }
    }
}