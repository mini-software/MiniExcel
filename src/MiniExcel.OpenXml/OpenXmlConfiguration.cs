using MiniExcelLib.Core;
using MiniExcelLib.OpenXml.Styles;

namespace MiniExcelLib.OpenXml;

public class OpenXmlConfiguration : MiniExcelBaseConfiguration
{
    internal static OpenXmlConfiguration Default => new();
    
    public bool FillMergedCells { get; set; }
    public TableStyles TableStyles { get; set; } = TableStyles.Default;
    public bool AutoFilter { get; set; } = true;
    public bool RightToLeft { get; set; } = false;
    public int FreezeRowCount { get; set; } = 1;
    public int FreezeColumnCount { get; set; } = 0;
    public bool EnableConvertByteArray { get; set; } = true;
    public bool EnableWriteFilePath{ get; set; } = true;
    public bool IgnoreTemplateParameterMissing { get; set; } = true;
    public bool EnableWriteNullValueCell { get; set; } = true;
    public bool WriteEmptyStringAsNull { get; set; } = false;
    public bool TrimColumnNames { get; set; } = true;
    public bool IgnoreEmptyRows { get; set; } = false;
    
    public bool EnableSharedStringCache { get; set; } = true;
    public long SharedStringCacheSize { get; set; } = 5 * 1024 * 1024;
        
    /// <summary>
    /// The directory where the shared strings cache files are stored.
    /// It defaults to the system's temporary folder.
    /// </summary>
    public string SharedStringCachePath { get; set; } = Path.GetTempPath();

    public OpenXmlStyleOptions StyleOptions { get; set; } = new();
    public DynamicExcelSheetAttribute[]? DynamicSheets { get; set; }
    
    /// <summary>
    /// Calculate column widths automatically from each column value.
    /// </summary>
    public bool EnableAutoWidth { get; set; }
    public double MinWidth { get; set; } = 9.28515625;
    public double MaxWidth { get; set; } = 200;
}

public enum TableStyles
{
    None,
    Default
}
