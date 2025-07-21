using MiniExcelLib.Core.OpenXml.Attributes;
using MiniExcelLib.Core.OpenXml.Styles;

namespace MiniExcelLib.Core.OpenXml;

public class OpenXmlConfiguration : MiniExcelBaseConfiguration
{
    internal static OpenXmlConfiguration Default => new();
    
    public bool FillMergedCells { get; set; }
    public TableStyles TableStyles { get; set; } = TableStyles.Default;
    public bool AutoFilter { get; set; } = true;
    public int FreezeRowCount { get; set; } = 1;
    public int FreezeColumnCount { get; set; } = 0;
    public bool EnableConvertByteArray { get; set; } = true;
    public bool IgnoreTemplateParameterMissing { get; set; } = true;
    public bool EnableWriteNullValueCell { get; set; } = true;
    public bool WriteEmptyStringAsNull { get; set; } = false;
    public bool TrimColumnNames { get; set; } = true;
    public bool IgnoreEmptyRows { get; set; } = false;
    public bool EnableSharedStringCache { get; set; } = true;
    public long SharedStringCacheSize { get; set; } = 5 * 1024 * 1024;
    public OpenXmlStyleOptions StyleOptions { get; set; } = new();
    public DynamicExcelSheet[]? DynamicSheets { get; set; }
    public bool EnableWriteFilePath{ get; set; } = true;
    
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