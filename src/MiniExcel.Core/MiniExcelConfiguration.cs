namespace MiniExcelLib.Core;

public interface IMiniExcelConfiguration;

public abstract class MiniExcelBaseConfiguration : IMiniExcelConfiguration
{
    public CultureInfo Culture { get; set; } = CultureInfo.InvariantCulture;
    public DynamicExcelColumn[]? DynamicColumns { get; set; } = [];
    public int BufferSize { get; set; } = 1024 * 512;
    public bool FastMode { get; set; }
        
    /// <summary>
    /// When exporting using DataReader, the data not in DynamicColumn will be filtered.
    /// </summary>
    public bool DynamicColumnFirst { get; set; } = false;
}