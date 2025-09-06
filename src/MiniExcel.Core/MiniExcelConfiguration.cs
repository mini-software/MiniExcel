using MiniExcelLib.Core.Attributes;

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

    /// <summary>
    /// Sets the options to how and when encountered DateTime values can be converted to DateOnly values.
    /// </summary>
    public DateOnlyConversionMode DateOnlyConversionMode { get; set; } = DateOnlyConversionMode.None;
}


/// <summary>
/// Specifies how DateTime values should be converted to DateOnly.
/// </summary>
public enum DateOnlyConversionMode
{
    /// <summary>
    /// No conversion is applied; DateOnly values are not transformed.
    /// </summary>
    None,

    /// <summary>
    /// Converts DateTime to DateOnly by enforcing midnight (00:00:00) as the time component.
    /// </summary>
    EnforceMidnight,

    /// <summary>
    /// Converts DateTime to DateOnly by ignoring the time part completely, assuming the time component is not critical.
    /// </summary>
    IgnoreTimePart
}
