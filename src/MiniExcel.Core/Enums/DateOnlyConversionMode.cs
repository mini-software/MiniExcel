namespace MiniExcelLib.Core.Enums;


public enum DateOnlyConversionMode
{
    /// <summary>
    /// No conversion is applied and DateOnly values are not transformed.
    /// </summary>
    None,

    /// <summary>
    /// Allows conversion from DateTime to DateOnly only if the time component is exactly midnight (00:00:00).
    /// </summary>
    RequireMidnight,

    /// <summary>
    /// Converts DateTime to DateOnly by ignoring the time part completely, assuming the time component is not critical.
    /// </summary>
    IgnoreTimePart
}