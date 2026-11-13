using MiniExcelLib.OpenXml.Styles;

namespace MiniExcelLib.OpenXml.Models;

/// <summary>
/// A value to export together with a formula or style of its own, overriding the column's settings for that cell.
/// </summary>
/// <param name="value">The value written to the cell. When <see cref="Formula"/> is set, it is the cached result of the formula.</param>
public sealed class OpenXmlCellValue(object? value)
{
    /// <summary>
    /// The value written to the cell, or the cached result when <see cref="Formula"/> is set.
    /// </summary>
    public object? Value { get; } = value;

    /// <summary>
    /// A formula for the cell, without the leading equals sign.
    /// </summary>
    public string? Formula { get; set; }

    /// <summary>
    /// A number format for the cell, replacing the column's format.
    /// </summary>
    public string? NumberFormat { get; set; }

    /// <summary>
    /// The style of the cell, such as its font color.
    /// </summary>
    public OpenXmlCellStyle Style { get; } = new();
}
