using System;

namespace MiniExcelLibs.Enums;

/// <summary>
/// Excel image display mode (whether the image aligns/scales with cells).
/// </summary>
public enum XlsxImgType
{
    /// <summary>
    /// The image moves with the cell but does not scale (OneCellAnchor).
    /// Typically used when the image is bound to a single starting cell.
    /// </summary>
    OneCellAnchor,

    /// <summary>
    /// The image floats over the worksheet, maintaining a fixed position regardless of cell changes (AbsoluteAnchor).
    /// </summary>
    AbsoluteAnchor,

    /// <summary>
    /// The image is embedded in the cell and moves and scales with the cell (TwoCellAnchor).
    /// </summary>
    TwoCellAnchor
}