namespace MiniExcelLib.Core.Enums;

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
    /// Drawing anchored between two cells; moves and scales with the cells (TwoCellAnchor).
    /// Still a floating drawing — not Excel 365 Place in Cell. Prefer <see cref="PlaceInCell"/> for true in-cell images.
    /// </summary>
    TwoCellAnchor,

    /// <summary>
    /// Excel 365 "Place in Cell": the image is stored as the cell value via richData (not a drawing).
    /// Requires Microsoft 365; older Excel versions show #VALUE! in the cell.
    /// WidthPx, HeightPx, and Location are ignored — size follows the cell.
    /// </summary>
    PlaceInCell
}