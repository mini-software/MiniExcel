using System.Drawing;
using MiniExcelLib.Core.Enums;

namespace MiniExcelLib.OpenXml.Picture;

public class MiniExcelPicture
{
    public byte[] ImageBytes { get; set; } = [];
    public string? SheetName { get; set; }
    public string? PictureType { get; set; }
    public string? CellAddress { get; set; }

    /// <summary>
    /// Only takes effect when the image is in AbsoluteAnchor floating mode.
    /// Ignored for <see cref="XlsxImgType.PlaceInCell"/>.
    /// </summary>
    public Point Location { get; set; }

    /// <summary>
    /// Drawing anchor mode, or <see cref="XlsxImgType.PlaceInCell"/> for Excel 365 in-cell images.
    /// </summary>
    public XlsxImgType ImgType { get; set; }

    internal int ColumnNumber => CellReferenceConverter.TryParseCellReference(CellAddress, out var column, out _)
        ? column - 1
        : throw new InvalidDataException($"Value {CellAddress} is not a valid cell reference.");

    internal int RowNumber => CellReferenceConverter.TryParseCellReference(CellAddress, out var _, out var row)
        ? row - 1
        : throw new InvalidDataException($"Value {CellAddress} is not a valid cell reference.");

    /// <summary>
    /// Image width in pixels for drawing anchors. Ignored for <see cref="XlsxImgType.PlaceInCell"/>.
    /// </summary>
    public int WidthPx { get; set; } = 80;

    /// <summary>
    /// Image height in pixels for drawing anchors. Ignored for <see cref="XlsxImgType.PlaceInCell"/>.
    /// </summary>
    public int HeightPx { get; set; } = 24;
}
