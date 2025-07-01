using MiniExcelLibs.Enums;
using MiniExcelLibs.Utils;
using System.Drawing;

namespace MiniExcelLibs.Picture;

public class MiniExcelPicture
{
    public byte[] ImageBytes { get; set; } = [];
    public string? SheetName { get; set; }
    public string? PictureType { get; set; }
    public string? CellAddress { get; set; }
    /// <summary>
    /// Only takes effect when the image is in AbsoluteAnchor floating mode
    /// </summary>
	public Point Location { get; set; }
	public XlsxImgType ImgType { get; set; }
    internal int ColumnNumber => ReferenceHelper.ConvertCellToXY(CellAddress).Item1 -1;
    internal int RowNumber => ReferenceHelper.ConvertCellToXY(CellAddress).Item2 - 1;
    
    public int WidthPx { get; set; } = 80;
    public int HeightPx { get; set; } = 24;
}