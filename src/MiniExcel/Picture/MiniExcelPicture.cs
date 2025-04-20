using MiniExcelLibs.Utils;

namespace MiniExcelLibs.Picture
{
    public class MiniExcelPicture
    {
        public byte[] ImageBytes { get; set; }
        public string SheetName { get; set; }
        public string PictureType { get; set; }
        public string CellAddress { get; set; }
        internal int ColumnNumber => ReferenceHelper.ConvertCellToXY(CellAddress).Item1 -1;
        internal int RowNumber => ReferenceHelper.ConvertCellToXY(CellAddress).Item2 - 1;
        public int WidthPx { get; set; }
        public int HeightPx { get; set; }
    }
}