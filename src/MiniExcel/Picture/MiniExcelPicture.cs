namespace MiniExcelLibs.Picture
{
    public class MiniExcelPicture
    {
        public byte[] ImageBytes { get; set; }
        public string SheetName { get; set; }
        public string PictureType { get; set; }
        public int ColumnNumber { get; set; }
        public int RowNumber { get; set; }
        public int WidthPx { get; set; }
        public int HeightPx { get; set; }
    }
}