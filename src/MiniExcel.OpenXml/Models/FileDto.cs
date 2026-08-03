namespace MiniExcelLib.OpenXml.Models;

internal class FileDto
{
    internal int SheetIndex { get; set; }
    internal int RowIndex { get; set; }
    internal int CellIndex { get; set; }
    internal string Id => $"rFileId_{SheetIndex}_{RowIndex + 1}_{CellIndex + 1}";
    internal string Path => $"xl/media/{Id}.{Extension}";
    internal bool IsImage { get; set; }
    internal string Extension { get; set; }
    internal byte[] Contents { get; set; }
}