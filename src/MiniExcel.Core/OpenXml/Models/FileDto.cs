namespace MiniExcelLib.Core.OpenXml.Models;

internal class FileDto
{
    internal string ID { get; set; } = $"R{Guid.NewGuid():N}";
    internal string Extension { get; set; }
    internal string Path => $"xl/media/{ID}.{Extension}";
    internal string Path2 => $"/xl/media/{ID}.{Extension}";
    internal byte[] Byte { get; set; }
    internal int RowIndex { get; set; }
    internal int CellIndex { get; set; }
    internal bool IsImage { get; set; }
    internal int SheetId { get; set; }
}