using System;

namespace MiniExcelLibs.OpenXml.Models
{
    internal class FileDto
    {
        public string ID { get; set; } = $"R{Guid.NewGuid():N}";
        public string Extension { get; set; }
        public string Path { get { return $"xl/media/{ID}.{Extension}"; } }
        public string Path2 { get { return $"/xl/media/{ID}.{Extension}"; } }
        public byte[] Byte { get; set; }
        public int RowIndex { get; set; }
        public int CellIndex { get; set; }
        public bool IsImage { get; set; } = false;
        public int SheetId { get; set; }
    }
}