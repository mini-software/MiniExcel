using System;

namespace MiniExcelLibs.OpenXml.Models
{
    internal class SheetDto
    {
        public string ID { get; set; } = $"R{Guid.NewGuid():N}";
        public string Name { get; set; }
        public int SheetIdx { get; set; }
        public string Path { get { return $"xl/worksheets/sheet{SheetIdx}.xml"; } }

        public string State { get; set; }
    }
}