using System;

namespace MiniExcelLibs.OpenXml.Models
{
    internal class DrawingDto
    {
        public string ID { get; set; } = $"R{Guid.NewGuid():N}";
    }
}