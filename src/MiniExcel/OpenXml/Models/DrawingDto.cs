using System;

namespace MiniExcelLibs.OpenXml.Models;

internal class DrawingDto
{
    internal string ID { get; set; } = $"R{Guid.NewGuid():N}";
}