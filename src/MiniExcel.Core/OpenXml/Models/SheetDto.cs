namespace MiniExcelLib.Core.OpenXml.Models;

internal class SheetDto
{
    internal string ID { get; set; } = $"R{Guid.NewGuid():N}";
    internal string? Name { get; set; }
    internal int SheetIdx { get; set; }
    internal string Path => $"xl/worksheets/sheet{SheetIdx}.xml";

    internal string State { get; set; }
}