using MiniExcelLib.OpenXml.Constants;

namespace MiniExcelLib.OpenXml.Models;

internal class SheetDto
{
    internal string ID { get; set; } = $"R{Guid.NewGuid():N}";
    internal string? Name { get; set; }
    internal int SheetIdx { get; set; }
    internal string Path => ExcelFileNames.Worksheet(SheetIdx);

    internal string State { get; set; }
}
