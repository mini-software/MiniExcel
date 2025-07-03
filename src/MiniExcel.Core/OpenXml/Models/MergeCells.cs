namespace MiniExcelLib.Core.OpenXml.Models;

internal class MergeCells
{
    public Dictionary<string, object?> MergesValues { get; set; } = [];
    public Dictionary<string, string> MergesMap { get; set; } =  [];
}