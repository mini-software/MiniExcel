namespace MiniExcelLib.OpenXml.Tests.Configuration;

internal static class AutoAdjustTestParameters
{
    internal const int Column1MaLen = 32;
    internal const int Column2MaxLen = 16;
    private const int Column3MaxLen = 2;
    private const int Column4MaxLen = 100;

    public static List<string[]> GetTestData() => [
        [
            new('1', Column1MaLen), 
            new('2', Column2MaxLen / 2),
            new('3', Column3MaxLen / 2),
            new('4', Column4MaxLen)
        ],
        [
            new('1', Column1MaLen / 2), 
            new('2', Column2MaxLen),
            new('3', Column3MaxLen), 
            new('4', Column4MaxLen)
        ] ];

    public static List<Dictionary<string, object>> GetDictionaryTestData() => GetTestData()
        .Select(row => row
            .Select((value, i) => (value, i))
            .ToDictionary(x => $"Column{x.i}", object (x) => x.value))
        .ToList();

    public static OpenXmlConfiguration GetConfiguration() => new()
    {
        EnableAutoWidth = true,
        FastMode = true,
        MaxWidth = 50
    };
}

internal class ImgExportTestDto
{
    public string? Name { get; set; }

    [MiniExcelColumn(Name = "图片", Width = 100)]
    public byte[]? Img { get; set; }
}
