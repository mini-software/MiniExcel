using MiniExcelLib.Core.Helpers;
using MiniExcelLib.Tests.Common.Utils;

namespace MiniExcelLib.OpenXml.Tests;

public class MiniExcelOpenXmlConfigurationTest
{
    private readonly OpenXmlImporter _excelImporter =  MiniExcel.Importers.GetOpenXmlImporter();
    private readonly OpenXmlExporter _excelExporter =  MiniExcel.Exporters.GetOpenXmlExporter();    
    
    [Fact]
    public async Task DisableWriteFilePathTest()
    {
        ImgExportTestDto[] value =
        [
            new() { Name = "github", Img = await File.ReadAllBytesAsync(PathHelper.GetFile("images/github_logo.png")) },
            new() { Name = "google", Img = await File.ReadAllBytesAsync(PathHelper.GetFile("images/google_logo.png")) },
            new() { Name = "microsoft", Img = await File.ReadAllBytesAsync(PathHelper.GetFile("images/microsoft_logo.png")) },
            new() { Name = "reddit", Img = await File.ReadAllBytesAsync(PathHelper.GetFile("images/reddit_logo.png")) },
        ];

        var path = PathHelper.GetFile("xlsx/Test_EnableWriteFilePath.xlsx");
        await _excelExporter.ExportAsync(path, value, configuration: new OpenXmlConfiguration { EnableWriteFilePath = false }, overwriteFile: true);
        Assert.True(File.Exists(path));

        var rows = await _excelImporter.QueryAsync<ImgExportTestDto>(path).CreateListAsync();
        Assert.True(rows.All(x => x.Img is null or []));
    }
    
    private class ImgExportTestDto
    {
        public string? Name { get; set; }

        [MiniExcelColumn(Name = "图片", Width = 100)]
        public byte[]? Img { get; set; }
    }
}
