using MiniExcelLib.Core.Helpers;
using MiniExcelLib.Tests.Common.Utils;

namespace MiniExcelLib.Tests;

public class MiniExcelOpenXmlConfigurationTest
{
    private readonly OpenXmlImporter _excelImporter =  MiniExcel.Importers.GetOpenXmlImporter();
    private readonly OpenXmlExporter _excelExporter =  MiniExcel.Exporters.GetOpenXmlExporter();    
    [Fact]
    public async Task EnableWriteFilePathTest()
    {
        var img = await new HttpClient().GetByteArrayAsync("https://user-images.githubusercontent.com/12729184/150462383-ad9931b3-ed8d-4221-a1d6-66f799743433.png");
        ImgExportTestDto[] value =
        [
            new() { Name = "github", Img = await File.ReadAllBytesAsync(PathHelper.GetFile("images/github_logo.png")) },
            new() { Name = "google", Img = await File.ReadAllBytesAsync(PathHelper.GetFile("images/google_logo.png")) },
            new() { Name = "microsoft", Img = await File.ReadAllBytesAsync(PathHelper.GetFile("images/microsoft_logo.png")) },
            new() { Name = "reddit", Img = await File.ReadAllBytesAsync(PathHelper.GetFile("images/reddit_logo.png")) },
            new() { Name = "statck_overflow", Img = await File.ReadAllBytesAsync(PathHelper.GetFile("images/statck_overflow_logo.png")) },
            new() { Name = "statck_over", Img = img }
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

