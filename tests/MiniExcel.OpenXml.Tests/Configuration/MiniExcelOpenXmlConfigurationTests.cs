using MiniExcelLib.Tests.Common.Utils;

namespace MiniExcelLib.OpenXml.Tests.Configuration;

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

        using var path = AutoDeletingPath.Create();
        await _excelExporter.ExportAsync(path.ToString(), value, configuration: new OpenXmlConfiguration { EnableWriteFilePath = false }, overwriteFile: true);

        var rows = await _excelImporter.QueryAsync<ImgExportTestDto>(path.ToString()).ToListAsync();
        Assert.True(rows.All(x => x.Img is null or []));
    }
}
