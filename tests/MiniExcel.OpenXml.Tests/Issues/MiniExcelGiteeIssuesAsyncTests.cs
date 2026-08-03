using MiniExcelLib.Tests.Common.Utils;

namespace MiniExcelLib.OpenXml.Tests.Issues;

public class MiniExcelGiteeIssuesAsyncTests
{
    private readonly OpenXmlImporter _excelImporter =  MiniExcel.Importers.GetOpenXmlImporter();
    private readonly OpenXmlExporter _excelExporter =  MiniExcel.Exporters.GetOpenXmlExporter();

    // https://gitee.com/dotnetchina/MiniExcel/issues/I3OSKV
    // When exporting, the pure numeric string will be forcibly converted to a numeric type, resulting in the loss of the end data
    [Fact]
    public async Task IssueI3OSKV()
    {
        using var path1 = AutoDeletingPath.Create();
        var value1 = new[] { new { Test = "12345678901234567890" } };
        await _excelExporter.ExportAsync(path1.ToString(), value1);

        var result1 = await _excelImporter.QueryAsync(path1.ToString(), true).FirstAsync();
        Assert.Equal("12345678901234567890", result1.Test);

        using var path2 = AutoDeletingPath.Create();
        var value2 = new[] { new { Test = 123456.789 } };
        await _excelExporter.ExportAsync(path2.ToString(), value2);

        var result2 = await _excelImporter.QueryAsync(path2.ToString(), true).FirstAsync();
        Assert.Equal(123456.789, result2.Test);
    }
}
