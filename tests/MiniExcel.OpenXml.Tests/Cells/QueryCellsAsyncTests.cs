using System.Drawing;
using OfficeOpenXml.Drawing;
using LicenseContext = OfficeOpenXml.LicenseContext;

namespace MiniExcelLib.OpenXml.Tests.Cells;

public class QueryCellsAsyncTests
{
    private readonly OpenXmlImporter _excelImporter = MiniExcelV2.Importers.GetOpenXmlImporter();
    private readonly OpenXmlExporter _excelExporter = MiniExcelV2.Exporters.GetOpenXmlExporter();

    public QueryCellsAsyncTests()
    {
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
    }

    internal static MemoryStream CreateWorkbook()
    {
        using var package = new ExcelPackage();
        var sheet = package.Workbook.Worksheets.Add("Sheet1");

        sheet.Cells["A1"].Value = 1;
        sheet.Cells["A2"].Value = 2;
        sheet.Cells["A2"].Style.Font.Color.SetColor(Color.Red);
        sheet.Cells["A3"].Value = 3;
        sheet.Cells["A3"].Style.Font.Color.SetColor(eThemeSchemeColor.Text1);
        sheet.Cells["A4"].Value = "text";
        sheet.Cells["A4"].Style.Font.Color.SetColor(eThemeSchemeColor.Accent2);
        sheet.Cells["A4"].Style.Font.Color.Tint = 0.4m;

        // Assigning one formula to a range makes EPPlus store it as a shared formula
        sheet.Cells["B1:B3"].Formula = "A1*2";
        sheet.Cells["C1"].Formula = "\"x\"&A1";

        package.Workbook.Calculate();

        var stream = new MemoryStream();
        package.SaveAs(stream);
        stream.Position = 0;
        return stream;
    }

    [Fact]
    public async Task QueryCells_ReturnsValuesAndReferencesAsync()
    {
        using var stream = CreateWorkbook();
        var cells = await _excelImporter.QueryCellsAsync(stream).ToDictionaryAsync(c => c.Reference);

        Assert.Equal(1d, cells["A1"].Value);
        Assert.Equal(1, cells["A1"].RowIndex);
        Assert.Equal(1, cells["A1"].ColumnIndex);
        Assert.Equal("text", cells["A4"].Value);
        Assert.Equal(4, cells["A4"].RowIndex);
        Assert.Equal(3, cells["C1"].ColumnIndex);
    }

    [Fact]
    public async Task QueryCells_ReportsFormulasAsync()
    {
        using var stream = CreateWorkbook();
        var cells = await _excelImporter.QueryCellsAsync(stream).ToDictionaryAsync(c => c.Reference);

        Assert.False(cells["A1"].HasFormula);
        Assert.Null(cells["A1"].Formula);

        Assert.True(cells["B1"].HasFormula);
        Assert.Equal("A1*2", cells["B1"].Formula);
        Assert.Equal(2d, cells["B1"].Value);

        // Cells sharing the formula of B1 carry no formula text of their own
        Assert.True(cells["B3"].HasFormula);
        Assert.Null(cells["B3"].Formula);
        Assert.Equal(6d, cells["B3"].Value);

        Assert.True(cells["C1"].HasFormula);
        Assert.Equal("\"x\"&A1", cells["C1"].Formula);
        Assert.Equal("x1", cells["C1"].Value);
    }

    [Fact]
    public async Task QueryCells_ReportsFontColorsAsync()
    {
        using var stream = CreateWorkbook();
        var cells = await _excelImporter.QueryCellsAsync(stream).ToDictionaryAsync(c => c.Reference);

        Assert.Equal("FFFF0000", cells["A2"].FontColor?.Rgb);
        Assert.Null(cells["A2"].FontColor?.Theme);

        Assert.Equal(1, cells["A3"].FontColor?.Theme);
        Assert.Null(cells["A3"].FontColor?.Rgb);
        Assert.Equal(0, cells["A3"].FontColor?.Tint);

        Assert.Equal(5, cells["A4"].FontColor?.Theme);
        Assert.Equal(0.4, cells["A4"].FontColor!.Tint, 6);

        Assert.Null(cells["A1"].FontColor?.Rgb);
        Assert.Null(cells["A1"].FontColor?.Theme);
    }

    [Fact]
    public async Task QueryCells_SelectsSheetByNameAsync()
    {
        using var stream = new MemoryStream();
        var sheets = new Dictionary<string, object>
        {
            ["First"] = new[] { new { Value = "first" } },
            ["Second"] = new[] { new { Value = "second" } }
        };
        await _excelExporter.ExportAsync(stream, sheets);
        stream.Position = 0;

        var cells = await _excelImporter.QueryCellsAsync(stream, "Second").ToListAsync();

        Assert.Equal(["A1", "A2"], cells.Select(c => c.Reference));
        Assert.Equal(["Value", "second"], cells.Select(c => c.Value));
        Assert.All(cells, c => Assert.False(c.HasFormula));
    }
}
