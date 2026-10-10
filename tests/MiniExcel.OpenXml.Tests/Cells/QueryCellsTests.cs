using LicenseContext = OfficeOpenXml.LicenseContext;

namespace MiniExcelLib.OpenXml.Tests.Cells;

public class QueryCellsTests
{
    private readonly OpenXmlImporter _excelImporter = MiniExcelV2.Importers.GetOpenXmlImporter();

    public QueryCellsTests()
    {
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
    }

    [Fact]
    public void QueryCells_ReportsFormulasAndFontColors()
    {
        using var stream = QueryCellsAsyncTests.CreateWorkbook();
        var cells = _excelImporter.QueryCells(stream).ToDictionary(c => c.Reference);

        Assert.False(cells["A1"].HasFormula);
        Assert.True(cells["B1"].HasFormula);
        Assert.Equal("A1*2", cells["B1"].Formula);
        Assert.Equal(2d, cells["B1"].Value);
        Assert.True(cells["B2"].HasFormula);
        Assert.Null(cells["B2"].Formula);

        Assert.Equal("FFFF0000", cells["A2"].FontColor?.Rgb);
        Assert.Equal(1, cells["A3"].FontColor?.Theme);
    }
}
