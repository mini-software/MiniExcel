using DocumentFormat.OpenXml.Office2010.Excel;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Wordprocessing;
using MiniExcelLib.OpenXml.Models;
using MiniExcelLib.Tests.Common.Utils;

namespace MiniExcelLib.OpenXml.Tests.SheetInformations;

public class MiniExcelOpenXmlSheetInformationsAsync
{
    private readonly OpenXmlImporter _excelImporter = MiniExcel.Importers.GetOpenXmlImporter();

    [Fact]
    public async Task GetSheetDimensionsAsyncTest()
    {
        var path = PathHelper.GetFile("xlsx/TestMultiSheet.xlsx");
        var dimensions = await _excelImporter.GetSheetDimensionsAsync(path);
        
        Assert.NotNull(dimensions);
        Assert.NotEmpty(dimensions);
        Assert.Equal(3, dimensions.Count);
        
        Assert.Equal(12, dimensions[0].Rows.Count);
        Assert.Equal(4, dimensions[0].Columns.Count);
        Assert.Equal(5, dimensions[2].Rows.Count);
        Assert.Equal(2, dimensions[2].Columns.Count);
    }

    [Fact]
    public async Task GetSheetDimensionsAsyncWithLeaveOpenTest()
    {
        var path = PathHelper.GetFile("xlsx/TestMultiSheet.xlsx");

        await using var stream = File.OpenRead(path);
        var dimensions = await _excelImporter.GetSheetDimensionsAsync(stream, leaveOpen: true);
        
        Assert.NotNull(dimensions);
        Assert.Equal(3, dimensions.Count);

        Assert.Equal(12, dimensions[0].Rows.Count);
        Assert.Equal(4, dimensions[0].Columns.Count);
        Assert.Equal(5, dimensions[2].Rows.Count);
        Assert.Equal(2, dimensions[2].Columns.Count);

        Assert.True(stream.CanRead);
    }

    [Fact]
    public async Task GetSheetInformationsAsyncTest()
    {
        var path = PathHelper.GetFile("xlsx/TestMultiSheet.xlsx");
        var sheetInfos = await _excelImporter.GetSheetInformationsAsync(path);
        
        Assert.NotNull(sheetInfos);
        Assert.NotEmpty(sheetInfos);
        Assert.Equal(3, sheetInfos.Count);

        for (int i = 0; i < 3; i++)
        {
            Assert.Equal($"Sheet{i + 1}", sheetInfos[i].Name);
            Assert.Equal((uint)i, sheetInfos[i].Index);
            Assert.Equal(SheetState.Visible, sheetInfos[i].State);
        }
    }

    [Fact]
    public async Task GetSheetInformationsAsyncWithLeaveOpenTest()
    {
        var path = PathHelper.GetFile("xlsx/TestMultiSheet.xlsx");

        await using var stream = File.OpenRead(path);
        var sheetInfos = await _excelImporter.GetSheetInformationsAsync(stream, leaveOpen: true);

        Assert.NotNull(sheetInfos);
        Assert.NotEmpty(sheetInfos);
        Assert.Equal(3, sheetInfos.Count);

        for (int i = 0; i < 3; i++)
        {
            Assert.Equal($"Sheet{i + 1}", sheetInfos[i].Name);
            Assert.Equal((uint)i, sheetInfos[i].Index);
            Assert.Equal(SheetState.Visible, sheetInfos[i].State);
        }

        Assert.True(stream.CanRead);
    }

    [Fact]
    public async Task GetColumnNamesAsyncTest()
    {
        var path = PathHelper.GetFile("xlsx/TestTypeMapping.xlsx");

        await using var stream = File.OpenRead(path);
        var columnNames = await _excelImporter.GetColumnNamesAsync(stream, hasHeaderRow: true);

        Assert.Equal(["ID", "Name", "BoD", "Age", "VIP", "Mail", "Points", "IgnoredProperty"], columnNames);
    }

    [Fact]
    public async Task GetColumnNamesAsyncWithoutHeaderTest()
    {
        var path = PathHelper.GetFile("xlsx/TestTypeMapping.xlsx");

        await using var stream = File.OpenRead(path);
        var columnNames = await _excelImporter.GetColumnNamesAsync(path);

        Assert.Equal(["A", "B", "C", "D", "E", "F", "G", "H"], columnNames);
    }

    [Fact]
    public async Task GetColumnsFromEmptyFileAsyncTest()
    {
        var path = PathHelper.GetFile("xlsx/TestEmpty.xlsx");
        var columns = await _excelImporter.GetColumnNamesAsync(path);
        Assert.Empty(columns);
    }
}
