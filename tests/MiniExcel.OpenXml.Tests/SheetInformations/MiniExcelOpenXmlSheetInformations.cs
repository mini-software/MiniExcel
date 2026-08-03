using MiniExcelLib.OpenXml.Models;
using MiniExcelLib.OpenXml.Tests.Utils;
using MiniExcelLib.Tests.Common.Utils;
using System.Reflection;

namespace MiniExcelLib.OpenXml.Tests.SheetInformations;

public class MiniExcelOpenXmlSheetInformations
{
    private readonly OpenXmlImporter _excelImporter = MiniExcel.Importers.GetOpenXmlImporter();

    [Fact]
    public void GetSheetDimensionsTest()
    {
        var path = PathHelper.GetFile("xlsx/TestMultiSheet.xlsx");
        var dimensions = _excelImporter.GetSheetDimensions(path);

        Assert.NotNull(dimensions);
        Assert.NotEmpty(dimensions);
        Assert.Equal(3, dimensions.Count);

        Assert.Equal(12, dimensions[0].Rows.Count);
        Assert.Equal(4, dimensions[0].Columns.Count);
        Assert.Equal(5, dimensions[2].Rows.Count);
        Assert.Equal(2, dimensions[2].Columns.Count);
    }

    [Fact]
    public void GetSheetDimensionsWithLeaveOpenTest()
    {
        var path = PathHelper.GetFile("xlsx/TestMultiSheet.xlsx");

        using var stream = File.OpenRead(path);
        var dimensions = _excelImporter.GetSheetDimensions(stream, leaveOpen: true);

        Assert.NotNull(dimensions);
        Assert.Equal(3, dimensions.Count);

        Assert.Equal(12, dimensions[0].Rows.Count);
        Assert.Equal(4, dimensions[0].Columns.Count);
        Assert.Equal(5, dimensions[2].Rows.Count);
        Assert.Equal(2, dimensions[2].Columns.Count);

        Assert.True(stream.CanRead);
    }

    [Fact]
    public void GetSheetInformationsTest()
    {
        var path = PathHelper.GetFile("xlsx/TestMultiSheet.xlsx");
        var sheetInfos = _excelImporter.GetSheetInformations(path);

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
    public void GetSheetInformationsWithLeaveOpenTest()
    {
        var path = PathHelper.GetFile("xlsx/TestMultiSheet.xlsx");

        using var stream = File.OpenRead(path);
        var sheetInfos = _excelImporter.GetSheetInformations(stream, leaveOpen: true);

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
    public void GetColumnNamesTest()
    {
        var path = PathHelper.GetFile("xlsx/TestTypeMapping.xlsx");

        using var stream = File.OpenRead(path);
        var columnNames = _excelImporter.GetColumnNames(stream, hasHeaderRow: true);

        Assert.Equal(["ID", "Name", "BoD", "Age", "VIP", "Mail", "Points", "IgnoredProperty"], columnNames);
    }

    [Fact]
    public void GetColumnNamesWithoutHeaderTest()
    {
        var path = PathHelper.GetFile("xlsx/TestTypeMapping.xlsx");

        using var stream = File.OpenRead(path);
        var columnNames = _excelImporter.GetColumnNames(path);

        Assert.Equal(["A", "B", "C", "D", "E", "F", "G", "H"], columnNames);
    }

    [Fact]
    public void GetColumnsFromEmptyFileTest()
    {
        var path = PathHelper.GetFile("xlsx/TestEmpty.xlsx");
        var columns = _excelImporter.GetColumnNames(path);
        Assert.Empty(columns);
    }
}
