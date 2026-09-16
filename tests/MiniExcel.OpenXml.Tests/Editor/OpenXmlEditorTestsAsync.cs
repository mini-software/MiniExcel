using System.Drawing;
using ClosedXML.Excel;
using MiniExcelLib.Tests.Common.Utils;

namespace MiniExcelLib.OpenXml.Tests.Editor;

public class OpenXmlEditorTestsAsync
{
    [Fact]
    public async Task SaveAppliesUpdatesInCellOrderAndPreservesExistingStyle()
    {
        using var path = AutoDeletingPath.Create();
        using (var workbook = new XLWorkbook())
        {
            var worksheet = workbook.AddWorksheet("Data");
            var firstCell = worksheet.Cell("A1");
            firstCell.Value = 12.34;
            firstCell.Style.NumberFormat.Format = "0.00";
            firstCell.Style.Fill.BackgroundColor = XLColor.Yellow;
            firstCell.Style.Border.LeftBorder = XLBorderStyleValues.Thin;
            firstCell.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            worksheet.Cell("X100").Value = "last";
            workbook.SaveAs(path.ToString());
        }

        await MiniExcelV2.Editors.GetOpenXmlEditor()
            .StartEditingPipeline(path.ToString())
            .UpdateCellStyle("X100", style => style.FontColor = Color.Blue, "Data")
            .UpdateCellStyle("A1", style => style.FontColor = Color.Red, "Data")
            .SaveChangesAsync();

        using var updatedWorkbook = new XLWorkbook(path.ToString());
        var updatedWorksheet = updatedWorkbook.Worksheet("Data");
        var firstCellStyle = updatedWorksheet.Cell("A1").Style;

        Assert.Equal(Color.Red.ToArgb(), firstCellStyle.Font.FontColor.Color.ToArgb());
        Assert.Equal(Color.Blue.ToArgb(), updatedWorksheet.Cell("X100").Style.Font.FontColor.Color.ToArgb());
        Assert.Equal("0.00", firstCellStyle.NumberFormat.Format);
        Assert.Equal(XLColor.Yellow, firstCellStyle.Fill.BackgroundColor);
        Assert.Equal(XLBorderStyleValues.Thin, firstCellStyle.Border.LeftBorder);
        Assert.Equal(XLAlignmentHorizontalValues.Center, firstCellStyle.Alignment.Horizontal);
    }

    [Fact]
    public async Task LastUpdateWinsForTheSameCell()
    {
        using var path = AutoDeletingPath.Create();
        CreateWorkbook(path.ToString());

        await MiniExcelV2.Editors.GetOpenXmlEditor()
            .StartEditingPipeline(path.ToString())
            .UpdateCellStyle("A1", style => style.FontColor = Color.Red)
            .UpdateCellStyle("A1", style => style.FontColor = Color.Blue)
            .SaveChangesAsync();

        using var workbook = new XLWorkbook(path.ToString());
        Assert.Equal(Color.Blue.ToArgb(), workbook.Worksheet(1).Cell("A1").Style.Font.FontColor.Color.ToArgb());
    }

    [Fact]
    public async Task SaveAsyncUpdatesASelectedWorksheetInAStream()
    {
        using var stream = new MemoryStream();
        using (var workbook = new XLWorkbook())
        {
            workbook.AddWorksheet("First").Cell("A1").Value = "first";
            workbook.AddWorksheet("Second").Cell("A1").Value = "second";
            workbook.SaveAs(stream);
        }

        await MiniExcelV2.Editors.GetOpenXmlEditor()
            .StartEditingPipeline(stream, leaveOpen: true)
            .UpdateCellStyle("A1", style => style.FontColor = Color.Blue, "Second")
            .SaveChangesAsync();

        stream.Position = 0;
        using var updatedWorkbook = new XLWorkbook(stream);
        Assert.NotEqual(Color.Blue.ToArgb(), updatedWorkbook.Worksheet("First").Cell("A1").Style.Font.FontColor.Color.ToArgb());
        Assert.Equal(Color.Blue.ToArgb(), updatedWorkbook.Worksheet("Second").Cell("A1").Style.Font.FontColor.Color.ToArgb());
    }

    [Fact]
    public async Task FailedSaveLeavesTheOriginalWorkbookUnchanged()
    {
        using var path = AutoDeletingPath.Create();
        CreateWorkbook(path.ToString());

        var editor = MiniExcelV2.Editors.GetOpenXmlEditor()
            .StartEditingPipeline(path.ToString())
            .UpdateCellStyle("A1", style => style.FontColor = Color.Red)
            .UpdateCellStyle("A2", style => style.FontColor = Color.Blue);

        await Assert.ThrowsAsync<InvalidDataException>(() => editor.SaveChangesAsync());

        using var workbook = new XLWorkbook(path.ToString());
        Assert.NotEqual(Color.Red.ToArgb(), workbook.Worksheet(1).Cell("A1").Style.Font.FontColor.Color.ToArgb());
    }

    private static void CreateWorkbook(string path)
    {
        using var workbook = new XLWorkbook();
        workbook.AddWorksheet("Data").Cell("A1").Value = "value";
        workbook.SaveAs(path);
    }
}
