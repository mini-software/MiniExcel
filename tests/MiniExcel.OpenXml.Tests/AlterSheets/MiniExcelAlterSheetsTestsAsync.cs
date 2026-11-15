using ClosedXML.Excel;
using MiniExcelLib.OpenXml.Models;
using MiniExcelLib.Tests.Common.Utils;
using static MiniExcelLib.OpenXml.Tests.Utils.SheetHelper;

namespace MiniExcelLib.OpenXml.Tests.AlterSheets;

public class MiniExcelAlterSheetsTestAsync
{
    private readonly OpenXmlEditor _excelEditor = MiniExcelV2.Editors.GetOpenXmlEditor();
    
    [Fact]
    public async Task AlterSheetAsync_WhenNewNameProvided_RenamesWorksheet()
    {
        // Arrange
        const string originalName = "Sheet1";
        const string newName = "RenamedSheet";
        await using var stream = CreateTestWorkbookStream();

        // Act
        await _excelEditor.AlterSheetInfoAsync(stream, originalName, newSheetName: newName);

        // Assert
        stream.Position = 0; // Reset to read the saved results
        using var workbook = new XLWorkbook(stream);
        
        Assert.False(workbook.Worksheets.Contains(originalName));
        Assert.True(workbook.Worksheets.Contains(newName));
    }

    [Fact]
    public async Task AlterSheetAsync_WhenNewIndexProvided_MovesWorksheet()
    {
        // Arrange
        const string targetSheet = "Sheet1";
        const int newIndex = 2; 
        await using var stream = CreateTestWorkbookStream();

        // Act
        await _excelEditor.AlterSheetInfoAsync(stream, targetSheet, newSheetIndex: newIndex);

        // Assert
        stream.Position = 0;
        using var workbook = new XLWorkbook(stream);
        
        // Assert that the sheet at the new index is indeed our target sheet
        Assert.Equal(targetSheet, workbook.Worksheet(newIndex + 1).Name);
    }

    [Fact]
    public async Task AlterSheetAsync_WhenNewStateProvided_ChangesVisibility()
    {
        // Arrange
        const string targetSheet = "Sheet2";
        await using var stream = CreateTestWorkbookStream();
        
        // Act
        await _excelEditor.AlterSheetInfoAsync(stream, targetSheet, newSheetState: SheetState.Hidden);

        // Assert
        stream.Position = 0;
        using var workbook = new XLWorkbook(stream);
        
        var sheet = workbook.Worksheet(targetSheet);
        Assert.Equal(XLWorksheetVisibility.Hidden, sheet.Visibility);
    }

    [Fact]
    public async Task AlterSheetAsync_WhenAllPropertiesProvided_UpdatesAllSuccessfully()
    {
        // Arrange
        const string originalName = "Sheet3";
        const string newName = "SecretData";
        const int newIndex = 0;
        const SheetState newState = SheetState.VeryHidden;
        await using var stream = CreateTestWorkbookStream();

        // Act
        await _excelEditor.AlterSheetInfoAsync(
            stream, 
            originalName, 
            newSheetName: newName, 
            newSheetIndex: newIndex, 
            newSheetState: newState);

        // Assert
        stream.Position = 0;
        using var workbook = new XLWorkbook(stream);
        
        // 1. Check Name
        Assert.False(workbook.Worksheets.Contains(originalName));
        var modifiedSheet = workbook.Worksheet(newName);
        Assert.NotNull(modifiedSheet);

        // 2. Check Index (Should now be the first sheet)
        Assert.Equal(newName, workbook.Worksheet(newIndex + 1).Name);

        // 3. Check State
        Assert.Equal(XLWorksheetVisibility.VeryHidden, modifiedSheet.Visibility);
    }

    [Fact]
    public async Task AlterSheetAsync_WhenNoOptionalParametersProvided_LeavesSheetUnchanged()
    {
        // Arrange
        const string targetSheet = "Sheet1";
        using var path = AutoDeletingPath.Create();
        using (var stream = CreateTestWorkbookStream())
        {
            await using var fs = File.OpenWrite(path.FilePath);
            stream.Position = 0;
            await stream.CopyToAsync(fs);
        }

        // Act
        await _excelEditor.AlterSheetInfoAsync(path.FilePath, targetSheet);

        // Assert
        using var workbook = new XLWorkbook(path.FilePath);
        var sheet = workbook.Worksheet(targetSheet);

        // Ensure defaults remain intact
        Assert.NotNull(sheet);
        Assert.Equal("Sheet1", workbook.Worksheet(1).Name);
        Assert.Equal(XLWorksheetVisibility.Visible, sheet.Visibility);
    }
}