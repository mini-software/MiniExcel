using MiniExcelLib.OpenXml.Models;
using static MiniExcelLib.OpenXml.Tests.Utils.SheetHelper;

namespace MiniExcelLib.OpenXml.Tests.AlterSheets;

public class MiniExcelAlterSheetTests
{
    private readonly OpenXmlExporter _excelExporter = MiniExcel.Exporters.GetOpenXmlExporter();
    
    [Fact]
    public void AlterSheet_WhenNewNameProvided_RenamesWorksheet()
    {
        // Arrange
        const string originalName = "Sheet1";
        const string newName = "RenamedSheet";
        using var stream = CreateTestWorkbookStream();

        // Act
        _excelExporter.AlterSheet(stream, originalName, newSheetName: newName);

        // Assert
        stream.Position = 0;
        using var package = new ExcelPackage(stream);
        
        Assert.Null(package.Workbook.Worksheets[originalName]);
        Assert.NotNull(package.Workbook.Worksheets[newName]);
    }

    [Fact]
    public void AlterSheet_WhenNewIndexProvided_MovesWorksheet()
    {
        // Arrange
        const string targetSheet = "Sheet1";
        const int newIndex = 2; 
        using var stream = CreateTestWorkbookStream();

        // Act
        _excelExporter.AlterSheet(stream, targetSheet, newSheetIndex: newIndex);

        // Assert
        stream.Position = 0;
        using var package = new ExcelPackage(stream);
        
        // Assert that the sheet at the new index is indeed our target sheet
        Assert.Equal(targetSheet, package.Workbook.Worksheets[newIndex].Name);
    }

    [Fact]
    public void AlterSheet_WhenNewStateProvided_ChangesVisibility()
    {
        // Arrange
        const string targetSheet = "Sheet2";
        using var stream = CreateTestWorkbookStream();
        
        // Act
        _excelExporter.AlterSheet(stream, targetSheet, newSheetState: SheetState.Hidden);

        // Assert
        stream.Position = 0;
        using var package = new ExcelPackage(stream);
        
        var sheet = package.Workbook.Worksheets[targetSheet];
        Assert.Equal(eWorkSheetHidden.Hidden, sheet.Hidden);
    }

    [Fact]
    public void AlterSheet_WhenAllPropertiesProvided_UpdatesAllSuccessfully()
    {
        // Arrange
        const string originalName = "Sheet3";
        const string newName = "SecretData";
        const int newIndex = 0;
        const SheetState newState = SheetState.VeryHidden;
        using var stream = CreateTestWorkbookStream();

        // Act
        _excelExporter.AlterSheet(
            stream, 
            originalName, 
            newSheetName: newName, 
            newSheetIndex: newIndex, 
            newSheetState: newState);

        // Assert
        stream.Position = 0;
        using var package = new ExcelPackage(stream);
        
        // 1. Check Name
        Assert.Null(package.Workbook.Worksheets[originalName]);
        var modifiedSheet = package.Workbook.Worksheets[newName];
        Assert.NotNull(modifiedSheet);

        // 2. Check Index (Should now be the first sheet)
        Assert.Equal(newName, package.Workbook.Worksheets[newIndex].Name);

        // 3. Check State
        Assert.Equal(eWorkSheetHidden.VeryHidden, modifiedSheet.Hidden);
    }

    [Fact]
    public void AlterSheet_WhenNoOptionalParametersProvided_LeavesSheetUnchanged()
    {
        // Arrange
        const string targetSheet = "Sheet1";
        using var stream = CreateTestWorkbookStream();

        // Act
        _excelExporter.AlterSheet(stream, targetSheet);

        // Assert
        stream.Position = 0;
        using var package = new ExcelPackage(stream);
        var sheet = package.Workbook.Worksheets[targetSheet];

        // Ensure defaults remain intact
        Assert.NotNull(sheet);
        Assert.Equal("Sheet1", package.Workbook.Worksheets[0].Name); 
        Assert.Equal(eWorkSheetHidden.Visible, sheet.Hidden);
    }
}