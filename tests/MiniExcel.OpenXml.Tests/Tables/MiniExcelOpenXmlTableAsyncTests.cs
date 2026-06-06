using MiniExcelLib.Tests.Common.Utils;

namespace MiniExcelLib.OpenXml.Tests.Tables;

public class MiniExcelOpenXmlTableAsyncTests
{
    private readonly OpenXmlImporter _excelImporter =  MiniExcel.Importers.GetOpenXmlImporter();

    /// <summary>
    /// Tests querying a named table from a file path with dynamic results.
    /// </summary>
    [Fact]
    public async Task QueryTableAsync_FromFilePath_ReturnsDynamicRows()
    {
        // Arrange
        var path = PathHelper.GetFile("xlsx/TestQueryTable.xlsx");
        
        // Act
        var rows = await _excelImporter.QueryTableAsync(path).ToListAsync();
        
        // Assert
        Assert.Equal(3, rows.Count);
        Assert.Equal("aaa", rows[0].Col1);
        Assert.Equal(123D, rows[0].Col2);
        Assert.Equal(new DateTime(2026, 5, 17), rows[0].Col3);
    }

    /// <summary>
    /// Tests querying a named table from a stream with dynamic results.
    /// </summary>
    [Fact]
    public async Task QueryTableAsync_FromStream_ReturnsDynamicRows()
    {
        // Arrange
        var path = PathHelper.GetFile("xlsx/TestQueryTable.xlsx");
        await using var stream = File.OpenRead(path);

        // Act
        var rows = await _excelImporter.QueryTableAsync(stream).ToListAsync();

        // Assert
        Assert.Equal(3, rows.Count);
        Assert.Equal("bbb", rows[1].Col1);
        Assert.Equal(456D, rows[1].Col2);
        Assert.Equal(new DateTime(2026, 5, 18), rows[1].Col3);
    }

    /// <summary>
    /// Tests querying a named table from a file path with strongly-typed results.
    /// </summary>
    [Fact]
    public async Task QueryTableAsync_Generic_FromFilePath_ReturnsTypedRows()
    {
        // Arrange
        var path = PathHelper.GetFile("xlsx/TestQueryTable.xlsx");
        
        // Act
        var rows = await _excelImporter.QueryTableAsync<QueryTableTestModel>(path).ToListAsync();
        
        // Assert
        Assert.Equal(3, rows.Count);
        Assert.Equal("aaa", rows[0].Col1);
        Assert.Equal(123D, rows[0].Col2);
        Assert.Equal(new DateTime(2026, 5, 17), rows[0].Col3);
    }

    /// <summary>
    /// Tests querying a named table from a stream with strongly-typed results.
    /// </summary>
    [Fact]
    public async Task QueryTableAsync_Generic_FromStream_ReturnsTypedRows()
    {
        // Arrange
        var path = PathHelper.GetFile("xlsx/TestQueryTable.xlsx");
        await using var stream = File.OpenRead(path);
        
        // Act
        var rows = await _excelImporter.QueryTableAsync<QueryTableTestModel>(stream).ToListAsync();
        
        // Assert
        Assert.Equal(3, rows.Count);
        Assert.Equal("ccc", rows[2].Col1);
        Assert.Equal(789D, rows[2].Col2);
        Assert.Equal(new DateTime(2026, 5, 19), rows[2].Col3);
    }

    /// <summary>
    /// Tests querying multiple tables from the same sheet.
    /// </summary>
    [Fact]
    public async Task QueryTableAsync_MultipleTablesInSheet_ReturnsCorrectTableData()
    {
        // Arrange
        var path = PathHelper.GetFile("xlsx/TestQueryTable.xlsx");
        
        // Act
        var table1 = await _excelImporter.QueryTableAsync(path).ToListAsync();
        var table2 = await _excelImporter.QueryTableAsync(path, "Sheet1", "Table2").ToListAsync();
        
        // Assert
        Assert.NotEmpty(table1);
        Assert.NotEmpty(table2);
                
        // Assert
        Assert.Equal(3, table1.Count);
        Assert.Equal("aaa", table1[0].Col1);
        Assert.Equal(123D, table1[0].Col2);
        Assert.Equal(new DateTime(2026, 5, 17), table1[0].Col3);

        Assert.Equal(2, table2.Count);
        Assert.Equal("test", table2[0].Prop1);
        Assert.Equal(11D, table2[0].Prop2);
        Assert.Equal("aaa", table2[0].Prop3);
        Assert.Equal(new TimeSpan(10, 30, 0), table2[0].Prop4.TimeOfDay);
    }

    /// <summary>
    /// Tests QueryTableAsync with custom sheet and table names.
    /// </summary>
    [Fact]
    public async Task QueryTableAsync_WithCustomSheetAndTableNames_ReturnsCorrectData()
    {
        // Arrange
        var path = PathHelper.GetFile("xlsx/TestQueryTable.xlsx");
        
        // Act
        var rows = await _excelImporter.QueryTableAsync(path, "CustomSheet", "CustomTable").ToListAsync();
        
        // Assert
        Assert.NotEmpty(rows);
    }
}
