using MiniExcelLib.Tests.Common.Utils;

namespace MiniExcelLib.OpenXml.Tests.MultipleSheets;

public class MiniExcelOpenXmlMultipleSheetAsyncTests
{
    private readonly OpenXmlImporter _excelImporter =  MiniExcel.Importers.GetOpenXmlImporter();
    
    [Fact]
    public async Task SpecifySheetNameQueryTest()
    {
        var path = PathHelper.GetFile("xlsx/TestMultiSheet.xlsx");
        
        var rows1 = await _excelImporter.QueryAsync(path, sheetName: "Sheet3").ToListAsync();
        Assert.Equal(5, rows1.Count);
        Assert.Equal(3, rows1[0].A);
        Assert.Equal(3, rows1[0].B);

        var rows2 = await _excelImporter.QueryAsync(path, sheetName: "Sheet2").ToListAsync();
        Assert.Equal(12, rows2.Count);
        Assert.Equal(1, rows2[0].A);
        Assert.Equal(1, rows2[0].B);

        var rows3 = await _excelImporter.QueryAsync(path, sheetName: "Sheet1").ToListAsync();
        Assert.Equal(12, rows3.Count);
        Assert.Equal(2, rows3[0].A);
        Assert.Equal(2, rows3[0].B);
        await Assert.ThrowsAsync<InvalidOperationException>(async () => await _excelImporter.QueryAsync(path, sheetName: "xxxx").ToListAsync());

        await using var stream = File.OpenRead(path);
        
        var rows4 = await _excelImporter.QueryAsync(stream, sheetName: "Sheet3").Cast<IDictionary<string, object>>().ToListAsync();
        Assert.Equal(5, rows4.Count);
        Assert.Equal(3d, rows4[0]["A"]);
        Assert.Equal(3d, rows4[0]["B"]);

        var rows5 = await _excelImporter.QueryAsync(stream, sheetName: "Sheet2").Cast<IDictionary<string, object>>().ToListAsync();
        Assert.Equal(12, rows5.Count);
        Assert.Equal(1d, rows5[0]["A"]);
        Assert.Equal(1d, rows5[0]["B"]);
        
        var rows6 = await _excelImporter.QueryAsync(stream, sheetName: "Sheet1").Cast<IDictionary<string, object>>().ToListAsync();
        Assert.Equal(12, rows6.Count);
        Assert.Equal(2d, rows6[0]["A"]);
        Assert.Equal(2d, rows6[0]["B"]);

        var rows7 = await _excelImporter.QueryAsync(stream, sheetName: "Sheet1").Cast<IDictionary<string, object>>().ToListAsync();
        Assert.Equal(12, rows7.Count);
        Assert.Equal(2d, rows7[0]["A"]);
        Assert.Equal(2d, rows7[0]["B"]);
    }

    [Fact]
    public async Task MultiSheetsQueryBasicTest()
    {
        var path = PathHelper.GetFile("xlsx/TestMultiSheet.xlsx");
        await using var stream = File.OpenRead(path);
        _ = await _excelImporter.QueryAsync(stream, sheetName: "Sheet1").ToListAsync();
        _ = await _excelImporter.QueryAsync(stream, sheetName: "Sheet2").ToListAsync();
        _ = await _excelImporter.QueryAsync(stream, sheetName: "Sheet3").ToListAsync();
    }

    [Fact]
    public async Task MultiSheetsQueryTest()
    {
        var path = PathHelper.GetFile("xlsx/TestMultiSheet.xlsx");
        
        var sheetNames1 = await _excelImporter.GetSheetNamesAsync(path);
        foreach (var sheetName in sheetNames1)
        {
            _ = await _excelImporter.QueryAsync(path, sheetName: sheetName).ToListAsync();
        }
            
        Assert.Equal(["Sheet1", "Sheet2", "Sheet3"], sheetNames1);

        
        await using var stream = File.OpenRead(path);
        var sheetNames2 = await  _excelImporter.GetSheetNamesAsync(stream);
            
        Assert.Equal(["Sheet1", "Sheet2", "Sheet3"], sheetNames2);
        foreach (var sheetName in sheetNames2)
        {
            _ = await _excelImporter.QueryAsync(stream, sheetName: sheetName).ToListAsync();
        }
    }
}
