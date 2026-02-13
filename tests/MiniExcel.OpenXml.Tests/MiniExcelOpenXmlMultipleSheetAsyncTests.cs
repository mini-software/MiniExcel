using MiniExcelLib.Tests.Common.Utils;

namespace MiniExcelLib.OpenXml.Tests;

public class MiniExcelOpenXmlMultipleSheetAsyncTests
{
    private readonly OpenXmlImporter _excelImporter =  MiniExcel.Importers.GetOpenXmlImporter();
    
    [Fact]
    public async Task SpecifySheetNameQueryTest()
    {
        var path = PathHelper.GetFile("xlsx/TestMultiSheet.xlsx");
        {
            var q =  _excelImporter.QueryAsync(path, sheetName: "Sheet3").ToBlockingEnumerable();
            var rows = q.ToList();
            Assert.Equal(5, rows.Count);
            Assert.Equal(3, rows[0].A);
            Assert.Equal(3, rows[0].B);
        }
        {
            var q =  _excelImporter.QueryAsync(path, sheetName: "Sheet2").ToBlockingEnumerable();
            var rows = q.ToList();
            Assert.Equal(12, rows.Count);
            Assert.Equal(1, rows[0].A);
            Assert.Equal(1, rows[0].B);
        }
        {
            var q =  _excelImporter.QueryAsync(path, sheetName: "Sheet1").ToBlockingEnumerable();
            var rows = q.ToList();
            Assert.Equal(12, rows.Count);
            Assert.Equal(2, rows[0].A);
            Assert.Equal(2, rows[0].B);
        }
        {
            await Assert.ThrowsAsync<InvalidOperationException>(() =>
            {
                _ =  _excelImporter.QueryAsync(path, sheetName: "xxxx").ToBlockingEnumerable().ToList();
                return Task.CompletedTask;
            });
        }

        await using var stream = File.OpenRead(path);
        
        {
            var rows =  _excelImporter.QueryAsync(stream, sheetName: "Sheet3").ToBlockingEnumerable().Cast<IDictionary<string, object>>().ToList();
            Assert.Equal(5, rows.Count);
            Assert.Equal(3d, rows[0]["A"]);
            Assert.Equal(3d, rows[0]["B"]);
        }
        {
            var rows =  _excelImporter.QueryAsync(stream, sheetName: "Sheet2").ToBlockingEnumerable().Cast<IDictionary<string, object>>().ToList();
            Assert.Equal(12, rows.Count);
            Assert.Equal(1d, rows[0]["A"]);
            Assert.Equal(1d, rows[0]["B"]);
        }
        {
            var rows =  _excelImporter.QueryAsync(stream, sheetName: "Sheet1").ToBlockingEnumerable().Cast<IDictionary<string, object>>().ToList();
            Assert.Equal(12, rows.Count);
            Assert.Equal(2d, rows[0]["A"]);
            Assert.Equal(2d, rows[0]["B"]);
        }
        {
            var rows =  _excelImporter.QueryAsync(stream, sheetName: "Sheet1").ToBlockingEnumerable().Cast<IDictionary<string, object>>().ToList();
            Assert.Equal(12, rows.Count);
            Assert.Equal(2d, rows[0]["A"]);
            Assert.Equal(2d, rows[0]["B"]);
        }
    }

    [Fact]
    public async Task MultiSheetsQueryBasicTest()
    {
        var path = PathHelper.GetFile("xlsx/TestMultiSheet.xlsx");
        await using var stream = File.OpenRead(path);
        _ =  _excelImporter.QueryAsync(stream, sheetName: "Sheet1").ToBlockingEnumerable();
        _ =  _excelImporter.QueryAsync(stream, sheetName: "Sheet2").ToBlockingEnumerable();
        _ =  _excelImporter.QueryAsync(stream, sheetName: "Sheet3").ToBlockingEnumerable();
    }

    [Fact]
    public async Task MultiSheetsQueryTest()
    {
        var path = PathHelper.GetFile("xlsx/TestMultiSheet.xlsx");
        {
            var sheetNames = await  _excelImporter.GetSheetNamesAsync(path);
            foreach (var sheetName in sheetNames)
            {
                _ =  _excelImporter.QueryAsync(path, sheetName: sheetName).ToBlockingEnumerable();
            }
            
            Assert.Equal(new[] { "Sheet1", "Sheet2", "Sheet3" }, sheetNames);
        }

        {
            await using var stream = File.OpenRead(path);
            var sheetNames = await  _excelImporter.GetSheetNamesAsync(stream);
            
            Assert.Equal(new[] { "Sheet1", "Sheet2", "Sheet3" }, sheetNames);
            foreach (var sheetName in sheetNames)
            {
                _ = _excelImporter.QueryAsync(stream, sheetName: sheetName).ToBlockingEnumerable().ToList();
            }
        }
    }
}