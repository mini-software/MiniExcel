namespace MiniExcelLib.Tests;

public class MiniExcelOpenXmlMultipleSheetAsyncTests
{
    private readonly OpenXmlImporter _importer =  MiniExcel.GetImporterProvider().GetExcelImporter();
    
    [Fact]
    public async Task SpecifySheetNameQueryTest()
    {
        const string path = "../../../../../samples/xlsx/TestMultiSheet.xlsx";
        {
            var q =  _importer.QueryExcelAsync(path, sheetName: "Sheet3").ToBlockingEnumerable();
            var rows = q.ToList();
            Assert.Equal(5, rows.Count);
            Assert.Equal(3, rows[0].A);
            Assert.Equal(3, rows[0].B);
        }
        {
            var q =  _importer.QueryExcelAsync(path, sheetName: "Sheet2").ToBlockingEnumerable();
            var rows = q.ToList();
            Assert.Equal(12, rows.Count);
            Assert.Equal(1, rows[0].A);
            Assert.Equal(1, rows[0].B);
        }
        {
            var q =  _importer.QueryExcelAsync(path, sheetName: "Sheet1").ToBlockingEnumerable();
            var rows = q.ToList();
            Assert.Equal(12, rows.Count);
            Assert.Equal(2, rows[0].A);
            Assert.Equal(2, rows[0].B);
        }
        {
            await Assert.ThrowsAsync<InvalidOperationException>(() =>
            {
                _ =  _importer.QueryExcelAsync(path, sheetName: "xxxx").ToBlockingEnumerable().ToList();
                return Task.CompletedTask;
            });
        }

        await using var stream = File.OpenRead(path);
        
        {
            var rows =  _importer.QueryExcelAsync(stream, sheetName: "Sheet3").ToBlockingEnumerable().Cast<IDictionary<string, object>>().ToList();
            Assert.Equal(5, rows.Count);
            Assert.Equal(3d, rows[0]["A"]);
            Assert.Equal(3d, rows[0]["B"]);
        }
        {
            var rows =  _importer.QueryExcelAsync(stream, sheetName: "Sheet2").ToBlockingEnumerable().Cast<IDictionary<string, object>>().ToList();
            Assert.Equal(12, rows.Count);
            Assert.Equal(1d, rows[0]["A"]);
            Assert.Equal(1d, rows[0]["B"]);
        }
        {
            var rows =  _importer.QueryExcelAsync(stream, sheetName: "Sheet1").ToBlockingEnumerable().Cast<IDictionary<string, object>>().ToList();
            Assert.Equal(12, rows.Count);
            Assert.Equal(2d, rows[0]["A"]);
            Assert.Equal(2d, rows[0]["B"]);
        }
        {
            var rows =  _importer.QueryExcelAsync(stream, sheetName: "Sheet1").ToBlockingEnumerable().Cast<IDictionary<string, object>>().ToList();
            Assert.Equal(12, rows.Count);
            Assert.Equal(2d, rows[0]["A"]);
            Assert.Equal(2d, rows[0]["B"]);
        }
    }

    [Fact]
    public async Task MultiSheetsQueryBasicTest()
    {
        const string path = "../../../../../samples/xlsx/TestMultiSheet.xlsx";
        await using var stream = File.OpenRead(path);
        _ =  _importer.QueryExcelAsync(stream, sheetName: "Sheet1").ToBlockingEnumerable();
        _ =  _importer.QueryExcelAsync(stream, sheetName: "Sheet2").ToBlockingEnumerable();
        _ =  _importer.QueryExcelAsync(stream, sheetName: "Sheet3").ToBlockingEnumerable();
    }

    [Fact]
    public async Task MultiSheetsQueryTest()
    {
        const string path = "../../../../../samples/xlsx/TestMultiSheet.xlsx";
        {
            var sheetNames = (await  _importer.GetSheetNamesAsync(path)).ToList();
            foreach (var sheetName in sheetNames)
            {
                _ =  _importer.QueryExcelAsync(path, sheetName: sheetName).ToBlockingEnumerable();
            }
            Assert.Equal(new[] { "Sheet1", "Sheet2", "Sheet3" }, sheetNames);
        }

        {
            await using var stream = File.OpenRead(path);
            var sheetNames = (await  _importer.GetSheetNamesAsync(stream)).ToList();
            Assert.Equal(new[] { "Sheet1", "Sheet2", "Sheet3" }, sheetNames);
            foreach (var sheetName in sheetNames)
            {
                _ =  _importer.QueryExcelAsync(stream, sheetName: sheetName).ToBlockingEnumerable().ToList();
            }
        }
    }
}