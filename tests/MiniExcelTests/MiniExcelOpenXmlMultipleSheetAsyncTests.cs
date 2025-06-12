using Xunit;

namespace MiniExcelLibs.Tests;

public class MiniExcelOpenXmlMultipleSheetAsyncTests
{
    [Fact]
    public async Task SpecifySheetNameQueryTest()
    {
        const string path = "../../../../../samples/xlsx/TestMultiSheet.xlsx";
        {
            var q = MiniExcel.QueryAsync(path, sheetName: "Sheet3").ToBlockingEnumerable();
            var rows = q.ToList();
            Assert.Equal(5, rows.Count);
            Assert.Equal(3, rows[0].A);
            Assert.Equal(3, rows[0].B);
        }
        {
            var q = MiniExcel.QueryAsync(path, sheetName: "Sheet2").ToBlockingEnumerable();
            var rows = q.ToList();
            Assert.Equal(12, rows.Count);
            Assert.Equal(1, rows[0].A);
            Assert.Equal(1, rows[0].B);
        }
        {
            var q = MiniExcel.QueryAsync(path, sheetName: "Sheet1").ToBlockingEnumerable();
            var rows = q.ToList();
            Assert.Equal(12, rows.Count);
            Assert.Equal(2, rows[0].A);
            Assert.Equal(2, rows[0].B);
        }
        {
            await Assert.ThrowsAsync<InvalidOperationException>(async () =>
            {
                var rows = MiniExcel.QueryAsync(path, sheetName: "xxxx").ToBlockingEnumerable().ToList();
            });
        }

        await using var stream = File.OpenRead(path);
        
        {
            var rows = (stream.QueryAsync(sheetName: "Sheet3").ToBlockingEnumerable()).Cast<IDictionary<string, object>>().ToList();
            Assert.Equal(5, rows.Count);
            Assert.Equal(3d, rows[0]["A"]);
            Assert.Equal(3d, rows[0]["B"]);
        }
        {
            var rows = (stream.QueryAsync(sheetName: "Sheet2").ToBlockingEnumerable()).Cast<IDictionary<string, object>>().ToList();
            Assert.Equal(12, rows.Count);
            Assert.Equal(1d, rows[0]["A"]);
            Assert.Equal(1d, rows[0]["B"]);
        }
        {
            var rows = (stream.QueryAsync(sheetName: "Sheet1").ToBlockingEnumerable()).Cast<IDictionary<string, object>>().ToList();
            Assert.Equal(12, rows.Count);
            Assert.Equal(2d, rows[0]["A"]);
            Assert.Equal(2d, rows[0]["B"]);
        }
        {
            var rows = (stream.QueryAsync(sheetName: "Sheet1").ToBlockingEnumerable()).Cast<IDictionary<string, object>>().ToList();
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
        _ = stream.QueryAsync(sheetName: "Sheet1").ToBlockingEnumerable();
        _ = stream.QueryAsync(sheetName: "Sheet2").ToBlockingEnumerable();
        _ = stream.QueryAsync(sheetName: "Sheet3").ToBlockingEnumerable();
    }

    [Fact]
    public async Task MultiSheetsQueryTest()
    {
        const string path = "../../../../../samples/xlsx/TestMultiSheet.xlsx";
        {
            var sheetNames = MiniExcel.GetSheetNames(path).ToList();
            foreach (var sheetName in sheetNames)
            {
                var rows = MiniExcel.QueryAsync(path, sheetName: sheetName).ToBlockingEnumerable();
            }
            Assert.Equal(new[] { "Sheet1", "Sheet2", "Sheet3" }, sheetNames);
        }

        {
            await using var stream = File.OpenRead(path);
            var sheetNames = stream.GetSheetNames().ToList();
            Assert.Equal(new[] { "Sheet1", "Sheet2", "Sheet3" }, sheetNames);
            foreach (var sheetName in sheetNames)
            {
                var rows = (stream.QueryAsync(sheetName: sheetName).ToBlockingEnumerable()).ToList();
            }
        }
    }
}