using Xunit;

namespace MiniExcelLibs.Tests;

public class MiniExcelOpenXmlMultipleSheetAsyncTests
{
    [Fact]
    public async Task SpecifySheetNameQueryTest()
    {
        const string path = "../../../../../samples/xlsx/TestMultiSheet.xlsx";
        {
            var q = await MiniExcel.QueryAsync(path, sheetName: "Sheet3");
            var rows = q.ToList();
            Assert.Equal(5, rows.Count);
            Assert.Equal(3, rows[0].A);
            Assert.Equal(3, rows[0].B);
        }
        {
            var q = await MiniExcel.QueryAsync(path, sheetName: "Sheet2");
            var rows = q.ToList();
            Assert.Equal(12, rows.Count);
            Assert.Equal(1, rows[0].A);
            Assert.Equal(1, rows[0].B);
        }
        {
            var q = await MiniExcel.QueryAsync(path, sheetName: "Sheet1");
            var rows = q.ToList();
            Assert.Equal(12, rows.Count);
            Assert.Equal(2, rows[0].A);
            Assert.Equal(2, rows[0].B);
        }
        {
            await Assert.ThrowsAsync<InvalidOperationException>(async () =>
            {
                var rows = (await MiniExcel.QueryAsync(path, sheetName: "xxxx")).ToList();
            });
        }

        await using var stream = File.OpenRead(path);
        
        {
            var rows = (await stream.QueryAsync(sheetName: "Sheet3")).Cast<IDictionary<string, object>>().ToList();
            Assert.Equal(5, rows.Count);
            Assert.Equal(3d, rows[0]["A"]);
            Assert.Equal(3d, rows[0]["B"]);
        }
        {
            var rows = (await stream.QueryAsync(sheetName: "Sheet2")).Cast<IDictionary<string, object>>().ToList();
            Assert.Equal(12, rows.Count);
            Assert.Equal(1d, rows[0]["A"]);
            Assert.Equal(1d, rows[0]["B"]);
        }
        {
            var rows = (await stream.QueryAsync(sheetName: "Sheet1")).Cast<IDictionary<string, object>>().ToList();
            Assert.Equal(12, rows.Count);
            Assert.Equal(2d, rows[0]["A"]);
            Assert.Equal(2d, rows[0]["B"]);
        }
        {
            var rows = (await stream.QueryAsync(sheetName: "Sheet1")).Cast<IDictionary<string, object>>().ToList();
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
        _ = await stream.QueryAsync(sheetName: "Sheet1");
        _ = await stream.QueryAsync(sheetName: "Sheet2");
        _ = await stream.QueryAsync(sheetName: "Sheet3");
    }

    [Fact]
    public async Task MultiSheetsQueryTest()
    {
        const string path = "../../../../../samples/xlsx/TestMultiSheet.xlsx";
        {
            var sheetNames = MiniExcel.GetSheetNames(path).ToList();
            foreach (var sheetName in sheetNames)
            {
                var rows = await MiniExcel.QueryAsync(path, sheetName: sheetName);
            }
            Assert.Equal(new[] { "Sheet2", "Sheet1", "Sheet3" }, sheetNames);
        }

        {
            await using var stream = File.OpenRead(path);
            var sheetNames = stream.GetSheetNames().ToList();
            Assert.Equal(new[] { "Sheet2", "Sheet1", "Sheet3" }, sheetNames);
            foreach (var sheetName in sheetNames)
            {
                var rows = (await stream.QueryAsync(sheetName: sheetName)).ToList();
            }
        }
    }
}