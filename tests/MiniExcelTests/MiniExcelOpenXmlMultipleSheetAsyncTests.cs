using Xunit;
using System.Linq;
using System;
using System.IO;
using System.Threading.Tasks;

namespace MiniExcelLibs.Tests
{
    public partial class MiniExcelOpenXmlMultipleSheetAsyncTests
    {
        [Fact]
        public async Task SpecifySheetNameQueryTest()
        {
            var path = @"../../../../../samples/xlsx/TestMultiSheet.xlsx";
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
                await Assert.ThrowsAsync<InvalidOperationException>(async() => {
                    var q= await MiniExcel.QueryAsync(path, sheetName: "xxxx");
                    q.ToList();
                });
            }

            using (var stream = File.OpenRead(path))
            {
                {
                    var q = await stream.QueryAsync(sheetName: "Sheet3");
                    var rows = q.ToList();
                    Assert.Equal(5, rows.Count);
                    Assert.Equal(3d, rows[0]["A"]);
                    Assert.Equal(3d, rows[0]["B"]);
                }
                {
                    var q = await stream.QueryAsync(sheetName: "Sheet2");
                    var rows = q.ToList();
                    Assert.Equal(12, rows.Count);
                    Assert.Equal(1d, rows[0]["A"]);
                    Assert.Equal(1d, rows[0]["B"]);
                }
                {
                    var q = await stream.QueryAsync(sheetName: "Sheet1");
                    var rows = q.ToList();
                    Assert.Equal(12, rows.Count);
                    Assert.Equal(2d, rows[0]["A"]);
                    Assert.Equal(2d, rows[0]["B"]);
                }
                {
                    var q = await stream.QueryAsync(sheetName: "Sheet1");
                    var rows = q.ToList();
                    Assert.Equal(12, rows.Count);
                    Assert.Equal(2d, rows[0]["A"]);
                    Assert.Equal(2d, rows[0]["B"]);
                }
            }
        }

        [Fact]
        public async Task MultiSheetsQueryBasicTest()
        {
            var path = @"../../../../../samples/xlsx/TestMultiSheet.xlsx";
            using (var stream = File.OpenRead(path))
            {
                var sheet1 = await stream.QueryAsync(sheetName: "Sheet1");
                var sheet2 = await stream.QueryAsync(sheetName: "Sheet2");
                var sheet3 = await stream.QueryAsync(sheetName: "Sheet3");
            }
        }

        [Fact]
        public async Task MultiSheetsQueryTest()
        {
            var path = @"../../../../../samples/xlsx/TestMultiSheet.xlsx";
            {
                var sheetNames = MiniExcel.GetSheetNames(path).ToList();
                foreach (var sheetName in sheetNames)
                {
                    var rows = await MiniExcel.QueryAsync(path, sheetName: sheetName);
                }

                Assert.Equal(new[] { "Sheet2", "Sheet1", "Sheet3" }, sheetNames);
            }

            {
                using (var stream = File.OpenRead(path))
                {
                    var sheetNames = stream.GetSheetNames().ToList();
                    Assert.Equal(new[] { "Sheet2", "Sheet1", "Sheet3" }, sheetNames);
                    foreach (var sheetName in sheetNames)
                    {
                        var rows = await stream.QueryAsync(sheetName: sheetName);
                    }
                }
            }
        }
    }
}