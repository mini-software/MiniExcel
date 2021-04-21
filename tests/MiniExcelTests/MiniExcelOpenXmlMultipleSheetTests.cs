using Xunit;
using System.Linq;
using System;
using System.IO;

namespace MiniExcelLibs.Tests
{
    public partial class MiniExcelOpenXmlMultipleSheetTests
    {
        [Fact]
        public void SpecifySheetNameQueryTest()
        {
            var path = @"../../../../../samples/xlsx/TestMultiSheet.xlsx";
            {
                var rows = MiniExcel.Query(path, sheetName: "Sheet3").ToList();
                Assert.Equal(5, rows.Count);
                Assert.Equal(3, rows[0].A);
                Assert.Equal(3, rows[0].B);
            }
            {
                var rows = MiniExcel.Query(path, sheetName: "Sheet2").ToList();
                Assert.Equal(12, rows.Count);
                Assert.Equal(1, rows[0].A);
                Assert.Equal(1, rows[0].B);
            }
            {
                var rows = MiniExcel.Query(path, sheetName: "Sheet1").ToList();
                Assert.Equal(12, rows.Count);
                Assert.Equal(2, rows[0].A);
                Assert.Equal(2, rows[0].B);
            }
            {
                Assert.Throws<InvalidOperationException>(() => MiniExcel.Query(path, sheetName: "xxxx").ToList());
            }

            using (var stream = File.OpenRead(path))
            {
                {
                    var rows = stream.Query(sheetName: "Sheet3").ToList();
                    Assert.Equal(5, rows.Count);
                    Assert.Equal(3, rows[0].A);
                    Assert.Equal(3, rows[0].B);
                }
                {
                    var rows = stream.Query(sheetName: "Sheet2").ToList();
                    Assert.Equal(12, rows.Count);
                    Assert.Equal(1, rows[0].A);
                    Assert.Equal(1, rows[0].B);
                }
                {
                    var rows = stream.Query(sheetName: "Sheet1").ToList();
                    Assert.Equal(12, rows.Count);
                    Assert.Equal(2, rows[0].A);
                    Assert.Equal(2, rows[0].B);
                }
                {
                    var rows = stream.Query(sheetName: "Sheet1").ToList();
                    Assert.Equal(12, rows.Count);
                    Assert.Equal(2, rows[0].A);
                    Assert.Equal(2, rows[0].B);
                }
            }
        }

        [Fact]
        public void MultiSheetsQueryBasicTest()
        {
            var path = @"../../../../../samples/xlsx/TestMultiSheet.xlsx";
            using (var stream = File.OpenRead(path))
            {
                var sheet1 = stream.Query(sheetName: "Sheet1");
                var sheet2 = stream.Query(sheetName: "Sheet2");
                var sheet3 = stream.Query(sheetName: "Sheet3");
            }
        }

        [Fact]
        public void MultiSheetsQueryTest()
        {
            var path = @"../../../../../samples/xlsx/TestMultiSheet.xlsx";
            {
                var sheetNames = MiniExcel.GetSheetNames(path).ToList();
                foreach (var sheetName in sheetNames)
                {
                    var rows = MiniExcel.Query(path, sheetName: sheetName);
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
                        var rows = stream.Query(sheetName: sheetName);
                    }
                }
            }
        }
    }
}