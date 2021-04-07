using Xunit;
using System.Linq;
using System;

namespace MiniExcelLibs.Tests
{
    public partial class MiniExcelOpenXmlMultipleSheetTests
    {
        [Fact]
        public void SpecifySheetNameQueryTest()
        {
            var path = @"..\..\..\..\..\samples\xlsx\TestMultiSheet.xlsx";
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
        }
    }
}