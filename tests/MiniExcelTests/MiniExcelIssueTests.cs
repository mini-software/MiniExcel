using Xunit;
using System;
using System.Linq;
using System.IO;

namespace MiniExcelLibs.Tests
{
    public partial class MiniExcelIssueTests
    {
        ///https://github.com/shps951023/MiniExcel/issues/138
        [Fact]
        public void Issue138()
        {
            var path = @"..\..\..\..\..\samples\xlsx\TestIssue138.xlsx";
            {
                var rows = MiniExcel.Query(path, true).ToList();
                Assert.Equal(6, rows.Count);
            }
            {

                var rows = MiniExcel.Query<Issue138ExcelRow>(path).ToList();
                Assert.Equal(6, rows.Count);
            }
        }

        public class Issue138ExcelRow
        {
            public DateTime? date { get; set; }
            public int? 實單每日損益 { get; set; }
            public int? 程式每日損益 { get; set; }
            public string 商品 { get; set; }
            public string 滿倉口數 { get; set; }
            public double? 波段 { get; set; }
            public double? 當沖 { get; set; }
        }
    }
}