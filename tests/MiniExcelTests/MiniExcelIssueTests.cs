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

                foreach (var index in new[] { 0, 2, 5 })
                {
                    Assert.Equal(1, rows[index].實單每日損益);
                    Assert.Equal(2, rows[index].程式每日損益);
                    Assert.Equal("測試商品1", rows[index].商品);
                    Assert.Equal(111.11, rows[index].滿倉口數);
                    Assert.Equal(111.11, rows[index].波段);
                    Assert.Equal(111.11, rows[index].當沖);
                }

                foreach (var index in new[] { 1, 3, 4 })
                {
                    Assert.Null(rows[index].實單每日損益);
                    Assert.Null(rows[index].程式每日損益);
                    Assert.Null(rows[index].商品);
                    Assert.Null(rows[index].滿倉口數);
                    Assert.Null(rows[index].波段);
                    Assert.Null(rows[index].當沖);
                }
            }
            {

                var rows = MiniExcel.Query<Issue138ExcelRow>(path).ToList();
                Assert.Equal(6, rows.Count);
                Assert.Equal(new DateTime(2021,3,1), rows[0].date);

                foreach (var index in new[] { 0,2,5})
                {
                    Assert.Equal(1, rows[index].實單每日損益);
                    Assert.Equal(2, rows[index].程式每日損益);
                    Assert.Equal("測試商品1", rows[index].商品);
                    Assert.Equal(111.11, rows[index].滿倉口數);
                    Assert.Equal(111.11, rows[index].波段);
                    Assert.Equal(111.11, rows[index].當沖);
                }

                foreach (var index in new[] { 1, 3, 4 })
                {
                    Assert.Null(rows[index].實單每日損益);
                    Assert.Null(rows[index].程式每日損益);
                    Assert.Null(rows[index].商品);
                    Assert.Null(rows[index].滿倉口數);
                    Assert.Null(rows[index].波段);
                    Assert.Null(rows[index].當沖);
                }
            }
        }

        public class Issue138ExcelRow
        {
            public DateTime? date { get; set; }
            public int? 實單每日損益 { get; set; }
            public int? 程式每日損益 { get; set; }
            public string 商品 { get; set; }
            public double? 滿倉口數 { get; set; }
            public double? 波段 { get; set; }
            public double? 當沖 { get; set; }
        }
    }
}