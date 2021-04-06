using Xunit;
using System;
using System.Linq;
using System.IO;
using System.Collections.Generic;

namespace MiniExcelLibs.Tests
{
    public partial class MiniExcelIssueTests
    {
        /// <summary>
        /// https://github.com/shps951023/MiniExcel/issues/153
        /// </summary>
        [Fact]
        public void Issue153()
        {
            var path = @"..\..\..\..\..\samples\xlsx\TestIssue153.xlsx";
            var rows = MiniExcel.Query(path,true).First() as IDictionary<string,object>;
            Assert.Equal(new[] { "序号", "代号", "新代号", "名称", "XXX" , "部门名称" , "单位" , "ERP工时   (小时)A", "工时(秒) A/3600", "标准人工工时(秒)", "生产标准机器工时(秒)", "财务、标准机器工时(秒)", "更新日期", "产品机种", "备注","最近一次修改前的标准工时(秒)", "最近一次修改前的标准机时(秒)", "备注1" }
                ,rows.Keys);
        }

        /// <summary>
        /// https://github.com/shps951023/MiniExcel/issues/137
        /// </summary>
        [Fact]
        public void Issue137()
        {
            var path = @"..\..\..\..\..\samples\xlsx\TestIssue137.xlsx";

            {
                var rows = MiniExcel.Query(path).ToList();
                var first = rows[0] as IDictionary<string, object>; //![image](https://user-images.githubusercontent.com/12729184/113266322-ba06e400-9307-11eb-9521-d36abfda75cc.png)
                Assert.Equal(new[] { "A", "B", "C", "D", "E", "F", "G", "H" }, first.Keys.ToArray());
                Assert.Equal(11, rows.Count);
                {
                    var row = rows[0] as IDictionary<string, object>;
                    Assert.Equal("比例", row["A"]);
                    Assert.Equal("商品", row["B"]);
                    Assert.Equal("滿倉口數", row["C"]);
                    Assert.Equal(" ", row["D"]);
                    Assert.Equal(" ", row["E"]);
                    Assert.Equal(" ", row["F"]);
                    Assert.Equal(Double.Parse("0"), row["G"]);
                    Assert.Equal("1為港幣 0為台幣", row["H"]);
                }
                {
                    var row = rows[1] as IDictionary<string, object>;
                    Assert.Equal(double.Parse("1"), row["A"]);
                    Assert.Equal("MTX", row["B"]);
                    Assert.Equal(double.Parse("10"), row["C"]);
                    Assert.Null(row["D"]);
                    Assert.Null(row["E"]);
                    Assert.Null(row["F"]);
                    Assert.Null(row["G"]);
                    Assert.Null(row["H"]);
                }
                {
                    var row = rows[2] as IDictionary<string, object>;
                    Assert.Equal(double.Parse("0.95"), row["A"]);
                }
            }

            // dynamic query with head
            {
                var rows = MiniExcel.Query(path, true).ToList();
                var first = rows[0] as IDictionary<string, object>; //![image](https://user-images.githubusercontent.com/12729184/113266322-ba06e400-9307-11eb-9521-d36abfda75cc.png)
                Assert.Equal(new[] { "比例", "商品", "滿倉口數", "0", "1為港幣 0為台幣" }, first.Keys.ToArray());
                Assert.Equal(10, rows.Count);
                {
                    var row = rows[0] as IDictionary<string, object>;
                    Assert.Equal(double.Parse("1"), row["比例"]);
                    Assert.Equal("MTX", row["商品"]);
                    Assert.Equal(double.Parse("10"), row["滿倉口數"]);
                    Assert.Null(row["0"]);
                    Assert.Null(row["1為港幣 0為台幣"]);
                }

                {
                    var row = rows[1] as IDictionary<string, object>;
                    Assert.Equal(double.Parse("0.95"), row["比例"]);
                }
            }

            {
                var rows = MiniExcel.Query<Issue137ExcelRow>(path).ToList();
                Assert.Equal(10, rows.Count);
                {
                    var row = rows[0] ;
                    Assert.Equal(1, row.比例);
                    Assert.Equal("MTX", row.商品);
                    Assert.Equal(10, row.滿倉口數);
                }

                {
                    var row = rows[1];
                    Assert.Equal(0.95, row.比例);
                }
            }
        }

        public class Issue137ExcelRow
        {
            public double? 比例 { get; set; }
            public string 商品 { get; set; }
            public int? 滿倉口數 { get; set; }
        }


        /// <summary>
        /// https://github.com/shps951023/MiniExcel/issues/138
        /// </summary>
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