using ClosedXML.Excel;
using Dapper;
using ExcelDataReader;
using MiniExcelLibs.Attributes;
using MiniExcelLibs.OpenXml;
using MiniExcelLibs.Tests.Utils;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SQLite;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.IO.Packaging;
using System.Linq;
using Xunit;
using Xunit.Abstractions;
using static MiniExcelLibs.Tests.Utils.MiniExcelOpenXml;

namespace MiniExcelLibs.Tests
{
    public class MiniExcelOpenXmlTests
    {
        private readonly ITestOutputHelper output;
        public MiniExcelOpenXmlTests(ITestOutputHelper output)
        {
            this.output = output;
        }

        [Fact]
        public void GetColumnsTest()
        {
            {
                var path = "../../../../../samples/xlsx/TestTypeMapping.xlsx";
                var columns = MiniExcel.GetColumns(path);
                Assert.Equal(["A", "B", "C", "D", "E", "F", "G", "H"], columns);
            }

            {
                var path = "../../../../../samples/xlsx/TestTypeMapping.xlsx";
                var columns = MiniExcel.GetColumns(path);
                Assert.Equal(8, columns.Count);
            }

            {
                var path = "../../../../../samples/xlsx/TestEmpty.xlsx";
                var columns = MiniExcel.GetColumns(path);
                Assert.Null(columns);
            }
        }

        [Fact]
        public void SaveAsControlChracter()
        {
            string path = GetTempXlsxPath();
            char[] chars =
            [
                '\u0000','\u0001','\u0002','\u0003','\u0004','\u0005','\u0006','\u0007','\u0008',
                '\u0009', //<HT>
                '\u000A', //<LF>
                '\u000B','\u000C',
                 '\u000D', //<CR>
                '\u000E','\u000F','\u0010','\u0011','\u0012','\u0013','\u0014','\u0015','\u0016',
                 '\u0017','\u0018','\u0019','\u001A','\u001B','\u001C','\u001D','\u001E','\u001F','\u007F'
            ];
            var input = chars.Select(s => new { Test = s.ToString() });
            MiniExcel.SaveAs(path, input);

            var rows2 = MiniExcel.Query(path, true).Select(s => s.Test).ToArray();
            var rows1 = MiniExcel.Query<SaveAsControlChracterVO>(path).Select(s => s.Test).ToArray();
        }

        public class SaveAsControlChracterVO
        {
            public string Test { get; set; }
        }

        public class ExcelAttributeDemo
        {
            [ExcelColumnName("Column1")]
            public string Test1 { get; set; }
            [ExcelColumnName("Column2")]
            public string Test2 { get; set; }
            [ExcelIgnore]
            public string Test3 { get; set; }
            [ExcelColumnIndex("I")] // system will convert "I" to 8 index
            public string Test4 { get; set; }
            public string Test5 { get; } //wihout set will ignore
            public string Test6 { get; private set; } //un-public set will ignore
            [ExcelColumnIndex(3)] // start with 0
            public string Test7 { get; set; }
        }

        [Fact]
        public void CustomAttributeWihoutVaildPropertiesTest()
        {
            var path = "../../../../../samples/xlsx/TestCustomExcelColumnAttribute.xlsx";
            Assert.Throws<InvalidOperationException>(() => MiniExcel.Query<CustomAttributesWihoutVaildPropertiesTestPoco>(path).ToList());
        }

        [Fact]
        public void QueryCustomAttributesTest()
        {
            var path = "../../../../../samples/xlsx/TestCustomExcelColumnAttribute.xlsx";
            var rows = MiniExcel.Query<ExcelAttributeDemo>(path).ToList();
            Assert.Equal("Column1", rows[0].Test1);
            Assert.Equal("Column2", rows[0].Test2);
            Assert.Null(rows[0].Test3);
            Assert.Equal("Test7", rows[0].Test4);
            Assert.Null(rows[0].Test5);
            Assert.Null(rows[0].Test6);
            Assert.Equal("Test4", rows[0].Test7);
        }

        [Fact]
        public void SaveAsCustomAttributesTest()
        {
            string path = GetTempXlsxPath();
            var input = Enumerable.Range(1, 3).Select(
                s => new ExcelAttributeDemo
                {
                    Test1 = "Test1",
                    Test2 = "Test2",
                    Test3 = "Test3",
                    Test4 = "Test4",
                }
            );
            MiniExcel.SaveAs(path, input);
            {
                var rows = MiniExcel.Query(path, true).ToList();
                var first = rows[0] as IDictionary<string, object>;
                Assert.Equal(new[] { "Column1", "Column2", "Test5", "Test7", "Test6", "Test4" }, first.Keys);
                Assert.Equal("Test1", rows[0].Column1);
                Assert.Equal("Test2", rows[0].Column2);
                Assert.Equal("Test4", rows[0].Test4);
                Assert.Null(rows[0].Test5);
                Assert.Null(rows[0].Test6);

                Assert.Equal(3, rows.Count);
            }
        }

        private static string GetTempXlsxPath()
        {
            return Path.Combine(Path.GetTempPath(), $"{Guid.NewGuid()}.xlsx");
        }

        public class CustomAttributesWihoutVaildPropertiesTestPoco
        {
            [ExcelIgnore]
            public string Test3 { get; set; }
            public string Test5 { get; }
            public string Test6 { get; private set; }
        }



        [Fact]
        public void QueryCastToIDictionary()
        {
            var path = @"../../../../../samples/xlsx/TestCenterEmptyRow/TestCenterEmptyRow.xlsx";
            foreach (IDictionary<string, object> row in MiniExcel.Query(path))
            {

            }
        }

        [Fact]
        public void QueryRangeToIDictionary()
        {
            var path = "../../../../../samples/xlsx/TestCenterEmptyRow/TestCenterEmptyRow.xlsx";
            // tips：Only uppercase letters are effective
            var rows = MiniExcel.QueryRange(path, startCell: "A2", endCell: "C7")
                .Cast<IDictionary<string, object>>()
                .ToList();
            Assert.Equal(5, rows.Count);
            Assert.Equal(3, rows[0].Count);
            Assert.Equal(2d, rows[1]["B"]);
            Assert.Equal(null!, rows[2]["A"]);
        }

        [Fact]
        public void CenterEmptyRowsQueryTest()
        {
            var path = "../../../../../samples/xlsx/TestCenterEmptyRow/TestCenterEmptyRow.xlsx";
            using (var stream = File.OpenRead(path))
            {
                var rows = stream.Query().ToList();

                Assert.Equal("a", rows[0].A);
                Assert.Equal("b", rows[0].B);
                Assert.Equal("c", rows[0].C);
                Assert.Equal("d", rows[0].D);

                Assert.Equal(1, rows[1].A);
                Assert.Equal(null, rows[1].B);
                Assert.Equal(3, rows[1].C);
                Assert.Equal(null, rows[1].D);

                Assert.Equal(null, rows[2].A);
                Assert.Equal(2, rows[2].B);
                Assert.Equal(null, rows[2].C);
                Assert.Equal(4, rows[2].D);

                Assert.Equal(null, rows[3].A);
                Assert.Equal(null, rows[3].B);
                Assert.Equal(null, rows[3].C);
                Assert.Equal(null, rows[3].D);

                Assert.Equal(1, rows[4].A);
                Assert.Equal(null, rows[4].B);
                Assert.Equal(3, rows[4].C);
                Assert.Equal(null, rows[4].D);

                Assert.Equal(null, rows[5].A);
                Assert.Equal(2, rows[5].B);
                Assert.Equal(null, rows[5].C);
                Assert.Equal(4, rows[5].D);
            }

            using (var stream = File.OpenRead(path))
            {
                var rows = stream.Query(useHeaderRow: true).ToList();

                Assert.Equal(1, rows[0].a);
                Assert.Equal(null, rows[0].b);
                Assert.Equal(3, rows[0].c);
                Assert.Equal(null, rows[0].d);

                Assert.Equal(null, rows[1].a);
                Assert.Equal(2, rows[1].b);
                Assert.Equal(null, rows[1].c);
                Assert.Equal(4, rows[1].d);

                Assert.Equal(null, rows[2].a);
                Assert.Equal(null, rows[2].b);
                Assert.Equal(null, rows[2].c);
                Assert.Equal(null, rows[2].d);

                Assert.Equal(1, rows[3].a);
                Assert.Equal(null, rows[3].b);
                Assert.Equal(3, rows[3].c);
                Assert.Equal(null, rows[3].d);

                Assert.Equal(null, rows[4].a);
                Assert.Equal(2, rows[4].b);
                Assert.Equal(null, rows[4].c);
                Assert.Equal(4, rows[4].d);
            }
        }
        
        [Fact]
        public void TestEmptyRowsQuerySelfClosingTag()
        {
            var path = "../../../../../samples/xlsx/TestEmptySelfClosingRow.xlsx";
            using (var stream = File.OpenRead(path))
            {
                var rows = stream.Query().ToList();

                Assert.Equal(null, rows[0].A);
                Assert.Equal(1, rows[1].A);
                Assert.Equal(null, rows[2].A);
                Assert.Equal(2, rows[3].A);
                Assert.Equal(null, rows[4].A);
                Assert.Equal(null, rows[5].A);
                Assert.Equal(null, rows[6].A);
                Assert.Equal(null, rows[7].A);
                Assert.Equal(null, rows[8].A);
                Assert.Equal(1, rows[9].A);
            }
        }

        [Fact]
        public void TestDynamicQueryBasic_WithoutHead()
        {
            var path = "../../../../../samples/xlsx/TestDynamicQueryBasic_WithoutHead.xlsx";
            using (var stream = File.OpenRead(path))
            {
                var rows = stream.Query().ToList();

                Assert.Equal("MiniExcel", rows[0].A);
                Assert.Equal(1, rows[0].B);
                Assert.Equal("Github", rows[1].A);
                Assert.Equal(2, rows[1].B);
            }
        }

        [Fact]
        public void TestDynamicQueryBasic_useHeaderRow()
        {
            var path = "../../../../../samples/xlsx/TestDynamicQueryBasic.xlsx";
            using (var stream = File.OpenRead(path))
            {
                var rows = stream.Query(useHeaderRow: true).ToList();

                Assert.Equal("MiniExcel", rows[0].Column1);
                Assert.Equal(1, rows[0].Column2);
                Assert.Equal("Github", rows[1].Column1);
                Assert.Equal(2, rows[1].Column2);
            }

            {
                var rows = MiniExcel.Query(path, useHeaderRow: true).ToList();

                Assert.Equal("MiniExcel", rows[0].Column1);
                Assert.Equal(1, rows[0].Column2);
                Assert.Equal("Github", rows[1].Column1);
                Assert.Equal(2, rows[1].Column2);
            }
        }

        public class DemoPocoHelloWorld
        {
            public string HelloWorld1 { get; set; }
        }

        public class UserAccount
        {
            public Guid ID { get; set; }
            public string Name { get; set; }
            public DateTime BoD { get; set; }
            public int Age { get; set; }
            public bool VIP { get; set; }
            public decimal Points { get; set; }
            public int IgnoredProperty => 1;
        }

        [Fact]
        public void QueryStrongTypeMapping_Test()
        {
            var path = "../../../../../samples/xlsx/TestTypeMapping.xlsx";
            using (var stream = File.OpenRead(path))
            {
                var rows = stream.Query<UserAccount>().ToList();

                Assert.Equal(100, rows.Count);

                Assert.Equal(Guid.Parse("78DE23D2-DCB6-BD3D-EC67-C112BBC322A2"), rows[0].ID);
                Assert.Equal("Wade", rows[0].Name);
                Assert.Equal(DateTime.ParseExact("27/09/2020", "dd/MM/yyyy", CultureInfo.InvariantCulture), rows[0].BoD);
                Assert.Equal(36, rows[0].Age);
                Assert.False(rows[0].VIP);
                Assert.Equal(5019.12m, rows[0].Points);
                Assert.Equal(1, rows[0].IgnoredProperty);
            }

            {
                var rows = MiniExcel.Query<UserAccount>(path).ToList();

                Assert.Equal(100, rows.Count);

                Assert.Equal(Guid.Parse("78DE23D2-DCB6-BD3D-EC67-C112BBC322A2"), rows[0].ID);
                Assert.Equal("Wade", rows[0].Name);
                Assert.Equal(DateTime.ParseExact("27/09/2020", "dd/MM/yyyy", CultureInfo.InvariantCulture), rows[0].BoD);
                Assert.Equal(36, rows[0].Age);
                Assert.False(rows[0].VIP);
                Assert.Equal(5019.12m, rows[0].Points);
                Assert.Equal(1, rows[0].IgnoredProperty);
            }
        }


        public class AutoCheckType
        {
            public Guid? Guid { get; set; }
            public bool? Bool { get; set; }
            public DateTime? Datetime { get; set; }
            public string String { get; set; }
        }

        [Fact]
        public void AutoCheckTypeTest()
        {
            var path = "../../../../../samples/xlsx/TestTypeMapping_AutoCheckFormat.xlsx";
            using (var stream = FileHelper.OpenRead(path))
            {
                var rows = stream.Query<AutoCheckType>().ToList();
            }
        }

        public class ExcelUriDemo
        {
            public string Name { get; set; }
            public int Age { get; set; }
            public Uri Url { get; set; }
        }

        [Fact]
        public void UriMappingTest()
        {
            var path = "../../../../../samples/xlsx/TestUriMapping.xlsx";
            using (var stream = File.OpenRead(path))
            {
                var rows = stream.Query<ExcelUriDemo>().ToList();

                Assert.Equal("Felix", rows[1].Name);
                Assert.Equal(44, rows[1].Age);
                Assert.Equal(new Uri("https://friendly-utilization.net"), rows[1].Url);
            }
        }

        public class SimpleAccount
        {
            public string Name { get; set; }
            public int Age { get; set; }
            public string Mail { get; set; }
            public decimal Points { get; set; }
        }
        [Fact]
        public void TrimColumnNamesTest()
        {
            var path = "../../../../../samples/xlsx/TestTrimColumnNames.xlsx";
            var rows = MiniExcel.Query<SimpleAccount>(path).ToList();

            Assert.Equal("Raymond", rows[4].Name);
            Assert.Equal(18, rows[4].Age);
            Assert.Equal("sagittis.lobortis@leoMorbi.com", rows[4].Mail);
            Assert.Equal(8209.76m, rows[4].Points);
        }
        
        [Fact]
        public void TestDatetimeSpanFormat_ClosedXml()
        {
            var path = "../../../../../samples/xlsx/TestDatetimeSpanFormat_ClosedXml.xlsx";
            using (var stream = FileHelper.OpenRead(path))
            {
                var row = stream.Query().First();
                var a = row.A;
                var b = row.B;
                Assert.Equal(DateTime.Parse("2021-03-20T23:39:42.3130000"), (DateTime)a);
                Assert.Equal(TimeSpan.FromHours(10), (TimeSpan)b);
            }
        }

        [Fact]
        public void LargeFileQueryStrongTypeMapping_Test()
        {
            var path = "../../../../../benchmarks/MiniExcel.Benchmarks/Test1,000,000x10.xlsx";
            using (var stream = File.OpenRead(path))
            {
                var rows = stream.Query<DemoPocoHelloWorld>().Take(2).ToList();

                Assert.Equal("HelloWorld2", rows[0].HelloWorld1);
                Assert.Equal("HelloWorld3", rows[1].HelloWorld1);
            }
            {
                var rows = MiniExcel.Query<DemoPocoHelloWorld>(path).Take(2).ToList();

                Assert.Equal("HelloWorld2", rows[0].HelloWorld1);
                Assert.Equal("HelloWorld3", rows[1].HelloWorld1);
            }
        }

        [Theory]
        [InlineData("../../../../../samples/xlsx/ExcelDataReaderCollections/TestChess.xlsx")]
        [InlineData("../../../../../samples/xlsx/TestCenterEmptyRow/TestCenterEmptyRow.xlsx")]
        public void QueryExcelDataReaderCheckTest(string path)
        {
#if NETCOREAPP3_1 || NET5_0
            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
#endif

            DataSet exceldatareaderResult;
            using (var stream = File.OpenRead(path))
            using (var reader = ExcelReaderFactory.CreateReader(stream))
            {
                exceldatareaderResult = reader.AsDataSet();
            }

            using (var stream = File.OpenRead(path))
            {
                var rows = stream.Query().ToList();
                Assert.Equal(exceldatareaderResult.Tables[0].Rows.Count, rows.Count);
                foreach (IDictionary<string, object> row in rows)
                {
                    var rowIndex = rows.IndexOf(row);
                    var keys = row.Keys;
                    foreach (var key in keys)
                    {
                        var eV = exceldatareaderResult.Tables[0].Rows[rowIndex][Helpers.GetColumnIndex(key)];
                        var v = row[key] == null ? DBNull.Value : row[key];
                        Assert.Equal(eV, v);
                    }
                }
            }
        }

        [Fact]
        public void QueryCustomStyle()
        {
            var path = "../../../../../samples/xlsx/TestWihoutRAttribute.xlsx";
            using (var stream = File.OpenRead(path))
            {

            }
        }

        [Fact]
        public void QuerySheetWithoutRAttribute()
        {
            var path = "../../../../../samples/xlsx/TestWihoutRAttribute.xlsx";
            using (var stream = File.OpenRead(path))
            {
                var rows = stream.Query().ToList();
                var keys = (rows.First() as IDictionary<string, object>)!.Keys;

                Assert.Equal(2, rows.Count);
                Assert.Equal(5, keys.Count);

                Assert.Equal(1, rows[0].A);
                //Assert.Equal(@""" <> +}{/nHello World]", (string)rows[0].B);
                Assert.Equal(null, rows[0].C);
                Assert.Equal(null, rows[0].D);
                Assert.Equal(null, rows[0].E);

                Assert.Equal(1, rows[1].A);
                Assert.Equal("\"<>+}{\\nHello World", rows[1].B);
                Assert.Equal(true, rows[1].C);
                Assert.Equal("2021-03-16T19:10:21", rows[1].D);
            }
        }

        [Fact]
        public void FixDimensionJustOneColumnParsingError_Test()
        {
            {
                var path = "../../../../../samples/xlsx/TestDimensionC3.xlsx";
                using (var stream = File.OpenRead(path))
                {
                    var rows = stream.Query().ToList();
                    var keys = (rows.First() as IDictionary<string, object>)!.Keys;
                    Assert.Equal(3, keys.Count);
                    Assert.Equal(2, rows.Count);
                }
            }
        }

        public class SaveAsFileWithDimensionByICollectionTestType
        {
            public string A { get; set; }
            public string B { get; set; }
        }
        [Fact]
        public void SaveAsFileWithDimensionByICollection()
        {
            //List<strongtype>
            {
                var path = Path.Combine(Path.GetTempPath(), $"{Guid.NewGuid()}.xlsx");
                List<SaveAsFileWithDimensionByICollectionTestType> values =
                [
                    new() { A = "A", B = "B" },
                    new() { A = "A", B = "B" }
                ];
                MiniExcel.SaveAs(path, values);
                {
                    using (var stream = File.OpenRead(path))
                    {
                        var rows = stream.Query(useHeaderRow: false).ToList();
                        Assert.Equal(3, rows.Count);
                        Assert.Equal("A", rows[0].A);
                        Assert.Equal("A", rows[1].A);
                        Assert.Equal("A", rows[2].A);
                    }
                    using (var stream = File.OpenRead(path))
                    {
                        var rows = stream.Query(useHeaderRow: true).ToList();
                        Assert.Equal(2, rows.Count);
                        Assert.Equal("A", rows[0].A);
                        Assert.Equal("A", rows[1].A);
                    }
                }

                Assert.Equal("A1:B3", Helpers.GetFirstSheetDimensionRefValue(path));
                File.Delete(path);

                MiniExcel.SaveAs(path, values, false);
                Assert.Equal("A1:B2", Helpers.GetFirstSheetDimensionRefValue(path));
                File.Delete(path);
            }

            //List<strongtype> empty
            {
                var path = Path.Combine(Path.GetTempPath(), $"{Guid.NewGuid()}.xlsx");
                var values = new List<SaveAsFileWithDimensionByICollectionTestType>()
                {
                };
                MiniExcel.SaveAs(path, values, false);
                {
                    using (var stream = File.OpenRead(path))
                    {
                        var rows = stream.Query(useHeaderRow: false).ToList();
                        Assert.Empty(rows);
                    }
                    Assert.Equal("A1:B1", Helpers.GetFirstSheetDimensionRefValue(path));
                }
                File.Delete(path);


                MiniExcel.SaveAs(path, values, true);
                {
                    using (var stream = File.OpenRead(path))
                    {
                        var rows = stream.Query(useHeaderRow: false).ToList();
                        Assert.Single(rows);
                    }
                }
                Assert.Equal("A1:B1", Helpers.GetFirstSheetDimensionRefValue(path));
                File.Delete(path);
            }

            //Array<anoymous>
            {
                var path = Path.Combine(Path.GetTempPath(), $"{Guid.NewGuid()}.xlsx");
                var values = new[]
                {
                    new {A="A",B="B"},
                    new {A="A",B="B"},
                };
                MiniExcel.SaveAs(path, values);
                {
                    using (var stream = File.OpenRead(path))
                    {
                        var rows = stream.Query(useHeaderRow: false).ToList();
                        Assert.Equal(3, rows.Count);
                        Assert.Equal("A", rows[0].A);
                        Assert.Equal("A", rows[1].A);
                        Assert.Equal("A", rows[2].A);
                    }
                    using (var stream = File.OpenRead(path))
                    {
                        var rows = stream.Query(useHeaderRow: true).ToList();
                        Assert.Equal(2, rows.Count);
                        Assert.Equal("A", rows[0].A);
                        Assert.Equal("A", rows[1].A);
                    }
                }

                Assert.Equal("A1:B3", Helpers.GetFirstSheetDimensionRefValue(path));
                File.Delete(path);

                MiniExcel.SaveAs(path, values, false);
                Assert.Equal("A1:B2", Helpers.GetFirstSheetDimensionRefValue(path));
                File.Delete(path);
            }

            // without properties
            {
                var path = Path.Combine(Path.GetTempPath(), $"{Guid.NewGuid()}.xlsx");
                var values = new List<int>();
                Assert.Throws<NotImplementedException>(() => MiniExcel.SaveAs(path, values));
            }
        }

        [Fact]
        public void SaveAsFileWithDimension()
        {
            {
                var path = Path.Combine(Path.GetTempPath(), $"{Guid.NewGuid()}.xlsx");
                var table = new DataTable();
                MiniExcel.SaveAs(path, table);
                Assert.Equal("A1", Helpers.GetFirstSheetDimensionRefValue(path));
                {
                    using (var stream = File.OpenRead(path))
                    {
                        var rows = stream.Query().ToList();
                        Assert.Single(rows); //TODO:
                    }
                    File.Delete(path);
                }

                MiniExcel.SaveAs(path, table, printHeader: false);
                Assert.Equal("A1", Helpers.GetFirstSheetDimensionRefValue(path));
                File.Delete(path);
            }

            {
                var path = Path.Combine(Path.GetTempPath(), $"{Guid.NewGuid()}.xlsx");
                var table = new DataTable();
                {
                    table.Columns.Add("a", typeof(string));
                    table.Columns.Add("b", typeof(decimal));
                    table.Columns.Add("c", typeof(bool));
                    table.Columns.Add("d", typeof(DateTime));
                    table.Rows.Add(@"""<>+-*//}{\\n", 1234567890);
                    table.Rows.Add("<test>Hello World</test>", -1234567890, false, DateTime.Now);
                }
                MiniExcel.SaveAs(path, table);
                Assert.Equal("A1:D3", Helpers.GetFirstSheetDimensionRefValue(path));

                {
                    using (var stream = File.OpenRead(path))
                    {
                        var rows = stream.Query(useHeaderRow: true).ToList();
                        Assert.Equal(2, rows.Count);
                        Assert.Equal(@"""<>+-*//}{\\n", rows[0].a);
                        Assert.Equal(1234567890, rows[0].b);
                        Assert.Equal(null, rows[0].c);
                        Assert.Equal(null, rows[0].d);
                    }

                    using (var stream = File.OpenRead(path))
                    {
                        var rows = stream.Query().ToList();
                        Assert.Equal(3, rows.Count);
                        Assert.Equal("a", rows[0].A);
                        Assert.Equal("b", rows[0].B);
                        Assert.Equal("c", rows[0].C);
                        Assert.Equal("d", rows[0].D);
                    }
                }

                File.Delete(path);

                MiniExcel.SaveAs(path, table, printHeader: false);
                Assert.Equal("A1:D2", Helpers.GetFirstSheetDimensionRefValue(path));
                File.Delete(path);
            }

            //TODO:StartCell

            {
                var path = Path.Combine(Path.GetTempPath(), $"{Guid.NewGuid()}.xlsx");
                var table = new DataTable();
                {
                    table.Columns.Add("a", typeof(string));
                    table.Rows.Add("A");
                    table.Rows.Add("B");
                }
                MiniExcel.SaveAs(path, table);
                Assert.Equal("A3", Helpers.GetFirstSheetDimensionRefValue(path));
                File.Delete(path);
            }
        }

        [Fact]
        public void SaveAsByDataTableTest()
        {
            {
                var now = DateTime.Now;
                var path = Path.Combine(Path.GetTempPath(), $"{Guid.NewGuid()}.xlsx");

                var table = new DataTable();
                {
                    table.Columns.Add("a", typeof(string));
                    table.Columns.Add("b", typeof(decimal));
                    table.Columns.Add("c", typeof(bool));
                    table.Columns.Add("d", typeof(DateTime));
                    table.Rows.Add(@"""<>+-*//}{\\n", 1234567890, true, now);
                    table.Rows.Add("<test>Hello World</test>", -1234567890, false, now.Date);
                }

                MiniExcel.SaveAs(path, table, sheetName: "R&D");

                using (var p = new ExcelPackage(new FileInfo(path)))
                {
                    var ws = p.Workbook.Worksheets.First();

                    Assert.True(ws.Cells["A1"].Value.ToString() == "a");
                    Assert.True(ws.Cells["B1"].Value.ToString() == "b");
                    Assert.True(ws.Cells["C1"].Value.ToString() == "c");
                    Assert.True(ws.Cells["D1"].Value.ToString() == "d");

                    Assert.True(ws.Cells["A2"].Value.ToString() == @"""<>+-*//}{\\n");
                    Assert.True(ws.Cells["B2"].Value.ToString() == "1234567890");
                    Assert.True(ws.Cells["C2"].Value.ToString() == true.ToString());
                    Assert.True(ws.Cells["D2"].Value.ToString() == now.ToString());

                    Assert.True(ws.Name == "R&D");
                }

                File.Delete(path);
            }
            {
                var path = Path.Combine(Path.GetTempPath(), $"{Guid.NewGuid()}.xlsx");
                var table = new DataTable();
                {
                    table.Columns.Add("Column1", typeof(string));
                    table.Columns.Add("Column2", typeof(int));
                    table.Rows.Add("MiniExcel", 1);
                    table.Rows.Add("Github", 2);
                }

                MiniExcel.SaveAs(path, table);

                File.Delete(path);
            }
        }

        [Fact]
        public void QueryByLINQExtensionsAvoidLargeFileOOMTest()
        {
            var path = "../../../../../benchmarks/MiniExcel.Benchmarks/Test1,000,000x10.xlsx";

            {
                var row = MiniExcel.Query(path).First();
                Assert.Equal("HelloWorld1", row.A);
            }

            using (var stream = File.OpenRead(path))
            {
                var row = stream.Query().First();
                Assert.Equal("HelloWorld1", row.A);
            }

            {
                var rows = MiniExcel.Query(path).Take(10);
                Assert.Equal(10, rows.Count());
            }
        }

        [Fact]
        public void EmptyTest()
        {
            {
                var path = Path.Combine(Path.GetTempPath(), $"{Guid.NewGuid()}.xlsx");
                using (var connection = Db.GetConnection("Data Source=:memory:"))
                {
                    var rows = connection.Query("with cte as (select 1 id,2 val) select * from cte where 1=2");
                    MiniExcel.SaveAs(path, rows);
                }
                using (var stream = File.OpenRead(path))
                {
                    var rows = stream.Query(useHeaderRow: true).ToList();
                    Assert.Empty(rows);
                }
                File.Delete(path);
            }
        }

        [Fact]
        public void SaveAsByIEnumerableIDictionary()
        {
            var path = Path.Combine(Path.GetTempPath(), $"{Guid.NewGuid()}.xlsx");

            {
                var values = new List<Dictionary<string, object>>()
                {
                    new() { { "Column1", "MiniExcel" }, { "Column2", 1 } },
                    new() { { "Column1", "Github" }, { "Column2", 2 } },
                };
                var sheets = new Dictionary<string, object>
                {
                    ["R&D"] = values,
                    ["success!"] = values
                };
                MiniExcel.SaveAs(path, sheets);

                using (var stream = File.OpenRead(path))
                {
                    var rows = stream.Query(useHeaderRow: false).ToList();
                    Assert.Equal("Column1", rows[0].A);
                    Assert.Equal("Column2", rows[0].B);
                    Assert.Equal("MiniExcel", rows[1].A);
                    Assert.Equal(1, rows[1].B);
                    Assert.Equal("Github", rows[2].A);
                    Assert.Equal(2, rows[2].B);

                    Assert.Equal("R&D", stream.GetSheetNames()[0]);
                }

                using (var stream = File.OpenRead(path))
                {
                    var rows = stream.Query(useHeaderRow: true).ToList();

                    Assert.Equal(2, rows.Count);
                    Assert.Equal("MiniExcel", rows[0].Column1);
                    Assert.Equal(1, rows[0].Column2);
                    Assert.Equal("Github", rows[1].Column1);
                    Assert.Equal(2, rows[1].Column2);

                    Assert.Equal("success!", stream.GetSheetNames()[1]);
                }

                Assert.Equal("A1:B3", Helpers.GetFirstSheetDimensionRefValue(path));
                File.Delete(path);
            }

            {
                var values = new List<Dictionary<int, object>>()
                {
                    new() { { 1, "MiniExcel"}, { 2, 1 } },
                    new() { { 1, "Github" }, { 2, 2 } },
                };
                MiniExcel.SaveAs(path, values);

                using (var stream = File.OpenRead(path))
                {
                    var rows = stream.Query(useHeaderRow: false).ToList();
                    Assert.Equal(3, rows.Count);
                }

                Assert.Equal("A1:B3", Helpers.GetFirstSheetDimensionRefValue(path));
                File.Delete(path);
            }
        }

        [Fact]
        public void SaveAsFrozenRowsAndColumnsTest()
        {
            var config = new OpenXmlConfiguration
            {
                FreezeRowCount = 1,
                FreezeColumnCount = 2
            };

            {
                // Test enumerable
                var path = Path.Combine(Path.GetTempPath(), $"{Guid.NewGuid()}.xlsx");
                MiniExcel.SaveAs(
                    path,
                    new[] 
                    {
                        new { Column1 = "MiniExcel", Column2 = 1 },
                        new { Column1 = "Github", Column2 = 2}
                    },
                    configuration: config
                );

                using (var stream = File.OpenRead(path))
                {
                    var rows = stream.Query(useHeaderRow: true).ToList();

                    Assert.Equal("MiniExcel", rows[0].Column1);
                    Assert.Equal(1, rows[0].Column2);
                    Assert.Equal("Github", rows[1].Column1);
                    Assert.Equal(2, rows[1].Column2);
                }

                Assert.Equal("A1:B3", Helpers.GetFirstSheetDimensionRefValue(path));
                File.Delete(path);
            }

            {
                // test table
                var table = new DataTable();
                {
                    table.Columns.Add("a", typeof(string));
                    table.Columns.Add("b", typeof(decimal));
                    table.Columns.Add("c", typeof(bool));
                    table.Columns.Add("d", typeof(DateTime));
                    table.Rows.Add("some text", 1234567890, true, DateTime.Now);
                    table.Rows.Add("<test>Hello World</test>", -1234567890, false, DateTime.Now.Date);
                }
                var pathTable = Path.Combine(Path.GetTempPath(), $"{Guid.NewGuid()}.xlsx");
                MiniExcel.SaveAs(pathTable, table, configuration: config);

                Assert.Equal("A1:D3", Helpers.GetFirstSheetDimensionRefValue(pathTable));


                // data reader
                var reader = table.CreateDataReader();
                var pathReader = Path.Combine(Path.GetTempPath(), $"{Guid.NewGuid()}.xlsx");

                MiniExcel.SaveAs(pathReader, reader, configuration: config);
                Assert.Equal("A1:D3", Helpers.GetFirstSheetDimensionRefValue(pathTable));
            }

        }

        [Fact]
        public void SaveAsByDapperRows()
        {
            var path = Path.Combine(Path.GetTempPath(), $"{Guid.NewGuid()}.xlsx");


            // Dapper Query
            using (var connection = Db.GetConnection("Data Source=:memory:"))
            {
                var rows = connection.Query("select 'MiniExcel' as Column1,1 as Column2 union all select 'Github',2");
                MiniExcel.SaveAs(path, rows);
            }

            Assert.Equal("A1:B3", Helpers.GetFirstSheetDimensionRefValue(path));

            using (var stream = File.OpenRead(path))
            {
                var rows = stream.Query(useHeaderRow: true).ToList();

                Assert.Equal("MiniExcel", rows[0].Column1);
                Assert.Equal(1, rows[0].Column2);
                Assert.Equal("Github", rows[1].Column1);
                Assert.Equal(2, rows[1].Column2);
            }

            File.Delete(path);

            // Empty
            using (var connection = Db.GetConnection("Data Source=:memory:"))
            {
                var rows = connection.Query("with cte as (select 'MiniExcel' as Column1,1 as Column2 union all select 'Github',2)select * from cte where 1=2").ToList();
                MiniExcel.SaveAs(path, rows);
            }

            using (var stream = File.OpenRead(path))
            {
                var rows = stream.Query(useHeaderRow: false).ToList();
                Assert.Empty(rows);
            }

            using (var stream = File.OpenRead(path))
            {
                var rows = stream.Query(useHeaderRow: true).ToList();
                Assert.Empty(rows);
            }

            Assert.Equal("A1", Helpers.GetFirstSheetDimensionRefValue(path));
            File.Delete(path);


            // ToList
            using (var connection = Db.GetConnection("Data Source=:memory:"))
            {
                var rows = connection.Query("select 'MiniExcel' as Column1,1 as Column2 union all select 'Github',2").ToList();
                MiniExcel.SaveAs(path, rows);
            }

            Assert.Equal("A1:B3", Helpers.GetFirstSheetDimensionRefValue(path));

            using (var stream = File.OpenRead(path))
            {
                var rows = stream.Query(useHeaderRow: false).ToList();

                Assert.Equal("Column1", rows[0].A);
                Assert.Equal("Column2", rows[0].B);
                Assert.Equal("MiniExcel", rows[1].A);
                Assert.Equal(1, rows[1].B);
                Assert.Equal("Github", rows[2].A);
                Assert.Equal(2, rows[2].B);
            }

            using (var stream = File.OpenRead(path))
            {
                var rows = stream.Query(useHeaderRow: true).ToList();

                Assert.Equal("MiniExcel", rows[0].Column1);
                Assert.Equal(1, rows[0].Column2);
                Assert.Equal("Github", rows[1].Column1);
                Assert.Equal(2, rows[1].Column2);
            }
            File.Delete(path);
        }


        public class Demo
        {
            public string Column1 { get; set; }
            public decimal Column2 { get; set; }
        }
        [Fact]
        public void QueryByStrongTypeParameterTest()
        {
            var path = Path.Combine(Path.GetTempPath(), $"{Guid.NewGuid()}.xlsx");

            List<Demo> values =
            [
                new() { Column1 = "MiniExcel", Column2 = 1 },
                new() { Column1 = "Github", Column2 = 2 }
            ];
            MiniExcel.SaveAs(path, values);


            using (var stream = File.OpenRead(path))
            {
                var rows = stream.Query(useHeaderRow: true).ToList();

                Assert.Equal("MiniExcel", rows[0].Column1);
                Assert.Equal(1, rows[0].Column2);
                Assert.Equal("Github", rows[1].Column1);
                Assert.Equal(2, rows[1].Column2);
            }

            File.Delete(path);
        }

        [Fact]
        public void QueryByDictionaryStringAndObjectParameterTest()
        {
            var path = Path.Combine(Path.GetTempPath(), $"{Guid.NewGuid()}.xlsx");

            List<Dictionary<string, object>> values =
            [
                new() { { "Column1", "MiniExcel" }, { "Column2", 1 } },
                new() { { "Column1", "Github" }, { "Column2", 2 } }
            ];
            MiniExcel.SaveAs(path, values);


            using (var stream = File.OpenRead(path))
            {
                var rows = stream.Query(useHeaderRow: true).ToList();

                Assert.Equal("MiniExcel", rows[0].Column1);
                Assert.Equal(1, rows[0].Column2);
                Assert.Equal("Github", rows[1].Column1);
                Assert.Equal(2, rows[1].Column2);
            }

            File.Delete(path);
        }

        [Fact]
        public void SQLiteInsertTest()
        {
            // Avoid SQL Insert Large Size Xlsx OOM
            var path = "../../../../../samples/xlsx/Test5x2.xlsx";
            var tempSqlitePath = Path.Combine(Path.GetTempPath(), $"{Guid.NewGuid()}.db");
            var connectionString = $"Data Source={tempSqlitePath};Version=3;";

            using (var connection = new SQLiteConnection(connectionString))
            {
                connection.Execute("create table T (A varchar(20),B varchar(20));");
            }

            using (var connection = new SQLiteConnection(connectionString))
            {
                connection.Open();
                using (var transaction = connection.BeginTransaction())
                using (var stream = File.OpenRead(path))
                {
                    var rows = stream.Query();
                    foreach (var row in rows)
                        connection.Execute("insert into T (A,B) values (@A,@B)", new { row.A, row.B }, transaction: transaction);
                    transaction.Commit();
                }
            }

            using (var connection = new SQLiteConnection(connectionString))
            {
                var result = connection.Query("select * from T");
                Assert.Equal(5, result.Count());
            }

            File.Delete(tempSqlitePath);
        }

        [Fact]
        public void SaveAsBasicCreateTest()
        {
            var path = Path.Combine(Path.GetTempPath(), $"{Guid.NewGuid()}.xlsx");
            var rowsWritten = MiniExcel.SaveAs(path, new[] 
            {
                new { Column1 = "MiniExcel", Column2 = 1 },
                new { Column1 = "Github", Column2 = 2}
            });
            Assert.Single(rowsWritten);
            Assert.Equal(2, rowsWritten[0]);

            using (var stream = File.OpenRead(path))
            {
                var rows = stream.Query(useHeaderRow: true).ToList();

                Assert.Equal("MiniExcel", rows[0].Column1);
                Assert.Equal(1, rows[0].Column2);
                Assert.Equal("Github", rows[1].Column1);
                Assert.Equal(2, rows[1].Column2);
            }

            Assert.Equal("A1:B3", Helpers.GetFirstSheetDimensionRefValue(path));

            File.Delete(path);
        }

        [Fact]
        public void SaveAsBasicStreamTest()
        {
            {
                var path = Path.Combine(Path.GetTempPath(), $"{Guid.NewGuid()}.xlsx");
                var values = new[] 
                {
                    new { Column1 = "MiniExcel", Column2 = 1 },
                    new { Column1 = "Github", Column2 = 2}
                };
                using (var stream = new FileStream(path, FileMode.CreateNew))
                {
                    var rowsWritten = stream.SaveAs(values);
                    Assert.Single(rowsWritten);
                    Assert.Equal(2, rowsWritten[0]);
                }

                using (var stream = File.OpenRead(path))
                {
                    var rows = stream.Query(useHeaderRow: true).ToList();

                    Assert.Equal("MiniExcel", rows[0].Column1);
                    Assert.Equal(1, rows[0].Column2);
                    Assert.Equal("Github", rows[1].Column1);
                    Assert.Equal(2, rows[1].Column2);
                }

                File.Delete(path);
            }
            {
                var path = Path.Combine(Path.GetTempPath(), $"{Guid.NewGuid()}.xlsx");
                var values = new[] 
                {
                    new { Column1 = "MiniExcel", Column2 = 1 },
                    new { Column1 = "Github", Column2 = 2}
                };
                using (var stream = new MemoryStream())
                using (var fileStream = new FileStream(path, FileMode.Create))
                {
                    var rowsWritten = stream.SaveAs(values);
                    stream.Seek(0, SeekOrigin.Begin);
                    stream.CopyTo(fileStream);
                    Assert.Single(rowsWritten);
                    Assert.Equal(2, rowsWritten[0]);
                }

                using (var stream = File.OpenRead(path))
                {
                    var rows = stream.Query(useHeaderRow: true).ToList();

                    Assert.Equal("MiniExcel", rows[0].Column1);
                    Assert.Equal(1, rows[0].Column2);
                    Assert.Equal("Github", rows[1].Column1);
                    Assert.Equal(2, rows[1].Column2);
                }

                File.Delete(path);
            }
        }

        [Fact]
        public void SaveAsSpecialAndTypeCreateTest()
        {
            var path = Path.Combine(Path.GetTempPath(), $"{Guid.NewGuid()}.xlsx");
            var rowsWritten = MiniExcel.SaveAs(path, new[] 
            {
                  new { a = @"""<>+-*//}{\\n", b = 1234567890,c = true,d=DateTime.Now },
                  new { a = "<test>Hello World</test>", b = -1234567890,c=false,d=DateTime.Now.Date}
            });
            Assert.Single(rowsWritten);
            Assert.Equal(2, rowsWritten[0]);
            
            var info = new FileInfo(path);
            Assert.True(info.FullName == path);

            File.Delete(path);
        }

        [Fact]
        public void SaveAsFileEpplusCanReadTest()
        {
            var now = DateTime.Now;
            var path = Path.Combine(Path.GetTempPath(), $"{Guid.NewGuid()}.xlsx");
            var rowsWritten = MiniExcel.SaveAs(path, new[]
            {
                new { a = @"""<>+-*//}{\\n", b = 1234567890,c = true,d= now},
                new { a = "<test>Hello World</test>", b = -1234567890,c=false,d=now.Date}
             });
            Assert.Single(rowsWritten);
            Assert.Equal(2, rowsWritten[0]);
            
            using (var p = new ExcelPackage(new FileInfo(path)))
            {
                var ws = p.Workbook.Worksheets.First();

                Assert.True(ws.Cells["A1"].Value.ToString() == "a");
                Assert.True(ws.Cells["B1"].Value.ToString() == "b");
                Assert.True(ws.Cells["C1"].Value.ToString() == "c");
                Assert.True(ws.Cells["D1"].Value.ToString() == "d");

                Assert.True(ws.Cells["A2"].Value.ToString() == @"""<>+-*//}{\\n");
                Assert.True(ws.Cells["B2"].Value.ToString() == "1234567890");
                Assert.True(ws.Cells["C2"].Value.ToString() == true.ToString());
                Assert.True(ws.Cells["D2"].Value.ToString() == now.ToString());
            }
            File.Delete(path);
        }

        [Fact]
        public void SavaAsClosedXmlCanReadTest()
        {
            var now = DateTime.Now;
            var path = Path.Combine(Path.GetTempPath(), $"{Guid.NewGuid()}.xlsx");
            var rowsWritten = MiniExcel.SaveAs(path, new[] 
            {
                new { a = @"""<>+-*//}{\\n", b = 1234567890,c = true,d= now},
                new { a = "<test>Hello World</test>", b = -1234567890,c=false,d=now.Date}
            }, sheetName: "R&D");
            Assert.Single(rowsWritten);
            Assert.Equal(2, rowsWritten[0]);
            
            using (var workbook = new XLWorkbook(path))
            {
                var ws = workbook.Worksheets.First();

                Assert.True(ws.Cell("A1").Value.ToString() == "a");
                Assert.True(ws.Cell("D1").Value.ToString() == "d");
                Assert.True(ws.Cell("B1").Value.ToString() == "b");
                Assert.True(ws.Cell("C1").Value.ToString() == "c");

                Assert.True(ws.Cell("A2").Value.ToString() == @"""<>+-*//}{\\n");
                Assert.True(ws.Cell("B2").Value.ToString() == "1234567890");
                Assert.True(ws.Cell("C2").Value.ToString() == true.ToString());
                Assert.True(ws.Cell("D2").Value.ToString() == now.ToString());

                Assert.True(ws.Name == "R&D");
            }
            File.Delete(path);
        }

        [Fact]
        public void ContentTypeUriContentTypeReadCheckTest()
        {
            var now = DateTime.Now;
            var path = Path.Combine(Path.GetTempPath(), $"{Guid.NewGuid()}.xlsx");
            var rowsWritten = MiniExcel.SaveAs(path, new[] 
            {
                new { a = @"""<>+-*//}{\\n", b = 1234567890,c = true,d= now},
                new { a = "<test>Hello World</test>", b = -1234567890,c=false,d=now.Date}
            });
            Assert.Single(rowsWritten);
            Assert.Equal(2, rowsWritten[0]);
            
            using (var zip = Package.Open(path, FileMode.Open))
            {
                var allParts = zip.GetParts()
                    .Select(s => new { s.CompressionOption, s.ContentType, s.Uri, s.Package.GetType().Name })
                    .ToDictionary(s => s.Uri.ToString(), s => s);
                
                Assert.True(allParts["/xl/styles.xml"].ContentType == "application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml");
                Assert.True(allParts["/xl/workbook.xml"].ContentType == "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml");
                Assert.True(allParts["/xl/worksheets/sheet1.xml"].ContentType == "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml");
                Assert.True(allParts["/xl/_rels/workbook.xml.rels"].ContentType == "application/vnd.openxmlformats-package.relationships+xml");
                Assert.True(allParts["/_rels/.rels"].ContentType == "application/vnd.openxmlformats-package.relationships+xml");
            }
            File.Delete(path);
        }

        [Fact]
        public void TestStirctOpenXml()
        {
            var path = "../../../../../samples/xlsx/TestStrictOpenXml.xlsx";
            var columns = MiniExcel.GetColumns(path);
            Assert.Equal(new[] { "A", "B", "C" }, columns);

            var rows = MiniExcel.Query(path).ToList();
            Assert.Equal(rows[0].A, "title1");
            Assert.Equal(rows[0].B, "title2");
            Assert.Equal(rows[0].C, "title3");
            Assert.Equal(rows[1].A, "value1");
            Assert.Equal(rows[1].B, "value2");
            Assert.Equal(rows[1].C, "value3");
        }

        [Fact]
        public void SharedStringCacheTest()
        {
            var path = "../../../../../benchmarks/MiniExcel.Benchmarks/Test1,000,000x10_SharingStrings.xlsx";
            var stopWatch = new Stopwatch();
            stopWatch.Start();
            MiniExcel.Query(path, configuration: new OpenXmlConfiguration() { EnableSharedStringCache = true }).First();
            Process currentProcess = Process.GetCurrentProcess();
            long totalBytesOfMemoryUsed = currentProcess.WorkingSet64;
            output.WriteLine("totalBytesOfMemoryUsed: " + totalBytesOfMemoryUsed);
            output.WriteLine("elapsedMilliseconds: " + stopWatch.ElapsedMilliseconds);
            stopWatch.Stop();
        }

        [Fact]
        public void SharedStringNoCacheTest()
        {
            var path = "../../../../../benchmarks/MiniExcel.Benchmarks/Test1,000,000x10_SharingStrings.xlsx";
            var stopWatch = new Stopwatch();
            stopWatch.Start();
            MiniExcel.Query(path).First();
            Process currentProcess = Process.GetCurrentProcess();
            long totalBytesOfMemoryUsed = currentProcess.WorkingSet64;
            output.WriteLine("totalBytesOfMemoryUsed: " + totalBytesOfMemoryUsed);
            output.WriteLine("elapsedMilliseconds: " + stopWatch.ElapsedMilliseconds);
            stopWatch.Stop();
        }

        [Fact]
        public void DynamicColumnsConfigurationIsUsedWhenCreatingExcelUsingIDataReader()
        {
            var path = Path.Combine(Path.GetTempPath(), $"{Guid.NewGuid()}.xlsx");
            var dateTime = DateTime.Now;
            var onlyDate = DateOnly.FromDateTime(dateTime);
            var table = new DataTable();
            {
                table.Columns.Add("Column1", typeof(string));
                table.Columns.Add("Column2", typeof(int));
                table.Columns.Add("Column3", typeof(DateTime));
                table.Columns.Add("Column4", typeof(DateOnly));
                table.Rows.Add("MiniExcel", 1, dateTime, onlyDate);
                table.Rows.Add("Github", 2, dateTime, onlyDate);
            }

            var configuration = new OpenXmlConfiguration
            {
                DynamicColumns = new[]
                {
                    new DynamicExcelColumn("Column1")
                    {
                        Name = "Name of something",
                        Index = 0,
                        Width = 150
                    },
                    new DynamicExcelColumn("Column2")
                    {
                        Name = "Its value",
                        Index = 1,
                        Width = 150
                    },
                    new DynamicExcelColumn("Column3")
                    {
                        Name = "Its Date",
                        Index = 2,
                        Width = 150,
                        Format = "dd.mm.yyyy hh:mm:ss",
                    }

                }
            };
            var reader = table.CreateDataReader();

            MiniExcel.SaveAs(path, reader, configuration: configuration);

            using (var stream = File.OpenRead(path))
            {
                var rows = stream.Query(useHeaderRow: true)
                    .Select(x => (IDictionary<string, object>)x)
                    .ToList();

                Assert.Contains("Name of something", rows[0]);
                Assert.Contains("Its value", rows[0]);
                Assert.Contains("Its Date", rows[0]);
                Assert.Contains("Column4", rows[0]);
                Assert.Contains("Name of something", rows[1]);
                Assert.Contains("Its value", rows[1]);
                Assert.Contains("Its Date", rows[1]);
                Assert.Contains("Column4", rows[1]);

                Assert.Equal("MiniExcel", rows[0]["Name of something"]);
                Assert.Equal(1D, rows[0]["Its value"]);
                Assert.Equal(dateTime, (DateTime)rows[0]["Its Date"], TimeSpan.FromMilliseconds(10d));
                Assert.Equal(onlyDate.ToDateTime(TimeOnly.MinValue), (DateTime)rows[0]["Column4"]);
                Assert.Equal("Github", rows[1]["Name of something"]);
                Assert.Equal(2D, rows[1]["Its value"]);
                Assert.Equal(dateTime, (DateTime)rows[1]["Its Date"], TimeSpan.FromMilliseconds(10d));
                Assert.Equal(onlyDate.ToDateTime(TimeOnly.MinValue), (DateTime)rows[1]["Column4"]);
            }
        }

        [Fact]
        public void DynamicColumnsConfigurationIsUsedWhenCreatingExcelUsingDataTable()
        {
            var path = Path.Combine(Path.GetTempPath(), $"{Guid.NewGuid()}.xlsx");
            var dateTime = DateTime.Now;
            var onlyDate = DateOnly.FromDateTime(dateTime);
            var table = new DataTable();
            {
                table.Columns.Add("Column1", typeof(string));
                table.Columns.Add("Column2", typeof(int));
                table.Columns.Add("Column3", typeof(DateTime));
                table.Columns.Add("Column4", typeof(DateOnly));
                table.Rows.Add("MiniExcel", 1, dateTime, onlyDate);
                table.Rows.Add("Github", 2, dateTime, onlyDate);
            }

            var configuration = new OpenXmlConfiguration
            {
                DynamicColumns = new[]
                {
                    new DynamicExcelColumn("Column1")
                    {
                        Name = "Name of something",
                        Index = 0,
                        Width = 150
                    },
                    new DynamicExcelColumn("Column2")
                    {
                        Name = "Its value",
                        Index = 1,
                        Width = 150
                    },
                    new DynamicExcelColumn("Column3")
                    {
                        Name = "Its Date",
                        Index = 2,
                        Width = 150,
                        Format = "dd.mm.yyyy hh:mm:ss"
                    }
                }
            };

            MiniExcel.SaveAs(path, table, configuration: configuration);

            using (var stream = File.OpenRead(path))
            {
                var rows = stream.Query(useHeaderRow: true)
                    .Select(x => (IDictionary<string, object>)x)
                    .Select(x => (IDictionary<string, object>)x)
                    .ToList();

                Assert.Contains("Name of something", rows[0]);
                Assert.Contains("Its value", rows[0]);
                Assert.Contains("Its Date", rows[0]);
                Assert.Contains("Column4", rows[0]);
                Assert.Contains("Name of something", rows[1]);
                Assert.Contains("Its value", rows[1]);
                Assert.Contains("Its Date", rows[1]);
                Assert.Contains("Column4", rows[1]);


                Assert.Equal("MiniExcel", rows[0]["Name of something"]);
                Assert.Equal(1D, rows[0]["Its value"]);
                Assert.Equal(dateTime, (DateTime)rows[0]["Its Date"], TimeSpan.FromMilliseconds(10d));
                Assert.Equal(onlyDate.ToDateTime(TimeOnly.MinValue), (DateTime)rows[0]["Column4"]);
                Assert.Equal("Github", rows[1]["Name of something"]);
                Assert.Equal(2D, rows[1]["Its value"]);
                Assert.Equal(dateTime, (DateTime)rows[1]["Its Date"], TimeSpan.FromMilliseconds(10d));
                Assert.Equal(onlyDate.ToDateTime(TimeOnly.MinValue), (DateTime)rows[1]["Column4"]);
            }
        }

        [Fact]
        public void InsertSheetTest()
        {
            var now = DateTime.Now;
            var path = PathHelper.GetTempPath();
            {
                var table = new DataTable();
                {
                    table.Columns.Add("a", typeof(string));
                    table.Columns.Add("b", typeof(decimal));
                    table.Columns.Add("c", typeof(bool));
                    table.Columns.Add("d", typeof(DateTime));
                    table.Rows.Add(@"""<>+-*//}{\\n", 1234567890, true, now);
                    table.Rows.Add("<test>Hello World</test>", -1234567890, false, now.Date);
                }

                var rowsWritten = MiniExcel.Insert(path, table, sheetName: "Sheet1");
                Assert.Equal(2, rowsWritten);
                
                using (var p = new ExcelPackage(new FileInfo(path)))
                {
                    var sheet1 = p.Workbook.Worksheets[0];

                    Assert.True(sheet1.Cells["A1"].Value.ToString() == "a");
                    Assert.True(sheet1.Cells["B1"].Value.ToString() == "b");
                    Assert.True(sheet1.Cells["C1"].Value.ToString() == "c");
                    Assert.True(sheet1.Cells["D1"].Value.ToString() == "d");

                    Assert.True(sheet1.Cells["A2"].Value.ToString() == @"""<>+-*//}{\\n");
                    Assert.True(sheet1.Cells["B2"].Value.ToString() == "1234567890");
                    Assert.True(sheet1.Cells["C2"].Value.ToString() == true.ToString());
                    Assert.True(sheet1.Cells["D2"].Value.ToString() == now.ToString());

                    Assert.True(sheet1.Name == "Sheet1");
                }
            }
            {
                var table = new DataTable();
                {
                    table.Columns.Add("Column1", typeof(string));
                    table.Columns.Add("Column2", typeof(int));
                    table.Rows.Add("MiniExcel", 1);
                    table.Rows.Add("Github", 2);
                }

                var rowsWritten = MiniExcel.Insert(path, table, sheetName: "Sheet2");
                Assert.Equal(2, rowsWritten);
                
                using (var p = new ExcelPackage(new FileInfo(path)))
                {
                    var sheet2 = p.Workbook.Worksheets[1];

                    Assert.True(sheet2.Cells["A1"].Value.ToString() == "Column1");
                    Assert.True(sheet2.Cells["B1"].Value.ToString() == "Column2");

                    Assert.True(sheet2.Cells["A2"].Value.ToString() == "MiniExcel");
                    Assert.True(sheet2.Cells["B2"].Value.ToString() == "1");

                    Assert.True(sheet2.Cells["A3"].Value.ToString() == "Github");
                    Assert.True(sheet2.Cells["B3"].Value.ToString() == "2");

                    Assert.True(sheet2.Name == "Sheet2");
                }
            }
            {
                var table = new DataTable();
                {
                    table.Columns.Add("Column1", typeof(string));
                    table.Columns.Add("Column2", typeof(DateTime));
                    table.Rows.Add("Test", now);
                }

                var rowsWritten = MiniExcel.Insert(path, table, sheetName: "Sheet2", printHeader: false, configuration: new OpenXmlConfiguration
                {
                    FastMode = true,
                    AutoFilter = false,
                    TableStyles = TableStyles.None,
                    DynamicColumns = new[]
                    {
                        new DynamicExcelColumn("Column2")
                        {
                            Name = "Its Date",
                            Index = 1,
                            Width = 150,
                            Format = "dd.mm.yyyy hh:mm:ss",
                        }
                    }
                }, overwriteSheet: true);
                Assert.Equal(1, rowsWritten);
                
                using (var p = new ExcelPackage(new FileInfo(path)))
                {
                    var sheet2 = p.Workbook.Worksheets[1];

                    Assert.True(sheet2.Cells["A1"].Value.ToString() == "Test");
                    Assert.True(sheet2.Cells["B1"].Text == now.ToString("dd.MM.yyyy HH:mm:ss"));
                    Assert.True(sheet2.Name == "Sheet2");
                }
            }
            {
                var table = new DataTable();
                {
                    table.Columns.Add("Column1", typeof(string));
                    table.Columns.Add("Column2", typeof(DateTime));
                    table.Rows.Add("MiniExcel", now);
                    table.Rows.Add("Github", now);
                }

                var rowsWritten = MiniExcel.Insert(path, table, sheetName: "Sheet3", configuration: new OpenXmlConfiguration
                {
                    FastMode = true,
                    AutoFilter = false,
                    TableStyles = TableStyles.None,
                    DynamicColumns =
                    [
                        new DynamicExcelColumn("Column2")
                        {
                            Name = "Its Date",
                            Index = 1,
                            Width = 150,
                            Format = "dd.mm.yyyy hh:mm:ss",
                        }
                    ]
                });
                Assert.Equal(2, rowsWritten);
                
                using (var p = new ExcelPackage(new FileInfo(path)))
                {
                    var sheet3 = p.Workbook.Worksheets[2];

                    Assert.True(sheet3.Cells["A1"].Value.ToString() == "Column1");
                    Assert.True(sheet3.Cells["B1"].Value.ToString() == "Its Date");

                    Assert.True(sheet3.Cells["A2"].Value.ToString() == "MiniExcel");
                    Assert.True(sheet3.Cells["B2"].Text == now.ToString("dd.MM.yyyy HH:mm:ss"));

                    Assert.True(sheet3.Cells["A3"].Value.ToString() == "Github");
                    Assert.True(sheet3.Cells["B3"].Text == now.ToString("dd.MM.yyyy HH:mm:ss"));

                    Assert.True(sheet3.Name == "Sheet3");
                }
            }
        }
    }
}