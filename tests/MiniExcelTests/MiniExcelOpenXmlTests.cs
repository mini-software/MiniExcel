using Xunit;
using System;
using System.Linq;
using System.IO;
using OfficeOpenXml;
using ClosedXML.Excel;
using System.IO.Packaging;
using System.Data;
using ExcelDataReader;
using System.Collections.Generic;
using System.Data.SQLite;
using Dapper;
using System.Globalization;
using static MiniExcelLibs.Tests.Utils.MiniExcelOpenXml;
using MiniExcelLibs.Tests.Utils;
using MiniExcelLibs.Attributes;

namespace MiniExcelLibs.Tests
{
    public partial class MiniExcelOpenXmlTests
    {

        [Fact]
        public void GetColumnsTest()
        {
            {
                var path = @"../../../../../samples/xlsx/TestTypeMapping.xlsx";
                var columns = MiniExcel.GetColumns(path);
                Assert.Equal(new[] { "A", "B", "C", "D", "E", "F", "G", "H" }, columns);
            }

            {
                var path = @"../../../../../samples/xlsx/TestTypeMapping.xlsx";
                var columns = MiniExcel.GetColumns(path);
                Assert.Equal(8, columns.Count);
            }

            {
                var path = @"../../../../../samples/xlsx/TestEmpty.xlsx";
                var columns = MiniExcel.GetColumns(path);
                Assert.Null(columns);
            }
        }

        [Fact]
        public void SaveAsControlChracter()
        {
            string path = GetTempXlsxPath();
            char[] chars = new char[] {'\u0000','\u0001','\u0002','\u0003','\u0004','\u0005','\u0006','\u0007','\u0008',
                '\u0009', //<HT>
	            '\u000A', //<LF>
	            '\u000B','\u000C',
                 '\u000D', //<CR>
	            '\u000E','\u000F','\u0010','\u0011','\u0012','\u0013','\u0014','\u0015','\u0016',
                 '\u0017','\u0018','\u0019','\u001A','\u001B','\u001C','\u001D','\u001E','\u001F','\u007F'
            };
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
            var path = @"../../../../../samples/xlsx/TestCustomExcelColumnAttribute.xlsx";
            Assert.Throws<System.InvalidOperationException>(() => MiniExcel.Query<CustomAttributesWihoutVaildPropertiesTestPoco>(path).ToList());
        }

        [Fact]
        public void QueryCustomAttributesTest()
        {
            var path = @"../../../../../samples/xlsx/TestCustomExcelColumnAttribute.xlsx";
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



        [Fact()]
        public void QueryCastToIDictionary()
        {
            var path = @"../../../../../samples/xlsx/TestCenterEmptyRow/TestCenterEmptyRow.xlsx";
            foreach (IDictionary<string, object> row in MiniExcel.Query(path))
            {

            }
        }
        [Fact()]
        public void CenterEmptyRowsQueryTest()
        {
            var path = @"../../../../../samples/xlsx/TestCenterEmptyRow/TestCenterEmptyRow.xlsx";
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

        [Fact()]
        public void TestDynamicQueryBasic_WithoutHead()
        {
            var path = @"../../../../../samples/xlsx/TestDynamicQueryBasic_WithoutHead.xlsx";
            using (var stream = File.OpenRead(path))
            {
                var rows = stream.Query().ToList();

                Assert.Equal("MiniExcel", rows[0].A);
                Assert.Equal(1, rows[0].B);
                Assert.Equal("Github", rows[1].A);
                Assert.Equal(2, rows[1].B);
            }
        }

        [Fact()]
        public void TestDynamicQueryBasic_useHeaderRow()
        {
            var path = @"../../../../../samples/xlsx/TestDynamicQueryBasic.xlsx";
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
            public int IgnoredProperty { get { return 1; } }
        }

        [Fact()]
        public void QueryStrongTypeMapping_Test()
        {
            var path = @"../../../../../samples/xlsx/TestTypeMapping.xlsx";
            using (var stream = File.OpenRead(path))
            {
                var rows = stream.Query<UserAccount>().ToList();

                Assert.Equal(100, rows.Count());

                Assert.Equal(Guid.Parse("78DE23D2-DCB6-BD3D-EC67-C112BBC322A2"), rows[0].ID);
                Assert.Equal("Wade", rows[0].Name);
                Assert.Equal(DateTime.ParseExact("27/09/2020", "dd/MM/yyyy", CultureInfo.InvariantCulture), rows[0].BoD);
                Assert.Equal(36, rows[0].Age);
                Assert.False(rows[0].VIP);
                Assert.Equal(decimal.Parse("5019.12"), rows[0].Points);
                Assert.Equal(1, rows[0].IgnoredProperty);
            }

            {
                var rows = MiniExcel.Query<UserAccount>(path).ToList();

                Assert.Equal(100, rows.Count());

                Assert.Equal(Guid.Parse("78DE23D2-DCB6-BD3D-EC67-C112BBC322A2"), rows[0].ID);
                Assert.Equal("Wade", rows[0].Name);
                Assert.Equal(DateTime.ParseExact("27/09/2020", "dd/MM/yyyy", CultureInfo.InvariantCulture), rows[0].BoD);
                Assert.Equal(36, rows[0].Age);
                Assert.False(rows[0].VIP);
                Assert.Equal(decimal.Parse("5019.12"), rows[0].Points);
                Assert.Equal(1, rows[0].IgnoredProperty);
            }
        }


        public class AutoCheckType
        {
            public Guid? @guid { get; set; }
            public bool? @bool { get; set; }
            public DateTime? datetime { get; set; }
            public string @string { get; set; }
        }

        [Fact()]
        public void AutoCheckTypeTest()
        {
            var path = @"../../../../../samples/xlsx/TestTypeMapping_AutoCheckFormat.xlsx";
            using (var stream = FileHelper.OpenRead(path))
            {
                var rows = stream.Query<AutoCheckType>().ToList();
            }
        }

        [Fact()]
        public void TestDatetimeSpanFormat_ClosedXml()
        {
            var path = @"../../../../../samples/xlsx/TestDatetimeSpanFormat_ClosedXml.xlsx";
            using (var stream = FileHelper.OpenRead(path))
            {
                var row = stream.Query().First();
                var a = row.A;
                var b = row.B;
                Assert.Equal(DateTime.Parse("2021-03-20T23:39:42.3130000"), (DateTime)a);
                Assert.Equal(TimeSpan.FromHours(10), (TimeSpan)b);
            }
        }

        [Fact()]
        public void LargeFileQueryStrongTypeMapping_Test()
        {
            var path = @"../../../../../benchmarks/MiniExcel.Benchmarks/Test1,000,000x10.xlsx";
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

        [Theory()]
        [InlineData(@"../../../../../samples/xlsx/ExcelDataReaderCollections/TestChess.xlsx")]
        [InlineData(@"../../../../../samples/xlsx/TestCenterEmptyRow/TestCenterEmptyRow.xlsx")]
        public void QueryExcelDataReaderCheckTest(string path)
        {
#if NETCOREAPP3_1 || NET5_0
            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
#endif

            DataSet exceldatareaderResult;
            using (var stream = File.OpenRead(path))
            using (var reader = ExcelDataReader.ExcelReaderFactory.CreateReader(stream))
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
                        var eV = exceldatareaderResult.Tables[0].Rows[rowIndex][MiniExcelLibs.Tests.Utils.Helpers.GetColumnIndex(key)];
                        var v = row[key] == null ? DBNull.Value : row[key];
                        Assert.Equal(eV, v);
                    }
                }
            }
        }

        [Fact()]
        public void QueryCustomStyle()
        {
            var path = @"../../../../../samples/xlsx/TestWihoutRAttribute.xlsx";
            using (var stream = File.OpenRead(path))
            {

            }
        }

        [Fact()]
        public void QuerySheetWithoutRAttribute()
        {
            var path = @"../../../../../samples/xlsx/TestWihoutRAttribute.xlsx";
            using (var stream = File.OpenRead(path))
            {
                var rows = stream.Query().ToList();
                var keys = (rows.First() as IDictionary<string, object>).Keys;

                Assert.Equal(2, rows.Count());
                Assert.Equal(5, keys.Count());

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

        [Fact()]
        public void FixDimensionJustOneColumnParsingError_Test()
        {
            {
                var path = @"../../../../../samples/xlsx/TestDimensionC3.xlsx";
                using (var stream = File.OpenRead(path))
                {
                    var rows = stream.Query().ToList();
                    var keys = (rows.First() as IDictionary<string, object>).Keys;
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
        [Fact()]
        public void SaveAsFileWithDimensionByICollection()
        {
            //List<strongtype>
            {
                var path = Path.Combine(Path.GetTempPath(), $"{Guid.NewGuid()}.xlsx");
                var values = new List<SaveAsFileWithDimensionByICollectionTestType>()
                {
                    new SaveAsFileWithDimensionByICollectionTestType{A="A",B="B"},
                    new SaveAsFileWithDimensionByICollectionTestType{A="A",B="B"},
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
                    Assert.Equal("A1", Helpers.GetFirstSheetDimensionRefValue(path));
                }
                File.Delete(path);


                MiniExcel.SaveAs(path, values, true);
                {
                    using (var stream = File.OpenRead(path))
                    {
                        var rows = stream.Query(useHeaderRow: false).ToList();
                        Assert.Empty(rows);
                    }
                }
                Assert.Equal("A1", Helpers.GetFirstSheetDimensionRefValue(path));
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
                MiniExcel.SaveAs(path, values);
                using (var stream = File.OpenRead(path))
                {
                    var rows = stream.Query(useHeaderRow: true).ToList();
                    Assert.Empty(rows);
                }
                File.Delete(path);
                //Assert.Throws<InvalidOperationException>(() => MiniExcel.SaveAs(path, values));
            }
        }

        [Fact()]
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
                    table.Rows.Add(@"<test>Hello World</test>", -1234567890, false, DateTime.Now);
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
                    table.Rows.Add(@"A");
                    table.Rows.Add(@"B");
                }
                MiniExcel.SaveAs(path, table);
                Assert.Equal("A3", Helpers.GetFirstSheetDimensionRefValue(path));
                File.Delete(path);
            }
        }

        [Fact()]
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
                    table.Rows.Add(@"<test>Hello World</test>", -1234567890, false, now.Date);
                }

                MiniExcel.SaveAs(path, table);

                using (var p = new ExcelPackage(new FileInfo(path)))
                {
                    var ws = p.Workbook.Worksheets.First();

                    Assert.True(ws.Cells["A1"].Value.ToString() == "a");
                    Assert.True(ws.Cells["B1"].Value.ToString() == "b");
                    Assert.True(ws.Cells["C1"].Value.ToString() == "c");
                    Assert.True(ws.Cells["D1"].Value.ToString() == "d");

                    Assert.True(ws.Cells["A2"].Value.ToString() == @"""<>+-*//}{\\n");
                    Assert.True(ws.Cells["B2"].Value.ToString() == @"1234567890");
                    Assert.True(ws.Cells["C2"].Value.ToString() == true.ToString());
                    Assert.True(ws.Cells["D2"].Value.ToString() == now.ToString());
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
            }
        }

        [Fact()]
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
                    var rows = connection.Query(@"with cte as (select 1 id,2 val) select * from cte where 1=2");
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
                    new Dictionary<string, object>(){ { "Column1","MiniExcel"},{ "Column2", 1} },
                     new Dictionary<string, object>(){ { "Column1", "Github" },{ "Column2", 2} },
                };
                MiniExcel.SaveAs(path, values);

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

                    Assert.Equal(2, rows.Count);
                    Assert.Equal("MiniExcel", rows[0].Column1);
                    Assert.Equal(1, rows[0].Column2);
                    Assert.Equal("Github", rows[1].Column1);
                    Assert.Equal(2, rows[1].Column2);
                }

                Assert.Equal("A1:B3", Helpers.GetFirstSheetDimensionRefValue(path));
                File.Delete(path);
            }

            {
                var values = new List<Dictionary<int, object>>()
                {
                    new Dictionary<int, object>(){ { 1,"MiniExcel"},{ 2, 1} },
                     new Dictionary<int, object>(){ { 1, "Github" },{ 2, 2} },
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

        [Fact()]
        public void SaveAsByDapperRows()
        {
            var path = Path.Combine(Path.GetTempPath(), $"{Guid.NewGuid()}.xlsx");


            // Dapper Query
            using (var connection = Db.GetConnection("Data Source=:memory:"))
            {
                var rows = connection.Query(@"select 'MiniExcel' as Column1,1 as Column2 union all select 'Github',2");
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
                var rows = connection.Query(@"with cte as (select 'MiniExcel' as Column1,1 as Column2 union all select 'Github',2)select * from cte where 1=2").ToList();
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
                var rows = connection.Query(@"select 'MiniExcel' as Column1,1 as Column2 union all select 'Github',2").ToList();
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
        [Fact()]
        public void QueryByStrongTypeParameterTest()
        {
            var path = Path.Combine(Path.GetTempPath(), $"{Guid.NewGuid()}.xlsx");

            var values = new List<Demo>()
            {
                new Demo { Column1= "MiniExcel" ,Column2 = 1 },
                new Demo { Column1 = "Github", Column2 = 2 }
            };
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

        [Fact()]
        public void QueryByDictionaryStringAndObjectParameterTest()
        {
            var path = Path.Combine(Path.GetTempPath(), $"{Guid.NewGuid()}.xlsx");

            var values = new List<Dictionary<string, object>>()
            {
                new Dictionary<string,object>{{ "Column1", "MiniExcel" }, { "Column2", 1 } },
                new Dictionary<string,object>{{ "Column1", "Github" }, { "Column2", 2 } }
            };
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

        [Fact()]
        public void SQLiteInsertTest()
        {
            // Avoid SQL Insert Large Size Xlsx OOM
            var path = @"../../../../../samples/xlsx/Test5x2.xlsx";
            var tempSqlitePath = Path.Combine(Path.GetTempPath(), $"{Guid.NewGuid()}.db");
            var connectionString = $"Data Source={tempSqlitePath};Version=3;";

            using (var connection = new SQLiteConnection(connectionString))
            {
                connection.Execute(@"create table T (A varchar(20),B varchar(20));");
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

        [Fact()]
        public void SaveAsBasicCreateTest()
        {
            var path = Path.Combine(Path.GetTempPath(), $"{Guid.NewGuid()}.xlsx");
            MiniExcel.SaveAs(path, new[] {
                  new { Column1 = "MiniExcel", Column2 = 1 },
                  new { Column1 = "Github", Column2 = 2}
            });

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

        [Fact()]
        public void SaveAsBasicStreamTest()
        {
            {
                var path = Path.Combine(Path.GetTempPath(), $"{Guid.NewGuid()}.xlsx");
                var values = new[] {
                      new { Column1 = "MiniExcel", Column2 = 1 },
                      new { Column1 = "Github", Column2 = 2}
                };
                using (var stream = new FileStream(path, FileMode.CreateNew))
                {
                    stream.SaveAs(values);
                }

                using (var stream = File.OpenRead(path))
                {
                    var rows = stream.Query(useHeaderRow: true).ToList();

                    Assert.Equal("MiniExcel", rows[0].Column1);
                    Assert.Equal(1, rows[0].Column2);
                    Assert.Equal("Github", rows[1].Column1);
                    Assert.Equal(2, rows[1].Column2);
                };

                File.Delete(path);
            }
            {
                var path = Path.Combine(Path.GetTempPath(), $"{Guid.NewGuid()}.xlsx");
                var values = new[] {
                      new { Column1 = "MiniExcel", Column2 = 1 },
                      new { Column1 = "Github", Column2 = 2}
                };
                using (var stream = new MemoryStream())
                using (var fileStream = new FileStream(path, FileMode.Create))
                {
                    stream.SaveAs(values);
                    stream.Seek(0, SeekOrigin.Begin);
                    stream.CopyTo(fileStream);
                }

                using (var stream = File.OpenRead(path))
                {
                    var rows = stream.Query(useHeaderRow: true).ToList();

                    Assert.Equal("MiniExcel", rows[0].Column1);
                    Assert.Equal(1, rows[0].Column2);
                    Assert.Equal("Github", rows[1].Column1);
                    Assert.Equal(2, rows[1].Column2);
                };

                File.Delete(path);
            }
        }

        [Fact()]
        public void SaveAsSpecialAndTypeCreateTest()
        {
            var path = Path.Combine(Path.GetTempPath(), $"{Guid.NewGuid()}.xlsx");
            MiniExcel.SaveAs(path, new[] {
                  new { a = @"""<>+-*//}{\\n", b = 1234567890,c = true,d=DateTime.Now },
                  new { a = "<test>Hello World</test>", b = -1234567890,c=false,d=DateTime.Now.Date}
             });
            var info = new FileInfo(path);

            Assert.True(info.FullName == path);

            File.Delete(path);
        }

        [Fact()]
        public void SaveAsFileEpplusCanReadTest()
        {
            var now = DateTime.Now;
            var path = Path.Combine(Path.GetTempPath(), $"{Guid.NewGuid()}.xlsx");
            MiniExcel.SaveAs(path, new[] {
                  new { a = @"""<>+-*//}{\\n", b = 1234567890,c = true,d= now},
                  new { a = "<test>Hello World</test>", b = -1234567890,c=false,d=now.Date}
             });
            using (var p = new ExcelPackage(new FileInfo(path)))
            {
                var ws = p.Workbook.Worksheets.First();

                Assert.True(ws.Cells["A1"].Value.ToString() == "a");
                Assert.True(ws.Cells["B1"].Value.ToString() == "b");
                Assert.True(ws.Cells["C1"].Value.ToString() == "c");
                Assert.True(ws.Cells["D1"].Value.ToString() == "d");

                Assert.True(ws.Cells["A2"].Value.ToString() == @"""<>+-*//}{\\n");
                Assert.True(ws.Cells["B2"].Value.ToString() == @"1234567890");
                Assert.True(ws.Cells["C2"].Value.ToString() == true.ToString());
                Assert.True(ws.Cells["D2"].Value.ToString() == now.ToString());
            }
            File.Delete(path);
        }

        [Fact()]
        public void SavaAsClosedXmlCanReadTest()
        {
            var now = DateTime.Now;
            var path = Path.Combine(Path.GetTempPath(), $"{Guid.NewGuid()}.xlsx");
            MiniExcel.SaveAs(path, new[] {
                  new { a = @"""<>+-*//}{\\n", b = 1234567890,c = true,d= now},
                  new { a = "<test>Hello World</test>", b = -1234567890,c=false,d=now.Date}
             });
            using (var workbook = new XLWorkbook(path))
            {
                var ws = workbook.Worksheets.First();

                Assert.True(ws.Cell("A1").Value.ToString() == "a");
                Assert.True(ws.Cell("D1").Value.ToString() == "d");
                Assert.True(ws.Cell("B1").Value.ToString() == "b");
                Assert.True(ws.Cell("C1").Value.ToString() == "c");

                Assert.True(ws.Cell("A2").Value.ToString() == @"""<>+-*//}{\\n");
                Assert.True(ws.Cell("B2").Value.ToString() == @"1234567890");
                Assert.True(ws.Cell("C2").Value.ToString() == true.ToString());
                Assert.True(ws.Cell("D2").Value.ToString() == now.ToString());
            }
            File.Delete(path);
        }

        [Fact()]
        public void ContentTypeUriContentTypeReadCheckTest()
        {
            var now = DateTime.Now;
            var path = Path.Combine(Path.GetTempPath(), $"{Guid.NewGuid()}.xlsx");
            MiniExcel.SaveAs(path, new[] {
                  new { a = @"""<>+-*//}{\\n", b = 1234567890,c = true,d= now},
                  new { a = "<test>Hello World</test>", b = -1234567890,c=false,d=now.Date}
             });
            using (Package zip = System.IO.Packaging.Package.Open(path, FileMode.Open))
            {
                var allParts = zip.GetParts().Select(s => new { s.CompressionOption, s.ContentType, s.Uri, s.Package.GetType().Name })
                    .ToDictionary(s => s.Uri.ToString(), s => s)
                    ;
                Assert.True(allParts[@"/xl/styles.xml"].ContentType == "application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml");
                Assert.True(allParts[@"/xl/workbook.xml"].ContentType == "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml");
                Assert.True(allParts[@"/xl/worksheets/sheet1.xml"].ContentType == "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml");
                Assert.True(allParts[@"/xl/_rels/workbook.xml.rels"].ContentType == "application/vnd.openxmlformats-package.relationships+xml");
                Assert.True(allParts[@"/_rels/.rels"].ContentType == "application/vnd.openxmlformats-package.relationships+xml");
            }
            File.Delete(path);
        }
    }
}