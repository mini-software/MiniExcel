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
using MiniExcelLibs.Utils;
using System.Threading;
using System.Data.SQLite;
using Dapper;
using System.Globalization;
using MiniExcelLibs.Tests.Utils;
using System.IO.Compression;
using System.Text;
using System.Xml;
using System.Xml.Linq;

namespace MiniExcelLibs.Tests
{
    public partial class MiniExcelHelperTests
    {
        [Fact()]
        public void CenterEmptyRowsQueryTest()
        {
            var path = @"..\..\..\..\..\samples\xlsx\TestCenterEmptyRow\TestCenterEmptyRow.xlsx";
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
            var path = @"..\..\..\..\..\samples\xlsx\TestDynamicQueryBasic_WithoutHead.xlsx";
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
            var path = @"..\..\..\..\..\samples\xlsx\TestDynamicQueryBasic.xlsx";
            using (var stream = File.OpenRead(path))
            {
                var rows = stream.Query(useHeaderRow: true).ToList();

                Assert.Equal("MiniExcel", rows[0].Column1);
                Assert.Equal(1, rows[0].Column2);
                Assert.Equal("Github", rows[1].Column1);
                Assert.Equal(2, rows[1].Column2);
            }
        }



        public class DemoPocoHelloWorld
        {
            public string HelloWorld { get; set; }
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
            var path = @"..\..\..\..\..\samples\xlsx\TestTypeMapping.xlsx";
            using (var stream = File.OpenRead(path))
            {
                var rows = stream.Query<UserAccount>().ToList();

                Assert.Equal(100,rows.Count());

                Assert.Equal(Guid.Parse("78DE23D2-DCB6-BD3D-EC67-C112BBC322A2"), rows[0].ID);
                Assert.Equal("Wade", rows[0].Name);
                Assert.Equal(DateTime.ParseExact("27/09/2020","dd/MM/yyyy", CultureInfo.InvariantCulture), rows[0].BoD);
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
            var path = @"..\..\..\..\..\samples\xlsx\TestTypeMapping_AutoCheckFormat.xlsx";
            using (var stream = FileHelper.OpenRead(path))
            {
                var rows = stream.Query<AutoCheckType>().ToList();
            }
        }

        [Fact()]
        public void TestDatetimeSpanFormat_ClosedXml()
        {
            var path = @"..\..\..\..\..\samples\xlsx\TestDatetimeSpanFormat_ClosedXml.xlsx";
            using (var stream = FileHelper.OpenRead(path))
            {
                var row = stream.QueryFirst();
                var a = row.A;
                var b = row.B;
                Assert.Equal(DateTime.Parse("2021-03-20T23:39:42.3130000"), (DateTime)a);
                Assert.Equal(TimeSpan.FromHours(10), (TimeSpan)b);
            }
        }

        [Fact()]
        public void LargeFileQueryStrongTypeMapping_Test()
        {
            var path = @"..\..\..\..\..\samples\xlsx\Test1,000,000x10\Test1,000,000x10.xlsx";
            using (var stream = File.OpenRead(path))
            {
                var rows = stream.Query<DemoPocoHelloWorld>().Take(2).ToList();

                Assert.Equal("HelloWorld", rows[0].HelloWorld);
                Assert.Equal("HelloWorld", rows[1].HelloWorld);
            }
        }



        [Theory()]
        [InlineData(@"..\..\..\..\..\samples\xlsx\ExcelDataReaderCollections\TestChess.xlsx")]
        [InlineData(@"..\..\..\..\..\samples\xlsx\TestCenterEmptyRow\TestCenterEmptyRow.xlsx")]
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
            var path = @"..\..\..\..\..\samples\xlsx\TestWihoutRAttribute.xlsx";
            using (var stream = File.OpenRead(path))
            {

            }
        }

        [Fact()]
        public void QuerySheetWithoutRAttribute()
        {
            var path = @"..\..\..\..\..\samples\xlsx\TestWihoutRAttribute.xlsx";
            using (var stream = File.OpenRead(path))
            {
                var rows = stream.Query().ToList();
                var keys = (rows.First() as IDictionary<string, object>).Keys;

                Assert.Equal(2, rows.Count());
                Assert.Equal(5, keys.Count());

                Assert.Equal(1, rows[0].A);
                //Assert.Equal(@""" <> +}{\nHello World]", (string)rows[0].B);
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
                var path = @"..\..\..\..\..\samples\xlsx\TestDimensionC3.xlsx";
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
                Assert.Equal("A1:B3", GetFirstSheetDimensionRefValue(path));
                File.Delete(path);
            }

            //Array<anoymous>
            {
                var path = Path.Combine(Path.GetTempPath(), $"{Guid.NewGuid()}.xlsx");
                var values = new []
                {
                    new {A="A",B="B"},
                    new {A="A",B="B"},
                };
                MiniExcel.SaveAs(path, values);
                Assert.Equal("A1:B3", GetFirstSheetDimensionRefValue(path));
                File.Delete(path);
            }

            // without properties
            {
                var path = Path.Combine(Path.GetTempPath(), $"{Guid.NewGuid()}.xlsx");
                var values = new List<int>();
                Assert.Throws<InvalidOperationException>(() => MiniExcel.SaveAs(path, values));
            }
        }

        [Fact()]
        public void SaveAsFileWithDimension()
        {
            {
                var path = Path.Combine(Path.GetTempPath(), $"{Guid.NewGuid()}.xlsx");
                var table = new DataTable();
                MiniExcel.SaveAs(path, table);
                Assert.Equal("A1", GetFirstSheetDimensionRefValue(path));
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
                Assert.Equal("A1:D2", GetFirstSheetDimensionRefValue(path));
                File.Delete(path);
            }

            {
                var path = Path.Combine(Path.GetTempPath(), $"{Guid.NewGuid()}.xlsx");
                var table = new DataTable();
                {
                    table.Columns.Add("a", typeof(string));
                    table.Rows.Add(@"A");
                    table.Rows.Add(@"B");
                }
                MiniExcel.SaveAs(path, table);
                Assert.Equal("A2", GetFirstSheetDimensionRefValue(path));
                File.Delete(path);
            }
        }

        private static string GetFirstSheetDimensionRefValue(string path)
        {
            string refV;
            using (var stream = File.OpenRead(path))
            using (ZipArchive archive = new ZipArchive(stream, ZipArchiveMode.Read, false, Encoding.UTF8))
            {
                var sheet = archive.Entries.Single(w => w.FullName.StartsWith("xl/worksheets/sheet1", StringComparison.OrdinalIgnoreCase));
                using (var sheetStream = sheet.Open())
                {
                    var dimension = XElement.Load(sheetStream)
                         .Descendants("dimension");
                    refV = dimension.Single().Attribute("ref").Value;
                }
            }

            return refV;
        }

        //[Theory()]
        //[InlineData(@"..\..\..\..\..\samples\xlsx\ExcelDataReaderCollections\TestOpen\TestOpen.xlsx")]
        //        public void QueryExcelDataReaderCheckTypeMappingTest(string path)
        //        {
        //#if NETCOREAPP3_1 || NET5_0
        //            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
        //#endif

        //            DataSet exceldatareaderResult;
        //            using (var stream = File.OpenRead(path))
        //            using (var reader = ExcelReaderFactory.CreateReader(stream))
        //            {
        //                exceldatareaderResult = reader.AsDataSet();
        //            }

        //            using (var stream = File.OpenRead(path))
        //            {
        //                var rows = stream.Query().ToList();
        //                Assert.Equal(exceldatareaderResult.Tables[0].Rows.Count, rows.Count);
        //                foreach (IDictionary<string, object> row in rows)
        //                {
        //                    var rowIndex = rows.IndexOf(row);
        //                    var keys = row.Keys;
        //                    foreach (var key in keys)
        //                    {
        //                        var eV = exceldatareaderResult.Tables[0].Rows[rowIndex][int.Parse(key)];
        //                        var v = row[key] == null ? DBNull.Value : row[key];
        //                        Assert.Equal(eV, v);
        //                    }
        //                }
        //            }
        //        }

        [Fact()]
        public void CreateDataTableTest()
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
        public void QueryFirstAvoidLargeFileOOMTest()
        {
            var path = @"..\..\..\..\..\samples\xlsx\Test1,000,000x10\Test1,000,000x10.xlsx";
            using (var stream = File.OpenRead(path))
                Assert.Equal("HelloWorld", stream.QueryFirst().A);
        }

        [Fact()]
        public void SQLiteInsertTest()
        {
            // Avoid SQL Insert Large Size Xlsx OOM
            var path = @"..\..\..\..\..\samples\xlsx\Test5x2.xlsx";
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
        public void BasicCreateTest()
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

            File.Delete(path);
        }

        [Fact()]
        public void BasicSaveAsStreamTest()
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
        public void SpecialAndTypeCreateTest()
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
        public void EpplusCanReadTest()
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
        public void ClosedXmlCanReadTest()
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