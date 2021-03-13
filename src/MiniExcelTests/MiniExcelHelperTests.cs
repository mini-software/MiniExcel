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

namespace MiniExcelLibs.Tests
{
    public class MiniExcelHelperTests
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
                var rows = stream.Query(useHeaderRow:true).ToList();

                Assert.Equal("MiniExcel", rows[0].Column1);
                Assert.Equal(1, rows[0].Column2);
                Assert.Equal("Github", rows[1].Column1);
                Assert.Equal(2, rows[1].Column2);
            }
        }

        //TODO:
        //[Fact()]
        public void QueryAvoidOOMSqlInsertTest()
        {
            var path = @"..\..\..\..\..\samples\xlsx\TestCenterEmptyRow\TestCenterEmptyRow.xlsx";
            using (var stream = File.OpenRead(path))
            {
                var rows = stream.Query();
                foreach (var item in rows)
                {

                }
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
                Assert.Equal(exceldatareaderResult.Tables[0].Rows.Count , rows.Count);
                foreach (IDictionary<string, object> row in rows)
                {
                    var rowIndex = rows.IndexOf(row);
                    var keys = row.Keys;
                    foreach (var key in keys)
                    {
                        var eV = exceldatareaderResult.Tables[0].Rows[rowIndex][MiniExcelLibs.Utils.Helpers.GetColumnIndex(key)];
                        var v = row[key]==null?DBNull.Value:row[key];
                        Assert.Equal(eV, v);
                    }
                }
            }
        }

        //[Theory()]
        //[InlineData(@"..\..\..\..\..\samples\xlsx\ExcelDataReaderCollections\TestOpen\TestOpen.xlsx")]
        public void QueryExcelDataReaderCheckTypeMappingTest(string path)
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
                        var eV = exceldatareaderResult.Tables[0].Rows[rowIndex][int.Parse(key)];
                        var v = row[key] == null ? DBNull.Value : row[key];
                        Assert.Equal(eV, v);
                    }
                }
            }
        }

        [Fact()]
        public void CreateDataTableTest()
        {
            var now = DateTime.Now;
            var path = Path.Combine(Path.GetTempPath(), $"{Guid.NewGuid()}.xlsx");
            var table = new DataTable();
            {
                table.Columns.Add("a", typeof(string));
                table.Columns.Add("b", typeof(decimal));
                table.Columns.Add("c", typeof(bool));
                table.Columns.Add("d", typeof(DateTime));
                table.Rows.Add(@"""<>+-*//}{\\n", 1234567890,true, now);
                table.Rows.Add(@"<test>Hello World</test>", -1234567890,false, now.Date);
            }

            MiniExcel.Create(path, table);

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
        public void CreateTest()
        {
            var path = Path.Combine(Path.GetTempPath(), $"{Guid.NewGuid()}.xlsx");
            MiniExcel.Create(path, new[] {
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
            MiniExcel.Create(path, new[] {
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
            MiniExcel.Create(path, new[] {
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
            MiniExcel.Create(path, new[] {
                  new { a = @"""<>+-*//}{\\n", b = 1234567890,c = true,d= now},
                  new { a = "<test>Hello World</test>", b = -1234567890,c=false,d=now.Date}
             });
            using (Package zip = System.IO.Packaging.Package.Open(path, FileMode.Open))
            {
                var allParts = zip.GetParts().Select(s => new { s.CompressionOption, s.ContentType, s.Uri, s.Package.GetType().Name })
                    .ToDictionary(s=>s.Uri.ToString(),s=>s)
                    ;
                Assert.True(allParts[@"/xl/styles.xml"].ContentType == "application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml") ;
                Assert.True(allParts[@"/xl/workbook.xml"].ContentType == "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml");
                Assert.True(allParts[@"/xl/worksheets/sheet1.xml"].ContentType == "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml");
                Assert.True(allParts[@"/xl/_rels/workbook.xml.rels"].ContentType == "application/vnd.openxmlformats-package.relationships+xml");
                Assert.True(allParts[@"/_rels/.rels"].ContentType == "application/vnd.openxmlformats-package.relationships+xml");
            }
            File.Delete(path);
        }
    }
}