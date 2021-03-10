using Xunit;
using MiniExcelLibs;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using OfficeOpenXml;
using ClosedXML.Excel;
using System.IO.Packaging;
using System.Data;

namespace MiniExcelLibs.Tests
{
    public class MiniExcelHelperTests
    {
        [Fact()]
        public void QueryTest()
        {
            var path = @"..\..\..\..\..\samples\xlsx\TestCenterEmptyRow\TestCenterEmptyRow.xlsx";
            using (var stream = File.OpenRead(path))
            {
                var rows = stream.Query();

                Assert.Equal("a", rows[0][0]);
                Assert.Equal("b", rows[0][1]);
                Assert.Equal("c", rows[0][2]);
                Assert.Equal("d", rows[0][3]);

                Assert.Equal("1", rows[1][0]);
                Assert.Null(rows[1][1]);
                Assert.Equal("3", rows[1][2]);
                Assert.Null(rows[1][3]);

                Assert.Null(rows[2][0]);
                Assert.Equal("2", rows[2][1]);
                Assert.Null(rows[2][2]);
                Assert.Equal("4", rows[2][3]);

                Assert.Null(rows[3][0]);
                Assert.Null(rows[3][1]);
                Assert.Null(rows[3][2]);
                Assert.Null(rows[3][3]);
            }

            {
                var rows = MiniExcel.Query(path);

                Assert.Equal("a", rows[0][0]);
                Assert.Equal("b", rows[0][1]);
                Assert.Equal("c", rows[0][2]);
                Assert.Equal("d", rows[0][3]);

                Assert.Equal("1", rows[1][0]);
                Assert.Null(rows[1][1]);
                Assert.Equal("3", rows[1][2]);
                Assert.Null(rows[1][3]);

                Assert.Null(rows[2][0]);
                Assert.Equal("2", rows[2][1]);
                Assert.Null(rows[2][2]);
                Assert.Equal("4", rows[2][3]);

                Assert.Null(rows[3][0]);
                Assert.Null(rows[3][1]);
                Assert.Null(rows[3][2]);
                Assert.Null(rows[3][3]);
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