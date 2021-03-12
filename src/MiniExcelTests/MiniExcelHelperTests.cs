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
using System.Dynamic;

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
                foreach (var item in rows)
                {
                    
                }
            }
        }

        [Fact()]
        public void QueryExcelDataReaderCheckTest()
        {
#if NETCOREAPP3_1 || NET5_0
            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
#endif
            var path = @"..\..\..\..\..\samples\xlsx\TestCenterEmptyRow\TestCenterEmptyRow.xlsx";

            DataSet exceldatareaderResult;
            using (var stream = File.OpenRead(path))
            using (var reader = ExcelReaderFactory.CreateReader(stream))
            {
                exceldatareaderResult = reader.AsDataSet();
            }

            using (var stream = File.OpenRead(path))
            {
                var rows = stream.Query().ToList();
                foreach (IDictionary<string, object> row in rows)
                {
                    var rowIndex = rows.IndexOf(row);
                    var keys = row.Keys;
                    foreach (var key in keys)
                    {
                        var eV = exceldatareaderResult.Tables[0].Rows[rowIndex][int.Parse(key)];
                        var v = row[key];
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