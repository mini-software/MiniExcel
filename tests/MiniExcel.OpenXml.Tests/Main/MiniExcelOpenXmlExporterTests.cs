using ClosedXML.Excel;
using MiniExcelLib.OpenXml.Tests.Utils;
using MiniExcelLib.Tests.Common;
using MiniExcelLib.Tests.Common.Utils;

namespace MiniExcelLib.OpenXml.Tests.Main;

public class MiniExcelOpenXmlExporterTests(ITestOutputHelper output)
{
    private readonly ITestOutputHelper _output = output;
    
    private readonly OpenXmlImporter _excelImporter =  MiniExcel.Importers.GetOpenXmlImporter();
    private readonly OpenXmlExporter _excelExporter =  MiniExcel.Exporters.GetOpenXmlExporter();
   
    static MiniExcelOpenXmlExporterTests()
    {
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
    }
    
    [Fact]
    public void SaveAsControlChracter()
    {
        using var path = AutoDeletingPath.Create();
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
         _excelExporter.Export(path.ToString(), input);

        var rows2 = _excelImporter.Query(path.ToString(), true).Select(s => s.Test).ToArray();
        var rows1 = _excelImporter.Query<SaveAsControlChracterVO>(path.ToString()).Select(s => s.Test).ToArray();
    }

    [Fact]
    public void SaveAsCustomAttributesTest()
    {
        using var path = AutoDeletingPath.Create();
        var input = Enumerable.Range(1, 3)
            .Select(_ => new ExcelAttributeDemo
            {
                Test1 = "Test1",
                Test2 = "Test2",
                Test3 = "Test3",
                Test4 = "Test4",
            });

         _excelExporter.Export(path.ToString(), input);
        var rows = _excelImporter.Query(path.ToString(), true).ToList();
        var first = rows[0] as IDictionary<string, object>;

        Assert.Equal(3, rows.Count);
        Assert.Equal(["Column1", "Column2", "Test5", "Test7", "Test6", "Test4"], first?.Keys);
        Assert.Equal("Test1", rows[0].Column1);
        Assert.Equal("Test2", rows[0].Column2);
        Assert.Equal("Test4", rows[0].Test4);
        Assert.Null(rows[0].Test5);
        Assert.Null(rows[0].Test6);
    }

    [Fact]
    public void SaveAsFileWithDimensionByICollection()
    {
        //List<strongtype>
        {
            using var file = AutoDeletingPath.Create();
            var path = file.ToString();

            List<SaveAsFileWithDimensionByICollectionTestType> values =
            [
                new() { A = "A", B = "B" },
                new() { A = "A", B = "B" }
            ];
             _excelExporter.Export(path, values);

            using (var stream = File.OpenRead(path))
            {
                var rows = _excelImporter.Query(stream, hasHeaderRow: false).ToList();
                Assert.Equal(3, rows.Count);
                Assert.Equal("A", rows[0].A);
                Assert.Equal("A", rows[1].A);
                Assert.Equal("A", rows[2].A);
            }
            using (var stream = File.OpenRead(path))
            {
                var rows = _excelImporter.Query(stream, hasHeaderRow: true).ToList();
                Assert.Equal(2, rows.Count);
                Assert.Equal("A", rows[0].A);
                Assert.Equal("A", rows[1].A);
            }
            Assert.Equal("A1:B3", SheetHelper.GetFirstSheetDimensionRefValue(path));

             _excelExporter.Export(path, values, false, overwriteFile: true);
            Assert.Equal("A1:B2", SheetHelper.GetFirstSheetDimensionRefValue(path));
        }

        //List<strongtype> empty
        {
            using var file = AutoDeletingPath.Create();
            var path = file.ToString();
            List<SaveAsFileWithDimensionByICollectionTestType> values = [];

             _excelExporter.Export(path, values, false);
            {
                using (var stream = File.OpenRead(path))
                {
                    var rows = _excelImporter.Query(stream, hasHeaderRow: false).ToList();
                    Assert.Empty(rows);
                }
                Assert.Equal("A1:B1", SheetHelper.GetFirstSheetDimensionRefValue(path));
            }

            _excelExporter.Export(path, values, overwriteFile: true);
            {
                using var stream = File.OpenRead(path);
                var rows = _excelImporter.Query(stream, hasHeaderRow: false).ToList();
                Assert.Single(rows);
            }
            Assert.Equal("A1:B1", SheetHelper.GetFirstSheetDimensionRefValue(path));
        }

        //Array<anoymous>
        {
            using var file = AutoDeletingPath.Create();
            var path = file.ToString();
            var values = new[]
            {
                new {A="A",B="B"},
                new {A="A",B="B"},
            };
             _excelExporter.Export(path, values);
            {
                using (var stream = File.OpenRead(path))
                {
                    var rows = _excelImporter.Query(stream, hasHeaderRow: false).ToList();
                    Assert.Equal(3, rows.Count);
                    Assert.Equal("A", rows[0].A);
                    Assert.Equal("A", rows[1].A);
                    Assert.Equal("A", rows[2].A);
                }
                using (var stream = File.OpenRead(path))
                {
                    var rows = _excelImporter.Query(stream, hasHeaderRow: true).ToList();
                    Assert.Equal(2, rows.Count);
                    Assert.Equal("A", rows[0].A);
                    Assert.Equal("A", rows[1].A);
                }
            }
            Assert.Equal("A1:B3", SheetHelper.GetFirstSheetDimensionRefValue(path));

             _excelExporter.Export(path, values, false, overwriteFile: true);
            Assert.Equal("A1:B2", SheetHelper.GetFirstSheetDimensionRefValue(path));
        }

        // without properties
        {
            using var path = AutoDeletingPath.Create();
            var values = new List<int>();
            Assert.Throws<NotSupportedException>(() =>  _excelExporter.Export(path.ToString(), values));
        }
    }

    [Fact]
    public void SaveAsFileWithDimension()
    {
        {
            using var file = AutoDeletingPath.Create();
            var path = file.ToString();

            var table = new DataTable();
            _excelExporter.Export(path, table);
            Assert.Equal("A1", SheetHelper.GetFirstSheetDimensionRefValue(path));
            
            var rows = _excelImporter.Query(path).ToList();
            Assert.Empty(rows);

            _excelExporter.Export(path, table, printHeader: false, overwriteFile: true);
            Assert.Equal("A1", SheetHelper.GetFirstSheetDimensionRefValue(path));
        }

        {
            using var file = AutoDeletingPath.Create();
            var path = file.ToString();

            var table = new DataTable();
            table.Columns.Add("a", typeof(string));
            table.Columns.Add("b", typeof(decimal));
            table.Columns.Add("c", typeof(bool));
            table.Columns.Add("d", typeof(DateTime));
            table.Rows.Add(@"""<>+-*//}{\\n", 1234567890);
            table.Rows.Add("<test>Hello World</test>", -1234567890, false, DateTime.Now);

            _excelExporter.Export(path, table);
            Assert.Equal("A1:D3", SheetHelper.GetFirstSheetDimensionRefValue(path));

            var rowsWithHeader = _excelImporter.Query(path, hasHeaderRow: true).ToList();
            Assert.Equal(2, rowsWithHeader.Count);
            Assert.Equal(@"""<>+-*//}{\\n", rowsWithHeader[0].a);
            Assert.Equal(1234567890, rowsWithHeader[0].b);
            Assert.Null(rowsWithHeader[0].c);
            Assert.Null(rowsWithHeader[0].d);

            var rowsNoHeader = _excelImporter.Query(path).ToList();
            Assert.Equal(3, rowsNoHeader.Count);
            Assert.Equal("a", rowsNoHeader[0].A);
            Assert.Equal("b", rowsNoHeader[0].B);
            Assert.Equal("c", rowsNoHeader[0].C);
            Assert.Equal("d", rowsNoHeader[0].D);

            _excelExporter.Export(path, table, printHeader: false, overwriteFile: true);
            Assert.Equal("A1:D2", SheetHelper.GetFirstSheetDimensionRefValue(path));
        }

        //TODO:StartCell
        {
            using var path = AutoDeletingPath.Create();

            var table = new DataTable();
            table.Columns.Add("a", typeof(string));
            table.Rows.Add("A");
            table.Rows.Add("B");

             _excelExporter.Export(path.ToString(), table);
            Assert.Equal("A1:A3", SheetHelper.GetFirstSheetDimensionRefValue(path.ToString()));
        }
    }

    [Fact]
    public void SaveAsByDataTableTest()
    {
        {
            var now = new DateTime(2026, 6, 2, 15, 1, 47);
            using var file = AutoDeletingPath.Create();
            var path = file.ToString();

            var table = new DataTable();
            table.Columns.Add("a", typeof(string));
            table.Columns.Add("b", typeof(decimal));
            table.Columns.Add("c", typeof(bool));
            table.Columns.Add("d", typeof(DateTime));
            table.Rows.Add(@"""<>+-*//}{\\n", 1234567890, true, now);
            table.Rows.Add("<test>Hello World</test>", -1234567890, false, now.Date);

             _excelExporter.Export(path, table, sheetName: "R&D");

            using var p = new ExcelPackage(new FileInfo(path));
            var ws = p.Workbook.Worksheets.First();

            Assert.Equal("a", ws.Cells["A1"].Value.ToString());
            Assert.Equal("b", ws.Cells["B1"].Value.ToString());
            Assert.Equal("c", ws.Cells["C1"].Value.ToString());
            Assert.Equal("d", ws.Cells["D1"].Value.ToString());

            Assert.Equal(@"""<>+-*//}{\\n", ws.Cells["A2"].Value.ToString());
            Assert.Equal("1234567890", ws.Cells["B2"].Value.ToString());
            Assert.True(ws.Cells["C2"].GetCellValue<bool>());
            Assert.Equal(now, ws.Cells["D2"].GetCellValue<DateTime>());

            Assert.Equal("R&D", ws.Name);
        }
        {
            using var path = AutoDeletingPath.Create();
            var table = new DataTable();
            table.Columns.Add("Column1", typeof(string));
            table.Columns.Add("Column2", typeof(int));
            table.Rows.Add("MiniExcel", 1);
            table.Rows.Add("Github", 2);

             _excelExporter.Export(path.ToString(), table);
        }
    }

    [Fact]
    public void EmptyTest()
    {
        using var path = AutoDeletingPath.Create();
        using (var connection = Db.GetConnection("Data Source=:memory:"))
        {
            var rows = connection.Query("with cte as (select 1 id,2 val) select * from cte where 1=2");
             _excelExporter.Export(path.ToString(), rows);
        }
        using (var stream = File.OpenRead(path.ToString()))
        {
            var rows = _excelImporter.Query(stream, hasHeaderRow: true).ToList();
            Assert.Empty(rows);
        }
    }

    [Fact]
    public void SaveAsByIEnumerableIDictionary()
    {
        using var file = AutoDeletingPath.Create();
        var path = file.ToString();

        {
            var values = new List<Dictionary<string, object>>
            {
                new() { { "Column1", "MiniExcel" }, { "Column2", 1 } },
                new() { { "Column1", "Github" }, { "Column2", 2 } }
            };
            var sheets = new Dictionary<string, object>
            {
                ["R&D"] = values,
                ["success!"] = values
            };
             _excelExporter.Export(path, sheets);

            using (var stream = File.OpenRead(path))
            {
                var rows = _excelImporter.Query(stream, hasHeaderRow: false, leaveOpen: true).ToList();
                Assert.Equal("Column1", rows[0].A);
                Assert.Equal("Column2", rows[0].B);
                Assert.Equal("MiniExcel", rows[1].A);
                Assert.Equal(1, rows[1].B);
                Assert.Equal("Github", rows[2].A);
                Assert.Equal(2, rows[2].B);

                Assert.Equal("R&D", _excelImporter.GetSheetNames(stream)[0]);
            }

            using (var stream = File.OpenRead(path))
            {
                var rows = _excelImporter.Query(stream, hasHeaderRow: true, leaveOpen: true).ToList();

                Assert.Equal(2, rows.Count);
                Assert.Equal("MiniExcel", rows[0].Column1);
                Assert.Equal(1, rows[0].Column2);
                Assert.Equal("Github", rows[1].Column1);
                Assert.Equal(2, rows[1].Column2);

                Assert.Equal("success!", _excelImporter.GetSheetNames(stream)[1]);
            }

            Assert.Equal("A1:B3", SheetHelper.GetFirstSheetDimensionRefValue(path));
        }

        {
            var values = new List<Dictionary<int, object>>
            {
                new() { { 1, "MiniExcel"}, { 2, 1 } },
                new() { { 1, "Github" }, { 2, 2 } },
            };
             _excelExporter.Export(path, values, overwriteFile: true);

            using (var stream = File.OpenRead(path))
            {
                var rows = _excelImporter.Query(stream, hasHeaderRow: false).ToList();
                Assert.Equal(3, rows.Count);
            }

            Assert.Equal("A1:B3", SheetHelper.GetFirstSheetDimensionRefValue(path));
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

        // Test enumerable
        using var path = AutoDeletingPath.Create();
        _excelExporter.Export(
            path.ToString(),
            new[]
            {
                new { Column1 = "MiniExcel", Column2 = 1 },
                new { Column1 = "Github", Column2 = 2 }
            },
            configuration: config
        );

        using (var stream = File.OpenRead(path.ToString()))
        {
            var rows = _excelImporter.Query(stream, hasHeaderRow: true).ToList();

            Assert.Equal("MiniExcel", rows[0].Column1);
            Assert.Equal(1, rows[0].Column2);
            Assert.Equal("Github", rows[1].Column1);
            Assert.Equal(2, rows[1].Column2);
        }

        Assert.Equal("A1:B3", SheetHelper.GetFirstSheetDimensionRefValue(path.ToString()));

        // test table
        var table = new DataTable();
        table.Columns.Add("a", typeof(string));
        table.Columns.Add("b", typeof(decimal));
        table.Columns.Add("c", typeof(bool));
        table.Columns.Add("d", typeof(DateTime));
        table.Rows.Add("some text", 1234567890, true, DateTime.Now);
        table.Rows.Add("<test>Hello World</test>", -1234567890, false, DateTime.Now.Date);

        using var pathTable = AutoDeletingPath.Create();
        _excelExporter.Export(pathTable.ToString(), table, configuration: config);
        Assert.Equal("A1:D3", SheetHelper.GetFirstSheetDimensionRefValue(pathTable.ToString()));

        // data reader
        var reader = table.CreateDataReader();
        using var pathReader = AutoDeletingPath.Create();

         _excelExporter.Export(pathReader.ToString(), reader, configuration: config, overwriteFile: true);
        Assert.Equal("A1:D3", SheetHelper.GetFirstSheetDimensionRefValue(pathTable.ToString())); //TODO: fix datareader not writing ref dimension (also in async version)
    }

    [Fact]
    public void SaveAsByDapperRows()
    {
        using var file = AutoDeletingPath.Create();
        var path = file.ToString();

        // Dapper Query
        using (var connection = Db.GetConnection("Data Source=:memory:"))
        {
            var rows = connection.Query("select 'MiniExcel' as Column1,1 as Column2 union all select 'Github',2");
             _excelExporter.Export(path, rows);
        }

        Assert.Equal("A1:B3", SheetHelper.GetFirstSheetDimensionRefValue(path));

        using (var stream = File.OpenRead(path))
        {
            var rows = _excelImporter.Query(stream, hasHeaderRow: true).ToList();

            Assert.Equal("MiniExcel", rows[0].Column1);
            Assert.Equal(1, rows[0].Column2);
            Assert.Equal("Github", rows[1].Column1);
            Assert.Equal(2, rows[1].Column2);
        }

        // Empty
        using (var connection = Db.GetConnection("Data Source=:memory:"))
        {
            var rows = connection.Query("with cte as (select 'MiniExcel' as Column1,1 as Column2 union all select 'Github',2)select * from cte where 1=2").ToList();
             _excelExporter.Export(path, rows, overwriteFile: true);
        }

        using (var stream = File.OpenRead(path))
        {
            var rows = _excelImporter.Query(stream, hasHeaderRow: false).ToList();
            Assert.Empty(rows);
        }

        using (var stream = File.OpenRead(path))
        {
            var rows = _excelImporter.Query(stream, hasHeaderRow: true).ToList();
            Assert.Empty(rows);
        }

        Assert.Equal("A1", SheetHelper.GetFirstSheetDimensionRefValue(path));

        // ToList
        using (var connection = Db.GetConnection("Data Source=:memory:"))
        {
            var rows = connection.Query("select 'MiniExcel' as Column1,1 as Column2 union all select 'Github',2").ToList();
             _excelExporter.Export(path, rows, overwriteFile: true);
        }

        Assert.Equal("A1:B3", SheetHelper.GetFirstSheetDimensionRefValue(path));

        using (var stream = File.OpenRead(path))
        {
            var rows = _excelImporter.Query(stream, hasHeaderRow: false).ToList();

            Assert.Equal("Column1", rows[0].A);
            Assert.Equal("Column2", rows[0].B);
            Assert.Equal("MiniExcel", rows[1].A);
            Assert.Equal(1, rows[1].B);
            Assert.Equal("Github", rows[2].A);
            Assert.Equal(2, rows[2].B);
        }

        using (var stream = File.OpenRead(path))
        {
            var rows = _excelImporter.Query(stream, hasHeaderRow: true).ToList();

            Assert.Equal("MiniExcel", rows[0].Column1);
            Assert.Equal(1, rows[0].Column2);
            Assert.Equal("Github", rows[1].Column1);
            Assert.Equal(2, rows[1].Column2);
        }
    }

    [Fact]
    public void SQLiteInsertTest()
    {
        // Avoid SQL Insert Large Size Xlsx OOM
        var path = PathHelper.GetFile("xlsx/Test5x2.xlsx");
        var tempSqlitePath = AutoDeletingPath.Create(Path.GetTempPath(), $"{Guid.NewGuid()}.db");
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
                var rows = _excelImporter.Query(stream);
                foreach (var row in rows)
                {
                    _ = connection.Execute("insert into T (A,B) values (@A,@B)", new { row.A, row.B }, transaction: transaction);
                }

                transaction.Commit();
            }
        }

        using (var connection = new SQLiteConnection(connectionString))
        {
            var result = connection.Query("select * from T");
            Assert.Equal(5, result.Count());
        }
    }

    [Fact]
    public void SaveAsBasicCreateTest()
    {
        using var path = AutoDeletingPath.Create();

        var rowsWritten = _excelExporter.Export(path.ToString(), new[]
        {
            new { Column1 = "MiniExcel", Column2 = 1 },
            new { Column1 = "Github", Column2 = 2}
        });

        Assert.Single(rowsWritten);
        Assert.Equal(2, rowsWritten[0]);

        using (var stream = File.OpenRead(path.ToString()))
        {
            var rows = _excelImporter.Query(stream, hasHeaderRow: true).ToList();

            Assert.Equal("MiniExcel", rows[0].Column1);
            Assert.Equal(1, rows[0].Column2);
            Assert.Equal("Github", rows[1].Column1);
            Assert.Equal(2, rows[1].Column2);
        }

        Assert.Equal("A1:B3", SheetHelper.GetFirstSheetDimensionRefValue(path.ToString()));
    }

    [Fact]
    public void SaveAsBasicStreamTest()
    {
        {
            using var path = AutoDeletingPath.Create();
            var values = new[]
            {
                new { Column1 = "MiniExcel", Column2 = 1 },
                new { Column1 = "Github", Column2 = 2 }
            };
            using (var stream = new FileStream(path.ToString(), FileMode.CreateNew))
            {
                var rowsWritten = _excelExporter.Export(stream, values);
                Assert.Single(rowsWritten);
                Assert.Equal(2, rowsWritten[0]);
            }

            using (var stream = File.OpenRead(path.ToString()))
            {
                var rows = _excelImporter.Query(stream, hasHeaderRow: true).ToList();

                Assert.Equal("MiniExcel", rows[0].Column1);
                Assert.Equal(1, rows[0].Column2);
                Assert.Equal("Github", rows[1].Column1);
                Assert.Equal(2, rows[1].Column2);
            }
        }
        {
            using var path = AutoDeletingPath.Create();
            var values = new[]
            {
                new { Column1 = "MiniExcel", Column2 = 1 },
                new { Column1 = "Github", Column2 = 2}
            };
            using (var stream = new MemoryStream())
            using (var fileStream = new FileStream(path.ToString(), FileMode.Create))
            {
                var rowsWritten = _excelExporter.Export(stream, values);
                stream.Seek(0, SeekOrigin.Begin);
                stream.CopyTo(fileStream);
                Assert.Single(rowsWritten);
                Assert.Equal(2, rowsWritten[0]);
            }

            using (var stream = File.OpenRead(path.ToString()))
            {
                var rows = _excelImporter.Query(stream, hasHeaderRow: true).ToList();

                Assert.Equal("MiniExcel", rows[0].Column1);
                Assert.Equal(1, rows[0].Column2);
                Assert.Equal("Github", rows[1].Column1);
                Assert.Equal(2, rows[1].Column2);
            }
        }
    }

    [Fact]
    public void SaveAsSpecialAndTypeCreateTest()
    {
        using var path = AutoDeletingPath.Create();
        var rowsWritten = _excelExporter.Export(path.ToString(), new[]
        {
            new { a = @"""<>+-*//}{\\n", b = 1234567890, c = true, d = DateTime.Now },
            new { a = "<test>Hello World</test>", b = -1234567890, c = false, d = DateTime.Now.Date }
        });
        Assert.Single(rowsWritten);
        Assert.Equal(2, rowsWritten[0]);

        var info = new FileInfo(path.ToString());
        Assert.True(info.FullName == path.ToString());
    }

    [Fact]
    public void SaveAsFileEpplusCanReadTest()
    {
        var now = new DateTime(2026, 6, 2, 15, 2, 33);
        using var path = AutoDeletingPath.Create();
        var rowsWritten = _excelExporter.Export(path.ToString(), new[]
        {
            new { a = @"""<>+-*//}{\\n", b = 1234567890, c = true, d = now},
            new { a = "<test>Hello World</test>", b = -1234567890, c = false, d = now.Date }
        });
        Assert.Single(rowsWritten);
        Assert.Equal(2, rowsWritten[0]);

        using var p = new ExcelPackage(new FileInfo(path.ToString()));
        var ws = p.Workbook.Worksheets.First();

        Assert.Equal("a", ws.Cells["A1"].Value.ToString());
        Assert.Equal("b", ws.Cells["B1"].Value.ToString());
        Assert.Equal("c", ws.Cells["C1"].Value.ToString());
        Assert.Equal("d", ws.Cells["D1"].Value.ToString());

        Assert.Equal(@"""<>+-*//}{\\n", ws.Cells["A2"].Value.ToString());
        Assert.Equal("1234567890", ws.Cells["B2"].Value.ToString());
        Assert.True(ws.Cells["C2"].GetCellValue<bool>());
        Assert.Equal(now, ws.Cells["D2"].GetValue<DateTime>());
    }

    [Fact]
    public void SavaAsClosedXmlCanReadTest()
    {
        var now = new DateTime(2026, 6, 2, 15, 3, 19);
        using var path = AutoDeletingPath.Create();
        var rowsWritten = _excelExporter.Export(path.ToString(), new[]
        {
            new { a = @"""<>+-*//}{\\n", b = 1234567890, c = true, d = now },
            new { a = "<test>Hello World</test>", b = -1234567890, c = false, d = now.Date }
        }, sheetName: "R&D");

        Assert.Single(rowsWritten);
        Assert.Equal(2, rowsWritten[0]);

        using var workbook = new XLWorkbook(path.ToString());
        var ws = workbook.Worksheets.First();

        Assert.Equal("a", ws.Cell("A1").Value.ToString());
        Assert.Equal("d", ws.Cell("D1").Value.ToString());
        Assert.Equal("b", ws.Cell("B1").Value.ToString());
        Assert.Equal("c", ws.Cell("C1").Value.ToString());

        Assert.Equal(@"""<>+-*//}{\\n", ws.Cell("A2").Value.ToString());
        Assert.Equal("1234567890", ws.Cell("B2").Value.ToString());
        Assert.True(ws.Cell("C2").GetValue<bool>());
        Assert.Equal(now, ws.Cell("D2").GetDateTime());

        Assert.Equal("R&D", ws.Name);
    }

    [Fact]
    public void ContentTypeUriContentTypeReadCheckTest()
    {
        var now = DateTime.Now;
        using var path = AutoDeletingPath.Create();
        var rowsWritten = _excelExporter.Export(path.ToString(), new[]
        {
            new { a = @"""<>+-*//}{\\n", b = 1234567890, c = true, d= now },
            new { a = "<test>Hello World</test>", b = -1234567890, c = false, d = now.Date }
        });
        Assert.Single(rowsWritten);
        Assert.Equal(2, rowsWritten[0]);

        using var zip = Package.Open(path.ToString(), FileMode.Open);
        var allParts = zip.GetParts()
            .Select(s => new { s.CompressionOption, s.ContentType, s.Uri, s.Package.GetType().Name })
            .ToDictionary(s => s.Uri.ToString(), s => s);

        Assert.Equal("application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml", allParts["/xl/styles.xml"].ContentType);
        Assert.Equal("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml", allParts["/xl/workbook.xml"].ContentType);
        Assert.Equal("application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml", allParts["/xl/worksheets/sheet1.xml"].ContentType);
        Assert.Equal("application/vnd.openxmlformats-package.relationships+xml", allParts["/xl/_rels/workbook.xml.rels"].ContentType);
        Assert.Equal("application/vnd.openxmlformats-package.relationships+xml", allParts["/_rels/.rels"].ContentType);
    }

    [Fact]
    public void DynamicColumnsConfigurationIsUsedWhenCreatingExcelUsingIDataReader()
    {
        using var path = AutoDeletingPath.Create();
        var dateTime = DateTime.Now;
        var onlyDate = DateOnly.FromDateTime(dateTime);

        var table = new DataTable();
        table.Columns.Add("Column1", typeof(string));
        table.Columns.Add("Column2", typeof(int));
        table.Columns.Add("Column3", typeof(DateTime));
        table.Columns.Add("Column4", typeof(DateOnly));
        table.Rows.Add("MiniExcel", 1, dateTime, onlyDate);
        table.Rows.Add("Github", 2, dateTime, onlyDate);

        var configuration = new OpenXmlConfiguration
        {
            DynamicColumns =
            [
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
            ]
        };
        var reader = table.CreateDataReader();
         _excelExporter.Export(path.ToString(), reader, configuration: configuration);

        using var stream = File.OpenRead(path.ToString());
        var rows = _excelImporter.Query(stream, hasHeaderRow: true)
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

    [Fact]
    public void DynamicColumnsConfigurationIsUsedWhenCreatingExcelUsingDataTable()
    {
        using var path = AutoDeletingPath.Create();
        var dateTime = DateTime.Now;
        var onlyDate = DateOnly.FromDateTime(dateTime);

        var table = new DataTable();
        table.Columns.Add("Column1", typeof(string));
        table.Columns.Add("Column2", typeof(int));
        table.Columns.Add("Column3", typeof(DateTime));
        table.Columns.Add("Column4", typeof(DateOnly));
        table.Rows.Add("MiniExcel", 1, dateTime, onlyDate);
        table.Rows.Add("Github", 2, dateTime, onlyDate);

        var configuration = new OpenXmlConfiguration
        {
            DynamicColumns =
            [
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
            ]
        };
         _excelExporter.Export(path.ToString(), table, configuration: configuration);

        using var stream = File.OpenRead(path.ToString());
        var rows = _excelImporter.Query(stream, hasHeaderRow: true)
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

    [Fact]
    public void InsertSheetTest()
    {
        var now = new DateTime(2026, 6, 2, 15, 4, 51);
        using var file = AutoDeletingPath.Create();
        var path = file.ToString();

        {
            var table = new DataTable();
            table.Columns.Add("a", typeof(string));
            table.Columns.Add("b", typeof(decimal));
            table.Columns.Add("c", typeof(bool));
            table.Columns.Add("d", typeof(DateTime));
            table.Rows.Add(@"""<>+-*//}{\\n", 1234567890, true, now);
            table.Rows.Add("<test>Hello World</test>", -1234567890, false, now.Date);

            var rowsWritten = _excelExporter.InsertSheet(path, table, sheetName: "Sheet1");
            Assert.Equal(2, rowsWritten);

            using var p = new ExcelPackage(path);
            var sheet1 = p.Workbook.Worksheets[0];

            Assert.Equal("a", sheet1.Cells["A1"].Value.ToString());
            Assert.Equal("b", sheet1.Cells["B1"].Value.ToString());
            Assert.Equal("c", sheet1.Cells["C1"].Value.ToString());
            Assert.Equal("d", sheet1.Cells["D1"].Value.ToString());

            Assert.Equal(@"""<>+-*//}{\\n", sheet1.Cells["A2"].Value.ToString());
            Assert.Equal("1234567890", sheet1.Cells["B2"].Value.ToString());
            Assert.True(sheet1.Cells["C2"].GetCellValue<bool>());
            Assert.Equal(now, sheet1.Cells["D2"].GetCellValue<DateTime>());

            Assert.Equal("Sheet1", sheet1.Name);
        }
        {
            var table = new DataTable();
            table.Columns.Add("Column1", typeof(string));
            table.Columns.Add("Column2", typeof(int));
            table.Rows.Add("MiniExcel", 1);
            table.Rows.Add("Github", 2);

            var rowsWritten = _excelExporter.InsertSheet(path, table, sheetName: "Sheet2");
            Assert.Equal(2, rowsWritten);

            using var p = new ExcelPackage(path);
            var sheet2 = p.Workbook.Worksheets[1];

            Assert.Equal("Column1", sheet2.Cells["A1"].Value.ToString());
            Assert.Equal("Column2", sheet2.Cells["B1"].Value.ToString());

            Assert.Equal("MiniExcel", sheet2.Cells["A2"].Value.ToString());
            Assert.Equal("1", sheet2.Cells["B2"].Value.ToString());

            Assert.Equal("Github", sheet2.Cells["A3"].Value.ToString());
            Assert.Equal("2", sheet2.Cells["B3"].Value.ToString());

            Assert.Equal("Sheet2", sheet2.Name);
        }
        {
            var table = new DataTable();
            table.Columns.Add("Column1", typeof(string));
            table.Columns.Add("Column2", typeof(DateTime));
            table.Rows.Add("Test", now);
        
            var rowsWritten = _excelExporter.InsertSheet(path, table, sheetName: "Sheet2", printHeader: false, configuration: new OpenXmlConfiguration
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
            }, overwriteSheet: true);
        
            Assert.Equal(1, rowsWritten);
        
            using var p = new ExcelPackage(path);
            var sheet2 = p.Workbook.Worksheets[1];
        
            Assert.Equal("Test", sheet2.Cells["A1"].Value.ToString());
            Assert.True(sheet2.Cells["B1"].Text == now.ToString("dd.MM.yyyy HH:mm:ss"));
            Assert.Equal("Sheet2", sheet2.Name);
        }
        {
            var table = new DataTable();
            table.Columns.Add("Column1", typeof(string));
            table.Columns.Add("Column2", typeof(DateTime));
            table.Rows.Add("MiniExcel", now);
            table.Rows.Add("Github", now);

            var rowsWritten = _excelExporter.InsertSheet(path, table, sheetName: "Sheet3", configuration: new OpenXmlConfiguration
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

            using var p = new ExcelPackage(path);
            var sheet3 = p.Workbook.Worksheets[2];

            Assert.Equal("Column1", sheet3.Cells["A1"].Value.ToString());
            Assert.Equal("Its Date", sheet3.Cells["B1"].Value.ToString());

            Assert.Equal("MiniExcel", sheet3.Cells["A2"].Value.ToString());
            Assert.True(sheet3.Cells["B2"].Text == now.ToString("dd.MM.yyyy HH:mm:ss"));

            Assert.Equal("Github", sheet3.Cells["A3"].Value.ToString());
            Assert.True(sheet3.Cells["B3"].Text == now.ToString("dd.MM.yyyy HH:mm:ss"));

            Assert.Equal("Sheet3", sheet3.Name);
        }
    }

    [Fact]
    public void CopyAndInsertSheetTest()
    {
        var now = DateTime.Now;
        var dt = new DateTime(now.Year, now.Month, now.Day, now.Hour, now.Minute, now.Second);

        using var firstFile = AutoDeletingPath.Create();
        using var secondFile = AutoDeletingPath.Create();
        using var thirdFile = AutoDeletingPath.Create();

        var firstPath = firstFile.ToString();
        var secondPath = secondFile.ToString();
        var thirdPath = thirdFile.ToString();

        {
            using var table = new DataTable();
            table.Columns.Add("a", typeof(string));
            table.Columns.Add("b", typeof(decimal));
            table.Columns.Add("c", typeof(bool));
            table.Columns.Add("d", typeof(DateTime));
            table.Rows.Add(@"""<>+-*//}{\\n", 1234567890, true, dt);
            table.Rows.Add("<test>Hello World</test>", -1234567890, false, dt.Date);
            _excelExporter.Export(firstPath, table);
        }
        {
            using var table = new DataTable();
            table.Columns.Add("Column1", typeof(string));
            table.Columns.Add("Column2", typeof(int));
            table.Rows.Add("MiniExcel", 1);
            table.Rows.Add("Github", 2);

            _excelExporter.CopyAndAddSheet(firstPath, secondPath, table, sheetName: "Sheet2");
            using var p = new ExcelPackage(secondPath);
            var sheet2 = p.Workbook.Worksheets[1];

            Assert.Equal("Column1", sheet2.Cells["A1"].Value.ToString());
            Assert.Equal("Column2", sheet2.Cells["B1"].Value.ToString());

            Assert.Equal("MiniExcel", sheet2.Cells["A2"].Value.ToString());
            Assert.Equal(1, (double)sheet2.Cells["B2"].Value);

            Assert.Equal("Github", sheet2.Cells["A3"].Value.ToString());
            Assert.Equal(2, (double)sheet2.Cells["B3"].Value);

            Assert.Equal("Sheet2", sheet2.Name);
        }
        {
            _excelExporter.CopyAndAddSheet(secondPath, thirdPath, new[] { new { Column1 = "Test", Column2 = dt } }, sheetName: "Sheet2", printHeader: false, configuration: new OpenXmlConfiguration
            {
                AutoFilter = false,
                TableStyles = TableStyles.None,
                DynamicColumns =
                [
                    new DynamicExcelColumn("Column2")
                    {
                        Name = "Date",
                        Index = 1,
                        Width = 150,
                        Format = "dd.mm.yyyy hh:mm:ss"
                    }
                ]
            }, overwriteSheet: true);

            using var p = new ExcelPackage(thirdPath);
            var sheet2 = p.Workbook.Worksheets[1];

            Assert.Equal("Sheet2", sheet2.Name);
            Assert.Equal("Test", sheet2.Cells["A1"].Value);
            Assert.Equal(dt.ToString("dd.MM.yyyy HH:mm:ss"), sheet2.Cells["B1"].Text);
        }
        {
            using var table = new DataTable();
            table.Columns.Add("Column1", typeof(string));
            table.Columns.Add("Column2", typeof(DateTime));
            table.Rows.Add("MiniExcel", dt);
            table.Rows.Add("Github", dt);
            using var reader = table.CreateDataReader();

            using var fs = File.OpenRead(thirdPath);
            using var ms = new MemoryStream();

            _excelExporter.CopyAndAddSheet(fs, ms, reader, sheetName: "Sheet3", configuration: new OpenXmlConfiguration
            {
                AutoFilter = false,
                TableStyles = TableStyles.None,
                DynamicColumns =
                [
                    new DynamicExcelColumn("Column2")
                    {
                        Name = "Date",
                        Index = 1,
                        Width = 150,
                        Format = "dd.mm.yyyy hh:mm:ss"
                    }
                ]
            });

            using var p = new ExcelPackage(ms);
            var sheet3 = p.Workbook.Worksheets[2];

            Assert.Equal("Column1", sheet3.Cells["A1"].Value);
            Assert.Equal("Date", sheet3.Cells["B1"].Value);

            Assert.Equal("MiniExcel", sheet3.Cells["A2"].Value);
            Assert.Equal(dt.ToString("dd.MM.yyyy HH:mm:ss"), sheet3.Cells["B2"].Text);

            Assert.Equal("Github", sheet3.Cells["A3"].Value);
            Assert.Equal(dt.ToString("dd.MM.yyyy HH:mm:ss"), sheet3.Cells["B3"].Text);

            Assert.Equal("Sheet3", sheet3.Name);
        }
    }

    [Fact]
    public void ExportAndQueryMixedFieldAndPropertyTest()
    {
        using var path = AutoDeletingPath.Create();
        var input = new[] { new MixedFieldPropertyTest { Field1 = "F", Prop1 = "P" } };

         _excelExporter.Export(path.ToString(), input);

        var rows = _excelImporter.Query(path.ToString(), true).ToList();
        var first = rows[0] as IDictionary<string, object>;

        Assert.Contains("F1", first!.Keys);
        Assert.Contains("P1", first.Keys);
    }

    [Fact]
    public void ExportAndQueryFieldsWithoutAttributeTest()
    {
        using var path = AutoDeletingPath.Create();
        var input = new[] { new FieldsWithoutAttributeTest { NotMappedField = "NO", MappedField = "YES" } };

         _excelExporter.Export(path.ToString(), input);

        var rows = _excelImporter.Query(path.ToString(), true).Cast<IDictionary<string, object>>().ToList();
        Assert.Contains("Mapped", rows[0].Keys);
        Assert.DoesNotContain("NotMappedField", rows[0].Keys);
    }
    
    [Fact]
    public async Task InvalidSheetNameCharactersShouldThrow()
    {
        await using var ms1 = new MemoryStream();
        Assert.Throws<ArgumentException>(() => _excelExporter.Export(ms1, Array.Empty<object>(), sheetName: "Sheet?"));
        
        await using var ms2 = new MemoryStream();
        Assert.Throws<ArgumentException>(() => _excelExporter.InsertSheet(ms2, Array.Empty<object>(), sheetName: "Sheet[]"));
        
        await using var ms3 = new MemoryStream();
        using var package = new ExcelPackage(ms3);
        package.Workbook.Worksheets.Add("Sheet1");
        package.Save();
        
        ms1.Seek(0, SeekOrigin.Begin);
        Assert.Throws<ArgumentException>(() => _excelExporter.AlterSheet(ms3, "Sheet1", "Sheet*"));
    }
    
    [Theory]
    [InlineData("")]
    [InlineData("it")]
    [InlineData("zh")]
    public void LocalizationTest(string cultureId)
    {
        var ogCulture = CultureInfo.CurrentUICulture;

        try
        {
            CultureInfo.CurrentUICulture = new CultureInfo(cultureId);

            using var ms = new MemoryStream();
            _excelExporter.Export(ms, Array.Empty<LocalizationSupportDto>());
            ms.Seek(0, SeekOrigin.Begin);
            
            using var package = new ExcelPackage(ms);
            var cells = package.Workbook.Worksheets[0].Cells;

            var (firstName, lastName, address, age) = cultureId switch
            {
                "" => ("Name", "Surname", "Address", "Age"),
                "it" => ("Nome", "Cognome", "Indirizzo", "Età"),
                "zh" => ("名", "姓", "地址", "年龄"),
                _ => throw new UnreachableException()
            };
            
            Assert.Equal(firstName, cells["A1"].Value);
            Assert.Equal(lastName, cells["B1"].Value);
            Assert.Equal(address, cells["C1"].Value);
            Assert.Equal(age, cells["D1"].Value);
        }
        finally
        {
            CultureInfo.CurrentUICulture = ogCulture;
        }
    }

    [Theory]
    [InlineData("")]
    [InlineData("it")]
    [InlineData("zh")]
    public void LocalizationTestDynamicColumns(string cultureId)
    {
        var ogCulture = CultureInfo.CurrentUICulture;

        try
        {
            CultureInfo.CurrentUICulture = new CultureInfo(cultureId);

            DynamicExcelColumn[] cols = [
                new("FirstName") { ResourceType = typeof(Localization) }, 
                new("LastName") { ResourceType = typeof(Localization) },
                new("Address") { ResourceType = typeof(Localization) },
                new("Age") { ResourceType = typeof(Localization) }
            ]; 

            using var stream = new MemoryStream();
            _excelExporter.Export(
                stream, 
                new[] { new { FirstName = "", LastName = "", Address = "", Age = 0 } },
                configuration: new OpenXmlConfiguration { DynamicColumns =  cols });
            
            stream.Seek(0, SeekOrigin.Begin);
            using var package = new ExcelPackage(stream);
            var cells = package.Workbook.Worksheets[0].Cells;

            var (firstName, lastName, address, age) = cultureId switch
            {
                "" => ("Name", "Surname", "Address", "Age"),
                "it" => ("Nome", "Cognome", "Indirizzo", "Età"),
                "zh" => ("名", "姓", "地址", "年龄"),
                _ => throw new UnreachableException()
            };
            
            Assert.Equal(firstName, cells["A1"].Value);
            Assert.Equal(lastName, cells["B1"].Value);
            Assert.Equal(address, cells["C1"].Value);
            Assert.Equal(age, cells["D1"].Value);
        }
        finally
        {
            CultureInfo.CurrentUICulture = ogCulture;
        }
    }
}
