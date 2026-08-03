using ClosedXML.Excel;
using MiniExcelLib.OpenXml.Tests.Utils;
using MiniExcelLib.Tests.Common;
using MiniExcelLib.Tests.Common.Utils;

namespace MiniExcelLib.OpenXml.Tests.Main;

public class MiniExcelOpenXmlExporterAsyncTests
{
    private readonly OpenXmlImporter _excelImporter =  MiniExcel.Importers.GetOpenXmlImporter();
    private readonly OpenXmlExporter _excelExporter =  MiniExcel.Exporters.GetOpenXmlExporter();
   
    static MiniExcelOpenXmlExporterAsyncTests()
    {
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
    }
    
    [Fact]
    public async Task SaveAsControlChracter()
    {
        using var file = AutoDeletingPath.Create();
        var path = file.ToString();

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
        await _excelExporter.ExportAsync(path, input);

        var rows2 = await _excelImporter.QueryAsync(path, true).ToListAsync();
        var rows1 = await _excelImporter.QueryAsync<SaveAsControlChracterVO>(path).ToListAsync();
    }

    [Fact]
    public async Task SaveAsCustomAttributesTest()
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

        await _excelExporter.ExportAsync(path.ToString(), input);
        var rows = await _excelImporter.QueryAsync(path.ToString(), true).ToListAsync();
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
    public async Task SaveAsCustomAttributes2Test()
    {
        using var path = AutoDeletingPath.Create();
        var input = Enumerable.Range(1, 3)
            .Select(_ => new ExcelAttributeDemo2
            {
                Test1 = "Test1",
                Test2 = "Test2",
                Test3 = "Test3",
                Test4 = "Test4",
            });

        await _excelExporter.ExportAsync(path.ToString(), input);
        var rows = await _excelImporter.QueryAsync(path.ToString(), true).ToListAsync();
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
    public async Task SaveAsFileWithDimensionByICollection()
    {
        //List<strongtype>
        {
            List<SaveAsFileWithDimensionByICollectionTestType> values =
            [
                new() { A = "A", B = "B" },
                new() { A = "A", B = "B" }
            ];

            using (var file = AutoDeletingPath.Create())
            {
                var path = file.ToString();
                await _excelExporter.ExportAsync(path, values);
                await using (var stream = File.OpenRead(path))
                {
                    var rows = await _excelImporter.QueryAsync(stream, hasHeaderRow: false).Cast<IDictionary<string, object>>().ToListAsync();
                    Assert.Equal(3, rows.Count);
                    Assert.Equal("A", rows[0]["A"]);
                    Assert.Equal("A", rows[1]["A"]);
                    Assert.Equal("A", rows[2]["A"]);
                }

                await using (var stream = File.OpenRead(path))
                {
                    var rows = await _excelImporter.QueryAsync(stream, hasHeaderRow: true).Cast<IDictionary<string, object>>().ToListAsync();
                    Assert.Equal(2, rows.Count);
                    Assert.Equal("A", rows[0]["A"]);
                    Assert.Equal("A", rows[1]["A"]);
                }

                Assert.Equal("A1:B3", SheetHelper.GetFirstSheetDimensionRefValue(path));
            }

            using var newPath = AutoDeletingPath.Create();
            await _excelExporter.ExportAsync(newPath.ToString(), values, false);
            Assert.Equal("A1:B2", SheetHelper.GetFirstSheetDimensionRefValue(newPath.ToString()));
        }

        //List<strongtype> empty
        {
            List<SaveAsFileWithDimensionByICollectionTestType> values = [];
            using (var file = AutoDeletingPath.Create())
            {
                var path = file.ToString();
                await _excelExporter.ExportAsync(path, values, false);
                await using (var stream = File.OpenRead(path))
                {
                    var rows = await _excelImporter.QueryAsync(stream, hasHeaderRow: false).ToListAsync();
                    Assert.Empty(rows);
                }

                Assert.Equal("A1:B1", SheetHelper.GetFirstSheetDimensionRefValue(path));
            }

            using (var file = AutoDeletingPath.Create())
            {
                var path = file.ToString();
                await _excelExporter.ExportAsync(path, values);
                {
                    await using var stream = File.OpenRead(path);
                    var rows = await _excelImporter.QueryAsync(stream, hasHeaderRow: false).ToListAsync();
                    Assert.Single(rows);
                }
                Assert.Equal("A1:B1", SheetHelper.GetFirstSheetDimensionRefValue(path));
            }
        }

        //Array<anoymous>
        {
            var values = new[]
            {
                new { A = "A", B = "B" },
                new { A = "A", B = "B" },
            };

            using (var file = AutoDeletingPath.Create())
            {
                var path = file.ToString();
                await _excelExporter.ExportAsync(path, values);

                await using (var stream = File.OpenRead(path))
                {
                    var rows = await _excelImporter.QueryAsync(stream, hasHeaderRow: false).Cast<IDictionary<string, object>>().ToListAsync();
                    Assert.Equal(3, rows.Count);
                    Assert.Equal("A", rows[0]["A"]);
                    Assert.Equal("A", rows[1]["A"]);
                    Assert.Equal("A", rows[2]["A"]);
                }

                await using (var stream = File.OpenRead(path))
                {
                    var rows = await _excelImporter.QueryAsync(stream, hasHeaderRow: true).Cast<IDictionary<string, object>>().ToListAsync();
                    Assert.Equal(2, rows.Count);
                    Assert.Equal("A", rows[0]["A"]);
                    Assert.Equal("A", rows[1]["A"]);
                }

                Assert.Equal("A1:B3", SheetHelper.GetFirstSheetDimensionRefValue(path));
            }

            using (var path = AutoDeletingPath.Create())
            {
                await _excelExporter.ExportAsync(path.ToString(), values, false);
                Assert.Equal("A1:B2", SheetHelper.GetFirstSheetDimensionRefValue(path.ToString()));
            }
        }

        // without properties
        {
            using var path = AutoDeletingPath.Create();
            var values = new List<int>();
            await Assert.ThrowsAsync<NotSupportedException>(() =>  _excelExporter.ExportAsync(path.ToString(), values));
        }
    }

    [Fact]
    public async Task SaveAsFileWithDimension()
    {
        {
            using var file = AutoDeletingPath.Create();
            var path = file.ToString();

            using var table = new DataTable();
            await _excelExporter.ExportAsync(path, table);
            Assert.Equal("A1", SheetHelper.GetFirstSheetDimensionRefValue(path));
            
            var rows = await _excelImporter.QueryAsync(path).ToListAsync();
            Assert.Empty(rows);
            
            await _excelExporter.ExportAsync(path, table, printHeader: false, overwriteFile: true);
            Assert.Equal("A1", SheetHelper.GetFirstSheetDimensionRefValue(path));
        }
        {
            using var file = AutoDeletingPath.Create();
            var path = file.ToString();
            using var table = new DataTable();

            table.Columns.Add("a", typeof(string));
            table.Columns.Add("b", typeof(decimal));
            table.Columns.Add("c", typeof(bool));
            table.Columns.Add("d", typeof(DateTime));
            table.Rows.Add(@"""<>+-*//}{\\n", 1234567890);
            table.Rows.Add("<test>Hello World</test>", -1234567890, false, DateTime.Now);

            await _excelExporter.ExportAsync(path, table);
            Assert.Equal("A1:D3", SheetHelper.GetFirstSheetDimensionRefValue(path));

            await using (var stream = File.OpenRead(path))
            {
                var rows = await _excelImporter.QueryAsync(stream, hasHeaderRow: true)
                    .Cast<IDictionary<string, object>>()
                    .ToListAsync();
 
                Assert.Equal(2, rows.Count);
                Assert.Equal(@"""<>+-*//}{\\n", rows[0]["a"]);
                Assert.Equal(1234567890d, rows[0]["b"]);
                Assert.Null(rows[0]["c"]);
                Assert.Null(rows[0]["d"]);
            }

            await using (var stream = File.OpenRead(path))
            {
                var rows = await _excelImporter.QueryAsync(stream)
                    .Cast<IDictionary<string, object>>()
                    .ToListAsync();

                Assert.Equal(3, rows.Count);
                Assert.Equal("a", rows[0]["A"]);
                Assert.Equal("b", rows[0]["B"]);
                Assert.Equal("c", rows[0]["C"]);
                Assert.Equal("d", rows[0]["D"]);
            }

            await _excelExporter.ExportAsync(path, table, printHeader: false, overwriteFile: true);
            Assert.Equal("A1:D2", SheetHelper.GetFirstSheetDimensionRefValue(path));
        }

        {
            using var path = AutoDeletingPath.Create();
            using var table = new DataTable();
            table.Columns.Add("a", typeof(string));
            table.Rows.Add("A");
            table.Rows.Add("B");

            await _excelExporter.ExportAsync(path.ToString(), table);
            Assert.Equal("A1:A3", SheetHelper.GetFirstSheetDimensionRefValue(path.ToString()));
        }
    }

    [Fact]
    public async Task SaveAsByDataTableTest()
    {
        using var file = AutoDeletingPath.Create();
        var path = file.ToString();

        var now = DateTime.Now;
        var dt = new DateTime(now.Year, now.Month, now.Day, now.Hour, now.Minute, now.Second);

        using var table = new DataTable();
        table.Columns.Add("a", typeof(string));
        table.Columns.Add("b", typeof(decimal));
        table.Columns.Add("c", typeof(bool));
        table.Columns.Add("d", typeof(DateTime));
        table.Rows.Add(@"""<>+-*//}{\\n", 1234567890, true, dt);
        await _excelExporter.ExportAsync(path, table);

        using var p = new ExcelPackage(path);
        var cells = p.Workbook.Worksheets[0].Cells;

        Assert.Equal("a", cells["A1"].Text);
        Assert.Equal("b", cells["B1"].Text);
        Assert.Equal("c", cells["C1"].Text);
        Assert.Equal("d", cells["D1"].Text);

        Assert.Equal(@"""<>+-*//}{\\n", cells["A2"].Value);
        Assert.Equal(1234567890, (double)cells["B2"].Value);
        Assert.True((bool)cells["C2"].Value);
        Assert.Equal(dt, (DateTime)cells["D2"].Value);
    }

    [Fact]
    public async Task SaveEmptyDataReaderTest()
    {
        using var path = AutoDeletingPath.Create();
        await using (var connection = Db.GetConnection())
        {
            var rows = await connection.QueryAsync("with cte as (select 1 id,2 val) select * from cte where 1=2");
            await _excelExporter.ExportAsync(path.ToString(), rows);
        }

        await using (var stream = File.OpenRead(path.ToString()))
        {
            var row = await _excelImporter.QueryAsync(stream, hasHeaderRow: true).ToListAsync();
            Assert.Empty(row);
        }
    }

    [Fact]
    public async Task SaveEmptyAsyncEnumerableTest()
    {
        using var path = AutoDeletingPath.Create();
        await using (var connection = Db.GetConnection())
        {
            var data = Array.Empty<object>().ToAsyncEnumerable();
            await _excelExporter.ExportAsync(path.ToString(), data);
        }

        await using (var stream = File.OpenRead(path.ToString()))
        {
            var data = await _excelImporter.QueryAsync(stream).ToListAsync();
            Assert.Empty(data);
        }
    }

    [Fact]
    public async Task SaveAsByIEnumerableIDictionary()
    {
        {
            using var file = AutoDeletingPath.Create();
            var path = file.ToString();
            List<Dictionary<string, object>> values =
            [
                new() { { "Column1", "MiniExcel" }, { "Column2", 1 } },
                new() { { "Column1", "Github" }, { "Column2", 2 } }
            ];
            await _excelExporter.ExportAsync(path, values);

            await using (var stream = File.OpenRead(path))
            {
                var rows = await _excelImporter.QueryAsync(stream, hasHeaderRow: false).Cast<IDictionary<string, object>>().ToListAsync();
                Assert.Equal("Column1", rows[0]["A"]);
                Assert.Equal("Column2", rows[0]["B"]);
                Assert.Equal("MiniExcel", rows[1]["A"]);
                Assert.Equal(1d, rows[1]["B"]);
                Assert.Equal("Github", rows[2]["A"]);
                Assert.Equal(2d, rows[2]["B"]);
            }

            await using (var stream = File.OpenRead(path))
            {
                var rows = await _excelImporter.QueryAsync(stream, hasHeaderRow: true).Cast<IDictionary<string, object>>().ToListAsync();
                Assert.Equal(2, rows.Count);
                Assert.Equal("MiniExcel", rows[0]["Column1"]);
                Assert.Equal(1d, rows[0]["Column2"]);
                Assert.Equal("Github", rows[1]["Column1"]);
                Assert.Equal(2d, rows[1]["Column2"]);
            }

            Assert.Equal("A1:B3", SheetHelper.GetFirstSheetDimensionRefValue(path));
        }

        {
            using var file = AutoDeletingPath.Create();
            var path = file.ToString();

            List<Dictionary<int, object>> values =
            [
                new() { { 1, "MiniExcel" }, { 2, 1 } },
                new() { { 1, "Github" }, { 2, 2 } }
            ];
            await _excelExporter.ExportAsync(path, values);

            await using (var stream = File.OpenRead(path))
            {
                var rows = await _excelImporter.QueryAsync(stream, hasHeaderRow: false).ToListAsync();
                Assert.Equal(3, rows.Count);
            }
            Assert.Equal("A1:B3", SheetHelper.GetFirstSheetDimensionRefValue(path));
        }
    }

    [Fact]
    public async Task SaveAsByIEnumerableIDictionaryWithDynamicConfiguration()
    {
        using var file = AutoDeletingPath.Create();
        var path = file.ToString();

        var dynamicColumns = new[]
        {
            new DynamicExcelColumn("Column1") { Name = "Name Column" },
            new DynamicExcelColumn("Column2") { Name = "Value Column" }
        };
        var config = new OpenXmlConfiguration
        {
            DynamicColumns = dynamicColumns
        };
        List<Dictionary<string, object>> values =
        [
            new() { { "Column1", "MiniExcel" }, { "Column2", 1 } },
            new() { { "Column1", "Github" }, { "Column2", 2 } }
        ];
        await _excelExporter.ExportAsync(path, values, configuration: config);

        await using (var stream = File.OpenRead(path))
        {
            var rows = await _excelImporter.QueryAsync(stream, hasHeaderRow: true).Cast<IDictionary<string, object>>().ToListAsync();
            Assert.Equal(2, rows.Count);
            Assert.Equal("Name Column", rows[0].Keys.ElementAt(0));
            Assert.Equal("Value Column", rows[0].Keys.ElementAt(1));
            Assert.Equal("MiniExcel", rows[0].Values.ElementAt(0));
            Assert.Equal(1d, rows[0].Values.ElementAt(1));
            Assert.Equal("Github", rows[1].Values.ElementAt(0));
            Assert.Equal(2d, rows[1].Values.ElementAt(1));
        }
        Assert.Equal("A1:B3", SheetHelper.GetFirstSheetDimensionRefValue(path));
    }

    [Fact]
    public async Task SaveAsFrozenRowsAndColumnsTest()
    {
        var config = new OpenXmlConfiguration
        {
            FreezeRowCount = 1,
            FreezeColumnCount = 2
        };

        // Test enumerable
        using var file = AutoDeletingPath.Create();
        var path = file.ToString();

        await _excelExporter.ExportAsync(
            path,
            new[]
            {
                new { Column1 = "MiniExcel", Column2 = 1 },
                new { Column1 = "Github", Column2 = 2}
            },
            configuration: config
        );

        await using (var stream = File.OpenRead(path))
        {
            var rows = await _excelImporter.QueryAsync(stream, hasHeaderRow: true).ToListAsync();

            Assert.Equal("MiniExcel", rows[0].Column1);
            Assert.Equal(1, rows[0].Column2);
            Assert.Equal("Github", rows[1].Column1);
            Assert.Equal(2, rows[1].Column2);
        }
        Assert.Equal("A1:B3", SheetHelper.GetFirstSheetDimensionRefValue(path));

        // test table
        var table = new DataTable();
        table.Columns.Add("a", typeof(string));
        table.Columns.Add("b", typeof(decimal));
        table.Columns.Add("c", typeof(bool));
        table.Columns.Add("d", typeof(DateTime));
        table.Rows.Add("some text", 1234567890, true, DateTime.Now);
        table.Rows.Add("<test>Hello World</test>", -1234567890, false, DateTime.Now.Date);

        using var pathTable = AutoDeletingPath.Create();
        await _excelExporter.ExportAsync(pathTable.ToString(), table, configuration: config);
        Assert.Equal("A1:D3", SheetHelper.GetFirstSheetDimensionRefValue(pathTable.ToString()));

        // data reader
        await using var reader = table.CreateDataReader();
        using var pathReader = AutoDeletingPath.Create();
        await _excelExporter.ExportAsync(pathReader.ToString(), reader, configuration: config);
        Assert.Equal("A1:D3", SheetHelper.GetFirstSheetDimensionRefValue(pathTable.ToString()));
    }

    [Fact]
    public async Task SaveAsByDapperRows()
    {
        using var file = AutoDeletingPath.Create();
        var path = file.ToString();

        // Dapper Query
        await using (var connection = Db.GetConnection())
        {
            var rows = await connection.QueryAsync("select 'MiniExcel' as Column1,1 as Column2 union all select 'Github',2");
            await _excelExporter.ExportAsync(path, rows);
        }
        Assert.Equal("A1:B3", SheetHelper.GetFirstSheetDimensionRefValue(path));

        await using (var stream = File.OpenRead(path))
        {
            var rows = await _excelImporter.QueryAsync(stream, hasHeaderRow: true).Cast<IDictionary<string, object>>().ToListAsync();
            Assert.Equal("MiniExcel", rows[0]["Column1"]);
            Assert.Equal(1d, rows[0]["Column2"]);
            Assert.Equal("Github", rows[1]["Column1"]);
            Assert.Equal(2d, rows[1]["Column2"]);
        }

        // Empty
        await using (var connection = Db.GetConnection())
        {
            var rows = await connection.QueryAsync("with cte as (select 'MiniExcel' as Column1,1 as Column2 union all select 'Github',2)select * from cte where 1=2");
            await _excelExporter.ExportAsync(path, rows.AsList(), overwriteFile: true);
        }

        await using (var stream = File.OpenRead(path))
        {
            var rows = await _excelImporter.QueryAsync(stream, hasHeaderRow: false).ToListAsync();
            Assert.Empty(rows);
        }

        await using (var stream = File.OpenRead(path))
        {
            var rows = await _excelImporter.QueryAsync(stream, hasHeaderRow: true).ToListAsync();
            Assert.Empty(rows);
        }
        Assert.Equal("A1", SheetHelper.GetFirstSheetDimensionRefValue(path));

        // ToList
        await using (var connection = Db.GetConnection())
        {
            var rows = (await connection.QueryAsync("select 'MiniExcel' as Column1,1 as Column2 union all select 'Github',2")).ToList();
            await _excelExporter.ExportAsync(path, rows, overwriteFile: true);
        }
        Assert.Equal("A1:B3", SheetHelper.GetFirstSheetDimensionRefValue(path));

        await using (var stream = File.OpenRead(path))
        {
            var rows = await _excelImporter.QueryAsync(stream, hasHeaderRow: false).Cast<IDictionary<string, object>>().ToListAsync();
            Assert.Equal("Column1", rows[0]["A"]);
            Assert.Equal("Column2", rows[0]["B"]);
            Assert.Equal("MiniExcel", rows[1]["A"]);
            Assert.Equal(1d, rows[1]["B"]);
            Assert.Equal("Github", rows[2]["A"]);
            Assert.Equal(2d, rows[2]["B"]);
        }

        await using (var stream = File.OpenRead(path))
        {
            var rows = await _excelImporter.QueryAsync(stream, hasHeaderRow: true).Cast<IDictionary<string, object>>().ToListAsync();
            Assert.Equal("MiniExcel", rows[0]["Column1"]);
            Assert.Equal(1d, rows[0]["Column2"]);
            Assert.Equal("Github", rows[1]["Column1"]);
            Assert.Equal(2d, rows[1]["Column2"]);
        }
    }

    [Fact]
    public async Task SQLiteInsertTest()
    {
        // Async Task SQL Insert Large Size Xlsx OOM
        var path = PathHelper.GetFile("xlsx/Test5x2.xlsx");
        using var tempSqlitePath = AutoDeletingPath.Create(Path.GetTempPath(), $"{Guid.NewGuid()}.db");
        var connectionString = $"Data Source={tempSqlitePath};Version=3;";

        await using (var connection = new SQLiteConnection(connectionString))
        {
            await connection.ExecuteAsync("create table T (A varchar(20),B varchar(20));");
        }

        await using (var connection = new SQLiteConnection(connectionString))
        {
            await connection.OpenAsync();
            await using (var transaction = connection.BeginTransaction())
            await using (var stream = File.OpenRead(path))
            {
                var rows = _excelImporter.QueryAsync(stream).Cast<IDictionary<string, object>>();
                await foreach (var row in rows)
                    await connection.ExecuteAsync(
                        "insert into T (A,B) values (@A,@B)",
                        new { A = row["A"], B = row["B"] },
                        transaction: transaction
                    );

                await transaction.CommitAsync();
            }
        }

        await using (var connection = new SQLiteConnection(connectionString))
        {
            var result = await connection.QueryAsync("select * from T");
            Assert.Equal(5, result.Count());
        }
    }

    [Fact]
    public async Task SaveAsBasicCreateTest()
    {
        using var file = AutoDeletingPath.Create();
        var path = file.ToString();

        await _excelExporter.ExportAsync(path, new[]
        {
            new { Column1 = "MiniExcel", Column2 = 1 },
            new { Column1 = "Github", Column2 = 2}
        });

        await using (var stream = File.OpenRead(path))
        {
            var rows = await _excelImporter.QueryAsync(stream, hasHeaderRow: true).Cast<IDictionary<string, object>>().ToListAsync();
            Assert.Equal("MiniExcel", rows[0]["Column1"]);
            Assert.Equal(1d, rows[0]["Column2"]);
            Assert.Equal("Github", rows[1]["Column1"]);
            Assert.Equal(2d, rows[1]["Column2"]);
        }
        Assert.Equal("A1:B3", SheetHelper.GetFirstSheetDimensionRefValue(path));
    }

    [Fact]
    public async Task SaveAsBasicStreamTest()
    {
        {
            using var file = AutoDeletingPath.Create();
            var path = file.ToString();

            object[] values =
            [
                new { Column1 = "MiniExcel", Column2 = 1 },
                new { Column1 = "Github", Column2 = 2}
            ];

            await using (var stream = new FileStream(path, FileMode.CreateNew))
            {
                await _excelExporter.ExportAsync(stream, values);
            }

            await using (var stream = File.OpenRead(path))
            {
                var rows = await _excelImporter.QueryAsync(stream, hasHeaderRow: true)
                    .Cast<IDictionary<string, object>>()
                    .ToListAsync();

                Assert.Equal("MiniExcel", rows[0]["Column1"]);
                Assert.Equal(1d, rows[0]["Column2"]);
                Assert.Equal("Github", rows[1]["Column1"]);
                Assert.Equal(2d, rows[1]["Column2"]);
            }
        }
        {
            using var file = AutoDeletingPath.Create();
            var path = file.ToString();
            var values = new[]
            {
                new { Column1 = "MiniExcel", Column2 = 1 },
                new { Column1 = "Github", Column2 = 2}
            };

            await using (var stream = new MemoryStream())
            await using (var fileStream = new FileStream(path, FileMode.Create))
            {
                await _excelExporter.ExportAsync(stream, values);
                stream.Seek(0, SeekOrigin.Begin);
                await stream.CopyToAsync(fileStream);
            }

            await using (var stream = File.OpenRead(path))
            {
                var rows = await _excelImporter.QueryAsync(stream, hasHeaderRow: true)
                    .Cast<IDictionary<string, object>>()
                    .ToListAsync();

                Assert.Equal("MiniExcel", rows[0]["Column1"]);
                Assert.Equal(1d, rows[0]["Column2"]);
                Assert.Equal("Github", rows[1]["Column1"]);
                Assert.Equal(2d, rows[1]["Column2"]);
            }
        }
    }

    [Fact]
    public async Task SaveAsSpecialAndTypeCreateTest()
    {
        using var path = AutoDeletingPath.Create();
        await _excelExporter.ExportAsync(path.ToString(), new[]
        {
            new { a = @"""<>+-*//}{\\n", b = 1234567890, c = true, d = DateTime.Now },
            new { a = "<test>Hello World</test>", b = -1234567890, c = false, d = DateTime.Now.Date}
        });
        var info = new FileInfo(path.ToString());
        Assert.True(info.FullName == path.ToString());
    }

    [Fact]
    public async Task SaveAsFileEpplusCanReadTest()
    {
        using var path = AutoDeletingPath.Create();

        var now = DateTime.Now;
        var dt = new DateTime(now.Year, now.Month, now.Day, now.Hour, now.Minute, now.Second);

        await _excelExporter.ExportAsync(path.ToString(), new[]
        {
            new { a = @"""<>+-*//}{\\n", b = 1234567890, c = true, d = dt},
            new { a = "<test>Hello World</test>", b = -1234567890, c = false, d = dt.Date}
        });

        using var p = new ExcelPackage(path.ToString());
        var cells = p.Workbook.Worksheets[0].Cells;

        Assert.Equal("a", cells["A1"].Value.ToString());
        Assert.Equal("b", cells["B1"].Value.ToString());
        Assert.Equal("c", cells["C1"].Value.ToString());
        Assert.Equal("d", cells["D1"].Value.ToString());

        Assert.Equal(@"""<>+-*//}{\\n", cells["A2"].Value);
        Assert.Equal(1234567890, (double)cells["B2"].Value);
        Assert.True((bool)cells["C2"].Value);
        Assert.Equal(dt, (DateTime)cells["D2"].Value);

        Assert.Equal("<test>Hello World</test>", cells["A3"].Value);
        Assert.Equal(-1234567890, (double)cells["B3"].Value);
        Assert.False((bool)cells["C3"].Value);
        Assert.Equal(dt.Date, (DateTime)cells["D3"].Value);
    }

    [Fact]
    public async Task SavaAsClosedXmlCanReadTest()
    {
        var now = DateTime.Now;
        var dt = new DateTime(now.Year, now.Month, now.Day, now.Hour, now.Minute, now.Second);

        using var path = AutoDeletingPath.Create();
        await _excelExporter.ExportAsync(path.ToString(), new[]
        {
            new { a = @"""<>+-*//}{\\n", b = 1234567890, c = true, d = dt },
            new { a = "<test>Hello World</test>", b = -1234567890, c = false, d = dt.Date }
        });

        using var workbook = new XLWorkbook(path.ToString());
        var ws = workbook.Worksheets.First();

        Assert.Equal(@"""<>+-*//}{\\n", ws.Cell("A2").Value);
        Assert.Equal(1234567890, (double)ws.Cell("B2").Value);
        Assert.True((bool)ws.Cell("C2").Value);
        Assert.Equal(dt, ws.Cell("D2").Value);

        Assert.Equal("<test>Hello World</test>", ws.Cell("A3").Value);
        Assert.Equal(-1234567890, (double)ws.Cell("B3").Value);
        Assert.False((bool)ws.Cell("C3").Value);
        Assert.Equal(dt.Date, ws.Cell("D3").Value);
    }

    [Fact]
    public async Task ContentTypeUriContentTypeReadCheckTest()
    {
        var now = DateTime.Now;
        using var path = AutoDeletingPath.Create();

        await _excelExporter.ExportAsync(path.ToString(), new[]
        {
            new { a = @"""<>+-*//}{\\n", b = 1234567890, c = true, d = now},
            new { a = "<test>Hello World</test>", b = -1234567890, c = false, d = now.Date}
        });
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
    public async Task DynamicColumnsConfigurationIsUsedWhenCreatingExcelUsingIDataReader()
    {
        using var path = AutoDeletingPath.Create();
        var dateTime = DateTime.Now;
        var onlyDate = DateOnly.FromDateTime(dateTime);

        using var table = new DataTable();
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
        await using var reader = table.CreateDataReader();
        await _excelExporter.ExportAsync(path.ToString(), reader, configuration: configuration);

        await using var stream = File.OpenRead(path.ToString());
        var rows = await _excelImporter.QueryAsync(stream, hasHeaderRow: true)
            .Cast<IDictionary<string, object>>()
            .ToListAsync();

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
    public async Task DynamicColumnsConfigurationIsUsedWhenCreatingExcelUsingDataTable()
    {
        using var path = AutoDeletingPath.Create();
        var dateTime = DateTime.Now;
        var onlyDate = DateOnly.FromDateTime(dateTime);

        using var table = new DataTable();
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
        await _excelExporter.ExportAsync(path.ToString(), table, configuration: configuration);

        await using var stream = File.OpenRead(path.ToString());
        var rows = await _excelImporter.QueryAsync(stream, hasHeaderRow: true)
            .Cast<IDictionary<string, object>>()
            .ToListAsync();

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
    public async Task SaveAsByMiniExcelDataReader()
    {
        using var path1 = AutoDeletingPath.Create();
        var values = new List<Demo>
        {
            new() { Column1= "MiniExcel" ,Column2 = 1 },
            new() { Column1 = "Github", Column2 = 2 }
        };
        await _excelExporter.ExportAsync(path1.ToString(), values);

        using var path2 = AutoDeletingPath.Create();
        await using var reader = _excelImporter.GetDataReader(path1.ToString(), true);
        
        await _excelExporter.ExportAsync(path2.ToString(), reader);
        var results = await _excelImporter.QueryAsync<Demo>(path2.ToString()).ToListAsync();

        Assert.Equal(2, results.Count);
        Assert.Equal("MiniExcel", results.First().Column1);
        Assert.Equal(1, results.First().Column2);
        Assert.Equal("Github", results.Last().Column1);
        Assert.Equal(2, results.Last().Column2);
    }

    [Fact]
    public async Task InsertSheetTest()
    {
        var now = DateTime.Now;
        var dt = new DateTime(now.Year, now.Month, now.Day, now.Hour, now.Minute, now.Second);
        
        using var file = AutoDeletingPath.Create();
        var path = file.ToString();

        {
            using var table = new DataTable();
            table.Columns.Add("a", typeof(string));
            table.Columns.Add("b", typeof(decimal));
            table.Columns.Add("c", typeof(bool));
            table.Columns.Add("d", typeof(DateTime));
            table.Rows.Add(@"""<>+-*//}{\\n", 1234567890, true, dt);
            table.Rows.Add("<test>Hello World</test>", -1234567890, false, dt.Date);
            await _excelExporter.InsertSheetAsync(path, table, sheetName: "Sheet1");

            using var p = new ExcelPackage(path);
            var sheet1 = p.Workbook.Worksheets[0];

            Assert.Equal("Sheet1", sheet1.Name);
            Assert.Equal("a", sheet1.Cells["A1"].Value.ToString());
            Assert.Equal("b", sheet1.Cells["B1"].Value.ToString());
            Assert.Equal("c", sheet1.Cells["C1"].Value.ToString());
            Assert.Equal("d", sheet1.Cells["D1"].Value.ToString());

            Assert.Equal(@"""<>+-*//}{\\n", sheet1.Cells["A2"].Value);
            Assert.Equal(1234567890, (double)sheet1.Cells["B2"].Value);
            Assert.True((bool)sheet1.Cells["C2"].Value);
            Assert.Equal(dt, (DateTime)sheet1.Cells["D2"].Value);

            Assert.Equal("<test>Hello World</test>", sheet1.Cells["A3"].Value);
            Assert.Equal(-1234567890, (double)sheet1.Cells["B3"].Value);
            Assert.False((bool)sheet1.Cells["C3"].Value);
            Assert.Equal(dt.Date, (DateTime)sheet1.Cells["D3"].Value);
        }
        {
            using var table = new DataTable();
            table.Columns.Add("Column1", typeof(string));
            table.Columns.Add("Column2", typeof(int));
            table.Rows.Add("MiniExcel", 1);
            table.Rows.Add("Github", 2);

            await _excelExporter.InsertSheetAsync(path, table, sheetName: "Sheet2");
            using var p = new ExcelPackage(path);
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
            using var table = new DataTable();
            table.Columns.Add("Column1", typeof(string));
            table.Columns.Add("Column2", typeof(DateTime));
            table.Rows.Add("Test", dt);

            await _excelExporter.InsertSheetAsync(path, table, sheetName: "Sheet2", printHeader: false, configuration: new OpenXmlConfiguration
            {
                FastMode = true,
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

            using var p = new ExcelPackage(path);
            var sheet2 = p.Workbook.Worksheets[1];

            Assert.Equal("Sheet2", sheet2.Name);
            Assert.Equal("Test", sheet2.Cells["A1"].Value);
            Assert.Equal(dt.ToString("dd.MM.yyyy HH:mm:ss"), sheet2.Cells["B1"].Text );
        }
        {
            using var table = new DataTable();
            table.Columns.Add("Column1", typeof(string));
            table.Columns.Add("Column2", typeof(DateTime));
            table.Rows.Add("MiniExcel", dt);
            table.Rows.Add("Github", dt);

            await _excelExporter.InsertSheetAsync(path, table, sheetName: "Sheet3", configuration: new OpenXmlConfiguration
            {
                FastMode = true,
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

            using var p = new ExcelPackage(path);
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
    public async Task CopyAndInsertSheetTest()
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
            await _excelExporter.ExportAsync(firstPath, table);
        }
        {
            using var table = new DataTable();
            table.Columns.Add("Column1", typeof(string));
            table.Columns.Add("Column2", typeof(int));
            table.Rows.Add("MiniExcel", 1);
            table.Rows.Add("Github", 2);

            await _excelExporter.CopyAndAddSheetAsync(firstPath, secondPath, table, sheetName: "Sheet2");
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
            await _excelExporter.CopyAndAddSheetAsync(secondPath, thirdPath, new[] { new { Column1 = "Test", Column2 = dt } }, sheetName: "Sheet2", printHeader: false, configuration: new OpenXmlConfiguration
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
            Assert.Equal(dt.ToString("dd.MM.yyyy HH:mm:ss"), sheet2.Cells["B1"].Text );
        }
        {
            using var table = new DataTable();
            table.Columns.Add("Column1", typeof(string));
            table.Columns.Add("Column2", typeof(DateTime));
            table.Rows.Add("MiniExcel", dt);
            table.Rows.Add("Github", dt);
            await using var reader = table.CreateDataReader();

            await using var fs = File.OpenRead(thirdPath);
            await using var ms = new MemoryStream();

            await _excelExporter.CopyAndAddSheetAsync(fs, ms, reader, sheetName: "Sheet3", configuration: new OpenXmlConfiguration
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
    public async Task SaveAsByAsyncEnumerable()
    {
        using var path = AutoDeletingPath.Create();

        await _excelExporter.ExportAsync(path.ToString(), GetValues());
        var results = await _excelImporter.QueryAsync(path.ToString(), hasHeaderRow: true).ToListAsync();

        Assert.Equal(2, results.Count);
        Assert.Equal("MiniExcel", results[0].Column1);
        Assert.Equal(1, results[0].Column2);
        Assert.Equal("Github", results[^1].Column1);
        Assert.Equal(2, results[^1].Column2);
        return;

        static async IAsyncEnumerable<Demo> GetValues()
        {
            await Task.CompletedTask;
            yield return new Demo { Column1 = "MiniExcel", Column2 = 1 };
            await Task.CompletedTask;
            yield return new Demo { Column1 = "Github", Column2 = 2 };
        }
    }
    
    [Fact]
    public async Task ExportDataTableWithProgressTest()
    {
        var dataTable = new DataTable();
        dataTable.Columns.Add("Id", typeof(int));
        dataTable.Columns.Add("Name", typeof(string));
        dataTable.Columns.Add("Date", typeof(DateTime));
        dataTable.Rows.Add(1, "Alice", DateTime.Now);
        dataTable.Rows.Add(2, DBNull.Value, DateTime.UtcNow);
        dataTable.Rows.Add(3, "Alice", DateTime.Now.Date);

        var progress = new SimpleProgress();
        using var ms = new MemoryStream();
        var rowCounts = await _excelExporter.ExportAsync(ms, dataTable, progress: progress);
        Assert.Single(rowCounts);
        Assert.Equal(3, rowCounts.First());

        //Confirm the progress report is correct
        var cellCount = dataTable.Columns.Count * dataTable.Rows.Count;
        Assert.Equal(cellCount, progress.Value);

        ms.Seek(0, SeekOrigin.Begin);
        var resultDataTable = await _excelImporter.QueryAsDataTableAsync(ms, leaveOpen: true);

        //Confirm the data is correct
        Assert.Equal(dataTable.Rows.Count, resultDataTable.Rows.Count);
        Assert.Equal(dataTable.Columns.Count, resultDataTable.Columns.Count);
        for (var i = 0; i < dataTable.Rows.Count; i++)
        {
            for (var j = 0; j < dataTable.Columns.Count; j++)
            {
                //We compare string values because types change after writing and reading them back (e.g. int becomes double)
                Assert.Equal(dataTable.Rows[i][j].ToString(), resultDataTable.Rows[i][j].ToString());
            }
        }
    }

    [Fact]
    public async Task NumericFormattingWithMiniExcelFormatAttributeTest()
    {
        using var file = AutoDeletingPath.Create();
        var path = file.ToString();

        NumericFormattingTestDto[] testData =
        [
            new(currency: 1234.56m,
                alignedCurrency: 9876.54m,
                percentage: 0.85m,
                scientificNotation: 1234567890.123d,
                fixedDecimal: 42.123456m,
                phoneNumber: 5551234567,
                veryLongNumber: 155043269579349,
                customFormat: 999.999
            ),

            new(currency: -500.00m,
                alignedCurrency: -250.75m,
                percentage: 0.42m,
                scientificNotation: 987654321.456d,
                fixedDecimal: 15.5m,
                phoneNumber: 4155552671,
                veryLongNumber: 20573068629711152,
                customFormat: 100.012
            )
        ];

        await _excelExporter.ExportAsync(path, testData);

        using var package = new ExcelPackage(path);
        var cells = package.Workbook.Worksheets[0].Cells;

        // Verify headers
        Assert.Equal("Currency", cells["A1"].Value);
        Assert.Equal("AlignedCurrency", cells["B1"].Value);
        Assert.Equal("Percentage", cells["C1"].Value);
        Assert.Equal("ScientificNotation", cells["D1"].Value);
        Assert.Equal("FixedDecimal", cells["F1"].Value);
        Assert.Equal("PhoneNumber", cells["G1"].Value);
        Assert.Equal("VeryLongNumber", cells["H1"].Value);
        Assert.Equal("CustomFormat", cells["I1"].Value);

        // Verify first row of data
        Assert.Equal(1234.56, cells["A2"].Value);
        Assert.Equal("\"$\"#,##0.00", cells["A2"].Style.Numberformat.Format);

        Assert.Equal(9876.54, cells["B2"].Value);
        Assert.Equal("$#,##0.00_);($#,##0.00)", cells["B2"].Style.Numberformat.Format);

        Assert.Equal(0.85, cells["C2"].Value);
        Assert.Equal("0%", cells["C2"].Style.Numberformat.Format);

        Assert.Equal(1234567890.123, cells["D2"].Value);
        Assert.Equal("0.00E+00", cells["D2"].Style.Numberformat.Format);

        Assert.Equal(42.123456, cells["F2"].Value);
        Assert.Equal("0.000000", cells["F2"].Style.Numberformat.Format);

        Assert.Equal(5551234567, Convert.ToInt64(cells["G2"].Value));
        Assert.Equal("[<=9999999]###-####;(###) ###-####", cells["G2"].Style.Numberformat.Format);

        Assert.Equal(155043269579349, Convert.ToInt64(cells["H2"].Value));
        Assert.Equal("#", cells["H2"].Style.Numberformat.Format);

        Assert.Equal(999.999, cells["I2"].Value);
        Assert.Equal("0.000", cells["I2"].Style.Numberformat.Format);

        // Verify second row of data
        Assert.Equal(-500.00, cells["A3"].Value);
        Assert.Equal("\"$\"#,##0.00", cells["A3"].Style.Numberformat.Format);

        Assert.Equal(-250.75, cells["B3"].Value);
        Assert.Equal("$#,##0.00_);($#,##0.00)", cells["B3"].Style.Numberformat.Format);

        Assert.Equal(0.42, cells["C3"].Value);
        Assert.Equal("0%", cells["C3"].Style.Numberformat.Format);

        Assert.Equal(987654321.456, cells["D3"].Value);
        Assert.Equal("0.00E+00", cells["D3"].Style.Numberformat.Format);

        Assert.Equal(15.5, cells["F3"].Value);
        Assert.Equal("0.000000", cells["F3"].Style.Numberformat.Format);

        Assert.Equal(4155552671, Convert.ToInt64(cells["G3"].Value));
        Assert.Equal("[<=9999999]###-####;(###) ###-####", cells["G3"].Style.Numberformat.Format);

        Assert.Equal(20573068629711152, Convert.ToInt64(cells["H3"].Value));
        Assert.Equal("#", cells["H3"].Style.Numberformat.Format);

        Assert.Equal(100.012, cells["I3"].Value);
        Assert.Equal("0.000", cells["I3"].Style.Numberformat.Format);
    }

    [Fact]
    public async Task DateTimeFormattingWithMiniExcelFormatAttributeTest()
    {
        using var file = AutoDeletingPath.Create();
        var path = file.ToString();

        // Create fixed DateTime values for consistent testing
        var baseDate = new DateTime(2026, 5, 8, 14, 30, 45);
        var baseTime = new TimeSpan(14, 30, 45);

        DateTimeFormattingTestDto[] testData =
        [
            new(
                shortDate: baseDate,
                longDate: baseDate,
                dateWithTime: baseDate,
                timeOnly: baseTime,
                isoDateTime: baseDate,
                customDateTime: baseDate,
                monthYear: baseDate
            ),
            new(
                shortDate: new DateTime(2020, 12, 25),
                longDate: new DateTime(2020, 12, 25),
                dateWithTime: new DateTime(2020, 12, 25, 8, 15, 30),
                timeOnly: new TimeSpan(8, 15, 30),
                isoDateTime: new DateTime(2020, 12, 25, 8, 15, 30),
                customDateTime: new DateTime(2020, 12, 25, 8, 15, 30),
                monthYear: new DateTime(2020, 12, 25)
            )
        ];

        await _excelExporter.ExportAsync(path, testData);

        using var package = new ExcelPackage(path);
        var cells = package.Workbook.Worksheets[0].Cells;

        // Verify headers
        Assert.Equal("ShortDate", cells["A1"].Value);
        Assert.Equal("LongDate", cells["B1"].Value);
        Assert.Equal("DateWithTime", cells["C1"].Value);
        Assert.Equal("TimeOnly", cells["D1"].Value);
        Assert.Equal("IsoDateTime", cells["E1"].Value);
        Assert.Equal("CustomDateTime", cells["F1"].Value);
        Assert.Equal("MonthYear", cells["G1"].Value);

        // Verify first row
        Assert.Equal(baseDate, GetDateTime(cells["A2"].Value));
        Assert.Equal("mm/dd/yyyy", cells["A2"].Style.Numberformat.Format);

        // Long date format (dddd, mmmm dd, yyyy)
        Assert.Equal(baseDate, GetDateTime(cells["B2"].Value));
        Assert.Equal("dddd, mmmm dd, yyyy", cells["B2"].Style.Numberformat.Format);

        // Date with time (yyyy-mm-dd hh:mm:ss)
        Assert.Equal(baseDate, GetDateTime(cells["C2"].Value));
        Assert.Equal("yyyy-mm-dd hh:mm:ss", cells["C2"].Style.Numberformat.Format);

        // Time only format ([h]:mm:ss)
        Assert.Equal(baseTime, GetDateTime(cells["D2"].Value).TimeOfDay);
        Assert.Equal("[h]:mm:ss", cells["D2"].Style.Numberformat.Format);

        // ISO 8601 format (yyyy-mm-ddThh:mm:ss)
        Assert.Equal(baseDate, GetDateTime(cells["E2"].Value));
        Assert.Equal("yyyy-mm-dd\"T\"hh:mm:ss", cells["E2"].Style.Numberformat.Format);

        // Custom format (dd.mm.yyyy hh:mm)
        Assert.Equal(baseDate, GetDateTime(cells["F2"].Value));
        Assert.Equal("dd.mm.yyyy hh:mm", cells["F2"].Style.Numberformat.Format);

        // Month/Year format (mmmm yyyy)
        Assert.Equal(baseDate, GetDateTime(cells["G2"].Value));
        Assert.Equal("mmmm yyyy", cells["G2"].Style.Numberformat.Format);

        // Verify second row
        var secondRowDate = new DateTime(2020, 12, 25);
        var secondRowTime = new TimeSpan(8, 15, 30);

        Assert.Equal(secondRowDate, GetDateTime(cells["A3"].Value));
        Assert.Equal("mm/dd/yyyy", cells["A3"].Style.Numberformat.Format);

        Assert.Equal(secondRowDate, GetDateTime(cells["B3"].Value));
        Assert.Equal("dddd, mmmm dd, yyyy", cells["B3"].Style.Numberformat.Format);

        Assert.Equal(new DateTime(2020, 12, 25, 8, 15, 30), GetDateTime(cells["C3"].Value));
        Assert.Equal("yyyy-mm-dd hh:mm:ss", cells["C3"].Style.Numberformat.Format);

        Assert.Equal(secondRowTime, GetDateTime(cells["D3"].Value).TimeOfDay);
        Assert.Equal("[h]:mm:ss", cells["D3"].Style.Numberformat.Format);

        Assert.Equal(new DateTime(2020, 12, 25, 8, 15, 30), GetDateTime(cells["E3"].Value));
        Assert.Equal("yyyy-mm-dd\"T\"hh:mm:ss", cells["E3"].Style.Numberformat.Format);

        Assert.Equal(new DateTime(2020, 12, 25, 8, 15, 30), GetDateTime(cells["F3"].Value));
        Assert.Equal("dd.mm.yyyy hh:mm", cells["F3"].Style.Numberformat.Format);

        Assert.Equal(secondRowDate, GetDateTime(cells["G3"].Value));
        Assert.Equal("mmmm yyyy", cells["G3"].Style.Numberformat.Format);
        return;

        static DateTime GetDateTime(object value) => DateTime.FromOADate((double)value);
    }
    
    [Theory]
    [InlineData("")]
    [InlineData("it")]
    [InlineData("zh")]
    public async Task LocalizationTest(string cultureId)
    {
        var ogCulture = CultureInfo.CurrentUICulture;

        try
        {
            CultureInfo.CurrentUICulture = new CultureInfo(cultureId);

            await using var ms = new MemoryStream();
            await _excelExporter.ExportAsync(ms, Array.Empty<LocalizationSupportDto>());
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
    public async Task LocalizationTestDynamicColumns(string cultureId)
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

            await using var stream = new MemoryStream();
            await _excelExporter.ExportAsync(
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
