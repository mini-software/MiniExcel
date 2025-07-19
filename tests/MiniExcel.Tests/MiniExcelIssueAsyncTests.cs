using MiniExcelLib.Tests.Common.Utils;
using MiniExcelLib.Attributes;
using Newtonsoft.Json;

namespace MiniExcelLib.Tests;

public class MiniExcelIssueAsyncTests(ITestOutputHelper output)
{
    private readonly ITestOutputHelper _output = output;
    
    private readonly OpenXmlImporter _importer =  MiniExcel.Importer.GetExcelImporter();
    private readonly OpenXmlExporter _exporter =  MiniExcel.Exporter.GetExcelExporter();
    private readonly OpenXmlTemplater _templater =  MiniExcel.Templater.GetExcelTemplater();
    
    /// <summary>
    /// [SaveAsByTemplate support DateTime custom format · Issue #255 · mini-software/MiniExcel]
    /// (https://github.com/mini-software/MiniExcel/issues/255)
    /// </summary>
    [Fact]
    public async Task Issue255()
    {
        //tempalte
        {
            var templatePath = PathHelper.GetFile("xlsx/TestsIssue255_Template.xlsx");
            using var path = AutoDeletingPath.Create();
            var value = new
            {
                Issue255DTO = new[] 
                {
                    new Issue255DTO { Time = new DateTime(2021, 01, 01), Time2 = new DateTime(2021, 01, 01) }
                }
            };
            
            await  _templater.ApplyXlsxTemplateAsync(path.ToString(), templatePath, value);
            var q =  _importer.QueryExcelAsync(path.ToString()).ToBlockingEnumerable();
            var rows = q.ToList();
            
            Assert.Equal("2021", rows[1].A.ToString());
            Assert.Equal("2021", rows[1].B.ToString());
        }
        //saveas
        {
            using var path = AutoDeletingPath.Create();
            var value = new[] 
            {
                new Issue255DTO
                {
                    Time = new DateTime(2021, 01, 01),
                    Time2 = new DateTime(2021, 01, 01)
                }
            };
            var rowsWritten = await  _exporter.ExportExcelAsync(path.ToString(), value);
            Assert.Single(rowsWritten);
            Assert.Equal(1, rowsWritten[0]);
                
            var q =  _importer.QueryExcelAsync(path.ToString()).ToBlockingEnumerable();
            var rows = q.ToList();
            Assert.Equal("2021", rows[1].A.ToString());
            Assert.Equal("2021", rows[1].B.ToString());
        }
    }

    private class Issue255DTO
    {
        [MiniExcelFormat("yyyy")]
        public DateTime Time { get; set; }

        [MiniExcelColumn(Format = "yyyy")]
        public DateTime Time2 { get; set; }
    }

    /// <summary>
    /// [Dynamic QueryAsync custom format not using mapping format · Issue #256]
    /// (https://github.com/mini-software/MiniExcel/issues/256)
    /// </summary>
    [Fact]
    public async Task Issue256()
    {
        var path = PathHelper.GetFile("xlsx/TestIssue256.xlsx");
        var q =  _importer.QueryExcelAsync(path, false).ToBlockingEnumerable();
        var rows = q.ToList();
        
        Assert.Equal(new DateTime(2003, 4, 16), rows[1].A);
        Assert.Equal(new DateTime(2004, 4, 16), rows[1].B);
    }


    /// <summary>
    /// No error exception throw when reading xls file #242
    /// </summary>
    [Fact]
    public async Task Issue242()
    {
        var path = PathHelper.GetFile("xls/TestIssue242.xls");
        Assert.Throws<InvalidDataException>(() => _ =  _importer.QueryExcelAsync(path).ToBlockingEnumerable().ToList());

        await using var stream = File.OpenRead(path);
        Assert.Throws<InvalidDataException>(() => _ =  _importer.QueryExcelAsync(stream).ToBlockingEnumerable().ToList());
    }

    /// <summary>
    /// Csv type mapping QueryAsync error "cannot be converted to xxx type" #243
    /// </summary>
    [Fact]
    public async Task Issue243()
    {
        using var path = AutoDeletingPath.Create(ExcelType.Csv);
        var value = new[] 
        {
            new { Name ="Jack",Age=25,InDate=new DateTime(2021,01,03)},
            new { Name ="Henry",Age=36,InDate=new DateTime(2020,05,03)},
        };
        
        var rowsWritten = await  _exporter.ExportExcelAsync(path.ToString(), value);
        Assert.Single(rowsWritten);
        Assert.Equal(2, rowsWritten[0]);

        var q =  _importer.QueryExcelAsync<Issue243Dto>(path.ToString()).ToBlockingEnumerable();
        var rows = q.ToList();
        
        Assert.Equal("Jack", rows[0].Name);
        Assert.Equal(25, rows[0].Age);
        Assert.Equal(new DateTime(2021, 01, 03), rows[0].InDate);

        Assert.Equal("Henry", rows[1].Name);
        Assert.Equal(36, rows[1].Age);
        Assert.Equal(new DateTime(2020, 05, 03), rows[1].InDate);
    }

    private class Issue243Dto
    {
        public string Name { get; set; }
        public int Age { get; set; }
        public DateTime InDate { get; set; }
    }

    /// <summary>
    /// Support Custom Datetime format #241
    /// </summary>
    [Fact]
    public async Task Issue241()
    {
        Issue241Dto[] value =
        [
            new() { Name="Jack",InDate=new DateTime(2021,01,04) },
            new() { Name="Henry",InDate=new DateTime(2020,04,05) }
        ];

        // csv
        {
            using var file = AutoDeletingPath.Create(ExcelType.Csv);
            var path = file.ToString();
            var rowsWritten = await  _exporter.ExportExcelAsync(path, value);
            
            Assert.Single(rowsWritten);
            Assert.Equal(2, rowsWritten[0]);

            {
                var q =  _importer.QueryExcelAsync(path, true).ToBlockingEnumerable();
                var rows = q.ToList();
                
                Assert.Equal(rows[0].InDate, "01 04, 2021");
                Assert.Equal(rows[1].InDate, "04 05, 2020");
            }

            {
                var q =  _importer.QueryExcelAsync<Issue241Dto>(path).ToBlockingEnumerable();
                var rows = q.ToList();
                
                Assert.Equal(rows[0].InDate, new DateTime(2021, 01, 04));
                Assert.Equal(rows[1].InDate, new DateTime(2020, 04, 05));
            }
        }

        // xlsx
        {
            using var file = AutoDeletingPath.Create();
            var path = file.ToString();
            var rowsWritten = await  _exporter.ExportExcelAsync(path, value);
            
            Assert.Single(rowsWritten);
            Assert.Equal(2, rowsWritten[0]);

            {
                var q =  _importer.QueryExcelAsync(path, true).ToBlockingEnumerable();
                var rows = q.ToList();
                
                Assert.Equal(rows[0].InDate, "01 04, 2021");
                Assert.Equal(rows[1].InDate, "04 05, 2020");
            }

            {
                var q =  _importer.QueryExcelAsync<Issue241Dto>(path).ToBlockingEnumerable();
                var rows = q.ToList();
                
                Assert.Equal(rows[0].InDate, new DateTime(2021, 01, 04));
                Assert.Equal(rows[1].InDate, new DateTime(2020, 04, 05));
            }
        }
    }

    private class Issue241Dto
    {
        public string Name { get; set; }

        [MiniExcelFormat("MM dd, yyyy")]
        public DateTime InDate { get; set; }
    }

    /// <summary>
    /// SaveAs Default Template #132
    /// </summary>
    [Fact]
    public async Task Issue132()
    {
        {
            using var path = AutoDeletingPath.Create();
            var value = new[] 
            {
                new { Name ="Jack", Age=25, InDate=new DateTime(2021,01,03)},
                new { Name ="Henry", Age=36, InDate=new DateTime(2020,05,03)},
            };

            await  _exporter.ExportExcelAsync(path.ToString(), value);
        }

        {
            using var path = AutoDeletingPath.Create();
            var value = new[]
            {
                new { Name ="Jack", Age=25, InDate=new DateTime(2021,01,03)},
                new { Name ="Henry", Age=36, InDate=new DateTime(2020,05,03)},
            };
            var config = new OpenXmlConfiguration
            {
                TableStyles = TableStyles.None
            };
            var rowsWritten = await  _exporter.ExportExcelAsync(path.ToString(), value, configuration: config);
            
            Assert.Single(rowsWritten);
            Assert.Equal(2, rowsWritten[0]);
        }

        {
            using var path = AutoDeletingPath.Create();
            var value = JsonConvert.DeserializeObject<DataTable>(
                JsonConvert.SerializeObject(new[] 
                {
                    new { Name ="Jack", Age=25,InDate=new DateTime(2021,01,03)},
                    new { Name ="Henry", Age=36,InDate=new DateTime(2020,05,03)},
                })
            );
            var rowsWritten = await  _exporter.ExportExcelAsync(path.ToString(), value);
            
            Assert.Single(rowsWritten);
            Assert.Equal(2, rowsWritten[0]);
        }
    }

    /// <summary>
    /// Support SaveAs by DataSet #235
    /// </summary>
    [Fact]
    public async Task Issue235()
    {
        using var file = AutoDeletingPath.Create();
        var path = file.ToString();
        var sheets = new DataSet();

        var users = JsonConvert.DeserializeObject<DataTable>(JsonConvert.SerializeObject(new[] { new { Name = "Jack", Age = 25 }, new { Name = "Mike", Age = 44 } }));
        users.TableName = "users";
        sheets.Tables.Add(users);
        
        var department = JsonConvert.DeserializeObject<DataTable>(JsonConvert.SerializeObject(new[] { new { ID = "01", Name = "HR" }, new { ID = "02", Name = "IT" } }));
        department.TableName = "department";
        sheets.Tables.Add(department);

        var rowsWritten = await  _exporter.ExportExcelAsync(path, sheets);
        Assert.Equal(2, rowsWritten.Length);
        Assert.Equal(2, rowsWritten[0]);

        var sheetNames = await  _importer.GetSheetNamesAsync(path);
        Assert.Equal("users", sheetNames[0]);
        Assert.Equal("department", sheetNames[1]);

        {
            var q =  _importer.QueryExcelAsync(path, true, sheetName: "users").ToBlockingEnumerable();
            var rows = q.ToList();
            Assert.Equal("Jack", rows[0].Name);
            Assert.Equal(25, rows[0].Age);
            Assert.Equal("Mike", rows[1].Name);
            Assert.Equal(44, rows[1].Age);
        }
        {
            var q =  _importer.QueryExcelAsync(path, true, sheetName: "department").ToBlockingEnumerable();
            var rows = q.ToList();
            Assert.Equal("01", rows[0].ID);
            Assert.Equal("HR", rows[0].Name);
            Assert.Equal("02", rows[1].ID);
            Assert.Equal("IT", rows[1].Name);
        }
    }

    /// <summary>
    /// QueryAsDataTable A2=5.5 , A3=0.55/1.1 will case double type check error #233
    /// </summary>
    [Fact]
    public async Task Issue233()
    {
        var path = PathHelper.GetFile("xlsx/TestIssue233.xlsx");
            

        var dt = await  _importer.QueryExcelAsDataTableAsync(path);

            
        var rows = dt.Rows;
            
        Assert.Equal(0.55, rows[0]["Size"]);
        Assert.Equal("0.55/1.1", rows[1]["Size"]);
    }

    /// <summary>
    /// Csv QueryAsync split comma not correct #237
    /// https://github.com/mini-software/MiniExcel/issues/237
    /// </summary>
    [Fact]
    public async Task Issue237()
    {
        using var path = AutoDeletingPath.Create(ExcelType.Csv);
        var value = new[]
        {
            new{ id="\"\"1,2,3\"\""},
            new{ id="1,2,3"},
        };
        await  _exporter.ExportExcelAsync(path.ToString(), value);

        var q =  _importer.QueryExcelAsync(path.ToString(), true).ToBlockingEnumerable();
        var rows = q.ToList();

        Assert.Equal("\"\"1,2,3\"\"", rows[0].id);
        Assert.Equal("1,2,3", rows[1].id);
    }

    /// <summary>
    /// SaveAs support multiple sheets #234
    /// </summary>
    [Fact]
    public async Task Issue234()
    {
        using var file = AutoDeletingPath.Create();
        var path = file.ToString();
        
        var users = new[]
        {
            new { Name = "Jack", Age = 25 },
            new { Name = "Mike", Age = 44 }
        };
        var department = new[]
        {
            new { ID = "01", Name = "HR" },
            new { ID = "02", Name = "IT" }
        };
        var sheets = new Dictionary<string, object>
        {
            ["users"] = users,
            ["department"] = department
        };
        var rowsWritten = await  _exporter.ExportExcelAsync(path, sheets);
        
        Assert.Equal(2, rowsWritten.Length);
        Assert.Equal(2, rowsWritten[0]);


        var sheetNames =  _importer.GetSheetNames(path);
        Assert.Equal("users", sheetNames[0]);
        Assert.Equal("department", sheetNames[1]);

        {
            var q =  _importer.QueryExcelAsync(path, true, sheetName: "users").ToBlockingEnumerable();
            var rows = q.ToList();
            
            Assert.Equal("Jack", rows[0].Name);
            Assert.Equal(25, rows[0].Age);
            Assert.Equal("Mike", rows[1].Name);
            Assert.Equal(44, rows[1].Age);
        }
        {
            var q =  _importer.QueryExcelAsync(path, true, sheetName: "department").ToBlockingEnumerable();
            var rows = q.ToList();
            
            Assert.Equal("01", rows[0].ID);
            Assert.Equal("HR", rows[0].Name);
            Assert.Equal("02", rows[1].ID);
            Assert.Equal("IT", rows[1].Name);
        }
    }

    /// <summary>
    /// SaveAs By Reader Closed error : 'Error! Invalid attempt to call FieldCount when reader is closed' #230
    /// https://github.com/mini-software/MiniExcel/issues/230
    /// </summary>
    [Fact]
    public async Task Issue230()
    {
        await using var conn = Db.GetConnection("Data Source=:memory:");
        await conn.OpenAsync();
        await using var cmd = conn.CreateCommand();
        cmd.CommandText = "select 1 id union all select 2";
        
        await using (var reader = await cmd.ExecuteReaderAsync(CommandBehavior.CloseConnection))
        {
            while (await reader.ReadAsync())
            {
                for (int i = 0; i < reader.FieldCount; i++)
                {
                    var result = $"{reader.GetName(i)} , {reader.GetValue(i)}";
                    _output.WriteLine(result);
                }
            }
        }

        await using var conn2 = Db.GetConnection("Data Source=:memory:");
        await conn2.OpenAsync();
        await using var cmd2 = conn2.CreateCommand();
        cmd2.CommandText = "select 1 id union all select 2";
        
        await using (var reader = await cmd2.ExecuteReaderAsync(CommandBehavior.CloseConnection))
        {
            while (await reader.ReadAsync())
            {
                for (int i = 0; i < reader.FieldCount; i++)
                {
                    var result = $"{reader.GetName(i)} , {reader.GetValue(i)}";
                    _output.WriteLine(result);
                }
            }
        }

        await using var conn3 = Db.GetConnection("Data Source=:memory:");
        await conn3.OpenAsync();
        await using var cmd3 = conn3.CreateCommand();
        cmd3.CommandText = "select 1 id union all select 2";
        
        await using (var reader = await cmd3.ExecuteReaderAsync(CommandBehavior.CloseConnection))
        {
            using var path = AutoDeletingPath.Create();
            var rowsWritten = await  _exporter.ExportExcelAsync(path.ToString(), reader, printHeader: true);
            
            Assert.Single(rowsWritten);
            Assert.Equal(2, rowsWritten[0]);
                
            var q =  _importer.QueryExcelAsync(path.ToString(), true).ToBlockingEnumerable();
            var rows = q.ToList();
            
            Assert.Equal(1, rows[0].id);
            Assert.Equal(2, rows[1].id);
        }
    }

    /// <summary>
    /// v0.14.3 QueryAsDataTable error "Cannot set Column to be null" #229
    /// https://github.com/mini-software/MiniExcel/issues/229
    /// </summary>
    [Fact]
    public async Task Issue229()
    {
        var path = PathHelper.GetFile("xlsx/TestIssue229.xlsx");
            
        var dt = await  _importer.QueryExcelAsDataTableAsync(path);
            
        foreach (DataColumn column in dt.Columns)
        {
            var v = dt.Rows[3][column];
            Assert.Equal(DBNull.Value, v);
        }
    }

    /// <summary>
    /// [QueryAsync Merge cells data · Issue #122 · mini-software/MiniExcel]
    /// (https://github.com/mini-software/MiniExcel/issues/122)
    /// </summary>
    [Fact]
    public Task Issue122()
    {
        var config = new OpenXmlConfiguration
        {
            FillMergedCells = true
        };
            
        var path1 = PathHelper.GetFile("xlsx/TestIssue122.xlsx");
        var rows1 =  _importer.QueryExcelAsync(path1, useHeaderRow: true, configuration: config).ToBlockingEnumerable().ToList();
        
        Assert.Equal("HR", rows1[0].Department);
        Assert.Equal("HR", rows1[1].Department);
        Assert.Equal("HR", rows1[2].Department);
        Assert.Equal("IT", rows1[3].Department);
        Assert.Equal("IT", rows1[4].Department);
        Assert.Equal("IT", rows1[5].Department);

        var path2 = PathHelper.GetFile("xlsx/TestIssue122_2.xlsx");
        var rows2 =  _importer.QueryExcelAsync(path2, useHeaderRow: true, configuration: config).ToBlockingEnumerable().ToList();

        Assert.Equal("V1", rows2[2].Test1);
        Assert.Equal("V2", rows2[5].Test2);
        Assert.Equal("V3", rows2[1].Test3);
        Assert.Equal("V4", rows2[2].Test4);
        Assert.Equal("V5", rows2[3].Test5);
        Assert.Equal("V6", rows2[5].Test5);
        
        return Task.CompletedTask;
    }

    /// <summary>
    /// [Support Xlsm AutoCheck · Issue #227 · mini-software/MiniExcel]
    /// (https://github.com/mini-software/MiniExcel/issues/227)
    /// </summary>
    [Fact]
    public async Task Issue227()
    {
        {
            var path = PathHelper.GetTempPath("xlsm");
            Assert.Throws<NotSupportedException>(() =>  _exporter.ExportExcel(path, new[] { new { V = "A1" }, new { V = "A2" } }));
            File.Delete(path);
        }

        {
            var path = PathHelper.GetFile("xlsx/TestIssue227.xlsm");
            {
                var q =  _importer.QueryExcelAsync<MiniExcelOpenXmlTests.UserAccount>(path).ToBlockingEnumerable();
                var rows = q.ToList();
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
                await using var stream = File.OpenRead(path);
                var q =  _importer.QueryExcelAsync<MiniExcelOpenXmlTests.UserAccount>(stream).ToBlockingEnumerable();
                var rows = q.ToList();
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
    }

    /// <summary>
    /// https://github.com/mini-software/MiniExcel/issues/226
    /// Fix SaveAsByTemplate single column demension index error #226
    /// </summary>
    [Fact]
    public async Task Issue226()
    {
        using var path = AutoDeletingPath.Create();
        var templatePath = PathHelper.GetFile("xlsx/TestIssue226.xlsx");
        await  _templater.ApplyXlsxTemplateAsync(path.ToString(), templatePath, new { employees = new[] { new { name = "123" }, new { name = "123" } } });
        Assert.Equal("A1:A3", SheetHelper.GetFirstSheetDimensionRefValue(path.ToString()));
    }

    /// <summary>
    /// ASP.NET Webform gridview datasource can't use miniexcel queryasdatatable · Issue #223]
    /// (https://github.com/mini-software/MiniExcel/issues/223)
    /// </summary>
    [Fact]
    public async Task Issue223()
    {
        List<Dictionary<string, object?>> value =
        [
            new() { { "A", null }, { "B", null } },
            new() { { "A", 123 }, { "B", new DateTime(2021, 1, 1) } },
            new() { { "A", Guid.NewGuid() }, { "B", "HelloWorld" } }
        ];
        using var path = AutoDeletingPath.Create();
        var rowsWritten = await  _exporter.ExportExcelAsync(path.ToString(), value);
        Assert.Single(rowsWritten);
        Assert.Equal(3, rowsWritten[0]);


        var dt = await  _importer.QueryExcelAsDataTableAsync(path.ToString());
#pragma warning restore CS0618
        var columns = dt.Columns;
        Assert.Equal(typeof(object), columns[0].DataType);
        Assert.Equal(typeof(object), columns[1].DataType);

        Assert.Equal(123.0, dt.Rows[1]["A"]);
        Assert.Equal("HelloWorld", dt.Rows[2]["B"]);
    }

    /// <summary>
    /// [Custom yyyy-MM-dd format not convert datetime · Issue #222]
    /// (https://github.com/mini-software/MiniExcel/issues/222)
    /// </summary>
    [Fact]
    public async Task Issue222()
    {
        var path = PathHelper.GetFile("xlsx/TestIssue222.xlsx");
        var q =  _importer.QueryExcelAsync(path).ToBlockingEnumerable();
        var rows = q.ToList();
        Assert.Equal(typeof(DateTime), rows[1].A.GetType());
        Assert.Equal(new DateTime(2021, 4, 29), rows[1].A);
    }

    /// <summary>
    /// QueryAsync Support StartCell #147
    /// https://github.com/mini-software/MiniExcel/issues/147
    /// </summary>
    [Fact]
    public async Task Issue147()
    {
        {
            var path = PathHelper.GetFile("xlsx/TestIssue147.xlsx");
            var q =  _importer.QueryExcelAsync(path, useHeaderRow: false, startCell: "C3", sheetName: "Sheet1").ToBlockingEnumerable();
            var rows = q.ToList();
            
            Assert.Equal(["C", "D", "E"], (rows[0] as IDictionary<string, object>)?.Keys);
            Assert.Equal(["Column1", "Column2", "Column3"], new[] { rows[0].C as string, rows[0].D as string, rows[0].E as string });
            Assert.Equal(["C4", "D4", "E4"], new[] { rows[1].C as string, rows[1].D as string, rows[1].E as string });
            Assert.Equal(["C9", "D9", "E9"], new[] { rows[6].C as string, rows[6].D as string, rows[6].E as string });
            Assert.Equal(["C12", "D12", "E12"], new[] { rows[9].C as string, rows[9].D as string, rows[9].E as string });
            Assert.Equal(["C13", "D13", "E13"], new[] { rows[10].C as string, rows[10].D as string, rows[10].E as string });
            
            foreach (var i in new[] { 4, 5, 7, 8 })
            {
                Assert.Equal([null, null, null], new[] { rows[i].C as string, rows[i].D as string, rows[i].E as string });
            }
            Assert.Equal(11, rows.Count);

            var columns = await  _importer.GetExcelColumnsAsync(path, startCell: "C3");
            Assert.Equal(["C", "D", "E"], columns);
        }

        {
            var path = PathHelper.GetFile("xlsx/TestIssue147.xlsx");
            var q =  _importer.QueryExcelAsync(path, useHeaderRow: true, startCell: "C3", sheetName: "Sheet1").ToBlockingEnumerable();
            var rows = q.ToList();
            
            Assert.Equal(["Column1", "Column2", "Column3"], (rows[0] as IDictionary<string, object>)?.Keys);
            Assert.Equal(["C4", "D4", "E4"], new[] { rows[0].Column1 as string, rows[0].Column2 as string, rows[0].Column3 as string });
            Assert.Equal(["C9", "D9", "E9"], new[] { rows[5].Column1 as string, rows[5].Column2 as string, rows[5].Column3 as string });
            Assert.Equal(["C12", "D12", "E12"], new[] { rows[8].Column1 as string, rows[8].Column2 as string, rows[8].Column3 as string });
            Assert.Equal(["C13", "D13", "E13"], new[] { rows[9].Column1 as string, rows[9].Column2 as string, rows[9].Column3 as string });
            
            foreach (var i in new[] { 3, 4, 6, 7 })
            {
                Assert.Equal([null, null, null], new[] { rows[i].Column1 as string, rows[i].Column2 as string, rows[i].Column3 as string });
            }
            Assert.Equal(10, rows.Count);
                
            var columns = await  _importer.GetExcelColumnsAsync(path, useHeaderRow: true, startCell: "C3");
            Assert.Equal(["Column1", "Column2", "Column3"], columns);
        }
    }


    /// <summary>
    /// [Can SaveAs support iDataReader export to avoid the dataTable consuming too much memory · Issue #211 · mini-software/MiniExcel]
    /// (https://github.com/mini-software/MiniExcel/issues/211)
    /// </summary>
    [Fact]
    public async Task Issue211()
    {
        using var path = AutoDeletingPath.Create();
        using var tempSqlitePath = AutoDeletingPath.Create(Path.GetTempPath(), $"{Guid.NewGuid()}.db");
        var connectionString = $"Data Source={tempSqlitePath};Version=3;";

        await using var connection = new SQLiteConnection(connectionString);
        using var reader = await connection.ExecuteReaderAsync("select 1 Test1,2 Test2 union all select 3 , 4 union all select 5 ,6");

        var rowsWritten = await  _exporter.ExportExcelAsync(path.ToString(), reader);
        Assert.Single(rowsWritten);
        Assert.Equal(3, rowsWritten[0]);

        var q =  _importer.QueryExcelAsync(path.ToString(), true).ToBlockingEnumerable();
        var rows = q.ToList();
        Assert.Equal(1.0, rows[0].Test1);
        Assert.Equal(2.0, rows[0].Test2);
        Assert.Equal(3.0, rows[1].Test1);
        Assert.Equal(4.0, rows[1].Test2);
    }

    [Fact]
    public async Task EmptyDataReaderIssue()
    {
        using var path = AutoDeletingPath.Create();
        using var tempSqlitePath = AutoDeletingPath.Create(Path.GetTempPath(), $"{Guid.NewGuid()}.db");
        var connectionString = $"Data Source={tempSqlitePath};Version=3;";

        await using (var connection1 = new SQLiteConnection(connectionString))
        {
            await connection1.ExecuteAsync("CREATE TABLE test (id int PRIMARY KEY, name TEXT)");
        }

        await using var connection2 = new SQLiteConnection(connectionString);
        using var reader = await connection2.ExecuteReaderAsync("SELECT * FROM test");

        var rowsWritten = await  _exporter.ExportExcelAsync(path.ToString(), reader);
        Assert.Single(rowsWritten);
        Assert.Equal(0, rowsWritten[0]);

        var q =  _importer.QueryExcelAsync(path.ToString(), true).ToBlockingEnumerable();
        var rows = q.ToList();
        Assert.Empty(rows);
    }

    /// <summary>
    /// [When reading Excel, can return IDataReader and DataTable to facilitate the import of database. Like ExcelDataReader provide reader.AsDataSet() · Issue #216 · mini-software/MiniExcel](https://github.com/mini-software/MiniExcel/issues/216)
    /// </summary>
    [Fact]
    public async Task Issue216()
    {
        using var path = AutoDeletingPath.Create();
        var value = new[] { new { Test1 = "1", Test2 = 2 }, new { Test1 = "3", Test2 = 4 } };
        var rowsWritten = await  _exporter.ExportExcelAsync(path.ToString(), value);
        Assert.Single(rowsWritten);
        Assert.Equal(2, rowsWritten[0]);

        {
            var table = await  _importer.QueryExcelAsDataTableAsync(path.ToString());
            Assert.Equal("Test1", table.Columns[0].ColumnName);
            Assert.Equal("Test2", table.Columns[1].ColumnName);
            Assert.Equal("1", table.Rows[0]["Test1"]);
            Assert.Equal(2.0, table.Rows[0]["Test2"]);
            Assert.Equal("3", table.Rows[1]["Test1"]);
            Assert.Equal(4.0, table.Rows[1]["Test2"]);
        }

        {
            var dt = await  _importer.QueryExcelAsDataTableAsync(path.ToString(), false);
            Assert.Equal("Test1", dt.Rows[0]["A"]);
            Assert.Equal("Test2", dt.Rows[0]["B"]);
            Assert.Equal("1", dt.Rows[1]["A"]);
            Assert.Equal(2.0, dt.Rows[1]["B"]);
            Assert.Equal("3", dt.Rows[2]["A"]);
            Assert.Equal(4.0, dt.Rows[2]["B"]);
        }
    }

    /// <summary>
    /// https://gitee.com/dotnetchina/MiniExcel/issues/I3OSKV
    /// When exporting, the pure numeric string will be forcibly converted to a numeric type, resulting in the loss of the end data
    /// </summary>
    [Fact]
    public async Task IssueI3OSKV()
    {
        {
            using var path = AutoDeletingPath.Create();
            var value = new[] { new { Test = "12345678901234567890" } };
            await  _exporter.ExportExcelAsync(path.ToString(), value);

            var q =  _importer.QueryExcelAsync(path.ToString(), true).ToBlockingEnumerable();
            var A2 = q.First().Test;
            Assert.Equal("12345678901234567890", A2);
        }

        {
            using var path = AutoDeletingPath.Create();
            var value = new[] { new { Test = 123456.789 } };
            await  _exporter.ExportExcelAsync(path.ToString(), value);

            var q =  _importer.QueryExcelAsync(path.ToString(), true).ToBlockingEnumerable();
            var A2 = q.First().Test;
            Assert.Equal(123456.789, A2);
        }
    }

    /// <summary>
    /// [Dynamic QueryAsync can't summary numeric cell value default, need to cast · Issue #220 · mini-software/MiniExcel]
    /// (https://github.com/mini-software/MiniExcel/issues/220)
    /// </summary>
    [Fact]
    public async Task Issue220()
    {
        var path = PathHelper.GetFile("xlsx/TestIssue220.xlsx");
        var rows =  _importer.QueryExcelAsync(path, useHeaderRow: true).ToBlockingEnumerable();
        var result = rows
            .GroupBy(s => s.PRT_ID)
            .Select(g => new
            {
                PRT_ID = g.Key,
                Apr = g.Sum(d => (double?)d.Apr),
                May = g.Sum(d => (double?)d.May),
                Jun = g.Sum(d => (double?)d.Jun),
            })
            .ToList();
        
        Assert.Equal(91843.25, result[0].Jun);
        Assert.Equal(50000.99, result[1].Jun);
    }

    /// <summary>
    /// Optimize stream excel type check
    /// https://github.com/mini-software/MiniExcel/issues/215
    /// </summary>
    [Fact]
    public async Task Issue215()
    {
        await using var stream = new MemoryStream();
        await  _exporter.ExportExcelAsync(stream, new[] { new { V = "test1" }, new { V = "test2" } });
        
        var q =  _importer.QueryExcelAsync(stream, true).ToBlockingEnumerable().Cast<IDictionary<string, object>>();
        var rows = q.ToList();
        
        Assert.Equal("test1", rows[0]["V"]);
        Assert.Equal("test2", rows[1]["V"]);
    }

    /// <summary>
    /// DataTable recommended to use Caption for column name first, then use columname
    /// https://github.com/mini-software/MiniExcel/issues/217
    /// </summary>
    [Fact]
    public async Task Issue217()
    {
        using var table = new DataTable();
        table.Columns.Add("CustomerID");
        table.Columns.Add("CustomerName").Caption = "Name";
        table.Columns.Add("CreditLimit").Caption = "Limit";
        table.Rows.Add(1, "Jonathan", 23.44);
        table.Rows.Add(2, "Bill", 56.87);

        // openxml
        {
            using var path = AutoDeletingPath.Create();
            var rowsWritten = await  _exporter.ExportExcelAsync(path.ToString(), table);
            Assert.Single(rowsWritten);
            Assert.Equal(2, rowsWritten[0]);
                
            var q =  _importer.QueryExcelAsync(path.ToString()).ToBlockingEnumerable();
            var rows = q.ToList();
            Assert.Equal("Name", rows[0].B);
            Assert.Equal("Limit", rows[0].C);
        }

        // csv
        {
            using var path = AutoDeletingPath.Create(ExcelType.Csv);
            await  _exporter.ExportExcelAsync(path.ToString(), table);

            var q =  _importer.QueryExcelAsync(path.ToString()).ToBlockingEnumerable();
            var rows = q.ToList();
            Assert.Equal("Name", rows[0].B);
            Assert.Equal("Limit", rows[0].C);
        }
    }

    /// <summary>
    /// _ _exporter.ExportXlsx(path, table,sheetName:“Name”) ，the actual sheetName is Sheet1
    /// https://github.com/mini-software/MiniExcel/issues/212
    /// </summary>
    [Fact]
    public async Task Issue212()
    {
        const string sheetName = "Demo";
        
        using var path = AutoDeletingPath.Create();
        await  _exporter.ExportExcelAsync(path.ToString(), new[] { new { x = 1, y = 2 } }, sheetName: sheetName);

        var actualSheetName =  _importer.GetSheetNames(path.ToString()).ToList()[0];
        Assert.Equal(sheetName, actualSheetName);
    }

    /// <summary>
    /// Version &lt;= v0.13.1 Template merge row list rendering has no merge
    /// https://github.com/mini-software/MiniExcel/issues/207
    /// </summary>
    [Fact]
    public async Task Issue207()
    {
        {
            const string templatePath = "../../../../../samples/xlsx/TestIssue207_2.xlsx";
            using var file = AutoDeletingPath.Create();
            var path = file.ToString();
            
            var value = new
            {
                project = new[] 
                {
                    new {name = "項目1", content="[]內容1,[]內容2,[]內容3,[]內容4,[]內容5"},
                    new {name = "項目2", content="[]內容1,[]內容2,[]內容3,[]內容4,[]內容5"},
                    new {name = "項目3", content="[]內容1,[]內容2,[]內容3,[]內容4,[]內容5"},
                    new {name = "項目4", content="[]內容1,[]內容2,[]內容3,[]內容4,[]內容5"}
                }
            };

            await  _templater.ApplyXlsxTemplateAsync(path, templatePath, value);
            var q =  _importer.QueryExcelAsync(path).ToBlockingEnumerable();
            var rows = q.ToList();

            Assert.Equal("項目1", rows[0].A);
            Assert.Equal("[]內容1,[]內容2,[]內容3,[]內容4,[]內容5", rows[0].B);
            Assert.Equal("項目2", rows[2].A);
            Assert.Equal("[]內容1,[]內容2,[]內容3,[]內容4,[]內容5", rows[2].B);
            Assert.Equal("項目3", rows[4].A);
            Assert.Equal("[]內容1,[]內容2,[]內容3,[]內容4,[]內容5", rows[4].B);
            Assert.Equal("項目4", rows[6].A);
            Assert.Equal("[]內容1,[]內容2,[]內容3,[]內容4,[]內容5", rows[6].B);

            Assert.Equal("Test1", rows[8].A);
            Assert.Equal("Test2", rows[8].B);
            Assert.Equal("Test3", rows[8].C);

            Assert.Equal("項目1", rows[12].A);
            Assert.Equal("[]內容1,[]內容2,[]內容3,[]內容4,[]內容5", rows[12].B);
            Assert.Equal("項目2", rows[13].A);
            Assert.Equal("[]內容1,[]內容2,[]內容3,[]內容4,[]內容5", rows[13].B);
            Assert.Equal("項目3", rows[14].A);
            Assert.Equal("[]內容1,[]內容2,[]內容3,[]內容4,[]內容5", rows[14].B);
            Assert.Equal("項目4", rows[15].A);
            Assert.Equal("[]內容1,[]內容2,[]內容3,[]內容4,[]內容5", rows[15].B);

            var demension = SheetHelper.GetFirstSheetDimensionRefValue(path);
            Assert.Equal("A1:C16", demension);
        }

        {
            const string templatePath = "../../../../../samples/xlsx/TestIssue207_Template_Merge_row_list_rendering_without_merge/template.xlsx";
            using var file = AutoDeletingPath.Create();
            var path = file.ToString();
            
            var value = new
            {
                project = new[] 
                {
                    new {name = "項目1", content="[]內容1,[]內容2,[]內容3,[]內容4,[]內容5"},
                    new {name = "項目2", content="[]內容1,[]內容2,[]內容3,[]內容4,[]內容5"},
                    new {name = "項目3", content="[]內容1,[]內容2,[]內容3,[]內容4,[]內容5"},
                    new {name = "項目4", content="[]內容1,[]內容2,[]內容3,[]內容4,[]內容5"}
                }
            };

            await  _templater.ApplyXlsxTemplateAsync(path, templatePath, value);

            var q =  _importer.QueryExcelAsync(path).ToBlockingEnumerable();
            var rows = q.ToList();
            Assert.Equal("項目1", rows[0].A);
            Assert.Equal("[]內容1,[]內容2,[]內容3,[]內容4,[]內容5", rows[0].C);
            Assert.Equal("項目2", rows[3].A);
            Assert.Equal("[]內容1,[]內容2,[]內容3,[]內容4,[]內容5", rows[3].C);
            Assert.Equal("項目3", rows[6].A);
            Assert.Equal("[]內容1,[]內容2,[]內容3,[]內容4,[]內容5", rows[6].C);
            Assert.Equal("項目4", rows[9].A);
            Assert.Equal("[]內容1,[]內容2,[]內容3,[]內容4,[]內容5", rows[9].C);
            
            var demension = SheetHelper.GetFirstSheetDimensionRefValue(path);
            Assert.Equal("A1:E15", demension);
        }
    }

    /// <summary>
    /// https://github.com/mini-software/MiniExcel/issues/87
    /// </summary>
    [Fact]
    public async Task Issue87()
    {
        const string templatePath = "../../../../../samples/xlsx/TestTemplateCenterEmpty.xlsx";
        using var path = AutoDeletingPath.Create();
        var value = new
        {
            Tests = Enumerable.Range(1, 5).Select((s, i) => new { test1 = i, test2 = i })
        };

        await using var stream = File.OpenRead(templatePath);
        var q =  _importer.QueryExcelAsync(templatePath).ToBlockingEnumerable();
        var rows = q.ToList();

        await  _templater.ApplyXlsxTemplateAsync(path.ToString(), templatePath, value);
    }

    /// <summary>
    /// https://github.com/mini-software/MiniExcel/issues/206
    /// </summary>
    [Fact]
    public async Task Issue206()
    {
        const string templatePath = "../../../../../samples/xlsx/TestTemplateBasicIEmumerableFill.xlsx";
        {
            using var path = AutoDeletingPath.Create();

            var dt = new DataTable();
            {
                dt.Columns.Add("name");
                dt.Columns.Add("department");
            }
            var value = new Dictionary<string, object>
            {
                ["employees"] = dt
            };
            await  _templater.ApplyXlsxTemplateAsync(path.ToString(), templatePath, value);

            var dimension = SheetHelper.GetFirstSheetDimensionRefValue(path.ToString());
            Assert.Equal("A1:B2", dimension);
        }

        {
            using var path = AutoDeletingPath.Create();

            var dt = new DataTable();
            {
                dt.Columns.Add("name");
                dt.Columns.Add("department");
                dt.Rows.Add("Jack", "HR");
            }
            var value = new Dictionary<string, object> { ["employees"] = dt };
            await  _templater.ApplyXlsxTemplateAsync(path.ToString(), templatePath, value);

            var dimension = SheetHelper.GetFirstSheetDimensionRefValue(path.ToString());
            Assert.Equal("A1:B2", dimension);
        }
    }


    /// <summary>
    /// https://github.com/mini-software/MiniExcel/issues/193
    /// </summary>
    [Fact]
    public async Task Issue193()
    {
        {
            const string templatePath = "../../../../../samples/xlsx/TestTemplateComplexWithNamespacePrefix.xlsx";
            using var file = AutoDeletingPath.Create();
            var path = file.ToString();

            // 1. By Class
            var value = new
            {
                title = "FooCompany",
                managers = new[] 
                {
                    new {name="Jack",department="HR"},
                    new {name="Loan",department="IT"}
                },
                employees = new[] 
                {
                    new {name="Wade",department="HR"},
                    new {name="Felix",department="HR"},
                    new {name="Eric",department="IT"},
                    new {name="Keaton",department="IT"}
                }
            };
            await  _templater.ApplyXlsxTemplateAsync(path, templatePath, value);

            foreach (var sheetName in  _importer.GetSheetNames(path))
            {
                var q =  _importer.QueryExcelAsync(path, sheetName: sheetName).ToBlockingEnumerable();
                var rows = q.ToList();
                Assert.Equal(9, rows.Count);

                Assert.Equal("FooCompany", rows[0].A);
                Assert.Equal("Jack", rows[2].B);
                Assert.Equal("HR", rows[2].C);
                Assert.Equal("Loan", rows[3].B);
                Assert.Equal("IT", rows[3].C);

                Assert.Equal("Wade", rows[5].B);
                Assert.Equal("HR", rows[5].C);
                Assert.Equal("Felix", rows[6].B);
                Assert.Equal("HR", rows[6].C);

                Assert.Equal("Eric", rows[7].B);
                Assert.Equal("IT", rows[7].C);
                Assert.Equal("Keaton", rows[8].B);
                Assert.Equal("IT", rows[8].C);

                var demension = SheetHelper.GetFirstSheetDimensionRefValue(path);
                Assert.Equal("A1:C9", demension);

                //TODO:row can't contain xmlns
                //![image](https://user-images.githubusercontent.com/12729184/114998840-ead44500-9ed3-11eb-8611-58afb98faed9.png)

            }
        }

        {
            const string templatePath = "../../../../../samples/xlsx/TestTemplateComplex.xlsx";
            using var file = AutoDeletingPath.Create();
            var path = file.ToString();

            // 2. By Dictionary
            var value = new Dictionary<string, object>
            {
                ["title"] = "FooCompany",
                ["managers"] = new[] 
                {
                    new {name="Jack",department="HR"},
                    new {name="Loan",department="IT"}
                },
                ["employees"] = new[] 
                {
                    new {name="Wade",department="HR"},
                    new {name="Felix",department="HR"},
                    new {name="Eric",department="IT"},
                    new {name="Keaton",department="IT"}
                }
            };
            await  _templater.ApplyXlsxTemplateAsync(path, templatePath, value);

            var q =  _importer.QueryExcelAsync(path).ToBlockingEnumerable();
            var rows = q.ToList();
            Assert.Equal("FooCompany", rows[0].A);
            Assert.Equal("Jack", rows[2].B);
            Assert.Equal("HR", rows[2].C);
            Assert.Equal("Loan", rows[3].B);
            Assert.Equal("IT", rows[3].C);

            Assert.Equal("Wade", rows[5].B);
            Assert.Equal("HR", rows[5].C);
            Assert.Equal("Felix", rows[6].B);
            Assert.Equal("HR", rows[6].C);

            Assert.Equal("Eric", rows[7].B);
            Assert.Equal("IT", rows[7].C);
            Assert.Equal("Keaton", rows[8].B);
            Assert.Equal("IT", rows[8].C);

            var demension = SheetHelper.GetFirstSheetDimensionRefValue(path);
            Assert.Equal("A1:C9", demension);
        }
    }

    [Fact]
    public async Task Issue142_Query()
    {
        const string path = "../../../../../samples/xlsx/TestIssue142.xlsx";
        const string pathCsv = "../../../../../samples/xlsx/TestIssue142.csv";
        {
            var rows =  _importer.QueryExcelAsync<Issue142VoExcelColumnNameNotFound>(path).ToBlockingEnumerable().ToList();
            Assert.Equal(0, rows[0].MyProperty1);
        }

        {
            await Assert.ThrowsAsync<ArgumentException>(async () =>
            {
                var q =  _importer.QueryExcelAsync<Issue142VoOverIndex>(path).ToBlockingEnumerable().ToList();
            });
        }

        {
            var q =  _importer.QueryExcelAsync<Issue142VO>(path).ToBlockingEnumerable();
            var rows = q.ToList();
            Assert.Equal("CustomColumnName", rows[0].MyProperty1);
            Assert.Null(rows[0].MyProperty7);
            Assert.Equal("MyProperty2", rows[0].MyProperty2);
            Assert.Equal("MyProperty103", rows[0].MyProperty3);
            Assert.Equal("MyProperty100", rows[0].MyProperty4);
            Assert.Equal("MyProperty102", rows[0].MyProperty5);
            Assert.Equal("MyProperty6", rows[0].MyProperty6);
        }

        {
            var q =  _importer.QueryExcelAsync<Issue142VO>(path).ToBlockingEnumerable();
            var rows = q.ToList();
            Assert.Equal("CustomColumnName", rows[0].MyProperty1);
            Assert.Null(rows[0].MyProperty7);
            Assert.Equal("MyProperty2", rows[0].MyProperty2);
            Assert.Equal("MyProperty103", rows[0].MyProperty3);
            Assert.Equal("MyProperty100", rows[0].MyProperty4);
            Assert.Equal("MyProperty102", rows[0].MyProperty5);
            Assert.Equal("MyProperty6", rows[0].MyProperty6);
        }
    }

    private class Issue142VO
    {
        [MiniExcelColumnName("CustomColumnName")]
        public string MyProperty1 { get; set; }  //index = 1
        [MiniExcelIgnore]
        public string MyProperty7 { get; set; } //index = null
        public string MyProperty2 { get; set; } //index = 3
        [MiniExcelColumnIndex(6)]
        public string MyProperty3 { get; set; } //index = 6
        [MiniExcelColumnIndex("A")] // equal column index 0
        public string MyProperty4 { get; set; }
        [MiniExcelColumnIndex(2)]
        public string MyProperty5 { get; set; } //index = 2
        public string MyProperty6 { get; set; } //index = 4
    }

    private class Issue142VoOverIndex
    {
        [MiniExcelColumnIndex("Z")]
        public int MyProperty1 { get; set; }
    }

    private class Issue142VoExcelColumnNameNotFound
    {
        [MiniExcelColumnIndex("B")]
        public int MyProperty1 { get; set; }
    }

    /// <summary>
    /// https://github.com/mini-software/MiniExcel/issues/150
    /// </summary>
    [Fact]
    public async Task Issue150()
    {
        var path = PathHelper.GetTempFilePath();
    
        await Assert.ThrowsAnyAsync<NotSupportedException>(async () => await  _exporter.ExportExcelAsync(path, new[] { 1, 2 }));
        File.Delete(path);
  
        await Assert.ThrowsAnyAsync<NotSupportedException>(async () => await  _exporter.ExportExcelAsync(path, new[] { "1", "2" }));
        File.Delete(path);
        
        await Assert.ThrowsAnyAsync<NotSupportedException>(async () => await  _exporter.ExportExcelAsync(path, new[] { '1', '2' }));
        File.Delete(path);
        
        await Assert.ThrowsAnyAsync<NotSupportedException>(async () => await  _exporter.ExportExcelAsync(path, new[] { DateTime.Now }));
        File.Delete(path);
        
        await Assert.ThrowsAnyAsync<NotSupportedException>(async () => await  _exporter.ExportExcelAsync(path, new[] { Guid.NewGuid() }));
        File.Delete(path);
    }

    /// <summary>
    /// https://github.com/mini-software/MiniExcel/issues/157
    /// </summary>
    [Fact]
    public async Task Issue157()
    {
        {
            using var file = AutoDeletingPath.Create();
            var path = file.ToString();
            
            _output.WriteLine("==== SaveAs by strongly type ====");
            var input = JsonConvert.DeserializeObject<IEnumerable<MiniExcelOpenXmlTests.UserAccount>>(
                """
                [
                  {
                    "ID":"78de23d2-dcb6-bd3d-ec67-c112bbc322a2",
                    "Name":"Wade",
                    "BoD":"2020-09-27T00:00:00",
                    "Age":5019,
                    "VIP":false,
                    "Points":5019.12,
                    "IgnoredProperty":null
                  },
                  {
                    "ID":"20d3bfce-27c3-ad3e-4f70-35c81c7e8e45",
                    "Name":"Felix",
                    "BoD":"2020-10-25T00:00:00",
                    "Age":7028,
                    "VIP":true,
                    "Points":7028.46,
                    "IgnoredProperty":null
                  },
                  {
                    "ID":"52013bf0-9aeb-48e6-e5f5-e9500afb034f",
                    "Name":"Phelan",
                    "BoD":"2021-10-04T00:00:00",
                    "Age":3836,
                    "VIP":true,
                    "Points":3835.7,
                    "IgnoredProperty":null
                  },
                    {
                    "ID":"3b97b87c-7afe-664f-1af5-6914d313ae25",
                    "Name":"Samuel",
                    "BoD":"2020-06-21T00:00:00",
                    "Age":9352,
                    "VIP":false,
                    "Points":9351.71,
                    "IgnoredProperty":null
                  },
                  {
                    "ID":"9a989c43-d55f-5306-0d2f-0fbafae135bb",
                    "Name":"Raymond",
                    "BoD":"2021-07-12T00:00:00",
                    "Age":8210,
                    "VIP":true,
                    "Points":8209.76,
                    "IgnoredProperty":null
                  }
                ]
                """);
            var rowsWritten = await  _exporter.ExportExcelAsync(path, input);
            Assert.Single(rowsWritten);
            Assert.Equal(5, rowsWritten[0]);

            var q =  _importer.QueryExcelAsync(path, sheetName: "Sheet1").ToBlockingEnumerable();
            var rows = q.ToList();
            Assert.Equal(6, rows.Count);
            Assert.Equal("Sheet1",  _importer.GetSheetNames(path).First());

            using var p = new ExcelPackage(new FileInfo(path));
            var ws = p.Workbook.Worksheets.First();
            Assert.Equal("Sheet1", ws.Name);
            Assert.Equal("Sheet1", p.Workbook.Worksheets["Sheet1"].Name);
        }
        {
            const string path = "../../../../../samples/xlsx/TestIssue157.xlsx";
            {
                var q =  _importer.QueryExcelAsync(path, sheetName: "Sheet1").ToBlockingEnumerable();
                var rows = q.ToList();
                Assert.Equal(6, rows.Count);
                Assert.Equal("Sheet1",  _importer.GetSheetNames(path).First());
            }
            using (var p = new ExcelPackage(new FileInfo(path)))
            {
                var ws = p.Workbook.Worksheets.First();
                Assert.Equal("Sheet1", ws.Name);
                Assert.Equal("Sheet1", p.Workbook.Worksheets["Sheet1"].Name);
            }

            {
                var q =  _importer.QueryExcelAsync<MiniExcelOpenXmlTests.UserAccount>(path, sheetName: "Sheet1").ToBlockingEnumerable();
                var rows = q.ToList();
                Assert.Equal(5, rows.Count);

                Assert.Equal(Guid.Parse("78DE23D2-DCB6-BD3D-EC67-C112BBC322A2"), rows[0].ID);
                Assert.Equal("Wade", rows[0].Name);
                Assert.Equal(DateTime.ParseExact("27/09/2020", "dd/MM/yyyy", CultureInfo.InvariantCulture), rows[0].BoD);
                Assert.False(rows[0].VIP);
                Assert.Equal(5019m, rows[0].Points);
                Assert.Equal(1, rows[0].IgnoredProperty);
            }
        }
    }

    /// <summary>
    /// https://github.com/mini-software/MiniExcel/issues/149
    /// </summary>
    [Fact]
    public async Task Issue149()
    {
        char[] chars =
        [
            '\u0000', '\u0001', '\u0002', '\u0003', '\u0004', '\u0005', '\u0006', '\u0007', '\u0008',
            '\u0009', //<HT>
            '\u000A', //<LF>
            '\u000B', '\u000C',
            '\u000D', //<CR>
            '\u000E', '\u000F', '\u0010', '\u0011', '\u0012', '\u0013', '\u0014', '\u0015', '\u0016',
            '\u0017', '\u0018', '\u0019', '\u001A', '\u001B', '\u001C', '\u001D', '\u001E', '\u001F', '\u007F'
        ];
        var strings = chars.Select(s => s.ToString()).ToArray();
        
        {
            const string path = "../../../../../samples/xlsx/TestIssue149.xlsx";
            var q =  _importer.QueryExcelAsync(path).ToBlockingEnumerable();
            var rows = q.Select(s => (string)s.A).ToList();
            
            for (int i = 0; i < chars.Length; i++)
            {
                if (i == 13)
                    continue;
                
                Assert.Equal(strings[i], rows[i]);
            }
        }

        {
            using var file = AutoDeletingPath.Create();
            var path = file.ToString();

            var input = chars.Select(s => new { Test = s.ToString() });
            await  _exporter.ExportExcelAsync(path, input);

            var q =  _importer.QueryExcelAsync(path, true).ToBlockingEnumerable();

            var rows = q.Select(s => (string)s.Test).ToList();
            for (int i = 0; i < chars.Length; i++)
            {
                _output.WriteLine($"{i}, {chars[i]}, {rows[i]}");
                if (i is 13 or 9 or 10)
                    continue;
                
                Assert.Equal(strings[i], rows[i]);
            }
        }

        {
            using var file = AutoDeletingPath.Create();
            var path = file.ToString();

            var input = chars.Select(s => new { Test = s.ToString() });
            await  _exporter.ExportExcelAsync(path, input);

            var q =  _importer.QueryExcelAsync<Issue149VO>(path).ToBlockingEnumerable();
            var rows = q.Select(s => s.Test).ToList();
            
            for (int i = 0; i < chars.Length; i++)
            {
                _output.WriteLine($"{i}, {chars[i]}, {rows[i]}");
                if (i is 13 or 9 or 10)
                    continue;
                
                Assert.Equal(strings[i], rows[i]);
            }
        }
    }

    private class Issue149VO
    {
        public string Test { get; set; }
    }

    /// <summary>
    /// https://github.com/mini-software/MiniExcel/issues/153
    /// </summary>
    [Fact]
    public async Task Issue153()
    {
        const string path = "../../../../../samples/xlsx/TestIssue153.xlsx";
        var q =  _importer.QueryExcelAsync(path, true).ToBlockingEnumerable();
        var rows = q.First() as IDictionary<string, object>;

        Assert.Equal(
        [
            "序号", "代号", "新代号", "名称", "XXX", "部门名称", "单位", "ERP工时   (小时)A", "工时(秒) A/3600", "标准人工工时(秒)",
            "生产标准机器工时(秒)", "财务、标准机器工时(秒)", "更新日期", "产品机种", "备注", "最近一次修改前的标准工时(秒)", "最近一次修改前的标准机时(秒)", "备注1"
        ], rows?.Keys);
    }

    /// <summary>
    /// https://github.com/mini-software/MiniExcel/issues/137
    /// </summary>
    [Fact]
    public async Task Issue137()
    {
        var path = "../../../../../samples/xlsx/TestIssue137.xlsx";

        {
            var q =  _importer.QueryExcelAsync(path).ToBlockingEnumerable();
            var rows = q.ToList();
            var first = rows[0] as IDictionary<string, object>; // https://user-images.githubusercontent.com/12729184/113266322-ba06e400-9307-11eb-9521-d36abfda75cc.png
            Assert.Equal(["A", "B", "C", "D", "E", "F", "G", "H"], first?.Keys.ToArray());
            Assert.Equal(11, rows.Count);

            {
                var row = rows[0] as IDictionary<string, object>;
                Assert.Equal("比例", row!["A"]);
                Assert.Equal("商品", row["B"]);
                Assert.Equal("滿倉口數", row["C"]);
                Assert.Equal(" ", row["D"]);
                Assert.Equal(" ", row["E"]);
                Assert.Equal(" ", row["F"]);
                Assert.Equal(0.0, row["G"]);
                Assert.Equal("1為港幣 0為台幣", row["H"]);
            }
            {
                var row = rows[1] as IDictionary<string, object>;
                Assert.Equal(1.0, row!["A"]);
                Assert.Equal("MTX", row["B"]);
                Assert.Equal(10.0, row["C"]);
                Assert.Null(row["D"]);
                Assert.Null(row["E"]);
                Assert.Null(row["F"]);
                Assert.Null(row["G"]);
                Assert.Null(row["H"]);
            }
            {
                var row = rows[2] as IDictionary<string, object>;
                Assert.Equal(0.95, row!["A"]);
            }
        }

        // dynamic query with head
        {
            var q =  _importer.QueryExcelAsync(path, true).ToBlockingEnumerable();
            var rows = q.ToList();
            var first = rows[0] as IDictionary<string, object>; // https://user-images.githubusercontent.com/12729184/113266322-ba06e400-9307-11eb-9521-d36abfda75cc.png
            Assert.Equal(["比例", "商品", "滿倉口數", "0", "1為港幣 0為台幣"], first?.Keys.ToArray());
            Assert.Equal(10, rows.Count);
            
            {
                var row = rows[0] as IDictionary<string, object>;
                Assert.Equal(1.0, row!["比例"]);
                Assert.Equal("MTX", row["商品"]);
                Assert.Equal(10.0, row["滿倉口數"]);
                Assert.Null(row["0"]);
                Assert.Null(row["1為港幣 0為台幣"]);
            }

            {
                var row = rows[1] as IDictionary<string, object>;
                Assert.Equal(0.95, row!["比例"]);
            }
        }

        {
            var q =  _importer.QueryExcelAsync<Issue137ExcelRow>(path).ToBlockingEnumerable();
            var rows = q.ToList();
            Assert.Equal(10, rows.Count);
            
            {
                var row = rows[0];
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

    private class Issue137ExcelRow
    {
        public double? 比例 { get; set; }
        public string 商品 { get; set; }
        public int? 滿倉口數 { get; set; }
    }


    /// <summary>
    /// https://github.com/mini-software/MiniExcel/issues/138
    /// </summary>
    [Fact]
    public async Task Issue138()
    {
        const string path = "../../../../../samples/xlsx/TestIssue138.xlsx";
        {
            var q =  _importer.QueryExcelAsync(path, true).ToBlockingEnumerable();
            var rows = q.ToList();
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

            var q =  _importer.QueryExcelAsync<Issue138ExcelRow>(path).ToBlockingEnumerable();
            var rows = q.ToList();
            Assert.Equal(6, rows.Count);
            Assert.Equal(new DateTime(2021, 3, 1), rows[0].Date);

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
    }

    private class Issue138ExcelRow
    {
        public DateTime? Date { get; set; }
        public int? 實單每日損益 { get; set; }
        public int? 程式每日損益 { get; set; }
        public string 商品 { get; set; }
        public double? 滿倉口數 { get; set; }
        public double? 波段 { get; set; }
        public double? 當沖 { get; set; }
    }
}