using MiniExcelLib.Core.Exceptions;
using MiniExcelLib.OpenXml.Tests.Utils;
using MiniExcelLib.Tests.Common.Utils;

namespace MiniExcelLib.OpenXml.Tests.Issues;

public class MiniExcelGithubIssuesAsyncTests(ITestOutputHelper output)
{
    private readonly ITestOutputHelper _output = output;
    
    private readonly OpenXmlImporter _excelImporter =  MiniExcel.Importers.GetOpenXmlImporter();
    private readonly OpenXmlExporter _excelExporter =  MiniExcel.Exporters.GetOpenXmlExporter();
    private readonly OpenXmlTemplater _excelTemplater =  MiniExcel.Templaters.GetOpenXmlTemplater();

    static MiniExcelGithubIssuesAsyncTests()
    {
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
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
        await using var reader = await connection2.ExecuteReaderAsync("SELECT * FROM test");

        var rowsWritten = await  _excelExporter.ExportAsync(path.ToString(), reader);
        Assert.Single(rowsWritten);
        Assert.Equal(0, rowsWritten[0]);

        var rows = await _excelImporter.QueryAsync(path.ToString(), true).ToListAsync();
        Assert.Empty(rows);
    }

    [Fact]
    public async Task Issue87()
    {
        var templatePath = PathHelper.GetFile("xlsx/TestTemplateCenterEmpty.xlsx");
        using var path = AutoDeletingPath.Create();
        var value = new
        {
            Tests = Enumerable.Range(1, 5).Select((_, i) => new { test1 = i, test2 = i })
        };

        await using var stream = File.OpenRead(templatePath);
        _ = await _excelImporter.QueryAsync(templatePath).ToListAsync();
        await _excelTemplater.FillTemplateAsync(path.ToString(), templatePath, value);
    }

    // QueryAsync Merge cells data
    [Fact]
    public async Task Issue122()
    {
        var config = new OpenXmlConfiguration
        {
            FillMergedCells = true
        };
            
        var path1 = PathHelper.GetFile("xlsx/TestIssue122.xlsx");
        var rows1 = await _excelImporter.QueryAsync(path1, hasHeaderRow: true, configuration: config).ToListAsync();
        
        Assert.Equal("HR", rows1[0].Department);
        Assert.Equal("HR", rows1[1].Department);
        Assert.Equal("HR", rows1[2].Department);
        Assert.Equal("IT", rows1[3].Department);
        Assert.Equal("IT", rows1[4].Department);
        Assert.Equal("IT", rows1[5].Department);

        var path2 = PathHelper.GetFile("xlsx/TestIssue122_2.xlsx");
        var rows2 = await _excelImporter.QueryAsync(path2, hasHeaderRow: true, configuration: config).ToListAsync();

        Assert.Equal("V1", rows2[2].Test1);
        Assert.Equal("V2", rows2[5].Test2);
        Assert.Equal("V3", rows2[1].Test3);
        Assert.Equal("V4", rows2[2].Test4);
        Assert.Equal("V5", rows2[3].Test5);
        Assert.Equal("V6", rows2[5].Test5);
    }

    // SaveAs Default Template
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

            await  _excelExporter.ExportAsync(path.ToString(), value);
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
            var rowsWritten = await  _excelExporter.ExportAsync(path.ToString(), value, configuration: config);
            
            Assert.Single(rowsWritten);
            Assert.Equal(2, rowsWritten[0]);
        }

        {
            using var path = AutoDeletingPath.Create();

            var dt = new DataTable();
            dt.Columns.Add("Name");
            dt.Columns.Add("Age");
            dt.Columns.Add("Date");

            dt.Rows.Add("Jack", 25, new DateTime(2021, 01, 03));
            dt.Rows.Add("Henry", 36, new DateTime(2021, 01, 03));

            var rowsWritten = await  _excelExporter.ExportAsync(path.ToString(), dt);
            
            Assert.Single(rowsWritten);
            Assert.Equal(2, rowsWritten[0]);
        }
    } 

    [Fact]
    public async Task Issue137()
    {
        var path = PathHelper.GetFile("xlsx/TestIssue137.xlsx");
        {
            var rows = await _excelImporter.QueryAsync(path).Cast<IDictionary<string, object>>().ToListAsync();
            Assert.Equal(["A", "B", "C", "D", "E", "F", "G", "H"], rows[0].Keys.ToArray());
            Assert.Equal(11, rows.Count);

            var row1 = rows[0];
            Assert.Equal("比例", row1["A"]);
            Assert.Equal("商品", row1["B"]);
            Assert.Equal("滿倉口數", row1["C"]);
            Assert.Equal(" ", row1["D"]);
            Assert.Equal(" ", row1["E"]);
            Assert.Equal(" ", row1["F"]);
            Assert.Equal(0.0, row1["G"]);
            Assert.Equal("1為港幣 0為台幣", row1["H"]);

            var row2 = rows[1];
            Assert.Equal(1.0, row2["A"]);
            Assert.Equal("MTX", row2["B"]);
            Assert.Equal(10.0, row2["C"]);
            Assert.Null(row2["D"]);
            Assert.Null(row2["E"]);
            Assert.Null(row2["F"]);
            Assert.Null(row2["G"]);
            Assert.Null(row2["H"]);

            var row3 = rows[2];
            Assert.Equal(0.95, row3["A"]);
        }

        // dynamic query with head
        {
            var rows = await _excelImporter.QueryAsync(path, true).Cast<IDictionary<string, object>>().ToListAsync();
            var first = rows[0]; // https://user-images.githubusercontent.com/12729184/113266322-ba06e400-9307-11eb-9521-d36abfda75cc.png
            Assert.Equal(["比例", "商品", "滿倉口數", "0", "1為港幣 0為台幣"], first.Keys.ToArray());
            Assert.Equal(10, rows.Count);
            
            var row1 = rows[0];
            Assert.Equal(1.0, row1["比例"]);
            Assert.Equal("MTX", row1["商品"]);
            Assert.Equal(10.0, row1["滿倉口數"]);
            Assert.Null(row1["0"]);
            Assert.Null(row1["1為港幣 0為台幣"]);

            var row2 = rows[1];
            Assert.Equal(0.95, row2["比例"]);
        }

        {
            var rows = await _excelImporter.QueryAsync<Issue137Dto>(path).ToListAsync();
            Assert.Equal(10, rows.Count);
            
            var row1 = rows[0];
            Assert.Equal(1, row1.比例);
            Assert.Equal("MTX", row1.商品);
            Assert.Equal(10, row1.滿倉口數);

            var row2 = rows[1];
            Assert.Equal(0.95, row2.比例);
        }
    }

    [Fact]
    public async Task Issue138()
    {
        var path = PathHelper.GetFile("xlsx/TestIssue138.xlsx");
        {
            var rows = await _excelImporter.QueryAsync(path, true).ToListAsync();
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

            var rows = await _excelImporter.QueryAsync<Issue138Dto>(path).ToListAsync();
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

    [Fact]
    public async Task Issue142()
    {
        var path = PathHelper.GetFile("xlsx/TestIssue142.xlsx");
        var rows = await _excelImporter.QueryAsync<Issue142DtoVariant2>(path).ToListAsync();

        Assert.Equal(0, rows[0].MyProperty1);
        await Assert.ThrowsAsync<InvalidMappingException>(async () =>
        {
            _ = await _excelImporter.QueryAsync<Issue142DtoVariant1>(path).ToListAsync();
        });

        var rows1 = await _excelImporter.QueryAsync<Issue142Dto>(path).ToListAsync();
        Assert.Equal("CustomColumnName", rows1[0].MyProperty1);
        Assert.Null(rows1[0].MyProperty7);
        Assert.Equal("MyProperty2", rows1[0].MyProperty2);
        Assert.Equal("MyProperty103", rows1[0].MyProperty3);
        Assert.Equal("MyProperty100", rows1[0].MyProperty4);
        Assert.Equal("MyProperty102", rows1[0].MyProperty5);
        Assert.Equal("MyProperty6", rows1[0].MyProperty6);

        var rows2 = await _excelImporter.QueryAsync<Issue142Dto>(path).ToListAsync();
        Assert.Equal("CustomColumnName", rows2[0].MyProperty1);
        Assert.Null(rows2[0].MyProperty7);
        Assert.Equal("MyProperty2", rows2[0].MyProperty2);
        Assert.Equal("MyProperty103", rows2[0].MyProperty3);
        Assert.Equal("MyProperty100", rows2[0].MyProperty4);
        Assert.Equal("MyProperty102", rows2[0].MyProperty5);
        Assert.Equal("MyProperty6", rows2[0].MyProperty6);
    }

    // QueryAsync Support StartCell
    [Fact]
    public async Task Issue147()
    {
        var path1 = PathHelper.GetFile("xlsx/TestIssue147.xlsx");
        var rows1 = await _excelImporter.QueryAsync(path1, hasHeaderRow: false, startCell: "C3", sheetName: "Sheet1").ToListAsync();
            
        Assert.Equal(["C", "D", "E"], (rows1[0] as IDictionary<string, object>)?.Keys);
        Assert.Equal(new[]{ "Column1", "Column2", "Column3" }, new[] { rows1[0].C as string, rows1[0].D as string, rows1[0].E as string });
        Assert.Equal(new[]{ "C4", "D4", "E4" }, new[] { rows1[1].C as string, rows1[1].D as string, rows1[1].E as string });
        Assert.Equal(new[]{ "C9", "D9", "E9" }, new[] { rows1[6].C as string, rows1[6].D as string, rows1[6].E as string });
        Assert.Equal(new[]{ "C12", "D12", "E12" }, new[] { rows1[9].C as string, rows1[9].D as string, rows1[9].E as string });
        Assert.Equal(new[]{ "C13", "D13", "E13" }, new[] { rows1[10].C as string, rows1[10].D as string, rows1[10].E as string });
            
        foreach (var i in new[] { 4, 5, 7, 8 })
        {
            Assert.Equal(new string?[]{null, null, null}, new[] { rows1[i].C as string, rows1[i].D as string, rows1[i].E as string });
        }
        Assert.Equal(11, rows1.Count);

        var columns1 = await _excelImporter.GetColumnNamesAsync(path1, startCell: "C3");
        Assert.Equal(["C", "D", "E"], columns1);

        
        var path2 = PathHelper.GetFile("xlsx/TestIssue147.xlsx");
        var rows2 = await _excelImporter.QueryAsync(path2, hasHeaderRow: true, startCell: "C3", sheetName: "Sheet1").ToListAsync();
            
        Assert.Equal(["Column1", "Column2", "Column3"], (rows2[0] as IDictionary<string, object>)?.Keys);
        Assert.Equal(new[]{"C4", "D4", "E4"}, new[] { rows2[0].Column1 as string, rows2[0].Column2 as string, rows2[0].Column3 as string });
        Assert.Equal(new[]{"C9", "D9", "E9"}, new[] { rows2[5].Column1 as string, rows2[5].Column2 as string, rows2[5].Column3 as string });
        Assert.Equal(new[]{"C12", "D12", "E12"}, new[] { rows2[8].Column1 as string, rows2[8].Column2 as string, rows2[8].Column3 as string });
        Assert.Equal(new[]{"C13", "D13", "E13"}, new[] { rows2[9].Column1 as string, rows2[9].Column2 as string, rows2[9].Column3 as string });
            
        foreach (var i in new[] { 3, 4, 6, 7 })
        {
            Assert.Equal(new string?[]{null, null, null}, new[] { rows2[i].Column1 as string, rows2[i].Column2 as string, rows2[i].Column3 as string });
        }
        Assert.Equal(10, rows2.Count);
                
        var columns2 = await  _excelImporter.GetColumnNamesAsync(path2, hasHeaderRow: true, startCell: "C3");
        Assert.Equal(["Column1", "Column2", "Column3"], columns2);
    }

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
        
        var path1 = PathHelper.GetFile("xlsx/TestIssue149.xlsx");
        var rows1 = await _excelImporter.QueryAsync(path1).Select(s => (string)s.A).ToListAsync();
            
        for (int i = 0; i < chars.Length; i++)
        {
            if (i != 13) 
                Assert.Equal(strings[i], rows1[i]);
        }

        using var file1 = AutoDeletingPath.Create();
        var path2 = file1.ToString();

        var input1 = chars.Select(s => new { Test = s.ToString() });
        await _excelExporter.ExportAsync(path2, input1);

        var rows2 = await _excelImporter.QueryAsync(path2, true).Select(s => (string)s.Test).ToListAsync();
        for (int i = 0; i < chars.Length; i++)
        {
            _output.WriteLine($"{i}, {chars[i]}, {rows2[i]}");
            if (i is not (9 or 10 or 13)) 
                Assert.Equal(strings[i], rows2[i]);
        }

        using var file2 = AutoDeletingPath.Create();
        var path3 = file2.ToString();

        var input2 = chars.Select(s => new { Test = s.ToString() });
        await _excelExporter.ExportAsync(path3, input2);

        var rows = await _excelImporter.QueryAsync<Issue149VO>(path3).Select(s => s.Test).ToListAsync();
        for (int i = 0; i < chars.Length; i++)
        {
            _output.WriteLine($"{i}, {chars[i]}, {rows[i]}");
            if (i is not (13 or 9 or 10))
                Assert.Equal(strings[i], rows[i]);
        }
    }

    [Fact]
    public async Task Issue150()
    {
        using var filePath = AutoDeletingPath.Create();
        var path = filePath.ToString();
    
        await Assert.ThrowsAnyAsync<NotSupportedException>(async () => await  _excelExporter.ExportAsync(path, new[] { 1, 2 }, overwriteFile: true));
        await Assert.ThrowsAnyAsync<NotSupportedException>(async () => await  _excelExporter.ExportAsync(path, new[] { "1", "2" }, overwriteFile: true));
        await Assert.ThrowsAnyAsync<NotSupportedException>(async () => await  _excelExporter.ExportAsync(path, new[] { '1', '2' }, overwriteFile: true));
        await Assert.ThrowsAnyAsync<NotSupportedException>(async () => await  _excelExporter.ExportAsync(path, new[] { DateTime.Now }, overwriteFile: true));
        await Assert.ThrowsAnyAsync<NotSupportedException>(async () => await  _excelExporter.ExportAsync(path, new[] { Guid.NewGuid() }, overwriteFile: true));
    }

    [Fact]
    public async Task Issue153()
    {
        var path = PathHelper.GetFile("xlsx/TestIssue153.xlsx");
        var row = (IDictionary<string, object>)await _excelImporter.QueryAsync(path, true).FirstAsync();

        Assert.Equal(
        [
            "序号", "代号", "新代号", "名称", "XXX", "部门名称", "单位", "ERP工时   (小时)A", "工时(秒) A/3600", "标准人工工时(秒)",
            "生产标准机器工时(秒)", "财务、标准机器工时(秒)", "更新日期", "产品机种", "备注", "最近一次修改前的标准工时(秒)", "最近一次修改前的标准机时(秒)", "备注1"
        ], row.Keys);
    }

    [Fact]
    public async Task Issue157()
    {
        {
            using var file = AutoDeletingPath.Create();
            var path = file.ToString();

            List<UserAccount> data = 
            [
                new()
                {
                    ID = new Guid("78de23d2-dcb6-bd3d-ec67-c112bbc322a2"),
                    Name = "Wade",
                    BoD = new DateTime(2020, 9, 27),
                    Points = 5019.12m
                },
                new()
                {
                    ID = new Guid("20d3bfce-27c3-ad3e-4f70-35c81c7e8e45"),
                    Name = "Felix",
                    BoD = new DateTime(2020, 10, 25),
                    Points = 7028.46m
                },
                new()
                {
                    ID = new Guid("52013bf0-9aeb-48e6-e5f5-e9500afb034f"),
                    Name = "Phelan",
                    BoD = new DateTime(2020, 10, 25),
                    Points = 3835.7m,
                    VIP = true
                },
                new()
                {
                    ID = new Guid("3b97b87c-7afe-664f-1af5-6914d313ae25"),
                    Name = "Samuel",
                    BoD = new DateTime(2020, 6, 21),
                    Points = 9351.71m
                },
                new()
                {
                    ID = new Guid("9a989c43-d55f-5306-0d2f-0fbafae135bb"),
                    Name = "Raymond",
                    BoD = new DateTime(2021, 7, 12),
                    Points = 8209.76m,
                    VIP = true
                }
            ];

            var rowsWritten = await _excelExporter.ExportAsync(path, data);
            Assert.Single(rowsWritten);
            Assert.Equal(5, rowsWritten[0]);

            var q = await _excelImporter.QueryAsync(path, sheetName: "Sheet1").ToListAsync();
            var rows = q.ToList();
            Assert.Equal(6, rows.Count);
            Assert.Equal("Sheet1", (await _excelImporter.GetSheetNamesAsync(path))[0]);

            using var p = new ExcelPackage(new FileInfo(path));
            var ws = p.Workbook.Worksheets[0];
            Assert.Equal("Sheet1", ws.Name);
            Assert.Equal("Sheet1", p.Workbook.Worksheets["Sheet1"].Name);
        }
        {
            var path = PathHelper.GetFile("xlsx/TestIssue157.xlsx");
            {
                var rows = await _excelImporter.QueryAsync(path, sheetName: "Sheet1").ToListAsync();
                Assert.Equal(6, rows.Count);
                Assert.Equal("Sheet1", (await _excelImporter.GetSheetNamesAsync(path))[0]);
            }
            using (var p = new ExcelPackage(new FileInfo(path)))
            {
                var ws = p.Workbook.Worksheets.First();
                Assert.Equal("Sheet1", ws.Name);
                Assert.Equal("Sheet1", p.Workbook.Worksheets["Sheet1"].Name);
            }

            {
                var rows = await _excelImporter.QueryAsync<UserAccount>(path, sheetName: "Sheet1").ToListAsync();
                Assert.Equal(5, rows.Count);

                Assert.Equal(Guid.Parse("78DE23D2-DCB6-BD3D-EC67-C112BBC322A2"), rows[0].ID);
                Assert.Equal("Wade", rows[0].Name);
                Assert.Equal(DateTime.ParseExact("27/09/2020", "dd/MM/yyyy", CultureInfo.InvariantCulture), rows[0].BoD);
                Assert.False(rows[0].VIP);
                Assert.Equal(5019.12m, rows[0].Points);
                Assert.Equal(1, rows[0].IgnoredProperty);
            }
        }
    }

    [Fact]
    public async Task Issue193()
    {
        {
            var templatePath = PathHelper.GetFile("xlsx/TestTemplateComplexWithNamespacePrefix.xlsx");
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
            await  _excelTemplater.FillTemplateAsync(path, templatePath, value);

            foreach (var sheetName in await _excelImporter.GetSheetNamesAsync(path))
            {
                var rows = await _excelImporter.QueryAsync(path, sheetName: sheetName).ToListAsync();
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
            var templatePath = PathHelper.GetFile("xlsx/TestTemplateComplex.xlsx");
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
            await  _excelTemplater.FillTemplateAsync(path, templatePath, value);

            var rows = await _excelImporter.QueryAsync(path).ToListAsync();
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
    public async Task Issue206()
    {
        var templatePath = PathHelper.GetFile("xlsx/TestTemplateBasicIEmumerableFill.xlsx");
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
            await  _excelTemplater.FillTemplateAsync(path.ToString(), templatePath, value);

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
            await  _excelTemplater.FillTemplateAsync(path.ToString(), templatePath, value);

            var dimension = SheetHelper.GetFirstSheetDimensionRefValue(path.ToString());
            Assert.Equal("A1:B2", dimension);
        }
    }

    // Template merge row list rendering has no merge
    [Fact]
    public async Task Issue207()
    {
        {
            var templatePath = PathHelper.GetFile("xlsx/TestIssue207_2.xlsx");
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

            await  _excelTemplater.FillTemplateAsync(path, templatePath, value);
            var rows = await _excelImporter.QueryAsync(path).ToListAsync();

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
            var templatePath = PathHelper.GetFile("xlsx/TestIssue207_Template_Merge_row_list_rendering_without_merge/template.xlsx");
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

            await  _excelTemplater.FillTemplateAsync(path, templatePath, value);

            var rows = await _excelImporter.QueryAsync(path).ToListAsync();
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

    // SaveAs support for IDataReader
    [Fact]
    public async Task Issue211()
    {
        using var path = AutoDeletingPath.Create();
        using var tempSqlitePath = AutoDeletingPath.Create(Path.GetTempPath(), $"{Guid.NewGuid()}.db");
        var connectionString = $"Data Source={tempSqlitePath};Version=3;";

        await using var connection = new SQLiteConnection(connectionString);
        await using var reader = await connection.ExecuteReaderAsync("select 1 Test1,2 Test2 union all select 3 , 4 union all select 5 ,6");

        var rowsWritten = await  _excelExporter.ExportAsync(path.ToString(), reader);
        Assert.Single(rowsWritten);
        Assert.Equal(3, rowsWritten[0]);

        var rows = await _excelImporter.QueryAsync(path.ToString(), true).ToListAsync();
        Assert.Equal(1.0, rows[0].Test1);
        Assert.Equal(2.0, rows[0].Test2);
        Assert.Equal(3.0, rows[1].Test1);
        Assert.Equal(4.0, rows[1].Test2);
    }

    // _exporter.ExportXlsx(path, table,sheetName:“Name”) ，final sheetName is incorrectly Sheet1
    [Fact]
    public async Task Issue212()
    {
        const string sheetName = "Demo";
        
        using var path = AutoDeletingPath.Create();
        await  _excelExporter.ExportAsync(path.ToString(), new[] { new { x = 1, y = 2 } }, sheetName: sheetName);

        var actualSheetName =  (await _excelImporter.GetSheetNamesAsync(path.ToString()))[0];
        Assert.Equal(sheetName, actualSheetName);
    }

    // When reading Excel, can return IDataReader and DataTable to facilitate the import of database. Like ExcelDataReader provide reader.AsDataSet()
    [Fact]
    public async Task Issue216()
    {
        using var path = AutoDeletingPath.Create();

        var value = new[] { new { Test1 = "1", Test2 = 2 }, new { Test1 = "3", Test2 = 4 } };
        var rowsWritten = await  _excelExporter.ExportAsync(path.ToString(), value);
        Assert.Single(rowsWritten);
        Assert.Equal(2, rowsWritten[0]);

        using var table = await _excelImporter.QueryAsDataTableAsync(path.ToString());
        Assert.Equal("Test1", table.Columns[0].ColumnName);
        Assert.Equal("Test2", table.Columns[1].ColumnName);
        Assert.Equal("1", table.Rows[0]["Test1"]);
        Assert.Equal(2.0, table.Rows[0]["Test2"]);
        Assert.Equal("3", table.Rows[1]["Test1"]);
        Assert.Equal(4.0, table.Rows[1]["Test2"]);

        using var dt = await _excelImporter.QueryAsDataTableAsync(path.ToString(), false);
        Assert.Equal("Test1", dt.Rows[0]["A"]);
        Assert.Equal("Test2", dt.Rows[0]["B"]);
        Assert.Equal("1", dt.Rows[1]["A"]);
        Assert.Equal(2.0, dt.Rows[1]["B"]);
        Assert.Equal("3", dt.Rows[2]["A"]);
        Assert.Equal(4.0, dt.Rows[2]["B"]);
    }

    // DataTable recommended to use Caption for column name first, then use columname
    [Fact]
    public async Task Issue217()
    {
        using var table = new DataTable();
        table.Columns.Add("CustomerID");
        table.Columns.Add("CustomerName").Caption = "Name";
        table.Columns.Add("CreditLimit").Caption = "Limit";
        table.Rows.Add(1, "Jonathan", 23.44);
        table.Rows.Add(2, "Bill", 56.87);

        using var path = AutoDeletingPath.Create();
        var rowsWritten = await  _excelExporter.ExportAsync(path.ToString(), table);
        Assert.Single(rowsWritten);
        Assert.Equal(2, rowsWritten[0]);
                
        var rows = await _excelImporter.QueryAsync(path.ToString()).ToListAsync();
        Assert.Equal("Name", rows[0].B);
        Assert.Equal("Limit", rows[0].C);
    }

    // Dynamic QueryAsync can't summary numeric cell value default, need to cast
    [Fact]
    public async Task Issue220()
    {
        var path = PathHelper.GetFile("xlsx/TestIssue220.xlsx");
        var rows = _excelImporter.QueryAsync(path, hasHeaderRow: true);
        var result = await rows
            .GroupBy(s => s.PRT_ID)
            .Select(g => new
            {
                PRT_ID = g.Key,
                Apr = g.Sum(d => (double?)d.Apr),
                May = g.Sum(d => (double?)d.May),
                Jun = g.Sum(d => (double?)d.Jun),
            })
            .ToListAsync();
        
        Assert.Equal(91843.25, result[0].Jun);
        Assert.Equal(50000.99, result[1].Jun);
    }

    // Custom yyyy-MM-dd format is not converted to datetime
    [Fact]
    public async Task Issue222()
    {
        var path = PathHelper.GetFile("xlsx/TestIssue222.xlsx");
        var rows = await _excelImporter.QueryAsync(path).ToListAsync();
        Assert.Equal(typeof(DateTime), rows[1].A.GetType());
        Assert.Equal(new DateTime(2021, 4, 29), rows[1].A);
    }

    // ASP.NET Webform gridview datasource can't use miniexcel queryasdatatable
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
        var rowsWritten = await _excelExporter.ExportAsync(path.ToString(), value);
        Assert.Single(rowsWritten);
        Assert.Equal(3, rowsWritten[0]);


        using var dt = await _excelImporter.QueryAsDataTableAsync(path.ToString());
        var columns = dt.Columns;
        Assert.Equal(typeof(object), columns[0].DataType);
        Assert.Equal(typeof(object), columns[1].DataType);

        Assert.Equal(123.0, dt.Rows[1]["A"]);
        Assert.Equal("HelloWorld", dt.Rows[2]["B"]);
    }

    /// SaveAsByTemplate single column demension index error
    [Fact]
    public async Task Issue226()
    {
        using var path = AutoDeletingPath.Create();
        var templatePath = PathHelper.GetFile("xlsx/TestIssue226.xlsx");
        await _excelTemplater.FillTemplateAsync(path.ToString(), templatePath, new { employees = new[] { new { name = "123" }, new { name = "123" } } });
        Assert.Equal("A1:A3", SheetHelper.GetFirstSheetDimensionRefValue(path.ToString()));
    }

    // Support Xlsm AutoCheck 
    [Fact]
    public async Task Issue227()
    {
        var xlsmPath = AutoDeletingPath.Create(Path.GetTempPath(), $"{Guid.NewGuid()}.xlsm");
        Assert.Throws<NotSupportedException>(() =>  _excelExporter.Export(xlsmPath.FilePath, new[] { new { V = "A1" }, new { V = "A2" } }));

        var path = PathHelper.GetFile("xlsx/TestIssue227.xlsm");
        var rows1 = await _excelImporter.QueryAsync<UserAccount>(path).ToListAsync();
        Assert.Equal(100, rows1.Count);

        Assert.Equal(Guid.Parse("78DE23D2-DCB6-BD3D-EC67-C112BBC322A2"), rows1[0].ID);
        Assert.Equal("Wade", rows1[0].Name);
        Assert.Equal(DateTime.ParseExact("27/09/2020", "dd/MM/yyyy", CultureInfo.InvariantCulture), rows1[0].BoD);
        Assert.Equal(36, rows1[0].Age);
        Assert.False(rows1[0].VIP);
        Assert.Equal(5019.12m, rows1[0].Points);
        Assert.Equal(1, rows1[0].IgnoredProperty);

        
        await using var stream = File.OpenRead(path);
        var rows2 = await _excelImporter.QueryAsync<UserAccount>(stream).ToListAsync();
        Assert.Equal(100, rows2.Count);

        Assert.Equal(Guid.Parse("78DE23D2-DCB6-BD3D-EC67-C112BBC322A2"), rows2[0].ID);
        Assert.Equal("Wade", rows2[0].Name);
        Assert.Equal(DateTime.ParseExact("27/09/2020", "dd/MM/yyyy", CultureInfo.InvariantCulture), rows2[0].BoD);
        Assert.Equal(36, rows2[0].Age);
        Assert.False(rows2[0].VIP);
        Assert.Equal(5019.12m, rows2[0].Points);
        Assert.Equal(1, rows2[0].IgnoredProperty);
    }

    // QueryAsDataTable error "Cannot set Column to be null"
    [Fact]
    public async Task Issue229()
    {
        var path = PathHelper.GetFile("xlsx/TestIssue229.xlsx");
            
        using var dt = await _excelImporter.QueryAsDataTableAsync(path);
            
        foreach (DataColumn column in dt.Columns)
        {
            var v = dt.Rows[3][column];
            Assert.Equal(DBNull.Value, v);
        }
    }

    // SaveAs By data reader error : 'Invalid attempt to call FieldCount when reader is closed'
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

            var rowsWritten = await  _excelExporter.ExportAsync(path.ToString(), reader, printHeader: true);
            Assert.Single(rowsWritten);
            Assert.Equal(2, rowsWritten[0]);
                
            var rows = await _excelImporter.QueryAsync(path.ToString(), true).ToListAsync();
            Assert.Equal(1, rows[0].id);
            Assert.Equal(2, rows[1].id);
        }
    }

    // QueryAsDataTable A2=5.5 , A3=0.55/1.1 will case double type check error
    [Fact]
    public async Task Issue233()
    {
        var path = PathHelper.GetFile("xlsx/TestIssue233.xlsx");

        using var dt = await _excelImporter.QueryAsDataTableAsync(path);
        var rows = dt.Rows;
            
        Assert.Equal(0.55, rows[0]["Size"]);
        Assert.Equal("0.55/1.1", rows[1]["Size"]);
    }

    // SaveAs support multiple sheets
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
        var rowsWritten = await _excelExporter.ExportAsync(path, sheets);
        
        Assert.Equal(2, rowsWritten.Length);
        Assert.Equal(2, rowsWritten[0]);


        var sheetNames = await _excelImporter.GetSheetNamesAsync(path);
        Assert.Equal("users", sheetNames[0]);
        Assert.Equal("department", sheetNames[1]);

        {
            var rows = await _excelImporter.QueryAsync(path, true, sheetName: "users").ToListAsync();
            Assert.Equal("Jack", rows[0].Name);
            Assert.Equal(25, rows[0].Age);
            Assert.Equal("Mike", rows[1].Name);
            Assert.Equal(44, rows[1].Age);
        }
        {
            var rows = await  _excelImporter.QueryAsync(path, true, sheetName: "department").ToListAsync();
            Assert.Equal("01", rows[0].ID);
            Assert.Equal("HR", rows[0].Name);
            Assert.Equal("02", rows[1].ID);
            Assert.Equal("IT", rows[1].Name);
        }
    }

    /// Support SaveAs by DataSet
    [Fact]
    public async Task Issue235()
    {
        using var file = AutoDeletingPath.Create();
        var path = file.ToString();

        var users = new DataTable { TableName = "users" };
        users.Columns.Add("Name", typeof(string));
        users.Columns.Add("Age", typeof(int));
        users.Rows.Add("Jack", 25);
        users.Rows.Add("Mike", 44);

        var departments = new DataTable { TableName = "departments" };
        departments.Columns.Add("ID");
        departments.Columns.Add("Name");
        departments.Rows.Add("01", "HR");
        departments.Rows.Add("02", "IT");

        DataSet sheets = new();
        sheets.Tables.Add(users);
        sheets.Tables.Add(departments);

        var rowsWritten = await  _excelExporter.ExportAsync(path, sheets);
        Assert.Equal(2, rowsWritten.Length);
        Assert.Equal(2, rowsWritten[0]);

        var sheetNames = await  _excelImporter.GetSheetNamesAsync(path);
        Assert.Equal("users", sheetNames[0]);
        Assert.Equal("departments", sheetNames[1]);

        var rows1 = await _excelImporter.QueryAsync(path, true, sheetName: "users").ToListAsync();
        Assert.Equal("Jack", rows1[0].Name);
        Assert.Equal(25, rows1[0].Age);
        Assert.Equal("Mike", rows1[1].Name);
        Assert.Equal(44, rows1[1].Age);

        var rows2 = await _excelImporter.QueryAsync(path, true, sheetName: "departments").ToListAsync();
        Assert.Equal("01", rows2[0].ID);
        Assert.Equal("HR", rows2[0].Name);
        Assert.Equal("02", rows2[1].ID);
        Assert.Equal("IT", rows2[1].Name);
    }

    /// Support Custom Datetime format
    [Fact]
    public async Task Issue241()
    {
        var date1 = new DateTime(2021, 01, 04);
        var date2 = new DateTime(2020, 04, 05);

        Issue241Dto[] value =
        [
            new() { Name = "Jack", InDate = date1 },
            new() { Name = "Henry", InDate = date2 }
        ];

        using var file = AutoDeletingPath.Create();
        var path = file.ToString();
        var rowsWritten = await _excelExporter.ExportAsync(path, value);
            
        Assert.Single(rowsWritten);
        Assert.Equal(2, rowsWritten[0]);

        using var package = new ExcelPackage(path);
        var cells = package.Workbook.Worksheets[0].Cells;

        Assert.Equal(date1, DateTime.FromOADate((double)cells["B2"].Value));
        Assert.Equal("01 04, 2021", cells["B2"].Text);
        Assert.Equal(date2, DateTime.FromOADate((double)cells["B3"].Value));
        Assert.Equal("04 05, 2020", cells["B3"].Text);
    }

    /// No error exception throw when reading xls file
    [Fact]
    public async Task Issue242()
    {
        var path = PathHelper.GetFile("xls/TestIssue242.xls");
        await Assert.ThrowsAsync<InvalidDataException>(async () => _ = await _excelImporter.QueryAsync(path).ToListAsync());

        await using var stream = File.OpenRead(path);
        await Assert.ThrowsAsync<InvalidDataException>(async () => _ = await _excelImporter.QueryAsync(stream).ToListAsync());
    }

    // SaveAsByTemplate support DateTime custom format
    [Fact]
    public async Task Issue255()
    {
        var dt1 = new DateTime(2021, 01, 01);
        var dt2 = new DateTime(2022, 01, 01);
       
        //template
        {
            var templatePath = PathHelper.GetFile("xlsx/TestsIssue255_Template.xlsx");
            await using var ms = new MemoryStream();
            var value = new
            {
                Issue255DTO = new[] { new Issue255DTO { Time = dt1, Time2 = dt2 } }
            };
            
            await _excelTemplater.FillTemplateAsync(ms, templatePath, value);

            ms.Seek(0, SeekOrigin.Begin);
            using var package = new ExcelPackage(ms);
            var cells = package.Workbook.Worksheets[0].Cells;

            Assert.Equal("2021", cells["A2"].Text);
            Assert.Equal("2022", cells["B2"].Text);
        }
        //export
        {
            await using var ms = new MemoryStream();
            Issue255DTO[] value = 
            [
                new() { Time = dt1, Time2 = dt2 }
            ];

            var rowsWritten = await  _excelExporter.ExportAsync(ms, value);
            Assert.Single(rowsWritten);
            Assert.Equal(1, rowsWritten[0]);

            ms.Seek(0, SeekOrigin.Begin);
            using var package = new ExcelPackage(ms);

            var cells = package.Workbook.Worksheets[0].Cells;
            Assert.Equal(dt1, DateTime.FromOADate((double)cells["A2"].Value));
            Assert.Equal("2021", cells["A2"].Text);
            Assert.Equal(dt2, DateTime.FromOADate((double)cells["B2"].Value));
            Assert.Equal("2022", cells["B2"].Text);
        }
    }

    // Dynamic QueryAsync custom format not using mapping format
    [Fact]
    public async Task Issue256()
    {
        var path = PathHelper.GetFile("xlsx/TestIssue256.xlsx");
        var q = await _excelImporter.QueryAsync(path).ToListAsync();
        var rows = q.ToList();
        
        Assert.Equal(new DateTime(2003, 4, 16), rows[1].A);
        Assert.Equal(new DateTime(2004, 4, 16), rows[1].B);
    }

    [Fact]
    public async Task Issue520()
    {
        await using var ms = new MemoryStream();

        Issue520Dto[] data = [new(542, DateTime.Today, 300)];

        await _excelExporter.ExportAsync(ms, data);
        ms.Seek(0, SeekOrigin.Begin);
        
        using var package = new ExcelPackage(ms);
        var cells = package.Workbook.Worksheets.First().Cells;
        
        Assert.Equal(542.0, cells["A2"].Value);
        Assert.Equal(DateTime.Today, DateTime.FromOADate((double)cells["B2"].Value));
        Assert.Equal(300.0, cells["C2"].Value);
    }

    [Fact]
    public async Task TestIssue627()
    {
        var data = new[] { new { LongNumber = "1550432695793487872" } };

        var config = new OpenXmlConfiguration
        {
            DynamicColumns =
            [
                new DynamicExcelColumn("LongNumber") { Format = "@" }
            ]
        };

        await using var ms = new MemoryStream();
        await _excelExporter.ExportAsync(ms, data, configuration: config);
        ms.Seek(0, SeekOrigin.Begin);

        using var package = new ExcelPackage(ms);
        var cell = package.Workbook.Worksheets[0].Cells["A2"];

        Assert.Equal("1550432695793487872", cell.GetValue<string>());
        Assert.Equal("@", cell.Style.Numberformat.Format);
    }

    [Fact]
    public async Task TestIssue658()
    {
        static IEnumerable<Issue658Dto> GetTestData()
        {
            yield return new() { FirstName = "Unit", LastName = "Test" };
            yield return new() { FirstName = "Unit1", LastName = "Test1" };
            yield return new() { FirstName = "Unit2", LastName = "Test2" };
        }

        using var memoryStream = new MemoryStream();
        var testData = GetTestData().ToList();
        await _excelExporter.ExportAsync(memoryStream, testData, configuration: new OpenXmlConfiguration
        {
            FastMode = true,
        });

        memoryStream.Position = 0;
        var queryData = await _excelImporter.QueryAsync<Issue658Dto>(memoryStream).ToListAsync();
        Assert.Equal(testData.Count, queryData.Count);

        var i = 0;
        foreach (var data in testData)
        {
            Assert.Equal(data.FirstName, queryData[i].FirstName);
            Assert.Equal(data.LastName, queryData[i].LastName);
            i++;
        }
    }

    [Fact]
    public async Task TestIssue951()
    {
        var templatePath = PathHelper.GetFile("xlsx/TestTemplateEasyFill.xlsx");
        using var path = AutoDeletingPath.Create();
        
        var value = new Issue951Dto
        {
            Name = "Jack",
            CreateDate = new DateTime(2021, 01, 01),
            VIP = true,
            Points = 123
        };

        // must not throw because of indexer
        await _excelTemplater.FillTemplateAsync(path.ToString(), templatePath, value);
    }
}
