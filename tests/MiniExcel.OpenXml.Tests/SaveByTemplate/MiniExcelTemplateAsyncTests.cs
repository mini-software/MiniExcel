using MiniExcelLib.OpenXml.Tests.Utils;
using MiniExcelLib.Tests.Common.Utils;

namespace MiniExcelLib.OpenXml.Tests.SaveByTemplate;

public class MiniExcelTemplateAsyncTests
{
    private readonly OpenXmlImporter _excelImporter =  MiniExcel.Importers.GetOpenXmlImporter();
    private readonly OpenXmlTemplater _excelTemplater =  MiniExcel.Templaters.GetOpenXmlTemplater();
    
    [Fact]
    public async Task DatatableTemptyRowTest()
    {
        var templatePath = PathHelper.GetFile("xlsx/TestTemplateComplex.xlsx");
        {
            using var path = AutoDeletingPath.Create();
            var managers = new DataTable();
            {
                managers.Columns.Add("name");
                managers.Columns.Add("department");
            }
            var employees = new DataTable();
            {
                employees.Columns.Add("name");
                employees.Columns.Add("department");
            }
            var value = new Dictionary<string, object>
            {
                ["title"] = "FooCompany",
                ["managers"] = managers,
                ["employees"] = employees
            };
            await _excelTemplater.FillTemplateAsync(path.ToString(), templatePath, value);
            var rows = await _excelImporter.QueryAsync(path.ToString()).ToListAsync();

            var dimension = SheetHelper.GetFirstSheetDimensionRefValue(path.ToString());
            Assert.Equal("A1:C5", dimension);
        }
        {
            using var path = AutoDeletingPath.Create();
            var managers = new DataTable();
            managers.Columns.Add("name");
            managers.Columns.Add("department");
            managers.Rows.Add("Jack", "HR");
            
            var employees = new DataTable();
            employees.Columns.Add("name");
            employees.Columns.Add("department");
            employees.Rows.Add("Wade", "HR");
            
            var value = new Dictionary<string, object>()
            {
                ["title"] = "FooCompany",
                ["managers"] = managers,
                ["employees"] = employees
            };
            
            await _excelTemplater.FillTemplateAsync(path.ToString(), templatePath, value);
            var rows = await _excelImporter.QueryAsync(path.ToString()).ToListAsync();

            var dimension = SheetHelper.GetFirstSheetDimensionRefValue(path.ToString());
            Assert.Equal("A1:C5", dimension);
        }
    }

    [Fact]
    public async Task DatatableTest()
    {
        var templatePath = PathHelper.GetFile("xlsx/TestTemplateComplex.xlsx");
        var path = AutoDeletingPath.Create();
        
        var managers = new DataTable();
        managers.Columns.Add("name");
        managers.Columns.Add("department");
        managers.Rows.Add("Jack", "HR");
        managers.Rows.Add("Loan", "IT");
        
        var employees = new DataTable();
        employees.Columns.Add("name");
        employees.Columns.Add("department");
        employees.Rows.Add("Wade", "HR");
        employees.Rows.Add("Felix", "HR");
        employees.Rows.Add("Eric", "IT");
        employees.Rows.Add("Keaton", "IT");
        
        var value = new Dictionary<string, object>
        {
            ["title"] = "FooCompany",
            ["managers"] = managers,
            ["employees"] = employees
        };
        await _excelTemplater.FillTemplateAsync(path.ToString(), templatePath, value);
        var rows = await _excelImporter.QueryAsync(path.ToString()).ToListAsync();

        var dimension = SheetHelper.GetFirstSheetDimensionRefValue(path.ToString());
        Assert.Equal("A1:C9", dimension);
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

        {
            rows = await _excelImporter.QueryAsync(path.ToString(), sheetName: "Sheet2").ToListAsync();
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

            dimension = SheetHelper.GetFirstSheetDimensionRefValue(path.ToString());
            Assert.Equal("A1:C9", dimension);
        }
    }

    [Fact]
    public async Task DapperTemplateTest()
    {
        var templatePath = PathHelper.GetFile("xlsx/TestTemplateComplex.xlsx");
        using var path = AutoDeletingPath.Create();

        var connection = Db.GetConnection("Data Source=:memory:");
        var value = new Dictionary<string, object>
        {
            ["title"] = "FooCompany",
            ["managers"] = connection.Query("select 'Jack' name,'HR' department union all select 'Loan','IT'"),
            ["employees"] = connection.Query(@"select 'Wade' name,'HR' department union all select 'Felix','HR' union all select 'Eric','IT' union all select 'Keaton','IT'")
        };
        await _excelTemplater.FillTemplateAsync(path.ToString(), templatePath, value);

        {
            var rows = await _excelImporter.QueryAsync(path.ToString()).ToListAsync();

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

            var dimension = SheetHelper.GetFirstSheetDimensionRefValue(path.ToString());
            Assert.Equal("A1:C9", dimension);
        }

        {
            var rows = await _excelImporter.QueryAsync(path.ToString(), sheetName: "Sheet2").ToListAsync();
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

            var dimension = SheetHelper.GetFirstSheetDimensionRefValue(path.ToString());
            Assert.Equal("A1:C9", dimension);
        }
    }

    [Fact]
    public async Task DictionaryTemplateTest()
    {
        var templatePath = PathHelper.GetFile("xlsx/TestTemplateComplex.xlsx");
        var path = Path.Combine(Path.GetTempPath(), $"{Guid.NewGuid().ToString()}.xlsx");

        var value = new Dictionary<string, object>
        {
            ["title"] = "FooCompany",
            ["managers"] = new[]
            {
                new Dictionary<string, object>{["name"]="Jack",["department"]="HR"},
                new Dictionary<string, object>{["name"]="Loan",["department"]="IT"}
            },
            ["employees"] = new[] 
            {
                new Dictionary<string, object>{["name"]="Wade",["department"]="HR"},
                new Dictionary<string, object>{["name"]="Felix",["department"]="HR"},
                new Dictionary<string, object>{["name"]="Eric",["department"]="IT"},
                new Dictionary<string, object>{["name"]="Keaton",["department"]="IT"}
            }
        };
        await _excelTemplater.FillTemplateAsync(path, templatePath, value);

        {
            var rows = _excelImporter.Query(path).ToList();

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

            var dimension = SheetHelper.GetFirstSheetDimensionRefValue(path);
            Assert.Equal("A1:C9", dimension);
        }

        {
            var rows = _excelImporter.Query(path, sheetName: "Sheet2").ToList();

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

            var dimension = SheetHelper.GetFirstSheetDimensionRefValue(path);
            Assert.Equal("A1:C9", dimension);
        }
    }

    [Fact]
    public async Task TestGithubProject()
    {
        var templatePath = PathHelper.GetFile("xlsx/TestTemplateGithubProjects.xlsx");
        using var path = AutoDeletingPath.Create();
        
        var projects = new[]
        {
            new {Name = "MiniExcel",Link="https://github.com/mini-software/MiniExcel",Star=146, CreateTime=new DateTime(2021,03,01)},
            new {Name = "HtmlTableHelper",Link="https://github.com/mini-software/HtmlTableHelper",Star=16, CreateTime=new DateTime(2020,02,01)},
            new {Name = "PocoClassGenerator",Link="https://github.com/mini-software/PocoClassGenerator",Star=16, CreateTime=new DateTime(2019,03,17)}
        };
        var value = new
        {
            User = "ITWeiHan",
            Projects = projects,
            TotalStar = projects.Sum(s => s.Star)
        };
        await _excelTemplater.FillTemplateAsync(path.ToString(), templatePath, value);

        var rows = await _excelImporter.QueryAsync(path.ToString()).ToListAsync();
        Assert.Equal("ITWeiHan Github Projects", rows[0].B);
        Assert.Equal("Total Star : 178", rows[8].C);

        var dimension = SheetHelper.GetFirstSheetDimensionRefValue(path.ToString());
        Assert.Equal("A1:D9", dimension);
    }

    private class TestIEnumerableTypePoco
    {
        public string @string { get; set; }
        public int? @int { get; set; }
        public decimal? @decimal { get; set; }
        public double? @double { get; set; }
        public DateTime? datetime { get; set; }
        public bool? @bool { get; set; }
        public Guid? Guid { get; set; }
    }
    [Fact]
    public async Task TestIEnumerableType()
    {
        var templatePath = PathHelper.GetFile("xlsx/TestIEnumerableType.xlsx");
        using var path = AutoDeletingPath.Create();

        var poco = new TestIEnumerableTypePoco
        {
            @string = "string",
            @int = 123, 
            @decimal = 123.45m,
            @double = 123.33,
            datetime = new DateTime(2021, 4, 1), 
            @bool = true,
            Guid = Guid.NewGuid()
        };
            
        var value = new
        {
            Ts = new[] {
                poco,
                new TestIEnumerableTypePoco(),
                null,
                new TestIEnumerableTypePoco(),
                poco
            }
        };
        await _excelTemplater.FillTemplateAsync(path.ToString(), templatePath, value);

        var rows = _excelImporter.Query<TestIEnumerableTypePoco>(path.ToString()).ToList();
        Assert.Equal(poco.@string, rows[0].@string);
        Assert.Equal(poco.@int, rows[0].@int);
        Assert.Equal(poco.@double, rows[0].@double);
        Assert.Equal(poco.@decimal, rows[0].@decimal);
        Assert.Equal(poco.@bool, rows[0].@bool);
        Assert.Equal(poco.datetime, rows[0].datetime);
        Assert.Equal(poco.Guid, rows[0].Guid);

        // special input null but query is empty vo
        for (int i = 1; i <= 3; i++)
        {
            Assert.Null(rows[i].@string);
            Assert.Null(rows[i].@int);
            Assert.Null(rows[i].@double);
            Assert.Null(rows[i].@decimal);
            Assert.Null(rows[i].@bool);
            Assert.Null(rows[i].datetime);
            Assert.Null(rows[i].Guid);
        }

        Assert.Equal(poco.@string, rows[4].@string);
        Assert.Equal(poco.@int, rows[4].@int);
        Assert.Equal(poco.@double, rows[4].@double);
        Assert.Equal(poco.@decimal, rows[4].@decimal);
        Assert.Equal(poco.@bool, rows[4].@bool);
        Assert.Equal(poco.datetime, rows[4].datetime);
        Assert.Equal(poco.Guid, rows[4].Guid);

        var dimension = SheetHelper.GetFirstSheetDimensionRefValue(path.ToString());
        Assert.Equal("A1:G6", dimension);
    }

    [Fact]
    public async Task TestTemplateTypeMapping()
    {
        var templatePath = PathHelper.GetFile("xlsx/TestITemplateTypeAutoMapping.xlsx");
        using var path = AutoDeletingPath.Create();

        var value = new TestIEnumerableTypePoco
        {
            @string = "string",
            @int = 123,
            @decimal = 123.45m, 
            @double = 123.33,
            datetime = new DateTime(2021, 4, 1),
            @bool = true,
            Guid = Guid.NewGuid()
        };
        await _excelTemplater.FillTemplateAsync(path.ToString(), templatePath, value);

        var rows = _excelImporter.Query<TestIEnumerableTypePoco>(path.ToString()).ToList();
        Assert.Equal(value.@string, rows[0].@string);
        Assert.Equal(value.@int, rows[0].@int);
        Assert.Equal(value.@double, rows[0].@double);
        Assert.Equal(value.@decimal, rows[0].@decimal);
        Assert.Equal(value.@bool, rows[0].@bool);
        Assert.Equal(value.datetime, rows[0].datetime);
        Assert.Equal(value.Guid, rows[0].Guid);

        var dimension = SheetHelper.GetFirstSheetDimensionRefValue(path.ToString());
        Assert.Equal("A1:G2", dimension);
    }

    [Fact]
    public async Task TemplateCenterEmptyTest()
    {
        var templatePath = PathHelper.GetFile("xlsx/TestTemplateCenterEmpty.xlsx");
        using var path = AutoDeletingPath.Create();
        var value = new
        {
            Tests = Enumerable.Range(1, 5).Select(i => new { test1 = i, test2 = i })
        };
        await _excelTemplater.FillTemplateAsync(path.ToString(), templatePath, value);
    }

    [Fact]
    public async Task TemplateAsyncBasicTest()
    {
        var templatePath = PathHelper.GetFile("xlsx/TestTemplateEasyFill.xlsx");
        {
            using var path = AutoDeletingPath.Create();
            
            // 1. By POCO
            var value = new
            {
                Name = "Jack",
                CreateDate = new DateTime(2021, 01, 01),
                VIP = true,
                Points = 123
            };
            await _excelTemplater.FillTemplateAsync(path.ToString(), templatePath, value);

            var rows = await _excelImporter.QueryAsync(path.ToString()).ToListAsync();
            Assert.Equal("Jack", rows[1].A);
            Assert.Equal("2021-01-01 00:00:00", rows[1].B);
            Assert.Equal(true, rows[1].C);
            Assert.Equal(123, rows[1].D);
            Assert.Equal("Jack has 123 points", rows[1].E);

            var dimension = SheetHelper.GetFirstSheetDimensionRefValue(path.ToString());
            Assert.Equal("A1:E2", dimension);
        }

        {
            var path = Path.Combine(Path.GetTempPath(), $"{Guid.NewGuid().ToString()}.xlsx");
            var templateBytes = File.ReadAllBytes(templatePath);
            // 1. By POCO
            var value = new
            {
                Name = "Jack",
                CreateDate = new DateTime(2021, 01, 01),
                VIP = true,
                Points = 123
            };
            await _excelTemplater.FillTemplateAsync(path, templateBytes, value);

            var rows = await _excelImporter.QueryAsync(path).ToListAsync();
            Assert.Equal("Jack", rows[1].A);
            Assert.Equal("2021-01-01 00:00:00", rows[1].B);
            Assert.Equal(true, rows[1].C);
            Assert.Equal(123, rows[1].D);
            Assert.Equal("Jack has 123 points", rows[1].E);

            var dimension = SheetHelper.GetFirstSheetDimensionRefValue(path);
            Assert.Equal("A1:E2", dimension);
        }

        {
            using var path = AutoDeletingPath.Create();
            var templateBytes = await File.ReadAllBytesAsync(templatePath);
            
            // 1. By POCO
            var value = new
            {
                Name = "Jack",
                CreateDate = new DateTime(2021, 01, 01),
                VIP = true,
                Points = 123
            };
            await using (var stream = File.Create(path.ToString()))
            {
                await _excelTemplater.FillTemplateAsync(stream, templateBytes, value);
            }

            var rows = await _excelImporter.QueryAsync(path.ToString()).ToListAsync();
            Assert.Equal("Jack", rows[1].A);
            Assert.Equal("2021-01-01 00:00:00", rows[1].B);
            Assert.Equal(true, rows[1].C);
            Assert.Equal(123, rows[1].D);
            Assert.Equal("Jack has 123 points", rows[1].E);

            var dimension = SheetHelper.GetFirstSheetDimensionRefValue(path.ToString());
            Assert.Equal("A1:E2", dimension);
        }

        {
            using var path = AutoDeletingPath.Create();
            
            // 2. By Dictionary
            var value = new Dictionary<string, object>()
            {
                ["Name"] = "Jack",
                ["CreateDate"] = new DateTime(2021, 01, 01),
                ["VIP"] = true,
                ["Points"] = 123
            };
            await _excelTemplater.FillTemplateAsync(path.ToString(), templatePath, value);

            var rows = await _excelImporter.QueryAsync(path.ToString()).ToListAsync();
            Assert.Equal("Jack", rows[1].A);
            Assert.Equal("2021-01-01 00:00:00", rows[1].B);
            Assert.Equal(true, rows[1].C);
            Assert.Equal(123, rows[1].D);
            Assert.Equal("Jack has 123 points", rows[1].E);

            var dimension = SheetHelper.GetFirstSheetDimensionRefValue(path.ToString());
            Assert.Equal("A1:E2", dimension);
        }
    }

    [Fact]
    public async Task TestIEnumerable()
    {
        var templatePath = PathHelper.GetFile("xlsx/TestTemplateBasicIEmumerableFill.xlsx");
        {
            using var path = AutoDeletingPath.Create();

            //1. By POCO
            var value = new
            {
                employees = new[] 
                {
                    new { name = "Jack", department = "HR" },
                    new { name = "Lisa", department = "HR" },
                    new { name = "John", department = "HR" },
                    new { name = "Mike", department = "IT" },
                    new { name = "Neo", department = "IT" },
                    new { name = "Loan", department = "IT "}
                }
            };
            await _excelTemplater.FillTemplateAsync(path.ToString(), templatePath, value);

            var dimension = SheetHelper.GetFirstSheetDimensionRefValue(path.ToString());
            Assert.Equal("A1:B7", dimension);
        }

        {
            using var path = AutoDeletingPath.Create();

            //2. By Dictionary
            var value = new Dictionary<string, object>
            {
                ["employees"] = new[] 
                {
                    new { name = "Jack", department = "HR" },
                    new { name = "Lisa", department = "HR" },
                    new { name = "John", department = "HR" },
                    new { name = "Mike", department = "IT" },
                    new { name = "Neo", department = "IT" },
                    new { name = "Loan", department = "IT "}
                }
            };
            await _excelTemplater.FillTemplateAsync(path.ToString(), templatePath, value);

            var dimension = SheetHelper.GetFirstSheetDimensionRefValue(path.ToString());
            Assert.Equal("A1:B7", dimension);
        }

        {
            using var path = AutoDeletingPath.Create();

            //3. By DataTable
            var dt = new DataTable();
            {
                dt.Columns.Add("name");
                dt.Columns.Add("department");
                dt.Rows.Add("Jack", "HR");
                dt.Rows.Add("Lisa", "HR");
                dt.Rows.Add("John", "HR");
                dt.Rows.Add("Mike", "IT");
                dt.Rows.Add("Neo", "IT");
                dt.Rows.Add("Loan", "IT");
            }
            var value = new Dictionary<string, object>
            {
                ["employees"] = dt
            };
            await _excelTemplater.FillTemplateAsync(path.ToString(), templatePath, value);

            var dimension = SheetHelper.GetFirstSheetDimensionRefValue(path.ToString());
            Assert.Equal("A1:B7", dimension);
        }
    }

    [Fact]
    public async Task TestIEnumerableGrouped()
    {
        var templatePath = PathHelper.GetFile("xlsx/TestTemplateBasicIEmumerableFillGroup.xlsx");
        {
            using var path = AutoDeletingPath.Create();

            //1. By POCO
            var value = new
            {
                employees = new[] 
                {
                    new { name = "Jack", department = "HR" },
                    new { name = "Lisa", department = "HR" },
                    new { name = "John", department = "HR" },
                    new { name = "Mike", department = "IT" },
                    new { name = "Neo", department = "IT" },
                    new { name = "Loan", department = "IT" }
                }
            };
            await _excelTemplater.FillTemplateAsync(path.ToString(), templatePath, value);

            var dimension = SheetHelper.GetFirstSheetDimensionRefValue(path.ToString());
            Assert.Equal("A1:B18", dimension);
        }

        {
            using var path = AutoDeletingPath.Create();

            //2. By Dictionary
            var value = new Dictionary<string, object>()
            {
                ["employees"] = new[] 
                {
                    new { name = "Jack", department = "HR" },
                    new { name = "Jack", department = "HR" },
                    new { name = "John", department = "HR" },
                    new { name = "John", department = "IT" },
                    new { name = "Neo", department = "IT" },
                    new { name = "Loan", department = "IT "}
                }
            };
            await _excelTemplater.FillTemplateAsync(path.ToString(), templatePath, value);

            var dimension = SheetHelper.GetFirstSheetDimensionRefValue(path.ToString());
            Assert.Equal("A1:B18", dimension);
        }

        {
            using var path = AutoDeletingPath.Create();

            //3. By DataTable
            var dt = new DataTable();
            {
                dt.Columns.Add("name");
                dt.Columns.Add("department");
                dt.Rows.Add("Jack", "HR");
                dt.Rows.Add("Lisa", "HR");
                dt.Rows.Add("John", "HR");
                dt.Rows.Add("Mike", "IT");
                dt.Rows.Add("Neo", "IT");
                dt.Rows.Add("Loan", "IT");
            }
            var value = new Dictionary<string, object>
            {
                ["employees"] = dt
            };
            await _excelTemplater.FillTemplateAsync(path.ToString(), templatePath, value);

            var dimension = SheetHelper.GetFirstSheetDimensionRefValue(path.ToString());
            Assert.Equal("A1:B18", dimension);
        }
    }

    [Fact]
    public async Task TestIEnumerableConditional()
    {
        var templatePath = PathHelper.GetFile("xlsx/TestTemplateBasicIEmumerableFillConditional.xlsx");
        {
            using var path = AutoDeletingPath.Create();

            //1. By POCO
            var value = new
            {
                employees = new[] 
                {
                    new { name = "Jack", department = "HR" },
                    new { name = "Lisa", department = "HR" },
                    new { name = "John", department = "HR" },
                    new { name = "Mike", department = "IT" },
                    new { name = "Neo", department = "IT" },
                    new { name = "Loan", department = "IT "}
                }
            };
            await _excelTemplater.FillTemplateAsync(path.ToString(), templatePath, value);

            var dimension = SheetHelper.GetFirstSheetDimensionRefValue(path.ToString());
            Assert.Equal("A1:B18", dimension);
        }

        {
            using var path = AutoDeletingPath.Create();

            //2. By Dictionary
            var value = new Dictionary<string, object>
            {
                ["employees"] = new[] 
                {
                    new { name = "Jack", department = "HR" },
                    new { name = "Jack", department = "HR" },
                    new { name = "John", department = "HR" },
                    new { name = "John", department = "IT" },
                    new { name = "Neo", department = "IT" },
                    new { name = "Loan", department = "IT" }
                }
            };
            await _excelTemplater.FillTemplateAsync(path.ToString(), templatePath, value);

            var dimension = SheetHelper.GetFirstSheetDimensionRefValue(path.ToString());
            Assert.Equal("A1:B18", dimension);
        }

        {
            using var path = AutoDeletingPath.Create();

            //3. By DataTable
            var dt = new DataTable();
            {
                dt.Columns.Add("name");
                dt.Columns.Add("department");
                dt.Rows.Add("Jack", "HR");
                dt.Rows.Add("Lisa", "HR");
                dt.Rows.Add("John", "HR");
                dt.Rows.Add("Mike", "IT");
                dt.Rows.Add("Neo", "IT");
                dt.Rows.Add("Loan", "IT");
            }
            var value = new Dictionary<string, object>
            {
                ["employees"] = dt
            };
            await _excelTemplater.FillTemplateAsync(path.ToString(), templatePath, value);

            var dimension = SheetHelper.GetFirstSheetDimensionRefValue(path.ToString());
            Assert.Equal("A1:B18", dimension);
        }
    }

    [Fact]
    public async Task TemplateTest()
    {
        var templatePath = PathHelper.GetFile("xlsx/TestTemplateComplex.xlsx");
        {
            var path = AutoDeletingPath.Create();

            // 1. By Class
            var value = new
            {
                title = "FooCompany",
                managers = new[] 
                {
                    new { name = "Jack", department = "HR" },
                    new { name = "Loan", department = "IT" }
                },
                employees = new[] 
                {
                    new { name = "Wade", department = "HR" },
                    new { name = "Felix", department = "HR" },
                    new { name = "Eric", department = "IT" },
                    new { name = "Keaton", department = "IT" }
                }
            };
            await _excelTemplater.FillTemplateAsync(path.ToString(), templatePath, value);

            {
                var rows = await _excelImporter.QueryAsync(path.ToString()).ToListAsync();
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

                var dimension = SheetHelper.GetFirstSheetDimensionRefValue(path.ToString());
                Assert.Equal("A1:C9", dimension);
            }

            {
                var rows = await _excelImporter.QueryAsync(path.ToString(), sheetName: "Sheet2").ToListAsync();
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

                var dimension = SheetHelper.GetFirstSheetDimensionRefValue(path.ToString());
                Assert.Equal("A1:C9", dimension);
            }
        }

        {
            using var path = AutoDeletingPath.Create();

            // 2. By Dictionary
            var value = new Dictionary<string, object>
            {
                ["title"] = "FooCompany",
                ["managers"] = new[] 
                {
                    new { name = "Jack", department = "HR" },
                    new { name = "Loan", department = "IT" }
                },
                ["employees"] = new[] 
                {
                    new { name = "Wade", department = "HR" },
                    new { name = "Felix", department = "HR" },
                    new { name = "Eric", department = "IT" },
                    new { name = "Keaton", department = "IT" }
                }
            };
            await _excelTemplater.FillTemplateAsync(path.ToString(), templatePath, value);

            var rows = await _excelImporter.QueryAsync(path.ToString()).ToListAsync();
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

            var dimension = SheetHelper.GetFirstSheetDimensionRefValue(path.ToString());
            Assert.Equal("A1:C9", dimension);
        }
    }

    [Fact]
    public async Task SaveAsByTemplateAsync_TakeCancel_Throws_TaskCanceledException()
    {
        var templatePath = PathHelper.GetFile("xlsx/TestTemplateEasyFill.xlsx");
        await Assert.ThrowsAsync<OperationCanceledException>(async () =>
        {
            using var cts = new CancellationTokenSource();
            using var path = AutoDeletingPath.Create();

            var value = new Dictionary<string, object>
            {
                ["Name"] = "Jack",
                ["CreateDate"] = new DateTime(2021, 01, 01),
                ["VIP"] = true,
                ["Points"] = 123
            };

            await cts.CancelAsync();
            await _excelTemplater.FillTemplateAsync(path.ToString(), templatePath, value, cancellationToken: cts.Token);
        });
    }
    
    
    [Fact]
    public async Task TestMergeSameCellsWithTagAsync()
    {
        var path = PathHelper.GetFile("xlsx/TestMergeWithTag.xlsx");
        using var mergedFilePath = AutoDeletingPath.Create();

        await _excelTemplater.MergeSameCellsAsync(mergedFilePath.ToString(), path);
        var mergedCells = SheetHelper.GetFirstSheetMergedCells(mergedFilePath.ToString());

        Assert.Equal("A2:A4", mergedCells[0]);
        Assert.Equal("C3:C4", mergedCells[1]);
        Assert.Equal("A7:A8", mergedCells[2]);
    }

    [Fact]
    public async Task TestMergeSameCellsWithLimitTagAsync()
    {
        var path = PathHelper.GetFile("xlsx/TestMergeWithLimitTag.xlsx");
        using var mergedFilePath = AutoDeletingPath.Create();

        await _excelTemplater.MergeSameCellsAsync(mergedFilePath.ToString(), path);
        var mergedCells = SheetHelper.GetFirstSheetMergedCells(mergedFilePath.ToString());

        Assert.Equal("A3:A4", mergedCells[0]);
        Assert.Equal("C3:C6", mergedCells[1]);
        Assert.Equal("A5:A6", mergedCells[2]);
    }
    
    
    [Fact]
    public async Task TestIEnumerableWithFormulas()
    {
        var templatePath = PathHelper.GetFile("xlsx/TestTemplateBasicIEnumerableFillWithFormulas.xlsx");
        using var path = AutoDeletingPath.Create();

        var value = new
        {
            employees = new[]
            {
                new { name = "Jack", department = "HR", salary = 90000 },
                new { name = "Lisa", department = "HR", salary = 150000 },
                new { name = "John", department = "HR", salary = 64000 },
                new { name = "Mike", department = "IT", salary = 87000 },
                new { name = "Neo", department = "IT", salary = 98000 },
                new { name = "Joan", department = "IT", salary = 120000 }
            }
        };
        await _excelTemplater.FillTemplateAsync(path.FilePath, templatePath, value, true);

        var dimension = SheetHelper.GetFirstSheetDimensionRefValue(path.FilePath);
        Assert.Equal("A1:C13", dimension);
    }
}
