using Dapper;
using MiniExcelLibs;
using MiniExcelLibs.Tests.Utils;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using Xunit;

namespace MiniExcelTests
{
    public class MiniExcelTemplateAsyncTests
    {
        [Fact]
        public async Task DatatableTemptyRowTest()
        {
            {
                var path = Path.Combine(Path.GetTempPath(), $"{Guid.NewGuid().ToString()}.xlsx");
                var templatePath = @"../../../../../samples/xlsx/TestTemplateComplex.xlsx";
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
                var value = new Dictionary<string, object>()
                {
                    ["title"] = "FooCompany",
                    ["managers"] = managers,
                    ["employees"] = employees
                };
                await MiniExcel.SaveAsByTemplateAsync(path, templatePath, value);
                {
                    var rows = MiniExcel.Query(path).ToList();

                    var demension = Helpers.GetFirstSheetDimensionRefValue(path);
                    Assert.Equal("A1:C5", demension);
                }
            }
            {
                var path = Path.Combine(Path.GetTempPath(), $"{Guid.NewGuid().ToString()}.xlsx");
                var templatePath = @"../../../../../samples/xlsx/TestTemplateComplex.xlsx";
                var managers = new DataTable();
                {
                    managers.Columns.Add("name");
                    managers.Columns.Add("department");
                    managers.Rows.Add("Jack", "HR");
                }
                var employees = new DataTable();
                {
                    employees.Columns.Add("name");
                    employees.Columns.Add("department");
                    employees.Rows.Add("Wade", "HR");
                }
                var value = new Dictionary<string, object>()
                {
                    ["title"] = "FooCompany",
                    ["managers"] = managers,
                    ["employees"] = employees
                };
                await MiniExcel.SaveAsByTemplateAsync(path, templatePath, value);
                {
                    var rows = MiniExcel.Query(path).ToList();

                    var demension = Helpers.GetFirstSheetDimensionRefValue(path);
                    Assert.Equal("A1:C5", demension);
                }
            }
        }

        [Fact]
        public async Task DatatableTest()
        {
            var path = Path.Combine(Path.GetTempPath(), $"{Guid.NewGuid().ToString()}.xlsx");
            var templatePath = @"../../../../../samples/xlsx/TestTemplateComplex.xlsx";
            var managers = new DataTable();
            {
                managers.Columns.Add("name");
                managers.Columns.Add("department");
                managers.Rows.Add("Jack", "HR");
                managers.Rows.Add("Loan", "IT");
            }
            var employees = new DataTable();
            {
                employees.Columns.Add("name");
                employees.Columns.Add("department");
                employees.Rows.Add("Wade", "HR");
                employees.Rows.Add("Felix", "HR");
                employees.Rows.Add("Eric", "IT");
                employees.Rows.Add("Keaton", "IT");
            }
            var value = new Dictionary<string, object>()
            {
                ["title"] = "FooCompany",
                ["managers"] = managers,
                ["employees"] = employees
            };
            await MiniExcel.SaveAsByTemplateAsync(path, templatePath, value);
            {
                var rows = MiniExcel.Query(path).ToList();

                var demension = Helpers.GetFirstSheetDimensionRefValue(path);
                Assert.Equal("A1:C9", demension);

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
            }

            {
                var rows = MiniExcel.Query(path, sheetName: "Sheet2").ToList();

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

                var demension = Helpers.GetFirstSheetDimensionRefValue(path);
                Assert.Equal("A1:C9", demension);
            }
        }


        [Fact]
        public async Task DapperTemplateTest()
        {
            var path = Path.Combine(Path.GetTempPath(), $"{Guid.NewGuid().ToString()}.xlsx");
            var templatePath = @"../../../../../samples/xlsx/TestTemplateComplex.xlsx";

            var connection = Db.GetConnection("Data Source=:memory:");
            var value = new Dictionary<string, object>()
            {
                ["title"] = "FooCompany",
                ["managers"] = connection.Query("select 'Jack' name,'HR' department union all select 'Loan','IT'"),
                ["employees"] = connection.Query(@"select 'Wade' name,'HR' department union all select 'Felix','HR' union all select 'Eric','IT' union all select 'Keaton','IT'")
            };
            await MiniExcel.SaveAsByTemplateAsync(path, templatePath, value);

            {
                var rows = MiniExcel.Query(path).ToList();

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

                var demension = Helpers.GetFirstSheetDimensionRefValue(path);
                Assert.Equal("A1:C9", demension);
            }

            {
                var rows = MiniExcel.Query(path, sheetName: "Sheet2").ToList();

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

                var demension = Helpers.GetFirstSheetDimensionRefValue(path);
                Assert.Equal("A1:C9", demension);
            }
        }


        [Fact]
        public async Task DictionaryTemplateTest()
        {
            var path = Path.Combine(Path.GetTempPath(), $"{Guid.NewGuid().ToString()}.xlsx");
            var templatePath = @"../../../../../samples/xlsx/TestTemplateComplex.xlsx";

            var value = new Dictionary<string, object>()
            {
                ["title"] = "FooCompany",
                ["managers"] = new[] {
                    new Dictionary<string, object>{["name"]="Jack",["department"]="HR"},
                    new Dictionary<string, object>{["name"]="Loan",["department"]="IT"}
                },
                ["employees"] = new[] {
                    new Dictionary<string, object>{["name"]="Wade",["department"]="HR"},
                    new Dictionary<string, object>{["name"]="Felix",["department"]="HR"},
                    new Dictionary<string, object>{["name"]="Eric",["department"]="IT"},
                    new Dictionary<string, object>{["name"]="Keaton",["department"]="IT"}
                }
            };
            await MiniExcel.SaveAsByTemplateAsync(path, templatePath, value);

            {
                var rows = MiniExcel.Query(path).ToList();

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

                var demension = Helpers.GetFirstSheetDimensionRefValue(path);
                Assert.Equal("A1:C9", demension);
            }

            {
                var rows = MiniExcel.Query(path, sheetName: "Sheet2").ToList();

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

                var demension = Helpers.GetFirstSheetDimensionRefValue(path);
                Assert.Equal("A1:C9", demension);
            }
        }

        [Fact]
        public async Task TestGithubProject()
        {
            var path = Path.Combine(Path.GetTempPath(), $"{Guid.NewGuid().ToString()}.xlsx");
            var templatePath = @"../../../../../samples/xlsx/TestTemplateGithubProjects.xlsx";
            var projects = new[]
            {
                new {Name = "MiniExcel",Link="https://github.com/shps951023/MiniExcel",Star=146, CreateTime=new DateTime(2021,03,01)},
                new {Name = "HtmlTableHelper",Link="https://github.com/shps951023/HtmlTableHelper",Star=16, CreateTime=new DateTime(2020,02,01)},
                new {Name = "PocoClassGenerator",Link="https://github.com/shps951023/PocoClassGenerator",Star=16, CreateTime=new DateTime(2019,03,17)}
            };
            var value = new
            {
                User = "ITWeiHan",
                Projects = projects,
                TotalStar = projects.Sum(s => s.Star)
            };
            await MiniExcel.SaveAsByTemplateAsync(path, templatePath, value);

            var rows = MiniExcel.Query(path).ToList();
            Assert.Equal("ITWeiHan Github Projects", rows[0].B);
            Assert.Equal("Total Star : 178", rows[8].C);

            var demension = Helpers.GetFirstSheetDimensionRefValue(path);
            Assert.Equal("A1:D9", demension);
        }

        public class TestIEnumerableTypePoco
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
            {
                var path = Path.Combine(Path.GetTempPath(), $"{Guid.NewGuid().ToString()}.xlsx");
                var templatePath = @"../../../../../samples/xlsx/TestIEnumerableType.xlsx";

                //1. By POCO
                var poco = new TestIEnumerableTypePoco { @string = "string", @int = 123, @decimal = decimal.Parse("123.45"), @double = (double)123.33, @datetime = new DateTime(2021, 4, 1), @bool = true, @Guid = Guid.NewGuid() };
                var value = new
                {
                    Ts = new[] {
                        poco,
                        new TestIEnumerableTypePoco{},
                        null,
                        new TestIEnumerableTypePoco{},
                        poco
                    }
                };
                await MiniExcel.SaveAsByTemplateAsync(path, templatePath, value);

                var rows = MiniExcel.Query<TestIEnumerableTypePoco>(path).ToList();
                Assert.Equal(poco.@string, rows[0].@string);
                Assert.Equal(poco.@int, rows[0].@int);
                Assert.Equal(poco.@double, rows[0].@double);
                Assert.Equal(poco.@decimal, rows[0].@decimal);
                Assert.Equal(poco.@bool, rows[0].@bool);
                Assert.Equal(poco.datetime, rows[0].datetime);
                Assert.Equal(poco.Guid, rows[0].Guid);

                Assert.Null(rows[1].@string);
                Assert.Null(rows[1].@int);
                Assert.Null(rows[1].@double);
                Assert.Null(rows[1].@decimal);
                Assert.Null(rows[1].@bool);
                Assert.Null(rows[1].datetime);
                Assert.Null(rows[1].Guid);

                // special input null but query is empty vo
                Assert.Null(rows[2].@string);
                Assert.Null(rows[2].@int);
                Assert.Null(rows[2].@double);
                Assert.Null(rows[2].@decimal);
                Assert.Null(rows[2].@bool);
                Assert.Null(rows[2].datetime);
                Assert.Null(rows[2].Guid);

                Assert.Null(rows[3].@string);
                Assert.Null(rows[3].@int);
                Assert.Null(rows[3].@double);
                Assert.Null(rows[3].@decimal);
                Assert.Null(rows[3].@bool);
                Assert.Null(rows[3].datetime);
                Assert.Null(rows[3].Guid);


                Assert.Equal(poco.@string, rows[4].@string);
                Assert.Equal(poco.@int, rows[4].@int);
                Assert.Equal(poco.@double, rows[4].@double);
                Assert.Equal(poco.@decimal, rows[4].@decimal);
                Assert.Equal(poco.@bool, rows[4].@bool);
                Assert.Equal(poco.datetime, rows[4].datetime);
                Assert.Equal(poco.Guid, rows[4].Guid);

                var demension = Helpers.GetFirstSheetDimensionRefValue(path);
                Assert.Equal("A1:G6", demension);
            }
        }

        [Fact]
        public async Task TestTemplateTypeMapping()
        {
            {
                var path = Path.Combine(Path.GetTempPath(), $"{Guid.NewGuid().ToString()}.xlsx");
                var templatePath = @"../../../../../samples/xlsx/TestITemplateTypeAutoMapping.xlsx";

                //1. By POCO
                var value = new TestIEnumerableTypePoco { @string = "string", @int = 123, @decimal = decimal.Parse("123.45"), @double = (double)123.33, @datetime = new DateTime(2021, 4, 1), @bool = true, @Guid = Guid.NewGuid() };
                await MiniExcel.SaveAsByTemplateAsync(path, templatePath, value);

                var rows = MiniExcel.Query<TestIEnumerableTypePoco>(path).ToList();
                Assert.Equal(value.@string, rows[0].@string);
                Assert.Equal(value.@int, rows[0].@int);
                Assert.Equal(value.@double, rows[0].@double);
                Assert.Equal(value.@decimal, rows[0].@decimal);
                Assert.Equal(value.@bool, rows[0].@bool);
                Assert.Equal(value.datetime, rows[0].datetime);
                Assert.Equal(value.Guid, rows[0].Guid);

                var demension = Helpers.GetFirstSheetDimensionRefValue(path);
                Assert.Equal("A1:G2", demension);
            }
        }

        [Fact]
        public async Task TemplateCenterEmptyTest()
        {
            var path = Path.Combine(Path.GetTempPath(), $"{Guid.NewGuid().ToString()}.xlsx");
            var templatePath = @"../../../../../samples/xlsx/TestTemplateCenterEmpty.xlsx";
            var value = new
            {
                Tests = Enumerable.Range(1, 5).Select((s, i) => new { test1 = i, test2 = i })
            };
            await MiniExcel.SaveAsByTemplateAsync(path, templatePath, value);
        }

        [Fact]
        public async Task TemplateAsyncBasiTest()
        {
            {
                var path = Path.Combine(Path.GetTempPath(), $"{Guid.NewGuid().ToString()}.xlsx");
                var templatePath = @"../../../../../samples/xlsx/TestTemplateEasyFill.xlsx";
                // 1. By POCO
                var value = new
                {
                    Name = "Jack",
                    CreateDate = new DateTime(2021, 01, 01),
                    VIP = true,
                    Points = 123
                };
                await MiniExcel.SaveAsByTemplateAsync(path, templatePath, value);

                var rows = MiniExcel.Query(path).ToList();
                Assert.Equal("Jack", rows[1].A);
                Assert.Equal("2021-01-01 00:00:00", rows[1].B);
                Assert.Equal(true, rows[1].C);
                Assert.Equal(123, rows[1].D);
                Assert.Equal("Jack has 123 points", rows[1].E);

                var demension = Helpers.GetFirstSheetDimensionRefValue(path);
                Assert.Equal("A1:E2", demension);
            }

            {
                var path = Path.Combine(Path.GetTempPath(), $"{Guid.NewGuid().ToString()}.xlsx");
                var templatePath = @"../../../../../samples/xlsx/TestTemplateEasyFill.xlsx";
                var templateBytes = File.ReadAllBytes(templatePath);
                // 1. By POCO
                var value = new
                {
                    Name = "Jack",
                    CreateDate = new DateTime(2021, 01, 01),
                    VIP = true,
                    Points = 123
                };
                await MiniExcel.SaveAsByTemplateAsync(path, templateBytes, value);

                var rows = MiniExcel.Query(path).ToList();
                Assert.Equal("Jack", rows[1].A);
                Assert.Equal("2021-01-01 00:00:00", rows[1].B);
                Assert.Equal(true, rows[1].C);
                Assert.Equal(123, rows[1].D);
                Assert.Equal("Jack has 123 points", rows[1].E);

                var demension = Helpers.GetFirstSheetDimensionRefValue(path);
                Assert.Equal("A1:E2", demension);
            }

            {
                var path = Path.Combine(Path.GetTempPath(), $"{Guid.NewGuid().ToString()}.xlsx");
                var templatePath = @"../../../../../samples/xlsx/TestTemplateEasyFill.xlsx";
                var templateBytes = File.ReadAllBytes(templatePath);
                // 1. By POCO
                var value = new
                {
                    Name = "Jack",
                    CreateDate = new DateTime(2021, 01, 01),
                    VIP = true,
                    Points = 123
                };
                using (var stream = File.Create(path))
                {
                    await stream.SaveAsByTemplateAsync(templateBytes, value);
                }

                var rows = MiniExcel.Query(path).ToList();
                Assert.Equal("Jack", rows[1].A);
                Assert.Equal("2021-01-01 00:00:00", rows[1].B);
                Assert.Equal(true, rows[1].C);
                Assert.Equal(123, rows[1].D);
                Assert.Equal("Jack has 123 points", rows[1].E);

                var demension = Helpers.GetFirstSheetDimensionRefValue(path);
                Assert.Equal("A1:E2", demension);
            }

            {
                var path = Path.Combine(Path.GetTempPath(), $"{Guid.NewGuid().ToString()}.xlsx");
                var templatePath = @"../../../../../samples/xlsx/TestTemplateEasyFill.xlsx";
                // 2. By Dictionary
                var value = new Dictionary<string, object>()
                {
                    ["Name"] = "Jack",
                    ["CreateDate"] = new DateTime(2021, 01, 01),
                    ["VIP"] = true,
                    ["Points"] = 123
                };
                await MiniExcel.SaveAsByTemplateAsync(path, templatePath, value);

                var rows = MiniExcel.Query(path).ToList();
                Assert.Equal("Jack", rows[1].A);
                Assert.Equal("2021-01-01 00:00:00", rows[1].B);
                Assert.Equal(true, rows[1].C);
                Assert.Equal(123, rows[1].D);
                Assert.Equal("Jack has 123 points", rows[1].E);

                var demension = Helpers.GetFirstSheetDimensionRefValue(path);
                Assert.Equal("A1:E2", demension);
            }
        }

        [Fact]
        public async Task TestIEnumerable()
        {
            {
                var path = Path.Combine(Path.GetTempPath(), $"{Guid.NewGuid().ToString()}.xlsx");
                var templatePath = @"../../../../../samples/xlsx/TestTemplateBasicIEmumerableFill.xlsx";

                //1. By POCO
                var value = new
                {
                    employees = new[] {
                        new {name="Jack",department="HR"},
                        new {name="Lisa",department="HR"},
                        new {name="John",department="HR"},
                        new {name="Mike",department="IT"},
                        new {name="Neo",department="IT"},
                        new {name="Loan",department="IT"}
                    }
                };
                await MiniExcel.SaveAsByTemplateAsync(path, templatePath, value);

                var demension = Helpers.GetFirstSheetDimensionRefValue(path);
                Assert.Equal("A1:B7", demension);
            }

            {
                var path = Path.Combine(Path.GetTempPath(), $"{Guid.NewGuid().ToString()}.xlsx");
                var templatePath = @"../../../../../samples/xlsx/TestTemplateBasicIEmumerableFill.xlsx";

                //2. By Dictionary
                var value = new Dictionary<string, object>()
                {
                    ["employees"] = new[] {
                        new {name="Jack",department="HR"},
                        new {name="Lisa",department="HR"},
                        new {name="John",department="HR"},
                        new {name="Mike",department="IT"},
                        new {name="Neo",department="IT"},
                        new {name="Loan",department="IT"}
                    }
                };
                await MiniExcel.SaveAsByTemplateAsync(path, templatePath, value);

                var demension = Helpers.GetFirstSheetDimensionRefValue(path);
                Assert.Equal("A1:B7", demension);
            }

            {
                var path = Path.Combine(Path.GetTempPath(), $"{Guid.NewGuid().ToString()}.xlsx");
                var templatePath = @"../../../../../samples/xlsx/TestTemplateBasicIEmumerableFill.xlsx";

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
                var value = new Dictionary<string, object>()
                {
                    ["employees"] = dt
                };
                await MiniExcel.SaveAsByTemplateAsync(path, templatePath, value);

                var demension = Helpers.GetFirstSheetDimensionRefValue(path);
                Assert.Equal("A1:B7", demension);
            }
        }

        [Fact]
        public async Task TemplateTest()
        {
            {
                var path = Path.Combine(Path.GetTempPath(), $"{Guid.NewGuid().ToString()}.xlsx");
                var templatePath = @"../../../../../samples/xlsx/TestTemplateComplex.xlsx";

                // 1. By Class
                var value = new
                {
                    title = "FooCompany",
                    managers = new[] {
                        new {name="Jack",department="HR"},
                        new {name="Loan",department="IT"}
                    },
                    employees = new[] {
                        new {name="Wade",department="HR"},
                        new {name="Felix",department="HR"},
                        new {name="Eric",department="IT"},
                        new {name="Keaton",department="IT"}
                    }
                };
                await MiniExcel.SaveAsByTemplateAsync(path, templatePath, value);

                {
                    var rows = MiniExcel.Query(path).ToList();

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

                    var demension = Helpers.GetFirstSheetDimensionRefValue(path);
                    Assert.Equal("A1:C9", demension);
                }

                {
                    var rows = MiniExcel.Query(path, sheetName: "Sheet2").ToList();

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

                    var demension = Helpers.GetFirstSheetDimensionRefValue(path);
                    Assert.Equal("A1:C9", demension);
                }
            }


            {
                var path = Path.Combine(Path.GetTempPath(), $"{Guid.NewGuid().ToString()}.xlsx");
                var templatePath = @"../../../../../samples/xlsx/TestTemplateComplex.xlsx";

                // 2. By Dictionary
                var value = new Dictionary<string, object>()
                {
                    ["title"] = "FooCompany",
                    ["managers"] = new[] {
                        new {name="Jack",department="HR"},
                        new {name="Loan",department="IT"}
                    },
                    ["employees"] = new[] {
                        new {name="Wade",department="HR"},
                        new {name="Felix",department="HR"},
                        new {name="Eric",department="IT"},
                        new {name="Keaton",department="IT"}
                    }
                };
                await MiniExcel.SaveAsByTemplateAsync(path, templatePath, value);

                var rows = MiniExcel.Query(path).ToList();
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

                var demension = Helpers.GetFirstSheetDimensionRefValue(path);
                Assert.Equal("A1:C9", demension);
            }

        }
    }
}
