using System.Data;
using Dapper;
using MiniExcelLib.Core.Enums;
using MiniExcelLib.Core.OpenXml.Picture;
using MiniExcelLib.Tests.Utils;
using Importer = MiniExcelLib.MiniExcel.Importer;
using Templater = MiniExcelLib.MiniExcel.Templater;
using Xunit;

namespace MiniExcelLib.Tests.SaveByTemplate;

public class MiniExcelTemplateTests
{
    [Fact]
    public void TestImageType()
    {
        const string templatePath = "../../../../../samples/xlsx/TestImageType.xlsx";
        {
            string absolutePath = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, templatePath));

            using var path = AutoDeletingPath.Create();
            File.Copy(absolutePath, path.FilePath, overwrite: true); // Copy the template file

            var img1Bytes = File.ReadAllBytes("../../../../../samples/images/TestIssue327.png");  // Use your local image
            var img2Bytes = File.ReadAllBytes("../../../../../samples/images/TestIssue327.png");  // Use your local image
            var img3Bytes = File.ReadAllBytes("../../../../../samples/images/TestIssue327.png");  // Use your local image

            var pictures = new[]
            {
                new MiniExcelPicture
                {
                    CellAddress = "B2",
                    ImageBytes = img1Bytes,
                    PictureType = "png",
                    ImgType = XlsxImgType.AbsoluteAnchor,
                    Location = new System.Drawing.Point(255,255),
                    WidthPx = 1920,
                    HeightPx = 1032
                },
                new MiniExcelPicture
                {
                    CellAddress = "D4",
                    ImageBytes = img2Bytes,
                    PictureType = "png",
                    ImgType = XlsxImgType.TwoCellAnchor,
                    WidthPx = 1920,
                    HeightPx = 1032
                },
                new MiniExcelPicture
                {
                    CellAddress = "F6",
                    ImageBytes = img3Bytes,
                    PictureType = "png",
                    ImgType = XlsxImgType.OneCellAnchor,
                    WidthPx = 1920,
                    HeightPx = 1032
                }
            };

            // Act
            MiniExcel.AddPicture(path.ToString(), pictures);

            // Assert
            using var zip = ZipFile.OpenRead(path.FilePath);
            var mediaEntries = zip.Entries.Where(x => x.FullName.StartsWith("xl/media/")).ToList();
            Assert.Equal(pictures.Length, mediaEntries.Count);

            // Assert (use EPPlus to verify that images are inserted correctly)
            using (var package = new ExcelPackage(new FileInfo(path.FilePath)))
            {
                var sheet = package.Workbook.Worksheets[0];
                var picB2 = sheet.Drawings.OfType<ExcelPicture>()
                    .FirstOrDefault(p => p.EditAs == eEditAs.Absolute);

                Assert.NotNull(picB2);
                Assert.Equal(1920 * 9525, picB2.Size.Width);
                Assert.Equal(1032 * 9525, picB2.Size.Height);
                //Console.WriteLine("✅ AbsoluteAnchor image exists and the size is as expected (1920x1032)");

                //Console.WriteLine("✅ Image inserted successfully (B2 - AbsoluteAnchor)");

                // Validate image at D4 (ImgType.TwoCellAnchor)
                var picD4 = sheet.Drawings.OfType<ExcelPicture>()
                    .FirstOrDefault(p => p.EditAs == eEditAs.TwoCell && p.From != null && p.From.Column == 3 && p.From.Row == 3);
                Assert.NotNull(picD4);
                //Console.WriteLine("✅ Image inserted successfully (D4 - TwoCellAnchor)");

                // Validate image at F6 (ImgType.OneCellAnchor)
                var picF6 = sheet.Drawings.OfType<ExcelPicture>()
                    .FirstOrDefault(p => p.EditAs == eEditAs.OneCell && p.From != null && p.From.Column == 5 && p.From.Row == 5);
                Assert.NotNull(picF6);
                //Console.WriteLine("✅ Image inserted successfully (F6 - OneCellAnchor)");
            }
        }
    }
    
    [Fact]
    public void DatatableTemptyRowTest()
    {
        const string templatePath = "../../../../../samples/xlsx/TestTemplateComplex.xlsx";
        {
            using var path = AutoDeletingPath.Create();

            var managers = new DataTable();
            managers.Columns.Add("name");
            managers.Columns.Add("department");
            
            var employees = new DataTable();
            employees.Columns.Add("name");
            employees.Columns.Add("department");
            
            var value = new Dictionary<string, object>()
            {
                ["title"] = "FooCompany",
                ["managers"] = managers,
                ["employees"] = employees
            };
            Templater.ApplyXlsxTemplate(path.ToString(), templatePath, value);
            
            var rows = Importer.QueryXlsx(path.ToString()).ToList();
            var dimension = Helpers.GetFirstSheetDimensionRefValue(path.ToString());
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
            Templater.ApplyXlsxTemplate(path.ToString(), templatePath, value);
            
            var rows = Importer.QueryXlsx(path.ToString()).ToList();
            var dimension = Helpers.GetFirstSheetDimensionRefValue(path.ToString());
            Assert.Equal("A1:C5", dimension);
        }
    }

    [Fact]
    public void DatatableTest()
    {
        const string templatePath = "../../../../../samples/xlsx/TestTemplateComplex.xlsx";
        using var file = AutoDeletingPath.Create();
        var path = file.ToString();

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
        
        var value = new Dictionary<string, object>()
        {
            ["title"] = "FooCompany",
            ["managers"] = managers,
            ["employees"] = employees
        };
        Templater.ApplyXlsxTemplate(path, templatePath, value);

        {
            var rows = Importer.QueryXlsx(path).ToList();

            var dimension = Helpers.GetFirstSheetDimensionRefValue(path);
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
        }

        {
            var rows = Importer.QueryXlsx(path, sheetName: "Sheet2").ToList();
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

            var dimension = Helpers.GetFirstSheetDimensionRefValue(path);
            Assert.Equal("A1:C9", dimension);
        }
    }


    [Fact]
    public void DapperTemplateTest()
    {
        const string templatePath = "../../../../../samples/xlsx/TestTemplateComplex.xlsx";
        using var file = AutoDeletingPath.Create();
        var path = file.ToString();

        var connection = Db.GetConnection("Data Source=:memory:");
        var value = new Dictionary<string, object>
        {
            ["title"] = "FooCompany",
            ["managers"] = connection.Query("select 'Jack' name,'HR' department union all select 'Loan','IT'"),
            ["employees"] = connection.Query("select 'Wade' name,'HR' department union all select 'Felix','HR' union all select 'Eric','IT' union all select 'Keaton','IT'")
        };
        Templater.ApplyXlsxTemplate(path, templatePath, value);

        {
            var rows = Importer.QueryXlsx(path).ToList();
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

            var dimension = Helpers.GetFirstSheetDimensionRefValue(path);
            Assert.Equal("A1:C9", dimension);
        }

        {
            var rows = Importer.QueryXlsx(path, sheetName: "Sheet2").ToList();
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

            var dimension = Helpers.GetFirstSheetDimensionRefValue(path);
            Assert.Equal("A1:C9", dimension);
        }
    }


    [Fact]
    public void DictionaryTemplateTest()
    {
        const string templatePath = "../../../../../samples/xlsx/TestTemplateComplex.xlsx";
        using var file = AutoDeletingPath.Create();
        var path = file.ToString();

        var value = new Dictionary<string, object>
        {
            ["title"] = "FooCompany",
            ["managers"] = new[]
            {
                new Dictionary<string, object> { ["name"] = "Jack", ["department"] = "HR" },
                new Dictionary<string, object> { ["name"] = "Loan", ["department"] = "IT" }
            },
            ["employees"] = new[]
            {
                new Dictionary<string, object> { ["name"] = "Wade", ["department"] = "HR" },
                new Dictionary<string, object> { ["name"] = "Felix", ["department"] = "HR" },
                new Dictionary<string, object> { ["name"] = "Eric", ["department"] = "IT" },
                new Dictionary<string, object> { ["name"] = "Keaton", ["department"] = "IT" }
            }
        };
        Templater.ApplyXlsxTemplate(path, templatePath, value);

        {
            var rows = Importer.QueryXlsx(path).ToList();
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

            var dimension = Helpers.GetFirstSheetDimensionRefValue(path);
            Assert.Equal("A1:C9", dimension);
        }

        {
            var rows = Importer.QueryXlsx(path, sheetName: "Sheet2").ToList();
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

            var dimension = Helpers.GetFirstSheetDimensionRefValue(path);
            Assert.Equal("A1:C9", dimension);
        }
    }

    private class Employee
    {
        public string name { get; set; }
        public string department { get; set; }
    }

    [Fact]
    public void GroupTemplateTest()
    {
        const string templatePath = "../../../../../samples/xlsx/TestTemplateBasicIEmumerableFillGroup.xlsx";
        var file = AutoDeletingPath.Create();
        var path = file.ToString();

        var value = new
        {
            employees = new List<Employee>
            {
                new() { name = "Jack", department = "HR" },
                new() { name = "Jack", department = "IT" },
                new() { name = "Loan", department = "IT" },
                new() { name = "Eric", department = "IT" },
                new() { name = "Eric", department = "HR" },
                new() { name = "Keaton", department = "IT" },
                new() { name = "Felix", department = "HR" }
            }
        };
        Templater.ApplyXlsxTemplate(path, templatePath, value);

        var rows = Importer.QueryXlsx(path).ToList();
        Assert.Equal(16, rows.Count);

        Assert.Equal("Jack", rows[1].A);
        Assert.Equal("Jack", rows[2].A);
        Assert.Equal("HR", rows[2].B);
        Assert.Equal("Jack", rows[3].A);
        Assert.Equal("IT", rows[3].B);

        Assert.Equal("Loan", rows[4].A);
        Assert.Equal("Loan", rows[5].A);
        Assert.Equal("IT", rows[5].B);

        Assert.Equal("Eric", rows[6].A);
        Assert.Equal("Eric", rows[7].A);
        Assert.Equal("IT", rows[7].B);
        Assert.Equal("Eric", rows[8].A);
        Assert.Equal("HR", rows[8].B);

        Assert.Equal("Keaton", rows[9].A);
        Assert.Equal("Keaton", rows[10].A);
        Assert.Equal("IT", rows[10].B);

        Assert.Equal("Felix", rows[11].A);
        Assert.Equal("Felix", rows[12].A);
        Assert.Equal("HR", rows[12].B);

        var dimension = Helpers.GetFirstSheetDimensionRefValue(path);
        Assert.Equal("A1:B20", dimension);
    }

    [Fact]
    public void TestGithubProject()
    {
        const string templatePath = "../../../../../samples/xlsx/TestTemplateGithubProjects.xlsx";
        var path = AutoDeletingPath.Create();

        var projects = new[]
        {
            new
            {
                Name = "MiniExcel",
                Link = "https://github.com/mini-software/MiniExcel",
                Star = 146,
                CreateTime = new DateTime(2021, 03, 01)
            },
            new
            {
                Name = "HtmlTableHelper",
                Link = "https://github.com/mini-software/HtmlTableHelper",
                Star = 16,
                CreateTime = new DateTime(2020, 02, 01)
            },
            new
            {
                Name = "PocoClassGenerator",
                Link = "https://github.com/mini-software/PocoClassGenerator",
                Star = 16,
                CreateTime = new DateTime(2019, 03, 17)
            }
        };
        var value = new
        {
            User = "ITWeiHan",
            Projects = projects,
            TotalStar = projects.Sum(s => s.Star)
        };
        Templater.ApplyXlsxTemplate(path.ToString(), templatePath, value);

        var rows = Importer.QueryXlsx(path.ToString()).ToList();
        Assert.Equal("ITWeiHan Github Projects", rows[0].B);
        Assert.Equal("Total Star : 178", rows[8].C);

        var dimension = Helpers.GetFirstSheetDimensionRefValue(path.ToString());
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
    public void TestIEnumerableType()
    {
        {
            const string templatePath = "../../../../../samples/xlsx/TestIEnumerableType.xlsx";
            using var path = AutoDeletingPath.Create();

            //1. By POCO
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
                Ts = new[]
                {
                    poco,
                    new TestIEnumerableTypePoco(),
                    null,
                    new TestIEnumerableTypePoco(),
                    poco
                }
            };
            Templater.ApplyXlsxTemplate(path.ToString(), templatePath, value);

            var rows = Importer.QueryXlsx<TestIEnumerableTypePoco>(path.ToString()).ToList();
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

            var dimension = Helpers.GetFirstSheetDimensionRefValue(path.ToString());
            Assert.Equal("A1:G6", dimension);
        }
    }

    [Fact]
    public void TestTemplateTypeMapping()
    {
        {
            const string templatePath = "../../../../../samples/xlsx/TestITemplateTypeAutoMapping.xlsx";
            using var path = AutoDeletingPath.Create();

            //1. By POCO
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
            Templater.ApplyXlsxTemplate(path.ToString(), templatePath, value);

            var rows = Importer.QueryXlsx<TestIEnumerableTypePoco>(path.ToString()).ToList();
            Assert.Equal(value.@string, rows[0].@string);
            Assert.Equal(value.@int, rows[0].@int);
            Assert.Equal(value.@double, rows[0].@double);
            Assert.Equal(value.@decimal, rows[0].@decimal);
            Assert.Equal(value.@bool, rows[0].@bool);
            Assert.Equal(value.datetime, rows[0].datetime);
            Assert.Equal(value.Guid, rows[0].Guid);

            var dimension = Helpers.GetFirstSheetDimensionRefValue(path.ToString());
            Assert.Equal("A1:G2", dimension);
        }
    }

    [Fact]
    public void TemplateCenterEmptyTest()
    {
        const string templatePath = "../../../../../samples/xlsx/TestTemplateCenterEmpty.xlsx";
        using var path = AutoDeletingPath.Create();
        var value = new
        {
            Tests = Enumerable.Range(1, 5).Select(i => new { test1 = i, test2 = i })
        };
        Templater.ApplyXlsxTemplate(path.ToString(), templatePath, value);
    }

    [Fact]
    public void TemplateBasicTest()
    {
        const string templatePath = "../../../../../samples/xlsx/TestTemplateEasyFill.xlsx";
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
            Templater.ApplyXlsxTemplate(path.ToString(), templatePath, value);

            var rows = Importer.QueryXlsx(path.ToString()).ToList();
            Assert.Equal("Jack", rows[1].A);
            Assert.Equal("2021-01-01 00:00:00", rows[1].B);
            Assert.Equal(true, rows[1].C);
            Assert.Equal(123, rows[1].D);
            Assert.Equal("Jack has 123 points", rows[1].E);

            var dimension = Helpers.GetFirstSheetDimensionRefValue(path.ToString());
            Assert.Equal("A1:E2", dimension);
        }

        {
            using var path = AutoDeletingPath.Create();
            var templateBytes = File.ReadAllBytes(templatePath);
            // 1. By POCO
            var value = new
            {
                Name = "Jack",
                CreateDate = new DateTime(2021, 01, 01),
                VIP = true,
                Points = 123
            };
            Templater.ApplyXlsxTemplate(path.ToString(), templateBytes, value);

            var rows = Importer.QueryXlsx(path.ToString()).ToList();
            Assert.Equal("Jack", rows[1].A);
            Assert.Equal("2021-01-01 00:00:00", rows[1].B);
            Assert.Equal(true, rows[1].C);
            Assert.Equal(123, rows[1].D);
            Assert.Equal("Jack has 123 points", rows[1].E);

            var dimension = Helpers.GetFirstSheetDimensionRefValue(path.ToString());
            Assert.Equal("A1:E2", dimension);
        }

        {
            var path = AutoDeletingPath.Create();
            var templateBytes = File.ReadAllBytes(templatePath);
            
            // 1. By POCO
            var value = new
            {
                Name = "Jack",
                CreateDate = new DateTime(2021, 01, 01),
                VIP = true,
                Points = 123
            };
            using (var stream = File.Create(path.ToString()))
            {
                Templater.ApplyXlsxTemplate(stream, templateBytes, value);
            }

            var rows = Importer.QueryXlsx(path.ToString()).ToList();
            Assert.Equal("Jack", rows[1].A);
            Assert.Equal("2021-01-01 00:00:00", rows[1].B);
            Assert.Equal(true, rows[1].C);
            Assert.Equal(123, rows[1].D);
            Assert.Equal("Jack has 123 points", rows[1].E);

            var dimension = Helpers.GetFirstSheetDimensionRefValue(path.ToString());
            Assert.Equal("A1:E2", dimension);
        }

        {
            using var path = AutoDeletingPath.Create();
            
            // 2. By Dictionary
            var value = new Dictionary<string, object>
            {
                ["Name"] = "Jack",
                ["CreateDate"] = new DateTime(2021, 01, 01),
                ["VIP"] = true,
                ["Points"] = 123
            };
            Templater.ApplyXlsxTemplate(path.ToString(), templatePath, value);

            var rows = Importer.QueryXlsx(path.ToString()).ToList();
            Assert.Equal("Jack", rows[1].A);
            Assert.Equal("2021-01-01 00:00:00", rows[1].B);
            Assert.Equal(true, rows[1].C);
            Assert.Equal(123, rows[1].D);
            Assert.Equal("Jack has 123 points", rows[1].E);

            var dimension = Helpers.GetFirstSheetDimensionRefValue(path.ToString());
            Assert.Equal("A1:E2", dimension);
        }
    }

    [Fact]
    public void TestIEnumerable()
    {
        const string templatePath = "../../../../../samples/xlsx/TestTemplateBasicIEmumerableFill.xlsx";
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
            Templater.ApplyXlsxTemplate(path.ToString(), templatePath, value);

            var dimension = Helpers.GetFirstSheetDimensionRefValue(path.ToString());
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
                    new { name = "Loan", department = "IT" }
                }
            };
            Templater.ApplyXlsxTemplate(path.ToString(), templatePath, value);

            var dimension = Helpers.GetFirstSheetDimensionRefValue(path.ToString());
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
            Templater.ApplyXlsxTemplate(path.ToString(), templatePath, value);

            var dimension = Helpers.GetFirstSheetDimensionRefValue(path.ToString());
            Assert.Equal("A1:B7", dimension);
        }
    }

    [Fact]
    public void TestIEnumerableWithFormulas()
    {
        const string templatePath = "../../../../../samples/xlsx/TestTemplateBasicIEnumerableFillWithFormulas.xlsx";
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
        Templater.ApplyXlsxTemplate(path.ToString(), templatePath, value);

        var dimension = Helpers.GetFirstSheetDimensionRefValue(path.ToString());
        Assert.Equal("A1:C13", dimension);
    }

    [Fact]
    public void TemplateTest()
    {
        const string templatePath = "../../../../../samples/xlsx/TestTemplateComplex.xlsx";
        {
            using var path = AutoDeletingPath.Create();

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
            Templater.ApplyXlsxTemplate(path.ToString(), templatePath, value);

            {
                var rows = Importer.QueryXlsx(path.ToString()).ToList();

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

                var dimension = Helpers.GetFirstSheetDimensionRefValue(path.ToString());
                Assert.Equal("A1:C9", dimension);
            }

            {
                var rows = Importer.QueryXlsx(path.ToString(), sheetName: "Sheet2").ToList();

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

                var dimension = Helpers.GetFirstSheetDimensionRefValue(path.ToString());
                Assert.Equal("A1:C9", dimension);
            }
        }


        {
            using var path = AutoDeletingPath.Create();

            // 2. By Dictionary
            var value = new Dictionary<string, object>()
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
            Templater.ApplyXlsxTemplate(path.ToString(), templatePath, value);

            var rows = Importer.QueryXlsx(path.ToString()).ToList();
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

            var dimension = Helpers.GetFirstSheetDimensionRefValue(path.ToString());
            Assert.Equal("A1:C9", dimension);
        }
    }

    [Fact]
    public void MergeSameCellsWithTagTest()
    {
        const string path = "../../../../../samples/xlsx/TestMergeWithTag.xlsx";
        using var mergedFilePath = AutoDeletingPath.Create();

        Templater.MergeSameCells(mergedFilePath.ToString(), path);
        var mergedCells = Helpers.GetFirstSheetMergedCells(mergedFilePath.ToString());

        Assert.Equal("A2:A4", mergedCells[0]);
        Assert.Equal("C3:C4", mergedCells[1]);
        Assert.Equal("A7:A8", mergedCells[2]);
    }

    [Fact]
    public void MergeSameCellsWithLimitTagTest()
    {
        const string path = "../../../../../samples/xlsx/TestMergeWithLimitTag.xlsx";
        using var mergedFilePath = AutoDeletingPath.Create();

        Templater.MergeSameCells(mergedFilePath.ToString(), path);
        var mergedCells = Helpers.GetFirstSheetMergedCells(mergedFilePath.ToString());

        Assert.Equal("A3:A4", mergedCells[0]);
        Assert.Equal("C3:C6", mergedCells[1]);
        Assert.Equal("A5:A6", mergedCells[2]);
    }
}