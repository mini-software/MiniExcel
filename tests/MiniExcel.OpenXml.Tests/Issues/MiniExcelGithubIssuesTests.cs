using System.Text.RegularExpressions;
using MiniExcelLib.Core.Enums;
using MiniExcelLib.Core.Exceptions;
using MiniExcelLib.OpenXml.Picture;
using MiniExcelLib.OpenXml.Tests.Utils;
using MiniExcelLib.Tests.Common.Utils;
using NPOI.XSSF.UserModel;
using LicenseContext = OfficeOpenXml.LicenseContext;

namespace MiniExcelLib.OpenXml.Tests.Issues;

public class MiniExcelGithubIssuesTests(ITestOutputHelper output)
{
    private readonly ITestOutputHelper _output = output;
    
    private readonly OpenXmlImporter _excelImporter =  MiniExcel.Importers.GetOpenXmlImporter();
    private readonly OpenXmlExporter _excelExporter =  MiniExcel.Exporters.GetOpenXmlExporter();
    private readonly OpenXmlTemplater _excelTemplater =  MiniExcel.Templaters.GetOpenXmlTemplater();

    static MiniExcelGithubIssuesTests()
    {
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
    }

    private static bool IsDateFormatString(string formatCode) => DateTimeHelper.IsDateTimeFormat(formatCode);

    [Fact]
    public void TestIssue_DataReaderSupportDimension()
    {
        using var table = new DataTable();

        table.Columns.Add("id", typeof(int));
        table.Columns.Add("name", typeof(string));
        table.Rows.Add(1, "Jack");
        table.Rows.Add(2, "Mike");

        using var path = AutoDeletingPath.Create();
        using var reader = table.CreateDataReader();
        var config = new OpenXmlConfiguration { FastMode = true };
        _excelExporter.Export(path.ToString(), reader, configuration: config);
        var xml = SheetHelper.GetZipFileContent(path.ToString(), "xl/worksheets/sheet1.xml");

        Assert.Contains("<x:autoFilter ref=\"A1:B3\" />", xml);
        Assert.Contains("<x:dimension ref=\"A1:B3\" />", xml);
    }

    [Fact]
    public void Issue87()
    {
        var templatePath = PathHelper.GetFile("xlsx/TestTemplateCenterEmpty.xlsx");
        using var path = AutoDeletingPath.Create();
        var value = new
        {
            Tests = Enumerable.Range(1, 5).Select((_, i) => new { test1 = i, test2 = i })
        };

        var rows = _excelImporter.Query(templatePath).ToList();
        _excelTemplater.FillTemplate(path.ToString(), templatePath, value);
    }

    [Fact]
    public void TestIssue117()
    {
        var cacheFull = new SharedStringsDiskCache(Path.GetTempPath());
        for (int i = 0; i < 100; i++)
        {
            cacheFull[i] = i.ToString();
        }
        for (int i = 0; i < 100; i++)
        {
            Assert.Equal(i.ToString(), cacheFull[i]);
        }
        Assert.Equal(100, cacheFull.Count);

        var cacheEmpty = new SharedStringsDiskCache(Path.GetTempPath());
        Assert.Empty(cacheEmpty);
    }

    // Query Merge cells data
    [Fact]
    public void Issue122()
    {
        var config = new OpenXmlConfiguration
        {
            FillMergedCells = true
        };

        var path1 = PathHelper.GetFile("xlsx/TestIssue122.xlsx");
        var rows1 = _excelImporter.Query(path1, hasHeaderRow: true, configuration: config).ToList();
        Assert.Equal("HR", rows1[0].Department);
        Assert.Equal("HR", rows1[1].Department);
        Assert.Equal("HR", rows1[2].Department);
        Assert.Equal("IT", rows1[3].Department);
        Assert.Equal("IT", rows1[4].Department);
        Assert.Equal("IT", rows1[5].Department);

        var path2 = PathHelper.GetFile("xlsx/TestIssue122_2.xlsx");
        var rows2 = _excelImporter.Query(path2, hasHeaderRow: true, configuration: config).ToList();
        Assert.Equal("V1", rows2[2].Test1);
        Assert.Equal("V2", rows2[5].Test2);
        Assert.Equal("V3", rows2[1].Test3);
        Assert.Equal("V4", rows2[2].Test4);
        Assert.Equal("V5", rows2[3].Test5);
        Assert.Equal("V6", rows2[5].Test5);
    }

    [Fact]
    public void Issue132()
    {
        {
            using var path = AutoDeletingPath.Create();
            var value = new[] {
                new { name = "Jack", Age = 25, InDate = new DateTime(2021,01,03)},
                new { name = "Henry", Age = 36, InDate = new DateTime(2020,05,03)},
            };

            _excelExporter.Export(path.ToString(), value);
        }

        {
            using var path = AutoDeletingPath.Create();
            var value = new[]
            {
                new { name = "Jack", Age = 25, InDate = new DateTime(2021,01,03)},
                new { name = "Henry", Age = 36, InDate = new DateTime(2020,05,03)},
            };
            var config = new OpenXmlConfiguration
            {
                TableStyles = TableStyles.None
            };
            _excelExporter.Export(path.ToString(), value, configuration: config);
        }

        {
            using var path = AutoDeletingPath.Create();

            var dt = new DataTable();
            dt.Columns.Add("Name");
            dt.Columns.Add("Age");
            dt.Columns.Add("Date");

            dt.Rows.Add("Jack", 25, new DateTime(2021, 01, 03));
            dt.Rows.Add("Henry", 36, new DateTime(2021, 01, 03));

            _excelExporter.Export(path.ToString(), dt);
        }
    }

    [Fact]
    public void TestIssues133()
    {
        {
            using var path = AutoDeletingPath.Create();

            var value = new DataTable();
            value.Columns.Add("Id");
            value.Columns.Add("Name");
            _excelExporter.Export(path.ToString(), value);
            var rows = _excelImporter.Query(path.ToString()).ToList();

            Assert.Equal("Id", rows[0].A);
            Assert.Equal("Name", rows[0].B);
            Assert.Single(rows);
            Assert.Equal("A1:B1", SheetHelper.GetFirstSheetDimensionRefValue(path.ToString()));
        }

        {
            using var path = AutoDeletingPath.Create();

            var value = Array.Empty<TestIssues133Dto>();
            _excelExporter.Export(path.ToString(), value);
            var rows = _excelImporter.Query(path.ToString()).ToList();

            Assert.Equal("Id", rows[0].A);
            Assert.Equal("Name", rows[0].B);
            Assert.Single(rows);
            Assert.Equal("A1:B1", SheetHelper.GetFirstSheetDimensionRefValue(path.ToString()));
        }
    }

    [Fact]
    public void Issue137()
    {
        var path = PathHelper.GetFile("xlsx/TestIssue137.xlsx");
        {
            var rows = _excelImporter.Query(path).ToList();
            var first = (IDictionary<string, object>)rows[0]; // https://user-images.githubusercontent.com/12729184/113266322-ba06e400-9307-11eb-9521-d36abfda75cc.png
            Assert.Equal(["A", "B", "C", "D", "E", "F", "G", "H"], first.Keys.ToArray());
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
            var rows = _excelImporter.Query(path, true).ToList();
            var first = (IDictionary<string, object>)rows[0]; //![image](https://user-images.githubusercontent.com/12729184/113266322-ba06e400-9307-11eb-9521-d36abfda75cc.png)
            Assert.Equal(["比例", "商品", "滿倉口數", "0", "1為港幣 0為台幣"], first.Keys.ToArray());
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
            var rows = _excelImporter.Query<Issue137Dto>(path).ToList();
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

    [Fact]
    public void Issue138()
    {
        var path = PathHelper.GetFile("xlsx/TestIssue138.xlsx");
        var rows1 = _excelImporter.Query(path, true).ToList();
        Assert.Equal(6, rows1.Count);

        foreach (var index in new[] { 0, 2, 5 })
        {
            Assert.Equal(1, rows1[index].實單每日損益);
            Assert.Equal(2, rows1[index].程式每日損益);
            Assert.Equal("測試商品1", rows1[index].商品);
            Assert.Equal(111.11, rows1[index].滿倉口數);
            Assert.Equal(111.11, rows1[index].波段);
            Assert.Equal(111.11, rows1[index].當沖);
        }

        foreach (var index in new[] { 1, 3, 4 })
        {
            Assert.Null(rows1[index].實單每日損益);
            Assert.Null(rows1[index].程式每日損益);
            Assert.Null(rows1[index].商品);
            Assert.Null(rows1[index].滿倉口數);
            Assert.Null(rows1[index].波段);
            Assert.Null(rows1[index].當沖);
        }

        var rows2 = _excelImporter.Query<Issue138Dto>(path).ToList();
        Assert.Equal(6, rows2.Count);
        Assert.Equal(new DateTime(2021, 3, 1), rows2[0].Date);

        foreach (var index in new[] { 0, 2, 5 })
        {
            Assert.Equal(1, rows2[index].實單每日損益);
            Assert.Equal(2, rows2[index].程式每日損益);
            Assert.Equal("測試商品1", rows2[index].商品);
            Assert.Equal(111.11, rows2[index].滿倉口數);
            Assert.Equal(111.11, rows2[index].波段);
            Assert.Equal(111.11, rows2[index].當沖);
        }

        foreach (var index in new[] { 1, 3, 4 })
        {
            Assert.Null(rows2[index].實單每日損益);
            Assert.Null(rows2[index].程式每日損益);
            Assert.Null(rows2[index].商品);
            Assert.Null(rows2[index].滿倉口數);
            Assert.Null(rows2[index].波段);
            Assert.Null(rows2[index].當沖);
        }
    }

    [Fact]
    public void Issue149()
    {
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
        var strings = chars.Select(s => s.ToString()).ToArray();

        {
            var path = PathHelper.GetFile("xlsx/TestIssue149.xlsx");
            var rows = _excelImporter.Query(path).Select(s => (string)s.A).ToList();
            for (int i = 0; i < chars.Length; i++)
            {
                //output.WriteLine($"{i} , {chars[i]} , {rows[i]}");
                if (i == 13)
                    continue;

                Assert.Equal(strings[i], rows[i]);
            }
        }

        {
            using var path = AutoDeletingPath.Create();
            var input = chars.Select(s => new { Test = s.ToString() });
            _excelExporter.Export(path.ToString(), input);

            var rows = _excelImporter.Query(path.ToString(), true).Select(s => (string)s.Test).ToList();
            for (int i = 0; i < chars.Length; i++)
            {
                _output.WriteLine($"{i}, {chars[i]}, {rows[i]}");
                if (i is 13 or 9 or 10)
                    continue;

                Assert.Equal(strings[i], rows[i]);
            }
        }

        {
            using var path = AutoDeletingPath.Create();
            var input = chars.Select(s => new { Test = s.ToString() });
            _excelExporter.Export(path.ToString(), input);

            var rows = _excelImporter.Query<Issue149VO>(path.ToString()).Select(s => s.Test).ToList();
            for (int i = 0; i < chars.Length; i++)
            {
                _output.WriteLine($"{i}, {chars[i]}, {rows[i]}");
                if (i is 13 or 9 or 10)
                    continue;

                Assert.Equal(strings[i], rows[i]);
            }
        }
    }

    [Fact]
    public void Issue150()
    {
        var path = Path.Combine(Path.GetTempPath(), $"{Guid.NewGuid()}.xlsx");

        Assert.Throws<NotSupportedException>(() =>  _excelExporter.Export(path, new[] { 1, 2 }));
        File.Delete(path);

        Assert.Throws<NotSupportedException>(() =>  _excelExporter.Export(path, new[] { "1", "2" }));
        File.Delete(path);

        Assert.Throws<NotSupportedException>(() =>  _excelExporter.Export(path, new[] { '1', '2' }));
        File.Delete(path);

        Assert.Throws<NotSupportedException>(() =>  _excelExporter.Export(path, new[] { DateTime.Now }));
        File.Delete(path);

        Assert.Throws<NotSupportedException>(() =>  _excelExporter.Export(path, new[] { Guid.NewGuid() }));
        File.Delete(path);
    }

    [Fact]
    public void Issue153()
    {
        var path = PathHelper.GetFile("xlsx/TestIssue153.xlsx");
        var rows = _excelImporter.Query(path, true).First() as IDictionary<string, object>;
        Assert.Equal(
        [
            "序号", "代号", "新代号", "名称", "XXX", "部门名称", "单位", "ERP工时   (小时)A", "工时(秒) A/3600", "标准人工工时(秒)",
            "生产标准机器工时(秒)", "财务、标准机器工时(秒)", "更新日期", "产品机种", "备注", "最近一次修改前的标准工时(秒)", "最近一次修改前的标准机时(秒)", "备注1"
        ], rows?.Keys);
    }

    [Fact]
    public void Issue157()
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

            _excelExporter.Export(path, data);

            var rows = _excelImporter.Query(path, sheetName: "Sheet1").ToList();
            Assert.Equal(6, rows.Count);
            Assert.Equal("Sheet1", _excelImporter.GetSheetNames(path).First());

            using var p = new ExcelPackage(new FileInfo(path));
            var ws = p.Workbook.Worksheets.First();
            Assert.Equal("Sheet1", ws.Name);
            Assert.Equal("Sheet1", p.Workbook.Worksheets["Sheet1"].Name);
        }
        {
            var path = PathHelper.GetFile("xlsx/TestIssue157.xlsx");

            {
                var rows = _excelImporter.Query(path, sheetName: "Sheet1").ToList();
                Assert.Equal(6, rows.Count);
                Assert.Equal("Sheet1", _excelImporter.GetSheetNames(path).First());
            }
            using (var p = new ExcelPackage(new FileInfo(path)))
            {
                var ws = p.Workbook.Worksheets.First();
                Assert.Equal("Sheet1", ws.Name);
                Assert.Equal("Sheet1", p.Workbook.Worksheets["Sheet1"].Name);
            }

            {
                var rows = _excelImporter.Query<UserAccount>(path, sheetName: "Sheet1").ToList();
                Assert.Equal(5, rows.Count);

                Assert.Equal(new Guid("78DE23D2-DCB6-BD3D-EC67-C112BBC322A2"), rows[0].ID);
                Assert.Equal("Wade", rows[0].Name);
                Assert.Equal(new DateTime(2020,9,27), rows[0].BoD);
                Assert.False(rows[0].VIP);
                Assert.Equal(5019.12m, rows[0].Points);
                Assert.Equal(1, rows[0].IgnoredProperty);
            }
        }
    }

    // Query Support StartCell
    [Fact]
    public void Issue147()
    {
        var path1 = PathHelper.GetFile("xlsx/TestIssue147.xlsx");
        var rows1 = _excelImporter.Query(path1, hasHeaderRow: false, startCell: "C3", sheetName: "Sheet1").ToList();

        Assert.Equal(["C", "D", "E"], (rows1[0] as IDictionary<string, object>)?.Keys);
        Assert.Equal(new[]{"Column1", "Column2", "Column3"}, new[] { rows1[0].C as string, rows1[0].D as string, rows1[0].E as string });
        Assert.Equal(new[]{"C4", "D4", "E4"}, new[] { rows1[1].C as string, rows1[1].D as string, rows1[1].E as string });
        Assert.Equal(new[]{"C9", "D9", "E9"}, new[] { rows1[6].C as string, rows1[6].D as string, rows1[6].E as string });
        Assert.Equal(new[]{"C12", "D12", "E12"}, new[] { rows1[9].C as string, rows1[9].D as string, rows1[9].E as string });
        Assert.Equal(new[]{"C13", "D13", "E13"}, new[] { rows1[10].C as string, rows1[10].D as string, rows1[10].E as string });
        foreach (var i in new[] { 4, 5, 7, 8 })
            Assert.Equal(expected: [null, null, null], new[] { rows1[i].C as string, rows1[i].D as string, rows1[i].E as string });

        Assert.Equal(11, rows1.Count);


        var columns1 = _excelImporter.GetColumnNames(path1, startCell: "C3");
        Assert.Equal(["C", "D", "E"], columns1);

        var path2 = PathHelper.GetFile("xlsx/TestIssue147.xlsx");
        var rows2 = _excelImporter.Query(path2, hasHeaderRow: true, startCell: "C3", sheetName: "Sheet1").ToList();

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

        var columns2 = _excelImporter.GetColumnNames(path2, hasHeaderRow: true, startCell: "C3");
        Assert.Equal(["Column1", "Column2", "Column3"], columns2);
    }

    [Fact]
    public void TestIssue186()
    {
        var originPath = PathHelper.GetFile("xlsx/TestIssue186_Template.xlsx");
        using var path = AutoDeletingPath.Create();
        File.Copy(originPath, path.FilePath);

        MiniExcelPicture[] images =
        [
            new()
            {
                ImageBytes = File.ReadAllBytes(PathHelper.GetFile("images/github_logo.png")),
                SheetName = null, // default null is first sheet
                CellAddress = "C3", // required
            },
            new()
            {
                ImageBytes = File.ReadAllBytes(PathHelper.GetFile("images/google_logo.png")),
                PictureType = "image/png", // default PictureType = image/png
                SheetName = "Demo",
                CellAddress = "C9", // required
                WidthPx = 100,
                HeightPx = 100
            }
        ];

        _excelTemplater.AddPicture(path.FilePath, images);
    }

    // SaveAs default theme support filter mode - https://github.com/mini-software/MiniExcel/issues/190
    [Fact]
    public void TestIssue190()
    {
        {
            using var path = AutoDeletingPath.Create();
            var value = new TestIssue190Dto[] { };
            _excelExporter.Export(path.ToString(), value, configuration: new OpenXmlConfiguration { AutoFilter = false });

            var sheetXml = SheetHelper.GetZipFileContent(path.ToString(), "xl/worksheets/sheet1.xml");
            Assert.DoesNotContain("<x:autoFilter ref=\"A1:C1\" />", sheetXml);
        }
        {
            using var path = AutoDeletingPath.Create();
            var value = new TestIssue190Dto[] { };
            _excelExporter.Export(path.ToString(), value);

            var sheetXml = SheetHelper.GetZipFileContent(path.ToString(), "xl/worksheets/sheet1.xml");
            Assert.Contains("<x:autoFilter ref=\"A1:C1\" />", sheetXml);
        }
        {
            using var path = AutoDeletingPath.Create();
            TestIssue190Dto[] value =
            [
                new() { ID = 1, Name = "Jack", Age = 32 },
                new() { ID = 2, Name = "Lisa", Age = 45 }
            ];
            _excelExporter.Export(path.ToString(), value);

            var sheetXml = SheetHelper.GetZipFileContent(path.ToString(), "xl/worksheets/sheet1.xml");
            Assert.Contains("<x:autoFilter ref=\"A1:C3\" />", sheetXml);
        }
    }

    [Fact]
    public void Issue193()
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
             _excelTemplater.FillTemplate(path, templatePath, value);

            foreach (var sheetName in _excelImporter.GetSheetNames(path))
            {
                var rows = _excelImporter.Query(path, sheetName: sheetName).ToList();
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

                /*TODO:row can't contain xmlns*/
                // https://user-images.githubusercontent.com/12729184/114998840-ead44500-9ed3-11eb-8611-58afb98faed9.png

            }
        }

        {
            var templatePath = PathHelper.GetFile("xlsx/TestTemplateComplex.xlsx");
            using var path = AutoDeletingPath.Create();

            // 2. By Dictionary
            var value = new Dictionary<string, object>
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
             _excelTemplater.FillTemplate(path.ToString(), templatePath, value);
            var rows = _excelImporter.Query(path.ToString()).ToList();

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
    public void Issue206()
    {
        {
            var templatePath = PathHelper.GetFile("xlsx/TestTemplateBasicIEmumerableFill.xlsx");
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
            _excelTemplater.FillTemplate(path.ToString(), templatePath, value);

            var dimension = SheetHelper.GetFirstSheetDimensionRefValue(path.ToString());
            Assert.Equal("A1:B2", dimension);
        }

        {
            var templatePath = PathHelper.GetFile("xlsx/TestTemplateBasicIEmumerableFill.xlsx");
            using var path = AutoDeletingPath.Create();

            using var dt = new DataTable();
            dt.Columns.Add("name");
            dt.Columns.Add("department");
            dt.Rows.Add("Jack", "HR");

            var value = new Dictionary<string, object> { ["employees"] = dt };
            _excelTemplater.FillTemplate(path.ToString(), templatePath, value);

            var dimension = SheetHelper.GetFirstSheetDimensionRefValue(path.ToString());
            Assert.Equal("A1:B2", dimension);
        }
    }
    
    // Template merge row list rendering has no merge
    [Fact]
    public void Issue207()
    {
        {
            var tempaltePath = PathHelper.GetFile("xlsx/TestIssue207_2.xlsx");
            using var file = AutoDeletingPath.Create();
            var path = file.ToString();

            var value = new
            {
                project = new[] {
                    new {name = "項目1",content="[]內容1,[]內容2,[]內容3,[]內容4,[]內容5"},
                    new {name = "項目2",content="[]內容1,[]內容2,[]內容3,[]內容4,[]內容5"},
                    new {name = "項目3",content="[]內容1,[]內容2,[]內容3,[]內容4,[]內容5"},
                    new {name = "項目4",content="[]內容1,[]內容2,[]內容3,[]內容4,[]內容5"},
                }
            };

             _excelTemplater.FillTemplate(path, tempaltePath, value);
            var rows = _excelImporter.Query(path).ToList();

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

            var dimension = SheetHelper.GetFirstSheetDimensionRefValue(path);
            Assert.Equal("A1:C16", dimension);
        }

        {
            var tempaltePath = PathHelper.GetFile("xlsx/TestIssue207_Template_Merge_row_list_rendering_without_merge/template.xlsx");
            using var file = AutoDeletingPath.Create();
            var path = file.ToString();

            var value = new
            {
                project = new[] {
                    new {name = "項目1",content="[]內容1,[]內容2,[]內容3,[]內容4,[]內容5"},
                    new {name = "項目2",content="[]內容1,[]內容2,[]內容3,[]內容4,[]內容5"},
                    new {name = "項目3",content="[]內容1,[]內容2,[]內容3,[]內容4,[]內容5"},
                    new {name = "項目4",content="[]內容1,[]內容2,[]內容3,[]內容4,[]內容5"},
                }
            };

             _excelTemplater.FillTemplate(path, tempaltePath, value);
            var rows = _excelImporter.Query(path).ToList();

            Assert.Equal("項目1", rows[0].A);
            Assert.Equal("[]內容1,[]內容2,[]內容3,[]內容4,[]內容5", rows[0].C);
            Assert.Equal("項目2", rows[3].A);
            Assert.Equal("[]內容1,[]內容2,[]內容3,[]內容4,[]內容5", rows[3].C);
            Assert.Equal("項目3", rows[6].A);
            Assert.Equal("[]內容1,[]內容2,[]內容3,[]內容4,[]內容5", rows[6].C);
            Assert.Equal("項目4", rows[9].A);
            Assert.Equal("[]內容1,[]內容2,[]內容3,[]內容4,[]內容5", rows[9].C);
            var dimension = SheetHelper.GetFirstSheetDimensionRefValue(path);
            Assert.Equal("A1:E15", dimension);
        }
    }

    [Fact]
    public void Issue208()
    {
        var path = PathHelper.GetFile("xlsx/TestIssue208.xlsx");
        var columns = _excelImporter.GetColumnNames(path).ToList();
        Assert.Equal(16384, columns.Count);
        Assert.Equal("XFD", columns[16383]);
    }

    // Query type conversion error - https://github.com/mini-software/MiniExcel/issues/309
    [Fact]
    public void TestIssue209()
    {
        try
        {
            var path = PathHelper.GetFile("xlsx/TestIssue309.xlsx");
            _ = _excelImporter.Query<TestIssue209Dto>(path).ToList();
        }
        catch (ValueNotAssignableException ex)
        {
            Assert.Equal("SEQ", ex.ColumnName);
            Assert.Equal(4, ex.Row);
            Assert.Equal("Error", ex.Value);
            Assert.Equal(typeof(int), ex.ColumnType);
            Assert.Equal("The value Error cannot be assigned to type Int32.", ex.Message);
        }
    }

    // SaveAs support for IDataReader based export
    [Fact]
    public void Issue211()
    {
        using var path = AutoDeletingPath.Create();
        var tempSqlitePath = AutoDeletingPath.Create(Path.GetTempPath(), $"{Guid.NewGuid()}.db");
        var connectionString = $"Data Source={tempSqlitePath};Version=3;";

        using var connection = new SQLiteConnection(connectionString);
        var reader = connection.ExecuteReader(@"select 1 Test1,2 Test2 union all select 3 , 4 union all select 5 ,6");
        _excelExporter.Export(path.ToString(), reader);
        var rows = _excelImporter.Query(path.ToString(), true).ToList();

        Assert.Equal(1.0, rows[0].Test1);
        Assert.Equal(2.0, rows[0].Test2);
        Assert.Equal(3.0, rows[1].Test1);
        Assert.Equal(4.0, rows[1].Test2);
    }

    // _exporter.Export(path, table,sheetName:“Name”) sheetName is incorrectly Sheet1
    [Fact]
    public void Issue212()
    {
        using var path = AutoDeletingPath.Create();
        _excelExporter.Export(path.ToString(), new[] { new { x = 1, y = 2 } }, sheetName: "Demo");

        var actualSheetName = _excelImporter.GetSheetNames(path.ToString()).ToList()[0];
        Assert.Equal("Demo", actualSheetName);
    }

    // Support reading Excel by IDataReader and returning DataTable
    [Fact]
    public void Issue216()
    {
        using var path = AutoDeletingPath.Create();
        var value = new[]
        {
            new { Test1 = "1", Test2 = 2 },
            new { Test1 = "3", Test2 = 4 }
        };
        _excelExporter.Export(path.ToString(), value);

        using var table = _excelImporter.QueryAsDataTable(path.ToString());
        Assert.Equal("Test1", table.Columns[0].ColumnName);
        Assert.Equal("Test2", table.Columns[1].ColumnName);
        Assert.Equal("1", table.Rows[0]["Test1"]);
        Assert.Equal(2.0, table.Rows[0]["Test2"]);
        Assert.Equal("3", table.Rows[1]["Test1"]);
        Assert.Equal(4.0, table.Rows[1]["Test2"]);


        using var dt = _excelImporter.QueryAsDataTable(path.ToString(), false);
        Assert.Equal("Test1", dt.Rows[0]["A"]);
        Assert.Equal("Test2", dt.Rows[0]["B"]);
        Assert.Equal("1", dt.Rows[1]["A"]);
        Assert.Equal(2.0, dt.Rows[1]["B"]);
        Assert.Equal("3", dt.Rows[2]["A"]);
        Assert.Equal(4.0, dt.Rows[2]["B"]);
    }

    // DataTable recommended to use Caption for column name first, then use columname
    [Fact]
    public void Issue217()
    {
        using var table = new DataTable();
        table.Columns.Add("CustomerID");
        table.Columns.Add("CustomerName").Caption = "Name";
        table.Columns.Add("CreditLimit").Caption = "Limit";
        table.Rows.Add(1, "Jonathan", 23.44);
        table.Rows.Add(2, "Bill", 56.87);

        using var path = AutoDeletingPath.Create();
        _excelExporter.Export(path.ToString(), table);

        var rows = _excelImporter.Query(path.ToString()).ToList();
        Assert.Equal("Name", rows[0].B);
        Assert.Equal("Limit", rows[0].C);
    }

    // Dynamic Query can't summary numeric cell value default, need to cast
    [Fact]
    public void Issue220()
    {
        var path = PathHelper.GetFile("xlsx/TestIssue220.xlsx");
        var rows = _excelImporter.Query(path, hasHeaderRow: true);
        var result = rows
            .GroupBy(s => s.PRT_ID)
            .Select(g => new
            {
                PRT_ID = g.Key,
                Apr = g.Sum(x => (double?)x.Apr),
                May = g.Sum(x => (double?)x.May),
                Jun = g.Sum(x => (double?)x.Jun)
            })
            .ToList();

        Assert.Equal(91843.25, result[0].Jun);
        Assert.Equal(50000.99, result[1].Jun);
    }

    // Custom yyyy-MM-dd format not converted to datetime
    [Fact]
    public void Issue222()
    {
        var path = PathHelper.GetFile("xlsx/TestIssue222.xlsx");
        var rows = _excelImporter.Query(path).ToList();
        Assert.Equal(typeof(DateTime), rows[1].A.GetType());
        Assert.Equal(new DateTime(2021, 4, 29), rows[1].A);
    }

    [Fact]
    public void Issue223()
    {
        List<Dictionary<string, object?>> value =
        [
            new() { { "A", null }, { "B", null } },
            new() { { "A", 123 }, { "B", new DateTime(2021, 1, 1) } },
            new() { { "A", Guid.NewGuid() }, { "B", "HelloWorld" } }
        ];
        using var path = AutoDeletingPath.Create();
        _excelExporter.Export(path.ToString(), value);
        
        using var dt = _excelImporter.QueryAsDataTable(path.ToString());

        var columns = dt.Columns;
        Assert.Equal(typeof(object), columns[0].DataType);
        Assert.Equal(typeof(object), columns[1].DataType);

        Assert.Equal(123.0, dt.Rows[1]["A"]);
        Assert.Equal("HelloWorld", dt.Rows[2]["B"]);
    }

    // Fix SaveAsByTemplate single column dimension index error
    [Fact]
    public void Issue226()
    {
        using var path = AutoDeletingPath.Create();
        var templatePath = PathHelper.GetFile("xlsx/TestIssue226.xlsx");
        _excelTemplater.FillTemplate(path.ToString(), templatePath, new { employees = new[] { new { name = "123" }, new { name = "123" } } });
        Assert.Equal("A1:A3", SheetHelper.GetFirstSheetDimensionRefValue(path.ToString()));
    }

    // Support Xlsm AutoCheck
    [Fact]
    public void Issue227()
    {
        var path = PathHelper.GetTempPath("xlsm");
        Assert.Throws<NotSupportedException>(() =>  _excelExporter.Export(path, new[] { new { V = "A1" }, new { V = "A2" } }));
        File.Delete(path);

        var path1 = PathHelper.GetFile("xlsx/TestIssue227.xlsm");
        var rows1 = _excelImporter.Query<UserAccount>(path1).ToList();
        Assert.Equal(100, rows1.Count);

        Assert.Equal(Guid.Parse("78DE23D2-DCB6-BD3D-EC67-C112BBC322A2"), rows1[0].ID);
        Assert.Equal("Wade", rows1[0].Name);
        Assert.Equal(DateTime.ParseExact("27/09/2020", "dd/MM/yyyy", CultureInfo.InvariantCulture), rows1[0].BoD);
        Assert.Equal(36, rows1[0].Age);
        Assert.False(rows1[0].VIP);
        Assert.Equal(5019.12m, rows1[0].Points);
        Assert.Equal(1, rows1[0].IgnoredProperty);

        using var stream = File.OpenRead(path1);
        var rows2 = _excelImporter.Query<UserAccount>(stream).ToList();
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
    public void Issue229()
    {
        var path = PathHelper.GetFile("xlsx/TestIssue229.xlsx");

        using var dt = _excelImporter.QueryAsDataTable(path);

        foreach (DataColumn column in dt.Columns)
        {
            var v = dt.Rows[3][column];
            Assert.Equal(DBNull.Value, v);
        }
    }

    // SaveAs By data reader error : 'Invalid attempt to call FieldCount when reader is closed'
    [Fact]
    public void Issue230()
    {
        using var conn = Db.GetConnection("Data Source=:memory:");
        conn.Open();
        using var cmd = conn.CreateCommand();
        cmd.CommandText = "select 1 id union all select 2";

        using (var reader = cmd.ExecuteReader(CommandBehavior.CloseConnection))
        {
            while (reader.Read())
            {
                for (int i = 0; i < reader.FieldCount; i++)
                {
                    var result = $"{reader.GetName(i)} , {reader.GetValue(i)}";
                    _output.WriteLine(result);
                }
            }
        }

        using var conn2 = Db.GetConnection("Data Source=:memory:");
        conn2.Open();
        using var cmd2 = conn2.CreateCommand();
        cmd2.CommandText = "select 1 id union all select 2";
        using (var reader = cmd2.ExecuteReader(CommandBehavior.CloseConnection))
        {
            while (reader.Read())
            {
                for (int i = 0; i < reader.FieldCount; i++)
                {
                    var result = $"{reader.GetName(i)}, {reader.GetValue(i)}";
                    _output.WriteLine(result);
                }
            }
        }

        using var conn3 = Db.GetConnection("Data Source=:memory:");
        conn3.Open();
        using var cmd3 = conn3.CreateCommand();
        cmd3.CommandText = "select 1 id union all select 2";
        using (var reader = cmd3.ExecuteReader(CommandBehavior.CloseConnection))
        {
            using var path = AutoDeletingPath.Create();
            _excelExporter.Export(path.ToString(), reader, printHeader: true);
            var rows = _excelImporter.Query(path.ToString(), true).ToList();
            Assert.Equal(1, rows[0].id);
            Assert.Equal(2, rows[1].id);
        }
    }

    // QueryAsDataTable A2=5.5 , A3=0.55/1.1 will case double type check error
    [Fact]
    public void Issue233()
    {
        var path = PathHelper.GetFile("xlsx/TestIssue233.xlsx");

        var dt = _excelImporter.QueryAsDataTable(path);

        var rows = dt.Rows;

        Assert.Equal(0.55, rows[0]["Size"]);
        Assert.Equal("0.55/1.1", rows[1]["Size"]);
    }

    /// SaveAs support for multiple sheets
    [Fact]
    public void Issue234()
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
        _excelExporter.Export(path, sheets);

        var sheetNames = _excelImporter.GetSheetNames(path);
        Assert.Equal("users", sheetNames[0]);
        Assert.Equal("department", sheetNames[1]);

        {
            var rows = _excelImporter.Query(path, true, sheetName: "users").ToList();
            Assert.Equal("Jack", rows[0].Name);
            Assert.Equal(25, rows[0].Age);
            Assert.Equal("Mike", rows[1].Name);
            Assert.Equal(44, rows[1].Age);
        }
        {
            var rows = _excelImporter.Query(path, true, sheetName: "department").ToList();
            Assert.Equal("01", rows[0].ID);
            Assert.Equal("HR", rows[0].Name);
            Assert.Equal("02", rows[1].ID);
            Assert.Equal("IT", rows[1].Name);
        }
    }

    // Support SaveAs by DataSet
    [Fact]
    public void Issue235()
    {
        using var path = AutoDeletingPath.Create();
        
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

        DataSet dataSet = new();
        dataSet.Tables.Add(users);
        dataSet.Tables.Add(departments);

        var rowsWritten = _excelExporter.Export(path.ToString(), dataSet);
        Assert.Equal(2, rowsWritten.Length);
        Assert.Equal(2, rowsWritten[0]);

        var sheetNames = _excelImporter.GetSheetNames(path.ToString());
        Assert.Equal("users", sheetNames[0]);
        Assert.Equal("departments", sheetNames[1]);

        var rows1 = _excelImporter.Query(path.ToString(), true, sheetName: "users").ToList();
        Assert.Equal("Jack", rows1[0].Name);
        Assert.Equal(25, rows1[0].Age);
        Assert.Equal("Mike", rows1[1].Name);
        Assert.Equal(44, rows1[1].Age);

        var rows2 = _excelImporter.Query(path.ToString(), true, sheetName: "departments").ToList();
        Assert.Equal("01", rows2[0].ID);
        Assert.Equal("HR", rows2[0].Name);
        Assert.Equal("02", rows2[1].ID);
        Assert.Equal("IT", rows2[1].Name);
    }
    
    [Fact]
    public void TestIssue240()
    {
        using var path = AutoDeletingPath.Create();
        var templatePath = PathHelper.GetFile("xlsx/TestIssue24020201.xlsx");
        var data = new Dictionary<string, object>
        {
            ["title"] = "This's title",
            ["B"] = new List<Dictionary<string, object>>
            {
                new() { { "specialMark", 1 } },
                new() { { "specialMark", 2 } },
                new() { { "specialMark", 3 } },
            }
        };
         _excelTemplater.FillTemplate(path.ToString(), templatePath, data);
    }

    // Support Custom Datetime format
    [Fact]
    public void Issue241()
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
        var rowsWritten =  _excelExporter.Export(path, value);
            
        Assert.Single(rowsWritten);
        Assert.Equal(2, rowsWritten[0]);

        using var package = new ExcelPackage(path);
        var cells = package.Workbook.Worksheets.First().Cells;

        Assert.Equal(date1, DateTime.FromOADate((double)cells["B2"].Value));
        Assert.Equal("01 04, 2021", cells["B2"].Text);
        Assert.Equal(date2, DateTime.FromOADate((double)cells["B3"].Value));
        Assert.Equal("04 05, 2020", cells["B3"].Text);
    }

    // No error exception throw when reading xls file #242
    [Fact]
    public void Issue242()
    {
        var path = PathHelper.GetFile("xls/TestIssue242.xls");
        Assert.Throws<InvalidDataException>(() => _excelImporter.Query(path).ToList());

        using var stream = File.OpenRead(path);
        Assert.Throws<InvalidDataException>(() => _excelImporter.Query(stream).ToList());
    }

    // SaveAsByTemplate support DateTime custom format
    [Fact]
    public void Issue255()
    {
        var dt1 = new DateTime(2021, 01, 01);
        var dt2 = new DateTime(2022, 01, 01);
       
        //template
        {
            var templatePath = PathHelper.GetFile("xlsx/TestsIssue255_Template.xlsx");
            using var ms = new MemoryStream();
            var value = new
            {
                Issue255DTO = new[] { new Issue255DTO { Time = dt1, Time2 = dt2 } }
            };
            
            _excelTemplater.FillTemplate(ms, templatePath, value);

            ms.Seek(0, SeekOrigin.Begin);
            using var package = new ExcelPackage(ms);
            var cells = package.Workbook.Worksheets[0].Cells;

            Assert.Equal("2021", cells["A2"].Text);
            Assert.Equal("2022", cells["B2"].Text);
        }
        //export
        {
            using var ms = new MemoryStream();
            Issue255DTO[] value = 
            [
                new() { Time = dt1, Time2 = dt2 }
            ];

            var rowsWritten = _excelExporter.Export(ms, value);
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

    // Dynamic Query custom format not using mapping format
    [Fact]
    public void Issue256()
    {
        var path = PathHelper.GetFile("xlsx/TestIssue256.xlsx");
        var rows = _excelImporter.Query(path, false).ToList();
        Assert.Equal(new DateTime(2003, 4, 16), rows[1].A);
        Assert.Equal(new DateTime(2004, 4, 16), rows[1].B);
    }

    // custom format contains specific format (eg:`#,##0.000_);[Red]\(#,##0.000\)`), automatic converter will convert double to datetime
    [Fact]
    public void TestIssue267()
    {
        var path = PathHelper.GetFile("/xlsx/TestIssue267.xlsx");
        var row = _excelImporter.Query(path).SingleOrDefault();
        Assert.Equal(10618, row!.A);
        Assert.Equal("2021-02-23", row.B);
        Assert.Equal(43.199999999999996, row.C);
        Assert.Equal(1.2, row.D);
        Assert.Equal(new DateTime(2021, 7, 5), row.E);
        Assert.Equal(new DateTime(2021, 7, 5, 15, 2, 46), row.F);
    }

    [Fact]
    public void TestIssue268_DateFormat()
    {
        Assert.True(IsDateFormatString("dd/mm/yyyy"));
        Assert.True(IsDateFormatString("dd-mmm-yy"));
        Assert.True(IsDateFormatString("dd-mmmm"));
        Assert.True(IsDateFormatString("mmm-yy"));
        Assert.True(IsDateFormatString("h:mm AM/PM"));
        Assert.True(IsDateFormatString("h:mm:ss AM/PM"));
        Assert.True(IsDateFormatString("hh:mm"));
        Assert.True(IsDateFormatString("hh:mm:ss"));
        Assert.True(IsDateFormatString("dd/mm/yyyy hh:mm"));
        Assert.True(IsDateFormatString("mm:ss"));
        Assert.True(IsDateFormatString("mm:ss.0"));
        Assert.True(IsDateFormatString("[$-809]dd mmmm yyyy"));
        Assert.False(IsDateFormatString("#,##0;[Red]-#,##0"));
        Assert.False(IsDateFormatString(@"#,##0.000_);[Red]\(#,##0.000\)"));
        Assert.False(IsDateFormatString("0_);[Red](0)"));
        Assert.False(IsDateFormatString(@"0\h"));
        Assert.False(IsDateFormatString("0\"h\""));
        Assert.False(IsDateFormatString("0%"));
        Assert.False(IsDateFormatString("General"));
        Assert.False(IsDateFormatString(@"_-* #,##0\ _P_t_s_-;\-* #,##0\ _P_t_s_-;_-* "" - ""??\ _P_t_s_-;_-@_- "));
    }

    /// <summary>
    /// Custom excel zip can't read and show Number of entries expected in End Of Central Directory does not correspond to number of entries in Central Directory. #272
    /// </summary>
    [Fact]
    public void TestIssue272()
    {
        var path = PathHelper.GetFile("/xlsx/TestIssue272.xlsx");
        Assert.Throws<InvalidDataException>(() => _excelImporter.Query(path).ToList());
    }

    // Support column width attribute - https://github.com/mini-software/MiniExcel/issues/280
    [Fact]
    public void TestIssue280()
    {
        TestIssue280Dto[] value =
        [
            new() { ID = 1, Name = "Jack" },
            new() { ID = 2, Name = "Mike" }
        ];
        using var path = AutoDeletingPath.Create();
        _excelExporter.Export(path.ToString(), value);
    }

    // Create Multiple Sheets from IDataReader is bugged
    [Fact]
    public void TestIssue283()
    {
        using var path = AutoDeletingPath.Create();
        using (var cn = Db.GetConnection())
        {
            var sheets = new Dictionary<string, object>
            {
                { "sheet01", cn.ExecuteReader("select 'v1' col1") },
                { "sheet02", cn.ExecuteReader("select 'v2' col1") }
            };
            var rows = _excelExporter.Export(path.ToString(), sheets);
            Assert.Equal(2, rows.Length);
        }

        var sheetNames = _excelImporter.GetSheetNames(path.ToString());
        Assert.Equal(["sheet01", "sheet02"], sheetNames);
    }

    [Fact]
    public void TestIssue286()
    {
        TestIssue286Dto[] values =
        [
            new() { E = TestIssue286Enum.VIP1 },
            new() { E = TestIssue286Enum.VIP2 }
        ];
        using var path = AutoDeletingPath.Create();
        _excelExporter.Export(path.ToString(), values);
        var rows = _excelImporter.Query(path.ToString(), true).ToList();

        Assert.Equal("VIP1", rows[0].E);
        Assert.Equal("VIP2", rows[1].E);
    }

    [Fact]
    public void TestIssue289()
    {
        using var path = AutoDeletingPath.Create();
        DescriptionEnumDto[] value =
        [
            new() { Name = "0001", UserType = DescriptionEnum.V1 },
            new() { Name = "0002", UserType = DescriptionEnum.V2 },
            new() { Name = "0003", UserType = DescriptionEnum.V3 }
        ];
        _excelExporter.Export(path.ToString(), value);

        var rows = _excelImporter.Query<DescriptionEnumDto>(path.ToString()).ToList();

        Assert.Equal(DescriptionEnum.V1, rows[0].UserType);
        Assert.Equal(DescriptionEnum.V2, rows[1].UserType);
        Assert.Equal(DescriptionEnum.V3, rows[2].UserType);
    }

    /// Prefix and suffix blank space are lost after SaveAs - https://github.com/mini-software/MiniExcel/issues/294
    [Fact]
    public void TestIssue294()
    {
        {
            using var path = AutoDeletingPath.Create();
            var value = new[] { new { Name = "   Jack" } };
            _excelExporter.Export(path.ToString(), value);
            var sheetXml = SheetHelper.GetZipFileContent(path.ToString(), "xl/worksheets/sheet1.xml");
            Assert.Contains("xml:space=\"preserve\"", sheetXml);
        }
        {
            using var path = AutoDeletingPath.Create();
            var value = new[] { new { Name = "Ja ck" } };
            _excelExporter.Export(path.ToString(), value);
            var sheetXml = SheetHelper.GetZipFileContent(path.ToString(), "xl/worksheets/sheet1.xml");
            Assert.DoesNotContain("xml:space=\"preserve\"", sheetXml);
        }
        {
            using var path = AutoDeletingPath.Create();
            var value = new[] { new { Name = "Jack   " } };
            _excelExporter.Export(path.ToString(), value);
            var sheetXml = SheetHelper.GetZipFileContent(path.ToString(), "xl/worksheets/sheet1.xml");
            Assert.Contains("xml:space=\"preserve\"", sheetXml);
        }
    }

    // SaveAs support Image type - https://github.com/mini-software/MiniExcel/issues/304
    [Fact]
    public void TestIssue304()
    {
        var path = PathHelper.GetTempFilePath();
        var value = new[]
        {
            new { Name="github", Image=File.ReadAllBytes(PathHelper.GetFile("images/github_logo.png"))},
            new { Name="google", Image=File.ReadAllBytes(PathHelper.GetFile("images/google_logo.png"))},
            new { Name="microsoft", Image=File.ReadAllBytes(PathHelper.GetFile("images/microsoft_logo.png"))},
            new { Name="reddit", Image=File.ReadAllBytes(PathHelper.GetFile("images/reddit_logo.png"))},
            new { Name="stackoverflow", Image=File.ReadAllBytes(PathHelper.GetFile("images/stackoverflow_logo.png"))},
        };
        _excelExporter.Export(path, value);

        Assert.Contains("/xl/media/", SheetHelper.GetZipFileContent(path, "xl/drawings/_rels/drawing1.xml.rels"));
        Assert.Contains("ext cx=\"609600\" cy=\"190500\"", SheetHelper.GetZipFileContent(path, "xl/drawings/drawing1.xml"));
        Assert.Contains("/xl/drawings/drawing1.xml", SheetHelper.GetZipFileContent(path, "[Content_Types].xml"));
        Assert.Contains("drawing r:id=", SheetHelper.GetZipFileContent(path, "xl/worksheets/sheet1.xml"));
        Assert.Contains("../drawings/drawing1.xml", SheetHelper.GetZipFileContent(path, "xl/worksheets/_rels/sheet1.xml.rels"));
    }

    // https://github.com/mini-software/MiniExcel/issues/305
    [Fact]
    public async Task TestIssue305()
    {
        var dt = new DateTime(2022, 01, 22);

        using var path = AutoDeletingPath.Create();
        TestIssueI49RZHDto[] value =
        [
            new() { dd = dt },
            new() { dd = null }
        ];
        await _excelExporter.ExportAsync(path.FilePath, value, overwriteFile: true);

        using var package = new ExcelPackage(path.ToString());
        var cells = package.Workbook.Worksheets[0].Cells;
        
        Assert.Equal(dt, DateTime.FromOADate((double)cells["A2"].Value));
        Assert.Equal("22-01-2022", cells["A2"].Text);
    }

    [Fact]
    public async Task TestIssue307()
    {
        using var file = AutoDeletingPath.Create();
        var path = file.ToString();
        var value = new[] { new { id = 1, name = "Jack" } };

        await _excelExporter.ExportAsync(path, value);
        Assert.Throws<IOException>(() => _excelExporter.Export(path, value));

        await _excelExporter.ExportAsync(path, value, overwriteFile: true);
        await Assert.ThrowsAsync<IOException>(async () => await _excelExporter.ExportAsync(path, value));
        await _excelExporter.ExportAsync(path, value, overwriteFile: true);
    }

    [Fact]
    public void TestIssue310()
    {
        using var path = AutoDeletingPath.Create();
        var value = new[] { new TestIssue310Dto { V1 = null }, new TestIssue310Dto { V1 = 2 } };
        _excelExporter.Export(path.ToString(), value);
        var rows = _excelImporter.Query<TestIssue310Dto>(path.ToString()).ToList();
    }

    [Fact]
    public void TestIssue310_Fix497()
    {
        using var path = AutoDeletingPath.Create();
        var value = new[]
        {
            new TestIssue310Dto { V1 = null },
            new TestIssue310Dto { V1 = 2 }
        };
        _excelExporter.Export(path.ToString(), value, configuration: new OpenXmlConfiguration { EnableWriteNullValueCell = false });
        var rows = _excelImporter.Query<TestIssue310Dto>(path.ToString()).ToList();
    }

    [Fact]
    public void TestIssue312()
    {
        using var path = AutoDeletingPath.Create();
        TestIssue312Dto[] value =
        [
            new() { Value = 12_345.6789 },
            new() { Value = null }
        ];
       _excelExporter.Export(path.ToString(), value);

       using var package = new ExcelPackage(path.ToString());
       var cells = package.Workbook.Worksheets[0].Cells;

       var fmt = cells["A2"].Style.Numberformat.Format;
       Assert.Equal(12_345.68.ToString(fmt), cells["A2"].Text);
       Assert.Equal(12_345.6789, (double)cells["A2"].Value);
    }

    // SaveAs and Query support btye[] base64 converter - https://github.com/mini-software/MiniExcel/issues/318
    [Fact]
    public void TestIssue318()
    {
        var imageByte = File.ReadAllBytes(PathHelper.GetFile("images/github_logo.png"));
        using var path = AutoDeletingPath.Create();
        var value = new[]
        {
            new { Name="github", Image=imageByte},
        };
        _excelExporter.Export(path.ToString(), value);


        // import to byte[]
        {
            const string expectedBase64 = "iVBORw0KGgoAAAANSUhEUgAAABwAAAAcCAIAAAD9b0jDAAAACXBIWXMAAA7EAAAOxAGVKw4bAAAAEXRFWHRTb2Z0d2FyZQBTbmlwYXN0ZV0Xzt0AAALNSURBVEiJ7ZVLTBNBGMdndrfdIofy0ERbCgcFeYRuCy2JGOPNRA9qeIZS6YEEogQj0YMmGOqDSATxQaLRxKtRID4SgjGelUBpaQvGZ7kpII8aWtjd2dkdDxsJoS1pIh6M/k+z8833m/3+8+0OJISArRa15cT/0D8CZTYPe32+Zy+GxjzjMzOzAACDYafdZquqOG7hzJtkwUQthRC6cavv0eN+QRTBujUQQp1OV1dbffZMq1arTRaqKIok4eZTrSNjHqIo6gIIIQBgbQwpal+Z/f7dPo2GoaiNHtJut3vjPhBe7+kdfvW61Mq1nGyaX1xYjkRzsk2Z6Rm8IOTvzWs73SLwwqjHK4jCgf3lcV6VxGgiECji7AXm0gvtHYQQnue/zy8ghCRJWlxaWuV5Qsilq9cKzLYiiz04ORVLiHP6A4NPRQlhjLWsVpZlnU63Y3umRqNhGCYjPV3HsrIsMwyDsYQQejIwGEuIA/WMT1AAaDSahnoHTdPKL1vXPKVp2umoZVkWAOj1+ZOCzs7NKYTo9XqjYRcAgKIo9ZRUu9VxltGYZTQAAL5+m0kKijEmAPCrqyJCcRuOECKI4lL4ByEEYykpaE62iQIgurLi9wchhLIsry8fYwwh9PomwuEwACDbZEoKauHMgKJSU1PbOy6Hpqdpml5fPsMwn7+EOru6IYQAghKrJSloTVUFURSX02G3lRw+WulqbA4EJ9XQh4+f2s6dr65zhkLTEEIKwtqaylhCnG/fauFO1Nfde/Bw6Hm/0WiYevc+LU2vhlK2pQwNvwQAsCwrYexyOrji4lhCnOaXZRljXONoOHTk2Ju3I/5AcC3EC0JZ+cE9Bea8IqursUkUker4BsWBqpIk6aL7Sm4htzvfvByJqJORaDS3kMsvLuns6kYIJcpNCFU17pvouXlHEET1URDEnt7bo2OezbMS/vp+R3/PdfKPQ38Ccg0E/CDcpY8AAAAASUVORK5CYII=";
            var rows = _excelImporter.Query(path.ToString(), true).ToList();
            var actulBase64 = Convert.ToBase64String((byte[])rows[0].Image);
            Assert.Equal(expectedBase64, actulBase64);
        }

        // import to base64 string
        {
            var config = new OpenXmlConfiguration { EnableConvertByteArray = false };
            var rows = _excelImporter.Query(path.ToString(), true, configuration: config).ToList();
            var image = (string)rows[0].Image;
            Assert.StartsWith("@@@fileid@@@,xl/media/", image);
        }

    }
    
    // https://github.com/mini-software/MiniExcel/issues/325
    [Fact]
    public void TestIssue325()
    {
        using var path = AutoDeletingPath.Create();
        var value = new Dictionary<string, object>
        {
            { "sheet1",new[]{ new { id = 1, date = DateTime.Parse("2022-01-01") } }},
            { "sheet2",new[]{ new { id = 2, date = DateTime.Parse("2022-01-01") } }},
        };
        _excelExporter.Export(path.ToString(), value);

        var xml = SheetHelper.GetZipFileContent(path.ToString(), "xl/worksheets/_rels/sheet2.xml.rels");
        var cnt = Regex.Matches(xml, "Id=\"drawing2\"").Count;
        Assert.True(cnt == 1);
    }

    [Fact]
    public void TestIssue327()
    {
        using var path = AutoDeletingPath.Create();
        var value = new[]
        {
            new { id = 1, file = File.ReadAllBytes(PathHelper.GetFile("xlsx/Issue327/TestIssue327.png")) },
            new { id = 2, file = File.ReadAllBytes(PathHelper.GetFile("xlsx/Issue327/TestIssue327.txt")) },
            new { id = 3, file = File.ReadAllBytes(PathHelper.GetFile("xlsx/Issue327/TestIssue327.html")) },
        };
        _excelExporter.Export(path.ToString(), value);
        var rows = _excelImporter.Query(path.ToString(), true).ToList();

        Assert.Equal(value[0].file, rows[0].file);
        Assert.Equal(value[1].file, rows[1].file);
        Assert.Equal(value[2].file, rows[2].file);
        Assert.Equal("Hello MiniExcel", Encoding.UTF8.GetString(rows[1].file));
        Assert.Equal("<html>Hello MiniExcel</html>", Encoding.UTF8.GetString(rows[2].file));
    }

    [Fact]
    public void TestIssue328()
    {
        using var path = AutoDeletingPath.Create();
        var value = new[]
        {
            new
            {
                id=1,
                name="Jack",
                indate=new DateTime(2022,5,13),
                file = File.ReadAllBytes(PathHelper.GetFile("xlsx/Issue327/TestIssue327.png"))
            },
            new
            {
                id=2,
                name="Henry",
                indate=new DateTime(2022,4,10),
                file = File.ReadAllBytes(PathHelper.GetFile("xlsx/Issue327/TestIssue327.txt"))
            },
        };
        _excelExporter.Export(path.ToString(), value);

        var rowIndx = 0;
        using var reader = _excelImporter.GetDataReader(path.ToString(), true);

        Assert.Equal("id", reader.GetName(0));
        Assert.Equal("name", reader.GetName(1));
        Assert.Equal("indate", reader.GetName(2));
        Assert.Equal("file", reader.GetName(3));

        while (reader.Read())
        {
            for (int i = 0; i < reader.FieldCount; i++)
            {
                var v = reader.GetValue(i);
                if (rowIndx == 0 && i == 0) Assert.Equal(1.0, v);
                if (rowIndx == 0 && i == 1) Assert.Equal("Jack", v);
                if (rowIndx == 0 && i == 2) Assert.Equal(new DateTime(2022, 5, 13), v);
                if (rowIndx == 0 && i == 3) Assert.Equal(File.ReadAllBytes(PathHelper.GetFile("xlsx/Issue327/TestIssue327.png")), v);
                if (rowIndx == 1 && i == 0) Assert.Equal(2.0, v);
                if (rowIndx == 1 && i == 1) Assert.Equal("Henry", v);
                if (rowIndx == 1 && i == 2) Assert.Equal(new DateTime(2022, 4, 10), v);
                if (rowIndx == 1 && i == 3) Assert.Equal(File.ReadAllBytes(PathHelper.GetFile("xlsx/Issue327/TestIssue327.txt")), v);
            }
            rowIndx++;
        }

        //TODO:How to resolve empty body sheet?
    }

    [Fact]
    public void TestIssue331()
    {
        var cln = CultureInfo.CurrentCulture.Name;
        CultureInfo.CurrentCulture = CultureInfo.GetCultureInfo("cs-CZ");

        var data = Enumerable.Range(1, 10).Select(x => new TestIssue331Dto
        {
            Number = x,
            Text = $"Number {x}",
            DecimalNumber = x / 2m,
            DoubleNumber = x / 2d
        });

        using var path = AutoDeletingPath.Create();
        _excelExporter.Export(path.ToString(), data);

        var rows = _excelImporter.Query(path.ToString(), startCell: "A2").ToArray();
        Assert.Equal(1.5, rows[2].B);
        Assert.Equal(1.5, rows[2].C);

        CultureInfo.CurrentCulture = CultureInfo.GetCultureInfo(cln);
    }

    [Fact]
    public void TestIssue331_2()
    {
        var cln = CultureInfo.CurrentCulture.Name;
        CultureInfo.CurrentCulture = CultureInfo.GetCultureInfo("cs-CZ");

        var config = new OpenXmlConfiguration
        {
            Culture = CultureInfo.GetCultureInfo("cs-CZ")
        };

        var rnd = new Random();
        var data = Enumerable.Range(1, 100).Select(x => new TestIssue331Dto
        {
            Number = x,
            Text = $"Number {x}",
            DecimalNumber = (decimal)rnd.NextDouble(),
            DoubleNumber = rnd.NextDouble()
        });

        using var path = AutoDeletingPath.Create();
        _excelExporter.Export(path.ToString(), data, configuration: config);
        CultureInfo.CurrentCulture = CultureInfo.GetCultureInfo(cln);
    }

    // Excel was unable to open the file https://github.com/mini-software/MiniExcel/issues/343
    [Fact]
    public void TestIssue343()
    {
        var date = DateTime.Parse("2022-03-17 09:32:06.111", CultureInfo.InvariantCulture);
        using var path = AutoDeletingPath.Create();

        CultureInfo.CurrentCulture = CultureInfo.GetCultureInfo("ff-Latn");
        var table = new DataTable();
        {
            table.Columns.Add("time1", typeof(DateTime));
            table.Columns.Add("time2", typeof(DateTime));
            table.Rows.Add(date, date);
            table.Rows.Add(date, date);
        }
        var reader = table.CreateDataReader();
        _excelExporter.Export(path.ToString(), reader);

        var rows = _excelImporter.Query(path.ToString(), true).ToArray();
        Assert.Equal(date, rows[0].time1);
        Assert.Equal(date, rows[0].time2);
    }

    [Fact]
    public void TestIssue352()
    {
        {
            using var table = new DataTable();
            table.Columns.Add("id", typeof(int));
            table.Columns.Add("name", typeof(string));
            table.Rows.Add(1, "Jack");
            table.Rows.Add(2, "Mike");

            using var path = AutoDeletingPath.Create();
            var reader = table.CreateDataReader();
            _excelExporter.Export(path.ToString(), reader);
            var xml = SheetHelper.GetZipFileContent(path.ToString(), "xl/worksheets/sheet1.xml");
            var cnt = Regex.Count(xml, "<x:autoFilter ref=\"A1:B3\" />");
        }
        {
            using var table = new DataTable();
            table.Columns.Add("id", typeof(int));
            table.Columns.Add("name", typeof(string));
            table.Rows.Add(1, "Jack");
            table.Rows.Add(2, "Mike");

            using var path = AutoDeletingPath.Create();
            var reader = table.CreateDataReader();
            _excelExporter.Export(path.ToString(), reader, false);
            var xml = SheetHelper.GetZipFileContent(path.ToString(), "xl/worksheets/sheet1.xml");
            var cnt = Regex.Count(xml, "<x:autoFilter ref=\"A1:B2\" />");
        }
        {
            using var table = new DataTable();
            table.Columns.Add("id", typeof(int));
            table.Columns.Add("name", typeof(string));

            using var path = AutoDeletingPath.Create();
            var reader = table.CreateDataReader();
            _excelExporter.Export(path.ToString(), reader);
            var xml = SheetHelper.GetZipFileContent(path.ToString(), "xl/worksheets/sheet1.xml");
            var cnt = Regex.Count(xml, "<x:autoFilter ref=\"A1:B1\" />");
        }
    }

    [Fact]
    public void TestIssue360()
    {
        var path = PathHelper.GetFile("xlsx/NotDuplicateSharedStrings_10x100.xlsx");
        var config = new OpenXmlConfiguration { SharedStringCacheSize = 1 };
        var sheets = _excelImporter.GetSheetNames(path);
        foreach (var sheetName in sheets)
        {
            _ = _excelImporter.QueryAsDataTable(path, hasHeaderRow: true, sheetName: sheetName, configuration: config);
        }
    }

    [Fact]
    public void TestIssue369()
    {
        var config = new OpenXmlConfiguration
        {
            DynamicColumns =
            [
                new DynamicExcelColumn("id") { Ignore=true },
                new DynamicExcelColumn("name") { Index=1, Width=10 },
                new DynamicExcelColumn("createdate") { Index=0, Format="yyyy-MM-dd", Width=15 },
                new DynamicExcelColumn("point") { Index=2, Name="Account Point" }
            ]
        };
        using var path = AutoDeletingPath.Create();
        var value = new[] { new { id = 1, name = "Jack", createdate = new DateTime(2022, 04, 12), point = 123.456 } };
        _excelExporter.Export(path.ToString(), value, configuration: config);

        var rows = _excelImporter.Query(path.ToString()).ToList();
        Assert.Equal("createdate", rows[0].A);
        Assert.Equal(new DateTime(2022, 04, 12), rows[1].A);
        Assert.Equal("name", rows[0].B);
        Assert.Equal("Jack", rows[1].B);
        Assert.Equal("Account Point", rows[0].C);
        Assert.Equal(123.456, rows[1].C);
    }

    [Fact]
    public void TestIssue370()
    {
        var config = new OpenXmlConfiguration
        {
            DynamicColumns =
            [
                new DynamicExcelColumn("Id") { Ignore = true },
                new DynamicExcelColumn("Name") { Index = 1,Width = 10 },
                new DynamicExcelColumn("Date") { Index = 0, Format="yyyy-MM-dd", Width = 15 },
                new DynamicExcelColumn("Point") { Index = 2, Name = "Account Point" }
            ]
        };
        using var path = AutoDeletingPath.Create();
        List<Dictionary<string, object>> value =
        [
            new()
            {
                ["Id"] = 1,
                ["Name"] = "Jack",
                ["Date"] = new DateTime(2022, 04, 12),
                ["Point"] = 123.456
            }
        ];
        _excelExporter.Export(path.ToString(), value, configuration: config);

        var rows = _excelImporter.Query(path.ToString()).ToList();
        Assert.Equal("Date", rows[0].A);
        Assert.Equal(new DateTime(2022, 04, 12), rows[1].A);
        Assert.Equal("Name", rows[0].B);
        Assert.Equal("Jack", rows[1].B);
        Assert.Equal("Account Point", rows[0].C);
        Assert.Equal(123.456, rows[1].C);    
    }

    [Theory]
    [InlineData(true, 1)]
    [InlineData(false, 0)]
    public void TestIssue401(bool autoFilter, int count)
    {
        // Test for DataTable
        {
            var table = new DataTable();
            {
                table.Columns.Add("id", typeof(int));
                table.Columns.Add("name", typeof(string));
                table.Rows.Add(1, "Jack");
                table.Rows.Add(2, "Mike");
            }

            var reader = table.CreateDataReader();
            using var path = AutoDeletingPath.Create();
            var config = new OpenXmlConfiguration { AutoFilter = autoFilter };
            _excelExporter.Export(path.ToString(), reader, configuration: config);

            var xml = SheetHelper.GetZipFileContent(path.ToString(), "xl/worksheets/sheet1.xml");
            var cnt = Regex.Matches(xml, "<x:autoFilter ref=\"A1:B3\" />").Count;
            Assert.Equal(count, cnt);
        }
        {
            var table = new DataTable();
            {
                table.Columns.Add("id", typeof(int));
                table.Columns.Add("name", typeof(string));
                table.Rows.Add(1, "Jack");
                table.Rows.Add(2, "Mike");
            }
            var reader = table.CreateDataReader();
            using var path = AutoDeletingPath.Create();
            var config = new OpenXmlConfiguration { AutoFilter = autoFilter };
            _excelExporter.Export(path.ToString(), reader, false, configuration: config);

            var xml = SheetHelper.GetZipFileContent(path.ToString(), "xl/worksheets/sheet1.xml");
            var cnt = Regex.Matches(xml, "<x:autoFilter ref=\"A1:B2\" />").Count;
            Assert.Equal(count, cnt);
        }
        {
            var table = new DataTable();
            {
                table.Columns.Add("id", typeof(int));
                table.Columns.Add("name", typeof(string));
            }
            var reader = table.CreateDataReader();
            using var path = AutoDeletingPath.Create();
            var config = new OpenXmlConfiguration { AutoFilter = autoFilter };
            _excelExporter.Export(path.ToString(), reader, configuration: config);

            var xml = SheetHelper.GetZipFileContent(path.ToString(), "xl/worksheets/sheet1.xml");
            var cnt = Regex.Matches(xml, "<x:autoFilter ref=\"A1:B1\" />").Count;
            Assert.Equal(count, cnt);
        }

        // Test for DataReader
        {
            using var path = AutoDeletingPath.Create();
            var config = new OpenXmlConfiguration { AutoFilter = autoFilter };
            using (var connection = Db.GetConnection("Data Source=:memory:"))
            {
                connection.Open();

                using var command = connection.CreateCommand();
                command.CommandText =
                    """
                    SELECT
                        'MiniExcel' as Column1,
                        1 as Column2 

                    UNION ALL 
                    SELECT 'Github', 2
                    """;

                using var reader = command.ExecuteReader();
                _excelExporter.Export(path.ToString(), reader, configuration: config);
            }

            var xml = SheetHelper.GetZipFileContent(path.ToString(), "xl/worksheets/sheet1.xml");
            var cnt = Regex.Matches(xml, "autoFilter").Count;
            Assert.Equal(count, cnt);
        }

        {
            var xlsxPath = PathHelper.GetFile("xlsx/Test5x2.xlsx");
            using var tempSqlitePath = AutoDeletingPath.Create(Path.GetTempPath(), $"{Guid.NewGuid()}.db");
            var connectionString = $"Data Source={tempSqlitePath};Version=3;";

            using (var connection = new SQLiteConnection(connectionString))
            {
                connection.Execute("create table T (A varchar(20),B varchar(20));");
            }

            using (var connection = new SQLiteConnection(connectionString))
            {
                connection.Open();
                using (var transaction = connection.BeginTransaction())
                using (var stream = File.OpenRead(xlsxPath))
                {
                    var rows = _excelImporter.Query(stream);
                    foreach (var row in rows)
                        connection.Execute(
                            "insert into T (A,B) values (@A,@B)",
                            new { row.A, row.B },
                            transaction: transaction);

                    transaction.Commit();
                }
            }

            using var path = AutoDeletingPath.Create();
            var config = new OpenXmlConfiguration { AutoFilter = autoFilter };
            using (var connection = new SQLiteConnection(connectionString))
            {
                using var command = new SQLiteCommand("select * from T", connection);
                connection.Open();
                using var reader = command.ExecuteReader();
                _excelExporter.Export(path.ToString(), reader, configuration: config);
            }

            var xml = SheetHelper.GetZipFileContent(path.ToString(), "xl/worksheets/sheet1.xml");
            var cnt = Regex.Matches(xml, "autoFilter").Count;
            Assert.Equal(count, cnt);
        }
    }

    // https://github.com/MiniExcel/MiniExcel/issues/405)
    [Fact]
    public void TestIssue405()
    {
        using var path = AutoDeletingPath.Create();
        var value = new[] { new { id = 1, name = "test" } };
        _excelExporter.Export(path.ToString(), value);

        var xml = SheetHelper.GetZipFileContent(path.ToString(), "xl/sharedStrings.xml");
        Assert.StartsWith("<x:sst xmlns:x=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"", xml);
    }

    [Fact]
    public void TestIssue408()
    {
        using var reader = _excelImporter.GetDataReader(PathHelper.GetFile("xlsx/TestTypeMapping.xlsx"), true);
        
        var dt = new DataTable();
        dt.Load(reader);
        
        Assert.Equal(100, dt.Rows.Count);
        Assert.Equal("78DE23D2-DCB6-BD3D-EC67-C112BBC322A2", dt.Rows[0]["ID"]);
        Assert.Equal("Wade", dt.Rows[0]["Name"]);
        Assert.Equal("27/09/2020", dt.Rows[0]["BoD"]);
        Assert.Equal(36d, dt.Rows[0]["Age"]);
        Assert.False(Convert.ToBoolean(dt.Rows[0]["VIP"]));
        Assert.Equal(5019.12, dt.Rows[0]["Points"]);
    }
    
    [Fact]
    public void TestIssue409()
    {
        var path = PathHelper.GetFile("xlsx/TestIssue409.xlsx");
        var config = new OpenXmlConfiguration 
        {
            Culture = new CultureInfo("ru")
            {
                NumberFormat = { NumberDecimalSeparator = "," }
            }
        };

        var query = _excelImporter.Query<Issue409Dto>(path, configuration: config).ToList();

        Assert.Equal(0.002886, query[0].Quantity);
        Assert.Equal(4.1E-05, query[1].Quantity);
        Assert.Equal(0.02586, query[2].Quantity);
        Assert.Equal(0.000217, query[3].Quantity);
        Assert.Equal(17.4024812, query[4].Quantity);
        Assert.Equal(1.43E-06, query[5].Quantity);
        Assert.Equal(9.9E-06, query[6].Quantity);
    }

    // https://github.com/MiniExcel/MiniExcel/issues/413)
    [Fact]
    public void TestIssue413()
    {
        using var path = AutoDeletingPath.Create();
        var value = new
        {
            list = new List<Dictionary<string, object>>
            {
                new() { { "id","001"},{ "time",new DateTime(2022,12,25)} },
                new() { { "id","002"},{ "time",new DateTime(2022,9,23)} },
            }
        };
        var templatePath = PathHelper.GetFile("xlsx/TestIssue413.xlsx");
         _excelTemplater.FillTemplate(path.ToString(), templatePath, value);
        var rows = _excelImporter.Query(path.ToString()).ToList();

        Assert.Equal("2022-12-25 00:00:00", rows[1].B);
        Assert.Equal("2022-09-23 00:00:00", rows[2].B);
    }

    [Fact]
    public void Issue422()
    {
        var items = new[]
        {
            new { Row1 = "1", Row2 = "2" },
            new { Row1 = "3", Row2 = "4" }
        };
        var enumerableWithCount = new Issue422Enumerable(items);

        using var path = AutoDeletingPath.Create();
        _excelExporter.Export(path.ToString(), enumerableWithCount);

        Assert.Equal(1, enumerableWithCount.GetEnumeratorCount);
    }

    // Exception : MiniExcelLibs.Core.Exceptions.ExcelInvalidCastException: 'ColumnName : Date, CellRow : 2, Value : 2021-01-31 10:03:00 +08:00, it can't cast to DateTimeOffset type.'
    [Fact]
    public void TestIssue430()
    {
        using var path = AutoDeletingPath.Create();
        TestIssue430Dto[] value =
        [
            new() { Date = new DateTimeOffset(2021, 1, 31, 10, 3, 0, TimeSpan.FromHours(5)) }
        ];
        _excelExporter.Export(path.ToString(), value);

        var testValue = _excelImporter.Query(path.ToString(), hasHeaderRow: true).First();
        Assert.Equal("2021-01-31 10:03:00", testValue.Date.ToString("yyyy-MM-dd HH:mm:ss"));
    }

    [Fact]
    public void Issue459()
    {
        var template = PathHelper.GetFile("xlsx/TestIssue459.xlsx");
        using var ms = new MemoryStream();
        var values = new
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

         _excelTemplater.FillTemplate(ms, template, values);
    }

    [Fact]
    public void Issue520()
    {
        using var ms = new MemoryStream();
        Issue520Dto[] data = [new(542, DateTime.Today, 300)];
        _excelExporter.Export(ms, data);
        ms.Seek(0, SeekOrigin.Begin);
        
        using var package = new ExcelPackage(ms);
        var cells = package.Workbook.Worksheets.First().Cells;
        
        Assert.Equal(542.0, cells["A2"].Value);
        Assert.Equal(542.0.ToString("R$ #,##0.00"), cells["A2"].Text);
        Assert.Equal(DateTime.Today, DateTime.FromOADate((double)cells["B2"].Value));
        Assert.Equal(DateTime.Today.ToString("dd/MM/yyyy"), cells["B2"].Text);
        Assert.Equal(300.0, cells["C2"].Value);
        Assert.Equal(300.0.ToString("R$ #,##0.00"), cells["C2"].Text);
    }
    
    [Fact]
    public void Issue527()
    {
        List<DescriptionEnumDto> row =
        [
            new() { Name = "Bill", UserType = DescriptionEnum.V1 },
            new() { Name = "Bob", UserType = DescriptionEnum.V2 }
        ];

        var value = new { t = row };
        var template = PathHelper.GetFile("xlsx/Issue527Template.xlsx");

        using var path = AutoDeletingPath.Create();
        _excelTemplater.FillTemplate(path.FilePath, template, value);

        var rows = _excelImporter.Query(path.FilePath).ToList();
        Assert.Equal("General User", rows[1].B);
        Assert.Equal("General Administrator", rows[2].B);
    }
    
    [Fact]
    public void Issue542()
    {
        var path = PathHelper.GetFile("xlsx/TestIssue542.xlsx");

        var resultWithoutFirstRow = _excelImporter.Query<Issue542>(path).ToList();
        var resultWithFirstRow = _excelImporter.Query<Issue542>(path, treatHeaderAsData: true).ToList();

        Assert.Equal(15, resultWithoutFirstRow.Count);
        Assert.Equal(16, resultWithFirstRow.Count);

        Assert.Equal("Felix", resultWithoutFirstRow[0].Name);
        Assert.Equal("Wade", resultWithFirstRow[0].Name);
    }

    [Fact]
    public void TestIssue549()
    {
        var data = new[]
        {
            new{ id = 1, name = "jack" },
            new{ id = 2, name = "mike" }
        };
        
        using var file = AutoDeletingPath.Create();
        var path = file.ToString();

        _excelExporter.Export(path, data);
        var rows = _excelImporter.Query(path, true).ToList();
        {
            using var stream = new FileStream(path, FileMode.Open, FileAccess.Read);
            using var workbook = new XSSFWorkbook(stream);

            var sheet = workbook.GetSheetAt(0);
            var a2 = sheet.GetRow(1).GetCell(0);
            var b2 = sheet.GetRow(1).GetCell(1);
            Assert.Equal((string)rows[0].id.ToString(), a2.NumericCellValue.ToString(CultureInfo.InvariantCulture));
            Assert.Equal((string)rows[0].name.ToString(), b2.StringCellValue);
        }
    }

    [Fact]
    public void TestIssue553()
    {
        using var path = AutoDeletingPath.Create();
        var templatePath = PathHelper.GetFile("xlsx/TestIssue553.xlsx");
        var data = new
        {
            B = new[]
            {
                new{ ITM = 1 },
                new{ ITM = 2 },
                new{ ITM = 3 }
            }
        };
        _excelTemplater.FillTemplate(path.ToString(), templatePath, data);

        var rows = _excelImporter.Query(path.ToString()).ToList();
        Assert.Equal(rows[2].A, 1);
        Assert.Equal(rows[3].A, 2);
        Assert.Equal(rows[4].A, 3);
    }

    [Fact]
    public void TestIssue584()
    {
        var excelconfig = new OpenXmlConfiguration
        {
            FastMode = true,
            DynamicColumns =
            [
                new DynamicExcelColumn("Id") { Ignore = true }
            ]
        };

        using var conn = Db.GetConnection();
        conn.Open();

        using var cmd = conn.CreateCommand();
        cmd.CommandText =
            """
            WITH test('Id', 'Name') AS (
                VALUES 
                    (1, 'test1'), 
                    (2, 'test2'), 
                    (3, 'test3')
                )
            SELECT * FROM test;
            """;
        using var reader = cmd.ExecuteReader();

        using var path = AutoDeletingPath.Create();
        _excelExporter.Export(path.FilePath, reader, configuration: excelconfig, overwriteFile: true);

        var rows = _excelImporter.Query(path.FilePath).ToList();
        Assert.All(rows, x => Assert.Single(x));
        Assert.Equal("Name", rows[0].A);
    }

    [Fact]
    public void Issue585()
    {
        var path = PathHelper.GetFile("xlsx/TestIssue585.xlsx");

        var items1 = _excelImporter.Query<Issue585Variant1>(path);
        Assert.Equal(2, items1.Count());

        var items2 = _excelImporter.Query<Issue585Variant2>(path);
        Assert.Equal(2, items2.Count());

        var items3 = _excelImporter.Query<Issue585Variant3>(path);
        Assert.Equal(2, items3.Count());
    }

    [Fact]
    public void Issue606_1()
    {
        // excel max rows: 1,048,576
        // before changes: 1999999 => 25.8 GB mem
        //  after changes: 1999999 => peaks at 3.2 GB mem (10:20 min)
        //  after changes:  100000 => peaks at 222 MB mem (34 sec)

        var value = new
        {
            Title = "My Title",
            OrderInfo = Enumerable
                .Range(1, 100)
                .Select(_ => new
                {
                    Standard = "standard",
                    RegionName = "region",
                    DealerName = "department",
                    SalesPointName = "region",
                    CustomerName = "customer",
                    IdentityType = "aaaaaa",
                    IdentitySeries = "ssssss",
                    IdentityNumber = "nnnnn",
                    BirthDate = "date",
                    TariffPlanName = "plan",
                    PhoneNumber = "num",
                    SimCardIcc = "sss",
                    BisContractNumber = "eee",
                    CreatedAt = "dd.mm.yyyy",
                    UserDescription = "fhtyrhthrthrt",
                    UserName = "dfsfsdfds",
                    PaymentsAmount = "dfhgdfgadfgdfg",
                    OrderState = "agafgdafgadfgd",
                })
        };

        var path = Path.Combine
        (
            Path.GetTempPath(),
            string.Concat(nameof(MiniExcelGithubIssuesTests), "_", nameof(Issue606_1), ".xlsx")
        );

        var templateFileName = PathHelper.GetFile("xlsx/TestIssue606_Template.xlsx");
         _excelTemplater.FillTemplate(path, Path.GetFullPath(templateFileName), value);
        File.Delete(path);
    }

    [Fact]
    public void TestIssue627()
    {
        var config = new OpenXmlConfiguration
        {
            AutoFilter = false,
            DynamicColumns =
            [
                new DynamicExcelColumn("long2") { Format = "@", Width = 25 }
            ]
        };

        using var path = AutoDeletingPath.Create();
        var value = new[] { new { long2 = "1550432695793487872" } };
        var rowsWritten =  _excelExporter.Export(path.ToString(), value, configuration: config);

        Assert.Single(rowsWritten);
        Assert.Equal(1, rowsWritten[0]);
    }

    [Fact]
    public void Issue632_1()
    {
        //https://github.com/mini-software/MiniExcel/issues/632
        var values = Enumerable.Range(1, 100)
            .Select(item => new Dictionary<string, object>
            {
                { "Id", item },
                { "Time", DateTime.Now.ToLocalTime() },
                { "CPU Usage (%)", Math.Round(56.345, 1) },
                { "Memory Usage (%)", Math.Round(98.234, 1) },
                { "Disk Usage (%)", Math.Round(32.456, 1) },
                { "CPU Temperature (°C)", Math.Round(74.234, 1) },
                { "Voltage (V)", Math.Round(6.3223, 1) },
                { "Network Usage (Kb/s)", Math.Round(4503.23422, 1) },
                { "Instrument", "QT800050" }
            })
            .ToList();

        var config = new OpenXmlConfiguration
        {
            TableStyles = TableStyles.None,
            DynamicColumns = [new DynamicExcelColumn("Time") { Index = 0, Width = 20, Format = "d.MM.yyyy" }]
        };

        var path = Path.Combine(
            Path.GetTempPath(),
            string.Concat(nameof(MiniExcelGithubIssuesTests), "_", nameof(Issue632_1), ".xlsx")
        );

         _excelExporter.Export(path, values, configuration: config, overwriteFile: true);
        File.Delete(path);
    }

    [Fact]
    public void Issue_658()
    {
        static IEnumerable<Issue658Dto> GetTestData()
        {
            yield return new() { FirstName = "Unit", LastName = "Test" };
            yield return new() { FirstName = "Unit1", LastName = "Test1" };
            yield return new() { FirstName = "Unit2", LastName = "Test2" };
        }

        using var memoryStream = new MemoryStream();
        var testData = GetTestData().ToList();
        var rowsWritten = _excelExporter.Export(memoryStream, testData, configuration: new OpenXmlConfiguration
        {
            FastMode = true
        });
        Assert.Single(rowsWritten);
        Assert.Equal(3, rowsWritten[0]);

        memoryStream.Position = 0;

        var queryData = _excelImporter.Query<Issue658Dto>(memoryStream).ToList();

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
    public void Issue_686()
    {
        var path = PathHelper.GetFile("xlsx/TestIssue686.xlsx");
        Assert.Throws<InvalidDataException>(() =>
            _excelImporter.QueryRange(path, hasHeaderRow: false, startCell: "ZZFF10", endCell: "ZZFF11").First());

        Assert.Throws<InvalidDataException>(() =>
            _excelImporter.QueryRange(path, hasHeaderRow: false, startCell: "ZZFF@@10", endCell: "ZZFF@@11").First());
    }

    [Fact]
    public void Test_Issue_693_SaveSheetWithLongName()
    {
        using var path1 = AutoDeletingPath.Create();
        using var path2 = AutoDeletingPath.Create();

        List<Dictionary<string, object>> data = [new() { ["First"] = 1, ["Second"] = 2 }];
        Assert.Throws<ArgumentException>(() =>  _excelExporter.Export(path1.ToString(), data, sheetName: "Some Really Looooooooooong Sheet Name"));
         _excelExporter.Export(path2.ToString(), new List<Dictionary<string, object>>());
        Assert.Throws<ArgumentException>(() =>  _excelExporter.InsertSheet(path2.ToString(), data, sheetName: "Some Other Very Looooooong Sheet Name"));
    }

    [Fact]
    public void Test_Issue_697_EmptyRowsStronglyTypedQuery()
    {
        var path = PathHelper.GetFile("xlsx/TestIssue697.xlsx");
        var rowsIgnoreEmpty = _excelImporter.Query<Issue697Dto>(path, configuration: new OpenXmlConfiguration { IgnoreEmptyRows = true }).ToList();
        var rowsCountEmpty = _excelImporter.Query<Issue697Dto>(path).ToList();
        Assert.Equal(4, rowsIgnoreEmpty.Count);
        Assert.Equal(5, rowsCountEmpty.Count);
    }

    [Fact]
    public void Issue_710()
    {
        var values = new[] { new { Column1 = "MiniExcel", Column2 = 1, Column3 = "Test" } };
        using var memoryStream = new MemoryStream();
        _excelExporter.Export(memoryStream, values, configuration: new OpenXmlConfiguration
        {
            FastMode = true
        });

        memoryStream.Position = 0;
        using var dataReader = _excelImporter.GetDataReader(memoryStream, hasHeaderRow: false);

        dataReader.Read();
        for (int i = 0; i < dataReader.FieldCount; i++)
        {
            var columnName = dataReader.GetName(i);
            var ordinal = dataReader.GetOrdinal(columnName);

            Assert.Equal(i, ordinal);
        }
    }

    [Fact]
    public void Issue_732_First_Sheet_Active()
    {
        var path1 = PathHelper.GetFile("xlsx/TestIssue732_1.xlsx");
        var path2 = PathHelper.GetFile("xlsx/TestIssue732_2.xlsx");
        var path3 = PathHelper.GetFile("xlsx/TestIssue732_3.xlsx");

        var info1 = _excelImporter.GetSheetInformations(path1);
        var info2 = _excelImporter.GetSheetInformations(path2);
        var info3 = _excelImporter.GetSheetInformations(path3);

        Assert.Equal(0u, info1.SingleOrDefault(x => x.Active)?.Index); // first sheet is active
        Assert.Equal(1u, info2.SingleOrDefault(x => x.Active)?.Index); // second sheet is active
        Assert.Equal(0u, info3.SingleOrDefault(x => x.Active)?.Index); // only one sheet in file
    }

    [Fact]
    public void TestIssue750()
    {
        var templatePath = PathHelper.GetFile("xlsx/TestIssue20250403_SaveAsByTemplate_OPT.xlsx");
        var memoryBefore = GC.GetTotalMemory(true);

        using var path = AutoDeletingPath.Create();
        var data = new Dictionary<string, object>
        {
            ["list"] = Enumerable.Range(0, 10_000)
                .Select(_ => new { value1 = Guid.NewGuid(), value2 = Guid.NewGuid(), })
        };
         _excelTemplater.FillTemplate(path.ToString(), templatePath, data);

        var rows = _excelImporter.Query(path.ToString())
            .Skip(1453)
            .Take(2)
            .ToList();

        Assert.True(((string)rows[0].A).Length > 9);

        var memoryAfter = GC.GetTotalMemory(true);
        var memoryIncrease = memoryAfter - memoryBefore;

        _output.WriteLine($"memoryIncrease: {memoryIncrease}");
    }

    [Fact]
    public void TestIssue751()
    {
        var templatePath = PathHelper.GetFile("xlsx/TestIssue20250403_SaveAsByTemplate_OPT.xlsx");

        using var path = AutoDeletingPath.Create();
        var list = Enumerable.Range(0, 10)
                .Select(_ => new { value1 = Guid.NewGuid(), value2 = Guid.NewGuid(), }).ToList();
        var data = new Dictionary<string, object>
        {
            ["list"] = list
        };
         _excelTemplater.FillTemplate(path.ToString(), templatePath, data);

        using var stream = File.OpenRead(path.ToString());
        using var workbook = new XSSFWorkbook(stream);
        var sheet = workbook.GetSheetAt(0);
        var row = sheet.GetRow(0);
        var cell = row.GetCell(0);
        Assert.Equal("value1", cell.ToString());
        Assert.Equal(0, cell.ColumnIndex);
        Assert.Equal(0, cell.RowIndex);
        Assert.Equal("value2", row.GetCell(1).ToString());
        Assert.Equal(1, row.GetCell(1).ColumnIndex);
        Assert.Equal(0, row.GetCell(1).RowIndex);
        for (int i = 0; i < list.Count; i++)
        {
            var rowData = sheet.GetRow(i + 1);
            Assert.Equal(list[i].value1.ToString(), rowData.GetCell(0).ToString());
            Assert.Equal(list[i].value2.ToString(), rowData.GetCell(1).ToString());
        }
    }

    [Fact]
    public void TestIssue763()
    {
        var path = PathHelper.GetFile("xlsx/TestIssue763.xlsx");
        var rows = _excelImporter.QueryRange(path, startCell: "A3", endCell: "J3").ToArray();
        Assert.Equal("A3", rows[0].A);
        Assert.Null(rows[0].J);
    }

    [Fact]
    public void TestIssue768()
    {
        var templatePath = PathHelper.GetFile("xlsx/TestIssue768.xlsx");
        using var path = AutoDeletingPath.Create();

        var list = Enumerable.Range(0, 10)
            .Select(_ => new
            {
                value1 = Guid.NewGuid(),
                value2 = Guid.NewGuid()
            }
            )
            .ToList();

        var data = new Dictionary<string, object>
        {
            ["list"] = list
        };

         _excelTemplater.FillTemplate(path.ToString(), templatePath, data);
        var rows = _excelImporter.Query(path.ToString(), startCell: "A16").ToList();

        Assert.Equal(list[0].value1.ToString(), rows[0].A.ToString());
        Assert.Equal(list[1].value1.ToString(), rows[1].A.ToString());
    }

    [Fact]
    public void TestIssue771()
    {
        var template = PathHelper.GetFile("xlsx/TestIssue771.xlsx");
        using var path = AutoDeletingPath.Create();

        var value = new
        {
            list = GetEnumerable(),
            list2 = GetEnumerable(),
            list3 = GetEnumerable(),
            list4 = GetEnumerable(),
            list5 = GetEnumerable(),
            list6 = GetEnumerable(),
            list7 = GetEnumerable(),
            list8 = GetEnumerable(),
            list9 = GetEnumerable(),
            list10 = GetEnumerable(),
            list11 = GetEnumerable(),
            list12 = GetEnumerable()
        };

         _excelTemplater.FillTemplate(path.FilePath, template, value);
        var rows = _excelImporter.Query(path.FilePath).ToList();

        Assert.Equal("2025-1", rows[2].B);
        Assert.Null(rows[3].B);
        Assert.Null(rows[4].B);
        Assert.Equal("2025-2", rows[5].B);
        return;

        IEnumerable<object> GetEnumerable() => Enumerable.Range(0, 3).Select(s => new { ID = Guid.NewGuid(), level = s });
    }

    [Fact]
    public void TestIssue772()
    {
        var path = PathHelper.GetFile("xlsx/TestIssue772.xlsx");
        var testValue = _excelImporter.Query(path, sheetName: "Supply plan(daily)", startCell: "A1")
            .Skip(19)
            .First().C
            .ToString();

        Assert.Equal("01108083-1Delta", testValue);
    }

    [Fact]
    public void TestIssue773()
    {
        var templatePath = PathHelper.GetFile("xlsx/TestIssue773_Template.xlsx");
        List<dynamic> a =
        [
            new { Id = 1, Name = "Bill", A = "a1", B = "b1", C = "c1", D = "d1", E = "e1", F = "f1", G = "g1", H = "H1" },
            new { Id = 2, Name = "Steve", A = "a2", B = "b2", C = "c2", D = "d2", E = "e2", F = "f2", G = "g2", H = "H2" },
            new { Id = 3, Name = "Ram", A = "a3", B = "b3", C = "c3", D = "d3", E = "e3", F = "f3", G = "g3", H = "H3" }
        ];

        var fill = new { t = a };
        using var path = AutoDeletingPath.Create();

         _excelTemplater.FillTemplate(path.FilePath, templatePath, fill);
        var rows = _excelImporter.Query(path.FilePath).ToList();

        Assert.Equal("H1", rows[4].AF);
        Assert.Equal("c3", rows[6].AA);
        Assert.Equal("Ram", rows[6].B);
    }

    [Fact]
    public void TestIssue789()
    {
        var path = PathHelper.GetTempPath();
        var value = new[] {
            new Dictionary<string, object> { {"no","1"} },
            new Dictionary<string, object> { {"no","2"} },
            new Dictionary<string, object> { {"no","3"} },
        };
        _excelExporter.Export(path, value);

        var xml = SheetHelper.GetZipFileContent(path, "xl/worksheets/sheet1.xml");

        Assert.Contains("<x:autoFilter ref=\"A1:A4\" />", xml);
    }

    [Fact]
    public void TestIssue809()
    {
        var path = PathHelper.GetFile("xlsx/TestIssue809.xlsx");
        var rows = _excelImporter.Query(path).ToList();

        Assert.Equal(3, rows.Count);
        Assert.Null(rows[0].A);
        Assert.Equal(2, rows[2].B);
    }

    [Fact]
    public void TestIssue814()
    {
        var originPath = PathHelper.GetFile("xlsx/TestIssue186_Template.xlsx");
        using var path = AutoDeletingPath.Create();
        File.Copy(originPath, path.FilePath);

        MiniExcelPicture[] images =
        [
            new()
           {
               ImageBytes = File.ReadAllBytes(PathHelper.GetFile("images/github_logo.png")),
               SheetName = null, // default null is first sheet  
               CellAddress = "C3", // required  
           },
           new()
           {
               ImageBytes = File.ReadAllBytes(PathHelper.GetFile("images/google_logo.png")),
               PictureType = "image/png", // default PictureType = image/png  
               SheetName = "Demo",
               CellAddress = "C9", // required  
               WidthPx = 500,
               HeightPx = 500
           },
           new()
           {
               ImageBytes = File.ReadAllBytes(PathHelper.GetFile("images/google_logo.png")),
               PictureType = "image/png", // default PictureType = image/png  
               SheetName = "Demo",
               CellAddress = "E9", // required  
               WidthPx = 800,
               HeightPx = 850
           }
        ];

        _excelTemplater.AddPicture(path.FilePath, images);

        using var package = new ExcelPackage(new FileInfo(path.FilePath));

        // Check picture in the first sheet (C3)  
        var firstSheet = package.Workbook.Worksheets[0];
        var pictureInC3 = firstSheet.Drawings.OfType<OfficeOpenXml.Drawing.ExcelPicture>().FirstOrDefault(p => p.From.Column == 2 && p.From.Row == 2);
        Assert.NotNull(pictureInC3);

        // Check picture in the "Demo" sheet (C9)  
        var demoSheet = package.Workbook.Worksheets["Demo"];
        var pictureInC9 = demoSheet.Drawings.OfType<OfficeOpenXml.Drawing.ExcelPicture>().FirstOrDefault(p => p.From.Column == 2 && p.From.Row == 8);
        Assert.NotNull(pictureInC9);
    }

    [Fact]
    public void TestIssue815()
    {
        var originPath = PathHelper.GetFile("xlsx/TestIssue186_Template.xlsx");
        using var path = AutoDeletingPath.Create();
        File.Copy(originPath, path.FilePath);
        
        MiniExcelPicture[] images =
        [
            new()
            {
                ImageBytes = File.ReadAllBytes(PathHelper.GetFile("images/github_logo.png")),
                SheetName = null, // default null is first sheet  
                CellAddress = "C3", // required  
            },
            new()
            {
                ImageBytes = File.ReadAllBytes(PathHelper.GetFile("images/google_logo.png")),
                PictureType = "image/png", // default PictureType = image/png  
                SheetName = "Demo",
                CellAddress = "C9", // required  
                WidthPx = 500,
                HeightPx = 500
            },
            new()
            {
                ImageBytes = File.ReadAllBytes(PathHelper.GetFile("images/google_logo.png")),
                PictureType = "image/png", // default PictureType = image/png  
                SheetName = "Demo",
                CellAddress = "E9", // required  
                WidthPx = 800,
                HeightPx = 850
            }
        ];

        _excelTemplater.AddPicture(path.FilePath, images);

        using var package = new ExcelPackage(new FileInfo(path.FilePath));
        // Check picture in the first sheet (C3)  
        var firstSheet = package.Workbook.Worksheets[0];
        var pictureInC3 = firstSheet.Drawings.OfType<OfficeOpenXml.Drawing.ExcelPicture>().FirstOrDefault(p => p.From.Column == 2 && p.From.Row == 2);
        Assert.NotNull(pictureInC3);

        // Check picture in the "Demo" sheet (C9)  
        var demoSheet = package.Workbook.Worksheets["Demo"];
        var pictureInC9 = demoSheet.Drawings.OfType<OfficeOpenXml.Drawing.ExcelPicture>().FirstOrDefault(p => p.From.Column == 2 && p.From.Row == 8);
        Assert.NotNull(pictureInC9);

        // TODO:check C3 image WidthPx = 80px, HeightPx = 24px, C9 WidthPx=500,HeightPx=500 
    }

    [Fact]
    public void TestIssue816()
    {
        var originPath = PathHelper.GetFile("xlsx/TestIssue186_Template.xlsx");
        using var path = AutoDeletingPath.Create();
        File.Copy(originPath, path.FilePath);
        {
            MiniExcelPicture[] images =
            [
                new()
                {
                    ImageBytes = File.ReadAllBytes(PathHelper.GetFile("images/github_logo.png")),
                    SheetName = null, // default null is first sheet  
                    CellAddress = "C3", // required  
                },
                
                new()
                {
                    ImageBytes = File.ReadAllBytes(PathHelper.GetFile("images/google_logo.png")),
                    PictureType = "image/png", // default PictureType = image/png  
                    SheetName = "Demo",
                    CellAddress = "C9", // required  
                    WidthPx = 500,
                    HeightPx = 500
                },
                
                new()
                {
                    ImageBytes = File.ReadAllBytes(PathHelper.GetFile("images/google_logo.png")),
                    PictureType = "image/png", // default PictureType = image/png  
                    SheetName = "Demo",
                    CellAddress = "E9", // required  
                    WidthPx = 800,
                    HeightPx = 850
                }
            ];

            _excelTemplater.AddPicture(path.FilePath, images);

            using var package = new ExcelPackage(new FileInfo(path.FilePath));
            
            // Check picture in the first sheet (C3)  
            var firstSheet = package.Workbook.Worksheets[0];
            var pictureInC3 = firstSheet.Drawings
                .OfType<OfficeOpenXml.Drawing.ExcelPicture>()
                .FirstOrDefault(p => p.From.Column == 2 && p.From.Row == 2);
            
            Assert.NotNull(pictureInC3);

            // Check picture in the "Demo" sheet (C9)  
            var demoSheet = package.Workbook.Worksheets["Demo"];
            var pictureInC9 = demoSheet.Drawings
                .OfType<OfficeOpenXml.Drawing.ExcelPicture>()
                .FirstOrDefault(p => p.From.Column == 2 && p.From.Row == 8);
            
            Assert.NotNull(pictureInC9);
        }

        {
            MiniExcelPicture[] images =
            [
                new()
                {
                    ImageBytes = File.ReadAllBytes(PathHelper.GetFile("images/github_logo.png")),
                    SheetName = null, // default null is first sheet  
                    CellAddress = "D3", // required  
                },
                
                new()
                {
                    ImageBytes = File.ReadAllBytes(PathHelper.GetFile("images/google_logo.png")),
                    PictureType = "image/png", // default PictureType = image/png  
                    SheetName = "Demo",
                    CellAddress = "D9", // required  
                    WidthPx = 500,
                    HeightPx = 500
                },
                
                new()
                {
                    ImageBytes = File.ReadAllBytes(PathHelper.GetFile("images/google_logo.png")),
                    PictureType = "image/png", // default PictureType = image/png  
                    SheetName = "Demo",
                    CellAddress = "F9", // required  
                    WidthPx = 800,
                    HeightPx = 850
                }
            ];

            _excelTemplater.AddPicture(path.FilePath, images);

            using var package = new ExcelPackage(new FileInfo(path.FilePath));
            // Check picture in the first sheet (C3)  
            var firstSheet1 = package.Workbook.Worksheets[0];
            var pictureInC3 = firstSheet1.Drawings
                .OfType<OfficeOpenXml.Drawing.ExcelPicture>()
                .FirstOrDefault(p => p.From.Column == 2 && p.From.Row == 2);
                
            Assert.NotNull(pictureInC3);

            // Check picture in the "Demo" sheet (C9)  
            var demoSheet1 = package.Workbook.Worksheets["Demo"];
            var pictureInC9 = demoSheet1.Drawings
                .OfType<OfficeOpenXml.Drawing.ExcelPicture>()
                .FirstOrDefault(p => p.From.Column == 2 && p.From.Row == 8);
                
            Assert.NotNull(pictureInC9);

            // Check picture in the first sheet (D3)
            var firstSheet2 = package.Workbook.Worksheets[0];
            var pictureInD3 = firstSheet2.Drawings
                .OfType<OfficeOpenXml.Drawing.ExcelPicture>()
                .FirstOrDefault(p => p.From.Column == 3 && p.From.Row == 2);
                
            Assert.NotNull(pictureInD3);

            // Check picture in the "Demo" sheet (D9)
            var demoSheet2 = package.Workbook.Worksheets["Demo"];
            var pictureInD9 = demoSheet2.Drawings
                .OfType<OfficeOpenXml.Drawing.ExcelPicture>()
                .FirstOrDefault(p => p.From.Column == 3 && p.From.Row == 8);
                
            Assert.NotNull(pictureInD9);

            // Check picture in the "Demo" sheet (F9)
            var pictureInF9 = demoSheet2.Drawings
                .OfType<OfficeOpenXml.Drawing.ExcelPicture>()
                .FirstOrDefault(p => p.From.Column == 5 && p.From.Row == 8);
                
            Assert.NotNull(pictureInF9);
        }
    }

    [Theory]
    [InlineData("DateTimeMidnight", DateOnlyConversionMode.None, true)]
    [InlineData("DateTimeNotMidnight", DateOnlyConversionMode.None, true)]
    [InlineData("DateTimeMidnight", DateOnlyConversionMode.RequireMidnight, false)]
    [InlineData("DateTimeNotMidnight", DateOnlyConversionMode.RequireMidnight, true)]
    [InlineData("DateTimeMidnight", DateOnlyConversionMode.IgnoreTimePart, false)]
    [InlineData("DateTimeNotMidnight", DateOnlyConversionMode.IgnoreTimePart, false)]
    public void TestIssue869(string fileName, DateOnlyConversionMode mode, bool throwsException)
    {
        var path = PathHelper.GetFile($"xlsx/TestIssue869/{fileName}.xlsx");
        var config = new OpenXmlConfiguration { DateOnlyConversionMode = mode };
        
        List<Issue869> TestFn() => _excelImporter.Query<Issue869>(path, configuration: config).ToList();
        if (throwsException)
        {
            Assert.Throws<ValueNotAssignableException>(TestFn);
        }
        else
        {
            try
            {
                var result = TestFn();
                Assert.Equal(new DateOnly(2025, 1, 1), result[0].Date);
            }
            catch (Exception ex)
            {
                Assert.Fail($"No exception should be thrown, but one was still thrown: {ex}.");
            }
        }
    }

    [Fact]
    public void TestIssue876()
    {
        var someTable = new[] 
        {
            new { Name = "Jack", Age = 25 }, 
        };

        var sheets = new Dictionary<string, object>
        {
            ["SomeVeryLongNameWithMoreThan31Characters"] = someTable
        };

        Assert.Throws<ArgumentException>(() =>
        {
            using var outputPath = AutoDeletingPath.Create();
            _excelExporter.Export(outputPath.ToString(), sheets);
        });
    }
    
    [Fact]
    public void TestIssue880_ShouldThrowNotSerializableException()
    {
        Issue880[] toExport = [new() { Test = "test" }];
        
        Assert.Throws<MemberNotSerializableException>(() =>
        {
            using var ms = new MemoryStream();
            _excelExporter.Export(ms, toExport);
        });
    }
    
    [Fact]
    public void TestIssue881()
    {
        Assert.Throws<ValueNotAssignableException>(() =>
        {
            _ = _excelImporter.Query<Issue409Dto>(PathHelper.GetFile("xlsx/TestIssue881.xlsx")).ToList();
        });
    }

    [Fact]
    public void TestIssue888_ShouldIgnoreFrame()
    {
        var xlsxPath = PathHelper.GetFile("xlsx/Issue888_DataWithFrame.xlsx");

        using var stream = File.OpenRead(xlsxPath);
        var dataRead = _excelImporter.Query<Issue888Dto>(stream, startCell: "A2").ToArray();

        Assert.Equal("Key1", dataRead[0].Key);
        Assert.Equal("Value1", dataRead[0].Value);
        Assert.Equal("Key2", dataRead[1].Key);
        Assert.Equal("Value2", dataRead[1].Value);
    }

    [Fact]
    public void TestIssue915()
    {
        var templatePath = PathHelper.GetFile("xlsx/TestIssue915.xlsx");
        var value = new Dictionary<string,object>
        {
            ["Data"] = new[] 
            { 
                new { Name = "Hill", Altitude = 6m }, 
                new { Name = "Mount", Altitude = 7.4m }, 
                new { Name = "Peak", Altitude = 8.6m } 
            }
        };
        
        using var path = AutoDeletingPath.Create();
        _excelTemplater.FillTemplate(path.ToString(), templatePath, value);

        var result = _excelImporter.Query(path.ToString(), true).ToList();
        
        Assert.Equal(6, result[0].Altitude);
        Assert.Equal(7.4, result[1].Altitude);
        Assert.Equal(8.6, result[2].Altitude);
    }

    [Fact]
    public void TestIssue951()
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

        // must not throw
        _excelTemplater.FillTemplate(path.ToString(), templatePath, value);
    }
}
