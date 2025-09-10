using System.Collections;
using System.ComponentModel;
using System.Data;
using System.Data.SQLite;
using System.Globalization;
using System.Text;
using System.Text.RegularExpressions;
using ClosedXML.Excel;
using Dapper;
using MiniExcelLibs.Attributes;
using MiniExcelLibs.Csv;
using MiniExcelLibs.Exceptions;
using MiniExcelLibs.OpenXml;
using MiniExcelLibs.Picture;
using MiniExcelLibs.Tests.Utils;
using MiniExcelLibs.Utils;
using Newtonsoft.Json;
using NPOI.XSSF.UserModel;
using OfficeOpenXml;
using OfficeOpenXml.Drawing;
using Xunit;
using Xunit.Abstractions;
using static MiniExcelLibs.Tests.MiniExcelOpenXmlTests;
using TableStyles = MiniExcelLibs.OpenXml.TableStyles;

namespace MiniExcelLibs.Tests;

public class MiniExcelIssueTests(ITestOutputHelper output)
{
    private readonly ITestOutputHelper _output = output;

    /// <summary>
    /// https://github.com/mini-software/MiniExcel/issues/549
    /// </summary>
    [Fact]
    public void TestIssue549()
    {
        var data = new[]
        {
            new{id=1,name="jack"},
            new{id=2,name="mike"},
        };
        using var file = AutoDeletingPath.Create();
        var path = file.ToString();

        MiniExcel.SaveAs(path, data);
        var rows = MiniExcel.Query(path, true).ToList();
        {
            using var stream = new FileStream(path, FileMode.Open, FileAccess.Read);
            using var workbook = new XSSFWorkbook(stream);

            var sheet = workbook.GetSheetAt(0);
            var a2 = sheet.GetRow(1).GetCell(0);
            var b2 = sheet.GetRow(1).GetCell(1);
            Assert.Equal((string)rows[0].id.ToString(), a2.NumericCellValue.ToString());
            Assert.Equal((string)rows[0].name.ToString(), b2.StringCellValue);
        }
    }

    [Fact]
    public void TestIssue24020201()
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
        MiniExcel.SaveAsByTemplate(path.ToString(), templatePath, data);
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
                new{ ITM=1 },
                new{ ITM=2 },
                new{ ITM=3 }
            }
        };
        MiniExcel.SaveAsByTemplate(path.ToString(), templatePath, data);

        var rows = MiniExcel.Query(path.ToString()).ToList();
        Assert.Equal(rows[2].A, 1);
        Assert.Equal(rows[3].A, 2);
        Assert.Equal(rows[4].A, 3);
    }

    [Fact]
    public void TestPR10()
    {
        var path = PathHelper.GetFile("csv/TestIssue142.csv");
        var config = new CsvConfiguration
        {
            SplitFn = row => Regex.Split(row, "[\t,](?=(?:[^\"]|\"[^\"]*\")*$)")
                .Select(s => Regex.Replace(s.Replace("\"\"", "\""), "^\"|\"$", ""))
                .ToArray()
        };
        var rows = MiniExcel.Query(path, configuration: config).ToList();
    }

    [Fact]
    public void TestIssue289()
    {
        using var path = AutoDeletingPath.Create();
        DescriptionEnumDto[] value =
        [
            new() { Name="0001", UserType=DescriptionEnum.V1 },
            new() { Name="0002", UserType=DescriptionEnum.V2 },
            new() { Name="0003", UserType=DescriptionEnum.V3 }
        ];
        MiniExcel.SaveAs(path.ToString(), value);

        var rows = MiniExcel.Query<DescriptionEnumDto>(path.ToString()).ToList();

        Assert.Equal(DescriptionEnum.V1, rows[0].UserType);
        Assert.Equal(DescriptionEnum.V2, rows[1].UserType);
        Assert.Equal(DescriptionEnum.V3, rows[2].UserType);
    }

    private class DescriptionEnumDto
    {
        public string Name { get; set; }
        public DescriptionEnum UserType { get; set; }
    }

    private enum DescriptionEnum
    {
        [Description("General User")] V1,
        [Description("General Administrator")] V2,
        [Description("Super Administrator")] V3
    }

    /// <summary>
    /// https://gitee.com/dotnetchina/MiniExcel/issues/I4X92G
    /// </summary>
    [Fact]
    public void TestIssueI4X92G()
    {
        using var file = AutoDeletingPath.Create(ExcelType.CSV);
        var path = file.ToString();

        {
            var value = new[]
            {
                new { ID = 1, Name = "Jack", InDate = new DateTime(2021,01,03)},
                new { ID = 2, Name = "Henry", InDate = new DateTime(2020,05,03)}
            };
            MiniExcel.SaveAs(path, value);
            var content = File.ReadAllText(path);
            Assert.Equal(
                """
                 ID,Name,InDate
                 1,Jack,"2021-01-03 00:00:00"
                 2,Henry,"2020-05-03 00:00:00"

                 """,
                content);
        }
        {
            var value = new { ID = 3, Name = "Mike", InDate = new DateTime(2021, 04, 23) };
            var rowsWritten = MiniExcel.Insert(path, value);
            Assert.Equal(1, rowsWritten);

            var content = File.ReadAllText(path);
            Assert.Equal(
                """
                 ID,Name,InDate
                 1,Jack,"2021-01-03 00:00:00"
                 2,Henry,"2020-05-03 00:00:00"
                 3,Mike,"2021-04-23 00:00:00"

                 """,
                content);
        }
        {
            var value = new[]
            {
                new { ID=4,Name ="Frank",InDate=new DateTime(2021,06,07)},
                new { ID=5,Name ="Gloria",InDate=new DateTime(2022,05,03)},
            };
            var rowsWritten = MiniExcel.Insert(path, value);
            Assert.Equal(2, rowsWritten);

            var content = File.ReadAllText(path);
            Assert.Equal(
                """
                 ID,Name,InDate
                 1,Jack,"2021-01-03 00:00:00"
                 2,Henry,"2020-05-03 00:00:00"
                 3,Mike,"2021-04-23 00:00:00"
                 4,Frank,"2021-06-07 00:00:00"
                 5,Gloria,"2022-05-03 00:00:00"

                 """,
                content);
        }
    }


    /// <summary>
    /// Exception : MiniExcelLibs.Exceptions.ExcelInvalidCastException: 'ColumnName : Date, CellRow : 2, Value : 2021-01-31 10:03:00 +08:00, it can't cast to DateTimeOffset type.'
    /// </summary>
    [Fact]
    public void TestIssue430()
    {
        using var path = AutoDeletingPath.Create();
        var value = new[]
        {
            new TestIssue430Dto{ Date=DateTimeOffset.Parse("2021-01-31 10:03:00 +05:00")}
        };
        MiniExcel.SaveAs(path.ToString(), value);
        var rows = MiniExcel.Query<TestIssue430Dto>(path.ToString()).ToArray();
        Assert.Equal("2021-01-31 10:03:00 +05:00", rows[0].Date.ToString("yyyy-MM-dd HH:mm:ss zzz"));
    }

    private class TestIssue430Dto
    {
        [ExcelFormat("yyyy-MM-dd HH:mm:ss zzz")]
        public DateTimeOffset Date { get; set; }
    }

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
        MiniExcel.SaveAs(path.ToString(), reader, configuration: config);
        var xml = Helpers.GetZipFileContent(path.ToString(), "xl/worksheets/sheet1.xml");

        Assert.Contains("<x:autoFilter ref=\"A1:B3\" />", xml);
        Assert.Contains("<x:dimension ref=\"A1:B3\" />", xml);
    }

    /// <summary>
    /// [ · Issue #413 · MiniExcel/MiniExcel]
    /// (https://github.com/MiniExcel/MiniExcel/issues/413)
    /// </summary>
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
        MiniExcel.SaveAsByTemplate(path.ToString(), templatePath, value);
        var rows = MiniExcel.Query(path.ToString()).ToList();

        Assert.Equal("2022-12-25 00:00:00", rows[1].B);
        Assert.Equal("2022-09-23 00:00:00", rows[2].B);
    }

    /// <summary>
    /// [SaveAs Support empty sharedstring · Issue #405 · MiniExcel/MiniExcel]
    /// (https://github.com/MiniExcel/MiniExcel/issues/405)
    /// </summary>
    [Fact]
    public void TestIssue405()
    {
        using var path = AutoDeletingPath.Create();
        var value = new[] { new { id = 1, name = "test" } };
        MiniExcel.SaveAs(path.ToString(), value);

        var xml = Helpers.GetZipFileContent(path.ToString(), "xl/sharedStrings.xml");
        Assert.StartsWith("<sst xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"", xml);
    }

    /// <summary>
    /// Using stream.SaveAs will close the Stream automatically when Specifying excelType
    /// https://gitee.com/dotnetchina/MiniExcel/issues/I57WMM
    /// </summary>
    [Fact]
    public void TestIssueI57WMM()
    {
        Dictionary<string, object>[] sheets = [new() { ["ID"] = "0001", ["Name"] = "Jack" }];
        using var stream = new MemoryStream();
        stream.SaveAs(sheets, excelType: ExcelType.CSV);
        stream.Position = 0;

        // convert stream to string
        using var reader = new StreamReader(stream);
        var text = reader.ReadToEnd();

        Assert.Equal("ID,Name\r\n0001,Jack\r\n", text);
    }

    [Fact]
    public void TestIssue370()
    {
        var config = new OpenXmlConfiguration
        {
            DynamicColumns =
            [
                new DynamicExcelColumn("id") { Ignore=true },
                new DynamicExcelColumn("name") { Index=1,Width=10 },
                new DynamicExcelColumn("createdate") { Index=0, Format="yyyy-MM-dd", Width=15 },
                new DynamicExcelColumn("point") { Index=2, Name="Account Point" }
            ]
        };
        using var path = AutoDeletingPath.Create();
        var json = JsonConvert.SerializeObject(new[]
        {
            new
            {
                id = 1,
                name = "Jack",
                createdate = new DateTime(2022, 04, 12),
                point = 123.456
            }
        }, Formatting.Indented);

        var value = JsonConvert.DeserializeObject<List<Dictionary<string, object>>>(json);
        MiniExcel.SaveAs(path.ToString(), value, configuration: config);

        var rows = MiniExcel.Query(path.ToString()).ToList();
        Assert.Equal("createdate", rows[0].A);
        Assert.Equal(new DateTime(2022, 04, 12), rows[1].A);
        Assert.Equal("name", rows[0].B);
        Assert.Equal("Jack", rows[1].B);
        Assert.Equal("Account Point", rows[0].C);
        Assert.Equal(123.456, rows[1].C);
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
        MiniExcel.SaveAs(path.ToString(), value, configuration: config);

        var rows = MiniExcel.Query(path.ToString()).ToList();
        Assert.Equal("createdate", rows[0].A);
        Assert.Equal(new DateTime(2022, 04, 12), rows[1].A);
        Assert.Equal("name", rows[0].B);
        Assert.Equal("Jack", rows[1].B);
        Assert.Equal("Account Point", rows[0].C);
        Assert.Equal(123.456, rows[1].C);
    }

    [Fact]
    public void TestIssueI4ZYUU()
    {
        using var path = AutoDeletingPath.Create();
        TestIssueI4ZYUUDto[] value = [new() { MyProperty = "1", MyProperty2 = new DateTime(2022, 10, 15) }];
        MiniExcel.SaveAs(path.ToString(), value);

        var rows = MiniExcel.Query(path.ToString()).ToList();
        Assert.Equal("2022-10", rows[1].B);

        using var workbook = new XLWorkbook(path.ToString());
        var ws = workbook.Worksheet(1);

        Assert.True(ws.Column("A").Width > 0);
        Assert.True(ws.Column("B").Width > 0);
    }

    private class TestIssueI4ZYUUDto
    {
        [ExcelColumn(Name = "ID", Index = 0)]
        public string MyProperty { get; set; }
        [ExcelColumn(Name = "CreateDate", Index = 1, Format = "yyyy-MM", Width = 100)]
        public DateTime MyProperty2 { get; set; }
    }

    [Fact]
    public void TestIssue360()
    {
        var path = PathHelper.GetFile("xlsx/NotDuplicateSharedStrings_10x100.xlsx");
        var config = new OpenXmlConfiguration { SharedStringCacheSize = 1 };
        var sheets = MiniExcel.GetSheetNames(path);
        foreach (var sheetName in sheets)
        {
            var dt = MiniExcel.QueryAsDataTable(path, useHeaderRow: true, sheetName: sheetName, configuration: config);
        }
    }

    [Fact]
    public void TestIssue117()
    {
        {
            var cache = new SharedStringsDiskCache();
            for (int i = 0; i < 100; i++)
            {
                cache[i] = i.ToString();
            }
            for (int i = 0; i < 100; i++)
            {
                Assert.Equal(i.ToString(), cache[i]);
            }
            Assert.Equal(100, cache.Count);
        }
        {
            var cache = new SharedStringsDiskCache();
            Assert.Empty(cache);
        }
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
            MiniExcel.SaveAs(path.ToString(), reader);
            var xml = Helpers.GetZipFileContent(path.ToString(), "xl/worksheets/sheet1.xml");
            var cnt = Regex.Matches(xml, "<x:autoFilter ref=\"A1:B3\" />").Count;
        }
        {
            using var table = new DataTable();
            table.Columns.Add("id", typeof(int));
            table.Columns.Add("name", typeof(string));
            table.Rows.Add(1, "Jack");
            table.Rows.Add(2, "Mike");

            using var path = AutoDeletingPath.Create();
            var reader = table.CreateDataReader();
            MiniExcel.SaveAs(path.ToString(), reader, false);
            var xml = Helpers.GetZipFileContent(path.ToString(), "xl/worksheets/sheet1.xml");
            var cnt = Regex.Matches(xml, "<x:autoFilter ref=\"A1:B2\" />").Count;
        }
        {
            using var table = new DataTable();
            table.Columns.Add("id", typeof(int));
            table.Columns.Add("name", typeof(string));

            using var path = AutoDeletingPath.Create();
            var reader = table.CreateDataReader();
            MiniExcel.SaveAs(path.ToString(), reader);
            var xml = Helpers.GetZipFileContent(path.ToString(), "xl/worksheets/sheet1.xml");
            var cnt = Regex.Matches(xml, "<x:autoFilter ref=\"A1:B1\" />").Count;
        }
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
            MiniExcel.SaveAs(path.ToString(), reader, configuration: config);

            var xml = Helpers.GetZipFileContent(path.ToString(), "xl/worksheets/sheet1.xml");
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
            MiniExcel.SaveAs(path.ToString(), reader, false, configuration: config);

            var xml = Helpers.GetZipFileContent(path.ToString(), "xl/worksheets/sheet1.xml");
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
            MiniExcel.SaveAs(path.ToString(), reader, configuration: config);

            var xml = Helpers.GetZipFileContent(path.ToString(), "xl/worksheets/sheet1.xml");
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
                MiniExcel.SaveAs(path.ToString(), reader, configuration: config);
            }

            var xml = Helpers.GetZipFileContent(path.ToString(), "xl/worksheets/sheet1.xml");
            var cnt = Regex.Matches(xml, "autoFilter").Count;
            Assert.Equal(count, cnt);
        }

        {
            const string xlsxPath = "../../../../../samples/xlsx/Test5x2.xlsx";
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
                    var rows = stream.Query();
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
                MiniExcel.SaveAs(path.ToString(), reader, configuration: config);
            }

            var xml = Helpers.GetZipFileContent(path.ToString(), "xl/worksheets/sheet1.xml");
            var cnt = Regex.Matches(xml, "autoFilter").Count;
            Assert.Equal(count, cnt);
        }
    }

    [Fact]
    public async Task TestIssue307()
    {
        using var file = AutoDeletingPath.Create();
        var path = file.ToString();
        var value = new[] { new { id = 1, name = "Jack" } };

        await MiniExcel.SaveAsAsync(path, value);
        Assert.Throws<IOException>(() => MiniExcel.SaveAs(path, value));

        await MiniExcel.SaveAsAsync(path, value, overwriteFile: true);
        await Assert.ThrowsAsync<IOException>(async () => await MiniExcel.SaveAsAsync(path, value));
        await MiniExcel.SaveAsAsync(path, value, overwriteFile: true);
    }

    [Fact]
    public void TestIssue310()
    {
        using var path = AutoDeletingPath.Create();
        var value = new[] { new TestIssue310Dto { V1 = null }, new TestIssue310Dto { V1 = 2 } };
        MiniExcel.SaveAs(path.ToString(), value);
        var rows = MiniExcel.Query<TestIssue310Dto>(path.ToString()).ToList();
    }

    [Fact]
    public void TestIssue310Fix497()
    {
        using var path = AutoDeletingPath.Create();
        var value = new[]
        {
            new TestIssue310Dto { V1 = null },
            new TestIssue310Dto { V1 = 2 }
        };
        MiniExcel.SaveAs(path.ToString(), value, configuration: new OpenXmlConfiguration { EnableWriteNullValueCell = false });
        var rows = MiniExcel.Query<TestIssue310Dto>(path.ToString()).ToList();
    }

    private class TestIssue310Dto
    {
        public int? V1 { get; set; }
    }

    /// <summary>
    /// Excel was unable to open the file https://github.com/mini-software/MiniExcel/issues/343
    /// </summary>
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
        MiniExcel.SaveAs(path.ToString(), reader);

        var rows = MiniExcel.Query(path.ToString(), true).ToArray();
        Assert.Equal(date, rows[0].time1);
        Assert.Equal(date, rows[0].time2);
    }

    [Fact]
    public void TestIssueI4YCLQ_2()
    {
        var c = ExcelOpenXmlUtils.ConvertColumnName(1);
        var c2 = ExcelOpenXmlUtils.ConvertColumnName(3);
        var path = PathHelper.GetFile("xlsx/TestIssueI4YCLQ_2.xlsx");
        var rows = MiniExcel.Query<TestIssueI4YCLQ_2Dto>(path, startCell: "B2").ToList();

        Assert.Null(rows[0].站点编码);
        Assert.Equal("N1", rows[0].站址名称);
        Assert.Equal("a", rows[0].值1);
        Assert.Equal("b", rows[0].值2);
        Assert.Equal("c", rows[0].值3);
        Assert.Equal("A1", rows[0].资源ID);
        Assert.Equal("A", rows[0].值4);
        Assert.Equal("B", rows[0].值5);
        Assert.Equal("C", rows[0].值6);
        Assert.Null(rows[0].值7);
        Assert.Null(rows[0].值8);
    }

    private class TestIssueI4YCLQ_2Dto
    {
        [ExcelColumnIndex("A")]
        public string 站点编码 { get; set; }
        [ExcelColumnIndex("B")]
        public string 站址名称 { get; set; }
        [ExcelColumnIndex("C")]
        public string 值1 { get; set; }
        [ExcelColumnIndex("D")]
        public string 值2 { get; set; }
        [ExcelColumnIndex("E")]
        public string 值3 { get; set; }
        [ExcelColumnIndex("F")]
        public string 资源ID { get; set; }
        [ExcelColumnIndex("G")]
        public string 值4 { get; set; }
        [ExcelColumnIndex("H")]
        public string 值5 { get; set; }
        [ExcelColumnIndex("I")]
        public string 值6 { get; set; }
        public string 值7 { get; set; }
        [ExcelColumnName("NotExist")]
        public string 值8 { get; set; }
    }

    [Fact]
    public async Task TestIssue338()
    {
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
        {
            var path = PathHelper.GetFile("csv/TestIssue338.csv");
            var row = (await MiniExcel.QueryAsync(path)).FirstOrDefault();
            Assert.Equal("���Ĳ�������", row!.A);
        }
        {
            var path = PathHelper.GetFile("csv/TestIssue338.csv");
            var config = new CsvConfiguration
            {
                StreamReaderFunc = stream => new StreamReader(stream, Encoding.GetEncoding("gb2312"))
            };
            var row = (await MiniExcel.QueryAsync(path, configuration: config)).FirstOrDefault();
            Assert.Equal("中文测试内容", row!.A);
        }
        {
            var path = PathHelper.GetFile("csv/TestIssue338.csv");
            var config = new CsvConfiguration
            {
                StreamReaderFunc = stream => new StreamReader(stream, Encoding.GetEncoding("gb2312"))
            };
            await using (var stream = new FileStream(path, FileMode.Open, FileAccess.Read))
            {
                var row = (await stream.QueryAsync(configuration: config, excelType: ExcelType.CSV)).FirstOrDefault();
                Assert.Equal("中文测试内容", row!.A);
            }
        }
    }

    [Fact]
    public void TestIssueI4WM67()
    {
        using var path = AutoDeletingPath.Create();
        var templatePath = PathHelper.GetFile("xlsx/TestIssueI4WM67.xlsx");
        var value = new Dictionary<string, object>
        {
            ["users"] = Array.Empty<TestIssueI4WM67Dto>()
        };
        MiniExcel.SaveAsByTemplate(path.ToString(), templatePath, value);
        var rows = MiniExcel.Query(path.ToString()).ToList();
        Assert.Single(rows);
    }

    private class TestIssueI4WM67Dto
    {
        public int ID { get; set; }
        public string Name { get; set; }
    }

    [Fact]
    public void TestIssueI4WXFB()
    {
        {
            using var path = AutoDeletingPath.Create();
            var templatePath = PathHelper.GetFile("xlsx/TestIssueI4WXFB.xlsx");
            var value = new Dictionary<string, object>
            {
                ["Name"] = "Jack",
                ["Amount"] = 1000,
                ["Department"] = "HR"
            };
            MiniExcel.SaveAsByTemplate(path.ToString(), templatePath, value);
        }

        {
            var config = new OpenXmlConfiguration
            {
                IgnoreTemplateParameterMissing = false
            };
            using var path = AutoDeletingPath.Create();
            var templatePath = PathHelper.GetFile("xlsx/TestIssueI4WXFB.xlsx");
            var value = new Dictionary<string, object>
            {
                ["Name"] = "Jack",
                ["Amount"] = 1000,
                ["Department"] = "HR"
            };
            Assert.Throws<KeyNotFoundException>(() => MiniExcel.SaveAsByTemplate(path.ToString(), templatePath, value, config));
        }
    }

    [Fact]
    public void TestIssueI4WDA9()
    {
        using var path = AutoDeletingPath.Create(ExcelType.CSV);
        var value = new DataTable();
        {
            value.Columns.Add("\"name\"");
            value.Rows.Add("\"Jack\"");
        }

        MiniExcel.SaveAs(path.ToString(), value);
        Assert.Equal("\"\"\"name\"\"\"\r\n\"\"\"Jack\"\"\"\r\n", File.ReadAllText(path.ToString()));
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
        MiniExcel.SaveAs(path.ToString(), data, configuration: config);
        CultureInfo.CurrentCulture = CultureInfo.GetCultureInfo(cln);
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
        MiniExcel.SaveAs(path.ToString(), data);

        var rows = MiniExcel.Query(path.ToString(), startCell: "A2").ToArray();
        Assert.Equal(1.5, rows[2].B);
        Assert.Equal(1.5, rows[2].C);

        CultureInfo.CurrentCulture = CultureInfo.GetCultureInfo(cln);
    }

    private class TestIssue331Dto
    {
        public int Number { get; set; }
        public decimal DecimalNumber { get; set; }
        public double DoubleNumber { get; set; }
        public string Text { get; set; }
    }

    [Fact]
    public void TestIssueI4TXGT()
    {
        using var file = AutoDeletingPath.Create();
        var path = file.ToString();
        var value = new[] { new TestIssueI4TXGTDto { ID = 1, Name = "Apple", Spc = "X", Up = 6999 } };

        MiniExcel.SaveAs(path, value);
        {
            var rows = MiniExcel.Query(path).ToList();
            Assert.Equal("ID", rows[0].A);
            Assert.Equal("Name", rows[0].B);
            Assert.Equal("Specification", rows[0].C);
            Assert.Equal("Unit Price", rows[0].D);
        }
        {
            var rows = MiniExcel.Query<TestIssueI4TXGTDto>(path).ToList();
            Assert.Equal(1, rows[0].ID);
            Assert.Equal("Apple", rows[0].Name);
            Assert.Equal("X", rows[0].Spc);
            Assert.Equal(6999, rows[0].Up);
        }
    }

    private class TestIssueI4TXGTDto
    {
        public int ID { get; set; }
        public string Name { get; set; }
        [DisplayName("Specification")]
        public string Spc { get; set; }
        [DisplayName("Unit Price")]
        public decimal Up { get; set; }
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
                file = File.ReadAllBytes(PathHelper.GetFile("images/TestIssue327.png"))
            },
            new
            {
                id=2,
                name="Henry",
                indate=new DateTime(2022,4,10),
                file = File.ReadAllBytes(PathHelper.GetFile("other/TestIssue327.txt"))
            },
        };
        MiniExcel.SaveAs(path.ToString(), value);

        var rowIndx = 0;
        using var reader = MiniExcel.GetReader(path.ToString(), true);

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
                if (rowIndx == 0 && i == 3) Assert.Equal(File.ReadAllBytes(PathHelper.GetFile("images/TestIssue327.png")), v);
                if (rowIndx == 1 && i == 0) Assert.Equal(2.0, v);
                if (rowIndx == 1 && i == 1) Assert.Equal("Henry", v);
                if (rowIndx == 1 && i == 2) Assert.Equal(new DateTime(2022, 4, 10), v);
                if (rowIndx == 1 && i == 3) Assert.Equal(File.ReadAllBytes(PathHelper.GetFile("other/TestIssue327.txt")), v);
            }
            rowIndx++;
        }

        //TODO:How to resolve empty body sheet?
    }

    [Fact]
    public void TestIssue327()
    {
        using var path = AutoDeletingPath.Create();
        var value = new[]
        {
            new { id = 1, file = File.ReadAllBytes(PathHelper.GetFile("images/TestIssue327.png")) },
            new { id = 2, file = File.ReadAllBytes(PathHelper.GetFile("other/TestIssue327.txt")) },
            new { id = 3, file = File.ReadAllBytes(PathHelper.GetFile("other/TestIssue327.html")) },
        };
        MiniExcel.SaveAs(path.ToString(), value);
        var rows = MiniExcel.Query(path.ToString(), true).ToList();

        Assert.Equal(value[0].file, rows[0].file);
        Assert.Equal(value[1].file, rows[1].file);
        Assert.Equal(value[2].file, rows[2].file);
        Assert.Equal("Hello MiniExcel", Encoding.UTF8.GetString(rows[1].file));
        Assert.Equal("<html>Hello MiniExcel</html>", Encoding.UTF8.GetString(rows[2].file));
    }

    [Fact]
    public void TestIssue316()
    {
        // XLSX
        {
            {
                using var file = AutoDeletingPath.Create();
                var path = file.ToString();
                var value = new[]
                {
                    new{ Amount=123_456.789M, CreateTime=DateTime.Parse("2018-01-31",CultureInfo.InvariantCulture)}
                };
                var config = new OpenXmlConfiguration
                {
                    Culture = new CultureInfo("fr-FR"),
                };
                MiniExcel.SaveAs(path, value, configuration: config);

                //Datetime error
                {
                    Assert.Throws<ExcelInvalidCastException>(() =>
                    {
                        var conf = new OpenXmlConfiguration
                        {
                            Culture = new CultureInfo("en-US"),
                        };
                        _ = MiniExcel.Query<TestIssue316Dto>(path, configuration: conf).ToList();
                    });
                }

                // dynamic
                {
                    var rows = MiniExcel.Query(path, true).ToList();
                    Assert.Equal("123456,789", rows[0].Amount);
                    Assert.Equal("31/01/2018 00:00:00", rows[0].CreateTime);
                }
            }

            // type
            {
                using var file = AutoDeletingPath.Create();
                var path = file.ToString();
                var value = new[]
                {
                    new{ Amount=123_456.789M, CreateTime=DateTime.Parse("2018-05-12", CultureInfo.InvariantCulture)}
                };
                {
                    var config = new OpenXmlConfiguration
                    {
                        Culture = new CultureInfo("fr-FR"),
                    };
                    MiniExcel.SaveAs(path, value, configuration: config);
                }

                {
                    var rows = MiniExcel.Query(path, true).ToList();
                    Assert.Equal("123456,789", rows[0].Amount);
                    Assert.Equal("12/05/2018 00:00:00", rows[0].CreateTime);
                }

                {
                    var config = new OpenXmlConfiguration
                    {
                        Culture = new CultureInfo("en-US"),
                    };
                    var rows = MiniExcel.Query<TestIssue316Dto>(path, configuration: config).ToList();

                    Assert.Equal("2018-12-05 00:00:00", rows[0].CreateTime.ToString("yyyy-MM-dd HH:mm:ss"));
                    Assert.Equal(123456789m, rows[0].Amount);
                }

                {
                    var config = new OpenXmlConfiguration
                    {
                        Culture = new CultureInfo("fr-FR"),
                    };
                    var rows = MiniExcel.Query<TestIssue316Dto>(path, configuration: config).ToList();

                    Assert.Equal("2018-05-12 00:00:00", rows[0].CreateTime.ToString("yyyy-MM-dd HH:mm:ss"));
                    Assert.Equal(123456.789m, rows[0].Amount);
                }
            }
        }

        // CSV
        {
            {
                using var file = AutoDeletingPath.Create(ExcelType.CSV);
                var path = file.ToString();
                var value = new[]
                {
                    new{ Amount=123_456.789M, CreateTime=DateTime.Parse("2018-01-31", CultureInfo.InvariantCulture)}
                };
                var config = new CsvConfiguration
                {
                    Culture = new CultureInfo("fr-FR"),
                };
                MiniExcel.SaveAs(path, value, configuration: config);

                //Datetime error
                {
                    Assert.Throws<ExcelInvalidCastException>(() =>
                    {
                        var conf = new CsvConfiguration
                        {
                            Culture = new CultureInfo("en-US"),
                        };
                        _ = MiniExcel.Query<TestIssue316Dto>(path, configuration: conf).ToList();
                    });
                }

                // dynamic
                {
                    var rows = MiniExcel.Query(path, true).ToList();
                    Assert.Equal("123456,789", rows[0].Amount);
                    Assert.Equal("31/01/2018 00:00:00", rows[0].CreateTime);
                }
            }

            // type
            {
                var path = PathHelper.GetTempFilePath("csv");
                var value = new[]
                {
                    new{ Amount=123_456.789M, CreateTime=DateTime.Parse("2018-05-12", CultureInfo.InvariantCulture)}
                };
                {
                    var config = new CsvConfiguration
                    {
                        Culture = new CultureInfo("fr-FR"),
                    };
                    MiniExcel.SaveAs(path, value, configuration: config);
                }

                {
                    var rows = MiniExcel.Query(path, true).ToList();
                    Assert.Equal("123456,789", rows[0].Amount);
                    Assert.Equal("12/05/2018 00:00:00", rows[0].CreateTime);
                }

                {
                    var config = new CsvConfiguration
                    {
                        Culture = new CultureInfo("en-US"),
                    };
                    var rows = MiniExcel.Query<TestIssue316Dto>(path, configuration: config).ToList();

                    Assert.Equal("2018-12-05 00:00:00", rows[0].CreateTime.ToString("yyyy-MM-dd HH:mm:ss"));
                    Assert.Equal(123456789m, rows[0].Amount);
                }

                {
                    var config = new CsvConfiguration
                    {
                        Culture = new CultureInfo("fr-FR"),
                    };
                    var rows = MiniExcel.Query<TestIssue316Dto>(path, configuration: config).ToList();

                    Assert.Equal("2018-05-12 00:00:00", rows[0].CreateTime.ToString("yyyy-MM-dd HH:mm:ss"));
                    Assert.Equal(123456.789m, rows[0].Amount);
                }
            }
        }
    }

    private class TestIssue316Dto
    {
        public decimal Amount { get; set; }
        public DateTime CreateTime { get; set; }
    }

    /// <summary>
    /// https://github.com/mini-software/MiniExcel/issues/325
    /// </summary>
    [Fact]
    public void TestIssue325()
    {
        using var path = AutoDeletingPath.Create();
        var value = new Dictionary<string, object>
        {
            { "sheet1",new[]{ new { id = 1, date = DateTime.Parse("2022-01-01") } }},
            { "sheet2",new[]{ new { id = 2, date = DateTime.Parse("2022-01-01") } }},
        };
        MiniExcel.SaveAs(path.ToString(), value);

        var xml = Helpers.GetZipFileContent(path.ToString(), "xl/worksheets/_rels/sheet2.xml.rels");
        var cnt = Regex.Matches(xml, "Id=\"drawing2\"").Count;
        Assert.True(cnt == 1);
    }

    /// <summary>
    /// https://gitee.com/dotnetchina/MiniExcel/issues/I49RZH
    /// https://github.com/mini-software/MiniExcel/issues/305
    /// </summary>
    [Fact]
    public void TestIssueI49RZH()
    {
        // xlsx
        {
            using var path = AutoDeletingPath.Create();
            var value = new[]
            {
                new TestIssueI49RZHDto{ dd = DateTimeOffset.Parse("2022-01-22")},
                new TestIssueI49RZHDto{ dd = null}
            };
            MiniExcel.SaveAs(path.ToString(), value);

            var rows = MiniExcel.Query(path.ToString()).ToList();
            Assert.Equal("2022-01-22", rows[1].A);
        }

        //TODO:CSV
        {
            using var path = AutoDeletingPath.Create(ExcelType.CSV);
            var value = new[]
            {
                new TestIssueI49RZHDto{ dd = DateTimeOffset.Parse("2022-01-22")},
                new TestIssueI49RZHDto{ dd = null}
            };
            MiniExcel.SaveAs(path.ToString(), value);

            var rows = MiniExcel.Query(path.ToString()).ToList();
            Assert.Equal("2022-01-22", rows[1].A);
        }
    }

    private class TestIssueI49RZHDto
    {
        [ExcelFormat("yyyy-MM-dd")]
        public DateTimeOffset? dd { get; set; }
    }

    /// <summary>
    /// https://github.com/mini-software/MiniExcel/issues/312
    /// </summary>
    [Fact]
    public void TestIssue312()
    {
        //xlsx
        {
            using var path = AutoDeletingPath.Create();
            TestIssue312Dto[] value =
            [
                new() { Value = 12345.6789},
                new() { Value = null}
            ];
            MiniExcel.SaveAs(path.ToString(), value);

            var rows = MiniExcel.Query(path.ToString()).ToList();
            Assert.Equal("12,345.68", rows[1].A);
        }

        //TODO:CSV
        {
            using var path = AutoDeletingPath.Create(ExcelType.CSV);
            TestIssue312Dto[] value =
            [
                new() { Value = 12345.6789},
                new() { Value = null}
            ];
            MiniExcel.SaveAs(path.ToString(), value);

            var rows = MiniExcel.Query(path.ToString()).ToList();
            Assert.Equal("12,345.68", rows[1].A);
        }
    }

    private class TestIssue312Dto
    {
        [ExcelFormat("0,0.00")]
        public double? Value { get; set; }
    }

    /// <summary>
    /// Query type conversion error
    /// https://github.com/mini-software/MiniExcel/issues/309
    /// </summary>
    [Fact]
    public void TestIssue209()
    {
        try
        {
            var path = PathHelper.GetFile("xlsx/TestIssue309.xlsx");
            var rows = MiniExcel.Query<TestIssue209Dto>(path).ToList();
        }
        catch (ExcelInvalidCastException ex)
        {
            Assert.Equal("SEQ", ex.ColumnName);
            Assert.Equal(4, ex.Row);
            Assert.Equal("Error", ex.Value);
            Assert.Equal(typeof(int), ex.InvalidCastType);
            Assert.Equal("ColumnName: SEQ, CellRow: 4, Value: Error. The value cannot be cast to type Int32.", ex.Message);
        }
    }

    private class TestIssue209Dto
    {
        public int ID { get; set; }
        public string Name { get; set; }
        public int SEQ { get; set; }
    }

    /// <summary>
    /// [SaveAs and Query support btye[] base64 converter · Issue #318 · mini-software/MiniExcel](https://github.com/mini-software/MiniExcel/issues/318)
    /// </summary>
    [Fact]
    public void TestIssue318()
    {
        var imageByte = File.ReadAllBytes(PathHelper.GetFile("images/github_logo.png"));
        using var path = AutoDeletingPath.Create();
        var value = new[]
        {
            new { Name="github", Image=imageByte},
        };
        MiniExcel.SaveAs(path.ToString(), value);


        // import to byte[]
        {
            const string expectedBase64 = "iVBORw0KGgoAAAANSUhEUgAAABwAAAAcCAIAAAD9b0jDAAAACXBIWXMAAA7EAAAOxAGVKw4bAAAAEXRFWHRTb2Z0d2FyZQBTbmlwYXN0ZV0Xzt0AAALNSURBVEiJ7ZVLTBNBGMdndrfdIofy0ERbCgcFeYRuCy2JGOPNRA9qeIZS6YEEogQj0YMmGOqDSATxQaLRxKtRID4SgjGelUBpaQvGZ7kpII8aWtjd2dkdDxsJoS1pIh6M/k+z8833m/3+8+0OJISArRa15cT/0D8CZTYPe32+Zy+GxjzjMzOzAACDYafdZquqOG7hzJtkwUQthRC6cavv0eN+QRTBujUQQp1OV1dbffZMq1arTRaqKIok4eZTrSNjHqIo6gIIIQBgbQwpal+Z/f7dPo2GoaiNHtJut3vjPhBe7+kdfvW61Mq1nGyaX1xYjkRzsk2Z6Rm8IOTvzWs73SLwwqjHK4jCgf3lcV6VxGgiECji7AXm0gvtHYQQnue/zy8ghCRJWlxaWuV5Qsilq9cKzLYiiz04ORVLiHP6A4NPRQlhjLWsVpZlnU63Y3umRqNhGCYjPV3HsrIsMwyDsYQQejIwGEuIA/WMT1AAaDSahnoHTdPKL1vXPKVp2umoZVkWAOj1+ZOCzs7NKYTo9XqjYRcAgKIo9ZRUu9VxltGYZTQAAL5+m0kKijEmAPCrqyJCcRuOECKI4lL4ByEEYykpaE62iQIgurLi9wchhLIsry8fYwwh9PomwuEwACDbZEoKauHMgKJSU1PbOy6Hpqdpml5fPsMwn7+EOru6IYQAghKrJSloTVUFURSX02G3lRw+WulqbA4EJ9XQh4+f2s6dr65zhkLTEEIKwtqaylhCnG/fauFO1Nfde/Bw6Hm/0WiYevc+LU2vhlK2pQwNvwQAsCwrYexyOrji4lhCnOaXZRljXONoOHTk2Ju3I/5AcC3EC0JZ+cE9Bea8IqursUkUker4BsWBqpIk6aL7Sm4htzvfvByJqJORaDS3kMsvLuns6kYIJcpNCFU17pvouXlHEET1URDEnt7bo2OezbMS/vp+R3/PdfKPQ38Ccg0E/CDcpY8AAAAASUVORK5CYII=";
            var rows = MiniExcel.Query(path.ToString(), true).ToList();
            var actulBase64 = Convert.ToBase64String((byte[])rows[0].Image);
            Assert.Equal(expectedBase64, actulBase64);
        }

        // import to base64 string
        {
            var config = new OpenXmlConfiguration { EnableConvertByteArray = false };
            var rows = MiniExcel.Query(path.ToString(), true, configuration: config).ToList();
            var image = (string)rows[0].Image;
            Assert.StartsWith("@@@fileid@@@,xl/media/", image);
        }

    }


    /// <summary>
    /// SaveAs support Image type · Issue #304  https://github.com/mini-software/MiniExcel/issues/304
    /// </summary>
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
            new { Name="statck_overflow", Image=File.ReadAllBytes(PathHelper.GetFile("images/statck_overflow_logo.png"))},
        };
        MiniExcel.SaveAs(path, value);

        Assert.Contains("/xl/media/", Helpers.GetZipFileContent(path, "xl/drawings/_rels/drawing1.xml.rels"));
        Assert.Contains("ext cx=\"609600\" cy=\"190500\"", Helpers.GetZipFileContent(path, "xl/drawings/drawing1.xml"));
        Assert.Contains("/xl/drawings/drawing1.xml", Helpers.GetZipFileContent(path, "[Content_Types].xml"));
        Assert.Contains("drawing r:id=", Helpers.GetZipFileContent(path, "xl/worksheets/sheet1.xml"));
        Assert.Contains("../drawings/drawing1.xml", Helpers.GetZipFileContent(path, "xl/worksheets/_rels/sheet1.xml.rels"));
    }

    /// <summary>
    /// https://gitee.com/dotnetchina/MiniExcel/issues/I4HL54
    /// </summary>
    [Fact]
    public void TestIssueI4HL54()
    {
        using var cn = Db.GetConnection();

        using var reader = cn.ExecuteReader(@"select 'Hello World1' Text union all select 'Hello World2'");
        var templatePath = PathHelper.GetFile("xlsx/TestIssueI4HL54_Template.xlsx");
        using var path = AutoDeletingPath.Create();
        var value = new Dictionary<string, object>
        {
            { "Texts",reader}
        };
        MiniExcel.SaveAsByTemplate(path.ToString(), templatePath, value);

        var rows = MiniExcel.Query(path.ToString(), true).ToList();
        Assert.Equal("Hello World1", rows[0].Text);
        Assert.Equal("Hello World2", rows[1].Text);
    }

    /// <summary>
    /// [Prefix and suffix blank space will lost after SaveAs · Issue #294 · mini-software/MiniExcel]
    /// (https://github.com/mini-software/MiniExcel/issues/294)
    /// </summary>
    [Fact]
    public void TestIssue294()
    {
        {
            using var path = AutoDeletingPath.Create();
            var value = new[] { new { Name = "   Jack" } };
            MiniExcel.SaveAs(path.ToString(), value);
            var sheetXml = Helpers.GetZipFileContent(path.ToString(), "xl/worksheets/sheet1.xml");
            Assert.Contains("xml:space=\"preserve\"", sheetXml);
        }
        {
            using var path = AutoDeletingPath.Create();
            var value = new[] { new { Name = "Ja ck" } };
            MiniExcel.SaveAs(path.ToString(), value);
            var sheetXml = Helpers.GetZipFileContent(path.ToString(), "xl/worksheets/sheet1.xml");
            Assert.DoesNotContain("xml:space=\"preserve\"", sheetXml);
        }
        {
            using var path = AutoDeletingPath.Create();
            var value = new[] { new { Name = "Jack   " } };
            MiniExcel.SaveAs(path.ToString(), value);
            var sheetXml = Helpers.GetZipFileContent(path.ToString(), "xl/worksheets/sheet1.xml");
            Assert.Contains("xml:space=\"preserve\"", sheetXml);
        }
    }

    /// <summary>
    /// Column '' does not belong to table when csv convert to datatable #298
    /// https://github.com/mini-software/MiniExcel/issues/298
    /// </summary>
    [Fact]
    public void TestIssue298()
    {
        var path = PathHelper.GetFile("/csv/TestIssue298.csv");
#pragma warning disable CS0618 // Type or member is obsolete
        var dt = MiniExcel.QueryAsDataTable(path);
#pragma warning restore CS0618
        Assert.Equal(["ID", "Name", "Age"], dt.Columns.Cast<DataColumn>().Select(x => x.ColumnName));
    }

    /// <summary>
    /// SaveAsByTemplate if there is & in the cell value, it will be &amp;
    /// https://gitee.com/dotnetchina/MiniExcel/issues/I4DQUN
    /// </summary>
    [Fact]
    public void TestIssueI4DQUN()
    {
        var templatePath = PathHelper.GetFile("xlsx/TestIssueI4DQUN.xlsx");
        using var path = AutoDeletingPath.Create();
        var value = new Dictionary<string, object>
        {
            { "Title", "Hello & World < , > , \" , '" },
            { "Details", new[] { new { Value = "Hello & Value < , > , \" , '" } } },
        };
        MiniExcel.SaveAsByTemplate(path.ToString(), templatePath, value);

        var sheetXml = Helpers.GetZipFileContent(path.ToString(), "xl/worksheets/sheet1.xml");
        Assert.Contains("<v>Hello &amp; World &lt; , &gt; , \" , '</v>", sheetXml);
        Assert.Contains("<v>Hello &amp; Value &lt; , &gt; , \" , '</v>", sheetXml);
    }

    /// <summary>
    /// [SaveAs default theme support filter mode · Issue #190 · mini-software/MiniExcel](https://github.com/mini-software/MiniExcel/issues/190)
    /// </summary>
    [Fact]
    public void TestIssue190()
    {
        {
            using var path = AutoDeletingPath.Create();
            var value = new TestIssue190Dto[] { };
            MiniExcel.SaveAs(path.ToString(), value, configuration: new OpenXmlConfiguration { AutoFilter = false });

            var sheetXml = Helpers.GetZipFileContent(path.ToString(), "xl/worksheets/sheet1.xml");
            Assert.DoesNotContain("<x:autoFilter ref=\"A1:C1\" />", sheetXml);
        }
        {
            using var path = AutoDeletingPath.Create();
            var value = new TestIssue190Dto[] { };
            MiniExcel.SaveAs(path.ToString(), value);

            var sheetXml = Helpers.GetZipFileContent(path.ToString(), "xl/worksheets/sheet1.xml");
            Assert.Contains("<x:autoFilter ref=\"A1:C1\" />", sheetXml);
        }
        {
            using var path = AutoDeletingPath.Create();
            TestIssue190Dto[] value =
            [
                new() { ID = 1, Name = "Jack", Age = 32 },
                new() { ID = 2, Name = "Lisa", Age = 45 }
            ];
            MiniExcel.SaveAs(path.ToString(), value);

            var sheetXml = Helpers.GetZipFileContent(path.ToString(), "xl/worksheets/sheet1.xml");
            Assert.Contains("<x:autoFilter ref=\"A1:C3\" />", sheetXml);
        }
    }

    private class TestIssue190Dto
    {
        public int ID { get; set; }
        public string Name { get; set; }
        public int Age { get; set; }
    }

    /// <summary>
    /// [According to the XLSX to CSV example, there will be data loss if there is no header. · Issue #292 · mini-software/MiniExcel](https://github.com/mini-software/MiniExcel/issues/292)
    /// </summary>
    [Fact]
    public void TestIssue292()
    {
        {
            var xlsxPath = PathHelper.GetFile("/xlsx/TestIssue292.xlsx");
            using var path = AutoDeletingPath.Create(ExcelType.CSV);
            MiniExcel.ConvertXlsxToCsv(xlsxPath, path.ToString());

            var actualCotent = File.ReadAllText(path.ToString());
            Assert.Equal(
                """
                Name,Age,Name,Age
                Jack,22,Mike,25
                Henry,44,Jerry,44

                """,
                actualCotent);
        }

        {
            var csvPath = PathHelper.GetFile("/csv/TestIssue292.csv");
            using var path = AutoDeletingPath.Create();
            MiniExcel.ConvertCsvToXlsx(csvPath, path.ToString());

            var rows = MiniExcel.Query(path.ToString()).ToList();
            Assert.Equal(3, rows.Count);
            Assert.Equal("Name", rows[0].A);
            Assert.Equal("Age", rows[0].B);
            Assert.Equal("Name", rows[0].C);
            Assert.Equal("Age", rows[0].D);
            Assert.Equal("Jack", rows[1].A);
            Assert.Equal("22", rows[1].B);
            Assert.Equal("Mike", rows[1].C);
            Assert.Equal("25", rows[1].D);
        }
    }

    /// <summary>
    /// [Csv Query then SaveAs will throw "Stream was not readable." exception · Issue #293 · mini-software/MiniExcel](https://github.com/mini-software/MiniExcel/issues/293)
    /// </summary>
    [Fact]
    public void TestIssue293()
    {
        var path = PathHelper.GetFile("/csv/Test5x2.csv");
        using var tempPath = AutoDeletingPath.Create();
        using var csv = File.OpenRead(path);
        var value = csv.Query(useHeaderRow: false, excelType: ExcelType.CSV);
        MiniExcel.SaveAs(tempPath.ToString(), value, printHeader: false, excelType: ExcelType.XLSX);
    }

    [Fact]
    public void TestIssueI49RYZ()
    {
        I49RYZDto[] values =
        [
            new() { Name="Jack", UserType=I49RYZUserType.V1 },
            new() { Name="Leo", UserType=I49RYZUserType.V2 },
            new() { Name="Henry", UserType=I49RYZUserType.V3 },
            new() { Name="Lisa", UserType=null }
        ];
        using var path = AutoDeletingPath.Create();
        MiniExcel.SaveAs(path.ToString(), values);
        var rows = MiniExcel.Query(path.ToString(), true).ToList();
        Assert.Equal("GeneralUser", rows[0].UserType);
        Assert.Equal("SuperAdministrator", rows[1].UserType);
        Assert.Equal("GeneralAdministrator", rows[2].UserType);
        Assert.Null(rows[3].UserType);
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
        MiniExcel.SaveAs(path.ToString(), values);
        var rows = MiniExcel.Query(path.ToString(), true).ToList();

        Assert.Equal("VIP1", rows[0].E);
        Assert.Equal("VIP2", rows[1].E);
    }

    private class TestIssue286Dto
    {
        public TestIssue286Enum E { get; set; }
    }

    private enum TestIssue286Enum { VIP1, VIP2 }

    private enum I49RYZUserType
    {
        [Description("GeneralUser")] V1 = 0,
        [Description("SuperAdministrator")] V2 = 1,
        [Description("GeneralAdministrator")] V3 = 2
    }

    private class I49RYZDto
    {
        public string Name { get; set; }
        public I49RYZUserType? UserType { get; set; }
    }


    /// <summary>
    /// Create Multiple Sheets from IDataReader have Bug #283
    /// </summary>
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
            var rows = MiniExcel.SaveAs(path.ToString(), sheets);
            Assert.Equal(2, rows.Length);
        }

        var sheetNames = MiniExcel.GetSheetNames(path.ToString());
        Assert.Equal(["sheet01", "sheet02"], sheetNames);
    }

    /// <summary>
    /// https://gitee.com/dotnetchina/MiniExcel/issues/I40QA5
    /// </summary>
    [Fact]
    public void TestIssueI40QA5()
    {
        {
            var path = PathHelper.GetFile("/xlsx/TestIssueI40QA5_1.xlsx");
            var rows = MiniExcel.Query<TestIssueI40QA5Dto>(path).ToList();
            Assert.Equal("E001", rows[0].Empno);
            Assert.Equal("E002", rows[1].Empno);
        }
        {
            var path = PathHelper.GetFile("/xlsx/TestIssueI40QA5_2.xlsx");
            var rows = MiniExcel.Query<TestIssueI40QA5Dto>(path).ToList();
            Assert.Equal("E001", rows[0].Empno);
            Assert.Equal("E002", rows[1].Empno);
        }
        {
            var path = PathHelper.GetFile("/xlsx/TestIssueI40QA5_3.xlsx");
            var rows = MiniExcel.Query<TestIssueI40QA5Dto>(path).ToList();
            Assert.Equal("E001", rows[0].Empno);
            Assert.Equal("E002", rows[1].Empno);
        }
        {
            var path = PathHelper.GetFile("/xlsx/TestIssueI40QA5_4.xlsx");
            var rows = MiniExcel.Query<TestIssueI40QA5Dto>(path).ToList();
            Assert.Null(rows[0].Empno);
            Assert.Null(rows[1].Empno);
        }
    }

    private class TestIssueI40QA5Dto
    {
        [ExcelColumnName(excelColumnName: "EmployeeNo", aliases: new[] { "EmpNo", "No" })]
        public string Empno { get; set; }
        public string Name { get; set; }
    }

    [Fact]
    public void TestIssues133()
    {
        {
            using var path = AutoDeletingPath.Create();

            var value = new DataTable();
            value.Columns.Add("Id");
            value.Columns.Add("Name");
            MiniExcel.SaveAs(path.ToString(), value);
            var rows = MiniExcel.Query(path.ToString()).ToList();

            Assert.Equal("Id", rows[0].A);
            Assert.Equal("Name", rows[0].B);
            Assert.Single(rows);
            Assert.Equal("A1:B1", Helpers.GetFirstSheetDimensionRefValue(path.ToString()));
        }

        {
            using var path = AutoDeletingPath.Create();

            var value = Array.Empty<TestIssues133Dto>();
            MiniExcel.SaveAs(path.ToString(), value);
            var rows = MiniExcel.Query(path.ToString()).ToList();

            Assert.Equal("Id", rows[0].A);
            Assert.Equal("Name", rows[0].B);
            Assert.Single(rows);
            Assert.Equal("A1:B1", Helpers.GetFirstSheetDimensionRefValue(path.ToString()));
        }
    }

    private class TestIssues133Dto
    {
        public string Id { get; set; }
        public string Name { get; set; }
    }

    /// <summary>
    /// Semicolon expected
    /// </summary>
    [Fact]
    public void TestIssueI45TF5_2()
    {
        {
            var value = new[] { new Dictionary<string, object> { { "Col1&Col2", "V1&V2" } } };
            var path = PathHelper.GetTempPath();
            MiniExcel.SaveAs(path, value);
            //System.Xml.XmlException : '<' is an unexpected token. The expected token is ';'.
            Helpers.GetZipFileContent(path, "xl/worksheets/sheet1.xml"); //check illegal format or not
        }

        {
            using var dt = new DataTable();
            dt.Columns.Add("Col1&Col2");
            dt.Rows.Add("V1&V2");
            var path = PathHelper.GetTempPath();
            MiniExcel.SaveAs(path, dt);
            //System.Xml.XmlException : '<' is an unexpected token. The expected token is ';'.
            Helpers.GetZipFileContent(path, "xl/worksheets/sheet1.xml"); //check illegal format or not
        }
    }

    [Fact]
    public void TestIssueI45TF5()
    {
        using var path = AutoDeletingPath.Create();
        MiniExcel.SaveAs(path.ToString(), new[] { new { C1 = "1&2;3,4", C2 = "1&2;3,4" } });
        var sheet1Xml = Helpers.GetZipFileContent(path.ToString(), "xl/worksheets/sheet1.xml");
        Assert.DoesNotContain("<x:cols>", sheet1Xml);
    }

    /// <summary>
    /// [Support column width attribute · Issue #280 · mini-software/MiniExcel](https://github.com/mini-software/MiniExcel/issues/280)
    /// </summary>
    [Fact]
    public void TestIssue280()
    {
        TestIssue280Dto[] value =
        [
            new() { ID = 1, Name = "Jack" },
            new() { ID = 2, Name = "Mike" }
        ];
        using var path = AutoDeletingPath.Create();
        MiniExcel.SaveAs(path.ToString(), value);
    }

    private class TestIssue280Dto
    {
        [ExcelColumnWidth(20)]
        public int ID { get; set; }
        [ExcelColumnWidth(15.50)]
        public string Name { get; set; }
    }

    /// <summary>
    /// Csv not support QueryAsDataTable #279 https://github.com/mini-software/MiniExcel/issues/279
    /// </summary>
    [Fact]
    public void TestIssue279()
    {
        var path = PathHelper.GetFile("/csv/TestHeader.csv");
#pragma warning disable CS0618 // Type or member is obsolete
        using var dt = MiniExcel.QueryAsDataTable(path, true, null, ExcelType.CSV);
#pragma warning restore CS0618
        Assert.Equal("A1", dt.Rows[0]["Column1"]);
        Assert.Equal("A2", dt.Rows[1]["Column1"]);
        Assert.Equal("B1", dt.Rows[0]["Column2"]);
        Assert.Equal("B2", dt.Rows[1]["Column2"]);
    }

    /// <summary>
    /// Custom excel zip can't read and show Number of entries expected in End Of Central Directory does not correspond to number of entries in Central Directory. #272
    /// </summary>
    [Fact]
    public void TestIssue272()
    {
        var path = PathHelper.GetFile("/xlsx/TestIssue272.xlsx");
        try
        {
            var rows = MiniExcel.Query(path).ToList();
        }
        catch (Exception e)
        {
            Assert.Equal(typeof(InvalidDataException), e.GetType());
            Assert.StartsWith("It's not legal excel zip, please check or issue for me.", e.Message);
        }
    }

    /// <summary>
    /// v0.16.0-0.17.1 custom format contains specific format (eg:`#,##0.000_);[Red]\(#,##0.000\)`), automatic converter will convert double to datetime #267
    /// </summary>
    [Fact]
    public void TestIssue267()
    {
        var path = PathHelper.GetFile("/xlsx/TestIssue267.xlsx");
        var row = MiniExcel.Query(path).SingleOrDefault();
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

    private static bool IsDateFormatString(string formatCode) => DateTimeHelper.IsDateTimeFormat(formatCode);

    [Fact]
    public void TestIssueI3X2ZL()
    {
        try
        {
            var path = PathHelper.GetFile("xlsx/TestIssueI3X2ZL_datetime_error.xlsx");
            var rows = MiniExcel.Query<IssueI3X2ZLDTO>(path, startCell: "B3").ToList();
        }
        catch (InvalidCastException ex)
        {
            Assert.Equal(
                "ColumnName: Col2, CellRow: 6, Value: error. The value cannot be cast to type DateTime.",
                ex.Message
            );
        }

        try
        {
            var path = PathHelper.GetFile("xlsx/TestIssueI3X2ZL_int_error.xlsx");
            var rows = MiniExcel.Query<IssueI3X2ZLDTO>(path).ToList();
        }
        catch (InvalidCastException ex)
        {
            Assert.Equal(
                "ColumnName: Col1, CellRow: 3, Value: error. The value cannot be cast to type Int32.",
                ex.Message
            );
        }
    }

    private class IssueI3X2ZLDTO
    {
        public int Col1 { get; set; }
        public DateTime Col2 { get; set; }
    }

    /// <summary>
    /// [Convert csv to xlsx · Issue #261 · mini-software/MiniExcel](https://github.com/mini-software/MiniExcel/issues/261)
    /// </summary>
    [Fact]
    public void TestIssue261()
    {
        var csvPath = PathHelper.GetFile("csv/TestCsvToXlsx.csv");
        using var path = AutoDeletingPath.Create();
        CsvToXlsx(csvPath, path.ToString());
        var rows = MiniExcel.Query(path.ToString()).ToList();

        Assert.Equal("Name", rows[0].A);
        Assert.Equal("Jack", rows[1].A);
        Assert.Equal("Neo", rows[2].A);
        Assert.Null(rows[3].A);
        Assert.Null(rows[4].A);
        Assert.Equal("Age", rows[0].B);
        Assert.Equal("34", rows[1].B);
        Assert.Equal("26", rows[2].B);
        Assert.Null(rows[3].B);
        Assert.Null(rows[4].B);
    }

    private static void CsvToXlsx(string csvPath, string xlsxPath)
    {
        var value = MiniExcel.Query(csvPath, true);
        MiniExcel.SaveAs(xlsxPath, value);
    }

    /// <summary>
    /// [SaveAsByTemplate support DateTime custom format · Issue #255 · mini-software/MiniExcel]
    /// (https://github.com/mini-software/MiniExcel/issues/255)
    /// </summary>
    [Fact]
    public void Issue255()
    {
        //template
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
            MiniExcel.SaveAsByTemplate(path.ToString(), templatePath, value);
            var rows = MiniExcel.Query(path.ToString()).ToList();
            Assert.Equal("2021", rows[1].A.ToString());
            Assert.Equal("2021", rows[1].B.ToString());
        }
        //saveas
        {
            using var path = AutoDeletingPath.Create();
            var value = new[]
            {
                new Issue255DTO { Time = new DateTime(2021, 01, 01) }
            };
            MiniExcel.SaveAs(path.ToString(), value);
            var rows = MiniExcel.Query(path.ToString()).ToList();
            Assert.Equal("2021", rows[1].A.ToString());
        }
    }

    private class Issue255DTO
    {
        [ExcelFormat("yyyy")]
        public DateTime Time { get; set; }

        [ExcelColumn(Format = "yyyy")]
        public DateTime Time2 { get; set; }
    }

    /// <summary>
    /// [Dynamic Query custom format not using mapping format · Issue #256]
    /// (https://github.com/mini-software/MiniExcel/issues/256)
    /// </summary>
    [Fact]
    public void Issue256()
    {
        var path = PathHelper.GetFile("xlsx/TestIssue256.xlsx");
        var rows = MiniExcel.Query(path, false).ToList();
        Assert.Equal(new DateTime(2003, 4, 16), rows[1].A);
        Assert.Equal(new DateTime(2004, 4, 16), rows[1].B);
    }


    /// <summary>
    /// Csv SaveAs by datareader with encoding default show messy code #253
    /// </summary>
    [Fact]
    public void Issue253()
    {
        {
            var value = new[] { new { col1 = "世界你好" } };
            using var path = AutoDeletingPath.Create(ExcelType.CSV);
            MiniExcel.SaveAs(path.ToString(), value);
            const string expected =
                """
                col1
                世界你好

                """;
            Assert.Equal(expected, File.ReadAllText(path.ToString()));
        }

        {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
            var value = new[] { new { col1 = "世界你好" } };
            using var path = AutoDeletingPath.Create(ExcelType.CSV);
            var config = new CsvConfiguration
            {
                StreamWriterFunc = stream => new StreamWriter(stream, Encoding.GetEncoding("gb2312"))
            };
            MiniExcel.SaveAs(path.ToString(), value, excelType: ExcelType.CSV, configuration: config);
            const string expected =
                """
                col1
                �������

                """;
            Assert.Equal(expected, File.ReadAllText(path.ToString()));
        }

        using var cn = Db.GetConnection();

        {
            var value = cn.ExecuteReader("select '世界你好' col1");
            using var path = AutoDeletingPath.Create(ExcelType.CSV);
            MiniExcel.SaveAs(path.ToString(), value);
            const string expected =
                """
                col1
                世界你好

                """;
            Assert.Equal(expected, File.ReadAllText(path.ToString()));
        }
    }

    /// <summary>
    /// [CSV SaveAs support datareader · Issue #251 · mini-software/MiniExcel](https://github.com/mini-software/MiniExcel/issues/251)
    /// </summary>
    [Fact]
    public void Issue251()
    {
        using var cn = Db.GetConnection();
        using var reader = cn.ExecuteReader(@"select '""<>+-*//}{\\n' a,1234567890 b union all select '<test>Hello World</test>',-1234567890");
        using var path = AutoDeletingPath.Create(ExcelType.CSV);
        MiniExcel.SaveAs(path.ToString(), reader);
        const string expected =
            """"
            a,b
            """<>+-*//}{\\n",1234567890
            "<test>Hello World</test>",-1234567890

            """";

        Assert.Equal(expected, File.ReadAllText(path.ToString()));
    }

    /// <summary>
    /// No error exception throw when reading xls file #242
    /// </summary>
    [Fact]
    public void Issue242()
    {
        var path = PathHelper.GetFile("xls/TestIssue242.xls");
        Assert.Throws<NotSupportedException>(() => MiniExcel.Query(path).ToList());

        using var stream = File.OpenRead(path);
        Assert.Throws<InvalidDataException>(() => stream.Query().ToList());
    }

    /// <summary>
    /// Csv type mapping Query error "cannot be converted to xxx type" #243
    /// </summary>
    [Fact]
    public void Issue243()
    {
        using var path = AutoDeletingPath.Create(ExcelType.CSV);
        var value = new[]
        {
            new { Name = "Jack", Age = 25, InDate = new DateTime(2021,01,03) },
            new { Name = "Henry",  Age = 36, InDate = new DateTime(2020,05,03) },
        };
        MiniExcel.SaveAs(path.ToString(), value);

        var rows = MiniExcel.Query<Issue243Dto>(path.ToString()).ToList();
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
    public void Issue241()
    {

        Issue241Dto[] value =
        [
            new() { Name = "Jack", InDate = new DateTime(2021,01,04) },
            new() { Name = "Henry", InDate = new DateTime(2020,04,05) }
        ];

        // csv
        {
            using var path = AutoDeletingPath.Create(ExcelType.CSV);
            MiniExcel.SaveAs(path.ToString(), value);

            {
                var rows = MiniExcel.Query(path.ToString(), true).ToList();
                Assert.Equal(rows[0].InDate, "01 04, 2021");
                Assert.Equal(rows[1].InDate, "04 05, 2020");
            }

            {
                var rows = MiniExcel.Query<Issue241Dto>(path.ToString()).ToList();
                Assert.Equal(rows[0].InDate, new DateTime(2021, 01, 04));
                Assert.Equal(rows[1].InDate, new DateTime(2020, 04, 05));
            }
        }

        // xlsx
        {
            using var path = AutoDeletingPath.Create();
            MiniExcel.SaveAs(path.ToString(), value);

            {
                var rows = MiniExcel.Query(path.ToString(), true).ToList();
                Assert.Equal(rows[0].InDate, "01 04, 2021");
                Assert.Equal(rows[1].InDate, "04 05, 2020");
            }

            {
                var rows = MiniExcel.Query<Issue241Dto>(path.ToString()).ToList();
                Assert.Equal(rows[0].InDate, new DateTime(2021, 01, 04));
                Assert.Equal(rows[1].InDate, new DateTime(2020, 04, 05));
            }
        }
    }

    private class Issue241Dto
    {
        public string Name { get; set; }

        [ExcelFormat("MM dd, yyyy")]
        public DateTime InDate { get; set; }
    }

    /// <summary>
    /// SaveAs Default Template #132
    /// </summary>
    [Fact]
    public void Issue132()
    {
        {
            using var path = AutoDeletingPath.Create();
            var value = new[] {
                new { name = "Jack", Age = 25, InDate = new DateTime(2021,01,03)},
                new { name = "Henry", Age = 36, InDate = new DateTime(2020,05,03)},
            };

            MiniExcel.SaveAs(path.ToString(), value);
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
            MiniExcel.SaveAs(path.ToString(), value, configuration: config);
        }

        {
            using var path = AutoDeletingPath.Create();
            var value = JsonConvert.DeserializeObject<DataTable>(
                JsonConvert.SerializeObject(new[]
                {
                    new { name = "Jack", Age = 25, InDate = new DateTime(2021,01,03)},
                    new { name = "Henry", Age = 36, InDate = new DateTime(2020,05,03)},
                })
            );
            MiniExcel.SaveAs(path.ToString(), value);
        }
    }

    /// <summary>
    /// Support SaveAs by DataSet #235
    /// </summary>
    [Fact]
    public void Issue235()
    {
        using var path = AutoDeletingPath.Create();

        DataSet dataSet = new();
        var users = JsonConvert.DeserializeObject<DataTable>(JsonConvert.SerializeObject(new[]
        {
            new { Name = "Jack", Age = 25 },
            new { Name = "Mike", Age = 44 }
        }));
        users!.TableName = "users";

        var department = JsonConvert.DeserializeObject<DataTable>(JsonConvert.SerializeObject(new[]
        {
            new { ID = "01", Name = "HR" },
            new { ID = "02", Name = "IT" }
        }));
        department!.TableName = "department";

        dataSet.Tables.Add(users);
        dataSet.Tables.Add(department);

        var rowsWritten = MiniExcel.SaveAs(path.ToString(), dataSet);
        Assert.Equal(2, rowsWritten.Length);
        Assert.Equal(2, rowsWritten[0]);

        var sheetNames = MiniExcel.GetSheetNames(path.ToString());
        Assert.Equal("users", sheetNames[0]);
        Assert.Equal("department", sheetNames[1]);

        {
            var rows = MiniExcel.Query(path.ToString(), true, sheetName: "users").ToList();
            Assert.Equal("Jack", rows[0].Name);
            Assert.Equal(25, rows[0].Age);
            Assert.Equal("Mike", rows[1].Name);
            Assert.Equal(44, rows[1].Age);
        }
        {
            var rows = MiniExcel.Query(path.ToString(), true, sheetName: "department").ToList();
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
    public void Issue233()
    {
        var path = PathHelper.GetFile("xlsx/TestIssue233.xlsx");
#pragma warning disable CS0618 // Type or member is obsolete
        using var dt = MiniExcel.QueryAsDataTable(path);
#pragma warning restore CS0618
        var rows = dt.Rows;

        Assert.Equal(0.55, rows[0]["Size"]);
        Assert.Equal("0.55/1.1", rows[1]["Size"]);
    }

    /// <summary>
    /// Csv Query split comma not correct #237
    /// https://github.com/mini-software/MiniExcel/issues/237
    /// </summary>
    [Fact]
    public void Issue237()
    {
        var value = new[]
        {
            new{ id="\"\"1,2,3\"\""},
            new{ id="1,2,3"},
        };
        using var path = AutoDeletingPath.Create(ExcelType.CSV);
        MiniExcel.SaveAs(path.ToString(), value);

        var rows = MiniExcel.Query(path.ToString(), true).ToList();

        Assert.Equal("\"\"1,2,3\"\"", rows[0].id);
        Assert.Equal("1,2,3", rows[1].id);
    }

    /// <summary>
    /// SaveAs support multiple sheets #234
    /// </summary>
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
        MiniExcel.SaveAs(path, sheets);

        var sheetNames = MiniExcel.GetSheetNames(path);
        Assert.Equal("users", sheetNames[0]);
        Assert.Equal("department", sheetNames[1]);

        {
            var rows = MiniExcel.Query(path, true, sheetName: "users").ToList();
            Assert.Equal("Jack", rows[0].Name);
            Assert.Equal(25, rows[0].Age);
            Assert.Equal("Mike", rows[1].Name);
            Assert.Equal(44, rows[1].Age);
        }
        {
            var rows = MiniExcel.Query(path, true, sheetName: "department").ToList();
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
            MiniExcel.SaveAs(path.ToString(), reader, printHeader: true);
            var rows = MiniExcel.Query(path.ToString(), true).ToList();
            Assert.Equal(1, rows[0].id);
            Assert.Equal(2, rows[1].id);
        }
    }

    /// <summary>
    /// v0.14.3 QueryAsDataTable error "Cannot set Column to be null" #229
    /// https://github.com/mini-software/MiniExcel/issues/229
    /// </summary>
    [Fact]
    public void Issue229()
    {
        var path = PathHelper.GetFile("xlsx/TestIssue229.xlsx");
#pragma warning disable CS0618 // Type or member is obsolete
        using var dt = MiniExcel.QueryAsDataTable(path);
#pragma warning restore CS0618
        foreach (DataColumn column in dt.Columns)
        {
            var v = dt.Rows[3][column];
            Assert.Equal(DBNull.Value, v);
        }
    }

    /// <summary>
    /// [Query Merge cells data · Issue #122 · mini-software/MiniExcel]
    /// (https://github.com/mini-software/MiniExcel/issues/122)
    /// </summary>
    [Fact]
    public void Issue122()
    {
        var config = new OpenXmlConfiguration
        {
            FillMergedCells = true
        };
        {
            var path = PathHelper.GetFile("xlsx/TestIssue122.xlsx");
            {
                var rows = MiniExcel.Query(path, useHeaderRow: true, configuration: config).ToList();
                Assert.Equal("HR", rows[0].Department);
                Assert.Equal("HR", rows[1].Department);
                Assert.Equal("HR", rows[2].Department);
                Assert.Equal("IT", rows[3].Department);
                Assert.Equal("IT", rows[4].Department);
                Assert.Equal("IT", rows[5].Department);
            }
        }

        {
            var path = PathHelper.GetFile("xlsx/TestIssue122_2.xlsx");
            {
                var rows = MiniExcel.Query(path, useHeaderRow: true, configuration: config).ToList();
                Assert.Equal("V1", rows[2].Test1);
                Assert.Equal("V2", rows[5].Test2);
                Assert.Equal("V3", rows[1].Test3);
                Assert.Equal("V4", rows[2].Test4);
                Assert.Equal("V5", rows[3].Test5);
                Assert.Equal("V6", rows[5].Test5);
            }
        }
    }

    /// <summary>
    /// [Support Xlsm AutoCheck · Issue #227 · mini-software/MiniExcel]
    /// (https://github.com/mini-software/MiniExcel/issues/227)
    /// </summary>
    [Fact]
    public void Issue227()
    {
        {
            var path = PathHelper.GetTempPath("xlsm");
            Assert.Throws<NotSupportedException>(() => MiniExcel.SaveAs(path, new[] { new { V = "A1" }, new { V = "A2" } }));
            File.Delete(path);
        }

        {
            var path = PathHelper.GetFile("xlsx/TestIssue227.xlsm");
            {
                var rows = MiniExcel.Query<UserAccount>(path).ToList();
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
                using var stream = File.OpenRead(path);
                var rows = stream.Query<UserAccount>().ToList();
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
    /// Fix SaveAsByTemplate single column dimension index error #226
    /// </summary>
    [Fact]
    public void Issue226()
    {
        using var path = AutoDeletingPath.Create();
        var templatePath = PathHelper.GetFile("xlsx/TestIssue226.xlsx");
        MiniExcel.SaveAsByTemplate(path.ToString(), templatePath, new { employees = new[] { new { name = "123" }, new { name = "123" } } });
        Assert.Equal("A1:A3", Helpers.GetFirstSheetDimensionRefValue(path.ToString()));
    }

    /// <summary>
    /// ASP.NET Webform gridview datasource can't use miniexcel queryasdatatable · Issue #223]
    /// (https://github.com/mini-software/MiniExcel/issues/223)
    /// </summary>
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
        MiniExcel.SaveAs(path.ToString(), value);

#pragma warning disable CS0618 // Type or member is obsolete
        using var dt = MiniExcel.QueryAsDataTable(path.ToString());
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
    public void Issue222()
    {
        var path = PathHelper.GetFile("xlsx/TestIssue222.xlsx");
        var rows = MiniExcel.Query(path).ToList();
        Assert.Equal(typeof(DateTime), rows[1].A.GetType());
        Assert.Equal(new DateTime(2021, 4, 29), rows[1].A);
    }

    /// <summary>
    /// Query Support StartCell #147
    /// https://github.com/mini-software/MiniExcel/issues/147
    /// </summary>
    [Fact]
    public void Issue147()
    {
        {
            var path = PathHelper.GetFile("xlsx/TestIssue147.xlsx");
            var rows = MiniExcel.Query(path, useHeaderRow: false, startCell: "C3", sheetName: "Sheet1").ToList();

            Assert.Equal(["C", "D", "E"], (rows[0] as IDictionary<string, object>)?.Keys);
            Assert.Equal(["Column1", "Column2", "Column3"], new[] { rows[0].C as string, rows[0].D as string, rows[0].E as string });
            Assert.Equal(["C4", "D4", "E4"], new[] { rows[1].C as string, rows[1].D as string, rows[1].E as string });
            Assert.Equal(["C9", "D9", "E9"], new[] { rows[6].C as string, rows[6].D as string, rows[6].E as string });
            Assert.Equal(["C12", "D12", "E12"], new[] { rows[9].C as string, rows[9].D as string, rows[9].E as string });
            Assert.Equal(["C13", "D13", "E13"], new[] { rows[10].C as string, rows[10].D as string, rows[10].E as string });
            foreach (var i in new[] { 4, 5, 7, 8 })
                Assert.Equal(expected: [null, null, null], new[] { rows[i].C as string, rows[i].D as string, rows[i].E as string });

            Assert.Equal(11, rows.Count);


            var columns = MiniExcel.GetColumns(path, startCell: "C3");
            Assert.Equal(["C", "D", "E"], columns);
        }

        {
            var path = PathHelper.GetFile("xlsx/TestIssue147.xlsx");
            var rows = MiniExcel.Query(path, useHeaderRow: true, startCell: "C3", sheetName: "Sheet1").ToList();

            Assert.Equal(["Column1", "Column2", "Column3"], (rows[0] as IDictionary<string, object>)?.Keys);
            Assert.Equal(["C4", "D4", "E4"], new[] { rows[0].Column1 as string, rows[0].Column2 as string, rows[0].Column3 as string });
            Assert.Equal(["C9", "D9", "E9"], new[] { rows[5].Column1 as string, rows[5].Column2 as string, rows[5].Column3 as string });
            Assert.Equal(["C12", "D12", "E12"], new[] { rows[8].Column1 as string, rows[8].Column2 as string, rows[8].Column3 as string });
            Assert.Equal(["C13", "D13", "E13"], new[] { rows[9].Column1 as string, rows[9].Column2 as string, rows[9].Column3 as string });
            foreach (var i in new[] { 3, 4, 6, 7 })
                Assert.Equal([null, null, null], new[] { rows[i].Column1 as string, rows[i].Column2 as string, rows[i].Column3 as string });

            Assert.Equal(10, rows.Count);

            var columns = MiniExcel.GetColumns(path, useHeaderRow: true, startCell: "C3");
            Assert.Equal(["Column1", "Column2", "Column3"], columns);
        }
    }


    /// <summary>
    /// [Can SaveAs support iDataReader export to avoid the dataTable consuming too much memory · Issue #211 · mini-software/MiniExcel]
    /// (https://github.com/mini-software/MiniExcel/issues/211)
    /// </summary>
    [Fact]
    public void Issue211()
    {
        using var path = AutoDeletingPath.Create();
        var tempSqlitePath = AutoDeletingPath.Create(Path.GetTempPath(), $"{Guid.NewGuid()}.db");
        var connectionString = $"Data Source={tempSqlitePath};Version=3;";

        using var connection = new SQLiteConnection(connectionString);
        var reader = connection.ExecuteReader(@"select 1 Test1,2 Test2 union all select 3 , 4 union all select 5 ,6");
        MiniExcel.SaveAs(path.ToString(), reader);
        var rows = MiniExcel.Query(path.ToString(), true).ToList();

        Assert.Equal(1.0, rows[0].Test1);
        Assert.Equal(2.0, rows[0].Test2);
        Assert.Equal(3.0, rows[1].Test1);
        Assert.Equal(4.0, rows[1].Test2);
    }

    /// <summary>
    /// [When reading Excel, can return IDataReader and DataTable to facilitate the import of database. Like ExcelDataReader provide reader.AsDataSet() · Issue #216 · mini-software/MiniExcel](https://github.com/mini-software/MiniExcel/issues/216)
    /// </summary>
    [Fact]
    public void Issue216()
    {
        using var path = AutoDeletingPath.Create();
        var value = new[]
        {
            new { Test1 = "1", Test2 = 2 },
            new { Test1 = "3", Test2 = 4 }
        };
        MiniExcel.SaveAs(path.ToString(), value);

        {
#pragma warning disable CS0618 // Type or member is obsolete
            using var table = MiniExcel.QueryAsDataTable(path.ToString());
#pragma warning restore CS0618
            Assert.Equal("Test1", table.Columns[0].ColumnName);
            Assert.Equal("Test2", table.Columns[1].ColumnName);
            Assert.Equal("1", table.Rows[0]["Test1"]);
            Assert.Equal(2.0, table.Rows[0]["Test2"]);
            Assert.Equal("3", table.Rows[1]["Test1"]);
            Assert.Equal(4.0, table.Rows[1]["Test2"]);
        }

        {
#pragma warning disable CS0618 // Type or member is obsolete
            using var dt = MiniExcel.QueryAsDataTable(path.ToString(), false);
#pragma warning restore CS0618
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
    public void IssueI3OSKV()
    {
        {
            using var path = AutoDeletingPath.Create();
            var value = new[] { new { Test = "12345678901234567890" } };
            MiniExcel.SaveAs(path.ToString(), value);

            var A2 = MiniExcel.Query(path.ToString(), true).First().Test;
            Assert.Equal("12345678901234567890", A2);
        }

        {
            using var path = AutoDeletingPath.Create();
            var value = new[] { new { Test = 123456.789 } };
            MiniExcel.SaveAs(path.ToString(), value);

            var A2 = MiniExcel.Query(path.ToString(), true).First().Test;
            Assert.Equal(123456.789, A2);
        }
    }


    /// <summary>
    /// [Dynamic Query can't summary numeric cell value default, need to cast · Issue #220 · mini-software/MiniExcel]
    /// (https://github.com/mini-software/MiniExcel/issues/220)
    /// </summary>
    [Fact]
    public void Issue220()
    {
        var path = PathHelper.GetFile("xlsx/TestIssue220.xlsx");
        var rows = MiniExcel.Query(path, useHeaderRow: true);
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

    /// <summary>
    /// Optimize stream excel type check
    /// https://github.com/mini-software/MiniExcel/issues/215
    /// </summary>
    [Fact]
    public void Issue215()
    {
        using var stream = new MemoryStream();
        stream.SaveAs(new[] { new { V = "test1" }, new { V = "test2" } });
        var rows = stream.Query(true).ToList();

        Assert.Equal("test1", rows[0].V);
        Assert.Equal("test2", rows[1].V);
    }

    /// <summary>
    /// Support Enum Mapping
    /// https://github.com/mini-software/MiniExcel/issues/89
    /// </summary>
    [Fact]
    public void Issue89()
    {
        //csv
        {
            const string text =
                """
               State
               OnDuty
               Fired
               Leave
               """;

            using var stream = new MemoryStream();
            using var writer = new StreamWriter(stream);

            writer.Write(text);
            writer.Flush();
            stream.Position = 0;
            var rows = stream.Query<Issue89VO>(excelType: ExcelType.CSV).ToList();

            Assert.Equal(Issue89VO.WorkState.OnDuty, rows[0].State);
            Assert.Equal(Issue89VO.WorkState.Fired, rows[1].State);
            Assert.Equal(Issue89VO.WorkState.Leave, rows[2].State);

            using var path = AutoDeletingPath.Create(ExcelType.CSV);
            MiniExcel.SaveAs(path.ToString(), rows);
            var rows2 = MiniExcel.Query<Issue89VO>(path.ToString()).ToList();

            Assert.Equal(Issue89VO.WorkState.OnDuty, rows2[0].State);
            Assert.Equal(Issue89VO.WorkState.Fired, rows2[1].State);
            Assert.Equal(Issue89VO.WorkState.Leave, rows2[2].State);
        }

        //xlsx
        {
            var path = PathHelper.GetFile("xlsx/TestIssue89.xlsx");
            var rows = MiniExcel.Query<Issue89VO>(path).ToList();

            Assert.Equal(Issue89VO.WorkState.OnDuty, rows[0].State);
            Assert.Equal(Issue89VO.WorkState.Fired, rows[1].State);
            Assert.Equal(Issue89VO.WorkState.Leave, rows[2].State);

            using var xlsxPath = AutoDeletingPath.Create();
            MiniExcel.SaveAs(xlsxPath.ToString(), rows);
            var rows2 = MiniExcel.Query<Issue89VO>(xlsxPath.ToString()).ToList();

            Assert.Equal(Issue89VO.WorkState.OnDuty, rows2[0].State);
            Assert.Equal(Issue89VO.WorkState.Fired, rows2[1].State);
            Assert.Equal(Issue89VO.WorkState.Leave, rows2[2].State);
        }
    }

    public class Issue89VO
    {
        public WorkState State { get; set; }

        public enum WorkState
        {
            OnDuty,
            Leave,
            Fired
        }
    }

    /// <summary>
    /// DataTable recommended to use Caption for column name first, then use columname
    /// https://github.com/mini-software/MiniExcel/issues/217
    /// </summary>
    [Fact]
    public void Issue217()
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
            MiniExcel.SaveAs(path.ToString(), table);

            var rows = MiniExcel.Query(path.ToString()).ToList();
            Assert.Equal("Name", rows[0].B);
            Assert.Equal("Limit", rows[0].C);
        }

        // csv
        {
            using var path = AutoDeletingPath.Create(ExcelType.CSV);
            MiniExcel.SaveAs(path.ToString(), table);

            var rows = MiniExcel.Query(path.ToString()).ToList();
            Assert.Equal("Name", rows[0].B);
            Assert.Equal("Limit", rows[0].C);
        }
    }

    /// <summary>
    /// MiniExcel.SaveAs(path, table,sheetName:“Name”) ，the actual sheetName is Sheet1
    /// https://github.com/mini-software/MiniExcel/issues/212
    /// </summary>
    [Fact]
    public void Issue212()
    {
        const string sheetName = "Demo";
        using var path = AutoDeletingPath.Create();
        MiniExcel.SaveAs(path.ToString(), new[] { new { x = 1, y = 2 } }, sheetName: sheetName);

        var actualSheetName = MiniExcel.GetSheetNames(path.ToString()).ToList()[0];
        Assert.Equal(sheetName, actualSheetName);
    }

    /// <summary>
    /// Version &lt;= v0.13.1 Template merge row list rendering has no merge
    /// https://github.com/mini-software/MiniExcel/issues/207
    /// </summary>
    [Fact]
    public void Issue207()
    {
        {
            const string tempaltePath = "../../../../../samples/xlsx/TestIssue207_2.xlsx";
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

            MiniExcel.SaveAsByTemplate(path, tempaltePath, value);
            var rows = MiniExcel.Query(path).ToList();

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

            var dimension = Helpers.GetFirstSheetDimensionRefValue(path);
            Assert.Equal("A1:C16", dimension);
        }

        {
            const string tempaltePath = "../../../../../samples/xlsx/TestIssue207_Template_Merge_row_list_rendering_without_merge/template.xlsx";
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

            MiniExcel.SaveAsByTemplate(path, tempaltePath, value);
            var rows = MiniExcel.Query(path).ToList();

            Assert.Equal("項目1", rows[0].A);
            Assert.Equal("[]內容1,[]內容2,[]內容3,[]內容4,[]內容5", rows[0].C);
            Assert.Equal("項目2", rows[3].A);
            Assert.Equal("[]內容1,[]內容2,[]內容3,[]內容4,[]內容5", rows[3].C);
            Assert.Equal("項目3", rows[6].A);
            Assert.Equal("[]內容1,[]內容2,[]內容3,[]內容4,[]內容5", rows[6].C);
            Assert.Equal("項目4", rows[9].A);
            Assert.Equal("[]內容1,[]內容2,[]內容3,[]內容4,[]內容5", rows[9].C);
            var dimension = Helpers.GetFirstSheetDimensionRefValue(path);
            Assert.Equal("A1:E15", dimension);
        }
    }

    /// <summary>
    /// https://github.com/mini-software/MiniExcel/issues/87
    /// </summary>
    [Fact]
    public void Issue87()
    {
        const string templatePath = "../../../../../samples/xlsx/TestTemplateCenterEmpty.xlsx";
        using var path = AutoDeletingPath.Create();
        var value = new
        {
            Tests = Enumerable.Range(1, 5).Select((_, i) => new { test1 = i, test2 = i })
        };

        var rows = MiniExcel.Query(templatePath).ToList();
        MiniExcel.SaveAsByTemplate(path.ToString(), templatePath, value);
    }

    /// <summary>
    /// https://github.com/mini-software/MiniExcel/issues/208
    /// </summary>
    [Fact]
    public void Issue208()
    {
        const string path = "../../../../../samples/xlsx/TestIssue208.xlsx";
        var columns = MiniExcel.GetColumns(path).ToList();
        Assert.Equal(16384, columns.Count);
        Assert.Equal("XFD", columns[16383]);
    }

    /// <summary>
    /// https://github.com/mini-software/MiniExcel/issues/206
    /// </summary>
    [Fact]
    public void Issue206()
    {
        {
            const string templatePath = "../../../../../samples/xlsx/TestTemplateBasicIEmumerableFill.xlsx";
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
            MiniExcel.SaveAsByTemplate(path.ToString(), templatePath, value);

            var dimension = Helpers.GetFirstSheetDimensionRefValue(path.ToString());
            Assert.Equal("A1:B2", dimension);
        }

        {
            const string templatePath = "../../../../../samples/xlsx/TestTemplateBasicIEmumerableFill.xlsx";
            using var path = AutoDeletingPath.Create();

            using var dt = new DataTable();
            dt.Columns.Add("name");
            dt.Columns.Add("department");
            dt.Rows.Add("Jack", "HR");

            var value = new Dictionary<string, object> { ["employees"] = dt };
            MiniExcel.SaveAsByTemplate(path.ToString(), templatePath, value);

            var dimension = Helpers.GetFirstSheetDimensionRefValue(path.ToString());
            Assert.Equal("A1:B2", dimension);
        }
    }


    /// <summary>
    /// https://github.com/mini-software/MiniExcel/issues/193
    /// </summary>
    [Fact]
    public void Issue193()
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
            MiniExcel.SaveAsByTemplate(path, templatePath, value);

            foreach (var sheetName in MiniExcel.GetSheetNames(path))
            {
                var rows = MiniExcel.Query(path, sheetName: sheetName).ToList();
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

                /*TODO:row can't contain xmlns*/
                // https://user-images.githubusercontent.com/12729184/114998840-ead44500-9ed3-11eb-8611-58afb98faed9.png

            }
        }

        {
            const string templatePath = "../../../../../samples/xlsx/TestTemplateComplex.xlsx";
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
            MiniExcel.SaveAsByTemplate(path.ToString(), templatePath, value);
            var rows = MiniExcel.Query(path.ToString()).ToList();

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
    public void Issue142()
    {
        {
            using var file = AutoDeletingPath.Create();
            var path = file.ToString();
            Issue142VO[] values =
            [
                new()
                {
                    MyProperty1 = "MyProperty1", MyProperty2 = "MyProperty2", MyProperty3 = "MyProperty3",
                    MyProperty4 = "MyProperty4", MyProperty5 = "MyProperty5", MyProperty6 = "MyProperty6",
                    MyProperty7 = "MyProperty7"
                }
            ];
            var rowsWritten = MiniExcel.SaveAs(path, values);
            Assert.Single(rowsWritten);
            Assert.Equal(1, rowsWritten[0]);

            {
                var rows = MiniExcel.Query(path).ToList();

                Assert.Equal("MyProperty4", rows[0].A);
                Assert.Equal("CustomColumnName", rows[0].B);
                Assert.Equal("MyProperty5", rows[0].C);
                Assert.Equal("MyProperty2", rows[0].D);
                Assert.Equal("MyProperty6", rows[0].E);
                Assert.Equal(null, rows[0].F);
                Assert.Equal("MyProperty3", rows[0].G);

                Assert.Equal("MyProperty4", rows[0].A);
                Assert.Equal("CustomColumnName", rows[0].B);
                Assert.Equal("MyProperty5", rows[0].C);
                Assert.Equal("MyProperty2", rows[0].D);
                Assert.Equal("MyProperty6", rows[0].E);
                Assert.Equal(null, rows[0].F);
                Assert.Equal("MyProperty3", rows[0].G);
            }

            {
                var rows = MiniExcel.Query<Issue142VO>(path).ToList();

                Assert.Equal("MyProperty4", rows[0].MyProperty4);
                Assert.Equal("MyProperty1", rows[0].MyProperty1);
                Assert.Equal("MyProperty5", rows[0].MyProperty5);
                Assert.Equal("MyProperty2", rows[0].MyProperty2);
                Assert.Equal("MyProperty6", rows[0].MyProperty6);
                Assert.Null(rows[0].MyProperty7);
                Assert.Equal("MyProperty3", rows[0].MyProperty3);
            }
        }

        {
            using var file = AutoDeletingPath.Create(ExcelType.CSV);
            var path = file.ToString();
            Issue142VO[] values =
            [
                new()
                {
                    MyProperty1 = "MyProperty1", MyProperty2 = "MyProperty2", MyProperty3 = "MyProperty3",
                    MyProperty4 = "MyProperty4", MyProperty5 = "MyProperty5", MyProperty6 = "MyProperty6",
                    MyProperty7 = "MyProperty7"
                }
            ];
            var rowsWritten = MiniExcel.SaveAs(path, values);
            Assert.Single(rowsWritten);
            Assert.Equal(1, rowsWritten[0]);

            const string expected =
                """
                MyProperty4,CustomColumnName,MyProperty5,MyProperty2,MyProperty6,,MyProperty3
                MyProperty4,MyProperty1,MyProperty5,MyProperty2,MyProperty6,,MyProperty3

                """;

            Assert.Equal(expected, File.ReadAllText(path));

            {
                var rows = MiniExcel.Query<Issue142VO>(path).ToList();

                Assert.Equal("MyProperty4", rows[0].MyProperty4);
                Assert.Equal("MyProperty1", rows[0].MyProperty1);
                Assert.Equal("MyProperty5", rows[0].MyProperty5);
                Assert.Equal("MyProperty2", rows[0].MyProperty2);
                Assert.Equal("MyProperty6", rows[0].MyProperty6);
                Assert.Null(rows[0].MyProperty7);
                Assert.Equal("MyProperty3", rows[0].MyProperty3);
            }
        }

        {
            using var path = AutoDeletingPath.Create(ExcelType.CSV);
            Issue142VoDuplicateColumnName[] input =
            [
                new() { MyProperty1 = 0, MyProperty2 = 0, MyProperty3 = 0, MyProperty4 = 0 }
            ];
            Assert.Throws<InvalidOperationException>(() => MiniExcel.SaveAs(path.ToString(), input));
        }
    }

    [Fact]
    public void Issue142_Query()
    {
        const string path = "../../../../../samples/xlsx/TestIssue142.xlsx";
        const string csvPath = "../../../../../samples/csv/TestIssue142.csv";
        {
            var rows = MiniExcel.Query<Issue142VoExcelColumnNameNotFound>(path).ToList();
            Assert.Equal(0, rows[0].MyProperty1);
        }

        Assert.Throws<ArgumentException>(() => MiniExcel.Query<Issue142VoOverIndex>(path).ToList());

        var rowsXlsx = MiniExcel.Query<Issue142VO>(path).ToList();
        Assert.Equal("CustomColumnName", rowsXlsx[0].MyProperty1);
        Assert.Null(rowsXlsx[0].MyProperty7);
        Assert.Equal("MyProperty2", rowsXlsx[0].MyProperty2);
        Assert.Equal("MyProperty103", rowsXlsx[0].MyProperty3);
        Assert.Equal("MyProperty100", rowsXlsx[0].MyProperty4);
        Assert.Equal("MyProperty102", rowsXlsx[0].MyProperty5);
        Assert.Equal("MyProperty6", rowsXlsx[0].MyProperty6);

        var rowsCsv = MiniExcel.Query<Issue142VO>(csvPath).ToList();
        Assert.Equal("CustomColumnName", rowsCsv[0].MyProperty1);
        Assert.Null(rowsCsv[0].MyProperty7);
        Assert.Equal("MyProperty2", rowsCsv[0].MyProperty2);
        Assert.Equal("MyProperty103", rowsCsv[0].MyProperty3);
        Assert.Equal("MyProperty100", rowsCsv[0].MyProperty4);
        Assert.Equal("MyProperty102", rowsCsv[0].MyProperty5);
        Assert.Equal("MyProperty6", rowsCsv[0].MyProperty6);
    }

    private class Issue142VO
    {
        [ExcelColumnName("CustomColumnName")]
        public string MyProperty1 { get; set; }  //index = 1
        [ExcelIgnore]
        public string MyProperty7 { get; set; } //index = null
        public string MyProperty2 { get; set; } //index = 3
        [ExcelColumnIndex(6)]
        public string MyProperty3 { get; set; } //index = 6
        [ExcelColumnIndex("A")] // equal column index 0
        public string MyProperty4 { get; set; }
        [ExcelColumnIndex(2)]
        public string MyProperty5 { get; set; } //index = 2
        public string MyProperty6 { get; set; } //index = 4
    }

    private class Issue142VoDuplicateColumnName
    {
        [ExcelColumnIndex("A")]
        public int MyProperty1 { get; set; }
        [ExcelColumnIndex("A")]
        public int MyProperty2 { get; set; }

        public int MyProperty3 { get; set; }
        [ExcelColumnIndex("B")]
        public int MyProperty4 { get; set; }
    }

    private class Issue142VoOverIndex
    {
        [ExcelColumnIndex("Z")]
        public int MyProperty1 { get; set; }
    }

    private class Issue142VoExcelColumnNameNotFound
    {
        [ExcelColumnIndex("B")]
        public int MyProperty1 { get; set; }
    }

    /// <summary>
    /// https://github.com/mini-software/MiniExcel/issues/150
    /// </summary>
    [Fact]
    public void Issue150()
    {
        var path = Path.Combine(Path.GetTempPath(), $"{Guid.NewGuid()}.xlsx");

        Assert.Throws<NotSupportedException>(() => MiniExcel.SaveAs(path, new[] { 1, 2 }));
        File.Delete(path);

        Assert.Throws<NotSupportedException>(() => MiniExcel.SaveAs(path, new[] { "1", "2" }));
        File.Delete(path);

        Assert.Throws<NotSupportedException>(() => MiniExcel.SaveAs(path, new[] { '1', '2' }));
        File.Delete(path);

        Assert.Throws<NotSupportedException>(() => MiniExcel.SaveAs(path, new[] { DateTime.Now }));
        File.Delete(path);

        Assert.Throws<NotSupportedException>(() => MiniExcel.SaveAs(path, new[] { Guid.NewGuid() }));
        File.Delete(path);
    }

    /// <summary>
    /// https://github.com/mini-software/MiniExcel/issues/157
    /// </summary>
    [Fact]
    public void Issue157()
    {
        {
            using var file = AutoDeletingPath.Create();
            var path = file.ToString();
            _output.WriteLine("==== SaveAs by strongly type ====");

            var input = JsonConvert.DeserializeObject<IEnumerable<UserAccount>>(
                """
                [
                  {
                    "ID":"78de23d2-dcb6-bd3d-ec67-c112bbc322a2",
                    "Name":"Wade","BoD":"2020-09-27T00:00:00",
                    "Age":5019,"VIP":false,"Points":5019.12,
                    "IgnoredProperty":null
                  },
                  {
                    "ID":"20d3bfce-27c3-ad3e-4f70-35c81c7e8e45",
                    "Name":"Felix","BoD":"2020-10-25T00:00:00",
                    "Age":7028,"VIP":true,"Points":7028.46,
                    "IgnoredProperty":null
                  },
                  {
                    "ID":"52013bf0-9aeb-48e6-e5f5-e9500afb034f",
                    "Name":"Phelan","BoD":"2021-10-04T00:00:00",
                    "Age":3836,"VIP":true,"Points":3835.7,
                    "IgnoredProperty":null
                  },
                  {
                    "ID":"3b97b87c-7afe-664f-1af5-6914d313ae25",
                    "Name":"Samuel","BoD":"2020-06-21T00:00:00",
                    "Age":9352,"VIP":false,"Points":9351.71,
                    "IgnoredProperty":null
                  },
                  {
                    "ID":"9a989c43-d55f-5306-0d2f-0fbafae135bb",
                    "Name":"Raymond","BoD":"2021-07-12T00:00:00",
                    "Age":8210,"VIP":true,"Points":8209.76,
                    "IgnoredProperty":null
                  }
                ]
                """);
            MiniExcel.SaveAs(path, input);

            var rows = MiniExcel.Query(path, sheetName: "Sheet1").ToList();
            Assert.Equal(6, rows.Count);
            Assert.Equal("Sheet1", MiniExcel.GetSheetNames(path).First());

            using var p = new ExcelPackage(new FileInfo(path));
            var ws = p.Workbook.Worksheets.First();
            Assert.Equal("Sheet1", ws.Name);
            Assert.Equal("Sheet1", p.Workbook.Worksheets["Sheet1"].Name);
        }
        {
            const string path = "../../../../../samples/xlsx/TestIssue157.xlsx";

            {
                var rows = MiniExcel.Query(path, sheetName: "Sheet1").ToList();
                Assert.Equal(6, rows.Count);
                Assert.Equal("Sheet1", MiniExcel.GetSheetNames(path).First());
            }
            using (var p = new ExcelPackage(new FileInfo(path)))
            {
                var ws = p.Workbook.Worksheets.First();
                Assert.Equal("Sheet1", ws.Name);
                Assert.Equal("Sheet1", p.Workbook.Worksheets["Sheet1"].Name);
            }

            {
                var rows = MiniExcel.Query<UserAccount>(path, sheetName: "Sheet1").ToList();
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
            const string path = "../../../../../samples/xlsx/TestIssue149.xlsx";
            var rows = MiniExcel.Query(path).Select(s => (string)s.A).ToList();
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
            MiniExcel.SaveAs(path.ToString(), input);

            var rows = MiniExcel.Query(path.ToString(), true).Select(s => (string)s.Test).ToList();
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
            MiniExcel.SaveAs(path.ToString(), input);

            var rows = MiniExcel.Query<Issue149VO>(path.ToString()).Select(s => s.Test).ToList();
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
    public void Issue153()
    {
        const string path = "../../../../../samples/xlsx/TestIssue153.xlsx";
        var rows = MiniExcel.Query(path, true).First() as IDictionary<string, object>;
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
    public void Issue137()
    {
        const string path = "../../../../../samples/xlsx/TestIssue137.xlsx";

        {
            var rows = MiniExcel.Query(path).ToList();
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
            var rows = MiniExcel.Query(path, true).ToList();
            var first = rows[0] as IDictionary<string, object>; //![image](https://user-images.githubusercontent.com/12729184/113266322-ba06e400-9307-11eb-9521-d36abfda75cc.png)
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
            var rows = MiniExcel.Query<Issue137ExcelRow>(path).ToList();
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
    public void Issue138()
    {
        const string path = "../../../../../samples/xlsx/TestIssue138.xlsx";
        {
            var rows = MiniExcel.Query(path, true).ToList();
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

            var rows = MiniExcel.Query<Issue138ExcelRow>(path).ToList();
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

    /// <summary>
    /// https://gitee.com/dotnetchina/MiniExcel/issues/I50VD5
    /// </summary>
    [Fact]
    public void IssueI50VD5()
    {
        using var file = AutoDeletingPath.Create();
        var path = file.ToString();

        List<dynamic> list1 =
        [
            new { Name = "github", Image = File.ReadAllBytes(PathHelper.GetFile("images/github_logo.png")) },
            new { Name = "google", Image = File.ReadAllBytes(PathHelper.GetFile("images/google_logo.png")) },
            new { Name = "microsoft", Image = File.ReadAllBytes(PathHelper.GetFile("images/microsoft_logo.png")) },
            new { Name = "reddit", Image = File.ReadAllBytes(PathHelper.GetFile("images/reddit_logo.png")) },
            new { Name = "statck_overflow", Image = File.ReadAllBytes(PathHelper.GetFile("images/statck_overflow_logo.png")) }
        ];

        List<dynamic> list2 =
        [
            new { Id = 1, Name = "github", Image = File.ReadAllBytes(PathHelper.GetFile("images/github_logo.png")) },
            new { Id = 2, Name = "google", Image = File.ReadAllBytes(PathHelper.GetFile("images/google_logo.png")) }
        ];

        var sheets = new Dictionary<string, object>
        {
            ["A"] = list1,
            ["B"] = list2,
        };
        MiniExcel.SaveAs(path, sheets);

        {
            Assert.Contains("/xl/media/", Helpers.GetZipFileContent(path, "xl/drawings/_rels/drawing1.xml.rels"));
            Assert.Contains("ext cx=\"609600\" cy=\"190500\"", Helpers.GetZipFileContent(path, "xl/drawings/drawing1.xml"));
            Assert.Contains("/xl/drawings/drawing1.xml", Helpers.GetZipFileContent(path, "[Content_Types].xml"));
            Assert.Contains("drawing r:id=\"drawing1\"", Helpers.GetZipFileContent(path, "xl/worksheets/sheet1.xml"));
            Assert.Contains("../drawings/drawing1.xml", Helpers.GetZipFileContent(path, "xl/worksheets/_rels/sheet1.xml.rels"));

            Assert.Contains("/xl/media/", Helpers.GetZipFileContent(path, "xl/drawings/_rels/drawing2.xml.rels"));
            Assert.Contains("ext cx=\"609600\" cy=\"190500\"", Helpers.GetZipFileContent(path, "xl/drawings/drawing2.xml"));
            Assert.Contains("/xl/drawings/drawing1.xml", Helpers.GetZipFileContent(path, "[Content_Types].xml"));
            Assert.Contains("drawing r:id=\"drawing2\"", Helpers.GetZipFileContent(path, "xl/worksheets/sheet2.xml"));
            Assert.Contains("../drawings/drawing2.xml", Helpers.GetZipFileContent(path, "xl/worksheets/_rels/sheet2.xml.rels"));
        }
    }

    private class Issue422Enumerable(IEnumerable inner) : IEnumerable
    {
        private readonly IEnumerable _inner = inner;
        public int GetEnumeratorCount { get; private set; }

        public IEnumerator GetEnumerator()
        {
            GetEnumeratorCount++;
            return _inner.GetEnumerator();
        }
    }

    /// <summary>
    /// https://github.com/mini-software/MiniExcel/issues/422
    /// </summary>
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
        MiniExcel.SaveAs(path.ToString(), enumerableWithCount);
        Assert.Equal(1, enumerableWithCount.GetEnumeratorCount);
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

        ms.SaveAsByTemplate(template, values);
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
        MiniExcel.SaveAsByTemplate(path.FilePath, template, value);

        var rows = MiniExcel.Query(path.FilePath).ToList();
        Assert.Equal("General User", rows[1].B);
        Assert.Equal("General Administrator", rows[2].B);
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
        MiniExcel.SaveAs(path.FilePath, reader, configuration: excelconfig, overwriteFile: true);

        var rows = MiniExcel.Query(path.FilePath).ToList();
        Assert.All(rows, x => Assert.Single(x));
        Assert.Equal("Name", rows[0].A);
    }

    private class Issue585VO1
    {
        public string Col1 { get; set; }
        public string Col2 { get; set; }
        public string Col3 { get; set; }
    }

    private class Issue585VO2
    {
        public string Col1 { get; set; }

        [ExcelColumnName("Col2")]
        public string Prop2 { get; set; }

        public string Col3 { get; set; }
    }

    private class Issue585VO3
    {
        public string Col1 { get; set; }

        [ExcelColumnIndex("B")]
        public string Prop2 { get; set; }

        public string Col3 { get; set; }
    }

    [Fact]
    public void Issue585()
    {
        const string path = "../../../../../samples/xlsx/TestIssue585.xlsx";

        var items1 = MiniExcel.Query<Issue585VO1>(path);
        Assert.Equal(2, items1.Count());

        var items2 = MiniExcel.Query<Issue585VO2>(path);
        Assert.Equal(2, items2.Count());

        var items3 = MiniExcel.Query<Issue585VO3>(path);
        Assert.Equal(2, items3.Count());
    }

    private class Issue542
    {
        [ExcelColumnIndex(0)] public Guid ID { get; set; }
        [ExcelColumnIndex(1)] public string Name { get; set; }
    }

    [Fact]
    public void Issue_542()
    {
        const string path = "../../../../../samples/xlsx/TestIssue542.xlsx";

        var resultWithoutFirstRow = MiniExcel.Query<Issue542>(path).ToList();
        var resultWithFirstRow = MiniExcel.Query<Issue542>(path, hasHeader: false).ToList();

        Assert.Equal(15, resultWithoutFirstRow.Count);
        Assert.Equal(16, resultWithFirstRow.Count);

        Assert.Equal("Felix", resultWithoutFirstRow[0].Name);
        Assert.Equal("Wade", resultWithFirstRow[0].Name);
    }

    private class Issue507V01
    {
        public string A { get; set; }
        public DateTime B { get; set; }
        public string C { get; set; }
        public int D { get; set; }
    }

    private class Issue507V02
    {
        public DateTime B { get; set; }
        public int D { get; set; }
    }

    [Fact]
    public void Issue507_1()
    {
        //Problem with multi-line when using Query func
        //https://github.com/mini-software/MiniExcel/issues/507

        var path = Path.Combine(Path.GetTempPath(), string.Concat(nameof(MiniExcelIssueTests), "_", nameof(Issue507_1), ".csv"));
        var values = new Issue507V01[]
        {
            new() { A = "Github", B = DateTime.Parse("2021-01-01"), C = "abcd", D = 123 },
            new() { A = "Microsoft \nTest 1", B = DateTime.Parse("2021-02-01"), C = "efgh", D = 123 },
            new() { A = "Microsoft \rTest 2", B = DateTime.Parse("2021-02-01"), C = "ab\nc\nd", D = 123 },
            new() { A = "Microsoft\"\" \r\nTest\n3", B = DateTime.Parse("2021-02-01"), C = "a\"\"\nb\n\nc", D = 123 },
        };

        var config = new CsvConfiguration
        {
            //AlwaysQuote = true,
            ReadLineBreaksWithinQuotes = true,
        };

        // create
        using (var stream = File.Create(path))
        {
            stream.SaveAs(values, excelType: ExcelType.CSV, configuration: config);
        }

        // read
        var getRowsInfo = MiniExcel.Query<Issue507V01>(path, excelType: ExcelType.CSV, configuration: config).ToArray();

        Assert.Equal(values.Length, getRowsInfo.Length);

        Assert.Equal("Github", getRowsInfo[0].A);
        Assert.Equal("abcd", getRowsInfo[0].C);

        Assert.Equal($"Microsoft {config.NewLine}Test 1", getRowsInfo[1].A);
        Assert.Equal("efgh", getRowsInfo[1].C);

        Assert.Equal($"Microsoft {config.NewLine}Test 2", getRowsInfo[2].A);
        Assert.Equal($"ab{config.NewLine}c{config.NewLine}d", getRowsInfo[2].C);

        Assert.Equal($"""Microsoft"" {config.NewLine}Test{config.NewLine}3""", getRowsInfo[3].A);
        Assert.Equal($"""a""{config.NewLine}b{config.NewLine}{config.NewLine}c""", getRowsInfo[3].C);

        File.Delete(path);
    }

    [Fact]
    public void Issue507_2()
    {
        //Problem with multi-line when using Query func
        //https://github.com/mini-software/MiniExcel/issues/507

        var path = Path.Combine(Path.GetTempPath(), string.Concat(nameof(MiniExcelIssueTests), "_", nameof(Issue507_2), ".csv"));
        var values = new Issue507V02[]
        {
            new() { B = DateTime.Parse("2021-01-01"), D = 123 },
            new() { B = DateTime.Parse("2021-02-01"), D = 123 },
            new() { B = DateTime.Parse("2021-02-01"), D = 123 },
            new() { B = DateTime.Parse("2021-02-01"), D = 123 },
        };

        var config = new CsvConfiguration
        {
            //AlwaysQuote = true,
            ReadLineBreaksWithinQuotes = true,
        };

        // create
        using (var stream = File.Create(path))
        {
            stream.SaveAs(values, excelType: ExcelType.CSV, configuration: config);
        }

        // read
        var getRowsInfo = MiniExcel.Query<Issue507V02>(path, excelType: ExcelType.CSV, configuration: config).ToArray();
        Assert.Equal(values.Length, getRowsInfo.Length);

        File.Delete(path);
    }

    [Fact]
    public void Issue507_3_MismatchedQuoteCsv()
    {
        //Problem with multi-line when using Query func
        //https://github.com/mini-software/MiniExcel/issues/507

        var config = new CsvConfiguration
        {
            //AlwaysQuote = true,
            ReadLineBreaksWithinQuotes = true,
        };

        // create
        using var stream = new MemoryStream(Encoding.UTF8.GetBytes("A,B,C\n\"r1a: no end quote,r1b,r1c"));

        // read
        var getRowsInfo = stream.Query(excelType: ExcelType.CSV, configuration: config).ToArray();
        Assert.Equal(2, getRowsInfo.Length);
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
            string.Concat(nameof(MiniExcelIssueTests), "_", nameof(Issue606_1), ".xlsx")
        );

        const string templateFileName = "../../../../../samples/xlsx/TestIssue606_Template.xlsx";
        MiniExcel.SaveAsByTemplate(path, Path.GetFullPath(templateFileName), value);
        File.Delete(path);
    }

    /// <summary>
    /// https://github.com/mini-software/MiniExcel/issues/627
    /// </summary>
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
        var rowsWritten = MiniExcel.SaveAs(path.ToString(), value, configuration: config);

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
            string.Concat(nameof(MiniExcelIssueTests), "_", nameof(Issue632_1), ".xlsx")
        );

        MiniExcel.SaveAs(path, values, excelType: ExcelType.XLSX, configuration: config, overwriteFile: true);
        File.Delete(path);
    }

    private class Issue658TestData
    {
        public string FirstName { get; set; }
        public string LastName { get; set; }
    }

    /// <summary>
    /// https://github.com/mini-software/MiniExcel/issues/658
    /// </summary>
    [Fact]
    public void Issue_658()
    {
        static IEnumerable<Issue658TestData> GetTestData()
        {
            yield return new() { FirstName = "Unit", LastName = "Test" };
            yield return new() { FirstName = "Unit1", LastName = "Test1" };
            yield return new() { FirstName = "Unit2", LastName = "Test2" };
        }

        using var memoryStream = new MemoryStream();
        var testData = GetTestData();
        var rowsWritten = memoryStream.SaveAs(testData, configuration: new OpenXmlConfiguration
        {
            FastMode = true
        });
        Assert.Single(rowsWritten);
        Assert.Equal(3, rowsWritten[0]);

        memoryStream.Position = 0;

        var queryData = memoryStream.Query<Issue658TestData>().ToList();

        Assert.Equal(testData.Count(), queryData.Count);

        var i = 0;
        foreach (var data in testData)
        {
            Assert.Equal(data.FirstName, queryData[i].FirstName);
            Assert.Equal(data.LastName, queryData[i].LastName);
            i++;
        }
    }

    /// <summary>
    /// https://github.com/mini-software/MiniExcel/issues/658
    /// </summary>
    /// <returns></returns>
    [Fact]
    public async Task Issue_658_async()
    {
        static IEnumerable<Issue658TestData> GetTestData()
        {
            yield return new() { FirstName = "Unit", LastName = "Test" };
            yield return new() { FirstName = "Unit1", LastName = "Test1" };
            yield return new() { FirstName = "Unit2", LastName = "Test2" };
        }

        using var memoryStream = new MemoryStream();
        var testData = GetTestData();
        await memoryStream.SaveAsAsync(testData, configuration: new OpenXmlConfiguration
        {
            FastMode = true,
        });

        memoryStream.Position = 0;

        var queryData = (await memoryStream.QueryAsync<Issue658TestData>()).ToList();

        Assert.Equal(testData.Count(), queryData.Count);

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
            MiniExcel.QueryRange(path, useHeaderRow: false, startCell: "ZZFF10", endCell: "ZZFF11").First());

        Assert.Throws<InvalidDataException>(() =>
            MiniExcel.QueryRange(path, useHeaderRow: false, startCell: "ZZFF@@10", endCell: "ZZFF@@11").First());
    }

    [Fact]
    public void Test_Issue_693_SaveSheetWithLongName()
    {
        using var path1 = AutoDeletingPath.Create();
        using var path2 = AutoDeletingPath.Create();

        List<Dictionary<string, object>> data = [new() { ["First"] = 1, ["Second"] = 2 }];
        Assert.Throws<ArgumentException>(() => MiniExcel.SaveAs(path1.ToString(), data, sheetName: "Some Really Looooooooooong Sheet Name"));
        MiniExcel.SaveAs(path2.ToString(), new List<Dictionary<string, object>>());
        Assert.Throws<ArgumentException>(() => MiniExcel.Insert(path2.ToString(), data, sheetName: "Some Other Very Looooooong Sheet Name"));
    }

    private class Issue697
    {
        public int First { get; set; }
        public int Second { get; set; }
        public int Third { get; set; }
        public int Fourth { get; set; }
    }
    [Fact]
    public void Test_Issue_697_EmptyRowsStronglyTypedQuery()
    {
        const string path = "../../../../../samples/xlsx/TestIssue697.xlsx";
        var rowsIgnoreEmpty = MiniExcel.Query<Issue697>(path, configuration: new OpenXmlConfiguration { IgnoreEmptyRows = true }).ToList();
        var rowsCountEmpty = MiniExcel.Query<Issue697>(path).ToList();
        Assert.Equal(4, rowsIgnoreEmpty.Count);
        Assert.Equal(5, rowsCountEmpty.Count);
    }

    [Fact]
    public void Issue_710()
    {
        var values = new[] { new { Column1 = "MiniExcel", Column2 = 1, Column3 = "Test" } };
        using var memoryStream = new MemoryStream();
        memoryStream.SaveAs(values, configuration: new OpenXmlConfiguration
        {
            FastMode = true
        });

        memoryStream.Position = 0;
        using var dataReader = memoryStream.GetReader(useHeaderRow: false);

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
        const string path1 = "../../../../../samples/xlsx/TestIssue732_1.xlsx";
        const string path2 = "../../../../../samples/xlsx/TestIssue732_2.xlsx";
        const string path3 = "../../../../../samples/xlsx/TestIssue732_3.xlsx";

        var info1 = MiniExcel.GetSheetInformations(path1);
        var info2 = MiniExcel.GetSheetInformations(path2);
        var info3 = MiniExcel.GetSheetInformations(path3);

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
        MiniExcel.SaveAsByTemplate(path.ToString(), templatePath, data);

        var rows = MiniExcel.Query(path.ToString())
            .Skip(1453)
            .Take(2)
            .ToList();

        Assert.True(((string)rows[0].A).Length > 9);

        var memoryAfter = GC.GetTotalMemory(true);
        var memoryIncrease = memoryAfter - memoryBefore;

        _output.WriteLine($"memoryIncrease: {memoryIncrease}");
    }


    /// <summary>
    /// https://github.com/mini-software/MiniExcel/issues/751
    /// Optimize CleanXml method #751
    /// </summary>
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
        MiniExcel.SaveAsByTemplate(path.ToString(), templatePath, data);

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

    /// <summary>
    /// https://github.com/mini-software/MiniExcel/issues/763
    /// Optimize CleanXml method #751
    /// </summary>
    [Fact]
    public void TestIssue763()
    {
        var path = PathHelper.GetFile("xlsx/TestIssue763.xlsx");
        var rows = MiniExcel.QueryRange(path, startCell: "A3", endCell: "J3").ToArray();
        Assert.Equal("A3", rows[0].A);
        Assert.Equal(null, rows[0].J);
    }

    /// <summary>
    /// https://github.com/mini-software/MiniExcel/issues/768
    /// Optimize CleanXml method #751
    /// </summary>
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

        MiniExcel.SaveAsByTemplate(path.ToString(), templatePath, data);
        var rows = MiniExcel.Query(path.ToString(), startCell: "A16").ToList();

        Assert.Equal(list[0].value1.ToString(), rows[0].A.ToString());
        Assert.Equal(list[1].value1.ToString(), rows[1].A.ToString());
    }

    /// <summary>
    /// https://github.com/mini-software/MiniExcel/issues/186
    /// </summary>
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

        MiniExcel.AddPicture(path.FilePath, images);
    }

    /// <summary>
    /// https://github.com/mini-software/MiniExcel/issues/771
    /// </summary>
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

        MiniExcel.SaveAsByTemplate(path.FilePath, template, value);
        var rows = MiniExcel.Query(path.FilePath).ToList();

        Assert.Equal("2025-1", rows[2].B);
        Assert.Equal(null, rows[3].B);
        Assert.Equal(null, rows[4].B);
        Assert.Equal("2025-2", rows[5].B);
        return;

        IEnumerable<object> GetEnumerable() => Enumerable.Range(0, 3).Select(s => new { ID = Guid.NewGuid(), level = s });
    }

    /// <summary>
    /// https://github.com/mini-software/MiniExcel/issues/772
    /// </summary>
    [Fact]
    public void TestIssue772()
    {
        var path = PathHelper.GetFile("xlsx/TestIssue772.xlsx");
        var rows = MiniExcel.Query(path, sheetName: "Supply plan(daily)", startCell: "A1")
            .Cast<IDictionary<string, object>>()
            .ToArray();

        Assert.Equal("01108083-1Delta", (string)rows[19]["C"]);
    }

    /// <summary>
    /// https://github.com/mini-software/MiniExcel/issues/773
    /// </summary>
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

        MiniExcel.SaveAsByTemplate(path.FilePath, templatePath, fill);
        var rows = MiniExcel.Query(path.FilePath).ToList();

        Assert.Equal("H1", rows[4].AF);
        Assert.Equal("c3", rows[6].AA);
        Assert.Equal("Ram", rows[6].B);
    }

    /// <summary>
    /// https://github.com/mini-software/MiniExcel/issues/789
    /// </summary>
    [Fact]
    public void TestIssue789()
    {
        var path = PathHelper.GetTempPath();
        var value = new[] {
            new Dictionary<string, object> { {"no","1"} },
            new Dictionary<string, object> { {"no","2"} },
            new Dictionary<string, object> { {"no","3"} },
        };
        MiniExcel.SaveAs(path, value);

        var xml = Helpers.GetZipFileContent(path.ToString(), "xl/worksheets/sheet1.xml");

        Assert.Contains("<x:autoFilter ref=\"A1:A4\" />", xml);
    }
    
    /// <summary>
    /// https://github.com/mini-software/MiniExcel/issues/814
    /// </summary>
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

        MiniExcel.AddPicture(path.FilePath, images);

        using var package = new ExcelPackage(new FileInfo(path.FilePath));

        // Check picture in the first sheet (C3)  
        var firstSheet = package.Workbook.Worksheets[0];
        var pictureInC3 = firstSheet.Drawings.OfType<ExcelPicture>().FirstOrDefault(p => p.From.Column == 2 && p.From.Row == 2);
        Assert.NotNull(pictureInC3);

        // Check picture in the "Demo" sheet (C9)  
        var demoSheet = package.Workbook.Worksheets["Demo"];
        var pictureInC9 = demoSheet.Drawings.OfType<ExcelPicture>().FirstOrDefault(p => p.From.Column == 2 && p.From.Row == 8);
        Assert.NotNull(pictureInC9);
    }

    /// <summary>
    /// https://github.com/mini-software/MiniExcel/issues/815
    /// </summary>
    [Fact]
    public void TestIssue815()
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

            MiniExcel.AddPicture(path.FilePath, images);

            using (var package = new ExcelPackage(new FileInfo(path.FilePath)))
            {
                // Check picture in the first sheet (C3)  
                var firstSheet = package.Workbook.Worksheets[0];
                var pictureInC3 = firstSheet.Drawings.OfType<ExcelPicture>().FirstOrDefault(p => p.From.Column == 2 && p.From.Row == 2);
                Assert.NotNull(pictureInC3);

                // Check picture in the "Demo" sheet (C9)  
                var demoSheet = package.Workbook.Worksheets["Demo"];
                var pictureInC9 = demoSheet.Drawings.OfType<ExcelPicture>().FirstOrDefault(p => p.From.Column == 2 && p.From.Row == 8);
                Assert.NotNull(pictureInC9);
            }
        }

        // TODO:check C3 image WidthPx = 80px, HeightPx = 24px, C9 WidthPx=500,HeightPx=500 
    }

    /// <summary>
    /// https://github.com/mini-software/MiniExcel/issues/816
    /// </summary>
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

            MiniExcel.AddPicture(path.FilePath, images);

            using (var package = new ExcelPackage(new FileInfo(path.FilePath)))
            {
                // Check picture in the first sheet (C3)  
                var firstSheet = package.Workbook.Worksheets[0];
                var pictureInC3 = firstSheet.Drawings
                    .OfType<ExcelPicture>()
                    .FirstOrDefault(p => p.From.Column == 2 && p.From.Row == 2);
                
                Assert.NotNull(pictureInC3);

                // Check picture in the "Demo" sheet (C9)  
                var demoSheet = package.Workbook.Worksheets["Demo"];
                var pictureInC9 = demoSheet.Drawings
                    .OfType<ExcelPicture>()
                    .FirstOrDefault(p => p.From.Column == 2 && p.From.Row == 8);
                
                Assert.NotNull(pictureInC9);
            }
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

            MiniExcel.AddPicture(path.FilePath, images);

            using (var package = new ExcelPackage(new FileInfo(path.FilePath)))
            {
                {
                    // Check picture in the first sheet (C3)  
                    var firstSheet = package.Workbook.Worksheets[0];
                    var pictureInC3 = firstSheet.Drawings.
                        OfType<ExcelPicture>()
                        .FirstOrDefault(p => p.From.Column == 2 && p.From.Row == 2);
                    
                    Assert.NotNull(pictureInC3);

                    // Check picture in the "Demo" sheet (C9)  
                    var demoSheet = package.Workbook.Worksheets["Demo"];
                    var pictureInC9 = demoSheet.Drawings
                        .OfType<ExcelPicture>()
                        .FirstOrDefault(p => p.From.Column == 2 && p.From.Row == 8);
                    
                    Assert.NotNull(pictureInC9);
                }

                {
                    // Check picture in the first sheet (D3)
                    var firstSheet = package.Workbook.Worksheets[0];
                    var pictureInD3 = firstSheet.Drawings
                        .OfType<ExcelPicture>()
                        .FirstOrDefault(p => p.From.Column == 3 && p.From.Row == 2);
                    
                    Assert.NotNull(pictureInD3);

                    // Check picture in the "Demo" sheet (D9)
                    var demoSheet = package.Workbook.Worksheets["Demo"];
                    var pictureInD9 = demoSheet.Drawings
                        .OfType<ExcelPicture>()
                        .FirstOrDefault(p => p.From.Column == 3 && p.From.Row == 8);
                    
                    Assert.NotNull(pictureInD9);

                    // Check picture in the "Demo" sheet (F9)
                    var pictureInF9 = demoSheet.Drawings
                        .OfType<ExcelPicture>()
                        .FirstOrDefault(p => p.From.Column == 5 && p.From.Row == 8);
                    
                    Assert.NotNull(pictureInF9);
                }
            }

        }
    }

    // https://github.com/mini-software/MiniExcel/issues/809
    [Fact]
    public void TestIssue809()
    {
        var path = PathHelper.GetFile("xlsx/TestIssue809.xlsx");
        var rows = MiniExcel.Query(path).ToList();

        Assert.Equal(3, rows.Count);
        Assert.Equal(null, rows[0].A);
        Assert.Equal(2, rows[2].B);
    }
    
    
    private class Issue869
    {
        public string? Name { get; set; }
        public DateOnly? Date { get; set; }
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

        var testFn = () => MiniExcel.Query<Issue869>(path, configuration: config).ToList();
        if (throwsException)
        {
            Assert.Throws<ExcelInvalidCastException>(testFn);
        }
        else
        {
            try
            {
                var result = testFn();
                Assert.Equal(new DateOnly(2025, 1, 1), result[0].Date);
            }
            catch (Exception ex)
            {
                Assert.Fail($"No exception should be thrown, but one was still thrown: {ex}.");
            }
        }
    }
}