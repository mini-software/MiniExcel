using CsvHelper;
using MiniExcelLibs.Tests.Utils;
using System.Data;
using System.Globalization;
using System.Text;
using Xunit;

namespace MiniExcelLibs.Tests;

public class MiniExcelCsvAsycTests
{
    [Fact]
    public async Task Gb2312_Encoding_Read_Test()
    {
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
        var path = PathHelper.GetFile("csv/gb2312_Encoding_Read_Test.csv");
        var config = new Csv.CsvConfiguration
        {
            StreamReaderFunc = stream => new StreamReader(stream, encoding: Encoding.GetEncoding("gb2312"))
        };
        var q = await MiniExcel.QueryAsync(path, true, excelType: ExcelType.CSV, configuration: config);
        var rows = q.ToList();
        Assert.Equal("世界你好", rows[0].栏位1);
    }

    [Fact]
    public async Task SeperatorTest()
    {
        using var file = AutoDeletingPath.Create(ExcelType.CSV);
        var path = file.ToString();

        List<Dictionary<string, object>> values =
        [
            new()
            {
                { "a", @"""<>+-*//}{\\n" },
                { "b", 1234567890 },
                { "c", true },
                { "d", new DateTime(2021, 1, 1) }
            },

            new()
            {
                { "a", "<test>Hello World</test>" },
                { "b", -1234567890 },
                { "c", false },
                { "d", new DateTime(2021, 1, 2) }
            }
        ];
            
        var rowsWritten = await MiniExcel.SaveAsAsync(path, values, configuration: new Csv.CsvConfiguration { Seperator = ';' });
        Assert.Equal(2, rowsWritten[0]);
            
        const string expected =
            """"
            a;b;c;d
            """<>+-*//}{\\n";1234567890;True;"2021-01-01 00:00:00"
            "<test>Hello World</test>";-1234567890;False;"2021-01-02 00:00:00"

            """";
        
        Assert.Equal(expected, await File.ReadAllTextAsync(path));
    }

    [Fact]
    public async Task SaveAsByDictionary()
    {
        {
            using var file = AutoDeletingPath.Create(ExcelType.CSV);
            var path = file.ToString();

            var table = new List<Dictionary<string, object>>();
            await MiniExcel.SaveAsAsync(path, table);
            Assert.Equal("\r\n", await File.ReadAllTextAsync(path));
        }

        {
            using var file = AutoDeletingPath.Create(ExcelType.CSV);
            var path = file.ToString();

            var table = new Dictionary<string, object>(); //TODO
            Assert.Throws<NotSupportedException>(() => MiniExcel.SaveAs(path, table));
        }

        {
            using var file = AutoDeletingPath.Create(ExcelType.CSV);
            var path = file.ToString();

            List<Dictionary<string, object>> values =
            [
                new()
                {
                    { "a", @"""<>+-*//}{\\n" },
                    { "b", 1234567890 },
                    { "c", true },
                    { "d", new DateTime(2021, 1, 1) }
                },
                new()
                {
                    { "a", "<test>Hello World</test>" },
                    { "b", -1234567890 },
                    { "c", false },
                    { "d", new DateTime(2021, 1, 2) }
                }

            ];
            var rowsWritten = await MiniExcel.SaveAsAsync(path, values);
            Assert.Equal(2, rowsWritten[0]);

            using (var reader = new StreamReader(path))
            using (var csv = new CsvReader(reader, CultureInfo.InvariantCulture))
            {
                var records = csv.GetRecords<dynamic>().ToList();
                Assert.Equal(@"""<>+-*//}{\\n", records[0].a);
                Assert.Equal("1234567890", records[0].b);
                Assert.Equal("True", records[0].c);
                Assert.Equal("2021-01-01 00:00:00", records[0].d);

                Assert.Equal("<test>Hello World</test>", records[1].a);
                Assert.Equal("-1234567890", records[1].b);
                Assert.Equal("False", records[1].c);
                Assert.Equal("2021-01-02 00:00:00", records[1].d);
            }
        }

        {
            using var file = AutoDeletingPath.Create(ExcelType.CSV);
            var path = file.ToString();
            
            List<Dictionary<int, object>> values =
            [
                new()
                {
                    { 1, @"""<>+-*//}{\\n" },
                    { 2, 1234567890 },
                    { 3, true },
                    { 4, new DateTime(2021, 1, 1) }
                },

                new()
                {
                    { 1, "<test>Hello World</test>" },
                    { 2, -1234567890 },
                    { 3, false },
                    { 4, new DateTime(2021, 1, 2) }
                }
            ];
            var rowsWritten = await MiniExcel.SaveAsAsync(path, values);
            Assert.Equal(2, rowsWritten[0]);

            using (var reader = new StreamReader(path))
            using (var csv = new CsvReader(reader, CultureInfo.InvariantCulture))
            {
                var records = csv.GetRecords<dynamic>().ToList();
                {
                    var row = records[0] as IDictionary<string, object>;
                    Assert.Equal(@"""<>+-*//}{\\n", row["1"]);
                    Assert.Equal("1234567890", row["2"]);
                    Assert.Equal("True", row["3"]);
                    Assert.Equal("2021-01-01 00:00:00", row["4"]);
                }
                {
                    var row = records[1] as IDictionary<string, object>;
                    Assert.Equal("<test>Hello World</test>", row["1"]);
                    Assert.Equal("-1234567890", row["2"]);
                    Assert.Equal("False", row["3"]);
                    Assert.Equal("2021-01-02 00:00:00", row["4"]);
                }
            }
        }
    }

    [Fact]
    public async Task SaveAsByDataTableTest()
    {
        {
            using var file = AutoDeletingPath.Create(ExcelType.CSV);
            var path = file.ToString();

            var table = new DataTable();
            await MiniExcel.SaveAsAsync(path, table);

            var text = File.ReadAllText(path);
            Assert.Equal("\r\n", text);
        }

        {
            var path = Path.Combine(Path.GetTempPath(), $"{Guid.NewGuid()}.csv");

            var table = new DataTable();
            {
                table.Columns.Add("a", typeof(string));
                table.Columns.Add("b", typeof(decimal));
                table.Columns.Add("c", typeof(bool));
                table.Columns.Add("d", typeof(DateTime));
                table.Rows.Add(@"""<>+-*//}{\\n", 1234567890, true, new DateTime(2021, 1, 1));
                table.Rows.Add("<test>Hello World</test>", -1234567890, false, new DateTime(2021, 1, 2));
            }

            var rowsWritten = await MiniExcel.SaveAsAsync(path, table);
            Assert.Equal(2, rowsWritten[0]);
                
            using (var reader = new StreamReader(path))
            using (var csv = new CsvReader(reader, CultureInfo.InvariantCulture))
            {
                var records = csv.GetRecords<dynamic>().ToList();
                Assert.Equal(@"""<>+-*//}{\\n", records[0].a);
                Assert.Equal("1234567890", records[0].b);
                Assert.Equal("True", records[0].c);
                Assert.Equal("2021-01-01 00:00:00", records[0].d);

                Assert.Equal("<test>Hello World</test>", records[1].a);
                Assert.Equal("-1234567890", records[1].b);
                Assert.Equal("False", records[1].c);
                Assert.Equal("2021-01-02 00:00:00", records[1].d);
            }
        }
    }


    private class Test
    {
        public string c1 { get; set; }
        public string c2 { get; set; }
    }

    [Fact]
    public async Task CsvExcelTypeTest()
    {
        {
            using var file = AutoDeletingPath.Create(ExcelType.CSV);
            var path = file.ToString();

            var input = new[] { new { A = "Test1", B = "Test2" } };
            await MiniExcel.SaveAsAsync(path, input);

            var texts = File.ReadAllLines(path);
            Assert.Equal("A,B", texts[0]);
            Assert.Equal("Test1,Test2", texts[1]);

            {
                var q = await MiniExcel.QueryAsync(path);
                var rows = q.ToList();
                Assert.Equal("A", rows[0].A);
                Assert.Equal("B", rows[0].B);
                Assert.Equal("Test1", rows[1].A);
                Assert.Equal("Test2", rows[1].B);
            }

            using (var reader = new StreamReader(path))
            using (var csv = new CsvReader(reader, CultureInfo.InvariantCulture))
            {
                var rows = csv.GetRecords<dynamic>().ToList();
                Assert.Equal("Test1", rows[0].A);
                Assert.Equal("Test2", rows[0].B);
            }
        }
    }

    [Fact]
    public async Task Create2x2_Test()
    {
        using var file = AutoDeletingPath.Create(ExcelType.CSV);
        var path = file.ToString();

        await MiniExcel.SaveAsAsync(path, new[] {
            new { c1 = "A1" ,c2 = "B1"},
            new { c1 = "A2" ,c2 = "B2"},
        });

        using (var stream = File.OpenRead(path))
        {
            var rows = stream.Query(useHeaderRow: true, excelType: ExcelType.CSV).ToList();
            Assert.Equal("A1", rows[0].c1);
            Assert.Equal("B1", rows[0].c2);
            Assert.Equal("A2", rows[1].c1);
            Assert.Equal("B2", rows[1].c2);
        }

        {
            var rows = MiniExcel.Query(path, useHeaderRow: true, excelType: ExcelType.CSV).ToList();
            Assert.Equal("A1", rows[0].c1);
            Assert.Equal("B1", rows[0].c2);
            Assert.Equal("A2", rows[1].c1);
            Assert.Equal("B2", rows[1].c2);
        }
    }

    [Fact]
    public async Task CsvTypeMappingTest()
    {
        using var file = AutoDeletingPath.Create(ExcelType.CSV);
        var path = file.ToString();

        await MiniExcel.SaveAsAsync(path, new[] {
            new { c1 = "A1" ,c2 = "B1"},
            new { c1 = "A2" ,c2 = "B2"},
        });

        using (var stream = File.OpenRead(path))
        {
            var rows = stream.Query<Test>(excelType: ExcelType.CSV).ToList();
            Assert.Equal("A1", rows[0].c1);
            Assert.Equal("B1", rows[0].c2);
            Assert.Equal("A2", rows[1].c1);
            Assert.Equal("B2", rows[1].c2);
        }

        {
            var rows = MiniExcel.Query<Test>(path, excelType: ExcelType.CSV).ToList();
            Assert.Equal("A1", rows[0].c1);
            Assert.Equal("B1", rows[0].c2);
            Assert.Equal("A2", rows[1].c1);
            Assert.Equal("B2", rows[1].c2);
        }
    }

    [Fact]
    public async Task CsvReadEmptyStringAsNullTest()
    {
        using var file = AutoDeletingPath.Create(ExcelType.CSV);
        var path = file.ToString();
        
        await MiniExcel.SaveAsAsync(path, new[] 
        {
            new { c1 = "A1", c2 = (string)null},
            new { c1 = (string)null, c2 = (string)null},
        });

        using (var stream = File.OpenRead(path))
        {
            var rows = stream.Query<Test>(excelType: ExcelType.CSV).ToList();
            Assert.Equal("A1", rows[0].c1);
            Assert.Equal(string.Empty, rows[0].c2);
            Assert.Equal(string.Empty, rows[1].c1);
            Assert.Equal(string.Empty, rows[1].c2);
        }

        {
            var rows = MiniExcel.Query<Test>(path, excelType: ExcelType.CSV).ToList();
            Assert.Equal("A1", rows[0].c1);
            Assert.Equal(string.Empty, rows[0].c2);
            Assert.Equal(string.Empty, rows[1].c1);
            Assert.Equal(string.Empty, rows[1].c2);
        }

        var config = new Csv.CsvConfiguration { ReadEmptyStringAsNull = true };
        using (var stream = File.OpenRead(path))
        {
            var rows = stream.Query<Test>(excelType: ExcelType.CSV, configuration: config).ToList();
            Assert.Equal("A1", rows[0].c1);
            Assert.Null(rows[0].c2);
            Assert.Null(rows[1].c1);
            Assert.Null(rows[1].c2);
        }

        {
            var rows = MiniExcel.Query<Test>(path, excelType: ExcelType.CSV, configuration: config).ToList();
            Assert.Equal("A1", rows[0].c1);
            Assert.Null(rows[0].c2);
            Assert.Null(rows[1].c1);
            Assert.Null(rows[1].c2);
        }
    }

    [Fact]
    public async Task SaveAsByAsyncEnumerable()
    {
        using var file = AutoDeletingPath.Create();
        var path = file.ToString();

#pragma warning disable CS1998 // Async method lacks 'await' operators and will run synchronously
        static async IAsyncEnumerable<Test> GetValues()
        {
            yield return new Test { c1 = "A1", c2 = "B1" };
            yield return new Test { c1 = "A2", c2 = "B2" };
        }
#pragma warning restore CS1998 // Async method lacks 'await' operators and will run synchronously

        var rowsWritten = await MiniExcel.SaveAsAsync(path, GetValues());
        Assert.Equal(2, rowsWritten[0]);
    
        var results = MiniExcel.Query<Test>(path).ToList();
        Assert.Equal(2, results.Count);
        Assert.Equal("A1", results[0].c1);
        Assert.Equal("B1", results[0].c2);
        Assert.Equal("A2", results[1].c1);
        Assert.Equal("B2", results[1].c2);
    }
}