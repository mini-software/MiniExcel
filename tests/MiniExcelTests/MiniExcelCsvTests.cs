using CsvHelper;
using MiniExcelLibs.Attributes;
using MiniExcelLibs.Exceptions;
using MiniExcelLibs.Tests.Utils;
using System.Data;
using System.Globalization;
using System.Text;
using Xunit;

namespace MiniExcelLibs.Tests;

public class MiniExcelCsvTests
{
    [Fact]
    public void gb2312_Encoding_Read_Test()
    {
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
        var path = PathHelper.GetFile("csv/gb2312_Encoding_Read_Test.csv");
        var config = new Csv.CsvConfiguration
        {
            StreamReaderFunc = stream => new StreamReader(stream, encoding: Encoding.GetEncoding("gb2312"))
        };
        var rows = MiniExcel.Query(path, true, excelType: ExcelType.CSV, configuration: config).ToList();
        Assert.Equal("世界你好", rows[0].栏位1);
    }

    [Fact]
    public void SeperatorTest()
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
        var rowsWritten = MiniExcel.SaveAs(path, values, configuration: new Csv.CsvConfiguration { Seperator = ';' });
        Assert.Equal(2, rowsWritten[0]);
            
        const string expected =
            """"
            a;b;c;d
            """<>+-*//}{\\n";1234567890;True;"2021-01-01 00:00:00"
            "<test>Hello World</test>";-1234567890;False;"2021-01-02 00:00:00"

            """";
        Assert.Equal(expected, File.ReadAllText(path));
    }

    [Fact]
    public void AlwaysQuoteTest()
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
        
        MiniExcel.SaveAs(path, values, configuration: new Csv.CsvConfiguration { AlwaysQuote = true });
        const string expected = 
            """"
            "a","b","c","d"
            """<>+-*//}{\\n","1234567890","True","2021-01-01 00:00:00"
            "<test>Hello World</test>","-1234567890","False","2021-01-02 00:00:00"

            """";
        Assert.Equal(expected, File.ReadAllText(path));
    }

    [Fact]
    public void QuoteSpecialCharacters()
    {
        using var file = AutoDeletingPath.Create(ExcelType.CSV);
        var path = file.ToString();

        List<Dictionary<string, object>> values =
        [
            new()
            {
                { "a", "potato,banana" },
                { "b", "text\ntest" },
                { "c", "text\rpotato" },
                { "d", new DateTime(2021, 1, 1) }
            }

        ];
        var rowsWritten = MiniExcel.SaveAs(path, values, configuration: new Csv.CsvConfiguration());
        Assert.Equal(1, rowsWritten[0]);
            
        const string expected = "a,b,c,d\r\n\"potato,banana\",\"text\ntest\",\"text\rpotato\",\"2021-01-01 00:00:00\"\r\n";
        Assert.Equal(expected, File.ReadAllText(path));
    }

    [Fact]
    public void SaveAsByDictionary()
    {
        {
            using var path = AutoDeletingPath.Create(ExcelType.CSV);
            var table = new List<Dictionary<string, object>>();
            MiniExcel.SaveAs(path.ToString(), table);
            Assert.Equal("\r\n", File.ReadAllText(path.ToString()));
        }

        {
            using var path = AutoDeletingPath.Create(ExcelType.CSV);

            var table = new Dictionary<string, object>(); //TODO
            Assert.Throws<NotSupportedException>(() => MiniExcel.SaveAs(path.ToString(), table));
        }

        {
            using var path = AutoDeletingPath.Create(ExcelType.CSV);
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
            
            var rowsWritten = MiniExcel.SaveAs(path.ToString(), values);
            Assert.Equal(2, rowsWritten[0]);

            using var reader = new StreamReader(path.ToString());
            using var csv = new CsvReader(reader, CultureInfo.InvariantCulture);
            
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

        {
            using var path = AutoDeletingPath.Create(ExcelType.CSV);
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
            MiniExcel.SaveAs(path.ToString(), values);

            using (var reader = new StreamReader(path.ToString()))
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
    public void SaveAsByDataTableTest()
    {
        {
            using var path = AutoDeletingPath.Create(ExcelType.CSV);
            
            var table = new DataTable();
            MiniExcel.SaveAs(path.ToString(), table);

            var text = File.ReadAllText(path.ToString());
            Assert.Equal("\r\n", text);
        }

        {
            using var path = AutoDeletingPath.Create(ExcelType.CSV);

            var table = new DataTable();
            {
                table.Columns.Add("a", typeof(string));
                table.Columns.Add("b", typeof(decimal));
                table.Columns.Add("c", typeof(bool));
                table.Columns.Add("d", typeof(DateTime));
                table.Rows.Add(@"""<>+-*//}{\\n", 1234567890, true, new DateTime(2021, 1, 1));
                table.Rows.Add("<test>Hello World</test>", -1234567890, false, new DateTime(2021, 1, 2));
            }

            var rowsWritten = MiniExcel.SaveAs(path.ToString(), table);
            Assert.Equal(2, rowsWritten[0]);

            using (var reader = new StreamReader(path.ToString()))
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

    private class TestWithAlias
    {
        [ExcelColumnName(excelColumnName: "c1", aliases: ["column1", "col1"])]
        public string c1 { get; set; }
        [ExcelColumnName(excelColumnName: "c2", aliases: ["column2", "col2"])]
        public string c2 { get; set; }
    }

    [Fact]
    public void CsvExcelTypeTest()
    {
        {
            using var file = AutoDeletingPath.Create(ExcelType.CSV);
            var path = file.ToString();
            
            var input = new[] { new { A = "Test1", B = "Test2" } };
            MiniExcel.SaveAs(path.ToString(), input);

            var texts = File.ReadAllLines(path.ToString());
            Assert.Equal("A,B", texts[0]);
            Assert.Equal("Test1,Test2", texts[1]);

            {
                var rows = MiniExcel.Query(path.ToString()).ToList();
                Assert.Equal("A", rows[0].A);
                Assert.Equal("B", rows[0].B);
                Assert.Equal("Test1", rows[1].A);
                Assert.Equal("Test2", rows[1].B);
            }

            using var reader = new StreamReader(path.ToString());
            using var csv = new CsvReader(reader, CultureInfo.InvariantCulture);
            {
                var rows = csv.GetRecords<dynamic>().ToList();
                Assert.Equal("Test1", rows[0].A);
                Assert.Equal("Test2", rows[0].B);
            }
        }
    }

    [Fact]
    public void Create2x2_Test()
    {
        using var file = AutoDeletingPath.Create(ExcelType.CSV);
        var path = file.ToString();

        MiniExcel.SaveAs(path, new[] 
        {
            new { c1 = "A1", c2 = "B1"},
            new { c1 = "A2", c2 = "B2"},
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
    public void CsvTypeMappingTest()
    {
        using var file = AutoDeletingPath.Create(ExcelType.CSV);
        var path = file.ToString();

        MiniExcel.SaveAs(path, new[] 
        {
            new { c1 = "A1", c2 = "B1"},
            new { c1 = "A2", c2 = "B2"},
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
    public void CsvColumnNotFoundTest()
    {
        using var file = AutoDeletingPath.Create(ExcelType.CSV);
        var path = file.ToString();

        File.WriteAllLines(path, ["c1,c2", "v1"]);

        using (var stream = File.OpenRead(path))
        {
            var exception = Assert.Throws<ExcelColumnNotFoundException>(() => stream.Query<Test>(excelType: ExcelType.CSV).ToList());

            Assert.Equal("c2", exception.ColumnName);
            Assert.Equal(2, exception.RowIndex);
            Assert.Null(exception.ColumnIndex);
            Assert.True(exception.RowValues is IDictionary<string, object>);
            Assert.Equal(1, ((IDictionary<string, object>)exception.RowValues).Count);
        }

        {
            var exception = Assert.Throws<ExcelColumnNotFoundException>(() => MiniExcel.Query<Test>(path, excelType: ExcelType.CSV).ToList());

            Assert.Equal("c2", exception.ColumnName);
            Assert.Equal(2, exception.RowIndex);
            Assert.Null(exception.ColumnIndex);
            Assert.True(exception.RowValues is IDictionary<string, object>);
            Assert.Equal(1, ((IDictionary<string, object>)exception.RowValues).Count);
        }
    }

    [Fact]
    public void CsvColumnNotFoundWithAliasTest()
    {
        using var file = AutoDeletingPath.Create(ExcelType.CSV);
        var path = file.ToString();

        File.WriteAllLines(path, ["col1,col2", "v1"]);
        using (var stream = File.OpenRead(path))
        {
            var exception = Assert.Throws<ExcelColumnNotFoundException>(() => stream.Query<TestWithAlias>(excelType: ExcelType.CSV).ToList());

            Assert.Equal("c2", exception.ColumnName);
            Assert.Equal(2, exception.RowIndex);
            Assert.Null(exception.ColumnIndex);
            Assert.True(exception.RowValues is IDictionary<string, object>);
            Assert.Equal(1, ((IDictionary<string, object>)exception.RowValues).Count);
        }

        {
            var exception = Assert.Throws<ExcelColumnNotFoundException>(() => MiniExcel.Query<TestWithAlias>(path, excelType: ExcelType.CSV).ToList());

            Assert.Equal("c2", exception.ColumnName);
            Assert.Equal(2, exception.RowIndex);
            Assert.Null(exception.ColumnIndex);
            Assert.True(exception.RowValues is IDictionary<string, object>);
            Assert.Equal(1, ((IDictionary<string, object>)exception.RowValues).Count);
        }
    }

    [Fact]
    public void Delimiters_Test()
    {
        //TODO:Datetime have default format like yyyy-MM-dd HH:mm:ss ?
        {
            Assert.Equal(Generate("\"\"\""), MiniExcelGenerateCsv("\"\"\""));
            Assert.Equal(Generate(","), MiniExcelGenerateCsv(","));
            Assert.Equal(Generate(" "), MiniExcelGenerateCsv(" "));
            Assert.Equal(Generate(";"), MiniExcelGenerateCsv(";"));
            Assert.Equal(Generate("\t"), MiniExcelGenerateCsv("\t"));
        }
    }

    private static string Generate(string value)
    {
        using var file = AutoDeletingPath.Create(ExcelType.CSV);
        var path = file.ToString();

        using (var writer = new StreamWriter(path))
        using (var csv = new CsvWriter(writer, CultureInfo.InvariantCulture))
        {
            var records = Enumerable.Range(1, 1).Select(_ => new { v1 = value, v2 = value });
            csv.WriteRecords(records);
        }

        var content = File.ReadAllText(path);
        return content;
    }

    private static string MiniExcelGenerateCsv(string value)
    {
        using var file = AutoDeletingPath.Create(ExcelType.CSV);
        var path = file.ToString();

        using (var stream = File.Create(path))
        {
            IEnumerable<object> records = [new { v1 = value, v2 = value }];
            var rowsWritten = stream.SaveAs(records, excelType: ExcelType.CSV);
            Assert.Equal(1, rowsWritten[0]);
        }

        var content = File.ReadAllText(path);
        return content;
    }
}