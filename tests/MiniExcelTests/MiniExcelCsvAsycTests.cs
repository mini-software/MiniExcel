using CsvReader = CsvHelper.CsvReader;

namespace MiniExcelLib.Tests;

public class MiniExcelCsvAsycTests
{
    private readonly MiniExcelImporter _importer =  MiniExcel.GetImporter();
    private readonly MiniExcelExporter _exporter =  MiniExcel.GetExporter();
    
    [Fact]
    public async Task Gb2312_Encoding_Read_Test()
    {
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
        var path = PathHelper.GetFile("csv/gb2312_Encoding_Read_Test.csv");
        var config = new CsvConfiguration
        {
            StreamReaderFunc = stream => new StreamReader(stream, encoding: Encoding.GetEncoding("gb2312"))
        };
        var q = _importer.QueryCsvAsync(path, true, configuration: config).ToBlockingEnumerable();
        var rows = q.ToList();
        Assert.Equal("世界你好", rows[0].栏位1);
    }

    [Fact]
    public async Task SeperatorTest()
    {
        using var file = AutoDeletingPath.Create(ExcelType.Csv);
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
            
        var rowsWritten = await _exporter.ExportCsvAsync(path, values, configuration: new CsvConfiguration { Seperator = ';' });
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
            using var file = AutoDeletingPath.Create(ExcelType.Csv);
            var path = file.ToString();

            var table = new List<Dictionary<string, object>>();
            await _exporter.ExportCsvAsync(path, table);
            Assert.Equal("\r\n", await File.ReadAllTextAsync(path));
        }

        {
            using var file = AutoDeletingPath.Create(ExcelType.Csv);
            var path = file.ToString();

            var table = new Dictionary<string, object>(); //TODO
            Assert.Throws<NotSupportedException>(() => _exporter.ExportCsv(path, table));
        }

        {
            using var file = AutoDeletingPath.Create(ExcelType.Csv);
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
            var rowsWritten = await _exporter.ExportCsvAsync(path, values);
            Assert.Equal(2, rowsWritten[0]);

            using var reader = new StreamReader(path);
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
            using var file = AutoDeletingPath.Create(ExcelType.Csv);
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
            
            var rowsWritten = await _exporter.ExportCsvAsync(path, values);
            Assert.Equal(2, rowsWritten[0]);

            using var reader = new StreamReader(path);
            using var csv = new CsvReader(reader, CultureInfo.InvariantCulture);
            var records = csv.GetRecords<dynamic>().ToList();
            
            var row1 = records[0] as IDictionary<string, object>;
            Assert.Equal(@"""<>+-*//}{\\n", row1!["1"]);
            Assert.Equal("1234567890", row1["2"]);
            Assert.Equal("True", row1["3"]);
            Assert.Equal("2021-01-01 00:00:00", row1["4"]);
            
            var row2 = records[1] as IDictionary<string, object>;
            Assert.Equal("<test>Hello World</test>", row2!["1"]);
            Assert.Equal("-1234567890", row2["2"]);
            Assert.Equal("False", row2["3"]);
            Assert.Equal("2021-01-02 00:00:00", row2["4"]);
        }
    }

    [Fact]
    public async Task SaveAsByDataTableTest()
    {
        using var file1 = AutoDeletingPath.Create(ExcelType.Csv);
        var path1 = file1.ToString();

        var emptyTable = new DataTable();
        await _exporter.ExportCsvAsync(path1, emptyTable);

        var text = await File.ReadAllTextAsync(path1);
        Assert.Equal("\r\n", text);

        
        using var file2= AutoDeletingPath.Create(ExcelType.Csv);
        var path2 = file2.ToString();

        var table = new DataTable();
        table.Columns.Add("a", typeof(string));
        table.Columns.Add("b", typeof(decimal));
        table.Columns.Add("c", typeof(bool));
        table.Columns.Add("d", typeof(DateTime));
        table.Rows.Add(@"""<>+-*//}{\\n", 1234567890, true, new DateTime(2021, 1, 1));
        table.Rows.Add("<test>Hello World</test>", -1234567890, false, new DateTime(2021, 1, 2));

        var rowsWritten = await _exporter.ExportCsvAsync(path2, table);
        Assert.Equal(2, rowsWritten[0]);

        using var reader = new StreamReader(path2);
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


    private class Test
    {
        public string? c1 { get; set; }
        public string? c2 { get; set; }
    }

    [Fact]
    public async Task CsvExcelTypeTest()
    {
        using var file = AutoDeletingPath.Create(ExcelType.Csv);
        var path = file.ToString();

        var input = new[] { new { A = "Test1", B = "Test2" } };
        await _exporter.ExportCsvAsync(path, input);

        var texts = await File.ReadAllLinesAsync(path);
        Assert.Equal("A,B", texts[0]);
        Assert.Equal("Test1,Test2", texts[1]);

        var q = _importer.QueryCsvAsync(path).ToBlockingEnumerable();
        var rows1 = q.ToList();

        Assert.Equal("A", rows1[0].A);
        Assert.Equal("B", rows1[0].B);
        Assert.Equal("Test1", rows1[1].A);
        Assert.Equal("Test2", rows1[1].B);

        using var reader = new StreamReader(path);
        using var csv = new CsvReader(reader, CultureInfo.InvariantCulture);
        var rows2 = csv.GetRecords<dynamic>().ToList();
        
        Assert.Equal("Test1", rows2[0].A);
        Assert.Equal("Test2", rows2[0].B);
    }

    [Fact]
    public async Task Create2x2_Test()
    {
        using var file = AutoDeletingPath.Create(ExcelType.Csv);
        var path = file.ToString();

        await _exporter.ExportCsvAsync(path, new[] 
        {
            new { c1 = "A1", c2 = "B1"},
            new { c1 = "A2", c2 = "B2"},
        });

        await using (var stream = File.OpenRead(path))
        {
            var rows = _importer.QueryCsvAsync(stream, useHeaderRow: true).ToBlockingEnumerable().ToList();
            Assert.Equal("A1", rows[0].c1);
            Assert.Equal("B1", rows[0].c2);
            Assert.Equal("A2", rows[1].c1);
            Assert.Equal("B2", rows[1].c2);
        }

        {
            var rows = _importer.QueryCsvAsync(path, useHeaderRow: true).ToBlockingEnumerable().ToList();
            Assert.Equal("A1", rows[0].c1);
            Assert.Equal("B1", rows[0].c2);
            Assert.Equal("A2", rows[1].c1);
            Assert.Equal("B2", rows[1].c2);
        }
    }

    [Fact]
    public async Task CsvTypeMappingTest()
    {
        using var file = AutoDeletingPath.Create(ExcelType.Csv);
        var path = file.ToString();

        await _exporter.ExportCsvAsync(path, new[] 
        {
            new { c1 = "A1", c2 = "B1"},
            new { c1 = "A2", c2 = "B2"}
        });

        await using (var stream = File.OpenRead(path))
        {
            var rows = _importer.QueryCsv<Test>(stream).ToList();
            Assert.Equal("A1", rows[0].c1);
            Assert.Equal("B1", rows[0].c2);
            Assert.Equal("A2", rows[1].c1);
            Assert.Equal("B2", rows[1].c2);
        }

        {
            var rows = _importer.QueryCsv<Test>(path).ToList();
            Assert.Equal("A1", rows[0].c1);
            Assert.Equal("B1", rows[0].c2);
            Assert.Equal("A2", rows[1].c1);
            Assert.Equal("B2", rows[1].c2);
        }
    }

    [Fact]
    public async Task CsvReadEmptyStringAsNullTest()
    {
        using var file = AutoDeletingPath.Create(ExcelType.Csv);
        var path = file.ToString();
        
        await _exporter.ExportCsvAsync(path, new[] 
        {
            new { c1 = (string?)"A1", c2 = (string?)null},
            new { c1 = (string?)null, c2 = (string?)null}
        });

        await using (var stream = File.OpenRead(path))
        {
            var rows = _importer.QueryCsv<Test>(stream).ToList();
            Assert.Equal("A1", rows[0].c1);
            Assert.Equal(string.Empty, rows[0].c2);
            Assert.Equal(string.Empty, rows[1].c1);
            Assert.Equal(string.Empty, rows[1].c2);
        }

        {
            var rows = _importer.QueryCsv<Test>(path).ToList();
            Assert.Equal("A1", rows[0].c1);
            Assert.Equal(string.Empty, rows[0].c2);
            Assert.Equal(string.Empty, rows[1].c1);
            Assert.Equal(string.Empty, rows[1].c2);
        }

        var config = new CsvConfiguration { ReadEmptyStringAsNull = true };
        await using (var stream = File.OpenRead(path))
        {
            var rows = _importer.QueryCsv<Test>(stream, configuration: config).ToList();
            Assert.Equal("A1", rows[0].c1);
            Assert.Null(rows[0].c2);
            Assert.Null(rows[1].c1);
            Assert.Null(rows[1].c2);
        }

        {
            var rows = _importer.QueryCsv<Test>(path, configuration: config).ToList();
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

        var rowsWritten = await _exporter.ExportCsvAsync(path, GetValues());
        Assert.Equal(2, rowsWritten[0]);
    
        var results = _importer.QueryCsv<Test>(path).ToList();
        Assert.Equal(2, results.Count);
        Assert.Equal("A1", results[0].c1);
        Assert.Equal("B1", results[0].c2);
        Assert.Equal("A2", results[1].c1);
        Assert.Equal("B2", results[1].c2);
    }
}