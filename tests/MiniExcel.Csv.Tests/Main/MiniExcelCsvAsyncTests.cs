namespace MiniExcelLib.Csv.Tests.Main;

public class MiniExcelCsvAsyncTests
{
    private readonly CsvExporter _csvExporter = MiniExcel.Exporters.GetCsvExporter();
    private readonly CsvImporter _csvImporter = MiniExcel.Importers.GetCsvImporter();
    
    [Fact]
    public async Task Gb2312_Encoding_Read_Test()
    {
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
        var path = PathHelper.GetFile("csv/gb2312_Encoding_Read_Test.csv");
        var config = new CsvConfiguration
        {
            StreamReaderFunc = stream => new StreamReader(stream, encoding: Encoding.GetEncoding("gb2312"))
        };
        var rows = await _csvImporter.QueryAsync(path, true, configuration: config).ToListAsync();
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
            
        var rowsWritten = await _csvExporter.ExportAsync(path, values, configuration: new CsvConfiguration { Seperator = ';' });
        Assert.Equal(2, rowsWritten);
            
        const string expected =
            """"
            a;b;c;d
            """<>+-*//}{\\n";1234567890;True;"2021-01-01 00:00:00"
            "<test>Hello World</test>";-1234567890;False;"2021-01-02 00:00:00"

            """";
        
        Assert.Equal(expected, await File.ReadAllTextAsync(path));
    }

    [Fact]
    public async Task WriteNullValueTest()
    {
        using var path = AutoDeletingPath.Create(ExcelType.Csv);
        await _csvExporter.ExportAsync(path.FilePath, null!);
        Assert.Equal("", File.ReadAllText(path.FilePath));
    }

    [Fact]
    public async Task SaveAsByDictionary()
    {
        {
            using var file = AutoDeletingPath.Create(ExcelType.Csv);
            var path = file.ToString();

            var table = new List<Dictionary<string, object>>();
            await _csvExporter.ExportAsync(path, table);
            Assert.Equal("\r\n", await File.ReadAllTextAsync(path));
        }

        {
            using var file = AutoDeletingPath.Create(ExcelType.Csv);
            var path = file.ToString();

            var table = new Dictionary<string, object>(); //TODO
            Assert.Throws<NotSupportedException>(() => _csvExporter.Export(path, table));
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
            var rowsWritten = await _csvExporter.ExportAsync(path, values);
            Assert.Equal(2, rowsWritten);

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
            
            var rowsWritten = await _csvExporter.ExportAsync(path, values);
            Assert.Equal(2, rowsWritten);

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
        await _csvExporter.ExportAsync(path1, emptyTable);

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

        var rowsWritten = await _csvExporter.ExportAsync(path2, table);
        Assert.Equal(2, rowsWritten);

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

    [Fact]
    public async Task CsvExcelTypeTest()
    {
        using var file = AutoDeletingPath.Create(ExcelType.Csv);
        var path = file.ToString();

        var input = new[] { new { A = "Test1", B = "Test2" } };
        await _csvExporter.ExportAsync(path, input);

        var texts = await File.ReadAllLinesAsync(path);
        Assert.Equal("A,B", texts[0]);
        Assert.Equal("Test1,Test2", texts[1]);

        var q = _csvImporter.QueryAsync(path).ToBlockingEnumerable();
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

        await _csvExporter.ExportAsync(path, new[] 
        {
            new { C1 = "A1", C2 = "B1"},
            new { C1 = "A2", C2 = "B2"},
        });

        await using (var stream = File.OpenRead(path))
        {
            var rows = _csvImporter.QueryAsync(stream, hasHeaderRow: true).ToBlockingEnumerable().ToList();
            Assert.Equal("A1", rows[0].C1);
            Assert.Equal("B1", rows[0].C2);
            Assert.Equal("A2", rows[1].C1);
            Assert.Equal("B2", rows[1].C2);
        }

        {
            var rows = _csvImporter.QueryAsync(path, hasHeaderRow: true).ToBlockingEnumerable().ToList();
            Assert.Equal("A1", rows[0].C1);
            Assert.Equal("B1", rows[0].C2);
            Assert.Equal("A2", rows[1].C1);
            Assert.Equal("B2", rows[1].C2);
        }
    }

    [Fact]
    public async Task CsvTypeMappingTest()
    {
        using var file = AutoDeletingPath.Create(ExcelType.Csv);
        var path = file.ToString();

        await _csvExporter.ExportAsync(path, new[] 
        {
            new { C1 = "A1", C2 = "B1"},
            new { C1 = "A2", C2 = "B2"}
        });

        await using (var stream = File.OpenRead(path))
        {
            var rows = _csvImporter.Query<TestDto>(stream).ToList();
            Assert.Equal("A1", rows[0].C1);
            Assert.Equal("B1", rows[0].C2);
            Assert.Equal("A2", rows[1].C1);
            Assert.Equal("B2", rows[1].C2);
        }

        {
            var rows = _csvImporter.Query<TestDto>(path).ToList();
            Assert.Equal("A1", rows[0].C1);
            Assert.Equal("B1", rows[0].C2);
            Assert.Equal("A2", rows[1].C1);
            Assert.Equal("B2", rows[1].C2);
        }
    }

    [Fact]
    public async Task CsvReadEmptyStringAsNullTest()
    {
        using var file = AutoDeletingPath.Create(ExcelType.Csv);
        var path = file.ToString();
        
        await _csvExporter.ExportAsync(path, new[] 
        {
            new { C1 = (string?)"A1", C2 = (string?)null},
            new { C1 = (string?)null, C2 = (string?)null}
        });

        await using (var stream = File.OpenRead(path))
        {
            var rows = _csvImporter.Query<TestDto>(stream).ToList();
            Assert.Equal("A1", rows[0].C1);
            Assert.Equal(string.Empty, rows[0].C2);
            Assert.Equal(string.Empty, rows[1].C1);
            Assert.Equal(string.Empty, rows[1].C2);
        }

        {
            var rows = _csvImporter.Query<TestDto>(path).ToList();
            Assert.Equal("A1", rows[0].C1);
            Assert.Equal(string.Empty, rows[0].C2);
            Assert.Equal(string.Empty, rows[1].C1);
            Assert.Equal(string.Empty, rows[1].C2);
        }

        var config = new CsvConfiguration { ReadEmptyStringAsNull = true };
        await using (var stream = File.OpenRead(path))
        {
            var rows = _csvImporter.Query<TestDto>(stream, configuration: config).ToList();
            Assert.Equal("A1", rows[0].C1);
            Assert.Null(rows[0].C2);
            Assert.Null(rows[1].C1);
            Assert.Null(rows[1].C2);
        }

        {
            var rows = _csvImporter.Query<TestDto>(path, configuration: config).ToList();
            Assert.Equal("A1", rows[0].C1);
            Assert.Null(rows[0].C2);
            Assert.Null(rows[1].C1);
            Assert.Null(rows[1].C2);
        }
    }

    [Fact]
    public async Task SaveAsByAsyncEnumerable()
    {
        using var file = AutoDeletingPath.Create();
        var path = file.ToString();

        static async IAsyncEnumerable<TestDto> GetValues()
        {
            yield return await Task.FromResult(new TestDto { C1 = "A1", C2 = "B1" });
            yield return await Task.FromResult(new TestDto { C1 = "A2", C2 = "B2" });
        }

        var rowsWritten = await _csvExporter.ExportAsync(path, GetValues());
        Assert.Equal(2, rowsWritten);
    
        var results = _csvImporter.Query<TestDto>(path).ToList();
        Assert.Equal(2, results.Count);
        Assert.Equal("A1", results[0].C1);
        Assert.Equal("B1", results[0].C2);
        Assert.Equal("A2", results[1].C1);
        Assert.Equal("B2", results[1].C2);
    }

    [Fact]
    public async Task AppendToCsvTest()
    {
        using var file = AutoDeletingPath.Create(ExcelType.Csv);
        var path = file.ToString();

        {
            var value = new[]
            {
                new { ID = 1, Name = "Jack", InDate = new DateTime(2021,01,03) },
                new { ID = 2, Name = "Henry", InDate = new DateTime(2020,05,03) },
            };
            await _csvExporter.AppendAsync(path, value);

            var content = await File.ReadAllTextAsync(path);
            Assert.Equal(
                """
                ID,Name,InDate
                1,Jack,"2021-01-03 00:00:00"
                2,Henry,"2020-05-03 00:00:00"

                """, content);
        }
        {
            var value = new { ID = 3, Name = "Mike", InDate = new DateTime(2021, 04, 23) };
            await _csvExporter.AppendAsync(path, value);
            
            var content = await File.ReadAllTextAsync(path);
            Assert.Equal(
                """
                ID,Name,InDate
                1,Jack,"2021-01-03 00:00:00"
                2,Henry,"2020-05-03 00:00:00"
                3,Mike,"2021-04-23 00:00:00"

                """, content);
        }
        {
            var value = new[]
            {
                new { ID = 4, Name = "Frank", InDate = new DateTime(2021,06,07) },
                new { ID = 5, Name = "Gloria", InDate = new DateTime(2022,05,03) }
            };
            await _csvExporter.AppendAsync(path, value);

            var content = await File.ReadAllTextAsync(path);
            Assert.Equal(
                """
                ID,Name,InDate
                1,Jack,"2021-01-03 00:00:00"
                2,Henry,"2020-05-03 00:00:00"
                3,Mike,"2021-04-23 00:00:00"
                4,Frank,"2021-06-07 00:00:00"
                5,Gloria,"2022-05-03 00:00:00"

                """, content);
        }
    }

    [Fact]
    public async Task ExportDataTableWithProgressTest()
    {
        var dataTable = new DataTable();
        dataTable.Columns.Add("Id", typeof(int));
        dataTable.Columns.Add("Name", typeof(string));
        dataTable.Columns.Add("Date", typeof(DateTime));
        dataTable.Rows.Add(1, "Alice", new DateTime(1900, 1, 1, 1, 0, 0));
        dataTable.Rows.Add(2, DBNull.Value, new DateTime(1901, 2, 2, 2, 0, 0));
        dataTable.Rows.Add(3, "Alice", DateTime.Now.Date);

        using var path = AutoDeletingPath.Create();

        var progress = new SimpleProgress();
        var rowCounts = await _csvExporter.ExportAsync(path.FilePath, dataTable, progress: progress);
        Assert.Equal(3, rowCounts);

        //Confirm the progress report is correct
        var cellCount = dataTable.Columns.Count * dataTable.Rows.Count;
        Assert.Equal(cellCount, progress.Value);

        var resultDataTable = await _csvImporter.QueryAsDataTableAsync(path.FilePath);

        //Confirm the data is correct
        Assert.Equal(dataTable.Rows.Count, resultDataTable.Rows.Count);
        Assert.Equal(dataTable.Columns.Count, resultDataTable.Columns.Count);
        for (var i = 0; i < dataTable.Rows.Count; i++)
        {
            for (var j = 0; j < dataTable.Columns.Count; j++)
            {
                if (dataTable.Columns[j].DataType == typeof(DateTime))
                {
                    //We need to compare Dates properly as they will be formatted differently in CSV
                    //Note: if dates have millisecond precision that will be lost when saving to CSV
                    DateTime.TryParse(resultDataTable.Rows[i][j].ToString(), out var resultDate);
                    Assert.Equal((DateTime)dataTable.Rows[i][j], resultDate);
                }
                else
                {
                    //We compare string values because types change after writing and reading them back
                    Assert.Equal(dataTable.Rows[i][j].ToString(), resultDataTable.Rows[i][j].ToString());
                }
            }
        }
    }

    [Fact]
    public async Task GetColumnNamesTest()
    {
        var path = PathHelper.GetFile(@"csv/TestHeader.csv");
        var cols = (await _csvImporter.GetColumnNamesAsync(path, true)).ToArray();
        Assert.Equal("Column1", cols[0]);
        Assert.Equal("Column2", cols[1]);
    }

    [Fact]
    public async Task GetColumnNamesEmptyTest()
    {
        await using var ms = new MemoryStream();
        var cols = await _csvImporter.GetColumnNamesAsync(ms);
        Assert.Empty(cols);
    }
}
