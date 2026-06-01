namespace MiniExcelLib.Csv.Tests.Issues;

public class AsyncIssueTests
{
    private readonly CsvExporter _csvExporter = MiniExcel.Exporters.GetCsvExporter();
    private readonly CsvImporter _csvImporter = MiniExcel.Importers.GetCsvImporter();

    // Support Enum Mapping
    [Fact]
    public async Task Issue89()
    {
        const string text =
            """
            State
            OnDuty
            Fired
            Leave
            """;

        await using var stream = new MemoryStream();
        await using var writer = new StreamWriter(stream);

        await writer.WriteAsync(text);
        await writer.FlushAsync();

        stream.Position = 0;
        var rows1 = await _csvImporter.QueryAsync<Issue89Dto>(stream).ToListAsync();

        Assert.Equal(Issue89Dto.WorkState.OnDuty, rows1[0].State);
        Assert.Equal(Issue89Dto.WorkState.Fired, rows1[1].State);
        Assert.Equal(Issue89Dto.WorkState.Leave, rows1[2].State);

        var outputPath = PathHelper.GetTempPath();
        var rowsWritten = await MiniExcel.Exporters.GetOpenXmlExporter().ExportAsync(outputPath, rows1);
        Assert.Single(rowsWritten);
        Assert.Equal(3, rowsWritten[0]);

        var rows2 = await MiniExcel.Importers.GetOpenXmlImporter().QueryAsync<Issue89Dto>(outputPath).ToListAsync();
        Assert.Equal(Issue89Dto.WorkState.OnDuty, rows2[0].State);
        Assert.Equal(Issue89Dto.WorkState.Fired, rows2[1].State);
        Assert.Equal(Issue89Dto.WorkState.Leave, rows2[2].State);
    }

    [Fact]
    public async Task Issue142()
    {
        using var file = AutoDeletingPath.Create(ExcelType.Csv);
        var path = file.ToString();

        await _csvExporter.ExportAsync(path, new[] { new Issue142Dto { MyProperty1 = "MyProperty1", MyProperty2 = "MyProperty2", MyProperty3 = "MyProperty3", MyProperty4 = "MyProperty4", MyProperty5 = "MyProperty5", MyProperty6 = "MyProperty6", MyProperty7 = "MyProperty7" } });
        const string expected =
            """
            MyProperty4,CustomColumnName,MyProperty5,MyProperty2,MyProperty6,,MyProperty3
            MyProperty4,MyProperty1,MyProperty5,MyProperty2,MyProperty6,,MyProperty3

            """;

        Assert.Equal(expected, await File.ReadAllTextAsync(path));
        var rows = await _csvImporter.QueryAsync<Issue142Dto>(path).ToListAsync();

        Assert.Equal("MyProperty4", rows[0].MyProperty4);
        Assert.Equal("MyProperty1", rows[0].MyProperty1);
        Assert.Equal("MyProperty5", rows[0].MyProperty5);
        Assert.Equal("MyProperty2", rows[0].MyProperty2);
        Assert.Equal("MyProperty6", rows[0].MyProperty6);
        Assert.Null(rows[0].MyProperty7);
        Assert.Equal("MyProperty3", rows[0].MyProperty3);
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

        using var path = AutoDeletingPath.Create(ExcelType.Csv);
        await  _csvExporter.ExportAsync(path.ToString(), table);

        var rows = await _csvImporter.QueryAsync(path.ToString()).ToListAsync();
        Assert.Equal("Name", rows[0].B);
        Assert.Equal("Limit", rows[0].C);
    }

    
    // Csv QueryAsync split comma not correct
    [Fact]
    public async Task Issue237()
    {
        using var path = AutoDeletingPath.Create(ExcelType.Csv);
        var value = new[]
        {
            new{ id = "\"\"1,2,3\"\"" },
            new{ id = "1,2,3" }
        };
        await  _csvExporter.ExportAsync(path.ToString(), value);

        var rows = await _csvImporter.QueryAsync(path.ToString(), true).ToListAsync();
        Assert.Equal("\"\"1,2,3\"\"", rows[0].id);
        Assert.Equal("1,2,3", rows[1].id);
    }

    // Support Custom Datetime format
    [Fact]
    public async Task Issue241()
    {
        Issue241Dto[] value =
        [
            new() { Name = "Jack", InDate = new DateTime(2021, 01, 04) },
            new() { Name = "Henry", InDate = new DateTime(2020, 04, 05) }
        ];
        
        using var file = AutoDeletingPath.Create(ExcelType.Csv);
        var path = file.ToString();
        var rowsWritten = await _csvExporter.ExportAsync(path, value);

        Assert.Equal(2, rowsWritten);

        var rows1 = await _csvImporter.QueryAsync(path, true).ToListAsync();
        Assert.Equal(rows1[0].InDate, "01 04, 2021");
        Assert.Equal(rows1[1].InDate, "04 05, 2020");

        var rows2 = await _csvImporter.QueryAsync<Issue241Dto>(path).ToListAsync();
        Assert.Equal(rows2[0].InDate, new DateTime(2021, 01, 04));
        Assert.Equal(rows2[1].InDate, new DateTime(2020, 04, 05));
    }

    
    // Csv type mapping QueryAsync error "cannot be converted to xxx type"
    [Fact]
    public async Task Issue243()
    {
        using var path = AutoDeletingPath.Create(ExcelType.Csv);
        var value = new[] 
        {
            new { Name = "Jack", Age = 25, InDate = new DateTime(2021,01,03) },
            new { Name = "Henry", Age = 36, InDate = new DateTime(2020,05,03) }
        };
        
        var rowsWritten = await  _csvExporter.ExportAsync(path.ToString(), value);
        Assert.Equal(2, rowsWritten);

        var rows = await _csvImporter.QueryAsync<Issue243Dto>(path.ToString()).ToListAsync();

        Assert.Equal("Jack", rows[0].Name);
        Assert.Equal(25, rows[0].Age);
        Assert.Equal(new DateTime(2021, 01, 03), rows[0].InDate);

        Assert.Equal("Henry", rows[1].Name);
        Assert.Equal(36, rows[1].Age);
        Assert.Equal(new DateTime(2020, 05, 03), rows[1].InDate);
    }

    // CSV support export from datareader
    [Fact]
    public async Task Issue251()
    {
        await using var cn = Db.GetConnection();
        var reader = await cn.ExecuteReaderAsync(@"select '""<>+-*//}{\\n' a,1234567890 b union all select '<test>Hello World</test>',-1234567890");

        using var path = AutoDeletingPath.Create(ExcelType.Csv);
        var rowsWritten = await _csvExporter.ExportAsync(path.ToString(), reader);

        Assert.Equal(2, rowsWritten);

        const string expected =
            """"
            a,b
            """<>+-*//}{\\n",1234567890
            "<test>Hello World</test>",-1234567890

            """";

        Assert.Equal(expected, await File.ReadAllTextAsync(path.ToString()));
    }

    // Csv export from datareader with default encoding shows incorrect result
    [Fact]
    public async Task Issue253()
    {
        {
            var value = new[] { new { col1 = "世界你好" } };
            using var path = AutoDeletingPath.Create(ExcelType.Csv);

            await _csvExporter.ExportAsync(path.ToString(), value);
            const string expected =
                """
                col1
                世界你好

                """;

            Assert.Equal(expected, await File.ReadAllTextAsync(path.ToString()));
        }

        {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
            var value = new[] { new { col1 = "世界你好" } };
            using var path = AutoDeletingPath.Create(ExcelType.Csv);

            var config = new CsvConfiguration
            {
                StreamWriterFunc = stream => new StreamWriter(stream, Encoding.GetEncoding("gb2312"))
            };

            await _csvExporter.ExportAsync(path.ToString(), value, configuration: config);
            const string expected =
                """
                col1
                �������

                """;

            Assert.Equal(expected, await File.ReadAllTextAsync(path.ToString()));
        }

        await using var cn = Db.GetConnection();

        {
            var value = await cn.ExecuteReaderAsync("select '世界你好' col1");
            using var path = AutoDeletingPath.Create(ExcelType.Csv);
            await _csvExporter.ExportAsync(path.ToString(), value);
            const string expected =
                """
                col1
                世界你好

                """;

            Assert.Equal(expected, await File.ReadAllTextAsync(path.ToString()));
        }
    }
}
