using MiniExcelLib.OpenXml.Api;

namespace MiniExcelLib.Csv.Tests;

public class AsyncIssueTests
{
    private readonly CsvExporter _csvExporter = MiniExcel.Exporters.GetCsvExporter();
    private readonly CsvImporter _csvImporter = MiniExcel.Importers.GetCsvImporter();

    private readonly OpenXmlExporter _openXmlExporter = MiniExcel.Exporters.GetOpenXmlExporter();
    private readonly OpenXmlImporter _openXmlImporter = MiniExcel.Importers.GetOpenXmlImporter();
    /// <summary>
    /// Csv SaveAs by datareader with encoding default show messy code #253
    /// </summary>
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

    /// <summary>
    /// [CSV SaveAs support datareader · Issue #251 · mini-software/MiniExcel](https://github.com/mini-software/MiniExcel/issues/251)
    /// </summary>
    [Fact]
    public async Task Issue251()
    {
        await using var cn = Db.GetConnection();
        var reader = await cn.ExecuteReaderAsync(@"select '""<>+-*//}{\\n' a,1234567890 b union all select '<test>Hello World</test>',-1234567890");

        using var path = AutoDeletingPath.Create(ExcelType.Csv);
        var rowsWritten = await _csvExporter.ExportAsync(path.ToString(), reader);

        Assert.Single(rowsWritten);
        Assert.Equal(2, rowsWritten[0]);

        const string expected =
            """"
            a,b
            """<>+-*//}{\\n",1234567890
            "<test>Hello World</test>",-1234567890

            """";

        Assert.Equal(expected, await File.ReadAllTextAsync(path.ToString()));
    }

    private class Issue89Dto
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
    /// Support Enum Mapping
    /// https://github.com/mini-software/MiniExcel/issues/89
    /// </summary>
    [Fact]
    public async Task Issue89()
    {
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
            var q = _csvImporter.QueryAsync<Issue89Dto>(stream).ToBlockingEnumerable();
            var rows = q.ToList();

            Assert.Equal(Issue89Dto.WorkState.OnDuty, rows[0].State);
            Assert.Equal(Issue89Dto.WorkState.Fired, rows[1].State);
            Assert.Equal(Issue89Dto.WorkState.Leave, rows[2].State);

            var outputPath = PathHelper.GetTempPath();
            var rowsWritten = await _openXmlExporter.ExportAsync(outputPath, rows);
            Assert.Single(rowsWritten);
            Assert.Equal(3, rowsWritten[0]);

            var q2 = _openXmlImporter.QueryAsync<Issue89Dto>(outputPath).ToBlockingEnumerable();
            var rows2 = q2.ToList();
            Assert.Equal(Issue89Dto.WorkState.OnDuty, rows2[0].State);
            Assert.Equal(Issue89Dto.WorkState.Fired, rows2[1].State);
            Assert.Equal(Issue89Dto.WorkState.Leave, rows2[2].State);
        }

        //xlsx
        {
            var path = PathHelper.GetFile("xlsx/TestIssue89.xlsx");
            var q = _openXmlImporter.QueryAsync<Issue89Dto>(path).ToBlockingEnumerable();
            var rows = q.ToList();
            Assert.Equal(Issue89Dto.WorkState.OnDuty, rows[0].State);
            Assert.Equal(Issue89Dto.WorkState.Fired, rows[1].State);
            Assert.Equal(Issue89Dto.WorkState.Leave, rows[2].State);

            var outputPath = PathHelper.GetTempPath();
            var rowsWritten = await _openXmlExporter.ExportAsync(outputPath, rows);
            Assert.Single(rowsWritten);
            Assert.Equal(3, rowsWritten[0]);

            var q1 = _openXmlImporter.QueryAsync<Issue89Dto>(outputPath).ToBlockingEnumerable();
            var rows2 = q1.ToList();
            Assert.Equal(Issue89Dto.WorkState.OnDuty, rows2[0].State);
            Assert.Equal(Issue89Dto.WorkState.Fired, rows2[1].State);
            Assert.Equal(Issue89Dto.WorkState.Leave, rows2[2].State);
        }
    }

    private class Issue142VoDuplicateColumnName
    {
        [MiniExcelColumnIndex("A")]
        public int MyProperty1 { get; set; }
        
        [MiniExcelColumnIndex("A")]
        public int MyProperty2 { get; set; }

        public int MyProperty3 { get; set; }
        [MiniExcelColumnIndex("B")]
        
        public int MyProperty4 { get; set; }
    }

    [Fact]
    public async Task Issue142()
    {
        {
            using var file = AutoDeletingPath.Create();
            var path = file.ToString();
            await _openXmlExporter.ExportAsync(path, new[] { new Issue142Dto { MyProperty1 = "MyProperty1", MyProperty2 = "MyProperty2", MyProperty3 = "MyProperty3", MyProperty4 = "MyProperty4", MyProperty5 = "MyProperty5", MyProperty6 = "MyProperty6", MyProperty7 = "MyProperty7" } });

            {
                var q = _openXmlImporter.QueryAsync(path).ToBlockingEnumerable();
                var rows = q.ToList();
                Assert.Equal("MyProperty4", rows[0].A);
                Assert.Equal("CustomColumnName", rows[0].B); //note
                Assert.Equal("MyProperty5", rows[0].C);
                Assert.Equal("MyProperty2", rows[0].D);
                Assert.Equal("MyProperty6", rows[0].E);
                Assert.Null(rows[0].F);
                Assert.Equal("MyProperty3", rows[0].G);

                Assert.Equal("MyProperty4", rows[0].A);
                Assert.Equal("CustomColumnName", rows[0].B); //note
                Assert.Equal("MyProperty5", rows[0].C);
                Assert.Equal("MyProperty2", rows[0].D);
                Assert.Equal("MyProperty6", rows[0].E);
                Assert.Null(rows[0].F);
                Assert.Equal("MyProperty3", rows[0].G);
            }

            {
                var q = _openXmlImporter.QueryAsync<Issue142Dto>(path).ToBlockingEnumerable();
                var rows = q.ToList();

                Assert.Equal("MyProperty4", rows[0].MyProperty4);
                Assert.Equal("MyProperty1", rows[0].MyProperty1); //note
                Assert.Equal("MyProperty5", rows[0].MyProperty5);
                Assert.Equal("MyProperty2", rows[0].MyProperty2);
                Assert.Equal("MyProperty6", rows[0].MyProperty6);
                Assert.Null(rows[0].MyProperty7);
                Assert.Equal("MyProperty3", rows[0].MyProperty3);
            }
        }

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

            {
                var q = _csvImporter.QueryAsync<Issue142Dto>(path).ToBlockingEnumerable();
                var rows = q.ToList();

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
            using var path = AutoDeletingPath.Create();
            Issue142VoDuplicateColumnName[] input = [new() { MyProperty1 = 0, MyProperty2 = 0, MyProperty3 = 0, MyProperty4 = 0 }];
            Assert.Throws<InvalidOperationException>(() => _openXmlExporter.Export(path.ToString(), input));
        }
    }
    
    /// <summary>
    /// DataTable recommended to use Caption for column name first, then use columname
    /// https://github.com/mini-software/MiniExcel/issues/217
    /// </summary>
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

        var q =  _csvImporter.QueryAsync(path.ToString()).ToBlockingEnumerable();
        var rows = q.ToList();
        Assert.Equal("Name", rows[0].B);
        Assert.Equal("Limit", rows[0].C);
    }

    
    /// <summary>
    /// Csv QueryAsync split comma not correct #237
    /// https://github.com/mini-software/MiniExcel/issues/237
    /// </summary>
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

        var q =  _csvImporter.QueryAsync(path.ToString(), true).ToBlockingEnumerable();
        var rows = q.ToList();

        Assert.Equal("\"\"1,2,3\"\"", rows[0].id);
        Assert.Equal("1,2,3", rows[1].id);
    }


    /// <summary>
    /// Support Custom Datetime format #241
    /// </summary>
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

        Assert.Single(rowsWritten);
        Assert.Equal(2, rowsWritten[0]);

        var q1 = _csvImporter.QueryAsync(path, true).ToBlockingEnumerable();
        var rows1 = q1.ToList();
        Assert.Equal(rows1[0].InDate, "01 04, 2021");
        Assert.Equal(rows1[1].InDate, "04 05, 2020");

        var q2 = _csvImporter.QueryAsync<Issue241Dto>(path).ToBlockingEnumerable();
        var rows2 = q2.ToList();
        Assert.Equal(rows2[0].InDate, new DateTime(2021, 01, 04));
        Assert.Equal(rows2[1].InDate, new DateTime(2020, 04, 05));
    }

    
    /// <summary>
    /// Csv type mapping QueryAsync error "cannot be converted to xxx type" #243
    /// </summary>
    [Fact]
    public async Task Issue243()
    {
        using var path = AutoDeletingPath.Create(ExcelType.Csv);
        var value = new[] 
        {
            new { Name ="Jack",Age=25,InDate=new DateTime(2021,01,03)},
            new { Name ="Henry",Age=36,InDate=new DateTime(2020,05,03)},
        };
        
        var rowsWritten = await  _csvExporter.ExportAsync(path.ToString(), value);
        Assert.Single(rowsWritten);
        Assert.Equal(2, rowsWritten[0]);

        var q =  _csvImporter.QueryAsync<Issue243Dto>(path.ToString()).ToBlockingEnumerable();
        var rows = q.ToList();
        
        Assert.Equal("Jack", rows[0].Name);
        Assert.Equal(25, rows[0].Age);
        Assert.Equal(new DateTime(2021, 01, 03), rows[0].InDate);

        Assert.Equal("Henry", rows[1].Name);
        Assert.Equal(36, rows[1].Age);
        Assert.Equal(new DateTime(2020, 05, 03), rows[1].InDate);
    }

    #region Duplicated
    private class Issue142Dto
    {
        [MiniExcelColumnName("CustomColumnName")]
        public string? MyProperty1 { get; set; }  //index = 1
        [MiniExcelIgnore]
        public string? MyProperty7 { get; set; } //index = null
        public string? MyProperty2 { get; set; } //index = 3
        [MiniExcelColumnIndex(6)]
        public string? MyProperty3 { get; set; } //index = 6
        [MiniExcelColumnIndex("A")] // equal column index 0
        public string? MyProperty4 { get; set; }
        [MiniExcelColumnIndex(2)]
        public string? MyProperty5 { get; set; } //index = 2
        public string? MyProperty6 { get; set; } //index = 4
    }
    
    private class Issue241Dto
    {
        public string? Name { get; set; }

        [MiniExcelFormat("MM dd, yyyy")]
        public DateTime InDate { get; set; }
    }
    
    private class Issue243Dto
    {
        public string? Name { get; set; }
        public int Age { get; set; }
        public DateTime InDate { get; set; }
    }
    #endregion
}
