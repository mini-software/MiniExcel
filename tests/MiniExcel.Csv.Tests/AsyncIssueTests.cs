using MiniExcelLib.Tests.Common.Utils;

namespace MiniExcel.Csv.Tests;

public class AsyncIssueTests
{
    private readonly MiniExcelExporter _exporter = MiniExcelLib.MiniExcel.GetExporter();
    private readonly MiniExcelImporter _importer = MiniExcelLib.MiniExcel.GetImporter();

    /// <summary>
    /// Csv SaveAs by datareader with encoding default show messy code #253
    /// </summary>
    [Fact]
    public async Task Issue253()
    {
        {
            var value = new[] { new { col1 = "世界你好" } };
            using var path = AutoDeletingPath.Create(ExcelType.Csv);

            await _exporter.ExportCsvAsync(path.ToString(), value);
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

            await _exporter.ExportCsvAsync(path.ToString(), value, configuration: config);
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
            await _exporter.ExportCsvAsync(path.ToString(), value);
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
        var rowsWritten = await _exporter.ExportCsvAsync(path.ToString(), reader);

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

    private class Issue89VO
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
        //csv
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
            var q = _importer.QueryCsvAsync<Issue89VO>(stream).ToBlockingEnumerable();
            var rows = q.ToList();

            Assert.Equal(Issue89VO.WorkState.OnDuty, rows[0].State);
            Assert.Equal(Issue89VO.WorkState.Fired, rows[1].State);
            Assert.Equal(Issue89VO.WorkState.Leave, rows[2].State);

            var outputPath = PathHelper.GetTempPath("xlsx");
            var rowsWritten = await _exporter.ExportXlsxAsync(outputPath, rows);
            Assert.Single(rowsWritten);
            Assert.Equal(3, rowsWritten[0]);

            var q2 = _importer.QueryXlsxAsync<Issue89VO>(outputPath).ToBlockingEnumerable();
            var rows2 = q2.ToList();
            Assert.Equal(Issue89VO.WorkState.OnDuty, rows2[0].State);
            Assert.Equal(Issue89VO.WorkState.Fired, rows2[1].State);
            Assert.Equal(Issue89VO.WorkState.Leave, rows2[2].State);
        }

        //xlsx
        {
            var path = PathHelper.GetFile("xlsx/TestIssue89.xlsx");
            var q = _importer.QueryXlsxAsync<Issue89VO>(path).ToBlockingEnumerable();
            var rows = q.ToList();
            Assert.Equal(Issue89VO.WorkState.OnDuty, rows[0].State);
            Assert.Equal(Issue89VO.WorkState.Fired, rows[1].State);
            Assert.Equal(Issue89VO.WorkState.Leave, rows[2].State);

            var outputPath = PathHelper.GetTempPath();
            var rowsWritten = await _exporter.ExportXlsxAsync(outputPath, rows);
            Assert.Single(rowsWritten);
            Assert.Equal(3, rowsWritten[0]);

            var q1 = _importer.QueryXlsxAsync<Issue89VO>(outputPath).ToBlockingEnumerable();
            var rows2 = q1.ToList();
            Assert.Equal(Issue89VO.WorkState.OnDuty, rows2[0].State);
            Assert.Equal(Issue89VO.WorkState.Fired, rows2[1].State);
            Assert.Equal(Issue89VO.WorkState.Leave, rows2[2].State);
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
            await _exporter.ExportXlsxAsync(path, new[] { new Issue142VO { MyProperty1 = "MyProperty1", MyProperty2 = "MyProperty2", MyProperty3 = "MyProperty3", MyProperty4 = "MyProperty4", MyProperty5 = "MyProperty5", MyProperty6 = "MyProperty6", MyProperty7 = "MyProperty7" } });

            {
                var q = _importer.QueryXlsxAsync(path).ToBlockingEnumerable();
                var rows = q.ToList();
                Assert.Equal("MyProperty4", rows[0].A);
                Assert.Equal("CustomColumnName", rows[0].B); //note
                Assert.Equal("MyProperty5", rows[0].C);
                Assert.Equal("MyProperty2", rows[0].D);
                Assert.Equal("MyProperty6", rows[0].E);
                Assert.Equal(null, rows[0].F);
                Assert.Equal("MyProperty3", rows[0].G);

                Assert.Equal("MyProperty4", rows[0].A);
                Assert.Equal("CustomColumnName", rows[0].B); //note
                Assert.Equal("MyProperty5", rows[0].C);
                Assert.Equal("MyProperty2", rows[0].D);
                Assert.Equal("MyProperty6", rows[0].E);
                Assert.Equal(null, rows[0].F);
                Assert.Equal("MyProperty3", rows[0].G);
            }

            {
                var q = _importer.QueryXlsxAsync<Issue142VO>(path).ToBlockingEnumerable();
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
            await _exporter.ExportCsvAsync(path, new[] { new Issue142VO { MyProperty1 = "MyProperty1", MyProperty2 = "MyProperty2", MyProperty3 = "MyProperty3", MyProperty4 = "MyProperty4", MyProperty5 = "MyProperty5", MyProperty6 = "MyProperty6", MyProperty7 = "MyProperty7" } });
            const string expected =
                """
                MyProperty4,CustomColumnName,MyProperty5,MyProperty2,MyProperty6,,MyProperty3
                MyProperty4,MyProperty1,MyProperty5,MyProperty2,MyProperty6,,MyProperty3

                """;
            Assert.Equal(expected, await File.ReadAllTextAsync(path));

            {
                var q = _importer.QueryCsvAsync<Issue142VO>(path).ToBlockingEnumerable();
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
            Assert.Throws<InvalidOperationException>(() => _exporter.ExportXlsx(path.ToString(), input));
        }
    }

    #region Duplicated
    private class Issue142VO
    {
        [MiniExcelColumnName("CustomColumnName")]
        public string MyProperty1 { get; set; }  //index = 1
        [MiniExcelIgnore]
        public string MyProperty7 { get; set; } //index = null
        public string MyProperty2 { get; set; } //index = 3
        [MiniExcelColumnIndex(6)]
        public string MyProperty3 { get; set; } //index = 6
        [MiniExcelColumnIndex("A")] // equal column index 0
        public string MyProperty4 { get; set; }
        [MiniExcelColumnIndex(2)]
        public string MyProperty5 { get; set; } //index = 2
        public string MyProperty6 { get; set; } //index = 4
    }
    #endregion

}
