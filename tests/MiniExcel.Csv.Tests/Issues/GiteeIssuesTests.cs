namespace MiniExcelLib.Csv.Tests.Issues;

public class GiteeIssuesTests
{
    private readonly CsvExporter _csvExporter = MiniExcel.Exporters.GetCsvExporter();

    // https://gitee.com/dotnetchina/MiniExcel/issues/I4X92G
    [Fact]
    public void TestIssueI4X92G()
    {
        using var file = AutoDeletingPath.Create(ExcelType.Csv);
        var path = file.ToString();

        {
            var value = new[]
            {
                new { ID = 1, Name = "Jack", InDate = new DateTime(2021,01,03)},
                new { ID = 2, Name = "Henry", InDate = new DateTime(2020,05,03)}
            };
            _csvExporter.Export(path, value);
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
            var rowsWritten = _csvExporter.Append(path, value);
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
            var rowsWritten = _csvExporter.Append(path, value);
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

    [Fact]
    public void TestIssueI4Wda9()
    {
        using var path = AutoDeletingPath.Create(ExcelType.Csv);
        var value = new DataTable();
        {
            value.Columns.Add("\"name\"");
            value.Rows.Add("\"Jack\"");
        }

        _csvExporter.Export(path.ToString(), value);
        Assert.Equal("\"\"\"name\"\"\"\r\n\"\"\"Jack\"\"\"\r\n", File.ReadAllText(path.ToString()));
    }

    // Using stream.SaveAs will close the Stream automatically when Specifying excelType
    // https://gitee.com/dotnetchina/MiniExcel/issues/I57WMM
    [Fact]
    public void TestIssueGiteeI57()
    {
        Dictionary<string, object>[] sheets = [new() { ["ID"] = "0001", ["Name"] = "Jack" }];
        using var stream = new MemoryStream();

        var config = new CsvConfiguration { StreamWriterFunc = x => new StreamWriter(x, Encoding.Default, leaveOpen: true) };
        _csvExporter.Export(stream, sheets, configuration: config);
        stream.Seek(0, SeekOrigin.Begin);

        // convert stream to string
        using var reader = new StreamReader(stream);
        var text = reader.ReadToEnd();

        Assert.Equal("ID,Name\r\n0001,Jack\r\n", text);
    }
}