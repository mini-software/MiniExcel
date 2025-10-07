namespace MiniExcelLib.Tests;

public class MiniExcelWriterTests
{
    private readonly MiniExcelExporterProvider _excelExporterProvider = new();
    private readonly MiniExcelImporterProvider _excelImporterProvider = new();

    [Fact]
    public async Task ExportDataTableWithProgressTest()
    {
        var dataTable = new DataTable();
        dataTable.Columns.Add("Id", typeof(int));
        dataTable.Columns.Add("Name", typeof(string));
        dataTable.Columns.Add("Date", typeof(DateTime));
        dataTable.Rows.Add(1, "Alice", DateTime.Now);
        dataTable.Rows.Add(2, DBNull.Value, DateTime.UtcNow);
        dataTable.Rows.Add(3, "Alice", DateTime.Now.Date);

        var progress = new SimpleProgress();
        using var ms = new MemoryStream();
        var exporter = _excelExporterProvider.GetOpenXmlExporter();
        var rowCounts = await exporter.ExportAsync(ms, dataTable, progress: progress);
        Assert.Single(rowCounts);
        Assert.Equal(3, rowCounts.First());

        //Confirm the progress report is correct
        var cellCount = dataTable.Columns.Count * dataTable.Rows.Count;
        Assert.Equal(cellCount, progress.Value);

        ms.Seek(0, SeekOrigin.Begin);
        var importer = _excelImporterProvider.GetOpenXmlImporter();
        var resultDataTable = importer.QueryAsDataTable(ms);

        //Confirm the data is correct
        Assert.Equal(dataTable.Rows.Count, resultDataTable.Rows.Count);
        Assert.Equal(dataTable.Columns.Count, resultDataTable.Columns.Count);
        for (var i = 0; i < dataTable.Rows.Count; i++)
        {
            for (var j = 0; j < dataTable.Columns.Count; j++)
            {
                //We compare string values because types change after writing and reading them back (e.g. int becomes double)
                Assert.Equal(dataTable.Rows[i][j].ToString(), resultDataTable.Rows[i][j].ToString());
            }
        }
    }

    private class SimpleProgress: IProgress<int>
    {
        public int Value { get; private set; }
        public void Report(int value)
        {
            Value += value;
        }
    }
}
