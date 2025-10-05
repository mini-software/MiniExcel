namespace MiniExcelLib.Tests;

public class MiniExcelWriterTests
{
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

        // We need to use the file system because the CsvExporter automatically disposes the stream
        var tempFilePath = Path.GetTempFileName();

        using (var fs = File.Create(tempFilePath))
        {

            var progress = new SimpleProgress();
            var exporter = new CsvExporter();
            var rowCounts = await exporter.ExportAsync(fs, dataTable, progress: progress);
            Assert.Single(rowCounts);
            Assert.Equal(3, rowCounts.First());

            //Confirm the progress report is correct
            var cellCount = dataTable.Columns.Count * dataTable.Rows.Count;
            Assert.Equal(cellCount, progress.Value);
        }

        using (var fs = File.OpenRead(tempFilePath))
        {
            var importer = new CsvImporter();
            var resultDataTable = importer.QueryAsDataTable(fs);

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
