using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using MiniExcelLib.Core.OpenXml.Models;

namespace MiniExcelLib.Tests;

public class MiniExcelAutoAdjustWidthTests
{
    private readonly MiniExcelExporter _exporter = MiniExcel.GetExporter();
    
    [Fact]
    public async Task AutoAdjustWidthThrowsExceptionWithoutFastMode_Async()
    {
        using var file = AutoDeletingPath.Create();
        var path = file.ToString();

        await Assert.ThrowsAsync<InvalidOperationException>(() => _exporter.ExportXlsxAsync(path, AutoAdjustTestParameters.GetDictionaryTestData(), configuration: new OpenXmlConfiguration
        {
            EnableAutoWidth = true,
        }));
    }

    [Fact]
    public void AutoAdjustWidthThrowsExceptionWithoutFastMode()
    {
        using var file = AutoDeletingPath.Create();
        var path = file.ToString();

        Assert.Throws<InvalidOperationException>(() => _exporter.ExportXlsx(path, AutoAdjustTestParameters.GetDictionaryTestData(), configuration: new OpenXmlConfiguration
        {
            EnableAutoWidth = true,
        }));
    }

    [Fact]
    public async Task AutoAdjustWidthEnumerable_Async()
    {
        using var file = AutoDeletingPath.Create();
        var path = file.ToString();

        var configuration = AutoAdjustTestParameters.GetConfiguration();
        await _exporter.ExportXlsxAsync(path, AutoAdjustTestParameters.GetDictionaryTestData(), configuration: configuration);

        AssertExpectedWidth(path, configuration);
    }

    [Fact]
    public void AutoAdjustWidthEnumerable()
    {
        using var file = AutoDeletingPath.Create();
        var path = file.ToString();

        var configuration = AutoAdjustTestParameters.GetConfiguration();
        _exporter.ExportXlsx(path, AutoAdjustTestParameters.GetDictionaryTestData(), configuration: configuration);

        AssertExpectedWidth(path, configuration);
    }

    [Fact]
    public async Task AutoAdjustWidthDataReader_Async()
    {
        using var file = AutoDeletingPath.Create();
        var path = file.ToString();

        var configuration = AutoAdjustTestParameters.GetConfiguration();

        await using (var connection = Db.GetConnection("Data Source=:memory:"))
        {
            await using var command = new SQLiteCommand(Db.GenerateDummyQuery(AutoAdjustTestParameters.GetDictionaryTestData()), connection);
            connection.Open();
            await using var reader = command.ExecuteReader();
            await _exporter.ExportXlsxAsync(path, reader, configuration: configuration);
        }

        AssertExpectedWidth(path, configuration);
    }

    [Fact]
    public void AutoAdjustWidthDataReader()
    {
        using var file = AutoDeletingPath.Create();
        var path = file.ToString();

        var configuration = AutoAdjustTestParameters.GetConfiguration();

        using (var connection = Db.GetConnection("Data Source=:memory:"))
        {
            using var command = new SQLiteCommand(Db.GenerateDummyQuery(AutoAdjustTestParameters.GetDictionaryTestData()), connection);
            connection.Open();
            using var reader = command.ExecuteReader();
            _exporter.ExportXlsx(path, reader, configuration: configuration);
        }

        AssertExpectedWidth(path, configuration);
    }


    [Fact]
    public async Task AutoAdjustWidthDataTable_Async()
    {
        using var file = AutoDeletingPath.Create();
        var path = file.ToString();

        var table = new DataTable();
        table.Columns.Add("Column1", typeof(string));
        table.Columns.Add("Column2", typeof(string));
        table.Columns.Add("Column3", typeof(string));
        table.Columns.Add("Column4", typeof(string));

        foreach (var row in AutoAdjustTestParameters.GetTestData())
        {
            table.Rows.Add(row);
        }

        var configuration = AutoAdjustTestParameters.GetConfiguration();
        await _exporter.ExportXlsxAsync(path, table, configuration: configuration);

        AssertExpectedWidth(path, configuration);
    }

    [Fact]
    public void AutoAdjustWidthDataTable()
    {
        using var file = AutoDeletingPath.Create();
        var path = file.ToString();

        var table = new DataTable();
        table.Columns.Add("Column1", typeof(string));
        table.Columns.Add("Column2", typeof(string));
        table.Columns.Add("Column3", typeof(string));
        table.Columns.Add("Column4", typeof(string));

        foreach (var row in AutoAdjustTestParameters.GetTestData())
        {
            table.Rows.Add(row);
        }

        var configuration = AutoAdjustTestParameters.GetConfiguration();
        _exporter.ExportXlsx(path, table, configuration: configuration);

        AssertExpectedWidth(path, configuration);
    }

    private static void AssertExpectedWidth(string path, OpenXmlConfiguration configuration)
    {
        using var document = SpreadsheetDocument.Open(path, false);
        var worksheetPart = document.WorkbookPart.WorksheetParts.First();

        var columns = worksheetPart.Worksheet.GetFirstChild<Columns>();
        Assert.False(columns is null, "No column width information was written.");
        foreach (var column in columns.Elements<Column>())
        {
            var expectedWidth = column.Min.Value switch
            {
                1 => ExcelWidthCollection.GetApproximateTextWidth(AutoAdjustTestParameters.column1MaxStringLength),
                2 => ExcelWidthCollection.GetApproximateTextWidth(AutoAdjustTestParameters.column2MaxStringLength),
                3 => configuration.MinWidth,
                4 => configuration.MaxWidth,
                _ => throw new Exception("Unexpected column"),
            };

            Assert.Equal(expectedWidth, column.Width?.Value);
        }
    }

    private static class AutoAdjustTestParameters
    {
        public const int column1MaxStringLength = 32;
        public const int column2MaxStringLength = 16;
        public const int column3MaxStringLength = 2;
        public const int column4MaxStringLength = 100;
        public const int minStringLength = 8;
        public const int maxStringLength = 50;

        public static List<string[]> GetTestData() =>
        [
            new string[]
            {
                new('1', column1MaxStringLength), new('2', column2MaxStringLength / 2),
                new('3', column3MaxStringLength / 2), new('4', column1MaxStringLength)
            },
            new string[]
            {
                new('1', column1MaxStringLength / 2), new('2', column2MaxStringLength),
                new('3', column3MaxStringLength), new('4', column4MaxStringLength)
            }
        ];

        public static List<Dictionary<string, object>> GetDictionaryTestData() => GetTestData()
            .Select(row => row
                .Select((value, i) => (value, i))
                .ToDictionary(x => $"Column{x.i}", object (x) => x.value))
            .ToList();

        public static OpenXmlConfiguration GetConfiguration() => new()
        {
            EnableAutoWidth = true,
            FastMode = true,
            MinWidth = ExcelWidthCollection.GetApproximateTextWidth(minStringLength),
            MaxWidth = ExcelWidthCollection.GetApproximateTextWidth(maxStringLength)
        };
    }
}