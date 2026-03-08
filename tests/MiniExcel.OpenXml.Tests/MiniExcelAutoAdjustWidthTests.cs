using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using MiniExcelLib.OpenXml.Models;
using MiniExcelLib.Tests.Common.Utils;

namespace MiniExcelLib.OpenXml.Tests;

public class MiniExcelAutoAdjustWidthTests
{
    private readonly OpenXmlExporter _excelExporter =  MiniExcel.Exporters.GetOpenXmlExporter();
    
    [Fact]
    public async Task AutoAdjustWidthThrowsExceptionWithoutFastMode_Async()
    {
        using var file = AutoDeletingPath.Create();
        var path = file.ToString();

        await Assert.ThrowsAsync<InvalidOperationException>(() => _excelExporter.ExportAsync(path, AutoAdjustTestParameters.GetDictionaryTestData(), configuration: new OpenXmlConfiguration
        {
            EnableAutoWidth = true,
        }));
    }

    [Fact]
    public void AutoAdjustWidthThrowsExceptionWithoutFastMode()
    {
        using var file = AutoDeletingPath.Create();
        var path = file.ToString();

        Assert.Throws<InvalidOperationException>(() => _excelExporter.Export(path, AutoAdjustTestParameters.GetDictionaryTestData(), configuration: new OpenXmlConfiguration
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
        await _excelExporter.ExportAsync(path, AutoAdjustTestParameters.GetDictionaryTestData(), configuration: configuration);

        AssertExpectedWidth(path, configuration);
    }

    [Fact]
    public void AutoAdjustWidthEnumerable()
    {
        using var file = AutoDeletingPath.Create();
        var path = file.ToString();

        var configuration = AutoAdjustTestParameters.GetConfiguration();
        var data = AutoAdjustTestParameters.GetDictionaryTestData();
        _excelExporter.Export(path, data, configuration: configuration);

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
            await using var reader = await command.ExecuteReaderAsync();
            await _excelExporter.ExportAsync(path, reader, configuration: configuration, overwriteFile: true);
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
            _excelExporter.Export(path, reader, configuration: configuration);
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

        foreach (object[] row in AutoAdjustTestParameters.GetTestData())
        {
            table.Rows.Add(row);
        }

        var configuration = AutoAdjustTestParameters.GetConfiguration();
        await _excelExporter.ExportAsync(path, table, configuration: configuration);

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

        foreach (object[] row in AutoAdjustTestParameters.GetTestData())
        {
            table.Rows.Add(row);
        }

        var configuration = AutoAdjustTestParameters.GetConfiguration();
        _excelExporter.Export(path, table, configuration: configuration);

        AssertExpectedWidth(path, configuration);
    }

    private static void AssertExpectedWidth(string path, OpenXmlConfiguration configuration)
    {
        using var document = SpreadsheetDocument.Open(path, false);
        var worksheetPart = document.WorkbookPart?.WorksheetParts.First();

        var columns = worksheetPart?.Worksheet.GetFirstChild<Columns>();
        Assert.False(columns is null, "No column width information was written.");
        
        foreach (var column in columns.Elements<Column>())
        {
            var expectedWidth = column.Min?.Value switch
            {
                1 => AutoAdjustTestParameters.Column1MaLen,
                2 => AutoAdjustTestParameters.Column2MaxLen,
                3 => configuration.MinWidth,
                4 => configuration.MaxWidth,
                _ => throw new UnreachableException()
            };
            Assert.Equal(ExcelColumnWidth.GetWidthFromTextLength(expectedWidth), Math.Round(column.Width!.Value, 8));
        }
    }

    private static class AutoAdjustTestParameters
    {
        internal const int Column1MaLen = 32;
        internal const int Column2MaxLen = 16;
        private const int Column3MaxLen = 2;
        private const int Column4MaxLen = 100;

        public static List<string[]> GetTestData() => [
        [
            new('1', Column1MaLen), 
            new('2', Column2MaxLen / 2),
            new('3', Column3MaxLen / 2),
            new('4', Column4MaxLen)
        ],
        [
            new('1', Column1MaLen / 2), 
            new('2', Column2MaxLen),
            new('3', Column3MaxLen), 
            new('4', Column4MaxLen)
        ] ];

        public static List<Dictionary<string, object>> GetDictionaryTestData() => GetTestData()
            .Select(row => row
                .Select((value, i) => (value, i))
                .ToDictionary(x => $"Column{x.i}", object (x) => x.value))
            .ToList();

        public static OpenXmlConfiguration GetConfiguration() => new()
        {
            EnableAutoWidth = true,
            FastMode = true,
            MaxWidth = 50
        };
    }
}
