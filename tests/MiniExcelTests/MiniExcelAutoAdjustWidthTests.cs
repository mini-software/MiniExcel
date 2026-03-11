using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using MiniExcelLibs.OpenXml;
using MiniExcelLibs.Tests.Utils;
using System.Data;
using System.Data.SQLite;
using System.Diagnostics;
using Xunit;

namespace MiniExcelLibs.Tests;

public class MiniExcelAutoAdjustWidthTests
{
    [Fact]
    public async Task AutoAdjustWidthThrowsExceptionWithoutFastMode_Async()
    {
        using var file = AutoDeletingPath.Create();
        var path = file.ToString();

        await Assert.ThrowsAsync<InvalidOperationException>(() => MiniExcel.SaveAsAsync(path, AutoAdjustTestParameters.GetDictionaryTestData(), configuration: new OpenXmlConfiguration
        {
            EnableAutoWidth = true,
        }));
    }

    [Fact]
    public void AutoAdjustWidthThrowsExceptionWithoutFastMode()
    {
        using var file = AutoDeletingPath.Create();
        var path = file.ToString();

        Assert.Throws<InvalidOperationException>(() => MiniExcel.SaveAs(path, AutoAdjustTestParameters.GetDictionaryTestData(), configuration: new OpenXmlConfiguration
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
        await MiniExcel.SaveAsAsync(path, AutoAdjustTestParameters.GetDictionaryTestData(), configuration: configuration);

        AssertExpectedWidth(path, configuration);
    }

    [Fact]
    public void AutoAdjustWidthEnumerable()
    {
        using var file = AutoDeletingPath.Create();
        var path = file.ToString();

        var configuration = AutoAdjustTestParameters.GetConfiguration();
        MiniExcel.SaveAs(path, AutoAdjustTestParameters.GetDictionaryTestData(), configuration: configuration);

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
            await MiniExcel.SaveAsAsync(path, reader, configuration: configuration);
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
            MiniExcel.SaveAs(path, reader, configuration: configuration);
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
        await MiniExcel.SaveAsAsync(path, table, configuration: configuration);

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
        MiniExcel.SaveAs(path, table, configuration: configuration);

        AssertExpectedWidth(path, configuration);
    }

    private static void AssertExpectedWidth(string path, OpenXmlConfiguration configuration)
    {
        using var document = SpreadsheetDocument.Open(path, false);
        var worksheetPart = document.WorkbookPart.WorksheetParts.First();

        var columns = worksheetPart.Worksheet.GetFirstChild<Columns>();
        Assert.False(columns == null, "No column width information was written.");
        
        foreach (var column in columns.Elements<Column>())
        {
            var expectedWidth = column.Min?.Value switch
            {
                1 => AutoAdjustTestParameters.Column1MaxLen,
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
        internal const int Column1MaxLen = 32;
        internal const int Column2MaxLen = 16;
        private const int Column3MaxLen = 2;
        private const int Column4MaxLen = 100;

        public static List<string[]> GetTestData() =>
        [
            [
                new('1', Column1MaxLen), 
                new('2', Column2MaxLen / 2),
                new('3', Column3MaxLen / 2),
                new('4', Column4MaxLen)
            ],
            [
                new('1', Column1MaxLen / 2), 
                new('2', Column2MaxLen),
                new('3', Column3MaxLen), 
                new('4', Column4MaxLen)
            ]
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
            MaxWidth = 50
        };
    }
}