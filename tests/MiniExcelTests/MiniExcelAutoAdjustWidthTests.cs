﻿using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using MiniExcelLibs.OpenXml;
using MiniExcelLibs.Tests.Utils;
using System.Data;
using System.Data.SQLite;
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
            var expectedWidth = column.Min.Value switch
            {
                1 => ExcelWidthCollection.GetApproximateRequiredCalibriWidth(AutoAdjustTestParameters.column1MaxStringLength),
                2 => ExcelWidthCollection.GetApproximateRequiredCalibriWidth(AutoAdjustTestParameters.column2MaxStringLength),
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
            MinWidth = ExcelWidthCollection.GetApproximateRequiredCalibriWidth(minStringLength),
            MaxWidth = ExcelWidthCollection.GetApproximateRequiredCalibriWidth(maxStringLength)
        };
    }
}