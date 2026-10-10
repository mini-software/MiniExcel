using System.Drawing;
using System.Xml.Linq;
using ClosedXML.Excel;
using MiniExcelLib.OpenXml.Models;

namespace MiniExcelLib.OpenXml.Tests.CellValues;

public class OpenXmlCellValueAsyncTests
{
    private readonly OpenXmlExporter _exporter = MiniExcelV2.Exporters.GetOpenXmlExporter();

    private class FormattedRow
    {
        [MiniExcelFormat("0.000")]
        public object? Amount { get; set; }

        public object? When { get; set; }
    }

    private async Task<XLWorkbook> ExportAsync(object value, OpenXmlConfiguration? configuration = null)
    {
        var stream = new MemoryStream();
        await _exporter.ExportAsync(stream, value, configuration: configuration);
        stream.Position = 0;
        return new XLWorkbook(stream);
    }

    [Fact]
    public async Task FormulaIsWrittenWithItsCachedResultAsync()
    {
        Dictionary<string, object?>[] rows =
        [
            new() { ["A"] = 3, ["B"] = new OpenXmlCellValue(6) { Formula = "A2*2" }, ["C"] = new OpenXmlCellValue("x3") { Formula = "=\"x\"&A2" } }
        ];

        using var workbook = await ExportAsync(rows);
        var sheet = workbook.Worksheet(1);

        Assert.Equal("A2*2", sheet.Cell("B2").FormulaA1);
        Assert.Equal(6, sheet.Cell("B2").CachedValue.GetNumber());
        Assert.Equal("\"x\"&A2", sheet.Cell("C2").FormulaA1);
        Assert.Equal("x3", sheet.Cell("C2").CachedValue.GetText());
        Assert.False(sheet.Cell("A2").HasFormula);
    }

    [Fact]
    public async Task FontColorAndNumberFormatApplyToTheCellOnlyAsync()
    {
        Dictionary<string, object?>[] rows =
        [
            new() { ["Value"] = new OpenXmlCellValue(0.25) { NumberFormat = "0.0%", FontColor = Color.Red } },
            new() { ["Value"] = new OpenXmlCellValue("grey") { FontColor = Color.Gray } },
            new() { ["Value"] = 1.5 }
        ];

        using var workbook = await ExportAsync(rows);
        var sheet = workbook.Worksheet(1);

        Assert.Equal(0.25, sheet.Cell("A2").GetDouble());
        Assert.Equal("0.0%", sheet.Cell("A2").Style.NumberFormat.Format);
        Assert.Equal(Color.Red.ToArgb(), sheet.Cell("A2").Style.Font.FontColor.Color.ToArgb());
        Assert.Equal("grey", sheet.Cell("A3").GetString());
        Assert.Equal(Color.Gray.ToArgb(), sheet.Cell("A3").Style.Font.FontColor.Color.ToArgb());
        Assert.Equal(1.5, sheet.Cell("A4").GetDouble());
        Assert.NotEqual(Color.Red.ToArgb(), sheet.Cell("A4").Style.Font.FontColor.Color.ToArgb());
    }

    [Fact]
    public async Task FontColorKeepsTheColumnFormatAsync()
    {
        Dictionary<string, object?>[] rows =
        [
            new() { ["Amount"] = new OpenXmlCellValue(1.23456) { FontColor = Color.Blue } },
            new() { ["Amount"] = 2.5 }
        ];
        var configuration = new OpenXmlConfiguration { DynamicColumns = [new DynamicExcelColumn("Amount") { Format = "0.000" }] };

        using var workbook = await ExportAsync(rows, configuration);
        var sheet = workbook.Worksheet(1);

        Assert.Equal(1.23456, sheet.Cell("A2").GetDouble());
        Assert.Equal("0.000", sheet.Cell("A2").Style.NumberFormat.Format);
        Assert.Equal(Color.Blue.ToArgb(), sheet.Cell("A2").Style.Font.FontColor.Color.ToArgb());
        Assert.Equal("0.000", sheet.Cell("A3").Style.NumberFormat.Format);
    }

    [Fact]
    public async Task WrappedValuesInTypedRowsUseTheirOwnTypeAsync()
    {
        FormattedRow[] rows =
        [
            new() { Amount = new OpenXmlCellValue(1.23456) { FontColor = Color.Blue }, When = new OpenXmlCellValue(new DateTime(2024, 5, 6)) { FontColor = Color.Blue } }
        ];

        using var workbook = await ExportAsync(rows);
        var sheet = workbook.Worksheet(1);

        Assert.Equal(1.23456, sheet.Cell("A2").GetDouble());
        Assert.Equal("0.000", sheet.Cell("A2").Style.NumberFormat.Format);
        Assert.Equal(new DateTime(2024, 5, 6), sheet.Cell("B2").GetDateTime());
        Assert.Equal(14, sheet.Cell("B2").Style.NumberFormat.NumberFormatId);
        Assert.Equal(Color.Blue.ToArgb(), sheet.Cell("B2").Style.Font.FontColor.Color.ToArgb());
    }

    [Fact]
    public async Task CellsWithTheSameStyleShareOneFormatAsync()
    {
        var rows = Enumerable.Range(1, 50)
            .Select(i => new Dictionary<string, object?> { ["Value"] = new OpenXmlCellValue(i) { NumberFormat = "0.00", FontColor = Color.Gray } })
            .ToArray();

        var stream = new MemoryStream();
        await _exporter.ExportAsync(stream, rows);
        stream.Position = 0;

        using var archive = new ZipArchive(stream);
        using var reader = new StreamReader(archive.GetEntry("xl/styles.xml")!.Open());
        var styles = XDocument.Parse(await reader.ReadToEndAsync());
        XNamespace ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

        Assert.Single(styles.Descendants(ns + "numFmt"), e => (string?)e.Attribute("formatCode") == "0.00");
        Assert.Single(styles.Descendants(ns + "font"), e => (string?)e.Element(ns + "color")?.Attribute("rgb") == "FF808080");
    }

    [Fact]
    public async Task NullValueKeepsItsStyleAsync()
    {
        Dictionary<string, object?>[] rows = [new() { ["Value"] = new OpenXmlCellValue(null) { FontColor = Color.Red } }];

        using var workbook = await ExportAsync(rows);
        var cell = workbook.Worksheet(1).Cell("A2");

        Assert.True(cell.IsEmpty());
        Assert.Equal(Color.Red.ToArgb(), cell.Style.Font.FontColor.Color.ToArgb());
    }

    [Fact]
    public async Task MinimalStylesSupportCellValuesAsync()
    {
        Dictionary<string, object?>[] rows = [new() { ["Value"] = new OpenXmlCellValue(0.5) { NumberFormat = "0%", FontColor = Color.Green } }];

        using var workbook = await ExportAsync(rows, new OpenXmlConfiguration { TableStyles = TableStyles.None });
        var cell = workbook.Worksheet(1).Cell("A2");

        Assert.Equal("0%", cell.Style.NumberFormat.Format);
        Assert.Equal(Color.Green.ToArgb(), cell.Style.Font.FontColor.Color.ToArgb());
    }

    [Fact]
    public async Task InvalidNumberFormatThrowsAsync()
    {
        Dictionary<string, object?>[] rows = [new() { ["Value"] = new OpenXmlCellValue(1) { NumberFormat = "0.0\"" } }];

        await Assert.ThrowsAsync<InvalidOperationException>(() => _exporter.ExportAsync(new MemoryStream(), rows));
    }
}
