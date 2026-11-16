using ClosedXML.Excel;
using MiniExcelLib.Core.Enums;
using MiniExcelLib.OpenXml.Styles;

namespace MiniExcelLib.OpenXml.Tests.Tables;

public class MiniExcelOpenXmlTableExportAsyncTests
{
    private readonly OpenXmlExporter _exporter = MiniExcelV2.Exporters.GetOpenXmlExporter();
    private readonly OpenXmlImporter _importer = MiniExcelV2.Importers.GetOpenXmlImporter();

    private class Weld
    {
        public int Id { get; set; }
        public string? Name { get; set; }
        public double Depth { get; set; }
    }

    private class IndexedWeld
    {
        [MiniExcelColumnIndex(0)]
        public int Id { get; set; }

        [MiniExcelColumnIndex(2)]
        public double Depth { get; set; }
    }

    private static readonly Weld[] Welds =
    [
        new() { Id = 1, Name = "first", Depth = 0.1 },
        new() { Id = 2, Name = "second", Depth = 0.2 }
    ];

    private async Task<MemoryStream> ExportAsync(object value, OpenXmlConfiguration configuration, bool printHeader = true)
    {
        var stream = new MemoryStream();
        await _exporter.ExportAsync(stream, value, printHeader, configuration: configuration);
        stream.Position = 0;
        return stream;
    }

    [Fact]
    public async Task DataIsWrittenAsATableInTheChosenStyleAsync()
    {
        var configuration = new OpenXmlConfiguration
        {
            TableStyles = TableStyles.None,
            Table = new OpenXmlTableOptions { Name = "Welds", Style = "TableStyleMedium15", ShowFilterButton = false }
        };

        using var stream = await ExportAsync(Welds, configuration);
        using var workbook = new XLWorkbook(stream);
        var table = workbook.Worksheet(1).Tables.Single();

        Assert.Equal("Welds", table.Name);
        Assert.Equal("A1:C3", table.RangeAddress.ToString());
        Assert.Equal(["Id", "Name", "Depth"], table.Fields.Select(f => f.Name));
        Assert.Equal(XLTableTheme.TableStyleMedium15, table.Theme);
        Assert.False(table.ShowAutoFilter);
        Assert.True(table.ShowRowStripes);
        Assert.Null(workbook.Worksheet(1).AutoFilter.Range);

        stream.Position = 0;
        var rows = await _importer.QueryTableAsync(stream, tableName: "Welds").ToListAsync();
        Assert.Equal(2, rows.Count);
        Assert.Equal("second", (string)rows[1].Name);
    }

    [Fact]
    public async Task HeaderStyleAppliesWithoutCellStylesAsync()
    {
        var configuration = new OpenXmlConfiguration
        {
            TableStyles = TableStyles.None,
            StyleOptions = new OpenXmlStyleOptions
            {
                HeaderStyle = new OpenXmlHeaderStyle
                {
                    WrapText = true,
                    HorizontalAlignment = HorizontalCellAlignment.Center,
                    VerticalAlignment = VerticalCellAlignment.Center,
                    RowHeight = 45,
                }
            },
            Table = new OpenXmlTableOptions { Style = "TableStyleMedium15" }
        };

        using var stream = await ExportAsync(Welds, configuration);
        using var workbook = new XLWorkbook(stream);
        var sheet = workbook.Worksheet(1);
        var header = sheet.Cell("A1").Style;

        Assert.True(header.Alignment.WrapText);
        Assert.Equal(XLAlignmentHorizontalValues.Center, header.Alignment.Horizontal);
        Assert.Equal(XLAlignmentVerticalValues.Center, header.Alignment.Vertical);
        Assert.Equal(XLFillPatternValues.None, header.Fill.PatternType);
        Assert.Equal(45, sheet.Row(1).Height);
        Assert.False(sheet.Cell("A2").Style.Alignment.WrapText);
    }

    [Fact]
    public async Task EachSheetGetsItsOwnTableAsync()
    {
        var sheets = new Dictionary<string, object> { ["First"] = Welds, ["Second"] = Welds.Take(1) };

        using var stream = await ExportAsync(sheets, new OpenXmlConfiguration { Table = new OpenXmlTableOptions() });
        using var workbook = new XLWorkbook(stream);

        Assert.Equal("Table1", workbook.Worksheet("First").Tables.Single().Name);
        Assert.Equal("A1:C3", workbook.Worksheet("First").Tables.Single().RangeAddress.ToString());
        Assert.Equal("Table2", workbook.Worksheet("Second").Tables.Single().Name);
        Assert.Equal("A1:C2", workbook.Worksheet("Second").Tables.Single().RangeAddress.ToString());
        Assert.True(workbook.Worksheet("First").Tables.Single().ShowAutoFilter);
    }

    [Fact]
    public async Task HeaderWithoutDataStillMakesATableAsync()
    {
        using var stream = await ExportAsync(Array.Empty<Weld>(), new OpenXmlConfiguration { Table = new OpenXmlTableOptions() });
        using var workbook = new XLWorkbook(stream);

        Assert.Equal("A1:C2", workbook.Worksheet(1).Tables.Single().RangeAddress.ToString());
    }

    [Fact]
    public async Task TableWithoutHeaderRowAsync()
    {
        using var stream = await ExportAsync(Welds, new OpenXmlConfiguration { Table = new OpenXmlTableOptions() }, printHeader: false);
        using var workbook = new XLWorkbook(stream);
        var table = workbook.Worksheet(1).Tables.Single();

        Assert.False(table.ShowHeaderRow);
        Assert.Equal("A1:C2", table.RangeAddress.ToString());
    }

    [Fact]
    public async Task ColumnGapsAreRejectedAsync()
    {
        IndexedWeld[] rows = [new() { Id = 1, Depth = 0.1 }];

        await Assert.ThrowsAsync<NotSupportedException>(() => ExportAsync(rows, new OpenXmlConfiguration { Table = new OpenXmlTableOptions() }));
    }

    [Fact]
    public async Task InvalidTableNameIsRejectedAsync()
    {
        var configuration = new OpenXmlConfiguration { Table = new OpenXmlTableOptions { Name = "My Welds" } };

        await Assert.ThrowsAsync<ArgumentException>(() => ExportAsync(Welds, configuration));
    }

    [Fact]
    public async Task InsertingASheetWithATableIsRejectedAsync()
    {
        using var stream = await ExportAsync(Welds, new OpenXmlConfiguration());
        var configuration = new OpenXmlConfiguration { FastMode = true, Table = new OpenXmlTableOptions() };

        await Assert.ThrowsAsync<NotSupportedException>(() => _exporter.InsertSheetAsync(stream, Welds, "Second", configuration: configuration));
    }
}
