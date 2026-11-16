using ClosedXML.Excel;

namespace MiniExcelLib.OpenXml.Tests.Tables;

public class MiniExcelOpenXmlTableExportTests
{
    private readonly OpenXmlExporter _exporter = MiniExcelV2.Exporters.GetOpenXmlExporter();

    [Fact]
    public void DataIsWrittenAsATable()
    {
        var rows = new[] { new { Id = 1, Name = "first" }, new { Id = 2, Name = "second" } };
        var configuration = new OpenXmlConfiguration { Table = new OpenXmlTableOptions { Style = "TableStyleLight9" } };

        var stream = new MemoryStream();
        _exporter.Export(stream, rows, configuration: configuration);
        stream.Position = 0;

        using var workbook = new XLWorkbook(stream);
        var table = workbook.Worksheet(1).Tables.Single();
        Assert.Equal("Table1", table.Name);
        Assert.Equal("A1:B3", table.RangeAddress.ToString());
        Assert.Equal(XLTableTheme.TableStyleLight9, table.Theme);
    }
}
