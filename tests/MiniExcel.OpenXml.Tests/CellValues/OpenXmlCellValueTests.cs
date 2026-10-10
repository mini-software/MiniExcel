using System.Drawing;
using ClosedXML.Excel;
using MiniExcelLib.OpenXml.Models;

namespace MiniExcelLib.OpenXml.Tests.CellValues;

public class OpenXmlCellValueTests
{
    private readonly OpenXmlExporter _exporter = MiniExcelV2.Exporters.GetOpenXmlExporter();

    [Fact]
    public void FormulaFontColorAndNumberFormatAreWritten()
    {
        Dictionary<string, object?>[] rows =
        [
            new() { ["A"] = 3, ["B"] = new OpenXmlCellValue(0.6) { Formula = "A2/5", NumberFormat = "0.0%", FontColor = Color.Gray } }
        ];

        var stream = new MemoryStream();
        _exporter.Export(stream, rows);
        stream.Position = 0;

        using var workbook = new XLWorkbook(stream);
        var cell = workbook.Worksheet(1).Cell("B2");

        Assert.Equal("A2/5", cell.FormulaA1);
        Assert.Equal(0.6, cell.CachedValue.GetNumber());
        Assert.Equal("0.0%", cell.Style.NumberFormat.Format);
        Assert.Equal(Color.Gray.ToArgb(), cell.Style.Font.FontColor.Color.ToArgb());
    }
}
