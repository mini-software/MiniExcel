using System.Xml.Linq;
using ClosedXML.Excel;
using MiniExcelLib.Core.Helpers;

namespace MiniExcelLib.OpenXml.Tests.Cells;

public class QueryCellsWithMetadataAsyncTests
{
    private readonly OpenXmlImporter _excelImporter = MiniExcelV2.Importers.GetOpenXmlImporter();
    private readonly OpenXmlExporter _excelExporter = MiniExcelV2.Exporters.GetOpenXmlExporter();

    internal static MemoryStream CreateWorkbook()
    {
        using var workbook = new XLWorkbook();
        var sheet = workbook.AddWorksheet("Sheet1");
        sheet.Cell("A1").Value = 1;
        sheet.Cell("A2").Value = 2;
        sheet.Cell("A2").Style.Font.FontColor = XLColor.Red;
        sheet.Cell("A3").Value = 3;
        sheet.Cell("A3").Style.Font.FontColor = XLColor.FromTheme(XLThemeColor.Text1);
        sheet.Cell("A4").Value = "text";
        sheet.Cell("A4").Style.Font.FontColor = XLColor.FromTheme(XLThemeColor.Accent2, 0.4);

        var stream = new MemoryStream();
        workbook.SaveAs(stream);

        // Write shared formula records and cached values explicitly; ClosedXML saves ordinary formulas.
        using (var archive = new ZipArchive(stream, ZipArchiveMode.Update, leaveOpen: true))
        {
            var entry = archive.GetEntry("xl/worksheets/sheet1.xml")!;
            XDocument document;
            using (var input = entry.Open())
                document = XDocument.Load(input);
            XNamespace ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
            var rows = document.Descendants(ns + "row").ToArray();
            rows[0].Add(
                new XElement(ns + "c", new XAttribute("r", "B1"),
                    new XElement(ns + "f", new XAttribute("t", "shared"),
                        new XAttribute("si", "0"), new XAttribute("ref", "B1:B3"), "A1*2"),
                    new XElement(ns + "v", "2")),
                new XElement(ns + "c", new XAttribute("r", "C1"), new XAttribute("t", "str"),
                    new XElement(ns + "f", "\"x\"&A1"), new XElement(ns + "v", "x1")));
            for (var index = 1; index < 3; index++)
            {
                rows[index].Add(new XElement(ns + "c", new XAttribute("r", $"B{index + 1}"),
                    new XElement(ns + "f", new XAttribute("t", "shared"), new XAttribute("si", "0")),
                    new XElement(ns + "v", (index + 1) * 2)));
            }

            entry.Delete();
            using (var output = archive.CreateEntry("xl/worksheets/sheet1.xml").Open())
                document.Save(output);

            var stylesEntry = archive.GetEntry("xl/styles.xml")!;
            XDocument styles;
            using (var input = stylesEntry.Open())
                styles = XDocument.Load(input);
            styles.Descendants(ns + "font").First().Element(ns + "color")?.Remove();
            stylesEntry.Delete();
            using var stylesOutput = archive.CreateEntry("xl/styles.xml").Open();
            styles.Save(stylesOutput);
        }
        stream.Position = 0;
        return stream;
    }

    [Fact]
    public async Task QueryCellsWithMetadata_ReturnsValuesAndReferencesAsync()
    {
        using var stream = CreateWorkbook();
        var cells = await _excelImporter.QueryCellsWithMetadataAsync(stream).ToDictionaryAsync(c => c.Reference);

        Assert.Equal(1d, cells["A1"].Value);
        Assert.Equal(1, cells["A1"].RowIndex);
        Assert.Equal(1, cells["A1"].ColumnIndex);
        Assert.Equal("text", cells["A4"].Value);
        Assert.Equal(4, cells["A4"].RowIndex);
        Assert.Equal(3, cells["C1"].ColumnIndex);
    }

    [Fact]
    public async Task QueryCellsWithMetadata_ReportsFormulasAsync()
    {
        using var stream = CreateWorkbook();
        var cells = await _excelImporter.QueryCellsWithMetadataAsync(stream).ToDictionaryAsync(c => c.Reference);

        Assert.False(cells["A1"].HasFormula);
        Assert.Null(cells["A1"].Formula);

        Assert.True(cells["B1"].HasFormula);
        Assert.Equal("A1*2", cells["B1"].Formula);
        Assert.Equal(2d, cells["B1"].Value);

        // Cells sharing the formula of B1 carry no formula text of their own
        Assert.True(cells["B3"].HasFormula);
        Assert.Null(cells["B3"].Formula);
        Assert.Equal(6d, cells["B3"].Value);

        Assert.True(cells["C1"].HasFormula);
        Assert.Equal("\"x\"&A1", cells["C1"].Formula);
        Assert.Equal("x1", cells["C1"].Value);
    }

    [Fact]
    public async Task QueryCellsWithMetadata_ReportsFontColorsAsync()
    {
        using var stream = CreateWorkbook();
        var cells = await _excelImporter.QueryCellsWithMetadataAsync(stream).ToDictionaryAsync(c => c.Reference);

        Assert.Equal("FFFF0000", cells["A2"].FontColor?.Rgb);
        Assert.Null(cells["A2"].FontColor?.Theme);

        Assert.Equal(1, cells["A3"].FontColor?.Theme);
        Assert.Null(cells["A3"].FontColor?.Rgb);
        Assert.Equal(0, cells["A3"].FontColor?.Tint);

        Assert.Equal(5, cells["A4"].FontColor?.Theme);
        Assert.Equal(0.4, cells["A4"].FontColor!.Tint, 6);

        Assert.Null(cells["A1"].FontColor?.Rgb);
        Assert.Null(cells["A1"].FontColor?.Theme);
    }

    [Fact]
    public async Task QueryCellsWithMetadata_ReportsNumberFormatsAsync()
    {
        using var workbook = new XLWorkbook();
        var sheet = workbook.AddWorksheet("Sheet1");
        sheet.Cell("A1").Value = 0.233;
        sheet.Cell("A1").Style.NumberFormat.Format = "0.0%";
        sheet.Cell("A2").Value = 5;

        await using var stream = new MemoryStream();
        workbook.SaveAs(stream);
        stream.Position = 0;

        var cells = await _excelImporter.QueryCellsWithMetadataAsync(stream).ToDictionaryAsync(c => c.Reference);

        Assert.Equal("0.0%", cells["A1"].NumberFormat);
        Assert.Equal(0.233, cells["A1"].Value);
        Assert.Null(cells["A2"].NumberFormat);
    }

    [Fact]
    public async Task QueryCellsWithMetadata_SelectsSheetByNameAsync()
    {
        using var stream = new MemoryStream();
        var sheets = new Dictionary<string, object>
        {
            ["First"] = new[] { new { Value = "first" } },
            ["Second"] = new[] { new { Value = "second" } }
        };
        await _excelExporter.ExportAsync(stream, sheets);
        stream.Position = 0;

        var cells = await _excelImporter.QueryCellsWithMetadataAsync(stream, "Second").ToListAsync();

        Assert.Equal(["A1", "A2"], cells.Select(c => c.Reference));
        Assert.Equal(["Value", "second"], cells.Select(c => c.Value));
        Assert.All(cells, c => Assert.False(c.HasFormula));
    }

    [Fact]
    public async Task QueryCellsWithMetadata_RejectsOutOfRangeCellReferenceAsync()
    {
        using var stream = new MemoryStream();
        await _excelExporter.ExportAsync(stream, new[] { new { Value = "value" } });

        using (var archive = new ZipArchive(stream, ZipArchiveMode.Update, leaveOpen: true))
        {
            var entry = archive.GetEntry("xl/worksheets/sheet1.xml")!;

            string xml;
            using (var reader = new StreamReader(await entry.OpenAsync()))
            {
                xml = await reader.ReadToEndAsync();
            }

            Assert.Contains("r=\"A2\"", xml);
            entry.Delete();

            var newEntry = archive.CreateEntry("xl/worksheets/sheet1.xml");
            await using var entryStream = await newEntry.OpenAsync();
            await using var writer = new StreamWriter(entryStream);

            await writer.WriteAsync(xml.Replace("r=\"A2\"", "r=\"XFE2\""));
        }
        stream.Position = 0;

        var exception = await Assert.ThrowsAsync<InvalidDataException>(() => _excelImporter.QueryCellsWithMetadataAsync(stream).ToListAsync().AsTask());
        Assert.Contains("XFE2", exception.Message);
    }
}
