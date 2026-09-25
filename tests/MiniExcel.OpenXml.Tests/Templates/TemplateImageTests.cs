using System.Xml.Linq;
using ClosedXML.Excel;
using MiniExcelLib.Tests.Common.Utils;
using OfficeOpenXml.Drawing;

namespace MiniExcelLib.OpenXml.Tests.Templates;

/// <summary>
/// Template image support: a byte[] property resolved from a template placeholder must be rendered
/// as an embedded image, consistently with the SaveAs pipeline (issue #972 / #604).
/// </summary>
public class TemplateImageTests(ITestOutputHelper output)
{
    private static readonly XNamespace SpreadsheetNs = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
    private static readonly XNamespace DrawingNs = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing";

    private readonly OpenXmlTemplater _templater = MiniExcelV2.Templaters.GetOpenXmlTemplater();
    private readonly ITestOutputHelper _output = output;

    private static byte[] TestPng() => File.ReadAllBytes(PathHelper.GetFile("xlsx/Issue327/TestIssue327.png"));

    private static string GetSheetXml(string xlsxPath, int sheetIndex = 1)
    {
        using var zip = ZipFile.OpenRead(xlsxPath);
        var entry = zip.GetEntry($"xl/worksheets/sheet{sheetIndex}.xml");
        Assert.NotNull(entry);
        using var reader = new StreamReader(entry!.Open());
        return reader.ReadToEnd();
    }

    private static IReadOnlyList<string> GetMediaEntries(string xlsxPath)
    {
        using var zip = ZipFile.OpenRead(xlsxPath);
        return zip.Entries
            .Where(e => e.FullName.StartsWith("xl/media/", StringComparison.OrdinalIgnoreCase))
            .Select(e => e.FullName)
            .ToList();
    }

    private static void AssertPackageIsValid(string xlsxPath)
    {
        using var zip = ZipFile.OpenRead(xlsxPath);

        // The worksheet must reference the drawing part.
        using (var sheetStream = zip.GetEntry("xl/worksheets/sheet1.xml")!.Open())
        {
            var sheetDoc = XDocument.Load(sheetStream);
            Assert.NotNull(sheetDoc.Descendants(SpreadsheetNs + "drawing").FirstOrDefault());
        }

        // The persisted content types must declare the image extension and the drawing part.
        using (var contentTypesStream = zip.GetEntry("[Content_Types].xml")!.Open())
        {
            var contentTypes = XDocument.Load(contentTypesStream).ToString();
            Assert.Contains("image/png", contentTypes);
            Assert.Contains("/xl/drawings/drawing1.xml", contentTypes);
            Assert.Contains("application/vnd.openxmlformats-officedocument.drawing+xml", contentTypes);
        }
    }

    private static void AssertPackageIsValidAndHasImages(string xlsxPath, int expectedImages)
    {
        AssertPackageIsValid(xlsxPath);

        // A real OpenXML consumer must be able to open the workbook and see the images.
        using (var package = new ExcelPackage(new FileInfo(xlsxPath)))
        {
            Assert.Equal(expectedImages, package.Workbook.Worksheets[0].Drawings.OfType<ExcelPicture>().Count());
        }

        // OPC-level check: the drawing part must resolve to a declared content type, otherwise Excel
        // repairs the file by dropping the drawing (the failure EPPlus does not surface).
        using var opcPackage = Package.Open(xlsxPath, FileMode.Open, FileAccess.Read, FileShare.Read);
        var drawingPart = opcPackage.GetPart(new Uri("/xl/drawings/drawing1.xml", UriKind.Relative));
        Assert.Equal("application/vnd.openxmlformats-officedocument.drawing+xml", drawingPart.ContentType);
    }

    private static (long Width, long Height) GetAnchorSize(string xlsxPath)
    {
        using var zip = ZipFile.OpenRead(xlsxPath);
        using var drawingStream = zip.GetEntry("xl/drawings/drawing1.xml")!.Open();
        var ext = XDocument.Load(drawingStream).Descendants(DrawingNs + "ext").First();
        return ((long)ext.Attribute("cx")!, (long)ext.Attribute("cy")!);
    }

    [Fact]
    public void ScalarByteArray_IsRenderedAsImage()
    {
        using var template = AutoDeletingPath.Create();
        using (var wb = new XLWorkbook())
        {
            var ws = wb.AddWorksheet("Sheet1");
            ws.Cell("A1").Value = "{{Logo}}";
            wb.SaveAs(template.FilePath);
        }

        using var path = AutoDeletingPath.Create();
        _templater.FillTemplate(path.ToString(), template.FilePath, new { Logo = TestPng() });

        var sheet = GetSheetXml(path.ToString());
        _output.WriteLine(sheet);

        Assert.Single(GetMediaEntries(path.ToString()));
        Assert.DoesNotContain("System.Byte[]", sheet);
        Assert.DoesNotContain("{{Logo}}", sheet);
        AssertPackageIsValidAndHasImages(path.ToString(), expectedImages: 1);
    }

    [Fact]
    public void DefaultImageSize_IsUsedWhenRowHasNoExplicitHeight()
    {
        using var template = AutoDeletingPath.Create();
        using (var wb = new XLWorkbook())
        {
            var ws = wb.AddWorksheet("Sheet1");
            ws.Cell("A1").Value = "{{Logo}}";
            wb.SaveAs(template.FilePath);
        }

        using var path = AutoDeletingPath.Create();
        _templater.FillTemplate(path.ToString(), template.FilePath, new { Logo = TestPng() });

        var (width, height) = GetAnchorSize(path.ToString());
        Assert.Equal(609600L, width);
        Assert.Equal(190500L, height);
    }

    [Fact]
    public void RowHeight_ScalesImagePreservingAspectRatio()
    {
        using var template = AutoDeletingPath.Create();
        using (var wb = new XLWorkbook())
        {
            var ws = wb.AddWorksheet("Sheet1");
            ws.Cell("A1").Value = "{{Logo}}";
            ws.Row(1).Height = 30;
            wb.SaveAs(template.FilePath);
        }

        using var path = AutoDeletingPath.Create();
        _templater.FillTemplate(path.ToString(), template.FilePath, new { Logo = TestPng() });

        Assert.Single(GetMediaEntries(path.ToString()));

        // The fixture is 1920x1032; a 30pt row is 30 * 12700 = 381000 EMU tall and the width keeps the ratio.
        var (width, height) = GetAnchorSize(path.ToString());
        Assert.Equal(381000L, height);
        Assert.Equal((long)Math.Round(381000 * (1920.0 / 1032.0)), width);

        AssertPackageIsValidAndHasImages(path.ToString(), expectedImages: 1);
    }

    [Fact]
    public void CollectionRowHeight_ScalesEachImage()
    {
        using var template = AutoDeletingPath.Create();
        using (var wb = new XLWorkbook())
        {
            var ws = wb.AddWorksheet("Sheet1");
            ws.Cell("A1").Value = "{{Products.Name}}";
            ws.Cell("B1").Value = "{{Products.Image}}";
            ws.Row(1).Height = 40;
            wb.SaveAs(template.FilePath);
        }

        using var path = AutoDeletingPath.Create();
        var image = TestPng();
        _templater.FillTemplate(path.ToString(), template.FilePath, new
        {
            Products = new[]
            {
                new { Name = "A", Image = image },
                new { Name = "B", Image = image },
            }
        });

        using var zip = ZipFile.OpenRead(path.ToString());
        using var drawingStream = zip.GetEntry("xl/drawings/drawing1.xml")!.Open();
        var sizes = XDocument.Load(drawingStream).Descendants(DrawingNs + "ext")
            .Select(ext => ((long)ext.Attribute("cx")!, (long)ext.Attribute("cy")!))
            .ToList();

        Assert.Equal(2, sizes.Count);
        Assert.All(sizes, size => Assert.Equal(40L * 12700, size.Item2));
    }

    [Fact]
    public void NestedByteArray_IsRenderedAsImage()
    {
        using var template = AutoDeletingPath.Create();
        using (var wb = new XLWorkbook())
        {
            var ws = wb.AddWorksheet("Sheet1");
            ws.Cell("A1").Value = "{{Company.Logo}}";
            wb.SaveAs(template.FilePath);
        }

        using var path = AutoDeletingPath.Create();
        _templater.FillTemplate(path.ToString(), template.FilePath, new { Company = new { Logo = TestPng() } });

        var sheet = GetSheetXml(path.ToString());
        _output.WriteLine(sheet);

        Assert.Single(GetMediaEntries(path.ToString()));
        Assert.DoesNotContain("System.Byte[]", sheet);
        Assert.DoesNotContain("{{Company.Logo}}", sheet);
        AssertPackageIsValidAndHasImages(path.ToString(), expectedImages: 1);
    }

    [Fact]
    public void CollectionByteArray_IsRenderedAsImagePerRow()
    {
        using var template = AutoDeletingPath.Create();
        using (var wb = new XLWorkbook())
        {
            var ws = wb.AddWorksheet("Sheet1");
            ws.Cell("A1").Value = "{{Products.Name}}";
            ws.Cell("B1").Value = "{{Products.Image}}";
            wb.SaveAs(template.FilePath);
        }

        using var path = AutoDeletingPath.Create();
        var image = TestPng();
        _templater.FillTemplate(path.ToString(), template.FilePath, new
        {
            Products = new[]
            {
                new { Name = "A", Image = image },
                new { Name = "B", Image = image },
                new { Name = "C", Image = image },
            }
        });

        var sheet = GetSheetXml(path.ToString());
        _output.WriteLine(sheet);

        Assert.Equal(3, GetMediaEntries(path.ToString()).Count);
        Assert.DoesNotContain("System.Byte[]", sheet);

        // Each expanded row must have an image anchored to its final row.
        using (var zip = ZipFile.OpenRead(path.ToString()))
        using (var drawingStream = zip.GetEntry("xl/drawings/drawing1.xml")!.Open())
        {
            var drawingXml = XDocument.Load(drawingStream);
            var anchors = drawingXml.Descendants(DrawingNs + "oneCellAnchor").ToList();
            Assert.Equal(3, anchors.Count);

            var anchorRows = anchors
                .Select(a => (int)a.Element(DrawingNs + "from")!.Element(DrawingNs + "row")!)
                .OrderBy(r => r)
                .ToList();
            Assert.Equal([0, 1, 2], anchorRows);
        }

        AssertPackageIsValidAndHasImages(path.ToString(), expectedImages: 3);
    }

    [Fact]
    public void NullImage_DoesNotProduceImageOrPlaceholder()
    {
        using var template = AutoDeletingPath.Create();
        using (var wb = new XLWorkbook())
        {
            var ws = wb.AddWorksheet("Sheet1");
            ws.Cell("A1").Value = "{{Logo}}";
            wb.SaveAs(template.FilePath);
        }

        using var path = AutoDeletingPath.Create();
        _templater.FillTemplate(path.ToString(), template.FilePath, new { Logo = (byte[]?)null });

        var sheet = GetSheetXml(path.ToString());
        Assert.Empty(GetMediaEntries(path.ToString()));
        Assert.DoesNotContain("System.Byte[]", sheet);
        Assert.DoesNotContain("{{Logo}}", sheet);
    }

    [Fact]
    public void DisabledByteArrayConversion_DoesNotProduceImage()
    {
        using var template = AutoDeletingPath.Create();
        using (var wb = new XLWorkbook())
        {
            var ws = wb.AddWorksheet("Sheet1");
            ws.Cell("A1").Value = "{{Logo}}";
            wb.SaveAs(template.FilePath);
        }

        using var path = AutoDeletingPath.Create();
        _templater.FillTemplate(path.ToString(), template.FilePath, new { Logo = TestPng() },
            configuration: new OpenXmlConfiguration { EnableConvertByteArray = false });

        Assert.Empty(GetMediaEntries(path.ToString()));
    }

    [Fact]
    public void MultipleImageProperties_AreRenderedIndependently()
    {
        using var template = AutoDeletingPath.Create();
        using (var wb = new XLWorkbook())
        {
            var ws = wb.AddWorksheet("Sheet1");
            ws.Cell("A1").Value = "{{Photo}}";
            ws.Cell("B1").Value = "{{Signature}}";
            wb.SaveAs(template.FilePath);
        }

        using var path = AutoDeletingPath.Create();
        _templater.FillTemplate(path.ToString(), template.FilePath, new { Photo = TestPng(), Signature = TestPng() });

        Assert.Equal(2, GetMediaEntries(path.ToString()).Count);

        using var zip = ZipFile.OpenRead(path.ToString());
        using var drawingStream = zip.GetEntry("xl/drawings/drawing1.xml")!.Open();
        var anchors = XDocument.Load(drawingStream).Descendants(DrawingNs + "oneCellAnchor").ToList();
        Assert.Equal(2, anchors.Count);
        Assert.Equal([0, 1], anchors.Select(a => (int)a.Element(DrawingNs + "from")!.Element(DrawingNs + "col")!).OrderBy(c => c).ToList());
    }

    [Fact]
    public void MultipleSheetsWithImages_EmitOneDrawingPerSheet()
    {
        using var template = AutoDeletingPath.Create();
        using (var wb = new XLWorkbook())
        {
            var first = wb.AddWorksheet("First");
            first.Cell("A1").Value = "{{Logo}}";
            var second = wb.AddWorksheet("Second");
            second.Cell("A1").Value = "{{Logo}}";
            wb.SaveAs(template.FilePath);
        }

        using var path = AutoDeletingPath.Create();
        _templater.FillTemplate(path.ToString(), template.FilePath, new { Logo = TestPng() });

        Assert.Equal(2, GetMediaEntries(path.ToString()).Count);

        using var zip = ZipFile.OpenRead(path.ToString());
        Assert.NotNull(zip.GetEntry("xl/drawings/drawing1.xml"));
        Assert.NotNull(zip.GetEntry("xl/drawings/drawing2.xml"));
        Assert.NotNull(zip.GetEntry("xl/worksheets/_rels/sheet1.xml.rels"));
        Assert.NotNull(zip.GetEntry("xl/worksheets/_rels/sheet2.xml.rels"));
    }

    [Fact]
    public void TemplateWithExistingImage_PreservesItAndAddsNewImage()
    {
        using var template = AutoDeletingPath.Create();
        using (var wb = new XLWorkbook())
        {
            var ws = wb.AddWorksheet("Sheet1");
            ws.Cell("A1").Value = "{{Logo}}";
            using var stream = new MemoryStream(TestPng());
            ws.AddPicture(stream).MoveTo(ws.Cell("D10"));
            wb.SaveAs(template.FilePath);
        }

        using var path = AutoDeletingPath.Create();
        _templater.FillTemplate(path.ToString(), template.FilePath, new { Logo = TestPng() });

        Assert.Equal(2, GetMediaEntries(path.ToString()).Count);

        using var package = new ExcelPackage(new FileInfo(path.ToString()));
        Assert.Equal(2, package.Workbook.Worksheets[0].Drawings.OfType<ExcelPicture>().Count());
    }

    [Fact]
    public void TemplateWithExistingPictures_AssignsUniquePictureIds()
    {
        using var template = AutoDeletingPath.Create();
        using (var wb = new XLWorkbook())
        {
            var ws = wb.AddWorksheet("Sheet1");
            ws.Cell("A1").Value = "{{Logo}}";
            using (var first = new MemoryStream(TestPng()))
                ws.AddPicture(first).MoveTo(ws.Cell("D10"));
            using (var second = new MemoryStream(TestPng()))
                ws.AddPicture(second).MoveTo(ws.Cell("D20"));
            wb.SaveAs(template.FilePath);
        }

        using var path = AutoDeletingPath.Create();
        _templater.FillTemplate(path.ToString(), template.FilePath, new { Logo = TestPng() });

        using var zip = ZipFile.OpenRead(path.ToString());
        using var drawingStream = zip.GetEntry("xl/drawings/drawing1.xml")!.Open();
        var pictureIds = XDocument.Load(drawingStream).Descendants()
            .Where(element => element.Name.LocalName == "cNvPr")
            .Select(element => (int)element.Attribute("id")!)
            .ToList();

        Assert.Equal(3, pictureIds.Count);
        Assert.Equal(pictureIds.Count, pictureIds.Distinct().Count());
    }

    [Fact]
    public void ParametrizedSheets_RenderImagesPerGeneratedSheet()
    {
        using var template = AutoDeletingPath.Create();
        using (var wb = new XLWorkbook())
        {
            var ws = wb.AddWorksheet("$Items$");
            ws.Cell("A1").Value = "{{Name}}";
            ws.Cell("B1").Value = "{{Image}}";
            wb.SaveAs(template.FilePath);
        }

        using var path = AutoDeletingPath.Create();
        var image = TestPng();
        _templater.FillTemplate(path.ToString(), template.FilePath, new
        {
            Items = new[]
            {
                new { Name = "A", Image = image },
                new { Name = "B", Image = image },
            }
        });

        Assert.Equal(2, GetMediaEntries(path.ToString()).Count);

        using var zip = ZipFile.OpenRead(path.ToString());
        Assert.NotNull(zip.GetEntry("xl/drawings/drawing1.xml"));
        Assert.NotNull(zip.GetEntry("xl/drawings/drawing2.xml"));
    }

    [Fact]
    public void LegacyFacade_SaveAsByTemplate_RendersImages()
    {
        using var template = AutoDeletingPath.Create();
        using (var wb = new XLWorkbook())
        {
            var ws = wb.AddWorksheet("Sheet1");
            ws.Cell("A1").Value = "{{Logo}}";
            wb.SaveAs(template.FilePath);
        }

        using var path = AutoDeletingPath.Create();
        MiniExcelLibs.MiniExcel.SaveAsByTemplate(path.ToString(), template.FilePath, new { Logo = TestPng() });

        Assert.Single(GetMediaEntries(path.ToString()));
        AssertPackageIsValidAndHasImages(path.ToString(), expectedImages: 1);
    }

    [Fact]
    public void TemplateWithExistingSheetRels_MergesDrawingRelationship()
    {
        using var template = AutoDeletingPath.Create();
        using (var wb = new XLWorkbook())
        {
            var ws = wb.AddWorksheet("Sheet1");
            ws.Cell("A1").Value = "{{Logo}}";
            ws.Cell("B1").SetHyperlink(new XLHyperlink("https://example.com"));
            wb.SaveAs(template.FilePath);
        }

        using var path = AutoDeletingPath.Create();
        _templater.FillTemplate(path.ToString(), template.FilePath, new { Logo = TestPng() });

        Assert.Single(GetMediaEntries(path.ToString()));

        using var zip = ZipFile.OpenRead(path.ToString());
        using var relsStream = zip.GetEntry("xl/worksheets/_rels/sheet1.xml.rels")!.Open();
        var rels = XDocument.Load(relsStream).ToString();
        Assert.Contains("relationships/hyperlink", rels); // pre-existing relationship preserved
        Assert.Contains("relationships/drawing", rels);   // our drawing relationship merged in
    }
}
