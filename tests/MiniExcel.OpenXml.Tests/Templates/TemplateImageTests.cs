using System.Reflection;
using System.Xml.Linq;
using ClosedXML.Excel;
using MiniExcelLib.OpenXml.Templates;
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

    private static byte[] OtherPng() => File.ReadAllBytes(PathHelper.GetFile("images/github_logo.png"));

    /// <summary>
    /// Builds a large, unique image payload whose magic bytes classify it as a PNG. The size makes a
    /// leak meaningful and the uniqueness guarantees the array is not shared with any fixture.
    /// </summary>
    private static byte[] CreateLargeRecognisableImage(int size = 1_000_000)
    {
        var image = new byte[size];
        // PNG signature: enough for ImageHelper.GetImageFormat to recognise the bytes as an image.
        image[0] = 137; image[1] = 80; image[2] = 78; image[3] = 71;
        image[4] = 13; image[5] = 10; image[6] = 26; image[7] = 10;
        return image;
    }

    private static void CollectGarbage()
    {
        GC.Collect();
        GC.WaitForPendingFinalizers();
        GC.Collect();
    }

    /// <summary>
    /// Fills a template with a large image and returns only a <see cref="WeakReference"/> to its
    /// bytes. Every strong reference to the bytes lives and dies inside this method, so the caller
    /// can verify through garbage collection that the template pipeline did not retain them.
    /// </summary>
    private static WeakReference FillTemplateAndTrackImageBytes(string templatePath)
    {
        var image = CreateLargeRecognisableImage();
        var weakReference = new WeakReference(image);

        using var output = new MemoryStream();
        var openXmlTemplate = new OpenXmlTemplate(output, null, new OpenXmlValueExtractor());
        openXmlTemplate.SaveAsByTemplate(templatePath, new { Logo = image });

        return weakReference;
    }

    private static IReadOnlyList<string?> GetEmbedIds(string xlsxPath, string drawingPath = "xl/drawings/drawing1.xml")
    {
        var relationshipNs = XNamespace.Get("http://schemas.openxmlformats.org/officeDocument/2006/relationships");
        using var zip = ZipFile.OpenRead(xlsxPath);
        using var drawingStream = zip.GetEntry(drawingPath)!.Open();
        return XDocument.Load(drawingStream).Descendants()
            .Where(element => element.Name.LocalName == "blip")
            .Select(element => (string?)element.Attribute(relationshipNs + "embed"))
            .ToList();
    }

    /// <summary>
    /// Verifies the complete drawing relationship chain for a drawing part: every <c>r:embed</c>
    /// resolves to a declared relationship, every image relationship targets an existing media part,
    /// the picture ids are unique and the generated media parts are not overwritten.
    /// </summary>
    private static void AssertDrawingReferencesIntegrity(string xlsxPath, string drawingPath = "xl/drawings/drawing1.xml")
    {
        using var zip = ZipFile.OpenRead(xlsxPath);
        var relsPath = $"xl/drawings/_rels/{Path.GetFileName(drawingPath)}.rels";
        Assert.NotNull(zip.GetEntry(relsPath));

        using (var drawingStream = zip.GetEntry(drawingPath)!.Open())
        {
            var pictureIds = XDocument.Load(drawingStream).Descendants()
                .Where(element => element.Name.LocalName == "cNvPr")
                .Select(element => (string?)element.Attribute("id"))
                .ToList();
            Assert.Equal(pictureIds.Count, pictureIds.Distinct().Count());
        }

        var embeds = GetEmbedIds(xlsxPath, drawingPath);

        using var relsStream = zip.GetEntry(relsPath)!.Open();
        var imageRelationships = XDocument.Load(relsStream).Descendants()
            .Where(element => element.Name.LocalName == "Relationship")
            .Where(element => element.Attribute("Type")?.Value.EndsWith("/image", StringComparison.Ordinal) is true)
            .ToList();

        var relationshipIds = imageRelationships.Select(rel => (string?)rel.Attribute("Id")).ToList();
        Assert.Equal(relationshipIds.Count, relationshipIds.Distinct().Count());

        // Every anchor must resolve to a declared relationship ...
        Assert.All(embeds, embed => Assert.Contains(embed, relationshipIds));

        // ... and every relationship must resolve to a media part that exists in the package.
        Assert.Equal(embeds.Count, imageRelationships.Count);
        foreach (var target in imageRelationships.Select(rel => rel.Attribute("Target")!.Value.TrimStart('/')))
        {
            Assert.NotNull(zip.GetEntry(target));
        }
    }

    /// <summary>
    /// Resolves the drawing part a worksheet relationship points at, normalizing both the relative
    /// (<c>../drawings/..</c>) and absolute (<c>/xl/drawings/..</c>) target forms.
    /// </summary>
    private static string GetDrawingPartForSheet(string xlsxPath, int sheetIndex)
    {
        using var zip = ZipFile.OpenRead(xlsxPath);
        using var relsStream = zip.GetEntry($"xl/worksheets/_rels/sheet{sheetIndex}.xml.rels")!.Open();
        var target = XDocument.Load(relsStream).Descendants()
            .Where(element => element.Name.LocalName == "Relationship")
            .First(element => element.Attribute("Type")?.Value.EndsWith("/drawing", StringComparison.Ordinal) is true)
            .Attribute("Target")!.Value
            .Replace('\\', '/');

        return target.StartsWith("../", StringComparison.Ordinal)
            ? "xl/" + target[3..]
            : target.TrimStart('/');
    }

    private static int GetPictureCount(string xlsxPath, string drawingPath)
    {
        using var zip = ZipFile.OpenRead(xlsxPath);
        using var drawingStream = zip.GetEntry(drawingPath)!.Open();
        return XDocument.Load(drawingStream).Descendants()
            .Count(element => element.Name.LocalName == "blip");
    }

    /// <summary>
    /// Drops the <c>xmlns:r</c> declaration from a template worksheet that does not otherwise use the
    /// relationships namespace, to exercise the generated <c>r:id</c> binding.
    /// </summary>
    private static void RemoveWorksheetRelationshipsNamespace(string xlsxPath)
    {
        using var zip = ZipFile.Open(xlsxPath, ZipArchiveMode.Update);
        var entry = zip.GetEntry("xl/worksheets/sheet1.xml")!;

        string content;
        using (var reader = new StreamReader(entry.Open()))
        {
            content = reader.ReadToEnd();
        }

        entry.Delete();

        var updated = zip.CreateEntry("xl/worksheets/sheet1.xml");
        using var writer = new StreamWriter(updated.Open());
        writer.Write(content.Replace(
            " xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\"",
            string.Empty,
            StringComparison.Ordinal));
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
    public void TwoImagesInSameCell_AreRenderedAsDistinctPictures()
    {
        using var template = AutoDeletingPath.Create();
        using (var wb = new XLWorkbook())
        {
            var ws = wb.AddWorksheet("Sheet1");
            ws.Cell("A1").Value = "{{Image1}} {{Image2}}";
            wb.SaveAs(template.FilePath);
        }

        using var path = AutoDeletingPath.Create();
        _templater.FillTemplate(path.ToString(), template.FilePath, new { Image1 = TestPng(), Image2 = OtherPng() });

        // Both images share the same anchor cell; each one must still get its own media part and
        // relationship instead of overwriting the other.
        Assert.Equal(2, GetMediaEntries(path.ToString()).Count);
        Assert.Equal(2, GetEmbedIds(path.ToString()).Distinct().Count());

        AssertDrawingReferencesIntegrity(path.ToString());
        AssertPackageIsValidAndHasImages(path.ToString(), expectedImages: 2);
    }

    [Fact]
    public void CollectionWithMultipleImageColumns_ProducesDistinctParts()
    {
        using var template = AutoDeletingPath.Create();
        using (var wb = new XLWorkbook())
        {
            var ws = wb.AddWorksheet("Sheet1");
            ws.Cell("A1").Value = "{{Products.Name}}";
            ws.Cell("B1").Value = "{{Products.Image1}}";
            ws.Cell("C1").Value = "{{Products.Image2}}";
            wb.SaveAs(template.FilePath);
        }

        using var path = AutoDeletingPath.Create();
        _templater.FillTemplate(path.ToString(), template.FilePath, new
        {
            Products = new[]
            {
                new { Name = "A", Image1 = TestPng(), Image2 = OtherPng() },
                new { Name = "B", Image1 = OtherPng(), Image2 = TestPng() },
            }
        });

        Assert.Equal(4, GetMediaEntries(path.ToString()).Count());

        AssertDrawingReferencesIntegrity(path.ToString());
        AssertPackageIsValidAndHasImages(path.ToString(), expectedImages: 4);
    }

    [Fact]
    public void LargeCollectionWithImages_EmitsEveryImage()
    {
        using var template = AutoDeletingPath.Create();
        using (var wb = new XLWorkbook())
        {
            var ws = wb.AddWorksheet("Sheet1");
            ws.Cell("A1").Value = "{{Products.Name}}";
            ws.Cell("B1").Value = "{{Products.Image}}";
            wb.SaveAs(template.FilePath);
        }

        var image = TestPng();
        var products = Enumerable.Range(0, 64)
            .Select(i => new { Name = $"P{i}", Image = image })
            .ToArray();

        using var path = AutoDeletingPath.Create();
        _templater.FillTemplate(path.ToString(), template.FilePath, new { Products = products });

        // Progressive collection expansion must not drop or overwrite any image.
        Assert.Equal(64, GetMediaEntries(path.ToString()).Count());

        var sheet = GetSheetXml(path.ToString());
        Assert.DoesNotContain("@@@imageid@@@", sheet);

        AssertDrawingReferencesIntegrity(path.ToString());
        AssertPackageIsValidAndHasImages(path.ToString(), expectedImages: 64);
    }

    [Fact]
    public void ImageState_IsReleasedWhenTheCallCompletes()
    {
        using var template = AutoDeletingPath.Create();
        using (var wb = new XLWorkbook())
        {
            var ws = wb.AddWorksheet("Sheet1");
            ws.Cell("A1").Value = "{{Products.Name}}";
            ws.Cell("B1").Value = "{{Products.Image}}";
            wb.SaveAs(template.FilePath);
        }

        var image = TestPng();
        var products = Enumerable.Range(0, 8).Select(i => new { Name = $"P{i}", Image = image }).ToArray();

        using var output = new MemoryStream();
        var openXmlTemplate = new OpenXmlTemplate(output, null, new OpenXmlValueExtractor());
        openXmlTemplate.SaveAsByTemplate(template.FilePath, new { Products = products });

        // Exactly measuring peak memory is not reliable in a test; instead this checks the ownership
        // invariant of the image pipeline: once the call returns, the pending, reuse and emission
        // collections are empty, so the template instance no longer pins any image bytes.
        var type = typeof(OpenXmlTemplate);
        foreach (var fieldName in new[] { "_pendingImages", "_capturedImages", "_files" })
        {
            var field = type.GetField(fieldName, BindingFlags.NonPublic | BindingFlags.Instance);
            Assert.NotNull(field);

            var collection = field!.GetValue(openXmlTemplate);
            Assert.NotNull(collection);
            Assert.Equal(0, (int)collection!.GetType().GetProperty("Count")!.GetValue(collection)!);
        }
    }

    [Fact]
    public void ImageBytes_AreCollectableAfterTheCallCompletes()
    {
        using var template = AutoDeletingPath.Create();
        using (var wb = new XLWorkbook())
        {
            var ws = wb.AddWorksheet("Sheet1");
            ws.Cell("A1").Value = "{{Logo}}";
            wb.SaveAs(template.FilePath);
        }

        // The helper owns every strong reference to the image bytes and hands back only a weak one.
        // A leak anywhere in the template pipeline (media capture, pending/captured dictionaries,
        // emission state, or the output archive) would keep the bytes alive and fail this assertion.
        var weakReference = FillTemplateAndTrackImageBytes(template.FilePath);

        CollectGarbage();

        Assert.False(weakReference.IsAlive,
            "The template pipeline retained the image bytes after the call completed.");
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

        AssertDrawingReferencesIntegrity(path.ToString());
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

        AssertDrawingReferencesIntegrity(path.ToString());
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

    [Fact]
    public void ExistingPictureOnAnotherSheet_DoesNotReuseItsDrawingPart()
    {
        using var template = AutoDeletingPath.Create();
        using (var wb = new XLWorkbook())
        {
            // Sheet 1 needs a generated drawing while sheet 2 already owns a template drawing
            // (xl/drawings/drawing1.xml). The generated part must not reuse that filename.
            var first = wb.AddWorksheet("First");
            first.Cell("A1").Value = "{{Logo}}";

            var second = wb.AddWorksheet("Second");
            using (var stream = new MemoryStream(TestPng()))
                second.AddPicture(stream).MoveTo(second.Cell("A1"));

            wb.SaveAs(template.FilePath);
        }

        using var path = AutoDeletingPath.Create();
        _templater.FillTemplate(path.ToString(), template.FilePath, new { Logo = TestPng() });

        var firstDrawing = GetDrawingPartForSheet(path.ToString(), 1);
        var secondDrawing = GetDrawingPartForSheet(path.ToString(), 2);

        // Each sheet must reference its own drawing part, and both parts must exist.
        Assert.NotEqual(firstDrawing, secondDrawing);
        using (var zip = ZipFile.OpenRead(path.ToString()))
        {
            Assert.NotNull(zip.GetEntry(firstDrawing));
            Assert.NotNull(zip.GetEntry(secondDrawing));
        }

        // The template picture stays on sheet 2 and the generated image lands on sheet 1.
        Assert.Equal(1, GetPictureCount(path.ToString(), firstDrawing));
        Assert.Equal(1, GetPictureCount(path.ToString(), secondDrawing));

        AssertDrawingReferencesIntegrity(path.ToString(), firstDrawing);
        AssertDrawingReferencesIntegrity(path.ToString(), secondDrawing);

        using var package = new ExcelPackage(new FileInfo(path.ToString()));
        Assert.Single(package.Workbook.Worksheets[0].Drawings.OfType<ExcelPicture>());
        Assert.Single(package.Workbook.Worksheets[1].Drawings.OfType<ExcelPicture>());
    }

    [Fact]
    public void SheetWithComment_PlacesDrawingBeforeLegacyDrawing()
    {
        using var template = AutoDeletingPath.Create();
        using (var wb = new XLWorkbook())
        {
            var ws = wb.AddWorksheet("Sheet1");
            ws.Cell("A1").Value = "{{Logo}}";
            ws.Cell("A1").GetComment().AddText("a comment");
            wb.SaveAs(template.FilePath);
        }

        using var path = AutoDeletingPath.Create();
        _templater.FillTemplate(path.ToString(), template.FilePath, new { Logo = TestPng() });

        using var zip = ZipFile.OpenRead(path.ToString());
        using var sheetStream = zip.GetEntry("xl/worksheets/sheet1.xml")!.Open();
        var worksheet = XDocument.Load(sheetStream).Root!;
        var afterSheetData = worksheet.Element(SpreadsheetNs + "sheetData")!
            .ElementsAfterSelf()
            .Select(element => element.Name.LocalName)
            .ToList();

        // The worksheet schema requires <drawing> before <legacyDrawing> (used by comments).
        Assert.Contains("drawing", afterSheetData);
        Assert.Contains("legacyDrawing", afterSheetData);
        Assert.True(
            afterSheetData.IndexOf("drawing") < afterSheetData.IndexOf("legacyDrawing"),
            $"Expected <drawing> before <legacyDrawing> but got: {string.Join(", ", afterSheetData)}");

        // Both the image and the comment relationship survive.
        Assert.Single(GetMediaEntries(path.ToString()));
        using var relsStream = zip.GetEntry("xl/worksheets/_rels/sheet1.xml.rels")!.Open();
        var rels = XDocument.Load(relsStream).ToString();
        Assert.Contains("relationships/drawing", rels);
        Assert.Contains("relationships/vmlDrawing", rels);
    }

    [Fact]
    public void WorksheetWithoutRelationshipsNamespace_DeclaresItForGeneratedDrawing()
    {
        using var template = AutoDeletingPath.Create();
        using (var wb = new XLWorkbook())
        {
            var ws = wb.AddWorksheet("Sheet1");
            ws.Cell("A1").Value = "{{Logo}}";
            wb.SaveAs(template.FilePath);
        }

        // A template may not declare xmlns:r at all; the generated r:id must still be bound.
        RemoveWorksheetRelationshipsNamespace(template.FilePath);

        using var path = AutoDeletingPath.Create();
        _templater.FillTemplate(path.ToString(), template.FilePath, new { Logo = TestPng() });

        var relationshipNs = XNamespace.Get("http://schemas.openxmlformats.org/officeDocument/2006/relationships");

        using var zip = ZipFile.OpenRead(path.ToString());
        using var sheetStream = zip.GetEntry("xl/worksheets/sheet1.xml")!.Open();

        // XDocument.Load throws on an unbound prefix, so this also proves the declaration is present.
        var drawing = XDocument.Load(sheetStream).Descendants(SpreadsheetNs + "drawing").Single();
        Assert.Equal("rDrawing1", (string?)drawing.Attribute(relationshipNs + "id"));

        AssertPackageIsValidAndHasImages(path.ToString(), expectedImages: 1);
    }

    [Fact]
    public void MalformedTiffBytes_DoNotAbortTemplateRendering()
    {
        using var template = AutoDeletingPath.Create();
        var templatePath = template.FilePath; 
        
        using (var wb = new XLWorkbook())
        {
            var ws = wb.AddWorksheet("Sheet1");
            ws.Cell("A1").Value = "{{Logo}}";
            wb.SaveAs(template.FilePath);
        }

        // TIFF header whose IFD offset is int.MaxValue: the size parser must return null instead of
        // overflowing, and the export must still complete.
        byte[] malformedTiff = [(byte)'I', (byte)'I', 0x2A, 0x00, 0xFF, 0xFF, 0xFF, 0x7F];

        using var path = AutoDeletingPath.Create();
        var finalPath = path.FilePath;
        
        var exception = Record.Exception(() =>
            _templater.FillTemplate(finalPath, templatePath, new { Logo = malformedTiff }));

        Assert.Null(exception);

        // The bytes are still recognised as an image, so a picture is emitted with the default size.
        Assert.Single(GetMediaEntries(path.ToString()));
        Assert.DoesNotContain("System.Byte[]", GetSheetXml(path.ToString()));
    }
}
