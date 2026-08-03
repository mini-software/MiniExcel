using System.Xml.Linq;
using MiniExcelLib.Core.Enums;
using MiniExcelLib.OpenXml.Picture;
using MiniExcelLib.Tests.Common.Utils;

namespace MiniExcelLib.OpenXml.Tests;

public class MiniExcelPlaceInCellTests
{
    private readonly OpenXmlExporter _excelExporter = MiniExcel.Exporters.GetOpenXmlExporter();
    private readonly OpenXmlTemplater _excelTemplater = MiniExcel.Templaters.GetOpenXmlTemplater();

    [Fact]
    public void PlaceInCell_OnMemoryStream_DoesNotWriteToDisk()
    {
        using var stream = new MemoryStream();
        _excelExporter.Export(stream, new[] { new { Name = "A" } });
        stream.Position = 0;

        var imageBytes = File.ReadAllBytes(PathHelper.GetFile("images/github_logo.png"));
        _excelTemplater.AddPicture(stream,
            new MiniExcelPicture
            {
                ImageBytes = imageBytes,
                CellAddress = "C3",
                ImgType = XlsxImgType.PlaceInCell,
                PictureType = "image/png"
            });

        stream.Position = 0;
        using var zip = new ZipArchive(stream, ZipArchiveMode.Read, leaveOpen: true);

        Assert.Contains(zip.Entries, e => e.FullName.StartsWith("xl/media/image", StringComparison.OrdinalIgnoreCase));
        Assert.NotNull(zip.GetEntry("xl/metadata.xml"));
        Assert.NotNull(zip.GetEntry("xl/richData/rdrichvalue.xml"));

        XNamespace ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
        var sheet = XDocument.Load(zip.GetEntry("xl/worksheets/sheet1.xml")!.Open());
        var cell = sheet.Descendants(ns + "c").First(c => (string?)c.Attribute("r") == "C3");
        Assert.Equal("e", (string?)cell.Attribute("t"));
        Assert.Equal("1", (string?)cell.Attribute("vm"));
        Assert.Equal("#VALUE!", (string?)cell.Element(ns + "v"));
    }

    [Fact]
    public void PlaceInCell_WritesRichDataPartsAndCellValue()
    {
        using var path = AutoDeletingPath.Create();
        _excelExporter.Export(path.FilePath, new[] { new { Name = "A" } });

        var imageBytes = File.ReadAllBytes(PathHelper.GetFile("images/github_logo.png"));
        _excelTemplater.AddPicture(path.FilePath,
            new MiniExcelPicture
            {
                ImageBytes = imageBytes,
                CellAddress = "C3",
                ImgType = XlsxImgType.PlaceInCell,
                PictureType = "image/png"
            });

        using var zip = ZipFile.OpenRead(path.FilePath);
        Assert.Contains(zip.Entries, e => e.FullName.StartsWith("xl/media/image", StringComparison.OrdinalIgnoreCase));
        Assert.NotNull(zip.GetEntry("xl/metadata.xml"));
        Assert.NotNull(zip.GetEntry("xl/richData/rdrichvalue.xml"));
        Assert.NotNull(zip.GetEntry("xl/richData/rdrichvaluestructure.xml"));
        Assert.NotNull(zip.GetEntry("xl/richData/rdRichValueTypes.xml"));
        Assert.NotNull(zip.GetEntry("xl/richData/richValueRel.xml"));
        Assert.NotNull(zip.GetEntry("xl/richData/_rels/richValueRel.xml.rels"));

        var contentTypes = XDocument.Load(zip.GetEntry("[Content_Types].xml")!.Open());
        Assert.Contains(contentTypes.Root!.Elements(),
            e => (string?)e.Attribute("PartName") == "/xl/metadata.xml");
        Assert.Contains(contentTypes.Root!.Elements(),
            e => (string?)e.Attribute("PartName") == "/xl/richData/rdrichvalue.xml");

        var workbookRels = XDocument.Load(zip.GetEntry("xl/_rels/workbook.xml.rels")!.Open());
        var relTypes = workbookRels.Root!.Elements()
            .Select(e => (string?)e.Attribute("Type"))
            .ToHashSet(StringComparer.Ordinal);
        Assert.Contains("http://schemas.openxmlformats.org/officeDocument/2006/relationships/sheetMetadata", relTypes);
        Assert.Contains("http://schemas.microsoft.com/office/2017/06/relationships/rdRichValue", relTypes);
        Assert.Contains("http://schemas.microsoft.com/office/2022/10/relationships/richValueRel", relTypes);

        var sheet = XDocument.Load(zip.GetEntry("xl/worksheets/sheet1.xml")!.Open());
        XNamespace ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
        var cell = sheet.Descendants(ns + "c")
            .FirstOrDefault(c => (string?)c.Attribute("r") == "C3");
        Assert.NotNull(cell);
        Assert.Equal("e", (string?)cell.Attribute("t"));
        Assert.Equal("1", (string?)cell.Attribute("vm"));
        Assert.Equal("#VALUE!", (string?)cell.Element(ns + "v"));

        var dimension = sheet.Root!.Element(ns + "dimension");
        Assert.NotNull(dimension);
        var dimRef = (string?)dimension.Attribute("ref");
        Assert.False(string.IsNullOrEmpty(dimRef));
        Assert.Contains("C3", dimRef, StringComparison.OrdinalIgnoreCase);

        var structure = XDocument.Load(zip.GetEntry("xl/richData/rdrichvaluestructure.xml")!.Open());
        XNamespace rd = "http://schemas.microsoft.com/office/spreadsheetml/2017/richdata";
        Assert.Contains(structure.Root!.Elements(rd + "s"),
            s => (string?)s.Attribute("t") == "_localImage");
    }

    [Fact]
    public void PlaceInCell_MultipleImages_IncrementVmAndRelIndexes()
    {
        using var path = AutoDeletingPath.Create();
        _excelExporter.Export(path.FilePath, new[] { new { Name = "A" } });

        var img1 = File.ReadAllBytes(PathHelper.GetFile("images/github_logo.png"));
        var img2 = File.ReadAllBytes(PathHelper.GetFile("images/google_logo.png"));

        _excelTemplater.AddPicture(path.FilePath,
            new MiniExcelPicture
            {
                ImageBytes = img1,
                CellAddress = "A1",
                ImgType = XlsxImgType.PlaceInCell
            },
            new MiniExcelPicture
            {
                ImageBytes = img2,
                CellAddress = "B2",
                ImgType = XlsxImgType.PlaceInCell
            });

        using var zip = ZipFile.OpenRead(path.FilePath);
        XNamespace ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
        XNamespace rd = "http://schemas.microsoft.com/office/spreadsheetml/2017/richdata";
        XNamespace rvr = "http://schemas.microsoft.com/office/spreadsheetml/2022/richvaluerel";
        XNamespace pkg = "http://schemas.openxmlformats.org/package/2006/relationships";

        var sheet = XDocument.Load(zip.GetEntry("xl/worksheets/sheet1.xml")!.Open());
        var a1 = sheet.Descendants(ns + "c").First(c => (string?)c.Attribute("r") == "A1");
        var b2 = sheet.Descendants(ns + "c").First(c => (string?)c.Attribute("r") == "B2");
        Assert.Equal("1", (string?)a1.Attribute("vm"));
        Assert.Equal("2", (string?)b2.Attribute("vm"));

        var rvData = XDocument.Load(zip.GetEntry("xl/richData/rdrichvalue.xml")!.Open());
        var rvs = rvData.Root!.Elements(rd + "rv").ToList();
        Assert.Equal(2, rvs.Count);
        Assert.Equal("0", rvs[0].Elements(rd + "v").First().Value);
        Assert.Equal("5", rvs[0].Elements(rd + "v").Skip(1).First().Value);
        Assert.Equal("1", rvs[1].Elements(rd + "v").First().Value);

        var richValueRel = XDocument.Load(zip.GetEntry("xl/richData/richValueRel.xml")!.Open());
        Assert.Equal(2, richValueRel.Root!.Elements(rvr + "rel").Count());

        var rels = XDocument.Load(zip.GetEntry("xl/richData/_rels/richValueRel.xml.rels")!.Open());
        Assert.Equal(2, rels.Root!.Elements(pkg + "Relationship").Count());

        var metadata = XDocument.Load(zip.GetEntry("xl/metadata.xml")!.Open());
        var valueMetadata = metadata.Root!.Element(ns + "valueMetadata");
        Assert.Equal("2", (string?)valueMetadata!.Attribute("count"));
    }

    [Fact]
    public void PlaceInCell_CanMixWithDrawingAnchor()
    {
        using var path = AutoDeletingPath.Create();
        _excelExporter.Export(path.FilePath, new[] { new { Name = "A" } });

        var img1 = File.ReadAllBytes(PathHelper.GetFile("images/github_logo.png"));
        var img2 = File.ReadAllBytes(PathHelper.GetFile("images/google_logo.png"));

        _excelTemplater.AddPicture(path.FilePath,
            new MiniExcelPicture
            {
                ImageBytes = img1,
                CellAddress = "C3",
                ImgType = XlsxImgType.PlaceInCell
            },
            new MiniExcelPicture
            {
                ImageBytes = img2,
                CellAddress = "E5",
                ImgType = XlsxImgType.OneCellAnchor,
                WidthPx = 50,
                HeightPx = 50
            });

        using var zip = ZipFile.OpenRead(path.FilePath);
        Assert.NotNull(zip.GetEntry("xl/metadata.xml"));
        Assert.NotNull(zip.GetEntry("xl/richData/rdrichvalue.xml"));
        Assert.Contains(zip.Entries, e => e.FullName.StartsWith("xl/drawings/drawing", StringComparison.OrdinalIgnoreCase));

        XNamespace ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
        var sheet = XDocument.Load(zip.GetEntry("xl/worksheets/sheet1.xml")!.Open());
        var c3 = sheet.Descendants(ns + "c").First(c => (string?)c.Attribute("r") == "C3");
        Assert.Equal("e", (string?)c3.Attribute("t"));
        Assert.NotNull(sheet.Root!.Element(ns + "drawing"));
    }
}
