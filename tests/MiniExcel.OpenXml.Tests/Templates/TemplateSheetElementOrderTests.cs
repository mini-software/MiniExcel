using System.Xml.Linq;
using ClosedXML.Excel;
using MiniExcelLib.Tests.Common.Utils;

namespace MiniExcelLib.OpenXml.Tests.Templates;

/// <summary>
/// Regression test for #1077: a template sheet with sheetProtection and mergeCells must keep the
/// ECMA-376 child order (sheetData, sheetProtection, mergeCells, ...) after the template is filled.
/// </summary>
public class TemplateSheetElementOrderTests
{
    private readonly OpenXmlTemplater _templater = MiniExcelV2.Templaters.GetOpenXmlTemplater();

    [Fact]
    public void TemplateWithSheetProtectionAndMergeCells_WritesSheetProtectionBeforeMergeCells()
    {
        using var template = AutoDeletingPath.Create();
        using (var wb = new XLWorkbook())
        {
            var ws = wb.AddWorksheet("Sheet1");
            ws.Cell("A1").Value = "{{name}}";
            ws.Cell("B1").Value = "merged";
            ws.Range("B1:C1").Merge();
            ws.Protect();
            wb.SaveAs(template.FilePath);
        }

        using var path = AutoDeletingPath.Create();
        _templater.FillTemplate(path.ToString(), template.FilePath, new Dictionary<string, object?> { ["name"] = "Jack" });

        using var zip = ZipFile.OpenRead(path.ToString());
        using var stream = zip.GetEntry("xl/worksheets/sheet1.xml")!.Open();
        var names = XDocument.Load(stream).Root!.Elements().Select(e => e.Name.LocalName).ToList();

        var protection = names.IndexOf("sheetProtection");
        var merge = names.IndexOf("mergeCells");
        Assert.True(protection >= 0, "sheetProtection is missing: " + string.Join(",", names));
        Assert.True(merge >= 0, "mergeCells is missing: " + string.Join(",", names));
        Assert.True(protection < merge, "sheetProtection must come before mergeCells: " + string.Join(",", names));
    }

    [Fact]
    public void TemplateWithAllElementsBetweenSheetDataAndMergeCells_KeepsSchemaOrderWithoutDuplicates()
    {
        const string sheetXml =
            "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">" +
            "<dimension ref=\"A1:C1\"/>" +
            "<sheetData><row r=\"1\"><c r=\"A1\" t=\"inlineStr\"><is><t>{{name}}</t></is></c>" +
            "<c r=\"B1\" t=\"inlineStr\"><is><t>merged</t></is></c><c r=\"C1\"/></row></sheetData>" +
            "<sheetCalcPr fullCalcOnLoad=\"1\"/>" +
            "<sheetProtection sheet=\"1\"/>" +
            "<protectedRanges><protectedRange sqref=\"A1\" name=\"r1\"/></protectedRanges>" +
            "<scenarios/>" +
            "<autoFilter ref=\"A1:C1\"/>" +
            "<sortState ref=\"A1:C1\"/>" +
            "<dataConsolidate/>" +
            "<customSheetViews><customSheetView guid=\"{00000000-0000-0000-0000-000000000001}\"/></customSheetViews>" +
            "<mergeCells count=\"1\"><mergeCell ref=\"B1:C1\"/></mergeCells>" +
            "<phoneticPr fontId=\"0\"/>" +
            "<pageMargins left=\"0.7\" right=\"0.7\" top=\"0.75\" bottom=\"0.75\" header=\"0.3\" footer=\"0.3\"/>" +
            "</worksheet>";

        using var template = AutoDeletingPath.Create();
        using (var archive = ZipFile.Open(template.FilePath, ZipArchiveMode.Create))
        {
            void AddEntry(string name, string content)
            {
                using var writer = new StreamWriter(archive.CreateEntry(name).Open());
                writer.Write(content);
            }

            AddEntry("[Content_Types].xml",
                "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">" +
                "<Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>" +
                "<Default Extension=\"xml\" ContentType=\"application/xml\"/>" +
                "<Override PartName=\"/xl/workbook.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml\"/>" +
                "<Override PartName=\"/xl/worksheets/sheet1.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml\"/>" +
                "</Types>");
            AddEntry("_rels/.rels",
                "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">" +
                "<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" Target=\"xl/workbook.xml\"/>" +
                "</Relationships>");
            AddEntry("xl/workbook.xml",
                "<workbook xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">" +
                "<sheets><sheet name=\"Sheet1\" sheetId=\"1\" r:id=\"rId1\"/></sheets></workbook>");
            AddEntry("xl/_rels/workbook.xml.rels",
                "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">" +
                "<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\" Target=\"worksheets/sheet1.xml\"/>" +
                "</Relationships>");
            AddEntry("xl/worksheets/sheet1.xml", sheetXml);
        }

        using var path = AutoDeletingPath.Create();
        _templater.FillTemplate(path.ToString(), template.FilePath, new Dictionary<string, object?> { ["name"] = "Jack" });

        using var zip = ZipFile.OpenRead(path.ToString());
        using var stream = zip.GetEntry("xl/worksheets/sheet1.xml")!.Open();
        var names = XDocument.Load(stream).Root!.Elements().Select(e => e.Name.LocalName).ToList();

        string[] expected =
        [
            "dimension", "sheetData", "sheetCalcPr", "sheetProtection", "protectedRanges", "scenarios",
            "autoFilter", "sortState", "dataConsolidate", "customSheetViews", "mergeCells", "phoneticPr", "pageMargins"
        ];
        Assert.Equal(expected, names);
    }
}

