using System.IO.Compression;
using System.Xml;
using ClosedXML.Excel;
using MiniExcelLibs.Tests.Utils;
using Xunit;

namespace MiniExcelLibs.Tests.SaveByTemplate;

/// <summary>
/// Regression tests for calcChain.xml handling and '$='-formula cell serialization in template
/// rendering. A template containing any formula carries a calcChain part whose entries point at
/// cell addresses; row insertion shifts formula cells, so the chain must be regenerated from the
/// rendered output — never left stale, and never written empty (a calcChain with zero &lt;c&gt;
/// entries is schema-invalid and Excel refuses to open the whole file).
/// </summary>
public class MiniExcelTemplateCalcChainTests
{
    [Fact]
    public void TemplateWithStaticFormula_DoesNotWriteStaleOrEmptyCalcChain()
    {
        // A template with a static Excel formula (below an IEnumerable row) carries a calcChain
        // pointing at the formula's pre-render address. After rows are inserted the address is
        // stale — the rendered package must not contain a stale or empty calcChain.
        using var template = AutoDeletingPath.Create();
        using (var wb = new XLWorkbook())
        {
            var ws = wb.AddWorksheet("Sheet1");
            ws.Cell("A1").Value = "{{title}}";
            ws.Cell("A3").Value = "{{items.Name}}";
            ws.Cell("B3").Value = "{{items.Qty}}";
            ws.Cell("B5").FormulaA1 = "SUM(B3:B4)";
            wb.SaveAs(template.FilePath);
        }

        using var path = AutoDeletingPath.Create();
        Dictionary<string, object?> data = new()
        {
            ["title"] = "FooCompany",
            ["items"] = new[]
            {
                new { Name = "A", Qty = 1 },
                new { Name = "B", Qty = 2 },
            }
        };
        MiniExcel.SaveAsByTemplate(path.ToString(), template.FilePath, data);

        using var zip = ZipFile.OpenRead(path.ToString());
        var calcChain = zip.GetEntry("xl/calcChain.xml");
        if (calcChain != null)
        {
            using var reader = new StreamReader(calcChain.Open());
            var content = reader.ReadToEnd();
            Assert.Contains("<c ", content); // an empty calcChain is schema-invalid
            Assert.DoesNotContain(@"r=""B5""", content); // the pre-render address is stale after the row shift
        }
    }

    [Fact]
    public void DollarFormulaInSparseRow_WritesValidFormulaCellAndCorrectCalcChainRef()
    {
        // The '$=' formula sits in column D of a row whose only other cell is in column A, so the
        // formula cell's child-list position (1) differs from its column (D) — the calcChain ref
        // must come from the cell's own address. The rendered cell must be a real formula element
        // in the spreadsheetml namespace, not an inline string.
        using var template = AutoDeletingPath.Create();
        using (var wb = new XLWorkbook())
        {
            var ws = wb.AddWorksheet("Sheet1");
            // the static formula makes the authoring library emit a calcChain part, so the
            // regeneration path runs; without one the template carries no chain to regenerate
            ws.Cell("F1").FormulaA1 = "1+1";
            ws.Cell("A5").Value = "{{items.Name}}";
            ws.Cell("B5").Value = "{{items.Qty}}";
            ws.Cell("A7").Value = "Total";
            ws.Cell("D7").Value = "$=SUM(B{{$enumrowstart}}:B{{$enumrowend}})";
            wb.SaveAs(template.FilePath);
        }

        using var path = AutoDeletingPath.Create();
        Dictionary<string, object?> data = new()
        {
            ["title"] = "FooCompany",
            ["items"] = new[]
            {
                new { Name = "A", Qty = 1 },
                new { Name = "B", Qty = 2 },
            }
        };
        MiniExcel.SaveAsByTemplate(path.ToString(), template.FilePath, data);

        using var zip = ZipFile.OpenRead(path.ToString());

        // the formula cell: <c r="D8"> (two items shift row 7 to 8) with a namespaced <f> child and no inlineStr type
        var doc = new XmlDocument();
        using (var sheet = zip.GetEntry("xl/worksheets/sheet1.xml")!.Open())
        {
            doc.Load(sheet);
        }

        var ns = new XmlNamespaceManager(doc.NameTable);
        ns.AddNamespace("x", "http://schemas.openxmlformats.org/spreadsheetml/2006/main");
        var formulaCell = doc.SelectSingleNode("//x:c[@r='D8']", ns) as XmlElement;
        Assert.NotNull(formulaCell);
        Assert.Equal("SUM(B5:B6)", formulaCell.SelectSingleNode("x:f", ns)?.InnerText);
        Assert.NotEqual("inlineStr", formulaCell.GetAttribute("t"));

        // and the regenerated calcChain points at the formula's real address, not the one derived
        // from the cell's position in the row (which would be column B here)
        using var chainReader = new StreamReader(zip.GetEntry("xl/calcChain.xml")!.Open());
        var chain = chainReader.ReadToEnd();
        Assert.Contains(@"r=""D8""", chain);
        Assert.DoesNotContain(@"r=""B8""", chain);
    }
}
