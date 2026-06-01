using MiniExcelLib.OpenXml.Tests.Utils;
using MiniExcelLib.Tests.Common.Utils;

namespace MiniExcelLib.OpenXml.Tests.Issues;

public class MiniExcelGiteeIssuesTests
{
    private readonly OpenXmlImporter _excelImporter =  MiniExcel.Importers.GetOpenXmlImporter();
    private readonly OpenXmlExporter _excelExporter =  MiniExcel.Exporters.GetOpenXmlExporter();
    private readonly OpenXmlTemplater _excelTemplater =  MiniExcel.Templaters.GetOpenXmlTemplater();

    [Fact]
    public void TestIssueI4ZYUU()
    {
        using var path = AutoDeletingPath.Create();
        
        var dt = new DateTime(2022, 10, 15);
        TestIssueI4ZYUUDto[] value = [new() { MyProperty = "1", MyProperty2 = dt }];
        _excelExporter.Export(path.ToString(), value);

        using var workbook = new ClosedXML.Excel.XLWorkbook(path.ToString());
        var ws = workbook.Worksheet(1);

        Assert.Equal(dt, ws.Cell(2, "B").Value.GetDateTime());
        Assert.Equal("2022-10", ws.Cell(2, "B").GetFormattedString());
        Assert.True(ws.Column("A").Width > 0);
        Assert.True(ws.Column("B").Width > 0);
    }

    [Fact]
    public void TestIssueI4YCLQ_2()
    {
        var path = PathHelper.GetFile("xlsx/TestIssueI4YCLQ_2.xlsx");
        var rows = _excelImporter.Query<TestIssueI4YCLQ_2Dto>(path, startCell: "B2").ToList();

        Assert.Null(rows[0].站点编码);
        Assert.Equal("N1", rows[0].站址名称);
        Assert.Equal("a", rows[0].值1);
        Assert.Equal("b", rows[0].值2);
        Assert.Equal("c", rows[0].值3);
        Assert.Equal("A1", rows[0].资源ID);
        Assert.Equal("A", rows[0].值4);
        Assert.Equal("B", rows[0].值5);
        Assert.Equal("C", rows[0].值6);
        Assert.Null(rows[0].值7);
        Assert.Null(rows[0].值8);
    }

    [Fact]
    public void TestIssueI4WM67()
    {
        using var path = AutoDeletingPath.Create();
        var templatePath = PathHelper.GetFile("xlsx/TestIssueI4WM67.xlsx");
        var value = new Dictionary<string, object>
        {
            ["users"] = Array.Empty<TestIssueI4WM67Dto>()
        };
         _excelTemplater.FillTemplate(path.ToString(), templatePath, value);
        var rows = _excelImporter.Query(path.ToString()).ToList();
        Assert.Single(rows);
    }

    [Fact]
    public void TestIssueI4WXFB()
    {
        {
            using var path = AutoDeletingPath.Create();
            var templatePath = PathHelper.GetFile("xlsx/TestIssueI4WXFB.xlsx");
            var value = new Dictionary<string, object>
            {
                ["Name"] = "Jack",
                ["Amount"] = 1000,
                ["Department"] = "HR"
            };
             _excelTemplater.FillTemplate(path.ToString(), templatePath, value);
        }

        {
            var config = new OpenXmlConfiguration
            {
                IgnoreTemplateParameterMissing = false
            };
            using var path = AutoDeletingPath.Create();
            var templatePath = PathHelper.GetFile("xlsx/TestIssueI4WXFB.xlsx");
            var value = new Dictionary<string, object>
            {
                ["Name"] = "Jack",
                ["Amount"] = 1000,
                ["Department"] = "HR"
            };
            Assert.Throws<KeyNotFoundException>(() => _excelTemplater.FillTemplate(path.ToString(), templatePath, value, configuration: config));
        }
    }

    [Fact]
    public void TestIssueI4TXGT()
    {
        using var file = AutoDeletingPath.Create();
        var path = file.ToString();
        var value = new[] { new TestIssueI4TXGTDto { ID = 1, Name = "Apple", Spc = "X", Up = 6999 } };

        _excelExporter.Export(path, value);
        var rows1 = _excelImporter.Query(path).ToList();
        Assert.Equal("ID", rows1[0].A);
        Assert.Equal("Name", rows1[0].B);
        Assert.Equal("Specification", rows1[0].C);
        Assert.Equal("Unit Price", rows1[0].D);

        var rows2 = _excelImporter.Query<TestIssueI4TXGTDto>(path).ToList();
        Assert.Equal(1, rows2[0].ID);
        Assert.Equal("Apple", rows2[0].Name);
        Assert.Equal("X", rows2[0].Spc);
        Assert.Equal(6999, rows2[0].Up);
    }

    // https://gitee.com/dotnetchina/MiniExcel/issues/I4HL54
    [Fact]
    public void TestIssueI4HL54()
    {
        using var cn = Db.GetConnection();

        using var reader = cn.ExecuteReader(@"select 'Hello World1' Text union all select 'Hello World2'");
        var templatePath = PathHelper.GetFile("xlsx/TestIssueI4HL54_Template.xlsx");
        using var path = AutoDeletingPath.Create();
        var value = new Dictionary<string, object>
        {
            { "Texts",reader}
        };
        _excelTemplater.FillTemplate(path.ToString(), templatePath, value);

        var rows = _excelImporter.Query(path.ToString(), true).ToList();
        Assert.Equal("Hello World1", rows[0].Text);
        Assert.Equal("Hello World2", rows[1].Text);
    }

    // SaveAsByTemplate if there is & in the cell value, it will be &amp;
    // https://gitee.com/dotnetchina/MiniExcel/issues/I4DQUN
    [Fact]
    public void TestIssueI4DQUN()
    {
        var templatePath = PathHelper.GetFile("xlsx/TestIssueI4DQUN.xlsx");
        using var path = AutoDeletingPath.Create();
        var value = new Dictionary<string, object>
        {
            { "Title", "Hello & World < , > , \" , '" },
            { "Details", new[] { new { Value = "Hello & Value < , > , \" , '" } } },
        };
        _excelTemplater.FillTemplate(path.ToString(), templatePath, value);

        var sheetXml = SheetHelper.GetZipFileContent(path.ToString(), "xl/worksheets/sheet1.xml");
        // Template now uses inlineStr format (<is><t>...</t></is>) instead of SharedStrings (<v>...</v>)
        Assert.Contains("<t>Hello &amp; World &lt; , &gt; , \" , '</t>", sheetXml);
        Assert.Contains("<t>Hello &amp; Value &lt; , &gt; , \" , '</t>", sheetXml);
    }

    [Fact]
    public void TestIssueI49RYZ()
    {
        DescriptionEnumDto[] values =
        [
            new() { Name = "Jack", UserType = DescriptionEnum.V1 },
            new() { Name = "Leo", UserType = DescriptionEnum.V2 },
            new() { Name = "Henry", UserType = DescriptionEnum.V3 },
            new() { Name = "Lisa", UserType = null }
        ];

        using var path = AutoDeletingPath.Create();
        _excelExporter.Export(path.ToString(), values);
        var rows = _excelImporter.Query(path.ToString(), true).ToList();
        Assert.Equal("General User", rows[0].UserType);
        Assert.Equal("General Administrator", rows[1].UserType);
        Assert.Equal("Super Administrator", rows[2].UserType);
        Assert.Null(rows[3].UserType);
    }

    // https://gitee.com/dotnetchina/MiniExcel/issues/I40QA5
    [Fact]
    public void TestIssueI40QA5()
    {
        {
            var path = PathHelper.GetFile("/xlsx/TestIssueI40QA5_1.xlsx");
            var rows = _excelImporter.Query<TestIssueI40QA5Dto>(path).ToList();
            Assert.Equal("E001", rows[0].Empno);
            Assert.Equal("E002", rows[1].Empno);
        }
        {
            var path = PathHelper.GetFile("/xlsx/TestIssueI40QA5_2.xlsx");
            var rows = _excelImporter.Query<TestIssueI40QA5Dto>(path).ToList();
            Assert.Equal("E001", rows[0].Empno);
            Assert.Equal("E002", rows[1].Empno);
        }
        {
            var path = PathHelper.GetFile("/xlsx/TestIssueI40QA5_3.xlsx");
            var rows = _excelImporter.Query<TestIssueI40QA5Dto>(path).ToList();
            Assert.Equal("E001", rows[0].Empno);
            Assert.Equal("E002", rows[1].Empno);
        }
        {
            var path = PathHelper.GetFile("/xlsx/TestIssueI40QA5_4.xlsx");
            var rows = _excelImporter.Query<TestIssueI40QA5Dto>(path).ToList();
            Assert.Null(rows[0].Empno);
            Assert.Null(rows[1].Empno);
        }
    }

    // Semicolon expected
    [Fact]
    public void TestIssueI45TF5_2()
    {
        var value = new[] { new Dictionary<string, object> { { "Col1&Col2", "V1&V2" } } };

        using var path1 = AutoDeletingPath.Create();
        _excelExporter.Export(path1.ToString(), value);
        //System.Xml.XmlException : '<' is an unexpected token. The expected token is ';'.
        SheetHelper.GetZipFileContent(path1.ToString(), "xl/worksheets/sheet1.xml"); //check illegal format or not

        using var dt = new DataTable();
        dt.Columns.Add("Col1&Col2");
        dt.Rows.Add("V1&V2");

        using var path2 = AutoDeletingPath.Create();
        _excelExporter.Export(path2.FilePath, dt);
        //System.Xml.XmlException : '<' is an unexpected token. The expected token is ';'.
        SheetHelper.GetZipFileContent(path2.FilePath, "xl/worksheets/sheet1.xml"); //check illegal format or not
    }

    [Fact]
    public void TestIssueI45TF5()
    {
        using var path = AutoDeletingPath.Create();
        _excelExporter.Export(path.ToString(), new[] { new { C1 = "1&2;3,4", C2 = "1&2;3,4" } });
        var sheet1Xml = SheetHelper.GetZipFileContent(path.ToString(), "xl/worksheets/sheet1.xml");
        Assert.DoesNotContain("<x:cols>", sheet1Xml);
    }
    
    [Fact]
    public void TestIssueI3X2ZL()
    {
        try
        {
            var path = PathHelper.GetFile("xlsx/TestIssueI3X2ZL_datetime_error.xlsx");
            var rows = _excelImporter.Query<IssueI3X2ZLDTO>(path, startCell: "B3").ToList();
        }
        catch (InvalidCastException ex)
        {
            Assert.Equal(
                "The value error cannot be assigned to type DateTime.",
                ex.Message
            );
        }

        try
        {
            var path = PathHelper.GetFile("xlsx/TestIssueI3X2ZL_int_error.xlsx");
            var rows = _excelImporter.Query<IssueI3X2ZLDTO>(path).ToList();
        }
        catch (InvalidCastException ex)
        {
            Assert.Equal(
                "The value error cannot be assigned to type Int32.",
                ex.Message
            );
        }
    }

    // https://gitee.com/dotnetchina/MiniExcel/issues/I3OSKV
    // When exporting, the pure numeric string will be forcibly converted to a numeric type, resulting in the loss of the end data
    [Fact]
    public void IssueI3OSKV()
    {
        using var path1 = AutoDeletingPath.Create();
        var value1 = new[] { new { Test = "12345678901234567890" } };
        _excelExporter.Export(path1.ToString(), value1);

        var result1 = _excelImporter.Query(path1.ToString(), true).First();
        Assert.Equal("12345678901234567890", result1.Test);

        using var path2 = AutoDeletingPath.Create();
        var value2 = new[] { new { Test = 123456.789 } };
        _excelExporter.Export(path2.ToString(), value2);

        var result2 = _excelImporter.Query(path2.ToString(), true).First();
        Assert.Equal(123456.789, result2.Test);
    }

    // https://gitee.com/dotnetchina/MiniExcel/issues/I50VD5
    [Fact]
    public void IssueI50VD5()
    {
        using var file = AutoDeletingPath.Create();
        var path = file.ToString();

        List<dynamic> list1 =
        [
            new { Name = "github", Image = File.ReadAllBytes(PathHelper.GetFile("images/github_logo.png")) },
            new { Name = "google", Image = File.ReadAllBytes(PathHelper.GetFile("images/google_logo.png")) },
            new { Name = "microsoft", Image = File.ReadAllBytes(PathHelper.GetFile("images/microsoft_logo.png")) },
            new { Name = "reddit", Image = File.ReadAllBytes(PathHelper.GetFile("images/reddit_logo.png")) },
            new { Name = "stackoverflow", Image = File.ReadAllBytes(PathHelper.GetFile("images/stackoverflow_logo.png")) }
        ];

        List<dynamic> list2 =
        [
            new { Id = 1, Name = "github", Image = File.ReadAllBytes(PathHelper.GetFile("images/github_logo.png")) },
            new { Id = 2, Name = "google", Image = File.ReadAllBytes(PathHelper.GetFile("images/google_logo.png")) }
        ];

        var sheets = new Dictionary<string, object>
        {
            ["A"] = list1,
            ["B"] = list2,
        };
         _excelExporter.Export(path, sheets);

        {
            Assert.Contains("/xl/media/", SheetHelper.GetZipFileContent(path, "xl/drawings/_rels/drawing1.xml.rels"));
            Assert.Contains("ext cx=\"609600\" cy=\"190500\"", SheetHelper.GetZipFileContent(path, "xl/drawings/drawing1.xml"));
            Assert.Contains("/xl/drawings/drawing1.xml", SheetHelper.GetZipFileContent(path, "[Content_Types].xml"));
            Assert.Contains("drawing r:id=\"drawing1\"", SheetHelper.GetZipFileContent(path, "xl/worksheets/sheet1.xml"));
            Assert.Contains("../drawings/drawing1.xml", SheetHelper.GetZipFileContent(path, "xl/worksheets/_rels/sheet1.xml.rels"));

            Assert.Contains("/xl/media/", SheetHelper.GetZipFileContent(path, "xl/drawings/_rels/drawing2.xml.rels"));
            Assert.Contains("ext cx=\"609600\" cy=\"190500\"", SheetHelper.GetZipFileContent(path, "xl/drawings/drawing2.xml"));
            Assert.Contains("/xl/drawings/drawing1.xml", SheetHelper.GetZipFileContent(path, "[Content_Types].xml"));
            Assert.Contains("drawing r:id=\"drawing2\"", SheetHelper.GetZipFileContent(path, "xl/worksheets/sheet2.xml"));
            Assert.Contains("../drawings/drawing2.xml", SheetHelper.GetZipFileContent(path, "xl/worksheets/_rels/sheet2.xml.rels"));
        }
    }
}