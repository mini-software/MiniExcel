using ExcelDataReader;
using MiniExcelLib.Core.Exceptions;
using MiniExcelLib.OpenXml.Tests.Utils;
using MiniExcelLib.Tests.Common.Utils;

namespace MiniExcelLib.OpenXml.Tests.Main;

public class MiniExcelOpenXmlImporterTests(ITestOutputHelper output)
{
    private readonly ITestOutputHelper _output = output;
    
    private readonly OpenXmlImporter _excelImporter =  MiniExcel.Importers.GetOpenXmlImporter();
    private readonly OpenXmlExporter _excelExporter =  MiniExcel.Exporters.GetOpenXmlExporter();
   
    static MiniExcelOpenXmlImporterTests()
    {
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
    }

    [Fact]
    public void CustomAttributeWihoutVaildPropertiesTest()
    {
        var path = PathHelper.GetFile("xlsx/TestCustomExcelColumnAttribute.xlsx");
        Assert.Throws<InvalidMappingException>(() =>  _excelImporter.Query<CustomAttributesWihoutVaildPropertiesTestPoco>(path).ToList());
    }

    [Fact]
    public void QueryCustomAttributesTest()
    {
        var path = PathHelper.GetFile("xlsx/TestCustomExcelColumnAttribute.xlsx");
        var rows = _excelImporter.Query<ExcelAttributeDemo>(path).ToList();

        Assert.Equal("Column1", rows[0].Test1);
        Assert.Equal("Column2", rows[0].Test2);
        Assert.Null(rows[0].Test3);
        Assert.Equal("Test7", rows[0].Test4);
        Assert.Null(rows[0].Test5);
        Assert.Null(rows[0].Test6);
        Assert.Equal("Test4", rows[0].Test7);
    }

    [Fact]
    public void QueryCastToIDictionary()
    {
        var path = PathHelper.GetFile("xlsx/TestCenterEmptyRow/TestCenterEmptyRow.xlsx");
        foreach (IDictionary<string, object> row in  _excelImporter.Query(path))
        {
            _ = row;
        }
    }

    [Fact]
    public void QueryRangeToIDictionary()
    {
        var path = PathHelper.GetFile("xlsx/TestCenterEmptyRow/TestCenterEmptyRow.xlsx");
        // tips：Only uppercase letters are effective
        var rows = _excelImporter.QueryRange(path, startCell: "A2", endCell: "C7")
            .Cast<IDictionary<string, object>>()
            .ToList();

        Assert.Equal(5, rows.Count);
        Assert.Equal(3, rows[0].Count);
        Assert.Equal(2d, rows[1]["B"]);
        Assert.Equal(null!, rows[2]["A"]);

        rows = _excelImporter.QueryRange(path, startRowIndex: 2, startColumnIndex: 1, endRowIndex: 7, endColumnIndex: 3)
            .Cast<IDictionary<string, object>>()
            .ToList();

        Assert.Equal(5, rows.Count);
        Assert.Equal(3, rows[0].Count);
        Assert.Equal(2d, rows[1]["B"]);
        Assert.Equal(null!, rows[2]["A"]);

        rows = _excelImporter.QueryRange(path, startRowIndex:2, startColumnIndex: 1, endRowIndex: 3)
          .Cast<IDictionary<string, object>>()
          .ToList();
        Assert.Equal(2, rows.Count);
        Assert.Equal(4, rows[0].Count);
        Assert.Equal(4d, rows[1]["D"]);

        rows = _excelImporter.QueryRange(path, startRowIndex: 2, startColumnIndex: 1, endColumnIndex: 3)
        .Cast<IDictionary<string, object>>()
        .ToList();
        Assert.Equal(5, rows.Count);
        Assert.Equal(3, rows[0].Count);
        Assert.Equal(3d, rows[3]["C"]);
    }

    [Fact]
    public void CenterEmptyRowsQueryTest()
    {
        var path = PathHelper.GetFile("xlsx/TestCenterEmptyRow/TestCenterEmptyRow.xlsx");
        using (var stream = File.OpenRead(path))
        {
            var rows = _excelImporter.Query(stream).ToList();

            Assert.Equal("a", rows[0].A);
            Assert.Equal("b", rows[0].B);
            Assert.Equal("c", rows[0].C);
            Assert.Equal("d", rows[0].D);

            Assert.Equal(1, rows[1].A);
            Assert.Null(rows[1].B);
            Assert.Equal(3, rows[1].C);
            Assert.Null(rows[1].D);

            Assert.Null(rows[2].A);
            Assert.Equal(2, rows[2].B);
            Assert.Null(rows[2].C);
            Assert.Equal(4, rows[2].D);

            Assert.Null(rows[3].A);
            Assert.Null(rows[3].B);
            Assert.Null(rows[3].C);
            Assert.Null(rows[3].D);

            Assert.Equal(1, rows[4].A);
            Assert.Null(rows[4].B);
            Assert.Equal(3, rows[4].C);
            Assert.Null(rows[4].D);

            Assert.Null(rows[5].A);
            Assert.Equal(2, rows[5].B);
            Assert.Null(rows[5].C);
            Assert.Equal(4, rows[5].D);
        }

        using (var stream = File.OpenRead(path))
        {
            var rows = _excelImporter.Query(stream, hasHeaderRow: true).ToList();

            Assert.Equal(1, rows[0].a);
            Assert.Null(rows[0].b);
            Assert.Equal(3, rows[0].c);
            Assert.Null(rows[0].d);

            Assert.Null(rows[1].a);
            Assert.Equal(2, rows[1].b);
            Assert.Null(rows[1].c);
            Assert.Equal(4, rows[1].d);

            Assert.Null(rows[2].a);
            Assert.Null(rows[2].b);
            Assert.Null(rows[2].c);
            Assert.Null(rows[2].d);

            Assert.Equal(1, rows[3].a);
            Assert.Null(rows[3].b);
            Assert.Equal(3, rows[3].c);
            Assert.Null(rows[3].d);

            Assert.Null(rows[4].a);
            Assert.Equal(2, rows[4].b);
            Assert.Null(rows[4].c);
            Assert.Equal(4, rows[4].d);
        }
    }

    [Fact]
    public void TestEmptyRowsQuerySelfClosingTag()
    {
        var path = PathHelper.GetFile("xlsx/TestEmptySelfClosingRow.xlsx");
        using var stream = File.OpenRead(path);
        var rows = _excelImporter.Query(stream).ToList();

        Assert.Null(rows[0].A);
        Assert.Equal(1, rows[1].A);
        Assert.Null(rows[2].A);
        Assert.Equal(2, rows[3].A);
        Assert.Null(rows[4].A);
        Assert.Null(rows[5].A);
        Assert.Null(rows[6].A);
        Assert.Null(rows[7].A);
        Assert.Null(rows[8].A);
        Assert.Equal(1, rows[9].A);
    }

    [Fact]
    public void TestDynamicQueryBasic_WithoutHead()
    {
        var path = PathHelper.GetFile("xlsx/TestDynamicQueryBasic_WithoutHead.xlsx");
        using var stream = File.OpenRead(path);
        var rows = _excelImporter.Query(stream).ToList();

        Assert.Equal("MiniExcel", rows[0].A);
        Assert.Equal(1, rows[0].B);
        Assert.Equal("Github", rows[1].A);
        Assert.Equal(2, rows[1].B);
    }

    [Fact]
    public void TestDynamicQueryBasic_hasHeaderRow()
    {
        var path = PathHelper.GetFile("xlsx/TestDynamicQueryBasic.xlsx");
        using (var stream = File.OpenRead(path))
        {
            var rows = _excelImporter.Query(stream, hasHeaderRow: true).ToList();

            Assert.Equal("MiniExcel", rows[0].Column1);
            Assert.Equal(1, rows[0].Column2);
            Assert.Equal("Github", rows[1].Column1);
            Assert.Equal(2, rows[1].Column2);
        }

        {
            var rows = _excelImporter.Query(path, hasHeaderRow: true).ToList();

            Assert.Equal("MiniExcel", rows[0].Column1);
            Assert.Equal(1, rows[0].Column2);
            Assert.Equal("Github", rows[1].Column1);
            Assert.Equal(2, rows[1].Column2);
        }
    }

    [Fact]
    public void QueryStrongTypeMapping_Test()
    {
        var path = PathHelper.GetFile("xlsx/TestTypeMapping.xlsx");
        using (var stream = File.OpenRead(path))
        {
            var rows = _excelImporter.Query<UserAccount>(stream).ToList();
            Assert.Equal(100, rows.Count);

            Assert.Equal(Guid.Parse("78DE23D2-DCB6-BD3D-EC67-C112BBC322A2"), rows[0].ID);
            Assert.Equal("Wade", rows[0].Name);
            Assert.Equal(DateTime.ParseExact("27/09/2020", "dd/MM/yyyy", CultureInfo.InvariantCulture), rows[0].BoD);
            Assert.Equal(36, rows[0].Age);
            Assert.False(rows[0].VIP);
            Assert.Equal(5019.12m, rows[0].Points);
            Assert.Equal(1, rows[0].IgnoredProperty);
        }

        {
            var rows = _excelImporter.Query<UserAccount>(path).ToList();
            Assert.Equal(100, rows.Count);

            Assert.Equal(Guid.Parse("78DE23D2-DCB6-BD3D-EC67-C112BBC322A2"), rows[0].ID);
            Assert.Equal("Wade", rows[0].Name);
            Assert.Equal(DateTime.ParseExact("27/09/2020", "dd/MM/yyyy", CultureInfo.InvariantCulture), rows[0].BoD);
            Assert.Equal(36, rows[0].Age);
            Assert.False(rows[0].VIP);
            Assert.Equal(5019.12m, rows[0].Points);
            Assert.Equal(1, rows[0].IgnoredProperty);
        }
    }

    [Fact]
    public void QueryRangeWithCellReferencesTest()
    {
        var path = PathHelper.GetFile("xlsx/TestTypeMapping.xlsx");

        using var stream = File.OpenRead(path);
        var rows = _excelImporter.QueryRange(stream, startCell: "B6", endCell: "D9")
            .Cast<IDictionary<string, object>>()
            .ToList();

        Assert.Equal(4, rows.Count);
        Assert.Equal(3, rows[0].Count);

        Assert.Equal("Raymond", rows[0]["B"]);
        Assert.Equal(new DateTime(2021, 12, 7), rows[0]["C"]);
        Assert.Equal(18d, rows[0]["D"]);
        Assert.Equal("Clarke", rows[1]["B"]);
        Assert.Equal(new DateTime(2021, 10, 16), rows[1]["C"]);
        Assert.Equal(60d, rows[1]["D"]);
        Assert.Equal("Eric", rows[2]["B"]);
        Assert.Equal(new DateTime(2020, 6, 24), rows[2]["C"]);
        Assert.Equal(30d, rows[2]["D"]);
    }

    [Fact]
    public void QueryRangeWithIndexCoordinatesTest()
    {
        var path = PathHelper.GetFile("xlsx/TestTypeMapping.xlsx");

        using var stream = File.OpenRead(path);
        var rows = _excelImporter.QueryRange(stream, startRowIndex: 6, startColumnIndex: 2, endRowIndex: 9, endColumnIndex: 4)
            .Cast<IDictionary<string, object>>()
            .ToList();

        Assert.Equal(4, rows.Count);
        Assert.Equal(3, rows[0].Count);

        Assert.Equal("Raymond", rows[0]["B"]);
        Assert.Equal(new DateTime(2021, 12, 7), rows[0]["C"]);
        Assert.Equal(18d, rows[0]["D"]);
        Assert.Equal("Clarke", rows[1]["B"]);
        Assert.Equal(new DateTime(2021, 10, 16), rows[1]["C"]);
        Assert.Equal(60d, rows[1]["D"]);
        Assert.Equal("Eric", rows[2]["B"]);
        Assert.Equal(new DateTime(2020, 6, 24), rows[2]["C"]);
        Assert.Equal(30d, rows[2]["D"]);
    }

    [Fact]
    public void QueryRangeWithCellReferencesAndHeaderTest()
    {
        var path = PathHelper.GetFile("xlsx/TestQueryRange.xlsx");
        var rows = _excelImporter.QueryRange(path, hasHeaderRow: true, startCell: "C3", endCell: "E6")
            .Cast<IDictionary<string, object>>()
            .ToList();

        Assert.Equal(3, rows.Count);
        Assert.Equal(3, rows[0].Count);

        Assert.Equal("Wade", rows[0]["Name"]);
        Assert.Equal(new DateTime(2020, 9, 27), rows[0]["BoD"]);
        Assert.Equal(36d, rows[0]["Age"]);
        Assert.Equal("Felix", rows[1]["Name"]);
        Assert.Equal(new DateTime(2020, 10, 25), rows[1]["BoD"]);
        Assert.Equal(44d, rows[1]["Age"]);
        Assert.Equal("Phelan", rows[2]["Name"]);
        Assert.Equal(new DateTime(2021, 4, 10), rows[2]["BoD"]);
        Assert.Equal(33d, rows[2]["Age"]);
    }

    [Fact]
    public void QueryRangeWithIndexCoordinatesAndHeaderTest()
    {
        var path = PathHelper.GetFile("xlsx/TestQueryRange.xlsx");
        var rows = _excelImporter.QueryRange(path, hasHeaderRow: true, startCell: "C3", endCell: "E6")
            .Cast<IDictionary<string, object>>()
            .ToList();

        Assert.Equal(3, rows.Count);
        Assert.Equal(3, rows[0].Count);

        Assert.Equal("Wade", rows[0]["Name"]);
        Assert.Equal(new DateTime(2020, 9, 27), rows[0]["BoD"]);
        Assert.Equal(36d, rows[0]["Age"]);
        Assert.Equal("Felix", rows[1]["Name"]);
        Assert.Equal(new DateTime(2020, 10, 25), rows[1]["BoD"]);
        Assert.Equal(44d, rows[1]["Age"]);
        Assert.Equal("Phelan", rows[2]["Name"]);
        Assert.Equal(new DateTime(2021, 4, 10), rows[2]["BoD"]);
        Assert.Equal(33d, rows[2]["Age"]);
    }

    [Fact]
    public void AutoCheckTypeTest()
    {
        var path = PathHelper.GetFile("xlsx/TestTypeMapping_AutoCheckFormat.xlsx");
        using var stream = File.OpenRead(path);
        var rows = _excelImporter.Query<AutoCheckType>(stream).ToList();
    }

    [Fact]
    public void UriMappingTest()
    {
        var path = PathHelper.GetFile("xlsx/TestUriMapping.xlsx");
        using var stream = File.OpenRead(path);
        var rows = _excelImporter.Query<ExcelUriDemo>(stream).ToList();

        Assert.Equal("Felix", rows[1].Name);
        Assert.Equal(44, rows[1].Age);
        Assert.Equal(new Uri("https://friendly-utilization.net"), rows[1].Url);
    }

    [Fact]
    public void TrimColumnNamesTest()
    {
        var path = PathHelper.GetFile("xlsx/TestTrimColumnNames.xlsx");
        var rows = _excelImporter.Query<SimpleAccount>(path).ToList();

        Assert.Equal("Raymond", rows[4].Name);
        Assert.Equal(18, rows[4].Age);
        Assert.Equal("sagittis.lobortis@leoMorbi.com", rows[4].Mail);
        Assert.Equal(8209.76m, rows[4].Points);
    }

    [Fact]
    public void TestDatetimeSpanFormat_ClosedXml()
    {
        var path = PathHelper.GetFile("xlsx/TestDatetimeSpanFormat_ClosedXml.xlsx");
        using var stream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read);

        var row = _excelImporter.Query(stream).First();
        var a = row.A;
        var b = row.B;
        Assert.Equal(DateTime.Parse("2021-03-20T23:39:42.3130000"), (DateTime)a);
        Assert.Equal(TimeSpan.FromHours(10), (TimeSpan)b);
    }

    [Theory]
    [InlineData("../../../../data/xlsx/ExcelDataReaderCollections/TestChess.xlsx")]
    [InlineData("../../../../data/xlsx/TestCenterEmptyRow/TestCenterEmptyRow.xlsx")]
    public void QueryDataReaderCheckTest(string path)
    {
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        using var fs = File.OpenRead(path);
        using var reader = ExcelReaderFactory.CreateReader(fs);
        var exceldatareaderResult = reader.AsDataSet();

        using var stream = File.OpenRead(path);
        var rows = _excelImporter.Query(stream).ToList();
        Assert.Equal(exceldatareaderResult.Tables[0].Rows.Count, rows.Count);

        foreach (IDictionary<string, object?> row in rows)
        {
            var rowIndex = rows.IndexOf(row);
            foreach (var (key, value) in row)
            {
                var eV = exceldatareaderResult.Tables[0].Rows[rowIndex][SheetHelper.GetColumnIndex(key)];
                var v = value ?? DBNull.Value;
                Assert.Equal(eV, v);
            }
        }
    }

    [Fact]
    public void QuerySheetWithoutRAttribute()
    {
        var path = PathHelper.GetFile("xlsx/TestWihoutRAttribute.xlsx");
        using var stream = File.OpenRead(path);
        var rows = _excelImporter.Query(stream).ToList();
        var keys = (rows.First() as IDictionary<string, object>)!.Keys;

        Assert.Equal(2, rows.Count);
        Assert.Equal(5, keys.Count);

        Assert.Equal(1, rows[0].A);
        Assert.Null(rows[0].C);
        Assert.Null(rows[0].D);
        Assert.Null(rows[0].E);

        Assert.Equal(1, rows[1].A);
        Assert.Equal("\"<>+}{\\nHello World", rows[1].B);
        Assert.Equal(true, rows[1].C);
        Assert.Equal("2021-03-16T19:10:21", rows[1].D);
    }

    [Fact]
    public void FixDimensionJustOneColumnParsingError_Test()
    {
        var path = PathHelper.GetFile("xlsx/TestDimensionC3.xlsx");
        using var stream = File.OpenRead(path);
        var rows = _excelImporter.Query(stream).ToList();
        var keys = ((IDictionary<string, object>)rows.First()).Keys;
        Assert.Equal(3, keys.Count);
        Assert.Equal(2, rows.Count);
    }

    [Fact]
    public void QueryByStrongTypeParameterTest()
    {
        using var path = AutoDeletingPath.Create();
        List<Demo> values =
        [
            new() { Column1 = "MiniExcel", Column2 = 1 },
            new() { Column1 = "Github", Column2 = 2 }
        ];
         _excelExporter.Export(path.ToString(), values);

        using var stream = File.OpenRead(path.ToString());
        var rows = _excelImporter.Query(stream, hasHeaderRow: true).ToList();

        Assert.Equal("MiniExcel", rows[0].Column1);
        Assert.Equal(1, rows[0].Column2);
        Assert.Equal("Github", rows[1].Column1);
        Assert.Equal(2, rows[1].Column2);
    }

    [Fact]
    public void QueryByDictionaryStringAndObjectParameterTest()
    {
        using var path = AutoDeletingPath.Create();
        List<Dictionary<string, object>> values =
        [
            new() { { "Column1", "MiniExcel" }, { "Column2", 1 } },
            new() { { "Column1", "Github" }, { "Column2", 2 } }
        ];
         _excelExporter.Export(path.ToString(), values);

        using var stream = File.OpenRead(path.ToString());
        var rows = _excelImporter.Query(stream, hasHeaderRow: true).ToList();

        Assert.Equal("MiniExcel", rows[0].Column1);
        Assert.Equal(1, rows[0].Column2);
        Assert.Equal("Github", rows[1].Column1);
        Assert.Equal(2, rows[1].Column2);
    }

    [Fact]
    public void TestStirctOpenXml()
    {
        var path = PathHelper.GetFile("xlsx/TestStrictOpenXml.xlsx");
        var columns = _excelImporter.GetColumnNames (path);
        Assert.Equal(["A", "B", "C"], columns);

        var rows = _excelImporter.Query(path).ToList();
        Assert.Equal(rows[0].A, "title1");
        Assert.Equal(rows[0].B, "title2");
        Assert.Equal(rows[0].C, "title3");
        Assert.Equal(rows[1].A, "value1");
        Assert.Equal(rows[1].B, "value2");
        Assert.Equal(rows[1].C, "value3");
    }

    [Fact]
    public void DateOnlySupportTest()
    {
        var query = _excelImporter.Query<DateOnlyTest>(PathHelper.GetFile("xlsx/TestDateOnlyMapping.xlsx")).ToList();

        Assert.Equal(new DateOnly(2020, 9, 27), query[0].Date);
        Assert.Equal(new DateOnly(2020, 10, 25), query[1].Date);
        Assert.Equal(new DateOnly(2021, 10, 4), query[2].Date);

        Assert.Equal(new DateOnly(2020, 9, 27), query[0].DateWithFormat);
        Assert.Equal(new DateOnly(2020, 10, 25), query[1].DateWithFormat);
        Assert.Equal(new DateOnly(2020, 6, 1), query[7].DateWithFormat);
    }

    [Fact]
    public void SheetDimensionsTest()
    {
        var path1 = PathHelper.GetFile("xlsx/TestTypeMapping.xlsx");
        var dim1 = _excelImporter.GetSheetDimensions(path1);
        Assert.Equal("A1", dim1[0].StartCell);
        Assert.Equal("H101", dim1[0].EndCell);
        Assert.Equal(101, dim1[0].Rows.Count);
        Assert.Equal(8, dim1[0].Columns.Count);
        Assert.Equal(1, dim1[0].Rows.StartIndex);
        Assert.Equal(101, dim1[0].Rows.EndIndex);
        Assert.Equal(1, dim1[0].Columns.StartIndex);
        Assert.Equal(8, dim1[0].Columns.EndIndex);

        var path2 = PathHelper.GetFile("xlsx/TestNoDimension.xlsx");
        var dim2 = _excelImporter.GetSheetDimensions(path2);
        Assert.Equal(101, dim2[0].Rows.Count);
        Assert.Equal(7, dim2[0].Columns.Count);
        Assert.Equal(1, dim2[0].Rows.StartIndex);
        Assert.Equal(101, dim2[0].Rows.EndIndex);
        Assert.Equal(1, dim2[0].Columns.StartIndex);
        Assert.Equal(7, dim2[0].Columns.EndIndex);
    }

    [Fact]
    public void SheetDimensionsTest_MultiSheet()
    {
        var path = PathHelper.GetFile("xlsx/TestMultiSheet.xlsx");
        var dim = _excelImporter.GetSheetDimensions(path);

        Assert.Equal("A1", dim[0].StartCell);
        Assert.Equal("D12", dim[0].EndCell);
        Assert.Equal(12, dim[0].Rows.Count);
        Assert.Equal(4, dim[0].Columns.Count);
        Assert.Equal(1, dim[0].Rows.StartIndex);
        Assert.Equal(12, dim[0].Rows.EndIndex);
        Assert.Equal(1, dim[0].Columns.StartIndex);
        Assert.Equal(4, dim[0].Columns.EndIndex);

        Assert.Equal("A1", dim[1].StartCell);
        Assert.Equal("D12", dim[1].EndCell);
        Assert.Equal(12, dim[1].Rows.Count);
        Assert.Equal(4, dim[1].Columns.Count);
        Assert.Equal(1, dim[1].Rows.StartIndex);
        Assert.Equal(12, dim[1].Rows.EndIndex);
        Assert.Equal(1, dim[1].Columns.StartIndex);
        Assert.Equal(4, dim[1].Columns.EndIndex);

        Assert.Equal("A1", dim[2].StartCell);
        Assert.Equal("B5", dim[2].EndCell);
        Assert.Equal(5, dim[2].Rows.Count);
        Assert.Equal(2, dim[2].Columns.Count);
        Assert.Equal(1, dim[2].Rows.StartIndex);
        Assert.Equal(5, dim[2].Rows.EndIndex);
        Assert.Equal(1, dim[2].Columns.StartIndex);
        Assert.Equal(2, dim[2].Columns.EndIndex);
    }

    [Fact]
    public void ExportAndQueryFieldsStrongMappingTest()
    {
        using var path = AutoDeletingPath.Create();
        var input = Enumerable.Range(1, 3)
            .Select(i => new ExcelFieldMappingTest
            {
                Test1 = $"T{i}",
                Test2 = i,
                Test = i + (decimal)i / 10
            });

        _excelExporter.Export(path.ToString(), input);

        var rows = _excelImporter.Query<ExcelFieldMappingTest>(path.ToString()).ToList();
        Assert.Equal(3, rows.Count);
        Assert.Equal("T1", rows[0].Test1);
        Assert.Equal(1, rows[0].Test2);
        Assert.Equal(1.1m, rows[0].Test);
    }

    [Fact]
    public void QueryFieldsAsDynamicTest()
    {
        using var path = AutoDeletingPath.Create();
        ExcelFieldMappingTest[] input = [new() { Test1 = "X1", Test2 = 5, Test = 7.3m }];
        
        _excelExporter.Export(path.ToString(), input);

        var rows = _excelImporter.Query(path.ToString(), true).ToList();
        var first = rows[0] as IDictionary<string, object>;

        // Column headers should include the column names from field attributes 
        Assert.Contains("Column1", first!.Keys);
        Assert.Contains("Column2", first.Keys);
    }
}
