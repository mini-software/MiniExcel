using ClosedXML.Excel;
using ExcelDataReader;
using MiniExcelLib.OpenXml.Tests.Utils;
using MiniExcelLib.Tests.Common.Utils;
using FileHelper = MiniExcelLib.OpenXml.Tests.Utils.FileHelper;

namespace MiniExcelLib.OpenXml.Tests;

public class MiniExcelOpenXmlTests(ITestOutputHelper output)
{
    private readonly ITestOutputHelper _output = output;
    
    private readonly OpenXmlImporter _excelImporter =  MiniExcel.Importers.GetOpenXmlImporter();
    private readonly OpenXmlExporter _excelExporter =  MiniExcel.Exporters.GetOpenXmlExporter();
   
    [Fact]
    public void GetColumnsTest()
    {
        var tmPath = PathHelper.GetFile("xlsx/TestTypeMapping.xlsx");
        var tePath = PathHelper.GetFile("xlsx/TestEmpty.xlsx");
        {
            var columns =  _excelImporter.GetColumnNames (tmPath);
            Assert.Equal(["A", "B", "C", "D", "E", "F", "G", "H"], columns);
        }

        {
            var columns =  _excelImporter.GetColumnNames (tmPath);
            Assert.Equal(8, columns.Count);
        }

        {
            var columns =  _excelImporter.GetColumnNames (tePath);
            Assert.Empty(columns);
        }
    }

    [Fact]
    public void SaveAsControlChracter()
    {
        using var path = AutoDeletingPath.Create();
        char[] chars =
        [
            '\u0000','\u0001','\u0002','\u0003','\u0004','\u0005','\u0006','\u0007','\u0008',
            '\u0009', //<HT>
            '\u000A', //<LF>
            '\u000B','\u000C',
            '\u000D', //<CR>
            '\u000E','\u000F','\u0010','\u0011','\u0012','\u0013','\u0014','\u0015','\u0016',
            '\u0017','\u0018','\u0019','\u001A','\u001B','\u001C','\u001D','\u001E','\u001F','\u007F'
        ];
        var input = chars.Select(s => new { Test = s.ToString() });
         _excelExporter.Export(path.ToString(), input);

        var rows2 =  _excelImporter.Query(path.ToString(), true).Select(s => s.Test).ToArray();
        var rows1 =  _excelImporter.Query<SaveAsControlChracterVO>(path.ToString()).Select(s => s.Test).ToArray();
    }

    private class SaveAsControlChracterVO
    {
        public string Test { get; set; }
    }

    private class ExcelAttributeDemo
    {
        [MiniExcelColumnName("Column1")]
        public string Test1 { get; set; }
        [MiniExcelColumnName("Column2")]
        public string Test2 { get; set; }
        [MiniExcelIgnore]
        public string Test3 { get; set; }
        [MiniExcelColumnIndex("I")] // system will convert "I" to 8 index
        public string Test4 { get; set; }
        public string Test5 { get; } //wihout set will ignore
        public string Test6 { get; private set; } //un-public set will ignore
        [MiniExcelColumnIndex(3)] // start with 0
        public string Test7 { get; set; }
    }

    [Fact]
    public void CustomAttributeWihoutVaildPropertiesTest()
    {
        var path = PathHelper.GetFile("xlsx/TestCustomExcelColumnAttribute.xlsx");
        Assert.Throws<InvalidOperationException>(() =>  _excelImporter.Query<CustomAttributesWihoutVaildPropertiesTestPoco>(path).ToList());
    }

    [Fact]
    public void QueryCustomAttributesTest()
    {
        var path = PathHelper.GetFile("xlsx/TestCustomExcelColumnAttribute.xlsx");
        var rows =  _excelImporter.Query<ExcelAttributeDemo>(path).ToList();

        Assert.Equal("Column1", rows[0].Test1);
        Assert.Equal("Column2", rows[0].Test2);
        Assert.Null(rows[0].Test3);
        Assert.Equal("Test7", rows[0].Test4);
        Assert.Null(rows[0].Test5);
        Assert.Null(rows[0].Test6);
        Assert.Equal("Test4", rows[0].Test7);
    }

    [Fact]
    public void SaveAsCustomAttributesTest()
    {
        using var path = AutoDeletingPath.Create();
        var input = Enumerable.Range(1, 3)
            .Select(_ => new ExcelAttributeDemo
            {
                Test1 = "Test1",
                Test2 = "Test2",
                Test3 = "Test3",
                Test4 = "Test4",
            });

         _excelExporter.Export(path.ToString(), input);
        var rows =  _excelImporter.Query(path.ToString(), true).ToList();
        var first = rows[0] as IDictionary<string, object>;

        Assert.Equal(3, rows.Count);
        Assert.Equal(["Column1", "Column2", "Test5", "Test7", "Test6", "Test4"], first?.Keys);
        Assert.Equal("Test1", rows[0].Column1);
        Assert.Equal("Test2", rows[0].Column2);
        Assert.Equal("Test4", rows[0].Test4);
        Assert.Null(rows[0].Test5);
        Assert.Null(rows[0].Test6);
    }

    private class CustomAttributesWihoutVaildPropertiesTestPoco
    {
        [MiniExcelIgnore]
        public string Test3 { get; set; }
        public string Test5 { get; }
        public string Test6 { get; private set; }
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
        var rows =  _excelImporter.QueryRange(path, startCell: "A2", endCell: "C7")
            .Cast<IDictionary<string, object>>()
            .ToList();

        Assert.Equal(5, rows.Count);
        Assert.Equal(3, rows[0].Count);
        Assert.Equal(2d, rows[1]["B"]);
        Assert.Equal(null!, rows[2]["A"]);

        rows =  _excelImporter.QueryRange(path, startRowIndex: 2, startColumnIndex: 1, endRowIndex: 7, endColumnIndex: 3)
            .Cast<IDictionary<string, object>>()
            .ToList();

        Assert.Equal(5, rows.Count);
        Assert.Equal(3, rows[0].Count);
        Assert.Equal(2d, rows[1]["B"]);
        Assert.Equal(null!, rows[2]["A"]);

        rows =  _excelImporter.QueryRange(path, startRowIndex:2, startColumnIndex: 1, endRowIndex: 3)
          .Cast<IDictionary<string, object>>()
          .ToList();
        Assert.Equal(2, rows.Count);
        Assert.Equal(4, rows[0].Count);
        Assert.Equal(4d, rows[1]["D"]);

        rows =  _excelImporter.QueryRange(path, startRowIndex: 2, startColumnIndex: 1, endColumnIndex: 3)
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
            var rows =  _excelImporter.Query(stream).ToList();

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
            var rows =  _excelImporter.Query(stream, useHeaderRow: true).ToList();

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
        var rows =  _excelImporter.Query(stream).ToList();

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
        var rows =  _excelImporter.Query(stream).ToList();

        Assert.Equal("MiniExcel", rows[0].A);
        Assert.Equal(1, rows[0].B);
        Assert.Equal("Github", rows[1].A);
        Assert.Equal(2, rows[1].B);
    }

    [Fact]
    public void TestDynamicQueryBasic_useHeaderRow()
    {
        var path = PathHelper.GetFile("xlsx/TestDynamicQueryBasic.xlsx");
        using (var stream = File.OpenRead(path))
        {
            var rows =  _excelImporter.Query(stream, useHeaderRow: true).ToList();

            Assert.Equal("MiniExcel", rows[0].Column1);
            Assert.Equal(1, rows[0].Column2);
            Assert.Equal("Github", rows[1].Column1);
            Assert.Equal(2, rows[1].Column2);
        }

        {
            var rows =  _excelImporter.Query(path, useHeaderRow: true).ToList();

            Assert.Equal("MiniExcel", rows[0].Column1);
            Assert.Equal(1, rows[0].Column2);
            Assert.Equal("Github", rows[1].Column1);
            Assert.Equal(2, rows[1].Column2);
        }
    }

    private class DemoPocoHelloWorld
    {
        public string HelloWorld1 { get; set; }
    }

    public class UserAccount
    {
        public Guid ID { get; set; }
        public string Name { get; set; }
        public DateTime BoD { get; set; }
        public int Age { get; set; }
        public bool VIP { get; set; }
        public decimal Points { get; set; }
        public int IgnoredProperty => 1;
    }

    [Fact]
    public void QueryStrongTypeMapping_Test()
    {
        var path = PathHelper.GetFile("xlsx/TestTypeMapping.xlsx");
        using (var stream = File.OpenRead(path))
        {
            var rows =  _excelImporter.Query<UserAccount>(stream).ToList();
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
            var rows =  _excelImporter.Query<UserAccount>(path).ToList();
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

    private class AutoCheckType
    {
        public Guid? Guid { get; set; }
        public bool? Bool { get; set; }
        public DateTime? Datetime { get; set; }
        public string String { get; set; }
    }

    [Fact]
    public void AutoCheckTypeTest()
    {
        var path = PathHelper.GetFile("xlsx/TestTypeMapping_AutoCheckFormat.xlsx");
        using var stream = FileHelper.OpenRead(path);
        var rows =  _excelImporter.Query<AutoCheckType>(stream).ToList();
    }

    private class ExcelUriDemo
    {
        public string Name { get; set; }
        public int Age { get; set; }
        public Uri Url { get; set; }
    }

    [Fact]
    public void UriMappingTest()
    {
        var path = PathHelper.GetFile("xlsx/TestUriMapping.xlsx");
        using var stream = File.OpenRead(path);
        var rows =  _excelImporter.Query<ExcelUriDemo>(stream).ToList();

        Assert.Equal("Felix", rows[1].Name);
        Assert.Equal(44, rows[1].Age);
        Assert.Equal(new Uri("https://friendly-utilization.net"), rows[1].Url);
    }

    private class SimpleAccount
    {
        public string Name { get; set; }
        public int Age { get; set; }
        public string Mail { get; set; }
        public decimal Points { get; set; }
    }
    [Fact]
    public void TrimColumnNamesTest()
    {
        var path = PathHelper.GetFile("xlsx/TestTrimColumnNames.xlsx");
        var rows =  _excelImporter.Query<SimpleAccount>(path).ToList();

        Assert.Equal("Raymond", rows[4].Name);
        Assert.Equal(18, rows[4].Age);
        Assert.Equal("sagittis.lobortis@leoMorbi.com", rows[4].Mail);
        Assert.Equal(8209.76m, rows[4].Points);
    }

    [Fact]
    public void TestDatetimeSpanFormat_ClosedXml()
    {
        var path = PathHelper.GetFile("xlsx/TestDatetimeSpanFormat_ClosedXml.xlsx");
        using var stream = FileHelper.OpenRead(path);

        var row =  _excelImporter.Query(stream).First();
        var a = row.A;
        var b = row.B;
        Assert.Equal(DateTime.Parse("2021-03-20T23:39:42.3130000"), (DateTime)a);
        Assert.Equal(TimeSpan.FromHours(10), (TimeSpan)b);
    }

    [Fact]
    public void LargeFileQueryStrongTypeMapping_Test()
    {
        const string path = "../../../../../benchmarks/MiniExcel.Benchmarks/Test1,000,000x10.xlsx";
        using (var stream = File.OpenRead(path))
        {
            var rows =  _excelImporter.Query<DemoPocoHelloWorld>(stream).Take(2).ToList();

            Assert.Equal("HelloWorld2", rows[0].HelloWorld1);
            Assert.Equal("HelloWorld3", rows[1].HelloWorld1);
        }
        {
            var rows =  _excelImporter.Query<DemoPocoHelloWorld>(path).Take(2).ToList();

            Assert.Equal("HelloWorld2", rows[0].HelloWorld1);
            Assert.Equal("HelloWorld3", rows[1].HelloWorld1);
        }
    }

    [Theory]
    [InlineData("../../../../data/xlsx/ExcelDataReaderCollections/TestChess.xlsx")]
    [InlineData("../../../../data/xlsx/TestCenterEmptyRow/TestCenterEmptyRow.xlsx")]
    public void QueryDataReaderCheckTest(string path)
    {
#if NETCOREAPP3_1_OR_GREATER
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
#endif
        using var fs = File.OpenRead(path);
        using var reader = ExcelReaderFactory.CreateReader(fs);
        var exceldatareaderResult = reader.AsDataSet();

        using var stream = File.OpenRead(path);
        var rows =  _excelImporter.Query(stream).ToList();
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
    public void QueryCustomStyle()
    {
        var path = PathHelper.GetFile("xlsx/TestWihoutRAttribute.xlsx");
        using (var stream = File.OpenRead(path))
        {
            // TODO: does this need filling? 
        }
    }

    [Fact]
    public void QuerySheetWithoutRAttribute()
    {
        var path = PathHelper.GetFile("xlsx/TestWihoutRAttribute.xlsx");
        using var stream = File.OpenRead(path);
        var rows =  _excelImporter.Query(stream).ToList();
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
        var rows =  _excelImporter.Query(stream).ToList();
        var keys = ((IDictionary<string, object>)rows.First()).Keys;
        Assert.Equal(3, keys.Count);
        Assert.Equal(2, rows.Count);
    }

    private class SaveAsFileWithDimensionByICollectionTestType
    {
        public string A { get; set; }
        public string B { get; set; }
    }
    [Fact]
    public void SaveAsFileWithDimensionByICollection()
    {
        //List<strongtype>
        {
            using var file = AutoDeletingPath.Create();
            var path = file.ToString();

            List<SaveAsFileWithDimensionByICollectionTestType> values =
            [
                new() { A = "A", B = "B" },
                new() { A = "A", B = "B" }
            ];
             _excelExporter.Export(path, values);

            using (var stream = File.OpenRead(path))
            {
                var rows =  _excelImporter.Query(stream, useHeaderRow: false).ToList();
                Assert.Equal(3, rows.Count);
                Assert.Equal("A", rows[0].A);
                Assert.Equal("A", rows[1].A);
                Assert.Equal("A", rows[2].A);
            }
            using (var stream = File.OpenRead(path))
            {
                var rows =  _excelImporter.Query(stream, useHeaderRow: true).ToList();
                Assert.Equal(2, rows.Count);
                Assert.Equal("A", rows[0].A);
                Assert.Equal("A", rows[1].A);
            }
            Assert.Equal("A1:B3", SheetHelper.GetFirstSheetDimensionRefValue(path));

             _excelExporter.Export(path, values, false, overwriteFile: true);
            Assert.Equal("A1:B2", SheetHelper.GetFirstSheetDimensionRefValue(path));
        }

        //List<strongtype> empty
        {
            using var file = AutoDeletingPath.Create();
            var path = file.ToString();
            List<SaveAsFileWithDimensionByICollectionTestType> values = [];

             _excelExporter.Export(path, values, false);
            {
                using (var stream = File.OpenRead(path))
                {
                    var rows =  _excelImporter.Query(stream, useHeaderRow: false).ToList();
                    Assert.Empty(rows);
                }
                Assert.Equal("A1:B1", SheetHelper.GetFirstSheetDimensionRefValue(path));
            }

            _excelExporter.Export(path, values, overwriteFile: true);
            {
                using var stream = File.OpenRead(path);
                var rows =  _excelImporter.Query(stream, useHeaderRow: false).ToList();
                Assert.Single(rows);
            }
            Assert.Equal("A1:B1", SheetHelper.GetFirstSheetDimensionRefValue(path));
        }

        //Array<anoymous>
        {
            using var file = AutoDeletingPath.Create();
            var path = file.ToString();
            var values = new[]
            {
                new {A="A",B="B"},
                new {A="A",B="B"},
            };
             _excelExporter.Export(path, values);
            {
                using (var stream = File.OpenRead(path))
                {
                    var rows =  _excelImporter.Query(stream, useHeaderRow: false).ToList();
                    Assert.Equal(3, rows.Count);
                    Assert.Equal("A", rows[0].A);
                    Assert.Equal("A", rows[1].A);
                    Assert.Equal("A", rows[2].A);
                }
                using (var stream = File.OpenRead(path))
                {
                    var rows =  _excelImporter.Query(stream, useHeaderRow: true).ToList();
                    Assert.Equal(2, rows.Count);
                    Assert.Equal("A", rows[0].A);
                    Assert.Equal("A", rows[1].A);
                }
            }
            Assert.Equal("A1:B3", SheetHelper.GetFirstSheetDimensionRefValue(path));

             _excelExporter.Export(path, values, false, overwriteFile: true);
            Assert.Equal("A1:B2", SheetHelper.GetFirstSheetDimensionRefValue(path));
        }

        // without properties
        {
            using var path = AutoDeletingPath.Create();
            var values = new List<int>();
            Assert.Throws<NotSupportedException>(() =>  _excelExporter.Export(path.ToString(), values));
        }
    }

    [Fact]
    public void SaveAsFileWithDimension()
    {
        {
            using var file = AutoDeletingPath.Create();
            var path = file.ToString();

            var table = new DataTable();
             _excelExporter.Export(path, table);
            Assert.Equal("A1", SheetHelper.GetFirstSheetDimensionRefValue(path));
            {
                using var stream = File.OpenRead(path);
                var rows =  _excelImporter.Query(stream).ToList();
                Assert.Single(rows);
            }

             _excelExporter.Export(path, table, printHeader: false, overwriteFile: true);
            Assert.Equal("A1", SheetHelper.GetFirstSheetDimensionRefValue(path));
        }

        {
            using var file = AutoDeletingPath.Create();
            var path = file.ToString();

            var table = new DataTable();
            table.Columns.Add("a", typeof(string));
            table.Columns.Add("b", typeof(decimal));
            table.Columns.Add("c", typeof(bool));
            table.Columns.Add("d", typeof(DateTime));
            table.Rows.Add(@"""<>+-*//}{\\n", 1234567890);
            table.Rows.Add("<test>Hello World</test>", -1234567890, false, DateTime.Now);

             _excelExporter.Export(path, table);
            Assert.Equal("A1:D3", SheetHelper.GetFirstSheetDimensionRefValue(path));

            using (var stream = File.OpenRead(path))
            {
                var rows =  _excelImporter.Query(stream, useHeaderRow: true).ToList();
                Assert.Equal(2, rows.Count);
                Assert.Equal(@"""<>+-*//}{\\n", rows[0].a);
                Assert.Equal(1234567890, rows[0].b);
                Assert.Null(rows[0].c);
                Assert.Null(rows[0].d);
            }

            using (var stream = File.OpenRead(path))
            {
                var rows =  _excelImporter.Query(stream).ToList();
                Assert.Equal(3, rows.Count);
                Assert.Equal("a", rows[0].A);
                Assert.Equal("b", rows[0].B);
                Assert.Equal("c", rows[0].C);
                Assert.Equal("d", rows[0].D);
            }

            _excelExporter.Export(path, table, printHeader: false, overwriteFile: true);
            Assert.Equal("A1:D2", SheetHelper.GetFirstSheetDimensionRefValue(path));
        }

        //TODO:StartCell
        {
            using var path = AutoDeletingPath.Create();

            var table = new DataTable();
            table.Columns.Add("a", typeof(string));
            table.Rows.Add("A");
            table.Rows.Add("B");

             _excelExporter.Export(path.ToString(), table);
            Assert.Equal("A1:A3", SheetHelper.GetFirstSheetDimensionRefValue(path.ToString()));
        }
    }

    [Fact]
    public void SaveAsByDataTableTest()
    {
        {
            var now = DateTime.Now;
            using var file = AutoDeletingPath.Create();
            var path = file.ToString();

            var table = new DataTable();
            table.Columns.Add("a", typeof(string));
            table.Columns.Add("b", typeof(decimal));
            table.Columns.Add("c", typeof(bool));
            table.Columns.Add("d", typeof(DateTime));
            table.Rows.Add(@"""<>+-*//}{\\n", 1234567890, true, now);
            table.Rows.Add("<test>Hello World</test>", -1234567890, false, now.Date);

             _excelExporter.Export(path, table, sheetName: "R&D");

            using var p = new ExcelPackage(new FileInfo(path));
            var ws = p.Workbook.Worksheets.First();

            Assert.True(ws.Cells["A1"].Value.ToString() == "a");
            Assert.True(ws.Cells["B1"].Value.ToString() == "b");
            Assert.True(ws.Cells["C1"].Value.ToString() == "c");
            Assert.True(ws.Cells["D1"].Value.ToString() == "d");

            Assert.True(ws.Cells["A2"].Value.ToString() == @"""<>+-*//}{\\n");
            Assert.True(ws.Cells["B2"].Value.ToString() == "1234567890");
            Assert.True(ws.Cells["C2"].Value.ToString() == true.ToString());
            Assert.True(ws.Cells["D2"].Value.ToString() == now.ToString());

            Assert.True(ws.Name == "R&D");
        }
        {
            using var path = AutoDeletingPath.Create();
            var table = new DataTable();
            table.Columns.Add("Column1", typeof(string));
            table.Columns.Add("Column2", typeof(int));
            table.Rows.Add("MiniExcel", 1);
            table.Rows.Add("Github", 2);

             _excelExporter.Export(path.ToString(), table);
        }
    }

    [Fact]
    public void QueryByLINQExtensionsAvoidLargeFileOOMTest()
    {
        const string path = "../../../../../benchmarks/MiniExcel.Benchmarks/Test1,000,000x10.xlsx";

        var query1 =  _excelImporter.Query(path).First();
        Assert.Equal("HelloWorld1", query1.A);

        using (var stream = File.OpenRead(path))
        {
            var query2 =  _excelImporter.Query(stream).First();
            Assert.Equal("HelloWorld1", query2.A);
        }

        var query3 =  _excelImporter.Query(path).Take(10);
        Assert.Equal(10, query3.Count());
    }

    [Fact]
    public void EmptyTest()
    {
        using var path = AutoDeletingPath.Create();
        using (var connection = Db.GetConnection("Data Source=:memory:"))
        {
            var rows = connection.Query("with cte as (select 1 id,2 val) select * from cte where 1=2");
             _excelExporter.Export(path.ToString(), rows);
        }
        using (var stream = File.OpenRead(path.ToString()))
        {
            var rows =  _excelImporter.Query(stream, useHeaderRow: true).ToList();
            Assert.Empty(rows);
        }
    }

    [Fact]
    public void SaveAsByIEnumerableIDictionary()
    {
        using var file = AutoDeletingPath.Create();
        var path = file.ToString();

        {
            var values = new List<Dictionary<string, object>>
            {
                new() { { "Column1", "MiniExcel" }, { "Column2", 1 } },
                new() { { "Column1", "Github" }, { "Column2", 2 } }
            };
            var sheets = new Dictionary<string, object>
            {
                ["R&D"] = values,
                ["success!"] = values
            };
             _excelExporter.Export(path, sheets);

            using (var stream = File.OpenRead(path))
            {
                var rows =  _excelImporter.Query(stream, useHeaderRow: false).ToList();
                Assert.Equal("Column1", rows[0].A);
                Assert.Equal("Column2", rows[0].B);
                Assert.Equal("MiniExcel", rows[1].A);
                Assert.Equal(1, rows[1].B);
                Assert.Equal("Github", rows[2].A);
                Assert.Equal(2, rows[2].B);

                Assert.Equal("R&D",  _excelImporter.GetSheetNames(stream)[0]);
            }

            using (var stream = File.OpenRead(path))
            {
                var rows =  _excelImporter.Query(stream, useHeaderRow: true).ToList();

                Assert.Equal(2, rows.Count);
                Assert.Equal("MiniExcel", rows[0].Column1);
                Assert.Equal(1, rows[0].Column2);
                Assert.Equal("Github", rows[1].Column1);
                Assert.Equal(2, rows[1].Column2);

                Assert.Equal("success!",  _excelImporter.GetSheetNames(stream)[1]);
            }

            Assert.Equal("A1:B3", SheetHelper.GetFirstSheetDimensionRefValue(path));
        }

        {
            var values = new List<Dictionary<int, object>>
            {
                new() { { 1, "MiniExcel"}, { 2, 1 } },
                new() { { 1, "Github" }, { 2, 2 } },
            };
             _excelExporter.Export(path, values, overwriteFile: true);

            using (var stream = File.OpenRead(path))
            {
                var rows =  _excelImporter.Query(stream, useHeaderRow: false).ToList();
                Assert.Equal(3, rows.Count);
            }

            Assert.Equal("A1:B3", SheetHelper.GetFirstSheetDimensionRefValue(path));
        }
    }

    [Fact]
    public void SaveAsFrozenRowsAndColumnsTest()
    {
        var config = new OpenXmlConfiguration
        {
            FreezeRowCount = 1,
            FreezeColumnCount = 2
        };

        // Test enumerable
        using var path = AutoDeletingPath.Create();
         _excelExporter.Export(
            path.ToString(),
            new[]
            {
                new { Column1 = "MiniExcel", Column2 = 1 },
                new { Column1 = "Github", Column2 = 2 }
            },
            configuration: config
        );

        using (var stream = File.OpenRead(path.ToString()))
        {
            var rows =  _excelImporter.Query(stream, useHeaderRow: true).ToList();

            Assert.Equal("MiniExcel", rows[0].Column1);
            Assert.Equal(1, rows[0].Column2);
            Assert.Equal("Github", rows[1].Column1);
            Assert.Equal(2, rows[1].Column2);
        }

        Assert.Equal("A1:B3", SheetHelper.GetFirstSheetDimensionRefValue(path.ToString()));

        // test table
        var table = new DataTable();
        table.Columns.Add("a", typeof(string));
        table.Columns.Add("b", typeof(decimal));
        table.Columns.Add("c", typeof(bool));
        table.Columns.Add("d", typeof(DateTime));
        table.Rows.Add("some text", 1234567890, true, DateTime.Now);
        table.Rows.Add("<test>Hello World</test>", -1234567890, false, DateTime.Now.Date);

        using var pathTable = AutoDeletingPath.Create();
         _excelExporter.Export(pathTable.ToString(), table, configuration: config);
        Assert.Equal("A1:D3", SheetHelper.GetFirstSheetDimensionRefValue(pathTable.ToString()));

        // data reader
        var reader = table.CreateDataReader();
        using var pathReader = AutoDeletingPath.Create();

         _excelExporter.Export(pathReader.ToString(), reader, configuration: config, overwriteFile: true);
        Assert.Equal("A1:D3", SheetHelper.GetFirstSheetDimensionRefValue(pathTable.ToString())); //TODO: fix datareader not writing ref dimension (also in async version)
    }

    [Fact]
    public void SaveAsByDapperRows()
    {
        using var file = AutoDeletingPath.Create();
        var path = file.ToString();

        // Dapper Query
        using (var connection = Db.GetConnection("Data Source=:memory:"))
        {
            var rows = connection.Query("select 'MiniExcel' as Column1,1 as Column2 union all select 'Github',2");
             _excelExporter.Export(path, rows);
        }

        Assert.Equal("A1:B3", SheetHelper.GetFirstSheetDimensionRefValue(path));

        using (var stream = File.OpenRead(path))
        {
            var rows =  _excelImporter.Query(stream, useHeaderRow: true).ToList();

            Assert.Equal("MiniExcel", rows[0].Column1);
            Assert.Equal(1, rows[0].Column2);
            Assert.Equal("Github", rows[1].Column1);
            Assert.Equal(2, rows[1].Column2);
        }

        // Empty
        using (var connection = Db.GetConnection("Data Source=:memory:"))
        {
            var rows = connection.Query("with cte as (select 'MiniExcel' as Column1,1 as Column2 union all select 'Github',2)select * from cte where 1=2").ToList();
             _excelExporter.Export(path, rows, overwriteFile: true);
        }

        using (var stream = File.OpenRead(path))
        {
            var rows =  _excelImporter.Query(stream, useHeaderRow: false).ToList();
            Assert.Empty(rows);
        }

        using (var stream = File.OpenRead(path))
        {
            var rows =  _excelImporter.Query(stream, useHeaderRow: true).ToList();
            Assert.Empty(rows);
        }

        Assert.Equal("A1", SheetHelper.GetFirstSheetDimensionRefValue(path));

        // ToList
        using (var connection = Db.GetConnection("Data Source=:memory:"))
        {
            var rows = connection.Query("select 'MiniExcel' as Column1,1 as Column2 union all select 'Github',2").ToList();
             _excelExporter.Export(path, rows, overwriteFile: true);
        }

        Assert.Equal("A1:B3", SheetHelper.GetFirstSheetDimensionRefValue(path));

        using (var stream = File.OpenRead(path))
        {
            var rows =  _excelImporter.Query(stream, useHeaderRow: false).ToList();

            Assert.Equal("Column1", rows[0].A);
            Assert.Equal("Column2", rows[0].B);
            Assert.Equal("MiniExcel", rows[1].A);
            Assert.Equal(1, rows[1].B);
            Assert.Equal("Github", rows[2].A);
            Assert.Equal(2, rows[2].B);
        }

        using (var stream = File.OpenRead(path))
        {
            var rows =  _excelImporter.Query(stream, useHeaderRow: true).ToList();

            Assert.Equal("MiniExcel", rows[0].Column1);
            Assert.Equal(1, rows[0].Column2);
            Assert.Equal("Github", rows[1].Column1);
            Assert.Equal(2, rows[1].Column2);
        }
    }

    private class Demo
    {
        public string Column1 { get; set; }
        public decimal Column2 { get; set; }
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
        var rows =  _excelImporter.Query(stream, useHeaderRow: true).ToList();

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
        var rows =  _excelImporter.Query(stream, useHeaderRow: true).ToList();

        Assert.Equal("MiniExcel", rows[0].Column1);
        Assert.Equal(1, rows[0].Column2);
        Assert.Equal("Github", rows[1].Column1);
        Assert.Equal(2, rows[1].Column2);
    }

    [Fact]
    public void SQLiteInsertTest()
    {
        // Avoid SQL Insert Large Size Xlsx OOM
        var path = PathHelper.GetFile("xlsx/Test5x2.xlsx");
        var tempSqlitePath = AutoDeletingPath.Create(Path.GetTempPath(), $"{Guid.NewGuid()}.db");
        var connectionString = $"Data Source={tempSqlitePath};Version=3;";

        using (var connection = new SQLiteConnection(connectionString))
        {
            connection.Execute("create table T (A varchar(20),B varchar(20));");
        }

        using (var connection = new SQLiteConnection(connectionString))
        {
            connection.Open();
            using (var transaction = connection.BeginTransaction())
            using (var stream = File.OpenRead(path))
            {
                var rows =  _excelImporter.Query(stream);
                foreach (var row in rows)
                {
                    _ = connection.Execute("insert into T (A,B) values (@A,@B)", new { row.A, row.B }, transaction: transaction);
                }

                transaction.Commit();
            }
        }

        using (var connection = new SQLiteConnection(connectionString))
        {
            var result = connection.Query("select * from T");
            Assert.Equal(5, result.Count());
        }
    }

    [Fact]
    public void SaveAsBasicCreateTest()
    {
        using var path = AutoDeletingPath.Create();

        var rowsWritten =  _excelExporter.Export(path.ToString(), new[]
        {
            new { Column1 = "MiniExcel", Column2 = 1 },
            new { Column1 = "Github", Column2 = 2}
        });

        Assert.Single(rowsWritten);
        Assert.Equal(2, rowsWritten[0]);

        using (var stream = File.OpenRead(path.ToString()))
        {
            var rows =  _excelImporter.Query(stream, useHeaderRow: true).ToList();

            Assert.Equal("MiniExcel", rows[0].Column1);
            Assert.Equal(1, rows[0].Column2);
            Assert.Equal("Github", rows[1].Column1);
            Assert.Equal(2, rows[1].Column2);
        }

        Assert.Equal("A1:B3", SheetHelper.GetFirstSheetDimensionRefValue(path.ToString()));
    }

    [Fact]
    public void SaveAsBasicStreamTest()
    {
        {
            using var path = AutoDeletingPath.Create();
            var values = new[]
            {
                new { Column1 = "MiniExcel", Column2 = 1 },
                new { Column1 = "Github", Column2 = 2 }
            };
            using (var stream = new FileStream(path.ToString(), FileMode.CreateNew))
            {
                var rowsWritten =  _excelExporter.Export(stream, values);
                Assert.Single(rowsWritten);
                Assert.Equal(2, rowsWritten[0]);
            }

            using (var stream = File.OpenRead(path.ToString()))
            {
                var rows =  _excelImporter.Query(stream, useHeaderRow: true).ToList();

                Assert.Equal("MiniExcel", rows[0].Column1);
                Assert.Equal(1, rows[0].Column2);
                Assert.Equal("Github", rows[1].Column1);
                Assert.Equal(2, rows[1].Column2);
            }
        }
        {
            using var path = AutoDeletingPath.Create();
            var values = new[]
            {
                new { Column1 = "MiniExcel", Column2 = 1 },
                new { Column1 = "Github", Column2 = 2}
            };
            using (var stream = new MemoryStream())
            using (var fileStream = new FileStream(path.ToString(), FileMode.Create))
            {
                var rowsWritten =  _excelExporter.Export(stream, values);
                stream.Seek(0, SeekOrigin.Begin);
                stream.CopyTo(fileStream);
                Assert.Single(rowsWritten);
                Assert.Equal(2, rowsWritten[0]);
            }

            using (var stream = File.OpenRead(path.ToString()))
            {
                var rows =  _excelImporter.Query(stream, useHeaderRow: true).ToList();

                Assert.Equal("MiniExcel", rows[0].Column1);
                Assert.Equal(1, rows[0].Column2);
                Assert.Equal("Github", rows[1].Column1);
                Assert.Equal(2, rows[1].Column2);
            }
        }
    }

    [Fact]
    public void SaveAsSpecialAndTypeCreateTest()
    {
        using var path = AutoDeletingPath.Create();
        var rowsWritten =  _excelExporter.Export(path.ToString(), new[]
        {
            new { a = @"""<>+-*//}{\\n", b = 1234567890, c = true, d = DateTime.Now },
            new { a = "<test>Hello World</test>", b = -1234567890, c = false, d = DateTime.Now.Date }
        });
        Assert.Single(rowsWritten);
        Assert.Equal(2, rowsWritten[0]);

        var info = new FileInfo(path.ToString());
        Assert.True(info.FullName == path.ToString());
    }

    [Fact]
    public void SaveAsFileEpplusCanReadTest()
    {
        var now = DateTime.Now;
        using var path = AutoDeletingPath.Create();
        var rowsWritten =  _excelExporter.Export(path.ToString(), new[]
        {
            new { a = @"""<>+-*//}{\\n", b = 1234567890, c = true, d = now},
            new { a = "<test>Hello World</test>", b = -1234567890, c = false, d = now.Date }
        });
        Assert.Single(rowsWritten);
        Assert.Equal(2, rowsWritten[0]);

        using var p = new ExcelPackage(new FileInfo(path.ToString()));
        var ws = p.Workbook.Worksheets.First();

        Assert.True(ws.Cells["A1"].Value.ToString() == "a");
        Assert.True(ws.Cells["B1"].Value.ToString() == "b");
        Assert.True(ws.Cells["C1"].Value.ToString() == "c");
        Assert.True(ws.Cells["D1"].Value.ToString() == "d");

        Assert.True(ws.Cells["A2"].Value.ToString() == @"""<>+-*//}{\\n");
        Assert.True(ws.Cells["B2"].Value.ToString() == "1234567890");
        Assert.True(ws.Cells["C2"].Value.ToString() == true.ToString());
        Assert.True(ws.Cells["D2"].Value.ToString() == now.ToString());
    }

    [Fact]
    public void SavaAsClosedXmlCanReadTest()
    {
        var now = DateTime.Now;
        using var path = AutoDeletingPath.Create();
        var rowsWritten =  _excelExporter.Export(path.ToString(), new[]
        {
            new { a = @"""<>+-*//}{\\n", b = 1234567890, c = true, d = now },
            new { a = "<test>Hello World</test>", b = -1234567890, c = false, d = now.Date }
        }, sheetName: "R&D");

        Assert.Single(rowsWritten);
        Assert.Equal(2, rowsWritten[0]);

        using var workbook = new XLWorkbook(path.ToString());
        var ws = workbook.Worksheets.First();

        Assert.True(ws.Cell("A1").Value.ToString() == "a");
        Assert.True(ws.Cell("D1").Value.ToString() == "d");
        Assert.True(ws.Cell("B1").Value.ToString() == "b");
        Assert.True(ws.Cell("C1").Value.ToString() == "c");

        Assert.True(ws.Cell("A2").Value.ToString() == @"""<>+-*//}{\\n");
        Assert.True(ws.Cell("B2").Value.ToString() == "1234567890");
        Assert.Equal(bool.TrueString, ws.Cell("C2").Value.ToString(), ignoreCase: true);
        Assert.True(ws.Cell("D2").Value.ToString() == now.ToString());

        Assert.True(ws.Name == "R&D");
    }

    [Fact]
    public void ContentTypeUriContentTypeReadCheckTest()
    {
        var now = DateTime.Now;
        using var path = AutoDeletingPath.Create();
        var rowsWritten =  _excelExporter.Export(path.ToString(), new[]
        {
            new { a = @"""<>+-*//}{\\n", b = 1234567890, c = true, d= now },
            new { a = "<test>Hello World</test>", b = -1234567890, c = false, d = now.Date }
        });
        Assert.Single(rowsWritten);
        Assert.Equal(2, rowsWritten[0]);

        using var zip = Package.Open(path.ToString(), FileMode.Open);
        var allParts = zip.GetParts()
            .Select(s => new { s.CompressionOption, s.ContentType, s.Uri, s.Package.GetType().Name })
            .ToDictionary(s => s.Uri.ToString(), s => s);

        Assert.True(allParts["/xl/styles.xml"].ContentType == "application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml");
        Assert.True(allParts["/xl/workbook.xml"].ContentType == "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml");
        Assert.True(allParts["/xl/worksheets/sheet1.xml"].ContentType == "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml");
        Assert.True(allParts["/xl/_rels/workbook.xml.rels"].ContentType == "application/vnd.openxmlformats-package.relationships+xml");
        Assert.True(allParts["/_rels/.rels"].ContentType == "application/vnd.openxmlformats-package.relationships+xml");
    }

    [Fact]
    public void TestStirctOpenXml()
    {
        var path = PathHelper.GetFile("xlsx/TestStrictOpenXml.xlsx");
        var columns =  _excelImporter.GetColumnNames (path);
        Assert.Equal(["A", "B", "C"], columns);

        var rows =  _excelImporter.Query(path).ToList();
        Assert.Equal(rows[0].A, "title1");
        Assert.Equal(rows[0].B, "title2");
        Assert.Equal(rows[0].C, "title3");
        Assert.Equal(rows[1].A, "value1");
        Assert.Equal(rows[1].B, "value2");
        Assert.Equal(rows[1].C, "value3");
    }

    [Fact]
    public void SharedStringCacheTest()
    {
        const string path = "../../../../../benchmarks/MiniExcel.Benchmarks/Test1,000,000x10_SharingStrings.xlsx";

        var ts = Stopwatch.GetTimestamp();
        _ =  _excelImporter.Query(path, configuration: new OpenXmlConfiguration { EnableSharedStringCache = true }).First();
        using var currentProcess = Process.GetCurrentProcess();
        var totalBytesOfMemoryUsed = currentProcess.WorkingSet64;

        _output.WriteLine("totalBytesOfMemoryUsed: " + totalBytesOfMemoryUsed);
        _output.WriteLine("elapsedMilliseconds: " + Stopwatch.GetElapsedTime(ts).TotalMilliseconds);
    }

    [Fact]
    public void SharedStringNoCacheTest()
    {
        const string path = "../../../../../benchmarks/MiniExcel.Benchmarks/Test1,000,000x10_SharingStrings.xlsx";

        var ts = Stopwatch.GetTimestamp();
        _ =  _excelImporter.Query(path).First();
        using var currentProcess = Process.GetCurrentProcess();
        var totalBytesOfMemoryUsed = currentProcess.WorkingSet64;
        _output.WriteLine("totalBytesOfMemoryUsed: " + totalBytesOfMemoryUsed);
        _output.WriteLine("elapsedMilliseconds: " + Stopwatch.GetElapsedTime(ts).TotalMilliseconds);
    }

    [Fact]
    public void DynamicColumnsConfigurationIsUsedWhenCreatingExcelUsingIDataReader()
    {
        using var path = AutoDeletingPath.Create();
        var dateTime = DateTime.Now;
        var onlyDate = DateOnly.FromDateTime(dateTime);

        var table = new DataTable();
        table.Columns.Add("Column1", typeof(string));
        table.Columns.Add("Column2", typeof(int));
        table.Columns.Add("Column3", typeof(DateTime));
        table.Columns.Add("Column4", typeof(DateOnly));
        table.Rows.Add("MiniExcel", 1, dateTime, onlyDate);
        table.Rows.Add("Github", 2, dateTime, onlyDate);

        var configuration = new OpenXmlConfiguration
        {
            DynamicColumns =
            [
                new DynamicExcelColumn("Column1")
                {
                    Name = "Name of something",
                    Index = 0,
                    Width = 150
                },
                new DynamicExcelColumn("Column2")
                {
                    Name = "Its value",
                    Index = 1,
                    Width = 150
                },
                new DynamicExcelColumn("Column3")
                {
                    Name = "Its Date",
                    Index = 2,
                    Width = 150,
                    Format = "dd.mm.yyyy hh:mm:ss",
                }
            ]
        };
        var reader = table.CreateDataReader();
         _excelExporter.Export(path.ToString(), reader, configuration: configuration);

        using var stream = File.OpenRead(path.ToString());
        var rows =  _excelImporter.Query(stream, useHeaderRow: true)
            .Select(x => (IDictionary<string, object>)x)
            .ToList();

        Assert.Contains("Name of something", rows[0]);
        Assert.Contains("Its value", rows[0]);
        Assert.Contains("Its Date", rows[0]);
        Assert.Contains("Column4", rows[0]);
        Assert.Contains("Name of something", rows[1]);
        Assert.Contains("Its value", rows[1]);
        Assert.Contains("Its Date", rows[1]);
        Assert.Contains("Column4", rows[1]);

        Assert.Equal("MiniExcel", rows[0]["Name of something"]);
        Assert.Equal(1D, rows[0]["Its value"]);
        Assert.Equal(dateTime, (DateTime)rows[0]["Its Date"], TimeSpan.FromMilliseconds(10d));
        Assert.Equal(onlyDate.ToDateTime(TimeOnly.MinValue), (DateTime)rows[0]["Column4"]);
        Assert.Equal("Github", rows[1]["Name of something"]);
        Assert.Equal(2D, rows[1]["Its value"]);
        Assert.Equal(dateTime, (DateTime)rows[1]["Its Date"], TimeSpan.FromMilliseconds(10d));
        Assert.Equal(onlyDate.ToDateTime(TimeOnly.MinValue), (DateTime)rows[1]["Column4"]);
    }

    [Fact]
    public void DynamicColumnsConfigurationIsUsedWhenCreatingExcelUsingDataTable()
    {
        using var path = AutoDeletingPath.Create();
        var dateTime = DateTime.Now;
        var onlyDate = DateOnly.FromDateTime(dateTime);

        var table = new DataTable();
        table.Columns.Add("Column1", typeof(string));
        table.Columns.Add("Column2", typeof(int));
        table.Columns.Add("Column3", typeof(DateTime));
        table.Columns.Add("Column4", typeof(DateOnly));
        table.Rows.Add("MiniExcel", 1, dateTime, onlyDate);
        table.Rows.Add("Github", 2, dateTime, onlyDate);

        var configuration = new OpenXmlConfiguration
        {
            DynamicColumns =
            [
                new DynamicExcelColumn("Column1")
                {
                    Name = "Name of something",
                    Index = 0,
                    Width = 150
                },
                new DynamicExcelColumn("Column2")
                {
                    Name = "Its value",
                    Index = 1,
                    Width = 150
                },
                new DynamicExcelColumn("Column3")
                {
                    Name = "Its Date",
                    Index = 2,
                    Width = 150,
                    Format = "dd.mm.yyyy hh:mm:ss"
                }
            ]
        };
         _excelExporter.Export(path.ToString(), table, configuration: configuration);

        using var stream = File.OpenRead(path.ToString());
        var rows =  _excelImporter.Query(stream, useHeaderRow: true)
            .Select(x => (IDictionary<string, object>)x)
            .Select(x => (IDictionary<string, object>)x)
            .ToList();

        Assert.Contains("Name of something", rows[0]);
        Assert.Contains("Its value", rows[0]);
        Assert.Contains("Its Date", rows[0]);
        Assert.Contains("Column4", rows[0]);
        Assert.Contains("Name of something", rows[1]);
        Assert.Contains("Its value", rows[1]);
        Assert.Contains("Its Date", rows[1]);
        Assert.Contains("Column4", rows[1]);


        Assert.Equal("MiniExcel", rows[0]["Name of something"]);
        Assert.Equal(1D, rows[0]["Its value"]);
        Assert.Equal(dateTime, (DateTime)rows[0]["Its Date"], TimeSpan.FromMilliseconds(10d));
        Assert.Equal(onlyDate.ToDateTime(TimeOnly.MinValue), (DateTime)rows[0]["Column4"]);
        Assert.Equal("Github", rows[1]["Name of something"]);
        Assert.Equal(2D, rows[1]["Its value"]);
        Assert.Equal(dateTime, (DateTime)rows[1]["Its Date"], TimeSpan.FromMilliseconds(10d));
        Assert.Equal(onlyDate.ToDateTime(TimeOnly.MinValue), (DateTime)rows[1]["Column4"]);
    }

    [Fact]
    public void InsertSheetTest()
    {
        var now = DateTime.Now;
        using var file = AutoDeletingPath.Create();
        var path = file.ToString();

        {
            var table = new DataTable();
            table.Columns.Add("a", typeof(string));
            table.Columns.Add("b", typeof(decimal));
            table.Columns.Add("c", typeof(bool));
            table.Columns.Add("d", typeof(DateTime));
            table.Rows.Add(@"""<>+-*//}{\\n", 1234567890, true, now);
            table.Rows.Add("<test>Hello World</test>", -1234567890, false, now.Date);

            var rowsWritten =  _excelExporter.InsertSheet(path, table, sheetName: "Sheet1");
            Assert.Equal(2, rowsWritten);

            using var p = new ExcelPackage(new FileInfo(path));
            var sheet1 = p.Workbook.Worksheets[0];

            Assert.True(sheet1.Cells["A1"].Value.ToString() == "a");
            Assert.True(sheet1.Cells["B1"].Value.ToString() == "b");
            Assert.True(sheet1.Cells["C1"].Value.ToString() == "c");
            Assert.True(sheet1.Cells["D1"].Value.ToString() == "d");

            Assert.True(sheet1.Cells["A2"].Value.ToString() == @"""<>+-*//}{\\n");
            Assert.True(sheet1.Cells["B2"].Value.ToString() == "1234567890");
            Assert.True(sheet1.Cells["C2"].Value.ToString() == true.ToString());
            Assert.True(sheet1.Cells["D2"].Value.ToString() == now.ToString());

            Assert.True(sheet1.Name == "Sheet1");
        }
        {
            var table = new DataTable();
            table.Columns.Add("Column1", typeof(string));
            table.Columns.Add("Column2", typeof(int));
            table.Rows.Add("MiniExcel", 1);
            table.Rows.Add("Github", 2);

            var rowsWritten =  _excelExporter.InsertSheet(path, table, sheetName: "Sheet2");
            Assert.Equal(2, rowsWritten);

            using var p = new ExcelPackage(new FileInfo(path));
            var sheet2 = p.Workbook.Worksheets[1];

            Assert.True(sheet2.Cells["A1"].Value.ToString() == "Column1");
            Assert.True(sheet2.Cells["B1"].Value.ToString() == "Column2");

            Assert.True(sheet2.Cells["A2"].Value.ToString() == "MiniExcel");
            Assert.True(sheet2.Cells["B2"].Value.ToString() == "1");

            Assert.True(sheet2.Cells["A3"].Value.ToString() == "Github");
            Assert.True(sheet2.Cells["B3"].Value.ToString() == "2");

            Assert.True(sheet2.Name == "Sheet2");
        }
        {
            var table = new DataTable();
            table.Columns.Add("Column1", typeof(string));
            table.Columns.Add("Column2", typeof(DateTime));
            table.Rows.Add("Test", now);

            var rowsWritten =  _excelExporter.InsertSheet(path, table, sheetName: "Sheet2", printHeader: false, configuration: new OpenXmlConfiguration
            {
                FastMode = true,
                AutoFilter = false,
                TableStyles = TableStyles.None,
                DynamicColumns =
                [
                    new DynamicExcelColumn("Column2")
                    {
                        Name = "Its Date",
                        Index = 1,
                        Width = 150,
                        Format = "dd.mm.yyyy hh:mm:ss",
                    }
                ]
            }, overwriteSheet: true);

            Assert.Equal(1, rowsWritten);

            using var p = new ExcelPackage(new FileInfo(path));
            var sheet2 = p.Workbook.Worksheets[1];

            Assert.True(sheet2.Cells["A1"].Value.ToString() == "Test");
            Assert.True(sheet2.Cells["B1"].Text == now.ToString("dd.MM.yyyy HH:mm:ss"));
            Assert.True(sheet2.Name == "Sheet2");
        }
        {
            var table = new DataTable();
            table.Columns.Add("Column1", typeof(string));
            table.Columns.Add("Column2", typeof(DateTime));
            table.Rows.Add("MiniExcel", now);
            table.Rows.Add("Github", now);

            var rowsWritten =  _excelExporter.InsertSheet(path, table, sheetName: "Sheet3", configuration: new OpenXmlConfiguration
            {
                FastMode = true,
                AutoFilter = false,
                TableStyles = TableStyles.None,
                DynamicColumns =
                [
                    new DynamicExcelColumn("Column2")
                    {
                        Name = "Its Date",
                        Index = 1,
                        Width = 150,
                        Format = "dd.mm.yyyy hh:mm:ss",
                    }
                ]
            });
            Assert.Equal(2, rowsWritten);

            using var p = new ExcelPackage(new FileInfo(path));
            var sheet3 = p.Workbook.Worksheets[2];

            Assert.True(sheet3.Cells["A1"].Value.ToString() == "Column1");
            Assert.True(sheet3.Cells["B1"].Value.ToString() == "Its Date");

            Assert.True(sheet3.Cells["A2"].Value.ToString() == "MiniExcel");
            Assert.True(sheet3.Cells["B2"].Text == now.ToString("dd.MM.yyyy HH:mm:ss"));

            Assert.True(sheet3.Cells["A3"].Value.ToString() == "Github");
            Assert.True(sheet3.Cells["B3"].Text == now.ToString("dd.MM.yyyy HH:mm:ss"));

            Assert.True(sheet3.Name == "Sheet3");
        }
    }

    private class DateOnlyTest
    {
        public DateOnly Date { get; set; }
        [MiniExcelFormat("d.M.yyyy")] public DateOnly DateWithFormat { get; set; }
    }

    [Fact]
    public void DateOnlySupportTest()
    {
        var query =  _excelImporter.Query<DateOnlyTest>(PathHelper.GetFile("xlsx/TestDateOnlyMapping.xlsx")).ToList();

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
        var dim1 =  _excelImporter.GetSheetDimensions(path1);
        Assert.Equal("A1", dim1[0].StartCell);
        Assert.Equal("H101", dim1[0].EndCell);
        Assert.Equal(101, dim1[0].Rows.Count);
        Assert.Equal(8, dim1[0].Columns.Count);
        Assert.Equal(1, dim1[0].Rows.StartIndex);
        Assert.Equal(101, dim1[0].Rows.EndIndex);
        Assert.Equal(1, dim1[0].Columns.StartIndex);
        Assert.Equal(8, dim1[0].Columns.EndIndex);

        var path2 = PathHelper.GetFile("xlsx/TestNoDimension.xlsx");
        var dim2 =  _excelImporter.GetSheetDimensions(path2);
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
        var dim =  _excelImporter.GetSheetDimensions(path);

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
}