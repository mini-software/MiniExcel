using ClosedXML.Excel;
using Dapper;
using ExcelDataReader;
using MiniExcelLibs.Attributes;
using MiniExcelLibs.OpenXml;
using MiniExcelLibs.Tests.Utils;
using OfficeOpenXml;
using System.Data;
using System.Data.SQLite;
using System.Globalization;
using System.IO.Packaging;
using System.Text;
using Xunit;

namespace MiniExcelLibs.Tests;

public class MiniExcelOpenXmlAsyncTests
{
    [Fact]
    public async Task SaveAsControlChracter()
    {
        using var file = AutoDeletingPath.Create();
        var path = file.ToString();
        
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
        await MiniExcel.SaveAsAsync(path, input);

        var rows2 = (await MiniExcel.QueryAsync(path, true)).ToArray();
        var rows1 = (await MiniExcel.QueryAsync<SaveAsControlChracterVO>(path)).ToArray();
    }

    private class SaveAsControlChracterVO
    {
        public string Test { get; set; }
    }

    private class ExcelAttributeDemo
    {
        [ExcelColumnName("Column1")]
        public string Test1 { get; set; }
        [ExcelColumnName("Column2")]
        public string Test2 { get; set; }
        [ExcelIgnore]
        public string Test3 { get; set; }
        [ExcelColumnIndex("I")] // system will convert "I" to 8 index
        public string Test4 { get; set; }
        public string Test5 { get; } //wihout set will ignore
        public string Test6 { get; private set; } //un-public set will ignore
        [ExcelColumnIndex(3)] // start with 0
        public string Test7 { get; set; }
    }

    private class ExcelAttributeDemo2
    {
        [ExcelColumn(Name = "Column1")]
        public string Test1 { get; set; }
        [ExcelColumn(Name = "Column2")]
        public string Test2 { get; set; }
        [ExcelColumn(Ignore = true)]
        public string Test3 { get; set; }
        [ExcelColumn(IndexName = "I")] // system will convert "I" to 8 index
        public string Test4 { get; set; }
        public string Test5 { get; } //wihout set will ignore
        public string Test6 { get; private set; } //un-public set will ignore
        [ExcelColumn(Index = 3)] // start with 0
        public string Test7 { get; set; }
    }

    [Fact]
    public async Task CustomAttributeWihoutVaildPropertiesTest()
    {
        const string path = "../../../../../samples/xlsx/TestCustomExcelColumnAttribute.xlsx";
        await Assert.ThrowsAsync<InvalidOperationException>(async () =>
        {
            _ = (await MiniExcel.QueryAsync<CustomAttributesWihoutVaildPropertiesTestPoco>(path)).ToList();
        });
    }

    [Fact]
    public async Task QueryCustomAttributesTest()
    {
        const string path = "../../../../../samples/xlsx/TestCustomExcelColumnAttribute.xlsx";
        var rows = (await MiniExcel.QueryAsync<ExcelAttributeDemo>(path)).ToList();

        Assert.Equal("Column1", rows[0].Test1);
        Assert.Equal("Column2", rows[0].Test2);
        Assert.Null(rows[0].Test3);
        Assert.Equal("Test7", rows[0].Test4);
        Assert.Null(rows[0].Test5);
        Assert.Null(rows[0].Test6);
        Assert.Equal("Test4", rows[0].Test7);
    }

    [Fact]
    public async Task QueryCustomAttributes2Test()
    {
        const string path = "../../../../../samples/xlsx/TestCustomExcelColumnAttribute.xlsx";
        var rows = (await MiniExcel.QueryAsync<ExcelAttributeDemo2>(path)).ToList();
        
        Assert.Equal("Column1", rows[0].Test1);
        Assert.Equal("Column2", rows[0].Test2);
        Assert.Null(rows[0].Test3);
        Assert.Equal("Test7", rows[0].Test4);
        Assert.Null(rows[0].Test5);
        Assert.Null(rows[0].Test6);
        Assert.Equal("Test4", rows[0].Test7);
    }

    [Fact]
    public async Task SaveAsCustomAttributesTest()
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
        
        await MiniExcel.SaveAsAsync(path.ToString(), input);
        var d = await MiniExcel.QueryAsync(path.ToString(), true);
        var rows = d.ToList();
        var first = rows[0] as IDictionary<string, object>;
        
        Assert.Equal(3, rows.Count);
        Assert.Equal(["Column1", "Column2", "Test5", "Test7", "Test6", "Test4"], first?.Keys);
        Assert.Equal("Test1", rows[0].Column1);
        Assert.Equal("Test2", rows[0].Column2);
        Assert.Equal("Test4", rows[0].Test4);
        Assert.Null(rows[0].Test5);
        Assert.Null(rows[0].Test6);
    }

    [Fact]
    public async Task SaveAsCustomAttributes2Test()
    {
        using var path = AutoDeletingPath.Create();
        var input = Enumerable.Range(1, 3)
            .Select(_ => new ExcelAttributeDemo2
            {
                Test1 = "Test1",
                Test2 = "Test2",
                Test3 = "Test3",
                Test4 = "Test4",
            });
        
        await MiniExcel.SaveAsAsync(path.ToString(), input);
        var d = await MiniExcel.QueryAsync(path.ToString(), true);
        var rows = d.ToList();
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
        [ExcelIgnore]
        public string Test3 { get; set; }
        public string Test5 { get; }
        public string Test6 { get; private set; }
    }

    [Fact]
    public async Task QueryCastToIDictionary()
    {
        const string path = "../../../../../samples/xlsx/TestCenterEmptyRow/TestCenterEmptyRow.xlsx";
        foreach (IDictionary<string, object> row in await MiniExcel.QueryAsync(path))
        {
            _ = row;
        }
    }
    
    [Fact]
    public async Task CenterEmptyRowsQueryTest()
    {
        const string path = "../../../../../samples/xlsx/TestCenterEmptyRow/TestCenterEmptyRow.xlsx";
        await using (var stream = File.OpenRead(path))
        {
            var d = (await stream.QueryAsync()).Cast<IDictionary<string, object>>();
            var rows = d.ToList();
            Assert.Equal("a", rows[0]["A"]);
            Assert.Equal("b", rows[0]["B"]);
            Assert.Equal("c", rows[0]["C"]);
            Assert.Equal("d", rows[0]["D"]);

            Assert.Equal(1d, rows[1]["A"]);
            Assert.Null(rows[1]["B"]);
            Assert.Equal(3d, rows[1]["C"]);
            Assert.Null(rows[1]["D"]);

            Assert.Null(rows[2]["A"]);
            Assert.Equal(2d, rows[2]["B"]);
            Assert.Null(rows[2]["C"]);
            Assert.Equal(4d, rows[2]["D"]);

            Assert.Null(rows[3]["A"]);
            Assert.Null(rows[3]["B"]);
            Assert.Null(rows[3]["C"]);
            Assert.Null(rows[3]["D"]);

            Assert.Equal(1d, rows[4]["A"]);
            Assert.Null(rows[4]["B"]);
            Assert.Equal(3d, rows[4]["C"]);
            Assert.Null(rows[4]["D"]);

            Assert.Null(rows[5]["A"]);
            Assert.Equal(2d, rows[5]["B"]);
            Assert.Null(rows[5]["C"]);
            Assert.Equal(4d, rows[5]["D"]);

        }

        await using (var stream = File.OpenRead(path))
        {
            var d = (await stream.QueryAsync(useHeaderRow: true)).Cast<IDictionary<string, object>>();
            var rows = d.ToList();
            Assert.Equal(1d, rows[0]["a"]);
            Assert.Null(rows[0]["b"]);
            Assert.Equal(3d, rows[0]["c"]);
            Assert.Null(rows[0]["d"]);

            Assert.Null(rows[1]["a"]);
            Assert.Equal(2d, rows[1]["b"]);
            Assert.Null(rows[1]["c"]);
            Assert.Equal(4d, rows[1]["d"]);

            Assert.Null(rows[2]["a"]);
            Assert.Null(rows[2]["b"]);
            Assert.Null(rows[2]["c"]);
            Assert.Null(rows[2]["d"]);

            Assert.Equal(1d, rows[3]["a"]);
            Assert.Null(rows[3]["b"]);
            Assert.Equal(3d, rows[3]["c"]);
            Assert.Null(rows[3]["d"]);

            Assert.Null(rows[4]["a"]);
            Assert.Equal(2d, rows[4]["b"]);
            Assert.Null(rows[4]["c"]);
            Assert.Equal(4d, rows[4]["d"]);
        }
    }

    [Fact]
    public async Task TestDynamicQueryBasic_WithoutHead()
    {
        const string path = "../../../../../samples/xlsx/TestDynamicQueryBasic_WithoutHead.xlsx";
        await using var stream = File.OpenRead(path);
        var d = (await stream.QueryAsync()).Cast<IDictionary<string, object>>();
        var rows = d.ToList();

        Assert.Equal("MiniExcel", rows[0]["A"]);
        Assert.Equal(1d, rows[0]["B"]);
        Assert.Equal("Github", rows[1]["A"]);
        Assert.Equal(2d, rows[1]["B"]);
    }

    [Fact]
    public async Task TestDynamicQueryBasic_useHeaderRow()
    {
        const string path = "../../../../../samples/xlsx/TestDynamicQueryBasic.xlsx";
        await using (var stream = File.OpenRead(path))
        {
            var d = (await stream.QueryAsync(useHeaderRow: true)).Cast<IDictionary<string, object>>();
            var rows = d.ToList();
            Assert.Equal("MiniExcel", rows[0]["Column1"]);
            Assert.Equal(1d, rows[0]["Column2"]);
            Assert.Equal("Github", rows[1]["Column1"]);
            Assert.Equal(2d, rows[1]["Column2"]);
        }

        {
            var d = await MiniExcel.QueryAsync(path, useHeaderRow: true);
            var rows = d.ToList();
            Assert.Equal("MiniExcel", rows[0].Column1);
            Assert.Equal(1d, rows[0].Column2);
            Assert.Equal("Github", rows[1].Column1);
            Assert.Equal(2d, rows[1].Column2);
        }
    }

    private class DemoPocoHelloWorld
    {
        public string HelloWorld1 { get; set; }
    }

    private class UserAccount
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
    public async Task QueryStrongTypeMapping_Test()
    {
        const string path = "../../../../../samples/xlsx/TestTypeMapping.xlsx";
        await using (var stream = File.OpenRead(path))
        {
            var d = await stream.QueryAsync<UserAccount>();
            var rows = d.ToList();
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
            var rows = MiniExcel.Query<UserAccount>(path).ToList();
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
        public Guid? @guid { get; set; }
        public bool? @bool { get; set; }
        public DateTime? datetime { get; set; }
        public string @string { get; set; }
    }

    [Fact]
    public async Task AutoCheckTypeTest()
    {
        const string path = "../../../../../samples/xlsx/TestTypeMapping_AutoCheckFormat.xlsx";
        await using var stream = FileHelper.OpenRead(path);
        _ = (await stream.QueryAsync<AutoCheckType>()).ToList();
    }

    [Fact]
    public async Task TestDatetimeSpanFormat_ClosedXml()
    {
        const string path = "../../../../../samples/xlsx/TestDatetimeSpanFormat_ClosedXml.xlsx";
        await using var stream = FileHelper.OpenRead(path);
        
        var d = (await stream.QueryAsync()).Cast<IDictionary<string, object>>();
        var row = d.First();
        var a = row["A"];
        var b = row["B"];
        
        Assert.Equal(DateTime.Parse("2021-03-20T23:39:42.3130000"), (DateTime)a);
        Assert.Equal(TimeSpan.FromHours(10), (TimeSpan)b);
    }

    [Fact]
    public async Task LargeFileQueryStrongTypeMapping_Test()
    {
        const string path = "../../../../../benchmarks/MiniExcel.Benchmarks/Test1,000,000x10.xlsx";
        await using (var stream = File.OpenRead(path))
        {
            var d = await stream.QueryAsync<DemoPocoHelloWorld>();
            var rows = d.Take(2).ToList();
            Assert.Equal("HelloWorld2", rows[0].HelloWorld1);
            Assert.Equal("HelloWorld3", rows[1].HelloWorld1);
        }
        {
            var d = await MiniExcel.QueryAsync<DemoPocoHelloWorld>(path);
            var rows = d.Take(2).ToList();
            Assert.Equal("HelloWorld2", rows[0].HelloWorld1);
            Assert.Equal("HelloWorld3", rows[1].HelloWorld1);
        }
    }

    [Theory]
    [InlineData("../../../../../samples/xlsx/ExcelDataReaderCollections/TestChess.xlsx")]
    [InlineData("../../../../../samples/xlsx/TestCenterEmptyRow/TestCenterEmptyRow.xlsx")]
    public async Task QueryExcelDataReaderCheckTest(string path)
    {
#if NETCOREAPP3_1_OR_GREATER
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
#endif

        await using var fs = File.OpenRead(path);
        using var reader = ExcelDataReader.ExcelReaderFactory.CreateReader(fs);
        var exceldatareaderResult = reader.AsDataSet();
        await using var stream = File.OpenRead(path);

        var d = await stream.QueryAsync();
        var rows = d.ToList();
        Assert.Equal(exceldatareaderResult.Tables[0].Rows.Count, rows.Count);
        
        foreach (IDictionary<string, object?> row in rows)
        {
            var rowIndex = rows.IndexOf(row);
            foreach (var (key, value) in row)
            {
                var eV = exceldatareaderResult.Tables[0].Rows[rowIndex][Helpers.GetColumnIndex(key)];
                var v = value ?? DBNull.Value;
                Assert.Equal(eV, v);
            }
        }
    }

    [Fact]
    public async Task QuerySheetWithoutRAttribute()
    {
        const string path = "../../../../../samples/xlsx/TestWihoutRAttribute.xlsx";
        await using var stream = File.OpenRead(path);
        
        var d = (await stream.QueryAsync()).Cast<IDictionary<string, object>>();
        var rows = d.ToList();
        var keys = rows.First().Keys;

        Assert.Equal(2, rows.Count);
        Assert.Equal(5, keys.Count);

        Assert.Equal(1d, rows[0]["A"]);
        Assert.Null(rows[0]["C"]);
        Assert.Null(rows[0]["D"]);
        Assert.Null(rows[0]["E"]);

        Assert.Equal(1d, rows[1]["A"]);
        Assert.Equal("\"<>+}{\\nHello World", rows[1]["B"]);
        Assert.Equal(true, rows[1]["C"]);
        Assert.Equal("2021-03-16T19:10:21", rows[1]["D"]);
    }

    [Fact]
    public async Task FixDimensionJustOneColumnParsingError_Test()
    {
        const string path = "../../../../../samples/xlsx/TestDimensionC3.xlsx";
        await using var stream = File.OpenRead(path);
        var d = await stream.QueryAsync();
        var rows = d.ToList();
        var keys = (rows.First() as IDictionary<string, object>)?.Keys;
        Assert.Equal(3, keys?.Count);
        Assert.Equal(2, rows.Count);
    }

    private class SaveAsFileWithDimensionByICollectionTestType
    {
        public string A { get; set; }
        public string B { get; set; }
    }
    [Fact]
    public async Task SaveAsFileWithDimensionByICollection()
    {
        //List<strongtype>
        {
            List<SaveAsFileWithDimensionByICollectionTestType> values=
            [
                new() { A = "A", B = "B" },
                new() { A = "A", B = "B" }
            ];
            
            using (var file = AutoDeletingPath.Create())
            {
                var path = file.ToString();
                await MiniExcel.SaveAsAsync(path, values);
                await using (var stream = File.OpenRead(path))
                {
                    var d = (await stream.QueryAsync(useHeaderRow: false)).Cast<IDictionary<string, object>>();
                    var rows = d.ToList();
                    Assert.Equal(3, rows.Count);
                    Assert.Equal("A", rows[0]["A"]);
                    Assert.Equal("A", rows[1]["A"]);
                    Assert.Equal("A", rows[2]["A"]);
                }

                await using (var stream = File.OpenRead(path))
                {
                    var d = (await stream.QueryAsync(useHeaderRow: true)).Cast<IDictionary<string, object>>();
                    var rows = d.ToList();
                    Assert.Equal(2, rows.Count);
                    Assert.Equal("A", rows[0]["A"]);
                    Assert.Equal("A", rows[1]["A"]);
                }

                Assert.Equal("A1:B3", Helpers.GetFirstSheetDimensionRefValue(path));
            }
            
            using var newPath = AutoDeletingPath.Create();
            await MiniExcel.SaveAsAsync(newPath.ToString(), values, false);
            Assert.Equal("A1:B2", Helpers.GetFirstSheetDimensionRefValue(newPath.ToString()));
        }

        //List<strongtype> empty
        {
            List<SaveAsFileWithDimensionByICollectionTestType> values = [];
            using (var file = AutoDeletingPath.Create())
            {
                var path = file.ToString();
                await MiniExcel.SaveAsAsync(path, values, false);
                await using (var stream = File.OpenRead(path))
                {
                    var d = await stream.QueryAsync(useHeaderRow: false);
                    var rows = d.ToList();
                    Assert.Empty(rows);
                }

                Assert.Equal("A1:B1", Helpers.GetFirstSheetDimensionRefValue(path));
            }

            using (var file = AutoDeletingPath.Create())
            {
                var path = file.ToString();
                await MiniExcel.SaveAsAsync(path, values);
                {
                    await using var stream = File.OpenRead(path);
                    var d = await stream.QueryAsync(useHeaderRow: false);
                    var rows = d.ToList();
                    Assert.Single(rows);
                }
                Assert.Equal("A1:B1", Helpers.GetFirstSheetDimensionRefValue(path));
            }
        }

        //Array<anoymous>
        {
            var values = new[]
            {
                new { A = "A", B = "B" },
                new { A = "A", B = "B" },
            };
            
            using (var file = AutoDeletingPath.Create())
            {
                var path = file.ToString();
                await MiniExcel.SaveAsAsync(path, values);

                await using (var stream = File.OpenRead(path))
                {
                    var d = (await stream.QueryAsync(useHeaderRow: false)).Cast<IDictionary<string, object>>();
                    var rows = d.ToList();
                    Assert.Equal(3, rows.Count);
                    Assert.Equal("A", rows[0]["A"]);
                    Assert.Equal("A", rows[1]["A"]);
                    Assert.Equal("A", rows[2]["A"]);
                }

                await using (var stream = File.OpenRead(path))
                {
                    var d = (await stream.QueryAsync(useHeaderRow: true)).Cast<IDictionary<string, object>>();
                    var rows = d.ToList();
                    Assert.Equal(2, rows.Count);
                    Assert.Equal("A", rows[0]["A"]);
                    Assert.Equal("A", rows[1]["A"]);
                }

                Assert.Equal("A1:B3", Helpers.GetFirstSheetDimensionRefValue(path));
            }

            using (var path = AutoDeletingPath.Create())
            {
                await MiniExcel.SaveAsAsync(path.ToString(), values, false);
                Assert.Equal("A1:B2", Helpers.GetFirstSheetDimensionRefValue(path.ToString()));
            }
        }

        // without properties
        {
            using var path = AutoDeletingPath.Create();
            var values = new List<int>();
            await Assert.ThrowsAsync<NotSupportedException>(() => MiniExcel.SaveAsAsync(path.ToString(), values));
        }
    }

    [Fact]
    public async Task SaveAsFileWithDimension()
    {
        {
            using var file = AutoDeletingPath.Create();
            var path = file.ToString();

            using var table = new DataTable();
            await MiniExcel.SaveAsAsync(path, table);
            Assert.Equal("A1", Helpers.GetFirstSheetDimensionRefValue(path));
            {
                await using var stream = File.OpenRead(path);
                var d = await stream.QueryAsync();
                var rows = d.ToList();
                Assert.Single(rows); 
            }
            await MiniExcel.SaveAsAsync(path, table, printHeader: false, overwriteFile: true);
            Assert.Equal("A1", Helpers.GetFirstSheetDimensionRefValue(path));
        }
        {
            using var file = AutoDeletingPath.Create();
            var path = file.ToString();
            using var table = new DataTable();
            
            table.Columns.Add("a", typeof(string));
            table.Columns.Add("b", typeof(decimal));
            table.Columns.Add("c", typeof(bool));
            table.Columns.Add("d", typeof(DateTime));
            table.Rows.Add(@"""<>+-*//}{\\n", 1234567890);
            table.Rows.Add("<test>Hello World</test>", -1234567890, false, DateTime.Now);

            await MiniExcel.SaveAsAsync(path, table);
            Assert.Equal("A1:D3", Helpers.GetFirstSheetDimensionRefValue(path));

            await using (var stream = File.OpenRead(path))
            {
                var d = (await stream.QueryAsync(useHeaderRow: true)).Cast<IDictionary<string, object>>();
                var rows = d.ToList();
                Assert.Equal(2, rows.Count);
                Assert.Equal(@"""<>+-*//}{\\n", rows[0]["a"]);
                Assert.Equal(1234567890d, rows[0]["b"]);
                Assert.Null(rows[0]["c"]);
                Assert.Null(rows[0]["d"]);
            }

            await using (var stream = File.OpenRead(path))
            {
                var d = (await stream.QueryAsync()).Cast<IDictionary<string, object>>();
                var rows = d.ToList();
                Assert.Equal(3, rows.Count);
                Assert.Equal("a", rows[0]["A"]);
                Assert.Equal("b", rows[0]["B"]);
                Assert.Equal("c", rows[0]["C"]);
                Assert.Equal("d", rows[0]["D"]);
            }

            await MiniExcel.SaveAsAsync(path, table, printHeader: false, overwriteFile:true);
            Assert.Equal("A1:D2", Helpers.GetFirstSheetDimensionRefValue(path));
        }

        //TODO:StartCell
        {
            using var path = AutoDeletingPath.Create();
            using var table = new DataTable();
            table.Columns.Add("a", typeof(string));
            table.Rows.Add("A");
            table.Rows.Add("B");
            
            await MiniExcel.SaveAsAsync(path.ToString(), table);
            Assert.Equal("A3", Helpers.GetFirstSheetDimensionRefValue(path.ToString()));
        }
    }

    [Fact]
    public async Task SaveAsByDataTableTest()
    {
        {
            using var file = AutoDeletingPath.Create();
            var path = file.ToString();

            var now = DateTime.Now;

            using var table = new DataTable();
            table.Columns.Add("a", typeof(string));
            table.Columns.Add("b", typeof(decimal));
            table.Columns.Add("c", typeof(bool));
            table.Columns.Add("d", typeof(DateTime));
            table.Rows.Add(@"""<>+-*//}{\\n", 1234567890, true, now);
            table.Rows.Add("<test>Hello World</test>", -1234567890, false, now.Date);

            await MiniExcel.SaveAsAsync(path, table);

            using var p = new ExcelPackage(new FileInfo(path));
            var ws = p.Workbook.Worksheets.First();

            Assert.True(ws.Cells["A1"].Value.ToString() == "a");
            Assert.True(ws.Cells["B1"].Value.ToString() == "b");
            Assert.True(ws.Cells["C1"].Value.ToString() == "c");
            Assert.True(ws.Cells["D1"].Value.ToString() == "d");

            Assert.True(ws.Cells["A2"].Value.ToString() == @"""<>+-*//}{\\n");
            Assert.True(ws.Cells["B2"].Value.ToString() == @"1234567890");
            Assert.True(ws.Cells["C2"].Value.ToString() == true.ToString());
            Assert.True(ws.Cells["D2"].Value.ToString() == now.ToString());
        }
        {
            using var path = AutoDeletingPath.Create();
            using var table = new DataTable();
            table.Columns.Add("Column1", typeof(string));
            table.Columns.Add("Column2", typeof(int));
            table.Rows.Add("MiniExcel", 1);
            table.Rows.Add("Github", 2);

            await MiniExcel.SaveAsAsync(path.ToString(), table);
        }
    }

    [Fact]
    public async Task QueryByLINQExtensionsVoidTaskLargeFileOOMTest()
    {
        const string path = "../../../../../benchmarks/MiniExcel.Benchmarks/Test1,000,000x10.xlsx";

        {
            var row = (await MiniExcel.QueryAsync(path)).First();
            Assert.Equal("HelloWorld1", row.A);
        }

        await using (var stream = File.OpenRead(path))
        {
            var d = (await stream.QueryAsync()).Cast<IDictionary<string, object>>();
            var row = d.First();
            Assert.Equal("HelloWorld1", row["A"]);
        }

        {
            var d = (await MiniExcel.QueryAsync(path)).Cast<IDictionary<string, object>>();
            var rows = d.Take(10);
            Assert.Equal(10, rows.Count());
        }
    }

    [Fact]
    public async Task EmptyTest()
    {
        using var path = AutoDeletingPath.Create();

        await using (var connection = Db.GetConnection("Data Source=:memory:"))
        {
            var rows = await connection.QueryAsync("with cte as (select 1 id,2 val) select * from cte where 1=2");
            await MiniExcel.SaveAsAsync(path.ToString(), rows);
        }

        await using (var stream = File.OpenRead(path.ToString()))
        {
            var row = await stream.QueryAsync(useHeaderRow: true);
            Assert.Empty(row);
        }
    }

    [Fact]
    public async Task SaveAsByIEnumerableIDictionary()
    {

        {
            using var file = AutoDeletingPath.Create();
            var path = file.ToString();
            List<Dictionary<string, object>> values =
            [
                new() { { "Column1", "MiniExcel" }, { "Column2", 1 } },
                new() { { "Column1", "Github" }, { "Column2", 2 } }
            ];
            await MiniExcel.SaveAsAsync(path, values);

            await using (var stream = File.OpenRead(path))
            {
                var d = (await stream.QueryAsync(useHeaderRow: false)).Cast<IDictionary<string, object>>();
                var rows = d.ToList();
                Assert.Equal("Column1", rows[0]["A"]);
                Assert.Equal("Column2", rows[0]["B"]);
                Assert.Equal("MiniExcel", rows[1]["A"]);
                Assert.Equal(1d, rows[1]["B"]);
                Assert.Equal("Github", rows[2]["A"]);
                Assert.Equal(2d, rows[2]["B"]);
            }

            await using (var stream = File.OpenRead(path))
            {
                var d = (await stream.QueryAsync(useHeaderRow: true)).Cast<IDictionary<string, object>>();
                var rows = d.ToList();
                Assert.Equal(2, rows.Count);
                Assert.Equal("MiniExcel", rows[0]["Column1"]);
                Assert.Equal(1d, rows[0]["Column2"]);
                Assert.Equal("Github", rows[1]["Column1"]);
                Assert.Equal(2d, rows[1]["Column2"]);
            }

            Assert.Equal("A1:B3", Helpers.GetFirstSheetDimensionRefValue(path));
        }

        {
            using var file = AutoDeletingPath.Create();
            var path = file.ToString();

            List<Dictionary<int, object>> values =
            [
                new() { { 1, "MiniExcel" }, { 2, 1 } },
                new() { { 1, "Github" }, { 2, 2 } }
            ];
            await MiniExcel.SaveAsAsync(path, values);

            await using (var stream = File.OpenRead(path))
            {
                var d = await stream.QueryAsync(useHeaderRow: false);
                var rows = d.ToList();
                Assert.Equal(3, rows.Count);
            }
            Assert.Equal("A1:B3", Helpers.GetFirstSheetDimensionRefValue(path));
        }
    }

    [Fact]
    public async Task SaveAsByIEnumerableIDictionaryWithDynamicConfiguration()
    {
        using var file = AutoDeletingPath.Create();
        var path = file.ToString();

        var dynamicColumns = new[]
        {
            new DynamicExcelColumn("Column1") { Name = "Name Column" },
            new DynamicExcelColumn("Column2") { Name = "Value Column" }
        };
        var config = new OpenXmlConfiguration
        {
            DynamicColumns = dynamicColumns
        };
        List<Dictionary<string, object>> values =
        [
            new() { { "Column1", "MiniExcel" }, { "Column2", 1 } },
            new() { { "Column1", "Github" }, { "Column2", 2 } }
        ];
        await MiniExcel.SaveAsAsync(path, values, configuration: config);

        await using (var stream = File.OpenRead(path))
        {
            var d = (await stream.QueryAsync(useHeaderRow: true)).Cast<IDictionary<string, object>>();
            var rows = d.ToList();
            Assert.Equal(2, rows.Count);
            Assert.Equal("Name Column", rows[0].Keys.ElementAt(0));
            Assert.Equal("Value Column", rows[0].Keys.ElementAt(1));
            Assert.Equal("MiniExcel", rows[0].Values.ElementAt(0));
            Assert.Equal(1d, rows[0].Values.ElementAt(1));
            Assert.Equal("Github", rows[1].Values.ElementAt(0));
            Assert.Equal(2d, rows[1].Values.ElementAt(1));
        }
        Assert.Equal("A1:B3", Helpers.GetFirstSheetDimensionRefValue(path));
    }


    [Fact]
    public async Task SaveAsFrozenRowsAndColumnsTest()
    {
        var config = new OpenXmlConfiguration
        {
            FreezeRowCount = 1,
            FreezeColumnCount = 2
        };

        // Test enumerable
        using var file = AutoDeletingPath.Create();
        var path = file.ToString();

        await MiniExcel.SaveAsAsync(
            path,
            new[] 
            {
                new { Column1 = "MiniExcel", Column2 = 1 },
                new { Column1 = "Github", Column2 = 2}
            },
            configuration: config
        );

        await using (var stream = File.OpenRead(path))
        {
            var rows = (await stream.QueryAsync(useHeaderRow: true)).ToList();

            Assert.Equal("MiniExcel", rows[0].Column1);
            Assert.Equal(1, rows[0].Column2);
            Assert.Equal("Github", rows[1].Column1);
            Assert.Equal(2, rows[1].Column2);
        }
        Assert.Equal("A1:B3", Helpers.GetFirstSheetDimensionRefValue(path));

        // test table
        var table = new DataTable();
        table.Columns.Add("a", typeof(string));
        table.Columns.Add("b", typeof(decimal));
        table.Columns.Add("c", typeof(bool));
        table.Columns.Add("d", typeof(DateTime));
        table.Rows.Add("some text", 1234567890, true, DateTime.Now);
        table.Rows.Add("<test>Hello World</test>", -1234567890, false, DateTime.Now.Date);
        
        using var pathTable = AutoDeletingPath.Create();
        await MiniExcel.SaveAsAsync(pathTable.ToString(), table, configuration: config);
        Assert.Equal("A1:D3", Helpers.GetFirstSheetDimensionRefValue(pathTable.ToString()));

        // data reader
        await using var reader = table.CreateDataReader();
        using var pathReader = AutoDeletingPath.Create();
        await MiniExcel.SaveAsAsync(pathReader.ToString(), reader, configuration: config);
        Assert.Equal("A1:D3", Helpers.GetFirstSheetDimensionRefValue(pathTable.ToString()));
    }

    [Fact]
    public async Task SaveAsByDapperRows()
    {
        using var file = AutoDeletingPath.Create();
        var path = file.ToString();

        // Dapper Query
        await using (var connection = Db.GetConnection("Data Source=:memory:"))
        {
            var rows = await connection.QueryAsync("select 'MiniExcel' as Column1,1 as Column2 union all select 'Github',2");
            await MiniExcel.SaveAsAsync(path, rows);
        }
        Assert.Equal("A1:B3", Helpers.GetFirstSheetDimensionRefValue(path));

        await using (var stream = File.OpenRead(path))
        {
            var rows = (await stream.QueryAsync(useHeaderRow: true)).Cast<IDictionary<string, object>>().ToList();
            Assert.Equal("MiniExcel", rows[0]["Column1"]);
            Assert.Equal(1d, rows[0]["Column2"]);
            Assert.Equal("Github", rows[1]["Column1"]);
            Assert.Equal(2d, rows[1]["Column2"]);
        }

        // Empty
        await using (var connection = Db.GetConnection("Data Source=:memory:"))
        {
            var rows = (await connection.QueryAsync("with cte as (select 'MiniExcel' as Column1,1 as Column2 union all select 'Github',2)select * from cte where 1=2")).ToList();
            await MiniExcel.SaveAsAsync(path, rows, overwriteFile: true);
        }

        await using (var stream = File.OpenRead(path))
        {
            var rows = (await stream.QueryAsync(useHeaderRow: false)).ToList();
            Assert.Empty(rows);
        }

        await using (var stream = File.OpenRead(path))
        {
            var rows = (await stream.QueryAsync(useHeaderRow: true)).ToList();
            Assert.Empty(rows);
        }
        Assert.Equal("A1", Helpers.GetFirstSheetDimensionRefValue(path));

        // ToList
        await using (var connection = Db.GetConnection("Data Source=:memory:"))
        {
            var rows = (await connection.QueryAsync("select 'MiniExcel' as Column1,1 as Column2 union all select 'Github',2")).ToList();
            await MiniExcel.SaveAsAsync(path, rows, overwriteFile: true);
        }
        Assert.Equal("A1:B3", Helpers.GetFirstSheetDimensionRefValue(path));

        await using (var stream = File.OpenRead(path))
        {
            var rows = (await stream.QueryAsync(useHeaderRow: false)).Cast<IDictionary<string, object>>().ToList();
            Assert.Equal("Column1", rows[0]["A"]);
            Assert.Equal("Column2", rows[0]["B"]);
            Assert.Equal("MiniExcel", rows[1]["A"]);
            Assert.Equal(1d, rows[1]["B"]);
            Assert.Equal("Github", rows[2]["A"]);
            Assert.Equal(2d, rows[2]["B"]);
        }

        await using (var stream = File.OpenRead(path))
        {
            var rows = (await stream.QueryAsync(useHeaderRow: true)).Cast<IDictionary<string, object>>().ToList();
            Assert.Equal("MiniExcel", rows[0]["Column1"]);
            Assert.Equal(1d, rows[0]["Column2"]);
            Assert.Equal("Github", rows[1]["Column1"]);
            Assert.Equal(2d, rows[1]["Column2"]);
        }
    }


    private class Demo
    {
        public string Column1 { get; set; }
        public decimal Column2 { get; set; }
    }
    [Fact]
    public async Task QueryByStrongTypeParameterTest()
    {
        using var file = AutoDeletingPath.Create();
        var path = file.ToString();
        List<Demo> values =
        [
            new() { Column1 = "MiniExcel", Column2 = 1 },
            new() { Column1 = "Github", Column2 = 2 }
        ];
        await MiniExcel.SaveAsAsync(path, values);

        await using var stream = File.OpenRead(path);
        var rows = (await stream.QueryAsync(useHeaderRow: true)).Cast<IDictionary<string, object>>().ToList();
        
        Assert.Equal("MiniExcel", rows[0]["Column1"]);
        Assert.Equal(1d, rows[0]["Column2"]);
        Assert.Equal("Github", rows[1]["Column1"]);
        Assert.Equal(2d, rows[1]["Column2"]);
    }

    [Fact]
    public async Task QueryByDictionaryStringAndObjectParameterTest()
    {
        using var file = AutoDeletingPath.Create();
        var path = file.ToString();
        List<Dictionary<string, object>> values =
        [
            new() { { "Column1", "MiniExcel" }, { "Column2", 1 } },
            new() { { "Column1", "Github" }, { "Column2", 2 } }
        ];
        await MiniExcel.SaveAsAsync(path, values);

        await using var stream = File.OpenRead(path);
        var rows = (await stream.QueryAsync(useHeaderRow: true)).Cast<IDictionary<string, object>>().ToList();
        
        Assert.Equal("MiniExcel", rows[0]["Column1"]);
        Assert.Equal(1d, rows[0]["Column2"]);
        Assert.Equal("Github", rows[1]["Column1"]);
        Assert.Equal(2d, rows[1]["Column2"]);
    }

    [Fact]
    public async Task SQLiteInsertTest()
    {
        // Async Task SQL Insert Large Size Xlsx OOM
        const string path = "../../../../../samples/xlsx/Test5x2.xlsx";
        using var tempSqlitePath = AutoDeletingPath.Create(Path.GetTempPath(), $"{Guid.NewGuid()}.db");
        var connectionString = $"Data Source={tempSqlitePath};Version=3;";

        await using (var connection = new SQLiteConnection(connectionString))
        {
            await connection.ExecuteAsync("create table T (A varchar(20),B varchar(20));");
        }

        await using (var connection = new SQLiteConnection(connectionString))
        {
            await connection.OpenAsync();
            await using (var transaction = connection.BeginTransaction())
            await using (var stream = File.OpenRead(path))
            {
                var rows = (await stream.QueryAsync()).Cast<IDictionary<string, object>>();
                foreach (var row in rows)
                    await connection.ExecuteAsync(
                        "insert into T (A,B) values (@A,@B)",
                        new { A = row["A"], B = row["B"] }, 
                        transaction: transaction
                    );
                
                await transaction.CommitAsync();
            }
        }

        await using (var connection = new SQLiteConnection(connectionString))
        {
            var result = await connection.QueryAsync("select * from T");
            Assert.Equal(5, result.Count());
        }
    }

    [Fact]
    public async Task SaveAsBasicCreateTest()
    {
        using var file = AutoDeletingPath.Create();
        var path = file.ToString();
        
        await MiniExcel.SaveAsAsync(path, new[] 
        {
            new { Column1 = "MiniExcel", Column2 = 1 },
            new { Column1 = "Github", Column2 = 2}
        });

        await using (var stream = File.OpenRead(path))
        {
            var d = (await stream.QueryAsync(useHeaderRow: true)).Cast<IDictionary<string, object>>();
            var rows = d.ToList();
            Assert.Equal("MiniExcel", rows[0]["Column1"]);
            Assert.Equal(1d, rows[0]["Column2"]);
            Assert.Equal("Github", rows[1]["Column1"]);
            Assert.Equal(2d, rows[1]["Column2"]);
        }
        Assert.Equal("A1:B3", Helpers.GetFirstSheetDimensionRefValue(path));
    }

    [Fact]
    public async Task SaveAsBasicStreamTest()
    {
        {
            using var file = AutoDeletingPath.Create();
            var path = file.ToString();
            
            var values = new[] 
            {
                new { Column1 = "MiniExcel", Column2 = 1 },
                new { Column1 = "Github", Column2 = 2}
            };
            await using (var stream = new FileStream(path, FileMode.CreateNew))
            {
                await stream.SaveAsAsync(values);
            }

            await using (var stream = File.OpenRead(path))
            {
                var rows = (await stream.QueryAsync(useHeaderRow: true)).Cast<IDictionary<string, object>>().ToList();
                Assert.Equal("MiniExcel", rows[0]["Column1"]);
                Assert.Equal(1d, rows[0]["Column2"]);
                Assert.Equal("Github", rows[1]["Column1"]);
                Assert.Equal(2d, rows[1]["Column2"]);
            }
        }
        {
            using var file = AutoDeletingPath.Create();
            var path = file.ToString();
            var values = new[] 
            {
                new { Column1 = "MiniExcel", Column2 = 1 },
                new { Column1 = "Github", Column2 = 2}
            };
            await using (var stream = new MemoryStream())
            await using (var fileStream = new FileStream(path, FileMode.Create))
            {
                await stream.SaveAsAsync(values);
                stream.Seek(0, SeekOrigin.Begin);
                await stream.CopyToAsync(fileStream);
            }

            await using (var stream = File.OpenRead(path))
            {
                var rows = (await stream.QueryAsync(useHeaderRow: true)).Cast<IDictionary<string, object>>().ToList();
                Assert.Equal("MiniExcel", rows[0]["Column1"]);
                Assert.Equal(1d, rows[0]["Column2"]);
                Assert.Equal("Github", rows[1]["Column1"]);
                Assert.Equal(2d, rows[1]["Column2"]);
            }
        }
    }

    [Fact]
    public async Task SaveAsSpecialAndTypeCreateTest()
    {
        using var path = AutoDeletingPath.Create();
        await MiniExcel.SaveAsAsync(path.ToString(), new[] 
        {
            new { a = @"""<>+-*//}{\\n", b = 1234567890, c = true, d = DateTime.Now },
            new { a = "<test>Hello World</test>", b = -1234567890, c = false, d = DateTime.Now.Date}
        });
        var info = new FileInfo(path.ToString());
        Assert.True(info.FullName == path.ToString());
    }

    [Fact]
    public async Task SaveAsFileEpplusCanReadTest()
    {
        using var path = AutoDeletingPath.Create();
        var now = DateTime.Now;

        await MiniExcel.SaveAsAsync(path.ToString(), new[] 
        {
            new { a = @"""<>+-*//}{\\n", b = 1234567890, c = true, d = now},
            new { a = "<test>Hello World</test>", b = -1234567890, c = false, d = now.Date}
        });
        
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
    public async Task SavaAsClosedXmlCanReadTest()
    {
        var now = DateTime.Now;
        using var path = AutoDeletingPath.Create();

        await MiniExcel.SaveAsAsync(path.ToString(), new[] 
        {
            new { a = @"""<>+-*//}{\\n", b = 1234567890, c = true, d = now},
            new { a = "<test>Hello World</test>", b = -1234567890, c = false, d = now.Date}
        });
        using var workbook = new XLWorkbook(path.ToString());
        var ws = workbook.Worksheets.First();

        Assert.True(ws.Cell("A1").Value.ToString() == "a");
        Assert.True(ws.Cell("D1").Value.ToString() == "d");
        Assert.True(ws.Cell("B1").Value.ToString() == "b");
        Assert.True(ws.Cell("C1").Value.ToString() == "c");

        Assert.True(ws.Cell("A2").Value.ToString() == @"""<>+-*//}{\\n");
        Assert.True(ws.Cell("B2").Value.ToString() == "1234567890");
        Assert.True(ws.Cell("C2").Value.ToString() == true.ToString());
        Assert.True(ws.Cell("D2").Value.ToString() == now.ToString());
    }

    [Fact]
    public async Task ContentTypeUriContentTypeReadCheckTest()
    {
        var now = DateTime.Now;
        using var path = AutoDeletingPath.Create();

        await MiniExcel.SaveAsAsync(path.ToString(), new[]
        {
            new { a = @"""<>+-*//}{\\n", b = 1234567890, c = true, d = now},
            new { a = "<test>Hello World</test>", b = -1234567890, c = false, d = now.Date}
        });
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
    public async Task ReadBigExcel_TakeCancel_Throws_TaskCanceledException()
    {
        await Assert.ThrowsAsync<TaskCanceledException>(async () =>
        {
            const string path = "../../../../../samples/xlsx/bigExcel.xlsx";
            using var cts = new CancellationTokenSource();

            cts.CancelAsync();

            await using var stream = FileHelper.OpenRead(path);
            var rows = (await stream.QueryAsync(cancellationToken: cts.Token)).ToList();
        });
    }

    [Fact]
    public async Task ReadBigExcel_Prcoessing_TakeCancel_Throws_TaskCanceledException()
    {
        await Assert.ThrowsAsync<OperationCanceledException>(async () =>
        {
            const string path = "../../../../../samples/xlsx/bigExcel.xlsx";
            var cts = new CancellationTokenSource();

            var cancelTask = Task.Run(async () =>
            {
                await Task.Delay(2000, cts.Token);
                await cts.CancelAsync();
                cts.Token.ThrowIfCancellationRequested();
            });

            await using var stream = FileHelper.OpenRead(path);
            var d = stream.QueryAsync(cancellationToken: cts.Token);
            await cancelTask;
            _ = (await d).ToList();
        });
    }

    [Fact]
    public async Task DynamicColumnsConfigurationIsUsedWhenCreatingExcelUsingIDataReader()
    {
        using var path = AutoDeletingPath.Create();
        var dateTime = DateTime.Now;
        var onlyDate = DateOnly.FromDateTime(dateTime);
        
        using var table = new DataTable();
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
        await using var reader = table.CreateDataReader();
        await MiniExcel.SaveAsAsync(path.ToString(), reader, configuration: configuration);

        await using var stream = File.OpenRead(path.ToString());
        var rows = (await stream.QueryAsync(useHeaderRow: true))
            .Cast<IDictionary<string, object>>()
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
    public async Task DynamicColumnsConfigurationIsUsedWhenCreatingExcelUsingDataTable()
    {
        using var path = AutoDeletingPath.Create();
        var dateTime = DateTime.Now;
        var onlyDate = DateOnly.FromDateTime(dateTime);
        
        using var table = new DataTable();
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
        await MiniExcel.SaveAsAsync(path.ToString(), table, configuration: configuration);

        await using var stream = File.OpenRead(path.ToString());
        var rows = (await stream.QueryAsync(useHeaderRow: true))
            .Cast<IDictionary<string, object>>()
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
    public async Task SaveAsByMiniExcelDataReader()
    {
        using var path1 = AutoDeletingPath.Create();
        var values = new List<Demo>
        {
            new() { Column1= "MiniExcel" ,Column2 = 1 },
            new() { Column1 = "Github", Column2 = 2 }
        };
        await MiniExcel.SaveAsAsync(path1.ToString(), values);

        await using (IMiniExcelDataReader? reader = MiniExcel.GetReader(path1.ToString(), true))
        {
            using var path2 = AutoDeletingPath.Create();
            await MiniExcel.SaveAsAsync(path2.ToString(), reader);
            var results = (await MiniExcel.QueryAsync<Demo>(path2.ToString())).ToList();

            Assert.True(results.Count == 2);
            Assert.True(results.First().Column1 == "MiniExcel");
            Assert.True(results.First().Column2 == 1);
            Assert.True(results.Last().Column1 == "Github");
            Assert.True(results.Last().Column2 == 2);
        }
    }

    [Fact]
    public async Task InsertSheetTest()
    {
        var now = DateTime.Now;
        using var file = AutoDeletingPath.Create();
        var path = file.ToString();

        {
            using var table = new DataTable();
            table.Columns.Add("a", typeof(string));
            table.Columns.Add("b", typeof(decimal));
            table.Columns.Add("c", typeof(bool));
            table.Columns.Add("d", typeof(DateTime));
            table.Rows.Add(@"""<>+-*//}{\\n", 1234567890, true, now);
            table.Rows.Add("<test>Hello World</test>", -1234567890, false, now.Date);

            await MiniExcel.InsertAsync(path, table, sheetName: "Sheet1");
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
            using var table = new DataTable();
            table.Columns.Add("Column1", typeof(string));
            table.Columns.Add("Column2", typeof(int));
            table.Rows.Add("MiniExcel", 1);
            table.Rows.Add("Github", 2);

            await MiniExcel.InsertAsync(path, table, sheetName: "Sheet2");
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
            using var table = new DataTable();
            table.Columns.Add("Column1", typeof(string));
            table.Columns.Add("Column2", typeof(DateTime));
            table.Rows.Add("Test", now);

            await MiniExcel.InsertAsync(path, table, sheetName: "Sheet2", printHeader: false, configuration: new OpenXmlConfiguration
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

            using var p = new ExcelPackage(new FileInfo(path));
            var sheet2 = p.Workbook.Worksheets[1];

            Assert.True(sheet2.Cells["A1"].Value.ToString() == "Test");
            Assert.True(sheet2.Cells["B1"].Text == now.ToString("dd.MM.yyyy HH:mm:ss"));
            Assert.True(sheet2.Name == "Sheet2");
        }
        {
            using var table = new DataTable();
            table.Columns.Add("Column1", typeof(string));
            table.Columns.Add("Column2", typeof(DateTime));
            table.Rows.Add("MiniExcel", now);
            table.Rows.Add("Github", now);

            await MiniExcel.InsertAsync(path, table, sheetName: "Sheet3", configuration: new OpenXmlConfiguration
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

    [Fact]
    public async Task InsertCsvTest()
    {
        using var file = AutoDeletingPath.Create(ExcelType.CSV);
        var path = file.ToString();

        {
            var value = new[] 
            {
                new { ID=1,Name ="Jack",InDate=new DateTime(2021,01,03)},
                new { ID=2,Name ="Henry",InDate=new DateTime(2020,05,03)},
            };
            await MiniExcel.SaveAsAsync(path, value);
            var content = await File.ReadAllTextAsync(path);
            Assert.Equal(
                """
                ID,Name,InDate
                1,Jack,"2021-01-03 00:00:00"
                2,Henry,"2020-05-03 00:00:00"

                """, content);
        }
        {
            var value = new { ID = 3, Name = "Mike", InDate = new DateTime(2021, 04, 23) };
            await MiniExcel.InsertAsync(path, value);
            var content = await File.ReadAllTextAsync(path);
            Assert.Equal(
                """
                ID,Name,InDate
                1,Jack,"2021-01-03 00:00:00"
                2,Henry,"2020-05-03 00:00:00"
                3,Mike,"2021-04-23 00:00:00"

                """, content);
        }
        {
            var value = new[]
            {
                new { ID=4,Name ="Frank",InDate=new DateTime(2021,06,07)},
                new { ID=5,Name ="Gloria",InDate=new DateTime(2022,05,03)},
            };
            
            await MiniExcel.InsertAsync(path, value);
            var content = await File.ReadAllTextAsync(path);
            Assert.Equal(
                """
                ID,Name,InDate
                1,Jack,"2021-01-03 00:00:00"
                2,Henry,"2020-05-03 00:00:00"
                3,Mike,"2021-04-23 00:00:00"
                4,Frank,"2021-06-07 00:00:00"
                5,Gloria,"2022-05-03 00:00:00"

                """, content);
        }
    }

    [Fact]
    public async Task SaveAsByAsyncEnumerable()
    {
        using var path = AutoDeletingPath.Create();

#pragma warning disable CS1998 // Async method lacks 'await' operators and will run synchronously
        static async IAsyncEnumerable<Demo> GetValues()
        {
            yield return new Demo { Column1 = "MiniExcel", Column2 = 1 };
            yield return new Demo { Column1 = "Github", Column2 = 2 };
        }
#pragma warning restore CS1998

        await MiniExcel.SaveAsAsync(path.ToString(), GetValues());
        var results = MiniExcel.Query<Demo>(path.ToString()).ToList();

        Assert.True(results.Count == 2);
        Assert.True(results.First().Column1 == "MiniExcel");
        Assert.True(results.First().Column2 == 1);
        Assert.True(results.Last().Column1 == "Github");
        Assert.True(results.Last().Column2 == 2);
    }
}