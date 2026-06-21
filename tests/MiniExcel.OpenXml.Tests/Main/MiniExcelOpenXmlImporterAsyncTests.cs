using ExcelDataReader;
using MiniExcelLib.Core.Exceptions;
using MiniExcelLib.OpenXml.Tests.Utils;
using MiniExcelLib.Tests.Common.Utils;

namespace MiniExcelLib.OpenXml.Tests.Main;

public class MiniExcelOpenXmlImporterAsyncTests
{
    private readonly OpenXmlImporter _excelImporter =  MiniExcel.Importers.GetOpenXmlImporter();
    private readonly OpenXmlExporter _excelExporter =  MiniExcel.Exporters.GetOpenXmlExporter();
   
    static MiniExcelOpenXmlImporterAsyncTests()
    {
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
    }

    [Fact]
    public async Task CustomAttributeWihoutVaildPropertiesTest()
    {
        var path = PathHelper.GetFile("xlsx/TestCustomExcelColumnAttribute.xlsx");
        await Assert.ThrowsAsync<InvalidMappingException>(async () => await _excelImporter.QueryAsync<CustomAttributesWihoutVaildPropertiesTestPoco>(path).ToListAsync());
    }

    [Fact]
    public async Task QueryCustomAttributesTest()
    {
        var path = PathHelper.GetFile("xlsx/TestCustomExcelColumnAttribute.xlsx");
        var rows = await _excelImporter.QueryAsync<ExcelAttributeDemo>(path).ToListAsync();

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
        var path = PathHelper.GetFile("xlsx/TestCustomExcelColumnAttribute.xlsx");
        var rows = await _excelImporter.QueryAsync<ExcelAttributeDemo2>(path).ToListAsync();

        Assert.Equal("Column1", rows[0].Test1);
        Assert.Equal("Column2", rows[0].Test2);
        Assert.Null(rows[0].Test3);
        Assert.Equal("Test7", rows[0].Test4);
        Assert.Null(rows[0].Test5);
        Assert.Null(rows[0].Test6);
        Assert.Equal("Test4", rows[0].Test7);
    }

    [Fact]
    public async Task QueryCastToIDictionary()
    {
        var path = PathHelper.GetFile("xlsx/TestCenterEmptyRow/TestCenterEmptyRow.xlsx");
        await foreach (IDictionary<string, object> row in _excelImporter.QueryAsync(path))
        {
            _ = row;
        }
    }

    [Fact]
    public async Task CenterEmptyRowsQueryTest()
    {
        var path = PathHelper.GetFile("xlsx/TestCenterEmptyRow/TestCenterEmptyRow.xlsx");
        await using (var stream = File.OpenRead(path))
        {
            var rows = await _excelImporter.QueryAsync(stream).Cast<IDictionary<string, object>>().ToListAsync();
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
            var rows = await _excelImporter.QueryAsync(stream, hasHeaderRow: true).Cast<IDictionary<string, object>>().ToListAsync();
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
        var path = PathHelper.GetFile("xlsx/TestDynamicQueryBasic_WithoutHead.xlsx");
        await using var stream = File.OpenRead(path);
        var rows = await _excelImporter.QueryAsync(stream).Cast<IDictionary<string, object>>().ToListAsync();

        Assert.Equal("MiniExcel", rows[0]["A"]);
        Assert.Equal(1d, rows[0]["B"]);
        Assert.Equal("Github", rows[1]["A"]);
        Assert.Equal(2d, rows[1]["B"]);
    }

    [Fact]
    public async Task TestDynamicQueryBasic_useHeaderRow()
    {
        var path = PathHelper.GetFile("xlsx/TestDynamicQueryBasic.xlsx");
        await using (var stream = File.OpenRead(path))
        {
            var rows = await _excelImporter.QueryAsync(stream, hasHeaderRow: true).Cast<IDictionary<string, object>>().ToListAsync();
            Assert.Equal("MiniExcel", rows[0]["Column1"]);
            Assert.Equal(1d, rows[0]["Column2"]);
            Assert.Equal("Github", rows[1]["Column1"]);
            Assert.Equal(2d, rows[1]["Column2"]);
        }

        {
            var rows = await _excelImporter.QueryAsync(path, hasHeaderRow: true).ToListAsync();
            Assert.Equal("MiniExcel", rows[0].Column1);
            Assert.Equal(1d, rows[0].Column2);
            Assert.Equal("Github", rows[1].Column1);
            Assert.Equal(2d, rows[1].Column2);
        }
    }

    [Fact]
    public async Task QueryStrongTypeMapping_Test()
    {
        var path = PathHelper.GetFile("xlsx/TestTypeMapping.xlsx");
        await using (var stream = File.OpenRead(path))
        {
            var rows = await _excelImporter.QueryAsync<UserAccount>(stream).ToListAsync();
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
            var rows =  _excelImporter.Query(path, hasHeaderRow: true).ToList();
            Assert.Equal(100, rows.Count);

            Assert.Equal("78DE23D2-DCB6-BD3D-EC67-C112BBC322A2", rows[0].ID);
            Assert.Equal("Wade", rows[0].Name);
            Assert.Equal(new DateTime(2020, 9, 27), rows[0].BoD);
            Assert.Equal(36, rows[0].Age);
            Assert.False(rows[0].VIP);
            Assert.Equal(5019.12d, rows[0].Points);
            Assert.Null(rows[0].IgnoredProperty);
        }
    }

    [Fact]
    public async Task AutoCheckTypeTest()
    {
        var path = PathHelper.GetFile("xlsx/TestTypeMapping_AutoCheckFormat.xlsx");
        await using var stream = File.OpenRead(path);
        _ =  await _excelImporter.QueryAsync<AutoCheckType>(stream).ToListAsync();
    }

    [Fact]
    public async Task TestDatetimeSpanFormat_ClosedXml()
    {
        var path = PathHelper.GetFile("xlsx/TestDatetimeSpanFormat_ClosedXml.xlsx");
        await using var stream = File.OpenRead(path);

        var row = await _excelImporter.QueryAsync(stream).Cast<IDictionary<string, object>>().FirstAsync();
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
            var rows = await _excelImporter.QueryAsync<DemoPocoHelloWorld>(stream).Take(2).ToListAsync();
            Assert.Equal("HelloWorld2", rows[0].HelloWorld1);
            Assert.Equal("HelloWorld3", rows[1].HelloWorld1);
        }
        {
            var rows = await _excelImporter.QueryAsync<DemoPocoHelloWorld>(path).Take(2).ToListAsync();
            Assert.Equal("HelloWorld2", rows[0].HelloWorld1);
            Assert.Equal("HelloWorld3", rows[1].HelloWorld1);
        }
    }

    [Theory]
    [InlineData("../../../../data/xlsx/ExcelDataReaderCollections/TestChess.xlsx")]
    [InlineData("../../../../data/xlsx/TestCenterEmptyRow/TestCenterEmptyRow.xlsx")]
    public async Task QueryExcelDataReaderCheckTest(string path)
    {
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        await using var fs = File.OpenRead(path);
        using var reader = ExcelReaderFactory.CreateReader(fs);
        var exceldatareaderResult = reader.AsDataSet();
        await using var stream = File.OpenRead(path);

        var rows = await _excelImporter.QueryAsync(stream).ToListAsync();
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
    public async Task QuerySheetWithoutRAttribute()
    {
        var path = PathHelper.GetFile("xlsx/TestWihoutRAttribute.xlsx");
        await using var stream = File.OpenRead(path);

        var rows = await _excelImporter.QueryAsync(stream).Cast<IDictionary<string, object>>().ToListAsync();
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
        var path = PathHelper.GetFile("xlsx/TestDimensionC3.xlsx");
        await using var stream = File.OpenRead(path);

        var rows = await _excelImporter.QueryAsync(stream).ToListAsync();
        var keys = (rows.First() as IDictionary<string, object>)?.Keys;
        
        Assert.Equal(3, keys?.Count);
        Assert.Equal(2, rows.Count);
    }

    [Fact]
    public async Task QueryByLINQExtensionsVoidTaskLargeFileOOMTest()
    {
        const string path = "../../../../../benchmarks/MiniExcel.Benchmarks/Test1,000,000x10.xlsx";

        {
            var row = await _excelImporter.QueryAsync(path).FirstAsync();
            Assert.Equal("HelloWorld1", row.A);
        }

        await using (var stream = File.OpenRead(path))
        {
            var row = await _excelImporter.QueryAsync(stream).Cast<IDictionary<string, object>>().FirstAsync();
            Assert.Equal("HelloWorld1", row["A"]);
        }

        {
            var count = await _excelImporter.QueryAsync(path)
                .Cast<IDictionary<string, object>>()
                .Take(10)
                .CountAsync();

            Assert.Equal(10, count);
        }
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
        await _excelExporter.ExportAsync(path, values);

        await using var stream = File.OpenRead(path);
        var rows = await _excelImporter.QueryAsync(stream, hasHeaderRow: true).Cast<IDictionary<string, object>>().ToListAsync();

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
        await _excelExporter.ExportAsync(path, values);

        await using var stream = File.OpenRead(path);
        var rows = await _excelImporter.QueryAsync(stream, hasHeaderRow: true).Cast<IDictionary<string, object>>().ToListAsync();

        Assert.Equal("MiniExcel", rows[0]["Column1"]);
        Assert.Equal(1d, rows[0]["Column2"]);
        Assert.Equal("Github", rows[1]["Column1"]);
        Assert.Equal(2d, rows[1]["Column2"]);
    }

    [Fact]
    public async Task QueryRangeAsyncWithCellReferencesTest()
    {
        var path = PathHelper.GetFile("xlsx/TestTypeMapping.xlsx");

        await using var stream = File.OpenRead(path);
        var rows = await _excelImporter.QueryRangeAsync(stream, startCell: "B6", endCell: "D9")
            .Cast<IDictionary<string, object>>()
            .ToListAsync();

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
    public async Task QueryRangeAsyncWithIndexCoordinatesTest()
    {
        var path = PathHelper.GetFile("xlsx/TestTypeMapping.xlsx");

        await using var stream = File.OpenRead(path);
        var rows = await _excelImporter.QueryRangeAsync(stream, startRowIndex: 6, startColumnIndex: 2, endRowIndex: 9, endColumnIndex: 4)
            .Cast<IDictionary<string, object>>()
            .ToListAsync();

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
    public async Task QueryRangeAsyncWithCellReferencesAndHeaderTest()
    {
        var path = PathHelper.GetFile("xlsx/TestQueryRange.xlsx");
        var rows = await _excelImporter.QueryRangeAsync(path, hasHeaderRow: true, startCell: "C3", endCell: "E6")
            .Cast<IDictionary<string, object>>()
            .ToListAsync();

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
    public async Task QueryRangeAsyncWithIndexCoordinatesAndHeaderTest()
    {
        var path = PathHelper.GetFile("xlsx/TestQueryRange.xlsx");
        var rows = await _excelImporter.QueryRangeAsync(path, hasHeaderRow: true, startCell: "C3", endCell: "E6")
            .Cast<IDictionary<string, object>>()
            .ToListAsync();

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
    public async Task ReadBigExcel_TakeCancel_Throws_TaskCanceledException()
    {
        await Assert.ThrowsAsync<OperationCanceledException>(async () =>
        {
            var path = PathHelper.GetFile("xlsx/bigExcel.xlsx");
            using var cts = new CancellationTokenSource();

            await cts.CancelAsync();
            await using var stream = File.OpenRead(path);
            _ = await _excelImporter.QueryAsync(stream, cancellationToken: cts.Token).ToListAsync(cts.Token);
        });
    }

    [Fact]
    public async Task ReadBigExcel_Processing_TakeCancel_Throws_TaskCanceledException()
    {
        await Assert.ThrowsAsync<OperationCanceledException>(async () =>
        {
            var cts = new CancellationTokenSource();

            var exportTask = _excelImporter.QueryAsync(PathHelper.GetFile("xlsx/bigExcel.xlsx"), cancellationToken: cts.Token).ToListAsync(cts.Token);
            await cts.CancelAsync();
            await exportTask;
        });
    }

    [Fact]
    public async Task InvalidSheetNameCharactersShouldThrow()
    {
        await using var ms1 = new MemoryStream();
        await Assert.ThrowsAsync<ArgumentException>(() => _excelExporter.ExportAsync(ms1, Array.Empty<object>(), sheetName: "Sheet?"));
        
        await using var ms2 = new MemoryStream();
        await Assert.ThrowsAsync<ArgumentException>(() => _excelExporter.InsertSheetAsync(ms2, Array.Empty<object>(), sheetName: "Sheet[]"));
        
        await using var ms3 = new MemoryStream();
        using var package = new ExcelPackage(ms3);
        package.Workbook.Worksheets.Add("Sheet1");
        await package.SaveAsync();
        
        ms1.Seek(0, SeekOrigin.Begin);
        await Assert.ThrowsAsync<ArgumentException>(() => _excelExporter.AlterSheetAsync(ms3, "Sheet1", "Sheet*"));
    }

    [Fact]
    public async Task MultipleResultSets()
    {
        await using var stream =File.OpenRead(PathHelper.GetFile("xlsx/TestTypeMapping.xlsx"));
        await using var dr = await OpenXmlDataReader.CreateAsync(stream, hasHeaderRow: true, leaveOpen: true);
        await dr.ReadAsync();
        var v1 = dr.GetValue(0);
        var nr = await dr.NextResultAsync();
        await dr.ReadAsync();
        var v2 = dr.GetValue(0);
    }

    [Fact]
    public void MultipleResultSets2()
    {
        using var stream = File.OpenRead(PathHelper.GetFile("xlsx/TestMultiSheet.xlsx"));
        using var dr = OpenXmlDataReader.Create(stream, leaveOpen: true);
        dr.Read();
        var v1 = dr.GetValue(0);
        var nr = dr.NextResult();
        dr.Read();
        var v2 = dr.GetValue(0);
    }
}
