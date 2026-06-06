using MiniExcelLib.Tests.Common.Utils;

namespace MiniExcelLib.OpenXml.Tests.DataReader;

public class OpenXmlDataReaderAsyncTests
{
    private readonly OpenXmlImporter _excelImporter = MiniExcel.Importers.GetOpenXmlImporter();

    static OpenXmlDataReaderAsyncTests()
    {
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
    }

    [Fact]
    public async Task GetDataReader_WithSimpleData_ReturnsValidDataReader()
    {
        var path = PathHelper.GetFile("xlsx/TestTypeMapping.xlsx");
        await using var stream = File.OpenRead(path);
        await using var reader = await _excelImporter.GetAsyncDataReader(stream, hasHeaderRow: true);
        
        Assert.Equal(8, reader.FieldCount);
        Assert.Equal("ID", reader.GetName(0));
        Assert.Equal("Name", reader.GetName(1));
        Assert.Equal("BoD", reader.GetName(2));
        Assert.Equal("Age", reader.GetName(3));
        Assert.Equal("VIP", reader.GetName(4));
        Assert.Equal("Mail", reader.GetName(5));
        Assert.Equal("Points", reader.GetName(6));
        
        Assert.True(await reader.ReadAsync());
        Assert.Equal(Guid.Parse("78DE23D2-DCB6-BD3D-EC67-C112BBC322A2"), reader.GetGuid(0));
        Assert.Equal("Wade", reader.GetString(1));
        Assert.Equal(new DateTime(2020, 9, 27), reader.GetDateTime(2));
        Assert.Equal(36, reader.GetInt32(3));
        Assert.False(reader.GetBoolean(4));
        Assert.Equal(5019.12, reader.GetDouble(6));
    }

    [Fact]
    public async Task GetDataReader_WithoutHeaderRow_ReturnsDataWithoutHeaders()
    {
        var path = PathHelper.GetFile("xlsx/TestStrictOpenXml.xlsx");
        await using var stream = File.OpenRead(path);
        await using var reader = await _excelImporter.GetAsyncDataReader(stream, hasHeaderRow: false);
        
        // First row should be headers when hasHeaderRow is false
        Assert.True(await reader.ReadAsync());
        Assert.Equal("A", reader.GetName(0));
        Assert.Equal("B", reader.GetName(1));
        Assert.Equal("C", reader.GetName(2));
    }

    [Fact]
    public async Task GetDataReader_WithSpecificSheet_ReadsCorrectSheet()
    {
        var path = PathHelper.GetFile("xlsx/TestMultiSheet.xlsx");
        await using var stream = File.OpenRead(path);
        await using var reader = await _excelImporter.GetAsyncDataReader(stream, hasHeaderRow: true, sheetName: "Sheet2");
        
        Assert.Equal("Sheet2", reader.GetWorksheetName());
        Assert.True(await reader.ReadAsync());
        Assert.Equal(1d, reader.GetValue(0));
    }

    [Fact]
    public async Task GetDataReader_WithStartCell_SkipsToStartingCell()
    {
        var path = PathHelper.GetFile("xlsx/TestTypeMapping.xlsx");
        await using var stream = File.OpenRead(path);
        await using var reader = await _excelImporter.GetAsyncDataReader(stream, hasHeaderRow: false, startCell: "C3");

        Assert.True(await reader.ReadAsync());
        Assert.Equal(new DateTime(2020, 10, 25), reader.GetDateTime(0));
        Assert.Equal(44, reader.GetInt32(1));
        Assert.True(reader.GetBoolean(2));
        Assert.Equal("elit.elit.fermentum@enim.edu", reader.GetString(3));
        Assert.Equal(7028.46, reader.GetDouble(4));
    }

    [Fact]
    public async Task GetDataReader_WithLeaveOpenFalse_DisposesStream()
    {
        var path = PathHelper.GetFile("xlsx/TestTypeMapping.xlsx");

        var stream = File.OpenRead(path);
        await using (var reader = await _excelImporter.GetAsyncDataReader(stream, hasHeaderRow: false, leaveOpen: false))
        {
            await reader.ReadAsync();
        }

        Assert.False(stream.CanRead);
        Assert.Throws<ObjectDisposedException>(() => stream.Seek(0, SeekOrigin.Begin));
    }

    [Fact]
    public async Task GetDataReader_WithLeaveOpenTrue_DoesNotDisposeStream()
    {
        var path = PathHelper.GetFile("xlsx/TestTypeMapping.xlsx");

        var stream = File.OpenRead(path);
        await using (var reader = await _excelImporter.GetAsyncDataReader(stream, hasHeaderRow: false, leaveOpen: true))
        {
            await reader.ReadAsync();
        }

        Assert.True(stream.CanRead);
    }

    [Fact]
    public async Task GetDataReader_GetSchemaTable_ReturnsColumnInfo()
    {
        var path = PathHelper.GetFile("xlsx/TestTypeMapping.xlsx");
        await using var reader = await _excelImporter.GetAsyncDataReader(path, hasHeaderRow: true);
        using var schemaTable = reader.GetSchemaTable();
        
        Assert.NotNull(schemaTable);
        Assert.Equal(8, schemaTable.Rows.Count);
    }

    [Fact]
    public async Task GetDataReader_GetOrdinal_ReturnsColumnIndex()
    {
        var path = PathHelper.GetFile("xlsx/TestTypeMapping.xlsx");
        await using var reader = await _excelImporter.GetAsyncDataReader(path, hasHeaderRow: true);

        Assert.Equal(0, reader.GetOrdinal("ID"));
        Assert.Equal(1, reader.GetOrdinal("Name"));
        Assert.Equal(2, reader.GetOrdinal("BoD"));
        Assert.Equal(3, reader.GetOrdinal("Age"));
        Assert.Equal(4, reader.GetOrdinal("VIP"));
        Assert.Equal(5, reader.GetOrdinal("Mail"));
        Assert.Equal(6, reader.GetOrdinal("Points"));
    }

    [Fact]
    public async Task GetDataReader_WithMultipleSheets_ReadsAllSheets()
    {
        var path = PathHelper.GetFile("xlsx/TestMultiSheet.xlsx");
        await using var reader = await _excelImporter.GetAsyncDataReader(path);

        Assert.Equal("Sheet1", reader.GetWorksheetName());
        Assert.True(await reader.ReadAsync());
        Assert.Equal(2d, reader.GetValue(0));
        
        // Move to next result set
        Assert.True(await reader.NextResultAsync());
        Assert.Equal("Sheet2", reader.GetWorksheetName());
        Assert.True(await reader.ReadAsync());
        Assert.Equal(1d, reader.GetValue(0));
        Assert.True(await reader.NextResultAsync());
        Assert.Equal("Sheet3", reader.GetWorksheetName());
        Assert.True(await reader.ReadAsync());
        Assert.Equal(3d, reader.GetValue(0));
        Assert.False(await reader.NextResultAsync());
    }
}
