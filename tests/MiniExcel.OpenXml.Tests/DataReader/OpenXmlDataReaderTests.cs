using MiniExcelLib.Tests.Common.Utils;

namespace MiniExcelLib.OpenXml.Tests.DataReader;

public class OpenXmlDataReaderTests
{
    private readonly OpenXmlImporter _excelImporter = MiniExcel.Importers.GetOpenXmlImporter();

    static OpenXmlDataReaderTests()
    {
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
    }

    [Fact]
    public async Task GetDataReader_WithSimpleData_ReturnsValidDataReader()
    {
        var path = PathHelper.GetFile("xlsx/TestTypeMapping.xlsx");
        using var stream = File.OpenRead(path);
        using var reader = _excelImporter.GetDataReader(stream, hasHeaderRow: true);
        
        Assert.Equal(8, reader.FieldCount);
        Assert.Equal("ID", reader.GetName(0));
        Assert.Equal("Name", reader.GetName(1));
        Assert.Equal("BoD", reader.GetName(2));
        Assert.Equal("Age", reader.GetName(3));
        Assert.Equal("VIP", reader.GetName(4));
        Assert.Equal("Mail", reader.GetName(5));
        Assert.Equal("Points", reader.GetName(6));
        
        Assert.True(reader.Read());
        Assert.Equal(Guid.Parse("78DE23D2-DCB6-BD3D-EC67-C112BBC322A2"), reader.GetGuid(0));
        Assert.False(reader.IsDBNull(1));
        Assert.Equal("Wade", reader.GetString(1));
        Assert.Equal(new DateTime(2020, 9, 27), reader.GetDateTime(2));
        Assert.Equal(36, reader.GetInt16(3));
        Assert.Equal(36, reader.GetInt32(3));
        Assert.Equal(36, reader.GetInt64(3));
        Assert.False(reader.GetBoolean(4));
        Assert.Equal(5019.12f, reader.GetFloat(6));
        Assert.Equal(5019.12d, reader.GetDouble(6));
        Assert.Equal(5019.12m, reader.GetDecimal(6));
    }

    [Fact]
    public void GetDataReader_WithoutHeaderRow_ReturnsDataWithoutHeaders()
    {
        var path = PathHelper.GetFile("xlsx/TestStrictOpenXml.xlsx");
        using var stream = File.OpenRead(path);
        using var reader = _excelImporter.GetDataReader(stream, hasHeaderRow: false);
        
        // First row should be headers when hasHeaderRow is false
        Assert.True(reader.Read());
        Assert.Equal("A", reader.GetName(0));
        Assert.Equal("B", reader.GetName(1));
        Assert.Equal("C", reader.GetName(2));
    }

    [Fact]
    public void GetDataReader_WithSpecificSheet_ReadsCorrectSheet()
    {
        var path = PathHelper.GetFile("xlsx/TestMultiSheet.xlsx");
        using var stream = File.OpenRead(path);
        using var reader = _excelImporter.GetDataReader(stream, hasHeaderRow: true, sheetName: "Sheet2");
        
        Assert.Equal("Sheet2", reader.GetWorksheetName());
        Assert.True(reader.Read());
        Assert.Equal(1d, reader.GetValue(0));
    }

    [Fact]
    public void GetDataReader_WithStartCell_SkipsToStartingCell()
    {
        var path = PathHelper.GetFile("xlsx/TestTypeMapping.xlsx");
        using var stream = File.OpenRead(path);
        using var reader = _excelImporter.GetDataReader(stream, hasHeaderRow: false, startCell: "C3");

        Assert.True(reader.Read());
        Assert.Equal(new DateTime(2020, 10, 25), reader.GetDateTime(0));
        Assert.Equal(44, reader.GetInt32(1));
        Assert.True(reader.GetBoolean(2));
        Assert.Equal("elit.elit.fermentum@enim.edu", reader.GetString(3));
        Assert.Equal(7028.46, reader.GetDouble(4));
    }

    [Fact]
    public void GetDataReader_WithLeaveOpenFalse_DisposesStream()
    {
        var path = PathHelper.GetFile("xlsx/TestTypeMapping.xlsx");

        var stream = File.OpenRead(path);
        using (var reader = _excelImporter.GetDataReader(stream, hasHeaderRow: false, leaveOpen: false))
        {
            reader.Read();
        }

        Assert.False(stream.CanRead);
        Assert.Throws<ObjectDisposedException>(() => stream.Seek(0, SeekOrigin.Begin));
    }

    [Fact]
    public void GetDataReader_WithLeaveOpenTrue_DoesNotDisposeStream()
    {
        var path = PathHelper.GetFile("xlsx/TestTypeMapping.xlsx");

        var stream = File.OpenRead(path);
        using (var reader = _excelImporter.GetDataReader(stream, hasHeaderRow: false, leaveOpen: true))
        {
            reader.Read();
        }

        Assert.True(stream.CanRead);
    }

    [Fact]
    public void GetDataReader_GetSchemaTable_ReturnsColumnInfo()
    {
        var path = PathHelper.GetFile("xlsx/TestTypeMapping.xlsx");
        using var reader = _excelImporter.GetDataReader(path, hasHeaderRow: true);
        using var schemaTable = reader.GetSchemaTable();
        
        Assert.NotNull(schemaTable);
        Assert.Equal(8, schemaTable.Rows.Count);
    }

    [Fact]
    public void GetDataReader_GetOrdinal_ReturnsColumnIndex()
    {
        var path = PathHelper.GetFile("xlsx/TestTypeMapping.xlsx");
        using var reader = _excelImporter.GetDataReader(path, hasHeaderRow: true);

        Assert.Equal(0, reader.GetOrdinal("ID"));
        Assert.Equal(1, reader.GetOrdinal("Name"));
        Assert.Equal(2, reader.GetOrdinal("BoD"));
        Assert.Equal(3, reader.GetOrdinal("Age"));
        Assert.Equal(4, reader.GetOrdinal("VIP"));
        Assert.Equal(5, reader.GetOrdinal("Mail"));
        Assert.Equal(6, reader.GetOrdinal("Points"));
    }

    [Fact]
    public void GetDataReader_WithMultipleSheets_ReadsAllSheets()
    {
        var path = PathHelper.GetFile("xlsx/TestMultiSheet.xlsx");
        using var reader = _excelImporter.GetDataReader(path);

        Assert.Equal("Sheet1", reader.GetWorksheetName());
        Assert.True(reader.Read());
        Assert.Equal(2d, reader.GetValue(0));
        
        // Move to next result set
        Assert.True(reader.NextResult());
        Assert.Equal("Sheet2", reader.GetWorksheetName());
        Assert.True(reader.Read());
        Assert.Equal(1d, reader.GetValue(0));
        Assert.True(reader.NextResult());
        Assert.Equal("Sheet3", reader.GetWorksheetName());
        Assert.True(reader.Read());
        Assert.Equal(3d, reader.GetValue(0));
        Assert.False(reader.NextResult());
    }

    [Fact]
    public void GetDataReader_ReadingAfterDispose_ThrowsException()
    {
        var path = PathHelper.GetFile("xlsx/TestMultiSheet.xlsx");
        var reader = _excelImporter.GetDataReader(path);

        reader.Dispose();
        Assert.Throws<ObjectDisposedException>(() => reader.Read());
    }
}
