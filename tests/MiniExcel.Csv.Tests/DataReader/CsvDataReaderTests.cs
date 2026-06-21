namespace MiniExcelLib.Csv.Tests.DataReader;

public class CsvDataReaderTests
{
    private readonly CsvImporter _csvImporter = MiniExcel.Importers.GetCsvImporter();

    [Fact]
    public void GetDataReader_WithSimpleData_ReturnsValidDataReader()
    {
        var path = PathHelper.GetFile("csv/TestDataReaderHeader.csv");
        using var stream = File.OpenRead(path);
        using var reader = _csvImporter.GetDataReader(stream, hasHeaderRow: true);
        
        Assert.Equal("Name", reader.GetName(0));
        Assert.Equal("Age", reader.GetName(1));
        Assert.True(reader.Read());
        Assert.Equal("John", reader.GetString(0));
        Assert.Equal(30, reader.GetInt32(1));
        Assert.True(reader.Read());
        Assert.Equal("Jane", reader.GetString(0));
        Assert.Equal(25, reader.GetInt32(1));
        Assert.False(reader.Read());
    }

    [Fact]
    public void GetDataReader_WithoutHeaderRow_ReturnsDataWithoutHeaders()
    {
        var path = PathHelper.GetFile("csv/TestDataReaderNoHeader.csv");
        using var stream = File.OpenRead(path);
        using var reader = _csvImporter.GetDataReader(stream, hasHeaderRow: false);

        Assert.Equal("A", reader.GetName(0));
        Assert.Equal("B", reader.GetName(1));
        Assert.True(reader.Read());
        Assert.Equal("Value1", reader.GetValue(0));
        Assert.Equal("Value2", reader.GetValue(1));
        Assert.True(reader.Read());
        Assert.Equal("Value3", reader.GetValue(0));
        Assert.Equal("Value4", reader.GetValue(1));
        Assert.False(reader.Read());
    }

    [Fact]
    public void GetDataReader_WithCustomSeparator_AppliesConfiguration()
    {
        var path = PathHelper.GetFile("csv/TestDataReaderCustomSeparator.csv");
        using var stream = File.OpenRead(path);

        var importConfig = new CsvConfiguration { Seperator = ';' };
        using var reader = _csvImporter.GetDataReader(stream, hasHeaderRow: true, configuration: importConfig);
        
        Assert.Equal("Name", reader.GetName(0));
        Assert.Equal("Age", reader.GetName(1));
        Assert.True(reader.Read());
        Assert.Equal("John", reader.GetValue(0));
        Assert.Equal("30", reader.GetValue(1));
        Assert.True(reader.Read());
        Assert.Equal("Jane", reader.GetValue(0));
        Assert.Equal("25", reader.GetValue(1));
        Assert.False(reader.Read());
    }

    [Fact]
    public void GetDataReader_WithLeaveOpenFalse_DisposesStream()
    {
        var path = PathHelper.GetFile("csv/TestDataReaderHeader.csv");
        var stream = File.OpenRead(path);
        using (var reader = _csvImporter.GetDataReader(stream, hasHeaderRow: false, leaveOpen: false))
        {
            reader.Read();
        }

        Assert.False(stream.CanRead);
        Assert.Throws<ObjectDisposedException>(() => stream.Seek(0, SeekOrigin.Begin));
    }

    [Fact]
    public void GetDataReader_WithLeaveOpenTrue_DoesNotDisposeStream()
    {
        var path = PathHelper.GetFile("csv/TestDataReaderHeader.csv");
        using var stream = File.OpenRead(path);
        using (var reader = _csvImporter.GetDataReader(stream, hasHeaderRow: false, leaveOpen: true))
        {
            reader.Read();
        }

        Assert.True(stream.CanRead);
    }

    [Fact]
    public void GetDataReader_GetSchemaTable_ReturnsColumnInfo()
    {
        var path = PathHelper.GetFile("csv/TestDataReaderHeader.csv");
        using var stream = File.OpenRead(path);

        using var reader = _csvImporter.GetDataReader(stream, hasHeaderRow: true);
        using var schemaTable = reader.GetSchemaTable();
        
        Assert.NotNull(schemaTable);
        Assert.Equal(2, schemaTable.Rows.Count);
    }

    [Fact]
    public void GetDataReader_GetOrdinal_ReturnsColumnIndex()
    {
        var path = PathHelper.GetFile("csv/TestDataReaderHeader.csv");
        using var stream = File.OpenRead(path);
        using var reader = _csvImporter.GetDataReader(stream, hasHeaderRow: true);

        Assert.Equal(0, reader.GetOrdinal("Name"));
        Assert.Equal(1, reader.GetOrdinal("Age"));
    }

    [Fact]
    public void GetDataReader_NextResult_ThrowsNotSupportedException()
    {
        var path = PathHelper.GetFile("csv/TestDataReaderHeader.csv");
        using var reader = _csvImporter.GetDataReader(path, hasHeaderRow: true);
        
        Assert.Throws<NotSupportedException>(() => reader.NextResult());
    }
}
