namespace MiniExcelLib.Csv.Tests.DataReader;

public class CsvDataReaderAsyncTests
{
    private readonly CsvImporter _csvImporter = MiniExcel.Importers.GetCsvImporter();

    [Fact]
    public async Task GetDataReader_WithSimpleData_ReturnsValidDataReader()
    {
        var path = PathHelper.GetFile("csv/TestDataReaderHeader.csv");
        await using var stream = File.OpenRead(path);
        await using var reader = await _csvImporter.GetAsyncDataReader(stream, hasHeaderRow: true);
        
        Assert.Equal("Name", reader.GetName(0));
        Assert.Equal("Age", reader.GetName(1));
        Assert.True(await reader.ReadAsync());
        Assert.Equal("John", reader.GetString(0));
        Assert.Equal(30, reader.GetInt32(1));
        Assert.True(await reader.ReadAsync());
        Assert.Equal("Jane", reader.GetString(0));
        Assert.Equal(25, reader.GetInt32(1));
        Assert.False(await reader.ReadAsync());
    }

    [Fact]
    public async Task GetDataReader_WithoutHeaderRow_ReturnsDataWithoutHeaders()
    {
        var path = PathHelper.GetFile("csv/TestDataReaderNoHeader.csv");
        await using var stream = File.OpenRead(path);
        await using var reader = await _csvImporter.GetAsyncDataReader(stream, hasHeaderRow: false);

        Assert.Equal("A", reader.GetName(0));
        Assert.Equal("B", reader.GetName(1));
        Assert.True(await reader.ReadAsync());
        Assert.Equal("Value1", reader.GetValue(0));
        Assert.Equal("Value2", reader.GetValue(1));
        Assert.True(await reader.ReadAsync());
        Assert.Equal("Value3", reader.GetValue(0));
        Assert.Equal("Value4", reader.GetValue(1));
        Assert.False(await reader.ReadAsync());
    }

    [Fact]
    public async Task GetDataReader_WithCustomSeparator_AppliesConfiguration()
    {
        var path = PathHelper.GetFile("csv/TestDataReaderCustomSeparator.csv");
        await using var stream = File.OpenRead(path);

        var importConfig = new CsvConfiguration { Seperator = ';' };
        await using var reader = await _csvImporter.GetAsyncDataReader(stream, hasHeaderRow: true, configuration: importConfig);
        
        Assert.Equal("Name", reader.GetName(0));
        Assert.Equal("Age", reader.GetName(1));
        Assert.True(await reader.ReadAsync());
        Assert.Equal("John", reader.GetValue(0));
        Assert.Equal("30", reader.GetValue(1));
        Assert.True(await reader.ReadAsync());
        Assert.Equal("Jane", reader.GetValue(0));
        Assert.Equal("25", reader.GetValue(1));
        Assert.False(await reader.ReadAsync());
    }

    [Fact]
    public async Task GetDataReader_WithLeaveOpenFalse_DisposesStream()
    {
        var path = PathHelper.GetFile("csv/TestDataReaderHeader.csv");
        var stream = File.OpenRead(path);
        await using (var reader = await _csvImporter.GetAsyncDataReader(stream, hasHeaderRow: false, leaveOpen: false))
        {
            await reader.ReadAsync();
        }

        Assert.False(stream.CanRead);
        await Assert.ThrowsAsync<ObjectDisposedException>(() => stream.ReadAsync([], 0, 0));
    }

    [Fact]
    public async Task GetDataReader_WithLeaveOpenTrue_DoesNotDisposeStream()
    {
        var path = PathHelper.GetFile("csv/TestDataReaderHeader.csv");
        await using var stream = File.OpenRead(path);
        await using (var reader = await _csvImporter.GetAsyncDataReader(stream, hasHeaderRow: false, leaveOpen: true))
        {
            await reader.ReadAsync();
        }

        Assert.True(stream.CanRead);
    }

    [Fact]
    public async Task GetDataReader_GetSchemaTable_ReturnsColumnInfo()
    {
        var path = PathHelper.GetFile("csv/TestDataReaderHeader.csv");
        await using var stream = File.OpenRead(path);

        await using var reader = await _csvImporter.GetAsyncDataReader(stream, hasHeaderRow: true);
        using var schemaTable = reader.GetSchemaTable();
        
        Assert.NotNull(schemaTable);
        Assert.Equal(2, schemaTable.Rows.Count);
    }

    [Fact]
    public async Task GetDataReader_GetOrdinal_ReturnsColumnIndex()
    {
        var path = PathHelper.GetFile("csv/TestDataReaderHeader.csv");
        await using var stream = File.OpenRead(path);
        await using var reader = await _csvImporter.GetAsyncDataReader(stream, hasHeaderRow: true);

        Assert.Equal(0, reader.GetOrdinal("Name"));
        Assert.Equal(1, reader.GetOrdinal("Age"));
    }

    [Fact]
    public async Task GetDataReader_NextResult_ThrowsNotSupportedException()
    {
        var path = PathHelper.GetFile("csv/TestDataReaderHeader.csv");
        await using var stream = File.OpenRead(path);
        await using var reader = await _csvImporter.GetAsyncDataReader(stream, hasHeaderRow: true);
        
        await Assert.ThrowsAsync<NotSupportedException>(async () => await reader.NextResultAsync());
    }
}
