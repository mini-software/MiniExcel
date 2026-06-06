using MiniExcelLib.Core;

namespace MiniExcelLib.Csv;

public sealed class CsvDataReader : MiniExcelDataReaderBase
{
    private CsvDataReader(CsvReader reader, bool hasHeaderRow, bool isAsyncSource, CsvConfiguration configuration)
        : base(reader, hasHeaderRow, isAsyncSource, configuration)
    {
    }
    
    internal static CsvDataReader Create(Stream stream, bool hasHeaderRow = false, CsvConfiguration? configuration = null, bool leaveOpen = false)
    {
        CsvReader? reader = null;
        CsvDataReader? dataReader = null;
        configuration ??= CsvConfiguration.Default;

        try
        {
            reader = new CsvReader(stream, configuration, leaveOpen);
            dataReader = new CsvDataReader(reader, hasHeaderRow, isAsyncSource: false, configuration)
            {
                Source = reader.Query(hasHeaderRow, null, "A1").GetEnumerator(),
            };

            if (dataReader.Source!.MoveNext())
            {
                dataReader.Columns = dataReader.Source.Current?.Keys.ToList() ?? [];
                dataReader.FieldCount = dataReader.Columns.Count;
            }
            else
            {
                dataReader.IsEmpty = true;
            }

            var result = dataReader;
            dataReader = null;
            reader = null;
            stream = null!;

            return result;
        }
        finally
        {
            dataReader?.Dispose();
            reader?.Dispose();

            if (!leaveOpen)
                ((Stream?)stream)?.Dispose();
        }
    }

    internal static async Task<CsvDataReader> CreateAsync(Stream stream, bool hasHeaderRow = false, CsvConfiguration? configuration = null, bool leaveOpen = false, CancellationToken cancellationToken = default)
    {
        CsvReader? reader = null;
        CsvDataReader? dataReader = null;
        configuration ??= CsvConfiguration.Default;

        try
        {
            reader = new CsvReader(stream, configuration, leaveOpen);
            dataReader = new CsvDataReader(reader, hasHeaderRow, isAsyncSource: true, configuration)
            {
                AsyncSource = reader.QueryAsync(hasHeaderRow, null, "A1", cancellationToken).GetAsyncEnumerator(cancellationToken)
            };

            if (await dataReader.AsyncSource.MoveNextAsync().ConfigureAwait(false))
            {
                dataReader.Columns = dataReader.AsyncSource.Current?.Keys.ToList() ?? [];
                dataReader.FieldCount = dataReader.Columns.Count;
            }
            else
            {
                dataReader.IsEmpty = true;
            }

            var result = dataReader;
            dataReader = null;
            reader = null;
            stream = null!;
            
            return result;
        }
        finally
        {
            if (dataReader is not null)
                await dataReader.DisposeAsync().ConfigureAwait(false);
            
            if (reader?.DisposeAsync() is { } disposeTask)
                await disposeTask.ConfigureAwait(false);
            
            if (!leaveOpen && (Stream?)stream is not null)
                await stream.DisposeAsync().ConfigureAwait(false);
        }
    }

    /// <summary>
    /// This method will throw <see cref="NotSupportedException" />
    /// </summary>
    public override bool NextResult() 
        => throw new NotSupportedException();

    /// <summary>
    /// This method will throw <see cref="NotSupportedException" />
    /// </summary>
    public override Task<bool> NextResultAsync(CancellationToken cancellationToken = default)
        => throw new NotSupportedException();
}
