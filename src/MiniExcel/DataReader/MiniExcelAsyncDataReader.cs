namespace MiniExcelLib.DataReader;

// todo: this is way improvable, ideally the sync and async implementations into a single datareader
public class MiniExcelAsyncDataReader : MiniExcelDataReaderBase, IAsyncDisposable
{
    private readonly IAsyncEnumerator<IDictionary<string, object?>> _source;
    
    private readonly Stream _stream;
    private List<string> _keys;
    private int _fieldCount;

    private bool _isFirst = true;
    private bool _disposed = false;

    /// <summary>
    /// Initializes a new instance of the <see cref="MiniExcelDataReader"/> class.
    /// </summary>

    internal MiniExcelAsyncDataReader(Stream? stream, IAsyncEnumerable<IDictionary<string, object?>> values)
    {
        _stream = stream ?? throw new ArgumentNullException(nameof(stream));
        _source = values.GetAsyncEnumerator();
    }

    public static async Task<MiniExcelAsyncDataReader> CreateAsync(Stream? stream, IAsyncEnumerable<IDictionary<string, object?>> values)
    {
        var reader = new MiniExcelAsyncDataReader(stream, values);
        if (await reader._source.MoveNextAsync().ConfigureAwait(false))
        {
            reader._keys = reader._source.Current?.Keys.ToList() ?? [];
            reader._fieldCount = reader._keys.Count;
        }
        return reader;
    }


    /// <inheritdoc/>
    public override object? GetValue(int i)
    {
        if (_source.Current is null)
            throw new InvalidOperationException("No current row available.");
        
        return _source.Current[_keys[i]];
    }

    /// <inheritdoc/>
    public override int FieldCount => _fieldCount;

    /// <summary>
    /// This method will throw a NotSupportedException. Please use ReadAsync or the synchronous MiniExcelDataReader implementation.
    /// </summary>
    public override bool Read() => throw new NotSupportedException("Use the ReadAsync method instead.");

    public override async Task<bool> ReadAsync(CancellationToken cancellationToken = default)
    {
        if (_isFirst)
        {
            _isFirst = false;
            return await Task.FromResult(true).ConfigureAwait(false);
        }
        
        return await _source.MoveNextAsync().ConfigureAwait(false);
    }

    /// <inheritdoc/>
    public override string GetName(int i)
    {
        return _keys[i];
    }

    /// <inheritdoc/>
    public override int GetOrdinal(string name)
    {
        return _keys.IndexOf(name);
    }

    /// <inheritdoc/>
    protected override void Dispose(bool disposing)
    {
        if (!_disposed)
        {
            if (disposing)
            {
                _stream?.Dispose();
                _source.DisposeAsync().GetAwaiter().GetResult();
            }
            _disposed = true;
        }
        base.Dispose(disposing);
    }

    /// <summary>
    /// Disposes the object.
    /// </summary>
    public new void Dispose()
    {
        Dispose(true);
        GC.SuppressFinalize(this);
    }

    public async ValueTask DisposeAsync()
    {
        _stream?.Dispose();
        await _source.DisposeAsync().ConfigureAwait(false);
        
        GC.SuppressFinalize(this);
    }
}