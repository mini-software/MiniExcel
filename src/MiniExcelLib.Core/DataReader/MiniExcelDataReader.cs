namespace MiniExcelLib.Core.DataReader;

public class MiniExcelDataReader : MiniExcelDataReaderBase
{
    private readonly IEnumerator<IDictionary<string, object?>> _source;
    private readonly Stream _stream;
    private readonly List<string> _keys;
    private readonly int _fieldCount;

    private bool _isFirst = true;
    private bool _disposed = false;

    /// <summary>
    /// Initializes a new instance of the <see cref="MiniExcelDataReader"/> class.
    /// </summary>
    internal MiniExcelDataReader(Stream? stream, IEnumerable<IDictionary<string, object?>> values)
    {
        _stream = stream ?? throw new ArgumentNullException(nameof(stream));
        _source = values.GetEnumerator();
        
        if (_source.MoveNext())
        {
            _keys = _source.Current?.Keys.ToList() ?? [];
            _fieldCount = _keys.Count;
        }
    }

    public static MiniExcelDataReader Create(Stream? stream, IEnumerable<IDictionary<string, object?>> values) => new(stream, values);
    
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
    /// <inheritdoc/>
    /// </summary>
    /// <returns></returns>
    public override bool Read()
    {
        if (_isFirst)
        {
            _isFirst = false;
            return true;
        }
        return _source.MoveNext();
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
                _source?.Dispose();
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
}