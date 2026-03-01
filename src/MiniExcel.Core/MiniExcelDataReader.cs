namespace MiniExcelLib.Core;

public sealed class MiniExcelDataReader : IMiniExcelDataReader
{
    private readonly IEnumerator<IDictionary<string, object?>>? _source;
    private readonly IAsyncEnumerator<IDictionary<string, object?>>? _asyncSource;
    private readonly Stream _stream;

    private List<string> _columns = [];
    private DataTable? _schema;

    private readonly bool _isAsyncSource;
    private bool _isFirst = true;

    public object this[int i]
        => GetValue(i);

    public object this[string name]
        => GetValue(GetOrdinal(name));

    public int Depth { get; private set; } = -1;
    public int FieldCount { get; private set; }
    public bool IsClosed { get; private set; }
    public int RecordsAffected => 0;

    
    private MiniExcelDataReader(Stream? stream, IEnumerable<IDictionary<string, object?>>? values)
    {
        _stream = stream ?? throw new ArgumentNullException(nameof(stream));
        _source = values?.GetEnumerator() ?? throw new ArgumentNullException(nameof(values));
    }
    
    public static MiniExcelDataReader Create(Stream? stream, IEnumerable<IDictionary<string, object?>> values)
    {
        var reader = new MiniExcelDataReader(stream, values);
        if (reader._source!.MoveNext())
        {
            reader._columns = reader._source.Current?.Keys.ToList() ?? [];
            reader.FieldCount = reader._columns.Count;
            reader.Depth++;
        }
        return reader;
    }

    private MiniExcelDataReader(Stream? stream, IAsyncEnumerable<IDictionary<string, object?>>? values)
    {
        _stream = stream ?? throw new ArgumentNullException(nameof(stream));
        _asyncSource = values?.GetAsyncEnumerator() ?? throw new ArgumentNullException(nameof(values));
        _isAsyncSource =  true;
    }

    public static async Task<MiniExcelDataReader> CreateAsync(Stream? stream, IAsyncEnumerable<IDictionary<string, object?>> values)
    {
        var reader = new MiniExcelDataReader(stream, values);
        if (await reader._asyncSource!.MoveNextAsync().ConfigureAwait(false))
        {
            reader._columns = reader._asyncSource.Current.Keys.ToList();
            reader.FieldCount = reader._columns.Count;
            reader.Depth++;
        }
        return reader;
    }

    public bool Read()
    {
        if (IsClosed)
            throw new InvalidOperationException("The data reader has been closed");
        
        if (_isAsyncSource)
            throw new InvalidOperationException("The data reader was configured to execute asynchronously");
        
        Depth++;
        if (_isFirst)
        {
            _isFirst = false;
            return true;
        }

        return _source!.MoveNext();
    }

    public async Task<bool> ReadAsync(CancellationToken cancellationToken = default)
    {
        if (!_isAsyncSource)
            return await Task.FromResult(Read()).ConfigureAwait(false);

        if (IsClosed)
            throw new InvalidOperationException("The data reader has been closed");

        Depth++;
        if (_isFirst)
        {
            _isFirst = false;
            return true;
        }
        
        return await _asyncSource!.MoveNextAsync().ConfigureAwait(false);
    }

    public IDataReader GetData(int i)
        => throw new NotSupportedException();

    public Type GetFieldType(int i)
        => typeof(object);
    
    public string GetDataTypeName(int i)
        => typeof(object).FullName!;

    /// <summary>
    /// This method will alway throw a <see cref="NotSupportedException" />
    /// </summary>
    public long GetBytes(int i, long fieldOffset, byte[]? buffer, int bufferoffset, int length) 
        => throw new NotSupportedException("MiniExcelDataReader does not support this method");

    public long GetChars(int i, long fieldoffset, char[]? buffer, int bufferoffset, int length)
    {
        var s = GetString(i).Substring((int)fieldoffset, length);

        for (int j = bufferoffset; j < s.Length; j++)
            buffer.AsSpan()[j] = s.AsSpan()[j];
        
        return s.Length;
    }
    
    public bool GetBoolean(int i) => GetValue(i) switch
    {
        bool b => b,
        null => throw new InvalidOperationException("The value is null"),
        var value => Convert.ToBoolean(value)
    };

    public byte GetByte(int i) => GetValue(i) is { } value
        ? Convert.ToByte(value)
        : throw new InvalidOperationException("The value is null");

    public char GetChar(int i) => GetValue(i) switch
    {
        char c => c,
        null => throw new InvalidOperationException("The value is null"),
        var value => Convert.ToChar(value)
    };

    public DateTime GetDateTime(int i) => GetValue(i) switch
    {
        DateTime dt => dt,
        double d => DateTime.FromOADate(d),
        null => throw new InvalidOperationException("The value is null"),
        var value => Convert.ToDateTime(value)
    };

    public decimal GetDecimal(int i) => GetValue(i) switch
    {
        decimal d => d,
        null => throw new InvalidOperationException("The value is null"),
        var value => Convert.ToDecimal(value)
    };

    public float GetFloat(int i) => GetValue(i) switch
    {
        float f => f,
        null => throw new InvalidOperationException("The value is null"),
        var value => Convert.ToSingle(value)
    };

    public double GetDouble(int i) => GetValue(i) switch
    {
        double d => d,
        null => throw new InvalidOperationException("The value is null"),
        var value => Convert.ToDouble(value)
    };

    public Guid GetGuid(int i) => GetValue(i) switch
    {
        Guid g => g,
        string s => Guid.Parse(s),
        byte[] b => new Guid(b),
        null => throw new InvalidOperationException("The value is null"),
        var value => throw new InvalidCastException($"The value {value} cannot be cast to Guid"),
    };

    public short GetInt16(int i) => GetValue(i) switch
    {
        short s => s,
        null => throw new InvalidOperationException("The value is null"),
        var value => Convert.ToInt16(value)
    };

    public int GetInt32(int i) => GetValue(i) switch
    {
        int s => s,
        null => throw new InvalidOperationException("The value is null"),
        var value => Convert.ToInt32(value)
    };

    public long GetInt64(int i) => GetValue(i) switch
    {
        long l => l,
        null => throw new InvalidOperationException("The value is null"),
        var value => Convert.ToInt64(value)
    };

    public string GetString(int i) => GetValue(i) switch
    {
        string s => s,
        var value => Convert.ToString(value) ?? ""
    };

    public object GetValue(int i)
    {
        var currentRow = _isAsyncSource
            ? _asyncSource!.Current
            : _source!.Current!;

        return currentRow is not null
            ? currentRow[_columns[i]] ?? DBNull.Value
            : throw new InvalidOperationException("Current row is not available.");
    }

    public int GetValues(object?[] values)
    {
        var count = Math.Min(values.Length, FieldCount);

        for (int i = 0; i < values.Length; i++) 
            values[i] = GetValue(i);

        return count;
    }

    public bool IsDBNull(int i) 
        => GetValue(i) is null or DBNull;
    
    public string GetName(int i)
        => _columns[i];

    public int GetOrdinal(string name)
        => _columns.IndexOf(name);
    
    public DataTable GetSchemaTable()
    {
        if (_schema is null)
        {
            _schema = new DataTable();
            _schema.Columns.Add("ColumnOrdinal");
            _schema.Columns.Add("ColumnName");
            
            for (int i = 0; i < _columns.Count; i++)
            {
                _schema.Rows.Add(i, _columns[i]);
            }
        }
        
        return _schema;
    }

    /// <summary>
    /// This method will alway return false
    /// </summary>
    public bool NextResult() => false;

    public void Close()
    {
        if (IsClosed) 
            return;

        if (_isAsyncSource)
        {
            if (_asyncSource is IDisposable disposable)
                disposable.Dispose();
            else 
                _asyncSource!.DisposeAsync(); // fire and forget if all other options are exhausted
        }
        else
        {
            _source!.Dispose();
        }

        _stream.Dispose();
        IsClosed = true;
    }

    public async Task CloseAsync()
    {
        if (IsClosed)
            return;

        if (_isAsyncSource)
        {
            await _asyncSource!.DisposeAsync().ConfigureAwait(false);
        }
        _source?.Dispose();

#if NETCOREAPP3_0_OR_GREATER
        await _stream!.DisposeAsync().ConfigureAwait(false);
#else
        _stream.Dispose();
#endif

        IsClosed = true;
    }

    public void Dispose()
        => Close();

    public async ValueTask DisposeAsync()
        => await CloseAsync().ConfigureAwait(false);
}
