namespace MiniExcelLib.Core;

public abstract class MiniExcelDataReaderBase : IMiniExcelDataReader
{
    protected readonly IMiniExcelReader MiniExcelReader;

    protected IEnumerator<IDictionary<string, object?>>? Source;
    protected IAsyncEnumerator<IDictionary<string, object?>>? AsyncSource;
    protected readonly bool IsAsyncSource;

    protected readonly Dictionary<string, int> Ordinals = [];
    
    protected bool IsEmpty;
    protected bool HasHeaderRow;

    protected List<string> Columns = [];
    protected DataTable? Schema;

    protected bool IsFirstRow = true;

    public virtual object this[int i]
        => GetValue(i);

    public virtual object this[string name]
        => GetValue(GetOrdinal(name));

    public int Depth => 0;
    public int RecordsAffected => -1;
    public int FieldCount { get; protected set; }

    private bool _disposed;
    public bool IsClosed { get; protected set; }


    protected MiniExcelDataReaderBase(IMiniExcelReader miniExcelReader, bool hasHeaderRow, bool isAsyncSource)
    {
        MiniExcelReader = miniExcelReader;
        HasHeaderRow = hasHeaderRow;
        IsAsyncSource =  isAsyncSource;
    }

    public virtual bool Read()
    {
        if (_disposed)
            throw new ObjectDisposedException("This data reader has been disposed.");
            
        if (IsClosed)
            throw new InvalidOperationException("This data reader has been closed.");
        
        if (IsAsyncSource)
            throw new InvalidOperationException("This data reader was configured to execute asynchronously");
        
        if (IsFirstRow)
        {
            IsFirstRow = false;
            return !IsEmpty;
        }

        return Source!.MoveNext();
    }

    public virtual async Task<bool> ReadAsync(CancellationToken cancellationToken = default)
    {
        if (_disposed)
            throw new ObjectDisposedException("This data reader has been disposed.");

        if (IsClosed)
            throw new InvalidOperationException("The data reader has been closed");

        if (!IsAsyncSource)
            return await Task.FromResult(Read()).ConfigureAwait(false);

        if (IsFirstRow)
        {
            IsFirstRow = false;
            return !IsEmpty;
        }
        
        return await AsyncSource!.MoveNextAsync().ConfigureAwait(false);
    }

    public virtual IDataReader GetData(int i)
        => throw new NotSupportedException();

    public virtual Type GetFieldType(int i)
        => typeof(object);
    
    public virtual string GetDataTypeName(int i)
        => typeof(object).FullName!;

    /// <summary>
    /// This method will alway throw a <see cref="NotSupportedException" />
    /// </summary>
    public virtual long GetBytes(int i, long fieldOffset, byte[]? buffer, int bufferoffset, int length) 
        => throw new NotSupportedException("MiniExcelDataReader does not support this method");

    public virtual long GetChars(int i, long fieldoffset, char[]? buffer, int bufferoffset, int length)
    {
        var s = GetString(i);
        var len = Math.Min(length, s.Length - (int)fieldoffset);
        var subs = s.Substring((int)fieldoffset, len);

        if (buffer is not null)
            subs.AsSpan().CopyTo(buffer.AsSpan(bufferoffset));

        return subs.Length;
    }
    
    public virtual bool GetBoolean(int i) => GetValue(i) switch
    {
        bool b => b,
        null => throw new InvalidOperationException("The value is null"),
        var value => Convert.ToBoolean(value)
    };

    public virtual byte GetByte(int i) => GetValue(i) is { } value
        ? Convert.ToByte(value)
        : throw new InvalidOperationException("The value is null");

    public virtual char GetChar(int i) => GetValue(i) switch
    {
        char c => c,
        null => throw new InvalidOperationException("The value is null"),
        var value => Convert.ToChar(value)
    };

    public virtual DateTime GetDateTime(int i) => GetValue(i) switch
    {
        DateTime dt => dt,
        double d => DateTime.FromOADate(d),
        null => throw new InvalidOperationException("The value is null"),
        var value => Convert.ToDateTime(value)
    };

    public virtual decimal GetDecimal(int i) => GetValue(i) switch
    {
        decimal d => d,
        null => throw new InvalidOperationException("The value is null"),
        var value => Convert.ToDecimal(value)
    };

    public virtual float GetFloat(int i) => GetValue(i) switch
    {
        float f => f,
        null => throw new InvalidOperationException("The value is null"),
        var value => Convert.ToSingle(value)
    };

    public virtual double GetDouble(int i) => GetValue(i) switch
    {
        double d => d,
        null => throw new InvalidOperationException("The value is null"),
        var value => Convert.ToDouble(value)
    };

    public virtual Guid GetGuid(int i) => GetValue(i) switch
    {
        Guid g => g,
        string s => Guid.Parse(s),
        byte[] b => new Guid(b),
        null => throw new InvalidOperationException("The value is null"),
        var value => throw new InvalidCastException($"The value {value} cannot be cast to Guid"),
    };

    public virtual short GetInt16(int i) => GetValue(i) switch
    {
        short s => s,
        null => throw new InvalidOperationException("The value is null"),
        var value => Convert.ToInt16(value)
    };

    public virtual int GetInt32(int i) => GetValue(i) switch
    {
        int s => s,
        null => throw new InvalidOperationException("The value is null"),
        var value => Convert.ToInt32(value)
    };

    public virtual long GetInt64(int i) => GetValue(i) switch
    {
        long l => l,
        null => throw new InvalidOperationException("The value is null"),
        var value => Convert.ToInt64(value)
    };

    public virtual string GetString(int i) => GetValue(i) switch
    {
        string s => s,
        var value => Convert.ToString(value) ?? ""
    };

    public virtual object GetValue(int i)
    {
        var currentRow = IsAsyncSource
            ? AsyncSource?.Current
            : Source?.Current;

        return currentRow is not null
            ? currentRow[Columns[i]] ?? DBNull.Value
            : throw new InvalidOperationException("Current row is not available.");
    }

    public virtual int GetValues(object?[] values)
    {
        var count = Math.Min(values.Length, FieldCount);

        for (int i = 0; i < count; i++) 
            values[i] = GetValue(i);

        return count;
    }

    public virtual bool IsDBNull(int i) 
        => GetValue(i) is null or DBNull;
    
    public virtual string GetName(int i)
        => Columns[i];

    public int GetOrdinal(string name)
    {
        if (name is null)
            throw new ArgumentNullException(nameof(name));

        if (Ordinals.TryGetValue(name, out var ordinal))
            return ordinal;

        var ord = Columns.IndexOf(name);
        Ordinals[name] = ord;
            
        return ord;
    }

    public DataTable GetSchemaTable()
    {
        if (Schema is null)
        {
            Schema = new DataTable();
            Schema.Columns.Add("ColumnOrdinal", typeof(int));
            Schema.Columns.Add("ColumnName", typeof(string));
            
            for (int i = 0; i < Columns.Count; i++)
            {
                Schema.Rows.Add(i, Columns[i]);
            }
        }
        
        return Schema;
    }
    
    public virtual bool NextResult()
        => throw new NotImplementedException();

    public virtual Task<bool> NextResultAsync(CancellationToken cancellationToken = default)
        => throw new NotImplementedException();


    public void Close() 
        => Dispose();

    public async Task CloseAsync()
        => await DisposeAsync().ConfigureAwait(false);

    protected virtual void Dispose(bool disposing)
    {
        if (_disposed)
            return;

        if (disposing)
        {
            if (IsAsyncSource)
            {
                if (AsyncSource is IDisposable disposable) disposable.Dispose();
                // necessary fallback when the data reader is being disposed synchronously despite having being initialized asynchronously  
                else Task.Run(async () => await AsyncSource!.DisposeAsync().ConfigureAwait(false)).GetAwaiter().GetResult();
            }
            else
            {
                Source!.Dispose();
            }

            MiniExcelReader.Dispose();
            Schema?.Dispose();

            IsClosed = true;
        }

        _disposed = true;
    }

    public void Dispose()
    {
        Dispose(true);
        GC.SuppressFinalize(this);
    }

    protected virtual async ValueTask DisposeAsyncCore()
    {
        if (_disposed)
            return;

        if (IsAsyncSource) 
            await AsyncSource!.DisposeAsync().ConfigureAwait(false);

        Schema?.Dispose();
        await MiniExcelReader.DisposeAsync().ConfigureAwait(false);
        
        IsClosed = true;
    }

    public async ValueTask DisposeAsync()
    {
        await DisposeAsyncCore().ConfigureAwait(false);
        Dispose(false);

        GC.SuppressFinalize(this);
    }
}
