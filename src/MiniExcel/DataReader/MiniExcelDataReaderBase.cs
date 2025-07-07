namespace MiniExcelLib.DataReader;

/// <summary>
/// IMiniExcelDataReader Base Class
/// </summary>
public abstract class MiniExcelDataReaderBase : IMiniExcelDataReader
{
    /// <summary>
    /// <inheritdoc/>
    /// </summary>
    /// <param name="i"></param>
    /// <returns></returns>
    public virtual object? this[int i] => null;

    /// <summary>
    /// <inheritdoc/>
    /// </summary>
    /// <param name="name"></param>
    /// <returns></returns>
    public virtual object? this[string name] => null;

    /// <summary>
    /// <inheritdoc/>
    /// </summary>
    public virtual int Depth { get; } = 0;

    /// <summary>
    /// <inheritdoc/>
    /// </summary>
    public virtual bool IsClosed { get; } = false;

    /// <summary>
    /// <inheritdoc/>
    /// </summary>
    public virtual int RecordsAffected { get; } = 0;

    /// <summary>
    /// <inheritdoc/>
    /// </summary>
    public virtual int FieldCount { get; } = 0;

    /// <summary>
    /// <inheritdoc/>
    /// </summary>
    /// <param name="i"></param>
    /// <returns></returns>
    public virtual bool GetBoolean(int i) => false;

    /// <summary>
    /// <inheritdoc/>
    /// </summary>
    /// <param name="i"></param>
    /// <returns></returns>
    public virtual byte GetByte(int i) => byte.MinValue;

    /// <summary>
    /// <inheritdoc/>
    /// </summary>
    /// <param name="i"></param>
    /// <param name="fieldOffset"></param>
    /// <param name="buffer"></param>
    /// <param name="bufferOffset"></param>
    /// <param name="length"></param>
    /// <returns></returns>
    public virtual long GetBytes(int i, long fieldOffset, byte[]? buffer, int bufferOffset, int length) => 0;

    /// <summary>
    /// <inheritdoc/>
    /// </summary>
    /// <param name="i"></param>
    /// <returns></returns>
    public virtual char GetChar(int i) => char.MinValue;

    /// <summary>
    /// <inheritdoc/>
    /// </summary>
    /// <param name="i"></param>
    /// <param name="fieldOffset"></param>
    /// <param name="buffer"></param>
    /// <param name="bufferOffset"></param>
    /// <param name="length"></param>
    /// <returns></returns>
    public virtual long GetChars(int i, long fieldOffset, char[]? buffer, int bufferOffset, int length) => 0;

    /// <summary>
    /// <inheritdoc/>
    /// </summary>
    /// <param name="i"></param>
    /// <returns></returns>
    public virtual IDataReader? GetData(int i) => null;

    /// <summary>
    /// <inheritdoc/>
    /// </summary>
    /// <param name="i"></param>
    /// <returns></returns>
    public virtual string GetDataTypeName(int i) => string.Empty;

    /// <summary>
    /// <inheritdoc/>
    /// </summary>
    /// <param name="i"></param>
    /// <returns></returns>
    public virtual DateTime GetDateTime(int i) => DateTime.MinValue;

    /// <summary>
    /// <inheritdoc/>
    /// </summary>
    /// <param name="i"></param>
    /// <returns></returns>
    public virtual decimal GetDecimal(int i) => 0;

    /// <summary>
    /// <inheritdoc/>
    /// </summary>
    /// <param name="i"></param>
    /// <returns></returns>
    public virtual double GetDouble(int i) => 0;

    /// <summary>
    /// <inheritdoc/>
    /// </summary>
    /// <param name="i"></param>
    /// <returns></returns>
    public virtual Type GetFieldType(int i) => null!;

    /// <summary>
    /// <inheritdoc/>
    /// </summary>
    /// <param name="i"></param>
    /// <returns></returns>
    public virtual float GetFloat(int i) => 0f;

    /// <summary>
    /// <inheritdoc/>
    /// </summary>
    /// <param name="i"></param>
    /// <returns></returns>
    public virtual Guid GetGuid(int i) => Guid.Empty;

    /// <summary>
    /// <inheritdoc/>
    /// </summary>
    /// <param name="i"></param>
    /// <returns></returns>
    public virtual short GetInt16(int i) => 0;

    /// <summary>
    /// <inheritdoc/>
    /// </summary>
    /// <param name="i"></param>
    /// <returns></returns>
    public virtual int GetInt32(int i) => 0;

    /// <summary>
    /// <inheritdoc/>
    /// </summary>
    /// <param name="i"></param>
    /// <returns></returns>
    public virtual long GetInt64(int i) => 0;

    /// <summary>
    /// <inheritdoc/>
    /// </summary>
    /// <param name="name"></param>
    /// <returns></returns>
    public virtual int GetOrdinal(string name) => 0;

    /// <summary>
    /// <inheritdoc/>
    /// </summary>
    /// <returns></returns>
    public virtual DataTable? GetSchemaTable() => null;

    /// <summary>
    /// <inheritdoc/>
    /// </summary>
    /// <param name="i"></param>
    /// <returns></returns>
    public virtual string GetString(int i) => string.Empty;

    /// <summary>
    /// <inheritdoc/>
    /// </summary>
    /// <param name="values"></param>
    /// <returns></returns>
    public virtual int GetValues(object[] values) => 0;

    /// <summary>
    /// <inheritdoc/>
    /// </summary>
    /// <param name="i"></param>
    /// <returns></returns>
    public virtual bool IsDBNull(int i) => false;

    /// <summary>
    /// <inheritdoc/>
    /// </summary>
    /// <returns></returns>
    public virtual bool NextResult() => false;

    /// <summary>
    /// <inheritdoc/>
    /// </summary>
    /// <param name="cancellationToken"></param>
    /// <returns></returns>
    public virtual Task<bool> NextResultAsync(CancellationToken cancellationToken = default)
    {
        if (cancellationToken.IsCancellationRequested)
            return Task.FromCanceled<bool>(cancellationToken);

        try
        {
            return NextResult() ? Task.FromResult(true) : Task.FromResult(false);
        }
        catch (Exception e)
        {
            return Task.FromException<bool>(e);
        }
    }

    /// <summary>
    /// <inheritdoc/>
    /// </summary>
    /// <param name="i"></param>
    /// <returns></returns>
    public abstract string GetName(int i);

    /// <summary>
    /// <inheritdoc/>
    /// </summary>
    /// <param name="i"></param>
    /// <param name="cancellationToken"></param>
    /// <returns></returns>
    public virtual Task<string> GetNameAsync(int i, CancellationToken cancellationToken = default)
    {
        if (cancellationToken.IsCancellationRequested)
            return Task.FromCanceled<string>(cancellationToken);
        
        try
        {
            return Task.FromResult(GetName(i));
        }
        catch (Exception e)
        {
            return Task.FromException<string>(e);
        }
    }

    /// <summary>
    /// <inheritdoc/>
    /// </summary>
    /// <param name="i"></param>
    /// <returns></returns>
    public abstract object? GetValue(int i);

    /// <summary>
    /// <inheritdoc/>
    /// </summary>
    /// <param name="i"></param>
    /// <param name="cancellationToken"></param>
    /// <returns></returns>
    public virtual Task<object?> GetValueAsync(int i, CancellationToken cancellationToken = default)
    {
        if (cancellationToken.IsCancellationRequested)
            return Task.FromCanceled<object?>(cancellationToken);

        try
        {
            return Task.FromResult(GetValue(i));
        }
        catch (Exception e)
        {
            return Task.FromException<object?>(e);
        }
    }

    /// <summary>
    /// <inheritdoc/>
    /// </summary>
    /// <returns></returns>
    public abstract bool Read();

    /// <summary>
    /// <inheritdoc/>
    /// </summary>
    /// <param name="cancellationToken"></param>
    /// <returns></returns>
    public virtual Task<bool> ReadAsync(CancellationToken cancellationToken = default)
    {
        if (cancellationToken.IsCancellationRequested)
            return Task.FromCanceled<bool>(cancellationToken);

        try
        {
            return Read() ? Task.FromResult(true) : Task.FromResult(false);
        }
        catch (Exception e)
        {
            return Task.FromException<bool>(e);
        }
    }

    /// <summary>
    /// <inheritdoc/>
    /// </summary>
    public virtual void Close()
    {
    }

    /// <summary>
    /// <inheritdoc/>
    /// </summary>
    /// <returns></returns>
    public virtual Task CloseAsync()
    {
        try
        {
            Close();
            return Task.CompletedTask;
        }
        catch (Exception e)
        {
            return Task.FromException(e);
        }
    }

    /// <summary>
    /// <inheritdoc/>
    /// </summary>
    public void Dispose()
    {
        Dispose(true);
        GC.SuppressFinalize(this);
    }

#if NET8_0_OR_GREATER
    /// <summary>
    /// <inheritdoc/>
    /// </summary>
    /// <returns></returns>
    /// <exception cref="NotImplementedException"></exception>
    public virtual ValueTask DisposeAsync()
    {
        Dispose();
        return default;
    }
#endif

    /// <summary>
    /// <inheritdoc cref="Dispose" />
    /// </summary>
    /// <param name="disposing"></param>
    protected virtual void Dispose(bool disposing)
    {
        if (disposing)
        {
            Close();
        }
    }
}