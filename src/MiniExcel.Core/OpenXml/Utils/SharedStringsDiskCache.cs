namespace MiniExcelLib.Core.OpenXml.Utils;

internal class SharedStringsDiskCache : IDictionary<int, string>, IDisposable
{
    private const int ExcelCellMaxLength = 32767;
    private static readonly Encoding Encoding = new UTF8Encoding(true);
        
    private readonly FileStream _positionFs;
    private readonly FileStream _lengthFs;
    private readonly FileStream _valueFs;
    
    private long _maxIndex = -1;
    private bool _disposedValue;
    
    public int Count => checked((int)(_maxIndex + 1));

    public string this[int key]
    {
        get => GetValue(key);
        set => Add(key, value);
    }
    
    public bool ContainsKey(int key) => key <= _maxIndex;

    public SharedStringsDiskCache(string sharedStringsCachePath)
    {
        var dir = Path.GetDirectoryName(sharedStringsCachePath);
        if (!Directory.Exists(dir))
            throw new DirectoryNotFoundException($"\"{dir}\" is not a valid path for the shared strings cache.");

        var prefix = $"{Path.GetRandomFileName()}_miniexcel";
        _positionFs = new FileStream(Path.Combine(dir, $"{prefix}_position"), FileMode.OpenOrCreate);
        _lengthFs = new FileStream(Path.Combine(dir, $"{prefix}_length"), FileMode.OpenOrCreate);
        _valueFs = new FileStream(Path.Combine(dir, $"{prefix}_data"), FileMode.OpenOrCreate);
    }

    // index must start with 0-N
    private void Add(int index, string value)
    {
        if (index > _maxIndex)
            _maxIndex = index;
            
        var valueBs = Encoding.GetBytes(value);
        if (value.Length > ExcelCellMaxLength) //check info length, becasue cell string max length is 47483647
            throw new ArgumentOutOfRangeException("", "Excel one cell max length is 32,767 characters");
            
        _positionFs.Write(BitConverter.GetBytes(_valueFs.Position), 0, 4);
        _lengthFs.Write(BitConverter.GetBytes(valueBs.Length), 0, 4);
        _valueFs.Write(valueBs, 0, valueBs.Length);
    }

    private string GetValue(int index)
    {
        _positionFs.Position = index * 4;
        var bytes = new byte[4];
        _ = _positionFs.Read(bytes, 0, 4);
        var position = BitConverter.ToInt32(bytes, 0);
            
        bytes = new byte[4];
        _lengthFs.Position = index * 4;
        _ = _lengthFs.Read(bytes, 0, 4);
        var length = BitConverter.ToInt32(bytes, 0);
            
        bytes = new byte[length];
        _valueFs.Position = position;
        _ = _valueFs.Read(bytes, 0, length);

        return Encoding.GetString(bytes);
    }

    public ICollection<int> Keys => throw new NotImplementedException();
    public ICollection<string> Values => throw new NotImplementedException();
    public bool IsReadOnly => throw new NotImplementedException();
    public bool Remove(int key)
    {
        throw new NotImplementedException();
    }

    public bool TryGetValue(int key, out string value)
    {
        throw new NotImplementedException();
    }

    public void Add(KeyValuePair<int, string> item)
    {
        throw new NotImplementedException();
    }

    public void Clear()
    {
        throw new NotImplementedException();
    }

    public bool Contains(KeyValuePair<int, string> item)
    {
        throw new NotImplementedException();
    }

    public void CopyTo(KeyValuePair<int, string>[] array, int arrayIndex)
    {
        throw new NotImplementedException();
    }

    public bool Remove(KeyValuePair<int, string> item)
    {
        throw new NotImplementedException();
    }

    public IEnumerator<KeyValuePair<int, string>> GetEnumerator()
    {
        for (int i = 0; i < _maxIndex; i++)
            yield return new KeyValuePair<int, string>(i, this[i]);
    }

    IEnumerator IEnumerable.GetEnumerator()
    {
        for (int i = 0; i < _maxIndex; i++)
            yield return this[i];
    }

    void IDictionary<int, string>.Add(int key, string value)
    {
        throw new NotImplementedException();
    }
    
    
    ~SharedStringsDiskCache()
    {
        Dispose(disposing: false);
    }

    public void Dispose()
    {
        Dispose(disposing: true);
        GC.SuppressFinalize(this);
    }
    
    protected virtual void Dispose(bool disposing)
    {
        if (!_disposedValue)
        {
            if (disposing)
            {
                // TODO: dispose managed state (managed objects)
            }
                
            _positionFs.Dispose();
            if (File.Exists(_positionFs.Name))
                File.Delete(_positionFs.Name);
                
            _lengthFs.Dispose();
            if (File.Exists(_lengthFs.Name))
                File.Delete(_lengthFs.Name);
                
            _valueFs.Dispose();
            if (File.Exists(_valueFs.Name))
                File.Delete(_valueFs.Name);
                
            _disposedValue = true;
        }
    }
}