namespace MiniExcelLib.OpenXml.Utils;

internal sealed class SharedStringsDiskCache : IDictionary<int, string>, IDisposable
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

    public SharedStringsDiskCache(string sharedStringsCacheDir)
    {
        if (string.IsNullOrWhiteSpace(sharedStringsCacheDir) || !Directory.Exists(sharedStringsCacheDir))
            throw new DirectoryNotFoundException($"\"{sharedStringsCacheDir}\" is not a valid directory for the shared strings cache.");

        var prefix = $"{Path.GetRandomFileName()}_miniexcel";
        _positionFs = new FileStream(Path.Combine(sharedStringsCacheDir, $"{prefix}_position"), FileMode.OpenOrCreate);
        _lengthFs = new FileStream(Path.Combine(sharedStringsCacheDir, $"{prefix}_length"), FileMode.OpenOrCreate);
        _valueFs = new FileStream(Path.Combine(sharedStringsCacheDir, $"{prefix}_data"), FileMode.OpenOrCreate);
    }

    // index must start with 0-N
    public void Add(int index, string value)
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
        if (index > _maxIndex)
            throw new KeyNotFoundException();
        
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

    public ICollection<int> Keys => throw new NotSupportedException();
    public ICollection<string> Values => throw new NotSupportedException();
    public bool IsReadOnly => throw new NotImplementedException();
    public bool Remove(int key)
    {
        throw new NotSupportedException();
    }

    public bool TryGetValue(int key, out string value)
    {
        throw new NotSupportedException();
    }

    public void Add(KeyValuePair<int, string> item) => Add(item.Key, item.Value);

    public void Clear()
    {
        throw new NotSupportedException();
    }

    public bool Contains(KeyValuePair<int, string> item)
    {
        throw new NotSupportedException();
    }

    public void CopyTo(KeyValuePair<int, string>[] array, int arrayIndex)
    {
        throw new NotSupportedException();
    }

    public bool Remove(KeyValuePair<int, string> item)
    {
        throw new NotSupportedException();
    }

    public IEnumerator<KeyValuePair<int, string>> GetEnumerator()
    {
        for (int i = 0; i <= _maxIndex; i++)
            yield return new KeyValuePair<int, string>(i, this[i]);
    }

    IEnumerator IEnumerable.GetEnumerator() => GetEnumerator();

    public void Dispose()
    {
        if (_disposedValue)
            return;

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
