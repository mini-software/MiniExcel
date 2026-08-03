using System.Resources;

namespace MiniExcelLib.Core.Attributes;

public class MiniExcelColumnAttribute : MiniExcelAttributeBase
{
    private ResourceManager? _resourceManager;

    public string? Name { get; set; }
    public string[]? Aliases { get; set; } = [];
    public string? Format { get; set; }
    public bool Hidden { get; set; }
    public bool Ignore { get; set; }

    internal int FormatId { get; private set; } = -1;
    public double Width { get; set; } = 8.42857143;
    public ColumnType Type { get; set; } = ColumnType.Value;

    private int _index = -1;
    public int Index
    {
        get => _index;
        set => Init(value);
    }

    private string? _indexName;
    public string? IndexName
    {
        get => _indexName;
        set => Init(CellReferenceConverter.GetNumericalIndex(value), value);
    }

    private Type? _resourceType;
    public Type? ResourceType
    {
        get => _resourceType;
        set
        {
            if (_resourceType == value)
                return;

            _resourceType = value;
            if (value is null)
                return;

            const BindingFlags bindingFlags = BindingFlags.Static | BindingFlags.Public | BindingFlags.NonPublic;
            if (value.GetProperty(nameof(ResourceManager), bindingFlags) is { } property && 
                property.GetValue(null) is ResourceManager resourceManager)
            {
                _resourceManager = resourceManager;
            }
            else
            {
                _resourceManager = new ResourceManager(value);
            }
        }
    }

    internal string? GetColumnName(string? resourceKey = null)
    {
        if (Name is not null)
            return _resourceManager?.GetString(Name) ?? Name;

        if (resourceKey is not null)
            return _resourceManager?.GetString(resourceKey) ?? resourceKey;

        return null;
    }

    private void Init(int index, string? columnName = null)
    {
        if (index < 0)
            throw new ArgumentOutOfRangeException(nameof(index), index, $"Column index {index} must be greater or equal to zero.");

        _index = index;
        _indexName ??= columnName ?? CellReferenceConverter.GetAlphabeticalIndex(index);
    }

    public void SetFormatId(int formatId) => FormatId = formatId;
}

public class DynamicExcelColumn(string key) : MiniExcelColumnAttribute
{
    public string Key { get; set; } = key;
    public Func<object?, object?>? CustomFormatter { get; set; }
}

public enum ColumnType { Value, Formula }
