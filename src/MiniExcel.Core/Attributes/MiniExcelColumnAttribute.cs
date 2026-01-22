namespace MiniExcelLib.Core.Attributes;

[AttributeUsage(AttributeTargets.Property | AttributeTargets.Field)]
public class MiniExcelColumnAttribute : Attribute
{
    private int _index = -1;
    private string? _xName;

    public string? Name { get; set; }
     public string[]? Aliases { get; set; } = [];
     public string? Format { get; set; }
     public bool Ignore { get; set; }
    
    internal int FormatId { get; private set; } = -1;
    public double Width { get; set; } = 9.28515625;
    public ColumnType Type { get; set; } = ColumnType.Value;

    public int Index
    {
        get => _index;
        set => Init(value);
    }

    public string? IndexName
    {
        get => _xName;
        set => Init(CellReferenceConverter.GetNumericalIndex(value), value);
    }

    private void Init(int index, string? columnName = null)
    {
        if (index < 0)
            throw new ArgumentOutOfRangeException(nameof(index), index, $"Column index {index} must be greater or equal to zero.");

        _index = index;
        _xName ??= columnName ?? CellReferenceConverter.GetAlphabeticalIndex(index);
    }
    
    public void SetFormatId(int formatId) =>  FormatId = formatId;
}

public enum ColumnType { Value, Formula }

public class DynamicExcelColumn : MiniExcelColumnAttribute
{
    public string Key { get; set; }
    public Func<object, object>? CustomFormatter { get; set; }

    public DynamicExcelColumn(string key)
    {
        Key = key;
    }
}
