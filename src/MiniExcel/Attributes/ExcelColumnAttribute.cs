using System;
using MiniExcelLibs.Utils;

namespace MiniExcelLibs.Attributes;

[AttributeUsage(AttributeTargets.Property | AttributeTargets.Field)]
public class ExcelColumnAttribute : Attribute
{
    private int _index = -1;
    private string? _xName;

    internal int FormatId { get; set; } = -1;

    public string? Name { get; set; }
    public string[]? Aliases { get; set; } = [];
    public double Width { get; set; } = 9.28515625;
    public string? Format { get; set; }
    public bool Ignore { get; set; }
    public ColumnType Type { get; set; } = ColumnType.Value;

    public int Index
    {
        get => _index;
        set => Init(value);
    }

    public string? IndexName
    {
        get => _xName;
        set => Init(ColumnHelper.GetColumnIndex(value), value);
    }

    private void Init(int index, string? columnName = null)
    {
        if (index < 0)
            throw new ArgumentOutOfRangeException(nameof(index), index, $"Column index {index} must be greater or equal to zero.");

        _index = index;
        _xName ??= columnName ?? ColumnHelper.GetAlphabetColumnName(index);
    }
}

public enum ColumnType { Value, Formula }

public class DynamicExcelColumn : ExcelColumnAttribute
{
    public string Key { get; set; }
    public Func<object, object> CustomFormatter { get; set; }

    public DynamicExcelColumn(string key)
    {
        Key = key;
    }
}