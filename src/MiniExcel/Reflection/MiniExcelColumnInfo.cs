namespace MiniExcelLib.Reflection;

public class MiniExcelColumnInfo
{
    public object Key { get; set; }
    public int? ExcelColumnIndex { get; set; }
    public string? ExcelColumnName { get; set; }
    public string[]? ExcelColumnAliases { get; set; } = [];
    public MiniExcelProperty Property { get; set; }
    public Type ExcludeNullableType { get; set; }
    public bool Nullable { get; internal set; }
    public string? ExcelFormat { get; internal set; }
    public double? ExcelColumnWidth { get; internal set; }
    public string? ExcelIndexName { get; internal set; }
    public bool ExcelIgnore { get; internal set; }
    public int ExcelFormatId { get; internal set; }
    public ColumnType ExcelColumnType { get; internal set; }
    public Func<object, object>? CustomFormatter { get; set; }
}