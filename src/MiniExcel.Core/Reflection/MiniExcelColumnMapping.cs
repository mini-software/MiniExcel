using MiniExcelLib.Core.Attributes;

namespace MiniExcelLib.Core.Reflection;

public class MiniExcelColumnMapping
{
    public object Key { get; set; }
    public MiniExcelMemberAccessor MemberAccessor { get; set; }
    public Type ExcludeNullableType { get; set; }
    public bool Nullable { get; internal set; }
    public int? ExcelColumnIndex { get; set; }
    public string? ExcelColumnName { get; set; }
    public string[]? ExcelColumnAliases { get; set; } = [];
    public string? ExcelFormat { get; internal set; }
    public double? ExcelColumnWidth { get; internal set; }
    public string? ExcelIndexName { get; internal set; }
    public bool ExcelHiddenColumn { get; internal set; }
    public bool ExcelIgnoreColumn { get; internal set; }
    public int ExcelFormatId { get; internal set; }
    public ColumnType ExcelColumnType { get; internal set; }
    public Func<object?, object?>? CustomFormatter { get; set; }
}