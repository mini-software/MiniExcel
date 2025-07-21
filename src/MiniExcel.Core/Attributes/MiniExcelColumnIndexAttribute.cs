using MiniExcelLib.Core.Helpers;

namespace MiniExcelLib.Core.Attributes;

[AttributeUsage(AttributeTargets.Property | AttributeTargets.Field)]
public class MiniExcelColumnIndexAttribute : Attribute
{
    public int ExcelColumnIndex { get; set; }
    internal string? ExcelXName { get; set; }
    public MiniExcelColumnIndexAttribute(string columnName) => Init(ColumnHelper.GetColumnIndex(columnName), columnName);
    public MiniExcelColumnIndexAttribute(int columnIndex) => Init(columnIndex);

    private void Init(int columnIndex, string? columnName = null)
    {
        if (columnIndex < 0)
            throw new ArgumentOutOfRangeException(nameof(columnIndex), columnIndex, $"Column index {columnIndex} must be greater or equal to zero.");
        
        ExcelXName ??= columnName ?? ColumnHelper.GetAlphabetColumnName(columnIndex);
        ExcelColumnIndex = columnIndex;
    }
}