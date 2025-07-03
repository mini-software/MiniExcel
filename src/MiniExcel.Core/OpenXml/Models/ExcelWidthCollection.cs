using MiniExcelLib.Core.Reflection;

namespace MiniExcelLib.Core.OpenXml.Models;

public sealed class ExcelColumnWidth
{
    public int Index { get; set; }
    public double Width { get; set; }

    internal static IEnumerable<ExcelColumnWidth> FromProps(ICollection<MiniExcelColumnInfo> props, double? minWidth = null)
    {
        var i = 1;
        foreach (var p in props)
        {
            if (p?.ExcelColumnWidth is not null || minWidth is not null)
            {
                var colIndex = p?.ExcelColumnIndex + 1;
                yield return new ExcelColumnWidth
                {
                    Index = colIndex ?? i,
                    Width = p?.ExcelColumnWidth ?? minWidth.Value
                };
            }

            i++;
        }
    }
}

public sealed class ExcelWidthCollection
{
    private readonly Dictionary<int, ExcelColumnWidth> _columnWidths;
    private readonly double _maxWidth;

    public IEnumerable<ExcelColumnWidth> Columns => _columnWidths.Values;

    internal ExcelWidthCollection(double minWidth, double maxWidth, ICollection<MiniExcelColumnInfo> props)
    {
        _maxWidth = maxWidth;
        _columnWidths = ExcelColumnWidth.FromProps(props, minWidth).ToDictionary(x => x.Index);
    }

    public void AdjustWidth(int columnIndex, string columnValue)
    {
        if (!string.IsNullOrEmpty(columnValue) && _columnWidths.TryGetValue(columnIndex, out var currentWidth))
        {
            var adjustedWidth = Math.Max(currentWidth.Width, GetApproximateTextWidth(columnValue.Length));
            currentWidth.Width = Math.Min(_maxWidth, adjustedWidth);
        }
    }

    /// <summary>
    /// Get the approximate width of the given text for Calibri 11pt
    /// </summary>
    /// <remarks>
    /// Rounds the result to 2 decimal places.
    /// </remarks>
    public static double GetApproximateTextWidth(int textLength)
    {
        const double characterWidthFactor = 1.2;  // Estimated factor for Calibri, 11pt
        const double padding = 2;  // Add some padding for extra spacing

        var excelColumnWidth = textLength * characterWidthFactor + padding;
        return Math.Round(excelColumnWidth, 2);
    }
}