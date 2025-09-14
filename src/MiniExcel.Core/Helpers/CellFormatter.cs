namespace MiniExcelLib.Core.Helpers;

/// <summary>
/// Utility class for formatting cell values consistently across the mapping system.
/// Centralizes Excel-specific formatting logic to reduce code duplication.
/// </summary>
internal static class CellFormatter
{
    /// <summary>
    /// Excel epoch date used for date/time calculations.
    /// Excel treats dates as days since this date.
    /// </summary>
    public static readonly DateTime ExcelEpoch = new(1899, 12, 30);

    /// <summary>
    /// Formats a value for Excel cell output, returning both the formatted string and cell type.
    /// </summary>
    /// <param name="value">The value to format</param>
    /// <returns>A tuple containing the formatted value and the Excel cell type</returns>
    public static (string? value, string? type) FormatCellValue(object? value)
    {
        if (value is null)
            return (null, null);
        
        switch (value)
        {
            case string s:
                // Use inline string to avoid shared string table
                return (s, "inlineStr");
                
            case DateTime dt:
                // Excel stores dates as numbers
                var excelDate = (dt - ExcelEpoch).TotalDays;
                return (excelDate.ToString(CultureInfo.InvariantCulture), null);
                
            case DateTimeOffset dto:
                var excelDateOffset = (dto.DateTime - ExcelEpoch).TotalDays;
                return (excelDateOffset.ToString(CultureInfo.InvariantCulture), null);
                
            case bool b:
                return (b ? "1" : "0", "b");
                
            case byte:
            case sbyte:
            case short:
            case ushort:
            case int:
            case uint:
            case long:
            case ulong:
            case float:
            case double:
            case decimal:
                return (Convert.ToString(value, CultureInfo.InvariantCulture), null);
                
            default:
                // Convert to string
                return (value.ToString(), "inlineStr");
        }
    }

}