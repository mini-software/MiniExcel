namespace MiniExcelLib.OpenXml.Charts;

/// <summary>
/// The values or categories of a chart series: either a range of cells or literal values stored in the chart itself.
/// </summary>
public sealed class OpenXmlChartData
{
    private OpenXmlChartData(string? formula, double[]? numbers, string[]? strings)
    {
        Formula = formula;
        Numbers = numbers;
        Strings = strings;
    }

    internal string? Formula { get; }
    internal double[]? Numbers { get; }
    internal string[]? Strings { get; }

    /// <summary>
    /// Data read from a range of cells, such as <c>B2:B100</c>, on a sheet of the same workbook.
    /// </summary>
    public static OpenXmlChartData FromRange(string sheetName, string range)
    {
        if (string.IsNullOrEmpty(range))
            throw new ArgumentException("The range must not be empty.", nameof(range));

        var cells = range.Split(':');
        if (cells.Length > 2)
            throw new ArgumentException($"Value {range} is not a valid cell range.", nameof(range));

        if (!CellReferenceConverter.TryParseCellReference(cells[0], out var firstColumn, out var firstRow) ||
            !CellReferenceConverter.TryParseCellReference(cells[^1], out var lastColumn, out var lastRow))
            throw new ArgumentException($"Value {range} is not a valid cell range.", nameof(range));

        return FromRange(sheetName, firstRow, firstColumn, lastRow, lastColumn);
    }

    /// <summary>
    /// Data read from a range of cells on a sheet of the same workbook. Rows and columns are 1-based.
    /// </summary>
    public static OpenXmlChartData FromRange(string sheetName, int firstRow, int firstColumn, int lastRow, int lastColumn)
    {
        if (string.IsNullOrEmpty(sheetName))
            throw new ArgumentException("The sheet name must not be empty.", nameof(sheetName));
        if (firstRow < 1 || firstColumn < 1 || lastRow < firstRow || lastColumn < firstColumn)
            throw new ArgumentOutOfRangeException(nameof(lastRow), "The range must start at row and column 1 or later and must not end before it starts.");

        var first = AbsoluteReference(firstColumn, firstRow);
        var last = AbsoluteReference(lastColumn, lastRow);
        var range = first == last ? first : $"{first}:{last}";

        return new OpenXmlChartData($"'{sheetName.Replace("'", "''")}'!{range}", null, null);
    }

    /// <summary>
    /// Numbers stored in the chart itself.
    /// </summary>
    public static OpenXmlChartData FromValues(params double[] values)
    {
        if (values is null or [])
            throw new ArgumentException("At least one value is required.", nameof(values));
        if (values.Any(v => double.IsNaN(v) || double.IsInfinity(v)))
            throw new ArgumentException("Chart values must be finite numbers.", nameof(values));

        return new OpenXmlChartData(null, values, null);
    }

    /// <summary>
    /// Text stored in the chart itself, typically category labels.
    /// </summary>
    public static OpenXmlChartData FromValues(params string[] values)
    {
        if (values is null or [])
            throw new ArgumentException("At least one value is required.", nameof(values));

        return new OpenXmlChartData(null, null, values);
    }

    private static string AbsoluteReference(int column, int row)
        => $"${CellReferenceConverter.GetAlphabeticalIndex(column - 1)}${row}";
}
