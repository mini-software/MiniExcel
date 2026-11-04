namespace MiniExcelLib.OpenXml.Charts;

/// <summary>
/// A sheet that holds only a chart. Export it as one of the sheets of a workbook, next to the sheets with its data.
/// </summary>
public class OpenXmlChartSheet(OpenXmlChart chart)
{
    public OpenXmlChart Chart { get; } = chart ?? throw new ArgumentNullException(nameof(chart));
}
