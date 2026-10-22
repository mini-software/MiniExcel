namespace MiniExcelLib.OpenXml.Charts;

/// <summary>
/// A group of series drawn the same way on the same axes. A chart with several plots combines them, for example columns
/// on the secondary axes behind a line on the primary axes.
/// </summary>
public class OpenXmlChartPlot(OpenXmlChartType type)
{
    public OpenXmlChartType Type { get; } = type;

    /// <summary>
    /// Whether the plot uses the secondary axes of the chart instead of the primary ones.
    /// </summary>
    public bool UseSecondaryAxes { get; set; }

    /// <summary>
    /// The gap between column or bar clusters as a percentage of the bar width, from 0 to 500. Ignored by other plots.
    /// </summary>
    public int GapWidth { get; set; } = 150;

    public List<OpenXmlChartSeries> Series { get; } = [];

    internal bool IsScatter => Type is OpenXmlChartType.Scatter or OpenXmlChartType.ScatterLines;

    internal bool IsBar => Type is OpenXmlChartType.Column or OpenXmlChartType.StackedColumn or OpenXmlChartType.Bar or OpenXmlChartType.StackedBar;
}
