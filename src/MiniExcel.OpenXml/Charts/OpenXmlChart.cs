namespace MiniExcelLib.OpenXml.Charts;

/// <summary>
/// A chart made of one or more plots that share primary and, optionally, secondary axes.
/// </summary>
public class OpenXmlChart
{
    public string? Title { get; set; }

    /// <summary>
    /// The title font size in points. When null, Excel picks it.
    /// </summary>
    public double? TitleFontSize { get; set; }

    public bool TitleBold { get; set; }

    /// <summary>
    /// Where the legend is placed. When null, the chart has no legend.
    /// </summary>
    public OpenXmlChartLegendPosition? LegendPosition { get; set; } = OpenXmlChartLegendPosition.Right;

    /// <summary>
    /// The style of the chart border. When null, Excel picks it.
    /// </summary>
    public OpenXmlChartLine? Border { get; set; }

    /// <summary>
    /// The horizontal axis of the primary plots, or the vertical one for bar plots.
    /// </summary>
    public OpenXmlChartAxis XAxis { get; } = new();

    /// <summary>
    /// The value axis of the primary plots.
    /// </summary>
    public OpenXmlChartAxis YAxis { get; } = new() { MajorGridlines = new OpenXmlChartLine() };

    /// <summary>
    /// The category or X axis of the plots on the secondary axes. Hidden unless made visible.
    /// </summary>
    public OpenXmlChartAxis SecondaryXAxis { get; } = new() { Visible = false };

    /// <summary>
    /// The value axis of the plots on the secondary axes, drawn on the opposite side of the primary one.
    /// </summary>
    public OpenXmlChartAxis SecondaryYAxis { get; } = new();

    public List<OpenXmlChartPlot> Plots { get; } = [];
}
