using System.Drawing;

namespace MiniExcelLib.OpenXml.Charts;

/// <summary>
/// One series of a chart plot.
/// </summary>
/// <param name="values">The values of the series: the Y values of a scatter plot, or the bar heights and line points of the other plots.</param>
/// <param name="categories">The X values of a scatter plot, or the category labels of the other plots. When omitted, points are numbered from 1.</param>
public class OpenXmlChartSeries(OpenXmlChartData values, OpenXmlChartData? categories = null)
{
    public OpenXmlChartData Values { get; } = values ?? throw new ArgumentNullException(nameof(values));

    public OpenXmlChartData? Categories { get; } = categories;

    /// <summary>
    /// The name shown in the legend.
    /// </summary>
    public string? Name { get; set; }

    /// <summary>
    /// The line color of scatter and line series, or the fill color of column and bar series. The alpha channel sets
    /// the transparency. When empty, Excel picks the color.
    /// </summary>
    public Color Color { get; set; }

    /// <summary>
    /// The line width in points, for scatter and line series. When null, Excel picks the width.
    /// </summary>
    public double? LineWidth { get; set; }

    /// <summary>
    /// Whether the series has an entry in the legend.
    /// </summary>
    public bool ShowInLegend { get; set; } = true;
}
