namespace MiniExcelLib.OpenXml.Charts;

public class OpenXmlChartAxis
{
    /// <summary>
    /// Whether the axis line, tick marks and labels are drawn. A hidden axis still scales its plots.
    /// </summary>
    public bool Visible { get; set; } = true;

    public string? Title { get; set; }

    /// <summary>
    /// The lowest value on the axis. When null, Excel picks it. Ignored by category axes.
    /// </summary>
    public double? Minimum { get; set; }

    /// <summary>
    /// The highest value on the axis. When null, Excel picks it. Ignored by category axes.
    /// </summary>
    public double? Maximum { get; set; }

    /// <summary>
    /// The distance between major tick marks. When null, Excel picks it. Ignored by category axes.
    /// </summary>
    public double? MajorUnit { get; set; }

    /// <summary>
    /// The number format of the labels, such as <c>#,##0</c>. When null, the labels follow the source data.
    /// </summary>
    public string? NumberFormat { get; set; }

    /// <summary>
    /// The rotation of the labels in degrees from -90 to 90, counterclockwise like in Excel. When null, Excel picks it.
    /// </summary>
    public int? LabelRotation { get; set; }

    public OpenXmlChartTickMark MajorTickMark { get; set; } = OpenXmlChartTickMark.Outside;

    public OpenXmlChartTickMark MinorTickMark { get; set; } = OpenXmlChartTickMark.None;

    /// <summary>
    /// The style of the axis line. When null, Excel picks it.
    /// </summary>
    public OpenXmlChartLine? Line { get; set; }

    /// <summary>
    /// The style of the major gridlines. When null, no gridlines are drawn.
    /// </summary>
    public OpenXmlChartLine? MajorGridlines { get; set; }
}
