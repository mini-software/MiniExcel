using System.Drawing;

namespace MiniExcelLib.OpenXml.Charts;

/// <summary>
/// The style of a line in a chart: an axis, gridlines or the chart border.
/// </summary>
public class OpenXmlChartLine
{
    /// <summary>
    /// The line color. The alpha channel sets the transparency. When empty, Excel picks the color.
    /// </summary>
    public Color Color { get; set; }

    /// <summary>
    /// The line width in points. When null, Excel picks the width.
    /// </summary>
    public double? Width { get; set; }
}
