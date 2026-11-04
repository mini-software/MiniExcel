using System.Drawing;

namespace MiniExcelLib.OpenXml.Charts;

internal static class ChartXml
{
    private const int PrimaryXAxisId = 500000001;
    private const int PrimaryYAxisId = 500000002;
    private const int SecondaryXAxisId = 500000003;
    private const int SecondaryYAxisId = 500000004;

    // The size Excel gives a new chart sheet; it is rescaled to the window when opened
    private const long ChartSheetWidthEmu = 9309100;
    private const long ChartSheetHeightEmu = 6072188;

    private enum AxisKind
    {
        Scatter,
        Vertical,
        Horizontal
    }

    internal static string ChartSheet(int sheetIndex) => XmlHelper.MinifyXml(
        $"""
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <chartsheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
            <sheetViews><sheetView zoomToFit="1" workbookViewId="0"/></sheetViews>
            <drawing r:id="rDrawing{sheetIndex}"/>
        </chartsheet>
        """);

    internal static string ChartRelationship(int chartIndex)
        => $"""<Relationship Id="rChart{chartIndex}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="/{ExcelFileNames.Chart(chartIndex)}" />""";

    internal static string ChartSheetAnchor(int chartIndex)
        => $"""
            <xdr:absoluteAnchor>
                <xdr:pos x="0" y="0"/>
                <xdr:ext cx="{ChartSheetWidthEmu}" cy="{ChartSheetHeightEmu}"/>
                <xdr:graphicFrame macro="">
                    <xdr:nvGraphicFramePr>
                        <xdr:cNvPr id="{chartIndex + 1}" name="Chart {chartIndex}"/>
                        <xdr:cNvGraphicFramePr><a:graphicFrameLocks noGrp="1"/></xdr:cNvGraphicFramePr>
                    </xdr:nvGraphicFramePr>
                    <xdr:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></xdr:xfrm>
                    <a:graphic>
                        <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart">
                            <c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" r:id="rChart{chartIndex}"/>
                        </a:graphicData>
                    </a:graphic>
                </xdr:graphicFrame>
                <xdr:clientData/>
            </xdr:absoluteAnchor>
            """;

    internal static void Validate(OpenXmlChart chart)
    {
        if (chart.Plots.Count == 0)
            throw new InvalidOperationException("A chart needs at least one plot.");
        if (chart.Plots.Any(p => p.Series.Count == 0))
            throw new InvalidOperationException("Every plot of a chart needs at least one series.");
        if (chart.Plots.All(p => p.UseSecondaryAxes))
            throw new InvalidOperationException("At least one plot of a chart must use the primary axes.");
        if (chart.Plots.SelectMany(p => p.Series).Any(s => s.Values.Strings is not null))
            throw new InvalidOperationException("Series values must be numbers or a cell range.");
        if (chart.Plots.Any(p => p.GapWidth is < 0 or > 500))
            throw new InvalidOperationException("The gap width of a plot must be from 0 to 500.");
        if (new[] { chart.XAxis, chart.YAxis, chart.SecondaryXAxis, chart.SecondaryYAxis }.Any(a => a.LabelRotation is < -90 or > 90))
            throw new InvalidOperationException("The label rotation of an axis must be from -90 to 90 degrees.");
        if (new[] { chart.XAxis, chart.YAxis, chart.SecondaryXAxis, chart.SecondaryYAxis }.SelectMany(a => new[] { a.Minimum, a.Maximum, a.MajorUnit }).Any(v => v is { } d && (double.IsNaN(d) || double.IsInfinity(d))))
            throw new InvalidOperationException("The minimum, maximum and major unit of an axis must be finite numbers.");

        foreach (var group in chart.Plots.GroupBy(p => p.UseSecondaryAxes))
        {
            var axes = group.Key ? "secondary" : "primary";
            if (group.Select(GetAxisKind).Distinct().Count() > 1)
                throw new InvalidOperationException($"The plots on the {axes} axes can't share them: scatter, column and line, and horizontal bar plots each need their own axes.");
            if (group.Count(p => p.IsBar) > 1 || group.Count(p => p.IsScatter) > 1 || group.Count(p => p.Type == OpenXmlChartType.Line) > 1)
                throw new InvalidOperationException($"The {axes} axes have more than one plot of the same kind; add the series to a single plot instead.");
        }
    }

    internal static string Chart(OpenXmlChart chart)
    {
        var sb = new StringBuilder();
        sb.Append("""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>""");
        sb.Append("""<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">""");
        sb.Append("""<c:roundedCorners val="0"/>""");
        sb.Append("<c:chart>");

        if (chart.Title is null)
        {
            sb.Append("""<c:autoTitleDeleted val="1"/>""");
        }
        else
        {
            AppendTitle(sb, chart.Title, chart.TitleFontSize, chart.TitleBold, vertical: false);
            sb.Append("""<c:autoTitleDeleted val="0"/>""");
        }

        sb.Append("<c:plotArea><c:layout/>");

        // Excel lists the plots by chart type, columns and bars before lines before scatter, primary axes first
        var plots = chart.Plots.OrderBy(p => p.IsBar ? 0 : p.IsScatter ? 2 : 1).ThenBy(p => p.UseSecondaryAxes).ToList();

        var seriesIndex = 0;
        foreach (var plot in plots)
        {
            AppendPlot(sb, plot, ref seriesIndex);
        }

        var primary = chart.Plots.First(p => !p.UseSecondaryAxes);
        AppendAxes(sb, GetAxisKind(primary), chart.XAxis, chart.YAxis, PrimaryXAxisId, PrimaryYAxisId, secondary: false);

        if (chart.Plots.FirstOrDefault(p => p.UseSecondaryAxes) is { } secondary)
        {
            AppendAxes(sb, GetAxisKind(secondary), chart.SecondaryXAxis, chart.SecondaryYAxis, SecondaryXAxisId, SecondaryYAxisId, secondary: true);
        }

        sb.Append("</c:plotArea>");

        if (chart.LegendPosition is { } legendPosition)
        {
            sb.Append($"""<c:legend><c:legendPos val="{GetLegendPosition(legendPosition)}"/>""");
            var entries = GetLegendOrder(plots, legendPosition);
            for (var i = 0; i < entries.Count; i++)
            {
                if (!entries[i].ShowInLegend)
                    sb.Append($"""<c:legendEntry><c:idx val="{i}"/><c:delete val="1"/></c:legendEntry>""");
            }
            sb.Append("""<c:overlay val="0"/></c:legend>""");
        }

        sb.Append("""<c:plotVisOnly val="1"/><c:dispBlanksAs val="gap"/>""");
        sb.Append("</c:chart>");

        if (chart.Border is { } border)
        {
            AppendLineProperties(sb, border);
        }

        sb.Append("</c:chartSpace>");
        return sb.ToString();
    }

    private static AxisKind GetAxisKind(OpenXmlChartPlot plot) => plot switch
    {
        { IsScatter: true } => AxisKind.Scatter,
        { Type: OpenXmlChartType.Bar or OpenXmlChartType.StackedBar } => AxisKind.Horizontal,
        _ => AxisKind.Vertical
    };

    // A legendEntry index is the position of the entry in the legend as Excel draws it, not a series index. Excel
    // lists horizontal bars bottom up, and stacked columns top down beside a vertical legend, so those are reversed
    private static List<OpenXmlChartSeries> GetLegendOrder(List<OpenXmlChartPlot> plots, OpenXmlChartLegendPosition position)
    {
        var verticalLegend = position is OpenXmlChartLegendPosition.Left or OpenXmlChartLegendPosition.Right or OpenXmlChartLegendPosition.TopRight;
        var entries = new List<OpenXmlChartSeries>();
        foreach (var plot in plots)
        {
            var reversed = plot.Type == OpenXmlChartType.Bar || (plot.Type == OpenXmlChartType.StackedColumn && verticalLegend);
            entries.AddRange(reversed ? Enumerable.Reverse(plot.Series) : plot.Series);
        }

        return entries;
    }

    private static void AppendPlot(StringBuilder sb, OpenXmlChartPlot plot, ref int seriesIndex)
    {
        var (xAxisId, yAxisId) = plot.UseSecondaryAxes ? (SecondaryXAxisId, SecondaryYAxisId) : (PrimaryXAxisId, PrimaryYAxisId);

        switch (plot.Type)
        {
            case OpenXmlChartType.Scatter or OpenXmlChartType.ScatterLines:
                sb.Append("""<c:scatterChart><c:scatterStyle val="lineMarker"/><c:varyColors val="0"/>""");
                break;
            case OpenXmlChartType.Line:
                sb.Append("""<c:lineChart><c:grouping val="standard"/><c:varyColors val="0"/>""");
                break;
            default:
                var direction = plot.Type is OpenXmlChartType.Bar or OpenXmlChartType.StackedBar ? "bar" : "col";
                var grouping = plot.Type is OpenXmlChartType.StackedColumn or OpenXmlChartType.StackedBar ? "stacked" : "clustered";
                sb.Append($"""<c:barChart><c:barDir val="{direction}"/><c:grouping val="{grouping}"/><c:varyColors val="0"/>""");
                break;
        }

        foreach (var series in plot.Series)
        {
            AppendSeries(sb, plot, series, seriesIndex++);
        }

        switch (plot.Type)
        {
            case OpenXmlChartType.Scatter or OpenXmlChartType.ScatterLines:
                sb.Append($"""<c:axId val="{xAxisId}"/><c:axId val="{yAxisId}"/></c:scatterChart>""");
                break;
            case OpenXmlChartType.Line:
                sb.Append($"""<c:marker val="1"/><c:axId val="{xAxisId}"/><c:axId val="{yAxisId}"/></c:lineChart>""");
                break;
            default:
                sb.Append($"""<c:gapWidth val="{plot.GapWidth}"/>""");
                if (plot.Type is OpenXmlChartType.StackedColumn or OpenXmlChartType.StackedBar)
                    sb.Append("""<c:overlap val="100"/>""");
                sb.Append($"""<c:axId val="{xAxisId}"/><c:axId val="{yAxisId}"/></c:barChart>""");
                break;
        }
    }

    private static void AppendSeries(StringBuilder sb, OpenXmlChartPlot plot, OpenXmlChartSeries series, int index)
    {
        sb.Append($"""<c:ser><c:idx val="{index}"/><c:order val="{index}"/>""");
        if (series.Name is not null)
            sb.Append($"<c:tx><c:v>{XmlHelper.EncodeXml(series.Name)}</c:v></c:tx>");

        if (plot.IsBar)
        {
            if (!series.Color.IsEmpty)
                sb.Append($"<c:spPr><a:solidFill>{Color(series.Color)}</a:solidFill></c:spPr>");
            sb.Append("""<c:invertIfNegative val="0"/>""");
        }
        else if (plot.Type is OpenXmlChartType.Scatter)
        {
            // Markers only: Excel stores this as a line marker series whose line isn't drawn
            sb.Append("""<c:spPr><a:ln w="19050"><a:noFill/></a:ln></c:spPr>""");
            sb.Append("""<c:marker><c:symbol val="circle"/><c:size val="5"/>""");
            if (!series.Color.IsEmpty)
                sb.Append($"<c:spPr><a:solidFill>{Color(series.Color)}</a:solidFill><a:ln><a:solidFill>{Color(series.Color)}</a:solidFill></a:ln></c:spPr>");
            sb.Append("</c:marker>");
        }
        else
        {
            AppendLineProperties(sb, new OpenXmlChartLine { Color = series.Color, Width = series.LineWidth });
            sb.Append("""<c:marker><c:symbol val="none"/></c:marker>""");
        }

        if (plot.IsScatter)
        {
            if (series.Categories is { } xValues)
                AppendData(sb, "c:xVal", xValues, textRange: false);
            AppendData(sb, "c:yVal", series.Values, textRange: false);
            sb.Append("""<c:smooth val="0"/>""");
        }
        else
        {
            if (series.Categories is { } categories)
                AppendData(sb, "c:cat", categories, textRange: true);
            AppendData(sb, "c:val", series.Values, textRange: false);
            if (plot.Type is OpenXmlChartType.Line)
                sb.Append("""<c:smooth val="0"/>""");
        }

        sb.Append("</c:ser>");
    }

    private static void AppendData(StringBuilder sb, string element, OpenXmlChartData data, bool textRange)
    {
        sb.Append($"<{element}>");
        if (data.Formula is { } formula)
        {
            // Excel refills the cache from the cells on load, but it numbers legend entries before that and skips
            // series without one, so an empty cache keeps the hidden entries pointing at the right series
            var (reference, cache) = textRange
                ? ("c:strRef", """<c:strCache><c:ptCount val="0"/></c:strCache>""")
                : ("c:numRef", """<c:numCache><c:formatCode>General</c:formatCode><c:ptCount val="0"/></c:numCache>""");
            sb.Append($"<{reference}><c:f>{XmlHelper.EncodeXml(formula)}</c:f>{cache}</{reference}>");
        }
        else if (data.Numbers is { } numbers)
        {
            sb.Append($"""<c:numLit><c:formatCode>General</c:formatCode><c:ptCount val="{numbers.Length}"/>""");
            for (var i = 0; i < numbers.Length; i++)
            {
                sb.Append($"""<c:pt idx="{i}"><c:v>{Number(numbers[i])}</c:v></c:pt>""");
            }
            sb.Append("</c:numLit>");
        }
        else
        {
            var strings = data.Strings!;
            sb.Append($"""<c:strLit><c:ptCount val="{strings.Length}"/>""");
            for (var i = 0; i < strings.Length; i++)
            {
                sb.Append($"""<c:pt idx="{i}"><c:v>{XmlHelper.EncodeXml(strings[i])}</c:v></c:pt>""");
            }
            sb.Append("</c:strLit>");
        }
        sb.Append($"</{element}>");
    }

    private static void AppendAxes(StringBuilder sb, AxisKind kind, OpenXmlChartAxis xAxis, OpenXmlChartAxis yAxis, int xAxisId, int yAxisId, bool secondary)
    {
        var (xPosition, yPosition) = (kind, secondary) switch
        {
            (AxisKind.Horizontal, false) => ("l", "b"),
            (AxisKind.Horizontal, true) => ("r", "t"),
            (_, false) => ("b", "l"),
            (_, true) => ("t", "r")
        };

        // The value axis crosses its partner at the far end on the secondary axes, so it sits opposite the primary one
        var yCrosses = secondary ? "max" : "autoZero";
        var crossBetween = kind == AxisKind.Scatter ? "midCat" : "between";

        if (kind == AxisKind.Scatter)
        {
            AppendValueAxis(sb, xAxis, xAxisId, yAxisId, xPosition, "autoZero", crossBetween);
        }
        else
        {
            AppendCategoryAxis(sb, xAxis, xAxisId, yAxisId, xPosition);
        }

        AppendValueAxis(sb, yAxis, yAxisId, xAxisId, yPosition, yCrosses, crossBetween);
    }

    private static void AppendValueAxis(StringBuilder sb, OpenXmlChartAxis axis, int axisId, int crossAxisId, string position, string crosses, string crossBetween)
    {
        sb.Append($"""<c:valAx><c:axId val="{axisId}"/><c:scaling><c:orientation val="minMax"/>""");
        if (axis.Maximum is { } max)
            sb.Append($"""<c:max val="{Number(max)}"/>""");
        if (axis.Minimum is { } min)
            sb.Append($"""<c:min val="{Number(min)}"/>""");
        sb.Append("</c:scaling>");

        AppendAxisBody(sb, axis, crossAxisId, position);

        sb.Append($"""<c:crosses val="{crosses}"/><c:crossBetween val="{crossBetween}"/>""");
        if (axis.MajorUnit is { } majorUnit)
            sb.Append($"""<c:majorUnit val="{Number(majorUnit)}"/>""");
        sb.Append("</c:valAx>");
    }

    private static void AppendCategoryAxis(StringBuilder sb, OpenXmlChartAxis axis, int axisId, int crossAxisId, string position)
    {
        sb.Append($"""<c:catAx><c:axId val="{axisId}"/><c:scaling><c:orientation val="minMax"/></c:scaling>""");
        AppendAxisBody(sb, axis, crossAxisId, position);
        sb.Append("""<c:crosses val="autoZero"/><c:auto val="1"/><c:lblAlgn val="ctr"/><c:lblOffset val="100"/><c:noMultiLvlLbl val="0"/></c:catAx>""");
    }

    // The elements every axis shares, from delete to crossAx
    private static void AppendAxisBody(StringBuilder sb, OpenXmlChartAxis axis, int crossAxisId, string position)
    {
        sb.Append($"""<c:delete val="{(axis.Visible ? 0 : 1)}"/><c:axPos val="{position}"/>""");

        if (axis.MajorGridlines is { } gridlines)
        {
            sb.Append("<c:majorGridlines>");
            AppendLineProperties(sb, gridlines);
            sb.Append("</c:majorGridlines>");
        }

        if (axis.Title is { } title)
            AppendTitle(sb, title, null, false, vertical: position is "l" or "r");

        sb.Append(axis.NumberFormat is { } format
            ? $"""<c:numFmt formatCode="{XmlHelper.EncodeXml(format)}" sourceLinked="0"/>"""
            : """<c:numFmt formatCode="General" sourceLinked="1"/>""");

        sb.Append($"""<c:majorTickMark val="{GetTickMark(axis.MajorTickMark)}"/><c:minorTickMark val="{GetTickMark(axis.MinorTickMark)}"/><c:tickLblPos val="nextTo"/>""");

        if (axis.Line is { } line)
            AppendLineProperties(sb, line);

        if (axis.LabelRotation is { } rotation)
        {
            // DrawingML turns text clockwise, while Excel's dialog turns it counterclockwise
            sb.Append($"""<c:txPr><a:bodyPr rot="{-rotation * 60000}" vert="horz"/><a:lstStyle/><a:p><a:pPr><a:defRPr/></a:pPr><a:endParaRPr lang="en-US"/></a:p></c:txPr>""");
        }

        sb.Append($"""<c:crossAx val="{crossAxisId}"/>""");
    }

    private static void AppendTitle(StringBuilder sb, string text, double? fontSize, bool bold, bool vertical)
    {
        var size = fontSize is { } points ? $" sz=\"{(int)Math.Round(points * 100)}\"" : "";
        var weight = bold ? " b=\"1\"" : " b=\"0\"";
        var rotation = vertical ? """ rot="-5400000" vert="horz" """ : " ";

        sb.Append($"<c:title><c:tx><c:rich><a:bodyPr{rotation}/><a:lstStyle/>");
        sb.Append($"<a:p><a:pPr><a:defRPr{size}{weight}/></a:pPr><a:r><a:rPr lang=\"en-US\"{size}{weight}/><a:t>{XmlHelper.EncodeXml(text)}</a:t></a:r></a:p>");
        sb.Append("""</c:rich></c:tx><c:overlay val="0"/></c:title>""");
    }

    private static void AppendLineProperties(StringBuilder sb, OpenXmlChartLine line)
    {
        if (line.Color.IsEmpty && line.Width is null)
            return;

        var width = line.Width is { } points ? $" w=\"{(long)Math.Round(points * 12700)}\"" : "";
        sb.Append($"<c:spPr><a:ln{width}>");
        if (!line.Color.IsEmpty)
            sb.Append($"<a:solidFill>{Color(line.Color)}</a:solidFill>");
        sb.Append("</a:ln></c:spPr>");
    }

    private static string Color(Color color)
    {
        var rgb = $"{color.R:X2}{color.G:X2}{color.B:X2}";
        return color.A == 255
            ? $"""<a:srgbClr val="{rgb}"/>"""
            : $"""<a:srgbClr val="{rgb}"><a:alpha val="{(int)Math.Round(color.A * 100000 / 255.0)}"/></a:srgbClr>""";
    }

    private static string Number(double value) => value.ToString("R", CultureInfo.InvariantCulture);

    private static string GetTickMark(OpenXmlChartTickMark tickMark) => tickMark switch
    {
        OpenXmlChartTickMark.None => "none",
        OpenXmlChartTickMark.Inside => "in",
        OpenXmlChartTickMark.Outside => "out",
        OpenXmlChartTickMark.Cross => "cross",
        _ => throw new InvalidEnumArgumentException(nameof(tickMark), (int)tickMark, typeof(OpenXmlChartTickMark))
    };

    private static string GetLegendPosition(OpenXmlChartLegendPosition position) => position switch
    {
        OpenXmlChartLegendPosition.Right => "r",
        OpenXmlChartLegendPosition.Left => "l",
        OpenXmlChartLegendPosition.Top => "t",
        OpenXmlChartLegendPosition.Bottom => "b",
        OpenXmlChartLegendPosition.TopRight => "tr",
        _ => throw new InvalidEnumArgumentException(nameof(position), (int)position, typeof(OpenXmlChartLegendPosition))
    };
}
