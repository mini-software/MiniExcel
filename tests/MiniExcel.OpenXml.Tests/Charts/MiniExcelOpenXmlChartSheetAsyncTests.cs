using System.Drawing;
using MiniExcelLib.OpenXml.Charts;
using OfficeOpenXml;
using OfficeOpenXml.Drawing.Chart;

namespace MiniExcelLib.OpenXml.Tests.Charts;

public class MiniExcelOpenXmlChartSheetAsyncTests
{
    private readonly OpenXmlExporter _exporter = MiniExcelV2.Exporters.GetOpenXmlExporter();
    private readonly OpenXmlImporter _importer = MiniExcelV2.Importers.GetOpenXmlImporter();

    public MiniExcelOpenXmlChartSheetAsyncTests()
    {
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
    }

    private static readonly object[] Readings =
    [
        new { Odometer = 0.0, Speed = 1.5 },
        new { Odometer = 10.0, Speed = 4.5 },
        new { Odometer = 20.0, Speed = 6.5 }
    ];

    private static OpenXmlChart SpeedChart()
    {
        var chart = new OpenXmlChart
        {
            Title = "Speed",
            TitleFontSize = 18,
            TitleBold = true,
            LegendPosition = OpenXmlChartLegendPosition.Bottom,
            Border = new OpenXmlChartLine { Width = 0.75 }
        };
        chart.XAxis.Title = "Odometer (m)";
        chart.XAxis.Minimum = 0;
        chart.XAxis.Maximum = 20;
        chart.XAxis.NumberFormat = "#,##0";
        chart.XAxis.LabelRotation = -75;
        chart.YAxis.Title = "Speed (m/s)";
        chart.YAxis.Minimum = 0;
        chart.YAxis.Maximum = 8;
        chart.YAxis.MajorGridlines = new OpenXmlChartLine { Color = Color.Gray, Width = 0.75 };
        chart.SecondaryYAxis.Visible = false;
        chart.SecondaryYAxis.Minimum = 0;
        chart.SecondaryYAxis.Maximum = 8;

        chart.Plots.Add(new OpenXmlChartPlot(OpenXmlChartType.ScatterLines)
        {
            Series =
            {
                new OpenXmlChartSeries(OpenXmlChartData.FromRange("Raw", "B2:B4"), OpenXmlChartData.FromRange("Raw", "A2:A4"))
                {
                    Name = "Speed",
                    Color = Color.Black,
                    LineWidth = 2,
                    ShowInLegend = false
                }
            }
        });

        // Bands behind the line: two full-width columns stacked to the band limits
        var bands = new OpenXmlChartPlot(OpenXmlChartType.StackedColumn) { UseSecondaryAxes = true, GapWidth = 0 };
        var categories = OpenXmlChartData.FromValues(0, 20);
        bands.Series.Add(new OpenXmlChartSeries(OpenXmlChartData.FromValues(5, 5), categories) { Name = "Standard", Color = Color.FromArgb(64, Color.Green) });
        bands.Series.Add(new OpenXmlChartSeries(OpenXmlChartData.FromValues(1, 1), categories) { Name = "Modified", Color = Color.FromArgb(64, Color.Orange) });
        bands.Series.Add(new OpenXmlChartSeries(OpenXmlChartData.FromValues(2, 2), categories) { Name = "Best efforts", Color = Color.FromArgb(64, Color.Red) });
        chart.Plots.Add(bands);

        return chart;
    }

    private async Task<MemoryStream> ExportAsync(OpenXmlChart chart)
    {
        var sheets = new Dictionary<string, object>
        {
            ["Raw"] = Readings,
            ["Speed Graph"] = new OpenXmlChartSheet(chart)
        };

        var stream = new MemoryStream();
        await _exporter.ExportAsync(stream, sheets);
        stream.Position = 0;
        return stream;
    }

    [Fact]
    public async Task ChartSheetPlotsTheDataSheetAsync()
    {
        using var stream = await ExportAsync(SpeedChart());

        using var package = new ExcelPackage(stream);
        var chartSheet = Assert.IsType<ExcelChartsheet>(package.Workbook.Worksheets["Speed Graph"]);
        var chart = chartSheet.Chart;
        var line = Assert.IsType<ExcelScatterChart>(chart.PlotArea.ChartTypes.Single(c => c is ExcelScatterChart));

        Assert.Equal("Speed", chart.Title.Text);
        Assert.Equal("Odometer (m)", line.XAxis.Title.Text);
        Assert.Equal("Speed (m/s)", line.YAxis.Title.Text);
        Assert.Equal(0, line.YAxis.MinValue);
        Assert.Equal(8, line.YAxis.MaxValue);
        Assert.Equal("#,##0", line.XAxis.Format);

        var series = line.Series.Single();
        Assert.Equal("'Raw'!$B$2:$B$4", series.Series);
        Assert.Equal("'Raw'!$A$2:$A$4", series.XSeries);
    }

    [Fact]
    public async Task SecondaryPlotIsCombinedWithThePrimaryOneAsync()
    {
        using var stream = await ExportAsync(SpeedChart());

        using var package = new ExcelPackage(stream);
        var chart = ((ExcelChartsheet)package.Workbook.Worksheets["Speed Graph"]).Chart;

        Assert.Equal(2, chart.PlotArea.ChartTypes.Count);
        var bands = Assert.IsType<ExcelBarChart>(chart.PlotArea.ChartTypes.Single(c => c is ExcelBarChart));
        Assert.Equal(3, bands.Series.Count);
        Assert.Equal(0, bands.GapWidth);
    }

    [Fact]
    public async Task ChartXmlKeepsStylesLegendAndLiteralsAsync()
    {
        using var stream = await ExportAsync(SpeedChart());
        using var archive = new ZipArchive(stream, ZipArchiveMode.Read);

        using var reader = new StreamReader(archive.GetEntry("xl/charts/chart2.xml")!.Open());
        var chartXml = await reader.ReadToEndAsync();

        // Excel lists columns before lines, so the line added first is series and legend entry 3
        Assert.Contains("""<c:scatterChart><c:scatterStyle val="lineMarker"/><c:varyColors val="0"/><c:ser><c:idx val="3"/>""", chartXml);
        Assert.Contains("""<c:legendEntry><c:idx val="3"/><c:delete val="1"/></c:legendEntry>""", chartXml);
        Assert.Contains("""<a:srgbClr val="008000"><a:alpha val="25098"/></a:srgbClr>""", chartXml);
        Assert.Contains("""<c:numLit><c:formatCode>General</c:formatCode><c:ptCount val="2"/><c:pt idx="0"><c:v>5</c:v></c:pt>""", chartXml);
        Assert.Contains("""<a:bodyPr rot="4500000" vert="horz"/>""", chartXml);
        Assert.Contains("""<c:barChart><c:barDir val="col"/><c:grouping val="stacked"/>""", chartXml);
        Assert.Contains("""<c:crosses val="max"/>""", chartXml);
    }

    [Theory]
    [InlineData(OpenXmlChartLegendPosition.Right, 2)]
    [InlineData(OpenXmlChartLegendPosition.Bottom, 0)]
    public async Task HiddenLegendEntryFollowsExcelsLegendOrderAsync(OpenXmlChartLegendPosition position, int expectedEntry)
    {
        // Beside a vertical legend Excel lists stacked columns top down, so the first series is the last entry
        var chart = new OpenXmlChart { LegendPosition = position };
        var plot = new OpenXmlChartPlot(OpenXmlChartType.StackedColumn);
        plot.Series.Add(new OpenXmlChartSeries(OpenXmlChartData.FromValues(1, 2)) { ShowInLegend = false });
        plot.Series.Add(new OpenXmlChartSeries(OpenXmlChartData.FromValues(3, 4)));
        plot.Series.Add(new OpenXmlChartSeries(OpenXmlChartData.FromValues(5, 6)));
        chart.Plots.Add(plot);

        using var stream = await ExportAsync(chart);
        using var archive = new ZipArchive(stream, ZipArchiveMode.Read);
        using var reader = new StreamReader(archive.GetEntry("xl/charts/chart2.xml")!.Open());
        var chartXml = await reader.ReadToEndAsync();

        Assert.Contains($"""<c:legendEntry><c:idx val="{expectedEntry}"/><c:delete val="1"/></c:legendEntry>""", chartXml);
    }

    [Fact]
    public async Task DataSheetIsStillReadableAsync()
    {
        using var stream = await ExportAsync(SpeedChart());

        var names = await _importer.GetSheetNamesAsync(stream, leaveOpen: true);
        Assert.Equal(["Raw", "Speed Graph"], names);

        stream.Position = 0;
        var rows = await _importer.QueryAsync(stream, hasHeaderRow: true, sheetName: "Raw").ToListAsync();
        Assert.Equal(3, rows.Count);
        Assert.Equal(6.5, (double)rows[2].Speed);
    }

    [Fact]
    public async Task InsertingASheetKeepsTheChartSheetAsync()
    {
        using var stream = await ExportAsync(SpeedChart());

        await _exporter.InsertSheetAsync(stream, new[] { new { Note = "added" } }, sheetName: "Notes");

        stream.Position = 0;
        using var package = new ExcelPackage(stream);
        Assert.IsType<ExcelChartsheet>(package.Workbook.Worksheets["Speed Graph"]);
        Assert.Equal("added", package.Workbook.Worksheets["Notes"].Cells["A2"].Value);
    }

    [Fact]
    public async Task ChartSheetCanNotBeInsertedAsync()
    {
        using var stream = await ExportAsync(SpeedChart());

        await Assert.ThrowsAsync<NotSupportedException>(() => _exporter.InsertSheetAsync(stream, new OpenXmlChartSheet(SpeedChart()), sheetName: "Another"));
    }

    [Fact]
    public async Task ScatterAndColumnPlotsCanNotShareAxesAsync()
    {
        var chart = SpeedChart();
        chart.Plots[1].UseSecondaryAxes = false;

        var exception = await Assert.ThrowsAsync<InvalidOperationException>(() => ExportAsync(chart));
        Assert.Contains("primary axes", exception.Message);
    }

    [Fact]
    public async Task ChartWithoutPlotsIsRejectedAsync()
    {
        await Assert.ThrowsAsync<InvalidOperationException>(() => ExportAsync(new OpenXmlChart()));
    }

    [Theory]
    [InlineData("Raw", "A2:A10", "'Raw'!$A$2:$A$10")]
    [InlineData("O'Brien", "c3", "'O''Brien'!$C$3")]
    public void RangeIsAbsoluteAndQuoted(string sheetName, string range, string expected)
    {
        var data = OpenXmlChartData.FromRange(sheetName, range);
        Assert.Equal(expected, data.Formula);
    }

    [Theory]
    [InlineData("A")]
    [InlineData("A1:B2:C3")]
    [InlineData("B5:A1")]
    public void InvalidRangeIsRejected(string range)
    {
        Assert.ThrowsAny<ArgumentException>(() => OpenXmlChartData.FromRange("Raw", range));
    }
}
