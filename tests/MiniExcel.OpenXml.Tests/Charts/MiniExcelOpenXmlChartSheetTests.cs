using MiniExcelLib.OpenXml.Charts;
using OfficeOpenXml;
using OfficeOpenXml.Drawing.Chart;

namespace MiniExcelLib.OpenXml.Tests.Charts;

public class MiniExcelOpenXmlChartSheetTests
{
    private readonly OpenXmlExporter _exporter = MiniExcelV2.Exporters.GetOpenXmlExporter();

    public MiniExcelOpenXmlChartSheetTests()
    {
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
    }

    [Fact]
    public void ColumnChartSheetIsWritten()
    {
        var chart = new OpenXmlChart { Title = "Sales" };
        chart.Plots.Add(new OpenXmlChartPlot(OpenXmlChartType.Column)
        {
            Series = { new OpenXmlChartSeries(OpenXmlChartData.FromRange("Data", "B2:B3"), OpenXmlChartData.FromRange("Data", "A2:A3")) { Name = "Sales" } }
        });

        var sheets = new Dictionary<string, object>
        {
            ["Chart"] = new OpenXmlChartSheet(chart),
            ["Data"] = new[] { new { Month = "Jan", Sales = 10 }, new { Month = "Feb", Sales = 12 } }
        };

        using var stream = new MemoryStream();
        _exporter.Export(stream, sheets);
        stream.Position = 0;

        using var package = new ExcelPackage(stream);
        var chartSheet = Assert.IsType<ExcelChartsheet>(package.Workbook.Worksheets["Chart"]);
        var barChart = Assert.IsType<ExcelBarChart>(chartSheet.Chart);
        Assert.Equal(eChartType.ColumnClustered, barChart.ChartType);
        Assert.Equal("'Data'!$A$2:$A$3", barChart.Series.Single().XSeries);
        Assert.Equal("Jan", package.Workbook.Worksheets["Data"].Cells["A2"].Value);
    }
}
