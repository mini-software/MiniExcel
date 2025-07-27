using BenchmarkDotNet.Attributes;
using MiniExcelLib.Benchmarks.Utils;
using MiniExcelLib.Core;

namespace MiniExcelLib.Benchmarks.BenchmarkSections;

public class XlsxAsyncBenchmark : BenchmarkBase
{
    private OpenXmlExporter _exporter;
    private OpenXmlTemplater _templater;
    
    [GlobalSetup]
    public void Setup()
    {
        _exporter = MiniExcel.Exporters.GetOpenXmlExporter();
        _templater = MiniExcel.Templaters.GetOpenXmlTemplater();
    }
    
    [Benchmark(Description = "MiniExcel Create Xlsx Async")]
    public async Task MiniExcelCreateAsyncTest()
    {
        using var path = AutoDeletingPath.Create();
        await using var stream = File.Create(path.FilePath);

        await _exporter.ExportAsync(stream, GetValue());
    }

    [Benchmark(Description = "MiniExcel Generate Template Async")]
    public async Task MiniExcel_Template_Generate_Async_Test()
    {
        const string templatePath = "TestTemplateBasicIEmumerableFill.xlsx";

        using var path = AutoDeletingPath.Create();
        var value = new
        {
            employees = Enumerable.Range(1, RowCount)
                .Select(s => new 
                {
                    name = "Jack",
                    department = "HR"
                })
        };
     
        await _templater.ApplyTemplateAsync(path.FilePath, templatePath, value);
    }
}
