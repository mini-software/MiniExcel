using BenchmarkDotNet.Attributes;
using MiniExcelLib.Benchmarks.Utils;

namespace MiniExcelLib.Benchmarks.BenchmarkSections;

public class XlsxAsyncBenchmark : BenchmarkBase
{
    private MiniExcelExporter _exporter;
    private MiniExcelTemplater _templater;
    
    [GlobalSetup]
    public void Setup()
    {
        _exporter = new MiniExcelExporter();
        _templater = new MiniExcelTemplater();
    }
    
    [Benchmark(Description = "MiniExcel Create Xlsx Async")]
    public async Task MiniExcelCreateAsyncTest()
    {
        using var path = AutoDeletingPath.Create();
        await using var stream = File.Create(path.FilePath);

        await _exporter.ExportXlsxAsync(stream, GetValue());
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
     
        await _templater.ApplyXlsxTemplateAsync(path.FilePath, templatePath, value);
    }
}
