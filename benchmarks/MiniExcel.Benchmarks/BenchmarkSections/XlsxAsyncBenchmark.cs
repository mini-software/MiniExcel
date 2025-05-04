using BenchmarkDotNet.Attributes;
using MiniExcelLibs.Benchmarks.Utils;

namespace MiniExcelLibs.Benchmarks.BenchmarkSections;

public class XlsxAsyncBenchmark : BenchmarkBase
{
    [Benchmark(Description = "MiniExcel Create Xlsx Async")]
    public async Task MiniExcelCreateAsyncTest()
    {
        using var path = AutoDeletingPath.Create();
        using var stream = File.Create(path.FilePath);

        await stream.SaveAsAsync(GetValue());
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
     
        await MiniExcel.SaveAsByTemplateAsync(path.FilePath, templatePath, value);
    }
}
