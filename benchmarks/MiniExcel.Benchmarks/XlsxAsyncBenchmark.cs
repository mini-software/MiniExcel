using System;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using BenchmarkDotNet.Attributes;

namespace MiniExcelLibs.Benchmarks
{
    public class XlsxAsyncBenchmark : BenchmarkBase
    {
        [Benchmark(Description = "MiniExcel Async Create Xlsx")]
        public async Task MiniExcelCreateAsyncTest()
        {
            var value = Getvalue();
            var path = Path.Combine(Path.GetTempPath(), $"{Guid.NewGuid().ToString()}.xlsx");
            using (var stream = File.Create(path))
                await stream.SaveAsAsync(value);
            File.Delete(path);
        }

        [Benchmark(Description = "MiniExcel Async Template Generate")]
        public async Task MiniExcel_Template_Generate_Async_Test()
        {
            {
                var path = Path.Combine(Path.GetTempPath(), $"{Guid.NewGuid().ToString()}.xlsx");
                const string templatePath = @"TestTemplateBasicIEmumerableFill.xlsx";
                var value = new
                {
                    employees = Enumerable.Range(1, rowCount).Select(s => new { name = "Jack", department = "HR" })
                };
                await MiniExcel.SaveAsByTemplateAsync(path, templatePath, value);
            }
        }
    }
}
