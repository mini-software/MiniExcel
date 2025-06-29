using BenchmarkDotNet.Running;
using MiniExcelLib.Benchmarks;
using MiniExcelLib.Benchmarks.BenchmarkSections;

if (Environment.GetEnvironmentVariable("BenchmarkMode") == "Automatic")
{
    var section = Environment.GetEnvironmentVariable("BenchmarkSection");
    var benchmark = section?.ToLowerInvariant().Trim() switch
    {
        "query" => typeof(QueryXlsxBenchmark),
        "create" => typeof(CreateXlsxBenchmark),
        "template" => typeof(TemplateXlsxBenchmark),
        _ => throw new ArgumentException($"Benchmark section {section} does not exist")
    };
    
    BenchmarkRunner.Run(benchmark, BenchmarkConfig.Default, args);
}
else
{
    BenchmarkSwitcher
        .FromTypes(
        [
            typeof(QueryXlsxBenchmark),
            typeof(CreateXlsxBenchmark),
            typeof(TemplateXlsxBenchmark)
        ])
        .Run(args, BenchmarkConfig.Default);
}