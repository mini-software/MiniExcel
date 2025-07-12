using BenchmarkDotNet.Running;
using MiniExcelLib.Benchmarks;
using MiniExcelLib.Benchmarks.BenchmarkSections;

if (Environment.GetEnvironmentVariable("BenchmarkMode") == "Automatic")
{
    var section = Environment.GetEnvironmentVariable("BenchmarkSection");
    var benchmark = section?.ToLowerInvariant().Trim() switch
    {
        "query" => typeof(QueryExcelBenchmark),
        "create" => typeof(CreateExcelBenchmark),
        "template" => typeof(TemplateExcelBenchmark),
        _ => throw new ArgumentException($"Benchmark section {section} does not exist")
    };
    
    BenchmarkRunner.Run(benchmark, BenchmarkConfig.Default, args);
}
else
{
    BenchmarkSwitcher
        .FromTypes(
        [
            typeof(QueryExcelBenchmark),
            typeof(CreateExcelBenchmark),
            typeof(TemplateExcelBenchmark)
        ])
        .Run(args, BenchmarkConfig.Default);
}