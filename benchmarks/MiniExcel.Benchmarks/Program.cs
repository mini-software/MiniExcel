using System.ComponentModel;
using BenchmarkDotNet.Running;
using MiniExcelLibs.Benchmarks;
using MiniExcelLibs.Benchmarks.BenchmarkSections;

if (Environment.GetEnvironmentVariable("BenchmarkMode") == "Automatic")
{
    var section = Environment.GetEnvironmentVariable("BenchmarkSection");
    var benchmark = section?.ToLowerInvariant().Trim() switch
    {
        "query" => typeof(QueryXlsxBenchmark),
        "create" => typeof(CreateXlsxBenchmark),
        "template" => typeof(TemplateXlsxBenchmark),
        _ => throw new InvalidEnumArgumentException($"Benchmark section {section} does not exist")
    };
    
    BenchmarkRunner.Run(benchmark, new Config(), args);
}
else
    BenchmarkSwitcher
        .FromTypes([typeof(CreateXlsxBenchmark)])
        .Run(args, new Config());