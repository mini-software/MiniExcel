using BenchmarkDotNet.Columns;
using BenchmarkDotNet.Configs;
using BenchmarkDotNet.Diagnosers;
using BenchmarkDotNet.Exporters;
using BenchmarkDotNet.Exporters.Csv;
using BenchmarkDotNet.Jobs;
using BenchmarkDotNet.Loggers;
using BenchmarkDotNet.Order;

namespace MiniExcelLib.Benchmarks;

internal class BenchmarkConfig : ManualConfig
{
    private const int Launches = 1;
    private const int Warmups = 3;
    private const int Unrolls = 3;
    private const int Iterations = 3;

    public static BenchmarkConfig Default => new();
    
    private BenchmarkConfig()
    {
        AddLogger(ConsoleLogger.Default);

        AddExporter(CsvExporter.Default);
        AddExporter(MarkdownExporter.GitHub);
        AddExporter(HtmlExporter.Default);

        AddDiagnoser(MemoryDiagnoser.Default);
        AddColumn(TargetMethodColumn.Method);
        AddColumn(StatisticColumn.Mean);
        AddColumn(StatisticColumn.StdDev);
        AddColumn(StatisticColumn.Error);
        AddColumn(BaselineRatioColumn.RatioMean);
        AddColumnProvider(DefaultColumnProviders.Metrics);

        AddJob(Job.ShortRun
            .WithLaunchCount(Launches)
            .WithWarmupCount(Warmups)
            .WithUnrollFactor(Unrolls)
            .WithIterationCount(Iterations));

        Orderer = new DefaultOrderer(SummaryOrderPolicy.FastestToSlowest);
        Options |= ConfigOptions.JoinSummary;
    }
}
