using BenchmarkDotNet.Running;
using MiniExcelLibs.Benchmarks;


#if DEBUG
new XlsxBenchmark().Epplus_QueryFirst_Test();
#else
BenchmarkSwitcher
    .FromTypes([typeof(XlsxBenchmark)])
    .Run(args, new Config());
#endif

Console.Read();
