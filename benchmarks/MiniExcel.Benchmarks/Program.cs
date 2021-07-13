using System;
using System.ComponentModel;
using System.Diagnostics;
using BenchmarkDotNet.Configs;
using BenchmarkDotNet.Running;
using MiniExcelLibs;
using OfficeOpenXml;

namespace MiniExcelLibs.Benchmarks
{
    class Program
    {
        static void Main(string[] args)
        {
#if DEBUG
            new XlsxBenchmark().Epplus_QueryFirst_Test();
#else
            BenchmarkSwitcher.FromTypes(new[]{typeof(XlsxBenchmark)}).Run(args, new Config());
#endif
            Console.Read();
        }
    }
}
