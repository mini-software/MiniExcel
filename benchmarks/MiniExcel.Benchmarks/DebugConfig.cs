using BenchmarkDotNet.Analysers;
using BenchmarkDotNet.Columns;
using BenchmarkDotNet.Configs;
using BenchmarkDotNet.Diagnosers;
using BenchmarkDotNet.Exporters;
using BenchmarkDotNet.Filters;
using BenchmarkDotNet.Jobs;
using BenchmarkDotNet.Loggers;
using BenchmarkDotNet.Order;
using BenchmarkDotNet.Reports;
using BenchmarkDotNet.Toolchains.InProcess.Emit;
using BenchmarkDotNet.Validators;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;

namespace MiniExcelLibs.Benchmarks
{
    public class DebugConfig : ManualConfig
    {
        public new IOrderer Orderer => DefaultOrderer.Instance;

        public new SummaryStyle SummaryStyle => SummaryStyle.Default;

        public new ConfigUnionRule UnionRule => ConfigUnionRule.Union;

        public new string ArtifactsPath => Path.Combine(Directory.GetCurrentDirectory(), "BenchmarkDotNet.Artifacts");

        public new CultureInfo CultureInfo => null;

        public new ConfigOptions Options => ConfigOptions.KeepBenchmarkFiles | ConfigOptions.DisableOptimizationsValidator;

        public new IEnumerable<Job> GetJobs()
        {
            return new Job[1]
            {
                 JobMode<Job>.Default.WithToolchain(new InProcessEmitToolchain(TimeSpan.FromHours(1.0), logOutput: true))
            };
        }


        public new IEnumerable<IValidator> GetValidators()
        {
            return Array.Empty<IValidator>();
        }

        public new IEnumerable<IColumnProvider> GetColumnProviders()
        {
            return DefaultColumnProviders.Instance;
        }

        public new IEnumerable<IExporter> GetExporters()
        {
            return Array.Empty<IExporter>();
        }

        public new IEnumerable<ILogger> GetLoggers()
        {
            return new ILogger[1]
            {
                    ConsoleLogger.Default
            };
        }

        public new IEnumerable<IDiagnoser> GetDiagnosers()
        {
            return Array.Empty<IDiagnoser>();
        }

        public new IEnumerable<IAnalyser> GetAnalysers()
        {
            return Array.Empty<IAnalyser>();
        }

        public new IEnumerable<HardwareCounter> GetHardwareCounters()
        {
            return Array.Empty<HardwareCounter>();
        }

        public new IEnumerable<IFilter> GetFilters()
        {
            return Array.Empty<IFilter>();
        }

        public new IEnumerable<BenchmarkLogicalGroupRule> GetLogicalGroupRules()
        {
            return Array.Empty<BenchmarkLogicalGroupRule>();
        }
    }
}
