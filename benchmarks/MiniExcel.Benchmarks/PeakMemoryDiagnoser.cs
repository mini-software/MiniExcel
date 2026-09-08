using System.ComponentModel;
using System.Diagnostics;
using BenchmarkDotNet.Analysers;
using BenchmarkDotNet.Columns;
using BenchmarkDotNet.Diagnosers;
using BenchmarkDotNet.Engines;
using BenchmarkDotNet.Exporters;
using BenchmarkDotNet.Loggers;
using BenchmarkDotNet.Reports;
using BenchmarkDotNet.Running;
using BenchmarkDotNet.Validators;

namespace MiniExcelLibs.Benchmarks;

public sealed class PeakMemoryDiagnoser : IDiagnoser
{
    private const int SampleIntervalMilliseconds = 20;
    private readonly Dictionary<BenchmarkCase, PeakMemory> _results = new();
    private CancellationTokenSource? _cancellation;
    private Task? _sampler;

    public IEnumerable<string> Ids => [nameof(PeakMemoryDiagnoser)];
    public IEnumerable<IExporter> Exporters => [];
    public IEnumerable<IAnalyser> Analysers => [];

    public RunMode GetRunMode(BenchmarkCase benchmarkCase) => RunMode.ExtraRun;

    public void Handle(HostSignal signal, DiagnoserActionParameters parameters)
    {
        if (signal == HostSignal.BeforeActualRun)
            Start(parameters);
        else if (signal is HostSignal.AfterActualRun or HostSignal.AfterAll or HostSignal.AfterProcessExit)
            Stop();
    }

    public IEnumerable<Metric> ProcessResults(DiagnoserResults results)
    {
        if (!_results.TryGetValue(results.BenchmarkCase, out var peakMemory))
            yield break;

        yield return new Metric(PeakWorkingSetMetricDescriptor.Instance, peakMemory.WorkingSet);
        yield return new Metric(PeakPrivateBytesMetricDescriptor.Instance, peakMemory.PrivateBytes);
    }

    public void DisplayResults(ILogger logger)
    {
    }

    public IEnumerable<ValidationError> Validate(ValidationParameters validationParameters) => [];

    private void Start(DiagnoserActionParameters parameters)
    {
        Stop();

        if (parameters.Process == null)
            return;

        var peakMemory = new PeakMemory();
        _results[parameters.BenchmarkCase] = peakMemory;
        _cancellation = new CancellationTokenSource();
        _sampler = Task.Run(() => Sample(parameters.Process, peakMemory, _cancellation.Token));
    }

    private void Stop()
    {
        if (_cancellation == null || _sampler == null)
            return;

        try
        {
            _cancellation.Cancel();
            _sampler.GetAwaiter().GetResult();
        }
        finally
        {
            _cancellation.Dispose();
            _cancellation = null;
            _sampler = null;
        }
    }

    private static void Sample(Process process, PeakMemory peakMemory, CancellationToken cancellationToken)
    {
        while (!cancellationToken.IsCancellationRequested)
        {
            try
            {
                process.Refresh();
                peakMemory.WorkingSet = Math.Max(peakMemory.WorkingSet, process.WorkingSet64);
                peakMemory.PrivateBytes = Math.Max(peakMemory.PrivateBytes, process.PrivateMemorySize64);
            }
            catch (InvalidOperationException)
            {
                break;
            }
            catch (Win32Exception)
            {
                break;
            }

            cancellationToken.WaitHandle.WaitOne(SampleIntervalMilliseconds);
        }
    }

    private sealed class PeakMemory
    {
        public long WorkingSet { get; set; }
        public long PrivateBytes { get; set; }
    }

    private abstract class PeakMemoryMetricDescriptor : IMetricDescriptor
    {
        public abstract string Id { get; }
        public abstract string DisplayName { get; }
        public abstract string Legend { get; }
        public string NumberFormat => "0.##";
        public UnitType UnitType => UnitType.Size;
        public string Unit => "B";
        public bool TheGreaterTheBetter => false;
        public int PriorityInCategory => 0;
        public bool GetIsAvailable(Metric metric) => metric.Value > 0;
    }

    private sealed class PeakWorkingSetMetricDescriptor : PeakMemoryMetricDescriptor
    {
        public static readonly IMetricDescriptor Instance = new PeakWorkingSetMetricDescriptor();
        public override string Id => "PeakWorkingSet";
        public override string DisplayName => "Peak Working Set";
        public override string Legend => $"Peak working set sampled every {SampleIntervalMilliseconds} ms during the diagnostic run";
    }

    private sealed class PeakPrivateBytesMetricDescriptor : PeakMemoryMetricDescriptor
    {
        public static readonly IMetricDescriptor Instance = new PeakPrivateBytesMetricDescriptor();
        public override string Id => "PeakPrivateBytes";
        public override string DisplayName => "Peak Private Bytes";
        public override string Legend => $"Peak private bytes sampled every {SampleIntervalMilliseconds} ms during the diagnostic run";
    }
}