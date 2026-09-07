using System.Diagnostics;
using System.Globalization;
using System.Reflection;
using MiniExcelLibs;

const string workerArgument = "--worker";

if (args.Contains(workerArgument, StringComparer.OrdinalIgnoreCase))
{
    RunWorker(GetFilePath(args));
    return;
}

await RunControllerAsync(GetFilePath(args));

static async Task RunControllerAsync(string filePath)
{
    if (!File.Exists(filePath))
        throw new FileNotFoundException("The benchmark workbook was not found.", filePath);

    var assemblyPath = Assembly.GetExecutingAssembly().Location;
    var processPath = Environment.ProcessPath
        ?? throw new InvalidOperationException("Unable to determine the current process path.");
    var startInfo = new ProcessStartInfo
    {
        FileName = processPath,
        RedirectStandardOutput = true,
        RedirectStandardError = true,
        UseShellExecute = false
    };

    if (Path.GetFileNameWithoutExtension(processPath).Equals("dotnet", StringComparison.OrdinalIgnoreCase))
        startInfo.ArgumentList.Add(assemblyPath);

    startInfo.ArgumentList.Add(workerArgument);
    startInfo.ArgumentList.Add("--file");
    startInfo.ArgumentList.Add(filePath);

    using var process = Process.Start(startInfo)
        ?? throw new InvalidOperationException("Unable to start the measurement worker.");
    var outputTask = process.StandardOutput.ReadToEndAsync();
    var errorTask = process.StandardError.ReadToEndAsync();
    long peakWorkingSet = 0;
    long peakPrivateBytes = 0;

    while (true)
    {
        try
        {
            if (process.HasExited)
                break;
            process.Refresh();
            peakWorkingSet = Math.Max(peakWorkingSet, process.WorkingSet64);
            peakPrivateBytes = Math.Max(peakPrivateBytes, process.PrivateMemorySize64);
        }
        catch (InvalidOperationException) when (process.HasExited)
        {
            break;
        }
        await Task.Delay(20);
    }

    await process.WaitForExitAsync();
    var output = await outputTask;
    var error = await errorTask;

    if (process.ExitCode != 0)
        throw new InvalidOperationException($"Measurement worker failed:{Environment.NewLine}{error}");

    var resultLine = output.Split(Environment.NewLine, StringSplitOptions.RemoveEmptyEntries)
        .Single(line => line.StartsWith("RESULT|", StringComparison.Ordinal));
    var result = WorkerResult.Parse(resultLine);
    var miniExcelAssembly = typeof(MiniExcel).Assembly;
    var miniExcelVersion = miniExcelAssembly
        .GetCustomAttribute<AssemblyInformationalVersionAttribute>()?.InformationalVersion
        ?? miniExcelAssembly.GetName().Version?.ToString()
        ?? "unknown";

    Console.WriteLine($"MiniExcel version:      {miniExcelVersion}");
    Console.WriteLine($"File:                  {filePath}");
    Console.WriteLine($"Rows enumerated:       {result.RowCount:N0}");
    Console.WriteLine($"Elapsed:               {result.ElapsedMilliseconds:N0} ms");
    Console.WriteLine($"Total managed allocated: {FormatBytes(result.AllocatedBytes)}");
    Console.WriteLine($"Peak working set:      {FormatBytes(peakWorkingSet)}");
    Console.WriteLine($"Peak private bytes:    {FormatBytes(peakPrivateBytes)}");
    Console.WriteLine();
    Console.WriteLine("Total managed allocated is cumulative allocation, not peak memory usage.");
}

static void RunWorker(string filePath)
{
    GC.Collect(GC.MaxGeneration, GCCollectionMode.Forced, blocking: true, compacting: true);
    var allocatedBefore = GC.GetTotalAllocatedBytes(precise: true);
    var stopwatch = Stopwatch.StartNew();
    long rowCount = 0;

    foreach (var _ in MiniExcel.Query(filePath))
        rowCount++;

    stopwatch.Stop();
    var allocatedBytes = GC.GetTotalAllocatedBytes(precise: true) - allocatedBefore;
    Console.WriteLine(FormattableString.Invariant(
        $"RESULT|{rowCount}|{stopwatch.ElapsedMilliseconds}|{allocatedBytes}"));
}

static string GetFilePath(string[] arguments)
{
    var fileIndex = Array.FindIndex(arguments,
        argument => argument.Equals("--file", StringComparison.OrdinalIgnoreCase));
    return fileIndex >= 0 && fileIndex + 1 < arguments.Length
        ? Path.GetFullPath(arguments[fileIndex + 1])
        : Path.Combine(AppContext.BaseDirectory, "Test1,000,000x10.xlsx");
}

static string FormatBytes(long bytes)
{
    const double bytesPerMegabyte = 1024 * 1024;
    return $"{bytes / bytesPerMegabyte:N1} MB";
}

internal readonly record struct WorkerResult(long RowCount, long ElapsedMilliseconds, long AllocatedBytes)
{
    public static WorkerResult Parse(string value)
    {
        var parts = value.Split('|');
        return new WorkerResult(
            long.Parse(parts[1], CultureInfo.InvariantCulture),
            long.Parse(parts[2], CultureInfo.InvariantCulture),
            long.Parse(parts[3], CultureInfo.InvariantCulture));
    }
}