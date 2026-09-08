# Benchmark Guide 

**Peak Memory Usage:**

The benchmark pipeline includes a custom `PeakMemoryDiagnoser`. It samples the
benchmark process during the actual workload phase of a separate diagnostic run
and adds `Peak Working Set` and `Peak Private Bytes` columns to the
BenchmarkDotNet reports. The regular `Allocated` column remains the cumulative
managed allocation per operation and must not be interpreted as peak memory
usage.

The peak values are sampled every 20 ms, so they are approximations. Run the
query benchmarks with:

```powershell
$env:BenchmarkMode = "Automatic"
$env:BenchmarkSection = "Query"
$env:BenchmarkPeakMemory = "true"
dotnet run --project benchmarks/MiniExcel.Benchmarks -c Release -f net10.0
```

Visual Studio `Diagnostic Tools` can still be used for a detailed memory trace:

![image-20250601142913055](https://github.com/user-attachments/assets/439abeb1-4856-4851-88d3-1eac1b223ed6)