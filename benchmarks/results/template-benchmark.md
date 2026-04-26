```

BenchmarkDotNet v0.15.8, Linux Ubuntu 24.04.4 LTS (Noble Numbat)
AMD EPYC 9V74 2.87GHz, 1 CPU, 4 logical and 2 physical cores
.NET SDK 10.0.203
  [Host]   : .NET 10.0.7 (10.0.7, 10.0.726.21808), X64 RyuJIT x86-64-v3
  ShortRun : .NET 10.0.7 (10.0.7, 10.0.726.21808), X64 RyuJIT x86-64-v3


```
| Method                               | Mean         | StdDev     | Error       | Gen0        | Gen1       | Gen2      | Allocated  |
|------------------------------------- |-------------:|-----------:|------------:|------------:|-----------:|----------:|-----------:|
| &#39;MiniExcel Mapping Fill Template&#39;    |     3.663 ms |  0.0304 ms |   0.5550 ms |   2510.4167 |  2489.5833 | 2484.3750 |    15.4 MB |
| &#39;MiniExcel Fill Template&#39;            |   425.232 ms |  4.6240 ms |  84.3592 ms |  35166.6667 |   333.3333 |         - |  563.19 MB |
| &#39;ClosedXml.Report Generate Template&#39; | 5,070.920 ms | 50.5414 ms | 922.0625 ms | 168500.0000 | 55666.6667 | 5833.3333 | 2799.89 MB |
