```

BenchmarkDotNet v0.15.8, Linux Ubuntu 24.04.4 LTS (Noble Numbat)
AMD EPYC 9V74 2.86GHz, 1 CPU, 4 logical and 2 physical cores
.NET SDK 10.0.301
  [Host]   : .NET 10.0.9 (10.0.9, 10.0.926.27113), X64 RyuJIT x86-64-v3
  ShortRun : .NET 10.0.9 (10.0.9, 10.0.926.27113), X64 RyuJIT x86-64-v3


```
| Method                               | Mean         | StdDev     | Error        | Gen0        | Gen1       | Gen2      | Allocated  |
|------------------------------------- |-------------:|-----------:|-------------:|------------:|-----------:|----------:|-----------:|
| &#39;MiniExcel Mapping Fill Template&#39;    |     3.261 ms |  0.2537 ms |     4.628 ms |   2348.9583 |  2328.1250 | 2322.9167 |   15.41 MB |
| &#39;MiniExcel Fill Template&#39;            |   460.505 ms |  3.6797 ms |    67.132 ms |  37333.3333 |   333.3333 |         - |  596.96 MB |
| &#39;ClosedXml.Report Generate Template&#39; | 5,047.767 ms | 64.9253 ms | 1,184.479 ms | 169833.3333 | 56166.6667 | 6833.3333 | 2799.89 MB |
