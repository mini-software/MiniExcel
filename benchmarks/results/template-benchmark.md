```

BenchmarkDotNet v0.15.0, Linux Ubuntu 24.04.4 LTS (Noble Numbat)
AMD EPYC 9V74, 1 CPU, 4 logical and 2 physical cores
.NET SDK 10.0.301
  [Host]   : .NET 10.0.9 (10.0.926.27113), X64 RyuJIT AVX2
  ShortRun : .NET 10.0.9 (10.0.926.27113), X64 RyuJIT AVX2


```
| Method                               | Mean     | StdDev   | Error     | Gen0         | Gen1        | Gen2      | Allocated |
|------------------------------------- |---------:|---------:|----------:|-------------:|------------:|----------:|----------:|
| &#39;MiniExcel Template Generate&#39;        |  4.249 s | 0.0114 s |  0.2075 s |  349666.6667 |   1666.6667 |         - |   5.45 GB |
| &#39;ClosedXml.Report Template Generate&#39; | 68.699 s | 0.7211 s | 13.1555 s | 1582333.3333 | 485333.3333 | 6666.6667 |   26.3 GB |
