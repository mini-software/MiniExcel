```

BenchmarkDotNet v0.15.0, Linux Ubuntu 24.04.3 LTS (Noble Numbat)
AMD EPYC 7763, 1 CPU, 4 logical and 2 physical cores
.NET SDK 10.0.100
  [Host]   : .NET 10.0.0 (10.0.25.52411), X64 RyuJIT AVX2
  ShortRun : .NET 10.0.0 (10.0.25.52411), X64 RyuJIT AVX2


```
| Method                               | Mean     | StdDev   | Error    | Gen0         | Gen1        | Gen2      | Allocated |
|------------------------------------- |---------:|---------:|---------:|-------------:|------------:|----------:|----------:|
| &#39;MiniExcel Template Generate&#39;        |  3.461 s | 0.0133 s | 0.2422 s |  308166.6667 |   1500.0000 |  166.6667 |    4.8 GB |
| &#39;ClosedXml.Report Template Generate&#39; | 62.902 s | 0.5341 s | 9.7444 s | 1549000.0000 | 496000.0000 | 6833.3333 |  25.77 GB |
