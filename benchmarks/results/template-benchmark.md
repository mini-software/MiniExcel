```

BenchmarkDotNet v0.15.0, Linux Ubuntu 24.04.4 LTS (Noble Numbat)
AMD EPYC 7763, 1 CPU, 4 logical and 2 physical cores
.NET SDK 10.0.300
  [Host]   : .NET 10.0.8 (10.0.826.23019), X64 RyuJIT AVX2
  ShortRun : .NET 10.0.8 (10.0.826.23019), X64 RyuJIT AVX2


```
| Method                               | Mean     | StdDev   | Error    | Gen0         | Gen1        | Gen2      | Allocated |
|------------------------------------- |---------:|---------:|---------:|-------------:|------------:|----------:|----------:|
| &#39;MiniExcel Template Generate&#39;        |  4.305 s | 0.0113 s | 0.2056 s |  349666.6667 |   1666.6667 |  166.6667 |   5.45 GB |
| &#39;ClosedXml.Report Template Generate&#39; | 62.220 s | 0.3061 s | 5.5844 s | 1548500.0000 | 496166.6667 | 6333.3333 |  25.77 GB |
