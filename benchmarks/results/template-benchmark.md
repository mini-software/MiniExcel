```

BenchmarkDotNet v0.15.0, Linux Ubuntu 24.04.4 LTS (Noble Numbat)
AMD EPYC 7763, 1 CPU, 4 logical and 2 physical cores
.NET SDK 10.0.202
  [Host]   : .NET 10.0.6 (10.0.626.17701), X64 RyuJIT AVX2
  ShortRun : .NET 10.0.6 (10.0.626.17701), X64 RyuJIT AVX2


```
| Method                               | Mean     | StdDev   | Error    | Gen0         | Gen1        | Gen2      | Allocated |
|------------------------------------- |---------:|---------:|---------:|-------------:|------------:|----------:|----------:|
| &#39;MiniExcel Template Generate&#39;        |  4.347 s | 0.0057 s | 0.1036 s |  349666.6667 |   1666.6667 |  166.6667 |   5.45 GB |
| &#39;ClosedXml.Report Template Generate&#39; | 65.624 s | 0.3025 s | 5.5178 s | 1581833.3333 | 485166.6667 | 6166.6667 |   26.3 GB |
