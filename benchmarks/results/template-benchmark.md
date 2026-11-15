```

BenchmarkDotNet v0.15.0, Linux Ubuntu 24.04.5 LTS (Noble Numbat)
AMD EPYC 9V45, 1 CPU, 4 logical and 2 physical cores
.NET SDK 11.0.100
  [Host]   : .NET 10.0.15 (10.0.1526.53008), X64 RyuJIT AVX-512F+CD+BW+DQ+VL+VBMI
  ShortRun : .NET 10.0.15 (10.0.1526.53008), X64 RyuJIT AVX-512F+CD+BW+DQ+VL+VBMI


```
| Method                               | Mean     | StdDev   | Error     | Gen0         | Gen1        | Gen2      | Allocated |
|------------------------------------- |---------:|---------:|----------:|-------------:|------------:|----------:|----------:|
| &#39;MiniExcel Template Generate&#39;        |  2.313 s | 0.0370 s |  0.6755 s |  349666.6667 |   1833.3333 |  166.6667 |   5.45 GB |
| &#39;ClosedXml.Report Template Generate&#39; | 43.510 s | 0.7137 s | 13.0212 s | 1582166.6667 | 484833.3333 | 6833.3333 |  26.27 GB |
