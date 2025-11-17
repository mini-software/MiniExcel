```

BenchmarkDotNet v0.15.0, Linux Ubuntu 24.04.3 LTS (Noble Numbat)
AMD EPYC 7763, 1 CPU, 4 logical and 2 physical cores
.NET SDK 10.0.100
  [Host]   : .NET 10.0.0 (10.0.25.52411), X64 RyuJIT AVX2
  ShortRun : .NET 10.0.0 (10.0.25.52411), X64 RyuJIT AVX2


```
| Method                               | Mean     | StdDev   | Error    | Gen0        | Gen1        | Gen2      | Allocated |
|------------------------------------- |---------:|---------:|---------:|------------:|------------:|----------:|----------:|
| &#39;MiniExcel Create Xlsx&#39;              |  4.278 s | 0.0211 s | 0.3845 s | 251666.6667 |   1666.6667 | 1333.3333 |   3.92 GB |
| &#39;OpenXmlSdk Create Xlsx by DOM mode&#39; | 19.611 s | 0.0421 s | 0.7674 s | 307000.0000 | 306833.3333 | 3833.3333 |   6.22 GB |
| &#39;ClosedXml Create Xlsx&#39;              | 20.232 s | 0.1549 s | 2.8260 s | 195500.0000 |  54500.0000 | 4166.6667 |   4.48 GB |
| &#39;Epplus Create Xlsx&#39;                 | 22.058 s | 0.0051 s | 0.0935 s |  88333.3333 |  17000.0000 | 5333.3333 |   2.51 GB |
| &#39;NPOI Create Xlsx&#39;                   | 35.951 s | 0.2129 s | 3.8832 s | 962833.3333 | 447166.6667 | 3500.0000 |  16.82 GB |
