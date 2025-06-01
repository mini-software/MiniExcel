```

BenchmarkDotNet v0.15.0, Linux Ubuntu 24.04.2 LTS (Noble Numbat)
AMD EPYC 7763, 1 CPU, 4 logical and 2 physical cores
.NET SDK 9.0.300
  [Host]   : .NET 9.0.5 (9.0.525.21509), X64 RyuJIT AVX2
  ShortRun : .NET 9.0.5 (9.0.525.21509), X64 RyuJIT AVX2

MiniExcel 1.41.2
OpenXmlSdk 3.3.0
ClosedXml 0.105.0
Epplus 7.7.2

```
| Method                                       | *Highest  Point Memory |     Mean |   StdDev |    Error |         Gen0 |        Gen1 |      Gen2 | Allocated |
| -------------------------------------------- | ---------------------- | -------: | -------: | -------: | -----------: | ----------: | --------: | --------: |
| &#39;MiniExcel Template Generate&#39;        | 24 MB                  |  3.340 s | 0.0235 s | 0.4295 s |  220166.6667 |   1000.0000 |         - |   3.43 GB |
| &#39;ClosedXml.Report Template Generate&#39; | 3156 MB                | 70.884 s | 0.4144 s | 7.5599 s | 1584833.3333 | 508500.0000 | 6500.0000 |  26.34 GB |
