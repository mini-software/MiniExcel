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
| Method                                       | *Highest  Point Memory |     Mean |   StdDev |    Error |        Gen0 |        Gen1 |      Gen2 | Allocated |
| -------------------------------------------- | ---------------------- | -------: | -------: | -------: | ----------: | ----------: | --------: | --------: |
| &#39;MiniExcel Create Xlsx&#39;              | 33 MB                  |  4.347 s | 0.0114 s | 0.2078 s | 251666.6667 |   1833.3333 | 1666.6667 |   3.92 GB |
| &#39;OpenXmlSdk Create Xlsx by DOM mode&#39; | 4126MB                 | 22.359 s | 0.1548 s | 2.8238 s | 307000.0000 | 306833.3333 | 3833.3333 |   6.22 GB |
| &#39;ClosedXml Create Xlsx&#39;              | 525 MB                 | 22.707 s | 0.0647 s | 1.1801 s | 195500.0000 |  54666.6667 | 4166.6667 |   4.48 GB |
| &#39;Epplus Create Xlsx&#39;                 | 617 MB                 | 22.822 s | 0.0275 s | 0.5017 s |  89000.0000 |  17500.0000 | 6000.0000 |   2.51 GB |
