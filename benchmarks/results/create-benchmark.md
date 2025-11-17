```

BenchmarkDotNet v0.15.6, Linux Ubuntu 24.04.3 LTS (Noble Numbat)
AMD EPYC 7763 2.45GHz, 1 CPU, 4 logical and 2 physical cores
.NET SDK 10.0.100
  [Host]   : .NET 10.0.0 (10.0.0, 10.0.25.52411), X64 RyuJIT x86-64-v3
  ShortRun : .NET 10.0.0 (10.0.0, 10.0.25.52411), X64 RyuJIT x86-64-v3


```
| Method                                      | Mean     | StdDev   | Error    | Gen0        | Gen1        | Gen2      | Allocated |
|-------------------------------------------- |---------:|---------:|---------:|------------:|------------:|----------:|----------:|
| &#39;MiniExcel Create Xlsx with Simple Mapping&#39; |  3.590 s | 0.0282 s | 0.5148 s | 213333.3333 |   1333.3333 | 1166.6667 |   3.32 GB |
| &#39;MiniExcel Create Xlsx&#39;                     |  4.503 s | 0.0137 s | 0.2506 s | 251666.6667 |   1666.6667 | 1333.3333 |   3.92 GB |
| &#39;ClosedXml Create Xlsx&#39;                     | 19.846 s | 0.0988 s | 1.8023 s | 195500.0000 |  54500.0000 | 4166.6667 |   4.48 GB |
| &#39;OpenXmlSdk Create Xlsx by DOM mode&#39;        | 19.966 s | 0.0908 s | 1.6574 s | 307000.0000 | 306833.3333 | 3833.3333 |   6.22 GB |
| &#39;Epplus Create Xlsx&#39;                        | 21.634 s | 0.0503 s | 0.9169 s |  88333.3333 |  17000.0000 | 5333.3333 |   2.51 GB |
| &#39;NPOI Create Xlsx&#39;                          | 36.795 s | 0.3467 s | 6.3256 s | 963500.0000 | 447833.3333 | 4000.0000 |  16.82 GB |
