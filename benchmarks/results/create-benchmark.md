```

BenchmarkDotNet v0.15.0, Linux Ubuntu 24.04.4 LTS (Noble Numbat)
AMD EPYC 7763, 1 CPU, 4 logical and 2 physical cores
.NET SDK 10.0.300
  [Host]   : .NET 10.0.8 (10.0.826.23019), X64 RyuJIT AVX2
  ShortRun : .NET 10.0.8 (10.0.826.23019), X64 RyuJIT AVX2


```
| Method                               | Mean     | StdDev   | Error    | Gen0        | Gen1        | Gen2      | Allocated |
|------------------------------------- |---------:|---------:|---------:|------------:|------------:|----------:|----------:|
| &#39;MiniExcel Create Xlsx&#39;              |  4.171 s | 0.0050 s | 0.0911 s | 251666.6667 |   1666.6667 | 1333.3333 |   3.92 GB |
| &#39;ClosedXml Create Xlsx&#39;              | 20.286 s | 0.0318 s | 0.5803 s | 195500.0000 |  54500.0000 | 4166.6667 |   4.48 GB |
| &#39;OpenXmlSdk Create Xlsx by DOM mode&#39; | 20.581 s | 0.0079 s | 0.1434 s | 307000.0000 | 306833.3333 | 3833.3333 |   6.22 GB |
| &#39;Epplus Create Xlsx&#39;                 | 21.780 s | 0.0388 s | 0.7081 s |  88833.3333 |  17333.3333 | 5833.3333 |   2.51 GB |
| &#39;NPOI Create Xlsx&#39;                   | 38.620 s | 0.4143 s | 7.5591 s | 962833.3333 | 447333.3333 | 3500.0000 |  16.82 GB |
