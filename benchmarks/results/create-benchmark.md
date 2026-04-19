```

BenchmarkDotNet v0.15.0, Linux Ubuntu 24.04.4 LTS (Noble Numbat)
AMD EPYC 7763, 1 CPU, 4 logical and 2 physical cores
.NET SDK 10.0.202
  [Host]   : .NET 10.0.6 (10.0.626.17701), X64 RyuJIT AVX2
  ShortRun : .NET 10.0.6 (10.0.626.17701), X64 RyuJIT AVX2


```
| Method                               | Mean     | StdDev   | Error    | Gen0        | Gen1        | Gen2      | Allocated |
|------------------------------------- |---------:|---------:|---------:|------------:|------------:|----------:|----------:|
| &#39;MiniExcel Create Xlsx&#39;              |  4.031 s | 0.0063 s | 0.1149 s | 251666.6667 |   1666.6667 | 1333.3333 |   3.92 GB |
| &#39;OpenXmlSdk Create Xlsx by DOM mode&#39; | 19.418 s | 0.1846 s | 3.3671 s | 307000.0000 | 306833.3333 | 3833.3333 |   6.22 GB |
| &#39;ClosedXml Create Xlsx&#39;              | 19.960 s | 0.0377 s | 0.6880 s | 195500.0000 |  54500.0000 | 4166.6667 |   4.48 GB |
| &#39;Epplus Create Xlsx&#39;                 | 21.278 s | 0.0393 s | 0.7165 s |  88333.3333 |  17000.0000 | 5333.3333 |   2.51 GB |
| &#39;NPOI Create Xlsx&#39;                   | 35.446 s | 0.0999 s | 1.8222 s | 962833.3333 | 447333.3333 | 3500.0000 |  16.82 GB |
