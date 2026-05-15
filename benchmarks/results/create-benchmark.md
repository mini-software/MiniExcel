```

BenchmarkDotNet v0.15.0, Linux Ubuntu 24.04.4 LTS (Noble Numbat)
AMD EPYC 7763, 1 CPU, 4 logical and 2 physical cores
.NET SDK 10.0.300
  [Host]   : .NET 10.0.8 (10.0.826.23019), X64 RyuJIT AVX2
  ShortRun : .NET 10.0.8 (10.0.826.23019), X64 RyuJIT AVX2


```
| Method                               | Mean     | StdDev   | Error    | Gen0        | Gen1        | Gen2      | Allocated |
|------------------------------------- |---------:|---------:|---------:|------------:|------------:|----------:|----------:|
| &#39;MiniExcel Create Xlsx&#39;              |  3.961 s | 0.0292 s | 0.5327 s | 251666.6667 |   1666.6667 | 1333.3333 |   3.92 GB |
| &#39;OpenXmlSdk Create Xlsx by DOM mode&#39; | 19.379 s | 0.0468 s | 0.8530 s | 307000.0000 | 306833.3333 | 3833.3333 |   6.22 GB |
| &#39;ClosedXml Create Xlsx&#39;              | 19.950 s | 0.1189 s | 2.1691 s | 195500.0000 |  54500.0000 | 4166.6667 |   4.48 GB |
| &#39;Epplus Create Xlsx&#39;                 | 21.761 s | 0.0049 s | 0.0897 s |  88333.3333 |  17000.0000 | 5333.3333 |   2.51 GB |
| &#39;NPOI Create Xlsx&#39;                   | 35.261 s | 0.1067 s | 1.9475 s | 962833.3333 | 447166.6667 | 3500.0000 |  16.82 GB |
