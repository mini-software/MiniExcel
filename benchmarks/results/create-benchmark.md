```

BenchmarkDotNet v0.15.0, Linux Ubuntu 24.04.4 LTS (Noble Numbat)
AMD EPYC 9V74, 1 CPU, 4 logical and 2 physical cores
.NET SDK 10.0.301
  [Host]   : .NET 10.0.9 (10.0.926.27113), X64 RyuJIT AVX2
  ShortRun : .NET 10.0.9 (10.0.926.27113), X64 RyuJIT AVX2


```
| Method                               | Mean     | StdDev   | Error    | Gen0        | Gen1        | Gen2      | Allocated |
|------------------------------------- |---------:|---------:|---------:|------------:|------------:|----------:|----------:|
| &#39;MiniExcel Create Xlsx&#39;              |  4.238 s | 0.0030 s | 0.0549 s | 251500.0000 |   1500.0000 | 1166.6667 |   3.92 GB |
| &#39;OpenXmlSdk Create Xlsx by DOM mode&#39; | 18.494 s | 0.0451 s | 0.8223 s | 307000.0000 | 306833.3333 | 3833.3333 |   6.22 GB |
| &#39;ClosedXml Create Xlsx&#39;              | 20.486 s | 0.1200 s | 2.1886 s | 195500.0000 |  54500.0000 | 4166.6667 |   4.48 GB |
| &#39;Epplus Create Xlsx&#39;                 | 21.503 s | 0.0451 s | 0.8230 s |  88333.3333 |  17000.0000 | 5333.3333 |   2.51 GB |
| &#39;NPOI Create Xlsx&#39;                   | 38.968 s | 0.1316 s | 2.4013 s | 963666.6667 | 448666.6667 | 4333.3333 |  16.82 GB |
