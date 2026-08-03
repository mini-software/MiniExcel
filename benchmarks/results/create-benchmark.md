```

BenchmarkDotNet v0.15.8, Linux Ubuntu 24.04.4 LTS (Noble Numbat)
AMD EPYC 9V74 3.69GHz, 1 CPU, 4 logical and 2 physical cores
.NET SDK 10.0.301
  [Host]   : .NET 10.0.9 (10.0.9, 10.0.926.27113), X64 RyuJIT x86-64-v4
  ShortRun : .NET 10.0.9 (10.0.9, 10.0.926.27113), X64 RyuJIT x86-64-v4


```
| Method                                      | Mean       | StdDev   | Error    | Gen0        | Gen1       | Gen2      | Allocated  |
|-------------------------------------------- |-----------:|---------:|---------:|------------:|-----------:|----------:|-----------:|
| &#39;MiniExcel Create Xlsx&#39;                     |   595.5 ms |  7.03 ms | 128.2 ms |  27000.0000 |  1666.6667 | 1500.0000 |  425.84 MB |
| &#39;MiniExcel Create Xlsx with Simple Mapping&#39; |   605.3 ms |  5.79 ms | 105.7 ms |  24166.6667 |  2000.0000 | 2000.0000 |  371.47 MB |
| &#39;ClosedXml Create Xlsx&#39;                     | 1,277.5 ms | 10.92 ms | 199.2 ms |  21833.3333 |  8000.0000 | 2833.3333 |  401.22 MB |
| &#39;OpenXmlSdk Create Xlsx by DOM mode&#39;        | 1,695.4 ms |  6.48 ms | 118.2 ms |  45166.6667 | 45000.0000 | 3333.3333 |   762.9 MB |
| &#39;Epplus Create Xlsx&#39;                        | 1,800.8 ms | 26.90 ms | 490.7 ms |  12166.6667 |  6333.3333 | 4166.6667 |  288.25 MB |
| &#39;NPOI Create Xlsx&#39;                          | 3,005.6 ms | 15.12 ms | 275.8 ms | 103500.0000 | 49000.0000 | 3166.6667 | 1741.95 MB |
