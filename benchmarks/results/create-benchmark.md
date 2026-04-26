```

BenchmarkDotNet v0.15.8, Linux Ubuntu 24.04.4 LTS (Noble Numbat)
AMD EPYC 9V74 2.60GHz, 1 CPU, 4 logical and 2 physical cores
.NET SDK 10.0.203
  [Host]   : .NET 10.0.7 (10.0.7, 10.0.726.21808), X64 RyuJIT x86-64-v3
  ShortRun : .NET 10.0.7 (10.0.7, 10.0.726.21808), X64 RyuJIT x86-64-v3


```
| Method                                      | Mean       | StdDev    | Error      | Gen0       | Gen1       | Gen2      | Allocated  |
|-------------------------------------------- |-----------:|----------:|-----------:|-----------:|-----------:|----------:|-----------:|
| &#39;MiniExcel Create Xlsx with Simple Mapping&#39; |   732.7 ms |   5.87 ms |   107.0 ms | 23666.6667 |  1333.3333 | 1333.3333 |  372.16 MB |
| &#39;MiniExcel Create Xlsx&#39;                     |   789.0 ms |   6.05 ms |   110.4 ms | 27166.6667 |  1333.3333 | 1166.6667 |   432.7 MB |
| &#39;ClosedXml Create Xlsx&#39;                     | 1,621.7 ms |  12.02 ms |   219.3 ms | 21833.3333 |  8000.0000 | 2833.3333 |  401.23 MB |
| &#39;OpenXmlSdk Create Xlsx by DOM mode&#39;        | 2,210.4 ms |  35.61 ms |   649.7 ms | 45166.6667 | 45000.0000 | 3333.3333 |   762.9 MB |
| &#39;Epplus Create Xlsx&#39;                        | 2,316.7 ms |   9.48 ms |   173.0 ms | 12166.6667 |  6333.3333 | 4166.6667 |  288.25 MB |
| &#39;NPOI Create Xlsx&#39;                          | 3,991.9 ms | 149.43 ms | 2,726.1 ms | 98833.3333 | 48500.0000 | 3000.0000 | 1748.04 MB |
