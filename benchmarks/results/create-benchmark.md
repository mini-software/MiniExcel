```

BenchmarkDotNet v0.15.0, Linux Ubuntu 24.04.5 LTS (Noble Numbat)
AMD EPYC 9V74, 1 CPU, 4 logical and 2 physical cores
.NET SDK 11.0.100
  [Host]   : .NET 10.0.15 (10.0.1526.53008), X64 RyuJIT AVX-512F+CD+BW+DQ+VL+VBMI
  ShortRun : .NET 10.0.15 (10.0.1526.53008), X64 RyuJIT AVX-512F+CD+BW+DQ+VL+VBMI


```
| Method                               | Mean     | StdDev   | Error     | Gen0        | Gen1        | Gen2      | Allocated |
|------------------------------------- |---------:|---------:|----------:|------------:|------------:|----------:|----------:|
| &#39;MiniExcel Create Xlsx&#39;              |  3.530 s | 0.0101 s |  0.1834 s | 252333.3333 |   2333.3333 | 2000.0000 |   3.92 GB |
| &#39;ClosedXml Create Xlsx&#39;              | 16.626 s | 0.1346 s |  2.4554 s | 195500.0000 |  54500.0000 | 4166.6667 |   4.48 GB |
| &#39;Epplus Create Xlsx&#39;                 | 17.126 s | 0.0426 s |  0.7773 s |  88666.6667 |  17166.6667 | 5666.6667 |   2.51 GB |
| &#39;OpenXmlSdk Create Xlsx by DOM mode&#39; | 17.193 s | 0.0714 s |  1.3023 s | 307000.0000 | 306833.3333 | 3833.3333 |   6.22 GB |
| &#39;NPOI Create Xlsx&#39;                   | 34.607 s | 1.5350 s | 28.0050 s | 963333.3333 | 447666.6667 | 3833.3333 |  16.82 GB |
