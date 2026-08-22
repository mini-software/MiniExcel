```

BenchmarkDotNet v0.15.0, Linux Ubuntu 24.04.4 LTS (Noble Numbat)
AMD EPYC 9V74, 1 CPU, 4 logical and 2 physical cores
.NET SDK 10.0.400
  [Host]   : .NET 10.0.11 (10.0.1126.37416), X64 RyuJIT AVX2
  ShortRun : .NET 10.0.11 (10.0.1126.37416), X64 RyuJIT AVX2


```
| Method                               | Mean     | StdDev   | Error    | Gen0        | Gen1        | Gen2      | Allocated |
|------------------------------------- |---------:|---------:|---------:|------------:|------------:|----------:|----------:|
| &#39;MiniExcel Create Xlsx&#39;              |  4.233 s | 0.0508 s | 0.9270 s | 251500.0000 |   1500.0000 | 1166.6667 |   3.92 GB |
| &#39;OpenXmlSdk Create Xlsx by DOM mode&#39; | 18.473 s | 0.0197 s | 0.3585 s | 307000.0000 | 306833.3333 | 3833.3333 |   6.22 GB |
| &#39;ClosedXml Create Xlsx&#39;              | 20.504 s | 0.0660 s | 1.2038 s | 195500.0000 |  54500.0000 | 4166.6667 |   4.48 GB |
| &#39;Epplus Create Xlsx&#39;                 | 21.302 s | 0.0367 s | 0.6692 s |  88333.3333 |  17000.0000 | 5333.3333 |   2.51 GB |
| &#39;NPOI Create Xlsx&#39;                   | 37.339 s | 0.0616 s | 1.1245 s | 963833.3333 | 448166.6667 | 4333.3333 |  16.82 GB |
