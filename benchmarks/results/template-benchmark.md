```

BenchmarkDotNet v0.15.6, Linux Ubuntu 24.04.3 LTS (Noble Numbat)
Intel Xeon Platinum 8370C CPU 2.80GHz (Max: 3.36GHz), 1 CPU, 4 logical and 2 physical cores
.NET SDK 10.0.100
  [Host]   : .NET 10.0.0 (10.0.0, 10.0.25.52411), X64 RyuJIT x86-64-v4
  ShortRun : .NET 10.0.0 (10.0.0, 10.0.25.52411), X64 RyuJIT x86-64-v4


```
| Method                                | Mean          | StdDev        | Error          | Gen0         | Gen1        | Gen2      | Allocated   |
|-------------------------------------- |--------------:|--------------:|---------------:|-------------:|------------:|----------:|------------:|
| &#39;MiniExcel Mapping Template Generate&#39; |      2.941 ms |     0.0528 ms |      0.9636 ms |    2312.5000 |   2302.0833 | 2302.0833 |    15.28 MB |
| &#39;MiniExcel Template Generate&#39;         |  3,234.435 ms |    15.5778 ms |    284.1962 ms |  217833.3333 |   1166.6667 |  166.6667 |  5212.93 MB |
| &#39;ClosedXml.Report Template Generate&#39;  | 62,563.080 ms | 1,186.0532 ms | 21,638.0257 ms | 1034333.3333 | 377500.0000 | 6333.3333 | 26397.81 MB |
