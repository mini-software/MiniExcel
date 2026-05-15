```

BenchmarkDotNet v0.15.0, Linux Ubuntu 24.04.4 LTS (Noble Numbat)
AMD EPYC 9V74, 1 CPU, 4 logical and 2 physical cores
.NET SDK 10.0.300
  [Host]   : .NET 10.0.8 (10.0.826.23019), X64 RyuJIT AVX2
  ShortRun : .NET 10.0.8 (10.0.826.23019), X64 RyuJIT AVX2


```
| Method                               | Mean     | StdDev   | Error    | Gen0         | Gen1        | Gen2      | Allocated |
|------------------------------------- |---------:|---------:|---------:|-------------:|------------:|----------:|----------:|
| &#39;MiniExcel Template Generate&#39;        |  4.234 s | 0.0184 s | 0.3361 s |  349666.6667 |   1666.6667 |  166.6667 |   5.45 GB |
| &#39;ClosedXml.Report Template Generate&#39; | 68.069 s | 0.2216 s | 4.0434 s | 1582166.6667 | 485500.0000 | 6500.0000 |   26.3 GB |
