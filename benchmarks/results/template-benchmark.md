```

BenchmarkDotNet v0.15.0, Linux Ubuntu 24.04.4 LTS (Noble Numbat)
INTEL XEON PLATINUM 8573C 2.30GHz, 1 CPU, 4 logical and 2 physical cores
.NET SDK 10.0.400
  [Host]   : .NET 10.0.11 (10.0.1126.37416), X64 RyuJIT AVX-512F+CD+BW+DQ+VL+VBMI
  ShortRun : .NET 10.0.11 (10.0.1126.37416), X64 RyuJIT AVX-512F+CD+BW+DQ+VL+VBMI


```
| Method                               | Mean     | StdDev   | Error    | Gen0        | Gen1        | Gen2      | Allocated |
|------------------------------------- |---------:|---------:|---------:|------------:|------------:|----------:|----------:|
| &#39;MiniExcel Template Generate&#39;        |  3.142 s | 0.0389 s | 0.7093 s |  69833.3333 |    333.3333 |         - |   5.45 GB |
| &#39;ClosedXml.Report Template Generate&#39; | 55.832 s | 0.2313 s | 4.2191 s | 320833.3333 | 186500.0000 | 5833.3333 |  26.29 GB |
