
This project is part of the [.NET Foundation](https://dotnetfoundation.org/projects/project-detail/miniexcel) and operates under their code of conduct.

---

### Introduction

MiniExcel is simple and efficient to avoid OOM's .NET processing Excel tool.

At present, most popular frameworks need to load all the data into the memory to facilitate operation, but it will cause memory consumption problems. MiniExcel tries to use algorithm from a stream to reduce the original 1000 MB occupation to a few MB to avoid OOM(out of memory).

![image](https://user-images.githubusercontent.com/12729184/113086657-ab8bd000-9214-11eb-9563-c970ac1ee35e.png)


### Features

- Low memory consumption, avoid OOM (out of memory) and full GC
- Supports real time operation of each row of data
- Supports LINQ deferred execution, it can do low-consumption, fast paging and other complex queries
- Lightweight, without Microsoft Office installed, no COM+, DLL size is less than 400KB
- Easy API style to read/write/fill excel

### Get Started

- [Import/Query Excel](#getstart1)

- [Export/Create Excel](#getstart2)

- [Excel Template](#getstart3)

- [Excel Column Name/Index/Ignore Attribute](#getstart4)

- [Examples](#getstart5)



### Installation

You can install the package [from NuGet](https://www.nuget.org/packages/MiniExcel)

### Release Notes

Please Check [Release Notes](docs)

### TODO

Please Check  [TODO](https://github.com/mini-software/MiniExcel/projects/1?fullscreen=true)

### Performance

The code for the benchmarks can be found in [MiniExcel.Benchmarks](https://github.com/mini-software/MiniExcel/tree/master/benchmarks/MiniExcel.Benchmarks).
To run all the benchmarks use:

```bash
dotnet run -project .\benchmarks\MiniExcel.Benchmarks -c Release -f net9.0 -filter * --join
```

Hardware and settings used are the following:
```
BenchmarkDotNet v0.15.0, Linux Ubuntu 24.04.2 LTS (Noble Numbat)
AMD EPYC 7763, 1 CPU, 4 logical and 2 physical cores
.NET SDK 9.0.300
  [Host]   : .NET 9.0.5 (9.0.525.21509), X64 RyuJIT AVX2
  ShortRun : .NET 9.0.5 (9.0.525.21509), X64 RyuJIT AVX2
```

#### Import/Query Excel

The file used to test performance is [**Test1,000,000x10.xlsx**](https://github.com/mini-software/MiniExcel/tree/master/benchmarks/MiniExcel.Benchmarks/Test1%2C000%2C000x10.xlsx), a 32MB document containing 1,000,000 rows * 10 columns whose cells are filled with the string "HelloWorld".

| Method                               |             Mean |         StdDev |            Error |        Gen0 |        Gen1 |      Gen2 |     Allocated |
|--------------------------------------|-----------------:|---------------:|-----------------:|------------:|------------:|----------:|--------------:|
| &#39;MiniExcel QueryFirst&#39;       |         63.70 μs |       0.337 μs |         6.144 μs |      2.9297 |      2.7669 |         - |      49.67 KB |
| &#39;ExcelDataReader QueryFirst&#39; |  5,010,679.51 μs |  53,245.186 μs |   971,390.400 μs | 105000.0000 |    333.3333 |         - | 1717272.56 KB |
| &#39;MiniExcel Query&#39;            |  9,172,286.91 μs |  12,805.326 μs |   233,616.824 μs | 448500.0000 |   4666.6667 |         - | 7327883.36 KB |
| &#39;ExcelDataReader Query&#39;      | 10,609,617.09 μs |  29,055.953 μs |   530,088.745 μs | 275666.6667 |  68666.6667 |         - | 4504691.87 KB |
| &#39;Epplus QueryFirst&#39;          | 13,770,656.24 μs |  45,909.809 μs |   837,565.827 μs | 174333.3333 |  88833.3333 | 4333.3333 | 3700587.76 KB |
| &#39;Epplus Query&#39;               | 19,257,306.83 μs |  63,117.956 μs | 1,151,506.486 μs | 452333.3333 |  90500.0000 | 5333.3333 | 8223933.16 KB |
| &#39;ClosedXml Query&#39;            | 31,070,263.83 μs | 342,973.671 μs | 6,257,116.502 μs | 401666.6667 | 104166.6667 | 3333.3333 | 6822559.68 KB |
| &#39;ClosedXml QueryFirst&#39;       | 31,141,877.48 μs |  21,006.538 μs |   383,237.459 μs | 402166.6667 | 104833.3333 | 3833.3333 |  6738357.8 KB |
| &#39;OpenXmlSDK QueryFirst&#39;      | 31,750,686.63 μs | 263,328.569 μs | 4,804,093.357 μs | 374666.6667 | 374500.0000 | 3166.6667 | 6069266.96 KB |
| &#39;OpenXmlSDK Query&#39;           | 32,919,119.46 μs | 411,395.682 μs | 7,505,388.691 μs | 374666.6667 | 374500.0000 | 3166.6667 | 6078467.83 KB |


#### Export/Create Excel

Logic: create a total of 10,000,000 "HelloWorld" cells Excel document

| Method                                       |     Mean |   StdDev |    Error |        Gen0 |        Gen1 |      Gen2 | Allocated |
|----------------------------------------------|---------:|---------:|---------:|------------:|------------:|----------:|----------:|
| &#39;MiniExcel Create Xlsx&#39;              |  4.427 s | 0.0056 s | 0.1023 s | 251666.6667 |   1833.3333 | 1666.6667 |   3.92 GB |
| &#39;OpenXmlSdk Create Xlsx by DOM mode&#39; | 22.729 s | 0.1226 s | 2.2374 s | 307000.0000 | 306833.3333 | 3833.3333 |   6.22 GB |
| &#39;ClosedXml Create Xlsx&#39;              | 22.851 s | 0.0190 s | 0.3473 s | 195500.0000 |  54500.0000 | 4166.6667 |   4.48 GB |
| &#39;Epplus Create Xlsx&#39;                 | 23.027 s | 0.0088 s | 0.1596 s |  89000.0000 |  17500.0000 | 6000.0000 |   2.51 GB |

Warning: these results may be outdated. You can find the benchmarks for the latest release [here](https://github.com/mini-software/MiniExcel/tree/master/benchmarks/results).


### Documents

https://github.com/mini-software/MiniExcel
