
This project is part of the [.NET Foundation](https://dotnetfoundation.org/projects/project-detail/miniexcel) and operates under their code of conduct.

---

### Introduction

MiniExcel is simple and efficient to avoid OOM's .NET processing Excel tool.

At present, most popular frameworks need to load all the data into the memory to facilitate operation, but it will cause memory consumption problems. MiniExcel tries to use algorithm from a stream to reduce the original 1000 MB occupation to a few MB to avoid OOM(out of memory).

![image](https://user-images.githubusercontent.com/12729184/113086657-ab8bd000-9214-11eb-9563-c970ac1ee35e.png)


### Features

- Low memory consumption, avoid OOM (out of memory) and full GC
- Support `real-time` operation of each row of data
- Support LINQ deferred execution, it can do low-consumption, fast paging and other complex queries
- Lightweight, without Microsoft Office installed, no COM+, DLL size is less than 150KB
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

Please Check  [TODO](https://github.com/shps951023/MiniExcel/projects/1?fullscreen=true)

### Performance

Benchmarks logic can be found in  [MiniExcel.Benchmarks](benchmarks/MiniExcel.Benchmarks/Program.cs) , and test cli

```bash
dotnet run -p .\benchmarks\MiniExcel.Benchmarks\ -c Release -f netcoreapp3.1 -- -f * --join
```

Output from the latest run is :

```bash
BenchmarkDotNet=v0.12.1, OS=Windows 10.0.19042
Intel Core i7-7700 CPU 3.60GHz (Kaby Lake), 1 CPU, 8 logical and 4 physical cores
  [Host]     : .NET Framework 4.8 (4.8.4341.0), X64 RyuJIT
  Job-ZYYABG : .NET Framework 4.8 (4.8.4341.0), X64 RyuJIT
IterationCount=3  LaunchCount=3  WarmupCount=3
```

Benchmark History :  [Link](https://github.com/shps951023/MiniExcel/issues/276)



#### Import/Query Excel

Logic : [**Test1,000,000x10.xlsx**](benchmarks/MiniExcel.Benchmarks/Test1%2C000%2C000x10.xlsx)  as performance test basic file, 1,000,000 rows * 10 columns  "HelloWorld" cells, 23 MB file size


| Library      | Method                       | Max Memory Usage |         Mean |
| ---------------------------- | -------------: | ---------------: | ---------------: |
| MiniExcel | 'MiniExcel QueryFirst'       |       0.109 MB | 0.0007264 sec |
| ExcelDataReader | 'ExcelDataReader QueryFirst' |       15.24 MB | 10.66421 sec |
| MiniExcel  | 'MiniExcel Query'            |        17.3 MB | 14.17933 sec |
| ExcelDataReader | 'ExcelDataReader Query'      |        17.3 MB | 22.56508 sec |
| Epplus    | 'Epplus QueryFirst'          |       1,452 MB | 18.19801 sec |
| Epplus        | 'Epplus Query'               |       1,451 MB | 23.64747 sec |
| OpenXmlSDK | 'OpenXmlSDK Query'           |       1,412 MB | 52.00327 sec |
| OpenXmlSDK | 'OpenXmlSDK QueryFirst'      |       1,413 MB | 52.34865 sec |
| ClosedXml | 'ClosedXml QueryFirst'       |       2,158 MB | 66.18897 sec |
| ClosedXml  | 'ClosedXml Query'            |       2,184 MB | 191.43412 sec |

#### Export/Create Excel

Logic : create a total of 10,000,000 "HelloWorld" excel

| Library            | Method                   | Max Memory Usage |         Mean |
| ------------------------ | -------------: | ---------------: | -----------: |
| MiniExcel | 'MiniExcel Create Xlsx'  |          15 MB | 11.53181 sec |
| Epplus | 'Epplus Create Xlsx'     |       1,204 MB | 22.50971 sec |
| OpenXmlSdk | 'OpenXmlSdk Create Xlsx' |       2,621 MB | 42.47399 sec |
| ClosedXml | 'ClosedXml Create Xlsx'  |       7,141 MB | 140.93992 sec |



### Documents

https://github.com/mini-software/MiniExcel
