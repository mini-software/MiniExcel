[![NuGet](https://img.shields.io/nuget/v/MiniExcel.svg)](https://www.nuget.org/packages/MiniExcel)  [![](https://img.shields.io/nuget/dt/MiniExcel.svg)](https://www.nuget.org/packages/MiniExcel)  [![Build status](https://ci.appveyor.com/api/projects/status/b2vustrwsuqx45f4/branch/master?svg=true)](https://ci.appveyor.com/project/shps951023/miniexcel/branch/master) [![.NET Framework](https://img.shields.io/badge/.NET%20Framework-%3E%3D%204.6.1-red.svg)](#)  [![.NET Standard](https://img.shields.io/badge/.NET%20Standard-%3E%3D%202.0-red.svg)](#) [![.NET](https://img.shields.io/badge/.NET%20-%3E%3D%205.0-red.svg)](#) 

---

[English](https://github.com/shps951023/MiniExcel) / [繁體中文](https://github.com/shps951023/MiniExcel/blob/master/README.zh-tw.md) / [简体中文](https://github.com/shps951023/MiniExcel/blob/master/README.zh-Hans.md)

---

### Introduction

MiniExcel is simple and efficient to avoid OOM's .NET processing Excel tool.

At present, most popular frameworks need to load all the data into the memory to facilitate operation, but it will cause memory consumption problems. MiniExcel tries to use algorithm from a stream to reduce the original 1000 MB occupation to a few MB to avoid OOM(out of memory).

![image](https://user-images.githubusercontent.com/12729184/113086657-ab8bd000-9214-11eb-9563-c970ac1ee35e.png)

### Features
- Low memory consumption, avoid OOM (out of memory) and full GC
- Support `real-time` operation of each row of data
  ![miniexcel_lazy_load](https://user-images.githubusercontent.com/12729184/111034290-e5588a80-844f-11eb-8c84-6fdb6fb8f403.gif)
- Support LINQ deferred execution, it can do low-consumption, fast paging and other complex queries
  ![queryfirst](https://user-images.githubusercontent.com/12729184/111072392-6037a900-8515-11eb-9693-5ce2dad1e460.gif)
- Lightweight, does not with any third-party dependencies, DLL is less than 100KB
- Easy Dapper API style

### Installation

You can install the package [from NuGet](https://www.nuget.org/packages/MiniExcel)

### Release Notes

Please Check [Release Notes](https://github.com/shps951023/MiniExcel/tree/master/docs)

### Discussions / TODO 

Please Check [Discussions](https://github.com/shps951023/MiniExcel/discussions) / [TODO](https://github.com/shps951023/MiniExcel/projects/1?fullscreen=true)

### Performance

[**Test1,000,000x10.xlsx**](https://github.com/shps951023/MiniExcel/blob/master/samples/xlsx/Test1%2C000%2C000x10/Test1%2C000%2C000x10.xlsx) as performance test basic file,A total of 10,000,000 "HelloWorld" with a file size of 23 MB

Benchmarks  logic can be found in  [MiniExcel.Benchmarks](https://github.com/shps951023/MiniExcel/tree/master/benchmarks/MiniExcel.Benchmarks) , and test cli

```
dotnet run -p .\benchmarks\MiniExcel.Benchmarks\ -c Release -f netcoreapp3.1 -- -f * --join
```

Output from the latest run is :  

```
BenchmarkDotNet=v0.12.1, OS=Windows 10.0.19042
Intel Core i7-7700 CPU 3.60GHz (Kaby Lake), 1 CPU, 8 logical and 4 physical cores
  [Host]     : .NET Framework 4.8 (4.8.4341.0), X64 RyuJIT
  Job-ZYYABG : .NET Framework 4.8 (4.8.4341.0), X64 RyuJIT
IterationCount=3  LaunchCount=3  WarmupCount=3  
```

| Method                       | Max Memory Usage |             Mean |        Gen 0 |       Gen 1 |      Gen 2 |
| ---------------------------- | ---------------: | ---------------: | -----------: | ----------: | ---------: |
| 'MiniExcel QueryFirst'       |         0.109 MB |         726.4 us |            - |           - |          - |
| 'ExcelDataReader QueryFirst' |         15.24 MB |  10,664,238.2 us |  566000.0000 |   1000.0000 |          - |
| 'MiniExcel Query'            |          17.3 MB |  14,179,334.8 us |  367000.0000 |  96000.0000 |  7000.0000 |
| 'ExcelDataReader Query'      |          17.3 MB |  22,565,088.7 us | 1210000.0000 |   2000.0000 |          - |
| 'Epplus QueryFirst'          |         1,452 MB |  18,198,015.4 us |  535000.0000 | 132000.0000 |  9000.0000 |
| 'Epplus Query'               |         1,451 MB |  23,647,471.1 us | 1451000.0000 | 133000.0000 |  9000.0000 |
| 'OpenXmlSDK Query'           |         1,412 MB |  52,003,270.1 us |  978000.0000 | 353000.0000 | 11000.0000 |
| 'OpenXmlSDK QueryFirst'      |         1,413 MB |  52,348,659.1 us |  978000.0000 | 353000.0000 | 11000.0000 |
| 'ClosedXml QueryFirst'       |         2,158 MB |  66,188,979.6 us | 2156000.0000 | 575000.0000 |  9000.0000 |
| 'ClosedXml Query'            |         2,184 MB | 191,434,126.6 us | 2165000.0000 | 577000.0000 | 10000.0000 |


| Method                   | Max Memory Usage |             Mean |        Gen 0 |        Gen 1 |      Gen 2 |
| ------------------------ | ---------------: | ---------------: | -----------: | -----------: | ---------: |
| 'MiniExcel Create Xlsx'  |            15 MB |  11,531,819.8 us | 1020000.0000 |            - |          - |
| 'Epplus Create Xlsx'     |         1,204 MB |  22,509,717.7 us | 1370000.0000 |   60000.0000 | 30000.0000 |
| 'OpenXmlSdk Create Xlsx' |         2,621 MB |  42,473,998.9 us | 1370000.0000 |  460000.0000 | 50000.0000 |
| 'ClosedXml Create Xlsx'  |         7,141 MB | 140,939,928.6 us | 5520000.0000 | 1500000.0000 | 80000.0000 |




### Execute a query and map the results to a strongly typed IEnumerable [[Try it]](https://dotnetfiddle.net/w5WD1J)

Recommand to use Stream.Query because of better efficiency.

```C#
public class UserAccount
{
    public Guid ID { get; set; }
    public string Name { get; set; }
    public DateTime BoD { get; set; }
    public int Age { get; set; }
    public bool VIP { get; set; }
    public decimal Points { get; set; }
}

var rows = MiniExcel.Query<UserAccount>(path);

// or

using (var stream = File.OpenRead(path))
    var rows = stream.Query<UserAccount>();
```

![image](https://user-images.githubusercontent.com/12729184/111107423-c8c46b80-8591-11eb-982f-c97a2dafb379.png)


### Execute a query and map it to a list of dynamic objects without using head [[Try it]](https://dotnetfiddle.net/w5WD1J)

* dynamic key is `A.B.C.D..`

| MiniExcel     | 1     |
| -------- | -------- |
| Github     | 2     |

```C#

var rows = MiniExcel.Query(path).ToList();

// or 
using (var stream = File.OpenRead(path))
{
    var rows = stream.Query().ToList();
                
    Assert.Equal("MiniExcel", rows[0].A);
    Assert.Equal(1, rows[0].B);
    Assert.Equal("Github", rows[1].A);
    Assert.Equal(2, rows[1].B);
}
```

### Execute a query with first header row [[Try it]](https://dotnetfiddle.net/w5WD1J)

note : same column name use last right one 

Input Excel :  

| Column1 | Column2 |
| -------- | -------- |
| MiniExcel     | 1     |
| Github     | 2     |


```C#

var rows = MiniExcel.Query(useHeaderRow:true).ToList();

// or

using (var stream = File.OpenRead(path))
{
    var rows = stream.Query(useHeaderRow:true).ToList();

    Assert.Equal("MiniExcel", rows[0].Column1);
    Assert.Equal(1, rows[0].Column2);
    Assert.Equal("Github", rows[1].Column1);
    Assert.Equal(2, rows[1].Column2);
}
```

### Query Support LINQ Extension First/Take/Skip ...etc

Query First
```C#
var row = MiniExcel.Query(path).First();
Assert.Equal("HelloWorld", row.A);

// or

using (var stream = File.OpenRead(path))
{
    var row = stream.Query().First();
    Assert.Equal("HelloWorld", row.A);
}
```

Performance between MiniExcel/ExcelDataReader/ClosedXML/EPPlus  
![queryfirst](https://user-images.githubusercontent.com/12729184/111072392-6037a900-8515-11eb-9693-5ce2dad1e460.gif)


### Create Excel file [[Try it]](https://dotnetfiddle.net/w5WD1J)

1. Must be a non-abstract type with a public parameterless constructor .

2. MiniExcel support parameter IEnumerable Deferred Execution, If you want to use least memory, please do not call methods such as ToList

e.g : ToList or not memory usage  
![image](https://user-images.githubusercontent.com/12729184/112587389-752b0b00-8e38-11eb-8a52-cfb76c57e5eb.png)



Anonymous or strongly type: 
```C#
var path = Path.Combine(Path.GetTempPath(), $"{Guid.NewGuid()}.xlsx");
MiniExcel.SaveAs(path, new[] {
    new { Column1 = "MiniExcel", Column2 = 1 },
    new { Column1 = "Github", Column2 = 2}
});
```

Datatable:  
```C#
var path = Path.Combine(Path.GetTempPath(), $"{Guid.NewGuid()}.xlsx");
var table = new DataTable();
{
    table.Columns.Add("Column1", typeof(string));
    table.Columns.Add("Column2", typeof(decimal));
    table.Rows.Add("MiniExcel", 1);
    table.Rows.Add("Github", 2);
}

MiniExcel.SaveAs(path, table);
```

Dapper:  
```C#
using (var connection = GetConnection(connectionString))
{
    var rows = connection.Query(@"select 'MiniExcel' as Column1,1 as Column2 union all select 'Github',2");
    MiniExcel.SaveAs(path, rows);
}
```

`IEnumerable<IDictionary<string, object>>`
```C#
var values = new List<Dictionary<string, object>>()
{
    new Dictionary<string,object>{{ "Column1", "MiniExcel" }, { "Column2", 1 } },
    new Dictionary<string,object>{{ "Column1", "Github" }, { "Column2", 2 } }
};
MiniExcel.SaveAs(path, values);
```

Create File Result : 

| Column1 | Column2 |
| -------- | -------- |
| MiniExcel     | 1     |
| Github     | 2     |

### SaveAs Stream [[Try it]](https://dotnetfiddle.net/JOen0e)

```C#
using (var stream = File.Create(path))
{
    stream.SaveAs(values);
}
```



### Excel Column Name/Ignore Attribute

e.g

input excel :  

| Test1 | Test2 | Test3 | Test4 | Test5 | Test6 | Column1 | Column2 |
| ----- | ----- | ----- | ----- | ----- | ----- | ------- | ------- |
| Test1 | Test2 | Test3 | Test4 | Test5 | Test6 | Column1 | Column2 |

```C#
public class ExcelAttributeDemo
{
    [ExcelColumnName("Column1")]
    public string Test1 { get; set; }
    [ExcelColumnName("Column2")]
    public string Test2 { get; set; }
    [ExcelIgnore]
    public string Test3 { get; set; }
    public string Test4 { get; set; }
    public string Test5 { get; }
    public string Test6 { get; private set; }
}

var rows = MiniExcel.Query<ExcelAttributeDemo>(path).ToList();
Assert.Equal("Column1", rows[0].Test1);
Assert.Equal("Column2", rows[0].Test2);
Assert.Null(rows[0].Test3);
Assert.Equal("Test4", rows[0].Test4);
Assert.Null(rows[0].Test5);
Assert.Null(rows[0].Test6);
```



### SQLite & Dapper `Large Size File` SQL Insert Avoid OOM

note : please don't call ToList/ToArray methods after Query, it'll load all data into memory

```C#
using (var connection = new SQLiteConnection(connectionString))
{
    connection.Open();
    using (var transaction = connection.BeginTransaction())
    using (var stream = File.OpenRead(path))
    {
	   var rows = stream.Query();
	   foreach (var row in rows)
			 connection.Execute("insert into T (A,B) values (@A,@B)", new { row.A, row.B }, transaction: transaction);
	   transaction.Commit();
    }
}
```

performance:
![image](https://user-images.githubusercontent.com/12729184/111072579-2dda7b80-8516-11eb-9843-c01a1edc88ec.png)


### ASP.NET Core 3.1 or MVC 5 Download Excel Xlsx API Demo

```C#
public class ExcelController : Controller
{
    public IActionResult Download()
    {
        var values = new[] {
            new { Column1 = "MiniExcel", Column2 = 1 },
            new { Column1 = "Github", Column2 = 2}
        };
        var stream = new MemoryStream();
        stream.SaveAs(values);
        return File(stream,
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            "demo.xlsx");
    }
}
```

### Excel Type Auto Check

Default system will auto check file path or stream is from xlsx or csv, but if you need to specify type, it can use excelType parameter.
```C#
stream.SaveAs(excelType:ExcelType.CSV);
//or
stream.SaveAs(excelType:ExcelType.XLSX);
//or
stream.Query(excelType:ExcelType.CSV);
//or
stream.Query(excelType:ExcelType.XLSX);
```



### Dynamic Query cast row to `IDictionary<string,object>` 

```C#
foreach(IDictionary<string,object> row = MiniExcel.Query(path))
{
    //..
}
```



### Limitations and caveats 

- Not support xls and encrypted file now

### Reference

- [ExcelDataReader](https://github.com/ExcelDataReader/ExcelDataReader)   
- [StackExchange/Dapper](https://github.com/StackExchange/Dapper)    

### Contributors :  

![](https://contrib.rocks/image?repo=shps951023/MiniExcel)