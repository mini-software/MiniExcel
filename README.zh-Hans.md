[![NuGet](https://img.shields.io/nuget/v/MiniExcel.svg)](https://www.nuget.org/packages/MiniExcel)  [![](https://img.shields.io/nuget/dt/MiniExcel.svg)](https://www.nuget.org/packages/MiniExcel)  [![Build status](https://ci.appveyor.com/api/projects/status/b2vustrwsuqx45f4/branch/master?svg=true)](https://ci.appveyor.com/project/shps951023/miniexcel/branch/master) [![.NET Framework](https://img.shields.io/badge/.NET%20Framework-%3E%3D%204.6.1-red.svg)](#)  [![.NET Standard](https://img.shields.io/badge/.NET%20Standard-%3E%3D%202.0-red.svg)](#) [![.NET](https://img.shields.io/badge/.NET%20-%3E%3D%205.0-red.svg)](#) 

---

[English](https://github.com/shps951023/MiniExcel) / [繁體中文](https://github.com/shps951023/MiniExcel/blob/master/README.zh-tw.md) / [简体中文](https://github.com/shps951023/MiniExcel/blob/master/README.zh-Hans.md)

---

### 简介

MiniExcel 简单、高效避免OOM的.NET处理Excel工具。


目前主流框架大多需要将资料全载入到内存方便操作，但这会导致内存消耗问题，MiniExcel 尝试以 Stream 角度写底层算法逻辑，能让原本1000多MB占用降低到几MB，避免内存不够情况。

![image](https://user-images.githubusercontent.com/12729184/113120478-33d59980-9244-11eb-8675-a49651c8af67.png)

### 特点

- 低内存耗用，避免OOM(out of memoery)、频繁 Full GC 情况
- 支持`即时`操作每行数据
  ![miniexcel_lazy_load](https://user-images.githubusercontent.com/12729184/111034290-e5588a80-844f-11eb-8c84-6fdb6fb8f403.gif)
- 兼具搭配 LINQ 延迟查询特性，能办到低消耗、快速分页等复杂查询  
  图片:与主流框架对比的消耗、效率差  
  ![queryfirst](https://user-images.githubusercontent.com/12729184/111072392-6037a900-8515-11eb-9693-5ce2dad1e460.gif)
- 轻量，不依赖任何套件，DLL小于100KB
- 简便操作的 Dapper API 风格

### 安装

请查看 [from NuGet](https://www.nuget.org/packages/MiniExcel)

### 更新日志

请查看 [Release Notes](https://github.com/shps951023/MiniExcel/tree/master/docs)

### Discussions / TODO 

请查看 [Discussions](https://github.com/shps951023/MiniExcel/discussions) / [TODO](https://github.com/shps951023/MiniExcel/projects/1?fullscreen=true)

### 性能测试

以 [**Test1,000,000x10.xlsx**](https://github.com/shps951023/MiniExcel/blob/master/samples/xlsx/Test1%2C000%2C000x10/Test1%2C000%2C000x10.xlsx) 做基准与主流框架做性能测试，总共 1千万笔 "HelloWorld"，文件大小 23 MB   

Benchmarks  逻辑可以在 [MiniExcel.Benchmarks](https://github.com/shps951023/MiniExcel/tree/master/benchmarks/MiniExcel.Benchmarks) 查看或是提交 PR，运行指令

```
dotnet run -p .\benchmarks\MiniExcel.Benchmarks\ -c Release -f netcoreapp3.1 -- -f * --join
```

最后一次运行结果 :  

```
BenchmarkDotNet=v0.12.1, OS=Windows 10.0.19042
Intel Core i7-7700 CPU 3.60GHz (Kaby Lake), 1 CPU, 8 logical and 4 physical cores
  [Host]     : .NET Framework 4.8 (4.8.4341.0), X64 RyuJIT
  Job-ZYYABG : .NET Framework 4.8 (4.8.4341.0), X64 RyuJIT
IterationCount=3  LaunchCount=3  WarmupCount=3  
```

| Method                       | 最大内存耗用 |         平均时间 |        Gen 0 |       Gen 1 |      Gen 2 |
| ---------------------------- | -----------: | ---------------: | -----------: | ----------: | ---------: |
| 'MiniExcel QueryFirst'       |     0.109 MB |         726.4 us |            - |           - |          - |
| 'ExcelDataReader QueryFirst' |     15.24 MB |  10,664,238.2 us |  566000.0000 |   1000.0000 |          - |
| 'MiniExcel Query'            |      17.3 MB |  14,179,334.8 us |  367000.0000 |  96000.0000 |  7000.0000 |
| 'ExcelDataReader Query'      |      17.3 MB |  22,565,088.7 us | 1210000.0000 |   2000.0000 |          - |
| 'Epplus QueryFirst'          |     1,452 MB |  18,198,015.4 us |  535000.0000 | 132000.0000 |  9000.0000 |
| 'Epplus Query'               |     1,451 MB |  23,647,471.1 us | 1451000.0000 | 133000.0000 |  9000.0000 |
| 'OpenXmlSDK Query'           |     1,412 MB |  52,003,270.1 us |  978000.0000 | 353000.0000 | 11000.0000 |
| 'OpenXmlSDK QueryFirst'      |     1,413 MB |  52,348,659.1 us |  978000.0000 | 353000.0000 | 11000.0000 |
| 'ClosedXml QueryFirst'       |     2,158 MB |  66,188,979.6 us | 2156000.0000 | 575000.0000 |  9000.0000 |
| 'ClosedXml Query'            |     2,184 MB | 191,434,126.6 us | 2165000.0000 | 577000.0000 | 10000.0000 |


| Method                   | 最大内存耗用 |         平均时间 |        Gen 0 |        Gen 1 |      Gen 2 |
| ------------------------ | -----------: | ---------------: | -----------: | -----------: | ---------: |
| 'MiniExcel Create Xlsx'  |        15 MB |  11,531,819.8 us | 1020000.0000 |            - |          - |
| 'Epplus Create Xlsx'     |     1,204 MB |  22,509,717.7 us | 1370000.0000 |   60000.0000 | 30000.0000 |
| 'OpenXmlSdk Create Xlsx' |     2,621 MB |  42,473,998.9 us | 1370000.0000 |  460000.0000 | 50000.0000 |
| 'ClosedXml Create Xlsx'  |     7,141 MB | 140,939,928.6 us | 5520000.0000 | 1500000.0000 | 80000.0000 |

### Query 查询 Excel 返回`强型别` IEnumerable 数据 [[Try it]](https://dotnetfiddle.net/w5WD1J)

推荐使用 Stream.Query 效率会相对较好。

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


### Query 查询 Excel 返回`Dynamic` IEnumerable 数据 [[Try it]](https://dotnetfiddle.net/w5WD1J)

* Key 系统预设为 `A,B,C,D...Z`

| MiniExcel | 1    |
| --------- | ---- |
| Github    | 2    |

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

### 查询数据以第一行数据当Key [[Try it]](https://dotnetfiddle.net/w5WD1J)

note : 同名以右边数据为准   

Input Excel :    

| Column1   | Column2 |
| --------- | ------- |
| MiniExcel | 1       |
| Github    | 2       |


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

### Query 查询支援延迟加载(Deferred Execution)，能配合LINQ First/Take/Skip办到低消耗、高效率复杂查询

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

### 建立 Excel 文件 [[Try it]](https://dotnetfiddle.net/w5WD1J)

1. 必须是 non-abstract 类别有 public parameterless constructor    
2. MiniExcel SaveAs 支援 `IEnumerable参数``延迟查询`，除非必要请不要使用 ToList 等方法读取全部数据到内存   

图片 : 是否呼叫 ToList 的内存差别  
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

output : 

| Column1   | Column2 |
| --------- | ------- |
| MiniExcel | 1       |
| Github    | 2       |

### SaveAs 支援 Stream [[Try it]](https://dotnetfiddle.net/JOen0e)

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





### 例子 : SQLite & Dapper 读取大数据新增到数据库

note : 请不要呼叫 call ToList/ToArray 等方法，这会将所有数据读到内存内

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

效能:
![image](https://user-images.githubusercontent.com/12729184/111072579-2dda7b80-8516-11eb-9843-c01a1edc88ec.png)


### 例子 : ASP.NET Core 3.1 or MVC 5 下载 Excel Xlsx API Demo

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

### Excel 类别自动判断

MiniExcel 预设会根据扩展名或是 Stream 类别判断是 xlsx 还是 csv，但会有失准时候，请自行指定。

```C#
stream.SaveAs(excelType:ExcelType.CSV);
//or
stream.SaveAs(excelType:ExcelType.XLSX);
//or
stream.Query(excelType:ExcelType.CSV);
//or
stream.Query(excelType:ExcelType.XLSX);
```



### Dynamic Query 转换 `IDictionary<string,object>` 数据

```C#
foreach(IDictionary<string,object> row = MiniExcel.Query(path))
{
    //..
}
```



### 局限与警告

- 目前不支援 xls (97-2003) 或是加密文件。
- 不支援样式、字体、宽度等`修改`，因为 MiniExcel 概念是只专注于值数据，借此降低内存消耗跟提升效率。

### 参考

- 读取逻辑 :  [ExcelDataReader](https://github.com/ExcelDataReader/ExcelDataReader)   
- API 设计方式 :　[StackExchange/Dapper](https://github.com/StackExchange/Dapper)    

### Contributors :  

![](https://contrib.rocks/image?repo=shps951023/MiniExcel)