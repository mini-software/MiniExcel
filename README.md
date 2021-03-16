[![NuGet](https://img.shields.io/nuget/v/MiniExcel.svg)](https://www.nuget.org/packages/MiniExcel)  [![](https://img.shields.io/nuget/dt/MiniExcel.svg)](https://www.nuget.org/packages/MiniExcel)  [![Build status](https://ci.appveyor.com/api/projects/status/b2vustrwsuqx45f4/branch/master?svg=true)](https://ci.appveyor.com/project/shps951023/miniexcel/branch/master)
### Features
- `Avoid large file OOM(out of memoery)` by IEnumerable Lazy `Step By Step getting one row cells` not until all rows read in memory  
e.g:  Comparison of MiniExcel Query and ExcelDataReader/EPPlus/ClosedXml of reading large Xlsx File
![miniexcel_lazy_load](https://user-images.githubusercontent.com/12729184/111034290-e5588a80-844f-11eb-8c84-6fdb6fb8f403.gif)
- Support .NET Standard 2.0/.NET 4.6/.NET 5
- Mini without any third party library dependencies
- Support dynamic/type mapping query and create by AnonymousType/DapperRows/List/Array/Set/Enumrable/DataTable/Dictionary
- [Dapper](https://github.com/StackExchange/Dapper) query style 

### Installation

You can install the package [from NuGet](https://www.nuget.org/packages/MiniExcel)

### Release Notes

Please Check [Release Notes](https://github.com/shps951023/MiniExcel/tree/master/docs)

### Execute a query and map the results to a strongly typed IEnumerable

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

using (var stream = File.OpenRead(path))
    var rows = stream.Query<UserAccount>();
```

![image](https://user-images.githubusercontent.com/12729184/111107423-c8c46b80-8591-11eb-982f-c97a2dafb379.png)


### Execute a query and map it to a list of dynamic objects without using head

| MiniExcel     | 1     | 
| -------- | -------- | 
| Github     | 2     | 

```C#
using (var stream = File.OpenRead(path))
{
    var rows = stream.Query().ToList();
                
    Assert.Equal("MiniExcel", rows[0].A);
    Assert.Equal(1, rows[0].B);
    Assert.Equal("Github", rows[1].A);
    Assert.Equal(2, rows[1].B);
}
```

### Execute a query with first header row

| Column1 | Column2 | 
| -------- | -------- | 
| MiniExcel     | 1     |  
| Github     | 2     | 


```C#
using (var stream = File.OpenRead(path))
{
    var rows = stream.Query(useHeaderRow:true).ToList();

    Assert.Equal("MiniExcel", rows[0].Column1);
    Assert.Equal(1, rows[0].Column2);
    Assert.Equal("Github", rows[1].Column1);
    Assert.Equal(2, rows[1].Column2);
}
```

### Query Mapping Type

### Create Excel Xlsx file by ICollection Anonymous Type/Datatable
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

Create File Result : 

| Column1 | Column2 | 
| -------- | -------- | 
| MiniExcel     | 1     |  
| Github     | 2     | 

### SaveAs Stream

```C#
using (var stream = new FileStream(path, FileMode.CreateNew))
{
    stream.SaveAs(values);
}
```

### Query First

```C#
using (var stream = File.OpenRead(path))
    Assert.Equal("HelloWorld", stream.QueryFirst().A);
```

performance:  MiniExcel/ExcelDataReader/ClosedXML/EPPlus  
![queryfirst](https://user-images.githubusercontent.com/12729184/111072392-6037a900-8515-11eb-9693-5ce2dad1e460.gif)


### SQLite & Dapper `Large Size File` SQL Insert Avoid OOM (out of memory) 

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

### TODO

Please Check [Project · todo](https://github.com/shps951023/MiniExcel/projects/1?fullscreen=true)

### Limitations and caveats 

- Custom datetime/timespan format can't mapping to DateTime/TimeSpan type
- Same column name use last right one 
- Must be a non-abstract type with a public parameterless constructor 

### Reference 

- Query logic learn from [ExcelDataReader](https://github.com/ExcelDataReader/ExcelDataReader)
- Query Style learn from [StackExchange/Dapper](https://github.com/StackExchange/Dapper)