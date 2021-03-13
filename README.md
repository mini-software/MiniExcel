| [![NuGet](https://img.shields.io/nuget/v/MiniExcel.svg)](https://www.nuget.org/packages/MiniExcel) | ![](https://img.shields.io/nuget/dt/MiniExcel.svg) | 
| -------- | -------- | 

### Features
- Support IEnumerable Lazy & Real `Step By Step one row read` not until all rows data read in memory
e.g:  Comparison of MiniExcel Query and ExcelDataReader of Reading Large Xlsx File
![](https://user-images.githubusercontent.com/12729184/110884175-9f9ca480-831f-11eb-9795-cf0b9f386955.gif)
- Mini (DLL Size Only 20KB) and Easy to use.
- Support .NET Standard 2.0/.NET 4.6/.NET 5
- Without Any Third Party Library Dependencies
- Support Anonymous Types,Dapper Dynamic Query,List/Array/Set/Enumrable,DataTable,Dictionary

### Installation

You can install the package [from NuGet](https://www.nuget.org/packages/MiniExcel)

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


### TODO

Please Check [Issues · todo](https://github.com/shps951023/MiniExcel/labels/todo)

## Release Notes

Please Check [Release Notes](https://github.com/shps951023/MiniExcel/tree/master/docs)