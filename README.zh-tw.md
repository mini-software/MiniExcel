[![NuGet](https://img.shields.io/nuget/v/MiniExcel.svg)](https://www.nuget.org/packages/MiniExcel)  [![](https://img.shields.io/nuget/dt/MiniExcel.svg)](https://www.nuget.org/packages/MiniExcel)  [![Build status](https://ci.appveyor.com/api/projects/status/b2vustrwsuqx45f4/branch/master?svg=true)](https://ci.appveyor.com/project/shps951023/miniexcel/branch/master) [![.NET Framework](https://img.shields.io/badge/.NET%20Framework-%3E%3D%204.6.1-red.svg)](#)  [![.NET Standard](https://img.shields.io/badge/.NET%20Standard-%3E%3D%202.0-red.svg)](#) [![.NET](https://img.shields.io/badge/.NET%20-%3E%3D%205.0-red.svg)](#) 
 
MiniExcel 簡單、高效避免OOM的.NET處理Excel工具。

---

- [English](README.md)
- [繁體中文](README.zh-tw.md)

---


目前主流框架大多需要將資料全載入到記憶體方便操作，但會導致記憶體消耗問題，MiniExcel嘗試以 stream 角度重寫底層算法邏輯，能讓原本1000多MB占用降低到幾MB，避免記憶體不夠情況。

### 特點
- 低記憶體耗用，避免OOM(out of memoery)
- 支持`即時`操作一行一行資料
![miniexcel_lazy_load](https://user-images.githubusercontent.com/12729184/111034290-e5588a80-844f-11eb-8c84-6fdb6fb8f403.gif)
- 兼具搭配 LINQ 延遲查詢特性，能辦到低消耗、快速分頁等複雜查詢
圖片:與主流框架對比的消耗、效率差  
![queryfirst](https://user-images.githubusercontent.com/12729184/111072392-6037a900-8515-11eb-9693-5ce2dad1e460.gif)
- 輕量，不依賴任何套件，DLL小於100KB

### 安裝

請查看 [from NuGet](https://www.nuget.org/packages/MiniExcel)

### 更新日誌

請查看 [Release Notes](https://github.com/shps951023/MiniExcel/tree/master/docs)

### TODO

請查看 [Project · todo](https://github.com/shps951023/MiniExcel/projects/1?fullscreen=true)

### Query 查詢 Excel 返回`強型別` IEnumerable 資料 [[Try it]](https://dotnetfiddle.net/w5WD1J)

推薦使用 Stream.Query 效率會相對較好。

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


### Query 查詢 Excel 返回`Dynamic` IEnumerable 資料 [[Try it]](https://dotnetfiddle.net/w5WD1J)

* Key 系統預設為 `A,B,C,D...Z`

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

### 查詢資料以第一行數據當Key [[Try it]](https://dotnetfiddle.net/w5WD1J)

note : 同名以右邊數據為準

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

### Query 查詢支援延遲加載(Deferred Execution)，能配合LINQ First/Take/Skip辦到低消耗、高效率複雜查詢

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


### 建立 Excel 檔案 [[Try it]](https://dotnetfiddle.net/w5WD1J)

1. 必須是 non-abstract 類別有公開建構式
2. MiniExcel SaveAs 支援 `IEnumerable參數``延遲查詢`，除非必要請不要使用 ToList 等方法讀取全部資料到記憶體，請看圖片了解差異

圖片 : 是否呼叫 ToList 的記憶體差別
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

| Column1 | Column2 | 
| -------- | -------- | 
| MiniExcel     | 1     |  
| Github     | 2     | 

### SaveAs 支援 Stream [[Try it]](https://dotnetfiddle.net/JOen0e)

```C#
using (var stream = File.Create(path))
{
    stream.SaveAs(values);
}
```


### 例子 : SQLite & Dapper 讀取大數據新增到資料庫

note : 請不要呼叫 call ToList/ToArray 等方法，這會將所有資料讀到記憶體內

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


### ASP.NET Core 3.1 or MVC 5 下載 Excel Xlsx API Demo

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

### Excel 類別自動判斷

MiniExcel 預設會根據擴展名或是 Stream 類別判斷是 xlsx 還是 csv，但會有失準時候，請自行指定。
```C#
stream.SaveAs(excelType:ExcelType.CSV);
//or
stream.SaveAs(excelType:ExcelType.XLSX);
//or
stream.Query(excelType:ExcelType.CSV);
//or
stream.Query(excelType:ExcelType.XLSX);
```


### 不足之處
- 目前不支援 xls 或是加密檔案。

### 參考
- 讀取邏輯　:  [ExcelDataReader](https://github.com/ExcelDataReader/ExcelDataReader)   
- API 設計方式 :　[StackExchange/Dapper](https://github.com/StackExchange/Dapper)    
