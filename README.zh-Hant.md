<div align="center">
<p><a href="https://www.nuget.org/packages/MiniExcel"><img src="https://img.shields.io/nuget/v/MiniExcel.svg" alt="NuGet"></a>  <a href="https://www.nuget.org/packages/MiniExcel"><img src="https://img.shields.io/nuget/dt/MiniExcel.svg" alt=""></a>
<a href="https://ci.appveyor.com/project/shps951023/miniexcel/branch/master"><img src="https://ci.appveyor.com/api/projects/status/b2vustrwsuqx45f4/branch/master?svg=true" alt="Build status"></a>
<a href="https://gitee.com/dotnetchina/MiniExcel"><img src="https://gitee.com/dotnetchina/MiniExcel/badge/star.svg" alt="star"></a> <a href="https://github.com/shps951023/MiniExcel" rel="nofollow"><img src="https://img.shields.io/github/stars/shps951023/MiniExcel?logo=github" alt="GitHub stars"></a>
<a href="https://www.nuget.org/packages/MiniExcel"><img src="https://img.shields.io/badge/.NET-%3E%3D%204.5-red.svg" alt="version"></a>
</p>
</div>

---

<div align="center">
<p><strong><a href="README.md">English</a> | <a href="README.zh-CN.md">简体中文</a> | <a href="README.zh-Hant.md">繁體中文</a></strong></p>
</div>

---

[<img align="right" src="https://github.com/dotnet-foundation/swag/blob/main/logo/dotnetfoundation_v4.png?raw=true" width="100" />](https://www.dotnetfoundation.org/)

<div align="center">
<p>該專案是 <a href="https://www.dotnetfoundation.org/">.NET 基金會</a> 的一部分，並根據其<a href="https://www.dotnetfoundation.org/code-of-conduct">行為準則</a> 運作。</p>
</div>

---

<div align="center">
<p> 您的 <a href="https://github.com/shps951023/miniexcel">Star</a> 和 <a href="https://miniexcel.github.io">贊助</a> 能幫助 MiniExcel 成長 </p>
</div>

---

### [🎥影片教學](https://www.udemy.com/course/miniexcel-tutorial/?referralCode=732E11323F1E8064F96C)

---

### 簡介

MiniExcel 簡單、高效避免OOM的.NET處理Excel查、寫、填充工具。

目前主流框架大多需要將資料全載入到記憶體方便操作，但這會導致記憶體消耗問題，MiniExcel 嘗試以 Stream 角度寫底層算法邏輯，能讓原本1000多MB占用降低到幾MB，避免記憶體不夠情況。

![image](https://user-images.githubusercontent.com/12729184/113084691-1804d000-9211-11eb-9b08-cbb89d9ecdc2.png)

### 特點
- 低記憶體耗用，避免OOM、頻繁 Full GC 情況
- 支持`即時`操作每行資料
- 兼具搭配 LINQ 延遲查詢特性，能辦到低消耗、快速分頁等複雜查詢
- 輕量，不需要安裝 Microsoft Office、COM+，DLL小於150KB
- 簡便操作的 API 風格



### 快速開始

- [導入、讀取 Excel](#getstart1)
- [導出 、創建 Excel](#getstart2)
- [模板填充、創建 Excel](#getstart3)
- [Excel Column Name/Index/Ignore Attribute](#getstart4)
- [例子](#getstart5)


### 安裝

請查看 [NuGet](https://www.nuget.org/packages/MiniExcel)

### 更新日誌

請查看 [Release Notes](docs)

### TODO

請查看 [TODO](https://github.com/shps951023/MiniExcel/projects/1?fullscreen=true)

### 性能比較、測試

Benchmarks  邏輯可以在 [MiniExcel.Benchmarks](benchmarks/MiniExcel.Benchmarks/Program.cs) 查看或是提交 PR，運行指令

```bash
dotnet run -p .\benchmarks\MiniExcel.Benchmarks\ -c Release -f netcoreapp3.1 -- -f * --join
```

最後一次運行規格、結果 :

```bash
BenchmarkDotNet=v0.12.1, OS=Windows 10.0.19042
Intel Core i7-7700 CPU 3.60GHz (Kaby Lake), 1 CPU, 8 logical and 4 physical cores
  [Host]     : .NET Framework 4.8 (4.8.4341.0), X64 RyuJIT
  Job-ZYYABG : .NET Framework 4.8 (4.8.4341.0), X64 RyuJIT
IterationCount=3  LaunchCount=3  WarmupCount=3
```

Benchmark History :  [Link](https://github.com/shps951023/MiniExcel/issues/276)



#### 導入、查詢 Excel 比較

邏輯 : 以 [**Test1,000,000x10.xlsx**](benchmarks/MiniExcel.Benchmarks/Test1%2C000%2C000x10.xlsx) 做基準與主流框架做性能測試，總共 1,000,000 行 * 10 列筆 "HelloWorld"，文件大小 23 MB


| Library      | Method                       | 最大記憶體耗用 |         平均時間 |
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

#### 導出、創建 Excel 比較

邏輯 : 創建1千萬筆 "HelloWorld"

| Library            | Method                   | 最大記憶體耗用 |         平均時間 |
| ------------------------ | -------------: | ---------------: | -----------: |
| MiniExcel | 'MiniExcel Create Xlsx'  |          15 MB | 11.53181 sec |
| Epplus | 'Epplus Create Xlsx'     |       1,204 MB | 22.50971 sec |
| OpenXmlSdk | 'OpenXmlSdk Create Xlsx' |       2,621 MB | 42.47399 sec |
| ClosedXml | 'ClosedXml Create Xlsx'  |       7,141 MB | 140.93992 sec |





### 讀/導入 Excel <a name="getstart1"></a>

- 支持任何 stream 类型 : FileStream,MemoryStream



#### 1. Query 查詢 Excel 返回`強型別` IEnumerable 資料 [[Try it]](https://dotnetfiddle.net/w5WD1J)

```csharp
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


#### 2. Query 查詢 Excel 返回`Dynamic` IEnumerable 資料 [[Try it]](https://dotnetfiddle.net/w5WD1J)

* Key 系統預設為 `A,B,C,D...Z`

| MiniExcel     | 1     |
| -------- | -------- |
| Github     | 2     |

```csharp

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

#### 3. 查詢資料以第一行數據當Key [[Try it]](https://dotnetfiddle.net/w5WD1J)

注意 : 同名以右邊數據為準

Input Excel :

| Column1 | Column2 |
| -------- | -------- |
| MiniExcel     | 1     |
| Github     | 2     |


```csharp

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

#### 4. Query 查詢支援延遲加載(Deferred Execution)，能配合LINQ First/Take/Skip辦到低消耗、高效率複雜查詢

舉例 : 查詢第一筆資料

```csharp
var row = MiniExcel.Query(path).First();
Assert.Equal("HelloWorld", row.A);

// or

using (var stream = File.OpenRead(path))
{
    var row = stream.Query().First();
    Assert.Equal("HelloWorld", row.A);
}
```

與其他框架效率比較 :

![queryfirst](https://user-images.githubusercontent.com/12729184/111072392-6037a900-8515-11eb-9693-5ce2dad1e460.gif)

#### 5. 查詢指定 Sheet 名稱

```csharp
MiniExcel.Query(path, sheetName: "SheetName");
//or
stream.Query(sheetName: "SheetName");
```

#### 6. 查詢所有 Sheet 名稱跟資料

```csharp
var sheetNames = MiniExcel.GetSheetNames(path);
foreach (var sheetName in sheetNames)
{
    var rows = MiniExcel.Query(path, sheetName: sheetName);
}
```

#### 7. 查詢所有欄(列)

```csharp
var columns = MiniExcel.GetColumns(path); // e.g result : ["A","B"...]

var cnt = columns.Count;  // get column count
```

#### 8. Dynamic Query 轉成 `IDictionary<string,object>` 資料

```csharp
foreach(IDictionary<string,object> row in MiniExcel.Query(path))
{
    //..
}

// or
var rows = MiniExcel.Query(path).Cast<IDictionary<string,object>>();
```



#### 9. Query 讀 Excel 返回 DataTable

提醒 : 不建議使用，因為DataTable會將數據`全載入內存`，失去MiniExcel低記憶體消耗功能。

```C#
var table = MiniExcel.QueryAsDataTable(path, useHeaderRow: true);
```

![image](https://user-images.githubusercontent.com/12729184/116673475-07917200-a9d6-11eb-947e-a6f68cce58df.png)

#### 10. 指定單元格開始讀取資料

```csharp
MiniExcel.Query(path,useHeaderRow:true,startCell:"B3")
```

![image](https://user-images.githubusercontent.com/12729184/117260316-8593c400-ae81-11eb-9877-c087b7ac2b01.png)

#### 11. 合併的單元格填充

注意 : 效率相對於`沒有使用合併填充`來說差
底層原因 : OpenXml 標准將 mergeCells 放在文件最下方，導致需要遍歷兩次 sheetxml

```csharp
    var config = new OpenXmlConfiguration()
    {
        FsillMergedCells = true
    };
    var rows = MiniExcel.Query(path, configuration: config);
```

![image](https://user-images.githubusercontent.com/12729184/117973630-3527d500-b35f-11eb-95c3-bde255f8114e.png)

支持不固定長寬多行列填充

![image](https://user-images.githubusercontent.com/12729184/117973820-6d2f1800-b35f-11eb-88d8-555063938108.png)


#### 12. 讀取大文件硬碟緩存 (Disk-Base Cache - SharedString)

概念 : MiniExcel 當判斷文件 SharedString 大小超過 5MB，預設會使用本地緩存，如 [10x100000.xlsx](https://github.com/MiniExcel/MiniExcel/files/8403819/NotDuplicateSharedStrings_10x100000.xlsx)(一百萬筆數據)，讀取不開啟本地緩存需要最高記憶體使用約195MB，開啟後降為65MB。但要特別注意，此優化是以`時間換取記憶體減少`，所以讀取效率會變慢，此例子讀取時間從 7.4 秒提高到 27.2 秒，假如不需要能用以下代碼關閉硬碟緩存

```csharp
var config = new OpenXmlConfiguration { EnableSharedStringCache = false };
MiniExcel.Query(path,configuration: config)
```

也能使用 SharedStringCacheSize 調整 sharedString 文件大小超過指定大小才做硬碟緩存
```csharp
var config = new OpenXmlConfiguration { SharedStringCacheSize=500*1024*1024 };
MiniExcel.Query(path, configuration: config);
```


![image](https://user-images.githubusercontent.com/12729184/161411851-1c3f72a7-33b3-4944-84dc-ffc1d16747dd.png)

![image](https://user-images.githubusercontent.com/12729184/161411825-17f53ec7-bef4-4b16-b234-e24799ea41b0.png)



### 寫/導出 Excel  <a name="getstart2"></a>

1. 必須是非abstract 類別有公開無參數構造函數
2. MiniExcel SaveAs 支援 `IEnumerable參數延遲查詢`，除非必要請不要使用 ToList 等方法讀取全部資料到記憶體

圖片 : 是否呼叫 ToList 的記憶體差別

#### ![image](https://user-images.githubusercontent.com/12729184/112587389-752b0b00-8e38-11eb-8a52-cfb76c57e5eb.png)



#### 1. 支持集合<匿名類別>或是<強型別> [[Try it]](https://dotnetfiddle.net/w5WD1J)

```csharp
var path = Path.Combine(Path.GetTempPath(), $"{Guid.NewGuid()}.xlsx");
MiniExcel.SaveAs(path, new[] {
    new { Column1 = "MiniExcel", Column2 = 1 },
    new { Column1 = "Github", Column2 = 2}
});
```



#### 2. `IEnumerable<IDictionary<string, object>>`

```csharp
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



#### 3.  IDataReader

- 推薦使用，可以避免載入全部數據到記憶體

```csharp
MiniExcel.SaveAs(path, reader);
```

![image](https://user-images.githubusercontent.com/12729184/121275378-149a5e80-c8bc-11eb-85fe-5453552134f0.png)

推薦 DataReader 多表格導出方式(建議使用 Dapper ExecuteReader )

```csharp
using (var cnn = Connection)
{
    cnn.Open();
    var sheets = new Dictionary<string,object>();
    sheets.Add("sheet1", cnn.ExecuteReader("select 1 id"));
    sheets.Add("sheet2", cnn.ExecuteReader("select 2 id"));
    MiniExcel.SaveAs("Demo.xlsx", sheets);
}
```



####  4. Datatable

- `不推薦使用`，會將數據全載入記憶體
- 優先使用 Caption 當欄位名稱

```csharp
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

####  5. Dapper Query

感謝 @shaofing #552 更正，低內存請使用 `CommandDefinition + CommandFlags.NoCache`，如下

```csharp
using (var connection = GetConnection(connectionString))
{
    var rows = connection.Query(
        new CommandDefinition(
            @"select 'MiniExcel' as Column1,1 as Column2 union all select 'Github',2"
            , flags: CommandFlags.NoCache)
        );
    MiniExcel.SaveAs(path, rows);
}
```
上面的方法已知的問題：不能使用異步QueryAsync的方法，會報連接已經關閉的異常

以下寫法會將數據全載入內存

```csharp
using (var connection = GetConnection(connectionString))
{
    var rows = connection.Query(@"select 'MiniExcel' as Column1,1 as Column2 union all select 'Github',2");
    MiniExcel.SaveAs(path, rows);
}
```



#### 6. SaveAs 支持 Stream，生成文件不落地 [[Try it]](https://dotnetfiddle.net/JOen0e)

```csharp
using (var stream = new MemoryStream()) //支持 FileStream,MemoryStream..等
{
    stream.SaveAs(values);
}
```

像是 API 導出 Excel

```csharp
public IActionResult DownloadExcel()
{
    var values = new[] {
        new { Column1 = "MiniExcel", Column2 = 1 },
        new { Column1 = "Github", Column2 = 2}
    };

    var memoryStream = new MemoryStream();
    memoryStream.SaveAs(values);
    memoryStream.Seek(0, SeekOrigin.Begin);
    return new FileStreamResult(memoryStream, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    {
        FileDownloadName = "demo.xlsx"
    };
}
```



#### 7. 創建多個工作表(Sheet)

```csharp
// 1. Dictionary<string,object>
var users = new[] { new { Name = "Jack", Age = 25 }, new { Name = "Mike", Age = 44 } };
var department = new[] { new { ID = "01", Name = "HR" }, new { ID = "02", Name = "IT" } };
var sheets = new Dictionary<string, object>
{
    ["users"] = users,
    ["department"] = department
};
MiniExcel.SaveAs(path, sheets);

// 2. DataSet
var sheets = new DataSet();
sheets.Add(UsersDataTable);
sheets.Add(DepartmentDataTable);
//..
MiniExcel.SaveAs(path, sheets);
```

![image](https://user-images.githubusercontent.com/12729184/118130875-6e7c4580-b430-11eb-9b82-22f02716bd63.png)

#### 8. 表格樣式選擇

預設樣式

![image](https://user-images.githubusercontent.com/12729184/138234373-cfa97109-b71f-4711-b7f5-0eaaa4a0a3a6.png)

不需要樣式

```csharp
var config = new OpenXmlConfiguration()
{
    TableStyles = TableStyles.None
};
MiniExcel.SaveAs(path, value,configuration:config);
```

![image](https://user-images.githubusercontent.com/12729184/118784917-f3e57700-b8c2-11eb-8718-8d955b1bc197.png)





#### 9. AutoFilter 篩選

從 0.19.0 支持，可藉由 OpenXmlConfiguration.AutoFilter 設定，預設為True。關閉 AutoFilter 方式 :

```csharp
MiniExcel.SaveAs(path, value, configuration: new OpenXmlConfiguration() { AutoFilter = false });
```



#### 10. 圖片生成

```csharp
var value = new[] {
    new { Name="github",Image=File.ReadAllBytes(PathHelper.GetFile("images/github_logo.png"))},
    new { Name="google",Image=File.ReadAllBytes(PathHelper.GetFile("images/google_logo.png"))},
    new { Name="microsoft",Image=File.ReadAllBytes(PathHelper.GetFile("images/microsoft_logo.png"))},
    new { Name="reddit",Image=File.ReadAllBytes(PathHelper.GetFile("images/reddit_logo.png"))},
    new { Name="statck_overflow",Image=File.ReadAllBytes(PathHelper.GetFile("images/statck_overflow_logo.png"))},
};
MiniExcel.SaveAs(path, value);
```

![image](https://user-images.githubusercontent.com/12729184/150462383-ad9931b3-ed8d-4221-a1d6-66f799743433.png)



#### 11. Byte Array 文件導出

從 1.22.0 開始，當值類型為 `byte[]` 系統預設會轉成保存文件路徑以便導入時轉回 `byte[]`，如不想轉換可以將 `OpenXmlConfiguration.EnableConvertByteArray` 改為 `false`，能提升系統效率。

![image](https://user-images.githubusercontent.com/12729184/153702334-c3b834f4-6ae4-4ddf-bd4e-e5005d5d8c6a.png)

#### 12. 垂直合併相同的單元格

只支持 `xlsx` 格式合併單元格

```csharp
var mergedFilePath = Path.Combine(Path.GetTempPath(), $"{Guid.NewGuid().ToString()}.xlsx");

var path = @"../../../../../samples/xlsx/TestMergeSameCells.xlsx";

MiniExcel.MergeSameCells(mergedFilePath, path);
```

```csharp
var memoryStream = new MemoryStream();

var path = @"../../../../../samples/xlsx/TestMergeSameCells.xlsx";

memoryStream.MergeSameCells(path);
```

合併前後對比

![before_merge_cells](https://user-images.githubusercontent.com/38832863/219970175-913b3d04-d714-4279-a7a4-6cefb7aa6ce8.PNG)
![after_merge_cells](https://user-images.githubusercontent.com/38832863/219970176-e78c491a-2f90-45a7-a4a2-425c5708d38c.PNG)



#### 13. 是否寫入 null values cell

預設:

```csharp
DataTable dt = new DataTable();

/* ... */

DataRow dr = dt.NewRow();

dr["Name1"] = "Somebody once";
dr["Name2"] = null;
dr["Name3"] = "told me.";

dt.Rows.Add(dr);

MiniExcel.SaveAs(@"C:\temp\Book1.xlsx", dt);
```

![image](https://user-images.githubusercontent.com/31481586/241419441-c4f27e8f-3f87-46db-a10f-08665864c874.png)

```xml
<x:row r="2">
    <x:c r="A2" t ="str" s="2">
        <x:v>Somebody once</x:v>
    </x:c>
    <x:c r="B2" t ="str" s="2">
        <x:v></x:v>
    </x:c>
    <x:c r="C2" t ="str" s="2">
        <x:v>told me.</x:v>
    </x:c>
</x:row>
```

設定不寫入:

```csharp
OpenXmlConfiguration configuration = new OpenXmlConfiguration()
{
     EnableWriteNullValueCell = false // Default value is true.
};

MiniExcel.SaveAs(@"C:\temp\Book1.xlsx", dt, configuration: configuration);
```

![image](https://user-images.githubusercontent.com/31481586/241419455-3c0aec8a-4e5f-4d83-b7ec-6572124c165d.png)


```xml
<x:row r="2">
    <x:c r="A2" t ="str" s="2">
        <x:v>Somebody once</x:v>
    </x:c>
    <x:c r="B2" s="2"></x:c>
    <x:c r="C2" t ="str" s="2">
        <x:v>told me.</x:v>
    </x:c>
</x:row>
```



### 模板填充 Excel <a name="getstart3"></a>

- 宣告方式類似 Vue 模板 `{{變量名稱}}`, 或是集合渲染 `{{集合名稱.欄位名稱}}`
- 集合渲染支持 IEnumerable/DataTable/DapperRow



#### 1. 基本填充

模板:
![image](https://user-images.githubusercontent.com/12729184/114537556-ed8d2b00-9c84-11eb-8303-a69f62c41e5b.png)

最終效果:
![image](https://user-images.githubusercontent.com/12729184/114537490-d8180100-9c84-11eb-8c69-db58692f3a85.png)

代碼:
```csharp
// 1. By POCO
var value = new
{
    Name = "Jack",
    CreateDate = new DateTime(2021, 01, 01),
    VIP = true,
    Points = 123
};
MiniExcel.SaveAsByTemplate(path, templatePath, value);


// 2. By Dictionary
var value = new Dictionary<string, object>()
{
    ["Name"] = "Jack",
    ["CreateDate"] = new DateTime(2021, 01, 01),
    ["VIP"] = true,
    ["Points"] = 123
};
MiniExcel.SaveAsByTemplate(path, templatePath, value);
```



#### 2. IEnumerable/DataTable 數據填充

> Note1: 同行從左往右以第一個 IEnumerableUse 當列表來源 (不支持同列多集合)

模板:
![image](https://user-images.githubusercontent.com/12729184/114564652-14f2f080-9ca3-11eb-831f-09e3fedbc5fc.png)

最終效果:
![image](https://user-images.githubusercontent.com/12729184/114564204-b2015980-9ca2-11eb-900d-e21249f93f7c.png)

代碼:

```csharp
//1. By POCO
var value = new
{
    employees = new[] {
        new {name="Jack",department="HR"},
        new {name="Lisa",department="HR"},
        new {name="John",department="HR"},
        new {name="Mike",department="IT"},
        new {name="Neo",department="IT"},
        new {name="Loan",department="IT"}
    }
};
MiniExcel.SaveAsByTemplate(path, templatePath, value);

//2. By Dictionary
var value = new Dictionary<string, object>()
{
    ["employees"] = new[] {
        new {name="Jack",department="HR"},
        new {name="Lisa",department="HR"},
        new {name="John",department="HR"},
        new {name="Mike",department="IT"},
        new {name="Neo",department="IT"},
        new {name="Loan",department="IT"}
    }
};
MiniExcel.SaveAsByTemplate(path, templatePath, value);
```



#### 3. 複雜數據填充

> Note: 支持多 sheet 填充,並共用同一組參數

模板:

![image](https://user-images.githubusercontent.com/12729184/114565255-acf0da00-9ca3-11eb-8a7f-8131b2265ae8.png)

最終效果:

![image](https://user-images.githubusercontent.com/12729184/114565329-bf6b1380-9ca3-11eb-85e3-3969e8bf6378.png)

代碼:

```csharp
// 1. By POCO
var value = new
{
    title = "FooCompany",
    managers = new[] {
        new {name="Jack",department="HR"},
        new {name="Loan",department="IT"}
    },
    employees = new[] {
        new {name="Wade",department="HR"},
        new {name="Felix",department="HR"},
        new {name="Eric",department="IT"},
        new {name="Keaton",department="IT"}
    }
};
MiniExcel.SaveAsByTemplate(path, templatePath, value);

// 2. By Dictionary
var value = new Dictionary<string, object>()
{
    ["title"] = "FooCompany",
    ["managers"] = new[] {
        new {name="Jack",department="HR"},
        new {name="Loan",department="IT"}
    },
    ["employees"] = new[] {
        new {name="Wade",department="HR"},
        new {name="Felix",department="HR"},
        new {name="Eric",department="IT"},
        new {name="Keaton",department="IT"}
    }
};
MiniExcel.SaveAsByTemplate(path, templatePath, value);
```

#### 4. 大數據填充效率比較

> NOTE: 在 MiniExcel 使用 IEnumerable 延遲 ( 不ToList ) 可以節省記憶體使用

![image](https://user-images.githubusercontent.com/12729184/114577091-5046ec80-9cae-11eb-924b-087c7becf8da.png)



#### 5. Cell 值自動類別對應

模板

![image](https://user-images.githubusercontent.com/12729184/114802504-64830a80-9dd0-11eb-8d56-8e8c401b3ace.png)

最終效果

![image](https://user-images.githubusercontent.com/12729184/114802419-43221e80-9dd0-11eb-9ffe-a2ce34fe7076.png)

類別

```csharp
public class Poco
{
    public string @string { get; set; }
    public int? @int { get; set; }
    public decimal? @decimal { get; set; }
    public double? @double { get; set; }
    public DateTime? datetime { get; set; }
    public bool? @bool { get; set; }
    public Guid? Guid { get; set; }
}
```

代碼

```csharp
var poco = new TestIEnumerableTypePoco { @string = "string", @int = 123, @decimal = decimal.Parse("123.45"), @double = (double)123.33, @datetime = new DateTime(2021, 4, 1), @bool = true, @Guid = Guid.NewGuid() };
var value = new
{
    Ts = new[] {
        poco,
        new TestIEnumerableTypePoco{},
        null,
        poco
    }
};
MiniExcel.SaveAsByTemplate(path, templatePath, value);
```



#### 6. Example :  列出 Github 專案

模板

![image](https://user-images.githubusercontent.com/12729184/115068665-221f1200-9f25-11eb-9820-3d7d9638cb03.png)

最終效果

![image](https://user-images.githubusercontent.com/12729184/115068685-2814f300-9f25-11eb-96b5-0e7f21297f4d.png)

代碼

```csharp
var projects = new[]
{
    new {Name = "MiniExcel",Link="https://github.com/shps951023/MiniExcel",Star=146, CreateTime=new DateTime(2021,03,01)},
    new {Name = "HtmlTableHelper",Link="https://github.com/shps951023/HtmlTableHelper",Star=16, CreateTime=new DateTime(2020,02,01)},
    new {Name = "PocoClassGenerator",Link="https://github.com/shps951023/PocoClassGenerator",Star=16, CreateTime=new DateTime(2019,03,17)}
};
var value = new
{
    User = "ITWeiHan",
    Projects = projects,
    TotalStar = projects.Sum(s => s.Star)
};
MiniExcel.SaveAsByTemplate(path, templatePath, value);
```

#### 7. 分組數據填充

```csharp
var value = new Dictionary<string, object>()
{
    ["employees"] = new[] {
        new {name="Jack",department="HR"},
        new {name="Jack",department="HR"},
        new {name="John",department="HR"},
        new {name="John",department="IT"},
        new {name="Neo",department="IT"},
        new {name="Loan",department="IT"}
    }
};
MiniExcel.SaveAsByTemplate(path, templatePath, value);
```
##### 1. 使用`@group` tag 和 @header` tag

Before

![before_with_header](https://user-images.githubusercontent.com/38832863/218646717-21b9d57a-2be2-4e9a-801b-ae212231d2b4.PNG)

After

![after_with_header](https://user-images.githubusercontent.com/38832863/218646721-58a7a340-7004-4bc2-af24-cffcb2c20737.PNG)

##### 2. 使用 @group tag 沒有 @header tag

Before

![before_without_header](https://user-images.githubusercontent.com/38832863/218646873-b12417fa-801b-4890-8e96-669ed3b43902.PNG)

After

![after_without_header](https://user-images.githubusercontent.com/38832863/218646872-622461ba-342e-49ee-834f-b91ad9c2dac3.PNG)

##### 3. 沒有 @group tag

Before

![without_group](https://user-images.githubusercontent.com/38832863/218646975-f52a68eb-e031-43b5-abaa-03b67c052d1a.PNG)

After

![without_group_after](https://user-images.githubusercontent.com/38832863/218646974-4a3c0e07-7c66-4088-ad07-b4ad3695b7e1.PNG)

#### 8. DataTable 當參數

```csharp
var managers = new DataTable();
{
    managers.Columns.Add("name");
    managers.Columns.Add("department");
    managers.Rows.Add("Jack", "HR");
    managers.Rows.Add("Loan", "IT");
}
var value = new Dictionary<string, object>()
{
    ["title"] = "FooCompany",
    ["managers"] = managers,
};
MiniExcel.SaveAsByTemplate(path, templatePath, value);
```

#### 9. 其他

##### 1. 檢查模版參數

從 V1.24.0 版本開始，預設忽略模版不存在的參數Key，IgnoreTemplateParameterMissing 可以決定是否拋出錯誤

```csharp
var config = new OpenXmlConfiguration()
{
    IgnoreTemplateParameterMissing = false,
};
MiniExcel.SaveAsByTemplate(path, templatePath, value, config)
```

![image](https://user-images.githubusercontent.com/12729184/157464332-e316f829-54aa-4c84-a5aa-9aef337b668d.png)



### Excel 列屬性 (Excel Column Attribute) <a name="getstart4"></a>



#### 1. 指定列名稱、指定第幾列、是否忽略該列

Excel例子

![image](https://user-images.githubusercontent.com/12729184/114230869-3e163700-99ac-11eb-9a90-2039d4b4b313.png)


代碼

```csharp
public class ExcelAttributeDemo
{
    [ExcelColumnName("Column1")]
    public string Test1 { get; set; }
    [ExcelColumnName("Column2")]
    public string Test2 { get; set; }
    [ExcelIgnore]
    public string Test3 { get; set; }
    [ExcelColumnIndex("I")] // 系統會自動轉換"I"為第8列
    public string Test4 { get; set; }
    public string Test5 { get; } //系統會忽略此列
    public string Test6 { get; private set; } //set非公開,系統會忽略
    [ExcelColumnIndex(3)] // 從0開始索引
    public string Test7 { get; set; }
}

var rows = MiniExcel.Query<ExcelAttributeDemo>(path).ToList();
Assert.Equal("Column1", rows[0].Test1);
Assert.Equal("Column2", rows[0].Test2);
Assert.Null(rows[0].Test3);
Assert.Equal("Test7", rows[0].Test4);
Assert.Null(rows[0].Test5);
Assert.Null(rows[0].Test6);
Assert.Equal("Test4", rows[0].Test7);
```

#### 2. 自定義Format格式 (ExcelFormatAttribute)

從 V0.21.0 開始支持有 `ToString(string content)` 的類別 format

類別

```csharp
public class Dto
{
    public string Name { get; set; }

    [ExcelFormat("MMMM dd, yyyy")]
    public DateTime InDate { get; set; }
}
```

代碼

```csharp
var value = new Dto[] {
    new Issue241Dto{ Name="Jack",InDate=new DateTime(2021,01,04)},
    new Issue241Dto{ Name="Henry",InDate=new DateTime(2020,04,05)},
};
MiniExcel.SaveAs(path, value);
```

效果

![image](https://user-images.githubusercontent.com/12729184/118910788-ab2bcd80-b957-11eb-8d42-bfce36621b1b.png)

Query 支持自定義格式轉換

![image](https://user-images.githubusercontent.com/12729184/118911286-87b55280-b958-11eb-9a88-c8ff403d240a.png)

#### 3. 指定列寬(ExcelColumnWidthAttribute)

```csharp
public class Dto
{
    [ExcelColumnWidth(20)]
    public int ID { get; set; }
    [ExcelColumnWidth(15.50)]
    public string Name { get; set; }
}
```

#### 4. 多列名對應同一屬性

```csharp
public class Dto
{
    [ExcelColumnName(excelColumnName:"EmployeeNo",aliases:new[] { "EmpNo","No" })]
    public string Empno { get; set; }
    public string Name { get; set; }
}
```



#### 5. System.ComponentModel.DisplayNameAttribute = ExcelColumnName.excelColumnNameAttribute

從 1.24.0 開始支持 System.ComponentModel.DisplayNameAttribute 等同於 ExcelColumnName.excelColumnNameAttribute 效果

```C#
public class TestIssueI4TXGTDto
{
    public int ID { get; set; }
    public string Name { get; set; }
    [DisplayName("Specification")]
    public string Spc { get; set; }
    [DisplayName("Unit Price")]
    public decimal Up { get; set; }
}
```

#### 6. ExcelColumnAttribute

從 1.26.0 版本開始，可以簡化多Attribute寫法
```csharp
        public class TestIssueI4ZYUUDto
        {
            [ExcelColumn(Name = "ID",Index =0)]
            public string MyProperty { get; set; }
            [ExcelColumn(Name = "CreateDate", Index = 1,Format ="yyyy-MM",Width =100)]
            public DateTime MyProperty2 { get; set; }
        }
```

#### 7. DynamicColumnAttribute 動態設定 Column

從 1.26.0 版本開始，可以動態設定 Column 的屬性
```csharp
            var config = new OpenXmlConfiguration
            {
                DynamicColumns = new DynamicExcelColumn[] {
                    new DynamicExcelColumn("id"){Ignore=true},
                    new DynamicExcelColumn("name"){Index=1,Width=10},
                    new DynamicExcelColumn("createdate"){Index=0,Format="yyyy-MM-dd",Width=15},
                    new DynamicExcelColumn("point"){Index=2,Name="Account Point"},
                }
            };
            var path = PathHelper.GetTempPath();
            var value = new[] { new { id = 1, name = "Jack", createdate = new DateTime(2022, 04, 12) ,point = 123.456} };
            MiniExcel.SaveAs(path, value, configuration: config);
```
![image](https://user-images.githubusercontent.com/12729184/164510353-5aecbc4e-c3ce-41e8-b6cf-afd55eb23b68.png)





### 新增、刪除、修改

#### 新增

v1.28.0 開始支持 CSV 插入新增，在最後一行新增N筆數據

```csharp
// 原始數據
{
    var value = new[] {
          new { ID=1,Name ="Jack",InDate=new DateTime(2021,01,03)},
          new { ID=2,Name ="Henry",InDate=new DateTime(2020,05,03)},
    };
    MiniExcel.SaveAs(path, value);
}
// 最後一行新增一行數據
{
    var value = new { ID=3,Name = "Mike", InDate = new DateTime(2021, 04, 23) };
    MiniExcel.Insert(path, value);
}
// 最後一行新增N行數據
{
    var value = new[] {
          new { ID=4,Name ="Frank",InDate=new DateTime(2021,06,07)},
          new { ID=5,Name ="Gloria",InDate=new DateTime(2022,05,03)},
    };
    MiniExcel.Insert(path, value);
}
```

![image](https://user-images.githubusercontent.com/12729184/191023733-1e2fa732-db5c-4a3a-9722-b891fe5aa069.png)



#### 刪除(未完成)

#### 修改(未完成)



### Excel 類別自動判斷 <a name="getstart5"></a>

- MiniExcel 預設會根據`文件擴展名`判斷是 xlsx 還是 csv，但會有失準時候，請自行指定。
- Stream 類別無法判斷來源於哪種 excel 請自行指定

```csharp
stream.SaveAs(excelType:ExcelType.CSV);
//or
stream.SaveAs(excelType:ExcelType.XLSX);
//or
stream.Query(excelType:ExcelType.CSV);
//or
stream.Query(excelType:ExcelType.XLSX);
```



### CSV

#### 概念

- 預設全以字串類型返回，預設不會轉換為數字或者日期，除非有強型別定義泛型



#### 自定分隔符

預設以 `,` 作為分隔符，自定義請修改 `Seperator` 屬性

```csharp
var config = new MiniExcelLibs.Csv.CsvConfiguration()
{
    Seperator=';'
};
MiniExcel.SaveAs(path, values,configuration: config);
```



#### 自定義換行符

預設以 `\r\n` 作為換行符，自定義請修改 `NewLine` 屬性

```csharp
var config = new MiniExcelLibs.Csv.CsvConfiguration()
{
    NewLine='\n'
};
MiniExcel.SaveAs(path, values,configuration: config);
```

在 V1.30.1 版本開始支持動態更換換行符 (thanks @hyzx86)

```csharp
var config = new CsvConfiguration()
{
    SplitFn = (row) => Regex.Split(row, $"[\t,](?=(?:[^\"]|\"[^\"]*\")*$)")
        .Select(s => Regex.Replace(s.Replace("\"\"", "\""), "^\"|\"$", "")).ToArray()
};
var rows = MiniExcel.Query(path, configuration: config).ToList();
```

#### 自定義編碼

- 預設編碼為「從Byte順序標記檢測編碼」(detectEncodingFromByteOrderMarks: true)
- 有自定義編碼需求，請修改 StreamReaderFunc /  StreamWriterFunc 屬性

```csharp
// Read
var config = new MiniExcelLibs.Csv.CsvConfiguration()
{
    StreamReaderFunc = (stream) => new StreamReader(stream,Encoding.GetEncoding("gb2312"))
};
var rows = MiniExcel.Query(path, true,excelType:ExcelType.CSV,configuration: config);

// Write
var config = new MiniExcelLibs.Csv.CsvConfiguration()
{
    StreamWriterFunc = (stream) => new StreamWriter(stream, Encoding.GetEncoding("gb2312"))
};
MiniExcel.SaveAs(path, value,excelType:ExcelType.CSV, configuration: config);
```


### DataReader

#### 1. GetReader

从 1.23.0 版本开始能获取 DataReader

```csharp
    using (var reader = MiniExcel.GetReader(path,true))
    {
        while (reader.Read())
        {
            for (int i = 0; i < reader.FieldCount; i++)
            {
                var value = reader.GetValue(i);
            }
        }
    }
```



### 異步 Async

- 從 v0.17.0 版本開始支持異步 (感謝[isdaniel ( SHIH,BING-SIOU)](https://github.com/isdaniel))

```csharp
public static Task SaveAsAsync(string path, object value, bool printHeader = true, string sheetName = "Sheet1", ExcelType excelType = ExcelType.UNKNOWN, IConfiguration configuration = null)
public static Task SaveAsAsync(this Stream stream, object value, bool printHeader = true, string sheetName = "Sheet1", ExcelType excelType = ExcelType.XLSX, IConfiguration configuration = null)
public static Task<IEnumerable<dynamic>> QueryAsync(string path, bool useHeaderRow = false, string sheetName = null, ExcelType excelType = ExcelType.UNKNOWN, string startCell = "A1", IConfiguration configuration = null)
public static Task<IEnumerable<T>> QueryAsync<T>(this Stream stream, string sheetName = null, ExcelType excelType = ExcelType.UNKNOWN, string startCell = "A1", IConfiguration configuration = null) where T : class, new()
public static Task<IEnumerable<T>> QueryAsync<T>(string path, string sheetName = null, ExcelType excelType = ExcelType.UNKNOWN, string startCell = "A1", IConfiguration configuration = null) where T : class, new()
public static Task<IEnumerable<IDictionary<string, object>>> QueryAsync(this Stream stream, bool useHeaderRow = false, string sheetName = null, ExcelType excelType = ExcelType.UNKNOWN, string startCell = "A1", IConfiguration configuration = null)
public static Task SaveAsByTemplateAsync(this Stream stream, string templatePath, object value)
public static Task SaveAsByTemplateAsync(this Stream stream, byte[] templateBytes, object value)
public static Task SaveAsByTemplateAsync(string path, string templatePath, object value)
public static Task SaveAsByTemplateAsync(string path, byte[] templateBytes, object value)
public static Task<DataTable> QueryAsDataTableAsync(string path, bool useHeaderRow = true, string sheetName = null, ExcelType excelType = ExcelType.UNKNOWN, string startCell = "A1", IConfiguration configuration = null)
```

- 從 v1.25.0 開始支持 `cancellationToken`。



### 其他

#### 1. 映射枚舉(enum)

系統會自動映射(注意:大小寫不敏感)

![image](https://user-images.githubusercontent.com/12729184/116210595-9784b100-a775-11eb-936f-8e7a8b435961.png)

從V0.18.0版本開始支持Enum Description

```csharp
public class Dto
{
    public string Name { get; set; }
    public I49RYZUserType UserType { get; set; }
}

public enum Type
{
    [Description("General User")]
    V1,
    [Description("General Administrator")]
    V2,
    [Description("Super Administrator")]
    V3
}
```

![image](https://user-images.githubusercontent.com/12729184/133116630-27cc7161-099a-48b8-9784-cd1e443af3d1.png)

從 1.30.0 版本開始支持由 Description 轉回 Enum 功能，感謝 @KaneLeung



#### 2. CSV 轉 XLSX 或是 XLSX 轉 CSV

```csharp
MiniExcel.ConvertXlsxToCsv(xlsxPath, csvPath);
MiniExcel.ConvertXlsxToCsv(xlsxStream, csvStream);
MiniExcel.ConvertCsvToXlsx(csvPath, xlsxPath);
MiniExcel.ConvertCsvToXlsx(csvStream, xlsxStream);
```
```csharp
using (var excelStream = new FileStream(path: filePath, FileMode.Open, FileAccess.Read))
using (var csvStream = new MemoryStream())
{
   MiniExcel.ConvertXlsxToCsv(excelStream, csvStream);
}
```

#### 3. 自定義 CultureInfo

從 1.22.0 版本開始，可以使用以下代碼自定義文化資訊，系統預設 `CultureInfo.InvariantCulture`。

```csharp
var config = new CsvConfiguration()
{
    Culture = new CultureInfo("fr-FR"),
};
MiniExcel.SaveAs(path, value, configuration: config);

// or
MiniExcel.Query(path,configuration: config);
```

#### 4. 導出自定義 Buffer Size
```csharp
    public abstract class Configuration : IConfiguration
    {
        public int BufferSize { get; set; } = 1024 * 512;
    }
```

#### 5. FastMode

系統不會限制記憶體，達到更快的效率

```csharp
var config = new OpenXmlConfiguration() { FastMode = true };
MiniExcel.SaveAs(path, reader,configuration:config);
```



### 範例

#### 1. SQLite & Dapper 讀取大數據新增到資料庫

Note : 請不要呼叫 call ToList/ToArray 等方法，這會將所有資料讀到記憶體內

```csharp
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


#### 2. ASP.NET Core 3.1 or MVC 5 下載/上傳 Excel Xlsx API Demo [Try it](tests/MiniExcel.Tests.AspNetCore)

```csharp
public class ApiController : Controller
{
    public IActionResult Index()
    {
        return new ContentResult
        {
            ContentType = "text/html",
            StatusCode = (int)HttpStatusCode.OK,
            Content = @"<html><body>
<a href='api/DownloadExcel'>DownloadExcel</a><br>
<a href='api/DownloadExcelFromTemplatePath'>DownloadExcelFromTemplatePath</a><br>
<a href='api/DownloadExcelFromTemplateBytes'>DownloadExcelFromTemplateBytes</a><br>
<p>Upload Excel</p>
<form method='post' enctype='multipart/form-data' action='/api/uploadexcel'>
    <input type='file' name='excel'> <br>
    <input type='submit' >
</form>
</body></html>"
        };
    }

    public IActionResult DownloadExcel()
    {
        var values = new[] {
            new { Column1 = "MiniExcel", Column2 = 1 },
            new { Column1 = "Github", Column2 = 2}
        };
        var memoryStream = new MemoryStream();
        memoryStream.SaveAs(values);
        memoryStream.Seek(0, SeekOrigin.Begin);
        return new FileStreamResult(memoryStream, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        {
            FileDownloadName = "demo.xlsx"
        };
    }

    public IActionResult DownloadExcelFromTemplatePath()
    {
        string templatePath = "TestTemplateComplex.xlsx";

        Dictionary<string, object> value = new Dictionary<string, object>()
        {
            ["title"] = "FooCompany",
            ["managers"] = new[] {
                new {name="Jack",department="HR"},
                new {name="Loan",department="IT"}
            },
            ["employees"] = new[] {
                new {name="Wade",department="HR"},
                new {name="Felix",department="HR"},
                new {name="Eric",department="IT"},
                new {name="Keaton",department="IT"}
            }
        };

        MemoryStream memoryStream = new MemoryStream();
        memoryStream.SaveAsByTemplate(templatePath, value);
        memoryStream.Seek(0, SeekOrigin.Begin);
        return new FileStreamResult(memoryStream, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        {
            FileDownloadName = "demo.xlsx"
        };
    }

    private static Dictionary<string, Byte[]> TemplateBytesCache = new Dictionary<string, byte[]>();

    static ApiController()
    {
        string templatePath = "TestTemplateComplex.xlsx";
        byte[] bytes = System.IO.File.ReadAllBytes(templatePath);
        TemplateBytesCache.Add(templatePath, bytes);
    }

    public IActionResult DownloadExcelFromTemplateBytes()
    {
        byte[] bytes = TemplateBytesCache["TestTemplateComplex.xlsx"];

        Dictionary<string, object> value = new Dictionary<string, object>()
        {
            ["title"] = "FooCompany",
            ["managers"] = new[] {
                new {name="Jack",department="HR"},
                new {name="Loan",department="IT"}
            },
            ["employees"] = new[] {
                new {name="Wade",department="HR"},
                new {name="Felix",department="HR"},
                new {name="Eric",department="IT"},
                new {name="Keaton",department="IT"}
            }
        };

        MemoryStream memoryStream = new MemoryStream();
        memoryStream.SaveAsByTemplate(bytes, value);
        memoryStream.Seek(0, SeekOrigin.Begin);
        return new FileStreamResult(memoryStream, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        {
            FileDownloadName = "demo.xlsx"
        };
    }

    public IActionResult UploadExcel(IFormFile excel)
    {
        var stream = new MemoryStream();
        excel.CopyTo(stream);

        foreach (var item in stream.Query(true))
        {
            // do your logic etc.
        }

        return Ok("File uploaded successfully");
    }
}
```

####  3. 分頁查詢

```csharp
void Main()
{
    var rows = MiniExcel.Query(path);

    Console.WriteLine("==== No.1 Page ====");
    Console.WriteLine(Page(rows,pageSize:3,page:1));
    Console.WriteLine("==== No.50 Page ====");
    Console.WriteLine(Page(rows,pageSize:3,page:50));
    Console.WriteLine("==== No.5000 Page ====");
    Console.WriteLine(Page(rows,pageSize:3,page:5000));
}

public static IEnumerable<T> Page<T>(IEnumerable<T> en, int pageSize, int page)
{
    return en.Skip(page * pageSize).Take(pageSize);
}
```

![20210419](https://user-images.githubusercontent.com/12729184/114679083-6ef4c400-9d3e-11eb-9f78-a86daa45fe46.gif)

#### 4. WebForm不落地導出Excel

```csharp
var fileName = "Demo.xlsx";
var sheetName = "Sheet1";
HttpResponse response = HttpContext.Current.Response;
response.Clear();
response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
response.AddHeader("Content-Disposition", $"attachment;filename=\"{fileName}\"");
var values = new[] {
    new { Column1 = "MiniExcel", Column2 = 1 },
    new { Column1 = "Github", Column2 = 2}
};
var memoryStream = new MemoryStream();
memoryStream.SaveAs(values, sheetName: sheetName);
memoryStream.Seek(0, SeekOrigin.Begin);
memoryStream.CopyTo(Response.OutputStream);
response.End();
```

#### 5. 動態 i18n 多國語言跟權限管理

像例子一樣，建立一個方法處理 i18n 跟權限管理，並搭配 `yield return 返回 IEnumerable<Dictionary<string, object>>`，即可達到動態、低記憶體處理效果

```csharp
void Main()
{
    var value = new Order[] {
        new Order(){OrderNo = "SO01",CustomerID="C001",ProductID="P001",Qty=100,Amt=500},
        new Order(){OrderNo = "SO02",CustomerID="C002",ProductID="P002",Qty=300,Amt=400},
    };

    Console.WriteLine("en-Us and Sales role");
    {
        var path = Path.GetTempPath() + Guid.NewGuid() + ".xlsx";
        var lang = "en-US";
        var role = "Sales";
        MiniExcel.SaveAs(path, GetOrders(lang, role, value));
        MiniExcel.Query(path, true).Dump();
    }

    Console.WriteLine("zh-CN and PMC role");
    {
        var path = Path.GetTempPath() + Guid.NewGuid() + ".xlsx";
        var lang = "zh-CN";
        var role = "PMC";
        MiniExcel.SaveAs(path, GetOrders(lang, role, value));
        MiniExcel.Query(path, true).Dump();
    }
}

private IEnumerable<Dictionary<string, object>> GetOrders(string lang, string role, Order[] orders)
{
    foreach (var order in orders)
    {
        var newOrder = new Dictionary<string, object>();

        if (lang == "zh-CN")
        {
            newOrder.Add("客戶編號", order.CustomerID);
            newOrder.Add("訂單編號", order.OrderNo);
            newOrder.Add("產品編號", order.ProductID);
            newOrder.Add("數量", order.Qty);
            if (role == "Sales")
                newOrder.Add("價格", order.Amt);
            yield return newOrder;
        }
        else if (lang == "en-US")
        {
            newOrder.Add("Customer ID", order.CustomerID);
            newOrder.Add("Order No", order.OrderNo);
            newOrder.Add("Product ID", order.ProductID);
            newOrder.Add("Quantity", order.Qty);
            if (role == "Sales")
                newOrder.Add("Amount", order.Amt);
            yield return newOrder;
        }
        else
        {
            throw new InvalidDataException($"lang {lang} wrong");
        }
    }
}

public class Order
{
    public string OrderNo { get; set; }
    public string CustomerID { get; set; }
    public decimal Qty { get; set; }
    public string ProductID { get; set; }
    public decimal Amt { get; set; }
}
```

![image](https://user-images.githubusercontent.com/12729184/118939964-d24bc480-b982-11eb-88dd-f06655f6121a.png)



### FAQ 常見問題

#### Q: Excel 表頭標題名稱跟 class 屬性名稱不一致，如何對應?

A. 請使用 ExcelColumnName 作 mapping

![image](https://user-images.githubusercontent.com/12729184/116020475-eac50980-a678-11eb-8804-129e87200e5e.png)



#### Q. 多工作表(sheet)如何導出/查詢資料?

A. 使用 `GetSheetNames `方法搭配 Query 的 sheetName 參數



```csharp
var sheets = MiniExcel.GetSheetNames(path);
foreach (var sheet in sheets)
{
    Console.WriteLine($"sheet name : {sheet} ");
    var rows = MiniExcel.Query(path,useHeaderRow:true,sheetName:sheet);
    Console.WriteLine(rows);
}
```

![image](https://user-images.githubusercontent.com/12729184/116199841-2a1f5300-a76a-11eb-90a3-6710561cf6db.png)



#### Q. 是否使用 Count 會載入全部數據到記憶體

不會，圖片測試一百萬行*十列資料，簡單測試，內存最大使用 < 60MB，花費13.65秒

![image](https://user-images.githubusercontent.com/12729184/117118518-70586000-adc3-11eb-9ce3-2ba76cf8b5e5.png)

#### Q. Query如何使用整數索引取值?

Query 預設索引為字串Key : A,B,C....，想要改為數字索引，請建立以下方法自行轉換

```csharp
void Main()
{
    var path = @"D:\git\MiniExcel\samples\xlsx\TestTypeMapping.xlsx";
    var rows = MiniExcel.Query(path,true);
    foreach (var r in ConvertToIntIndexRows(rows))
    {
        Console.Write($"column 0 : {r[0]} ,column 1 : {r[1]}");
        Console.WriteLine();
    }
}

private IEnumerable<Dictionary<int, object>> ConvertToIntIndexRows(IEnumerable<object> rows)
{
    ICollection<string> keys = null;
    var isFirst = true;
    foreach (IDictionary<string,object> r in rows)
    {
        if(isFirst)
        {
            keys = r.Keys;
            isFirst = false;
        }

        var dic = new Dictionary<int, object>();
        var index = 0;
        foreach (var key in keys)
            dic[index++] = r[key];
        yield return dic;
    }
}
```

#### Q. 導出時數組為空時生成沒有標題空 Excel

因為 MiniExcel 使用類似 JSON.NET 動態從值獲取類別機制簡化 API 操作，沒有數據就無法獲取類別。可以查看[ issue #133](https://github.com/shps951023/MiniExcel/issues/133) 了解。

![image](https://user-images.githubusercontent.com/12729184/122639771-546c0c00-d12e-11eb-800c-498db27889ca.png)

> 強型別和 DataTable 會生成表頭，但 Dicionary 依舊是空 Excel

#### Q. 如何人為空白行中止遍歷?

常發生人為不小心在最後幾行留下空白行情況，MiniExcel可以搭配 `LINQ TakeWhile`實現空白行中斷遍歷。

![image](https://user-images.githubusercontent.com/12729184/130209137-162621c2-f337-4479-9996-beeac65bc4d4.png)

#### Q. 不想要空白行如何去除?
![image](https://user-images.githubusercontent.com/12729184/137873865-7107d8f5-eb59-42db-903a-44e80589f1b2.png)


IEnumerable版本

```csharp
public static IEnumerable<dynamic> QueryWithoutEmptyRow(Stream stream, bool useHeaderRow, string sheetName, ExcelType excelType, string startCell, IConfiguration configuration)
{
    var rows = stream.Query(useHeaderRow,sheetName,excelType,startCell,configuration);
    foreach (IDictionary<string,object> row in rows)
    {
        if(row.Keys.Any(key=>row[key]!=null))
            yield return row;
    }
}
```



DataTable版本

```csharp
public static DataTable QueryAsDataTableWithoutEmptyRow(Stream stream, bool useHeaderRow, string sheetName, ExcelType excelType, string startCell, IConfiguration configuration)
{
    if (sheetName == null && excelType != ExcelType.CSV) /*Issue #279*/
        sheetName = stream.GetSheetNames().First();

    var dt = new DataTable(sheetName);
    var first = true;
    var rows = stream.Query(useHeaderRow,sheetName,excelType,startCell,configuration);
    foreach (IDictionary<string, object> row in rows)
    {
        if (first)
        {

            foreach (var key in row.Keys)
            {
                var column = new DataColumn(key, typeof(object)) { Caption = key };
                dt.Columns.Add(column);
            }

            dt.BeginLoadData();
            first = false;
        }

        var newRow = dt.NewRow();
        var isNull=true;
        foreach (var key in row.Keys)
        {
            var _v = row[key];
            if(_v!=null)
                isNull = false;
            newRow[key] = _v;
        }

        if(!isNull)
            dt.Rows.Add(newRow);
    }

    dt.EndLoadData();
    return dt;
}
```



#### Q. 保存如何取代MiniExcel.SaveAs(path, value)，文件存在系統會報已存在錯誤?

請改以Stream自行管控Stream行為，如

```C#
    using (var stream = File.Create("Demo.xlsx"))
        MiniExcel.SaveAs(stream,value);
```



從V1.25.0版本開始，支持 overwriteFile 參數，方便調整是否要覆蓋已存在文件

```csharp
    MiniExcel.SaveAs(path, value, overwriteFile: true);
```



### 侷限與警告

- 目前不支援 xls (97-2003) 或是加密檔案
- xlsm 只支持查詢



### 參考

[ExcelDataReader](https://github.com/ExcelDataReader/ExcelDataReader)  / [ClosedXML](https://github.com/ClosedXML/ClosedXML) / [Dapper](https://github.com/DapperLib/Dapper) / [ExcelNumberFormat](https://github.com/andersnm/ExcelNumberFormat)



### 感謝名單

####  [Jetbrains](https://www.jetbrains.com/)

![jetbrains-variant-2](https://user-images.githubusercontent.com/12729184/123997015-8456c180-da02-11eb-829a-aec476fe8e94.png)

感謝提供免費IDE支持此專案 ([License](https://user-images.githubusercontent.com/12729184/123988233-6ab17c00-d9fa-11eb-8739-2a08c6a4a263.png))

### 收益流水
目前收益 https://github.com/mini-software/MiniExcel/issues/560#issue-2080619180



### Contributors

![](https://contrib.rocks/image?repo=shps951023/MiniExcel)