## MiniExcel

<div>
    <a href="https://www.nuget.org/packages/MiniExcel">
        <img src="https://img.shields.io/nuget/v/MiniExcel.svg" alt="NuGet">
    </a>
    <a href="https://www.nuget.org/packages/MiniExcel">
        <img src="https://img.shields.io/nuget/dt/MiniExcel.svg" alt="">
    </a>
    <a href="https://ci.appveyor.com/project/mini-software/miniexcel/branch/master">
        <img src="https://ci.appveyor.com/api/projects/status/b2vustrwsuqx45f4/branch/master?svg=true" alt="Build status">
    </a>
    <a href="https://gitee.com/dotnetchina/MiniExcel">
        <img src="https://gitee.com/dotnetchina/MiniExcel/badge/star.svg" alt="star">
    </a>
    <a href="https://github.com/mini-software/MiniExcel" rel="nofollow">
        <img src="https://img.shields.io/github/stars/mini-software/MiniExcel?logo=github" alt="GitHub stars">
    </a>
    <a href="https://www.nuget.org/packages/MiniExcel">
        <img src="https://img.shields.io/badge/.NET-%3E%3D%204.5-red.svg" alt="version">
    </a>
    <a href="https://deepwiki.com/mini-software/MiniExcel">
        <img src="https://deepwiki.com/badge.svg" alt="Ask DeepWiki">
    </a>
</div>

---

MiniExcel is a simple and efficient Excel processing tool for .NET, specifically designed to minimize memory usage.

At present, most popular frameworks need to load all the data from an Excel document into memory to facilitate operations, but this may cause memory consumption problems. MiniExcel's approach is different: the data is processed row by row in a streaming manner, reducing the original consumption from potentially hundreds of megabytes to just a few megabytes, effectively preventing out-of-memory(OOM) issues.

```mermaid
flowchart LR
    A1(["Excel analysis<br>process"]) --> A2{{"Unzipping<br>XLSX file"}} --> A3{{"Parsing<br>OpenXML"}} --> A4{{"Model<br>conversion"}} --> A5(["Output"])

    B1(["Other Excel<br>Frameworks"]) --> B2{{"Memory"}} --> B3{{"Memory"}} --> B4{{"Workbooks &<br>Worksheets"}} --> B5(["All rows at<br>the same time"])

    C1(["MiniExcel"]) --> C2{{"Stream"}} --> C3{{"Stream"}} --> C4{{"POCO or dynamic"}} --> C5(["Deferred execution<br>row by row"])

    classDef analysis fill:#D0E8FF,stroke:#1E88E5,color:#0D47A1,font-weight:bold;
    classDef others fill:#FCE4EC,stroke:#EC407A,color:#880E4F,font-weight:bold;
    classDef miniexcel fill:#E8F5E9,stroke:#388E3C,color:#1B5E20,font-weight:bold;

    class A1,A2,A3,A4,A5 analysis;
    class B1,B2,B3,B4,B5 others;
    class C1,C2,C3,C4,C5 miniexcel;
```

### Features

- Minimizes memory consumption, preventing out-of-memory (OOM) errors and avoiding full garbage collections
- Enables real-time, row-level data operations for better performance on large datasets
- Supports LINQ with deferred execution, allowing for fast, memory-efficient paging and complex queries
- Lightweight, without the need for Microsoft Office or COM+ components, and a size under 800KB
- Simple and intuitive API style to import, export, and template Excel worksheets

### Quickstart

#### Importing

You can query worksheets and map the data either to strongly typed classes or dynamic objects:

```csharp
public class UserAccount
{
    public Guid ID { get; set; }
    public string Name { get; set; }
    public DateTime DateOfBirth { get; set; }
    public int Age { get; set; }
    public bool Vip { get; set; }
    public decimal Points { get; set; }
}

var userRows = MiniExcel.Query<UserAccount>(path);

// or simply

var dynamicRows = MiniExcel.Query(path);
```

#### Exporting

There are multiple ways to exprt data to an Excel document:

```csharp
// From strongly typed objects

var values = new[]
{
    new { Name = "MiniExcel", Value = 1 },
    new { Name = "Github", Value = 2 }
};
MiniExcel.SaveAs(yourPath, values);


// From anonymous objects

public class TestType
{
    public string Name { get; set; }
    public int Value { get; set; }
}

TestType[] values =
[
    new TestType { Name = "MiniExcel", Value = 1 },
    new TestType { Name = "Github", Value = 2 }
];
MiniExcel.SaveAs(yourPath, values);


//From a IEnumerable<IDictionary<string, object>>

new List<Dictionary<string, object>>() dicts =
[
    new Dictionary<string, object> { { "Name", "MiniExcel" }, { "Value", 1 } },
    new Dictionary<string, object> { { "Name", "Github" }, { "Value", 2 } }
];
MiniExcel.SaveAs(yourPath, dicts);


// Directly from a IDataReader

using var connection = yourConnectionProvider.GetConnection();
connection.Open();

using var cmd = connection.CreateCommand();
cmd.CommandText = """
    SELECT 'MiniExcel' AS "Name", 1 AS "Value"
    UNION ALL
    SELECT 'Github', 2
    """;

using var reader = cmd.ExecuteReader();
MiniExcel.SaveAs(yourPath, reader);


// From a DataTable

var table = new DataTable();
table.Columns.Add("Name", typeof(string));
table.Columns.Add("Value", typeof(int));
table.Rows.Add("MiniExcel", 1);
table.Rows.Add("Github", 2);

MiniExcel.SaveAs(path, table);
```
