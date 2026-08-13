# MiniExcel Rust XLSX MVP

该目录包含 MiniExcel 基础 XLSX 读写流程的实验性 Rust 实现。目前它属于研究分支，尚未发布到 crates.io，也不会替代现有 .NET 包。

[English](README.md)

## 当前能力

- 从路径读取 `.xlsx`。
- 通过 `MiniExcel::query()` 和 `MiniExcel::query_as()` 以有界内存流式读取 worksheet。
- 枚举工作表并按名称选择工作表。
- 使用稳定列顺序的动态行，可选首行表头。
- 通过 Serde 将行反序列化为 Rust 结构体。
- 支持 A1 起始单元格、表头修剪和可选空行过滤。
- 从动态行或 Serde 结构体创建新的 `.xlsx` 工作簿。
- 读取时可选择工作表，写入输出到文件路径。
- 支持字符串、布尔值、整数、浮点数、空单元格、Excel 错误、日期、时间、日期时间和时长。

项目使用 Rust 2024，最低支持 Rust 1.85.0。

## 构建

在仓库根目录运行：

```bash
cargo +1.85.0 check --manifest-path rust/Cargo.toml --workspace --all-targets --locked
cargo test --manifest-path rust/Cargo.toml --workspace --all-targets --locked
```

仓库会提交 workspace 的 `Cargo.lock`，确保本地研究与 CI 使用同一依赖图。

## 公开 API

`MiniExcel` 是唯一公开的行为入口。reader、writer、ZIP/XML parser 和 iterator 具体类型全部保留在 crate 内部。crate 根其余导出仅为数据与配置契约：`CellValue`、`DynamicRow`、`CellReference`、`ReadOptions`、`WriteOptions`、`HeaderMode`、`Error` 和 `Result`。日期/时间 Serde adapter 位于 `serde_helpers`。

## 简洁的流式 Query

最接近 `MiniExcel.Query` 的 Rust 写法是返回迭代器：

```rust
use miniexcel::MiniExcel;

for row in MiniExcel::query("book.xlsx")? {
    let row = row?;
    println!("{:?}", row["A"]);
}
# Ok::<(), miniexcel::Error>(())
```

worksheet XML 会被增量解压和解析，行数据通过有界 channel 交给迭代器，并随迭代逐行映射。因此可直接使用 `take`、`filter`、`find` 等操作，不必收集全部行；丢弃迭代器会停止其 worker。工作表、表头、起始单元格和空行选项可通过 `MiniExcel::query_with_options()` 设置。

类型化查询使用同一模式：

```rust
# use serde::Deserialize;
use miniexcel::MiniExcel;

#[derive(Deserialize)]
#[serde(rename_all = "PascalCase")]
struct Record {
    name: String,
}

for record in MiniExcel::query_as::<Record>("book.xlsx")? {
    println!("{}", record?.name);
}
# Ok::<(), miniexcel::Error>(())
```

`MiniExcel::query()` 和 `query_as()` 接收路径，因为迭代器存活期间由 worker 持有 ZIP archive。具体 iterator 类型刻意隐藏在 crate 内部。

> **内存边界：** 流式路径会在内存中保留工作簿元数据、样式、shared-string table、少量行 channel 和 parser buffer，但不会保留完整 worksheet XML 或所有行。为了在 `<dimension>` 缺失/过期时仍提供稳定的全局列 schema，并排除末尾仅带样式的空行，它会先做一次有界内存元数据扫描，再进行流式输出。峰值内存仍可能随 shared-string table 或单个超大行增长，但不会随 worksheet 总行数增长。

## 动态读取

```rust
use miniexcel::{HeaderMode, MiniExcel, ReadOptions};

let options = ReadOptions::new()
    .with_sheet_name("Data")
    .with_header_mode(HeaderMode::FirstRow);

for row in MiniExcel::query_with_options("book.xlsx", &options)? {
    println!("{:?}", row?["Name"]);
}
# Ok::<(), miniexcel::Error>(())
```

默认的 `HeaderMode::Auto` 表示：`query()` 默认没有表头，`query_as()` 默认使用第一行作为表头。

没有表头时，动态键使用真实 Excel 列名，例如 `A`、`B`、`AA`。为了兼容 MiniExcel，默认保留空行；可通过 `with_ignore_empty_rows(true)` 删除所有单元格都为空的行。

## 类型化读取

```rust
use chrono::NaiveDate;
use miniexcel::MiniExcel;
use serde::Deserialize;

#[derive(Deserialize)]
#[serde(rename_all = "PascalCase")]
struct Release {
    name: String,
    version: u32,
    #[serde(deserialize_with = "miniexcel::serde_helpers::deserialize_date")]
    released_on: NaiveDate,
}

let rows = MiniExcel::query_as::<Release>("book.xlsx")?
    .collect::<miniexcel::Result<Vec<_>>>()?;
# Ok::<(), miniexcel::Error>(())
```

支持 Serde 的 `rename`、`alias`、`default`、`skip` 和 `Option` 语义。首期不移植 MiniExcel 专用的列索引 Attribute。

## 动态写入

```rust
use miniexcel::{CellValue, DynamicRow, MiniExcel, WriteOptions};

let mut row = DynamicRow::new();
row.insert("Name".to_owned(), CellValue::String("MiniExcel".to_owned()));
row.insert("Version".to_owned(), CellValue::Int(2));

MiniExcel::save_as_with_options(
    "book.xlsx",
    &[row],
    &WriteOptions::new().with_sheet_name("Data"),
)?;
# Ok::<(), miniexcel::Error>(())
```

动态 schema 按所有行中键第一次出现的顺序合并，缺失值写为空单元格。需要显式 schema 或仅写表头时，请使用 `MiniExcel::save_as_with_schema()`。

## 类型化写入

```rust
use chrono::NaiveDate;
use miniexcel::{MiniExcel, WriteOptions};
use serde::Serialize;

#[derive(Serialize)]
#[serde(rename_all = "PascalCase")]
struct Release {
    name: String,
    #[serde(serialize_with = "miniexcel::serde_helpers::serialize_date_to_excel")]
    released_on: NaiveDate,
}

let values = [Release {
    name: "MiniExcel Rust".to_owned(),
    released_on: NaiveDate::from_ymd_opt(2026, 8, 13).unwrap(),
}];
let options = WriteOptions::new()
    .with_sheet_name("Releases")
    .with_column_format("ReleasedOn", "yyyy-mm-dd");

MiniExcel::save_as_serialized_with_options("releases.xlsx", &values, &options)?;
# Ok::<(), miniexcel::Error>(())
```

列格式的键是经过 Serde 重命名后的最终字段/表头名称。类型化写入支持结构体及结构体集合；Map 和 `flatten` 应改用动态 API。

## 重要语义

- 未指定工作表时选择工作簿顺序中的第一张表，而不是 active tab。
- 能精确表示为 `i64` 的 XLSX 数值返回 `CellValue::Int`，其他数值返回 `Float`。
- Excel 序列日期不总能区分纯日期、纯时间和日期时间，因此动态读取统一为 `CellValue::DateTime`；ISO 值会尽量保留更具体的类型。
- 公式只读取缓存值，不返回公式表达式。
- `MiniExcel::query()` 和 `query_as()` 会从路径严格流式解析 worksheet XML。
- 流式查询是同步接口，每个活动 query 使用一个 worker thread；首期不包含 async I/O。
- 写入只创建新工作簿并覆盖目标路径，不能修改已有工作簿。

## 首期不包含

CSV、`.xls`、`.xlsb`、`.ods`、模板、宏、图片、合并单元格操作、结束范围坐标、公式写入、通用样式系统、WASM 和修改已有工作簿均延后实现。

依赖选择和行为对照请查看[兼容性研究记录](docs/compatibility.md)。