# MiniExcel Rust XLSX MVP

This directory contains an experimental Rust implementation of MiniExcel's basic XLSX read and write workflows. It is a research track, is not published to crates.io, and does not replace the .NET packages.

[简体中文](README.zh-CN.md)

## Status

The MVP currently supports:

- Reading `.xlsx` files from paths.
- Bounded-memory worksheet streaming through `MiniExcel::query()` and `MiniExcel::query_as()`.
- Listing worksheets and selecting a worksheet by name.
- Dynamic rows with stable column order and optional header rows.
- Typed row deserialization through Serde.
- A1 start cells, header trimming, and optional empty-row filtering.
- Creating new `.xlsx` workbooks from dynamic rows or Serde structs.
- Worksheet selection for reads and path-based workbook output.
- Strings, booleans, integers, floating-point values, empty cells, Excel errors, dates, times, datetimes, and durations.

The implementation uses Rust 2024 with an MSRV of Rust 1.85.0.

## Build

Run commands from the repository root:

```bash
cargo +1.85.0 check --manifest-path rust/Cargo.toml --workspace --all-targets --locked
cargo test --manifest-path rust/Cargo.toml --workspace --all-targets --locked
```

The workspace lockfile is committed so CI and local research use the same dependency graph.

## .NET Parity

.NET and Rust consume the same versioned behavior contract at `tests/data/contracts/xlsx-parity-v1.json`. It covers the common dynamic and typed query surface with the same XLSX fixtures and canonical expected values.

```bash
cargo +1.85.0 test --manifest-path rust/Cargo.toml -p miniexcel --test parity_contract --locked
dotnet test tests/MiniExcel.OpenXml.Tests/MiniExcel.OpenXml.Tests.csproj --framework net10.0 --filter "FullyQualifiedName~RustParityContractTests"
```

Both commands must pass for a behavior to be considered equivalent. See [Compatibility and research notes](docs/compatibility.md#net-parity-contract) for normalization rules and the explicit version 1 scope.

## Public API

`MiniExcel` is the only public behavior entry point. Reader, writer, ZIP/XML parser, and iterator implementation types are internal. The remaining root exports are data and configuration contracts: `CellValue`, `DynamicRow`, `CellReference`, `ReadOptions`, `WriteOptions`, `HeaderMode`, `Error`, and `Result`. Date/time Serde adapters are available under `serde_helpers`.

## Simple Streaming Query

The closest Rust equivalent to `MiniExcel.Query` is an iterator:

```rust
use miniexcel::MiniExcel;

for row in MiniExcel::query("book.xlsx")? {
    let row = row?;
    println!("{:?}", row["A"]);
}
# Ok::<(), miniexcel::Error>(())
```

Worksheet XML is decompressed and parsed incrementally. Rows are delivered through a bounded channel and mapped as the iterator advances, so callers can use operations such as `take`, `filter`, and `find` without collecting every row. Dropping the iterator stops its worker. Use `MiniExcel::query_with_options()` for worksheet, header, start-cell, and empty-row options.

Typed rows use the same model:

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

`MiniExcel::query()` and `query_as()` accept paths because a worker owns the ZIP archive while the iterator is alive. Their concrete iterator types are intentionally hidden.

> **Memory boundary:** the streaming path keeps workbook metadata, styles, and the shared-string table in memory, plus a small row channel and parser buffers. It does not retain worksheet XML or all worksheet rows. It performs one bounded-memory metadata pass before the streaming pass so every dynamic row has a stable global column schema and explicitly declared style-only rows are preserved even when `<dimension>` is missing or stale. Peak memory can still grow with the shared-string table or a single exceptionally large row, but not with the full worksheet row count.

## Dynamic Reading

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

`HeaderMode::Auto` is the default. It means no header for `query()` and a first-row header for `query_as()`.

Without headers, dynamic keys use the actual Excel column names such as `A`, `B`, and `AA`. Empty rows are retained by default to match MiniExcel. Use `with_ignore_empty_rows(true)` to filter rows whose cells are all empty.

## Typed Reading

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

Serde `rename`, `alias`, `default`, `skip`, and `Option` semantics are supported. MiniExcel-specific column-index attributes are not part of the MVP.

## Dynamic Writing

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

Dynamic schemas are the union of row keys in first-seen order. Missing values are written as blank cells. Use `MiniExcel::save_as_with_schema()` when an explicit schema is required, including header-only exports.

## Typed Writing

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

The column-format key is the final Serde field/header name. Typed Serde writing supports structs and vectors of structs; maps and `flatten` are handled through the dynamic API instead.

## Important Semantics

- The default worksheet is the first workbook worksheet, not the active tab.
- Dynamic XLSX numbers with an exact `i64` representation are returned as `CellValue::Int`; other numeric values remain `Float`.
- Excel serial dates cannot always distinguish date-only, time-only, and datetime intent. Dynamic serial values are normalized to `CellValue::DateTime`; ISO values retain the more specific variant when possible.
- Formula expressions are not returned. Reading uses their cached values.
- `MiniExcel::query()` and `query_as()` strictly stream worksheet XML from paths.
- Streaming is synchronous and uses one worker thread per active query. Async I/O is not part of the MVP.
- Writing creates new workbooks and overwrites target paths. It cannot modify an existing workbook.

## Non-Goals For This MVP

CSV, `.xls`, `.xlsb`, `.ods`, templates, macros, images, merged-cell operations, arbitrary range end coordinates, formula authoring, a general style system, WASM, and editing existing workbooks are deferred.

See [Compatibility and research notes](docs/compatibility.md) for dependency choices and behavior mapping.