# MiniExcel Rust XLSX MVP

This directory contains an experimental Rust implementation of MiniExcel's basic XLSX read and write workflows. It is a research track, is not published to crates.io, and does not replace the .NET packages.

[简体中文](README.zh-CN.md)

## Status

The MVP currently supports:

- Reading `.xlsx` files from paths or `Read + Seek` sources.
- Bounded-memory worksheet streaming through `MiniExcel::query()` and `MiniExcel::query_as()`.
- Listing worksheets and selecting a worksheet by name.
- Dynamic rows with stable column order and optional header rows.
- Typed row deserialization through Serde.
- A1 start cells, header trimming, and optional empty-row filtering.
- Creating new `.xlsx` workbooks from dynamic rows or Serde structs.
- Multiple worksheets and output to paths, byte buffers, or `Write + Send` targets.
- Strings, booleans, integers, floating-point values, empty cells, Excel errors, dates, times, datetimes, and durations.

The implementation uses Rust 2024 with an MSRV of Rust 1.85.0.

## Build

Run commands from the repository root:

```bash
cargo +1.85.0 check --manifest-path rust/Cargo.toml --workspace --all-targets --locked
cargo test --manifest-path rust/Cargo.toml --workspace --all-targets --locked
```

The workspace lockfile is committed so CI and local research use the same dependency graph.

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

`MiniExcel::query()` and `query_as()` currently accept paths because a worker owns the ZIP archive while the iterator is alive. `XlsxReader::from_reader()` remains available for arbitrary `Read + Seek` sources, but that compatibility path uses calamine and materializes the selected worksheet.

> **Memory boundary:** the streaming path keeps workbook metadata, styles, and the shared-string table in memory, plus a small row channel and parser buffers. It does not retain worksheet XML or all worksheet rows. It performs one bounded-memory metadata pass before the streaming pass so every dynamic row has a stable global column schema and trailing style-only rows are excluded even when `<dimension>` is missing or stale. Peak memory can still grow with the shared-string table or a single exceptionally large row, but not with the full worksheet row count.

## Dynamic Reading

```rust
use miniexcel::{HeaderMode, ReadOptions, XlsxReader};

let mut reader = XlsxReader::open("book.xlsx")?;
let options = ReadOptions::new()
    .with_sheet_name("Data")
    .with_header_mode(HeaderMode::FirstRow);

let rows = reader.read_rows(&options)?;
println!("{:?}", rows[0]["Name"]);
# Ok::<(), miniexcel::Error>(())
```

`HeaderMode::Auto` is the default. It means no header for `read_rows()` and a first-row header for `deserialize()`.

Without headers, dynamic keys use the actual Excel column names such as `A`, `B`, and `AA`. Empty rows are retained by default to match MiniExcel. Use `with_ignore_empty_rows(true)` to filter rows whose cells are all empty.

## Typed Reading

```rust
use chrono::NaiveDate;
use miniexcel::{ReadOptions, XlsxReader};
use serde::Deserialize;

#[derive(Deserialize)]
#[serde(rename_all = "PascalCase")]
struct Release {
    name: String,
    version: u32,
    #[serde(deserialize_with = "miniexcel::serde_helpers::deserialize_date")]
    released_on: NaiveDate,
}

let mut reader = XlsxReader::open("book.xlsx")?;
let rows: Vec<Release> = reader.deserialize(&ReadOptions::default())?;
# Ok::<(), miniexcel::Error>(())
```

Serde `rename`, `alias`, `default`, `skip`, and `Option` semantics are supported. MiniExcel-specific column-index attributes are not part of the MVP.

## Dynamic Writing

```rust
use miniexcel::{CellValue, DynamicRow, WriteOptions, XlsxWriter};

let mut row = DynamicRow::new();
row.insert("Name".to_owned(), CellValue::String("MiniExcel".to_owned()));
row.insert("Version".to_owned(), CellValue::Int(2));

let mut writer = XlsxWriter::new();
writer.add_rows(&[row], &WriteOptions::new().with_sheet_name("Data"))?;
writer.save("book.xlsx")?;
# Ok::<(), miniexcel::Error>(())
```

Dynamic schemas are the union of row keys in first-seen order. Missing values are written as blank cells. Use `add_rows_with_schema()` when an explicit schema is required, including header-only exports.

## Typed Writing

```rust
use chrono::NaiveDate;
use miniexcel::{WriteOptions, XlsxWriter};
use serde::Serialize;

#[derive(Serialize)]
#[serde(rename_all = "PascalCase")]
struct Release {
    name: String,
    #[serde(serialize_with = "miniexcel::serde_helpers::serialize_datetime_to_excel")]
    released_on: NaiveDate,
}

let values = [Release {
    name: "MiniExcel Rust".to_owned(),
    released_on: NaiveDate::from_ymd_opt(2026, 8, 13).unwrap(),
}];
let options = WriteOptions::new()
    .with_sheet_name("Releases")
    .with_column_format("ReleasedOn", "yyyy-mm-dd");

let mut writer = XlsxWriter::new();
writer.add_serialized(&values, &options)?;
let bytes = writer.to_bytes()?;
# Ok::<(), miniexcel::Error>(())
```

The column-format key is the final Serde field/header name. Typed Serde writing supports structs and vectors of structs; maps and `flatten` are handled through the dynamic API instead.

## Important Semantics

- The default worksheet is the first workbook worksheet, not the active tab.
- Dynamic XLSX numbers with an exact `i64` representation are returned as `CellValue::Int`; other numeric values remain `Float`.
- Excel serial dates cannot always distinguish date-only, time-only, and datetime intent. Dynamic serial values are normalized to `CellValue::DateTime`; ISO values retain the more specific variant when possible.
- Formula expressions are not returned. Reading uses their cached values.
- `MiniExcel::query()` and `query_as()` strictly stream worksheet XML from paths. `XlsxReader` remains a calamine-backed, materialized compatibility API for arbitrary `Read + Seek` inputs.
- Streaming is synchronous and uses one worker thread per active query. Async I/O is not part of the MVP.
- Writing creates new workbooks and overwrites target paths. It cannot modify an existing workbook.

## Non-Goals For This MVP

CSV, `.xls`, `.xlsb`, `.ods`, templates, macros, images, merged-cell operations, arbitrary range end coordinates, formula authoring, a general style system, WASM, and editing existing workbooks are deferred.

See [Compatibility and research notes](docs/compatibility.md) for dependency choices and behavior mapping.