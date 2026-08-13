# Rust XLSX Compatibility Notes

## Goal

The Rust MVP implements the smallest useful MiniExcel-style XLSX read/write surface while keeping an idiomatic Rust API. It uses a focused OOXML pull parser for bounded-memory path queries, calamine for the general `Read + Seek` compatibility reader and Serde conversion, and rust_xlsxwriter for workbook generation.

## Dependency Baseline

| Dependency | Locked API line | Role | License | MSRV note |
| --- | --- | --- | --- | --- |
| `calamine` | 0.35 | XLSX parsing and Serde row deserialization | MIT | 0.35 declares Rust 1.83 |
| `rust_xlsxwriter` | 0.96 | New XLSX workbook generation and Serde serialization | MIT OR Apache-2.0 | 0.96 declares Rust 1.83 |
| `serde` | 1.x | Typed mapping | MIT OR Apache-2.0 | Resolved by the workspace lockfile |
| `chrono` | 0.4 | Timezone-free Excel date/time values | MIT OR Apache-2.0 | Resolved by the workspace lockfile |
| `indexmap` | 2.x | Stable dynamic column ordering | MIT OR Apache-2.0 | Resolved by the workspace lockfile |
| `quick-xml` | 0.39 | Incremental OOXML parsing | MIT | Locked and checked with Rust 1.85 |
| `thiserror` | 2.x | Public error composition | MIT OR Apache-2.0 | Resolved by the workspace lockfile |
| `zip` | 7.2 | Incremental worksheet entry decompression | MIT | Locked and checked with Rust 1.85 |

The latest `calamine 0.36` and `rust_xlsxwriter 0.97` require Rust 1.88. The MVP pins the preceding API lines so the declared Rust 1.85 MSRV is executable rather than aspirational.

## API Mapping

| MiniExcel V2 concept | Rust MVP | Notes |
| --- | --- | --- |
| OpenXML importer | `MiniExcel` / `XlsxReader<R>` | Strict path streaming or materialized `Read + Seek` compatibility mode |
| Dynamic `Query` | `MiniExcel::query()` | Streams owned `IndexMap<String, CellValue>` rows with bounded buffering |
| Typed `Query<T>` | `MiniExcel::query_as<T>()` | Streams rows and applies Serde mapping one row at a time |
| General reader query | `XlsxReader::query()` / `query_as<T>()` | Lazy mapping over a calamine-materialized worksheet `Range` |
| `GetSheetNames` | `sheet_names()` | Workbook order is preserved |
| `startCell` | `ReadOptions::with_start_cell()` | A1 start only; no end coordinate in M1 |
| `IgnoreEmptyRows` | `ReadOptions::with_ignore_empty_rows()` | Defaults to `false` for MiniExcel compatibility |
| OpenXML exporter | `XlsxWriter` | Creates new workbooks only |
| Dynamic export | `add_rows()` / `add_rows_with_schema()` | Map serialization is implemented manually |
| Typed export | `add_serialized<T>()` | Uses `rust_xlsxwriter` Serde support |
| Path/stream export | `save()`, `to_bytes()`, `save_to_writer()` | Writer output does not require `Seek` |

`MiniExcel` provides the simple static path facade familiar to .NET users. `XlsxReader` and `XlsxWriter` remain available when callers need explicit ownership and state. Options use builder methods and all failures return `Result`.

## Compatibility Defaults

- `read_rows()` with `HeaderMode::Auto` uses column letters and treats the first row as data.
- `deserialize()` with `HeaderMode::Auto` consumes the first selected row as headers.
- The first worksheet in workbook order is selected when no name is supplied.
- Empty rows between the selected start and last used cell are retained by default.
- Typed header strings are trimmed by default. Dynamic headers follow the .NET behavior and retain non-blank text as stored.
- Blank dynamic headers are omitted. Duplicate dynamic headers retain their first key position while later columns overwrite the value.
- A missing dynamic cell is represented by `CellValue::Empty`, not by omission from a known schema.
- Writer row counts exclude the header row.

## Type Mapping

| XLSX value | Dynamic Rust value |
| --- | --- |
| Empty | `CellValue::Empty` |
| Boolean | `CellValue::Bool` |
| Exact integral number in `i64` range | `CellValue::Int` |
| Other number | `CellValue::Float` |
| Shared/inline string | `CellValue::String` |
| Excel serial date/time | `CellValue::DateTime` |
| Excel duration | `CellValue::Duration` |
| ISO date/time | `Date`, `Time`, or `DateTime` when parseable |
| Cell error | `CellValue::Error` |
| Formula | Cached result value only |

Typed conversions are delegated to calamine's Serde deserializer. The public `serde_helpers` module adds strict chrono helpers that convert an invalid value into the library's contextual `Error::Deserialize` path.

For typed writing, chrono values must use `serialize_datetime_to_excel` (or its optional variant) and a corresponding `WriteOptions::with_column_format()` entry. Otherwise standard chrono Serde behavior writes text rather than an Excel serial date.

## Memory And I/O Model

`MiniExcel::query()` and `query_as()` use a dedicated path-streaming backend. A worker owns the ZIP archive, reads workbook relationships, styles, and shared strings, then processes worksheet XML with quick-xml. A bounded channel holds at most eight parsed rows. Dropping the public iterator disconnects the channel and joins the worker, so an early `take` or `find` stops further work.

The backend makes two sequential, bounded-memory passes over the selected worksheet entry. The first records only the maximum used column and final row containing a cell. This is required for MiniExcel-compatible stable dynamic schemas when legal files omit `<dimension>`, and to avoid exposing trailing style-only row elements. The second pass emits rows. Worksheet XML and prior rows are never retained; memory consists primarily of shared strings, styles, parser buffers, the current row, and the bounded channel.

`XlsxReader<R>` is the compatibility path for arbitrary `Read + Seek` inputs and sheet-name inspection. Its `query()` / `query_as()` map rows lazily, but calamine materializes the selected worksheet first. `read_rows()` and `deserialize()` collect those iterators.

`XlsxWriter` can emit to a non-seekable `Write + Send` target, but rust_xlsxwriter assembles a new ZIP package. It cannot patch or insert sheets into an existing workbook.

## Test Sources

Rust integration tests reuse the repository's existing files under `tests/data/xlsx`, including:

- Dynamic header and no-header files.
- Center and self-closing empty rows.
- Typed value and trimmed-header mapping.
- Multiple worksheets.
- Cells without explicit `r` attributes.
- A typed conversion failure with a verified Excel row number.
- Strict streaming A1 starts, empty-row filtering, dates, trimmed headers, and early typed errors.

Writer tests generate temporary workbooks and read them back through `XlsxReader`, covering dynamic and typed values, dates, multiple output targets, empty schemas, path overwrite behavior, and worksheet-name validation.

## Deferred Work

CSV providers, old Excel formats, templates, images, merged-cell APIs, formula authoring, general styling, modifying existing workbooks, async I/O, streaming from caller-owned readers, WASM, and publication policy require separate design and acceptance milestones.