# Rust XLSX Compatibility Notes

## Goal

The Rust MVP implements the smallest useful MiniExcel-style XLSX read/write surface behind one `MiniExcel` facade. It uses a focused OOXML pull parser for bounded-memory path queries, calamine data and Serde conversion internally, and rust_xlsxwriter for workbook generation.

## Dependency Baseline

| Dependency | Locked API line | Role | License | MSRV note |
| --- | --- | --- | --- | --- |
| `calamine` | 0.35 | XLSX parsing and Serde row deserialization | MIT | 0.35 declares Rust 1.83 |
| `clap` | 4.6 | Local CLI argument parsing | MIT OR Apache-2.0 | 4.6 declares Rust 1.85 |
| `rust_xlsxwriter` | 0.96 | New XLSX workbook generation and Serde serialization | MIT OR Apache-2.0 | 0.96 declares Rust 1.83 |
| `serde` | 1.x | Typed mapping | MIT OR Apache-2.0 | Resolved by the workspace lockfile |
| `chrono` | 0.4 | Timezone-free Excel date/time values | MIT OR Apache-2.0 | Resolved by the workspace lockfile |
| `indexmap` | 2.x | Stable dynamic column ordering | MIT OR Apache-2.0 | Resolved by the workspace lockfile |
| `quick-xml` | 0.39 | Incremental OOXML parsing | MIT | Locked and checked with Rust 1.85 |
| `serde_json` | 1.x | Shared parity contracts and CLI JSON output | MIT OR Apache-2.0 | Checked with Rust 1.85 |
| `thiserror` | 2.x | Public error composition | MIT OR Apache-2.0 | Resolved by the workspace lockfile |
| `zip` | 7.2 | Incremental worksheet entry decompression | MIT | Locked and checked with Rust 1.85 |

The latest `calamine 0.36` and `rust_xlsxwriter 0.97` require Rust 1.88. The MVP pins the preceding API lines so the declared Rust 1.85 MSRV is executable rather than aspirational.

## API Mapping

| MiniExcel V2 concept | Rust MVP | Notes |
| --- | --- | --- |
| OpenXML importer | `MiniExcel` | Concrete reader/parser types are internal |
| Dynamic `Query` | `MiniExcel::query()` | Streams owned `IndexMap<String, CellValue>` rows with bounded buffering |
| Typed `Query<T>` | `MiniExcel::query_as<T>()` | Streams rows and applies Serde mapping one row at a time |
| `QueryRange` | `ReadOptions::with_start_cell()` / `with_end_cell()` | Inclusive A1 range for dynamic and typed reads |
| `GetSheetNames` | `MiniExcel::get_sheet_names()` | Workbook order is preserved |
| `GetColumns` | `MiniExcel::get_columns()` | Returns selected dynamic keys or an empty vector |
| `startCell` | `ReadOptions::with_start_cell()` | A1 start coordinate |
| `IgnoreEmptyRows` | `ReadOptions::with_ignore_empty_rows()` | Defaults to `false` for MiniExcel compatibility |
| OpenXML exporter | `MiniExcel::save_as*()` | Concrete writer type is internal; creates new workbooks only |
| Dynamic export | `save_as()` / `save_as_with_schema()` | Map serialization is implemented internally |
| Typed export | `save_as_serialized<T>()` | Uses Serde mapping internally |

`MiniExcel` is the only public behavior entry point. Reader, writer, parser, and concrete iterator types are crate-internal. Public supporting types are limited to row/cell values, options, errors/results, and Serde date/time helpers.

## Compatibility Defaults

- `MiniExcel::query()` with `HeaderMode::Auto` uses column letters and treats the first row as data.
- `MiniExcel::query_as()` with `HeaderMode::Auto` consumes the first selected row as headers.
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

For typed writing, chrono values must use the matching MiniExcel helper (`serialize_date_to_excel`, `serialize_datetime_to_excel`, or `serialize_time_to_excel`) and a corresponding `WriteOptions::with_column_format()` entry. Otherwise standard chrono Serde behavior writes text rather than an Excel serial value.

## Memory And I/O Model

`MiniExcel::query()` and `query_as()` use a dedicated path-streaming backend. A worker owns the ZIP archive, reads workbook relationships, styles, and shared strings, then processes worksheet XML with quick-xml. A bounded channel holds at most eight parsed rows. Dropping the public iterator disconnects the channel and joins the worker, so an early `take` or `find` stops further work.

The backend makes two sequential, bounded-memory passes over the selected worksheet entry. The first records only the maximum used column and final explicitly declared row. This is required for MiniExcel-compatible stable dynamic schemas when legal files omit `<dimension>`, and to preserve style-only row elements like the .NET reader. The second pass emits rows. Worksheet XML and prior rows are never retained; memory consists primarily of shared strings, styles, parser buffers, the current row, and the bounded channel.

The internal writer assembles a new ZIP package. The public facade writes to paths and cannot patch or insert sheets into an existing workbook.

## Test Sources

Rust integration tests reuse the repository's existing files under `tests/data/xlsx`, including:

- Dynamic header and no-header files.
- Center and self-closing empty rows.
- Typed value and trimmed-header mapping.
- Multiple worksheets.
- Cells without explicit `r` attributes.
- A typed conversion failure with a verified Excel row number.
- Strict streaming A1 starts, empty-row filtering, dates, trimmed headers, and early typed errors.

Writer tests generate temporary workbooks through `MiniExcel::save_as*()` and read them back through `MiniExcel::query*()`, covering dynamic and typed values, dates, empty schemas, path overwrite behavior, and worksheet-name validation. The WASM adapter has native unit tests, while Browser Lab Playwright tests cover generated-workbook rendering, query controls, inclusive end ranges, and desktop/mobile viewports.

## .NET Parity Contract

Behavior shared by .NET and Rust is defined in `tests/data/contracts/xlsx-parity-v1.json`. This file is the single expected-data source for:

- `tests/MiniExcel.OpenXml.Tests/Compatibility/RustParityContractTests.cs`
- `rust/miniexcel/tests/parity_contract.rs`

Both adapters use their public APIs, query the same XLSX fixtures, normalize language-specific representations, and compare sheet order, row counts, column order, selected values, and common conversion-error context. Normalization maps null/empty cells, booleans, numbers, GUIDs, datetimes, durations, and strings to stable tagged text. In particular, integral .NET `double` and Rust `CellValue::Int` values compare as the same number, and ISO date strings compare with chrono date/time values.

Run both sides from the repository root:

```bash
cargo +1.85.0 test --manifest-path rust/Cargo.toml -p miniexcel --test parity_contract --locked
dotnet test tests/MiniExcel.OpenXml.Tests/MiniExcel.OpenXml.Tests.csproj --framework net10.0 --filter "FullyQualifiedName~RustParityContractTests"
```

The Rust workflow runs the Rust contract on Linux and Windows and runs the .NET contract on Linux. The regular .NET workflow also discovers the parity tests. A compatibility change is complete only when the shared contract is updated deliberately and both adapters pass it.

The contract covers only the current common surface: dynamic/typed path queries, inclusive range queries, column-name discovery, header behavior, sheet selection/order, A1 starts, empty/style-only rows, inferred cell references, scalar/date/duration mapping, trimmed typed headers, and conversion-error row/value context. Async APIs, DataReader, templates, and writing parity remain outside version 1 and must not be described as equivalent yet.

## .NET Coverage Boundary

| .NET surface | Rust status | Shared contract |
| --- | --- | --- |
| Dynamic and typed XLSX query | Implemented | Yes |
| `QueryRange` with A1 coordinates | Implemented | Yes |
| `GetSheetNames` and `GetColumns` | Implemented | Yes |
| New-workbook `SaveAs` | Implemented and roundtrip-tested | Not yet |
| Byte-array query/write for WASM | Implemented | Rust/browser tests |
| Async APIs, DataReader, stream ownership | Deferred | No |
| Sheet information/dimensions, insert/edit | Deferred | No |
| CSV and legacy formats | Deferred | No |
| Templates, pictures, merges, comments | Deferred | No |

This matrix is the coverage claim: Rust does not yet provide complete API parity with the current .NET packages.

## Deferred Work

CSV providers, old Excel formats, templates, images, merged-cell APIs, formula authoring, general styling, modifying existing workbooks, async I/O, streaming from caller-owned readers, and publication policy require separate design and acceptance milestones.