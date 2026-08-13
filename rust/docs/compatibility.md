# Rust XLSX Compatibility Notes

## Goal

The Rust MVP implements the smallest useful MiniExcel-style XLSX read/write surface while keeping an idiomatic Rust API. It wraps mature format libraries instead of maintaining a second OOXML engine in this repository.

## Dependency Baseline

| Dependency | Locked API line | Role | License | MSRV note |
| --- | --- | --- | --- | --- |
| `calamine` | 0.35 | XLSX parsing and Serde row deserialization | MIT | 0.35 declares Rust 1.83 |
| `rust_xlsxwriter` | 0.96 | New XLSX workbook generation and Serde serialization | MIT OR Apache-2.0 | 0.96 declares Rust 1.83 |
| `serde` | 1.x | Typed mapping | MIT OR Apache-2.0 | Resolved by the workspace lockfile |
| `chrono` | 0.4 | Timezone-free Excel date/time values | MIT OR Apache-2.0 | Resolved by the workspace lockfile |
| `indexmap` | 2.x | Stable dynamic column ordering | MIT OR Apache-2.0 | Resolved by the workspace lockfile |
| `thiserror` | 2.x | Public error composition | MIT OR Apache-2.0 | Resolved by the workspace lockfile |

The latest `calamine 0.36` and `rust_xlsxwriter 0.97` require Rust 1.88. The MVP pins the preceding API lines so the declared Rust 1.85 MSRV is executable rather than aspirational.

## API Mapping

| MiniExcel V2 concept | Rust MVP | Notes |
| --- | --- | --- |
| OpenXML importer | `XlsxReader<R>` | Generic over `Read + Seek` |
| Dynamic `Query` | `read_rows()` | Returns owned `IndexMap<String, CellValue>` rows |
| Typed `Query<T>` | `deserialize<T>()` | Uses Serde field naming and helpers |
| `GetSheetNames` | `sheet_names()` | Workbook order is preserved |
| `startCell` | `ReadOptions::with_start_cell()` | A1 start only; no end coordinate in M1 |
| `IgnoreEmptyRows` | `ReadOptions::with_ignore_empty_rows()` | Defaults to `false` for MiniExcel compatibility |
| OpenXML exporter | `XlsxWriter` | Creates new workbooks only |
| Dynamic export | `add_rows()` / `add_rows_with_schema()` | Map serialization is implemented manually |
| Typed export | `add_serialized<T>()` | Uses `rust_xlsxwriter` Serde support |
| Path/stream export | `save()`, `to_bytes()`, `save_to_writer()` | Writer output does not require `Seek` |

The Rust API does not reproduce the static provider facade. Reader and writer values carry ownership and state directly, while options use builder methods and all failures return `Result`.

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

`XlsxReader` accepts a stream-like source but calamine materializes worksheet ranges. Returning owned rows also allocates mapped values. This differs from MiniExcel's .NET `IAsyncEnumerable` streaming implementation and is documented as an MVP limitation.

`XlsxWriter` can emit to a non-seekable `Write + Send` target, but rust_xlsxwriter assembles a new ZIP package. It cannot patch or insert sheets into an existing workbook.

## Test Sources

Rust integration tests reuse the repository's existing files under `tests/data/xlsx`, including:

- Dynamic header and no-header files.
- Center and self-closing empty rows.
- Typed value and trimmed-header mapping.
- Multiple worksheets.
- Cells without explicit `r` attributes.
- A typed conversion failure with a verified Excel row number.

Writer tests generate temporary workbooks and read them back through `XlsxReader`, covering dynamic and typed values, dates, multiple output targets, empty schemas, path overwrite behavior, and worksheet-name validation.

## Deferred Work

CSV providers, old Excel formats, templates, images, merged-cell APIs, formulas, general styling, modifying existing workbooks, true streaming, async I/O, WASM, and publication policy require separate design and acceptance milestones.