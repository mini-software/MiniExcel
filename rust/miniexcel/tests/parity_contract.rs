mod common;

use std::collections::HashMap;
use std::fs;
use std::path::PathBuf;

use chrono::{NaiveDate, NaiveDateTime, NaiveTime, Timelike};
use miniexcel::{CellReference, CellValue, HeaderMode, MiniExcel, ReadOptions};
use serde::Deserialize;

#[derive(Debug, Deserialize)]
#[serde(rename_all = "camelCase")]
struct ParityContract {
    version: u32,
    dynamic_cases: Vec<DynamicCase>,
    typed_cases: Vec<TypedCase>,
    error_cases: Vec<ErrorCase>,
}

#[derive(Debug, Deserialize)]
#[serde(rename_all = "camelCase")]
struct DynamicCase {
    name: String,
    fixture: String,
    has_header: bool,
    sheet_name: Option<String>,
    start_cell: String,
    end_cell: Option<String>,
    ignore_empty_rows: bool,
    expected_sheet_names: Option<Vec<String>>,
    row_count: usize,
    expected_columns: Option<Vec<String>>,
    samples: Vec<RowSample>,
}

#[derive(Debug, Deserialize)]
#[serde(rename_all = "camelCase")]
struct TypedCase {
    name: String,
    model: String,
    fixture: String,
    row_count: usize,
    samples: Vec<RowSample>,
}

#[derive(Debug, Deserialize)]
#[serde(rename_all = "camelCase")]
struct ErrorCase {
    name: String,
    model: String,
    fixture: String,
    expected_row: usize,
    expected_value: String,
}

#[derive(Debug, Deserialize)]
#[serde(rename_all = "camelCase")]
struct RowSample {
    row_index: usize,
    cells: HashMap<String, String>,
}

#[derive(Debug, Deserialize)]
#[serde(rename_all = "PascalCase")]
struct ParityUserAccount {
    #[serde(rename = "ID")]
    id: String,
    name: String,
    #[serde(rename = "BoD", deserialize_with = "miniexcel::serde_helpers::deserialize_date")]
    born_on: NaiveDate,
    age: u32,
    #[serde(rename = "VIP")]
    vip: bool,
    points: f64,
}

#[derive(Debug, Deserialize)]
#[serde(rename_all = "PascalCase")]
struct ParitySimpleAccount {
    name: Option<String>,
    age: u32,
    mail: Option<String>,
    points: f64,
}

#[derive(Debug, Deserialize)]
#[serde(rename_all = "PascalCase")]
#[allow(dead_code)]
struct ParityInvalidSequence {
    #[serde(rename = "ID")]
    id: u32,
    name: Option<String>,
    #[serde(rename = "SEQ")]
    sequence: u32,
}

#[test]
fn dynamic_queries_match_dotnet_contract() {
    let contract = load_contract();
    assert_eq!(contract.version, 1, "unsupported parity contract version");

    for test_case in &contract.dynamic_cases {
        let path = common::fixture(&test_case.fixture);
        if let Some(expected) = &test_case.expected_sheet_names {
            let actual = MiniExcel::get_sheet_names(&path)
                .unwrap_or_else(|error| panic!("{}: read sheet names: {error}", test_case.name));
            assert_eq!(&actual, expected, "{}: sheet names", test_case.name);
        }

        let start_cell: CellReference = test_case
            .start_cell
            .parse()
            .unwrap_or_else(|error| panic!("{}: invalid start cell: {error}", test_case.name));
        let mut options = ReadOptions::new()
            .with_start_cell(start_cell)
            .with_header_mode(if test_case.has_header {
                HeaderMode::FirstRow
            } else {
                HeaderMode::None
            })
            .with_ignore_empty_rows(test_case.ignore_empty_rows);
        if let Some(sheet_name) = &test_case.sheet_name {
            options = options.with_sheet_name(sheet_name);
        }
        if let Some(end_cell) = &test_case.end_cell {
            options =
                options.with_end_cell(end_cell.parse().unwrap_or_else(|error| {
                    panic!("{}: invalid end cell: {error}", test_case.name)
                }));
        }

        let rows = MiniExcel::query_with_options(path, &options)
            .unwrap_or_else(|error| panic!("{}: create query: {error}", test_case.name))
            .collect::<miniexcel::Result<Vec<_>>>()
            .unwrap_or_else(|error| panic!("{}: read rows: {error}", test_case.name));
        assert_eq!(rows.len(), test_case.row_count, "{}: row count", test_case.name);

        if let Some(expected) = &test_case.expected_columns {
            let actual = rows
                .first()
                .unwrap_or_else(|| panic!("{}: expected at least one row", test_case.name))
                .keys()
                .cloned()
                .collect::<Vec<_>>();
            assert_eq!(&actual, expected, "{}: column order", test_case.name);
            if test_case.end_cell.is_none() {
                let actual = MiniExcel::get_columns(common::fixture(&test_case.fixture), &options)
                    .unwrap_or_else(|error| {
                        panic!("{}: get column names: {error}", test_case.name)
                    });
                assert_eq!(&actual, expected, "{}: get column names", test_case.name);
            }
        }

        for sample in &test_case.samples {
            let row = rows.get(sample.row_index).unwrap_or_else(|| {
                panic!("{}: missing sample row {}", test_case.name, sample.row_index)
            });
            for (column, expected) in &sample.cells {
                let actual = row.get(column).unwrap_or_else(|| {
                    panic!(
                        "{}: row {} does not contain column {column}",
                        test_case.name, sample.row_index
                    )
                });
                assert_eq!(
                    normalize_cell(actual),
                    *expected,
                    "{}: row {}, column {column}",
                    test_case.name,
                    sample.row_index
                );
            }
        }
    }
}

#[test]
fn typed_queries_match_dotnet_contract() {
    let contract = load_contract();

    for test_case in &contract.typed_cases {
        let rows = normalized_typed_rows(test_case);
        assert_eq!(rows.len(), test_case.row_count, "{}: row count", test_case.name);

        for sample in &test_case.samples {
            let row = rows.get(sample.row_index).unwrap_or_else(|| {
                panic!("{}: missing sample row {}", test_case.name, sample.row_index)
            });
            for (field, expected) in &sample.cells {
                let actual = row.get(field).unwrap_or_else(|| {
                    panic!(
                        "{}: row {} does not contain field {field}",
                        test_case.name, sample.row_index
                    )
                });
                assert_eq!(
                    actual, expected,
                    "{}: row {}, field {field}",
                    test_case.name, sample.row_index
                );
            }
        }
    }
}

#[test]
fn conversion_errors_match_dotnet_contract() {
    let contract = load_contract();

    for test_case in &contract.error_cases {
        let path = common::fixture(&test_case.fixture);
        let error = match test_case.model.as_str() {
            "invalidSequence" => MiniExcel::query_as::<ParityInvalidSequence>(path)
                .unwrap_or_else(|error| panic!("{}: create typed query: {error}", test_case.name))
                .collect::<miniexcel::Result<Vec<_>>>()
                .expect_err("typed conversion should fail"),
            model => panic!("{}: unsupported error model '{model}'", test_case.name),
        };
        let message = error.to_string();
        assert!(
            message.contains(&format!("row {}", test_case.expected_row)),
            "{}: expected row {} in '{message}'",
            test_case.name,
            test_case.expected_row
        );
        assert!(
            message.contains(&test_case.expected_value),
            "{}: expected value '{}' in '{message}'",
            test_case.name,
            test_case.expected_value
        );
    }
}

fn load_contract() -> ParityContract {
    let path = PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("..")
        .join("..")
        .join("tests")
        .join("data")
        .join("contracts")
        .join("xlsx-parity-v1.json");
    let json = fs::read_to_string(path).expect("read shared parity contract");
    serde_json::from_str(&json).expect("parse shared parity contract")
}

fn normalized_typed_rows(test_case: &TypedCase) -> Vec<HashMap<String, String>> {
    let path = common::fixture(&test_case.fixture);
    match test_case.model.as_str() {
        "userAccount" => MiniExcel::query_as::<ParityUserAccount>(path)
            .unwrap_or_else(|error| panic!("{}: create typed query: {error}", test_case.name))
            .collect::<miniexcel::Result<Vec<_>>>()
            .unwrap_or_else(|error| panic!("{}: read typed rows: {error}", test_case.name))
            .into_iter()
            .map(|row| {
                HashMap::from([
                    ("ID".to_owned(), normalize_guid(&row.id)),
                    ("Name".to_owned(), normalize_string(&row.name)),
                    ("BoD".to_owned(), normalize_datetime(row.born_on.and_time(NaiveTime::MIN))),
                    ("Age".to_owned(), format!("number:{}", row.age)),
                    ("VIP".to_owned(), format!("bool:{}", row.vip)),
                    ("Points".to_owned(), format!("number:{}", row.points)),
                ])
            })
            .collect(),
        "simpleAccount" => MiniExcel::query_as::<ParitySimpleAccount>(path)
            .unwrap_or_else(|error| panic!("{}: create typed query: {error}", test_case.name))
            .collect::<miniexcel::Result<Vec<_>>>()
            .unwrap_or_else(|error| panic!("{}: read typed rows: {error}", test_case.name))
            .into_iter()
            .map(|row| {
                HashMap::from([
                    (
                        "Name".to_owned(),
                        row.name.as_deref().map_or_else(|| "empty:".to_owned(), normalize_string),
                    ),
                    ("Age".to_owned(), format!("number:{}", row.age)),
                    (
                        "Mail".to_owned(),
                        row.mail.as_deref().map_or_else(|| "empty:".to_owned(), normalize_string),
                    ),
                    ("Points".to_owned(), format!("number:{}", row.points)),
                ])
            })
            .collect(),
        model => panic!("{}: unsupported typed model '{model}'", test_case.name),
    }
}

fn normalize_cell(value: &CellValue) -> String {
    match value {
        CellValue::Empty => "empty:".to_owned(),
        CellValue::Bool(value) => format!("bool:{value}"),
        CellValue::Int(value) => format!("number:{value}"),
        CellValue::Float(value) => format!("number:{value}"),
        CellValue::String(value) => normalize_string(value),
        CellValue::Date(value) => normalize_datetime(value.and_time(NaiveTime::MIN)),
        CellValue::Time(value) => format!("time:{}", value.format("%H:%M:%S")),
        CellValue::DateTime(value) => normalize_datetime(*value),
        CellValue::Duration(value) => format!("duration:{}", value.num_milliseconds()),
        CellValue::Error(value) => format!("error:{value}"),
    }
}

fn normalize_string(value: &str) -> String {
    if let Some(guid) = canonical_guid(value) {
        return format!("guid:{guid}");
    }
    if let Ok(value) = NaiveDateTime::parse_from_str(value, "%Y-%m-%dT%H:%M:%S") {
        return normalize_datetime(value);
    }
    if let Ok(value) = NaiveDateTime::parse_from_str(value, "%Y-%m-%d %H:%M:%S") {
        return normalize_datetime(value);
    }
    if let Ok(value) = NaiveDate::parse_from_str(value, "%Y-%m-%d") {
        return normalize_datetime(value.and_time(NaiveTime::MIN));
    }
    format!("string:{value}")
}

fn normalize_guid(value: &str) -> String {
    canonical_guid(value)
        .map(|value| format!("guid:{value}"))
        .unwrap_or_else(|| normalize_string(value))
}

fn canonical_guid(value: &str) -> Option<String> {
    let bytes = value.as_bytes();
    if bytes.len() != 36
        || !bytes.iter().enumerate().all(|(index, byte)| {
            matches!(index, 8 | 13 | 18 | 23)
                .then_some(*byte == b'-')
                .unwrap_or_else(|| byte.is_ascii_hexdigit())
        })
    {
        return None;
    }
    Some(value.to_ascii_uppercase())
}

fn normalize_datetime(value: NaiveDateTime) -> String {
    if value.nanosecond() == 0 {
        format!("datetime:{}", value.format("%Y-%m-%dT%H:%M:%S"))
    } else {
        format!("datetime:{}", value.format("%Y-%m-%dT%H:%M:%S%.f"))
    }
}
