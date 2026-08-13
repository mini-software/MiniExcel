mod common;

use chrono::NaiveDate;
use miniexcel::{Error, MiniExcel, ReadOptions, XlsxReader};
use serde::Deserialize;

#[derive(Debug, Deserialize)]
#[serde(rename_all = "PascalCase")]
struct UserAccount {
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
struct SimpleAccount {
    name: Option<String>,
    age: u32,
    mail: Option<String>,
    points: f64,
}

#[test]
fn deserializes_typed_rows_with_dates() {
    let mut reader =
        XlsxReader::open(common::fixture("TestTypeMapping.xlsx")).expect("open fixture");

    let rows: Vec<UserAccount> =
        reader.deserialize(&ReadOptions::default()).expect("deserialize rows");

    assert_eq!(rows.len(), 100);
    assert_eq!(rows[0].id, "78DE23D2-DCB6-BD3D-EC67-C112BBC322A2");
    assert_eq!(rows[0].name, "Wade");
    assert_eq!(rows[0].born_on, NaiveDate::from_ymd_opt(2020, 9, 27).unwrap());
    assert_eq!(rows[0].age, 36);
    assert!(!rows[0].vip);
    assert_eq!(rows[0].points, 5019.12);
}

#[test]
fn trims_headers_for_typed_rows_by_default() {
    let mut reader =
        XlsxReader::open(common::fixture("TestTrimColumnNames.xlsx")).expect("open fixture");

    let rows: Vec<SimpleAccount> =
        reader.deserialize(&ReadOptions::default()).expect("deserialize rows");

    assert_eq!(rows[4].name.as_deref(), Some("Raymond"));
    assert_eq!(rows[4].age, 18);
    assert_eq!(rows[4].mail.as_deref(), Some("sagittis.lobortis@leoMorbi.com"));
    assert_eq!(rows[4].points, 8209.76);
}

#[test]
fn reports_missing_sheets() {
    let mut reader =
        XlsxReader::open(common::fixture("TestMultiSheet.xlsx")).expect("open fixture");
    let error = reader
        .read_rows(&ReadOptions::new().with_sheet_name("missing"))
        .expect_err("missing sheet should fail");

    assert!(error.to_string().contains("missing"));
}

#[derive(Debug, Deserialize)]
#[serde(rename_all = "PascalCase")]
#[allow(dead_code)]
struct InvalidSequence {
    #[serde(rename = "ID")]
    id: u32,
    name: Option<String>,
    #[serde(rename = "SEQ")]
    sequence: u32,
}

#[test]
fn reports_sheet_and_excel_row_for_mapping_errors() {
    let mut reader = XlsxReader::open(common::fixture("TestIssue309.xlsx")).expect("open fixture");
    let error = reader
        .deserialize::<InvalidSequence>(&ReadOptions::default())
        .expect_err("invalid sequence should fail");

    match error {
        Error::Deserialize { sheet, row, source } => {
            assert_eq!(sheet, "Sheet1");
            assert_eq!(row, 4);
            assert!(source.to_string().contains("SEQ") || source.to_string().contains("Error"));
        }
        other => panic!("unexpected error: {other}"),
    }
}

#[test]
fn typed_query_maps_rows_lazily() {
    let mut rows = MiniExcel::query_as::<InvalidSequence>(common::fixture("TestIssue309.xlsx"))
        .expect("create typed query");

    assert!(rows.next().expect("Excel row 2").is_ok());
    assert!(rows.next().expect("Excel row 3").is_ok());
    let error = rows.next().expect("Excel row 4").expect_err("row 4 should fail");

    match error {
        Error::Deserialize { sheet, row, .. } => {
            assert_eq!(sheet, "Sheet1");
            assert_eq!(row, 4);
        }
        other => panic!("unexpected error: {other}"),
    }
}

#[test]
fn typed_streaming_query_deserializes_dates() {
    let mut rows = MiniExcel::query_as::<UserAccount>(common::fixture("TestTypeMapping.xlsx"))
        .expect("create typed streaming query");

    let first = rows.next().expect("first row").expect("deserialize first row");
    assert_eq!(first.id, "78DE23D2-DCB6-BD3D-EC67-C112BBC322A2");
    assert_eq!(first.name, "Wade");
    assert_eq!(first.born_on, NaiveDate::from_ymd_opt(2020, 9, 27).unwrap());
    assert_eq!(first.age, 36);
    assert!(!first.vip);
    assert_eq!(first.points, 5019.12);
}

#[test]
fn typed_streaming_query_trims_headers() {
    let rows = MiniExcel::query_as::<SimpleAccount>(common::fixture("TestTrimColumnNames.xlsx"))
        .expect("create trimmed-header query")
        .collect::<miniexcel::Result<Vec<_>>>()
        .expect("deserialize trimmed-header rows");

    assert_eq!(rows[4].name.as_deref(), Some("Raymond"));
    assert_eq!(rows[4].age, 18);
    assert_eq!(rows[4].mail.as_deref(), Some("sagittis.lobortis@leoMorbi.com"));
    assert_eq!(rows[4].points, 8209.76);
}
