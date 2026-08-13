use std::io::Cursor;

use chrono::{Duration, NaiveDate, NaiveDateTime, NaiveTime};
use miniexcel::{
    CellValue, DynamicRow, HeaderMode, ReadOptions, WriteOptions, XlsxReader, XlsxWriter,
};
use serde::{Deserialize, Serialize};

#[derive(Debug, Deserialize, Serialize)]
#[serde(rename_all = "PascalCase")]
struct Release {
    name: String,
    version: u32,
    #[serde(
        serialize_with = "miniexcel::serde_helpers::serialize_datetime_to_excel",
        deserialize_with = "miniexcel::serde_helpers::deserialize_date"
    )]
    released_on: NaiveDate,
    #[serde(skip)]
    internal: bool,
}

#[test]
fn writes_dynamic_rows_and_reads_them_back() {
    let date = NaiveDate::from_ymd_opt(2025, 8, 13).unwrap();
    let time = NaiveTime::from_hms_opt(14, 30, 15).unwrap();
    let datetime = NaiveDateTime::new(date, time);

    let mut first = DynamicRow::new();
    first.insert("Name".to_owned(), CellValue::String("MiniExcel".to_owned()));
    first.insert("Count".to_owned(), CellValue::Int(2));
    first.insert("Ratio".to_owned(), CellValue::Float(1.25));
    first.insert("Active".to_owned(), CellValue::Bool(true));
    first.insert("Date".to_owned(), CellValue::Date(date));
    first.insert("Time".to_owned(), CellValue::Time(time));
    first.insert("Created".to_owned(), CellValue::DateTime(datetime));
    first.insert("Elapsed".to_owned(), CellValue::Duration(Duration::hours(27)));
    first.insert("Missing".to_owned(), CellValue::Empty);

    let mut second = DynamicRow::new();
    second.insert("Name".to_owned(), CellValue::String("Rust".to_owned()));
    second.insert("Later".to_owned(), CellValue::String("union column".to_owned()));

    let mut writer = XlsxWriter::new();
    let summary = writer
        .add_rows(&[first, second], &WriteOptions::new().with_sheet_name("Data"))
        .expect("add rows");
    assert_eq!(summary.sheet_name(), "Data");
    assert_eq!(summary.rows_written(), 2);

    let bytes = writer.to_bytes().expect("write workbook");
    let mut reader = XlsxReader::from_reader(Cursor::new(bytes)).expect("open generated workbook");
    assert_eq!(reader.sheet_names(), ["Data"]);

    let options = ReadOptions::new().with_sheet_name("Data").with_header_mode(HeaderMode::FirstRow);
    let rows = reader.read_rows(&options).expect("read generated rows");

    assert_eq!(rows.len(), 2);
    assert_eq!(rows[0]["Name"], CellValue::String("MiniExcel".to_owned()));
    assert_eq!(rows[0]["Count"], CellValue::Int(2));
    assert_eq!(rows[0]["Ratio"], CellValue::Float(1.25));
    assert_eq!(rows[0]["Active"], CellValue::Bool(true));
    assert_eq!(rows[0]["Elapsed"], CellValue::Duration(Duration::hours(27)));
    assert!(rows[0]["Missing"].is_empty());
    assert!(rows[1]["Count"].is_empty());
    assert_eq!(rows[1]["Later"], CellValue::String("union column".to_owned()));
    assert_eq!(rows[0].keys().last().map(String::as_str), Some("Later"));
}

#[test]
fn writes_without_headers() {
    let mut row = DynamicRow::new();
    row.insert("Name".to_owned(), CellValue::String("MiniExcel".to_owned()));
    row.insert("Count".to_owned(), CellValue::Int(1));

    let mut writer = XlsxWriter::new();
    writer.add_rows(&[row], &WriteOptions::new().with_print_header(false)).expect("add rows");
    let bytes = writer.to_bytes().expect("write workbook");

    let mut reader = XlsxReader::from_reader(Cursor::new(bytes)).expect("open generated workbook");
    let rows = reader.read_rows(&ReadOptions::default()).expect("read generated rows");
    assert_eq!(rows[0]["A"], CellValue::String("MiniExcel".to_owned()));
    assert_eq!(rows[0]["B"], CellValue::Int(1));
}

#[test]
fn rejects_invalid_and_duplicate_sheet_names() {
    let mut row = DynamicRow::new();
    row.insert("Value".to_owned(), CellValue::Int(1));

    let mut writer = XlsxWriter::new();
    writer
        .add_rows(&[row.clone()], &WriteOptions::new().with_sheet_name("Data"))
        .expect("add first sheet");

    assert!(writer.add_rows(&[row.clone()], &WriteOptions::new().with_sheet_name("data")).is_err());
    assert!(writer.add_rows(&[row], &WriteOptions::new().with_sheet_name("invalid/name")).is_err());
}

#[test]
fn writes_serde_structs_with_dates() {
    let releases = [
        Release {
            name: "MiniExcel".to_owned(),
            version: 2,
            released_on: NaiveDate::from_ymd_opt(2025, 1, 2).unwrap(),
            internal: true,
        },
        Release {
            name: "MiniExcel Rust".to_owned(),
            version: 1,
            released_on: NaiveDate::from_ymd_opt(2026, 8, 13).unwrap(),
            internal: false,
        },
    ];

    let options = WriteOptions::new()
        .with_sheet_name("Releases")
        .with_column_format("ReleasedOn", "yyyy-mm-dd");
    let mut writer = XlsxWriter::new();
    let summary = writer.add_serialized(&releases, &options).expect("serialize rows");
    assert_eq!(summary.rows_written(), 2);

    let mut bytes = Vec::new();
    writer.save_to_writer(&mut bytes).expect("write to generic writer");
    let mut reader = XlsxReader::from_reader(Cursor::new(bytes)).expect("open generated workbook");
    let rows: Vec<Release> = reader
        .deserialize(&ReadOptions::new().with_sheet_name("Releases"))
        .expect("deserialize generated rows");

    assert_eq!(rows.len(), 2);
    assert_eq!(rows[0].name, "MiniExcel");
    assert_eq!(rows[0].version, 2);
    assert_eq!(rows[0].released_on, NaiveDate::from_ymd_opt(2025, 1, 2).unwrap());
    assert!(!rows[0].internal);
    assert_eq!(rows[1].released_on, NaiveDate::from_ymd_opt(2026, 8, 13).unwrap());
}

#[test]
fn writes_explicit_empty_schema_and_overwrites_paths() {
    let temp_dir = tempfile::tempdir().expect("create temp directory");
    let path = temp_dir.path().join("output.xlsx");
    let schema = vec!["Value".to_owned()];

    let mut writer = XlsxWriter::new();
    writer
        .add_rows_with_schema(&schema, &[], &WriteOptions::default())
        .expect("add header-only sheet");
    writer.save(&path).expect("save first workbook");

    let mut first_reader = XlsxReader::open(&path).expect("open first workbook");
    let first_rows = first_reader
        .read_rows(&ReadOptions::new().with_header_mode(HeaderMode::FirstRow))
        .expect("read header-only sheet");
    assert!(first_rows.is_empty());

    let mut row = DynamicRow::new();
    row.insert("Value".to_owned(), CellValue::String("replacement".to_owned()));
    let mut replacement = XlsxWriter::new();
    replacement.add_rows(&[row], &WriteOptions::default()).expect("add replacement row");
    replacement.save(&path).expect("overwrite workbook");

    let mut second_reader = XlsxReader::open(path).expect("open replacement workbook");
    let rows = second_reader
        .read_rows(&ReadOptions::new().with_header_mode(HeaderMode::FirstRow))
        .expect("read replacement row");
    assert_eq!(rows[0]["Value"], CellValue::String("replacement".to_owned()));
}

#[test]
fn requires_schema_for_empty_default_exports() {
    let mut writer = XlsxWriter::new();
    assert!(writer.add_rows(&[], &WriteOptions::default()).is_err());
    assert!(writer.add_serialized::<Release>(&[], &WriteOptions::default()).is_err());
}
