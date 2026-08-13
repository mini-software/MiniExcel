mod common;

use chrono::NaiveDate;
use miniexcel::{CellReference, CellValue, HeaderMode, MiniExcel, ReadOptions, XlsxReader};

#[test]
fn queries_rows_through_the_simple_facade() {
    let mut rows = MiniExcel::query(common::fixture("TestDynamicQueryBasic_WithoutHead.xlsx"))
        .expect("create query");

    let first = rows.next().expect("first row").expect("read first row");
    assert_eq!(first["A"], CellValue::String("MiniExcel".to_owned()));
    assert_eq!(first["B"], CellValue::Int(1));
    assert_eq!(rows.count(), 1);
}

#[test]
fn streaming_query_can_stop_early() {
    let mut rows =
        MiniExcel::query(common::fixture("TestTypeMapping.xlsx")).expect("create streaming query");
    assert!(rows.next().expect("first row").is_ok());
    drop(rows);
}

#[test]
fn streaming_query_preserves_self_closing_empty_rows() {
    let rows = MiniExcel::query(common::fixture("TestEmptySelfClosingRow.xlsx"))
        .expect("create streaming query")
        .collect::<miniexcel::Result<Vec<_>>>()
        .expect("stream rows");

    assert_eq!(rows.len(), 10);
    assert!(rows[0]["A"].is_empty());
    assert_eq!(rows[1]["A"], CellValue::Int(1));
    assert!(rows[2]["A"].is_empty());
    assert_eq!(rows[3]["A"], CellValue::Int(2));
    assert!(rows[4..9].iter().all(|row| row["A"].is_empty()));
    assert_eq!(rows[9]["A"], CellValue::Int(1));
}

#[test]
fn streaming_query_selects_sheets_and_infers_missing_cell_references() {
    let options = ReadOptions::new().with_sheet_name("Sheet3");
    let rows = MiniExcel::query_with_options(common::fixture("TestMultiSheet.xlsx"), &options)
        .expect("create Sheet3 query")
        .collect::<miniexcel::Result<Vec<_>>>()
        .expect("stream Sheet3");
    assert_eq!(rows.len(), 5);
    assert_eq!(rows[0]["A"], CellValue::Int(3));
    assert_eq!(rows[0]["B"], CellValue::Int(3));

    let rows = MiniExcel::query(common::fixture("TestWihoutRAttribute.xlsx"))
        .expect("create missing-reference query")
        .collect::<miniexcel::Result<Vec<_>>>()
        .expect("stream missing-reference rows");
    assert_eq!(rows.len(), 2);
    assert_eq!(rows[0].len(), 5);
    assert_eq!(rows[1]["B"], CellValue::String("\"<>+}{\\nHello World".to_owned()));
    assert_eq!(rows[1]["C"], CellValue::Bool(true));
}

#[test]
fn streaming_query_honors_start_cells_and_empty_row_filtering() {
    let start_cell: CellReference = "B6".parse().expect("valid A1 reference");
    let options = ReadOptions::new().with_start_cell(start_cell);
    let first = MiniExcel::query_with_options(common::fixture("TestTypeMapping.xlsx"), &options)
        .expect("create start-cell query")
        .next()
        .expect("first selected row")
        .expect("stream first selected row");
    assert_eq!(first["B"], CellValue::String("Raymond".to_owned()));
    assert_eq!(first["D"], CellValue::Int(18));

    let options = ReadOptions::new().with_ignore_empty_rows(true);
    let rows = MiniExcel::query_with_options(
        common::fixture("TestCenterEmptyRow/TestCenterEmptyRow.xlsx"),
        &options,
    )
    .expect("create empty-row query")
    .collect::<miniexcel::Result<Vec<_>>>()
    .expect("stream non-empty rows");
    assert_eq!(rows.len(), 5);
    assert!(rows.iter().all(|row| row.values().any(|value| !value.is_empty())));
}

#[test]
fn streaming_query_reads_cached_formula_values() {
    let path = common::fixture("TestIssue157.xlsx");
    let streamed = MiniExcel::query(&path)
        .expect("create formula streaming query")
        .collect::<miniexcel::Result<Vec<_>>>()
        .expect("stream cached formula values");
    let mut reader = XlsxReader::open(path).expect("open formula fixture");
    let materialized = reader.read_rows(&ReadOptions::default()).expect("read formula fixture");

    assert_eq!(streamed, materialized);
}

#[test]
fn reads_dynamic_rows_without_headers() {
    let mut reader = XlsxReader::open(common::fixture("TestDynamicQueryBasic_WithoutHead.xlsx"))
        .expect("open fixture");

    let rows = reader.read_rows(&ReadOptions::default()).expect("read rows");

    assert_eq!(rows.len(), 2);
    assert_eq!(rows[0]["A"], CellValue::String("MiniExcel".to_owned()));
    assert_eq!(rows[0]["B"], CellValue::Int(1));
    assert_eq!(rows[1]["A"], CellValue::String("Github".to_owned()));
    assert_eq!(rows[1]["B"], CellValue::Int(2));
}

#[test]
fn reads_dynamic_rows_with_headers() {
    let mut reader =
        XlsxReader::open(common::fixture("TestDynamicQueryBasic.xlsx")).expect("open fixture");
    let options = ReadOptions::new().with_header_mode(HeaderMode::FirstRow);

    let rows = reader.read_rows(&options).expect("read rows");

    assert_eq!(rows.len(), 2);
    assert_eq!(rows[0]["Column1"], CellValue::String("MiniExcel".to_owned()));
    assert_eq!(rows[0]["Column2"], CellValue::Int(1));
    assert_eq!(rows[1]["Column1"], CellValue::String("Github".to_owned()));
    assert_eq!(rows[1]["Column2"], CellValue::Int(2));
}

#[test]
fn preserves_and_can_ignore_empty_rows() {
    let path = common::fixture("TestCenterEmptyRow/TestCenterEmptyRow.xlsx");
    let mut reader = XlsxReader::open(&path).expect("open fixture");
    let rows = reader.read_rows(&ReadOptions::default()).expect("read rows");

    assert_eq!(rows.len(), 6);
    assert!(rows[3].values().all(CellValue::is_empty));

    let mut reader = XlsxReader::open(path).expect("open fixture");
    let options = ReadOptions::new().with_ignore_empty_rows(true);
    let rows = reader.read_rows(&options).expect("read rows");
    assert_eq!(rows.len(), 5);
    assert!(rows.iter().all(|row| row.values().any(|value| !value.is_empty())));
}

#[test]
fn selects_sheets_in_workbook_order() {
    let mut reader =
        XlsxReader::open(common::fixture("TestMultiSheet.xlsx")).expect("open fixture");
    assert_eq!(reader.sheet_names(), ["Sheet1", "Sheet2", "Sheet3"]);

    let rows =
        reader.read_rows(&ReadOptions::new().with_sheet_name("Sheet3")).expect("read Sheet3");
    assert_eq!(rows.len(), 5);
    assert_eq!(rows[0]["A"], CellValue::Int(3));
    assert_eq!(rows[0]["B"], CellValue::Int(3));
}

#[test]
fn preserves_self_closing_empty_rows() {
    let mut reader =
        XlsxReader::open(common::fixture("TestEmptySelfClosingRow.xlsx")).expect("open fixture");
    let rows = reader.read_rows(&ReadOptions::default()).expect("read rows");

    assert_eq!(rows.len(), 10);
    assert!(rows[0]["A"].is_empty());
    assert_eq!(rows[1]["A"], CellValue::Int(1));
    assert!(rows[2]["A"].is_empty());
    assert_eq!(rows[3]["A"], CellValue::Int(2));
    assert!(rows[4..9].iter().all(|row| row["A"].is_empty()));
    assert_eq!(rows[9]["A"], CellValue::Int(1));
}

#[test]
fn reads_from_an_a1_start_cell() {
    let mut reader =
        XlsxReader::open(common::fixture("TestTypeMapping.xlsx")).expect("open fixture");
    let start_cell: CellReference = "B6".parse().expect("valid A1 reference");
    let options = ReadOptions::new().with_start_cell(start_cell);

    let rows = reader.read_rows(&options).expect("read rows");

    assert_eq!(rows[0]["B"], CellValue::String("Raymond".to_owned()));
    assert_eq!(
        rows[0]["C"],
        CellValue::DateTime(
            NaiveDate::from_ymd_opt(2021, 12, 7).unwrap().and_hms_opt(0, 0, 0).unwrap()
        )
    );
    assert_eq!(rows[0]["D"], CellValue::Int(18));
}

#[test]
fn reads_cells_without_r_attributes() {
    let mut reader =
        XlsxReader::open(common::fixture("TestWihoutRAttribute.xlsx")).expect("open fixture");
    let rows = reader.read_rows(&ReadOptions::default()).expect("read rows");

    assert_eq!(rows.len(), 2);
    assert_eq!(rows[0].len(), 5);
    assert_eq!(rows[0]["A"], CellValue::Int(1));
    assert!(rows[0]["C"].is_empty());
    assert!(rows[0]["D"].is_empty());
    assert!(rows[0]["E"].is_empty());
    assert_eq!(rows[1]["A"], CellValue::Int(1));
    assert_eq!(rows[1]["B"], CellValue::String("\"<>+}{\\nHello World".to_owned()));
    assert_eq!(rows[1]["C"], CellValue::Bool(true));
}
