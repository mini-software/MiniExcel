use std::path::PathBuf;

use chrono::NaiveDate;
use miniexcel::{CellValue, DynamicRow, WriteOptions, XlsxWriter};

fn main() -> miniexcel::Result<()> {
    let output = std::env::args()
        .nth(1)
        .map_or_else(|| std::env::temp_dir().join("miniexcel-rust-example.xlsx"), PathBuf::from);

    let mut first = DynamicRow::new();
    first.insert("Name".to_owned(), CellValue::String("MiniExcel".to_owned()));
    first.insert("Version".to_owned(), CellValue::Int(2));
    first.insert(
        "ReleasedOn".to_owned(),
        CellValue::Date(NaiveDate::from_ymd_opt(2025, 1, 2).expect("valid example date")),
    );

    let mut second = DynamicRow::new();
    second.insert("Name".to_owned(), CellValue::String("MiniExcel Rust".to_owned()));
    second.insert("Version".to_owned(), CellValue::Int(1));

    let mut writer = XlsxWriter::new();
    let summary =
        writer.add_rows(&[first, second], &WriteOptions::new().with_sheet_name("Releases"))?;
    writer.save(&output)?;

    println!(
        "Wrote {} rows to '{}' in {}",
        summary.rows_written(),
        summary.sheet_name(),
        output.display()
    );
    Ok(())
}
