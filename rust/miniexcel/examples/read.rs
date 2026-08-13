use std::path::PathBuf;

use miniexcel::{HeaderMode, ReadOptions, XlsxReader};

fn main() -> miniexcel::Result<()> {
    let path = std::env::args().nth(1).map_or_else(default_fixture, PathBuf::from);
    let mut reader = XlsxReader::open(&path)?;
    let options = ReadOptions::new().with_header_mode(HeaderMode::FirstRow);

    println!("Sheets: {:?}", reader.sheet_names());
    for row in reader.read_rows(&options)?.into_iter().take(5) {
        println!("{row:?}");
    }
    Ok(())
}

fn default_fixture() -> PathBuf {
    PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("..")
        .join("..")
        .join("tests")
        .join("data")
        .join("xlsx")
        .join("TestDynamicQueryBasic.xlsx")
}
