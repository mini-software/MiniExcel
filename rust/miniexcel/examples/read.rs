use std::path::PathBuf;

use miniexcel::{HeaderMode, MiniExcel, ReadOptions};

fn main() -> miniexcel::Result<()> {
    let path = std::env::args().nth(1).map_or_else(default_fixture, PathBuf::from);
    let options = ReadOptions::new().with_header_mode(HeaderMode::FirstRow);

    for row in MiniExcel::query_with_options(path, &options)?.take(5) {
        println!("{:?}", row?);
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
