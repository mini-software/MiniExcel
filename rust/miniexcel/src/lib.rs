#![forbid(unsafe_code)]

//! Experimental Rust XLSX support for MiniExcel.

mod cell;
mod error;
mod options;
mod reader;
pub mod serde_helpers;
mod writer;

pub use cell::{CellReference, CellValue, DynamicRow};
pub use error::{Error, Result};
pub use options::{HeaderMode, ReadOptions, WriteOptions, WriteSummary};
pub use reader::XlsxReader;
pub use writer::XlsxWriter;
