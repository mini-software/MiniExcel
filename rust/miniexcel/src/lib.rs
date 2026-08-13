#![forbid(unsafe_code)]

//! Experimental Rust XLSX support for MiniExcel.

mod cell;
mod error;
mod facade;
mod options;
mod reader;
pub mod serde_helpers;
mod streaming;
mod writer;

pub use cell::{CellReference, CellValue, DynamicRow};
pub use error::{Error, Result};
pub use facade::MiniExcel;
pub use options::{HeaderMode, ReadOptions, WriteOptions, WriteSummary};
pub use reader::{DynamicRows, TypedRows, XlsxReader};
pub use streaming::{StreamingRows, StreamingTypedRows};
pub use writer::XlsxWriter;
