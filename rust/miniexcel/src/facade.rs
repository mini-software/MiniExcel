use std::path::Path;

use calamine::{Reader, Xlsx, open_workbook};
use serde::Serialize;
use serde::de::DeserializeOwned;

use crate::streaming::{StreamingRows, StreamingTypedRows};
use crate::writer::XlsxWriter;
use crate::{DynamicRow, ReadOptions, Result, WriteOptions};

/// Convenience entry points for the common path-based MiniExcel workflow.
pub struct MiniExcel;

impl MiniExcel {
    /// Returns worksheet names in workbook order.
    pub fn get_sheet_names(path: impl AsRef<Path>) -> Result<Vec<String>> {
        let workbook: Xlsx<_> = open_workbook(path)?;
        Ok(workbook.sheet_names())
    }

    /// Returns worksheet names from an in-memory XLSX workbook.
    pub fn get_sheet_names_from_bytes(bytes: &[u8]) -> Result<Vec<String>> {
        crate::streaming::sheet_names_from_bytes(bytes)
    }

    /// Streams dynamic rows from the first worksheet without a header row.
    pub fn query(
        path: impl AsRef<Path>,
    ) -> Result<Box<dyn Iterator<Item = Result<DynamicRow>> + Send>> {
        Self::query_with_options(path, &ReadOptions::default())
    }

    /// Streams dynamic rows using explicit read options.
    pub fn query_with_options(
        path: impl AsRef<Path>,
        options: &ReadOptions,
    ) -> Result<Box<dyn Iterator<Item = Result<DynamicRow>> + Send>> {
        Ok(Box::new(StreamingRows::open(path, options)?))
    }

    /// Returns the selected dynamic column names, or an empty vector when no data rows exist.
    pub fn get_columns(path: impl AsRef<Path>, options: &ReadOptions) -> Result<Vec<String>> {
        let mut rows = Self::query_with_options(path, options)?;
        Ok(rows.next().transpose()?.map_or_else(Vec::new, |row| row.into_keys().collect()))
    }

    /// Reads dynamic rows from an in-memory XLSX workbook.
    ///
    /// Unlike path queries, this method materializes the selected rows and is intended for
    /// browser uploads and other environments without filesystem access.
    pub fn query_bytes(bytes: &[u8], options: &ReadOptions) -> Result<Vec<DynamicRow>> {
        crate::streaming::query_bytes(bytes, options)
    }

    /// Streams and deserializes rows from the first worksheet through Serde.
    pub fn query_as<T>(path: impl AsRef<Path>) -> Result<Box<dyn Iterator<Item = Result<T>> + Send>>
    where
        T: DeserializeOwned + 'static,
    {
        Self::query_as_with_options(path, &ReadOptions::default())
    }

    /// Streams and deserializes rows through Serde using explicit read options.
    pub fn query_as_with_options<T>(
        path: impl AsRef<Path>,
        options: &ReadOptions,
    ) -> Result<Box<dyn Iterator<Item = Result<T>> + Send>>
    where
        T: DeserializeOwned + 'static,
    {
        Ok(Box::new(StreamingTypedRows::open(path, options)?))
    }

    /// Creates a new XLSX workbook from dynamic rows.
    pub fn save_as(path: impl AsRef<Path>, rows: &[DynamicRow]) -> Result<()> {
        Self::save_as_with_options(path, rows, &WriteOptions::default())
    }

    /// Creates a new XLSX workbook from dynamic rows using explicit options.
    pub fn save_as_with_options(
        path: impl AsRef<Path>,
        rows: &[DynamicRow],
        options: &WriteOptions,
    ) -> Result<()> {
        let mut writer = XlsxWriter::new();
        writer.add_rows(rows, options)?;
        writer.save(path)
    }

    /// Creates an in-memory XLSX workbook from dynamic rows.
    pub fn save_as_bytes(rows: &[DynamicRow], options: &WriteOptions) -> Result<Vec<u8>> {
        let mut writer = XlsxWriter::new();
        writer.add_rows(rows, options)?;
        writer.save_to_bytes()
    }

    /// Creates a new XLSX workbook using an explicit dynamic schema.
    pub fn save_as_with_schema(
        path: impl AsRef<Path>,
        schema: &[String],
        rows: &[DynamicRow],
        options: &WriteOptions,
    ) -> Result<()> {
        let mut writer = XlsxWriter::new();
        writer.add_rows_with_schema(schema, rows, options)?;
        writer.save(path)
    }

    /// Creates a new XLSX workbook from Serde-serializable rows.
    pub fn save_as_serialized<T>(path: impl AsRef<Path>, rows: &[T]) -> Result<()>
    where
        T: Serialize,
    {
        Self::save_as_serialized_with_options(path, rows, &WriteOptions::default())
    }

    /// Creates a new XLSX workbook from Serde rows using explicit options.
    pub fn save_as_serialized_with_options<T>(
        path: impl AsRef<Path>,
        rows: &[T],
        options: &WriteOptions,
    ) -> Result<()>
    where
        T: Serialize,
    {
        let mut writer = XlsxWriter::new();
        writer.add_serialized(rows, options)?;
        writer.save(path)
    }
}
