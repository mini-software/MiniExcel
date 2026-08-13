use std::path::Path;

use serde::de::DeserializeOwned;

use crate::{ReadOptions, Result, StreamingRows, StreamingTypedRows};

/// Convenience entry points for the common path-based MiniExcel workflow.
pub struct MiniExcel;

impl MiniExcel {
    /// Streams dynamic rows from the first worksheet without a header row.
    pub fn query(path: impl AsRef<Path>) -> Result<StreamingRows> {
        Self::query_with_options(path, &ReadOptions::default())
    }

    /// Streams dynamic rows using explicit read options.
    pub fn query_with_options(
        path: impl AsRef<Path>,
        options: &ReadOptions,
    ) -> Result<StreamingRows> {
        StreamingRows::open(path, options)
    }

    /// Streams and deserializes rows from the first worksheet through Serde.
    pub fn query_as<T>(path: impl AsRef<Path>) -> Result<StreamingTypedRows<T>>
    where
        T: DeserializeOwned,
    {
        Self::query_as_with_options(path, &ReadOptions::default())
    }

    /// Streams and deserializes rows through Serde using explicit read options.
    pub fn query_as_with_options<T>(
        path: impl AsRef<Path>,
        options: &ReadOptions,
    ) -> Result<StreamingTypedRows<T>>
    where
        T: DeserializeOwned,
    {
        StreamingTypedRows::open(path, options)
    }
}
