mod ooxml;

use std::iter::FusedIterator;
use std::marker::PhantomData;
use std::path::Path;

use calamine::{Data, RangeDeserializerBuilder};
use serde::de::DeserializeOwned;

use crate::reader::{column_names, header_names, row_to_range, to_cell_value, trim_header_row};
use crate::{DynamicRow, Error, ReadOptions, Result};

use self::ooxml::StreamingRawRows;

enum Headers {
    FirstRow(Vec<Option<String>>),
    ColumnLetters { start_column: usize, headers: Option<Vec<Option<String>>> },
}

/// A bounded-memory iterator over dynamic XLSX rows.
pub struct StreamingRows {
    rows: StreamingRawRows,
    headers: Headers,
}

impl StreamingRows {
    pub(crate) fn open(path: impl AsRef<Path>, options: &ReadOptions) -> Result<Self> {
        let mut rows = StreamingRawRows::open(path, options)?;
        let headers = if options.uses_headers(false) {
            let headers =
                rows.next().transpose()?.map_or_else(Vec::new, |row| header_names(&row.values));
            Headers::FirstRow(headers)
        } else {
            Headers::ColumnLetters { start_column: options.start_cell().column(), headers: None }
        };
        Ok(Self { rows, headers })
    }
}

impl Iterator for StreamingRows {
    type Item = Result<DynamicRow>;

    fn next(&mut self) -> Option<Self::Item> {
        let selected_row = match self.rows.next()? {
            Ok(row) => row,
            Err(error) => return Some(Err(error)),
        };
        let headers = match &mut self.headers {
            Headers::FirstRow(headers) => headers,
            Headers::ColumnLetters { start_column, headers } => headers
                .get_or_insert_with(|| column_names(*start_column, selected_row.values.len())),
        };
        let mut row = DynamicRow::with_capacity(headers.len());
        for (column, header) in headers.iter().enumerate() {
            let Some(header) = header else {
                continue;
            };
            let value =
                selected_row.values.get(column).map_or(crate::CellValue::Empty, to_cell_value);
            row.insert(header.clone(), value);
        }
        Some(Ok(row))
    }
}

impl FusedIterator for StreamingRows {}

/// A bounded-memory iterator that deserializes XLSX rows through Serde.
pub struct StreamingTypedRows<T> {
    rows: StreamingRawRows,
    headers: Option<Vec<Data>>,
    sheet_name: String,
    marker: PhantomData<fn() -> T>,
}

impl<T> StreamingTypedRows<T>
where
    T: DeserializeOwned,
{
    pub(crate) fn open(path: impl AsRef<Path>, options: &ReadOptions) -> Result<Self> {
        let mut rows = StreamingRawRows::open(path, options)?;
        let sheet_name = rows.sheet_name().to_owned();
        let headers = if options.uses_headers(true) {
            rows.next().transpose()?.map(|mut row| {
                if options.trim_headers() {
                    trim_header_row(&mut row.values);
                }
                row.values
            })
        } else {
            None
        };
        Ok(Self { rows, headers, sheet_name, marker: PhantomData })
    }
}

impl<T> Iterator for StreamingTypedRows<T>
where
    T: DeserializeOwned,
{
    type Item = Result<T>;

    fn next(&mut self) -> Option<Self::Item> {
        let row = match self.rows.next()? {
            Ok(row) => row,
            Err(error) => return Some(Err(error)),
        };
        let range = row_to_range(self.headers.as_deref(), &row.values);
        let mut builder = RangeDeserializerBuilder::new();
        builder.has_headers(self.headers.is_some());
        let result = builder
            .from_range::<Data, T>(&range)
            .and_then(|mut iterator| {
                iterator.next().unwrap_or_else(|| {
                    Err(calamine::DeError::Custom(
                        "the selected Excel row did not produce a value".to_owned(),
                    ))
                })
            })
            .map_err(|source| Error::Deserialize {
                sheet: self.sheet_name.clone(),
                row: row.excel_row + 1,
                source,
            });
        Some(result)
    }
}

impl<T> FusedIterator for StreamingTypedRows<T> where T: DeserializeOwned {}
