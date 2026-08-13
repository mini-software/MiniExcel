use std::fs::File;
use std::io::{BufReader, Read, Seek};
use std::iter::FusedIterator;
use std::marker::PhantomData;
use std::path::Path;

use calamine::{Data, DataType, Range, RangeDeserializerBuilder, Reader, Xlsx};
use serde::de::DeserializeOwned;

use crate::{CellValue, DynamicRow, Error, ReadOptions, Result};

pub struct XlsxReader<R> {
    workbook: Xlsx<R>,
}

/// An iterator that maps one selected worksheet row at a time.
pub struct DynamicRows {
    rows: SelectedRows,
    headers: Vec<Option<String>>,
}

impl DynamicRows {
    fn new(range: Range<Data>, options: &ReadOptions) -> Self {
        let mut rows = SelectedRows::new(range, options);
        let headers = if options.uses_headers(false) {
            rows.next().map_or_else(Vec::new, |row| header_names(&row.values))
        } else {
            column_names(options.start_cell().column(), rows.width())
        };
        Self { rows, headers }
    }
}

impl Iterator for DynamicRows {
    type Item = Result<DynamicRow>;

    fn next(&mut self) -> Option<Self::Item> {
        let selected_row = self.rows.next()?;
        let mut row = DynamicRow::with_capacity(self.headers.len());
        for (column, header) in self.headers.iter().enumerate() {
            let Some(header) = header else {
                continue;
            };
            let value = selected_row.values.get(column).map_or(CellValue::Empty, to_cell_value);
            row.insert(header.clone(), value);
        }
        Some(Ok(row))
    }

    fn size_hint(&self) -> (usize, Option<usize>) {
        self.rows.size_hint()
    }
}

impl FusedIterator for DynamicRows {}

/// An iterator that deserializes one selected worksheet row at a time.
pub struct TypedRows<T> {
    rows: SelectedRows,
    headers: Option<Vec<Data>>,
    sheet_name: String,
    marker: PhantomData<fn() -> T>,
}

impl<T> TypedRows<T> {
    fn new(sheet_name: String, range: Range<Data>, options: &ReadOptions) -> Self {
        let mut rows = SelectedRows::new(range, options);
        let headers = if options.uses_headers(true) {
            rows.next().map(|mut row| {
                if options.trim_headers() {
                    trim_header_row(&mut row.values);
                }
                row.values
            })
        } else {
            None
        };
        Self { rows, headers, sheet_name, marker: PhantomData }
    }
}

impl<T> Iterator for TypedRows<T>
where
    T: DeserializeOwned,
{
    type Item = Result<T>;

    fn next(&mut self) -> Option<Self::Item> {
        let row = self.rows.next()?;
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

    fn size_hint(&self) -> (usize, Option<usize>) {
        self.rows.size_hint()
    }
}

impl<T> FusedIterator for TypedRows<T> where T: DeserializeOwned {}

impl XlsxReader<BufReader<File>> {
    pub fn open(path: impl AsRef<Path>) -> Result<Self> {
        let file = File::open(path)?;
        Self::from_reader(BufReader::new(file))
    }
}

impl<R> XlsxReader<R>
where
    R: Read + Seek,
{
    pub fn from_reader(reader: R) -> Result<Self> {
        Ok(Self { workbook: Xlsx::new(reader)? })
    }

    #[must_use]
    pub fn sheet_names(&self) -> Vec<String> {
        self.workbook.sheet_names()
    }

    pub fn query(&mut self, options: &ReadOptions) -> Result<DynamicRows> {
        let (_, range) = self.read_range(options)?;
        Ok(DynamicRows::new(range, options))
    }

    pub fn read_rows(&mut self, options: &ReadOptions) -> Result<Vec<DynamicRow>> {
        self.query(options)?.collect()
    }

    pub fn deserialize<T>(&mut self, options: &ReadOptions) -> Result<Vec<T>>
    where
        T: DeserializeOwned,
    {
        self.query_as(options)?.collect()
    }

    pub fn query_as<T>(&mut self, options: &ReadOptions) -> Result<TypedRows<T>>
    where
        T: DeserializeOwned,
    {
        let (sheet_name, range) = self.read_range(options)?;
        Ok(TypedRows::new(sheet_name, range, options))
    }

    fn read_range(&mut self, options: &ReadOptions) -> Result<(String, Range<Data>)> {
        let sheet_names = self.workbook.sheet_names();
        let sheet_name = match options.sheet_name() {
            Some(sheet_name) if sheet_names.iter().any(|candidate| candidate == sheet_name) => {
                sheet_name.to_owned()
            }
            Some(sheet_name) => return Err(Error::SheetNotFound(sheet_name.to_owned())),
            None => sheet_names.into_iter().next().ok_or(Error::NoWorksheets)?,
        };
        let range = self.workbook.worksheet_range(&sheet_name)?;
        Ok((sheet_name, range))
    }
}

pub(crate) struct SelectedRow {
    pub(crate) excel_row: usize,
    pub(crate) values: Vec<Data>,
}

struct SelectedRows {
    range: Range<Data>,
    next_row: usize,
    end_row: Option<usize>,
    start_column: usize,
    end_column: usize,
    ignore_empty_rows: bool,
}

impl SelectedRows {
    fn new(range: Range<Data>, options: &ReadOptions) -> Self {
        let start = options.start_cell();
        let (end_row, end_column) = range.end().map_or((None, start.column()), |(row, column)| {
            let row = row as usize;
            let column = column as usize;
            if start.row() > row || start.column() > column {
                (None, start.column())
            } else {
                (Some(row), column)
            }
        });
        Self {
            range,
            next_row: start.row(),
            end_row,
            start_column: start.column(),
            end_column,
            ignore_empty_rows: options.ignore_empty_rows(),
        }
    }

    fn width(&self) -> usize {
        self.end_row.map_or(0, |_| self.end_column - self.start_column + 1)
    }
}

impl Iterator for SelectedRows {
    type Item = SelectedRow;

    fn next(&mut self) -> Option<Self::Item> {
        let end_row = self.end_row?;
        while self.next_row <= end_row {
            let excel_row = self.next_row;
            self.next_row += 1;
            let values = (self.start_column..=self.end_column)
                .map(|column| {
                    self.range
                        .get_value((excel_row as u32, column as u32))
                        .cloned()
                        .unwrap_or(Data::Empty)
                })
                .collect::<Vec<_>>();
            if self.ignore_empty_rows && values.iter().all(DataType::is_empty) {
                continue;
            }
            return Some(SelectedRow { excel_row, values });
        }
        None
    }

    fn size_hint(&self) -> (usize, Option<usize>) {
        let remaining = self.end_row.map_or(0, |end_row| {
            if self.next_row > end_row { 0 } else { end_row - self.next_row + 1 }
        });
        if self.ignore_empty_rows { (0, Some(remaining)) } else { (remaining, Some(remaining)) }
    }
}

impl FusedIterator for SelectedRows {}

pub(crate) fn header_names(values: &[Data]) -> Vec<Option<String>> {
    values
        .iter()
        .map(|value| {
            if value.is_empty() || value.to_string().trim().is_empty() {
                None
            } else {
                Some(value.to_string())
            }
        })
        .collect()
}

pub(crate) fn column_names(start_column: usize, width: usize) -> Vec<Option<String>> {
    (start_column..start_column + width).map(|column| Some(column_name(column))).collect()
}

fn column_name(mut column: usize) -> String {
    let mut letters = Vec::with_capacity(3);
    column += 1;
    while column > 0 {
        column -= 1;
        letters.push(char::from(b'A' + (column % 26) as u8));
        column /= 26;
    }
    letters.iter().rev().collect()
}

pub(crate) fn trim_header_row(values: &mut [Data]) {
    for value in values {
        if let Data::String(header) = value {
            *header = header.trim().to_owned();
        }
    }
}

pub(crate) fn row_to_range(headers: Option<&[Data]>, values: &[Data]) -> Range<Data> {
    let height = usize::from(headers.is_some()) + 1;
    let width = values.len();
    if width == 0 {
        return Range::empty();
    }

    let mut range = Range::new((0, 0), ((height - 1) as u32, (width - 1) as u32));
    if let Some(headers) = headers {
        for (column_index, value) in headers.iter().enumerate() {
            range.set_value((0, column_index as u32), value.clone());
        }
    }
    let row_index = u32::from(headers.is_some());
    for (column_index, value) in values.iter().enumerate() {
        range.set_value((row_index, column_index as u32), value.clone());
    }
    range
}

pub(crate) fn to_cell_value(value: &Data) -> CellValue {
    match value {
        Data::Empty => CellValue::Empty,
        Data::Bool(value) => CellValue::Bool(*value),
        Data::Int(value) => CellValue::Int(*value),
        Data::Float(value)
            if value.is_finite()
                && value.fract() == 0.0
                && *value >= i64::MIN as f64
                && *value < -(i64::MIN as f64) =>
        {
            CellValue::Int(*value as i64)
        }
        Data::Float(value) => CellValue::Float(*value),
        Data::String(value) => CellValue::String(value.clone()),
        Data::DateTime(value) if value.is_duration() => {
            value.as_duration().map_or(CellValue::Float(value.as_f64()), CellValue::Duration)
        }
        Data::DateTime(value) => {
            value.as_datetime().map_or(CellValue::Float(value.as_f64()), CellValue::DateTime)
        }
        Data::DateTimeIso(iso) if iso.contains('T') || iso.contains(' ') => {
            value.as_datetime().map_or_else(|| CellValue::String(iso.clone()), CellValue::DateTime)
        }
        Data::DateTimeIso(iso) if iso.contains(':') => {
            value.as_time().map_or_else(|| CellValue::String(iso.clone()), CellValue::Time)
        }
        Data::DateTimeIso(iso) => {
            value.as_date().map_or_else(|| CellValue::String(iso.clone()), CellValue::Date)
        }
        Data::DurationIso(iso) => {
            value.as_duration().map_or_else(|| CellValue::String(iso.clone()), CellValue::Duration)
        }
        Data::Error(error) => CellValue::Error(error.to_string()),
    }
}
