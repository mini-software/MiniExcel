use std::fs::File;
use std::io::{BufReader, Read, Seek};
use std::path::Path;

use calamine::{Data, DataType, Range, RangeDeserializerBuilder, Reader, Xlsx};
use serde::de::DeserializeOwned;

use crate::{CellValue, DynamicRow, Error, ReadOptions, Result};

pub struct XlsxReader<R> {
    workbook: Xlsx<R>,
}

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

    pub fn read_rows(&mut self, options: &ReadOptions) -> Result<Vec<DynamicRow>> {
        let (sheet_name, range) = self.read_range(options)?;
        let rows = select_rows(&range, options);
        if rows.is_empty() {
            return Ok(Vec::new());
        }

        let uses_headers = options.uses_headers(false);
        let (headers, data_rows) = if uses_headers {
            (header_names(&rows[0].values), &rows[1..])
        } else {
            (column_names(options.start_cell().column(), rows[0].values.len()), rows.as_slice())
        };

        let mut result = Vec::with_capacity(data_rows.len());
        for selected_row in data_rows {
            let mut row = DynamicRow::with_capacity(headers.len());
            for (column, header) in headers.iter().enumerate() {
                let Some(header) = header else {
                    continue;
                };
                let value = selected_row.values.get(column).map_or(CellValue::Empty, to_cell_value);
                row.insert(header.clone(), value);
            }
            result.push(row);
        }

        debug_assert!(!sheet_name.is_empty());
        Ok(result)
    }

    pub fn deserialize<T>(&mut self, options: &ReadOptions) -> Result<Vec<T>>
    where
        T: DeserializeOwned,
    {
        let (sheet_name, range) = self.read_range(options)?;
        let mut rows = select_rows(&range, options);
        if rows.is_empty() {
            return Ok(Vec::new());
        }

        let uses_headers = options.uses_headers(true);
        if uses_headers && options.trim_headers() {
            trim_header_row(&mut rows[0].values);
        }

        let normalized = rows_to_range(&rows);
        let mut builder = RangeDeserializerBuilder::new();
        builder.has_headers(uses_headers);
        let iterator = builder.from_range::<Data, T>(&normalized).map_err(|source| {
            Error::Deserialize { sheet: sheet_name.clone(), row: rows[0].excel_row + 1, source }
        })?;

        let data_offset = usize::from(uses_headers);
        let mut result = Vec::with_capacity(rows.len().saturating_sub(data_offset));
        for (index, item) in iterator.enumerate() {
            let row = rows
                .get(index + data_offset)
                .map_or(rows[0].excel_row + index + data_offset + 1, |row| row.excel_row + 1);
            result.push(item.map_err(|source| Error::Deserialize {
                sheet: sheet_name.clone(),
                row,
                source,
            })?);
        }
        Ok(result)
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

struct SelectedRow {
    excel_row: usize,
    values: Vec<Data>,
}

fn select_rows(range: &Range<Data>, options: &ReadOptions) -> Vec<SelectedRow> {
    let Some((end_row, end_column)) = range.end() else {
        return Vec::new();
    };
    let start = options.start_cell();
    let end_row = end_row as usize;
    let end_column = end_column as usize;
    if start.row() > end_row || start.column() > end_column {
        return Vec::new();
    }

    let mut rows = Vec::with_capacity(end_row - start.row() + 1);
    for row_index in start.row()..=end_row {
        let mut values = Vec::with_capacity(end_column - start.column() + 1);
        for column_index in start.column()..=end_column {
            values.push(
                range
                    .get_value((row_index as u32, column_index as u32))
                    .cloned()
                    .unwrap_or(Data::Empty),
            );
        }

        if options.ignore_empty_rows() && values.iter().all(DataType::is_empty) {
            continue;
        }
        rows.push(SelectedRow { excel_row: row_index, values });
    }
    rows
}

fn header_names(values: &[Data]) -> Vec<Option<String>> {
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

fn column_names(start_column: usize, width: usize) -> Vec<Option<String>> {
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

fn trim_header_row(values: &mut [Data]) {
    for value in values {
        if let Data::String(header) = value {
            *header = header.trim().to_owned();
        }
    }
}

fn rows_to_range(rows: &[SelectedRow]) -> Range<Data> {
    let width = rows.first().map_or(0, |row| row.values.len());
    if rows.is_empty() || width == 0 {
        return Range::empty();
    }

    let mut range = Range::new((0, 0), ((rows.len() - 1) as u32, (width - 1) as u32));
    for (row_index, row) in rows.iter().enumerate() {
        for (column_index, value) in row.values.iter().enumerate() {
            range.set_value((row_index as u32, column_index as u32), value.clone());
        }
    }
    range
}

fn to_cell_value(value: &Data) -> CellValue {
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
