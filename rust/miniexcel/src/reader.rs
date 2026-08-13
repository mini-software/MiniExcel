use calamine::{Data, DataType, Range};

use crate::CellValue;

pub(crate) struct SelectedRow {
    pub(crate) excel_row: usize,
    pub(crate) values: Vec<Data>,
}

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
