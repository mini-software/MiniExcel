use std::collections::HashSet;
use std::io::Write;
use std::path::Path;

use indexmap::IndexSet;
use rust_xlsxwriter::{CustomSerializeField, Format, SerializeFieldOptions, Workbook, Worksheet};
use serde::Serialize;

use crate::{CellValue, DynamicRow, Error, Result, WriteOptions, WriteSummary};

const MAX_EXCEL_ROWS: usize = 1_048_576;
const MAX_EXCEL_COLUMNS: usize = 16_384;

pub struct XlsxWriter {
    workbook: Workbook,
    sheet_names: HashSet<String>,
}

impl XlsxWriter {
    #[must_use]
    pub fn new() -> Self {
        Self::default()
    }

    pub fn add_rows(
        &mut self,
        rows: &[DynamicRow],
        options: &WriteOptions,
    ) -> Result<WriteSummary> {
        let mut schema = IndexSet::new();
        for row in rows {
            schema.extend(row.keys().cloned());
        }

        if schema.is_empty() && (!rows.is_empty() || options.print_header()) {
            return Err(Error::MissingSchema);
        }

        let schema: Vec<String> = schema.into_iter().collect();
        self.add_rows_with_schema(&schema, rows, options)
    }

    pub fn add_rows_with_schema(
        &mut self,
        schema: &[String],
        rows: &[DynamicRow],
        options: &WriteOptions,
    ) -> Result<WriteSummary> {
        validate_sheet_name(options.sheet_name(), &self.sheet_names)?;
        validate_schema(schema)?;
        validate_dimensions(rows.len(), schema.len(), options.print_header())?;

        let mut worksheet = Worksheet::new();
        worksheet.set_name(options.sheet_name())?;

        let mut output_row = 0_u32;
        if options.print_header() {
            for (column, header) in schema.iter().enumerate() {
                worksheet.write_string(0, column as u16, header)?;
            }
            output_row = 1;
        }

        let formats = CellFormats::new(options);
        for row in rows {
            for (column, header) in schema.iter().enumerate() {
                let value = row.get(header).unwrap_or(&CellValue::Empty);
                write_cell(&mut worksheet, output_row, column as u16, value, &formats)?;
            }
            output_row += 1;
        }

        self.workbook.push_worksheet(worksheet);
        self.sheet_names.insert(normalized_sheet_name(options.sheet_name()));
        Ok(WriteSummary::new(options.sheet_name().to_owned(), rows.len()))
    }

    pub fn save(&mut self, path: impl AsRef<Path>) -> Result<()> {
        self.workbook.save(path)?;
        Ok(())
    }

    pub fn to_bytes(&mut self) -> Result<Vec<u8>> {
        Ok(self.workbook.save_to_buffer()?)
    }

    pub fn save_to_writer<W>(&mut self, writer: W) -> Result<()>
    where
        W: Write + Send,
    {
        self.workbook.save_to_writer(writer)?;
        Ok(())
    }

    pub fn add_serialized<T>(&mut self, rows: &[T], options: &WriteOptions) -> Result<WriteSummary>
    where
        T: Serialize,
    {
        validate_sheet_name(options.sheet_name(), &self.sheet_names)?;
        validate_dimensions(rows.len(), 1, options.print_header())?;
        let Some(first) = rows.first() else {
            if options.print_header() {
                return Err(Error::MissingSchema);
            }

            let mut worksheet = Worksheet::new();
            worksheet.set_name(options.sheet_name())?;
            self.workbook.push_worksheet(worksheet);
            self.sheet_names.insert(normalized_sheet_name(options.sheet_name()));
            return Ok(WriteSummary::new(options.sheet_name().to_owned(), 0));
        };

        let mut worksheet = Worksheet::new();
        worksheet.set_name(options.sheet_name())?;
        let custom_headers: Vec<CustomSerializeField> = options
            .column_formats()
            .iter()
            .map(|(field_name, number_format)| {
                CustomSerializeField::new(field_name)
                    .set_value_format(Format::new().set_num_format(number_format))
            })
            .collect();
        let mut header_options = SerializeFieldOptions::new().hide_headers(!options.print_header());
        if !custom_headers.is_empty() {
            header_options = header_options.set_custom_headers(&custom_headers);
        }
        worksheet.serialize_headers_with_options(0, 0, first, &header_options)?;
        for row in rows {
            worksheet.serialize(row)?;
        }

        self.workbook.push_worksheet(worksheet);
        self.sheet_names.insert(normalized_sheet_name(options.sheet_name()));
        Ok(WriteSummary::new(options.sheet_name().to_owned(), rows.len()))
    }
}

impl Default for XlsxWriter {
    fn default() -> Self {
        Self { workbook: Workbook::new(), sheet_names: HashSet::new() }
    }
}

struct CellFormats {
    blank: Format,
    date: Format,
    time: Format,
    datetime: Format,
    duration: Format,
}

impl CellFormats {
    fn new(options: &WriteOptions) -> Self {
        Self {
            blank: Format::new().set_num_format("@"),
            date: Format::new().set_num_format(options.date_format()),
            time: Format::new().set_num_format(options.time_format()),
            datetime: Format::new().set_num_format(options.datetime_format()),
            duration: Format::new().set_num_format(options.duration_format()),
        }
    }
}

fn write_cell(
    worksheet: &mut Worksheet,
    row: u32,
    column: u16,
    value: &CellValue,
    formats: &CellFormats,
) -> Result<()> {
    match value {
        CellValue::Empty => {
            worksheet.write_blank(row, column, &formats.blank)?;
        }
        CellValue::Bool(value) => {
            worksheet.write_boolean(row, column, *value)?;
        }
        CellValue::Int(value) => {
            worksheet.write(row, column, *value)?;
        }
        CellValue::Float(value) => {
            worksheet.write_number(row, column, *value)?;
        }
        CellValue::String(value) | CellValue::Error(value) => {
            worksheet.write_string(row, column, value)?;
        }
        CellValue::Date(value) => {
            worksheet.write_datetime_with_format(row, column, value, &formats.date)?;
        }
        CellValue::Time(value) => {
            worksheet.write_datetime_with_format(row, column, value, &formats.time)?;
        }
        CellValue::DateTime(value) => {
            worksheet.write_datetime_with_format(row, column, value, &formats.datetime)?;
        }
        CellValue::Duration(value) => {
            let excel_days = value.num_milliseconds() as f64 / 86_400_000.0;
            worksheet.write_number_with_format(row, column, excel_days, &formats.duration)?;
        }
    }
    Ok(())
}

fn validate_sheet_name(name: &str, existing_names: &HashSet<String>) -> Result<()> {
    if name.is_empty() {
        return Err(invalid_sheet_name(name, "name cannot be blank"));
    }
    if name.chars().count() > 31 {
        return Err(invalid_sheet_name(name, "name cannot exceed 31 characters"));
    }
    if name.chars().any(|character| matches!(character, '[' | ']' | ':' | '*' | '?' | '/' | '\\')) {
        return Err(invalid_sheet_name(name, "name contains an invalid character"));
    }
    if name.starts_with('\'') || name.ends_with('\'') {
        return Err(invalid_sheet_name(name, "name cannot start or end with an apostrophe"));
    }
    if existing_names.contains(&normalized_sheet_name(name)) {
        return Err(Error::DuplicateSheetName(name.to_owned()));
    }
    Ok(())
}

fn invalid_sheet_name(name: &str, reason: &'static str) -> Error {
    Error::InvalidSheetName { name: name.to_owned(), reason }
}

fn normalized_sheet_name(name: &str) -> String {
    name.to_lowercase()
}

fn validate_schema(schema: &[String]) -> Result<()> {
    let mut names = HashSet::with_capacity(schema.len());
    for name in schema {
        if !names.insert(name) {
            return Err(Error::DuplicateColumnName(name.clone()));
        }
    }
    Ok(())
}

fn validate_dimensions(rows: usize, columns: usize, print_header: bool) -> Result<()> {
    let output_rows = rows.saturating_add(usize::from(print_header));
    if output_rows > MAX_EXCEL_ROWS || columns > MAX_EXCEL_COLUMNS {
        return Err(Error::WorksheetLimit { rows: output_rows, columns });
    }
    Ok(())
}
