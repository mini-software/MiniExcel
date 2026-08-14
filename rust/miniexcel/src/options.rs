use indexmap::IndexMap;

use crate::CellReference;

#[derive(Clone, Copy, Debug, Default, Eq, PartialEq)]
pub enum HeaderMode {
    #[default]
    Auto,
    None,
    FirstRow,
}

#[derive(Clone, Debug, Eq, PartialEq)]
pub struct ReadOptions {
    sheet_name: Option<String>,
    start_cell: CellReference,
    end_cell: Option<CellReference>,
    header_mode: HeaderMode,
    ignore_empty_rows: bool,
    trim_headers: bool,
}

impl ReadOptions {
    #[must_use]
    pub fn new() -> Self {
        Self::default()
    }

    #[must_use]
    pub fn with_sheet_name(mut self, sheet_name: impl Into<String>) -> Self {
        self.sheet_name = Some(sheet_name.into());
        self
    }

    #[must_use]
    pub const fn with_start_cell(mut self, start_cell: CellReference) -> Self {
        self.start_cell = start_cell;
        self
    }

    #[must_use]
    pub const fn with_end_cell(mut self, end_cell: CellReference) -> Self {
        self.end_cell = Some(end_cell);
        self
    }

    #[must_use]
    pub const fn with_header_mode(mut self, header_mode: HeaderMode) -> Self {
        self.header_mode = header_mode;
        self
    }

    #[must_use]
    pub const fn with_ignore_empty_rows(mut self, ignore_empty_rows: bool) -> Self {
        self.ignore_empty_rows = ignore_empty_rows;
        self
    }

    #[must_use]
    pub const fn with_trim_headers(mut self, trim_headers: bool) -> Self {
        self.trim_headers = trim_headers;
        self
    }

    #[must_use]
    pub(crate) fn sheet_name(&self) -> Option<&str> {
        self.sheet_name.as_deref()
    }

    #[must_use]
    pub(crate) const fn start_cell(&self) -> CellReference {
        self.start_cell
    }

    #[must_use]
    pub(crate) const fn end_cell(&self) -> Option<CellReference> {
        self.end_cell
    }

    #[must_use]
    pub(crate) const fn ignore_empty_rows(&self) -> bool {
        self.ignore_empty_rows
    }

    #[must_use]
    pub(crate) const fn trim_headers(&self) -> bool {
        self.trim_headers
    }

    pub(crate) const fn uses_headers(&self, typed: bool) -> bool {
        match self.header_mode {
            HeaderMode::Auto => typed,
            HeaderMode::None => false,
            HeaderMode::FirstRow => true,
        }
    }
}

impl Default for ReadOptions {
    fn default() -> Self {
        Self {
            sheet_name: None,
            start_cell: CellReference::A1,
            end_cell: None,
            header_mode: HeaderMode::Auto,
            ignore_empty_rows: false,
            trim_headers: true,
        }
    }
}

#[derive(Clone, Debug, Eq, PartialEq)]
pub struct WriteOptions {
    sheet_name: String,
    print_header: bool,
    date_format: String,
    time_format: String,
    datetime_format: String,
    duration_format: String,
    column_formats: IndexMap<String, String>,
}

impl WriteOptions {
    #[must_use]
    pub fn new() -> Self {
        Self::default()
    }

    #[must_use]
    pub fn with_sheet_name(mut self, sheet_name: impl Into<String>) -> Self {
        self.sheet_name = sheet_name.into();
        self
    }

    #[must_use]
    pub const fn with_print_header(mut self, print_header: bool) -> Self {
        self.print_header = print_header;
        self
    }

    #[must_use]
    pub fn with_date_format(mut self, format: impl Into<String>) -> Self {
        self.date_format = format.into();
        self
    }

    #[must_use]
    pub fn with_time_format(mut self, format: impl Into<String>) -> Self {
        self.time_format = format.into();
        self
    }

    #[must_use]
    pub fn with_datetime_format(mut self, format: impl Into<String>) -> Self {
        self.datetime_format = format.into();
        self
    }

    #[must_use]
    pub fn with_duration_format(mut self, format: impl Into<String>) -> Self {
        self.duration_format = format.into();
        self
    }

    #[must_use]
    pub fn with_column_format(
        mut self,
        field_name: impl Into<String>,
        format: impl Into<String>,
    ) -> Self {
        self.column_formats.insert(field_name.into(), format.into());
        self
    }

    #[must_use]
    pub(crate) fn sheet_name(&self) -> &str {
        &self.sheet_name
    }

    #[must_use]
    pub(crate) const fn print_header(&self) -> bool {
        self.print_header
    }

    #[must_use]
    pub(crate) fn date_format(&self) -> &str {
        &self.date_format
    }

    #[must_use]
    pub(crate) fn time_format(&self) -> &str {
        &self.time_format
    }

    #[must_use]
    pub(crate) fn datetime_format(&self) -> &str {
        &self.datetime_format
    }

    #[must_use]
    pub(crate) fn duration_format(&self) -> &str {
        &self.duration_format
    }

    #[must_use]
    pub(crate) fn column_formats(&self) -> &IndexMap<String, String> {
        &self.column_formats
    }
}

impl Default for WriteOptions {
    fn default() -> Self {
        Self {
            sheet_name: "Sheet1".to_owned(),
            print_header: true,
            date_format: "yyyy-mm-dd".to_owned(),
            time_format: "hh:mm:ss".to_owned(),
            datetime_format: "yyyy-mm-dd hh:mm:ss".to_owned(),
            duration_format: "[h]:mm:ss".to_owned(),
            column_formats: IndexMap::new(),
        }
    }
}
