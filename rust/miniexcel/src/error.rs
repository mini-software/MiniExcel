use thiserror::Error;

#[derive(Debug, Error)]
#[error(transparent)]
pub struct Error(#[from] ErrorKind);

#[derive(Debug, Error)]
enum ErrorKind {
    #[error("I/O error: {0}")]
    Io(#[from] std::io::Error),

    #[error("failed to read XLSX data: {0}")]
    Read(#[from] calamine::XlsxError),

    #[error("failed to stream XLSX data: {0}")]
    Stream(String),

    #[error("failed to write XLSX data: {0}")]
    Write(#[from] rust_xlsxwriter::XlsxError),

    #[error("invalid A1 cell reference: {0}")]
    InvalidCellReference(String),

    #[error("worksheet '{0}' was not found")]
    SheetNotFound(String),

    #[error("the workbook does not contain any worksheets")]
    NoWorksheets,

    #[error("invalid worksheet name '{name}': {reason}")]
    InvalidSheetName { name: String, reason: &'static str },

    #[error("worksheet name '{0}' is already in use")]
    DuplicateSheetName(String),

    #[error("cannot write headers for an empty data set without an explicit schema")]
    MissingSchema,

    #[error("worksheet data exceeds Excel limits: {rows} rows, {columns} columns")]
    WorksheetLimit { rows: usize, columns: usize },

    #[error("column name '{0}' appears more than once in the schema")]
    DuplicateColumnName(String),

    #[error("failed to deserialize worksheet '{sheet}' at Excel row {row}: {source}")]
    Deserialize {
        sheet: String,
        row: usize,
        #[source]
        source: calamine::DeError,
    },
}

impl Error {
    pub(crate) fn stream(message: impl Into<String>) -> Self {
        ErrorKind::Stream(message.into()).into()
    }

    pub(crate) fn invalid_cell_reference(reference: impl Into<String>) -> Self {
        ErrorKind::InvalidCellReference(reference.into()).into()
    }

    pub(crate) fn sheet_not_found(sheet_name: impl Into<String>) -> Self {
        ErrorKind::SheetNotFound(sheet_name.into()).into()
    }

    pub(crate) fn no_worksheets() -> Self {
        ErrorKind::NoWorksheets.into()
    }

    pub(crate) fn invalid_sheet_name(name: impl Into<String>, reason: &'static str) -> Self {
        ErrorKind::InvalidSheetName { name: name.into(), reason }.into()
    }

    pub(crate) fn duplicate_sheet_name(name: impl Into<String>) -> Self {
        ErrorKind::DuplicateSheetName(name.into()).into()
    }

    pub(crate) fn missing_schema() -> Self {
        ErrorKind::MissingSchema.into()
    }

    pub(crate) fn worksheet_limit(rows: usize, columns: usize) -> Self {
        ErrorKind::WorksheetLimit { rows, columns }.into()
    }

    pub(crate) fn duplicate_column_name(name: impl Into<String>) -> Self {
        ErrorKind::DuplicateColumnName(name.into()).into()
    }

    pub(crate) fn deserialize(
        sheet: impl Into<String>,
        row: usize,
        source: calamine::DeError,
    ) -> Self {
        ErrorKind::Deserialize { sheet: sheet.into(), row, source }.into()
    }
}

impl From<std::io::Error> for Error {
    fn from(source: std::io::Error) -> Self {
        ErrorKind::Io(source).into()
    }
}

impl From<calamine::XlsxError> for Error {
    fn from(source: calamine::XlsxError) -> Self {
        ErrorKind::Read(source).into()
    }
}

impl From<rust_xlsxwriter::XlsxError> for Error {
    fn from(source: rust_xlsxwriter::XlsxError) -> Self {
        ErrorKind::Write(source).into()
    }
}

pub type Result<T> = std::result::Result<T, Error>;
