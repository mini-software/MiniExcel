use thiserror::Error;

#[derive(Debug, Error)]
pub enum Error {
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

pub type Result<T> = std::result::Result<T, Error>;
