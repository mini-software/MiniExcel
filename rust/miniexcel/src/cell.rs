use std::fmt;
use std::str::FromStr;

use chrono::{Duration, NaiveDate, NaiveDateTime, NaiveTime};
use indexmap::IndexMap;
use serde::{Deserialize, Serialize};

use crate::{Error, Result};

const MAX_EXCEL_COLUMN: usize = 16_383;
const MAX_EXCEL_ROW: usize = 1_048_575;

#[derive(Clone, Debug, PartialEq, Serialize, Deserialize)]
pub enum CellValue {
    Empty,
    Bool(bool),
    Int(i64),
    Float(f64),
    String(String),
    Date(NaiveDate),
    Time(NaiveTime),
    DateTime(NaiveDateTime),
    Duration(Duration),
    Error(String),
}

impl CellValue {
    #[must_use]
    pub const fn is_empty(&self) -> bool {
        matches!(self, Self::Empty)
    }
}

pub type DynamicRow = IndexMap<String, CellValue>;

#[derive(Clone, Copy, Debug, Eq, Hash, PartialEq)]
pub struct CellReference {
    row: usize,
    column: usize,
}

impl CellReference {
    pub const A1: Self = Self { row: 0, column: 0 };

    pub fn new(row: usize, column: usize) -> Result<Self> {
        if row > MAX_EXCEL_ROW || column > MAX_EXCEL_COLUMN {
            return Err(Error::InvalidCellReference(format!(
                "row {}, column {}",
                row + 1,
                column + 1
            )));
        }

        Ok(Self { row, column })
    }

    #[must_use]
    pub const fn row(self) -> usize {
        self.row
    }

    #[must_use]
    pub const fn column(self) -> usize {
        self.column
    }
}

impl Default for CellReference {
    fn default() -> Self {
        Self::A1
    }
}

impl fmt::Display for CellReference {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        let mut column = self.column + 1;
        let mut letters = [0_u8; 3];
        let mut length = 0;

        while column > 0 {
            column -= 1;
            letters[length] = b'A' + (column % 26) as u8;
            length += 1;
            column /= 26;
        }

        for letter in letters[..length].iter().rev() {
            formatter.write_str(char::from(*letter).encode_utf8(&mut [0; 4]))?;
        }

        write!(formatter, "{}", self.row + 1)
    }
}

impl FromStr for CellReference {
    type Err = Error;

    fn from_str(value: &str) -> Result<Self> {
        let reference = value.trim();
        let bytes = reference.as_bytes();
        let mut index = 0;

        if bytes.get(index) == Some(&b'$') {
            index += 1;
        }

        let column_start = index;
        let mut column = 0_usize;
        while let Some(byte) = bytes.get(index).copied() {
            if !byte.is_ascii_alphabetic() {
                break;
            }

            column = column
                .checked_mul(26)
                .and_then(|current| {
                    current.checked_add(usize::from(byte.to_ascii_uppercase() - b'A' + 1))
                })
                .ok_or_else(|| Error::InvalidCellReference(reference.to_owned()))?;
            index += 1;
        }

        if index == column_start || column == 0 || column - 1 > MAX_EXCEL_COLUMN {
            return Err(Error::InvalidCellReference(reference.to_owned()));
        }

        if bytes.get(index) == Some(&b'$') {
            index += 1;
        }

        let row_start = index;
        while bytes.get(index).is_some_and(u8::is_ascii_digit) {
            index += 1;
        }

        if index == row_start || index != bytes.len() {
            return Err(Error::InvalidCellReference(reference.to_owned()));
        }

        let row = reference[row_start..]
            .parse::<usize>()
            .map_err(|_| Error::InvalidCellReference(reference.to_owned()))?;
        if row == 0 || row - 1 > MAX_EXCEL_ROW {
            return Err(Error::InvalidCellReference(reference.to_owned()));
        }

        Ok(Self { row: row - 1, column: column - 1 })
    }
}

impl TryFrom<&str> for CellReference {
    type Error = Error;

    fn try_from(value: &str) -> Result<Self> {
        value.parse()
    }
}

#[cfg(test)]
mod tests {
    use super::CellReference;

    #[test]
    fn parses_and_formats_a1_references() {
        for (value, row, column, canonical) in [
            ("A1", 0, 0, "A1"),
            ("$b$2", 1, 1, "B2"),
            ("AA10", 9, 26, "AA10"),
            ("XFD1048576", 1_048_575, 16_383, "XFD1048576"),
        ] {
            let reference: CellReference = value.parse().expect("valid cell reference");
            assert_eq!(reference.row(), row);
            assert_eq!(reference.column(), column);
            assert_eq!(reference.to_string(), canonical);
        }
    }

    #[test]
    fn rejects_invalid_a1_references() {
        for value in ["", "A", "1", "A0", "XFE1", "A1048577", "A1x", "$$A1"] {
            assert!(value.parse::<CellReference>().is_err(), "{value} should be invalid");
        }
    }
}
