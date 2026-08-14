use std::collections::HashSet;
use std::env;
use std::error::Error;
use std::fs;
use std::io::{self, Write};
use std::path::{Path, PathBuf};
use std::process::{Command as ProcessCommand, ExitCode};

use chrono::NaiveDate;
use clap::{Args, Parser, Subcommand, ValueEnum};
use miniexcel::{
    CellReference, CellValue, DynamicRow, HeaderMode, MiniExcel, ReadOptions, WriteOptions,
};
use serde_json::{Map, Number, Value};

type CliResult<T> = Result<T, Box<dyn Error>>;

#[derive(Debug, Parser)]
#[command(name = "miniexcel", version, about = "Inspect and locally test MiniExcel XLSX behavior")]
struct Cli {
    #[command(subcommand)]
    command: CliCommand,
}

#[derive(Debug, Subcommand)]
enum CliCommand {
    /// List worksheet names in workbook order.
    Sheets {
        /// Path to an XLSX workbook.
        file: PathBuf,
    },
    /// Stream rows from an XLSX workbook.
    Query(QueryArgs),
    /// Create and read back a small workbook.
    WriteDemo {
        /// Destination XLSX path.
        output: PathBuf,
    },
    /// Run the shared Rust/.NET XLSX parity contract.
    Parity(ParityArgs),
}

#[derive(Debug, Args)]
struct QueryArgs {
    /// Path to an XLSX workbook.
    file: PathBuf,

    /// Worksheet name. The first worksheet is used when omitted.
    #[arg(long)]
    sheet: Option<String>,

    /// Treat the first selected row as column headers.
    #[arg(long)]
    header: bool,

    /// First cell to read, in A1 notation.
    #[arg(long, default_value = "A1")]
    start_cell: CellReference,

    /// Last cell to read, in A1 notation. Reads to the worksheet end when omitted.
    #[arg(long)]
    end_cell: Option<CellReference>,

    /// Omit rows whose selected cells are all empty.
    #[arg(long)]
    ignore_empty_rows: bool,

    /// Maximum rows to print. Use 0 for all rows.
    #[arg(long, default_value_t = 20)]
    limit: usize,

    /// Output representation.
    #[arg(long, value_enum, default_value = "table")]
    format: OutputFormat,
}

#[derive(Clone, Copy, Debug, ValueEnum)]
enum OutputFormat {
    Table,
    Json,
    Jsonl,
}

#[derive(Debug, Args)]
struct ParityArgs {
    /// Repository root. Auto-detected from the current directory when omitted.
    #[arg(long)]
    repo_root: Option<PathBuf>,

    /// Target framework for the .NET parity test.
    #[arg(long, default_value = "net10.0")]
    framework: String,

    /// Run only the Rust parity test.
    #[arg(long, conflicts_with = "dotnet_only")]
    rust_only: bool,

    /// Run only the .NET parity test.
    #[arg(long, conflicts_with = "rust_only")]
    dotnet_only: bool,
}

fn main() -> ExitCode {
    match run(Cli::parse()) {
        Ok(()) => ExitCode::SUCCESS,
        Err(error) => {
            eprintln!("error: {error}");
            ExitCode::FAILURE
        }
    }
}

fn run(cli: Cli) -> CliResult<()> {
    match cli.command {
        CliCommand::Sheets { file } => list_sheets(&file),
        CliCommand::Query(arguments) => query(arguments),
        CliCommand::WriteDemo { output } => write_demo(&output),
        CliCommand::Parity(arguments) => parity(arguments),
    }
}

fn list_sheets(file: &Path) -> CliResult<()> {
    for (index, sheet_name) in MiniExcel::get_sheet_names(file)?.into_iter().enumerate() {
        println!("{}\t{}", index + 1, sheet_name);
    }
    Ok(())
}

fn query(arguments: QueryArgs) -> CliResult<()> {
    let mut options = ReadOptions::new()
        .with_start_cell(arguments.start_cell)
        .with_header_mode(if arguments.header { HeaderMode::FirstRow } else { HeaderMode::None })
        .with_ignore_empty_rows(arguments.ignore_empty_rows);
    if let Some(sheet_name) = arguments.sheet {
        options = options.with_sheet_name(sheet_name);
    }
    if let Some(end_cell) = arguments.end_cell {
        options = options.with_end_cell(end_cell);
    }

    let rows = MiniExcel::query_with_options(&arguments.file, &options)?;
    match arguments.format {
        OutputFormat::Jsonl => print_json_lines(rows, arguments.limit),
        OutputFormat::Json => {
            let rows = collect_rows(rows, arguments.limit)?;
            println!("{}", serde_json::to_string_pretty(&rows_to_json(&rows))?);
            Ok(())
        }
        OutputFormat::Table => {
            let rows = collect_rows(rows, arguments.limit)?;
            print_table(&rows);
            Ok(())
        }
    }
}

fn collect_rows(
    mut rows: impl Iterator<Item = miniexcel::Result<DynamicRow>>,
    limit: usize,
) -> CliResult<Vec<DynamicRow>> {
    let mut output = Vec::new();
    while limit == 0 || output.len() < limit {
        let Some(row) = rows.next() else {
            break;
        };
        output.push(row?);
    }
    Ok(output)
}

fn print_json_lines(
    mut rows: impl Iterator<Item = miniexcel::Result<DynamicRow>>,
    limit: usize,
) -> CliResult<()> {
    let stdout = io::stdout();
    let mut output = stdout.lock();
    let mut count = 0;
    while limit == 0 || count < limit {
        let Some(row) = rows.next() else {
            break;
        };
        serde_json::to_writer(&mut output, &row_to_json(&row?))?;
        output.write_all(b"\n")?;
        count += 1;
    }
    Ok(())
}

fn rows_to_json(rows: &[DynamicRow]) -> Value {
    Value::Array(rows.iter().map(row_to_json).collect())
}

fn row_to_json(row: &DynamicRow) -> Value {
    let values = row
        .iter()
        .map(|(column, value)| (column.clone(), cell_to_json(value)))
        .collect::<Map<_, _>>();
    Value::Object(values)
}

fn cell_to_json(value: &CellValue) -> Value {
    match value {
        CellValue::Empty => Value::Null,
        CellValue::Bool(value) => Value::Bool(*value),
        CellValue::Int(value) => Value::Number(Number::from(*value)),
        CellValue::Float(value) => Number::from_f64(*value)
            .map(Value::Number)
            .unwrap_or_else(|| Value::String(value.to_string())),
        CellValue::String(value) | CellValue::Error(value) => Value::String(value.clone()),
        CellValue::Date(value) => Value::String(value.format("%Y-%m-%d").to_string()),
        CellValue::Time(value) => Value::String(value.format("%H:%M:%S%.f").to_string()),
        CellValue::DateTime(value) => {
            Value::String(value.format("%Y-%m-%dT%H:%M:%S%.f").to_string())
        }
        CellValue::Duration(value) => Value::String(format!("{}ms", value.num_milliseconds())),
    }
}

fn print_table(rows: &[DynamicRow]) {
    if rows.is_empty() {
        println!("No rows.");
        return;
    }

    let mut seen = HashSet::new();
    let mut columns = Vec::new();
    for row in rows {
        for column in row.keys() {
            if seen.insert(column.clone()) {
                columns.push(column.clone());
            }
        }
    }

    println!("| {} |", columns.join(" | "));
    println!("| {} |", columns.iter().map(|_| "---").collect::<Vec<_>>().join(" | "));
    for row in rows {
        let values = columns
            .iter()
            .map(|column| row.get(column).map_or_else(String::new, display_cell))
            .collect::<Vec<_>>();
        println!("| {} |", values.join(" | "));
    }
}

fn display_cell(value: &CellValue) -> String {
    let value = match value {
        CellValue::Empty => String::new(),
        CellValue::Bool(value) => value.to_string(),
        CellValue::Int(value) => value.to_string(),
        CellValue::Float(value) => value.to_string(),
        CellValue::String(value) => value.clone(),
        CellValue::Date(value) => value.format("%Y-%m-%d").to_string(),
        CellValue::Time(value) => value.format("%H:%M:%S%.f").to_string(),
        CellValue::DateTime(value) => value.format("%Y-%m-%dT%H:%M:%S%.f").to_string(),
        CellValue::Duration(value) => format!("{}ms", value.num_milliseconds()),
        CellValue::Error(value) => value.clone(),
    };
    value.replace('|', "\\|").replace('\r', "\\r").replace('\n', "\\n").replace('\t', "\\t")
}

fn write_demo(output: &Path) -> CliResult<()> {
    if let Some(parent) = output.parent().filter(|parent| !parent.as_os_str().is_empty()) {
        fs::create_dir_all(parent)?;
    }

    let mut first = DynamicRow::new();
    first.insert("Name".to_owned(), CellValue::String("MiniExcel".to_owned()));
    first.insert("Version".to_owned(), CellValue::Int(2));
    first.insert("Active".to_owned(), CellValue::Bool(true));
    first.insert(
        "ReleasedOn".to_owned(),
        CellValue::Date(NaiveDate::from_ymd_opt(2026, 8, 13).expect("valid demo date")),
    );

    let mut second = DynamicRow::new();
    second.insert("Name".to_owned(), CellValue::String("MiniExcel Rust".to_owned()));
    second.insert("Version".to_owned(), CellValue::Int(1));
    second.insert("Active".to_owned(), CellValue::Bool(true));

    let rows = [first, second];
    let write_options = WriteOptions::new().with_sheet_name("Demo");
    MiniExcel::save_as_with_options(output, &rows, &write_options)?;

    let read_options =
        ReadOptions::new().with_sheet_name("Demo").with_header_mode(HeaderMode::FirstRow);
    let row_count = MiniExcel::query_with_options(output, &read_options)?
        .try_fold(0usize, |count, row| row.map(|_| count + 1))?;
    println!("Wrote and read back {row_count} rows: {}", output.display());
    Ok(())
}

fn parity(arguments: ParityArgs) -> CliResult<()> {
    let repo_root = find_repo_root(arguments.repo_root.as_deref())?;
    println!("Repository: {}", repo_root.display());

    if !arguments.dotnet_only {
        let mut command = ProcessCommand::new("cargo");
        command.current_dir(&repo_root).args([
            "+1.85.0",
            "test",
            "--manifest-path",
            "rust/Cargo.toml",
            "-p",
            "miniexcel",
            "--test",
            "parity_contract",
            "--locked",
        ]);
        run_process("Rust parity", &mut command)?;
    }

    if !arguments.rust_only {
        let mut command = ProcessCommand::new("dotnet");
        command.current_dir(&repo_root).args([
            "test",
            "tests/MiniExcel.OpenXml.Tests/MiniExcel.OpenXml.Tests.csproj",
            "--framework",
            &arguments.framework,
            "--filter",
            "FullyQualifiedName~RustParityContractTests",
            "--verbosity",
            "minimal",
        ]);
        run_process(".NET parity", &mut command)?;
    }

    println!("Parity checks passed.");
    Ok(())
}

fn find_repo_root(explicit: Option<&Path>) -> CliResult<PathBuf> {
    if let Some(root) = explicit {
        return validate_repo_root(root.to_owned());
    }

    let current = env::current_dir()?;
    for candidate in current.ancestors() {
        if is_repo_root(candidate) {
            return Ok(candidate.to_owned());
        }
    }

    let compiled_root = PathBuf::from(env!("CARGO_MANIFEST_DIR")).join("..").join("..");
    validate_repo_root(compiled_root)
}

fn validate_repo_root(root: PathBuf) -> CliResult<PathBuf> {
    if is_repo_root(&root) {
        Ok(root)
    } else {
        Err(io::Error::new(
            io::ErrorKind::NotFound,
            format!("'{}' is not a MiniExcel repository root", root.display()),
        )
        .into())
    }
}

fn is_repo_root(path: &Path) -> bool {
    path.join("rust/Cargo.toml").is_file()
        && path.join("tests/MiniExcel.OpenXml.Tests/MiniExcel.OpenXml.Tests.csproj").is_file()
}

fn run_process(label: &str, command: &mut ProcessCommand) -> CliResult<()> {
    println!("\n== {label} ==");
    let status = command.status()?;
    if status.success() {
        Ok(())
    } else {
        Err(io::Error::other(format!("{label} failed with {status}")).into())
    }
}

#[cfg(test)]
mod tests {
    use std::cell::Cell;

    use miniexcel::DynamicRow;

    use super::{Cli, CliCommand, OutputFormat, collect_rows};
    use clap::Parser;

    #[test]
    fn parses_query_options() {
        let cli = Cli::try_parse_from([
            "miniexcel",
            "query",
            "book.xlsx",
            "--sheet",
            "Data",
            "--header",
            "--start-cell",
            "B2",
            "--limit",
            "5",
            "--format",
            "jsonl",
        ])
        .expect("parse query arguments");

        let CliCommand::Query(arguments) = cli.command else {
            panic!("expected query command");
        };
        assert_eq!(arguments.sheet.as_deref(), Some("Data"));
        assert!(arguments.header);
        assert_eq!(arguments.start_cell.to_string(), "B2");
        assert_eq!(arguments.limit, 5);
        assert!(matches!(arguments.format, OutputFormat::Jsonl));
    }

    #[test]
    fn row_limit_does_not_overread_the_stream() {
        let pulls = Cell::new(0usize);
        let rows = std::iter::from_fn(|| {
            pulls.set(pulls.get() + 1);
            Some(Ok(DynamicRow::new()))
        });

        let output = collect_rows(rows, 2).expect("collect limited rows");

        assert_eq!(output.len(), 2);
        assert_eq!(pulls.get(), 2);
    }
}
