use std::collections::HashMap;
use std::fs::File;
use std::io::{BufReader, Read, Seek};
use std::path::{Path, PathBuf};
use std::sync::mpsc::{self, Receiver, SyncSender};
use std::sync::{
    Arc,
    atomic::{AtomicBool, Ordering},
};
use std::thread::{self, JoinHandle};

use calamine::{CellErrorType, Data, DataType, ExcelDateTime, ExcelDateTimeType};
use quick_xml::Reader;
use quick_xml::encoding::Decoder;
use quick_xml::events::{BytesRef, BytesStart, Event};
use zip::ZipArchive;
use zip::result::ZipError;

use crate::reader::SelectedRow;
use crate::{Error, ReadOptions, Result};

const ROW_BUFFER_SIZE: usize = 8;

pub(super) struct StreamingRawRows {
    receiver: Option<Receiver<Result<SelectedRow>>>,
    worker: Option<JoinHandle<()>>,
    sheet_name: String,
    cancelled: Arc<AtomicBool>,
}

impl StreamingRawRows {
    pub(super) fn open(path: impl AsRef<Path>, options: &ReadOptions) -> Result<Self> {
        let path = path.as_ref().to_owned();
        let file = File::open(&path)?;
        let (ready_sender, ready_receiver) = mpsc::sync_channel(0);
        let (row_sender, row_receiver) = mpsc::sync_channel(ROW_BUFFER_SIZE);
        let cancelled = Arc::new(AtomicBool::new(false));
        let worker_cancelled = Arc::clone(&cancelled);
        let options = options.clone();
        let worker =
            thread::Builder::new().name("miniexcel-xlsx-stream".to_owned()).spawn(move || {
                worker_main(
                    path,
                    BufReader::new(file),
                    options,
                    worker_cancelled,
                    ready_sender,
                    row_sender,
                )
            })?;

        let ready = match ready_receiver.recv() {
            Ok(ready) => ready,
            Err(_) => Err(Error::stream(
                "the XLSX streaming worker stopped during initialization".to_owned(),
            )),
        };
        match ready {
            Ok(sheet_name) => Ok(Self {
                receiver: Some(row_receiver),
                worker: Some(worker),
                sheet_name,
                cancelled,
            }),
            Err(error) => {
                drop(row_receiver);
                let _ = worker.join();
                Err(error)
            }
        }
    }

    pub(super) fn sheet_name(&self) -> &str {
        &self.sheet_name
    }
}

impl Iterator for StreamingRawRows {
    type Item = Result<SelectedRow>;

    fn next(&mut self) -> Option<Self::Item> {
        self.receiver.as_ref()?.recv().ok()
    }
}

impl Drop for StreamingRawRows {
    fn drop(&mut self) {
        self.cancelled.store(true, Ordering::Relaxed);
        self.receiver.take();
        if let Some(worker) = self.worker.take() {
            let _ = worker.join();
        }
    }
}

struct WorkbookContext {
    sheet_name: String,
    sheet_path: String,
    shared_strings: Vec<String>,
    styles: Vec<CellFormat>,
    is_1904: bool,
}

#[derive(Clone, Copy, Debug, Default, Eq, PartialEq)]
enum CellFormat {
    #[default]
    Other,
    DateTime,
    TimeDelta,
}

#[derive(Default)]
struct WorkbookInfo {
    sheets: Vec<SheetInfo>,
    is_1904: bool,
}

struct SheetInfo {
    name: String,
    relationship_id: String,
}

fn worker_main<R>(
    path: PathBuf,
    reader: R,
    options: ReadOptions,
    cancelled: Arc<AtomicBool>,
    ready_sender: SyncSender<Result<String>>,
    row_sender: SyncSender<Result<SelectedRow>>,
) where
    R: Read + Seek,
{
    let mut archive = match ZipArchive::new(reader) {
        Ok(archive) => archive,
        Err(error) => {
            let _ = ready_sender
                .send(Err(stream_error(format!("cannot open '{}':", path.display()), error)));
            return;
        }
    };
    let context = match prepare_workbook(&mut archive, &options) {
        Ok(context) => context,
        Err(error) => {
            let _ = ready_sender.send(Err(error));
            return;
        }
    };
    if ready_sender.send(Ok(context.sheet_name.clone())).is_err() {
        return;
    }
    let extent = match scan_worksheet_extent(&mut archive, &context.sheet_path, &cancelled) {
        Ok(extent) => extent,
        Err(error) => {
            let _ = row_sender.send(Err(error));
            return;
        }
    };
    if cancelled.load(Ordering::Relaxed) {
        return;
    }
    if let Err(error) =
        stream_worksheet(&mut archive, context, extent, &options, &cancelled, &row_sender)
    {
        let _ = row_sender.send(Err(error));
    }
}

fn prepare_workbook<R>(
    archive: &mut ZipArchive<R>,
    options: &ReadOptions,
) -> Result<WorkbookContext>
where
    R: Read + Seek,
{
    let workbook = read_workbook_info(archive)?;
    let sheet = match options.sheet_name() {
        Some(sheet_name) => workbook
            .sheets
            .iter()
            .find(|sheet| sheet.name == sheet_name)
            .ok_or_else(|| Error::sheet_not_found(sheet_name))?,
        None => workbook.sheets.first().ok_or_else(Error::no_worksheets)?,
    };
    let sheet_path = read_relationship_target(archive, &sheet.relationship_id)?;
    let shared_strings = read_shared_strings(archive)?;
    let styles = read_styles(archive)?;
    Ok(WorkbookContext {
        sheet_name: sheet.name.clone(),
        sheet_path,
        shared_strings,
        styles,
        is_1904: workbook.is_1904,
    })
}

fn read_workbook_info<R>(archive: &mut ZipArchive<R>) -> Result<WorkbookInfo>
where
    R: Read + Seek,
{
    let file = archive
        .by_name("xl/workbook.xml")
        .map_err(|error| stream_error("cannot read xl/workbook.xml:", error))?;
    let mut xml = Reader::from_reader(BufReader::new(file));
    let mut buffer = Vec::new();
    let mut workbook = WorkbookInfo::default();
    loop {
        match xml
            .read_event_into(&mut buffer)
            .map_err(|error| stream_error("invalid xl/workbook.xml:", error))?
        {
            Event::Start(event) | Event::Empty(event)
                if is_name(event.name().as_ref(), b"workbookPr") =>
            {
                workbook.is_1904 = attribute(&event, xml.decoder(), b"date1904")?
                    .is_some_and(|value| value == "1" || value.eq_ignore_ascii_case("true"));
            }
            Event::Start(event) | Event::Empty(event)
                if is_name(event.name().as_ref(), b"sheet") =>
            {
                let Some(name) = attribute(&event, xml.decoder(), b"name")? else {
                    buffer.clear();
                    continue;
                };
                let Some(relationship_id) = attribute(&event, xml.decoder(), b"id")? else {
                    buffer.clear();
                    continue;
                };
                workbook.sheets.push(SheetInfo { name, relationship_id });
            }
            Event::Eof => break,
            _ => {}
        }
        buffer.clear();
    }
    Ok(workbook)
}

fn read_relationship_target<R>(archive: &mut ZipArchive<R>, relationship_id: &str) -> Result<String>
where
    R: Read + Seek,
{
    let file = archive
        .by_name("xl/_rels/workbook.xml.rels")
        .map_err(|error| stream_error("cannot read workbook relationships:", error))?;
    let mut xml = Reader::from_reader(BufReader::new(file));
    let mut buffer = Vec::new();
    loop {
        match xml
            .read_event_into(&mut buffer)
            .map_err(|error| stream_error("invalid workbook relationships:", error))?
        {
            Event::Start(event) | Event::Empty(event)
                if is_name(event.name().as_ref(), b"Relationship") =>
            {
                if attribute(&event, xml.decoder(), b"Id")?.as_deref() == Some(relationship_id) {
                    let target = attribute(&event, xml.decoder(), b"Target")?.ok_or_else(|| {
                        Error::stream(format!(
                            "worksheet relationship '{relationship_id}' has no target"
                        ))
                    })?;
                    return Ok(normalize_zip_path(&target));
                }
            }
            Event::Eof => break,
            _ => {}
        }
        buffer.clear();
    }
    Err(Error::stream(format!("worksheet relationship '{relationship_id}' was not found")))
}

fn read_shared_strings<R>(archive: &mut ZipArchive<R>) -> Result<Vec<String>>
where
    R: Read + Seek,
{
    let file = match archive.by_name("xl/sharedStrings.xml") {
        Ok(file) => file,
        Err(ZipError::FileNotFound) => return Ok(Vec::new()),
        Err(error) => return Err(stream_error("cannot read shared strings:", error)),
    };
    let mut xml = Reader::from_reader(BufReader::new(file));
    let mut buffer = Vec::new();
    let mut strings = Vec::new();
    let mut current = String::new();
    let mut in_item = false;
    let mut in_text = false;
    loop {
        match xml
            .read_event_into(&mut buffer)
            .map_err(|error| stream_error("invalid shared strings XML:", error))?
        {
            Event::Start(event) if is_name(event.name().as_ref(), b"si") => {
                current.clear();
                in_item = true;
            }
            Event::Empty(event) if is_name(event.name().as_ref(), b"si") => {
                strings.push(String::new());
            }
            Event::Start(event) if in_item && is_name(event.name().as_ref(), b"t") => {
                in_text = true;
            }
            Event::Text(text) if in_text => append_text(&text, &mut current)?,
            Event::GeneralRef(reference) if in_text => append_reference(&reference, &mut current)?,
            Event::End(event) if is_name(event.name().as_ref(), b"t") => in_text = false,
            Event::End(event) if is_name(event.name().as_ref(), b"si") => {
                strings.push(decode_excel_escapes(&current));
                in_item = false;
            }
            Event::Eof => break,
            _ => {}
        }
        buffer.clear();
    }
    Ok(strings)
}

fn read_styles<R>(archive: &mut ZipArchive<R>) -> Result<Vec<CellFormat>>
where
    R: Read + Seek,
{
    let file = match archive.by_name("xl/styles.xml") {
        Ok(file) => file,
        Err(ZipError::FileNotFound) => return Ok(Vec::new()),
        Err(error) => return Err(stream_error("cannot read styles:", error)),
    };
    let mut xml = Reader::from_reader(BufReader::new(file));
    let mut buffer = Vec::new();
    let mut custom_formats = HashMap::new();
    let mut styles = Vec::new();
    let mut in_cell_formats = false;
    loop {
        match xml
            .read_event_into(&mut buffer)
            .map_err(|error| stream_error("invalid styles XML:", error))?
        {
            Event::Start(event) if is_name(event.name().as_ref(), b"cellXfs") => {
                in_cell_formats = true;
            }
            Event::End(event) if is_name(event.name().as_ref(), b"cellXfs") => {
                in_cell_formats = false;
            }
            Event::Start(event) | Event::Empty(event)
                if is_name(event.name().as_ref(), b"numFmt") =>
            {
                let id = attribute(&event, xml.decoder(), b"numFmtId")?
                    .and_then(|value| value.parse::<u32>().ok());
                let format = attribute(&event, xml.decoder(), b"formatCode")?;
                if let (Some(id), Some(format)) = (id, format) {
                    custom_formats.insert(id, classify_custom_format(&format));
                }
            }
            Event::Start(event) | Event::Empty(event)
                if in_cell_formats && is_name(event.name().as_ref(), b"xf") =>
            {
                let id = attribute(&event, xml.decoder(), b"numFmtId")?
                    .and_then(|value| value.parse::<u32>().ok())
                    .unwrap_or(0);
                styles.push(custom_formats.get(&id).copied().unwrap_or_else(|| builtin_format(id)));
            }
            Event::Eof => break,
            _ => {}
        }
        buffer.clear();
    }
    Ok(styles)
}

#[derive(Clone, Copy, Default)]
enum CellKind {
    #[default]
    Number,
    SharedString,
    InlineString,
    Boolean,
    Error,
    String,
    IsoDate,
}

#[derive(Clone, Copy)]
enum Capture {
    Value,
    InlineText,
}

struct CellState {
    column: usize,
    style: usize,
    kind: CellKind,
    value: String,
    inline_text: String,
    capture: Option<Capture>,
}

struct RowState {
    excel_row: usize,
    cells: Vec<(usize, Data)>,
    next_column: usize,
}

#[derive(Clone, Copy, Default)]
struct WorksheetExtent {
    end_row: Option<usize>,
    end_column: Option<usize>,
}

fn scan_worksheet_extent<R>(
    archive: &mut ZipArchive<R>,
    sheet_path: &str,
    cancelled: &AtomicBool,
) -> Result<WorksheetExtent>
where
    R: Read + Seek,
{
    let file = archive
        .by_name(sheet_path)
        .map_err(|error| stream_error("cannot scan worksheet XML:", error))?;
    let mut xml = Reader::from_reader(BufReader::new(file));
    let mut buffer = Vec::new();
    let mut extent = WorksheetExtent::default();
    let mut current_row = None;
    let mut last_declared_row = None;
    let mut next_column = 0;
    let mut in_sheet_data = false;

    loop {
        if cancelled.load(Ordering::Relaxed) {
            return Ok(extent);
        }
        match xml
            .read_event_into(&mut buffer)
            .map_err(|error| stream_error("invalid worksheet XML during scan:", error))?
        {
            Event::Start(event) if is_name(event.name().as_ref(), b"sheetData") => {
                in_sheet_data = true;
            }
            Event::End(event) if is_name(event.name().as_ref(), b"sheetData") => break,
            Event::Start(event) | Event::Empty(event)
                if in_sheet_data && is_name(event.name().as_ref(), b"row") =>
            {
                let row = row_index(&event, xml.decoder(), last_declared_row)?;
                current_row = Some(row);
                last_declared_row = Some(row);
                next_column = 0;
            }
            Event::Start(event) | Event::Empty(event)
                if current_row.is_some() && is_name(event.name().as_ref(), b"c") =>
            {
                let column = attribute(&event, xml.decoder(), b"r")?
                    .and_then(|reference| parse_column(&reference))
                    .unwrap_or(next_column);
                next_column = column.saturating_add(1);
                extent.end_row = current_row;
                extent.end_column = Some(extent.end_column.map_or(column, |end| end.max(column)));
            }
            Event::End(event) if is_name(event.name().as_ref(), b"row") => {
                current_row = None;
            }
            Event::Eof => break,
            _ => {}
        }
        buffer.clear();
    }
    Ok(extent)
}

fn stream_worksheet<R>(
    archive: &mut ZipArchive<R>,
    context: WorkbookContext,
    extent: WorksheetExtent,
    options: &ReadOptions,
    cancelled: &AtomicBool,
    sender: &SyncSender<Result<SelectedRow>>,
) -> Result<()>
where
    R: Read + Seek,
{
    let file = archive
        .by_name(&context.sheet_path)
        .map_err(|error| stream_error("cannot read worksheet XML:", error))?;
    let mut xml = Reader::from_reader(BufReader::new(file));
    let mut buffer = Vec::new();
    let mut current_row = None;
    let mut current_cell = None;
    let mut last_declared_row = None;
    let mut next_output_row = options.start_cell().row();
    let mut in_sheet_data = false;

    loop {
        if cancelled.load(Ordering::Relaxed) {
            return Ok(());
        }
        match xml
            .read_event_into(&mut buffer)
            .map_err(|error| stream_error("invalid worksheet XML:", error))?
        {
            Event::Start(event) if is_name(event.name().as_ref(), b"sheetData") => {
                in_sheet_data = true;
            }
            Event::End(event) if is_name(event.name().as_ref(), b"sheetData") => break,
            Event::Start(event) if in_sheet_data && is_name(event.name().as_ref(), b"row") => {
                let excel_row = row_index(&event, xml.decoder(), last_declared_row)?;
                if extent.end_row.is_none_or(|end_row| excel_row > end_row) {
                    break;
                }
                if !emit_missing_rows(
                    &mut next_output_row,
                    excel_row,
                    extent.end_column,
                    options,
                    sender,
                ) {
                    return Ok(());
                }
                last_declared_row = Some(excel_row);
                current_row = Some(RowState { excel_row, cells: Vec::new(), next_column: 0 });
            }
            Event::Empty(event) if in_sheet_data && is_name(event.name().as_ref(), b"row") => {
                let excel_row = row_index(&event, xml.decoder(), last_declared_row)?;
                if extent.end_row.is_none_or(|end_row| excel_row > end_row) {
                    break;
                }
                if !emit_missing_rows(
                    &mut next_output_row,
                    excel_row,
                    extent.end_column,
                    options,
                    sender,
                ) {
                    return Ok(());
                }
                last_declared_row = Some(excel_row);
                let row = RowState { excel_row, cells: Vec::new(), next_column: 0 };
                if !emit_row(row, &mut next_output_row, extent.end_column, options, sender) {
                    return Ok(());
                }
            }
            Event::Start(event)
                if current_row.is_some() && is_name(event.name().as_ref(), b"c") =>
            {
                current_cell = Some(start_cell(
                    &event,
                    xml.decoder(),
                    current_row.as_mut().expect("row checked above"),
                )?);
            }
            Event::Empty(event)
                if current_row.is_some() && is_name(event.name().as_ref(), b"c") =>
            {
                let cell = start_cell(
                    &event,
                    xml.decoder(),
                    current_row.as_mut().expect("row checked above"),
                )?;
                finish_cell(cell, current_row.as_mut().expect("row checked above"), &context)?;
            }
            Event::Start(event)
                if current_cell.is_some() && is_name(event.name().as_ref(), b"v") =>
            {
                current_cell.as_mut().expect("cell checked above").capture = Some(Capture::Value);
            }
            Event::Start(event)
                if current_cell.is_some() && is_name(event.name().as_ref(), b"t") =>
            {
                current_cell.as_mut().expect("cell checked above").capture =
                    Some(Capture::InlineText);
            }
            Event::Text(text) if current_cell.is_some() => {
                append_cell_text(&text, current_cell.as_mut().expect("cell checked above"))?;
            }
            Event::GeneralRef(reference) if current_cell.is_some() => {
                append_cell_reference(
                    &reference,
                    current_cell.as_mut().expect("cell checked above"),
                )?;
            }
            Event::End(event)
                if current_cell.is_some()
                    && (is_name(event.name().as_ref(), b"v")
                        || is_name(event.name().as_ref(), b"t")) =>
            {
                current_cell.as_mut().expect("cell checked above").capture = None;
            }
            Event::End(event) if is_name(event.name().as_ref(), b"c") => {
                if let (Some(cell), Some(row)) = (current_cell.take(), current_row.as_mut()) {
                    finish_cell(cell, row, &context)?;
                }
            }
            Event::End(event) if is_name(event.name().as_ref(), b"row") => {
                if let Some(row) = current_row.take() {
                    if !emit_row(row, &mut next_output_row, extent.end_column, options, sender) {
                        return Ok(());
                    }
                }
            }
            Event::Eof => break,
            _ => {}
        }
        buffer.clear();
    }
    Ok(())
}

fn start_cell(event: &BytesStart<'_>, decoder: Decoder, row: &mut RowState) -> Result<CellState> {
    let column = attribute(event, decoder, b"r")?
        .and_then(|reference| parse_column(&reference))
        .unwrap_or(row.next_column);
    row.next_column = column.saturating_add(1);
    let style =
        attribute(event, decoder, b"s")?.and_then(|value| value.parse::<usize>().ok()).unwrap_or(0);
    let kind = match attribute(event, decoder, b"t")?.as_deref() {
        Some("s") => CellKind::SharedString,
        Some("inlineStr") => CellKind::InlineString,
        Some("b") => CellKind::Boolean,
        Some("e") => CellKind::Error,
        Some("str") => CellKind::String,
        Some("d") => CellKind::IsoDate,
        _ => CellKind::Number,
    };
    Ok(CellState {
        column,
        style,
        kind,
        value: String::new(),
        inline_text: String::new(),
        capture: None,
    })
}

fn finish_cell(cell: CellState, row: &mut RowState, context: &WorkbookContext) -> Result<()> {
    let data = match cell.kind {
        CellKind::SharedString => {
            if cell.value.is_empty() {
                Data::Empty
            } else {
                let index = cell
                    .value
                    .parse::<usize>()
                    .map_err(|error| stream_error("invalid shared string index:", error))?;
                Data::String(context.shared_strings.get(index).cloned().ok_or_else(|| {
                    Error::stream(format!("shared string index {index} is out of range"))
                })?)
            }
        }
        CellKind::InlineString => Data::String(decode_excel_escapes(&cell.inline_text)),
        CellKind::Boolean => {
            Data::Bool(cell.value == "1" || cell.value.eq_ignore_ascii_case("true"))
        }
        CellKind::Error => Data::Error(parse_cell_error(&cell.value)?),
        CellKind::String => Data::String(decode_excel_escapes(&cell.value)),
        CellKind::IsoDate => Data::DateTimeIso(cell.value),
        CellKind::Number if cell.value.is_empty() => Data::Empty,
        CellKind::Number => {
            let value = cell
                .value
                .parse::<f64>()
                .map_err(|error| stream_error("invalid numeric cell value:", error))?;
            match context.styles.get(cell.style).copied().unwrap_or_default() {
                CellFormat::DateTime => Data::DateTime(ExcelDateTime::new(
                    value,
                    ExcelDateTimeType::DateTime,
                    context.is_1904,
                )),
                CellFormat::TimeDelta => Data::DateTime(ExcelDateTime::new(
                    value,
                    ExcelDateTimeType::TimeDelta,
                    context.is_1904,
                )),
                CellFormat::Other => Data::Float(value),
            }
        }
    };
    row.cells.push((cell.column, data));
    Ok(())
}

fn emit_missing_rows(
    next_output_row: &mut usize,
    target_row: usize,
    end_column: Option<usize>,
    options: &ReadOptions,
    sender: &SyncSender<Result<SelectedRow>>,
) -> bool {
    while *next_output_row < target_row {
        let row = RowState { excel_row: *next_output_row, cells: Vec::new(), next_column: 0 };
        if !emit_row(row, next_output_row, end_column, options, sender) {
            return false;
        }
    }
    true
}

fn emit_row(
    row: RowState,
    next_output_row: &mut usize,
    end_column: Option<usize>,
    options: &ReadOptions,
    sender: &SyncSender<Result<SelectedRow>>,
) -> bool {
    *next_output_row = (*next_output_row).max(row.excel_row.saturating_add(1));
    if row.excel_row < options.start_cell().row() {
        return true;
    }
    let start_column = options.start_cell().column();
    let mut width = end_column
        .filter(|column| *column >= start_column)
        .map_or(0, |column| column - start_column + 1);
    if let Some(max_column) = row.cells.iter().map(|(column, _)| *column).max() {
        if max_column >= start_column {
            width = width.max(max_column - start_column + 1);
        }
    }
    let mut values = vec![Data::Empty; width];
    for (column, value) in row.cells {
        if column >= start_column {
            values[column - start_column] = value;
        }
    }
    if options.ignore_empty_rows() && values.iter().all(DataType::is_empty) {
        return true;
    }
    sender.send(Ok(SelectedRow { excel_row: row.excel_row, values })).is_ok()
}

fn row_index(event: &BytesStart<'_>, decoder: Decoder, previous: Option<usize>) -> Result<usize> {
    Ok(attribute(event, decoder, b"r")?
        .and_then(|value| value.parse::<usize>().ok())
        .and_then(|value| value.checked_sub(1))
        .unwrap_or_else(|| previous.map_or(0, |row| row.saturating_add(1))))
}

fn append_cell_text(text: &quick_xml::events::BytesText<'_>, cell: &mut CellState) -> Result<()> {
    match cell.capture {
        Some(Capture::Value) => append_text(text, &mut cell.value),
        Some(Capture::InlineText) => append_text(text, &mut cell.inline_text),
        None => Ok(()),
    }
}

fn append_cell_reference(reference: &BytesRef<'_>, cell: &mut CellState) -> Result<()> {
    match cell.capture {
        Some(Capture::Value) => append_reference(reference, &mut cell.value),
        Some(Capture::InlineText) => append_reference(reference, &mut cell.inline_text),
        None => Ok(()),
    }
}

fn append_text(text: &quick_xml::events::BytesText<'_>, target: &mut String) -> Result<()> {
    let text = text.xml10_content().map_err(|error| stream_error("invalid XML text:", error))?;
    target.push_str(&text);
    Ok(())
}

fn append_reference(reference: &BytesRef<'_>, target: &mut String) -> Result<()> {
    let decoded =
        reference.decode().map_err(|error| stream_error("invalid XML reference:", error))?;
    match decoded.as_ref() {
        "lt" => target.push('<'),
        "gt" => target.push('>'),
        "amp" => target.push('&'),
        "quot" => target.push('"'),
        "apos" => target.push('\''),
        _ => {
            if let Some(value) = reference
                .resolve_char_ref()
                .map_err(|error| stream_error("invalid XML character reference:", error))?
            {
                target.push(value);
            } else {
                return Err(Error::stream(format!("unrecognized XML entity '&{decoded};'")));
            }
        }
    }
    Ok(())
}

fn attribute(event: &BytesStart<'_>, decoder: Decoder, name: &[u8]) -> Result<Option<String>> {
    for attribute in event.attributes().with_checks(false) {
        let attribute = attribute.map_err(|error| stream_error("invalid XML attribute:", error))?;
        if is_name(attribute.key.as_ref(), name) {
            return attribute
                .decode_and_unescape_value(decoder)
                .map(|value| Some(value.into_owned()))
                .map_err(|error| stream_error("invalid XML attribute value:", error));
        }
    }
    Ok(None)
}

fn is_name(actual: &[u8], expected: &[u8]) -> bool {
    actual.rsplit(|byte| *byte == b':').next() == Some(expected)
}

fn normalize_zip_path(target: &str) -> String {
    let target = target.replace('\\', "/");
    let path = if target.starts_with('/') {
        target.trim_start_matches('/').to_owned()
    } else if target.starts_with("xl/") {
        target
    } else {
        format!("xl/{target}")
    };
    let mut parts = Vec::new();
    for part in path.split('/') {
        match part {
            "" | "." => {}
            ".." => {
                parts.pop();
            }
            _ => parts.push(part),
        }
    }
    parts.join("/")
}

fn parse_column(reference: &str) -> Option<usize> {
    let mut column = 0usize;
    let mut found = false;
    for byte in reference.bytes() {
        if byte == b'$' && !found {
            continue;
        }
        if !byte.is_ascii_alphabetic() {
            break;
        }
        found = true;
        column = column
            .checked_mul(26)?
            .checked_add(usize::from(byte.to_ascii_uppercase() - b'A' + 1))?;
    }
    found.then(|| column - 1)
}

fn parse_cell_error(value: &str) -> Result<CellErrorType> {
    match value {
        "#DIV/0!" => Ok(CellErrorType::Div0),
        "#N/A" => Ok(CellErrorType::NA),
        "#NAME?" => Ok(CellErrorType::Name),
        "#NULL!" => Ok(CellErrorType::Null),
        "#NUM!" => Ok(CellErrorType::Num),
        "#REF!" => Ok(CellErrorType::Ref),
        "#VALUE!" => Ok(CellErrorType::Value),
        "#DATA!" | "#GETTING_DATA" => Ok(CellErrorType::GettingData),
        _ => Err(Error::stream(format!("unknown Excel cell error '{value}'"))),
    }
}

fn builtin_format(id: u32) -> CellFormat {
    match id {
        14..=22 | 45 | 47 => CellFormat::DateTime,
        46 => CellFormat::TimeDelta,
        _ => CellFormat::Other,
    }
}

fn classify_custom_format(format: &str) -> CellFormat {
    let characters = format.as_bytes();
    let mut index = 0;
    while index < characters.len() {
        match characters[index] {
            b';' => break,
            b'\\' | b'_' | b'*' => index = index.saturating_add(2),
            b'"' => {
                index += 1;
                while index < characters.len() && characters[index] != b'"' {
                    index += 1;
                }
                index = index.saturating_add(1);
            }
            b'[' => {
                let start = index + 1;
                index = start;
                while index < characters.len() && characters[index] != b']' {
                    index += 1;
                }
                let token = &format[start..index];
                if !token.is_empty()
                    && token
                        .bytes()
                        .all(|byte| matches!(byte.to_ascii_lowercase(), b'h' | b'm' | b's'))
                {
                    return CellFormat::TimeDelta;
                }
                index = index.saturating_add(1);
            }
            byte if matches!(byte.to_ascii_lowercase(), b'd' | b'm' | b'y' | b'h' | b's') => {
                return CellFormat::DateTime;
            }
            _ => index += 1,
        }
    }
    CellFormat::Other
}

fn decode_excel_escapes(value: &str) -> String {
    let bytes = value.as_bytes();
    let mut result = String::with_capacity(value.len());
    let mut index = 0;
    while index < bytes.len() {
        if index + 7 <= bytes.len()
            && bytes[index] == b'_'
            && matches!(bytes[index + 1], b'x' | b'X')
            && bytes[index + 6] == b'_'
        {
            if let Ok(code) = u16::from_str_radix(&value[index + 2..index + 6], 16) {
                if let Some(character) = char::from_u32(u32::from(code)) {
                    result.push(character);
                    index += 7;
                    continue;
                }
            }
        }
        let character = value[index..].chars().next().expect("index is in bounds");
        result.push(character);
        index += character.len_utf8();
    }
    result
}

fn stream_error(context: impl std::fmt::Display, error: impl std::fmt::Display) -> Error {
    Error::stream(format!("{context} {error}"))
}

#[cfg(test)]
mod tests {
    use super::{CellFormat, classify_custom_format};

    #[test]
    fn classifies_custom_excel_number_formats() {
        assert_eq!(classify_custom_format("yyyy-mm-dd"), CellFormat::DateTime);
        assert_eq!(classify_custom_format("h:mm AM/PM"), CellFormat::DateTime);
        assert_eq!(classify_custom_format("[h]:mm:ss"), CellFormat::TimeDelta);
        assert_eq!(classify_custom_format("0.00"), CellFormat::Other);
        assert_eq!(classify_custom_format("[Red][>=100]0.00"), CellFormat::Other);
        assert_eq!(classify_custom_format("\"days\" 0"), CellFormat::Other);
    }
}
