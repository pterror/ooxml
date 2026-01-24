//! Lazy/streaming API for parsing worksheets without full materialization.
//!
//! This module provides an alternative to the eager parsing API that avoids
//! allocating all rows and cells upfront. Instead, rows are parsed on-demand
//! as the iterator advances.
//!
//! # Use Cases
//!
//! - **Large files**: When memory is a concern and you don't need random access.
//! - **Early termination**: When you only need to read the first N rows.
//! - **Streaming processing**: When you process rows one-at-a-time and discard.
//!
//! # Example
//!
//! ```ignore
//! use ooxml_sml::lazy::LazyWorksheet;
//!
//! let xml = std::fs::read("xl/worksheets/sheet1.xml")?;
//! let worksheet = LazyWorksheet::new(&xml)?;
//!
//! for row in worksheet.rows() {
//!     let row = row?;
//!     println!("Row {}: {} cells", row.row_num(), row.cell_count());
//!     for cell in row.cells() {
//!         println!("  {:?}", cell);
//!     }
//! }
//! ```
//!
//! # Limitations
//!
//! - **One-shot iteration**: Once a row is consumed, you cannot go back.
//! - **No random access**: Cannot jump to row N directly.
//! - **No modification**: Read-only streaming.

use crate::generated_parsers::ParseError;
use quick_xml::Reader;
use quick_xml::events::{BytesStart, Event};
use std::io::{BufRead, Cursor};

/// A lazy worksheet that parses rows on-demand.
///
/// This provides streaming access to worksheet data without loading
/// all rows into memory at once.
pub struct LazyWorksheet<'a> {
    /// The raw XML data.
    xml: &'a [u8],
}

impl<'a> LazyWorksheet<'a> {
    /// Create a new lazy worksheet from raw XML bytes.
    ///
    /// This does not parse any rows immediately - parsing happens
    /// when you iterate over `rows()`.
    pub fn new(xml: &'a [u8]) -> Self {
        Self { xml }
    }

    /// Returns an iterator over rows, parsing each on-demand.
    ///
    /// Each call creates a fresh iterator from the beginning.
    /// Rows are parsed as you iterate - no data is loaded until needed.
    pub fn rows(&self) -> RowIterator<Cursor<&'a [u8]>> {
        RowIterator::new(Reader::from_reader(Cursor::new(self.xml)))
    }
}

/// Iterator over worksheet rows that parses on-demand.
pub struct RowIterator<R: BufRead> {
    reader: Reader<R>,
    buf: Vec<u8>,
    in_sheet_data: bool,
    finished: bool,
}

impl<R: BufRead> RowIterator<R> {
    fn new(reader: Reader<R>) -> Self {
        Self {
            reader,
            buf: Vec::new(),
            in_sheet_data: false,
            finished: false,
        }
    }
}

impl<R: BufRead> Iterator for RowIterator<R> {
    type Item = Result<LazyRow, ParseError>;

    fn next(&mut self) -> Option<Self::Item> {
        if self.finished {
            return None;
        }

        loop {
            self.buf.clear();
            match self.reader.read_event_into(&mut self.buf) {
                Ok(Event::Start(e)) => {
                    let name = e.name();
                    match name.as_ref() {
                        b"sheetData" => {
                            self.in_sheet_data = true;
                        }
                        b"row" if self.in_sheet_data => {
                            // Parse this row
                            match parse_lazy_row(&mut self.reader, &e) {
                                Ok(row) => return Some(Ok(row)),
                                Err(e) => return Some(Err(e)),
                            }
                        }
                        _ => {
                            // Skip other elements inside sheetData
                            if self.in_sheet_data
                                && let Err(e) = skip_element(&mut self.reader)
                            {
                                return Some(Err(e));
                            }
                        }
                    }
                }
                Ok(Event::Empty(e)) => {
                    let name = e.name();
                    if name.as_ref() == b"row" && self.in_sheet_data {
                        // Empty row (no cells)
                        match parse_lazy_row_from_empty(&e) {
                            Ok(row) => return Some(Ok(row)),
                            Err(e) => return Some(Err(e)),
                        }
                    }
                }
                Ok(Event::End(e)) => {
                    if e.name().as_ref() == b"sheetData" {
                        self.in_sheet_data = false;
                        self.finished = true;
                        return None;
                    }
                }
                Ok(Event::Eof) => {
                    self.finished = true;
                    return None;
                }
                Err(e) => return Some(Err(ParseError::Xml(e))),
                _ => {}
            }
        }
    }
}

/// A lazily-parsed row with its cells.
///
/// Unlike the generated `Row` type, this is a lightweight struct
/// that owns only the parsed data for this single row.
#[derive(Debug, Clone)]
pub struct LazyRow {
    /// 1-based row number.
    pub row_num: Option<u32>,
    /// Cells in this row.
    pub cells: Vec<LazyCell>,
    /// Style index.
    pub style: Option<u32>,
    /// Custom height.
    pub height: Option<f64>,
    /// Whether the row is hidden.
    pub hidden: Option<bool>,
}

impl LazyRow {
    /// Get the 1-based row number.
    pub fn row_num(&self) -> Option<u32> {
        self.row_num
    }

    /// Get the number of cells in this row.
    pub fn cell_count(&self) -> usize {
        self.cells.len()
    }

    /// Iterate over cells.
    pub fn cells(&self) -> impl Iterator<Item = &LazyCell> {
        self.cells.iter()
    }

    /// Get a cell by column letter (e.g., "A", "B", "AA").
    pub fn cell(&self, col: &str) -> Option<&LazyCell> {
        self.cells.iter().find(|c| {
            c.reference
                .as_ref()
                .map(|r| r.starts_with(col) && r[col.len()..].chars().all(|c| c.is_ascii_digit()))
                .unwrap_or(false)
        })
    }
}

/// A lazily-parsed cell.
#[derive(Debug, Clone)]
pub struct LazyCell {
    /// Cell reference (e.g., "A1", "B5").
    pub reference: Option<String>,
    /// Cell type (s=shared string, n=number, b=boolean, etc.).
    pub cell_type: Option<String>,
    /// Style index.
    pub style: Option<u32>,
    /// Cell value (raw string).
    pub value: Option<String>,
    /// Formula (raw string).
    pub formula: Option<String>,
}

impl LazyCell {
    /// Get the cell reference.
    pub fn reference(&self) -> Option<&str> {
        self.reference.as_deref()
    }

    /// Get the raw value string.
    pub fn value(&self) -> Option<&str> {
        self.value.as_deref()
    }

    /// Get the formula string.
    pub fn formula(&self) -> Option<&str> {
        self.formula.as_deref()
    }

    /// Check if this is a shared string reference.
    pub fn is_shared_string(&self) -> bool {
        self.cell_type.as_deref() == Some("s")
    }

    /// Check if this is a number.
    pub fn is_number(&self) -> bool {
        self.cell_type.as_deref() == Some("n") || self.cell_type.is_none()
    }

    /// Check if this is a boolean.
    pub fn is_boolean(&self) -> bool {
        self.cell_type.as_deref() == Some("b")
    }

    /// Try to parse value as a number.
    pub fn as_f64(&self) -> Option<f64> {
        self.value.as_ref()?.parse().ok()
    }

    /// Try to parse value as an integer.
    pub fn as_i64(&self) -> Option<i64> {
        self.value.as_ref()?.parse().ok()
    }

    /// Get the shared string index (if this is a shared string cell).
    pub fn shared_string_index(&self) -> Option<usize> {
        if self.is_shared_string() {
            self.value.as_ref()?.parse().ok()
        } else {
            None
        }
    }
}

// =============================================================================
// Internal parsing functions
// =============================================================================

fn parse_lazy_row<R: BufRead>(
    reader: &mut Reader<R>,
    start: &BytesStart,
) -> Result<LazyRow, ParseError> {
    let mut row = LazyRow {
        row_num: None,
        cells: Vec::new(),
        style: None,
        height: None,
        hidden: None,
    };

    // Parse attributes
    for attr in start.attributes().filter_map(|a| a.ok()) {
        let val = String::from_utf8_lossy(&attr.value);
        match attr.key.as_ref() {
            b"r" => row.row_num = val.parse().ok(),
            b"s" => row.style = val.parse().ok(),
            b"ht" => row.height = val.parse().ok(),
            b"hidden" => row.hidden = Some(val == "1" || val == "true"),
            _ => {}
        }
    }

    // Parse child elements (cells)
    let mut buf = Vec::new();
    loop {
        buf.clear();
        match reader.read_event_into(&mut buf)? {
            Event::Start(e) => {
                if e.name().as_ref() == b"c" {
                    row.cells.push(parse_lazy_cell(reader, &e)?);
                } else {
                    skip_element(reader)?;
                }
            }
            Event::Empty(e) => {
                if e.name().as_ref() == b"c" {
                    row.cells.push(parse_lazy_cell_from_empty(&e)?);
                }
            }
            Event::End(e) => {
                if e.name().as_ref() == b"row" {
                    break;
                }
            }
            Event::Eof => break,
            _ => {}
        }
    }

    Ok(row)
}

fn parse_lazy_row_from_empty(start: &BytesStart) -> Result<LazyRow, ParseError> {
    let mut row = LazyRow {
        row_num: None,
        cells: Vec::new(),
        style: None,
        height: None,
        hidden: None,
    };

    for attr in start.attributes().filter_map(|a| a.ok()) {
        let val = String::from_utf8_lossy(&attr.value);
        match attr.key.as_ref() {
            b"r" => row.row_num = val.parse().ok(),
            b"s" => row.style = val.parse().ok(),
            b"ht" => row.height = val.parse().ok(),
            b"hidden" => row.hidden = Some(val == "1" || val == "true"),
            _ => {}
        }
    }

    Ok(row)
}

fn parse_lazy_cell<R: BufRead>(
    reader: &mut Reader<R>,
    start: &BytesStart,
) -> Result<LazyCell, ParseError> {
    let mut cell = LazyCell {
        reference: None,
        cell_type: None,
        style: None,
        value: None,
        formula: None,
    };

    // Parse attributes
    for attr in start.attributes().filter_map(|a| a.ok()) {
        let val = String::from_utf8_lossy(&attr.value);
        match attr.key.as_ref() {
            b"r" => cell.reference = Some(val.into_owned()),
            b"t" => cell.cell_type = Some(val.into_owned()),
            b"s" => cell.style = val.parse().ok(),
            _ => {}
        }
    }

    // Parse child elements (v, f)
    let mut buf = Vec::new();
    loop {
        buf.clear();
        match reader.read_event_into(&mut buf)? {
            Event::Start(e) => match e.name().as_ref() {
                b"v" => {
                    cell.value = Some(read_text_content(reader)?);
                }
                b"f" => {
                    cell.formula = Some(read_text_content(reader)?);
                }
                _ => {
                    skip_element(reader)?;
                }
            },
            Event::Empty(_) => {
                // Empty v or f element - no content
            }
            Event::End(e) => {
                if e.name().as_ref() == b"c" {
                    break;
                }
            }
            Event::Eof => break,
            _ => {}
        }
    }

    Ok(cell)
}

fn parse_lazy_cell_from_empty(start: &BytesStart) -> Result<LazyCell, ParseError> {
    let mut cell = LazyCell {
        reference: None,
        cell_type: None,
        style: None,
        value: None,
        formula: None,
    };

    for attr in start.attributes().filter_map(|a| a.ok()) {
        let val = String::from_utf8_lossy(&attr.value);
        match attr.key.as_ref() {
            b"r" => cell.reference = Some(val.into_owned()),
            b"t" => cell.cell_type = Some(val.into_owned()),
            b"s" => cell.style = val.parse().ok(),
            _ => {}
        }
    }

    Ok(cell)
}

fn read_text_content<R: BufRead>(reader: &mut Reader<R>) -> Result<String, ParseError> {
    let mut text = String::new();
    let mut buf = Vec::new();
    loop {
        buf.clear();
        match reader.read_event_into(&mut buf)? {
            Event::Text(e) => {
                text.push_str(&e.decode().unwrap_or_default());
            }
            Event::CData(e) => {
                text.push_str(&e.decode().unwrap_or_default());
            }
            Event::End(_) => break,
            Event::Eof => break,
            _ => {}
        }
    }
    Ok(text)
}

fn skip_element<R: BufRead>(reader: &mut Reader<R>) -> Result<(), ParseError> {
    let mut depth = 1u32;
    let mut buf = Vec::new();
    loop {
        buf.clear();
        match reader.read_event_into(&mut buf)? {
            Event::Start(_) => depth += 1,
            Event::End(_) => {
                depth -= 1;
                if depth == 0 {
                    break;
                }
            }
            Event::Eof => break,
            _ => {}
        }
    }
    Ok(())
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_lazy_worksheet_empty() {
        let xml = br#"<?xml version="1.0"?>
            <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
                <sheetData/>
            </worksheet>"#;

        let ws = LazyWorksheet::new(xml);
        let rows: Vec<_> = ws.rows().collect();
        assert!(rows.is_empty());
    }

    #[test]
    fn test_lazy_worksheet_single_row() {
        let xml = br#"<?xml version="1.0"?>
            <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
                <sheetData>
                    <row r="1">
                        <c r="A1" t="s"><v>0</v></c>
                        <c r="B1"><v>42</v></c>
                    </row>
                </sheetData>
            </worksheet>"#;

        let ws = LazyWorksheet::new(xml);
        let rows: Vec<_> = ws.rows().collect();
        assert_eq!(rows.len(), 1);

        let row = rows[0].as_ref().unwrap();
        assert_eq!(row.row_num(), Some(1));
        assert_eq!(row.cell_count(), 2);

        let cell_a = row.cell("A").unwrap();
        assert_eq!(cell_a.reference(), Some("A1"));
        assert!(cell_a.is_shared_string());
        assert_eq!(cell_a.shared_string_index(), Some(0));

        let cell_b = row.cell("B").unwrap();
        assert_eq!(cell_b.reference(), Some("B1"));
        assert!(cell_b.is_number());
        assert_eq!(cell_b.as_f64(), Some(42.0));
    }

    #[test]
    fn test_lazy_worksheet_multiple_rows() {
        let xml = br#"<?xml version="1.0"?>
            <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
                <sheetData>
                    <row r="1"><c r="A1"><v>1</v></c></row>
                    <row r="2"><c r="A2"><v>2</v></c></row>
                    <row r="3"><c r="A3"><v>3</v></c></row>
                </sheetData>
            </worksheet>"#;

        let ws = LazyWorksheet::new(xml);
        let mut count = 0;
        for (i, row) in ws.rows().enumerate() {
            let row = row.unwrap();
            assert_eq!(row.row_num(), Some((i + 1) as u32));
            count += 1;
        }
        assert_eq!(count, 3);
    }

    #[test]
    fn test_lazy_row_with_formula() {
        let xml = br#"<?xml version="1.0"?>
            <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
                <sheetData>
                    <row r="1">
                        <c r="A1"><f>SUM(B1:B10)</f><v>55</v></c>
                    </row>
                </sheetData>
            </worksheet>"#;

        let ws = LazyWorksheet::new(xml);
        let row = ws.rows().next().unwrap().unwrap();
        let cell = row.cell("A").unwrap();
        assert_eq!(cell.formula(), Some("SUM(B1:B10)"));
        assert_eq!(cell.value(), Some("55"));
    }
}
