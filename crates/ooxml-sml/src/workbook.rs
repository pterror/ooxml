//! Workbook API for reading and writing Excel files.
//!
//! This module provides the main entry point for working with XLSX files.

use crate::error::{Error, Result};
use ooxml::{Package, Relationships};
use quick_xml::Reader;
use quick_xml::events::Event;
use std::collections::HashMap;
use std::fs::File;
use std::io::{BufReader, Cursor, Read, Seek};
use std::path::Path;

// Relationship types (ECMA-376 Part 1)
const REL_OFFICE_DOCUMENT: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument";
const REL_SHARED_STRINGS: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings";

/// An Excel workbook.
///
/// This is the main entry point for reading XLSX files.
pub struct Workbook<R: Read + Seek> {
    package: Package<R>,
    /// Path to the workbook part (e.g., "xl/workbook.xml").
    workbook_path: String,
    /// Workbook-level relationships.
    workbook_rels: Relationships,
    /// Sheet metadata (name, relationship ID).
    sheet_info: Vec<SheetInfo>,
    /// Shared string table.
    shared_strings: Vec<String>,
}

/// Metadata about a sheet.
#[derive(Debug, Clone)]
struct SheetInfo {
    name: String,
    #[allow(dead_code)]
    sheet_id: u32,
    rel_id: String,
}

impl Workbook<BufReader<File>> {
    /// Open a workbook from a file path.
    pub fn open<P: AsRef<Path>>(path: P) -> Result<Self> {
        let file = File::open(path)?;
        Self::from_reader(BufReader::new(file))
    }
}

impl<R: Read + Seek> Workbook<R> {
    /// Open a workbook from a reader.
    pub fn from_reader(reader: R) -> Result<Self> {
        let mut package = Package::open(reader)?;

        // Find the workbook part via root relationships
        let root_rels = package.read_relationships()?;
        let workbook_rel = root_rels
            .get_by_type(REL_OFFICE_DOCUMENT)
            .ok_or_else(|| Error::Invalid("Missing workbook relationship".into()))?;
        let workbook_path = workbook_rel.target.clone();

        // Load workbook relationships
        let workbook_rels = package
            .read_part_relationships(&workbook_path)
            .unwrap_or_default();

        // Parse workbook.xml to get sheet list
        let workbook_xml = package.read_part(&workbook_path)?;
        let sheet_info = parse_workbook_sheets(&workbook_xml)?;

        // Load shared strings if present
        let shared_strings = if let Some(rel) = workbook_rels.get_by_type(REL_SHARED_STRINGS) {
            let path = resolve_path(&workbook_path, &rel.target);
            if let Ok(data) = package.read_part(&path) {
                parse_shared_strings(&data)?
            } else {
                Vec::new()
            }
        } else {
            Vec::new()
        };

        Ok(Self {
            package,
            workbook_path,
            workbook_rels,
            sheet_info,
            shared_strings,
        })
    }

    /// Get the number of sheets in the workbook.
    pub fn sheet_count(&self) -> usize {
        self.sheet_info.len()
    }

    /// Get sheet names.
    pub fn sheet_names(&self) -> Vec<&str> {
        self.sheet_info.iter().map(|s| s.name.as_str()).collect()
    }

    /// Get a sheet by index.
    pub fn sheet(&mut self, index: usize) -> Result<Sheet> {
        let info = self
            .sheet_info
            .get(index)
            .ok_or_else(|| Error::Invalid(format!("Sheet index {} out of range", index)))?
            .clone();

        self.load_sheet(&info)
    }

    /// Get a sheet by name.
    pub fn sheet_by_name(&mut self, name: &str) -> Result<Sheet> {
        let info = self
            .sheet_info
            .iter()
            .find(|s| s.name == name)
            .ok_or_else(|| Error::Invalid(format!("Sheet '{}' not found", name)))?
            .clone();

        self.load_sheet(&info)
    }

    /// Load all sheets.
    pub fn sheets(&mut self) -> Result<Vec<Sheet>> {
        let infos: Vec<_> = self.sheet_info.clone();
        infos.iter().map(|info| self.load_sheet(info)).collect()
    }

    /// Load a sheet's data.
    fn load_sheet(&mut self, info: &SheetInfo) -> Result<Sheet> {
        // Find the sheet path from relationships
        let rel = self.workbook_rels.get(&info.rel_id).ok_or_else(|| {
            Error::Invalid(format!("Missing relationship for sheet '{}'", info.name))
        })?;

        let path = resolve_path(&self.workbook_path, &rel.target);
        let data = self.package.read_part(&path)?;

        parse_sheet(&data, &info.name, &self.shared_strings)
    }
}

/// A worksheet in the workbook.
#[derive(Debug, Clone)]
pub struct Sheet {
    name: String,
    rows: Vec<Row>,
}

impl Sheet {
    /// Get the sheet name.
    pub fn name(&self) -> &str {
        &self.name
    }

    /// Get all rows in the sheet.
    pub fn rows(&self) -> &[Row] {
        &self.rows
    }

    /// Get a specific row by 1-based index.
    pub fn row(&self, row_num: u32) -> Option<&Row> {
        self.rows.iter().find(|r| r.row_num == row_num)
    }

    /// Get a cell by reference (e.g., "A1", "B2").
    pub fn cell(&self, reference: &str) -> Option<&Cell> {
        let (col, row) = parse_cell_reference(reference)?;
        self.row(row)?.cell_at_column(col)
    }

    /// Get the used range dimensions.
    pub fn dimensions(&self) -> Option<(u32, u32, u32, u32)> {
        if self.rows.is_empty() {
            return None;
        }

        let min_row = self.rows.iter().map(|r| r.row_num).min().unwrap();
        let max_row = self.rows.iter().map(|r| r.row_num).max().unwrap();
        let min_col = self
            .rows
            .iter()
            .flat_map(|r| r.cells.iter().map(|c| c.column))
            .min()
            .unwrap_or(1);
        let max_col = self
            .rows
            .iter()
            .flat_map(|r| r.cells.iter().map(|c| c.column))
            .max()
            .unwrap_or(1);

        Some((min_row, min_col, max_row, max_col))
    }
}

/// A row in a worksheet.
#[derive(Debug, Clone)]
pub struct Row {
    row_num: u32,
    cells: Vec<Cell>,
}

impl Row {
    /// Get the 1-based row number.
    pub fn row_num(&self) -> u32 {
        self.row_num
    }

    /// Get all cells in the row.
    pub fn cells(&self) -> &[Cell] {
        &self.cells
    }

    /// Get a cell at a specific 1-based column index.
    pub fn cell_at_column(&self, col: u32) -> Option<&Cell> {
        self.cells.iter().find(|c| c.column == col)
    }

    /// Get a cell by column letter (e.g., "A", "B", "AA").
    pub fn cell(&self, col_letter: &str) -> Option<&Cell> {
        let col = column_letter_to_number(col_letter)?;
        self.cell_at_column(col)
    }
}

/// A cell in a worksheet.
#[derive(Debug, Clone)]
pub struct Cell {
    /// Cell reference (e.g., "A1").
    reference: String,
    /// 1-based column number.
    column: u32,
    /// Cell value.
    value: CellValue,
    /// Formula (if any).
    formula: Option<String>,
}

impl Cell {
    /// Get the cell reference (e.g., "A1").
    pub fn reference(&self) -> &str {
        &self.reference
    }

    /// Get the 1-based column number.
    pub fn column(&self) -> u32 {
        self.column
    }

    /// Get the cell value.
    pub fn value(&self) -> &CellValue {
        &self.value
    }

    /// Get the formula (if any).
    pub fn formula(&self) -> Option<&str> {
        self.formula.as_deref()
    }

    /// Get the value as a string (for display).
    pub fn value_as_string(&self) -> String {
        match &self.value {
            CellValue::Empty => String::new(),
            CellValue::String(s) => s.clone(),
            CellValue::Number(n) => n.to_string(),
            CellValue::Boolean(b) => if *b { "TRUE" } else { "FALSE" }.to_string(),
            CellValue::Error(e) => e.clone(),
        }
    }

    /// Try to get the value as a number.
    pub fn value_as_number(&self) -> Option<f64> {
        match &self.value {
            CellValue::Number(n) => Some(*n),
            CellValue::String(s) => s.parse().ok(),
            _ => None,
        }
    }

    /// Try to get the value as a boolean.
    pub fn value_as_bool(&self) -> Option<bool> {
        match &self.value {
            CellValue::Boolean(b) => Some(*b),
            CellValue::Number(n) => Some(*n != 0.0),
            CellValue::String(s) => match s.to_lowercase().as_str() {
                "true" | "1" => Some(true),
                "false" | "0" => Some(false),
                _ => None,
            },
            _ => None,
        }
    }
}

/// The value of a cell.
#[derive(Debug, Clone, PartialEq)]
pub enum CellValue {
    /// Empty cell.
    Empty,
    /// String value.
    String(String),
    /// Numeric value.
    Number(f64),
    /// Boolean value.
    Boolean(bool),
    /// Error value (e.g., "#DIV/0!").
    Error(String),
}

// ============================================================================
// Parsing
// ============================================================================

/// Parse the workbook.xml to extract sheet information.
fn parse_workbook_sheets(xml: &[u8]) -> Result<Vec<SheetInfo>> {
    let mut reader = Reader::from_reader(Cursor::new(xml));
    let mut buf = Vec::new();
    let mut sheets = Vec::new();
    let mut in_sheets = false;

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) | Ok(Event::Empty(e)) => {
                let name = e.name();
                let name = name.as_ref();
                if name == b"sheets" {
                    in_sheets = true;
                } else if in_sheets && name == b"sheet" {
                    let mut sheet_name = String::new();
                    let mut sheet_id = 0u32;
                    let mut rel_id = String::new();

                    for attr in e.attributes().filter_map(|a| a.ok()) {
                        match attr.key.as_ref() {
                            b"name" => {
                                sheet_name = String::from_utf8_lossy(&attr.value).into_owned();
                            }
                            b"sheetId" => {
                                sheet_id =
                                    String::from_utf8_lossy(&attr.value).parse().unwrap_or(0);
                            }
                            key if key == b"r:id" || key == b"id" => {
                                rel_id = String::from_utf8_lossy(&attr.value).into_owned();
                            }
                            _ => {}
                        }
                    }

                    if !sheet_name.is_empty() && !rel_id.is_empty() {
                        sheets.push(SheetInfo {
                            name: sheet_name,
                            sheet_id,
                            rel_id,
                        });
                    }
                }
            }
            Ok(Event::End(e)) => {
                let name = e.name();
                if name.as_ref() == b"sheets" {
                    in_sheets = false;
                }
            }
            Ok(Event::Eof) => break,
            Err(e) => return Err(Error::Xml(e)),
            _ => {}
        }
        buf.clear();
    }

    Ok(sheets)
}

/// Parse the shared strings table.
fn parse_shared_strings(xml: &[u8]) -> Result<Vec<String>> {
    let mut reader = Reader::from_reader(Cursor::new(xml));
    let mut buf = Vec::new();
    let mut strings = Vec::new();
    let mut current_string = String::new();
    let mut in_si = false;
    let mut in_t = false;

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) => {
                let name = e.name();
                let name = name.as_ref();
                if name == b"si" {
                    in_si = true;
                    current_string.clear();
                } else if in_si && name == b"t" {
                    in_t = true;
                }
            }
            Ok(Event::Empty(e)) => {
                // Handle self-closing <t/> (empty string)
                let name = e.name();
                if in_si && name.as_ref() == b"t" {
                    // Empty text element, nothing to add
                }
            }
            Ok(Event::Text(e)) => {
                if in_t {
                    current_string.push_str(&e.decode().unwrap_or_default());
                }
            }
            Ok(Event::End(e)) => {
                let name = e.name();
                let name = name.as_ref();
                if name == b"si" {
                    strings.push(std::mem::take(&mut current_string));
                    in_si = false;
                } else if name == b"t" {
                    in_t = false;
                }
            }
            Ok(Event::Eof) => break,
            Err(e) => return Err(Error::Xml(e)),
            _ => {}
        }
        buf.clear();
    }

    Ok(strings)
}

/// Parse a worksheet.
fn parse_sheet(xml: &[u8], name: &str, shared_strings: &[String]) -> Result<Sheet> {
    let mut reader = Reader::from_reader(Cursor::new(xml));
    let mut buf = Vec::new();
    let mut rows: HashMap<u32, Vec<Cell>> = HashMap::new();

    let mut current_row: u32 = 0;
    let mut current_cell_ref = String::new();
    let mut current_cell_type = String::new();
    let mut current_cell_value = String::new();
    let mut current_formula = Option::<String>::None;
    let mut in_cell = false;
    let mut in_value = false;
    let mut in_formula = false;

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) => {
                let tag = e.name();
                let tag = tag.as_ref();
                match tag {
                    b"row" => {
                        for attr in e.attributes().filter_map(|a| a.ok()) {
                            if attr.key.as_ref() == b"r" {
                                current_row =
                                    String::from_utf8_lossy(&attr.value).parse().unwrap_or(0);
                            }
                        }
                    }
                    b"c" => {
                        in_cell = true;
                        current_cell_ref.clear();
                        current_cell_type.clear();
                        current_cell_value.clear();
                        current_formula = None;

                        for attr in e.attributes().filter_map(|a| a.ok()) {
                            match attr.key.as_ref() {
                                b"r" => {
                                    current_cell_ref =
                                        String::from_utf8_lossy(&attr.value).into_owned();
                                }
                                b"t" => {
                                    current_cell_type =
                                        String::from_utf8_lossy(&attr.value).into_owned();
                                }
                                _ => {}
                            }
                        }
                    }
                    b"v" if in_cell => {
                        in_value = true;
                    }
                    b"f" if in_cell => {
                        in_formula = true;
                        current_formula = Some(String::new());
                    }
                    _ => {}
                }
            }
            Ok(Event::Empty(e)) => {
                // Handle empty cells <c r="A1"/>
                let tag = e.name();
                if tag.as_ref() == b"c" {
                    let mut cell_ref = String::new();
                    for attr in e.attributes().filter_map(|a| a.ok()) {
                        if attr.key.as_ref() == b"r" {
                            cell_ref = String::from_utf8_lossy(&attr.value).into_owned();
                        }
                    }
                    if !cell_ref.is_empty() {
                        let column = parse_cell_reference(&cell_ref).map(|(c, _)| c).unwrap_or(1);
                        let cell = Cell {
                            reference: cell_ref,
                            column,
                            value: CellValue::Empty,
                            formula: None,
                        };
                        rows.entry(current_row).or_default().push(cell);
                    }
                }
            }
            Ok(Event::Text(e)) => {
                if in_value {
                    current_cell_value.push_str(&e.decode().unwrap_or_default());
                } else if in_formula && let Some(ref mut f) = current_formula {
                    f.push_str(&e.decode().unwrap_or_default());
                }
            }
            Ok(Event::End(e)) => {
                let tag = e.name();
                let tag = tag.as_ref();
                match tag {
                    b"c" if in_cell => {
                        let column = parse_cell_reference(&current_cell_ref)
                            .map(|(c, _)| c)
                            .unwrap_or(1);

                        let value = match current_cell_type.as_str() {
                            "s" => {
                                // Shared string index
                                let idx: usize = current_cell_value.parse().unwrap_or(0);
                                shared_strings
                                    .get(idx)
                                    .map(|s| CellValue::String(s.clone()))
                                    .unwrap_or(CellValue::Empty)
                            }
                            "str" | "inlineStr" => {
                                // Inline string
                                CellValue::String(std::mem::take(&mut current_cell_value))
                            }
                            "b" => {
                                // Boolean
                                CellValue::Boolean(current_cell_value == "1")
                            }
                            "e" => {
                                // Error
                                CellValue::Error(std::mem::take(&mut current_cell_value))
                            }
                            _ => {
                                // Number (default) or empty
                                if current_cell_value.is_empty() {
                                    CellValue::Empty
                                } else if let Ok(n) = current_cell_value.parse::<f64>() {
                                    CellValue::Number(n)
                                } else {
                                    CellValue::String(std::mem::take(&mut current_cell_value))
                                }
                            }
                        };

                        let cell = Cell {
                            reference: std::mem::take(&mut current_cell_ref),
                            column,
                            value,
                            formula: current_formula.take(),
                        };

                        rows.entry(current_row).or_default().push(cell);
                        in_cell = false;
                    }
                    b"v" => in_value = false,
                    b"f" => in_formula = false,
                    _ => {}
                }
            }
            Ok(Event::Eof) => break,
            Err(e) => return Err(Error::Xml(e)),
            _ => {}
        }
        buf.clear();
    }

    // Convert HashMap to sorted Vec<Row>
    let mut row_list: Vec<Row> = rows
        .into_iter()
        .map(|(row_num, cells)| Row { row_num, cells })
        .collect();
    row_list.sort_by_key(|r| r.row_num);

    // Sort cells within each row by column
    for row in &mut row_list {
        row.cells.sort_by_key(|c| c.column);
    }

    Ok(Sheet {
        name: name.to_string(),
        rows: row_list,
    })
}

// ============================================================================
// Utilities
// ============================================================================

/// Convert a column letter to a 1-based column number.
/// E.g., "A" -> 1, "B" -> 2, "Z" -> 26, "AA" -> 27
fn column_letter_to_number(col: &str) -> Option<u32> {
    let mut result: u32 = 0;
    for c in col.chars() {
        if !c.is_ascii_alphabetic() {
            return None;
        }
        result = result * 26 + (c.to_ascii_uppercase() as u32 - 'A' as u32 + 1);
    }
    if result == 0 { None } else { Some(result) }
}

/// Parse a cell reference like "A1" into (column, row).
fn parse_cell_reference(reference: &str) -> Option<(u32, u32)> {
    let col_end = reference
        .chars()
        .take_while(|c| c.is_ascii_alphabetic())
        .count();
    if col_end == 0 {
        return None;
    }

    let col_str = &reference[..col_end];
    let row_str = &reference[col_end..];

    let col = column_letter_to_number(col_str)?;
    let row: u32 = row_str.parse().ok()?;

    Some((col, row))
}

/// Resolve a relative path against a base path.
fn resolve_path(base: &str, target: &str) -> String {
    if target.starts_with('/') {
        return target.to_string();
    }

    // Get the directory of the base path
    let base_dir = if let Some(idx) = base.rfind('/') {
        &base[..=idx]
    } else {
        "/"
    };

    format!("{}{}", base_dir, target)
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_column_letter_to_number() {
        assert_eq!(column_letter_to_number("A"), Some(1));
        assert_eq!(column_letter_to_number("B"), Some(2));
        assert_eq!(column_letter_to_number("Z"), Some(26));
        assert_eq!(column_letter_to_number("AA"), Some(27));
        assert_eq!(column_letter_to_number("AB"), Some(28));
        assert_eq!(column_letter_to_number("AZ"), Some(52));
        assert_eq!(column_letter_to_number("BA"), Some(53));
        assert_eq!(column_letter_to_number(""), None);
        assert_eq!(column_letter_to_number("1"), None);
    }

    #[test]
    fn test_parse_cell_reference() {
        assert_eq!(parse_cell_reference("A1"), Some((1, 1)));
        assert_eq!(parse_cell_reference("B2"), Some((2, 2)));
        assert_eq!(parse_cell_reference("Z10"), Some((26, 10)));
        assert_eq!(parse_cell_reference("AA100"), Some((27, 100)));
        assert_eq!(parse_cell_reference("1A"), None);
        assert_eq!(parse_cell_reference(""), None);
    }

    #[test]
    fn test_resolve_path() {
        assert_eq!(
            resolve_path("/xl/workbook.xml", "worksheets/sheet1.xml"),
            "/xl/worksheets/sheet1.xml"
        );
        assert_eq!(
            resolve_path("/xl/workbook.xml", "/xl/sharedStrings.xml"),
            "/xl/sharedStrings.xml"
        );
    }

    #[test]
    fn test_parse_shared_strings() {
        let xml = r#"<?xml version="1.0"?>
        <sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
            <si><t>Hello</t></si>
            <si><t>World</t></si>
            <si><t></t></si>
        </sst>"#;

        let strings = parse_shared_strings(xml.as_bytes()).unwrap();
        assert_eq!(strings, vec!["Hello", "World", ""]);
    }
}
