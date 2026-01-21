//! Workbook API for reading and writing Excel files.
//!
//! This module provides the main entry point for working with XLSX files.

use crate::error::{Error, Result};
use ooxml_opc::{Package, Relationships};
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
const REL_STYLES: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles";
const REL_COMMENTS: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments";

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
    /// Stylesheet (number formats, fonts, fills, borders, cell formats).
    styles: Stylesheet,
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

        // Load styles if present
        let styles = if let Some(rel) = workbook_rels.get_by_type(REL_STYLES) {
            let path = resolve_path(&workbook_path, &rel.target);
            if let Ok(data) = package.read_part(&path) {
                parse_styles(&data)?
            } else {
                Stylesheet::default()
            }
        } else {
            Stylesheet::default()
        };

        Ok(Self {
            package,
            workbook_path,
            workbook_rels,
            sheet_info,
            shared_strings,
            styles,
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

    /// Get the workbook stylesheet.
    pub fn styles(&self) -> &Stylesheet {
        &self.styles
    }

    /// Load a sheet's data.
    fn load_sheet(&mut self, info: &SheetInfo) -> Result<Sheet> {
        // Find the sheet path from relationships
        let rel = self.workbook_rels.get(&info.rel_id).ok_or_else(|| {
            Error::Invalid(format!("Missing relationship for sheet '{}'", info.name))
        })?;

        let path = resolve_path(&self.workbook_path, &rel.target);
        let data = self.package.read_part(&path)?;

        let mut sheet = parse_sheet(&data, &info.name, &self.shared_strings, &self.styles)?;

        // Try to load comments for this sheet
        if let Ok(sheet_rels) = self.package.read_part_relationships(&path)
            && let Some(comments_rel) = sheet_rels.get_by_type(REL_COMMENTS)
        {
            let comments_path = resolve_path(&path, &comments_rel.target);
            if let Ok(comments_data) = self.package.read_part(&comments_path) {
                sheet.comments = parse_comments(&comments_data)?;
            }
        }

        Ok(sheet)
    }
}

/// A merged cell range.
///
/// ECMA-376 Part 1, Section 18.3.1.55 (mergeCell).
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct MergedCell {
    /// The range reference (e.g., "A1:B2").
    reference: String,
    /// Top-left row (1-based).
    start_row: u32,
    /// Top-left column (1-based).
    start_col: u32,
    /// Bottom-right row (1-based).
    end_row: u32,
    /// Bottom-right column (1-based).
    end_col: u32,
}

impl MergedCell {
    /// Get the range reference (e.g., "A1:B2").
    pub fn reference(&self) -> &str {
        &self.reference
    }

    /// Get the top-left cell reference.
    pub fn start_cell(&self) -> String {
        format!(
            "{}{}",
            column_number_to_letter(self.start_col),
            self.start_row
        )
    }

    /// Get the bottom-right cell reference.
    pub fn end_cell(&self) -> String {
        format!("{}{}", column_number_to_letter(self.end_col), self.end_row)
    }

    /// Get the start row (1-based).
    pub fn start_row(&self) -> u32 {
        self.start_row
    }

    /// Get the start column (1-based).
    pub fn start_col(&self) -> u32 {
        self.start_col
    }

    /// Get the end row (1-based).
    pub fn end_row(&self) -> u32 {
        self.end_row
    }

    /// Get the end column (1-based).
    pub fn end_col(&self) -> u32 {
        self.end_col
    }

    /// Check if a cell at (row, col) is within this merged range.
    pub fn contains(&self, row: u32, col: u32) -> bool {
        row >= self.start_row && row <= self.end_row && col >= self.start_col && col <= self.end_col
    }

    /// Check if a cell reference is within this merged range.
    pub fn contains_ref(&self, reference: &str) -> bool {
        if let Some((col, row)) = parse_cell_reference(reference) {
            self.contains(row, col)
        } else {
            false
        }
    }
}

/// Column information (width and style).
///
/// ECMA-376 Part 1, Section 18.3.1.13 (col).
#[derive(Debug, Clone)]
pub struct ColumnInfo {
    /// First column in the range (1-based).
    pub min: u32,
    /// Last column in the range (1-based).
    pub max: u32,
    /// Column width in character units (Excel default is 8.43).
    pub width: Option<f64>,
    /// Whether the column is hidden.
    pub hidden: bool,
    /// Whether the width was explicitly set (not default).
    pub custom_width: bool,
    /// Style index for the column.
    pub style_index: Option<u32>,
}

impl ColumnInfo {
    /// Check if this column info applies to a specific column number.
    pub fn applies_to(&self, col: u32) -> bool {
        col >= self.min && col <= self.max
    }
}

/// A cell comment (note).
///
/// ECMA-376 Part 1, Section 18.7 (Comments).
#[derive(Debug, Clone)]
pub struct Comment {
    /// Cell reference (e.g., "A1").
    reference: String,
    /// Author of the comment.
    author: Option<String>,
    /// Comment text content.
    text: String,
}

impl Comment {
    /// Get the cell reference (e.g., "A1").
    pub fn reference(&self) -> &str {
        &self.reference
    }

    /// Get the author of the comment.
    pub fn author(&self) -> Option<&str> {
        self.author.as_deref()
    }

    /// Get the comment text.
    pub fn text(&self) -> &str {
        &self.text
    }
}

/// A worksheet in the workbook.
#[derive(Debug, Clone)]
pub struct Sheet {
    name: String,
    rows: Vec<Row>,
    merged_cells: Vec<MergedCell>,
    columns: Vec<ColumnInfo>,
    comments: Vec<Comment>,
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

    /// Get all merged cell ranges in this sheet.
    pub fn merged_cells(&self) -> &[MergedCell] {
        &self.merged_cells
    }

    /// Check if a cell is part of a merged range.
    ///
    /// Returns the merged cell range if the cell is within one.
    pub fn merged_cell_at(&self, reference: &str) -> Option<&MergedCell> {
        self.merged_cells
            .iter()
            .find(|mc| mc.contains_ref(reference))
    }

    /// Check if a cell at (row, col) is part of a merged range.
    pub fn merged_cell_at_pos(&self, row: u32, col: u32) -> Option<&MergedCell> {
        self.merged_cells.iter().find(|mc| mc.contains(row, col))
    }

    /// Get all column definitions.
    pub fn columns(&self) -> &[ColumnInfo] {
        &self.columns
    }

    /// Get column info for a specific column (1-based).
    pub fn column_info(&self, col: u32) -> Option<&ColumnInfo> {
        self.columns.iter().find(|c| c.applies_to(col))
    }

    /// Get the width of a specific column (1-based).
    ///
    /// Returns None if no custom width is set.
    pub fn column_width(&self, col: u32) -> Option<f64> {
        self.column_info(col).and_then(|c| c.width)
    }

    /// Get all comments in this sheet.
    pub fn comments(&self) -> &[Comment] {
        &self.comments
    }

    /// Get the comment for a specific cell (e.g., "A1").
    pub fn comment(&self, reference: &str) -> Option<&Comment> {
        self.comments.iter().find(|c| c.reference == reference)
    }

    /// Check if a cell has a comment.
    pub fn has_comment(&self, reference: &str) -> bool {
        self.comment(reference).is_some()
    }
}

/// A row in a worksheet.
#[derive(Debug, Clone)]
pub struct Row {
    row_num: u32,
    cells: Vec<Cell>,
    /// Row height in points (Excel default is ~15).
    height: Option<f64>,
    /// Whether the row is hidden.
    hidden: bool,
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

    /// Get the row height in points.
    ///
    /// Returns None if using default height.
    pub fn height(&self) -> Option<f64> {
        self.height
    }

    /// Check if the row is hidden.
    pub fn is_hidden(&self) -> bool {
        self.hidden
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
    /// Style index (into workbook.styles().cell_formats).
    style_index: Option<u32>,
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

    /// Get the style index for this cell.
    ///
    /// Use this with `workbook.styles().cell_formats` to get formatting details.
    /// Returns `None` if no style is applied.
    pub fn style_index(&self) -> Option<u32> {
        self.style_index
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

/// Stylesheet containing number formats, fonts, fills, and cell styles.
///
/// ECMA-376 Part 1, Section 18.8 (Styles).
#[derive(Debug, Clone, Default)]
pub struct Stylesheet {
    /// Number formats (custom format codes).
    pub number_formats: Vec<NumberFormat>,
    /// Font definitions.
    pub fonts: Vec<Font>,
    /// Fill definitions.
    pub fills: Vec<Fill>,
    /// Border definitions.
    pub borders: Vec<Border>,
    /// Cell format records (combines font, fill, border, number format).
    pub cell_formats: Vec<CellFormat>,
}

/// A number format definition.
///
/// ECMA-376 Part 1, Section 18.8.30 (numFmt).
#[derive(Debug, Clone)]
pub struct NumberFormat {
    /// Format ID (built-in formats use IDs 0-163).
    pub id: u32,
    /// Format code (e.g., "0.00", "#,##0", "yyyy-mm-dd").
    pub code: String,
}

/// A font definition.
///
/// ECMA-376 Part 1, Section 18.8.22 (font).
#[derive(Debug, Clone, Default)]
pub struct Font {
    /// Font name (e.g., "Calibri", "Arial").
    pub name: Option<String>,
    /// Font size in points.
    pub size: Option<f64>,
    /// Bold.
    pub bold: bool,
    /// Italic.
    pub italic: bool,
    /// Underline.
    pub underline: bool,
    /// Strikethrough.
    pub strike: bool,
    /// Font color (RGB hex without #, or theme reference).
    pub color: Option<String>,
}

/// A fill definition.
///
/// ECMA-376 Part 1, Section 18.8.20 (fill).
#[derive(Debug, Clone, Default)]
pub struct Fill {
    /// Pattern type (e.g., "solid", "none", "gray125").
    pub pattern_type: Option<String>,
    /// Foreground color (RGB hex).
    pub fg_color: Option<String>,
    /// Background color (RGB hex).
    pub bg_color: Option<String>,
}

/// A border definition.
///
/// ECMA-376 Part 1, Section 18.8.4 (border).
#[derive(Debug, Clone, Default)]
pub struct Border {
    /// Left border style.
    pub left: Option<BorderSide>,
    /// Right border style.
    pub right: Option<BorderSide>,
    /// Top border style.
    pub top: Option<BorderSide>,
    /// Bottom border style.
    pub bottom: Option<BorderSide>,
}

/// A single border side.
#[derive(Debug, Clone, Default)]
pub struct BorderSide {
    /// Border style (e.g., "thin", "medium", "thick", "dashed").
    pub style: Option<String>,
    /// Border color (RGB hex).
    pub color: Option<String>,
}

/// A cell format record combining style elements.
///
/// ECMA-376 Part 1, Section 18.8.45 (xf).
#[derive(Debug, Clone, Default)]
pub struct CellFormat {
    /// Index into number_formats (or built-in format ID).
    pub number_format_id: u32,
    /// Index into fonts.
    pub font_id: u32,
    /// Index into fills.
    pub fill_id: u32,
    /// Index into borders.
    pub border_id: u32,
    /// Horizontal alignment.
    pub horizontal_align: Option<String>,
    /// Vertical alignment.
    pub vertical_align: Option<String>,
    /// Text wrap.
    pub wrap_text: bool,
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

/// Parse the styles.xml file.
///
/// ECMA-376 Part 1, Section 18.8 (Styles).
fn parse_styles(xml: &[u8]) -> Result<Stylesheet> {
    let mut reader = Reader::from_reader(Cursor::new(xml));
    let mut buf = Vec::new();
    let mut styles = Stylesheet::default();

    // Parsing state
    let mut in_num_fmts = false;
    let mut in_fonts = false;
    let mut in_fills = false;
    let mut in_borders = false;
    let mut in_cell_xfs = false;
    let mut in_font = false;
    let mut in_fill = false;
    let mut in_border = false;
    let mut in_xf = false;

    let mut current_font = Font::default();
    let mut current_fill = Fill::default();
    let mut current_border = Border::default();
    let mut current_xf = CellFormat::default();

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) => {
                let tag = e.name();
                let tag = tag.as_ref();
                match tag {
                    b"numFmts" => in_num_fmts = true,
                    b"fonts" => in_fonts = true,
                    b"fills" => in_fills = true,
                    b"borders" => in_borders = true,
                    b"cellXfs" => in_cell_xfs = true,
                    b"font" if in_fonts => {
                        in_font = true;
                        current_font = Font::default();
                    }
                    b"fill" if in_fills => {
                        in_fill = true;
                        current_fill = Fill::default();
                    }
                    b"patternFill" if in_fill => {
                        for attr in e.attributes().filter_map(|a| a.ok()) {
                            if attr.key.as_ref() == b"patternType" {
                                current_fill.pattern_type =
                                    Some(String::from_utf8_lossy(&attr.value).into_owned());
                            }
                        }
                    }
                    b"border" if in_borders => {
                        in_border = true;
                        current_border = Border::default();
                    }
                    b"left" | b"right" | b"top" | b"bottom" if in_border => {
                        // Border sides with content are handled in End event
                    }
                    b"xf" if in_cell_xfs => {
                        in_xf = true;
                        current_xf = CellFormat::default();
                        for attr in e.attributes().filter_map(|a| a.ok()) {
                            let val = String::from_utf8_lossy(&attr.value);
                            match attr.key.as_ref() {
                                b"numFmtId" => {
                                    current_xf.number_format_id = val.parse().unwrap_or(0)
                                }
                                b"fontId" => current_xf.font_id = val.parse().unwrap_or(0),
                                b"fillId" => current_xf.fill_id = val.parse().unwrap_or(0),
                                b"borderId" => current_xf.border_id = val.parse().unwrap_or(0),
                                _ => {}
                            }
                        }
                    }
                    b"alignment" if in_xf => {
                        for attr in e.attributes().filter_map(|a| a.ok()) {
                            let val = String::from_utf8_lossy(&attr.value);
                            match attr.key.as_ref() {
                                b"horizontal" => {
                                    current_xf.horizontal_align = Some(val.into_owned())
                                }
                                b"vertical" => current_xf.vertical_align = Some(val.into_owned()),
                                b"wrapText" => current_xf.wrap_text = val == "1" || val == "true",
                                _ => {}
                            }
                        }
                    }
                    _ => {}
                }
            }
            Ok(Event::Empty(e)) => {
                let tag = e.name();
                let tag = tag.as_ref();
                match tag {
                    b"numFmt" if in_num_fmts => {
                        let mut id = 0u32;
                        let mut code = String::new();
                        for attr in e.attributes().filter_map(|a| a.ok()) {
                            let val = String::from_utf8_lossy(&attr.value);
                            match attr.key.as_ref() {
                                b"numFmtId" => id = val.parse().unwrap_or(0),
                                b"formatCode" => code = val.into_owned(),
                                _ => {}
                            }
                        }
                        styles.number_formats.push(NumberFormat { id, code });
                    }
                    b"name" if in_font => {
                        for attr in e.attributes().filter_map(|a| a.ok()) {
                            if attr.key.as_ref() == b"val" {
                                current_font.name =
                                    Some(String::from_utf8_lossy(&attr.value).into_owned());
                            }
                        }
                    }
                    b"sz" if in_font => {
                        for attr in e.attributes().filter_map(|a| a.ok()) {
                            if attr.key.as_ref() == b"val" {
                                current_font.size =
                                    String::from_utf8_lossy(&attr.value).parse().ok();
                            }
                        }
                    }
                    b"b" if in_font => current_font.bold = true,
                    b"i" if in_font => current_font.italic = true,
                    b"u" if in_font => current_font.underline = true,
                    b"strike" if in_font => current_font.strike = true,
                    b"color" if in_font => {
                        for attr in e.attributes().filter_map(|a| a.ok()) {
                            if attr.key.as_ref() == b"rgb" {
                                current_font.color =
                                    Some(String::from_utf8_lossy(&attr.value).into_owned());
                            }
                        }
                    }
                    b"patternFill" if in_fill => {
                        // Handle self-closing patternFill
                        for attr in e.attributes().filter_map(|a| a.ok()) {
                            if attr.key.as_ref() == b"patternType" {
                                current_fill.pattern_type =
                                    Some(String::from_utf8_lossy(&attr.value).into_owned());
                            }
                        }
                    }
                    b"fgColor" if in_fill => {
                        for attr in e.attributes().filter_map(|a| a.ok()) {
                            if attr.key.as_ref() == b"rgb" {
                                current_fill.fg_color =
                                    Some(String::from_utf8_lossy(&attr.value).into_owned());
                            }
                        }
                    }
                    b"bgColor" if in_fill => {
                        for attr in e.attributes().filter_map(|a| a.ok()) {
                            if attr.key.as_ref() == b"rgb" {
                                current_fill.bg_color =
                                    Some(String::from_utf8_lossy(&attr.value).into_owned());
                            }
                        }
                    }
                    b"left" | b"right" | b"top" | b"bottom" if in_border => {
                        let mut side = BorderSide::default();
                        for attr in e.attributes().filter_map(|a| a.ok()) {
                            if attr.key.as_ref() == b"style" {
                                side.style =
                                    Some(String::from_utf8_lossy(&attr.value).into_owned());
                            }
                        }
                        match tag {
                            b"left" => current_border.left = Some(side),
                            b"right" => current_border.right = Some(side),
                            b"top" => current_border.top = Some(side),
                            b"bottom" => current_border.bottom = Some(side),
                            _ => {}
                        }
                    }
                    b"xf" if in_cell_xfs => {
                        let mut xf = CellFormat::default();
                        for attr in e.attributes().filter_map(|a| a.ok()) {
                            let val = String::from_utf8_lossy(&attr.value);
                            match attr.key.as_ref() {
                                b"numFmtId" => xf.number_format_id = val.parse().unwrap_or(0),
                                b"fontId" => xf.font_id = val.parse().unwrap_or(0),
                                b"fillId" => xf.fill_id = val.parse().unwrap_or(0),
                                b"borderId" => xf.border_id = val.parse().unwrap_or(0),
                                _ => {}
                            }
                        }
                        styles.cell_formats.push(xf);
                    }
                    _ => {}
                }
            }
            Ok(Event::End(e)) => {
                let tag = e.name();
                let tag = tag.as_ref();
                match tag {
                    b"numFmts" => in_num_fmts = false,
                    b"fonts" => in_fonts = false,
                    b"fills" => in_fills = false,
                    b"borders" => in_borders = false,
                    b"cellXfs" => in_cell_xfs = false,
                    b"font" if in_fonts => {
                        styles.fonts.push(std::mem::take(&mut current_font));
                        in_font = false;
                    }
                    b"fill" if in_fills => {
                        styles.fills.push(std::mem::take(&mut current_fill));
                        in_fill = false;
                    }
                    b"border" if in_borders => {
                        styles.borders.push(std::mem::take(&mut current_border));
                        in_border = false;
                    }
                    b"xf" if in_cell_xfs => {
                        styles.cell_formats.push(std::mem::take(&mut current_xf));
                        in_xf = false;
                    }
                    _ => {}
                }
            }
            Ok(Event::Eof) => break,
            Err(e) => return Err(Error::Xml(e)),
            _ => {}
        }
        buf.clear();
    }

    Ok(styles)
}

/// Parse a worksheet.
fn parse_sheet(
    xml: &[u8],
    name: &str,
    shared_strings: &[String],
    _styles: &Stylesheet,
) -> Result<Sheet> {
    let mut reader = Reader::from_reader(Cursor::new(xml));
    let mut buf = Vec::new();
    let mut rows: HashMap<u32, (Vec<Cell>, Option<f64>, bool)> = HashMap::new(); // (cells, height, hidden)
    let mut merged_cells: Vec<MergedCell> = Vec::new();
    let mut columns: Vec<ColumnInfo> = Vec::new();

    let mut current_row: u32 = 0;
    let mut current_row_height: Option<f64> = None;
    let mut current_row_hidden = false;
    let mut current_cell_ref = String::new();
    let mut current_cell_type = String::new();
    let mut current_cell_style: Option<u32> = None;
    let mut current_cell_value = String::new();
    let mut current_formula = Option::<String>::None;
    let mut in_cell = false;
    let mut in_value = false;
    let mut in_formula = false;
    let mut in_merge_cells = false;
    let mut in_cols = false;

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) => {
                let tag = e.name();
                let tag = tag.as_ref();
                match tag {
                    b"cols" => {
                        in_cols = true;
                    }
                    b"row" => {
                        current_row_height = None;
                        current_row_hidden = false;
                        for attr in e.attributes().filter_map(|a| a.ok()) {
                            match attr.key.as_ref() {
                                b"r" => {
                                    current_row =
                                        String::from_utf8_lossy(&attr.value).parse().unwrap_or(0);
                                }
                                b"ht" => {
                                    current_row_height =
                                        String::from_utf8_lossy(&attr.value).parse().ok();
                                }
                                b"hidden" => {
                                    let val = String::from_utf8_lossy(&attr.value);
                                    current_row_hidden = val == "1" || val == "true";
                                }
                                _ => {}
                            }
                        }
                    }
                    b"c" => {
                        in_cell = true;
                        current_cell_ref.clear();
                        current_cell_type.clear();
                        current_cell_style = None;
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
                                b"s" => {
                                    current_cell_style =
                                        String::from_utf8_lossy(&attr.value).parse().ok();
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
                    b"mergeCells" => {
                        in_merge_cells = true;
                    }
                    _ => {}
                }
            }
            Ok(Event::Empty(e)) => {
                let tag = e.name();
                let tag_ref = tag.as_ref();

                // Handle empty cells <c r="A1"/>
                if tag_ref == b"c" {
                    let mut cell_ref = String::new();
                    let mut cell_style: Option<u32> = None;
                    for attr in e.attributes().filter_map(|a| a.ok()) {
                        match attr.key.as_ref() {
                            b"r" => {
                                cell_ref = String::from_utf8_lossy(&attr.value).into_owned();
                            }
                            b"s" => {
                                cell_style = String::from_utf8_lossy(&attr.value).parse().ok();
                            }
                            _ => {}
                        }
                    }
                    if !cell_ref.is_empty() {
                        let column = parse_cell_reference(&cell_ref).map(|(c, _)| c).unwrap_or(1);
                        let cell = Cell {
                            reference: cell_ref,
                            column,
                            value: CellValue::Empty,
                            formula: None,
                            style_index: cell_style,
                        };
                        rows.entry(current_row)
                            .or_insert_with(|| (Vec::new(), current_row_height, current_row_hidden))
                            .0
                            .push(cell);
                    }
                }
                // Handle column <col min="1" max="1" width="15"/>
                else if in_cols && tag_ref == b"col" {
                    let mut col_info = ColumnInfo {
                        min: 1,
                        max: 1,
                        width: None,
                        hidden: false,
                        custom_width: false,
                        style_index: None,
                    };
                    for attr in e.attributes().filter_map(|a| a.ok()) {
                        let val = String::from_utf8_lossy(&attr.value);
                        match attr.key.as_ref() {
                            b"min" => col_info.min = val.parse().unwrap_or(1),
                            b"max" => col_info.max = val.parse().unwrap_or(1),
                            b"width" => col_info.width = val.parse().ok(),
                            b"hidden" => col_info.hidden = val == "1" || val == "true",
                            b"customWidth" => col_info.custom_width = val == "1" || val == "true",
                            b"style" => col_info.style_index = val.parse().ok(),
                            _ => {}
                        }
                    }
                    columns.push(col_info);
                }
                // Handle merged cell <mergeCell ref="A1:B2"/>
                else if in_merge_cells && tag_ref == b"mergeCell" {
                    for attr in e.attributes().filter_map(|a| a.ok()) {
                        if attr.key.as_ref() == b"ref" {
                            let range_ref = String::from_utf8_lossy(&attr.value).into_owned();
                            if let Some(mc) = parse_merge_cell_range(&range_ref) {
                                merged_cells.push(mc);
                            }
                        }
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
                            style_index: current_cell_style.take(),
                        };

                        rows.entry(current_row)
                            .or_insert_with(|| (Vec::new(), current_row_height, current_row_hidden))
                            .0
                            .push(cell);
                        in_cell = false;
                    }
                    b"v" => in_value = false,
                    b"f" => in_formula = false,
                    b"cols" => in_cols = false,
                    b"mergeCells" => in_merge_cells = false,
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
        .map(|(row_num, (cells, height, hidden))| Row {
            row_num,
            cells,
            height,
            hidden,
        })
        .collect();
    row_list.sort_by_key(|r| r.row_num);

    // Sort cells within each row by column
    for row in &mut row_list {
        row.cells.sort_by_key(|c| c.column);
    }

    Ok(Sheet {
        name: name.to_string(),
        rows: row_list,
        merged_cells,
        columns,
        comments: Vec::new(),
    })
}

/// Parse comments from a comments XML file.
///
/// ECMA-376 Part 1, Section 18.7 (Comments).
fn parse_comments(xml: &[u8]) -> Result<Vec<Comment>> {
    let mut reader = Reader::from_reader(Cursor::new(xml));
    let mut buf = Vec::new();
    let mut comments = Vec::new();
    let mut authors: Vec<String> = Vec::new();

    // Parsing state
    let mut in_authors = false;
    let mut in_author = false;
    let mut in_comment_list = false;
    let mut in_comment = false;
    let mut in_text = false;
    let mut in_t = false;

    let mut current_ref = String::new();
    let mut current_author_id: Option<usize> = None;
    let mut current_text = String::new();
    let mut author_text = String::new();

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) => {
                let name = e.name();
                let name = name.as_ref();
                match name {
                    b"authors" => in_authors = true,
                    b"author" => {
                        in_author = true;
                        author_text.clear();
                    }
                    b"commentList" => in_comment_list = true,
                    b"comment" if in_comment_list => {
                        in_comment = true;
                        current_ref.clear();
                        current_author_id = None;
                        current_text.clear();

                        for attr in e.attributes().filter_map(|a| a.ok()) {
                            match attr.key.as_ref() {
                                b"ref" => {
                                    current_ref = String::from_utf8_lossy(&attr.value).into_owned();
                                }
                                b"authorId" => {
                                    current_author_id =
                                        String::from_utf8_lossy(&attr.value).parse().ok();
                                }
                                _ => {}
                            }
                        }
                    }
                    b"text" if in_comment => in_text = true,
                    b"t" if in_text => in_t = true,
                    b"r" if in_text => {
                        // Text run inside comment text
                    }
                    _ => {}
                }
            }
            Ok(Event::Text(e)) => {
                let text = e.decode().unwrap_or_default();
                if in_author {
                    author_text.push_str(&text);
                } else if in_t {
                    current_text.push_str(&text);
                }
            }
            Ok(Event::End(e)) => {
                let name = e.name();
                let name = name.as_ref();
                match name {
                    b"authors" => in_authors = false,
                    b"author" => {
                        in_author = false;
                        if in_authors {
                            authors.push(std::mem::take(&mut author_text));
                        }
                    }
                    b"commentList" => in_comment_list = false,
                    b"comment" if in_comment => {
                        in_comment = false;
                        if !current_ref.is_empty() {
                            let author = current_author_id.and_then(|id| authors.get(id).cloned());
                            comments.push(Comment {
                                reference: std::mem::take(&mut current_ref),
                                author,
                                text: std::mem::take(&mut current_text),
                            });
                        }
                    }
                    b"text" if in_text => in_text = false,
                    b"t" if in_t => in_t = false,
                    _ => {}
                }
            }
            Ok(Event::Eof) => break,
            Err(e) => return Err(Error::Xml(e)),
            _ => {}
        }
        buf.clear();
    }

    Ok(comments)
}

// ============================================================================
// Utilities
// ============================================================================

/// Convert a 1-based column number to letters.
/// E.g., 1 -> "A", 2 -> "B", 26 -> "Z", 27 -> "AA"
fn column_number_to_letter(mut col: u32) -> String {
    let mut result = String::new();
    while col > 0 {
        col -= 1;
        result.insert(0, (b'A' + (col % 26) as u8) as char);
        col /= 26;
    }
    result
}

/// Parse a merge cell range like "A1:B2" into a MergedCell.
fn parse_merge_cell_range(range: &str) -> Option<MergedCell> {
    let parts: Vec<&str> = range.split(':').collect();
    if parts.len() != 2 {
        return None;
    }

    let (start_col, start_row) = parse_cell_reference(parts[0])?;
    let (end_col, end_row) = parse_cell_reference(parts[1])?;

    Some(MergedCell {
        reference: range.to_string(),
        start_row,
        start_col,
        end_row,
        end_col,
    })
}

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

    #[test]
    fn test_parse_styles() {
        let xml = r#"<?xml version="1.0"?>
        <styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
            <numFmts count="1">
                <numFmt numFmtId="164" formatCode="0.00%"/>
            </numFmts>
            <fonts count="2">
                <font>
                    <name val="Calibri"/>
                    <sz val="11"/>
                </font>
                <font>
                    <b/>
                    <name val="Arial"/>
                    <sz val="14"/>
                    <color rgb="FF0000FF"/>
                </font>
            </fonts>
            <fills count="2">
                <fill>
                    <patternFill patternType="none"/>
                </fill>
                <fill>
                    <patternFill patternType="solid">
                        <fgColor rgb="FFFFFF00"/>
                    </patternFill>
                </fill>
            </fills>
            <borders count="1">
                <border>
                    <left style="thin"/>
                    <right/>
                    <top/>
                    <bottom/>
                </border>
            </borders>
            <cellXfs count="2">
                <xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>
                <xf numFmtId="164" fontId="1" fillId="1" borderId="0"/>
            </cellXfs>
        </styleSheet>"#;

        let styles = parse_styles(xml.as_bytes()).unwrap();

        // Check number formats
        assert_eq!(styles.number_formats.len(), 1);
        assert_eq!(styles.number_formats[0].id, 164);
        assert_eq!(styles.number_formats[0].code, "0.00%");

        // Check fonts
        assert_eq!(styles.fonts.len(), 2);
        assert_eq!(styles.fonts[0].name, Some("Calibri".to_string()));
        assert_eq!(styles.fonts[0].size, Some(11.0));
        assert!(!styles.fonts[0].bold);
        assert_eq!(styles.fonts[1].name, Some("Arial".to_string()));
        assert!(styles.fonts[1].bold);
        assert_eq!(styles.fonts[1].color, Some("FF0000FF".to_string()));

        // Check fills
        assert_eq!(styles.fills.len(), 2);
        assert_eq!(styles.fills[0].pattern_type, Some("none".to_string()));
        assert_eq!(styles.fills[1].pattern_type, Some("solid".to_string()));
        assert_eq!(styles.fills[1].fg_color, Some("FFFFFF00".to_string()));

        // Check borders
        assert_eq!(styles.borders.len(), 1);
        assert!(styles.borders[0].left.is_some());
        assert_eq!(
            styles.borders[0].left.as_ref().unwrap().style,
            Some("thin".to_string())
        );

        // Check cell formats
        assert_eq!(styles.cell_formats.len(), 2);
        assert_eq!(styles.cell_formats[0].font_id, 0);
        assert_eq!(styles.cell_formats[1].number_format_id, 164);
        assert_eq!(styles.cell_formats[1].font_id, 1);
        assert_eq!(styles.cell_formats[1].fill_id, 1);
    }

    #[test]
    fn test_parse_comments() {
        let xml = r#"<?xml version="1.0"?>
        <comments xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
            <authors>
                <author>John Doe</author>
                <author>Jane Smith</author>
            </authors>
            <commentList>
                <comment ref="A1" authorId="0">
                    <text><t>This is a comment on A1</t></text>
                </comment>
                <comment ref="B2" authorId="1">
                    <text>
                        <r><t>Multi-line</t></r>
                        <r><t> comment</t></r>
                    </text>
                </comment>
            </commentList>
        </comments>"#;

        let comments = parse_comments(xml.as_bytes()).unwrap();
        assert_eq!(comments.len(), 2);

        // First comment
        assert_eq!(comments[0].reference(), "A1");
        assert_eq!(comments[0].author(), Some("John Doe"));
        assert_eq!(comments[0].text(), "This is a comment on A1");

        // Second comment (with multiple text runs)
        assert_eq!(comments[1].reference(), "B2");
        assert_eq!(comments[1].author(), Some("Jane Smith"));
        assert_eq!(comments[1].text(), "Multi-line comment");
    }
}
