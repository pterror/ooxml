//! Extension traits for generated OOXML types.
//!
//! This module provides convenience methods for the generated types via extension traits.
//! See ADR-003 for the architectural rationale.
//!
//! # Design
//!
//! Extension traits are split into two categories:
//!
//! - **Pure traits** (`CellExt`, `RowExt`): Methods that don't need external context
//! - **Resolve traits** (`CellResolveExt`): Methods that need `ResolveContext` for
//!   shared strings, styles, etc.
//!
//! # Example
//!
//! ```ignore
//! use ooxml_sml::ext::{CellExt, CellResolveExt, ResolveContext};
//! use ooxml_sml::types::Cell;
//!
//! let cell: &Cell = /* ... */;
//!
//! // Pure methods - no context needed
//! let col = cell.column_number();
//! let row = cell.row_number();
//!
//! // Resolved methods - context required
//! let ctx = ResolveContext::new(shared_strings, stylesheet);
//! let value = cell.value_as_string(&ctx);
//! ```

use crate::parsers::{FromXml, ParseError};
use crate::types::{Cell, CellType, Row, SheetData, Worksheet};
use quick_xml::Reader;
use quick_xml::events::Event;
use std::io::Cursor;

/// Resolved cell value (typed).
#[derive(Debug, Clone, PartialEq)]
pub enum CellValue {
    /// Empty cell
    Empty,
    /// String value (from shared strings or inline)
    String(String),
    /// Numeric value
    Number(f64),
    /// Boolean value
    Boolean(bool),
    /// Error value (e.g., "#REF!", "#VALUE!")
    Error(String),
}

impl CellValue {
    /// Check if the value is empty.
    pub fn is_empty(&self) -> bool {
        matches!(self, CellValue::Empty)
    }

    /// Get as string for display.
    pub fn to_display_string(&self) -> String {
        match self {
            CellValue::Empty => String::new(),
            CellValue::String(s) => s.clone(),
            CellValue::Number(n) => n.to_string(),
            CellValue::Boolean(b) => if *b { "TRUE" } else { "FALSE" }.to_string(),
            CellValue::Error(e) => e.clone(),
        }
    }

    /// Try to get as number.
    pub fn as_number(&self) -> Option<f64> {
        match self {
            CellValue::Number(n) => Some(*n),
            CellValue::String(s) => s.parse().ok(),
            _ => None,
        }
    }

    /// Try to get as boolean.
    pub fn as_bool(&self) -> Option<bool> {
        match self {
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

/// Context for resolving cell values.
///
/// Contains shared strings table and stylesheet needed to convert
/// raw XML values into typed `CellValue`s.
#[derive(Debug, Clone, Default)]
pub struct ResolveContext {
    /// Shared string table (index -> string)
    pub shared_strings: Vec<String>,
    // Future: stylesheet, themes, etc.
}

impl ResolveContext {
    /// Create a new resolve context.
    pub fn new(shared_strings: Vec<String>) -> Self {
        Self { shared_strings }
    }

    /// Get a shared string by index.
    pub fn shared_string(&self, index: usize) -> Option<&str> {
        self.shared_strings.get(index).map(|s| s.as_str())
    }
}

// =============================================================================
// Cell Extension Traits
// =============================================================================

/// Pure extension methods for `Cell` (no context needed).
pub trait CellExt {
    /// Get the cell reference string (e.g., "A1", "B5").
    fn reference_str(&self) -> Option<&str>;

    /// Parse column number from reference (1-based, e.g., "B5" -> 2).
    fn column_number(&self) -> Option<u32>;

    /// Parse row number from reference (1-based, e.g., "B5" -> 5).
    fn row_number(&self) -> Option<u32>;

    /// Check if cell has a formula.
    fn has_formula(&self) -> bool;

    /// Get the formula text (if any).
    fn formula_text(&self) -> Option<&str>;

    /// Get the raw value string (before resolution).
    fn raw_value(&self) -> Option<&str>;

    /// Get the cell type.
    fn cell_type(&self) -> Option<CellType>;

    /// Check if this is a shared string cell.
    fn is_shared_string(&self) -> bool;

    /// Check if this is a number cell.
    fn is_number(&self) -> bool;

    /// Check if this is a boolean cell.
    fn is_boolean(&self) -> bool;

    /// Check if this is an error cell.
    fn is_error(&self) -> bool;
}

impl CellExt for Cell {
    fn reference_str(&self) -> Option<&str> {
        self.reference.as_deref()
    }

    fn column_number(&self) -> Option<u32> {
        let reference = self.reference.as_ref()?;
        parse_column(reference)
    }

    fn row_number(&self) -> Option<u32> {
        let reference = self.reference.as_ref()?;
        parse_row(reference)
    }

    fn has_formula(&self) -> bool {
        self.formula.is_some()
    }

    fn formula_text(&self) -> Option<&str> {
        // TODO: CellFormula text content not yet captured by codegen
        // For now, return None - formula presence can be checked with has_formula()
        self.formula.as_ref().map(|_| "(formula)" as &str)
    }

    fn raw_value(&self) -> Option<&str> {
        self.value.as_deref()
    }

    fn cell_type(&self) -> Option<CellType> {
        self.cell_type
    }

    fn is_shared_string(&self) -> bool {
        matches!(self.cell_type, Some(CellType::S))
    }

    fn is_number(&self) -> bool {
        matches!(self.cell_type, Some(CellType::N)) || self.cell_type.is_none()
    }

    fn is_boolean(&self) -> bool {
        matches!(self.cell_type, Some(CellType::B))
    }

    fn is_error(&self) -> bool {
        matches!(self.cell_type, Some(CellType::E))
    }
}

/// Extension methods for `Cell` that require resolution context.
pub trait CellResolveExt {
    /// Resolve the cell value to a typed `CellValue`.
    fn resolved_value(&self, ctx: &ResolveContext) -> CellValue;

    /// Get value as display string.
    fn value_as_string(&self, ctx: &ResolveContext) -> String;

    /// Try to get value as number.
    fn value_as_number(&self, ctx: &ResolveContext) -> Option<f64>;

    /// Try to get value as boolean.
    fn value_as_bool(&self, ctx: &ResolveContext) -> Option<bool>;
}

impl CellResolveExt for Cell {
    fn resolved_value(&self, ctx: &ResolveContext) -> CellValue {
        let raw = match &self.value {
            Some(v) => v.as_str(),
            None => return CellValue::Empty,
        };

        match &self.cell_type {
            Some(CellType::S) => {
                // Shared string - raw value is index
                if let Ok(idx) = raw.parse::<usize>()
                    && let Some(s) = ctx.shared_string(idx)
                {
                    return CellValue::String(s.to_string());
                }
                CellValue::Error(format!("#REF! (invalid shared string index: {})", raw))
            }
            Some(CellType::B) => {
                // Boolean
                CellValue::Boolean(raw == "1" || raw.eq_ignore_ascii_case("true"))
            }
            Some(CellType::E) => {
                // Error
                CellValue::Error(raw.to_string())
            }
            Some(CellType::Str) | Some(CellType::InlineStr) => {
                // Inline string
                CellValue::String(raw.to_string())
            }
            Some(CellType::N) | None => {
                // Number (or default, which is number)
                if raw.is_empty() {
                    CellValue::Empty
                } else if let Ok(n) = raw.parse::<f64>() {
                    CellValue::Number(n)
                } else {
                    // Fallback to string if not a valid number
                    CellValue::String(raw.to_string())
                }
            }
        }
    }

    fn value_as_string(&self, ctx: &ResolveContext) -> String {
        self.resolved_value(ctx).to_display_string()
    }

    fn value_as_number(&self, ctx: &ResolveContext) -> Option<f64> {
        self.resolved_value(ctx).as_number()
    }

    fn value_as_bool(&self, ctx: &ResolveContext) -> Option<bool> {
        self.resolved_value(ctx).as_bool()
    }
}

// =============================================================================
// Row Extension Traits
// =============================================================================

/// Pure extension methods for `Row` (no context needed).
pub trait RowExt {
    /// Get the 1-based row number.
    fn row_number(&self) -> Option<u32>;

    /// Get the number of cells in this row.
    fn cell_count(&self) -> usize;

    /// Check if row is empty (no cells).
    fn is_empty(&self) -> bool;

    /// Get a cell by column number (1-based).
    fn cell_at_column(&self, col: u32) -> Option<&Cell>;

    /// Iterate over cells.
    fn cells_iter(&self) -> impl Iterator<Item = &Cell>;
}

impl RowExt for Row {
    fn row_number(&self) -> Option<u32> {
        self.reference
    }

    fn cell_count(&self) -> usize {
        self.cells.len()
    }

    fn is_empty(&self) -> bool {
        self.cells.is_empty()
    }

    fn cell_at_column(&self, col: u32) -> Option<&Cell> {
        self.cells
            .iter()
            .find(|c| {
                c.reference
                    .as_ref()
                    .and_then(|r| parse_column(r))
                    .map(|c_col| c_col == col)
                    .unwrap_or(false)
            })
            .map(|c| c.as_ref())
    }

    fn cells_iter(&self) -> impl Iterator<Item = &Cell> {
        self.cells.iter().map(|c| c.as_ref())
    }
}

// =============================================================================
// Worksheet Parsing
// =============================================================================

/// Parse a worksheet from XML bytes using the generated FromXml parser.
///
/// This is the recommended way to parse worksheet XML, as it uses the
/// spec-compliant generated types and is faster than serde.
pub fn parse_worksheet(xml: &[u8]) -> Result<Worksheet, ParseError> {
    let mut reader = Reader::from_reader(Cursor::new(xml));
    let mut buf = Vec::new();

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) => return Worksheet::from_xml(&mut reader, &e, false),
            Ok(Event::Empty(e)) => return Worksheet::from_xml(&mut reader, &e, true),
            Ok(Event::Eof) => break,
            Err(e) => return Err(ParseError::Xml(e)),
            _ => {}
        }
        buf.clear();
    }
    Err(ParseError::UnexpectedElement(
        "no worksheet element found".to_string(),
    ))
}

// =============================================================================
// Worksheet Extension Traits
// =============================================================================

/// Pure extension methods for `Worksheet` (no context needed).
pub trait WorksheetExt {
    /// Get the sheet data (rows and cells).
    fn sheet_data(&self) -> &SheetData;

    /// Get the number of rows.
    fn row_count(&self) -> usize;

    /// Check if the worksheet is empty.
    fn is_empty(&self) -> bool;

    /// Get a row by 1-based row number.
    fn row(&self, row_num: u32) -> Option<&Row>;

    /// Get a cell by reference (e.g., "A1", "B5").
    fn cell(&self, reference: &str) -> Option<&Cell>;

    /// Iterate over all rows.
    fn rows(&self) -> impl Iterator<Item = &Row>;

    /// Check if the worksheet has an auto-filter.
    fn has_auto_filter(&self) -> bool;

    /// Check if the worksheet has merged cells.
    fn has_merged_cells(&self) -> bool;

    /// Check if the worksheet has conditional formatting.
    fn has_conditional_formatting(&self) -> bool;

    /// Check if the worksheet has data validations.
    fn has_data_validations(&self) -> bool;

    /// Check if the worksheet has freeze panes.
    fn has_freeze_panes(&self) -> bool;
}

impl WorksheetExt for Worksheet {
    fn sheet_data(&self) -> &SheetData {
        &self.sheet_data
    }

    fn row_count(&self) -> usize {
        self.sheet_data.row.len()
    }

    fn is_empty(&self) -> bool {
        self.sheet_data.row.is_empty()
    }

    fn row(&self, row_num: u32) -> Option<&Row> {
        self.sheet_data
            .row
            .iter()
            .find(|r| r.reference == Some(row_num))
            .map(|r| r.as_ref())
    }

    fn cell(&self, reference: &str) -> Option<&Cell> {
        let col = parse_column(reference)?;
        let row_num = parse_row(reference)?;
        let row = self.row(row_num)?;
        row.cells
            .iter()
            .find(|c| {
                c.reference
                    .as_ref()
                    .and_then(|r| parse_column(r))
                    .map(|c_col| c_col == col)
                    .unwrap_or(false)
            })
            .map(|c| c.as_ref())
    }

    fn rows(&self) -> impl Iterator<Item = &Row> {
        self.sheet_data.row.iter().map(|r| r.as_ref())
    }

    fn has_auto_filter(&self) -> bool {
        self.auto_filter.is_some()
    }

    fn has_merged_cells(&self) -> bool {
        self.merged_cells.is_some()
    }

    fn has_conditional_formatting(&self) -> bool {
        !self.conditional_formatting.is_empty()
    }

    fn has_data_validations(&self) -> bool {
        self.data_validations.is_some()
    }

    fn has_freeze_panes(&self) -> bool {
        // Check if any sheet view has a pane with frozen state
        self.sheet_views
            .as_ref()
            .is_some_and(|views| views.sheet_view.iter().any(|sv| sv.pane.is_some()))
    }
}

/// Extension methods for `SheetData`.
pub trait SheetDataExt {
    /// Get a row by 1-based row number.
    fn row(&self, row_num: u32) -> Option<&Row>;

    /// Iterate over rows.
    fn rows(&self) -> impl Iterator<Item = &Row>;
}

impl SheetDataExt for SheetData {
    fn row(&self, row_num: u32) -> Option<&Row> {
        self.row
            .iter()
            .find(|r| r.reference == Some(row_num))
            .map(|r| r.as_ref())
    }

    fn rows(&self) -> impl Iterator<Item = &Row> {
        self.row.iter().map(|r| r.as_ref())
    }
}

// =============================================================================
// Helpers
// =============================================================================

/// Parse column letters from a cell reference (e.g., "AB5" -> 28).
fn parse_column(reference: &str) -> Option<u32> {
    let mut col: u32 = 0;
    for ch in reference.chars() {
        if ch.is_ascii_alphabetic() {
            col = col * 26 + (ch.to_ascii_uppercase() as u32 - 'A' as u32 + 1);
        } else {
            break;
        }
    }
    if col > 0 { Some(col) } else { None }
}

/// Parse row number from a cell reference (e.g., "AB5" -> 5).
fn parse_row(reference: &str) -> Option<u32> {
    let digits: String = reference.chars().filter(|c| c.is_ascii_digit()).collect();
    digits.parse().ok()
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_parse_column() {
        assert_eq!(parse_column("A1"), Some(1));
        assert_eq!(parse_column("B5"), Some(2));
        assert_eq!(parse_column("Z1"), Some(26));
        assert_eq!(parse_column("AA1"), Some(27));
        assert_eq!(parse_column("AB1"), Some(28));
        assert_eq!(parse_column("AZ1"), Some(52));
        assert_eq!(parse_column("BA1"), Some(53));
    }

    #[test]
    fn test_parse_row() {
        assert_eq!(parse_row("A1"), Some(1));
        assert_eq!(parse_row("B5"), Some(5));
        assert_eq!(parse_row("AA100"), Some(100));
        assert_eq!(parse_row("ZZ9999"), Some(9999));
    }

    #[test]
    fn test_cell_ext() {
        let cell = Cell {
            reference: Some("B5".to_string()),
            cell_type: Some(CellType::N),
            value: Some("42.5".to_string()),
            formula: None,
            style_index: None,
            cm: None,
            vm: None,
            placeholder: None,
            is: None,
            extension_list: None,
        };

        assert_eq!(cell.column_number(), Some(2));
        assert_eq!(cell.row_number(), Some(5));
        assert!(!cell.has_formula());
        assert!(cell.is_number());
        assert!(!cell.is_shared_string());
    }

    #[test]
    fn test_cell_resolve_number() {
        let cell = Cell {
            reference: Some("A1".to_string()),
            cell_type: Some(CellType::N),
            value: Some("123.45".to_string()),
            formula: None,
            style_index: None,
            cm: None,
            vm: None,
            placeholder: None,
            is: None,
            extension_list: None,
        };

        let ctx = ResolveContext::default();
        assert_eq!(cell.resolved_value(&ctx), CellValue::Number(123.45));
        assert_eq!(cell.value_as_string(&ctx), "123.45");
        assert_eq!(cell.value_as_number(&ctx), Some(123.45));
    }

    #[test]
    fn test_cell_resolve_shared_string() {
        let cell = Cell {
            reference: Some("A1".to_string()),
            cell_type: Some(CellType::S),
            value: Some("0".to_string()), // Index into shared strings
            formula: None,
            style_index: None,
            cm: None,
            vm: None,
            placeholder: None,
            is: None,
            extension_list: None,
        };

        let ctx = ResolveContext::new(vec!["Hello".to_string(), "World".to_string()]);
        assert_eq!(
            cell.resolved_value(&ctx),
            CellValue::String("Hello".to_string())
        );
        assert_eq!(cell.value_as_string(&ctx), "Hello");
    }

    #[test]
    fn test_cell_resolve_boolean() {
        let cell = Cell {
            reference: Some("A1".to_string()),
            cell_type: Some(CellType::B),
            value: Some("1".to_string()),
            formula: None,
            style_index: None,
            cm: None,
            vm: None,
            placeholder: None,
            is: None,
            extension_list: None,
        };

        let ctx = ResolveContext::default();
        assert_eq!(cell.resolved_value(&ctx), CellValue::Boolean(true));
        assert_eq!(cell.value_as_string(&ctx), "TRUE");
        assert_eq!(cell.value_as_bool(&ctx), Some(true));
    }

    #[test]
    fn test_parse_worksheet() {
        let xml = br#"<?xml version="1.0" encoding="UTF-8"?>
        <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
            <sheetData>
                <row r="1" spans="1:3">
                    <c r="A1" t="s"><v>0</v></c>
                    <c r="B1"><v>42.5</v></c>
                    <c r="C1" t="b"><v>1</v></c>
                </row>
                <row r="2">
                    <c r="A2"><v>100</v></c>
                </row>
            </sheetData>
        </worksheet>"#;

        let worksheet = parse_worksheet(xml).expect("parse failed");

        assert_eq!(worksheet.row_count(), 2);
        assert!(!worksheet.is_empty());

        // Test row access
        let row1 = worksheet.row(1).expect("row 1 should exist");
        assert_eq!(row1.cells.len(), 3);

        let row2 = worksheet.row(2).expect("row 2 should exist");
        assert_eq!(row2.cells.len(), 1);

        // Test cell access by reference
        let cell_a1 = worksheet.cell("A1").expect("A1 should exist");
        assert_eq!(cell_a1.value.as_deref(), Some("0"));
        assert!(cell_a1.is_shared_string());

        let cell_b1 = worksheet.cell("B1").expect("B1 should exist");
        assert_eq!(cell_b1.value.as_deref(), Some("42.5"));
        assert!(cell_b1.is_number());

        // Test non-existent cell
        assert!(worksheet.cell("Z99").is_none());
    }

    #[test]
    fn test_worksheet_ext_with_resolve() {
        let xml = br#"<?xml version="1.0" encoding="UTF-8"?>
        <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
            <sheetData>
                <row r="1">
                    <c r="A1" t="s"><v>0</v></c>
                    <c r="B1" t="s"><v>1</v></c>
                </row>
            </sheetData>
        </worksheet>"#;

        let worksheet = parse_worksheet(xml).expect("parse failed");
        let ctx = ResolveContext::new(vec!["Hello".to_string(), "World".to_string()]);

        let cell_a1 = worksheet.cell("A1").expect("A1 should exist");
        assert_eq!(cell_a1.value_as_string(&ctx), "Hello");

        let cell_b1 = worksheet.cell("B1").expect("B1 should exist");
        assert_eq!(cell_b1.value_as_string(&ctx), "World");
    }
}
