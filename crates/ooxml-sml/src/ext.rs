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
        self.formula.as_ref().and_then(|f| f.text.as_deref())
    }

    fn raw_value(&self) -> Option<&str> {
        self.value.as_deref()
    }

    fn cell_type(&self) -> Option<CellType> {
        self.cell_type
    }

    fn is_shared_string(&self) -> bool {
        matches!(self.cell_type, Some(CellType::SharedString))
    }

    fn is_number(&self) -> bool {
        matches!(self.cell_type, Some(CellType::Number)) || self.cell_type.is_none()
    }

    fn is_boolean(&self) -> bool {
        matches!(self.cell_type, Some(CellType::Boolean))
    }

    fn is_error(&self) -> bool {
        matches!(self.cell_type, Some(CellType::Error))
    }
}

/// Extension methods for `Cell` that require a [`ResolveContext`] to dereference
/// shared strings and interpret cell types.
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
            Some(CellType::SharedString) => {
                // Shared string - raw value is index
                if let Ok(idx) = raw.parse::<usize>()
                    && let Some(s) = ctx.shared_string(idx)
                {
                    return CellValue::String(s.to_string());
                }
                CellValue::Error(format!("#REF! (invalid shared string index: {})", raw))
            }
            Some(CellType::Boolean) => {
                // Boolean
                CellValue::Boolean(raw == "1" || raw.eq_ignore_ascii_case("true"))
            }
            Some(CellType::Error) => {
                // Error
                CellValue::Error(raw.to_string())
            }
            Some(CellType::String) | Some(CellType::InlineString) => {
                // Inline string
                CellValue::String(raw.to_string())
            }
            Some(CellType::Number) | None => {
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

/// Pure extension methods for `Row` that do not require external context.
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
        self.cells.iter().find(|c| {
            c.reference
                .as_ref()
                .and_then(|r| parse_column(r))
                .map(|c_col| c_col == col)
                .unwrap_or(false)
        })
    }

    fn cells_iter(&self) -> impl Iterator<Item = &Cell> {
        self.cells.iter()
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
    #[cfg(feature = "sml-filtering")]
    fn has_auto_filter(&self) -> bool;

    /// Check if the worksheet has merged cells.
    fn has_merged_cells(&self) -> bool;

    /// Check if the worksheet has conditional formatting.
    #[cfg(feature = "sml-styling")]
    fn has_conditional_formatting(&self) -> bool;

    /// Check if the worksheet has data validations.
    #[cfg(feature = "sml-validation")]
    fn has_data_validations(&self) -> bool;

    /// Check if the worksheet has freeze panes.
    #[cfg(feature = "sml-structure")]
    fn has_freeze_panes(&self) -> bool;

    /// Get the height of a row in points (if specified).
    ///
    /// `row_num` is 1-based. Returns `None` if the row does not exist or has no
    /// explicit height set. ECMA-376 Part 1, §18.3.1.73 (`@ht` attribute).
    #[cfg(feature = "sml-styling")]
    fn get_row_height(&self, row_num: u32) -> Option<f64>;

    /// Get the width of a column in characters (if specified).
    ///
    /// `col_idx` is 1-based (A=1, B=2, …). Returns `None` if no column definition
    /// covers this index or if the column has no explicit width.
    /// ECMA-376 Part 1, §18.3.1.13 (`@width` attribute).
    #[cfg(feature = "sml-styling")]
    fn get_column_width(&self, col_idx: u32) -> Option<f64>;
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
    }

    fn cell(&self, reference: &str) -> Option<&Cell> {
        let col = parse_column(reference)?;
        let row_num = parse_row(reference)?;
        let row = self.row(row_num)?;
        row.cells.iter().find(|c| {
            c.reference
                .as_ref()
                .and_then(|r| parse_column(r))
                .map(|c_col| c_col == col)
                .unwrap_or(false)
        })
    }

    fn rows(&self) -> impl Iterator<Item = &Row> {
        self.sheet_data.row.iter()
    }

    #[cfg(feature = "sml-filtering")]
    fn has_auto_filter(&self) -> bool {
        self.auto_filter.is_some()
    }

    fn has_merged_cells(&self) -> bool {
        self.merged_cells.is_some()
    }

    #[cfg(feature = "sml-styling")]
    fn has_conditional_formatting(&self) -> bool {
        !self.conditional_formatting.is_empty()
    }

    #[cfg(feature = "sml-validation")]
    fn has_data_validations(&self) -> bool {
        self.data_validations.is_some()
    }

    #[cfg(feature = "sml-structure")]
    fn has_freeze_panes(&self) -> bool {
        // Check if any sheet view has a pane with frozen state
        self.sheet_views
            .as_ref()
            .is_some_and(|views| views.sheet_view.iter().any(|sv| sv.pane.is_some()))
    }

    #[cfg(feature = "sml-styling")]
    fn get_row_height(&self, row_num: u32) -> Option<f64> {
        self.sheet_data
            .row
            .iter()
            .find(|r| r.reference == Some(row_num))
            .and_then(|r| r.height)
    }

    #[cfg(feature = "sml-styling")]
    fn get_column_width(&self, col_idx: u32) -> Option<f64> {
        self.cols
            .iter()
            .flat_map(|cols| &cols.col)
            .find(|c| c.start_column <= col_idx && col_idx <= c.end_column)
            .and_then(|c| c.width)
    }
}

/// Extension methods for `SheetData`, providing row lookup and iteration.
pub trait SheetDataExt {
    /// Get a row by 1-based row number.
    fn row(&self, row_num: u32) -> Option<&Row>;

    /// Iterate over rows.
    fn rows(&self) -> impl Iterator<Item = &Row>;
}

impl SheetDataExt for SheetData {
    fn row(&self, row_num: u32) -> Option<&Row> {
        self.row.iter().find(|r| r.reference == Some(row_num))
    }

    fn rows(&self) -> impl Iterator<Item = &Row> {
        self.row.iter()
    }
}

// =============================================================================
// ResolvedSheet - High-level wrapper with automatic value resolution
// =============================================================================

/// A worksheet with bound resolution context for convenient value access.
///
/// This is the high-level API for reading worksheets. It wraps a generated
/// `types::Worksheet` and provides methods that automatically resolve values
/// using the shared string table.
///
/// # Example
///
/// ```ignore
/// let sheet = ResolvedSheet::new(name, worksheet, shared_strings);
///
/// // Iterate rows and get resolved values
/// for row in sheet.rows() {
///     for cell in row.cells_iter() {
///         println!("{}", sheet.cell_value_string(cell));
///     }
/// }
///
/// // Direct cell access
/// if let Some(cell) = sheet.cell("A1") {
///     println!("A1 = {}", sheet.cell_value_string(cell));
/// }
/// ```
#[derive(Debug, Clone)]
pub struct ResolvedSheet {
    /// Sheet name
    name: String,
    /// The underlying worksheet data (generated type)
    worksheet: Worksheet,
    /// Resolution context for shared strings
    context: ResolveContext,
    /// Comments (loaded separately from comments.xml)
    comments: Vec<Comment>,
    /// Charts (loaded separately via drawing relationships)
    charts: Vec<Chart>,
    /// Pivot tables (loaded from sheet relationships)
    #[cfg(feature = "sml-pivot")]
    pivot_tables: Vec<crate::types::CTPivotTableDefinition>,
}

/// A comment (note) attached to a cell, as returned by [`ResolvedSheet::comments`].
#[derive(Debug, Clone)]
pub struct Comment {
    /// Cell reference (e.g., "A1")
    pub reference: String,
    /// Comment author (if available)
    pub author: Option<String>,
    /// Comment text
    pub text: String,
}

/// A chart embedded in the worksheet, as returned by [`ResolvedSheet::charts`].
#[derive(Debug, Clone)]
pub struct Chart {
    /// Chart title (if available)
    pub title: Option<String>,
    /// Chart type
    pub chart_type: ChartType,
}

/// The type of an embedded chart as exposed by [`ResolvedSheet`].
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ChartType {
    /// Bar chart (horizontal bars).
    Bar,
    /// Column chart (vertical bars).
    Column,
    /// Line chart.
    Line,
    /// Pie chart.
    Pie,
    /// Area chart.
    Area,
    /// Scatter (XY) chart.
    Scatter,
    /// Doughnut chart.
    Doughnut,
    /// Radar (spider) chart.
    Radar,
    /// Surface chart.
    Surface,
    /// Bubble chart.
    Bubble,
    /// Stock (OHLC) chart.
    Stock,
    /// Unrecognized chart type.
    Unknown,
}

impl ResolvedSheet {
    /// Create a new resolved sheet.
    pub fn new(name: String, worksheet: Worksheet, shared_strings: Vec<String>) -> Self {
        Self {
            name,
            worksheet,
            context: ResolveContext::new(shared_strings),
            comments: Vec::new(),
            charts: Vec::new(),
            #[cfg(feature = "sml-pivot")]
            pivot_tables: Vec::new(),
        }
    }

    /// Create a resolved sheet with comments, charts, and pivot tables.
    pub fn with_extras(
        name: String,
        worksheet: Worksheet,
        shared_strings: Vec<String>,
        comments: Vec<Comment>,
        charts: Vec<Chart>,
        #[cfg(feature = "sml-pivot")] pivot_tables: Vec<crate::types::CTPivotTableDefinition>,
    ) -> Self {
        Self {
            name,
            worksheet,
            context: ResolveContext::new(shared_strings),
            comments,
            charts,
            #[cfg(feature = "sml-pivot")]
            pivot_tables,
        }
    }

    /// Get the sheet name.
    pub fn name(&self) -> &str {
        &self.name
    }

    /// Get the underlying worksheet (generated type).
    pub fn worksheet(&self) -> &Worksheet {
        &self.worksheet
    }

    /// Get the resolution context.
    pub fn context(&self) -> &ResolveContext {
        &self.context
    }

    // -------------------------------------------------------------------------
    // Row/Cell Access (delegating to WorksheetExt)
    // -------------------------------------------------------------------------

    /// Get the number of rows.
    pub fn row_count(&self) -> usize {
        self.worksheet.row_count()
    }

    /// Check if the sheet is empty.
    pub fn is_empty(&self) -> bool {
        self.worksheet.is_empty()
    }

    /// Get a row by 1-based row number.
    pub fn row(&self, row_num: u32) -> Option<&Row> {
        self.worksheet.row(row_num)
    }

    /// Iterate over all rows.
    pub fn rows(&self) -> impl Iterator<Item = &Row> {
        self.worksheet.rows()
    }

    /// Get a cell by reference (e.g., "A1").
    pub fn cell(&self, reference: &str) -> Option<&Cell> {
        self.worksheet.cell(reference)
    }

    // -------------------------------------------------------------------------
    // Value Resolution (convenience methods)
    // -------------------------------------------------------------------------

    /// Get a cell's resolved value.
    pub fn cell_value(&self, cell: &Cell) -> CellValue {
        cell.resolved_value(&self.context)
    }

    /// Get a cell's value as a display string.
    pub fn cell_value_string(&self, cell: &Cell) -> String {
        cell.value_as_string(&self.context)
    }

    /// Get a cell's value as a number (if applicable).
    pub fn cell_value_number(&self, cell: &Cell) -> Option<f64> {
        cell.value_as_number(&self.context)
    }

    /// Get a cell's value as a boolean (if applicable).
    pub fn cell_value_bool(&self, cell: &Cell) -> Option<bool> {
        cell.value_as_bool(&self.context)
    }

    /// Get the value at a cell reference as a string.
    pub fn value_at(&self, reference: &str) -> Option<String> {
        self.cell(reference).map(|c| self.cell_value_string(c))
    }

    /// Get the value at a cell reference as a number.
    pub fn number_at(&self, reference: &str) -> Option<f64> {
        self.cell(reference).and_then(|c| self.cell_value_number(c))
    }

    // -------------------------------------------------------------------------
    // Sheet Features
    // -------------------------------------------------------------------------

    /// Check if the sheet has an auto-filter.
    #[cfg(feature = "sml-filtering")]
    pub fn has_auto_filter(&self) -> bool {
        self.worksheet.has_auto_filter()
    }

    /// Check if the sheet has merged cells.
    pub fn has_merged_cells(&self) -> bool {
        self.worksheet.has_merged_cells()
    }

    /// Check if the sheet has conditional formatting.
    #[cfg(feature = "sml-styling")]
    pub fn has_conditional_formatting(&self) -> bool {
        self.worksheet.has_conditional_formatting()
    }

    /// Check if the sheet has data validations.
    #[cfg(feature = "sml-validation")]
    pub fn has_data_validations(&self) -> bool {
        self.worksheet.has_data_validations()
    }

    /// Check if the sheet has freeze panes.
    #[cfg(feature = "sml-structure")]
    pub fn has_freeze_panes(&self) -> bool {
        self.worksheet.has_freeze_panes()
    }

    // -------------------------------------------------------------------------
    // Comments
    // -------------------------------------------------------------------------

    /// Get all comments.
    pub fn comments(&self) -> &[Comment] {
        &self.comments
    }

    /// Get the comment for a specific cell.
    pub fn comment(&self, reference: &str) -> Option<&Comment> {
        self.comments.iter().find(|c| c.reference == reference)
    }

    /// Check if a cell has a comment.
    pub fn has_comment(&self, reference: &str) -> bool {
        self.comment(reference).is_some()
    }

    // -------------------------------------------------------------------------
    // Charts
    // -------------------------------------------------------------------------

    /// Get all charts.
    pub fn charts(&self) -> &[Chart] {
        &self.charts
    }

    // -------------------------------------------------------------------------
    // Pivot Tables
    // -------------------------------------------------------------------------

    /// Get all pivot tables on this sheet.
    ///
    /// Pivot tables are loaded from the sheet's relationships (pivotTable parts).
    /// Requires the `sml-pivot` feature.
    ///
    /// ECMA-376 Part 1, Section 18.10 (PivotTable).
    #[cfg(feature = "sml-pivot")]
    pub fn pivot_tables(&self) -> &[crate::types::CTPivotTableDefinition] {
        &self.pivot_tables
    }

    // -------------------------------------------------------------------------
    // Dimensions & Structure
    // -------------------------------------------------------------------------

    /// Get the used range dimensions: (min_row, min_col, max_row, max_col).
    ///
    /// Returns None if the sheet is empty.
    pub fn dimensions(&self) -> Option<(u32, u32, u32, u32)> {
        if self.worksheet.sheet_data.row.is_empty() {
            return None;
        }

        let mut min_row = u32::MAX;
        let mut max_row = 0u32;
        let mut min_col = u32::MAX;
        let mut max_col = 0u32;

        for row in &self.worksheet.sheet_data.row {
            if let Some(row_num) = row.reference {
                min_row = min_row.min(row_num);
                max_row = max_row.max(row_num);
            }
            for cell in &row.cells {
                if let Some(col) = cell.column_number() {
                    min_col = min_col.min(col);
                    max_col = max_col.max(col);
                }
            }
        }

        if min_row == u32::MAX {
            None
        } else {
            Some((min_row, min_col, max_row, max_col))
        }
    }

    /// Get merged cell ranges (raw data).
    pub fn merged_cells(&self) -> Option<&crate::types::MergedCells> {
        self.worksheet.merged_cells.as_deref()
    }

    /// Get conditional formatting rules (raw data).
    #[cfg(feature = "sml-styling")]
    pub fn conditional_formatting(&self) -> &[crate::types::ConditionalFormatting] {
        &self.worksheet.conditional_formatting
    }

    /// Get data validations (raw data).
    #[cfg(feature = "sml-validation")]
    pub fn data_validations(&self) -> Option<&crate::types::DataValidations> {
        self.worksheet.data_validations.as_deref()
    }

    /// Get the auto-filter configuration (raw data).
    #[cfg(feature = "sml-filtering")]
    pub fn auto_filter(&self) -> Option<&crate::types::AutoFilter> {
        self.worksheet.auto_filter.as_deref()
    }

    /// Get the sheet views (contains freeze pane info).
    pub fn sheet_views(&self) -> Option<&crate::types::SheetViews> {
        self.worksheet.sheet_views.as_deref()
    }

    /// Get the freeze pane configuration (if any).
    #[cfg(feature = "sml-structure")]
    pub fn freeze_pane(&self) -> Option<&crate::types::Pane> {
        self.worksheet
            .sheet_views
            .as_ref()
            .and_then(|views| views.sheet_view.first())
            .and_then(|view| view.pane.as_deref())
    }

    /// Get column definitions.
    #[cfg(feature = "sml-styling")]
    pub fn columns(&self) -> &[crate::types::Columns] {
        &self.worksheet.cols
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

// =============================================================================
// Pivot Table Extension Traits
// =============================================================================

/// Extension methods for `types::CTPivotTableDefinition`.
///
/// Provides convenient access to pivot table name, location, and field indices.
/// Gated on the `sml-pivot` feature.
///
/// ECMA-376 Part 1, Section 18.10.1.73 (pivotTableDefinition).
#[cfg(feature = "sml-pivot")]
pub trait PivotTableExt {
    /// Get the pivot table name.
    fn name(&self) -> &str;

    /// Get the cell range for this pivot table (e.g., "A1:D10").
    ///
    /// This is the `ref` attribute of the `location` element.
    fn location_reference(&self) -> &str;

    /// Get the names of all data fields (value fields).
    ///
    /// Returns each data field's `name` attribute when present, or an empty
    /// string for unnamed data fields.
    fn data_field_names(&self) -> Vec<&str>;

    /// Get the field indices used as row fields.
    ///
    /// Each value is the `x` attribute of a `<field>` element in `<rowFields>`.
    /// Negative values (e.g., `-2`) represent the special "data" axis field.
    fn row_field_indices(&self) -> Vec<i32>;

    /// Get the field indices used as column fields.
    ///
    /// Each value is the `x` attribute of a `<field>` element in `<colFields>`.
    /// Negative values (e.g., `-2`) represent the special "data" axis field.
    fn col_field_indices(&self) -> Vec<i32>;
}

#[cfg(feature = "sml-pivot")]
impl PivotTableExt for crate::types::CTPivotTableDefinition {
    fn name(&self) -> &str {
        &self.name
    }

    fn location_reference(&self) -> &str {
        &self.location.reference
    }

    fn data_field_names(&self) -> Vec<&str> {
        self.data_fields
            .as_ref()
            .map(|df| {
                df.data_field
                    .iter()
                    .map(|f| f.name.as_deref().unwrap_or(""))
                    .collect()
            })
            .unwrap_or_default()
    }

    fn row_field_indices(&self) -> Vec<i32> {
        self.row_fields
            .as_ref()
            .map(|rf| rf.field.iter().map(|f| f.x).collect())
            .unwrap_or_default()
    }

    fn col_field_indices(&self) -> Vec<i32> {
        self.col_fields
            .as_ref()
            .map(|cf| cf.field.iter().map(|f| f.x).collect())
            .unwrap_or_default()
    }
}

// =============================================================================
// Conditional Formatting Extension Traits
// =============================================================================

/// Extension methods for `types::ConditionalFormatting`.
///
/// Provides convenient access to cell range and contained rules.
/// Gated on the `sml-styling` feature.
///
/// ECMA-376 Part 1, Section 18.3.1.18 (conditionalFormatting).
#[cfg(feature = "sml-styling")]
pub trait ConditionalFormattingExt {
    /// Get the cell range this formatting applies to (sqref attribute).
    ///
    /// For example, `"A1:B10"` or `"A1:A10 C1:C10"` (space-separated ranges).
    fn cell_range(&self) -> Option<&str>;

    /// Get the conditional formatting rules.
    fn rules(&self) -> &[crate::types::ConditionalRule];

    /// Get the number of rules.
    fn rule_count(&self) -> usize;
}

#[cfg(feature = "sml-styling")]
impl ConditionalFormattingExt for crate::types::ConditionalFormatting {
    fn cell_range(&self) -> Option<&str> {
        self.square_reference.as_deref()
    }

    fn rules(&self) -> &[crate::types::ConditionalRule] {
        &self.cf_rule
    }

    fn rule_count(&self) -> usize {
        self.cf_rule.len()
    }
}

/// Extension methods for `types::ConditionalRule`.
///
/// Provides convenient access to rule type and visualization sub-elements.
/// Gated on the `sml-styling` feature.
///
/// ECMA-376 Part 1, Section 18.3.1.10 (cfRule).
#[cfg(feature = "sml-styling")]
pub trait ConditionalRuleExt {
    /// Get the rule type (e.g., ColorScale, DataBar, CellIs, Expression).
    fn rule_type(&self) -> Option<&crate::types::ConditionalType>;

    /// Get the rule priority (lower number = higher priority).
    fn priority(&self) -> i32;

    /// Check if this rule uses a color scale visualization.
    fn has_color_scale(&self) -> bool;

    /// Check if this rule uses a data bar visualization.
    fn has_data_bar(&self) -> bool;

    /// Check if this rule uses an icon set visualization.
    fn has_icon_set(&self) -> bool;

    /// Get the formula expressions associated with this rule.
    ///
    /// For `cellIs` rules, contains 1-2 formulas (operands).
    /// For `expression` rules, contains 1 formula.
    /// For visualization rules (colorScale, dataBar, iconSet), empty.
    fn formulas(&self) -> &[crate::types::STFormula];
}

#[cfg(feature = "sml-styling")]
impl ConditionalRuleExt for crate::types::ConditionalRule {
    fn rule_type(&self) -> Option<&crate::types::ConditionalType> {
        self.r#type.as_ref()
    }

    fn priority(&self) -> i32 {
        self.priority
    }

    fn has_color_scale(&self) -> bool {
        self.color_scale.is_some()
    }

    fn has_data_bar(&self) -> bool {
        self.data_bar.is_some()
    }

    fn has_icon_set(&self) -> bool {
        self.icon_set.is_some()
    }

    fn formulas(&self) -> &[crate::types::STFormula] {
        &self.formula
    }
}

// =============================================================================
// Font / Fill / Border / Format Extension Traits
// =============================================================================

/// Extension methods for `types::Font` (ECMA-376 §18.8.22, CT_Font).
///
/// Provides convenient access to font properties such as bold, italic, name,
/// and size. All accessors are gated on the `sml-styling` feature because the
/// underlying fields are only present when that feature is enabled.
#[cfg(feature = "sml-styling")]
pub trait FontExt {
    /// Check if bold is set.
    fn is_bold(&self) -> bool;
    /// Check if italic is set.
    fn is_italic(&self) -> bool;
    /// Check if strikethrough is set.
    fn is_strikethrough(&self) -> bool;
    /// Check if outline is set.
    fn is_outline(&self) -> bool;
    /// Check if shadow is set.
    fn is_shadow(&self) -> bool;
    /// Check if condense is set.
    fn is_condense(&self) -> bool;
    /// Check if extend is set.
    fn is_extend(&self) -> bool;
    /// Get the typeface name (e.g. `"Calibri"`).
    fn font_name(&self) -> Option<&str>;
    /// Get the font size in points (e.g. `11.0`).
    fn font_size(&self) -> Option<f64>;
    /// Get the font colour.
    fn font_color(&self) -> Option<&crate::types::Color>;
    /// Get the vertical alignment (baseline / superscript / subscript).
    fn vertical_align(&self) -> Option<crate::types::VerticalAlignRun>;
    /// Get the font scheme (none / major / minor).
    fn font_scheme(&self) -> Option<crate::types::FontScheme>;
}

#[cfg(feature = "sml-styling")]
impl FontExt for crate::types::Font {
    fn is_bold(&self) -> bool {
        self.b.as_ref().is_some_and(|v| v.value.unwrap_or(false))
    }

    fn is_italic(&self) -> bool {
        self.i.as_ref().is_some_and(|v| v.value.unwrap_or(false))
    }

    fn is_strikethrough(&self) -> bool {
        self.strike
            .as_ref()
            .is_some_and(|v| v.value.unwrap_or(false))
    }

    fn is_outline(&self) -> bool {
        self.outline
            .as_ref()
            .is_some_and(|v| v.value.unwrap_or(false))
    }

    fn is_shadow(&self) -> bool {
        self.shadow
            .as_ref()
            .is_some_and(|v| v.value.unwrap_or(false))
    }

    fn is_condense(&self) -> bool {
        self.condense
            .as_ref()
            .is_some_and(|v| v.value.unwrap_or(false))
    }

    fn is_extend(&self) -> bool {
        self.extend
            .as_ref()
            .is_some_and(|v| v.value.unwrap_or(false))
    }

    fn font_name(&self) -> Option<&str> {
        self.name.as_deref().map(|n| n.value.as_str())
    }

    fn font_size(&self) -> Option<f64> {
        self.sz.as_deref().map(|s| s.value)
    }

    fn font_color(&self) -> Option<&crate::types::Color> {
        self.color.as_deref()
    }

    fn vertical_align(&self) -> Option<crate::types::VerticalAlignRun> {
        self.vert_align.as_deref().map(|v| v.value)
    }

    fn font_scheme(&self) -> Option<crate::types::FontScheme> {
        self.scheme.as_deref().map(|s| s.value)
    }
}

/// Extension methods for `types::Fill` (ECMA-376 §18.8.20, CT_Fill).
///
/// Provides access to the fill sub-type (pattern or gradient).
/// Gated on the `sml-styling` feature.
#[cfg(feature = "sml-styling")]
pub trait FillExt {
    /// Get the pattern fill definition, if this fill uses a pattern.
    fn pattern_fill(&self) -> Option<&crate::types::PatternFill>;
    /// Get the gradient fill definition, if this fill uses a gradient.
    fn gradient_fill(&self) -> Option<&crate::types::GradientFill>;
    /// Check if this fill has any fill type set (pattern or gradient).
    fn has_fill(&self) -> bool;
}

#[cfg(feature = "sml-styling")]
impl FillExt for crate::types::Fill {
    fn pattern_fill(&self) -> Option<&crate::types::PatternFill> {
        self.pattern_fill.as_deref()
    }

    fn gradient_fill(&self) -> Option<&crate::types::GradientFill> {
        self.gradient_fill.as_deref()
    }

    fn has_fill(&self) -> bool {
        self.pattern_fill.is_some() || self.gradient_fill.is_some()
    }
}

/// Extension methods for `types::PatternFill` (ECMA-376 §18.8.32, CT_PatternFill).
///
/// Provides access to the pattern type and foreground/background colours.
pub trait PatternFillExt {
    /// Get the pattern type (solid, dark-grid, etc.).
    fn pattern_type(&self) -> Option<crate::types::PatternType>;
    /// Get the foreground colour of the pattern.
    fn foreground_color(&self) -> Option<&crate::types::Color>;
    /// Get the background colour of the pattern.
    fn background_color(&self) -> Option<&crate::types::Color>;
}

impl PatternFillExt for crate::types::PatternFill {
    fn pattern_type(&self) -> Option<crate::types::PatternType> {
        self.pattern_type
    }

    fn foreground_color(&self) -> Option<&crate::types::Color> {
        self.fg_color.as_deref()
    }

    fn background_color(&self) -> Option<&crate::types::Color> {
        self.bg_color.as_deref()
    }
}

/// Extension methods for `types::Border` (ECMA-376 §18.8.4, CT_Border).
///
/// Provides access to individual border sides and diagonal flags.
/// Gated on the `sml-styling` feature.
#[cfg(feature = "sml-styling")]
pub trait BorderExt {
    /// Get the left border properties.
    fn left_border(&self) -> Option<&crate::types::BorderProperties>;
    /// Get the right border properties.
    fn right_border(&self) -> Option<&crate::types::BorderProperties>;
    /// Get the top border properties.
    fn top_border(&self) -> Option<&crate::types::BorderProperties>;
    /// Get the bottom border properties.
    fn bottom_border(&self) -> Option<&crate::types::BorderProperties>;
    /// Get the diagonal border properties.
    fn diagonal_border(&self) -> Option<&crate::types::BorderProperties>;
    /// Check if the diagonal-up line is shown.
    fn is_diagonal_up(&self) -> bool;
    /// Check if the diagonal-down line is shown.
    fn is_diagonal_down(&self) -> bool;
    /// Check if the outline border is applied.
    fn is_outline_applied(&self) -> bool;
}

#[cfg(feature = "sml-styling")]
impl BorderExt for crate::types::Border {
    fn left_border(&self) -> Option<&crate::types::BorderProperties> {
        self.left.as_deref()
    }

    fn right_border(&self) -> Option<&crate::types::BorderProperties> {
        self.right.as_deref()
    }

    fn top_border(&self) -> Option<&crate::types::BorderProperties> {
        self.top.as_deref()
    }

    fn bottom_border(&self) -> Option<&crate::types::BorderProperties> {
        self.bottom.as_deref()
    }

    fn diagonal_border(&self) -> Option<&crate::types::BorderProperties> {
        self.diagonal.as_deref()
    }

    fn is_diagonal_up(&self) -> bool {
        self.diagonal_up.unwrap_or(false)
    }

    fn is_diagonal_down(&self) -> bool {
        self.diagonal_down.unwrap_or(false)
    }

    fn is_outline_applied(&self) -> bool {
        self.outline.unwrap_or(false)
    }
}

/// Extension methods for `types::BorderProperties` (ECMA-376 §18.8.3, CT_BorderPr).
///
/// Provides access to border line style and colour.
pub trait BorderPropertiesExt {
    /// Get the border line style.
    fn border_style(&self) -> Option<crate::types::BorderStyle>;
    /// Get the border colour.
    fn border_color(&self) -> Option<&crate::types::Color>;
    /// Check if a border style is set.
    fn has_style(&self) -> bool;
}

impl BorderPropertiesExt for crate::types::BorderProperties {
    fn border_style(&self) -> Option<crate::types::BorderStyle> {
        self.style
    }

    fn border_color(&self) -> Option<&crate::types::Color> {
        self.color.as_deref()
    }

    fn has_style(&self) -> bool {
        self.style.is_some()
    }
}

/// Extension methods for `types::CellAlignment` (ECMA-376 §18.8.1, CT_CellAlignment).
///
/// Provides access to horizontal/vertical alignment, text rotation, and wrapping.
/// Gated on the `sml-styling` feature.
#[cfg(feature = "sml-styling")]
pub trait CellAlignmentExt {
    /// Get the horizontal alignment.
    fn horizontal_alignment(&self) -> Option<crate::types::HorizontalAlignment>;
    /// Get the vertical alignment.
    fn vertical_alignment(&self) -> Option<crate::types::VerticalAlignment>;
    /// Get the text rotation in degrees (0–180 or 255 for vertical text).
    fn text_rotation(&self) -> Option<u32>;
    /// Check if text wrapping is enabled.
    fn is_wrap_text(&self) -> bool;
    /// Check if shrink-to-fit is enabled.
    fn is_shrink_to_fit(&self) -> bool;
    /// Get the indent level.
    fn indent(&self) -> Option<u32>;
}

#[cfg(feature = "sml-styling")]
impl CellAlignmentExt for crate::types::CellAlignment {
    fn horizontal_alignment(&self) -> Option<crate::types::HorizontalAlignment> {
        self.horizontal
    }

    fn vertical_alignment(&self) -> Option<crate::types::VerticalAlignment> {
        self.vertical
    }

    fn text_rotation(&self) -> Option<u32> {
        self.text_rotation.as_deref().and_then(|r| r.parse().ok())
    }

    fn is_wrap_text(&self) -> bool {
        self.wrap_text.unwrap_or(false)
    }

    fn is_shrink_to_fit(&self) -> bool {
        self.shrink_to_fit.unwrap_or(false)
    }

    fn indent(&self) -> Option<u32> {
        self.indent
    }
}

/// Extension methods for `types::CellProtection` (ECMA-376 §18.8.13, CT_CellProtection).
///
/// Provides access to cell locking and formula-hiding flags.
/// Gated on the `sml-protection` feature.
#[cfg(feature = "sml-protection")]
pub trait CellProtectionExt {
    /// Check if the cell is locked (default: `true` per OOXML spec).
    fn is_locked(&self) -> bool;
    /// Check if the formula is hidden from the formula bar.
    fn is_formula_hidden(&self) -> bool;
}

#[cfg(feature = "sml-protection")]
impl CellProtectionExt for crate::types::CellProtection {
    fn is_locked(&self) -> bool {
        // Per ECMA-376, the default value is `true` — cells are locked unless
        // explicitly set to `false`.
        self.locked.unwrap_or(true)
    }

    fn is_formula_hidden(&self) -> bool {
        self.hidden.unwrap_or(false)
    }
}

/// Extension methods for `types::Format` (ECMA-376 §18.8.45, CT_Xf).
///
/// The `xf` element is used in both `cellXfs` (cell formats) and
/// `cellStyleXfs` (named style formats). It carries index references into
/// the font, fill, border, and numFmt tables, plus optional alignment and
/// protection overrides.
pub trait FormatExt {
    /// Get the number format ID (index into `numFmts`).
    #[cfg(feature = "sml-styling")]
    fn number_format_id(&self) -> u32;
    /// Get the font ID (index into `fonts`).
    #[cfg(feature = "sml-styling")]
    fn font_id(&self) -> u32;
    /// Get the fill ID (index into `fills`).
    #[cfg(feature = "sml-styling")]
    fn fill_id(&self) -> u32;
    /// Get the border ID (index into `borders`).
    #[cfg(feature = "sml-styling")]
    fn border_id(&self) -> u32;
    /// Check if alignment properties are applied from this format.
    #[cfg(feature = "sml-styling")]
    fn apply_alignment(&self) -> bool;
    /// Check if font properties are applied from this format.
    #[cfg(feature = "sml-styling")]
    fn apply_font(&self) -> bool;
    /// Check if fill properties are applied from this format.
    #[cfg(feature = "sml-styling")]
    fn apply_fill(&self) -> bool;
    /// Check if border properties are applied from this format.
    #[cfg(feature = "sml-styling")]
    fn apply_border(&self) -> bool;
    /// Get the cell alignment override (if any).
    #[cfg(feature = "sml-styling")]
    fn alignment(&self) -> Option<&crate::types::CellAlignment>;
    /// Get the cell protection override (if any).
    #[cfg(feature = "sml-protection")]
    fn protection(&self) -> Option<&crate::types::CellProtection>;
}

impl FormatExt for crate::types::Format {
    #[cfg(feature = "sml-styling")]
    fn number_format_id(&self) -> u32 {
        self.number_format_id.unwrap_or(0)
    }

    #[cfg(feature = "sml-styling")]
    fn font_id(&self) -> u32 {
        self.font_id.unwrap_or(0)
    }

    #[cfg(feature = "sml-styling")]
    fn fill_id(&self) -> u32 {
        self.fill_id.unwrap_or(0)
    }

    #[cfg(feature = "sml-styling")]
    fn border_id(&self) -> u32 {
        self.border_id.unwrap_or(0)
    }

    #[cfg(feature = "sml-styling")]
    fn apply_alignment(&self) -> bool {
        self.apply_alignment.unwrap_or(false)
    }

    #[cfg(feature = "sml-styling")]
    fn apply_font(&self) -> bool {
        self.apply_font.unwrap_or(false)
    }

    #[cfg(feature = "sml-styling")]
    fn apply_fill(&self) -> bool {
        self.apply_fill.unwrap_or(false)
    }

    #[cfg(feature = "sml-styling")]
    fn apply_border(&self) -> bool {
        self.apply_border.unwrap_or(false)
    }

    #[cfg(feature = "sml-styling")]
    fn alignment(&self) -> Option<&crate::types::CellAlignment> {
        self.alignment.as_deref()
    }

    #[cfg(feature = "sml-protection")]
    fn protection(&self) -> Option<&crate::types::CellProtection> {
        self.protection.as_deref()
    }
}

/// Extension for `types::Worksheet` to access conditional formatting.
///
/// Gated on the `sml-styling` feature.
#[cfg(feature = "sml-styling")]
pub trait WorksheetConditionalFormattingExt {
    /// Get all conditional formatting rules on this worksheet.
    fn conditional_formattings(&self) -> &[crate::types::ConditionalFormatting];
}

#[cfg(feature = "sml-styling")]
impl WorksheetConditionalFormattingExt for crate::types::Worksheet {
    fn conditional_formattings(&self) -> &[crate::types::ConditionalFormatting] {
        &self.conditional_formatting
    }
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
    #[cfg(feature = "full")]
    fn test_cell_ext() {
        let cell = Cell {
            reference: Some("B5".to_string()),
            cell_type: Some(CellType::Number),
            value: Some("42.5".to_string()),
            formula: None,
            style_index: None,
            cm: None,
            vm: None,
            placeholder: None,
            is: None,
            extension_list: None,
            #[cfg(feature = "extra-attrs")]
            extra_attrs: Default::default(),
            #[cfg(feature = "extra-children")]
            extra_children: Default::default(),
        };

        assert_eq!(cell.column_number(), Some(2));
        assert_eq!(cell.row_number(), Some(5));
        assert!(!cell.has_formula());
        assert!(cell.is_number());
        assert!(!cell.is_shared_string());
    }

    #[test]
    #[cfg(feature = "full")]
    fn test_cell_resolve_number() {
        let cell = Cell {
            reference: Some("A1".to_string()),
            cell_type: Some(CellType::Number),
            value: Some("123.45".to_string()),
            formula: None,
            style_index: None,
            cm: None,
            vm: None,
            placeholder: None,
            is: None,
            extension_list: None,
            #[cfg(feature = "extra-attrs")]
            extra_attrs: Default::default(),
            #[cfg(feature = "extra-children")]
            extra_children: Default::default(),
        };

        let ctx = ResolveContext::default();
        assert_eq!(cell.resolved_value(&ctx), CellValue::Number(123.45));
        assert_eq!(cell.value_as_string(&ctx), "123.45");
        assert_eq!(cell.value_as_number(&ctx), Some(123.45));
    }

    #[test]
    #[cfg(feature = "full")]
    fn test_cell_resolve_shared_string() {
        let cell = Cell {
            reference: Some("A1".to_string()),
            cell_type: Some(CellType::SharedString),
            value: Some("0".to_string()), // Index into shared strings
            formula: None,
            style_index: None,
            cm: None,
            vm: None,
            placeholder: None,
            is: None,
            extension_list: None,
            #[cfg(feature = "extra-attrs")]
            extra_attrs: Default::default(),
            #[cfg(feature = "extra-children")]
            extra_children: Default::default(),
        };

        let ctx = ResolveContext::new(vec!["Hello".to_string(), "World".to_string()]);
        assert_eq!(
            cell.resolved_value(&ctx),
            CellValue::String("Hello".to_string())
        );
        assert_eq!(cell.value_as_string(&ctx), "Hello");
    }

    #[test]
    #[cfg(feature = "full")]
    fn test_cell_resolve_boolean() {
        let cell = Cell {
            reference: Some("A1".to_string()),
            cell_type: Some(CellType::Boolean),
            value: Some("1".to_string()),
            formula: None,
            style_index: None,
            cm: None,
            vm: None,
            placeholder: None,
            is: None,
            extension_list: None,
            #[cfg(feature = "extra-attrs")]
            extra_attrs: Default::default(),
            #[cfg(feature = "extra-children")]
            extra_children: Default::default(),
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

    #[test]
    fn test_resolved_sheet() {
        let xml = br#"<?xml version="1.0" encoding="UTF-8"?>
        <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
            <sheetData>
                <row r="1">
                    <c r="A1" t="s"><v>0</v></c>
                    <c r="B1"><v>42.5</v></c>
                    <c r="C1" t="b"><v>1</v></c>
                </row>
                <row r="2">
                    <c r="A2" t="s"><v>1</v></c>
                    <c r="B2"><v>100</v></c>
                </row>
            </sheetData>
        </worksheet>"#;

        let worksheet = parse_worksheet(xml).expect("parse failed");
        let shared_strings = vec!["Hello".to_string(), "World".to_string()];
        let sheet = ResolvedSheet::new("Sheet1".to_string(), worksheet, shared_strings);

        // Basic info
        assert_eq!(sheet.name(), "Sheet1");
        assert_eq!(sheet.row_count(), 2);
        assert!(!sheet.is_empty());

        // Cell access with auto-resolution
        let cell_a1 = sheet.cell("A1").expect("A1");
        assert_eq!(sheet.cell_value_string(cell_a1), "Hello");

        let cell_b1 = sheet.cell("B1").expect("B1");
        assert_eq!(sheet.cell_value_number(cell_b1), Some(42.5));

        let cell_c1 = sheet.cell("C1").expect("C1");
        assert_eq!(sheet.cell_value_bool(cell_c1), Some(true));

        // Convenience methods
        assert_eq!(sheet.value_at("A1"), Some("Hello".to_string()));
        assert_eq!(sheet.value_at("A2"), Some("World".to_string()));
        assert_eq!(sheet.number_at("B1"), Some(42.5));
        assert_eq!(sheet.number_at("B2"), Some(100.0));

        // Non-existent cell
        assert!(sheet.value_at("Z99").is_none());
    }

    #[test]
    #[cfg(feature = "sml-styling")]
    fn test_conditional_formatting_cell_range() {
        use crate::workbook::bootstrap;
        let xml = br#"<?xml version="1.0" encoding="UTF-8"?>
        <conditionalFormatting sqref="A1:B5" xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
        </conditionalFormatting>"#;
        let cf: crate::types::ConditionalFormatting = bootstrap(xml).expect("parse failed");
        assert_eq!(cf.cell_range(), Some("A1:B5"));
        assert_eq!(cf.rule_count(), 0);
    }

    #[test]
    #[cfg(feature = "sml-styling")]
    fn test_conditional_rule_type() {
        use crate::workbook::bootstrap;
        let xml = br#"<?xml version="1.0" encoding="UTF-8"?>
        <cfRule type="colorScale" priority="1" xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
            <colorScale>
                <cfvo type="min"/>
                <cfvo type="max"/>
                <color rgb="FF0000"/>
                <color rgb="00FF00"/>
            </colorScale>
        </cfRule>"#;
        let rule: crate::types::ConditionalRule = bootstrap(xml).expect("parse failed");
        assert_eq!(
            rule.rule_type(),
            Some(&crate::types::ConditionalType::ColorScale)
        );
        assert_eq!(rule.priority(), 1);
    }

    #[test]
    #[cfg(feature = "sml-styling")]
    fn test_conditional_rule_has_color_scale() {
        use crate::workbook::bootstrap;
        let xml = br#"<?xml version="1.0" encoding="UTF-8"?>
        <cfRule type="colorScale" priority="1" xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
            <colorScale>
                <cfvo type="min"/>
                <cfvo type="max"/>
                <color rgb="FF0000"/>
                <color rgb="00FF00"/>
            </colorScale>
        </cfRule>"#;
        let rule: crate::types::ConditionalRule = bootstrap(xml).expect("parse failed");
        assert!(rule.has_color_scale());
        assert!(!rule.has_data_bar());
        assert!(!rule.has_icon_set());
    }

    #[test]
    #[cfg(feature = "sml-pivot")]
    fn test_pivot_table_name() {
        use crate::workbook::bootstrap;
        let xml = br#"<?xml version="1.0" encoding="UTF-8"?>
        <pivotTableDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
            name="PivotTable1" cacheId="1" dataCaption="Values">
            <location ref="A1:D10" firstHeaderRow="1" firstDataRow="2" firstDataCol="1"/>
        </pivotTableDefinition>"#;
        let pt: crate::types::CTPivotTableDefinition = bootstrap(xml).expect("parse failed");
        assert_eq!(pt.name(), "PivotTable1");
    }

    #[test]
    #[cfg(feature = "sml-pivot")]
    fn test_pivot_table_location() {
        use crate::workbook::bootstrap;
        let xml = br#"<?xml version="1.0" encoding="UTF-8"?>
        <pivotTableDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
            name="PivotTable1" cacheId="1" dataCaption="Values">
            <location ref="A1:D10" firstHeaderRow="1" firstDataRow="2" firstDataCol="1"/>
        </pivotTableDefinition>"#;
        let pt: crate::types::CTPivotTableDefinition = bootstrap(xml).expect("parse failed");
        assert_eq!(pt.location_reference(), "A1:D10");
    }

    // -------------------------------------------------------------------------
    // FontExt tests
    // -------------------------------------------------------------------------

    #[cfg(feature = "sml-styling")]
    #[test]
    fn test_font_ext_bold_italic() {
        use crate::workbook::bootstrap;
        let xml = br#"<?xml version="1.0" encoding="UTF-8"?>
        <font xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
            <b val="1"/>
            <i val="0"/>
            <name val="Calibri"/>
            <sz val="11"/>
        </font>"#;
        let font: crate::types::Font = bootstrap(xml).expect("parse failed");
        assert!(font.is_bold());
        assert!(!font.is_italic());
        assert_eq!(font.font_name(), Some("Calibri"));
        assert_eq!(font.font_size(), Some(11.0));
    }

    #[cfg(feature = "sml-styling")]
    #[test]
    fn test_font_ext_defaults() {
        use crate::workbook::bootstrap;
        let xml = br#"<?xml version="1.0" encoding="UTF-8"?>
        <font xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"/>
        "#;
        let font: crate::types::Font = bootstrap(xml).expect("parse failed");
        assert!(!font.is_bold());
        assert!(!font.is_italic());
        assert!(!font.is_strikethrough());
        assert!(font.font_name().is_none());
        assert!(font.font_size().is_none());
        assert!(font.font_color().is_none());
        assert!(font.vertical_align().is_none());
        assert!(font.font_scheme().is_none());
    }

    // -------------------------------------------------------------------------
    // FillExt / PatternFillExt tests
    // -------------------------------------------------------------------------

    #[cfg(feature = "sml-styling")]
    #[test]
    fn test_fill_ext_pattern() {
        use crate::workbook::bootstrap;
        let xml = br#"<?xml version="1.0" encoding="UTF-8"?>
        <fill xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
            <patternFill patternType="solid"/>
        </fill>"#;
        let fill: crate::types::Fill = bootstrap(xml).expect("parse failed");
        assert!(fill.has_fill());
        assert!(fill.pattern_fill().is_some());
        assert!(fill.gradient_fill().is_none());
    }

    #[cfg(feature = "sml-styling")]
    #[test]
    fn test_fill_ext_no_fill() {
        use crate::workbook::bootstrap;
        let xml = br#"<?xml version="1.0" encoding="UTF-8"?>
        <fill xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
            <patternFill patternType="none"/>
        </fill>"#;
        let fill: crate::types::Fill = bootstrap(xml).expect("parse failed");
        assert!(fill.has_fill()); // has a pattern fill element, even if "none"
        let pf = fill.pattern_fill().unwrap();
        use crate::ext::PatternFillExt;
        use crate::types::PatternType;
        assert_eq!(pf.pattern_type(), Some(PatternType::None));
    }

    // -------------------------------------------------------------------------
    // BorderExt tests
    // -------------------------------------------------------------------------

    #[cfg(feature = "sml-styling")]
    #[test]
    fn test_border_ext_diagonal_flags() {
        use crate::workbook::bootstrap;
        let xml = br#"<?xml version="1.0" encoding="UTF-8"?>
        <border diagonalUp="1" diagonalDown="0"
                xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
        </border>"#;
        let border: crate::types::Border = bootstrap(xml).expect("parse failed");
        assert!(border.is_diagonal_up());
        assert!(!border.is_diagonal_down());
        assert!(border.left_border().is_none());
        assert!(border.diagonal_border().is_none());
    }

    // -------------------------------------------------------------------------
    // CellAlignmentExt tests
    // -------------------------------------------------------------------------

    #[cfg(feature = "sml-styling")]
    #[test]
    fn test_cell_alignment_ext() {
        use crate::workbook::bootstrap;
        let xml = br#"<?xml version="1.0" encoding="UTF-8"?>
        <alignment horizontal="center" vertical="bottom" wrapText="1" indent="2"
                   xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"/>
        "#;
        let align: crate::types::CellAlignment = bootstrap(xml).expect("parse failed");
        use crate::types::{HorizontalAlignment, VerticalAlignment};
        assert_eq!(
            align.horizontal_alignment(),
            Some(HorizontalAlignment::Center)
        );
        assert_eq!(align.vertical_alignment(), Some(VerticalAlignment::Bottom));
        assert!(align.is_wrap_text());
        assert!(!align.is_shrink_to_fit());
        assert_eq!(align.indent(), Some(2));
    }

    // -------------------------------------------------------------------------
    // FormatExt tests
    // -------------------------------------------------------------------------

    #[cfg(feature = "sml-styling")]
    #[test]
    fn test_format_ext_ids() {
        use crate::workbook::bootstrap;
        let xml = br#"<?xml version="1.0" encoding="UTF-8"?>
        <xf numFmtId="4" fontId="1" fillId="2" borderId="3" applyFont="1" applyFill="0"
            xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"/>
        "#;
        let fmt: crate::types::Format = bootstrap(xml).expect("parse failed");
        use crate::ext::FormatExt;
        assert_eq!(fmt.number_format_id(), 4);
        assert_eq!(fmt.font_id(), 1);
        assert_eq!(fmt.fill_id(), 2);
        assert_eq!(fmt.border_id(), 3);
        assert!(fmt.apply_font());
        assert!(!fmt.apply_fill());
        assert!(fmt.alignment().is_none());
    }

    // -------------------------------------------------------------------------
    // WorksheetExt row/column dimension tests
    // -------------------------------------------------------------------------

    #[cfg(feature = "sml-styling")]
    #[test]
    fn test_worksheet_get_row_height() {
        let xml = br#"<?xml version="1.0" encoding="UTF-8"?>
        <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
            <sheetData>
                <row r="1" ht="20" customHeight="1">
                    <c r="A1"><v>1</v></c>
                </row>
                <row r="2">
                    <c r="A2"><v>2</v></c>
                </row>
            </sheetData>
        </worksheet>"#;
        let ws = parse_worksheet(xml).expect("parse failed");
        assert_eq!(ws.get_row_height(1), Some(20.0));
        assert_eq!(ws.get_row_height(2), None); // no explicit height
        assert_eq!(ws.get_row_height(99), None); // row doesn't exist
    }
}
