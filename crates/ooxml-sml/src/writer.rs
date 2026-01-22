//! Excel workbook writing support.
//!
//! This module provides `WorkbookBuilder` for creating new Excel files.
//!
//! # Example
//!
//! ```no_run
//! use ooxml_sml::WorkbookBuilder;
//!
//! let mut wb = WorkbookBuilder::new();
//! let sheet = wb.add_sheet("Sheet1");
//! sheet.set_cell("A1", "Hello");
//! sheet.set_cell("B1", 42.0);
//! sheet.set_cell("A2", true);
//! wb.save("output.xlsx")?;
//! # Ok::<(), ooxml_sml::Error>(())
//! ```

use crate::error::Result;
use ooxml_opc::PackageWriter;
use std::collections::HashMap;
use std::fs::File;
use std::io::{BufWriter, Seek, Write};
use std::path::Path;

// Content types
const CT_WORKBOOK: &str =
    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml";
const CT_WORKSHEET: &str =
    "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml";
const CT_SHARED_STRINGS: &str =
    "application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml";
const CT_STYLES: &str = "application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml";
const CT_RELATIONSHIPS: &str = "application/vnd.openxmlformats-package.relationships+xml";
const CT_XML: &str = "application/xml";

// Relationship types
const REL_OFFICE_DOCUMENT: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument";
const REL_WORKSHEET: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet";
const REL_SHARED_STRINGS: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings";
const REL_STYLES: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles";

// Namespaces
const NS_SPREADSHEET: &str = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const NS_RELATIONSHIPS: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

/// A value that can be written to a cell.
#[derive(Debug, Clone)]
pub enum WriteCellValue {
    /// String value.
    String(String),
    /// Numeric value.
    Number(f64),
    /// Boolean value.
    Boolean(bool),
    /// Formula (the formula text, not the result).
    Formula(String),
    /// Empty cell.
    Empty,
}

impl From<&str> for WriteCellValue {
    fn from(s: &str) -> Self {
        WriteCellValue::String(s.to_string())
    }
}

impl From<String> for WriteCellValue {
    fn from(s: String) -> Self {
        WriteCellValue::String(s)
    }
}

impl From<f64> for WriteCellValue {
    fn from(n: f64) -> Self {
        WriteCellValue::Number(n)
    }
}

impl From<i32> for WriteCellValue {
    fn from(n: i32) -> Self {
        WriteCellValue::Number(n as f64)
    }
}

impl From<i64> for WriteCellValue {
    fn from(n: i64) -> Self {
        WriteCellValue::Number(n as f64)
    }
}

impl From<bool> for WriteCellValue {
    fn from(b: bool) -> Self {
        WriteCellValue::Boolean(b)
    }
}

/// A cell style for formatting.
///
/// Use `CellStyleBuilder` to create styles, then apply them with `set_cell_style`.
#[derive(Debug, Clone, Default, PartialEq)]
pub struct CellStyle {
    /// Font formatting.
    pub font: Option<FontStyle>,
    /// Fill (background) formatting.
    pub fill: Option<FillStyle>,
    /// Border formatting.
    pub border: Option<BorderStyle>,
    /// Number format code (e.g., "0.00", "#,##0", "yyyy-mm-dd").
    pub number_format: Option<String>,
    /// Horizontal alignment.
    pub horizontal_alignment: Option<HorizontalAlignment>,
    /// Vertical alignment.
    pub vertical_alignment: Option<VerticalAlignment>,
    /// Text wrap.
    pub wrap_text: bool,
}

impl CellStyle {
    /// Create a new empty cell style.
    pub fn new() -> Self {
        Self::default()
    }

    /// Set the font style.
    pub fn with_font(mut self, font: FontStyle) -> Self {
        self.font = Some(font);
        self
    }

    /// Set the fill style.
    pub fn with_fill(mut self, fill: FillStyle) -> Self {
        self.fill = Some(fill);
        self
    }

    /// Set the border style.
    pub fn with_border(mut self, border: BorderStyle) -> Self {
        self.border = Some(border);
        self
    }

    /// Set the number format code.
    pub fn with_number_format(mut self, format: impl Into<String>) -> Self {
        self.number_format = Some(format.into());
        self
    }

    /// Set horizontal alignment.
    pub fn with_horizontal_alignment(mut self, align: HorizontalAlignment) -> Self {
        self.horizontal_alignment = Some(align);
        self
    }

    /// Set vertical alignment.
    pub fn with_vertical_alignment(mut self, align: VerticalAlignment) -> Self {
        self.vertical_alignment = Some(align);
        self
    }

    /// Enable text wrapping.
    pub fn with_wrap_text(mut self, wrap: bool) -> Self {
        self.wrap_text = wrap;
        self
    }
}

/// Font style for cell formatting.
#[derive(Debug, Clone, Default, PartialEq)]
pub struct FontStyle {
    /// Font name (e.g., "Arial", "Calibri").
    pub name: Option<String>,
    /// Font size in points.
    pub size: Option<f64>,
    /// Bold text.
    pub bold: bool,
    /// Italic text.
    pub italic: bool,
    /// Underline style.
    pub underline: Option<UnderlineStyle>,
    /// Strikethrough.
    pub strikethrough: bool,
    /// Font color as RGB hex (e.g., "FF0000" for red).
    pub color: Option<String>,
}

impl FontStyle {
    /// Create a new empty font style.
    pub fn new() -> Self {
        Self::default()
    }

    /// Set the font name.
    pub fn with_name(mut self, name: impl Into<String>) -> Self {
        self.name = Some(name.into());
        self
    }

    /// Set the font size.
    pub fn with_size(mut self, size: f64) -> Self {
        self.size = Some(size);
        self
    }

    /// Set bold.
    pub fn bold(mut self) -> Self {
        self.bold = true;
        self
    }

    /// Set italic.
    pub fn italic(mut self) -> Self {
        self.italic = true;
        self
    }

    /// Set underline.
    pub fn underline(mut self, style: UnderlineStyle) -> Self {
        self.underline = Some(style);
        self
    }

    /// Set strikethrough.
    pub fn strikethrough(mut self) -> Self {
        self.strikethrough = true;
        self
    }

    /// Set the font color (RGB hex, e.g., "FF0000" for red).
    pub fn with_color(mut self, color: impl Into<String>) -> Self {
        self.color = Some(color.into());
        self
    }
}

/// Underline style for fonts.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub enum UnderlineStyle {
    #[default]
    Single,
    Double,
    SingleAccounting,
    DoubleAccounting,
}

impl UnderlineStyle {
    fn to_xml_value(self) -> &'static str {
        match self {
            UnderlineStyle::Single => "single",
            UnderlineStyle::Double => "double",
            UnderlineStyle::SingleAccounting => "singleAccounting",
            UnderlineStyle::DoubleAccounting => "doubleAccounting",
        }
    }
}

/// Fill style for cell background.
#[derive(Debug, Clone, Default, PartialEq)]
pub struct FillStyle {
    /// Fill pattern type.
    pub pattern: FillPattern,
    /// Foreground color (pattern color) as RGB hex.
    pub fg_color: Option<String>,
    /// Background color as RGB hex.
    pub bg_color: Option<String>,
}

impl FillStyle {
    /// Create a new empty fill style.
    pub fn new() -> Self {
        Self::default()
    }

    /// Create a solid fill with the given color.
    pub fn solid(color: impl Into<String>) -> Self {
        Self {
            pattern: FillPattern::Solid,
            fg_color: Some(color.into()),
            bg_color: None,
        }
    }

    /// Set the pattern type.
    pub fn with_pattern(mut self, pattern: FillPattern) -> Self {
        self.pattern = pattern;
        self
    }

    /// Set the foreground color.
    pub fn with_fg_color(mut self, color: impl Into<String>) -> Self {
        self.fg_color = Some(color.into());
        self
    }

    /// Set the background color.
    pub fn with_bg_color(mut self, color: impl Into<String>) -> Self {
        self.bg_color = Some(color.into());
        self
    }
}

/// Fill pattern types.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub enum FillPattern {
    #[default]
    None,
    Solid,
    MediumGray,
    DarkGray,
    LightGray,
    DarkHorizontal,
    DarkVertical,
    DarkDown,
    DarkUp,
    DarkGrid,
    DarkTrellis,
    LightHorizontal,
    LightVertical,
    LightDown,
    LightUp,
    LightGrid,
    LightTrellis,
    Gray125,
    Gray0625,
}

impl FillPattern {
    fn to_xml_value(self) -> &'static str {
        match self {
            FillPattern::None => "none",
            FillPattern::Solid => "solid",
            FillPattern::MediumGray => "mediumGray",
            FillPattern::DarkGray => "darkGray",
            FillPattern::LightGray => "lightGray",
            FillPattern::DarkHorizontal => "darkHorizontal",
            FillPattern::DarkVertical => "darkVertical",
            FillPattern::DarkDown => "darkDown",
            FillPattern::DarkUp => "darkUp",
            FillPattern::DarkGrid => "darkGrid",
            FillPattern::DarkTrellis => "darkTrellis",
            FillPattern::LightHorizontal => "lightHorizontal",
            FillPattern::LightVertical => "lightVertical",
            FillPattern::LightDown => "lightDown",
            FillPattern::LightUp => "lightUp",
            FillPattern::LightGrid => "lightGrid",
            FillPattern::LightTrellis => "lightTrellis",
            FillPattern::Gray125 => "gray125",
            FillPattern::Gray0625 => "gray0625",
        }
    }
}

/// Border style for cells.
#[derive(Debug, Clone, Default, PartialEq)]
pub struct BorderStyle {
    /// Left border.
    pub left: Option<BorderSideStyle>,
    /// Right border.
    pub right: Option<BorderSideStyle>,
    /// Top border.
    pub top: Option<BorderSideStyle>,
    /// Bottom border.
    pub bottom: Option<BorderSideStyle>,
    /// Diagonal border.
    pub diagonal: Option<BorderSideStyle>,
    /// Diagonal up.
    pub diagonal_up: bool,
    /// Diagonal down.
    pub diagonal_down: bool,
}

impl BorderStyle {
    /// Create a new empty border style.
    pub fn new() -> Self {
        Self::default()
    }

    /// Create a border with all sides using the same style.
    pub fn all(style: BorderLineStyle, color: Option<String>) -> Self {
        let side = BorderSideStyle { style, color };
        Self {
            left: Some(side.clone()),
            right: Some(side.clone()),
            top: Some(side.clone()),
            bottom: Some(side),
            diagonal: None,
            diagonal_up: false,
            diagonal_down: false,
        }
    }

    /// Set the left border.
    pub fn with_left(mut self, style: BorderLineStyle, color: Option<String>) -> Self {
        self.left = Some(BorderSideStyle { style, color });
        self
    }

    /// Set the right border.
    pub fn with_right(mut self, style: BorderLineStyle, color: Option<String>) -> Self {
        self.right = Some(BorderSideStyle { style, color });
        self
    }

    /// Set the top border.
    pub fn with_top(mut self, style: BorderLineStyle, color: Option<String>) -> Self {
        self.top = Some(BorderSideStyle { style, color });
        self
    }

    /// Set the bottom border.
    pub fn with_bottom(mut self, style: BorderLineStyle, color: Option<String>) -> Self {
        self.bottom = Some(BorderSideStyle { style, color });
        self
    }
}

/// Style for a single border side.
#[derive(Debug, Clone, Default, PartialEq)]
pub struct BorderSideStyle {
    /// Line style.
    pub style: BorderLineStyle,
    /// Color as RGB hex.
    pub color: Option<String>,
}

/// Border line styles.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub enum BorderLineStyle {
    #[default]
    None,
    Thin,
    Medium,
    Dashed,
    Dotted,
    Thick,
    Double,
    Hair,
    MediumDashed,
    DashDot,
    MediumDashDot,
    DashDotDot,
    MediumDashDotDot,
    SlantDashDot,
}

impl BorderLineStyle {
    fn to_xml_value(self) -> &'static str {
        match self {
            BorderLineStyle::None => "none",
            BorderLineStyle::Thin => "thin",
            BorderLineStyle::Medium => "medium",
            BorderLineStyle::Dashed => "dashed",
            BorderLineStyle::Dotted => "dotted",
            BorderLineStyle::Thick => "thick",
            BorderLineStyle::Double => "double",
            BorderLineStyle::Hair => "hair",
            BorderLineStyle::MediumDashed => "mediumDashed",
            BorderLineStyle::DashDot => "dashDot",
            BorderLineStyle::MediumDashDot => "mediumDashDot",
            BorderLineStyle::DashDotDot => "dashDotDot",
            BorderLineStyle::MediumDashDotDot => "mediumDashDotDot",
            BorderLineStyle::SlantDashDot => "slantDashDot",
        }
    }
}

/// Horizontal alignment.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub enum HorizontalAlignment {
    #[default]
    General,
    Left,
    Center,
    Right,
    Fill,
    Justify,
    CenterContinuous,
    Distributed,
}

impl HorizontalAlignment {
    fn to_xml_value(self) -> &'static str {
        match self {
            HorizontalAlignment::General => "general",
            HorizontalAlignment::Left => "left",
            HorizontalAlignment::Center => "center",
            HorizontalAlignment::Right => "right",
            HorizontalAlignment::Fill => "fill",
            HorizontalAlignment::Justify => "justify",
            HorizontalAlignment::CenterContinuous => "centerContinuous",
            HorizontalAlignment::Distributed => "distributed",
        }
    }
}

/// Vertical alignment.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub enum VerticalAlignment {
    Top,
    Center,
    #[default]
    Bottom,
    Justify,
    Distributed,
}

impl VerticalAlignment {
    fn to_xml_value(self) -> &'static str {
        match self {
            VerticalAlignment::Top => "top",
            VerticalAlignment::Center => "center",
            VerticalAlignment::Bottom => "bottom",
            VerticalAlignment::Justify => "justify",
            VerticalAlignment::Distributed => "distributed",
        }
    }
}

/// A cell being built in a sheet.
#[derive(Debug, Clone)]
struct BuilderCell {
    value: WriteCellValue,
    style: Option<CellStyle>,
}

/// Column width definition for writing.
#[derive(Debug, Clone)]
struct ColumnWidth {
    min: u32,
    max: u32,
    width: f64,
}

/// Row height definition for writing.
#[derive(Debug, Clone)]
struct RowHeight {
    row: u32,
    height: f64,
}

/// A sheet being built.
#[derive(Debug)]
pub struct SheetBuilder {
    name: String,
    cells: HashMap<(u32, u32), BuilderCell>,
    merged_cells: Vec<String>,
    column_widths: Vec<ColumnWidth>,
    row_heights: Vec<RowHeight>,
}

impl SheetBuilder {
    fn new(name: impl Into<String>) -> Self {
        Self {
            name: name.into(),
            cells: HashMap::new(),
            merged_cells: Vec::new(),
            column_widths: Vec::new(),
            row_heights: Vec::new(),
        }
    }

    /// Set a cell value by reference (e.g., "A1", "B2").
    pub fn set_cell(&mut self, reference: &str, value: impl Into<WriteCellValue>) {
        if let Some((row, col)) = parse_cell_reference(reference) {
            self.cells.insert(
                (row, col),
                BuilderCell {
                    value: value.into(),
                    style: None,
                },
            );
        }
    }

    /// Set a cell value with a style by reference.
    pub fn set_cell_styled(
        &mut self,
        reference: &str,
        value: impl Into<WriteCellValue>,
        style: CellStyle,
    ) {
        if let Some((row, col)) = parse_cell_reference(reference) {
            self.cells.insert(
                (row, col),
                BuilderCell {
                    value: value.into(),
                    style: Some(style),
                },
            );
        }
    }

    /// Set a cell value by row and column (1-based).
    pub fn set_cell_at(&mut self, row: u32, col: u32, value: impl Into<WriteCellValue>) {
        self.cells.insert(
            (row, col),
            BuilderCell {
                value: value.into(),
                style: None,
            },
        );
    }

    /// Set a cell value with a style by row and column.
    pub fn set_cell_at_styled(
        &mut self,
        row: u32,
        col: u32,
        value: impl Into<WriteCellValue>,
        style: CellStyle,
    ) {
        self.cells.insert(
            (row, col),
            BuilderCell {
                value: value.into(),
                style: Some(style),
            },
        );
    }

    /// Apply a style to an existing cell.
    pub fn set_cell_style(&mut self, reference: &str, style: CellStyle) {
        if let Some((row, col)) = parse_cell_reference(reference)
            && let Some(cell) = self.cells.get_mut(&(row, col))
        {
            cell.style = Some(style);
        }
    }

    /// Set a formula in a cell.
    pub fn set_formula(&mut self, reference: &str, formula: impl Into<String>) {
        if let Some((row, col)) = parse_cell_reference(reference) {
            self.cells.insert(
                (row, col),
                BuilderCell {
                    value: WriteCellValue::Formula(formula.into()),
                    style: None,
                },
            );
        }
    }

    /// Set a formula with a style in a cell.
    pub fn set_formula_styled(
        &mut self,
        reference: &str,
        formula: impl Into<String>,
        style: CellStyle,
    ) {
        if let Some((row, col)) = parse_cell_reference(reference) {
            self.cells.insert(
                (row, col),
                BuilderCell {
                    value: WriteCellValue::Formula(formula.into()),
                    style: Some(style),
                },
            );
        }
    }

    /// Merge cells in a range (e.g., "A1:B2").
    ///
    /// Note: The value of the merged cell should be set in the top-left cell.
    pub fn merge_cells(&mut self, range: &str) {
        self.merged_cells.push(range.to_string());
    }

    /// Set the width of a column (in character units, Excel default is ~8.43).
    ///
    /// Column is specified by letter (e.g., "A", "B", "AA").
    pub fn set_column_width(&mut self, col: &str, width: f64) {
        if let Some(col_num) = column_letter_to_number(col) {
            self.column_widths.push(ColumnWidth {
                min: col_num,
                max: col_num,
                width,
            });
        }
    }

    /// Set the width of a range of columns.
    ///
    /// Columns are specified by letter (e.g., "A:C" for columns A through C).
    pub fn set_column_width_range(&mut self, start_col: &str, end_col: &str, width: f64) {
        if let (Some(min), Some(max)) = (
            column_letter_to_number(start_col),
            column_letter_to_number(end_col),
        ) {
            self.column_widths.push(ColumnWidth { min, max, width });
        }
    }

    /// Set the height of a row (in points, Excel default is ~15).
    pub fn set_row_height(&mut self, row: u32, height: f64) {
        self.row_heights.push(RowHeight { row, height });
    }

    /// Get the sheet name.
    pub fn name(&self) -> &str {
        &self.name
    }
}

/// Builder for creating Excel workbooks.
#[derive(Debug)]
pub struct WorkbookBuilder {
    sheets: Vec<SheetBuilder>,
    shared_strings: Vec<String>,
    string_index: HashMap<String, usize>,
    // Style collections (populated during write)
    fonts: Vec<FontStyle>,
    font_index: HashMap<FontStyleKey, usize>,
    fills: Vec<FillStyle>,
    fill_index: HashMap<FillStyleKey, usize>,
    borders: Vec<BorderStyle>,
    border_index: HashMap<BorderStyleKey, usize>,
    number_formats: Vec<String>,
    number_format_index: HashMap<String, u32>,
    cell_formats: Vec<CellFormatRecord>,
    cell_format_index: HashMap<CellFormatKey, usize>,
}

// Helper types for style deduplication
#[derive(Debug, Clone, Hash, PartialEq, Eq)]
struct FontStyleKey {
    name: Option<String>,
    size_bits: Option<u64>, // f64 as bits for hashing
    bold: bool,
    italic: bool,
    underline: Option<String>,
    strikethrough: bool,
    color: Option<String>,
}

impl From<&FontStyle> for FontStyleKey {
    fn from(f: &FontStyle) -> Self {
        Self {
            name: f.name.clone(),
            size_bits: f.size.map(|s| s.to_bits()),
            bold: f.bold,
            italic: f.italic,
            underline: f.underline.map(|u| u.to_xml_value().to_string()),
            strikethrough: f.strikethrough,
            color: f.color.clone(),
        }
    }
}

#[derive(Debug, Clone, Hash, PartialEq, Eq)]
struct FillStyleKey {
    pattern: String,
    fg_color: Option<String>,
    bg_color: Option<String>,
}

impl From<&FillStyle> for FillStyleKey {
    fn from(f: &FillStyle) -> Self {
        Self {
            pattern: f.pattern.to_xml_value().to_string(),
            fg_color: f.fg_color.clone(),
            bg_color: f.bg_color.clone(),
        }
    }
}

#[derive(Debug, Clone, Hash, PartialEq, Eq)]
struct BorderStyleKey {
    left: Option<(String, Option<String>)>,
    right: Option<(String, Option<String>)>,
    top: Option<(String, Option<String>)>,
    bottom: Option<(String, Option<String>)>,
}

impl From<&BorderStyle> for BorderStyleKey {
    fn from(b: &BorderStyle) -> Self {
        Self {
            left: b
                .left
                .as_ref()
                .map(|s| (s.style.to_xml_value().to_string(), s.color.clone())),
            right: b
                .right
                .as_ref()
                .map(|s| (s.style.to_xml_value().to_string(), s.color.clone())),
            top: b
                .top
                .as_ref()
                .map(|s| (s.style.to_xml_value().to_string(), s.color.clone())),
            bottom: b
                .bottom
                .as_ref()
                .map(|s| (s.style.to_xml_value().to_string(), s.color.clone())),
        }
    }
}

#[derive(Debug, Clone, Hash, PartialEq, Eq)]
struct CellFormatKey {
    font_id: usize,
    fill_id: usize,
    border_id: usize,
    num_fmt_id: u32,
    horizontal: Option<String>,
    vertical: Option<String>,
    wrap_text: bool,
}

#[derive(Debug, Clone)]
struct CellFormatRecord {
    font_id: usize,
    fill_id: usize,
    border_id: usize,
    num_fmt_id: u32,
    horizontal: Option<HorizontalAlignment>,
    vertical: Option<VerticalAlignment>,
    wrap_text: bool,
}

impl Default for WorkbookBuilder {
    fn default() -> Self {
        Self::new()
    }
}

impl WorkbookBuilder {
    /// Create a new workbook builder.
    pub fn new() -> Self {
        Self {
            sheets: Vec::new(),
            shared_strings: Vec::new(),
            string_index: HashMap::new(),
            fonts: Vec::new(),
            font_index: HashMap::new(),
            fills: Vec::new(),
            fill_index: HashMap::new(),
            borders: Vec::new(),
            border_index: HashMap::new(),
            number_formats: Vec::new(),
            number_format_index: HashMap::new(),
            cell_formats: Vec::new(),
            cell_format_index: HashMap::new(),
        }
    }

    /// Add a new sheet to the workbook.
    pub fn add_sheet(&mut self, name: impl Into<String>) -> &mut SheetBuilder {
        self.sheets.push(SheetBuilder::new(name));
        self.sheets.last_mut().unwrap()
    }

    /// Get a mutable reference to a sheet by index.
    pub fn sheet_mut(&mut self, index: usize) -> Option<&mut SheetBuilder> {
        self.sheets.get_mut(index)
    }

    /// Get the number of sheets.
    pub fn sheet_count(&self) -> usize {
        self.sheets.len()
    }

    /// Save the workbook to a file.
    pub fn save<P: AsRef<Path>>(self, path: P) -> Result<()> {
        let file = File::create(path)?;
        let writer = BufWriter::new(file);
        self.write(writer)
    }

    /// Write the workbook to a writer.
    pub fn write<W: Write + Seek>(mut self, writer: W) -> Result<()> {
        // Collect all strings and styles first
        self.collect_shared_strings();
        self.collect_styles();

        let has_styles = !self.cell_formats.is_empty();

        let mut pkg = PackageWriter::new(writer);

        // Add default content types
        pkg.add_default_content_type("rels", CT_RELATIONSHIPS);
        pkg.add_default_content_type("xml", CT_XML);

        // Build root relationships
        let root_rels = format!(
            r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="{}" Target="xl/workbook.xml"/>
</Relationships>"#,
            REL_OFFICE_DOCUMENT
        );

        // Build workbook relationships
        let mut wb_rels = String::new();
        wb_rels.push_str(r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>"#);
        wb_rels.push('\n');
        wb_rels.push_str(
            r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">"#,
        );
        wb_rels.push('\n');

        let mut next_rel_id = 1;
        for (i, _sheet) in self.sheets.iter().enumerate() {
            wb_rels.push_str(&format!(
                r#"  <Relationship Id="rId{}" Type="{}" Target="worksheets/sheet{}.xml"/>"#,
                next_rel_id,
                REL_WORKSHEET,
                i + 1
            ));
            wb_rels.push('\n');
            next_rel_id += 1;
        }

        // Add styles relationship if we have styles
        if has_styles {
            wb_rels.push_str(&format!(
                r#"  <Relationship Id="rId{}" Type="{}" Target="styles.xml"/>"#,
                next_rel_id, REL_STYLES
            ));
            wb_rels.push('\n');
            next_rel_id += 1;
        }

        // Add shared strings relationship if we have strings
        if !self.shared_strings.is_empty() {
            wb_rels.push_str(&format!(
                r#"  <Relationship Id="rId{}" Type="{}" Target="sharedStrings.xml"/>"#,
                next_rel_id, REL_SHARED_STRINGS
            ));
            wb_rels.push('\n');
        }

        wb_rels.push_str("</Relationships>");

        // Build workbook.xml
        let mut workbook_xml = String::new();
        workbook_xml.push_str(r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>"#);
        workbook_xml.push('\n');
        workbook_xml.push_str(&format!(
            r#"<workbook xmlns="{}" xmlns:r="{}">"#,
            NS_SPREADSHEET, NS_RELATIONSHIPS
        ));
        workbook_xml.push('\n');
        workbook_xml.push_str("  <sheets>\n");

        for (i, sheet) in self.sheets.iter().enumerate() {
            let sheet_id = i + 1;
            let rel_id = i + 1;
            workbook_xml.push_str(&format!(
                r#"    <sheet name="{}" sheetId="{}" r:id="rId{}"/>"#,
                escape_xml(&sheet.name),
                sheet_id,
                rel_id
            ));
            workbook_xml.push('\n');
        }

        workbook_xml.push_str("  </sheets>\n");
        workbook_xml.push_str("</workbook>");

        // Write parts to package
        pkg.add_part("_rels/.rels", CT_RELATIONSHIPS, root_rels.as_bytes())?;
        pkg.add_part(
            "xl/_rels/workbook.xml.rels",
            CT_RELATIONSHIPS,
            wb_rels.as_bytes(),
        )?;
        pkg.add_part("xl/workbook.xml", CT_WORKBOOK, workbook_xml.as_bytes())?;

        // Write styles if any
        if has_styles {
            let styles_xml = self.serialize_styles();
            pkg.add_part("xl/styles.xml", CT_STYLES, styles_xml.as_bytes())?;
        }

        // Write each sheet
        for (i, sheet) in self.sheets.iter().enumerate() {
            let sheet_num = i + 1;
            let sheet_xml = self.serialize_sheet(sheet);
            let part_name = format!("xl/worksheets/sheet{}.xml", sheet_num);
            pkg.add_part(&part_name, CT_WORKSHEET, sheet_xml.as_bytes())?;
        }

        // Write shared strings if any
        if !self.shared_strings.is_empty() {
            let ss_xml = self.serialize_shared_strings();
            pkg.add_part("xl/sharedStrings.xml", CT_SHARED_STRINGS, ss_xml.as_bytes())?;
        }

        pkg.finish()?;
        Ok(())
    }

    /// Collect all strings from cells into shared string table.
    fn collect_shared_strings(&mut self) {
        for sheet in &self.sheets {
            for cell in sheet.cells.values() {
                if let WriteCellValue::String(s) = &cell.value
                    && !self.string_index.contains_key(s)
                {
                    let idx = self.shared_strings.len();
                    self.shared_strings.push(s.clone());
                    self.string_index.insert(s.clone(), idx);
                }
            }
        }
    }

    /// Collect all styles from cells and build style indices.
    fn collect_styles(&mut self) {
        // Add default font (required by Excel)
        let default_font = FontStyle::new().with_name("Calibri").with_size(11.0);
        self.get_or_add_font(&default_font);

        // Add required default fills (required by Excel: none and gray125)
        let none_fill = FillStyle::new();
        let gray_fill = FillStyle::new().with_pattern(FillPattern::Gray125);
        self.get_or_add_fill(&none_fill);
        self.get_or_add_fill(&gray_fill);

        // Add default border (required by Excel)
        let default_border = BorderStyle::new();
        self.get_or_add_border(&default_border);

        // First collect all styles into a Vec to avoid borrow issues
        let styles: Vec<CellStyle> = self
            .sheets
            .iter()
            .flat_map(|sheet| sheet.cells.values())
            .filter_map(|cell| cell.style.clone())
            .collect();

        // Then add them to the style collections
        for style in &styles {
            self.get_or_add_cell_format(style);
        }
    }

    /// Get or add a font, returning its index.
    fn get_or_add_font(&mut self, font: &FontStyle) -> usize {
        let key = FontStyleKey::from(font);
        if let Some(&idx) = self.font_index.get(&key) {
            return idx;
        }
        let idx = self.fonts.len();
        self.fonts.push(font.clone());
        self.font_index.insert(key, idx);
        idx
    }

    /// Get or add a fill, returning its index.
    fn get_or_add_fill(&mut self, fill: &FillStyle) -> usize {
        let key = FillStyleKey::from(fill);
        if let Some(&idx) = self.fill_index.get(&key) {
            return idx;
        }
        let idx = self.fills.len();
        self.fills.push(fill.clone());
        self.fill_index.insert(key, idx);
        idx
    }

    /// Get or add a border, returning its index.
    fn get_or_add_border(&mut self, border: &BorderStyle) -> usize {
        let key = BorderStyleKey::from(border);
        if let Some(&idx) = self.border_index.get(&key) {
            return idx;
        }
        let idx = self.borders.len();
        self.borders.push(border.clone());
        self.border_index.insert(key, idx);
        idx
    }

    /// Get or add a number format, returning its ID.
    fn get_or_add_number_format(&mut self, format: &str) -> u32 {
        if let Some(&id) = self.number_format_index.get(format) {
            return id;
        }
        // Custom number formats start at 164
        let id = 164 + self.number_formats.len() as u32;
        self.number_formats.push(format.to_string());
        self.number_format_index.insert(format.to_string(), id);
        id
    }

    /// Get or add a cell format, returning its index (xfId).
    fn get_or_add_cell_format(&mut self, style: &CellStyle) -> usize {
        let font_id = style.font.as_ref().map_or(0, |f| self.get_or_add_font(f));
        let fill_id = style.fill.as_ref().map_or(0, |f| self.get_or_add_fill(f));
        let border_id = style
            .border
            .as_ref()
            .map_or(0, |b| self.get_or_add_border(b));
        let num_fmt_id = style
            .number_format
            .as_ref()
            .map_or(0, |f| self.get_or_add_number_format(f));

        let key = CellFormatKey {
            font_id,
            fill_id,
            border_id,
            num_fmt_id,
            horizontal: style
                .horizontal_alignment
                .map(|a| a.to_xml_value().to_string()),
            vertical: style
                .vertical_alignment
                .map(|a| a.to_xml_value().to_string()),
            wrap_text: style.wrap_text,
        };

        if let Some(&idx) = self.cell_format_index.get(&key) {
            return idx;
        }

        let record = CellFormatRecord {
            font_id,
            fill_id,
            border_id,
            num_fmt_id,
            horizontal: style.horizontal_alignment,
            vertical: style.vertical_alignment,
            wrap_text: style.wrap_text,
        };

        let idx = self.cell_formats.len();
        self.cell_formats.push(record);
        self.cell_format_index.insert(key, idx);
        idx
    }

    /// Get the style index for a cell (returns 0 if no style, or actual index + 1).
    fn get_cell_style_index(&self, style: &Option<CellStyle>) -> Option<usize> {
        style.as_ref().map(|s| {
            let font_id = s.font.as_ref().map_or(0, |f| {
                let key = FontStyleKey::from(f);
                *self.font_index.get(&key).unwrap_or(&0)
            });
            let fill_id = s.fill.as_ref().map_or(0, |f| {
                let key = FillStyleKey::from(f);
                *self.fill_index.get(&key).unwrap_or(&0)
            });
            let border_id = s.border.as_ref().map_or(0, |b| {
                let key = BorderStyleKey::from(b);
                *self.border_index.get(&key).unwrap_or(&0)
            });
            let num_fmt_id = s
                .number_format
                .as_ref()
                .map_or(0, |f| *self.number_format_index.get(f).unwrap_or(&0));

            let key = CellFormatKey {
                font_id,
                fill_id,
                border_id,
                num_fmt_id,
                horizontal: s.horizontal_alignment.map(|a| a.to_xml_value().to_string()),
                vertical: s.vertical_alignment.map(|a| a.to_xml_value().to_string()),
                wrap_text: s.wrap_text,
            };

            // Return index + 1 because index 0 is reserved for default style
            self.cell_format_index.get(&key).map_or(0, |&idx| idx + 1)
        })
    }

    /// Serialize styles to XML.
    fn serialize_styles(&self) -> String {
        let mut xml = String::new();
        xml.push_str(r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>"#);
        xml.push('\n');
        xml.push_str(&format!(r#"<styleSheet xmlns="{}">"#, NS_SPREADSHEET));
        xml.push('\n');

        // Number formats
        if !self.number_formats.is_empty() {
            xml.push_str(&format!(
                "  <numFmts count=\"{}\">\n",
                self.number_formats.len()
            ));
            for (i, fmt) in self.number_formats.iter().enumerate() {
                let id = 164 + i as u32;
                xml.push_str(&format!(
                    r#"    <numFmt numFmtId="{}" formatCode="{}"/>"#,
                    id,
                    escape_xml(fmt)
                ));
                xml.push('\n');
            }
            xml.push_str("  </numFmts>\n");
        }

        // Fonts
        xml.push_str(&format!("  <fonts count=\"{}\">\n", self.fonts.len()));
        for font in &self.fonts {
            xml.push_str("    <font>\n");
            if font.bold {
                xml.push_str("      <b/>\n");
            }
            if font.italic {
                xml.push_str("      <i/>\n");
            }
            if font.strikethrough {
                xml.push_str("      <strike/>\n");
            }
            if let Some(u) = &font.underline {
                xml.push_str(&format!(r#"      <u val="{}"/>"#, u.to_xml_value()));
                xml.push('\n');
            }
            if let Some(size) = font.size {
                xml.push_str(&format!(r#"      <sz val="{}"/>"#, size));
                xml.push('\n');
            }
            if let Some(color) = &font.color {
                xml.push_str(&format!(r#"      <color rgb="FF{}"/>"#, color));
                xml.push('\n');
            }
            if let Some(name) = &font.name {
                xml.push_str(&format!(r#"      <name val="{}"/>"#, escape_xml(name)));
                xml.push('\n');
            }
            xml.push_str("    </font>\n");
        }
        xml.push_str("  </fonts>\n");

        // Fills
        xml.push_str(&format!("  <fills count=\"{}\">\n", self.fills.len()));
        for fill in &self.fills {
            xml.push_str("    <fill>\n");
            xml.push_str(&format!(
                r#"      <patternFill patternType="{}">"#,
                fill.pattern.to_xml_value()
            ));
            xml.push('\n');
            if let Some(fg) = &fill.fg_color {
                xml.push_str(&format!(r#"        <fgColor rgb="FF{}"/>"#, fg));
                xml.push('\n');
            }
            if let Some(bg) = &fill.bg_color {
                xml.push_str(&format!(r#"        <bgColor rgb="FF{}"/>"#, bg));
                xml.push('\n');
            }
            xml.push_str("      </patternFill>\n");
            xml.push_str("    </fill>\n");
        }
        xml.push_str("  </fills>\n");

        // Borders
        xml.push_str(&format!("  <borders count=\"{}\">\n", self.borders.len()));
        for border in &self.borders {
            let mut diagonal_attrs = String::new();
            if border.diagonal_up {
                diagonal_attrs.push_str(r#" diagonalUp="1""#);
            }
            if border.diagonal_down {
                diagonal_attrs.push_str(r#" diagonalDown="1""#);
            }
            xml.push_str(&format!("    <border{}>\n", diagonal_attrs));

            // Left
            self.serialize_border_side(&mut xml, "left", &border.left);
            // Right
            self.serialize_border_side(&mut xml, "right", &border.right);
            // Top
            self.serialize_border_side(&mut xml, "top", &border.top);
            // Bottom
            self.serialize_border_side(&mut xml, "bottom", &border.bottom);
            // Diagonal
            self.serialize_border_side(&mut xml, "diagonal", &border.diagonal);

            xml.push_str("    </border>\n");
        }
        xml.push_str("  </borders>\n");

        // Cell style XFs (required, at least one default)
        xml.push_str("  <cellStyleXfs count=\"1\">\n");
        xml.push_str("    <xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\"/>\n");
        xml.push_str("  </cellStyleXfs>\n");

        // Cell XFs
        let xf_count = self.cell_formats.len() + 1; // +1 for default
        xml.push_str(&format!("  <cellXfs count=\"{}\">\n", xf_count));
        // Default format
        xml.push_str(
            "    <xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\" xfId=\"0\"/>\n",
        );

        for xf in &self.cell_formats {
            let mut attrs = format!(
                r#"numFmtId="{}" fontId="{}" fillId="{}" borderId="{}" xfId="0""#,
                xf.num_fmt_id, xf.font_id, xf.fill_id, xf.border_id
            );

            if xf.font_id > 0 {
                attrs.push_str(r#" applyFont="1""#);
            }
            if xf.fill_id > 0 {
                attrs.push_str(r#" applyFill="1""#);
            }
            if xf.border_id > 0 {
                attrs.push_str(r#" applyBorder="1""#);
            }
            if xf.num_fmt_id > 0 {
                attrs.push_str(r#" applyNumberFormat="1""#);
            }

            let has_alignment = xf.horizontal.is_some() || xf.vertical.is_some() || xf.wrap_text;
            if has_alignment {
                attrs.push_str(r#" applyAlignment="1""#);
                xml.push_str(&format!("    <xf {}>\n", attrs));

                let mut align_attrs = Vec::new();
                if let Some(h) = xf.horizontal {
                    align_attrs.push(format!(r#"horizontal="{}""#, h.to_xml_value()));
                }
                if let Some(v) = xf.vertical {
                    align_attrs.push(format!(r#"vertical="{}""#, v.to_xml_value()));
                }
                if xf.wrap_text {
                    align_attrs.push(r#"wrapText="1""#.to_string());
                }
                xml.push_str(&format!("      <alignment {}/>\n", align_attrs.join(" ")));
                xml.push_str("    </xf>\n");
            } else {
                xml.push_str(&format!("    <xf {}/>\n", attrs));
            }
        }
        xml.push_str("  </cellXfs>\n");

        // Cell styles (required)
        xml.push_str("  <cellStyles count=\"1\">\n");
        xml.push_str("    <cellStyle name=\"Normal\" xfId=\"0\" builtinId=\"0\"/>\n");
        xml.push_str("  </cellStyles>\n");

        xml.push_str("</styleSheet>");
        xml
    }

    /// Serialize a border side element.
    fn serialize_border_side(&self, xml: &mut String, name: &str, side: &Option<BorderSideStyle>) {
        if let Some(s) = side {
            if s.style != BorderLineStyle::None {
                xml.push_str(&format!(
                    r#"      <{} style="{}">"#,
                    name,
                    s.style.to_xml_value()
                ));
                xml.push('\n');
                if let Some(color) = &s.color {
                    xml.push_str(&format!(r#"        <color rgb="FF{}"/>"#, color));
                    xml.push('\n');
                }
                xml.push_str(&format!("      </{}>\n", name));
            } else {
                xml.push_str(&format!("      <{}/>\n", name));
            }
        } else {
            xml.push_str(&format!("      <{}/>\n", name));
        }
    }

    /// Serialize a sheet to XML.
    fn serialize_sheet(&self, sheet: &SheetBuilder) -> String {
        let mut xml = String::new();
        xml.push_str(r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>"#);
        xml.push('\n');
        xml.push_str(&format!(r#"<worksheet xmlns="{}">"#, NS_SPREADSHEET));
        xml.push('\n');

        // Write column widths if any
        if !sheet.column_widths.is_empty() {
            xml.push_str("  <cols>\n");
            for col in &sheet.column_widths {
                xml.push_str(&format!(
                    r#"    <col min="{}" max="{}" width="{}" customWidth="1"/>"#,
                    col.min, col.max, col.width
                ));
                xml.push('\n');
            }
            xml.push_str("  </cols>\n");
        }

        xml.push_str("  <sheetData>\n");

        // Group cells by row
        let mut rows: HashMap<u32, Vec<(u32, &BuilderCell)>> = HashMap::new();
        for ((row, col), cell) in &sheet.cells {
            rows.entry(*row).or_default().push((*col, cell));
        }

        // Build row height lookup
        let row_height_map: HashMap<u32, f64> = sheet
            .row_heights
            .iter()
            .map(|rh| (rh.row, rh.height))
            .collect();

        // Sort rows and serialize
        let mut row_nums: Vec<_> = rows.keys().copied().collect();
        row_nums.sort();

        for row_num in row_nums {
            let cells = rows.get(&row_num).unwrap();
            let mut sorted_cells: Vec<_> = cells.clone();
            sorted_cells.sort_by_key(|(col, _)| *col);

            // Include row height if set
            if let Some(height) = row_height_map.get(&row_num) {
                xml.push_str(&format!(
                    r#"    <row r="{}" ht="{}" customHeight="1">"#,
                    row_num, height
                ));
            } else {
                xml.push_str(&format!(r#"    <row r="{}">"#, row_num));
            }
            xml.push('\n');

            for (col, cell) in sorted_cells {
                let ref_str = column_to_letter(col) + &row_num.to_string();
                xml.push_str(&self.serialize_cell(&ref_str, cell));
            }

            xml.push_str("    </row>\n");
        }

        xml.push_str("  </sheetData>\n");

        // Write merged cells if any
        if !sheet.merged_cells.is_empty() {
            xml.push_str(&format!(
                "  <mergeCells count=\"{}\">\n",
                sheet.merged_cells.len()
            ));
            for range in &sheet.merged_cells {
                xml.push_str(&format!(r#"    <mergeCell ref="{}"/>"#, escape_xml(range)));
                xml.push('\n');
            }
            xml.push_str("  </mergeCells>\n");
        }

        xml.push_str("</worksheet>");
        xml
    }

    /// Serialize a cell to XML.
    fn serialize_cell(&self, reference: &str, cell: &BuilderCell) -> String {
        let style_attr = self
            .get_cell_style_index(&cell.style)
            .filter(|&s| s > 0)
            .map(|s| format!(r#" s="{}""#, s))
            .unwrap_or_default();

        match &cell.value {
            WriteCellValue::String(s) => {
                let idx = self.string_index.get(s).unwrap_or(&0);
                format!(
                    r#"      <c r="{}" t="s"{}><v>{}</v></c>"#,
                    reference, style_attr, idx
                ) + "\n"
            }
            WriteCellValue::Number(n) => {
                format!(
                    r#"      <c r="{}"{}><v>{}</v></c>"#,
                    reference, style_attr, n
                ) + "\n"
            }
            WriteCellValue::Boolean(b) => {
                let val = if *b { "1" } else { "0" };
                format!(
                    r#"      <c r="{}" t="b"{}><v>{}</v></c>"#,
                    reference, style_attr, val
                ) + "\n"
            }
            WriteCellValue::Formula(f) => {
                format!(
                    r#"      <c r="{}"{}><f>{}</f></c>"#,
                    reference,
                    style_attr,
                    escape_xml(f)
                ) + "\n"
            }
            WriteCellValue::Empty => {
                if !style_attr.is_empty() {
                    format!(r#"      <c r="{}"{}></c>"#, reference, style_attr) + "\n"
                } else {
                    String::new()
                }
            }
        }
    }

    /// Serialize shared strings table to XML.
    fn serialize_shared_strings(&self) -> String {
        let mut xml = String::new();
        xml.push_str(r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>"#);
        xml.push('\n');
        xml.push_str(&format!(
            r#"<sst xmlns="{}" count="{}" uniqueCount="{}">"#,
            NS_SPREADSHEET,
            self.shared_strings.len(),
            self.shared_strings.len()
        ));
        xml.push('\n');

        for s in &self.shared_strings {
            xml.push_str(&format!("  <si><t>{}</t></si>\n", escape_xml(s)));
        }

        xml.push_str("</sst>");
        xml
    }
}

/// Parse a cell reference like "A1" into (row, col).
fn parse_cell_reference(reference: &str) -> Option<(u32, u32)> {
    let mut col_part = String::new();
    let mut row_part = String::new();

    for c in reference.chars() {
        if c.is_ascii_alphabetic() {
            col_part.push(c.to_ascii_uppercase());
        } else if c.is_ascii_digit() {
            row_part.push(c);
        }
    }

    let col = column_letter_to_number(&col_part)?;
    let row: u32 = row_part.parse().ok()?;

    Some((row, col))
}

/// Convert column letters to number (A=1, B=2, ..., Z=26, AA=27).
fn column_letter_to_number(letters: &str) -> Option<u32> {
    if letters.is_empty() {
        return None;
    }

    let mut result: u32 = 0;
    for c in letters.chars() {
        if !c.is_ascii_alphabetic() {
            return None;
        }
        result = result * 26 + (c.to_ascii_uppercase() as u32 - 'A' as u32 + 1);
    }
    Some(result)
}

/// Convert column number to letters (1=A, 2=B, ..., 26=Z, 27=AA).
fn column_to_letter(mut col: u32) -> String {
    let mut result = String::new();
    while col > 0 {
        col -= 1;
        result.insert(0, (b'A' + (col % 26) as u8) as char);
        col /= 26;
    }
    result
}

/// Escape XML special characters.
fn escape_xml(s: &str) -> String {
    s.replace('&', "&amp;")
        .replace('<', "&lt;")
        .replace('>', "&gt;")
        .replace('"', "&quot;")
        .replace('\'', "&apos;")
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
    }

    #[test]
    fn test_column_to_letter() {
        assert_eq!(column_to_letter(1), "A");
        assert_eq!(column_to_letter(2), "B");
        assert_eq!(column_to_letter(26), "Z");
        assert_eq!(column_to_letter(27), "AA");
        assert_eq!(column_to_letter(28), "AB");
        assert_eq!(column_to_letter(52), "AZ");
    }

    #[test]
    fn test_parse_cell_reference() {
        assert_eq!(parse_cell_reference("A1"), Some((1, 1)));
        assert_eq!(parse_cell_reference("B2"), Some((2, 2)));
        assert_eq!(parse_cell_reference("Z10"), Some((10, 26)));
        assert_eq!(parse_cell_reference("AA1"), Some((1, 27)));
    }

    #[test]
    fn test_workbook_builder() {
        let mut wb = WorkbookBuilder::new();
        let sheet = wb.add_sheet("Test");
        sheet.set_cell("A1", "Hello");
        sheet.set_cell("B1", 42.0);
        sheet.set_cell("C1", true);
        sheet.set_formula("D1", "SUM(A1:C1)");

        assert_eq!(wb.sheet_count(), 1);
    }

    #[test]
    fn test_roundtrip_simple() {
        use std::io::Cursor;

        let mut wb = WorkbookBuilder::new();
        let sheet = wb.add_sheet("Sheet1");
        sheet.set_cell("A1", "Test Value");
        sheet.set_cell("B1", 123.45);

        // Write to memory
        let mut buffer = Cursor::new(Vec::new());
        wb.write(&mut buffer).unwrap();

        // Read back
        buffer.set_position(0);
        let mut workbook = crate::Workbook::from_reader(buffer).unwrap();
        let read_sheet = workbook.sheet(0).unwrap();

        assert_eq!(read_sheet.name(), "Sheet1");

        let cell_a1 = read_sheet.cell("A1").unwrap();
        assert_eq!(cell_a1.value_as_string(), "Test Value");

        let cell_b1 = read_sheet.cell("B1").unwrap();
        assert_eq!(cell_b1.value_as_number(), Some(123.45));
    }

    #[test]
    fn test_roundtrip_merged_cells() {
        use std::io::Cursor;

        let mut wb = WorkbookBuilder::new();
        let sheet = wb.add_sheet("Sheet1");
        sheet.set_cell("A1", "Merged Header");
        sheet.merge_cells("A1:C1");
        sheet.set_cell("A2", "Data 1");
        sheet.set_cell("B2", "Data 2");
        sheet.merge_cells("A3:B4");
        sheet.set_cell("A3", "Block");

        // Write to memory
        let mut buffer = Cursor::new(Vec::new());
        wb.write(&mut buffer).unwrap();

        // Read back
        buffer.set_position(0);
        let mut workbook = crate::Workbook::from_reader(buffer).unwrap();
        let read_sheet = workbook.sheet(0).unwrap();

        // Check merged cells were preserved
        let merged = read_sheet.merged_cells();
        assert_eq!(merged.len(), 2);
        assert_eq!(merged[0].reference(), "A1:C1");
        assert_eq!(merged[1].reference(), "A3:B4");

        // Check helper methods
        assert!(merged[0].contains(1, 1)); // A1
        assert!(merged[0].contains(1, 2)); // B1
        assert!(merged[0].contains(1, 3)); // C1
        assert!(!merged[0].contains(1, 4)); // D1 - outside

        // Check start/end cells
        assert_eq!(merged[0].start_cell(), "A1");
        assert_eq!(merged[0].end_cell(), "C1");
        assert_eq!(merged[1].start_cell(), "A3");
        assert_eq!(merged[1].end_cell(), "B4");

        // Check merged_cell_at helper
        assert!(read_sheet.merged_cell_at("A1").is_some());
        assert!(read_sheet.merged_cell_at("B1").is_some());
        assert!(read_sheet.merged_cell_at("D1").is_none());

        // Cell values should still be accessible
        let cell_a1 = read_sheet.cell("A1").unwrap();
        assert_eq!(cell_a1.value_as_string(), "Merged Header");
    }

    #[test]
    fn test_roundtrip_dimensions() {
        use std::io::Cursor;

        let mut wb = WorkbookBuilder::new();
        let sheet = wb.add_sheet("Sheet1");
        sheet.set_cell("A1", "Header 1");
        sheet.set_cell("B1", "Header 2");
        sheet.set_cell("C1", "Header 3");
        sheet.set_cell("A2", "Data");

        // Set column widths
        sheet.set_column_width("A", 20.0);
        sheet.set_column_width_range("B", "C", 15.5);

        // Set row heights
        sheet.set_row_height(1, 25.0);
        sheet.set_row_height(2, 18.0);

        // Write to memory
        let mut buffer = Cursor::new(Vec::new());
        wb.write(&mut buffer).unwrap();

        // Read back
        buffer.set_position(0);
        let mut workbook = crate::Workbook::from_reader(buffer).unwrap();
        let read_sheet = workbook.sheet(0).unwrap();

        // Check column widths were preserved
        let columns = read_sheet.columns();
        assert_eq!(columns.len(), 2);

        // Column A (col 1)
        assert_eq!(columns[0].min, 1);
        assert_eq!(columns[0].max, 1);
        assert_eq!(columns[0].width, Some(20.0));

        // Columns B-C (cols 2-3)
        assert_eq!(columns[1].min, 2);
        assert_eq!(columns[1].max, 3);
        assert_eq!(columns[1].width, Some(15.5));

        // Test helper methods
        assert_eq!(read_sheet.column_width(1), Some(20.0));
        assert_eq!(read_sheet.column_width(2), Some(15.5));
        assert_eq!(read_sheet.column_width(3), Some(15.5));
        assert_eq!(read_sheet.column_width(4), None);

        // Check row heights were preserved
        let row1 = read_sheet.row(1).unwrap();
        assert_eq!(row1.height(), Some(25.0));

        let row2 = read_sheet.row(2).unwrap();
        assert_eq!(row2.height(), Some(18.0));
    }
}
