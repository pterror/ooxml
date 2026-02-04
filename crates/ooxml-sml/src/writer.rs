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
use crate::generated_serializers::ToXml;
use crate::types;
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
const CT_COMMENTS: &str =
    "application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml";
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
const REL_COMMENTS: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments";

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

/// A conditional formatting rule for writing.
#[derive(Debug, Clone)]
pub struct ConditionalFormat {
    /// Cell range the rule applies to (e.g., "A1:C10").
    pub range: String,
    /// The rules in this conditional format.
    pub rules: Vec<ConditionalFormatRule>,
}

/// A single conditional formatting rule.
#[derive(Debug, Clone)]
pub struct ConditionalFormatRule {
    /// Rule type.
    pub rule_type: crate::ConditionalRuleType,
    /// Priority (lower = higher priority).
    pub priority: u32,
    /// Differential formatting ID.
    pub dxf_id: Option<u32>,
    /// Operator for cellIs rules.
    pub operator: Option<String>,
    /// Formula(s) for the rule.
    pub formulas: Vec<String>,
    /// Text for containsText/beginsWith/endsWith rules.
    pub text: Option<String>,
}

impl ConditionalFormat {
    /// Create a new conditional format for a range.
    pub fn new(range: impl Into<String>) -> Self {
        Self {
            range: range.into(),
            rules: Vec::new(),
        }
    }

    /// Add a cell value comparison rule.
    pub fn add_cell_is_rule(
        mut self,
        operator: &str,
        formula: impl Into<String>,
        priority: u32,
        dxf_id: Option<u32>,
    ) -> Self {
        self.rules.push(ConditionalFormatRule {
            rule_type: crate::ConditionalRuleType::CellIs,
            priority,
            dxf_id,
            operator: Some(operator.to_string()),
            formulas: vec![formula.into()],
            text: None,
        });
        self
    }

    /// Add a formula expression rule.
    pub fn add_expression_rule(
        mut self,
        formula: impl Into<String>,
        priority: u32,
        dxf_id: Option<u32>,
    ) -> Self {
        self.rules.push(ConditionalFormatRule {
            rule_type: crate::ConditionalRuleType::Expression,
            priority,
            dxf_id,
            operator: None,
            formulas: vec![formula.into()],
            text: None,
        });
        self
    }

    /// Add a duplicate values rule.
    pub fn add_duplicate_values_rule(mut self, priority: u32, dxf_id: Option<u32>) -> Self {
        self.rules.push(ConditionalFormatRule {
            rule_type: crate::ConditionalRuleType::DuplicateValues,
            priority,
            dxf_id,
            operator: None,
            formulas: Vec::new(),
            text: None,
        });
        self
    }

    /// Add a contains text rule.
    pub fn add_contains_text_rule(
        mut self,
        text: impl Into<String>,
        priority: u32,
        dxf_id: Option<u32>,
    ) -> Self {
        let text = text.into();
        self.rules.push(ConditionalFormatRule {
            rule_type: crate::ConditionalRuleType::ContainsText,
            priority,
            dxf_id,
            operator: Some("containsText".to_string()),
            formulas: Vec::new(),
            text: Some(text),
        });
        self
    }
}

/// A data validation rule for writing.
#[derive(Debug, Clone)]
pub struct DataValidationBuilder {
    /// Cell range the validation applies to (e.g., "A1:C10").
    pub range: String,
    /// Validation type.
    pub validation_type: crate::DataValidationType,
    /// Comparison operator.
    pub operator: crate::DataValidationOperator,
    /// First formula/value.
    pub formula1: Option<String>,
    /// Second formula/value (for between/notBetween operators).
    pub formula2: Option<String>,
    /// Allow blank cells.
    pub allow_blank: bool,
    /// Show input message when cell is selected.
    pub show_input_message: bool,
    /// Show error message on invalid input.
    pub show_error_message: bool,
    /// Error alert style.
    pub error_style: crate::DataValidationErrorStyle,
    /// Error title.
    pub error_title: Option<String>,
    /// Error message.
    pub error_message: Option<String>,
    /// Input prompt title.
    pub prompt_title: Option<String>,
    /// Input prompt message.
    pub prompt_message: Option<String>,
}

impl DataValidationBuilder {
    /// Create a new data validation for a range.
    pub fn new(range: impl Into<String>) -> Self {
        Self {
            range: range.into(),
            validation_type: crate::DataValidationType::None,
            operator: crate::DataValidationOperator::Between,
            formula1: None,
            formula2: None,
            allow_blank: true,
            show_input_message: true,
            show_error_message: true,
            error_style: crate::DataValidationErrorStyle::Stop,
            error_title: None,
            error_message: None,
            prompt_title: None,
            prompt_message: None,
        }
    }

    /// Create a list validation (dropdown) from a range or comma-separated values.
    pub fn list(range: impl Into<String>, source: impl Into<String>) -> Self {
        Self {
            range: range.into(),
            validation_type: crate::DataValidationType::List,
            operator: crate::DataValidationOperator::Between,
            formula1: Some(source.into()),
            formula2: None,
            allow_blank: true,
            show_input_message: true,
            show_error_message: true,
            error_style: crate::DataValidationErrorStyle::Stop,
            error_title: None,
            error_message: None,
            prompt_title: None,
            prompt_message: None,
        }
    }

    /// Create a whole number validation.
    pub fn whole_number(
        range: impl Into<String>,
        operator: crate::DataValidationOperator,
        value1: impl Into<String>,
    ) -> Self {
        Self {
            range: range.into(),
            validation_type: crate::DataValidationType::Whole,
            operator,
            formula1: Some(value1.into()),
            formula2: None,
            allow_blank: true,
            show_input_message: true,
            show_error_message: true,
            error_style: crate::DataValidationErrorStyle::Stop,
            error_title: None,
            error_message: None,
            prompt_title: None,
            prompt_message: None,
        }
    }

    /// Create a decimal number validation.
    pub fn decimal(
        range: impl Into<String>,
        operator: crate::DataValidationOperator,
        value1: impl Into<String>,
    ) -> Self {
        Self {
            range: range.into(),
            validation_type: crate::DataValidationType::Decimal,
            operator,
            formula1: Some(value1.into()),
            formula2: None,
            allow_blank: true,
            show_input_message: true,
            show_error_message: true,
            error_style: crate::DataValidationErrorStyle::Stop,
            error_title: None,
            error_message: None,
            prompt_title: None,
            prompt_message: None,
        }
    }

    /// Set the second value/formula for between/notBetween operators.
    pub fn with_formula2(mut self, formula2: impl Into<String>) -> Self {
        self.formula2 = Some(formula2.into());
        self
    }

    /// Set whether blank cells are allowed.
    pub fn with_allow_blank(mut self, allow: bool) -> Self {
        self.allow_blank = allow;
        self
    }

    /// Set the error style.
    pub fn with_error_style(mut self, style: crate::DataValidationErrorStyle) -> Self {
        self.error_style = style;
        self
    }

    /// Set the error message.
    pub fn with_error(mut self, title: impl Into<String>, message: impl Into<String>) -> Self {
        self.error_title = Some(title.into());
        self.error_message = Some(message.into());
        self
    }

    /// Set the input prompt message.
    pub fn with_prompt(mut self, title: impl Into<String>, message: impl Into<String>) -> Self {
        self.prompt_title = Some(title.into());
        self.prompt_message = Some(message.into());
        self.show_input_message = true;
        self
    }
}

/// A defined name (named range) for writing.
///
/// Defined names can reference ranges, formulas, or constants.
/// They can be global (workbook scope) or local (sheet scope).
///
/// # Example
///
/// ```ignore
/// let mut wb = WorkbookBuilder::new();
/// // Global defined name
/// wb.add_defined_name("MyRange", "Sheet1!$A$1:$B$10");
/// // Sheet-scoped defined name
/// wb.add_defined_name_with_scope("LocalName", "Sheet1!$C$1", 0);
/// ```
#[derive(Debug, Clone)]
pub struct DefinedNameBuilder {
    /// The name (e.g., "MyRange", "_xlnm.Print_Area").
    pub name: String,
    /// The formula or reference (e.g., "Sheet1!$A$1:$B$10").
    pub reference: String,
    /// Optional sheet index if this name is scoped to a specific sheet.
    pub local_sheet_id: Option<u32>,
    /// Optional comment/description.
    pub comment: Option<String>,
    /// Whether this is a hidden name.
    pub hidden: bool,
}

impl DefinedNameBuilder {
    /// Create a new defined name with global scope.
    pub fn new(name: impl Into<String>, reference: impl Into<String>) -> Self {
        Self {
            name: name.into(),
            reference: reference.into(),
            local_sheet_id: None,
            comment: None,
            hidden: false,
        }
    }

    /// Create a new defined name with sheet scope.
    pub fn with_sheet_scope(
        name: impl Into<String>,
        reference: impl Into<String>,
        sheet_index: u32,
    ) -> Self {
        Self {
            name: name.into(),
            reference: reference.into(),
            local_sheet_id: Some(sheet_index),
            comment: None,
            hidden: false,
        }
    }

    /// Add a comment to the defined name.
    pub fn with_comment(mut self, comment: impl Into<String>) -> Self {
        self.comment = Some(comment.into());
        self
    }

    /// Mark the defined name as hidden.
    pub fn hidden(mut self) -> Self {
        self.hidden = true;
        self
    }

    /// Create a print area defined name for a sheet.
    ///
    /// # Example
    ///
    /// ```ignore
    /// let print_area = DefinedNameBuilder::print_area(0, "Sheet1!$A$1:$G$20");
    /// wb.add_defined_name_builder(print_area);
    /// ```
    pub fn print_area(sheet_index: u32, reference: impl Into<String>) -> Self {
        Self {
            name: "_xlnm.Print_Area".to_string(),
            reference: reference.into(),
            local_sheet_id: Some(sheet_index),
            comment: None,
            hidden: false,
        }
    }

    /// Create a print titles defined name for a sheet (repeating rows/columns).
    ///
    /// # Example
    ///
    /// ```ignore
    /// // Repeat rows 1-2 on each printed page
    /// let print_titles = DefinedNameBuilder::print_titles(0, "Sheet1!$1:$2");
    /// wb.add_defined_name_builder(print_titles);
    /// ```
    pub fn print_titles(sheet_index: u32, reference: impl Into<String>) -> Self {
        Self {
            name: "_xlnm.Print_Titles".to_string(),
            reference: reference.into(),
            local_sheet_id: Some(sheet_index),
            comment: None,
            hidden: false,
        }
    }
}

/// A cell comment (note) for writing.
///
/// # Example
///
/// ```ignore
/// let mut wb = WorkbookBuilder::new();
/// let sheet = wb.add_sheet("Sheet1");
/// sheet.add_comment("A1", "This is a comment");
/// sheet.add_comment_with_author("B1", "Another comment", "John Doe");
/// ```
#[derive(Debug, Clone)]
pub struct CommentBuilder {
    /// Cell reference (e.g., "A1").
    pub reference: String,
    /// Comment text content.
    pub text: String,
    /// Author of the comment (optional).
    pub author: Option<String>,
}

impl CommentBuilder {
    /// Create a new comment.
    pub fn new(reference: impl Into<String>, text: impl Into<String>) -> Self {
        Self {
            reference: reference.into(),
            text: text.into(),
            author: None,
        }
    }

    /// Create a new comment with an author.
    pub fn with_author(
        reference: impl Into<String>,
        text: impl Into<String>,
        author: impl Into<String>,
    ) -> Self {
        Self {
            reference: reference.into(),
            text: text.into(),
            author: Some(author.into()),
        }
    }

    /// Set the author of the comment.
    pub fn author(mut self, author: impl Into<String>) -> Self {
        self.author = Some(author.into());
        self
    }
}

/// A sheet being built.
#[derive(Debug)]
pub struct SheetBuilder {
    name: String,
    cells: HashMap<(u32, u32), BuilderCell>,
    merged_cells: Vec<String>,
    column_widths: Vec<ColumnWidth>,
    row_heights: Vec<RowHeight>,
    conditional_formats: Vec<ConditionalFormat>,
    data_validations: Vec<DataValidationBuilder>,
    comments: Vec<CommentBuilder>,
}

impl SheetBuilder {
    fn new(name: impl Into<String>) -> Self {
        Self {
            name: name.into(),
            cells: HashMap::new(),
            merged_cells: Vec::new(),
            column_widths: Vec::new(),
            row_heights: Vec::new(),
            conditional_formats: Vec::new(),
            data_validations: Vec::new(),
            comments: Vec::new(),
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

    /// Add conditional formatting to the sheet.
    pub fn add_conditional_format(&mut self, cf: ConditionalFormat) {
        self.conditional_formats.push(cf);
    }

    /// Add data validation to the sheet.
    pub fn add_data_validation(&mut self, dv: DataValidationBuilder) {
        self.data_validations.push(dv);
    }

    /// Add a comment (note) to a cell.
    ///
    /// # Example
    ///
    /// ```ignore
    /// sheet.add_comment("A1", "This is a comment");
    /// ```
    pub fn add_comment(&mut self, reference: impl Into<String>, text: impl Into<String>) {
        self.comments.push(CommentBuilder::new(reference, text));
    }

    /// Add a comment (note) with an author to a cell.
    ///
    /// # Example
    ///
    /// ```ignore
    /// sheet.add_comment_with_author("A1", "Review needed", "John Doe");
    /// ```
    pub fn add_comment_with_author(
        &mut self,
        reference: impl Into<String>,
        text: impl Into<String>,
        author: impl Into<String>,
    ) {
        self.comments
            .push(CommentBuilder::with_author(reference, text, author));
    }

    /// Add a comment using a builder for full control.
    pub fn add_comment_builder(&mut self, comment: CommentBuilder) {
        self.comments.push(comment);
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
    defined_names: Vec<DefinedNameBuilder>,
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
            defined_names: Vec::new(),
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

    /// Add a defined name (named range) with global (workbook) scope.
    ///
    /// # Example
    ///
    /// ```ignore
    /// let mut wb = WorkbookBuilder::new();
    /// wb.add_sheet("Sheet1");
    /// wb.add_defined_name("MyRange", "Sheet1!$A$1:$B$10");
    /// ```
    pub fn add_defined_name(&mut self, name: impl Into<String>, reference: impl Into<String>) {
        self.defined_names
            .push(DefinedNameBuilder::new(name, reference));
    }

    /// Add a defined name (named range) with sheet scope.
    ///
    /// Sheet-scoped names are only visible within the specified sheet.
    ///
    /// # Example
    ///
    /// ```ignore
    /// let mut wb = WorkbookBuilder::new();
    /// wb.add_sheet("Sheet1");
    /// // This name is only visible in Sheet1 (index 0)
    /// wb.add_defined_name_with_scope("LocalRange", "Sheet1!$A$1:$B$10", 0);
    /// ```
    pub fn add_defined_name_with_scope(
        &mut self,
        name: impl Into<String>,
        reference: impl Into<String>,
        sheet_index: u32,
    ) {
        self.defined_names
            .push(DefinedNameBuilder::with_sheet_scope(
                name,
                reference,
                sheet_index,
            ));
    }

    /// Add a defined name using a builder for full control.
    ///
    /// # Example
    ///
    /// ```ignore
    /// let mut wb = WorkbookBuilder::new();
    /// wb.add_sheet("Sheet1");
    ///
    /// let name = DefinedNameBuilder::new("MyRange", "Sheet1!$A$1:$B$10")
    ///     .with_comment("Sales data range")
    ///     .hidden();
    /// wb.add_defined_name_builder(name);
    /// ```
    pub fn add_defined_name_builder(&mut self, builder: DefinedNameBuilder) {
        self.defined_names.push(builder);
    }

    /// Add a print area for a sheet.
    ///
    /// # Example
    ///
    /// ```ignore
    /// let mut wb = WorkbookBuilder::new();
    /// wb.add_sheet("Sheet1");
    /// wb.set_print_area(0, "Sheet1!$A$1:$G$20");
    /// ```
    pub fn set_print_area(&mut self, sheet_index: u32, reference: impl Into<String>) {
        self.defined_names
            .push(DefinedNameBuilder::print_area(sheet_index, reference));
    }

    /// Add print titles (repeating rows/columns) for a sheet.
    ///
    /// # Example
    ///
    /// ```ignore
    /// let mut wb = WorkbookBuilder::new();
    /// wb.add_sheet("Sheet1");
    /// // Repeat rows 1-2 on each printed page
    /// wb.set_print_titles(0, "Sheet1!$1:$2");
    /// ```
    pub fn set_print_titles(&mut self, sheet_index: u32, reference: impl Into<String>) {
        self.defined_names
            .push(DefinedNameBuilder::print_titles(sheet_index, reference));
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

        // Add defined names if any
        if !self.defined_names.is_empty() {
            workbook_xml.push_str("  <definedNames>\n");
            for dn in &self.defined_names {
                workbook_xml.push_str("    <definedName name=\"");
                workbook_xml.push_str(&escape_xml(&dn.name));
                workbook_xml.push('"');
                if let Some(sheet_id) = dn.local_sheet_id {
                    workbook_xml.push_str(&format!(" localSheetId=\"{}\"", sheet_id));
                }
                if let Some(ref comment) = dn.comment {
                    workbook_xml.push_str(" comment=\"");
                    workbook_xml.push_str(&escape_xml(comment));
                    workbook_xml.push('"');
                }
                if dn.hidden {
                    workbook_xml.push_str(" hidden=\"1\"");
                }
                workbook_xml.push('>');
                workbook_xml.push_str(&escape_xml(&dn.reference));
                workbook_xml.push_str("</definedName>\n");
            }
            workbook_xml.push_str("  </definedNames>\n");
        }

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

        // Write each sheet and its related parts (comments, etc.)
        for (i, sheet) in self.sheets.iter().enumerate() {
            let sheet_num = i + 1;
            let sheet_xml = self.serialize_sheet(sheet)?;
            let part_name = format!("xl/worksheets/sheet{}.xml", sheet_num);
            pkg.add_part(&part_name, CT_WORKSHEET, &sheet_xml)?;

            // Write comments if the sheet has any
            if !sheet.comments.is_empty() {
                let comments_xml = self.serialize_comments(sheet)?;
                let comments_part = format!("xl/comments{}.xml", sheet_num);
                pkg.add_part(&comments_part, CT_COMMENTS, &comments_xml)?;

                // Write sheet relationships (for comments)
                let sheet_rels = format!(
                    r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="{}" Target="../comments{}.xml"/>
</Relationships>"#,
                    REL_COMMENTS, sheet_num
                );
                let rels_part = format!("xl/worksheets/_rels/sheet{}.xml.rels", sheet_num);
                pkg.add_part(&rels_part, CT_RELATIONSHIPS, sheet_rels.as_bytes())?;
            }
        }

        // Write shared strings if any
        if !self.shared_strings.is_empty() {
            let ss_xml = self.serialize_shared_strings()?;
            pkg.add_part("xl/sharedStrings.xml", CT_SHARED_STRINGS, &ss_xml)?;
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

    /// Serialize comments to XML.
    ///
    /// ECMA-376 Part 1, Section 18.7 (Comments).
    fn serialize_comments(&self, sheet: &SheetBuilder) -> Result<Vec<u8>> {
        // Collect unique authors and build author->index mapping
        let mut authors: Vec<String> = Vec::new();
        let mut author_index: HashMap<String, u32> = HashMap::new();

        for comment in &sheet.comments {
            let author = comment.author.clone().unwrap_or_default();
            if !author_index.contains_key(&author) {
                author_index.insert(author.clone(), authors.len() as u32);
                authors.push(author);
            }
        }

        // Build comment list
        let comment_list: Vec<Box<types::Comment>> = sheet
            .comments
            .iter()
            .map(|c| {
                let author = c.author.clone().unwrap_or_default();
                let author_id = *author_index.get(&author).unwrap_or(&0);

                Box::new(types::Comment {
                    #[cfg(feature = "sml-comments")]
                    reference: c.reference.clone(),
                    #[cfg(feature = "sml-comments")]
                    author_id,
                    #[cfg(feature = "sml-comments")]
                    guid: None,
                    #[cfg(feature = "sml-comments")]
                    shape_id: None,
                    #[cfg(feature = "sml-comments")]
                    text: Box::new(types::RichString {
                        cell_type: None,
                        reference: vec![Box::new(types::RichTextElement {
                            r_pr: None,
                            cell_type: c.text.clone(),
                            #[cfg(feature = "extra-children")]
                            extra_children: Vec::new(),
                        })],
                        r_ph: Vec::new(),
                        phonetic_pr: None,
                        #[cfg(feature = "extra-children")]
                        extra_children: Vec::new(),
                    }),
                    comment_pr: None,
                    #[cfg(feature = "extra-attrs")]
                    extra_attrs: Default::default(),
                    #[cfg(feature = "extra-children")]
                    extra_children: Vec::new(),
                })
            })
            .collect();

        let comments = types::Comments {
            authors: Box::new(types::Authors {
                author: authors,
                #[cfg(feature = "extra-children")]
                extra_children: Vec::new(),
            }),
            comment_list: Box::new(types::CommentList {
                comment: comment_list,
                #[cfg(feature = "extra-children")]
                extra_children: Vec::new(),
            }),
            extension_list: None,
            #[cfg(feature = "extra-children")]
            extra_children: Vec::new(),
        };

        serialize_with_namespaces(&comments, "comments")
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

    /// Serialize a sheet to XML using generated types.
    fn serialize_sheet(&self, sheet: &SheetBuilder) -> Result<Vec<u8>> {
        // Build column definitions
        #[cfg(feature = "sml-styling")]
        let cols: Vec<Box<types::Columns>> = if sheet.column_widths.is_empty() {
            Vec::new()
        } else {
            vec![Box::new(types::Columns {
                col: sheet
                    .column_widths
                    .iter()
                    .map(|cw| {
                        Box::new(types::Column {
                            #[cfg(feature = "sml-styling")]
                            start_column: cw.min,
                            #[cfg(feature = "sml-styling")]
                            end_column: cw.max,
                            #[cfg(feature = "sml-styling")]
                            width: Some(cw.width),
                            #[cfg(feature = "sml-styling")]
                            style: None,
                            #[cfg(feature = "sml-structure")]
                            hidden: None,
                            #[cfg(feature = "sml-styling")]
                            best_fit: None,
                            #[cfg(feature = "sml-styling")]
                            custom_width: Some(true),
                            #[cfg(feature = "sml-styling")]
                            phonetic: None,
                            #[cfg(feature = "sml-structure")]
                            outline_level: None,
                            #[cfg(feature = "sml-structure")]
                            collapsed: None,
                            #[cfg(feature = "extra-attrs")]
                            extra_attrs: Default::default(),
                        })
                    })
                    .collect(),
                #[cfg(feature = "extra-children")]
                extra_children: Vec::new(),
            })]
        };

        // Group cells by row
        let mut rows_map: HashMap<u32, Vec<(u32, &BuilderCell)>> = HashMap::new();
        for ((row, col), cell) in &sheet.cells {
            rows_map.entry(*row).or_default().push((*col, cell));
        }

        // Build row height lookup
        let row_height_map: HashMap<u32, f64> = sheet
            .row_heights
            .iter()
            .map(|rh| (rh.row, rh.height))
            .collect();

        // Sort and build rows
        let mut row_nums: Vec<_> = rows_map.keys().copied().collect();
        row_nums.sort();

        let rows: Vec<Box<types::Row>> = row_nums
            .iter()
            .map(|&row_num| {
                let cells_data = rows_map.get(&row_num).unwrap();
                let mut sorted_cells: Vec<_> = cells_data.clone();
                sorted_cells.sort_by_key(|(col, _)| *col);

                let cells: Vec<Box<types::Cell>> = sorted_cells
                    .iter()
                    .map(|(col, cell)| {
                        let ref_str = column_to_letter(*col) + &row_num.to_string();
                        self.build_cell(&ref_str, cell)
                    })
                    .collect();

                Box::new(types::Row {
                    reference: Some(row_num),
                    cell_spans: None,
                    style_index: None,
                    #[cfg(feature = "sml-styling")]
                    custom_format: None,
                    #[cfg(feature = "sml-styling")]
                    height: row_height_map.get(&row_num).copied(),
                    #[cfg(feature = "sml-structure")]
                    hidden: None,
                    #[cfg(feature = "sml-styling")]
                    custom_height: row_height_map.get(&row_num).map(|_| true),
                    #[cfg(feature = "sml-structure")]
                    outline_level: None,
                    #[cfg(feature = "sml-structure")]
                    collapsed: None,
                    #[cfg(feature = "sml-styling")]
                    thick_top: None,
                    #[cfg(feature = "sml-styling")]
                    thick_bot: None,
                    #[cfg(feature = "sml-i18n")]
                    placeholder: None,
                    cells,
                    #[cfg(feature = "sml-extensions")]
                    extension_list: None,
                    #[cfg(feature = "extra-attrs")]
                    extra_attrs: Default::default(),
                    #[cfg(feature = "extra-children")]
                    extra_children: Vec::new(),
                })
            })
            .collect();

        // Build merged cells
        let merged_cells = if sheet.merged_cells.is_empty() {
            None
        } else {
            Some(Box::new(types::MergedCells {
                count: Some(sheet.merged_cells.len() as u32),
                merge_cell: sheet
                    .merged_cells
                    .iter()
                    .map(|range| {
                        Box::new(types::MergedCell {
                            reference: range.clone(),
                            #[cfg(feature = "extra-attrs")]
                            extra_attrs: Default::default(),
                        })
                    })
                    .collect(),
                #[cfg(feature = "extra-attrs")]
                extra_attrs: Default::default(),
                #[cfg(feature = "extra-children")]
                extra_children: Vec::new(),
            }))
        };

        // Build worksheet
        // Note: conditionalFormatting and dataValidations require more complex type
        // mapping and are left as manual XML appending for now (TODO: migrate these)
        let worksheet = types::Worksheet {
            #[cfg(feature = "sml-styling")]
            sheet_properties: None,
            dimension: None,
            sheet_views: None,
            #[cfg(feature = "sml-styling")]
            sheet_format: None,
            #[cfg(feature = "sml-styling")]
            cols,
            sheet_data: Box::new(types::SheetData {
                row: rows,
                #[cfg(feature = "extra-children")]
                extra_children: Vec::new(),
            }),
            #[cfg(feature = "sml-formulas")]
            sheet_calc_pr: None,
            #[cfg(feature = "sml-protection")]
            sheet_protection: None,
            #[cfg(feature = "sml-protection")]
            protected_ranges: None,
            #[cfg(feature = "sml-formulas-advanced")]
            scenarios: None,
            #[cfg(feature = "sml-filtering")]
            auto_filter: None,
            #[cfg(feature = "sml-filtering")]
            sort_state: None,
            #[cfg(feature = "sml-formulas-advanced")]
            data_consolidate: None,
            #[cfg(feature = "sml-structure")]
            custom_sheet_views: None,
            merged_cells,
            #[cfg(feature = "sml-i18n")]
            phonetic_pr: None,
            #[cfg(feature = "sml-styling")]
            conditional_formatting: Vec::new(), // TODO: migrate conditional formatting
            #[cfg(feature = "sml-validation")]
            data_validations: None, // TODO: migrate data validations
            #[cfg(feature = "sml-hyperlinks")]
            hyperlinks: None,
            #[cfg(feature = "sml-layout")]
            print_options: None,
            #[cfg(feature = "sml-layout")]
            page_margins: None,
            #[cfg(feature = "sml-layout")]
            page_setup: None,
            #[cfg(feature = "sml-layout")]
            header_footer: None,
            #[cfg(feature = "sml-layout")]
            row_breaks: None,
            #[cfg(feature = "sml-layout")]
            col_breaks: None,
            #[cfg(feature = "sml-metadata")]
            custom_properties: None,
            #[cfg(feature = "sml-formulas-advanced")]
            cell_watches: None,
            #[cfg(feature = "sml-validation")]
            ignored_errors: None,
            #[cfg(feature = "sml-metadata")]
            smart_tags: None,
            #[cfg(feature = "sml-drawings")]
            drawing: None,
            #[cfg(feature = "sml-comments")]
            legacy_drawing: None,
            #[cfg(feature = "sml-layout")]
            legacy_drawing_h_f: None,
            #[cfg(feature = "sml-drawings")]
            drawing_h_f: None,
            #[cfg(feature = "sml-drawings")]
            picture: None,
            #[cfg(feature = "sml-external")]
            ole_objects: None,
            #[cfg(feature = "sml-external")]
            controls: None,
            #[cfg(feature = "sml-external")]
            web_publish_items: None,
            #[cfg(feature = "sml-tables")]
            table_parts: None,
            #[cfg(feature = "sml-extensions")]
            extension_list: None,
            #[cfg(feature = "extra-children")]
            extra_children: Vec::new(),
        };

        serialize_with_namespaces(&worksheet, "worksheet")
    }

    /// Build a Cell type from builder data.
    fn build_cell(&self, reference: &str, cell: &BuilderCell) -> Box<types::Cell> {
        let style_index = self
            .get_cell_style_index(&cell.style)
            .filter(|&s| s > 0)
            .map(|s| s as u32);

        let (cell_type, value, formula) = match &cell.value {
            WriteCellValue::String(s) => {
                let idx = self.string_index.get(s).unwrap_or(&0);
                (
                    Some(types::CellType::SharedString),
                    Some(idx.to_string()),
                    None,
                )
            }
            WriteCellValue::Number(n) => (None, Some(n.to_string()), None),
            WriteCellValue::Boolean(b) => {
                let val = if *b { "1" } else { "0" };
                (Some(types::CellType::Boolean), Some(val.to_string()), None)
            }
            WriteCellValue::Formula(f) => (
                None,
                None,
                Some(Box::new(types::CellFormula {
                    text: Some(f.clone()),
                    cell_type: None,
                    #[cfg(feature = "sml-formulas-advanced")]
                    aca: None,
                    reference: None,
                    #[cfg(feature = "sml-formulas-advanced")]
                    dt2_d: None,
                    #[cfg(feature = "sml-formulas-advanced")]
                    dtr: None,
                    #[cfg(feature = "sml-formulas-advanced")]
                    del1: None,
                    #[cfg(feature = "sml-formulas-advanced")]
                    del2: None,
                    #[cfg(feature = "sml-formulas-advanced")]
                    r1: None,
                    #[cfg(feature = "sml-formulas-advanced")]
                    r2: None,
                    #[cfg(feature = "sml-formulas-advanced")]
                    ca: None,
                    si: None,
                    #[cfg(feature = "sml-formulas-advanced")]
                    bx: None,
                    #[cfg(feature = "extra-attrs")]
                    extra_attrs: Default::default(),
                    #[cfg(feature = "extra-children")]
                    extra_children: Vec::new(),
                })),
            ),
            WriteCellValue::Empty => (None, None, None),
        };

        Box::new(types::Cell {
            reference: Some(reference.to_string()),
            style_index,
            cell_type,
            #[cfg(feature = "sml-metadata")]
            cm: None,
            #[cfg(feature = "sml-metadata")]
            vm: None,
            #[cfg(feature = "sml-i18n")]
            placeholder: None,
            formula,
            value,
            is: None,
            #[cfg(feature = "sml-extensions")]
            extension_list: None,
            #[cfg(feature = "extra-attrs")]
            extra_attrs: Default::default(),
            #[cfg(feature = "extra-children")]
            extra_children: Vec::new(),
        })
    }

    /// Serialize shared strings table to XML using generated types.
    fn serialize_shared_strings(&self) -> Result<Vec<u8>> {
        let count = self.shared_strings.len() as u32;
        let sst = types::SharedStrings {
            count: Some(count),
            unique_count: Some(count),
            si: self
                .shared_strings
                .iter()
                .map(|s| {
                    Box::new(types::RichString {
                        cell_type: Some(s.clone()),
                        reference: Vec::new(),
                        r_ph: Vec::new(),
                        phonetic_pr: None,
                        #[cfg(feature = "extra-children")]
                        extra_children: Vec::new(),
                    })
                })
                .collect(),
            extension_list: None,
            #[cfg(feature = "extra-attrs")]
            extra_attrs: Default::default(),
            #[cfg(feature = "extra-children")]
            extra_children: Vec::new(),
        };
        serialize_with_namespaces(&sst, "sst")
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

// =============================================================================
// ToXml serialization helpers
// =============================================================================

/// Namespace declarations for SML root elements.
const NS_DECLS: &[(&str, &str)] = &[("xmlns", NS_SPREADSHEET)];

/// Serialize a ToXml value with namespace declarations and XML declaration.
fn serialize_with_namespaces(value: &impl ToXml, tag: &str) -> Result<Vec<u8>> {
    use quick_xml::Writer;
    use quick_xml::events::{BytesEnd, BytesStart, Event};

    let inner = Vec::new();
    let mut writer = Writer::new(inner);

    // Write start tag with namespace declarations + type's own attrs
    let start = BytesStart::new(tag);
    let start = value.write_attrs(start);
    let mut start = start;
    for &(key, val) in NS_DECLS {
        start.push_attribute((key, val));
    }

    if value.is_empty_element() {
        writer.write_event(Event::Empty(start))?;
    } else {
        writer.write_event(Event::Start(start))?;
        value.write_children(&mut writer)?;
        writer.write_event(Event::End(BytesEnd::new(tag)))?;
    }

    let inner = writer.into_inner();
    let mut buf = Vec::with_capacity(
        b"<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n".len() + inner.len(),
    );
    buf.extend_from_slice(b"<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n");
    buf.extend_from_slice(&inner);
    Ok(buf)
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
        let read_sheet = workbook.resolved_sheet(0).unwrap();

        assert_eq!(read_sheet.name(), "Sheet1");
        assert_eq!(read_sheet.value_at("A1"), Some("Test Value".to_string()));
        assert_eq!(read_sheet.number_at("B1"), Some(123.45));
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
        let read_sheet = workbook.resolved_sheet(0).unwrap();

        // Check merged cells were preserved
        let merged = read_sheet.merged_cells().expect("Should have merged cells");
        assert_eq!(merged.merge_cell.len(), 2);
        assert_eq!(merged.merge_cell[0].reference.as_str(), "A1:C1");
        assert_eq!(merged.merge_cell[1].reference.as_str(), "A3:B4");

        // Cell values should still be accessible
        assert_eq!(read_sheet.value_at("A1"), Some("Merged Header".to_string()));
    }

    #[test]
    #[cfg(feature = "full")]
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
        let read_sheet = workbook.resolved_sheet(0).unwrap();

        // Check column widths were preserved
        // Structure: Worksheet has Vec<Columns>, each Columns has Vec<Column>
        let columns = read_sheet.columns();
        assert!(!columns.is_empty());

        // Collect all column definitions
        let all_cols: Vec<_> = columns.iter().flat_map(|c| &c.col).collect();
        assert_eq!(all_cols.len(), 2);

        // Column A (col 1)
        assert_eq!(all_cols[0].start_column, 1);
        assert_eq!(all_cols[0].end_column, 1);
        assert_eq!(all_cols[0].width, Some(20.0));

        // Columns B-C (cols 2-3)
        assert_eq!(all_cols[1].start_column, 2);
        assert_eq!(all_cols[1].end_column, 3);
        assert_eq!(all_cols[1].width, Some(15.5));

        // Check row heights were preserved
        let row1 = read_sheet.row(1).unwrap();
        assert_eq!(row1.height, Some(25.0));

        let row2 = read_sheet.row(2).unwrap();
        assert_eq!(row2.height, Some(18.0));
    }

    #[test]
    #[cfg(feature = "full")]
    #[ignore = "blocked on migrating conditional_formatting to generated ToXml types"]
    fn test_roundtrip_conditional_formatting() {
        use std::io::Cursor;

        let mut wb = WorkbookBuilder::new();
        let sheet = wb.add_sheet("Sheet1");
        sheet.set_cell("A1", 10.0);
        sheet.set_cell("A2", 20.0);
        sheet.set_cell("A3", 30.0);

        // Add conditional formatting: highlight cells > 15
        let cf = ConditionalFormat::new("A1:A3")
            .add_cell_is_rule("greaterThan", "15", 1, None)
            .add_expression_rule("$A1>$A2", 2, None);
        sheet.add_conditional_format(cf);

        // Add another rule for duplicates
        let cf2 = ConditionalFormat::new("B1:B10").add_duplicate_values_rule(1, None);
        sheet.add_conditional_format(cf2);

        // Write to memory
        let mut buffer = Cursor::new(Vec::new());
        wb.write(&mut buffer).unwrap();

        // Read back
        buffer.set_position(0);
        let mut workbook = crate::Workbook::from_reader(buffer).unwrap();
        let read_sheet = workbook.resolved_sheet(0).unwrap();

        // Check conditional formatting was preserved
        let cfs = read_sheet.conditional_formatting();
        assert_eq!(cfs.len(), 2);

        // First conditional format has range A1:A3 and 2 rules
        assert_eq!(cfs[0].square_reference.as_deref(), Some("A1:A3"));
        assert_eq!(cfs[0].cf_rule.len(), 2);

        // Second conditional format has range B1:B10 and 1 rule
        assert_eq!(cfs[1].square_reference.as_deref(), Some("B1:B10"));
        assert_eq!(cfs[1].cf_rule.len(), 1);
    }

    #[test]
    #[cfg(feature = "full")]
    #[ignore = "blocked on migrating data_validations to generated ToXml types"]
    fn test_roundtrip_data_validation() {
        use std::io::Cursor;

        let mut wb = WorkbookBuilder::new();
        let sheet = wb.add_sheet("Sheet1");
        sheet.set_cell("A1", 10.0);

        // Add a list validation
        let dv = DataValidationBuilder::list("A1:A10", "\"Yes,No,Maybe\"")
            .with_error("Invalid Input", "Please select from the list")
            .with_prompt("Select", "Choose a value");
        sheet.add_data_validation(dv);

        // Add a whole number validation
        let dv2 = DataValidationBuilder::whole_number(
            "B1:B10",
            crate::DataValidationOperator::GreaterThan,
            "0",
        )
        .with_error("Invalid Number", "Please enter a positive number");
        sheet.add_data_validation(dv2);

        // Write to memory
        let mut buffer = Cursor::new(Vec::new());
        wb.write(&mut buffer).unwrap();

        // Read back
        buffer.set_position(0);
        let mut workbook = crate::Workbook::from_reader(buffer).unwrap();
        let read_sheet = workbook.resolved_sheet(0).unwrap();

        // Check data validations were preserved
        let dvs = read_sheet
            .data_validations()
            .expect("Should have data validations");
        assert_eq!(dvs.data_validation.len(), 2);

        // First validation: list for A1:A10
        let dv0 = &dvs.data_validation[0];
        assert_eq!(dv0.square_reference.as_str(), "A1:A10");
        assert_eq!(dv0.error_title.as_deref(), Some("Invalid Input"));
        assert_eq!(dv0.error.as_deref(), Some("Please select from the list"));
        assert_eq!(dv0.prompt_title.as_deref(), Some("Select"));
        assert_eq!(dv0.prompt.as_deref(), Some("Choose a value"));

        // Second validation: whole number for B1:B10
        let dv1 = &dvs.data_validation[1];
        assert_eq!(dv1.square_reference.as_str(), "B1:B10");
        assert_eq!(dv1.error_title.as_deref(), Some("Invalid Number"));
    }

    #[test]
    fn test_roundtrip_defined_names() {
        use std::io::Cursor;

        let mut wb = WorkbookBuilder::new();
        wb.add_sheet("Sheet1");
        wb.add_sheet("Sheet2");

        // Add a global defined name
        wb.add_defined_name("GlobalRange", "Sheet1!$A$1:$B$10");

        // Add a sheet-scoped defined name
        wb.add_defined_name_with_scope("LocalRange", "Sheet1!$C$1:$D$5", 0);

        // Add a defined name with comment using builder
        let dn = DefinedNameBuilder::new("DataRange", "Sheet2!$A$1:$Z$100")
            .with_comment("Main data table");
        wb.add_defined_name_builder(dn);

        // Add print area
        wb.set_print_area(0, "Sheet1!$A$1:$G$20");

        // Write to memory
        let mut buffer = Cursor::new(Vec::new());
        wb.write(&mut buffer).unwrap();

        // Read back
        buffer.set_position(0);
        let workbook = crate::Workbook::from_reader(buffer).unwrap();

        // Check defined names were preserved
        let names = workbook.defined_names();
        assert_eq!(names.len(), 4);

        // Check global range
        let global = workbook.defined_name("GlobalRange").unwrap();
        assert_eq!(global.name, "GlobalRange");
        assert_eq!(global.reference, "Sheet1!$A$1:$B$10");
        assert!(global.local_sheet_id.is_none());

        // Check sheet-scoped range
        let local = workbook.defined_name_in_sheet("LocalRange", 0).unwrap();
        assert_eq!(local.name, "LocalRange");
        assert_eq!(local.reference, "Sheet1!$C$1:$D$5");
        assert_eq!(local.local_sheet_id, Some(0));

        // Check data range with comment
        let data = workbook.defined_name("DataRange").unwrap();
        assert_eq!(data.name, "DataRange");
        assert_eq!(data.reference, "Sheet2!$A$1:$Z$100");
        assert_eq!(data.comment.as_deref(), Some("Main data table"));

        // Check print area (built-in name)
        let print_area = workbook
            .defined_name_in_sheet("_xlnm.Print_Area", 0)
            .unwrap();
        assert_eq!(print_area.reference, "Sheet1!$A$1:$G$20");
        assert!(print_area.is_builtin());
    }

    #[test]
    fn test_roundtrip_comments() {
        use std::io::Cursor;

        let mut wb = WorkbookBuilder::new();
        let sheet = wb.add_sheet("Sheet1");
        sheet.set_cell("A1", "Hello");
        sheet.set_cell("B1", 42.0);

        // Add comments
        sheet.add_comment("A1", "This is a simple comment");
        sheet.add_comment_with_author("B1", "Review this value", "John Doe");

        // Add a comment using the builder
        let comment = CommentBuilder::new("C1", "Builder comment").author("Jane Smith");
        sheet.add_comment_builder(comment);

        // Write to memory
        let mut buffer = Cursor::new(Vec::new());
        wb.write(&mut buffer).unwrap();

        // Read back
        buffer.set_position(0);
        let mut workbook = crate::Workbook::from_reader(buffer).unwrap();
        let read_sheet = workbook.resolved_sheet(0).unwrap();

        // Check comments were preserved
        let comments = read_sheet.comments();
        assert_eq!(comments.len(), 3);

        // First comment (ext::Comment has public fields)
        let c1 = read_sheet.comment("A1").unwrap();
        assert_eq!(c1.reference, "A1");
        assert_eq!(c1.text, "This is a simple comment");
        assert!(c1.author.is_none() || c1.author.as_ref().is_some_and(|a| a.is_empty()));

        // Second comment
        let c2 = read_sheet.comment("B1").unwrap();
        assert_eq!(c2.reference, "B1");
        assert_eq!(c2.text, "Review this value");
        assert_eq!(c2.author.as_deref(), Some("John Doe"));

        // Third comment
        let c3 = read_sheet.comment("C1").unwrap();
        assert_eq!(c3.reference, "C1");
        assert_eq!(c3.text, "Builder comment");
        assert_eq!(c3.author.as_deref(), Some("Jane Smith"));

        // Check helper method
        assert!(read_sheet.has_comment("A1"));
        assert!(read_sheet.has_comment("B1"));
        assert!(!read_sheet.has_comment("D1"));
    }
}
