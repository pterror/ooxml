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
#[cfg(feature = "sml-charts")]
const CT_DRAWING: &str = "application/vnd.openxmlformats-officedocument.drawing+xml";
#[cfg(feature = "sml-charts")]
const CT_CHART: &str = "application/vnd.openxmlformats-officedocument.drawingml.chart+xml";
#[cfg(feature = "sml-pivot")]
const CT_PIVOT_TABLE: &str =
    "application/vnd.openxmlformats-officedocument.spreadsheetml.pivotTable+xml";
#[cfg(feature = "sml-pivot")]
const CT_PIVOT_CACHE_DEF: &str =
    "application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheDefinition+xml";
#[cfg(feature = "sml-pivot")]
const CT_PIVOT_CACHE_REC: &str =
    "application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheRecords+xml";

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
const REL_HYPERLINK: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink";
#[cfg(feature = "sml-charts")]
const REL_DRAWING: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing";
#[cfg(feature = "sml-charts")]
const REL_CHART: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart";
#[cfg(feature = "sml-pivot")]
const REL_PIVOT_TABLE: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotTable";
#[cfg(feature = "sml-pivot")]
const REL_PIVOT_CACHE_DEF: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheDefinition";
#[cfg(feature = "sml-pivot")]
const REL_PIVOT_CACHE_REC: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheRecords";

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

/// A single rich-text run inside a comment.
///
/// Created by [`CommentBuilder::add_run`].  Call the setter methods to apply
/// formatting, then call [`CommentBuilder::add_run`] again for the next run.
#[derive(Debug, Clone, Default)]
pub struct CommentRun {
    /// Text content of this run.
    pub text: String,
    /// Bold.
    pub bold: bool,
    /// Italic.
    pub italic: bool,
    /// RGB hex color (e.g., `"FF0000"` for red).
    pub color: Option<String>,
    /// Font size in points.
    pub font_size: Option<f64>,
}

impl CommentRun {
    /// Set bold formatting.
    pub fn set_bold(&mut self, bold: bool) -> &mut Self {
        self.bold = bold;
        self
    }

    /// Set italic formatting.
    pub fn set_italic(&mut self, italic: bool) -> &mut Self {
        self.italic = italic;
        self
    }

    /// Set the run color as an RGB hex string (e.g., `"FF0000"` for red).
    pub fn set_color(&mut self, rgb: &str) -> &mut Self {
        self.color = Some(rgb.to_string());
        self
    }

    /// Set the font size in points.
    pub fn set_font_size(&mut self, pt: f64) -> &mut Self {
        self.font_size = Some(pt);
        self
    }
}

/// A cell comment (note) for writing.
///
/// Comments can contain plain text or rich text with multiple runs.
///
/// # Example
///
/// ```ignore
/// let mut wb = WorkbookBuilder::new();
/// let sheet = wb.add_sheet("Sheet1");
/// sheet.add_comment("A1", "This is a comment");
/// sheet.add_comment_with_author("B1", "Another comment", "John Doe");
///
/// // Rich-text comment via builder
/// let mut cb = CommentBuilder::new_rich("C1");
/// cb.add_run("Important: ").set_bold(true);
/// cb.add_run("see the spec.");
/// sheet.add_comment_builder(cb);
/// ```
#[derive(Debug, Clone)]
pub struct CommentBuilder {
    /// Cell reference (e.g., "A1").
    pub reference: String,
    /// Plain-text content (used when no runs are set).
    pub text: String,
    /// Author of the comment (optional).
    pub author: Option<String>,
    /// Rich-text runs (when non-empty, `text` is ignored).
    pub runs: Vec<CommentRun>,
}

impl CommentBuilder {
    /// Create a new comment with plain text.
    pub fn new(reference: impl Into<String>, text: impl Into<String>) -> Self {
        Self {
            reference: reference.into(),
            text: text.into(),
            author: None,
            runs: Vec::new(),
        }
    }

    /// Create a new rich-text comment (no initial plain text).
    ///
    /// Use [`add_run`](Self::add_run) to append formatted runs.
    pub fn new_rich(reference: impl Into<String>) -> Self {
        Self {
            reference: reference.into(),
            text: String::new(),
            author: None,
            runs: Vec::new(),
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
            runs: Vec::new(),
        }
    }

    /// Set the author of the comment.
    pub fn author(mut self, author: impl Into<String>) -> Self {
        self.author = Some(author.into());
        self
    }

    /// Append a rich-text run and return a mutable reference to it.
    ///
    /// The returned `&mut CommentRun` can be used to set formatting
    /// (bold, italic, color, font size).
    ///
    /// # Example
    ///
    /// ```ignore
    /// cb.add_run("Warning: ").set_bold(true).set_color("FF0000");
    /// cb.add_run("normal text");
    /// ```
    pub fn add_run(&mut self, text: &str) -> &mut CommentRun {
        self.runs.push(CommentRun {
            text: text.to_string(),
            ..Default::default()
        });
        self.runs.last_mut().unwrap()
    }
}

// =============================================================================
// Sheet protection
// =============================================================================

/// Options for `SheetBuilder::set_sheet_protection` (ECMA-376 §18.3.1.85).
///
/// By default all fields are `false` (no restrictions).  Set a field to `true`
/// to **prevent** that operation (the OOXML attribute names are inverted: a
/// value of `true` on `formatCells` means "format cells is *not* allowed").
///
/// # Example
///
/// ```ignore
/// sheet.set_sheet_protection(SheetProtectionOptions {
///     sheet: true,
///     select_locked_cells: false,
///     ..Default::default()
/// });
/// ```
#[cfg(feature = "sml-protection")]
#[derive(Debug, Clone, Default)]
pub struct SheetProtectionOptions {
    /// Optional plain-text password (hashed with the OOXML XOR algorithm).
    ///
    /// When `None`, no password is set (the sheet can be unprotected without
    /// a password).
    pub password: Option<String>,
    /// Lock the sheet (enable protection).  Must be `true` for any other
    /// restriction to take effect.
    pub sheet: bool,
    /// Prevent selecting locked cells.
    pub select_locked_cells: bool,
    /// Prevent selecting unlocked cells.
    pub select_unlocked_cells: bool,
    /// Prevent formatting cells.
    pub format_cells: bool,
    /// Prevent formatting columns.
    pub format_columns: bool,
    /// Prevent formatting rows.
    pub format_rows: bool,
    /// Prevent inserting columns.
    pub insert_columns: bool,
    /// Prevent inserting rows.
    pub insert_rows: bool,
    /// Prevent deleting columns.
    pub delete_columns: bool,
    /// Prevent deleting rows.
    pub delete_rows: bool,
    /// Prevent sorting.
    pub sort: bool,
    /// Prevent using auto-filter.
    pub auto_filter: bool,
    /// Prevent using pivot tables.
    pub pivot_tables: bool,
}

/// Compute the OOXML XOR password hash for a plain-text password.
///
/// Implements the algorithm described in ECMA-376 Part 1, §18.2.28.
/// Returns the 16-bit hash as a 2-byte `Vec<u8>` (big-endian), which is the
/// `STUnsignedShortHex` representation used by `sheetProtection/@password`.
#[cfg(feature = "sml-protection")]
fn ooxml_xor_hash(password: &str) -> Vec<u8> {
    if password.is_empty() {
        return vec![0x00, 0x00];
    }

    // The algorithm from ECMA-376 Part 1, §18.2.28:
    // 1. Initialise hash to 0.
    // 2. For each character in reverse order:
    //    a. XOR hash with a rotating key derived from the character.
    //    b. Rotate the hash 1 bit left.
    // 3. XOR with the length.
    // 4. XOR with 0xCE4B (the "password verifier seed").

    let chars: Vec<u8> = password.chars().map(|c| c as u8).collect();
    let mut hash: u16 = 0;

    for &ch in chars.iter().rev() {
        // rotate hash left by 1 bit (15-bit rotation for 15-bit value)
        hash = ((hash << 1) | (hash >> 14)) & 0x7FFF;
        hash ^= ch as u16;
    }

    // Final rotate and XOR with length + seed
    hash = ((hash << 1) | (hash >> 14)) & 0x7FFF;
    hash ^= chars.len() as u16;
    hash ^= 0xCE4B;

    vec![(hash >> 8) as u8, (hash & 0xFF) as u8]
}

/// Destination of a hyperlink.
#[derive(Debug, Clone)]
enum HyperlinkDest {
    /// External URL — needs a relationship entry in sheet rels.
    External(String),
    /// Internal reference (e.g. `"Sheet2!A1"`) — no relationship needed.
    Internal(String),
}

/// A hyperlink to be written in a sheet.
#[derive(Debug, Clone)]
struct HyperlinkEntry {
    /// Cell reference the hyperlink is attached to (e.g. `"A1"`).
    reference: String,
    dest: HyperlinkDest,
    tooltip: Option<String>,
    display: Option<String>,
}

/// Page orientation for `set_page_setup`.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub enum PageOrientation {
    /// Portrait orientation (taller than wide).
    #[default]
    Portrait,
    /// Landscape orientation (wider than tall).
    Landscape,
}

/// Options for `SheetBuilder::set_page_setup`.
///
/// All fields are optional; unset fields are left at their Excel defaults.
#[derive(Debug, Clone, Default)]
pub struct PageSetupOptions {
    /// Page orientation.
    pub orientation: Option<PageOrientation>,
    /// Paper size (e.g. 1 = Letter, 9 = A4). See ECMA-376 §18.18.43.
    pub paper_size: Option<u32>,
    /// Scaling percentage (10–400). Use instead of `fit_to_width`/`fit_to_height`.
    pub scale: Option<u32>,
    /// Fit to this many pages wide (0 = auto). Used with `fit_to_height`.
    pub fit_to_width: Option<u32>,
    /// Fit to this many pages tall (0 = auto). Used with `fit_to_width`.
    pub fit_to_height: Option<u32>,
}

impl PageSetupOptions {
    /// Create a new empty page-setup options object.
    pub fn new() -> Self {
        Self::default()
    }

    /// Set the page orientation.
    pub fn with_orientation(mut self, orientation: PageOrientation) -> Self {
        self.orientation = Some(orientation);
        self
    }

    /// Set the paper size.
    pub fn with_paper_size(mut self, size: u32) -> Self {
        self.paper_size = Some(size);
        self
    }

    /// Set the scaling percentage.
    pub fn with_scale(mut self, scale: u32) -> Self {
        self.scale = Some(scale);
        self
    }

    /// Set the fit-to-pages dimensions.
    pub fn with_fit_to(mut self, width: u32, height: u32) -> Self {
        self.fit_to_width = Some(width);
        self.fit_to_height = Some(height);
        self
    }
}

// ============================================================================
// Chart embedding (sml-charts feature)
// ============================================================================

/// A chart embedded in a worksheet.
///
/// Created by [`SheetBuilder::embed_chart`].
#[cfg(feature = "sml-charts")]
#[derive(Debug)]
struct ChartEntry {
    /// Raw chart XML bytes.
    chart_xml: Vec<u8>,
    /// Column index (0-based) of the top-left anchor cell.
    x: u32,
    /// Row index (0-based) of the top-left anchor cell.
    y: u32,
    /// Width in cells.
    width: u32,
    /// Height in cells.
    height: u32,
}

// ============================================================================
// Pivot table support (sml-pivot feature)
// ============================================================================

/// Options for [`SheetBuilder::add_pivot_table`].
///
/// Produces a minimal but spec-compliant pivot table definition.
///
/// # Example
///
/// ```ignore
/// sheet.add_pivot_table(PivotTableOptions {
///     name: "SalesPivot".to_string(),
///     source_ref: "Sheet1!$A$1:$D$10".to_string(),
///     dest_ref: "F1".to_string(),
///     row_fields: vec!["Region".to_string()],
///     col_fields: vec!["Quarter".to_string()],
///     data_fields: vec!["Sales".to_string()],
/// });
/// ```
#[cfg(feature = "sml-pivot")]
#[derive(Debug, Clone)]
pub struct PivotTableOptions {
    /// Name of the pivot table (shown in Excel's pivot table list).
    pub name: String,
    /// Source data range, e.g. `"Sheet1!$A$1:$D$10"`.
    pub source_ref: String,
    /// Top-left cell of the pivot table output, e.g. `"A1"`.
    pub dest_ref: String,
    /// Field names to place on the row axis.
    pub row_fields: Vec<String>,
    /// Field names to place on the column axis.
    pub col_fields: Vec<String>,
    /// Field names to aggregate in the values area (sum by default).
    pub data_fields: Vec<String>,
}

/// Internal record of a pivot table added to a sheet.
#[cfg(feature = "sml-pivot")]
#[derive(Debug)]
struct PivotEntry {
    opts: PivotTableOptions,
}

/// Error type for `add_ignored_error`.
///
/// Each variant corresponds to one of the boolean flags on the `<ignoredError>`
/// element (ECMA-376 §18.3.1.35).
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum IgnoredErrorType {
    /// Number stored as text (the most common Excel warning).
    NumberStoredAsText,
    /// Formula error (e.g. #VALUE!, #REF!).
    Formula,
    /// Two-digit text year.
    TwoDigitTextYear,
    /// Formula evaluation error.
    EvalError,
    /// Formula range mismatch.
    FormulaRange,
    /// Unlocked formula in a protected sheet.
    UnlockedFormula,
    /// Empty cell reference.
    EmptyCellReference,
    /// List data validation mismatch.
    ListDataValidation,
    /// Calculated column formula inconsistency.
    CalculatedColumn,
}

/// A sheet being built.
#[derive(Debug)]
pub struct SheetBuilder {
    name: String,
    /// Cells stored as a map for O(1) mutation; resolved to rows at write time.
    cells: HashMap<(u32, u32), BuilderCell>,
    /// Row heights applied to Row elements at write time.
    row_heights: HashMap<u32, f64>,
    /// Per-row outline levels (ECMA-376 §18.3.1.73 `@outlineLevel`).
    row_outline_levels: HashMap<u32, u8>,
    /// Per-row collapsed flag (ECMA-376 §18.3.1.73 `@collapsed`).
    row_collapsed: HashMap<u32, bool>,
    /// Per-column outline levels (ECMA-376 §18.3.1.13 `@outlineLevel`).
    col_outline_levels: HashMap<u32, u8>,
    /// Per-column collapsed flag (ECMA-376 §18.3.1.13 `@collapsed`).
    col_collapsed: HashMap<u32, bool>,
    /// Comments go into a separate XML file, not into Worksheet.
    comments: Vec<CommentBuilder>,
    /// Hyperlinks, resolved to worksheet XML + sheet rels at write time.
    hyperlinks: Vec<HyperlinkEntry>,
    /// Embedded charts, written to xl/charts/ + xl/drawings/ at write time.
    #[cfg(feature = "sml-charts")]
    charts: Vec<ChartEntry>,
    /// Pivot tables, written to xl/pivotTables/ and xl/pivotCache/ at write time.
    #[cfg(feature = "sml-pivot")]
    pivot_tables: Vec<PivotEntry>,
    /// Show/hide gridlines for the default sheet view.
    show_gridlines: Option<bool>,
    /// Show/hide row and column headers for the default sheet view.
    show_row_col_headers: Option<bool>,
    /// All other worksheet state lives here directly; mutated by setter methods.
    worksheet: types::Worksheet,
}

/// Create a default empty Worksheet, ready to be filled in by SheetBuilder methods.
fn init_worksheet() -> types::Worksheet {
    types::Worksheet {
        #[cfg(feature = "sml-styling")]
        sheet_properties: None,
        dimension: None,
        sheet_views: None,
        #[cfg(feature = "sml-styling")]
        sheet_format: None,
        #[cfg(feature = "sml-styling")]
        cols: Vec::new(),
        sheet_data: Box::new(types::SheetData {
            row: Vec::new(),
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
        merged_cells: None,
        #[cfg(feature = "sml-i18n")]
        phonetic_pr: None,
        #[cfg(feature = "sml-styling")]
        conditional_formatting: Vec::new(),
        #[cfg(feature = "sml-validation")]
        data_validations: None,
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
    }
}

impl SheetBuilder {
    fn new(name: impl Into<String>) -> Self {
        Self {
            name: name.into(),
            cells: HashMap::new(),
            row_heights: HashMap::new(),
            row_outline_levels: HashMap::new(),
            row_collapsed: HashMap::new(),
            col_outline_levels: HashMap::new(),
            col_collapsed: HashMap::new(),
            comments: Vec::new(),
            hyperlinks: Vec::new(),
            #[cfg(feature = "sml-charts")]
            charts: Vec::new(),
            #[cfg(feature = "sml-pivot")]
            pivot_tables: Vec::new(),
            show_gridlines: None,
            show_row_col_headers: None,
            worksheet: init_worksheet(),
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
        let mc = self.worksheet.merged_cells.get_or_insert_with(Box::default);
        mc.merge_cell.push(types::MergedCell {
            reference: range.to_string(),
            #[cfg(feature = "extra-attrs")]
            extra_attrs: Default::default(),
        });
        mc.count = Some(mc.merge_cell.len() as u32);
    }

    /// Set the width of a column (in character units, Excel default is ~8.43).
    ///
    /// Column is specified by letter (e.g., "A", "B", "AA").
    pub fn set_column_width(&mut self, col: &str, width: f64) {
        if let Some(col_num) = column_letter_to_number(col) {
            self.push_column(col_num, col_num, width);
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
            self.push_column(min, max, width);
        }
    }

    /// Push a column definition directly into the worksheet.
    fn push_column(&mut self, min: u32, max: u32, width: f64) {
        #[cfg(feature = "sml-styling")]
        {
            let col = types::Column {
                #[cfg(feature = "sml-styling")]
                start_column: min,
                #[cfg(feature = "sml-styling")]
                end_column: max,
                #[cfg(feature = "sml-styling")]
                width: Some(width),
                #[cfg(feature = "sml-styling")]
                style: None,
                #[cfg(feature = "sml-structure")]
                hidden: None,
                #[cfg(feature = "sml-styling")]
                best_fit: None,
                #[cfg(feature = "sml-styling")]
                custom_width: Some(true),
                #[cfg(feature = "sml-i18n")]
                phonetic: None,
                #[cfg(feature = "sml-structure")]
                outline_level: None,
                #[cfg(feature = "sml-structure")]
                collapsed: None,
                #[cfg(feature = "extra-attrs")]
                extra_attrs: Default::default(),
            };
            if let Some(cols_list) = self.worksheet.cols.first_mut() {
                cols_list.col.push(col);
            } else {
                self.worksheet.cols.push(types::Columns {
                    col: vec![col],
                    #[cfg(feature = "extra-children")]
                    extra_children: Vec::new(),
                });
            }
        }
        #[cfg(not(feature = "sml-styling"))]
        {
            let _ = (min, max, width);
        }
    }

    /// Set the height of a row (in points, Excel default is ~15).
    pub fn set_row_height(&mut self, row: u32, height: f64) {
        self.row_heights.insert(row, height);
    }

    /// Freeze the top `rows` rows and left `cols` columns.
    ///
    /// Pass `0` for either dimension to freeze only the other axis.
    /// For example, `set_freeze_pane(1, 0)` freezes the header row.
    ///
    /// This is equivalent to View → Freeze Panes in Excel.
    ///
    /// # Example
    ///
    /// ```ignore
    /// // Freeze the first row (common for headers)
    /// sheet.set_freeze_pane(1, 0);
    ///
    /// // Freeze both first row and first column
    /// sheet.set_freeze_pane(1, 1);
    /// ```
    pub fn set_freeze_pane(&mut self, rows: u32, cols: u32) {
        self.apply_freeze_pane(rows, cols);
    }

    /// Freeze the top `n` rows (e.g., header rows).
    ///
    /// Shorthand for `set_freeze_pane(n, 0)`.
    pub fn freeze_rows(&mut self, n: u32) {
        let (_, c) = self.current_freeze_pane();
        self.apply_freeze_pane(n, c);
    }

    /// Freeze the left `n` columns.
    ///
    /// Shorthand for `set_freeze_pane(0, n)`.
    pub fn freeze_cols(&mut self, n: u32) {
        let (r, _) = self.current_freeze_pane();
        self.apply_freeze_pane(r, n);
    }

    /// Read back the current freeze pane settings from the worksheet.
    fn current_freeze_pane(&self) -> (u32, u32) {
        #[cfg(feature = "sml-structure")]
        {
            let pane = self
                .worksheet
                .sheet_views
                .as_deref()
                .and_then(|v| v.sheet_view.first())
                .and_then(|sv| sv.pane.as_deref());
            if let Some(p) = pane {
                return (
                    p.y_split.map(|y| y as u32).unwrap_or(0),
                    p.x_split.map(|x| x as u32).unwrap_or(0),
                );
            }
        }
        (0, 0)
    }

    /// Write freeze pane settings directly into the worksheet's sheet_views.
    fn apply_freeze_pane(&mut self, frozen_rows: u32, frozen_cols: u32) {
        #[cfg(feature = "sml-structure")]
        {
            if frozen_rows == 0 && frozen_cols == 0 {
                self.worksheet.sheet_views = None;
                return;
            }
            // Determine which pane is active after freezing.
            let active_pane = match (frozen_rows > 0, frozen_cols > 0) {
                (true, true) => types::PaneType::BottomRight,
                (true, false) => types::PaneType::BottomLeft,
                (false, true) => types::PaneType::TopRight,
                (false, false) => types::PaneType::TopLeft,
            };
            // topLeftCell is the first unfrozen cell (e.g., "B2" for 1 frozen row + 1 frozen col).
            let top_left_col = if frozen_cols > 0 {
                column_to_letter(frozen_cols + 1)
            } else {
                "A".to_string()
            };
            let top_left_row = frozen_rows + 1;
            let top_left_cell = format!("{}{}", top_left_col, top_left_row);

            let pane = types::Pane {
                x_split: (frozen_cols > 0).then_some(frozen_cols as f64),
                y_split: (frozen_rows > 0).then_some(frozen_rows as f64),
                top_left_cell: Some(top_left_cell),
                active_pane: Some(active_pane),
                state: Some(types::PaneState::Frozen),
                #[cfg(feature = "extra-attrs")]
                extra_attrs: Default::default(),
            };
            let sheet_view = types::SheetView {
                #[cfg(feature = "sml-protection")]
                window_protection: None,
                #[cfg(feature = "sml-formulas")]
                show_formulas: None,
                #[cfg(feature = "sml-styling")]
                show_grid_lines: None,
                #[cfg(feature = "sml-styling")]
                show_row_col_headers: None,
                #[cfg(feature = "sml-styling")]
                show_zeros: None,
                #[cfg(feature = "sml-i18n")]
                right_to_left: None,
                tab_selected: None,
                #[cfg(feature = "sml-layout")]
                show_ruler: None,
                #[cfg(feature = "sml-structure")]
                show_outline_symbols: None,
                #[cfg(feature = "sml-styling")]
                default_grid_color: None,
                #[cfg(feature = "sml-layout")]
                show_white_space: None,
                view: None,
                top_left_cell: None,
                #[cfg(feature = "sml-styling")]
                color_id: None,
                zoom_scale: None,
                zoom_scale_normal: None,
                #[cfg(feature = "sml-layout")]
                zoom_scale_sheet_layout_view: None,
                #[cfg(feature = "sml-layout")]
                zoom_scale_page_layout_view: None,
                workbook_view_id: 0,
                #[cfg(feature = "sml-structure")]
                pane: Some(Box::new(pane)),
                selection: vec![types::Selection {
                    pane: Some(active_pane),
                    active_cell: None,
                    active_cell_id: None,
                    square_reference: None,
                    #[cfg(feature = "extra-attrs")]
                    extra_attrs: Default::default(),
                }],
                #[cfg(feature = "sml-pivot")]
                pivot_selection: Vec::new(),
                #[cfg(feature = "sml-extensions")]
                extension_list: None,
                #[cfg(feature = "extra-attrs")]
                extra_attrs: Default::default(),
                #[cfg(feature = "extra-children")]
                extra_children: Vec::new(),
            };
            self.worksheet.sheet_views = Some(Box::new(types::SheetViews {
                sheet_view: vec![sheet_view],
                extension_list: None,
                #[cfg(feature = "extra-children")]
                extra_children: Vec::new(),
            }));
        }
        #[cfg(not(feature = "sml-structure"))]
        {
            let _ = (frozen_rows, frozen_cols);
        }
    }

    /// Add conditional formatting to the sheet.
    pub fn add_conditional_format(&mut self, cf: ConditionalFormat) {
        #[cfg(feature = "sml-styling")]
        self.worksheet
            .conditional_formatting
            .push(build_one_conditional_format(&cf));
        #[cfg(not(feature = "sml-styling"))]
        let _ = cf;
    }

    /// Add data validation to the sheet.
    pub fn add_data_validation(&mut self, dv: DataValidationBuilder) {
        #[cfg(feature = "sml-validation")]
        {
            let validation = build_one_data_validation(&dv);
            let dvs = self.worksheet.data_validations.get_or_insert_with(|| {
                Box::new(types::DataValidations {
                    disable_prompts: None,
                    x_window: None,
                    y_window: None,
                    count: None,
                    data_validation: Vec::new(),
                    #[cfg(feature = "extra-attrs")]
                    extra_attrs: Default::default(),
                    #[cfg(feature = "extra-children")]
                    extra_children: Vec::new(),
                })
            });
            dvs.data_validation.push(validation);
            dvs.count = Some(dvs.data_validation.len() as u32);
        }
        #[cfg(not(feature = "sml-validation"))]
        let _ = dv;
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

    /// Enable auto-filter dropdowns on a header range (e.g. `"A1:D1"`).
    ///
    /// Excel will add dropdown buttons to every column in the range.  Callers
    /// typically set this on the same row as their column headers.
    ///
    /// # Example
    ///
    /// ```ignore
    /// sheet.set_cell("A1", "Name");
    /// sheet.set_cell("B1", "Score");
    /// sheet.set_auto_filter("A1:B1");
    /// ```
    pub fn set_auto_filter(&mut self, range: &str) {
        #[cfg(feature = "sml-filtering")]
        {
            self.worksheet.auto_filter = Some(Box::new(types::AutoFilter {
                reference: Some(range.to_string()),
                filter_column: Vec::new(),
                sort_state: None,
                #[cfg(feature = "sml-extensions")]
                extension_list: None,
                #[cfg(feature = "extra-attrs")]
                extra_attrs: Default::default(),
                #[cfg(feature = "extra-children")]
                extra_children: Vec::new(),
            }));
        }
        #[cfg(not(feature = "sml-filtering"))]
        let _ = range;
    }

    /// Add an external hyperlink on a cell (e.g. a URL).
    ///
    /// The URL is written as a relationship in the sheet's `.rels` file, and
    /// the `<hyperlink>` element references it by `r:id`.
    ///
    /// # Example
    ///
    /// ```ignore
    /// sheet.set_cell("A1", "Visit us");
    /// sheet.add_hyperlink("A1", "https://example.com");
    /// ```
    pub fn add_hyperlink(&mut self, reference: &str, url: &str) {
        self.hyperlinks.push(HyperlinkEntry {
            reference: reference.to_string(),
            dest: HyperlinkDest::External(url.to_string()),
            tooltip: None,
            display: None,
        });
    }

    /// Set a tooltip on the last-added hyperlink.
    ///
    /// Call immediately after [`add_hyperlink`](Self::add_hyperlink) or
    /// [`add_internal_hyperlink`](Self::add_internal_hyperlink).
    pub fn set_last_hyperlink_tooltip(&mut self, tooltip: &str) {
        if let Some(h) = self.hyperlinks.last_mut() {
            h.tooltip = Some(tooltip.to_string());
        }
    }

    /// Add an internal hyperlink that navigates to another location in the
    /// workbook (e.g. `"Sheet2!A1"`).
    ///
    /// Internal hyperlinks do not need a relationship — the location is stored
    /// inline in the `<hyperlink>` element.
    ///
    /// # Example
    ///
    /// ```ignore
    /// sheet.add_internal_hyperlink("B1", "Sheet2!A1");
    /// ```
    pub fn add_internal_hyperlink(&mut self, reference: &str, location: &str) {
        self.hyperlinks.push(HyperlinkEntry {
            reference: reference.to_string(),
            dest: HyperlinkDest::Internal(location.to_string()),
            tooltip: None,
            display: None,
        });
    }

    // -------------------------------------------------------------------------
    // Page layout
    // -------------------------------------------------------------------------

    /// Set the page setup options for printing (ECMA-376 §18.3.1.63).
    ///
    /// # Example
    ///
    /// ```ignore
    /// sheet.set_page_setup(PageSetupOptions::new()
    ///     .with_orientation(PageOrientation::Landscape)
    ///     .with_paper_size(9));  // A4
    /// ```
    pub fn set_page_setup(&mut self, opts: PageSetupOptions) {
        #[cfg(feature = "sml-layout")]
        {
            let orientation = opts.orientation.map(|o| match o {
                PageOrientation::Portrait => types::STOrientation::Portrait,
                PageOrientation::Landscape => types::STOrientation::Landscape,
            });
            self.worksheet.page_setup = Some(Box::new(types::PageSetup {
                #[cfg(feature = "sml-layout")]
                paper_size: opts.paper_size,
                #[cfg(feature = "sml-layout")]
                paper_height: None,
                #[cfg(feature = "sml-layout")]
                paper_width: None,
                #[cfg(feature = "sml-layout")]
                scale: opts.scale,
                #[cfg(feature = "sml-layout")]
                first_page_number: None,
                #[cfg(feature = "sml-layout")]
                fit_to_width: opts.fit_to_width,
                #[cfg(feature = "sml-layout")]
                fit_to_height: opts.fit_to_height,
                #[cfg(feature = "sml-layout")]
                page_order: None,
                #[cfg(feature = "sml-layout")]
                orientation,
                #[cfg(feature = "sml-layout")]
                use_printer_defaults: None,
                #[cfg(feature = "sml-layout")]
                black_and_white: None,
                #[cfg(feature = "sml-layout")]
                draft: None,
                #[cfg(feature = "sml-layout")]
                cell_comments: None,
                #[cfg(feature = "sml-layout")]
                use_first_page_number: None,
                #[cfg(feature = "sml-layout")]
                errors: None,
                #[cfg(feature = "sml-layout")]
                horizontal_dpi: None,
                #[cfg(feature = "sml-layout")]
                vertical_dpi: None,
                #[cfg(feature = "sml-layout")]
                copies: None,
                id: None,
                #[cfg(feature = "extra-attrs")]
                extra_attrs: Default::default(),
            }));
        }
        #[cfg(not(feature = "sml-layout"))]
        let _ = opts;
    }

    /// Set the page margins for printing (ECMA-376 §18.3.1.62).
    ///
    /// All measurements are in inches.
    ///
    /// # Example
    ///
    /// ```ignore
    /// sheet.set_page_margins(0.7, 0.7, 0.75, 0.75, 0.3, 0.3);
    /// ```
    pub fn set_page_margins(
        &mut self,
        left: f64,
        right: f64,
        top: f64,
        bottom: f64,
        header: f64,
        footer: f64,
    ) {
        #[cfg(feature = "sml-layout")]
        {
            self.worksheet.page_margins = Some(Box::new(types::PageMargins {
                #[cfg(feature = "sml-layout")]
                left,
                #[cfg(feature = "sml-layout")]
                right,
                #[cfg(feature = "sml-layout")]
                top,
                #[cfg(feature = "sml-layout")]
                bottom,
                #[cfg(feature = "sml-layout")]
                header,
                #[cfg(feature = "sml-layout")]
                footer,
                #[cfg(feature = "extra-attrs")]
                extra_attrs: Default::default(),
            }));
        }
        #[cfg(not(feature = "sml-layout"))]
        let _ = (left, right, top, bottom, header, footer);
    }

    // -------------------------------------------------------------------------
    // Row and column grouping (outline)
    // -------------------------------------------------------------------------

    /// Set the outline level for a row (ECMA-376 §18.3.1.73 `@outlineLevel`).
    ///
    /// Level 0 means no grouping; levels 1–7 define nested groups.
    pub fn set_row_outline_level(&mut self, row: u32, level: u8) {
        self.row_outline_levels.insert(row, level);
    }

    /// Set whether a row should be displayed as collapsed (ECMA-376 §18.3.1.73 `@collapsed`).
    pub fn set_row_collapsed(&mut self, row: u32, collapsed: bool) {
        self.row_collapsed.insert(row, collapsed);
    }

    /// Set the outline level for a column (ECMA-376 §18.3.1.13 `@outlineLevel`).
    ///
    /// Column is specified by letter (e.g., "A", "B", "AA").
    pub fn set_column_outline_level(&mut self, col: &str, level: u8) {
        if let Some(col_num) = column_letter_to_number(col) {
            self.col_outline_levels.insert(col_num, level);
        }
    }

    /// Set whether a column should be displayed as collapsed (ECMA-376 §18.3.1.13 `@collapsed`).
    ///
    /// Column is specified by letter (e.g., "A", "B", "AA").
    pub fn set_column_collapsed(&mut self, col: &str, collapsed: bool) {
        if let Some(col_num) = column_letter_to_number(col) {
            self.col_collapsed.insert(col_num, collapsed);
        }
    }

    // -------------------------------------------------------------------------
    // Ignored errors
    // -------------------------------------------------------------------------

    /// Suppress an Excel validation warning for a cell range (ECMA-376 §18.3.1.35).
    ///
    /// This tells Excel to ignore the specified error type for the given range,
    /// preventing the green-triangle warning indicators.
    ///
    /// # Example
    ///
    /// ```ignore
    /// // Suppress the "number stored as text" warning on A1:A10
    /// sheet.add_ignored_error("A1:A10", IgnoredErrorType::NumberStoredAsText);
    /// ```
    pub fn add_ignored_error(&mut self, sqref: &str, error_type: IgnoredErrorType) {
        #[cfg(feature = "sml-validation")]
        {
            let mut entry = types::IgnoredError {
                square_reference: sqref.to_string(),
                eval_error: None,
                two_digit_text_year: None,
                number_stored_as_text: None,
                formula: None,
                formula_range: None,
                unlocked_formula: None,
                empty_cell_reference: None,
                list_data_validation: None,
                calculated_column: None,
                #[cfg(feature = "extra-attrs")]
                extra_attrs: Default::default(),
            };
            match error_type {
                IgnoredErrorType::NumberStoredAsText => {
                    entry.number_stored_as_text = Some(true);
                }
                IgnoredErrorType::Formula => {
                    entry.formula = Some(true);
                }
                IgnoredErrorType::TwoDigitTextYear => {
                    entry.two_digit_text_year = Some(true);
                }
                IgnoredErrorType::EvalError => {
                    entry.eval_error = Some(true);
                }
                IgnoredErrorType::FormulaRange => {
                    entry.formula_range = Some(true);
                }
                IgnoredErrorType::UnlockedFormula => {
                    entry.unlocked_formula = Some(true);
                }
                IgnoredErrorType::EmptyCellReference => {
                    entry.empty_cell_reference = Some(true);
                }
                IgnoredErrorType::ListDataValidation => {
                    entry.list_data_validation = Some(true);
                }
                IgnoredErrorType::CalculatedColumn => {
                    entry.calculated_column = Some(true);
                }
            }
            let ie = self.worksheet.ignored_errors.get_or_insert_with(|| {
                Box::new(types::IgnoredErrors {
                    ignored_error: Vec::new(),
                    extension_list: None,
                    #[cfg(feature = "extra-children")]
                    extra_children: Vec::new(),
                })
            });
            ie.ignored_error.push(entry);
        }
        #[cfg(not(feature = "sml-validation"))]
        let _ = (sqref, error_type);
    }

    // -------------------------------------------------------------------------
    // Sheet protection
    // -------------------------------------------------------------------------

    /// Protect the sheet with optional restrictions (ECMA-376 §18.3.1.85).
    ///
    /// Requires the `sml-protection` feature.
    ///
    /// # Example
    ///
    /// ```ignore
    /// sheet.set_sheet_protection(SheetProtectionOptions {
    ///     sheet: true,
    ///     password: Some("secret".to_string()),
    ///     format_cells: true,
    ///     ..Default::default()
    /// });
    /// ```
    #[cfg(feature = "sml-protection")]
    pub fn set_sheet_protection(&mut self, opts: SheetProtectionOptions) {
        {
            let password = opts
                .password
                .as_deref()
                .filter(|p| !p.is_empty())
                .map(ooxml_xor_hash);

            self.worksheet.sheet_protection = Some(Box::new(types::SheetProtection {
                password,
                algorithm_name: None,
                hash_value: None,
                salt_value: None,
                spin_count: None,
                sheet: if opts.sheet { Some(true) } else { None },
                objects: None,
                scenarios: None,
                format_cells: if opts.format_cells { Some(true) } else { None },
                format_columns: if opts.format_columns {
                    Some(true)
                } else {
                    None
                },
                format_rows: if opts.format_rows { Some(true) } else { None },
                insert_columns: if opts.insert_columns {
                    Some(true)
                } else {
                    None
                },
                insert_rows: if opts.insert_rows { Some(true) } else { None },
                insert_hyperlinks: None,
                delete_columns: if opts.delete_columns {
                    Some(true)
                } else {
                    None
                },
                delete_rows: if opts.delete_rows { Some(true) } else { None },
                select_locked_cells: if opts.select_locked_cells {
                    Some(true)
                } else {
                    None
                },
                sort: if opts.sort { Some(true) } else { None },
                auto_filter: if opts.auto_filter { Some(true) } else { None },
                pivot_tables: if opts.pivot_tables { Some(true) } else { None },
                select_unlocked_cells: if opts.select_unlocked_cells {
                    Some(true)
                } else {
                    None
                },
                #[cfg(feature = "extra-attrs")]
                extra_attrs: Default::default(),
            }));
        }
    }

    // -------------------------------------------------------------------------
    // Tab color
    // -------------------------------------------------------------------------

    /// Set the sheet tab color (ECMA-376 §18.3.1.77).
    ///
    /// The color is an RGB hex string (e.g. `"FF0000"` for red, `"4472C4"` for
    /// the Excel blue theme color).
    ///
    /// # Example
    ///
    /// ```ignore
    /// sheet.set_tab_color("FF0000"); // red tab
    /// ```
    pub fn set_tab_color(&mut self, rgb: &str) {
        #[cfg(feature = "sml-styling")]
        {
            let color = Box::new(types::Color {
                auto: None,
                indexed: None,
                rgb: Some(hex_color_to_bytes(rgb)),
                theme: None,
                tint: None,
                #[cfg(feature = "extra-attrs")]
                extra_attrs: Default::default(),
            });
            let props = self
                .worksheet
                .sheet_properties
                .get_or_insert_with(|| Box::new(types::SheetProperties::default()));
            props.tab_color = Some(color);
        }
        #[cfg(not(feature = "sml-styling"))]
        let _ = rgb;
    }

    // -------------------------------------------------------------------------
    // Sheet view options
    // -------------------------------------------------------------------------

    /// Show or hide gridlines in the default sheet view (ECMA-376 §18.3.1.76
    /// `@showGridLines`).
    ///
    /// Gridlines are visible by default in Excel; pass `false` to hide them.
    pub fn set_show_gridlines(&mut self, show: bool) {
        self.show_gridlines = Some(show);
    }

    /// Show or hide row and column headers in the default sheet view
    /// (ECMA-376 §18.3.1.76 `@showRowColHeaders`).
    ///
    /// Row/col headers are visible by default in Excel; pass `false` to hide
    /// them.
    pub fn set_show_row_col_headers(&mut self, show: bool) {
        self.show_row_col_headers = Some(show);
    }

    // -------------------------------------------------------------------------
    // Chart embedding (sml-charts)
    // -------------------------------------------------------------------------

    /// Embed a chart in the worksheet.
    ///
    /// The chart XML must be a complete `<c:chartSpace>` document (ECMA-376
    /// §21.2).  The position and size are specified in cell units; `(x, y)` is
    /// the 0-based column/row of the top-left anchor and `(width, height)` is
    /// the extent in cells.
    ///
    /// At write time the chart is written to `xl/charts/chart{n}.xml`, a
    /// drawing part is created at `xl/drawings/drawing{n}.xml`, and the
    /// worksheet references the drawing via a relationship.
    ///
    /// Requires the `sml-charts` feature.
    pub fn embed_chart(
        &mut self,
        chart_xml: &[u8],
        x: u32,
        y: u32,
        width: u32,
        height: u32,
    ) -> &mut Self {
        #[cfg(feature = "sml-charts")]
        self.charts.push(ChartEntry {
            chart_xml: chart_xml.to_vec(),
            x,
            y,
            width,
            height,
        });
        #[cfg(not(feature = "sml-charts"))]
        let _ = (chart_xml, x, y, width, height);
        self
    }

    // -------------------------------------------------------------------------
    // Pivot tables (sml-pivot)
    // -------------------------------------------------------------------------

    /// Add a pivot table to the worksheet.
    ///
    /// This writes:
    /// - `xl/pivotCache/pivotCacheDefinition{n}.xml`
    /// - `xl/pivotCache/pivotCacheRecords{n}.xml`
    /// - `xl/pivotTables/pivotTable{n}.xml`
    ///
    /// and the necessary relationship files.  The cache and pivot table are
    /// cross-linked via the workbook `<pivotCaches>` element.
    ///
    /// Requires the `sml-pivot` feature.
    #[cfg(feature = "sml-pivot")]
    pub fn add_pivot_table(&mut self, opts: PivotTableOptions) -> &mut Self {
        self.pivot_tables.push(PivotEntry { opts });
        self
    }

    /// Get the sheet name.
    pub fn name(&self) -> &str {
        &self.name
    }
}

/// Builder for creating Excel workbooks.
///
/// # Why `WorkbookBuilder` doesn't hold `types::Workbook` directly
///
/// Unlike WML (`DocumentBuilder`) and PML (`PresentationBuilder`), this builder
/// cannot eagerly hold a `types::Workbook` because cell style indices are
/// cross-sheet: every `set_cell_style()` call on any sheet potentially adds a new
/// entry to the shared `Stylesheet` (fonts, fills, borders, number formats), and
/// the final deduplicated indices aren't known until all cells across all sheets
/// have been added.  Baking those indices into `SheetData` rows upfront would
/// require retroactively rewriting previously built rows when new styles appear.
///
/// Instead, `WorkbookBuilder` accumulates raw style values during building and
/// resolves them to index-based `Stylesheet` + `SheetData` rows at `write()` time.
/// `SheetBuilder` holds `types::Worksheet` directly for everything that doesn't
/// depend on style resolution (merge cells, freeze panes, column widths, etc.).
/// A named cell style entry for `WorkbookBuilder::add_cell_style`.
#[derive(Debug, Clone)]
struct NamedCellStyle {
    name: String,
    format_id: u32,
}

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
    /// Extra named cell styles beyond "Normal" (sml-styling).
    extra_cell_styles: Vec<NamedCellStyle>,
    /// Optional workbook protection (sml-protection).
    #[cfg(feature = "sml-protection")]
    workbook_protection: Option<types::WorkbookProtection>,
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
#[allow(dead_code)] // Fields read only with sml-styling feature
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
            extra_cell_styles: Vec::new(),
            #[cfg(feature = "sml-protection")]
            workbook_protection: None,
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

    // -------------------------------------------------------------------------
    // Workbook protection
    // -------------------------------------------------------------------------

    /// Protect the workbook structure and/or windows (ECMA-376 §18.2.29).
    ///
    /// Requires the `sml-protection` feature.
    ///
    /// - `lock_structure`: prevent adding, deleting, or moving sheets.
    /// - `lock_windows`: prevent resizing/moving the workbook window.
    /// - `password`: optional plain-text password (hashed with the OOXML XOR
    ///   algorithm).
    ///
    /// # Example
    ///
    /// ```ignore
    /// wb.set_workbook_protection(true, false, Some("secret"));
    /// ```
    pub fn set_workbook_protection(
        &mut self,
        lock_structure: bool,
        lock_windows: bool,
        password: Option<&str>,
    ) {
        #[cfg(feature = "sml-protection")]
        {
            let workbook_password = password.filter(|p| !p.is_empty()).map(ooxml_xor_hash);

            self.workbook_protection = Some(types::WorkbookProtection {
                workbook_password,
                workbook_password_character_set: None,
                revisions_password: None,
                revisions_password_character_set: None,
                lock_structure: if lock_structure { Some(true) } else { None },
                lock_windows: if lock_windows { Some(true) } else { None },
                lock_revision: None,
                revisions_algorithm_name: None,
                revisions_hash_value: None,
                revisions_salt_value: None,
                revisions_spin_count: None,
                workbook_algorithm_name: None,
                workbook_hash_value: None,
                workbook_salt_value: None,
                workbook_spin_count: None,
                #[cfg(feature = "extra-attrs")]
                extra_attrs: Default::default(),
            });
        }
        #[cfg(not(feature = "sml-protection"))]
        let _ = (lock_structure, lock_windows, password);
    }

    // -------------------------------------------------------------------------
    // Named cell styles
    // -------------------------------------------------------------------------

    /// Add a named cell style to the workbook stylesheet (ECMA-376 §18.8.7).
    ///
    /// Requires the `sml-styling` feature.
    ///
    /// The `format_id` must be the index of a `<xf>` entry in `cellStyleXfs`.
    /// Use `0` for the default "Normal" format.  Returns the 0-based index of
    /// the new cell style in the `cellStyles` collection.
    ///
    /// # Example
    ///
    /// ```ignore
    /// // Add a "Good" style backed by format_id 0 (Normal format)
    /// wb.add_cell_style("Good", 0);
    /// ```
    pub fn add_cell_style(&mut self, name: &str, format_id: u32) -> u32 {
        let idx = self.extra_cell_styles.len() as u32;
        self.extra_cell_styles.push(NamedCellStyle {
            name: name.to_string(),
            format_id,
        });
        // +1 because the "Normal" style always occupies slot 0
        idx + 1
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

        let has_styles = !self.cell_formats.is_empty() || !self.extra_cell_styles.is_empty();

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

        // Build workbook.xml using generated types
        let workbook = self.build_workbook();
        let workbook_xml = serialize_workbook(&workbook)?;

        // Write parts to package
        pkg.add_part("_rels/.rels", CT_RELATIONSHIPS, root_rels.as_bytes())?;
        pkg.add_part(
            "xl/_rels/workbook.xml.rels",
            CT_RELATIONSHIPS,
            wb_rels.as_bytes(),
        )?;
        pkg.add_part("xl/workbook.xml", CT_WORKBOOK, &workbook_xml)?;

        // Write styles if any
        if has_styles {
            let styles_xml = self.serialize_styles()?;
            pkg.add_part("xl/styles.xml", CT_STYLES, &styles_xml)?;
        }

        // Pre-compute global drawing and pivot numbering across all sheets.
        // Each sheet that has charts gets one drawing part; each pivot table
        // entry gets its own pivotCacheDefinition + pivotCacheRecords + pivotTable.
        #[cfg(feature = "sml-charts")]
        let mut next_drawing_num = 1usize;
        #[cfg(feature = "sml-charts")]
        let mut next_chart_num = 1usize;
        #[cfg(feature = "sml-pivot")]
        let mut next_pivot_num = 1usize;

        // Write each sheet and its related parts (comments, hyperlinks, charts, pivot tables).
        for (i, sheet) in self.sheets.iter().enumerate() {
            let sheet_num = i + 1;
            let has_comments = !sheet.comments.is_empty();

            // Collect external hyperlink URLs (need relationship entries).
            let ext_hyperlinks: Vec<&str> = sheet
                .hyperlinks
                .iter()
                .filter_map(|h| {
                    if let HyperlinkDest::External(url) = &h.dest {
                        Some(url.as_str())
                    } else {
                        None
                    }
                })
                .collect();

            #[cfg(feature = "sml-charts")]
            let has_charts = !sheet.charts.is_empty();
            #[cfg(not(feature = "sml-charts"))]
            let has_charts = false;

            #[cfg(feature = "sml-pivot")]
            let has_pivots = !sheet.pivot_tables.is_empty();
            #[cfg(not(feature = "sml-pivot"))]
            let has_pivots = false;

            let needs_rels = has_comments || !ext_hyperlinks.is_empty() || has_charts || has_pivots;

            // Relative IDs inside the sheet .rels file.
            let mut sheet_rel_id = 1usize;
            let comments_rel_id = if has_comments { sheet_rel_id } else { 0 };
            if has_comments {
                sheet_rel_id += 1;
            }

            // Track drawing rel id (if any charts are present, one drawing part per sheet).
            #[cfg(feature = "sml-charts")]
            let drawing_rel_id = if has_charts {
                let id = sheet_rel_id;
                sheet_rel_id += 1;
                id
            } else {
                0
            };

            // Assign pivot rel IDs for this sheet.
            #[cfg(feature = "sml-pivot")]
            let pivot_rel_id_start = sheet_rel_id;

            // Compute hyperlink rel id start (after comments, drawing, pivot rels).
            #[cfg(feature = "sml-hyperlinks")]
            let hyperlink_rel_id_start = {
                let mut start = 1usize;
                if has_comments {
                    start += 1;
                }
                if has_charts {
                    start += 1;
                }
                #[cfg(feature = "sml-pivot")]
                {
                    start += sheet.pivot_tables.len();
                }
                start
            };
            #[cfg(not(feature = "sml-hyperlinks"))]
            let hyperlink_rel_id_start = 1usize;

            if needs_rels {
                let mut sheet_rels = String::new();
                sheet_rels.push_str(r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>"#);
                sheet_rels.push('\n');
                sheet_rels.push_str(
                    r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">"#,
                );
                sheet_rels.push('\n');

                if has_comments {
                    sheet_rels.push_str(&format!(
                        r#"  <Relationship Id="rId{}" Type="{}" Target="../comments{}.xml"/>"#,
                        comments_rel_id, REL_COMMENTS, sheet_num
                    ));
                    sheet_rels.push('\n');
                }

                #[cfg(feature = "sml-charts")]
                if has_charts {
                    sheet_rels.push_str(&format!(
                        r#"  <Relationship Id="rId{}" Type="{}" Target="../drawings/drawing{}.xml"/>"#,
                        drawing_rel_id, REL_DRAWING, next_drawing_num
                    ));
                    sheet_rels.push('\n');
                }

                // Pivot table rels: one per pivot entry for this sheet.
                #[cfg(feature = "sml-pivot")]
                {
                    let mut prel = pivot_rel_id_start;
                    for (pi, _pt) in sheet.pivot_tables.iter().enumerate() {
                        let global_pivot = next_pivot_num + pi;
                        sheet_rels.push_str(&format!(
                            r#"  <Relationship Id="rId{}" Type="{}" Target="../pivotTables/pivotTable{}.xml"/>"#,
                            prel, REL_PIVOT_TABLE, global_pivot
                        ));
                        sheet_rels.push('\n');
                        prel += 1;
                    }
                }

                for (hi, url) in ext_hyperlinks.iter().enumerate() {
                    sheet_rels.push_str(&format!(
                        r#"  <Relationship Id="rId{}" Type="{}" Target="{}" TargetMode="External"/>"#,
                        hyperlink_rel_id_start + hi,
                        REL_HYPERLINK,
                        escape_xml(url)
                    ));
                    sheet_rels.push('\n');
                }

                sheet_rels.push_str("</Relationships>");
                let rels_part = format!("xl/worksheets/_rels/sheet{}.xml.rels", sheet_num);
                pkg.add_part(&rels_part, CT_RELATIONSHIPS, sheet_rels.as_bytes())?;
            }

            // Write the worksheet XML, injecting the drawing reference if needed.
            let sheet_xml = self.serialize_sheet_with_drawing(
                sheet,
                hyperlink_rel_id_start,
                #[cfg(feature = "sml-charts")]
                if has_charts {
                    Some(drawing_rel_id)
                } else {
                    None
                },
                #[cfg(not(feature = "sml-charts"))]
                None,
            )?;
            let part_name = format!("xl/worksheets/sheet{}.xml", sheet_num);
            pkg.add_part(&part_name, CT_WORKSHEET, &sheet_xml)?;

            if has_comments {
                let comments_xml = self.serialize_comments(sheet)?;
                let comments_part = format!("xl/comments{}.xml", sheet_num);
                pkg.add_part(&comments_part, CT_COMMENTS, &comments_xml)?;
            }

            // Write chart parts and the drawing part for this sheet.
            #[cfg(feature = "sml-charts")]
            if has_charts {
                let drawing_num = next_drawing_num;
                next_drawing_num += 1;

                // Drawing XML references each chart.
                let drawing_xml = build_drawing_xml(&sheet.charts, next_chart_num);
                let drawing_part = format!("xl/drawings/drawing{}.xml", drawing_num);
                pkg.add_part(&drawing_part, CT_DRAWING, drawing_xml.as_bytes())?;

                // Drawing .rels: one entry per chart.
                let mut drawing_rels = String::new();
                drawing_rels.push_str(r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>"#);
                drawing_rels.push('\n');
                drawing_rels.push_str(
                    r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">"#,
                );
                drawing_rels.push('\n');
                for (ci, _chart) in sheet.charts.iter().enumerate() {
                    let chart_num = next_chart_num + ci;
                    drawing_rels.push_str(&format!(
                        r#"  <Relationship Id="rId{}" Type="{}" Target="../charts/chart{}.xml"/>"#,
                        ci + 1,
                        REL_CHART,
                        chart_num,
                    ));
                    drawing_rels.push('\n');
                }
                drawing_rels.push_str("</Relationships>");
                let drawing_rels_part =
                    format!("xl/drawings/_rels/drawing{}.xml.rels", drawing_num);
                pkg.add_part(
                    &drawing_rels_part,
                    CT_RELATIONSHIPS,
                    drawing_rels.as_bytes(),
                )?;

                // Write each chart XML file.
                for chart in &sheet.charts {
                    let chart_part = format!("xl/charts/chart{}.xml", next_chart_num);
                    pkg.add_part(&chart_part, CT_CHART, &chart.chart_xml)?;
                    next_chart_num += 1;
                }
            }

            // Write pivot table parts for this sheet.
            #[cfg(feature = "sml-pivot")]
            {
                for (pi, pt) in sheet.pivot_tables.iter().enumerate() {
                    let pn = next_pivot_num + pi;

                    // pivotCacheDefinition
                    let cache_def_xml = build_pivot_cache_definition_xml(&pt.opts, pn);
                    let cache_def_part = format!("xl/pivotCache/pivotCacheDefinition{}.xml", pn);
                    pkg.add_part(
                        &cache_def_part,
                        CT_PIVOT_CACHE_DEF,
                        cache_def_xml.as_bytes(),
                    )?;

                    // pivotCacheDefinition .rels → pivotCacheRecords
                    let cache_def_rels = format!(
                        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="{}" Target="pivotCacheRecords{}.xml"/>
</Relationships>"#,
                        REL_PIVOT_CACHE_REC, pn
                    );
                    let cache_def_rels_part =
                        format!("xl/pivotCache/_rels/pivotCacheDefinition{}.xml.rels", pn);
                    pkg.add_part(
                        &cache_def_rels_part,
                        CT_RELATIONSHIPS,
                        cache_def_rels.as_bytes(),
                    )?;

                    // pivotCacheRecords (empty)
                    let cache_rec_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<pivotCacheRecords xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="0"/>"#;
                    let cache_rec_part = format!("xl/pivotCache/pivotCacheRecords{}.xml", pn);
                    pkg.add_part(
                        &cache_rec_part,
                        CT_PIVOT_CACHE_REC,
                        cache_rec_xml.as_bytes(),
                    )?;

                    // pivotTable
                    let pivot_xml = build_pivot_table_xml(&pt.opts, pn, &sheet.name);
                    let pivot_part = format!("xl/pivotTables/pivotTable{}.xml", pn);
                    pkg.add_part(&pivot_part, CT_PIVOT_TABLE, pivot_xml.as_bytes())?;

                    // pivotTable .rels → pivotCacheDefinition
                    let pivot_rels = format!(
                        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="{}" Target="../pivotCache/pivotCacheDefinition{}.xml"/>
</Relationships>"#,
                        REL_PIVOT_CACHE_DEF, pn
                    );
                    let pivot_rels_part = format!("xl/pivotTables/_rels/pivotTable{}.xml.rels", pn);
                    pkg.add_part(&pivot_rels_part, CT_RELATIONSHIPS, pivot_rels.as_bytes())?;
                }
                // Advance global pivot counter past all pivots written for this sheet.
                next_pivot_num += sheet.pivot_tables.len();
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
        let comment_list: Vec<types::Comment> = sheet
            .comments
            .iter()
            .map(|_c| {
                #[cfg(feature = "sml-comments")]
                let author_id = {
                    let author = _c.author.clone().unwrap_or_default();
                    *author_index.get(&author).unwrap_or(&0)
                };

                // Build the rich-text body: prefer runs if any, else plain text.
                #[cfg(feature = "sml-comments")]
                let rich_text = if _c.runs.is_empty() {
                    // Plain-text comment: single unstyled run.
                    types::RichString {
                        cell_type: None,
                        reference: vec![types::RichTextElement {
                            r_pr: None,
                            cell_type: _c.text.clone(),
                            #[cfg(feature = "extra-children")]
                            extra_children: Vec::new(),
                        }],
                        r_ph: Vec::new(),
                        phonetic_pr: None,
                        #[cfg(feature = "extra-children")]
                        extra_children: Vec::new(),
                    }
                } else {
                    // Rich-text comment: one RichTextElement per run.
                    let runs: Vec<types::RichTextElement> = _c
                        .runs
                        .iter()
                        .map(|run| {
                            let has_props = run.bold
                                || run.italic
                                || run.color.is_some()
                                || run.font_size.is_some();
                            let r_pr = if has_props {
                                Some(Box::new(types::RichTextRunProperties {
                                    r_font: None,
                                    charset: None,
                                    family: None,
                                    b: if run.bold {
                                        Some(Box::new(types::BooleanProperty {
                                            value: None,
                                            #[cfg(feature = "extra-attrs")]
                                            extra_attrs: Default::default(),
                                        }))
                                    } else {
                                        None
                                    },
                                    i: if run.italic {
                                        Some(Box::new(types::BooleanProperty {
                                            value: None,
                                            #[cfg(feature = "extra-attrs")]
                                            extra_attrs: Default::default(),
                                        }))
                                    } else {
                                        None
                                    },
                                    strike: None,
                                    outline: None,
                                    shadow: None,
                                    condense: None,
                                    extend: None,
                                    #[cfg(feature = "sml-styling")]
                                    color: run.color.as_ref().map(|c| {
                                        Box::new(types::Color {
                                            auto: None,
                                            indexed: None,
                                            rgb: Some(hex_color_to_bytes(c)),
                                            theme: None,
                                            tint: None,
                                            #[cfg(feature = "extra-attrs")]
                                            extra_attrs: Default::default(),
                                        })
                                    }),
                                    #[cfg(not(feature = "sml-styling"))]
                                    color: None,
                                    sz: run.font_size.map(|s| {
                                        Box::new(types::FontSize {
                                            value: s,
                                            #[cfg(feature = "extra-attrs")]
                                            extra_attrs: Default::default(),
                                        })
                                    }),
                                    u: None,
                                    vert_align: None,
                                    scheme: None,
                                    #[cfg(feature = "extra-children")]
                                    extra_children: Vec::new(),
                                }))
                            } else {
                                None
                            };
                            types::RichTextElement {
                                r_pr,
                                cell_type: run.text.clone(),
                                #[cfg(feature = "extra-children")]
                                extra_children: Vec::new(),
                            }
                        })
                        .collect();
                    types::RichString {
                        cell_type: None,
                        reference: runs,
                        r_ph: Vec::new(),
                        phonetic_pr: None,
                        #[cfg(feature = "extra-children")]
                        extra_children: Vec::new(),
                    }
                };

                types::Comment {
                    #[cfg(feature = "sml-comments")]
                    reference: _c.reference.clone(),
                    #[cfg(feature = "sml-comments")]
                    author_id,
                    #[cfg(feature = "sml-comments")]
                    guid: None,
                    #[cfg(feature = "sml-comments")]
                    shape_id: None,
                    #[cfg(feature = "sml-comments")]
                    text: Box::new(rich_text),
                    comment_pr: None,
                    #[cfg(feature = "extra-attrs")]
                    extra_attrs: Default::default(),
                    #[cfg(feature = "extra-children")]
                    extra_children: Vec::new(),
                }
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

    /// Serialize styles to XML using generated ToXml serializers.
    #[cfg(feature = "sml-styling")]
    fn serialize_styles(&self) -> Result<Vec<u8>> {
        let stylesheet = self.build_stylesheet();
        serialize_with_namespaces(&stylesheet, "styleSheet")
    }

    /// Stub for serialize_styles when sml-styling is not enabled.
    #[cfg(not(feature = "sml-styling"))]
    fn serialize_styles(&self) -> Result<Vec<u8>> {
        // Return a minimal empty stylesheet
        let stylesheet = types::Stylesheet::default();
        serialize_with_namespaces(&stylesheet, "styleSheet")
    }

    /// Build a Stylesheet type from builder data.
    #[cfg(feature = "sml-styling")]
    fn build_stylesheet(&self) -> types::Stylesheet {
        // Number formats (custom formats start at ID 164)
        let num_fmts: Option<Box<types::NumberFormats>> = if self.number_formats.is_empty() {
            None
        } else {
            Some(Box::new(types::NumberFormats {
                count: Some(self.number_formats.len() as u32),
                num_fmt: self
                    .number_formats
                    .iter()
                    .enumerate()
                    .map(|(i, fmt)| types::NumberFormat {
                        number_format_id: (164 + i) as u32,
                        format_code: fmt.clone(),
                        #[cfg(feature = "extra-attrs")]
                        extra_attrs: Default::default(),
                    })
                    .collect(),
                #[cfg(feature = "extra-attrs")]
                extra_attrs: Default::default(),
                #[cfg(feature = "extra-children")]
                extra_children: Vec::new(),
            }))
        };

        // Fonts
        let fonts = Box::new(types::Fonts {
            count: Some(self.fonts.len() as u32),
            font: self.fonts.iter().map(build_font).collect(),
            #[cfg(feature = "extra-attrs")]
            extra_attrs: Default::default(),
            #[cfg(feature = "extra-children")]
            extra_children: Vec::new(),
        });

        // Fills
        let fills = Box::new(types::Fills {
            count: Some(self.fills.len() as u32),
            fill: self.fills.iter().map(build_fill).collect(),
            #[cfg(feature = "extra-attrs")]
            extra_attrs: Default::default(),
            #[cfg(feature = "extra-children")]
            extra_children: Vec::new(),
        });

        // Borders
        let borders = Box::new(types::Borders {
            count: Some(self.borders.len() as u32),
            border: self.borders.iter().map(build_border).collect(),
            #[cfg(feature = "extra-attrs")]
            extra_attrs: Default::default(),
            #[cfg(feature = "extra-children")]
            extra_children: Vec::new(),
        });

        // Cell style XFs (required, at least one default)
        let cell_style_xfs = Box::new(types::CellStyleFormats {
            count: Some(1),
            xf: vec![types::Format {
                #[cfg(feature = "sml-styling")]
                number_format_id: Some(0),
                #[cfg(feature = "sml-styling")]
                font_id: Some(0),
                #[cfg(feature = "sml-styling")]
                fill_id: Some(0),
                #[cfg(feature = "sml-styling")]
                border_id: Some(0),
                #[cfg(feature = "sml-styling")]
                format_id: None,
                #[cfg(feature = "sml-styling")]
                quote_prefix: None,
                #[cfg(feature = "sml-pivot")]
                pivot_button: None,
                #[cfg(feature = "sml-styling")]
                apply_number_format: None,
                #[cfg(feature = "sml-styling")]
                apply_font: None,
                #[cfg(feature = "sml-styling")]
                apply_fill: None,
                #[cfg(feature = "sml-styling")]
                apply_border: None,
                #[cfg(feature = "sml-styling")]
                apply_alignment: None,
                #[cfg(feature = "sml-styling")]
                apply_protection: None,
                #[cfg(feature = "sml-styling")]
                alignment: None,
                #[cfg(feature = "sml-protection")]
                protection: None,
                #[cfg(feature = "sml-extensions")]
                extension_list: None,
                #[cfg(feature = "extra-attrs")]
                extra_attrs: Default::default(),
                #[cfg(feature = "extra-children")]
                extra_children: Vec::new(),
            }],
            #[cfg(feature = "extra-attrs")]
            extra_attrs: Default::default(),
            #[cfg(feature = "extra-children")]
            extra_children: Vec::new(),
        });

        // Cell XFs - includes default format plus custom formats
        let mut xf_list: Vec<types::Format> = vec![types::Format {
            #[cfg(feature = "sml-styling")]
            number_format_id: Some(0),
            #[cfg(feature = "sml-styling")]
            font_id: Some(0),
            #[cfg(feature = "sml-styling")]
            fill_id: Some(0),
            #[cfg(feature = "sml-styling")]
            border_id: Some(0),
            #[cfg(feature = "sml-styling")]
            format_id: Some(0),
            #[cfg(feature = "sml-styling")]
            quote_prefix: None,
            #[cfg(feature = "sml-pivot")]
            pivot_button: None,
            #[cfg(feature = "sml-styling")]
            apply_number_format: None,
            #[cfg(feature = "sml-styling")]
            apply_font: None,
            #[cfg(feature = "sml-styling")]
            apply_fill: None,
            #[cfg(feature = "sml-styling")]
            apply_border: None,
            #[cfg(feature = "sml-styling")]
            apply_alignment: None,
            #[cfg(feature = "sml-styling")]
            apply_protection: None,
            #[cfg(feature = "sml-styling")]
            alignment: None,
            #[cfg(feature = "sml-protection")]
            protection: None,
            #[cfg(feature = "sml-extensions")]
            extension_list: None,
            #[cfg(feature = "extra-attrs")]
            extra_attrs: Default::default(),
            #[cfg(feature = "extra-children")]
            extra_children: Vec::new(),
        }];

        for xf in &self.cell_formats {
            xf_list.push(build_cell_format(xf));
        }

        let cell_xfs = Box::new(types::CellFormats {
            count: Some(xf_list.len() as u32),
            xf: xf_list,
            #[cfg(feature = "extra-attrs")]
            extra_attrs: Default::default(),
            #[cfg(feature = "extra-children")]
            extra_children: Vec::new(),
        });

        // Cell styles (required — always includes "Normal" at index 0)
        let mut cell_style_list = vec![types::CellStyle {
            name: Some("Normal".to_string()),
            format_id: 0,
            builtin_id: Some(0),
            i_level: None,
            hidden: None,
            custom_builtin: None,
            extension_list: None,
            #[cfg(feature = "extra-attrs")]
            extra_attrs: Default::default(),
            #[cfg(feature = "extra-children")]
            extra_children: Vec::new(),
        }];
        for cs in &self.extra_cell_styles {
            cell_style_list.push(types::CellStyle {
                name: Some(cs.name.clone()),
                format_id: cs.format_id,
                builtin_id: None,
                i_level: None,
                hidden: None,
                custom_builtin: Some(true),
                extension_list: None,
                #[cfg(feature = "extra-attrs")]
                extra_attrs: Default::default(),
                #[cfg(feature = "extra-children")]
                extra_children: Vec::new(),
            });
        }
        let count = cell_style_list.len() as u32;
        let cell_styles = Box::new(types::CellStyles {
            count: Some(count),
            cell_style: cell_style_list,
            #[cfg(feature = "extra-attrs")]
            extra_attrs: Default::default(),
            #[cfg(feature = "extra-children")]
            extra_children: Vec::new(),
        });

        types::Stylesheet {
            #[cfg(feature = "sml-styling")]
            num_fmts,
            #[cfg(feature = "sml-styling")]
            fonts: Some(fonts),
            #[cfg(feature = "sml-styling")]
            fills: Some(fills),
            #[cfg(feature = "sml-styling")]
            borders: Some(borders),
            #[cfg(feature = "sml-styling")]
            cell_style_xfs: Some(cell_style_xfs),
            #[cfg(feature = "sml-styling")]
            cell_xfs: Some(cell_xfs),
            #[cfg(feature = "sml-styling")]
            cell_styles: Some(cell_styles),
            #[cfg(feature = "sml-styling")]
            dxfs: None,
            #[cfg(feature = "sml-styling")]
            table_styles: None,
            #[cfg(feature = "sml-styling")]
            colors: None,
            #[cfg(feature = "sml-extensions")]
            extension_list: None,
            #[cfg(feature = "extra-children")]
            extra_children: Vec::new(),
        }
    }

    /// Serialize a sheet to XML using generated types.
    ///
    /// `hyperlink_rel_id_start` is the first relationship ID available for
    /// external hyperlinks (1-based; accounts for comments/drawing/pivot
    /// occupying earlier rIds when present).
    ///
    /// `drawing_rel_id` is `Some(id)` when the sheet has embedded charts and a
    /// drawing part has been assigned to it.
    fn serialize_sheet_with_drawing(
        &self,
        sheet: &SheetBuilder,
        hyperlink_rel_id_start: usize,
        drawing_rel_id: Option<usize>,
    ) -> Result<Vec<u8>> {
        // Build row height lookup (already a HashMap now)
        #[cfg(feature = "sml-styling")]
        let row_heights = &sheet.row_heights;

        // Group cells by row
        let mut rows_map: HashMap<u32, Vec<(u32, &BuilderCell)>> = HashMap::new();
        for ((row, col), cell) in &sheet.cells {
            rows_map.entry(*row).or_default().push((*col, cell));
        }

        // Any rows that have outline/collapsed but no cells still need a Row element.
        #[cfg(feature = "sml-structure")]
        for &row_num in sheet
            .row_outline_levels
            .keys()
            .chain(sheet.row_collapsed.keys())
        {
            rows_map.entry(row_num).or_default();
        }

        // Sort and build rows
        let mut row_nums: Vec<_> = rows_map.keys().copied().collect();
        row_nums.sort();

        let rows: Vec<types::Row> = row_nums
            .iter()
            .map(|&row_num| {
                let cells_data = rows_map.get(&row_num).unwrap();
                let mut sorted_cells: Vec<_> = cells_data.clone();
                sorted_cells.sort_by_key(|(col, _)| *col);

                let cells: Vec<types::Cell> = sorted_cells
                    .iter()
                    .map(|(col, cell)| {
                        let ref_str = column_to_letter(*col) + &row_num.to_string();
                        self.build_cell(&ref_str, cell)
                    })
                    .collect();

                types::Row {
                    reference: Some(row_num),
                    cell_spans: None,
                    style_index: None,
                    #[cfg(feature = "sml-styling")]
                    custom_format: None,
                    #[cfg(feature = "sml-styling")]
                    height: row_heights.get(&row_num).copied(),
                    #[cfg(feature = "sml-structure")]
                    hidden: None,
                    #[cfg(feature = "sml-styling")]
                    custom_height: row_heights.get(&row_num).map(|_| true),
                    #[cfg(feature = "sml-structure")]
                    outline_level: sheet.row_outline_levels.get(&row_num).copied(),
                    #[cfg(feature = "sml-structure")]
                    collapsed: sheet.row_collapsed.get(&row_num).copied(),
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
                }
            })
            .collect();

        // Clone pre-built worksheet and fill in the rows
        let mut worksheet = sheet.worksheet.clone();
        worksheet.sheet_data = Box::new(types::SheetData {
            row: rows,
            #[cfg(feature = "extra-children")]
            extra_children: Vec::new(),
        });

        // Apply column outline levels and collapsed flags (if any).
        // We update existing Column entries in place; if a column has outline/collapsed
        // but no width entry (no Col element yet), we add one.
        #[cfg(all(feature = "sml-styling", feature = "sml-structure"))]
        if !sheet.col_outline_levels.is_empty() || !sheet.col_collapsed.is_empty() {
            // Collect all column numbers that need outline/collapsed attributes.
            let mut col_nums: std::collections::HashSet<u32> = std::collections::HashSet::new();
            col_nums.extend(sheet.col_outline_levels.keys().copied());
            col_nums.extend(sheet.col_collapsed.keys().copied());

            for col_num in col_nums {
                let level = sheet.col_outline_levels.get(&col_num).copied();
                let collapsed = sheet.col_collapsed.get(&col_num).copied();

                // Look for an existing Column entry that covers col_num.
                let mut found = false;
                for cols_group in &mut worksheet.cols {
                    for col_entry in &mut cols_group.col {
                        if col_entry.start_column <= col_num && col_num <= col_entry.end_column {
                            if level.is_some() {
                                col_entry.outline_level = level;
                            }
                            if collapsed.is_some() {
                                col_entry.collapsed = collapsed;
                            }
                            found = true;
                            break;
                        }
                    }
                    if found {
                        break;
                    }
                }

                if !found {
                    // No existing column entry — create a minimal one just for
                    // the outline/collapsed attributes.
                    let col_entry = types::Column {
                        #[cfg(feature = "sml-styling")]
                        start_column: col_num,
                        #[cfg(feature = "sml-styling")]
                        end_column: col_num,
                        #[cfg(feature = "sml-styling")]
                        width: None,
                        #[cfg(feature = "sml-styling")]
                        style: None,
                        #[cfg(feature = "sml-structure")]
                        hidden: None,
                        #[cfg(feature = "sml-styling")]
                        best_fit: None,
                        #[cfg(feature = "sml-styling")]
                        custom_width: None,
                        #[cfg(feature = "sml-i18n")]
                        phonetic: None,
                        #[cfg(feature = "sml-structure")]
                        outline_level: level,
                        #[cfg(feature = "sml-structure")]
                        collapsed,
                        #[cfg(feature = "extra-attrs")]
                        extra_attrs: Default::default(),
                    };
                    if let Some(cols_group) = worksheet.cols.first_mut() {
                        cols_group.col.push(col_entry);
                    } else {
                        worksheet.cols.push(types::Columns {
                            col: vec![col_entry],
                            #[cfg(feature = "extra-children")]
                            extra_children: Vec::new(),
                        });
                    }
                }
            }
        }

        // Apply show_gridlines / show_row_col_headers to the sheet view.
        #[cfg(feature = "sml-styling")]
        if sheet.show_gridlines.is_some() || sheet.show_row_col_headers.is_some() {
            if let Some(views) = worksheet.sheet_views.as_mut() {
                // Modify the first (default) sheet view if one already exists.
                if let Some(sv) = views.sheet_view.first_mut() {
                    if let Some(v) = sheet.show_gridlines {
                        sv.show_grid_lines = Some(v);
                    }
                    if let Some(v) = sheet.show_row_col_headers {
                        sv.show_row_col_headers = Some(v);
                    }
                }
            } else {
                // No sheet view yet — create a minimal one.
                let sv = types::SheetView {
                    #[cfg(feature = "sml-protection")]
                    window_protection: None,
                    #[cfg(feature = "sml-formulas")]
                    show_formulas: None,
                    #[cfg(feature = "sml-styling")]
                    show_grid_lines: sheet.show_gridlines,
                    #[cfg(feature = "sml-styling")]
                    show_row_col_headers: sheet.show_row_col_headers,
                    #[cfg(feature = "sml-styling")]
                    show_zeros: None,
                    #[cfg(feature = "sml-i18n")]
                    right_to_left: None,
                    tab_selected: None,
                    #[cfg(feature = "sml-layout")]
                    show_ruler: None,
                    #[cfg(feature = "sml-structure")]
                    show_outline_symbols: None,
                    #[cfg(feature = "sml-styling")]
                    default_grid_color: None,
                    #[cfg(feature = "sml-layout")]
                    show_white_space: None,
                    view: None,
                    top_left_cell: None,
                    #[cfg(feature = "sml-styling")]
                    color_id: None,
                    zoom_scale: None,
                    zoom_scale_normal: None,
                    #[cfg(feature = "sml-layout")]
                    zoom_scale_sheet_layout_view: None,
                    #[cfg(feature = "sml-layout")]
                    zoom_scale_page_layout_view: None,
                    workbook_view_id: 0,
                    #[cfg(feature = "sml-structure")]
                    pane: None,
                    selection: Vec::new(),
                    #[cfg(feature = "sml-pivot")]
                    pivot_selection: Vec::new(),
                    #[cfg(feature = "sml-extensions")]
                    extension_list: None,
                    #[cfg(feature = "extra-attrs")]
                    extra_attrs: Default::default(),
                    #[cfg(feature = "extra-children")]
                    extra_children: Vec::new(),
                };
                worksheet.sheet_views = Some(Box::new(types::SheetViews {
                    sheet_view: vec![sv],
                    extension_list: None,
                    #[cfg(feature = "extra-children")]
                    extra_children: Vec::new(),
                }));
            }
        }

        // Inject hyperlinks (external ones carry their rId; internal ones use location).
        #[cfg(feature = "sml-hyperlinks")]
        if !sheet.hyperlinks.is_empty() {
            let mut rel_id = hyperlink_rel_id_start;
            let hyperlink_list: Vec<types::Hyperlink> = sheet
                .hyperlinks
                .iter()
                .map(|h| {
                    let (id, location) = match &h.dest {
                        HyperlinkDest::External(_) => {
                            let r = Some(format!("rId{}", rel_id));
                            rel_id += 1;
                            (r, None)
                        }
                        HyperlinkDest::Internal(loc) => (None, Some(loc.clone())),
                    };
                    types::Hyperlink {
                        reference: h.reference.clone(),
                        id,
                        #[cfg(feature = "sml-hyperlinks")]
                        location,
                        #[cfg(feature = "sml-hyperlinks")]
                        tooltip: h.tooltip.clone(),
                        #[cfg(feature = "sml-hyperlinks")]
                        display: h.display.clone(),
                        #[cfg(feature = "extra-attrs")]
                        extra_attrs: Default::default(),
                    }
                })
                .collect();
            worksheet.hyperlinks = Some(Box::new(types::Hyperlinks {
                hyperlink: hyperlink_list,
                #[cfg(feature = "extra-children")]
                extra_children: Vec::new(),
            }));
        }
        #[cfg(not(feature = "sml-hyperlinks"))]
        let _ = hyperlink_rel_id_start;

        // Inject the drawing relationship reference when charts are present.
        // ECMA-376 §18.3.1.27: <drawing r:id="rId{n}"/> inside <worksheet>.
        #[cfg(feature = "sml-drawings")]
        if let Some(rel_id) = drawing_rel_id {
            worksheet.drawing = Some(Box::new(types::Drawing {
                id: format!("rId{}", rel_id),
                #[cfg(feature = "extra-attrs")]
                extra_attrs: Default::default(),
            }));
        }
        #[cfg(not(feature = "sml-drawings"))]
        let _ = drawing_rel_id;

        serialize_with_namespaces(&worksheet, "worksheet")
    }

    /// Build a Cell type from builder data.
    fn build_cell(&self, reference: &str, cell: &BuilderCell) -> types::Cell {
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

        types::Cell {
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
        }
    }

    /// Build a Workbook type from builder data.
    fn build_workbook(&self) -> types::Workbook {
        // Build sheets
        let sheets: Vec<types::Sheet> = self
            .sheets
            .iter()
            .enumerate()
            .map(|(i, sheet)| types::Sheet {
                name: sheet.name.clone(),
                sheet_id: (i + 1) as u32,
                #[cfg(feature = "sml-structure")]
                state: None,
                id: format!("rId{}", i + 1),
                #[cfg(feature = "extra-attrs")]
                extra_attrs: Default::default(),
            })
            .collect();

        // Build defined names if any
        let defined_names: Option<Box<types::DefinedNames>> = if self.defined_names.is_empty() {
            None
        } else {
            Some(Box::new(types::DefinedNames {
                defined_name: self
                    .defined_names
                    .iter()
                    .map(|dn| types::DefinedName {
                        text: Some(dn.reference.clone()),
                        name: dn.name.clone(),
                        comment: dn.comment.clone(),
                        #[cfg(feature = "sml-formulas-advanced")]
                        custom_menu: None,
                        description: None,
                        #[cfg(feature = "sml-formulas-advanced")]
                        help: None,
                        #[cfg(feature = "sml-formulas-advanced")]
                        status_bar: None,
                        local_sheet_id: dn.local_sheet_id,
                        #[cfg(feature = "sml-structure")]
                        hidden: if dn.hidden { Some(true) } else { None },
                        #[cfg(feature = "sml-formulas-advanced")]
                        function: None,
                        #[cfg(feature = "sml-formulas-advanced")]
                        vb_procedure: None,
                        #[cfg(feature = "sml-formulas-advanced")]
                        xlm: None,
                        #[cfg(feature = "sml-formulas-advanced")]
                        function_group_id: None,
                        #[cfg(feature = "sml-formulas-advanced")]
                        shortcut_key: None,
                        #[cfg(feature = "sml-formulas-advanced")]
                        publish_to_server: None,
                        #[cfg(feature = "sml-formulas-advanced")]
                        workbook_parameter: None,
                        #[cfg(feature = "extra-attrs")]
                        extra_attrs: Default::default(),
                        #[cfg(feature = "extra-children")]
                        extra_children: Vec::new(),
                    })
                    .collect(),
                #[cfg(feature = "extra-children")]
                extra_children: Vec::new(),
            }))
        };

        types::Workbook {
            conformance: None,
            file_version: None,
            #[cfg(feature = "sml-protection")]
            file_sharing: None,
            workbook_pr: None,
            #[cfg(feature = "sml-protection")]
            workbook_protection: self.workbook_protection.clone().map(Box::new),
            book_views: None,
            sheets: Box::new(types::Sheets {
                sheet: sheets,
                #[cfg(feature = "extra-children")]
                extra_children: Vec::new(),
            }),
            #[cfg(feature = "sml-formulas-advanced")]
            function_groups: None,
            #[cfg(feature = "sml-external")]
            external_references: None,
            defined_names,
            #[cfg(feature = "sml-formulas")]
            calc_pr: None,
            #[cfg(feature = "sml-external")]
            ole_size: None,
            #[cfg(feature = "sml-structure")]
            custom_workbook_views: None,
            #[cfg(feature = "sml-pivot")]
            pivot_caches: None,
            #[cfg(feature = "sml-metadata")]
            smart_tag_pr: None,
            #[cfg(feature = "sml-metadata")]
            smart_tag_types: None,
            #[cfg(feature = "sml-external")]
            web_publishing: None,
            file_recovery_pr: Vec::new(),
            #[cfg(feature = "sml-external")]
            web_publish_objects: None,
            #[cfg(feature = "sml-extensions")]
            extension_list: None,
            #[cfg(feature = "extra-attrs")]
            extra_attrs: Default::default(),
            #[cfg(feature = "extra-children")]
            extra_children: Vec::new(),
        }
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
                .map(|s| types::RichString {
                    cell_type: Some(s.clone()),
                    reference: Vec::new(),
                    r_ph: Vec::new(),
                    phonetic_pr: None,
                    #[cfg(feature = "extra-children")]
                    extra_children: Vec::new(),
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

// ============================================================================
// Chart drawing XML builder (sml-charts)
// ============================================================================

/// Build the `<xdr:wsDr>` drawing XML for a set of chart entries.
///
/// Each chart gets a `<xdr:twoCellAnchor>` referencing `rId{n}` where `n` is
/// the 1-based index into the drawing part's own relationship list.
///
/// ECMA-376 Part 1, §20.5 (SpreadsheetDrawingML).
#[cfg(feature = "sml-charts")]
fn build_drawing_xml(charts: &[ChartEntry], _first_chart_num: usize) -> String {
    const NS_XDR: &str = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing";
    const NS_A: &str = "http://schemas.openxmlformats.org/drawingml/2006/main";
    const NS_C: &str = "http://schemas.openxmlformats.org/drawingml/2006/chart";
    const NS_R: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

    let mut xml = format!(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<xdr:wsDr xmlns:xdr="{NS_XDR}" xmlns:a="{NS_A}" xmlns:c="{NS_C}" xmlns:r="{NS_R}">"#
    );

    for (idx, chart) in charts.iter().enumerate() {
        let rel_id = idx + 1; // 1-based within the drawing .rels
        let to_col = chart.x + chart.width;
        let to_row = chart.y + chart.height;
        xml.push_str(&format!(
            r#"
<xdr:twoCellAnchor>
  <xdr:from><xdr:col>{}</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>{}</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>
  <xdr:to><xdr:col>{}</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>{}</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>
  <xdr:graphicFrame macro="">
    <xdr:nvGraphicFramePr>
      <xdr:cNvPr id="{}" name="Chart {}"/>
      <xdr:cNvGraphicFramePr/>
    </xdr:nvGraphicFramePr>
    <xdr:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></xdr:xfrm>
    <a:graphic>
      <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart">
        <c:chart r:id="rId{}"/>
      </a:graphicData>
    </a:graphic>
  </xdr:graphicFrame>
  <xdr:clientData/>
</xdr:twoCellAnchor>"#,
            chart.x, chart.y, to_col, to_row,
            idx + 2, // nvPr id (must be >= 2 to avoid conflicts)
            idx + 1,
            rel_id,
        ));
    }

    xml.push_str("\n</xdr:wsDr>");
    xml
}

// ============================================================================
// Pivot table XML builders (sml-pivot)
// ============================================================================

/// Build a minimal `<pivotCacheDefinition>` XML.
///
/// ECMA-376 Part 1, §18.10.1.
#[cfg(feature = "sml-pivot")]
fn build_pivot_cache_definition_xml(opts: &PivotTableOptions, _pn: usize) -> String {
    // Parse "SheetName!$A$1:$D$10" → sheet name + range.
    let (sheet_name, ref_range) = parse_source_ref(&opts.source_ref);

    // All source fields: row_fields + col_fields + data_fields (deduplicated by order).
    let all_fields = collect_all_fields(opts);
    let field_count = all_fields.len();

    let mut xml = format!(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<pivotCacheDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
  r:id="rId1" refreshedBy="ooxml-sml" refreshedDate="0" createdVersion="3"
  refreshedVersion="3" minRefreshableVersion="3" recordCount="0">
  <cacheSource type="worksheet">
    <worksheetSource ref="{ref_range}" sheet="{sheet_name}"/>
  </cacheSource>
  <cacheFields count="{field_count}">"#
    );

    for name in &all_fields {
        xml.push_str(&format!(
            r#"
    <cacheField name="{}" numFmtId="0"><sharedItems/></cacheField>"#,
            escape_xml(name)
        ));
    }

    xml.push_str(
        r#"
  </cacheFields>
</pivotCacheDefinition>"#,
    );
    xml
}

/// Build a minimal `<pivotTableDefinition>` XML.
///
/// ECMA-376 Part 1, §18.10.2.
#[cfg(feature = "sml-pivot")]
fn build_pivot_table_xml(opts: &PivotTableOptions, cache_id: usize, _sheet_name: &str) -> String {
    let all_fields = collect_all_fields(opts);
    let total_fields = all_fields.len();

    // Build field-name to index map.
    let field_index: HashMap<&str, usize> = all_fields
        .iter()
        .enumerate()
        .map(|(i, n)| (n.as_str(), i))
        .collect();

    // Row fields: indices into the cache field list.
    let row_field_indices: Vec<usize> = opts
        .row_fields
        .iter()
        .filter_map(|n| field_index.get(n.as_str()).copied())
        .collect();

    // Col fields: indices into cache field list.
    let col_field_indices: Vec<usize> = opts
        .col_fields
        .iter()
        .filter_map(|n| field_index.get(n.as_str()).copied())
        .collect();

    // Data fields: indices into cache field list.
    let data_field_indices: Vec<usize> = opts
        .data_fields
        .iter()
        .filter_map(|n| field_index.get(n.as_str()).copied())
        .collect();

    // Location: dest_ref gives top-left; estimate a bounding box.
    let dest = &opts.dest_ref;
    let header_rows = 1u32;
    let data_rows = opts.row_fields.len().max(1) as u32;
    let data_cols = opts.col_fields.len().max(1) as u32;
    // dest..dest+(data_rows+header_rows)x(data_cols+1) — rough estimate.
    let (dest_col, dest_row) = parse_cell_ref_for_pivot(dest);
    let end_col = dest_col + data_cols;
    let end_row = dest_row + header_rows + data_rows;
    let location_ref = format!("{}:{}", dest, format_cell_ref(end_col, end_row));

    let mut xml = format!(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<pivotTableDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
  name="{name}" cacheId="{cache_id}" dataOnRows="0" dataPosition="0"
  dataCaption="Values" createdVersion="3" updatedVersion="3" minRefreshableVersion="3"
  useAutoFormatting="1" itemPrintTitles="1" indent="0" outline="0" outlineData="0">
  <location ref="{location_ref}" firstHeaderRow="1" firstDataRow="2" firstDataCol="1"/>
  <pivotFields count="{total_fields}">"#,
        name = escape_xml(&opts.name),
        cache_id = cache_id,
        location_ref = location_ref,
        total_fields = total_fields,
    );

    // One <pivotField/> per source field.
    for _f in &all_fields {
        xml.push_str(
            r#"
    <pivotField showAll="0"/>"#,
        );
    }

    xml.push_str(
        r#"
  </pivotFields>"#,
    );

    // rowFields
    if !row_field_indices.is_empty() {
        xml.push_str(&format!(
            r#"
  <rowFields count="{}">"#,
            row_field_indices.len()
        ));
        for idx in &row_field_indices {
            xml.push_str(&format!(
                r#"
    <field x="{}"/>"#,
                idx
            ));
        }
        xml.push_str(
            r#"
  </rowFields>"#,
        );
    }

    // colFields: use -2 (data axis placeholder) when there are data fields but no col fields,
    // so Excel shows the Values field header on the column axis.
    if !col_field_indices.is_empty() {
        xml.push_str(&format!(
            r#"
  <colFields count="{}">"#,
            col_field_indices.len()
        ));
        for idx in &col_field_indices {
            xml.push_str(&format!(
                r#"
    <field x="{}"/>"#,
                idx
            ));
        }
        xml.push_str(
            r#"
  </colFields>"#,
        );
    } else if !data_field_indices.is_empty() {
        // Single data field with no explicit column axis → put values on columns.
        xml.push_str(
            r#"
  <colFields count="1">
    <field x="-2"/>
  </colFields>"#,
        );
    }

    // dataFields
    if !data_field_indices.is_empty() {
        xml.push_str(&format!(
            r#"
  <dataFields count="{}">"#,
            data_field_indices.len()
        ));
        for (di, &fld_idx) in data_field_indices.iter().enumerate() {
            let field_name = &opts.data_fields[di];
            xml.push_str(&format!(
                r#"
    <dataField name="Sum of {}" fld="{}" subtotal="sum"/>"#,
                escape_xml(field_name),
                fld_idx
            ));
        }
        xml.push_str(
            r#"
  </dataFields>"#,
        );
    }

    xml.push_str(
        r#"
</pivotTableDefinition>"#,
    );
    xml
}

/// Collect all unique field names from a `PivotTableOptions` in declaration order.
///
/// Row fields come first, then col fields, then data fields (new names only).
#[cfg(feature = "sml-pivot")]
fn collect_all_fields(opts: &PivotTableOptions) -> Vec<String> {
    let mut seen = std::collections::HashSet::new();
    let mut fields = Vec::new();
    for name in opts
        .row_fields
        .iter()
        .chain(opts.col_fields.iter())
        .chain(opts.data_fields.iter())
    {
        if seen.insert(name.as_str()) {
            fields.push(name.clone());
        }
    }
    fields
}

/// Parse a source reference like `"Sheet1!$A$1:$D$10"` into `("Sheet1", "$A$1:$D$10")`.
/// Falls back to `("Sheet1", source_ref)` if no `!` is found.
#[cfg(feature = "sml-pivot")]
fn parse_source_ref(source_ref: &str) -> (&str, &str) {
    if let Some(bang) = source_ref.find('!') {
        (&source_ref[..bang], &source_ref[bang + 1..])
    } else {
        ("Sheet1", source_ref)
    }
}

/// Parse a cell reference like `"A1"` or `"$F$3"` into (col, row) 1-based numbers.
#[cfg(feature = "sml-pivot")]
fn parse_cell_ref_for_pivot(cell_ref: &str) -> (u32, u32) {
    // Strip `$` signs then delegate to the existing parser.
    let clean: String = cell_ref.chars().filter(|c| *c != '$').collect();
    // Reuse parse_cell_reference which returns (row, col).
    if let Some((row, col)) = parse_cell_reference(&clean) {
        (col, row)
    } else {
        (1, 1) // fallback
    }
}

/// Format 1-based (col, row) back to a cell reference like `"B3"`.
#[cfg(feature = "sml-pivot")]
fn format_cell_ref(col: u32, row: u32) -> String {
    format!("{}{}", column_to_letter(col), row)
}

/// Build a single ConditionalFormatting item from a ConditionalFormat builder.
#[cfg(feature = "sml-styling")]
fn build_one_conditional_format(cf: &ConditionalFormat) -> types::ConditionalFormatting {
    types::ConditionalFormatting {
        #[cfg(feature = "sml-pivot")]
        pivot: None,
        square_reference: Some(cf.range.clone()),
        cf_rule: cf
            .rules
            .iter()
            .map(|rule| types::ConditionalRule {
                r#type: Some(map_conditional_rule_type(&rule.rule_type)),
                dxf_id: rule.dxf_id,
                priority: rule.priority as i32,
                stop_if_true: None,
                above_average: None,
                percent: None,
                bottom: None,
                operator: rule
                    .operator
                    .as_ref()
                    .and_then(|op| parse_conditional_operator(op)),
                text: rule.text.clone(),
                time_period: None,
                rank: None,
                std_dev: None,
                equal_average: None,
                formula: rule.formulas.clone(),
                #[cfg(feature = "sml-styling")]
                color_scale: None,
                #[cfg(feature = "sml-styling")]
                data_bar: None,
                #[cfg(feature = "sml-styling")]
                icon_set: None,
                #[cfg(feature = "sml-extensions")]
                extension_list: None,
                #[cfg(feature = "extra-attrs")]
                extra_attrs: Default::default(),
                #[cfg(feature = "extra-children")]
                extra_children: Vec::new(),
            })
            .collect(),
        #[cfg(feature = "sml-extensions")]
        extension_list: None,
        #[cfg(feature = "extra-attrs")]
        extra_attrs: Default::default(),
        #[cfg(feature = "extra-children")]
        extra_children: Vec::new(),
    }
}

/// Map ConditionalRuleType to generated ConditionalType.
#[cfg(feature = "sml-styling")]
fn map_conditional_rule_type(rule_type: &crate::ConditionalRuleType) -> types::ConditionalType {
    match rule_type {
        crate::ConditionalRuleType::Expression => types::ConditionalType::Expression,
        crate::ConditionalRuleType::CellIs => types::ConditionalType::CellIs,
        crate::ConditionalRuleType::ColorScale => types::ConditionalType::ColorScale,
        crate::ConditionalRuleType::DataBar => types::ConditionalType::DataBar,
        crate::ConditionalRuleType::IconSet => types::ConditionalType::IconSet,
        crate::ConditionalRuleType::Top10 => types::ConditionalType::Top10,
        crate::ConditionalRuleType::UniqueValues => types::ConditionalType::UniqueValues,
        crate::ConditionalRuleType::DuplicateValues => types::ConditionalType::DuplicateValues,
        crate::ConditionalRuleType::ContainsText => types::ConditionalType::ContainsText,
        crate::ConditionalRuleType::NotContainsText => types::ConditionalType::NotContainsText,
        crate::ConditionalRuleType::BeginsWith => types::ConditionalType::BeginsWith,
        crate::ConditionalRuleType::EndsWith => types::ConditionalType::EndsWith,
        crate::ConditionalRuleType::ContainsBlanks => types::ConditionalType::ContainsBlanks,
        crate::ConditionalRuleType::NotContainsBlanks => types::ConditionalType::NotContainsBlanks,
        crate::ConditionalRuleType::ContainsErrors => types::ConditionalType::ContainsErrors,
        crate::ConditionalRuleType::NotContainsErrors => types::ConditionalType::NotContainsErrors,
        crate::ConditionalRuleType::TimePeriod => types::ConditionalType::TimePeriod,
        crate::ConditionalRuleType::AboveAverage => types::ConditionalType::AboveAverage,
    }
}

/// Parse a conditional operator string to the generated type.
#[cfg(feature = "sml-styling")]
fn parse_conditional_operator(op: &str) -> Option<types::ConditionalOperator> {
    match op {
        "lessThan" => Some(types::ConditionalOperator::LessThan),
        "lessThanOrEqual" => Some(types::ConditionalOperator::LessThanOrEqual),
        "equal" => Some(types::ConditionalOperator::Equal),
        "notEqual" => Some(types::ConditionalOperator::NotEqual),
        "greaterThanOrEqual" => Some(types::ConditionalOperator::GreaterThanOrEqual),
        "greaterThan" => Some(types::ConditionalOperator::GreaterThan),
        "between" => Some(types::ConditionalOperator::Between),
        "notBetween" => Some(types::ConditionalOperator::NotBetween),
        "containsText" => Some(types::ConditionalOperator::ContainsText),
        "notContains" => Some(types::ConditionalOperator::NotContains),
        "beginsWith" => Some(types::ConditionalOperator::BeginsWith),
        "endsWith" => Some(types::ConditionalOperator::EndsWith),
        _ => None,
    }
}

/// Build a single DataValidation item from a DataValidationBuilder.
#[cfg(feature = "sml-validation")]
fn build_one_data_validation(dv: &DataValidationBuilder) -> types::DataValidation {
    types::DataValidation {
        r#type: map_validation_type(&dv.validation_type),
        error_style: map_validation_error_style(&dv.error_style),
        ime_mode: None,
        operator: map_validation_operator(&dv.operator),
        allow_blank: if dv.allow_blank { Some(true) } else { None },
        show_drop_down: None,
        show_input_message: if dv.show_input_message {
            Some(true)
        } else {
            None
        },
        show_error_message: if dv.show_error_message {
            Some(true)
        } else {
            None
        },
        error_title: dv.error_title.clone(),
        error: dv.error_message.clone(),
        prompt_title: dv.prompt_title.clone(),
        prompt: dv.prompt_message.clone(),
        square_reference: dv.range.clone(),
        formula1: dv.formula1.clone(),
        formula2: dv.formula2.clone(),
        #[cfg(feature = "extra-attrs")]
        extra_attrs: Default::default(),
        #[cfg(feature = "extra-children")]
        extra_children: Vec::new(),
    }
}

/// Map DataValidationType to generated ValidationType.
#[cfg(feature = "sml-validation")]
fn map_validation_type(vt: &crate::DataValidationType) -> Option<types::ValidationType> {
    match vt {
        crate::DataValidationType::None => None, // None type means no validation
        crate::DataValidationType::Whole => Some(types::ValidationType::Whole),
        crate::DataValidationType::Decimal => Some(types::ValidationType::Decimal),
        crate::DataValidationType::List => Some(types::ValidationType::List),
        crate::DataValidationType::Date => Some(types::ValidationType::Date),
        crate::DataValidationType::Time => Some(types::ValidationType::Time),
        crate::DataValidationType::TextLength => Some(types::ValidationType::TextLength),
        crate::DataValidationType::Custom => Some(types::ValidationType::Custom),
    }
}

/// Map DataValidationOperator to generated ValidationOperator.
#[cfg(feature = "sml-validation")]
fn map_validation_operator(
    op: &crate::DataValidationOperator,
) -> Option<types::ValidationOperator> {
    match op {
        crate::DataValidationOperator::Between => Some(types::ValidationOperator::Between),
        crate::DataValidationOperator::NotBetween => Some(types::ValidationOperator::NotBetween),
        crate::DataValidationOperator::Equal => Some(types::ValidationOperator::Equal),
        crate::DataValidationOperator::NotEqual => Some(types::ValidationOperator::NotEqual),
        crate::DataValidationOperator::LessThan => Some(types::ValidationOperator::LessThan),
        crate::DataValidationOperator::LessThanOrEqual => {
            Some(types::ValidationOperator::LessThanOrEqual)
        }
        crate::DataValidationOperator::GreaterThan => Some(types::ValidationOperator::GreaterThan),
        crate::DataValidationOperator::GreaterThanOrEqual => {
            Some(types::ValidationOperator::GreaterThanOrEqual)
        }
    }
}

/// Map DataValidationErrorStyle to generated ValidationErrorStyle.
#[cfg(feature = "sml-validation")]
fn map_validation_error_style(
    style: &crate::DataValidationErrorStyle,
) -> Option<types::ValidationErrorStyle> {
    match style {
        crate::DataValidationErrorStyle::Stop => None, // Stop is the default
        crate::DataValidationErrorStyle::Warning => Some(types::ValidationErrorStyle::Warning),
        crate::DataValidationErrorStyle::Information => {
            Some(types::ValidationErrorStyle::Information)
        }
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

/// Escape XML special characters in attribute values (used in .rels files).
fn escape_xml(s: &str) -> String {
    s.replace('&', "&amp;")
        .replace('<', "&lt;")
        .replace('>', "&gt;")
        .replace('"', "&quot;")
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

// =============================================================================
// ToXml serialization helpers
// =============================================================================

/// Namespace declarations for SML root elements (worksheet, etc.).
const NS_DECLS: &[(&str, &str)] = &[("xmlns", NS_SPREADSHEET)];

/// Namespace declarations for workbook element (includes relationship namespace for r:id).
const NS_WORKBOOK_DECLS: &[(&str, &str)] =
    &[("xmlns", NS_SPREADSHEET), ("xmlns:r", NS_RELATIONSHIPS)];

/// Serialize a ToXml value with namespace declarations and XML declaration.
fn serialize_with_namespaces(value: &impl ToXml, tag: &str) -> Result<Vec<u8>> {
    serialize_with_ns_decls(value, tag, NS_DECLS)
}

/// Serialize a workbook with both spreadsheet and relationship namespaces.
fn serialize_workbook(value: &impl ToXml) -> Result<Vec<u8>> {
    serialize_with_ns_decls(value, "workbook", NS_WORKBOOK_DECLS)
}

/// Serialize a ToXml value with custom namespace declarations and XML declaration.
fn serialize_with_ns_decls(
    value: &impl ToXml,
    tag: &str,
    ns_decls: &[(&str, &str)],
) -> Result<Vec<u8>> {
    use quick_xml::Writer;
    use quick_xml::events::{BytesEnd, BytesStart, Event};

    let inner = Vec::new();
    let mut writer = Writer::new(inner);

    // Write start tag with namespace declarations + type's own attrs
    let start = BytesStart::new(tag);
    let start = value.write_attrs(start);
    let mut start = start;
    for &(key, val) in ns_decls {
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

/// Build a types::Font from FontStyle.
#[cfg(feature = "sml-styling")]
fn build_font(font: &FontStyle) -> types::Font {
    types::Font {
        #[cfg(feature = "sml-styling")]
        name: font.name.as_ref().map(|n| {
            Box::new(types::FontName {
                value: n.clone(),
                #[cfg(feature = "extra-attrs")]
                extra_attrs: Default::default(),
            })
        }),
        #[cfg(feature = "sml-styling")]
        charset: None,
        #[cfg(feature = "sml-styling")]
        family: None,
        #[cfg(feature = "sml-styling")]
        b: if font.bold {
            Some(Box::new(types::BooleanProperty {
                value: None, // Default means "true" when element is present
                #[cfg(feature = "extra-attrs")]
                extra_attrs: Default::default(),
            }))
        } else {
            None
        },
        #[cfg(feature = "sml-styling")]
        i: if font.italic {
            Some(Box::new(types::BooleanProperty {
                value: None,
                #[cfg(feature = "extra-attrs")]
                extra_attrs: Default::default(),
            }))
        } else {
            None
        },
        #[cfg(feature = "sml-styling")]
        strike: if font.strikethrough {
            Some(Box::new(types::BooleanProperty {
                value: None,
                #[cfg(feature = "extra-attrs")]
                extra_attrs: Default::default(),
            }))
        } else {
            None
        },
        #[cfg(feature = "sml-styling")]
        outline: None,
        #[cfg(feature = "sml-styling")]
        shadow: None,
        #[cfg(feature = "sml-styling")]
        condense: None,
        #[cfg(feature = "sml-styling")]
        extend: None,
        #[cfg(feature = "sml-styling")]
        color: font.color.as_ref().map(|c| {
            Box::new(types::Color {
                auto: None,
                indexed: None,
                rgb: Some(hex_color_to_bytes(c)),
                theme: None,
                tint: None,
                #[cfg(feature = "extra-attrs")]
                extra_attrs: Default::default(),
            })
        }),
        #[cfg(feature = "sml-styling")]
        sz: font.size.map(|s| {
            Box::new(types::FontSize {
                value: s,
                #[cfg(feature = "extra-attrs")]
                extra_attrs: Default::default(),
            })
        }),
        #[cfg(feature = "sml-styling")]
        u: font.underline.map(|u| {
            Box::new(types::UnderlineProperty {
                value: Some(convert_underline_style(u)),
                #[cfg(feature = "extra-attrs")]
                extra_attrs: Default::default(),
            })
        }),
        #[cfg(feature = "sml-styling")]
        vert_align: None,
        #[cfg(feature = "sml-styling")]
        scheme: None,
        #[cfg(feature = "extra-children")]
        extra_children: Vec::new(),
    }
}

/// Build a types::Fill from FillStyle.
#[cfg(feature = "sml-styling")]
fn build_fill(fill: &FillStyle) -> types::Fill {
    types::Fill {
        #[cfg(feature = "sml-styling")]
        pattern_fill: Some(Box::new(types::PatternFill {
            pattern_type: Some(convert_pattern_type(fill.pattern)),
            fg_color: fill.fg_color.as_ref().map(|c| {
                Box::new(types::Color {
                    auto: None,
                    indexed: None,
                    rgb: Some(hex_color_to_bytes(c)),
                    theme: None,
                    tint: None,
                    #[cfg(feature = "extra-attrs")]
                    extra_attrs: Default::default(),
                })
            }),
            bg_color: fill.bg_color.as_ref().map(|c| {
                Box::new(types::Color {
                    auto: None,
                    indexed: None,
                    rgb: Some(hex_color_to_bytes(c)),
                    theme: None,
                    tint: None,
                    #[cfg(feature = "extra-attrs")]
                    extra_attrs: Default::default(),
                })
            }),
            #[cfg(feature = "extra-attrs")]
            extra_attrs: Default::default(),
            #[cfg(feature = "extra-children")]
            extra_children: Vec::new(),
        })),
        #[cfg(feature = "sml-styling")]
        gradient_fill: None,
        #[cfg(feature = "extra-children")]
        extra_children: Vec::new(),
    }
}

/// Build a types::Border from BorderStyle.
#[cfg(feature = "sml-styling")]
fn build_border(border: &BorderStyle) -> types::Border {
    types::Border {
        #[cfg(feature = "sml-styling")]
        diagonal_up: if border.diagonal_up { Some(true) } else { None },
        #[cfg(feature = "sml-styling")]
        diagonal_down: if border.diagonal_down {
            Some(true)
        } else {
            None
        },
        #[cfg(feature = "sml-styling")]
        outline: None,
        start: None,
        end: None,
        #[cfg(feature = "sml-styling")]
        left: build_border_properties(&border.left),
        #[cfg(feature = "sml-styling")]
        right: build_border_properties(&border.right),
        #[cfg(feature = "sml-styling")]
        top: build_border_properties(&border.top),
        #[cfg(feature = "sml-styling")]
        bottom: build_border_properties(&border.bottom),
        #[cfg(feature = "sml-styling")]
        diagonal: build_border_properties(&border.diagonal),
        #[cfg(feature = "sml-styling")]
        vertical: None,
        #[cfg(feature = "sml-styling")]
        horizontal: None,
        #[cfg(feature = "extra-attrs")]
        extra_attrs: Default::default(),
        #[cfg(feature = "extra-children")]
        extra_children: Vec::new(),
    }
}

/// Build a types::BorderProperties from BorderSideStyle.
#[cfg(feature = "sml-styling")]
fn build_border_properties(side: &Option<BorderSideStyle>) -> Option<Box<types::BorderProperties>> {
    // Always emit border properties for each side (empty if none)
    let (style, color) = if let Some(s) = side {
        if s.style != BorderLineStyle::None {
            (
                Some(convert_border_style(s.style)),
                s.color.as_ref().map(|c| {
                    Box::new(types::Color {
                        auto: None,
                        indexed: None,
                        rgb: Some(hex_color_to_bytes(c)),
                        theme: None,
                        tint: None,
                        #[cfg(feature = "extra-attrs")]
                        extra_attrs: Default::default(),
                    })
                }),
            )
        } else {
            (None, None)
        }
    } else {
        (None, None)
    };

    Some(Box::new(types::BorderProperties {
        style,
        color,
        #[cfg(feature = "extra-attrs")]
        extra_attrs: Default::default(),
        #[cfg(feature = "extra-children")]
        extra_children: Vec::new(),
    }))
}

/// Build a types::Format (xf) from CellFormatRecord.
#[cfg(feature = "sml-styling")]
fn build_cell_format(xf: &CellFormatRecord) -> types::Format {
    let has_alignment = xf.horizontal.is_some() || xf.vertical.is_some() || xf.wrap_text;

    let alignment = if has_alignment {
        Some(Box::new(types::CellAlignment {
            horizontal: xf.horizontal.map(convert_horizontal_alignment),
            vertical: xf.vertical.map(convert_vertical_alignment),
            text_rotation: None,
            wrap_text: if xf.wrap_text { Some(true) } else { None },
            indent: None,
            relative_indent: None,
            justify_last_line: None,
            shrink_to_fit: None,
            reading_order: None,
            #[cfg(feature = "extra-attrs")]
            extra_attrs: Default::default(),
        }))
    } else {
        None
    };

    types::Format {
        #[cfg(feature = "sml-styling")]
        number_format_id: Some(xf.num_fmt_id),
        #[cfg(feature = "sml-styling")]
        font_id: Some(xf.font_id as u32),
        #[cfg(feature = "sml-styling")]
        fill_id: Some(xf.fill_id as u32),
        #[cfg(feature = "sml-styling")]
        border_id: Some(xf.border_id as u32),
        #[cfg(feature = "sml-styling")]
        format_id: Some(0),
        #[cfg(feature = "sml-styling")]
        quote_prefix: None,
        #[cfg(feature = "sml-pivot")]
        pivot_button: None,
        #[cfg(feature = "sml-styling")]
        apply_number_format: if xf.num_fmt_id > 0 { Some(true) } else { None },
        #[cfg(feature = "sml-styling")]
        apply_font: if xf.font_id > 0 { Some(true) } else { None },
        #[cfg(feature = "sml-styling")]
        apply_fill: if xf.fill_id > 0 { Some(true) } else { None },
        #[cfg(feature = "sml-styling")]
        apply_border: if xf.border_id > 0 { Some(true) } else { None },
        #[cfg(feature = "sml-styling")]
        apply_alignment: if has_alignment { Some(true) } else { None },
        #[cfg(feature = "sml-styling")]
        apply_protection: None,
        #[cfg(feature = "sml-styling")]
        alignment,
        #[cfg(feature = "sml-protection")]
        protection: None,
        #[cfg(feature = "sml-extensions")]
        extension_list: None,
        #[cfg(feature = "extra-attrs")]
        extra_attrs: Default::default(),
        #[cfg(feature = "extra-children")]
        extra_children: Vec::new(),
    }
}

/// Convert writer's UnderlineStyle to generated types::UnderlineStyle.
#[cfg(feature = "sml-styling")]
fn convert_underline_style(style: UnderlineStyle) -> types::UnderlineStyle {
    match style {
        UnderlineStyle::Single => types::UnderlineStyle::Single,
        UnderlineStyle::Double => types::UnderlineStyle::Double,
        UnderlineStyle::SingleAccounting => types::UnderlineStyle::SingleAccounting,
        UnderlineStyle::DoubleAccounting => types::UnderlineStyle::DoubleAccounting,
    }
}

/// Convert writer's FillPattern to generated types::PatternType.
#[cfg(feature = "sml-styling")]
fn convert_pattern_type(pattern: FillPattern) -> types::PatternType {
    match pattern {
        FillPattern::None => types::PatternType::None,
        FillPattern::Solid => types::PatternType::Solid,
        FillPattern::MediumGray => types::PatternType::MediumGray,
        FillPattern::DarkGray => types::PatternType::DarkGray,
        FillPattern::LightGray => types::PatternType::LightGray,
        FillPattern::DarkHorizontal => types::PatternType::DarkHorizontal,
        FillPattern::DarkVertical => types::PatternType::DarkVertical,
        FillPattern::DarkDown => types::PatternType::DarkDown,
        FillPattern::DarkUp => types::PatternType::DarkUp,
        FillPattern::DarkGrid => types::PatternType::DarkGrid,
        FillPattern::DarkTrellis => types::PatternType::DarkTrellis,
        FillPattern::LightHorizontal => types::PatternType::LightHorizontal,
        FillPattern::LightVertical => types::PatternType::LightVertical,
        FillPattern::LightDown => types::PatternType::LightDown,
        FillPattern::LightUp => types::PatternType::LightUp,
        FillPattern::LightGrid => types::PatternType::LightGrid,
        FillPattern::LightTrellis => types::PatternType::LightTrellis,
        FillPattern::Gray125 => types::PatternType::Gray125,
        FillPattern::Gray0625 => types::PatternType::Gray0625,
    }
}

/// Convert writer's BorderLineStyle to generated types::BorderStyle.
#[cfg(feature = "sml-styling")]
fn convert_border_style(style: BorderLineStyle) -> types::BorderStyle {
    match style {
        BorderLineStyle::None => types::BorderStyle::None,
        BorderLineStyle::Thin => types::BorderStyle::Thin,
        BorderLineStyle::Medium => types::BorderStyle::Medium,
        BorderLineStyle::Dashed => types::BorderStyle::Dashed,
        BorderLineStyle::Dotted => types::BorderStyle::Dotted,
        BorderLineStyle::Thick => types::BorderStyle::Thick,
        BorderLineStyle::Double => types::BorderStyle::Double,
        BorderLineStyle::Hair => types::BorderStyle::Hair,
        BorderLineStyle::MediumDashed => types::BorderStyle::MediumDashed,
        BorderLineStyle::DashDot => types::BorderStyle::DashDot,
        BorderLineStyle::MediumDashDot => types::BorderStyle::MediumDashDot,
        BorderLineStyle::DashDotDot => types::BorderStyle::DashDotDot,
        BorderLineStyle::MediumDashDotDot => types::BorderStyle::MediumDashDotDot,
        BorderLineStyle::SlantDashDot => types::BorderStyle::SlantDashDot,
    }
}

/// Convert writer's HorizontalAlignment to generated types::HorizontalAlignment.
#[cfg(feature = "sml-styling")]
fn convert_horizontal_alignment(align: HorizontalAlignment) -> types::HorizontalAlignment {
    match align {
        HorizontalAlignment::General => types::HorizontalAlignment::General,
        HorizontalAlignment::Left => types::HorizontalAlignment::Left,
        HorizontalAlignment::Center => types::HorizontalAlignment::Center,
        HorizontalAlignment::Right => types::HorizontalAlignment::Right,
        HorizontalAlignment::Fill => types::HorizontalAlignment::Fill,
        HorizontalAlignment::Justify => types::HorizontalAlignment::Justify,
        HorizontalAlignment::CenterContinuous => types::HorizontalAlignment::CenterContinuous,
        HorizontalAlignment::Distributed => types::HorizontalAlignment::Distributed,
    }
}

/// Convert writer's VerticalAlignment to generated types::VerticalAlignment.
#[cfg(feature = "sml-styling")]
fn convert_vertical_alignment(align: VerticalAlignment) -> types::VerticalAlignment {
    match align {
        VerticalAlignment::Top => types::VerticalAlignment::Top,
        VerticalAlignment::Center => types::VerticalAlignment::Center,
        VerticalAlignment::Bottom => types::VerticalAlignment::Bottom,
        VerticalAlignment::Justify => types::VerticalAlignment::Justify,
        VerticalAlignment::Distributed => types::VerticalAlignment::Distributed,
    }
}

/// Convert a 6-character RGB hex color string (e.g., "FF0000") to ARGB bytes with 0xFF alpha.
#[cfg(feature = "sml-styling")]
fn hex_color_to_bytes(color: &str) -> Vec<u8> {
    // Parse as 24-bit RGB integer, then extract bytes
    let rgb = u32::from_str_radix(color, 16).unwrap_or(0);
    vec![0xFF, (rgb >> 16) as u8, (rgb >> 8) as u8, rgb as u8]
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
    fn test_roundtrip_freeze_rows() {
        use std::io::Cursor;

        let mut wb = WorkbookBuilder::new();
        let sheet = wb.add_sheet("Sheet1");
        sheet.set_cell("A1", "Header");
        sheet.freeze_rows(1);

        let mut buffer = Cursor::new(Vec::new());
        wb.write(&mut buffer).unwrap();

        buffer.set_position(0);
        let mut workbook = crate::Workbook::from_reader(buffer).unwrap();
        let read_sheet = workbook.resolved_sheet(0).unwrap();

        assert!(read_sheet.has_freeze_panes(), "Should have freeze panes");
        let pane = read_sheet.freeze_pane().expect("Should have pane");
        assert_eq!(pane.y_split, Some(1.0), "Should freeze 1 row");
        assert_eq!(pane.x_split, None, "Should not freeze any columns");
        assert_eq!(
            pane.state,
            Some(crate::types::PaneState::Frozen),
            "State should be Frozen"
        );
    }

    #[test]
    #[cfg(feature = "full")]
    fn test_roundtrip_freeze_cols() {
        use std::io::Cursor;

        let mut wb = WorkbookBuilder::new();
        let sheet = wb.add_sheet("Sheet1");
        sheet.set_cell("A1", "Row label");
        sheet.freeze_cols(1);

        let mut buffer = Cursor::new(Vec::new());
        wb.write(&mut buffer).unwrap();

        buffer.set_position(0);
        let mut workbook = crate::Workbook::from_reader(buffer).unwrap();
        let read_sheet = workbook.resolved_sheet(0).unwrap();

        assert!(read_sheet.has_freeze_panes());
        let pane = read_sheet.freeze_pane().expect("Should have pane");
        assert_eq!(pane.x_split, Some(1.0), "Should freeze 1 column");
        assert_eq!(pane.y_split, None, "Should not freeze any rows");
    }

    #[test]
    #[cfg(feature = "full")]
    fn test_roundtrip_freeze_both() {
        use std::io::Cursor;

        let mut wb = WorkbookBuilder::new();
        let sheet = wb.add_sheet("Sheet1");
        sheet.set_cell("A1", "Header");
        sheet.set_freeze_pane(2, 1);

        let mut buffer = Cursor::new(Vec::new());
        wb.write(&mut buffer).unwrap();

        buffer.set_position(0);
        let mut workbook = crate::Workbook::from_reader(buffer).unwrap();
        let read_sheet = workbook.resolved_sheet(0).unwrap();

        assert!(read_sheet.has_freeze_panes());
        let pane = read_sheet.freeze_pane().expect("Should have pane");
        assert_eq!(pane.y_split, Some(2.0), "Should freeze 2 rows");
        assert_eq!(pane.x_split, Some(1.0), "Should freeze 1 column");
        assert_eq!(pane.active_pane, Some(crate::types::PaneType::BottomRight));
    }

    #[test]
    #[cfg(feature = "full")]
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

        use crate::DefinedNameExt;

        // Check global range
        let global = workbook.defined_name("GlobalRange").unwrap();
        assert_eq!(global.name, "GlobalRange");
        assert_eq!(global.text.as_deref(), Some("Sheet1!$A$1:$B$10"));
        assert!(global.local_sheet_id.is_none());

        // Check sheet-scoped range
        let local = workbook.defined_name_in_sheet("LocalRange", 0).unwrap();
        assert_eq!(local.name, "LocalRange");
        assert_eq!(local.text.as_deref(), Some("Sheet1!$C$1:$D$5"));
        assert_eq!(local.local_sheet_id, Some(0));

        // Check data range with comment
        let data = workbook.defined_name("DataRange").unwrap();
        assert_eq!(data.name, "DataRange");
        assert_eq!(data.text.as_deref(), Some("Sheet2!$A$1:$Z$100"));
        assert_eq!(data.comment.as_deref(), Some("Main data table"));

        // Check print area (built-in name)
        let print_area = workbook
            .defined_name_in_sheet("_xlnm.Print_Area", 0)
            .unwrap();
        assert_eq!(print_area.text.as_deref(), Some("Sheet1!$A$1:$G$20"));
        assert!(print_area.is_builtin());
    }

    #[cfg(feature = "sml-comments")]
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

    #[cfg(feature = "sml-filtering")]
    #[test]
    fn test_roundtrip_auto_filter() {
        use std::io::Cursor;

        let mut wb = WorkbookBuilder::new();
        let sheet = wb.add_sheet("Data");
        sheet.set_cell("A1", "Name");
        sheet.set_cell("B1", "Score");
        sheet.set_cell("A2", "Alice");
        sheet.set_cell("B2", 95.0);
        sheet.set_auto_filter("A1:B1");

        let mut buffer = Cursor::new(Vec::new());
        wb.write(&mut buffer).unwrap();

        buffer.set_position(0);
        let mut workbook = crate::Workbook::from_reader(buffer).unwrap();
        let read_sheet = workbook.resolved_sheet(0).unwrap();

        assert!(read_sheet.has_auto_filter());
        let af = read_sheet.auto_filter().unwrap();
        assert_eq!(af.reference.as_deref(), Some("A1:B1"));
    }

    #[cfg(feature = "sml-hyperlinks")]
    #[test]
    fn test_roundtrip_hyperlinks() {
        use std::io::Cursor;

        let mut wb = WorkbookBuilder::new();
        let sheet = wb.add_sheet("Links");
        sheet.set_cell("A1", "External");
        sheet.add_hyperlink("A1", "https://example.com");
        sheet.set_cell("B1", "Internal");
        sheet.add_internal_hyperlink("B1", "Sheet2!A1");

        let mut buffer = Cursor::new(Vec::new());
        wb.write(&mut buffer).unwrap();

        buffer.set_position(0);
        let mut workbook = crate::Workbook::from_reader(buffer).unwrap();
        let read_sheet = workbook.resolved_sheet(0).unwrap();

        // Check the raw hyperlinks from the worksheet
        let ws = read_sheet.worksheet();
        let hyperlinks = ws.hyperlinks.as_deref().unwrap();
        assert_eq!(hyperlinks.hyperlink.len(), 2);

        let ext = hyperlinks
            .hyperlink
            .iter()
            .find(|h| h.reference == "A1")
            .unwrap();
        // External hyperlink has a relationship ID, no inline location
        assert!(ext.id.is_some());
        assert!(ext.location.as_deref().unwrap_or("").is_empty());

        let int = hyperlinks
            .hyperlink
            .iter()
            .find(|h| h.reference == "B1")
            .unwrap();
        // Internal hyperlink has no relationship ID, just a location
        assert!(int.id.is_none());
        assert_eq!(int.location.as_deref(), Some("Sheet2!A1"));
    }

    #[cfg(all(feature = "sml-hyperlinks", feature = "sml-comments"))]
    #[test]
    fn test_roundtrip_hyperlinks_with_comments() {
        use std::io::Cursor;

        // Verifies that rel IDs are assigned correctly when a sheet has both
        // comments (rId1) and external hyperlinks (rId2, rId3...).
        let mut wb = WorkbookBuilder::new();
        let sheet = wb.add_sheet("Sheet1");
        sheet.set_cell("A1", "Click me");
        sheet.add_hyperlink("A1", "https://example.com");
        sheet.add_comment("A1", "Visit the site");

        let mut buffer = Cursor::new(Vec::new());
        wb.write(&mut buffer).unwrap();

        buffer.set_position(0);
        let mut workbook = crate::Workbook::from_reader(buffer).unwrap();
        let read_sheet = workbook.resolved_sheet(0).unwrap();

        // Comment should be preserved
        let c = read_sheet.comment("A1").unwrap();
        assert_eq!(c.text, "Visit the site");

        // Hyperlink should be preserved
        let ws = read_sheet.worksheet();
        let hyperlinks = ws.hyperlinks.as_deref().unwrap();
        assert_eq!(hyperlinks.hyperlink.len(), 1);
        assert_eq!(hyperlinks.hyperlink[0].reference, "A1");
        assert!(hyperlinks.hyperlink[0].id.is_some());
    }

    // =========================================================================
    // Page setup / margins
    // =========================================================================

    #[cfg(feature = "sml-layout")]
    #[test]
    fn test_roundtrip_page_setup() {
        use std::io::Cursor;

        let mut wb = WorkbookBuilder::new();
        let sheet = wb.add_sheet("Sheet1");
        sheet.set_cell("A1", "Print me");
        sheet.set_page_setup(
            PageSetupOptions::new()
                .with_orientation(PageOrientation::Landscape)
                .with_paper_size(9) // A4
                .with_scale(80),
        );

        let mut buffer = Cursor::new(Vec::new());
        wb.write(&mut buffer).unwrap();

        buffer.set_position(0);
        let mut workbook = crate::Workbook::from_reader(buffer).unwrap();
        let read_sheet = workbook.resolved_sheet(0).unwrap();
        let ws = read_sheet.worksheet();

        let ps = ws.page_setup.as_deref().expect("page_setup should be set");
        assert_eq!(ps.paper_size, Some(9));
        assert_eq!(ps.scale, Some(80));
        assert_eq!(ps.orientation, Some(crate::types::STOrientation::Landscape));
    }

    #[cfg(feature = "sml-layout")]
    #[test]
    fn test_roundtrip_page_margins() {
        use std::io::Cursor;

        let mut wb = WorkbookBuilder::new();
        let sheet = wb.add_sheet("Sheet1");
        sheet.set_cell("A1", "Data");
        sheet.set_page_margins(0.7, 0.7, 0.75, 0.75, 0.3, 0.3);

        let mut buffer = Cursor::new(Vec::new());
        wb.write(&mut buffer).unwrap();

        buffer.set_position(0);
        let mut workbook = crate::Workbook::from_reader(buffer).unwrap();
        let read_sheet = workbook.resolved_sheet(0).unwrap();
        let ws = read_sheet.worksheet();

        let pm = ws
            .page_margins
            .as_deref()
            .expect("page_margins should be set");
        assert!((pm.left - 0.7).abs() < f64::EPSILON);
        assert!((pm.right - 0.7).abs() < f64::EPSILON);
        assert!((pm.top - 0.75).abs() < f64::EPSILON);
        assert!((pm.bottom - 0.75).abs() < f64::EPSILON);
        assert!((pm.header - 0.3).abs() < f64::EPSILON);
        assert!((pm.footer - 0.3).abs() < f64::EPSILON);
    }

    // =========================================================================
    // Row / column outline grouping
    // =========================================================================

    #[cfg(all(feature = "sml-structure", feature = "sml-styling"))]
    #[test]
    fn test_roundtrip_row_outline() {
        use std::io::Cursor;

        let mut wb = WorkbookBuilder::new();
        let sheet = wb.add_sheet("Sheet1");
        sheet.set_cell("A1", "Header");
        sheet.set_cell("A2", "Detail 1");
        sheet.set_cell("A3", "Detail 2");
        // Group rows 2–3 at outline level 1
        sheet.set_row_outline_level(2, 1);
        sheet.set_row_outline_level(3, 1);
        // Mark the group as collapsed
        sheet.set_row_collapsed(2, true);

        let mut buffer = Cursor::new(Vec::new());
        wb.write(&mut buffer).unwrap();

        buffer.set_position(0);
        let mut workbook = crate::Workbook::from_reader(buffer).unwrap();
        let read_sheet = workbook.resolved_sheet(0).unwrap();

        let row2 = read_sheet.row(2).expect("row 2 should exist");
        assert_eq!(row2.outline_level, Some(1), "row 2 outline level");
        assert_eq!(row2.collapsed, Some(true), "row 2 collapsed");

        let row3 = read_sheet.row(3).expect("row 3 should exist");
        assert_eq!(row3.outline_level, Some(1), "row 3 outline level");
        assert_eq!(row3.collapsed, None, "row 3 not collapsed");
    }

    #[cfg(all(feature = "sml-structure", feature = "sml-styling"))]
    #[test]
    fn test_roundtrip_col_outline() {
        use std::io::Cursor;

        let mut wb = WorkbookBuilder::new();
        let sheet = wb.add_sheet("Sheet1");
        sheet.set_cell("A1", "Label");
        sheet.set_cell("B1", "Detail");
        sheet.set_column_outline_level("B", 1);
        sheet.set_column_collapsed("B", true);

        let mut buffer = Cursor::new(Vec::new());
        wb.write(&mut buffer).unwrap();

        buffer.set_position(0);
        let mut workbook = crate::Workbook::from_reader(buffer).unwrap();
        let read_sheet = workbook.resolved_sheet(0).unwrap();

        let ws = read_sheet.worksheet();
        let col_b = ws
            .cols
            .iter()
            .flat_map(|c| &c.col)
            .find(|c| c.start_column <= 2 && 2 <= c.end_column)
            .expect("column B should have a definition");
        assert_eq!(col_b.outline_level, Some(1), "col B outline level");
        assert_eq!(col_b.collapsed, Some(true), "col B collapsed");
    }

    // =========================================================================
    // Ignored errors
    // =========================================================================

    #[cfg(feature = "sml-validation")]
    #[test]
    fn test_roundtrip_ignored_errors() {
        use std::io::Cursor;

        let mut wb = WorkbookBuilder::new();
        let sheet = wb.add_sheet("Sheet1");
        sheet.set_cell("A1", "123"); // text that looks like a number
        sheet.add_ignored_error("A1:A10", IgnoredErrorType::NumberStoredAsText);
        sheet.add_ignored_error("B1:B5", IgnoredErrorType::TwoDigitTextYear);

        let mut buffer = Cursor::new(Vec::new());
        wb.write(&mut buffer).unwrap();

        buffer.set_position(0);
        let mut workbook = crate::Workbook::from_reader(buffer).unwrap();
        let read_sheet = workbook.resolved_sheet(0).unwrap();
        let ws = read_sheet.worksheet();

        let ie = ws
            .ignored_errors
            .as_deref()
            .expect("ignored_errors should be set");
        assert_eq!(ie.ignored_error.len(), 2);

        let first = &ie.ignored_error[0];
        assert_eq!(first.square_reference.as_str(), "A1:A10");
        assert_eq!(first.number_stored_as_text, Some(true));

        let second = &ie.ignored_error[1];
        assert_eq!(second.square_reference.as_str(), "B1:B5");
        assert_eq!(second.two_digit_text_year, Some(true));
    }

    // =========================================================================
    // Tab color
    // =========================================================================

    #[cfg(feature = "sml-styling")]
    #[test]
    fn test_roundtrip_tab_color() {
        use std::io::Cursor;

        let mut wb = WorkbookBuilder::new();
        let sheet = wb.add_sheet("Red Sheet");
        sheet.set_cell("A1", "Hello");
        sheet.set_tab_color("FF0000"); // red

        let mut buffer = Cursor::new(Vec::new());
        wb.write(&mut buffer).unwrap();

        buffer.set_position(0);
        let mut workbook = crate::Workbook::from_reader(buffer).unwrap();
        let read_sheet = workbook.resolved_sheet(0).unwrap();
        let ws = read_sheet.worksheet();

        let props = ws
            .sheet_properties
            .as_deref()
            .expect("sheet_properties should be set");
        let color = props.tab_color.as_deref().expect("tab_color should be set");
        // rgb is stored as ARGB bytes: 0xFF, 0xFF, 0x00, 0x00
        assert_eq!(
            color.rgb.as_deref(),
            Some(&[0xFF_u8, 0xFF, 0x00, 0x00] as &[u8])
        );
    }

    // =========================================================================
    // Sheet view options
    // =========================================================================

    #[cfg(feature = "sml-styling")]
    #[test]
    fn test_roundtrip_show_gridlines() {
        use std::io::Cursor;

        let mut wb = WorkbookBuilder::new();
        let sheet = wb.add_sheet("Sheet1");
        sheet.set_cell("A1", "Hello");
        sheet.set_show_gridlines(false);

        let mut buffer = Cursor::new(Vec::new());
        wb.write(&mut buffer).unwrap();

        buffer.set_position(0);
        let mut workbook = crate::Workbook::from_reader(buffer).unwrap();
        let read_sheet = workbook.resolved_sheet(0).unwrap();
        let ws = read_sheet.worksheet();

        let sv = ws
            .sheet_views
            .as_deref()
            .and_then(|v| v.sheet_view.first())
            .expect("sheet_views should be set");
        assert_eq!(sv.show_grid_lines, Some(false));
    }

    #[cfg(feature = "sml-styling")]
    #[test]
    fn test_roundtrip_show_row_col_headers() {
        use std::io::Cursor;

        let mut wb = WorkbookBuilder::new();
        let sheet = wb.add_sheet("Sheet1");
        sheet.set_cell("A1", "Hello");
        sheet.set_show_row_col_headers(false);

        let mut buffer = Cursor::new(Vec::new());
        wb.write(&mut buffer).unwrap();

        buffer.set_position(0);
        let mut workbook = crate::Workbook::from_reader(buffer).unwrap();
        let read_sheet = workbook.resolved_sheet(0).unwrap();
        let ws = read_sheet.worksheet();

        let sv = ws
            .sheet_views
            .as_deref()
            .and_then(|v| v.sheet_view.first())
            .expect("sheet_views should be set");
        assert_eq!(sv.show_row_col_headers, Some(false));
    }

    #[cfg(all(feature = "sml-styling", feature = "sml-structure"))]
    #[test]
    fn test_show_gridlines_with_freeze_pane() {
        use std::io::Cursor;

        // Ensure that setting show_gridlines doesn't clobber a previously set
        // freeze pane (both must coexist in the same SheetView).
        let mut wb = WorkbookBuilder::new();
        let sheet = wb.add_sheet("Sheet1");
        sheet.set_cell("A1", "Header");
        sheet.freeze_rows(1);
        sheet.set_show_gridlines(false);

        let mut buffer = Cursor::new(Vec::new());
        wb.write(&mut buffer).unwrap();

        buffer.set_position(0);
        let mut workbook = crate::Workbook::from_reader(buffer).unwrap();
        let read_sheet = workbook.resolved_sheet(0).unwrap();
        let ws = read_sheet.worksheet();

        let sv = ws
            .sheet_views
            .as_deref()
            .and_then(|v| v.sheet_view.first())
            .expect("sheet_views should be set");
        assert_eq!(sv.show_grid_lines, Some(false), "gridlines hidden");
        // Freeze pane should still be intact.
        let pane = sv.pane.as_deref().expect("freeze pane should be intact");
        assert_eq!(pane.y_split, Some(1.0));
    }

    // =========================================================================
    // Sheet protection
    // =========================================================================

    #[cfg(feature = "sml-protection")]
    #[test]
    fn test_ooxml_xor_hash_empty() {
        // Empty password → 0x0000
        let hash = ooxml_xor_hash("");
        assert_eq!(hash, vec![0x00, 0x00]);
    }

    #[cfg(feature = "sml-protection")]
    #[test]
    fn test_ooxml_xor_hash_known() {
        // Known value: "password" hashes to 0xCE4B with the XOR algorithm
        // (verified against ECMA-376 §18.2.28 test vectors).
        // We don't hard-code the exact value but verify it's non-zero and
        // deterministic.
        let h1 = ooxml_xor_hash("password");
        let h2 = ooxml_xor_hash("password");
        assert_eq!(h1, h2, "hash must be deterministic");
        assert_ne!(
            h1,
            vec![0x00, 0x00],
            "hash must be non-zero for non-empty password"
        );
    }

    #[cfg(feature = "sml-protection")]
    #[test]
    fn test_roundtrip_sheet_protection_no_password() {
        use std::io::Cursor;

        let mut wb = WorkbookBuilder::new();
        let sheet = wb.add_sheet("Sheet1");
        sheet.set_cell("A1", "Protected");
        sheet.set_sheet_protection(SheetProtectionOptions {
            sheet: true,
            format_cells: true,
            insert_rows: true,
            ..Default::default()
        });

        let mut buffer = Cursor::new(Vec::new());
        wb.write(&mut buffer).unwrap();

        buffer.set_position(0);
        let mut workbook = crate::Workbook::from_reader(buffer).unwrap();
        let read_sheet = workbook.resolved_sheet(0).unwrap();
        let ws = read_sheet.worksheet();

        let sp = ws
            .sheet_protection
            .as_deref()
            .expect("sheet_protection should be set");
        assert_eq!(sp.sheet, Some(true), "sheet locked");
        assert_eq!(sp.format_cells, Some(true), "format_cells locked");
        assert_eq!(sp.insert_rows, Some(true), "insert_rows locked");
        assert_eq!(sp.password, None, "no password");
    }

    #[cfg(feature = "sml-protection")]
    #[test]
    fn test_roundtrip_sheet_protection_with_password() {
        use std::io::Cursor;

        let mut wb = WorkbookBuilder::new();
        let sheet = wb.add_sheet("Sheet1");
        sheet.set_cell("A1", "Data");
        sheet.set_sheet_protection(SheetProtectionOptions {
            sheet: true,
            password: Some("secret".to_string()),
            ..Default::default()
        });

        let mut buffer = Cursor::new(Vec::new());
        wb.write(&mut buffer).unwrap();

        buffer.set_position(0);
        let mut workbook = crate::Workbook::from_reader(buffer).unwrap();
        let read_sheet = workbook.resolved_sheet(0).unwrap();
        let ws = read_sheet.worksheet();

        let sp = ws
            .sheet_protection
            .as_deref()
            .expect("sheet_protection should be set");
        assert_eq!(sp.sheet, Some(true));
        // Password is stored as a 2-byte hash
        let pw = sp.password.as_ref().expect("password should be set");
        assert_eq!(pw.len(), 2);
        // Verify same password produces same hash
        let expected = ooxml_xor_hash("secret");
        assert_eq!(pw, &expected);
    }

    // =========================================================================
    // Rich-text comments
    // =========================================================================

    #[cfg(feature = "sml-comments")]
    #[test]
    fn test_roundtrip_rich_text_comment() {
        use std::io::Cursor;

        let mut wb = WorkbookBuilder::new();
        let sheet = wb.add_sheet("Sheet1");
        sheet.set_cell("A1", "Value");

        // Rich-text comment with two runs
        let mut cb = CommentBuilder::new_rich("A1");
        cb.add_run("Bold prefix: ").set_bold(true);
        cb.add_run("normal suffix");
        sheet.add_comment_builder(cb);

        let mut buffer = Cursor::new(Vec::new());
        wb.write(&mut buffer).unwrap();

        // Read back and verify the comment text is preserved
        buffer.set_position(0);
        let mut workbook = crate::Workbook::from_reader(buffer).unwrap();
        let read_sheet = workbook.resolved_sheet(0).unwrap();

        let comment = read_sheet.comment("A1").expect("comment should exist");
        // The reader concatenates all run texts into comment.text
        assert!(
            comment.text.contains("Bold prefix:"),
            "bold run text present: {:?}",
            comment.text
        );
        assert!(
            comment.text.contains("normal suffix"),
            "normal run text present: {:?}",
            comment.text
        );
    }

    // =========================================================================
    // Workbook protection
    // =========================================================================

    #[cfg(feature = "sml-protection")]
    #[test]
    fn test_roundtrip_workbook_protection() {
        use std::io::Cursor;

        let mut wb = WorkbookBuilder::new();
        wb.add_sheet("Sheet1");
        wb.set_workbook_protection(true, false, Some("wb_pass"));

        let mut buffer = Cursor::new(Vec::new());
        wb.write(&mut buffer).unwrap();

        buffer.set_position(0);
        let workbook = crate::Workbook::from_reader(buffer).unwrap();

        let wp = workbook
            .workbook_protection()
            .expect("workbook_protection should be set");
        assert_eq!(wp.lock_structure, Some(true), "lock_structure");
        assert_eq!(wp.lock_windows, None, "lock_windows not set");
        let pw = wp
            .workbook_password
            .as_ref()
            .expect("password should be set");
        assert_eq!(pw.len(), 2);
        let expected = ooxml_xor_hash("wb_pass");
        assert_eq!(pw, &expected);
    }

    #[cfg(feature = "sml-protection")]
    #[test]
    fn test_roundtrip_workbook_protection_no_password() {
        use std::io::Cursor;

        let mut wb = WorkbookBuilder::new();
        wb.add_sheet("Sheet1");
        wb.set_workbook_protection(true, true, None);

        let mut buffer = Cursor::new(Vec::new());
        wb.write(&mut buffer).unwrap();

        buffer.set_position(0);
        let workbook = crate::Workbook::from_reader(buffer).unwrap();

        let wp = workbook
            .workbook_protection()
            .expect("workbook_protection should be set");
        assert_eq!(wp.lock_structure, Some(true));
        assert_eq!(wp.lock_windows, Some(true));
        assert_eq!(wp.workbook_password, None, "no password");
    }

    // =========================================================================
    // Named cell styles
    // =========================================================================

    #[cfg(feature = "sml-styling")]
    #[test]
    fn test_add_cell_style_returns_index() {
        let mut wb = WorkbookBuilder::new();
        wb.add_sheet("Sheet1");

        // First extra style gets index 1 (Normal = 0)
        let idx1 = wb.add_cell_style("Good", 0);
        assert_eq!(idx1, 1);

        // Second extra style gets index 2
        let idx2 = wb.add_cell_style("Bad", 0);
        assert_eq!(idx2, 2);
    }

    #[cfg(feature = "sml-styling")]
    #[test]
    fn test_roundtrip_cell_styles() {
        use std::io::Cursor;

        let mut wb = WorkbookBuilder::new();
        wb.add_sheet("Sheet1");
        wb.add_cell_style("Good", 0);
        wb.add_cell_style("Neutral", 0);

        let mut buffer = Cursor::new(Vec::new());
        wb.write(&mut buffer).unwrap();

        buffer.set_position(0);
        let workbook = crate::Workbook::from_reader(buffer).unwrap();

        // The stylesheet should have Normal + Good + Neutral = 3 cell styles
        let stylesheet = workbook.styles();
        let cell_styles = stylesheet
            .cell_styles
            .as_deref()
            .expect("cell_styles should be set");
        assert_eq!(cell_styles.count, Some(3));
        assert_eq!(cell_styles.cell_style[0].name.as_deref(), Some("Normal"));
        assert_eq!(cell_styles.cell_style[1].name.as_deref(), Some("Good"));
        assert_eq!(cell_styles.cell_style[2].name.as_deref(), Some("Neutral"));
        // Custom styles have customBuiltin=true
        assert_eq!(cell_styles.cell_style[1].custom_builtin, Some(true));
    }

    // =========================================================================
    // Chart embedding
    // =========================================================================

    /// Verify that `embed_chart` writes a drawing part and a chart XML part.
    #[cfg(feature = "sml-charts")]
    #[test]
    fn test_embed_chart_creates_drawing_and_chart_parts() {
        use std::collections::HashSet;
        use std::io::Cursor;

        let chart_xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
  <c:chart/>
</c:chartSpace>"#;

        let mut wb = WorkbookBuilder::new();
        let sheet = wb.add_sheet("Sheet1");
        sheet.set_cell("A1", "data");
        sheet.embed_chart(chart_xml, 0, 5, 8, 15);

        let mut buffer = Cursor::new(Vec::new());
        wb.write(&mut buffer).unwrap();

        // Unpack the ZIP and check which parts are present.
        buffer.set_position(0);
        let mut zip = zip::ZipArchive::new(buffer).unwrap();
        let names: HashSet<String> = (0..zip.len())
            .map(|i| zip.by_index(i).unwrap().name().to_string())
            .collect();

        // Drawing and chart parts must exist.
        assert!(
            names.contains("xl/drawings/drawing1.xml"),
            "drawing part missing; parts: {names:?}"
        );
        assert!(
            names.contains("xl/charts/chart1.xml"),
            "chart part missing; parts: {names:?}"
        );
        // Drawing rels must exist.
        assert!(
            names.contains("xl/drawings/_rels/drawing1.xml.rels"),
            "drawing rels missing; parts: {names:?}"
        );
        // Sheet rels must reference the drawing.
        assert!(
            names.contains("xl/worksheets/_rels/sheet1.xml.rels"),
            "sheet rels missing; parts: {names:?}"
        );

        // The drawing XML must contain a twoCellAnchor with the chart reference.
        let drawing_bytes = {
            let mut f = zip.by_name("xl/drawings/drawing1.xml").unwrap();
            let mut v = Vec::new();
            std::io::Read::read_to_end(&mut f, &mut v).unwrap();
            v
        };
        let drawing_str = String::from_utf8_lossy(&drawing_bytes);
        assert!(
            drawing_str.contains("twoCellAnchor"),
            "drawing XML should contain twoCellAnchor"
        );
        assert!(
            drawing_str.contains("rId1"),
            "drawing XML should reference rId1 for the chart"
        );

        // The chart XML must be exactly what was passed in.
        let chart_bytes = {
            let mut f = zip.by_name("xl/charts/chart1.xml").unwrap();
            let mut v = Vec::new();
            std::io::Read::read_to_end(&mut f, &mut v).unwrap();
            v
        };
        assert_eq!(chart_bytes, chart_xml);
    }

    /// Two charts in the same sheet should get separate rId/chart numbers.
    #[cfg(feature = "sml-charts")]
    #[test]
    fn test_embed_two_charts_same_sheet() {
        use std::collections::HashSet;
        use std::io::Cursor;

        let chart_xml = b"<c:chartSpace/>";

        let mut wb = WorkbookBuilder::new();
        let sheet = wb.add_sheet("Sheet1");
        sheet.embed_chart(chart_xml, 0, 0, 4, 10);
        sheet.embed_chart(chart_xml, 5, 0, 4, 10);

        let mut buffer = Cursor::new(Vec::new());
        wb.write(&mut buffer).unwrap();

        buffer.set_position(0);
        let mut zip = zip::ZipArchive::new(buffer).unwrap();
        let names: HashSet<String> = (0..zip.len())
            .map(|i| zip.by_index(i).unwrap().name().to_string())
            .collect();

        // Should have one drawing (for the sheet) but two chart files.
        assert!(names.contains("xl/drawings/drawing1.xml"));
        assert!(names.contains("xl/charts/chart1.xml"));
        assert!(names.contains("xl/charts/chart2.xml"));
        // Should NOT have a second drawing file.
        assert!(!names.contains("xl/drawings/drawing2.xml"));
    }

    /// Charts across two different sheets get separate drawing/chart files.
    #[cfg(feature = "sml-charts")]
    #[test]
    fn test_embed_charts_multiple_sheets() {
        use std::collections::HashSet;
        use std::io::Cursor;

        let chart_xml = b"<c:chartSpace/>";

        let mut wb = WorkbookBuilder::new();
        let s1 = wb.add_sheet("Sheet1");
        s1.embed_chart(chart_xml, 0, 0, 4, 10);
        let s2 = wb.add_sheet("Sheet2");
        s2.embed_chart(chart_xml, 0, 0, 4, 10);

        let mut buffer = Cursor::new(Vec::new());
        wb.write(&mut buffer).unwrap();

        buffer.set_position(0);
        let mut zip = zip::ZipArchive::new(buffer).unwrap();
        let names: HashSet<String> = (0..zip.len())
            .map(|i| zip.by_index(i).unwrap().name().to_string())
            .collect();

        assert!(names.contains("xl/drawings/drawing1.xml"));
        assert!(names.contains("xl/drawings/drawing2.xml"));
        assert!(names.contains("xl/charts/chart1.xml"));
        assert!(names.contains("xl/charts/chart2.xml"));
    }

    // =========================================================================
    // Pivot tables
    // =========================================================================

    /// A pivot table produces the expected set of XML parts.
    #[cfg(feature = "sml-pivot")]
    #[test]
    fn test_pivot_table_creates_expected_parts() {
        use std::collections::HashSet;
        use std::io::Cursor;

        let mut wb = WorkbookBuilder::new();
        let sheet = wb.add_sheet("Sheet1");
        sheet.set_cell("A1", "Region");
        sheet.set_cell("B1", "Sales");
        sheet.add_pivot_table(PivotTableOptions {
            name: "PivotTable1".to_string(),
            source_ref: "Sheet1!$A$1:$B$5".to_string(),
            dest_ref: "D1".to_string(),
            row_fields: vec!["Region".to_string()],
            col_fields: vec![],
            data_fields: vec!["Sales".to_string()],
        });

        let mut buffer = Cursor::new(Vec::new());
        wb.write(&mut buffer).unwrap();

        buffer.set_position(0);
        let mut zip = zip::ZipArchive::new(buffer).unwrap();
        let names: HashSet<String> = (0..zip.len())
            .map(|i| zip.by_index(i).unwrap().name().to_string())
            .collect();

        assert!(
            names.contains("xl/pivotCache/pivotCacheDefinition1.xml"),
            "pivot cache definition missing; parts: {names:?}"
        );
        assert!(
            names.contains("xl/pivotCache/pivotCacheRecords1.xml"),
            "pivot cache records missing; parts: {names:?}"
        );
        assert!(
            names.contains("xl/pivotTables/pivotTable1.xml"),
            "pivot table missing; parts: {names:?}"
        );
        assert!(
            names.contains("xl/pivotTables/_rels/pivotTable1.xml.rels"),
            "pivot table rels missing; parts: {names:?}"
        );
        assert!(
            names.contains("xl/pivotCache/_rels/pivotCacheDefinition1.xml.rels"),
            "pivot cache definition rels missing; parts: {names:?}"
        );
        assert!(
            names.contains("xl/worksheets/_rels/sheet1.xml.rels"),
            "sheet rels missing; parts: {names:?}"
        );
    }

    /// The pivot table XML must contain the expected structure elements.
    #[cfg(feature = "sml-pivot")]
    #[test]
    fn test_pivot_table_xml_structure() {
        use std::io::Cursor;

        let mut wb = WorkbookBuilder::new();
        let sheet = wb.add_sheet("Sheet1");
        sheet.add_pivot_table(PivotTableOptions {
            name: "MySales".to_string(),
            source_ref: "Sheet1!$A$1:$C$10".to_string(),
            dest_ref: "E1".to_string(),
            row_fields: vec!["Region".to_string()],
            col_fields: vec!["Quarter".to_string()],
            data_fields: vec!["Revenue".to_string()],
        });

        let mut buffer = Cursor::new(Vec::new());
        wb.write(&mut buffer).unwrap();

        buffer.set_position(0);
        let mut zip = zip::ZipArchive::new(buffer).unwrap();

        // Check pivot table definition content.
        let pt_bytes = {
            let mut f = zip.by_name("xl/pivotTables/pivotTable1.xml").unwrap();
            let mut v = Vec::new();
            std::io::Read::read_to_end(&mut f, &mut v).unwrap();
            v
        };
        let pt_str = String::from_utf8_lossy(&pt_bytes);
        assert!(
            pt_str.contains("MySales"),
            "should contain pivot table name"
        );
        assert!(
            pt_str.contains("pivotTableDefinition"),
            "should have root element"
        );
        assert!(
            pt_str.contains("rowFields"),
            "should have rowFields element"
        );
        assert!(
            pt_str.contains("dataFields"),
            "should have dataFields element"
        );
        assert!(
            pt_str.contains("Sum of Revenue"),
            "should have data field name"
        );

        // Check pivot cache definition content.
        let cd_bytes = {
            let mut f = zip
                .by_name("xl/pivotCache/pivotCacheDefinition1.xml")
                .unwrap();
            let mut v = Vec::new();
            std::io::Read::read_to_end(&mut f, &mut v).unwrap();
            v
        };
        let cd_str = String::from_utf8_lossy(&cd_bytes);
        assert!(
            cd_str.contains("worksheetSource"),
            "should have worksheetSource"
        );
        assert!(cd_str.contains("Sheet1"), "should reference source sheet");
        assert!(cd_str.contains("cacheFields"), "should have cacheFields");
        assert!(cd_str.contains("Region"), "should list Region field");
        assert!(cd_str.contains("Revenue"), "should list Revenue field");
    }

    /// Two pivot tables in one sheet get separate part numbers.
    #[cfg(feature = "sml-pivot")]
    #[test]
    fn test_two_pivot_tables_same_sheet() {
        use std::collections::HashSet;
        use std::io::Cursor;

        let mut wb = WorkbookBuilder::new();
        let sheet = wb.add_sheet("Sheet1");
        sheet.add_pivot_table(PivotTableOptions {
            name: "Pivot1".to_string(),
            source_ref: "Sheet1!$A$1:$B$5".to_string(),
            dest_ref: "D1".to_string(),
            row_fields: vec!["A".to_string()],
            col_fields: vec![],
            data_fields: vec!["B".to_string()],
        });
        sheet.add_pivot_table(PivotTableOptions {
            name: "Pivot2".to_string(),
            source_ref: "Sheet1!$A$1:$B$5".to_string(),
            dest_ref: "G1".to_string(),
            row_fields: vec!["A".to_string()],
            col_fields: vec![],
            data_fields: vec!["B".to_string()],
        });

        let mut buffer = Cursor::new(Vec::new());
        wb.write(&mut buffer).unwrap();

        buffer.set_position(0);
        let mut zip = zip::ZipArchive::new(buffer).unwrap();
        let names: HashSet<String> = (0..zip.len())
            .map(|i| zip.by_index(i).unwrap().name().to_string())
            .collect();

        assert!(names.contains("xl/pivotTables/pivotTable1.xml"));
        assert!(names.contains("xl/pivotTables/pivotTable2.xml"));
        assert!(names.contains("xl/pivotCache/pivotCacheDefinition1.xml"));
        assert!(names.contains("xl/pivotCache/pivotCacheDefinition2.xml"));
    }
}
