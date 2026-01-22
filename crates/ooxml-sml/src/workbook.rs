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
    /// Defined names (named ranges).
    defined_names: Vec<DefinedName>,
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

        // Parse workbook.xml to get sheet list and defined names
        let workbook_xml = package.read_part(&workbook_path)?;
        let sheet_info = parse_workbook_sheets(&workbook_xml)?;
        let defined_names = parse_defined_names(&workbook_xml)?;

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
            defined_names,
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

    /// Get all defined names (named ranges).
    pub fn defined_names(&self) -> &[DefinedName] {
        &self.defined_names
    }

    /// Get a defined name by its name.
    ///
    /// For names with sheet scope, use `defined_name_in_sheet` instead.
    pub fn defined_name(&self, name: &str) -> Option<&DefinedName> {
        self.defined_names
            .iter()
            .find(|d| d.name.eq_ignore_ascii_case(name) && d.local_sheet_id.is_none())
    }

    /// Get a defined name by its name within a specific sheet scope.
    pub fn defined_name_in_sheet(&self, name: &str, sheet_index: u32) -> Option<&DefinedName> {
        self.defined_names
            .iter()
            .find(|d| d.name.eq_ignore_ascii_case(name) && d.local_sheet_id == Some(sheet_index))
    }

    /// Get all global defined names (workbook scope).
    pub fn global_defined_names(&self) -> impl Iterator<Item = &DefinedName> {
        self.defined_names
            .iter()
            .filter(|d| d.local_sheet_id.is_none())
    }

    /// Get all defined names scoped to a specific sheet.
    pub fn sheet_defined_names(&self, sheet_index: u32) -> impl Iterator<Item = &DefinedName> {
        self.defined_names
            .iter()
            .filter(move |d| d.local_sheet_id == Some(sheet_index))
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

/// Conditional formatting rules for a range.
///
/// ECMA-376 Part 1, Section 18.3.1.18 (conditionalFormatting).
#[derive(Debug, Clone)]
pub struct ConditionalFormatting {
    /// Cell range(s) this formatting applies to (e.g., "A1:B10" or "A1:B10 C1:D10").
    pub ranges: String,
    /// Formatting rules (in priority order).
    pub rules: Vec<ConditionalRule>,
}

impl ConditionalFormatting {
    /// Get the cell ranges as individual strings.
    pub fn range_list(&self) -> Vec<&str> {
        self.ranges.split_whitespace().collect()
    }

    /// Check if a cell reference is within the conditional formatting range.
    pub fn contains_ref(&self, reference: &str) -> bool {
        for range in self.range_list() {
            if range_contains_ref(range, reference) {
                return true;
            }
        }
        false
    }
}

/// A single conditional formatting rule.
///
/// ECMA-376 Part 1, Section 18.3.1.10 (cfRule).
#[derive(Debug, Clone)]
pub struct ConditionalRule {
    /// Rule type (e.g., "cellIs", "colorScale", "dataBar").
    pub rule_type: ConditionalRuleType,
    /// Priority (lower number = higher priority).
    pub priority: u32,
    /// Differential formatting ID (reference to dxf in styles).
    pub dxf_id: Option<u32>,
    /// Operator for comparison rules.
    pub operator: Option<String>,
    /// Formula(s) used in the rule.
    pub formulas: Vec<String>,
    /// Stop if this rule matches.
    pub stop_if_true: bool,
    /// For aboveAverage type: whether to match above average.
    pub above_average: Option<bool>,
    /// For aboveAverage type: whether to include equal values.
    pub equal_average: Option<bool>,
    /// For top10 type: rank value.
    pub rank: Option<u32>,
    /// For top10 type: whether it's top (true) or bottom (false).
    pub top: Option<bool>,
    /// For top10 type: whether rank is a percentage.
    pub percent: Option<bool>,
    /// Text value for containsText, beginsWith, endsWith rules.
    pub text: Option<String>,
    /// Time period for timePeriod rules.
    pub time_period: Option<String>,
    /// Color scale configuration.
    pub color_scale: Option<ColorScale>,
    /// Data bar configuration.
    pub data_bar: Option<DataBar>,
    /// Icon set configuration.
    pub icon_set: Option<IconSet>,
}

impl ConditionalRule {
    /// Check if this is a cell value comparison rule.
    pub fn is_cell_is(&self) -> bool {
        matches!(self.rule_type, ConditionalRuleType::CellIs)
    }

    /// Check if this is a color scale rule.
    pub fn is_color_scale(&self) -> bool {
        matches!(self.rule_type, ConditionalRuleType::ColorScale)
    }

    /// Check if this is a data bar rule.
    pub fn is_data_bar(&self) -> bool {
        matches!(self.rule_type, ConditionalRuleType::DataBar)
    }

    /// Check if this is an icon set rule.
    pub fn is_icon_set(&self) -> bool {
        matches!(self.rule_type, ConditionalRuleType::IconSet)
    }
}

/// Type of conditional formatting rule.
///
/// ECMA-376 Part 1, Section 18.18.12 (ST_CfType).
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ConditionalRuleType {
    /// Expression-based rule.
    Expression,
    /// Cell value comparison.
    CellIs,
    /// Color scale gradient.
    ColorScale,
    /// Data bar visualization.
    DataBar,
    /// Icon set.
    IconSet,
    /// Top N values.
    Top10,
    /// Unique values.
    UniqueValues,
    /// Duplicate values.
    DuplicateValues,
    /// Contains specified text.
    ContainsText,
    /// Does not contain specified text.
    NotContainsText,
    /// Begins with specified text.
    BeginsWith,
    /// Ends with specified text.
    EndsWith,
    /// Contains blanks.
    ContainsBlanks,
    /// Does not contain blanks.
    NotContainsBlanks,
    /// Contains errors.
    ContainsErrors,
    /// Does not contain errors.
    NotContainsErrors,
    /// Time period comparison.
    TimePeriod,
    /// Above or below average.
    AboveAverage,
}

impl ConditionalRuleType {
    /// Parse from the cfRule type attribute string.
    pub fn parse(s: &str) -> Option<Self> {
        match s {
            "expression" => Some(Self::Expression),
            "cellIs" => Some(Self::CellIs),
            "colorScale" => Some(Self::ColorScale),
            "dataBar" => Some(Self::DataBar),
            "iconSet" => Some(Self::IconSet),
            "top10" => Some(Self::Top10),
            "uniqueValues" => Some(Self::UniqueValues),
            "duplicateValues" => Some(Self::DuplicateValues),
            "containsText" => Some(Self::ContainsText),
            "notContainsText" => Some(Self::NotContainsText),
            "beginsWith" => Some(Self::BeginsWith),
            "endsWith" => Some(Self::EndsWith),
            "containsBlanks" => Some(Self::ContainsBlanks),
            "notContainsBlanks" => Some(Self::NotContainsBlanks),
            "containsErrors" => Some(Self::ContainsErrors),
            "notContainsErrors" => Some(Self::NotContainsErrors),
            "timePeriod" => Some(Self::TimePeriod),
            "aboveAverage" => Some(Self::AboveAverage),
            _ => None,
        }
    }
}

/// Color scale for conditional formatting.
#[derive(Debug, Clone)]
pub struct ColorScale {
    /// Color values (2 or 3 for two-color or three-color scale).
    pub colors: Vec<ColorScaleValue>,
}

/// A single color value in a color scale.
#[derive(Debug, Clone)]
pub struct ColorScaleValue {
    /// Value type (min, max, num, percent, percentile, formula).
    pub value_type: String,
    /// Value (for num, percent, percentile, formula types).
    pub value: Option<String>,
    /// Color in ARGB format.
    pub color: Option<String>,
}

/// Data bar configuration for conditional formatting.
#[derive(Debug, Clone)]
pub struct DataBar {
    /// Minimum value type.
    pub min_type: String,
    /// Minimum value (optional).
    pub min_value: Option<String>,
    /// Maximum value type.
    pub max_type: String,
    /// Maximum value (optional).
    pub max_value: Option<String>,
    /// Bar color in ARGB format.
    pub color: Option<String>,
    /// Show value in cell.
    pub show_value: bool,
}

/// Icon set configuration for conditional formatting.
#[derive(Debug, Clone)]
pub struct IconSet {
    /// Icon set name (e.g., "3Arrows", "4Rating", "5Quarters").
    pub icon_set: String,
    /// Show icon only (no cell value).
    pub show_value: bool,
    /// Reverse icon order.
    pub reverse: bool,
    /// Threshold values for icons.
    pub values: Vec<IconSetValue>,
}

/// A threshold value in an icon set.
#[derive(Debug, Clone)]
pub struct IconSetValue {
    /// Value type (percent, num, percentile, formula).
    pub value_type: String,
    /// Threshold value.
    pub value: Option<String>,
}

/// Data validation rule for a range of cells.
///
/// ECMA-376 Part 1, Section 18.3.1.32 (dataValidation).
#[derive(Debug, Clone)]
pub struct DataValidation {
    /// Cell range(s) this validation applies to.
    pub ranges: String,
    /// Validation type.
    pub validation_type: DataValidationType,
    /// Comparison operator.
    pub operator: DataValidationOperator,
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
    pub error_style: DataValidationErrorStyle,
    /// Error title.
    pub error_title: Option<String>,
    /// Error message.
    pub error_message: Option<String>,
    /// Input prompt title.
    pub prompt_title: Option<String>,
    /// Input prompt message.
    pub prompt_message: Option<String>,
}

impl DataValidation {
    /// Get the cell ranges as individual strings.
    pub fn range_list(&self) -> Vec<&str> {
        self.ranges.split_whitespace().collect()
    }

    /// Check if a cell reference is within the validation range.
    pub fn contains_ref(&self, reference: &str) -> bool {
        for range in self.range_list() {
            if range_contains_ref(range, reference) {
                return true;
            }
        }
        false
    }

    /// Check if this is a list validation (dropdown).
    pub fn is_list(&self) -> bool {
        matches!(self.validation_type, DataValidationType::List)
    }
}

/// Type of data validation.
///
/// ECMA-376 Part 1, Section 18.18.21 (ST_DataValidationType).
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub enum DataValidationType {
    /// No validation.
    #[default]
    None,
    /// Whole number.
    Whole,
    /// Decimal number.
    Decimal,
    /// List of values (dropdown).
    List,
    /// Date.
    Date,
    /// Time.
    Time,
    /// Text length.
    TextLength,
    /// Custom formula.
    Custom,
}

impl DataValidationType {
    /// Parse from the dataValidation type attribute.
    fn parse(s: &str) -> Self {
        match s {
            "none" => Self::None,
            "whole" => Self::Whole,
            "decimal" => Self::Decimal,
            "list" => Self::List,
            "date" => Self::Date,
            "time" => Self::Time,
            "textLength" => Self::TextLength,
            "custom" => Self::Custom,
            _ => Self::None,
        }
    }
}

/// Comparison operator for data validation.
///
/// ECMA-376 Part 1, Section 18.18.22 (ST_DataValidationOperator).
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub enum DataValidationOperator {
    /// Between formula1 and formula2.
    #[default]
    Between,
    /// Not between formula1 and formula2.
    NotBetween,
    /// Equal to formula1.
    Equal,
    /// Not equal to formula1.
    NotEqual,
    /// Less than formula1.
    LessThan,
    /// Less than or equal to formula1.
    LessThanOrEqual,
    /// Greater than formula1.
    GreaterThan,
    /// Greater than or equal to formula1.
    GreaterThanOrEqual,
}

impl DataValidationOperator {
    /// Parse from the dataValidation operator attribute.
    fn parse(s: &str) -> Self {
        match s {
            "between" => Self::Between,
            "notBetween" => Self::NotBetween,
            "equal" => Self::Equal,
            "notEqual" => Self::NotEqual,
            "lessThan" => Self::LessThan,
            "lessThanOrEqual" => Self::LessThanOrEqual,
            "greaterThan" => Self::GreaterThan,
            "greaterThanOrEqual" => Self::GreaterThanOrEqual,
            _ => Self::Between,
        }
    }
}

/// Error alert style for data validation.
///
/// ECMA-376 Part 1, Section 18.18.23 (ST_DataValidationErrorStyle).
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub enum DataValidationErrorStyle {
    /// Stop: Prevents the user from entering invalid data.
    #[default]
    Stop,
    /// Warning: Warns the user but allows invalid data.
    Warning,
    /// Information: Informs the user but allows invalid data.
    Information,
}

impl DataValidationErrorStyle {
    /// Parse from the dataValidation errorStyle attribute.
    fn parse(s: &str) -> Self {
        match s {
            "stop" => Self::Stop,
            "warning" => Self::Warning,
            "information" => Self::Information,
            _ => Self::Stop,
        }
    }
}

/// Check if a cell reference is within a range.
fn range_contains_ref(range: &str, reference: &str) -> bool {
    let (cell_col, cell_row) = match parse_cell_reference(reference) {
        Some(r) => r,
        None => return false,
    };

    if let Some((start, end)) = range.split_once(':') {
        // Range like "A1:B10"
        let (start_col, start_row) = match parse_cell_reference(start) {
            Some(r) => r,
            None => return false,
        };
        let (end_col, end_row) = match parse_cell_reference(end) {
            Some(r) => r,
            None => return false,
        };
        cell_row >= start_row && cell_row <= end_row && cell_col >= start_col && cell_col <= end_col
    } else {
        // Single cell like "A1"
        if let Some((range_col, range_row)) = parse_cell_reference(range) {
            cell_row == range_row && cell_col == range_col
        } else {
            false
        }
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
    conditional_formats: Vec<ConditionalFormatting>,
    data_validations: Vec<DataValidation>,
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

    /// Get all conditional formatting rules in this sheet.
    pub fn conditional_formats(&self) -> &[ConditionalFormatting] {
        &self.conditional_formats
    }

    /// Get conditional formatting rules that apply to a specific cell.
    pub fn conditional_formats_for(&self, reference: &str) -> Vec<&ConditionalFormatting> {
        self.conditional_formats
            .iter()
            .filter(|cf| cf.contains_ref(reference))
            .collect()
    }

    /// Check if a cell has any conditional formatting.
    pub fn has_conditional_formatting(&self, reference: &str) -> bool {
        self.conditional_formats
            .iter()
            .any(|cf| cf.contains_ref(reference))
    }

    /// Get all data validation rules in this sheet.
    pub fn data_validations(&self) -> &[DataValidation] {
        &self.data_validations
    }

    /// Get data validation for a specific cell.
    pub fn data_validation(&self, reference: &str) -> Option<&DataValidation> {
        self.data_validations
            .iter()
            .find(|dv| dv.contains_ref(reference))
    }

    /// Check if a cell has data validation.
    pub fn has_data_validation(&self, reference: &str) -> bool {
        self.data_validation(reference).is_some()
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

impl Stylesheet {
    /// Get a number format code by ID.
    ///
    /// Looks up custom formats first, then falls back to built-in formats.
    pub fn format_code(&self, id: u32) -> Option<String> {
        // Check custom formats first
        if let Some(fmt) = self.number_formats.iter().find(|f| f.id == id) {
            return Some(fmt.code.clone());
        }

        // Fall back to built-in formats
        builtin_format_code(id).map(|s| s.to_string())
    }

    /// Check if a format ID represents a date/time format.
    pub fn is_date_format(&self, id: u32) -> bool {
        if let Some(code) = self.format_code(id) {
            is_date_format_code(&code)
        } else {
            // Check built-in date format IDs (14-22, 45-47)
            matches!(id, 14..=22 | 45..=47)
        }
    }
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

impl NumberFormat {
    /// Check if this format represents a date/time format.
    ///
    /// Date formats contain patterns like y, m, d, h, s but not in contexts
    /// like [Red] color codes or escaped characters.
    pub fn is_date_format(&self) -> bool {
        is_date_format_code(&self.code)
    }
}

/// A named range (defined name) in the workbook.
///
/// ECMA-376 Part 1, Section 18.2.5 (definedName).
/// Named ranges can be global (workbook scope) or local (sheet scope).
#[derive(Debug, Clone)]
pub struct DefinedName {
    /// The name of the defined range (e.g., "MyRange", "_xlnm.Print_Area").
    pub name: String,
    /// The formula or reference (e.g., "Sheet1!$A$1:$B$10").
    pub reference: String,
    /// Optional sheet index if this name is scoped to a specific sheet.
    /// If None, the name is global (workbook scope).
    pub local_sheet_id: Option<u32>,
    /// Optional comment/description.
    pub comment: Option<String>,
    /// Whether this is a hidden name.
    pub hidden: bool,
}

impl DefinedName {
    /// Check if this is a built-in Excel name (prefixed with "_xlnm.").
    ///
    /// Built-in names include:
    /// - _xlnm.Print_Area
    /// - _xlnm.Print_Titles
    /// - _xlnm._FilterDatabase
    /// - _xlnm.Criteria
    /// - _xlnm.Extract
    pub fn is_builtin(&self) -> bool {
        self.name.starts_with("_xlnm.")
    }

    /// Get the scope of this defined name.
    pub fn scope(&self) -> DefinedNameScope {
        match self.local_sheet_id {
            Some(id) => DefinedNameScope::Sheet(id),
            None => DefinedNameScope::Workbook,
        }
    }
}

/// The scope of a defined name.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum DefinedNameScope {
    /// Global workbook scope.
    Workbook,
    /// Local to a specific sheet (by index).
    Sheet(u32),
}

/// Get the format code for a built-in format ID.
///
/// Excel has built-in formats with IDs 0-49. Custom formats start at 164.
/// Reference: ECMA-376 Part 1, Section 18.8.30.
pub fn builtin_format_code(id: u32) -> Option<&'static str> {
    match id {
        0 => Some("General"),
        1 => Some("0"),
        2 => Some("0.00"),
        3 => Some("#,##0"),
        4 => Some("#,##0.00"),
        9 => Some("0%"),
        10 => Some("0.00%"),
        11 => Some("0.00E+00"),
        12 => Some("# ?/?"),
        13 => Some("# ??/??"),
        14 => Some("mm-dd-yy"),
        15 => Some("d-mmm-yy"),
        16 => Some("d-mmm"),
        17 => Some("mmm-yy"),
        18 => Some("h:mm AM/PM"),
        19 => Some("h:mm:ss AM/PM"),
        20 => Some("h:mm"),
        21 => Some("h:mm:ss"),
        22 => Some("m/d/yy h:mm"),
        37 => Some("#,##0 ;(#,##0)"),
        38 => Some("#,##0 ;[Red](#,##0)"),
        39 => Some("#,##0.00;(#,##0.00)"),
        40 => Some("#,##0.00;[Red](#,##0.00)"),
        45 => Some("mm:ss"),
        46 => Some("[h]:mm:ss"),
        47 => Some("mmss.0"),
        48 => Some("##0.0E+0"),
        49 => Some("@"),
        _ => None,
    }
}

/// Convert an Excel serial date number to (year, month, day).
///
/// Excel stores dates as the number of days since 1899-12-30 (in the 1900 system).
/// Serial 1 = January 1, 1900.
/// Note: Excel incorrectly treats 1900 as a leap year (Feb 29, 1900 = serial 60).
pub fn excel_date_to_ymd(serial: f64) -> Option<(i32, u32, u32)> {
    if serial < 1.0 {
        return None;
    }

    let mut days = serial.floor() as i32;

    // Handle Excel's leap year bug: serial 60 = Feb 29, 1900 which doesn't exist
    // For dates after this, we need to subtract 1
    if days > 60 {
        days -= 1;
    } else if days == 60 {
        // Feb 29, 1900 doesn't really exist, but Excel thinks it does
        return Some((1900, 2, 29));
    }

    // days is now the actual number of days since Dec 31, 1899
    // day 1 = Jan 1, 1900
    days -= 1; // Convert to 0-based

    // Calculate year
    let mut year = 1900;
    loop {
        let days_in_year = if is_leap_year(year) { 366 } else { 365 };
        if days < days_in_year {
            break;
        }
        days -= days_in_year;
        year += 1;
    }

    // Calculate month and day
    let leap = is_leap_year(year);
    let days_in_months: [i32; 12] = if leap {
        [31, 29, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31]
    } else {
        [31, 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31]
    };

    let mut month = 1u32;
    for &dim in &days_in_months {
        if days < dim {
            break;
        }
        days -= dim;
        month += 1;
    }

    Some((year, month, (days + 1) as u32))
}

/// Check if a year is a leap year.
fn is_leap_year(year: i32) -> bool {
    (year % 4 == 0 && year % 100 != 0) || (year % 400 == 0)
}

/// Convert an Excel serial date/time to (year, month, day, hour, minute, second).
pub fn excel_datetime_to_ymdhms(serial: f64) -> Option<(i32, u32, u32, u32, u32, u32)> {
    let (y, m, d) = excel_date_to_ymd(serial)?;

    // Extract time from fractional part
    let time_fraction = serial.fract();
    let total_seconds = (time_fraction * 86400.0).round() as u32;
    let hours = total_seconds / 3600;
    let minutes = (total_seconds % 3600) / 60;
    let seconds = total_seconds % 60;

    Some((y, m, d, hours, minutes, seconds))
}

/// Format an Excel date serial number as a string (YYYY-MM-DD).
pub fn format_excel_date(serial: f64) -> Option<String> {
    let (y, m, d) = excel_date_to_ymd(serial)?;
    Some(format!("{:04}-{:02}-{:02}", y, m, d))
}

/// Format an Excel datetime serial number as a string (YYYY-MM-DD HH:MM:SS).
pub fn format_excel_datetime(serial: f64) -> Option<String> {
    let (y, m, d, h, min, s) = excel_datetime_to_ymdhms(serial)?;
    if h == 0 && min == 0 && s == 0 {
        Some(format!("{:04}-{:02}-{:02}", y, m, d))
    } else {
        Some(format!(
            "{:04}-{:02}-{:02} {:02}:{:02}:{:02}",
            y, m, d, h, min, s
        ))
    }
}

/// Check if a format code represents a date/time format.
fn is_date_format_code(code: &str) -> bool {
    // Skip color codes like [Red], [Color1], etc.
    let code = code.to_lowercase();

    // Remove sections in square brackets (colors, conditions)
    let mut clean = String::new();
    let mut in_bracket = false;
    for c in code.chars() {
        match c {
            '[' => in_bracket = true,
            ']' => in_bracket = false,
            _ if !in_bracket => clean.push(c),
            _ => {}
        }
    }

    // Check for date/time tokens (not preceded by backslash escape)
    let date_tokens = ["y", "m", "d", "h", "s"];
    for token in date_tokens {
        if clean.contains(token) {
            // Make sure it's not just in a string literal
            return true;
        }
    }

    false
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

/// Parse defined names (named ranges) from workbook.xml.
fn parse_defined_names(xml: &[u8]) -> Result<Vec<DefinedName>> {
    let mut reader = Reader::from_reader(Cursor::new(xml));
    let mut buf = Vec::new();
    let mut names = Vec::new();
    let mut in_defined_names = false;
    let mut current_name: Option<(String, Option<u32>, Option<String>, bool)> = None;
    let mut current_text = String::new();

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) => {
                let tag = e.name();
                let tag = tag.as_ref();
                if tag == b"definedNames" {
                    in_defined_names = true;
                } else if in_defined_names && tag == b"definedName" {
                    let mut name = String::new();
                    let mut local_sheet_id = None;
                    let mut comment = None;
                    let mut hidden = false;

                    for attr in e.attributes().filter_map(|a| a.ok()) {
                        match attr.key.as_ref() {
                            b"name" => {
                                name = String::from_utf8_lossy(&attr.value).into_owned();
                            }
                            b"localSheetId" => {
                                local_sheet_id = String::from_utf8_lossy(&attr.value).parse().ok();
                            }
                            b"comment" => {
                                comment = Some(String::from_utf8_lossy(&attr.value).into_owned());
                            }
                            b"hidden" => {
                                hidden = &*attr.value == b"1" || &*attr.value == b"true";
                            }
                            _ => {}
                        }
                    }

                    current_name = Some((name, local_sheet_id, comment, hidden));
                    current_text.clear();
                }
            }
            Ok(Event::Text(e)) => {
                if current_name.is_some() {
                    current_text.push_str(&e.decode().unwrap_or_default());
                }
            }
            Ok(Event::End(e)) => {
                let tag = e.name();
                let tag = tag.as_ref();
                if tag == b"definedNames" {
                    in_defined_names = false;
                } else if tag == b"definedName" {
                    if let Some((name, local_sheet_id, comment, hidden)) = current_name.take()
                        && !name.is_empty()
                        && !current_text.is_empty()
                    {
                        names.push(DefinedName {
                            name,
                            reference: current_text.clone(),
                            local_sheet_id,
                            comment,
                            hidden,
                        });
                    }
                    current_text.clear();
                }
            }
            Ok(Event::Eof) => break,
            Err(e) => return Err(Error::Xml(e)),
            _ => {}
        }
        buf.clear();
    }

    Ok(names)
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

    // Conditional formatting state
    let mut conditional_formats: Vec<ConditionalFormatting> = Vec::new();
    let mut in_conditional_formatting = false;
    let mut current_cf_ranges = String::new();
    let mut current_cf_rules: Vec<ConditionalRule> = Vec::new();
    let mut in_cf_rule = false;
    let mut current_cf_rule: Option<ConditionalRule> = None;
    let mut current_cf_formula = String::new();
    let mut in_cf_formula = false;
    let mut current_color_scale: Option<ColorScale> = None;
    let mut current_data_bar: Option<DataBar> = None;
    let mut current_icon_set: Option<IconSet> = None;

    // Data validation state
    let mut data_validations: Vec<DataValidation> = Vec::new();
    let mut in_data_validation = false;
    let mut current_dv: Option<DataValidation> = None;
    let mut in_dv_formula1 = false;
    let mut in_dv_formula2 = false;
    let mut current_dv_formula1 = String::new();
    let mut current_dv_formula2 = String::new();

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
                    b"conditionalFormatting" => {
                        in_conditional_formatting = true;
                        current_cf_ranges.clear();
                        current_cf_rules.clear();
                        for attr in e.attributes().filter_map(|a| a.ok()) {
                            if attr.key.as_ref() == b"sqref" {
                                current_cf_ranges =
                                    String::from_utf8_lossy(&attr.value).into_owned();
                            }
                        }
                    }
                    b"cfRule" if in_conditional_formatting => {
                        in_cf_rule = true;
                        let mut rule = ConditionalRule {
                            rule_type: ConditionalRuleType::Expression,
                            priority: 0,
                            dxf_id: None,
                            operator: None,
                            formulas: Vec::new(),
                            stop_if_true: false,
                            above_average: None,
                            equal_average: None,
                            rank: None,
                            top: None,
                            percent: None,
                            text: None,
                            time_period: None,
                            color_scale: None,
                            data_bar: None,
                            icon_set: None,
                        };

                        for attr in e.attributes().filter_map(|a| a.ok()) {
                            let val = String::from_utf8_lossy(&attr.value);
                            match attr.key.as_ref() {
                                b"type" => {
                                    if let Some(rt) = ConditionalRuleType::parse(&val) {
                                        rule.rule_type = rt;
                                    }
                                }
                                b"priority" => {
                                    rule.priority = val.parse().unwrap_or(0);
                                }
                                b"dxfId" => {
                                    rule.dxf_id = val.parse().ok();
                                }
                                b"operator" => {
                                    rule.operator = Some(val.into_owned());
                                }
                                b"stopIfTrue" => {
                                    rule.stop_if_true = val == "1" || val == "true";
                                }
                                b"aboveAverage" => {
                                    rule.above_average = Some(val != "0" && val != "false");
                                }
                                b"equalAverage" => {
                                    rule.equal_average = Some(val == "1" || val == "true");
                                }
                                b"rank" => {
                                    rule.rank = val.parse().ok();
                                }
                                b"top" => {
                                    rule.top = Some(val != "0" && val != "false");
                                }
                                b"percent" => {
                                    rule.percent = Some(val == "1" || val == "true");
                                }
                                b"text" => {
                                    rule.text = Some(val.into_owned());
                                }
                                b"timePeriod" => {
                                    rule.time_period = Some(val.into_owned());
                                }
                                _ => {}
                            }
                        }
                        current_cf_rule = Some(rule);
                    }
                    b"colorScale" if in_cf_rule => {
                        current_color_scale = Some(ColorScale { colors: Vec::new() });
                    }
                    b"dataBar" if in_cf_rule => {
                        let mut data_bar = DataBar {
                            min_type: String::new(),
                            min_value: None,
                            max_type: String::new(),
                            max_value: None,
                            color: None,
                            show_value: true,
                        };
                        for attr in e.attributes().filter_map(|a| a.ok()) {
                            let val = String::from_utf8_lossy(&attr.value);
                            if attr.key.as_ref() == b"showValue" {
                                data_bar.show_value = val != "0" && val != "false";
                            }
                        }
                        current_data_bar = Some(data_bar);
                    }
                    b"iconSet" if in_cf_rule => {
                        let mut icon_set = IconSet {
                            icon_set: "3TrafficLights1".to_string(),
                            show_value: true,
                            reverse: false,
                            values: Vec::new(),
                        };
                        for attr in e.attributes().filter_map(|a| a.ok()) {
                            let val = String::from_utf8_lossy(&attr.value);
                            match attr.key.as_ref() {
                                b"iconSet" => icon_set.icon_set = val.into_owned(),
                                b"showValue" => icon_set.show_value = val != "0" && val != "false",
                                b"reverse" => icon_set.reverse = val == "1" || val == "true",
                                _ => {}
                            }
                        }
                        current_icon_set = Some(icon_set);
                    }
                    b"formula" if in_cf_rule => {
                        in_cf_formula = true;
                        current_cf_formula.clear();
                    }
                    b"dataValidation" => {
                        in_data_validation = true;
                        let mut dv = DataValidation {
                            ranges: String::new(),
                            validation_type: DataValidationType::None,
                            operator: DataValidationOperator::Between,
                            formula1: None,
                            formula2: None,
                            allow_blank: false,
                            show_input_message: false,
                            show_error_message: false,
                            error_style: DataValidationErrorStyle::Stop,
                            error_title: None,
                            error_message: None,
                            prompt_title: None,
                            prompt_message: None,
                        };

                        for attr in e.attributes().filter_map(|a| a.ok()) {
                            let val = String::from_utf8_lossy(&attr.value);
                            match attr.key.as_ref() {
                                b"sqref" => dv.ranges = val.into_owned(),
                                b"type" => dv.validation_type = DataValidationType::parse(&val),
                                b"operator" => dv.operator = DataValidationOperator::parse(&val),
                                b"allowBlank" => dv.allow_blank = val == "1" || val == "true",
                                b"showInputMessage" => {
                                    dv.show_input_message = val == "1" || val == "true"
                                }
                                b"showErrorMessage" => {
                                    dv.show_error_message = val == "1" || val == "true"
                                }
                                b"errorStyle" => {
                                    dv.error_style = DataValidationErrorStyle::parse(&val)
                                }
                                b"errorTitle" => dv.error_title = Some(val.into_owned()),
                                b"error" => dv.error_message = Some(val.into_owned()),
                                b"promptTitle" => dv.prompt_title = Some(val.into_owned()),
                                b"prompt" => dv.prompt_message = Some(val.into_owned()),
                                _ => {}
                            }
                        }
                        current_dv = Some(dv);
                        current_dv_formula1.clear();
                        current_dv_formula2.clear();
                    }
                    b"formula1" if in_data_validation => {
                        in_dv_formula1 = true;
                        current_dv_formula1.clear();
                    }
                    b"formula2" if in_data_validation => {
                        in_dv_formula2 = true;
                        current_dv_formula2.clear();
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
                // Handle cfvo (conditional format value object) - used in colorScale, dataBar, iconSet
                else if tag_ref == b"cfvo" {
                    let mut val_type = String::new();
                    let mut val = None;
                    for attr in e.attributes().filter_map(|a| a.ok()) {
                        let attr_val = String::from_utf8_lossy(&attr.value);
                        match attr.key.as_ref() {
                            b"type" => val_type = attr_val.into_owned(),
                            b"val" => val = Some(attr_val.into_owned()),
                            _ => {}
                        }
                    }

                    if let Some(ref mut cs) = current_color_scale {
                        cs.colors.push(ColorScaleValue {
                            value_type: val_type.clone(),
                            value: val.clone(),
                            color: None,
                        });
                    } else if let Some(ref mut db) = current_data_bar {
                        if db.min_type.is_empty() {
                            db.min_type = val_type;
                            db.min_value = val;
                        } else {
                            db.max_type = val_type;
                            db.max_value = val;
                        }
                    } else if let Some(ref mut is) = current_icon_set {
                        is.values.push(IconSetValue {
                            value_type: val_type,
                            value: val,
                        });
                    }
                }
                // Handle color element in colorScale or dataBar
                else if tag_ref == b"color" {
                    let mut color = None;
                    for attr in e.attributes().filter_map(|a| a.ok()) {
                        if attr.key.as_ref() == b"rgb" {
                            color = Some(String::from_utf8_lossy(&attr.value).into_owned());
                        }
                    }

                    if let Some(ref mut cs) = current_color_scale {
                        if let Some(last) = cs.colors.last_mut() {
                            last.color = color;
                        }
                    } else if let Some(ref mut db) = current_data_bar {
                        db.color = color;
                    }
                }
            }
            Ok(Event::Text(e)) => {
                if in_value {
                    current_cell_value.push_str(&e.decode().unwrap_or_default());
                } else if in_formula && let Some(ref mut f) = current_formula {
                    f.push_str(&e.decode().unwrap_or_default());
                } else if in_cf_formula {
                    current_cf_formula.push_str(&e.decode().unwrap_or_default());
                } else if in_dv_formula1 {
                    current_dv_formula1.push_str(&e.decode().unwrap_or_default());
                } else if in_dv_formula2 {
                    current_dv_formula2.push_str(&e.decode().unwrap_or_default());
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
                    b"formula" if in_cf_rule => {
                        if !current_cf_formula.is_empty()
                            && let Some(ref mut rule) = current_cf_rule
                        {
                            rule.formulas.push(std::mem::take(&mut current_cf_formula));
                        }
                        in_cf_formula = false;
                    }
                    b"colorScale" if in_cf_rule => {
                        if let Some(ref mut rule) = current_cf_rule {
                            rule.color_scale = current_color_scale.take();
                        }
                    }
                    b"dataBar" if in_cf_rule => {
                        if let Some(ref mut rule) = current_cf_rule {
                            rule.data_bar = current_data_bar.take();
                        }
                    }
                    b"iconSet" if in_cf_rule => {
                        if let Some(ref mut rule) = current_cf_rule {
                            rule.icon_set = current_icon_set.take();
                        }
                    }
                    b"cfRule" if in_cf_rule => {
                        if let Some(rule) = current_cf_rule.take() {
                            current_cf_rules.push(rule);
                        }
                        in_cf_rule = false;
                    }
                    b"conditionalFormatting" => {
                        if !current_cf_ranges.is_empty() {
                            conditional_formats.push(ConditionalFormatting {
                                ranges: std::mem::take(&mut current_cf_ranges),
                                rules: std::mem::take(&mut current_cf_rules),
                            });
                        }
                        in_conditional_formatting = false;
                    }
                    b"formula1" if in_data_validation => {
                        in_dv_formula1 = false;
                    }
                    b"formula2" if in_data_validation => {
                        in_dv_formula2 = false;
                    }
                    b"dataValidation" => {
                        if let Some(mut dv) = current_dv.take() {
                            if !current_dv_formula1.is_empty() {
                                dv.formula1 = Some(std::mem::take(&mut current_dv_formula1));
                            }
                            if !current_dv_formula2.is_empty() {
                                dv.formula2 = Some(std::mem::take(&mut current_dv_formula2));
                            }
                            data_validations.push(dv);
                        }
                        in_data_validation = false;
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
        conditional_formats,
        data_validations,
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

    #[test]
    fn test_excel_date_conversion() {
        // Test some known dates
        // January 1, 2000 = serial 36526
        assert_eq!(excel_date_to_ymd(36526.0), Some((2000, 1, 1)));

        // December 31, 1999 = serial 36525
        assert_eq!(excel_date_to_ymd(36525.0), Some((1999, 12, 31)));

        // January 1, 1900 = serial 1
        assert_eq!(excel_date_to_ymd(1.0), Some((1900, 1, 1)));

        // March 1, 1900 = serial 61 (after the leap year bug)
        assert_eq!(excel_date_to_ymd(61.0), Some((1900, 3, 1)));

        // Test datetime
        // Noon on Jan 1, 2000 = 36526.5
        assert_eq!(
            excel_datetime_to_ymdhms(36526.5),
            Some((2000, 1, 1, 12, 0, 0))
        );

        // Format functions
        assert_eq!(format_excel_date(36526.0), Some("2000-01-01".to_string()));
        assert_eq!(
            format_excel_datetime(36526.5),
            Some("2000-01-01 12:00:00".to_string())
        );
    }

    #[test]
    fn test_builtin_format_codes() {
        assert_eq!(builtin_format_code(0), Some("General"));
        assert_eq!(builtin_format_code(1), Some("0"));
        assert_eq!(builtin_format_code(14), Some("mm-dd-yy"));
        assert_eq!(builtin_format_code(22), Some("m/d/yy h:mm"));
        assert_eq!(builtin_format_code(49), Some("@"));
        assert_eq!(builtin_format_code(100), None);
    }

    #[test]
    fn test_is_date_format() {
        assert!(is_date_format_code("mm-dd-yy"));
        assert!(is_date_format_code("yyyy-mm-dd"));
        assert!(is_date_format_code("d-mmm-yy"));
        assert!(is_date_format_code("h:mm:ss"));
        assert!(is_date_format_code("[Red]yyyy-mm-dd")); // With color code
        assert!(!is_date_format_code("0.00"));
        assert!(!is_date_format_code("#,##0"));
        assert!(!is_date_format_code("General"));
        assert!(!is_date_format_code("@")); // Text format
    }

    #[test]
    fn test_parse_defined_names() {
        let xml = r#"<?xml version="1.0"?>
        <workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
            <sheets>
                <sheet name="Sheet1" sheetId="1" r:id="rId1"/>
            </sheets>
            <definedNames>
                <definedName name="MyRange">Sheet1!$A$1:$B$10</definedName>
                <definedName name="LocalName" localSheetId="0">Sheet1!$C$1:$C$5</definedName>
                <definedName name="_xlnm.Print_Area" localSheetId="0" comment="Print area">Sheet1!$A$1:$F$20</definedName>
                <definedName name="HiddenName" hidden="1">Sheet1!$Z$1</definedName>
            </definedNames>
        </workbook>"#;

        let names = parse_defined_names(xml.as_bytes()).unwrap();
        assert_eq!(names.len(), 4);

        // Global name
        assert_eq!(names[0].name, "MyRange");
        assert_eq!(names[0].reference, "Sheet1!$A$1:$B$10");
        assert!(names[0].local_sheet_id.is_none());
        assert!(!names[0].is_builtin());

        // Local name
        assert_eq!(names[1].name, "LocalName");
        assert_eq!(names[1].local_sheet_id, Some(0));

        // Built-in name with comment
        assert_eq!(names[2].name, "_xlnm.Print_Area");
        assert!(names[2].is_builtin());
        assert_eq!(names[2].comment, Some("Print area".to_string()));

        // Hidden name
        assert_eq!(names[3].name, "HiddenName");
        assert!(names[3].hidden);
    }
}
