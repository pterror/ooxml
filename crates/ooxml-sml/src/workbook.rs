//! Workbook API for reading and writing Excel files.
//!
//! This module provides the main entry point for working with XLSX files.

use crate::error::{Error, Result};
use crate::ext::{
    Chart as ExtChart, ChartType as ExtChartType, Comment as ExtComment, ResolvedSheet,
    parse_worksheet,
};
use ooxml_opc::{Package, Relationships};
use quick_xml::Reader;
use quick_xml::events::Event;
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
const REL_DRAWING: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing";
const REL_CHART: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart";
const REL_CHARTSHEET: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chartsheet";

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
    // =========================================================================
    // New API using generated types (ADR-003)
    // =========================================================================

    /// Get a sheet by index using the new generated parser.
    ///
    /// Returns a `ResolvedSheet` which wraps the generated `types::Worksheet`
    /// and provides automatic value resolution via extension traits.
    ///
    /// This is the recommended API for new code.
    pub fn resolved_sheet(&mut self, index: usize) -> Result<ResolvedSheet> {
        let info = self
            .sheet_info
            .get(index)
            .ok_or_else(|| Error::Invalid(format!("Sheet index {} out of range", index)))?
            .clone();

        self.load_resolved_sheet(&info)
    }

    /// Get a sheet by name using the new generated parser.
    ///
    /// Returns a `ResolvedSheet` which wraps the generated `types::Worksheet`
    /// and provides automatic value resolution via extension traits.
    pub fn resolved_sheet_by_name(&mut self, name: &str) -> Result<ResolvedSheet> {
        let info = self
            .sheet_info
            .iter()
            .find(|s| s.name == name)
            .ok_or_else(|| Error::Invalid(format!("Sheet '{}' not found", name)))?
            .clone();

        self.load_resolved_sheet(&info)
    }

    /// Load all sheets using the new generated parser.
    pub fn resolved_sheets(&mut self) -> Result<Vec<ResolvedSheet>> {
        let infos: Vec<_> = self.sheet_info.clone();
        infos
            .iter()
            .map(|info| self.load_resolved_sheet(info))
            .collect()
    }

    /// Load a sheet using the generated parser.
    fn load_resolved_sheet(&mut self, info: &SheetInfo) -> Result<ResolvedSheet> {
        // Find the sheet path from relationships
        let rel = self.workbook_rels.get(&info.rel_id).ok_or_else(|| {
            Error::Invalid(format!("Missing relationship for sheet '{}'", info.name))
        })?;

        let path = resolve_path(&self.workbook_path, &rel.target);
        let data = self.package.read_part(&path)?;

        // Check if this is a chartsheet or regular worksheet
        let is_chartsheet = rel.relationship_type == REL_CHARTSHEET;

        // Parse the worksheet using generated FromXml parser
        let worksheet = if is_chartsheet {
            // Chartsheets don't have the same structure - parse minimal empty worksheet XML
            // This ensures feature-gated fields are handled correctly by the generated parser
            let minimal_xml = br#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData/></worksheet>"#;
            parse_worksheet(minimal_xml)
                .map_err(|e| Error::Invalid(format!("Chartsheet parse error: {:?}", e)))?
        } else {
            parse_worksheet(&data).map_err(|e| Error::Invalid(format!("Parse error: {:?}", e)))?
        };

        // Load comments and charts
        let mut comments = Vec::new();
        let mut charts = Vec::new();

        if let Ok(sheet_rels) = self.package.read_part_relationships(&path) {
            // Load comments
            if !is_chartsheet && let Some(comments_rel) = sheet_rels.get_by_type(REL_COMMENTS) {
                let comments_path = resolve_path(&path, &comments_rel.target);
                if let Ok(comments_data) = self.package.read_part(&comments_path) {
                    comments = parse_comments_ext(&comments_data)?;
                }
            }

            // Load charts via drawing relationships
            if let Some(drawing_rel) = sheet_rels.get_by_type(REL_DRAWING) {
                let drawing_path = resolve_path(&path, &drawing_rel.target);
                if let Ok(drawing_rels) = self.package.read_part_relationships(&drawing_path) {
                    for rel in drawing_rels.iter() {
                        let chart_path = resolve_path(&drawing_path, &rel.target);
                        if rel.relationship_type == REL_CHART
                            && let Ok(chart_data) = self.package.read_part(&chart_path)
                            && let Ok(chart) = parse_chart_ext(&chart_data)
                        {
                            charts.push(chart);
                        }
                    }
                }
            }
        }

        Ok(ResolvedSheet::with_extras(
            info.name.clone(),
            worksheet,
            self.shared_strings.clone(),
            comments,
            charts,
        ))
    }
}

/// Parse comments for ext::Comment
fn parse_comments_ext(xml: &[u8]) -> Result<Vec<ExtComment>> {
    // Reuse existing comment parsing but convert to ext::Comment
    let old_comments = parse_comments(xml)?;
    Ok(old_comments
        .into_iter()
        .map(|c| ExtComment {
            reference: c.reference,
            author: c.author,
            text: c.text,
        })
        .collect())
}

/// Parse chart for ext::Chart
fn parse_chart_ext(xml: &[u8]) -> Result<ExtChart> {
    let old_chart = parse_chart(xml)?;
    Ok(ExtChart {
        title: old_chart.title,
        chart_type: match old_chart.chart_type {
            ChartType::Bar | ChartType::Bar3D => ExtChartType::Bar,
            ChartType::Line | ChartType::Line3D => ExtChartType::Line,
            ChartType::Pie | ChartType::Pie3D => ExtChartType::Pie,
            ChartType::Area | ChartType::Area3D => ExtChartType::Area,
            ChartType::Surface | ChartType::Surface3D => ExtChartType::Surface,
            ChartType::Scatter => ExtChartType::Scatter,
            ChartType::Doughnut => ExtChartType::Doughnut,
            ChartType::Radar => ExtChartType::Radar,
            ChartType::Bubble => ExtChartType::Bubble,
            ChartType::Stock => ExtChartType::Stock,
            ChartType::Unknown => ExtChartType::Unknown,
        },
    })
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

    /// Convert to XML attribute value.
    pub fn to_xml_value(self) -> &'static str {
        match self {
            Self::Expression => "expression",
            Self::CellIs => "cellIs",
            Self::ColorScale => "colorScale",
            Self::DataBar => "dataBar",
            Self::IconSet => "iconSet",
            Self::Top10 => "top10",
            Self::UniqueValues => "uniqueValues",
            Self::DuplicateValues => "duplicateValues",
            Self::ContainsText => "containsText",
            Self::NotContainsText => "notContainsText",
            Self::BeginsWith => "beginsWith",
            Self::EndsWith => "endsWith",
            Self::ContainsBlanks => "containsBlanks",
            Self::NotContainsBlanks => "notContainsBlanks",
            Self::ContainsErrors => "containsErrors",
            Self::NotContainsErrors => "notContainsErrors",
            Self::TimePeriod => "timePeriod",
            Self::AboveAverage => "aboveAverage",
        }
    }
}

/// A chart embedded in a worksheet.
#[derive(Debug, Clone)]
pub struct Chart {
    title: Option<String>,
    chart_type: ChartType,
    series: Vec<ChartSeries>,
}

impl Chart {
    /// Get the chart title (if any).
    pub fn title(&self) -> Option<&str> {
        self.title.as_deref()
    }

    /// Get the chart type.
    pub fn chart_type(&self) -> ChartType {
        self.chart_type
    }

    /// Get all data series in the chart.
    pub fn series(&self) -> &[ChartSeries] {
        &self.series
    }
}

/// The type of chart.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub enum ChartType {
    Area,
    Area3D,
    Bar,
    Bar3D,
    Bubble,
    Doughnut,
    #[default]
    Line,
    Line3D,
    Pie,
    Pie3D,
    Radar,
    Scatter,
    Stock,
    Surface,
    Surface3D,
    Unknown,
}

/// A data series within a chart.
#[derive(Debug, Clone)]
pub struct ChartSeries {
    index: u32,
    name: Option<String>,
    category_ref: Option<String>,
    value_ref: Option<String>,
    categories: Vec<String>,
    values: Vec<f64>,
}

impl ChartSeries {
    /// Get the series index.
    pub fn index(&self) -> u32 {
        self.index
    }

    /// Get the series name (if any).
    pub fn name(&self) -> Option<&str> {
        self.name.as_deref()
    }

    /// Get the category data reference.
    pub fn category_ref(&self) -> Option<&str> {
        self.category_ref.as_deref()
    }

    /// Get the value data reference.
    pub fn value_ref(&self) -> Option<&str> {
        self.value_ref.as_deref()
    }

    /// Get the category labels.
    pub fn categories(&self) -> &[String] {
        &self.categories
    }

    /// Get the numeric values.
    pub fn values(&self) -> &[f64] {
        &self.values
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
    #[allow(dead_code)]
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

    /// Convert to XML attribute value.
    pub fn to_xml_value(self) -> &'static str {
        match self {
            Self::None => "none",
            Self::Whole => "whole",
            Self::Decimal => "decimal",
            Self::List => "list",
            Self::Date => "date",
            Self::Time => "time",
            Self::TextLength => "textLength",
            Self::Custom => "custom",
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
    #[allow(dead_code)]
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

    /// Convert to XML attribute value.
    pub fn to_xml_value(self) -> &'static str {
        match self {
            Self::Between => "between",
            Self::NotBetween => "notBetween",
            Self::Equal => "equal",
            Self::NotEqual => "notEqual",
            Self::LessThan => "lessThan",
            Self::LessThanOrEqual => "lessThanOrEqual",
            Self::GreaterThan => "greaterThan",
            Self::GreaterThanOrEqual => "greaterThanOrEqual",
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
    #[allow(dead_code)]
    fn parse(s: &str) -> Self {
        match s {
            "stop" => Self::Stop,
            "warning" => Self::Warning,
            "information" => Self::Information,
            _ => Self::Stop,
        }
    }

    /// Convert to XML attribute value.
    pub fn to_xml_value(self) -> &'static str {
        match self {
            Self::Stop => "stop",
            Self::Warning => "warning",
            Self::Information => "information",
        }
    }
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

/// Parse comments from a comments XML file.
///
/// ECMA-376 Part 1, Section 18.7 (Comments).
#[allow(dead_code)]
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

/// Parse a chart XML file.
#[allow(dead_code)]
fn parse_chart(xml: &[u8]) -> Result<Chart> {
    let mut reader = Reader::from_reader(Cursor::new(xml));
    let mut buf = Vec::new();

    let mut title: Option<String> = None;
    let mut chart_type = ChartType::Unknown;
    let mut series: Vec<ChartSeries> = Vec::new();

    let mut in_chart = false;
    let mut in_plot_area = false;
    let mut in_title = false;
    let mut in_title_tx = false;
    let mut in_title_rich = false;
    let mut in_title_p = false;
    let mut in_title_r = false;
    let mut in_title_t = false;
    let mut in_ser = false;
    let mut in_cat = false;
    let mut in_val = false;
    let mut in_str_ref = false;
    let mut in_num_ref = false;
    let mut in_str_cache = false;
    let mut in_num_cache = false;
    let mut in_pt = false;
    let mut in_v = false;
    let mut in_f = false;
    let mut in_tx = false;
    let mut in_ser_name_str_ref = false;

    let mut title_text = String::new();
    let mut current_series_idx: u32 = 0;
    let mut current_series_name: Option<String> = None;
    let mut current_cat_ref: Option<String> = None;
    let mut current_val_ref: Option<String> = None;
    let mut current_cat_values: Vec<String> = Vec::new();
    let mut current_val_values: Vec<f64> = Vec::new();
    let mut current_ref = String::new();
    let mut current_v = String::new();

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) => {
                let name = e.local_name();
                let name = name.as_ref();
                match name {
                    b"chart" => in_chart = true,
                    b"plotArea" if in_chart => in_plot_area = true,
                    b"lineChart" | b"line3DChart" if in_plot_area => {
                        chart_type = if name == b"line3DChart" {
                            ChartType::Line3D
                        } else {
                            ChartType::Line
                        };
                    }
                    b"barChart" | b"bar3DChart" if in_plot_area => {
                        chart_type = if name == b"bar3DChart" {
                            ChartType::Bar3D
                        } else {
                            ChartType::Bar
                        };
                    }
                    b"areaChart" | b"area3DChart" if in_plot_area => {
                        chart_type = if name == b"area3DChart" {
                            ChartType::Area3D
                        } else {
                            ChartType::Area
                        };
                    }
                    b"pieChart" | b"pie3DChart" if in_plot_area => {
                        chart_type = if name == b"pie3DChart" {
                            ChartType::Pie3D
                        } else {
                            ChartType::Pie
                        };
                    }
                    b"doughnutChart" if in_plot_area => chart_type = ChartType::Doughnut,
                    b"scatterChart" if in_plot_area => chart_type = ChartType::Scatter,
                    b"bubbleChart" if in_plot_area => chart_type = ChartType::Bubble,
                    b"radarChart" if in_plot_area => chart_type = ChartType::Radar,
                    b"stockChart" if in_plot_area => chart_type = ChartType::Stock,
                    b"surfaceChart" | b"surface3DChart" if in_plot_area => {
                        chart_type = if name == b"surface3DChart" {
                            ChartType::Surface3D
                        } else {
                            ChartType::Surface
                        };
                    }
                    b"title" if in_chart && !in_plot_area => {
                        in_title = true;
                        title_text.clear();
                    }
                    b"tx" if in_title => in_title_tx = true,
                    b"rich" if in_title_tx => in_title_rich = true,
                    b"p" if in_title_rich => in_title_p = true,
                    b"r" if in_title_p => in_title_r = true,
                    b"t" if in_title_r => in_title_t = true,
                    b"ser" if in_plot_area => {
                        in_ser = true;
                        current_series_idx = 0;
                        current_series_name = None;
                        current_cat_ref = None;
                        current_val_ref = None;
                        current_cat_values.clear();
                        current_val_values.clear();
                    }
                    b"idx" if in_ser => {
                        for attr in e.attributes().filter_map(|a| a.ok()) {
                            if attr.key.as_ref() == b"val" {
                                current_series_idx =
                                    String::from_utf8_lossy(&attr.value).parse().unwrap_or(0);
                            }
                        }
                    }
                    b"tx" if in_ser && !in_cat && !in_val => in_tx = true,
                    b"strRef" if in_tx => in_ser_name_str_ref = true,
                    b"v" if in_ser_name_str_ref => in_v = true,
                    b"cat" if in_ser => in_cat = true,
                    b"val" if in_ser => in_val = true,
                    b"strRef" if in_cat || in_val => {
                        in_str_ref = true;
                        current_ref.clear();
                    }
                    b"numRef" if in_cat || in_val => {
                        in_num_ref = true;
                        current_ref.clear();
                    }
                    b"strCache" if in_str_ref => in_str_cache = true,
                    b"numCache" if in_num_ref => in_num_cache = true,
                    b"pt" if in_str_cache || in_num_cache => {
                        in_pt = true;
                        current_v.clear();
                    }
                    b"v" if in_pt => in_v = true,
                    b"f" if (in_str_ref || in_num_ref) && !in_str_cache && !in_num_cache => {
                        in_f = true;
                        current_ref.clear();
                    }
                    _ => {}
                }
            }
            Ok(Event::Text(e)) => {
                let text = e.decode().unwrap_or_default();
                if in_title_t {
                    title_text.push_str(&text);
                } else if in_v {
                    current_v.push_str(&text);
                } else if in_f {
                    current_ref.push_str(&text);
                }
            }
            Ok(Event::End(e)) => {
                let name = e.local_name();
                let name = name.as_ref();
                match name {
                    b"chart" => in_chart = false,
                    b"plotArea" => in_plot_area = false,
                    b"title" if in_title => {
                        in_title = false;
                        if !title_text.is_empty() {
                            title = Some(std::mem::take(&mut title_text));
                        }
                    }
                    b"tx" if in_title => in_title_tx = false,
                    b"rich" if in_title_rich => in_title_rich = false,
                    b"p" if in_title_p => in_title_p = false,
                    b"r" if in_title_r => in_title_r = false,
                    b"t" if in_title_t => in_title_t = false,
                    b"ser" if in_ser => {
                        in_ser = false;
                        series.push(ChartSeries {
                            index: current_series_idx,
                            name: current_series_name.take(),
                            category_ref: current_cat_ref.take(),
                            value_ref: current_val_ref.take(),
                            categories: std::mem::take(&mut current_cat_values),
                            values: std::mem::take(&mut current_val_values),
                        });
                    }
                    b"tx" if in_tx => in_tx = false,
                    b"strRef" if in_ser_name_str_ref => in_ser_name_str_ref = false,
                    b"v" if in_v => {
                        in_v = false;
                        if in_pt {
                            if in_str_cache && in_cat {
                                current_cat_values.push(std::mem::take(&mut current_v));
                            } else if in_num_cache {
                                if let Ok(v) = current_v.parse::<f64>() {
                                    if in_cat {
                                        current_cat_values.push(current_v.clone());
                                    } else if in_val {
                                        current_val_values.push(v);
                                    }
                                }
                                current_v.clear();
                            }
                        } else if in_ser_name_str_ref {
                            current_series_name = Some(std::mem::take(&mut current_v));
                        }
                    }
                    b"cat" if in_cat => in_cat = false,
                    b"val" if in_val => in_val = false,
                    b"strRef" if in_str_ref => {
                        in_str_ref = false;
                        if in_cat && !current_ref.is_empty() {
                            current_cat_ref = Some(std::mem::take(&mut current_ref));
                        }
                    }
                    b"numRef" if in_num_ref => {
                        in_num_ref = false;
                        if in_cat && current_cat_ref.is_none() && !current_ref.is_empty() {
                            current_cat_ref = Some(std::mem::take(&mut current_ref));
                        } else if in_val && !current_ref.is_empty() {
                            current_val_ref = Some(std::mem::take(&mut current_ref));
                        }
                    }
                    b"strCache" if in_str_cache => in_str_cache = false,
                    b"numCache" if in_num_cache => in_num_cache = false,
                    b"pt" if in_pt => in_pt = false,
                    b"f" if in_f => {
                        in_f = false;
                        if (in_str_ref || in_num_ref) && !current_ref.is_empty() {
                            if in_cat && current_cat_ref.is_none() {
                                current_cat_ref = Some(current_ref.clone());
                            } else if in_val && current_val_ref.is_none() {
                                current_val_ref = Some(current_ref.clone());
                            }
                        }
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

    Ok(Chart {
        title,
        chart_type,
        series,
    })
}

// ============================================================================
/// Resolve a relative path against a base path.
fn resolve_path(base: &str, target: &str) -> String {
    let has_leading_slash = base.starts_with('/');

    if target.starts_with('/') {
        return target.to_string();
    }

    // Get the directory of the base path
    let base_dir = if let Some(idx) = base.rfind('/') {
        &base[..idx]
    } else {
        ""
    };

    // Build path segments, handling ".." for parent directory
    let mut parts: Vec<&str> = base_dir.split('/').filter(|s| !s.is_empty()).collect();
    for segment in target.split('/') {
        match segment {
            ".." => {
                parts.pop();
            }
            "." | "" => {}
            _ => {
                parts.push(segment);
            }
        }
    }

    let result = parts.join("/");
    if has_leading_slash {
        format!("/{}", result)
    } else {
        result
    }
}

#[cfg(test)]
mod tests {
    use super::*;

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
        // Parent directory handling
        assert_eq!(
            resolve_path("/xl/chartsheets/sheet1.xml", "../drawings/drawing1.xml"),
            "/xl/drawings/drawing1.xml"
        );
        assert_eq!(
            resolve_path("/xl/worksheets/sheet1.xml", "../comments1.xml"),
            "/xl/comments1.xml"
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
