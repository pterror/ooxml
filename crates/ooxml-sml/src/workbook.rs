//! Workbook API for reading and writing Excel files.
//!
//! This module provides the main entry point for working with XLSX files.

use crate::error::{Error, Result};
use crate::ext::{
    Chart as ExtChart, ChartType as ExtChartType, Comment as ExtComment, ResolvedSheet,
    parse_worksheet,
};
use crate::parsers::FromXml;
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
const REL_PIVOT_TABLE: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotTable";
const REL_PIVOT_CACHE_DEFINITION: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheDefinition";

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
    styles: crate::types::Stylesheet,
    /// Defined names (named ranges).
    defined_names: Vec<crate::types::DefinedName>,
    /// Workbook protection settings (if any).
    #[cfg(feature = "sml-protection")]
    workbook_protection: Option<crate::types::WorkbookProtection>,
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

        // Parse workbook.xml once to get sheet list and defined names
        let workbook_xml = package.read_part(&workbook_path)?;
        let wb: crate::types::Workbook = bootstrap(&workbook_xml)?;
        let sheet_info = wb
            .sheets
            .sheet
            .iter()
            .map(|s| SheetInfo {
                name: s.name.to_string(),
                sheet_id: s.sheet_id,
                rel_id: s.id.to_string(),
            })
            .collect();
        let defined_names = wb
            .defined_names
            .map(|dn| {
                let inner = *dn;
                inner.defined_name
            })
            .unwrap_or_default();
        #[cfg(feature = "sml-protection")]
        let workbook_protection = wb.workbook_protection.map(|b| *b);

        // Load shared strings if present
        let shared_strings = if let Some(rel) = workbook_rels.get_by_type(REL_SHARED_STRINGS) {
            let path = resolve_path(&workbook_path, &rel.target);
            if let Ok(data) = package.read_part(&path) {
                let sst: crate::types::SharedStrings = bootstrap(&data)?;
                extract_shared_strings(&sst)
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
                bootstrap(&data)?
            } else {
                crate::types::Stylesheet::default()
            }
        } else {
            crate::types::Stylesheet::default()
        };

        Ok(Self {
            package,
            workbook_path,
            workbook_rels,
            sheet_info,
            shared_strings,
            styles,
            defined_names,
            #[cfg(feature = "sml-protection")]
            workbook_protection,
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
    pub fn styles(&self) -> &crate::types::Stylesheet {
        &self.styles
    }

    /// Get the workbook stylesheet (alias for [`styles`](Self::styles)).
    pub fn stylesheet(&self) -> Option<&crate::types::Stylesheet> {
        Some(&self.styles)
    }

    /// Get the workbook protection settings (if any).
    ///
    /// Requires the `sml-protection` feature.
    ///
    /// ECMA-376 Part 1, Section 18.2.29 (workbookProtection).
    #[cfg(feature = "sml-protection")]
    pub fn workbook_protection(&self) -> Option<&crate::types::WorkbookProtection> {
        self.workbook_protection.as_ref()
    }

    /// Access the raw workbook protection for testing.
    ///
    /// Requires the `sml-protection` feature.
    #[cfg(feature = "sml-protection")]
    pub fn raw_workbook_protection(&self) -> Option<&crate::types::WorkbookProtection> {
        self.workbook_protection.as_ref()
    }

    /// Get all defined names (named ranges).
    pub fn defined_names(&self) -> &[crate::types::DefinedName] {
        &self.defined_names
    }

    /// Get a defined name by its name.
    ///
    /// For names with sheet scope, use `defined_name_in_sheet` instead.
    pub fn defined_name(&self, name: &str) -> Option<&crate::types::DefinedName> {
        self.defined_names
            .iter()
            .find(|d| d.name.eq_ignore_ascii_case(name) && d.local_sheet_id.is_none())
    }

    /// Get a defined name by its name within a specific sheet scope.
    pub fn defined_name_in_sheet(
        &self,
        name: &str,
        sheet_index: u32,
    ) -> Option<&crate::types::DefinedName> {
        self.defined_names
            .iter()
            .find(|d| d.name.eq_ignore_ascii_case(name) && d.local_sheet_id == Some(sheet_index))
    }

    /// Get all global defined names (workbook scope).
    pub fn global_defined_names(&self) -> impl Iterator<Item = &crate::types::DefinedName> {
        self.defined_names
            .iter()
            .filter(|d| d.local_sheet_id.is_none())
    }

    /// Get all defined names scoped to a specific sheet.
    pub fn sheet_defined_names(
        &self,
        sheet_index: u32,
    ) -> impl Iterator<Item = &crate::types::DefinedName> {
        self.defined_names
            .iter()
            .filter(move |d| d.local_sheet_id == Some(sheet_index))
    }

    /// Load all pivot cache definitions from workbook relationships.
    ///
    /// Pivot cache definitions are workbook-level parts linked via relationships of type
    /// `pivotCacheDefinition`. Each `types::CTPivotTableDefinition` references a cache
    /// by its `cache_id`, which matches the `r:id` attribute on the cache definition.
    ///
    /// ECMA-376 Part 1, Section 18.10 (PivotTable).
    pub fn pivot_caches(&mut self) -> Result<Vec<crate::types::PivotCacheDefinition>> {
        let mut caches = Vec::new();
        let rel_targets: Vec<String> = self
            .workbook_rels
            .get_all_by_type(REL_PIVOT_CACHE_DEFINITION)
            .map(|r| r.target.clone())
            .collect();
        for target in rel_targets {
            let path = resolve_path(&self.workbook_path, &target);
            if let Ok(data) = self.package.read_part(&path)
                && let Ok(cache) = bootstrap::<crate::types::PivotCacheDefinition>(&data)
            {
                caches.push(cache);
            }
        }
        Ok(caches)
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

    /// Get raw worksheet XML bytes for lazy/streaming parsing.
    ///
    /// This returns the raw XML data for a worksheet, which can be used with
    /// `LazyWorksheet` for memory-efficient streaming access without loading
    /// all rows into memory.
    ///
    /// # Example
    ///
    /// ```ignore
    /// use ooxml_sml::{Workbook, LazyWorksheet};
    ///
    /// let mut workbook = Workbook::open("large.xlsx")?;
    /// let xml = workbook.sheet_xml(0)?;
    /// let lazy = LazyWorksheet::new(&xml);
    ///
    /// // Stream rows without loading all into memory
    /// for row in lazy.rows() {
    ///     let row = row?;
    ///     // Process row...
    /// }
    /// ```
    pub fn sheet_xml(&mut self, index: usize) -> Result<Vec<u8>> {
        let info = self
            .sheet_info
            .get(index)
            .ok_or_else(|| Error::Invalid(format!("Sheet index {} out of range", index)))?
            .clone();

        let rel = self.workbook_rels.get(&info.rel_id).ok_or_else(|| {
            Error::Invalid(format!("Missing relationship for sheet '{}'", info.name))
        })?;

        let path = resolve_path(&self.workbook_path, &rel.target);
        Ok(self.package.read_part(&path)?)
    }

    /// Get raw worksheet XML bytes by sheet name.
    ///
    /// See `sheet_xml` for usage with `LazyWorksheet`.
    pub fn sheet_xml_by_name(&mut self, name: &str) -> Result<Vec<u8>> {
        let info = self
            .sheet_info
            .iter()
            .find(|s| s.name == name)
            .ok_or_else(|| Error::Invalid(format!("Sheet '{}' not found", name)))?
            .clone();

        let rel = self.workbook_rels.get(&info.rel_id).ok_or_else(|| {
            Error::Invalid(format!("Missing relationship for sheet '{}'", info.name))
        })?;

        let path = resolve_path(&self.workbook_path, &rel.target);
        Ok(self.package.read_part(&path)?)
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

        // Load comments, charts, and pivot tables
        let mut comments = Vec::new();
        let mut charts = Vec::new();
        #[cfg(feature = "sml-pivot")]
        let mut pivot_tables: Vec<crate::types::CTPivotTableDefinition> = Vec::new();

        if let Ok(sheet_rels) = self.package.read_part_relationships(&path) {
            // Load comments
            if !is_chartsheet && let Some(comments_rel) = sheet_rels.get_by_type(REL_COMMENTS) {
                let comments_path = resolve_path(&path, &comments_rel.target);
                if let Ok(comments_data) = self.package.read_part(&comments_path) {
                    comments = parse_comments_xml(&comments_data)?;
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

            // Load pivot tables from sheet relationships
            #[cfg(feature = "sml-pivot")]
            for rel in sheet_rels.get_all_by_type(REL_PIVOT_TABLE) {
                let pt_path = resolve_path(&path, &rel.target);
                if let Ok(pt_data) = self.package.read_part(&pt_path)
                    && let Ok(pt) = bootstrap::<crate::types::CTPivotTableDefinition>(&pt_data)
                {
                    pivot_tables.push(pt);
                }
            }
        }

        Ok(ResolvedSheet::with_extras(
            info.name.clone(),
            worksheet,
            self.shared_strings.clone(),
            comments,
            charts,
            #[cfg(feature = "sml-pivot")]
            pivot_tables,
        ))
    }
}

/// Parse comments using the generated FromXml parser.
fn parse_comments_xml(xml: &[u8]) -> Result<Vec<ExtComment>> {
    let comments: crate::types::Comments = bootstrap(xml)?;
    let authors = comments.authors.author.clone();
    #[cfg(not(feature = "sml-comments"))]
    {
        let _ = (authors, comments.comment_list);
        return Ok(Vec::new());
    }
    #[cfg(feature = "sml-comments")]
    Ok(comments
        .comment_list
        .comment
        .iter()
        .map(|c| {
            let author = authors.get(c.author_id as usize).cloned();
            let text = rich_string_to_plain(&c.text);
            ExtComment {
                reference: c.reference.clone(),
                author,
                text,
            }
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

/// Type of conditional formatting rule.
///
/// Used when writing conditional formatting via [`ConditionalFormat`].
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

/// A chart embedded in a worksheet (workbook-module representation).
///
/// Returned by [`Workbook::resolved_sheet`] as part of [`ResolvedSheet`];
/// see also [`ext::Chart`](crate::ext::Chart) which is the simpler public type.
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

/// The type of chart as reported by the workbook reader.
///
/// Covers all chart types defined in ECMA-376 Part 1, §21.2.
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

/// A data series within a chart, as parsed by the workbook reader.
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

/// The scope of a defined name: either workbook-global or limited to one sheet.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum DefinedNameScope {
    /// Global workbook scope.
    Workbook,
    /// Local to a specific sheet (by index).
    Sheet(u32),
}

// ============================================================================
// Extension Traits for Generated Types
// ============================================================================

/// Extension methods for the generated [`types::Stylesheet`](crate::types::Stylesheet).
pub trait StylesheetExt {
    /// Get a number format code by ID.
    ///
    /// Looks up custom formats first, then falls back to built-in formats.
    fn format_code(&self, id: u32) -> Option<String>;

    /// Check if a format ID represents a date/time format.
    fn is_date_format(&self, id: u32) -> bool;
}

impl StylesheetExt for crate::types::Stylesheet {
    fn format_code(&self, id: u32) -> Option<String> {
        #[cfg(feature = "sml-styling")]
        if let Some(num_fmts) = &self.num_fmts
            && let Some(fmt) = num_fmts.num_fmt.iter().find(|f| f.number_format_id == id)
        {
            return Some(fmt.format_code.clone());
        }
        builtin_format_code(id).map(|s| s.to_string())
    }

    fn is_date_format(&self, id: u32) -> bool {
        if let Some(code) = self.format_code(id) {
            is_date_format_code(&code)
        } else {
            // Check built-in date format IDs (14-22, 45-47)
            matches!(id, 14..=22 | 45..=47)
        }
    }
}

/// Extension methods for the generated [`types::DefinedName`](crate::types::DefinedName).
pub trait DefinedNameExt {
    /// Check if this is a built-in Excel name (prefixed with "_xlnm.").
    ///
    /// Built-in names include:
    /// - _xlnm.Print_Area
    /// - _xlnm.Print_Titles
    /// - _xlnm._FilterDatabase
    /// - _xlnm.Criteria
    /// - _xlnm.Extract
    fn is_builtin(&self) -> bool;

    /// Get the scope of this defined name.
    fn scope(&self) -> DefinedNameScope;
}

impl DefinedNameExt for crate::types::DefinedName {
    fn is_builtin(&self) -> bool {
        self.name.starts_with("_xlnm.")
    }

    fn scope(&self) -> DefinedNameScope {
        match self.local_sheet_id {
            Some(id) => DefinedNameScope::Sheet(id),
            None => DefinedNameScope::Workbook,
        }
    }
}

// ============================================================================
// Utility Functions
// ============================================================================

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

// ============================================================================
// Private Parsing Helpers
// ============================================================================

/// Bootstrap a generated type from raw XML bytes.
///
/// Scans the XML for the first element and calls `T::from_xml` on it.
pub(crate) fn bootstrap<T: FromXml>(xml: &[u8]) -> Result<T> {
    let mut reader = Reader::from_reader(Cursor::new(xml));
    let mut buf = Vec::new();
    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) => {
                return T::from_xml(&mut reader, &e, false)
                    .map_err(|e| Error::Invalid(format!("{e:?}")));
            }
            Ok(Event::Empty(e)) => {
                return T::from_xml(&mut reader, &e, true)
                    .map_err(|e| Error::Invalid(format!("{e:?}")));
            }
            Ok(Event::Eof) => break,
            Err(e) => return Err(Error::Xml(e)),
            _ => {}
        }
        buf.clear();
    }
    Err(Error::Invalid("element not found".into()))
}

/// Extract plain strings from a SharedStrings SST.
fn extract_shared_strings(sst: &crate::types::SharedStrings) -> Vec<String> {
    sst.si
        .iter()
        .map(|si| {
            if let Some(t) = &si.cell_type {
                t.clone()
            } else {
                si.reference.iter().map(|r| r.cell_type.as_str()).collect()
            }
        })
        .collect()
}

/// Extract plain text from a RichString.
fn rich_string_to_plain(rs: &crate::types::RichString) -> String {
    if let Some(t) = &rs.cell_type {
        t.clone()
    } else {
        rs.reference.iter().map(|r| r.cell_type.as_str()).collect()
    }
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
    fn test_bootstrap_shared_strings() {
        let xml = r#"<?xml version="1.0"?>
        <sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
            <si><t>Hello</t></si>
            <si><t>World</t></si>
            <si><t></t></si>
        </sst>"#;

        let sst: crate::types::SharedStrings = bootstrap(xml.as_bytes()).unwrap();
        let strings = extract_shared_strings(&sst);
        assert_eq!(strings, vec!["Hello", "World", ""]);
    }

    #[test]
    #[cfg(feature = "sml-styling")]
    fn test_stylesheet_ext() {
        let xml = r#"<?xml version="1.0"?>
        <styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
            <numFmts count="1">
                <numFmt numFmtId="164" formatCode="0.00%"/>
            </numFmts>
        </styleSheet>"#;

        let styles: crate::types::Stylesheet = bootstrap(xml.as_bytes()).unwrap();

        // Custom format lookup
        assert_eq!(styles.format_code(164), Some("0.00%".to_string()));

        // Built-in format fallback
        assert_eq!(styles.format_code(14), Some("mm-dd-yy".to_string()));
        assert_eq!(styles.format_code(0), Some("General".to_string()));
        assert_eq!(styles.format_code(100), None);

        // Date format detection
        assert!(styles.is_date_format(14));
        assert!(!styles.is_date_format(0));
        assert!(!styles.is_date_format(164)); // "0.00%" is not a date format
    }

    #[test]
    #[cfg(feature = "sml-comments")]
    fn test_parse_comments_xml() {
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

        let comments = parse_comments_xml(xml.as_bytes()).unwrap();
        assert_eq!(comments.len(), 2);

        assert_eq!(comments[0].reference, "A1");
        assert_eq!(comments[0].author, Some("John Doe".to_string()));
        assert_eq!(comments[0].text, "This is a comment on A1");

        assert_eq!(comments[1].reference, "B2");
        assert_eq!(comments[1].author, Some("Jane Smith".to_string()));
        assert_eq!(comments[1].text, "Multi-line comment");
    }

    #[test]
    fn test_bootstrap_defined_names() {
        let xml = r#"<?xml version="1.0"?>
        <workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
                  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
            <sheets>
                <sheet name="Sheet1" sheetId="1" r:id="rId1"/>
            </sheets>
            <definedNames>
                <definedName name="MyRange">Sheet1!$A$1:$B$10</definedName>
                <definedName name="LocalName" localSheetId="0">Sheet1!$C$1:$C$5</definedName>
                <definedName name="_xlnm.Print_Area" localSheetId="0" comment="Print area">Sheet1!$A$1:$F$20</definedName>
            </definedNames>
        </workbook>"#;

        let wb: crate::types::Workbook = bootstrap(xml.as_bytes()).unwrap();
        let names = wb
            .defined_names
            .map(|dn| {
                let inner = *dn;
                inner.defined_name
            })
            .unwrap_or_default();
        assert_eq!(names.len(), 3);

        // Global name
        assert_eq!(names[0].name, "MyRange");
        assert_eq!(names[0].text.as_deref(), Some("Sheet1!$A$1:$B$10"));
        assert!(names[0].local_sheet_id.is_none());
        assert!(!names[0].is_builtin());
        assert_eq!(names[0].scope(), DefinedNameScope::Workbook);

        // Local name
        assert_eq!(names[1].name, "LocalName");
        assert_eq!(names[1].local_sheet_id, Some(0));
        assert_eq!(names[1].scope(), DefinedNameScope::Sheet(0));

        // Built-in name with comment
        assert_eq!(names[2].name, "_xlnm.Print_Area");
        assert!(names[2].is_builtin());
        assert_eq!(names[2].comment.as_deref(), Some("Print area"));
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
}
