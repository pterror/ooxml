//! Word document API.
//!
//! This module provides the main entry point for working with DOCX files.
//!
//! # Example
//!
//! ```ignore
//! use ooxml_wml::Document;
//!
//! let doc = Document::open("document.docx")?;
//! for para in doc.body().paragraphs() {
//!     for run in para.runs() {
//!         print!("{}", run.text());
//!     }
//!     println!();
//! }
//! ```

use crate::error::{Error, Result};
use crate::raw_xml::{PositionedAttr, PositionedNode, RawXmlElement, RawXmlNode};
use crate::styles::{Styles, merge_run_properties};
use ooxml::{Package, Relationships, rel_type, rels_path_for};
use quick_xml::Reader;
use quick_xml::events::Event;
use std::fs::File;
use std::io::{BufReader, Read, Seek};
use std::path::Path;

/// A Word document (.docx file).
///
/// This is the main entry point for reading and writing Word documents.
pub struct Document<R> {
    package: Package<R>,
    body: Body,
    styles: Styles,
    /// Document part relationships (for images, hyperlinks, etc.)
    doc_rels: Relationships,
    /// Path to the document part (e.g., "word/document.xml")
    doc_path: String,
}

impl Document<BufReader<File>> {
    /// Open a Word document from a file path.
    pub fn open<P: AsRef<Path>>(path: P) -> Result<Self> {
        let file = File::open(path)?;
        let reader = BufReader::new(file);
        Self::from_reader(reader)
    }
}

impl<R: Read + Seek> Document<R> {
    /// Open a Word document from a reader.
    pub fn from_reader(reader: R) -> Result<Self> {
        let mut package = Package::open(reader)?;

        // Find the main document part via relationships
        let rels = package.read_relationships()?;
        let doc_rel = rels
            .get_by_type(rel_type::OFFICE_DOCUMENT)
            .ok_or_else(|| Error::MissingPart("main document relationship".into()))?;

        let doc_path = doc_rel.target.clone();

        // Parse the document XML
        let doc_xml = package.read_part(&doc_path)?;
        let body = parse_document(&doc_xml)?;

        // Load document-level relationships (for images, hyperlinks, etc.)
        let doc_rels_path = rels_path_for(&doc_path);
        let doc_rels = if package.has_part(&doc_rels_path) {
            let rels_xml = package.read_part(&doc_rels_path)?;
            Relationships::parse(&rels_xml[..])?
        } else {
            Relationships::new()
        };

        // Load styles if available
        let styles = if let Some(styles_rel) = rels.get_by_type(rel_type::STYLES) {
            let styles_xml = package.read_part(&styles_rel.target)?;
            Styles::parse(&styles_xml[..])?
        } else {
            Styles::new()
        };

        Ok(Self {
            package,
            body,
            styles,
            doc_rels,
            doc_path,
        })
    }

    /// Get the document body.
    pub fn body(&self) -> &Body {
        &self.body
    }

    /// Get a mutable reference to the document body.
    pub fn body_mut(&mut self) -> &mut Body {
        &mut self.body
    }

    /// Get the underlying package.
    pub fn package(&self) -> &Package<R> {
        &self.package
    }

    /// Get a mutable reference to the underlying package.
    pub fn package_mut(&mut self) -> &mut Package<R> {
        &mut self.package
    }

    /// Get the document styles.
    pub fn styles(&self) -> &Styles {
        &self.styles
    }

    /// Resolve effective run properties for a run, combining direct and style formatting.
    ///
    /// This merges: default run properties -> paragraph style -> character style -> direct formatting.
    pub fn resolve_run_formatting(&self, para: &Paragraph, run: &Run) -> RunProperties {
        let mut props = self.styles.default_run().clone();

        // Apply paragraph style's run properties
        if let Some(pstyle) = para.properties().and_then(|p| p.style.as_ref()) {
            let style_props = self.styles.resolve_run_properties(pstyle);
            merge_run_properties(&mut props, &style_props);
        }

        // Apply character style
        if let Some(rstyle) = run.properties().and_then(|p| p.style.as_ref()) {
            let style_props = self.styles.resolve_run_properties(rstyle);
            merge_run_properties(&mut props, &style_props);
        }

        // Apply direct formatting
        if let Some(direct) = run.properties() {
            merge_run_properties(&mut props, direct);
        }

        props
    }

    /// Extract all text from the document.
    ///
    /// Paragraphs are separated by newlines.
    pub fn text(&self) -> String {
        self.body.text()
    }

    /// Get image data by relationship ID.
    ///
    /// Looks up the relationship, reads the image file from the package,
    /// and returns the image data with its content type.
    ///
    /// # Example
    ///
    /// ```ignore
    /// for para in doc.body().paragraphs() {
    ///     for run in para.runs() {
    ///         for drawing in run.drawings() {
    ///             for image in drawing.images() {
    ///                 let data = doc.get_image_data(image.rel_id())?;
    ///                 println!("Image type: {}", data.content_type);
    ///             }
    ///         }
    ///     }
    /// }
    /// ```
    pub fn get_image_data(&mut self, rel_id: &str) -> Result<ImageData> {
        // Look up the relationship
        let rel = self
            .doc_rels
            .get(rel_id)
            .ok_or_else(|| Error::MissingPart(format!("image relationship {}", rel_id)))?;

        // Resolve the target path relative to the document
        let image_path = resolve_path(&self.doc_path, &rel.target);

        // Read the image data from the package
        let data = self.package.read_part(&image_path)?;

        // Determine content type from extension
        let content_type = content_type_from_path(&image_path);

        Ok(ImageData { content_type, data })
    }

    /// Get the URL for a hyperlink by its relationship ID.
    ///
    /// Returns None if the relationship doesn't exist.
    pub fn get_hyperlink_url(&self, rel_id: &str) -> Option<&str> {
        self.doc_rels.get(rel_id).map(|rel| rel.target.as_str())
    }

    /// Get document relationships (for advanced use).
    pub fn doc_relationships(&self) -> &Relationships {
        &self.doc_rels
    }
}

/// Resolve a relative path against a base path.
fn resolve_path(base: &str, relative: &str) -> String {
    // If the target is absolute (starts with /), use it directly (without the /)
    if let Some(stripped) = relative.strip_prefix('/') {
        return stripped.to_string();
    }

    // Otherwise, resolve relative to the base directory
    if let Some(slash_pos) = base.rfind('/') {
        format!("{}/{}", &base[..slash_pos], relative)
    } else {
        relative.to_string()
    }
}

/// Determine MIME content type from file extension.
fn content_type_from_path(path: &str) -> String {
    let ext = path.rsplit('.').next().unwrap_or("").to_lowercase();
    match ext.as_str() {
        "png" => "image/png",
        "jpg" | "jpeg" => "image/jpeg",
        "gif" => "image/gif",
        "bmp" => "image/bmp",
        "tiff" | "tif" => "image/tiff",
        "webp" => "image/webp",
        "svg" => "image/svg+xml",
        "emf" => "image/x-emf",
        "wmf" => "image/x-wmf",
        _ => "application/octet-stream",
    }
    .to_string()
}

/// The document body containing paragraphs and other block-level elements.
#[derive(Debug, Clone, Default)]
pub struct Body {
    /// Block-level content (paragraphs and tables in order).
    content: Vec<BlockContent>,
    /// Section properties (page size, margins, etc.).
    section_properties: Option<SectionProperties>,
    /// Unknown child elements preserved for round-trip fidelity.
    /// Stored with original position index for correct ordering during serialization.
    pub unknown_children: Vec<PositionedNode>,
}

/// Block-level content in the document body.
#[derive(Debug, Clone)]
#[allow(clippy::large_enum_variant)] // Table is larger due to properties/grid, but boxing adds indirection
pub enum BlockContent {
    /// A paragraph.
    Paragraph(Paragraph),
    /// A table.
    Table(Table),
}

impl Body {
    /// Create an empty body.
    pub fn new() -> Self {
        Self::default()
    }

    /// Get all paragraphs in the body (flattened, including those in tables).
    pub fn paragraphs(&self) -> Vec<&Paragraph> {
        let mut paras = Vec::new();
        for block in &self.content {
            match block {
                BlockContent::Paragraph(p) => paras.push(p),
                BlockContent::Table(t) => {
                    for row in t.rows() {
                        for cell in row.cells() {
                            for p in cell.paragraphs() {
                                paras.push(p);
                            }
                        }
                    }
                }
            }
        }
        paras
    }

    /// Get block-level content.
    pub fn content(&self) -> &[BlockContent] {
        &self.content
    }

    /// Get mutable reference to block-level content.
    pub fn content_mut(&mut self) -> &mut Vec<BlockContent> {
        &mut self.content
    }

    /// Get all tables in the body.
    pub fn tables(&self) -> impl Iterator<Item = &Table> {
        self.content.iter().filter_map(|b| match b {
            BlockContent::Table(t) => Some(t),
            _ => None,
        })
    }

    /// Add a new paragraph to the body.
    pub fn add_paragraph(&mut self) -> &mut Paragraph {
        self.content.push(BlockContent::Paragraph(Paragraph::new()));
        match self.content.last_mut() {
            Some(BlockContent::Paragraph(p)) => p,
            _ => unreachable!(),
        }
    }

    /// Add a new table to the body.
    pub fn add_table(&mut self) -> &mut Table {
        self.content.push(BlockContent::Table(Table::new()));
        match self.content.last_mut() {
            Some(BlockContent::Table(t)) => t,
            _ => unreachable!(),
        }
    }

    /// Extract all text from the body.
    pub fn text(&self) -> String {
        self.content
            .iter()
            .map(|block| match block {
                BlockContent::Paragraph(p) => p.text(),
                BlockContent::Table(t) => t.text(),
            })
            .collect::<Vec<_>>()
            .join("\n")
    }

    /// Get section properties (page size, margins, etc.).
    pub fn section_properties(&self) -> Option<&SectionProperties> {
        self.section_properties.as_ref()
    }

    /// Get mutable reference to section properties.
    pub fn section_properties_mut(&mut self) -> &mut Option<SectionProperties> {
        &mut self.section_properties
    }

    /// Set section properties.
    pub fn set_section_properties(&mut self, props: SectionProperties) {
        self.section_properties = Some(props);
    }
}

/// A table in the document.
///
/// Corresponds to the `<w:tbl>` element.
#[derive(Debug, Clone, Default)]
pub struct Table {
    /// Table properties.
    properties: Option<TableProperties>,
    /// Grid column definitions (defines the column structure).
    grid_columns: Vec<GridColumn>,
    /// Rows in the table.
    rows: Vec<Row>,
    /// Unknown child elements preserved for round-trip fidelity.
    /// Stored with original position index for correct ordering during serialization.
    pub unknown_children: Vec<PositionedNode>,
}

impl Table {
    /// Create an empty table.
    pub fn new() -> Self {
        Self::default()
    }

    /// Get rows in the table.
    pub fn rows(&self) -> &[Row] {
        &self.rows
    }

    /// Get mutable reference to rows.
    pub fn rows_mut(&mut self) -> &mut Vec<Row> {
        &mut self.rows
    }

    /// Add a new row to the table.
    pub fn add_row(&mut self) -> &mut Row {
        self.rows.push(Row::new());
        self.rows.last_mut().unwrap()
    }

    /// Get the number of rows.
    pub fn row_count(&self) -> usize {
        self.rows.len()
    }

    /// Get the number of columns (based on first row).
    pub fn column_count(&self) -> usize {
        self.rows.first().map(|r| r.cells().len()).unwrap_or(0)
    }

    /// Get table properties.
    pub fn properties(&self) -> Option<&TableProperties> {
        self.properties.as_ref()
    }

    /// Get mutable reference to table properties.
    pub fn properties_mut(&mut self) -> &mut Option<TableProperties> {
        &mut self.properties
    }

    /// Get grid column definitions.
    pub fn grid_columns(&self) -> &[GridColumn] {
        &self.grid_columns
    }

    /// Get mutable reference to grid columns.
    pub fn grid_columns_mut(&mut self) -> &mut Vec<GridColumn> {
        &mut self.grid_columns
    }

    /// Extract all text from the table.
    pub fn text(&self) -> String {
        self.rows
            .iter()
            .map(|r| r.text())
            .collect::<Vec<_>>()
            .join("\n")
    }
}

/// A row in a table.
///
/// Corresponds to the `<w:tr>` element.
#[derive(Debug, Clone, Default)]
pub struct Row {
    /// Row properties.
    properties: Option<RowProperties>,
    /// Cells in the row.
    cells: Vec<Cell>,
    /// Unknown child elements preserved for round-trip fidelity.
    /// Stored with original position index for correct ordering during serialization.
    pub unknown_children: Vec<PositionedNode>,
}

impl Row {
    /// Create an empty row.
    pub fn new() -> Self {
        Self::default()
    }

    /// Get cells in the row.
    pub fn cells(&self) -> &[Cell] {
        &self.cells
    }

    /// Get mutable reference to cells.
    pub fn cells_mut(&mut self) -> &mut Vec<Cell> {
        &mut self.cells
    }

    /// Add a new cell to the row.
    pub fn add_cell(&mut self) -> &mut Cell {
        self.cells.push(Cell::new());
        self.cells.last_mut().unwrap()
    }

    /// Get row properties.
    pub fn properties(&self) -> Option<&RowProperties> {
        self.properties.as_ref()
    }

    /// Get mutable reference to row properties.
    pub fn properties_mut(&mut self) -> &mut Option<RowProperties> {
        &mut self.properties
    }

    /// Check if this row is a header row.
    pub fn is_header(&self) -> bool {
        self.properties
            .as_ref()
            .map(|p| p.is_header)
            .unwrap_or(false)
    }

    /// Extract all text from the row (cells separated by tabs).
    pub fn text(&self) -> String {
        self.cells
            .iter()
            .map(|c| c.text())
            .collect::<Vec<_>>()
            .join("\t")
    }
}

/// A cell in a table row.
///
/// Corresponds to the `<w:tc>` element.
#[derive(Debug, Clone, Default)]
pub struct Cell {
    /// Paragraphs in the cell.
    paragraphs: Vec<Paragraph>,
    /// Cell properties (borders, shading, merge, etc.).
    properties: Option<CellProperties>,
    /// Unknown child elements preserved for round-trip fidelity.
    /// Stored with original position index for correct ordering during serialization.
    pub unknown_children: Vec<PositionedNode>,
}

impl Cell {
    /// Create an empty cell.
    pub fn new() -> Self {
        Self::default()
    }

    /// Get paragraphs in the cell.
    pub fn paragraphs(&self) -> &[Paragraph] {
        &self.paragraphs
    }

    /// Get mutable reference to paragraphs.
    pub fn paragraphs_mut(&mut self) -> &mut Vec<Paragraph> {
        &mut self.paragraphs
    }

    /// Add a paragraph to the cell.
    pub fn add_paragraph(&mut self) -> &mut Paragraph {
        self.paragraphs.push(Paragraph::new());
        self.paragraphs.last_mut().unwrap()
    }

    /// Get cell properties.
    pub fn properties(&self) -> Option<&CellProperties> {
        self.properties.as_ref()
    }

    /// Get mutable reference to cell properties.
    pub fn properties_mut(&mut self) -> &mut Option<CellProperties> {
        &mut self.properties
    }

    /// Set cell properties.
    pub fn set_properties(&mut self, props: CellProperties) {
        self.properties = Some(props);
    }

    /// Extract all text from the cell.
    pub fn text(&self) -> String {
        self.paragraphs
            .iter()
            .map(|p| p.text())
            .collect::<Vec<_>>()
            .join(" ")
    }
}

/// A paragraph in the document.
///
/// Corresponds to the `<w:p>` element.
#[derive(Debug, Clone, Default)]
pub struct Paragraph {
    /// Content in the paragraph (runs and hyperlinks).
    content: Vec<ParagraphContent>,
    /// Paragraph properties (style, alignment, etc.).
    properties: Option<ParagraphProperties>,
    /// Unknown child elements preserved for round-trip fidelity.
    /// Stored with original position index for correct ordering during serialization.
    pub unknown_children: Vec<PositionedNode>,
}

/// Content that can appear inside a paragraph.
#[derive(Debug, Clone)]
pub enum ParagraphContent {
    /// A text run.
    Run(Run),
    /// A hyperlink containing runs.
    Hyperlink(Hyperlink),
}

/// A hyperlink in the document.
///
/// Corresponds to the `<w:hyperlink>` element. Contains runs that form
/// the clickable link text.
#[derive(Debug, Clone, Default)]
pub struct Hyperlink {
    /// Runs inside the hyperlink.
    runs: Vec<Run>,
    /// Relationship ID for external URLs.
    rel_id: Option<String>,
    /// Anchor for internal document links (bookmarks).
    anchor: Option<String>,
    /// Unknown child elements preserved for round-trip fidelity.
    /// Stored with original position index for correct ordering during serialization.
    pub unknown_children: Vec<PositionedNode>,
    /// Unknown attributes preserved for round-trip fidelity.
    /// Stored with original position index for correct ordering during serialization.
    pub unknown_attrs: Vec<PositionedAttr>,
}

impl Paragraph {
    /// Create an empty paragraph.
    pub fn new() -> Self {
        Self::default()
    }

    /// Get all runs in the paragraph (flattened, including those in hyperlinks).
    pub fn runs(&self) -> Vec<&Run> {
        let mut runs = Vec::new();
        for content in &self.content {
            match content {
                ParagraphContent::Run(run) => runs.push(run),
                ParagraphContent::Hyperlink(link) => {
                    for run in &link.runs {
                        runs.push(run);
                    }
                }
            }
        }
        runs
    }

    /// Get paragraph content (runs and hyperlinks).
    pub fn content(&self) -> &[ParagraphContent] {
        &self.content
    }

    /// Get mutable reference to paragraph content.
    pub fn content_mut(&mut self) -> &mut Vec<ParagraphContent> {
        &mut self.content
    }

    /// Add a new run to the paragraph.
    pub fn add_run(&mut self) -> &mut Run {
        self.content.push(ParagraphContent::Run(Run::new()));
        match self.content.last_mut() {
            Some(ParagraphContent::Run(run)) => run,
            _ => unreachable!(),
        }
    }

    /// Add a new hyperlink to the paragraph.
    pub fn add_hyperlink(&mut self) -> &mut Hyperlink {
        self.content
            .push(ParagraphContent::Hyperlink(Hyperlink::new()));
        match self.content.last_mut() {
            Some(ParagraphContent::Hyperlink(link)) => link,
            _ => unreachable!(),
        }
    }

    /// Get all hyperlinks in the paragraph.
    pub fn hyperlinks(&self) -> impl Iterator<Item = &Hyperlink> {
        self.content.iter().filter_map(|c| match c {
            ParagraphContent::Hyperlink(link) => Some(link),
            _ => None,
        })
    }

    /// Get paragraph properties.
    pub fn properties(&self) -> Option<&ParagraphProperties> {
        self.properties.as_ref()
    }

    /// Set paragraph properties.
    pub fn set_properties(&mut self, props: ParagraphProperties) {
        self.properties = Some(props);
    }

    /// Extract all text from the paragraph.
    pub fn text(&self) -> String {
        self.content
            .iter()
            .map(|c| match c {
                ParagraphContent::Run(run) => run.text().to_string(),
                ParagraphContent::Hyperlink(link) => link.text(),
            })
            .collect()
    }
}

impl Hyperlink {
    /// Create an empty hyperlink.
    pub fn new() -> Self {
        Self::default()
    }

    /// Get runs inside the hyperlink.
    pub fn runs(&self) -> &[Run] {
        &self.runs
    }

    /// Get mutable reference to runs.
    pub fn runs_mut(&mut self) -> &mut Vec<Run> {
        &mut self.runs
    }

    /// Add a new run to the hyperlink.
    pub fn add_run(&mut self) -> &mut Run {
        self.runs.push(Run::new());
        self.runs.last_mut().unwrap()
    }

    /// Get the relationship ID (for external URLs).
    pub fn rel_id(&self) -> Option<&str> {
        self.rel_id.as_deref()
    }

    /// Set the relationship ID.
    pub fn set_rel_id(&mut self, rel_id: impl Into<String>) -> &mut Self {
        self.rel_id = Some(rel_id.into());
        self
    }

    /// Get the anchor (for internal document links).
    pub fn anchor(&self) -> Option<&str> {
        self.anchor.as_deref()
    }

    /// Set the anchor for internal links.
    pub fn set_anchor(&mut self, anchor: impl Into<String>) -> &mut Self {
        self.anchor = Some(anchor.into());
        self
    }

    /// Check if this is an external hyperlink.
    pub fn is_external(&self) -> bool {
        self.rel_id.is_some()
    }

    /// Extract all text from the hyperlink.
    pub fn text(&self) -> String {
        self.runs.iter().map(|r| r.text()).collect()
    }
}

/// Paragraph alignment (justification).
///
/// Corresponds to the `<w:jc>` element values.
/// ECMA-376 Part 1, Section 17.18.44 (ST_Jc).
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub enum Alignment {
    /// Align left (default for LTR).
    #[default]
    Left,
    /// Center alignment.
    Center,
    /// Align right.
    Right,
    /// Justified (both edges aligned).
    Both,
    /// Distributed alignment (for CJK).
    Distribute,
}

impl Alignment {
    /// Parse from the w:val attribute value.
    pub fn parse(s: &str) -> Self {
        match s {
            "left" | "start" => Alignment::Left,
            "center" => Alignment::Center,
            "right" | "end" => Alignment::Right,
            "both" | "justify" => Alignment::Both,
            "distribute" => Alignment::Distribute,
            _ => Alignment::Left,
        }
    }

    /// Convert to the w:val attribute value.
    pub fn as_str(&self) -> &'static str {
        match self {
            Alignment::Left => "left",
            Alignment::Center => "center",
            Alignment::Right => "right",
            Alignment::Both => "both",
            Alignment::Distribute => "distribute",
        }
    }
}

/// Section properties defining page layout.
///
/// Corresponds to the `<w:sectPr>` element.
/// ECMA-376 Part 1, Section 17.6.17 (sectPr).
#[derive(Debug, Clone, Default)]
pub struct SectionProperties {
    /// Page size (width and height).
    pub page_size: Option<PageSize>,
    /// Page margins.
    pub margins: Option<PageMargins>,
    /// Unknown child elements preserved for round-trip fidelity.
    pub unknown_children: Vec<PositionedNode>,
}

/// Page size definition.
///
/// Corresponds to the `<w:pgSz>` element.
/// ECMA-376 Part 1, Section 17.6.13 (pgSz).
#[derive(Debug, Clone, Copy)]
pub struct PageSize {
    /// Page width in twips (1/20 of a point, 1440 twips = 1 inch).
    pub width: u32,
    /// Page height in twips.
    pub height: u32,
    /// Page orientation.
    pub orientation: PageOrientation,
}

impl Default for PageSize {
    fn default() -> Self {
        // US Letter: 8.5" x 11" = 12240 x 15840 twips
        Self {
            width: 12240,
            height: 15840,
            orientation: PageOrientation::Portrait,
        }
    }
}

/// Page orientation.
///
/// ECMA-376 Part 1, Section 17.18.65 (ST_PageOrientation).
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub enum PageOrientation {
    /// Portrait orientation (default).
    #[default]
    Portrait,
    /// Landscape orientation.
    Landscape,
}

impl PageOrientation {
    /// Parse from the w:orient attribute value.
    pub fn parse(s: &str) -> Self {
        match s {
            "landscape" => PageOrientation::Landscape,
            _ => PageOrientation::Portrait,
        }
    }

    /// Convert to the w:orient attribute value.
    pub fn as_str(&self) -> &'static str {
        match self {
            PageOrientation::Portrait => "portrait",
            PageOrientation::Landscape => "landscape",
        }
    }
}

/// Page margins.
///
/// Corresponds to the `<w:pgMar>` element.
/// ECMA-376 Part 1, Section 17.6.11 (pgMar).
#[derive(Debug, Clone, Copy)]
pub struct PageMargins {
    /// Top margin in twips.
    pub top: i32,
    /// Bottom margin in twips.
    pub bottom: i32,
    /// Left margin in twips.
    pub left: u32,
    /// Right margin in twips.
    pub right: u32,
    /// Header distance from page edge in twips.
    pub header: Option<u32>,
    /// Footer distance from page edge in twips.
    pub footer: Option<u32>,
    /// Gutter margin in twips.
    pub gutter: Option<u32>,
}

impl Default for PageMargins {
    fn default() -> Self {
        // Standard 1" margins = 1440 twips
        Self {
            top: 1440,
            bottom: 1440,
            left: 1440,
            right: 1440,
            header: None,
            footer: None,
            gutter: None,
        }
    }
}

/// Properties of a paragraph.
///
/// Corresponds to the `<w:pPr>` element.
#[derive(Debug, Clone, Default)]
pub struct ParagraphProperties {
    /// Style ID reference.
    pub style: Option<String>,
    /// Numbering/list properties.
    pub numbering: Option<NumberingProperties>,
    /// Paragraph alignment/justification.
    pub alignment: Option<Alignment>,
    /// Spacing before the paragraph in twips (1/20 of a point).
    pub spacing_before: Option<u32>,
    /// Spacing after the paragraph in twips.
    pub spacing_after: Option<u32>,
    /// Line spacing in twips (or 240ths of a line if line_rule is Auto).
    pub spacing_line: Option<u32>,
    /// Left indentation in twips.
    pub indent_left: Option<i32>,
    /// Right indentation in twips.
    pub indent_right: Option<i32>,
    /// First line indentation in twips (positive for indent, use indent_hanging for hanging).
    pub indent_first_line: Option<i32>,
    /// Hanging indentation in twips (overrides indent_first_line if set).
    pub indent_hanging: Option<u32>,
    /// Paragraph borders.
    pub borders: Option<ParagraphBorders>,
    /// Paragraph shading/background.
    pub shading: Option<CellShading>,
    /// Outline level (0-9, used for TOC and document map).
    pub outline_level: Option<u8>,
    /// Keep paragraph with next paragraph on same page.
    pub keep_next: bool,
    /// Keep all lines of paragraph on same page.
    pub keep_lines: bool,
    /// Start paragraph on new page.
    pub page_break_before: bool,
    /// Enable widow/orphan control (prevent single lines at top/bottom of page).
    pub widow_control: Option<bool>,
    /// Unknown child elements preserved for round-trip fidelity.
    /// Stored with original position index for correct ordering during serialization.
    pub unknown_children: Vec<PositionedNode>,
}

/// Paragraph borders.
///
/// Corresponds to the `<w:pBdr>` element.
/// ECMA-376 Part 1, Section 17.3.1.24 (pBdr).
#[derive(Debug, Clone, Default)]
pub struct ParagraphBorders {
    /// Top border.
    pub top: Option<Border>,
    /// Bottom border.
    pub bottom: Option<Border>,
    /// Left border.
    pub left: Option<Border>,
    /// Right border.
    pub right: Option<Border>,
    /// Border between this and previous paragraph (if identical borders).
    pub between: Option<Border>,
    /// Border around paragraph (shorthand for all sides).
    pub bar: Option<Border>,
}

/// Numbering properties for a paragraph.
///
/// Corresponds to the `<w:numPr>` element. References a numbering definition
/// and specifies the indentation level.
#[derive(Debug, Clone)]
pub struct NumberingProperties {
    /// Numbering definition ID (references an entry in numbering.xml).
    pub num_id: u32,
    /// Indentation level (0-8). 0 is the first level.
    pub ilvl: u32,
}

/// Properties of a table.
///
/// Corresponds to the `<w:tblPr>` element.
/// ECMA-376 Part 1, Section 17.4.59 (tblPr).
#[derive(Debug, Clone, Default)]
pub struct TableProperties {
    /// Table width.
    pub width: Option<TableWidth>,
    /// Table justification/alignment.
    pub justification: Option<TableJustification>,
    /// Table indentation from leading margin.
    pub indent: Option<i32>,
    /// Table borders.
    pub borders: Option<TableBorders>,
    /// Table shading/background.
    pub shading: Option<CellShading>,
    /// Table layout algorithm.
    pub layout: Option<TableLayout>,
    /// Unknown child elements preserved for round-trip fidelity.
    pub unknown_children: Vec<PositionedNode>,
}

/// Table width specification.
///
/// Corresponds to the `<w:tblW>` element.
#[derive(Debug, Clone, Copy)]
pub struct TableWidth {
    /// Width value (interpretation depends on width_type).
    pub width: i32,
    /// Width type.
    pub width_type: WidthType,
}

/// Table justification/alignment.
///
/// ECMA-376 Part 1, Section 17.18.45 (ST_Jc) for table alignment.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub enum TableJustification {
    /// Left aligned (default).
    #[default]
    Left,
    /// Center aligned.
    Center,
    /// Right aligned.
    Right,
}

impl TableJustification {
    /// Parse from the w:val attribute value.
    pub fn parse(s: &str) -> Self {
        match s {
            "center" => Self::Center,
            "right" | "end" => Self::Right,
            _ => Self::Left,
        }
    }

    /// Convert to the w:val attribute value.
    pub fn as_str(&self) -> &'static str {
        match self {
            Self::Left => "left",
            Self::Center => "center",
            Self::Right => "right",
        }
    }
}

/// Table layout algorithm.
///
/// ECMA-376 Part 1, Section 17.18.87 (ST_TblLayoutType).
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub enum TableLayout {
    /// Autofit layout (default) - columns resize based on content.
    #[default]
    Autofit,
    /// Fixed layout - columns maintain specified widths.
    Fixed,
}

impl TableLayout {
    /// Parse from the w:type attribute value.
    pub fn parse(s: &str) -> Self {
        match s {
            "fixed" => Self::Fixed,
            _ => Self::Autofit,
        }
    }

    /// Convert to the w:type attribute value.
    pub fn as_str(&self) -> &'static str {
        match self {
            Self::Autofit => "autofit",
            Self::Fixed => "fixed",
        }
    }
}

/// Table borders.
///
/// Corresponds to the `<w:tblBorders>` element.
#[derive(Debug, Clone, Default)]
pub struct TableBorders {
    /// Top border.
    pub top: Option<Border>,
    /// Bottom border.
    pub bottom: Option<Border>,
    /// Left (start) border.
    pub left: Option<Border>,
    /// Right (end) border.
    pub right: Option<Border>,
    /// Inside horizontal borders (between rows).
    pub inside_h: Option<Border>,
    /// Inside vertical borders (between columns).
    pub inside_v: Option<Border>,
}

/// Table grid column width.
///
/// Corresponds to the `<w:gridCol>` element within `<w:tblGrid>`.
#[derive(Debug, Clone, Copy)]
pub struct GridColumn {
    /// Column width in twips.
    pub width: u32,
}

/// Properties of a table row.
///
/// Corresponds to the `<w:trPr>` element.
/// ECMA-376 Part 1, Section 17.4.78 (trPr).
#[derive(Debug, Clone, Default)]
pub struct RowProperties {
    /// Row height specification.
    pub height: Option<RowHeight>,
    /// Whether this row should be repeated as a header row on each page.
    pub is_header: bool,
    /// Whether this row can be split across pages.
    pub cant_split: bool,
    /// Unknown child elements preserved for round-trip fidelity.
    pub unknown_children: Vec<PositionedNode>,
}

/// Row height specification.
///
/// Corresponds to the `<w:trHeight>` element.
#[derive(Debug, Clone, Copy)]
pub struct RowHeight {
    /// Height value in twips.
    pub value: u32,
    /// Height rule.
    pub rule: HeightRule,
}

/// Height rule for table rows.
///
/// ECMA-376 Part 1, Section 17.18.37 (ST_HeightRule).
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub enum HeightRule {
    /// Automatically determined height based on content (default).
    #[default]
    Auto,
    /// Exact height (content may be clipped).
    Exact,
    /// Minimum height (row can grow if needed).
    AtLeast,
}

impl HeightRule {
    /// Parse from the w:hRule attribute value.
    pub fn parse(s: &str) -> Self {
        match s {
            "exact" => Self::Exact,
            "atLeast" => Self::AtLeast,
            _ => Self::Auto,
        }
    }

    /// Convert to the w:hRule attribute value.
    pub fn as_str(&self) -> &'static str {
        match self {
            Self::Auto => "auto",
            Self::Exact => "exact",
            Self::AtLeast => "atLeast",
        }
    }
}

/// Properties of a table cell.
///
/// Corresponds to the `<w:tcPr>` element.
/// ECMA-376 Part 1, Section 17.4.66 (tcPr).
#[derive(Debug, Clone, Default)]
pub struct CellProperties {
    /// Cell width in twips.
    pub width: Option<CellWidth>,
    /// Cell borders.
    pub borders: Option<CellBorders>,
    /// Cell shading/background.
    pub shading: Option<CellShading>,
    /// Horizontal span (number of grid columns this cell covers).
    pub grid_span: Option<u32>,
    /// Vertical merge type.
    pub vertical_merge: Option<VerticalMerge>,
    /// Vertical alignment of content within the cell.
    pub vertical_align: Option<CellVerticalAlign>,
    /// Unknown child elements preserved for round-trip fidelity.
    pub unknown_children: Vec<PositionedNode>,
}

/// Cell width specification.
///
/// Corresponds to the `<w:tcW>` element.
#[derive(Debug, Clone, Copy)]
pub struct CellWidth {
    /// Width value (interpretation depends on width_type).
    pub width: u32,
    /// Width type.
    pub width_type: WidthType,
}

/// Width measurement type.
///
/// ECMA-376 Part 1, Section 17.18.107 (ST_TblWidth).
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub enum WidthType {
    /// Width in twips.
    #[default]
    Dxa,
    /// Width in fiftieths of a percent (pct).
    Pct,
    /// Automatically determined width.
    Auto,
    /// No width specified.
    Nil,
}

impl WidthType {
    /// Parse from the w:type attribute value.
    pub fn parse(s: &str) -> Self {
        match s {
            "dxa" => Self::Dxa,
            "pct" => Self::Pct,
            "auto" => Self::Auto,
            "nil" => Self::Nil,
            _ => Self::Dxa,
        }
    }

    /// Convert to the w:type attribute value.
    pub fn as_str(&self) -> &'static str {
        match self {
            Self::Dxa => "dxa",
            Self::Pct => "pct",
            Self::Auto => "auto",
            Self::Nil => "nil",
        }
    }
}

/// Cell borders.
///
/// Corresponds to the `<w:tcBorders>` element.
#[derive(Debug, Clone, Default)]
pub struct CellBorders {
    /// Top border.
    pub top: Option<Border>,
    /// Bottom border.
    pub bottom: Option<Border>,
    /// Left (start) border.
    pub left: Option<Border>,
    /// Right (end) border.
    pub right: Option<Border>,
    /// Inside horizontal border (for merged cells).
    pub inside_h: Option<Border>,
    /// Inside vertical border (for merged cells).
    pub inside_v: Option<Border>,
}

/// A border definition.
///
/// ECMA-376 Part 1, Section 17.3.4 (border).
#[derive(Debug, Clone)]
pub struct Border {
    /// Border style.
    pub style: BorderStyle,
    /// Border width in eighths of a point.
    pub size: Option<u32>,
    /// Border color (hex RGB, e.g., "FF0000" for red).
    pub color: Option<String>,
    /// Space between border and content in points.
    pub space: Option<u32>,
}

/// Border style.
///
/// ECMA-376 Part 1, Section 17.18.2 (ST_Border).
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub enum BorderStyle {
    /// No border.
    #[default]
    Nil,
    /// No border (same as nil).
    None,
    /// Single line border.
    Single,
    /// Thick single line.
    Thick,
    /// Double line border.
    Double,
    /// Dotted border.
    Dotted,
    /// Dashed border.
    Dashed,
    /// Dot-dash pattern.
    DotDash,
    /// Dot-dot-dash pattern.
    DotDotDash,
    /// Triple line border.
    Triple,
    /// Thin-thick small gap.
    ThinThickSmallGap,
    /// Thick-thin small gap.
    ThickThinSmallGap,
    /// Thin-thick-thin small gap.
    ThinThickThinSmallGap,
    /// Thin-thick medium gap.
    ThinThickMediumGap,
    /// Thick-thin medium gap.
    ThickThinMediumGap,
    /// Thin-thick-thin medium gap.
    ThinThickThinMediumGap,
    /// Thin-thick large gap.
    ThinThickLargeGap,
    /// Thick-thin large gap.
    ThickThinLargeGap,
    /// Thin-thick-thin large gap.
    ThinThickThinLargeGap,
    /// Wavy border.
    Wave,
    /// Double wavy border.
    DoubleWave,
    /// Dash small gap.
    DashSmallGap,
    /// Dash-dot stroked.
    DashDotStroked,
    /// 3D emboss effect.
    ThreeDEmboss,
    /// 3D engrave effect.
    ThreeDEngrave,
    /// Outset border.
    Outset,
    /// Inset border.
    Inset,
}

impl BorderStyle {
    /// Parse from the w:val attribute value.
    pub fn parse(s: &str) -> Self {
        match s {
            "nil" => Self::Nil,
            "none" => Self::None,
            "single" => Self::Single,
            "thick" => Self::Thick,
            "double" => Self::Double,
            "dotted" => Self::Dotted,
            "dashed" => Self::Dashed,
            "dotDash" => Self::DotDash,
            "dotDotDash" => Self::DotDotDash,
            "triple" => Self::Triple,
            "thinThickSmallGap" => Self::ThinThickSmallGap,
            "thickThinSmallGap" => Self::ThickThinSmallGap,
            "thinThickThinSmallGap" => Self::ThinThickThinSmallGap,
            "thinThickMediumGap" => Self::ThinThickMediumGap,
            "thickThinMediumGap" => Self::ThickThinMediumGap,
            "thinThickThinMediumGap" => Self::ThinThickThinMediumGap,
            "thinThickLargeGap" => Self::ThinThickLargeGap,
            "thickThinLargeGap" => Self::ThickThinLargeGap,
            "thinThickThinLargeGap" => Self::ThinThickThinLargeGap,
            "wave" => Self::Wave,
            "doubleWave" => Self::DoubleWave,
            "dashSmallGap" => Self::DashSmallGap,
            "dashDotStroked" => Self::DashDotStroked,
            "threeDEmboss" => Self::ThreeDEmboss,
            "threeDEngrave" => Self::ThreeDEngrave,
            "outset" => Self::Outset,
            "inset" => Self::Inset,
            _ => Self::Single,
        }
    }

    /// Convert to the w:val attribute value.
    pub fn as_str(&self) -> &'static str {
        match self {
            Self::Nil => "nil",
            Self::None => "none",
            Self::Single => "single",
            Self::Thick => "thick",
            Self::Double => "double",
            Self::Dotted => "dotted",
            Self::Dashed => "dashed",
            Self::DotDash => "dotDash",
            Self::DotDotDash => "dotDotDash",
            Self::Triple => "triple",
            Self::ThinThickSmallGap => "thinThickSmallGap",
            Self::ThickThinSmallGap => "thickThinSmallGap",
            Self::ThinThickThinSmallGap => "thinThickThinSmallGap",
            Self::ThinThickMediumGap => "thinThickMediumGap",
            Self::ThickThinMediumGap => "thickThinMediumGap",
            Self::ThinThickThinMediumGap => "thinThickThinMediumGap",
            Self::ThinThickLargeGap => "thinThickLargeGap",
            Self::ThickThinLargeGap => "thickThinLargeGap",
            Self::ThinThickThinLargeGap => "thinThickThinLargeGap",
            Self::Wave => "wave",
            Self::DoubleWave => "doubleWave",
            Self::DashSmallGap => "dashSmallGap",
            Self::DashDotStroked => "dashDotStroked",
            Self::ThreeDEmboss => "threeDEmboss",
            Self::ThreeDEngrave => "threeDEngrave",
            Self::Outset => "outset",
            Self::Inset => "inset",
        }
    }
}

/// Cell shading/background.
///
/// Corresponds to the `<w:shd>` element.
#[derive(Debug, Clone, Default)]
pub struct CellShading {
    /// Fill color (hex RGB, e.g., "FFFF00" for yellow).
    pub fill: Option<String>,
    /// Pattern color.
    pub color: Option<String>,
    /// Shading pattern.
    pub pattern: Option<ShadingPattern>,
}

/// Shading pattern type.
///
/// ECMA-376 Part 1, Section 17.18.78 (ST_Shd).
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub enum ShadingPattern {
    /// No pattern (solid fill).
    #[default]
    Clear,
    /// Solid fill.
    Solid,
    /// Horizontal stripe pattern.
    HorzStripe,
    /// Vertical stripe pattern.
    VertStripe,
    /// Reverse diagonal stripe.
    ReverseDiagStripe,
    /// Diagonal stripe.
    DiagStripe,
    /// Horizontal cross pattern.
    HorzCross,
    /// Diagonal cross pattern.
    DiagCross,
    /// Thin horizontal stripe.
    ThinHorzStripe,
    /// Thin vertical stripe.
    ThinVertStripe,
    /// Thin reverse diagonal stripe.
    ThinReverseDiagStripe,
    /// Thin diagonal stripe.
    ThinDiagStripe,
    /// Thin horizontal cross.
    ThinHorzCross,
    /// Thin diagonal cross.
    ThinDiagCross,
    /// 5% pattern.
    Pct5,
    /// 10% pattern.
    Pct10,
    /// 12.5% pattern.
    Pct12,
    /// 15% pattern.
    Pct15,
    /// 20% pattern.
    Pct20,
    /// 25% pattern.
    Pct25,
    /// 30% pattern.
    Pct30,
    /// 35% pattern.
    Pct35,
    /// 37.5% pattern.
    Pct37,
    /// 40% pattern.
    Pct40,
    /// 45% pattern.
    Pct45,
    /// 50% pattern.
    Pct50,
    /// 55% pattern.
    Pct55,
    /// 60% pattern.
    Pct60,
    /// 62.5% pattern.
    Pct62,
    /// 65% pattern.
    Pct65,
    /// 70% pattern.
    Pct70,
    /// 75% pattern.
    Pct75,
    /// 80% pattern.
    Pct80,
    /// 85% pattern.
    Pct85,
    /// 87.5% pattern.
    Pct87,
    /// 90% pattern.
    Pct90,
    /// 95% pattern.
    Pct95,
}

impl ShadingPattern {
    /// Parse from the w:val attribute value.
    pub fn parse(s: &str) -> Self {
        match s {
            "clear" => Self::Clear,
            "solid" => Self::Solid,
            "horzStripe" => Self::HorzStripe,
            "vertStripe" => Self::VertStripe,
            "reverseDiagStripe" => Self::ReverseDiagStripe,
            "diagStripe" => Self::DiagStripe,
            "horzCross" => Self::HorzCross,
            "diagCross" => Self::DiagCross,
            "thinHorzStripe" => Self::ThinHorzStripe,
            "thinVertStripe" => Self::ThinVertStripe,
            "thinReverseDiagStripe" => Self::ThinReverseDiagStripe,
            "thinDiagStripe" => Self::ThinDiagStripe,
            "thinHorzCross" => Self::ThinHorzCross,
            "thinDiagCross" => Self::ThinDiagCross,
            "pct5" => Self::Pct5,
            "pct10" => Self::Pct10,
            "pct12" => Self::Pct12,
            "pct15" => Self::Pct15,
            "pct20" => Self::Pct20,
            "pct25" => Self::Pct25,
            "pct30" => Self::Pct30,
            "pct35" => Self::Pct35,
            "pct37" => Self::Pct37,
            "pct40" => Self::Pct40,
            "pct45" => Self::Pct45,
            "pct50" => Self::Pct50,
            "pct55" => Self::Pct55,
            "pct60" => Self::Pct60,
            "pct62" => Self::Pct62,
            "pct65" => Self::Pct65,
            "pct70" => Self::Pct70,
            "pct75" => Self::Pct75,
            "pct80" => Self::Pct80,
            "pct85" => Self::Pct85,
            "pct87" => Self::Pct87,
            "pct90" => Self::Pct90,
            "pct95" => Self::Pct95,
            _ => Self::Clear,
        }
    }
}

/// Vertical merge state for a cell.
///
/// ECMA-376 Part 1, Section 17.18.99 (ST_Merge).
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum VerticalMerge {
    /// Start of a vertically merged region.
    Restart,
    /// Continuation of a vertically merged region.
    Continue,
}

impl VerticalMerge {
    /// Parse from the w:val attribute value.
    pub fn parse(s: &str) -> Option<Self> {
        match s {
            "restart" => Some(Self::Restart),
            "continue" => Some(Self::Continue),
            _ => None,
        }
    }

    /// Convert to the w:val attribute value.
    pub fn as_str(&self) -> &'static str {
        match self {
            Self::Restart => "restart",
            Self::Continue => "continue",
        }
    }
}

/// Vertical alignment of cell content.
///
/// ECMA-376 Part 1, Section 17.18.101 (ST_VerticalJc).
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub enum CellVerticalAlign {
    /// Align to top.
    #[default]
    Top,
    /// Align to center.
    Center,
    /// Align to bottom.
    Bottom,
}

impl CellVerticalAlign {
    /// Parse from the w:val attribute value.
    pub fn parse(s: &str) -> Self {
        match s {
            "top" => Self::Top,
            "center" => Self::Center,
            "bottom" => Self::Bottom,
            "both" => Self::Center, // Treat justify as center
            _ => Self::Top,
        }
    }

    /// Convert to the w:val attribute value.
    pub fn as_str(&self) -> &'static str {
        match self {
            Self::Top => "top",
            Self::Center => "center",
            Self::Bottom => "bottom",
        }
    }
}

/// A text run in the document.
///
/// Corresponds to the `<w:r>` element. A run is a contiguous range of text
/// with the same formatting. Runs may also contain drawings (images).
#[derive(Debug, Clone, Default)]
pub struct Run {
    /// Text content in the run.
    text: String,
    /// Run properties (bold, italic, etc.).
    properties: Option<RunProperties>,
    /// Drawings (images) in the run.
    drawings: Vec<Drawing>,
    /// Whether this run contains a page break.
    page_break: bool,
    /// Unknown child elements preserved for round-trip fidelity.
    /// Stored with original position index for correct ordering during serialization.
    pub unknown_children: Vec<PositionedNode>,
    /// Unknown attributes preserved for round-trip fidelity.
    /// Stored with original position index for correct ordering during serialization.
    pub unknown_attrs: Vec<PositionedAttr>,
}

/// A drawing element containing images.
///
/// Corresponds to the `<w:drawing>` element. Currently only inline drawings
/// are supported (not floating/anchored).
#[derive(Debug, Clone, Default)]
pub struct Drawing {
    /// Images in this drawing.
    images: Vec<InlineImage>,
    /// Unknown child elements preserved for round-trip fidelity.
    /// Stored with original position index for correct ordering during serialization.
    pub unknown_children: Vec<PositionedNode>,
}

/// An inline image in a drawing.
///
/// Represents an image embedded in the document via DrawingML.
/// References image data through a relationship ID.
#[derive(Debug, Clone)]
pub struct InlineImage {
    /// Relationship ID referencing the image file (e.g., "rId4").
    rel_id: String,
    /// Width in EMUs (English Metric Units). 914400 EMUs = 1 inch.
    width_emu: Option<i64>,
    /// Height in EMUs.
    height_emu: Option<i64>,
    /// Optional description/alt text for the image.
    description: Option<String>,
    /// Unknown child elements preserved for round-trip fidelity.
    /// Stored with original position index for correct ordering during serialization.
    pub unknown_children: Vec<PositionedNode>,
}

/// Image data loaded from the package.
#[derive(Debug, Clone)]
pub struct ImageData {
    /// MIME content type (e.g., "image/png", "image/jpeg").
    pub content_type: String,
    /// Raw image bytes.
    pub data: Vec<u8>,
}

impl Run {
    /// Create an empty run.
    pub fn new() -> Self {
        Self::default()
    }

    /// Get the text content.
    pub fn text(&self) -> &str {
        &self.text
    }

    /// Set the text content.
    pub fn set_text(&mut self, text: impl Into<String>) {
        self.text = text.into();
    }

    /// Get run properties.
    pub fn properties(&self) -> Option<&RunProperties> {
        self.properties.as_ref()
    }

    /// Set run properties.
    pub fn set_properties(&mut self, props: RunProperties) {
        self.properties = Some(props);
    }

    /// Check if the run is bold.
    pub fn is_bold(&self) -> bool {
        self.properties.as_ref().is_some_and(|p| p.bold)
    }

    /// Check if the run is italic.
    pub fn is_italic(&self) -> bool {
        self.properties.as_ref().is_some_and(|p| p.italic)
    }

    /// Check if the run is underlined.
    pub fn is_underline(&self) -> bool {
        self.properties
            .as_ref()
            .is_some_and(|p| p.underline.is_some())
    }

    /// Get the underline style, if any.
    pub fn underline_style(&self) -> Option<UnderlineStyle> {
        self.properties.as_ref().and_then(|p| p.underline)
    }

    /// Check if the run has strikethrough.
    pub fn is_strikethrough(&self) -> bool {
        self.properties.as_ref().is_some_and(|p| p.strike)
    }

    /// Check if the run has double strikethrough.
    pub fn is_double_strikethrough(&self) -> bool {
        self.properties.as_ref().is_some_and(|p| p.double_strike)
    }

    /// Get the highlight color, if any.
    pub fn highlight(&self) -> Option<HighlightColor> {
        self.properties.as_ref().and_then(|p| p.highlight)
    }

    /// Get the vertical alignment (superscript/subscript), if any.
    pub fn vertical_align(&self) -> Option<VerticalAlign> {
        self.properties.as_ref().and_then(|p| p.vertical_align)
    }

    /// Check if the run is superscript.
    pub fn is_superscript(&self) -> bool {
        self.properties
            .as_ref()
            .is_some_and(|p| p.vertical_align == Some(VerticalAlign::Superscript))
    }

    /// Check if the run is subscript.
    pub fn is_subscript(&self) -> bool {
        self.properties
            .as_ref()
            .is_some_and(|p| p.vertical_align == Some(VerticalAlign::Subscript))
    }

    /// Check if the run is all caps.
    pub fn is_all_caps(&self) -> bool {
        self.properties.as_ref().is_some_and(|p| p.all_caps)
    }

    /// Check if the run is small caps.
    pub fn is_small_caps(&self) -> bool {
        self.properties.as_ref().is_some_and(|p| p.small_caps)
    }

    /// Get drawings (images) in this run.
    pub fn drawings(&self) -> &[Drawing] {
        &self.drawings
    }

    /// Get mutable reference to drawings.
    pub fn drawings_mut(&mut self) -> &mut Vec<Drawing> {
        &mut self.drawings
    }

    /// Check if this run contains any images.
    pub fn has_images(&self) -> bool {
        self.drawings.iter().any(|d| !d.images.is_empty())
    }

    /// Check if this run contains a page break.
    pub fn has_page_break(&self) -> bool {
        self.page_break
    }

    /// Set whether this run contains a page break.
    pub fn set_page_break(&mut self, has_break: bool) -> &mut Self {
        self.page_break = has_break;
        self
    }
}

impl Drawing {
    /// Create an empty drawing.
    pub fn new() -> Self {
        Self::default()
    }

    /// Get images in this drawing.
    pub fn images(&self) -> &[InlineImage] {
        &self.images
    }

    /// Get mutable reference to images.
    pub fn images_mut(&mut self) -> &mut Vec<InlineImage> {
        &mut self.images
    }

    /// Add an image to this drawing.
    pub fn add_image(&mut self, rel_id: impl Into<String>) -> &mut InlineImage {
        self.images.push(InlineImage::new(rel_id));
        self.images.last_mut().unwrap()
    }
}

impl InlineImage {
    /// Create a new inline image with the given relationship ID.
    pub fn new(rel_id: impl Into<String>) -> Self {
        Self {
            rel_id: rel_id.into(),
            width_emu: None,
            height_emu: None,
            description: None,
            unknown_children: Vec::new(), // (position, node) pairs
        }
    }

    /// Get the relationship ID.
    pub fn rel_id(&self) -> &str {
        &self.rel_id
    }

    /// Get width in EMUs (914400 EMUs = 1 inch).
    pub fn width_emu(&self) -> Option<i64> {
        self.width_emu
    }

    /// Get height in EMUs.
    pub fn height_emu(&self) -> Option<i64> {
        self.height_emu
    }

    /// Get width in inches.
    pub fn width_inches(&self) -> Option<f64> {
        self.width_emu.map(|e| e as f64 / 914400.0)
    }

    /// Get height in inches.
    pub fn height_inches(&self) -> Option<f64> {
        self.height_emu.map(|e| e as f64 / 914400.0)
    }

    /// Set width in EMUs.
    pub fn set_width_emu(&mut self, emu: i64) -> &mut Self {
        self.width_emu = Some(emu);
        self
    }

    /// Set height in EMUs.
    pub fn set_height_emu(&mut self, emu: i64) -> &mut Self {
        self.height_emu = Some(emu);
        self
    }

    /// Set width in inches.
    pub fn set_width_inches(&mut self, inches: f64) -> &mut Self {
        self.width_emu = Some((inches * 914400.0) as i64);
        self
    }

    /// Set height in inches.
    pub fn set_height_inches(&mut self, inches: f64) -> &mut Self {
        self.height_emu = Some((inches * 914400.0) as i64);
        self
    }

    /// Get the description/alt text.
    pub fn description(&self) -> Option<&str> {
        self.description.as_deref()
    }

    /// Set the description/alt text.
    pub fn set_description(&mut self, desc: impl Into<String>) -> &mut Self {
        self.description = Some(desc.into());
        self
    }
}

/// Underline style for text.
///
/// Corresponds to the `w:val` attribute of `<w:u>`.
/// Reference: ECMA-376 Part 1, §17.18.99 (ST_Underline)
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub enum UnderlineStyle {
    /// Single line underline (default).
    #[default]
    Single,
    /// Double line underline.
    Double,
    /// Thick single line.
    Thick,
    /// Dotted underline.
    Dotted,
    /// Dotted heavy underline.
    DottedHeavy,
    /// Dashed underline.
    Dash,
    /// Dashed heavy underline.
    DashedHeavy,
    /// Long dashed underline.
    DashLong,
    /// Long dashed heavy underline.
    DashLongHeavy,
    /// Dot-dash underline.
    DotDash,
    /// Dot-dash heavy underline.
    DashDotHeavy,
    /// Dot-dot-dash underline.
    DotDotDash,
    /// Dot-dot-dash heavy underline.
    DashDotDotHeavy,
    /// Wavy underline.
    Wave,
    /// Wavy heavy underline.
    WavyHeavy,
    /// Double wavy underline.
    WavyDouble,
    /// Words only (spaces not underlined).
    Words,
}

impl UnderlineStyle {
    /// Parse from the w:val attribute value.
    pub fn parse(s: &str) -> Option<Self> {
        match s {
            "single" => Some(Self::Single),
            "double" => Some(Self::Double),
            "thick" => Some(Self::Thick),
            "dotted" => Some(Self::Dotted),
            "dottedHeavy" => Some(Self::DottedHeavy),
            "dash" => Some(Self::Dash),
            "dashedHeavy" => Some(Self::DashedHeavy),
            "dashLong" => Some(Self::DashLong),
            "dashLongHeavy" => Some(Self::DashLongHeavy),
            "dotDash" => Some(Self::DotDash),
            "dashDotHeavy" => Some(Self::DashDotHeavy),
            "dotDotDash" => Some(Self::DotDotDash),
            "dashDotDotHeavy" => Some(Self::DashDotDotHeavy),
            "wave" => Some(Self::Wave),
            "wavyHeavy" => Some(Self::WavyHeavy),
            "wavyDouble" => Some(Self::WavyDouble),
            "words" => Some(Self::Words),
            "none" => None,
            _ => Some(Self::Single), // Default to single for unknown values
        }
    }

    /// Convert to the w:val attribute value.
    pub fn as_str(&self) -> &'static str {
        match self {
            Self::Single => "single",
            Self::Double => "double",
            Self::Thick => "thick",
            Self::Dotted => "dotted",
            Self::DottedHeavy => "dottedHeavy",
            Self::Dash => "dash",
            Self::DashedHeavy => "dashedHeavy",
            Self::DashLong => "dashLong",
            Self::DashLongHeavy => "dashLongHeavy",
            Self::DotDash => "dotDash",
            Self::DashDotHeavy => "dashDotHeavy",
            Self::DotDotDash => "dotDotDash",
            Self::DashDotDotHeavy => "dashDotDotHeavy",
            Self::Wave => "wave",
            Self::WavyHeavy => "wavyHeavy",
            Self::WavyDouble => "wavyDouble",
            Self::Words => "words",
        }
    }
}

/// Highlight color for text.
///
/// Corresponds to the `w:val` attribute of `<w:highlight>`.
/// Reference: ECMA-376 Part 1, §17.18.40 (ST_HighlightColor)
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum HighlightColor {
    Black,
    Blue,
    Cyan,
    DarkBlue,
    DarkCyan,
    DarkGray,
    DarkGreen,
    DarkMagenta,
    DarkRed,
    DarkYellow,
    Green,
    LightGray,
    Magenta,
    Red,
    White,
    Yellow,
}

impl HighlightColor {
    /// Parse from the w:val attribute value.
    pub fn parse(s: &str) -> Option<Self> {
        match s {
            "black" => Some(Self::Black),
            "blue" => Some(Self::Blue),
            "cyan" => Some(Self::Cyan),
            "darkBlue" => Some(Self::DarkBlue),
            "darkCyan" => Some(Self::DarkCyan),
            "darkGray" => Some(Self::DarkGray),
            "darkGreen" => Some(Self::DarkGreen),
            "darkMagenta" => Some(Self::DarkMagenta),
            "darkRed" => Some(Self::DarkRed),
            "darkYellow" => Some(Self::DarkYellow),
            "green" => Some(Self::Green),
            "lightGray" => Some(Self::LightGray),
            "magenta" => Some(Self::Magenta),
            "red" => Some(Self::Red),
            "white" => Some(Self::White),
            "yellow" => Some(Self::Yellow),
            "none" => None,
            _ => None,
        }
    }

    /// Convert to the w:val attribute value.
    pub fn as_str(&self) -> &'static str {
        match self {
            Self::Black => "black",
            Self::Blue => "blue",
            Self::Cyan => "cyan",
            Self::DarkBlue => "darkBlue",
            Self::DarkCyan => "darkCyan",
            Self::DarkGray => "darkGray",
            Self::DarkGreen => "darkGreen",
            Self::DarkMagenta => "darkMagenta",
            Self::DarkRed => "darkRed",
            Self::DarkYellow => "darkYellow",
            Self::Green => "green",
            Self::LightGray => "lightGray",
            Self::Magenta => "magenta",
            Self::Red => "red",
            Self::White => "white",
            Self::Yellow => "yellow",
        }
    }
}

/// Vertical alignment for text (superscript/subscript).
///
/// Corresponds to the `w:val` attribute of `<w:vertAlign>`.
/// Reference: ECMA-376 Part 1, §17.18.96 (ST_VerticalAlignRun)
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum VerticalAlign {
    /// Normal baseline alignment.
    Baseline,
    /// Superscript.
    Superscript,
    /// Subscript.
    Subscript,
}

impl VerticalAlign {
    /// Parse from the w:val attribute value.
    pub fn parse(s: &str) -> Option<Self> {
        match s {
            "baseline" => Some(Self::Baseline),
            "superscript" => Some(Self::Superscript),
            "subscript" => Some(Self::Subscript),
            _ => None,
        }
    }

    /// Convert to the w:val attribute value.
    pub fn as_str(&self) -> &'static str {
        match self {
            Self::Baseline => "baseline",
            Self::Superscript => "superscript",
            Self::Subscript => "subscript",
        }
    }
}

/// Properties of a text run.
///
/// Corresponds to the `<w:rPr>` element.
#[derive(Debug, Clone, Default)]
pub struct RunProperties {
    /// Bold formatting.
    pub bold: bool,
    /// Italic formatting.
    pub italic: bool,
    /// Underline style. None means no underline.
    pub underline: Option<UnderlineStyle>,
    /// Strike-through formatting.
    pub strike: bool,
    /// Double strike-through formatting.
    pub double_strike: bool,
    /// Font size in half-points.
    pub size: Option<u32>,
    /// Font name.
    pub font: Option<String>,
    /// Style ID reference.
    pub style: Option<String>,
    /// Text color as hex RGB (e.g., "FF0000" for red, without # prefix).
    pub color: Option<String>,
    /// Highlight/background color.
    pub highlight: Option<HighlightColor>,
    /// Vertical alignment (superscript/subscript).
    pub vertical_align: Option<VerticalAlign>,
    /// All capitals.
    pub all_caps: bool,
    /// Small capitals.
    pub small_caps: bool,
    /// Unknown child elements preserved for round-trip fidelity.
    /// Stored with original position index for correct ordering during serialization.
    pub unknown_children: Vec<PositionedNode>,
}

// XML element names (local names)
const EL_BODY: &[u8] = b"body";
const EL_P: &[u8] = b"p";
const EL_R: &[u8] = b"r";
const EL_T: &[u8] = b"t";
const EL_PPR: &[u8] = b"pPr";
const EL_RPR: &[u8] = b"rPr";
const EL_PSTYLE: &[u8] = b"pStyle";
const EL_RSTYLE: &[u8] = b"rStyle";
const EL_B: &[u8] = b"b";
const EL_I: &[u8] = b"i";
const EL_U: &[u8] = b"u";
const EL_STRIKE: &[u8] = b"strike";
const EL_DSTRIKE: &[u8] = b"dstrike";
const EL_HIGHLIGHT: &[u8] = b"highlight";
const EL_VERT_ALIGN: &[u8] = b"vertAlign";
const EL_CAPS: &[u8] = b"caps";
const EL_SMALL_CAPS: &[u8] = b"smallCaps";
const EL_SZ: &[u8] = b"sz";
const EL_RFONTS: &[u8] = b"rFonts";
const EL_COLOR: &[u8] = b"color";
const EL_BR: &[u8] = b"br";
const EL_TAB: &[u8] = b"tab";
const EL_TBL: &[u8] = b"tbl";
const EL_TR: &[u8] = b"tr";
const EL_TC: &[u8] = b"tc";

// Table properties elements
const EL_TBL_PR: &[u8] = b"tblPr";
const EL_TBL_W: &[u8] = b"tblW";
const EL_TBL_BORDERS: &[u8] = b"tblBorders";
const EL_TBL_IND: &[u8] = b"tblInd";
const EL_TBL_LAYOUT: &[u8] = b"tblLayout";
const EL_TBL_GRID: &[u8] = b"tblGrid";
const EL_GRID_COL: &[u8] = b"gridCol";

// Row properties elements
const EL_TR_PR: &[u8] = b"trPr";
const EL_TR_HEIGHT: &[u8] = b"trHeight";
const EL_TBL_HEADER: &[u8] = b"tblHeader";
const EL_CANT_SPLIT: &[u8] = b"cantSplit";

// Cell properties elements
const EL_TC_PR: &[u8] = b"tcPr";
const EL_TC_W: &[u8] = b"tcW";
const EL_TC_BORDERS: &[u8] = b"tcBorders";
const EL_GRID_SPAN: &[u8] = b"gridSpan";
const EL_V_MERGE: &[u8] = b"vMerge";
const EL_V_ALIGN: &[u8] = b"vAlign";
const EL_SHD: &[u8] = b"shd";

// Border element names (shared for table, cell, paragraph borders)
const EL_TOP: &[u8] = b"top";
const EL_BOTTOM: &[u8] = b"bottom";
const EL_LEFT: &[u8] = b"left";
const EL_RIGHT: &[u8] = b"right";
const EL_INSIDE_H: &[u8] = b"insideH";
const EL_INSIDE_V: &[u8] = b"insideV";

// Hyperlink element
const EL_HYPERLINK: &[u8] = b"hyperlink";

// Numbering elements
const EL_NUMPR: &[u8] = b"numPr";
const EL_NUMID: &[u8] = b"numId";
const EL_ILVL: &[u8] = b"ilvl";

// Paragraph formatting elements
const EL_JC: &[u8] = b"jc";
const EL_SPACING: &[u8] = b"spacing";
const EL_IND: &[u8] = b"ind";
const EL_P_BDR: &[u8] = b"pBdr";
const EL_OUTLINE_LVL: &[u8] = b"outlineLvl";
const EL_KEEP_NEXT: &[u8] = b"keepNext";
const EL_KEEP_LINES: &[u8] = b"keepLines";
const EL_PAGE_BREAK_BEFORE: &[u8] = b"pageBreakBefore";
const EL_WIDOW_CONTROL: &[u8] = b"widowControl";
const EL_BETWEEN: &[u8] = b"between";
const EL_BAR: &[u8] = b"bar";

// Drawing element names (DrawingML)
const EL_DRAWING: &[u8] = b"drawing";
const EL_INLINE: &[u8] = b"inline";
const EL_EXTENT: &[u8] = b"extent";
const EL_DOCPR: &[u8] = b"docPr";
const EL_BLIP: &[u8] = b"blip";

// Section properties elements
const EL_SECT_PR: &[u8] = b"sectPr";
const EL_PG_SZ: &[u8] = b"pgSz";
const EL_PG_MAR: &[u8] = b"pgMar";

/// Parse a document.xml file into a Body.
fn parse_document(xml: &[u8]) -> Result<Body> {
    let mut reader = Reader::from_reader(xml);
    reader.config_mut().trim_text(false); // Preserve whitespace in text

    let mut buf = Vec::new();
    let mut body = Body::new();
    let mut in_body = false;
    let mut current_para: Option<Paragraph> = None;
    let mut current_run: Option<Run> = None;
    let mut current_ppr: Option<ParagraphProperties> = None;
    let mut current_rpr: Option<RunProperties> = None;
    let mut in_ppr = false;
    let mut in_rpr = false;
    let mut in_text = false;

    // Table parsing state
    let mut current_table: Option<Table> = None;
    let mut current_row: Option<Row> = None;
    let mut current_cell: Option<Cell> = None;

    // Hyperlink parsing state
    let mut current_hyperlink: Option<Hyperlink> = None;

    // Numbering parsing state
    let mut in_numpr = false;
    let mut current_numid: Option<u32> = None;
    let mut current_ilvl: Option<u32> = None;

    // Paragraph border parsing state
    let mut in_p_bdr = false;
    let mut current_p_borders: Option<ParagraphBorders> = None;

    // Drawing/image parsing state
    let mut current_drawing: Option<Drawing> = None;
    let mut current_image: Option<InlineImage> = None;

    // Section properties parsing state
    let mut current_sect_pr: Option<SectionProperties> = None;
    let mut in_sect_pr = false;
    let mut sect_pr_child_idx: usize = 0;

    // Table properties parsing state
    let mut current_tbl_pr: Option<TableProperties> = None;
    let mut in_tbl_pr = false;
    let mut in_tbl_borders = false;
    let mut current_tbl_borders: Option<TableBorders> = None;
    let mut _tbl_pr_child_idx: usize = 0;

    // Table grid parsing state
    let mut in_tbl_grid = false;

    // Row properties parsing state
    let mut current_tr_pr: Option<RowProperties> = None;
    let mut in_tr_pr = false;
    let mut _tr_pr_child_idx: usize = 0;

    // Cell properties parsing state
    let mut current_tc_pr: Option<CellProperties> = None;
    let mut in_tc_pr = false;
    let mut in_tc_borders = false;
    let mut current_tc_borders: Option<CellBorders> = None;
    let mut _tc_pr_child_idx: usize = 0;

    // Child position counters for round-trip ordering preservation
    let mut body_child_idx: usize = 0;
    let mut table_child_idx: usize = 0;
    let mut row_child_idx: usize = 0;
    let mut cell_child_idx: usize = 0;
    let mut para_child_idx: usize = 0;
    let mut hyperlink_child_idx: usize = 0;
    let mut run_child_idx: usize = 0;
    let mut ppr_child_idx: usize = 0;
    let mut rpr_child_idx: usize = 0;

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) => {
                let name = e.name();
                let local = local_name(name.as_ref());
                match local {
                    name if name == EL_BODY => {
                        in_body = true;
                        body_child_idx = 0;
                    }
                    name if name == EL_TBL && in_body => {
                        current_table = Some(Table::new());
                        table_child_idx = 0;
                        body_child_idx += 1;
                    }
                    name if name == EL_TBL_PR && current_table.is_some() => {
                        in_tbl_pr = true;
                        current_tbl_pr = Some(TableProperties::default());
                        _tbl_pr_child_idx = 0;
                        table_child_idx += 1;
                    }
                    name if name == EL_TBL_BORDERS && in_tbl_pr => {
                        in_tbl_borders = true;
                        current_tbl_borders = Some(TableBorders::default());
                        _tbl_pr_child_idx += 1;
                    }
                    name if name == EL_TBL_GRID && current_table.is_some() => {
                        in_tbl_grid = true;
                        table_child_idx += 1;
                    }
                    name if name == EL_TR && current_table.is_some() => {
                        current_row = Some(Row::new());
                        row_child_idx = 0;
                        table_child_idx += 1;
                    }
                    name if name == EL_TR_PR && current_row.is_some() => {
                        in_tr_pr = true;
                        current_tr_pr = Some(RowProperties::default());
                        _tr_pr_child_idx = 0;
                        row_child_idx += 1;
                    }
                    name if name == EL_TC && current_row.is_some() => {
                        current_cell = Some(Cell::new());
                        cell_child_idx = 0;
                        row_child_idx += 1;
                    }
                    name if name == EL_TC_PR && current_cell.is_some() => {
                        in_tc_pr = true;
                        current_tc_pr = Some(CellProperties::default());
                        _tc_pr_child_idx = 0;
                        cell_child_idx += 1;
                    }
                    name if name == EL_TC_BORDERS && in_tc_pr => {
                        in_tc_borders = true;
                        current_tc_borders = Some(CellBorders::default());
                        _tc_pr_child_idx += 1;
                    }
                    name if name == EL_P && in_body => {
                        current_para = Some(Paragraph::new());
                        para_child_idx = 0;
                        if current_cell.is_some() {
                            cell_child_idx += 1;
                        } else {
                            body_child_idx += 1;
                        }
                    }
                    name if name == EL_HYPERLINK && current_para.is_some() => {
                        let mut hyperlink = Hyperlink::new();
                        // Extract known attributes and capture unknown ones
                        for (attr_idx, attr) in e.attributes().filter_map(|a| a.ok()).enumerate() {
                            let key = attr.key.as_ref();
                            if key == b"r:id" {
                                hyperlink.rel_id =
                                    Some(String::from_utf8_lossy(&attr.value).into_owned());
                            } else if key == b"w:anchor" || key == b"anchor" {
                                hyperlink.anchor =
                                    Some(String::from_utf8_lossy(&attr.value).into_owned());
                            } else {
                                // Unknown attribute - preserve with position
                                hyperlink.unknown_attrs.push(PositionedAttr::new(
                                    attr_idx,
                                    String::from_utf8_lossy(key),
                                    String::from_utf8_lossy(&attr.value),
                                ));
                            }
                        }
                        current_hyperlink = Some(hyperlink);
                        hyperlink_child_idx = 0;
                        para_child_idx += 1;
                    }
                    name if name == EL_R && current_para.is_some() => {
                        let mut run = Run::new();
                        // Capture unknown attributes with position
                        for (attr_idx, attr) in e.attributes().filter_map(|a| a.ok()).enumerate() {
                            // Run element has no known attributes we parse, capture all
                            run.unknown_attrs.push(PositionedAttr::new(
                                attr_idx,
                                String::from_utf8_lossy(attr.key.as_ref()),
                                String::from_utf8_lossy(&attr.value),
                            ));
                        }
                        current_run = Some(run);
                        run_child_idx = 0;
                        if current_hyperlink.is_some() {
                            hyperlink_child_idx += 1;
                        } else {
                            para_child_idx += 1;
                        }
                    }
                    name if name == EL_T && current_run.is_some() => {
                        in_text = true;
                        run_child_idx += 1;
                    }
                    name if name == EL_PPR && current_para.is_some() => {
                        in_ppr = true;
                        current_ppr = Some(ParagraphProperties::default());
                        ppr_child_idx = 0;
                        para_child_idx += 1;
                    }
                    name if name == EL_NUMPR && in_ppr => {
                        in_numpr = true;
                        current_numid = None;
                        current_ilvl = None;
                        ppr_child_idx += 1;
                    }
                    name if name == EL_P_BDR && in_ppr => {
                        in_p_bdr = true;
                        current_p_borders = Some(ParagraphBorders::default());
                        ppr_child_idx += 1;
                    }
                    name if name == EL_RPR && current_run.is_some() => {
                        in_rpr = true;
                        current_rpr = Some(RunProperties::default());
                        rpr_child_idx = 0;
                        run_child_idx += 1;
                    }
                    name if name == EL_DRAWING && current_run.is_some() => {
                        current_drawing = Some(Drawing::new());
                        run_child_idx += 1;
                    }
                    name if name == EL_INLINE && current_drawing.is_some() => {
                        // Start of an inline image - create with placeholder rel_id
                        current_image = Some(InlineImage::new(""));
                    }
                    name if name == EL_SECT_PR && in_body => {
                        // Section properties at document level (defines last section)
                        current_sect_pr = Some(SectionProperties::default());
                        in_sect_pr = true;
                        sect_pr_child_idx = 0;
                        body_child_idx += 1;
                    }
                    _ => {
                        // Only capture unknown elements when we're in a container context
                        // IMPORTANT: Don't capture while inside drawing/inline contexts -
                        // we need to continue parsing to find nested elements like blip
                        let should_capture = in_sect_pr
                            || in_rpr
                            || (in_ppr && !in_numpr)
                            || (current_run.is_some()
                                && current_drawing.is_none()
                                && current_image.is_none())
                            || current_hyperlink.is_some()
                            || (current_para.is_some()
                                && current_run.is_none()
                                && current_hyperlink.is_none())
                            || (current_cell.is_some() && current_para.is_none())
                            || (current_row.is_some() && current_cell.is_none())
                            || (current_table.is_some() && current_row.is_none())
                            || (in_body
                                && current_table.is_none()
                                && current_para.is_none()
                                && !in_sect_pr);

                        if should_capture {
                            // Capture unknown elements for round-trip preservation
                            let raw = RawXmlElement::from_reader(&mut reader, &e)?;
                            let node = RawXmlNode::Element(raw);
                            // Add to the innermost active container with position
                            if in_sect_pr {
                                if let Some(sect_pr) = current_sect_pr.as_mut() {
                                    sect_pr
                                        .unknown_children
                                        .push(PositionedNode::new(sect_pr_child_idx, node));
                                    sect_pr_child_idx += 1;
                                }
                            } else if in_rpr {
                                if let Some(rpr) = current_rpr.as_mut() {
                                    rpr.unknown_children
                                        .push(PositionedNode::new(rpr_child_idx, node));
                                    rpr_child_idx += 1;
                                }
                            } else if in_ppr && !in_numpr {
                                if let Some(ppr) = current_ppr.as_mut() {
                                    ppr.unknown_children
                                        .push(PositionedNode::new(ppr_child_idx, node));
                                    ppr_child_idx += 1;
                                }
                            } else if let Some(run) = current_run.as_mut() {
                                run.unknown_children
                                    .push(PositionedNode::new(run_child_idx, node));
                                run_child_idx += 1;
                            } else if let Some(hyperlink) = current_hyperlink.as_mut() {
                                hyperlink
                                    .unknown_children
                                    .push(PositionedNode::new(hyperlink_child_idx, node));
                                hyperlink_child_idx += 1;
                            } else if let Some(para) = current_para.as_mut() {
                                para.unknown_children
                                    .push(PositionedNode::new(para_child_idx, node));
                                para_child_idx += 1;
                            } else if let Some(cell) = current_cell.as_mut() {
                                cell.unknown_children
                                    .push(PositionedNode::new(cell_child_idx, node));
                                cell_child_idx += 1;
                            } else if let Some(row) = current_row.as_mut() {
                                row.unknown_children
                                    .push(PositionedNode::new(row_child_idx, node));
                                row_child_idx += 1;
                            } else if let Some(table) = current_table.as_mut() {
                                table
                                    .unknown_children
                                    .push(PositionedNode::new(table_child_idx, node));
                                table_child_idx += 1;
                            } else if in_body {
                                body.unknown_children
                                    .push(PositionedNode::new(body_child_idx, node));
                                body_child_idx += 1;
                            }
                        }
                        // If not in any container or inside drawing/image parsing, skip
                    }
                }
            }
            Ok(Event::Empty(e)) => {
                let name = e.name();
                let local = local_name(name.as_ref());
                match local {
                    name if name == EL_T && current_run.is_some() => {
                        // Empty text element, nothing to do
                    }
                    name if name == EL_BR && current_run.is_some() => {
                        // Check if this is a page break
                        let mut is_page_break = false;
                        for attr in e.attributes().filter_map(|a| a.ok()) {
                            if (attr.key.as_ref() == b"w:type" || attr.key.as_ref() == b"type")
                                && attr.value.as_ref() == b"page"
                            {
                                is_page_break = true;
                            }
                        }
                        if let Some(run) = current_run.as_mut() {
                            if is_page_break {
                                run.page_break = true;
                            } else {
                                // Regular line break
                                run.text.push('\n');
                            }
                        }
                    }
                    name if name == EL_TAB && current_run.is_some() => {
                        // Tab character
                        if let Some(run) = current_run.as_mut() {
                            run.text.push('\t');
                        }
                    }
                    name if name == EL_PSTYLE && in_ppr => {
                        if let Some(ppr) = current_ppr.as_mut() {
                            for attr in e.attributes().filter_map(|a| a.ok()) {
                                if attr.key.as_ref() == b"w:val" || attr.key.as_ref() == b"val" {
                                    ppr.style =
                                        Some(String::from_utf8_lossy(&attr.value).into_owned());
                                }
                            }
                        }
                    }
                    name if name == EL_JC && in_ppr => {
                        if let Some(ppr) = current_ppr.as_mut() {
                            for attr in e.attributes().filter_map(|a| a.ok()) {
                                if attr.key.as_ref() == b"w:val" || attr.key.as_ref() == b"val" {
                                    let val = std::str::from_utf8(&attr.value).unwrap_or("left");
                                    ppr.alignment = Some(Alignment::parse(val));
                                }
                            }
                        }
                    }
                    name if name == EL_SPACING && in_ppr => {
                        if let Some(ppr) = current_ppr.as_mut() {
                            for attr in e.attributes().filter_map(|a| a.ok()) {
                                let key = attr.key.as_ref();
                                if let Ok(s) = std::str::from_utf8(&attr.value) {
                                    match key {
                                        b"w:before" | b"before" => {
                                            ppr.spacing_before = s.parse().ok();
                                        }
                                        b"w:after" | b"after" => {
                                            ppr.spacing_after = s.parse().ok();
                                        }
                                        b"w:line" | b"line" => {
                                            ppr.spacing_line = s.parse().ok();
                                        }
                                        _ => {}
                                    }
                                }
                            }
                        }
                    }
                    name if name == EL_IND && in_ppr => {
                        if let Some(ppr) = current_ppr.as_mut() {
                            for attr in e.attributes().filter_map(|a| a.ok()) {
                                let key = attr.key.as_ref();
                                if let Ok(s) = std::str::from_utf8(&attr.value) {
                                    match key {
                                        b"w:left" | b"left" => {
                                            ppr.indent_left = s.parse().ok();
                                        }
                                        b"w:right" | b"right" => {
                                            ppr.indent_right = s.parse().ok();
                                        }
                                        b"w:firstLine" | b"firstLine" => {
                                            ppr.indent_first_line = s.parse().ok();
                                        }
                                        b"w:hanging" | b"hanging" => {
                                            ppr.indent_hanging = s.parse().ok();
                                        }
                                        _ => {}
                                    }
                                }
                            }
                        }
                    }
                    // Paragraph shading (w:shd in pPr)
                    name if name == EL_SHD && in_ppr && !in_p_bdr => {
                        if let Some(ppr) = current_ppr.as_mut() {
                            let mut shading = CellShading::default();
                            for attr in e.attributes().filter_map(|a| a.ok()) {
                                let key = attr.key.as_ref();
                                if let Ok(s) = std::str::from_utf8(&attr.value) {
                                    match key {
                                        b"w:fill" | b"fill" => {
                                            if s != "auto" {
                                                shading.fill = Some(s.to_string());
                                            }
                                        }
                                        b"w:color" | b"color" => {
                                            if s != "auto" {
                                                shading.color = Some(s.to_string());
                                            }
                                        }
                                        b"w:val" | b"val" => {
                                            shading.pattern = Some(ShadingPattern::parse(s));
                                        }
                                        _ => {}
                                    }
                                }
                            }
                            ppr.shading = Some(shading);
                            ppr_child_idx += 1;
                        }
                    }
                    // Outline level (w:outlineLvl)
                    name if name == EL_OUTLINE_LVL && in_ppr => {
                        if let Some(ppr) = current_ppr.as_mut() {
                            for attr in e.attributes().filter_map(|a| a.ok()) {
                                if (attr.key.as_ref() == b"w:val" || attr.key.as_ref() == b"val")
                                    && let Ok(s) = std::str::from_utf8(&attr.value)
                                {
                                    ppr.outline_level = s.parse().ok();
                                }
                            }
                            ppr_child_idx += 1;
                        }
                    }
                    // Keep with next (w:keepNext)
                    name if name == EL_KEEP_NEXT && in_ppr => {
                        if let Some(ppr) = current_ppr.as_mut() {
                            ppr.keep_next = parse_toggle_val(&e);
                            ppr_child_idx += 1;
                        }
                    }
                    // Keep lines together (w:keepLines)
                    name if name == EL_KEEP_LINES && in_ppr => {
                        if let Some(ppr) = current_ppr.as_mut() {
                            ppr.keep_lines = parse_toggle_val(&e);
                            ppr_child_idx += 1;
                        }
                    }
                    // Page break before (w:pageBreakBefore)
                    name if name == EL_PAGE_BREAK_BEFORE && in_ppr => {
                        if let Some(ppr) = current_ppr.as_mut() {
                            ppr.page_break_before = parse_toggle_val(&e);
                            ppr_child_idx += 1;
                        }
                    }
                    // Widow/orphan control (w:widowControl)
                    name if name == EL_WIDOW_CONTROL && in_ppr => {
                        if let Some(ppr) = current_ppr.as_mut() {
                            ppr.widow_control = Some(parse_toggle_val(&e));
                            ppr_child_idx += 1;
                        }
                    }
                    // Paragraph border elements (in pBdr)
                    name if name == EL_TOP && in_p_bdr => {
                        if let Some(borders) = current_p_borders.as_mut() {
                            borders.top = Some(parse_border(&e));
                        }
                    }
                    name if name == EL_BOTTOM && in_p_bdr => {
                        if let Some(borders) = current_p_borders.as_mut() {
                            borders.bottom = Some(parse_border(&e));
                        }
                    }
                    name if name == EL_LEFT && in_p_bdr => {
                        if let Some(borders) = current_p_borders.as_mut() {
                            borders.left = Some(parse_border(&e));
                        }
                    }
                    name if name == EL_RIGHT && in_p_bdr => {
                        if let Some(borders) = current_p_borders.as_mut() {
                            borders.right = Some(parse_border(&e));
                        }
                    }
                    name if name == EL_BETWEEN && in_p_bdr => {
                        if let Some(borders) = current_p_borders.as_mut() {
                            borders.between = Some(parse_border(&e));
                        }
                    }
                    name if name == EL_BAR && in_p_bdr => {
                        if let Some(borders) = current_p_borders.as_mut() {
                            borders.bar = Some(parse_border(&e));
                        }
                    }
                    name if name == EL_RSTYLE && in_rpr => {
                        if let Some(rpr) = current_rpr.as_mut() {
                            for attr in e.attributes().filter_map(|a| a.ok()) {
                                if attr.key.as_ref() == b"w:val" || attr.key.as_ref() == b"val" {
                                    rpr.style =
                                        Some(String::from_utf8_lossy(&attr.value).into_owned());
                                }
                            }
                        }
                    }
                    name if name == EL_NUMID && in_numpr => {
                        for attr in e.attributes().filter_map(|a| a.ok()) {
                            if (attr.key.as_ref() == b"w:val" || attr.key.as_ref() == b"val")
                                && let Ok(s) = std::str::from_utf8(&attr.value)
                            {
                                current_numid = s.parse().ok();
                            }
                        }
                    }
                    name if name == EL_ILVL && in_numpr => {
                        for attr in e.attributes().filter_map(|a| a.ok()) {
                            if (attr.key.as_ref() == b"w:val" || attr.key.as_ref() == b"val")
                                && let Ok(s) = std::str::from_utf8(&attr.value)
                            {
                                current_ilvl = s.parse().ok();
                            }
                        }
                    }
                    name if name == EL_B && in_rpr => {
                        if let Some(rpr) = current_rpr.as_mut() {
                            rpr.bold = parse_toggle_val(&e);
                        }
                    }
                    name if name == EL_I && in_rpr => {
                        if let Some(rpr) = current_rpr.as_mut() {
                            rpr.italic = parse_toggle_val(&e);
                        }
                    }
                    name if name == EL_U && in_rpr => {
                        if let Some(rpr) = current_rpr.as_mut() {
                            // Parse underline style from w:val attribute
                            let mut style = UnderlineStyle::Single;
                            for attr in e.attributes().filter_map(|a| a.ok()) {
                                if attr.key.as_ref() == b"w:val" || attr.key.as_ref() == b"val" {
                                    if let Ok(s) = std::str::from_utf8(&attr.value) {
                                        if let Some(parsed) = UnderlineStyle::parse(s) {
                                            style = parsed;
                                        } else {
                                            // "none" returns None, meaning no underline
                                            rpr.underline = None;
                                            break;
                                        }
                                    }
                                    rpr.underline = Some(style);
                                    break;
                                }
                            }
                            // If no val attribute, default to single underline
                            if rpr.underline.is_none() {
                                rpr.underline = Some(UnderlineStyle::Single);
                            }
                        }
                    }
                    name if name == EL_STRIKE && in_rpr => {
                        if let Some(rpr) = current_rpr.as_mut() {
                            rpr.strike = parse_toggle_val(&e);
                        }
                    }
                    name if name == EL_DSTRIKE && in_rpr => {
                        if let Some(rpr) = current_rpr.as_mut() {
                            rpr.double_strike = parse_toggle_val(&e);
                        }
                    }
                    name if name == EL_HIGHLIGHT && in_rpr => {
                        if let Some(rpr) = current_rpr.as_mut() {
                            for attr in e.attributes().filter_map(|a| a.ok()) {
                                if attr.key.as_ref() == b"w:val" || attr.key.as_ref() == b"val" {
                                    if let Ok(s) = std::str::from_utf8(&attr.value) {
                                        rpr.highlight = HighlightColor::parse(s);
                                    }
                                    break;
                                }
                            }
                        }
                    }
                    name if name == EL_VERT_ALIGN && in_rpr => {
                        if let Some(rpr) = current_rpr.as_mut() {
                            for attr in e.attributes().filter_map(|a| a.ok()) {
                                if attr.key.as_ref() == b"w:val" || attr.key.as_ref() == b"val" {
                                    if let Ok(s) = std::str::from_utf8(&attr.value) {
                                        rpr.vertical_align = VerticalAlign::parse(s);
                                    }
                                    break;
                                }
                            }
                        }
                    }
                    name if name == EL_CAPS && in_rpr => {
                        if let Some(rpr) = current_rpr.as_mut() {
                            rpr.all_caps = parse_toggle_val(&e);
                        }
                    }
                    name if name == EL_SMALL_CAPS && in_rpr => {
                        if let Some(rpr) = current_rpr.as_mut() {
                            rpr.small_caps = parse_toggle_val(&e);
                        }
                    }
                    name if name == EL_SZ && in_rpr => {
                        if let Some(rpr) = current_rpr.as_mut() {
                            for attr in e.attributes().filter_map(|a| a.ok()) {
                                if (attr.key.as_ref() == b"w:val" || attr.key.as_ref() == b"val")
                                    && let Ok(s) = std::str::from_utf8(&attr.value)
                                {
                                    rpr.size = s.parse().ok();
                                }
                            }
                        }
                    }
                    name if name == EL_RFONTS && in_rpr => {
                        if let Some(rpr) = current_rpr.as_mut() {
                            for attr in e.attributes().filter_map(|a| a.ok()) {
                                if attr.key.as_ref() == b"w:ascii" || attr.key.as_ref() == b"ascii"
                                {
                                    rpr.font =
                                        Some(String::from_utf8_lossy(&attr.value).into_owned());
                                    break;
                                }
                            }
                        }
                    }
                    name if name == EL_COLOR && in_rpr => {
                        if let Some(rpr) = current_rpr.as_mut() {
                            for attr in e.attributes().filter_map(|a| a.ok()) {
                                if attr.key.as_ref() == b"w:val" || attr.key.as_ref() == b"val" {
                                    let val = String::from_utf8_lossy(&attr.value).into_owned();
                                    // Skip "auto" which means use the default color
                                    if val != "auto" {
                                        rpr.color = Some(val);
                                    }
                                }
                            }
                        }
                    }
                    // Image extent (dimensions)
                    name if name == EL_EXTENT && current_image.is_some() => {
                        if let Some(img) = current_image.as_mut() {
                            for attr in e.attributes().filter_map(|a| a.ok()) {
                                match attr.key.as_ref() {
                                    b"cx" => {
                                        if let Ok(s) = std::str::from_utf8(&attr.value) {
                                            img.width_emu = s.parse().ok();
                                        }
                                    }
                                    b"cy" => {
                                        if let Ok(s) = std::str::from_utf8(&attr.value) {
                                            img.height_emu = s.parse().ok();
                                        }
                                    }
                                    _ => {}
                                }
                            }
                        }
                    }
                    // Image properties (description/alt text)
                    name if name == EL_DOCPR && current_image.is_some() => {
                        if let Some(img) = current_image.as_mut() {
                            for attr in e.attributes().filter_map(|a| a.ok()) {
                                if attr.key.as_ref() == b"descr" {
                                    img.description =
                                        Some(String::from_utf8_lossy(&attr.value).into_owned());
                                }
                            }
                        }
                    }
                    // Blip - contains the relationship ID to the image
                    name if name == EL_BLIP && current_image.is_some() => {
                        if let Some(img) = current_image.as_mut() {
                            for attr in e.attributes().filter_map(|a| a.ok()) {
                                // r:embed attribute contains the relationship ID
                                if attr.key.as_ref() == b"r:embed" || attr.key.as_ref() == b"embed"
                                {
                                    img.rel_id = String::from_utf8_lossy(&attr.value).into_owned();
                                }
                            }
                        }
                    }
                    // Page size (w:pgSz)
                    name if name == EL_PG_SZ && in_sect_pr => {
                        if let Some(sect_pr) = current_sect_pr.as_mut() {
                            let mut width = 12240u32; // Default US Letter width
                            let mut height = 15840u32; // Default US Letter height
                            let mut orient = PageOrientation::Portrait;
                            for attr in e.attributes().filter_map(|a| a.ok()) {
                                let key = attr.key.as_ref();
                                if let Ok(s) = std::str::from_utf8(&attr.value) {
                                    match key {
                                        b"w:w" | b"w" => {
                                            width = s.parse().unwrap_or(12240);
                                        }
                                        b"w:h" | b"h" => {
                                            height = s.parse().unwrap_or(15840);
                                        }
                                        b"w:orient" | b"orient" => {
                                            orient = PageOrientation::parse(s);
                                        }
                                        _ => {}
                                    }
                                }
                            }
                            sect_pr.page_size = Some(PageSize {
                                width,
                                height,
                                orientation: orient,
                            });
                            sect_pr_child_idx += 1;
                        }
                    }
                    // Page margins (w:pgMar)
                    name if name == EL_PG_MAR && in_sect_pr => {
                        if let Some(sect_pr) = current_sect_pr.as_mut() {
                            let mut margins = PageMargins::default();
                            for attr in e.attributes().filter_map(|a| a.ok()) {
                                let key = attr.key.as_ref();
                                if let Ok(s) = std::str::from_utf8(&attr.value) {
                                    match key {
                                        b"w:top" | b"top" => {
                                            margins.top = s.parse().unwrap_or(1440);
                                        }
                                        b"w:bottom" | b"bottom" => {
                                            margins.bottom = s.parse().unwrap_or(1440);
                                        }
                                        b"w:left" | b"left" => {
                                            margins.left = s.parse().unwrap_or(1440);
                                        }
                                        b"w:right" | b"right" => {
                                            margins.right = s.parse().unwrap_or(1440);
                                        }
                                        b"w:header" | b"header" => {
                                            margins.header = s.parse().ok();
                                        }
                                        b"w:footer" | b"footer" => {
                                            margins.footer = s.parse().ok();
                                        }
                                        b"w:gutter" | b"gutter" => {
                                            margins.gutter = s.parse().ok();
                                        }
                                        _ => {}
                                    }
                                }
                            }
                            sect_pr.margins = Some(margins);
                            sect_pr_child_idx += 1;
                        }
                    }
                    // Table width (w:tblW)
                    name if name == EL_TBL_W && in_tbl_pr => {
                        if let Some(tbl_pr) = current_tbl_pr.as_mut() {
                            let mut width: i32 = 0;
                            let mut width_type = WidthType::Dxa;
                            for attr in e.attributes().filter_map(|a| a.ok()) {
                                let key = attr.key.as_ref();
                                if let Ok(s) = std::str::from_utf8(&attr.value) {
                                    match key {
                                        b"w:w" | b"w" => {
                                            width = s.parse().unwrap_or(0);
                                        }
                                        b"w:type" | b"type" => {
                                            width_type = WidthType::parse(s);
                                        }
                                        _ => {}
                                    }
                                }
                            }
                            tbl_pr.width = Some(TableWidth { width, width_type });
                            _tbl_pr_child_idx += 1;
                        }
                    }
                    // Table justification (w:jc)
                    name if name == EL_JC && in_tbl_pr => {
                        if let Some(tbl_pr) = current_tbl_pr.as_mut() {
                            for attr in e.attributes().filter_map(|a| a.ok()) {
                                if (attr.key.as_ref() == b"w:val" || attr.key.as_ref() == b"val")
                                    && let Ok(s) = std::str::from_utf8(&attr.value)
                                {
                                    tbl_pr.justification = Some(TableJustification::parse(s));
                                }
                            }
                            _tbl_pr_child_idx += 1;
                        }
                    }
                    // Table indent (w:tblInd)
                    name if name == EL_TBL_IND && in_tbl_pr => {
                        if let Some(tbl_pr) = current_tbl_pr.as_mut() {
                            for attr in e.attributes().filter_map(|a| a.ok()) {
                                if (attr.key.as_ref() == b"w:w" || attr.key.as_ref() == b"w")
                                    && let Ok(s) = std::str::from_utf8(&attr.value)
                                {
                                    tbl_pr.indent = s.parse().ok();
                                }
                            }
                            _tbl_pr_child_idx += 1;
                        }
                    }
                    // Table layout (w:tblLayout)
                    name if name == EL_TBL_LAYOUT && in_tbl_pr => {
                        if let Some(tbl_pr) = current_tbl_pr.as_mut() {
                            for attr in e.attributes().filter_map(|a| a.ok()) {
                                if (attr.key.as_ref() == b"w:type" || attr.key.as_ref() == b"type")
                                    && let Ok(s) = std::str::from_utf8(&attr.value)
                                {
                                    tbl_pr.layout = Some(TableLayout::parse(s));
                                }
                            }
                            _tbl_pr_child_idx += 1;
                        }
                    }
                    // Table shading (w:shd in tblPr)
                    name if name == EL_SHD && in_tbl_pr && !in_tbl_borders => {
                        if let Some(tbl_pr) = current_tbl_pr.as_mut() {
                            let mut shading = CellShading::default();
                            for attr in e.attributes().filter_map(|a| a.ok()) {
                                let key = attr.key.as_ref();
                                if let Ok(s) = std::str::from_utf8(&attr.value) {
                                    match key {
                                        b"w:fill" | b"fill" => {
                                            if s != "auto" {
                                                shading.fill = Some(s.to_string());
                                            }
                                        }
                                        b"w:color" | b"color" => {
                                            if s != "auto" {
                                                shading.color = Some(s.to_string());
                                            }
                                        }
                                        b"w:val" | b"val" => {
                                            shading.pattern = Some(ShadingPattern::parse(s));
                                        }
                                        _ => {}
                                    }
                                }
                            }
                            tbl_pr.shading = Some(shading);
                            _tbl_pr_child_idx += 1;
                        }
                    }
                    // Grid column (w:gridCol)
                    name if name == EL_GRID_COL && in_tbl_grid => {
                        if let Some(table) = current_table.as_mut() {
                            let mut width = 0u32;
                            for attr in e.attributes().filter_map(|a| a.ok()) {
                                if (attr.key.as_ref() == b"w:w" || attr.key.as_ref() == b"w")
                                    && let Ok(s) = std::str::from_utf8(&attr.value)
                                {
                                    width = s.parse().unwrap_or(0);
                                }
                            }
                            table.grid_columns.push(GridColumn { width });
                        }
                    }
                    // Row height (w:trHeight)
                    name if name == EL_TR_HEIGHT && in_tr_pr => {
                        if let Some(tr_pr) = current_tr_pr.as_mut() {
                            let mut value = 0u32;
                            let mut rule = HeightRule::Auto;
                            for attr in e.attributes().filter_map(|a| a.ok()) {
                                let key = attr.key.as_ref();
                                if let Ok(s) = std::str::from_utf8(&attr.value) {
                                    match key {
                                        b"w:val" | b"val" => {
                                            value = s.parse().unwrap_or(0);
                                        }
                                        b"w:hRule" | b"hRule" => {
                                            rule = HeightRule::parse(s);
                                        }
                                        _ => {}
                                    }
                                }
                            }
                            tr_pr.height = Some(RowHeight { value, rule });
                            _tr_pr_child_idx += 1;
                        }
                    }
                    // Header row (w:tblHeader)
                    name if name == EL_TBL_HEADER && in_tr_pr => {
                        if let Some(tr_pr) = current_tr_pr.as_mut() {
                            tr_pr.is_header = true;
                            _tr_pr_child_idx += 1;
                        }
                    }
                    // Can't split row (w:cantSplit)
                    name if name == EL_CANT_SPLIT && in_tr_pr => {
                        if let Some(tr_pr) = current_tr_pr.as_mut() {
                            tr_pr.cant_split = true;
                            _tr_pr_child_idx += 1;
                        }
                    }
                    // Cell width (w:tcW)
                    name if name == EL_TC_W && in_tc_pr => {
                        if let Some(tc_pr) = current_tc_pr.as_mut() {
                            let mut width = 0u32;
                            let mut width_type = WidthType::Dxa;
                            for attr in e.attributes().filter_map(|a| a.ok()) {
                                let key = attr.key.as_ref();
                                if let Ok(s) = std::str::from_utf8(&attr.value) {
                                    match key {
                                        b"w:w" | b"w" => {
                                            width = s.parse().unwrap_or(0);
                                        }
                                        b"w:type" | b"type" => {
                                            width_type = WidthType::parse(s);
                                        }
                                        _ => {}
                                    }
                                }
                            }
                            tc_pr.width = Some(CellWidth { width, width_type });
                            _tc_pr_child_idx += 1;
                        }
                    }
                    // Grid span (w:gridSpan)
                    name if name == EL_GRID_SPAN && in_tc_pr => {
                        if let Some(tc_pr) = current_tc_pr.as_mut() {
                            for attr in e.attributes().filter_map(|a| a.ok()) {
                                if (attr.key.as_ref() == b"w:val" || attr.key.as_ref() == b"val")
                                    && let Ok(s) = std::str::from_utf8(&attr.value)
                                {
                                    tc_pr.grid_span = s.parse().ok();
                                }
                            }
                            _tc_pr_child_idx += 1;
                        }
                    }
                    // Vertical merge (w:vMerge)
                    name if name == EL_V_MERGE && in_tc_pr => {
                        if let Some(tc_pr) = current_tc_pr.as_mut() {
                            // If vMerge has no val attribute, it means "continue"
                            let mut found_val = false;
                            for attr in e.attributes().filter_map(|a| a.ok()) {
                                if (attr.key.as_ref() == b"w:val" || attr.key.as_ref() == b"val")
                                    && let Ok(s) = std::str::from_utf8(&attr.value)
                                {
                                    tc_pr.vertical_merge = VerticalMerge::parse(s);
                                    found_val = true;
                                }
                            }
                            if !found_val {
                                // Empty vMerge means continue
                                tc_pr.vertical_merge = Some(VerticalMerge::Continue);
                            }
                            _tc_pr_child_idx += 1;
                        }
                    }
                    // Vertical alignment (w:vAlign)
                    name if name == EL_V_ALIGN && in_tc_pr => {
                        if let Some(tc_pr) = current_tc_pr.as_mut() {
                            for attr in e.attributes().filter_map(|a| a.ok()) {
                                if (attr.key.as_ref() == b"w:val" || attr.key.as_ref() == b"val")
                                    && let Ok(s) = std::str::from_utf8(&attr.value)
                                {
                                    tc_pr.vertical_align = Some(CellVerticalAlign::parse(s));
                                }
                            }
                            _tc_pr_child_idx += 1;
                        }
                    }
                    // Shading (w:shd) - can appear in tcPr or rPr/pPr
                    name if name == EL_SHD && in_tc_pr => {
                        if let Some(tc_pr) = current_tc_pr.as_mut() {
                            let mut shading = CellShading {
                                fill: None,
                                color: None,
                                pattern: None,
                            };
                            for attr in e.attributes().filter_map(|a| a.ok()) {
                                let key = attr.key.as_ref();
                                if let Ok(s) = std::str::from_utf8(&attr.value) {
                                    match key {
                                        b"w:fill" | b"fill" => {
                                            if s != "auto" {
                                                shading.fill = Some(s.to_string());
                                            }
                                        }
                                        b"w:color" | b"color" => {
                                            if s != "auto" {
                                                shading.color = Some(s.to_string());
                                            }
                                        }
                                        b"w:val" | b"val" => {
                                            shading.pattern = Some(ShadingPattern::parse(s));
                                        }
                                        _ => {}
                                    }
                                }
                            }
                            tc_pr.shading = Some(shading);
                            _tc_pr_child_idx += 1;
                        }
                    }
                    // Border elements inside tblBorders
                    name if name == EL_TOP && in_tbl_borders => {
                        if let Some(borders) = current_tbl_borders.as_mut() {
                            borders.top = Some(parse_border(&e));
                        }
                    }
                    name if name == EL_BOTTOM && in_tbl_borders => {
                        if let Some(borders) = current_tbl_borders.as_mut() {
                            borders.bottom = Some(parse_border(&e));
                        }
                    }
                    name if name == EL_LEFT && in_tbl_borders => {
                        if let Some(borders) = current_tbl_borders.as_mut() {
                            borders.left = Some(parse_border(&e));
                        }
                    }
                    name if name == EL_RIGHT && in_tbl_borders => {
                        if let Some(borders) = current_tbl_borders.as_mut() {
                            borders.right = Some(parse_border(&e));
                        }
                    }
                    name if name == EL_INSIDE_H && in_tbl_borders => {
                        if let Some(borders) = current_tbl_borders.as_mut() {
                            borders.inside_h = Some(parse_border(&e));
                        }
                    }
                    name if name == EL_INSIDE_V && in_tbl_borders => {
                        if let Some(borders) = current_tbl_borders.as_mut() {
                            borders.inside_v = Some(parse_border(&e));
                        }
                    }
                    // Border elements inside tcBorders
                    name if name == EL_TOP && in_tc_borders => {
                        if let Some(borders) = current_tc_borders.as_mut() {
                            borders.top = Some(parse_border(&e));
                        }
                    }
                    name if name == EL_BOTTOM && in_tc_borders => {
                        if let Some(borders) = current_tc_borders.as_mut() {
                            borders.bottom = Some(parse_border(&e));
                        }
                    }
                    name if name == EL_LEFT && in_tc_borders => {
                        if let Some(borders) = current_tc_borders.as_mut() {
                            borders.left = Some(parse_border(&e));
                        }
                    }
                    name if name == EL_RIGHT && in_tc_borders => {
                        if let Some(borders) = current_tc_borders.as_mut() {
                            borders.right = Some(parse_border(&e));
                        }
                    }
                    name if name == EL_INSIDE_H && in_tc_borders => {
                        if let Some(borders) = current_tc_borders.as_mut() {
                            borders.inside_h = Some(parse_border(&e));
                        }
                    }
                    name if name == EL_INSIDE_V && in_tc_borders => {
                        if let Some(borders) = current_tc_borders.as_mut() {
                            borders.inside_v = Some(parse_border(&e));
                        }
                    }
                    _ => {
                        // Only capture unknown self-closing elements when in a container context
                        let should_capture = in_tbl_pr
                            || in_tr_pr
                            || in_tc_pr
                            || in_sect_pr
                            || in_rpr
                            || (in_ppr && !in_numpr)
                            || current_image.is_some()
                            || current_drawing.is_some()
                            || current_run.is_some()
                            || current_hyperlink.is_some()
                            || current_para.is_some()
                            || current_cell.is_some()
                            || current_row.is_some()
                            || current_table.is_some()
                            || in_body;

                        if should_capture {
                            // Capture unknown self-closing elements for round-trip preservation
                            let raw = RawXmlElement::from_empty(&e);
                            let node = RawXmlNode::Element(raw);
                            // Add to the innermost active container with position
                            if in_sect_pr {
                                if let Some(sect_pr) = current_sect_pr.as_mut() {
                                    sect_pr
                                        .unknown_children
                                        .push(PositionedNode::new(sect_pr_child_idx, node));
                                    sect_pr_child_idx += 1;
                                }
                            } else if in_rpr {
                                if let Some(rpr) = current_rpr.as_mut() {
                                    rpr.unknown_children
                                        .push(PositionedNode::new(rpr_child_idx, node));
                                    rpr_child_idx += 1;
                                }
                            } else if in_ppr && !in_numpr {
                                if let Some(ppr) = current_ppr.as_mut() {
                                    ppr.unknown_children
                                        .push(PositionedNode::new(ppr_child_idx, node));
                                    ppr_child_idx += 1;
                                }
                            } else if let Some(img) = current_image.as_mut() {
                                // Images don't track child positions (DrawingML is complex)
                                img.unknown_children.push(PositionedNode::new(0, node));
                            } else if let Some(drawing) = current_drawing.as_mut() {
                                // Drawings don't track child positions (DrawingML is complex)
                                drawing.unknown_children.push(PositionedNode::new(0, node));
                            } else if let Some(run) = current_run.as_mut() {
                                run.unknown_children
                                    .push(PositionedNode::new(run_child_idx, node));
                                run_child_idx += 1;
                            } else if let Some(hyperlink) = current_hyperlink.as_mut() {
                                hyperlink
                                    .unknown_children
                                    .push(PositionedNode::new(hyperlink_child_idx, node));
                                hyperlink_child_idx += 1;
                            } else if let Some(para) = current_para.as_mut() {
                                para.unknown_children
                                    .push(PositionedNode::new(para_child_idx, node));
                                para_child_idx += 1;
                            } else if let Some(cell) = current_cell.as_mut() {
                                cell.unknown_children
                                    .push(PositionedNode::new(cell_child_idx, node));
                                cell_child_idx += 1;
                            } else if let Some(row) = current_row.as_mut() {
                                row.unknown_children
                                    .push(PositionedNode::new(row_child_idx, node));
                                row_child_idx += 1;
                            } else if let Some(table) = current_table.as_mut() {
                                table
                                    .unknown_children
                                    .push(PositionedNode::new(table_child_idx, node));
                                table_child_idx += 1;
                            } else if in_body {
                                body.unknown_children
                                    .push(PositionedNode::new(body_child_idx, node));
                                body_child_idx += 1;
                            }
                        }
                        // If not in any container, silently skip (pre-body content)
                    }
                }
            }
            Ok(Event::Text(e)) => {
                if in_text && let Some(run) = current_run.as_mut() {
                    let text = e.decode().unwrap_or_default();
                    run.text.push_str(&text);
                }
            }
            Ok(Event::End(e)) => {
                let name = e.name();
                let local = local_name(name.as_ref());
                match local {
                    name if name == EL_BODY => {
                        in_body = false;
                    }
                    name if name == EL_TBL => {
                        if let Some(table) = current_table.take() {
                            body.content.push(BlockContent::Table(table));
                        }
                    }
                    name if name == EL_TBL_PR => {
                        if let Some(tbl_pr) = current_tbl_pr.take()
                            && let Some(table) = current_table.as_mut()
                        {
                            table.properties = Some(tbl_pr);
                        }
                        in_tbl_pr = false;
                    }
                    name if name == EL_TBL_BORDERS => {
                        if let Some(borders) = current_tbl_borders.take()
                            && let Some(tbl_pr) = current_tbl_pr.as_mut()
                        {
                            tbl_pr.borders = Some(borders);
                        }
                        in_tbl_borders = false;
                    }
                    name if name == EL_TBL_GRID => {
                        in_tbl_grid = false;
                    }
                    name if name == EL_TR => {
                        if let Some(mut row) = current_row.take()
                            && let Some(table) = current_table.as_mut()
                        {
                            row.properties = current_tr_pr.take();
                            table.rows.push(row);
                        }
                    }
                    name if name == EL_TR_PR => {
                        in_tr_pr = false;
                    }
                    name if name == EL_TC => {
                        if let Some(mut cell) = current_cell.take()
                            && let Some(row) = current_row.as_mut()
                        {
                            cell.properties = current_tc_pr.take();
                            row.cells.push(cell);
                        }
                    }
                    name if name == EL_TC_PR => {
                        in_tc_pr = false;
                    }
                    name if name == EL_TC_BORDERS => {
                        if let Some(borders) = current_tc_borders.take()
                            && let Some(tc_pr) = current_tc_pr.as_mut()
                        {
                            tc_pr.borders = Some(borders);
                        }
                        in_tc_borders = false;
                    }
                    name if name == EL_P && current_para.is_some() => {
                        if let Some(mut para) = current_para.take() {
                            para.properties = current_ppr.take();
                            // Add to cell if inside table, otherwise to body
                            if let Some(cell) = current_cell.as_mut() {
                                cell.paragraphs.push(para);
                            } else {
                                body.content.push(BlockContent::Paragraph(para));
                            }
                        }
                    }
                    name if name == EL_R && current_run.is_some() => {
                        if let Some(mut run) = current_run.take() {
                            run.properties = current_rpr.take();
                            // Add run to hyperlink if inside one, otherwise to paragraph
                            if let Some(hyperlink) = current_hyperlink.as_mut() {
                                hyperlink.runs.push(run);
                            } else if let Some(para) = current_para.as_mut() {
                                para.content.push(ParagraphContent::Run(run));
                            }
                        }
                    }
                    name if name == EL_HYPERLINK => {
                        if let Some(hyperlink) = current_hyperlink.take()
                            && let Some(para) = current_para.as_mut()
                        {
                            para.content.push(ParagraphContent::Hyperlink(hyperlink));
                        }
                    }
                    name if name == EL_T => {
                        in_text = false;
                    }
                    name if name == EL_NUMPR => {
                        in_numpr = false;
                        // Create numbering properties if we have both numId and ilvl
                        if let (Some(num_id), Some(ppr)) =
                            (current_numid.take(), current_ppr.as_mut())
                        {
                            ppr.numbering = Some(NumberingProperties {
                                num_id,
                                ilvl: current_ilvl.take().unwrap_or(0),
                            });
                        }
                    }
                    name if name == EL_P_BDR => {
                        if let Some(borders) = current_p_borders.take()
                            && let Some(ppr) = current_ppr.as_mut()
                        {
                            ppr.borders = Some(borders);
                        }
                        in_p_bdr = false;
                    }
                    name if name == EL_PPR => {
                        in_ppr = false;
                    }
                    name if name == EL_RPR => {
                        in_rpr = false;
                    }
                    // End of section properties - add to body
                    name if name == EL_SECT_PR => {
                        if let Some(sect_pr) = current_sect_pr.take() {
                            body.section_properties = Some(sect_pr);
                        }
                        in_sect_pr = false;
                    }
                    // End of inline image - add to current drawing
                    name if name == EL_INLINE => {
                        if let Some(img) = current_image.take()
                            && !img.rel_id.is_empty()
                            && let Some(drawing) = current_drawing.as_mut()
                        {
                            drawing.images.push(img);
                        }
                    }
                    // End of drawing - add to current run
                    name if name == EL_DRAWING => {
                        if let Some(drawing) = current_drawing.take()
                            && let Some(run) = current_run.as_mut()
                        {
                            run.drawings.push(drawing);
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

    Ok(body)
}

/// Extract the local name from a potentially namespaced element name.
fn local_name(name: &[u8]) -> &[u8] {
    // Handle both "w:p" and "p" formats
    if let Some(pos) = name.iter().position(|&b| b == b':') {
        &name[pos + 1..]
    } else {
        name
    }
}

/// Parse a toggle property value (like <w:b/> or <w:b w:val="true"/>).
///
/// Toggle properties are true if:
/// - Element is present with no val attribute
/// - Element has val="true", "1", or "on"
fn parse_toggle_val(e: &quick_xml::events::BytesStart) -> bool {
    for attr in e.attributes().filter_map(|a| a.ok()) {
        if attr.key.as_ref() == b"w:val" || attr.key.as_ref() == b"val" {
            return matches!(
                attr.value.as_ref(),
                b"true" | b"1" | b"on" | b"True" | b"On"
            );
        }
    }
    // No val attribute means true
    true
}

/// Parse a border element's attributes.
fn parse_border(e: &quick_xml::events::BytesStart) -> Border {
    let mut border = Border {
        style: BorderStyle::Single,
        size: None,
        color: None,
        space: None,
    };
    for attr in e.attributes().filter_map(|a| a.ok()) {
        let key = attr.key.as_ref();
        if let Ok(s) = std::str::from_utf8(&attr.value) {
            match key {
                b"w:val" | b"val" => {
                    border.style = BorderStyle::parse(s);
                }
                b"w:sz" | b"sz" => {
                    border.size = s.parse().ok();
                }
                b"w:color" | b"color" => {
                    if s != "auto" {
                        border.color = Some(s.to_string());
                    }
                }
                b"w:space" | b"space" => {
                    border.space = s.parse().ok();
                }
                _ => {}
            }
        }
    }
    border
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_parse_simple_document() {
        let xml = br#"<?xml version="1.0" encoding="UTF-8"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p>
      <w:r>
        <w:t>Hello, World!</w:t>
      </w:r>
    </w:p>
  </w:body>
</w:document>"#;

        let body = parse_document(xml).unwrap();

        assert_eq!(body.paragraphs().len(), 1);
        assert_eq!(body.paragraphs()[0].runs().len(), 1);
        assert_eq!(body.paragraphs()[0].runs()[0].text(), "Hello, World!");
        assert_eq!(body.text(), "Hello, World!");
    }

    #[test]
    fn test_parse_multiple_paragraphs() {
        let xml = br#"<?xml version="1.0" encoding="UTF-8"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p>
      <w:r><w:t>First paragraph.</w:t></w:r>
    </w:p>
    <w:p>
      <w:r><w:t>Second paragraph.</w:t></w:r>
    </w:p>
  </w:body>
</w:document>"#;

        let body = parse_document(xml).unwrap();

        assert_eq!(body.paragraphs().len(), 2);
        assert_eq!(body.text(), "First paragraph.\nSecond paragraph.");
    }

    #[test]
    fn test_parse_multiple_runs() {
        let xml = br#"<?xml version="1.0" encoding="UTF-8"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p>
      <w:r><w:t>Hello, </w:t></w:r>
      <w:r><w:t>World!</w:t></w:r>
    </w:p>
  </w:body>
</w:document>"#;

        let body = parse_document(xml).unwrap();

        assert_eq!(body.paragraphs().len(), 1);
        assert_eq!(body.paragraphs()[0].runs().len(), 2);
        assert_eq!(body.paragraphs()[0].text(), "Hello, World!");
    }

    #[test]
    fn test_parse_formatting() {
        let xml = br#"<?xml version="1.0" encoding="UTF-8"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p>
      <w:r>
        <w:rPr>
          <w:b/>
          <w:i/>
          <w:u w:val="single"/>
        </w:rPr>
        <w:t>Formatted text</w:t>
      </w:r>
    </w:p>
  </w:body>
</w:document>"#;

        let body = parse_document(xml).unwrap();
        let run = &body.paragraphs()[0].runs()[0];

        assert!(run.is_bold());
        assert!(run.is_italic());
        assert!(run.is_underline());
    }

    #[test]
    fn test_parse_toggle_values() {
        let xml = br#"<?xml version="1.0" encoding="UTF-8"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p>
      <w:r>
        <w:rPr>
          <w:b w:val="true"/>
          <w:i w:val="false"/>
        </w:rPr>
        <w:t>Test</w:t>
      </w:r>
    </w:p>
  </w:body>
</w:document>"#;

        let body = parse_document(xml).unwrap();
        let run = &body.paragraphs()[0].runs()[0];

        assert!(run.is_bold());
        assert!(!run.is_italic());
    }

    #[test]
    fn test_parse_paragraph_style() {
        let xml = br#"<?xml version="1.0" encoding="UTF-8"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p>
      <w:pPr>
        <w:pStyle w:val="Heading1"/>
      </w:pPr>
      <w:r><w:t>Title</w:t></w:r>
    </w:p>
  </w:body>
</w:document>"#;

        let body = parse_document(xml).unwrap();
        let para = &body.paragraphs()[0];

        assert_eq!(
            para.properties().and_then(|p| p.style.as_deref()),
            Some("Heading1")
        );
    }

    #[test]
    fn test_parse_line_breaks_and_tabs() {
        let xml = br#"<?xml version="1.0" encoding="UTF-8"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p>
      <w:r>
        <w:t>Line 1</w:t>
        <w:br/>
        <w:t>Line 2</w:t>
        <w:tab/>
        <w:t>After tab</w:t>
      </w:r>
    </w:p>
  </w:body>
</w:document>"#;

        let body = parse_document(xml).unwrap();
        let text = body.paragraphs()[0].runs()[0].text();

        assert_eq!(text, "Line 1\nLine 2\tAfter tab");
    }

    #[test]
    fn test_parse_simple_table() {
        let xml = br#"<?xml version="1.0" encoding="UTF-8"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:tbl>
      <w:tr>
        <w:tc>
          <w:p><w:r><w:t>A1</w:t></w:r></w:p>
        </w:tc>
        <w:tc>
          <w:p><w:r><w:t>B1</w:t></w:r></w:p>
        </w:tc>
      </w:tr>
      <w:tr>
        <w:tc>
          <w:p><w:r><w:t>A2</w:t></w:r></w:p>
        </w:tc>
        <w:tc>
          <w:p><w:r><w:t>B2</w:t></w:r></w:p>
        </w:tc>
      </w:tr>
    </w:tbl>
  </w:body>
</w:document>"#;

        let body = parse_document(xml).unwrap();

        // Should have one table
        let tables: Vec<_> = body.tables().collect();
        assert_eq!(tables.len(), 1);

        let table = tables[0];
        assert_eq!(table.row_count(), 2);
        assert_eq!(table.column_count(), 2);

        // Check cell content
        assert_eq!(table.rows()[0].cells()[0].text(), "A1");
        assert_eq!(table.rows()[0].cells()[1].text(), "B1");
        assert_eq!(table.rows()[1].cells()[0].text(), "A2");
        assert_eq!(table.rows()[1].cells()[1].text(), "B2");
    }

    #[test]
    fn test_parse_mixed_content() {
        let xml = br#"<?xml version="1.0" encoding="UTF-8"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p><w:r><w:t>Before table</w:t></w:r></w:p>
    <w:tbl>
      <w:tr>
        <w:tc><w:p><w:r><w:t>Cell</w:t></w:r></w:p></w:tc>
      </w:tr>
    </w:tbl>
    <w:p><w:r><w:t>After table</w:t></w:r></w:p>
  </w:body>
</w:document>"#;

        let body = parse_document(xml).unwrap();

        // Should have 3 block elements
        assert_eq!(body.content().len(), 3);

        // Check order
        assert!(matches!(body.content()[0], BlockContent::Paragraph(_)));
        assert!(matches!(body.content()[1], BlockContent::Table(_)));
        assert!(matches!(body.content()[2], BlockContent::Paragraph(_)));

        // All paragraphs (including table cells)
        assert_eq!(body.paragraphs().len(), 3);
    }

    #[test]
    fn test_table_text_extraction() {
        let xml = br#"<?xml version="1.0" encoding="UTF-8"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:tbl>
      <w:tr>
        <w:tc><w:p><w:r><w:t>A</w:t></w:r></w:p></w:tc>
        <w:tc><w:p><w:r><w:t>B</w:t></w:r></w:p></w:tc>
      </w:tr>
      <w:tr>
        <w:tc><w:p><w:r><w:t>C</w:t></w:r></w:p></w:tc>
        <w:tc><w:p><w:r><w:t>D</w:t></w:r></w:p></w:tc>
      </w:tr>
    </w:tbl>
  </w:body>
</w:document>"#;

        let body = parse_document(xml).unwrap();
        let tables: Vec<_> = body.tables().collect();
        let table = tables[0];

        // Row text uses tabs between cells
        assert_eq!(table.rows()[0].text(), "A\tB");
        assert_eq!(table.rows()[1].text(), "C\tD");

        // Table text uses newlines between rows
        assert_eq!(table.text(), "A\tB\nC\tD");
    }

    #[test]
    fn test_parse_inline_image() {
        // Simplified DrawingML structure for an inline image
        let xml = br#"<?xml version="1.0" encoding="UTF-8"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
            xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
            xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <w:body>
    <w:p>
      <w:r>
        <w:drawing>
          <wp:inline distT="0" distB="0" distL="0" distR="0">
            <wp:extent cx="914400" cy="457200"/>
            <wp:docPr id="1" name="Picture 1" descr="Test image"/>
            <a:graphic>
              <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">
                <a:blip r:embed="rId4"/>
              </a:graphicData>
            </a:graphic>
          </wp:inline>
        </w:drawing>
      </w:r>
    </w:p>
  </w:body>
</w:document>"#;

        let body = parse_document(xml).unwrap();
        let para = &body.paragraphs()[0];
        let run = &para.runs()[0];

        // Should have one drawing with one image
        assert_eq!(run.drawings().len(), 1);
        let drawing = &run.drawings()[0];
        assert_eq!(drawing.images().len(), 1);

        let image = &drawing.images()[0];
        assert_eq!(image.rel_id(), "rId4");
        assert_eq!(image.width_emu(), Some(914400)); // 1 inch
        assert_eq!(image.height_emu(), Some(457200)); // 0.5 inch
        assert_eq!(image.description(), Some("Test image"));

        // Check convenience methods
        assert!((image.width_inches().unwrap() - 1.0).abs() < 0.001);
        assert!((image.height_inches().unwrap() - 0.5).abs() < 0.001);

        // Run should report having images
        assert!(run.has_images());
    }

    #[test]
    fn test_resolve_path() {
        // Relative path resolution
        assert_eq!(
            resolve_path("word/document.xml", "media/image1.png"),
            "word/media/image1.png"
        );
        assert_eq!(
            resolve_path("word/document.xml", "../media/image1.png"),
            "word/../media/image1.png"
        );

        // Absolute path
        assert_eq!(
            resolve_path("word/document.xml", "/word/media/image1.png"),
            "word/media/image1.png"
        );
    }

    #[test]
    fn test_content_type_from_path() {
        assert_eq!(content_type_from_path("word/media/image1.png"), "image/png");
        assert_eq!(
            content_type_from_path("word/media/image2.jpg"),
            "image/jpeg"
        );
        assert_eq!(
            content_type_from_path("word/media/image3.JPEG"),
            "image/jpeg"
        );
        assert_eq!(content_type_from_path("word/media/image4.gif"), "image/gif");
        assert_eq!(
            content_type_from_path("word/media/unknown.xyz"),
            "application/octet-stream"
        );
    }

    #[test]
    fn test_parse_hyperlink() {
        let xml = br#"<?xml version="1.0" encoding="UTF-8"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <w:body>
    <w:p>
      <w:r><w:t>Click </w:t></w:r>
      <w:hyperlink r:id="rId5">
        <w:r><w:t>here</w:t></w:r>
      </w:hyperlink>
      <w:r><w:t> for more info.</w:t></w:r>
    </w:p>
  </w:body>
</w:document>"#;

        let body = parse_document(xml).unwrap();
        let para = &body.paragraphs()[0];

        // Should have 3 content items: run, hyperlink, run
        assert_eq!(para.content().len(), 3);

        // Check the hyperlink
        let hyperlinks: Vec<_> = para.hyperlinks().collect();
        assert_eq!(hyperlinks.len(), 1);
        let link = hyperlinks[0];
        assert_eq!(link.rel_id(), Some("rId5"));
        assert_eq!(link.text(), "here");
        assert!(link.is_external());

        // Full paragraph text should include hyperlink text
        assert_eq!(para.text(), "Click here for more info.");
    }

    #[test]
    fn test_parse_internal_hyperlink() {
        let xml = br#"<?xml version="1.0" encoding="UTF-8"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p>
      <w:hyperlink w:anchor="section1">
        <w:r><w:t>Jump to section 1</w:t></w:r>
      </w:hyperlink>
    </w:p>
  </w:body>
</w:document>"#;

        let body = parse_document(xml).unwrap();
        let para = &body.paragraphs()[0];
        let links: Vec<_> = para.hyperlinks().collect();

        assert_eq!(links.len(), 1);
        let link = links[0];
        assert_eq!(link.anchor(), Some("section1"));
        assert!(!link.is_external());
        assert_eq!(link.text(), "Jump to section 1");
    }

    #[test]
    fn test_unknown_elements_preserved() {
        // Document with unknown elements that should be preserved
        let xml = br#"<?xml version="1.0" encoding="UTF-8"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p>
      <w:pPr>
        <w:pStyle w:val="Normal"/>
        <w:customElement w:custom="value"/>
      </w:pPr>
      <w:r>
        <w:rPr>
          <w:b/>
          <w:unknownProp w:foo="bar"/>
        </w:rPr>
        <w:t>Hello</w:t>
      </w:r>
    </w:p>
  </w:body>
</w:document>"#;

        let body = parse_document(xml).unwrap();

        // Check that known elements are parsed
        let para = &body.paragraphs()[0];
        assert_eq!(para.text(), "Hello");
        assert!(para.runs()[0].is_bold());

        // Check that unknown elements in pPr are preserved
        let ppr = para.properties().unwrap();
        assert_eq!(ppr.unknown_children.len(), 1);
        if let crate::raw_xml::RawXmlNode::Element(elem) = &ppr.unknown_children[0].node {
            assert_eq!(elem.name, "w:customElement");
            assert_eq!(
                elem.attributes,
                vec![("w:custom".to_string(), "value".to_string())]
            );
        } else {
            panic!("Expected element node");
        }

        // Check that unknown elements in rPr are preserved
        let rpr = para.runs()[0].properties().unwrap();
        assert_eq!(rpr.unknown_children.len(), 1);
        if let crate::raw_xml::RawXmlNode::Element(elem) = &rpr.unknown_children[0].node {
            assert_eq!(elem.name, "w:unknownProp");
            assert_eq!(
                elem.attributes,
                vec![("w:foo".to_string(), "bar".to_string())]
            );
        } else {
            panic!("Expected element node");
        }
    }

    #[test]
    fn test_parse_section_properties() {
        let xml = br#"<?xml version="1.0" encoding="UTF-8"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p>
      <w:r><w:t>Test</w:t></w:r>
    </w:p>
    <w:sectPr>
      <w:pgSz w:w="12240" w:h="15840" w:orient="portrait"/>
      <w:pgMar w:top="1440" w:bottom="1440" w:left="1800" w:right="1800" w:header="720" w:footer="720"/>
    </w:sectPr>
  </w:body>
</w:document>"#;

        let body = parse_document(xml).unwrap();

        // Check that section properties are parsed
        let sect_pr = body
            .section_properties()
            .expect("should have section properties");

        // Check page size
        let pg_sz = sect_pr.page_size.as_ref().expect("should have page size");
        assert_eq!(pg_sz.width, 12240);
        assert_eq!(pg_sz.height, 15840);
        assert_eq!(pg_sz.orientation, PageOrientation::Portrait);

        // Check margins
        let margins = sect_pr.margins.as_ref().expect("should have margins");
        assert_eq!(margins.top, 1440);
        assert_eq!(margins.bottom, 1440);
        assert_eq!(margins.left, 1800);
        assert_eq!(margins.right, 1800);
        assert_eq!(margins.header, Some(720));
        assert_eq!(margins.footer, Some(720));
    }

    #[test]
    fn test_parse_section_properties_landscape() {
        let xml = br#"<?xml version="1.0" encoding="UTF-8"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p><w:r><w:t>Test</w:t></w:r></w:p>
    <w:sectPr>
      <w:pgSz w:w="15840" w:h="12240" w:orient="landscape"/>
    </w:sectPr>
  </w:body>
</w:document>"#;

        let body = parse_document(xml).unwrap();
        let sect_pr = body
            .section_properties()
            .expect("should have section properties");
        let pg_sz = sect_pr.page_size.as_ref().expect("should have page size");

        assert_eq!(pg_sz.width, 15840);
        assert_eq!(pg_sz.height, 12240);
        assert_eq!(pg_sz.orientation, PageOrientation::Landscape);
    }

    #[test]
    fn test_parse_cell_properties() {
        let xml = br#"<?xml version="1.0" encoding="UTF-8"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:tbl>
      <w:tr>
        <w:tc>
          <w:tcPr>
            <w:tcW w:w="2400" w:type="dxa"/>
            <w:gridSpan w:val="2"/>
            <w:vMerge w:val="restart"/>
            <w:shd w:fill="FFFF00" w:val="clear"/>
            <w:vAlign w:val="center"/>
            <w:tcBorders>
              <w:top w:val="single" w:sz="4" w:color="000000"/>
              <w:bottom w:val="double" w:sz="8"/>
            </w:tcBorders>
          </w:tcPr>
          <w:p><w:r><w:t>Cell content</w:t></w:r></w:p>
        </w:tc>
      </w:tr>
    </w:tbl>
  </w:body>
</w:document>"#;

        let body = parse_document(xml).unwrap();
        let table = body.tables().next().expect("should have table");
        let cell = &table.rows()[0].cells()[0];
        let props = cell.properties().expect("should have cell properties");

        // Check width
        let width = props.width.as_ref().expect("should have width");
        assert_eq!(width.width, 2400);
        assert_eq!(width.width_type, WidthType::Dxa);

        // Check grid span
        assert_eq!(props.grid_span, Some(2));

        // Check vertical merge
        assert_eq!(props.vertical_merge, Some(VerticalMerge::Restart));

        // Check shading
        let shading = props.shading.as_ref().expect("should have shading");
        assert_eq!(shading.fill, Some("FFFF00".to_string()));

        // Check vertical alignment
        assert_eq!(props.vertical_align, Some(CellVerticalAlign::Center));

        // Check borders
        let borders = props.borders.as_ref().expect("should have borders");
        let top = borders.top.as_ref().expect("should have top border");
        assert_eq!(top.style, BorderStyle::Single);
        assert_eq!(top.size, Some(4));
        assert_eq!(top.color, Some("000000".to_string()));

        let bottom = borders.bottom.as_ref().expect("should have bottom border");
        assert_eq!(bottom.style, BorderStyle::Double);
        assert_eq!(bottom.size, Some(8));
    }

    #[test]
    fn test_parse_cell_vertical_merge_continue() {
        let xml = br#"<?xml version="1.0" encoding="UTF-8"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:tbl>
      <w:tr>
        <w:tc>
          <w:tcPr>
            <w:vMerge/>
          </w:tcPr>
          <w:p><w:r><w:t>Merged</w:t></w:r></w:p>
        </w:tc>
      </w:tr>
    </w:tbl>
  </w:body>
</w:document>"#;

        let body = parse_document(xml).unwrap();
        let table = body.tables().next().expect("should have table");
        let cell = &table.rows()[0].cells()[0];
        let props = cell.properties().expect("should have cell properties");

        // Empty vMerge means continue
        assert_eq!(props.vertical_merge, Some(VerticalMerge::Continue));
    }
}
