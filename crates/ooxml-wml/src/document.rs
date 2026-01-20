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
    /// Unknown child elements preserved for round-trip fidelity.
    /// Stored with original position index for correct ordering during serialization.
    pub unknown_children: Vec<PositionedNode>,
}

/// Block-level content in the document body.
#[derive(Debug, Clone)]
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
}

/// A table in the document.
///
/// Corresponds to the `<w:tbl>` element.
#[derive(Debug, Clone, Default)]
pub struct Table {
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
    /// Unknown child elements preserved for round-trip fidelity.
    /// Stored with original position index for correct ordering during serialization.
    pub unknown_children: Vec<PositionedNode>,
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
        self.properties.as_ref().is_some_and(|p| p.underline)
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

/// Properties of a text run.
///
/// Corresponds to the `<w:rPr>` element.
#[derive(Debug, Clone, Default)]
pub struct RunProperties {
    /// Bold formatting.
    pub bold: bool,
    /// Italic formatting.
    pub italic: bool,
    /// Underline formatting.
    pub underline: bool,
    /// Strike-through formatting.
    pub strike: bool,
    /// Font size in half-points.
    pub size: Option<u32>,
    /// Font name.
    pub font: Option<String>,
    /// Style ID reference.
    pub style: Option<String>,
    /// Text color as hex RGB (e.g., "FF0000" for red, without # prefix).
    pub color: Option<String>,
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
const EL_SZ: &[u8] = b"sz";
const EL_RFONTS: &[u8] = b"rFonts";
const EL_COLOR: &[u8] = b"color";
const EL_BR: &[u8] = b"br";
const EL_TAB: &[u8] = b"tab";
const EL_TBL: &[u8] = b"tbl";
const EL_TR: &[u8] = b"tr";
const EL_TC: &[u8] = b"tc";

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

// Drawing element names (DrawingML)
const EL_DRAWING: &[u8] = b"drawing";
const EL_INLINE: &[u8] = b"inline";
const EL_EXTENT: &[u8] = b"extent";
const EL_DOCPR: &[u8] = b"docPr";
const EL_BLIP: &[u8] = b"blip";

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

    // Drawing/image parsing state
    let mut current_drawing: Option<Drawing> = None;
    let mut current_image: Option<InlineImage> = None;

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
                    name if name == EL_TR && current_table.is_some() => {
                        current_row = Some(Row::new());
                        row_child_idx = 0;
                        table_child_idx += 1;
                    }
                    name if name == EL_TC && current_row.is_some() => {
                        current_cell = Some(Cell::new());
                        cell_child_idx = 0;
                        row_child_idx += 1;
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
                    _ => {
                        // Only capture unknown elements when we're in a container context
                        // IMPORTANT: Don't capture while inside drawing/inline contexts -
                        // we need to continue parsing to find nested elements like blip
                        let should_capture = in_rpr
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
                            || (in_body && current_table.is_none() && current_para.is_none());

                        if should_capture {
                            // Capture unknown elements for round-trip preservation
                            let raw = RawXmlElement::from_reader(&mut reader, &e)?;
                            let node = RawXmlNode::Element(raw);
                            // Add to the innermost active container with position
                            if in_rpr {
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
                            // Underline is present if <w:u> exists with val != "none"
                            rpr.underline = true;
                            for attr in e.attributes().filter_map(|a| a.ok()) {
                                if (attr.key.as_ref() == b"w:val" || attr.key.as_ref() == b"val")
                                    && attr.value.as_ref() == b"none"
                                {
                                    rpr.underline = false;
                                }
                            }
                        }
                    }
                    name if name == EL_STRIKE && in_rpr => {
                        if let Some(rpr) = current_rpr.as_mut() {
                            rpr.strike = parse_toggle_val(&e);
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
                    _ => {
                        // Only capture unknown self-closing elements when in a container context
                        let should_capture = in_rpr
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
                            if in_rpr {
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
                    name if name == EL_TR => {
                        if let Some(row) = current_row.take()
                            && let Some(table) = current_table.as_mut()
                        {
                            table.rows.push(row);
                        }
                    }
                    name if name == EL_TC => {
                        if let Some(cell) = current_cell.take()
                            && let Some(row) = current_row.as_mut()
                        {
                            row.cells.push(cell);
                        }
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
                    name if name == EL_PPR => {
                        in_ppr = false;
                    }
                    name if name == EL_RPR => {
                        in_rpr = false;
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
}
