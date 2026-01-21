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
use ooxml_opc::{Package, Relationships, rel_type, rels_path_for};
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

    /// Load a header part by its relationship ID.
    ///
    /// Returns the parsed header content. The relationship ID comes from
    /// `SectionProperties.headers[].rel_id`.
    ///
    /// # Example
    /// ```ignore
    /// if let Some(sect_pr) = doc.body().section_properties() {
    ///     for header_ref in &sect_pr.headers {
    ///         let header = doc.get_header(&header_ref.rel_id)?;
    ///         for para in header.paragraphs() {
    ///             println!("{}", para.text());
    ///         }
    ///     }
    /// }
    /// ```
    pub fn get_header(&mut self, rel_id: &str) -> Result<HeaderPart> {
        let rel = self
            .doc_rels
            .get(rel_id)
            .ok_or_else(|| Error::MissingPart(format!("header relationship {}", rel_id)))?;

        let header_path = resolve_path(&self.doc_path, &rel.target);
        let header_xml = self.package.read_part(&header_path)?;
        parse_header_footer(&header_xml, true)
    }

    /// Load a footer part by its relationship ID.
    ///
    /// Returns the parsed footer content. The relationship ID comes from
    /// `SectionProperties.footers[].rel_id`.
    pub fn get_footer(&mut self, rel_id: &str) -> Result<FooterPart> {
        let rel = self
            .doc_rels
            .get(rel_id)
            .ok_or_else(|| Error::MissingPart(format!("footer relationship {}", rel_id)))?;

        let footer_path = resolve_path(&self.doc_path, &rel.target);
        let footer_xml = self.package.read_part(&footer_path)?;
        parse_header_footer(&footer_xml, false)
    }

    /// Load the footnotes part.
    ///
    /// Returns the parsed footnotes. Individual footnotes can be looked up by ID
    /// using `FootnotesPart::get()`.
    ///
    /// Returns `Error::MissingPart` if the document has no footnotes.xml.
    pub fn get_footnotes(&mut self) -> Result<FootnotesPart> {
        // Footnotes are referenced via relationship from document.xml.rels
        let footnotes_rel = self
            .doc_rels
            .get_by_type(
                "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes",
            )
            .ok_or_else(|| Error::MissingPart("footnotes relationship".into()))?;

        let footnotes_path = resolve_path(&self.doc_path, &footnotes_rel.target);
        let footnotes_xml = self.package.read_part(&footnotes_path)?;
        parse_footnotes(&footnotes_xml)
    }

    /// Load the endnotes part.
    ///
    /// Returns the parsed endnotes. Individual endnotes can be looked up by ID
    /// using `EndnotesPart::get()`.
    ///
    /// Returns `Error::MissingPart` if the document has no endnotes.xml.
    pub fn get_endnotes(&mut self) -> Result<EndnotesPart> {
        let endnotes_rel = self
            .doc_rels
            .get_by_type(
                "http://schemas.openxmlformats.org/officeDocument/2006/relationships/endnotes",
            )
            .ok_or_else(|| Error::MissingPart("endnotes relationship".into()))?;

        let endnotes_path = resolve_path(&self.doc_path, &endnotes_rel.target);
        let endnotes_xml = self.package.read_part(&endnotes_path)?;
        parse_footnotes(&endnotes_xml) // Same structure as footnotes
    }

    /// Load the comments part.
    ///
    /// Returns the parsed comments. Individual comments can be looked up by ID
    /// using `CommentsPart::get()`.
    ///
    /// Returns `Error::MissingPart` if the document has no comments.xml.
    pub fn get_comments(&mut self) -> Result<CommentsPart> {
        let comments_rel = self
            .doc_rels
            .get_by_type(
                "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments",
            )
            .ok_or_else(|| Error::MissingPart("comments relationship".into()))?;

        let comments_path = resolve_path(&self.doc_path, &comments_rel.target);
        let comments_xml = self.package.read_part(&comments_path)?;
        parse_comments(&comments_xml)
    }

    /// Load the document settings.
    ///
    /// Returns the parsed settings from word/settings.xml.
    ///
    /// Returns `Error::MissingPart` if the document has no settings.xml.
    pub fn get_settings(&mut self) -> Result<DocumentSettings> {
        let settings_rel = self
            .doc_rels
            .get_by_type(
                "http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings",
            )
            .ok_or_else(|| Error::MissingPart("settings relationship".into()))?;

        let settings_path = resolve_path(&self.doc_path, &settings_rel.target);
        let settings_xml = self.package.read_part(&settings_path)?;
        parse_settings(&settings_xml)
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

/// A header part containing block-level content.
///
/// Corresponds to the `<w:hdr>` element in header*.xml parts.
/// ECMA-376 Part 1, Section 17.10.3 (hdr).
#[derive(Debug, Clone, Default)]
pub struct HeaderPart {
    /// Block-level content (paragraphs and tables).
    content: Vec<BlockContent>,
    /// Unknown child elements preserved for round-trip fidelity.
    pub unknown_children: Vec<PositionedNode>,
}

impl HeaderPart {
    /// Create a new empty header.
    pub fn new() -> Self {
        Self::default()
    }

    /// Get the block-level content.
    pub fn content(&self) -> &[BlockContent] {
        &self.content
    }

    /// Get all paragraphs in the header (including those in tables and content controls).
    pub fn paragraphs(&self) -> Vec<&Paragraph> {
        collect_paragraphs(&self.content)
    }

    /// Extract all text from the header.
    pub fn text(&self) -> String {
        self.paragraphs()
            .iter()
            .map(|p| p.text())
            .collect::<Vec<_>>()
            .join("\n")
    }
}

/// A footer part containing block-level content.
///
/// Corresponds to the `<w:ftr>` element in footer*.xml parts.
/// ECMA-376 Part 1, Section 17.10.2 (ftr).
pub type FooterPart = HeaderPart;

/// Collect paragraphs from block content (including nested in tables/content controls).
fn collect_paragraphs(blocks: &[BlockContent]) -> Vec<&Paragraph> {
    let mut paras = Vec::new();
    for block in blocks {
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
            BlockContent::ContentControl(c) => {
                paras.extend(collect_paragraphs(&c.content));
            }
            BlockContent::CustomXml(c) => {
                paras.extend(collect_paragraphs(&c.content));
            }
        }
    }
    paras
}

/// Block-level content in the document body.
#[derive(Debug, Clone)]
#[allow(clippy::large_enum_variant)] // Table is larger due to properties/grid, but boxing adds indirection
pub enum BlockContent {
    /// A paragraph.
    Paragraph(Paragraph),
    /// A table.
    Table(Table),
    /// A content control (structured document tag).
    ContentControl(ContentControl),
    /// A custom XML block.
    CustomXml(CustomXml),
}

/// A content control (structured document tag).
///
/// Corresponds to the `<w:sdt>` element.
/// ECMA-376 Part 1, Section 17.5.2.29 (sdt).
#[derive(Debug, Clone, Default)]
pub struct ContentControl {
    /// The tag value (used for programmatic access).
    pub tag: Option<String>,
    /// Human-readable alias/title.
    pub alias: Option<String>,
    /// The content type (e.g., "text", "richText", "picture").
    pub content_type: Option<SdtContentType>,
    /// Block-level content within the control.
    pub content: Vec<BlockContent>,
    /// Unknown properties preserved for round-trip fidelity.
    pub unknown_children: Vec<PositionedNode>,
}

/// A custom XML block.
///
/// Corresponds to the `<w:customXml>` element.
/// ECMA-376 Part 1, Section 17.5.1.4 (customXml).
#[derive(Debug, Clone, Default)]
pub struct CustomXml {
    /// The namespace URI of the custom XML.
    pub uri: Option<String>,
    /// The element name within the namespace.
    pub element: Option<String>,
    /// Block-level content within the custom XML.
    pub content: Vec<BlockContent>,
    /// Unknown properties preserved for round-trip fidelity.
    pub unknown_children: Vec<PositionedNode>,
}

impl CustomXml {
    /// Create a new empty custom XML block.
    pub fn new() -> Self {
        Self::default()
    }

    /// Get the block-level content.
    pub fn content(&self) -> &[BlockContent] {
        &self.content
    }

    /// Get all paragraphs in the custom XML block.
    pub fn paragraphs(&self) -> Vec<&Paragraph> {
        collect_paragraphs(&self.content)
    }

    /// Extract all text from the custom XML block.
    pub fn text(&self) -> String {
        self.paragraphs()
            .iter()
            .map(|p| p.text())
            .collect::<Vec<_>>()
            .join("\n")
    }
}

/// Content type for a structured document tag.
///
/// ECMA-376 Part 1, Section 17.5.2.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum SdtContentType {
    /// Plain text content.
    Text,
    /// Rich text content.
    RichText,
    /// Combo box (dropdown list).
    ComboBox,
    /// Drop-down list.
    DropDownList,
    /// Date picker.
    Date,
    /// Document part gallery.
    DocPartGallery,
    /// Picture.
    Picture,
    /// Checkbox.
    CheckBox,
}

impl Body {
    /// Create an empty body.
    pub fn new() -> Self {
        Self::default()
    }

    /// Get all paragraphs in the body (flattened, including those in tables and content controls).
    pub fn paragraphs(&self) -> Vec<&Paragraph> {
        collect_paragraphs(&self.content)
    }

    /// Get block-level content.
    pub fn content(&self) -> &[BlockContent] {
        &self.content
    }

    /// Get mutable reference to block-level content.
    pub fn content_mut(&mut self) -> &mut Vec<BlockContent> {
        &mut self.content
    }

    /// Consume the body and return its parts.
    ///
    /// Used internally for converting parsed body content to HeaderPart/FooterPart.
    pub(crate) fn into_parts(self) -> (Vec<BlockContent>, Vec<PositionedNode>) {
        (self.content, self.unknown_children)
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
                BlockContent::ContentControl(c) => c.text(),
                BlockContent::CustomXml(c) => c.text(),
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

impl ContentControl {
    /// Create a new empty content control.
    pub fn new() -> Self {
        Self::default()
    }

    /// Get the tag value.
    pub fn tag(&self) -> Option<&str> {
        self.tag.as_deref()
    }

    /// Get the alias/title.
    pub fn alias(&self) -> Option<&str> {
        self.alias.as_deref()
    }

    /// Get the content type.
    pub fn content_type(&self) -> Option<SdtContentType> {
        self.content_type
    }

    /// Get the content.
    pub fn content(&self) -> &[BlockContent] {
        &self.content
    }

    /// Get mutable reference to content.
    pub fn content_mut(&mut self) -> &mut Vec<BlockContent> {
        &mut self.content
    }

    /// Extract all text from the content control.
    pub fn text(&self) -> String {
        self.content
            .iter()
            .map(|block| match block {
                BlockContent::Paragraph(p) => p.text(),
                BlockContent::Table(t) => t.text(),
                BlockContent::ContentControl(c) => c.text(),
                BlockContent::CustomXml(c) => c.text(),
            })
            .collect::<Vec<_>>()
            .join("\n")
    }
}

impl SdtContentType {
    /// Parse from the XML element name.
    pub fn parse(name: &str) -> Option<Self> {
        match name {
            "text" => Some(SdtContentType::Text),
            "richText" => Some(SdtContentType::RichText),
            "comboBox" => Some(SdtContentType::ComboBox),
            "dropDownList" => Some(SdtContentType::DropDownList),
            "date" => Some(SdtContentType::Date),
            "docPartGallery" | "docPartObj" => Some(SdtContentType::DocPartGallery),
            "picture" => Some(SdtContentType::Picture),
            "checkbox" => Some(SdtContentType::CheckBox),
            _ => None,
        }
    }

    /// Convert to XML element name.
    pub fn as_str(&self) -> &'static str {
        match self {
            SdtContentType::Text => "text",
            SdtContentType::RichText => "richText",
            SdtContentType::ComboBox => "comboBox",
            SdtContentType::DropDownList => "dropDownList",
            SdtContentType::Date => "date",
            SdtContentType::DocPartGallery => "docPartGallery",
            SdtContentType::Picture => "picture",
            SdtContentType::CheckBox => "checkbox",
        }
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
// Run is larger than other variants, but boxing would be a breaking API change.
// The size difference is acceptable for the ergonomic benefit of direct access.
#[allow(clippy::large_enum_variant)]
#[derive(Debug, Clone)]
pub enum ParagraphContent {
    /// A text run.
    Run(Run),
    /// A hyperlink containing runs.
    Hyperlink(Hyperlink),
    /// Start of a bookmark.
    BookmarkStart(BookmarkStart),
    /// End of a bookmark.
    BookmarkEnd(BookmarkEnd),
    /// Start of a comment range.
    CommentRangeStart(CommentRangeStart),
    /// End of a comment range.
    CommentRangeEnd(CommentRangeEnd),
    /// A simple field (e.g., PAGE, DATE).
    SimpleField(SimpleField),
}

/// A bookmark start marker.
///
/// Corresponds to the `<w:bookmarkStart>` element.
/// ECMA-376 Part 1, Section 17.13.6.2 (bookmarkStart).
#[derive(Debug, Clone)]
pub struct BookmarkStart {
    /// Unique ID for the bookmark (matches bookmarkEnd).
    pub id: u32,
    /// Name of the bookmark.
    pub name: String,
}

/// A bookmark end marker.
///
/// Corresponds to the `<w:bookmarkEnd>` element.
/// ECMA-376 Part 1, Section 17.13.6.1 (bookmarkEnd).
#[derive(Debug, Clone)]
pub struct BookmarkEnd {
    /// Unique ID for the bookmark (matches bookmarkStart).
    pub id: u32,
}

/// A comment range start marker.
///
/// Corresponds to the `<w:commentRangeStart>` element.
/// ECMA-376 Part 1, Section 17.13.4.4 (commentRangeStart).
#[derive(Debug, Clone)]
pub struct CommentRangeStart {
    /// Unique ID for the comment (references comment in comments.xml).
    pub id: u32,
}

/// A comment range end marker.
///
/// Corresponds to the `<w:commentRangeEnd>` element.
/// ECMA-376 Part 1, Section 17.13.4.3 (commentRangeEnd).
#[derive(Debug, Clone)]
pub struct CommentRangeEnd {
    /// Unique ID for the comment (matches commentRangeStart).
    pub id: u32,
}

/// A simple field in the document.
///
/// Corresponds to the `<w:fldSimple>` element.
/// ECMA-376 Part 1, Section 17.16.19 (fldSimple).
#[derive(Debug, Clone, Default)]
pub struct SimpleField {
    /// Field instruction code (e.g., "PAGE", "DATE", "TOC").
    pub instruction: String,
    /// Runs containing the field's displayed result.
    pub runs: Vec<Run>,
}

/// A field character marker in a complex field.
///
/// Corresponds to the `<w:fldChar>` element.
/// ECMA-376 Part 1, Section 17.16.18 (fldChar).
#[derive(Debug, Clone)]
pub struct FieldChar {
    /// Type of field character.
    pub field_type: FieldCharType,
}

/// Type of field character marker.
///
/// ECMA-376 Part 1, Section 17.18.29 (ST_FldCharType).
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum FieldCharType {
    /// Beginning of a complex field.
    Begin,
    /// Separator between field code and field result.
    Separate,
    /// End of a complex field.
    End,
}

impl FieldCharType {
    /// Parse from the w:fldCharType attribute value.
    pub fn parse(s: &str) -> Option<Self> {
        match s {
            "begin" => Some(FieldCharType::Begin),
            "separate" => Some(FieldCharType::Separate),
            "end" => Some(FieldCharType::End),
            _ => None,
        }
    }

    /// Convert to the w:fldCharType attribute value.
    pub fn as_str(&self) -> &'static str {
        match self {
            FieldCharType::Begin => "begin",
            FieldCharType::Separate => "separate",
            FieldCharType::End => "end",
        }
    }
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
                ParagraphContent::SimpleField(field) => {
                    for run in &field.runs {
                        runs.push(run);
                    }
                }
                ParagraphContent::BookmarkStart(_)
                | ParagraphContent::BookmarkEnd(_)
                | ParagraphContent::CommentRangeStart(_)
                | ParagraphContent::CommentRangeEnd(_) => {
                    // Bookmarks and comment ranges don't contain runs
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
                ParagraphContent::SimpleField(field) => {
                    field.runs.iter().map(|r| r.text()).collect::<String>()
                }
                ParagraphContent::BookmarkStart(_)
                | ParagraphContent::BookmarkEnd(_)
                | ParagraphContent::CommentRangeStart(_)
                | ParagraphContent::CommentRangeEnd(_) => String::new(),
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
    /// Section type (how the section break behaves).
    pub section_type: Option<SectionType>,
    /// Page size (width and height).
    pub page_size: Option<PageSize>,
    /// Page margins.
    pub margins: Option<PageMargins>,
    /// Column definitions.
    pub columns: Option<Columns>,
    /// Document grid settings.
    pub doc_grid: Option<DocGrid>,
    /// Header references for this section.
    pub headers: Vec<HeaderFooterRef>,
    /// Footer references for this section.
    pub footers: Vec<HeaderFooterRef>,
    /// Unknown child elements preserved for round-trip fidelity.
    pub unknown_children: Vec<PositionedNode>,
}

/// A reference to a header or footer part.
///
/// Corresponds to the `<w:headerReference>` and `<w:footerReference>` elements.
/// ECMA-376 Part 1, Section 17.10.5 (headerReference, footerReference).
#[derive(Debug, Clone)]
pub struct HeaderFooterRef {
    /// The relationship ID referencing the header/footer part.
    pub rel_id: String,
    /// The type of header/footer (default, first, even).
    pub hf_type: HeaderFooterType,
}

/// Type of header or footer.
///
/// ECMA-376 Part 1, Section 17.18.36 (ST_HdrFtr).
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub enum HeaderFooterType {
    /// Default header/footer used on most pages.
    #[default]
    Default,
    /// Header/footer for the first page only.
    First,
    /// Header/footer for even pages (when different odd/even is enabled).
    Even,
}

impl HeaderFooterType {
    /// Parse from the w:type attribute value.
    pub fn parse(s: &str) -> Self {
        match s {
            "first" => HeaderFooterType::First,
            "even" => HeaderFooterType::Even,
            _ => HeaderFooterType::Default,
        }
    }

    /// Convert to the w:type attribute value.
    pub fn as_str(&self) -> &'static str {
        match self {
            HeaderFooterType::Default => "default",
            HeaderFooterType::First => "first",
            HeaderFooterType::Even => "even",
        }
    }
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

/// Section type - determines how the section break behaves.
///
/// ECMA-376 Part 1, Section 17.18.77 (ST_SectionMark).
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub enum SectionType {
    /// Section begins on the same page (continuous section break).
    Continuous,
    /// Section begins on the next even page.
    EvenPage,
    /// Section begins in the next column (for multi-column sections).
    NextColumn,
    /// Section begins on the next page (default).
    #[default]
    NextPage,
    /// Section begins on the next odd page.
    OddPage,
}

impl SectionType {
    /// Parse from the w:val attribute value.
    pub fn parse(s: &str) -> Self {
        match s {
            "continuous" => SectionType::Continuous,
            "evenPage" => SectionType::EvenPage,
            "nextColumn" => SectionType::NextColumn,
            "oddPage" => SectionType::OddPage,
            _ => SectionType::NextPage,
        }
    }

    /// Convert to the w:val attribute value.
    pub fn as_str(&self) -> &'static str {
        match self {
            SectionType::Continuous => "continuous",
            SectionType::EvenPage => "evenPage",
            SectionType::NextColumn => "nextColumn",
            SectionType::NextPage => "nextPage",
            SectionType::OddPage => "oddPage",
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

/// Column definitions for a section.
///
/// Corresponds to the `<w:cols>` element.
/// ECMA-376 Part 1, Section 17.6.4 (cols).
#[derive(Debug, Clone, Default)]
pub struct Columns {
    /// Number of columns (default 1).
    pub num: Option<u32>,
    /// Space between columns in twips (default 720 = 0.5 inch).
    pub space: Option<u32>,
    /// Whether columns have equal width.
    pub equal_width: bool,
    /// Whether there's a separator line between columns.
    pub separator: bool,
    /// Individual column definitions (used when equal_width is false).
    pub columns: Vec<Column>,
}

/// A single column definition.
///
/// Corresponds to the `<w:col>` element within `<w:cols>`.
#[derive(Debug, Clone, Copy)]
pub struct Column {
    /// Column width in twips.
    pub width: u32,
    /// Space after this column in twips.
    pub space: Option<u32>,
}

/// Document grid settings for a section.
///
/// Corresponds to the `<w:docGrid>` element.
/// ECMA-376 Part 1, Section 17.6.5 (docGrid).
#[derive(Debug, Clone, Copy, Default)]
pub struct DocGrid {
    /// Grid type.
    pub grid_type: DocGridType,
    /// Line pitch (distance between lines) in twips.
    pub line_pitch: Option<u32>,
    /// Character pitch (distance between characters) in twips.
    pub char_space: Option<i32>,
}

/// Document grid type.
///
/// ECMA-376 Part 1, Section 17.18.14 (ST_DocGrid).
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub enum DocGridType {
    /// No document grid.
    #[default]
    Default,
    /// Line grid only.
    Lines,
    /// Line and character grid.
    LinesAndChars,
    /// Snap to characters.
    SnapToChars,
}

impl DocGridType {
    /// Parse from the w:type attribute value.
    pub fn parse(s: &str) -> Self {
        match s {
            "lines" => DocGridType::Lines,
            "linesAndChars" => DocGridType::LinesAndChars,
            "snapToChars" => DocGridType::SnapToChars,
            _ => DocGridType::Default,
        }
    }

    /// Convert to the w:type attribute value.
    pub fn as_str(&self) -> &'static str {
        match self {
            DocGridType::Default => "default",
            DocGridType::Lines => "lines",
            DocGridType::LinesAndChars => "linesAndChars",
            DocGridType::SnapToChars => "snapToChars",
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
    /// Custom tab stops.
    pub tabs: Vec<TabStop>,
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

/// A tab stop definition.
///
/// Corresponds to the `<w:tab>` element within `<w:tabs>`.
/// ECMA-376 Part 1, Section 17.3.1.38 (tab).
#[derive(Debug, Clone, Copy)]
pub struct TabStop {
    /// Position of the tab stop in twips.
    pub position: i32,
    /// Type of tab stop alignment.
    pub tab_type: TabStopType,
    /// Leader character to fill space before tab.
    pub leader: Option<TabLeader>,
}

/// Tab stop alignment type.
///
/// ECMA-376 Part 1, Section 17.18.83 (ST_TabJc).
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub enum TabStopType {
    /// Left-aligned tab stop (default).
    #[default]
    Left,
    /// Center-aligned tab stop.
    Center,
    /// Right-aligned tab stop.
    Right,
    /// Decimal-aligned tab stop.
    Decimal,
    /// Bar tab stop (draws a vertical line).
    Bar,
    /// Clear tab stop at this position.
    Clear,
}

impl TabStopType {
    /// Parse from the w:val attribute value.
    pub fn parse(s: &str) -> Self {
        match s {
            "center" => Self::Center,
            "right" | "end" => Self::Right,
            "decimal" | "num" => Self::Decimal,
            "bar" => Self::Bar,
            "clear" => Self::Clear,
            _ => Self::Left,
        }
    }

    /// Convert to the w:val attribute value.
    pub fn as_str(&self) -> &'static str {
        match self {
            Self::Left => "left",
            Self::Center => "center",
            Self::Right => "right",
            Self::Decimal => "decimal",
            Self::Bar => "bar",
            Self::Clear => "clear",
        }
    }
}

/// Tab leader character.
///
/// ECMA-376 Part 1, Section 17.18.82 (ST_TabTlc).
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum TabLeader {
    /// No leader (default).
    None,
    /// Dot leader (....).
    Dot,
    /// Hyphen leader (----).
    Hyphen,
    /// Underscore leader (____).
    Underscore,
    /// Heavy line leader.
    Heavy,
    /// Middle dot leader (····).
    MiddleDot,
}

impl TabLeader {
    /// Parse from the w:leader attribute value.
    pub fn parse(s: &str) -> Option<Self> {
        match s {
            "none" => Some(Self::None),
            "dot" => Some(Self::Dot),
            "hyphen" => Some(Self::Hyphen),
            "underscore" => Some(Self::Underscore),
            "heavy" => Some(Self::Heavy),
            "middleDot" => Some(Self::MiddleDot),
            _ => None,
        }
    }

    /// Convert to the w:leader attribute value.
    pub fn as_str(&self) -> &'static str {
        match self {
            Self::None => "none",
            Self::Dot => "dot",
            Self::Hyphen => "hyphen",
            Self::Underscore => "underscore",
            Self::Heavy => "heavy",
            Self::MiddleDot => "middleDot",
        }
    }
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
    /// VML pictures (legacy image format) in the run.
    vml_pictures: Vec<VmlPicture>,
    /// Embedded OLE objects in the run.
    embedded_objects: Vec<EmbeddedObject>,
    /// Symbols in the run.
    symbols: Vec<Symbol>,
    /// Field character marker (for complex fields).
    field_char: Option<FieldChar>,
    /// Field instruction text (for complex fields).
    instr_text: Option<String>,
    /// Footnote reference (if this run contains one).
    footnote_ref: Option<FootnoteReference>,
    /// Endnote reference (if this run contains one).
    endnote_ref: Option<EndnoteReference>,
    /// Comment reference (if this run contains one).
    comment_ref: Option<CommentReference>,
    /// Whether this run contains a page break.
    page_break: bool,
    /// Unknown child elements preserved for round-trip fidelity.
    /// Stored with original position index for correct ordering during serialization.
    pub unknown_children: Vec<PositionedNode>,
    /// Unknown attributes preserved for round-trip fidelity.
    /// Stored with original position index for correct ordering during serialization.
    pub unknown_attrs: Vec<PositionedAttr>,
}

/// A VML picture element (legacy image format).
///
/// Corresponds to the `<w:pict>` element which contains VML (Vector Markup Language)
/// content. This is a legacy format used in older Word documents. The content is
/// preserved as a raw XML element for roundtrip fidelity since VML is complex and
/// not fully parsed.
#[derive(Debug, Clone, Default)]
pub struct VmlPicture {
    /// Attributes on the w:pict element.
    pub attributes: Vec<(String, String)>,
    /// Child content preserved as raw XML nodes.
    pub children: Vec<RawXmlNode>,
}

/// An embedded OLE object.
///
/// Corresponds to the `<w:object>` element which contains embedded objects like
/// Excel charts, equations, or other OLE-linked content. The content is preserved
/// as raw XML for roundtrip fidelity since OLE objects are complex.
#[derive(Debug, Clone, Default)]
pub struct EmbeddedObject {
    /// Attributes on the w:object element.
    pub attributes: Vec<(String, String)>,
    /// Child content preserved as raw XML nodes.
    pub children: Vec<RawXmlNode>,
}

/// A drawing element containing images.
///
/// Corresponds to the `<w:drawing>` element. Supports both inline images
/// (positioned within text flow) and anchored images (floating with text wrap).
#[derive(Debug, Clone, Default)]
pub struct Drawing {
    /// Inline images in this drawing.
    images: Vec<InlineImage>,
    /// Anchored (floating) images in this drawing.
    anchored_images: Vec<AnchoredImage>,
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

/// An anchored (floating) image in a drawing.
///
/// Represents an image positioned relative to a reference point (page, column, paragraph, etc.)
/// with text wrapping options. Unlike inline images, anchored images can float and wrap.
/// ECMA-376 Part 1, Section 20.4.2.3 (anchor).
#[derive(Debug, Clone)]
pub struct AnchoredImage {
    /// Relationship ID referencing the image file (e.g., "rId4").
    rel_id: String,
    /// Width in EMUs (English Metric Units). 914400 EMUs = 1 inch.
    width_emu: Option<i64>,
    /// Height in EMUs.
    height_emu: Option<i64>,
    /// Optional description/alt text for the image.
    description: Option<String>,
    /// Whether the image is behind text (true) or in front (false).
    behind_doc: bool,
    /// Horizontal position offset from the reference in EMUs.
    pos_x: i64,
    /// Vertical position offset from the reference in EMUs.
    pos_y: i64,
    /// Text wrapping mode.
    wrap_type: WrapType,
    /// Unknown child elements preserved for round-trip fidelity.
    pub unknown_children: Vec<PositionedNode>,
}

/// Text wrapping type for anchored images.
///
/// ECMA-376 Part 1, Section 20.4.2.3.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub enum WrapType {
    /// No wrapping - text flows over/under the image.
    #[default]
    None,
    /// Square wrapping - text wraps around a bounding box.
    Square,
    /// Tight wrapping - text wraps closely around the image shape.
    Tight,
    /// Through wrapping - text wraps through transparent areas.
    Through,
    /// Top and bottom - text only above and below.
    TopAndBottom,
}

impl WrapType {
    /// Parse a wrap type from its XML element name.
    pub fn from_element(name: &[u8]) -> Option<Self> {
        match name {
            b"wrapNone" => Some(WrapType::None),
            b"wrapSquare" => Some(WrapType::Square),
            b"wrapTight" => Some(WrapType::Tight),
            b"wrapThrough" => Some(WrapType::Through),
            b"wrapTopAndBottom" => Some(WrapType::TopAndBottom),
            _ => None,
        }
    }

    /// Get the XML element name for this wrap type.
    pub fn as_element(&self) -> &'static str {
        match self {
            WrapType::None => "wrapNone",
            WrapType::Square => "wrapSquare",
            WrapType::Tight => "wrapTight",
            WrapType::Through => "wrapThrough",
            WrapType::TopAndBottom => "wrapTopAndBottom",
        }
    }
}

/// Image data loaded from the package.
#[derive(Debug, Clone)]
pub struct ImageData {
    /// MIME content type (e.g., "image/png", "image/jpeg").
    pub content_type: String,
    /// Raw image bytes.
    pub data: Vec<u8>,
}

/// A symbol character from a specific font.
///
/// Corresponds to the `<w:sym>` element.
/// ECMA-376 Part 1, Section 17.3.3.30 (sym).
#[derive(Debug, Clone)]
pub struct Symbol {
    /// Font name containing the symbol.
    pub font: String,
    /// Character code (typically hex, e.g., "F020").
    pub char_code: String,
}

/// A reference to a footnote.
///
/// Corresponds to the `<w:footnoteReference>` element.
/// ECMA-376 Part 1, Section 17.11.13 (footnoteReference).
#[derive(Debug, Clone)]
pub struct FootnoteReference {
    /// The footnote ID (references a footnote in word/footnotes.xml).
    pub id: u32,
}

/// A reference to an endnote.
///
/// Corresponds to the `<w:endnoteReference>` element.
/// ECMA-376 Part 1, Section 17.11.7 (endnoteReference).
#[derive(Debug, Clone)]
pub struct EndnoteReference {
    /// The endnote ID (references an endnote in word/endnotes.xml).
    pub id: u32,
}

/// A reference to a comment (the comment mark/bubble).
///
/// Corresponds to the `<w:commentReference>` element.
/// ECMA-376 Part 1, Section 17.13.4.5 (commentReference).
#[derive(Debug, Clone)]
pub struct CommentReference {
    /// The comment ID (references a comment in word/comments.xml).
    pub id: u32,
}

/// A single footnote.
///
/// Corresponds to the `<w:footnote>` element in word/footnotes.xml.
/// ECMA-376 Part 1, Section 17.11.10 (footnote).
#[derive(Debug, Clone)]
pub struct Footnote {
    /// The footnote ID (referenced by FootnoteReference).
    pub id: i32,
    /// The type of footnote (normal, separator, continuationSeparator).
    pub footnote_type: Option<String>,
    /// Block-level content (paragraphs, tables).
    content: Vec<BlockContent>,
}

impl Footnote {
    /// Get the block-level content.
    pub fn content(&self) -> &[BlockContent] {
        &self.content
    }

    /// Get all paragraphs in the footnote.
    pub fn paragraphs(&self) -> Vec<&Paragraph> {
        collect_paragraphs(&self.content)
    }

    /// Extract all text from the footnote.
    pub fn text(&self) -> String {
        self.paragraphs()
            .iter()
            .map(|p| p.text())
            .collect::<Vec<_>>()
            .join("\n")
    }
}

/// The footnotes part containing all footnotes.
///
/// Corresponds to the `<w:footnotes>` element in word/footnotes.xml.
#[derive(Debug, Clone, Default)]
pub struct FootnotesPart {
    /// All footnotes in the document.
    footnotes: Vec<Footnote>,
}

impl FootnotesPart {
    /// Get all footnotes.
    pub fn footnotes(&self) -> &[Footnote] {
        &self.footnotes
    }

    /// Get a footnote by its ID.
    pub fn get(&self, id: i32) -> Option<&Footnote> {
        self.footnotes.iter().find(|f| f.id == id)
    }
}

/// A single endnote.
///
/// Corresponds to the `<w:endnote>` element in word/endnotes.xml.
/// ECMA-376 Part 1, Section 17.11.2 (endnote).
pub type Endnote = Footnote;

/// The endnotes part containing all endnotes.
///
/// Corresponds to the `<w:endnotes>` element in word/endnotes.xml.
pub type EndnotesPart = FootnotesPart;

/// A single comment.
///
/// Corresponds to the `<w:comment>` element in word/comments.xml.
/// ECMA-376 Part 1, Section 17.13.4.2 (comment).
#[derive(Debug, Clone)]
pub struct Comment {
    /// The comment ID (referenced by CommentReference and comment ranges).
    pub id: i32,
    /// The author of the comment.
    pub author: Option<String>,
    /// The date/time of the comment.
    pub date: Option<String>,
    /// The author's initials.
    pub initials: Option<String>,
    /// Block-level content (paragraphs, tables).
    content: Vec<BlockContent>,
}

impl Comment {
    /// Get the block-level content.
    pub fn content(&self) -> &[BlockContent] {
        &self.content
    }

    /// Get all paragraphs in the comment.
    pub fn paragraphs(&self) -> Vec<&Paragraph> {
        collect_paragraphs(&self.content)
    }

    /// Extract all text from the comment.
    pub fn text(&self) -> String {
        self.paragraphs()
            .iter()
            .map(|p| p.text())
            .collect::<Vec<_>>()
            .join("\n")
    }
}

/// The comments part containing all comments.
///
/// Corresponds to the `<w:comments>` element in word/comments.xml.
#[derive(Debug, Clone, Default)]
pub struct CommentsPart {
    /// All comments in the document.
    comments: Vec<Comment>,
}

impl CommentsPart {
    /// Get all comments.
    pub fn comments(&self) -> &[Comment] {
        &self.comments
    }

    /// Get a comment by its ID.
    pub fn get(&self, id: i32) -> Option<&Comment> {
        self.comments.iter().find(|c| c.id == id)
    }
}

/// Document settings.
///
/// Corresponds to the `<w:settings>` element in word/settings.xml.
/// Contains document-wide settings like default tab stop, zoom level,
/// tracking changes settings, and various compatibility options.
///
/// ECMA-376 Part 1, Section 17.15.1 (Document Settings).
#[derive(Debug, Clone, Default)]
pub struct DocumentSettings {
    /// Default tab stop width in twentieths of a point (twips).
    /// Default is 720 twips (0.5 inch).
    pub default_tab_stop: Option<u32>,
    /// Document zoom percentage (e.g., 100 = 100%).
    pub zoom_percent: Option<u32>,
    /// Whether to display the document background shape.
    pub display_background_shape: bool,
    /// Whether track revisions (track changes) is enabled.
    pub track_revisions: bool,
    /// Whether to track moves separately in tracked changes.
    pub do_not_track_moves: bool,
    /// Whether to track formatting in tracked changes.
    pub do_not_track_formatting: bool,
    /// Spelling state - whether document has been fully checked.
    pub spelling_state: Option<ProofState>,
    /// Grammar state - whether document has been fully checked.
    pub grammar_state: Option<ProofState>,
    /// Character spacing control mode.
    pub character_spacing_control: Option<CharacterSpacingControl>,
    /// Compatibility mode (e.g., "15" for Word 2013+).
    pub compat_mode: Option<u32>,
    /// Unknown child elements preserved for round-trip fidelity.
    pub unknown_children: Vec<PositionedNode>,
}

/// Proof state for spelling/grammar checking.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ProofState {
    /// Document is clean (fully checked).
    Clean,
    /// Document is dirty (needs checking).
    Dirty,
}

/// Character spacing control mode.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum CharacterSpacingControl {
    /// Do not compress punctuation.
    DoNotCompress,
    /// Compress punctuation.
    CompressPunctuation,
    /// Compress punctuation and kana.
    CompressPunctuationAndJapaneseKana,
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

    /// Get VML pictures (legacy images) in this run.
    pub fn vml_pictures(&self) -> &[VmlPicture] {
        &self.vml_pictures
    }

    /// Get mutable reference to VML pictures.
    pub fn vml_pictures_mut(&mut self) -> &mut Vec<VmlPicture> {
        &mut self.vml_pictures
    }

    /// Get embedded OLE objects in this run.
    pub fn embedded_objects(&self) -> &[EmbeddedObject] {
        &self.embedded_objects
    }

    /// Get mutable reference to embedded objects.
    pub fn embedded_objects_mut(&mut self) -> &mut Vec<EmbeddedObject> {
        &mut self.embedded_objects
    }

    /// Check if this run contains any images (including VML pictures).
    pub fn has_images(&self) -> bool {
        self.drawings.iter().any(|d| !d.images.is_empty()) || !self.vml_pictures.is_empty()
    }

    /// Check if this run contains any embedded objects.
    pub fn has_embedded_objects(&self) -> bool {
        !self.embedded_objects.is_empty()
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

    /// Get symbols in this run.
    pub fn symbols(&self) -> &[Symbol] {
        &self.symbols
    }

    /// Add a symbol to this run.
    pub fn add_symbol(&mut self, symbol: Symbol) {
        self.symbols.push(symbol);
    }

    /// Get field character marker if present.
    pub fn field_char(&self) -> Option<&FieldChar> {
        self.field_char.as_ref()
    }

    /// Set field character marker.
    pub fn set_field_char(&mut self, field_char: FieldChar) {
        self.field_char = Some(field_char);
    }

    /// Get field instruction text if present.
    pub fn instr_text(&self) -> Option<&str> {
        self.instr_text.as_deref()
    }

    /// Set field instruction text.
    pub fn set_instr_text(&mut self, text: impl Into<String>) {
        self.instr_text = Some(text.into());
    }

    /// Get footnote reference if present.
    pub fn footnote_ref(&self) -> Option<&FootnoteReference> {
        self.footnote_ref.as_ref()
    }

    /// Set footnote reference.
    pub fn set_footnote_ref(&mut self, footnote_ref: FootnoteReference) {
        self.footnote_ref = Some(footnote_ref);
    }

    /// Get endnote reference if present.
    pub fn endnote_ref(&self) -> Option<&EndnoteReference> {
        self.endnote_ref.as_ref()
    }

    /// Set endnote reference.
    pub fn set_endnote_ref(&mut self, endnote_ref: EndnoteReference) {
        self.endnote_ref = Some(endnote_ref);
    }

    /// Get comment reference if present.
    pub fn comment_ref(&self) -> Option<&CommentReference> {
        self.comment_ref.as_ref()
    }

    /// Set comment reference.
    pub fn set_comment_ref(&mut self, comment_ref: CommentReference) {
        self.comment_ref = Some(comment_ref);
    }
}

impl Drawing {
    /// Create an empty drawing.
    pub fn new() -> Self {
        Self::default()
    }

    /// Get inline images in this drawing.
    pub fn images(&self) -> &[InlineImage] {
        &self.images
    }

    /// Get mutable reference to inline images.
    pub fn images_mut(&mut self) -> &mut Vec<InlineImage> {
        &mut self.images
    }

    /// Add an inline image to this drawing.
    pub fn add_image(&mut self, rel_id: impl Into<String>) -> &mut InlineImage {
        self.images.push(InlineImage::new(rel_id));
        self.images.last_mut().unwrap()
    }

    /// Get anchored (floating) images in this drawing.
    pub fn anchored_images(&self) -> &[AnchoredImage] {
        &self.anchored_images
    }

    /// Get mutable reference to anchored images.
    pub fn anchored_images_mut(&mut self) -> &mut Vec<AnchoredImage> {
        &mut self.anchored_images
    }

    /// Add an anchored image to this drawing.
    pub fn add_anchored_image(&mut self, rel_id: impl Into<String>) -> &mut AnchoredImage {
        self.anchored_images.push(AnchoredImage::new(rel_id));
        self.anchored_images.last_mut().unwrap()
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

impl AnchoredImage {
    /// Create a new anchored image with the given relationship ID.
    pub fn new(rel_id: impl Into<String>) -> Self {
        Self {
            rel_id: rel_id.into(),
            width_emu: None,
            height_emu: None,
            description: None,
            behind_doc: false,
            pos_x: 0,
            pos_y: 0,
            wrap_type: WrapType::None,
            unknown_children: Vec::new(),
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

    /// Check if the image is behind document text.
    pub fn is_behind_doc(&self) -> bool {
        self.behind_doc
    }

    /// Set whether the image is behind document text.
    pub fn set_behind_doc(&mut self, behind: bool) -> &mut Self {
        self.behind_doc = behind;
        self
    }

    /// Get horizontal position offset in EMUs.
    pub fn pos_x(&self) -> i64 {
        self.pos_x
    }

    /// Get vertical position offset in EMUs.
    pub fn pos_y(&self) -> i64 {
        self.pos_y
    }

    /// Set horizontal position offset in EMUs.
    pub fn set_pos_x(&mut self, emu: i64) -> &mut Self {
        self.pos_x = emu;
        self
    }

    /// Set vertical position offset in EMUs.
    pub fn set_pos_y(&mut self, emu: i64) -> &mut Self {
        self.pos_y = emu;
        self
    }

    /// Get the text wrapping type.
    pub fn wrap_type(&self) -> WrapType {
        self.wrap_type
    }

    /// Set the text wrapping type.
    pub fn set_wrap_type(&mut self, wrap: WrapType) -> &mut Self {
        self.wrap_type = wrap;
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
    /// Hidden text (w:vanish).
    pub hidden: bool,
    /// Run shading/background.
    pub shading: Option<CellShading>,
    /// Unknown child elements preserved for round-trip fidelity.
    /// Stored with original position index for correct ordering during serialization.
    pub unknown_children: Vec<PositionedNode>,
}

// XML element names (local names)
const EL_BODY: &[u8] = b"body";
const EL_HDR: &[u8] = b"hdr";
const EL_FTR: &[u8] = b"ftr";
const EL_FOOTNOTES: &[u8] = b"footnotes";
const EL_ENDNOTES: &[u8] = b"endnotes";
const EL_FOOTNOTE: &[u8] = b"footnote";
const EL_ENDNOTE: &[u8] = b"endnote";
const EL_COMMENTS: &[u8] = b"comments";
const EL_COMMENT: &[u8] = b"comment";
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
const EL_VANISH: &[u8] = b"vanish";
const EL_SZ: &[u8] = b"sz";
const EL_RFONTS: &[u8] = b"rFonts";
const EL_COLOR: &[u8] = b"color";
const EL_BR: &[u8] = b"br";
const EL_TAB: &[u8] = b"tab";
const EL_SYM: &[u8] = b"sym";
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

// Bookmark elements
const EL_BOOKMARK_START: &[u8] = b"bookmarkStart";
const EL_BOOKMARK_END: &[u8] = b"bookmarkEnd";

// Comment range elements
const EL_COMMENT_RANGE_START: &[u8] = b"commentRangeStart";
const EL_COMMENT_RANGE_END: &[u8] = b"commentRangeEnd";

// Field elements
const EL_FLD_SIMPLE: &[u8] = b"fldSimple";
const EL_FLD_CHAR: &[u8] = b"fldChar";
const EL_INSTR_TEXT: &[u8] = b"instrText";

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
const EL_TABS: &[u8] = b"tabs";
const EL_TAB_DEF: &[u8] = b"tab"; // Tab definition in w:tabs (not the tab character in runs)
const EL_BETWEEN: &[u8] = b"between";
const EL_BAR: &[u8] = b"bar";

// Drawing element names (DrawingML)
const EL_DRAWING: &[u8] = b"drawing";
const EL_INLINE: &[u8] = b"inline";
const EL_ANCHOR: &[u8] = b"anchor";
const EL_EXTENT: &[u8] = b"extent";
const EL_DOCPR: &[u8] = b"docPr";
const EL_BLIP: &[u8] = b"blip";
const EL_POS_H: &[u8] = b"positionH";
const EL_POS_V: &[u8] = b"positionV";
const EL_POS_OFFSET: &[u8] = b"posOffset";

// VML picture element (legacy image format)
const EL_PICT: &[u8] = b"pict";

// Embedded object element
const EL_OBJECT: &[u8] = b"object";

// Section properties elements
const EL_SECT_PR: &[u8] = b"sectPr";
const EL_PG_SZ: &[u8] = b"pgSz";
const EL_PG_MAR: &[u8] = b"pgMar";
const EL_SECT_TYPE: &[u8] = b"type";
const EL_COLS: &[u8] = b"cols";
const EL_COL: &[u8] = b"col";
const EL_DOC_GRID: &[u8] = b"docGrid";
const EL_HEADER_REFERENCE: &[u8] = b"headerReference";
const EL_FOOTER_REFERENCE: &[u8] = b"footerReference";
const EL_FOOTNOTE_REFERENCE: &[u8] = b"footnoteReference";
const EL_ENDNOTE_REFERENCE: &[u8] = b"endnoteReference";
const EL_COMMENT_REFERENCE: &[u8] = b"commentReference";

// Document settings elements
const EL_SETTINGS: &[u8] = b"settings";
const EL_DEFAULT_TAB_STOP: &[u8] = b"defaultTabStop";
const EL_ZOOM: &[u8] = b"zoom";
const EL_DISPLAY_BACKGROUND_SHAPE: &[u8] = b"displayBackgroundShape";
const EL_TRACK_REVISIONS: &[u8] = b"trackRevisions";
const EL_DO_NOT_TRACK_MOVES: &[u8] = b"doNotTrackMoves";
const EL_DO_NOT_TRACK_FORMATTING: &[u8] = b"doNotTrackFormatting";
const EL_PROOF_STATE: &[u8] = b"proofState";
const EL_CHAR_SPACE_CONTROL: &[u8] = b"characterSpacingControl";
const EL_COMPAT: &[u8] = b"compat";
const EL_COMPAT_SETTING: &[u8] = b"compatSetting";

// Content control (SDT) elements
const EL_SDT: &[u8] = b"sdt";
const EL_SDT_PR: &[u8] = b"sdtPr";
const EL_SDT_CONTENT: &[u8] = b"sdtContent";
const EL_TAG: &[u8] = b"tag";
const EL_ALIAS: &[u8] = b"alias";
const EL_CUSTOM_XML: &[u8] = b"customXml";

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
    let mut in_instr_text = false;

    // Table parsing state
    let mut current_table: Option<Table> = None;
    let mut current_row: Option<Row> = None;
    let mut current_cell: Option<Cell> = None;

    // Hyperlink parsing state
    let mut current_hyperlink: Option<Hyperlink> = None;

    // Simple field parsing state
    let mut current_fld_simple: Option<SimpleField> = None;

    // Content control (SDT) parsing state
    let mut current_sdt: Option<ContentControl> = None;
    let mut in_sdt_pr = false;
    let mut in_sdt_content = false;
    let mut _sdt_child_idx: usize = 0;

    // Custom XML parsing state
    let mut current_custom_xml: Option<CustomXml> = None;

    // Numbering parsing state
    let mut in_numpr = false;
    let mut current_numid: Option<u32> = None;
    let mut current_ilvl: Option<u32> = None;

    // Paragraph border parsing state
    let mut in_p_bdr = false;
    let mut current_p_borders: Option<ParagraphBorders> = None;

    // Tab stops parsing state
    let mut in_tabs = false;
    let mut current_tabs: Vec<TabStop> = Vec::new();

    // Drawing/image parsing state
    let mut current_drawing: Option<Drawing> = None;
    let mut current_image: Option<InlineImage> = None;
    let mut current_anchored_image: Option<AnchoredImage> = None;
    let mut in_pos_h = false;
    let mut in_pos_v = false;
    let mut in_pos_offset = false;

    // Section properties parsing state
    let mut current_sect_pr: Option<SectionProperties> = None;
    let mut in_sect_pr = false;
    let mut sect_pr_child_idx: usize = 0;

    // Columns parsing state (within section properties)
    let mut current_cols: Option<Columns> = None;
    let mut in_cols = false;

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
                    name if name == EL_BODY || name == EL_HDR || name == EL_FTR => {
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
                    // Content control (SDT)
                    name if name == EL_SDT && in_body => {
                        current_sdt = Some(ContentControl::new());
                        _sdt_child_idx = 0;
                        if current_cell.is_some() {
                            cell_child_idx += 1;
                        } else {
                            body_child_idx += 1;
                        }
                    }
                    // SDT properties
                    name if name == EL_SDT_PR && current_sdt.is_some() => {
                        in_sdt_pr = true;
                        _sdt_child_idx += 1;
                    }
                    // SDT content
                    name if name == EL_SDT_CONTENT && current_sdt.is_some() => {
                        in_sdt_content = true;
                        _sdt_child_idx += 1;
                    }
                    // Custom XML block
                    name if name == EL_CUSTOM_XML && in_body => {
                        let mut custom_xml = CustomXml::new();
                        for attr in e.attributes().flatten() {
                            let key = local_name(attr.key.as_ref());
                            if key == b"uri" {
                                custom_xml.uri =
                                    std::str::from_utf8(&attr.value).ok().map(String::from);
                            } else if key == b"element" {
                                custom_xml.element =
                                    std::str::from_utf8(&attr.value).ok().map(String::from);
                            }
                        }
                        current_custom_xml = Some(custom_xml);
                        if current_cell.is_some() {
                            cell_child_idx += 1;
                        } else {
                            body_child_idx += 1;
                        }
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
                    // Simple field (w:fldSimple)
                    name if name == EL_FLD_SIMPLE && current_para.is_some() => {
                        let mut field = SimpleField::default();
                        for attr in e.attributes().filter_map(|a| a.ok()) {
                            let key = attr.key.as_ref();
                            if key == b"w:instr" || key == b"instr" {
                                field.instruction =
                                    String::from_utf8_lossy(&attr.value).into_owned();
                            }
                        }
                        current_fld_simple = Some(field);
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
                    name if name == EL_INSTR_TEXT && current_run.is_some() => {
                        in_instr_text = true;
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
                    name if name == EL_TABS && in_ppr => {
                        in_tabs = true;
                        current_tabs.clear();
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
                    name if name == EL_PICT && current_run.is_some() => {
                        // VML picture (legacy image format) - capture entire content
                        let attributes = e
                            .attributes()
                            .filter_map(|a| a.ok())
                            .map(|a| {
                                (
                                    String::from_utf8_lossy(a.key.as_ref()).to_string(),
                                    String::from_utf8_lossy(&a.value).to_string(),
                                )
                            })
                            .collect();

                        // Parse children using RawXmlElement helper approach
                        let pict_elem = RawXmlElement::from_reader(&mut reader, &e)?;

                        let vml_pict = VmlPicture {
                            attributes,
                            children: pict_elem.children,
                        };

                        if let Some(run) = current_run.as_mut() {
                            run.vml_pictures.push(vml_pict);
                        }
                        run_child_idx += 1;
                    }
                    name if name == EL_OBJECT && current_run.is_some() => {
                        // Embedded OLE object - capture entire content
                        let attributes = e
                            .attributes()
                            .filter_map(|a| a.ok())
                            .map(|a| {
                                (
                                    String::from_utf8_lossy(a.key.as_ref()).to_string(),
                                    String::from_utf8_lossy(&a.value).to_string(),
                                )
                            })
                            .collect();

                        // Parse children using RawXmlElement helper approach
                        let obj_elem = RawXmlElement::from_reader(&mut reader, &e)?;

                        let embedded_obj = EmbeddedObject {
                            attributes,
                            children: obj_elem.children,
                        };

                        if let Some(run) = current_run.as_mut() {
                            run.embedded_objects.push(embedded_obj);
                        }
                        run_child_idx += 1;
                    }
                    name if name == EL_INLINE && current_drawing.is_some() => {
                        // Start of an inline image - create with placeholder rel_id
                        current_image = Some(InlineImage::new(""));
                    }
                    name if name == EL_ANCHOR && current_drawing.is_some() => {
                        // Start of an anchored (floating) image
                        let mut img = AnchoredImage::new("");
                        // Parse anchor attributes
                        for attr in e.attributes().filter_map(|a| a.ok()) {
                            let key = attr.key.as_ref();
                            if key == b"behindDoc"
                                && let Ok(s) = std::str::from_utf8(&attr.value)
                            {
                                img.behind_doc = matches!(s, "1" | "true");
                            }
                        }
                        current_anchored_image = Some(img);
                    }
                    name if name == EL_POS_H && current_anchored_image.is_some() => {
                        in_pos_h = true;
                    }
                    name if name == EL_POS_V && current_anchored_image.is_some() => {
                        in_pos_v = true;
                    }
                    name if name == EL_POS_OFFSET && (in_pos_h || in_pos_v) => {
                        in_pos_offset = true;
                    }
                    name if name == EL_SECT_PR && in_body => {
                        // Section properties at document level (defines last section)
                        current_sect_pr = Some(SectionProperties::default());
                        in_sect_pr = true;
                        sect_pr_child_idx = 0;
                        body_child_idx += 1;
                    }
                    // Start of columns definition (w:cols)
                    name if name == EL_COLS && in_sect_pr => {
                        let mut cols = Columns::default();
                        for attr in e.attributes().filter_map(|a| a.ok()) {
                            let key = attr.key.as_ref();
                            if let Ok(s) = std::str::from_utf8(&attr.value) {
                                match key {
                                    b"w:num" | b"num" => {
                                        cols.num = s.parse().ok();
                                    }
                                    b"w:space" | b"space" => {
                                        cols.space = s.parse().ok();
                                    }
                                    b"w:equalWidth" | b"equalWidth" => {
                                        cols.equal_width = matches!(s, "true" | "1" | "on");
                                    }
                                    b"w:sep" | b"sep" => {
                                        cols.separator = matches!(s, "true" | "1" | "on");
                                    }
                                    _ => {}
                                }
                            }
                        }
                        current_cols = Some(cols);
                        in_cols = true;
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
                    // Symbol character
                    name if name == EL_SYM && current_run.is_some() => {
                        let mut font = String::new();
                        let mut char_code = String::new();
                        for attr in e.attributes().filter_map(|a| a.ok()) {
                            let key = attr.key.as_ref();
                            if let Ok(s) = std::str::from_utf8(&attr.value) {
                                match key {
                                    b"w:font" | b"font" => {
                                        font = s.to_string();
                                    }
                                    b"w:char" | b"char" => {
                                        char_code = s.to_string();
                                    }
                                    _ => {}
                                }
                            }
                        }
                        if let Some(run) = current_run.as_mut() {
                            run.symbols.push(Symbol { font, char_code });
                        }
                    }
                    // Field character (complex field marker)
                    name if name == EL_FLD_CHAR && current_run.is_some() => {
                        for attr in e.attributes().filter_map(|a| a.ok()) {
                            let key = attr.key.as_ref();
                            if (key == b"w:fldCharType" || key == b"fldCharType")
                                && let Ok(s) = std::str::from_utf8(&attr.value)
                                && let Some(field_type) = FieldCharType::parse(s)
                                && let Some(run) = current_run.as_mut()
                            {
                                run.field_char = Some(FieldChar { field_type });
                            }
                        }
                    }
                    // Footnote reference
                    name if name == EL_FOOTNOTE_REFERENCE && current_run.is_some() => {
                        for attr in e.attributes().filter_map(|a| a.ok()) {
                            let key = attr.key.as_ref();
                            if (key == b"w:id" || key == b"id")
                                && let Ok(s) = std::str::from_utf8(&attr.value)
                                && let Ok(id) = s.parse::<u32>()
                                && let Some(run) = current_run.as_mut()
                            {
                                run.footnote_ref = Some(FootnoteReference { id });
                            }
                        }
                    }
                    // Endnote reference
                    name if name == EL_ENDNOTE_REFERENCE && current_run.is_some() => {
                        for attr in e.attributes().filter_map(|a| a.ok()) {
                            let key = attr.key.as_ref();
                            if (key == b"w:id" || key == b"id")
                                && let Ok(s) = std::str::from_utf8(&attr.value)
                                && let Ok(id) = s.parse::<u32>()
                                && let Some(run) = current_run.as_mut()
                            {
                                run.endnote_ref = Some(EndnoteReference { id });
                            }
                        }
                    }
                    // Comment reference
                    name if name == EL_COMMENT_REFERENCE && current_run.is_some() => {
                        for attr in e.attributes().filter_map(|a| a.ok()) {
                            let key = attr.key.as_ref();
                            if (key == b"w:id" || key == b"id")
                                && let Ok(s) = std::str::from_utf8(&attr.value)
                                && let Ok(id) = s.parse::<u32>()
                                && let Some(run) = current_run.as_mut()
                            {
                                run.comment_ref = Some(CommentReference { id });
                            }
                        }
                    }
                    // SDT tag
                    name if name == EL_TAG && in_sdt_pr => {
                        for attr in e.attributes().filter_map(|a| a.ok()) {
                            let key = attr.key.as_ref();
                            if (key == b"w:val" || key == b"val")
                                && let Ok(s) = std::str::from_utf8(&attr.value)
                                && let Some(sdt) = current_sdt.as_mut()
                            {
                                sdt.tag = Some(s.to_string());
                            }
                        }
                    }
                    // SDT alias
                    name if name == EL_ALIAS && in_sdt_pr => {
                        for attr in e.attributes().filter_map(|a| a.ok()) {
                            let key = attr.key.as_ref();
                            if (key == b"w:val" || key == b"val")
                                && let Ok(s) = std::str::from_utf8(&attr.value)
                                && let Some(sdt) = current_sdt.as_mut()
                            {
                                sdt.alias = Some(s.to_string());
                            }
                        }
                    }
                    // Bookmark start
                    name if name == EL_BOOKMARK_START && current_para.is_some() => {
                        let mut id = 0u32;
                        let mut bookmark_name = String::new();
                        for attr in e.attributes().filter_map(|a| a.ok()) {
                            let key = attr.key.as_ref();
                            if let Ok(s) = std::str::from_utf8(&attr.value) {
                                match key {
                                    b"w:id" | b"id" => {
                                        id = s.parse().unwrap_or(0);
                                    }
                                    b"w:name" | b"name" => {
                                        bookmark_name = s.to_string();
                                    }
                                    _ => {}
                                }
                            }
                        }
                        if let Some(para) = current_para.as_mut() {
                            para.content
                                .push(ParagraphContent::BookmarkStart(BookmarkStart {
                                    id,
                                    name: bookmark_name,
                                }));
                            para_child_idx += 1;
                        }
                    }
                    // Bookmark end
                    name if name == EL_BOOKMARK_END && current_para.is_some() => {
                        let mut id = 0u32;
                        for attr in e.attributes().filter_map(|a| a.ok()) {
                            if (attr.key.as_ref() == b"w:id" || attr.key.as_ref() == b"id")
                                && let Ok(s) = std::str::from_utf8(&attr.value)
                            {
                                id = s.parse().unwrap_or(0);
                            }
                        }
                        if let Some(para) = current_para.as_mut() {
                            para.content
                                .push(ParagraphContent::BookmarkEnd(BookmarkEnd { id }));
                            para_child_idx += 1;
                        }
                    }
                    // Comment range start
                    name if name == EL_COMMENT_RANGE_START && current_para.is_some() => {
                        let mut id = 0u32;
                        for attr in e.attributes().filter_map(|a| a.ok()) {
                            if (attr.key.as_ref() == b"w:id" || attr.key.as_ref() == b"id")
                                && let Ok(s) = std::str::from_utf8(&attr.value)
                            {
                                id = s.parse().unwrap_or(0);
                            }
                        }
                        if let Some(para) = current_para.as_mut() {
                            para.content.push(ParagraphContent::CommentRangeStart(
                                CommentRangeStart { id },
                            ));
                            para_child_idx += 1;
                        }
                    }
                    // Comment range end
                    name if name == EL_COMMENT_RANGE_END && current_para.is_some() => {
                        let mut id = 0u32;
                        for attr in e.attributes().filter_map(|a| a.ok()) {
                            if (attr.key.as_ref() == b"w:id" || attr.key.as_ref() == b"id")
                                && let Ok(s) = std::str::from_utf8(&attr.value)
                            {
                                id = s.parse().unwrap_or(0);
                            }
                        }
                        if let Some(para) = current_para.as_mut() {
                            para.content
                                .push(ParagraphContent::CommentRangeEnd(CommentRangeEnd { id }));
                            para_child_idx += 1;
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
                    // Tab stop definition (w:tab in w:tabs)
                    name if name == EL_TAB_DEF && in_tabs => {
                        let mut position: i32 = 0;
                        let mut tab_type = TabStopType::Left;
                        let mut leader = None;
                        for attr in e.attributes().filter_map(|a| a.ok()) {
                            let key = attr.key.as_ref();
                            if let Ok(s) = std::str::from_utf8(&attr.value) {
                                match key {
                                    b"w:pos" | b"pos" => {
                                        position = s.parse().unwrap_or(0);
                                    }
                                    b"w:val" | b"val" => {
                                        tab_type = TabStopType::parse(s);
                                    }
                                    b"w:leader" | b"leader" => {
                                        leader = TabLeader::parse(s);
                                    }
                                    _ => {}
                                }
                            }
                        }
                        current_tabs.push(TabStop {
                            position,
                            tab_type,
                            leader,
                        });
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
                    name if name == EL_VANISH && in_rpr => {
                        if let Some(rpr) = current_rpr.as_mut() {
                            rpr.hidden = parse_toggle_val(&e);
                        }
                    }
                    // Run shading (w:shd in rPr)
                    name if name == EL_SHD && in_rpr => {
                        if let Some(rpr) = current_rpr.as_mut() {
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
                            rpr.shading = Some(shading);
                            rpr_child_idx += 1;
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
                    // Image extent (dimensions) - for inline images
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
                    // Image extent (dimensions) - for anchored images
                    name if name == EL_EXTENT && current_anchored_image.is_some() => {
                        if let Some(img) = current_anchored_image.as_mut() {
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
                    // Image properties (description/alt text) - for inline images
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
                    // Image properties (description/alt text) - for anchored images
                    name if name == EL_DOCPR && current_anchored_image.is_some() => {
                        if let Some(img) = current_anchored_image.as_mut() {
                            for attr in e.attributes().filter_map(|a| a.ok()) {
                                if attr.key.as_ref() == b"descr" {
                                    img.description =
                                        Some(String::from_utf8_lossy(&attr.value).into_owned());
                                }
                            }
                        }
                    }
                    // Blip - contains the relationship ID to the image (inline)
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
                    // Blip - contains the relationship ID to the image (anchored)
                    name if name == EL_BLIP && current_anchored_image.is_some() => {
                        if let Some(img) = current_anchored_image.as_mut() {
                            for attr in e.attributes().filter_map(|a| a.ok()) {
                                // r:embed attribute contains the relationship ID
                                if attr.key.as_ref() == b"r:embed" || attr.key.as_ref() == b"embed"
                                {
                                    img.rel_id = String::from_utf8_lossy(&attr.value).into_owned();
                                }
                            }
                        }
                    }
                    // Wrap type for anchored images
                    name if current_anchored_image.is_some() => {
                        if let Some(wrap_type) = WrapType::from_element(name)
                            && let Some(img) = current_anchored_image.as_mut()
                        {
                            img.wrap_type = wrap_type;
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
                    // Section type (w:type)
                    name if name == EL_SECT_TYPE && in_sect_pr => {
                        if let Some(sect_pr) = current_sect_pr.as_mut() {
                            for attr in e.attributes().filter_map(|a| a.ok()) {
                                let key = attr.key.as_ref();
                                if let Ok(s) = std::str::from_utf8(&attr.value)
                                    && (key == b"w:val" || key == b"val")
                                {
                                    sect_pr.section_type = Some(SectionType::parse(s));
                                }
                            }
                            sect_pr_child_idx += 1;
                        }
                    }
                    // Columns (w:cols) - self-closing case
                    name if name == EL_COLS && in_sect_pr && !in_cols => {
                        if let Some(sect_pr) = current_sect_pr.as_mut() {
                            let mut cols = Columns::default();
                            for attr in e.attributes().filter_map(|a| a.ok()) {
                                let key = attr.key.as_ref();
                                if let Ok(s) = std::str::from_utf8(&attr.value) {
                                    match key {
                                        b"w:num" | b"num" => {
                                            cols.num = s.parse().ok();
                                        }
                                        b"w:space" | b"space" => {
                                            cols.space = s.parse().ok();
                                        }
                                        b"w:equalWidth" | b"equalWidth" => {
                                            cols.equal_width = matches!(s, "true" | "1" | "on");
                                        }
                                        b"w:sep" | b"sep" => {
                                            cols.separator = matches!(s, "true" | "1" | "on");
                                        }
                                        _ => {}
                                    }
                                }
                            }
                            sect_pr.columns = Some(cols);
                            sect_pr_child_idx += 1;
                        }
                    }
                    // Individual column definition (w:col)
                    name if name == EL_COL && in_cols => {
                        if let Some(cols) = current_cols.as_mut() {
                            let mut width = 0u32;
                            let mut space = None;
                            for attr in e.attributes().filter_map(|a| a.ok()) {
                                let key = attr.key.as_ref();
                                if let Ok(s) = std::str::from_utf8(&attr.value) {
                                    match key {
                                        b"w:w" | b"w" => {
                                            width = s.parse().unwrap_or(0);
                                        }
                                        b"w:space" | b"space" => {
                                            space = s.parse().ok();
                                        }
                                        _ => {}
                                    }
                                }
                            }
                            cols.columns.push(Column { width, space });
                        }
                    }
                    // Document grid (w:docGrid)
                    name if name == EL_DOC_GRID && in_sect_pr => {
                        if let Some(sect_pr) = current_sect_pr.as_mut() {
                            let mut doc_grid = DocGrid::default();
                            for attr in e.attributes().filter_map(|a| a.ok()) {
                                let key = attr.key.as_ref();
                                if let Ok(s) = std::str::from_utf8(&attr.value) {
                                    match key {
                                        b"w:type" | b"type" => {
                                            doc_grid.grid_type = DocGridType::parse(s);
                                        }
                                        b"w:linePitch" | b"linePitch" => {
                                            doc_grid.line_pitch = s.parse().ok();
                                        }
                                        b"w:charSpace" | b"charSpace" => {
                                            doc_grid.char_space = s.parse().ok();
                                        }
                                        _ => {}
                                    }
                                }
                            }
                            sect_pr.doc_grid = Some(doc_grid);
                            sect_pr_child_idx += 1;
                        }
                    }
                    // Header reference (w:headerReference)
                    name if name == EL_HEADER_REFERENCE && in_sect_pr => {
                        if let Some(sect_pr) = current_sect_pr.as_mut() {
                            let mut rel_id = String::new();
                            let mut hf_type = HeaderFooterType::Default;
                            for attr in e.attributes().filter_map(|a| a.ok()) {
                                let key = attr.key.as_ref();
                                if let Ok(s) = std::str::from_utf8(&attr.value) {
                                    match key {
                                        b"r:id" | b"id" => {
                                            rel_id = s.to_string();
                                        }
                                        b"w:type" | b"type" => {
                                            hf_type = HeaderFooterType::parse(s);
                                        }
                                        _ => {}
                                    }
                                }
                            }
                            if !rel_id.is_empty() {
                                sect_pr.headers.push(HeaderFooterRef { rel_id, hf_type });
                            }
                            sect_pr_child_idx += 1;
                        }
                    }
                    // Footer reference (w:footerReference)
                    name if name == EL_FOOTER_REFERENCE && in_sect_pr => {
                        if let Some(sect_pr) = current_sect_pr.as_mut() {
                            let mut rel_id = String::new();
                            let mut hf_type = HeaderFooterType::Default;
                            for attr in e.attributes().filter_map(|a| a.ok()) {
                                let key = attr.key.as_ref();
                                if let Ok(s) = std::str::from_utf8(&attr.value) {
                                    match key {
                                        b"r:id" | b"id" => {
                                            rel_id = s.to_string();
                                        }
                                        b"w:type" | b"type" => {
                                            hf_type = HeaderFooterType::parse(s);
                                        }
                                        _ => {}
                                    }
                                }
                            }
                            if !rel_id.is_empty() {
                                sect_pr.footers.push(HeaderFooterRef { rel_id, hf_type });
                            }
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
                let text = e.decode().unwrap_or_default();
                if in_text {
                    if let Some(run) = current_run.as_mut() {
                        run.text.push_str(&text);
                    }
                } else if in_instr_text && let Some(run) = current_run.as_mut() {
                    if let Some(ref mut instr) = run.instr_text {
                        instr.push_str(&text);
                    } else {
                        run.instr_text = Some(text.into_owned());
                    }
                } else if in_pos_offset && let Some(img) = current_anchored_image.as_mut() {
                    // Parse position offset value for anchored images
                    if let Ok(offset) = text.parse::<i64>() {
                        if in_pos_h {
                            img.pos_x = offset;
                        } else if in_pos_v {
                            img.pos_y = offset;
                        }
                    }
                }
            }
            Ok(Event::End(e)) => {
                let name = e.name();
                let local = local_name(name.as_ref());
                match local {
                    name if name == EL_BODY || name == EL_HDR || name == EL_FTR => {
                        in_body = false;
                    }
                    name if name == EL_TBL => {
                        if let Some(table) = current_table.take() {
                            if in_sdt_content {
                                if let Some(sdt) = current_sdt.as_mut() {
                                    sdt.content.push(BlockContent::Table(table));
                                }
                            } else if let Some(custom_xml) = current_custom_xml.as_mut() {
                                custom_xml.content.push(BlockContent::Table(table));
                            } else {
                                body.content.push(BlockContent::Table(table));
                            }
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
                            // Add to cell if inside table, SDT if inside content control,
                            // custom XML if inside that, otherwise to body
                            if let Some(cell) = current_cell.as_mut() {
                                cell.paragraphs.push(para);
                            } else if in_sdt_content {
                                if let Some(sdt) = current_sdt.as_mut() {
                                    sdt.content.push(BlockContent::Paragraph(para));
                                }
                            } else if let Some(custom_xml) = current_custom_xml.as_mut() {
                                custom_xml.content.push(BlockContent::Paragraph(para));
                            } else {
                                body.content.push(BlockContent::Paragraph(para));
                            }
                        }
                    }
                    name if name == EL_R && current_run.is_some() => {
                        if let Some(mut run) = current_run.take() {
                            run.properties = current_rpr.take();
                            // Add run to hyperlink/field if inside one, otherwise to paragraph
                            if let Some(hyperlink) = current_hyperlink.as_mut() {
                                hyperlink.runs.push(run);
                            } else if let Some(field) = current_fld_simple.as_mut() {
                                field.runs.push(run);
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
                    name if name == EL_FLD_SIMPLE => {
                        if let Some(field) = current_fld_simple.take()
                            && let Some(para) = current_para.as_mut()
                        {
                            para.content.push(ParagraphContent::SimpleField(field));
                        }
                    }
                    name if name == EL_T => {
                        in_text = false;
                    }
                    name if name == EL_INSTR_TEXT => {
                        in_instr_text = false;
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
                    name if name == EL_TABS => {
                        if let Some(ppr) = current_ppr.as_mut() {
                            ppr.tabs = std::mem::take(&mut current_tabs);
                        }
                        in_tabs = false;
                    }
                    name if name == EL_PPR => {
                        in_ppr = false;
                    }
                    name if name == EL_RPR => {
                        in_rpr = false;
                    }
                    // End of columns definition - add to section properties
                    name if name == EL_COLS => {
                        if let Some(cols) = current_cols.take()
                            && let Some(sect_pr) = current_sect_pr.as_mut()
                        {
                            sect_pr.columns = Some(cols);
                            sect_pr_child_idx += 1;
                        }
                        in_cols = false;
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
                    // End of anchored image - add to current drawing
                    name if name == EL_ANCHOR => {
                        if let Some(img) = current_anchored_image.take()
                            && !img.rel_id.is_empty()
                            && let Some(drawing) = current_drawing.as_mut()
                        {
                            drawing.anchored_images.push(img);
                        }
                    }
                    // End of position offset
                    name if name == EL_POS_OFFSET => {
                        in_pos_offset = false;
                    }
                    // End of horizontal position
                    name if name == EL_POS_H => {
                        in_pos_h = false;
                    }
                    // End of vertical position
                    name if name == EL_POS_V => {
                        in_pos_v = false;
                    }
                    // End of drawing - add to current run
                    name if name == EL_DRAWING => {
                        if let Some(drawing) = current_drawing.take()
                            && let Some(run) = current_run.as_mut()
                        {
                            run.drawings.push(drawing);
                        }
                    }
                    // End of SDT properties
                    name if name == EL_SDT_PR => {
                        in_sdt_pr = false;
                    }
                    // End of SDT content
                    name if name == EL_SDT_CONTENT => {
                        in_sdt_content = false;
                    }
                    // End of SDT - add to body or cell
                    name if name == EL_SDT => {
                        if let Some(sdt) = current_sdt.take() {
                            if let Some(cell) = current_cell.as_mut() {
                                // SDTs in table cells - add content directly to cell
                                // (cell doesn't support SDT directly, so flatten)
                                for block in sdt.content {
                                    if let BlockContent::Paragraph(para) = block {
                                        cell.paragraphs.push(para);
                                    }
                                }
                            } else {
                                body.content.push(BlockContent::ContentControl(sdt));
                            }
                        }
                    }
                    name if name == EL_CUSTOM_XML => {
                        if let Some(custom_xml) = current_custom_xml.take() {
                            if let Some(cell) = current_cell.as_mut() {
                                // Custom XML in table cells - flatten content
                                for block in custom_xml.content {
                                    if let BlockContent::Paragraph(para) = block {
                                        cell.paragraphs.push(para);
                                    }
                                }
                            } else {
                                body.content.push(BlockContent::CustomXml(custom_xml));
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

    Ok(body)
}

/// Parse a header or footer part.
///
/// Headers (`<w:hdr>`) and footers (`<w:ftr>`) have the same structure:
/// block-level content (paragraphs, tables) without section properties.
fn parse_header_footer(xml: &[u8], _is_header: bool) -> Result<HeaderPart> {
    // Headers/footers use the same content model as body
    // We reuse parse_document which now also handles hdr/ftr elements
    let body = parse_document(xml)?;
    let (content, unknown_children) = body.into_parts();

    Ok(HeaderPart {
        content,
        unknown_children,
    })
}

/// Parse a footnotes or endnotes part.
///
/// The structure is: `<w:footnotes>` containing `<w:footnote w:id="...">` elements,
/// each of which contains paragraphs.
fn parse_footnotes(xml: &[u8]) -> Result<FootnotesPart> {
    use quick_xml::Reader;
    use quick_xml::events::Event;

    let mut reader = Reader::from_reader(xml);
    reader.config_mut().trim_text(false);

    let mut buf = Vec::new();
    let mut footnotes = Vec::new();
    let mut current_footnote: Option<(i32, Option<String>, Vec<u8>)> = None;
    let mut depth = 0;

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) => {
                let name = e.name();
                let local = local_name(name.as_ref());
                if (local == EL_FOOTNOTE || local == EL_ENDNOTE) && current_footnote.is_none() {
                    // Parse footnote/endnote attributes
                    let mut id: i32 = 0;
                    let mut fn_type: Option<String> = None;
                    for attr in e.attributes().flatten() {
                        let key = local_name(attr.key.as_ref());
                        if key == b"id" {
                            if let Ok(s) = std::str::from_utf8(&attr.value) {
                                id = s.parse().unwrap_or(0);
                            }
                        } else if key == b"type" {
                            fn_type = std::str::from_utf8(&attr.value).ok().map(String::from);
                        }
                    }
                    // Start capturing inner XML
                    current_footnote = Some((id, fn_type, Vec::new()));
                    depth = 1;
                } else if current_footnote.is_some() {
                    depth += 1;
                    // Append this start tag to the captured content
                    if let Some((_, _, ref mut content)) = current_footnote {
                        content.extend_from_slice(b"<");
                        content.extend_from_slice(e.name().as_ref());
                        for attr in e.attributes().flatten() {
                            content.extend_from_slice(b" ");
                            content.extend_from_slice(attr.key.as_ref());
                            content.extend_from_slice(b"=\"");
                            content.extend_from_slice(&attr.value);
                            content.extend_from_slice(b"\"");
                        }
                        content.extend_from_slice(b">");
                    }
                }
            }
            Ok(Event::Empty(e)) => {
                if let Some((_, _, ref mut content)) = current_footnote {
                    content.extend_from_slice(b"<");
                    content.extend_from_slice(e.name().as_ref());
                    for attr in e.attributes().flatten() {
                        content.extend_from_slice(b" ");
                        content.extend_from_slice(attr.key.as_ref());
                        content.extend_from_slice(b"=\"");
                        content.extend_from_slice(&attr.value);
                        content.extend_from_slice(b"\"");
                    }
                    content.extend_from_slice(b"/>");
                }
            }
            Ok(Event::End(e)) => {
                let name = e.name();
                let local = local_name(name.as_ref());
                if current_footnote.is_some() {
                    depth -= 1;
                    if depth == 0 {
                        // End of footnote - parse the captured content
                        if let Some((id, fn_type, inner_xml)) = current_footnote.take() {
                            // Wrap in a body element for parsing
                            let mut wrapped = b"<?xml version=\"1.0\"?><w:body xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\">".to_vec();
                            wrapped.extend_from_slice(&inner_xml);
                            wrapped.extend_from_slice(b"</w:body>");

                            if let Ok(body) = parse_document(&wrapped) {
                                let (content, _) = body.into_parts();
                                footnotes.push(Footnote {
                                    id,
                                    footnote_type: fn_type,
                                    content,
                                });
                            }
                        }
                    } else {
                        // Append end tag to captured content
                        if let Some((_, _, ref mut content)) = current_footnote {
                            content.extend_from_slice(b"</");
                            content.extend_from_slice(e.name().as_ref());
                            content.extend_from_slice(b">");
                        }
                    }
                }
                if local == EL_FOOTNOTES || local == EL_ENDNOTES {
                    break;
                }
            }
            Ok(Event::Text(e)) => {
                if let Some((_, _, ref mut content)) = current_footnote {
                    content.extend_from_slice(&e);
                }
            }
            Ok(Event::Eof) => break,
            Err(e) => return Err(Error::Xml(e)),
            _ => {}
        }
        buf.clear();
    }

    Ok(FootnotesPart { footnotes })
}

/// Parse a comments part.
///
/// The structure is: `<w:comments>` containing `<w:comment w:id="...">` elements,
/// each of which contains paragraphs.
fn parse_comments(xml: &[u8]) -> Result<CommentsPart> {
    use quick_xml::Reader;
    use quick_xml::events::Event;

    let mut reader = Reader::from_reader(xml);
    reader.config_mut().trim_text(false);

    let mut buf = Vec::new();
    let mut comments = Vec::new();
    #[allow(clippy::type_complexity)]
    let mut current_comment: Option<(
        i32,
        Option<String>,
        Option<String>,
        Option<String>,
        Vec<u8>,
    )> = None;
    let mut depth = 0;

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) => {
                let name = e.name();
                let local = local_name(name.as_ref());
                if local == EL_COMMENT && current_comment.is_none() {
                    // Parse comment attributes
                    let mut id: i32 = 0;
                    let mut author: Option<String> = None;
                    let mut date: Option<String> = None;
                    let mut initials: Option<String> = None;
                    for attr in e.attributes().flatten() {
                        let key = local_name(attr.key.as_ref());
                        if key == b"id" {
                            if let Ok(s) = std::str::from_utf8(&attr.value) {
                                id = s.parse().unwrap_or(0);
                            }
                        } else if key == b"author" {
                            author = std::str::from_utf8(&attr.value).ok().map(String::from);
                        } else if key == b"date" {
                            date = std::str::from_utf8(&attr.value).ok().map(String::from);
                        } else if key == b"initials" {
                            initials = std::str::from_utf8(&attr.value).ok().map(String::from);
                        }
                    }
                    current_comment = Some((id, author, date, initials, Vec::new()));
                    depth = 1;
                } else if current_comment.is_some() {
                    depth += 1;
                    if let Some((_, _, _, _, ref mut content)) = current_comment {
                        content.extend_from_slice(b"<");
                        content.extend_from_slice(e.name().as_ref());
                        for attr in e.attributes().flatten() {
                            content.extend_from_slice(b" ");
                            content.extend_from_slice(attr.key.as_ref());
                            content.extend_from_slice(b"=\"");
                            content.extend_from_slice(&attr.value);
                            content.extend_from_slice(b"\"");
                        }
                        content.extend_from_slice(b">");
                    }
                }
            }
            Ok(Event::Empty(e)) => {
                if let Some((_, _, _, _, ref mut content)) = current_comment {
                    content.extend_from_slice(b"<");
                    content.extend_from_slice(e.name().as_ref());
                    for attr in e.attributes().flatten() {
                        content.extend_from_slice(b" ");
                        content.extend_from_slice(attr.key.as_ref());
                        content.extend_from_slice(b"=\"");
                        content.extend_from_slice(&attr.value);
                        content.extend_from_slice(b"\"");
                    }
                    content.extend_from_slice(b"/>");
                }
            }
            Ok(Event::End(e)) => {
                let name = e.name();
                let local = local_name(name.as_ref());
                if current_comment.is_some() {
                    depth -= 1;
                    if depth == 0 {
                        if let Some((id, author, date, initials, inner_xml)) =
                            current_comment.take()
                        {
                            let mut wrapped = b"<?xml version=\"1.0\"?><w:body xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\">".to_vec();
                            wrapped.extend_from_slice(&inner_xml);
                            wrapped.extend_from_slice(b"</w:body>");

                            if let Ok(body) = parse_document(&wrapped) {
                                let (content, _) = body.into_parts();
                                comments.push(Comment {
                                    id,
                                    author,
                                    date,
                                    initials,
                                    content,
                                });
                            }
                        }
                    } else if let Some((_, _, _, _, ref mut content)) = current_comment {
                        content.extend_from_slice(b"</");
                        content.extend_from_slice(e.name().as_ref());
                        content.extend_from_slice(b">");
                    }
                }
                if local == EL_COMMENTS {
                    break;
                }
            }
            Ok(Event::Text(e)) => {
                if let Some((_, _, _, _, ref mut content)) = current_comment {
                    content.extend_from_slice(&e);
                }
            }
            Ok(Event::Eof) => break,
            Err(e) => return Err(Error::Xml(e)),
            _ => {}
        }
        buf.clear();
    }

    Ok(CommentsPart { comments })
}

/// Parse a settings.xml file into DocumentSettings.
///
/// The structure is `<w:settings>` containing various setting elements.
fn parse_settings(xml: &[u8]) -> Result<DocumentSettings> {
    use quick_xml::Reader;
    use quick_xml::events::Event;

    let mut reader = Reader::from_reader(xml);
    reader.config_mut().trim_text(false);

    let mut buf = Vec::new();
    let mut settings = DocumentSettings::default();
    let mut in_settings = false;
    let mut in_compat = false;
    let mut child_idx: usize = 0;

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) => {
                let name = e.name();
                let local = local_name(name.as_ref());
                if local == EL_SETTINGS {
                    in_settings = true;
                } else if in_settings && local == EL_COMPAT {
                    in_compat = true;
                    child_idx += 1;
                } else if in_settings && !in_compat {
                    // Unknown element - preserve for roundtrip
                    let node = RawXmlElement::from_reader(&mut reader, &e)?;
                    settings
                        .unknown_children
                        .push(PositionedNode::new(child_idx, RawXmlNode::Element(node)));
                    child_idx += 1;
                }
            }
            Ok(Event::Empty(e)) => {
                let name = e.name();
                let local = local_name(name.as_ref());

                if !in_settings {
                    continue;
                }

                if local == EL_DEFAULT_TAB_STOP {
                    for attr in e.attributes().flatten() {
                        let key = local_name(attr.key.as_ref());
                        if key == b"val"
                            && let Ok(s) = std::str::from_utf8(&attr.value)
                        {
                            settings.default_tab_stop = s.parse().ok();
                        }
                    }
                    child_idx += 1;
                } else if local == EL_ZOOM {
                    for attr in e.attributes().flatten() {
                        let key = local_name(attr.key.as_ref());
                        if key == b"percent"
                            && let Ok(s) = std::str::from_utf8(&attr.value)
                        {
                            settings.zoom_percent = s.parse().ok();
                        }
                    }
                    child_idx += 1;
                } else if local == EL_DISPLAY_BACKGROUND_SHAPE {
                    settings.display_background_shape = parse_toggle_val(&e);
                    child_idx += 1;
                } else if local == EL_TRACK_REVISIONS {
                    settings.track_revisions = parse_toggle_val(&e);
                    child_idx += 1;
                } else if local == EL_DO_NOT_TRACK_MOVES {
                    settings.do_not_track_moves = parse_toggle_val(&e);
                    child_idx += 1;
                } else if local == EL_DO_NOT_TRACK_FORMATTING {
                    settings.do_not_track_formatting = parse_toggle_val(&e);
                    child_idx += 1;
                } else if local == EL_PROOF_STATE {
                    for attr in e.attributes().flatten() {
                        let key = local_name(attr.key.as_ref());
                        if let Ok(s) = std::str::from_utf8(&attr.value) {
                            if key == b"spelling" {
                                settings.spelling_state = match s {
                                    "clean" => Some(ProofState::Clean),
                                    "dirty" => Some(ProofState::Dirty),
                                    _ => None,
                                };
                            } else if key == b"grammar" {
                                settings.grammar_state = match s {
                                    "clean" => Some(ProofState::Clean),
                                    "dirty" => Some(ProofState::Dirty),
                                    _ => None,
                                };
                            }
                        }
                    }
                    child_idx += 1;
                } else if local == EL_CHAR_SPACE_CONTROL {
                    for attr in e.attributes().flatten() {
                        let key = local_name(attr.key.as_ref());
                        if key == b"val"
                            && let Ok(s) = std::str::from_utf8(&attr.value)
                        {
                            settings.character_spacing_control = match s {
                                "doNotCompress" => Some(CharacterSpacingControl::DoNotCompress),
                                "compressPunctuation" => {
                                    Some(CharacterSpacingControl::CompressPunctuation)
                                }
                                "compressPunctuationAndJapaneseKana" => Some(
                                    CharacterSpacingControl::CompressPunctuationAndJapaneseKana,
                                ),
                                _ => None,
                            };
                        }
                    }
                    child_idx += 1;
                } else if in_compat && local == EL_COMPAT_SETTING {
                    // Look for w:name="compatibilityMode"
                    for attr in e.attributes().flatten() {
                        let key = local_name(attr.key.as_ref());
                        if key == b"name" && attr.value.as_ref() == b"compatibilityMode" {
                            // Get the w:val attribute
                            for attr2 in e.attributes().flatten() {
                                let key2 = local_name(attr2.key.as_ref());
                                if key2 == b"val"
                                    && let Ok(s) = std::str::from_utf8(&attr2.value)
                                {
                                    settings.compat_mode = s.parse().ok();
                                }
                            }
                        }
                    }
                } else if !in_compat {
                    // Unknown empty element - preserve for roundtrip
                    let node = RawXmlElement::from_empty(&e);
                    settings
                        .unknown_children
                        .push(PositionedNode::new(child_idx, RawXmlNode::Element(node)));
                    child_idx += 1;
                }
            }
            Ok(Event::End(e)) => {
                let name = e.name();
                let local = local_name(name.as_ref());
                if local == EL_SETTINGS {
                    break;
                } else if local == EL_COMPAT {
                    in_compat = false;
                }
            }
            Ok(Event::Eof) => break,
            Err(e) => return Err(Error::Xml(e)),
            _ => {}
        }
        buf.clear();
    }

    Ok(settings)
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

    #[test]
    fn test_parse_header() {
        let xml = br#"<?xml version="1.0" encoding="UTF-8"?>
<w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:p>
    <w:r><w:t>Header content</w:t></w:r>
  </w:p>
</w:hdr>"#;

        let header = parse_header_footer(xml, true).unwrap();

        assert_eq!(header.paragraphs().len(), 1);
        assert_eq!(header.text(), "Header content");
    }

    #[test]
    fn test_parse_footer() {
        let xml = br#"<?xml version="1.0" encoding="UTF-8"?>
<w:ftr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:p>
    <w:r><w:t>Page </w:t></w:r>
  </w:p>
  <w:p>
    <w:r><w:t>Footer line 2</w:t></w:r>
  </w:p>
</w:ftr>"#;

        let footer = parse_header_footer(xml, false).unwrap();

        assert_eq!(footer.paragraphs().len(), 2);
        assert_eq!(footer.text(), "Page \nFooter line 2");
    }

    #[test]
    fn test_parse_footnotes() {
        let xml = br#"<?xml version="1.0" encoding="UTF-8"?>
<w:footnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:footnote w:type="separator" w:id="-1">
    <w:p><w:r><w:separator/></w:r></w:p>
  </w:footnote>
  <w:footnote w:id="1">
    <w:p><w:r><w:t>First footnote content.</w:t></w:r></w:p>
  </w:footnote>
  <w:footnote w:id="2">
    <w:p><w:r><w:t>Second footnote.</w:t></w:r></w:p>
  </w:footnote>
</w:footnotes>"#;

        let footnotes = parse_footnotes(xml).unwrap();

        assert_eq!(footnotes.footnotes().len(), 3);

        // Check the separator footnote
        let sep = footnotes.get(-1).expect("should have separator");
        assert_eq!(sep.footnote_type, Some("separator".to_string()));

        // Check the first real footnote
        let fn1 = footnotes.get(1).expect("should have footnote 1");
        assert_eq!(fn1.id, 1);
        assert_eq!(fn1.text(), "First footnote content.");

        // Check the second footnote
        let fn2 = footnotes.get(2).expect("should have footnote 2");
        assert_eq!(fn2.text(), "Second footnote.");
    }

    #[test]
    fn test_parse_comments() {
        let xml = br#"<?xml version="1.0" encoding="UTF-8"?>
<w:comments xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:comment w:id="0" w:author="John Doe" w:date="2024-01-15T10:30:00Z" w:initials="JD">
    <w:p><w:r><w:t>This needs revision.</w:t></w:r></w:p>
  </w:comment>
  <w:comment w:id="1" w:author="Jane Smith">
    <w:p><w:r><w:t>I agree with John.</w:t></w:r></w:p>
    <w:p><w:r><w:t>Let's discuss tomorrow.</w:t></w:r></w:p>
  </w:comment>
</w:comments>"#;

        let comments = parse_comments(xml).unwrap();

        assert_eq!(comments.comments().len(), 2);

        // Check the first comment
        let c0 = comments.get(0).expect("should have comment 0");
        assert_eq!(c0.id, 0);
        assert_eq!(c0.author, Some("John Doe".to_string()));
        assert_eq!(c0.date, Some("2024-01-15T10:30:00Z".to_string()));
        assert_eq!(c0.initials, Some("JD".to_string()));
        assert_eq!(c0.text(), "This needs revision.");

        // Check the second comment (multiple paragraphs)
        let c1 = comments.get(1).expect("should have comment 1");
        assert_eq!(c1.author, Some("Jane Smith".to_string()));
        assert_eq!(c1.date, None);
        assert_eq!(c1.paragraphs().len(), 2);
        assert_eq!(c1.text(), "I agree with John.\nLet's discuss tomorrow.");
    }
}
