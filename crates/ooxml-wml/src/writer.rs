//! Document writing and serialization.
//!
//! This module provides functionality for creating new Word documents
//! and saving existing documents.

use crate::document::{
    BlockContent, Body, Border, Cell, CellBorders, CellProperties, CellShading, CellWidth, Drawing,
    Hyperlink, InlineImage, NumberingProperties, PageOrientation, Paragraph, ParagraphContent,
    ParagraphProperties, Row, Run, RunProperties, SectionProperties, Table, VerticalMerge,
};
use crate::error::Result;
use crate::raw_xml::{PositionedAttr, PositionedNode, RawXmlNode};
use crate::styles::Styles;
use ooxml::{PackageWriter, Relationship, Relationships, content_type, rel_type};
use std::collections::HashMap;
use std::fs::File;
use std::io::{BufWriter, Seek, Write};
use std::path::Path;

/// WordprocessingML namespace.
pub const NS_W: &str = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
/// Relationships namespace.
pub const NS_R: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
/// WordprocessingML Drawing namespace.
pub const NS_WP: &str = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing";
/// DrawingML main namespace.
pub const NS_A: &str = "http://schemas.openxmlformats.org/drawingml/2006/main";
/// Picture namespace.
pub const NS_PIC: &str = "http://schemas.openxmlformats.org/drawingml/2006/picture";

/// A pending image to be written to the package.
#[derive(Clone)]
pub struct PendingImage {
    /// Raw image data.
    pub data: Vec<u8>,
    /// Content type (e.g., "image/png").
    pub content_type: String,
    /// Assigned relationship ID.
    pub rel_id: String,
    /// Generated filename (e.g., "image1.png").
    pub filename: String,
}

/// A pending hyperlink to be written to relationships.
#[derive(Clone)]
pub struct PendingHyperlink {
    /// Relationship ID.
    pub rel_id: String,
    /// Target URL.
    pub url: String,
}

/// List type for creating numbered or bulleted lists.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ListType {
    /// Bulleted list (uses bullet character).
    Bullet,
    /// Numbered list (uses decimal numbers: 1, 2, 3...).
    Decimal,
    /// Lowercase letter list (a, b, c...).
    LowerLetter,
    /// Uppercase letter list (A, B, C...).
    UpperLetter,
    /// Lowercase Roman numerals (i, ii, iii...).
    LowerRoman,
    /// Uppercase Roman numerals (I, II, III...).
    UpperRoman,
}

/// A numbering definition to be written to numbering.xml.
#[derive(Clone)]
pub struct PendingNumbering {
    /// Abstract numbering ID.
    pub abstract_num_id: u32,
    /// Concrete numbering ID (used in numPr).
    pub num_id: u32,
    /// List type.
    pub list_type: ListType,
}

/// Builder for creating new Word documents.
pub struct DocumentBuilder {
    body: Body,
    _styles: Styles, // TODO: serialize styles.xml
    /// Pending images to write, keyed by rel_id.
    images: HashMap<String, PendingImage>,
    /// Pending hyperlinks, keyed by rel_id.
    hyperlinks: HashMap<String, PendingHyperlink>,
    /// Numbering definitions, keyed by num_id.
    numberings: HashMap<u32, PendingNumbering>,
    /// Counter for generating unique IDs.
    next_rel_id: u32,
    /// Counter for generating unique numbering IDs.
    next_num_id: u32,
}

impl Default for DocumentBuilder {
    fn default() -> Self {
        Self::new()
    }
}

impl DocumentBuilder {
    /// Create a new document builder.
    pub fn new() -> Self {
        Self {
            body: Body::new(),
            _styles: Styles::new(),
            images: HashMap::new(),
            hyperlinks: HashMap::new(),
            numberings: HashMap::new(),
            next_rel_id: 1,
            next_num_id: 1,
        }
    }

    /// Add an image and return its relationship ID.
    ///
    /// The image data will be written to the package when save() is called.
    /// Use the returned rel_id when adding an InlineImage to a Run.
    pub fn add_image(&mut self, data: Vec<u8>, content_type: &str) -> String {
        let id = self.next_rel_id;
        self.next_rel_id += 1;

        let rel_id = format!("rId{}", id);
        let ext = extension_from_content_type(content_type);
        let filename = format!("image{}.{}", id, ext);

        self.images.insert(
            rel_id.clone(),
            PendingImage {
                data,
                content_type: content_type.to_string(),
                rel_id: rel_id.clone(),
                filename,
            },
        );

        rel_id
    }

    /// Add a hyperlink and return its relationship ID.
    ///
    /// Use the returned rel_id when creating a Hyperlink in a paragraph.
    pub fn add_hyperlink(&mut self, url: &str) -> String {
        let id = self.next_rel_id;
        self.next_rel_id += 1;

        let rel_id = format!("rId{}", id);

        self.hyperlinks.insert(
            rel_id.clone(),
            PendingHyperlink {
                rel_id: rel_id.clone(),
                url: url.to_string(),
            },
        );

        rel_id
    }

    /// Create a list definition and return its numbering ID.
    ///
    /// Use the returned num_id in NumberingProperties when adding list items.
    ///
    /// # Example
    ///
    /// ```ignore
    /// let num_id = builder.add_list(ListType::Bullet);
    /// let para = builder.body_mut().add_paragraph();
    /// para.set_properties(ParagraphProperties {
    ///     numbering: Some(NumberingProperties { num_id, ilvl: 0 }),
    ///     ..Default::default()
    /// });
    /// para.add_run().set_text("First list item");
    /// ```
    pub fn add_list(&mut self, list_type: ListType) -> u32 {
        let num_id = self.next_num_id;
        self.next_num_id += 1;

        self.numberings.insert(
            num_id,
            PendingNumbering {
                abstract_num_id: num_id, // Use same ID for simplicity
                num_id,
                list_type,
            },
        );

        num_id
    }

    /// Get a mutable reference to the body.
    pub fn body_mut(&mut self) -> &mut Body {
        &mut self.body
    }

    /// Add a paragraph with text.
    pub fn add_paragraph(&mut self, text: &str) -> &mut Self {
        let para = self.body.add_paragraph();
        para.add_run().set_text(text);
        self
    }

    /// Save the document to a file.
    pub fn save<P: AsRef<Path>>(self, path: P) -> Result<()> {
        let file = File::create(path)?;
        let writer = BufWriter::new(file);
        self.write(writer)
    }

    /// Write the document to a writer.
    pub fn write<W: Write + Seek>(self, writer: W) -> Result<()> {
        let mut pkg = PackageWriter::new(writer);

        // Add default content types
        pkg.add_default_content_type("rels", content_type::RELATIONSHIPS);
        pkg.add_default_content_type("xml", content_type::XML);

        // Add content types for images
        pkg.add_default_content_type("png", "image/png");
        pkg.add_default_content_type("jpg", "image/jpeg");
        pkg.add_default_content_type("jpeg", "image/jpeg");
        pkg.add_default_content_type("gif", "image/gif");

        // Write document.xml
        let doc_xml = serialize_document(&self.body);
        pkg.add_part(
            "word/document.xml",
            content_type::WORDPROCESSING_DOCUMENT,
            doc_xml.as_bytes(),
        )?;

        // Write package relationships
        let mut pkg_rels = Relationships::new();
        pkg_rels.add(Relationship::new(
            "rId1",
            rel_type::OFFICE_DOCUMENT,
            "word/document.xml",
        ));
        pkg.add_part(
            "_rels/.rels",
            content_type::RELATIONSHIPS,
            pkg_rels.serialize().as_bytes(),
        )?;

        // Build document relationships
        let mut doc_rels = Relationships::new();

        // Add image relationships and write image files
        for image in self.images.values() {
            doc_rels.add(Relationship::new(
                &image.rel_id,
                rel_type::IMAGE,
                format!("media/{}", image.filename),
            ));

            let image_path = format!("word/media/{}", image.filename);
            pkg.add_part(&image_path, &image.content_type, &image.data)?;
        }

        // Add hyperlink relationships (external)
        for hyperlink in self.hyperlinks.values() {
            doc_rels.add(Relationship::external(
                &hyperlink.rel_id,
                rel_type::HYPERLINK,
                &hyperlink.url,
            ));
        }

        // Write numbering.xml if we have any numbering definitions
        if !self.numberings.is_empty() {
            let num_xml = serialize_numbering(&self.numberings);
            pkg.add_part(
                "word/numbering.xml",
                content_type::WORDPROCESSING_NUMBERING,
                num_xml.as_bytes(),
            )?;

            // Add relationship from document to numbering
            let num_rel_id = format!("rId{}", self.next_rel_id);
            doc_rels.add(Relationship::new(
                &num_rel_id,
                rel_type::NUMBERING,
                "numbering.xml",
            ));
        }

        pkg.add_part(
            "word/_rels/document.xml.rels",
            content_type::RELATIONSHIPS,
            doc_rels.serialize().as_bytes(),
        )?;

        pkg.finish()?;
        Ok(())
    }
}

/// Serialize document body to XML.
pub fn serialize_document(body: &Body) -> String {
    let mut xml = String::new();

    // XML declaration
    xml.push_str(r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>"#);
    xml.push('\n');

    // Document element with namespaces (including DrawingML for images)
    xml.push_str(&format!(
        r#"<w:document xmlns:w="{}" xmlns:r="{}" xmlns:wp="{}" xmlns:a="{}" xmlns:pic="{}">"#,
        NS_W, NS_R, NS_WP, NS_A, NS_PIC
    ));

    // Body
    xml.push_str("<w:body>");
    serialize_body(body, &mut xml);
    xml.push_str("</w:body>");

    xml.push_str("</w:document>");
    xml
}

/// Serialize body contents.
fn serialize_body(body: &Body, xml: &mut String) {
    for block in body.content() {
        match block {
            BlockContent::Paragraph(para) => serialize_paragraph(para, xml),
            BlockContent::Table(table) => serialize_table(table, xml),
        }
    }
    // Section properties (must come after all block content)
    if let Some(sect_pr) = body.section_properties() {
        serialize_section_properties(sect_pr, xml);
    }
    // Write unknown children preserved for round-trip fidelity
    serialize_unknown_children(&body.unknown_children, xml);
}

/// Serialize section properties.
fn serialize_section_properties(props: &SectionProperties, xml: &mut String) {
    xml.push_str("<w:sectPr>");

    // Page size
    if let Some(pg_sz) = &props.page_size {
        xml.push_str("<w:pgSz w:w=\"");
        xml.push_str(&pg_sz.width.to_string());
        xml.push_str("\" w:h=\"");
        xml.push_str(&pg_sz.height.to_string());
        xml.push('"');
        if pg_sz.orientation == PageOrientation::Landscape {
            xml.push_str(" w:orient=\"landscape\"");
        }
        xml.push_str("/>");
    }

    // Page margins
    if let Some(margins) = &props.margins {
        xml.push_str("<w:pgMar w:top=\"");
        xml.push_str(&margins.top.to_string());
        xml.push_str("\" w:bottom=\"");
        xml.push_str(&margins.bottom.to_string());
        xml.push_str("\" w:left=\"");
        xml.push_str(&margins.left.to_string());
        xml.push_str("\" w:right=\"");
        xml.push_str(&margins.right.to_string());
        xml.push('"');
        if let Some(header) = margins.header {
            xml.push_str(" w:header=\"");
            xml.push_str(&header.to_string());
            xml.push('"');
        }
        if let Some(footer) = margins.footer {
            xml.push_str(" w:footer=\"");
            xml.push_str(&footer.to_string());
            xml.push('"');
        }
        if let Some(gutter) = margins.gutter {
            xml.push_str(" w:gutter=\"");
            xml.push_str(&gutter.to_string());
            xml.push('"');
        }
        xml.push_str("/>");
    }

    // Unknown children for round-trip preservation
    serialize_unknown_children(&props.unknown_children, xml);

    xml.push_str("</w:sectPr>");
}

/// Serialize a table.
fn serialize_table(table: &Table, xml: &mut String) {
    xml.push_str("<w:tbl>");
    for row in table.rows() {
        serialize_row(row, xml);
    }
    serialize_unknown_children(&table.unknown_children, xml);
    xml.push_str("</w:tbl>");
}

/// Serialize a table row.
fn serialize_row(row: &Row, xml: &mut String) {
    xml.push_str("<w:tr>");
    for cell in row.cells() {
        serialize_cell(cell, xml);
    }
    serialize_unknown_children(&row.unknown_children, xml);
    xml.push_str("</w:tr>");
}

/// Serialize a table cell.
fn serialize_cell(cell: &Cell, xml: &mut String) {
    xml.push_str("<w:tc>");
    // Cell properties must come first
    if let Some(props) = cell.properties() {
        serialize_cell_properties(props, xml);
    }
    for para in cell.paragraphs() {
        serialize_paragraph(para, xml);
    }
    serialize_unknown_children(&cell.unknown_children, xml);
    xml.push_str("</w:tc>");
}

/// Serialize cell properties.
fn serialize_cell_properties(props: &CellProperties, xml: &mut String) {
    xml.push_str("<w:tcPr>");

    // Cell width
    if let Some(width) = &props.width {
        serialize_cell_width(width, xml);
    }

    // Grid span (horizontal merge)
    if let Some(span) = props.grid_span
        && span > 1
    {
        xml.push_str("<w:gridSpan w:val=\"");
        xml.push_str(&span.to_string());
        xml.push_str("\"/>");
    }

    // Vertical merge
    if let Some(merge) = &props.vertical_merge {
        xml.push_str("<w:vMerge");
        match merge {
            VerticalMerge::Restart => xml.push_str(" w:val=\"restart\""),
            VerticalMerge::Continue => {} // Empty vMerge means continue
        }
        xml.push_str("/>");
    }

    // Borders
    if let Some(borders) = &props.borders {
        serialize_cell_borders(borders, xml);
    }

    // Shading
    if let Some(shading) = &props.shading {
        serialize_cell_shading(shading, xml);
    }

    // Vertical alignment
    if let Some(valign) = &props.vertical_align {
        xml.push_str("<w:vAlign w:val=\"");
        xml.push_str(valign.as_str());
        xml.push_str("\"/>");
    }

    // Unknown children for round-trip preservation
    serialize_unknown_children(&props.unknown_children, xml);

    xml.push_str("</w:tcPr>");
}

/// Serialize cell width.
fn serialize_cell_width(width: &CellWidth, xml: &mut String) {
    xml.push_str("<w:tcW w:w=\"");
    xml.push_str(&width.width.to_string());
    xml.push_str("\" w:type=\"");
    xml.push_str(width.width_type.as_str());
    xml.push_str("\"/>");
}

/// Serialize cell borders.
fn serialize_cell_borders(borders: &CellBorders, xml: &mut String) {
    xml.push_str("<w:tcBorders>");
    if let Some(border) = &borders.top {
        xml.push_str("<w:top");
        serialize_border_attrs(border, xml);
        xml.push_str("/>");
    }
    if let Some(border) = &borders.left {
        xml.push_str("<w:left");
        serialize_border_attrs(border, xml);
        xml.push_str("/>");
    }
    if let Some(border) = &borders.bottom {
        xml.push_str("<w:bottom");
        serialize_border_attrs(border, xml);
        xml.push_str("/>");
    }
    if let Some(border) = &borders.right {
        xml.push_str("<w:right");
        serialize_border_attrs(border, xml);
        xml.push_str("/>");
    }
    if let Some(border) = &borders.inside_h {
        xml.push_str("<w:insideH");
        serialize_border_attrs(border, xml);
        xml.push_str("/>");
    }
    if let Some(border) = &borders.inside_v {
        xml.push_str("<w:insideV");
        serialize_border_attrs(border, xml);
        xml.push_str("/>");
    }
    xml.push_str("</w:tcBorders>");
}

/// Serialize border attributes.
fn serialize_border_attrs(border: &Border, xml: &mut String) {
    xml.push_str(" w:val=\"");
    xml.push_str(border.style.as_str());
    xml.push('"');
    if let Some(sz) = border.size {
        xml.push_str(" w:sz=\"");
        xml.push_str(&sz.to_string());
        xml.push('"');
    }
    if let Some(color) = &border.color {
        xml.push_str(" w:color=\"");
        xml.push_str(color);
        xml.push('"');
    }
    if let Some(space) = border.space {
        xml.push_str(" w:space=\"");
        xml.push_str(&space.to_string());
        xml.push('"');
    }
}

/// Serialize cell shading.
fn serialize_cell_shading(shading: &CellShading, xml: &mut String) {
    xml.push_str("<w:shd");
    if let Some(pattern) = &shading.pattern {
        xml.push_str(" w:val=\"");
        xml.push_str(match pattern {
            crate::document::ShadingPattern::Clear => "clear",
            crate::document::ShadingPattern::Solid => "solid",
            _ => "clear", // Simplified - full mapping would be extensive
        });
        xml.push('"');
    }
    if let Some(fill) = &shading.fill {
        xml.push_str(" w:fill=\"");
        xml.push_str(fill);
        xml.push('"');
    }
    if let Some(color) = &shading.color {
        xml.push_str(" w:color=\"");
        xml.push_str(color);
        xml.push('"');
    }
    xml.push_str("/>");
}

/// Serialize a paragraph.
fn serialize_paragraph(para: &Paragraph, xml: &mut String) {
    xml.push_str("<w:p>");

    // Paragraph properties
    if let Some(props) = para.properties() {
        serialize_paragraph_properties(props, xml);
    }

    // Content (runs and hyperlinks)
    for content in para.content() {
        match content {
            ParagraphContent::Run(run) => serialize_run(run, xml),
            ParagraphContent::Hyperlink(link) => serialize_hyperlink(link, xml),
        }
    }

    serialize_unknown_children(&para.unknown_children, xml);
    xml.push_str("</w:p>");
}

/// Serialize a hyperlink.
fn serialize_hyperlink(link: &Hyperlink, xml: &mut String) {
    xml.push_str("<w:hyperlink");

    if let Some(rel_id) = link.rel_id() {
        xml.push_str(&format!(r#" r:id="{}""#, rel_id));
    }
    if let Some(anchor) = link.anchor() {
        xml.push_str(&format!(r#" w:anchor="{}""#, escape_xml(anchor)));
    }

    // Write unknown attributes preserved for round-trip fidelity
    serialize_unknown_attrs(&link.unknown_attrs, xml);

    xml.push('>');

    for run in link.runs() {
        serialize_run(run, xml);
    }

    serialize_unknown_children(&link.unknown_children, xml);
    xml.push_str("</w:hyperlink>");
}

/// Serialize paragraph properties.
fn serialize_paragraph_properties(props: &ParagraphProperties, xml: &mut String) {
    xml.push_str("<w:pPr>");

    if let Some(ref style) = props.style {
        xml.push_str(&format!(r#"<w:pStyle w:val="{}"/>"#, escape_xml(style)));
    }

    // Numbering properties
    if let Some(ref num_props) = props.numbering {
        serialize_numbering_properties(num_props, xml);
    }

    // Alignment (justification)
    if let Some(alignment) = props.alignment {
        xml.push_str(&format!(r#"<w:jc w:val="{}"/>"#, alignment.as_str()));
    }

    // Spacing
    if props.spacing_before.is_some()
        || props.spacing_after.is_some()
        || props.spacing_line.is_some()
    {
        xml.push_str("<w:spacing");
        if let Some(before) = props.spacing_before {
            xml.push_str(&format!(r#" w:before="{}""#, before));
        }
        if let Some(after) = props.spacing_after {
            xml.push_str(&format!(r#" w:after="{}""#, after));
        }
        if let Some(line) = props.spacing_line {
            xml.push_str(&format!(r#" w:line="{}""#, line));
        }
        xml.push_str("/>");
    }

    // Indentation
    if props.indent_left.is_some()
        || props.indent_right.is_some()
        || props.indent_first_line.is_some()
        || props.indent_hanging.is_some()
    {
        xml.push_str("<w:ind");
        if let Some(left) = props.indent_left {
            xml.push_str(&format!(r#" w:left="{}""#, left));
        }
        if let Some(right) = props.indent_right {
            xml.push_str(&format!(r#" w:right="{}""#, right));
        }
        if let Some(hanging) = props.indent_hanging {
            xml.push_str(&format!(r#" w:hanging="{}""#, hanging));
        } else if let Some(first_line) = props.indent_first_line {
            xml.push_str(&format!(r#" w:firstLine="{}""#, first_line));
        }
        xml.push_str("/>");
    }

    serialize_unknown_children(&props.unknown_children, xml);
    xml.push_str("</w:pPr>");
}

/// Serialize numbering properties (within pPr).
fn serialize_numbering_properties(props: &NumberingProperties, xml: &mut String) {
    xml.push_str("<w:numPr>");
    xml.push_str(&format!(r#"<w:ilvl w:val="{}"/>"#, props.ilvl));
    xml.push_str(&format!(r#"<w:numId w:val="{}"/>"#, props.num_id));
    xml.push_str("</w:numPr>");
}

/// Serialize a run.
fn serialize_run(run: &Run, xml: &mut String) {
    xml.push_str("<w:r");

    // Write unknown attributes preserved for round-trip fidelity
    serialize_unknown_attrs(&run.unknown_attrs, xml);

    xml.push('>');

    // Run properties
    if let Some(props) = run.properties() {
        serialize_run_properties(props, xml);
    }

    // Page break (if any)
    if run.has_page_break() {
        xml.push_str(r#"<w:br w:type="page"/>"#);
    }

    // Drawings (images)
    for drawing in run.drawings() {
        serialize_drawing(drawing, xml);
    }

    // Text content
    let text = run.text();
    if !text.is_empty() {
        // Handle text that needs xml:space="preserve"
        let needs_preserve = text.starts_with(' ')
            || text.ends_with(' ')
            || text.contains('\t')
            || text.contains('\n');

        if needs_preserve {
            xml.push_str(r#"<w:t xml:space="preserve">"#);
        } else {
            xml.push_str("<w:t>");
        }
        xml.push_str(&escape_xml(text));
        xml.push_str("</w:t>");
    }

    serialize_unknown_children(&run.unknown_children, xml);
    xml.push_str("</w:r>");
}

/// Serialize a drawing element.
fn serialize_drawing(drawing: &Drawing, xml: &mut String) {
    xml.push_str("<w:drawing>");
    for (idx, image) in drawing.images().iter().enumerate() {
        serialize_inline_image(image, idx + 1, xml);
    }
    serialize_unknown_children(&drawing.unknown_children, xml);
    xml.push_str("</w:drawing>");
}

/// Serialize an inline image.
///
/// Generates the DrawingML structure required for an inline image.
fn serialize_inline_image(image: &InlineImage, doc_id: usize, xml: &mut String) {
    // Default dimensions: 1 inch x 1 inch (914400 EMUs)
    let cx = image.width_emu().unwrap_or(914400);
    let cy = image.height_emu().unwrap_or(914400);
    let rel_id = image.rel_id();
    let descr = image.description().unwrap_or("Image");

    // Inline element with extent
    xml.push_str(r#"<wp:inline distT="0" distB="0" distL="0" distR="0">"#);
    xml.push_str(&format!(r#"<wp:extent cx="{}" cy="{}"/>"#, cx, cy));

    // Document properties
    xml.push_str(&format!(
        r#"<wp:docPr id="{}" name="Picture {}" descr="{}"/>"#,
        doc_id,
        doc_id,
        escape_xml(descr)
    ));

    // Graphic frame lock
    xml.push_str(
        r#"<wp:cNvGraphicFramePr><a:graphicFrameLocks noChangeAspect="1"/></wp:cNvGraphicFramePr>"#,
    );

    // Graphic container
    xml.push_str(r#"<a:graphic>"#);
    xml.push_str(
        r#"<a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">"#,
    );

    // Picture element
    xml.push_str(r#"<pic:pic>"#);

    // Non-visual properties
    xml.push_str(&format!(
        r#"<pic:nvPicPr><pic:cNvPr id="{}" name="Picture {}"/><pic:cNvPicPr/></pic:nvPicPr>"#,
        doc_id, doc_id
    ));

    // Blip fill (references the image relationship)
    xml.push_str(r#"<pic:blipFill>"#);
    xml.push_str(&format!(r#"<a:blip r:embed="{}"/>"#, rel_id));
    xml.push_str(r#"<a:stretch><a:fillRect/></a:stretch>"#);
    xml.push_str(r#"</pic:blipFill>"#);

    // Shape properties
    xml.push_str(r#"<pic:spPr>"#);
    xml.push_str(r#"<a:xfrm>"#);
    xml.push_str(r#"<a:off x="0" y="0"/>"#);
    xml.push_str(&format!(r#"<a:ext cx="{}" cy="{}"/>"#, cx, cy));
    xml.push_str(r#"</a:xfrm>"#);
    xml.push_str(r#"<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>"#);
    xml.push_str(r#"</pic:spPr>"#);

    xml.push_str(r#"</pic:pic>"#);
    xml.push_str(r#"</a:graphicData>"#);
    xml.push_str(r#"</a:graphic>"#);
    xml.push_str(r#"</wp:inline>"#);
}

/// Serialize run properties.
fn serialize_run_properties(props: &RunProperties, xml: &mut String) {
    // Only output if there are properties to write
    let has_props = props.bold
        || props.italic
        || props.underline.is_some()
        || props.strike
        || props.double_strike
        || props.size.is_some()
        || props.font.is_some()
        || props.style.is_some()
        || props.color.is_some()
        || props.highlight.is_some()
        || props.vertical_align.is_some()
        || props.all_caps
        || props.small_caps
        || !props.unknown_children.is_empty();

    if !has_props {
        return;
    }

    xml.push_str("<w:rPr>");

    if let Some(ref style) = props.style {
        xml.push_str(&format!(r#"<w:rStyle w:val="{}"/>"#, escape_xml(style)));
    }

    if let Some(ref font) = props.font {
        xml.push_str(&format!(r#"<w:rFonts w:ascii="{}"/>"#, escape_xml(font)));
    }

    if props.bold {
        xml.push_str("<w:b/>");
    }

    if props.italic {
        xml.push_str("<w:i/>");
    }

    if let Some(underline) = props.underline {
        xml.push_str(&format!(r#"<w:u w:val="{}"/>"#, underline.as_str()));
    }

    if props.strike {
        xml.push_str("<w:strike/>");
    }

    if props.double_strike {
        xml.push_str("<w:dstrike/>");
    }

    if props.all_caps {
        xml.push_str("<w:caps/>");
    }

    if props.small_caps {
        xml.push_str("<w:smallCaps/>");
    }

    if let Some(highlight) = props.highlight {
        xml.push_str(&format!(r#"<w:highlight w:val="{}"/>"#, highlight.as_str()));
    }

    if let Some(vertical_align) = props.vertical_align {
        xml.push_str(&format!(
            r#"<w:vertAlign w:val="{}"/>"#,
            vertical_align.as_str()
        ));
    }

    if let Some(size) = props.size {
        xml.push_str(&format!(r#"<w:sz w:val="{}"/>"#, size));
    }

    if let Some(ref color) = props.color {
        xml.push_str(&format!(r#"<w:color w:val="{}"/>"#, escape_xml(color)));
    }

    serialize_unknown_children(&props.unknown_children, xml);
    xml.push_str("</w:rPr>");
}

/// Escape special XML characters.
fn escape_xml(s: &str) -> String {
    let mut result = String::with_capacity(s.len());
    for c in s.chars() {
        match c {
            '&' => result.push_str("&amp;"),
            '<' => result.push_str("&lt;"),
            '>' => result.push_str("&gt;"),
            '"' => result.push_str("&quot;"),
            '\'' => result.push_str("&apos;"),
            _ => result.push(c),
        }
    }
    result
}

/// Serialize a RawXmlNode (preserved unknown element) to XML string.
fn serialize_raw_xml_node(node: &RawXmlNode, xml: &mut String) {
    match node {
        RawXmlNode::Element(elem) => serialize_raw_xml_element(elem, xml),
        RawXmlNode::Text(text) => xml.push_str(&escape_xml(text)),
        RawXmlNode::CData(text) => {
            xml.push_str("<![CDATA[");
            xml.push_str(text);
            xml.push_str("]]>");
        }
        RawXmlNode::Comment(text) => {
            xml.push_str("<!--");
            xml.push_str(text);
            xml.push_str("-->");
        }
    }
}

/// Serialize a RawXmlElement (preserved unknown element) to XML string.
fn serialize_raw_xml_element(elem: &crate::raw_xml::RawXmlElement, xml: &mut String) {
    xml.push('<');
    xml.push_str(&elem.name);

    for (key, value) in &elem.attributes {
        xml.push(' ');
        xml.push_str(key);
        xml.push_str("=\"");
        xml.push_str(&escape_xml(value));
        xml.push('"');
    }

    if elem.self_closing && elem.children.is_empty() {
        xml.push_str("/>");
    } else {
        xml.push('>');
        for child in &elem.children {
            serialize_raw_xml_node(child, xml);
        }
        xml.push_str("</");
        xml.push_str(&elem.name);
        xml.push('>');
    }
}

/// Serialize unknown children preserved for round-trip fidelity.
/// Children are sorted by position to maintain original order.
fn serialize_unknown_children(children: &[PositionedNode], xml: &mut String) {
    // Sort by position to interleave correctly with known elements
    let mut sorted: Vec<_> = children.iter().collect();
    sorted.sort_by_key(|pn| pn.position);
    for pn in sorted {
        serialize_raw_xml_node(&pn.node, xml);
    }
}

/// Serialize unknown attributes preserved for round-trip fidelity.
/// Attributes are sorted by position to maintain original order.
fn serialize_unknown_attrs(attrs: &[PositionedAttr], xml: &mut String) {
    // Sort by position to preserve original attribute order
    let mut sorted: Vec<_> = attrs.iter().collect();
    sorted.sort_by_key(|pa| pa.position);
    for pa in sorted {
        xml.push(' ');
        xml.push_str(&pa.name);
        xml.push_str("=\"");
        xml.push_str(&escape_xml(&pa.value));
        xml.push('"');
    }
}

/// Get file extension from MIME content type.
fn extension_from_content_type(content_type: &str) -> &'static str {
    match content_type {
        "image/png" => "png",
        "image/jpeg" => "jpg",
        "image/gif" => "gif",
        "image/bmp" => "bmp",
        "image/tiff" => "tiff",
        "image/webp" => "webp",
        "image/svg+xml" => "svg",
        "image/x-emf" | "image/emf" => "emf",
        "image/x-wmf" | "image/wmf" => "wmf",
        _ => "bin",
    }
}

/// Serialize numbering.xml content.
fn serialize_numbering(numberings: &HashMap<u32, PendingNumbering>) -> String {
    let mut xml = String::new();

    // XML declaration
    xml.push_str(r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>"#);
    xml.push('\n');

    // Numbering element with namespace
    xml.push_str(&format!(r#"<w:numbering xmlns:w="{}">"#, NS_W));

    // Sort numberings by num_id for deterministic output
    let mut sorted: Vec<_> = numberings.values().collect();
    sorted.sort_by_key(|n| n.num_id);

    // Write abstract numbering definitions
    for num in &sorted {
        serialize_abstract_num(num, &mut xml);
    }

    // Write concrete numbering instances
    for num in &sorted {
        xml.push_str(&format!(
            r#"<w:num w:numId="{}"><w:abstractNumId w:val="{}"/></w:num>"#,
            num.num_id, num.abstract_num_id
        ));
    }

    xml.push_str("</w:numbering>");
    xml
}

/// Serialize an abstract numbering definition.
fn serialize_abstract_num(num: &PendingNumbering, xml: &mut String) {
    xml.push_str(&format!(
        r#"<w:abstractNum w:abstractNumId="{}">"#,
        num.abstract_num_id
    ));

    // Level 0 definition (we only support single-level lists in v0.1)
    xml.push_str(r#"<w:lvl w:ilvl="0">"#);

    // Start value
    xml.push_str(r#"<w:start w:val="1"/>"#);

    // Number format and text based on list type
    let (num_fmt, lvl_text) = match num.list_type {
        ListType::Bullet => ("bullet", "\u{2022}"), // Bullet character
        ListType::Decimal => ("decimal", "%1."),
        ListType::LowerLetter => ("lowerLetter", "%1."),
        ListType::UpperLetter => ("upperLetter", "%1."),
        ListType::LowerRoman => ("lowerRoman", "%1."),
        ListType::UpperRoman => ("upperRoman", "%1."),
    };

    xml.push_str(&format!(r#"<w:numFmt w:val="{}"/>"#, num_fmt));
    xml.push_str(&format!(r#"<w:lvlText w:val="{}"/>"#, lvl_text));
    xml.push_str(r#"<w:lvlJc w:val="left"/>"#);

    // Paragraph properties (indentation)
    xml.push_str("<w:pPr>");
    xml.push_str(r#"<w:ind w:left="720" w:hanging="360"/>"#);
    xml.push_str("</w:pPr>");

    // Run properties for bullet lists (use Symbol font)
    if num.list_type == ListType::Bullet {
        xml.push_str("<w:rPr>");
        xml.push_str(r#"<w:rFonts w:ascii="Symbol" w:hAnsi="Symbol" w:hint="default"/>"#);
        xml.push_str("</w:rPr>");
    }

    xml.push_str("</w:lvl>");
    xml.push_str("</w:abstractNum>");
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_serialize_simple_document() {
        let mut body = Body::new();
        body.add_paragraph().add_run().set_text("Hello, World!");

        let xml = serialize_document(&body);

        assert!(xml.contains("<w:document"));
        assert!(xml.contains("<w:body>"));
        assert!(xml.contains("<w:p>"));
        assert!(xml.contains("<w:r>"));
        assert!(xml.contains("<w:t>Hello, World!</w:t>"));
    }

    #[test]
    fn test_serialize_with_formatting() {
        let mut body = Body::new();
        let run = body.add_paragraph().add_run();
        run.set_text("Bold text");
        run.set_properties(RunProperties {
            bold: true,
            italic: true,
            ..Default::default()
        });

        let xml = serialize_document(&body);

        assert!(xml.contains("<w:b/>"));
        assert!(xml.contains("<w:i/>"));
    }

    #[test]
    fn test_escape_xml_entities() {
        let mut body = Body::new();
        body.add_paragraph()
            .add_run()
            .set_text("Tom & Jerry <friends>");

        let xml = serialize_document(&body);

        assert!(xml.contains("Tom &amp; Jerry &lt;friends&gt;"));
    }

    #[test]
    fn test_preserve_whitespace() {
        let mut body = Body::new();
        body.add_paragraph().add_run().set_text("  leading spaces");

        let xml = serialize_document(&body);

        assert!(xml.contains(r#"xml:space="preserve""#));
    }

    #[test]
    fn test_document_builder() {
        let mut builder = DocumentBuilder::new();
        builder.add_paragraph("First paragraph");
        builder.add_paragraph("Second paragraph");

        let body = &builder.body;
        assert_eq!(body.paragraphs().len(), 2);
    }

    #[test]
    fn test_roundtrip_create_and_read() {
        use crate::Document;
        use std::io::Cursor;

        // Create a document
        let mut builder = DocumentBuilder::new();
        builder.add_paragraph("Test content");

        // Write to memory
        let mut buffer = Cursor::new(Vec::new());
        builder.write(&mut buffer).unwrap();

        // Read it back
        buffer.set_position(0);
        let doc = Document::from_reader(buffer).unwrap();

        assert_eq!(doc.body().paragraphs().len(), 1);
        assert_eq!(doc.text(), "Test content");
    }

    #[test]
    fn test_serialize_table() {
        let mut body = Body::new();
        let table = body.add_table();
        let row = table.add_row();
        row.add_cell().add_paragraph().add_run().set_text("A1");
        row.add_cell().add_paragraph().add_run().set_text("B1");

        let xml = serialize_document(&body);

        assert!(xml.contains("<w:tbl>"));
        assert!(xml.contains("<w:tr>"));
        assert!(xml.contains("<w:tc>"));
        assert!(xml.contains("<w:t>A1</w:t>"));
        assert!(xml.contains("<w:t>B1</w:t>"));
    }

    #[test]
    fn test_roundtrip_table() {
        use crate::Document;
        use std::io::Cursor;

        // Create a document with a table
        let mut builder = DocumentBuilder::new();
        builder.add_paragraph("Before table");
        let table = builder.body_mut().add_table();
        let row = table.add_row();
        row.add_cell().add_paragraph().add_run().set_text("Cell 1");
        row.add_cell().add_paragraph().add_run().set_text("Cell 2");
        builder
            .body_mut()
            .add_paragraph()
            .add_run()
            .set_text("After table");

        // Write to memory
        let mut buffer = Cursor::new(Vec::new());
        builder.write(&mut buffer).unwrap();

        // Read it back
        buffer.set_position(0);
        let doc = Document::from_reader(buffer).unwrap();

        // Verify structure
        assert_eq!(doc.body().content().len(), 3); // para, table, para
        assert_eq!(doc.body().tables().count(), 1);

        let table = doc.body().tables().next().unwrap();
        assert_eq!(table.row_count(), 1);
        assert_eq!(table.column_count(), 2);
        assert_eq!(table.rows()[0].cells()[0].text(), "Cell 1");
        assert_eq!(table.rows()[0].cells()[1].text(), "Cell 2");
    }

    #[test]
    fn test_serialize_inline_image() {
        use crate::document::Drawing;

        let mut body = Body::new();
        let run = body.add_paragraph().add_run();

        // Add a drawing with an image
        let mut drawing = Drawing::new();
        drawing
            .add_image("rId1")
            .set_width_inches(2.0)
            .set_height_inches(1.5)
            .set_description("Test image");
        run.drawings_mut().push(drawing);

        let xml = serialize_document(&body);

        // Check DrawingML structure
        assert!(xml.contains("<w:drawing>"));
        assert!(xml.contains("<wp:inline"));
        assert!(xml.contains(r#"r:embed="rId1""#));
        assert!(xml.contains("wp:extent"));
        assert!(xml.contains("pic:pic"));
        assert!(xml.contains(r#"descr="Test image""#));
    }

    #[test]
    fn test_document_builder_add_image() {
        let mut builder = DocumentBuilder::new();

        // Add an image via the builder
        let rel_id = builder.add_image(vec![0x89, 0x50, 0x4E, 0x47], "image/png");
        assert_eq!(rel_id, "rId1");

        // Add another image
        let rel_id2 = builder.add_image(vec![0xFF, 0xD8, 0xFF], "image/jpeg");
        assert_eq!(rel_id2, "rId2");

        // Verify the images are tracked
        assert_eq!(builder.images.len(), 2);
    }

    #[test]
    fn test_extension_from_content_type() {
        assert_eq!(extension_from_content_type("image/png"), "png");
        assert_eq!(extension_from_content_type("image/jpeg"), "jpg");
        assert_eq!(extension_from_content_type("image/gif"), "gif");
        assert_eq!(extension_from_content_type("unknown/type"), "bin");
    }
}
