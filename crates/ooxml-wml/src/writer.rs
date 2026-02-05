//! Document writing and serialization.
//!
//! This module provides functionality for creating new Word documents
//! using generated types and ToXml serializers.

use crate::error::Result;
use crate::generated_serializers::ToXml;
use crate::types;
use ooxml_opc::{PackageWriter, Relationship, Relationships, content_type, rel_type};
use ooxml_xml::{PositionedNode, RawXmlElement, RawXmlNode};
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

/// Standard namespace declarations used on root elements.
const NS_DECLS: &[(&str, &str)] = &[
    ("xmlns:w", NS_W),
    ("xmlns:r", NS_R),
    ("xmlns:wp", NS_WP),
    ("xmlns:a", NS_A),
    ("xmlns:pic", NS_PIC),
];

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
    /// Parse from the `w:type` attribute value.
    pub fn parse(s: &str) -> Self {
        match s {
            "first" => Self::First,
            "even" => Self::Even,
            _ => Self::Default,
        }
    }

    /// Convert to the `w:type` attribute value.
    pub fn as_str(&self) -> &'static str {
        match self {
            Self::Default => "default",
            Self::First => "first",
            Self::Even => "even",
        }
    }
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

/// A pending header to be written to the package.
#[derive(Clone)]
pub struct PendingHeader {
    /// Header content.
    pub body: types::HeaderFooter,
    /// Assigned relationship ID.
    pub rel_id: String,
    /// Header type (default, first, even).
    pub header_type: HeaderFooterType,
    /// Generated filename (e.g., "header1.xml").
    pub filename: String,
}

/// A pending footer to be written to the package.
#[derive(Clone)]
pub struct PendingFooter {
    /// Footer content.
    pub body: types::HeaderFooter,
    /// Assigned relationship ID.
    pub rel_id: String,
    /// Footer type (default, first, even).
    pub footer_type: HeaderFooterType,
    /// Generated filename (e.g., "footer1.xml").
    pub filename: String,
}

/// A pending footnote to be written to the package.
#[derive(Clone)]
pub struct PendingFootnote {
    /// Footnote ID (referenced by FootnoteReference).
    pub id: i32,
    /// Footnote content.
    pub body: types::FootnoteEndnote,
}

/// A pending endnote to be written to the package.
#[derive(Clone)]
pub struct PendingEndnote {
    /// Endnote ID (referenced by EndnoteReference).
    pub id: i32,
    /// Endnote content.
    pub body: types::FootnoteEndnote,
}

/// A pending comment to be written to the package.
#[derive(Clone)]
pub struct PendingComment {
    /// Comment ID (referenced by CommentReference and comment ranges).
    pub id: i32,
    /// Comment author.
    pub author: Option<String>,
    /// Comment date (ISO 8601 format).
    pub date: Option<String>,
    /// Comment initials.
    pub initials: Option<String>,
    /// Comment content.
    pub body: types::Comment,
}

/// Builder for header content.
///
/// Provides a fluent API for building header content.
pub struct HeaderBuilder<'a> {
    builder: &'a mut DocumentBuilder,
    rel_id: String,
}

impl<'a> HeaderBuilder<'a> {
    /// Get a mutable reference to the header body.
    pub fn body_mut(&mut self) -> &mut types::HeaderFooter {
        &mut self
            .builder
            .headers
            .get_mut(&self.rel_id)
            .expect("header should exist")
            .body
    }

    /// Add a paragraph with text to the header.
    pub fn add_paragraph(&mut self, text: &str) -> &mut Self {
        self.body_mut().add_paragraph().add_run().set_text(text);
        self
    }

    /// Get the relationship ID for this header.
    pub fn rel_id(&self) -> &str {
        &self.rel_id
    }
}

/// Builder for footer content.
///
/// Provides a fluent API for building footer content.
pub struct FooterBuilder<'a> {
    builder: &'a mut DocumentBuilder,
    rel_id: String,
}

impl<'a> FooterBuilder<'a> {
    /// Get a mutable reference to the footer body.
    pub fn body_mut(&mut self) -> &mut types::HeaderFooter {
        &mut self
            .builder
            .footers
            .get_mut(&self.rel_id)
            .expect("footer should exist")
            .body
    }

    /// Add a paragraph with text to the footer.
    pub fn add_paragraph(&mut self, text: &str) -> &mut Self {
        self.body_mut().add_paragraph().add_run().set_text(text);
        self
    }

    /// Get the relationship ID for this footer.
    pub fn rel_id(&self) -> &str {
        &self.rel_id
    }
}

/// Builder for footnote content.
pub struct FootnoteBuilder<'a> {
    builder: &'a mut DocumentBuilder,
    id: i32,
}

impl<'a> FootnoteBuilder<'a> {
    /// Get a mutable reference to the footnote body.
    pub fn body_mut(&mut self) -> &mut types::FootnoteEndnote {
        &mut self
            .builder
            .footnotes
            .get_mut(&self.id)
            .expect("footnote should exist")
            .body
    }

    /// Add a paragraph with text to the footnote.
    pub fn add_paragraph(&mut self, text: &str) -> &mut Self {
        self.body_mut().add_paragraph().add_run().set_text(text);
        self
    }

    /// Get the footnote ID for use in FootnoteReference.
    ///
    /// The returned ID is always positive (user-created footnotes start at 1).
    pub fn id(&self) -> u32 {
        self.id as u32
    }
}

/// Builder for endnote content.
pub struct EndnoteBuilder<'a> {
    builder: &'a mut DocumentBuilder,
    id: i32,
}

impl<'a> EndnoteBuilder<'a> {
    /// Get a mutable reference to the endnote body.
    pub fn body_mut(&mut self) -> &mut types::FootnoteEndnote {
        &mut self
            .builder
            .endnotes
            .get_mut(&self.id)
            .expect("endnote should exist")
            .body
    }

    /// Add a paragraph with text to the endnote.
    pub fn add_paragraph(&mut self, text: &str) -> &mut Self {
        self.body_mut().add_paragraph().add_run().set_text(text);
        self
    }

    /// Get the endnote ID for use in EndnoteReference.
    ///
    /// The returned ID is always positive (user-created endnotes start at 1).
    pub fn id(&self) -> u32 {
        self.id as u32
    }
}

/// Builder for comment content.
pub struct CommentBuilder<'a> {
    builder: &'a mut DocumentBuilder,
    id: i32,
}

impl<'a> CommentBuilder<'a> {
    /// Get a mutable reference to the comment body.
    pub fn body_mut(&mut self) -> &mut types::Comment {
        &mut self
            .builder
            .comments
            .get_mut(&self.id)
            .expect("comment should exist")
            .body
    }

    /// Add a paragraph with text to the comment.
    pub fn add_paragraph(&mut self, text: &str) -> &mut Self {
        self.body_mut().add_paragraph().add_run().set_text(text);
        self
    }

    /// Set the comment author.
    pub fn set_author(&mut self, author: &str) -> &mut Self {
        self.builder
            .comments
            .get_mut(&self.id)
            .expect("comment should exist")
            .author = Some(author.to_string());
        self
    }

    /// Set the comment date (ISO 8601 format, e.g., "2024-01-15T10:30:00Z").
    pub fn set_date(&mut self, date: &str) -> &mut Self {
        self.builder
            .comments
            .get_mut(&self.id)
            .expect("comment should exist")
            .date = Some(date.to_string());
        self
    }

    /// Set the comment initials.
    pub fn set_initials(&mut self, initials: &str) -> &mut Self {
        self.builder
            .comments
            .get_mut(&self.id)
            .expect("comment should exist")
            .initials = Some(initials.to_string());
        self
    }

    /// Get the comment ID for use in CommentReference and comment ranges.
    ///
    /// The returned ID is always positive (user-created comments start at 0).
    pub fn id(&self) -> u32 {
        self.id as u32
    }
}

// =============================================================================
// Drawing types (writer-only, produce DrawingML XML for images)
// =============================================================================

/// A drawing container for images.
///
/// This is a writer-side helper that produces the DrawingML XML for inline
/// and anchored images. Use `Drawing::build()` to convert to a generated
/// `types::CTDrawing` that can be added to a run.
#[derive(Debug, Clone, Default)]
pub struct Drawing {
    /// Inline images in this drawing.
    images: Vec<InlineImage>,
    /// Anchored (floating) images in this drawing.
    anchored_images: Vec<AnchoredImage>,
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

    /// Add an anchored (floating) image to this drawing.
    pub fn add_anchored_image(&mut self, rel_id: impl Into<String>) -> &mut AnchoredImage {
        self.anchored_images.push(AnchoredImage::new(rel_id));
        self.anchored_images.last_mut().unwrap()
    }

    /// Convert this drawing to a generated `CTDrawing` type.
    ///
    /// The `doc_id` counter is incremented for each image to produce unique IDs.
    pub fn build(self, doc_id: &mut usize) -> types::CTDrawing {
        let mut children = Vec::new();
        let mut child_idx = 0usize;

        for image in &self.images {
            let elem = build_inline_image_element(image, *doc_id);
            children.push(PositionedNode::new(child_idx, RawXmlNode::Element(elem)));
            child_idx += 1;
            *doc_id += 1;
        }

        for image in &self.anchored_images {
            let elem = build_anchored_image_element(image, *doc_id);
            children.push(PositionedNode::new(child_idx, RawXmlNode::Element(elem)));
            child_idx += 1;
            *doc_id += 1;
        }

        types::CTDrawing {
            #[cfg(feature = "extra-children")]
            extra_children: children,
        }
    }
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
}

impl InlineImage {
    /// Create a new inline image with the given relationship ID.
    pub fn new(rel_id: impl Into<String>) -> Self {
        Self {
            rel_id: rel_id.into(),
            width_emu: None,
            height_emu: None,
            description: None,
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

/// An anchored (floating) image in a drawing.
///
/// Represents an image positioned relative to a reference point with text
/// wrapping options. Unlike inline images, anchored images can float and wrap.
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

// =============================================================================
// DocumentBuilder
// =============================================================================

/// Builder for creating new Word documents.
pub struct DocumentBuilder {
    document: types::Document,
    /// Pending images to write, keyed by rel_id.
    images: HashMap<String, PendingImage>,
    /// Pending hyperlinks, keyed by rel_id.
    hyperlinks: HashMap<String, PendingHyperlink>,
    /// Numbering definitions, keyed by num_id.
    numberings: HashMap<u32, PendingNumbering>,
    /// Pending headers, keyed by rel_id.
    headers: HashMap<String, PendingHeader>,
    /// Pending footers, keyed by rel_id.
    footers: HashMap<String, PendingFooter>,
    /// Pending footnotes, keyed by ID.
    footnotes: HashMap<i32, PendingFootnote>,
    /// Pending endnotes, keyed by ID.
    endnotes: HashMap<i32, PendingEndnote>,
    /// Pending comments, keyed by ID.
    comments: HashMap<i32, PendingComment>,
    /// Counter for generating unique IDs.
    next_rel_id: u32,
    /// Counter for generating unique numbering IDs.
    next_num_id: u32,
    /// Counter for generating unique header IDs.
    next_header_id: u32,
    /// Counter for generating unique footer IDs.
    next_footer_id: u32,
    /// Counter for generating unique footnote IDs.
    /// Starts at 1 because 0 is reserved for the separator footnote.
    next_footnote_id: i32,
    /// Counter for generating unique endnote IDs.
    /// Starts at 1 because 0 is reserved for the separator endnote.
    next_endnote_id: i32,
    /// Counter for generating unique comment IDs.
    next_comment_id: i32,
    /// Counter for generating unique drawing/image IDs.
    next_drawing_id: usize,
}

impl Default for DocumentBuilder {
    fn default() -> Self {
        Self::new()
    }
}

impl DocumentBuilder {
    /// Create a new document builder.
    pub fn new() -> Self {
        let document = types::Document {
            #[cfg(feature = "wml-styling")]
            background: None,
            body: Some(Box::new(types::Body::default())),
            conformance: None,
            #[cfg(feature = "extra-attrs")]
            extra_attrs: std::collections::HashMap::new(),
            #[cfg(feature = "extra-children")]
            extra_children: Vec::new(),
        };

        Self {
            document,
            images: HashMap::new(),
            hyperlinks: HashMap::new(),
            numberings: HashMap::new(),
            headers: HashMap::new(),
            footers: HashMap::new(),
            footnotes: HashMap::new(),
            endnotes: HashMap::new(),
            comments: HashMap::new(),
            next_rel_id: 1,
            next_num_id: 1,
            next_header_id: 1,
            next_footer_id: 1,
            next_footnote_id: 1,
            next_endnote_id: 1,
            next_comment_id: 0,
            next_drawing_id: 1,
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

    /// Add a header and return a builder for its content.
    ///
    /// The header will be automatically linked to the document's section properties.
    pub fn add_header(&mut self, header_type: HeaderFooterType) -> HeaderBuilder<'_> {
        let id = self.next_rel_id;
        self.next_rel_id += 1;
        let header_num = self.next_header_id;
        self.next_header_id += 1;

        let rel_id = format!("rId{}", id);
        let filename = format!("header{}.xml", header_num);

        self.headers.insert(
            rel_id.clone(),
            PendingHeader {
                body: types::HeaderFooter::default(),
                rel_id: rel_id.clone(),
                header_type,
                filename,
            },
        );

        HeaderBuilder {
            builder: self,
            rel_id,
        }
    }

    /// Add a footer and return a builder for its content.
    ///
    /// The footer will be automatically linked to the document's section properties.
    pub fn add_footer(&mut self, footer_type: HeaderFooterType) -> FooterBuilder<'_> {
        let id = self.next_rel_id;
        self.next_rel_id += 1;
        let footer_num = self.next_footer_id;
        self.next_footer_id += 1;

        let rel_id = format!("rId{}", id);
        let filename = format!("footer{}.xml", footer_num);

        self.footers.insert(
            rel_id.clone(),
            PendingFooter {
                body: types::HeaderFooter::default(),
                rel_id: rel_id.clone(),
                footer_type,
                filename,
            },
        );

        FooterBuilder {
            builder: self,
            rel_id,
        }
    }

    /// Add a footnote and return a builder for its content.
    ///
    /// Use the returned `id` when adding a FootnoteReference to a Run.
    pub fn add_footnote(&mut self) -> FootnoteBuilder<'_> {
        let id = self.next_footnote_id;
        self.next_footnote_id += 1;

        self.footnotes.insert(
            id,
            PendingFootnote {
                id,
                body: types::FootnoteEndnote {
                    r#type: None,
                    id: id as i64,
                    block_content: Vec::new(),
                    #[cfg(feature = "extra-attrs")]
                    extra_attrs: std::collections::HashMap::new(),
                    #[cfg(feature = "extra-children")]
                    extra_children: Vec::new(),
                },
            },
        );

        FootnoteBuilder { builder: self, id }
    }

    /// Add an endnote and return a builder for its content.
    ///
    /// Use the returned `id` when adding an EndnoteReference to a Run.
    pub fn add_endnote(&mut self) -> EndnoteBuilder<'_> {
        let id = self.next_endnote_id;
        self.next_endnote_id += 1;

        self.endnotes.insert(
            id,
            PendingEndnote {
                id,
                body: types::FootnoteEndnote {
                    r#type: None,
                    id: id as i64,
                    block_content: Vec::new(),
                    #[cfg(feature = "extra-attrs")]
                    extra_attrs: std::collections::HashMap::new(),
                    #[cfg(feature = "extra-children")]
                    extra_children: Vec::new(),
                },
            },
        );

        EndnoteBuilder { builder: self, id }
    }

    /// Add a comment and return a builder for its content.
    ///
    /// Use the returned `id` when adding comment ranges and references to the document.
    pub fn add_comment(&mut self) -> CommentBuilder<'_> {
        let id = self.next_comment_id;
        self.next_comment_id += 1;

        self.comments.insert(
            id,
            PendingComment {
                id,
                author: None,
                date: None,
                initials: None,
                body: types::Comment {
                    id: 0,                 // set in build_comments
                    author: String::new(), // set in build_comments
                    date: None,
                    block_content: Vec::new(),
                    initials: None,
                    #[cfg(feature = "extra-attrs")]
                    extra_attrs: Default::default(),
                    #[cfg(feature = "extra-children")]
                    extra_children: Vec::new(),
                },
            },
        );

        CommentBuilder { builder: self, id }
    }

    /// Get a mutable reference to the document body.
    pub fn body_mut(&mut self) -> &mut types::Body {
        self.document
            .body
            .as_deref_mut()
            .expect("document body should exist")
    }

    /// Add a paragraph with text.
    pub fn add_paragraph(&mut self, text: &str) -> &mut Self {
        let para = self.body_mut().add_paragraph();
        para.add_run().set_text(text);
        self
    }

    /// Convert a Drawing helper to a CTDrawing using the builder's ID counter.
    ///
    /// Use this to create a `types::CTDrawing` from a `Drawing`, then add it
    /// to a run via `run.add_drawing(ct_drawing)`.
    pub fn build_drawing(&mut self, drawing: Drawing) -> types::CTDrawing {
        drawing.build(&mut self.next_drawing_id)
    }

    /// Save the document to a file.
    pub fn save<P: AsRef<Path>>(self, path: P) -> Result<()> {
        let file = File::create(path)?;
        let writer = BufWriter::new(file);
        self.write(writer)
    }

    /// Write the document to a writer.
    pub fn write<W: Write + Seek>(mut self, writer: W) -> Result<()> {
        let mut pkg = PackageWriter::new(writer);

        // Add default content types
        pkg.add_default_content_type("rels", content_type::RELATIONSHIPS);
        pkg.add_default_content_type("xml", content_type::XML);

        // Add content types for images
        pkg.add_default_content_type("png", "image/png");
        pkg.add_default_content_type("jpg", "image/jpeg");
        pkg.add_default_content_type("jpeg", "image/jpeg");
        pkg.add_default_content_type("gif", "image/gif");

        // Build document relationships
        let mut doc_rels = Relationships::new();

        // Add header/footer references to section properties
        if !self.headers.is_empty() || !self.footers.is_empty() {
            // Ensure body has section properties
            #[cfg(feature = "wml-layout")]
            {
                let body = self.document.body.as_deref_mut().expect("document body");
                if body.sect_pr.is_none() {
                    body.sect_pr = Some(Box::new(types::SectionProperties::default()));
                }
                let sect_pr = body.sect_pr.as_deref_mut().unwrap();

                // Add header references
                for header in self.headers.values() {
                    let hdr_ref = types::HeaderFooterReference {
                        id: header.rel_id.clone(),
                        r#type: match header.header_type {
                            HeaderFooterType::Default => types::STHdrFtr::Default,
                            HeaderFooterType::First => types::STHdrFtr::First,
                            HeaderFooterType::Even => types::STHdrFtr::Even,
                        },
                        #[cfg(feature = "extra-attrs")]
                        extra_attrs: std::collections::HashMap::new(),
                    };
                    sect_pr
                        .header_footer_refs
                        .push(types::HeaderFooterRef::HeaderReference(Box::new(hdr_ref)));
                }

                // Add footer references
                for footer in self.footers.values() {
                    let ftr_ref = types::HeaderFooterReference {
                        id: footer.rel_id.clone(),
                        r#type: match footer.footer_type {
                            HeaderFooterType::Default => types::STHdrFtr::Default,
                            HeaderFooterType::First => types::STHdrFtr::First,
                            HeaderFooterType::Even => types::STHdrFtr::Even,
                        },
                        #[cfg(feature = "extra-attrs")]
                        extra_attrs: std::collections::HashMap::new(),
                    };
                    sect_pr
                        .header_footer_refs
                        .push(types::HeaderFooterRef::FooterReference(Box::new(ftr_ref)));
                }
            }
        }

        // Set namespace declarations on the document's extra_attrs
        #[cfg(feature = "extra-attrs")]
        {
            for &(key, value) in NS_DECLS {
                self.document
                    .extra_attrs
                    .insert(key.to_string(), value.to_string());
            }
        }

        // Write document.xml
        let doc_xml = serialize_to_xml_bytes(&self.document, "w:document")?;
        pkg.add_part(
            "word/document.xml",
            content_type::WORDPROCESSING_DOCUMENT,
            &doc_xml,
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

        // Write headers and add relationships
        for header in self.headers.values() {
            let header_xml = serialize_with_namespaces(&header.body, "w:hdr")?;
            let header_path = format!("word/{}", header.filename);
            pkg.add_part(
                &header_path,
                content_type::WORDPROCESSING_HEADER,
                &header_xml,
            )?;

            doc_rels.add(Relationship::new(
                &header.rel_id,
                rel_type::HEADER,
                &header.filename,
            ));
        }

        // Write footers and add relationships
        for footer in self.footers.values() {
            let footer_xml = serialize_with_namespaces(&footer.body, "w:ftr")?;
            let footer_path = format!("word/{}", footer.filename);
            pkg.add_part(
                &footer_path,
                content_type::WORDPROCESSING_FOOTER,
                &footer_xml,
            )?;

            doc_rels.add(Relationship::new(
                &footer.rel_id,
                rel_type::FOOTER,
                &footer.filename,
            ));
        }

        // Write footnotes.xml if we have any footnotes
        if !self.footnotes.is_empty() {
            let fns = build_footnotes(&self.footnotes);
            let footnotes_xml = serialize_with_namespaces(&fns, "w:footnotes")?;
            pkg.add_part(
                "word/footnotes.xml",
                content_type::WORDPROCESSING_FOOTNOTES,
                &footnotes_xml,
            )?;

            let footnotes_rel_id = format!("rId{}", self.next_rel_id);
            self.next_rel_id += 1;
            doc_rels.add(Relationship::new(
                &footnotes_rel_id,
                rel_type::FOOTNOTES,
                "footnotes.xml",
            ));
        }

        // Write endnotes.xml if we have any endnotes
        if !self.endnotes.is_empty() {
            let ens = build_endnotes(&self.endnotes);
            let endnotes_xml = serialize_with_namespaces(&ens, "w:endnotes")?;
            pkg.add_part(
                "word/endnotes.xml",
                content_type::WORDPROCESSING_ENDNOTES,
                &endnotes_xml,
            )?;

            let endnotes_rel_id = format!("rId{}", self.next_rel_id);
            self.next_rel_id += 1;
            doc_rels.add(Relationship::new(
                &endnotes_rel_id,
                rel_type::ENDNOTES,
                "endnotes.xml",
            ));
        }

        // Write comments.xml if we have any comments
        if !self.comments.is_empty() {
            let comments = build_comments(&self.comments);
            let comments_xml = serialize_with_namespaces(&comments, "w:comments")?;
            pkg.add_part(
                "word/comments.xml",
                content_type::WORDPROCESSING_COMMENTS,
                &comments_xml,
            )?;

            let comments_rel_id = format!("rId{}", self.next_rel_id);
            self.next_rel_id += 1;
            doc_rels.add(Relationship::new(
                &comments_rel_id,
                rel_type::COMMENTS,
                "comments.xml",
            ));
        }

        // Write numbering.xml if we have any numbering definitions
        if !self.numberings.is_empty() {
            let numbering = build_numbering(&self.numberings);
            let num_xml = serialize_with_namespaces(&numbering, "w:numbering")?;
            pkg.add_part(
                "word/numbering.xml",
                content_type::WORDPROCESSING_NUMBERING,
                &num_xml,
            )?;

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

// =============================================================================
// Serialization helpers
// =============================================================================

/// Serialize a ToXml value to bytes with XML declaration prepended.
fn serialize_to_xml_bytes(value: &impl ToXml, tag: &str) -> Result<Vec<u8>> {
    let inner = Vec::new();
    let mut writer = quick_xml::Writer::new(inner);
    value.write_element(tag, &mut writer)?;
    let inner = writer.into_inner();
    let mut buf = Vec::with_capacity(
        b"<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n".len() + inner.len(),
    );
    buf.extend_from_slice(b"<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n");
    buf.extend_from_slice(&inner);
    Ok(buf)
}

/// Serialize a ToXml value with namespace declarations injected into the
/// root element's start tag. This is needed for types that don't have
/// `extra_attrs` (like Footnotes, Endnotes, Comments, Numbering, HeaderFooter).
fn serialize_with_namespaces(value: &impl ToXml, tag: &str) -> Result<Vec<u8>> {
    use quick_xml::events::{BytesEnd, BytesStart, Event};

    let inner = Vec::new();
    let mut writer = quick_xml::Writer::new(inner);

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
        b"<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n".len() + inner.len(),
    );
    buf.extend_from_slice(b"<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n");
    buf.extend_from_slice(&inner);
    Ok(buf)
}

// =============================================================================
// Build generated types from pending data
// =============================================================================

/// Build a separator footnote/endnote (required by Word).
///
/// Creates a FootnoteEndnote with a single paragraph containing a single run
/// with a separator or continuation separator element.
fn build_separator_ftn_edn(id: i64, ftn_type: types::STFtnEdn) -> types::FootnoteEndnote {
    let separator_content = match ftn_type {
        types::STFtnEdn::Separator => types::RunContent::Separator(Box::new(types::CTEmpty)),
        types::STFtnEdn::ContinuationSeparator => {
            types::RunContent::ContinuationSeparator(Box::new(types::CTEmpty))
        }
        _ => unreachable!("only Separator and ContinuationSeparator expected"),
    };

    let run = types::Run {
        rsid_r_pr: None,
        rsid_del: None,
        rsid_r: None,
        #[cfg(feature = "wml-styling")]
        r_pr: None,
        run_content: vec![separator_content],
        #[cfg(feature = "extra-attrs")]
        extra_attrs: std::collections::HashMap::new(),
        #[cfg(feature = "extra-children")]
        extra_children: Vec::new(),
    };

    let para = types::Paragraph {
        rsid_r_pr: None,
        rsid_r: None,
        rsid_del: None,
        rsid_p: None,
        rsid_r_default: None,
        #[cfg(feature = "wml-styling")]
        p_pr: None,
        paragraph_content: vec![types::ParagraphContent::R(Box::new(run))],
        #[cfg(feature = "extra-attrs")]
        extra_attrs: std::collections::HashMap::new(),
        #[cfg(feature = "extra-children")]
        extra_children: Vec::new(),
    };

    types::FootnoteEndnote {
        r#type: Some(ftn_type),
        id,
        block_content: vec![types::BlockContent::P(Box::new(para))],
        #[cfg(feature = "extra-attrs")]
        extra_attrs: std::collections::HashMap::new(),
        #[cfg(feature = "extra-children")]
        extra_children: Vec::new(),
    }
}

/// Build a Footnotes type from pending footnotes.
fn build_footnotes(footnotes: &HashMap<i32, PendingFootnote>) -> types::Footnotes {
    let mut fns = types::Footnotes {
        footnote: Vec::new(),
        #[cfg(feature = "extra-children")]
        extra_children: Vec::new(),
    };

    // Add separator footnotes (required by Word)
    fns.footnote
        .push(build_separator_ftn_edn(-1, types::STFtnEdn::Separator));
    fns.footnote.push(build_separator_ftn_edn(
        0,
        types::STFtnEdn::ContinuationSeparator,
    ));

    // Add user footnotes sorted by ID
    let mut sorted: Vec<_> = footnotes.values().collect();
    sorted.sort_by_key(|f| f.id);
    for footnote in sorted {
        fns.footnote.push(footnote.body.clone());
    }

    fns
}

/// Build an Endnotes type from pending endnotes.
fn build_endnotes(endnotes: &HashMap<i32, PendingEndnote>) -> types::Endnotes {
    let mut ens = types::Endnotes {
        endnote: Vec::new(),
        #[cfg(feature = "extra-children")]
        extra_children: Vec::new(),
    };

    // Add separator endnotes (required by Word)
    ens.endnote
        .push(build_separator_ftn_edn(-1, types::STFtnEdn::Separator));
    ens.endnote.push(build_separator_ftn_edn(
        0,
        types::STFtnEdn::ContinuationSeparator,
    ));

    // Add user endnotes sorted by ID
    let mut sorted: Vec<_> = endnotes.values().collect();
    sorted.sort_by_key(|e| e.id);
    for endnote in sorted {
        ens.endnote.push(endnote.body.clone());
    }

    ens
}

/// Build a Comments type from pending comments.
fn build_comments(comments: &HashMap<i32, PendingComment>) -> types::Comments {
    let mut result = types::Comments {
        comment: Vec::new(),
        #[cfg(feature = "extra-children")]
        extra_children: Vec::new(),
    };

    // Sort comments by ID for deterministic output
    let mut sorted: Vec<_> = comments.values().collect();
    sorted.sort_by_key(|c| c.id);

    for pc in sorted {
        let mut comment = pc.body.clone();
        comment.id = pc.id as i64;
        if let Some(ref author) = pc.author {
            comment.author = author.clone();
        }
        if let Some(ref date) = pc.date {
            comment.date = Some(date.clone());
        }
        if let Some(ref initials) = pc.initials {
            comment.initials = Some(initials.clone());
        }
        result.comment.push(comment);
    }

    result
}

/// Map ListType to STNumberFormat and level text.
fn list_type_to_num_fmt_and_text(list_type: ListType) -> (types::STNumberFormat, &'static str) {
    match list_type {
        ListType::Bullet => (types::STNumberFormat::Bullet, "\u{2022}"),
        ListType::Decimal => (types::STNumberFormat::Decimal, "%1."),
        ListType::LowerLetter => (types::STNumberFormat::LowerLetter, "%1."),
        ListType::UpperLetter => (types::STNumberFormat::UpperLetter, "%1."),
        ListType::LowerRoman => (types::STNumberFormat::LowerRoman, "%1."),
        ListType::UpperRoman => (types::STNumberFormat::UpperRoman, "%1."),
    }
}

/// Build a Numbering type from pending numbering definitions.
fn build_numbering(numberings: &HashMap<u32, PendingNumbering>) -> types::Numbering {
    let mut numbering = types::Numbering {
        #[cfg(feature = "wml-numbering")]
        num_pic_bullet: Vec::new(),
        abstract_num: Vec::new(),
        num: Vec::new(),
        #[cfg(feature = "wml-numbering")]
        num_id_mac_at_cleanup: None,
        #[cfg(feature = "extra-children")]
        extra_children: Vec::new(),
    };

    // Sort numberings by num_id for deterministic output
    let mut sorted: Vec<_> = numberings.values().collect();
    sorted.sort_by_key(|n| n.num_id);

    for pn in &sorted {
        let (num_fmt, lvl_text) = list_type_to_num_fmt_and_text(pn.list_type);

        let level = types::Level {
            ilvl: 0,
            tplc: None,
            tentative: None,
            start: Some(Box::new(types::CTDecimalNumber {
                value: 1,
                #[cfg(feature = "extra-attrs")]
                extra_attrs: std::collections::HashMap::new(),
            })),
            num_fmt: Some(Box::new(types::CTNumFmt {
                value: num_fmt,
                format: None,
                #[cfg(feature = "extra-attrs")]
                extra_attrs: std::collections::HashMap::new(),
            })),
            lvl_restart: None,
            paragraph_style: None,
            is_lgl: None,
            suff: None,
            lvl_text: Some(Box::new(types::CTLevelText {
                value: Some(lvl_text.to_string()),
                null: None,
                #[cfg(feature = "extra-attrs")]
                extra_attrs: std::collections::HashMap::new(),
            })),
            lvl_pic_bullet_id: None,
            legacy: None,
            lvl_jc: Some(Box::new(types::CTJc {
                value: types::STJc::Left,
                #[cfg(feature = "extra-attrs")]
                extra_attrs: std::collections::HashMap::new(),
            })),
            p_pr: None,
            r_pr: if pn.list_type == ListType::Bullet {
                // For bullet lists, use Symbol font
                Some(Box::new(build_bullet_run_properties()))
            } else {
                None
            },
            #[cfg(feature = "extra-attrs")]
            extra_attrs: std::collections::HashMap::new(),
            #[cfg(feature = "extra-children")]
            extra_children: Vec::new(),
        };

        let abs = types::AbstractNumbering {
            abstract_num_id: pn.abstract_num_id as i64,
            nsid: None,
            multi_level_type: None,
            tmpl: None,
            name: None,
            style_link: None,
            num_style_link: None,
            lvl: vec![level],
            #[cfg(feature = "extra-attrs")]
            extra_attrs: std::collections::HashMap::new(),
            #[cfg(feature = "extra-children")]
            extra_children: Vec::new(),
        };
        numbering.abstract_num.push(abs);

        let inst = types::NumberingInstance {
            num_id: pn.num_id as i64,
            abstract_num_id: Box::new(types::CTDecimalNumber {
                value: pn.abstract_num_id as i64,
                #[cfg(feature = "extra-attrs")]
                extra_attrs: std::collections::HashMap::new(),
            }),
            lvl_override: Vec::new(),
            #[cfg(feature = "extra-attrs")]
            extra_attrs: std::collections::HashMap::new(),
            #[cfg(feature = "extra-children")]
            extra_children: Vec::new(),
        };
        numbering.num.push(inst);
    }

    numbering
}

/// Build run properties for bullet list levels (Symbol font).
#[cfg(feature = "wml-styling")]
fn build_bullet_run_properties() -> types::RunProperties {
    types::RunProperties {
        fonts: Some(Box::new(types::Fonts {
            ascii: Some("Symbol".to_string()),
            h_ansi: Some("Symbol".to_string()),
            hint: Some(types::STHint::Default),
            ..Default::default()
        })),
        ..Default::default()
    }
}

/// Stub for when wml-styling is not enabled.
#[cfg(not(feature = "wml-styling"))]
fn build_bullet_run_properties() -> types::RunProperties {
    types::RunProperties::default()
}

// =============================================================================
// Drawing XML element builders
// =============================================================================

/// Build the `a:graphic` element containing a picture reference.
fn build_graphic_element(rel_id: &str, width: i64, height: i64, doc_id: usize) -> RawXmlElement {
    let blip = RawXmlElement {
        name: "a:blip".to_string(),
        attributes: vec![("r:embed".to_string(), rel_id.to_string())],
        children: vec![],
        self_closing: true,
    };

    let fill_rect = RawXmlElement {
        name: "a:fillRect".to_string(),
        attributes: vec![],
        children: vec![],
        self_closing: true,
    };

    let stretch = RawXmlElement {
        name: "a:stretch".to_string(),
        attributes: vec![],
        children: vec![RawXmlNode::Element(fill_rect)],
        self_closing: false,
    };

    let blip_fill = RawXmlElement {
        name: "pic:blipFill".to_string(),
        attributes: vec![],
        children: vec![RawXmlNode::Element(blip), RawXmlNode::Element(stretch)],
        self_closing: false,
    };

    let cnv_pr = RawXmlElement {
        name: "pic:cNvPr".to_string(),
        attributes: vec![
            ("id".to_string(), doc_id.to_string()),
            ("name".to_string(), format!("Picture {}", doc_id)),
        ],
        children: vec![],
        self_closing: true,
    };

    let cnv_pic_pr = RawXmlElement {
        name: "pic:cNvPicPr".to_string(),
        attributes: vec![],
        children: vec![],
        self_closing: true,
    };

    let nv_pic_pr = RawXmlElement {
        name: "pic:nvPicPr".to_string(),
        attributes: vec![],
        children: vec![RawXmlNode::Element(cnv_pr), RawXmlNode::Element(cnv_pic_pr)],
        self_closing: false,
    };

    let off = RawXmlElement {
        name: "a:off".to_string(),
        attributes: vec![
            ("x".to_string(), "0".to_string()),
            ("y".to_string(), "0".to_string()),
        ],
        children: vec![],
        self_closing: true,
    };

    let ext = RawXmlElement {
        name: "a:ext".to_string(),
        attributes: vec![
            ("cx".to_string(), width.to_string()),
            ("cy".to_string(), height.to_string()),
        ],
        children: vec![],
        self_closing: true,
    };

    let xfrm = RawXmlElement {
        name: "a:xfrm".to_string(),
        attributes: vec![],
        children: vec![RawXmlNode::Element(off), RawXmlNode::Element(ext)],
        self_closing: false,
    };

    let av_lst = RawXmlElement {
        name: "a:avLst".to_string(),
        attributes: vec![],
        children: vec![],
        self_closing: true,
    };

    let prst_geom = RawXmlElement {
        name: "a:prstGeom".to_string(),
        attributes: vec![("prst".to_string(), "rect".to_string())],
        children: vec![RawXmlNode::Element(av_lst)],
        self_closing: false,
    };

    let sp_pr = RawXmlElement {
        name: "pic:spPr".to_string(),
        attributes: vec![],
        children: vec![RawXmlNode::Element(xfrm), RawXmlNode::Element(prst_geom)],
        self_closing: false,
    };

    let pic = RawXmlElement {
        name: "pic:pic".to_string(),
        attributes: vec![],
        children: vec![
            RawXmlNode::Element(nv_pic_pr),
            RawXmlNode::Element(blip_fill),
            RawXmlNode::Element(sp_pr),
        ],
        self_closing: false,
    };

    let graphic_data = RawXmlElement {
        name: "a:graphicData".to_string(),
        attributes: vec![(
            "uri".to_string(),
            "http://schemas.openxmlformats.org/drawingml/2006/picture".to_string(),
        )],
        children: vec![RawXmlNode::Element(pic)],
        self_closing: false,
    };

    RawXmlElement {
        name: "a:graphic".to_string(),
        attributes: vec![],
        children: vec![RawXmlNode::Element(graphic_data)],
        self_closing: false,
    }
}

/// Build the `wp:inline` element for an inline image.
fn build_inline_image_element(image: &InlineImage, doc_id: usize) -> RawXmlElement {
    let width_emu = image.width_emu.unwrap_or(914400);
    let height_emu = image.height_emu.unwrap_or(914400);
    let desc = image.description.as_deref().unwrap_or("Image");

    let extent = RawXmlElement {
        name: "wp:extent".to_string(),
        attributes: vec![
            ("cx".to_string(), width_emu.to_string()),
            ("cy".to_string(), height_emu.to_string()),
        ],
        children: vec![],
        self_closing: true,
    };

    let doc_pr = RawXmlElement {
        name: "wp:docPr".to_string(),
        attributes: vec![
            ("id".to_string(), doc_id.to_string()),
            ("name".to_string(), format!("Picture {}", doc_id)),
            ("descr".to_string(), desc.to_string()),
        ],
        children: vec![],
        self_closing: true,
    };

    let graphic_frame_locks = RawXmlElement {
        name: "a:graphicFrameLocks".to_string(),
        attributes: vec![("noChangeAspect".to_string(), "1".to_string())],
        children: vec![],
        self_closing: true,
    };

    let cnv_graphic_frame_pr = RawXmlElement {
        name: "wp:cNvGraphicFramePr".to_string(),
        attributes: vec![],
        children: vec![RawXmlNode::Element(graphic_frame_locks)],
        self_closing: false,
    };

    let graphic = build_graphic_element(&image.rel_id, width_emu, height_emu, doc_id);

    RawXmlElement {
        name: "wp:inline".to_string(),
        attributes: vec![
            ("distT".to_string(), "0".to_string()),
            ("distB".to_string(), "0".to_string()),
            ("distL".to_string(), "0".to_string()),
            ("distR".to_string(), "0".to_string()),
        ],
        children: vec![
            RawXmlNode::Element(extent),
            RawXmlNode::Element(doc_pr),
            RawXmlNode::Element(cnv_graphic_frame_pr),
            RawXmlNode::Element(graphic),
        ],
        self_closing: false,
    }
}

/// Build the wrap type element for an anchored image.
fn build_wrap_element(wrap_type: WrapType) -> RawXmlElement {
    match wrap_type {
        WrapType::None => RawXmlElement {
            name: "wp:wrapNone".to_string(),
            attributes: vec![],
            children: vec![],
            self_closing: true,
        },
        WrapType::Square => RawXmlElement {
            name: "wp:wrapSquare".to_string(),
            attributes: vec![("wrapText".to_string(), "bothSides".to_string())],
            children: vec![],
            self_closing: true,
        },
        WrapType::Tight => {
            let polygon = build_default_wrap_polygon();
            RawXmlElement {
                name: "wp:wrapTight".to_string(),
                attributes: vec![("wrapText".to_string(), "bothSides".to_string())],
                children: vec![RawXmlNode::Element(polygon)],
                self_closing: false,
            }
        }
        WrapType::Through => {
            let polygon = build_default_wrap_polygon();
            RawXmlElement {
                name: "wp:wrapThrough".to_string(),
                attributes: vec![("wrapText".to_string(), "bothSides".to_string())],
                children: vec![RawXmlNode::Element(polygon)],
                self_closing: false,
            }
        }
        WrapType::TopAndBottom => RawXmlElement {
            name: "wp:wrapTopAndBottom".to_string(),
            attributes: vec![],
            children: vec![],
            self_closing: true,
        },
    }
}

/// Build a default rectangular wrap polygon.
fn build_default_wrap_polygon() -> RawXmlElement {
    let start = RawXmlElement {
        name: "wp:start".to_string(),
        attributes: vec![
            ("x".to_string(), "0".to_string()),
            ("y".to_string(), "0".to_string()),
        ],
        children: vec![],
        self_closing: true,
    };

    let line_to_1 = RawXmlElement {
        name: "wp:lineTo".to_string(),
        attributes: vec![
            ("x".to_string(), "0".to_string()),
            ("y".to_string(), "21600".to_string()),
        ],
        children: vec![],
        self_closing: true,
    };

    let line_to_2 = RawXmlElement {
        name: "wp:lineTo".to_string(),
        attributes: vec![
            ("x".to_string(), "21600".to_string()),
            ("y".to_string(), "21600".to_string()),
        ],
        children: vec![],
        self_closing: true,
    };

    let line_to_3 = RawXmlElement {
        name: "wp:lineTo".to_string(),
        attributes: vec![
            ("x".to_string(), "21600".to_string()),
            ("y".to_string(), "0".to_string()),
        ],
        children: vec![],
        self_closing: true,
    };

    let line_to_4 = RawXmlElement {
        name: "wp:lineTo".to_string(),
        attributes: vec![
            ("x".to_string(), "0".to_string()),
            ("y".to_string(), "0".to_string()),
        ],
        children: vec![],
        self_closing: true,
    };

    RawXmlElement {
        name: "wp:wrapPolygon".to_string(),
        attributes: vec![("edited".to_string(), "0".to_string())],
        children: vec![
            RawXmlNode::Element(start),
            RawXmlNode::Element(line_to_1),
            RawXmlNode::Element(line_to_2),
            RawXmlNode::Element(line_to_3),
            RawXmlNode::Element(line_to_4),
        ],
        self_closing: false,
    }
}

/// Build the `wp:anchor` element for an anchored (floating) image.
fn build_anchored_image_element(image: &AnchoredImage, doc_id: usize) -> RawXmlElement {
    let width_emu = image.width_emu.unwrap_or(914400);
    let height_emu = image.height_emu.unwrap_or(914400);
    let desc = image.description.as_deref().unwrap_or("Image");
    let behind_doc = if image.behind_doc { "1" } else { "0" };

    let simple_pos = RawXmlElement {
        name: "wp:simplePos".to_string(),
        attributes: vec![
            ("x".to_string(), "0".to_string()),
            ("y".to_string(), "0".to_string()),
        ],
        children: vec![],
        self_closing: true,
    };

    let pos_offset_h = RawXmlElement {
        name: "wp:posOffset".to_string(),
        attributes: vec![],
        children: vec![RawXmlNode::Text(image.pos_x.to_string())],
        self_closing: false,
    };

    let position_h = RawXmlElement {
        name: "wp:positionH".to_string(),
        attributes: vec![("relativeFrom".to_string(), "column".to_string())],
        children: vec![RawXmlNode::Element(pos_offset_h)],
        self_closing: false,
    };

    let pos_offset_v = RawXmlElement {
        name: "wp:posOffset".to_string(),
        attributes: vec![],
        children: vec![RawXmlNode::Text(image.pos_y.to_string())],
        self_closing: false,
    };

    let position_v = RawXmlElement {
        name: "wp:positionV".to_string(),
        attributes: vec![("relativeFrom".to_string(), "paragraph".to_string())],
        children: vec![RawXmlNode::Element(pos_offset_v)],
        self_closing: false,
    };

    let extent = RawXmlElement {
        name: "wp:extent".to_string(),
        attributes: vec![
            ("cx".to_string(), width_emu.to_string()),
            ("cy".to_string(), height_emu.to_string()),
        ],
        children: vec![],
        self_closing: true,
    };

    let effect_extent = RawXmlElement {
        name: "wp:effectExtent".to_string(),
        attributes: vec![
            ("l".to_string(), "0".to_string()),
            ("t".to_string(), "0".to_string()),
            ("r".to_string(), "0".to_string()),
            ("b".to_string(), "0".to_string()),
        ],
        children: vec![],
        self_closing: true,
    };

    let wrap = build_wrap_element(image.wrap_type);

    let doc_pr = RawXmlElement {
        name: "wp:docPr".to_string(),
        attributes: vec![
            ("id".to_string(), doc_id.to_string()),
            ("name".to_string(), format!("Picture {}", doc_id)),
            ("descr".to_string(), desc.to_string()),
        ],
        children: vec![],
        self_closing: true,
    };

    let graphic_frame_locks = RawXmlElement {
        name: "a:graphicFrameLocks".to_string(),
        attributes: vec![("noChangeAspect".to_string(), "1".to_string())],
        children: vec![],
        self_closing: true,
    };

    let cnv_graphic_frame_pr = RawXmlElement {
        name: "wp:cNvGraphicFramePr".to_string(),
        attributes: vec![],
        children: vec![RawXmlNode::Element(graphic_frame_locks)],
        self_closing: false,
    };

    let graphic = build_graphic_element(&image.rel_id, width_emu, height_emu, doc_id);

    RawXmlElement {
        name: "wp:anchor".to_string(),
        attributes: vec![
            ("distT".to_string(), "0".to_string()),
            ("distB".to_string(), "0".to_string()),
            ("distL".to_string(), "114300".to_string()),
            ("distR".to_string(), "114300".to_string()),
            ("simplePos".to_string(), "0".to_string()),
            ("relativeHeight".to_string(), "251658240".to_string()),
            ("behindDoc".to_string(), behind_doc.to_string()),
            ("locked".to_string(), "0".to_string()),
            ("layoutInCell".to_string(), "1".to_string()),
            ("allowOverlap".to_string(), "1".to_string()),
        ],
        children: vec![
            RawXmlNode::Element(simple_pos),
            RawXmlNode::Element(position_h),
            RawXmlNode::Element(position_v),
            RawXmlNode::Element(extent),
            RawXmlNode::Element(effect_extent),
            RawXmlNode::Element(wrap),
            RawXmlNode::Element(doc_pr),
            RawXmlNode::Element(cnv_graphic_frame_pr),
            RawXmlNode::Element(graphic),
        ],
        self_closing: false,
    }
}

// =============================================================================
// Utility functions
// =============================================================================

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

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_document_builder_simple() {
        let mut builder = DocumentBuilder::new();
        builder.add_paragraph("Hello, World!");
        builder.add_paragraph("Second paragraph");

        let body = builder.document.body.as_ref().unwrap();
        assert_eq!(body.block_content.len(), 2);
    }

    #[test]
    fn test_serialize_to_xml_bytes() {
        let doc = types::Document {
            background: None,
            body: Some(Box::new(types::Body::default())),
            conformance: None,
            #[cfg(feature = "extra-attrs")]
            extra_attrs: std::collections::HashMap::new(),
            #[cfg(feature = "extra-children")]
            extra_children: Vec::new(),
        };

        let bytes = serialize_to_xml_bytes(&doc, "w:document").unwrap();
        let xml = String::from_utf8(bytes).unwrap();
        assert!(xml.starts_with("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"));
        assert!(xml.contains("w:document"));
    }

    #[test]
    fn test_list_type_mapping() {
        let (fmt, text) = list_type_to_num_fmt_and_text(ListType::Bullet);
        assert!(matches!(fmt, types::STNumberFormat::Bullet));
        assert_eq!(text, "\u{2022}");

        let (fmt, text) = list_type_to_num_fmt_and_text(ListType::Decimal);
        assert!(matches!(fmt, types::STNumberFormat::Decimal));
        assert_eq!(text, "%1.");
    }

    #[test]
    fn test_extension_from_content_type() {
        assert_eq!(extension_from_content_type("image/png"), "png");
        assert_eq!(extension_from_content_type("image/jpeg"), "jpg");
        assert_eq!(extension_from_content_type("image/gif"), "gif");
        assert_eq!(extension_from_content_type("unknown/type"), "bin");
    }

    #[test]
    fn test_drawing_build() {
        let mut drawing = Drawing::new();
        drawing
            .add_image("rId1")
            .set_width_inches(1.0)
            .set_height_inches(1.0);

        let mut doc_id = 1;
        let ct_drawing = drawing.build(&mut doc_id);
        assert_eq!(doc_id, 2);

        #[cfg(feature = "extra-children")]
        assert_eq!(ct_drawing.extra_children.len(), 1);
        let _ = ct_drawing;
    }

    #[test]
    fn test_roundtrip_create_and_read() {
        use crate::Document;
        use crate::ext::BodyExt;
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
}
