//! Document writing and serialization.
//!
//! This module provides functionality for creating new Word documents
//! and saving existing documents.

use crate::document::{
    AnchoredImage, BlockContent, Body, Border, Cell, CellBorders, CellProperties, CellShading,
    CellWidth, ContentControl, CustomXml, DocGridType, Drawing, EmbeddedObject, GridColumn,
    HeaderFooterRef, HeaderFooterType, HeightRule, Hyperlink, InlineImage, NumberingProperties,
    PageOrientation, Paragraph, ParagraphBorders, ParagraphContent, ParagraphProperties, Row,
    RowHeight, RowProperties, Run, RunProperties, SectionProperties, TabStop, Table, TableBorders,
    TableProperties, TableWidth, VerticalMerge, VmlPicture, WrapType,
};
use crate::error::Result;
use crate::raw_xml::{PositionedAttr, PositionedNode, RawXmlNode};
use crate::styles::Styles;
use ooxml_opc::{PackageWriter, Relationship, Relationships, content_type, rel_type};
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

/// A pending header to be written to the package.
#[derive(Clone)]
pub struct PendingHeader {
    /// Header content (paragraphs, tables, etc.).
    pub body: Body,
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
    /// Footer content (paragraphs, tables, etc.).
    pub body: Body,
    /// Assigned relationship ID.
    pub rel_id: String,
    /// Footer type (default, first, even).
    pub footer_type: HeaderFooterType,
    /// Generated filename (e.g., "footer1.xml").
    pub filename: String,
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
    pub fn body_mut(&mut self) -> &mut Body {
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

/// A pending footnote to be written to the package.
#[derive(Clone)]
pub struct PendingFootnote {
    /// Footnote ID (referenced by FootnoteReference).
    pub id: i32,
    /// Footnote content (paragraphs, tables, etc.).
    pub body: Body,
}

/// A pending endnote to be written to the package.
#[derive(Clone)]
pub struct PendingEndnote {
    /// Endnote ID (referenced by EndnoteReference).
    pub id: i32,
    /// Endnote content (paragraphs, tables, etc.).
    pub body: Body,
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
    /// Comment content (paragraphs, tables, etc.).
    pub body: Body,
}

/// Builder for comment content.
pub struct CommentBuilder<'a> {
    builder: &'a mut DocumentBuilder,
    id: i32,
}

impl<'a> CommentBuilder<'a> {
    /// Get a mutable reference to the comment body.
    pub fn body_mut(&mut self) -> &mut Body {
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

/// Builder for footnote content.
pub struct FootnoteBuilder<'a> {
    builder: &'a mut DocumentBuilder,
    id: i32,
}

impl<'a> FootnoteBuilder<'a> {
    /// Get a mutable reference to the footnote body.
    pub fn body_mut(&mut self) -> &mut Body {
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
    pub fn body_mut(&mut self) -> &mut Body {
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

impl<'a> FooterBuilder<'a> {
    /// Get a mutable reference to the footer body.
    pub fn body_mut(&mut self) -> &mut Body {
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

    /// Add a header and return a builder for its content.
    ///
    /// The header will be automatically linked to the document's section properties.
    ///
    /// # Example
    ///
    /// ```ignore
    /// let mut builder = DocumentBuilder::new();
    /// let mut header = builder.add_header(HeaderFooterType::Default);
    /// header.add_paragraph("Document Title");
    /// ```
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
                body: Body::new(),
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
    ///
    /// # Example
    ///
    /// ```ignore
    /// let mut builder = DocumentBuilder::new();
    /// let mut footer = builder.add_footer(HeaderFooterType::Default);
    /// footer.add_paragraph("Page 1");
    /// ```
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
                body: Body::new(),
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
    /// Use the returned `id` when adding a `FootnoteReference` to a Run.
    ///
    /// # Example
    ///
    /// ```ignore
    /// use ooxml_wml::document::FootnoteReference;
    ///
    /// let mut builder = DocumentBuilder::new();
    /// let mut footnote = builder.add_footnote();
    /// footnote.add_paragraph("This is a footnote.");
    /// let footnote_id = footnote.id();
    ///
    /// // In a paragraph, add a reference to the footnote
    /// let run = builder.body_mut().add_paragraph().add_run();
    /// run.set_text("Some text");
    /// run.set_footnote_ref(FootnoteReference { id: footnote_id });
    /// ```
    pub fn add_footnote(&mut self) -> FootnoteBuilder<'_> {
        let id = self.next_footnote_id;
        self.next_footnote_id += 1;

        self.footnotes.insert(
            id,
            PendingFootnote {
                id,
                body: Body::new(),
            },
        );

        FootnoteBuilder { builder: self, id }
    }

    /// Add an endnote and return a builder for its content.
    ///
    /// Use the returned `id` when adding an `EndnoteReference` to a Run.
    ///
    /// # Example
    ///
    /// ```ignore
    /// use ooxml_wml::document::EndnoteReference;
    ///
    /// let mut builder = DocumentBuilder::new();
    /// let mut endnote = builder.add_endnote();
    /// endnote.add_paragraph("This is an endnote.");
    /// let endnote_id = endnote.id();
    ///
    /// // In a paragraph, add a reference to the endnote
    /// let run = builder.body_mut().add_paragraph().add_run();
    /// run.set_text("Some text");
    /// run.set_endnote_ref(EndnoteReference { id: endnote_id });
    /// ```
    pub fn add_endnote(&mut self) -> EndnoteBuilder<'_> {
        let id = self.next_endnote_id;
        self.next_endnote_id += 1;

        self.endnotes.insert(
            id,
            PendingEndnote {
                id,
                body: Body::new(),
            },
        );

        EndnoteBuilder { builder: self, id }
    }

    /// Add a comment and return a builder for its content.
    ///
    /// Use the returned `id` when adding comment ranges and references to the document.
    ///
    /// # Example
    ///
    /// ```ignore
    /// use ooxml_wml::document::{CommentReference, CommentRangeStart, CommentRangeEnd};
    ///
    /// let mut builder = DocumentBuilder::new();
    /// let mut comment = builder.add_comment();
    /// comment.set_author("John Doe");
    /// comment.add_paragraph("This needs review.");
    /// let comment_id = comment.id();
    ///
    /// // In a paragraph, mark the commented text with ranges and reference
    /// let para = builder.body_mut().add_paragraph();
    /// para.add_comment_range_start(comment_id);
    /// para.add_run().set_text("Commented text");
    /// para.add_comment_range_end(comment_id);
    /// // Add the comment reference (shows the comment marker)
    /// para.add_run().set_comment_ref(CommentReference { id: comment_id });
    /// ```
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
                body: Body::new(),
            },
        );

        CommentBuilder { builder: self, id }
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
            if self.body.section_properties().is_none() {
                self.body
                    .set_section_properties(SectionProperties::default());
            }
            let sect_pr = self.body.section_properties_mut().as_mut().unwrap();

            // Add header references
            for header in self.headers.values() {
                sect_pr.headers.push(HeaderFooterRef {
                    rel_id: header.rel_id.clone(),
                    hf_type: header.header_type,
                });
            }

            // Add footer references
            for footer in self.footers.values() {
                sect_pr.footers.push(HeaderFooterRef {
                    rel_id: footer.rel_id.clone(),
                    hf_type: footer.footer_type,
                });
            }
        }

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
            let header_xml = serialize_header(&header.body);
            let header_path = format!("word/{}", header.filename);
            pkg.add_part(
                &header_path,
                content_type::WORDPROCESSING_HEADER,
                header_xml.as_bytes(),
            )?;

            doc_rels.add(Relationship::new(
                &header.rel_id,
                rel_type::HEADER,
                &header.filename,
            ));
        }

        // Write footers and add relationships
        for footer in self.footers.values() {
            let footer_xml = serialize_footer(&footer.body);
            let footer_path = format!("word/{}", footer.filename);
            pkg.add_part(
                &footer_path,
                content_type::WORDPROCESSING_FOOTER,
                footer_xml.as_bytes(),
            )?;

            doc_rels.add(Relationship::new(
                &footer.rel_id,
                rel_type::FOOTER,
                &footer.filename,
            ));
        }

        // Write footnotes.xml if we have any footnotes
        if !self.footnotes.is_empty() {
            let footnotes_xml = serialize_footnotes(&self.footnotes);
            pkg.add_part(
                "word/footnotes.xml",
                content_type::WORDPROCESSING_FOOTNOTES,
                footnotes_xml.as_bytes(),
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
            let endnotes_xml = serialize_endnotes(&self.endnotes);
            pkg.add_part(
                "word/endnotes.xml",
                content_type::WORDPROCESSING_ENDNOTES,
                endnotes_xml.as_bytes(),
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
            let comments_xml = serialize_comments(&self.comments);
            pkg.add_part(
                "word/comments.xml",
                content_type::WORDPROCESSING_COMMENTS,
                comments_xml.as_bytes(),
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

/// Serialize header content to XML.
///
/// Headers use the `<w:hdr>` root element with the same content model as body.
pub fn serialize_header(body: &Body) -> String {
    let mut xml = String::new();

    // XML declaration
    xml.push_str(r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>"#);
    xml.push('\n');

    // Header element with namespaces
    xml.push_str(&format!(
        r#"<w:hdr xmlns:w="{}" xmlns:r="{}" xmlns:wp="{}" xmlns:a="{}" xmlns:pic="{}">"#,
        NS_W, NS_R, NS_WP, NS_A, NS_PIC
    ));

    // Content (paragraphs, tables, etc.)
    for block in body.content() {
        serialize_block_content(block, &mut xml);
    }
    serialize_unknown_children(&body.unknown_children, &mut xml);

    xml.push_str("</w:hdr>");
    xml
}

/// Serialize footer content to XML.
///
/// Footers use the `<w:ftr>` root element with the same content model as body.
pub fn serialize_footer(body: &Body) -> String {
    let mut xml = String::new();

    // XML declaration
    xml.push_str(r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>"#);
    xml.push('\n');

    // Footer element with namespaces
    xml.push_str(&format!(
        r#"<w:ftr xmlns:w="{}" xmlns:r="{}" xmlns:wp="{}" xmlns:a="{}" xmlns:pic="{}">"#,
        NS_W, NS_R, NS_WP, NS_A, NS_PIC
    ));

    // Content (paragraphs, tables, etc.)
    for block in body.content() {
        serialize_block_content(block, &mut xml);
    }
    serialize_unknown_children(&body.unknown_children, &mut xml);

    xml.push_str("</w:ftr>");
    xml
}

/// Serialize footnotes to XML.
///
/// Footnotes use the `<w:footnotes>` root element containing individual `<w:footnote>` elements.
fn serialize_footnotes(footnotes: &HashMap<i32, PendingFootnote>) -> String {
    let mut xml = String::new();

    // XML declaration
    xml.push_str(r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>"#);
    xml.push('\n');

    // Footnotes element with namespaces
    xml.push_str(&format!(
        r#"<w:footnotes xmlns:w="{}" xmlns:r="{}" xmlns:wp="{}" xmlns:a="{}" xmlns:pic="{}">"#,
        NS_W, NS_R, NS_WP, NS_A, NS_PIC
    ));

    // Add separator footnotes (required by Word)
    // ID -1: separator, ID 0: continuation separator
    xml.push_str(r#"<w:footnote w:type="separator" w:id="-1"><w:p><w:r><w:separator/></w:r></w:p></w:footnote>"#);
    xml.push_str(r#"<w:footnote w:type="continuationSeparator" w:id="0"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:footnote>"#);

    // Sort footnotes by ID for deterministic output
    let mut sorted: Vec<_> = footnotes.values().collect();
    sorted.sort_by_key(|f| f.id);

    // Serialize each footnote
    for footnote in sorted {
        xml.push_str(&format!(r#"<w:footnote w:id="{}">"#, footnote.id));
        for block in footnote.body.content() {
            serialize_block_content(block, &mut xml);
        }
        serialize_unknown_children(&footnote.body.unknown_children, &mut xml);
        xml.push_str("</w:footnote>");
    }

    xml.push_str("</w:footnotes>");
    xml
}

/// Serialize endnotes to XML.
///
/// Endnotes use the `<w:endnotes>` root element containing individual `<w:endnote>` elements.
fn serialize_endnotes(endnotes: &HashMap<i32, PendingEndnote>) -> String {
    let mut xml = String::new();

    // XML declaration
    xml.push_str(r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>"#);
    xml.push('\n');

    // Endnotes element with namespaces
    xml.push_str(&format!(
        r#"<w:endnotes xmlns:w="{}" xmlns:r="{}" xmlns:wp="{}" xmlns:a="{}" xmlns:pic="{}">"#,
        NS_W, NS_R, NS_WP, NS_A, NS_PIC
    ));

    // Add separator endnotes (required by Word)
    xml.push_str(r#"<w:endnote w:type="separator" w:id="-1"><w:p><w:r><w:separator/></w:r></w:p></w:endnote>"#);
    xml.push_str(r#"<w:endnote w:type="continuationSeparator" w:id="0"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:endnote>"#);

    // Sort endnotes by ID for deterministic output
    let mut sorted: Vec<_> = endnotes.values().collect();
    sorted.sort_by_key(|e| e.id);

    // Serialize each endnote
    for endnote in sorted {
        xml.push_str(&format!(r#"<w:endnote w:id="{}">"#, endnote.id));
        for block in endnote.body.content() {
            serialize_block_content(block, &mut xml);
        }
        serialize_unknown_children(&endnote.body.unknown_children, &mut xml);
        xml.push_str("</w:endnote>");
    }

    xml.push_str("</w:endnotes>");
    xml
}

/// Serialize comments to XML.
///
/// Comments use the `<w:comments>` root element containing individual `<w:comment>` elements.
fn serialize_comments(comments: &HashMap<i32, PendingComment>) -> String {
    let mut xml = String::new();

    // XML declaration
    xml.push_str(r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>"#);
    xml.push('\n');

    // Comments element with namespaces
    xml.push_str(&format!(
        r#"<w:comments xmlns:w="{}" xmlns:r="{}" xmlns:wp="{}" xmlns:a="{}" xmlns:pic="{}">"#,
        NS_W, NS_R, NS_WP, NS_A, NS_PIC
    ));

    // Sort comments by ID for deterministic output
    let mut sorted: Vec<_> = comments.values().collect();
    sorted.sort_by_key(|c| c.id);

    // Serialize each comment
    for comment in sorted {
        xml.push_str(&format!(r#"<w:comment w:id="{}""#, comment.id));

        // Add optional attributes
        if let Some(ref author) = comment.author {
            xml.push_str(&format!(r#" w:author="{}""#, escape_xml(author)));
        }
        if let Some(ref date) = comment.date {
            xml.push_str(&format!(r#" w:date="{}""#, date));
        }
        if let Some(ref initials) = comment.initials {
            xml.push_str(&format!(r#" w:initials="{}""#, escape_xml(initials)));
        }

        xml.push('>');

        // Comment content
        for block in comment.body.content() {
            serialize_block_content(block, &mut xml);
        }
        serialize_unknown_children(&comment.body.unknown_children, &mut xml);

        xml.push_str("</w:comment>");
    }

    xml.push_str("</w:comments>");
    xml
}

/// Serialize body contents.
fn serialize_body(body: &Body, xml: &mut String) {
    for block in body.content() {
        serialize_block_content(block, xml);
    }
    // Section properties (must come after all block content)
    if let Some(sect_pr) = body.section_properties() {
        serialize_section_properties(sect_pr, xml);
    }
    // Write unknown children preserved for round-trip fidelity
    serialize_unknown_children(&body.unknown_children, xml);
}

/// Serialize a block content element.
fn serialize_block_content(block: &BlockContent, xml: &mut String) {
    match block {
        BlockContent::Paragraph(para) => serialize_paragraph(para, xml),
        BlockContent::Table(table) => serialize_table(table, xml),
        BlockContent::ContentControl(sdt) => serialize_content_control(sdt, xml),
        BlockContent::CustomXml(custom) => serialize_custom_xml(custom, xml),
    }
}

/// Serialize a content control (SDT).
fn serialize_content_control(sdt: &ContentControl, xml: &mut String) {
    xml.push_str("<w:sdt>");

    // SDT properties
    let has_props = sdt.tag.is_some() || sdt.alias.is_some();
    if has_props {
        xml.push_str("<w:sdtPr>");
        if let Some(tag) = &sdt.tag {
            xml.push_str("<w:tag w:val=\"");
            xml.push_str(&escape_xml(tag));
            xml.push_str("\"/>");
        }
        if let Some(alias) = &sdt.alias {
            xml.push_str("<w:alias w:val=\"");
            xml.push_str(&escape_xml(alias));
            xml.push_str("\"/>");
        }
        xml.push_str("</w:sdtPr>");
    }

    // SDT content
    xml.push_str("<w:sdtContent>");
    for block in &sdt.content {
        serialize_block_content(block, xml);
    }
    xml.push_str("</w:sdtContent>");

    serialize_unknown_children(&sdt.unknown_children, xml);
    xml.push_str("</w:sdt>");
}

/// Serialize a custom XML block.
fn serialize_custom_xml(custom: &CustomXml, xml: &mut String) {
    xml.push_str("<w:customXml");
    if let Some(uri) = &custom.uri {
        xml.push_str(" w:uri=\"");
        xml.push_str(&escape_xml(uri));
        xml.push('"');
    }
    if let Some(element) = &custom.element {
        xml.push_str(" w:element=\"");
        xml.push_str(&escape_xml(element));
        xml.push('"');
    }
    xml.push('>');

    for block in &custom.content {
        serialize_block_content(block, xml);
    }

    serialize_unknown_children(&custom.unknown_children, xml);
    xml.push_str("</w:customXml>");
}

/// Serialize section properties.
fn serialize_section_properties(props: &SectionProperties, xml: &mut String) {
    xml.push_str("<w:sectPr>");

    // Section type
    if let Some(section_type) = &props.section_type {
        xml.push_str("<w:type w:val=\"");
        xml.push_str(section_type.as_str());
        xml.push_str("\"/>");
    }

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

    // Columns
    if let Some(cols) = &props.columns {
        xml.push_str("<w:cols");
        if let Some(num) = cols.num {
            xml.push_str(" w:num=\"");
            xml.push_str(&num.to_string());
            xml.push('"');
        }
        if let Some(space) = cols.space {
            xml.push_str(" w:space=\"");
            xml.push_str(&space.to_string());
            xml.push('"');
        }
        if cols.equal_width {
            xml.push_str(" w:equalWidth=\"true\"");
        }
        if cols.separator {
            xml.push_str(" w:sep=\"true\"");
        }
        if cols.columns.is_empty() {
            xml.push_str("/>");
        } else {
            xml.push('>');
            for col in &cols.columns {
                xml.push_str("<w:col w:w=\"");
                xml.push_str(&col.width.to_string());
                xml.push('"');
                if let Some(space) = col.space {
                    xml.push_str(" w:space=\"");
                    xml.push_str(&space.to_string());
                    xml.push('"');
                }
                xml.push_str("/>");
            }
            xml.push_str("</w:cols>");
        }
    }

    // Document grid
    if let Some(doc_grid) = &props.doc_grid {
        xml.push_str("<w:docGrid");
        if doc_grid.grid_type != DocGridType::Default {
            xml.push_str(" w:type=\"");
            xml.push_str(doc_grid.grid_type.as_str());
            xml.push('"');
        }
        if let Some(line_pitch) = doc_grid.line_pitch {
            xml.push_str(" w:linePitch=\"");
            xml.push_str(&line_pitch.to_string());
            xml.push('"');
        }
        if let Some(char_space) = doc_grid.char_space {
            xml.push_str(" w:charSpace=\"");
            xml.push_str(&char_space.to_string());
            xml.push('"');
        }
        xml.push_str("/>");
    }

    // Header references
    for header_ref in &props.headers {
        xml.push_str("<w:headerReference w:type=\"");
        xml.push_str(header_ref.hf_type.as_str());
        xml.push_str("\" r:id=\"");
        xml.push_str(&header_ref.rel_id);
        xml.push_str("\"/>");
    }

    // Footer references
    for footer_ref in &props.footers {
        xml.push_str("<w:footerReference w:type=\"");
        xml.push_str(footer_ref.hf_type.as_str());
        xml.push_str("\" r:id=\"");
        xml.push_str(&footer_ref.rel_id);
        xml.push_str("\"/>");
    }

    // Unknown children for round-trip preservation
    serialize_unknown_children(&props.unknown_children, xml);

    xml.push_str("</w:sectPr>");
}

/// Serialize a table.
fn serialize_table(table: &Table, xml: &mut String) {
    xml.push_str("<w:tbl>");
    if let Some(props) = table.properties() {
        serialize_table_properties(props, xml);
    }
    if !table.grid_columns().is_empty() {
        serialize_table_grid(table.grid_columns(), xml);
    }
    for row in table.rows() {
        serialize_row(row, xml);
    }
    serialize_unknown_children(&table.unknown_children, xml);
    xml.push_str("</w:tbl>");
}

/// Serialize table properties.
fn serialize_table_properties(props: &TableProperties, xml: &mut String) {
    xml.push_str("<w:tblPr>");
    if let Some(width) = &props.width {
        serialize_table_width(width, xml);
    }
    if let Some(justification) = props.justification {
        xml.push_str("<w:jc w:val=\"");
        xml.push_str(justification.as_str());
        xml.push_str("\"/>");
    }
    if let Some(indent) = props.indent {
        xml.push_str("<w:tblInd w:w=\"");
        xml.push_str(&indent.to_string());
        xml.push_str("\" w:type=\"dxa\"/>");
    }
    if let Some(borders) = &props.borders {
        serialize_table_borders(borders, xml);
    }
    if let Some(shading) = &props.shading {
        serialize_cell_shading(shading, xml);
    }
    if let Some(layout) = props.layout {
        xml.push_str("<w:tblLayout w:type=\"");
        xml.push_str(layout.as_str());
        xml.push_str("\"/>");
    }
    serialize_unknown_children(&props.unknown_children, xml);
    xml.push_str("</w:tblPr>");
}

/// Serialize table width.
fn serialize_table_width(width: &TableWidth, xml: &mut String) {
    xml.push_str("<w:tblW w:w=\"");
    xml.push_str(&width.width.to_string());
    xml.push_str("\" w:type=\"");
    xml.push_str(width.width_type.as_str());
    xml.push_str("\"/>");
}

/// Serialize table borders.
fn serialize_table_borders(borders: &TableBorders, xml: &mut String) {
    xml.push_str("<w:tblBorders>");
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
    xml.push_str("</w:tblBorders>");
}

/// Serialize table grid columns.
fn serialize_table_grid(columns: &[GridColumn], xml: &mut String) {
    xml.push_str("<w:tblGrid>");
    for col in columns {
        xml.push_str("<w:gridCol w:w=\"");
        xml.push_str(&col.width.to_string());
        xml.push_str("\"/>");
    }
    xml.push_str("</w:tblGrid>");
}

/// Serialize a table row.
fn serialize_row(row: &Row, xml: &mut String) {
    xml.push_str("<w:tr>");
    if let Some(props) = row.properties() {
        serialize_row_properties(props, xml);
    }
    for cell in row.cells() {
        serialize_cell(cell, xml);
    }
    serialize_unknown_children(&row.unknown_children, xml);
    xml.push_str("</w:tr>");
}

/// Serialize row properties.
fn serialize_row_properties(props: &RowProperties, xml: &mut String) {
    xml.push_str("<w:trPr>");
    if let Some(height) = &props.height {
        serialize_row_height(height, xml);
    }
    if props.is_header {
        xml.push_str("<w:tblHeader/>");
    }
    if props.cant_split {
        xml.push_str("<w:cantSplit/>");
    }
    serialize_unknown_children(&props.unknown_children, xml);
    xml.push_str("</w:trPr>");
}

/// Serialize row height.
fn serialize_row_height(height: &RowHeight, xml: &mut String) {
    xml.push_str("<w:trHeight w:val=\"");
    xml.push_str(&height.value.to_string());
    xml.push('"');
    if height.rule != HeightRule::Auto {
        xml.push_str(" w:hRule=\"");
        xml.push_str(height.rule.as_str());
        xml.push('"');
    }
    xml.push_str("/>");
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

    // Content (runs, hyperlinks, bookmarks)
    for content in para.content() {
        match content {
            ParagraphContent::Run(run) => serialize_run(run, xml),
            ParagraphContent::Hyperlink(link) => serialize_hyperlink(link, xml),
            ParagraphContent::BookmarkStart(bookmark) => {
                xml.push_str(&format!(
                    r#"<w:bookmarkStart w:id="{}" w:name="{}"/>"#,
                    bookmark.id,
                    escape_xml(&bookmark.name)
                ));
            }
            ParagraphContent::BookmarkEnd(bookmark) => {
                xml.push_str(&format!(r#"<w:bookmarkEnd w:id="{}"/>"#, bookmark.id));
            }
            ParagraphContent::CommentRangeStart(comment) => {
                xml.push_str(&format!(r#"<w:commentRangeStart w:id="{}"/>"#, comment.id));
            }
            ParagraphContent::CommentRangeEnd(comment) => {
                xml.push_str(&format!(r#"<w:commentRangeEnd w:id="{}"/>"#, comment.id));
            }
            ParagraphContent::SimpleField(field) => {
                xml.push_str("<w:fldSimple w:instr=\"");
                xml.push_str(&escape_xml(&field.instruction));
                xml.push_str("\">");
                for run in &field.runs {
                    serialize_run(run, xml);
                }
                xml.push_str("</w:fldSimple>");
            }
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

    // Paragraph borders
    if let Some(borders) = &props.borders {
        serialize_paragraph_borders(borders, xml);
    }

    // Paragraph shading
    if let Some(shading) = &props.shading {
        serialize_cell_shading(shading, xml);
    }

    // Outline level
    if let Some(level) = props.outline_level {
        xml.push_str(&format!(r#"<w:outlineLvl w:val="{}"/>"#, level));
    }

    // Flow control properties
    if props.keep_next {
        xml.push_str("<w:keepNext/>");
    }
    if props.keep_lines {
        xml.push_str("<w:keepLines/>");
    }
    if props.page_break_before {
        xml.push_str("<w:pageBreakBefore/>");
    }
    if let Some(widow_control) = props.widow_control {
        if widow_control {
            xml.push_str("<w:widowControl/>");
        } else {
            xml.push_str(r#"<w:widowControl w:val="0"/>"#);
        }
    }

    // Tab stops
    if !props.tabs.is_empty() {
        serialize_tab_stops(&props.tabs, xml);
    }

    serialize_unknown_children(&props.unknown_children, xml);
    xml.push_str("</w:pPr>");
}

/// Serialize tab stops.
fn serialize_tab_stops(tabs: &[TabStop], xml: &mut String) {
    xml.push_str("<w:tabs>");
    for tab in tabs {
        xml.push_str("<w:tab w:val=\"");
        xml.push_str(tab.tab_type.as_str());
        xml.push_str("\" w:pos=\"");
        xml.push_str(&tab.position.to_string());
        xml.push('"');
        if let Some(leader) = tab.leader {
            xml.push_str(" w:leader=\"");
            xml.push_str(leader.as_str());
            xml.push('"');
        }
        xml.push_str("/>");
    }
    xml.push_str("</w:tabs>");
}

/// Serialize paragraph borders.
fn serialize_paragraph_borders(borders: &ParagraphBorders, xml: &mut String) {
    xml.push_str("<w:pBdr>");
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
    if let Some(border) = &borders.between {
        xml.push_str("<w:between");
        serialize_border_attrs(border, xml);
        xml.push_str("/>");
    }
    if let Some(border) = &borders.bar {
        xml.push_str("<w:bar");
        serialize_border_attrs(border, xml);
        xml.push_str("/>");
    }
    xml.push_str("</w:pBdr>");
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

    // VML pictures (legacy images)
    for vml_pict in run.vml_pictures() {
        serialize_vml_picture(vml_pict, xml);
    }

    // Embedded objects
    for obj in run.embedded_objects() {
        serialize_embedded_object(obj, xml);
    }

    // Symbols
    for symbol in run.symbols() {
        xml.push_str("<w:sym w:font=\"");
        xml.push_str(&escape_xml(&symbol.font));
        xml.push_str("\" w:char=\"");
        xml.push_str(&escape_xml(&symbol.char_code));
        xml.push_str("\"/>");
    }

    // Field character (complex field marker)
    if let Some(field_char) = run.field_char() {
        xml.push_str("<w:fldChar w:fldCharType=\"");
        xml.push_str(field_char.field_type.as_str());
        xml.push_str("\"/>");
    }

    // Field instruction text
    if let Some(instr_text) = run.instr_text() {
        xml.push_str("<w:instrText>");
        xml.push_str(&escape_xml(instr_text));
        xml.push_str("</w:instrText>");
    }

    // Footnote reference
    if let Some(footnote_ref) = run.footnote_ref() {
        xml.push_str("<w:footnoteReference w:id=\"");
        xml.push_str(&footnote_ref.id.to_string());
        xml.push_str("\"/>");
    }

    // Endnote reference
    if let Some(endnote_ref) = run.endnote_ref() {
        xml.push_str("<w:endnoteReference w:id=\"");
        xml.push_str(&endnote_ref.id.to_string());
        xml.push_str("\"/>");
    }

    // Comment reference
    if let Some(comment_ref) = run.comment_ref() {
        xml.push_str("<w:commentReference w:id=\"");
        xml.push_str(&comment_ref.id.to_string());
        xml.push_str("\"/>");
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
    let mut doc_id = 1;
    for image in drawing.images() {
        serialize_inline_image(image, doc_id, xml);
        doc_id += 1;
    }
    for image in drawing.anchored_images() {
        serialize_anchored_image(image, doc_id, xml);
        doc_id += 1;
    }
    serialize_unknown_children(&drawing.unknown_children, xml);
    xml.push_str("</w:drawing>");
}

/// Serialize a VML picture element (legacy image format).
fn serialize_vml_picture(vml_pict: &VmlPicture, xml: &mut String) {
    xml.push_str("<w:pict");
    for (key, value) in &vml_pict.attributes {
        xml.push(' ');
        xml.push_str(key);
        xml.push_str("=\"");
        xml.push_str(&escape_xml(value));
        xml.push('"');
    }
    xml.push('>');

    // Serialize children using the RawXmlNode serialization
    for child in &vml_pict.children {
        serialize_raw_xml_node(child, xml);
    }

    xml.push_str("</w:pict>");
}

/// Serialize an embedded OLE object.
fn serialize_embedded_object(obj: &EmbeddedObject, xml: &mut String) {
    xml.push_str("<w:object");
    for (key, value) in &obj.attributes {
        xml.push(' ');
        xml.push_str(key);
        xml.push_str("=\"");
        xml.push_str(&escape_xml(value));
        xml.push('"');
    }
    xml.push('>');

    // Serialize children using the RawXmlNode serialization
    for child in &obj.children {
        serialize_raw_xml_node(child, xml);
    }

    xml.push_str("</w:object>");
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

/// Serialize an anchored (floating) image.
///
/// Generates the DrawingML structure required for an anchored image with text wrapping.
fn serialize_anchored_image(image: &AnchoredImage, doc_id: usize, xml: &mut String) {
    // Default dimensions: 1 inch x 1 inch (914400 EMUs)
    let cx = image.width_emu().unwrap_or(914400);
    let cy = image.height_emu().unwrap_or(914400);
    let rel_id = image.rel_id();
    let descr = image.description().unwrap_or("Image");
    let behind_doc = if image.is_behind_doc() { "1" } else { "0" };

    // Anchor element with positioning attributes
    xml.push_str(&format!(
        r#"<wp:anchor distT="0" distB="0" distL="114300" distR="114300" simplePos="0" relativeHeight="251658240" behindDoc="{}" locked="0" layoutInCell="1" allowOverlap="1">"#,
        behind_doc
    ));

    // Simple position (unused but required)
    xml.push_str(r#"<wp:simplePos x="0" y="0"/>"#);

    // Horizontal position
    xml.push_str(r#"<wp:positionH relativeFrom="column">"#);
    xml.push_str(&format!(
        r#"<wp:posOffset>{}</wp:posOffset>"#,
        image.pos_x()
    ));
    xml.push_str(r#"</wp:positionH>"#);

    // Vertical position
    xml.push_str(r#"<wp:positionV relativeFrom="paragraph">"#);
    xml.push_str(&format!(
        r#"<wp:posOffset>{}</wp:posOffset>"#,
        image.pos_y()
    ));
    xml.push_str(r#"</wp:positionV>"#);

    // Extent
    xml.push_str(&format!(r#"<wp:extent cx="{}" cy="{}"/>"#, cx, cy));

    // Effect extent (no effects)
    xml.push_str(r#"<wp:effectExtent l="0" t="0" r="0" b="0"/>"#);

    // Wrap type
    match image.wrap_type() {
        WrapType::None => xml.push_str(r#"<wp:wrapNone/>"#),
        WrapType::Square => xml.push_str(r#"<wp:wrapSquare wrapText="bothSides"/>"#),
        WrapType::Tight => xml.push_str(r#"<wp:wrapTight wrapText="bothSides"><wp:wrapPolygon edited="0"><wp:start x="0" y="0"/><wp:lineTo x="0" y="21600"/><wp:lineTo x="21600" y="21600"/><wp:lineTo x="21600" y="0"/><wp:lineTo x="0" y="0"/></wp:wrapPolygon></wp:wrapTight>"#),
        WrapType::Through => xml.push_str(r#"<wp:wrapThrough wrapText="bothSides"><wp:wrapPolygon edited="0"><wp:start x="0" y="0"/><wp:lineTo x="0" y="21600"/><wp:lineTo x="21600" y="21600"/><wp:lineTo x="21600" y="0"/><wp:lineTo x="0" y="0"/></wp:wrapPolygon></wp:wrapThrough>"#),
        WrapType::TopAndBottom => xml.push_str(r#"<wp:wrapTopAndBottom/>"#),
    }

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
    xml.push_str(r#"</wp:anchor>"#);
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
        || props.hidden
        || props.shading.is_some()
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

    if props.hidden {
        xml.push_str("<w:vanish/>");
    }

    if let Some(ref shading) = props.shading {
        serialize_cell_shading(shading, xml);
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

    #[test]
    fn test_serialize_header() {
        let mut body = Body::new();
        body.add_paragraph().add_run().set_text("Header text");

        let xml = serialize_header(&body);

        assert!(xml.contains("<w:hdr"));
        assert!(xml.contains("</w:hdr>"));
        assert!(xml.contains("<w:t>Header text</w:t>"));
    }

    #[test]
    fn test_serialize_footer() {
        let mut body = Body::new();
        body.add_paragraph().add_run().set_text("Footer text");

        let xml = serialize_footer(&body);

        assert!(xml.contains("<w:ftr"));
        assert!(xml.contains("</w:ftr>"));
        assert!(xml.contains("<w:t>Footer text</w:t>"));
    }

    #[test]
    fn test_document_builder_with_header_footer() {
        use crate::Document;
        use std::io::Cursor;

        // Create a document with header and footer
        let mut builder = DocumentBuilder::new();

        {
            let mut header = builder.add_header(HeaderFooterType::Default);
            header.add_paragraph("Document Header");
        }

        {
            let mut footer = builder.add_footer(HeaderFooterType::Default);
            footer.add_paragraph("Page Footer");
        }

        builder.add_paragraph("Body content");

        // Write to memory
        let mut buffer = Cursor::new(Vec::new());
        builder.write(&mut buffer).unwrap();

        // Read it back and verify
        buffer.set_position(0);
        let mut doc = Document::from_reader(buffer).unwrap();

        // Verify body content
        assert_eq!(doc.body().paragraphs().len(), 1);
        assert_eq!(doc.text(), "Body content");

        // Verify section properties have header/footer references
        let sect_pr = doc
            .body()
            .section_properties()
            .expect("should have section properties");
        assert_eq!(sect_pr.headers.len(), 1);
        assert_eq!(sect_pr.footers.len(), 1);

        // Get rel_ids before mutable borrows
        let header_rel_id = sect_pr.headers[0].rel_id.clone();
        let footer_rel_id = sect_pr.footers[0].rel_id.clone();

        // Verify we can read the header content
        let header = doc.get_header(&header_rel_id).unwrap();
        assert_eq!(header.text(), "Document Header");

        // Verify we can read the footer content
        let footer = doc.get_footer(&footer_rel_id).unwrap();
        assert_eq!(footer.text(), "Page Footer");
    }

    #[test]
    fn test_document_builder_with_footnote() {
        use crate::Document;
        use crate::document::FootnoteReference;
        use std::io::Cursor;

        // Create a document with a footnote
        let mut builder = DocumentBuilder::new();

        // Add a footnote first
        let footnote_id = {
            let mut footnote = builder.add_footnote();
            footnote.add_paragraph("This is a footnote.");
            footnote.id()
        };

        // Add body content with footnote reference
        {
            let para = builder.body_mut().add_paragraph();
            let run = para.add_run();
            run.set_text("Some text");
            run.set_footnote_ref(FootnoteReference { id: footnote_id });
        }

        // Write to memory
        let mut buffer = Cursor::new(Vec::new());
        builder.write(&mut buffer).unwrap();

        // Read it back and verify
        buffer.set_position(0);
        let mut doc = Document::from_reader(buffer).unwrap();

        // Verify body content has footnote reference
        let para = &doc.body().paragraphs()[0];
        let run = &para.runs()[0];
        assert!(run.footnote_ref().is_some());
        assert_eq!(run.footnote_ref().unwrap().id, footnote_id);

        // Verify we can read the footnotes part
        let footnotes = doc.get_footnotes().unwrap();

        // Find our footnote (ID 1) - cast to i32 for the get() method
        let fn1 = footnotes
            .get(footnote_id as i32)
            .expect("should find footnote");
        assert_eq!(fn1.text(), "This is a footnote.");
    }

    #[test]
    fn test_document_builder_with_endnote() {
        use crate::Document;
        use crate::document::EndnoteReference;
        use std::io::Cursor;

        // Create a document with an endnote
        let mut builder = DocumentBuilder::new();

        // Add an endnote first
        let endnote_id = {
            let mut endnote = builder.add_endnote();
            endnote.add_paragraph("This is an endnote.");
            endnote.id()
        };

        // Add body content with endnote reference
        {
            let para = builder.body_mut().add_paragraph();
            let run = para.add_run();
            run.set_text("Some text");
            run.set_endnote_ref(EndnoteReference { id: endnote_id });
        }

        // Write to memory
        let mut buffer = Cursor::new(Vec::new());
        builder.write(&mut buffer).unwrap();

        // Read it back and verify
        buffer.set_position(0);
        let mut doc = Document::from_reader(buffer).unwrap();

        // Verify body content has endnote reference
        let para = &doc.body().paragraphs()[0];
        let run = &para.runs()[0];
        assert!(run.endnote_ref().is_some());
        assert_eq!(run.endnote_ref().unwrap().id, endnote_id);

        // Verify we can read the endnotes part
        let endnotes = doc.get_endnotes().unwrap();

        // Find our endnote (ID 1) - cast to i32 for the get() method
        let en1 = endnotes
            .get(endnote_id as i32)
            .expect("should find endnote");
        assert_eq!(en1.text(), "This is an endnote.");
    }

    #[test]
    fn test_document_builder_with_comment() {
        use crate::Document;
        use crate::document::CommentReference;
        use std::io::Cursor;

        // Create a document with a comment
        let mut builder = DocumentBuilder::new();

        // Add a comment first
        let comment_id = {
            let mut comment = builder.add_comment();
            comment.set_author("Test Author");
            comment.add_paragraph("This needs review.");
            comment.id()
        };

        // Add body content with comment ranges and reference
        {
            let para = builder.body_mut().add_paragraph();
            para.add_comment_range_start(comment_id);
            para.add_run().set_text("Commented text");
            para.add_comment_range_end(comment_id);
            para.add_run()
                .set_comment_ref(CommentReference { id: comment_id });
        }

        // Write to memory
        let mut buffer = Cursor::new(Vec::new());
        builder.write(&mut buffer).unwrap();

        // Read it back and verify
        buffer.set_position(0);
        let mut doc = Document::from_reader(buffer).unwrap();

        // Verify body content
        assert!(!doc.body().paragraphs().is_empty());

        // Verify we can read the comments part
        let comments = doc.get_comments().unwrap();

        // Find our comment (ID 0)
        let c0 = comments
            .get(comment_id as i32)
            .expect("should find comment");
        assert_eq!(c0.author.as_deref(), Some("Test Author"));
        assert_eq!(c0.text(), "This needs review.");
    }
}
