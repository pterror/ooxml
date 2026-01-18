//! Document writing and serialization.
//!
//! This module provides functionality for creating new Word documents
//! and saving existing documents.

use crate::document::{Body, Paragraph, ParagraphProperties, Run, RunProperties};
use crate::error::Result;
use crate::styles::Styles;
use ooxml::{PackageWriter, Relationship, Relationships, content_type, rel_type};
use std::fs::File;
use std::io::{BufWriter, Seek, Write};
use std::path::Path;

/// WordprocessingML namespace.
pub const NS_W: &str = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
/// Relationships namespace.
pub const NS_R: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

/// Builder for creating new Word documents.
pub struct DocumentBuilder {
    body: Body,
    _styles: Styles, // TODO: serialize styles.xml
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
        }
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

        // Write document relationships (even if empty, Word expects it)
        let doc_rels = Relationships::new();
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

    // Document element with namespace
    xml.push_str(&format!(
        r#"<w:document xmlns:w="{}" xmlns:r="{}">"#,
        NS_W, NS_R
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
    for para in body.paragraphs() {
        serialize_paragraph(para, xml);
    }
}

/// Serialize a paragraph.
fn serialize_paragraph(para: &Paragraph, xml: &mut String) {
    xml.push_str("<w:p>");

    // Paragraph properties
    if let Some(props) = para.properties() {
        serialize_paragraph_properties(props, xml);
    }

    // Runs
    for run in para.runs() {
        serialize_run(run, xml);
    }

    xml.push_str("</w:p>");
}

/// Serialize paragraph properties.
fn serialize_paragraph_properties(props: &ParagraphProperties, xml: &mut String) {
    xml.push_str("<w:pPr>");

    if let Some(ref style) = props.style {
        xml.push_str(&format!(r#"<w:pStyle w:val="{}"/>"#, escape_xml(style)));
    }

    xml.push_str("</w:pPr>");
}

/// Serialize a run.
fn serialize_run(run: &Run, xml: &mut String) {
    xml.push_str("<w:r>");

    // Run properties
    if let Some(props) = run.properties() {
        serialize_run_properties(props, xml);
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

    xml.push_str("</w:r>");
}

/// Serialize run properties.
fn serialize_run_properties(props: &RunProperties, xml: &mut String) {
    // Only output if there are properties to write
    let has_props = props.bold
        || props.italic
        || props.underline
        || props.strike
        || props.size.is_some()
        || props.font.is_some()
        || props.style.is_some();

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

    if props.underline {
        xml.push_str(r#"<w:u w:val="single"/>"#);
    }

    if props.strike {
        xml.push_str("<w:strike/>");
    }

    if let Some(size) = props.size {
        xml.push_str(&format!(r#"<w:sz w:val="{}"/>"#, size));
    }

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
}
