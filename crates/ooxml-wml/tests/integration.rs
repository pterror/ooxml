//! Integration tests for ooxml-wml.
//!
//! Tests document reading, writing, and roundtripping.

use ooxml_wml::{Document, DocumentBuilder};
use std::io::Cursor;

/// Test creating a document and reading it back.
#[test]
fn test_roundtrip_simple_document() {
    // Create a document
    let mut builder = DocumentBuilder::new();
    builder.add_paragraph("Hello, World!");
    builder.add_paragraph("This is a test document.");

    // Write to memory
    let mut buffer = Cursor::new(Vec::new());
    builder.write(&mut buffer).unwrap();

    // Read it back
    buffer.set_position(0);
    let doc = Document::from_reader(buffer).unwrap();

    // Verify content
    assert_eq!(doc.body().paragraphs().len(), 2);
    assert_eq!(doc.text(), "Hello, World!\nThis is a test document.");
}

/// Test document with formatted text.
#[test]
fn test_roundtrip_formatted_text() {
    use ooxml_wml::RunProperties;

    // Create a document with formatting
    let mut builder = DocumentBuilder::new();
    {
        let para = builder.body_mut().add_paragraph();
        let run = para.add_run();
        run.set_text("Bold and italic");
        run.set_properties(RunProperties {
            bold: true,
            italic: true,
            ..Default::default()
        });
    }

    // Write to memory
    let mut buffer = Cursor::new(Vec::new());
    builder.write(&mut buffer).unwrap();

    // Read it back
    buffer.set_position(0);
    let doc = Document::from_reader(buffer).unwrap();

    // Verify content
    let para = &doc.body().paragraphs()[0];
    let run = &para.runs()[0];
    assert_eq!(run.text(), "Bold and italic");
    assert!(run.is_bold());
    assert!(run.is_italic());
}

/// Test reading package relationships.
#[test]
fn test_read_package_structure() {
    // Create a minimal document
    let mut builder = DocumentBuilder::new();
    builder.add_paragraph("Test");

    let mut buffer = Cursor::new(Vec::new());
    builder.write(&mut buffer).unwrap();

    buffer.set_position(0);
    let doc = Document::from_reader(buffer).unwrap();

    // Verify package structure
    assert!(doc.package().has_part("word/document.xml"));
    assert!(doc.package().has_part("_rels/.rels"));
    assert!(doc.package().has_part("[Content_Types].xml"));
}

/// Test document text extraction.
#[test]
fn test_text_extraction() {
    let mut builder = DocumentBuilder::new();
    builder.add_paragraph("Line 1");
    builder.add_paragraph("Line 2");
    builder.add_paragraph("Line 3");

    let mut buffer = Cursor::new(Vec::new());
    builder.write(&mut buffer).unwrap();

    buffer.set_position(0);
    let doc = Document::from_reader(buffer).unwrap();

    let text = doc.text();
    assert!(text.contains("Line 1"));
    assert!(text.contains("Line 2"));
    assert!(text.contains("Line 3"));
}

/// Test multiple runs in a paragraph.
#[test]
fn test_multiple_runs_roundtrip() {
    use ooxml_wml::RunProperties;

    let mut builder = DocumentBuilder::new();
    {
        let para = builder.body_mut().add_paragraph();

        // Add multiple runs with different formatting
        para.add_run().set_text("Normal ");

        let bold_run = para.add_run();
        bold_run.set_text("bold ");
        bold_run.set_properties(RunProperties {
            bold: true,
            ..Default::default()
        });

        para.add_run().set_text("normal again");
    }

    let mut buffer = Cursor::new(Vec::new());
    builder.write(&mut buffer).unwrap();

    buffer.set_position(0);
    let doc = Document::from_reader(buffer).unwrap();

    let para = &doc.body().paragraphs()[0];
    assert_eq!(para.runs().len(), 3);
    assert_eq!(para.text(), "Normal bold normal again");

    // Check formatting
    assert!(!para.runs()[0].is_bold());
    assert!(para.runs()[1].is_bold());
    assert!(!para.runs()[2].is_bold());
}
