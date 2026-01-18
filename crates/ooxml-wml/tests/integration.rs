//! Integration tests for ooxml-wml.
//!
//! Tests document reading, writing, and roundtripping.

use ooxml_wml::{Document, DocumentBuilder, Drawing};
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

/// Test creating a document with an inline image.
#[test]
fn test_roundtrip_image() {
    // Create a simple PNG image (1x1 pixel red)
    let png_data: Vec<u8> = vec![
        0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A, // PNG signature
        0x00, 0x00, 0x00, 0x0D, 0x49, 0x48, 0x44, 0x52, // IHDR chunk
        0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x01, // 1x1 pixels
        0x08, 0x02, 0x00, 0x00, 0x00, 0x90, 0x77, 0x53, 0xDE, // 8-bit RGB
        0x00, 0x00, 0x00, 0x0C, 0x49, 0x44, 0x41, 0x54, // IDAT chunk
        0x08, 0xD7, 0x63, 0xF8, 0xCF, 0xC0, 0x00, 0x00, // compressed data
        0x00, 0x03, 0x00, 0x01, 0x00, 0x18, 0xDD, 0x8D, 0xB4, 0x00, 0x00, 0x00, 0x00, 0x49, 0x45,
        0x4E, 0x44, // IEND chunk
        0xAE, 0x42, 0x60, 0x82,
    ];

    // Create a document with an image
    let mut builder = DocumentBuilder::new();
    let rel_id = builder.add_image(png_data.clone(), "image/png");

    {
        let para = builder.body_mut().add_paragraph();
        let run = para.add_run();

        // Add a drawing with the image
        let mut drawing = Drawing::new();
        drawing
            .add_image(&rel_id)
            .set_width_inches(2.0)
            .set_height_inches(1.5)
            .set_description("Test image");
        run.drawings_mut().push(drawing);
    }

    // Write to memory
    let mut buffer = Cursor::new(Vec::new());
    builder.write(&mut buffer).unwrap();

    // Read it back
    buffer.set_position(0);
    let mut doc = Document::from_reader(buffer).unwrap();

    // Verify the image was created
    assert!(doc.package().has_part("word/media/image1.png"));

    // Verify the drawing structure and get rel_id for later use
    let image_rel_id;
    {
        let para = &doc.body().paragraphs()[0];
        let run = &para.runs()[0];
        assert!(run.has_images());
        assert_eq!(run.drawings().len(), 1);

        let drawing = &run.drawings()[0];
        assert_eq!(drawing.images().len(), 1);

        let image = &drawing.images()[0];
        assert_eq!(image.rel_id(), &rel_id);
        assert!((image.width_inches().unwrap() - 2.0).abs() < 0.001);
        assert!((image.height_inches().unwrap() - 1.5).abs() < 0.001);
        assert_eq!(image.description(), Some("Test image"));

        image_rel_id = image.rel_id().to_string();
    }

    // Verify we can load the image data
    let image_data = doc.get_image_data(&image_rel_id).unwrap();
    assert_eq!(image_data.content_type, "image/png");
    assert_eq!(image_data.data, png_data);
}

/// Test creating a document with multiple images.
#[test]
fn test_multiple_images() {
    // Simple test data
    let png_data = vec![0x89, 0x50, 0x4E, 0x47, 0x00, 0x00];
    let jpg_data = vec![0xFF, 0xD8, 0xFF, 0xE0, 0x00, 0x00];

    let mut builder = DocumentBuilder::new();
    let rel_id1 = builder.add_image(png_data.clone(), "image/png");
    let rel_id2 = builder.add_image(jpg_data.clone(), "image/jpeg");

    // First paragraph with first image
    {
        let para = builder.body_mut().add_paragraph();
        let run = para.add_run();
        let mut drawing = Drawing::new();
        drawing.add_image(&rel_id1).set_width_inches(1.0);
        run.drawings_mut().push(drawing);
    }

    // Second paragraph with second image
    {
        let para = builder.body_mut().add_paragraph();
        let run = para.add_run();
        let mut drawing = Drawing::new();
        drawing.add_image(&rel_id2).set_width_inches(3.0);
        run.drawings_mut().push(drawing);
    }

    // Write and read back
    let mut buffer = Cursor::new(Vec::new());
    builder.write(&mut buffer).unwrap();

    buffer.set_position(0);
    let mut doc = Document::from_reader(buffer).unwrap();

    // Verify both images exist
    assert!(doc.package().has_part("word/media/image1.png"));
    assert!(doc.package().has_part("word/media/image2.jpg"));

    // Verify both paragraphs have images and collect rel_ids
    let (img1_rel_id, img2_rel_id);
    {
        let paras = doc.body().paragraphs();
        assert_eq!(paras.len(), 2);
        assert!(paras[0].runs()[0].has_images());
        assert!(paras[1].runs()[0].has_images());

        img1_rel_id = paras[0].runs()[0].drawings()[0].images()[0]
            .rel_id()
            .to_string();
        img2_rel_id = paras[1].runs()[0].drawings()[0].images()[0]
            .rel_id()
            .to_string();
    }

    // Load first image
    let data1 = doc.get_image_data(&img1_rel_id).unwrap();
    assert_eq!(data1.content_type, "image/png");
    assert_eq!(data1.data, png_data);

    // Load second image
    let data2 = doc.get_image_data(&img2_rel_id).unwrap();
    assert_eq!(data2.content_type, "image/jpeg");
    assert_eq!(data2.data, jpg_data);
}
