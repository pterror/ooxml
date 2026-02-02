//! Integration tests for ooxml-wml.
//!
//! Tests document reading, writing, and roundtripping.
//! Reading uses generated types accessed via extension traits.

use ooxml_wml::ext::{BodyExt, HyperlinkExt, ParagraphExt, RunExt, RunPropertiesExt};
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
    use ooxml_wml::ext::DrawingExt;

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
        let rel_ids = drawing.all_image_rel_ids();
        assert_eq!(rel_ids.len(), 1);
        image_rel_id = rel_ids[0].to_string();
        assert_eq!(image_rel_id, rel_id);
    }

    // Verify we can load the image data
    let image_data = doc.get_image_data(&image_rel_id).unwrap();
    assert_eq!(image_data.content_type, "image/png");
    assert_eq!(image_data.data, png_data);
}

/// Test creating a document with multiple images.
#[test]
fn test_multiple_images() {
    use ooxml_wml::ext::DrawingExt;

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

        img1_rel_id = paras[0].runs()[0].drawings()[0].all_image_rel_ids()[0].to_string();
        img2_rel_id = paras[1].runs()[0].drawings()[0].all_image_rel_ids()[0].to_string();
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

/// Test creating a document with a hyperlink.
#[test]
fn test_roundtrip_hyperlink() {
    let mut builder = DocumentBuilder::new();
    let rel_id = builder.add_hyperlink("https://example.com");

    {
        let para = builder.body_mut().add_paragraph();
        para.add_run().set_text("Visit ");

        let link = para.add_hyperlink();
        link.set_rel_id(&rel_id);
        link.add_run().set_text("our website");

        para.add_run().set_text(" for more info.");
    }

    // Write to memory
    let mut buffer = Cursor::new(Vec::new());
    builder.write(&mut buffer).unwrap();

    // Read it back
    buffer.set_position(0);
    let doc = Document::from_reader(buffer).unwrap();

    // Verify paragraph has hyperlink
    let para = &doc.body().paragraphs()[0];
    let links = para.hyperlinks();
    assert_eq!(links.len(), 1);
    let link = links[0];
    assert!(link.is_external());
    assert_eq!(link.text(), "our website");

    // Look up the URL
    let url = doc.get_hyperlink_url(link.rel_id().unwrap()).unwrap();
    assert_eq!(url, "https://example.com");

    // Full text
    assert_eq!(para.text(), "Visit our website for more info.");
}

/// Test creating a document with a numbered list.
#[test]
fn test_roundtrip_numbered_list() {
    use ooxml_wml::{ListType, NumberingProperties, ParagraphProperties};

    let mut builder = DocumentBuilder::new();
    let num_id = builder.add_list(ListType::Decimal);

    // Add three list items
    for item in &["First item", "Second item", "Third item"] {
        let para = builder.body_mut().add_paragraph();
        para.set_properties(ParagraphProperties {
            numbering: Some(NumberingProperties { num_id, ilvl: 0 }),
            ..Default::default()
        });
        para.add_run().set_text(*item);
    }

    // Write to memory
    let mut buffer = Cursor::new(Vec::new());
    builder.write(&mut buffer).unwrap();

    // Read it back
    buffer.set_position(0);
    let doc = Document::from_reader(buffer).unwrap();

    // Verify numbering.xml was created
    assert!(doc.package().has_part("word/numbering.xml"));

    // Verify paragraphs exist with correct text
    let paras = doc.body().paragraphs();
    assert_eq!(paras.len(), 3);

    // Verify text content
    assert_eq!(paras[0].text(), "First item");
    assert_eq!(paras[1].text(), "Second item");
    assert_eq!(paras[2].text(), "Third item");

    // Note: paragraph numbering properties (numPr) are in generated ParagraphProperties
    // extra_children since codegen doesn't yet flatten CTPPrBase into ParagraphProperties.
    // Full property verification is deferred to Phase 7 (writer migration).
}

/// Test creating a document with a bullet list.
#[test]
fn test_roundtrip_bullet_list() {
    use ooxml_wml::{ListType, NumberingProperties, ParagraphProperties};

    let mut builder = DocumentBuilder::new();
    let num_id = builder.add_list(ListType::Bullet);

    // Add bullet items
    for item in &["Apple", "Banana", "Cherry"] {
        let para = builder.body_mut().add_paragraph();
        para.set_properties(ParagraphProperties {
            numbering: Some(NumberingProperties { num_id, ilvl: 0 }),
            ..Default::default()
        });
        para.add_run().set_text(*item);
    }

    // Write to memory
    let mut buffer = Cursor::new(Vec::new());
    builder.write(&mut buffer).unwrap();

    // Read it back
    buffer.set_position(0);
    let doc = Document::from_reader(buffer).unwrap();

    // Verify content
    assert_eq!(doc.body().paragraphs().len(), 3);
    assert_eq!(doc.text(), "Apple\nBanana\nCherry");
}

/// Test creating a document with page breaks.
#[test]
fn test_roundtrip_page_break() {
    let mut builder = DocumentBuilder::new();

    // First page content
    builder.add_paragraph("Page 1 content");

    // Page break
    {
        let para = builder.body_mut().add_paragraph();
        para.add_run().set_page_break(true);
    }

    // Second page content
    builder.add_paragraph("Page 2 content");

    // Write to memory
    let mut buffer = Cursor::new(Vec::new());
    builder.write(&mut buffer).unwrap();

    // Read it back
    buffer.set_position(0);
    let doc = Document::from_reader(buffer).unwrap();

    // Verify structure
    assert_eq!(doc.body().paragraphs().len(), 3);

    // First paragraph - no page break
    assert!(!doc.body().paragraphs()[0].runs()[0].has_page_break());
    assert_eq!(doc.body().paragraphs()[0].text(), "Page 1 content");

    // Second paragraph - has page break
    assert!(doc.body().paragraphs()[1].runs()[0].has_page_break());

    // Third paragraph - no page break
    assert!(!doc.body().paragraphs()[2].runs()[0].has_page_break());
    assert_eq!(doc.body().paragraphs()[2].text(), "Page 2 content");
}

/// Test creating a document with colored text.
#[test]
fn test_roundtrip_text_color() {
    use ooxml_wml::RunProperties;

    let mut builder = DocumentBuilder::new();

    {
        let para = builder.body_mut().add_paragraph();

        // Red text
        let run = para.add_run();
        run.set_text("Red ");
        run.set_properties(RunProperties {
            color: Some("FF0000".to_string()),
            ..Default::default()
        });

        // Blue text
        let run = para.add_run();
        run.set_text("Blue ");
        run.set_properties(RunProperties {
            color: Some("0000FF".to_string()),
            ..Default::default()
        });

        // Green text (bold too)
        let run = para.add_run();
        run.set_text("Green");
        run.set_properties(RunProperties {
            color: Some("00FF00".to_string()),
            bold: true,
            ..Default::default()
        });
    }

    // Write to memory
    let mut buffer = Cursor::new(Vec::new());
    builder.write(&mut buffer).unwrap();

    // Read it back
    buffer.set_position(0);
    let doc = Document::from_reader(buffer).unwrap();

    // Verify colors
    let para = &doc.body().paragraphs()[0];
    assert_eq!(para.runs().len(), 3);

    let run0 = &para.runs()[0];
    assert_eq!(run0.text(), "Red ");
    assert_eq!(
        run0.properties().and_then(|p| p.color_hex()),
        Some("FF0000")
    );

    let run1 = &para.runs()[1];
    assert_eq!(run1.text(), "Blue ");
    assert_eq!(
        run1.properties().and_then(|p| p.color_hex()),
        Some("0000FF")
    );

    let run2 = &para.runs()[2];
    assert_eq!(run2.text(), "Green");
    assert_eq!(
        run2.properties().and_then(|p| p.color_hex()),
        Some("00FF00")
    );
    assert!(run2.is_bold());
}

/// Test creating a document with paragraph alignment, spacing, and indentation.
///
/// Note: The generated ParagraphProperties doesn't yet flatten CTPPrBase fields
/// (alignment, spacing, indent). These are captured as extra_children raw XML.
/// We verify the text content round-trips and the paragraph properties exist.
#[test]
fn test_roundtrip_paragraph_properties() {
    use ooxml_wml::{Alignment, ParagraphProperties};

    let mut builder = DocumentBuilder::new();

    // Centered paragraph
    {
        let para = builder.body_mut().add_paragraph();
        para.set_properties(ParagraphProperties {
            alignment: Some(Alignment::Center),
            ..Default::default()
        });
        para.add_run().set_text("Centered text");
    }

    // Right-aligned with spacing
    {
        let para = builder.body_mut().add_paragraph();
        para.set_properties(ParagraphProperties {
            alignment: Some(Alignment::Right),
            spacing_before: Some(240),
            spacing_after: Some(120),
            ..Default::default()
        });
        para.add_run().set_text("Right aligned with spacing");
    }

    // Indented paragraph
    {
        let para = builder.body_mut().add_paragraph();
        para.set_properties(ParagraphProperties {
            indent_left: Some(720),
            indent_first_line: Some(360),
            ..Default::default()
        });
        para.add_run().set_text("Indented paragraph");
    }

    // Write to memory
    let mut buffer = Cursor::new(Vec::new());
    builder.write(&mut buffer).unwrap();

    // Read it back
    buffer.set_position(0);
    let doc = Document::from_reader(buffer).unwrap();

    // Verify text content round-trips correctly
    let paras = doc.body().paragraphs();
    assert_eq!(paras.len(), 3);
    assert_eq!(paras[0].text(), "Centered text");
    assert_eq!(paras[1].text(), "Right aligned with spacing");
    assert_eq!(paras[2].text(), "Indented paragraph");

    // Verify paragraph properties are present (captured as extra_children)
    for para in &paras {
        assert!(
            para.properties().is_some(),
            "paragraph should have properties"
        );
    }
}

/// Test that unknown XML elements survive a roundtrip.
///
/// This verifies the core round-trip preservation functionality:
/// elements we don't explicitly understand are captured during parse
/// and written back during serialize.
#[test]
fn test_roundtrip_unknown_elements() {
    use ooxml_wml::{PositionedNode, RawXmlElement, RawXmlNode, RunProperties};

    // Create a document with unknown elements added programmatically
    let mut builder = DocumentBuilder::new();

    {
        let para = builder.body_mut().add_paragraph();

        // Add a run with text and unknown children in rPr
        let run = para.add_run();
        run.set_text("Hello with custom props");

        // Create RunProperties with unknown children
        let mut rpr = RunProperties {
            bold: true,
            ..Default::default()
        };

        // Add a fake unknown element to rPr
        let unknown_elem = RawXmlElement {
            name: "w:customTracking".to_string(),
            attributes: vec![("w:val".to_string(), "strict".to_string())],
            children: vec![],
            self_closing: true,
        };
        rpr.unknown_children
            .push(PositionedNode::new(0, RawXmlNode::Element(unknown_elem)));
        run.set_properties(rpr);
    }

    // Write to memory
    let mut buffer = Cursor::new(Vec::new());
    builder.write(&mut buffer).unwrap();

    // Read the raw XML to verify unknown element is present
    buffer.set_position(0);
    let doc = Document::from_reader(buffer).unwrap();

    // Verify the known content
    let para = &doc.body().paragraphs()[0];
    let run = &para.runs()[0];
    assert_eq!(run.text(), "Hello with custom props");
    assert!(run.is_bold());

    // Verify the unknown element was preserved in the generated RunProperties extra_children
    let rpr = run.properties().unwrap();
    assert!(
        !rpr.extra_children.is_empty(),
        "unknown children should be captured"
    );

    // Find the customTracking element in extra_children
    let found = rpr.extra_children.iter().any(|node| {
        if let RawXmlNode::Element(elem) = node {
            elem.name.ends_with("customTracking")
        } else {
            false
        }
    });
    assert!(found, "w:customTracking should be preserved");
}

/// Test roundtrip of extended font attributes (w:rFonts).
#[test]
fn test_roundtrip_fonts() {
    use ooxml_wml::{Fonts, RunProperties};

    let mut builder = DocumentBuilder::new();
    {
        let para = builder.body_mut().add_paragraph();
        let run = para.add_run();
        run.set_text("Text with fonts");
        run.set_properties(RunProperties {
            fonts: Some(Fonts {
                ascii: Some("Arial".to_string()),
                h_ansi: Some("Arial".to_string()),
                east_asia: Some("MS Gothic".to_string()),
                cs: Some("Arial".to_string()),
            }),
            ..Default::default()
        });
    }

    // Write to memory
    let mut buffer = Cursor::new(Vec::new());
    builder.write(&mut buffer).unwrap();

    // Read it back
    buffer.set_position(0);
    let doc = Document::from_reader(buffer).unwrap();

    // Verify the fonts were preserved
    let para = &doc.body().paragraphs()[0];
    let run = &para.runs()[0];
    let rpr = run.properties().unwrap();
    let fonts = rpr.fonts.as_ref().expect("fonts should be present");

    assert_eq!(fonts.ascii.as_deref(), Some("Arial"));
    assert_eq!(fonts.h_ansi.as_deref(), Some("Arial"));
    assert_eq!(fonts.east_asia.as_deref(), Some("MS Gothic"));
    assert_eq!(fonts.cs.as_deref(), Some("Arial"));
}

/// Test roundtrip of bookmarks.
#[test]
fn test_roundtrip_bookmarks() {
    use ooxml_wml::types::EGPContent;

    let mut builder = DocumentBuilder::new();
    {
        let para = builder.body_mut().add_paragraph();
        para.add_bookmark_start(0, "my_bookmark");
        para.add_run().set_text("Bookmarked text");
        para.add_bookmark_end(0);
    }

    // Write to memory
    let mut buffer = Cursor::new(Vec::new());
    builder.write(&mut buffer).unwrap();

    // Read it back
    buffer.set_position(0);
    let doc = Document::from_reader(buffer).unwrap();

    // Verify the bookmark was preserved
    let para = &doc.body().paragraphs()[0];

    // Check text
    assert_eq!(para.text(), "Bookmarked text");

    // Verify bookmark elements are present in the paragraph content
    let has_bookmark_start = para
        .p_content
        .iter()
        .any(|c| matches!(c.as_ref(), EGPContent::BookmarkStart(b) if b.name == "my_bookmark"));
    assert!(has_bookmark_start, "should have BookmarkStart");

    let has_bookmark_end = para
        .p_content
        .iter()
        .any(|c| matches!(c.as_ref(), EGPContent::BookmarkEnd(_)));
    assert!(has_bookmark_end, "should have BookmarkEnd");
}

/// Test that Document::write() produces a valid package that can be re-read.
///
/// Creates a document with DocumentBuilder, reads it via Document::from_reader(),
/// writes it back with Document::write(), then reads the result and verifies
/// text content is preserved.
#[test]
fn test_document_save_roundtrip() {
    let mut builder = DocumentBuilder::new();
    builder.add_paragraph("Hello, roundtrip!");
    builder.add_paragraph("Second paragraph.");

    // Write with builder
    let mut buffer = Cursor::new(Vec::new());
    builder.write(&mut buffer).unwrap();

    // Read with generated parser
    buffer.set_position(0);
    let mut doc = Document::from_reader(buffer).unwrap();

    // Save via generated serializer
    let mut out = Cursor::new(Vec::new());
    doc.write(&mut out).unwrap();

    // Re-read the saved output
    out.set_position(0);
    let doc2 = Document::from_reader(out).unwrap();

    assert_eq!(doc2.text(), "Hello, roundtrip!\nSecond paragraph.");
}

/// Test that Document::write() preserves all package parts (images, rels, etc.).
#[test]
fn test_document_save_preserves_parts() {
    let mut builder = DocumentBuilder::new();
    let png_data = vec![0x89, 0x50, 0x4E, 0x47, 0x00, 0x00];
    let rel_id = builder.add_image(png_data, "image/png");

    {
        let para = builder.body_mut().add_paragraph();
        let run = para.add_run();
        let mut drawing = Drawing::new();
        drawing.add_image(&rel_id).set_width_inches(1.0);
        run.drawings_mut().push(drawing);
    }
    builder.add_paragraph("Text content");

    let mut buffer = Cursor::new(Vec::new());
    builder.write(&mut buffer).unwrap();

    // Read and save
    buffer.set_position(0);
    let mut doc = Document::from_reader(buffer).unwrap();

    let mut out = Cursor::new(Vec::new());
    doc.write(&mut out).unwrap();

    // Verify all parts are present in the saved output
    out.set_position(0);
    let doc2 = Document::from_reader(out).unwrap();

    assert!(doc2.package().has_part("word/document.xml"));
    assert!(doc2.package().has_part("_rels/.rels"));
    assert!(doc2.package().has_part("word/media/image1.png"));
    assert!(doc2.text().contains("Text content"));
}

/// Test that modifications to the generated body are reflected after save.
#[test]
fn test_document_save_with_body_modification() {
    use ooxml_wml::types::{EGBlockLevelElts, EGPContent, EGRunInnerContent, Paragraph, Run, Text};

    let mut builder = DocumentBuilder::new();
    builder.add_paragraph("Original text");

    let mut buffer = Cursor::new(Vec::new());
    builder.write(&mut buffer).unwrap();

    // Read the document
    buffer.set_position(0);
    let mut doc = Document::from_reader(buffer).unwrap();

    // Add a new paragraph via the generated types
    let text = Text {
        text: Some("Added paragraph".to_string()),
        ..Default::default()
    };
    let mut run = Run::default();
    run.run_inner_content
        .push(Box::new(EGRunInnerContent::T(Box::new(text))));
    let mut para = Paragraph::default();
    para.p_content.push(Box::new(EGPContent::R(Box::new(run))));
    doc.body_mut()
        .block_level_elts
        .push(Box::new(EGBlockLevelElts::P(Box::new(para))));

    // Save
    let mut out = Cursor::new(Vec::new());
    doc.write(&mut out).unwrap();

    // Re-read and verify the modification
    out.set_position(0);
    let doc2 = Document::from_reader(out).unwrap();

    assert_eq!(doc2.text(), "Original text\nAdded paragraph");
}
