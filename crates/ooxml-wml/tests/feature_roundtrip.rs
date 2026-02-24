//! Feature-specific roundtrip tests for ooxml-wml.
//!
//! Tests document features that are not covered by integration.rs or parity_tests.rs:
//! headers/footers, footnotes, endnotes, comments, track changes, anchored images,
//! merged table cells, and section properties.

use ooxml_wml::ext::{BodyExt, CellExt, ParagraphExt, RowExt, RunExt, TableExt};
use ooxml_wml::{Document, DocumentBuilder, Drawing};
use std::io::Cursor;

// =============================================================================
// Helpers
// =============================================================================

/// Build a document, write to memory, read back.
fn roundtrip(builder: DocumentBuilder) -> Document<Cursor<Vec<u8>>> {
    let mut buffer = Cursor::new(Vec::new());
    builder.write(&mut buffer).unwrap();
    buffer.set_position(0);
    Document::from_reader(buffer).unwrap()
}

// =============================================================================
// 1. Headers and footers
// =============================================================================

/// Test creating a document with a default header and verifying the content round-trips.
///
/// The writer writes header content to `word/header1.xml` and adds a relationship
/// from `word/document.xml.rels`. On reading back, `get_header(rel_id)` uses the
/// relationship to load and parse that part.
///
/// The section properties on the body contain `<w:headerReference>` entries; we walk
/// `body.sect_pr.header_footer_refs` to find the rel ID via the `id` field on
/// `HeaderFooterReference` (ECMA-376 §17.10.5).
#[cfg(feature = "wml-layout")]
#[test]
fn test_header_footer_roundtrip() {
    use ooxml_wml::HeaderFooterType;
    use ooxml_wml::types::HeaderFooterRef;

    let mut builder = DocumentBuilder::new();
    builder.add_paragraph("Body content");

    {
        let mut hdr = builder.add_header(HeaderFooterType::Default);
        hdr.add_paragraph("Page header text");
    }

    {
        let mut ftr = builder.add_footer(HeaderFooterType::Default);
        ftr.add_paragraph("Page footer text");
    }

    let mut doc = roundtrip(builder);

    // Collect rel IDs from section properties before making mutable borrows.
    // HeaderFooterReference.id holds the r:id relationship ID directly.
    let (hdr_rel_id, ftr_rel_id) = {
        let sect = doc
            .body()
            .section_properties()
            .expect("section properties should exist");

        assert!(
            !sect.header_footer_refs.is_empty(),
            "section properties should have header/footer refs"
        );

        let mut hdr_id = None;
        let mut ftr_id = None;
        for r in &sect.header_footer_refs {
            match r {
                HeaderFooterRef::HeaderReference(h) => hdr_id = Some(h.id.clone()),
                HeaderFooterRef::FooterReference(f) => ftr_id = Some(f.id.clone()),
            }
        }

        (
            hdr_id.expect("should have a header reference"),
            ftr_id.expect("should have a footer reference"),
        )
    };

    // Load and verify header
    let header = doc.get_header(&hdr_rel_id).expect("should load header");
    let header_text: Vec<String> = header
        .block_content
        .iter()
        .filter_map(|bc| match bc {
            ooxml_wml::types::BlockContent::P(p) => Some(p.text()),
            _ => None,
        })
        .collect();
    assert!(
        header_text.iter().any(|t| t.contains("Page header text")),
        "header should contain expected text, got: {:?}",
        header_text
    );

    // Load and verify footer
    let footer = doc.get_footer(&ftr_rel_id).expect("should load footer");
    let footer_text: Vec<String> = footer
        .block_content
        .iter()
        .filter_map(|bc| match bc {
            ooxml_wml::types::BlockContent::P(p) => Some(p.text()),
            _ => None,
        })
        .collect();
    assert!(
        footer_text.iter().any(|t| t.contains("Page footer text")),
        "footer should contain expected text, got: {:?}",
        footer_text
    );
}

/// Test that even/first-page headers produce separate header references.
#[cfg(feature = "wml-layout")]
#[test]
fn test_multiple_header_types() {
    use ooxml_wml::HeaderFooterType;
    use ooxml_wml::types::HeaderFooterRef;

    let mut builder = DocumentBuilder::new();
    builder.add_paragraph("Content");

    builder
        .add_header(HeaderFooterType::Default)
        .add_paragraph("Default header");
    builder
        .add_header(HeaderFooterType::First)
        .add_paragraph("First page header");

    let doc = roundtrip(builder);

    let sect = doc
        .body()
        .section_properties()
        .expect("section properties should exist");

    // Count header references directly from header_footer_refs
    let hdr_count = sect
        .header_footer_refs
        .iter()
        .filter(|r| matches!(r, HeaderFooterRef::HeaderReference(_)))
        .count();
    assert_eq!(hdr_count, 2, "should have two header references");
}

// =============================================================================
// 2. Footnotes
// =============================================================================

/// Test adding a footnote, referencing it in a run, then reading back via get_footnotes().
#[test]
fn test_footnote_roundtrip() {
    let mut builder = DocumentBuilder::new();

    // Create a footnote
    let footnote_id;
    {
        let mut fn_builder = builder.add_footnote();
        fn_builder.add_paragraph("This is a footnote.");
        footnote_id = fn_builder.id();
    }

    // Add a paragraph with a footnote reference
    {
        let para = builder.body_mut().add_paragraph();
        let run = para.add_run();
        run.set_text("Text with footnote");
        let ref_run = para.add_run();
        ref_run.add_footnote_ref(footnote_id as i64);
    }

    let mut doc = roundtrip(builder);

    // Verify footnotes part is present in package
    assert!(
        doc.package().has_part("word/footnotes.xml"),
        "word/footnotes.xml should exist"
    );

    let footnotes = doc.get_footnotes().expect("should load footnotes");

    // The writer adds a separator footnote (id=-1 or id=0) plus our user footnote.
    // User footnotes have id >= 1.
    let user_footnotes: Vec<_> = footnotes.footnote.iter().filter(|f| f.id >= 1).collect();

    assert_eq!(
        user_footnotes.len(),
        1,
        "should have exactly one user footnote"
    );

    let fn_text: Vec<String> = user_footnotes[0]
        .block_content
        .iter()
        .filter_map(|bc| match bc {
            ooxml_wml::types::BlockContent::P(p) => Some(p.text()),
            _ => None,
        })
        .collect();
    assert!(
        fn_text.iter().any(|t| t.contains("This is a footnote.")),
        "footnote should contain expected text, got: {:?}",
        fn_text
    );
}

/// Test adding multiple footnotes.
#[test]
fn test_multiple_footnotes() {
    let mut builder = DocumentBuilder::new();

    let id1;
    let id2;
    {
        let mut fb = builder.add_footnote();
        fb.add_paragraph("First footnote");
        id1 = fb.id();
    }
    {
        let mut fb = builder.add_footnote();
        fb.add_paragraph("Second footnote");
        id2 = fb.id();
    }

    {
        let para = builder.body_mut().add_paragraph();
        para.add_run().set_text("First reference");
        para.add_run().add_footnote_ref(id1 as i64);
        para.add_run().set_text(" and second reference");
        para.add_run().add_footnote_ref(id2 as i64);
    }

    let mut doc = roundtrip(builder);
    let footnotes = doc.get_footnotes().expect("should load footnotes");

    let user_footnotes: Vec<_> = footnotes.footnote.iter().filter(|f| f.id >= 1).collect();
    assert_eq!(user_footnotes.len(), 2, "should have two user footnotes");
}

// =============================================================================
// 3. Endnotes
// =============================================================================

/// Test adding an endnote, referencing it in a run, then reading back via get_endnotes().
#[test]
fn test_endnote_roundtrip() {
    let mut builder = DocumentBuilder::new();

    let endnote_id;
    {
        let mut en_builder = builder.add_endnote();
        en_builder.add_paragraph("This is an endnote.");
        endnote_id = en_builder.id();
    }

    {
        let para = builder.body_mut().add_paragraph();
        let run = para.add_run();
        run.set_text("Text with endnote");
        let ref_run = para.add_run();
        ref_run.add_endnote_ref(endnote_id as i64);
    }

    let mut doc = roundtrip(builder);

    assert!(
        doc.package().has_part("word/endnotes.xml"),
        "word/endnotes.xml should exist"
    );

    let endnotes = doc.get_endnotes().expect("should load endnotes");

    let user_endnotes: Vec<_> = endnotes.endnote.iter().filter(|e| e.id >= 1).collect();
    assert_eq!(
        user_endnotes.len(),
        1,
        "should have exactly one user endnote"
    );

    let en_text: Vec<String> = user_endnotes[0]
        .block_content
        .iter()
        .filter_map(|bc| match bc {
            ooxml_wml::types::BlockContent::P(p) => Some(p.text()),
            _ => None,
        })
        .collect();
    assert!(
        en_text.iter().any(|t| t.contains("This is an endnote.")),
        "endnote should contain expected text, got: {:?}",
        en_text
    );
}

// =============================================================================
// 4. Comments
// =============================================================================

/// Test adding a comment with author/date and reading it back via get_comments().
#[test]
fn test_comment_roundtrip() {
    let mut builder = DocumentBuilder::new();

    let comment_id;
    {
        let mut cb = builder.add_comment();
        cb.set_author("Alice Reviewer");
        cb.set_date("2026-01-15T10:30:00Z");
        cb.add_paragraph("This is a comment.");
        comment_id = cb.id();
    }

    {
        let para = builder.body_mut().add_paragraph();
        para.add_comment_range_start(comment_id);
        para.add_run().set_text("Commented text");
        para.add_comment_range_end(comment_id);
        let ref_run = para.add_run();
        ref_run.add_comment_ref(comment_id as i64);
    }

    let mut doc = roundtrip(builder);

    assert!(
        doc.package().has_part("word/comments.xml"),
        "word/comments.xml should exist"
    );

    let comments = doc.get_comments().expect("should load comments");

    assert_eq!(comments.comment.len(), 1, "should have exactly one comment");

    let comment = &comments.comment[0];
    assert_eq!(
        comment.author, "Alice Reviewer",
        "comment author should match"
    );

    let comment_text: Vec<String> = comment
        .block_content
        .iter()
        .filter_map(|bc| match bc {
            ooxml_wml::types::BlockContent::P(p) => Some(p.text()),
            _ => None,
        })
        .collect();
    assert!(
        comment_text
            .iter()
            .any(|t| t.contains("This is a comment.")),
        "comment should contain expected text, got: {:?}",
        comment_text
    );
}

/// Test that multiple comments are all preserved.
#[test]
fn test_multiple_comments_roundtrip() {
    let mut builder = DocumentBuilder::new();

    {
        let mut cb = builder.add_comment();
        cb.set_author("Alice");
        cb.add_paragraph("First comment");
    }
    {
        let mut cb = builder.add_comment();
        cb.set_author("Bob");
        cb.add_paragraph("Second comment");
    }

    builder.add_paragraph("Document text");

    let mut doc = roundtrip(builder);
    let comments = doc.get_comments().expect("should load comments");
    assert_eq!(comments.comment.len(), 2, "should have two comments");

    let authors: Vec<&str> = comments.comment.iter().map(|c| c.author.as_str()).collect();
    assert!(authors.contains(&"Alice"), "should have Alice's comment");
    assert!(authors.contains(&"Bob"), "should have Bob's comment");
}

// =============================================================================
// 5. Track changes (read)
// =============================================================================

/// Test that tracked insertions and deletions are preserved across a roundtrip.
///
/// Uses `ins_run`/`del_run` helpers and verifies via `RevisionExt`:
/// `accepted_text()` and `rejected_text()`.
#[cfg(feature = "wml-track-changes")]
#[test]
fn test_track_changes_roundtrip() {
    use ooxml_wml::convenience::{del_run, ins_run};
    use ooxml_wml::ext::{RevisionExt, TrackChangeType};

    let mut builder = DocumentBuilder::new();

    {
        let para = builder.body_mut().add_paragraph();
        // Normal run
        para.add_run().set_text("Kept. ");
        // Tracked insertion
        para.paragraph_content.push(ins_run(
            1,
            "Alice",
            Some("2026-01-15T10:00:00Z"),
            "inserted text",
        ));
        // Tracked deletion
        para.paragraph_content
            .push(del_run(2, "Bob", None, "deleted text"));
    }

    let doc = roundtrip(builder);
    let para = &doc.body().paragraphs()[0];

    assert!(
        para.has_track_changes(),
        "paragraph should have track changes"
    );

    let changes = para.track_changes();
    assert_eq!(changes.len(), 2, "should have two tracked changes");

    // Check insertion
    let insertion = changes
        .iter()
        .find(|c| c.change_type == TrackChangeType::Insertion)
        .expect("should have an insertion");
    assert_eq!(insertion.author, "Alice");
    assert_eq!(insertion.text, "inserted text");
    assert_eq!(insertion.date.as_deref(), Some("2026-01-15T10:00:00Z"));

    // Check deletion
    let deletion = changes
        .iter()
        .find(|c| c.change_type == TrackChangeType::Deletion)
        .expect("should have a deletion");
    assert_eq!(deletion.author, "Bob");
    assert_eq!(deletion.text, "deleted text");

    // Accepted text: normal runs + insertions, no deletions
    let accepted = para.accepted_text();
    assert!(
        accepted.contains("Kept."),
        "accepted should contain kept text"
    );
    assert!(
        accepted.contains("inserted text"),
        "accepted should contain insertion"
    );
    assert!(
        !accepted.contains("deleted text"),
        "accepted should NOT contain deletion"
    );

    // Rejected text: normal runs + deletions, no insertions
    let rejected = para.rejected_text();
    assert!(
        rejected.contains("Kept."),
        "rejected should contain kept text"
    );
    assert!(
        rejected.contains("deleted text"),
        "rejected should contain deletion"
    );
    assert!(
        !rejected.contains("inserted text"),
        "rejected should NOT contain insertion"
    );
}

/// Test using the Paragraph convenience methods add_tracked_insertion / add_tracked_deletion.
#[cfg(feature = "wml-track-changes")]
#[test]
fn test_track_changes_convenience_methods() {
    use ooxml_wml::ext::RevisionExt;

    let mut builder = DocumentBuilder::new();
    {
        let para = builder.body_mut().add_paragraph();
        para.add_run().set_text("Start. ");
        para.add_tracked_insertion(10, "Carol", None, "new section");
        para.add_tracked_deletion(11, "Carol", None, "old section");
    }

    let doc = roundtrip(builder);
    let para = &doc.body().paragraphs()[0];

    assert!(para.has_track_changes());
    let changes = para.track_changes();
    assert_eq!(changes.len(), 2, "should have two tracked changes");
    assert_eq!(changes[0].author, "Carol");
    assert_eq!(changes[1].author, "Carol");
}

// =============================================================================
// 6. Anchored images
// =============================================================================

/// Test that anchored (floating) images survive a roundtrip.
///
/// Writes an anchored image with Square wrap, reads back, and verifies
/// the rel ID is discoverable via DrawingExt::anchored_image_rel_ids().
#[cfg(all(feature = "wml-drawings", feature = "extra-children"))]
#[test]
fn test_anchored_image_roundtrip() {
    use ooxml_wml::WrapType;
    use ooxml_wml::ext::DrawingExt;

    // Minimal valid PNG (1x1 px)
    let png_data: Vec<u8> = vec![
        0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A, 0x00, 0x00, 0x00, 0x0D, 0x49, 0x48, 0x44,
        0x52, 0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x01, 0x08, 0x02, 0x00, 0x00, 0x00, 0x90,
        0x77, 0x53, 0xDE, 0x00, 0x00, 0x00, 0x0C, 0x49, 0x44, 0x41, 0x54, 0x08, 0xD7, 0x63, 0xF8,
        0xCF, 0xC0, 0x00, 0x00, 0x00, 0x03, 0x00, 0x01, 0x00, 0x18, 0xDD, 0x8D, 0xB4, 0x00, 0x00,
        0x00, 0x00, 0x49, 0x45, 0x4E, 0x44, 0xAE, 0x42, 0x60, 0x82,
    ];

    let mut builder = DocumentBuilder::new();
    let rel_id = builder.add_image(png_data.clone(), "image/png");

    {
        let para = builder.body_mut().add_paragraph();
        let run = para.add_run();
        let mut drawing = Drawing::new();
        drawing
            .add_anchored_image(&rel_id)
            .set_width_inches(3.0)
            .set_height_inches(2.0)
            .set_wrap_type(WrapType::Square)
            .set_description("Anchored test image");
        let mut doc_id = 1usize;
        run.add_drawing(drawing.build(&mut doc_id));
    }

    let mut doc = roundtrip(builder);

    // Image file should be in the package
    assert!(
        doc.package().has_part("word/media/image1.png"),
        "image should be in package"
    );

    // Find the anchored image rel ID via DrawingExt
    let image_rel_id: String;
    {
        let para = &doc.body().paragraphs()[0];
        let run = &para.runs()[0];
        assert!(run.has_images(), "run should have images");
        let drawings = run.drawings();
        assert_eq!(drawings.len(), 1, "run should have one drawing");
        let drawing = drawings[0];
        let anchored_ids = drawing.anchored_image_rel_ids();
        assert_eq!(
            anchored_ids.len(),
            1,
            "drawing should have one anchored image rel ID"
        );
        image_rel_id = anchored_ids[0].to_string();
    }

    assert_eq!(image_rel_id, rel_id, "anchored image rel ID should match");

    // Verify image data is retrievable
    let image_data = doc.get_image_data(&image_rel_id).unwrap();
    assert_eq!(image_data.content_type, "image/png");
    assert_eq!(image_data.data, png_data);
}

// =============================================================================
// 7. Bookmarks (already covered in integration.rs but we add a richer variant)
// =============================================================================

/// Test multiple bookmarks in the same paragraph and across paragraphs.
#[test]
fn test_multiple_bookmarks_roundtrip() {
    use ooxml_wml::types::ParagraphContent;

    let mut builder = DocumentBuilder::new();

    // Para 1: two bookmarks
    {
        let para = builder.body_mut().add_paragraph();
        para.add_bookmark_start(1, "section_intro");
        para.add_run().set_text("Introduction");
        para.add_bookmark_end(1);
        para.add_bookmark_start(2, "section_intro_end");
        para.add_bookmark_end(2);
    }

    // Para 2: another bookmark
    {
        let para = builder.body_mut().add_paragraph();
        para.add_bookmark_start(3, "conclusion");
        para.add_run().set_text("Conclusion");
        para.add_bookmark_end(3);
    }

    let doc = roundtrip(builder);
    let paras = doc.body().paragraphs();
    assert_eq!(paras.len(), 2, "should have two paragraphs");

    // Para 1: check both bookmark starts
    let p1 = &paras[0];
    let bookmark_names: Vec<&str> = p1
        .paragraph_content
        .iter()
        .filter_map(|c| match c {
            ParagraphContent::BookmarkStart(b) => Some(b.name.as_str()),
            _ => None,
        })
        .collect();
    assert!(
        bookmark_names.contains(&"section_intro"),
        "should have section_intro bookmark"
    );
    assert!(
        bookmark_names.contains(&"section_intro_end"),
        "should have section_intro_end bookmark"
    );

    // Para 2: check conclusion bookmark
    let p2 = &paras[1];
    let found = p2
        .paragraph_content
        .iter()
        .any(|c| matches!(c, ParagraphContent::BookmarkStart(b) if b.name == "conclusion"));
    assert!(found, "should have conclusion bookmark");
}

// =============================================================================
// 8. Section properties (page layout)
// =============================================================================

/// Test setting page margins and size in section properties and reading them back.
#[cfg(feature = "wml-layout")]
#[test]
fn test_section_properties_roundtrip() {
    use ooxml_wml::ext::SectionPropertiesExt;
    use ooxml_wml::types::{PageMargins, PageSize, SectionProperties};

    let mut builder = DocumentBuilder::new();
    builder.add_paragraph("Content on a custom-sized page");

    let sect_pr = SectionProperties {
        pg_sz: Some(Box::new(PageSize {
            width: Some("12240".to_string()),  // Letter width (8.5 in)
            height: Some("15840".to_string()), // Letter height (11 in)
            orient: None,
            code: None,
            #[cfg(feature = "extra-attrs")]
            extra_attrs: Default::default(),
        })),
        pg_mar: Some(Box::new(PageMargins {
            top: "1440".to_string(),    // 1 inch
            bottom: "1440".to_string(), // 1 inch
            left: "1800".to_string(),   // 1.25 inch
            right: "1800".to_string(),  // 1.25 inch
            header: "720".to_string(),
            footer: "720".to_string(),
            gutter: "0".to_string(),
            #[cfg(feature = "extra-attrs")]
            extra_attrs: Default::default(),
        })),
        ..Default::default()
    };
    builder.body_mut().set_section_properties(sect_pr);

    let doc = roundtrip(builder);

    let sect = doc
        .body()
        .section_properties()
        .expect("section properties should exist");

    assert_eq!(
        sect.page_width_twips(),
        Some(12240),
        "page width should be 12240 twips"
    );
    assert_eq!(
        sect.page_height_twips(),
        Some(15840),
        "page height should be 15840 twips"
    );

    let margins = sect.page_margins().expect("page margins should exist");
    assert_eq!(margins.top, "1440", "top margin");
    assert_eq!(margins.bottom, "1440", "bottom margin");
    assert_eq!(margins.left, "1800", "left margin");
    assert_eq!(margins.right, "1800", "right margin");
}

/// Test landscape orientation round-trips correctly.
#[cfg(feature = "wml-layout")]
#[test]
fn test_landscape_orientation_roundtrip() {
    use ooxml_wml::ext::SectionPropertiesExt;
    use ooxml_wml::types::{PageSize, STPageOrientation, SectionProperties};

    let mut builder = DocumentBuilder::new();
    builder.add_paragraph("Landscape page");

    let sect_pr = SectionProperties {
        pg_sz: Some(Box::new(PageSize {
            width: Some("15840".to_string()),
            height: Some("12240".to_string()),
            orient: Some(STPageOrientation::Landscape),
            code: None,
            #[cfg(feature = "extra-attrs")]
            extra_attrs: Default::default(),
        })),
        ..Default::default()
    };
    builder.body_mut().set_section_properties(sect_pr);

    let doc = roundtrip(builder);

    let sect = doc.body().section_properties().expect("section properties");
    assert_eq!(
        sect.page_orientation(),
        Some(&STPageOrientation::Landscape),
        "orientation should be Landscape"
    );
    assert_eq!(sect.page_width_twips(), Some(15840));
    assert_eq!(sect.page_height_twips(), Some(12240));
}

// =============================================================================
// 9. Complex table (merged cells)
// =============================================================================

/// Test table with cells using grid span (horizontal merge).
///
/// ECMA-376 §17.4.17 (gridSpan) — a cell spanning multiple columns.
#[cfg(feature = "wml-tables")]
#[test]
fn test_gridspan_roundtrip() {
    use ooxml_wml::types::{CTDecimalNumber, TableCellProperties};

    let mut builder = DocumentBuilder::new();
    {
        let table = builder.body_mut().add_table();

        // Row 1: cell spanning 2 columns + one normal cell
        let row1 = table.add_row();
        {
            let cell = row1.add_cell();
            cell.add_paragraph().add_run().set_text("Spans 2 cols");
            // Apply gridSpan=2
            let tcp = TableCellProperties {
                grid_span: Some(Box::new(CTDecimalNumber {
                    value: 2,
                    #[cfg(feature = "extra-attrs")]
                    extra_attrs: Default::default(),
                })),
                ..Default::default()
            };
            cell.cell_properties = Some(Box::new(tcp));
        }
        {
            let cell = row1.add_cell();
            cell.add_paragraph().add_run().set_text("Cell C");
        }

        // Row 2: three individual cells
        let row2 = table.add_row();
        row2.add_cell()
            .add_paragraph()
            .add_run()
            .set_text("Cell A2");
        row2.add_cell()
            .add_paragraph()
            .add_run()
            .set_text("Cell B2");
        row2.add_cell()
            .add_paragraph()
            .add_run()
            .set_text("Cell C2");
    }

    let doc = roundtrip(builder);
    let tables = doc.body().tables();
    assert_eq!(tables.len(), 1, "table count");

    let table = &tables[0];
    assert_eq!(table.row_count(), 2, "row count");

    // Verify grid span on row 1, cell 0
    let rows = table.rows();
    let row1_cells = rows[0].cells();
    assert_eq!(
        row1_cells.len(),
        2,
        "row 1 should have 2 cells (1 spanning + 1 normal)"
    );
    assert_eq!(row1_cells[0].text(), "Spans 2 cols");

    #[cfg(feature = "wml-tables")]
    {
        let tcp = row1_cells[0]
            .properties()
            .expect("cell should have properties");
        let gs = tcp.grid_span.as_ref().expect("cell should have grid_span");
        assert_eq!(gs.value, 2, "grid_span should be 2");
    }
}

/// Test table with vertically merged cells (vMerge).
///
/// ECMA-376 §17.4.84 (vMerge) — vertical cell merging via restart and continuation.
#[cfg(feature = "wml-tables")]
#[test]
fn test_vmerge_roundtrip() {
    use ooxml_wml::types::{CTVMerge, STMerge, TableCellProperties};

    let mut builder = DocumentBuilder::new();
    {
        let table = builder.body_mut().add_table();

        // Row 1: merged start cell + normal cell
        let row1 = table.add_row();
        {
            let cell = row1.add_cell();
            cell.add_paragraph().add_run().set_text("Merged content");
            // vMerge restart (start of merge)
            let tcp = TableCellProperties {
                vertical_merge: Some(Box::new(CTVMerge {
                    value: Some(STMerge::Restart),
                    #[cfg(feature = "extra-attrs")]
                    extra_attrs: Default::default(),
                })),
                ..Default::default()
            };
            cell.cell_properties = Some(Box::new(tcp));
        }
        row1.add_cell()
            .add_paragraph()
            .add_run()
            .set_text("Normal R1C2");

        // Row 2: merge continuation cell + normal cell
        let row2 = table.add_row();
        {
            let cell = row2.add_cell();
            cell.add_paragraph(); // empty para required in OOXML for continuation cells
            // vMerge continuation (empty val = continuation)
            let tcp = TableCellProperties {
                vertical_merge: Some(Box::new(CTVMerge {
                    value: None,
                    #[cfg(feature = "extra-attrs")]
                    extra_attrs: Default::default(),
                })),
                ..Default::default()
            };
            cell.cell_properties = Some(Box::new(tcp));
        }
        row2.add_cell()
            .add_paragraph()
            .add_run()
            .set_text("Normal R2C2");
    }

    let doc = roundtrip(builder);
    let tables = doc.body().tables();
    assert_eq!(tables.len(), 1, "table count");

    let table = &tables[0];
    assert_eq!(table.row_count(), 2);

    let rows = table.rows();
    #[cfg(feature = "wml-tables")]
    {
        // Row 1, cell 0 should have vMerge=restart
        let r1c0_props = rows[0].cells()[0]
            .properties()
            .expect("should have cell properties");
        let vmerge = r1c0_props
            .vertical_merge
            .as_ref()
            .expect("should have vMerge");
        assert_eq!(
            vmerge.value,
            Some(STMerge::Restart),
            "first cell should have vMerge restart"
        );

        // Row 2, cell 0 should have vMerge=continuation (None value)
        let r2c0_props = rows[1].cells()[0]
            .properties()
            .expect("should have cell properties");
        let vmerge2 = r2c0_props
            .vertical_merge
            .as_ref()
            .expect("should have vMerge");
        assert_eq!(
            vmerge2.value, None,
            "second cell should have vMerge continuation (None)"
        );
    }
}

// =============================================================================
// 10. Document with combined features
// =============================================================================

/// Integration test combining multiple features: headers, footnotes, comments,
/// and section properties in a single document.
#[cfg(all(feature = "wml-layout", feature = "extra-attrs"))]
#[test]
fn test_combined_features_roundtrip() {
    use ooxml_wml::HeaderFooterType;
    use ooxml_wml::ext::SectionPropertiesExt;
    use ooxml_wml::types::{PageSize, SectionProperties};

    let mut builder = DocumentBuilder::new();

    // Add a header
    builder
        .add_header(HeaderFooterType::Default)
        .add_paragraph("Report Header");

    // Add a footnote
    let fn_id;
    {
        let mut fb = builder.add_footnote();
        fb.add_paragraph("Supporting reference");
        fn_id = fb.id();
    }

    // Add a comment
    let cm_id;
    {
        let mut cb = builder.add_comment();
        cb.set_author("Reviewer");
        cb.add_paragraph("Review comment");
        cm_id = cb.id();
    }

    // Add body content
    {
        let para = builder.body_mut().add_paragraph();
        para.add_comment_range_start(cm_id);
        para.add_run().set_text("Report body text");
        para.add_comment_range_end(cm_id);
        para.add_run().add_footnote_ref(fn_id as i64);
        para.add_run().add_comment_ref(cm_id as i64);
    }

    // Set section properties
    let sect_pr = SectionProperties {
        pg_sz: Some(Box::new(PageSize {
            width: Some("12240".to_string()),
            height: Some("15840".to_string()),
            orient: None,
            code: None,
            #[cfg(feature = "extra-attrs")]
            extra_attrs: Default::default(),
        })),
        ..Default::default()
    };
    builder.body_mut().set_section_properties(sect_pr);

    let mut doc = roundtrip(builder);

    // Verify all features
    assert!(
        doc.package().has_part("word/footnotes.xml"),
        "footnotes.xml"
    );
    assert!(doc.package().has_part("word/comments.xml"), "comments.xml");

    let footnotes = doc.get_footnotes().unwrap();
    let user_fns: Vec<_> = footnotes.footnote.iter().filter(|f| f.id >= 1).collect();
    assert_eq!(user_fns.len(), 1);

    let comments = doc.get_comments().unwrap();
    assert_eq!(comments.comment.len(), 1);
    assert_eq!(comments.comment[0].author, "Reviewer");

    let sect = doc.body().section_properties().expect("section properties");
    assert_eq!(sect.page_width_twips(), Some(12240));
    assert_eq!(sect.page_height_twips(), Some(15840));
}
