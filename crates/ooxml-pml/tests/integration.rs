//! Integration tests for ooxml-pml: builder → write → read roundtrip tests.
//!
//! These tests exercise the public API end-to-end: build a presentation with
//! `PresentationBuilder`, serialize it to an in-memory `Cursor`, then read it
//! back with `Presentation::from_reader` and assert on the result.

use ooxml_pml::{Presentation, PresentationBuilder, TableBuilder, TextRun};
use std::io::Cursor;

// ---------------------------------------------------------------------------
// Helper
// ---------------------------------------------------------------------------

/// Write `builder` to an in-memory buffer and read it back.
fn write_and_read(builder: PresentationBuilder) -> Presentation<Cursor<Vec<u8>>> {
    let mut buf = Cursor::new(Vec::new());
    builder.write(&mut buf).expect("write should succeed");
    buf.set_position(0);
    Presentation::from_reader(buf).expect("read should succeed")
}

// ---------------------------------------------------------------------------
// 1. Basic presentation roundtrip
// ---------------------------------------------------------------------------

/// A presentation with two slides can be written and read back, preserving
/// the slide count.
#[test]
fn test_presentation_roundtrip() {
    let mut builder = PresentationBuilder::new();
    builder.add_slide().add_title("Slide One");
    builder.add_slide().add_title("Slide Two");

    let pres = write_and_read(builder);
    assert_eq!(pres.slide_count(), 2);
}

// ---------------------------------------------------------------------------
// 2. Slide text content
// ---------------------------------------------------------------------------

/// Text added with `add_text` is preserved through the roundtrip.
#[test]
fn test_slide_text_content() {
    let mut builder = PresentationBuilder::new();
    {
        let slide = builder.add_slide();
        slide.add_text("Hello from ooxml-pml");
    }

    let mut pres = write_and_read(builder);
    let slide = pres.slide(0).unwrap();
    let text = slide.text();
    assert!(
        text.contains("Hello from ooxml-pml"),
        "expected text not found; got: {:?}",
        text
    );
}

// ---------------------------------------------------------------------------
// 3. Slide title
// ---------------------------------------------------------------------------

/// A title added with `add_title` is preserved and present in the slide text.
#[test]
fn test_slide_title() {
    let mut builder = PresentationBuilder::new();
    builder.add_slide().add_title("My Presentation Title");

    let mut pres = write_and_read(builder);
    let slide = pres.slide(0).unwrap();

    // The title shape should appear in slide shapes with the correct text.
    use ooxml_pml::ShapeExt;
    let title_shape = slide
        .shapes()
        .iter()
        .find(|s| s.name() == "Title")
        .expect("title shape should exist");

    let title_text = title_shape.text().unwrap_or_default();
    assert!(
        title_text.contains("My Presentation Title"),
        "title text mismatch; got: {:?}",
        title_text
    );
}

// ---------------------------------------------------------------------------
// 4. Multiple slides ordering
// ---------------------------------------------------------------------------

/// Three slides built in order should be read back in the same order.
#[test]
fn test_slide_ordering() {
    let mut builder = PresentationBuilder::new();
    builder.add_slide().add_title("First");
    builder.add_slide().add_title("Second");
    builder.add_slide().add_title("Third");

    let mut pres = write_and_read(builder);
    assert_eq!(pres.slide_count(), 3);

    let slides = pres.slides().unwrap();
    assert_eq!(slides[0].index(), 0);
    assert_eq!(slides[1].index(), 1);
    assert_eq!(slides[2].index(), 2);

    // Verify content is on the right slide.
    assert!(slides[0].text().contains("First"), "slide 0 text wrong");
    assert!(slides[1].text().contains("Second"), "slide 1 text wrong");
    assert!(slides[2].text().contains("Third"), "slide 2 text wrong");
}

// ---------------------------------------------------------------------------
// 5. Speaker notes
// ---------------------------------------------------------------------------

/// Speaker notes set with `set_notes` are preserved through the roundtrip.
#[test]
fn test_speaker_notes() {
    let mut builder = PresentationBuilder::new();
    {
        let slide = builder.add_slide();
        slide.add_title("Notes Demo");
        slide.set_notes("Remember to speak slowly.");
    }

    let mut pres = write_and_read(builder);
    let slide = pres.slide(0).unwrap();

    assert!(slide.has_notes(), "slide should have notes");
    let notes = slide.notes().expect("notes should be Some");
    assert!(
        notes.contains("Remember to speak slowly."),
        "notes text mismatch; got: {:?}",
        notes
    );
}

/// A slide without notes has `has_notes() == false`.
#[test]
fn test_no_speaker_notes() {
    let mut builder = PresentationBuilder::new();
    builder.add_slide().add_title("No Notes");

    let mut pres = write_and_read(builder);
    let slide = pres.slide(0).unwrap();
    assert!(!slide.has_notes());
    assert!(slide.notes().is_none());
}

// ---------------------------------------------------------------------------
// 6. Inline image roundtrip
// ---------------------------------------------------------------------------

/// A minimal valid 1×1 PNG embedded in a slide is preserved through the roundtrip.
///
/// PNG structure: 8-byte signature + IHDR chunk + IDAT chunk + IEND chunk.
/// This is the smallest possible valid PNG (1×1 pixel, 8-bit grayscale).
#[test]
fn test_image_roundtrip() {
    // Minimal 1×1 grayscale PNG (generated via Python's PIL, kept as hex).
    let png_bytes: &[u8] = &[
        0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A, // PNG signature
        0x00, 0x00, 0x00, 0x0D, 0x49, 0x48, 0x44, 0x52, // IHDR length + type
        0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x01, // width=1, height=1
        0x08, 0x00, 0x00, 0x00, 0x00, 0x3A, 0x7E, 0x9B, // bit depth=8, color=0 (gray), crc
        0x55, 0x00, 0x00, 0x00, 0x0A, 0x49, 0x44, 0x41, // IDAT length + type
        0x54, 0x78, 0x9C, 0x62, 0x00, 0x00, 0x00, 0x02, // zlib+deflate data
        0x00, 0x01, 0xE2, 0x21, 0xBC, 0x33, 0x00, 0x00, // IDAT data + crc
        0x00, 0x00, 0x49, 0x45, 0x4E, 0x44, 0xAE, 0x42, // IEND type + crc
        0x60, 0x82,
    ];

    let mut builder = PresentationBuilder::new();
    {
        let slide = builder.add_slide();
        slide.add_image(
            png_bytes.to_vec(),
            914400,
            914400, // x=1in, y=1in
            914400,
            914400, // width=1in, height=1in
        );
    }

    let mut buf = Cursor::new(Vec::new());
    builder.write(&mut buf).expect("write should succeed");
    buf.set_position(0);
    let mut pres = Presentation::from_reader(buf).expect("read should succeed");

    let slide = pres.slide(0).unwrap();

    // The slide should have exactly one picture.
    let pictures = slide.pictures();
    assert_eq!(pictures.len(), 1, "expected one picture");

    // The image bytes should be recoverable via get_image_data.
    let img = pres.get_image_data(&slide, &pictures[0]).unwrap();
    assert_eq!(img.data, png_bytes, "image bytes should match");
    assert_eq!(img.content_type, "image/png");
}

// ---------------------------------------------------------------------------
// 7. Table in slide
// ---------------------------------------------------------------------------

/// A table added with `add_table` is recovered through the roundtrip.
#[cfg(feature = "dml-tables")]
#[test]
fn test_table_roundtrip() {
    use ooxml_dml::TableCellExt;

    let mut builder = PresentationBuilder::new();
    {
        let slide = builder.add_slide();
        slide.add_title("Table Slide");

        let table = TableBuilder::new()
            .name("Test Table")
            .add_row(["Alpha", "Beta", "Gamma"])
            .add_row(["1", "2", "3"])
            .add_row(["X", "Y", "Z"]);

        // x=1in, y=2in, w=8in, h=2in (all in EMUs)
        slide.add_table(table, 914400, 1828800, 7315200, 1828800);
    }

    let mut pres = write_and_read(builder);
    let slide = pres.slide(0).unwrap();

    assert!(slide.has_tables(), "slide should have tables");
    assert_eq!(slide.table_count(), 1);

    let table = slide.table(0).unwrap();
    assert_eq!(table.row_count(), 3, "row count mismatch");
    assert_eq!(table.col_count(), 3, "col count mismatch");

    // Check cell content
    assert_eq!(table.cell(0, 0).unwrap().text(), "Alpha");
    assert_eq!(table.cell(0, 1).unwrap().text(), "Beta");
    assert_eq!(table.cell(0, 2).unwrap().text(), "Gamma");
    assert_eq!(table.cell(1, 0).unwrap().text(), "1");
    assert_eq!(table.cell(2, 2).unwrap().text(), "Z");
}

/// `to_text_grid` returns correct 2-D structure for a table.
#[cfg(feature = "dml-tables")]
#[test]
fn test_table_text_grid() {
    let mut builder = PresentationBuilder::new();
    {
        let slide = builder.add_slide();
        let table = TableBuilder::new()
            .add_row(["Row0Col0", "Row0Col1"])
            .add_row(["Row1Col0", "Row1Col1"]);
        slide.add_table(table, 0, 0, 9144000, 6858000);
    }

    let mut pres = write_and_read(builder);
    let slide = pres.slide(0).unwrap();

    let grid = slide.table(0).unwrap().to_text_grid();
    assert_eq!(grid.len(), 2);
    assert_eq!(grid[0], vec!["Row0Col0", "Row0Col1"]);
    assert_eq!(grid[1], vec!["Row1Col0", "Row1Col1"]);
}

/// An empty `TableBuilder` (no rows) does not produce a table on the slide.
#[test]
fn test_empty_table_not_added() {
    let mut builder = PresentationBuilder::new();
    {
        let slide = builder.add_slide();
        // add_table with no rows is a no-op per the writer implementation.
        slide.add_table(TableBuilder::new(), 0, 0, 9144000, 6858000);
    }

    let mut pres = write_and_read(builder);
    let slide = pres.slide(0).unwrap();
    assert!(!slide.has_tables(), "no table should have been added");
}

// ---------------------------------------------------------------------------
// 8. Hyperlink roundtrip
// ---------------------------------------------------------------------------

/// A hyperlink added with `add_hyperlink` is recoverable through the roundtrip.
#[test]
fn test_hyperlink_roundtrip() {
    let mut builder = PresentationBuilder::new();
    {
        let slide = builder.add_slide();
        slide.add_hyperlink(
            "Visit Rust",
            "https://www.rust-lang.org",
            457200,
            1600200,
            8229600,
            457200,
        );
    }

    let mut pres = write_and_read(builder);
    let slide = pres.slide(0).unwrap();

    assert!(slide.has_hyperlinks(), "slide should have hyperlinks");

    let links = slide.hyperlinks();
    assert_eq!(links.len(), 1);
    assert_eq!(links[0].text, "Visit Rust");

    let url = pres.resolve_hyperlink(&slide, &links[0].rel_id).unwrap();
    assert_eq!(url, "https://www.rust-lang.org");
}

/// `get_hyperlinks_with_urls` returns resolved (text, url) pairs.
#[test]
fn test_get_hyperlinks_with_urls() {
    let mut builder = PresentationBuilder::new();
    {
        let slide = builder.add_slide();
        slide.add_hyperlink("Rust", "https://www.rust-lang.org", 0, 0, 1000, 1000);
        slide.add_hyperlink("Crates", "https://crates.io", 0, 100000, 1000, 1000);
    }

    let mut pres = write_and_read(builder);
    let slide = pres.slide(0).unwrap();

    let links = pres.get_hyperlinks_with_urls(&slide).unwrap();
    assert_eq!(links.len(), 2);

    // Order is not guaranteed — collect into a map for comparison.
    let map: std::collections::HashMap<_, _> = links.into_iter().collect();
    assert_eq!(map["Rust"], "https://www.rust-lang.org");
    assert_eq!(map["Crates"], "https://crates.io");
}

/// Mixed text runs (plain + hyperlink) preserve both run types.
#[test]
fn test_mixed_text_runs_roundtrip() {
    let mut builder = PresentationBuilder::new();
    {
        let slide = builder.add_slide();
        slide.add_text_with_runs(
            vec![
                TextRun::text("Check out "),
                TextRun::hyperlink("the docs", "https://docs.rs"),
                TextRun::text(" for more."),
            ],
            457200,
            1600200,
            8229600,
            457200,
        );
    }

    let mut pres = write_and_read(builder);
    let slide = pres.slide(0).unwrap();

    assert!(slide.has_hyperlinks());
    let links = slide.hyperlinks();
    assert_eq!(links.len(), 1);
    assert_eq!(links[0].text, "the docs");

    let url = pres.resolve_hyperlink(&slide, &links[0].rel_id).unwrap();
    assert_eq!(url, "https://docs.rs");
}

// ---------------------------------------------------------------------------
// 9. Chart rel IDs (pml-charts feature)
// ---------------------------------------------------------------------------

/// A normal slide without any chart graphic frames has no chart rel IDs.
#[cfg(feature = "pml-charts")]
#[test]
fn test_chart_rel_ids_empty_for_normal_slide() {
    let mut builder = PresentationBuilder::new();
    builder.add_slide().add_text("No chart here");

    let mut pres = write_and_read(builder);
    let slide = pres.slide(0).unwrap();
    assert!(
        slide.chart_rel_ids().is_empty(),
        "expected no chart rel IDs on a plain text slide"
    );
}

// ---------------------------------------------------------------------------
// 10. SmartArt rel IDs (pml-charts feature)
// ---------------------------------------------------------------------------

/// A normal slide without any SmartArt has no SmartArt rel ID sets.
#[cfg(feature = "pml-charts")]
#[test]
fn test_smartart_rel_ids_empty_for_normal_slide() {
    let mut builder = PresentationBuilder::new();
    builder.add_slide().add_text("No SmartArt here");

    let mut pres = write_and_read(builder);
    let slide = pres.slide(0).unwrap();
    assert!(
        slide.smartart_rel_ids().is_empty(),
        "expected no SmartArt rel IDs on a plain text slide"
    );
}

// ---------------------------------------------------------------------------
// 11. Slide count
// ---------------------------------------------------------------------------

/// The slide count is correct for presentations with 0, 1, and N slides.
#[test]
fn test_slide_count_zero() {
    let builder = PresentationBuilder::new();
    let pres = write_and_read(builder);
    assert_eq!(pres.slide_count(), 0);
}

#[test]
fn test_slide_count_one() {
    let mut builder = PresentationBuilder::new();
    builder.add_slide();
    let pres = write_and_read(builder);
    assert_eq!(pres.slide_count(), 1);
}

#[test]
fn test_slide_count_many() {
    let mut builder = PresentationBuilder::new();
    for i in 0..10 {
        builder.add_slide().add_title(format!("Slide {}", i));
    }
    let pres = write_and_read(builder);
    assert_eq!(pres.slide_count(), 10);
}

// ---------------------------------------------------------------------------
// 12. Slide master and layout
// ---------------------------------------------------------------------------

/// Every written presentation contains exactly one slide master and one layout.
#[test]
fn test_slide_master_and_layout_present() {
    let mut builder = PresentationBuilder::new();
    builder.add_slide().add_title("Master Test");

    let pres = write_and_read(builder);

    assert_eq!(pres.slide_masters().len(), 1, "should have 1 slide master");
    assert_eq!(pres.slide_layouts().len(), 1, "should have 1 slide layout");

    let master = &pres.slide_masters()[0];
    assert_eq!(master.layout_count(), 1, "master should reference 1 layout");
}

/// `layout_by_name` finds the layout by its name attribute.
#[test]
fn test_layout_by_name() {
    let mut builder = PresentationBuilder::new();
    builder.add_slide();

    let pres = write_and_read(builder);

    // The minimal writer emits a layout named "Blank".
    let found = pres.layout_by_name("Blank");
    assert!(found.is_some(), "expected to find layout named 'Blank'");
}

// ---------------------------------------------------------------------------
// 13. Slide layout reference
// ---------------------------------------------------------------------------

/// Every slide written by `PresentationBuilder` has a layout rel ID.
#[test]
fn test_slide_has_layout_rel_id() {
    let mut builder = PresentationBuilder::new();
    builder.add_slide().add_title("Layout Ref Test");

    let mut pres = write_and_read(builder);
    let slide = pres.slide(0).unwrap();

    assert!(
        slide.layout_rel_id().is_some(),
        "slide should reference a layout"
    );
}

// ---------------------------------------------------------------------------
// 14. Slide index() accessor
// ---------------------------------------------------------------------------

#[test]
fn test_slide_index_accessor() {
    let mut builder = PresentationBuilder::new();
    for _ in 0..3 {
        builder.add_slide();
    }

    let mut pres = write_and_read(builder);
    for i in 0..3 {
        let slide = pres.slide(i).unwrap();
        assert_eq!(slide.index(), i);
    }
}

// ---------------------------------------------------------------------------
// 15. Out-of-range slide access returns error
// ---------------------------------------------------------------------------

#[test]
fn test_slide_out_of_range() {
    let mut builder = PresentationBuilder::new();
    builder.add_slide();

    let mut pres = write_and_read(builder);
    assert!(pres.slide(1).is_err(), "index 1 should be out of range");
    assert!(pres.slide(99).is_err(), "index 99 should be out of range");
}

// ---------------------------------------------------------------------------
// 16. Widescreen slide size
// ---------------------------------------------------------------------------

/// A widescreen presentation can be written and read back without error.
#[test]
fn test_widescreen_presentation() {
    let mut builder = PresentationBuilder::new();
    builder.set_widescreen();
    builder.add_slide().add_title("Widescreen");

    let pres = write_and_read(builder);
    assert_eq!(pres.slide_count(), 1);
}

// ---------------------------------------------------------------------------
// 17. Slide transition
// ---------------------------------------------------------------------------

#[cfg(feature = "pml-transitions")]
#[test]
fn test_slide_transition_roundtrip() {
    use ooxml_pml::{SlideTransition, TransitionSpeed, TransitionType};

    let mut builder = PresentationBuilder::new();
    {
        let slide = builder.add_slide();
        slide.add_title("Transition Slide");
        slide.set_transition(
            SlideTransition::new(TransitionType::Wipe)
                .with_speed(TransitionSpeed::Fast)
                .with_advance_after(2000),
        );
    }

    let mut pres = write_and_read(builder);
    let slide = pres.slide(0).unwrap();

    assert!(slide.has_transition(), "slide should have a transition");
    let t = slide.transition().unwrap();
    assert_eq!(t.transition_type, Some(TransitionType::Wipe));
    assert_eq!(t.speed, TransitionSpeed::Fast);
    assert_eq!(t.advance_time_ms, Some(2000));
}

/// A slide with no transition set has `has_transition() == false`.
#[cfg(feature = "pml-transitions")]
#[test]
fn test_no_transition() {
    let mut builder = PresentationBuilder::new();
    builder.add_slide().add_title("No Transition");

    let mut pres = write_and_read(builder);
    let slide = pres.slide(0).unwrap();
    assert!(!slide.has_transition());
    assert!(slide.transition().is_none());
}

// ---------------------------------------------------------------------------
// 18. `add_text_at` positions text at explicit coordinates
// ---------------------------------------------------------------------------

/// Text added with `add_text_at` is preserved through the roundtrip.
#[test]
fn test_add_text_at_roundtrip() {
    let mut builder = PresentationBuilder::new();
    {
        let slide = builder.add_slide();
        slide.add_text_at("Positioned text", 457200, 457200, 4572000, 914400);
    }

    let mut pres = write_and_read(builder);
    let slide = pres.slide(0).unwrap();
    assert!(
        slide.text().contains("Positioned text"),
        "text not found; got: {:?}",
        slide.text()
    );
}

// ---------------------------------------------------------------------------
// 19. Multiple text elements on one slide
// ---------------------------------------------------------------------------

#[test]
fn test_multiple_text_elements() {
    let mut builder = PresentationBuilder::new();
    {
        let slide = builder.add_slide();
        slide.add_title("Title Here");
        slide.add_text("Body line one");
        slide.add_text("Body line two");
    }

    let mut pres = write_and_read(builder);
    let slide = pres.slide(0).unwrap();
    let text = slide.text();

    assert!(text.contains("Title Here"), "title missing");
    assert!(text.contains("Body line one"), "body line 1 missing");
    assert!(text.contains("Body line two"), "body line 2 missing");
}

// ---------------------------------------------------------------------------
// 20. Image format detection: JPEG
// ---------------------------------------------------------------------------

/// A minimal JPEG is detected as JPEG format and stored with the correct content type.
#[test]
fn test_jpeg_image_roundtrip() {
    // Minimal JPEG: SOI marker + EOI marker (smallest valid JPEG).
    let jpeg_bytes: Vec<u8> = vec![
        0xFF, 0xD8, 0xFF, 0xE0, // SOI + APP0 marker
        0x00, 0x10, // APP0 length (16 bytes)
        0x4A, 0x46, 0x49, 0x46, 0x00, // "JFIF\0"
        0x01, 0x01, // version 1.1
        0x00, // aspect ratio units
        0x00, 0x01, 0x00, 0x01, // X/Y density = 1,1
        0x00, 0x00, // thumbnail size 0×0
        0xFF, 0xD9, // EOI
    ];

    let mut builder = PresentationBuilder::new();
    {
        let slide = builder.add_slide();
        slide.add_image(jpeg_bytes.clone(), 0, 0, 914400, 914400);
    }

    let mut buf = Cursor::new(Vec::new());
    builder.write(&mut buf).expect("write");
    buf.set_position(0);
    let mut pres = Presentation::from_reader(buf).expect("read");

    let slide = pres.slide(0).unwrap();
    let pictures = slide.pictures();
    assert_eq!(pictures.len(), 1);

    let img = pres.get_image_data(&slide, &pictures[0]).unwrap();
    assert_eq!(img.data, jpeg_bytes);
    assert_eq!(img.content_type, "image/jpeg");
}

// ---------------------------------------------------------------------------
// 21. Notes on multiple slides are independent
// ---------------------------------------------------------------------------

#[test]
fn test_notes_independence() {
    let mut builder = PresentationBuilder::new();
    {
        let s0 = builder.add_slide();
        s0.add_title("Slide A");
        s0.set_notes("Notes for A");
    }
    {
        let s1 = builder.add_slide();
        s1.add_title("Slide B");
        // no notes
    }
    {
        let s2 = builder.add_slide();
        s2.add_title("Slide C");
        s2.set_notes("Notes for C");
    }

    let mut pres = write_and_read(builder);

    let s0 = pres.slide(0).unwrap();
    let s1 = pres.slide(1).unwrap();
    let s2 = pres.slide(2).unwrap();

    assert!(s0.has_notes());
    assert_eq!(s0.notes().unwrap(), "Notes for A");

    assert!(!s1.has_notes());
    assert!(s1.notes().is_none());

    assert!(s2.has_notes());
    assert_eq!(s2.notes().unwrap(), "Notes for C");
}

// ---------------------------------------------------------------------------
// 22. `slides()` returns all slides
// ---------------------------------------------------------------------------

#[test]
fn test_slides_returns_all() {
    let mut builder = PresentationBuilder::new();
    for i in 0..5 {
        builder.add_slide().add_title(format!("Slide {}", i));
    }

    let mut pres = write_and_read(builder);
    let slides = pres.slides().unwrap();
    assert_eq!(slides.len(), 5);

    for (i, slide) in slides.iter().enumerate() {
        assert_eq!(slide.index(), i);
    }
}
