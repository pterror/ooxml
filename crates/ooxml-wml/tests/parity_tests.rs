//! Roundtrip tests: build → write → read using generated types + ext traits.
//!
//! These tests verify that documents created with `DocumentBuilder` (convenience API)
//! can be written to ZIP, read back with `Document::from_reader()` (generated parser),
//! and produce correct results through the ext trait API.

use ooxml_wml::ext::{
    BodyExt, CellExt, HyperlinkExt, ParagraphExt, RowExt, RunExt, RunPropertiesExt,
    SectionPropertiesExt, TableExt,
};
use ooxml_wml::{Document, DocumentBuilder};
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
// Tests
// =============================================================================

/// 3 plain text paragraphs — verify paragraph count and text.
#[test]
fn test_roundtrip_simple_document() {
    let mut builder = DocumentBuilder::new();
    builder.add_paragraph("First paragraph");
    builder.add_paragraph("Second paragraph");
    builder.add_paragraph("Third paragraph");

    let doc = roundtrip(builder);

    let paras = doc.body().paragraphs();
    assert_eq!(paras.len(), 3, "paragraph count");
    assert_eq!(paras[0].text(), "First paragraph");
    assert_eq!(paras[1].text(), "Second paragraph");
    assert_eq!(paras[2].text(), "Third paragraph");

    assert!(doc.body().text().contains("First paragraph"));
    assert!(doc.body().text().contains("Third paragraph"));
}

/// Runs with bold, italic, underline, strikethrough, font size, and color.
#[test]
fn test_roundtrip_formatted_text() {
    use ooxml_wml::types::STUnderline;

    let mut builder = DocumentBuilder::new();
    {
        let para = builder.body_mut().add_paragraph();

        let run = para.add_run();
        run.set_text("Bold");
        run.set_bold(true);

        let run = para.add_run();
        run.set_text("Italic");
        run.set_italic(true);

        let run = para.add_run();
        run.set_text("Underline");
        run.set_underline(STUnderline::Single);

        let run = para.add_run();
        run.set_text("Strike");
        run.set_strikethrough(true);

        let run = para.add_run();
        run.set_text("Sized");
        run.set_font_size(48);
        run.set_color("FF0000");
    }

    let doc = roundtrip(builder);
    let runs = doc.body().paragraphs()[0].runs();
    assert_eq!(runs.len(), 5, "run count");

    // Bold
    assert!(runs[0].is_bold());
    assert!(!runs[0].is_italic());

    // Italic
    assert!(!runs[1].is_bold());
    assert!(runs[1].is_italic());

    // Underline
    assert!(runs[2].is_underline());

    // Strikethrough
    assert!(runs[3].is_strikethrough());

    // Font size (half-points) and color
    assert_eq!(
        runs[4].properties().unwrap().font_size_half_points(),
        Some(48),
    );
    assert_eq!(runs[4].properties().unwrap().color_hex(), Some("FF0000"),);
}

/// 2x3 table with text in cells.
#[test]
fn test_roundtrip_basic_table() {
    let mut builder = DocumentBuilder::new();
    {
        let table = builder.body_mut().add_table();
        for r in 0..2 {
            let row = table.add_row();
            for c in 0..3 {
                let cell = row.add_cell();
                cell.add_paragraph().add_run().set_text(format!("R{r}C{c}"));
            }
        }
    }

    let doc = roundtrip(builder);
    let tables = doc.body().tables();
    assert_eq!(tables.len(), 1, "table count");

    let table = &tables[0];
    assert_eq!(table.row_count(), 2, "row count");

    for (r, row) in table.rows().iter().enumerate() {
        let cells = row.cells();
        assert_eq!(cells.len(), 3, "row {r} cell count");
        for (c, cell) in cells.iter().enumerate() {
            assert_eq!(cell.text(), format!("R{r}C{c}"), "row {r} cell {c}");
        }
    }
}

/// Paragraph with internal (anchor) hyperlink.
#[test]
fn test_roundtrip_hyperlinks() {
    let mut builder = DocumentBuilder::new();
    {
        let para = builder.body_mut().add_paragraph();
        let link = para.add_hyperlink();
        link.set_anchor("my_bookmark");
        link.add_run().set_text("Click here");
    }

    let doc = roundtrip(builder);
    let para = &doc.body().paragraphs()[0];
    let links = para.hyperlinks();

    assert_eq!(links.len(), 1, "hyperlink count");
    assert_eq!(links[0].text(), "Click here", "hyperlink text");
    assert_eq!(
        links[0].anchor_str(),
        Some("my_bookmark"),
        "hyperlink anchor"
    );
}

/// Run containing a page break.
#[test]
fn test_roundtrip_page_break() {
    let mut builder = DocumentBuilder::new();
    builder.add_paragraph("Before break");
    {
        let para = builder.body_mut().add_paragraph();
        para.add_run().set_page_break();
    }
    builder.add_paragraph("After break");

    let doc = roundtrip(builder);
    let paras = doc.body().paragraphs();

    assert_eq!(paras.len(), 3, "paragraph count");
    assert!(!paras[0].runs()[0].has_page_break());
    assert!(paras[1].runs()[0].has_page_break());
    assert!(!paras[2].runs()[0].has_page_break());
}

/// Multiple paragraphs with mixed runs: normal, bold, italic, bold+italic.
#[test]
fn test_roundtrip_multiple_paragraphs_with_formatting() {
    let mut builder = DocumentBuilder::new();

    // Paragraph 1: normal + bold
    {
        let para = builder.body_mut().add_paragraph();
        para.add_run().set_text("Normal ");
        let run = para.add_run();
        run.set_text("Bold");
        run.set_bold(true);
    }

    // Paragraph 2: italic + bold+italic
    {
        let para = builder.body_mut().add_paragraph();
        let run = para.add_run();
        run.set_text("Italic ");
        run.set_italic(true);
        let run = para.add_run();
        run.set_text("BoldItalic");
        run.set_bold(true);
        run.set_italic(true);
    }

    let doc = roundtrip(builder);
    let paras = doc.body().paragraphs();
    assert_eq!(paras.len(), 2);

    // Paragraph 1
    let p1_runs = paras[0].runs();
    assert_eq!(p1_runs.len(), 2);
    assert!(!p1_runs[0].is_bold());
    assert!(!p1_runs[0].is_italic());
    assert!(p1_runs[1].is_bold());
    assert!(!p1_runs[1].is_italic());

    // Paragraph 2
    let p2_runs = paras[1].runs();
    assert_eq!(p2_runs.len(), 2);
    assert!(!p2_runs[0].is_bold());
    assert!(p2_runs[0].is_italic());
    assert!(p2_runs[1].is_bold());
    assert!(p2_runs[1].is_italic());
}

/// Custom page size (landscape), margins.
#[cfg(feature = "wml-layout")]
#[test]
fn test_roundtrip_section_properties() {
    use ooxml_wml::types::{PageMargins, PageSize, STPageOrientation, SectionProperties};

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
        pg_mar: Some(Box::new(PageMargins {
            top: "720".to_string(),
            bottom: "720".to_string(),
            left: "1080".to_string(),
            right: "1080".to_string(),
            header: "360".to_string(),
            footer: "360".to_string(),
            gutter: "0".to_string(),
            #[cfg(feature = "extra-attrs")]
            extra_attrs: Default::default(),
        })),
        ..Default::default()
    };
    builder.body_mut().set_section_properties(sect_pr);

    let doc = roundtrip(builder);
    let sect = doc.body().section_properties().expect("section_properties");

    assert_eq!(sect.page_width_twips(), Some(15840), "page width twips");
    assert_eq!(sect.page_height_twips(), Some(12240), "page height twips");
    assert_eq!(sect.page_orientation(), Some(&STPageOrientation::Landscape));
}

/// Document with paragraphs + a table — text extraction includes table content.
#[test]
fn test_roundtrip_text_extraction() {
    let mut builder = DocumentBuilder::new();
    builder.add_paragraph("Above table");
    {
        let table = builder.body_mut().add_table();
        let row = table.add_row();
        row.add_cell().add_paragraph().add_run().set_text("Cell A");
        row.add_cell().add_paragraph().add_run().set_text("Cell B");
    }
    builder.add_paragraph("Below table");

    let doc = roundtrip(builder);

    // Direct paragraphs (not table cell paragraphs)
    let para_texts: Vec<_> = doc.body().paragraphs().iter().map(|p| p.text()).collect();
    assert_eq!(para_texts, vec!["Above table", "Below table"]);

    // Full text includes table content
    let text = doc.body().text();
    assert!(text.contains("Cell A"));
    assert!(text.contains("Cell B"));
}
