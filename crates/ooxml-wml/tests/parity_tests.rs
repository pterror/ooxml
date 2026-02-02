//! Parity tests: compare handwritten and generated parsers.
//!
//! These tests verify that `ext::parse_document` (generated parser) produces
//! semantically equivalent results to `Document::from_reader` (handwritten parser)
//! by building documents with `DocumentBuilder`, writing to in-memory ZIP, then
//! parsing with both paths and comparing results.
//!
//! # Remaining API-shape differences (not semantic disagreements)
//!
//! - **`Run.text()`**: Handwritten returns `&str` (single field);
//!   generated returns `String` (collects T/Tab/Br items). Same content.
//! - **`Paragraph.runs()`**: Handwritten includes insertion/deletion runs;
//!   generated only includes direct runs + hyperlink/field runs. Only matters
//!   for tracked-changes documents (not tested here).

use ooxml_wml::document::{PageMargins, PageOrientation, PageSize, SectionProperties};
use ooxml_wml::ext::{
    self, BodyExt, CellExt, DocumentExt, HyperlinkExt, ParagraphExt, RowExt, RunExt,
    RunPropertiesExt, SectionPropertiesExt, TableExt,
};
use ooxml_wml::{Document, DocumentBuilder, RunProperties};
use std::io::Cursor;

// =============================================================================
// Helpers
// =============================================================================

/// Build a document, return handwritten-parsed Document + raw document.xml bytes.
fn build_and_parse(builder: DocumentBuilder) -> (Document<Cursor<Vec<u8>>>, Vec<u8>) {
    let mut buffer = Cursor::new(Vec::new());
    builder.write(&mut buffer).unwrap();
    let bytes = buffer.into_inner();

    let mut doc = Document::from_reader(Cursor::new(bytes)).unwrap();
    let xml = doc.package_mut().read_part("word/document.xml").unwrap();

    (doc, xml)
}

/// Parse document.xml bytes with the generated parser.
fn parse_generated(xml: &[u8]) -> ooxml_wml::types::Document {
    ext::parse_document(xml).unwrap()
}

// =============================================================================
// Tests
// =============================================================================

/// 3 plain text paragraphs — compare paragraph count, text, and run counts.
#[test]
fn test_parity_simple_document() {
    let mut builder = DocumentBuilder::new();
    builder.add_paragraph("First paragraph");
    builder.add_paragraph("Second paragraph");
    builder.add_paragraph("Third paragraph");

    let (hw_doc, xml) = build_and_parse(builder);
    let gen_doc = parse_generated(&xml);

    let hw_paras = hw_doc.body().paragraphs();
    let gen_body = gen_doc.body().expect("generated body");
    let gen_paras = gen_body.paragraphs();

    assert_eq!(hw_paras.len(), gen_paras.len(), "paragraph count");

    for (i, (hw_p, gen_p)) in hw_paras.iter().zip(gen_paras.iter()).enumerate() {
        assert_eq!(hw_p.text(), gen_p.text(), "paragraph {i} text");
        assert_eq!(hw_p.runs().len(), gen_p.runs().len(), "paragraph {i} runs");
    }

    assert_eq!(hw_doc.body().text(), gen_body.text(), "full body text");
}

/// Runs with bold, italic, underline, strikethrough, font size, and color.
#[test]
fn test_parity_formatted_text() {
    use ooxml_wml::document::UnderlineStyle;

    let mut builder = DocumentBuilder::new();
    {
        let para = builder.body_mut().add_paragraph();

        let run = para.add_run();
        run.set_text("Bold");
        run.set_properties(RunProperties {
            bold: true,
            ..Default::default()
        });

        let run = para.add_run();
        run.set_text("Italic");
        run.set_properties(RunProperties {
            italic: true,
            ..Default::default()
        });

        let run = para.add_run();
        run.set_text("Underline");
        run.set_properties(RunProperties {
            underline: Some(UnderlineStyle::Single),
            ..Default::default()
        });

        let run = para.add_run();
        run.set_text("Strike");
        run.set_properties(RunProperties {
            strike: true,
            ..Default::default()
        });

        let run = para.add_run();
        run.set_text("Sized");
        run.set_properties(RunProperties {
            size: Some(48),
            color: Some("FF0000".to_string()),
            ..Default::default()
        });
    }

    let (hw_doc, xml) = build_and_parse(builder);
    let gen_doc = parse_generated(&xml);

    let hw_runs = hw_doc.body().paragraphs()[0].runs();
    let gen_body = gen_doc.body().unwrap();
    let gen_runs = gen_body.paragraphs()[0].runs();

    assert_eq!(hw_runs.len(), gen_runs.len(), "run count");

    // Bold
    assert_eq!(
        hw_runs[0].is_bold(),
        gen_runs[0].properties().unwrap().is_bold()
    );
    assert_eq!(
        hw_runs[0].is_italic(),
        gen_runs[0].properties().unwrap().is_italic()
    );

    // Italic
    assert_eq!(
        hw_runs[1].is_italic(),
        gen_runs[1].properties().unwrap().is_italic()
    );
    assert_eq!(
        hw_runs[1].is_bold(),
        gen_runs[1].properties().unwrap().is_bold()
    );

    // Underline
    assert_eq!(
        hw_runs[2].is_underline(),
        gen_runs[2].properties().unwrap().is_underline()
    );

    // Strikethrough
    assert_eq!(
        hw_runs[3].is_strikethrough(),
        gen_runs[3].properties().unwrap().is_strikethrough()
    );

    // Font size (half-points)
    assert_eq!(
        hw_runs[4].properties().and_then(|p| p.size),
        gen_runs[4].properties().unwrap().font_size_half_points(),
    );

    // Color (hex)
    assert_eq!(
        hw_runs[4].properties().and_then(|p| p.color.as_deref()),
        gen_runs[4].properties().unwrap().color_hex(),
    );
}

/// 2x3 table with text in cells.
#[test]
fn test_parity_basic_table() {
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

    let (hw_doc, xml) = build_and_parse(builder);
    let gen_doc = parse_generated(&xml);

    let hw_tables: Vec<_> = hw_doc.body().tables().collect();
    let gen_body = gen_doc.body().unwrap();
    let gen_tables = gen_body.tables();

    assert_eq!(hw_tables.len(), gen_tables.len(), "table count");

    let hw_table = hw_tables[0];
    let gen_table = gen_tables[0];

    assert_eq!(hw_table.rows().len(), gen_table.row_count(), "row count");

    for (ri, (hw_row, gen_row)) in hw_table
        .rows()
        .iter()
        .zip(gen_table.rows().iter())
        .enumerate()
    {
        let hw_cells = hw_row.cells();
        let gen_cells = gen_row.cells();
        assert_eq!(hw_cells.len(), gen_cells.len(), "row {ri} cell count");

        for (ci, (hw_cell, gen_cell)) in hw_cells.iter().zip(gen_cells.iter()).enumerate() {
            assert_eq!(hw_cell.text(), gen_cell.text(), "row {ri} cell {ci} text");
        }
    }
}

/// Paragraph with internal (anchor) hyperlink.
#[test]
fn test_parity_hyperlinks() {
    let mut builder = DocumentBuilder::new();
    {
        let para = builder.body_mut().add_paragraph();
        let link = para.add_hyperlink();
        link.set_anchor("my_bookmark");
        link.add_run().set_text("Click here");
    }

    let (hw_doc, xml) = build_and_parse(builder);
    let gen_doc = parse_generated(&xml);

    let hw_para = &hw_doc.body().paragraphs()[0];
    let gen_body = gen_doc.body().unwrap();
    let gen_para = &gen_body.paragraphs()[0];

    let hw_links: Vec<_> = hw_para.hyperlinks().collect();
    let gen_links = gen_para.hyperlinks();

    assert_eq!(hw_links.len(), gen_links.len(), "hyperlink count");
    assert_eq!(hw_links[0].text(), gen_links[0].text(), "hyperlink text");
    assert_eq!(
        hw_links[0].anchor().map(|s| s.to_string()),
        gen_links[0].anchor_str().map(|s| s.to_string()),
        "hyperlink anchor"
    );
}

/// Run containing a page break.
#[test]
fn test_parity_page_break() {
    let mut builder = DocumentBuilder::new();
    builder.add_paragraph("Before break");
    {
        let para = builder.body_mut().add_paragraph();
        para.add_run().set_page_break(true);
    }
    builder.add_paragraph("After break");

    let (hw_doc, xml) = build_and_parse(builder);
    let gen_doc = parse_generated(&xml);

    let hw_paras = hw_doc.body().paragraphs();
    let gen_body = gen_doc.body().unwrap();
    let gen_paras = gen_body.paragraphs();

    assert_eq!(hw_paras.len(), gen_paras.len(), "paragraph count");

    for (i, (hw_p, gen_p)) in hw_paras.iter().zip(gen_paras.iter()).enumerate() {
        if !hw_p.runs().is_empty() && !gen_p.runs().is_empty() {
            assert_eq!(
                hw_p.runs()[0].has_page_break(),
                gen_p.runs()[0].has_page_break(),
                "paragraph {i} page break"
            );
        }
    }

    // Specifically verify the break paragraph
    assert!(!hw_paras[0].runs()[0].has_page_break());
    assert!(hw_paras[1].runs()[0].has_page_break());
    assert!(!hw_paras[2].runs()[0].has_page_break());
}

/// Multiple paragraphs with mixed runs: normal, bold, italic, bold+italic.
#[test]
fn test_parity_multiple_paragraphs_with_formatting() {
    let mut builder = DocumentBuilder::new();

    // Paragraph 1: normal + bold
    {
        let para = builder.body_mut().add_paragraph();
        para.add_run().set_text("Normal ");
        let run = para.add_run();
        run.set_text("Bold");
        run.set_properties(RunProperties {
            bold: true,
            ..Default::default()
        });
    }

    // Paragraph 2: italic + bold+italic
    {
        let para = builder.body_mut().add_paragraph();
        let run = para.add_run();
        run.set_text("Italic ");
        run.set_properties(RunProperties {
            italic: true,
            ..Default::default()
        });
        let run = para.add_run();
        run.set_text("BoldItalic");
        run.set_properties(RunProperties {
            bold: true,
            italic: true,
            ..Default::default()
        });
    }

    let (hw_doc, xml) = build_and_parse(builder);
    let gen_doc = parse_generated(&xml);

    let hw_paras = hw_doc.body().paragraphs();
    let gen_body = gen_doc.body().unwrap();
    let gen_paras = gen_body.paragraphs();

    assert_eq!(hw_paras.len(), gen_paras.len());
    assert_eq!(hw_doc.body().text(), gen_body.text(), "full text");

    // Paragraph 1
    let hw_p1 = hw_paras[0].runs();
    let gen_p1 = gen_paras[0].runs();
    assert_eq!(hw_p1.len(), gen_p1.len());

    // Run 0: normal (no properties)
    assert!(!hw_p1[0].is_bold());
    assert!(!hw_p1[0].is_italic());
    assert!(!gen_p1[0].properties().is_some_and(|p| p.is_bold()));
    assert!(!gen_p1[0].properties().is_some_and(|p| p.is_italic()));

    // Run 1: bold only
    assert_eq!(
        hw_p1[1].is_bold(),
        gen_p1[1].properties().unwrap().is_bold()
    );
    assert_eq!(
        hw_p1[1].is_italic(),
        gen_p1[1].properties().unwrap().is_italic()
    );

    // Paragraph 2
    let hw_p2 = hw_paras[1].runs();
    let gen_p2 = gen_paras[1].runs();
    assert_eq!(hw_p2.len(), gen_p2.len());

    // Run 0: italic only
    assert_eq!(
        hw_p2[0].is_bold(),
        gen_p2[0].properties().unwrap().is_bold()
    );
    assert_eq!(
        hw_p2[0].is_italic(),
        gen_p2[0].properties().unwrap().is_italic()
    );

    // Run 1: bold+italic
    assert_eq!(
        hw_p2[1].is_bold(),
        gen_p2[1].properties().unwrap().is_bold()
    );
    assert_eq!(
        hw_p2[1].is_italic(),
        gen_p2[1].properties().unwrap().is_italic()
    );
}

/// Custom page size (landscape), margins.
#[cfg(feature = "wml-layout")]
#[test]
fn test_parity_section_properties() {
    use ooxml_wml::types::STPageOrientation;

    let mut builder = DocumentBuilder::new();
    builder.add_paragraph("Landscape page");

    builder
        .body_mut()
        .set_section_properties(SectionProperties {
            page_size: Some(PageSize {
                width: 15840,
                height: 12240,
                orientation: PageOrientation::Landscape,
            }),
            margins: Some(PageMargins {
                top: 720,
                bottom: 720,
                left: 1080,
                right: 1080,
                header: Some(360),
                footer: Some(360),
                gutter: Some(0),
            }),
            ..Default::default()
        });

    let (hw_doc, xml) = build_and_parse(builder);
    let gen_doc = parse_generated(&xml);

    let hw_sect = hw_doc
        .body()
        .section_properties()
        .expect("hw section_properties");
    let hw_pg = hw_sect.page_size.as_ref().expect("hw page_size");

    let gen_body = gen_doc.body().unwrap();
    let gen_sect = gen_body
        .section_properties()
        .expect("gen section_properties");

    assert_eq!(
        hw_pg.width as u64,
        gen_sect.page_width_twips().expect("gen width"),
        "page width twips"
    );
    assert_eq!(
        hw_pg.height as u64,
        gen_sect.page_height_twips().expect("gen height"),
        "page height twips"
    );
    assert_eq!(hw_pg.orientation, PageOrientation::Landscape);
    assert_eq!(
        gen_sect.page_orientation(),
        Some(&STPageOrientation::Landscape)
    );
}

/// Document with paragraphs + a table — both parsers should agree on
/// paragraph list, table structure, and full text extraction.
#[test]
fn test_parity_text_extraction() {
    let mut builder = DocumentBuilder::new();
    builder.add_paragraph("Above table");
    {
        let table = builder.body_mut().add_table();
        let row = table.add_row();
        row.add_cell().add_paragraph().add_run().set_text("Cell A");
        row.add_cell().add_paragraph().add_run().set_text("Cell B");
    }
    builder.add_paragraph("Below table");

    let (hw_doc, xml) = build_and_parse(builder);
    let gen_doc = parse_generated(&xml);
    let gen_body = gen_doc.body().unwrap();

    // Both sides return only direct paragraphs (not table cell paragraphs)
    let hw_para_texts: Vec<_> = hw_doc
        .body()
        .paragraphs()
        .iter()
        .map(|p| p.text())
        .collect();
    let gen_para_texts: Vec<_> = gen_body.paragraphs().iter().map(|p| p.text()).collect();
    assert_eq!(hw_para_texts, gen_para_texts, "paragraph texts");
    assert_eq!(hw_para_texts, vec!["Above table", "Below table"]);

    // Both sides include table text in Body.text()
    let hw_text = hw_doc.body().text();
    let gen_text = gen_body.text();
    assert_eq!(hw_text, gen_text, "full body text");
    assert!(hw_text.contains("Cell A"), "should include table text");
    assert!(hw_text.contains("Cell B"), "should include table text");
}
