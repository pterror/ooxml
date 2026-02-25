// Custom handwritten WML fixtures.

use ooxml_wml::DocumentBuilder;

/// Minimal valid 1x1 white PNG image bytes.
const MINIMAL_PNG: &[u8] = &[
    0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A, // PNG signature
    0x00, 0x00, 0x00, 0x0D, 0x49, 0x48, 0x44, 0x52, // IHDR length + type
    0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x01, // 1x1
    0x08, 0x02, 0x00, 0x00, 0x00, 0x90, 0x77, 0x53, // 8-bit RGB
    0xDE, 0x00, 0x00, 0x00, 0x0C, 0x49, 0x44, 0x41, // IDAT
    0x54, 0x08, 0xD7, 0x63, 0xF8, 0xFF, 0xFF, 0x3F, 0x00, 0x05, 0xFE, 0x02, 0xFE, 0xDC, 0xCC, 0x59,
    0xE7, 0x00, 0x00, 0x00, 0x00, 0x49, 0x45, 0x4E, // IEND
    0x44, 0xAE, 0x42, 0x60, 0x82,
];

fn write_builder(builder: DocumentBuilder) -> Vec<u8> {
    let mut buf = std::io::Cursor::new(Vec::new());
    builder.write(&mut buf).expect("write failed");
    buf.into_inner()
}

// =============================================================================
// wml/text/ — run formatting
// =============================================================================

pub fn fixture_wml_text_underline() -> crate::Fixture {
    let mut builder = DocumentBuilder::new();
    let para = builder.body_mut().add_paragraph();
    let run = para.add_run();
    run.set_text("underlined");
    run.set_underline(ooxml_wml::types::STUnderline::Single);
    crate::Fixture {
        path: "wml/text/underline.docx",
        description: "Underlined run",
        bytes: write_builder(builder),
        assertions: vec![
            crate::Assertion::RunText {
                para: 0,
                run: 0,
                expected: "underlined".into(),
            },
            crate::Assertion::RunUnderline {
                para: 0,
                run: 0,
                expected: true,
            },
        ],
    }
}

pub fn fixture_wml_text_double_strike() -> crate::Fixture {
    let mut builder = DocumentBuilder::new();
    let para = builder.body_mut().add_paragraph();
    let run = para.add_run();
    run.set_text("double-struck");
    run.set_double_strike(true);
    crate::Fixture {
        path: "wml/text/double-strike.docx",
        description: "Double strikethrough run",
        bytes: write_builder(builder),
        assertions: vec![crate::Assertion::RunText {
            para: 0,
            run: 0,
            expected: "double-struck".into(),
        }],
    }
}

pub fn fixture_wml_text_small_caps() -> crate::Fixture {
    let mut builder = DocumentBuilder::new();
    let para = builder.body_mut().add_paragraph();
    let run = para.add_run();
    run.set_text("small caps");
    run.set_small_caps(true);
    crate::Fixture {
        path: "wml/text/small-caps.docx",
        description: "Small caps run",
        bytes: write_builder(builder),
        assertions: vec![crate::Assertion::RunText {
            para: 0,
            run: 0,
            expected: "small caps".into(),
        }],
    }
}

pub fn fixture_wml_text_all_caps() -> crate::Fixture {
    let mut builder = DocumentBuilder::new();
    let para = builder.body_mut().add_paragraph();
    let run = para.add_run();
    run.set_text("all caps");
    run.set_all_caps(true);
    crate::Fixture {
        path: "wml/text/all-caps.docx",
        description: "All caps run",
        bytes: write_builder(builder),
        assertions: vec![crate::Assertion::RunText {
            para: 0,
            run: 0,
            expected: "all caps".into(),
        }],
    }
}

pub fn fixture_wml_text_shadow() -> crate::Fixture {
    let mut builder = DocumentBuilder::new();
    let para = builder.body_mut().add_paragraph();
    let run = para.add_run();
    run.set_text("shadow");
    run.set_shadow(true);
    crate::Fixture {
        path: "wml/text/shadow.docx",
        description: "Shadow text effect run",
        bytes: write_builder(builder),
        assertions: vec![crate::Assertion::RunText {
            para: 0,
            run: 0,
            expected: "shadow".into(),
        }],
    }
}

pub fn fixture_wml_text_emboss() -> crate::Fixture {
    let mut builder = DocumentBuilder::new();
    let para = builder.body_mut().add_paragraph();
    let run = para.add_run();
    run.set_text("emboss");
    run.set_emboss(true);
    crate::Fixture {
        path: "wml/text/emboss.docx",
        description: "Emboss text effect run",
        bytes: write_builder(builder),
        assertions: vec![crate::Assertion::RunText {
            para: 0,
            run: 0,
            expected: "emboss".into(),
        }],
    }
}

pub fn fixture_wml_text_outline() -> crate::Fixture {
    let mut builder = DocumentBuilder::new();
    let para = builder.body_mut().add_paragraph();
    let run = para.add_run();
    run.set_text("outline");
    run.set_outline(true);
    crate::Fixture {
        path: "wml/text/outline.docx",
        description: "Outline (hollow) text effect run",
        bytes: write_builder(builder),
        assertions: vec![crate::Assertion::RunText {
            para: 0,
            run: 0,
            expected: "outline".into(),
        }],
    }
}

pub fn fixture_wml_text_imprint() -> crate::Fixture {
    let mut builder = DocumentBuilder::new();
    let para = builder.body_mut().add_paragraph();
    let run = para.add_run();
    run.set_text("imprint");
    run.set_imprint(true);
    crate::Fixture {
        path: "wml/text/imprint.docx",
        description: "Imprint (engrave) text effect run",
        bytes: write_builder(builder),
        assertions: vec![crate::Assertion::RunText {
            para: 0,
            run: 0,
            expected: "imprint".into(),
        }],
    }
}

pub fn fixture_wml_text_vanish() -> crate::Fixture {
    let mut builder = DocumentBuilder::new();
    let para = builder.body_mut().add_paragraph();
    let run = para.add_run();
    run.set_text("hidden");
    run.set_vanish(true);
    crate::Fixture {
        path: "wml/text/vanish.docx",
        description: "Hidden (vanish) run",
        bytes: write_builder(builder),
        assertions: vec![crate::Assertion::RunText {
            para: 0,
            run: 0,
            expected: "hidden".into(),
        }],
    }
}

pub fn fixture_wml_text_color() -> crate::Fixture {
    let mut builder = DocumentBuilder::new();
    let para = builder.body_mut().add_paragraph();
    let run = para.add_run();
    run.set_text("red text");
    run.set_color("FF0000");
    crate::Fixture {
        path: "wml/text/color.docx",
        description: "Red colored run (FF0000)",
        bytes: write_builder(builder),
        assertions: vec![
            crate::Assertion::RunText {
                para: 0,
                run: 0,
                expected: "red text".into(),
            },
            crate::Assertion::RunColor {
                para: 0,
                run: 0,
                expected: Some("FF0000".into()),
            },
        ],
    }
}

pub fn fixture_wml_text_font_size() -> crate::Fixture {
    let mut builder = DocumentBuilder::new();
    let para = builder.body_mut().add_paragraph();
    let run = para.add_run();
    run.set_text("big text");
    // 24pt = 48 half-points
    run.set_font_size(48);
    crate::Fixture {
        path: "wml/text/font-size.docx",
        description: "Run with 24pt font size",
        bytes: write_builder(builder),
        assertions: vec![
            crate::Assertion::RunText {
                para: 0,
                run: 0,
                expected: "big text".into(),
            },
            crate::Assertion::RunFontSize {
                para: 0,
                run: 0,
                expected: 24.0,
                tolerance: 0.5,
            },
        ],
    }
}

pub fn fixture_wml_text_font_name() -> crate::Fixture {
    let mut builder = DocumentBuilder::new();
    let para = builder.body_mut().add_paragraph();
    let run = para.add_run();
    run.set_text("Arial text");
    run.set_fonts(ooxml_wml::types::Fonts {
        ascii: Some("Arial".to_string()),
        h_ansi: Some("Arial".to_string()),
        ..Default::default()
    });
    crate::Fixture {
        path: "wml/text/font-name.docx",
        description: "Run with Arial font",
        bytes: write_builder(builder),
        assertions: vec![
            crate::Assertion::RunText {
                para: 0,
                run: 0,
                expected: "Arial text".into(),
            },
            crate::Assertion::RunFontName {
                para: 0,
                run: 0,
                expected: Some("Arial".into()),
            },
        ],
    }
}

pub fn fixture_wml_text_bold_italic() -> crate::Fixture {
    let mut builder = DocumentBuilder::new();
    let para = builder.body_mut().add_paragraph();
    let run = para.add_run();
    run.set_text("bold italic");
    run.set_bold(true);
    run.set_italic(true);
    crate::Fixture {
        path: "wml/text/bold-italic.docx",
        description: "Run with both bold and italic",
        bytes: write_builder(builder),
        assertions: vec![
            crate::Assertion::RunText {
                para: 0,
                run: 0,
                expected: "bold italic".into(),
            },
            crate::Assertion::RunBold {
                para: 0,
                run: 0,
                expected: true,
            },
            crate::Assertion::RunItalic {
                para: 0,
                run: 0,
                expected: true,
            },
        ],
    }
}

// =============================================================================
// wml/text/unicode/
// =============================================================================

pub fn fixture_wml_text_unicode_latin_extended() -> crate::Fixture {
    let text = "café résumé naïve";
    let mut builder = DocumentBuilder::new();
    builder.body_mut().add_paragraph().add_run().set_text(text);
    crate::Fixture {
        path: "wml/text/unicode/latin-extended.docx",
        description: "Run with Latin extended characters",
        bytes: write_builder(builder),
        assertions: vec![crate::Assertion::RunText {
            para: 0,
            run: 0,
            expected: text.into(),
        }],
    }
}

pub fn fixture_wml_text_unicode_cjk() -> crate::Fixture {
    let text = "中文日本語한국어";
    let mut builder = DocumentBuilder::new();
    builder.body_mut().add_paragraph().add_run().set_text(text);
    crate::Fixture {
        path: "wml/text/unicode/cjk.docx",
        description: "Run with CJK characters",
        bytes: write_builder(builder),
        assertions: vec![crate::Assertion::RunText {
            para: 0,
            run: 0,
            expected: text.into(),
        }],
    }
}

pub fn fixture_wml_text_unicode_arabic() -> crate::Fixture {
    let text = "مرحبا بالعالم";
    let mut builder = DocumentBuilder::new();
    builder.body_mut().add_paragraph().add_run().set_text(text);
    crate::Fixture {
        path: "wml/text/unicode/arabic.docx",
        description: "Run with Arabic text",
        bytes: write_builder(builder),
        assertions: vec![crate::Assertion::RunText {
            para: 0,
            run: 0,
            expected: text.into(),
        }],
    }
}

pub fn fixture_wml_text_unicode_emoji() -> crate::Fixture {
    let text = "Hello 🌍🎉";
    let mut builder = DocumentBuilder::new();
    builder.body_mut().add_paragraph().add_run().set_text(text);
    crate::Fixture {
        path: "wml/text/unicode/emoji.docx",
        description: "Run with emoji characters",
        bytes: write_builder(builder),
        assertions: vec![crate::Assertion::RunText {
            para: 0,
            run: 0,
            expected: text.into(),
        }],
    }
}

pub fn fixture_wml_text_xml_special_chars() -> crate::Fixture {
    let text = "& < > \" '";
    let mut builder = DocumentBuilder::new();
    builder.body_mut().add_paragraph().add_run().set_text(text);
    crate::Fixture {
        path: "wml/text/unicode/xml-special-chars.docx",
        description: "Run with XML special characters that must round-trip safely",
        bytes: write_builder(builder),
        assertions: vec![crate::Assertion::RunText {
            para: 0,
            run: 0,
            expected: text.into(),
        }],
    }
}

pub fn fixture_wml_text_empty_run() -> crate::Fixture {
    let mut builder = DocumentBuilder::new();
    builder.body_mut().add_paragraph().add_run().set_text("");
    crate::Fixture {
        path: "wml/text/unicode/empty-run.docx",
        description: "A run with empty text",
        bytes: write_builder(builder),
        assertions: vec![crate::Assertion::RunText {
            para: 0,
            run: 0,
            expected: "".into(),
        }],
    }
}

// =============================================================================
// wml/paragraph/
// =============================================================================

pub fn fixture_wml_paragraph_align_left() -> crate::Fixture {
    let mut builder = DocumentBuilder::new();
    let para = builder.body_mut().add_paragraph();
    para.set_alignment(ooxml_wml::types::STJc::Left);
    para.add_run().set_text("left");
    crate::Fixture {
        path: "wml/paragraph/align-left.docx",
        description: "Left-aligned paragraph",
        bytes: write_builder(builder),
        assertions: vec![
            crate::Assertion::ParagraphAlign {
                para: 0,
                expected: "left".into(),
            },
            crate::Assertion::RunText {
                para: 0,
                run: 0,
                expected: "left".into(),
            },
        ],
    }
}

pub fn fixture_wml_paragraph_align_center() -> crate::Fixture {
    let mut builder = DocumentBuilder::new();
    let para = builder.body_mut().add_paragraph();
    para.set_alignment(ooxml_wml::types::STJc::Center);
    para.add_run().set_text("centered");
    crate::Fixture {
        path: "wml/paragraph/align-center.docx",
        description: "Center-aligned paragraph",
        bytes: write_builder(builder),
        assertions: vec![
            crate::Assertion::ParagraphAlign {
                para: 0,
                expected: "center".into(),
            },
            crate::Assertion::RunText {
                para: 0,
                run: 0,
                expected: "centered".into(),
            },
        ],
    }
}

pub fn fixture_wml_paragraph_align_right() -> crate::Fixture {
    let mut builder = DocumentBuilder::new();
    let para = builder.body_mut().add_paragraph();
    para.set_alignment(ooxml_wml::types::STJc::Right);
    para.add_run().set_text("right");
    crate::Fixture {
        path: "wml/paragraph/align-right.docx",
        description: "Right-aligned paragraph",
        bytes: write_builder(builder),
        assertions: vec![
            crate::Assertion::ParagraphAlign {
                para: 0,
                expected: "right".into(),
            },
            crate::Assertion::RunText {
                para: 0,
                run: 0,
                expected: "right".into(),
            },
        ],
    }
}

pub fn fixture_wml_paragraph_align_justify() -> crate::Fixture {
    let mut builder = DocumentBuilder::new();
    let para = builder.body_mut().add_paragraph();
    para.set_alignment(ooxml_wml::types::STJc::Both);
    para.add_run().set_text("justified");
    crate::Fixture {
        path: "wml/paragraph/align-justify.docx",
        description: "Justified paragraph",
        bytes: write_builder(builder),
        assertions: vec![
            crate::Assertion::ParagraphAlign {
                para: 0,
                expected: "both".into(),
            },
            crate::Assertion::RunText {
                para: 0,
                run: 0,
                expected: "justified".into(),
            },
        ],
    }
}

pub fn fixture_wml_paragraph_space_before() -> crate::Fixture {
    let mut builder = DocumentBuilder::new();
    let para = builder.body_mut().add_paragraph();
    para.set_space_before(240);
    para.add_run().set_text("space before");
    crate::Fixture {
        path: "wml/paragraph/space-before.docx",
        description: "Paragraph with 240 twips space before",
        bytes: write_builder(builder),
        assertions: vec![crate::Assertion::RunText {
            para: 0,
            run: 0,
            expected: "space before".into(),
        }],
    }
}

pub fn fixture_wml_paragraph_space_after() -> crate::Fixture {
    let mut builder = DocumentBuilder::new();
    let para = builder.body_mut().add_paragraph();
    para.set_space_after(240);
    para.add_run().set_text("space after");
    crate::Fixture {
        path: "wml/paragraph/space-after.docx",
        description: "Paragraph with 240 twips space after",
        bytes: write_builder(builder),
        assertions: vec![crate::Assertion::RunText {
            para: 0,
            run: 0,
            expected: "space after".into(),
        }],
    }
}

pub fn fixture_wml_paragraph_line_spacing() -> crate::Fixture {
    let mut builder = DocumentBuilder::new();
    let para = builder.body_mut().add_paragraph();
    para.set_line_spacing(360);
    para.add_run().set_text("1.5x spacing");
    crate::Fixture {
        path: "wml/paragraph/line-spacing.docx",
        description: "Paragraph with 360 twips (1.5x) line spacing",
        bytes: write_builder(builder),
        assertions: vec![crate::Assertion::RunText {
            para: 0,
            run: 0,
            expected: "1.5x spacing".into(),
        }],
    }
}

pub fn fixture_wml_paragraph_indent_left() -> crate::Fixture {
    let mut builder = DocumentBuilder::new();
    let para = builder.body_mut().add_paragraph();
    para.set_indent_left(720);
    para.add_run().set_text("indented left");
    crate::Fixture {
        path: "wml/paragraph/indent-left.docx",
        description: "Paragraph with 720 twips left indent",
        bytes: write_builder(builder),
        assertions: vec![crate::Assertion::RunText {
            para: 0,
            run: 0,
            expected: "indented left".into(),
        }],
    }
}

pub fn fixture_wml_paragraph_indent_first_line() -> crate::Fixture {
    let mut builder = DocumentBuilder::new();
    let para = builder.body_mut().add_paragraph();
    para.set_indent_first_line(360);
    para.add_run().set_text("first line indent");
    crate::Fixture {
        path: "wml/paragraph/indent-first-line.docx",
        description: "Paragraph with 360 twips first-line indent",
        bytes: write_builder(builder),
        assertions: vec![crate::Assertion::RunText {
            para: 0,
            run: 0,
            expected: "first line indent".into(),
        }],
    }
}

pub fn fixture_wml_paragraph_outline_level_1() -> crate::Fixture {
    let mut builder = DocumentBuilder::new();
    let para = builder.body_mut().add_paragraph();
    para.set_outline_level(0); // level 0 = heading 1 in OOXML outline
    para.add_run().set_text("Heading 1");
    crate::Fixture {
        path: "wml/paragraph/outline-level-1.docx",
        description: "Paragraph at outline level 1",
        bytes: write_builder(builder),
        assertions: vec![crate::Assertion::RunText {
            para: 0,
            run: 0,
            expected: "Heading 1".into(),
        }],
    }
}

pub fn fixture_wml_paragraph_outline_level_2() -> crate::Fixture {
    let mut builder = DocumentBuilder::new();
    let para = builder.body_mut().add_paragraph();
    para.set_outline_level(1);
    para.add_run().set_text("Heading 2");
    crate::Fixture {
        path: "wml/paragraph/outline-level-2.docx",
        description: "Paragraph at outline level 2",
        bytes: write_builder(builder),
        assertions: vec![crate::Assertion::RunText {
            para: 0,
            run: 0,
            expected: "Heading 2".into(),
        }],
    }
}

pub fn fixture_wml_paragraph_outline_level_3() -> crate::Fixture {
    let mut builder = DocumentBuilder::new();
    let para = builder.body_mut().add_paragraph();
    para.set_outline_level(2);
    para.add_run().set_text("Heading 3");
    crate::Fixture {
        path: "wml/paragraph/outline-level-3.docx",
        description: "Paragraph at outline level 3",
        bytes: write_builder(builder),
        assertions: vec![crate::Assertion::RunText {
            para: 0,
            run: 0,
            expected: "Heading 3".into(),
        }],
    }
}

pub fn fixture_wml_paragraph_page_break() -> crate::Fixture {
    let mut builder = DocumentBuilder::new();
    builder.body_mut().add_paragraph().add_page_break();
    crate::Fixture {
        path: "wml/paragraph/page-break.docx",
        description: "Paragraph containing a page break",
        bytes: write_builder(builder),
        assertions: vec![crate::Assertion::ParagraphCount { expected: 1 }],
    }
}

pub fn fixture_wml_paragraph_column_break() -> crate::Fixture {
    let mut builder = DocumentBuilder::new();
    builder.body_mut().add_paragraph().add_column_break();
    crate::Fixture {
        path: "wml/paragraph/column-break.docx",
        description: "Paragraph containing a column break",
        bytes: write_builder(builder),
        assertions: vec![crate::Assertion::ParagraphCount { expected: 1 }],
    }
}

// =============================================================================
// wml/list/
// =============================================================================

pub fn fixture_wml_list_bullet_single() -> crate::Fixture {
    let mut builder = DocumentBuilder::new();
    let num_id = builder.add_list(ooxml_wml::writer::ListType::Bullet);
    let para = builder.body_mut().add_paragraph();
    para.set_numbering(num_id, 0);
    para.add_run().set_text("bullet item");
    crate::Fixture {
        path: "wml/list/bullet-single.docx",
        description: "Single-item bulleted list",
        bytes: write_builder(builder),
        assertions: vec![
            crate::Assertion::RunText {
                para: 0,
                run: 0,
                expected: "bullet item".into(),
            },
            crate::Assertion::ParagraphListLevel {
                para: 0,
                expected: Some(0),
            },
        ],
    }
}

pub fn fixture_wml_list_decimal_single() -> crate::Fixture {
    let mut builder = DocumentBuilder::new();
    let num_id = builder.add_list(ooxml_wml::writer::ListType::Decimal);
    let para = builder.body_mut().add_paragraph();
    para.set_numbering(num_id, 0);
    para.add_run().set_text("numbered item");
    crate::Fixture {
        path: "wml/list/decimal-single.docx",
        description: "Single-item numbered list",
        bytes: write_builder(builder),
        assertions: vec![
            crate::Assertion::RunText {
                para: 0,
                run: 0,
                expected: "numbered item".into(),
            },
            crate::Assertion::ParagraphListLevel {
                para: 0,
                expected: Some(0),
            },
        ],
    }
}

pub fn fixture_wml_list_bullet_nested_2() -> crate::Fixture {
    let mut builder = DocumentBuilder::new();
    let num_id = builder.add_list(ooxml_wml::writer::ListType::Bullet);
    {
        let body = builder.body_mut();
        let p1 = body.add_paragraph();
        p1.set_numbering(num_id, 0);
        p1.add_run().set_text("top level");
        let p2 = body.add_paragraph();
        p2.set_numbering(num_id, 1);
        p2.add_run().set_text("nested level");
    }
    crate::Fixture {
        path: "wml/list/bullet-nested-2.docx",
        description: "Two-level nested bullet list",
        bytes: write_builder(builder),
        assertions: vec![
            crate::Assertion::RunText {
                para: 0,
                run: 0,
                expected: "top level".into(),
            },
            crate::Assertion::RunText {
                para: 1,
                run: 0,
                expected: "nested level".into(),
            },
            crate::Assertion::ParagraphListLevel {
                para: 0,
                expected: Some(0),
            },
            crate::Assertion::ParagraphListLevel {
                para: 1,
                expected: Some(1),
            },
        ],
    }
}

pub fn fixture_wml_list_custom_numbering() -> crate::Fixture {
    let mut builder = DocumentBuilder::new();
    let num_id = builder.add_custom_list(vec![ooxml_wml::writer::NumberingLevel {
        ilvl: 0,
        format: ooxml_wml::writer::ListType::Decimal,
        start: 1,
        text: "%1.".to_string(),
        indent_left: Some(720),
        hanging: Some(360),
    }]);
    let para = builder.body_mut().add_paragraph();
    para.set_numbering(num_id, 0);
    para.add_run().set_text("custom item");
    crate::Fixture {
        path: "wml/list/custom-numbering.docx",
        description: "Custom decimal list with %1. format",
        bytes: write_builder(builder),
        assertions: vec![
            crate::Assertion::RunText {
                para: 0,
                run: 0,
                expected: "custom item".into(),
            },
            crate::Assertion::ParagraphListLevel {
                para: 0,
                expected: Some(0),
            },
        ],
    }
}

// =============================================================================
// wml/table/
// =============================================================================

pub fn fixture_wml_table_basic_1x1() -> crate::Fixture {
    let mut builder = DocumentBuilder::new();
    {
        let table = builder.body_mut().add_table();
        let row = table.add_row();
        let cell = row.add_cell();
        cell.add_paragraph().add_run().set_text("cell");
    }
    crate::Fixture {
        path: "wml/table/basic-1x1.docx",
        description: "Basic 1-row 1-col table",
        bytes: write_builder(builder),
        assertions: vec![
            crate::Assertion::TableRows {
                table: 0,
                expected: 1,
            },
            crate::Assertion::TableCols {
                table: 0,
                row: 0,
                expected: 1,
            },
            crate::Assertion::TableCellText {
                table: 0,
                row: 0,
                col: 0,
                expected: "cell".into(),
            },
        ],
    }
}

pub fn fixture_wml_table_basic_2x3() -> crate::Fixture {
    let mut builder = DocumentBuilder::new();
    {
        let table = builder.body_mut().add_table();
        for r in 0..2u32 {
            let row = table.add_row();
            for c in 0..3u32 {
                let cell = row.add_cell();
                let text = format!("r{r}c{c}");
                cell.add_paragraph().add_run().set_text(&text);
            }
        }
    }
    crate::Fixture {
        path: "wml/table/basic-2x3.docx",
        description: "Basic 2-row 3-col table",
        bytes: write_builder(builder),
        assertions: vec![
            crate::Assertion::TableRows {
                table: 0,
                expected: 2,
            },
            crate::Assertion::TableCols {
                table: 0,
                row: 0,
                expected: 3,
            },
            crate::Assertion::TableCellText {
                table: 0,
                row: 0,
                col: 0,
                expected: "r0c0".into(),
            },
            crate::Assertion::TableCellText {
                table: 0,
                row: 1,
                col: 2,
                expected: "r1c2".into(),
            },
        ],
    }
}

pub fn fixture_wml_table_merged_horizontal() -> crate::Fixture {
    let mut builder = DocumentBuilder::new();
    {
        let table = builder.body_mut().add_table();
        let row = table.add_row();
        // Cell spanning 2 columns
        let cell = row.add_cell();
        cell.set_grid_span(2);
        cell.add_paragraph().add_run().set_text("merged");
    }
    crate::Fixture {
        path: "wml/table/merged-horizontal.docx",
        description: "Table row with two columns merged horizontally",
        bytes: write_builder(builder),
        assertions: vec![
            crate::Assertion::TableRows {
                table: 0,
                expected: 1,
            },
            crate::Assertion::TableCellColspan {
                table: 0,
                row: 0,
                col: 0,
                expected: 2,
            },
            crate::Assertion::TableCellText {
                table: 0,
                row: 0,
                col: 0,
                expected: "merged".into(),
            },
        ],
    }
}

pub fn fixture_wml_table_merged_vertical() -> crate::Fixture {
    use ooxml_wml::convenience::VMergeType;
    let mut builder = DocumentBuilder::new();
    {
        let table = builder.body_mut().add_table();
        // Row 0: restart merge in col 0, regular cell in col 1
        {
            let row = table.add_row();
            let cell0 = row.add_cell();
            cell0.set_vertical_merge(VMergeType::Restart);
            cell0.add_paragraph().add_run().set_text("merged");
            let cell1 = row.add_cell();
            cell1.add_paragraph().add_run().set_text("r0c1");
        }
        // Row 1: continue merge in col 0, regular cell in col 1
        {
            let row = table.add_row();
            let cell0 = row.add_cell();
            cell0.set_vertical_merge(VMergeType::Continue);
            cell0.add_paragraph();
            let cell1 = row.add_cell();
            cell1.add_paragraph().add_run().set_text("r1c1");
        }
    }
    crate::Fixture {
        path: "wml/table/merged-vertical.docx",
        description: "Table column spanning two rows vertically",
        bytes: write_builder(builder),
        assertions: vec![
            crate::Assertion::TableRows {
                table: 0,
                expected: 2,
            },
            crate::Assertion::TableCellText {
                table: 0,
                row: 0,
                col: 0,
                expected: "merged".into(),
            },
        ],
    }
}

pub fn fixture_wml_table_cell_background() -> crate::Fixture {
    let mut builder = DocumentBuilder::new();
    {
        let table = builder.body_mut().add_table();
        let row = table.add_row();
        let cell = row.add_cell();
        cell.set_background_color("FFD700"); // gold
        cell.add_paragraph().add_run().set_text("gold cell");
    }
    crate::Fixture {
        path: "wml/table/cell-background.docx",
        description: "Table cell with background color",
        bytes: write_builder(builder),
        assertions: vec![crate::Assertion::TableCellText {
            table: 0,
            row: 0,
            col: 0,
            expected: "gold cell".into(),
        }],
    }
}

pub fn fixture_wml_table_cell_borders() -> crate::Fixture {
    use ooxml_wml::convenience::BorderStyle;
    let mut builder = DocumentBuilder::new();
    {
        let table = builder.body_mut().add_table();
        let row = table.add_row();
        let cell = row.add_cell();
        cell.set_borders(BorderStyle::Single, 4, "000000");
        cell.add_paragraph().add_run().set_text("bordered");
    }
    crate::Fixture {
        path: "wml/table/cell-borders.docx",
        description: "Table cell with explicit borders",
        bytes: write_builder(builder),
        assertions: vec![crate::Assertion::TableCellText {
            table: 0,
            row: 0,
            col: 0,
            expected: "bordered".into(),
        }],
    }
}

pub fn fixture_wml_table_cell_padding() -> crate::Fixture {
    let mut builder = DocumentBuilder::new();
    {
        let table = builder.body_mut().add_table();
        let row = table.add_row();
        let cell = row.add_cell();
        cell.set_padding(144, 144, 144, 144); // 144 twips = ~0.1 inch
        cell.add_paragraph().add_run().set_text("padded");
    }
    crate::Fixture {
        path: "wml/table/cell-padding.docx",
        description: "Table cell with custom padding",
        bytes: write_builder(builder),
        assertions: vec![crate::Assertion::TableCellText {
            table: 0,
            row: 0,
            col: 0,
            expected: "padded".into(),
        }],
    }
}

pub fn fixture_wml_table_row_height() -> crate::Fixture {
    let mut builder = DocumentBuilder::new();
    {
        let table = builder.body_mut().add_table();
        let row = table.add_row();
        row.set_height(720); // 720 twips = 0.5 inch
        let cell = row.add_cell();
        cell.add_paragraph().add_run().set_text("tall row");
    }
    crate::Fixture {
        path: "wml/table/row-height.docx",
        description: "Table row with explicit height",
        bytes: write_builder(builder),
        assertions: vec![crate::Assertion::TableCellText {
            table: 0,
            row: 0,
            col: 0,
            expected: "tall row".into(),
        }],
    }
}

pub fn fixture_wml_table_table_width() -> crate::Fixture {
    use ooxml_wml::convenience::TableWidthUnit;
    let mut builder = DocumentBuilder::new();
    {
        let table = builder.body_mut().add_table();
        table.set_width(8640, TableWidthUnit::Dxa); // 8640 twips = 6 inches
        let row = table.add_row();
        let cell = row.add_cell();
        cell.add_paragraph().add_run().set_text("wide table");
    }
    crate::Fixture {
        path: "wml/table/table-width.docx",
        description: "Table with explicit width",
        bytes: write_builder(builder),
        assertions: vec![crate::Assertion::TableCellText {
            table: 0,
            row: 0,
            col: 0,
            expected: "wide table".into(),
        }],
    }
}

// =============================================================================
// wml/hyperlink/
// =============================================================================

pub fn fixture_wml_hyperlink_external() -> crate::Fixture {
    let mut builder = DocumentBuilder::new();
    let rel_id = builder.add_hyperlink("https://example.com");
    {
        let para = builder.body_mut().add_paragraph();
        let link = para.add_hyperlink();
        link.set_rel_id(&rel_id);
        link.add_run().set_text("example.com");
    }
    crate::Fixture {
        path: "wml/hyperlink/hyperlink-external.docx",
        description: "Paragraph with external hyperlink",
        bytes: write_builder(builder),
        assertions: vec![crate::Assertion::HyperlinkUrl {
            para: 0,
            run: 0,
            expected: Some("https://example.com".into()),
        }],
    }
}

// =============================================================================
// wml/image/
// =============================================================================

pub fn fixture_wml_image_inline() -> crate::Fixture {
    let mut builder = DocumentBuilder::new();
    let rel_id = builder.add_image(MINIMAL_PNG.to_vec(), "image/png");
    // Build drawing before borrowing body_mut so there's no borrow conflict
    let mut drawing = ooxml_wml::writer::Drawing::new();
    drawing
        .add_image(&rel_id)
        .set_width_emu(914400)
        .set_height_emu(914400);
    let ct_drawing = builder.build_drawing(drawing);
    {
        let para = builder.body_mut().add_paragraph();
        let run = para.add_run();
        run.add_drawing(ct_drawing);
    }
    crate::Fixture {
        path: "wml/image/image-inline.docx",
        description: "Inline PNG image",
        bytes: write_builder(builder),
        assertions: vec![crate::Assertion::ImageCount { expected: 1 }],
    }
}

pub fn fixture_wml_image_anchored() -> crate::Fixture {
    let mut builder = DocumentBuilder::new();
    let rel_id = builder.add_image(MINIMAL_PNG.to_vec(), "image/png");
    {
        let mut drawing = ooxml_wml::writer::Drawing::new();
        drawing
            .add_anchored_image(&rel_id)
            .set_width_emu(914400)
            .set_height_emu(914400)
            .set_wrap_type(ooxml_wml::writer::WrapType::Square);
        let ct_drawing = builder.build_drawing(drawing);
        let para = builder.body_mut().add_paragraph();
        let run = para.add_run();
        run.add_drawing(ct_drawing);
    }
    crate::Fixture {
        path: "wml/image/image-anchored.docx",
        description: "Anchored (floating) PNG image",
        bytes: write_builder(builder),
        assertions: vec![crate::Assertion::ImageCount { expected: 1 }],
    }
}

// =============================================================================
// wml/header-footer/
// =============================================================================

pub fn fixture_wml_header_default() -> crate::Fixture {
    let mut builder = DocumentBuilder::new();
    builder
        .add_header(ooxml_wml::writer::HeaderFooterType::Default)
        .add_paragraph("Header text");
    builder
        .body_mut()
        .add_paragraph()
        .add_run()
        .set_text("body");
    crate::Fixture {
        path: "wml/header-footer/header-default.docx",
        description: "Document with a default header",
        bytes: write_builder(builder),
        assertions: vec![crate::Assertion::RunText {
            para: 0,
            run: 0,
            expected: "body".into(),
        }],
    }
}

pub fn fixture_wml_footer_default() -> crate::Fixture {
    let mut builder = DocumentBuilder::new();
    builder
        .add_footer(ooxml_wml::writer::HeaderFooterType::Default)
        .add_paragraph("Footer text");
    builder
        .body_mut()
        .add_paragraph()
        .add_run()
        .set_text("body");
    crate::Fixture {
        path: "wml/header-footer/footer-default.docx",
        description: "Document with a default footer",
        bytes: write_builder(builder),
        assertions: vec![crate::Assertion::RunText {
            para: 0,
            run: 0,
            expected: "body".into(),
        }],
    }
}

pub fn fixture_wml_header_first_page() -> crate::Fixture {
    let mut builder = DocumentBuilder::new();
    builder
        .add_header(ooxml_wml::writer::HeaderFooterType::First)
        .add_paragraph("First page header");
    builder
        .body_mut()
        .add_paragraph()
        .add_run()
        .set_text("body");
    crate::Fixture {
        path: "wml/header-footer/header-first-page.docx",
        description: "Document with a first-page header",
        bytes: write_builder(builder),
        assertions: vec![crate::Assertion::RunText {
            para: 0,
            run: 0,
            expected: "body".into(),
        }],
    }
}

// =============================================================================
// wml/footnote/
// =============================================================================

pub fn fixture_wml_footnote_basic() -> crate::Fixture {
    let mut builder = DocumentBuilder::new();
    let footnote_id = {
        let mut fb = builder.add_footnote();
        fb.add_paragraph("Footnote text");
        fb.id()
    };
    {
        let para = builder.body_mut().add_paragraph();
        let run = para.add_run();
        run.set_text("text");
        run.add_footnote_ref(footnote_id as i64);
    }
    crate::Fixture {
        path: "wml/footnote/footnote-basic.docx",
        description: "Paragraph with one footnote",
        bytes: write_builder(builder),
        assertions: vec![crate::Assertion::RunText {
            para: 0,
            run: 0,
            expected: "text".into(),
        }],
    }
}

pub fn fixture_wml_endnote_basic() -> crate::Fixture {
    let mut builder = DocumentBuilder::new();
    let endnote_id = {
        let mut eb = builder.add_endnote();
        eb.add_paragraph("Endnote text");
        eb.id()
    };
    {
        let para = builder.body_mut().add_paragraph();
        let run = para.add_run();
        run.set_text("text");
        run.add_endnote_ref(endnote_id as i64);
    }
    crate::Fixture {
        path: "wml/footnote/endnote-basic.docx",
        description: "Paragraph with one endnote",
        bytes: write_builder(builder),
        assertions: vec![crate::Assertion::RunText {
            para: 0,
            run: 0,
            expected: "text".into(),
        }],
    }
}

// =============================================================================
// wml/comment/
// =============================================================================

pub fn fixture_wml_comment_basic() -> crate::Fixture {
    let mut builder = DocumentBuilder::new();
    let comment_id = {
        let mut cb = builder.add_comment();
        cb.set_author("Author");
        cb.add_paragraph("Comment text");
        cb.id()
    };
    {
        let para = builder.body_mut().add_paragraph();
        para.add_comment_range_start(comment_id);
        let run = para.add_run();
        run.set_text("commented text");
        run.add_comment_ref(comment_id as i64);
        para.add_comment_range_end(comment_id);
    }
    crate::Fixture {
        path: "wml/comment/comment-basic.docx",
        description: "Paragraph with one comment",
        bytes: write_builder(builder),
        assertions: vec![crate::Assertion::RunText {
            para: 0,
            run: 0,
            expected: "commented text".into(),
        }],
    }
}

// =============================================================================
// wml/track-changes/
// =============================================================================

pub fn fixture_wml_track_insert() -> crate::Fixture {
    let mut builder = DocumentBuilder::new();
    {
        let para = builder.body_mut().add_paragraph();
        para.add_tracked_insertion(1, "Alice", Some("2026-02-24T12:00:00Z"), "inserted text");
    }
    crate::Fixture {
        path: "wml/track-changes/track-insert.docx",
        description: "Document with a tracked insertion",
        bytes: write_builder(builder),
        assertions: vec![crate::Assertion::ParagraphCount { expected: 1 }],
    }
}

pub fn fixture_wml_track_delete() -> crate::Fixture {
    let mut builder = DocumentBuilder::new();
    {
        let para = builder.body_mut().add_paragraph();
        para.add_tracked_deletion(1, "Bob", None, "deleted text");
    }
    crate::Fixture {
        path: "wml/track-changes/track-delete.docx",
        description: "Document with a tracked deletion",
        bytes: write_builder(builder),
        assertions: vec![crate::Assertion::ParagraphCount { expected: 1 }],
    }
}

// =============================================================================
// wml/field/
// =============================================================================

pub fn fixture_wml_field_toc_basic() -> crate::Fixture {
    let mut builder = DocumentBuilder::new();
    builder
        .body_mut()
        .add_toc(ooxml_wml::convenience::TocOptions::default());
    crate::Fixture {
        path: "wml/field/toc-basic.docx",
        description: "Document with a basic TOC field",
        bytes: write_builder(builder),
        // TOC produces multiple paragraphs; just check at least one exists
        assertions: vec![crate::Assertion::ParagraphCount { expected: 2 }],
    }
}

// =============================================================================
// wml/form/
// =============================================================================

pub fn fixture_wml_form_text() -> crate::Fixture {
    let mut builder = DocumentBuilder::new();
    builder
        .body_mut()
        .add_form_field(ooxml_wml::convenience::FormFieldConfig {
            tag: Some("text_field".into()),
            label: Some("Name".into()),
            field_type: ooxml_wml::convenience::FormFieldType::PlainText,
            default_value: Some("Enter name here".into()),
            placeholder: None,
            list_items: vec![],
            date_format: None,
        });
    crate::Fixture {
        path: "wml/form/form-text.docx",
        description: "Document with a plain-text form field",
        bytes: write_builder(builder),
        assertions: vec![],
    }
}

pub fn fixture_wml_form_dropdown() -> crate::Fixture {
    let mut builder = DocumentBuilder::new();
    builder
        .body_mut()
        .add_form_field(ooxml_wml::convenience::FormFieldConfig {
            tag: Some("dropdown_field".into()),
            label: Some("Color".into()),
            field_type: ooxml_wml::convenience::FormFieldType::DropDownList,
            default_value: Some("Red".into()),
            placeholder: None,
            list_items: vec!["Red".into(), "Green".into(), "Blue".into()],
            date_format: None,
        });
    crate::Fixture {
        path: "wml/form/form-dropdown.docx",
        description: "Document with a dropdown list form field with 3 options",
        bytes: write_builder(builder),
        assertions: vec![],
    }
}

// =============================================================================
// wml/math/
// =============================================================================

pub fn fixture_wml_math_plain() -> crate::Fixture {
    let mut builder = DocumentBuilder::new();
    {
        let para = builder.body_mut().add_paragraph();
        para.add_math(ooxml_wml::convenience::OMathBuilder::plain("x"));
    }
    crate::Fixture {
        path: "wml/math/math-plain.docx",
        description: "Paragraph with plain math expression",
        bytes: write_builder(builder),
        assertions: vec![crate::Assertion::ParagraphCount { expected: 1 }],
    }
}

pub fn fixture_wml_math_fraction() -> crate::Fixture {
    let mut builder = DocumentBuilder::new();
    {
        let para = builder.body_mut().add_paragraph();
        para.add_math(ooxml_wml::convenience::OMathBuilder::fraction("1", "2"));
    }
    crate::Fixture {
        path: "wml/math/math-fraction.docx",
        description: "Paragraph with a fraction math expression",
        bytes: write_builder(builder),
        assertions: vec![crate::Assertion::ParagraphCount { expected: 1 }],
    }
}

pub fn fixture_wml_math_superscript() -> crate::Fixture {
    let mut builder = DocumentBuilder::new();
    {
        let para = builder.body_mut().add_paragraph();
        para.add_math(ooxml_wml::convenience::OMathBuilder::superscript("x", "2"));
    }
    crate::Fixture {
        path: "wml/math/math-superscript.docx",
        description: "Paragraph with a superscript math expression (x^2)",
        bytes: write_builder(builder),
        assertions: vec![crate::Assertion::ParagraphCount { expected: 1 }],
    }
}

pub fn fixture_wml_math_subscript() -> crate::Fixture {
    let mut builder = DocumentBuilder::new();
    {
        let para = builder.body_mut().add_paragraph();
        para.add_math(ooxml_wml::convenience::OMathBuilder::subscript("x", "1"));
    }
    crate::Fixture {
        path: "wml/math/math-subscript.docx",
        description: "Paragraph with a subscript math expression (x_1)",
        bytes: write_builder(builder),
        assertions: vec![crate::Assertion::ParagraphCount { expected: 1 }],
    }
}

pub fn fixture_wml_math_radical() -> crate::Fixture {
    let mut builder = DocumentBuilder::new();
    {
        let para = builder.body_mut().add_paragraph();
        para.add_math(ooxml_wml::convenience::OMathBuilder::radical("2"));
    }
    crate::Fixture {
        path: "wml/math/math-radical.docx",
        description: "Paragraph with a square root math expression",
        bytes: write_builder(builder),
        assertions: vec![crate::Assertion::ParagraphCount { expected: 1 }],
    }
}

pub fn fixture_wml_math_display() -> crate::Fixture {
    let mut builder = DocumentBuilder::new();
    {
        let para = builder.body_mut().add_paragraph();
        para.add_math(ooxml_wml::convenience::OMathBuilder::fraction("a", "b").as_display());
    }
    crate::Fixture {
        path: "wml/math/math-display.docx",
        description: "Paragraph with a display-mode (block) fraction",
        bytes: write_builder(builder),
        assertions: vec![crate::Assertion::ParagraphCount { expected: 1 }],
    }
}

// =============================================================================
// wml/bookmark/
// =============================================================================

pub fn fixture_wml_bookmark_basic() -> crate::Fixture {
    let mut builder = DocumentBuilder::new();
    {
        let para = builder.body_mut().add_paragraph();
        para.add_bookmark_start_u32(0, "my_bookmark");
        para.add_run().set_text("bookmarked text");
        para.add_bookmark_end_u32(0);
    }
    crate::Fixture {
        path: "wml/bookmark/bookmark-basic.docx",
        description: "Paragraph with a bookmark start and end",
        bytes: write_builder(builder),
        assertions: vec![
            crate::Assertion::RunText {
                para: 0,
                run: 0,
                expected: "bookmarked text".into(),
            },
            crate::Assertion::BookmarkNames {
                expected: vec!["my_bookmark".into()],
            },
        ],
    }
}

// =============================================================================
// wml/settings/
// =============================================================================

pub fn fixture_wml_settings_tab_stop() -> crate::Fixture {
    let mut builder = DocumentBuilder::new();
    builder.set_settings(ooxml_wml::writer::DocumentSettingsOptions {
        default_tab_stop: Some(720),
        ..Default::default()
    });
    builder
        .body_mut()
        .add_paragraph()
        .add_run()
        .set_text("tab stop doc");
    crate::Fixture {
        path: "wml/settings/settings-tab-stop.docx",
        description: "Document with default tab stop set to 720 twips",
        bytes: write_builder(builder),
        assertions: vec![crate::Assertion::RunText {
            para: 0,
            run: 0,
            expected: "tab stop doc".into(),
        }],
    }
}

pub fn fixture_wml_settings_even_odd_headers() -> crate::Fixture {
    let mut builder = DocumentBuilder::new();
    builder.set_settings(ooxml_wml::writer::DocumentSettingsOptions {
        even_and_odd_headers: true,
        ..Default::default()
    });
    builder
        .body_mut()
        .add_paragraph()
        .add_run()
        .set_text("even odd doc");
    crate::Fixture {
        path: "wml/settings/settings-even-odd-headers.docx",
        description: "Document with even-and-odd headers enabled",
        bytes: write_builder(builder),
        assertions: vec![crate::Assertion::RunText {
            para: 0,
            run: 0,
            expected: "even odd doc".into(),
        }],
    }
}

// =============================================================================
// wml/text-box/
// =============================================================================

// =============================================================================
// wml/edge-case/ — structural edge cases
// =============================================================================

pub fn fixture_wml_edge_empty_document() -> crate::Fixture {
    let builder = DocumentBuilder::new();
    crate::Fixture {
        path: "wml/edge-case/empty-document.docx",
        description: "Document with no paragraphs or content",
        bytes: write_builder(builder),
        assertions: vec![crate::Assertion::ParagraphCount { expected: 0 }],
    }
}

pub fn fixture_wml_edge_empty_paragraph() -> crate::Fixture {
    let mut builder = DocumentBuilder::new();
    builder.body_mut().add_paragraph();
    crate::Fixture {
        path: "wml/edge-case/empty-paragraph.docx",
        description: "Document with a single empty paragraph (no runs)",
        bytes: write_builder(builder),
        assertions: vec![
            crate::Assertion::ParagraphCount { expected: 1 },
            crate::Assertion::ParagraphText {
                para: 0,
                expected: "".into(),
            },
        ],
    }
}

pub fn fixture_wml_edge_many_paragraphs() -> crate::Fixture {
    let mut builder = DocumentBuilder::new();
    for i in 0..100 {
        let para = builder.body_mut().add_paragraph();
        para.add_run().set_text(format!("Paragraph {i}"));
    }
    crate::Fixture {
        path: "wml/edge-case/many-paragraphs.docx",
        description: "Document with 100 paragraphs",
        bytes: write_builder(builder),
        assertions: vec![
            crate::Assertion::ParagraphCount { expected: 100 },
            crate::Assertion::ParagraphText {
                para: 0,
                expected: "Paragraph 0".into(),
            },
            crate::Assertion::ParagraphText {
                para: 99,
                expected: "Paragraph 99".into(),
            },
        ],
    }
}

pub fn fixture_wml_edge_mixed_formatting_paragraph() -> crate::Fixture {
    let mut builder = DocumentBuilder::new();
    {
        let para = builder.body_mut().add_paragraph();
        para.add_run().set_text("normal ");
        let bold_run = para.add_run();
        bold_run.set_text("bold ");
        bold_run.set_bold(true);
        let italic_run = para.add_run();
        italic_run.set_text("italic ");
        italic_run.set_italic(true);
        let both_run = para.add_run();
        both_run.set_text("both");
        both_run.set_bold(true);
        both_run.set_italic(true);
    }
    crate::Fixture {
        path: "wml/edge-case/mixed-formatting.docx",
        description: "Single paragraph with normal, bold, italic, and bold+italic runs",
        bytes: write_builder(builder),
        assertions: vec![
            crate::Assertion::ParagraphCount { expected: 1 },
            crate::Assertion::RunText {
                para: 0,
                run: 0,
                expected: "normal ".into(),
            },
            crate::Assertion::RunBold {
                para: 0,
                run: 0,
                expected: false,
            },
            crate::Assertion::RunText {
                para: 0,
                run: 1,
                expected: "bold ".into(),
            },
            crate::Assertion::RunBold {
                para: 0,
                run: 1,
                expected: true,
            },
            crate::Assertion::RunText {
                para: 0,
                run: 2,
                expected: "italic ".into(),
            },
            crate::Assertion::RunItalic {
                para: 0,
                run: 2,
                expected: true,
            },
            crate::Assertion::RunText {
                para: 0,
                run: 3,
                expected: "both".into(),
            },
            crate::Assertion::RunBold {
                para: 0,
                run: 3,
                expected: true,
            },
            crate::Assertion::RunItalic {
                para: 0,
                run: 3,
                expected: true,
            },
        ],
    }
}

pub fn fixture_wml_edge_page_break() -> crate::Fixture {
    let mut builder = DocumentBuilder::new();
    {
        let para = builder.body_mut().add_paragraph();
        para.add_run().set_text("Before break");
    }
    {
        let para = builder.body_mut().add_paragraph();
        para.add_page_break();
    }
    {
        let para = builder.body_mut().add_paragraph();
        para.add_run().set_text("After break");
    }
    crate::Fixture {
        path: "wml/edge-case/page-break.docx",
        description: "Document with a page break between two paragraphs",
        bytes: write_builder(builder),
        assertions: vec![crate::Assertion::ParagraphText {
            para: 0,
            expected: "Before break".into(),
        }],
    }
}

pub fn fixture_wml_edge_table_then_paragraph() -> crate::Fixture {
    let mut builder = DocumentBuilder::new();
    {
        let table = builder.body_mut().add_table();
        let row = table.add_row();
        let cell = row.add_cell();
        cell.add_paragraph().add_run().set_text("cell");
    }
    {
        let para = builder.body_mut().add_paragraph();
        para.add_run().set_text("after table");
    }
    crate::Fixture {
        path: "wml/edge-case/table-then-paragraph.docx",
        description: "Document with a table followed by a paragraph",
        bytes: write_builder(builder),
        assertions: vec![crate::Assertion::TableCellText {
            table: 0,
            row: 0,
            col: 0,
            expected: "cell".into(),
        }],
    }
}

pub fn fixture_wml_edge_multiple_tables() -> crate::Fixture {
    let mut builder = DocumentBuilder::new();
    for i in 0..3 {
        let table = builder.body_mut().add_table();
        let row = table.add_row();
        let cell = row.add_cell();
        cell.add_paragraph()
            .add_run()
            .set_text(format!("table {i}"));
    }
    crate::Fixture {
        path: "wml/edge-case/multiple-tables.docx",
        description: "Document with three consecutive tables",
        bytes: write_builder(builder),
        assertions: vec![
            crate::Assertion::TableCellText {
                table: 0,
                row: 0,
                col: 0,
                expected: "table 0".into(),
            },
            crate::Assertion::TableCellText {
                table: 1,
                row: 0,
                col: 0,
                expected: "table 1".into(),
            },
            crate::Assertion::TableCellText {
                table: 2,
                row: 0,
                col: 0,
                expected: "table 2".into(),
            },
        ],
    }
}

// =============================================================================
// wml/text-box/
// =============================================================================

pub fn fixture_wml_textbox_basic() -> crate::Fixture {
    let mut builder = DocumentBuilder::new();
    {
        let mut drawing = ooxml_wml::writer::Drawing::new();
        drawing
            .add_text_box("text box content")
            .set_width_inches(2.0)
            .set_height_inches(1.0);
        let ct_drawing = builder.build_drawing(drawing);
        let para = builder.body_mut().add_paragraph();
        let run = para.add_run();
        run.add_drawing(ct_drawing);
    }
    crate::Fixture {
        path: "wml/text-box/textbox-basic.docx",
        description: "Paragraph containing a Drawing text box",
        bytes: write_builder(builder),
        assertions: vec![crate::Assertion::ParagraphCount { expected: 1 }],
    }
}
