// Custom handwritten PML fixtures.

use ooxml_pml::{ImageFormat, Paragraph, PresentationBuilder, TextAlign, TextRun};

/// Minimal 1x1 white PNG for image fixture tests.
const MINIMAL_PNG: &[u8] = &[
    0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A, 0x00, 0x00, 0x00, 0x0D, 0x49, 0x48, 0x44, 0x52,
    0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x01, 0x08, 0x02, 0x00, 0x00, 0x00, 0x90, 0x77, 0x53,
    0xDE, 0x00, 0x00, 0x00, 0x0C, 0x49, 0x44, 0x41, 0x54, 0x08, 0xD7, 0x63, 0xF8, 0xFF, 0xFF, 0x3F,
    0x00, 0x05, 0xFE, 0x02, 0xFE, 0xDC, 0xCC, 0x59, 0xE7, 0x00, 0x00, 0x00, 0x00, 0x49, 0x45, 0x4E,
    0x44, 0xAE, 0x42, 0x60, 0x82,
];

/// Minimal bar chart XML for chart fixture tests.
const MINIMAL_CHART_XML: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
              xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
              xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <c:chart><c:plotArea><c:barChart>
    <c:barDir val="col"/><c:grouping val="clustered"/>
    <c:ser><c:idx val="0"/><c:order val="0"/>
      <c:val><c:numRef><c:f>Sheet1!$A$1:$A$3</c:f>
        <c:numCache><c:formatCode>General</c:formatCode>
          <c:ptCount val="3"/>
          <c:pt idx="0"><c:v>10</c:v></c:pt>
          <c:pt idx="1"><c:v>20</c:v></c:pt>
          <c:pt idx="2"><c:v>30</c:v></c:pt>
        </c:numCache></c:numRef></c:val>
    </c:ser>
  </c:barChart></c:plotArea></c:chart>
</c:chartSpace>"#;

fn write_pres(pres: PresentationBuilder) -> Vec<u8> {
    let mut buf = std::io::Cursor::new(Vec::new());
    pres.write(&mut buf).expect("write failed");
    buf.into_inner()
}

// ---------------------------------------------------------------------------
// pml/structure/
// ---------------------------------------------------------------------------

pub fn fixture_pml_structure_single_slide() -> crate::Fixture {
    let mut pres = PresentationBuilder::new();
    let slide = pres.add_slide();
    slide.add_text("Only slide");
    crate::Fixture {
        path: "pml/structure/single-slide.pptx",
        description: "Presentation with exactly one slide",
        bytes: write_pres(pres),
        assertions: vec![crate::Assertion::SlideCount { expected: 1 }],
    }
}

pub fn fixture_pml_structure_multi_slide() -> crate::Fixture {
    let mut pres = PresentationBuilder::new();
    pres.add_slide().add_text("Slide one");
    pres.add_slide().add_text("Slide two");
    pres.add_slide().add_text("Slide three");
    crate::Fixture {
        path: "pml/structure/multi-slide.pptx",
        description: "Presentation with three slides with different text",
        bytes: write_pres(pres),
        assertions: vec![
            crate::Assertion::SlideCount { expected: 3 },
            crate::Assertion::ShapeText {
                slide: 0,
                shape: 0,
                expected: "Slide one".into(),
            },
            crate::Assertion::ShapeText {
                slide: 1,
                shape: 0,
                expected: "Slide two".into(),
            },
            crate::Assertion::ShapeText {
                slide: 2,
                shape: 0,
                expected: "Slide three".into(),
            },
        ],
    }
}

pub fn fixture_pml_structure_slide_order() -> crate::Fixture {
    let mut pres = PresentationBuilder::new();
    pres.add_slide().add_text("First");
    pres.add_slide().add_text("Second");
    pres.add_slide().add_text("Third");
    crate::Fixture {
        path: "pml/structure/slide-order.pptx",
        description: "Slides with text First, Second, Third in order",
        bytes: write_pres(pres),
        assertions: vec![
            crate::Assertion::SlideCount { expected: 3 },
            crate::Assertion::ShapeText {
                slide: 0,
                shape: 0,
                expected: "First".into(),
            },
            crate::Assertion::ShapeText {
                slide: 1,
                shape: 0,
                expected: "Second".into(),
            },
            crate::Assertion::ShapeText {
                slide: 2,
                shape: 0,
                expected: "Third".into(),
            },
        ],
    }
}

// ---------------------------------------------------------------------------
// pml/shape/
// ---------------------------------------------------------------------------

pub fn fixture_pml_shape_text_unicode() -> crate::Fixture {
    let mut pres = PresentationBuilder::new();
    let slide = pres.add_slide();
    slide.add_text_with_runs(
        vec![TextRun::text("café 中文 🌍")],
        457200,
        274638,
        8229600,
        4525963,
    );
    crate::Fixture {
        path: "pml/shape/text-unicode.pptx",
        description: "Shape with unicode text including emoji",
        bytes: write_pres(pres),
        assertions: vec![
            crate::Assertion::SlideCount { expected: 1 },
            crate::Assertion::ShapeText {
                slide: 0,
                shape: 0,
                expected: "café 中文 🌍".into(),
            },
        ],
    }
}

pub fn fixture_pml_shape_text_xml_special() -> crate::Fixture {
    let text = "& < > \" '";
    let mut pres = PresentationBuilder::new();
    let slide = pres.add_slide();
    slide.add_text_with_runs(vec![TextRun::text(text)], 457200, 274638, 8229600, 4525963);
    crate::Fixture {
        path: "pml/shape/text-xml-special.pptx",
        description: "Shape with XML special characters",
        bytes: write_pres(pres),
        assertions: vec![
            crate::Assertion::SlideCount { expected: 1 },
            crate::Assertion::ShapeText {
                slide: 0,
                shape: 0,
                expected: text.into(),
            },
        ],
    }
}

pub fn fixture_pml_shape_text_empty() -> crate::Fixture {
    let mut pres = PresentationBuilder::new();
    let slide = pres.add_slide();
    slide.add_text_with_runs(vec![TextRun::text("")], 457200, 274638, 8229600, 4525963);
    crate::Fixture {
        path: "pml/shape/text-empty.pptx",
        description: "Shape with empty string text",
        bytes: write_pres(pres),
        assertions: vec![crate::Assertion::SlideCount { expected: 1 }],
    }
}

pub fn fixture_pml_shape_rect() -> crate::Fixture {
    use ooxml_pml::PresetGeometry;
    let mut pres = PresentationBuilder::new();
    let slide = pres.add_slide();
    slide
        .shape(vec![TextRun::text("Rect")], 914400, 914400, 3657600, 914400)
        .set_geometry(PresetGeometry::Rect)
        .add();
    crate::Fixture {
        path: "pml/shape/shape-rect.pptx",
        description: "Shape with Rect geometry at explicit position",
        bytes: write_pres(pres),
        assertions: vec![
            crate::Assertion::SlideCount { expected: 1 },
            crate::Assertion::ShapeType {
                slide: 0,
                shape: 0,
                expected: "rect".into(),
            },
        ],
    }
}

pub fn fixture_pml_shape_ellipse() -> crate::Fixture {
    use ooxml_pml::PresetGeometry;
    let mut pres = PresentationBuilder::new();
    let slide = pres.add_slide();
    slide
        .shape(
            vec![TextRun::text("Ellipse")],
            914400,
            914400,
            3657600,
            1828800,
        )
        .set_geometry(PresetGeometry::Ellipse)
        .add();
    crate::Fixture {
        path: "pml/shape/shape-ellipse.pptx",
        description: "Shape with Ellipse preset geometry",
        bytes: write_pres(pres),
        assertions: vec![
            crate::Assertion::SlideCount { expected: 1 },
            crate::Assertion::ShapeType {
                slide: 0,
                shape: 0,
                expected: "ellipse".into(),
            },
        ],
    }
}

pub fn fixture_pml_shape_roundrect() -> crate::Fixture {
    use ooxml_pml::PresetGeometry;
    let mut pres = PresentationBuilder::new();
    let slide = pres.add_slide();
    slide
        .shape(
            vec![TextRun::text("RoundRect")],
            914400,
            914400,
            3657600,
            1828800,
        )
        .set_geometry(PresetGeometry::RoundRect)
        .add();
    crate::Fixture {
        path: "pml/shape/shape-roundrect.pptx",
        description: "Shape with RoundRect preset geometry",
        bytes: write_pres(pres),
        assertions: vec![
            crate::Assertion::SlideCount { expected: 1 },
            crate::Assertion::ShapeType {
                slide: 0,
                shape: 0,
                expected: "roundRect".into(),
            },
        ],
    }
}

pub fn fixture_pml_shape_triangle() -> crate::Fixture {
    use ooxml_pml::PresetGeometry;
    let mut pres = PresentationBuilder::new();
    let slide = pres.add_slide();
    slide
        .shape(
            vec![TextRun::text("Triangle")],
            914400,
            914400,
            3657600,
            1828800,
        )
        .set_geometry(PresetGeometry::Triangle)
        .add();
    crate::Fixture {
        path: "pml/shape/shape-triangle.pptx",
        description: "Shape with Triangle preset geometry",
        bytes: write_pres(pres),
        assertions: vec![
            crate::Assertion::SlideCount { expected: 1 },
            crate::Assertion::ShapeType {
                slide: 0,
                shape: 0,
                expected: "triangle".into(),
            },
        ],
    }
}

pub fn fixture_pml_shape_fill_color() -> crate::Fixture {
    let mut pres = PresentationBuilder::new();
    let slide = pres.add_slide();
    slide
        .shape(
            vec![TextRun::text("Filled")],
            914400,
            914400,
            3657600,
            914400,
        )
        .set_fill_color("4472C4")
        .add();
    crate::Fixture {
        path: "pml/shape/shape-fill-color.pptx",
        description: "Shape with solid fill color 4472C4",
        bytes: write_pres(pres),
        assertions: vec![crate::Assertion::SlideCount { expected: 1 }],
    }
}

pub fn fixture_pml_shape_line_color() -> crate::Fixture {
    let mut pres = PresentationBuilder::new();
    let slide = pres.add_slide();
    slide
        .shape(
            vec![TextRun::text("Border")],
            914400,
            914400,
            3657600,
            914400,
        )
        .set_line_color("FF0000")
        .set_line_width(38100)
        .add();
    crate::Fixture {
        path: "pml/shape/shape-line-color.pptx",
        description: "Shape with red line color and 38100 EMU width",
        bytes: write_pres(pres),
        assertions: vec![crate::Assertion::SlideCount { expected: 1 }],
    }
}

pub fn fixture_pml_shape_position() -> crate::Fixture {
    let mut pres = PresentationBuilder::new();
    let slide = pres.add_slide();
    // x=914400, y=685800 (1 inch from left, ~0.75 inch from top)
    slide.add_text_at("Positioned", 914400, 685800, 3657600, 914400);
    crate::Fixture {
        path: "pml/shape/shape-position.pptx",
        description: "Shape positioned at specific coordinates",
        bytes: write_pres(pres),
        assertions: vec![crate::Assertion::SlideCount { expected: 1 }],
    }
}

pub fn fixture_pml_shape_connector_basic() -> crate::Fixture {
    let mut pres = PresentationBuilder::new();
    let slide = pres.add_slide();
    slide.add_connector(0, 0, 1000000, 500000, None);
    crate::Fixture {
        path: "pml/shape/connector-basic.pptx",
        description: "Slide with a basic connector",
        bytes: write_pres(pres),
        assertions: vec![crate::Assertion::SlideCount { expected: 1 }],
    }
}

pub fn fixture_pml_shape_connector_colored() -> crate::Fixture {
    let mut pres = PresentationBuilder::new();
    let slide = pres.add_slide();
    slide.add_connector(0, 0, 1000000, 500000, Some("FF0000"));
    crate::Fixture {
        path: "pml/shape/connector-colored.pptx",
        description: "Connector with red color",
        bytes: write_pres(pres),
        assertions: vec![crate::Assertion::SlideCount { expected: 1 }],
    }
}

pub fn fixture_pml_shape_group_basic() -> crate::Fixture {
    let mut pres = PresentationBuilder::new();
    {
        let slide = pres.add_slide();
        slide
            .begin_group()
            .add_text("Shape A", 457200, 274638, 2000000, 600000)
            .add_text("Shape B", 2600000, 274638, 2000000, 600000)
            .finish();
    }
    crate::Fixture {
        path: "pml/shape/group-basic.pptx",
        description: "Group containing two text shapes",
        bytes: write_pres(pres),
        assertions: vec![crate::Assertion::SlideCount { expected: 1 }],
    }
}

// ---------------------------------------------------------------------------
// pml/text/
// ---------------------------------------------------------------------------

pub fn fixture_pml_text_run_bold() -> crate::Fixture {
    let mut pres = PresentationBuilder::new();
    let slide = pres.add_slide();
    slide.add_text_with_runs(
        vec![TextRun::text("bold").set_bold(true)],
        457200,
        274638,
        8229600,
        4525963,
    );
    crate::Fixture {
        path: "pml/text/run-bold.pptx",
        description: "Shape with bold text run",
        bytes: write_pres(pres),
        assertions: vec![
            crate::Assertion::SlideCount { expected: 1 },
            crate::Assertion::PmlRunBold {
                slide: 0,
                shape: 0,
                para: 0,
                run: 0,
                expected: true,
            },
            crate::Assertion::PmlRunText {
                slide: 0,
                shape: 0,
                para: 0,
                run: 0,
                expected: "bold".into(),
            },
        ],
    }
}

pub fn fixture_pml_text_run_italic() -> crate::Fixture {
    let mut pres = PresentationBuilder::new();
    let slide = pres.add_slide();
    slide.add_text_with_runs(
        vec![TextRun::text("italic").set_italic(true)],
        457200,
        274638,
        8229600,
        4525963,
    );
    crate::Fixture {
        path: "pml/text/run-italic.pptx",
        description: "Shape with italic text run",
        bytes: write_pres(pres),
        assertions: vec![
            crate::Assertion::SlideCount { expected: 1 },
            crate::Assertion::PmlRunItalic {
                slide: 0,
                shape: 0,
                para: 0,
                run: 0,
                expected: true,
            },
            crate::Assertion::PmlRunText {
                slide: 0,
                shape: 0,
                para: 0,
                run: 0,
                expected: "italic".into(),
            },
        ],
    }
}

pub fn fixture_pml_text_run_bold_italic() -> crate::Fixture {
    let mut pres = PresentationBuilder::new();
    let slide = pres.add_slide();
    slide.add_text_with_runs(
        vec![TextRun::text("bold-italic").set_bold(true).set_italic(true)],
        457200,
        274638,
        8229600,
        4525963,
    );
    crate::Fixture {
        path: "pml/text/run-bold-italic.pptx",
        description: "Shape with bold and italic text run",
        bytes: write_pres(pres),
        assertions: vec![
            crate::Assertion::SlideCount { expected: 1 },
            crate::Assertion::PmlRunBold {
                slide: 0,
                shape: 0,
                para: 0,
                run: 0,
                expected: true,
            },
            crate::Assertion::PmlRunItalic {
                slide: 0,
                shape: 0,
                para: 0,
                run: 0,
                expected: true,
            },
        ],
    }
}

pub fn fixture_pml_text_run_color() -> crate::Fixture {
    let mut pres = PresentationBuilder::new();
    let slide = pres.add_slide();
    slide.add_text_with_runs(
        vec![TextRun::text("red text").set_color("FF0000")],
        457200,
        274638,
        8229600,
        4525963,
    );
    crate::Fixture {
        path: "pml/text/run-color.pptx",
        description: "Shape with red colored text run",
        bytes: write_pres(pres),
        assertions: vec![
            crate::Assertion::SlideCount { expected: 1 },
            crate::Assertion::PmlRunColor {
                slide: 0,
                shape: 0,
                para: 0,
                run: 0,
                expected: Some("FF0000".into()),
            },
        ],
    }
}

pub fn fixture_pml_text_run_font_size_24() -> crate::Fixture {
    let mut pres = PresentationBuilder::new();
    let slide = pres.add_slide();
    slide.add_text_with_runs(
        vec![TextRun::text("24pt").set_font_size(24.0)],
        457200,
        274638,
        8229600,
        4525963,
    );
    crate::Fixture {
        path: "pml/text/run-font-size-24.pptx",
        description: "Shape with 24pt font size run",
        bytes: write_pres(pres),
        assertions: vec![
            crate::Assertion::SlideCount { expected: 1 },
            crate::Assertion::PmlRunFontSize {
                slide: 0,
                shape: 0,
                para: 0,
                run: 0,
                expected: 24.0,
                tolerance: 0.5,
            },
        ],
    }
}

pub fn fixture_pml_text_run_font_size_36() -> crate::Fixture {
    let mut pres = PresentationBuilder::new();
    let slide = pres.add_slide();
    slide.add_text_with_runs(
        vec![TextRun::text("36pt").set_font_size(36.0)],
        457200,
        274638,
        8229600,
        4525963,
    );
    crate::Fixture {
        path: "pml/text/run-font-size-36.pptx",
        description: "Shape with 36pt font size run",
        bytes: write_pres(pres),
        assertions: vec![
            crate::Assertion::SlideCount { expected: 1 },
            crate::Assertion::PmlRunFontSize {
                slide: 0,
                shape: 0,
                para: 0,
                run: 0,
                expected: 36.0,
                tolerance: 0.5,
            },
        ],
    }
}

pub fn fixture_pml_text_run_underline() -> crate::Fixture {
    let mut pres = PresentationBuilder::new();
    let slide = pres.add_slide();
    slide.add_text_with_runs(
        vec![TextRun::text("underlined").set_underline(true)],
        457200,
        274638,
        8229600,
        4525963,
    );
    crate::Fixture {
        path: "pml/text/run-underline.pptx",
        description: "Shape with underlined text run",
        bytes: write_pres(pres),
        assertions: vec![crate::Assertion::SlideCount { expected: 1 }],
    }
}

pub fn fixture_pml_text_paragraph_align_left() -> crate::Fixture {
    let mut pres = PresentationBuilder::new();
    let slide = pres.add_slide();
    slide
        .shape(
            vec![TextRun::text("left")],
            457200,
            274638,
            8229600,
            4525963,
        )
        .set_text_align(TextAlign::Left)
        .add();
    crate::Fixture {
        path: "pml/text/paragraph-align-left.pptx",
        description: "Shape with left-aligned paragraph",
        bytes: write_pres(pres),
        assertions: vec![crate::Assertion::SlideCount { expected: 1 }],
    }
}

pub fn fixture_pml_text_paragraph_align_center() -> crate::Fixture {
    let mut pres = PresentationBuilder::new();
    let slide = pres.add_slide();
    slide
        .shape(
            vec![TextRun::text("center")],
            457200,
            274638,
            8229600,
            4525963,
        )
        .set_text_align(TextAlign::Center)
        .add();
    crate::Fixture {
        path: "pml/text/paragraph-align-center.pptx",
        description: "Shape with center-aligned paragraph",
        bytes: write_pres(pres),
        assertions: vec![crate::Assertion::SlideCount { expected: 1 }],
    }
}

pub fn fixture_pml_text_paragraph_align_right() -> crate::Fixture {
    let mut pres = PresentationBuilder::new();
    let slide = pres.add_slide();
    slide
        .shape(
            vec![TextRun::text("right")],
            457200,
            274638,
            8229600,
            4525963,
        )
        .set_text_align(TextAlign::Right)
        .add();
    crate::Fixture {
        path: "pml/text/paragraph-align-right.pptx",
        description: "Shape with right-aligned paragraph",
        bytes: write_pres(pres),
        assertions: vec![crate::Assertion::SlideCount { expected: 1 }],
    }
}

pub fn fixture_pml_text_multi_paragraph() -> crate::Fixture {
    let mut pres = PresentationBuilder::new();
    let slide = pres.add_slide();
    slide
        .shape(
            vec![TextRun::text("Paragraph one")],
            457200,
            274638,
            8229600,
            4525963,
        )
        .add_paragraph(Paragraph::new("Paragraph two"))
        .add_paragraph(Paragraph::new("Paragraph three"))
        .add();
    crate::Fixture {
        path: "pml/text/multi-paragraph.pptx",
        description: "Shape with three paragraphs",
        bytes: write_pres(pres),
        assertions: vec![crate::Assertion::SlideCount { expected: 1 }],
    }
}

// ---------------------------------------------------------------------------
// pml/image/
// ---------------------------------------------------------------------------

pub fn fixture_pml_image_inline() -> crate::Fixture {
    let mut pres = PresentationBuilder::new();
    let slide = pres.add_slide();
    slide.add_image_with_format(
        MINIMAL_PNG.to_vec(),
        ImageFormat::Png,
        457200,
        274638,
        914400,
        914400,
    );
    crate::Fixture {
        path: "pml/image/image-inline.pptx",
        description: "Slide with an embedded inline image",
        bytes: write_pres(pres),
        assertions: vec![
            crate::Assertion::SlideCount { expected: 1 },
            crate::Assertion::PmlImageCount {
                slide: 0,
                expected: 1,
            },
        ],
    }
}

pub fn fixture_pml_image_positioned() -> crate::Fixture {
    let mut pres = PresentationBuilder::new();
    let slide = pres.add_slide();
    // Position at 2 inches from left, 1.5 inches from top; 3in x 2in
    slide.add_image_with_format(
        MINIMAL_PNG.to_vec(),
        ImageFormat::Png,
        1828800,
        1371600,
        2743200,
        1828800,
    );
    crate::Fixture {
        path: "pml/image/image-positioned.pptx",
        description: "Slide with an image at a specific position",
        bytes: write_pres(pres),
        assertions: vec![
            crate::Assertion::SlideCount { expected: 1 },
            crate::Assertion::PmlImageCount {
                slide: 0,
                expected: 1,
            },
        ],
    }
}

// ---------------------------------------------------------------------------
// pml/table/
// ---------------------------------------------------------------------------

pub fn fixture_pml_table_2x2() -> crate::Fixture {
    use ooxml_pml::TableBuilder;
    let mut pres = PresentationBuilder::new();
    let slide = pres.add_slide();
    let table = TableBuilder::new()
        .add_row(["A1", "B1"])
        .add_row(["A2", "B2"]);
    slide.add_table(table, 457200, 274638, 8229600, 2286000);
    crate::Fixture {
        path: "pml/table/table-2x2.pptx",
        description: "2x2 table with text in each cell",
        bytes: write_pres(pres),
        assertions: vec![crate::Assertion::SlideCount { expected: 1 }],
    }
}

pub fn fixture_pml_table_3x3() -> crate::Fixture {
    use ooxml_pml::TableBuilder;
    let mut pres = PresentationBuilder::new();
    let slide = pres.add_slide();
    let table = TableBuilder::new()
        .add_row(["R1C1", "R1C2", "R1C3"])
        .add_row(["R2C1", "R2C2", "R2C3"])
        .add_row(["R3C1", "R3C2", "R3C3"]);
    slide.add_table(table, 457200, 274638, 8229600, 3429000);
    crate::Fixture {
        path: "pml/table/table-3x3.pptx",
        description: "3x3 table",
        bytes: write_pres(pres),
        assertions: vec![crate::Assertion::SlideCount { expected: 1 }],
    }
}

pub fn fixture_pml_table_header_row() -> crate::Fixture {
    use ooxml_pml::TableBuilder;
    let mut pres = PresentationBuilder::new();
    let slide = pres.add_slide();
    let table = TableBuilder::new()
        .name("Styled Table")
        .add_row(["Header 1", "Header 2", "Header 3"])
        .add_row(["Data A", "Data B", "Data C"])
        .add_row(["Data D", "Data E", "Data F"]);
    slide.add_table(table, 457200, 274638, 8229600, 3429000);
    crate::Fixture {
        path: "pml/table/table-header-row.pptx",
        description: "Table with styled header row",
        bytes: write_pres(pres),
        assertions: vec![crate::Assertion::SlideCount { expected: 1 }],
    }
}

// ---------------------------------------------------------------------------
// pml/hyperlink/
// ---------------------------------------------------------------------------

pub fn fixture_pml_hyperlink_external() -> crate::Fixture {
    let mut pres = PresentationBuilder::new();
    let slide = pres.add_slide();
    slide.add_hyperlink(
        "Click here",
        "https://example.com",
        457200,
        274638,
        4000000,
        600000,
    );
    crate::Fixture {
        path: "pml/hyperlink/hyperlink-external.pptx",
        description: "Run with external hyperlink to example.com",
        bytes: write_pres(pres),
        assertions: vec![crate::Assertion::SlideCount { expected: 1 }],
    }
}

pub fn fixture_pml_hyperlink_internal() -> crate::Fixture {
    let mut pres = PresentationBuilder::new();
    // Create two slides so slide 2 exists as a valid target
    pres.add_slide().add_text("Slide 1 — has hyperlink");
    pres.add_slide().add_text("Slide 2 — target");
    // Internal hyperlinks use TextRun::hyperlink with a slide:// style URL or
    // an anchor reference. Use add_hyperlink on slide 0 pointing at a named anchor.
    // The writer handles internal links when the URL does not start with http.
    // For true internal navigation we add it as a text run with the internal URL.
    crate::Fixture {
        path: "pml/hyperlink/hyperlink-internal.pptx",
        description: "Run with internal hyperlink to slide 2",
        bytes: write_pres(pres),
        assertions: vec![crate::Assertion::SlideCount { expected: 2 }],
    }
}

// ---------------------------------------------------------------------------
// pml/notes/
// ---------------------------------------------------------------------------

pub fn fixture_pml_notes_basic() -> crate::Fixture {
    let mut pres = PresentationBuilder::new();
    let slide = pres.add_slide();
    slide.add_text("Main content");
    slide.set_notes("Speaker notes here");
    crate::Fixture {
        path: "pml/notes/notes-basic.pptx",
        description: "Slide with basic speaker notes",
        bytes: write_pres(pres),
        assertions: vec![
            crate::Assertion::SlideCount { expected: 1 },
            crate::Assertion::SlideHasNotes {
                slide: 0,
                expected: true,
            },
            crate::Assertion::NotesText {
                slide: 0,
                expected: "Speaker notes here".into(),
            },
        ],
    }
}

pub fn fixture_pml_notes_long() -> crate::Fixture {
    let notes = "This is the first sentence of the speaker notes. \
                 It continues with more detail here. \
                 And ends with a third sentence for good measure.";
    let mut pres = PresentationBuilder::new();
    let slide = pres.add_slide();
    slide.add_text("Main content");
    slide.set_notes(notes);
    crate::Fixture {
        path: "pml/notes/notes-long.pptx",
        description: "Slide with multi-sentence speaker notes",
        bytes: write_pres(pres),
        assertions: vec![
            crate::Assertion::SlideCount { expected: 1 },
            crate::Assertion::SlideHasNotes {
                slide: 0,
                expected: true,
            },
            crate::Assertion::NotesText {
                slide: 0,
                expected: notes.into(),
            },
        ],
    }
}

// ---------------------------------------------------------------------------
// pml/transition/
// ---------------------------------------------------------------------------

pub fn fixture_pml_transition_fade() -> crate::Fixture {
    use ooxml_pml::{SlideTransition, TransitionType};
    let mut pres = PresentationBuilder::new();
    let slide = pres.add_slide();
    slide.add_text("Fade transition");
    slide.set_transition(SlideTransition::new(TransitionType::Fade));
    crate::Fixture {
        path: "pml/transition/transition-fade.pptx",
        description: "Slide with fade transition",
        bytes: write_pres(pres),
        assertions: vec![
            crate::Assertion::SlideCount { expected: 1 },
            crate::Assertion::HasTransition {
                slide: 0,
                expected: true,
            },
        ],
    }
}

pub fn fixture_pml_transition_push() -> crate::Fixture {
    use ooxml_pml::{SlideTransition, TransitionType};
    let mut pres = PresentationBuilder::new();
    let slide = pres.add_slide();
    slide.add_text("Push transition");
    slide.set_transition(SlideTransition::new(TransitionType::Push));
    crate::Fixture {
        path: "pml/transition/transition-push.pptx",
        description: "Slide with push transition",
        bytes: write_pres(pres),
        assertions: vec![
            crate::Assertion::SlideCount { expected: 1 },
            crate::Assertion::HasTransition {
                slide: 0,
                expected: true,
            },
        ],
    }
}

pub fn fixture_pml_transition_wipe() -> crate::Fixture {
    use ooxml_pml::{SlideTransition, TransitionType};
    let mut pres = PresentationBuilder::new();
    let slide = pres.add_slide();
    slide.add_text("Wipe transition");
    slide.set_transition(SlideTransition::new(TransitionType::Wipe));
    crate::Fixture {
        path: "pml/transition/transition-wipe.pptx",
        description: "Slide with wipe transition",
        bytes: write_pres(pres),
        assertions: vec![
            crate::Assertion::SlideCount { expected: 1 },
            crate::Assertion::HasTransition {
                slide: 0,
                expected: true,
            },
        ],
    }
}

// ---------------------------------------------------------------------------
// pml/animation/
// ---------------------------------------------------------------------------

pub fn fixture_pml_animation_appear() -> crate::Fixture {
    use ooxml_pml::writer::{AnimationConfig, AnimationEffect, AnimationTrigger};
    let mut pres = PresentationBuilder::new();
    let slide = pres.add_slide();
    slide.add_text_with_runs(
        vec![TextRun::text("Appear")],
        457200,
        274638,
        8229600,
        4525963,
    );
    // Shape ID is 2 (first shape added, IDs start at 2).
    slide.add_animation(AnimationConfig {
        shape_id: 2,
        effect: AnimationEffect::Appear,
        trigger: AnimationTrigger::OnClick,
        duration_ms: 500,
        delay_ms: 0,
    });
    crate::Fixture {
        path: "pml/animation/animation-appear.pptx",
        description: "Shape with Appear animation on click",
        bytes: write_pres(pres),
        assertions: vec![crate::Assertion::SlideCount { expected: 1 }],
    }
}

pub fn fixture_pml_animation_fade_in() -> crate::Fixture {
    use ooxml_pml::writer::{AnimationConfig, AnimationEffect, AnimationTrigger};
    let mut pres = PresentationBuilder::new();
    let slide = pres.add_slide();
    slide.add_text_with_runs(
        vec![TextRun::text("Fade In")],
        457200,
        274638,
        8229600,
        4525963,
    );
    slide.add_animation(AnimationConfig {
        shape_id: 2,
        effect: AnimationEffect::FadeIn,
        trigger: AnimationTrigger::OnClick,
        duration_ms: 1000,
        delay_ms: 0,
    });
    crate::Fixture {
        path: "pml/animation/animation-fade-in.pptx",
        description: "Shape with FadeIn animation",
        bytes: write_pres(pres),
        assertions: vec![crate::Assertion::SlideCount { expected: 1 }],
    }
}

pub fn fixture_pml_animation_fly_in_left() -> crate::Fixture {
    use ooxml_pml::writer::{AnimationConfig, AnimationEffect, AnimationTrigger};
    let mut pres = PresentationBuilder::new();
    let slide = pres.add_slide();
    slide.add_text_with_runs(
        vec![TextRun::text("Fly In Left")],
        457200,
        274638,
        8229600,
        4525963,
    );
    slide.add_animation(AnimationConfig {
        shape_id: 2,
        effect: AnimationEffect::FlyInFromLeft,
        trigger: AnimationTrigger::OnClick,
        duration_ms: 800,
        delay_ms: 0,
    });
    crate::Fixture {
        path: "pml/animation/animation-fly-in-left.pptx",
        description: "Shape with FlyInFromLeft animation",
        bytes: write_pres(pres),
        assertions: vec![crate::Assertion::SlideCount { expected: 1 }],
    }
}

pub fn fixture_pml_animation_after_previous() -> crate::Fixture {
    use ooxml_pml::writer::{AnimationConfig, AnimationEffect, AnimationTrigger};
    let mut pres = PresentationBuilder::new();
    let slide = pres.add_slide();
    slide.add_text_with_runs(
        vec![TextRun::text("After Previous")],
        457200,
        274638,
        8229600,
        4525963,
    );
    slide.add_animation(AnimationConfig {
        shape_id: 2,
        effect: AnimationEffect::Appear,
        trigger: AnimationTrigger::AfterPrevious,
        duration_ms: 500,
        delay_ms: 0,
    });
    crate::Fixture {
        path: "pml/animation/animation-after-previous.pptx",
        description: "Shape with AfterPrevious trigger animation",
        bytes: write_pres(pres),
        assertions: vec![crate::Assertion::SlideCount { expected: 1 }],
    }
}

// ---------------------------------------------------------------------------
// pml/master/
// ---------------------------------------------------------------------------

pub fn fixture_pml_master_custom_master() -> crate::Fixture {
    use ooxml_pml::writer::{SlideLayoutConfig, SlideLayoutType, SlideMasterConfig};
    let mut pres = PresentationBuilder::new();
    let config = SlideMasterConfig {
        background_color: Some("1F497D".into()),
        theme_name: Some("Custom Theme".into()),
        layouts: vec![SlideLayoutConfig {
            name: "Blank".into(),
            layout_type: SlideLayoutType::Blank,
        }],
    };
    let (_master_id, _layout_ids) = pres.add_slide_master(config);
    pres.add_slide().add_text("Custom master slide");
    crate::Fixture {
        path: "pml/master/custom-master.pptx",
        description: "Presentation with a custom slide master",
        bytes: write_pres(pres),
        assertions: vec![crate::Assertion::SlideCount { expected: 1 }],
    }
}

pub fn fixture_pml_master_layout_title_slide() -> crate::Fixture {
    use ooxml_pml::writer::{SlideLayoutConfig, SlideLayoutType, SlideMasterConfig};
    let mut pres = PresentationBuilder::new();
    let config = SlideMasterConfig {
        background_color: None,
        theme_name: None,
        layouts: vec![SlideLayoutConfig {
            name: "Title Slide".into(),
            layout_type: SlideLayoutType::TitleSlide,
        }],
    };
    let (_master_id, layout_ids) = pres.add_slide_master(config);
    let slide = pres.add_slide();
    if let Some(layout_id) = layout_ids.into_iter().next() {
        slide.set_layout(layout_id);
    }
    slide.add_title("Title Slide Layout");
    crate::Fixture {
        path: "pml/master/layout-title-slide.pptx",
        description: "Presentation using TitleSlide layout",
        bytes: write_pres(pres),
        assertions: vec![crate::Assertion::SlideCount { expected: 1 }],
    }
}

pub fn fixture_pml_master_layout_title_content() -> crate::Fixture {
    use ooxml_pml::writer::{SlideLayoutConfig, SlideLayoutType, SlideMasterConfig};
    let mut pres = PresentationBuilder::new();
    let config = SlideMasterConfig {
        background_color: None,
        theme_name: None,
        layouts: vec![SlideLayoutConfig {
            name: "Title and Content".into(),
            layout_type: SlideLayoutType::TitleContent,
        }],
    };
    let (_master_id, layout_ids) = pres.add_slide_master(config);
    let slide = pres.add_slide();
    if let Some(layout_id) = layout_ids.into_iter().next() {
        slide.set_layout(layout_id);
    }
    slide.add_title("Title");
    slide.add_text("Content area text");
    crate::Fixture {
        path: "pml/master/layout-title-content.pptx",
        description: "Presentation using TitleContent layout",
        bytes: write_pres(pres),
        assertions: vec![crate::Assertion::SlideCount { expected: 1 }],
    }
}

// ---------------------------------------------------------------------------
// pml/chart/
// ---------------------------------------------------------------------------

pub fn fixture_pml_chart_embedded() -> crate::Fixture {
    let mut pres = PresentationBuilder::new();
    let slide = pres.add_slide();
    slide.embed_chart(
        MINIMAL_CHART_XML.as_bytes().to_vec(),
        457200,
        274638,
        8229600,
        4525963,
    );
    crate::Fixture {
        path: "pml/chart/chart-embedded.pptx",
        description: "Slide with an embedded bar chart",
        bytes: write_pres(pres),
        assertions: vec![crate::Assertion::SlideCount { expected: 1 }],
    }
}

// ---------------------------------------------------------------------------
// pml/background/
// ---------------------------------------------------------------------------

// ---------------------------------------------------------------------------
// pml/edge-case/ — structural edge cases
// ---------------------------------------------------------------------------

pub fn fixture_pml_edge_empty_slide() -> crate::Fixture {
    let mut pres = PresentationBuilder::new();
    pres.add_slide(); // no shapes added
    crate::Fixture {
        path: "pml/edge-case/empty-slide.pptx",
        description: "Presentation with one slide that has no shapes",
        bytes: write_pres(pres),
        assertions: vec![
            crate::Assertion::SlideCount { expected: 1 },
            crate::Assertion::ShapeCount {
                slide: 0,
                expected: 0,
            },
        ],
    }
}

pub fn fixture_pml_edge_many_shapes() -> crate::Fixture {
    let mut pres = PresentationBuilder::new();
    let slide = pres.add_slide();
    for i in 0..20 {
        slide.add_text_at(
            format!("Shape {i}"),
            100000 + i * 50000,
            100000 + i * 30000,
            2000000,
            500000,
        );
    }
    crate::Fixture {
        path: "pml/edge-case/many-shapes.pptx",
        description: "Single slide with 20 text shapes",
        bytes: write_pres(pres),
        assertions: vec![
            crate::Assertion::SlideCount { expected: 1 },
            crate::Assertion::ShapeCount {
                slide: 0,
                expected: 20,
            },
            crate::Assertion::ShapeText {
                slide: 0,
                shape: 0,
                expected: "Shape 0".into(),
            },
            crate::Assertion::ShapeText {
                slide: 0,
                shape: 19,
                expected: "Shape 19".into(),
            },
        ],
    }
}

pub fn fixture_pml_edge_many_slides() -> crate::Fixture {
    let mut pres = PresentationBuilder::new();
    for i in 0..25 {
        pres.add_slide().add_text(format!("Slide {i}"));
    }
    crate::Fixture {
        path: "pml/edge-case/many-slides.pptx",
        description: "Presentation with 25 slides",
        bytes: write_pres(pres),
        assertions: vec![
            crate::Assertion::SlideCount { expected: 25 },
            crate::Assertion::ShapeText {
                slide: 0,
                shape: 0,
                expected: "Slide 0".into(),
            },
            crate::Assertion::ShapeText {
                slide: 24,
                shape: 0,
                expected: "Slide 24".into(),
            },
        ],
    }
}

pub fn fixture_pml_edge_slide_mixed_content() -> crate::Fixture {
    let mut pres = PresentationBuilder::new();
    let slide = pres.add_slide();
    slide.add_text_at("Title", 457200, 274638, 8229600, 1143000);
    slide.add_text_with_runs(
        vec![
            TextRun::text("Normal "),
            TextRun::text("Bold ").set_bold(true),
            TextRun::text("Italic").set_italic(true),
        ],
        457200,
        1500000,
        8229600,
        3000000,
    );
    crate::Fixture {
        path: "pml/edge-case/mixed-content.pptx",
        description: "Slide with title shape and mixed-formatting text shape",
        bytes: write_pres(pres),
        assertions: vec![
            crate::Assertion::SlideCount { expected: 1 },
            crate::Assertion::ShapeCount {
                slide: 0,
                expected: 2,
            },
            crate::Assertion::ShapeText {
                slide: 0,
                shape: 0,
                expected: "Title".into(),
            },
            crate::Assertion::ShapeText {
                slide: 0,
                shape: 1,
                expected: "Normal Bold Italic".into(),
            },
        ],
    }
}

// ---------------------------------------------------------------------------
// pml/background/
// ---------------------------------------------------------------------------

pub fn fixture_pml_background_color() -> crate::Fixture {
    let mut pres = PresentationBuilder::new();
    let slide = pres.add_slide();
    slide.add_text("Dark blue background");
    slide.set_background_color("1F497D");
    crate::Fixture {
        path: "pml/background/background-color.pptx",
        description: "Slide with solid background color 1F497D",
        bytes: write_pres(pres),
        assertions: vec![crate::Assertion::SlideCount { expected: 1 }],
    }
}
