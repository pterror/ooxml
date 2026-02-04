//! Tests for WML generated FromXml parsers.
//!
//! These tests verify that the generated event-based parsers
//! can correctly parse WML XML snippets.

#[cfg(feature = "extra-children")]
use ooxml_wml::RawXmlNode;
use ooxml_wml::parsers::{FromXml, ParseError};
use ooxml_wml::types::*;
use quick_xml::Reader;
use quick_xml::events::Event;
use std::io::Cursor;

/// Helper to parse an XML string using the FromXml trait.
fn parse_from_xml<T: FromXml>(xml: &str) -> Result<T, ParseError> {
    let mut reader = Reader::from_reader(Cursor::new(xml.as_bytes()));
    let mut buf = Vec::new();
    loop {
        match reader.read_event_into(&mut buf)? {
            Event::Start(e) => return T::from_xml(&mut reader, &e, false),
            Event::Empty(e) => return T::from_xml(&mut reader, &e, true),
            Event::Eof => break,
            _ => {}
        }
        buf.clear();
    }
    Err(ParseError::UnexpectedElement(
        "no element found".to_string(),
    ))
}

#[test]
fn test_parse_run_with_text() {
    let xml = r#"<r><t>Hello World</t></r>"#;
    let run: Run = parse_from_xml(xml).expect("should parse run");
    assert!(run.r_pr.is_none());
    assert_eq!(run.run_content.len(), 1);
    // Verify it's a text element
    match &run.run_content[0] {
        RunContent::T(_) => {}
        other => panic!("expected T variant, got {:?}", other),
    }
}

#[test]
fn test_parse_run_with_properties_and_text() {
    let xml = r#"<r><rPr><b/><i/></rPr><t>Bold italic</t></r>"#;
    let run: Run = parse_from_xml(xml).expect("should parse run");
    assert!(run.r_pr.is_some());
    let rpr = run.r_pr.as_ref().unwrap();
    assert!(rpr.bold.is_some());
    assert!(rpr.italic.is_some());
    assert!(rpr.caps.is_none());
    assert_eq!(run.run_content.len(), 1);
}

#[test]
fn test_parse_run_properties_formatting() {
    let xml = r#"<rPr><b/><i/><u val="single"/><strike/><sz val="24"/></rPr>"#;
    let rpr: RunProperties = parse_from_xml(xml).expect("should parse rPr");
    assert!(rpr.bold.is_some());
    assert!(rpr.italic.is_some());
    assert!(rpr.underline.is_some());
    assert!(rpr.strikethrough.is_some());
    assert!(rpr.size.is_some());
    assert!(rpr.caps.is_none());
    assert!(rpr.small_caps.is_none());
}

#[test]
fn test_parse_run_properties_empty() {
    let xml = r#"<rPr/>"#;
    let rpr: RunProperties = parse_from_xml(xml).expect("should parse empty rPr");
    assert!(rpr.bold.is_none());
    assert!(rpr.italic.is_none());
}

#[test]
fn test_parse_paragraph_simple() {
    let xml = r#"<p><r><t>Hello</t></r></p>"#;
    let para: Paragraph = parse_from_xml(xml).expect("should parse paragraph");
    assert!(para.p_pr.is_none());
    assert_eq!(para.paragraph_content.len(), 1);
    // First content item should be a Run
    match &para.paragraph_content[0] {
        ParagraphContent::R(_) => {}
        other => panic!("expected R variant, got {:?}", other),
    }
}

#[test]
fn test_parse_paragraph_with_properties() {
    let xml = r#"<p><pPr><jc val="center"/></pPr><r><t>Centered</t></r></p>"#;
    let para: Paragraph = parse_from_xml(xml).expect("should parse paragraph");
    assert!(para.p_pr.is_some());
    assert_eq!(para.paragraph_content.len(), 1);
}

#[test]
fn test_parse_paragraph_multiple_runs() {
    let xml = r#"<p><r><t>Hello </t></r><r><rPr><b/></rPr><t>World</t></r></p>"#;
    let para: Paragraph = parse_from_xml(xml).expect("should parse paragraph");
    assert_eq!(para.paragraph_content.len(), 2);
}

#[test]
fn test_parse_body_with_paragraphs() {
    let xml = r#"<body><p><r><t>First</t></r></p><p><r><t>Second</t></r></p></body>"#;
    let body: Body = parse_from_xml(xml).expect("should parse body");
    assert_eq!(body.block_content.len(), 2);
    assert!(body.sect_pr.is_none());
}

#[test]
fn test_parse_body_with_section_properties() {
    let xml =
        r#"<body><p><r><t>Hello</t></r></p><sectPr><pgSz w="12240" h="15840"/></sectPr></body>"#;
    let body: Body = parse_from_xml(xml).expect("should parse body");
    assert_eq!(body.block_content.len(), 1);
    assert!(body.sect_pr.is_some());
}

#[test]
fn test_parse_document() {
    let xml = r#"<document conformance="transitional"><body><p><r><t>Hello</t></r></p></body></document>"#;
    let doc: Document = parse_from_xml(xml).expect("should parse document");
    assert!(doc.body.is_some());
    let body = doc.body.as_ref().unwrap();
    assert_eq!(body.block_content.len(), 1);
}

#[test]
fn test_parse_table_basic() {
    let xml = r#"<tbl><tblPr/><tblGrid><gridCol/></tblGrid><tr><tc><p><r><t>Cell</t></r></p></tc></tr></tbl>"#;
    let table: Table = parse_from_xml(xml).expect("should parse table");
    assert_eq!(table.rows.len(), 1);
    // First content should be a table row
    match &table.rows[0] {
        RowContent::Tr(_) => {}
        other => panic!("expected Tr variant, got {:?}", other),
    }
}

#[test]
fn test_parse_run_with_break() {
    let xml = r#"<r><t>Before</t><br/><t>After</t></r>"#;
    let run: Run = parse_from_xml(xml).expect("should parse run with break");
    assert_eq!(run.run_content.len(), 3);
    match &run.run_content[1] {
        RunContent::Br(_) => {}
        other => panic!("expected Br variant, got {:?}", other),
    }
}

#[test]
fn test_parse_run_with_attributes() {
    let xml = r#"<r rsidR="00A77427"><t>Text</t></r>"#;
    let run: Run = parse_from_xml(xml).expect("should parse run with rsid");
    assert!(run.rsid_r.is_some());
}

#[test]
fn test_parse_empty_paragraph() {
    let xml = r#"<p/>"#;
    let para: Paragraph = parse_from_xml(xml).expect("should parse empty paragraph");
    assert!(para.p_pr.is_none());
    assert!(para.paragraph_content.is_empty());
}

// =========================================================================
// ParagraphProperties
// =========================================================================

#[test]
fn test_parse_paragraph_properties_style_and_alignment() {
    let xml = r#"<pPr><pStyle val="Heading1"/><jc val="center"/></pPr>"#;
    let ppr: ParagraphProperties = parse_from_xml(xml).expect("should parse pPr");
    assert!(ppr.paragraph_style.is_some());
    assert_eq!(ppr.paragraph_style.as_ref().unwrap().value, "Heading1");
    assert!(ppr.justification.is_some());
}

#[test]
fn test_parse_paragraph_properties_numbering() {
    let xml = r#"<pPr><numPr><ilvl val="0"/><numId val="1"/></numPr></pPr>"#;
    let ppr: ParagraphProperties = parse_from_xml(xml).expect("should parse pPr with numbering");
    assert!(ppr.num_pr.is_some());
}

#[test]
fn test_parse_paragraph_properties_indentation_and_spacing() {
    let xml = r#"<pPr><ind left="720" hanging="360"/><spacing before="240" after="120"/></pPr>"#;
    let ppr: ParagraphProperties = parse_from_xml(xml).expect("should parse pPr with ind+spacing");
    assert!(ppr.indentation.is_some());
    assert!(ppr.spacing.is_some());
}

#[test]
fn test_parse_paragraph_properties_section() {
    let xml = r#"<pPr><sectPr><pgSz w="12240" h="15840"/></sectPr></pPr>"#;
    let ppr: ParagraphProperties = parse_from_xml(xml).expect("should parse pPr with sectPr");
    assert!(ppr.sect_pr.is_some());
}

// =========================================================================
// SectionProperties
// =========================================================================

#[test]
fn test_parse_section_properties_page_size() {
    let xml = r#"<sectPr><pgSz w="12240" h="15840"/></sectPr>"#;
    let sp: SectionProperties = parse_from_xml(xml).expect("should parse sectPr");
    assert!(sp.pg_sz.is_some());
    let pg_sz = sp.pg_sz.as_ref().unwrap();
    assert_eq!(pg_sz.width.as_deref(), Some("12240"));
    assert_eq!(pg_sz.height.as_deref(), Some("15840"));
}

#[test]
fn test_parse_section_properties_margins() {
    let xml = r#"<sectPr><pgMar top="1440" right="1440" bottom="1440" left="1440" header="720" footer="720" gutter="0"/></sectPr>"#;
    let sp: SectionProperties = parse_from_xml(xml).expect("should parse sectPr with margins");
    assert!(sp.pg_mar.is_some());
    let mar = sp.pg_mar.as_ref().unwrap();
    assert_eq!(mar.top, "1440");
    assert_eq!(mar.left, "1440");
}

#[test]
fn test_parse_section_properties_columns() {
    let xml = r#"<sectPr><cols space="720" num="2"/></sectPr>"#;
    let sp: SectionProperties = parse_from_xml(xml).expect("should parse sectPr with cols");
    assert!(sp.cols.is_some());
}

#[test]
fn test_parse_section_properties_title_page() {
    let xml = r#"<sectPr><titlePg/></sectPr>"#;
    let sp: SectionProperties = parse_from_xml(xml).expect("should parse sectPr with titlePg");
    assert!(sp.title_pg.is_some());
}

// =========================================================================
// Table with properties
// =========================================================================

#[test]
fn test_parse_table_with_full_properties() {
    let xml = r#"<tbl>
        <tblPr>
            <tblStyle val="TableGrid"/>
            <tblW w="0" type="auto"/>
            <tblLook val="04A0"/>
        </tblPr>
        <tblGrid><gridCol w="4680"/><gridCol w="4680"/></tblGrid>
        <tr>
            <trPr><cantSplit/></trPr>
            <tc>
                <tcPr><tcW w="4680" type="dxa"/><gridSpan val="1"/></tcPr>
                <p><r><t>A1</t></r></p>
            </tc>
            <tc>
                <tcPr><tcW w="4680" type="dxa"/></tcPr>
                <p><r><t>B1</t></r></p>
            </tc>
        </tr>
    </tbl>"#;
    let table: Table = parse_from_xml(xml).expect("should parse table with full properties");
    // Table properties
    assert!(table.table_properties.tbl_style.is_some());
    assert_eq!(
        table.table_properties.tbl_style.as_ref().unwrap().value,
        "TableGrid"
    );
    assert!(table.table_properties.tbl_w.is_some());
    assert!(table.table_properties.tbl_look.is_some());

    // Grid columns
    assert_eq!(table.tbl_grid.grid_col.len(), 2);

    // Row with properties
    assert_eq!(table.rows.len(), 1);
    if let RowContent::Tr(row) = &table.rows[0] {
        assert!(row.row_properties.is_some());
        // Two cells
        assert_eq!(row.cells.len(), 2);
        if let CellContent::Tc(cell) = &row.cells[0] {
            assert!(cell.cell_properties.is_some());
            let tc_pr = cell.cell_properties.as_ref().unwrap();
            assert!(tc_pr.tc_w.is_some());
            assert!(tc_pr.grid_span.is_some());
        } else {
            panic!("expected Tc variant");
        }
    } else {
        panic!("expected Tr variant");
    }
}

#[test]
fn test_parse_table_merged_cells() {
    let xml = r#"<tbl>
        <tblPr/><tblGrid><gridCol/><gridCol/></tblGrid>
        <tr>
            <tc><tcPr><vMerge val="restart"/></tcPr><p/></tc>
            <tc><p><r><t>B1</t></r></p></tc>
        </tr>
        <tr>
            <tc><tcPr><vMerge/></tcPr><p/></tc>
            <tc><p><r><t>B2</t></r></p></tc>
        </tr>
    </tbl>"#;
    let table: Table = parse_from_xml(xml).expect("should parse merged table");
    assert_eq!(table.rows.len(), 2);

    // First row, first cell has vMerge restart
    if let RowContent::Tr(row) = &table.rows[0] {
        if let CellContent::Tc(cell) = &row.cells[0] {
            let tc_pr = cell.cell_properties.as_ref().unwrap();
            assert!(tc_pr.vertical_merge.is_some());
        } else {
            panic!("expected Tc");
        }
    } else {
        panic!("expected Tr");
    }
}

// =========================================================================
// Style
// =========================================================================

#[test]
fn test_parse_style_paragraph() {
    let xml = r#"<style type="paragraph" styleId="Heading1" default="0">
        <name val="heading 1"/>
        <basedOn val="Normal"/>
        <next val="Normal"/>
        <qFormat/>
        <pPr><keepNext/><outlineLvl val="0"/></pPr>
        <rPr><b/><sz val="28"/></rPr>
    </style>"#;
    let style: Style = parse_from_xml(xml).expect("should parse style");
    assert_eq!(style.style_id.as_deref(), Some("Heading1"));
    assert!(style.name.is_some());
    assert_eq!(style.name.as_ref().unwrap().value, "heading 1");
    assert!(style.based_on.is_some());
    assert_eq!(style.based_on.as_ref().unwrap().value, "Normal");
    assert!(style.next.is_some());
    assert!(style.q_format.is_some());
    assert!(style.p_pr.is_some());
    assert!(style.r_pr.is_some());
}

#[test]
fn test_parse_style_character() {
    let xml = r#"<style type="character" styleId="Strong">
        <name val="Strong"/>
        <rPr><b/></rPr>
    </style>"#;
    let style: Style = parse_from_xml(xml).expect("should parse character style");
    assert_eq!(style.style_id.as_deref(), Some("Strong"));
    assert!(style.r_pr.is_some());
    assert!(style.p_pr.is_none());
}

// =========================================================================
// Hyperlink
// =========================================================================

#[test]
fn test_parse_paragraph_with_hyperlink() {
    let xml = r#"<p>
        <hyperlink anchor="bookmark1" tooltip="Click here">
            <r><rPr><u val="single"/></rPr><t>Link text</t></r>
        </hyperlink>
    </p>"#;
    let para: Paragraph = parse_from_xml(xml).expect("should parse paragraph with hyperlink");
    assert_eq!(para.paragraph_content.len(), 1);
    match &para.paragraph_content[0] {
        ParagraphContent::Hyperlink(h) => {
            assert_eq!(h.anchor.as_deref(), Some("bookmark1"));
            assert_eq!(h.tooltip.as_deref(), Some("Click here"));
            assert_eq!(h.paragraph_content.len(), 1);
        }
        other => panic!("expected Hyperlink variant, got {:?}", other),
    }
}

// =========================================================================
// Namespace-prefixed XML (real-world documents use w: prefix)
// =========================================================================

#[test]
fn test_parse_namespace_prefixed_document() {
    let xml = r#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
        <w:body>
            <w:p>
                <w:pPr><w:pStyle w:val="Normal"/></w:pPr>
                <w:r><w:rPr><w:b/></w:rPr><w:t>Hello</w:t></w:r>
            </w:p>
        </w:body>
    </w:document>"#;
    let doc: Document = parse_from_xml(xml).expect("should parse namespaced document");
    let body = doc.body.as_ref().expect("should have body");
    assert_eq!(body.block_content.len(), 1);

    // Verify we parsed through the namespace prefixes
    if let BlockContent::P(para) = &body.block_content[0] {
        assert!(para.p_pr.is_some());
        assert_eq!(para.paragraph_content.len(), 1);
        if let ParagraphContent::R(run) = &para.paragraph_content[0] {
            assert!(run.r_pr.is_some());
            let rpr = run.r_pr.as_ref().unwrap();
            assert!(rpr.bold.is_some());
        } else {
            panic!("expected Run");
        }
    } else {
        panic!("expected Paragraph");
    }
}

#[test]
fn test_parse_namespace_prefixed_table() {
    let xml = r#"<w:tbl xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
        <w:tblPr><w:tblStyle w:val="TableGrid"/></w:tblPr>
        <w:tblGrid><w:gridCol w:w="4680"/></w:tblGrid>
        <w:tr><w:tc><w:p><w:r><w:t>Cell</w:t></w:r></w:p></w:tc></w:tr>
    </w:tbl>"#;
    let table: Table = parse_from_xml(xml).expect("should parse namespaced table");
    assert_eq!(
        table.table_properties.tbl_style.as_ref().unwrap().value,
        "TableGrid"
    );
    assert_eq!(table.rows.len(), 1);
}

// =========================================================================
// Unknown element position tracking (ADR-004)
// =========================================================================

#[cfg(feature = "extra-children")]
#[test]
fn test_unknown_children_preserve_position() {
    // Parse a paragraph with a known child (run), an unknown child, then another run.
    // The unknown child should be captured in extra_children with the correct position.
    let xml = r#"<p>
        <r><t>First</t></r>
        <customElement attr="val"/>
        <r><t>Second</t></r>
    </p>"#;
    let para: Paragraph = parse_from_xml(xml).expect("should parse paragraph with unknown");
    assert_eq!(para.paragraph_content.len(), 2, "should have 2 known runs");
    assert_eq!(
        para.extra_children.len(),
        1,
        "should capture 1 unknown child"
    );

    // The unknown child was at position 1 (between the two runs at positions 0 and 2)
    let extra = &para.extra_children[0];
    assert_eq!(extra.position, 1, "unknown should be at position 1");
    if let RawXmlNode::Element(elem) = &extra.node {
        assert!(
            elem.name.ends_with("customElement"),
            "should capture customElement"
        );
    } else {
        panic!("expected Element node");
    }
}

#[cfg(feature = "extra-children")]
#[test]
fn test_unknown_children_multiple_positions() {
    // Unknown children at the start, middle, and end
    let xml = r#"<body>
        <unknown1/>
        <p><r><t>Para</t></r></p>
        <unknown2/>
        <p><r><t>Para2</t></r></p>
        <unknown3/>
    </body>"#;
    let body: Body = parse_from_xml(xml).expect("should parse body with unknowns");
    assert_eq!(body.block_content.len(), 2, "should have 2 paragraphs");
    assert_eq!(body.extra_children.len(), 3, "should capture 3 unknowns");

    // Verify positions: unknown1 at 0, para at 1, unknown2 at 2, para2 at 3, unknown3 at 4
    assert_eq!(body.extra_children[0].position, 0);
    assert_eq!(body.extra_children[1].position, 2);
    assert_eq!(body.extra_children[2].position, 4);
}

#[cfg(feature = "extra-children")]
#[test]
fn test_unknown_children_empty_element() {
    // Verify that self-closing unknown elements are captured with correct positions
    let xml = r#"<rPr><b/><unknownProp/><i/></rPr>"#;
    let rpr: RunProperties = parse_from_xml(xml).expect("should parse rPr with unknown");
    assert!(rpr.bold.is_some());
    assert!(rpr.italic.is_some());
    assert_eq!(rpr.extra_children.len(), 1);
    assert_eq!(rpr.extra_children[0].position, 1);
}
