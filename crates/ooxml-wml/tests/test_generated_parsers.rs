//! Tests for WML generated FromXml parsers.
//!
//! These tests verify that the generated event-based parsers
//! can correctly parse WML XML snippets.

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
    match run.run_content[0].as_ref() {
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
    match para.paragraph_content[0].as_ref() {
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
    match table.rows[0].as_ref() {
        RowContent::Tr(_) => {}
        other => panic!("expected Tr variant, got {:?}", other),
    }
}

#[test]
fn test_parse_run_with_break() {
    let xml = r#"<r><t>Before</t><br/><t>After</t></r>"#;
    let run: Run = parse_from_xml(xml).expect("should parse run with break");
    assert_eq!(run.run_content.len(), 3);
    match run.run_content[1].as_ref() {
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
