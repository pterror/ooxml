//! Roundtrip tests for WML generated ToXml serializers.
//!
//! These tests verify that parsing XML via FromXml and then serializing
//! via ToXml produces equivalent XML output.

use ooxml_wml::parsers::{FromXml, ParseError};
use ooxml_wml::serializers::{SerializeError, ToXml};
use ooxml_wml::types::*;
use quick_xml::Reader;
use quick_xml::Writer;
use quick_xml::events::Event;
use std::io::Cursor;

/// Parse an XML string using the FromXml trait.
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

/// Serialize a value to an XML string using the ToXml trait.
fn serialize_to_xml<T: ToXml>(value: &T, tag: &str) -> Result<String, SerializeError> {
    let mut writer = Writer::new(Vec::new());
    value.write_element(tag, &mut writer)?;
    Ok(String::from_utf8(writer.into_inner()).expect("valid utf-8"))
}

/// Parse → serialize → re-parse helper. Returns the re-parsed value.
fn roundtrip<T: FromXml + ToXml>(xml: &str, tag: &str) -> T {
    let parsed: T = parse_from_xml(xml).expect("initial parse should succeed");
    let serialized = serialize_to_xml(&parsed, tag).expect("serialization should succeed");
    parse_from_xml::<T>(&serialized).unwrap_or_else(|e| {
        panic!(
            "re-parse after roundtrip failed: {:?}\nserialized XML: {}",
            e, serialized
        )
    })
}

// ====================================================================
// Empty element tests
// ====================================================================

#[test]
fn test_serialize_empty_element() {
    let xml = r#"<b/>"#;
    let bold: CTOnOff = parse_from_xml(xml).expect("parse empty bold");
    let out = serialize_to_xml(&bold, "w:b").unwrap();
    // Empty element with no attributes should self-close
    assert_eq!(out, "<w:b/>");
}

#[test]
fn test_serialize_boolean_with_val() {
    let xml = r#"<b val="false"/>"#;
    let bold: CTOnOff = parse_from_xml(xml).expect("parse bold with val");
    let out = serialize_to_xml(&bold, "w:b").unwrap();
    assert!(
        out.contains("val="),
        "should contain val attribute: {}",
        out
    );
}

// ====================================================================
// Text content tests
// ====================================================================

#[test]
fn test_serialize_text_element() {
    let xml = r#"<t>Hello World</t>"#;
    let text: Text = parse_from_xml(xml).expect("parse text");
    let out = serialize_to_xml(&text, "w:t").unwrap();
    assert!(out.contains("Hello World"), "should contain text: {}", out);
    assert!(
        out.starts_with("<w:t>"),
        "should start with w:t tag: {}",
        out
    );
    assert!(
        out.ends_with("</w:t>"),
        "should end with closing tag: {}",
        out
    );
}

#[test]
fn test_roundtrip_text_element() {
    let xml = r#"<t>Hello World</t>"#;
    let rt: Text = roundtrip(xml, "w:t");
    assert_eq!(rt.text.as_deref(), Some("Hello World"));
}

// ====================================================================
// Attribute tests
// ====================================================================

#[test]
fn test_serialize_underline_with_val() {
    let xml = r#"<u val="single"/>"#;
    let u: CTUnderline = parse_from_xml(xml).expect("parse underline");
    let out = serialize_to_xml(&u, "w:u").unwrap();
    assert!(out.contains("single"), "should contain val=single: {}", out);
}

#[test]
fn test_roundtrip_underline() {
    let xml = r#"<u val="single"/>"#;
    let rt: CTUnderline = roundtrip(xml, "w:u");
    assert!(rt.value.is_some());
}

// ====================================================================
// Run tests
// ====================================================================

#[test]
fn test_serialize_run_with_text() {
    let xml = r#"<r><t>Hello</t></r>"#;
    let run: Run = parse_from_xml(xml).expect("parse run");
    let out = serialize_to_xml(&run, "w:r").unwrap();
    assert!(out.contains("Hello"), "should contain text: {}", out);
    assert!(out.contains("w:t"), "should contain w:t element: {}", out);
}

#[test]
fn test_roundtrip_run_with_text() {
    let xml = r#"<r><t>Hello</t></r>"#;
    let rt: Run = roundtrip(xml, "w:r");
    assert_eq!(rt.run_content.len(), 1);
    match &rt.run_content[0] {
        RunContent::T(t) => {
            assert_eq!(t.text.as_deref(), Some("Hello"));
        }
        other => panic!("expected T variant, got {:?}", other),
    }
}

#[test]
fn test_roundtrip_run_with_properties_and_text() {
    let xml = r#"<r><rPr><b/><i/></rPr><t>Bold italic</t></r>"#;
    let rt: Run = roundtrip(xml, "w:r");
    assert!(rt.r_pr.is_some());
    let rpr = rt.r_pr.as_ref().unwrap();
    assert!(rpr.bold.is_some());
    assert!(rpr.italic.is_some());
    assert_eq!(rt.run_content.len(), 1);
}

#[test]
fn test_roundtrip_run_with_break() {
    let xml = r#"<r><t>Before</t><br/><t>After</t></r>"#;
    let rt: Run = roundtrip(xml, "w:r");
    assert_eq!(rt.run_content.len(), 3);
    match &rt.run_content[1] {
        RunContent::Br(_) => {}
        other => panic!("expected Br variant, got {:?}", other),
    }
}

#[test]
fn test_roundtrip_run_with_rsid() {
    let xml = r#"<r rsidR="00A77427"><t>Text</t></r>"#;
    let rt: Run = roundtrip(xml, "w:r");
    assert!(rt.rsid_r.is_some());
    assert_eq!(rt.run_content.len(), 1);
}

// ====================================================================
// RunProperties tests
// ====================================================================

#[test]
fn test_serialize_empty_run_properties() {
    let xml = r#"<rPr/>"#;
    let rpr: RunProperties = parse_from_xml(xml).expect("parse empty rPr");
    let out = serialize_to_xml(&rpr, "w:rPr").unwrap();
    // Empty rPr should self-close
    assert_eq!(out, "<w:rPr/>");
}

#[test]
fn test_roundtrip_run_properties_formatting() {
    let xml = r#"<rPr><b/><i/><u val="single"/><strike/><sz val="24"/></rPr>"#;
    let rt: RunProperties = roundtrip(xml, "w:rPr");
    assert!(rt.bold.is_some());
    assert!(rt.italic.is_some());
    assert!(rt.underline.is_some());
    assert!(rt.strikethrough.is_some());
    assert!(rt.size.is_some());
}

// ====================================================================
// Paragraph tests
// ====================================================================

#[test]
fn test_serialize_empty_paragraph() {
    let xml = r#"<p/>"#;
    let para: Paragraph = parse_from_xml(xml).expect("parse empty p");
    let out = serialize_to_xml(&para, "w:p").unwrap();
    assert_eq!(out, "<w:p/>");
}

#[test]
fn test_roundtrip_paragraph_simple() {
    let xml = r#"<p><r><t>Hello</t></r></p>"#;
    let rt: Paragraph = roundtrip(xml, "w:p");
    assert!(rt.p_pr.is_none());
    assert_eq!(rt.paragraph_content.len(), 1);
    match &rt.paragraph_content[0] {
        ParagraphContent::R(_) => {}
        other => panic!("expected R variant, got {:?}", other),
    }
}

#[test]
fn test_roundtrip_paragraph_with_properties() {
    let xml = r#"<p><pPr><jc val="center"/></pPr><r><t>Centered</t></r></p>"#;
    let rt: Paragraph = roundtrip(xml, "w:p");
    assert!(rt.p_pr.is_some());
    assert_eq!(rt.paragraph_content.len(), 1);
}

#[test]
fn test_roundtrip_paragraph_multiple_runs() {
    let xml = r#"<p><r><t>Hello </t></r><r><rPr><b/></rPr><t>World</t></r></p>"#;
    let rt: Paragraph = roundtrip(xml, "w:p");
    assert_eq!(rt.paragraph_content.len(), 2);
}

// ====================================================================
// Body tests
// ====================================================================

#[test]
fn test_roundtrip_body_with_paragraphs() {
    let xml = r#"<body><p><r><t>First</t></r></p><p><r><t>Second</t></r></p></body>"#;
    let rt: Body = roundtrip(xml, "w:body");
    assert_eq!(rt.block_content.len(), 2);
    assert!(rt.sect_pr.is_none());
}

#[test]
fn test_roundtrip_body_with_section_properties() {
    let xml =
        r#"<body><p><r><t>Hello</t></r></p><sectPr><pgSz w="12240" h="15840"/></sectPr></body>"#;
    let rt: Body = roundtrip(xml, "w:body");
    assert_eq!(rt.block_content.len(), 1);
    assert!(rt.sect_pr.is_some());
}

// ====================================================================
// Document tests
// ====================================================================

#[test]
fn test_roundtrip_document() {
    let xml = r#"<document conformance="transitional"><body><p><r><t>Hello</t></r></p></body></document>"#;
    let rt: Document = roundtrip(xml, "w:document");
    assert!(rt.body.is_some());
    let body = rt.body.as_ref().unwrap();
    assert_eq!(body.block_content.len(), 1);
}

// ====================================================================
// Table tests
// ====================================================================

#[test]
fn test_roundtrip_table() {
    let xml = r#"<tbl><tblPr/><tblGrid><gridCol/></tblGrid><tr><tc><p><r><t>Cell</t></r></p></tc></tr></tbl>"#;
    let rt: Table = roundtrip(xml, "w:tbl");
    assert_eq!(rt.rows.len(), 1);
    match &rt.rows[0] {
        RowContent::Tr(_) => {}
        other => panic!("expected Tr variant, got {:?}", other),
    }
}

// ====================================================================
// Serialize output format verification
// ====================================================================

#[test]
fn test_serialize_preserves_namespace_prefix() {
    let xml = r#"<p><r><t>Hello</t></r></p>"#;
    let para: Paragraph = parse_from_xml(xml).expect("parse paragraph");
    let out = serialize_to_xml(&para, "w:p").unwrap();
    // Verify namespace prefixes are present
    assert!(out.starts_with("<w:p>"), "should start with <w:p>: {}", out);
    assert!(out.contains("<w:r>"), "should contain <w:r>: {}", out);
    assert!(out.contains("<w:t>"), "should contain <w:t>: {}", out);
    assert!(out.contains("</w:t>"), "should contain </w:t>: {}", out);
    assert!(out.contains("</w:r>"), "should contain </w:r>: {}", out);
    assert!(out.ends_with("</w:p>"), "should end with </w:p>: {}", out);
}

#[test]
fn test_serialize_self_closing_empty_elements() {
    let xml = r#"<rPr><b/><i/></rPr>"#;
    let rpr: RunProperties = parse_from_xml(xml).expect("parse rPr");
    let out = serialize_to_xml(&rpr, "w:rPr").unwrap();
    // b and i should be self-closing since they're empty
    assert!(
        out.contains("<w:b/>"),
        "bold should be self-closing: {}",
        out
    );
    assert!(
        out.contains("<w:i/>"),
        "italic should be self-closing: {}",
        out
    );
}

#[test]
fn test_serialize_attributes_on_start_tag() {
    let xml =
        r#"<body><p><r><t>Hello</t></r></p><sectPr><pgSz w="12240" h="15840"/></sectPr></body>"#;
    let body: Body = parse_from_xml(xml).expect("parse body");
    let out = serialize_to_xml(&body, "w:body").unwrap();
    // pgSz should have w and h attributes
    assert!(out.contains("w:pgSz"), "should contain pgSz: {}", out);
    assert!(out.contains("12240"), "should contain page width: {}", out);
    assert!(out.contains("15840"), "should contain page height: {}", out);
}
