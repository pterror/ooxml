//! Tests for DML generated FromXml parsers.
//!
//! These tests verify that the generated event-based parsers can correctly
//! parse DML XML snippets — including bare tags, namespace-prefixed tags,
//! and edge cases like XML entity references.

use ooxml_dml::parsers::{FromXml, ParseError};
use ooxml_dml::types::*;
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

// ---------------------------------------------------------------------------
// TextCharacterProperties (a:rPr)
// ---------------------------------------------------------------------------

#[test]
fn test_parse_text_char_props_empty() {
    let xml = r#"<rPr/>"#;
    let rpr: TextCharacterProperties = parse_from_xml(xml).expect("should parse empty rPr");
    assert_eq!(rpr.b, None);
    assert_eq!(rpr.i, None);
    assert_eq!(rpr.sz, None);
    assert_eq!(rpr.lang, None);
}

#[test]
fn test_parse_text_char_props_bold() {
    let xml = r#"<rPr b="1"/>"#;
    let rpr: TextCharacterProperties = parse_from_xml(xml).expect("should parse bold rPr");
    assert_eq!(rpr.b, Some(true));
    assert_eq!(rpr.i, None);
}

#[test]
fn test_parse_text_char_props_bold_false() {
    let xml = r#"<rPr b="0"/>"#;
    let rpr: TextCharacterProperties = parse_from_xml(xml).expect("should parse b=0 rPr");
    assert_eq!(rpr.b, Some(false));
}

#[test]
fn test_parse_text_char_props_italic() {
    let xml = r#"<rPr i="1"/>"#;
    let rpr: TextCharacterProperties = parse_from_xml(xml).expect("should parse italic rPr");
    assert_eq!(rpr.i, Some(true));
    assert_eq!(rpr.b, None);
}

#[test]
fn test_parse_text_char_props_bold_italic() {
    let xml = r#"<rPr b="1" i="1"/>"#;
    let rpr: TextCharacterProperties = parse_from_xml(xml).expect("should parse bold+italic rPr");
    assert_eq!(rpr.b, Some(true));
    assert_eq!(rpr.i, Some(true));
}

#[test]
fn test_parse_text_char_props_font_size() {
    // sz is in hundredths of a point: 2400 = 24pt
    let xml = r#"<rPr sz="2400"/>"#;
    let rpr: TextCharacterProperties = parse_from_xml(xml).expect("should parse font-size rPr");
    assert_eq!(rpr.sz, Some(2400));
}

#[test]
fn test_parse_text_char_props_lang() {
    let xml = r#"<rPr lang="en-US"/>"#;
    let rpr: TextCharacterProperties = parse_from_xml(xml).expect("should parse lang rPr");
    assert_eq!(rpr.lang.as_deref(), Some("en-US"));
}

#[test]
fn test_parse_text_char_props_font_size_and_lang() {
    let xml = r#"<rPr lang="en-US" sz="3600"/>"#;
    let rpr: TextCharacterProperties = parse_from_xml(xml).expect("should parse sz+lang rPr");
    assert_eq!(rpr.sz, Some(3600));
    assert_eq!(rpr.lang.as_deref(), Some("en-US"));
}

// ---------------------------------------------------------------------------
// TextRun (a:r)
// ---------------------------------------------------------------------------

#[test]
fn test_parse_text_run_basic() {
    let xml = r#"<r><t>Hello World</t></r>"#;
    let run: TextRun = parse_from_xml(xml).expect("should parse TextRun");
    assert_eq!(run.t, "Hello World");
    assert!(run.r_pr.is_none());
}

#[test]
fn test_parse_text_run_empty_text() {
    let xml = r#"<r><t></t></r>"#;
    let run: TextRun = parse_from_xml(xml).expect("should parse TextRun with empty text");
    assert_eq!(run.t, "");
}

#[test]
fn test_parse_text_run_bold() {
    let xml = r#"<r><rPr b="1"/><t>Bold</t></r>"#;
    let run: TextRun = parse_from_xml(xml).expect("should parse bold TextRun");
    assert_eq!(run.t, "Bold");
    let rpr = run.r_pr.as_ref().expect("should have rPr");
    assert_eq!(rpr.b, Some(true));
}

#[test]
fn test_parse_text_run_italic() {
    let xml = r#"<r><rPr i="1"/><t>Italic</t></r>"#;
    let run: TextRun = parse_from_xml(xml).expect("should parse italic TextRun");
    assert_eq!(run.t, "Italic");
    let rpr = run.r_pr.as_ref().expect("should have rPr");
    assert_eq!(rpr.i, Some(true));
}

#[test]
fn test_parse_text_run_font_size() {
    let xml = r#"<r><rPr sz="2400"/><t>Big text</t></r>"#;
    let run: TextRun = parse_from_xml(xml).expect("should parse TextRun with font size");
    assert_eq!(run.t, "Big text");
    let rpr = run.r_pr.as_ref().expect("should have rPr");
    assert_eq!(rpr.sz, Some(2400));
}

#[test]
fn test_parse_text_run_with_namespace() {
    let xml = r#"<a:r xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:t>Namespaced</a:t></a:r>"#;
    let run: TextRun = parse_from_xml(xml).expect("should parse namespaced TextRun");
    assert_eq!(run.t, "Namespaced");
}

#[test]
fn test_parse_text_run_xml_entities() {
    // Entity refs (&amp; etc.) must be decoded correctly.
    let xml = r#"<r><t>&amp; &lt; &gt; &quot; &apos;</t></r>"#;
    let run: TextRun = parse_from_xml(xml).expect("should parse TextRun with entities");
    assert_eq!(run.t, "& < > \" '");
}

#[test]
fn test_parse_text_run_unicode() {
    let xml = r#"<r><t>café 中文 🌍</t></r>"#;
    let run: TextRun = parse_from_xml(xml).expect("should parse TextRun with unicode");
    assert_eq!(run.t, "café 中文 🌍");
}

// ---------------------------------------------------------------------------
// TextParagraph (a:p)
// ---------------------------------------------------------------------------

#[test]
fn test_parse_text_paragraph_empty() {
    let xml = r#"<p/>"#;
    let para: TextParagraph = parse_from_xml(xml).expect("should parse empty paragraph");
    assert!(para.p_pr.is_none());
    assert!(para.text_run.is_empty());
    assert!(para.end_para_r_pr.is_none());
}

#[test]
fn test_parse_text_paragraph_with_run() {
    let xml = r#"<p><r><t>Hello</t></r></p>"#;
    let para: TextParagraph = parse_from_xml(xml).expect("should parse paragraph with run");
    assert_eq!(para.text_run.len(), 1);
    match &para.text_run[0] {
        EGTextRun::R(run) => assert_eq!(run.t, "Hello"),
        other => panic!("expected R variant, got {:?}", other),
    }
}

#[test]
fn test_parse_text_paragraph_multiple_runs() {
    let xml = r#"<p><r><t>Hello </t></r><r><rPr b="1"/><t>World</t></r></p>"#;
    let para: TextParagraph =
        parse_from_xml(xml).expect("should parse paragraph with multiple runs");
    assert_eq!(para.text_run.len(), 2);
    match &para.text_run[0] {
        EGTextRun::R(run) => assert_eq!(run.t, "Hello "),
        other => panic!("expected R variant, got {:?}", other),
    }
    match &para.text_run[1] {
        EGTextRun::R(run) => {
            assert_eq!(run.t, "World");
            assert_eq!(run.r_pr.as_ref().and_then(|r| r.b), Some(true));
        }
        other => panic!("expected R variant, got {:?}", other),
    }
}

#[test]
fn test_parse_text_paragraph_with_end_rpr() {
    let xml = r#"<p><r><t>text</t></r><endParaRPr lang="en-US"/></p>"#;
    let para: TextParagraph = parse_from_xml(xml).expect("should parse paragraph with endParaRPr");
    assert_eq!(para.text_run.len(), 1);
    assert!(para.end_para_r_pr.is_some());
}

#[test]
fn test_parse_text_paragraph_with_namespace() {
    let xml = r#"<a:p xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:r><a:t>Hi</a:t></a:r></a:p>"#;
    let para: TextParagraph = parse_from_xml(xml).expect("should parse namespaced paragraph");
    assert_eq!(para.text_run.len(), 1);
}

// ---------------------------------------------------------------------------
// TextBody (a:txBody)
// ---------------------------------------------------------------------------

#[test]
fn test_parse_text_body_minimal() {
    // bodyPr is required; no paragraphs needed.
    let xml = r#"<txBody><bodyPr/></txBody>"#;
    let body: TextBody = parse_from_xml(xml).expect("should parse minimal TextBody");
    assert!(body.p.is_empty());
    assert!(body.lst_style.is_none());
}

#[test]
fn test_parse_text_body_with_paragraph() {
    let xml = r#"<txBody><bodyPr/><p><r><t>Hello</t></r></p></txBody>"#;
    let body: TextBody = parse_from_xml(xml).expect("should parse TextBody with paragraph");
    assert_eq!(body.p.len(), 1);
    assert_eq!(body.p[0].text_run.len(), 1);
    match &body.p[0].text_run[0] {
        EGTextRun::R(run) => assert_eq!(run.t, "Hello"),
        other => panic!("expected R variant, got {:?}", other),
    }
}

#[test]
fn test_parse_text_body_multiple_paragraphs() {
    let xml = r#"<txBody><bodyPr/><p><r><t>First</t></r></p><p><r><t>Second</t></r></p></txBody>"#;
    let body: TextBody =
        parse_from_xml(xml).expect("should parse TextBody with multiple paragraphs");
    assert_eq!(body.p.len(), 2);
    match &body.p[0].text_run[0] {
        EGTextRun::R(run) => assert_eq!(run.t, "First"),
        _ => panic!("expected R"),
    }
    match &body.p[1].text_run[0] {
        EGTextRun::R(run) => assert_eq!(run.t, "Second"),
        _ => panic!("expected R"),
    }
}

#[test]
fn test_parse_text_body_with_namespace() {
    let xml = r#"<a:txBody xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:bodyPr/><a:p><a:r><a:t>Hi</a:t></a:r></a:p></a:txBody>"#;
    let body: TextBody = parse_from_xml(xml).expect("should parse namespaced TextBody");
    assert_eq!(body.p.len(), 1);
}

#[test]
fn test_parse_text_body_empty_paragraph() {
    let xml = r#"<txBody><bodyPr/><p/></txBody>"#;
    let body: TextBody = parse_from_xml(xml).expect("should parse TextBody with empty paragraph");
    assert_eq!(body.p.len(), 1);
    assert!(body.p[0].text_run.is_empty());
}
