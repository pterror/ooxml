//! Error-handling tests for malformed / adversarial DOCX inputs.
//!
//! These tests verify that the WML reader produces useful errors rather than
//! panicking when given invalid data.

use std::io::Cursor;

type Doc = ooxml_wml::Document<Cursor<Vec<u8>>>;

/// Feeding random bytes (not a ZIP) should produce an error, not a panic.
#[test]
fn test_not_a_zip() {
    let garbage = vec![0u8; 256];
    let result = Doc::from_reader(Cursor::new(garbage));
    assert!(result.is_err(), "should fail on non-ZIP input");
}

/// An empty byte stream should fail gracefully.
#[test]
fn test_empty_input() {
    let result = Doc::from_reader(Cursor::new(Vec::new()));
    assert!(result.is_err(), "should fail on empty input");
}

/// A valid ZIP with no OOXML parts should fail.
#[test]
fn test_empty_zip() {
    let mut buf = Cursor::new(Vec::new());
    {
        let _writer = zip::ZipWriter::new(&mut buf);
        // writer drops, finalizing the empty ZIP
    }
    let bytes = buf.into_inner();
    let result = Doc::from_reader(Cursor::new(bytes));
    assert!(result.is_err(), "should fail on empty ZIP");
}

/// A ZIP containing [Content_Types].xml but no document.xml.
#[test]
fn test_missing_document_xml() {
    let mut buf = Cursor::new(Vec::new());
    {
        use std::io::Write;
        let mut writer = zip::ZipWriter::new(&mut buf);
        let options = zip::write::SimpleFileOptions::default();
        writer.start_file("[Content_Types].xml", options).unwrap();
        writer
            .write_all(
                br#"<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
</Types>"#,
            )
            .unwrap();
        writer.finish().unwrap();
    }
    let bytes = buf.into_inner();
    let result = Doc::from_reader(Cursor::new(bytes));
    assert!(result.is_err(), "should fail when document.xml is missing");
}

/// A ZIP with document.xml containing malformed XML.
#[test]
fn test_malformed_xml_in_document() {
    let mut buf = Cursor::new(Vec::new());
    {
        use std::io::Write;
        let mut writer = zip::ZipWriter::new(&mut buf);
        let options = zip::write::SimpleFileOptions::default();

        writer.start_file("[Content_Types].xml", options).unwrap();
        writer
            .write_all(
                br#"<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>"#,
            )
            .unwrap();

        writer.start_file("_rels/.rels", options).unwrap();
        writer
            .write_all(
                br#"<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>"#,
            )
            .unwrap();

        writer.start_file("word/document.xml", options).unwrap();
        writer
            .write_all(b"this is not XML at all & has no elements")
            .unwrap();

        writer.finish().unwrap();
    }
    let bytes = buf.into_inner();
    let result = Doc::from_reader(Cursor::new(bytes));
    assert!(
        result.is_err(),
        "should fail when document.xml has malformed XML"
    );
}
