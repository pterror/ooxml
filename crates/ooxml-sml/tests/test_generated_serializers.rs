//! Roundtrip tests for SML generated ToXml serializers.
//!
//! These tests verify that parsing XML via FromXml and then serializing
//! via ToXml produces equivalent XML output.

#![cfg(feature = "full")]

use ooxml_sml::parsers::{FromXml, ParseError};
use ooxml_sml::serializers::{SerializeError, ToXml};
use ooxml_sml::types::*;
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
// Cell tests
// ====================================================================

#[test]
fn test_serialize_empty_cell() {
    let xml = r#"<c r="A1"/>"#;
    let cell: Cell = parse_from_xml(xml).expect("parse empty cell");
    let out = serialize_to_xml(&cell, "c").unwrap();
    assert!(
        out.contains("r=\"A1\""),
        "should contain cell reference: {}",
        out
    );
}

#[test]
fn test_serialize_cell_with_value() {
    let xml = r#"<c r="A1"><v>42</v></c>"#;
    let cell: Cell = parse_from_xml(xml).expect("parse cell with value");
    let out = serialize_to_xml(&cell, "c").unwrap();
    assert!(out.contains("42"), "should contain value: {}", out);
    // The serializer may add namespace prefix (sml:v or just v)
    assert!(
        out.contains("<v>") || out.contains(":v>"),
        "should contain v element: {}",
        out
    );
}

#[test]
fn test_roundtrip_cell_with_type() {
    let xml = r#"<c r="B2" t="s"><v>0</v></c>"#;
    let rt: Cell = roundtrip(xml, "c");
    assert_eq!(rt.reference.as_deref(), Some("B2"));
    assert_eq!(rt.cell_type, Some(CellType::SharedString));
    assert_eq!(rt.value.as_deref(), Some("0"));
}

#[test]
fn test_roundtrip_cell_with_formula() {
    let xml = r#"<c r="C3"><f>A1+B1</f><v>100</v></c>"#;
    let rt: Cell = roundtrip(xml, "c");
    assert_eq!(rt.reference.as_deref(), Some("C3"));
    assert!(rt.formula.is_some());
    let formula = rt.formula.as_ref().unwrap();
    assert_eq!(formula.text.as_deref(), Some("A1+B1"));
    assert_eq!(rt.value.as_deref(), Some("100"));
}

#[test]
fn test_roundtrip_cell_with_style() {
    let xml = r#"<c r="A1" s="5"><v>123</v></c>"#;
    let rt: Cell = roundtrip(xml, "c");
    assert_eq!(rt.reference.as_deref(), Some("A1"));
    assert_eq!(rt.style_index, Some(5));
}

// ====================================================================
// Row tests
// ====================================================================

#[test]
fn test_serialize_empty_row() {
    let xml = r#"<row r="1"/>"#;
    let row: Row = parse_from_xml(xml).expect("parse empty row");
    let out = serialize_to_xml(&row, "row").unwrap();
    assert!(
        out.contains("r=\"1\""),
        "should contain row number: {}",
        out
    );
}

#[test]
fn test_roundtrip_row_with_cells() {
    let xml = r#"<row r="1"><c r="A1"><v>Hello</v></c><c r="B1"><v>42</v></c></row>"#;
    let rt: Row = roundtrip(xml, "row");
    assert_eq!(rt.reference, Some(1));
    assert_eq!(rt.cells.len(), 2);
    assert_eq!(rt.cells[0].reference.as_deref(), Some("A1"));
    assert_eq!(rt.cells[1].reference.as_deref(), Some("B1"));
}

#[test]
fn test_roundtrip_row_with_height() {
    let xml = r#"<row r="5" ht="20" customHeight="1"/>"#;
    let rt: Row = roundtrip(xml, "row");
    assert_eq!(rt.reference, Some(5));
    assert_eq!(rt.height, Some(20.0));
    assert_eq!(rt.custom_height, Some(true));
}

#[test]
fn test_roundtrip_row_with_style() {
    let xml = r#"<row r="3" s="2" customFormat="1"/>"#;
    let rt: Row = roundtrip(xml, "row");
    assert_eq!(rt.reference, Some(3));
    assert_eq!(rt.style_index, Some(2));
    assert_eq!(rt.custom_format, Some(true));
}

// ====================================================================
// SheetData tests
// ====================================================================

#[test]
fn test_serialize_empty_sheet_data() {
    let xml = r#"<sheetData/>"#;
    let sd: SheetData = parse_from_xml(xml).expect("parse empty sheetData");
    let out = serialize_to_xml(&sd, "sheetData").unwrap();
    assert_eq!(out, "<sheetData/>");
}

#[test]
fn test_roundtrip_sheet_data_with_rows() {
    let xml = r#"<sheetData><row r="1"><c r="A1"><v>1</v></c></row><row r="2"><c r="A2"><v>2</v></c></row></sheetData>"#;
    let rt: SheetData = roundtrip(xml, "sheetData");
    assert_eq!(rt.row.len(), 2);
    assert_eq!(rt.row[0].reference, Some(1));
    assert_eq!(rt.row[1].reference, Some(2));
}

// ====================================================================
// Worksheet tests
// ====================================================================

#[test]
fn test_roundtrip_worksheet_minimal() {
    let xml = r#"<worksheet><sheetData/></worksheet>"#;
    let rt: Worksheet = roundtrip(xml, "worksheet");
    assert!(rt.sheet_data.row.is_empty());
}

#[test]
fn test_roundtrip_worksheet_with_dimension() {
    let xml = r#"<worksheet><dimension ref="A1:C10"/><sheetData/></worksheet>"#;
    let rt: Worksheet = roundtrip(xml, "worksheet");
    let dim = rt.dimension.expect("should have dimension");
    assert_eq!(dim.reference.as_str(), "A1:C10");
}

#[test]
fn test_roundtrip_worksheet_with_data() {
    let xml = r#"<worksheet><sheetData><row r="1"><c r="A1" t="s"><v>0</v></c><c r="B1"><v>42</v></c></row></sheetData></worksheet>"#;
    let rt: Worksheet = roundtrip(xml, "worksheet");
    assert_eq!(rt.sheet_data.row.len(), 1);
    let row = &rt.sheet_data.row[0];
    assert_eq!(row.cells.len(), 2);
}

#[test]
fn test_roundtrip_worksheet_with_merged_cells() {
    let xml = r#"<worksheet><sheetData/><mergeCells count="1"><mergeCell ref="A1:B2"/></mergeCells></worksheet>"#;
    let rt: Worksheet = roundtrip(xml, "worksheet");
    let mc = rt.merged_cells.expect("should have merge cells");
    assert_eq!(mc.merge_cell.len(), 1);
    assert_eq!(mc.merge_cell[0].reference.as_str(), "A1:B2");
}

// ====================================================================
// SharedStrings tests
// ====================================================================

#[test]
fn test_roundtrip_shared_strings_simple() {
    let xml = r#"<sst count="2" uniqueCount="2"><si><t>Hello</t></si><si><t>World</t></si></sst>"#;
    let rt: SharedStrings = roundtrip(xml, "sst");
    assert_eq!(rt.count, Some(2));
    assert_eq!(rt.unique_count, Some(2));
    assert_eq!(rt.si.len(), 2);
}

#[test]
fn test_roundtrip_shared_strings_with_rich_text() {
    let xml =
        r#"<sst count="1" uniqueCount="1"><si><r><t>Bold</t></r><r><t>Normal</t></r></si></sst>"#;
    let rt: SharedStrings = roundtrip(xml, "sst");
    assert_eq!(rt.si.len(), 1);
    let si = &rt.si[0];
    assert_eq!(si.reference.len(), 2);
}

// ====================================================================
// Workbook tests
// ====================================================================

#[test]
fn test_roundtrip_workbook_minimal() {
    let xml =
        r#"<workbook><sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets></workbook>"#;
    let rt: Workbook = roundtrip(xml, "workbook");
    assert_eq!(rt.sheets.sheet.len(), 1);
    assert_eq!(rt.sheets.sheet[0].name.as_str(), "Sheet1");
}

#[test]
fn test_roundtrip_workbook_with_properties() {
    let xml = r#"<workbook><workbookPr date1904="1"/><sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets></workbook>"#;
    let rt: Workbook = roundtrip(xml, "workbook");
    let props = rt.workbook_pr.expect("should have workbook properties");
    assert_eq!(props.date1904, Some(true));
}

// ====================================================================
// Stylesheet tests
// ====================================================================

#[test]
fn test_roundtrip_stylesheet_minimal() {
    let xml = r#"<styleSheet/>"#;
    let rt: Stylesheet = roundtrip(xml, "styleSheet");
    assert!(rt.fonts.is_none());
    assert!(rt.fills.is_none());
}

#[test]
fn test_roundtrip_stylesheet_with_fonts() {
    let xml = r#"<styleSheet><fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts></styleSheet>"#;
    let rt: Stylesheet = roundtrip(xml, "styleSheet");
    let fonts = rt.fonts.expect("should have fonts");
    assert_eq!(fonts.font.len(), 1);
}

// ====================================================================
// Serialize output format verification
// ====================================================================

#[test]
fn test_serialize_preserves_namespace_prefix() {
    let xml =
        r#"<worksheet><sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData></worksheet>"#;
    let ws: Worksheet = parse_from_xml(xml).expect("parse worksheet");
    let out = serialize_to_xml(&ws, "worksheet").unwrap();
    // Verify structure is preserved (may have sml: namespace prefix)
    assert!(
        out.contains("sheetData>"),
        "should contain sheetData: {}",
        out
    );
    assert!(
        out.contains("<row") || out.contains(":row"),
        "should contain row: {}",
        out
    );
    assert!(
        out.contains("<c") || out.contains(":c"),
        "should contain cell: {}",
        out
    );
    assert!(
        out.contains(":v>") || out.contains("<v>"),
        "should contain value: {}",
        out
    );
}

#[test]
fn test_serialize_self_closing_empty_elements() {
    let xml = r#"<sheetData/>"#;
    let sd: SheetData = parse_from_xml(xml).expect("parse sheetData");
    let out = serialize_to_xml(&sd, "sheetData").unwrap();
    // Empty sheetData should self-close
    assert_eq!(out, "<sheetData/>");
}

#[test]
fn test_serialize_boolean_attributes() {
    let xml = r#"<row r="1" hidden="1" customHeight="1"/>"#;
    let row: Row = parse_from_xml(xml).expect("parse row");
    let out = serialize_to_xml(&row, "row").unwrap();
    assert!(
        out.contains("hidden="),
        "should contain hidden attr: {}",
        out
    );
    assert!(
        out.contains("customHeight="),
        "should contain customHeight attr: {}",
        out
    );
}

#[test]
fn test_serialize_cell_formula_attributes() {
    let xml = r#"<c r="A1"><f t="shared" ref="A1:A10" si="0">ROW()</f><v>1</v></c>"#;
    let cell: Cell = parse_from_xml(xml).expect("parse cell");
    let out = serialize_to_xml(&cell, "c").unwrap();
    assert!(
        out.contains("t=\"shared\"") || out.contains("t='shared'"),
        "should contain formula type: {}",
        out
    );
    assert!(
        out.contains("ROW()"),
        "should contain formula text: {}",
        out
    );
}
