// Requires full feature set
#![cfg(feature = "full")]

//! Parity tests for generated event-based parsers vs serde deserialization.
//!
//! These tests ensure the generated FromXml parsers produce equivalent results
//! to quick-xml's serde deserialization.

use ooxml_sml::parsers::{FromXml, ParseError};
use ooxml_sml::types::*;
use quick_xml::Reader;
use quick_xml::de::from_str;
use quick_xml::events::Event;
use std::io::Cursor;

/// Helper to parse XML using the generated FromXml trait.
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

// =============================================================================
// AutoFilter tests
// =============================================================================

#[test]
fn test_autofilter_parity_simple() {
    let xml = r#"<autoFilter ref="A1:D10"/>"#;

    let serde_result: AutoFilter = from_str(xml).expect("serde parse failed");
    let fromxml_result: AutoFilter = parse_from_xml(xml).expect("fromxml parse failed");

    assert_eq!(serde_result.reference, fromxml_result.reference);
    assert_eq!(
        serde_result.filter_column.len(),
        fromxml_result.filter_column.len()
    );
}

#[test]
fn test_autofilter_parity_with_children() {
    let xml = r#"
    <autoFilter ref="A1:D10">
        <filterColumn colId="0"/>
        <filterColumn colId="1" hiddenButton="true"/>
    </autoFilter>
    "#;

    let serde_result: AutoFilter = from_str(xml).expect("serde parse failed");
    let fromxml_result: AutoFilter = parse_from_xml(xml).expect("fromxml parse failed");

    assert_eq!(serde_result.reference, fromxml_result.reference);
    assert_eq!(
        serde_result.filter_column.len(),
        fromxml_result.filter_column.len()
    );
    assert_eq!(serde_result.filter_column.len(), 2);

    // Compare filter columns
    for (s, f) in serde_result
        .filter_column
        .iter()
        .zip(fromxml_result.filter_column.iter())
    {
        assert_eq!(s.column_id, f.column_id);
        assert_eq!(s.hidden_button, f.hidden_button);
    }
}

// =============================================================================
// Row tests
// =============================================================================

#[test]
fn test_row_parity_simple() {
    let xml = r#"<row r="1" spans="1:5"/>"#;

    let serde_result: Row = from_str(xml).expect("serde parse failed");
    let fromxml_result: Row = parse_from_xml(xml).expect("fromxml parse failed");

    assert_eq!(serde_result.reference, fromxml_result.reference);
    assert_eq!(serde_result.cell_spans, fromxml_result.cell_spans);
}

#[test]
fn test_row_parity_with_cells() {
    let xml = r#"
    <row r="1" spans="1:3">
        <c r="A1" t="s"><v>0</v></c>
        <c r="B1"><v>42.5</v></c>
        <c r="C1" t="b"><v>1</v></c>
    </row>
    "#;

    let serde_result: Row = from_str(xml).expect("serde parse failed");
    let fromxml_result: Row = parse_from_xml(xml).expect("fromxml parse failed");

    assert_eq!(serde_result.reference, fromxml_result.reference);
    assert_eq!(serde_result.cells.len(), fromxml_result.cells.len());
    assert_eq!(serde_result.cells.len(), 3);

    // Compare cells
    for (s, f) in serde_result.cells.iter().zip(fromxml_result.cells.iter()) {
        assert_eq!(s.reference, f.reference);
        assert_eq!(s.cell_type, f.cell_type);
        // Compare value if present
        if let (Some(sv), Some(fv)) = (&s.value, &f.value) {
            assert_eq!(sv, fv);
        }
    }
}

// =============================================================================
// Cell tests
// =============================================================================

#[test]
fn test_cell_parity_simple() {
    let xml = r#"<c r="A1"><v>42</v></c>"#;

    let serde_result: Cell = from_str(xml).expect("serde parse failed");
    let fromxml_result: Cell = parse_from_xml(xml).expect("fromxml parse failed");

    assert_eq!(serde_result.reference, fromxml_result.reference);
    assert_eq!(serde_result.cell_type, fromxml_result.cell_type);
    assert_eq!(serde_result.style_index, fromxml_result.style_index);

    // Compare values
    assert!(serde_result.value.is_some());
    assert!(fromxml_result.value.is_some());
    assert_eq!(serde_result.value, fromxml_result.value);
}

#[test]
fn test_cell_parity_with_formula() {
    let xml = r#"<c r="A1"><f>SUM(B1:B10)</f><v>55</v></c>"#;

    let serde_result: Cell = from_str(xml).expect("serde parse failed");
    let fromxml_result: Cell = parse_from_xml(xml).expect("fromxml parse failed");

    assert_eq!(serde_result.reference, fromxml_result.reference);

    // Both should have formula
    assert!(serde_result.formula.is_some());
    assert!(fromxml_result.formula.is_some());

    // Both should have value
    assert!(serde_result.value.is_some());
    assert!(fromxml_result.value.is_some());
}

#[test]
fn test_cell_parity_with_type() {
    let xml = r#"<c r="A1" t="s" s="1"><v>0</v></c>"#;

    let serde_result: Cell = from_str(xml).expect("serde parse failed");
    let fromxml_result: Cell = parse_from_xml(xml).expect("fromxml parse failed");

    assert_eq!(serde_result.reference, fromxml_result.reference);
    assert_eq!(serde_result.cell_type, fromxml_result.cell_type);
    assert_eq!(serde_result.style_index, fromxml_result.style_index);
}

// =============================================================================
// SheetData tests
// =============================================================================

#[test]
fn test_sheetdata_parity() {
    let xml = r#"
    <sheetData>
        <row r="1" spans="1:3">
            <c r="A1" t="s"><v>0</v></c>
            <c r="B1"><v>123.45</v></c>
            <c r="C1" t="b"><v>1</v></c>
        </row>
        <row r="2">
            <c r="A2"><f>A1+B1</f><v>123.45</v></c>
        </row>
    </sheetData>
    "#;

    let serde_result: SheetData = from_str(xml).expect("serde parse failed");
    let fromxml_result: SheetData = parse_from_xml(xml).expect("fromxml parse failed");

    assert_eq!(serde_result.row.len(), fromxml_result.row.len());
    assert_eq!(serde_result.row.len(), 2);

    // Compare first row
    assert_eq!(
        serde_result.row[0].reference,
        fromxml_result.row[0].reference
    );
    assert_eq!(
        serde_result.row[0].cells.len(),
        fromxml_result.row[0].cells.len()
    );

    // Compare second row
    assert_eq!(
        serde_result.row[1].reference,
        fromxml_result.row[1].reference
    );
    assert_eq!(
        serde_result.row[1].cells.len(),
        fromxml_result.row[1].cells.len()
    );
}

// =============================================================================
// Filters tests
// =============================================================================

#[test]
fn test_filters_parity() {
    let xml = r#"<filters calendarType="gregorian"><filter val="test"/></filters>"#;

    let serde_result: Filters = from_str(xml).expect("serde parse failed");
    let fromxml_result: Filters = parse_from_xml(xml).expect("fromxml parse failed");

    assert_eq!(serde_result.calendar_type, fromxml_result.calendar_type);
    assert_eq!(serde_result.filter.len(), fromxml_result.filter.len());
}

// =============================================================================
// Large data test
// =============================================================================

#[test]
fn test_sheetdata_parity_large() {
    // Generate a larger dataset to stress test
    let mut xml = String::from("<sheetData>\n");
    for r in 1..=100 {
        xml.push_str(&format!("  <row r=\"{}\" spans=\"1:10\">\n", r));
        for c in 1..=10 {
            let col = col_letter(c);
            xml.push_str(&format!(
                "    <c r=\"{}{}\"><v>{}</v></c>\n",
                col,
                r,
                r * 10 + c
            ));
        }
        xml.push_str("  </row>\n");
    }
    xml.push_str("</sheetData>");

    let serde_result: SheetData = from_str(&xml).expect("serde parse failed");
    let fromxml_result: SheetData = parse_from_xml(&xml).expect("fromxml parse failed");

    assert_eq!(serde_result.row.len(), fromxml_result.row.len());
    assert_eq!(serde_result.row.len(), 100);

    // Verify all rows have correct cell counts
    for (s_row, f_row) in serde_result.row.iter().zip(fromxml_result.row.iter()) {
        assert_eq!(s_row.reference, f_row.reference);
        assert_eq!(s_row.cells.len(), f_row.cells.len());
        assert_eq!(s_row.cells.len(), 10);
    }
}

fn col_letter(col: usize) -> String {
    let mut result = String::new();
    let mut n = col;
    while n > 0 {
        n -= 1;
        result.insert(0, (b'A' + (n % 26) as u8) as char);
        n /= 26;
    }
    result
}
