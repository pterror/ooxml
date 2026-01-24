//! Roundtrip tests for SML (SpreadsheetML) codegen.
//!
//! These tests verify that parsing worksheet XML produces correct structures
//! that can be compared against expected values. Full XML roundtripping uses
//! the existing writer infrastructure.

// These tests require the full feature set
#![cfg(feature = "full")]

use ooxml_sml::parsers::{FromXml, ParseError};
use ooxml_sml::types::{CellType, Row, Worksheet};
use quick_xml::Reader;
use quick_xml::events::Event;
use std::io::Cursor;

fn parse_worksheet_xml(xml: &[u8]) -> Result<Worksheet, ParseError> {
    let mut reader = Reader::from_reader(Cursor::new(xml));
    let mut buf = Vec::new();

    loop {
        buf.clear();
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) if e.name().as_ref() == b"worksheet" => {
                return Worksheet::from_xml(&mut reader, &e, false);
            }
            Ok(Event::Empty(e)) if e.name().as_ref() == b"worksheet" => {
                return Worksheet::from_xml(&mut reader, &e, true);
            }
            Ok(Event::Eof) => {
                return Err(ParseError::UnexpectedElement(
                    "EOF before worksheet".to_string(),
                ));
            }
            Err(e) => return Err(ParseError::Xml(e)),
            _ => {}
        }
    }
}

#[test]
fn test_parse_minimal_worksheet() {
    let xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
    <sheetData/>
</worksheet>"#;

    let ws = parse_worksheet_xml(xml).expect("should parse");
    assert!(ws.sheet_data.row.is_empty());
}

#[test]
fn test_parse_worksheet_with_dimension() {
    let xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
    <dimension ref="A1:B2"/>
    <sheetData/>
</worksheet>"#;

    let ws = parse_worksheet_xml(xml).expect("should parse");
    let dim = ws.dimension.expect("should have dimension");
    assert_eq!(dim.reference.as_str(), "A1:B2");
}

#[test]
fn test_parse_worksheet_with_rows() {
    let xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
    <sheetData>
        <row r="1">
            <c r="A1" t="s"><v>0</v></c>
            <c r="B1"><v>42</v></c>
        </row>
        <row r="2">
            <c r="A2" t="s"><v>1</v></c>
            <c r="B2"><v>100</v></c>
        </row>
    </sheetData>
</worksheet>"#;

    let ws = parse_worksheet_xml(xml).expect("should parse");
    assert_eq!(ws.sheet_data.row.len(), 2);

    let row1: &Row = &ws.sheet_data.row[0];
    assert_eq!(row1.reference, Some(1));
    assert_eq!(row1.cells.len(), 2);

    let cell_a1 = &row1.cells[0];
    assert_eq!(cell_a1.reference.as_deref(), Some("A1"));
    assert_eq!(cell_a1.cell_type, Some(CellType::SharedString));
    assert_eq!(cell_a1.value.as_deref(), Some("0"));

    let cell_b1 = &row1.cells[1];
    assert_eq!(cell_b1.reference.as_deref(), Some("B1"));
    assert!(cell_b1.cell_type.is_none()); // Numbers don't need type
    assert_eq!(cell_b1.value.as_deref(), Some("42"));
}

#[test]
fn test_parse_worksheet_with_formula() {
    let xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
    <sheetData>
        <row r="1">
            <c r="A1"><v>10</v></c>
            <c r="B1"><v>20</v></c>
            <c r="C1"><f>A1+B1</f><v>30</v></c>
        </row>
    </sheetData>
</worksheet>"#;

    let ws = parse_worksheet_xml(xml).expect("should parse");
    let row = &ws.sheet_data.row[0];

    let cell_c1 = &row.cells[2];
    assert_eq!(cell_c1.reference.as_deref(), Some("C1"));
    let formula = cell_c1.formula.as_ref().expect("should have formula");
    assert_eq!(formula.text.as_str(), "A1+B1");
    assert_eq!(cell_c1.value.as_deref(), Some("30"));
}

#[test]
fn test_parse_worksheet_with_merged_cells() {
    let xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
    <sheetData>
        <row r="1">
            <c r="A1" t="s"><v>0</v></c>
        </row>
    </sheetData>
    <mergeCells count="2">
        <mergeCell ref="A1:C1"/>
        <mergeCell ref="A3:B4"/>
    </mergeCells>
</worksheet>"#;

    let ws = parse_worksheet_xml(xml).expect("should parse");
    let merged = ws.merged_cells.expect("should have merged cells");
    assert_eq!(merged.merge_cell.len(), 2);
    assert_eq!(merged.merge_cell[0].reference.as_str(), "A1:C1");
    assert_eq!(merged.merge_cell[1].reference.as_str(), "A3:B4");
}

#[test]
fn test_parse_worksheet_with_sheet_views() {
    let xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
    <sheetViews>
        <sheetView tabSelected="1" workbookViewId="0"/>
    </sheetViews>
    <sheetData/>
</worksheet>"#;

    let ws = parse_worksheet_xml(xml).expect("should parse");
    let views = ws.sheet_views.expect("should have sheet views");
    assert_eq!(views.sheet_view.len(), 1);

    let sv = &views.sheet_view[0];
    assert_eq!(sv.tab_selected, Some(true));
    assert_eq!(sv.workbook_view_id, 0);
}

#[test]
fn test_parse_worksheet_preserves_cell_styles() {
    let xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
    <sheetData>
        <row r="1" s="5" customFormat="1">
            <c r="A1" s="3"><v>100</v></c>
        </row>
    </sheetData>
</worksheet>"#;

    let ws = parse_worksheet_xml(xml).expect("should parse");
    let row = &ws.sheet_data.row[0];
    assert_eq!(row.style_index, Some(5));
    assert_eq!(row.custom_format, Some(true));

    let cell = &row.cells[0];
    assert_eq!(cell.style_index, Some(3));
}

/// Test that parsing and writing via WorkbookBuilder produces roundtrippable output.
#[test]
fn test_full_roundtrip_via_writer() {
    use ooxml_sml::{Workbook, WorkbookBuilder};
    use std::io::Cursor;

    // Create a workbook with various features
    let mut wb = WorkbookBuilder::new();
    let sheet = wb.add_sheet("TestSheet");
    sheet.set_cell("A1", "Header");
    sheet.set_cell("B1", 42.5);
    sheet.set_cell("C1", true);
    sheet.set_formula("D1", "A1&B1");
    sheet.merge_cells("A2:C2");

    // Write to memory
    let mut buffer = Cursor::new(Vec::new());
    wb.write(&mut buffer).expect("write should succeed");

    // Read back
    buffer.set_position(0);
    let mut workbook = Workbook::from_reader(buffer).expect("read should succeed");
    let sheet = workbook.resolved_sheet(0).expect("sheet should exist");

    // Verify data survived roundtrip
    assert_eq!(sheet.name(), "TestSheet");
    assert_eq!(sheet.value_at("A1"), Some("Header".to_string()));
    assert_eq!(sheet.number_at("B1"), Some(42.5));
    assert!(sheet.has_merged_cells());
}

// =============================================================================
// Serde roundtrip tests: serialize → deserialize
// =============================================================================

use ooxml_sml::types::*;
use quick_xml::de::from_str;
use quick_xml::se::Serializer;
use serde::Serialize;

/// Serialize a value to XML string.
fn to_xml_string<T: Serialize>(value: &T) -> String {
    let mut buffer = String::new();
    let mut ser = Serializer::new(&mut buffer);
    ser.indent(' ', 2);
    value.serialize(ser).expect("serialization should succeed");
    buffer
}

#[test]
fn test_serde_roundtrip_worksheet() {
    // Parse with event-based parser
    let xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
    <sheetData>
        <row r="1">
            <c r="A1" t="s"><v>0</v></c>
            <c r="B1"><v>42</v></c>
        </row>
    </sheetData>
</worksheet>"#;

    let ws1 = parse_worksheet_xml(xml).expect("should parse");

    // Serialize with serde
    let serialized = to_xml_string(&ws1);

    // Deserialize with serde
    let ws2: Worksheet = from_str(&serialized).expect("serde should deserialize");

    // Compare
    assert_eq!(ws1.sheet_data.row.len(), ws2.sheet_data.row.len());
    assert_eq!(
        ws1.sheet_data.row[0].cells.len(),
        ws2.sheet_data.row[0].cells.len()
    );
}

#[test]
fn test_serde_roundtrip_cell() {
    // Parse a cell from worksheet XML
    let xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
    <sheetData>
        <row r="1">
            <c r="A1" t="s" s="5"><v>Hello</v></c>
        </row>
    </sheetData>
</worksheet>"#;

    let ws = parse_worksheet_xml(xml).expect("should parse");
    let cell1 = &ws.sheet_data.row[0].cells[0];

    // Serialize
    let serialized = to_xml_string(cell1);

    // Deserialize
    let cell2: Cell = from_str(&serialized).expect("serde should deserialize");

    // Compare
    assert_eq!(cell1.reference, cell2.reference);
    assert_eq!(cell1.cell_type, cell2.cell_type);
    assert_eq!(cell1.style_index, cell2.style_index);
    assert_eq!(cell1.value, cell2.value);
}

#[test]
fn test_serde_roundtrip_row() {
    // Parse a row
    let xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
    <sheetData>
        <row r="5" spans="1:10" s="2" customFormat="1" ht="20.5" customHeight="1">
            <c r="A5"><v>100</v></c>
            <c r="B5" t="s"><v>1</v></c>
        </row>
    </sheetData>
</worksheet>"#;

    let ws = parse_worksheet_xml(xml).expect("should parse");
    let row1 = &ws.sheet_data.row[0];

    // Serialize just the row
    let serialized = to_xml_string(row1);

    // Deserialize
    let row2: Row = from_str(&serialized).expect("serde should deserialize row");

    // Compare
    assert_eq!(row1.reference, row2.reference);
    assert_eq!(row1.cell_spans, row2.cell_spans);
    assert_eq!(row1.style_index, row2.style_index);
    assert_eq!(row1.custom_format, row2.custom_format);
    assert_eq!(row1.height, row2.height);
    assert_eq!(row1.custom_height, row2.custom_height);
    assert_eq!(row1.cells.len(), row2.cells.len());
}

#[test]
fn test_serde_roundtrip_merged_cells() {
    let xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
    <sheetData/>
    <mergeCells count="2">
        <mergeCell ref="A1:C1"/>
        <mergeCell ref="D5:F10"/>
    </mergeCells>
</worksheet>"#;

    let ws = parse_worksheet_xml(xml).expect("should parse");
    let merged1 = ws.merged_cells.as_ref().expect("should have merged cells");

    // Serialize
    let serialized = to_xml_string(merged1);

    // Deserialize
    let merged2: MergedCells = from_str(&serialized).expect("serde should deserialize");

    // Compare
    assert_eq!(merged1.merge_cell.len(), merged2.merge_cell.len());
    for (m1, m2) in merged1.merge_cell.iter().zip(merged2.merge_cell.iter()) {
        assert_eq!(m1.reference, m2.reference);
    }
}

#[test]
fn test_serde_roundtrip_formula() {
    let xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
    <sheetData>
        <row r="1">
            <c r="C1"><f>SUM(A1:B1)</f><v>100</v></c>
        </row>
    </sheetData>
</worksheet>"#;

    let ws = parse_worksheet_xml(xml).expect("should parse");
    let formula1 = ws.sheet_data.row[0].cells[0]
        .formula
        .as_ref()
        .expect("should have formula");

    // Serialize
    let serialized = to_xml_string(formula1);

    // Deserialize
    let formula2: CellFormula = from_str(&serialized).expect("serde should deserialize");

    // Compare
    assert_eq!(formula1.text, formula2.text);
}

#[test]
fn test_serde_roundtrip_sheet_views() {
    let xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
    <sheetViews>
        <sheetView tabSelected="1" workbookViewId="0" zoomScale="85">
            <pane xSplit="1" ySplit="2" topLeftCell="B3" activePane="bottomRight" state="frozen"/>
        </sheetView>
    </sheetViews>
    <sheetData/>
</worksheet>"#;

    let ws = parse_worksheet_xml(xml).expect("should parse");
    let views1 = ws.sheet_views.as_ref().expect("should have sheet views");

    // Serialize
    let serialized = to_xml_string(views1);

    // Deserialize
    let views2: SheetViews = from_str(&serialized).expect("serde should deserialize");

    // Compare
    assert_eq!(views1.sheet_view.len(), views2.sheet_view.len());
    let sv1 = &views1.sheet_view[0];
    let sv2 = &views2.sheet_view[0];
    assert_eq!(sv1.tab_selected, sv2.tab_selected);
    assert_eq!(sv1.workbook_view_id, sv2.workbook_view_id);
    assert_eq!(sv1.zoom_scale, sv2.zoom_scale);

    // Pane comparison
    let pane1 = sv1.pane.as_ref().expect("should have pane");
    let pane2 = sv2.pane.as_ref().expect("should have pane");
    assert_eq!(pane1.x_split, pane2.x_split);
    assert_eq!(pane1.y_split, pane2.y_split);
    assert_eq!(pane1.top_left_cell, pane2.top_left_cell);
    assert_eq!(pane1.active_pane, pane2.active_pane);
    assert_eq!(pane1.state, pane2.state);
}

#[test]
fn test_boolean_serialize_format() {
    // OOXML uses "1"/"0" for booleans, but serde outputs "true"/"false"
    // This test documents current behavior - a proper fix requires custom
    // serialization in codegen
    let xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
    <sheetViews>
        <sheetView tabSelected="1" showGridLines="0" workbookViewId="0"/>
    </sheetViews>
    <sheetData/>
</worksheet>"#;

    let ws = parse_worksheet_xml(xml).expect("should parse");
    let serialized = to_xml_string(&ws);

    // Verify booleans are serialized (current behavior: "true"/"false")
    // TODO: Fix to output "1"/"0" for OOXML compliance
    // See: ECMA-376 Part 1, section 22.9.2.1 (xsd:boolean)
    assert!(
        serialized.contains("tabSelected="),
        "tabSelected should be present"
    );

    // Verify we can roundtrip (the important thing is data integrity)
    let ws2: Worksheet = from_str(&serialized).expect("serde should deserialize");
    let sv1 = &ws.sheet_views.as_ref().unwrap().sheet_view[0];
    let sv2 = &ws2.sheet_views.as_ref().unwrap().sheet_view[0];
    assert_eq!(sv1.tab_selected, sv2.tab_selected);
    assert_eq!(sv1.show_grid_lines, sv2.show_grid_lines);
}
