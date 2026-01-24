//! Tests using XML fixtures for comprehensive parsing verification.
//!
//! These tests load XML fixtures and verify that the generated parsers
//! correctly handle various OOXML features and edge cases.

// These tests require the full feature set
#![cfg(feature = "full")]

use ooxml_sml::parsers::{FromXml, ParseError};
use ooxml_sml::types::Worksheet;
use quick_xml::Reader;
use quick_xml::events::Event;
use std::io::Cursor;

/// Parse worksheet XML from bytes.
fn parse_worksheet(xml: &[u8]) -> Result<Worksheet, ParseError> {
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

/// Load a fixture file from the fixtures/xml directory.
fn load_fixture(name: &str) -> Vec<u8> {
    let path = format!(
        "{}/tests/fixtures/xml/{}.xml",
        env!("CARGO_MANIFEST_DIR"),
        name
    );
    std::fs::read(&path).unwrap_or_else(|e| panic!("Failed to load fixture {}: {}", name, e))
}

#[test]
fn test_fixture_basic_cells() {
    let xml = load_fixture("basic_cells");
    let ws = parse_worksheet(&xml).expect("should parse basic_cells");

    // Should have 2 rows
    assert_eq!(ws.sheet_data.row.len(), 2);

    // First row has 6 cells
    let row1 = &ws.sheet_data.row[0];
    assert_eq!(row1.cells.len(), 6);

    // Check cell types
    use ooxml_sml::types::CellType;

    // A1: number (implicit, no type)
    assert!(row1.cells[0].cell_type.is_none());
    assert_eq!(row1.cells[0].value.as_deref(), Some("42.5"));

    // B1: explicit number
    assert_eq!(row1.cells[1].cell_type, Some(CellType::Number));

    // C1: shared string
    assert_eq!(row1.cells[2].cell_type, Some(CellType::SharedString));

    // D1: boolean
    assert_eq!(row1.cells[3].cell_type, Some(CellType::Boolean));
    assert_eq!(row1.cells[3].value.as_deref(), Some("1"));

    // E1: error
    assert_eq!(row1.cells[4].cell_type, Some(CellType::Error));
    assert_eq!(row1.cells[4].value.as_deref(), Some("#DIV/0!"));

    // F1: inline string
    assert_eq!(row1.cells[5].cell_type, Some(CellType::InlineString));
}

#[test]
fn test_fixture_formulas() {
    let xml = load_fixture("formulas");
    let ws = parse_worksheet(&xml).expect("should parse formulas");

    // Check that formulas are parsed
    let row1 = &ws.sheet_data.row[0];

    // C1 has simple formula
    let cell_c1 = &row1.cells[2];
    assert!(cell_c1.formula.is_some());
    let formula = cell_c1.formula.as_ref().unwrap();
    assert_eq!(formula.text.as_str(), "A1+B1");
    assert_eq!(cell_c1.value.as_deref(), Some("30"));

    // D1 has cross-sheet reference
    let cell_d1 = &row1.cells[3];
    assert!(cell_d1.formula.is_some());
    assert!(cell_d1.formula.as_ref().unwrap().text.contains("Sheet2"));
}

#[test]
fn test_fixture_merged_cells() {
    let xml = load_fixture("merged_cells");
    let ws = parse_worksheet(&xml).expect("should parse merged_cells");

    // Should have merged cells
    let merged = ws.merged_cells.expect("should have merged cells");
    assert_eq!(merged.merge_cell.len(), 3);

    // Check merge references
    let refs: Vec<_> = merged
        .merge_cell
        .iter()
        .map(|m| m.reference.as_str())
        .collect();
    assert!(refs.contains(&"A1:D1"));
    assert!(refs.contains(&"A4:A5"));
    assert!(refs.contains(&"D4:E5"));
}

#[test]
fn test_fixture_freeze_panes() {
    let xml = load_fixture("freeze_panes");
    let ws = parse_worksheet(&xml).expect("should parse freeze_panes");

    // Should have sheet views
    let views = ws.sheet_views.expect("should have sheet views");
    assert_eq!(views.sheet_view.len(), 1);

    let view = &views.sheet_view[0];
    assert_eq!(view.tab_selected, Some(true));

    // Check pane
    let pane = view.pane.as_ref().expect("should have pane");
    assert_eq!(pane.y_split, Some(1.0));
    use ooxml_sml::types::PaneState;
    assert_eq!(pane.state, Some(PaneState::Frozen));
}

#[test]
fn test_fixture_freeze_panes_both() {
    let xml = load_fixture("freeze_panes_both");
    let ws = parse_worksheet(&xml).expect("should parse freeze_panes_both");

    let views = ws.sheet_views.expect("should have sheet views");
    let pane = views.sheet_view[0].pane.as_ref().expect("should have pane");

    // Both row and column frozen
    assert_eq!(pane.x_split, Some(1.0));
    assert_eq!(pane.y_split, Some(1.0));
    assert_eq!(pane.top_left_cell.as_deref(), Some("B2"));
}

#[test]
fn test_fixture_dimensions() {
    let xml = load_fixture("dimensions");
    let ws = parse_worksheet(&xml).expect("should parse dimensions");

    // Check dimension
    let dim = ws.dimension.expect("should have dimension");
    assert_eq!(dim.reference.as_str(), "A1:D10");

    // Check columns are parsed
    assert!(!ws.cols.is_empty());
}

#[test]
fn test_fixture_data_validation() {
    let xml = load_fixture("data_validation");
    let ws = parse_worksheet(&xml).expect("should parse data_validation");

    // Should have data validations
    let dvs = ws.data_validations.expect("should have data validations");
    assert_eq!(dvs.data_validation.len(), 5);

    // First is a list validation
    use ooxml_sml::types::ValidationType;
    let dv0 = &dvs.data_validation[0];
    assert_eq!(dv0.r#type, Some(ValidationType::List));
    assert_eq!(dv0.square_reference.as_str(), "A2:A10");
}

#[test]
fn test_fixture_conditional_formatting() {
    let xml = load_fixture("conditional_formatting");
    let ws = parse_worksheet(&xml).expect("should parse conditional_formatting");

    // Should have conditional formatting
    let cfs = &ws.conditional_formatting;
    assert_eq!(cfs.len(), 4);

    // First has cell value rules
    assert_eq!(cfs[0].square_reference.as_deref(), Some("A1:A10"));
    assert_eq!(cfs[0].cf_rule.len(), 2);
}

#[test]
fn test_fixture_hyperlinks() {
    let xml = load_fixture("hyperlinks");
    let ws = parse_worksheet(&xml).expect("should parse hyperlinks");

    // Should have hyperlinks
    let links = ws.hyperlinks.expect("should have hyperlinks");
    assert_eq!(links.hyperlink.len(), 5);

    // Check internal link
    let link2 = &links.hyperlink[1];
    assert_eq!(link2.location.as_deref(), Some("Sheet2!A1"));
}

#[test]
fn test_fixture_edge_sparse_data() {
    let xml = load_fixture("edge_sparse_data");
    let ws = parse_worksheet(&xml).expect("should parse edge_sparse_data");

    // Should handle sparse rows correctly
    let rows = &ws.sheet_data.row;
    assert_eq!(rows.len(), 4);

    // Check row numbers (should be 1, 100, 500, 1000)
    assert_eq!(rows[0].reference, Some(1));
    assert_eq!(rows[1].reference, Some(100));
    assert_eq!(rows[2].reference, Some(500));
    assert_eq!(rows[3].reference, Some(1000));

    // Check sparse column references
    let row500 = &rows[2];
    assert!(
        row500
            .cells
            .iter()
            .any(|c| c.reference.as_deref() == Some("AA500"))
    );
    assert!(
        row500
            .cells
            .iter()
            .any(|c| c.reference.as_deref() == Some("ZZ500"))
    );
}

#[test]
fn test_fixture_edge_unicode() {
    let xml = load_fixture("edge_unicode");
    let ws = parse_worksheet(&xml).expect("should parse edge_unicode");

    // Should have 10 rows of unicode content
    assert_eq!(ws.sheet_data.row.len(), 10);

    // Verify inline strings are parsed (checking they have inline content)
    for row in &ws.sheet_data.row {
        let cell = &row.cells[0];
        assert!(
            cell.is.is_some(),
            "Row {} should have inline string",
            row.reference.unwrap_or(0)
        );
    }
}

#[test]
fn test_fixture_edge_empty_elements() {
    let xml = load_fixture("edge_empty_elements");
    let ws = parse_worksheet(&xml).expect("should parse edge_empty_elements");

    // Should handle empty elements gracefully
    assert_eq!(ws.sheet_data.row.len(), 3);

    // First row has empty cells
    let row1 = &ws.sheet_data.row[0];
    assert_eq!(row1.cells.len(), 3);

    // Second row is empty but exists
    let row2 = &ws.sheet_data.row[1];
    assert!(row2.cells.is_empty());

    // Merged cells should be empty (count=0)
    if let Some(merged) = &ws.merged_cells {
        assert!(merged.merge_cell.is_empty());
    }
}

#[test]
fn test_fixture_edge_large_numbers() {
    let xml = load_fixture("edge_large_numbers");
    let ws = parse_worksheet(&xml).expect("should parse edge_large_numbers");

    assert_eq!(ws.sheet_data.row.len(), 10);

    // All cells should have values
    for row in &ws.sheet_data.row {
        assert!(
            row.cells[0].value.is_some(),
            "Row {} should have value",
            row.reference.unwrap_or(0)
        );
    }

    // Check specific values are preserved
    assert_eq!(
        ws.sheet_data.row[0].cells[0].value.as_deref(),
        Some("9999999999999999")
    );
    assert_eq!(
        ws.sheet_data.row[6].cells[0].value.as_deref(),
        Some("3.141592653589793")
    );
}
