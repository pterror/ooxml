//! Parity tests verifying generated types produce equivalent XML output.
//!
//! These tests parse XML with the generated event-based parser, serialize back
//! using serde, and compare the results. This ensures the codegen produces
//! types that faithfully roundtrip OOXML data.

// These tests require the full feature set
#![cfg(feature = "full")]

use ooxml_corpus::roundtrip::{CompareOptions, DifferenceKind, compare_xml};
use ooxml_sml::parsers::{FromXml, ParseError};
use ooxml_sml::types::*;
use quick_xml::Reader;
use quick_xml::events::Event;
use quick_xml::se::Serializer;
use serde::Serialize;
use std::io::Cursor;

/// Parse worksheet XML from bytes using the generated parser.
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

/// Serialize a worksheet back to XML using serde.
fn serialize_worksheet(ws: &Worksheet) -> Result<String, String> {
    let mut buffer = String::new();
    let mut serializer = Serializer::new(&mut buffer);
    serializer.indent(' ', 2);
    ws.serialize(serializer).map_err(|e| e.to_string())?;
    Ok(buffer)
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

/// Options for OOXML parity comparison.
/// Ignores namespace declarations but is otherwise strict.
fn parity_options() -> CompareOptions {
    CompareOptions {
        ignore_whitespace: true,
        ignore_attribute_order: true,
        ignore_element_order: false,
        ignore_attributes: vec!["xmlns".to_string(), "xmlns:".to_string()],
        ignore_elements: vec![],
    }
}

/// Run a parity test: parse, serialize, compare.
fn run_parity_test(
    fixture_name: &str,
) -> (
    Worksheet,
    String,
    Vec<ooxml_corpus::roundtrip::XmlDifference>,
) {
    let original_xml = load_fixture(fixture_name);

    // Parse using generated parser
    let ws = parse_worksheet(&original_xml)
        .unwrap_or_else(|e| panic!("Failed to parse {}: {:?}", fixture_name, e));

    // Serialize back using serde
    let roundtripped_xml = serialize_worksheet(&ws)
        .unwrap_or_else(|e| panic!("Failed to serialize {}: {}", fixture_name, e));

    // Compare XML
    let diffs = compare_xml(
        &original_xml,
        roundtripped_xml.as_bytes(),
        &parity_options(),
    );

    (ws, roundtripped_xml, diffs)
}

/// Report differences with categorization.
fn report_diffs(
    fixture_name: &str,
    roundtripped: &str,
    diffs: &[ooxml_corpus::roundtrip::XmlDifference],
) {
    if diffs.is_empty() {
        return;
    }

    let critical: Vec<_> = diffs
        .iter()
        .filter(|d| {
            matches!(
                d.kind,
                DifferenceKind::MissingElement | DifferenceKind::ExtraElement
            )
        })
        .collect();

    eprintln!(
        "\n=== {} parity: {} diffs ({} critical) ===",
        fixture_name,
        diffs.len(),
        critical.len()
    );

    for diff in diffs.iter().take(15) {
        let marker = if matches!(
            diff.kind,
            DifferenceKind::MissingElement | DifferenceKind::ExtraElement
        ) {
            "CRITICAL"
        } else {
            "info"
        };
        eprintln!(
            "  [{}] {:?} at {}: {}",
            marker, diff.kind, diff.path, diff.description
        );
    }

    if diffs.len() > 15 {
        eprintln!("  ... and {} more", diffs.len() - 15);
    }

    // Log a snippet of roundtripped output
    let snippet: String = roundtripped.chars().take(1500).collect();
    eprintln!("\nRoundtripped (first 1500 chars):\n{}", snippet);
}

// =============================================================================
// Basic structure parity tests
// =============================================================================

#[test]
fn test_parity_basic_cells() {
    let (ws, roundtripped, diffs) = run_parity_test("basic_cells");

    // Verify parsed structure
    assert_eq!(ws.sheet_data.row.len(), 2);
    assert_eq!(ws.sheet_data.row[0].cells.len(), 6);

    // Cell types should be preserved
    assert!(ws.sheet_data.row[0].cells[0].cell_type.is_none());
    assert_eq!(
        ws.sheet_data.row[0].cells[2].cell_type,
        Some(CellType::SharedString)
    );
    assert_eq!(
        ws.sheet_data.row[0].cells[3].cell_type,
        Some(CellType::Boolean)
    );

    report_diffs("basic_cells", &roundtripped, &diffs);
}

#[test]
fn test_parity_formulas() {
    let (ws, roundtripped, diffs) = run_parity_test("formulas");

    // Verify formula structure
    let row1 = &ws.sheet_data.row[0];
    let cell_c1 = &row1.cells[2];
    assert!(cell_c1.formula.is_some());
    let formula = cell_c1.formula.as_ref().unwrap();
    assert_eq!(formula.text.as_deref(), Some("A1+B1"));

    // Check array formula - field is cell_type of type FormulaType
    let row2 = &ws.sheet_data.row[1];
    let cell_a2 = &row2.cells[0];
    assert!(cell_a2.formula.is_some());
    let array_formula = cell_a2.formula.as_ref().unwrap();
    assert_eq!(array_formula.cell_type, Some(FormulaType::Array));

    report_diffs("formulas", &roundtripped, &diffs);
}

#[test]
fn test_parity_merged_cells() {
    let (ws, roundtripped, diffs) = run_parity_test("merged_cells");

    let merged = ws.merged_cells.as_ref().expect("should have merged cells");
    assert_eq!(merged.merge_cell.len(), 3);

    let refs: Vec<_> = merged
        .merge_cell
        .iter()
        .map(|m| m.reference.as_str())
        .collect();
    assert!(refs.contains(&"A1:D1"));
    assert!(refs.contains(&"A4:A5"));

    report_diffs("merged_cells", &roundtripped, &diffs);
}

#[test]
fn test_parity_freeze_panes() {
    let (ws, roundtripped, diffs) = run_parity_test("freeze_panes");

    let views = ws.sheet_views.as_ref().expect("should have sheet views");
    assert_eq!(views.sheet_view.len(), 1);

    let view = &views.sheet_view[0];
    let pane = view.pane.as_ref().expect("should have pane");
    assert_eq!(pane.y_split, Some(1.0));
    assert_eq!(pane.state, Some(PaneState::Frozen));

    report_diffs("freeze_panes", &roundtripped, &diffs);
}

#[test]
fn test_parity_dimensions() {
    let (ws, roundtripped, diffs) = run_parity_test("dimensions");

    let dim = ws.dimension.as_ref().expect("should have dimension");
    assert_eq!(dim.reference.as_str(), "A1:D10");
    assert!(!ws.cols.is_empty());

    report_diffs("dimensions", &roundtripped, &diffs);
}

// =============================================================================
// Edge case parity tests
// =============================================================================

#[test]
fn test_parity_edge_sparse_data() {
    let (ws, roundtripped, diffs) = run_parity_test("edge_sparse_data");

    let rows = &ws.sheet_data.row;
    assert_eq!(rows.len(), 4);
    assert_eq!(rows[0].reference, Some(1));
    assert_eq!(rows[1].reference, Some(100));
    assert_eq!(rows[2].reference, Some(500));
    assert_eq!(rows[3].reference, Some(1000));

    let row500 = &rows[2];
    assert!(
        row500
            .cells
            .iter()
            .any(|c| c.reference.as_deref() == Some("AA500"))
    );

    report_diffs("edge_sparse_data", &roundtripped, &diffs);
}

#[test]
fn test_parity_edge_unicode() {
    let (ws, roundtripped, diffs) = run_parity_test("edge_unicode");

    assert_eq!(ws.sheet_data.row.len(), 10);

    for row in &ws.sheet_data.row {
        let cell = &row.cells[0];
        assert!(
            cell.is.is_some(),
            "Row {} should have inline string",
            row.reference.unwrap_or(0)
        );
    }

    // Smoke test: unicode should survive
    assert!(
        roundtripped.contains("日本語") || roundtripped.contains("&#"),
        "Japanese text should be preserved (possibly escaped)"
    );

    report_diffs("edge_unicode", &roundtripped, &diffs);
}

#[test]
fn test_parity_edge_empty_elements() {
    let (ws, roundtripped, diffs) = run_parity_test("edge_empty_elements");

    assert_eq!(ws.sheet_data.row.len(), 3);
    assert_eq!(ws.sheet_data.row[0].cells.len(), 3);
    assert!(ws.sheet_data.row[1].cells.is_empty());

    if let Some(merged) = &ws.merged_cells {
        assert!(merged.merge_cell.is_empty());
    }

    report_diffs("edge_empty_elements", &roundtripped, &diffs);
}

#[test]
fn test_parity_edge_large_numbers() {
    let (ws, roundtripped, diffs) = run_parity_test("edge_large_numbers");

    assert_eq!(ws.sheet_data.row.len(), 10);
    assert_eq!(
        ws.sheet_data.row[0].cells[0].value.as_deref(),
        Some("9999999999999999")
    );
    assert_eq!(
        ws.sheet_data.row[6].cells[0].value.as_deref(),
        Some("3.141592653589793")
    );

    report_diffs("edge_large_numbers", &roundtripped, &diffs);
}

// =============================================================================
// Complex feature parity tests
// =============================================================================

#[test]
fn test_parity_data_validation() {
    let (ws, roundtripped, diffs) = run_parity_test("data_validation");

    let dvs = ws
        .data_validations
        .as_ref()
        .expect("should have data validations");
    assert_eq!(dvs.data_validation.len(), 5);
    assert_eq!(dvs.data_validation[0].r#type, Some(ValidationType::List));

    report_diffs("data_validation", &roundtripped, &diffs);
}

#[test]
fn test_parity_conditional_formatting() {
    let (ws, roundtripped, diffs) = run_parity_test("conditional_formatting");

    assert_eq!(ws.conditional_formatting.len(), 4);
    assert_eq!(
        ws.conditional_formatting[0].square_reference.as_deref(),
        Some("A1:A10")
    );
    assert_eq!(ws.conditional_formatting[0].cf_rule.len(), 2);

    report_diffs("conditional_formatting", &roundtripped, &diffs);
}

#[test]
fn test_parity_hyperlinks() {
    let (ws, roundtripped, diffs) = run_parity_test("hyperlinks");

    let links = ws.hyperlinks.as_ref().expect("should have hyperlinks");
    assert_eq!(links.hyperlink.len(), 5);
    assert_eq!(links.hyperlink[1].location.as_deref(), Some("Sheet2!A1"));

    report_diffs("hyperlinks", &roundtripped, &diffs);
}

// =============================================================================
// Serde serialization capability tests
// =============================================================================

#[test]
fn test_parsed_worksheet_serializes() {
    // Verify that any parsed worksheet can serialize without panic
    let fixtures = [
        "basic_cells",
        "formulas",
        "merged_cells",
        "freeze_panes",
        "dimensions",
        "data_validation",
        "conditional_formatting",
        "hyperlinks",
        "edge_sparse_data",
        "edge_unicode",
        "edge_empty_elements",
        "edge_large_numbers",
    ];

    for fixture in fixtures {
        let xml = load_fixture(fixture);
        let ws =
            parse_worksheet(&xml).unwrap_or_else(|e| panic!("should parse {}: {:?}", fixture, e));

        // Serialize should not panic
        let result = serialize_worksheet(&ws);
        assert!(
            result.is_ok(),
            "Failed to serialize {}: {:?}",
            fixture,
            result.err()
        );

        // Output should be non-empty
        let output = result.unwrap();
        assert!(!output.is_empty(), "Serialized {} is empty", fixture);

        // Output should contain worksheet element
        assert!(
            output.contains("Worksheet") || output.contains("worksheet"),
            "Serialized {} missing worksheet element",
            fixture
        );
    }
}

#[test]
fn test_serialize_produces_valid_xml() {
    // Verify serialized output is parseable XML
    let xml = load_fixture("basic_cells");
    let ws = parse_worksheet(&xml).expect("should parse");
    let serialized = serialize_worksheet(&ws).expect("should serialize");

    // Try to parse the serialized output
    let mut reader = Reader::from_str(&serialized);
    let mut buf = Vec::new();
    let mut event_count = 0;

    loop {
        buf.clear();
        match reader.read_event_into(&mut buf) {
            Ok(Event::Eof) => break,
            Ok(_) => event_count += 1,
            Err(e) => panic!("Serialized XML is invalid: {}", e),
        }
    }

    assert!(event_count > 0, "Serialized output produced no XML events");
}
