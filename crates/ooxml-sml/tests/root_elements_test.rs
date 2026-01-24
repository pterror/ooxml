//! Tests for root element types beyond Worksheet.
//!
//! Verifies that Workbook, SharedStrings, Stylesheet, and Comments
//! can be parsed and serialized correctly.

#![cfg(feature = "full")]

use ooxml_sml::parsers::{FromXml, ParseError};
use ooxml_sml::types::*;
use quick_xml::Reader;
use quick_xml::events::Event;
use quick_xml::se::Serializer;
use serde::Serialize;
use std::io::Cursor;

/// Load a fixture file.
fn load_fixture(name: &str) -> Vec<u8> {
    let path = format!(
        "{}/tests/fixtures/xml/{}.xml",
        env!("CARGO_MANIFEST_DIR"),
        name
    );
    std::fs::read(&path).unwrap_or_else(|e| panic!("Failed to load fixture {}: {}", name, e))
}

/// Generic parser for any root element type.
fn parse_root<T: FromXml>(xml: &[u8], element_name: &[u8]) -> Result<T, ParseError> {
    let mut reader = Reader::from_reader(Cursor::new(xml));
    let mut buf = Vec::new();

    loop {
        buf.clear();
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) if e.name().as_ref() == element_name => {
                return T::from_xml(&mut reader, &e, false);
            }
            Ok(Event::Empty(e)) if e.name().as_ref() == element_name => {
                return T::from_xml(&mut reader, &e, true);
            }
            Ok(Event::Eof) => {
                return Err(ParseError::UnexpectedElement(format!(
                    "EOF before {}",
                    String::from_utf8_lossy(element_name)
                )));
            }
            Err(e) => return Err(ParseError::Xml(e)),
            _ => {}
        }
    }
}

/// Serialize any type to XML.
fn serialize<T: Serialize>(value: &T) -> Result<String, String> {
    let mut buffer = String::new();
    let mut serializer = Serializer::new(&mut buffer);
    serializer.indent(' ', 2);
    value.serialize(serializer).map_err(|e| e.to_string())?;
    Ok(buffer)
}

// =============================================================================
// Workbook tests
// =============================================================================

#[test]
fn test_workbook_parse() {
    let xml = load_fixture("workbook");
    let wb: Workbook = parse_root(&xml, b"workbook").expect("should parse workbook");

    // Verify sheets (Box<Sheets>, not Option)
    assert_eq!(wb.sheets.sheet.len(), 3);
    assert_eq!(wb.sheets.sheet[0].name.as_str(), "Sheet1");
    assert_eq!(wb.sheets.sheet[2].state, Some(SheetState::Hidden));

    // Verify defined names (Option<Box<DefinedNames>>)
    assert!(wb.defined_names.is_some());
    let names = wb.defined_names.as_ref().unwrap();
    assert_eq!(names.defined_name.len(), 2);
}

#[test]
fn test_workbook_serialize() {
    let xml = load_fixture("workbook");
    let wb: Workbook = parse_root(&xml, b"workbook").expect("should parse");

    let serialized = serialize(&wb).expect("should serialize");

    // Should have correct root element
    assert!(
        serialized.starts_with("<workbook"),
        "Should start with <workbook, got: {}",
        &serialized[..50.min(serialized.len())]
    );

    // Should contain key elements
    assert!(
        serialized.contains("<sheets>") || serialized.contains("<sheets "),
        "Missing sheets element"
    );
    assert!(serialized.contains("<sheet "), "Missing sheet elements");
    assert!(serialized.contains("Sheet1"), "Missing sheet name");
}

#[test]
fn test_workbook_roundtrip() {
    let xml = load_fixture("workbook");
    let wb1: Workbook = parse_root(&xml, b"workbook").expect("should parse");

    // Serialize
    let serialized = serialize(&wb1).expect("should serialize");

    // Parse the serialized output
    let wb2: Workbook =
        parse_root(serialized.as_bytes(), b"workbook").expect("should parse serialized output");

    // Compare key fields
    assert_eq!(wb1.sheets.sheet.len(), wb2.sheets.sheet.len());
    for (sheet1, sheet2) in wb1.sheets.sheet.iter().zip(wb2.sheets.sheet.iter()) {
        assert_eq!(sheet1.name, sheet2.name);
        assert_eq!(sheet1.sheet_id, sheet2.sheet_id);
    }
}

// =============================================================================
// SharedStrings tests
// =============================================================================

#[test]
fn test_shared_strings_parse() {
    let xml = load_fixture("shared_strings");
    let sst: SharedStrings = parse_root(&xml, b"sst").expect("should parse shared strings");

    assert_eq!(sst.count, Some(10));
    assert_eq!(sst.unique_count, Some(7));
    assert_eq!(sst.si.len(), 7); // field is `si`, not `string_item`
}

#[test]
fn test_shared_strings_serialize() {
    let xml = load_fixture("shared_strings");
    let sst: SharedStrings = parse_root(&xml, b"sst").expect("should parse");

    let serialized = serialize(&sst).expect("should serialize");

    // Should have correct root element
    assert!(serialized.starts_with("<sst"), "Should start with <sst");

    // Should contain string items
    assert!(
        serialized.contains("<si>") || serialized.contains("<si ") || serialized.contains("<si/>"),
        "Missing si elements"
    );
}

#[test]
fn test_shared_strings_roundtrip() {
    let xml = load_fixture("shared_strings");
    let sst1: SharedStrings = parse_root(&xml, b"sst").expect("should parse");

    let serialized = serialize(&sst1).expect("should serialize");
    let sst2: SharedStrings =
        parse_root(serialized.as_bytes(), b"sst").expect("should parse serialized output");

    assert_eq!(sst1.si.len(), sst2.si.len());
}

// =============================================================================
// Stylesheet tests
// =============================================================================

#[test]
fn test_stylesheet_parse() {
    let xml = load_fixture("styles");
    let styles: Stylesheet = parse_root(&xml, b"styleSheet").expect("should parse stylesheet");

    // Verify fonts
    assert!(styles.fonts.is_some(), "fonts should be parsed");
    let fonts = styles.fonts.as_ref().unwrap();
    assert_eq!(fonts.font.len(), 3);

    // Verify fills
    assert!(styles.fills.is_some(), "fills should be parsed");
    let fills = styles.fills.as_ref().unwrap();
    assert_eq!(fills.fill.len(), 2);

    // Verify borders
    assert!(styles.borders.is_some(), "borders should be parsed");
    let borders = styles.borders.as_ref().unwrap();
    assert_eq!(borders.border.len(), 2);

    // Verify cell formats
    assert!(styles.cell_xfs.is_some(), "cell_xfs should be parsed");
    let xfs = styles.cell_xfs.as_ref().unwrap();
    assert_eq!(xfs.xf.len(), 3);

    // Verify cell style formats
    assert!(
        styles.cell_style_xfs.is_some(),
        "cell_style_xfs should be parsed"
    );
    let style_xfs = styles.cell_style_xfs.as_ref().unwrap();
    assert_eq!(style_xfs.xf.len(), 1);

    // Verify cell styles
    assert!(styles.cell_styles.is_some(), "cell_styles should be parsed");
    let cell_styles = styles.cell_styles.as_ref().unwrap();
    assert_eq!(cell_styles.cell_style.len(), 1);
}

#[test]
fn test_stylesheet_serialize() {
    let xml = load_fixture("styles");
    let styles: Stylesheet = parse_root(&xml, b"styleSheet").expect("should parse");

    let serialized = serialize(&styles).expect("should serialize");

    // Should have correct root element
    assert!(
        serialized.starts_with("<styleSheet"),
        "Should start with <styleSheet"
    );

    // Should contain key elements
    assert!(
        serialized.contains("<fonts") || serialized.contains("<font"),
        "Missing fonts"
    );
    assert!(
        serialized.contains("<fills") || serialized.contains("<fill"),
        "Missing fills"
    );
}

#[test]
fn test_stylesheet_roundtrip() {
    let xml = load_fixture("styles");
    let s1: Stylesheet = parse_root(&xml, b"styleSheet").expect("should parse");

    let serialized = serialize(&s1).expect("should serialize");
    let s2: Stylesheet =
        parse_root(serialized.as_bytes(), b"styleSheet").expect("should parse serialized output");

    // Compare counts
    if let (Some(f1), Some(f2)) = (&s1.fonts, &s2.fonts) {
        assert_eq!(f1.font.len(), f2.font.len());
    }
}

// =============================================================================
// Comments tests
// =============================================================================

#[test]
fn test_comments_parse() {
    let xml = load_fixture("comments");
    let comments: Comments = parse_root(&xml, b"comments").expect("should parse comments");

    // Verify authors (Box<Authors>, not Option)
    assert_eq!(comments.authors.author.len(), 2);
    assert_eq!(comments.authors.author[0].as_str(), "John Doe");

    // Verify comment list (Box<CommentList>, not Option)
    assert_eq!(comments.comment_list.comment.len(), 3);
    assert_eq!(comments.comment_list.comment[0].reference.as_str(), "A1");
}

#[test]
fn test_comments_serialize() {
    let xml = load_fixture("comments");
    let comments: Comments = parse_root(&xml, b"comments").expect("should parse");

    let serialized = serialize(&comments).expect("should serialize");

    // Should have correct root element
    assert!(
        serialized.starts_with("<comments"),
        "Should start with <comments"
    );

    // Should contain key elements
    assert!(
        serialized.contains("<authors>")
            || serialized.contains("<authors ")
            || serialized.contains("<author>"),
        "Missing authors"
    );
    assert!(
        serialized.contains("<commentList>")
            || serialized.contains("<commentList ")
            || serialized.contains("<comment "),
        "Missing comments"
    );
}

#[test]
fn test_comments_roundtrip() {
    let xml = load_fixture("comments");
    let c1: Comments = parse_root(&xml, b"comments").expect("should parse");

    let serialized = serialize(&c1).expect("should serialize");
    let c2: Comments =
        parse_root(serialized.as_bytes(), b"comments").expect("should parse serialized output");

    assert_eq!(c1.authors.author.len(), c2.authors.author.len());
    assert_eq!(c1.comment_list.comment.len(), c2.comment_list.comment.len());
}

// =============================================================================
// Element mapping coverage tests
// =============================================================================

#[test]
fn test_element_names_in_serialized_output() {
    // Parse worksheet and verify element names in serialized output
    let xml = load_fixture("basic_cells");
    let ws: Worksheet = parse_root(&xml, b"worksheet").expect("should parse");
    let serialized = serialize(&ws).expect("should serialize");

    // Verify key element names are lowercase as per OOXML spec
    let expected_elements = ["<worksheet", "<sheetData", "<row ", "<c ", "<v>"];

    for elem in expected_elements {
        assert!(
            serialized.contains(elem),
            "Missing expected element '{}' in serialized output.\nOutput: {}",
            elem,
            &serialized[..500.min(serialized.len())]
        );
    }

    // Verify NO PascalCase element names
    let unexpected = ["<Worksheet", "<SheetData", "<Row ", "<Cell "];
    for elem in unexpected {
        assert!(
            !serialized.contains(elem),
            "Found unexpected PascalCase element '{}' - mapping not applied",
            elem
        );
    }
}

#[test]
fn test_nested_element_names() {
    // Test that nested elements also have correct names
    let xml = load_fixture("freeze_panes");
    let ws: Worksheet = parse_root(&xml, b"worksheet").expect("should parse");
    let serialized = serialize(&ws).expect("should serialize");

    // sheetViews and nested elements
    assert!(serialized.contains("<sheetViews"), "Missing sheetViews");
    assert!(serialized.contains("<sheetView"), "Missing sheetView");
    assert!(serialized.contains("<pane"), "Missing pane");
}

#[test]
fn test_formula_element_name() {
    let xml = load_fixture("formulas");
    let ws: Worksheet = parse_root(&xml, b"worksheet").expect("should parse");
    let serialized = serialize(&ws).expect("should serialize");

    // Formula should be <f> not <CellFormula>
    assert!(
        serialized.contains("<f>") || serialized.contains("<f "),
        "Missing <f> element"
    );
    assert!(
        !serialized.contains("<CellFormula"),
        "Found <CellFormula> instead of <f>"
    );
}

#[test]
fn test_merged_cells_element_names() {
    let xml = load_fixture("merged_cells");
    let ws: Worksheet = parse_root(&xml, b"worksheet").expect("should parse");
    let serialized = serialize(&ws).expect("should serialize");

    assert!(serialized.contains("<mergeCells"), "Missing mergeCells");
    assert!(serialized.contains("<mergeCell "), "Missing mergeCell");
    assert!(
        !serialized.contains("<MergedCells"),
        "Found PascalCase MergedCells"
    );
}

// =============================================================================
// Round-trip integrity tests
// =============================================================================

#[test]
fn test_worksheet_roundtrip_integrity() {
    // Full round-trip: parse → serialize → parse → compare
    let xml = load_fixture("basic_cells");
    let ws1: Worksheet = parse_root(&xml, b"worksheet").expect("should parse");

    let serialized = serialize(&ws1).expect("should serialize");
    let ws2: Worksheet =
        parse_root(serialized.as_bytes(), b"worksheet").expect("should parse serialized output");

    // Compare data integrity
    assert_eq!(ws1.sheet_data.row.len(), ws2.sheet_data.row.len());
    for (r1, r2) in ws1.sheet_data.row.iter().zip(ws2.sheet_data.row.iter()) {
        assert_eq!(r1.reference, r2.reference);
        assert_eq!(r1.cells.len(), r2.cells.len());
        for (c1, c2) in r1.cells.iter().zip(r2.cells.iter()) {
            assert_eq!(c1.reference, c2.reference);
            assert_eq!(c1.value, c2.value);
            assert_eq!(c1.cell_type, c2.cell_type);
        }
    }
}

#[test]
fn test_all_fixtures_roundtrip() {
    // Verify all worksheet fixtures can roundtrip without data loss
    let fixtures = [
        "basic_cells",
        "formulas",
        "merged_cells",
        "freeze_panes",
        "freeze_panes_both",
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
        let ws1: Worksheet = parse_root(&xml, b"worksheet")
            .unwrap_or_else(|e| panic!("Failed to parse {}: {:?}", fixture, e));

        let serialized =
            serialize(&ws1).unwrap_or_else(|e| panic!("Failed to serialize {}: {}", fixture, e));

        let ws2: Worksheet = parse_root(serialized.as_bytes(), b"worksheet")
            .unwrap_or_else(|e| panic!("Failed to parse roundtripped {}: {:?}", fixture, e));

        // Basic integrity check
        assert_eq!(
            ws1.sheet_data.row.len(),
            ws2.sheet_data.row.len(),
            "Row count mismatch in {}",
            fixture
        );
    }
}
