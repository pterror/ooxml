//! Parity tests for worksheet parsing: generated vs hand-written.
//!
//! These tests verify that the generated `parse_worksheet()` function
//! produces equivalent results to the hand-written event parser.

// These tests require the full feature set
#![cfg(feature = "full")]

use ooxml_sml::ext::{
    CellExt, CellResolveExt, ResolveContext, RowExt, WorksheetExt, parse_worksheet,
};

/// Sample worksheet XML for testing
const SAMPLE_WORKSHEET: &[u8] = br#"<?xml version="1.0" encoding="UTF-8"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
    <dimension ref="A1:D5"/>
    <sheetViews>
        <sheetView tabSelected="1" workbookViewId="0">
            <pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/>
        </sheetView>
    </sheetViews>
    <cols>
        <col min="1" max="1" width="15" customWidth="1"/>
        <col min="2" max="2" width="20" customWidth="1"/>
    </cols>
    <sheetData>
        <row r="1" spans="1:4">
            <c r="A1" t="s"><v>0</v></c>
            <c r="B1" t="s"><v>1</v></c>
            <c r="C1" t="s"><v>2</v></c>
            <c r="D1" t="s"><v>3</v></c>
        </row>
        <row r="2" spans="1:4">
            <c r="A2" t="s"><v>4</v></c>
            <c r="B2"><v>100</v></c>
            <c r="C2"><v>200.5</v></c>
            <c r="D2" t="b"><v>1</v></c>
        </row>
        <row r="3" spans="1:4">
            <c r="A3" t="s"><v>5</v></c>
            <c r="B3"><v>150</v></c>
            <c r="C3"><v>300.75</v></c>
            <c r="D3" t="b"><v>0</v></c>
        </row>
        <row r="4" spans="1:4">
            <c r="A4" t="s"><v>6</v></c>
            <c r="B4"><f>SUM(B2:B3)</f><v>250</v></c>
            <c r="C4"><f>SUM(C2:C3)</f><v>501.25</v></c>
            <c r="D4" t="e"><v>#N/A</v></c>
        </row>
    </sheetData>
    <mergeCells count="1">
        <mergeCell ref="A5:D5"/>
    </mergeCells>
    <conditionalFormatting sqref="B2:C3">
        <cfRule type="cellIs" dxfId="0" priority="1" operator="greaterThan">
            <formula>100</formula>
        </cfRule>
    </conditionalFormatting>
    <dataValidations count="1">
        <dataValidation type="whole" allowBlank="1" showInputMessage="1" sqref="B2:B3">
            <formula1>0</formula1>
            <formula2>1000</formula2>
        </dataValidation>
    </dataValidations>
    <autoFilter ref="A1:D4">
        <filterColumn colId="0"/>
    </autoFilter>
</worksheet>"#;

/// Shared strings table for testing
fn test_shared_strings() -> Vec<String> {
    vec![
        "Name".to_string(),   // 0
        "Value1".to_string(), // 1
        "Value2".to_string(), // 2
        "Active".to_string(), // 3
        "Alice".to_string(),  // 4
        "Bob".to_string(),    // 5
        "Total".to_string(),  // 6
    ]
}

#[test]
fn test_generated_worksheet_parsing() {
    let worksheet = parse_worksheet(SAMPLE_WORKSHEET).expect("parse failed");

    // Basic structure
    assert_eq!(worksheet.row_count(), 4);
    assert!(!worksheet.is_empty());

    // Verify features detected
    assert!(worksheet.has_auto_filter());
    assert!(worksheet.has_merged_cells());
    assert!(worksheet.has_conditional_formatting());
    assert!(worksheet.has_data_validations());
    assert!(worksheet.has_freeze_panes());
}

#[test]
fn test_generated_worksheet_row_access() {
    let worksheet = parse_worksheet(SAMPLE_WORKSHEET).expect("parse failed");

    // Row 1 - headers (shared strings)
    let row1 = worksheet.row(1).expect("row 1 should exist");
    assert_eq!(row1.cell_count(), 4);

    // Row 2 - data with mixed types
    let row2 = worksheet.row(2).expect("row 2 should exist");
    assert_eq!(row2.cell_count(), 4);

    // Row 4 - with formulas and error
    let row4 = worksheet.row(4).expect("row 4 should exist");
    assert_eq!(row4.cell_count(), 4);

    // Non-existent row
    assert!(worksheet.row(10).is_none());
}

#[test]
fn test_generated_worksheet_cell_access() {
    let worksheet = parse_worksheet(SAMPLE_WORKSHEET).expect("parse failed");

    // Cell A1 - shared string
    let a1 = worksheet.cell("A1").expect("A1 should exist");
    assert!(a1.is_shared_string());
    assert_eq!(a1.raw_value(), Some("0"));
    assert_eq!(a1.column_number(), Some(1));
    assert_eq!(a1.row_number(), Some(1));

    // Cell B2 - number
    let b2 = worksheet.cell("B2").expect("B2 should exist");
    assert!(b2.is_number());
    assert_eq!(b2.raw_value(), Some("100"));

    // Cell C2 - decimal number
    let c2 = worksheet.cell("C2").expect("C2 should exist");
    assert_eq!(c2.raw_value(), Some("200.5"));

    // Cell D2 - boolean
    let d2 = worksheet.cell("D2").expect("D2 should exist");
    assert!(d2.is_boolean());
    assert_eq!(d2.raw_value(), Some("1"));

    // Cell B4 - formula
    let b4 = worksheet.cell("B4").expect("B4 should exist");
    assert!(b4.has_formula());
    assert_eq!(b4.raw_value(), Some("250"));

    // Cell D4 - error
    let d4 = worksheet.cell("D4").expect("D4 should exist");
    assert!(d4.is_error());
    assert_eq!(d4.raw_value(), Some("#N/A"));

    // Non-existent cell
    assert!(worksheet.cell("Z99").is_none());
}

#[test]
fn test_generated_worksheet_value_resolution() {
    let worksheet = parse_worksheet(SAMPLE_WORKSHEET).expect("parse failed");
    let ctx = ResolveContext::new(test_shared_strings());

    // A1 - shared string "Name"
    let a1 = worksheet.cell("A1").expect("A1");
    assert_eq!(a1.value_as_string(&ctx), "Name");

    // A2 - shared string "Alice"
    let a2 = worksheet.cell("A2").expect("A2");
    assert_eq!(a2.value_as_string(&ctx), "Alice");

    // B2 - number 100
    let b2 = worksheet.cell("B2").expect("B2");
    assert_eq!(b2.value_as_number(&ctx), Some(100.0));

    // C2 - decimal 200.5
    let c2 = worksheet.cell("C2").expect("C2");
    assert_eq!(c2.value_as_number(&ctx), Some(200.5));

    // D2 - boolean true
    let d2 = worksheet.cell("D2").expect("D2");
    assert_eq!(d2.value_as_bool(&ctx), Some(true));

    // D3 - boolean false
    let d3 = worksheet.cell("D3").expect("D3");
    assert_eq!(d3.value_as_bool(&ctx), Some(false));

    // D4 - error
    let d4 = worksheet.cell("D4").expect("D4");
    assert_eq!(d4.value_as_string(&ctx), "#N/A");
}

#[test]
fn test_generated_worksheet_iteration() {
    let worksheet = parse_worksheet(SAMPLE_WORKSHEET).expect("parse failed");
    let ctx = ResolveContext::new(test_shared_strings());

    // Iterate all rows and count cells
    let mut total_cells = 0;
    for row in worksheet.rows() {
        total_cells += row.cell_count();
    }
    assert_eq!(total_cells, 16); // 4 rows x 4 cells

    // Iterate row 1 and collect resolved values
    let row1 = worksheet.row(1).expect("row 1");
    let headers: Vec<String> = row1.cells_iter().map(|c| c.value_as_string(&ctx)).collect();
    assert_eq!(headers, vec!["Name", "Value1", "Value2", "Active"]);
}

#[test]
fn test_generated_worksheet_auto_filter() {
    let worksheet = parse_worksheet(SAMPLE_WORKSHEET).expect("parse failed");

    assert!(worksheet.has_auto_filter());

    let auto_filter = worksheet.auto_filter.as_ref().expect("auto_filter");
    assert_eq!(auto_filter.reference.as_deref(), Some("A1:D4"));
    assert_eq!(auto_filter.filter_column.len(), 1);
    assert_eq!(auto_filter.filter_column[0].column_id, 0);
}

#[test]
fn test_generated_worksheet_merged_cells() {
    let worksheet = parse_worksheet(SAMPLE_WORKSHEET).expect("parse failed");

    assert!(worksheet.has_merged_cells());

    let merged = worksheet.merged_cells.as_ref().expect("merged_cells");
    assert_eq!(merged.merge_cell.len(), 1);
    assert_eq!(merged.merge_cell[0].reference.as_str(), "A5:D5");
}

#[test]
fn test_generated_worksheet_conditional_formatting() {
    let worksheet = parse_worksheet(SAMPLE_WORKSHEET).expect("parse failed");

    assert!(worksheet.has_conditional_formatting());
    assert_eq!(worksheet.conditional_formatting.len(), 1);

    let cf = &worksheet.conditional_formatting[0];
    assert!(
        cf.square_reference
            .as_ref()
            .is_some_and(|s| s.contains("B2:C3"))
    );
}

#[test]
fn test_generated_worksheet_data_validation() {
    let worksheet = parse_worksheet(SAMPLE_WORKSHEET).expect("parse failed");

    assert!(worksheet.has_data_validations());

    let dvs = worksheet
        .data_validations
        .as_ref()
        .expect("data_validations");
    assert_eq!(dvs.data_validation.len(), 1);
}

#[test]
fn test_generated_worksheet_freeze_panes() {
    let worksheet = parse_worksheet(SAMPLE_WORKSHEET).expect("parse failed");

    assert!(worksheet.has_freeze_panes());

    let views = worksheet.sheet_views.as_ref().expect("sheet_views");
    let view = &views.sheet_view[0];
    let pane = view.pane.as_ref().expect("pane");

    assert_eq!(pane.y_split, Some(1.0));
    assert_eq!(pane.top_left_cell.as_deref(), Some("A2"));
}
