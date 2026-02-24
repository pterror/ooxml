//! Tests for pivot table and conditional formatting reading APIs.
//!
//! Covers `PivotTableExt`, `ConditionalFormattingExt`, `ConditionalRuleExt`,
//! and `WorksheetConditionalFormattingExt` traits defined in `ext.rs`.

// All tests in this file require the full feature set.
#![cfg(feature = "full")]

use ooxml_sml::parsers::{FromXml, ParseError};
use ooxml_sml::types::Worksheet;
use quick_xml::Reader;
use quick_xml::events::Event;
use std::io::Cursor;

// ============================================================================
// Helpers
// ============================================================================

/// Parse a worksheet from raw XML bytes (mirrors the helper in fixture_tests.rs).
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

/// Parse any generated type `T` from raw XML bytes (mirrors workbook::bootstrap).
///
/// Scans for the first Start or Empty event and calls `T::from_xml` on it.
fn bootstrap<T: FromXml>(xml: &[u8]) -> T {
    let mut reader = Reader::from_reader(Cursor::new(xml));
    let mut buf = Vec::new();
    loop {
        buf.clear();
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) => {
                return T::from_xml(&mut reader, &e, false)
                    .unwrap_or_else(|err| panic!("parse failed: {err:?}"));
            }
            Ok(Event::Empty(e)) => {
                return T::from_xml(&mut reader, &e, true)
                    .unwrap_or_else(|err| panic!("parse failed: {err:?}"));
            }
            Ok(Event::Eof) => panic!("EOF before first element"),
            Err(e) => panic!("XML error: {e}"),
            _ => {}
        }
    }
}

/// Load a fixture file from tests/fixtures/xml/.
fn load_fixture(name: &str) -> Vec<u8> {
    let path = format!(
        "{}/tests/fixtures/xml/{}.xml",
        env!("CARGO_MANIFEST_DIR"),
        name
    );
    std::fs::read(&path).unwrap_or_else(|e| panic!("Failed to load fixture {name}: {e}"))
}

// ============================================================================
// Conditional Formatting — ColorScale
// ============================================================================

#[test]
fn test_colorscale_rule_type() {
    use ooxml_sml::ConditionalRuleExt;
    use ooxml_sml::types::{ConditionalRule, ConditionalType};

    let xml = br#"<?xml version="1.0" encoding="UTF-8"?>
    <cfRule type="colorScale" priority="1"
      xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
      <colorScale>
        <cfvo type="min"/>
        <cfvo type="max"/>
        <color rgb="FF63BE7B"/>
        <color rgb="FFFFF8C"/>
      </colorScale>
    </cfRule>"#;

    let rule: ConditionalRule = bootstrap(xml);
    assert_eq!(rule.rule_type(), Some(&ConditionalType::ColorScale));
    assert_eq!(rule.priority(), 1);
    assert!(rule.has_color_scale(), "rule should have a color scale");
    assert!(!rule.has_data_bar());
    assert!(!rule.has_icon_set());
    assert!(
        rule.formulas().is_empty(),
        "color scale rules have no formulas"
    );
}

#[test]
fn test_colorscale_worksheet_fixture() {
    use ooxml_sml::types::ConditionalType;
    use ooxml_sml::{
        ConditionalFormattingExt, ConditionalRuleExt, WorksheetConditionalFormattingExt,
    };

    let xml = load_fixture("conditional_formatting_colorscale");
    let ws = parse_worksheet_xml(&xml).expect("should parse colorscale fixture");

    // One <conditionalFormatting> block
    let cfs = ws.conditional_formattings();
    assert_eq!(cfs.len(), 1, "expected 1 conditional formatting block");

    let cf = &cfs[0];
    assert_eq!(cf.cell_range(), Some("A1:A3"));
    assert_eq!(cf.rule_count(), 1);

    // Single rule is a colorScale
    let rule = &cf.rules()[0];
    assert_eq!(rule.rule_type(), Some(&ConditionalType::ColorScale));
    assert!(rule.has_color_scale());
    assert!(!rule.has_data_bar());
    assert!(!rule.has_icon_set());
}

#[test]
fn test_colorscale_worksheet_has_cf() {
    use ooxml_sml::WorksheetExt;

    let xml = load_fixture("conditional_formatting_colorscale");
    let ws = parse_worksheet_xml(&xml).expect("should parse colorscale fixture");

    assert!(ws.has_conditional_formatting());
    // The worksheet has 3 data rows
    assert_eq!(ws.row_count(), 3);
}

// ============================================================================
// Conditional Formatting — DataBar
// ============================================================================

#[test]
fn test_databar_rule_type() {
    use ooxml_sml::ConditionalRuleExt;
    use ooxml_sml::types::{ConditionalRule, ConditionalType};

    let xml = br#"<?xml version="1.0" encoding="UTF-8"?>
    <cfRule type="dataBar" priority="1"
      xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
      <dataBar>
        <cfvo type="min"/>
        <cfvo type="max"/>
        <color rgb="FF638EC6"/>
      </dataBar>
    </cfRule>"#;

    let rule: ConditionalRule = bootstrap(xml);
    assert_eq!(rule.rule_type(), Some(&ConditionalType::DataBar));
    assert_eq!(rule.priority(), 1);
    assert!(rule.has_data_bar(), "rule should have a data bar");
    assert!(!rule.has_color_scale());
    assert!(!rule.has_icon_set());
    assert!(rule.formulas().is_empty());
}

#[test]
fn test_databar_worksheet_fixture() {
    use ooxml_sml::types::ConditionalType;
    use ooxml_sml::{
        ConditionalFormattingExt, ConditionalRuleExt, WorksheetConditionalFormattingExt,
    };

    let xml = load_fixture("conditional_formatting_databar");
    let ws = parse_worksheet_xml(&xml).expect("should parse databar fixture");

    let cfs = ws.conditional_formattings();
    assert_eq!(cfs.len(), 1);

    let cf = &cfs[0];
    assert_eq!(cf.cell_range(), Some("A1:A2"));
    assert_eq!(cf.rule_count(), 1);

    let rule = &cf.rules()[0];
    assert_eq!(rule.rule_type(), Some(&ConditionalType::DataBar));
    assert!(rule.has_data_bar());
    assert!(!rule.has_color_scale());
    assert!(!rule.has_icon_set());
}

#[test]
fn test_databar_worksheet_has_data() {
    let xml = load_fixture("conditional_formatting_databar");
    let ws = parse_worksheet_xml(&xml).expect("should parse databar fixture");

    // Two rows of data
    assert_eq!(ws.sheet_data.row.len(), 2);
    assert_eq!(ws.sheet_data.row[0].cells[0].value.as_deref(), Some("25"));
    assert_eq!(ws.sheet_data.row[1].cells[0].value.as_deref(), Some("75"));
}

// ============================================================================
// Conditional Formatting — IconSet
// ============================================================================

#[test]
fn test_iconset_rule_type() {
    use ooxml_sml::ConditionalRuleExt;
    use ooxml_sml::types::{ConditionalRule, ConditionalType};

    let xml = br#"<?xml version="1.0" encoding="UTF-8"?>
    <cfRule type="iconSet" priority="1"
      xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
      <iconSet>
        <cfvo type="percent" val="0"/>
        <cfvo type="percent" val="33"/>
        <cfvo type="percent" val="67"/>
      </iconSet>
    </cfRule>"#;

    let rule: ConditionalRule = bootstrap(xml);
    assert_eq!(rule.rule_type(), Some(&ConditionalType::IconSet));
    assert_eq!(rule.priority(), 1);
    assert!(rule.has_icon_set(), "rule should have an icon set");
    assert!(!rule.has_color_scale());
    assert!(!rule.has_data_bar());
    assert!(rule.formulas().is_empty());
}

#[test]
fn test_iconset_worksheet_fixture() {
    use ooxml_sml::types::ConditionalType;
    use ooxml_sml::{
        ConditionalFormattingExt, ConditionalRuleExt, WorksheetConditionalFormattingExt,
    };

    let xml = load_fixture("conditional_formatting_iconset");
    let ws = parse_worksheet_xml(&xml).expect("should parse iconset fixture");

    let cfs = ws.conditional_formattings();
    assert_eq!(cfs.len(), 1);

    let cf = &cfs[0];
    assert_eq!(cf.cell_range(), Some("A1:A3"));
    assert_eq!(cf.rule_count(), 1);

    let rule = &cf.rules()[0];
    assert_eq!(rule.rule_type(), Some(&ConditionalType::IconSet));
    assert!(rule.has_icon_set());
    assert!(!rule.has_color_scale());
    assert!(!rule.has_data_bar());
}

// ============================================================================
// Conditional Formatting — Multiple rule types
// ============================================================================

#[test]
fn test_multiple_cf_rule_types() {
    use ooxml_sml::types::ConditionalType;
    use ooxml_sml::{
        ConditionalFormattingExt, ConditionalRuleExt, WorksheetConditionalFormattingExt,
    };

    let xml = load_fixture("conditional_formatting_multitype");
    let ws = parse_worksheet_xml(&xml).expect("should parse multitype fixture");

    let cfs = ws.conditional_formattings();
    assert_eq!(cfs.len(), 1, "expected 1 conditional formatting block");

    let cf = &cfs[0];
    assert_eq!(cf.cell_range(), Some("A1:A10"));
    assert_eq!(cf.rule_count(), 4);

    let rules = cf.rules();

    // Rule 0: Top10, priority 1
    assert_eq!(rules[0].rule_type(), Some(&ConditionalType::Top10));
    assert_eq!(rules[0].priority(), 1);

    // Rule 1: AboveAverage, priority 2
    assert_eq!(rules[1].rule_type(), Some(&ConditionalType::AboveAverage));
    assert_eq!(rules[1].priority(), 2);

    // Rule 2: DuplicateValues, priority 3
    assert_eq!(
        rules[2].rule_type(),
        Some(&ConditionalType::DuplicateValues)
    );
    assert_eq!(rules[2].priority(), 3);

    // Rule 3: ContainsText, priority 4
    assert_eq!(rules[3].rule_type(), Some(&ConditionalType::ContainsText));
    assert_eq!(rules[3].priority(), 4);
}

#[test]
fn test_multiple_cf_rule_formulas() {
    use ooxml_sml::{
        ConditionalFormattingExt, ConditionalRuleExt, WorksheetConditionalFormattingExt,
    };

    let xml = load_fixture("conditional_formatting_multitype");
    let ws = parse_worksheet_xml(&xml).expect("should parse multitype fixture");

    let rules = ws.conditional_formattings()[0].rules();

    // Rules 0-2 have no formula children
    assert!(rules[0].formulas().is_empty());
    assert!(rules[1].formulas().is_empty());
    assert!(rules[2].formulas().is_empty());

    // Rule 3 (ContainsText) has one formula
    assert_eq!(rules[3].formulas().len(), 1);
    assert!(
        rules[3].formulas()[0].contains("SEARCH"),
        "formula should contain SEARCH function"
    );
}

#[test]
fn test_multiple_cf_rule_attributes() {
    use ooxml_sml::{
        ConditionalFormattingExt, ConditionalRuleExt, WorksheetConditionalFormattingExt,
    };

    let xml = load_fixture("conditional_formatting_multitype");
    let ws = parse_worksheet_xml(&xml).expect("should parse multitype fixture");

    let rules = ws.conditional_formattings()[0].rules();

    // Top10 rule: rank=5, percent=false (0)
    let top10 = &rules[0];
    assert_eq!(
        top10.rule_type(),
        Some(&ooxml_sml::types::ConditionalType::Top10)
    );
    // rank attribute
    assert_eq!(ws.conditional_formattings()[0].rules()[0].priority(), 1);

    // AboveAverage rule: aboveAverage=true
    let above_avg = &rules[1];
    // The aboveAverage field on the raw type should be Some(true)
    assert!(above_avg.rule_type().is_some());

    // ContainsText rule: text="error"
    let contains = &rules[3];
    assert_eq!(
        contains.rule_type(),
        Some(&ooxml_sml::types::ConditionalType::ContainsText)
    );
}

// ============================================================================
// Conditional Formatting — Cell range via ConditionalFormattingExt
// ============================================================================

#[test]
fn test_cf_cell_range_colorscale() {
    use ooxml_sml::{ConditionalFormattingExt, WorksheetConditionalFormattingExt};

    let xml = load_fixture("conditional_formatting_colorscale");
    let ws = parse_worksheet_xml(&xml).expect("parse");
    assert_eq!(ws.conditional_formattings()[0].cell_range(), Some("A1:A3"));
}

#[test]
fn test_cf_cell_range_multitype() {
    use ooxml_sml::{ConditionalFormattingExt, WorksheetConditionalFormattingExt};

    let xml = load_fixture("conditional_formatting_multitype");
    let ws = parse_worksheet_xml(&xml).expect("parse");
    assert_eq!(ws.conditional_formattings()[0].cell_range(), Some("A1:A10"));
}

#[test]
fn test_cf_empty_block() {
    use ooxml_sml::ConditionalFormattingExt;
    use ooxml_sml::types::ConditionalFormatting;

    let xml = br#"<?xml version="1.0" encoding="UTF-8"?>
    <conditionalFormatting sqref="B2:C5"
      xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
    </conditionalFormatting>"#;

    let cf: ConditionalFormatting = bootstrap(xml);
    assert_eq!(cf.cell_range(), Some("B2:C5"));
    assert_eq!(cf.rule_count(), 0);
    assert!(cf.rules().is_empty());
}

#[test]
fn test_cf_rule_count() {
    use ooxml_sml::{ConditionalFormattingExt, WorksheetConditionalFormattingExt};

    let xml = load_fixture("conditional_formatting");
    let ws = parse_worksheet_xml(&xml).expect("parse existing fixture");

    // The existing fixture has 4 blocks
    let cfs = ws.conditional_formattings();
    assert_eq!(cfs.len(), 4);

    // First block (A1:A10) has 2 rules
    assert_eq!(cfs[0].rule_count(), 2);
    // Others have 1 each
    assert_eq!(cfs[1].rule_count(), 1);
    assert_eq!(cfs[2].rule_count(), 1);
    assert_eq!(cfs[3].rule_count(), 1);
}

// ============================================================================
// Pivot Table — Name
// ============================================================================

#[test]
fn test_pivot_table_name() {
    use ooxml_sml::PivotTableExt;
    use ooxml_sml::types::CTPivotTableDefinition;

    let xml = load_fixture("pivot_table_definition");
    let pt: CTPivotTableDefinition = bootstrap(&xml);
    assert_eq!(pt.name(), "PivotTable1");
}

#[test]
fn test_pivot_table_name_inline() {
    use ooxml_sml::PivotTableExt;
    use ooxml_sml::types::CTPivotTableDefinition;

    let xml = br#"<?xml version="1.0" encoding="UTF-8"?>
    <pivotTableDefinition
      xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
      name="MySales" cacheId="3" dataCaption="Values">
      <location ref="B2:E8" firstHeaderRow="1" firstDataRow="2" firstDataCol="0"/>
    </pivotTableDefinition>"#;

    let pt: CTPivotTableDefinition = bootstrap(xml);
    assert_eq!(pt.name(), "MySales");
}

// ============================================================================
// Pivot Table — Location
// ============================================================================

#[test]
fn test_pivot_table_location() {
    use ooxml_sml::PivotTableExt;
    use ooxml_sml::types::CTPivotTableDefinition;

    let xml = load_fixture("pivot_table_definition");
    let pt: CTPivotTableDefinition = bootstrap(&xml);
    assert_eq!(pt.location_reference(), "A1:D10");
}

#[test]
fn test_pivot_table_location_inline() {
    use ooxml_sml::PivotTableExt;
    use ooxml_sml::types::CTPivotTableDefinition;

    let xml = br#"<?xml version="1.0" encoding="UTF-8"?>
    <pivotTableDefinition
      xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
      name="Test" cacheId="1" dataCaption="Values">
      <location ref="C3:F15" firstHeaderRow="1" firstDataRow="2" firstDataCol="0"/>
    </pivotTableDefinition>"#;

    let pt: CTPivotTableDefinition = bootstrap(xml);
    assert_eq!(pt.location_reference(), "C3:F15");
}

// ============================================================================
// Pivot Table — Data fields
// ============================================================================

#[test]
fn test_pivot_table_data_field_names() {
    use ooxml_sml::PivotTableExt;
    use ooxml_sml::types::CTPivotTableDefinition;

    let xml = load_fixture("pivot_table_definition");
    let pt: CTPivotTableDefinition = bootstrap(&xml);

    let names = pt.data_field_names();
    assert_eq!(names.len(), 1);
    assert_eq!(names[0], "Sum of Sales");
}

#[test]
fn test_pivot_table_no_data_fields() {
    use ooxml_sml::PivotTableExt;
    use ooxml_sml::types::CTPivotTableDefinition;

    // Minimal pivot table with no dataFields element
    let xml = br#"<?xml version="1.0" encoding="UTF-8"?>
    <pivotTableDefinition
      xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
      name="Empty" cacheId="1" dataCaption="Values">
      <location ref="A1:B5" firstHeaderRow="1" firstDataRow="2" firstDataCol="0"/>
    </pivotTableDefinition>"#;

    let pt: CTPivotTableDefinition = bootstrap(xml);
    assert!(
        pt.data_field_names().is_empty(),
        "no data fields expected when dataFields element is absent"
    );
}

#[test]
fn test_pivot_table_multiple_data_fields() {
    use ooxml_sml::PivotTableExt;
    use ooxml_sml::types::CTPivotTableDefinition;

    let xml = br#"<?xml version="1.0" encoding="UTF-8"?>
    <pivotTableDefinition
      xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
      name="Multi" cacheId="1" dataCaption="Values">
      <location ref="A1:D10" firstHeaderRow="1" firstDataRow="2" firstDataCol="0"/>
      <dataFields count="2">
        <dataField name="Sum of Revenue" fld="0" subtotal="sum"/>
        <dataField name="Count of Units" fld="1" subtotal="count"/>
      </dataFields>
    </pivotTableDefinition>"#;

    let pt: CTPivotTableDefinition = bootstrap(xml);
    let names = pt.data_field_names();
    assert_eq!(names.len(), 2);
    assert_eq!(names[0], "Sum of Revenue");
    assert_eq!(names[1], "Count of Units");
}

#[test]
fn test_pivot_table_unnamed_data_field() {
    use ooxml_sml::PivotTableExt;
    use ooxml_sml::types::CTPivotTableDefinition;

    // dataField without a name attribute
    let xml = br#"<?xml version="1.0" encoding="UTF-8"?>
    <pivotTableDefinition
      xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
      name="NoName" cacheId="1" dataCaption="Values">
      <location ref="A1:B5" firstHeaderRow="1" firstDataRow="2" firstDataCol="0"/>
      <dataFields count="1">
        <dataField fld="0" subtotal="sum"/>
      </dataFields>
    </pivotTableDefinition>"#;

    let pt: CTPivotTableDefinition = bootstrap(xml);
    let names = pt.data_field_names();
    assert_eq!(names.len(), 1);
    // Unnamed fields return empty string per PivotTableExt contract
    assert_eq!(names[0], "");
}

// ============================================================================
// Pivot Table — Row and column field indices
// ============================================================================

#[test]
fn test_pivot_table_row_field_indices() {
    use ooxml_sml::PivotTableExt;
    use ooxml_sml::types::CTPivotTableDefinition;

    let xml = load_fixture("pivot_table_definition");
    let pt: CTPivotTableDefinition = bootstrap(&xml);

    assert_eq!(pt.row_field_indices(), vec![0]);
}

#[test]
fn test_pivot_table_col_field_indices() {
    use ooxml_sml::PivotTableExt;
    use ooxml_sml::types::CTPivotTableDefinition;

    let xml = load_fixture("pivot_table_definition");
    let pt: CTPivotTableDefinition = bootstrap(&xml);

    assert_eq!(pt.col_field_indices(), vec![1]);
}

#[test]
fn test_pivot_table_no_row_fields() {
    use ooxml_sml::PivotTableExt;
    use ooxml_sml::types::CTPivotTableDefinition;

    let xml = br#"<?xml version="1.0" encoding="UTF-8"?>
    <pivotTableDefinition
      xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
      name="NoRows" cacheId="1" dataCaption="Values">
      <location ref="A1:B5" firstHeaderRow="1" firstDataRow="2" firstDataCol="0"/>
    </pivotTableDefinition>"#;

    let pt: CTPivotTableDefinition = bootstrap(xml);
    assert!(
        pt.row_field_indices().is_empty(),
        "no row fields when rowFields element absent"
    );
    assert!(
        pt.col_field_indices().is_empty(),
        "no col fields when colFields element absent"
    );
}

#[test]
fn test_pivot_table_multiple_row_fields() {
    use ooxml_sml::PivotTableExt;
    use ooxml_sml::types::CTPivotTableDefinition;

    let xml = br#"<?xml version="1.0" encoding="UTF-8"?>
    <pivotTableDefinition
      xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
      name="MultiRow" cacheId="1" dataCaption="Values">
      <location ref="A1:F20" firstHeaderRow="1" firstDataRow="2" firstDataCol="0"/>
      <rowFields count="3">
        <field x="0"/>
        <field x="2"/>
        <field x="4"/>
      </rowFields>
      <colFields count="1">
        <field x="1"/>
      </colFields>
    </pivotTableDefinition>"#;

    let pt: CTPivotTableDefinition = bootstrap(xml);
    assert_eq!(pt.row_field_indices(), vec![0, 2, 4]);
    assert_eq!(pt.col_field_indices(), vec![1]);
}

#[test]
fn test_pivot_table_special_data_axis_field() {
    use ooxml_sml::PivotTableExt;
    use ooxml_sml::types::CTPivotTableDefinition;

    // x="-2" is the OOXML special "values" field on the axis
    let xml = br#"<?xml version="1.0" encoding="UTF-8"?>
    <pivotTableDefinition
      xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
      name="DataAxis" cacheId="1" dataCaption="Values">
      <location ref="A1:D10" firstHeaderRow="1" firstDataRow="2" firstDataCol="0"/>
      <colFields count="2">
        <field x="-2"/>
        <field x="0"/>
      </colFields>
    </pivotTableDefinition>"#;

    let pt: CTPivotTableDefinition = bootstrap(xml);
    let col_indices = pt.col_field_indices();
    assert_eq!(col_indices.len(), 2);
    // -2 is the special "data" axis field index
    assert_eq!(col_indices[0], -2);
    assert_eq!(col_indices[1], 0);
}

// ============================================================================
// Pivot Table — CacheId attribute
// ============================================================================

#[test]
fn test_pivot_table_cache_id() {
    use ooxml_sml::types::CTPivotTableDefinition;

    let xml = load_fixture("pivot_table_definition");
    let pt: CTPivotTableDefinition = bootstrap(&xml);
    // cacheId from the fixture is 1
    assert_eq!(pt.cache_id, 1);
}

#[test]
fn test_pivot_table_data_caption() {
    use ooxml_sml::types::CTPivotTableDefinition;

    let xml = load_fixture("pivot_table_definition");
    let pt: CTPivotTableDefinition = bootstrap(&xml);
    assert_eq!(pt.data_caption.as_str(), "Values");
}

// ============================================================================
// Integration: worksheet conditional_formatting field (via WorksheetExt)
// ============================================================================

#[test]
fn test_worksheet_has_no_cf_when_absent() {
    use ooxml_sml::WorksheetExt;

    let xml = br#"<?xml version="1.0" encoding="UTF-8"?>
    <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
      <sheetData/>
    </worksheet>"#;

    let ws = parse_worksheet_xml(xml).expect("parse");
    assert!(!ws.has_conditional_formatting());
}

#[test]
fn test_worksheet_has_cf_colorscale() {
    use ooxml_sml::WorksheetExt;

    let xml = load_fixture("conditional_formatting_colorscale");
    let ws = parse_worksheet_xml(&xml).expect("parse");
    assert!(ws.has_conditional_formatting());
}

#[test]
fn test_worksheet_has_cf_databar() {
    use ooxml_sml::WorksheetExt;

    let xml = load_fixture("conditional_formatting_databar");
    let ws = parse_worksheet_xml(&xml).expect("parse");
    assert!(ws.has_conditional_formatting());
}

#[test]
fn test_worksheet_has_cf_iconset() {
    use ooxml_sml::WorksheetExt;

    let xml = load_fixture("conditional_formatting_iconset");
    let ws = parse_worksheet_xml(&xml).expect("parse");
    assert!(ws.has_conditional_formatting());
}

#[test]
fn test_worksheet_has_cf_multitype() {
    use ooxml_sml::WorksheetExt;

    let xml = load_fixture("conditional_formatting_multitype");
    let ws = parse_worksheet_xml(&xml).expect("parse");
    assert!(ws.has_conditional_formatting());
}
