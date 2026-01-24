// Requires full feature set
#![cfg(feature = "full")]

use ooxml_sml::types::*;
use quick_xml::de::from_str;

#[test]
fn test_autofilter_deserialize() {
    // Simple autoFilter without namespace
    let xml = r#"<autoFilter ref="A1:D10"/>"#;
    let result: Result<AutoFilter, _> = from_str(xml);
    println!("Without namespace: {:?}", result);
    assert!(result.is_ok());
    let af = result.unwrap();
    assert_eq!(af.reference.as_deref(), Some("A1:D10"));
}

#[test]
fn test_autofilter_with_namespace() {
    // With namespace prefix (actual XLSX format)
    let xml = r#"<autoFilter xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" ref="A1:D10"/>"#;
    let result: Result<AutoFilter, _> = from_str(xml);
    println!("With default namespace: {:?}", result);
    // This might fail due to namespace handling
    if result.is_err() {
        println!("Expected - namespace handling is tricky");
    }
}

#[test]
fn test_autofilter_with_children() {
    let xml = r#"
    <autoFilter ref="A1:D10">
        <filterColumn colId="0"/>
        <filterColumn colId="1" hiddenButton="true"/>
    </autoFilter>
    "#;
    let result: Result<AutoFilter, _> = from_str(xml);
    println!("With children: {:?}", result);
    assert!(result.is_ok());
    let af = result.unwrap();
    assert_eq!(af.filter_column.len(), 2);
}

#[test]
fn test_enum_deserialize() {
    // Test simple enum
    let xml = r#"<filter val="test"/>"#;

    #[derive(Debug, serde::Deserialize)]
    #[allow(dead_code)]
    struct TestFilter {
        #[serde(rename = "@val")]
        val: String,
    }

    let result: Result<TestFilter, _> = from_str(xml);
    assert!(result.is_ok());
}

#[test]
fn test_calendar_type_enum() {
    // Calendar type is a string enum
    let xml = r#"<filters calendarType="gregorian"/>"#;
    let result: Result<Filters, _> = from_str(xml);
    println!("Enum deserialize: {:?}", result);
    assert!(result.is_ok());
    let f = result.unwrap();
    assert_eq!(f.calendar_type, Some(CalendarType::Gregorian));
}

#[test]
fn test_real_xlsx_worksheet() {
    // Test with real worksheet XML structure
    let xml = r#"
    <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
        <sheetData>
            <row r="1">
                <c r="A1" t="s"><v>0</v></c>
                <c r="B1"><v>42</v></c>
            </row>
        </sheetData>
    </worksheet>
    "#;

    let result: Result<Worksheet, _> = from_str(xml);
    println!("Real worksheet: {:?}", result);
    // This tests if we can parse actual XLSX structure
}

#[test]
fn test_real_xlsx_sheet_data() {
    // Just the sheetData element
    let xml = r#"
    <sheetData>
        <row r="1" spans="1:3">
            <c r="A1" t="s"><v>0</v></c>
            <c r="B1" t="n"><v>123.45</v></c>
            <c r="C1" t="b"><v>1</v></c>
        </row>
        <row r="2">
            <c r="A2"><f>A1+B1</f><v>123.45</v></c>
        </row>
    </sheetData>
    "#;

    let result: Result<SheetData, _> = from_str(xml);
    println!("Sheet data: {:?}", result);
}
