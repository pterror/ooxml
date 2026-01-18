use ooxml_codegen::parse_rnc;
use std::fs;

#[test]
fn test_parse_wml_rnc() {
    let path = concat!(
        env!("CARGO_MANIFEST_DIR"),
        "/../../spec/OfficeOpenXML-RELAXNG-Transitional/wml.rnc"
    );
    let input = fs::read_to_string(path).expect("failed to read wml.rnc");

    let schema = parse_rnc(&input).expect("failed to parse wml.rnc");

    // Verify we parsed the expected content
    assert_eq!(schema.namespaces.len(), 10);
    assert!(schema.definitions.len() > 400, "expected 400+ definitions");

    // Check some known definitions exist
    let def_names: Vec<_> = schema.definitions.iter().map(|d| d.name.as_str()).collect();
    assert!(def_names.contains(&"w_CT_Document"));
    assert!(def_names.contains(&"w_CT_Body"));
    assert!(def_names.contains(&"w_CT_P"));
    assert!(def_names.contains(&"w_CT_R"));
}
