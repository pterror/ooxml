use ooxml_codegen::{CodegenConfig, generate, parse_rnc};
use std::fs;

#[test]
fn test_generate_wml() {
    let path = concat!(
        env!("CARGO_MANIFEST_DIR"),
        "/../../spec/OfficeOpenXML-RELAXNG-Transitional/wml.rnc"
    );
    let input = fs::read_to_string(path).expect("failed to read wml.rnc");

    let schema = parse_rnc(&input).expect("failed to parse wml.rnc");

    let config = CodegenConfig {
        strip_prefix: Some("w_".to_string()),
        module_name: "wml".to_string(),
    };

    let code = generate(&schema, &config);

    // Print first 100 lines for inspection
    for line in code.lines().take(100) {
        println!("{}", line);
    }

    // Basic sanity checks (types keep CT_/ST_ prefixes)
    assert!(code.contains("pub enum STHighlightColor"));
    assert!(
        code.contains("pub struct CTBody"),
        "Expected CTBody struct in output"
    );

    // Check for some struct fields
    assert!(
        code.contains("pub struct CTDocument"),
        "Expected CTDocument struct"
    );
    assert!(
        code.contains("pub struct CTP "),
        "Expected CTP (paragraph) struct"
    );
}
