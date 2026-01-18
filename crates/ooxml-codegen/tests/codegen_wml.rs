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

    // Basic sanity checks
    assert!(code.contains("pub enum HighlightColor"));
    assert!(
        code.contains("pub struct Body"),
        "Expected Body struct in output"
    );

    // Check for some struct fields
    assert!(
        code.contains("pub struct Document"),
        "Expected Document struct"
    );
    assert!(
        code.contains("pub struct P "),
        "Expected P (paragraph) struct"
    );
}
