use ooxml_codegen::{CodegenConfig, generate, parse_rnc};
use std::fs;

#[test]
fn test_eg_definitions() {
    use ooxml_codegen::ast::Pattern;

    let path = concat!(
        env!("CARGO_MANIFEST_DIR"),
        "/../../spec/OfficeOpenXML-RELAXNG-Transitional/wml.rnc"
    );
    let input = fs::read_to_string(path).expect("failed to read wml.rnc");
    let schema = parse_rnc(&input).expect("failed to parse wml.rnc");

    let eg_defs: Vec<_> = schema
        .definitions
        .iter()
        .filter(|d| d.name.contains("_EG_"))
        .collect();

    println!("Found {} EG_* definitions:", eg_defs.len());

    // Check which ones are element choices
    for d in &eg_defs {
        let is_choice = matches!(&d.pattern, Pattern::Choice(_));
        let has_direct_element = if let Pattern::Choice(variants) = &d.pattern {
            variants.iter().any(is_direct_element)
        } else {
            false
        };
        if is_choice && has_direct_element {
            println!("  [ELEMENT CHOICE] {}", d.name);
        }
    }
}

fn is_direct_element(pattern: &ooxml_codegen::ast::Pattern) -> bool {
    use ooxml_codegen::ast::Pattern;
    match pattern {
        Pattern::Element { .. } => true,
        Pattern::Optional(inner) => is_direct_element(inner),
        _ => false,
    }
}

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
