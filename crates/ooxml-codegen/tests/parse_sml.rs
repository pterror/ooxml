use ooxml_codegen::{lex_rnc, parse_rnc};
use std::fs;

#[test]
fn test_parse_sml() {
    let path = concat!(
        env!("CARGO_MANIFEST_DIR"),
        "/../../spec/OfficeOpenXML-RELAXNG-Transitional/sml.rnc"
    );
    let input = fs::read_to_string(path).expect("failed to read sml.rnc");

    // First, let's see the tokens around the error position
    let tokens = lex_rnc(&input).expect("failed to lex sml.rnc");
    println!("Total tokens: {}", tokens.len());
    println!("\nTokens 15585-15605:");
    for (i, tok) in tokens.iter().enumerate().skip(15585).take(25) {
        println!("  {}: {:?}", i, tok);
    }

    let schema = parse_rnc(&input).expect("failed to parse sml.rnc");

    println!(
        "Parsed {} namespaces, {} definitions",
        schema.namespaces.len(),
        schema.definitions.len()
    );

    // Print first few definitions
    for def in schema.definitions.iter().take(10) {
        println!("  {}", def.name);
    }
}
