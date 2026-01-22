use ooxml_codegen::{CodegenConfig, Schema, generate, parse_rnc};
use std::fs;
use std::path::Path;

fn main() {
    let spec_dir = concat!(
        env!("CARGO_MANIFEST_DIR"),
        "/../../spec/OfficeOpenXML-RELAXNG-Transitional"
    );

    // Paths to schemas
    let dml_path = format!("{}/dml-main.rnc", spec_dir);
    let shared_path = format!("{}/shared-commonSimpleTypes.rnc", spec_dir);

    // Only regenerate if schemas change
    println!("cargo::rerun-if-changed={}", dml_path);
    println!("cargo::rerun-if-changed={}", shared_path);
    println!("cargo::rerun-if-changed=build.rs");

    // The generated file is committed at src/generated.rs
    // Only regenerate if OOXML_REGENERATE is set and specs exist
    let should_regenerate = std::env::var("OOXML_REGENERATE").is_ok();
    let dest_path = Path::new(env!("CARGO_MANIFEST_DIR")).join("src/generated.rs");

    if !should_regenerate {
        // Use the committed generated.rs - nothing to do
        return;
    }

    // Check if schema exists
    if !Path::new(&dml_path).exists() {
        eprintln!(
            "Warning: Schema not found at {}. Run scripts/download-spec.sh first.",
            dml_path
        );
        return;
    }

    eprintln!("Regenerating src/generated.rs from schemas...");

    // Parse the shared types schema first
    let mut combined_schema = if Path::new(&shared_path).exists() {
        let shared_input = fs::read_to_string(&shared_path).expect("failed to read shared types");
        parse_rnc(&shared_input).expect("failed to parse shared types")
    } else {
        Schema {
            namespaces: vec![],
            definitions: vec![],
        }
    };

    // Parse and merge the DML schema
    let dml_input = fs::read_to_string(&dml_path).expect("failed to read dml-main.rnc");
    let dml_schema = parse_rnc(&dml_input).expect("failed to parse dml-main.rnc");

    // Merge: add DML namespaces and definitions (DML takes precedence for duplicates)
    for ns in dml_schema.namespaces {
        if !combined_schema
            .namespaces
            .iter()
            .any(|n| n.prefix == ns.prefix)
        {
            combined_schema.namespaces.push(ns);
        }
    }
    combined_schema.definitions.extend(dml_schema.definitions);

    // Generate Rust code
    let config = CodegenConfig {
        strip_prefix: Some("a_".to_string()),
        module_name: "dml".to_string(),
    };
    let code = generate(&combined_schema, &config);

    // Write the generated code
    fs::write(&dest_path, code).expect("failed to write generated types");
    eprintln!(
        "Generated {} bytes to src/generated.rs",
        dest_path.metadata().map(|m| m.len()).unwrap_or(0)
    );
}
