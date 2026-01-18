use ooxml_codegen::{CodegenConfig, Schema, generate, parse_rnc};
use std::env;
use std::fs;
use std::path::Path;

fn main() {
    let out_dir = env::var("OUT_DIR").unwrap();
    let dest_path = Path::new(&out_dir).join("wml_types.rs");

    let spec_dir = concat!(
        env!("CARGO_MANIFEST_DIR"),
        "/../../spec/OfficeOpenXML-RELAXNG-Transitional"
    );

    // Paths to schemas
    let wml_path = format!("{}/wml.rnc", spec_dir);
    let shared_path = format!("{}/shared-commonSimpleTypes.rnc", spec_dir);

    // Only regenerate if schemas change
    println!("cargo::rerun-if-changed={}", wml_path);
    println!("cargo::rerun-if-changed={}", shared_path);
    println!("cargo::rerun-if-changed=build.rs");

    // Check if schema exists (it might not be downloaded yet)
    if !Path::new(&wml_path).exists() {
        fs::write(
            &dest_path,
            "// Schema not found. Run scripts/download-spec.sh first.\n",
        )
        .expect("failed to write stub");
        return;
    }

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

    // Parse and merge the WML schema
    let wml_input = fs::read_to_string(&wml_path).expect("failed to read wml.rnc");
    let wml_schema = parse_rnc(&wml_input).expect("failed to parse wml.rnc");

    // Merge: add WML namespaces and definitions (WML takes precedence for duplicates)
    for ns in wml_schema.namespaces {
        if !combined_schema
            .namespaces
            .iter()
            .any(|n| n.prefix == ns.prefix)
        {
            combined_schema.namespaces.push(ns);
        }
    }
    combined_schema.definitions.extend(wml_schema.definitions);

    // Generate Rust code
    let config = CodegenConfig {
        strip_prefix: Some("w_".to_string()),
        module_name: "wml".to_string(),
    };
    let code = generate(&combined_schema, &config);

    // Write the generated code
    fs::write(&dest_path, code).expect("failed to write generated types");
}
