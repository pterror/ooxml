use ooxml_codegen::{CodegenConfig, generate, parse_rnc};
use std::env;
use std::fs;
use std::path::Path;

fn main() {
    let out_dir = env::var("OUT_DIR").unwrap();
    let dest_path = Path::new(&out_dir).join("wml_types.rs");

    // Path to the WML schema
    let schema_path = concat!(
        env!("CARGO_MANIFEST_DIR"),
        "/../../spec/OfficeOpenXML-RELAXNG-Transitional/wml.rnc"
    );

    // Only regenerate if the schema changes
    println!("cargo::rerun-if-changed={}", schema_path);
    println!("cargo::rerun-if-changed=build.rs");

    // Check if schema exists (it might not be downloaded yet)
    if !Path::new(schema_path).exists() {
        // Generate a stub file with a warning
        fs::write(
            &dest_path,
            "// Schema not found. Run scripts/download-spec.sh first.\n",
        )
        .expect("failed to write stub");
        return;
    }

    // Parse the schema
    let input = fs::read_to_string(schema_path).expect("failed to read wml.rnc");
    let schema = parse_rnc(&input).expect("failed to parse wml.rnc");

    // Generate Rust code
    let config = CodegenConfig {
        strip_prefix: Some("w_".to_string()),
        module_name: "wml".to_string(),
    };
    let code = generate(&schema, &config);

    // Write the generated code
    fs::write(&dest_path, code).expect("failed to write generated types");
}
