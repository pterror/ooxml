use ooxml_codegen::{
    CodegenConfig, FeatureMappings, NameMappings, Schema, generate, generate_parsers,
    generate_serializers, parse_rnc,
};
use std::fs;
use std::path::Path;

fn main() {
    let spec_dir = concat!(
        env!("CARGO_MANIFEST_DIR"),
        "/../../spec/OfficeOpenXML-RELAXNG-Transitional"
    );
    let names_path = concat!(env!("CARGO_MANIFEST_DIR"), "/../../spec/ooxml-names.yaml");
    let features_path = concat!(
        env!("CARGO_MANIFEST_DIR"),
        "/../../spec/ooxml-features.yaml"
    );

    // Paths to schemas
    let dml_path = format!("{}/dml-main.rnc", spec_dir);
    let shared_path = format!("{}/shared-commonSimpleTypes.rnc", spec_dir);

    // Only regenerate if schemas change
    println!("cargo::rerun-if-changed={}", dml_path);
    println!("cargo::rerun-if-changed={}", shared_path);
    println!("cargo::rerun-if-changed={}", names_path);
    println!("cargo::rerun-if-changed={}", features_path);
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

    // Load name mappings if available
    let name_mappings = if Path::new(names_path).exists() {
        match NameMappings::from_yaml_file(Path::new(names_path)) {
            Ok(mappings) => {
                eprintln!("Loaded name mappings from {}", names_path);
                Some(mappings)
            }
            Err(e) => {
                eprintln!("Warning: Failed to load name mappings: {}", e);
                None
            }
        }
    } else {
        None
    };

    // Load feature mappings if available
    let feature_mappings = if Path::new(features_path).exists() {
        match FeatureMappings::from_yaml_file(Path::new(features_path)) {
            Ok(mappings) => {
                eprintln!("Loaded feature mappings from {}", features_path);
                Some(mappings)
            }
            Err(e) => {
                eprintln!("Warning: Failed to load feature mappings: {}", e);
                None
            }
        }
    } else {
        None
    };

    // Generate Rust code
    let config = CodegenConfig {
        strip_prefix: Some("a_".to_string()),
        module_name: "dml".to_string(),
        name_mappings,
        feature_mappings,
        ..Default::default()
    };
    let code = generate(&combined_schema, &config);

    // Write the generated code
    fs::write(&dest_path, code).expect("failed to write generated types");
    eprintln!(
        "Generated {} bytes to src/generated.rs",
        dest_path.metadata().map(|m| m.len()).unwrap_or(0)
    );

    // Generate event-based parsers
    // Enable with OOXML_GENERATE_PARSERS=1
    if std::env::var("OOXML_GENERATE_PARSERS").is_ok() {
        let parser_dest = Path::new(env!("CARGO_MANIFEST_DIR")).join("src/generated_parsers.rs");
        let parser_code = generate_parsers(&combined_schema, &config);
        fs::write(&parser_dest, parser_code).expect("failed to write generated parsers");
        eprintln!(
            "Generated {} bytes to src/generated_parsers.rs",
            parser_dest.metadata().map(|m| m.len()).unwrap_or(0)
        );
    }

    // Generate ToXml serializers
    // Enable with OOXML_GENERATE_SERIALIZERS=1
    if std::env::var("OOXML_GENERATE_SERIALIZERS").is_ok() {
        let serializer_dest =
            Path::new(env!("CARGO_MANIFEST_DIR")).join("src/generated_serializers.rs");
        let serializer_code = generate_serializers(&combined_schema, &config);
        fs::write(&serializer_dest, serializer_code)
            .expect("failed to write generated serializers");
        eprintln!(
            "Generated {} bytes to src/generated_serializers.rs",
            serializer_dest.metadata().map(|m| m.len()).unwrap_or(0)
        );
    }
}
