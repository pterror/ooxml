use ooxml_codegen::{
    CodegenConfig, FeatureMappings, NameMappings, Schema, analyze_schema, generate,
    generate_parsers, generate_serializers, parse_rnc,
};
use std::collections::HashMap;
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
    let pml_path = format!("{}/pml.rnc", spec_dir);
    let shared_path = format!("{}/shared-commonSimpleTypes.rnc", spec_dir);
    // Note: shared-relationshipReference.rnc is NOT included here because PML
    // imports DML types (which include STRelationshipId). Adding it would cause
    // duplicate type definitions and ambiguous imports.

    // Only regenerate if schemas change
    println!("cargo::rerun-if-changed={}", pml_path);
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
    if !Path::new(&pml_path).exists() {
        eprintln!(
            "Warning: Schema not found at {}. Run scripts/download-spec.sh first.",
            pml_path
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

    // Parse and merge the PML schema
    let pml_input = fs::read_to_string(&pml_path).expect("failed to read pml.rnc");
    let pml_schema = parse_rnc(&pml_input).expect("failed to parse pml.rnc");

    // Merge: add PML namespaces and definitions (PML takes precedence for duplicates)
    for ns in pml_schema.namespaces {
        if !combined_schema
            .namespaces
            .iter()
            .any(|n| n.prefix == ns.prefix)
        {
            combined_schema.namespaces.push(ns);
        }
    }
    combined_schema.definitions.extend(pml_schema.definitions);

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
    // DML types use "a_" prefix in schema (e.g., a_CT_Color)
    // The tuple is (crate_path, module_name) where module_name is used for name mapping lookups
    let mut cross_crate_type_prefix = HashMap::new();
    cross_crate_type_prefix.insert(
        "a_".to_string(),
        ("ooxml_dml::types::".to_string(), "dml".to_string()),
    );

    let config = CodegenConfig {
        strip_prefix: Some("p_".to_string()),
        module_name: "pml".to_string(),
        name_mappings,
        feature_mappings,
        // PML uses p: namespace prefix in real PPTX files
        xml_serialize_prefix: Some("p".to_string()),
        // PML references DML types for shared drawing elements
        cross_crate_imports: vec!["ooxml_dml::types::*".to_string()],
        // Map DML schema names to ooxml_dml crate types
        cross_crate_type_prefix,
        ..Default::default()
    };

    // Run static analysis if OOXML_ANALYZE is set
    if std::env::var("OOXML_ANALYZE").is_ok() {
        eprintln!("\n=== Static Analysis: PML ===");
        let report = analyze_schema(&combined_schema, &config);
        report.print("pml");
        eprintln!();
    }

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
