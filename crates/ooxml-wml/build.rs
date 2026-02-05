use ooxml_codegen::{
    CodegenConfig, FeatureMappings, NameMappings, Schema, analyze_schema, generate,
    generate_parsers, generate_serializers, parse_rnc,
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
    let wml_path = format!("{}/wml.rnc", spec_dir);
    let shared_path = format!("{}/shared-commonSimpleTypes.rnc", spec_dir);
    let rel_path = format!("{}/shared-relationshipReference.rnc", spec_dir);

    // Only regenerate if schemas change
    println!("cargo::rerun-if-changed={}", wml_path);
    println!("cargo::rerun-if-changed={}", shared_path);
    println!("cargo::rerun-if-changed={}", rel_path);
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
    if !Path::new(&wml_path).exists() {
        eprintln!(
            "Warning: Schema not found at {}. Run scripts/download-spec.sh first.",
            wml_path
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

    // Parse and merge the relationship reference schema (for r:id, r:embed, etc.)
    if Path::new(&rel_path).exists() {
        let rel_input =
            fs::read_to_string(&rel_path).expect("failed to read relationship references");
        let rel_schema = parse_rnc(&rel_input).expect("failed to parse relationship references");
        for ns in rel_schema.namespaces {
            if !combined_schema
                .namespaces
                .iter()
                .any(|n| n.prefix == ns.prefix)
            {
                combined_schema.namespaces.push(ns);
            }
        }
        combined_schema.definitions.extend(rel_schema.definitions);
    }

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
        strip_prefix: Some("w_".to_string()),
        module_name: "wml".to_string(),
        name_mappings,
        feature_mappings,
        // WML uses w: namespace prefix in real DOCX files
        xml_serialize_prefix: Some("w".to_string()),
        ..Default::default()
    };

    // Run static analysis if OOXML_ANALYZE is set
    if std::env::var("OOXML_ANALYZE").is_ok() {
        eprintln!("\n=== Static Analysis: WML ===");
        let report = analyze_schema(&combined_schema, &config);
        report.print("wml");
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
