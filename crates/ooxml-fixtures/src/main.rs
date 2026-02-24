//! Binary that generates all OOXML test fixtures.
//!
//! For each fixture defined in `spec/ooxml-fixture-spec.yaml` (and any
//! handwritten custom fixtures) this binary:
//!
//! 1. Calls the fixture function to produce bytes and assertions.
//! 2. Writes the raw bytes to `fixtures/{fixture.path}`.
//! 3. Writes a companion JSON manifest to `fixtures/{fixture.path}.json`.
//!
//! Run with:
//! ```text
//! cargo run -p ooxml-fixtures --bin generate-fixtures
//! ```

use ooxml_fixtures::manifest::write_manifest;
use std::path::Path;

fn main() {
    // Collect all fixtures from the generated module.
    // When spec/ooxml-fixture-spec.yaml exists and codegen runs, this vec will
    // be populated by the generated fixture functions.
    let fixtures: Vec<ooxml_fixtures::Fixture> = collect_fixtures();

    if fixtures.is_empty() {
        println!("No fixtures to generate (spec/ooxml-fixture-spec.yaml not found or empty).");
        return;
    }

    let out_dir = Path::new("fixtures");

    let mut written = 0usize;
    let mut errors = 0usize;

    for fixture in &fixtures {
        // Write OOXML file bytes.
        let ooxml_path = out_dir.join(fixture.path);
        if let Some(parent) = ooxml_path.parent()
            && let Err(e) = std::fs::create_dir_all(parent)
        {
            eprintln!(
                "ERROR: could not create directory {}: {e}",
                parent.display()
            );
            errors += 1;
            continue;
        }
        if let Err(e) = std::fs::write(&ooxml_path, &fixture.bytes) {
            eprintln!("ERROR: could not write {}: {e}", ooxml_path.display());
            errors += 1;
            continue;
        }

        // Write companion JSON manifest.
        if let Err(e) = write_manifest(fixture, out_dir) {
            eprintln!("ERROR: could not write manifest for {}: {e}", fixture.path);
            errors += 1;
            continue;
        }

        println!("  wrote {}", fixture.path);
        written += 1;
    }

    println!("\n{written} fixture(s) written, {errors} error(s).",);

    if errors > 0 {
        std::process::exit(1);
    }
}

/// Collect all fixtures: generated ones (from `spec/ooxml-fixture-spec.yaml`)
/// plus any handwritten custom fixtures.
fn collect_fixtures() -> Vec<ooxml_fixtures::Fixture> {
    // Generated fixtures are produced by build.rs and live in
    // `src/generated/mod.rs`.  The generated module exposes a
    // `all_fixtures()` function when it is non-empty.
    //
    // Custom fixtures are handwritten in `src/custom/{wml,sml,pml}.rs`.

    let mut fixtures: Vec<ooxml_fixtures::Fixture> = Vec::new();

    // Pull in generated fixtures (empty until spec YAML is present).
    fixtures.extend(ooxml_fixtures::generated::all_fixtures());

    fixtures
}
