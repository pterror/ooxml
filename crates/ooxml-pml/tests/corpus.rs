//! Corpus test - try parsing all PPTX files in the NapierOne corpus.

use ooxml_pml::Presentation;
use std::fs;
use std::path::Path;

#[test]
#[ignore] // Run with `cargo test -p ooxml-pml -- --ignored`
fn test_napierone_pptx_corpus() {
    let corpus_path = Path::new("/home/me/git/ooxml/corpora/napierone/PPTX");
    if !corpus_path.exists() {
        eprintln!("Corpus not found at {:?}, skipping", corpus_path);
        return;
    }

    let mut total = 0;
    let mut success = 0;
    let mut failures: Vec<(String, String)> = Vec::new();

    for entry in fs::read_dir(corpus_path).unwrap() {
        let entry = entry.unwrap();
        let path = entry.path();

        if path
            .extension()
            .is_some_and(|e| e.eq_ignore_ascii_case("pptx"))
        {
            total += 1;

            match Presentation::open(&path) {
                Ok(mut pres) => {
                    // Try to actually load the slides
                    match pres.slides() {
                        Ok(_) => success += 1,
                        Err(e) => {
                            failures.push((path.display().to_string(), e.to_string()));
                        }
                    }
                }
                Err(e) => {
                    failures.push((path.display().to_string(), e.to_string()));
                }
            }

            if total % 500 == 0 {
                eprintln!("Processed {}/{} files...", success, total);
            }
        }
    }

    eprintln!("\n=== Results ===");
    eprintln!("Total: {}", total);
    eprintln!(
        "Success: {} ({:.1}%)",
        success,
        (success as f64 / total as f64) * 100.0
    );
    eprintln!("Failures: {}", failures.len());

    if !failures.is_empty() {
        eprintln!("\nFirst 10 failures:");
        for (path, err) in failures.iter().take(10) {
            eprintln!("  {}: {}", path, err);
        }
    }

    // Assert a reasonable success rate
    let success_rate = success as f64 / total as f64;
    assert!(
        success_rate > 0.95,
        "Success rate {:.1}% is below 95% threshold",
        success_rate * 100.0
    );
}
