//! Integration tests for the ResolvedSheet API.
//!
//! These tests verify that the new spec-compliant generated types work correctly
//! with real XLSX files from the corpus.

use ooxml_sml::{CellExt, RowExt, Workbook};
use std::fs;
use std::path::Path;

/// Test basic resolved_sheet functionality with corpus files.
#[test]
#[ignore] // Run with `cargo test -p ooxml-sml -- --ignored`
fn test_resolved_sheet_corpus() {
    let corpus_path = Path::new("/home/me/git/ooxml/corpora/napierone/XLSX");
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
            .is_some_and(|e| e.eq_ignore_ascii_case("xlsx"))
        {
            total += 1;

            match Workbook::open(&path) {
                Ok(mut wb) => {
                    // Try to load sheets using the new resolved_sheet API
                    match wb.resolved_sheets() {
                        Ok(sheets) => {
                            // Verify we can read cell values
                            for sheet in &sheets {
                                // Test basic sheet properties
                                let _ = sheet.name();
                                let _ = sheet.row_count();
                                let _ = sheet.is_empty();
                                let _ = sheet.dimensions();
                                let _ = sheet.has_auto_filter();
                                let _ = sheet.has_merged_cells();
                                let _ = sheet.has_conditional_formatting();
                                let _ = sheet.has_data_validations();
                                let _ = sheet.has_freeze_panes();

                                // Test row/cell iteration
                                for row in sheet.rows().take(10) {
                                    let _ = row.row_number();
                                    for cell in row.cells_iter().take(10) {
                                        // Test extension traits
                                        let _ = cell.column_number();
                                        let _ = cell.has_formula();
                                        let _ = cell.is_shared_string();
                                        let _ = cell.raw_value();
                                        // Test value resolution
                                        let _ = sheet.cell_value(cell);
                                        let _ = sheet.cell_value_string(cell);
                                    }
                                }
                            }
                            success += 1;
                        }
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

    eprintln!("\n=== ResolvedSheet API Results ===");
    eprintln!("Total: {}", total);
    eprintln!(
        "Success: {} ({:.1}%)",
        success,
        if total > 0 {
            (success as f64 / total as f64) * 100.0
        } else {
            0.0
        }
    );
    eprintln!("Failures: {}", failures.len());

    if !failures.is_empty() {
        eprintln!("\nFirst 10 failures:");
        for (path, err) in failures.iter().take(10) {
            eprintln!("  {}: {}", path, err);
        }
    }

    if total > 0 {
        let success_rate = success as f64 / total as f64;
        assert!(
            success_rate > 0.95,
            "Success rate {:.1}% is below 95% threshold",
            success_rate * 100.0
        );
    }
}

/// Test formula text is correctly parsed.
#[test]
#[ignore]
fn test_formula_text_parsing() {
    let corpus_path = Path::new("/home/me/git/ooxml/corpora/napierone/XLSX");
    if !corpus_path.exists() {
        eprintln!("Corpus not found, skipping");
        return;
    }

    let mut formulas_found = 0;

    for entry in fs::read_dir(corpus_path).unwrap().take(50) {
        let entry = entry.unwrap();
        let path = entry.path();

        if !path
            .extension()
            .is_some_and(|e| e.eq_ignore_ascii_case("xlsx"))
        {
            continue;
        }

        let mut wb = match Workbook::open(&path) {
            Ok(wb) => wb,
            Err(_) => continue,
        };

        let sheets = match wb.resolved_sheets() {
            Ok(s) => s,
            Err(_) => continue,
        };

        for sheet in &sheets {
            for row in sheet.rows() {
                for cell in row.cells_iter() {
                    if cell.has_formula()
                        && let Some(formula) = cell.formula_text()
                        && !formula.is_empty()
                        && formula != "(formula)"
                    {
                        formulas_found += 1;
                    }
                }
            }
        }
    }

    eprintln!("Found {} formulas with text", formulas_found);
    // We expect to find at least some formulas in the corpus
    assert!(formulas_found > 0, "No formulas with text found in corpus");
}
