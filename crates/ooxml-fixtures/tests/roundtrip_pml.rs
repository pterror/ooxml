//! Fixture roundtrip tests for PML (PPTX).
//!
//! For every `.json` manifest under `fixtures/pml/` this test:
//! 1. Opens the companion `.pptx` file with the PML reader.
//! 2. Evaluates every assertion in the manifest.
//! 3. Collects all failures and reports them at the end.

use std::fs;
use std::io::Cursor;
use std::path::{Path, PathBuf};

use ooxml_dml::ext::{TextParagraphExt, TextRunExt};
use ooxml_dml::types::{EGColorChoice, EGFillProperties, EGGeometry};
use ooxml_pml::ext::ShapeExt;
use ooxml_pml::presentation::Presentation;
use serde_json::Value;

type PmlPresentation = Presentation<Cursor<Vec<u8>>>;

// ---------------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------------

fn fixtures_dir() -> PathBuf {
    let manifest_dir = env!("CARGO_MANIFEST_DIR");
    Path::new(manifest_dir).join("../../fixtures/pml")
}

fn find_json_files(dir: &Path) -> Vec<PathBuf> {
    let mut out = Vec::new();
    if let Ok(entries) = fs::read_dir(dir) {
        for entry in entries.flatten() {
            let path = entry.path();
            if path.is_dir() {
                out.extend(find_json_files(&path));
            } else if path.extension().is_some_and(|e| e == "json") {
                out.push(path);
            }
        }
    }
    out.sort();
    out
}

/// Extract the preset geometry type string of a shape, e.g. "rect", "ellipse".
fn shape_type_str(shape: &ooxml_pml::types::Shape) -> Option<String> {
    match shape.shape_properties.geometry.as_deref()? {
        EGGeometry::PrstGeom(pg) => Some(pg.preset.to_string()),
        _ => None,
    }
}

/// Extract the run color as an "RRGGBB" hex string from a DML text run.
fn run_color_hex(run: &ooxml_dml::types::TextRun) -> Option<String> {
    let fill = run.r_pr.as_ref()?.fill_properties.as_deref()?;
    let solid = match fill {
        EGFillProperties::SolidFill(s) => s,
        _ => return None,
    };
    let color_choice = solid.color_choice.as_deref()?;
    let srgb = match color_choice {
        EGColorChoice::SrgbClr(c) => c,
        _ => return None,
    };
    let b = &srgb.value;
    if b.len() >= 3 {
        Some(format!("{:02X}{:02X}{:02X}", b[0], b[1], b[2]))
    } else {
        None
    }
}

// ---------------------------------------------------------------------------
// Assertion dispatch
// ---------------------------------------------------------------------------

fn check_assertion(
    slide_count: usize,
    slides: &[ooxml_pml::Slide],
    a: &Value,
) -> Result<(), String> {
    let t = a["type"].as_str().unwrap_or("unknown");

    match t {
        // ---- presentation-level ----------------------------------------------
        "slide_count" => {
            let expected = a["expected"].as_u64().unwrap() as u32;
            let got = slide_count as u32;
            if got != expected {
                return Err(format!("slide_count: expected {expected}, got {got}"));
            }
        }

        "shape_count" => {
            let si = a["slide"].as_u64().unwrap() as usize;
            let expected = a["expected"].as_u64().unwrap() as u32;
            let got = slides.get(si).map(|s| s.shapes().len() as u32).unwrap_or(0);
            if got != expected {
                return Err(format!("shape_count[{si}]: expected {expected}, got {got}"));
            }
        }

        "shape_text" => {
            let si = a["slide"].as_u64().unwrap() as usize;
            let shi = a["shape"].as_u64().unwrap() as usize;
            let expected = a["expected"].as_str().unwrap();
            let got: String = slides
                .get(si)
                .and_then(|s| s.shapes().get(shi))
                .and_then(|shape| shape.text())
                .unwrap_or_default();
            if got != expected {
                return Err(format!(
                    "shape_text[{si}][{shi}]: expected {expected:?}, got {got:?}"
                ));
            }
        }

        "shape_type" => {
            let si = a["slide"].as_u64().unwrap() as usize;
            let shi = a["shape"].as_u64().unwrap() as usize;
            let expected = a["expected"].as_str().unwrap();
            let got: Option<String> = slides
                .get(si)
                .and_then(|s| s.shapes().get(shi))
                .and_then(shape_type_str);
            let got_ref = got.as_deref().unwrap_or("none");
            if got_ref != expected {
                return Err(format!(
                    "shape_type[{si}][{shi}]: expected {expected:?}, got {got_ref:?}"
                ));
            }
        }

        // ---- PML run-level ---------------------------------------------------
        "pml_run_text" => {
            let si = a["slide"].as_u64().unwrap() as usize;
            let shi = a["shape"].as_u64().unwrap() as usize;
            let pi = a["para"].as_u64().unwrap() as usize;
            let ri = a["run"].as_u64().unwrap() as usize;
            let expected = a["expected"].as_str().unwrap();
            let got: Option<String> = slides
                .get(si)
                .and_then(|s| s.shapes().get(shi))
                .and_then(|shape| shape.paragraphs().get(pi))
                .and_then(|para| para.runs().into_iter().nth(ri))
                .map(|run| run.text().to_string());
            let got_str = got.as_deref().unwrap_or("");
            if got_str != expected {
                return Err(format!(
                    "pml_run_text[{si}][{shi}][{pi}][{ri}]: expected {expected:?}, got {got_str:?}"
                ));
            }
        }

        "pml_run_bold" => {
            let si = a["slide"].as_u64().unwrap() as usize;
            let shi = a["shape"].as_u64().unwrap() as usize;
            let pi = a["para"].as_u64().unwrap() as usize;
            let ri = a["run"].as_u64().unwrap() as usize;
            let expected = a["expected"].as_bool().unwrap();
            let got = slides
                .get(si)
                .and_then(|s| s.shapes().get(shi))
                .and_then(|shape| shape.paragraphs().get(pi))
                .and_then(|para| para.runs().into_iter().nth(ri))
                .map(|run| run.is_bold())
                .unwrap_or(false);
            if got != expected {
                return Err(format!(
                    "pml_run_bold[{si}][{shi}][{pi}][{ri}]: expected {expected}, got {got}"
                ));
            }
        }

        "pml_run_italic" => {
            let si = a["slide"].as_u64().unwrap() as usize;
            let shi = a["shape"].as_u64().unwrap() as usize;
            let pi = a["para"].as_u64().unwrap() as usize;
            let ri = a["run"].as_u64().unwrap() as usize;
            let expected = a["expected"].as_bool().unwrap();
            let got = slides
                .get(si)
                .and_then(|s| s.shapes().get(shi))
                .and_then(|shape| shape.paragraphs().get(pi))
                .and_then(|para| para.runs().into_iter().nth(ri))
                .map(|run| run.is_italic())
                .unwrap_or(false);
            if got != expected {
                return Err(format!(
                    "pml_run_italic[{si}][{shi}][{pi}][{ri}]: expected {expected}, got {got}"
                ));
            }
        }

        "pml_run_color" => {
            let si = a["slide"].as_u64().unwrap() as usize;
            let shi = a["shape"].as_u64().unwrap() as usize;
            let pi = a["para"].as_u64().unwrap() as usize;
            let ri = a["run"].as_u64().unwrap() as usize;
            let expected: Option<&str> = a["expected"].as_str();
            let got: Option<String> = slides
                .get(si)
                .and_then(|s| s.shapes().get(shi))
                .and_then(|shape| shape.paragraphs().get(pi))
                .and_then(|para| para.runs().into_iter().nth(ri))
                .and_then(run_color_hex);
            let got_ref = got.as_deref();
            if got_ref != expected {
                return Err(format!(
                    "pml_run_color[{si}][{shi}][{pi}][{ri}]: expected {expected:?}, got {got_ref:?}"
                ));
            }
        }

        "pml_run_font_size" => {
            let si = a["slide"].as_u64().unwrap() as usize;
            let shi = a["shape"].as_u64().unwrap() as usize;
            let pi = a["para"].as_u64().unwrap() as usize;
            let ri = a["run"].as_u64().unwrap() as usize;
            let expected = a["expected"].as_f64().unwrap();
            let tolerance = a["tolerance"].as_f64().unwrap_or(0.0);
            // font_size() returns hundredths of a point; convert to points.
            let got: Option<f64> = slides
                .get(si)
                .and_then(|s| s.shapes().get(shi))
                .and_then(|shape| shape.paragraphs().get(pi))
                .and_then(|para| para.runs().into_iter().nth(ri))
                .and_then(|run| run.font_size())
                .map(|hp| hp as f64 / 100.0);
            match got {
                None => {
                    return Err(format!(
                        "pml_run_font_size[{si}][{shi}][{pi}][{ri}]: expected {expected}, got None"
                    ));
                }
                Some(v) if (v - expected).abs() > tolerance => {
                    return Err(format!(
                        "pml_run_font_size[{si}][{shi}][{pi}][{ri}]: expected {expected} ±{tolerance}, got {v}"
                    ));
                }
                _ => {}
            }
        }

        // ---- PML misc --------------------------------------------------------
        "slide_has_notes" => {
            let si = a["slide"].as_u64().unwrap() as usize;
            let expected = a["expected"].as_bool().unwrap();
            let got = slides.get(si).map(|s| s.has_notes()).unwrap_or(false);
            if got != expected {
                return Err(format!(
                    "slide_has_notes[{si}]: expected {expected}, got {got}"
                ));
            }
        }

        "notes_text" => {
            let si = a["slide"].as_u64().unwrap() as usize;
            let expected = a["expected"].as_str().unwrap();
            let got = slides.get(si).and_then(|s| s.notes()).unwrap_or("");
            if got != expected {
                return Err(format!(
                    "notes_text[{si}]: expected {expected:?}, got {got:?}"
                ));
            }
        }

        "pml_image_count" => {
            let si = a["slide"].as_u64().unwrap() as usize;
            let expected = a["expected"].as_u64().unwrap() as u32;
            let got = slides
                .get(si)
                .map(|s| s.pictures().len() as u32)
                .unwrap_or(0);
            if got != expected {
                return Err(format!(
                    "pml_image_count[{si}]: expected {expected}, got {got}"
                ));
            }
        }

        "has_transition" => {
            let si = a["slide"].as_u64().unwrap() as usize;
            let expected = a["expected"].as_bool().unwrap();
            let got = slides.get(si).map(|s| s.has_transition()).unwrap_or(false);
            if got != expected {
                return Err(format!(
                    "has_transition[{si}]: expected {expected}, got {got}"
                ));
            }
        }

        other => {
            return Err(format!("unknown assertion type: {other:?}"));
        }
    }

    Ok(())
}

// ---------------------------------------------------------------------------
// Test entry point
// ---------------------------------------------------------------------------

#[test]
fn roundtrip_pml() {
    let dir = fixtures_dir();
    let json_files = find_json_files(&dir);
    assert!(
        !json_files.is_empty(),
        "no PML fixture JSON files found in {}",
        dir.display()
    );

    let mut failures: Vec<String> = Vec::new();

    for json_path in json_files {
        let pptx_path = json_path.with_extension("");

        let json_bytes = match fs::read(&json_path) {
            Ok(b) => b,
            Err(e) => {
                failures.push(format!("[{}] read json: {e}", json_path.display()));
                continue;
            }
        };
        let manifest: Value = match serde_json::from_slice(&json_bytes) {
            Ok(v) => v,
            Err(e) => {
                failures.push(format!("[{}] parse json: {e}", json_path.display()));
                continue;
            }
        };

        let pptx_bytes = match fs::read(&pptx_path) {
            Ok(b) => b,
            Err(e) => {
                failures.push(format!("[{}] read pptx: {e}", pptx_path.display()));
                continue;
            }
        };

        let mut prs: PmlPresentation = match Presentation::from_reader(Cursor::new(pptx_bytes)) {
            Ok(p) => p,
            Err(e) => {
                failures.push(format!("[{}] open presentation: {e}", pptx_path.display()));
                continue;
            }
        };

        let slide_count = prs.slide_count();
        let slides = match prs.slides() {
            Ok(s) => s,
            Err(e) => {
                failures.push(format!("[{}] load slides: {e}", pptx_path.display()));
                continue;
            }
        };

        let assertions = manifest["assertions"].as_array().unwrap();
        for assertion in assertions {
            if let Err(msg) = check_assertion(slide_count, &slides, assertion) {
                failures.push(format!("[{}] {msg}", pptx_path.display()));
            }
        }
    }

    if !failures.is_empty() {
        panic!(
            "{} PML roundtrip failure(s):\n{}",
            failures.len(),
            failures.join("\n")
        );
    }
}
