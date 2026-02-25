//! Fixture roundtrip tests for WML (DOCX).
//!
//! For every `.json` manifest under `fixtures/wml/` this test:
//! 1. Opens the companion `.docx` file with the WML reader.
//! 2. Evaluates every assertion in the manifest.
//! 3. Collects all failures and reports them at the end.

use std::collections::BTreeSet;
use std::fs;
use std::io::Cursor;
use std::path::{Path, PathBuf};

use ooxml_wml::Document;
use ooxml_wml::ext::{
    BodyExt, CellExt, HyperlinkExt, ParagraphExt, RowExt, RunExt, RunPropertiesExt, TableExt,
};
use ooxml_wml::types;
use serde_json::Value;

type WmlDoc = Document<Cursor<Vec<u8>>>;

// ---------------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------------

fn fixtures_dir() -> PathBuf {
    let manifest_dir = env!("CARGO_MANIFEST_DIR");
    Path::new(manifest_dir).join("../../fixtures/wml")
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

/// Find the URL for the hyperlink that contains run `run_idx` in `para`.
///
/// Runs inside hyperlinks and standalone runs are counted together (matching
/// how `ParagraphExt::runs()` flattens them).
fn hyperlink_url_for_run(doc: &WmlDoc, para: &types::Paragraph, run_idx: usize) -> Option<String> {
    let mut counter = 0usize;
    for content in &para.paragraph_content {
        match content {
            types::ParagraphContent::R(_) => {
                if counter == run_idx {
                    return None;
                }
                counter += 1;
            }
            types::ParagraphContent::Hyperlink(hl) => {
                let hl_run_count = hl.runs().len();
                if counter + hl_run_count > run_idx {
                    return hl
                        .rel_id()
                        .and_then(|id| doc.get_hyperlink_url(id))
                        .map(|s| s.to_string());
                }
                counter += hl_run_count;
            }
            _ => {}
        }
    }
    None
}

/// Count total embedded images across all runs in the document body.
fn count_images(body: &types::Body) -> u32 {
    use ooxml_wml::ext::DrawingExt;
    let mut count = 0u32;
    for para in body.paragraphs() {
        for run in para.runs() {
            for drawing in run.drawings() {
                count += drawing.all_image_rel_ids().len() as u32;
            }
        }
    }
    count
}

/// Collect bookmark names from all paragraphs in the body (sorted).
fn collect_bookmark_names(body: &types::Body) -> Vec<String> {
    let mut names = BTreeSet::new();
    for para in body.paragraphs() {
        for content in &para.paragraph_content {
            if let types::ParagraphContent::BookmarkStart(bm) = content {
                names.insert(bm.name.clone());
            }
        }
    }
    names.into_iter().collect()
}

/// Compute the rowspan of a table cell at (row_idx, col_idx) by counting
/// how many consecutive rows have a `vMerge` continuation at the same column.
fn cell_rowspan(table: &types::Table, row_idx: usize, col_idx: usize) -> u32 {
    use ooxml_wml::types::STMerge;
    let rows = table.rows();
    let Some(start_row) = rows.get(row_idx) else {
        return 1;
    };
    let Some(cell) = start_row.cells().into_iter().nth(col_idx) else {
        return 1;
    };
    // Only count if this cell starts a rowspan.
    let is_restart = cell
        .properties()
        .and_then(|p| p.vertical_merge.as_ref())
        .is_some_and(|vm| {
            vm.value
                .as_ref()
                .is_none_or(|v| matches!(v, STMerge::Restart))
        });
    if !is_restart {
        return 1;
    }
    let mut span = 1u32;
    for r in &rows[row_idx + 1..] {
        let Some(c) = r.cells().into_iter().nth(col_idx) else {
            break;
        };
        let is_continue = c
            .properties()
            .and_then(|p| p.vertical_merge.as_ref())
            .is_some_and(|vm| matches!(vm.value.as_ref(), None | Some(STMerge::Continue)));
        if is_continue {
            span += 1;
        } else {
            break;
        }
    }
    span
}

// ---------------------------------------------------------------------------
// Assertion dispatch
// ---------------------------------------------------------------------------

fn check_assertion(doc: &WmlDoc, body: &types::Body, a: &Value) -> Result<(), String> {
    let t = a["type"].as_str().unwrap_or("unknown");

    match t {
        // ---- paragraph-level ------------------------------------------------
        "paragraph_count" => {
            let expected = a["expected"].as_u64().unwrap() as u32;
            let got = body.paragraphs().len() as u32;
            if got != expected {
                return Err(format!("paragraph_count: expected {expected}, got {got}"));
            }
        }

        "paragraph_text" => {
            let idx = a["para"].as_u64().unwrap() as usize;
            let expected = a["expected"].as_str().unwrap();
            let got = body
                .paragraphs()
                .into_iter()
                .nth(idx)
                .map(|p| p.text())
                .unwrap_or_default();
            if got != expected {
                return Err(format!(
                    "paragraph_text[{idx}]: expected {expected:?}, got {got:?}"
                ));
            }
        }

        "paragraph_style" => {
            let idx = a["para"].as_u64().unwrap() as usize;
            let expected = a["expected"].as_str().unwrap();
            let got = body
                .paragraphs()
                .into_iter()
                .nth(idx)
                .and_then(|p| p.properties())
                .and_then(|pr| pr.paragraph_style.as_ref())
                .map(|s| s.value.as_str())
                .unwrap_or("");
            if got != expected {
                return Err(format!(
                    "paragraph_style[{idx}]: expected {expected:?}, got {got:?}"
                ));
            }
        }

        "paragraph_align" => {
            let idx = a["para"].as_u64().unwrap() as usize;
            let expected = a["expected"].as_str().unwrap();
            let got = body
                .paragraphs()
                .into_iter()
                .nth(idx)
                .and_then(|p| p.properties())
                .and_then(|pr| pr.justification.as_ref())
                .map(|j| j.value.to_string())
                .unwrap_or_default();
            if got != expected {
                return Err(format!(
                    "paragraph_align[{idx}]: expected {expected:?}, got {got:?}"
                ));
            }
        }

        "paragraph_list_level" => {
            let idx = a["para"].as_u64().unwrap() as usize;
            let expected: Option<u32> = a["expected"].as_u64().map(|v| v as u32);
            let got: Option<u32> = body
                .paragraphs()
                .into_iter()
                .nth(idx)
                .and_then(|p| p.properties())
                .and_then(|pr| pr.num_pr.as_ref())
                .and_then(|n| n.ilvl.as_ref())
                .map(|l| l.value as u32);
            if got != expected {
                return Err(format!(
                    "paragraph_list_level[{idx}]: expected {expected:?}, got {got:?}"
                ));
            }
        }

        // ---- run-level -------------------------------------------------------
        "run_text" => {
            let pi = a["para"].as_u64().unwrap() as usize;
            let ri = a["run"].as_u64().unwrap() as usize;
            let expected = a["expected"].as_str().unwrap();
            let got = body
                .paragraphs()
                .into_iter()
                .nth(pi)
                .and_then(|p| p.runs().into_iter().nth(ri))
                .map(|r| r.text())
                .unwrap_or_default();
            if got != expected {
                return Err(format!(
                    "run_text[{pi}][{ri}]: expected {expected:?}, got {got:?}"
                ));
            }
        }

        "run_bold" => {
            let pi = a["para"].as_u64().unwrap() as usize;
            let ri = a["run"].as_u64().unwrap() as usize;
            let expected = a["expected"].as_bool().unwrap();
            let got = body
                .paragraphs()
                .into_iter()
                .nth(pi)
                .and_then(|p| p.runs().into_iter().nth(ri))
                .map(|r| r.is_bold())
                .unwrap_or(false);
            if got != expected {
                return Err(format!(
                    "run_bold[{pi}][{ri}]: expected {expected}, got {got}"
                ));
            }
        }

        "run_italic" => {
            let pi = a["para"].as_u64().unwrap() as usize;
            let ri = a["run"].as_u64().unwrap() as usize;
            let expected = a["expected"].as_bool().unwrap();
            let got = body
                .paragraphs()
                .into_iter()
                .nth(pi)
                .and_then(|p| p.runs().into_iter().nth(ri))
                .map(|r| r.is_italic())
                .unwrap_or(false);
            if got != expected {
                return Err(format!(
                    "run_italic[{pi}][{ri}]: expected {expected}, got {got}"
                ));
            }
        }

        "run_underline" => {
            let pi = a["para"].as_u64().unwrap() as usize;
            let ri = a["run"].as_u64().unwrap() as usize;
            let expected = a["expected"].as_bool().unwrap();
            let got = body
                .paragraphs()
                .into_iter()
                .nth(pi)
                .and_then(|p| p.runs().into_iter().nth(ri))
                .map(|r| r.is_underline())
                .unwrap_or(false);
            if got != expected {
                return Err(format!(
                    "run_underline[{pi}][{ri}]: expected {expected}, got {got}"
                ));
            }
        }

        "run_strikethrough" => {
            let pi = a["para"].as_u64().unwrap() as usize;
            let ri = a["run"].as_u64().unwrap() as usize;
            let expected = a["expected"].as_bool().unwrap();
            let got = body
                .paragraphs()
                .into_iter()
                .nth(pi)
                .and_then(|p| p.runs().into_iter().nth(ri))
                .map(|r| r.is_strikethrough())
                .unwrap_or(false);
            if got != expected {
                return Err(format!(
                    "run_strikethrough[{pi}][{ri}]: expected {expected}, got {got}"
                ));
            }
        }

        "run_color" => {
            let pi = a["para"].as_u64().unwrap() as usize;
            let ri = a["run"].as_u64().unwrap() as usize;
            let expected: Option<&str> = a["expected"].as_str();
            let got: Option<String> = body
                .paragraphs()
                .into_iter()
                .nth(pi)
                .and_then(|p| p.runs().into_iter().nth(ri))
                .and_then(|r| r.properties())
                .and_then(|pr| pr.color_hex())
                .map(|s| s.to_uppercase());
            let got_ref = got.as_deref();
            if got_ref != expected {
                return Err(format!(
                    "run_color[{pi}][{ri}]: expected {expected:?}, got {got_ref:?}"
                ));
            }
        }

        "run_font_size" => {
            let pi = a["para"].as_u64().unwrap() as usize;
            let ri = a["run"].as_u64().unwrap() as usize;
            let expected = a["expected"].as_f64().unwrap();
            let tolerance = a["tolerance"].as_f64().unwrap_or(0.0);
            let got = body
                .paragraphs()
                .into_iter()
                .nth(pi)
                .and_then(|p| p.runs().into_iter().nth(ri))
                .and_then(|r| r.properties())
                .and_then(|pr| pr.font_size_points());
            match got {
                None => {
                    return Err(format!(
                        "run_font_size[{pi}][{ri}]: expected {expected}, got None"
                    ));
                }
                Some(v) if (v - expected).abs() > tolerance => {
                    return Err(format!(
                        "run_font_size[{pi}][{ri}]: expected {expected} ±{tolerance}, got {v}"
                    ));
                }
                _ => {}
            }
        }

        "run_font_name" => {
            let pi = a["para"].as_u64().unwrap() as usize;
            let ri = a["run"].as_u64().unwrap() as usize;
            let expected: Option<&str> = a["expected"].as_str();
            let got: Option<String> = body
                .paragraphs()
                .into_iter()
                .nth(pi)
                .and_then(|p| p.runs().into_iter().nth(ri))
                .and_then(|r| r.properties())
                .and_then(|pr| pr.font_ascii())
                .map(|s| s.to_string());
            let got_ref = got.as_deref();
            if got_ref != expected {
                return Err(format!(
                    "run_font_name[{pi}][{ri}]: expected {expected:?}, got {got_ref:?}"
                ));
            }
        }

        // ---- table-level -----------------------------------------------------
        "table_rows" => {
            let ti = a["table"].as_u64().unwrap() as usize;
            let expected = a["expected"].as_u64().unwrap() as u32;
            let got = body
                .tables()
                .into_iter()
                .nth(ti)
                .map(|tbl| tbl.row_count() as u32)
                .unwrap_or(0);
            if got != expected {
                return Err(format!("table_rows[{ti}]: expected {expected}, got {got}"));
            }
        }

        "table_cols" => {
            let ti = a["table"].as_u64().unwrap() as usize;
            let ri = a["row"].as_u64().unwrap() as usize;
            let expected = a["expected"].as_u64().unwrap() as u32;
            let got = body
                .tables()
                .into_iter()
                .nth(ti)
                .and_then(|tbl| tbl.rows().into_iter().nth(ri))
                .map(|row| row.cells().len() as u32)
                .unwrap_or(0);
            if got != expected {
                return Err(format!(
                    "table_cols[{ti}][{ri}]: expected {expected}, got {got}"
                ));
            }
        }

        "table_cell_text" => {
            let ti = a["table"].as_u64().unwrap() as usize;
            let ri = a["row"].as_u64().unwrap() as usize;
            let ci = a["col"].as_u64().unwrap() as usize;
            let expected = a["expected"].as_str().unwrap();
            let got = body
                .tables()
                .into_iter()
                .nth(ti)
                .and_then(|tbl| tbl.rows().into_iter().nth(ri))
                .and_then(|row| row.cells().into_iter().nth(ci))
                .map(|cell| cell.text())
                .unwrap_or_default();
            if got != expected {
                return Err(format!(
                    "table_cell_text[{ti}][{ri}][{ci}]: expected {expected:?}, got {got:?}"
                ));
            }
        }

        "table_cell_colspan" => {
            let ti = a["table"].as_u64().unwrap() as usize;
            let ri = a["row"].as_u64().unwrap() as usize;
            let ci = a["col"].as_u64().unwrap() as usize;
            let expected = a["expected"].as_u64().unwrap() as u32;
            let got = body
                .tables()
                .into_iter()
                .nth(ti)
                .and_then(|tbl| tbl.rows().into_iter().nth(ri))
                .and_then(|row| row.cells().into_iter().nth(ci))
                .and_then(|cell| cell.properties())
                .and_then(|pr| pr.grid_span.as_ref())
                .map(|g| g.value as u32)
                .unwrap_or(1);
            if got != expected {
                return Err(format!(
                    "table_cell_colspan[{ti}][{ri}][{ci}]: expected {expected}, got {got}"
                ));
            }
        }

        "table_cell_rowspan" => {
            let ti = a["table"].as_u64().unwrap() as usize;
            let ri = a["row"].as_u64().unwrap() as usize;
            let ci = a["col"].as_u64().unwrap() as usize;
            let expected = a["expected"].as_u64().unwrap() as u32;
            let tables = body.tables();
            let got = tables
                .into_iter()
                .nth(ti)
                .map(|tbl| cell_rowspan(tbl, ri, ci))
                .unwrap_or(1);
            if got != expected {
                return Err(format!(
                    "table_cell_rowspan[{ti}][{ri}][{ci}]: expected {expected}, got {got}"
                ));
            }
        }

        // ---- misc -------------------------------------------------------------
        "hyperlink_url" => {
            let pi = a["para"].as_u64().unwrap() as usize;
            let ri = a["run"].as_u64().unwrap() as usize;
            let expected: Option<&str> = a["expected"].as_str();
            let para = body.paragraphs().into_iter().nth(pi);
            let got: Option<String> = para.and_then(|p| hyperlink_url_for_run(doc, p, ri));
            let got_ref = got.as_deref();
            if got_ref != expected {
                return Err(format!(
                    "hyperlink_url[{pi}][{ri}]: expected {expected:?}, got {got_ref:?}"
                ));
            }
        }

        "image_count" => {
            let expected = a["expected"].as_u64().unwrap() as u32;
            let got = count_images(body);
            if got != expected {
                return Err(format!("image_count: expected {expected}, got {got}"));
            }
        }

        "bookmark_names" => {
            let expected: Vec<String> = a["expected"]
                .as_array()
                .unwrap()
                .iter()
                .filter_map(|v| v.as_str().map(|s| s.to_string()))
                .collect();
            let got = collect_bookmark_names(body);
            if got != expected {
                return Err(format!(
                    "bookmark_names: expected {expected:?}, got {got:?}"
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
fn roundtrip_wml() {
    let dir = fixtures_dir();
    let json_files = find_json_files(&dir);
    assert!(
        !json_files.is_empty(),
        "no WML fixture JSON files found in {}",
        dir.display()
    );

    let mut failures: Vec<String> = Vec::new();

    for json_path in json_files {
        // Derive the OOXML file path by stripping the trailing ".json".
        let docx_path = json_path.with_extension(""); // strips ".json" → ".docx"

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

        let docx_bytes = match fs::read(&docx_path) {
            Ok(b) => b,
            Err(e) => {
                failures.push(format!("[{}] read docx: {e}", docx_path.display()));
                continue;
            }
        };

        let doc = match Document::from_reader(Cursor::new(docx_bytes)) {
            Ok(d) => d,
            Err(e) => {
                failures.push(format!("[{}] open document: {e}", docx_path.display()));
                continue;
            }
        };
        let body = doc.body();

        let assertions = manifest["assertions"].as_array().unwrap();
        for assertion in assertions {
            if let Err(msg) = check_assertion(&doc, body, assertion) {
                failures.push(format!("[{}] {msg}", docx_path.display()));
            }
        }
    }

    if !failures.is_empty() {
        panic!(
            "{} WML roundtrip failure(s):\n{}",
            failures.len(),
            failures.join("\n")
        );
    }
}
