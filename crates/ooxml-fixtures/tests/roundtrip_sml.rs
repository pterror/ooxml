//! Fixture roundtrip tests for SML (XLSX).
//!
//! For every `.json` manifest under `fixtures/sml/` this test:
//! 1. Opens the companion `.xlsx` file with the SML reader.
//! 2. Evaluates every assertion in the manifest.
//! 3. Collects all failures and reports them at the end.

use std::fs;
use std::io::Cursor;
use std::path::{Path, PathBuf};

use ooxml_sml::Workbook;
use ooxml_sml::ext::{CellExt, CellValue};
use ooxml_sml::types::Stylesheet;
use ooxml_sml::workbook::StylesheetExt;
use serde_json::Value;

type SmlWorkbook = Workbook<Cursor<Vec<u8>>>;

// ---------------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------------

fn fixtures_dir() -> PathBuf {
    let manifest_dir = env!("CARGO_MANIFEST_DIR");
    Path::new(manifest_dir).join("../../fixtures/sml")
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

/// Convert 0-indexed (row, col) to a cell reference like "A1".
fn idx_to_ref(row: usize, col: usize) -> String {
    let mut col_name = String::new();
    let mut c = col + 1; // convert to 1-based
    loop {
        col_name.insert(0, (b'A' + ((c - 1) % 26) as u8) as char);
        c = (c - 1) / 26;
        if c == 0 {
            break;
        }
    }
    format!("{}{}", col_name, row + 1)
}

/// Get the `CellValue` type as a string for the `cell_type` assertion.
fn cell_value_type_str(v: &CellValue) -> &'static str {
    match v {
        CellValue::Empty => "blank",
        CellValue::String(_) => "string",
        CellValue::Number(_) => "number",
        CellValue::Boolean(_) => "boolean",
        CellValue::Error(_) => "error",
    }
}

/// Get the font at `font_id` from the stylesheet (returns None if not found).
fn get_font(stylesheet: &Stylesheet, font_id: u32) -> Option<&ooxml_sml::types::Font> {
    stylesheet
        .fonts
        .as_ref()
        .and_then(|fonts| fonts.font.get(font_id as usize))
}

/// Get the Format (xf) at `style_index` from the stylesheet.
fn get_format(stylesheet: &Stylesheet, style_index: u32) -> Option<&ooxml_sml::types::Format> {
    stylesheet
        .cell_xfs
        .as_ref()
        .and_then(|xfs| xfs.xf.get(style_index as usize))
}

/// Get the Fill at `fill_id` from the stylesheet.
fn get_fill(stylesheet: &Stylesheet, fill_id: u32) -> Option<&ooxml_sml::types::Fill> {
    stylesheet
        .fills
        .as_ref()
        .and_then(|fills| fills.fill.get(fill_id as usize))
}

/// Convert a 4-byte ARGB `Vec<u8>` (bytes 1..4) to "RRGGBB".
fn argb_to_hex(rgb: &[u8]) -> Option<String> {
    if rgb.len() >= 4 {
        Some(format!("{:02X}{:02X}{:02X}", rgb[1], rgb[2], rgb[3]))
    } else if rgb.len() == 3 {
        Some(format!("{:02X}{:02X}{:02X}", rgb[0], rgb[1], rgb[2]))
    } else {
        None
    }
}

/// Check whether the cell at (row, col) is inside any merged region.
fn is_cell_merged(sheet: &ooxml_sml::ext::ResolvedSheet, row: usize, col: usize) -> bool {
    let cell_ref = idx_to_ref(row, col);
    let Some(merged_cells) = sheet.merged_cells() else {
        return false;
    };
    for mc in &merged_cells.merge_cell {
        if range_contains(&mc.reference, &cell_ref) {
            return true;
        }
    }
    false
}

/// Check if `cell_ref` (e.g. "A1") is within `range` (e.g. "A1:B2").
fn range_contains(range: &str, cell_ref: &str) -> bool {
    let Some((start, end)) = range.split_once(':') else {
        return range == cell_ref;
    };
    let (sc, sr) = split_cell_ref(start);
    let (ec, er) = split_cell_ref(end);
    let (cc, cr) = split_cell_ref(cell_ref);
    sc <= cc && cc <= ec && sr <= cr && cr <= er
}

/// Split a cell ref like "AB12" into ("AB", 12).
fn split_cell_ref(s: &str) -> (&str, u32) {
    let col_end = s.find(|c: char| c.is_ascii_digit()).unwrap_or(s.len());
    let col = &s[..col_end];
    let row: u32 = s[col_end..].parse().unwrap_or(0);
    (col, row)
}

// ---------------------------------------------------------------------------
// Assertion dispatch
// ---------------------------------------------------------------------------

fn check_assertion(
    workbook: &SmlWorkbook,
    stylesheet: &Option<Stylesheet>,
    sheet_cache: &mut [Option<ooxml_sml::ext::ResolvedSheet>],
    a: &Value,
) -> Result<(), String> {
    let t = a["type"].as_str().unwrap_or("unknown");

    match t {
        // ---- workbook-level --------------------------------------------------
        "sheet_count" => {
            let expected = a["expected"].as_u64().unwrap() as u32;
            let got = workbook.sheet_count() as u32;
            if got != expected {
                return Err(format!("sheet_count: expected {expected}, got {got}"));
            }
        }

        "sheet_name" => {
            let si = a["sheet"].as_u64().unwrap() as usize;
            let expected = a["expected"].as_str().unwrap();
            let names = workbook.sheet_names();
            let got = names.get(si).copied().unwrap_or("");
            if got != expected {
                return Err(format!(
                    "sheet_name[{si}]: expected {expected:?}, got {got:?}"
                ));
            }
        }

        // ---- cell-level ------------------------------------------------------
        "cell_type" | "cell_value" | "cell_formula" | "cell_format_code" | "cell_bold"
        | "cell_italic" | "cell_color" | "merged_region" => {
            let si = a["sheet"].as_u64().unwrap() as usize;
            let ri = a["row"].as_u64().unwrap() as usize;
            let ci = a["col"].as_u64().unwrap() as usize;
            let cell_ref = idx_to_ref(ri, ci);

            let sheet = match sheet_cache.get(si).and_then(|s| s.as_ref()) {
                Some(s) => s,
                None => return Err(format!("sheet {si} not loaded")),
            };

            match t {
                "cell_type" => {
                    let expected = a["expected"].as_str().unwrap();
                    let cell = sheet.cell(&cell_ref);
                    let value = cell
                        .map(|c| sheet.cell_value(c))
                        .unwrap_or(CellValue::Empty);
                    let got = cell_value_type_str(&value);
                    if got != expected {
                        return Err(format!(
                            "cell_type[{si}][{cell_ref}]: expected {expected:?}, got {got:?}"
                        ));
                    }
                }

                "cell_value" => {
                    let expected = a["expected"].as_str().unwrap();
                    let tolerance = a["tolerance"].as_f64().unwrap_or(0.0);
                    let cell = sheet.cell(&cell_ref);
                    let got_str = cell.map(|c| sheet.cell_value_string(c)).unwrap_or_default();
                    // Try numeric comparison when tolerance is set or expected looks numeric.
                    if let (Ok(exp_f), Some(got_f)) = (
                        expected.parse::<f64>(),
                        cell.and_then(|c| sheet.cell_value_number(c)),
                    ) {
                        if (got_f - exp_f).abs() > tolerance {
                            return Err(format!(
                                "cell_value[{si}][{cell_ref}]: expected {exp_f} ±{tolerance}, got {got_f}"
                            ));
                        }
                    } else if got_str != expected {
                        return Err(format!(
                            "cell_value[{si}][{cell_ref}]: expected {expected:?}, got {got_str:?}"
                        ));
                    }
                }

                "cell_formula" => {
                    let expected: Option<&str> = a["expected"].as_str();
                    let cell = sheet.cell(&cell_ref);
                    let got: Option<&str> = cell.and_then(|c| c.formula_text());
                    if got != expected {
                        return Err(format!(
                            "cell_formula[{si}][{cell_ref}]: expected {expected:?}, got {got:?}"
                        ));
                    }
                }

                "cell_format_code" => {
                    let expected: Option<&str> = a["expected"].as_str();
                    let cell = sheet.cell(&cell_ref);
                    let got: Option<String> = cell
                        .and_then(|c| c.style_index)
                        .and_then(|idx| stylesheet.as_ref().and_then(|ss| get_format(ss, idx)))
                        .and_then(|fmt| fmt.number_format_id)
                        .and_then(|id| {
                            stylesheet
                                .as_ref()
                                .and_then(|ss| ss.format_code(id))
                                .or_else(|| {
                                    ooxml_sml::workbook::builtin_format_code(id)
                                        .map(|s| s.to_string())
                                })
                        });
                    let got_ref = got.as_deref();
                    if got_ref != expected {
                        return Err(format!(
                            "cell_format_code[{si}][{cell_ref}]: expected {expected:?}, got {got_ref:?}"
                        ));
                    }
                }

                "cell_bold" => {
                    let expected = a["expected"].as_bool().unwrap();
                    let cell = sheet.cell(&cell_ref);
                    let got = cell
                        .and_then(|c| c.style_index)
                        .and_then(|idx| stylesheet.as_ref().and_then(|ss| get_format(ss, idx)))
                        .and_then(|fmt| fmt.font_id)
                        .and_then(|fid| stylesheet.as_ref().and_then(|ss| get_font(ss, fid)))
                        .and_then(|font| font.b.as_ref())
                        .map(|b| b.value.unwrap_or(true))
                        .unwrap_or(false);
                    if got != expected {
                        return Err(format!(
                            "cell_bold[{si}][{cell_ref}]: expected {expected}, got {got}"
                        ));
                    }
                }

                "cell_italic" => {
                    let expected = a["expected"].as_bool().unwrap();
                    let cell = sheet.cell(&cell_ref);
                    let got = cell
                        .and_then(|c| c.style_index)
                        .and_then(|idx| stylesheet.as_ref().and_then(|ss| get_format(ss, idx)))
                        .and_then(|fmt| fmt.font_id)
                        .and_then(|fid| stylesheet.as_ref().and_then(|ss| get_font(ss, fid)))
                        .and_then(|font| font.i.as_ref())
                        .map(|b| b.value.unwrap_or(true))
                        .unwrap_or(false);
                    if got != expected {
                        return Err(format!(
                            "cell_italic[{si}][{cell_ref}]: expected {expected}, got {got}"
                        ));
                    }
                }

                "cell_color" => {
                    let expected: Option<&str> = a["expected"].as_str();
                    let cell = sheet.cell(&cell_ref);
                    let got: Option<String> = cell
                        .and_then(|c| c.style_index)
                        .and_then(|idx| stylesheet.as_ref().and_then(|ss| get_format(ss, idx)))
                        .and_then(|fmt| fmt.fill_id)
                        .and_then(|fid| stylesheet.as_ref().and_then(|ss| get_fill(ss, fid)))
                        .and_then(|fill| fill.pattern_fill.as_ref())
                        .and_then(|pf| pf.fg_color.as_ref())
                        .and_then(|color| color.rgb.as_ref())
                        .and_then(|rgb| argb_to_hex(rgb));
                    let got_ref = got.as_deref();
                    if got_ref != expected {
                        return Err(format!(
                            "cell_color[{si}][{cell_ref}]: expected {expected:?}, got {got_ref:?}"
                        ));
                    }
                }

                "merged_region" => {
                    let expected = a["expected"].as_bool().unwrap();
                    let got = is_cell_merged(sheet, ri, ci);
                    if got != expected {
                        return Err(format!(
                            "merged_region[{si}][{cell_ref}]: expected {expected}, got {got}"
                        ));
                    }
                }

                _ => unreachable!(),
            }
        }

        "row_height" => {
            let si = a["sheet"].as_u64().unwrap() as usize;
            let ri = a["row"].as_u64().unwrap() as usize;
            let expected = a["expected"].as_f64().unwrap();
            let tolerance = a["tolerance"].as_f64().unwrap_or(0.0);

            let sheet = match sheet_cache.get(si).and_then(|s| s.as_ref()) {
                Some(s) => s,
                None => return Err(format!("sheet {si} not loaded")),
            };

            let row_num = (ri + 1) as u32;
            let got = sheet.row(row_num).and_then(|r| r.height);
            match got {
                None => {
                    return Err(format!(
                        "row_height[{si}][{ri}]: expected {expected}, got None"
                    ));
                }
                Some(v) if (v - expected).abs() > tolerance => {
                    return Err(format!(
                        "row_height[{si}][{ri}]: expected {expected} ±{tolerance}, got {v}"
                    ));
                }
                _ => {}
            }
        }

        "col_width" => {
            let si = a["sheet"].as_u64().unwrap() as usize;
            let ci = a["col"].as_u64().unwrap() as usize;
            let expected = a["expected"].as_f64().unwrap();
            let tolerance = a["tolerance"].as_f64().unwrap_or(0.0);

            let sheet = match sheet_cache.get(si).and_then(|s| s.as_ref()) {
                Some(s) => s,
                None => return Err(format!("sheet {si} not loaded")),
            };

            let col_1based = (ci + 1) as u32;
            let got: Option<f64> = sheet
                .columns()
                .iter()
                .flat_map(|cols| &cols.col)
                .find(|col| col.start_column <= col_1based && col_1based <= col.end_column)
                .and_then(|col| col.width);
            match got {
                None => {
                    return Err(format!(
                        "col_width[{si}][{ci}]: expected {expected}, got None"
                    ));
                }
                Some(v) if (v - expected).abs() > tolerance => {
                    return Err(format!(
                        "col_width[{si}][{ci}]: expected {expected} ±{tolerance}, got {v}"
                    ));
                }
                _ => {}
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
fn roundtrip_sml() {
    let dir = fixtures_dir();
    let json_files = find_json_files(&dir);
    assert!(
        !json_files.is_empty(),
        "no SML fixture JSON files found in {}",
        dir.display()
    );

    let mut failures: Vec<String> = Vec::new();

    for json_path in json_files {
        let xlsx_path = json_path.with_extension("");

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

        let xlsx_bytes = match fs::read(&xlsx_path) {
            Ok(b) => b,
            Err(e) => {
                failures.push(format!("[{}] read xlsx: {e}", xlsx_path.display()));
                continue;
            }
        };

        let mut workbook = match Workbook::from_reader(Cursor::new(xlsx_bytes)) {
            Ok(w) => w,
            Err(e) => {
                failures.push(format!("[{}] open workbook: {e}", xlsx_path.display()));
                continue;
            }
        };

        // Clone the stylesheet before mutably borrowing workbook for sheets.
        let stylesheet: Option<Stylesheet> = workbook.stylesheet().cloned();

        // Pre-load all sheets into a cache.
        let sheet_count = workbook.sheet_count();
        let mut sheet_cache: Vec<Option<ooxml_sml::ext::ResolvedSheet>> =
            Vec::with_capacity(sheet_count);
        for i in 0..sheet_count {
            match workbook.resolved_sheet(i) {
                Ok(s) => sheet_cache.push(Some(s)),
                Err(e) => {
                    failures.push(format!("[{}] load sheet {i}: {e}", xlsx_path.display()));
                    sheet_cache.push(None);
                }
            }
        }

        let assertions = manifest["assertions"].as_array().unwrap();
        for assertion in assertions {
            if let Err(msg) = check_assertion(&workbook, &stylesheet, &mut sheet_cache, assertion) {
                failures.push(format!("[{}] {msg}", xlsx_path.display()));
            }
        }
    }

    if !failures.is_empty() {
        panic!(
            "{} SML roundtrip failure(s):\n{}",
            failures.len(),
            failures.join("\n")
        );
    }
}
