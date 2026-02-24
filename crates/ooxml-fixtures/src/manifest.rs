//! Serialization of `Fixture` metadata to JSON manifests.
//!
//! Each fixture produces a companion `.json` file alongside the OOXML file.
//! The manifest records the fixture description and all assertions so that test
//! harnesses can verify the file without hard-coding expectations in test code.

use crate::{Assertion, Fixture};
use serde_json::{Map, Value, json};
use std::io;
use std::path::Path;

/// Serialize a single `Assertion` to a JSON object.
fn assertion_to_json(assertion: &Assertion) -> Value {
    match assertion {
        // WML — body
        Assertion::ParagraphCount { expected } => {
            json!({"type": "paragraph_count", "expected": expected})
        }
        Assertion::ParagraphText { para, expected } => {
            json!({"type": "paragraph_text", "para": para, "expected": expected})
        }
        Assertion::ParagraphStyle { para, expected } => {
            json!({"type": "paragraph_style", "para": para, "expected": expected})
        }
        Assertion::ParagraphAlign { para, expected } => {
            json!({"type": "paragraph_align", "para": para, "expected": expected})
        }
        Assertion::ParagraphListLevel { para, expected } => {
            json!({"type": "paragraph_list_level", "para": para, "expected": expected})
        }

        // WML — runs
        Assertion::RunText {
            para,
            run,
            expected,
        } => {
            json!({"type": "run_text", "para": para, "run": run, "expected": expected})
        }
        Assertion::RunBold {
            para,
            run,
            expected,
        } => {
            json!({"type": "run_bold", "para": para, "run": run, "expected": expected})
        }
        Assertion::RunItalic {
            para,
            run,
            expected,
        } => {
            json!({"type": "run_italic", "para": para, "run": run, "expected": expected})
        }
        Assertion::RunUnderline {
            para,
            run,
            expected,
        } => {
            json!({"type": "run_underline", "para": para, "run": run, "expected": expected})
        }
        Assertion::RunStrikethrough {
            para,
            run,
            expected,
        } => {
            json!({"type": "run_strikethrough", "para": para, "run": run, "expected": expected})
        }
        Assertion::RunColor {
            para,
            run,
            expected,
        } => {
            json!({"type": "run_color", "para": para, "run": run, "expected": expected})
        }
        Assertion::RunFontSize {
            para,
            run,
            expected,
            tolerance,
        } => {
            json!({"type": "run_font_size", "para": para, "run": run, "expected": expected, "tolerance": tolerance})
        }
        Assertion::RunFontName {
            para,
            run,
            expected,
        } => {
            json!({"type": "run_font_name", "para": para, "run": run, "expected": expected})
        }

        // WML — tables
        Assertion::TableRows { table, expected } => {
            json!({"type": "table_rows", "table": table, "expected": expected})
        }
        Assertion::TableCols {
            table,
            row,
            expected,
        } => {
            json!({"type": "table_cols", "table": table, "row": row, "expected": expected})
        }
        Assertion::TableCellText {
            table,
            row,
            col,
            expected,
        } => {
            json!({"type": "table_cell_text", "table": table, "row": row, "col": col, "expected": expected})
        }
        Assertion::TableCellColspan {
            table,
            row,
            col,
            expected,
        } => {
            json!({"type": "table_cell_colspan", "table": table, "row": row, "col": col, "expected": expected})
        }
        Assertion::TableCellRowspan {
            table,
            row,
            col,
            expected,
        } => {
            json!({"type": "table_cell_rowspan", "table": table, "row": row, "col": col, "expected": expected})
        }

        // WML — misc
        Assertion::HyperlinkUrl {
            para,
            run,
            expected,
        } => {
            json!({"type": "hyperlink_url", "para": para, "run": run, "expected": expected})
        }
        Assertion::ImageCount { expected } => {
            json!({"type": "image_count", "expected": expected})
        }
        Assertion::BookmarkNames { expected } => {
            json!({"type": "bookmark_names", "expected": expected})
        }

        // SML — workbook
        Assertion::SheetCount { expected } => {
            json!({"type": "sheet_count", "expected": expected})
        }
        Assertion::SheetName { sheet, expected } => {
            json!({"type": "sheet_name", "sheet": sheet, "expected": expected})
        }

        // SML — cells
        Assertion::CellType {
            sheet,
            row,
            col,
            expected,
        } => {
            json!({"type": "cell_type", "sheet": sheet, "row": row, "col": col, "expected": expected})
        }
        Assertion::CellValue {
            sheet,
            row,
            col,
            expected,
            tolerance,
        } => {
            json!({"type": "cell_value", "sheet": sheet, "row": row, "col": col, "expected": expected, "tolerance": tolerance})
        }
        Assertion::CellFormula {
            sheet,
            row,
            col,
            expected,
        } => {
            json!({"type": "cell_formula", "sheet": sheet, "row": row, "col": col, "expected": expected})
        }
        Assertion::CellFormatCode {
            sheet,
            row,
            col,
            expected,
        } => {
            json!({"type": "cell_format_code", "sheet": sheet, "row": row, "col": col, "expected": expected})
        }
        Assertion::CellBold {
            sheet,
            row,
            col,
            expected,
        } => {
            json!({"type": "cell_bold", "sheet": sheet, "row": row, "col": col, "expected": expected})
        }
        Assertion::CellItalic {
            sheet,
            row,
            col,
            expected,
        } => {
            json!({"type": "cell_italic", "sheet": sheet, "row": row, "col": col, "expected": expected})
        }
        Assertion::CellColor {
            sheet,
            row,
            col,
            expected,
        } => {
            json!({"type": "cell_color", "sheet": sheet, "row": row, "col": col, "expected": expected})
        }
        Assertion::MergedRegion {
            sheet,
            row,
            col,
            expected,
        } => {
            json!({"type": "merged_region", "sheet": sheet, "row": row, "col": col, "expected": expected})
        }
        Assertion::RowHeight {
            sheet,
            row,
            expected,
            tolerance,
        } => {
            json!({"type": "row_height", "sheet": sheet, "row": row, "expected": expected, "tolerance": tolerance})
        }
        Assertion::ColWidth {
            sheet,
            col,
            expected,
            tolerance,
        } => {
            json!({"type": "col_width", "sheet": sheet, "col": col, "expected": expected, "tolerance": tolerance})
        }

        // PML — presentation
        Assertion::SlideCount { expected } => {
            json!({"type": "slide_count", "expected": expected})
        }
        Assertion::ShapeCount { slide, expected } => {
            json!({"type": "shape_count", "slide": slide, "expected": expected})
        }
        Assertion::ShapeText {
            slide,
            shape,
            expected,
        } => {
            json!({"type": "shape_text", "slide": slide, "shape": shape, "expected": expected})
        }
        Assertion::ShapeType {
            slide,
            shape,
            expected,
        } => {
            json!({"type": "shape_type", "slide": slide, "shape": shape, "expected": expected})
        }

        // PML — runs
        Assertion::PmlRunText {
            slide,
            shape,
            para,
            run,
            expected,
        } => {
            json!({"type": "pml_run_text", "slide": slide, "shape": shape, "para": para, "run": run, "expected": expected})
        }
        Assertion::PmlRunBold {
            slide,
            shape,
            para,
            run,
            expected,
        } => {
            json!({"type": "pml_run_bold", "slide": slide, "shape": shape, "para": para, "run": run, "expected": expected})
        }
        Assertion::PmlRunItalic {
            slide,
            shape,
            para,
            run,
            expected,
        } => {
            json!({"type": "pml_run_italic", "slide": slide, "shape": shape, "para": para, "run": run, "expected": expected})
        }
        Assertion::PmlRunColor {
            slide,
            shape,
            para,
            run,
            expected,
        } => {
            json!({"type": "pml_run_color", "slide": slide, "shape": shape, "para": para, "run": run, "expected": expected})
        }
        Assertion::PmlRunFontSize {
            slide,
            shape,
            para,
            run,
            expected,
            tolerance,
        } => {
            json!({"type": "pml_run_font_size", "slide": slide, "shape": shape, "para": para, "run": run, "expected": expected, "tolerance": tolerance})
        }

        // PML — misc
        Assertion::SlideHasNotes { slide, expected } => {
            json!({"type": "slide_has_notes", "slide": slide, "expected": expected})
        }
        Assertion::NotesText { slide, expected } => {
            json!({"type": "notes_text", "slide": slide, "expected": expected})
        }
        Assertion::PmlImageCount { slide, expected } => {
            json!({"type": "pml_image_count", "slide": slide, "expected": expected})
        }
        Assertion::HasTransition { slide, expected } => {
            json!({"type": "has_transition", "slide": slide, "expected": expected})
        }
    }
}

/// Infer the OOXML format from the fixture file path extension.
fn format_from_path(path: &str) -> &'static str {
    if path.ends_with(".docx") {
        "wml"
    } else if path.ends_with(".xlsx") {
        "sml"
    } else if path.ends_with(".pptx") {
        "pml"
    } else {
        "unknown"
    }
}

/// Write a JSON manifest for `fixture` into `out_dir`.
///
/// The manifest is placed at `{out_dir}/{fixture.path}.json`.
/// All parent directories are created as needed.
pub fn write_manifest(fixture: &Fixture, out_dir: &Path) -> io::Result<()> {
    let manifest_path = out_dir.join(format!("{}.json", fixture.path));

    // Ensure parent directory exists.
    if let Some(parent) = manifest_path.parent() {
        std::fs::create_dir_all(parent)?;
    }

    let assertions: Vec<Value> = fixture.assertions.iter().map(assertion_to_json).collect();

    let mut doc = Map::new();
    doc.insert("version".into(), json!(1));
    doc.insert("format".into(), json!(format_from_path(fixture.path)));
    doc.insert("description".into(), json!(fixture.description));
    doc.insert("assertions".into(), Value::Array(assertions));

    let json_bytes = serde_json::to_vec_pretty(&Value::Object(doc)).map_err(io::Error::other)?;

    std::fs::write(&manifest_path, json_bytes)
}
