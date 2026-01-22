//! Example: Using the ResolvedSheet API (recommended)
//!
//! This example demonstrates the preferred way to access Excel data using
//! the spec-compliant generated types.
//!
//! Run with: cargo run --example resolved_sheet -- path/to/spreadsheet.xlsx

use ooxml_sml::{CellExt, CellResolveExt, RowExt, Workbook, WorksheetExt};
use std::env;

fn main() -> ooxml_sml::Result<()> {
    let args: Vec<String> = env::args().collect();
    if args.len() < 2 {
        eprintln!("Usage: {} <spreadsheet.xlsx>", args[0]);
        std::process::exit(1);
    }

    let path = &args[1];
    let mut workbook = Workbook::open(path)?;

    println!("Workbook has {} sheets", workbook.sheet_count());
    println!("Sheet names: {:?}", workbook.sheet_names());

    // Get the first sheet using the new resolved_sheet API
    let sheet = workbook.resolved_sheet(0)?;
    println!("\nSheet: {} ({} rows)", sheet.name(), sheet.row_count());

    // Access cells with automatic value resolution
    if let Some(cell) = sheet.cell("A1") {
        let value = sheet.cell_value(cell);
        println!("\nCell A1: {:?}", value);

        // Use extension traits for detailed access
        if cell.has_formula() {
            println!("  Formula: {}", cell.formula_text().unwrap_or("(unknown)"));
        }
        println!("  Is shared string: {}", cell.is_shared_string());
    }

    // Iterate over rows and cells
    println!("\nAll data:");
    for row in sheet.rows() {
        print!("Row {}: ", row.row_number().unwrap_or(0));
        for cell in row.cells_iter() {
            // Get resolved value using context
            let value = cell.value_as_string(sheet.context());
            print!("{}\t", value);
        }
        println!();
    }

    // Access worksheet features
    if sheet.worksheet().has_auto_filter() {
        println!("\nSheet has auto-filter enabled");
    }
    if let Some(merged) = &sheet.worksheet().merged_cells {
        println!("Merged cell ranges: {}", merged.merge_cell.len());
    }
    if sheet.worksheet().has_conditional_formatting() {
        println!(
            "Conditional formatting rules: {}",
            sheet.worksheet().conditional_formatting.len()
        );
    }

    // Access comments and charts
    if !sheet.comments().is_empty() {
        println!("\nComments ({}):", sheet.comments().len());
        for comment in sheet.comments() {
            println!("  {}: {}", comment.reference, comment.text);
        }
    }

    if !sheet.charts().is_empty() {
        println!("\nCharts ({}):", sheet.charts().len());
        for chart in sheet.charts() {
            println!(
                "  {}: {:?}",
                chart.title.as_deref().unwrap_or("(untitled)"),
                chart.chart_type
            );
        }
    }

    Ok(())
}
