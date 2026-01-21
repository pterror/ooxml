//! Example: Reading an Excel spreadsheet
//!
//! This example demonstrates how to open and read a .xlsx file.
//!
//! Run with: cargo run --example read_xlsx -- path/to/spreadsheet.xlsx

use ooxml_sml::Workbook;
use std::env;

fn main() -> ooxml_sml::Result<()> {
    let args: Vec<String> = env::args().collect();
    if args.len() < 2 {
        eprintln!("Usage: {} <spreadsheet.xlsx>", args[0]);
        std::process::exit(1);
    }

    let path = &args[1];
    println!("Opening: {}", path);

    let mut workbook = Workbook::open(path)?;

    // Print workbook info
    println!("\n=== Workbook Info ===");
    println!("Sheet count: {}", workbook.sheet_count());
    println!("Sheet names: {:?}", workbook.sheet_names());

    // Iterate through sheets
    for sheet in workbook.sheets()? {
        println!("\n=== Sheet: {} ===", sheet.name());

        // Print dimensions if available
        if let Some((min_row, min_col, max_row, max_col)) = sheet.dimensions() {
            println!(
                "Dimensions: rows {}..{}, cols {}..{}",
                min_row, max_row, min_col, max_col
            );
        }

        // Print all rows
        for row in sheet.rows() {
            print!("Row {:>3}: ", row.row_num());
            for cell in row.cells() {
                let value = cell.value_as_string();
                if !value.is_empty() {
                    print!("{}\t", value);
                }
            }
            println!();
        }
    }

    Ok(())
}
