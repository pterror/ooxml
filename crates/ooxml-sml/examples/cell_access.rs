//! Example: Accessing specific cells in an Excel spreadsheet
//!
//! This example demonstrates different ways to access cell data.
//!
//! Run with: cargo run --example cell_access -- path/to/spreadsheet.xlsx

use ooxml_sml::{Workbook, WorkbookCellValue as CellValue};
use std::env;

fn main() -> ooxml_sml::Result<()> {
    let args: Vec<String> = env::args().collect();
    if args.len() < 2 {
        eprintln!("Usage: {} <spreadsheet.xlsx>", args[0]);
        std::process::exit(1);
    }

    let path = &args[1];
    let mut workbook = Workbook::open(path)?;

    // Get the first sheet
    let sheet = workbook.sheet(0)?;
    println!("Sheet: {}", sheet.name());

    // Access cell by reference (e.g., "A1")
    if let Some(cell) = sheet.cell("A1") {
        println!("\nCell A1:");
        println!("  Reference: {}", cell.reference());
        println!("  Value: {:?}", cell.value());
        println!("  As string: {}", cell.value_as_string());

        // Check if there's a formula
        if let Some(formula) = cell.formula() {
            println!("  Formula: {}", formula);
        }
    }

    // Access cell by row and column
    if let Some(row) = sheet.row(1) {
        println!("\nRow 1 cells:");
        for cell in row.cells() {
            // Demonstrate type-specific access
            match cell.value() {
                CellValue::Number(n) => {
                    println!("  {}: Number = {}", cell.reference(), n);
                }
                CellValue::String(s) => {
                    println!("  {}: String = \"{}\"", cell.reference(), s);
                }
                CellValue::Boolean(b) => {
                    println!("  {}: Boolean = {}", cell.reference(), b);
                }
                CellValue::Error(e) => {
                    println!("  {}: Error = {}", cell.reference(), e);
                }
                CellValue::Empty => {
                    println!("  {}: Empty", cell.reference());
                }
            }
        }
    }

    // Access cells using column letters
    if let Some(row) = sheet.row(2)
        && let Some(cell) = row.cell("B")
    {
        println!("\nCell B2: {}", cell.value_as_string());
    }

    // Demonstrate numeric conversion
    println!("\nNumeric values in column A:");
    for row in sheet.rows() {
        if let Some(cell) = row.cell("A")
            && let Some(num) = cell.value_as_number()
        {
            println!("  Row {}: {}", row.row_num(), num);
        }
    }

    Ok(())
}
