//! Example: Creating an Excel spreadsheet
//!
//! This example demonstrates how to create a new .xlsx file.
//!
//! Run with: cargo run --example create_xlsx

use ooxml_sml::{CellResolveExt, RowExt, WorkbookBuilder};

fn main() -> ooxml_sml::Result<()> {
    let mut wb = WorkbookBuilder::new();

    // Create first sheet with sample data
    let sheet1 = wb.add_sheet("Sales Data");

    // Header row
    sheet1.set_cell("A1", "Product");
    sheet1.set_cell("B1", "Quantity");
    sheet1.set_cell("C1", "Price");
    sheet1.set_cell("D1", "Total");

    // Data rows
    sheet1.set_cell("A2", "Widget A");
    sheet1.set_cell("B2", 10.0);
    sheet1.set_cell("C2", 5.99);
    sheet1.set_formula("D2", "B2*C2");

    sheet1.set_cell("A3", "Widget B");
    sheet1.set_cell("B3", 25.0);
    sheet1.set_cell("C3", 12.50);
    sheet1.set_formula("D3", "B3*C3");

    sheet1.set_cell("A4", "Widget C");
    sheet1.set_cell("B4", 5.0);
    sheet1.set_cell("C4", 99.99);
    sheet1.set_formula("D4", "B4*C4");

    // Summary row
    sheet1.set_cell("A6", "Grand Total");
    sheet1.set_formula("D6", "SUM(D2:D4)");

    // Create second sheet with different data types
    let sheet2 = wb.add_sheet("Data Types");
    sheet2.set_cell("A1", "Type");
    sheet2.set_cell("B1", "Value");

    sheet2.set_cell("A2", "String");
    sheet2.set_cell("B2", "Hello, World!");

    sheet2.set_cell("A3", "Number");
    sheet2.set_cell("B3", 42.0);

    sheet2.set_cell("A4", "Boolean (true)");
    sheet2.set_cell("B4", true);

    sheet2.set_cell("A5", "Boolean (false)");
    sheet2.set_cell("B5", false);

    sheet2.set_cell("A6", "Formula");
    sheet2.set_formula("B6", "1+1");

    // Save the workbook
    let output_path = "sample_output.xlsx";
    wb.save(output_path)?;
    println!("Created: {}", output_path);

    // Verify by reading it back
    let mut workbook = ooxml_sml::Workbook::open(output_path)?;
    println!("\nVerifying created file:");
    println!("Sheet count: {}", workbook.sheet_count());
    println!("Sheet names: {:?}", workbook.sheet_names());

    for sheet in workbook.resolved_sheets()? {
        println!("\n=== {} ===", sheet.name());
        for row in sheet.rows() {
            print!("Row {:>2}: ", row.row_number().unwrap_or(0));
            for cell in row.cells_iter() {
                let value = cell.value_as_string(sheet.context());
                if !value.is_empty() {
                    print!("{}  ", value);
                }
            }
            println!();
        }
    }

    // Clean up
    std::fs::remove_file(output_path)?;
    println!("\nCleaned up test file.");

    Ok(())
}
