//! Example: Creating a PowerPoint presentation
//!
//! This example demonstrates how to create a new .pptx file with
//! various content types: titles, text, and tables.
//!
//! Run with: cargo run --example create_pptx

use ooxml_pml::{PresentationBuilder, TableBuilder, TableCellExt};

fn main() -> ooxml_pml::Result<()> {
    let mut pres = PresentationBuilder::new();

    // Slide 1: Title slide
    let slide1 = pres.add_slide();
    slide1.add_title("Welcome to OOXML");
    slide1.add_text("Creating PowerPoint files with Rust");

    // Slide 2: More text content
    let slide2 = pres.add_slide();
    slide2.add_title("Key Features");
    slide2.add_text("• Read existing presentations");
    slide2.add_text("• Write new presentations");
    slide2.add_text("• Tables with styled content");

    // Slide 3: Table
    let slide3 = pres.add_slide();
    slide3.add_title("Sample Data");

    let table = TableBuilder::new()
        .name("Quarterly Sales")
        .add_row(["Quarter", "Revenue", "Growth"])
        .add_row(["Q1", "$1.2M", "+5%"])
        .add_row(["Q2", "$1.4M", "+17%"])
        .add_row(["Q3", "$1.3M", "-7%"])
        .add_row(["Q4", "$1.8M", "+38%"]);

    // Position: x=1", y=2", width=8", height=2" (in EMUs: 914400 per inch)
    slide3.add_table(table, 914400, 1828800, 7315200, 1828800);

    // Slide 4: Speaker notes
    let slide4 = pres.add_slide();
    slide4.add_title("Summary");
    slide4.add_text("Thank you for using ooxml-pml!");
    slide4.set_notes("Remember to mention the GitHub repository.");

    // Save the presentation
    let output_path = "sample_presentation.pptx";
    pres.save(output_path)?;
    println!("Created: {}", output_path);

    // Verify by reading it back
    let mut presentation = ooxml_pml::Presentation::open(output_path)?;
    println!("\nVerifying created file:");
    println!("Slide count: {}", presentation.slide_count());

    for slide in presentation.slides()? {
        println!("\n=== Slide {} ===", slide.index() + 1);

        let text = slide.text();
        if !text.is_empty() {
            println!("{}", text);
        }

        // Show speaker notes if present
        if let Some(notes) = slide.notes() {
            println!("[Notes: {}]", notes);
        }

        // Show tables if present
        if slide.has_tables() {
            println!("\nTables: {}", slide.table_count());
            for (i, table) in slide.tables().iter().enumerate() {
                if let Some(name) = table.name() {
                    println!("  Table {}: \"{}\"", i + 1, name);
                }
                println!("  Dimensions: {}x{}", table.row_count(), table.col_count());

                // Print first row as preview
                if let Some(first_cell) = table.cell(0, 0) {
                    println!("  First cell: {}", first_cell.text());
                }
            }
        }
    }

    // Clean up
    std::fs::remove_file(output_path)?;
    println!("\nCleaned up test file.");

    Ok(())
}
