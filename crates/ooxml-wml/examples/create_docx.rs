//! Example: Creating a Word document
//!
//! This example demonstrates how to create a new .docx file with
//! formatted text, lists, and tables.
//!
//! Run with: cargo run --example create_docx

use ooxml_wml::{DocumentBuilder, ListType};

fn main() -> ooxml_wml::Result<()> {
    let mut builder = DocumentBuilder::new();

    // Add a title with bold formatting
    {
        let para = builder.body_mut().add_paragraph();
        let run = para.add_run();
        run.set_text("Document Title");
        run.set_bold(true);
        run.set_font_size(48); // 24pt (size is in half-points)
    }

    // Add a regular paragraph
    builder.add_paragraph("This is a simple paragraph with regular text.");

    // Add a paragraph with italic text
    {
        let para = builder.body_mut().add_paragraph();
        let run = para.add_run();
        run.set_text("This text is italic.");
        run.set_italic(true);
    }

    // Add a bulleted list
    let bullet_id = builder.add_list(ListType::Bullet);
    for item in [
        "First bullet point",
        "Second bullet point",
        "Third bullet point",
    ] {
        let para = builder.body_mut().add_paragraph();
        para.add_run().set_text(item);
        para.set_numbering(bullet_id, 0);
    }

    // Add a numbered list
    let num_id = builder.add_list(ListType::Decimal);
    for item in [
        "First numbered item",
        "Second numbered item",
        "Third numbered item",
    ] {
        let para = builder.body_mut().add_paragraph();
        para.add_run().set_text(item);
        para.set_numbering(num_id, 0);
    }

    // Add a simple table
    {
        let table = builder.body_mut().add_table();

        // Header row
        let header_row = table.add_row();
        for header in ["Column 1", "Column 2", "Column 3"] {
            let cell = header_row.add_cell();
            let para = cell.add_paragraph();
            let run = para.add_run();
            run.set_text(header);
            run.set_bold(true);
        }

        // Data rows
        for row_data in [["A1", "B1", "C1"], ["A2", "B2", "C2"], ["A3", "B3", "C3"]] {
            let row = table.add_row();
            for cell_text in row_data {
                let cell = row.add_cell();
                cell.add_paragraph().add_run().set_text(cell_text);
            }
        }
    }

    // Save the document
    let output_path = "output.docx";
    builder.save(output_path)?;
    println!("Document saved to: {}", output_path);

    Ok(())
}
