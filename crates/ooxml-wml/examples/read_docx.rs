//! Example: Reading a Word document
//!
//! This example demonstrates how to open and read a .docx file.
//!
//! Run with: cargo run --example read_docx -- path/to/document.docx

use ooxml_wml::Document;
use std::env;

fn main() -> ooxml_wml::Result<()> {
    let args: Vec<String> = env::args().collect();
    if args.len() < 2 {
        eprintln!("Usage: {} <document.docx>", args[0]);
        std::process::exit(1);
    }

    let path = &args[1];
    println!("Opening: {}", path);

    let doc = Document::open(path)?;

    // Print basic document info
    println!("\n=== Document Content ===\n");

    // Iterate through paragraphs
    for (i, para) in doc.body().paragraphs().iter().enumerate() {
        let text = para.text();
        if !text.is_empty() {
            println!("Paragraph {}: {}", i + 1, text);
        }
    }

    // Print full text
    println!("\n=== Full Text ===\n");
    println!("{}", doc.text());

    Ok(())
}
