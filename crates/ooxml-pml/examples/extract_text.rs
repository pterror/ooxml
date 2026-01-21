//! Example: Extracting text from a PowerPoint presentation
//!
//! This example demonstrates extracting all text content from slides.
//!
//! Run with: cargo run --example extract_text -- path/to/presentation.pptx

use ooxml_pml::Presentation;
use std::env;

fn main() -> ooxml_pml::Result<()> {
    let args: Vec<String> = env::args().collect();
    if args.len() < 2 {
        eprintln!("Usage: {} <presentation.pptx>", args[0]);
        std::process::exit(1);
    }

    let path = &args[1];
    let mut pres = Presentation::open(path)?;

    // Extract text from each slide
    for slide in pres.slides()? {
        let text = slide.text();
        if !text.is_empty() {
            println!("=== Slide {} ===\n{}\n", slide.index() + 1, text);
        }

        // Include speaker notes
        if let Some(notes) = slide.notes() {
            println!("[Notes: {}]\n", notes);
        }
    }

    Ok(())
}
