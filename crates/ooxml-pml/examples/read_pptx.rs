//! Example: Reading a PowerPoint presentation
//!
//! This example demonstrates how to open and read a .pptx file.
//!
//! Run with: cargo run --example read_pptx -- path/to/presentation.pptx

use ooxml_pml::Presentation;
use std::env;

fn main() -> ooxml_pml::Result<()> {
    let args: Vec<String> = env::args().collect();
    if args.len() < 2 {
        eprintln!("Usage: {} <presentation.pptx>", args[0]);
        std::process::exit(1);
    }

    let path = &args[1];
    println!("Opening: {}", path);

    let mut pres = Presentation::open(path)?;

    // Print presentation info
    println!("\n=== Presentation Info ===");
    println!("Slide count: {}", pres.slide_count());

    // Iterate through slides
    for slide in pres.slides()? {
        println!("\n=== Slide {} ===", slide.index() + 1);

        // Print all shapes with text
        for shape in slide.shapes() {
            if let Some(name) = shape.name() {
                print!("[{}] ", name);
            }
            if let Some(text) = shape.text() {
                println!("{}", text);
            } else if shape.name().is_some() {
                println!("(no text)");
            }
        }

        // Also show combined slide text
        let all_text = slide.text();
        if !all_text.is_empty() {
            println!("\n--- Full slide text ---");
            println!("{}", all_text);
        }

        // Show speaker notes if present
        if let Some(notes) = slide.notes() {
            println!("\n--- Speaker Notes ---");
            println!("{}", notes);
        }

        // Show pictures
        let pictures = slide.pictures();
        if !pictures.is_empty() {
            println!("\n--- Pictures ({}) ---", pictures.len());
            for pic in pictures {
                print!("  {}", pic.rel_id());
                if let Some(name) = pic.name() {
                    print!(" ({})", name);
                }
                if let Some(descr) = pic.description() {
                    print!(" - {}", descr);
                }
                println!();
            }
        }

        // Show transition if present
        if let Some(trans) = slide.transition() {
            println!("\n--- Transition ---");
            if let Some(ref tt) = trans.transition_type {
                println!("  Type: {:?}", tt);
            }
            println!("  Speed: {:?}", trans.speed);
            println!("  Advance on click: {}", trans.advance_on_click);
            if let Some(ms) = trans.advance_time_ms {
                println!("  Auto-advance: {}ms", ms);
            }
        }
    }

    Ok(())
}
