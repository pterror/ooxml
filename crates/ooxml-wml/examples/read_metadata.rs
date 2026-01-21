//! Example: Reading document metadata
//!
//! This example demonstrates how to read core and app properties
//! from a .docx file (title, author, word count, etc.).
//!
//! Run with: cargo run --example read_metadata -- path/to/document.docx

use ooxml_wml::Document;
use std::env;

fn main() -> ooxml_wml::Result<()> {
    let args: Vec<String> = env::args().collect();
    if args.len() < 2 {
        eprintln!("Usage: {} <document.docx>", args[0]);
        std::process::exit(1);
    }

    let path = &args[1];
    println!("Opening: {}\n", path);

    let doc = Document::open(path)?;

    // Print core properties (Dublin Core metadata)
    println!("=== Core Properties ===");
    if let Some(props) = doc.core_properties() {
        if let Some(title) = &props.title {
            println!("  Title:            {}", title);
        }
        if let Some(creator) = &props.creator {
            println!("  Author:           {}", creator);
        }
        if let Some(subject) = &props.subject {
            println!("  Subject:          {}", subject);
        }
        if let Some(description) = &props.description {
            println!("  Description:      {}", description);
        }
        if let Some(keywords) = &props.keywords {
            println!("  Keywords:         {}", keywords);
        }
        if let Some(category) = &props.category {
            println!("  Category:         {}", category);
        }
        if let Some(last_modified_by) = &props.last_modified_by {
            println!("  Last Modified By: {}", last_modified_by);
        }
        if let Some(revision) = &props.revision {
            println!("  Revision:         {}", revision);
        }
        if let Some(created) = &props.created {
            println!("  Created:          {}", created);
        }
        if let Some(modified) = &props.modified {
            println!("  Modified:         {}", modified);
        }
    } else {
        println!("  (No core properties found)");
    }

    // Print app properties (extended properties)
    println!("\n=== Application Properties ===");
    if let Some(props) = doc.app_properties() {
        if let Some(app) = &props.application {
            println!("  Application:      {}", app);
        }
        if let Some(version) = &props.app_version {
            println!("  App Version:      {}", version);
        }
        if let Some(company) = &props.company {
            println!("  Company:          {}", company);
        }
        if let Some(template) = &props.template {
            println!("  Template:         {}", template);
        }
        if let Some(pages) = props.pages {
            println!("  Pages:            {}", pages);
        }
        if let Some(words) = props.words {
            println!("  Words:            {}", words);
        }
        if let Some(characters) = props.characters {
            println!("  Characters:       {}", characters);
        }
        if let Some(paragraphs) = props.paragraphs {
            println!("  Paragraphs:       {}", paragraphs);
        }
        if let Some(lines) = props.lines {
            println!("  Lines:            {}", lines);
        }
        if let Some(total_time) = props.total_time {
            println!("  Editing Time:     {} minutes", total_time);
        }
    } else {
        println!("  (No app properties found)");
    }

    Ok(())
}
