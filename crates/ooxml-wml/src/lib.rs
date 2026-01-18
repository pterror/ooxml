//! WordprocessingML (DOCX) support for the ooxml library.
//!
//! This crate provides reading and writing of Word documents (.docx files).
//!
//! # Example
//!
//! ```ignore
//! use ooxml_wml::Document;
//!
//! // Open an existing document
//! let doc = Document::open("input.docx")?;
//! for para in doc.body().paragraphs() {
//!     println!("{}", para.text());
//! }
//!
//! // Create a new document
//! let mut doc = Document::new();
//! doc.body_mut().add_paragraph().add_run().set_text("Hello, World!");
//! doc.save("output.docx")?;
//! ```

// TODO: Document struct (main entry point)
// TODO: Body, Paragraph, Run, Text elements
// TODO: Formatting properties (bold, italic, etc.)
// TODO: Styles support
// TODO: Tables
// TODO: Lists/numbering
// TODO: Images
// TODO: Hyperlinks
