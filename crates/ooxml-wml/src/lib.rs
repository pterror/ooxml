//! WordprocessingML (DOCX) support for the ooxml library.
//!
//! This crate provides reading and writing of Word documents (.docx files).
//!
//! # Reading Documents
//!
//! ```no_run
//! use ooxml_wml::Document;
//!
//! let doc = Document::open("input.docx")?;
//! for para in doc.body().paragraphs() {
//!     println!("{}", para.text());
//! }
//! # Ok::<(), ooxml_wml::Error>(())
//! ```
//!
//! # Creating Documents
//!
//! ```no_run
//! use ooxml_wml::DocumentBuilder;
//!
//! let mut builder = DocumentBuilder::new();
//! builder.add_paragraph("Hello, World!");
//! builder.save("output.docx")?;
//! # Ok::<(), ooxml_wml::Error>(())
//! ```

pub mod document;
pub mod error;
pub mod styles;
pub mod writer;

// Generated types from ECMA-376 schema.
// Access via `ooxml_wml::types::*` for generated structs/enums.
pub mod types {
    #![allow(dead_code)]
    include!(concat!(env!("OUT_DIR"), "/wml_types.rs"));
}

pub use document::{
    BlockContent, Body, Cell, Document, Paragraph, ParagraphProperties, Row, Run, RunProperties,
    Table,
};
pub use error::{Error, Result};
pub use styles::{Style, StyleType, Styles};
pub use writer::DocumentBuilder;

// Re-export commonly used generated types at the crate root
pub use types::ns;
// TODO: Lists/numbering
// TODO: Images
// TODO: Hyperlinks
