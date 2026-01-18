//! WordprocessingML (DOCX) support for the ooxml library.
//!
//! This crate provides reading and writing of Word documents (.docx files).
//!
//! # Example
//!
//! ```no_run
//! use ooxml_wml::Document;
//!
//! // Open an existing document
//! let doc = Document::open("input.docx")?;
//! for para in doc.body().paragraphs() {
//!     println!("{}", para.text());
//! }
//! # Ok::<(), ooxml_wml::Error>(())
//! ```

pub mod document;
pub mod error;

// Generated types from ECMA-376 schema.
// Access via `ooxml_wml::types::*` for generated structs/enums.
pub mod types {
    #![allow(dead_code)]
    include!(concat!(env!("OUT_DIR"), "/wml_types.rs"));
}

pub use document::{Body, Document, Paragraph, ParagraphProperties, Run, RunProperties};
pub use error::{Error, Result};

// Re-export commonly used generated types at the crate root
pub use types::ns;

// TODO: Styles support
// TODO: Tables
// TODO: Lists/numbering
// TODO: Images
// TODO: Hyperlinks
