//! DrawingML (DML) support for the ooxml library.
//!
//! This crate provides shared DrawingML types used by Word (WML),
//! Excel (SML), and PowerPoint (PML) documents.
//!
//! DrawingML is defined in ECMA-376 Part 4 and provides common
//! elements for text formatting, shapes, images, and charts.
//!
//! # Text Content
//!
//! DrawingML text is structured as paragraphs containing runs:
//!
//! ```
//! use ooxml_dml::text::{Paragraph, Run};
//!
//! let mut para = Paragraph::new();
//! para.add_run(Run::new("Hello "));
//! para.add_run(Run::new("World"));
//! assert_eq!(para.text(), "Hello World");
//! ```

pub mod error;
pub mod text;

// Generated types from ECMA-376 schema.
// Access via `ooxml_dml::types::*` for generated structs/enums.
// This file is pre-generated and committed to avoid requiring spec downloads.
// To regenerate: OOXML_REGENERATE=1 cargo build -p ooxml-dml (with specs in /spec/)
#[allow(dead_code)]
pub mod generated;
pub use generated as types;

// TODO: DML parser/serializer generation has codegen issues with EG_* type handling
// Uncomment once codegen is fixed:
// pub mod generated_parsers;
// pub use generated_parsers as parsers;
// pub mod generated_serializers;
// pub use generated_serializers as serializers;

pub use error::{Error, Result};
pub use text::{
    Paragraph, ParagraphProperties, Run, RunProperties, TextAlignment, parse_text_body,
    parse_text_body_from_reader,
};
