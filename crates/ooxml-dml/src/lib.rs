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
//! DrawingML text is structured as paragraphs containing runs. Use the
//! extension traits from [`ext`] for convenient access:
//!
//! ```ignore
//! use ooxml_dml::ext::{TextBodyExt, TextParagraphExt, TextRunExt};
//! use ooxml_dml::types::TextBody;
//!
//! fn process_text(body: &TextBody) {
//!     for para in body.paragraphs() {
//!         println!("Paragraph: {}", para.text());
//!         for run in para.runs() {
//!             if run.is_bold() {
//!                 println!("  Bold: {}", run.text());
//!             }
//!         }
//!     }
//! }
//! ```

pub mod error;
pub mod ext;

// Generated types from ECMA-376 schema.
// Access via `ooxml_dml::types::*` for generated structs/enums.
// This file is pre-generated and committed to avoid requiring spec downloads.
// To regenerate: OOXML_REGENERATE=1 cargo build -p ooxml-dml (with specs in /spec/)
#[allow(dead_code)]
pub mod generated;
pub use generated as types;

pub mod generated_parsers;
pub use generated_parsers as parsers;
pub mod generated_serializers;
pub use generated_serializers as serializers;

pub use error::{Error, Result};
#[cfg(feature = "dml-text")]
pub use ext::{TextBodyExt, TextParagraphExt, TextRunExt};
