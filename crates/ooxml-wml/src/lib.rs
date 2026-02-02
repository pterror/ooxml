//! WordprocessingML (DOCX) support for the ooxml library.
//!
//! This crate provides reading and writing of Word documents (.docx files).
//!
//! # Reading Documents
//!
//! ```ignore
//! use ooxml_wml::Document;
//! use ooxml_wml::ext::{BodyExt, ParagraphExt};
//!
//! let doc = Document::open("input.docx")?;
//! for para in doc.body().paragraphs() {
//!     println!("{}", para.text());
//! }
//! ```
//!
//! # Creating Documents
//!
//! ```ignore
//! use ooxml_wml::DocumentBuilder;
//!
//! let mut builder = DocumentBuilder::new();
//! builder.add_paragraph("Hello, World!");
//! builder.save("output.docx")?;
//! ```

pub mod convenience;
pub mod document;
pub mod error;
pub mod ext;
pub mod writer;

// Generated types from ECMA-376 schema.
// Access via `ooxml_wml::types::*` for generated structs/enums.
// This file is pre-generated and committed to avoid requiring spec downloads.
// To regenerate: cargo build -p ooxml-wml (with specs in /spec/)
#[allow(dead_code)]
pub mod generated;
pub use generated as types;

pub mod generated_parsers;
pub use generated_parsers as parsers;

pub mod generated_serializers;
pub use generated_serializers as serializers;

// Metadata types from document.rs (OPC, not WML — not generated).
pub use document::{AppProperties, CoreProperties, Document, DocumentSettings, ImageData};

// Error types — always available.
pub use error::{Error, ParseContext, Result, position_to_line_col};
pub use ooxml_xml::{PositionedAttr, PositionedNode, RawXmlElement, RawXmlNode};

// Writer types.
pub use writer::{
    AnchoredImage, CommentBuilder, DocumentBuilder, Drawing, EndnoteBuilder, FooterBuilder,
    FootnoteBuilder, HeaderBuilder, HeaderFooterType, InlineImage, ListType, WrapType,
};

// Re-export commonly used generated types at the crate root.
pub use types::ns;

// Re-export MathZone from ooxml-omml for convenience.
pub use ooxml_omml::MathZone;
