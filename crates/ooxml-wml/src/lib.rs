//! WordprocessingML (DOCX) support for the ooxml library.
//!
//! This crate provides reading and writing of Word documents (.docx files).
//!
//! # Reading Documents
//!
//! Requires the `read` feature (enabled by default in `full`).
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
//! Requires the `write` feature (enabled by default in `full`).
//!
//! ```ignore
//! use ooxml_wml::DocumentBuilder;
//!
//! let mut builder = DocumentBuilder::new();
//! builder.add_paragraph("Hello, World!");
//! builder.save("output.docx")?;
//! ```

pub mod document;
pub mod error;

#[cfg(feature = "read")]
pub mod ext;

#[cfg(feature = "write")]
pub mod styles;

#[cfg(feature = "write")]
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

// Handwritten types from document.rs — used by both read (header/footer/footnotes)
// and write (DocumentBuilder). Always available.
pub use document::{
    Alignment, AppProperties, BlockContent, Body, BookmarkEnd, BookmarkStart, Cell,
    CommentRangeEnd, CommentRangeStart, CommentReference, CoreProperties, Deletion, Drawing,
    EmbeddedObject, EndnoteReference, Fonts, FootnoteReference, HeaderFooterType, Hyperlink,
    ImageData, InlineImage, Insertion, NumberingProperties, Paragraph, ParagraphContent,
    ParagraphProperties, Row, Run, RunProperties, Table, TextBoxContent, VmlPicture,
};

// Document reader — requires `read` feature.
#[cfg(feature = "read")]
pub use document::Document;

// Error types — always available.
pub use error::{Error, ParseContext, Result, position_to_line_col};
pub use ooxml_xml::{PositionedAttr, PositionedNode, RawXmlElement, RawXmlNode};

// Handwritten styles — only needed by the writer.
#[cfg(feature = "write")]
pub use styles::{Style, StyleType, Styles};

// Writer — requires `write` feature.
#[cfg(feature = "write")]
pub use writer::{
    CommentBuilder, DocumentBuilder, EndnoteBuilder, FooterBuilder, FootnoteBuilder, HeaderBuilder,
    ListType,
};

// Re-export commonly used generated types at the crate root.
pub use types::ns;

// Re-export MathZone from ooxml-omml for convenience.
pub use ooxml_omml::MathZone;

// ## Current Limitations (v0.1)
//
// Images:
// - Inline images only (no floating/anchored images)
// - No image cropping or effects
// - No linked images (external URLs)
// - Basic positioning only (no text wrapping options)
