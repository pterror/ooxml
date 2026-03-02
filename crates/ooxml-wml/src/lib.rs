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

/// Generated types from the ECMA-376 WordprocessingML schema.
///
/// These types map 1:1 to XML elements and attributes defined in ECMA-376 Part 1 §17.
/// They are produced by `ooxml-codegen` from the RELAX NG schemas and committed to avoid
/// requiring the schema files at build time. Use the extension traits in [`ext`] for
/// ergonomic access rather than working with these types directly.
///
/// Re-exported as [`types`].
#[allow(dead_code)]
pub mod generated;
/// Type aliases for the generated ECMA-376 types. See [`generated`] for details.
pub use generated as types;

/// Generated [`FromXml`](ooxml_xml::FromXml) parsers for all generated types.
///
/// Re-exported as [`parsers`].
pub mod generated_parsers;
/// Parsers for the generated ECMA-376 types. See [`generated_parsers`] for details.
pub use generated_parsers as parsers;

/// Generated [`ToXml`](ooxml_xml::ToXml) serializers for all generated types.
///
/// Re-exported as [`serializers`].
pub mod generated_serializers;
/// Serializers for the generated ECMA-376 types. See [`generated_serializers`] for details.
pub use generated_serializers as serializers;

// Metadata types from document.rs (OPC, not WML — not generated).
pub use document::{AppProperties, CoreProperties, Document, DocumentSettings, ImageData};

// Error types — always available.
pub use error::{Error, ParseContext, Result, position_to_line_col};
pub use ooxml_xml::{PositionedAttr, PositionedNode, RawXmlElement, RawXmlNode};

// Writer types.
pub use writer::{
    AnchoredImage, CommentBuilder, DocumentBuilder, Drawing, EndnoteBuilder, FooterBuilder,
    FootnoteBuilder, HeaderBuilder, HeaderFooterType, InlineImage, ListType, NumberingLevel,
    TextBox, WrapType,
};

// Re-export commonly used generated types at the crate root.
pub use types::ns;

// Re-export MathZone from ooxml-omml for convenience.
#[cfg(feature = "wml-math")]
pub use ooxml_omml::MathZone;
