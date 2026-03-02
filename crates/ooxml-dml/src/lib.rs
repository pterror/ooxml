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

/// Generated types from the ECMA-376 DrawingML schema.
///
/// These types map 1:1 to XML elements and attributes defined in ECMA-376 Part 4 §20–21.
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

pub use error::{Error, Result};
#[cfg(feature = "dml-diagrams")]
pub use ext::DataModelExt;
#[cfg(feature = "dml-charts")]
pub use ext::{ChartExt, ChartKind, ChartSpaceExt, ChartTitleExt, PlotAreaExt};
#[cfg(feature = "dml-tables")]
pub use ext::{TableCellExt, TableExt, TableRowExt};
#[cfg(feature = "dml-text")]
pub use ext::{TextBodyExt, TextParagraphExt, TextRunExt};
