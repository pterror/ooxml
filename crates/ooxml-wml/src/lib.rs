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
pub mod raw_xml;
pub mod styles;
pub mod writer;

// Generated types from ECMA-376 schema.
// Access via `ooxml_wml::types::*` for generated structs/enums.
pub mod types {
    #![allow(dead_code)]
    include!(concat!(env!("OUT_DIR"), "/wml_types.rs"));
}

pub use document::{
    Alignment, AppProperties, BlockContent, Body, BookmarkEnd, BookmarkStart, Cell,
    CommentRangeEnd, CommentRangeStart, CommentReference, CoreProperties, Deletion, Document,
    Drawing, EmbeddedObject, EndnoteReference, Fonts, FootnoteReference, HeaderFooterType,
    Hyperlink, ImageData, InlineImage, Insertion, NumberingProperties, Paragraph, ParagraphContent,
    ParagraphProperties, Row, Run, RunProperties, Table, TextBoxContent, VmlPicture,
};
pub use error::{Error, ParseContext, Result, position_to_line_col};
pub use raw_xml::{PositionedAttr, PositionedNode, RawXmlElement, RawXmlNode};
pub use styles::{Style, StyleType, Styles};
pub use writer::{
    CommentBuilder, DocumentBuilder, EndnoteBuilder, FooterBuilder, FootnoteBuilder, HeaderBuilder,
    ListType,
};

// Re-export commonly used generated types at the crate root
pub use types::ns;

// Re-export MathZone from ooxml-omml for convenience
pub use ooxml_omml::MathZone;

// ## Current Limitations (v0.1)
//
// Images:
// - Inline images only (no floating/anchored images)
// - No image cropping or effects
// - No linked images (external URLs)
// - Basic positioning only (no text wrapping options)
