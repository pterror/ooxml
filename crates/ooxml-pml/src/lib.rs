//! PresentationML (PPTX) support for the ooxml library.
//!
//! This crate provides reading and writing of PowerPoint presentations (.pptx files).
//!
//! # Reading Presentations
//!
//! ```no_run
//! use ooxml_pml::{Presentation, ShapeExt};
//!
//! let mut pres = Presentation::open("presentation.pptx")?;
//! println!("Slides: {}", pres.slide_count());
//! for slide in pres.slides()? {
//!     println!("Slide: {}", slide.index());
//!     for shape in slide.shapes() {
//!         if let Some(text) = shape.text() {
//!             println!("  Text: {}", text);
//!         }
//!     }
//! }
//! # Ok::<(), ooxml_pml::Error>(())
//! ```
//!
//! # Writing Presentations
//!
//! ```no_run
//! use ooxml_pml::PresentationBuilder;
//!
//! let mut pres = PresentationBuilder::new();
//! let slide = pres.add_slide();
//! slide.add_title("Hello World");
//! slide.add_text("Created with ooxml-pml");
//! pres.save("output.pptx")?;
//! # Ok::<(), ooxml_pml::Error>(())
//! ```

pub mod error;
pub mod ext;
pub mod presentation;
pub mod writer;

/// Generated types from the ECMA-376 PresentationML schema.
///
/// These types map 1:1 to XML elements and attributes defined in ECMA-376 Part 1 §19.
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
#[cfg(feature = "pml-charts")]
pub use presentation::SmartArtParts;
pub use presentation::{
    DiagramRelIds, Hyperlink, ImageData, Presentation, Slide, SlideLayout, SlideLayoutType,
    SlideMaster, Table, Transition, TransitionSpeed, TransitionType,
};
// Re-export generated types that replace handwritten ones
pub use types::{Picture, Shape};
// Re-export DML table types and extension traits for table access
pub use ooxml_dml::types::{CTTable, CTTableCell, CTTableRow};
pub use ooxml_dml::{TableCellExt, TableExt, TableRowExt};
pub use writer::{
    GroupBuilder, ImageFormat, Paragraph, PresentationBuilder, PresetGeometry, ShapeBuilder,
    SlideBuilder, SlideTransition, TableBuilder, TextAlign, TextRun,
};

// Extension traits for generated types
#[cfg(feature = "pml-notes")]
pub use ext::NotesSlideExt;
pub use ext::{
    CommonSlideDataExt, ConnectorExt, GraphicalObjectFrameExt, GroupShapeExt, PictureExt, ShapeExt,
    SlideExt, SlideLayoutExt, SlideMasterExt,
};
