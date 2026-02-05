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

// Generated types from ECMA-376 schema.
// Access via `ooxml_pml::types::*` for generated structs/enums.
// This file is pre-generated and committed to avoid requiring spec downloads.
// To regenerate: OOXML_REGENERATE=1 cargo build -p ooxml-pml (with specs in /spec/)
#[allow(dead_code)]
pub mod generated;
pub use generated as types;

pub mod generated_parsers;
pub use generated_parsers as parsers;
pub mod generated_serializers;
pub use generated_serializers as serializers;

pub use error::{Error, Result};
pub use presentation::{
    Hyperlink, ImageData, Presentation, Slide, SlideLayout, SlideLayoutType, SlideMaster, Table,
    TableCell, TableRow, Transition, TransitionSpeed, TransitionType,
};
// Re-export generated types that replace handwritten ones
pub use types::{Picture, Shape};
pub use writer::{
    ImageFormat, PresentationBuilder, SlideBuilder, SlideTransition, TableBuilder, TextRun,
};

// Extension traits for generated types
#[cfg(feature = "pml-notes")]
pub use ext::NotesSlideExt;
pub use ext::{
    CommonSlideDataExt, ConnectorExt, GraphicalObjectFrameExt, GroupShapeExt, PictureExt, ShapeExt,
    SlideExt, SlideLayoutExt, SlideMasterExt,
};
