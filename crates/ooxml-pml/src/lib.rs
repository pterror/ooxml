//! PresentationML (PPTX) support for the ooxml library.
//!
//! This crate provides reading and writing of PowerPoint presentations (.pptx files).
//!
//! # Reading Presentations
//!
//! ```no_run
//! use ooxml_pml::Presentation;
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
pub mod presentation;
pub mod writer;

pub use error::{Error, Result};
pub use presentation::{
    Hyperlink, ImageData, Picture, Presentation, Shape, Slide, SlideLayout, SlideLayoutType,
    SlideMaster, Table, TableCell, TableRow, Transition, TransitionSpeed, TransitionType,
};
pub use writer::{ImageFormat, PresentationBuilder, SlideBuilder, TableBuilder, TextRun};
