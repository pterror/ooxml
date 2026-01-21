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

pub mod error;
pub mod presentation;

pub use error::{Error, Result};
pub use presentation::{
    ImageData, Picture, Presentation, Shape, Slide, Transition, TransitionSpeed, TransitionType,
};
