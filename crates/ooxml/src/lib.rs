//! Core OOXML library: OPC packaging, relationships, and shared types.
//!
//! This crate provides the foundational types for working with Office Open XML files:
//! - OPC (Open Packaging Conventions) - ZIP-based package format
//! - Relationships - links between package parts
//! - Content types - MIME type mappings
//! - Core/App properties - document metadata
//!
//! Format-specific support is in separate crates:
//! - `ooxml-wml` - WordprocessingML (DOCX)
//! - `ooxml-sml` - SpreadsheetML (XLSX) - planned
//! - `ooxml-pml` - PresentationML (PPTX) - planned

pub mod error;
pub mod packaging;
pub mod relationships;

pub use error::{Error, Result};
