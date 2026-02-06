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
//! - `ooxml-sml` - SpreadsheetML (XLSX)
//! - `ooxml-pml` - PresentationML (PPTX)
//!
//! # Example
//!
//! ```no_run
//! use ooxml_opc::{Package, Relationships, rel_type, rels_path_for};
//! use std::fs::File;
//!
//! let file = File::open("document.docx")?;
//! let mut pkg = Package::open(file)?;
//!
//! // Read package relationships
//! let rels_data = pkg.read_part("_rels/.rels")?;
//! let rels = Relationships::parse(&rels_data[..])?;
//!
//! // Find the main document
//! if let Some(doc_rel) = rels.get_by_type(rel_type::OFFICE_DOCUMENT) {
//!     let doc_xml = pkg.read_part_string(&doc_rel.target)?;
//!     println!("Document: {}", doc_xml);
//! }
//! # Ok::<(), ooxml_opc::Error>(())
//! ```

pub mod error;
pub mod packaging;
pub mod relationships;

pub use error::{Error, Result};
pub use packaging::{ContentTypes, Package, PackageWriter, content_type};
pub use relationships::{Relationship, Relationships, TargetMode, rel_type, rels_path_for};
