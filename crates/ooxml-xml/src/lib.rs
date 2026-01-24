//! Shared XML utilities for the ooxml library.
//!
//! This crate provides common XML types and utilities used across OOXML format crates
//! for features like roundtrip fidelity (preserving unknown elements/attributes).

mod raw_xml;

pub use raw_xml::{PositionedAttr, PositionedNode, RawXmlElement, RawXmlNode};

/// Error type for XML operations.
#[derive(Debug, thiserror::Error)]
pub enum Error {
    #[error("XML error: {0}")]
    Xml(#[from] quick_xml::Error),
    #[error("IO error: {0}")]
    Io(#[from] std::io::Error),
    #[error("invalid XML: {0}")]
    Invalid(String),
}

/// Result type for XML operations.
pub type Result<T> = std::result::Result<T, Error>;
