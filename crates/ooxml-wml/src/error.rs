//! Error types for the ooxml-wml crate.

use thiserror::Error;

/// Result type for ooxml-wml operations.
pub type Result<T> = std::result::Result<T, Error>;

/// Errors that can occur when working with Word documents.
#[derive(Debug, Error)]
pub enum Error {
    /// Error from the core ooxml crate (packaging, relationships).
    #[error("package error: {0}")]
    Package(#[from] ooxml_opc::Error),

    /// XML parsing error.
    #[error("XML error: {0}")]
    Xml(#[from] quick_xml::Error),

    /// Invalid or malformed document structure.
    #[error("invalid document: {0}")]
    Invalid(String),

    /// Unsupported feature or element.
    #[error("unsupported: {0}")]
    Unsupported(String),

    /// Required part is missing from the package.
    #[error("missing part: {0}")]
    MissingPart(String),

    /// UTF-8 decoding error.
    #[error("UTF-8 error: {0}")]
    Utf8(#[from] std::string::FromUtf8Error),

    /// I/O error.
    #[error("I/O error: {0}")]
    Io(#[from] std::io::Error),
}
