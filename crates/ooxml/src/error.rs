//! Error types for the ooxml crate.

use thiserror::Error;

/// Result type for ooxml operations.
pub type Result<T> = std::result::Result<T, Error>;

/// Errors that can occur when working with OOXML files.
#[derive(Debug, Error)]
pub enum Error {
    /// IO error (file operations, ZIP handling).
    #[error("io error: {0}")]
    Io(#[from] std::io::Error),

    /// ZIP archive error.
    #[error("zip error: {0}")]
    Zip(#[from] zip::result::ZipError),

    /// XML parsing error.
    #[error("xml error: {0}")]
    Xml(#[from] quick_xml::Error),

    /// Invalid or malformed OOXML structure.
    #[error("invalid ooxml: {0}")]
    Invalid(String),

    /// Missing required part in the package.
    #[error("missing part: {0}")]
    MissingPart(String),

    /// Unsupported feature or element.
    #[error("unsupported: {0}")]
    Unsupported(String),
}
