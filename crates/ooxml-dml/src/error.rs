//! Error types for DrawingML parsing.

use thiserror::Error;

/// Errors that can occur when working with DrawingML content.
#[derive(Error, Debug)]
pub enum Error {
    /// XML parsing error.
    #[error("XML error: {0}")]
    Xml(#[from] quick_xml::Error),

    /// Invalid or malformed content.
    #[error("Invalid content: {0}")]
    Invalid(String),
}

/// Result type for DrawingML operations.
pub type Result<T> = std::result::Result<T, Error>;
