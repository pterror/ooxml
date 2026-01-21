//! Error types for SpreadsheetML parsing and writing.

use thiserror::Error;

/// Errors that can occur when working with XLSX files.
#[derive(Error, Debug)]
pub enum Error {
    /// Error from the OPC packaging layer.
    #[error("OPC error: {0}")]
    Opc(#[from] ooxml_opc::Error),

    /// XML parsing error.
    #[error("XML error: {0}")]
    Xml(#[from] quick_xml::Error),

    /// I/O error.
    #[error("I/O error: {0}")]
    Io(#[from] std::io::Error),

    /// Invalid or malformed content.
    #[error("Invalid content: {0}")]
    Invalid(String),

    /// Unsupported feature.
    #[error("Unsupported feature: {0}")]
    Unsupported(String),
}

/// Result type for SpreadsheetML operations.
pub type Result<T> = std::result::Result<T, Error>;
