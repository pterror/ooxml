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

    /// XML parsing error with location context.
    #[error("{context}: {message}")]
    Parse {
        /// Human-readable context (e.g., file path, element being parsed).
        context: String,
        /// The error message.
        message: String,
        /// Byte position in the source where the error occurred.
        position: Option<u64>,
    },

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

    /// Raw XML parsing error.
    #[error("raw XML error: {0}")]
    RawXml(#[from] ooxml_xml::Error),
}

impl Error {
    /// Create a parse error with context.
    pub fn parse(context: impl Into<String>, message: impl Into<String>) -> Self {
        Self::Parse {
            context: context.into(),
            message: message.into(),
            position: None,
        }
    }

    /// Create a parse error with context and position.
    pub fn parse_at(context: impl Into<String>, message: impl Into<String>, position: u64) -> Self {
        Self::Parse {
            context: context.into(),
            message: message.into(),
            position: Some(position),
        }
    }

    /// Add context to an existing error.
    pub fn with_context(self, context: impl Into<String>) -> Self {
        match self {
            Self::Xml(e) => Self::Parse {
                context: context.into(),
                message: e.to_string(),
                position: None,
            },
            Self::Parse {
                message, position, ..
            } => Self::Parse {
                context: context.into(),
                message,
                position,
            },
            other => other,
        }
    }

    /// Add position information to an existing error.
    pub fn at_position(self, position: u64) -> Self {
        match self {
            Self::Parse {
                context, message, ..
            } => Self::Parse {
                context,
                message,
                position: Some(position),
            },
            Self::Xml(e) => Self::Parse {
                context: String::new(),
                message: e.to_string(),
                position: Some(position),
            },
            other => other,
        }
    }
}

/// Context for tracking parsing location.
///
/// Used to provide better error messages with file paths and element context.
#[derive(Debug, Clone, Default)]
pub struct ParseContext {
    /// The file path being parsed (e.g., "word/document.xml").
    pub file_path: Option<String>,
    /// Stack of element names being parsed (for nested context).
    pub element_stack: Vec<String>,
}

impl ParseContext {
    /// Create a new parse context for a file.
    pub fn new(file_path: impl Into<String>) -> Self {
        Self {
            file_path: Some(file_path.into()),
            element_stack: Vec::new(),
        }
    }

    /// Push an element onto the context stack.
    pub fn push(&mut self, element: impl Into<String>) {
        self.element_stack.push(element.into());
    }

    /// Pop an element from the context stack.
    pub fn pop(&mut self) {
        self.element_stack.pop();
    }

    /// Get a human-readable description of the current context.
    pub fn describe(&self) -> String {
        let mut parts = Vec::new();
        if let Some(ref path) = self.file_path {
            parts.push(path.clone());
        }
        if !self.element_stack.is_empty() {
            parts.push(format!("in <{}>", self.element_stack.join("/")));
        }
        if parts.is_empty() {
            "unknown location".to_string()
        } else {
            parts.join(" ")
        }
    }

    /// Create an error with this context.
    pub fn error(&self, message: impl Into<String>) -> Error {
        Error::parse(self.describe(), message)
    }

    /// Create an error with this context and position.
    pub fn error_at(&self, message: impl Into<String>, position: u64) -> Error {
        Error::parse_at(self.describe(), message, position)
    }
}

/// Convert a byte position to line and column numbers.
///
/// Returns (line, column) where both are 1-indexed.
pub fn position_to_line_col(content: &[u8], position: u64) -> (usize, usize) {
    let position = position as usize;
    let content = if position <= content.len() {
        &content[..position]
    } else {
        content
    };

    let mut line = 1;
    let mut col = 1;

    for &byte in content {
        if byte == b'\n' {
            line += 1;
            col = 1;
        } else {
            col += 1;
        }
    }

    (line, col)
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_position_to_line_col() {
        let content = b"line1\nline2\nline3";
        assert_eq!(position_to_line_col(content, 0), (1, 1));
        assert_eq!(position_to_line_col(content, 5), (1, 6)); // at '\n'
        assert_eq!(position_to_line_col(content, 6), (2, 1)); // start of line2
        assert_eq!(position_to_line_col(content, 12), (3, 1)); // start of line3
    }

    #[test]
    fn test_parse_context() {
        let mut ctx = ParseContext::new("word/document.xml");
        assert_eq!(ctx.describe(), "word/document.xml");

        ctx.push("w:body");
        assert_eq!(ctx.describe(), "word/document.xml in <w:body>");

        ctx.push("w:p");
        assert_eq!(ctx.describe(), "word/document.xml in <w:body/w:p>");

        ctx.pop();
        assert_eq!(ctx.describe(), "word/document.xml in <w:body>");
    }

    #[test]
    fn test_error_with_context() {
        let err = Error::parse("word/document.xml", "unexpected element");
        assert!(err.to_string().contains("word/document.xml"));
        assert!(err.to_string().contains("unexpected element"));
    }
}
