//! Structured error analysis for OOXML parsing failures.
//!
//! This module provides tools to categorize and analyze parse errors,
//! replacing simple string matching with structured error information.

use serde::{Deserialize, Serialize};

/// Categorized error from parsing.
#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct AnalyzedError {
    /// High-level error category.
    pub category: ErrorCategory,
    /// More specific subcategory or element name.
    pub subcategory: Option<String>,
    /// Human-readable error message.
    pub message: String,
    /// Location information if available.
    pub location: Option<ErrorLocation>,
    /// Original error string for reference.
    pub raw_error: String,
}

/// High-level error categories.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash, Serialize, Deserialize)]
pub enum ErrorCategory {
    // Package-level errors
    /// ZIP archive corruption or read failure.
    ZipCorruption,
    /// Unsupported ZIP compression method.
    ZipUnsupportedCompression,
    /// Required part missing from package.
    MissingRequiredPart,
    /// Invalid content type declaration.
    InvalidContentType,

    // XML-level errors
    /// Malformed XML syntax.
    XmlMalformed,
    /// XML encoding error (UTF-8, etc).
    XmlEncodingError,
    /// XML namespace error.
    XmlNamespaceError,

    // Schema-level errors
    /// Unexpected element in document.
    UnexpectedElement,
    /// Missing required attribute.
    MissingRequiredAttribute,
    /// Invalid attribute value.
    InvalidAttributeValue,

    // Relationship errors
    /// Relationship target doesn't exist.
    BrokenRelationship,
    /// Referenced file missing.
    MissingTarget,

    // Parser limitations
    /// Feature not yet implemented.
    UnsupportedFeature,
    /// Element type not handled.
    NotImplemented,

    // I/O errors
    /// File read/write error.
    IoError,

    // Unknown
    /// Could not categorize error.
    Unknown,
}

impl ErrorCategory {
    /// Get a short string representation for display/storage.
    pub fn as_str(&self) -> &'static str {
        match self {
            ErrorCategory::ZipCorruption => "zip_corruption",
            ErrorCategory::ZipUnsupportedCompression => "zip_unsupported_compression",
            ErrorCategory::MissingRequiredPart => "missing_required_part",
            ErrorCategory::InvalidContentType => "invalid_content_type",
            ErrorCategory::XmlMalformed => "xml_malformed",
            ErrorCategory::XmlEncodingError => "xml_encoding",
            ErrorCategory::XmlNamespaceError => "xml_namespace",
            ErrorCategory::UnexpectedElement => "unexpected_element",
            ErrorCategory::MissingRequiredAttribute => "missing_attribute",
            ErrorCategory::InvalidAttributeValue => "invalid_attribute",
            ErrorCategory::BrokenRelationship => "broken_relationship",
            ErrorCategory::MissingTarget => "missing_target",
            ErrorCategory::UnsupportedFeature => "unsupported_feature",
            ErrorCategory::NotImplemented => "not_implemented",
            ErrorCategory::IoError => "io_error",
            ErrorCategory::Unknown => "unknown",
        }
    }

    /// Get a human-readable description.
    pub fn description(&self) -> &'static str {
        match self {
            ErrorCategory::ZipCorruption => "ZIP archive is corrupted or unreadable",
            ErrorCategory::ZipUnsupportedCompression => "ZIP uses unsupported compression",
            ErrorCategory::MissingRequiredPart => "Required document part is missing",
            ErrorCategory::InvalidContentType => "Invalid content type declaration",
            ErrorCategory::XmlMalformed => "XML syntax is malformed",
            ErrorCategory::XmlEncodingError => "XML encoding error (invalid UTF-8, etc)",
            ErrorCategory::XmlNamespaceError => "XML namespace error",
            ErrorCategory::UnexpectedElement => "Unexpected XML element found",
            ErrorCategory::MissingRequiredAttribute => "Required attribute is missing",
            ErrorCategory::InvalidAttributeValue => "Attribute has invalid value",
            ErrorCategory::BrokenRelationship => "Relationship references missing target",
            ErrorCategory::MissingTarget => "Referenced file is missing",
            ErrorCategory::UnsupportedFeature => "Feature is not yet supported",
            ErrorCategory::NotImplemented => "Element type not implemented",
            ErrorCategory::IoError => "File I/O error",
            ErrorCategory::Unknown => "Unknown error type",
        }
    }
}

/// Location information for an error.
#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct ErrorLocation {
    /// Path within the package (e.g., "word/document.xml").
    pub part_path: Option<String>,
    /// XPath-like element path (e.g., "/w:document/w:body/w:p[5]").
    pub element_path: Option<String>,
    /// Line number in XML file.
    pub line: Option<u32>,
    /// Column number in XML file.
    pub column: Option<u32>,
}

/// Analyze an ooxml-wml error and extract structured information.
pub fn analyze_error(error: &ooxml_wml::Error) -> AnalyzedError {
    let raw_error = error.to_string();

    match error {
        ooxml_wml::Error::Package(pkg_err) => analyze_package_error(pkg_err, &raw_error),
        ooxml_wml::Error::Xml(xml_err) => analyze_xml_error(xml_err, &raw_error),
        ooxml_wml::Error::Invalid(msg) => analyze_invalid_error(msg, &raw_error),
        ooxml_wml::Error::Unsupported(msg) => AnalyzedError {
            category: ErrorCategory::UnsupportedFeature,
            subcategory: extract_feature_name(msg),
            message: msg.clone(),
            location: None,
            raw_error,
        },
        ooxml_wml::Error::MissingPart(part) => AnalyzedError {
            category: ErrorCategory::MissingRequiredPart,
            subcategory: Some(part.clone()),
            message: format!("Missing required part: {}", part),
            location: Some(ErrorLocation {
                part_path: Some(part.clone()),
                element_path: None,
                line: None,
                column: None,
            }),
            raw_error,
        },
        ooxml_wml::Error::Utf8(_) => AnalyzedError {
            category: ErrorCategory::XmlEncodingError,
            subcategory: Some("UTF-8".to_string()),
            message: "UTF-8 decoding error".to_string(),
            location: None,
            raw_error,
        },
        ooxml_wml::Error::Io(io_err) => AnalyzedError {
            category: ErrorCategory::IoError,
            subcategory: Some(io_err.kind().to_string()),
            message: io_err.to_string(),
            location: None,
            raw_error,
        },
        ooxml_wml::Error::Parse {
            context,
            message,
            position,
        } => AnalyzedError {
            category: ErrorCategory::XmlMalformed,
            subcategory: Some(context.clone()),
            message: message.clone(),
            location: Some(ErrorLocation {
                part_path: Some(context.clone()),
                element_path: None,
                line: None,
                column: position.map(|p| p as u32),
            }),
            raw_error,
        },
    }
}

/// Analyze a package-level error.
fn analyze_package_error(error: &ooxml_opc::Error, raw_error: &str) -> AnalyzedError {
    match error {
        ooxml_opc::Error::Zip(zip_err) => {
            let zip_str = zip_err.to_string();
            let category = if zip_str.contains("invalid") || zip_str.contains("corrupt") {
                ErrorCategory::ZipCorruption
            } else if zip_str.contains("unsupported") || zip_str.contains("compression") {
                ErrorCategory::ZipUnsupportedCompression
            } else {
                ErrorCategory::ZipCorruption
            };
            AnalyzedError {
                category,
                subcategory: None,
                message: zip_str,
                location: None,
                raw_error: raw_error.to_string(),
            }
        }
        ooxml_opc::Error::Xml(xml_err) => analyze_xml_error(xml_err, raw_error),
        ooxml_opc::Error::Invalid(msg) => analyze_invalid_error(msg, raw_error),
        ooxml_opc::Error::MissingPart(part) => AnalyzedError {
            category: ErrorCategory::MissingRequiredPart,
            subcategory: Some(part.clone()),
            message: format!("Missing required part: {}", part),
            location: Some(ErrorLocation {
                part_path: Some(part.clone()),
                element_path: None,
                line: None,
                column: None,
            }),
            raw_error: raw_error.to_string(),
        },
        ooxml_opc::Error::Unsupported(msg) => AnalyzedError {
            category: ErrorCategory::UnsupportedFeature,
            subcategory: extract_feature_name(msg),
            message: msg.clone(),
            location: None,
            raw_error: raw_error.to_string(),
        },
        ooxml_opc::Error::Io(io_err) => AnalyzedError {
            category: ErrorCategory::IoError,
            subcategory: Some(io_err.kind().to_string()),
            message: io_err.to_string(),
            location: None,
            raw_error: raw_error.to_string(),
        },
    }
}

/// Analyze an XML parsing error.
fn analyze_xml_error(error: &quick_xml::Error, raw_error: &str) -> AnalyzedError {
    let msg = error.to_string();

    let category = if msg.contains("encoding") || msg.contains("UTF") || msg.contains("utf") {
        ErrorCategory::XmlEncodingError
    } else if msg.contains("namespace") {
        ErrorCategory::XmlNamespaceError
    } else {
        ErrorCategory::XmlMalformed
    };

    AnalyzedError {
        category,
        subcategory: None,
        message: msg,
        location: None, // quick-xml doesn't always provide position
        raw_error: raw_error.to_string(),
    }
}

/// Analyze an "Invalid" error message to categorize it.
fn analyze_invalid_error(msg: &str, raw_error: &str) -> AnalyzedError {
    let lower = msg.to_lowercase();

    let category = if lower.contains("element") || lower.contains("unexpected") {
        ErrorCategory::UnexpectedElement
    } else if lower.contains("attribute") && lower.contains("missing") {
        ErrorCategory::MissingRequiredAttribute
    } else if lower.contains("attribute") {
        ErrorCategory::InvalidAttributeValue
    } else if lower.contains("content type") {
        ErrorCategory::InvalidContentType
    } else if lower.contains("relationship") {
        ErrorCategory::BrokenRelationship
    } else {
        ErrorCategory::Unknown
    };

    AnalyzedError {
        category,
        subcategory: extract_element_name(msg),
        message: msg.to_string(),
        location: None,
        raw_error: raw_error.to_string(),
    }
}

/// Try to extract a feature name from an "unsupported" message.
fn extract_feature_name(msg: &str) -> Option<String> {
    // Look for quoted strings or element names
    if let Some(start) = msg.find('\'')
        && let Some(end) = msg[start + 1..].find('\'')
    {
        return Some(msg[start + 1..start + 1 + end].to_string());
    }
    if let Some(start) = msg.find('<')
        && let Some(end) = msg[start..].find('>')
    {
        return Some(msg[start..start + end + 1].to_string());
    }
    None
}

/// Try to extract an element name from an error message.
fn extract_element_name(msg: &str) -> Option<String> {
    // Look for w:xxx patterns or <xxx> patterns
    let patterns = [
        ("w:", " "),
        ("w:", ","),
        ("w:", ")"),
        ("<", ">"),
        ("<", " "),
    ];

    for (start_pat, end_pat) in patterns {
        if let Some(start) = msg.find(start_pat) {
            let after_start = start + start_pat.len();
            if let Some(end) = msg[after_start..].find(end_pat) {
                let name = &msg[start..after_start + end];
                if !name.is_empty() && name.len() < 50 {
                    return Some(name.to_string());
                }
            }
        }
    }
    None
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_error_category_str() {
        assert_eq!(ErrorCategory::ZipCorruption.as_str(), "zip_corruption");
        assert_eq!(
            ErrorCategory::MissingRequiredPart.as_str(),
            "missing_required_part"
        );
    }

    #[test]
    fn test_extract_feature_name() {
        assert_eq!(
            extract_feature_name("unsupported element 'w:sdt'"),
            Some("w:sdt".to_string())
        );
        assert_eq!(
            extract_feature_name("element <w:foo> not supported"),
            Some("<w:foo>".to_string())
        );
        assert_eq!(extract_feature_name("no special chars"), None);
    }

    #[test]
    fn test_extract_element_name() {
        assert_eq!(
            extract_element_name("unexpected element w:customXml in body"),
            Some("w:customXml".to_string())
        );
    }
}
