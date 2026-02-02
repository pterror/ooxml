//! Validation checks for OOXML documents.
//!
//! This module provides deep structural validation beyond basic parsing,
//! checking reference integrity, style definitions, and relationships.

use ooxml_wml::Document;
use ooxml_wml::ext::{CellExt, ParagraphExt, RowExt, RunExt, RunPropertiesExt, TableExt};
use ooxml_wml::types::BlockContent;
use serde::{Deserialize, Serialize};
use std::collections::HashSet;
use std::io::{Read, Seek};

/// Result of validating a document.
#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct ValidationResult {
    /// Whether the document passed all validation checks.
    pub is_valid: bool,
    /// Warnings (non-fatal issues).
    pub warnings: Vec<ValidationWarning>,
    /// Errors (fatal issues that indicate corruption or invalid structure).
    pub errors: Vec<ValidationError>,
}

impl ValidationResult {
    /// Create a new valid result with no issues.
    pub fn valid() -> Self {
        Self {
            is_valid: true,
            warnings: Vec::new(),
            errors: Vec::new(),
        }
    }

    /// Add a warning.
    pub fn add_warning(&mut self, warning: ValidationWarning) {
        self.warnings.push(warning);
    }

    /// Add an error (also sets is_valid to false).
    pub fn add_error(&mut self, error: ValidationError) {
        self.is_valid = false;
        self.errors.push(error);
    }

    /// Get total issue count.
    pub fn issue_count(&self) -> usize {
        self.warnings.len() + self.errors.len()
    }
}

/// Warning codes for non-fatal validation issues.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash, Serialize, Deserialize)]
pub enum WarningCode {
    /// Style referenced but not defined in styles.xml.
    UnresolvedStyleReference,
    /// Image missing alt text (accessibility).
    MissingImageAltText,
    /// Empty paragraph with no content.
    EmptyParagraph,
    /// Paragraph with only whitespace.
    WhitespaceOnlyParagraph,
    /// Hyperlink with empty target.
    EmptyHyperlinkTarget,
    /// Numbering reference without definition.
    UnresolvedNumberingReference,
    /// Font not commonly available.
    UncommonFont,
    /// Very deep nesting that may cause rendering issues.
    DeepNesting,
}

impl WarningCode {
    /// Get a short string identifier.
    pub fn as_str(&self) -> &'static str {
        match self {
            WarningCode::UnresolvedStyleReference => "unresolved_style",
            WarningCode::MissingImageAltText => "missing_alt_text",
            WarningCode::EmptyParagraph => "empty_paragraph",
            WarningCode::WhitespaceOnlyParagraph => "whitespace_only",
            WarningCode::EmptyHyperlinkTarget => "empty_hyperlink",
            WarningCode::UnresolvedNumberingReference => "unresolved_numbering",
            WarningCode::UncommonFont => "uncommon_font",
            WarningCode::DeepNesting => "deep_nesting",
        }
    }
}

/// A validation warning.
#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct ValidationWarning {
    /// Warning code.
    pub code: WarningCode,
    /// Human-readable message.
    pub message: String,
    /// Location in document if known.
    pub location: Option<String>,
}

/// Error codes for fatal validation issues.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash, Serialize, Deserialize)]
pub enum ErrorCode {
    /// Hyperlink references missing relationship.
    BrokenHyperlinkReference,
    /// Image references missing relationship or file.
    BrokenImageReference,
    /// Numbering definition referenced but missing.
    MissingNumberingDefinition,
    /// Required document part is missing.
    MissingRequiredPart,
    /// Circular style inheritance detected.
    CircularStyleInheritance,
    /// Invalid XML structure.
    InvalidStructure,
}

impl ErrorCode {
    /// Get a short string identifier.
    pub fn as_str(&self) -> &'static str {
        match self {
            ErrorCode::BrokenHyperlinkReference => "broken_hyperlink",
            ErrorCode::BrokenImageReference => "broken_image",
            ErrorCode::MissingNumberingDefinition => "missing_numbering",
            ErrorCode::MissingRequiredPart => "missing_part",
            ErrorCode::CircularStyleInheritance => "circular_style",
            ErrorCode::InvalidStructure => "invalid_structure",
        }
    }
}

/// A validation error.
#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct ValidationError {
    /// Error code.
    pub code: ErrorCode,
    /// Human-readable message.
    pub message: String,
    /// Location in document if known.
    pub location: Option<String>,
}

/// Common fonts that are widely available.
const COMMON_FONTS: &[&str] = &[
    "Arial",
    "Calibri",
    "Cambria",
    "Comic Sans MS",
    "Consolas",
    "Courier New",
    "Georgia",
    "Helvetica",
    "Impact",
    "Lucida Console",
    "Palatino",
    "Segoe UI",
    "Tahoma",
    "Times New Roman",
    "Trebuchet MS",
    "Verdana",
    // CJK fonts
    "MS Gothic",
    "MS Mincho",
    "SimSun",
    "SimHei",
    // Symbol fonts
    "Symbol",
    "Wingdings",
    "Webdings",
];

/// Validate a parsed document.
pub fn validate_document<R: Read + Seek>(doc: &Document<R>) -> ValidationResult {
    let mut result = ValidationResult::valid();

    // Get defined style IDs
    let defined_styles: HashSet<String> = doc
        .styles()
        .style
        .iter()
        .filter_map(|s| s.style_id.as_ref().map(|id| id.to_string()))
        .collect();

    // Check style references
    validate_style_references(doc, &defined_styles, &mut result);

    // Check for uncommon fonts
    validate_fonts(doc, &mut result);

    // Note: More validation checks can be added here as we enhance the parser
    // - validate_hyperlink_references
    // - validate_image_references
    // - validate_numbering_references

    result
}

/// Validate that all style references resolve to defined styles.
fn validate_style_references<R: Read + Seek>(
    doc: &Document<R>,
    defined_styles: &HashSet<String>,
    result: &mut ValidationResult,
) {
    let mut checked_styles: HashSet<String> = HashSet::new();

    for block in &doc.body().block_content {
        match block.as_ref() {
            BlockContent::P(para) => {
                // NOTE: Paragraph style references are in CTPPrBase (not yet flattened).
                // Checking character style references in runs instead.

                // Check run styles
                for run in para.runs() {
                    if let Some(props) = run.properties()
                        && let Some(style) = &props.run_style
                        && !defined_styles.contains(&*style.value)
                        && !checked_styles.contains(&*style.value)
                    {
                        result.add_warning(ValidationWarning {
                            code: WarningCode::UnresolvedStyleReference,
                            message: format!("Character style '{}' is not defined", style.value),
                            location: None,
                        });
                        checked_styles.insert(style.value.clone());
                    }
                }
            }
            BlockContent::Tbl(table) => {
                // Check styles in table cells
                for row in table.rows() {
                    for cell in row.cells() {
                        for para in cell.paragraphs() {
                            for run in para.runs() {
                                if let Some(props) = run.properties()
                                    && let Some(style) = &props.run_style
                                    && !defined_styles.contains(&*style.value)
                                    && !checked_styles.contains(&*style.value)
                                {
                                    result.add_warning(ValidationWarning {
                                        code: WarningCode::UnresolvedStyleReference,
                                        message: format!(
                                            "Character style '{}' is not defined",
                                            style.value
                                        ),
                                        location: Some("table cell".to_string()),
                                    });
                                    checked_styles.insert(style.value.clone());
                                }
                            }
                        }
                    }
                }
            }
            _ => {}
        }
    }
}

/// Validate fonts used in the document.
fn validate_fonts<R: Read + Seek>(doc: &Document<R>, result: &mut ValidationResult) {
    let mut checked_fonts: HashSet<String> = HashSet::new();
    let common_fonts_lower: HashSet<String> =
        COMMON_FONTS.iter().map(|f| f.to_lowercase()).collect();

    for block in &doc.body().block_content {
        if let BlockContent::P(para) = block.as_ref() {
            for run in para.runs() {
                if let Some(props) = run.properties()
                    && let Some(font) = props.font_ascii()
                {
                    let font_lower = font.to_lowercase();
                    if !common_fonts_lower.contains(&font_lower)
                        && !checked_fonts.contains(&font_lower)
                    {
                        result.add_warning(ValidationWarning {
                            code: WarningCode::UncommonFont,
                            message: format!("Font '{}' may not be available on all systems", font),
                            location: None,
                        });
                        checked_fonts.insert(font_lower);
                    }
                }
            }
        }
    }
}

/// Aggregate validation statistics across a corpus.
#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct CorpusValidationStats {
    /// Total documents validated.
    pub total_documents: u64,
    /// Documents that passed validation.
    pub valid_documents: u64,
    /// Documents with warnings only.
    pub documents_with_warnings: u64,
    /// Documents with errors.
    pub documents_with_errors: u64,
    /// Warning count by code.
    pub warnings_by_code: std::collections::HashMap<String, u64>,
    /// Error count by code.
    pub errors_by_code: std::collections::HashMap<String, u64>,
}

impl CorpusValidationStats {
    /// Create new empty stats.
    pub fn new() -> Self {
        Self::default()
    }

    /// Add a document's validation result to the aggregate stats.
    pub fn add(&mut self, result: &ValidationResult) {
        self.total_documents += 1;

        if result.is_valid {
            self.valid_documents += 1;
        }

        if !result.warnings.is_empty() && result.errors.is_empty() {
            self.documents_with_warnings += 1;
        }

        if !result.errors.is_empty() {
            self.documents_with_errors += 1;
        }

        for warning in &result.warnings {
            let key = warning.code.as_str().to_string();
            *self.warnings_by_code.entry(key).or_insert(0) += 1;
        }

        for error in &result.errors {
            let key = error.code.as_str().to_string();
            *self.errors_by_code.entry(key).or_insert(0) += 1;
        }
    }

    /// Get percentage of valid documents.
    pub fn valid_percentage(&self) -> f64 {
        if self.total_documents == 0 {
            0.0
        } else {
            (self.valid_documents as f64 / self.total_documents as f64) * 100.0
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_validation_result_valid() {
        let result = ValidationResult::valid();
        assert!(result.is_valid);
        assert!(result.warnings.is_empty());
        assert!(result.errors.is_empty());
    }

    #[test]
    fn test_add_warning() {
        let mut result = ValidationResult::valid();
        result.add_warning(ValidationWarning {
            code: WarningCode::UnresolvedStyleReference,
            message: "Test warning".to_string(),
            location: None,
        });

        assert!(result.is_valid); // Warnings don't make it invalid
        assert_eq!(result.warnings.len(), 1);
        assert_eq!(result.issue_count(), 1);
    }

    #[test]
    fn test_add_error() {
        let mut result = ValidationResult::valid();
        result.add_error(ValidationError {
            code: ErrorCode::BrokenImageReference,
            message: "Test error".to_string(),
            location: None,
        });

        assert!(!result.is_valid); // Errors make it invalid
        assert_eq!(result.errors.len(), 1);
    }

    #[test]
    fn test_corpus_stats() {
        let mut stats = CorpusValidationStats::new();

        // Add a valid document
        let valid_result = ValidationResult::valid();
        stats.add(&valid_result);

        // Add a document with warnings
        let mut warning_result = ValidationResult::valid();
        warning_result.add_warning(ValidationWarning {
            code: WarningCode::UncommonFont,
            message: "Test".to_string(),
            location: None,
        });
        stats.add(&warning_result);

        // Add a document with errors
        let mut error_result = ValidationResult::valid();
        error_result.add_error(ValidationError {
            code: ErrorCode::BrokenImageReference,
            message: "Test".to_string(),
            location: None,
        });
        stats.add(&error_result);

        assert_eq!(stats.total_documents, 3);
        assert_eq!(stats.valid_documents, 2); // First two are valid
        assert_eq!(stats.documents_with_warnings, 1);
        assert_eq!(stats.documents_with_errors, 1);
    }
}
