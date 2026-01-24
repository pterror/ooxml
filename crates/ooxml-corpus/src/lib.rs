//! Corpus analysis library for OOXML documents.
//!
//! This crate provides tools for analyzing collections of DOCX/XLSX/PPTX files
//! to discover patterns, edge cases, and test parser implementations.
//!
//! # Features
//!
//! - **Feature Detection**: Identify what OOXML features each document uses
//! - **Error Analysis**: Structured categorization of parse failures
//! - **Edge Case Detection**: Find unusual patterns and structures
//! - **Validation**: Deep structural and reference integrity checks
//! - **Persistence**: SQLite storage for analysis results
//! - **Fixture Synthesis**: Extract minimal test fixtures from interesting documents

pub mod coverage;
pub mod edge_cases;
pub mod error_analysis;
pub mod feature;
pub mod fixture;
pub mod persistence;
pub mod roundtrip;
pub mod sml_attrs;
pub mod validation;

pub use coverage::{
    CorpusCoverage, CoverageReport, DocumentCoverage, WML_ELEMENTS, extract_coverage_from_file,
    extract_coverage_from_reader,
};
pub use edge_cases::{
    CorpusEdgeCaseStats, DocumentEdgeCases, EdgeCase, EdgeCaseType, Severity, detect_edge_cases,
};
pub use error_analysis::{AnalyzedError, ErrorCategory, analyze_error};
pub use feature::{CorpusFeatureStats, DocumentFeatures, extract_features};
pub use fixture::{
    ExtractionResult, FixtureCriteria, FixtureManifest, batch_extract_fixtures,
    determine_interest_reason, extract_fixture,
};
pub use persistence::{CorpusDatabase, CorpusStats, StoredDocument, StoredError};
pub use roundtrip::{CompareOptions, DifferenceKind, RoundtripResult, XmlDifference, compare_xml};
pub use sml_attrs::{SmlAttrStats, analyze_xlsx_directory, analyze_xlsx_file};
pub use validation::{
    CorpusValidationStats, ErrorCode, ValidationError, ValidationResult, ValidationWarning,
    WarningCode, validate_document,
};
