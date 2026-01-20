//! Fixture synthesis for extracting minimal test cases from documents.
//!
//! This module provides tools to extract interesting documents from a corpus
//! and create minimal, focused test fixtures with metadata.

use crate::{DocumentEdgeCases, DocumentFeatures, EdgeCaseType, Severity};
use serde::{Deserialize, Serialize};
use std::fs;
use std::io::{Read, Write};
use std::path::Path;

/// Metadata about an extracted fixture.
#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct FixtureManifest {
    /// Original source file path.
    pub source_path: String,
    /// Corpus the document came from.
    pub corpus: String,
    /// Why this fixture is interesting.
    pub reason: String,
    /// Edge cases present in this document.
    pub edge_cases: Vec<String>,
    /// Key features of this document.
    pub features: FixtureFeaturesummary,
    /// Tags for categorization.
    pub tags: Vec<String>,
    /// Optional notes about the fixture.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub notes: Option<String>,
}

/// Summary of document features for the manifest.
#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct FixtureFeaturesummary {
    pub paragraph_count: u32,
    pub table_count: u32,
    pub image_count: u32,
    pub hyperlink_count: u32,
    pub has_lists: bool,
    pub has_formatting: bool,
}

impl From<&DocumentFeatures> for FixtureFeaturesummary {
    fn from(f: &DocumentFeatures) -> Self {
        Self {
            paragraph_count: f.paragraph_count,
            table_count: f.table_count,
            image_count: f.image_count,
            hyperlink_count: f.hyperlink_count,
            has_lists: f.has_numbering,
            has_formatting: f.has_bold || f.has_italic || f.has_color,
        }
    }
}

/// Criteria for selecting documents as fixtures.
#[derive(Debug, Clone, Default)]
pub struct FixtureCriteria {
    /// Minimum severity of edge cases to include.
    pub min_severity: Option<Severity>,
    /// Specific edge case types to look for.
    pub edge_case_types: Vec<EdgeCaseType>,
    /// Include documents with tables.
    pub has_tables: bool,
    /// Include documents with images.
    pub has_images: bool,
    /// Include documents with hyperlinks.
    pub has_hyperlinks: bool,
    /// Include documents with lists.
    pub has_lists: bool,
    /// Maximum paragraph count (for "minimal" fixtures).
    pub max_paragraphs: Option<u32>,
}

impl FixtureCriteria {
    /// Create criteria for "interesting" documents (unusual edge cases).
    pub fn interesting() -> Self {
        Self {
            min_severity: Some(Severity::Unusual),
            ..Default::default()
        }
    }

    /// Create criteria for documents with specific features.
    pub fn with_feature(feature: &str) -> Self {
        match feature {
            "tables" => Self {
                has_tables: true,
                ..Default::default()
            },
            "images" => Self {
                has_images: true,
                ..Default::default()
            },
            "hyperlinks" => Self {
                has_hyperlinks: true,
                ..Default::default()
            },
            "lists" => Self {
                has_lists: true,
                ..Default::default()
            },
            _ => Self::default(),
        }
    }

    /// Create criteria for minimal documents (good for unit tests).
    pub fn minimal() -> Self {
        Self {
            max_paragraphs: Some(50),
            ..Default::default()
        }
    }

    /// Check if a document matches the criteria.
    pub fn matches(&self, features: &DocumentFeatures, edge_cases: &DocumentEdgeCases) -> bool {
        // Check max paragraphs
        if let Some(max) = self.max_paragraphs
            && features.paragraph_count > max
        {
            return false;
        }

        // Check feature requirements
        if self.has_tables && features.table_count == 0 {
            return false;
        }
        if self.has_images && features.image_count == 0 {
            return false;
        }
        if self.has_hyperlinks && features.hyperlink_count == 0 {
            return false;
        }
        if self.has_lists && !features.has_numbering {
            return false;
        }

        // Check edge case severity
        if let Some(min_sev) = self.min_severity {
            let has_severity = edge_cases.cases.iter().any(|c| c.severity >= min_sev);
            if !has_severity {
                return false;
            }
        }

        // Check specific edge case types
        if !self.edge_case_types.is_empty() {
            let has_type = self.edge_case_types.iter().any(|t| edge_cases.has(*t));
            if !has_type {
                return false;
            }
        }

        true
    }
}

/// Result of extracting a fixture.
#[derive(Debug)]
pub struct ExtractionResult {
    /// Path where the fixture was saved.
    pub output_path: String,
    /// The manifest that was generated.
    pub manifest: FixtureManifest,
}

/// Extract a document as a test fixture.
///
/// Copies the document to the output directory and creates a manifest.yaml
/// with metadata about why the document is interesting.
pub fn extract_fixture(
    source_path: &Path,
    output_dir: &Path,
    corpus_name: &str,
    features: &DocumentFeatures,
    edge_cases: &DocumentEdgeCases,
    reason: &str,
) -> std::io::Result<ExtractionResult> {
    // Create output directory if needed
    fs::create_dir_all(output_dir)?;

    // Determine output filename (use original name or generate one)
    let source_name = source_path
        .file_name()
        .and_then(|n| n.to_str())
        .unwrap_or("document.docx");

    let output_path = output_dir.join(source_name);

    // Copy the document
    fs::copy(source_path, &output_path)?;

    // Generate manifest
    let manifest = FixtureManifest {
        source_path: source_path.display().to_string(),
        corpus: corpus_name.to_string(),
        reason: reason.to_string(),
        edge_cases: edge_cases
            .cases
            .iter()
            .map(|c| c.pattern_type.as_str().to_string())
            .collect(),
        features: FixtureFeaturesummary::from(features),
        tags: generate_tags(features, edge_cases),
        notes: None,
    };

    // Write manifest
    let manifest_path = output_dir.join("manifest.yaml");
    let manifest_yaml = serde_yaml::to_string(&manifest).map_err(std::io::Error::other)?;
    fs::write(&manifest_path, manifest_yaml)?;

    Ok(ExtractionResult {
        output_path: output_path.display().to_string(),
        manifest,
    })
}

/// Extract a fixture from a reader (e.g., from a ZIP archive).
pub fn extract_fixture_from_reader<R: Read>(
    mut reader: R,
    output_dir: &Path,
    filename: &str,
    corpus_name: &str,
    features: &DocumentFeatures,
    edge_cases: &DocumentEdgeCases,
    reason: &str,
) -> std::io::Result<ExtractionResult> {
    // Create output directory if needed
    fs::create_dir_all(output_dir)?;

    let output_path = output_dir.join(filename);

    // Write the document
    let mut output_file = fs::File::create(&output_path)?;
    let mut buffer = Vec::new();
    reader.read_to_end(&mut buffer)?;
    output_file.write_all(&buffer)?;

    // Generate manifest
    let manifest = FixtureManifest {
        source_path: filename.to_string(),
        corpus: corpus_name.to_string(),
        reason: reason.to_string(),
        edge_cases: edge_cases
            .cases
            .iter()
            .map(|c| c.pattern_type.as_str().to_string())
            .collect(),
        features: FixtureFeaturesummary::from(features),
        tags: generate_tags(features, edge_cases),
        notes: None,
    };

    // Write manifest
    let manifest_path = output_dir.join("manifest.yaml");
    let manifest_yaml = serde_yaml::to_string(&manifest).map_err(std::io::Error::other)?;
    fs::write(&manifest_path, manifest_yaml)?;

    Ok(ExtractionResult {
        output_path: output_path.display().to_string(),
        manifest,
    })
}

/// Generate tags based on document features and edge cases.
fn generate_tags(features: &DocumentFeatures, edge_cases: &DocumentEdgeCases) -> Vec<String> {
    let mut tags = Vec::new();

    // Feature-based tags
    if features.table_count > 0 {
        tags.push("tables".to_string());
    }
    if features.image_count > 0 {
        tags.push("images".to_string());
    }
    if features.hyperlink_count > 0 {
        tags.push("hyperlinks".to_string());
    }
    if features.has_numbering {
        tags.push("lists".to_string());
    }
    if features.has_bold || features.has_italic {
        tags.push("formatting".to_string());
    }
    if features.has_color {
        tags.push("colors".to_string());
    }

    // Severity-based tags
    if edge_cases
        .cases
        .iter()
        .any(|c| c.severity == Severity::Rare)
    {
        tags.push("rare".to_string());
    }
    if edge_cases
        .cases
        .iter()
        .any(|c| c.severity == Severity::Unusual)
    {
        tags.push("unusual".to_string());
    }

    // Size-based tags
    if features.paragraph_count <= 10 {
        tags.push("minimal".to_string());
    } else if features.paragraph_count >= 1000 {
        tags.push("large".to_string());
    }

    tags
}

/// Determine why a document is interesting based on its features and edge cases.
pub fn determine_interest_reason(
    features: &DocumentFeatures,
    edge_cases: &DocumentEdgeCases,
) -> Option<String> {
    // Check for rare edge cases first
    for case in &edge_cases.cases {
        if case.severity == Severity::Rare {
            return Some(format!("Rare edge case: {}", case.description));
        }
    }

    // Check for unusual edge cases
    for case in &edge_cases.cases {
        if case.severity == Severity::Unusual {
            return Some(format!("Unusual pattern: {}", case.description));
        }
    }

    // Check for notable combinations
    if features.table_count > 0 && features.image_count > 0 {
        return Some("Contains both tables and images".to_string());
    }

    if features.max_table_nesting > 1 {
        return Some(format!(
            "Nested tables ({} levels)",
            features.max_table_nesting
        ));
    }

    // Check for noteworthy cases
    for case in &edge_cases.cases {
        if case.severity == Severity::Noteworthy {
            return Some(format!("Notable: {}", case.description));
        }
    }

    None
}

/// Batch extract fixtures matching criteria from a list of analyzed documents.
pub fn batch_extract_fixtures(
    documents: &[(String, DocumentFeatures, DocumentEdgeCases)],
    output_base_dir: &Path,
    corpus_name: &str,
    criteria: &FixtureCriteria,
    max_fixtures: usize,
) -> std::io::Result<Vec<ExtractionResult>> {
    let mut results = Vec::new();
    let mut count = 0;

    for (path, features, edge_cases) in documents {
        if count >= max_fixtures {
            break;
        }

        if !criteria.matches(features, edge_cases) {
            continue;
        }

        let reason = determine_interest_reason(features, edge_cases)
            .unwrap_or_else(|| "Matches extraction criteria".to_string());

        // Create a subdirectory for each fixture
        let fixture_name = format!("fixture_{:04}", count);
        let fixture_dir = output_base_dir.join(&fixture_name);

        let source_path = Path::new(path);
        if source_path.exists() {
            match extract_fixture(
                source_path,
                &fixture_dir,
                corpus_name,
                features,
                edge_cases,
                &reason,
            ) {
                Ok(result) => {
                    results.push(result);
                    count += 1;
                }
                Err(e) => {
                    eprintln!("Warning: Failed to extract {}: {}", path, e);
                }
            }
        }
    }

    Ok(results)
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_criteria_matches_tables() {
        let criteria = FixtureCriteria::with_feature("tables");

        let features_with_tables = DocumentFeatures {
            table_count: 5,
            ..Default::default()
        };
        let features_without_tables = DocumentFeatures::default();
        let edge_cases = DocumentEdgeCases::default();

        assert!(criteria.matches(&features_with_tables, &edge_cases));
        assert!(!criteria.matches(&features_without_tables, &edge_cases));
    }

    #[test]
    fn test_criteria_max_paragraphs() {
        let criteria = FixtureCriteria {
            max_paragraphs: Some(100),
            ..Default::default()
        };

        let small_doc = DocumentFeatures {
            paragraph_count: 50,
            ..Default::default()
        };
        let large_doc = DocumentFeatures {
            paragraph_count: 500,
            ..Default::default()
        };
        let edge_cases = DocumentEdgeCases::default();

        assert!(criteria.matches(&small_doc, &edge_cases));
        assert!(!criteria.matches(&large_doc, &edge_cases));
    }

    #[test]
    fn test_generate_tags() {
        let features = DocumentFeatures {
            table_count: 2,
            image_count: 1,
            has_bold: true,
            paragraph_count: 5,
            ..Default::default()
        };
        let edge_cases = DocumentEdgeCases::default();

        let tags = generate_tags(&features, &edge_cases);

        assert!(tags.contains(&"tables".to_string()));
        assert!(tags.contains(&"images".to_string()));
        assert!(tags.contains(&"formatting".to_string()));
        assert!(tags.contains(&"minimal".to_string()));
    }

    #[test]
    fn test_feature_summary_from() {
        let features = DocumentFeatures {
            paragraph_count: 100,
            table_count: 5,
            image_count: 2,
            hyperlink_count: 10,
            has_numbering: true,
            has_bold: true,
            ..Default::default()
        };

        let summary = FixtureFeaturesummary::from(&features);

        assert_eq!(summary.paragraph_count, 100);
        assert_eq!(summary.table_count, 5);
        assert!(summary.has_lists);
        assert!(summary.has_formatting);
    }
}
