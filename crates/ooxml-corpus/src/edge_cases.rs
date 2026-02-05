//! Edge case detection for OOXML documents.
//!
//! This module identifies unusual patterns, structures, and features
//! that may stress-test parsers or represent interesting test fixtures.

use crate::DocumentFeatures;
use serde::{Deserialize, Serialize};

/// Types of edge cases we detect.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash, Serialize, Deserialize)]
pub enum EdgeCaseType {
    // Structural edge cases
    /// Table nested more than 2 levels deep.
    DeeplyNestedTable,
    /// Paragraph with more than 100 runs.
    ManyRunsInParagraph,
    /// Document with more than 1000 paragraphs.
    VeryLargeParagraphCount,
    /// Document with more than 50 tables.
    ManyTables,
    /// List nested more than 5 levels deep.
    DeeplyNestedList,

    // Formatting edge cases
    /// Uses more than 20 unique fonts.
    ManyUniqueFonts,
    /// Uses more than 50 unique colors.
    ManyUniqueColors,
    /// Very large font size (> 72pt).
    VeryLargeFontSize,

    // Content edge cases
    /// Contains embedded images.
    HasEmbeddedImages,
    /// Contains external hyperlinks.
    HasExternalHyperlinks,
    /// Contains page breaks.
    HasPageBreaks,

    // Style edge cases
    /// Uses more than 50 unique styles.
    ManyUniqueStyles,
    /// Missing style references (style ID not found).
    MissingStyleReference,

    // Producer quirks
    /// Heavy use of revision IDs (rsid attributes).
    HeavyRsidUsage,
    /// Contains custom XML parts.
    HasCustomXml,
    /// Contains form fields or content controls.
    HasContentControls,
}

impl EdgeCaseType {
    /// Get a short string identifier.
    pub fn as_str(&self) -> &'static str {
        match self {
            EdgeCaseType::DeeplyNestedTable => "deeply_nested_table",
            EdgeCaseType::ManyRunsInParagraph => "many_runs_in_paragraph",
            EdgeCaseType::VeryLargeParagraphCount => "very_large_paragraph_count",
            EdgeCaseType::ManyTables => "many_tables",
            EdgeCaseType::DeeplyNestedList => "deeply_nested_list",
            EdgeCaseType::ManyUniqueFonts => "many_unique_fonts",
            EdgeCaseType::ManyUniqueColors => "many_unique_colors",
            EdgeCaseType::VeryLargeFontSize => "very_large_font_size",
            EdgeCaseType::HasEmbeddedImages => "has_embedded_images",
            EdgeCaseType::HasExternalHyperlinks => "has_external_hyperlinks",
            EdgeCaseType::HasPageBreaks => "has_page_breaks",
            EdgeCaseType::ManyUniqueStyles => "many_unique_styles",
            EdgeCaseType::MissingStyleReference => "missing_style_reference",
            EdgeCaseType::HeavyRsidUsage => "heavy_rsid_usage",
            EdgeCaseType::HasCustomXml => "has_custom_xml",
            EdgeCaseType::HasContentControls => "has_content_controls",
        }
    }

    /// Get a human-readable description.
    pub fn description(&self) -> &'static str {
        match self {
            EdgeCaseType::DeeplyNestedTable => "Table nested more than 2 levels deep",
            EdgeCaseType::ManyRunsInParagraph => "Paragraph with more than 100 runs",
            EdgeCaseType::VeryLargeParagraphCount => "Document with more than 1000 paragraphs",
            EdgeCaseType::ManyTables => "Document with more than 50 tables",
            EdgeCaseType::DeeplyNestedList => "List nested more than 5 levels deep",
            EdgeCaseType::ManyUniqueFonts => "Uses more than 20 unique fonts",
            EdgeCaseType::ManyUniqueColors => "Uses more than 50 unique colors",
            EdgeCaseType::VeryLargeFontSize => "Very large font size (> 72pt)",
            EdgeCaseType::HasEmbeddedImages => "Contains embedded images",
            EdgeCaseType::HasExternalHyperlinks => "Contains external hyperlinks",
            EdgeCaseType::HasPageBreaks => "Contains page breaks",
            EdgeCaseType::ManyUniqueStyles => "Uses more than 50 unique styles",
            EdgeCaseType::MissingStyleReference => "References undefined style",
            EdgeCaseType::HeavyRsidUsage => "Heavy use of revision tracking IDs",
            EdgeCaseType::HasCustomXml => "Contains custom XML parts",
            EdgeCaseType::HasContentControls => "Contains form fields or content controls",
        }
    }
}

/// Severity of an edge case.
#[derive(Debug, Clone, Copy, PartialEq, Eq, PartialOrd, Ord, Hash, Serialize, Deserialize)]
pub enum Severity {
    /// Informational - common but notable.
    Info,
    /// Noteworthy - less common, good for testing.
    Noteworthy,
    /// Unusual - rare pattern worth investigating.
    Unusual,
    /// Rare - very uncommon, high value for testing.
    Rare,
}

impl Severity {
    /// Get a short string identifier.
    pub fn as_str(&self) -> &'static str {
        match self {
            Severity::Info => "info",
            Severity::Noteworthy => "noteworthy",
            Severity::Unusual => "unusual",
            Severity::Rare => "rare",
        }
    }
}

/// A detected edge case in a document.
#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct EdgeCase {
    /// Type of edge case.
    pub pattern_type: EdgeCaseType,
    /// Severity level.
    pub severity: Severity,
    /// Human-readable description with specifics.
    pub description: String,
    /// Optional location info (e.g., "paragraph 42", "table at `/w:body/w:tbl[3]`").
    pub location: Option<String>,
    /// Numeric value if applicable (e.g., nesting depth, count).
    pub value: Option<u32>,
}

/// Edge cases detected in a document.
#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct DocumentEdgeCases {
    /// All edge cases found.
    pub cases: Vec<EdgeCase>,
}

impl DocumentEdgeCases {
    /// Check if any edge cases were found.
    pub fn is_empty(&self) -> bool {
        self.cases.is_empty()
    }

    /// Get count of edge cases.
    pub fn len(&self) -> usize {
        self.cases.len()
    }

    /// Get edge cases by severity.
    pub fn by_severity(&self, severity: Severity) -> Vec<&EdgeCase> {
        self.cases
            .iter()
            .filter(|c| c.severity == severity)
            .collect()
    }

    /// Get edge cases by type.
    pub fn by_type(&self, pattern_type: EdgeCaseType) -> Vec<&EdgeCase> {
        self.cases
            .iter()
            .filter(|c| c.pattern_type == pattern_type)
            .collect()
    }

    /// Check if a specific edge case type is present.
    pub fn has(&self, pattern_type: EdgeCaseType) -> bool {
        self.cases.iter().any(|c| c.pattern_type == pattern_type)
    }

    /// Get the highest severity level present.
    pub fn max_severity(&self) -> Option<Severity> {
        self.cases.iter().map(|c| c.severity).max()
    }
}

/// Thresholds for edge case detection.
#[derive(Debug, Clone)]
pub struct EdgeCaseThresholds {
    /// Minimum table nesting depth to flag.
    pub min_table_nesting: u8,
    /// Minimum runs per paragraph to flag.
    pub min_runs_per_paragraph: u32,
    /// Minimum paragraph count to flag.
    pub min_paragraphs: u32,
    /// Minimum table count to flag.
    pub min_tables: u32,
    /// Minimum unique fonts to flag.
    pub min_unique_fonts: usize,
    /// Minimum unique colors to flag.
    pub min_unique_colors: usize,
    /// Minimum unique styles to flag.
    pub min_unique_styles: usize,
    /// Minimum font size (in half-points) to flag as very large.
    pub min_large_font_size: u32,
}

impl Default for EdgeCaseThresholds {
    fn default() -> Self {
        Self {
            min_table_nesting: 3,
            min_runs_per_paragraph: 100,
            min_paragraphs: 1000,
            min_tables: 50,
            min_unique_fonts: 20,
            min_unique_colors: 50,
            min_unique_styles: 50,
            min_large_font_size: 144, // 72pt in half-points
        }
    }
}

/// Detect edge cases from document features.
pub fn detect_edge_cases(features: &DocumentFeatures) -> DocumentEdgeCases {
    detect_edge_cases_with_thresholds(features, &EdgeCaseThresholds::default())
}

/// Detect edge cases with custom thresholds.
pub fn detect_edge_cases_with_thresholds(
    features: &DocumentFeatures,
    thresholds: &EdgeCaseThresholds,
) -> DocumentEdgeCases {
    let mut cases = Vec::new();

    // Structural checks
    if features.max_table_nesting >= thresholds.min_table_nesting {
        cases.push(EdgeCase {
            pattern_type: EdgeCaseType::DeeplyNestedTable,
            severity: if features.max_table_nesting >= 4 {
                Severity::Rare
            } else {
                Severity::Unusual
            },
            description: format!(
                "Table nesting depth of {} levels",
                features.max_table_nesting
            ),
            location: None,
            value: Some(features.max_table_nesting as u32),
        });
    }

    if features.paragraph_count >= thresholds.min_paragraphs {
        let severity = if features.paragraph_count >= 5000 {
            Severity::Rare
        } else if features.paragraph_count >= 2000 {
            Severity::Unusual
        } else {
            Severity::Noteworthy
        };
        cases.push(EdgeCase {
            pattern_type: EdgeCaseType::VeryLargeParagraphCount,
            severity,
            description: format!("{} paragraphs in document", features.paragraph_count),
            location: None,
            value: Some(features.paragraph_count),
        });
    }

    if features.table_count >= thresholds.min_tables {
        let severity = if features.table_count >= 100 {
            Severity::Rare
        } else {
            Severity::Unusual
        };
        cases.push(EdgeCase {
            pattern_type: EdgeCaseType::ManyTables,
            severity,
            description: format!("{} tables in document", features.table_count),
            location: None,
            value: Some(features.table_count),
        });
    }

    // Formatting checks
    if features.unique_fonts.len() >= thresholds.min_unique_fonts {
        cases.push(EdgeCase {
            pattern_type: EdgeCaseType::ManyUniqueFonts,
            severity: if features.unique_fonts.len() >= 50 {
                Severity::Rare
            } else {
                Severity::Unusual
            },
            description: format!("{} unique fonts used", features.unique_fonts.len()),
            location: None,
            value: Some(features.unique_fonts.len() as u32),
        });
    }

    if features.unique_colors.len() >= thresholds.min_unique_colors {
        cases.push(EdgeCase {
            pattern_type: EdgeCaseType::ManyUniqueColors,
            severity: Severity::Unusual,
            description: format!("{} unique colors used", features.unique_colors.len()),
            location: None,
            value: Some(features.unique_colors.len() as u32),
        });
    }

    // Check for very large font sizes
    if let Some(&max_size) = features.font_sizes.iter().max()
        && max_size >= thresholds.min_large_font_size
    {
        let pt_size = max_size / 2;
        cases.push(EdgeCase {
            pattern_type: EdgeCaseType::VeryLargeFontSize,
            severity: if pt_size >= 144 {
                Severity::Rare
            } else {
                Severity::Unusual
            },
            description: format!("Font size of {}pt used", pt_size),
            location: None,
            value: Some(max_size),
        });
    }

    // Style checks
    let total_styles = features.paragraph_style_refs.len() + features.character_style_refs.len();
    if total_styles >= thresholds.min_unique_styles {
        cases.push(EdgeCase {
            pattern_type: EdgeCaseType::ManyUniqueStyles,
            severity: Severity::Unusual,
            description: format!("{} unique style references", total_styles),
            location: None,
            value: Some(total_styles as u32),
        });
    }

    // Content presence checks (informational)
    if features.image_count > 0 {
        cases.push(EdgeCase {
            pattern_type: EdgeCaseType::HasEmbeddedImages,
            severity: Severity::Info,
            description: format!("{} embedded images", features.image_count),
            location: None,
            value: Some(features.image_count),
        });
    }

    if features.hyperlink_count > 0 {
        cases.push(EdgeCase {
            pattern_type: EdgeCaseType::HasExternalHyperlinks,
            severity: Severity::Info,
            description: format!("{} hyperlinks", features.hyperlink_count),
            location: None,
            value: Some(features.hyperlink_count),
        });
    }

    if features.page_break_count > 0 {
        cases.push(EdgeCase {
            pattern_type: EdgeCaseType::HasPageBreaks,
            severity: Severity::Info,
            description: format!("{} page breaks", features.page_break_count),
            location: None,
            value: Some(features.page_break_count),
        });
    }

    DocumentEdgeCases { cases }
}

/// Aggregate edge case statistics across a corpus.
#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct CorpusEdgeCaseStats {
    /// Total documents analyzed.
    pub total_documents: u64,
    /// Documents with at least one edge case.
    pub documents_with_edge_cases: u64,
    /// Count by edge case type.
    pub by_type: std::collections::HashMap<String, u64>,
    /// Count by severity.
    pub by_severity: std::collections::HashMap<String, u64>,
    /// Maximum values seen for numeric edge cases.
    pub max_values: std::collections::HashMap<String, u32>,
}

impl CorpusEdgeCaseStats {
    /// Create new empty stats.
    pub fn new() -> Self {
        Self::default()
    }

    /// Add a document's edge cases to the aggregate stats.
    pub fn add(&mut self, edge_cases: &DocumentEdgeCases) {
        self.total_documents += 1;

        if !edge_cases.is_empty() {
            self.documents_with_edge_cases += 1;
        }

        for case in &edge_cases.cases {
            let type_key = case.pattern_type.as_str().to_string();
            *self.by_type.entry(type_key.clone()).or_insert(0) += 1;

            let severity_key = case.severity.as_str().to_string();
            *self.by_severity.entry(severity_key).or_insert(0) += 1;

            if let Some(value) = case.value {
                let current_max = self.max_values.entry(type_key).or_insert(0);
                *current_max = (*current_max).max(value);
            }
        }
    }

    /// Get percentage of documents with edge cases.
    pub fn edge_case_percentage(&self) -> f64 {
        if self.total_documents == 0 {
            0.0
        } else {
            (self.documents_with_edge_cases as f64 / self.total_documents as f64) * 100.0
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_detect_large_document() {
        let features = DocumentFeatures {
            paragraph_count: 1500,
            ..Default::default()
        };

        let edge_cases = detect_edge_cases(&features);
        assert!(edge_cases.has(EdgeCaseType::VeryLargeParagraphCount));

        let case = edge_cases.by_type(EdgeCaseType::VeryLargeParagraphCount)[0];
        assert_eq!(case.severity, Severity::Noteworthy);
        assert_eq!(case.value, Some(1500));
    }

    #[test]
    fn test_detect_deeply_nested_table() {
        let features = DocumentFeatures {
            max_table_nesting: 4,
            ..Default::default()
        };

        let edge_cases = detect_edge_cases(&features);
        assert!(edge_cases.has(EdgeCaseType::DeeplyNestedTable));
        assert_eq!(edge_cases.max_severity(), Some(Severity::Rare));
    }

    #[test]
    fn test_detect_many_fonts() {
        let mut features = DocumentFeatures::default();
        for i in 0..25 {
            features.unique_fonts.insert(format!("Font{}", i));
        }

        let edge_cases = detect_edge_cases(&features);
        assert!(edge_cases.has(EdgeCaseType::ManyUniqueFonts));
    }

    #[test]
    fn test_info_severity_for_images() {
        let features = DocumentFeatures {
            image_count: 5,
            ..Default::default()
        };

        let edge_cases = detect_edge_cases(&features);
        assert!(edge_cases.has(EdgeCaseType::HasEmbeddedImages));

        let case = edge_cases.by_type(EdgeCaseType::HasEmbeddedImages)[0];
        assert_eq!(case.severity, Severity::Info);
    }

    #[test]
    fn test_corpus_stats_add() {
        let mut stats = CorpusEdgeCaseStats::new();

        let features = DocumentFeatures {
            paragraph_count: 2000,
            image_count: 3,
            ..Default::default()
        };

        let edge_cases = detect_edge_cases(&features);
        stats.add(&edge_cases);

        assert_eq!(stats.total_documents, 1);
        assert_eq!(stats.documents_with_edge_cases, 1);
        assert!(stats.by_type.contains_key("very_large_paragraph_count"));
        assert!(stats.by_type.contains_key("has_embedded_images"));
    }

    #[test]
    fn test_custom_thresholds() {
        let features = DocumentFeatures {
            paragraph_count: 500,
            ..Default::default()
        };

        // Default threshold is 1000, so no edge case
        let edge_cases = detect_edge_cases(&features);
        assert!(!edge_cases.has(EdgeCaseType::VeryLargeParagraphCount));

        // Custom threshold of 400
        let thresholds = EdgeCaseThresholds {
            min_paragraphs: 400,
            ..Default::default()
        };
        let edge_cases = detect_edge_cases_with_thresholds(&features, &thresholds);
        assert!(edge_cases.has(EdgeCaseType::VeryLargeParagraphCount));
    }
}
