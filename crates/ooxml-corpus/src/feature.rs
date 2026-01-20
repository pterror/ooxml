//! Feature detection for OOXML documents.
//!
//! This module provides tools to analyze what features a document uses,
//! enabling corpus-wide statistics and pattern detection.

use ooxml_wml::{Alignment, BlockContent, Document, ParagraphContent};
use serde::{Deserialize, Serialize};
use std::collections::{HashMap, HashSet};
use std::io::{Read, Seek};

/// Features detected in a single document.
#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct DocumentFeatures {
    // Structure counts
    /// Total number of paragraphs (including those in tables).
    pub paragraph_count: u32,
    /// Total number of runs across all paragraphs.
    pub run_count: u32,
    /// Number of tables in the document.
    pub table_count: u32,
    /// Maximum table nesting depth (1 = simple table, 2+ = nested).
    pub max_table_nesting: u8,
    /// Number of hyperlinks.
    pub hyperlink_count: u32,
    /// Number of images.
    pub image_count: u32,
    /// Number of page breaks.
    pub page_break_count: u32,

    // Formatting presence
    /// Document contains bold text.
    pub has_bold: bool,
    /// Document contains italic text.
    pub has_italic: bool,
    /// Document contains underlined text.
    pub has_underline: bool,
    /// Document contains strikethrough text.
    pub has_strike: bool,
    /// Document contains colored text.
    pub has_color: bool,
    /// Document contains explicit font sizes.
    pub has_font_size: bool,
    /// Document contains explicit font names.
    pub has_font_name: bool,

    // Paragraph properties
    /// Document uses paragraph alignment.
    pub has_alignment: bool,
    /// Document uses paragraph spacing.
    pub has_spacing: bool,
    /// Document uses paragraph indentation.
    pub has_indentation: bool,
    /// Document uses numbering/lists.
    pub has_numbering: bool,
    /// Number of list items (paragraphs with numbering).
    pub list_item_count: u32,

    // Style usage
    /// Number of styles defined in the document.
    pub style_count: u32,
    /// Paragraph style IDs referenced by paragraphs.
    pub paragraph_style_refs: HashSet<String>,
    /// Character style IDs referenced by runs.
    pub character_style_refs: HashSet<String>,

    // Collections for detailed analysis
    /// Unique text colors used (hex RGB without #).
    pub unique_colors: HashSet<String>,
    /// Unique font names used.
    pub unique_fonts: HashSet<String>,
    /// Font sizes used (in half-points).
    pub font_sizes: HashSet<u32>,
    /// Alignment types used.
    pub alignment_types: HashSet<String>,
}

/// Aggregate statistics across a corpus.
#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct CorpusFeatureStats {
    /// Total documents analyzed.
    pub total_documents: u64,

    // Feature prevalence (document counts)
    /// Documents containing at least one table.
    pub with_tables: u64,
    /// Documents containing at least one image.
    pub with_images: u64,
    /// Documents containing at least one hyperlink.
    pub with_hyperlinks: u64,
    /// Documents using numbering/lists.
    pub with_lists: u64,
    /// Documents using explicit paragraph styles.
    pub with_paragraph_styles: u64,
    /// Documents using character styles.
    pub with_character_styles: u64,
    /// Documents using text color.
    pub with_color: u64,
    /// Documents using bold text.
    pub with_bold: u64,
    /// Documents using italic text.
    pub with_italic: u64,
    /// Documents using paragraph alignment.
    pub with_alignment: u64,

    // Aggregate counts
    /// Total paragraphs across all documents.
    pub total_paragraphs: u64,
    /// Total tables across all documents.
    pub total_tables: u64,
    /// Total images across all documents.
    pub total_images: u64,
    /// Total hyperlinks across all documents.
    pub total_hyperlinks: u64,

    // Maximums
    /// Maximum paragraphs in any single document.
    pub max_paragraphs: u32,
    /// Maximum tables in any single document.
    pub max_tables: u32,
    /// Maximum images in any single document.
    pub max_images: u32,
    /// Maximum table nesting depth seen.
    pub max_table_nesting: u8,

    // Distributions
    /// Color usage across corpus (color -> document count).
    pub color_usage: HashMap<String, u64>,
    /// Font usage across corpus (font name -> document count).
    pub font_usage: HashMap<String, u64>,
    /// Style usage across corpus (style ID -> document count).
    pub style_usage: HashMap<String, u64>,
    /// Alignment type usage (alignment -> document count).
    pub alignment_usage: HashMap<String, u64>,
}

impl CorpusFeatureStats {
    /// Create new empty stats.
    pub fn new() -> Self {
        Self::default()
    }

    /// Add a document's features to the aggregate stats.
    pub fn add(&mut self, features: &DocumentFeatures) {
        self.total_documents += 1;

        // Prevalence
        if features.table_count > 0 {
            self.with_tables += 1;
        }
        if features.image_count > 0 {
            self.with_images += 1;
        }
        if features.hyperlink_count > 0 {
            self.with_hyperlinks += 1;
        }
        if features.has_numbering {
            self.with_lists += 1;
        }
        if !features.paragraph_style_refs.is_empty() {
            self.with_paragraph_styles += 1;
        }
        if !features.character_style_refs.is_empty() {
            self.with_character_styles += 1;
        }
        if features.has_color {
            self.with_color += 1;
        }
        if features.has_bold {
            self.with_bold += 1;
        }
        if features.has_italic {
            self.with_italic += 1;
        }
        if features.has_alignment {
            self.with_alignment += 1;
        }

        // Totals
        self.total_paragraphs += features.paragraph_count as u64;
        self.total_tables += features.table_count as u64;
        self.total_images += features.image_count as u64;
        self.total_hyperlinks += features.hyperlink_count as u64;

        // Maximums
        self.max_paragraphs = self.max_paragraphs.max(features.paragraph_count);
        self.max_tables = self.max_tables.max(features.table_count);
        self.max_images = self.max_images.max(features.image_count);
        self.max_table_nesting = self.max_table_nesting.max(features.max_table_nesting);

        // Distributions
        for color in &features.unique_colors {
            *self.color_usage.entry(color.clone()).or_insert(0) += 1;
        }
        for font in &features.unique_fonts {
            *self.font_usage.entry(font.clone()).or_insert(0) += 1;
        }
        for style in &features.paragraph_style_refs {
            *self.style_usage.entry(style.clone()).or_insert(0) += 1;
        }
        for style in &features.character_style_refs {
            *self.style_usage.entry(style.clone()).or_insert(0) += 1;
        }
        for alignment in &features.alignment_types {
            *self.alignment_usage.entry(alignment.clone()).or_insert(0) += 1;
        }
    }

    /// Calculate percentage of documents with a given feature.
    pub fn percentage(&self, count: u64) -> f64 {
        if self.total_documents == 0 {
            0.0
        } else {
            (count as f64 / self.total_documents as f64) * 100.0
        }
    }
}

/// Extract features from a parsed document.
pub fn extract_features<R: Read + Seek>(doc: &Document<R>) -> DocumentFeatures {
    let mut features = DocumentFeatures {
        style_count: doc.styles().iter().count() as u32,
        ..Default::default()
    };

    // Process body content
    for block in doc.body().content() {
        process_block(block, &mut features, 0);
    }

    features
}

/// Process a block-level element (paragraph or table).
fn process_block(block: &BlockContent, features: &mut DocumentFeatures, table_depth: u8) {
    match block {
        BlockContent::Paragraph(para) => {
            features.paragraph_count += 1;

            // Process paragraph properties
            if let Some(props) = para.properties() {
                if let Some(style) = &props.style {
                    features.paragraph_style_refs.insert(style.clone());
                }

                if props.numbering.is_some() {
                    features.has_numbering = true;
                    features.list_item_count += 1;
                }

                if let Some(alignment) = props.alignment {
                    features.has_alignment = true;
                    features
                        .alignment_types
                        .insert(alignment_to_string(alignment));
                }

                if props.spacing_before.is_some()
                    || props.spacing_after.is_some()
                    || props.spacing_line.is_some()
                {
                    features.has_spacing = true;
                }

                if props.indent_left.is_some()
                    || props.indent_right.is_some()
                    || props.indent_first_line.is_some()
                    || props.indent_hanging.is_some()
                {
                    features.has_indentation = true;
                }
            }

            // Process paragraph content
            for content in para.content() {
                match content {
                    ParagraphContent::Run(run) => {
                        process_run(run, features);
                    }
                    ParagraphContent::Hyperlink(link) => {
                        features.hyperlink_count += 1;
                        for run in link.runs() {
                            process_run(run, features);
                        }
                    }
                }
            }
        }
        BlockContent::Table(table) => {
            features.table_count += 1;
            let new_depth = table_depth + 1;
            features.max_table_nesting = features.max_table_nesting.max(new_depth);

            // Process table cells (which contain paragraphs)
            for row in table.rows() {
                for cell in row.cells() {
                    for para in cell.paragraphs() {
                        // Create a temporary block to reuse process_block
                        process_block(&BlockContent::Paragraph(para.clone()), features, new_depth);
                    }
                }
            }
        }
    }
}

/// Process a run and update features.
fn process_run(run: &ooxml_wml::Run, features: &mut DocumentFeatures) {
    features.run_count += 1;

    if run.has_page_break() {
        features.page_break_count += 1;
    }

    // Check for images
    for drawing in run.drawings() {
        features.image_count += drawing.images().len() as u32;
    }

    // Check run properties
    if let Some(props) = run.properties() {
        if props.bold {
            features.has_bold = true;
        }
        if props.italic {
            features.has_italic = true;
        }
        if props.underline.is_some() {
            features.has_underline = true;
        }
        if props.strike {
            features.has_strike = true;
        }

        if let Some(color) = &props.color {
            features.has_color = true;
            features.unique_colors.insert(color.clone());
        }

        if let Some(size) = props.size {
            features.has_font_size = true;
            features.font_sizes.insert(size);
        }

        if let Some(font) = &props.font {
            features.has_font_name = true;
            features.unique_fonts.insert(font.clone());
        }

        if let Some(style) = &props.style {
            features.character_style_refs.insert(style.clone());
        }
    }
}

/// Convert alignment enum to string for storage.
fn alignment_to_string(alignment: Alignment) -> String {
    alignment.as_str().to_string()
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_corpus_stats_add() {
        let mut stats = CorpusFeatureStats::new();

        let mut features = DocumentFeatures {
            paragraph_count: 10,
            table_count: 2,
            has_bold: true,
            ..Default::default()
        };
        features.unique_colors.insert("FF0000".to_string());

        stats.add(&features);

        assert_eq!(stats.total_documents, 1);
        assert_eq!(stats.total_paragraphs, 10);
        assert_eq!(stats.total_tables, 2);
        assert_eq!(stats.with_bold, 1);
        assert_eq!(stats.with_tables, 1);
        assert_eq!(stats.color_usage.get("FF0000"), Some(&1));
    }

    #[test]
    fn test_percentage() {
        let mut stats = CorpusFeatureStats::new();
        stats.total_documents = 100;
        stats.with_tables = 25;

        assert!((stats.percentage(stats.with_tables) - 25.0).abs() < 0.001);
    }
}
