//! Coverage tracking for OOXML spec elements.
//!
//! This module tracks which XML elements from the OOXML specification
//! appear in analyzed documents, helping identify test coverage gaps.

use serde::{Deserialize, Serialize};
use std::collections::{HashMap, HashSet};
use std::fs;
use std::io::{BufReader, Read, Seek};
use std::path::Path;

/// Known WordprocessingML (WML) elements from ECMA-376.
/// This is a subset of the most common elements.
pub const WML_ELEMENTS: &[&str] = &[
    // Document structure
    "w:document",
    "w:body",
    "w:sectPr",
    // Block-level
    "w:p",
    "w:tbl",
    "w:sdt",
    "w:customXml",
    // Paragraph content
    "w:r",
    "w:hyperlink",
    "w:bookmarkStart",
    "w:bookmarkEnd",
    "w:commentRangeStart",
    "w:commentRangeEnd",
    "w:fldSimple",
    "w:fldChar",
    // Run content
    "w:t",
    "w:tab",
    "w:br",
    "w:sym",
    "w:drawing",
    "w:pict",
    "w:object",
    // Paragraph properties
    "w:pPr",
    "w:pStyle",
    "w:jc",
    "w:ind",
    "w:spacing",
    "w:numPr",
    "w:pBdr",
    "w:shd",
    "w:tabs",
    "w:outlineLvl",
    "w:keepNext",
    "w:keepLines",
    "w:pageBreakBefore",
    "w:widowControl",
    // Run properties
    "w:rPr",
    "w:rStyle",
    "w:b",
    "w:i",
    "w:u",
    "w:strike",
    "w:dstrike",
    "w:color",
    "w:sz",
    "w:szCs",
    "w:rFonts",
    "w:highlight",
    "w:vertAlign",
    "w:caps",
    "w:smallCaps",
    "w:vanish",
    // Table elements
    "w:tr",
    "w:tc",
    "w:tblPr",
    "w:tblGrid",
    "w:gridCol",
    "w:trPr",
    "w:tcPr",
    "w:tblBorders",
    "w:tblCellMar",
    "w:tblW",
    "w:gridSpan",
    "w:vMerge",
    "w:tcBorders",
    "w:shd",
    // List/numbering
    "w:numId",
    "w:ilvl",
    "w:abstractNum",
    "w:num",
    "w:lvl",
    "w:lvlText",
    "w:numFmt",
    "w:start",
    // Styles
    "w:styles",
    "w:style",
    "w:basedOn",
    "w:name",
    "w:docDefaults",
    "w:rPrDefault",
    "w:pPrDefault",
    // Headers/footers
    "w:hdr",
    "w:ftr",
    "w:headerReference",
    "w:footerReference",
    // Footnotes/endnotes
    "w:footnotes",
    "w:footnote",
    "w:endnotes",
    "w:endnote",
    "w:footnoteReference",
    "w:endnoteReference",
    // Comments
    "w:comments",
    "w:comment",
    "w:commentReference",
    // Settings
    "w:settings",
    "w:compat",
    "w:compatSetting",
    // DrawingML elements (commonly seen in DOCX)
    "a:graphic",
    "a:graphicData",
    "wp:inline",
    "wp:anchor",
    "pic:pic",
    "pic:blipFill",
    "a:blip",
    // Relationships
    "Relationship",
    "Relationships",
];

/// Coverage data for a single document.
#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct DocumentCoverage {
    /// Elements found in this document.
    pub elements_found: HashSet<String>,
    /// Element counts (element -> count).
    pub element_counts: HashMap<String, u64>,
    /// Unknown elements (not in our spec list).
    pub unknown_elements: HashSet<String>,
}

impl DocumentCoverage {
    /// Create new empty coverage.
    pub fn new() -> Self {
        Self::default()
    }

    /// Record an element being seen.
    pub fn record_element(&mut self, element: &str) {
        self.elements_found.insert(element.to_string());
        *self.element_counts.entry(element.to_string()).or_insert(0) += 1;
    }

    /// Record an unknown element.
    pub fn record_unknown(&mut self, element: &str) {
        self.unknown_elements.insert(element.to_string());
    }

    /// Get coverage percentage against known elements.
    pub fn coverage_percentage(&self, known_elements: &[&str]) -> f64 {
        if known_elements.is_empty() {
            return 0.0;
        }
        let found = known_elements
            .iter()
            .filter(|e| self.elements_found.contains(**e))
            .count();
        (found as f64 / known_elements.len() as f64) * 100.0
    }
}

/// Aggregate coverage statistics across a corpus.
#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct CorpusCoverage {
    /// Total documents analyzed.
    pub total_documents: u64,
    /// All elements seen across corpus.
    pub elements_seen: HashSet<String>,
    /// Element occurrence counts (element -> document count).
    pub element_doc_counts: HashMap<String, u64>,
    /// Total element instance counts.
    pub element_total_counts: HashMap<String, u64>,
    /// Unknown elements seen.
    pub unknown_elements: HashSet<String>,
    /// Unknown element document counts.
    pub unknown_doc_counts: HashMap<String, u64>,
}

impl CorpusCoverage {
    /// Create new empty coverage.
    pub fn new() -> Self {
        Self::default()
    }

    /// Add a document's coverage to the corpus stats.
    pub fn add(&mut self, doc_coverage: &DocumentCoverage) {
        self.total_documents += 1;

        for element in &doc_coverage.elements_found {
            self.elements_seen.insert(element.clone());
            *self.element_doc_counts.entry(element.clone()).or_insert(0) += 1;
        }

        for (element, count) in &doc_coverage.element_counts {
            *self
                .element_total_counts
                .entry(element.clone())
                .or_insert(0) += count;
        }

        for element in &doc_coverage.unknown_elements {
            self.unknown_elements.insert(element.clone());
            *self.unknown_doc_counts.entry(element.clone()).or_insert(0) += 1;
        }
    }

    /// Get overall coverage percentage against known WML elements.
    pub fn wml_coverage_percentage(&self) -> f64 {
        let found = WML_ELEMENTS
            .iter()
            .filter(|e| self.elements_seen.contains(**e))
            .count();
        (found as f64 / WML_ELEMENTS.len() as f64) * 100.0
    }

    /// Get elements that were never seen.
    pub fn missing_elements(&self) -> Vec<&'static str> {
        WML_ELEMENTS
            .iter()
            .filter(|e| !self.elements_seen.contains(**e))
            .copied()
            .collect()
    }

    /// Get element frequency (document count) sorted by frequency.
    pub fn element_frequency(&self) -> Vec<(&str, u64)> {
        let mut freq: Vec<_> = self
            .element_doc_counts
            .iter()
            .map(|(e, c)| (e.as_str(), *c))
            .collect();
        freq.sort_by(|a, b| b.1.cmp(&a.1));
        freq
    }

    /// Get percentage of documents containing an element.
    pub fn element_percentage(&self, element: &str) -> f64 {
        if self.total_documents == 0 {
            return 0.0;
        }
        let count = self.element_doc_counts.get(element).copied().unwrap_or(0);
        (count as f64 / self.total_documents as f64) * 100.0
    }
}

/// Extract element coverage from a DOCX file by scanning XML.
pub fn extract_coverage_from_file(path: &Path) -> std::io::Result<DocumentCoverage> {
    let file = fs::File::open(path)?;
    let reader = BufReader::new(file);
    extract_coverage_from_reader(reader)
}

/// Extract element coverage from a reader (DOCX zip).
pub fn extract_coverage_from_reader<R: Read + Seek>(
    reader: R,
) -> std::io::Result<DocumentCoverage> {
    let mut coverage = DocumentCoverage::new();
    let known_elements: HashSet<&str> = WML_ELEMENTS.iter().copied().collect();

    let mut archive = match zip::ZipArchive::new(reader) {
        Ok(a) => a,
        Err(e) => return Err(std::io::Error::new(std::io::ErrorKind::InvalidData, e)),
    };

    // Scan XML files in the archive
    let xml_parts = [
        "word/document.xml",
        "word/styles.xml",
        "word/numbering.xml",
        "word/settings.xml",
        "word/footnotes.xml",
        "word/endnotes.xml",
        "word/comments.xml",
        "word/header1.xml",
        "word/header2.xml",
        "word/header3.xml",
        "word/footer1.xml",
        "word/footer2.xml",
        "word/footer3.xml",
        "word/_rels/document.xml.rels",
        "_rels/.rels",
        "[Content_Types].xml",
    ];

    for part_name in xml_parts {
        if let Ok(mut file) = archive.by_name(part_name) {
            let mut content = String::new();
            if file.read_to_string(&mut content).is_ok() {
                extract_elements_from_xml(&content, &known_elements, &mut coverage);
            }
        }
    }

    // Collect names of other XML files first
    let mut other_xml_files: Vec<String> = Vec::new();
    for i in 0..archive.len() {
        if let Ok(file) = archive.by_index(i) {
            let name = file.name().to_string();
            if name.ends_with(".xml") && !xml_parts.contains(&name.as_str()) {
                other_xml_files.push(name);
            }
        }
    }

    // Then process them
    for name in other_xml_files {
        if let Ok(mut file) = archive.by_name(&name) {
            let mut content = String::new();
            if file.read_to_string(&mut content).is_ok() {
                extract_elements_from_xml(&content, &known_elements, &mut coverage);
            }
        }
    }

    Ok(coverage)
}

/// Extract element names from XML content using simple regex-like parsing.
fn extract_elements_from_xml(
    content: &str,
    known_elements: &HashSet<&str>,
    coverage: &mut DocumentCoverage,
) {
    // Simple element extraction: find <prefix:name or <name patterns
    let mut i = 0;
    let bytes = content.as_bytes();

    while i < bytes.len() {
        if bytes[i] == b'<'
            && i + 1 < bytes.len()
            && bytes[i + 1] != b'/'
            && bytes[i + 1] != b'?'
            && bytes[i + 1] != b'!'
        {
            // Start of an element
            let start = i + 1;
            let mut end = start;

            // Find end of element name (space, >, or /)
            while end < bytes.len() {
                let b = bytes[end];
                if b == b' ' || b == b'>' || b == b'/' || b == b'\t' || b == b'\n' || b == b'\r' {
                    break;
                }
                end += 1;
            }

            if end > start {
                let element_name = &content[start..end];

                // Skip XML declarations and processing instructions
                if !element_name.starts_with('?') && !element_name.starts_with('!') {
                    coverage.record_element(element_name);

                    if !known_elements.contains(element_name) {
                        coverage.record_unknown(element_name);
                    }
                }
            }

            i = end;
        } else {
            i += 1;
        }
    }
}

/// Coverage report for display.
#[derive(Debug, Clone, Serialize)]
pub struct CoverageReport {
    /// Total documents analyzed.
    pub total_documents: u64,
    /// Overall WML coverage percentage.
    pub wml_coverage_percent: f64,
    /// Number of WML elements seen.
    pub wml_elements_seen: usize,
    /// Total WML elements in spec.
    pub wml_elements_total: usize,
    /// Missing WML elements.
    pub missing_elements: Vec<String>,
    /// Top elements by frequency.
    pub top_elements: Vec<(String, u64, f64)>, // (name, doc_count, percentage)
    /// Unknown elements found.
    pub unknown_elements: Vec<(String, u64)>, // (name, doc_count)
}

impl CoverageReport {
    /// Generate a report from corpus coverage.
    pub fn from_corpus(coverage: &CorpusCoverage) -> Self {
        let wml_seen: Vec<_> = WML_ELEMENTS
            .iter()
            .filter(|e| coverage.elements_seen.contains(**e))
            .collect();

        let missing: Vec<String> = coverage
            .missing_elements()
            .iter()
            .map(|s| s.to_string())
            .collect();

        let top_elements: Vec<_> = coverage
            .element_frequency()
            .into_iter()
            .take(30)
            .map(|(name, count)| {
                let pct = coverage.element_percentage(name);
                (name.to_string(), count, pct)
            })
            .collect();

        let mut unknown: Vec<_> = coverage
            .unknown_doc_counts
            .iter()
            .map(|(e, c)| (e.clone(), *c))
            .collect();
        unknown.sort_by(|a, b| b.1.cmp(&a.1));
        let unknown = unknown.into_iter().take(20).collect();

        Self {
            total_documents: coverage.total_documents,
            wml_coverage_percent: coverage.wml_coverage_percentage(),
            wml_elements_seen: wml_seen.len(),
            wml_elements_total: WML_ELEMENTS.len(),
            missing_elements: missing,
            top_elements,
            unknown_elements: unknown,
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_document_coverage() {
        let mut coverage = DocumentCoverage::new();
        coverage.record_element("w:p");
        coverage.record_element("w:r");
        coverage.record_element("w:p");

        assert!(coverage.elements_found.contains("w:p"));
        assert!(coverage.elements_found.contains("w:r"));
        assert_eq!(coverage.element_counts.get("w:p"), Some(&2));
    }

    #[test]
    fn test_corpus_coverage() {
        let mut corpus = CorpusCoverage::new();

        let mut doc1 = DocumentCoverage::new();
        doc1.record_element("w:p");
        doc1.record_element("w:r");

        let mut doc2 = DocumentCoverage::new();
        doc2.record_element("w:p");
        doc2.record_element("w:tbl");

        corpus.add(&doc1);
        corpus.add(&doc2);

        assert_eq!(corpus.total_documents, 2);
        assert_eq!(corpus.element_doc_counts.get("w:p"), Some(&2));
        assert_eq!(corpus.element_doc_counts.get("w:r"), Some(&1));
        assert_eq!(corpus.element_doc_counts.get("w:tbl"), Some(&1));
    }

    #[test]
    fn test_extract_elements() {
        let xml =
            r#"<w:document><w:body><w:p><w:r><w:t>Hello</w:t></w:r></w:p></w:body></w:document>"#;
        let known: HashSet<&str> = WML_ELEMENTS.iter().copied().collect();
        let mut coverage = DocumentCoverage::new();

        extract_elements_from_xml(xml, &known, &mut coverage);

        assert!(coverage.elements_found.contains("w:document"));
        assert!(coverage.elements_found.contains("w:body"));
        assert!(coverage.elements_found.contains("w:p"));
        assert!(coverage.elements_found.contains("w:r"));
        assert!(coverage.elements_found.contains("w:t"));
    }

    #[test]
    fn test_missing_elements() {
        let corpus = CorpusCoverage::new();
        let missing = corpus.missing_elements();

        // All elements should be missing in empty corpus
        assert_eq!(missing.len(), WML_ELEMENTS.len());
    }

    #[test]
    fn test_coverage_percentage() {
        let mut corpus = CorpusCoverage::new();

        // Add a doc with half the elements
        let mut doc = DocumentCoverage::new();
        for elem in WML_ELEMENTS.iter().take(WML_ELEMENTS.len() / 2) {
            doc.record_element(elem);
        }
        corpus.add(&doc);

        let coverage = corpus.wml_coverage_percentage();
        assert!(coverage > 45.0 && coverage < 55.0); // Should be around 50%
    }
}
