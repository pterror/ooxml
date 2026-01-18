//! Word document API.
//!
//! This module provides the main entry point for working with DOCX files.
//!
//! # Example
//!
//! ```ignore
//! use ooxml_wml::Document;
//!
//! let doc = Document::open("document.docx")?;
//! for para in doc.body().paragraphs() {
//!     for run in para.runs() {
//!         print!("{}", run.text());
//!     }
//!     println!();
//! }
//! ```

use crate::error::{Error, Result};
use crate::styles::{Styles, merge_run_properties};
use ooxml::{Package, rel_type};
use quick_xml::Reader;
use quick_xml::events::Event;
use std::fs::File;
use std::io::{BufReader, Read, Seek};
use std::path::Path;

/// A Word document (.docx file).
///
/// This is the main entry point for reading and writing Word documents.
pub struct Document<R> {
    package: Package<R>,
    body: Body,
    styles: Styles,
}

impl Document<BufReader<File>> {
    /// Open a Word document from a file path.
    pub fn open<P: AsRef<Path>>(path: P) -> Result<Self> {
        let file = File::open(path)?;
        let reader = BufReader::new(file);
        Self::from_reader(reader)
    }
}

impl<R: Read + Seek> Document<R> {
    /// Open a Word document from a reader.
    pub fn from_reader(reader: R) -> Result<Self> {
        let mut package = Package::open(reader)?;

        // Find the main document part via relationships
        let rels = package.read_relationships()?;
        let doc_rel = rels
            .get_by_type(rel_type::OFFICE_DOCUMENT)
            .ok_or_else(|| Error::MissingPart("main document relationship".into()))?;

        // Parse the document XML
        let doc_xml = package.read_part(&doc_rel.target)?;
        let body = parse_document(&doc_xml)?;

        // Load styles if available
        let styles = if let Some(styles_rel) = rels.get_by_type(rel_type::STYLES) {
            let styles_xml = package.read_part(&styles_rel.target)?;
            Styles::parse(&styles_xml[..])?
        } else {
            Styles::new()
        };

        Ok(Self {
            package,
            body,
            styles,
        })
    }

    /// Get the document body.
    pub fn body(&self) -> &Body {
        &self.body
    }

    /// Get a mutable reference to the document body.
    pub fn body_mut(&mut self) -> &mut Body {
        &mut self.body
    }

    /// Get the underlying package.
    pub fn package(&self) -> &Package<R> {
        &self.package
    }

    /// Get a mutable reference to the underlying package.
    pub fn package_mut(&mut self) -> &mut Package<R> {
        &mut self.package
    }

    /// Get the document styles.
    pub fn styles(&self) -> &Styles {
        &self.styles
    }

    /// Resolve effective run properties for a run, combining direct and style formatting.
    ///
    /// This merges: default run properties -> paragraph style -> character style -> direct formatting.
    pub fn resolve_run_formatting(&self, para: &Paragraph, run: &Run) -> RunProperties {
        let mut props = self.styles.default_run().clone();

        // Apply paragraph style's run properties
        if let Some(pstyle) = para.properties().and_then(|p| p.style.as_ref()) {
            let style_props = self.styles.resolve_run_properties(pstyle);
            merge_run_properties(&mut props, &style_props);
        }

        // Apply character style
        if let Some(rstyle) = run.properties().and_then(|p| p.style.as_ref()) {
            let style_props = self.styles.resolve_run_properties(rstyle);
            merge_run_properties(&mut props, &style_props);
        }

        // Apply direct formatting
        if let Some(direct) = run.properties() {
            merge_run_properties(&mut props, direct);
        }

        props
    }

    /// Extract all text from the document.
    ///
    /// Paragraphs are separated by newlines.
    pub fn text(&self) -> String {
        self.body.text()
    }
}

/// The document body containing paragraphs and other block-level elements.
#[derive(Debug, Clone, Default)]
pub struct Body {
    /// Paragraphs in the body.
    paragraphs: Vec<Paragraph>,
}

impl Body {
    /// Create an empty body.
    pub fn new() -> Self {
        Self::default()
    }

    /// Get paragraphs in the body.
    pub fn paragraphs(&self) -> &[Paragraph] {
        &self.paragraphs
    }

    /// Get a mutable reference to paragraphs.
    pub fn paragraphs_mut(&mut self) -> &mut Vec<Paragraph> {
        &mut self.paragraphs
    }

    /// Add a new paragraph to the body.
    pub fn add_paragraph(&mut self) -> &mut Paragraph {
        self.paragraphs.push(Paragraph::new());
        self.paragraphs.last_mut().unwrap()
    }

    /// Extract all text from the body.
    pub fn text(&self) -> String {
        self.paragraphs
            .iter()
            .map(|p| p.text())
            .collect::<Vec<_>>()
            .join("\n")
    }
}

/// A paragraph in the document.
///
/// Corresponds to the `<w:p>` element.
#[derive(Debug, Clone, Default)]
pub struct Paragraph {
    /// Runs in the paragraph.
    runs: Vec<Run>,
    /// Paragraph properties (style, alignment, etc.).
    properties: Option<ParagraphProperties>,
}

impl Paragraph {
    /// Create an empty paragraph.
    pub fn new() -> Self {
        Self::default()
    }

    /// Get runs in the paragraph.
    pub fn runs(&self) -> &[Run] {
        &self.runs
    }

    /// Get a mutable reference to runs.
    pub fn runs_mut(&mut self) -> &mut Vec<Run> {
        &mut self.runs
    }

    /// Add a new run to the paragraph.
    pub fn add_run(&mut self) -> &mut Run {
        self.runs.push(Run::new());
        self.runs.last_mut().unwrap()
    }

    /// Get paragraph properties.
    pub fn properties(&self) -> Option<&ParagraphProperties> {
        self.properties.as_ref()
    }

    /// Set paragraph properties.
    pub fn set_properties(&mut self, props: ParagraphProperties) {
        self.properties = Some(props);
    }

    /// Extract all text from the paragraph.
    pub fn text(&self) -> String {
        self.runs.iter().map(|r| r.text()).collect()
    }
}

/// Properties of a paragraph.
///
/// Corresponds to the `<w:pPr>` element.
#[derive(Debug, Clone, Default)]
pub struct ParagraphProperties {
    /// Style ID reference.
    pub style: Option<String>,
}

/// A text run in the document.
///
/// Corresponds to the `<w:r>` element. A run is a contiguous range of text
/// with the same formatting.
#[derive(Debug, Clone, Default)]
pub struct Run {
    /// Text content in the run.
    text: String,
    /// Run properties (bold, italic, etc.).
    properties: Option<RunProperties>,
}

impl Run {
    /// Create an empty run.
    pub fn new() -> Self {
        Self::default()
    }

    /// Get the text content.
    pub fn text(&self) -> &str {
        &self.text
    }

    /// Set the text content.
    pub fn set_text(&mut self, text: impl Into<String>) {
        self.text = text.into();
    }

    /// Get run properties.
    pub fn properties(&self) -> Option<&RunProperties> {
        self.properties.as_ref()
    }

    /// Set run properties.
    pub fn set_properties(&mut self, props: RunProperties) {
        self.properties = Some(props);
    }

    /// Check if the run is bold.
    pub fn is_bold(&self) -> bool {
        self.properties.as_ref().is_some_and(|p| p.bold)
    }

    /// Check if the run is italic.
    pub fn is_italic(&self) -> bool {
        self.properties.as_ref().is_some_and(|p| p.italic)
    }

    /// Check if the run is underlined.
    pub fn is_underline(&self) -> bool {
        self.properties.as_ref().is_some_and(|p| p.underline)
    }
}

/// Properties of a text run.
///
/// Corresponds to the `<w:rPr>` element.
#[derive(Debug, Clone, Default)]
pub struct RunProperties {
    /// Bold formatting.
    pub bold: bool,
    /// Italic formatting.
    pub italic: bool,
    /// Underline formatting.
    pub underline: bool,
    /// Strike-through formatting.
    pub strike: bool,
    /// Font size in half-points.
    pub size: Option<u32>,
    /// Font name.
    pub font: Option<String>,
    /// Style ID reference.
    pub style: Option<String>,
}

// XML element names (local names)
const EL_BODY: &[u8] = b"body";
const EL_P: &[u8] = b"p";
const EL_R: &[u8] = b"r";
const EL_T: &[u8] = b"t";
const EL_PPR: &[u8] = b"pPr";
const EL_RPR: &[u8] = b"rPr";
const EL_PSTYLE: &[u8] = b"pStyle";
const EL_RSTYLE: &[u8] = b"rStyle";
const EL_B: &[u8] = b"b";
const EL_I: &[u8] = b"i";
const EL_U: &[u8] = b"u";
const EL_STRIKE: &[u8] = b"strike";
const EL_SZ: &[u8] = b"sz";
const EL_RFONTS: &[u8] = b"rFonts";
const EL_BR: &[u8] = b"br";
const EL_TAB: &[u8] = b"tab";

/// Parse a document.xml file into a Body.
fn parse_document(xml: &[u8]) -> Result<Body> {
    let mut reader = Reader::from_reader(xml);
    reader.config_mut().trim_text(false); // Preserve whitespace in text

    let mut buf = Vec::new();
    let mut body = Body::new();
    let mut in_body = false;
    let mut current_para: Option<Paragraph> = None;
    let mut current_run: Option<Run> = None;
    let mut current_ppr: Option<ParagraphProperties> = None;
    let mut current_rpr: Option<RunProperties> = None;
    let mut in_ppr = false;
    let mut in_rpr = false;
    let mut in_text = false;

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) => {
                let name = e.name();
                let local = local_name(name.as_ref());
                match local {
                    name if name == EL_BODY => {
                        in_body = true;
                    }
                    name if name == EL_P && in_body => {
                        current_para = Some(Paragraph::new());
                    }
                    name if name == EL_R && current_para.is_some() => {
                        current_run = Some(Run::new());
                    }
                    name if name == EL_T && current_run.is_some() => {
                        in_text = true;
                    }
                    name if name == EL_PPR && current_para.is_some() => {
                        in_ppr = true;
                        current_ppr = Some(ParagraphProperties::default());
                    }
                    name if name == EL_RPR && current_run.is_some() => {
                        in_rpr = true;
                        current_rpr = Some(RunProperties::default());
                    }
                    _ => {}
                }
            }
            Ok(Event::Empty(e)) => {
                let name = e.name();
                let local = local_name(name.as_ref());
                match local {
                    name if name == EL_T && current_run.is_some() => {
                        // Empty text element, nothing to do
                    }
                    name if name == EL_BR && current_run.is_some() => {
                        // Line break
                        if let Some(run) = current_run.as_mut() {
                            run.text.push('\n');
                        }
                    }
                    name if name == EL_TAB && current_run.is_some() => {
                        // Tab character
                        if let Some(run) = current_run.as_mut() {
                            run.text.push('\t');
                        }
                    }
                    name if name == EL_PSTYLE && in_ppr => {
                        if let Some(ppr) = current_ppr.as_mut() {
                            for attr in e.attributes().filter_map(|a| a.ok()) {
                                if attr.key.as_ref() == b"w:val" || attr.key.as_ref() == b"val" {
                                    ppr.style =
                                        Some(String::from_utf8_lossy(&attr.value).into_owned());
                                }
                            }
                        }
                    }
                    name if name == EL_RSTYLE && in_rpr => {
                        if let Some(rpr) = current_rpr.as_mut() {
                            for attr in e.attributes().filter_map(|a| a.ok()) {
                                if attr.key.as_ref() == b"w:val" || attr.key.as_ref() == b"val" {
                                    rpr.style =
                                        Some(String::from_utf8_lossy(&attr.value).into_owned());
                                }
                            }
                        }
                    }
                    name if name == EL_B && in_rpr => {
                        if let Some(rpr) = current_rpr.as_mut() {
                            rpr.bold = parse_toggle_val(&e);
                        }
                    }
                    name if name == EL_I && in_rpr => {
                        if let Some(rpr) = current_rpr.as_mut() {
                            rpr.italic = parse_toggle_val(&e);
                        }
                    }
                    name if name == EL_U && in_rpr => {
                        if let Some(rpr) = current_rpr.as_mut() {
                            // Underline is present if <w:u> exists with val != "none"
                            rpr.underline = true;
                            for attr in e.attributes().filter_map(|a| a.ok()) {
                                if (attr.key.as_ref() == b"w:val" || attr.key.as_ref() == b"val")
                                    && attr.value.as_ref() == b"none"
                                {
                                    rpr.underline = false;
                                }
                            }
                        }
                    }
                    name if name == EL_STRIKE && in_rpr => {
                        if let Some(rpr) = current_rpr.as_mut() {
                            rpr.strike = parse_toggle_val(&e);
                        }
                    }
                    name if name == EL_SZ && in_rpr => {
                        if let Some(rpr) = current_rpr.as_mut() {
                            for attr in e.attributes().filter_map(|a| a.ok()) {
                                if (attr.key.as_ref() == b"w:val" || attr.key.as_ref() == b"val")
                                    && let Ok(s) = std::str::from_utf8(&attr.value)
                                {
                                    rpr.size = s.parse().ok();
                                }
                            }
                        }
                    }
                    name if name == EL_RFONTS && in_rpr => {
                        if let Some(rpr) = current_rpr.as_mut() {
                            for attr in e.attributes().filter_map(|a| a.ok()) {
                                if attr.key.as_ref() == b"w:ascii" || attr.key.as_ref() == b"ascii"
                                {
                                    rpr.font =
                                        Some(String::from_utf8_lossy(&attr.value).into_owned());
                                    break;
                                }
                            }
                        }
                    }
                    _ => {}
                }
            }
            Ok(Event::Text(e)) => {
                if in_text && let Some(run) = current_run.as_mut() {
                    let text = e.unescape().unwrap_or_default();
                    run.text.push_str(&text);
                }
            }
            Ok(Event::End(e)) => {
                let name = e.name();
                let local = local_name(name.as_ref());
                match local {
                    name if name == EL_BODY => {
                        in_body = false;
                    }
                    name if name == EL_P && current_para.is_some() => {
                        if let Some(mut para) = current_para.take() {
                            para.properties = current_ppr.take();
                            body.paragraphs.push(para);
                        }
                    }
                    name if name == EL_R && current_run.is_some() => {
                        if let Some(mut run) = current_run.take() {
                            run.properties = current_rpr.take();
                            if let Some(para) = current_para.as_mut() {
                                para.runs.push(run);
                            }
                        }
                    }
                    name if name == EL_T => {
                        in_text = false;
                    }
                    name if name == EL_PPR => {
                        in_ppr = false;
                    }
                    name if name == EL_RPR => {
                        in_rpr = false;
                    }
                    _ => {}
                }
            }
            Ok(Event::Eof) => break,
            Err(e) => return Err(Error::Xml(e)),
            _ => {}
        }
        buf.clear();
    }

    Ok(body)
}

/// Extract the local name from a potentially namespaced element name.
fn local_name(name: &[u8]) -> &[u8] {
    // Handle both "w:p" and "p" formats
    if let Some(pos) = name.iter().position(|&b| b == b':') {
        &name[pos + 1..]
    } else {
        name
    }
}

/// Parse a toggle property value (like <w:b/> or <w:b w:val="true"/>).
///
/// Toggle properties are true if:
/// - Element is present with no val attribute
/// - Element has val="true", "1", or "on"
fn parse_toggle_val(e: &quick_xml::events::BytesStart) -> bool {
    for attr in e.attributes().filter_map(|a| a.ok()) {
        if attr.key.as_ref() == b"w:val" || attr.key.as_ref() == b"val" {
            return matches!(
                attr.value.as_ref(),
                b"true" | b"1" | b"on" | b"True" | b"On"
            );
        }
    }
    // No val attribute means true
    true
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_parse_simple_document() {
        let xml = br#"<?xml version="1.0" encoding="UTF-8"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p>
      <w:r>
        <w:t>Hello, World!</w:t>
      </w:r>
    </w:p>
  </w:body>
</w:document>"#;

        let body = parse_document(xml).unwrap();

        assert_eq!(body.paragraphs().len(), 1);
        assert_eq!(body.paragraphs()[0].runs().len(), 1);
        assert_eq!(body.paragraphs()[0].runs()[0].text(), "Hello, World!");
        assert_eq!(body.text(), "Hello, World!");
    }

    #[test]
    fn test_parse_multiple_paragraphs() {
        let xml = br#"<?xml version="1.0" encoding="UTF-8"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p>
      <w:r><w:t>First paragraph.</w:t></w:r>
    </w:p>
    <w:p>
      <w:r><w:t>Second paragraph.</w:t></w:r>
    </w:p>
  </w:body>
</w:document>"#;

        let body = parse_document(xml).unwrap();

        assert_eq!(body.paragraphs().len(), 2);
        assert_eq!(body.text(), "First paragraph.\nSecond paragraph.");
    }

    #[test]
    fn test_parse_multiple_runs() {
        let xml = br#"<?xml version="1.0" encoding="UTF-8"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p>
      <w:r><w:t>Hello, </w:t></w:r>
      <w:r><w:t>World!</w:t></w:r>
    </w:p>
  </w:body>
</w:document>"#;

        let body = parse_document(xml).unwrap();

        assert_eq!(body.paragraphs().len(), 1);
        assert_eq!(body.paragraphs()[0].runs().len(), 2);
        assert_eq!(body.paragraphs()[0].text(), "Hello, World!");
    }

    #[test]
    fn test_parse_formatting() {
        let xml = br#"<?xml version="1.0" encoding="UTF-8"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p>
      <w:r>
        <w:rPr>
          <w:b/>
          <w:i/>
          <w:u w:val="single"/>
        </w:rPr>
        <w:t>Formatted text</w:t>
      </w:r>
    </w:p>
  </w:body>
</w:document>"#;

        let body = parse_document(xml).unwrap();
        let run = &body.paragraphs()[0].runs()[0];

        assert!(run.is_bold());
        assert!(run.is_italic());
        assert!(run.is_underline());
    }

    #[test]
    fn test_parse_toggle_values() {
        let xml = br#"<?xml version="1.0" encoding="UTF-8"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p>
      <w:r>
        <w:rPr>
          <w:b w:val="true"/>
          <w:i w:val="false"/>
        </w:rPr>
        <w:t>Test</w:t>
      </w:r>
    </w:p>
  </w:body>
</w:document>"#;

        let body = parse_document(xml).unwrap();
        let run = &body.paragraphs()[0].runs()[0];

        assert!(run.is_bold());
        assert!(!run.is_italic());
    }

    #[test]
    fn test_parse_paragraph_style() {
        let xml = br#"<?xml version="1.0" encoding="UTF-8"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p>
      <w:pPr>
        <w:pStyle w:val="Heading1"/>
      </w:pPr>
      <w:r><w:t>Title</w:t></w:r>
    </w:p>
  </w:body>
</w:document>"#;

        let body = parse_document(xml).unwrap();
        let para = &body.paragraphs()[0];

        assert_eq!(
            para.properties().and_then(|p| p.style.as_deref()),
            Some("Heading1")
        );
    }

    #[test]
    fn test_parse_line_breaks_and_tabs() {
        let xml = br#"<?xml version="1.0" encoding="UTF-8"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p>
      <w:r>
        <w:t>Line 1</w:t>
        <w:br/>
        <w:t>Line 2</w:t>
        <w:tab/>
        <w:t>After tab</w:t>
      </w:r>
    </w:p>
  </w:body>
</w:document>"#;

        let body = parse_document(xml).unwrap();
        let text = body.paragraphs()[0].runs()[0].text();

        assert_eq!(text, "Line 1\nLine 2\tAfter tab");
    }
}
