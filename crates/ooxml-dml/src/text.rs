//! DrawingML text types (a:p, a:r, a:t).
//!
//! These types represent text content in DrawingML, used for shape text
//! in PowerPoint and text boxes in Word/Excel.
//!
//! Reference: ECMA-376 Part 4, Section 21.1.2 (DrawingML - Text)

use quick_xml::Reader;
use quick_xml::events::Event;
use std::io::{BufRead, Cursor};

use crate::error::{Error, Result};

/// A DrawingML paragraph (`<a:p>`).
///
/// Contains text runs and paragraph-level formatting.
#[derive(Debug, Clone, Default)]
pub struct Paragraph {
    /// Text runs in the paragraph.
    runs: Vec<Run>,
    /// Paragraph properties (alignment, spacing, etc.)
    properties: Option<ParagraphProperties>,
}

impl Paragraph {
    /// Create a new empty paragraph.
    pub fn new() -> Self {
        Self::default()
    }

    /// Get the text runs.
    pub fn runs(&self) -> &[Run] {
        &self.runs
    }

    /// Get paragraph properties.
    pub fn properties(&self) -> Option<&ParagraphProperties> {
        self.properties.as_ref()
    }

    /// Extract all text from the paragraph.
    pub fn text(&self) -> String {
        self.runs.iter().map(|r| r.text.as_str()).collect()
    }

    /// Add a run to the paragraph.
    pub fn add_run(&mut self, run: Run) {
        self.runs.push(run);
    }
}

/// A DrawingML text run (`<a:r>`).
///
/// Contains text content and character-level formatting.
#[derive(Debug, Clone, Default)]
pub struct Run {
    /// Text content.
    text: String,
    /// Run properties (bold, italic, font, etc.)
    properties: Option<RunProperties>,
}

impl Run {
    /// Create a new run with text.
    pub fn new(text: impl Into<String>) -> Self {
        Self {
            text: text.into(),
            properties: None,
        }
    }

    /// Get the text content.
    pub fn text(&self) -> &str {
        &self.text
    }

    /// Get run properties.
    pub fn properties(&self) -> Option<&RunProperties> {
        self.properties.as_ref()
    }

    /// Set run properties.
    pub fn set_properties(&mut self, props: RunProperties) {
        self.properties = Some(props);
    }

    /// Check if bold.
    pub fn is_bold(&self) -> bool {
        self.properties.as_ref().is_some_and(|p| p.bold)
    }

    /// Check if italic.
    pub fn is_italic(&self) -> bool {
        self.properties.as_ref().is_some_and(|p| p.italic)
    }

    /// Check if the run has a hyperlink.
    pub fn has_hyperlink(&self) -> bool {
        self.properties
            .as_ref()
            .is_some_and(|p| p.hyperlink_rel_id.is_some())
    }

    /// Get the hyperlink relationship ID (for resolving via relationships).
    pub fn hyperlink_rel_id(&self) -> Option<&str> {
        self.properties
            .as_ref()
            .and_then(|p| p.hyperlink_rel_id.as_deref())
    }
}

/// Paragraph properties (`<a:pPr>`).
#[derive(Debug, Clone, Default)]
pub struct ParagraphProperties {
    /// Text alignment.
    pub alignment: Option<TextAlignment>,
    /// Indentation level (for bullets/numbering).
    pub level: Option<u32>,
}

/// Text alignment options.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum TextAlignment {
    Left,
    Center,
    Right,
    Justify,
    Distributed,
}

/// Run properties (`<a:rPr>`).
#[derive(Debug, Clone, Default)]
pub struct RunProperties {
    /// Bold text.
    pub bold: bool,
    /// Italic text.
    pub italic: bool,
    /// Underline.
    pub underline: bool,
    /// Strike-through.
    pub strikethrough: bool,
    /// Font size in hundredths of a point.
    pub font_size: Option<u32>,
    /// Font name.
    pub font_name: Option<String>,
    /// Text color (as hex RGB, e.g., "FF0000").
    pub color: Option<String>,
    /// Hyperlink reference (relationship ID for click action).
    pub hyperlink_rel_id: Option<String>,
}

// ============================================================================
// Parsing
// ============================================================================

/// Parse DrawingML text body content (`<a:txBody>` children).
///
/// Returns a list of paragraphs found in the text body.
pub fn parse_text_body(xml: &[u8]) -> Result<Vec<Paragraph>> {
    let mut reader = Reader::from_reader(Cursor::new(xml));
    parse_text_body_from_reader(&mut reader)
}

/// Parse DrawingML text body from an existing reader.
pub fn parse_text_body_from_reader<R: BufRead>(reader: &mut Reader<R>) -> Result<Vec<Paragraph>> {
    let mut buf = Vec::new();
    let mut paragraphs = Vec::new();
    let mut current_para: Option<Paragraph> = None;
    let mut current_run: Option<Run> = None;
    let mut current_props: Option<RunProperties> = None;
    let mut in_rpr = false;
    let mut in_t = false;
    let mut current_text = String::new();

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) => {
                let name = e.name();
                let name = name.as_ref();
                match name {
                    b"a:p" => {
                        current_para = Some(Paragraph::new());
                    }
                    b"a:r" => {
                        current_run = Some(Run::default());
                    }
                    b"a:rPr" => {
                        in_rpr = true;
                        current_props = Some(parse_run_properties(&e));
                    }
                    b"a:t" => {
                        in_t = true;
                        current_text.clear();
                    }
                    _ => {}
                }
            }
            Ok(Event::Text(e)) => {
                if in_t {
                    current_text.push_str(&e.decode().unwrap_or_default());
                }
            }
            Ok(Event::End(e)) => {
                let name = e.name();
                let name = name.as_ref();
                match name {
                    b"a:p" => {
                        if let Some(para) = current_para.take() {
                            paragraphs.push(para);
                        }
                    }
                    b"a:r" => {
                        if let Some(mut run) = current_run.take() {
                            if let Some(props) = current_props.take() {
                                run.properties = Some(props);
                            }
                            if let Some(para) = current_para.as_mut() {
                                para.runs.push(run);
                            }
                        }
                    }
                    b"a:rPr" => {
                        in_rpr = false;
                    }
                    b"a:t" => {
                        in_t = false;
                        if let Some(run) = current_run.as_mut() {
                            run.text = std::mem::take(&mut current_text);
                        }
                    }
                    b"a:txBody" => {
                        // End of text body
                        break;
                    }
                    _ => {}
                }
            }
            Ok(Event::Empty(e)) => {
                let name = e.name();
                let name = name.as_ref();
                if name == b"a:rPr" {
                    current_props = Some(parse_run_properties(&e));
                } else if in_rpr && name == b"a:hlinkClick" {
                    // Parse hyperlink click action
                    if let Some(props) = current_props.as_mut() {
                        for attr in e.attributes().filter_map(|a| a.ok()) {
                            let key = attr.key.as_ref();
                            // Check for r:id attribute (relationship ID)
                            if key == b"r:id" || key == b"id" {
                                props.hyperlink_rel_id =
                                    Some(String::from_utf8_lossy(&attr.value).into_owned());
                            }
                        }
                    }
                }
            }
            Ok(Event::Eof) => break,
            Err(e) => return Err(Error::Xml(e)),
            _ => {}
        }
        buf.clear();
    }

    Ok(paragraphs)
}

/// Parse run properties from a BytesStart element.
fn parse_run_properties(e: &quick_xml::events::BytesStart) -> RunProperties {
    let mut props = RunProperties::default();

    for attr in e.attributes().filter_map(|a| a.ok()) {
        let key = attr.key.as_ref();
        let value = String::from_utf8_lossy(&attr.value);
        match key {
            b"b" => props.bold = value == "1" || value == "true",
            b"i" => props.italic = value == "1" || value == "true",
            b"u" => props.underline = value != "none",
            b"strike" => props.strikethrough = value != "noStrike",
            b"sz" => props.font_size = value.parse().ok(),
            _ => {}
        }
    }

    props
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_parse_simple_text() {
        let xml = r#"<a:txBody>
            <a:p>
                <a:r>
                    <a:t>Hello World</a:t>
                </a:r>
            </a:p>
        </a:txBody>"#;

        let paragraphs = parse_text_body(xml.as_bytes()).unwrap();
        assert_eq!(paragraphs.len(), 1);
        assert_eq!(paragraphs[0].text(), "Hello World");
    }

    #[test]
    fn test_parse_formatted_text() {
        let xml = r#"<a:txBody>
            <a:p>
                <a:r>
                    <a:rPr b="1" i="1"/>
                    <a:t>Bold Italic</a:t>
                </a:r>
            </a:p>
        </a:txBody>"#;

        let paragraphs = parse_text_body(xml.as_bytes()).unwrap();
        assert_eq!(paragraphs.len(), 1);
        let run = &paragraphs[0].runs()[0];
        assert!(run.is_bold());
        assert!(run.is_italic());
    }

    #[test]
    fn test_parse_multiple_runs() {
        let xml = r#"<a:txBody>
            <a:p>
                <a:r><a:t>Hello </a:t></a:r>
                <a:r><a:t>World</a:t></a:r>
            </a:p>
        </a:txBody>"#;

        let paragraphs = parse_text_body(xml.as_bytes()).unwrap();
        assert_eq!(paragraphs.len(), 1);
        assert_eq!(paragraphs[0].runs().len(), 2);
        assert_eq!(paragraphs[0].text(), "Hello World");
    }

    #[test]
    fn test_parse_hyperlink() {
        let xml = r#"<a:txBody>
            <a:p>
                <a:r>
                    <a:rPr>
                        <a:hlinkClick r:id="rId1"/>
                    </a:rPr>
                    <a:t>Click here</a:t>
                </a:r>
            </a:p>
        </a:txBody>"#;

        let paragraphs = parse_text_body(xml.as_bytes()).unwrap();
        assert_eq!(paragraphs.len(), 1);
        let run = &paragraphs[0].runs()[0];
        assert!(run.has_hyperlink());
        assert_eq!(run.hyperlink_rel_id(), Some("rId1"));
        assert_eq!(run.text(), "Click here");
    }
}
