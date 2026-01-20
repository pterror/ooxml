//! Style definitions and resolution.
//!
//! Word documents use styles to define formatting that can be applied to
//! paragraphs and runs. Styles can inherit from other styles via `basedOn`.
//!
//! # Style Types
//!
//! - **Paragraph styles** - Applied to entire paragraphs (`<w:pStyle>`)
//! - **Character styles** - Applied to runs (`<w:rStyle>`)
//! - **Table styles** - Applied to tables
//! - **Numbering styles** - Applied to numbered/bulleted lists
//!
//! # Example
//!
//! ```ignore
//! let styles = Styles::parse(styles_xml)?;
//! if let Some(style) = styles.get("Heading1") {
//!     println!("Heading1 is bold: {}", style.run_properties().bold);
//! }
//! ```

use crate::document::{ParagraphProperties, RunProperties, UnderlineStyle};
use crate::error::{Error, Result};
use quick_xml::Reader;
use quick_xml::events::Event;
use std::collections::HashMap;
use std::io::Read;

/// A collection of style definitions from styles.xml.
#[derive(Debug, Clone, Default)]
pub struct Styles {
    /// Style definitions indexed by styleId.
    styles: HashMap<String, Style>,
    /// Default paragraph properties (from docDefaults).
    default_paragraph: ParagraphProperties,
    /// Default run properties (from docDefaults).
    default_run: RunProperties,
}

impl Styles {
    /// Create an empty styles collection.
    pub fn new() -> Self {
        Self::default()
    }

    /// Parse styles from XML (styles.xml content).
    pub fn parse<R: Read>(reader: R) -> Result<Self> {
        parse_styles(reader)
    }

    /// Get a style by ID.
    pub fn get(&self, style_id: &str) -> Option<&Style> {
        self.styles.get(style_id)
    }

    /// Get default paragraph properties.
    pub fn default_paragraph(&self) -> &ParagraphProperties {
        &self.default_paragraph
    }

    /// Get default run properties.
    pub fn default_run(&self) -> &RunProperties {
        &self.default_run
    }

    /// Resolve the effective run properties for a style, including inheritance.
    ///
    /// This walks up the `basedOn` chain and merges properties.
    pub fn resolve_run_properties(&self, style_id: &str) -> RunProperties {
        let mut props = self.default_run.clone();
        self.apply_run_properties_chain(style_id, &mut props, 0);
        props
    }

    /// Resolve the effective paragraph properties for a style, including inheritance.
    pub fn resolve_paragraph_properties(&self, style_id: &str) -> ParagraphProperties {
        let mut props = self.default_paragraph.clone();
        self.apply_paragraph_properties_chain(style_id, &mut props, 0);
        props
    }

    /// Apply run properties from a style chain (recursive with depth limit).
    fn apply_run_properties_chain(&self, style_id: &str, props: &mut RunProperties, depth: usize) {
        // Prevent infinite loops from circular references
        if depth > 20 {
            return;
        }

        if let Some(style) = self.styles.get(style_id) {
            // First apply base style properties
            if let Some(ref based_on) = style.based_on {
                self.apply_run_properties_chain(based_on, props, depth + 1);
            }

            // Then apply this style's properties (overrides base)
            if let Some(ref rpr) = style.run_properties {
                merge_run_properties(props, rpr);
            }
        }
    }

    /// Apply paragraph properties from a style chain (recursive with depth limit).
    fn apply_paragraph_properties_chain(
        &self,
        style_id: &str,
        props: &mut ParagraphProperties,
        depth: usize,
    ) {
        if depth > 20 {
            return;
        }

        if let Some(style) = self.styles.get(style_id) {
            if let Some(ref based_on) = style.based_on {
                self.apply_paragraph_properties_chain(based_on, props, depth + 1);
            }

            if let Some(ref ppr) = style.paragraph_properties {
                merge_paragraph_properties(props, ppr);
            }
        }
    }

    /// Iterate over all styles.
    pub fn iter(&self) -> impl Iterator<Item = (&str, &Style)> {
        self.styles.iter().map(|(k, v)| (k.as_str(), v))
    }

    /// Get the number of styles.
    pub fn len(&self) -> usize {
        self.styles.len()
    }

    /// Check if there are no styles.
    pub fn is_empty(&self) -> bool {
        self.styles.is_empty()
    }
}

/// A single style definition.
#[derive(Debug, Clone, Default)]
pub struct Style {
    /// Style ID (unique identifier).
    pub id: String,
    /// Style name (display name).
    pub name: Option<String>,
    /// Style type (paragraph, character, table, numbering).
    pub style_type: StyleType,
    /// ID of the style this one is based on.
    pub based_on: Option<String>,
    /// Whether this is the default style for its type.
    pub is_default: bool,
    /// Paragraph properties defined by this style.
    pub paragraph_properties: Option<ParagraphProperties>,
    /// Run properties defined by this style.
    pub run_properties: Option<RunProperties>,
}

/// Type of style.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub enum StyleType {
    /// Paragraph style (applied to entire paragraphs).
    #[default]
    Paragraph,
    /// Character style (applied to runs/text).
    Character,
    /// Table style.
    Table,
    /// Numbering style.
    Numbering,
}

impl StyleType {
    fn from_str(s: &str) -> Self {
        match s {
            "paragraph" => StyleType::Paragraph,
            "character" => StyleType::Character,
            "table" => StyleType::Table,
            "numbering" => StyleType::Numbering,
            _ => StyleType::Paragraph,
        }
    }
}

/// Merge source run properties into target (source overrides where set).
pub fn merge_run_properties(target: &mut RunProperties, source: &RunProperties) {
    if source.bold {
        target.bold = true;
    }
    if source.italic {
        target.italic = true;
    }
    if source.underline.is_some() {
        target.underline = source.underline;
    }
    if source.strike {
        target.strike = true;
    }
    if source.double_strike {
        target.double_strike = true;
    }
    if source.size.is_some() {
        target.size = source.size;
    }
    if source.font.is_some() {
        target.font = source.font.clone();
    }
    if source.style.is_some() {
        target.style = source.style.clone();
    }
    if source.color.is_some() {
        target.color = source.color.clone();
    }
    if source.highlight.is_some() {
        target.highlight = source.highlight;
    }
    if source.vertical_align.is_some() {
        target.vertical_align = source.vertical_align;
    }
    if source.all_caps {
        target.all_caps = true;
    }
    if source.small_caps {
        target.small_caps = true;
    }
}

/// Merge source paragraph properties into target.
fn merge_paragraph_properties(target: &mut ParagraphProperties, source: &ParagraphProperties) {
    if source.style.is_some() {
        target.style = source.style.clone();
    }
}

// XML element names
const EL_DOC_DEFAULTS: &[u8] = b"docDefaults";
const EL_RPR_DEFAULT: &[u8] = b"rPrDefault";
const EL_PPR_DEFAULT: &[u8] = b"pPrDefault";
const EL_STYLE: &[u8] = b"style";
const EL_NAME: &[u8] = b"name";
const EL_BASED_ON: &[u8] = b"basedOn";
const EL_RPR: &[u8] = b"rPr";
const EL_PPR: &[u8] = b"pPr";
const EL_B: &[u8] = b"b";
const EL_I: &[u8] = b"i";
const EL_U: &[u8] = b"u";
const EL_STRIKE: &[u8] = b"strike";
const EL_SZ: &[u8] = b"sz";
const EL_RFONTS: &[u8] = b"rFonts";

/// Parse styles.xml into a Styles collection.
fn parse_styles<R: Read>(reader: R) -> Result<Styles> {
    let mut xml = Reader::from_reader(std::io::BufReader::new(reader));
    xml.config_mut().trim_text(true);

    let mut styles = Styles::new();
    let mut buf = Vec::new();

    // Parsing state
    let mut in_doc_defaults = false;
    let mut in_rpr_default = false;
    let mut _in_ppr_default = false; // TODO: parse pPr defaults
    let mut current_style: Option<Style> = None;
    let mut in_style_rpr = false;
    let mut _in_style_ppr = false; // TODO: parse style pPr
    let mut current_rpr: Option<RunProperties> = None;

    loop {
        match xml.read_event_into(&mut buf) {
            Ok(Event::Start(e)) => {
                let name = e.name();
                let local = local_name(name.as_ref());

                match local {
                    name if name == EL_DOC_DEFAULTS => {
                        in_doc_defaults = true;
                    }
                    name if name == EL_RPR_DEFAULT && in_doc_defaults => {
                        in_rpr_default = true;
                    }
                    name if name == EL_PPR_DEFAULT && in_doc_defaults => {
                        _in_ppr_default = true;
                    }
                    name if name == EL_RPR => {
                        if in_rpr_default {
                            current_rpr = Some(RunProperties::default());
                        } else if current_style.is_some() {
                            in_style_rpr = true;
                            current_rpr = Some(RunProperties::default());
                        }
                    }
                    name if name == EL_PPR => {
                        if current_style.is_some() {
                            _in_style_ppr = true;
                        }
                    }
                    name if name == EL_STYLE => {
                        let mut style = Style::default();

                        for attr in e.attributes().filter_map(|a| a.ok()) {
                            let key = local_name(attr.key.as_ref());
                            match key {
                                b"type" => {
                                    let val = String::from_utf8_lossy(&attr.value);
                                    style.style_type = StyleType::from_str(&val);
                                }
                                b"styleId" => {
                                    style.id = String::from_utf8_lossy(&attr.value).into_owned();
                                }
                                b"default" => {
                                    let val = String::from_utf8_lossy(&attr.value);
                                    style.is_default = val == "1" || val == "true" || val == "on";
                                }
                                _ => {}
                            }
                        }

                        current_style = Some(style);
                    }
                    _ => {}
                }
            }
            Ok(Event::Empty(e)) => {
                let name = e.name();
                let local = local_name(name.as_ref());

                // Handle formatting elements in rPr
                if (in_rpr_default || in_style_rpr)
                    && let Some(ref mut rpr) = current_rpr
                {
                    match local {
                        name if name == EL_B => {
                            rpr.bold = parse_toggle_val(&e);
                        }
                        name if name == EL_I => {
                            rpr.italic = parse_toggle_val(&e);
                        }
                        name if name == EL_U => {
                            let mut style = UnderlineStyle::Single;
                            for attr in e.attributes().filter_map(|a| a.ok()) {
                                let key = local_name(attr.key.as_ref());
                                if key == b"val" {
                                    if let Ok(s) = std::str::from_utf8(&attr.value) {
                                        if let Some(parsed) = UnderlineStyle::parse(s) {
                                            style = parsed;
                                        } else {
                                            rpr.underline = None;
                                            break;
                                        }
                                    }
                                    rpr.underline = Some(style);
                                    break;
                                }
                            }
                            if rpr.underline.is_none() {
                                rpr.underline = Some(UnderlineStyle::Single);
                            }
                        }
                        name if name == EL_STRIKE => {
                            rpr.strike = parse_toggle_val(&e);
                        }
                        name if name == EL_SZ => {
                            for attr in e.attributes().filter_map(|a| a.ok()) {
                                let key = local_name(attr.key.as_ref());
                                if key == b"val"
                                    && let Ok(s) = std::str::from_utf8(&attr.value)
                                {
                                    rpr.size = s.parse().ok();
                                }
                            }
                        }
                        name if name == EL_RFONTS => {
                            for attr in e.attributes().filter_map(|a| a.ok()) {
                                let key = local_name(attr.key.as_ref());
                                if key == b"ascii" {
                                    rpr.font =
                                        Some(String::from_utf8_lossy(&attr.value).into_owned());
                                    break;
                                }
                            }
                        }
                        _ => {}
                    }
                }

                // Handle style child elements
                if let Some(ref mut style) = current_style {
                    match local {
                        name if name == EL_NAME => {
                            for attr in e.attributes().filter_map(|a| a.ok()) {
                                let key = local_name(attr.key.as_ref());
                                if key == b"val" {
                                    style.name =
                                        Some(String::from_utf8_lossy(&attr.value).into_owned());
                                }
                            }
                        }
                        name if name == EL_BASED_ON => {
                            for attr in e.attributes().filter_map(|a| a.ok()) {
                                let key = local_name(attr.key.as_ref());
                                if key == b"val" {
                                    style.based_on =
                                        Some(String::from_utf8_lossy(&attr.value).into_owned());
                                }
                            }
                        }
                        _ => {}
                    }
                }
            }
            Ok(Event::End(e)) => {
                let name = e.name();
                let local = local_name(name.as_ref());

                match local {
                    name if name == EL_DOC_DEFAULTS => {
                        in_doc_defaults = false;
                    }
                    name if name == EL_RPR_DEFAULT => {
                        if in_rpr_default {
                            if let Some(rpr) = current_rpr.take() {
                                styles.default_run = rpr;
                            }
                            in_rpr_default = false;
                        }
                    }
                    name if name == EL_PPR_DEFAULT => {
                        _in_ppr_default = false;
                    }
                    name if name == EL_RPR => {
                        if in_style_rpr {
                            if let Some(ref mut style) = current_style {
                                style.run_properties = current_rpr.take();
                            }
                            in_style_rpr = false;
                        } else if in_rpr_default && let Some(rpr) = current_rpr.take() {
                            styles.default_run = rpr;
                        }
                    }
                    name if name == EL_PPR => {
                        _in_style_ppr = false;
                    }
                    name if name == EL_STYLE => {
                        if let Some(style) = current_style.take()
                            && !style.id.is_empty()
                        {
                            styles.styles.insert(style.id.clone(), style);
                        }
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

    Ok(styles)
}

/// Extract local name from potentially namespaced element name.
fn local_name(name: &[u8]) -> &[u8] {
    if let Some(pos) = name.iter().position(|&b| b == b':') {
        &name[pos + 1..]
    } else {
        name
    }
}

/// Parse toggle property value.
fn parse_toggle_val(e: &quick_xml::events::BytesStart) -> bool {
    for attr in e.attributes().filter_map(|a| a.ok()) {
        let key = local_name(attr.key.as_ref());
        if key == b"val" {
            return matches!(
                attr.value.as_ref(),
                b"true" | b"1" | b"on" | b"True" | b"On"
            );
        }
    }
    true
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_parse_basic_styles() {
        let xml = br#"<?xml version="1.0" encoding="UTF-8"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:style w:type="paragraph" w:styleId="Normal" w:default="1">
    <w:name w:val="Normal"/>
  </w:style>
  <w:style w:type="paragraph" w:styleId="Heading1">
    <w:name w:val="Heading 1"/>
    <w:basedOn w:val="Normal"/>
    <w:rPr>
      <w:b/>
      <w:sz w:val="32"/>
    </w:rPr>
  </w:style>
</w:styles>"#;

        let styles = Styles::parse(&xml[..]).unwrap();

        assert_eq!(styles.len(), 2);

        let normal = styles.get("Normal").unwrap();
        assert_eq!(normal.name.as_deref(), Some("Normal"));
        assert!(normal.is_default);
        assert_eq!(normal.style_type, StyleType::Paragraph);

        let heading1 = styles.get("Heading1").unwrap();
        assert_eq!(heading1.name.as_deref(), Some("Heading 1"));
        assert_eq!(heading1.based_on.as_deref(), Some("Normal"));
        assert!(heading1.run_properties.as_ref().unwrap().bold);
        assert_eq!(heading1.run_properties.as_ref().unwrap().size, Some(32));
    }

    #[test]
    fn test_resolve_inherited_properties() {
        let xml = br#"<?xml version="1.0" encoding="UTF-8"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:style w:type="paragraph" w:styleId="Normal">
    <w:rPr>
      <w:rFonts w:ascii="Times New Roman"/>
      <w:sz w:val="24"/>
    </w:rPr>
  </w:style>
  <w:style w:type="paragraph" w:styleId="Heading1">
    <w:basedOn w:val="Normal"/>
    <w:rPr>
      <w:b/>
      <w:sz w:val="32"/>
    </w:rPr>
  </w:style>
</w:styles>"#;

        let styles = Styles::parse(&xml[..]).unwrap();

        // Heading1 should inherit font from Normal but override size
        let props = styles.resolve_run_properties("Heading1");
        assert!(props.bold);
        assert_eq!(props.size, Some(32)); // Overridden
        assert_eq!(props.font.as_deref(), Some("Times New Roman")); // Inherited
    }

    #[test]
    fn test_doc_defaults() {
        let xml = br#"<?xml version="1.0" encoding="UTF-8"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:docDefaults>
    <w:rPrDefault>
      <w:rPr>
        <w:rFonts w:ascii="Calibri"/>
        <w:sz w:val="22"/>
      </w:rPr>
    </w:rPrDefault>
  </w:docDefaults>
</w:styles>"#;

        let styles = Styles::parse(&xml[..]).unwrap();

        assert_eq!(styles.default_run().font.as_deref(), Some("Calibri"));
        assert_eq!(styles.default_run().size, Some(22));
    }

    #[test]
    fn test_character_style() {
        let xml = br#"<?xml version="1.0" encoding="UTF-8"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:style w:type="character" w:styleId="Strong">
    <w:name w:val="Strong"/>
    <w:rPr>
      <w:b/>
    </w:rPr>
  </w:style>
</w:styles>"#;

        let styles = Styles::parse(&xml[..]).unwrap();

        let strong = styles.get("Strong").unwrap();
        assert_eq!(strong.style_type, StyleType::Character);
        assert!(strong.run_properties.as_ref().unwrap().bold);
    }

    #[test]
    fn test_deep_inheritance() {
        let xml = br#"<?xml version="1.0" encoding="UTF-8"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:style w:type="paragraph" w:styleId="Base">
    <w:rPr>
      <w:rFonts w:ascii="Arial"/>
    </w:rPr>
  </w:style>
  <w:style w:type="paragraph" w:styleId="Level1">
    <w:basedOn w:val="Base"/>
    <w:rPr>
      <w:b/>
    </w:rPr>
  </w:style>
  <w:style w:type="paragraph" w:styleId="Level2">
    <w:basedOn w:val="Level1"/>
    <w:rPr>
      <w:i/>
    </w:rPr>
  </w:style>
</w:styles>"#;

        let styles = Styles::parse(&xml[..]).unwrap();

        // Level2 should have: Arial font (from Base), bold (from Level1), italic (own)
        let props = styles.resolve_run_properties("Level2");
        assert_eq!(props.font.as_deref(), Some("Arial"));
        assert!(props.bold);
        assert!(props.italic);
    }
}
