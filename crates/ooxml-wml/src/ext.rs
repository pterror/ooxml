//! Extension traits for generated WML types.
//!
//! This module provides convenience methods for the generated types via extension traits,
//! following the same pattern as SML's `ext.rs`. See ADR-003 for architectural rationale.
//!
//! # Design
//!
//! Extension traits are split into two categories:
//!
//! - **Pure traits** (`RunPropertiesExt`, `RunExt`, `ParagraphExt`, etc.): Methods that
//!   don't need external context
//! - **Resolve traits** (`RunResolveExt`): Methods that need `StyleContext` for
//!   style chain walking
//!
//! # Example
//!
//! ```ignore
//! use ooxml_wml::ext::{DocumentExt, BodyExt, ParagraphExt, RunExt};
//!
//! let doc: &types::Document = /* ... */;
//! if let Some(body) = doc.body() {
//!     for para in body.paragraphs() {
//!         println!("{}", para.text());
//!     }
//! }
//! ```

use crate::parsers::{FromXml, ParseError};
use crate::types;
use quick_xml::Reader;
use quick_xml::events::Event;
use std::io::Cursor;

// =============================================================================
// Helpers (private)
// =============================================================================

/// Check if a `CTOnOff` field represents "on" (ECMA-376 §17.17.4).
///
/// An omitted `val` attribute means "true" (the element's presence is the toggle).
/// Explicit values: "1", "true", "on" → true; "0", "false", "off" → false.
#[cfg_attr(
    not(any(feature = "wml-styling", feature = "wml-layout")),
    allow(dead_code)
)]
fn is_on(field: &Option<Box<types::CTOnOff>>) -> bool {
    match field {
        None => false,
        Some(ct) => match &ct.value {
            None => true, // element present with no val → on
            Some(v) => matches!(v.as_str(), "1" | "true" | "on"),
        },
    }
}

/// Tri-state check for style resolution: `None` = not specified, `Some(true/false)` = explicit.
#[cfg_attr(not(feature = "wml-styling"), allow(dead_code))]
fn check_toggle(field: &Option<Box<types::CTOnOff>>) -> Option<bool> {
    field.as_ref().map(|ct| match &ct.value {
        None => true,
        Some(v) => matches!(v.as_str(), "1" | "true" | "on"),
    })
}

/// Parse a half-point measurement string (e.g., "24" → 24 half-points = 12pt).
#[cfg_attr(not(feature = "wml-styling"), allow(dead_code))]
fn parse_half_points(s: &str) -> Option<u32> {
    s.parse::<u32>().ok()
}

// =============================================================================
// DocumentExt
// =============================================================================

/// Extension methods for `Document`.
pub trait DocumentExt {
    /// Get the document body (if present).
    fn body(&self) -> Option<&types::Body>;
}

impl DocumentExt for types::Document {
    fn body(&self) -> Option<&types::Body> {
        self.body.as_deref()
    }
}

// =============================================================================
// BodyExt
// =============================================================================

/// Extension methods for `Body`.
pub trait BodyExt {
    /// Get all paragraphs in the body.
    fn paragraphs(&self) -> Vec<&types::Paragraph>;

    /// Get all tables in the body.
    fn tables(&self) -> Vec<&types::Table>;

    /// Extract all text content from the body.
    fn text(&self) -> String;

    /// Get the document-level section properties (layout info).
    #[cfg(feature = "wml-layout")]
    fn section_properties(&self) -> Option<&types::SectionProperties>;
}

impl BodyExt for types::Body {
    fn paragraphs(&self) -> Vec<&types::Paragraph> {
        self.block_content
            .iter()
            .filter_map(|elt| match elt.as_ref() {
                types::BlockContent::P(p) => Some(p.as_ref()),
                _ => None,
            })
            .collect()
    }

    fn tables(&self) -> Vec<&types::Table> {
        self.block_content
            .iter()
            .filter_map(|elt| match elt.as_ref() {
                types::BlockContent::Tbl(t) => Some(t.as_ref()),
                _ => None,
            })
            .collect()
    }

    fn text(&self) -> String {
        let texts: Vec<String> = self
            .block_content
            .iter()
            .filter_map(|elt| match elt.as_ref() {
                types::BlockContent::P(p) => Some(p.text()),
                types::BlockContent::Tbl(t) => Some(t.text()),
                _ => None,
            })
            .collect();
        texts.join("\n")
    }

    #[cfg(feature = "wml-layout")]
    fn section_properties(&self) -> Option<&types::SectionProperties> {
        self.sect_pr.as_deref()
    }
}

// =============================================================================
// ParagraphExt
// =============================================================================

/// Extension methods for `Paragraph`.
pub trait ParagraphExt {
    /// Get all runs in this paragraph (including runs inside hyperlinks and simple fields).
    fn runs(&self) -> Vec<&types::Run>;

    /// Extract all text from this paragraph.
    fn text(&self) -> String;

    /// Get hyperlinks in this paragraph.
    fn hyperlinks(&self) -> Vec<&types::Hyperlink>;

    /// Get paragraph properties.
    #[cfg(feature = "wml-styling")]
    fn properties(&self) -> Option<&types::ParagraphProperties>;
}

impl ParagraphExt for types::Paragraph {
    fn runs(&self) -> Vec<&types::Run> {
        collect_runs_from_paragraph_content(&self.paragraph_content)
    }

    fn text(&self) -> String {
        let mut out = String::new();
        for content in &self.paragraph_content {
            collect_text_from_paragraph_content(content, &mut out);
        }
        out
    }

    fn hyperlinks(&self) -> Vec<&types::Hyperlink> {
        self.paragraph_content
            .iter()
            .filter_map(|c| match c.as_ref() {
                types::ParagraphContent::Hyperlink(h) => Some(h.as_ref()),
                _ => None,
            })
            .collect()
    }

    #[cfg(feature = "wml-styling")]
    fn properties(&self) -> Option<&types::ParagraphProperties> {
        self.p_pr.as_deref()
    }
}

/// Collect runs from paragraph content, including nested runs in hyperlinks and simple fields.
fn collect_runs_from_paragraph_content(
    content: &[Box<types::ParagraphContent>],
) -> Vec<&types::Run> {
    let mut runs = Vec::new();
    for item in content {
        match item.as_ref() {
            types::ParagraphContent::R(r) => runs.push(r.as_ref()),
            types::ParagraphContent::Hyperlink(h) => {
                runs.extend(collect_runs_from_paragraph_content(&h.paragraph_content));
            }
            types::ParagraphContent::FldSimple(f) => {
                runs.extend(collect_runs_from_paragraph_content(&f.paragraph_content));
            }
            _ => {}
        }
    }
    runs
}

/// Collect text from a single paragraph content item.
fn collect_text_from_paragraph_content(content: &types::ParagraphContent, out: &mut String) {
    match content {
        types::ParagraphContent::R(r) => out.push_str(&r.text()),
        types::ParagraphContent::Hyperlink(h) => {
            for item in &h.paragraph_content {
                collect_text_from_paragraph_content(item, out);
            }
        }
        types::ParagraphContent::FldSimple(f) => {
            for item in &f.paragraph_content {
                collect_text_from_paragraph_content(item, out);
            }
        }
        _ => {}
    }
}

// =============================================================================
// RunExt
// =============================================================================

/// Extension methods for `Run`.
pub trait RunExt {
    /// Extract text from this run.
    ///
    /// Collects `T` (text), `Tab` (→ `\t`), `Cr`/`Br`(non-page) (→ `\n`).
    fn text(&self) -> String;

    /// Get run properties.
    #[cfg(feature = "wml-styling")]
    fn properties(&self) -> Option<&types::RunProperties>;

    /// Check if this run contains a page break.
    fn has_page_break(&self) -> bool;

    /// Get all drawings in this run.
    #[cfg(feature = "wml-drawings")]
    fn drawings(&self) -> Vec<&types::CTDrawing>;

    /// Convenience: check if bold (delegates to properties).
    #[cfg(feature = "wml-styling")]
    fn is_bold(&self) -> bool;

    /// Convenience: check if italic (delegates to properties).
    #[cfg(feature = "wml-styling")]
    fn is_italic(&self) -> bool;

    /// Convenience: check if underlined (delegates to properties).
    #[cfg(feature = "wml-styling")]
    fn is_underline(&self) -> bool;

    /// Convenience: check if strikethrough (delegates to properties).
    #[cfg(feature = "wml-styling")]
    fn is_strikethrough(&self) -> bool;

    /// Check if this run contains any drawing elements (images).
    #[cfg(feature = "wml-drawings")]
    fn has_images(&self) -> bool;

    /// Get the footnote reference in this run, if any.
    fn footnote_ref(&self) -> Option<&types::FootnoteEndnoteRef>;

    /// Get the endnote reference in this run, if any.
    fn endnote_ref(&self) -> Option<&types::FootnoteEndnoteRef>;
}

impl RunExt for types::Run {
    fn text(&self) -> String {
        let mut out = String::new();
        for item in &self.run_content {
            match item.as_ref() {
                types::RunContent::T(t) => {
                    if let Some(ref text) = t.text {
                        out.push_str(text);
                    }
                }
                types::RunContent::Tab(_) => out.push('\t'),
                types::RunContent::Cr(_) => out.push('\n'),
                types::RunContent::Br(br) => {
                    // Page/column breaks aren't text; only text-wrapping breaks produce newlines
                    if !matches!(
                        br.r#type,
                        Some(types::STBrType::Page) | Some(types::STBrType::Column)
                    ) {
                        out.push('\n');
                    }
                }
                _ => {}
            }
        }
        out
    }

    #[cfg(feature = "wml-styling")]
    fn properties(&self) -> Option<&types::RunProperties> {
        self.r_pr.as_deref()
    }

    fn has_page_break(&self) -> bool {
        self.run_content.iter().any(|item| {
            matches!(
                item.as_ref(),
                types::RunContent::Br(br) if br.r#type == Some(types::STBrType::Page)
            )
        })
    }

    #[cfg(feature = "wml-drawings")]
    fn drawings(&self) -> Vec<&types::CTDrawing> {
        self.run_content
            .iter()
            .filter_map(|item| match item.as_ref() {
                types::RunContent::Drawing(d) => Some(d.as_ref()),
                _ => None,
            })
            .collect()
    }

    #[cfg(feature = "wml-styling")]
    fn is_bold(&self) -> bool {
        self.properties().is_some_and(|p| p.is_bold())
    }

    #[cfg(feature = "wml-styling")]
    fn is_italic(&self) -> bool {
        self.properties().is_some_and(|p| p.is_italic())
    }

    #[cfg(feature = "wml-styling")]
    fn is_underline(&self) -> bool {
        self.properties().is_some_and(|p| p.is_underline())
    }

    #[cfg(feature = "wml-styling")]
    fn is_strikethrough(&self) -> bool {
        self.properties().is_some_and(|p| p.is_strikethrough())
    }

    #[cfg(feature = "wml-drawings")]
    fn has_images(&self) -> bool {
        self.run_content
            .iter()
            .any(|item| matches!(item.as_ref(), types::RunContent::Drawing(_)))
    }

    fn footnote_ref(&self) -> Option<&types::FootnoteEndnoteRef> {
        self.run_content
            .iter()
            .find_map(|item| match item.as_ref() {
                types::RunContent::FootnoteReference(r) => Some(r.as_ref()),
                _ => None,
            })
    }

    fn endnote_ref(&self) -> Option<&types::FootnoteEndnoteRef> {
        self.run_content
            .iter()
            .find_map(|item| match item.as_ref() {
                types::RunContent::EndnoteReference(r) => Some(r.as_ref()),
                _ => None,
            })
    }
}

// =============================================================================
// DrawingExt
// =============================================================================

/// Extension methods for `CTDrawing` — extract image relationship IDs from raw XML.
///
/// Since `CTDrawing` captures its children as raw XML, this trait walks the tree
/// to find `<a:blip r:embed="..."/>` inside `<wp:inline>` and `<wp:anchor>` elements.
#[cfg(all(feature = "wml-drawings", feature = "extra-children"))]
pub trait DrawingExt {
    /// Get relationship IDs for inline images (`<wp:inline>` → `<a:blip r:embed="rId"/>`).
    fn inline_image_rel_ids(&self) -> Vec<&str>;

    /// Get relationship IDs for anchored images (`<wp:anchor>` → `<a:blip r:embed="rId"/>`).
    fn anchored_image_rel_ids(&self) -> Vec<&str>;

    /// Get all image relationship IDs (inline + anchored).
    fn all_image_rel_ids(&self) -> Vec<&str>;
}

#[cfg(all(feature = "wml-drawings", feature = "extra-children"))]
impl DrawingExt for types::CTDrawing {
    fn inline_image_rel_ids(&self) -> Vec<&str> {
        let mut ids = Vec::new();
        for child in &self.extra_children {
            if let ooxml_xml::RawXmlNode::Element(elem) = &child.node
                && local_name_of(&elem.name) == "inline"
            {
                collect_blip_rel_ids(elem, &mut ids);
            }
        }
        ids
    }

    fn anchored_image_rel_ids(&self) -> Vec<&str> {
        let mut ids = Vec::new();
        for child in &self.extra_children {
            if let ooxml_xml::RawXmlNode::Element(elem) = &child.node
                && local_name_of(&elem.name) == "anchor"
            {
                collect_blip_rel_ids(elem, &mut ids);
            }
        }
        ids
    }

    fn all_image_rel_ids(&self) -> Vec<&str> {
        let mut ids = self.inline_image_rel_ids();
        ids.extend(self.anchored_image_rel_ids());
        ids
    }
}

/// Extract the local name from a possibly-namespaced XML element name.
/// e.g. "wp:inline" → "inline", "a:blip" → "blip", "blip" → "blip".
#[cfg(all(feature = "wml-drawings", feature = "extra-children"))]
fn local_name_of(name: &str) -> &str {
    name.rsplit(':').next().unwrap_or(name)
}

/// Recursively walk a raw XML element tree and collect `r:embed` attribute values
/// from `<a:blip>` elements.
#[cfg(all(feature = "wml-drawings", feature = "extra-children"))]
fn collect_blip_rel_ids<'a>(elem: &'a ooxml_xml::RawXmlElement, ids: &mut Vec<&'a str>) {
    if local_name_of(&elem.name) == "blip" {
        for (attr_name, attr_val) in &elem.attributes {
            if attr_name == "r:embed" || local_name_of(attr_name) == "embed" {
                ids.push(attr_val.as_str());
            }
        }
    }
    for child in &elem.children {
        if let ooxml_xml::RawXmlNode::Element(child_elem) = child {
            collect_blip_rel_ids(child_elem, ids);
        }
    }
}

// =============================================================================
// RunPropertiesExt
// =============================================================================

/// Extension methods for `RunProperties` (ECMA-376 §17.3.2).
///
/// All toggle property checks follow the OOXML convention: element present
/// without `val` attribute means "on"; explicit `val` of "1"/"true"/"on" means on.
#[cfg(feature = "wml-styling")]
pub trait RunPropertiesExt {
    /// Check if bold is enabled.
    fn is_bold(&self) -> bool;

    /// Check if italic is enabled.
    fn is_italic(&self) -> bool;

    /// Check if any underline is set (not `none`).
    fn is_underline(&self) -> bool;

    /// Get the underline style.
    fn underline_style(&self) -> Option<&types::STUnderline>;

    /// Check if single strikethrough is enabled.
    fn is_strikethrough(&self) -> bool;

    /// Check if double strikethrough is enabled.
    fn is_double_strikethrough(&self) -> bool;

    /// Check if all-caps is enabled.
    fn is_all_caps(&self) -> bool;

    /// Check if small-caps is enabled.
    fn is_small_caps(&self) -> bool;

    /// Check if text is hidden (`<w:vanish/>`).
    fn is_hidden(&self) -> bool;

    /// Get highlight color.
    fn highlight_color(&self) -> Option<&types::STHighlightColor>;

    /// Get vertical alignment (superscript/subscript/baseline).
    fn vertical_alignment(&self) -> Option<&types::STVerticalAlignRun>;

    /// Check if superscript.
    fn is_superscript(&self) -> bool;

    /// Check if subscript.
    fn is_subscript(&self) -> bool;

    /// Get font size in half-points (e.g., 24 = 12pt).
    fn font_size_half_points(&self) -> Option<u32>;

    /// Get font size in points (e.g., 12.0).
    fn font_size_points(&self) -> Option<f64>;

    /// Get text color as hex string (e.g., "FF0000").
    fn color_hex(&self) -> Option<&str>;

    /// Get the referenced character style ID.
    fn style_id(&self) -> Option<&str>;

    /// Get the ASCII font name.
    fn font_ascii(&self) -> Option<&str>;

    /// Check if right-to-left text.
    fn is_rtl(&self) -> bool;
}

#[cfg(feature = "wml-styling")]
impl RunPropertiesExt for types::RunProperties {
    fn is_bold(&self) -> bool {
        is_on(&self.bold)
    }

    fn is_italic(&self) -> bool {
        is_on(&self.italic)
    }

    fn is_underline(&self) -> bool {
        self.underline
            .as_ref()
            .is_some_and(|u| !matches!(u.value, Some(types::STUnderline::None)))
    }

    fn underline_style(&self) -> Option<&types::STUnderline> {
        self.underline.as_ref().and_then(|u| u.value.as_ref())
    }

    fn is_strikethrough(&self) -> bool {
        is_on(&self.strikethrough)
    }

    fn is_double_strikethrough(&self) -> bool {
        is_on(&self.dstrike)
    }

    fn is_all_caps(&self) -> bool {
        is_on(&self.caps)
    }

    fn is_small_caps(&self) -> bool {
        is_on(&self.small_caps)
    }

    fn is_hidden(&self) -> bool {
        is_on(&self.vanish)
    }

    fn highlight_color(&self) -> Option<&types::STHighlightColor> {
        self.highlight.as_ref().map(|h| &h.value)
    }

    fn vertical_alignment(&self) -> Option<&types::STVerticalAlignRun> {
        self.vert_align.as_ref().map(|va| &va.value)
    }

    fn is_superscript(&self) -> bool {
        matches!(
            self.vert_align.as_ref().map(|va| &va.value),
            Some(types::STVerticalAlignRun::Superscript)
        )
    }

    fn is_subscript(&self) -> bool {
        matches!(
            self.vert_align.as_ref().map(|va| &va.value),
            Some(types::STVerticalAlignRun::Subscript)
        )
    }

    fn font_size_half_points(&self) -> Option<u32> {
        self.size
            .as_ref()
            .and_then(|sz| parse_half_points(&sz.value))
    }

    fn font_size_points(&self) -> Option<f64> {
        self.font_size_half_points().map(|hp| hp as f64 / 2.0)
    }

    fn color_hex(&self) -> Option<&str> {
        self.color.as_ref().map(|c| c.value.as_str())
    }

    fn style_id(&self) -> Option<&str> {
        self.run_style.as_ref().map(|s| s.value.as_str())
    }

    fn font_ascii(&self) -> Option<&str> {
        self.fonts.as_ref().and_then(|f| f.ascii.as_deref())
    }

    fn is_rtl(&self) -> bool {
        is_on(&self.rtl)
    }
}

// =============================================================================
// HyperlinkExt
// =============================================================================

/// Extension methods for `Hyperlink`.
pub trait HyperlinkExt {
    /// Get runs contained in this hyperlink.
    fn runs(&self) -> Vec<&types::Run>;

    /// Extract text from this hyperlink.
    fn text(&self) -> String;

    /// Get the anchor string (in-document bookmark reference).
    fn anchor_str(&self) -> Option<&str>;

    /// Get the relationship ID (`r:id` attribute) for external hyperlinks.
    #[cfg(feature = "extra-attrs")]
    fn rel_id(&self) -> Option<&str>;

    /// Check if this is an external hyperlink (has a relationship ID).
    #[cfg(feature = "extra-attrs")]
    fn is_external(&self) -> bool;
}

impl HyperlinkExt for types::Hyperlink {
    fn runs(&self) -> Vec<&types::Run> {
        collect_runs_from_paragraph_content(&self.paragraph_content)
    }

    fn text(&self) -> String {
        let mut out = String::new();
        for item in &self.paragraph_content {
            collect_text_from_paragraph_content(item, &mut out);
        }
        out
    }

    fn anchor_str(&self) -> Option<&str> {
        self.anchor.as_deref()
    }

    #[cfg(feature = "extra-attrs")]
    fn rel_id(&self) -> Option<&str> {
        self.extra_attrs.get("r:id").map(|s| s.as_str())
    }

    #[cfg(feature = "extra-attrs")]
    fn is_external(&self) -> bool {
        self.extra_attrs.contains_key("r:id")
    }
}

// =============================================================================
// TableExt
// =============================================================================

/// Extension methods for `Table`.
pub trait TableExt {
    /// Get all rows in this table.
    fn rows(&self) -> Vec<&types::CTRow>;

    /// Get the number of rows.
    fn row_count(&self) -> usize;

    /// Get table properties.
    fn properties(&self) -> &types::TableProperties;

    /// Extract all text from the table.
    fn text(&self) -> String;
}

impl TableExt for types::Table {
    fn rows(&self) -> Vec<&types::CTRow> {
        self.rows
            .iter()
            .filter_map(|c| match c.as_ref() {
                types::RowContent::Tr(row) => Some(row.as_ref()),
                _ => None,
            })
            .collect()
    }

    fn row_count(&self) -> usize {
        self.rows().len()
    }

    fn properties(&self) -> &types::TableProperties {
        &self.table_properties
    }

    fn text(&self) -> String {
        let row_texts: Vec<String> = self.rows().iter().map(|r| r.text()).collect();
        row_texts.join("\n")
    }
}

// =============================================================================
// RowExt
// =============================================================================

/// Extension methods for `CTRow`.
pub trait RowExt {
    /// Get all cells in this row.
    fn cells(&self) -> Vec<&types::TableCell>;

    /// Get row properties.
    #[cfg(feature = "wml-tables")]
    fn properties(&self) -> Option<&types::TableRowProperties>;

    /// Extract all text from the row.
    fn text(&self) -> String;
}

impl RowExt for types::CTRow {
    fn cells(&self) -> Vec<&types::TableCell> {
        self.cells
            .iter()
            .filter_map(|c| match c.as_ref() {
                types::CellContent::Tc(cell) => Some(cell.as_ref()),
                _ => None,
            })
            .collect()
    }

    #[cfg(feature = "wml-tables")]
    fn properties(&self) -> Option<&types::TableRowProperties> {
        self.row_properties.as_deref()
    }

    fn text(&self) -> String {
        let cell_texts: Vec<String> = self.cells().iter().map(|c| c.text()).collect();
        cell_texts.join("\t")
    }
}

// =============================================================================
// CellExt
// =============================================================================

/// Extension methods for `TableCell`.
pub trait CellExt {
    /// Get all paragraphs in this cell.
    fn paragraphs(&self) -> Vec<&types::Paragraph>;

    /// Get cell properties.
    #[cfg(feature = "wml-tables")]
    fn properties(&self) -> Option<&types::TableCellProperties>;

    /// Extract all text from the cell.
    fn text(&self) -> String;
}

impl CellExt for types::TableCell {
    fn paragraphs(&self) -> Vec<&types::Paragraph> {
        self.block_content
            .iter()
            .filter_map(|elt| match elt.as_ref() {
                types::BlockContent::P(p) => Some(p.as_ref()),
                _ => None,
            })
            .collect()
    }

    #[cfg(feature = "wml-tables")]
    fn properties(&self) -> Option<&types::TableCellProperties> {
        self.cell_properties.as_deref()
    }

    fn text(&self) -> String {
        let texts: Vec<String> = self.paragraphs().iter().map(|p| p.text()).collect();
        texts.join("\n")
    }
}

// =============================================================================
// SectionPropertiesExt
// =============================================================================

/// Extension methods for `SectionProperties` (ECMA-376 §17.6.17).
#[cfg(feature = "wml-layout")]
pub trait SectionPropertiesExt {
    /// Get the page size element.
    fn page_size(&self) -> Option<&types::PageSize>;

    /// Get the page margins element.
    fn page_margins(&self) -> Option<&types::PageMargins>;

    /// Get page width in twips.
    fn page_width_twips(&self) -> Option<u64>;

    /// Get page height in twips.
    fn page_height_twips(&self) -> Option<u64>;

    /// Get page orientation.
    fn page_orientation(&self) -> Option<&types::STPageOrientation>;

    /// Check if the section has a distinct title (first) page.
    fn has_title_page(&self) -> bool;

    /// Get header references (type + relationship ID from extra_attrs).
    #[cfg(feature = "extra-attrs")]
    fn header_references(&self) -> Vec<(&types::STHdrFtr, &str)>;

    /// Get footer references (type + relationship ID from extra_attrs).
    #[cfg(feature = "extra-attrs")]
    fn footer_references(&self) -> Vec<(&types::STHdrFtr, &str)>;
}

#[cfg(feature = "wml-layout")]
impl SectionPropertiesExt for types::SectionProperties {
    fn page_size(&self) -> Option<&types::PageSize> {
        self.pg_sz.as_deref()
    }

    fn page_margins(&self) -> Option<&types::PageMargins> {
        self.pg_mar.as_deref()
    }

    fn page_width_twips(&self) -> Option<u64> {
        self.pg_sz
            .as_ref()
            .and_then(|sz| sz.width.as_ref())
            .and_then(|w| w.parse::<u64>().ok())
    }

    fn page_height_twips(&self) -> Option<u64> {
        self.pg_sz
            .as_ref()
            .and_then(|sz| sz.height.as_ref())
            .and_then(|h| h.parse::<u64>().ok())
    }

    fn page_orientation(&self) -> Option<&types::STPageOrientation> {
        self.pg_sz.as_ref().and_then(|sz| sz.orient.as_ref())
    }

    fn has_title_page(&self) -> bool {
        is_on(&self.title_pg)
    }

    #[cfg(feature = "extra-attrs")]
    fn header_references(&self) -> Vec<(&types::STHdrFtr, &str)> {
        self.header_footer_refs
            .iter()
            .filter_map(|r| match r.as_ref() {
                types::HeaderFooterRef::HeaderReference(h) => {
                    h.extra_attrs.get("r:id").map(|id| (&h.r#type, id.as_str()))
                }
                _ => None,
            })
            .collect()
    }

    #[cfg(feature = "extra-attrs")]
    fn footer_references(&self) -> Vec<(&types::STHdrFtr, &str)> {
        self.header_footer_refs
            .iter()
            .filter_map(|r| match r.as_ref() {
                types::HeaderFooterRef::FooterReference(f) => {
                    f.extra_attrs.get("r:id").map(|id| (&f.r#type, id.as_str()))
                }
                _ => None,
            })
            .collect()
    }
}

// =============================================================================
// Style Resolution
// =============================================================================

/// Context for resolving run properties through the style inheritance chain.
///
/// OOXML styles form a `basedOn` chain. Resolution order (ECMA-376 §17.7.2):
/// 1. Direct run properties on the run
/// 2. Character style (referenced by `rPr/rStyle`)
/// 3. Walk the `basedOn` chain of the character style
/// 4. Document defaults (`docDefaults/rPrDefault/rPr`)
#[cfg(feature = "wml-styling")]
#[derive(Debug, Clone, Default)]
pub struct StyleContext {
    /// Styles indexed by styleId.
    pub styles: std::collections::HashMap<String, types::Style>,
    /// Default run properties from `docDefaults`.
    pub default_run_properties: Option<types::RunProperties>,
}

#[cfg(feature = "wml-styling")]
impl StyleContext {
    /// Build a `StyleContext` from a parsed `Styles` document.
    pub fn from_styles(styles_doc: &types::Styles) -> Self {
        let mut styles = std::collections::HashMap::new();
        for style in &styles_doc.style {
            if let Some(ref id) = style.style_id {
                styles.insert(id.clone(), style.as_ref().clone());
            }
        }

        let default_run_properties = styles_doc
            .doc_defaults
            .as_ref()
            .and_then(|dd| dd.r_pr_default.as_ref())
            .and_then(|rpd| rpd.r_pr.as_ref())
            .map(|rp| rp.as_ref().clone());

        Self {
            styles,
            default_run_properties,
        }
    }

    /// Look up a style by its ID.
    pub fn style(&self, id: &str) -> Option<&types::Style> {
        self.styles.get(id)
    }

    /// Walk the `basedOn` chain for a style, collecting run properties.
    /// Returns properties in order from most derived to least derived.
    /// Depth-limited to 20 to prevent infinite loops.
    fn collect_style_chain_rpr(&self, style_id: &str) -> Vec<&types::RunProperties> {
        let mut result = Vec::new();
        let mut current_id = Some(style_id.to_string());
        let mut depth = 0;

        while let Some(ref id) = current_id {
            if depth >= 20 {
                break;
            }
            if let Some(style) = self.styles.get(id) {
                if let Some(ref rpr) = style.r_pr {
                    result.push(rpr.as_ref());
                }
                current_id = style.based_on.as_ref().map(|b| b.value.clone());
            } else {
                break;
            }
            depth += 1;
        }
        result
    }
}

/// Extension methods for `Run` that resolve formatting through the style chain.
#[cfg(feature = "wml-styling")]
pub trait RunResolveExt {
    /// Resolve bold through direct → style chain → defaults.
    fn resolved_is_bold(&self, ctx: &StyleContext) -> bool;

    /// Resolve italic through direct → style chain → defaults.
    fn resolved_is_italic(&self, ctx: &StyleContext) -> bool;

    /// Resolve font size in half-points through direct → style chain → defaults.
    fn resolved_font_size_half_points(&self, ctx: &StyleContext) -> Option<u32>;

    /// Resolve ASCII font name through direct → style chain → defaults.
    fn resolved_font_ascii(&self, ctx: &StyleContext) -> Option<String>;

    /// Resolve text color hex through direct → style chain → defaults.
    fn resolved_color_hex(&self, ctx: &StyleContext) -> Option<String>;
}

#[cfg(feature = "wml-styling")]
impl RunResolveExt for types::Run {
    fn resolved_is_bold(&self, ctx: &StyleContext) -> bool {
        resolve_toggle(&self.r_pr, ctx, |rpr| &rpr.bold)
    }

    fn resolved_is_italic(&self, ctx: &StyleContext) -> bool {
        resolve_toggle(&self.r_pr, ctx, |rpr| &rpr.italic)
    }

    fn resolved_font_size_half_points(&self, ctx: &StyleContext) -> Option<u32> {
        resolve_option(&self.r_pr, ctx, |rpr| {
            rpr.size
                .as_ref()
                .and_then(|sz| parse_half_points(&sz.value))
        })
    }

    fn resolved_font_ascii(&self, ctx: &StyleContext) -> Option<String> {
        resolve_option(&self.r_pr, ctx, |rpr| {
            rpr.fonts.as_ref().and_then(|f| f.ascii.clone())
        })
    }

    fn resolved_color_hex(&self, ctx: &StyleContext) -> Option<String> {
        resolve_option(&self.r_pr, ctx, |rpr| {
            rpr.color.as_ref().map(|c| c.value.clone())
        })
    }
}

/// Resolve a toggle property through the style chain.
#[cfg(feature = "wml-styling")]
fn resolve_toggle(
    direct_rpr: &Option<Box<types::RunProperties>>,
    ctx: &StyleContext,
    accessor: impl Fn(&types::RunProperties) -> &Option<Box<types::CTOnOff>>,
) -> bool {
    // 1. Direct run properties
    if let Some(rpr) = direct_rpr {
        if let Some(val) = check_toggle(accessor(rpr)) {
            return val;
        }

        // 2. Style chain via rStyle
        if let Some(style_ref) = &rpr.run_style {
            for chain_rpr in ctx.collect_style_chain_rpr(&style_ref.value) {
                if let Some(val) = check_toggle(accessor(chain_rpr)) {
                    return val;
                }
            }
        }
    }

    // 3. Document defaults
    if let Some(defaults) = &ctx.default_run_properties
        && let Some(val) = check_toggle(accessor(defaults))
    {
        return val;
    }

    false
}

/// Resolve an optional property through the style chain.
#[cfg(feature = "wml-styling")]
fn resolve_option<T>(
    direct_rpr: &Option<Box<types::RunProperties>>,
    ctx: &StyleContext,
    accessor: impl Fn(&types::RunProperties) -> Option<T>,
) -> Option<T> {
    // 1. Direct run properties
    if let Some(rpr) = direct_rpr {
        if let val @ Some(_) = accessor(rpr) {
            return val;
        }

        // 2. Style chain via rStyle
        if let Some(style_ref) = &rpr.run_style {
            for chain_rpr in ctx.collect_style_chain_rpr(&style_ref.value) {
                if let val @ Some(_) = accessor(chain_rpr) {
                    return val;
                }
            }
        }
    }

    // 3. Document defaults
    if let Some(defaults) = &ctx.default_run_properties
        && let val @ Some(_) = accessor(defaults)
    {
        return val;
    }

    None
}

// =============================================================================
// Parsing Functions
// =============================================================================

/// Parse a `Document` from XML bytes using the generated `FromXml` parser.
///
/// This is the recommended way to parse document.xml content.
pub fn parse_document(xml: &[u8]) -> Result<types::Document, ParseError> {
    let mut reader = Reader::from_reader(Cursor::new(xml));
    let mut buf = Vec::new();

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) => return types::Document::from_xml(&mut reader, &e, false),
            Ok(Event::Empty(e)) => return types::Document::from_xml(&mut reader, &e, true),
            Ok(Event::Eof) => break,
            Err(e) => return Err(ParseError::Xml(e)),
            _ => {}
        }
        buf.clear();
    }
    Err(ParseError::UnexpectedElement(
        "no document element found".to_string(),
    ))
}

/// Parse a `Styles` document from XML bytes using the generated `FromXml` parser.
///
/// This is the recommended way to parse styles.xml content.
pub fn parse_styles(xml: &[u8]) -> Result<types::Styles, ParseError> {
    let mut reader = Reader::from_reader(Cursor::new(xml));
    let mut buf = Vec::new();

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) => return types::Styles::from_xml(&mut reader, &e, false),
            Ok(Event::Empty(e)) => return types::Styles::from_xml(&mut reader, &e, true),
            Ok(Event::Eof) => break,
            Err(e) => return Err(ParseError::Xml(e)),
            _ => {}
        }
        buf.clear();
    }
    Err(ParseError::UnexpectedElement(
        "no styles element found".to_string(),
    ))
}

/// Parse a header or footer from XML bytes using the generated `FromXml` parser.
pub fn parse_hdr_ftr(xml: &[u8]) -> Result<types::HeaderFooter, ParseError> {
    let mut reader = Reader::from_reader(Cursor::new(xml));
    let mut buf = Vec::new();

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) => return types::HeaderFooter::from_xml(&mut reader, &e, false),
            Ok(Event::Empty(e)) => return types::HeaderFooter::from_xml(&mut reader, &e, true),
            Ok(Event::Eof) => break,
            Err(e) => return Err(ParseError::Xml(e)),
            _ => {}
        }
        buf.clear();
    }
    Err(ParseError::UnexpectedElement(
        "no header/footer element found".to_string(),
    ))
}

/// Parse footnotes from XML bytes using the generated `FromXml` parser.
pub fn parse_footnotes(xml: &[u8]) -> Result<types::Footnotes, ParseError> {
    let mut reader = Reader::from_reader(Cursor::new(xml));
    let mut buf = Vec::new();

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) => return types::Footnotes::from_xml(&mut reader, &e, false),
            Ok(Event::Empty(e)) => return types::Footnotes::from_xml(&mut reader, &e, true),
            Ok(Event::Eof) => break,
            Err(e) => return Err(ParseError::Xml(e)),
            _ => {}
        }
        buf.clear();
    }
    Err(ParseError::UnexpectedElement(
        "no footnotes element found".to_string(),
    ))
}

/// Parse endnotes from XML bytes using the generated `FromXml` parser.
pub fn parse_endnotes(xml: &[u8]) -> Result<types::Endnotes, ParseError> {
    let mut reader = Reader::from_reader(Cursor::new(xml));
    let mut buf = Vec::new();

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) => return types::Endnotes::from_xml(&mut reader, &e, false),
            Ok(Event::Empty(e)) => return types::Endnotes::from_xml(&mut reader, &e, true),
            Ok(Event::Eof) => break,
            Err(e) => return Err(ParseError::Xml(e)),
            _ => {}
        }
        buf.clear();
    }
    Err(ParseError::UnexpectedElement(
        "no endnotes element found".to_string(),
    ))
}

/// Parse comments from XML bytes using the generated `FromXml` parser.
pub fn parse_comments(xml: &[u8]) -> Result<types::Comments, ParseError> {
    let mut reader = Reader::from_reader(Cursor::new(xml));
    let mut buf = Vec::new();

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) => return types::Comments::from_xml(&mut reader, &e, false),
            Ok(Event::Empty(e)) => return types::Comments::from_xml(&mut reader, &e, true),
            Ok(Event::Eof) => break,
            Err(e) => return Err(ParseError::Xml(e)),
            _ => {}
        }
        buf.clear();
    }
    Err(ParseError::UnexpectedElement(
        "no comments element found".to_string(),
    ))
}

// =============================================================================
// ResolvedDocument
// =============================================================================

/// A document with bound style context for convenient resolved access.
///
/// Wraps a generated `types::Document` and provides methods that automatically
/// resolve formatting through the style chain.
///
/// # Example
///
/// ```ignore
/// use ooxml_wml::ext::{ResolvedDocument, parse_document, parse_styles};
///
/// let doc = parse_document(doc_xml)?;
/// let styles = parse_styles(styles_xml)?;
/// let resolved = ResolvedDocument::new(doc, styles);
///
/// if let Some(body) = resolved.body() {
///     for para in body.paragraphs() {
///         println!("{}", para.text());
///     }
/// }
/// ```
#[cfg(feature = "wml-styling")]
pub struct ResolvedDocument {
    document: types::Document,
    context: StyleContext,
}

#[cfg(feature = "wml-styling")]
impl ResolvedDocument {
    /// Create a new resolved document from a parsed document and styles.
    pub fn new(document: types::Document, styles: types::Styles) -> Self {
        let context = StyleContext::from_styles(&styles);
        Self { document, context }
    }

    /// Create from a document with an existing style context.
    pub fn with_context(document: types::Document, context: StyleContext) -> Self {
        Self { document, context }
    }

    /// Get the underlying document.
    pub fn document(&self) -> &types::Document {
        &self.document
    }

    /// Get the style context.
    pub fn context(&self) -> &StyleContext {
        &self.context
    }

    /// Get the document body.
    pub fn body(&self) -> Option<&types::Body> {
        self.document.body()
    }

    /// Extract all text from the document.
    pub fn text(&self) -> String {
        self.document.body().map(|b| b.text()).unwrap_or_default()
    }

    /// Check if a run is bold (resolved through style chain).
    pub fn is_bold(&self, run: &types::Run) -> bool {
        run.resolved_is_bold(&self.context)
    }

    /// Check if a run is italic (resolved through style chain).
    pub fn is_italic(&self, run: &types::Run) -> bool {
        run.resolved_is_italic(&self.context)
    }

    /// Get resolved font size in half-points.
    pub fn font_size_half_points(&self, run: &types::Run) -> Option<u32> {
        run.resolved_font_size_half_points(&self.context)
    }

    /// Get resolved ASCII font name.
    pub fn font_ascii(&self, run: &types::Run) -> Option<String> {
        run.resolved_font_ascii(&self.context)
    }

    /// Get resolved text color hex.
    pub fn color_hex(&self, run: &types::Run) -> Option<String> {
        run.resolved_color_hex(&self.context)
    }
}

// =============================================================================
// Tests
// =============================================================================

#[cfg(test)]
#[allow(clippy::vec_box)]
mod tests {
    use super::*;

    // -------------------------------------------------------------------------
    // Helper tests
    // -------------------------------------------------------------------------

    #[test]
    fn test_is_on_none() {
        assert!(!is_on(&None));
    }

    #[test]
    fn test_is_on_present_no_val() {
        // Element present with no val attribute → on
        let field = Some(Box::new(types::CTOnOff {
            value: None,
            #[cfg(feature = "extra-attrs")]
            extra_attrs: Default::default(),
        }));
        assert!(is_on(&field));
    }

    #[test]
    fn test_is_on_explicit_true() {
        for val in &["1", "true", "on"] {
            let field = Some(Box::new(types::CTOnOff {
                value: Some(val.to_string()),
                #[cfg(feature = "extra-attrs")]
                extra_attrs: Default::default(),
            }));
            assert!(is_on(&field), "expected is_on for val={val}");
        }
    }

    #[test]
    fn test_is_on_explicit_false() {
        for val in &["0", "false", "off"] {
            let field = Some(Box::new(types::CTOnOff {
                value: Some(val.to_string()),
                #[cfg(feature = "extra-attrs")]
                extra_attrs: Default::default(),
            }));
            assert!(!is_on(&field), "expected !is_on for val={val}");
        }
    }

    #[test]
    fn test_check_toggle_none() {
        assert_eq!(check_toggle(&None), None);
    }

    #[test]
    fn test_check_toggle_present() {
        let field = Some(Box::new(types::CTOnOff {
            value: None,
            #[cfg(feature = "extra-attrs")]
            extra_attrs: Default::default(),
        }));
        assert_eq!(check_toggle(&field), Some(true));
    }

    #[test]
    fn test_parse_half_points() {
        assert_eq!(parse_half_points("24"), Some(24));
        assert_eq!(parse_half_points("0"), Some(0));
        assert_eq!(parse_half_points("abc"), None);
        assert_eq!(parse_half_points(""), None);
    }

    // -------------------------------------------------------------------------
    // RunPropertiesExt tests
    // -------------------------------------------------------------------------

    #[cfg(feature = "wml-styling")]
    fn make_run_properties() -> types::RunProperties {
        types::RunProperties {
            run_style: None,
            fonts: None,
            bold: None,
            b_cs: None,
            italic: None,
            i_cs: None,
            caps: None,
            small_caps: None,
            strikethrough: None,
            dstrike: None,
            outline: None,
            shadow: None,
            emboss: None,
            imprint: None,
            no_proof: None,
            snap_to_grid: None,
            vanish: None,
            web_hidden: None,
            color: None,
            spacing: None,
            width: None,
            kern: None,
            position: None,
            size: None,
            size_complex_script: None,
            highlight: None,
            underline: None,
            effect: None,
            bdr: None,
            shading: None,
            fit_text: None,
            vert_align: None,
            rtl: None,
            cs: None,
            em: None,
            lang: None,
            east_asian_layout: None,
            spec_vanish: None,
            o_math: None,
            r_pr_change: None,
            #[cfg(feature = "extra-children")]
            extra_children: Default::default(),
        }
    }

    #[cfg(feature = "wml-styling")]
    fn on_off(val: Option<&str>) -> Option<Box<types::CTOnOff>> {
        Some(Box::new(types::CTOnOff {
            value: val.map(|v| v.to_string()),
            #[cfg(feature = "extra-attrs")]
            extra_attrs: Default::default(),
        }))
    }

    #[test]
    #[cfg(feature = "wml-styling")]
    fn test_rpr_bold_italic() {
        let mut rpr = make_run_properties();
        assert!(!rpr.is_bold());
        assert!(!rpr.is_italic());

        rpr.bold = on_off(None); // present, no val → on
        rpr.italic = on_off(Some("true"));
        assert!(rpr.is_bold());
        assert!(rpr.is_italic());
    }

    #[test]
    #[cfg(feature = "wml-styling")]
    fn test_rpr_underline() {
        let mut rpr = make_run_properties();
        assert!(!rpr.is_underline());
        assert!(rpr.underline_style().is_none());

        rpr.underline = Some(Box::new(types::CTUnderline {
            value: Some(types::STUnderline::Single),
            color: None,
            theme_color: None,
            theme_tint: None,
            theme_shade: None,
            #[cfg(feature = "extra-attrs")]
            extra_attrs: Default::default(),
        }));
        assert!(rpr.is_underline());
        assert_eq!(rpr.underline_style(), Some(&types::STUnderline::Single));

        // "none" underline should not count as underlined
        rpr.underline = Some(Box::new(types::CTUnderline {
            value: Some(types::STUnderline::None),
            color: None,
            theme_color: None,
            theme_tint: None,
            theme_shade: None,
            #[cfg(feature = "extra-attrs")]
            extra_attrs: Default::default(),
        }));
        assert!(!rpr.is_underline());
    }

    #[test]
    #[cfg(feature = "wml-styling")]
    fn test_rpr_strikethrough() {
        let mut rpr = make_run_properties();
        rpr.strikethrough = on_off(None);
        assert!(rpr.is_strikethrough());
        assert!(!rpr.is_double_strikethrough());

        rpr.strikethrough = None;
        rpr.dstrike = on_off(Some("1"));
        assert!(!rpr.is_strikethrough());
        assert!(rpr.is_double_strikethrough());
    }

    #[test]
    #[cfg(feature = "wml-styling")]
    fn test_rpr_caps_hidden() {
        let mut rpr = make_run_properties();
        rpr.caps = on_off(None);
        rpr.vanish = on_off(Some("1"));
        assert!(rpr.is_all_caps());
        assert!(!rpr.is_small_caps());
        assert!(rpr.is_hidden());
    }

    #[test]
    #[cfg(feature = "wml-styling")]
    fn test_rpr_font_size() {
        let mut rpr = make_run_properties();
        assert!(rpr.font_size_half_points().is_none());

        rpr.size = Some(Box::new(types::CTHpsMeasure {
            value: "24".to_string(),
            #[cfg(feature = "extra-attrs")]
            extra_attrs: Default::default(),
        }));
        assert_eq!(rpr.font_size_half_points(), Some(24));
        assert_eq!(rpr.font_size_points(), Some(12.0));
    }

    #[test]
    #[cfg(feature = "wml-styling")]
    fn test_rpr_color() {
        let mut rpr = make_run_properties();
        assert!(rpr.color_hex().is_none());

        rpr.color = Some(Box::new(types::CTColor {
            value: "FF0000".to_string(),
            theme_color: None,
            theme_tint: None,
            theme_shade: None,
            #[cfg(feature = "extra-attrs")]
            extra_attrs: Default::default(),
        }));
        assert_eq!(rpr.color_hex(), Some("FF0000"));
    }

    #[test]
    #[cfg(feature = "wml-styling")]
    fn test_rpr_vertical_alignment() {
        let mut rpr = make_run_properties();
        assert!(!rpr.is_superscript());
        assert!(!rpr.is_subscript());

        rpr.vert_align = Some(Box::new(types::CTVerticalAlignRun {
            value: types::STVerticalAlignRun::Superscript,
            #[cfg(feature = "extra-attrs")]
            extra_attrs: Default::default(),
        }));
        assert!(rpr.is_superscript());
        assert!(!rpr.is_subscript());
    }

    #[test]
    #[cfg(feature = "wml-styling")]
    fn test_rpr_font_ascii() {
        let mut rpr = make_run_properties();
        assert!(rpr.font_ascii().is_none());

        rpr.fonts = Some(Box::new(types::Fonts {
            hint: None,
            ascii: Some("Arial".to_string()),
            h_ansi: None,
            east_asia: None,
            cs: None,
            ascii_theme: None,
            h_ansi_theme: None,
            east_asia_theme: None,
            cstheme: None,
            #[cfg(feature = "extra-attrs")]
            extra_attrs: Default::default(),
        }));
        assert_eq!(rpr.font_ascii(), Some("Arial"));
    }

    // -------------------------------------------------------------------------
    // RunExt tests
    // -------------------------------------------------------------------------

    fn make_text(s: &str) -> Box<types::RunContent> {
        Box::new(types::RunContent::T(Box::new(types::Text {
            text: Some(s.to_string()),
            #[cfg(feature = "extra-children")]
            extra_children: Default::default(),
        })))
    }

    fn make_tab() -> Box<types::RunContent> {
        Box::new(types::RunContent::Tab(Box::new(types::CTEmpty)))
    }

    fn make_br(br_type: Option<types::STBrType>) -> Box<types::RunContent> {
        Box::new(types::RunContent::Br(Box::new(types::CTBr {
            r#type: br_type,
            clear: None,
            #[cfg(feature = "extra-attrs")]
            extra_attrs: Default::default(),
        })))
    }

    fn make_cr() -> Box<types::RunContent> {
        Box::new(types::RunContent::Cr(Box::new(types::CTEmpty)))
    }

    fn make_run(content: Vec<Box<types::RunContent>>) -> types::Run {
        types::Run {
            rsid_r_pr: None,
            rsid_del: None,
            rsid_r: None,
            #[cfg(feature = "wml-styling")]
            r_pr: None,
            run_content: content,
            #[cfg(feature = "extra-attrs")]
            extra_attrs: Default::default(),
            #[cfg(feature = "extra-children")]
            extra_children: Default::default(),
        }
    }

    #[test]
    fn test_run_text_simple() {
        let run = make_run(vec![make_text("Hello"), make_text(" World")]);
        assert_eq!(run.text(), "Hello World");
    }

    #[test]
    fn test_run_text_with_tab_and_break() {
        let run = make_run(vec![
            make_text("A"),
            make_tab(),
            make_text("B"),
            make_br(None), // text wrapping break → newline
            make_text("C"),
        ]);
        assert_eq!(run.text(), "A\tB\nC");
    }

    #[test]
    fn test_run_text_page_break_not_text() {
        let run = make_run(vec![
            make_text("Before"),
            make_br(Some(types::STBrType::Page)),
            make_text("After"),
        ]);
        // Page breaks should not produce text
        assert_eq!(run.text(), "BeforeAfter");
        assert!(run.has_page_break());
    }

    #[test]
    fn test_run_text_cr() {
        let run = make_run(vec![make_text("A"), make_cr(), make_text("B")]);
        assert_eq!(run.text(), "A\nB");
    }

    #[test]
    fn test_run_no_page_break() {
        let run = make_run(vec![make_text("Hello")]);
        assert!(!run.has_page_break());
    }

    // -------------------------------------------------------------------------
    // ParagraphExt tests
    // -------------------------------------------------------------------------

    fn make_p_run(text: &str) -> Box<types::ParagraphContent> {
        Box::new(types::ParagraphContent::R(Box::new(make_run(vec![
            make_text(text),
        ]))))
    }

    fn make_paragraph(content: Vec<Box<types::ParagraphContent>>) -> types::Paragraph {
        types::Paragraph {
            rsid_r_pr: None,
            rsid_r: None,
            rsid_del: None,
            rsid_p: None,
            rsid_r_default: None,
            #[cfg(feature = "wml-styling")]
            p_pr: None,
            paragraph_content: content,
            #[cfg(feature = "extra-attrs")]
            extra_attrs: Default::default(),
            #[cfg(feature = "extra-children")]
            extra_children: Default::default(),
        }
    }

    #[test]
    fn test_paragraph_runs_and_text() {
        let para = make_paragraph(vec![make_p_run("Hello "), make_p_run("World")]);
        assert_eq!(para.runs().len(), 2);
        assert_eq!(para.text(), "Hello World");
    }

    #[test]
    fn test_paragraph_with_hyperlink() {
        let hyperlink = Box::new(types::ParagraphContent::Hyperlink(Box::new(
            types::Hyperlink {
                tgt_frame: None,
                tooltip: None,
                doc_location: None,
                history: None,
                anchor: Some("bookmark1".to_string()),
                paragraph_content: vec![make_p_run("link text")],
                #[cfg(feature = "extra-attrs")]
                extra_attrs: Default::default(),
                #[cfg(feature = "extra-children")]
                extra_children: Default::default(),
            },
        )));
        let para = make_paragraph(vec![make_p_run("Click "), hyperlink]);
        assert_eq!(para.runs().len(), 2);
        assert_eq!(para.text(), "Click link text");
        assert_eq!(para.hyperlinks().len(), 1);
        assert_eq!(para.hyperlinks()[0].anchor_str(), Some("bookmark1"));
    }

    #[test]
    fn test_paragraph_with_fld_simple() {
        let fld = Box::new(types::ParagraphContent::FldSimple(Box::new(
            types::CTSimpleField {
                instr: "PAGE".to_string(),
                fld_lock: None,
                dirty: None,
                fld_data: None,
                paragraph_content: vec![make_p_run("1")],
                #[cfg(feature = "extra-attrs")]
                extra_attrs: Default::default(),
                #[cfg(feature = "extra-children")]
                extra_children: Default::default(),
            },
        )));
        let para = make_paragraph(vec![make_p_run("Page "), fld]);
        assert_eq!(para.runs().len(), 2);
        assert_eq!(para.text(), "Page 1");
    }

    // -------------------------------------------------------------------------
    // BodyExt tests
    // -------------------------------------------------------------------------

    fn make_body(content: Vec<Box<types::BlockContent>>) -> types::Body {
        types::Body {
            block_content: content,
            #[cfg(feature = "wml-layout")]
            sect_pr: None,
            #[cfg(feature = "extra-children")]
            extra_children: Default::default(),
        }
    }

    #[test]
    fn test_body_paragraphs() {
        let p1 = Box::new(types::BlockContent::P(Box::new(make_paragraph(vec![
            make_p_run("First"),
        ]))));
        let p2 = Box::new(types::BlockContent::P(Box::new(make_paragraph(vec![
            make_p_run("Second"),
        ]))));
        let body = make_body(vec![p1, p2]);
        assert_eq!(body.paragraphs().len(), 2);
        assert_eq!(body.text(), "First\nSecond");
    }

    #[test]
    fn test_body_tables() {
        let tbl = Box::new(types::BlockContent::Tbl(Box::new(types::Table {
            range_markup: vec![],
            table_properties: Box::default(),
            tbl_grid: Box::default(),
            rows: vec![],
            #[cfg(feature = "extra-children")]
            extra_children: Default::default(),
        })));
        let body = make_body(vec![tbl]);
        assert_eq!(body.tables().len(), 1);
        assert_eq!(body.paragraphs().len(), 0);
    }

    // -------------------------------------------------------------------------
    // DocumentExt tests
    // -------------------------------------------------------------------------

    #[test]
    fn test_document_ext_body() {
        let doc = types::Document {
            background: None,
            body: Some(Box::new(make_body(vec![]))),
            conformance: None,
            #[cfg(feature = "extra-attrs")]
            extra_attrs: Default::default(),
            #[cfg(feature = "extra-children")]
            extra_children: Default::default(),
        };
        assert!(doc.body().is_some());

        let doc_no_body = types::Document {
            background: None,
            body: None,
            conformance: None,
            #[cfg(feature = "extra-attrs")]
            extra_attrs: Default::default(),
            #[cfg(feature = "extra-children")]
            extra_children: Default::default(),
        };
        assert!(doc_no_body.body().is_none());
    }

    // -------------------------------------------------------------------------
    // HyperlinkExt tests
    // -------------------------------------------------------------------------

    #[test]
    fn test_hyperlink_ext() {
        let h = types::Hyperlink {
            tgt_frame: None,
            tooltip: None,
            doc_location: None,
            history: None,
            anchor: Some("top".to_string()),
            paragraph_content: vec![make_p_run("click"), make_p_run(" here")],
            #[cfg(feature = "extra-attrs")]
            extra_attrs: Default::default(),
            #[cfg(feature = "extra-children")]
            extra_children: Default::default(),
        };
        assert_eq!(h.runs().len(), 2);
        assert_eq!(h.text(), "click here");
        assert_eq!(h.anchor_str(), Some("top"));
    }

    // -------------------------------------------------------------------------
    // Table/Row/Cell tests
    // -------------------------------------------------------------------------

    fn make_table_cell(text: &str) -> Box<types::CellContent> {
        Box::new(types::CellContent::Tc(Box::new(types::TableCell {
            id: None,
            cell_properties: None,
            block_content: vec![Box::new(types::BlockContent::P(Box::new(make_paragraph(
                vec![make_p_run(text)],
            ))))],
            #[cfg(feature = "extra-attrs")]
            extra_attrs: Default::default(),
            #[cfg(feature = "extra-children")]
            extra_children: Default::default(),
        })))
    }

    fn make_table_row(cells: Vec<Box<types::CellContent>>) -> Box<types::RowContent> {
        Box::new(types::RowContent::Tr(Box::new(types::CTRow {
            rsid_r_pr: None,
            rsid_r: None,
            rsid_del: None,
            rsid_tr: None,
            tbl_pr_ex: None,
            row_properties: None,
            cells,
            #[cfg(feature = "extra-attrs")]
            extra_attrs: Default::default(),
            #[cfg(feature = "extra-children")]
            extra_children: Default::default(),
        })))
    }

    fn make_table(rows: Vec<Box<types::RowContent>>) -> types::Table {
        types::Table {
            range_markup: vec![],
            table_properties: Box::default(),
            tbl_grid: Box::default(),
            rows,
            #[cfg(feature = "extra-children")]
            extra_children: Default::default(),
        }
    }

    #[test]
    fn test_table_rows_and_text() {
        let tbl = make_table(vec![
            make_table_row(vec![make_table_cell("A1"), make_table_cell("B1")]),
            make_table_row(vec![make_table_cell("A2"), make_table_cell("B2")]),
        ]);
        assert_eq!(tbl.row_count(), 2);
        assert_eq!(tbl.rows().len(), 2);
        assert_eq!(tbl.text(), "A1\tB1\nA2\tB2");
    }

    #[test]
    fn test_row_cells_and_text() {
        let row = types::CTRow {
            rsid_r_pr: None,
            rsid_r: None,
            rsid_del: None,
            rsid_tr: None,
            tbl_pr_ex: None,
            row_properties: None,
            cells: vec![make_table_cell("X"), make_table_cell("Y")],
            #[cfg(feature = "extra-attrs")]
            extra_attrs: Default::default(),
            #[cfg(feature = "extra-children")]
            extra_children: Default::default(),
        };
        assert_eq!(row.cells().len(), 2);
        assert_eq!(row.text(), "X\tY");
    }

    #[test]
    fn test_cell_paragraphs_and_text() {
        let cell = types::TableCell {
            id: None,
            cell_properties: None,
            block_content: vec![
                Box::new(types::BlockContent::P(Box::new(make_paragraph(vec![
                    make_p_run("Line 1"),
                ])))),
                Box::new(types::BlockContent::P(Box::new(make_paragraph(vec![
                    make_p_run("Line 2"),
                ])))),
            ],
            #[cfg(feature = "extra-attrs")]
            extra_attrs: Default::default(),
            #[cfg(feature = "extra-children")]
            extra_children: Default::default(),
        };
        assert_eq!(cell.paragraphs().len(), 2);
        assert_eq!(cell.text(), "Line 1\nLine 2");
    }

    // -------------------------------------------------------------------------
    // SectionPropertiesExt tests
    // -------------------------------------------------------------------------

    #[test]
    #[cfg(feature = "wml-layout")]
    fn test_section_properties_ext() {
        let sect_pr = types::SectionProperties {
            rsid_r_pr: None,
            rsid_del: None,
            rsid_r: None,
            rsid_sect: None,
            header_footer_refs: vec![],
            footnote_pr: None,
            endnote_pr: None,
            r#type: None,
            pg_sz: Some(Box::new(types::PageSize {
                width: Some("12240".to_string()),
                height: Some("15840".to_string()),
                orient: Some(types::STPageOrientation::Portrait),
                code: None,
                #[cfg(feature = "extra-attrs")]
                extra_attrs: Default::default(),
            })),
            pg_mar: Some(Box::new(types::PageMargins {
                top: "1440".to_string(),
                right: "1440".to_string(),
                bottom: "1440".to_string(),
                left: "1440".to_string(),
                header: "720".to_string(),
                footer: "720".to_string(),
                gutter: "0".to_string(),
                #[cfg(feature = "extra-attrs")]
                extra_attrs: Default::default(),
            })),
            paper_src: None,
            pg_borders: None,
            ln_num_type: None,
            pg_num_type: None,
            cols: None,
            form_prot: None,
            v_align: None,
            no_endnote: None,
            title_pg: on_off(None),
            text_direction: None,
            bidi: None,
            rtl_gutter: None,
            doc_grid: None,
            printer_settings: None,
            sect_pr_change: None,
            #[cfg(feature = "extra-attrs")]
            extra_attrs: Default::default(),
            #[cfg(feature = "extra-children")]
            extra_children: Default::default(),
        };

        assert_eq!(sect_pr.page_width_twips(), Some(12240));
        assert_eq!(sect_pr.page_height_twips(), Some(15840));
        assert_eq!(
            sect_pr.page_orientation(),
            Some(&types::STPageOrientation::Portrait)
        );
        assert!(sect_pr.has_title_page());
        assert!(sect_pr.page_size().is_some());
        assert!(sect_pr.page_margins().is_some());
    }

    // -------------------------------------------------------------------------
    // Parsing tests
    // -------------------------------------------------------------------------

    #[test]
    fn test_parse_document_simple() {
        // Generated parsers match on unprefixed element names, so use default namespace
        let xml = br#"<?xml version="1.0" encoding="UTF-8"?>
        <document xmlns="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
            <body>
                <p>
                    <r>
                        <t>Hello World</t>
                    </r>
                </p>
            </body>
        </document>"#;

        let doc = parse_document(xml).expect("parse_document failed");
        let body = doc.body().expect("body should exist");
        let paragraphs = body.paragraphs();
        assert_eq!(paragraphs.len(), 1);
        assert_eq!(paragraphs[0].text(), "Hello World");
    }

    #[test]
    fn test_parse_document_multiple_paragraphs() {
        let xml = br#"<?xml version="1.0" encoding="UTF-8"?>
        <document xmlns="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
            <body>
                <p>
                    <r><t>First</t></r>
                </p>
                <p>
                    <r><t>Second</t></r>
                </p>
            </body>
        </document>"#;

        let doc = parse_document(xml).expect("parse failed");
        let body = doc.body().expect("body");
        assert_eq!(body.paragraphs().len(), 2);
        assert_eq!(body.text(), "First\nSecond");
    }

    #[test]
    fn test_parse_styles_basic() {
        let xml = br#"<?xml version="1.0" encoding="UTF-8"?>
        <styles xmlns="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
            <style type="character" styleId="BoldStyle">
                <name val="Bold Style"/>
                <rPr>
                    <b/>
                </rPr>
            </style>
        </styles>"#;

        let styles = parse_styles(xml).expect("parse_styles failed");
        assert_eq!(styles.style.len(), 1);
        assert_eq!(styles.style[0].style_id.as_deref(), Some("BoldStyle"));
    }

    #[test]
    fn test_parse_document_no_element() {
        let xml = br#"<?xml version="1.0" encoding="UTF-8"?>"#;
        assert!(parse_document(xml).is_err());
    }

    // -------------------------------------------------------------------------
    // StyleContext + RunResolveExt tests
    // -------------------------------------------------------------------------

    #[test]
    #[cfg(feature = "wml-styling")]
    fn test_style_context_from_styles() {
        let xml = br#"<?xml version="1.0" encoding="UTF-8"?>
        <styles xmlns="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
            <docDefaults>
                <rPrDefault>
                    <rPr>
                        <sz val="24"/>
                    </rPr>
                </rPrDefault>
            </docDefaults>
            <style type="character" styleId="Strong">
                <name val="Strong"/>
                <rPr>
                    <b/>
                </rPr>
            </style>
        </styles>"#;

        let styles = parse_styles(xml).expect("parse");
        let ctx = StyleContext::from_styles(&styles);

        assert!(ctx.style("Strong").is_some());
        assert!(ctx.style("Nonexistent").is_none());
        assert!(ctx.default_run_properties.is_some());
        assert_eq!(
            ctx.default_run_properties
                .as_ref()
                .unwrap()
                .font_size_half_points(),
            Some(24)
        );
    }

    #[test]
    #[cfg(feature = "wml-styling")]
    fn test_resolve_bold_from_direct() {
        let run = types::Run {
            rsid_r_pr: None,
            rsid_del: None,
            rsid_r: None,
            r_pr: Some(Box::new({
                let mut rpr = make_run_properties();
                rpr.bold = on_off(None);
                rpr
            })),
            run_content: vec![make_text("bold")],
            #[cfg(feature = "extra-attrs")]
            extra_attrs: Default::default(),
            #[cfg(feature = "extra-children")]
            extra_children: Default::default(),
        };

        let ctx = StyleContext::default();
        assert!(run.resolved_is_bold(&ctx));
        assert!(!run.resolved_is_italic(&ctx));
    }

    #[test]
    #[cfg(feature = "wml-styling")]
    fn test_resolve_bold_from_style_chain() {
        // Set up: run references style "Emphasis" which is basedOn "Strong" which has bold
        let xml = br#"<?xml version="1.0" encoding="UTF-8"?>
        <styles xmlns="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
            <style type="character" styleId="Strong">
                <name val="Strong"/>
                <rPr>
                    <b/>
                    <sz val="28"/>
                </rPr>
            </style>
            <style type="character" styleId="Emphasis">
                <name val="Emphasis"/>
                <basedOn val="Strong"/>
                <rPr>
                    <i/>
                </rPr>
            </style>
        </styles>"#;

        let styles = parse_styles(xml).expect("parse");
        let ctx = StyleContext::from_styles(&styles);

        // Run references "Emphasis" style (which has italic, inherits bold from Strong)
        let run = types::Run {
            rsid_r_pr: None,
            rsid_del: None,
            rsid_r: None,
            r_pr: Some(Box::new({
                let mut rpr = make_run_properties();
                rpr.run_style = Some(Box::new(types::CTString {
                    value: "Emphasis".to_string(),
                    #[cfg(feature = "extra-attrs")]
                    extra_attrs: Default::default(),
                }));
                rpr
            })),
            run_content: vec![make_text("styled")],
            #[cfg(feature = "extra-attrs")]
            extra_attrs: Default::default(),
            #[cfg(feature = "extra-children")]
            extra_children: Default::default(),
        };

        assert!(run.resolved_is_bold(&ctx));
        assert!(run.resolved_is_italic(&ctx));
        assert_eq!(run.resolved_font_size_half_points(&ctx), Some(28));
    }

    #[test]
    #[cfg(feature = "wml-styling")]
    fn test_resolve_from_doc_defaults() {
        let xml = br#"<?xml version="1.0" encoding="UTF-8"?>
        <styles xmlns="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
            <docDefaults>
                <rPrDefault>
                    <rPr>
                        <sz val="22"/>
                        <rFonts ascii="Calibri"/>
                    </rPr>
                </rPrDefault>
            </docDefaults>
        </styles>"#;

        let styles = parse_styles(xml).expect("parse");
        let ctx = StyleContext::from_styles(&styles);

        // Run with no direct properties or style reference
        let run = types::Run {
            rsid_r_pr: None,
            rsid_del: None,
            rsid_r: None,
            r_pr: None,
            run_content: vec![make_text("default")],
            #[cfg(feature = "extra-attrs")]
            extra_attrs: Default::default(),
            #[cfg(feature = "extra-children")]
            extra_children: Default::default(),
        };

        assert!(!run.resolved_is_bold(&ctx));
        assert_eq!(run.resolved_font_size_half_points(&ctx), Some(22));
        assert_eq!(run.resolved_font_ascii(&ctx), Some("Calibri".to_string()));
    }

    #[test]
    #[cfg(feature = "wml-styling")]
    fn test_resolved_document() {
        let doc_xml = br#"<?xml version="1.0" encoding="UTF-8"?>
        <document xmlns="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
            <body>
                <p>
                    <r>
                        <rPr><b/></rPr>
                        <t>Bold text</t>
                    </r>
                </p>
            </body>
        </document>"#;

        let styles_xml = br#"<?xml version="1.0" encoding="UTF-8"?>
        <styles xmlns="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
        </styles>"#;

        let doc = parse_document(doc_xml).expect("parse doc");
        let styles = parse_styles(styles_xml).expect("parse styles");
        let resolved = ResolvedDocument::new(doc, styles);

        assert_eq!(resolved.text(), "Bold text");

        let body = resolved.body().expect("body");
        let paras = body.paragraphs();
        let runs = paras[0].runs();
        assert!(resolved.is_bold(runs[0]));
        assert!(!resolved.is_italic(runs[0]));
    }
}
