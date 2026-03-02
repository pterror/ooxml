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

/// Check if a `OnOffElement` field represents "on" (ECMA-376 §17.17.4).
///
/// An omitted `val` attribute means "true" (the element's presence is the toggle).
/// Explicit values: "1", "true", "on" → true; "0", "false", "off" → false.
#[cfg_attr(
    not(any(feature = "wml-styling", feature = "wml-layout")),
    allow(dead_code)
)]
fn is_on(field: &Option<Box<types::OnOffElement>>) -> bool {
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
fn check_toggle(field: &Option<Box<types::OnOffElement>>) -> Option<bool> {
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

/// Parse a twips measurement string (signed or unsigned).
#[cfg(feature = "wml-styling")]
fn parse_twips(s: &str) -> Option<i64> {
    s.parse::<i64>().ok()
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
// Table of Contents types
// =============================================================================

/// A single entry in a Table of Contents (ECMA-376 §17.12).
///
/// TOC entries use paragraph styles "TOC 1" through "TOC 9" (or "toc1"–"toc9").
/// The `level` is derived from the numeral in the style name.
#[derive(Debug, Clone, PartialEq)]
pub struct TocEntry {
    /// Heading level (1–9), derived from the paragraph style name.
    pub level: u8,
    /// Display text of the entry, extracted from paragraph runs.
    pub text: String,
    /// Page number, if present as the last numeric token in the paragraph.
    /// May be `None` or `0` for unsaved or newly-created documents.
    pub page: Option<u32>,
    /// Bookmark name if the entry is hyperlinked to a heading.
    /// Found in `ParagraphContent::BookmarkStart` inside the paragraph.
    pub bookmark: Option<String>,
}

/// A parsed Table of Contents (ECMA-376 §17.12.1).
///
/// Returned by [`BodyExt::table_of_contents`].
#[derive(Debug, Clone, PartialEq)]
pub struct TableOfContents {
    /// Ordered list of entries extracted from TOC-style paragraphs.
    pub entries: Vec<TocEntry>,
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

    /// Extract all Tables of Contents from this body.
    ///
    /// Scans the body for both SDT-wrapped and field-based TOCs.
    /// Each group of contiguous TOC-style paragraphs (or paragraphs inside
    /// an SDT that contains TOC entries) is returned as a separate
    /// [`TableOfContents`].
    ///
    /// TOC paragraphs use styles "TOC 1"–"TOC 9" or "toc1"–"toc9"
    /// (ECMA-376 §17.12.1).  Requires the `wml-styling` feature to
    /// detect paragraph styles; without it this always returns an empty vec.
    #[cfg(feature = "wml-styling")]
    fn table_of_contents(&self) -> Vec<TableOfContents>;

    /// Extract all form fields from this body (ECMA-376 §17.5.2).
    ///
    /// Walks all block content recursively (including table cells) and
    /// collects every SDT that has a recognisable form-field type.
    /// Requires the `wml-settings` feature; without it always returns an
    /// empty vec.
    #[cfg(feature = "wml-settings")]
    fn form_fields(&self) -> Vec<FormField>;
}

impl BodyExt for types::Body {
    fn paragraphs(&self) -> Vec<&types::Paragraph> {
        self.block_content
            .iter()
            .filter_map(|elt| match elt {
                types::BlockContent::P(p) => Some(p.as_ref()),
                _ => None,
            })
            .collect()
    }

    fn tables(&self) -> Vec<&types::Table> {
        self.block_content
            .iter()
            .filter_map(|elt| match elt {
                types::BlockContent::Tbl(t) => Some(t.as_ref()),
                _ => None,
            })
            .collect()
    }

    fn text(&self) -> String {
        let texts: Vec<String> = self
            .block_content
            .iter()
            .filter_map(|elt| match elt {
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

    #[cfg(feature = "wml-styling")]
    fn table_of_contents(&self) -> Vec<TableOfContents> {
        collect_tocs_from_block_content(&self.block_content)
    }

    #[cfg(feature = "wml-settings")]
    fn form_fields(&self) -> Vec<FormField> {
        collect_form_fields_from_block_content(&self.block_content)
    }
}

// =============================================================================
// TOC helpers (private)
// =============================================================================

/// Return the TOC level (1–9) for a paragraph style name, or `None` if not a
/// TOC style.
///
/// Recognises both display names ("TOC 1"–"TOC 9") and style IDs
/// ("toc1"–"toc9"), case-insensitively.
#[cfg(feature = "wml-styling")]
fn toc_style_level(style: &str) -> Option<u8> {
    let s = style.trim();

    // "TOC 1" … "TOC 9"  (display name, space-separated)
    if let Some(rest) = s.strip_prefix("TOC ").or_else(|| s.strip_prefix("toc "))
        && let Ok(n) = rest.trim().parse::<u8>()
        && (1..=9).contains(&n)
    {
        return Some(n);
    }

    // "toc1" … "toc9"  (style ID, no space)
    if let Some(rest) = s
        .strip_prefix("TOC")
        .or_else(|| s.strip_prefix("toc"))
        .filter(|r| r.len() == 1)
        && let Ok(n) = rest.parse::<u8>()
        && (1..=9).contains(&n)
    {
        return Some(n);
    }

    None
}

/// Return the TOC level for a paragraph, or `None` if it is not a TOC entry.
#[cfg(feature = "wml-styling")]
fn paragraph_toc_level(para: &types::Paragraph) -> Option<u8> {
    let style = para.p_pr.as_ref()?.paragraph_style.as_ref()?.value.as_str();
    toc_style_level(style)
}

/// Extract the text of a paragraph, stripping the trailing page number.
///
/// TOC paragraphs typically look like:  "Heading text\t42"
/// The page number is the last tab-separated token if it parses as a number.
/// Returns the trimmed text before the page number and the page number itself.
#[cfg(feature = "wml-styling")]
fn extract_toc_text_and_page(para: &types::Paragraph) -> (String, Option<u32>) {
    // Collect all run text first.
    let mut full = String::new();
    for content in &para.paragraph_content {
        collect_text_from_paragraph_content(content, &mut full);
    }

    // Split on the last tab and try to parse the tail as a page number.
    if let Some(tab_pos) = full.rfind('\t') {
        let tail = full[tab_pos + 1..].trim();
        if let Ok(page) = tail.parse::<u32>() {
            let text = full[..tab_pos].trim().to_string();
            return (text, Some(page));
        }
    }

    (full.trim().to_string(), None)
}

/// Find the first bookmark name embedded in a paragraph's content.
///
/// Hyperlinked TOC entries wrap their content in a `<w:hyperlink>` whose
/// anchor points to a bookmark on the heading.  The bookmark name is stored
/// in a `BookmarkStart` item at the paragraph level.
#[cfg(feature = "wml-styling")]
fn paragraph_bookmark(para: &types::Paragraph) -> Option<String> {
    for content in &para.paragraph_content {
        if let types::ParagraphContent::BookmarkStart(bm) = content {
            let name = bm.name.clone();
            if !name.is_empty() {
                return Some(name);
            }
        }
    }
    None
}

/// Convert a paragraph with a TOC style into a [`TocEntry`].
#[cfg(feature = "wml-styling")]
fn paragraph_to_toc_entry(para: &types::Paragraph, level: u8) -> TocEntry {
    let (text, page) = extract_toc_text_and_page(para);
    let bookmark = paragraph_bookmark(para);
    TocEntry {
        level,
        text,
        page,
        bookmark,
    }
}

/// Collect all TOC entries from a flat slice of [`BlockContent`] items.
///
/// Both SDT-wrapped and bare (field-based) TOC entries are detected.
/// Contiguous runs of TOC paragraphs (or SDTs containing TOC paragraphs) are
/// each returned as a separate [`TableOfContents`].
#[cfg(feature = "wml-styling")]
fn collect_tocs_from_block_content(blocks: &[types::BlockContent]) -> Vec<TableOfContents> {
    let mut result: Vec<TableOfContents> = Vec::new();
    // Accumulator for the current run of bare (non-SDT) TOC paragraphs.
    let mut current_entries: Vec<TocEntry> = Vec::new();

    for block in blocks {
        match block {
            types::BlockContent::P(para) => {
                if let Some(level) = paragraph_toc_level(para) {
                    current_entries.push(paragraph_to_toc_entry(para, level));
                } else {
                    // Non-TOC paragraph: flush any accumulated entries.
                    flush_toc(&mut current_entries, &mut result);
                }
            }
            types::BlockContent::Sdt(sdt) => {
                // Flush bare entries before handling the SDT.
                flush_toc(&mut current_entries, &mut result);

                // Extract TOC entries from the SDT content.
                let sdt_entries = collect_toc_entries_from_sdt(sdt);
                if !sdt_entries.is_empty() {
                    result.push(TableOfContents {
                        entries: sdt_entries,
                    });
                }
            }
            _ => {
                // Any other block (table, custom XML, …) ends a bare TOC run.
                flush_toc(&mut current_entries, &mut result);
            }
        }
    }

    // Flush any trailing bare entries.
    flush_toc(&mut current_entries, &mut result);

    result
}

/// Flush accumulated TOC entries into the result list.
#[cfg(feature = "wml-styling")]
fn flush_toc(entries: &mut Vec<TocEntry>, result: &mut Vec<TableOfContents>) {
    if !entries.is_empty() {
        result.push(TableOfContents {
            entries: std::mem::take(entries),
        });
    }
}

/// Collect all TOC entries from the content of an SDT block.
///
/// The `sdt_content` field holds [`BlockContentChoice`] items.  We walk those,
/// extracting paragraphs with TOC styles.
#[cfg(feature = "wml-styling")]
fn collect_toc_entries_from_sdt(sdt: &types::CTSdtBlock) -> Vec<TocEntry> {
    let content = match &sdt.sdt_content {
        Some(c) => c,
        None => return Vec::new(),
    };

    content
        .block_content
        .iter()
        .filter_map(|bc| match bc {
            types::BlockContentChoice::P(para) => {
                paragraph_toc_level(para).map(|lvl| paragraph_to_toc_entry(para, lvl))
            }
            _ => None,
        })
        .collect()
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

    /// Get paragraph alignment (justification). ECMA-376 §17.3.1.13.
    #[cfg(feature = "wml-styling")]
    fn alignment(&self) -> Option<types::STJc>;

    /// Get left indent in twips. ECMA-376 §17.3.1.12.
    ///
    /// Prefers `w:start` (OOXML) and falls back to `w:left` (compatibility).
    #[cfg(feature = "wml-styling")]
    fn indent_left(&self) -> Option<i64>;

    /// Get right indent in twips.
    ///
    /// Prefers `w:end` (OOXML) and falls back to `w:right` (compatibility).
    #[cfg(feature = "wml-styling")]
    fn indent_right(&self) -> Option<i64>;

    /// Get first-line additional indent in twips (positive = first line further right).
    ///
    /// Mutually exclusive with [`indent_hanging`].
    #[cfg(feature = "wml-styling")]
    fn indent_first_line(&self) -> Option<i64>;

    /// Get hanging indent in twips (positive = first line is that many twips to the *left* of the rest).
    ///
    /// Mutually exclusive with [`indent_first_line`].
    #[cfg(feature = "wml-styling")]
    fn indent_hanging(&self) -> Option<i64>;

    /// Get space before paragraph in twips. ECMA-376 §17.3.1.33.
    #[cfg(feature = "wml-styling")]
    fn space_before(&self) -> Option<i64>;

    /// Get space after paragraph in twips.
    #[cfg(feature = "wml-styling")]
    fn space_after(&self) -> Option<i64>;

    /// Get line spacing value in twips (240 = single, 360 = 1.5×, 480 = double for `auto` rule).
    #[cfg(feature = "wml-styling")]
    fn line_spacing(&self) -> Option<i64>;

    /// Get the line spacing rule, if set.
    #[cfg(feature = "wml-styling")]
    fn line_spacing_rule(&self) -> Option<types::STLineSpacingRule>;

    /// Get numbering (list) properties as `(num_id, ilvl)`. ECMA-376 §17.9.
    ///
    /// Returns `None` if this paragraph is not part of a list.
    #[cfg(feature = "wml-numbering")]
    fn numbering(&self) -> Option<(i64, i64)>;
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
            .filter_map(|c| match c {
                types::ParagraphContent::Hyperlink(h) => Some(h.as_ref()),
                _ => None,
            })
            .collect()
    }

    #[cfg(feature = "wml-styling")]
    fn properties(&self) -> Option<&types::ParagraphProperties> {
        self.p_pr.as_deref()
    }

    #[cfg(feature = "wml-styling")]
    fn alignment(&self) -> Option<types::STJc> {
        self.p_pr
            .as_deref()?
            .justification
            .as_deref()
            .map(|j| j.value)
    }

    #[cfg(feature = "wml-styling")]
    fn indent_left(&self) -> Option<i64> {
        let ind = self.p_pr.as_deref()?.indentation.as_deref()?;
        ind.start
            .as_deref()
            .or(ind.left.as_deref())
            .and_then(parse_twips)
    }

    #[cfg(feature = "wml-styling")]
    fn indent_right(&self) -> Option<i64> {
        let ind = self.p_pr.as_deref()?.indentation.as_deref()?;
        ind.end
            .as_deref()
            .or(ind.right.as_deref())
            .and_then(parse_twips)
    }

    #[cfg(feature = "wml-styling")]
    fn indent_first_line(&self) -> Option<i64> {
        let ind = self.p_pr.as_deref()?.indentation.as_deref()?;
        ind.first_line.as_deref().and_then(parse_twips)
    }

    #[cfg(feature = "wml-styling")]
    fn indent_hanging(&self) -> Option<i64> {
        let ind = self.p_pr.as_deref()?.indentation.as_deref()?;
        ind.hanging.as_deref().and_then(parse_twips)
    }

    #[cfg(feature = "wml-styling")]
    fn space_before(&self) -> Option<i64> {
        let spacing = self.p_pr.as_deref()?.spacing.as_deref()?;
        spacing.before.as_deref().and_then(parse_twips)
    }

    #[cfg(feature = "wml-styling")]
    fn space_after(&self) -> Option<i64> {
        let spacing = self.p_pr.as_deref()?.spacing.as_deref()?;
        spacing.after.as_deref().and_then(parse_twips)
    }

    #[cfg(feature = "wml-styling")]
    fn line_spacing(&self) -> Option<i64> {
        let spacing = self.p_pr.as_deref()?.spacing.as_deref()?;
        spacing.line.as_deref().and_then(parse_twips)
    }

    #[cfg(feature = "wml-styling")]
    fn line_spacing_rule(&self) -> Option<types::STLineSpacingRule> {
        self.p_pr.as_deref()?.spacing.as_deref()?.line_rule
    }

    #[cfg(feature = "wml-numbering")]
    fn numbering(&self) -> Option<(i64, i64)> {
        let num_pr = self.p_pr.as_deref()?.num_pr.as_deref()?;
        let num_id = num_pr.num_id.as_deref()?.value;
        let ilvl = num_pr.ilvl.as_deref()?.value;
        Some((num_id, ilvl))
    }
}

/// Collect runs from paragraph content, including nested runs in hyperlinks and simple fields.
fn collect_runs_from_paragraph_content(content: &[types::ParagraphContent]) -> Vec<&types::Run> {
    let mut runs = Vec::new();
    for item in content {
        match item {
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
            match item {
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
                item,
                types::RunContent::Br(br) if br.r#type == Some(types::STBrType::Page)
            )
        })
    }

    #[cfg(feature = "wml-drawings")]
    fn drawings(&self) -> Vec<&types::CTDrawing> {
        self.run_content
            .iter()
            .filter_map(|item| match item {
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
            .any(|item| matches!(item, types::RunContent::Drawing(_)))
    }

    fn footnote_ref(&self) -> Option<&types::FootnoteEndnoteRef> {
        self.run_content.iter().find_map(|item| match item {
            types::RunContent::FootnoteReference(r) => Some(r.as_ref()),
            _ => None,
        })
    }

    fn endnote_ref(&self) -> Option<&types::FootnoteEndnoteRef> {
        self.run_content.iter().find_map(|item| match item {
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
// DrawingChartExt — extract chart relationship IDs from CTDrawing
// =============================================================================

/// Extension methods for `CTDrawing` — extract chart relationship IDs from raw XML.
///
/// Charts in DOCX appear in `<wp:inline>` or `<wp:anchor>` elements inside
/// `<a:graphic>` → `<a:graphicData>` → `<c:chart r:id="rId..."/>`.
///
/// ECMA-376 Part 1, §20.4.2.8 (inline), §20.4.2.3 (anchor), §21.2.2.27 (chart).
#[cfg(all(feature = "wml-drawings", feature = "extra-children"))]
pub trait DrawingChartExt {
    /// Get relationship IDs for charts in inline drawings (`<wp:inline>`).
    fn inline_chart_rel_ids(&self) -> Vec<&str>;

    /// Get relationship IDs for charts in anchored drawings (`<wp:anchor>`).
    fn anchored_chart_rel_ids(&self) -> Vec<&str>;

    /// Get all chart relationship IDs (inline + anchored).
    fn all_chart_rel_ids(&self) -> Vec<&str>;
}

#[cfg(all(feature = "wml-drawings", feature = "extra-children"))]
impl DrawingChartExt for types::CTDrawing {
    fn inline_chart_rel_ids(&self) -> Vec<&str> {
        let mut ids = Vec::new();
        for child in &self.extra_children {
            if let ooxml_xml::RawXmlNode::Element(elem) = &child.node
                && local_name_of(&elem.name) == "inline"
            {
                collect_chart_rel_ids(elem, &mut ids);
            }
        }
        ids
    }

    fn anchored_chart_rel_ids(&self) -> Vec<&str> {
        let mut ids = Vec::new();
        for child in &self.extra_children {
            if let ooxml_xml::RawXmlNode::Element(elem) = &child.node
                && local_name_of(&elem.name) == "anchor"
            {
                collect_chart_rel_ids(elem, &mut ids);
            }
        }
        ids
    }

    fn all_chart_rel_ids(&self) -> Vec<&str> {
        let mut ids = self.inline_chart_rel_ids();
        ids.extend(self.anchored_chart_rel_ids());
        ids
    }
}

/// Recursively walk a raw XML element tree and collect `r:id` attribute values
/// from `<c:chart>` elements.
#[cfg(all(feature = "wml-drawings", feature = "extra-children"))]
fn collect_chart_rel_ids<'a>(elem: &'a ooxml_xml::RawXmlElement, ids: &mut Vec<&'a str>) {
    if local_name_of(&elem.name) == "chart" {
        for (attr_name, attr_val) in &elem.attributes {
            if attr_name == "r:id" || local_name_of(attr_name) == "id" {
                ids.push(attr_val.as_str());
            }
        }
    }
    for child in &elem.children {
        if let ooxml_xml::RawXmlNode::Element(child_elem) = child {
            collect_chart_rel_ids(child_elem, ids);
        }
    }
}

// =============================================================================
// TextBoxExt (DrawingML — modern text boxes)
// =============================================================================

/// Extension methods for `CTDrawing` — extract text from DrawingML text boxes.
///
/// Modern DOCX text boxes live inside `<wp:anchor>` elements within a `<w:drawing>`.
/// The content path is:
/// `<w:drawing>` → `<wp:anchor>` → `<a:graphic>` → `<a:graphicData>` →
/// `<wps:wsp>` → `<wps:txbx>` → `<w:txbxContent>` → paragraphs
///
/// Since `CTDrawing` captures all children as raw XML (`extra_children`), this trait
/// walks the tree recursively to find `w:txbxContent` elements and parses them.
///
/// ECMA-376 Part 1, §20.4.2.3 (anchor) and §20.1.2.2.19 (graphic).
#[cfg(all(feature = "wml-drawings", feature = "extra-children"))]
pub trait DrawingTextBoxExt {
    /// Extract the plain text of all text boxes in this drawing.
    ///
    /// Returns one `String` per text box found (anchored or inline).
    /// Each string contains the text of all paragraphs in that text box,
    /// joined with newlines.
    fn text_box_texts(&self) -> Vec<String>;
}

#[cfg(all(feature = "wml-drawings", feature = "extra-children"))]
impl DrawingTextBoxExt for types::CTDrawing {
    fn text_box_texts(&self) -> Vec<String> {
        let mut results = Vec::new();
        for child in &self.extra_children {
            if let ooxml_xml::RawXmlNode::Element(elem) = &child.node {
                collect_txbx_texts_from_raw(elem, &mut results);
            }
        }
        results
    }
}

/// Recursively walk a raw XML element tree and collect text from every
/// `w:txbxContent` element found anywhere in the subtree.
#[cfg(all(feature = "wml-drawings", feature = "extra-children"))]
fn collect_txbx_texts_from_raw(elem: &ooxml_xml::RawXmlElement, out: &mut Vec<String>) {
    if local_name_of(&elem.name) == "txbxContent" {
        // Found a text box content element — parse it and extract text.
        match elem.parse_as::<types::CTTxbxContent>() {
            Ok(content) => {
                let text = txbx_content_text(&content);
                out.push(text);
            }
            Err(_) => {
                // Parsing failed; skip this element silently.
            }
        }
        // Don't recurse into txbxContent children — we already parsed the whole subtree.
        return;
    }

    for child in &elem.children {
        if let ooxml_xml::RawXmlNode::Element(child_elem) = child {
            collect_txbx_texts_from_raw(child_elem, out);
        }
    }
}

/// Extract plain text from a `CTTxbxContent` by walking its block content.
#[cfg(all(feature = "wml-drawings", feature = "extra-children"))]
fn txbx_content_text(content: &types::CTTxbxContent) -> String {
    use crate::ext::{ParagraphExt, TableExt};
    let parts: Vec<String> = content
        .block_content
        .iter()
        .filter_map(|bc| match bc {
            types::BlockContent::P(p) => Some(p.text()),
            types::BlockContent::Tbl(t) => Some(t.text()),
            _ => None,
        })
        .collect();
    parts.join("\n")
}

// =============================================================================
// PictExt (VML — legacy text boxes)
// =============================================================================

/// Extension methods for `CTPicture` — extract text from VML text boxes.
///
/// Legacy DOCX text boxes (VML) appear as:
/// `<w:pict>` → `<v:shape>` → `<v:textbox>` → `<w:txbxContent>` → paragraphs
///
/// Since `CTPicture` captures all children as raw XML (`extra_children`), this
/// trait walks the tree to find `w:txbxContent` and parses it.
///
/// ECMA-376 Part 1, §17.3.3.21 (pict).
#[cfg(all(feature = "wml-drawings", feature = "extra-children"))]
pub trait PictExt {
    /// Extract the plain text of the first text box inside this picture element.
    ///
    /// VML picture elements typically contain at most one text box.
    /// Returns `None` if no text box content is found.
    fn text_box_text(&self) -> Option<String>;

    /// Extract the plain text of all text boxes inside this picture element.
    fn text_box_texts(&self) -> Vec<String>;
}

#[cfg(all(feature = "wml-drawings", feature = "extra-children"))]
impl PictExt for types::CTPicture {
    fn text_box_text(&self) -> Option<String> {
        self.text_box_texts().into_iter().next()
    }

    fn text_box_texts(&self) -> Vec<String> {
        let mut results = Vec::new();
        for child in &self.extra_children {
            if let ooxml_xml::RawXmlNode::Element(elem) = &child.node {
                collect_txbx_texts_from_raw(elem, &mut results);
            }
        }
        results
    }
}

// =============================================================================
// MathExt (OMML — Office Math Markup Language)
// =============================================================================

/// A math expression extracted from an `<m:oMath>` element (ECMA-376 Part 1 §22.1).
///
#[cfg(feature = "extra-children")]
#[derive(Debug, Clone)]
pub struct MathExpression {
    /// `true` for display (block) math (`<m:oMathPara>`), `false` for inline (`<m:oMath>`).
    pub is_display: bool,
    /// The parsed OMML math zone containing the structured math content.
    #[cfg(feature = "wml-math")]
    pub zone: ooxml_omml::MathZone,
}

#[cfg(all(feature = "extra-children", feature = "wml-math"))]
impl MathExpression {
    /// Extract plain text representation of the math content.
    pub fn text(&self) -> String {
        self.zone.text()
    }
}

/// Extension methods for types that may contain OMML math (ECMA-376 Part 1 §22.1).
///
/// Math in DOCX uses `<m:oMath>` (inline) or `<m:oMathPara>` (display/block)
/// elements from the namespace
/// `http://schemas.openxmlformats.org/officeDocument/2006/math`.
/// Because the OMML namespace is separate from the WordprocessingML namespace,
/// the generated parser stores these elements in `extra_children` on
/// `Paragraph`.  This trait walks those raw nodes to detect and extract math.
#[cfg(feature = "extra-children")]
pub trait MathExt {
    /// Return all math expressions contained in this element.
    fn math_expressions(&self) -> Vec<MathExpression>;

    /// Return `true` if this element contains at least one math expression.
    fn has_math(&self) -> bool {
        !self.math_expressions().is_empty()
    }
}

#[cfg(feature = "extra-children")]
impl MathExt for types::Paragraph {
    fn math_expressions(&self) -> Vec<MathExpression> {
        let mut out = Vec::new();
        for child in &self.extra_children {
            if let ooxml_xml::RawXmlNode::Element(elem) = &child.node {
                collect_math_from_raw(elem, &mut out);
            }
        }
        out
    }
}

#[cfg(feature = "extra-children")]
impl MathExt for types::Body {
    fn math_expressions(&self) -> Vec<MathExpression> {
        let mut out = Vec::new();
        for item in &self.block_content {
            if let types::BlockContent::P(p) = item {
                out.extend(p.math_expressions());
            }
        }
        out
    }
}

/// Walk a raw XML element and, when an `<m:oMath>` or `<m:oMathPara>` root is
/// found, collect a `MathExpression` and append it to `out`.
///
/// The function does **not** recurse into `<m:oMath>` children once the root
/// is identified — instead it hands the entire subtree to
/// [`collect_math_text`] to gather text leaves.
#[cfg(all(feature = "extra-children", feature = "wml-math"))]
fn parse_math_zone_from_element(elem: &ooxml_xml::RawXmlElement) -> ooxml_omml::MathZone {
    elem.parse_as::<ooxml_omml::MathZone>().unwrap_or_default()
}

#[cfg(feature = "extra-children")]
fn collect_math_from_raw(elem: &ooxml_xml::RawXmlElement, out: &mut Vec<MathExpression>) {
    let local = math_local_name(&elem.name);
    match local {
        "oMathPara" => {
            out.push(MathExpression {
                is_display: true,
                #[cfg(feature = "wml-math")]
                zone: {
                    // Display math: find the inner oMath child and parse it.
                    elem.children
                        .iter()
                        .filter_map(|c| match c {
                            ooxml_xml::RawXmlNode::Element(e)
                                if math_local_name(&e.name) == "oMath" =>
                            {
                                Some(parse_math_zone_from_element(e))
                            }
                            _ => None,
                        })
                        .next()
                        .unwrap_or_default()
                },
            });
        }
        "oMath" => {
            out.push(MathExpression {
                is_display: false,
                #[cfg(feature = "wml-math")]
                zone: parse_math_zone_from_element(elem),
            });
        }
        _ => {
            // Recurse into unrecognised wrapper elements.
            for child in &elem.children {
                if let ooxml_xml::RawXmlNode::Element(child_elem) = child {
                    collect_math_from_raw(child_elem, out);
                }
            }
        }
    }
}

/// Return the local name of a possibly-namespaced element name.
///
/// Handles both `m:oMath` (→ `"oMath"`) and bare `oMath`.  Strips any
/// namespace prefix; does **not** validate that the prefix is actually `m:`.
#[cfg(feature = "extra-children")]
#[inline]
fn math_local_name(name: &str) -> &str {
    name.rsplit(':').next().unwrap_or(name)
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

    /// Check if outline text effect is enabled. ECMA-376 §17.3.2.23.
    fn is_outline(&self) -> bool;

    /// Check if shadow text effect is enabled. ECMA-376 §17.3.2.31.
    fn is_shadow(&self) -> bool;

    /// Check if emboss text effect is enabled. ECMA-376 §17.3.2.13.
    fn is_emboss(&self) -> bool;

    /// Check if imprint (engrave) text effect is enabled. ECMA-376 §17.3.2.18.
    fn is_imprint(&self) -> bool;

    /// Check if spell/grammar check is suppressed for this run. ECMA-376 §17.3.2.22.
    fn is_no_proof(&self) -> bool;

    /// Check if content snaps to document grid. ECMA-376 §17.3.2.34.
    fn is_snap_to_grid(&self) -> bool;

    /// Check if text is hidden when rendered for the web view. ECMA-376 §17.3.2.44.
    fn is_web_hidden(&self) -> bool;

    /// Get character spacing adjustment in twips (positive = expand, negative = condense). ECMA-376 §17.3.2.35.
    fn character_spacing(&self) -> Option<i64>;

    /// Get text scale as a percentage (100 = normal width). ECMA-376 §17.3.2.43.
    fn text_scale_percent(&self) -> Option<u32>;

    /// Get kerning threshold in half-points. ECMA-376 §17.3.2.19.
    fn kerning(&self) -> Option<u32>;

    /// Get baseline shift in half-points (positive = raise, negative = lower). ECMA-376 §17.3.2.28.
    fn baseline_shift(&self) -> Option<i64>;

    /// Get language tag for this run. ECMA-376 §17.3.2.20.
    fn language(&self) -> Option<&types::LanguageElement>;
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

    fn is_outline(&self) -> bool {
        is_on(&self.outline)
    }

    fn is_shadow(&self) -> bool {
        is_on(&self.shadow)
    }

    fn is_emboss(&self) -> bool {
        is_on(&self.emboss)
    }

    fn is_imprint(&self) -> bool {
        is_on(&self.imprint)
    }

    fn is_no_proof(&self) -> bool {
        is_on(&self.no_proof)
    }

    fn is_snap_to_grid(&self) -> bool {
        is_on(&self.snap_to_grid)
    }

    fn is_web_hidden(&self) -> bool {
        is_on(&self.web_hidden)
    }

    fn character_spacing(&self) -> Option<i64> {
        self.spacing
            .as_ref()
            .and_then(|s| s.value.parse::<i64>().ok())
    }

    fn text_scale_percent(&self) -> Option<u32> {
        self.width.as_ref()?.value.as_deref()?.parse::<u32>().ok()
    }

    fn kerning(&self) -> Option<u32> {
        self.kern.as_ref().and_then(|k| parse_half_points(&k.value))
    }

    fn baseline_shift(&self) -> Option<i64> {
        self.position
            .as_ref()
            .and_then(|p| p.value.parse::<i64>().ok())
    }

    fn language(&self) -> Option<&types::LanguageElement> {
        self.lang.as_deref()
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
    fn rel_id(&self) -> Option<&str>;

    /// Check if this is an external hyperlink (has a relationship ID).
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
        #[cfg(feature = "wml-hyperlinks")]
        {
            self.anchor.as_deref()
        }
        #[cfg(not(feature = "wml-hyperlinks"))]
        {
            None
        }
    }

    fn rel_id(&self) -> Option<&str> {
        #[cfg(feature = "wml-hyperlinks")]
        {
            self.id.as_deref()
        }
        #[cfg(not(feature = "wml-hyperlinks"))]
        {
            None
        }
    }

    fn is_external(&self) -> bool {
        #[cfg(feature = "wml-hyperlinks")]
        {
            self.id.is_some()
        }
        #[cfg(not(feature = "wml-hyperlinks"))]
        {
            false
        }
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
            .filter_map(|c| match c {
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
            .filter_map(|c| match c {
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
            .filter_map(|elt| match elt {
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
            .filter_map(|r| match r {
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
            .filter_map(|r| match r {
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
                styles.insert(id.clone(), style.clone());
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

    /// Resolve underline through direct → style chain → defaults.
    fn resolved_is_underline(&self, ctx: &StyleContext) -> bool;

    /// Resolve single strikethrough through direct → style chain → defaults.
    fn resolved_is_strikethrough(&self, ctx: &StyleContext) -> bool;

    /// Resolve double strikethrough through direct → style chain → defaults.
    fn resolved_is_double_strikethrough(&self, ctx: &StyleContext) -> bool;

    /// Resolve all-caps through direct → style chain → defaults.
    fn resolved_is_all_caps(&self, ctx: &StyleContext) -> bool;

    /// Resolve small-caps through direct → style chain → defaults.
    fn resolved_is_small_caps(&self, ctx: &StyleContext) -> bool;

    /// Resolve hidden (`<w:vanish>`) through direct → style chain → defaults.
    fn resolved_is_hidden(&self, ctx: &StyleContext) -> bool;

    /// Resolve highlight color through direct → style chain → defaults.
    fn resolved_highlight_color(&self, ctx: &StyleContext) -> Option<types::STHighlightColor>;

    /// Resolve vertical alignment through direct → style chain → defaults.
    fn resolved_vertical_alignment(&self, ctx: &StyleContext) -> Option<types::STVerticalAlignRun>;
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

    fn resolved_is_underline(&self, ctx: &StyleContext) -> bool {
        resolve_option(&self.r_pr, ctx, |rpr| {
            rpr.underline
                .as_ref()
                .map(|u| !matches!(u.value, Some(types::STUnderline::None)))
        })
        .unwrap_or(false)
    }

    fn resolved_is_strikethrough(&self, ctx: &StyleContext) -> bool {
        resolve_toggle(&self.r_pr, ctx, |rpr| &rpr.strikethrough)
    }

    fn resolved_is_double_strikethrough(&self, ctx: &StyleContext) -> bool {
        resolve_toggle(&self.r_pr, ctx, |rpr| &rpr.dstrike)
    }

    fn resolved_is_all_caps(&self, ctx: &StyleContext) -> bool {
        resolve_toggle(&self.r_pr, ctx, |rpr| &rpr.caps)
    }

    fn resolved_is_small_caps(&self, ctx: &StyleContext) -> bool {
        resolve_toggle(&self.r_pr, ctx, |rpr| &rpr.small_caps)
    }

    fn resolved_is_hidden(&self, ctx: &StyleContext) -> bool {
        resolve_toggle(&self.r_pr, ctx, |rpr| &rpr.vanish)
    }

    fn resolved_highlight_color(&self, ctx: &StyleContext) -> Option<types::STHighlightColor> {
        resolve_option(&self.r_pr, ctx, |rpr| {
            rpr.highlight.as_ref().map(|h| h.value)
        })
    }

    fn resolved_vertical_alignment(&self, ctx: &StyleContext) -> Option<types::STVerticalAlignRun> {
        resolve_option(&self.r_pr, ctx, |rpr| {
            rpr.vert_align.as_ref().map(|va| va.value)
        })
    }
}

/// Resolve a toggle property through the style chain.
#[cfg(feature = "wml-styling")]
fn resolve_toggle(
    direct_rpr: &Option<Box<types::RunProperties>>,
    ctx: &StyleContext,
    accessor: impl Fn(&types::RunProperties) -> &Option<Box<types::OnOffElement>>,
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

/// Parse a chart part from XML bytes using the `ooxml_dml` generated `FromXml` parser.
///
/// This is used by `Document::get_chart()` to parse `word/charts/chartN.xml` parts.
/// Requires the `wml-charts` feature.
///
/// ECMA-376 Part 1, §21.2.2.27 (chartSpace).
#[cfg(feature = "wml-charts")]
pub(crate) fn parse_chart(xml: &[u8]) -> Result<ooxml_dml::types::ChartSpace, ParseError> {
    use ooxml_dml::parsers::FromXml as DmlFromXml;
    let mut reader = Reader::from_reader(Cursor::new(xml));
    let mut buf = Vec::new();

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) => {
                return ooxml_dml::types::ChartSpace::from_xml(&mut reader, &e, false)
                    .map_err(|e| ParseError::UnexpectedElement(e.to_string()));
            }
            Ok(Event::Empty(e)) => {
                return ooxml_dml::types::ChartSpace::from_xml(&mut reader, &e, true)
                    .map_err(|e| ParseError::UnexpectedElement(e.to_string()));
            }
            Ok(Event::Eof) => break,
            Err(e) => return Err(ParseError::Xml(e)),
            _ => {}
        }
        buf.clear();
    }
    Err(ParseError::UnexpectedElement(
        "no chartSpace element found".to_string(),
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

    /// Check if a run is underlined (resolved through style chain).
    pub fn is_underline(&self, run: &types::Run) -> bool {
        run.resolved_is_underline(&self.context)
    }

    /// Check if a run is struck through (resolved through style chain).
    pub fn is_strikethrough(&self, run: &types::Run) -> bool {
        run.resolved_is_strikethrough(&self.context)
    }

    /// Check if a run is double struck through (resolved through style chain).
    pub fn is_double_strikethrough(&self, run: &types::Run) -> bool {
        run.resolved_is_double_strikethrough(&self.context)
    }

    /// Check if a run is all-caps (resolved through style chain).
    pub fn is_all_caps(&self, run: &types::Run) -> bool {
        run.resolved_is_all_caps(&self.context)
    }

    /// Check if a run is small-caps (resolved through style chain).
    pub fn is_small_caps(&self, run: &types::Run) -> bool {
        run.resolved_is_small_caps(&self.context)
    }

    /// Check if a run is hidden (resolved through style chain).
    pub fn is_hidden(&self, run: &types::Run) -> bool {
        run.resolved_is_hidden(&self.context)
    }

    /// Get resolved highlight color.
    pub fn highlight_color(&self, run: &types::Run) -> Option<types::STHighlightColor> {
        run.resolved_highlight_color(&self.context)
    }

    /// Get resolved vertical alignment.
    pub fn vertical_alignment(&self, run: &types::Run) -> Option<types::STVerticalAlignRun> {
        run.resolved_vertical_alignment(&self.context)
    }
}

// =============================================================================
// RevisionExt / BodyRevisionExt
// =============================================================================

/// The type of a tracked change (ECMA-376 §17.13).
#[cfg(feature = "wml-track-changes")]
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum TrackChangeType {
    /// Content that was inserted (`<w:ins>`).
    Insertion,
    /// Content that was deleted (`<w:del>`).
    Deletion,
    /// Content that was moved away from this location (`<w:moveFrom>`).
    MoveFrom,
    /// Content that was moved to this location (`<w:moveTo>`).
    MoveTo,
}

/// A single tracked change in a paragraph (ECMA-376 §17.13.5).
#[cfg(feature = "wml-track-changes")]
#[derive(Debug, Clone)]
pub struct TrackChange {
    /// Revision ID (`w:id` attribute).
    pub id: i64,
    /// Author string (`w:author` attribute).
    pub author: String,
    /// Optional ISO 8601 date/time string (`w:date` attribute).
    pub date: Option<String>,
    /// The kind of change.
    pub change_type: TrackChangeType,
    /// Plain text extracted from the run content inside the change.
    pub text: String,
}

/// Extension methods for reading tracked changes from a paragraph (ECMA-376 §17.13).
#[cfg(feature = "wml-track-changes")]
pub trait RevisionExt {
    /// All tracked changes in this paragraph.
    fn track_changes(&self) -> Vec<TrackChange>;

    /// Text produced by accepting all tracked changes: insertions are kept,
    /// deletions are removed, normal runs are kept.
    fn accepted_text(&self) -> String;

    /// Text produced by rejecting all tracked changes: insertions are removed,
    /// deletions are restored, normal runs are kept.
    fn rejected_text(&self) -> String;

    /// Whether this paragraph contains any tracked changes.
    fn has_track_changes(&self) -> bool;
}

/// Extract plain text from a `CTRunTrackChange`'s `run_content` field.
#[cfg(feature = "wml-track-changes")]
fn text_from_run_track_change(tc: &types::CTRunTrackChange) -> String {
    let mut out = String::new();
    for item in &tc.run_content {
        if let types::RunContentChoice::R(run) = item {
            for rc in &run.run_content {
                match rc {
                    types::RunContent::T(t) => {
                        if let Some(ref s) = t.text {
                            out.push_str(s);
                        }
                    }
                    types::RunContent::Tab(_) => out.push('\t'),
                    types::RunContent::Cr(_) => out.push('\n'),
                    types::RunContent::Br(br) => {
                        if !matches!(
                            br.r#type,
                            Some(types::STBrType::Page) | Some(types::STBrType::Column)
                        ) {
                            out.push('\n');
                        }
                    }
                    // Also capture del-text for deletion change content
                    types::RunContent::DelText(t) => {
                        if let Some(ref s) = t.text {
                            out.push_str(s);
                        }
                    }
                    _ => {}
                }
            }
        }
    }
    out
}

#[cfg(feature = "wml-track-changes")]
impl RevisionExt for types::Paragraph {
    fn track_changes(&self) -> Vec<TrackChange> {
        let mut result = Vec::new();
        for item in &self.paragraph_content {
            let (tc, change_type) = match item {
                types::ParagraphContent::Ins(tc) => (tc.as_ref(), TrackChangeType::Insertion),
                types::ParagraphContent::Del(tc) => (tc.as_ref(), TrackChangeType::Deletion),
                types::ParagraphContent::MoveFrom(tc) => (tc.as_ref(), TrackChangeType::MoveFrom),
                types::ParagraphContent::MoveTo(tc) => (tc.as_ref(), TrackChangeType::MoveTo),
                _ => continue,
            };
            result.push(TrackChange {
                id: tc.id,
                author: tc.author.clone(),
                date: tc.date.clone(),
                change_type,
                text: text_from_run_track_change(tc),
            });
        }
        result
    }

    fn accepted_text(&self) -> String {
        let mut out = String::new();
        for item in &self.paragraph_content {
            match item {
                // Normal runs always included
                types::ParagraphContent::R(r) => {
                    out.push_str(&r.text());
                }
                // Insertions accepted → include text
                types::ParagraphContent::Ins(tc) | types::ParagraphContent::MoveTo(tc) => {
                    out.push_str(&text_from_run_track_change(tc));
                }
                // Deletions rejected → skip
                types::ParagraphContent::Del(_) | types::ParagraphContent::MoveFrom(_) => {}
                // Hyperlinks and simple fields: walk their paragraph_content
                types::ParagraphContent::Hyperlink(h) => {
                    for inner in &h.paragraph_content {
                        collect_text_from_paragraph_content(inner, &mut out);
                    }
                }
                types::ParagraphContent::FldSimple(f) => {
                    for inner in &f.paragraph_content {
                        collect_text_from_paragraph_content(inner, &mut out);
                    }
                }
                _ => {}
            }
        }
        out
    }

    fn rejected_text(&self) -> String {
        let mut out = String::new();
        for item in &self.paragraph_content {
            match item {
                // Normal runs always included
                types::ParagraphContent::R(r) => {
                    out.push_str(&r.text());
                }
                // Insertions rejected → skip
                types::ParagraphContent::Ins(_) | types::ParagraphContent::MoveTo(_) => {}
                // Deletions restored → include text
                types::ParagraphContent::Del(tc) | types::ParagraphContent::MoveFrom(tc) => {
                    out.push_str(&text_from_run_track_change(tc));
                }
                // Hyperlinks and simple fields: walk their paragraph_content
                types::ParagraphContent::Hyperlink(h) => {
                    for inner in &h.paragraph_content {
                        collect_text_from_paragraph_content(inner, &mut out);
                    }
                }
                types::ParagraphContent::FldSimple(f) => {
                    for inner in &f.paragraph_content {
                        collect_text_from_paragraph_content(inner, &mut out);
                    }
                }
                _ => {}
            }
        }
        out
    }

    fn has_track_changes(&self) -> bool {
        self.paragraph_content.iter().any(|item| {
            matches!(
                item,
                types::ParagraphContent::Ins(_)
                    | types::ParagraphContent::Del(_)
                    | types::ParagraphContent::MoveFrom(_)
                    | types::ParagraphContent::MoveTo(_)
            )
        })
    }
}

/// Extension methods for reading tracked changes from a document body (ECMA-376 §17.13).
#[cfg(feature = "wml-track-changes")]
pub trait BodyRevisionExt {
    /// All tracked changes in the document body across all paragraphs.
    fn all_track_changes(&self) -> Vec<TrackChange>;

    /// Full document text with all insertions accepted and deletions removed.
    fn accepted_text(&self) -> String;

    /// Full document text with all insertions rejected and deletions restored.
    fn rejected_text(&self) -> String;
}

/// Collect paragraphs from `BlockContent` items recursively (handles SDTs, custom XML, etc.).
#[cfg(feature = "wml-track-changes")]
fn paragraphs_from_block_content(blocks: &[types::BlockContent]) -> Vec<&types::Paragraph> {
    let mut result = Vec::new();
    for block in blocks {
        match block {
            types::BlockContent::P(p) => result.push(p.as_ref()),
            types::BlockContent::Tbl(t) => {
                for row in &t.rows {
                    if let types::RowContent::Tr(tr) = row {
                        for cell in &tr.cells {
                            if let types::CellContent::Tc(tc) = cell {
                                result.extend(paragraphs_from_block_content(&tc.block_content));
                            }
                        }
                    }
                }
            }
            types::BlockContent::Sdt(sdt) => {
                if let Some(content) = &sdt.sdt_content {
                    for inner in &content.block_content {
                        if let types::BlockContentChoice::P(p) = inner {
                            result.push(p.as_ref());
                        }
                    }
                }
            }
            _ => {}
        }
    }
    result
}

#[cfg(feature = "wml-track-changes")]
impl BodyRevisionExt for types::Body {
    fn all_track_changes(&self) -> Vec<TrackChange> {
        paragraphs_from_block_content(&self.block_content)
            .into_iter()
            .flat_map(|p| p.track_changes())
            .collect()
    }

    fn accepted_text(&self) -> String {
        let paras = paragraphs_from_block_content(&self.block_content);
        paras
            .iter()
            .map(|p| p.accepted_text())
            .collect::<Vec<_>>()
            .join("\n")
    }

    fn rejected_text(&self) -> String {
        let paras = paragraphs_from_block_content(&self.block_content);
        paras
            .iter()
            .map(|p| p.rejected_text())
            .collect::<Vec<_>>()
            .join("\n")
    }
}

// =============================================================================
// Form Fields (ECMA-376 §17.5.2)
// =============================================================================

/// The kind of a Structured Document Tag form control (ECMA-376 §17.5.2).
///
/// Determined by which child element is present inside `<w:sdtPr>`.
#[cfg(feature = "wml-settings")]
#[derive(Debug, Clone, PartialEq)]
pub enum FormFieldType {
    /// Plain-text input control (`<w:text>`).
    ///
    /// `multi_line` is `true` when `w:multiLine` is "1", "true", or "on".
    PlainText { multi_line: bool },
    /// Rich-text area (`<w:richText>`).
    RichText,
    /// Combo box with a fixed list of choices (`<w:comboBox>`).
    ComboBox { choices: Vec<String> },
    /// Drop-down list (`<w:dropDownList>`).
    DropDownList { choices: Vec<String> },
    /// Date picker (`<w:date>`).
    DatePicker { format: Option<String> },
    /// SDT whose type was not recognised by this library.
    Unknown,
}

/// A form field extracted from a Structured Document Tag (ECMA-376 §17.5.2).
///
/// Returned by [`FormFieldExt::form_field`] and collected by
/// [`BodyExt::form_fields`].
#[cfg(feature = "wml-settings")]
#[derive(Debug, Clone)]
pub struct FormField {
    /// Human-readable label (`<w:alias w:val="…"/>`).
    pub alias: Option<String>,
    /// Machine-readable tag (`<w:tag w:val="…"/>`).
    pub tag: Option<String>,
    /// The kind of control inferred from `<w:sdtPr>`.
    pub field_type: FormFieldType,
    /// Current text content of the SDT, extracted from its `sdtContent`.
    pub current_value: String,
}

/// Extract a [`FormField`] from an SDT properties block and its content text.
///
/// The `sdt_pr` argument must already be `Some`; call sites gate on that.
#[cfg(feature = "wml-settings")]
fn sdt_pr_to_form_field(sdt_pr: &types::CTSdtPr, current_value: String) -> FormField {
    let alias = sdt_pr.alias.as_ref().map(|a| a.value.clone());
    let tag = sdt_pr.tag.as_ref().map(|t| t.value.clone());

    let field_type = if let Some(text_elem) = &sdt_pr.text {
        let multi_line = text_elem
            .multi_line
            .as_deref()
            .map(|v| matches!(v, "1" | "true" | "on"))
            .unwrap_or(false);
        FormFieldType::PlainText { multi_line }
    } else if sdt_pr.rich_text.is_some() {
        FormFieldType::RichText
    } else if let Some(cb) = &sdt_pr.combo_box {
        let choices = cb
            .list_item
            .iter()
            .map(|item| {
                item.display_text
                    .clone()
                    .or_else(|| item.value.clone())
                    .unwrap_or_default()
            })
            .collect();
        FormFieldType::ComboBox { choices }
    } else if let Some(dd) = &sdt_pr.drop_down_list {
        let choices = dd
            .list_item
            .iter()
            .map(|item| {
                item.display_text
                    .clone()
                    .or_else(|| item.value.clone())
                    .unwrap_or_default()
            })
            .collect();
        FormFieldType::DropDownList { choices }
    } else if let Some(date) = &sdt_pr.date {
        let format = date.date_format.as_ref().map(|df| df.value.clone());
        FormFieldType::DatePicker { format }
    } else {
        FormFieldType::Unknown
    };

    FormField {
        alias,
        tag,
        field_type,
        current_value,
    }
}

/// Extension trait for extracting a [`FormField`] from an SDT element.
///
/// Implemented for both block-level (`CTSdtBlock`) and inline (`CTSdtRun`)
/// SDT variants. Returns `Some` whenever `sdt_pr` is present; returns `None`
/// if there are no properties at all (which is unusual but spec-legal).
///
/// ECMA-376 §17.5.2.
#[cfg(feature = "wml-settings")]
pub trait FormFieldExt {
    /// Extract a [`FormField`] if this SDT has `<w:sdtPr>` properties.
    fn form_field(&self) -> Option<FormField>;
}

#[cfg(feature = "wml-settings")]
impl FormFieldExt for types::CTSdtBlock {
    fn form_field(&self) -> Option<FormField> {
        let sdt_pr = self.sdt_pr.as_deref()?;
        let value = extract_text_from_block_sdt_content(self.sdt_content.as_deref());
        Some(sdt_pr_to_form_field(sdt_pr, value))
    }
}

#[cfg(feature = "wml-settings")]
impl FormFieldExt for types::CTSdtRun {
    fn form_field(&self) -> Option<FormField> {
        let sdt_pr = self.sdt_pr.as_deref()?;
        let value = extract_text_from_run_sdt_content(self.sdt_content.as_deref());
        Some(sdt_pr_to_form_field(sdt_pr, value))
    }
}

/// Extract plain text from `CTSdtContentBlock` by walking its block content
/// choices (paragraphs and nested tables).
#[cfg(feature = "wml-settings")]
fn extract_text_from_block_sdt_content(content: Option<&types::CTSdtContentBlock>) -> String {
    let content = match content {
        Some(c) => c,
        None => return String::new(),
    };
    let parts: Vec<String> = content
        .block_content
        .iter()
        .filter_map(|bc| match bc {
            types::BlockContentChoice::P(p) => Some(p.text()),
            types::BlockContentChoice::Tbl(t) => Some(t.text()),
            _ => None,
        })
        .collect();
    parts.join("\n")
}

/// Extract plain text from `CTSdtContentRun` by walking its paragraph content.
#[cfg(feature = "wml-settings")]
fn extract_text_from_run_sdt_content(content: Option<&types::CTSdtContentRun>) -> String {
    let content = match content {
        Some(c) => c,
        None => return String::new(),
    };
    let mut out = String::new();
    for item in &content.paragraph_content {
        collect_text_from_paragraph_content(item, &mut out);
    }
    out
}

/// Collect all form fields from a slice of [`types::BlockContent`] items,
/// recursing into tables and SDT block content.
#[cfg(feature = "wml-settings")]
fn collect_form_fields_from_block_content(blocks: &[types::BlockContent]) -> Vec<FormField> {
    let mut result = Vec::new();
    for block in blocks {
        match block {
            types::BlockContent::Sdt(sdt) => {
                if let Some(field) = sdt.form_field() {
                    result.push(field);
                }
                // Also look inside the SDT's block content for nested SDTs.
                if let Some(content) = &sdt.sdt_content {
                    for inner in &content.block_content {
                        collect_form_fields_from_block_content_choice(inner, &mut result);
                    }
                }
            }
            types::BlockContent::Tbl(t) => {
                for row in &t.rows {
                    if let types::RowContent::Tr(tr) = row {
                        for cell_content in &tr.cells {
                            if let types::CellContent::Tc(tc) = cell_content {
                                result.extend(collect_form_fields_from_block_content(
                                    &tc.block_content,
                                ));
                            }
                        }
                    }
                }
            }
            types::BlockContent::P(para) => {
                // Inline (run-level) SDTs inside paragraphs.
                for item in &para.paragraph_content {
                    if let types::ParagraphContent::Sdt(sdt_run) = item
                        && let Some(field) = sdt_run.form_field()
                    {
                        result.push(field);
                    }
                }
            }
            _ => {}
        }
    }
    result
}

/// Helper: collect form fields from a single [`types::BlockContentChoice`].
#[cfg(feature = "wml-settings")]
fn collect_form_fields_from_block_content_choice(
    item: &types::BlockContentChoice,
    result: &mut Vec<FormField>,
) {
    match item {
        types::BlockContentChoice::Sdt(sdt) => {
            if let Some(field) = sdt.form_field() {
                result.push(field);
            }
        }
        types::BlockContentChoice::P(para) => {
            for pc in &para.paragraph_content {
                if let types::ParagraphContent::Sdt(sdt_run) = pc
                    && let Some(field) = sdt_run.form_field()
                {
                    result.push(field);
                }
            }
        }
        _ => {}
    }
}

// =============================================================================
// Tests
// =============================================================================

#[cfg(test)]
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
        let field = Some(Box::new(types::OnOffElement {
            value: None,
            #[cfg(feature = "extra-attrs")]
            extra_attrs: Default::default(),
        }));
        assert!(is_on(&field));
    }

    #[test]
    fn test_is_on_explicit_true() {
        for val in &["1", "true", "on"] {
            let field = Some(Box::new(types::OnOffElement {
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
            let field = Some(Box::new(types::OnOffElement {
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
        let field = Some(Box::new(types::OnOffElement {
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
    fn on_off(val: Option<&str>) -> Option<Box<types::OnOffElement>> {
        Some(Box::new(types::OnOffElement {
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

        rpr.size = Some(Box::new(types::HpsMeasureElement {
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

    fn make_text(s: &str) -> types::RunContent {
        types::RunContent::T(Box::new(types::Text {
            text: Some(s.to_string()),
            #[cfg(feature = "extra-children")]
            extra_children: Default::default(),
        }))
    }

    fn make_tab() -> types::RunContent {
        types::RunContent::Tab(Box::new(types::CTEmpty))
    }

    fn make_br(br_type: Option<types::STBrType>) -> types::RunContent {
        types::RunContent::Br(Box::new(types::CTBr {
            r#type: br_type,
            clear: None,
            #[cfg(feature = "extra-attrs")]
            extra_attrs: Default::default(),
        }))
    }

    fn make_cr() -> types::RunContent {
        types::RunContent::Cr(Box::new(types::CTEmpty))
    }

    fn make_run(content: Vec<types::RunContent>) -> types::Run {
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

    fn make_p_run(text: &str) -> types::ParagraphContent {
        types::ParagraphContent::R(Box::new(make_run(vec![make_text(text)])))
    }

    fn make_paragraph(content: Vec<types::ParagraphContent>) -> types::Paragraph {
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
        let hyperlink = types::ParagraphContent::Hyperlink(Box::new(types::Hyperlink {
            id: None,
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
        }));
        let para = make_paragraph(vec![make_p_run("Click "), hyperlink]);
        assert_eq!(para.runs().len(), 2);
        assert_eq!(para.text(), "Click link text");
        assert_eq!(para.hyperlinks().len(), 1);
        assert_eq!(para.hyperlinks()[0].anchor_str(), Some("bookmark1"));
    }

    #[test]
    fn test_paragraph_with_fld_simple() {
        let fld = types::ParagraphContent::FldSimple(Box::new(types::CTSimpleField {
            instr: "PAGE".to_string(),
            fld_lock: None,
            dirty: None,
            fld_data: None,
            paragraph_content: vec![make_p_run("1")],
            #[cfg(feature = "extra-attrs")]
            extra_attrs: Default::default(),
            #[cfg(feature = "extra-children")]
            extra_children: Default::default(),
        }));
        let para = make_paragraph(vec![make_p_run("Page "), fld]);
        assert_eq!(para.runs().len(), 2);
        assert_eq!(para.text(), "Page 1");
    }

    // -------------------------------------------------------------------------
    // BodyExt tests
    // -------------------------------------------------------------------------

    fn make_body(content: Vec<types::BlockContent>) -> types::Body {
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
        let p1 = types::BlockContent::P(Box::new(make_paragraph(vec![make_p_run("First")])));
        let p2 = types::BlockContent::P(Box::new(make_paragraph(vec![make_p_run("Second")])));
        let body = make_body(vec![p1, p2]);
        assert_eq!(body.paragraphs().len(), 2);
        assert_eq!(body.text(), "First\nSecond");
    }

    #[test]
    fn test_body_tables() {
        let tbl = types::BlockContent::Tbl(Box::new(types::Table {
            range_markup: vec![],
            table_properties: Box::default(),
            tbl_grid: Box::default(),
            rows: vec![],
            #[cfg(feature = "extra-children")]
            extra_children: Default::default(),
        }));
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
            id: None,
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

    fn make_table_cell(text: &str) -> types::CellContent {
        types::CellContent::Tc(Box::new(types::TableCell {
            id: None,
            cell_properties: None,
            block_content: vec![types::BlockContent::P(Box::new(make_paragraph(vec![
                make_p_run(text),
            ])))],
            #[cfg(feature = "extra-attrs")]
            extra_attrs: Default::default(),
            #[cfg(feature = "extra-children")]
            extra_children: Default::default(),
        }))
    }

    fn make_table_row(cells: Vec<types::CellContent>) -> types::RowContent {
        types::RowContent::Tr(Box::new(types::CTRow {
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
        }))
    }

    fn make_table(rows: Vec<types::RowContent>) -> types::Table {
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
                types::BlockContent::P(Box::new(make_paragraph(vec![make_p_run("Line 1")]))),
                types::BlockContent::P(Box::new(make_paragraph(vec![make_p_run("Line 2")]))),
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

    // -------------------------------------------------------------------------
    // DrawingChartExt tests
    // -------------------------------------------------------------------------

    #[test]
    #[cfg(all(feature = "wml-drawings", feature = "extra-children"))]
    fn test_drawing_chart_rel_ids() {
        use super::DrawingChartExt;
        use ooxml_xml::{PositionedNode, RawXmlElement, RawXmlNode};

        // Build: wp:anchor → a:graphic → a:graphicData → c:chart r:id="rId5"
        let chart = RawXmlElement {
            name: "c:chart".to_string(),
            attributes: vec![("r:id".to_string(), "rId5".to_string())],
            children: vec![],
            self_closing: true,
        };
        let graphic_data = RawXmlElement {
            name: "a:graphicData".to_string(),
            attributes: vec![(
                "uri".to_string(),
                "http://schemas.openxmlformats.org/drawingml/2006/chart".to_string(),
            )],
            children: vec![RawXmlNode::Element(chart)],
            self_closing: false,
        };
        let graphic = RawXmlElement {
            name: "a:graphic".to_string(),
            attributes: vec![],
            children: vec![RawXmlNode::Element(graphic_data)],
            self_closing: false,
        };
        let anchor = RawXmlElement {
            name: "wp:anchor".to_string(),
            attributes: vec![],
            children: vec![RawXmlNode::Element(graphic)],
            self_closing: false,
        };

        let drawing = types::CTDrawing {
            extra_children: vec![PositionedNode::new(0, RawXmlNode::Element(anchor))],
        };

        let ids = drawing.all_chart_rel_ids();
        assert_eq!(ids, vec!["rId5"]);

        // anchored_chart_rel_ids should also return it
        assert_eq!(drawing.anchored_chart_rel_ids(), vec!["rId5"]);
        // inline should be empty
        assert!(drawing.inline_chart_rel_ids().is_empty());
    }

    #[test]
    #[cfg(all(feature = "wml-drawings", feature = "extra-children"))]
    fn test_drawing_no_charts() {
        use super::DrawingChartExt;
        use ooxml_xml::{PositionedNode, RawXmlElement, RawXmlNode};

        // Build an anchor with a blip (image), but no chart
        let blip = RawXmlElement {
            name: "a:blip".to_string(),
            attributes: vec![("r:embed".to_string(), "rId1".to_string())],
            children: vec![],
            self_closing: true,
        };
        let anchor = RawXmlElement {
            name: "wp:anchor".to_string(),
            attributes: vec![],
            children: vec![RawXmlNode::Element(blip)],
            self_closing: false,
        };

        let drawing = types::CTDrawing {
            extra_children: vec![PositionedNode::new(0, RawXmlNode::Element(anchor))],
        };

        assert!(drawing.all_chart_rel_ids().is_empty());
    }

    // -------------------------------------------------------------------------
    // DrawingTextBoxExt tests
    // -------------------------------------------------------------------------

    /// Build a minimal `CTDrawing` whose `extra_children` contains a `<wp:anchor>`
    /// that holds a `<w:txbxContent>` with the given paragraph text.
    #[cfg(all(feature = "wml-drawings", feature = "extra-children"))]
    fn make_drawing_with_textbox(text: &str) -> types::CTDrawing {
        use ooxml_xml::{PositionedNode, RawXmlElement, RawXmlNode};

        // Build the element tree bottom-up:
        // <wp:anchor>
        //   <wps:wsp>
        //     <wps:txbx>
        //       <w:txbxContent>
        //         <w:p>
        //           <w:r>
        //             <w:t>text</w:t>
        //           </w:r>
        //         </w:p>
        //       </w:txbxContent>
        //     </wps:txbx>
        //   </wps:wsp>
        // </wp:anchor>

        let t = RawXmlElement {
            name: "w:t".to_string(),
            attributes: vec![],
            children: vec![RawXmlNode::Text(text.to_string())],
            self_closing: false,
        };
        let r = RawXmlElement {
            name: "w:r".to_string(),
            attributes: vec![],
            children: vec![RawXmlNode::Element(t)],
            self_closing: false,
        };
        let p = RawXmlElement {
            name: "w:p".to_string(),
            attributes: vec![],
            children: vec![RawXmlNode::Element(r)],
            self_closing: false,
        };
        let txbx_content = RawXmlElement {
            name: "w:txbxContent".to_string(),
            attributes: vec![],
            children: vec![RawXmlNode::Element(p)],
            self_closing: false,
        };
        let txbx = RawXmlElement {
            name: "wps:txbx".to_string(),
            attributes: vec![],
            children: vec![RawXmlNode::Element(txbx_content)],
            self_closing: false,
        };
        let wsp = RawXmlElement {
            name: "wps:wsp".to_string(),
            attributes: vec![],
            children: vec![RawXmlNode::Element(txbx)],
            self_closing: false,
        };
        let anchor = RawXmlElement {
            name: "wp:anchor".to_string(),
            attributes: vec![],
            children: vec![RawXmlNode::Element(wsp)],
            self_closing: false,
        };

        types::CTDrawing {
            extra_children: vec![PositionedNode::new(0, RawXmlNode::Element(anchor))],
        }
    }

    #[test]
    #[cfg(all(feature = "wml-drawings", feature = "extra-children"))]
    fn test_drawing_text_box_texts_single() {
        use super::DrawingTextBoxExt;
        let drawing = make_drawing_with_textbox("Hello from text box");
        let texts = drawing.text_box_texts();
        assert_eq!(texts.len(), 1);
        assert_eq!(texts[0], "Hello from text box");
    }

    #[test]
    #[cfg(all(feature = "wml-drawings", feature = "extra-children"))]
    fn test_drawing_text_box_texts_empty() {
        use super::DrawingTextBoxExt;
        let drawing = types::CTDrawing {
            extra_children: vec![],
        };
        let texts = drawing.text_box_texts();
        assert!(texts.is_empty());
    }

    #[test]
    #[cfg(all(feature = "wml-drawings", feature = "extra-children"))]
    fn test_drawing_text_box_texts_multiple() {
        use super::DrawingTextBoxExt;
        use ooxml_xml::{PositionedNode, RawXmlElement, RawXmlNode};

        // Build two anchors each with a text box.
        fn make_anchor(text: &str) -> RawXmlElement {
            let t = RawXmlElement {
                name: "w:t".to_string(),
                attributes: vec![],
                children: vec![RawXmlNode::Text(text.to_string())],
                self_closing: false,
            };
            let r = RawXmlElement {
                name: "w:r".to_string(),
                attributes: vec![],
                children: vec![RawXmlNode::Element(t)],
                self_closing: false,
            };
            let p = RawXmlElement {
                name: "w:p".to_string(),
                attributes: vec![],
                children: vec![RawXmlNode::Element(r)],
                self_closing: false,
            };
            let txbx_content = RawXmlElement {
                name: "w:txbxContent".to_string(),
                attributes: vec![],
                children: vec![RawXmlNode::Element(p)],
                self_closing: false,
            };
            RawXmlElement {
                name: "wp:anchor".to_string(),
                attributes: vec![],
                children: vec![RawXmlNode::Element(txbx_content)],
                self_closing: false,
            }
        }

        let drawing = types::CTDrawing {
            extra_children: vec![
                PositionedNode::new(0, RawXmlNode::Element(make_anchor("First box"))),
                PositionedNode::new(1, RawXmlNode::Element(make_anchor("Second box"))),
            ],
        };

        let texts = drawing.text_box_texts();
        assert_eq!(texts.len(), 2);
        assert_eq!(texts[0], "First box");
        assert_eq!(texts[1], "Second box");
    }

    // -------------------------------------------------------------------------
    // PictExt tests (VML text boxes)
    // -------------------------------------------------------------------------

    /// Build a minimal `CTPicture` whose `extra_children` contains a `<v:shape>`
    /// that holds a `<v:textbox>` which holds a `<w:txbxContent>`.
    #[cfg(all(feature = "wml-drawings", feature = "extra-children"))]
    fn make_pict_with_textbox(text: &str) -> types::CTPicture {
        use ooxml_xml::{PositionedNode, RawXmlElement, RawXmlNode};

        let t = RawXmlElement {
            name: "w:t".to_string(),
            attributes: vec![],
            children: vec![RawXmlNode::Text(text.to_string())],
            self_closing: false,
        };
        let r = RawXmlElement {
            name: "w:r".to_string(),
            attributes: vec![],
            children: vec![RawXmlNode::Element(t)],
            self_closing: false,
        };
        let p = RawXmlElement {
            name: "w:p".to_string(),
            attributes: vec![],
            children: vec![RawXmlNode::Element(r)],
            self_closing: false,
        };
        let txbx_content = RawXmlElement {
            name: "w:txbxContent".to_string(),
            attributes: vec![],
            children: vec![RawXmlNode::Element(p)],
            self_closing: false,
        };
        let textbox = RawXmlElement {
            name: "v:textbox".to_string(),
            attributes: vec![],
            children: vec![RawXmlNode::Element(txbx_content)],
            self_closing: false,
        };
        let shape = RawXmlElement {
            name: "v:shape".to_string(),
            attributes: vec![("id".to_string(), "TextBox1".to_string())],
            children: vec![RawXmlNode::Element(textbox)],
            self_closing: false,
        };

        types::CTPicture {
            #[cfg(feature = "wml-drawings")]
            movie: None,
            #[cfg(feature = "wml-drawings")]
            control: None,
            extra_children: vec![PositionedNode::new(0, RawXmlNode::Element(shape))],
        }
    }

    #[test]
    #[cfg(all(feature = "wml-drawings", feature = "extra-children"))]
    fn test_pict_text_box_text() {
        use super::PictExt;
        let pict = make_pict_with_textbox("VML text box content");
        assert_eq!(
            pict.text_box_text(),
            Some("VML text box content".to_string())
        );
    }

    #[test]
    #[cfg(all(feature = "wml-drawings", feature = "extra-children"))]
    fn test_pict_text_box_text_none_when_empty() {
        use super::PictExt;
        let pict = types::CTPicture {
            #[cfg(feature = "wml-drawings")]
            movie: None,
            #[cfg(feature = "wml-drawings")]
            control: None,
            extra_children: vec![],
        };
        assert_eq!(pict.text_box_text(), None);
    }

    #[test]
    #[cfg(all(feature = "wml-drawings", feature = "extra-children"))]
    fn test_pict_text_box_texts() {
        use super::PictExt;
        let pict = make_pict_with_textbox("Hello");
        let texts = pict.text_box_texts();
        assert_eq!(texts.len(), 1);
        assert_eq!(texts[0], "Hello");
    }

    #[test]
    #[cfg(all(feature = "wml-drawings", feature = "extra-children"))]
    fn test_drawing_text_box_via_xml_parse() {
        // Integration test: build the drawing from raw XML, then extract text.
        use super::DrawingTextBoxExt;

        let xml = r#"<?xml version="1.0" encoding="UTF-8"?>
<document xmlns="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
          xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
          xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">
  <body>
    <p>
      <r>
        <drawing>
          <wp:anchor>
            <wps:wsp>
              <wps:txbx>
                <txbxContent>
                  <p><r><t>Anchored box text</t></r></p>
                </txbxContent>
              </wps:txbx>
            </wps:wsp>
          </wp:anchor>
        </drawing>
      </r>
    </p>
  </body>
</document>"#;

        let doc = parse_document(xml.as_bytes()).expect("parse");
        let body = doc.body().expect("body");
        let paras = body.paragraphs();
        assert!(!paras.is_empty());

        let run = &paras[0].runs()[0];
        let drawings = run.drawings();
        assert_eq!(drawings.len(), 1);

        let texts = drawings[0].text_box_texts();
        assert_eq!(texts.len(), 1);
        assert_eq!(texts[0], "Anchored box text");
    }

    // =========================================================================
    // TOC tests
    // =========================================================================

    /// Build a minimal paragraph XML with the given style name and run text.
    #[cfg(feature = "wml-styling")]
    fn toc_para_xml(style: &str, text: &str) -> String {
        format!(
            r#"<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:pPr><w:pStyle w:val="{style}"/></w:pPr>
  <w:r><w:t>{text}</w:t></w:r>
</w:p>"#,
        )
    }

    /// Build a document XML with the given body XML (pre-formatted).
    #[cfg(feature = "wml-styling")]
    fn doc_with_body(body_inner: &str) -> String {
        format!(
            r#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    {body_inner}
  </w:body>
</w:document>"#,
        )
    }

    #[test]
    #[cfg(feature = "wml-styling")]
    fn test_toc_no_entries() {
        // A body with no TOC-style paragraphs returns an empty vec.
        let xml = doc_with_body(
            r#"<w:p><w:pPr><w:pStyle w:val="Normal"/></w:pPr><w:r><w:t>Hello</w:t></w:r></w:p>"#,
        );
        let doc = parse_document(xml.as_bytes()).expect("parse");
        let body = doc.body().expect("body");
        let tocs = body.table_of_contents();
        assert!(tocs.is_empty(), "expected no TOCs, got: {tocs:?}");
    }

    #[test]
    #[cfg(feature = "wml-styling")]
    fn test_toc_levels() {
        // Three consecutive TOC-style paragraphs form a single TOC with correct levels.
        let p1 = toc_para_xml("TOC 1", "Chapter One");
        let p2 = toc_para_xml("TOC 2", "Section 1.1");
        let p3 = toc_para_xml("TOC 3", "Subsection 1.1.1");
        let xml = doc_with_body(&format!("{p1}{p2}{p3}"));
        let doc = parse_document(xml.as_bytes()).expect("parse");
        let body = doc.body().expect("body");
        let tocs = body.table_of_contents();
        assert_eq!(tocs.len(), 1);
        let toc = &tocs[0];
        assert_eq!(toc.entries.len(), 3);
        assert_eq!(toc.entries[0].level, 1);
        assert_eq!(toc.entries[0].text, "Chapter One");
        assert_eq!(toc.entries[1].level, 2);
        assert_eq!(toc.entries[1].text, "Section 1.1");
        assert_eq!(toc.entries[2].level, 3);
        assert_eq!(toc.entries[2].text, "Subsection 1.1.1");
    }

    #[test]
    #[cfg(feature = "wml-styling")]
    fn test_toc_style_id_form() {
        // Style IDs "toc1"/"toc2" (no space) are also recognised.
        let p1 = toc_para_xml("toc1", "First");
        let p2 = toc_para_xml("toc2", "Second");
        let xml = doc_with_body(&format!("{p1}{p2}"));
        let doc = parse_document(xml.as_bytes()).expect("parse");
        let body = doc.body().expect("body");
        let tocs = body.table_of_contents();
        assert_eq!(tocs.len(), 1);
        assert_eq!(tocs[0].entries[0].level, 1);
        assert_eq!(tocs[0].entries[1].level, 2);
    }

    #[test]
    #[cfg(feature = "wml-styling")]
    fn test_toc_entry_from_sdt() {
        // TOC entries inside an SDT block are extracted as a separate TableOfContents.
        let xml = doc_with_body(
            r#"<w:sdt xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:sdtContent>
    <w:p><w:pPr><w:pStyle w:val="TOC 1"/></w:pPr><w:r><w:t>Alpha</w:t></w:r></w:p>
    <w:p><w:pPr><w:pStyle w:val="TOC 2"/></w:pPr><w:r><w:t>Beta</w:t></w:r></w:p>
  </w:sdtContent>
</w:sdt>"#,
        );
        let doc = parse_document(xml.as_bytes()).expect("parse");
        let body = doc.body().expect("body");
        let tocs = body.table_of_contents();
        assert_eq!(tocs.len(), 1, "expected 1 TOC from SDT, got: {tocs:?}");
        assert_eq!(tocs[0].entries.len(), 2);
        assert_eq!(tocs[0].entries[0].level, 1);
        assert_eq!(tocs[0].entries[0].text, "Alpha");
        assert_eq!(tocs[0].entries[1].level, 2);
        assert_eq!(tocs[0].entries[1].text, "Beta");
    }

    #[test]
    #[cfg(feature = "wml-styling")]
    fn test_toc_page_number_extraction() {
        // A page number after a tab stop is extracted as `page`.
        let xml = doc_with_body(
            r#"<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:pPr><w:pStyle w:val="TOC 1"/></w:pPr>
  <w:r><w:t>My Chapter</w:t></w:r>
  <w:r><w:tab/></w:r>
  <w:r><w:t>42</w:t></w:r>
</w:p>"#,
        );
        let doc = parse_document(xml.as_bytes()).expect("parse");
        let body = doc.body().expect("body");
        let tocs = body.table_of_contents();
        assert_eq!(tocs.len(), 1);
        let entry = &tocs[0].entries[0];
        assert_eq!(entry.text, "My Chapter");
        assert_eq!(entry.page, Some(42));
    }

    #[test]
    #[cfg(feature = "wml-styling")]
    fn test_toc_non_toc_para_splits_groups() {
        // A non-TOC paragraph between two TOC runs produces two separate TOCs.
        let p1 = toc_para_xml("TOC 1", "First TOC entry");
        let normal = r#"<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:pPr><w:pStyle w:val="Normal"/></w:pPr>
  <w:r><w:t>Regular text</w:t></w:r>
</w:p>"#;
        let p2 = toc_para_xml("TOC 1", "Second TOC entry");
        let xml = doc_with_body(&format!("{p1}{normal}{p2}"));
        let doc = parse_document(xml.as_bytes()).expect("parse");
        let body = doc.body().expect("body");
        let tocs = body.table_of_contents();
        assert_eq!(tocs.len(), 2, "expected 2 separate TOCs");
        assert_eq!(tocs[0].entries[0].text, "First TOC entry");
        assert_eq!(tocs[1].entries[0].text, "Second TOC entry");
    }

    // -------------------------------------------------------------------------
    // RevisionExt / BodyRevisionExt tests
    // -------------------------------------------------------------------------

    /// Build a paragraph with an `<w:ins>` wrapping a run with `text`, plus an
    /// additional normal run with `suffix`.
    #[cfg(feature = "wml-track-changes")]
    fn make_para_with_ins(ins_text: &str, suffix: &str) -> types::Paragraph {
        use crate::convenience::ins_run;
        let mut para = types::Paragraph::default();
        para.paragraph_content
            .push(ins_run(1, "Alice", Some("2026-01-01T00:00:00Z"), ins_text));
        // Normal run
        let t = types::Text {
            text: Some(suffix.to_string()),
            #[cfg(feature = "extra-children")]
            extra_children: Vec::new(),
        };
        let run = types::Run {
            #[cfg(feature = "wml-track-changes")]
            rsid_r_pr: None,
            #[cfg(feature = "wml-track-changes")]
            rsid_del: None,
            #[cfg(feature = "wml-track-changes")]
            rsid_r: None,
            #[cfg(feature = "wml-styling")]
            r_pr: None,
            run_content: vec![types::RunContent::T(Box::new(t))],
            #[cfg(feature = "extra-attrs")]
            extra_attrs: Default::default(),
            #[cfg(feature = "extra-children")]
            extra_children: Vec::new(),
        };
        para.paragraph_content
            .push(types::ParagraphContent::R(Box::new(run)));
        para
    }

    /// Build a paragraph with a `<w:del>` wrapping a run with `del_text`, plus a
    /// normal run with `suffix`.
    #[cfg(feature = "wml-track-changes")]
    fn make_para_with_del(del_text: &str, suffix: &str) -> types::Paragraph {
        use crate::convenience::del_run;
        let mut para = types::Paragraph::default();
        para.paragraph_content
            .push(del_run(2, "Bob", None, del_text));
        let t = types::Text {
            text: Some(suffix.to_string()),
            #[cfg(feature = "extra-children")]
            extra_children: Vec::new(),
        };
        let run = types::Run {
            #[cfg(feature = "wml-track-changes")]
            rsid_r_pr: None,
            #[cfg(feature = "wml-track-changes")]
            rsid_del: None,
            #[cfg(feature = "wml-track-changes")]
            rsid_r: None,
            #[cfg(feature = "wml-styling")]
            r_pr: None,
            run_content: vec![types::RunContent::T(Box::new(t))],
            #[cfg(feature = "extra-attrs")]
            extra_attrs: Default::default(),
            #[cfg(feature = "extra-children")]
            extra_children: Vec::new(),
        };
        para.paragraph_content
            .push(types::ParagraphContent::R(Box::new(run)));
        para
    }

    #[test]
    #[cfg(feature = "wml-track-changes")]
    fn test_track_changes_accepted_text() {
        use super::RevisionExt;
        // Ins("hello") + Run(" world") → accepted = "hello world"
        let para = make_para_with_ins("hello", " world");
        assert_eq!(para.accepted_text(), "hello world");
    }

    #[test]
    #[cfg(feature = "wml-track-changes")]
    fn test_track_changes_rejected_text() {
        use super::RevisionExt;
        // Del("old") + Run(" word") → rejected = "old word"
        let para = make_para_with_del("old", " word");
        assert_eq!(para.rejected_text(), "old word");
    }

    #[test]
    #[cfg(feature = "wml-track-changes")]
    fn test_track_changes_accepted_text_excludes_deletions() {
        use super::RevisionExt;
        // Del("old") + Run(" word") → accepted = " word" (deletion excluded)
        let para = make_para_with_del("old", " word");
        assert_eq!(para.accepted_text(), " word");
    }

    #[test]
    #[cfg(feature = "wml-track-changes")]
    fn test_track_changes_rejected_text_excludes_insertions() {
        use super::RevisionExt;
        // Ins("hello") + Run(" world") → rejected = " world" (insertion excluded)
        let para = make_para_with_ins("hello", " world");
        assert_eq!(para.rejected_text(), " world");
    }

    #[test]
    #[cfg(feature = "wml-track-changes")]
    fn test_has_track_changes() {
        use super::RevisionExt;
        let para_with = make_para_with_ins("text", "");
        assert!(para_with.has_track_changes());

        // A plain paragraph with no tracked changes
        let plain = types::Paragraph::default();
        assert!(!plain.has_track_changes());
    }

    #[test]
    #[cfg(feature = "wml-track-changes")]
    fn test_track_changes_list() {
        use super::{RevisionExt, TrackChangeType};
        let para = make_para_with_ins("hello", " world");
        let changes = para.track_changes();
        assert_eq!(changes.len(), 1);
        let tc = &changes[0];
        assert_eq!(tc.id, 1);
        assert_eq!(tc.author, "Alice");
        assert_eq!(tc.date.as_deref(), Some("2026-01-01T00:00:00Z"));
        assert_eq!(tc.change_type, TrackChangeType::Insertion);
        assert_eq!(tc.text, "hello");
    }

    #[test]
    #[cfg(feature = "wml-track-changes")]
    fn test_track_changes_deletion_list() {
        use super::{RevisionExt, TrackChangeType};
        let para = make_para_with_del("old", " text");
        let changes = para.track_changes();
        assert_eq!(changes.len(), 1);
        let tc = &changes[0];
        assert_eq!(tc.id, 2);
        assert_eq!(tc.author, "Bob");
        assert_eq!(tc.date, None);
        assert_eq!(tc.change_type, TrackChangeType::Deletion);
        assert_eq!(tc.text, "old");
    }

    #[test]
    #[cfg(feature = "wml-track-changes")]
    fn test_body_revision_ext_all_track_changes() {
        use super::{BodyRevisionExt, TrackChangeType};
        let para1 = make_para_with_ins("inserted", "");
        let para2 = make_para_with_del("deleted", "");

        let body = types::Body {
            block_content: vec![
                types::BlockContent::P(Box::new(para1)),
                types::BlockContent::P(Box::new(para2)),
            ],
            #[cfg(feature = "wml-layout")]
            sect_pr: None,
            #[cfg(feature = "extra-children")]
            extra_children: Vec::new(),
        };
        let all = body.all_track_changes();
        assert_eq!(all.len(), 2);
        assert_eq!(all[0].change_type, TrackChangeType::Insertion);
        assert_eq!(all[0].text, "inserted");
        assert_eq!(all[1].change_type, TrackChangeType::Deletion);
        assert_eq!(all[1].text, "deleted");
    }

    #[test]
    #[cfg(feature = "wml-track-changes")]
    fn test_body_revision_ext_accepted_text() {
        use super::BodyRevisionExt;
        // Para 1: Ins("hello") + Run(" world")  → accepted = "hello world"
        // Para 2: Del("old") + Run(" text")     → accepted = " text"
        // joined with "\n"
        let para1 = make_para_with_ins("hello", " world");
        let para2 = make_para_with_del("old", " text");

        let body = types::Body {
            block_content: vec![
                types::BlockContent::P(Box::new(para1)),
                types::BlockContent::P(Box::new(para2)),
            ],
            #[cfg(feature = "wml-layout")]
            sect_pr: None,
            #[cfg(feature = "extra-children")]
            extra_children: Vec::new(),
        };
        assert_eq!(body.accepted_text(), "hello world\n text");
    }

    // -------------------------------------------------------------------------
    // FormFieldExt / BodyExt::form_fields tests
    // -------------------------------------------------------------------------

    /// Build a minimal `CTSdtPr` with only alias/tag set.
    #[cfg(feature = "wml-settings")]
    fn make_sdt_pr_base(alias: Option<&str>, tag: Option<&str>) -> types::CTSdtPr {
        types::CTSdtPr {
            r_pr: None,
            alias: alias.map(|s| {
                Box::new(types::CTString {
                    value: s.to_string(),
                    #[cfg(feature = "extra-attrs")]
                    extra_attrs: Default::default(),
                })
            }),
            tag: tag.map(|s| {
                Box::new(types::CTString {
                    value: s.to_string(),
                    #[cfg(feature = "extra-attrs")]
                    extra_attrs: Default::default(),
                })
            }),
            id: None,
            lock: None,
            placeholder: None,
            temporary: None,
            showing_plc_hdr: None,
            data_binding: None,
            label: None,
            tab_index: None,
            equation: None,
            combo_box: None,
            date: None,
            doc_part_obj: None,
            doc_part_list: None,
            drop_down_list: None,
            picture: None,
            rich_text: None,
            text: None,
            citation: None,
            group: None,
            bibliography: None,
            #[cfg(feature = "extra-children")]
            extra_children: Default::default(),
        }
    }

    /// Build a `CTSdtRun` containing a run with the given text.
    #[cfg(feature = "wml-settings")]
    fn make_sdt_run(sdt_pr: types::CTSdtPr, value_text: &str) -> types::CTSdtRun {
        let content = types::CTSdtContentRun {
            paragraph_content: vec![types::ParagraphContent::R(Box::new(make_run(vec![
                make_text(value_text),
            ])))],
            #[cfg(feature = "extra-children")]
            extra_children: Default::default(),
        };
        types::CTSdtRun {
            sdt_pr: Some(Box::new(sdt_pr)),
            sdt_end_pr: None,
            sdt_content: Some(Box::new(content)),
            #[cfg(feature = "extra-children")]
            extra_children: Default::default(),
        }
    }

    /// Build a `CTSdtBlock` containing a paragraph with the given text.
    #[cfg(feature = "wml-settings")]
    fn make_sdt_block(sdt_pr: types::CTSdtPr, value_text: &str) -> types::CTSdtBlock {
        let para = make_paragraph(vec![make_p_run(value_text)]);
        let content = types::CTSdtContentBlock {
            block_content: vec![types::BlockContentChoice::P(Box::new(para))],
            #[cfg(feature = "extra-children")]
            extra_children: Default::default(),
        };
        types::CTSdtBlock {
            sdt_pr: Some(Box::new(sdt_pr)),
            sdt_end_pr: None,
            sdt_content: Some(Box::new(content)),
            #[cfg(feature = "extra-children")]
            extra_children: Default::default(),
        }
    }

    #[test]
    #[cfg(feature = "wml-settings")]
    fn test_form_field_plain_text() {
        use super::{FormFieldExt, FormFieldType};
        let mut sdt_pr = make_sdt_pr_base(None, None);
        sdt_pr.text = Some(Box::new(types::CTSdtText {
            multi_line: None,
            #[cfg(feature = "extra-attrs")]
            extra_attrs: Default::default(),
        }));
        let sdt_run = make_sdt_run(sdt_pr, "my value");
        let field = sdt_run.form_field().expect("should have form field");
        assert_eq!(
            field.field_type,
            FormFieldType::PlainText { multi_line: false }
        );
        assert_eq!(field.current_value, "my value");
    }

    #[test]
    #[cfg(feature = "wml-settings")]
    fn test_form_field_plain_text_multiline() {
        use super::{FormFieldExt, FormFieldType};
        let mut sdt_pr = make_sdt_pr_base(None, None);
        sdt_pr.text = Some(Box::new(types::CTSdtText {
            multi_line: Some("1".to_string()),
            #[cfg(feature = "extra-attrs")]
            extra_attrs: Default::default(),
        }));
        let sdt_run = make_sdt_run(sdt_pr, "line1");
        let field = sdt_run.form_field().expect("should have form field");
        assert_eq!(
            field.field_type,
            FormFieldType::PlainText { multi_line: true }
        );
    }

    #[test]
    #[cfg(feature = "wml-settings")]
    fn test_form_field_combo_box() {
        use super::{FormFieldExt, FormFieldType};
        let mut sdt_pr = make_sdt_pr_base(None, None);
        sdt_pr.combo_box = Some(Box::new(types::CTSdtComboBox {
            last_value: Some("Option A".to_string()),
            list_item: vec![
                types::CTSdtListItem {
                    display_text: Some("Option A".to_string()),
                    value: Some("a".to_string()),
                    #[cfg(feature = "extra-attrs")]
                    extra_attrs: Default::default(),
                },
                types::CTSdtListItem {
                    display_text: Some("Option B".to_string()),
                    value: Some("b".to_string()),
                    #[cfg(feature = "extra-attrs")]
                    extra_attrs: Default::default(),
                },
            ],
            #[cfg(feature = "extra-attrs")]
            extra_attrs: Default::default(),
            #[cfg(feature = "extra-children")]
            extra_children: Default::default(),
        }));
        let sdt_block = make_sdt_block(sdt_pr, "Option A");
        let field = sdt_block.form_field().expect("should have form field");
        match &field.field_type {
            FormFieldType::ComboBox { choices } => {
                assert_eq!(
                    choices,
                    &vec!["Option A".to_string(), "Option B".to_string()]
                );
            }
            other => panic!("expected ComboBox, got {other:?}"),
        }
        assert_eq!(field.current_value, "Option A");
    }

    #[test]
    #[cfg(feature = "wml-settings")]
    fn test_form_field_dropdown() {
        use super::{FormFieldExt, FormFieldType};
        let mut sdt_pr = make_sdt_pr_base(None, None);
        sdt_pr.drop_down_list = Some(Box::new(types::CTSdtDropDownList {
            last_value: None,
            list_item: vec![
                types::CTSdtListItem {
                    display_text: Some("Red".to_string()),
                    value: Some("red".to_string()),
                    #[cfg(feature = "extra-attrs")]
                    extra_attrs: Default::default(),
                },
                types::CTSdtListItem {
                    display_text: Some("Blue".to_string()),
                    value: Some("blue".to_string()),
                    #[cfg(feature = "extra-attrs")]
                    extra_attrs: Default::default(),
                },
            ],
            #[cfg(feature = "extra-attrs")]
            extra_attrs: Default::default(),
            #[cfg(feature = "extra-children")]
            extra_children: Default::default(),
        }));
        let sdt_block = make_sdt_block(sdt_pr, "Red");
        let field = sdt_block.form_field().expect("should have form field");
        match &field.field_type {
            FormFieldType::DropDownList { choices } => {
                assert_eq!(choices, &vec!["Red".to_string(), "Blue".to_string()]);
            }
            other => panic!("expected DropDownList, got {other:?}"),
        }
        assert_eq!(field.current_value, "Red");
    }

    #[test]
    #[cfg(feature = "wml-settings")]
    fn test_form_field_alias_and_tag() {
        use super::{FormFieldExt, FormFieldType};
        let mut sdt_pr = make_sdt_pr_base(Some("Full Name"), Some("fullName"));
        sdt_pr.text = Some(Box::new(types::CTSdtText {
            multi_line: None,
            #[cfg(feature = "extra-attrs")]
            extra_attrs: Default::default(),
        }));
        let sdt_run = make_sdt_run(sdt_pr, "Jane Doe");
        let field = sdt_run.form_field().expect("should have form field");
        assert_eq!(field.alias.as_deref(), Some("Full Name"));
        assert_eq!(field.tag.as_deref(), Some("fullName"));
        assert_eq!(
            field.field_type,
            FormFieldType::PlainText { multi_line: false }
        );
        assert_eq!(field.current_value, "Jane Doe");
    }

    #[test]
    #[cfg(feature = "wml-settings")]
    fn test_form_field_rich_text() {
        use super::{FormFieldExt, FormFieldType};
        let mut sdt_pr = make_sdt_pr_base(None, None);
        sdt_pr.rich_text = Some(Box::new(types::CTEmpty));
        let sdt_block = make_sdt_block(sdt_pr, "rich content here");
        let field = sdt_block.form_field().expect("should have form field");
        assert_eq!(field.field_type, FormFieldType::RichText);
        assert_eq!(field.current_value, "rich content here");
    }

    #[test]
    #[cfg(feature = "wml-settings")]
    fn test_form_field_date_picker() {
        use super::{FormFieldExt, FormFieldType};
        let mut sdt_pr = make_sdt_pr_base(None, None);
        sdt_pr.date = Some(Box::new(types::CTSdtDate {
            full_date: Some("2026-02-24T00:00:00Z".to_string()),
            date_format: Some(Box::new(types::CTString {
                value: "yyyy-MM-dd".to_string(),
                #[cfg(feature = "extra-attrs")]
                extra_attrs: Default::default(),
            })),
            lid: None,
            store_mapped_data_as: None,
            calendar: None,
            #[cfg(feature = "extra-attrs")]
            extra_attrs: Default::default(),
            #[cfg(feature = "extra-children")]
            extra_children: Default::default(),
        }));
        let sdt_block = make_sdt_block(sdt_pr, "2026-02-24");
        let field = sdt_block.form_field().expect("should have form field");
        match &field.field_type {
            FormFieldType::DatePicker { format } => {
                assert_eq!(format.as_deref(), Some("yyyy-MM-dd"));
            }
            other => panic!("expected DatePicker, got {other:?}"),
        }
        assert_eq!(field.current_value, "2026-02-24");
    }

    #[test]
    #[cfg(feature = "wml-settings")]
    fn test_form_fields_from_body() {
        use super::{BodyExt, FormFieldType};

        // SDT 1: block-level plain text
        let sdt_pr1 = {
            let mut pr = make_sdt_pr_base(Some("First Name"), Some("firstName"));
            pr.text = Some(Box::new(types::CTSdtText {
                multi_line: None,
                #[cfg(feature = "extra-attrs")]
                extra_attrs: Default::default(),
            }));
            pr
        };
        let block_sdt = make_sdt_block(sdt_pr1, "John");

        // SDT 2: inline (run-level) combo box inside a paragraph
        let sdt_pr2 = {
            let mut pr = make_sdt_pr_base(Some("Color"), None);
            pr.combo_box = Some(Box::new(types::CTSdtComboBox {
                last_value: Some("Red".to_string()),
                list_item: vec![types::CTSdtListItem {
                    display_text: Some("Red".to_string()),
                    value: Some("red".to_string()),
                    #[cfg(feature = "extra-attrs")]
                    extra_attrs: Default::default(),
                }],
                #[cfg(feature = "extra-attrs")]
                extra_attrs: Default::default(),
                #[cfg(feature = "extra-children")]
                extra_children: Default::default(),
            }));
            pr
        };
        let inline_sdt = make_sdt_run(sdt_pr2, "Red");
        let para_with_sdt =
            make_paragraph(vec![types::ParagraphContent::Sdt(Box::new(inline_sdt))]);

        let body = make_body(vec![
            types::BlockContent::Sdt(Box::new(block_sdt)),
            types::BlockContent::P(Box::new(para_with_sdt)),
        ]);

        let fields = body.form_fields();
        assert_eq!(fields.len(), 2);

        assert_eq!(fields[0].alias.as_deref(), Some("First Name"));
        assert_eq!(fields[0].tag.as_deref(), Some("firstName"));
        assert_eq!(
            fields[0].field_type,
            FormFieldType::PlainText { multi_line: false }
        );
        assert_eq!(fields[0].current_value, "John");

        assert_eq!(fields[1].alias.as_deref(), Some("Color"));
        assert!(
            matches!(&fields[1].field_type, FormFieldType::ComboBox { choices } if choices == &["Red"])
        );
        assert_eq!(fields[1].current_value, "Red");
    }

    #[test]
    #[cfg(feature = "wml-settings")]
    fn test_form_field_no_sdt_pr_returns_none() {
        use super::FormFieldExt;
        let sdt_run = types::CTSdtRun {
            sdt_pr: None,
            sdt_end_pr: None,
            sdt_content: None,
            #[cfg(feature = "extra-children")]
            extra_children: Default::default(),
        };
        assert!(sdt_run.form_field().is_none());
    }

    // -------------------------------------------------------------------------
    // MathExt tests
    // -------------------------------------------------------------------------

    /// Build a `Paragraph` whose `extra_children` contains a minimal
    /// `<m:oMath>` element with the supplied math text.
    ///
    /// Structure:
    /// ```xml
    /// <m:oMath>
    ///   <m:r>
    ///     <m:t>text</m:t>
    ///   </m:r>
    /// </m:oMath>
    /// ```
    #[cfg(feature = "extra-children")]
    fn make_paragraph_with_inline_math(math_text: &str) -> types::Paragraph {
        use ooxml_xml::{PositionedNode, RawXmlElement, RawXmlNode};

        let t = RawXmlElement {
            name: "m:t".to_string(),
            attributes: vec![],
            children: vec![RawXmlNode::Text(math_text.to_string())],
            self_closing: false,
        };
        let r = RawXmlElement {
            name: "m:r".to_string(),
            attributes: vec![],
            children: vec![RawXmlNode::Element(t)],
            self_closing: false,
        };
        let o_math = RawXmlElement {
            name: "m:oMath".to_string(),
            attributes: vec![],
            children: vec![RawXmlNode::Element(r)],
            self_closing: false,
        };

        types::Paragraph {
            #[cfg(feature = "wml-track-changes")]
            rsid_r_pr: None,
            #[cfg(feature = "wml-track-changes")]
            rsid_r: None,
            #[cfg(feature = "wml-track-changes")]
            rsid_del: None,
            #[cfg(feature = "wml-track-changes")]
            rsid_p: None,
            #[cfg(feature = "wml-track-changes")]
            rsid_r_default: None,
            #[cfg(feature = "wml-styling")]
            p_pr: None,
            paragraph_content: vec![],
            #[cfg(feature = "extra-attrs")]
            extra_attrs: Default::default(),
            extra_children: vec![PositionedNode::new(0, RawXmlNode::Element(o_math))],
        }
    }

    /// Build a `Paragraph` whose `extra_children` contains a display
    /// `<m:oMathPara>` wrapping a `<m:oMath>`.
    ///
    /// Structure:
    /// ```xml
    /// <m:oMathPara>
    ///   <m:oMath>
    ///     <m:r><m:t>text</m:t></m:r>
    ///   </m:oMath>
    /// </m:oMathPara>
    /// ```
    #[cfg(feature = "extra-children")]
    fn make_paragraph_with_display_math(math_text: &str) -> types::Paragraph {
        use ooxml_xml::{PositionedNode, RawXmlElement, RawXmlNode};

        let t = RawXmlElement {
            name: "m:t".to_string(),
            attributes: vec![],
            children: vec![RawXmlNode::Text(math_text.to_string())],
            self_closing: false,
        };
        let r = RawXmlElement {
            name: "m:r".to_string(),
            attributes: vec![],
            children: vec![RawXmlNode::Element(t)],
            self_closing: false,
        };
        let o_math = RawXmlElement {
            name: "m:oMath".to_string(),
            attributes: vec![],
            children: vec![RawXmlNode::Element(r)],
            self_closing: false,
        };
        let o_math_para = RawXmlElement {
            name: "m:oMathPara".to_string(),
            attributes: vec![],
            children: vec![RawXmlNode::Element(o_math)],
            self_closing: false,
        };

        types::Paragraph {
            #[cfg(feature = "wml-track-changes")]
            rsid_r_pr: None,
            #[cfg(feature = "wml-track-changes")]
            rsid_r: None,
            #[cfg(feature = "wml-track-changes")]
            rsid_del: None,
            #[cfg(feature = "wml-track-changes")]
            rsid_p: None,
            #[cfg(feature = "wml-track-changes")]
            rsid_r_default: None,
            #[cfg(feature = "wml-styling")]
            p_pr: None,
            paragraph_content: vec![],
            #[cfg(feature = "extra-attrs")]
            extra_attrs: Default::default(),
            extra_children: vec![PositionedNode::new(0, RawXmlNode::Element(o_math_para))],
        }
    }

    #[test]
    #[cfg(feature = "extra-children")]
    fn test_math_expression_inline() {
        use super::MathExt;
        let para = make_paragraph_with_inline_math("x+y");
        let exprs = para.math_expressions();
        assert_eq!(exprs.len(), 1);
        assert!(!exprs[0].is_display);
        #[cfg(feature = "wml-math")]
        assert_eq!(exprs[0].text(), "x+y");
    }

    #[test]
    #[cfg(feature = "extra-children")]
    fn test_math_expression_display() {
        use super::MathExt;
        let para = make_paragraph_with_display_math("E=mc²");
        let exprs = para.math_expressions();
        assert_eq!(exprs.len(), 1);
        assert!(exprs[0].is_display);
        #[cfg(feature = "wml-math")]
        assert_eq!(exprs[0].text(), "E=mc²");
    }

    #[test]
    #[cfg(feature = "extra-children")]
    fn test_has_math_true() {
        use super::MathExt;
        let para = make_paragraph_with_inline_math("a²+b²=c²");
        assert!(para.has_math());
    }

    #[test]
    #[cfg(feature = "extra-children")]
    fn test_has_math_false() {
        use super::MathExt;
        // A paragraph with no math in extra_children.
        let para = types::Paragraph {
            #[cfg(feature = "wml-track-changes")]
            rsid_r_pr: None,
            #[cfg(feature = "wml-track-changes")]
            rsid_r: None,
            #[cfg(feature = "wml-track-changes")]
            rsid_del: None,
            #[cfg(feature = "wml-track-changes")]
            rsid_p: None,
            #[cfg(feature = "wml-track-changes")]
            rsid_r_default: None,
            #[cfg(feature = "wml-styling")]
            p_pr: None,
            paragraph_content: vec![],
            #[cfg(feature = "extra-attrs")]
            extra_attrs: Default::default(),
            extra_children: vec![],
        };
        assert!(!para.has_math());
    }

    #[test]
    #[cfg(feature = "extra-children")]
    fn test_body_math_expressions() {
        use super::MathExt;

        let para1 = make_paragraph_with_inline_math("x+y");
        let para2 = make_paragraph_with_display_math("∫f(x)dx");
        // A paragraph without math.
        let para3 = types::Paragraph {
            #[cfg(feature = "wml-track-changes")]
            rsid_r_pr: None,
            #[cfg(feature = "wml-track-changes")]
            rsid_r: None,
            #[cfg(feature = "wml-track-changes")]
            rsid_del: None,
            #[cfg(feature = "wml-track-changes")]
            rsid_p: None,
            #[cfg(feature = "wml-track-changes")]
            rsid_r_default: None,
            #[cfg(feature = "wml-styling")]
            p_pr: None,
            paragraph_content: vec![],
            #[cfg(feature = "extra-attrs")]
            extra_attrs: Default::default(),
            extra_children: vec![],
        };

        let body = types::Body {
            block_content: vec![
                types::BlockContent::P(Box::new(para1)),
                types::BlockContent::P(Box::new(para3)),
                types::BlockContent::P(Box::new(para2)),
            ],
            #[cfg(feature = "wml-layout")]
            sect_pr: None,
            extra_children: vec![],
        };

        let exprs = body.math_expressions();
        assert_eq!(exprs.len(), 2);
        assert!(!exprs[0].is_display);
        assert!(exprs[1].is_display);
    }
}
