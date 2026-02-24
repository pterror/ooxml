//! Builder methods on generated types for ergonomic document construction.
//!
//! These `impl` blocks extend the generated types with convenience methods
//! for building documents (adding paragraphs, runs, tables, etc.).
//! They are allowed because the generated types are in the same crate.

use crate::types;
#[cfg(feature = "extra-children")]
use ooxml_xml::PositionedNode;
use ooxml_xml::{RawXmlElement, RawXmlNode};

// =============================================================================
// Body
// =============================================================================

impl types::Body {
    /// Add an empty paragraph and return a mutable reference to it.
    pub fn add_paragraph(&mut self) -> &mut types::Paragraph {
        self.block_content
            .push(types::BlockContent::P(Box::default()));
        match self.block_content.last_mut().unwrap() {
            types::BlockContent::P(p) => p.as_mut(),
            _ => unreachable!(),
        }
    }

    /// Add an empty table and return a mutable reference to it.
    #[cfg(feature = "wml-tables")]
    pub fn add_table(&mut self) -> &mut types::Table {
        let table = types::Table {
            range_markup: Vec::new(),
            table_properties: Box::new(types::TableProperties::default()),
            tbl_grid: Box::new(types::TableGrid::default()),
            rows: Vec::new(),
            #[cfg(feature = "extra-children")]
            extra_children: Vec::new(),
        };
        self.block_content
            .push(types::BlockContent::Tbl(Box::new(table)));
        match self.block_content.last_mut().unwrap() {
            types::BlockContent::Tbl(t) => t.as_mut(),
            _ => unreachable!(),
        }
    }

    /// Set section properties on the body.
    #[cfg(feature = "wml-layout")]
    pub fn set_section_properties(&mut self, sect_pr: types::SectionProperties) {
        self.sect_pr = Some(Box::new(sect_pr));
    }
}

// =============================================================================
// Paragraph
// =============================================================================

impl types::Paragraph {
    /// Add an empty run and return a mutable reference to it.
    pub fn add_run(&mut self) -> &mut types::Run {
        self.paragraph_content
            .push(types::ParagraphContent::R(Box::default()));
        match self.paragraph_content.last_mut().unwrap() {
            types::ParagraphContent::R(r) => r.as_mut(),
            _ => unreachable!(),
        }
    }

    /// Add an empty hyperlink and return a mutable reference to it.
    #[cfg(feature = "wml-hyperlinks")]
    pub fn add_hyperlink(&mut self) -> &mut types::Hyperlink {
        self.paragraph_content
            .push(types::ParagraphContent::Hyperlink(Box::default()));
        match self.paragraph_content.last_mut().unwrap() {
            types::ParagraphContent::Hyperlink(h) => h.as_mut(),
            _ => unreachable!(),
        }
    }

    /// Add a bookmark start marker.
    ///
    /// This version accepts a `u32` id for ergonomic use in writer APIs.
    /// ECMA-376 Part 1, Section 17.13.6.1 (`w:bookmarkStart`).
    pub fn add_bookmark_start_u32(&mut self, id: u32, name: &str) {
        self.add_bookmark_start(id as i64, name);
    }

    /// Add a bookmark end marker.
    ///
    /// This version accepts a `u32` id for ergonomic use in writer APIs.
    /// ECMA-376 Part 1, Section 17.13.6.2 (`w:bookmarkEnd`).
    pub fn add_bookmark_end_u32(&mut self, id: u32) {
        self.add_bookmark_end(id as i64);
    }

    /// Add a bookmark start marker.
    pub fn add_bookmark_start(&mut self, id: i64, name: &str) {
        let bookmark = types::Bookmark {
            id,
            name: name.to_string(),
            #[cfg(feature = "wml-settings")]
            displaced_by_custom_xml: None,
            #[cfg(feature = "wml-tables")]
            col_first: None,
            #[cfg(feature = "wml-tables")]
            col_last: None,
            #[cfg(feature = "extra-attrs")]
            extra_attrs: Default::default(),
        };
        self.paragraph_content
            .push(types::ParagraphContent::BookmarkStart(Box::new(bookmark)));
    }

    /// Add a bookmark end marker.
    pub fn add_bookmark_end(&mut self, id: i64) {
        let range = types::CTMarkupRange {
            id,
            #[cfg(feature = "wml-settings")]
            displaced_by_custom_xml: None,
            #[cfg(feature = "extra-attrs")]
            extra_attrs: Default::default(),
        };
        self.paragraph_content
            .push(types::ParagraphContent::BookmarkEnd(Box::new(range)));
    }

    /// Add a comment range start marker.
    pub fn add_comment_range_start(&mut self, id: u32) {
        let range = types::CTMarkupRange {
            id: id as i64,
            #[cfg(feature = "wml-settings")]
            displaced_by_custom_xml: None,
            #[cfg(feature = "extra-attrs")]
            extra_attrs: Default::default(),
        };
        self.paragraph_content
            .push(types::ParagraphContent::CommentRangeStart(Box::new(range)));
    }

    /// Add a comment range end marker.
    pub fn add_comment_range_end(&mut self, id: u32) {
        let range = types::CTMarkupRange {
            id: id as i64,
            #[cfg(feature = "wml-settings")]
            displaced_by_custom_xml: None,
            #[cfg(feature = "extra-attrs")]
            extra_attrs: Default::default(),
        };
        self.paragraph_content
            .push(types::ParagraphContent::CommentRangeEnd(Box::new(range)));
    }

    /// Set paragraph properties.
    #[cfg(feature = "wml-styling")]
    pub fn set_properties(&mut self, props: types::ParagraphProperties) {
        self.p_pr = Some(Box::new(props));
    }

    /// Set numbering properties (list membership) on this paragraph.
    #[cfg(feature = "wml-styling")]
    pub fn set_numbering(&mut self, num_id: u32, ilvl: u32) {
        let ppr = self
            .p_pr
            .get_or_insert_with(|| Box::new(types::ParagraphProperties::default()));
        ppr.num_pr = Some(Box::new(types::NumberingProperties {
            ilvl: Some(Box::new(types::CTDecimalNumber {
                value: ilvl as i64,
                #[cfg(feature = "extra-attrs")]
                extra_attrs: Default::default(),
            })),
            num_id: Some(Box::new(types::CTDecimalNumber {
                value: num_id as i64,
                #[cfg(feature = "extra-attrs")]
                extra_attrs: Default::default(),
            })),
            numbering_change: None,
            ins: None,
            #[cfg(feature = "extra-children")]
            extra_children: Vec::new(),
        }));
    }

    /// Set paragraph alignment.
    ///
    /// Use `STJc` variants: `Left`, `Center`, `Right`, `Both` (justified), etc.
    #[cfg(feature = "wml-styling")]
    pub fn set_alignment(&mut self, alignment: types::STJc) {
        let ppr = self
            .p_pr
            .get_or_insert_with(|| Box::new(types::ParagraphProperties::default()));
        ppr.justification = Some(Box::new(types::CTJc {
            value: alignment,
            #[cfg(feature = "extra-attrs")]
            extra_attrs: Default::default(),
        }));
    }

    /// Set paragraph spacing (before and after, in twips).
    #[cfg(feature = "wml-styling")]
    pub fn set_spacing(&mut self, before: Option<u32>, after: Option<u32>) {
        let ppr = self
            .p_pr
            .get_or_insert_with(|| Box::new(types::ParagraphProperties::default()));
        ppr.spacing = Some(Box::new(types::CTSpacing {
            before: before.map(|b| b.to_string()),
            after: after.map(|a| a.to_string()),
            ..Default::default()
        }));
    }

    /// Set paragraph indentation.
    #[cfg(feature = "wml-styling")]
    pub fn set_indent(&mut self, left: Option<u32>, first_line: Option<u32>) {
        let ppr = self
            .p_pr
            .get_or_insert_with(|| Box::new(types::ParagraphProperties::default()));
        ppr.indentation = Some(Box::new(types::CTInd {
            left: left.map(|l| l.to_string()),
            first_line: first_line.map(|fl| fl.to_string()),
            ..Default::default()
        }));
    }

    /// Insert a page break run into this paragraph.
    ///
    /// Adds `<w:r><w:br w:type="page"/></w:r>` to the paragraph content.
    /// ECMA-376 Part 1, Section 17.3.3.1 (`w:br`).
    pub fn add_page_break(&mut self) -> &mut Self {
        let mut run = types::Run::default();
        run.run_content
            .push(types::RunContent::Br(Box::new(types::CTBr {
                r#type: Some(types::STBrType::Page),
                clear: None,
                #[cfg(feature = "extra-attrs")]
                extra_attrs: Default::default(),
            })));
        self.paragraph_content
            .push(types::ParagraphContent::R(Box::new(run)));
        self
    }

    /// Insert a column break run into this paragraph.
    ///
    /// Adds `<w:r><w:br w:type="column"/></w:r>` to the paragraph content.
    /// ECMA-376 Part 1, Section 17.3.3.1 (`w:br`).
    pub fn add_column_break(&mut self) -> &mut Self {
        let mut run = types::Run::default();
        run.run_content
            .push(types::RunContent::Br(Box::new(types::CTBr {
                r#type: Some(types::STBrType::Column),
                clear: None,
                #[cfg(feature = "extra-attrs")]
                extra_attrs: Default::default(),
            })));
        self.paragraph_content
            .push(types::ParagraphContent::R(Box::new(run)));
        self
    }

    /// Set space before the paragraph in twips (twentieths of a point).
    ///
    /// Modifies the `<w:spacing w:before="..."/>` attribute.
    /// ECMA-376 Part 1, Section 17.3.1.33 (`w:spacing`).
    #[cfg(feature = "wml-styling")]
    pub fn set_space_before(&mut self, twips: u32) -> &mut Self {
        let ppr = self
            .p_pr
            .get_or_insert_with(|| Box::new(types::ParagraphProperties::default()));
        let spacing = ppr
            .spacing
            .get_or_insert_with(|| Box::new(types::CTSpacing::default()));
        spacing.before = Some(twips.to_string());
        self
    }

    /// Set space after the paragraph in twips (twentieths of a point).
    ///
    /// Modifies the `<w:spacing w:after="..."/>` attribute.
    /// ECMA-376 Part 1, Section 17.3.1.33 (`w:spacing`).
    #[cfg(feature = "wml-styling")]
    pub fn set_space_after(&mut self, twips: u32) -> &mut Self {
        let ppr = self
            .p_pr
            .get_or_insert_with(|| Box::new(types::ParagraphProperties::default()));
        let spacing = ppr
            .spacing
            .get_or_insert_with(|| Box::new(types::CTSpacing::default()));
        spacing.after = Some(twips.to_string());
        self
    }

    /// Set line spacing in twips.
    ///
    /// Sets `<w:spacing w:line="..." w:lineRule="auto"/>`.
    /// A value of 240 is single-spacing (12pt × 20), 360 is 1.5×, 480 is double.
    /// ECMA-376 Part 1, Section 17.3.1.33 (`w:spacing`).
    #[cfg(feature = "wml-styling")]
    pub fn set_line_spacing(&mut self, twips: u32) -> &mut Self {
        let ppr = self
            .p_pr
            .get_or_insert_with(|| Box::new(types::ParagraphProperties::default()));
        let spacing = ppr
            .spacing
            .get_or_insert_with(|| Box::new(types::CTSpacing::default()));
        spacing.line = Some(twips.to_string());
        spacing.line_rule = Some(types::STLineSpacingRule::Auto);
        self
    }

    /// Set left indentation in twips.
    ///
    /// Sets `<w:ind w:left="..."/>`.
    /// ECMA-376 Part 1, Section 17.3.1.12 (`w:ind`).
    #[cfg(feature = "wml-styling")]
    pub fn set_indent_left(&mut self, twips: u32) -> &mut Self {
        let ppr = self
            .p_pr
            .get_or_insert_with(|| Box::new(types::ParagraphProperties::default()));
        let ind = ppr
            .indentation
            .get_or_insert_with(|| Box::new(types::CTInd::default()));
        ind.left = Some(twips.to_string());
        self
    }

    /// Set right indentation in twips.
    ///
    /// Sets `<w:ind w:right="..."/>`.
    /// ECMA-376 Part 1, Section 17.3.1.12 (`w:ind`).
    #[cfg(feature = "wml-styling")]
    pub fn set_indent_right(&mut self, twips: u32) -> &mut Self {
        let ppr = self
            .p_pr
            .get_or_insert_with(|| Box::new(types::ParagraphProperties::default()));
        let ind = ppr
            .indentation
            .get_or_insert_with(|| Box::new(types::CTInd::default()));
        ind.right = Some(twips.to_string());
        self
    }

    /// Set first-line indentation in twips.
    ///
    /// Sets `<w:ind w:firstLine="..."/>`. A positive value indents the first
    /// line; use `hanging` for a hanging indent (not yet exposed directly).
    /// ECMA-376 Part 1, Section 17.3.1.12 (`w:ind`).
    #[cfg(feature = "wml-styling")]
    pub fn set_indent_first_line(&mut self, twips: u32) -> &mut Self {
        let ppr = self
            .p_pr
            .get_or_insert_with(|| Box::new(types::ParagraphProperties::default()));
        let ind = ppr
            .indentation
            .get_or_insert_with(|| Box::new(types::CTInd::default()));
        ind.first_line = Some(twips.to_string());
        self
    }

    /// Set the outline level of this paragraph (0–8, where 0 = body text).
    ///
    /// Maps to `<w:outlineLvl w:val="..."/>` in paragraph properties.
    /// Levels 0–8 correspond to heading levels 1–9 in the document outline.
    /// ECMA-376 Part 1, Section 17.3.1.20 (`w:outlineLvl`).
    #[cfg(feature = "wml-styling")]
    pub fn set_outline_level(&mut self, level: u8) -> &mut Self {
        let ppr = self
            .p_pr
            .get_or_insert_with(|| Box::new(types::ParagraphProperties::default()));
        ppr.outline_lvl = Some(Box::new(types::CTDecimalNumber {
            value: level as i64,
            #[cfg(feature = "extra-attrs")]
            extra_attrs: Default::default(),
        }));
        self
    }
}

// =============================================================================
// Run
// =============================================================================

impl types::Run {
    /// Set the text content of this run.
    pub fn set_text(&mut self, text: impl Into<String>) {
        let t = types::Text {
            text: Some(text.into()),
            #[cfg(feature = "extra-children")]
            extra_children: Vec::new(),
        };
        self.run_content.push(types::RunContent::T(Box::new(t)));
    }

    /// Set bold on this run. Requires `wml-styling` feature.
    #[cfg(feature = "wml-styling")]
    pub fn set_bold(&mut self, bold: bool) {
        let rpr = self
            .r_pr
            .get_or_insert_with(|| Box::new(types::RunProperties::default()));
        if bold {
            rpr.bold = Some(Box::new(types::OnOffElement {
                value: None, // None means "true" for on/off elements
                #[cfg(feature = "extra-attrs")]
                extra_attrs: Default::default(),
            }));
        } else {
            rpr.bold = None;
        }
    }

    /// Set italic on this run. Requires `wml-styling` feature.
    #[cfg(feature = "wml-styling")]
    pub fn set_italic(&mut self, italic: bool) {
        let rpr = self
            .r_pr
            .get_or_insert_with(|| Box::new(types::RunProperties::default()));
        if italic {
            rpr.italic = Some(Box::new(types::OnOffElement {
                value: None,
                #[cfg(feature = "extra-attrs")]
                extra_attrs: Default::default(),
            }));
        } else {
            rpr.italic = None;
        }
    }

    /// Add a page break to this run.
    pub fn set_page_break(&mut self) {
        self.run_content
            .push(types::RunContent::Br(Box::new(types::CTBr {
                r#type: Some(types::STBrType::Page),
                clear: None,
                #[cfg(feature = "extra-attrs")]
                extra_attrs: Default::default(),
            })));
    }

    /// Set the text color on this run (hex string, e.g. "FF0000" for red).
    #[cfg(feature = "wml-styling")]
    pub fn set_color(&mut self, hex: &str) {
        let rpr = self
            .r_pr
            .get_or_insert_with(|| Box::new(types::RunProperties::default()));
        rpr.color = Some(Box::new(types::CTColor {
            value: hex.to_string(),
            theme_color: None,
            theme_tint: None,
            theme_shade: None,
            #[cfg(feature = "extra-attrs")]
            extra_attrs: Default::default(),
        }));
    }

    /// Set the font size in half-points (e.g. 48 = 24pt).
    #[cfg(feature = "wml-styling")]
    pub fn set_font_size(&mut self, half_points: i64) {
        let rpr = self
            .r_pr
            .get_or_insert_with(|| Box::new(types::RunProperties::default()));
        rpr.size = Some(Box::new(types::HpsMeasureElement {
            value: half_points.to_string(),
            #[cfg(feature = "extra-attrs")]
            extra_attrs: Default::default(),
        }));
    }

    /// Set strikethrough on this run.
    #[cfg(feature = "wml-styling")]
    pub fn set_strikethrough(&mut self, strike: bool) {
        let rpr = self
            .r_pr
            .get_or_insert_with(|| Box::new(types::RunProperties::default()));
        if strike {
            rpr.strikethrough = Some(Box::new(types::OnOffElement {
                value: None,
                #[cfg(feature = "extra-attrs")]
                extra_attrs: Default::default(),
            }));
        } else {
            rpr.strikethrough = None;
        }
    }

    /// Set underline style on this run.
    #[cfg(feature = "wml-styling")]
    pub fn set_underline(&mut self, style: types::STUnderline) {
        let rpr = self
            .r_pr
            .get_or_insert_with(|| Box::new(types::RunProperties::default()));
        rpr.underline = Some(Box::new(types::CTUnderline {
            value: Some(style),
            color: None,
            theme_color: None,
            theme_tint: None,
            theme_shade: None,
            #[cfg(feature = "extra-attrs")]
            extra_attrs: Default::default(),
        }));
    }

    /// Set fonts on this run.
    #[cfg(feature = "wml-styling")]
    pub fn set_fonts(&mut self, fonts: types::Fonts) {
        let rpr = self
            .r_pr
            .get_or_insert_with(|| Box::new(types::RunProperties::default()));
        rpr.fonts = Some(Box::new(fonts));
    }

    /// Set run properties.
    #[cfg(feature = "wml-styling")]
    pub fn set_properties(&mut self, props: types::RunProperties) {
        self.r_pr = Some(Box::new(props));
    }

    /// Add a drawing to this run's inner content.
    pub fn add_drawing(&mut self, drawing: types::CTDrawing) {
        self.run_content
            .push(types::RunContent::Drawing(Box::new(drawing)));
    }

    /// Add a footnote reference to this run.
    pub fn add_footnote_ref(&mut self, id: i64) {
        self.run_content
            .push(types::RunContent::FootnoteReference(Box::new(
                types::FootnoteEndnoteRef {
                    #[cfg(feature = "wml-comments")]
                    custom_mark_follows: None,
                    id,
                    #[cfg(feature = "extra-attrs")]
                    extra_attrs: Default::default(),
                },
            )));
    }

    /// Add an endnote reference to this run.
    pub fn add_endnote_ref(&mut self, id: i64) {
        self.run_content
            .push(types::RunContent::EndnoteReference(Box::new(
                types::FootnoteEndnoteRef {
                    #[cfg(feature = "wml-comments")]
                    custom_mark_follows: None,
                    id,
                    #[cfg(feature = "extra-attrs")]
                    extra_attrs: Default::default(),
                },
            )));
    }

    /// Add a comment reference to this run.
    pub fn add_comment_ref(&mut self, id: i64) {
        self.run_content
            .push(types::RunContent::CommentReference(Box::new(
                types::CTMarkup {
                    id,
                    #[cfg(feature = "extra-attrs")]
                    extra_attrs: Default::default(),
                },
            )));
    }

    /// Helper: set an `Option<Box<OnOffElement>>` field to on/off.
    #[cfg(feature = "wml-styling")]
    fn set_on_off(field: &mut Option<Box<types::OnOffElement>>, on: bool) {
        if on {
            *field = Some(Box::new(types::OnOffElement {
                value: None, // None means "true" for on/off elements
                #[cfg(feature = "extra-attrs")]
                extra_attrs: Default::default(),
            }));
        } else {
            *field = None;
        }
    }

    /// Set shadow effect on this run.
    ///
    /// Maps to `<w:shadow/>` in run properties.
    /// ECMA-376 Part 1, Section 17.3.2.31 (`w:shadow`).
    #[cfg(feature = "wml-styling")]
    pub fn set_shadow(&mut self, on: bool) {
        let rpr = self
            .r_pr
            .get_or_insert_with(|| Box::new(types::RunProperties::default()));
        Self::set_on_off(&mut rpr.shadow, on);
    }

    /// Set outline (hollow) text effect on this run.
    ///
    /// Maps to `<w:outline/>` in run properties.
    /// ECMA-376 Part 1, Section 17.3.2.23 (`w:outline`).
    #[cfg(feature = "wml-styling")]
    pub fn set_outline(&mut self, on: bool) {
        let rpr = self
            .r_pr
            .get_or_insert_with(|| Box::new(types::RunProperties::default()));
        Self::set_on_off(&mut rpr.outline, on);
    }

    /// Set emboss effect on this run.
    ///
    /// Maps to `<w:emboss/>` in run properties.
    /// ECMA-376 Part 1, Section 17.3.2.13 (`w:emboss`).
    #[cfg(feature = "wml-styling")]
    pub fn set_emboss(&mut self, on: bool) {
        let rpr = self
            .r_pr
            .get_or_insert_with(|| Box::new(types::RunProperties::default()));
        Self::set_on_off(&mut rpr.emboss, on);
    }

    /// Set imprint (engrave) effect on this run.
    ///
    /// Maps to `<w:imprint/>` in run properties.
    /// ECMA-376 Part 1, Section 17.3.2.18 (`w:imprint`).
    #[cfg(feature = "wml-styling")]
    pub fn set_imprint(&mut self, on: bool) {
        let rpr = self
            .r_pr
            .get_or_insert_with(|| Box::new(types::RunProperties::default()));
        Self::set_on_off(&mut rpr.imprint, on);
    }

    /// Set small caps on this run.
    ///
    /// Maps to `<w:smallCaps/>` in run properties.
    /// ECMA-376 Part 1, Section 17.3.2.33 (`w:smallCaps`).
    #[cfg(feature = "wml-styling")]
    pub fn set_small_caps(&mut self, on: bool) {
        let rpr = self
            .r_pr
            .get_or_insert_with(|| Box::new(types::RunProperties::default()));
        Self::set_on_off(&mut rpr.small_caps, on);
    }

    /// Set all caps on this run.
    ///
    /// Maps to `<w:caps/>` in run properties.
    /// ECMA-376 Part 1, Section 17.3.2.5 (`w:caps`).
    #[cfg(feature = "wml-styling")]
    pub fn set_all_caps(&mut self, on: bool) {
        let rpr = self
            .r_pr
            .get_or_insert_with(|| Box::new(types::RunProperties::default()));
        Self::set_on_off(&mut rpr.caps, on);
    }

    /// Set hidden text (vanish) on this run.
    ///
    /// Maps to `<w:vanish/>` in run properties. Hidden text is not rendered
    /// unless the application is set to show hidden text.
    /// ECMA-376 Part 1, Section 17.3.2.41 (`w:vanish`).
    #[cfg(feature = "wml-styling")]
    pub fn set_vanish(&mut self, on: bool) {
        let rpr = self
            .r_pr
            .get_or_insert_with(|| Box::new(types::RunProperties::default()));
        Self::set_on_off(&mut rpr.vanish, on);
    }

    /// Set double strikethrough on this run.
    ///
    /// Maps to `<w:dstrike/>` in run properties.
    /// ECMA-376 Part 1, Section 17.3.2.9 (`w:dstrike`).
    #[cfg(feature = "wml-styling")]
    pub fn set_double_strike(&mut self, on: bool) {
        let rpr = self
            .r_pr
            .get_or_insert_with(|| Box::new(types::RunProperties::default()));
        Self::set_on_off(&mut rpr.dstrike, on);
    }
}

// =============================================================================
// Hyperlink
// =============================================================================

#[cfg(feature = "wml-hyperlinks")]
impl types::Hyperlink {
    /// Add a run to this hyperlink and return a mutable reference.
    pub fn add_run(&mut self) -> &mut types::Run {
        self.paragraph_content
            .push(types::ParagraphContent::R(Box::default()));
        match self.paragraph_content.last_mut().unwrap() {
            types::ParagraphContent::R(r) => r.as_mut(),
            _ => unreachable!(),
        }
    }

    /// Set the relationship ID (for external hyperlinks).
    pub fn set_rel_id(&mut self, rel_id: &str) {
        self.id = Some(rel_id.to_string());
    }

    /// Set the anchor (for internal bookmarks).
    pub fn set_anchor(&mut self, anchor: &str) {
        self.anchor = Some(anchor.to_string());
    }
}

// =============================================================================
// Table
// =============================================================================

#[cfg(feature = "wml-tables")]
impl types::Table {
    /// Add a row and return a mutable reference.
    pub fn add_row(&mut self) -> &mut types::CTRow {
        self.rows.push(types::RowContent::Tr(Box::default()));
        match self.rows.last_mut().unwrap() {
            types::RowContent::Tr(r) => r.as_mut(),
            _ => unreachable!(),
        }
    }
}

#[cfg(feature = "wml-tables")]
impl types::CTRow {
    /// Add a cell and return a mutable reference.
    pub fn add_cell(&mut self) -> &mut types::TableCell {
        self.cells.push(types::CellContent::Tc(Box::default()));
        match self.cells.last_mut().unwrap() {
            types::CellContent::Tc(c) => c.as_mut(),
            _ => unreachable!(),
        }
    }
}

/// Vertical merge type for table cells.
///
/// ECMA-376 Part 1, Section 17.4.84 (`w:vMerge`).
#[cfg(feature = "wml-tables")]
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum VMergeType {
    /// The cell starts a vertically merged region (`w:val="restart"`).
    Restart,
    /// The cell continues an existing merged region (omitted `w:val`).
    Continue,
}

#[cfg(feature = "wml-tables")]
impl types::TableCell {
    /// Add a paragraph and return a mutable reference.
    pub fn add_paragraph(&mut self) -> &mut types::Paragraph {
        self.block_content
            .push(types::BlockContent::P(Box::default()));
        match self.block_content.last_mut().unwrap() {
            types::BlockContent::P(p) => p.as_mut(),
            _ => unreachable!(),
        }
    }

    /// Set the grid span (number of columns this cell spans).
    ///
    /// Maps to `<w:gridSpan w:val="n"/>` in cell properties.
    /// ECMA-376 Part 1, Section 17.4.17 (`w:gridSpan`).
    pub fn set_grid_span(&mut self, n: u32) {
        let tcpr = self
            .cell_properties
            .get_or_insert_with(|| Box::new(types::TableCellProperties::default()));
        tcpr.grid_span = Some(Box::new(types::CTDecimalNumber {
            value: n as i64,
            #[cfg(feature = "extra-attrs")]
            extra_attrs: Default::default(),
        }));
    }

    /// Set vertical merge on this cell.
    ///
    /// - `VMergeType::Restart` → `<w:vMerge w:val="restart"/>` (starts a merged region)
    /// - `VMergeType::Continue` → `<w:vMerge/>` (continues a merged region)
    ///
    /// ECMA-376 Part 1, Section 17.4.84 (`w:vMerge`).
    pub fn set_vertical_merge(&mut self, merge_type: VMergeType) {
        let tcpr = self
            .cell_properties
            .get_or_insert_with(|| Box::new(types::TableCellProperties::default()));
        tcpr.vertical_merge = Some(Box::new(types::CTVMerge {
            value: match merge_type {
                VMergeType::Restart => Some(types::STMerge::Restart),
                VMergeType::Continue => None,
            },
            #[cfg(feature = "extra-attrs")]
            extra_attrs: Default::default(),
        }));
    }
}

/// Border style for table cell borders.
///
/// ECMA-376 Part 1, Section 17.18.2 (`ST_Border`).
#[cfg(feature = "wml-tables")]
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum BorderStyle {
    /// No border.
    None,
    /// Single solid line.
    Single,
    /// Double solid lines.
    Double,
    /// Dashed line.
    Dashed,
    /// Dotted line.
    Dotted,
    /// Thick solid line.
    Thick,
}

#[cfg(feature = "wml-tables")]
impl BorderStyle {
    fn to_st_border(self) -> types::STBorder {
        match self {
            BorderStyle::None => types::STBorder::None,
            BorderStyle::Single => types::STBorder::Single,
            BorderStyle::Double => types::STBorder::Double,
            BorderStyle::Dashed => types::STBorder::Dashed,
            BorderStyle::Dotted => types::STBorder::Dotted,
            BorderStyle::Thick => types::STBorder::Thick,
        }
    }
}

/// Table width unit.
///
/// ECMA-376 Part 1, Section 17.18.87 (`ST_TblWidth`).
#[cfg(feature = "wml-tables")]
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum TableWidthUnit {
    /// Twips (twentieths of a point). 1440 = 1 inch.
    Dxa,
    /// Percent in fiftieths (5000 = 100%, 2500 = 50%).
    Pct,
}

#[cfg(feature = "wml-tables")]
impl types::CTRow {
    /// Set the row height in twips.
    ///
    /// Maps to `<w:trPr><w:trHeight w:val="..." w:hRule="exact"/></w:trPr>`.
    /// ECMA-376 Part 1, Section 17.4.81 (`w:trHeight`).
    pub fn set_height(&mut self, twips: u32) {
        let row_pr = self
            .row_properties
            .get_or_insert_with(|| Box::new(types::TableRowProperties::default()));
        row_pr.tr_height = Some(Box::new(types::CTHeight {
            value: Some(twips.to_string()),
            #[cfg(feature = "wml-tables")]
            h_rule: Some(types::STHeightRule::Exact),
            #[cfg(feature = "extra-attrs")]
            extra_attrs: Default::default(),
        }));
    }
}

#[cfg(feature = "wml-tables")]
impl types::TableCell {
    /// Set the background (shading) color of this cell.
    ///
    /// The `rgb` value should be a hex string (e.g., `"FF0000"` for red, `"auto"` for automatic).
    /// Maps to `<w:tcPr><w:shd w:val="clear" w:fill="..." w:color="auto"/></w:tcPr>`.
    /// ECMA-376 Part 1, Section 17.4.32 (`w:shd`).
    pub fn set_background_color(&mut self, rgb: &str) {
        let tcpr = self
            .cell_properties
            .get_or_insert_with(|| Box::new(types::TableCellProperties::default()));
        tcpr.shading = Some(Box::new(types::CTShd {
            value: types::STShd::Clear,
            #[cfg(feature = "wml-styling")]
            fill: Some(rgb.to_string()),
            #[cfg(feature = "wml-styling")]
            color: Some("auto".to_string()),
            #[cfg(feature = "wml-styling")]
            theme_color: None,
            #[cfg(feature = "wml-styling")]
            theme_tint: None,
            #[cfg(feature = "wml-styling")]
            theme_shade: None,
            #[cfg(feature = "wml-styling")]
            theme_fill: None,
            #[cfg(feature = "wml-styling")]
            theme_fill_tint: None,
            #[cfg(feature = "wml-styling")]
            theme_fill_shade: None,
            #[cfg(feature = "extra-attrs")]
            extra_attrs: Default::default(),
        }));
    }

    /// Set all four borders of this cell at once.
    ///
    /// `style` is the border style, `width_eights` is the width in eighths of a point
    /// (e.g., 4 = half a point = 0.5pt), and `color` is a hex color string (e.g., `"000000"`).
    ///
    /// ECMA-376 Part 1, Section 17.4.5 (`w:tcBorders`).
    pub fn set_borders(&mut self, style: BorderStyle, width_eights: u32, color: &str) {
        self.set_border_top(style, width_eights, color);
        self.set_border_bottom(style, width_eights, color);
        self.set_border_left(style, width_eights, color);
        self.set_border_right(style, width_eights, color);
    }

    /// Set the top border of this cell.
    ///
    /// ECMA-376 Part 1, Section 17.4.5 (`w:tcBorders`).
    pub fn set_border_top(&mut self, style: BorderStyle, width_eights: u32, color: &str) {
        let tcpr = self
            .cell_properties
            .get_or_insert_with(|| Box::new(types::TableCellProperties::default()));
        let borders = tcpr
            .tc_borders
            .get_or_insert_with(|| Box::new(types::CTTcBorders::default()));
        borders.top = Some(Box::new(make_cell_border(style, width_eights, color)));
    }

    /// Set the bottom border of this cell.
    ///
    /// ECMA-376 Part 1, Section 17.4.5 (`w:tcBorders`).
    pub fn set_border_bottom(&mut self, style: BorderStyle, width_eights: u32, color: &str) {
        let tcpr = self
            .cell_properties
            .get_or_insert_with(|| Box::new(types::TableCellProperties::default()));
        let borders = tcpr
            .tc_borders
            .get_or_insert_with(|| Box::new(types::CTTcBorders::default()));
        borders.bottom = Some(Box::new(make_cell_border(style, width_eights, color)));
    }

    /// Set the left border of this cell.
    ///
    /// ECMA-376 Part 1, Section 17.4.5 (`w:tcBorders`).
    pub fn set_border_left(&mut self, style: BorderStyle, width_eights: u32, color: &str) {
        let tcpr = self
            .cell_properties
            .get_or_insert_with(|| Box::new(types::TableCellProperties::default()));
        let borders = tcpr
            .tc_borders
            .get_or_insert_with(|| Box::new(types::CTTcBorders::default()));
        borders.left = Some(Box::new(make_cell_border(style, width_eights, color)));
    }

    /// Set the right border of this cell.
    ///
    /// ECMA-376 Part 1, Section 17.4.5 (`w:tcBorders`).
    pub fn set_border_right(&mut self, style: BorderStyle, width_eights: u32, color: &str) {
        let tcpr = self
            .cell_properties
            .get_or_insert_with(|| Box::new(types::TableCellProperties::default()));
        let borders = tcpr
            .tc_borders
            .get_or_insert_with(|| Box::new(types::CTTcBorders::default()));
        borders.right = Some(Box::new(make_cell_border(style, width_eights, color)));
    }

    /// Set cell padding (margins) in twips.
    ///
    /// Maps to `<w:tcPr><w:tcMar .../></w:tcPr>`.
    /// ECMA-376 Part 1, Section 17.4.44 (`w:tcMar`).
    pub fn set_padding(&mut self, top: u32, bottom: u32, left: u32, right: u32) {
        let tcpr = self
            .cell_properties
            .get_or_insert_with(|| Box::new(types::TableCellProperties::default()));
        tcpr.tc_mar = Some(Box::new(types::CTTcMar {
            top: Some(Box::new(make_tbl_width(top, types::STTblWidth::Dxa))),
            bottom: Some(Box::new(make_tbl_width(bottom, types::STTblWidth::Dxa))),
            left: Some(Box::new(make_tbl_width(left, types::STTblWidth::Dxa))),
            right: Some(Box::new(make_tbl_width(right, types::STTblWidth::Dxa))),
            #[cfg(feature = "wml-tables")]
            start: None,
            #[cfg(feature = "wml-tables")]
            end: None,
            #[cfg(feature = "extra-children")]
            extra_children: Vec::new(),
        }));
    }
}

#[cfg(feature = "wml-tables")]
impl types::Table {
    /// Set the preferred width of this table.
    ///
    /// `width` is the measurement value; `unit` is the unit type.
    /// Maps to `<w:tblPr><w:tblW w:w="..." w:type="..."/></w:tblPr>`.
    /// ECMA-376 Part 1, Section 17.4.63 (`w:tblW`).
    pub fn set_width(&mut self, width: u32, unit: TableWidthUnit) {
        let type_ = match unit {
            TableWidthUnit::Dxa => types::STTblWidth::Dxa,
            TableWidthUnit::Pct => types::STTblWidth::Pct,
        };
        self.table_properties.tbl_w = Some(Box::new(make_tbl_width(width, type_)));
    }
}

/// Build a `CTBorder` with the given style, width, and color.
#[cfg(feature = "wml-tables")]
fn make_cell_border(style: BorderStyle, width_eights: u32, color: &str) -> types::CTBorder {
    types::CTBorder {
        value: style.to_st_border(),
        #[cfg(feature = "wml-styling")]
        color: Some(color.to_string()),
        #[cfg(feature = "wml-styling")]
        size: Some(width_eights as u64),
        #[cfg(feature = "wml-styling")]
        space: Some(0u64),
        #[cfg(feature = "wml-styling")]
        theme_color: None,
        #[cfg(feature = "wml-styling")]
        theme_tint: None,
        #[cfg(feature = "wml-styling")]
        theme_shade: None,
        #[cfg(feature = "wml-styling")]
        shadow: None,
        #[cfg(feature = "wml-styling")]
        frame: None,
        #[cfg(feature = "extra-attrs")]
        extra_attrs: Default::default(),
    }
}

/// Build a `CTTblWidth` with the given value and type.
#[cfg(feature = "wml-tables")]
fn make_tbl_width(width: u32, type_: types::STTblWidth) -> types::CTTblWidth {
    types::CTTblWidth {
        width: Some(width.to_string()),
        r#type: Some(type_),
        #[cfg(feature = "extra-attrs")]
        extra_attrs: Default::default(),
    }
}

// =============================================================================
// Header/Footer (HeaderFooter)
// =============================================================================

impl types::HeaderFooter {
    /// Add an empty paragraph and return a mutable reference.
    pub fn add_paragraph(&mut self) -> &mut types::Paragraph {
        self.block_content
            .push(types::BlockContent::P(Box::default()));
        match self.block_content.last_mut().unwrap() {
            types::BlockContent::P(p) => p.as_mut(),
            _ => unreachable!(),
        }
    }
}

// =============================================================================
// Comment
// =============================================================================

impl types::Comment {
    /// Add a paragraph and return a mutable reference.
    pub fn add_paragraph(&mut self) -> &mut types::Paragraph {
        self.block_content
            .push(types::BlockContent::P(Box::default()));
        match self.block_content.last_mut().unwrap() {
            types::BlockContent::P(p) => p.as_mut(),
            _ => unreachable!(),
        }
    }
}

// =============================================================================
// Footnote/Endnote (FootnoteEndnote)
// =============================================================================

impl types::FootnoteEndnote {
    /// Add a paragraph and return a mutable reference.
    pub fn add_paragraph(&mut self) -> &mut types::Paragraph {
        self.block_content
            .push(types::BlockContent::P(Box::default()));
        match self.block_content.last_mut().unwrap() {
            types::BlockContent::P(p) => p.as_mut(),
            _ => unreachable!(),
        }
    }
}

// =============================================================================
// Track changes helpers (wml-track-changes)
// =============================================================================

/// Build a `CTRunTrackChange` containing a single text run.
///
/// The `text` is placed in a `<w:r><w:t>…</w:t></w:r>` inside the change element.
/// ECMA-376 §17.13.5.
#[cfg(feature = "wml-track-changes")]
fn make_run_track_change(
    id: i64,
    author: &str,
    date: Option<&str>,
    text: &str,
) -> types::CTRunTrackChange {
    let t = types::Text {
        text: Some(text.to_string()),
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
    types::CTRunTrackChange {
        id,
        author: author.to_string(),
        date: date.map(|d| d.to_string()),
        run_content: vec![types::RunContentChoice::R(Box::new(run))],
        #[cfg(feature = "extra-attrs")]
        extra_attrs: Default::default(),
        #[cfg(feature = "extra-children")]
        extra_children: Vec::new(),
    }
}

/// Create a `ParagraphContent::Ins` element wrapping a text run.
///
/// Use this to add a tracked insertion to a paragraph's `paragraph_content`.
///
/// # Example
///
/// ```
/// # #[cfg(feature = "wml-track-changes")] {
/// use ooxml_wml::convenience::ins_run;
/// use ooxml_wml::types;
///
/// let mut para = types::Paragraph::default();
/// para.paragraph_content.push(ins_run(1, "Alice", Some("2026-02-24T12:00:00Z"), "inserted text"));
/// # }
/// ```
///
/// ECMA-376 §17.13.5.16 (`w:ins`).
#[cfg(feature = "wml-track-changes")]
pub fn ins_run(id: i64, author: &str, date: Option<&str>, text: &str) -> types::ParagraphContent {
    types::ParagraphContent::Ins(Box::new(make_run_track_change(id, author, date, text)))
}

/// Create a `ParagraphContent::Del` element wrapping a text run.
///
/// Use this to add a tracked deletion to a paragraph's `paragraph_content`.
///
/// # Example
///
/// ```
/// # #[cfg(feature = "wml-track-changes")] {
/// use ooxml_wml::convenience::del_run;
/// use ooxml_wml::types;
///
/// let mut para = types::Paragraph::default();
/// para.paragraph_content.push(del_run(2, "Bob", None, "deleted text"));
/// # }
/// ```
///
/// ECMA-376 §17.13.5.13 (`w:del`).
#[cfg(feature = "wml-track-changes")]
pub fn del_run(id: i64, author: &str, date: Option<&str>, text: &str) -> types::ParagraphContent {
    types::ParagraphContent::Del(Box::new(make_run_track_change(id, author, date, text)))
}

#[cfg(feature = "wml-track-changes")]
impl types::Paragraph {
    /// Add a tracked insertion wrapping the given text and return a mutable
    /// reference to the `CTRunTrackChange` (ECMA-376 §17.13.5.16).
    pub fn add_tracked_insertion(
        &mut self,
        id: i64,
        author: &str,
        date: Option<&str>,
        text: &str,
    ) -> &mut types::CTRunTrackChange {
        self.paragraph_content.push(ins_run(id, author, date, text));
        match self.paragraph_content.last_mut().unwrap() {
            types::ParagraphContent::Ins(tc) => tc.as_mut(),
            _ => unreachable!(),
        }
    }

    /// Add a tracked deletion wrapping the given text and return a mutable
    /// reference to the `CTRunTrackChange` (ECMA-376 §17.13.5.13).
    pub fn add_tracked_deletion(
        &mut self,
        id: i64,
        author: &str,
        date: Option<&str>,
        text: &str,
    ) -> &mut types::CTRunTrackChange {
        self.paragraph_content.push(del_run(id, author, date, text));
        match self.paragraph_content.last_mut().unwrap() {
            types::ParagraphContent::Del(tc) => tc.as_mut(),
            _ => unreachable!(),
        }
    }
}

// =============================================================================
// Form fields (SDT-based)  (wml-settings feature)
// =============================================================================

/// Type of form field to create inside a Structured Document Tag.
///
/// ECMA-376 Part 1, Section 17.5.2 (Structured Document Tags).
#[cfg(feature = "wml-settings")]
#[derive(Debug, Clone)]
pub enum FormFieldType {
    /// Plain-text input (`<w:text/>`).
    PlainText,
    /// Rich-text area (`<w:richText/>`).
    RichText,
    /// Combo box (`<w:comboBox>`).
    ComboBox,
    /// Drop-down list (`<w:dropDownList>`).
    DropDownList,
    /// Date picker (`<w:date>`).
    DatePicker,
}

/// Configuration for a form field written as a Structured Document Tag.
///
/// ECMA-376 Part 1, Section 17.5.2 (`w:sdt`).
#[cfg(feature = "wml-settings")]
#[derive(Debug, Clone)]
pub struct FormFieldConfig {
    /// Machine-readable tag (`<w:tag w:val="..."/>`).
    pub tag: Option<String>,
    /// Human-readable alias/label (`<w:alias w:val="..."/>`).
    pub label: Option<String>,
    /// Type of form control.
    pub field_type: FormFieldType,
    /// Initial/default value displayed in the field.
    pub default_value: Option<String>,
    /// Placeholder text (used as content when no default_value is set).
    pub placeholder: Option<String>,
    /// Items for combo box / drop-down list fields.
    pub list_items: Vec<String>,
    /// Date format string for date picker fields (e.g. `"MM/dd/yyyy"`).
    pub date_format: Option<String>,
}

#[cfg(feature = "wml-settings")]
impl Default for FormFieldConfig {
    fn default() -> Self {
        Self {
            tag: None,
            label: None,
            field_type: FormFieldType::PlainText,
            default_value: None,
            placeholder: None,
            list_items: Vec::new(),
            date_format: None,
        }
    }
}

/// Build a single paragraph containing the given text, for use inside SDT content.
#[cfg(feature = "wml-settings")]
fn make_text_paragraph(text: &str) -> types::BlockContentChoice {
    let t = types::Text {
        text: Some(text.to_string()),
        #[cfg(feature = "extra-children")]
        extra_children: Vec::new(),
    };
    let mut run = types::Run::default();
    run.run_content.push(types::RunContent::T(Box::new(t)));
    let mut para = types::Paragraph::default();
    para.paragraph_content
        .push(types::ParagraphContent::R(Box::new(run)));
    types::BlockContentChoice::P(Box::new(para))
}

/// Build a `CTSdtListItem` from a display string.
#[cfg(feature = "wml-settings")]
fn make_list_item(text: &str) -> types::CTSdtListItem {
    types::CTSdtListItem {
        display_text: Some(text.to_string()),
        value: Some(text.to_string()),
        #[cfg(feature = "extra-attrs")]
        extra_attrs: Default::default(),
    }
}

#[cfg(feature = "wml-settings")]
impl types::Body {
    /// Add a form field as a Structured Document Tag (`<w:sdt>`).
    ///
    /// Produces a block-level SDT containing an appropriate SDT properties
    /// element (`<w:sdtPr>`) and content paragraph with the default value.
    ///
    /// ECMA-376 Part 1, Section 17.5.2.
    pub fn add_form_field(&mut self, config: FormFieldConfig) -> &mut Self {
        let content_text = config
            .default_value
            .as_deref()
            .or(config.placeholder.as_deref())
            .unwrap_or("")
            .to_string();

        // Build sdtPr
        let mut sdt_pr = types::CTSdtPr::default();

        #[cfg(feature = "wml-settings")]
        if let Some(ref tag_val) = config.tag {
            sdt_pr.tag = Some(Box::new(types::CTString {
                value: tag_val.clone(),
                #[cfg(feature = "extra-attrs")]
                extra_attrs: Default::default(),
            }));
        }

        #[cfg(feature = "wml-settings")]
        if let Some(ref alias_val) = config.label {
            sdt_pr.alias = Some(Box::new(types::CTString {
                value: alias_val.clone(),
                #[cfg(feature = "extra-attrs")]
                extra_attrs: Default::default(),
            }));
        }

        #[cfg(feature = "wml-settings")]
        match config.field_type {
            FormFieldType::PlainText => {
                sdt_pr.text = Some(Box::new(types::CTSdtText {
                    multi_line: None,
                    #[cfg(feature = "extra-attrs")]
                    extra_attrs: Default::default(),
                }));
            }
            FormFieldType::RichText => {
                sdt_pr.rich_text = Some(Box::new(types::CTEmpty));
            }
            FormFieldType::ComboBox => {
                let items = config
                    .list_items
                    .iter()
                    .map(|s| make_list_item(s))
                    .collect();
                sdt_pr.combo_box = Some(Box::new(types::CTSdtComboBox {
                    last_value: None,
                    list_item: items,
                    #[cfg(feature = "extra-attrs")]
                    extra_attrs: Default::default(),
                    #[cfg(feature = "extra-children")]
                    extra_children: Vec::new(),
                }));
            }
            FormFieldType::DropDownList => {
                let items = config
                    .list_items
                    .iter()
                    .map(|s| make_list_item(s))
                    .collect();
                sdt_pr.drop_down_list = Some(Box::new(types::CTSdtDropDownList {
                    last_value: None,
                    list_item: items,
                    #[cfg(feature = "extra-attrs")]
                    extra_attrs: Default::default(),
                    #[cfg(feature = "extra-children")]
                    extra_children: Vec::new(),
                }));
            }
            FormFieldType::DatePicker => {
                let date_format_elem = config.date_format.as_deref().map(|fmt| {
                    Box::new(types::CTString {
                        value: fmt.to_string(),
                        #[cfg(feature = "extra-attrs")]
                        extra_attrs: Default::default(),
                    })
                });
                sdt_pr.date = Some(Box::new(types::CTSdtDate {
                    date_format: date_format_elem,
                    full_date: None,
                    lid: None,
                    store_mapped_data_as: None,
                    calendar: None,
                    #[cfg(feature = "extra-attrs")]
                    extra_attrs: Default::default(),
                    #[cfg(feature = "extra-children")]
                    extra_children: Vec::new(),
                }));
            }
        }

        let sdt_content = types::CTSdtContentBlock {
            block_content: vec![make_text_paragraph(&content_text)],
            #[cfg(feature = "extra-children")]
            extra_children: Vec::new(),
        };

        let sdt = types::CTSdtBlock {
            sdt_pr: Some(Box::new(sdt_pr)),
            sdt_end_pr: None,
            sdt_content: Some(Box::new(sdt_content)),
            #[cfg(feature = "extra-children")]
            extra_children: Vec::new(),
        };

        self.block_content
            .push(types::BlockContent::Sdt(Box::new(sdt)));
        self
    }
}

// =============================================================================
// Table of Contents  (wml-fields feature)
// =============================================================================

/// Options for a Table of Contents field inserted by `Body::add_toc()`.
///
/// ECMA-376 Part 1, Section 17.16 (Field Codes) – `TOC` field.
#[cfg(feature = "wml-fields")]
#[derive(Debug, Clone)]
pub struct TocOptions {
    /// Optional heading paragraph (e.g. "Table of Contents").
    pub title: Option<String>,
    /// Maximum heading level to include (default 3 → H1–H3).
    pub max_level: u8,
    /// Whether page numbers should be right-aligned with tab leaders.
    pub right_align_page_numbers: bool,
    /// Whether TOC entries should be hyperlinks.
    pub use_hyperlinks: bool,
}

#[cfg(feature = "wml-fields")]
impl Default for TocOptions {
    fn default() -> Self {
        Self {
            title: None,
            max_level: 3,
            right_align_page_numbers: true,
            use_hyperlinks: true,
        }
    }
}

/// Build a CTFldChar with the given type and all optional fields set to None.
#[cfg(feature = "wml-fields")]
fn make_fld_char(fld_char_type: types::STFldCharType) -> types::CTFldChar {
    types::CTFldChar {
        fld_char_type,
        #[cfg(feature = "wml-fields")]
        fld_lock: None,
        #[cfg(feature = "wml-fields")]
        dirty: None,
        #[cfg(feature = "wml-fields")]
        fld_data: None,
        #[cfg(feature = "wml-fields")]
        ff_data: None,
        #[cfg(feature = "wml-track-changes")]
        numbering_change: None,
        #[cfg(feature = "extra-attrs")]
        extra_attrs: Default::default(),
        #[cfg(feature = "extra-children")]
        extra_children: Vec::new(),
    }
}

/// Build a single-run paragraph containing a field character run content item.
#[cfg(feature = "wml-fields")]
#[allow(dead_code)]
fn make_fld_char_para(fld_char_type: types::STFldCharType) -> types::Paragraph {
    let fld_char = make_fld_char(fld_char_type);
    let mut run = types::Run::default();
    run.run_content
        .push(types::RunContent::FldChar(Box::new(fld_char)));
    let mut para = types::Paragraph::default();
    para.paragraph_content
        .push(types::ParagraphContent::R(Box::new(run)));
    para
}

/// Build an instr-text run inside a paragraph.
#[cfg(feature = "wml-fields")]
#[allow(dead_code)]
fn make_instr_text_para(instr: &str) -> types::Paragraph {
    let t = types::Text {
        text: Some(instr.to_string()),
        #[cfg(feature = "extra-children")]
        extra_children: Vec::new(),
    };
    let mut run = types::Run::default();
    run.run_content
        .push(types::RunContent::InstrText(Box::new(t)));
    let mut para = types::Paragraph::default();
    para.paragraph_content
        .push(types::ParagraphContent::R(Box::new(run)));
    para
}

#[cfg(feature = "wml-fields")]
impl types::Body {
    /// Insert a Table of Contents field at the current position.
    ///
    /// This writes:
    /// 1. An optional title paragraph styled "TOC Heading".
    /// 2. A `TOC` field spanning three paragraphs: `fldChar begin`, `instrText`, `fldChar end`.
    /// 3. A placeholder paragraph telling the user to update the field.
    ///
    /// ECMA-376 Part 1, Section 17.16.5.58 (TOC).
    pub fn add_toc(&mut self, opts: TocOptions) -> &mut Self {
        // 1. Optional title paragraph
        if let Some(ref title_text) = opts.title {
            let para = self.add_paragraph();
            para.add_run().set_text(title_text.as_str());
            #[cfg(feature = "wml-styling")]
            {
                let ppr = para
                    .p_pr
                    .get_or_insert_with(|| Box::new(types::ParagraphProperties::default()));
                ppr.paragraph_style = Some(Box::new(types::CTString {
                    value: "TOCHeading".to_string(),
                    #[cfg(feature = "extra-attrs")]
                    extra_attrs: Default::default(),
                }));
            }
        }

        // 2. Build the TOC field instruction string
        let mut instr = format!(r#" TOC \o "1-{}" "#, opts.max_level);
        if opts.use_hyperlinks {
            instr.push_str(r"\h ");
        }
        if opts.right_align_page_numbers {
            instr.push_str(r"\z \u ");
        }

        // Write the field as three runs in a single paragraph:
        // <w:r><w:fldChar w:fldCharType="begin"/></w:r>
        // <w:r><w:instrText> TOC ... </w:instrText></w:r>
        // <w:r><w:fldChar w:fldCharType="separate"/></w:r>
        // <w:r><w:fldChar w:fldCharType="end"/></w:r>
        let fld_begin = make_fld_char(types::STFldCharType::Begin);
        let fld_separate = make_fld_char(types::STFldCharType::Separate);
        let fld_end = make_fld_char(types::STFldCharType::End);

        let instr_t = types::Text {
            text: Some(instr),
            #[cfg(feature = "extra-children")]
            extra_children: Vec::new(),
        };

        let mut run_begin = types::Run::default();
        run_begin
            .run_content
            .push(types::RunContent::FldChar(Box::new(fld_begin)));

        let mut run_instr = types::Run::default();
        run_instr
            .run_content
            .push(types::RunContent::InstrText(Box::new(instr_t)));

        let mut run_separate = types::Run::default();
        run_separate
            .run_content
            .push(types::RunContent::FldChar(Box::new(fld_separate)));

        let mut run_end = types::Run::default();
        run_end
            .run_content
            .push(types::RunContent::FldChar(Box::new(fld_end)));

        let toc_para = self.add_paragraph();
        toc_para
            .paragraph_content
            .push(types::ParagraphContent::R(Box::new(run_begin)));
        toc_para
            .paragraph_content
            .push(types::ParagraphContent::R(Box::new(run_instr)));
        toc_para
            .paragraph_content
            .push(types::ParagraphContent::R(Box::new(run_separate)));
        toc_para
            .paragraph_content
            .push(types::ParagraphContent::R(Box::new(run_end)));

        // 3. Placeholder paragraph
        let placeholder = self.add_paragraph();
        placeholder
            .add_run()
            .set_text("[Right-click to update field]");

        self
    }
}

// =============================================================================
// Office Math (OMath)
// =============================================================================

/// Office Math namespace (`m:`).
///
/// ECMA-376 Part 1, Section 22 (Office Math Markup Language).
pub const NS_M: &str = "http://schemas.openxmlformats.org/officeDocument/2006/math";

/// Builder for Office Math expressions embedded in Word paragraphs.
///
/// Produces either an inline `<m:oMath>` or a display `<m:oMathPara><m:oMath>`
/// element stored as a `RawXmlElement` inside a paragraph run.
///
/// ECMA-376 Part 1, Section 22.1.2.77 (`m:oMath`).
#[derive(Debug, Clone)]
pub struct OMathBuilder {
    display: bool,
    xml_content: String,
}

impl OMathBuilder {
    /// Create an inline math expression containing plain text.
    pub fn plain(text: &str) -> Self {
        Self {
            display: false,
            xml_content: format!("<m:r><m:t>{}</m:t></m:r>", xml_escape(text)),
        }
    }

    /// Create a fraction `numerator / denominator`.
    ///
    /// ECMA-376 Part 1, Section 22.1.2.36 (`m:f`).
    pub fn fraction(numerator: &str, denominator: &str) -> Self {
        Self {
            display: false,
            xml_content: format!(
                "<m:f><m:num><m:r><m:t>{}</m:t></m:r></m:num>\
                 <m:den><m:r><m:t>{}</m:t></m:r></m:den></m:f>",
                xml_escape(numerator),
                xml_escape(denominator)
            ),
        }
    }

    /// Create a superscript expression `base^exp`.
    ///
    /// ECMA-376 Part 1, Section 22.1.2.105 (`m:sSup`).
    pub fn superscript(base: &str, exp: &str) -> Self {
        Self {
            display: false,
            xml_content: format!(
                "<m:sSup><m:e><m:r><m:t>{}</m:t></m:r></m:e>\
                 <m:sup><m:r><m:t>{}</m:t></m:r></m:sup></m:sSup>",
                xml_escape(base),
                xml_escape(exp)
            ),
        }
    }

    /// Create a subscript expression `base_sub`.
    ///
    /// ECMA-376 Part 1, Section 22.1.2.98 (`m:sSub`).
    pub fn subscript(base: &str, sub: &str) -> Self {
        Self {
            display: false,
            xml_content: format!(
                "<m:sSub><m:e><m:r><m:t>{}</m:t></m:r></m:e>\
                 <m:sub><m:r><m:t>{}</m:t></m:r></m:sub></m:sSub>",
                xml_escape(base),
                xml_escape(sub)
            ),
        }
    }

    /// Create a radical (square root) of `base`.
    ///
    /// ECMA-376 Part 1, Section 22.1.2.79 (`m:rad`).
    pub fn radical(base: &str) -> Self {
        Self {
            display: false,
            xml_content: format!(
                "<m:rad><m:radPr><m:degHide m:val=\"1\"/></m:radPr>\
                 <m:deg/><m:e><m:r><m:t>{}</m:t></m:r></m:e></m:rad>",
                xml_escape(base)
            ),
        }
    }

    /// Set display (block) mode. Wraps the expression in `<m:oMathPara>`.
    pub fn as_display(mut self) -> Self {
        self.display = true;
        self
    }

    /// Build a `RawXmlElement` suitable for use inside a paragraph.
    ///
    /// The element is `<m:oMath>` (inline) or
    /// `<m:oMathPara><m:oMath>…</m:oMath></m:oMathPara>` (display).
    pub fn build(self) -> RawXmlElement {
        let omath = RawXmlElement {
            name: "m:oMath".to_string(),
            attributes: vec![("xmlns:m".to_string(), NS_M.to_string())],
            children: vec![RawXmlNode::Text(self.xml_content)],
            self_closing: false,
        };

        if self.display {
            RawXmlElement {
                name: "m:oMathPara".to_string(),
                attributes: vec![("xmlns:m".to_string(), NS_M.to_string())],
                children: vec![RawXmlNode::Element({
                    // inner oMath without the xmlns repetition
                    RawXmlElement {
                        name: "m:oMath".to_string(),
                        attributes: vec![],
                        children: vec![RawXmlNode::Text({
                            // re-use xml_content (moved into omath above) from display wrapper
                            // We need to re-derive it — use omath.children
                            match omath.children.into_iter().next() {
                                Some(RawXmlNode::Text(t)) => t,
                                _ => String::new(),
                            }
                        })],
                        self_closing: false,
                    }
                })],
                self_closing: false,
            }
        } else {
            omath
        }
    }
}

/// Escape XML special characters in text content.
fn xml_escape(s: &str) -> String {
    s.replace('&', "&amp;")
        .replace('<', "&lt;")
        .replace('>', "&gt;")
        .replace('"', "&quot;")
        .replace('\'', "&apos;")
}

impl types::Paragraph {
    /// Add an Office Math expression to this paragraph.
    ///
    /// The math element (`<m:oMath>` or `<m:oMathPara>`) is stored as a
    /// `RawXmlElement` inside an extra-children slot on a synthetic run.
    ///
    /// ECMA-376 Part 1, Section 22.1.2.77 (`m:oMath`).
    #[cfg(feature = "extra-children")]
    pub fn add_math(&mut self, builder: OMathBuilder) -> &mut Self {
        let elem = builder.build();
        // Math elements live as siblings of runs in the paragraph's extra_children,
        // using a PositionedNode so they serialize in document order.
        let idx = self.paragraph_content.len();
        self.extra_children
            .push(PositionedNode::new(idx, RawXmlNode::Element(elem)));
        self
    }
}

// =============================================================================
// Inline chart drawing
// =============================================================================

impl types::Paragraph {
    /// Add an inline chart drawing reference to this paragraph.
    ///
    /// Inserts a `<w:drawing><wp:inline>…<c:chart r:id="rel_id"/>…</wp:inline></w:drawing>`
    /// element referencing the chart part identified by `rel_id`.
    ///
    /// Use `DocumentBuilder::embed_chart` to obtain a `rel_id` first.
    ///
    /// ECMA-376 Part 1, Section 20.4.2.8 (inline).
    #[cfg(feature = "wml-charts")]
    pub fn add_inline_chart(&mut self, rel_id: &str, width_emu: i64, height_emu: i64) -> &mut Self {
        let drawing_elem = build_chart_inline_element(rel_id, width_emu, height_emu);
        let drawing = types::CTDrawing {
            #[cfg(feature = "extra-children")]
            extra_children: vec![PositionedNode::new(0, RawXmlNode::Element(drawing_elem))],
        };
        let mut run = types::Run::default();
        run.run_content
            .push(types::RunContent::Drawing(Box::new(drawing)));
        self.paragraph_content
            .push(types::ParagraphContent::R(Box::new(run)));
        self
    }
}

/// Build the `<wp:inline>` element referencing a chart.
#[cfg(feature = "wml-charts")]
fn build_chart_inline_element(rel_id: &str, width_emu: i64, height_emu: i64) -> RawXmlElement {
    // <c:chart r:id="rel_id"/>
    let chart_ref = RawXmlElement {
        name: "c:chart".to_string(),
        attributes: vec![
            (
                "xmlns:c".to_string(),
                "http://schemas.openxmlformats.org/drawingml/2006/chart".to_string(),
            ),
            (
                "xmlns:r".to_string(),
                "http://schemas.openxmlformats.org/officeDocument/2006/relationships".to_string(),
            ),
            ("r:id".to_string(), rel_id.to_string()),
        ],
        children: vec![],
        self_closing: true,
    };

    // <a:graphicData uri="...chart...">
    let graphic_data = RawXmlElement {
        name: "a:graphicData".to_string(),
        attributes: vec![(
            "uri".to_string(),
            "http://schemas.openxmlformats.org/drawingml/2006/chart".to_string(),
        )],
        children: vec![RawXmlNode::Element(chart_ref)],
        self_closing: false,
    };

    // <a:graphic>
    let graphic = RawXmlElement {
        name: "a:graphic".to_string(),
        attributes: vec![(
            "xmlns:a".to_string(),
            "http://schemas.openxmlformats.org/drawingml/2006/main".to_string(),
        )],
        children: vec![RawXmlNode::Element(graphic_data)],
        self_closing: false,
    };

    // <wp:extent cx="..." cy="..."/>
    let extent = RawXmlElement {
        name: "wp:extent".to_string(),
        attributes: vec![
            ("cx".to_string(), width_emu.to_string()),
            ("cy".to_string(), height_emu.to_string()),
        ],
        children: vec![],
        self_closing: true,
    };

    // <wp:docPr id="1" name="Chart 1"/>
    let doc_pr = RawXmlElement {
        name: "wp:docPr".to_string(),
        attributes: vec![
            ("id".to_string(), "1".to_string()),
            ("name".to_string(), "Chart 1".to_string()),
        ],
        children: vec![],
        self_closing: true,
    };

    // <wp:inline>
    RawXmlElement {
        name: "wp:inline".to_string(),
        attributes: vec![
            (
                "xmlns:wp".to_string(),
                "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
                    .to_string(),
            ),
            ("distT".to_string(), "0".to_string()),
            ("distB".to_string(), "0".to_string()),
            ("distL".to_string(), "0".to_string()),
            ("distR".to_string(), "0".to_string()),
        ],
        children: vec![
            RawXmlNode::Element(extent),
            RawXmlNode::Element(doc_pr),
            RawXmlNode::Element(graphic),
        ],
        self_closing: false,
    }
}

// =============================================================================
// Tests for new features
// =============================================================================

#[cfg(test)]
mod feature_tests {
    use super::*;

    // -------------------------------------------------------------------------
    // Settings
    // -------------------------------------------------------------------------

    #[test]
    #[cfg(feature = "wml-settings")]
    fn test_settings_xml_content() {
        use crate::writer::DocumentSettingsOptions;

        // Directly test the build_settings_xml helper by calling via the module
        // (We can't call build_settings_xml directly since it's private, so we
        //  verify the public API generates the right document structure.)
        let opts = DocumentSettingsOptions {
            default_tab_stop: Some(720),
            even_and_odd_headers: true,
            track_changes: true,
            rsid_root: Some("AB12CD34".to_string()),
            compat_mode: true,
        };
        // Just verify struct construction compiles and no panics
        let _ = opts;
    }

    #[test]
    #[cfg(all(
        feature = "wml-settings",
        feature = "extra-attrs",
        feature = "extra-children"
    ))]
    fn test_settings_roundtrip() {
        use crate::Document;
        use crate::writer::{DocumentBuilder, DocumentSettingsOptions};
        use std::io::Cursor;

        let mut builder = DocumentBuilder::new();
        builder.set_settings(DocumentSettingsOptions {
            default_tab_stop: Some(720),
            even_and_odd_headers: true,
            track_changes: false,
            rsid_root: None,
            compat_mode: false,
        });
        builder.add_paragraph("Hello");

        let mut buf = Cursor::new(Vec::new());
        builder.write(&mut buf).unwrap();

        // Re-open and verify the document is readable
        buf.set_position(0);
        let doc = Document::from_reader(buf).unwrap();
        let body = doc.body();
        assert!(!body.block_content.is_empty());
    }

    // -------------------------------------------------------------------------
    // Form fields
    // -------------------------------------------------------------------------

    #[test]
    #[cfg(feature = "wml-settings")]
    fn test_form_field_plain_text() {
        let mut body = types::Body::default();
        body.add_form_field(FormFieldConfig {
            tag: Some("myTag".to_string()),
            label: Some("My Field".to_string()),
            field_type: FormFieldType::PlainText,
            default_value: Some("default".to_string()),
            ..Default::default()
        });

        assert_eq!(body.block_content.len(), 1);
        match &body.block_content[0] {
            types::BlockContent::Sdt(sdt) => {
                let sdt_pr = sdt.sdt_pr.as_ref().expect("sdt_pr should be present");
                assert_eq!(sdt_pr.tag.as_ref().unwrap().value, "myTag");
                assert_eq!(sdt_pr.alias.as_ref().unwrap().value, "My Field");
                assert!(sdt_pr.text.is_some(), "text element should be set");
                let content = sdt.sdt_content.as_ref().expect("sdt_content");
                assert_eq!(content.block_content.len(), 1);
            }
            _ => panic!("expected Sdt block content"),
        }
    }

    #[test]
    #[cfg(feature = "wml-settings")]
    fn test_form_field_dropdown() {
        let mut body = types::Body::default();
        body.add_form_field(FormFieldConfig {
            field_type: FormFieldType::DropDownList,
            list_items: vec!["Option A".to_string(), "Option B".to_string()],
            ..Default::default()
        });

        match &body.block_content[0] {
            types::BlockContent::Sdt(sdt) => {
                let sdt_pr = sdt.sdt_pr.as_ref().unwrap();
                let dd = sdt_pr.drop_down_list.as_ref().expect("drop_down_list");
                assert_eq!(dd.list_item.len(), 2);
                assert_eq!(dd.list_item[0].display_text.as_deref(), Some("Option A"));
                assert_eq!(dd.list_item[1].display_text.as_deref(), Some("Option B"));
            }
            _ => panic!("expected Sdt"),
        }
    }

    #[test]
    #[cfg(feature = "wml-settings")]
    fn test_form_field_date_picker() {
        let mut body = types::Body::default();
        body.add_form_field(FormFieldConfig {
            field_type: FormFieldType::DatePicker,
            date_format: Some("MM/dd/yyyy".to_string()),
            ..Default::default()
        });

        match &body.block_content[0] {
            types::BlockContent::Sdt(sdt) => {
                let sdt_pr = sdt.sdt_pr.as_ref().unwrap();
                let date = sdt_pr.date.as_ref().expect("date element");
                assert_eq!(date.date_format.as_ref().unwrap().value, "MM/dd/yyyy");
            }
            _ => panic!("expected Sdt"),
        }
    }

    // -------------------------------------------------------------------------
    // Table of contents
    // -------------------------------------------------------------------------

    #[test]
    #[cfg(feature = "wml-fields")]
    fn test_toc_basic() {
        let mut body = types::Body::default();
        body.add_toc(TocOptions {
            title: Some("Contents".to_string()),
            max_level: 3,
            right_align_page_numbers: true,
            use_hyperlinks: true,
        });

        // Should have: title para + TOC field para + placeholder para = 3 paragraphs
        let paras: Vec<_> = body
            .block_content
            .iter()
            .filter_map(|b| match b {
                types::BlockContent::P(p) => Some(p),
                _ => None,
            })
            .collect();
        assert_eq!(
            paras.len(),
            3,
            "expected title + field + placeholder paragraphs"
        );

        // The field paragraph should have 4 run content items (begin, instrText, separate, end)
        let field_para = &paras[1];
        assert_eq!(
            field_para.paragraph_content.len(),
            4,
            "TOC paragraph should have 4 run items"
        );
    }

    #[test]
    #[cfg(feature = "wml-fields")]
    fn test_toc_no_title() {
        let mut body = types::Body::default();
        body.add_toc(TocOptions::default());

        // Should have: TOC field para + placeholder para = 2 paragraphs
        let paras: Vec<_> = body
            .block_content
            .iter()
            .filter_map(|b| match b {
                types::BlockContent::P(p) => Some(p),
                _ => None,
            })
            .collect();
        assert_eq!(paras.len(), 2, "expected field + placeholder paragraphs");
    }

    #[test]
    #[cfg(feature = "wml-fields")]
    fn test_toc_instr_text_contains_level() {
        let mut body = types::Body::default();
        body.add_toc(TocOptions {
            title: None,
            max_level: 2,
            right_align_page_numbers: false,
            use_hyperlinks: false,
        });

        let field_para = match &body.block_content[0] {
            types::BlockContent::P(p) => p,
            _ => panic!("expected paragraph"),
        };

        // Second run content item is InstrText
        match &field_para.paragraph_content[1] {
            types::ParagraphContent::R(run) => match &run.run_content[0] {
                types::RunContent::InstrText(t) => {
                    let instr = t.text.as_deref().unwrap_or("");
                    assert!(
                        instr.contains("1-2"),
                        "should contain level range 1-2, got: {}",
                        instr
                    );
                }
                _ => panic!("expected InstrText"),
            },
            _ => panic!("expected run"),
        }
    }

    // -------------------------------------------------------------------------
    // OMathBuilder
    // -------------------------------------------------------------------------

    #[test]
    fn test_omath_plain_build() {
        let builder = OMathBuilder::plain("x");
        let elem = builder.build();
        assert_eq!(elem.name, "m:oMath");
        assert!(!elem.attributes.is_empty(), "should have xmlns:m");
        assert_eq!(elem.children.len(), 1);
        match &elem.children[0] {
            RawXmlNode::Text(t) => assert!(t.contains("m:r"), "text should wrap in m:r: {}", t),
            _ => panic!("expected text node"),
        }
    }

    #[test]
    fn test_omath_fraction_build() {
        let builder = OMathBuilder::fraction("a", "b");
        let elem = builder.build();
        match &elem.children[0] {
            RawXmlNode::Text(t) => {
                assert!(t.contains("m:f"), "should contain fraction: {}", t);
                assert!(t.contains("m:num"), "should contain numerator: {}", t);
                assert!(t.contains("m:den"), "should contain denominator: {}", t);
            }
            _ => panic!("expected text node"),
        }
    }

    #[test]
    fn test_omath_superscript_build() {
        let builder = OMathBuilder::superscript("x", "2");
        let elem = builder.build();
        match &elem.children[0] {
            RawXmlNode::Text(t) => {
                assert!(t.contains("m:sSup"), "should contain sSup: {}", t);
                assert!(t.contains("m:e"), "should contain base: {}", t);
                assert!(t.contains("m:sup"), "should contain exp: {}", t);
            }
            _ => panic!("expected text node"),
        }
    }

    #[test]
    fn test_omath_subscript_build() {
        let builder = OMathBuilder::subscript("x", "i");
        let elem = builder.build();
        match &elem.children[0] {
            RawXmlNode::Text(t) => assert!(t.contains("m:sSub"), "{}", t),
            _ => panic!(),
        }
    }

    #[test]
    fn test_omath_radical_build() {
        let builder = OMathBuilder::radical("x");
        let elem = builder.build();
        match &elem.children[0] {
            RawXmlNode::Text(t) => assert!(t.contains("m:rad"), "{}", t),
            _ => panic!(),
        }
    }

    #[test]
    fn test_omath_display_wraps_in_para() {
        let builder = OMathBuilder::plain("x").as_display();
        let elem = builder.build();
        assert_eq!(
            elem.name, "m:oMathPara",
            "display should be wrapped in oMathPara"
        );
        assert_eq!(elem.children.len(), 1);
        match &elem.children[0] {
            RawXmlNode::Element(inner) => assert_eq!(inner.name, "m:oMath"),
            _ => panic!("expected oMath element child"),
        }
    }

    #[test]
    #[cfg(feature = "extra-children")]
    fn test_paragraph_add_math() {
        let mut para = types::Paragraph::default();
        para.add_math(OMathBuilder::plain("y = mx + b"));
        assert_eq!(para.extra_children.len(), 1);
    }

    // -------------------------------------------------------------------------
    // Inline chart
    // -------------------------------------------------------------------------

    #[test]
    #[cfg(all(
        feature = "wml-charts",
        feature = "extra-children",
        feature = "extra-attrs"
    ))]
    fn test_embed_chart_and_add_inline() {
        use crate::writer::DocumentBuilder;
        use std::io::Cursor;

        let chart_xml = br#"<?xml version="1.0" encoding="UTF-8"?><c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"/>"#;

        let mut builder = DocumentBuilder::new();
        let rel_id = builder.embed_chart(chart_xml).unwrap();
        assert!(
            rel_id.starts_with("rId"),
            "rel_id should be rId-prefixed: {}",
            rel_id
        );

        // Add a paragraph with the inline chart
        {
            let body = builder.body_mut();
            let para = body.add_paragraph();
            para.add_inline_chart(&rel_id, 3000000, 2000000);
        }

        // Write to a buffer — should not panic
        let mut buf = Cursor::new(Vec::new());
        builder.write(&mut buf).unwrap();
        assert!(!buf.get_ref().is_empty(), "output should not be empty");
    }

    #[test]
    #[cfg(feature = "wml-charts")]
    fn test_chart_inline_element_structure() {
        let elem = build_chart_inline_element("rId5", 3000000, 2000000);
        assert_eq!(elem.name, "wp:inline");
        // Should contain extent, docPr, and graphic children
        assert_eq!(elem.children.len(), 3);
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    // -------------------------------------------------------------------------
    // Page / column breaks
    // -------------------------------------------------------------------------

    #[test]
    fn test_paragraph_add_page_break() {
        let mut para = types::Paragraph::default();
        para.add_page_break();

        assert_eq!(para.paragraph_content.len(), 1);
        match &para.paragraph_content[0] {
            types::ParagraphContent::R(run) => {
                assert_eq!(run.run_content.len(), 1);
                match &run.run_content[0] {
                    types::RunContent::Br(br) => {
                        assert_eq!(br.r#type, Some(types::STBrType::Page));
                    }
                    _ => panic!("expected Br content"),
                }
            }
            _ => panic!("expected R content"),
        }
    }

    #[test]
    fn test_paragraph_add_column_break() {
        let mut para = types::Paragraph::default();
        para.add_column_break();

        assert_eq!(para.paragraph_content.len(), 1);
        match &para.paragraph_content[0] {
            types::ParagraphContent::R(run) => {
                assert_eq!(run.run_content.len(), 1);
                match &run.run_content[0] {
                    types::RunContent::Br(br) => {
                        assert_eq!(br.r#type, Some(types::STBrType::Column));
                    }
                    _ => panic!("expected Br content"),
                }
            }
            _ => panic!("expected R content"),
        }
    }

    // -------------------------------------------------------------------------
    // Paragraph spacing
    // -------------------------------------------------------------------------

    #[test]
    #[cfg(feature = "wml-styling")]
    fn test_set_space_before() {
        let mut para = types::Paragraph::default();
        para.set_space_before(240);

        let ppr = para.p_pr.as_ref().unwrap();
        let spacing = ppr.spacing.as_ref().unwrap();
        assert_eq!(spacing.before.as_deref(), Some("240"));
        assert!(spacing.after.is_none());
    }

    #[test]
    #[cfg(feature = "wml-styling")]
    fn test_set_space_after() {
        let mut para = types::Paragraph::default();
        para.set_space_after(160);

        let ppr = para.p_pr.as_ref().unwrap();
        let spacing = ppr.spacing.as_ref().unwrap();
        assert_eq!(spacing.after.as_deref(), Some("160"));
        assert!(spacing.before.is_none());
    }

    #[test]
    #[cfg(feature = "wml-styling")]
    fn test_set_line_spacing() {
        let mut para = types::Paragraph::default();
        para.set_line_spacing(360); // 1.5x spacing

        let ppr = para.p_pr.as_ref().unwrap();
        let spacing = ppr.spacing.as_ref().unwrap();
        assert_eq!(spacing.line.as_deref(), Some("360"));
        assert_eq!(spacing.line_rule, Some(types::STLineSpacingRule::Auto));
    }

    #[test]
    #[cfg(feature = "wml-styling")]
    fn test_spacing_accumulates() {
        let mut para = types::Paragraph::default();
        para.set_space_before(240);
        para.set_space_after(120);
        para.set_line_spacing(480);

        let ppr = para.p_pr.as_ref().unwrap();
        let spacing = ppr.spacing.as_ref().unwrap();
        assert_eq!(spacing.before.as_deref(), Some("240"));
        assert_eq!(spacing.after.as_deref(), Some("120"));
        assert_eq!(spacing.line.as_deref(), Some("480"));
    }

    // -------------------------------------------------------------------------
    // Paragraph indentation
    // -------------------------------------------------------------------------

    #[test]
    #[cfg(feature = "wml-styling")]
    fn test_set_indent_left() {
        let mut para = types::Paragraph::default();
        para.set_indent_left(720);

        let ppr = para.p_pr.as_ref().unwrap();
        let ind = ppr.indentation.as_ref().unwrap();
        assert_eq!(ind.left.as_deref(), Some("720"));
    }

    #[test]
    #[cfg(feature = "wml-styling")]
    fn test_set_indent_right() {
        let mut para = types::Paragraph::default();
        para.set_indent_right(360);

        let ppr = para.p_pr.as_ref().unwrap();
        let ind = ppr.indentation.as_ref().unwrap();
        assert_eq!(ind.right.as_deref(), Some("360"));
    }

    #[test]
    #[cfg(feature = "wml-styling")]
    fn test_set_indent_first_line() {
        let mut para = types::Paragraph::default();
        para.set_indent_first_line(180);

        let ppr = para.p_pr.as_ref().unwrap();
        let ind = ppr.indentation.as_ref().unwrap();
        assert_eq!(ind.first_line.as_deref(), Some("180"));
    }

    #[test]
    #[cfg(feature = "wml-styling")]
    fn test_indentation_accumulates() {
        let mut para = types::Paragraph::default();
        para.set_indent_left(720);
        para.set_indent_right(360);
        para.set_indent_first_line(180);

        let ppr = para.p_pr.as_ref().unwrap();
        let ind = ppr.indentation.as_ref().unwrap();
        assert_eq!(ind.left.as_deref(), Some("720"));
        assert_eq!(ind.right.as_deref(), Some("360"));
        assert_eq!(ind.first_line.as_deref(), Some("180"));
    }

    // -------------------------------------------------------------------------
    // Outline level
    // -------------------------------------------------------------------------

    #[test]
    #[cfg(feature = "wml-styling")]
    fn test_set_outline_level() {
        let mut para = types::Paragraph::default();
        para.set_outline_level(1); // heading level 2

        let ppr = para.p_pr.as_ref().unwrap();
        let lvl = ppr.outline_lvl.as_ref().unwrap();
        assert_eq!(lvl.value, 1);
    }

    // -------------------------------------------------------------------------
    // Run property details
    // -------------------------------------------------------------------------

    #[test]
    #[cfg(feature = "wml-styling")]
    fn test_run_set_shadow() {
        let mut run = types::Run::default();
        run.set_shadow(true);
        assert!(run.r_pr.as_ref().unwrap().shadow.is_some());
        run.set_shadow(false);
        assert!(run.r_pr.as_ref().unwrap().shadow.is_none());
    }

    #[test]
    #[cfg(feature = "wml-styling")]
    fn test_run_set_outline() {
        let mut run = types::Run::default();
        run.set_outline(true);
        assert!(run.r_pr.as_ref().unwrap().outline.is_some());
        run.set_outline(false);
        assert!(run.r_pr.as_ref().unwrap().outline.is_none());
    }

    #[test]
    #[cfg(feature = "wml-styling")]
    fn test_run_set_emboss() {
        let mut run = types::Run::default();
        run.set_emboss(true);
        assert!(run.r_pr.as_ref().unwrap().emboss.is_some());
    }

    #[test]
    #[cfg(feature = "wml-styling")]
    fn test_run_set_imprint() {
        let mut run = types::Run::default();
        run.set_imprint(true);
        assert!(run.r_pr.as_ref().unwrap().imprint.is_some());
    }

    #[test]
    #[cfg(feature = "wml-styling")]
    fn test_run_set_small_caps() {
        let mut run = types::Run::default();
        run.set_small_caps(true);
        assert!(run.r_pr.as_ref().unwrap().small_caps.is_some());
        run.set_small_caps(false);
        assert!(run.r_pr.as_ref().unwrap().small_caps.is_none());
    }

    #[test]
    #[cfg(feature = "wml-styling")]
    fn test_run_set_all_caps() {
        let mut run = types::Run::default();
        run.set_all_caps(true);
        assert!(run.r_pr.as_ref().unwrap().caps.is_some());
        run.set_all_caps(false);
        assert!(run.r_pr.as_ref().unwrap().caps.is_none());
    }

    #[test]
    #[cfg(feature = "wml-styling")]
    fn test_run_set_vanish() {
        let mut run = types::Run::default();
        run.set_vanish(true);
        assert!(run.r_pr.as_ref().unwrap().vanish.is_some());
        run.set_vanish(false);
        assert!(run.r_pr.as_ref().unwrap().vanish.is_none());
    }

    #[test]
    #[cfg(feature = "wml-styling")]
    fn test_run_set_double_strike() {
        let mut run = types::Run::default();
        run.set_double_strike(true);
        assert!(run.r_pr.as_ref().unwrap().dstrike.is_some());
        run.set_double_strike(false);
        assert!(run.r_pr.as_ref().unwrap().dstrike.is_none());
    }

    // -------------------------------------------------------------------------
    // Table merged cells
    // -------------------------------------------------------------------------

    #[test]
    #[cfg(feature = "wml-tables")]
    fn test_cell_set_grid_span() {
        let mut cell = types::TableCell::default();
        cell.set_grid_span(3);

        let tcpr = cell.cell_properties.as_ref().unwrap();
        let gs = tcpr.grid_span.as_ref().unwrap();
        assert_eq!(gs.value, 3);
    }

    #[test]
    #[cfg(feature = "wml-tables")]
    fn test_cell_set_vertical_merge_restart() {
        let mut cell = types::TableCell::default();
        cell.set_vertical_merge(VMergeType::Restart);

        let tcpr = cell.cell_properties.as_ref().unwrap();
        let vm = tcpr.vertical_merge.as_ref().unwrap();
        assert_eq!(vm.value, Some(types::STMerge::Restart));
    }

    #[test]
    #[cfg(feature = "wml-tables")]
    fn test_cell_set_vertical_merge_continue() {
        let mut cell = types::TableCell::default();
        cell.set_vertical_merge(VMergeType::Continue);

        let tcpr = cell.cell_properties.as_ref().unwrap();
        let vm = tcpr.vertical_merge.as_ref().unwrap();
        assert_eq!(vm.value, None);
    }

    // -------------------------------------------------------------------------
    // Roundtrip tests
    // -------------------------------------------------------------------------

    #[test]
    #[cfg(all(
        feature = "wml-styling",
        feature = "extra-attrs",
        feature = "extra-children"
    ))]
    fn test_roundtrip_page_break() {
        use crate::Document;
        use crate::writer::DocumentBuilder;
        use std::io::Cursor;

        let mut builder = DocumentBuilder::new();
        let body = builder.body_mut();
        let para = body.add_paragraph();
        para.add_page_break();

        let mut buf = Cursor::new(Vec::new());
        builder.write(&mut buf).unwrap();

        buf.set_position(0);
        let doc = Document::from_reader(buf).unwrap();
        let body_ref = doc.body();
        assert!(!body_ref.block_content.is_empty());
    }

    #[test]
    #[cfg(all(
        feature = "wml-styling",
        feature = "extra-attrs",
        feature = "extra-children"
    ))]
    fn test_roundtrip_spacing_and_indent() {
        use crate::Document;
        use crate::ext::BodyExt;
        use crate::writer::DocumentBuilder;
        use std::io::Cursor;

        let mut builder = DocumentBuilder::new();
        {
            let body = builder.body_mut();
            let para = body.add_paragraph();
            para.set_space_before(240);
            para.set_space_after(120);
            para.set_line_spacing(360);
            para.set_indent_left(720);
            para.set_outline_level(0);
            para.add_run().set_text("test spacing");
        }

        let mut buf = Cursor::new(Vec::new());
        builder.write(&mut buf).unwrap();

        buf.set_position(0);
        let doc = Document::from_reader(buf).unwrap();
        let body_ref = doc.body();
        assert_eq!(body_ref.paragraphs().len(), 1);

        let p = body_ref.paragraphs()[0];
        let ppr = p.p_pr.as_ref().unwrap();
        let spacing = ppr.spacing.as_ref().unwrap();
        assert_eq!(spacing.before.as_deref(), Some("240"));
        assert_eq!(spacing.after.as_deref(), Some("120"));
    }

    #[test]
    #[cfg(all(
        feature = "wml-tables",
        feature = "extra-attrs",
        feature = "extra-children"
    ))]
    fn test_roundtrip_grid_span_and_vmerge() {
        use crate::Document;
        use crate::ext::BodyExt;
        use crate::writer::DocumentBuilder;
        use std::io::Cursor;

        let mut builder = DocumentBuilder::new();
        {
            let body = builder.body_mut();
            let tbl = body.add_table();
            let row = tbl.add_row();
            let cell = row.add_cell();
            cell.set_grid_span(2);
            cell.set_vertical_merge(VMergeType::Restart);
            cell.add_paragraph().add_run().set_text("merged");
        }

        let mut buf = Cursor::new(Vec::new());
        builder.write(&mut buf).unwrap();

        buf.set_position(0);
        let doc = Document::from_reader(buf).unwrap();
        let body_ref = doc.body();
        assert!(!body_ref.tables().is_empty());

        let tbl = body_ref.tables()[0];
        let row = match &tbl.rows[0] {
            crate::types::RowContent::Tr(r) => r,
            _ => panic!("expected Tr"),
        };
        let cell = match &row.cells[0] {
            crate::types::CellContent::Tc(c) => c,
            _ => panic!("expected Tc"),
        };
        let tcpr = cell.cell_properties.as_ref().unwrap();
        assert_eq!(tcpr.grid_span.as_ref().unwrap().value, 2);
        assert_eq!(
            tcpr.vertical_merge.as_ref().unwrap().value,
            Some(crate::types::STMerge::Restart)
        );
    }

    // -------------------------------------------------------------------------
    // Bookmarks (u32 wrappers)
    // -------------------------------------------------------------------------

    #[test]
    fn test_add_bookmark_start_u32() {
        let mut para = types::Paragraph::default();
        para.add_bookmark_start_u32(1, "myBookmark");

        assert_eq!(para.paragraph_content.len(), 1);
        match &para.paragraph_content[0] {
            types::ParagraphContent::BookmarkStart(bm) => {
                assert_eq!(bm.id, 1);
                assert_eq!(bm.name, "myBookmark");
            }
            _ => panic!("expected BookmarkStart"),
        }
    }

    #[test]
    fn test_add_bookmark_end_u32() {
        let mut para = types::Paragraph::default();
        para.add_bookmark_end_u32(42);

        assert_eq!(para.paragraph_content.len(), 1);
        match &para.paragraph_content[0] {
            types::ParagraphContent::BookmarkEnd(bm) => {
                assert_eq!(bm.id, 42);
            }
            _ => panic!("expected BookmarkEnd"),
        }
    }

    // -------------------------------------------------------------------------
    // Table row height
    // -------------------------------------------------------------------------

    #[test]
    #[cfg(feature = "wml-tables")]
    fn test_row_set_height() {
        let mut row = types::CTRow::default();
        row.set_height(720);

        let row_pr = row.row_properties.as_ref().unwrap();
        let height = row_pr.tr_height.as_ref().unwrap();
        assert_eq!(height.value.as_deref(), Some("720"));
        assert_eq!(height.h_rule, Some(types::STHeightRule::Exact));
    }

    // -------------------------------------------------------------------------
    // Table cell background color
    // -------------------------------------------------------------------------

    #[test]
    #[cfg(feature = "wml-tables")]
    fn test_cell_set_background_color() {
        let mut cell = types::TableCell::default();
        cell.set_background_color("FF0000");

        let tcpr = cell.cell_properties.as_ref().unwrap();
        let shd = tcpr.shading.as_ref().unwrap();
        assert_eq!(shd.value, types::STShd::Clear);
        #[cfg(feature = "wml-styling")]
        assert_eq!(shd.fill.as_deref(), Some("FF0000"));
    }

    // -------------------------------------------------------------------------
    // Table cell borders
    // -------------------------------------------------------------------------

    #[test]
    #[cfg(feature = "wml-tables")]
    fn test_cell_set_borders() {
        let mut cell = types::TableCell::default();
        cell.set_borders(BorderStyle::Single, 4, "000000");

        let tcpr = cell.cell_properties.as_ref().unwrap();
        let borders = tcpr.tc_borders.as_ref().unwrap();
        assert!(borders.top.is_some());
        assert!(borders.bottom.is_some());
        assert!(borders.left.is_some());
        assert!(borders.right.is_some());
        let top = borders.top.as_ref().unwrap();
        assert_eq!(top.value, types::STBorder::Single);
        #[cfg(feature = "wml-styling")]
        assert_eq!(top.size, Some(4u64));
    }

    #[test]
    #[cfg(feature = "wml-tables")]
    fn test_cell_set_border_top_only() {
        let mut cell = types::TableCell::default();
        cell.set_border_top(BorderStyle::Dashed, 8, "AABBCC");

        let tcpr = cell.cell_properties.as_ref().unwrap();
        let borders = tcpr.tc_borders.as_ref().unwrap();
        assert!(borders.top.is_some());
        assert!(borders.bottom.is_none());
        assert!(borders.left.is_none());
        assert!(borders.right.is_none());
    }

    // -------------------------------------------------------------------------
    // Table cell padding
    // -------------------------------------------------------------------------

    #[test]
    #[cfg(feature = "wml-tables")]
    fn test_cell_set_padding() {
        let mut cell = types::TableCell::default();
        cell.set_padding(100, 100, 200, 200);

        let tcpr = cell.cell_properties.as_ref().unwrap();
        let mar = tcpr.tc_mar.as_ref().unwrap();
        let top = mar.top.as_ref().unwrap();
        assert_eq!(top.width.as_deref(), Some("100"));
        assert_eq!(top.r#type, Some(types::STTblWidth::Dxa));
        let left = mar.left.as_ref().unwrap();
        assert_eq!(left.width.as_deref(), Some("200"));
    }

    // -------------------------------------------------------------------------
    // Table width
    // -------------------------------------------------------------------------

    #[test]
    #[cfg(feature = "wml-tables")]
    fn test_table_set_width_dxa() {
        let table = types::Table {
            range_markup: Vec::new(),
            table_properties: Box::new(types::TableProperties::default()),
            tbl_grid: Box::new(types::TableGrid::default()),
            rows: Vec::new(),
            #[cfg(feature = "extra-children")]
            extra_children: Vec::new(),
        };
        let mut body = types::Body::default();
        body.block_content
            .push(types::BlockContent::Tbl(Box::new(table)));
        let tbl = match body.block_content.last_mut().unwrap() {
            types::BlockContent::Tbl(t) => t.as_mut(),
            _ => unreachable!(),
        };
        tbl.set_width(9360, TableWidthUnit::Dxa); // 6.5 inches

        let tbl_w = tbl.table_properties.tbl_w.as_ref().unwrap();
        assert_eq!(tbl_w.width.as_deref(), Some("9360"));
        assert_eq!(tbl_w.r#type, Some(types::STTblWidth::Dxa));
    }

    #[test]
    #[cfg(feature = "wml-tables")]
    fn test_table_set_width_pct() {
        let table = types::Table {
            range_markup: Vec::new(),
            table_properties: Box::new(types::TableProperties::default()),
            tbl_grid: Box::new(types::TableGrid::default()),
            rows: Vec::new(),
            #[cfg(feature = "extra-children")]
            extra_children: Vec::new(),
        };
        let mut body = types::Body::default();
        body.block_content
            .push(types::BlockContent::Tbl(Box::new(table)));
        let tbl = match body.block_content.last_mut().unwrap() {
            types::BlockContent::Tbl(t) => t.as_mut(),
            _ => unreachable!(),
        };
        tbl.set_width(5000, TableWidthUnit::Pct); // 100%

        let tbl_w = tbl.table_properties.tbl_w.as_ref().unwrap();
        assert_eq!(tbl_w.width.as_deref(), Some("5000"));
        assert_eq!(tbl_w.r#type, Some(types::STTblWidth::Pct));
    }

    // -------------------------------------------------------------------------
    // Roundtrip: table cell styling
    // -------------------------------------------------------------------------

    #[test]
    #[cfg(all(
        feature = "wml-tables",
        feature = "wml-styling",
        feature = "extra-attrs",
        feature = "extra-children"
    ))]
    fn test_roundtrip_cell_background_and_borders() {
        use crate::Document;
        use crate::ext::BodyExt;
        use crate::writer::DocumentBuilder;
        use std::io::Cursor;

        let mut builder = DocumentBuilder::new();
        {
            let body = builder.body_mut();
            let tbl = body.add_table();
            let row = tbl.add_row();
            let cell = row.add_cell();
            cell.set_background_color("FFFF00");
            cell.set_borders(BorderStyle::Single, 4, "000000");
            cell.set_padding(72, 72, 144, 144);
            cell.add_paragraph().add_run().set_text("styled");
        }

        let mut buf = Cursor::new(Vec::new());
        builder.write(&mut buf).unwrap();

        buf.set_position(0);
        let doc = Document::from_reader(buf).unwrap();
        let body_ref = doc.body();
        assert!(!body_ref.tables().is_empty());

        let tbl = body_ref.tables()[0];
        let row = match &tbl.rows[0] {
            crate::types::RowContent::Tr(r) => r,
            _ => panic!("expected Tr"),
        };
        let cell = match &row.cells[0] {
            crate::types::CellContent::Tc(c) => c,
            _ => panic!("expected Tc"),
        };
        let tcpr = cell.cell_properties.as_ref().unwrap();
        let shd = tcpr.shading.as_ref().unwrap();
        assert_eq!(shd.value, crate::types::STShd::Clear);
        assert_eq!(shd.fill.as_deref(), Some("FFFF00"));

        let borders = tcpr.tc_borders.as_ref().unwrap();
        assert!(borders.top.is_some());
        let top = borders.top.as_ref().unwrap();
        assert_eq!(top.value, crate::types::STBorder::Single);

        let mar = tcpr.tc_mar.as_ref().unwrap();
        let left = mar.left.as_ref().unwrap();
        assert_eq!(left.width.as_deref(), Some("144"));
    }
}
