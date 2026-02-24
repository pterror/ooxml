//! Builder methods on generated types for ergonomic document construction.
//!
//! These `impl` blocks extend the generated types with convenience methods
//! for building documents (adding paragraphs, runs, tables, etc.).
//! They are allowed because the generated types are in the same crate.

use crate::types;

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
}
