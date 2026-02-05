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
            displaced_by_custom_xml: None,
            col_first: None,
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
        ppr.num_pr = Some(Box::new(types::CTNumPr {
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
            rpr.bold = Some(Box::new(types::CTOnOff {
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
            rpr.italic = Some(Box::new(types::CTOnOff {
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
        rpr.size = Some(Box::new(types::CTHpsMeasure {
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
            rpr.strikethrough = Some(Box::new(types::CTOnOff {
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
