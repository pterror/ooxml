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
        self.block_level_elts
            .push(Box::new(types::EGBlockLevelElts::P(Box::default())));
        match self.block_level_elts.last_mut().unwrap().as_mut() {
            types::EGBlockLevelElts::P(p) => p.as_mut(),
            _ => unreachable!(),
        }
    }

    /// Add an empty table and return a mutable reference to it.
    #[cfg(feature = "wml-tables")]
    pub fn add_table(&mut self) -> &mut types::Table {
        let table = types::Table {
            range_markup_elements: Vec::new(),
            table_properties: Box::new(types::TableProperties::default()),
            tbl_grid: Box::new(types::TableGrid::default()),
            content_row_content: Vec::new(),
            #[cfg(feature = "extra-children")]
            extra_children: Vec::new(),
        };
        self.block_level_elts
            .push(Box::new(types::EGBlockLevelElts::Tbl(Box::new(table))));
        match self.block_level_elts.last_mut().unwrap().as_mut() {
            types::EGBlockLevelElts::Tbl(t) => t.as_mut(),
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
        self.p_content
            .push(Box::new(types::EGPContent::R(Box::default())));
        match self.p_content.last_mut().unwrap().as_mut() {
            types::EGPContent::R(r) => r.as_mut(),
            _ => unreachable!(),
        }
    }

    /// Add an empty hyperlink and return a mutable reference to it.
    #[cfg(feature = "wml-hyperlinks")]
    pub fn add_hyperlink(&mut self) -> &mut types::Hyperlink {
        self.p_content
            .push(Box::new(types::EGPContent::Hyperlink(Box::default())));
        match self.p_content.last_mut().unwrap().as_mut() {
            types::EGPContent::Hyperlink(h) => h.as_mut(),
            _ => unreachable!(),
        }
    }

    /// Add a bookmark start marker.
    pub fn add_bookmark_start(&mut self, id: i64, name: &str) {
        let mut bookmark = types::Bookmark {
            name: name.to_string(),
            #[cfg(feature = "extra-attrs")]
            extra_attrs: std::collections::HashMap::new(),
        };
        // id is an inherited attr from CT_MarkupRange → CT_Markup, stored in extra_attrs
        #[cfg(feature = "extra-attrs")]
        {
            bookmark
                .extra_attrs
                .insert("w:id".to_string(), id.to_string());
        }
        let _ = id; // suppress unused warning when extra-attrs is off
        self.p_content
            .push(Box::new(types::EGPContent::BookmarkStart(Box::new(
                bookmark,
            ))));
    }

    /// Add a bookmark end marker.
    pub fn add_bookmark_end(&mut self, id: i64) {
        let mut range = types::CTMarkupRange {
            displaced_by_custom_xml: None,
            #[cfg(feature = "extra-attrs")]
            extra_attrs: std::collections::HashMap::new(),
        };
        // id is inherited from CT_Markup, stored in extra_attrs
        #[cfg(feature = "extra-attrs")]
        {
            range.extra_attrs.insert("w:id".to_string(), id.to_string());
        }
        let _ = id;
        self.p_content
            .push(Box::new(types::EGPContent::BookmarkEnd(Box::new(range))));
    }

    /// Add a comment range start marker.
    pub fn add_comment_range_start(&mut self, id: u32) {
        let mut range = types::CTMarkupRange {
            displaced_by_custom_xml: None,
            #[cfg(feature = "extra-attrs")]
            extra_attrs: std::collections::HashMap::new(),
        };
        #[cfg(feature = "extra-attrs")]
        {
            range.extra_attrs.insert("w:id".to_string(), id.to_string());
        }
        let _ = id;
        self.p_content
            .push(Box::new(types::EGPContent::CommentRangeStart(Box::new(
                range,
            ))));
    }

    /// Add a comment range end marker.
    pub fn add_comment_range_end(&mut self, id: u32) {
        let mut range = types::CTMarkupRange {
            displaced_by_custom_xml: None,
            #[cfg(feature = "extra-attrs")]
            extra_attrs: std::collections::HashMap::new(),
        };
        #[cfg(feature = "extra-attrs")]
        {
            range.extra_attrs.insert("w:id".to_string(), id.to_string());
        }
        let _ = id;
        self.p_content
            .push(Box::new(types::EGPContent::CommentRangeEnd(Box::new(
                range,
            ))));
    }

    /// Set paragraph properties.
    #[cfg(feature = "wml-styling")]
    pub fn set_properties(&mut self, props: types::ParagraphProperties) {
        self.p_pr = Some(Box::new(props));
    }

    /// Set numbering properties (list membership) on this paragraph.
    ///
    /// This creates a `w:numPr` element with `w:ilvl` and `w:numId` children
    /// inside the paragraph properties' extra_children, since the generated
    /// ParagraphProperties doesn't yet flatten CTPPrBase fields.
    #[cfg(all(feature = "wml-styling", feature = "extra-children"))]
    pub fn set_numbering(&mut self, num_id: u32, ilvl: u32) {
        use ooxml_xml::{RawXmlElement, RawXmlNode};

        let ppr = self
            .p_pr
            .get_or_insert_with(|| Box::new(types::ParagraphProperties::default()));
        let num_pr = RawXmlElement {
            name: "w:numPr".to_string(),
            attributes: vec![],
            children: vec![
                RawXmlNode::Element(RawXmlElement {
                    name: "w:ilvl".to_string(),
                    attributes: vec![("w:val".to_string(), ilvl.to_string())],
                    children: vec![],
                    self_closing: true,
                }),
                RawXmlNode::Element(RawXmlElement {
                    name: "w:numId".to_string(),
                    attributes: vec![("w:val".to_string(), num_id.to_string())],
                    children: vec![],
                    self_closing: true,
                }),
            ],
            self_closing: false,
        };
        ppr.extra_children.push(RawXmlNode::Element(num_pr));
    }

    /// Set paragraph alignment.
    ///
    /// Creates a `w:jc` element in the paragraph properties' extra_children.
    /// Common values: "left", "center", "right", "both" (justified).
    #[cfg(all(feature = "wml-styling", feature = "extra-children"))]
    pub fn set_alignment(&mut self, alignment: &str) {
        use ooxml_xml::{RawXmlElement, RawXmlNode};

        let ppr = self
            .p_pr
            .get_or_insert_with(|| Box::new(types::ParagraphProperties::default()));
        let jc = RawXmlElement {
            name: "w:jc".to_string(),
            attributes: vec![("w:val".to_string(), alignment.to_string())],
            children: vec![],
            self_closing: true,
        };
        ppr.extra_children.push(RawXmlNode::Element(jc));
    }

    /// Set paragraph spacing (before and after, in twips).
    ///
    /// Creates a `w:spacing` element in the paragraph properties' extra_children.
    #[cfg(all(feature = "wml-styling", feature = "extra-children"))]
    pub fn set_spacing(&mut self, before: Option<u32>, after: Option<u32>) {
        use ooxml_xml::{RawXmlElement, RawXmlNode};

        let ppr = self
            .p_pr
            .get_or_insert_with(|| Box::new(types::ParagraphProperties::default()));
        let mut attrs = vec![];
        if let Some(b) = before {
            attrs.push(("w:before".to_string(), b.to_string()));
        }
        if let Some(a) = after {
            attrs.push(("w:after".to_string(), a.to_string()));
        }
        let spacing = RawXmlElement {
            name: "w:spacing".to_string(),
            attributes: attrs,
            children: vec![],
            self_closing: true,
        };
        ppr.extra_children.push(RawXmlNode::Element(spacing));
    }

    /// Set paragraph indentation.
    ///
    /// Creates a `w:ind` element in the paragraph properties' extra_children.
    #[cfg(all(feature = "wml-styling", feature = "extra-children"))]
    pub fn set_indent(&mut self, left: Option<u32>, first_line: Option<u32>) {
        use ooxml_xml::{RawXmlElement, RawXmlNode};

        let ppr = self
            .p_pr
            .get_or_insert_with(|| Box::new(types::ParagraphProperties::default()));
        let mut attrs = vec![];
        if let Some(l) = left {
            attrs.push(("w:left".to_string(), l.to_string()));
        }
        if let Some(fl) = first_line {
            attrs.push(("w:firstLine".to_string(), fl.to_string()));
        }
        let ind = RawXmlElement {
            name: "w:ind".to_string(),
            attributes: attrs,
            children: vec![],
            self_closing: true,
        };
        ppr.extra_children.push(RawXmlNode::Element(ind));
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
        self.run_inner_content
            .push(Box::new(types::EGRunInnerContent::T(Box::new(t))));
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
        self.run_inner_content
            .push(Box::new(types::EGRunInnerContent::Br(Box::new(
                types::CTBr {
                    r#type: Some(types::STBrType::Page),
                    clear: None,
                    #[cfg(feature = "extra-attrs")]
                    extra_attrs: Default::default(),
                },
            ))));
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
        self.run_inner_content
            .push(Box::new(types::EGRunInnerContent::Drawing(Box::new(
                drawing,
            ))));
    }

    /// Add a footnote reference to this run.
    pub fn add_footnote_ref(&mut self, id: i64) {
        self.run_inner_content
            .push(Box::new(types::EGRunInnerContent::FootnoteReference(
                Box::new(types::CTFtnEdnRef {
                    custom_mark_follows: None,
                    id,
                    #[cfg(feature = "extra-attrs")]
                    extra_attrs: Default::default(),
                }),
            )));
    }

    /// Add an endnote reference to this run.
    pub fn add_endnote_ref(&mut self, id: i64) {
        self.run_inner_content
            .push(Box::new(types::EGRunInnerContent::EndnoteReference(
                Box::new(types::CTFtnEdnRef {
                    custom_mark_follows: None,
                    id,
                    #[cfg(feature = "extra-attrs")]
                    extra_attrs: Default::default(),
                }),
            )));
    }

    /// Add a comment reference to this run.
    pub fn add_comment_ref(&mut self, id: i64) {
        self.run_inner_content
            .push(Box::new(types::EGRunInnerContent::CommentReference(
                Box::new(types::CTMarkup {
                    id,
                    #[cfg(feature = "extra-attrs")]
                    extra_attrs: Default::default(),
                }),
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
        self.p_content
            .push(Box::new(types::EGPContent::R(Box::default())));
        match self.p_content.last_mut().unwrap().as_mut() {
            types::EGPContent::R(r) => r.as_mut(),
            _ => unreachable!(),
        }
    }

    /// Set the relationship ID (for external hyperlinks).
    pub fn set_rel_id(&mut self, rel_id: &str) {
        #[cfg(feature = "extra-attrs")]
        {
            self.extra_attrs
                .insert("r:id".to_string(), rel_id.to_string());
        }
        #[cfg(not(feature = "extra-attrs"))]
        {
            let _ = rel_id;
        }
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
        self.content_row_content
            .push(Box::new(types::EGContentRowContent::Tr(Box::default())));
        match self.content_row_content.last_mut().unwrap().as_mut() {
            types::EGContentRowContent::Tr(r) => r.as_mut(),
            _ => unreachable!(),
        }
    }
}

#[cfg(feature = "wml-tables")]
impl types::CTRow {
    /// Add a cell and return a mutable reference.
    pub fn add_cell(&mut self) -> &mut types::TableCell {
        self.content_cell_content
            .push(Box::new(types::EGContentCellContent::Tc(Box::default())));
        match self.content_cell_content.last_mut().unwrap().as_mut() {
            types::EGContentCellContent::Tc(c) => c.as_mut(),
            _ => unreachable!(),
        }
    }
}

#[cfg(feature = "wml-tables")]
impl types::TableCell {
    /// Add a paragraph and return a mutable reference.
    pub fn add_paragraph(&mut self) -> &mut types::Paragraph {
        self.block_level_elts
            .push(Box::new(types::EGBlockLevelElts::P(Box::default())));
        match self.block_level_elts.last_mut().unwrap().as_mut() {
            types::EGBlockLevelElts::P(p) => p.as_mut(),
            _ => unreachable!(),
        }
    }
}

// =============================================================================
// Header/Footer (CTHdrFtr)
// =============================================================================

impl types::CTHdrFtr {
    /// Add an empty paragraph and return a mutable reference.
    pub fn add_paragraph(&mut self) -> &mut types::Paragraph {
        self.block_level_elts
            .push(Box::new(types::EGBlockLevelElts::P(Box::default())));
        match self.block_level_elts.last_mut().unwrap().as_mut() {
            types::EGBlockLevelElts::P(p) => p.as_mut(),
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
        self.block_level_elts
            .push(Box::new(types::EGBlockLevelElts::P(Box::default())));
        match self.block_level_elts.last_mut().unwrap().as_mut() {
            types::EGBlockLevelElts::P(p) => p.as_mut(),
            _ => unreachable!(),
        }
    }
}

// =============================================================================
// Footnote/Endnote (CTFtnEdn)
// =============================================================================

impl types::CTFtnEdn {
    /// Add a paragraph and return a mutable reference.
    pub fn add_paragraph(&mut self) -> &mut types::Paragraph {
        self.block_level_elts
            .push(Box::new(types::EGBlockLevelElts::P(Box::default())));
        match self.block_level_elts.last_mut().unwrap().as_mut() {
            types::EGBlockLevelElts::P(p) => p.as_mut(),
            _ => unreachable!(),
        }
    }
}
