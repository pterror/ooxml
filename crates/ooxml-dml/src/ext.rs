//! Extension traits for DrawingML types.
//!
//! Provides convenience methods for working with generated DML types.

#[cfg(feature = "dml-text")]
use crate::types::*;

/// Extension trait for [`TextBody`] providing convenience methods.
#[cfg(feature = "dml-text")]
pub trait TextBodyExt {
    /// Get all paragraphs in the text body.
    fn paragraphs(&self) -> &[TextParagraph];

    /// Extract all text from the text body.
    fn text(&self) -> String;
}

#[cfg(feature = "dml-text")]
impl TextBodyExt for TextBody {
    fn paragraphs(&self) -> &[TextParagraph] {
        &self.p
    }

    fn text(&self) -> String {
        self.p
            .iter()
            .map(|p| p.text())
            .collect::<Vec<_>>()
            .join("\n")
    }
}

/// Extension trait for [`TextParagraph`] providing convenience methods.
#[cfg(feature = "dml-text")]
pub trait TextParagraphExt {
    /// Get all text runs in the paragraph.
    fn runs(&self) -> Vec<&TextRun>;

    /// Extract all text from the paragraph.
    fn text(&self) -> String;

    /// Get the paragraph level (for bullets/numbering).
    fn level(&self) -> Option<i32>;

    /// Get the text alignment.
    fn alignment(&self) -> Option<STTextAlignType>;
}

#[cfg(feature = "dml-text")]
impl TextParagraphExt for TextParagraph {
    fn runs(&self) -> Vec<&TextRun> {
        self.text_run
            .iter()
            .filter_map(|tr| match tr {
                EGTextRun::R(run) => Some(run.as_ref()),
                _ => None,
            })
            .collect()
    }

    fn text(&self) -> String {
        self.text_run
            .iter()
            .filter_map(|tr| match tr {
                EGTextRun::R(run) => Some(run.t.as_str()),
                EGTextRun::Br(_) => Some("\n"),
                EGTextRun::Fld(fld) => fld.t.as_deref(),
            })
            .collect()
    }

    fn level(&self) -> Option<i32> {
        self.p_pr.as_ref().and_then(|p| p.lvl)
    }

    fn alignment(&self) -> Option<STTextAlignType> {
        self.p_pr.as_ref().and_then(|p| p.algn)
    }
}

/// Extension trait for [`TextRun`] providing convenience methods.
#[cfg(feature = "dml-text")]
pub trait TextRunExt {
    /// Get the text content.
    fn text(&self) -> &str;

    /// Check if the text is bold.
    fn is_bold(&self) -> bool;

    /// Check if the text is italic.
    fn is_italic(&self) -> bool;

    /// Check if the text is underlined.
    fn is_underlined(&self) -> bool;

    /// Check if the text has strikethrough.
    fn is_strikethrough(&self) -> bool;

    /// Get the font size in hundredths of a point.
    fn font_size(&self) -> Option<i32>;

    /// Check if the run has a hyperlink.
    fn has_hyperlink(&self) -> bool;

    /// Get the hyperlink relationship ID.
    fn hyperlink_rel_id(&self) -> Option<&str>;
}

#[cfg(feature = "dml-text")]
impl TextRunExt for TextRun {
    fn text(&self) -> &str {
        &self.t
    }

    fn is_bold(&self) -> bool {
        self.r_pr.as_ref().and_then(|p| p.b).unwrap_or(false)
    }

    fn is_italic(&self) -> bool {
        self.r_pr.as_ref().and_then(|p| p.i).unwrap_or(false)
    }

    fn is_underlined(&self) -> bool {
        self.r_pr
            .as_ref()
            .and_then(|p| p.u.as_ref())
            .is_some_and(|u| *u != STTextUnderlineType::None)
    }

    fn is_strikethrough(&self) -> bool {
        self.r_pr
            .as_ref()
            .and_then(|p| p.strike.as_ref())
            .is_some_and(|s| *s != STTextStrikeType::NoStrike)
    }

    fn font_size(&self) -> Option<i32> {
        self.r_pr.as_ref().and_then(|p| p.sz)
    }

    fn has_hyperlink(&self) -> bool {
        self.r_pr
            .as_ref()
            .and_then(|p| p.hlink_click.as_ref())
            .is_some()
    }

    fn hyperlink_rel_id(&self) -> Option<&str> {
        self.r_pr
            .as_ref()
            .and_then(|p| p.hlink_click.as_ref())
            .and_then(|h| h.id.as_deref())
    }
}

/// Extension trait for [`CTTable`] providing convenience methods.
#[cfg(feature = "dml-tables")]
pub trait TableExt {
    /// Get all rows in the table.
    fn rows(&self) -> &[CTTableRow];

    /// Get the number of rows.
    fn row_count(&self) -> usize;

    /// Get the number of columns (from grid, or first row if empty).
    fn col_count(&self) -> usize;

    /// Get a cell by row and column index (0-based).
    fn cell(&self, row: usize, col: usize) -> Option<&CTTableCell>;

    /// Get all cell text as a 2D vector.
    fn to_text_grid(&self) -> Vec<Vec<String>>;

    /// Get plain text representation (tab-separated values).
    fn text(&self) -> String;
}

#[cfg(feature = "dml-tables")]
impl TableExt for CTTable {
    fn rows(&self) -> &[CTTableRow] {
        &self.tr
    }

    fn row_count(&self) -> usize {
        self.tr.len()
    }

    fn col_count(&self) -> usize {
        self.tbl_grid.grid_col.len()
    }

    fn cell(&self, row: usize, col: usize) -> Option<&CTTableCell> {
        self.tr.get(row).and_then(|r| r.tc.get(col))
    }

    fn to_text_grid(&self) -> Vec<Vec<String>> {
        self.tr
            .iter()
            .map(|row| row.tc.iter().map(|c| c.text()).collect())
            .collect()
    }

    fn text(&self) -> String {
        self.tr
            .iter()
            .map(|row| {
                row.tc
                    .iter()
                    .map(|c| c.text())
                    .collect::<Vec<_>>()
                    .join("\t")
            })
            .collect::<Vec<_>>()
            .join("\n")
    }
}

/// Extension trait for [`CTTableRow`] providing convenience methods.
#[cfg(feature = "dml-tables")]
pub trait TableRowExt {
    /// Get all cells in this row.
    fn cells(&self) -> &[CTTableCell];

    /// Get a cell by column index (0-based).
    fn cell(&self, col: usize) -> Option<&CTTableCell>;

    /// Get the row height in EMUs (if specified).
    fn height_emu(&self) -> Option<i64>;
}

#[cfg(feature = "dml-tables")]
impl TableRowExt for CTTableRow {
    fn cells(&self) -> &[CTTableCell] {
        &self.tc
    }

    fn cell(&self, col: usize) -> Option<&CTTableCell> {
        self.tc.get(col)
    }

    fn height_emu(&self) -> Option<i64> {
        self.height.parse::<i64>().ok()
    }
}

/// Extension trait for [`CTTableCell`] providing convenience methods.
#[cfg(feature = "dml-tables")]
pub trait TableCellExt {
    /// Get the text body (paragraphs) if present.
    fn text_body(&self) -> Option<&TextBody>;

    /// Get the cell text (paragraphs joined with newlines).
    fn text(&self) -> String;

    /// Get the row span (number of rows this cell spans).
    fn row_span(&self) -> u32;

    /// Get the column span (number of columns this cell spans).
    fn col_span(&self) -> u32;

    /// Check if this cell spans multiple rows.
    fn has_row_span(&self) -> bool;

    /// Check if this cell spans multiple columns.
    fn has_col_span(&self) -> bool;

    /// Check if this cell is merged horizontally (continuation of previous cell).
    fn is_h_merge(&self) -> bool;

    /// Check if this cell is merged vertically (continuation of cell above).
    fn is_v_merge(&self) -> bool;
}

#[cfg(feature = "dml-tables")]
impl TableCellExt for CTTableCell {
    fn text_body(&self) -> Option<&TextBody> {
        self.tx_body.as_deref()
    }

    fn text(&self) -> String {
        self.tx_body
            .as_ref()
            .map(|tb| tb.text())
            .unwrap_or_default()
    }

    fn row_span(&self) -> u32 {
        self.row_span.map(|s| s.max(1) as u32).unwrap_or(1)
    }

    fn col_span(&self) -> u32 {
        self.grid_span.map(|s| s.max(1) as u32).unwrap_or(1)
    }

    fn has_row_span(&self) -> bool {
        self.row_span.is_some_and(|s| s > 1)
    }

    fn has_col_span(&self) -> bool {
        self.grid_span.is_some_and(|s| s > 1)
    }

    fn is_h_merge(&self) -> bool {
        self.h_merge.unwrap_or(false)
    }

    fn is_v_merge(&self) -> bool {
        self.v_merge.unwrap_or(false)
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_text_paragraph_text() {
        let para = TextParagraph {
            #[cfg(feature = "dml-text")]
            p_pr: None,
            text_run: vec![
                EGTextRun::R(Box::new(TextRun {
                    #[cfg(feature = "dml-text")]
                    r_pr: None,
                    #[cfg(feature = "dml-text")]
                    t: "Hello ".to_string(),
                    #[cfg(feature = "extra-children")]
                    extra_children: Vec::new(),
                })),
                EGTextRun::R(Box::new(TextRun {
                    #[cfg(feature = "dml-text")]
                    r_pr: None,
                    #[cfg(feature = "dml-text")]
                    t: "World".to_string(),
                    #[cfg(feature = "extra-children")]
                    extra_children: Vec::new(),
                })),
            ],
            #[cfg(feature = "dml-text")]
            end_para_r_pr: None,
            #[cfg(feature = "extra-children")]
            extra_children: Vec::new(),
        };

        assert_eq!(para.text(), "Hello World");
        assert_eq!(para.runs().len(), 2);
    }

    #[cfg(feature = "dml-text")]
    #[test]
    fn test_text_run_formatting() {
        let run = TextRun {
            r_pr: Some(Box::new(TextCharacterProperties {
                b: Some(true),
                i: Some(true),
                ..Default::default()
            })),
            t: "Bold Italic".to_string(),
            #[cfg(feature = "extra-children")]
            extra_children: Vec::new(),
        };

        assert!(run.is_bold());
        assert!(run.is_italic());
        assert!(!run.is_underlined());
        assert_eq!(run.text(), "Bold Italic");
    }
}
