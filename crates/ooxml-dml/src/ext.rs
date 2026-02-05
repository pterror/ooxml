//! Extension traits for DrawingML types.
//!
//! Provides convenience methods for working with generated DML types.

use crate::types::*;

/// Extension trait for [`TextBody`] providing convenience methods.
pub trait TextBodyExt {
    /// Get all paragraphs in the text body.
    fn paragraphs(&self) -> &[TextParagraph];

    /// Extract all text from the text body.
    fn text(&self) -> String;
}

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
pub trait TextParagraphExt {
    /// Get all text runs in the paragraph.
    fn runs(&self) -> Vec<&TextRun>;

    /// Extract all text from the paragraph.
    fn text(&self) -> String;

    /// Get the paragraph level (for bullets/numbering).
    #[cfg(feature = "dml-text")]
    fn level(&self) -> Option<i32>;

    /// Get the text alignment.
    #[cfg(feature = "dml-text")]
    fn alignment(&self) -> Option<STTextAlignType>;
}

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

    #[cfg(feature = "dml-text")]
    fn level(&self) -> Option<i32> {
        self.p_pr.as_ref().and_then(|p| p.lvl)
    }

    #[cfg(feature = "dml-text")]
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
