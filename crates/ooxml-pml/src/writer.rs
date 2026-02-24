//! PowerPoint presentation writing support.
//!
//! This module provides `PresentationBuilder` for creating new PPTX files.
//!
//! # Example
//!
//! ```no_run
//! use ooxml_pml::PresentationBuilder;
//!
//! let mut pres = PresentationBuilder::new();
//! let slide = pres.add_slide();
//! slide.add_title("Hello World");
//! slide.add_text("This is a presentation created with ooxml-pml");
//! pres.save("output.pptx")?;
//! # Ok::<(), ooxml_pml::Error>(())
//! ```

use crate::error::Result;
use crate::generated_serializers::ToXml;
use crate::types;
use ooxml_dml::types as dml;
use ooxml_opc::PackageWriter;
use std::fs::File;
use std::io::{BufWriter, Seek, Write};
use std::path::Path;

// Content types
const CT_PRESENTATION: &str =
    "application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml";
const CT_SLIDE: &str = "application/vnd.openxmlformats-officedocument.presentationml.slide+xml";
const CT_NOTES_SLIDE: &str =
    "application/vnd.openxmlformats-officedocument.presentationml.notesSlide+xml";
const CT_RELATIONSHIPS: &str = "application/vnd.openxmlformats-package.relationships+xml";
const CT_XML: &str = "application/xml";

// Image content types
const CT_JPEG: &str = "image/jpeg";
const CT_PNG: &str = "image/png";
const CT_GIF: &str = "image/gif";

// Relationship types
const REL_OFFICE_DOCUMENT: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument";
const REL_SLIDE: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide";
const REL_NOTES_SLIDE: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesSlide";
const REL_IMAGE: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image";
const REL_HYPERLINK: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink";
const REL_SLIDE_MASTER: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster";
const REL_SLIDE_LAYOUT: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout";

// Namespaces
const NS_PRES: &str = "http://schemas.openxmlformats.org/presentationml/2006/main";
const NS_DRAWING: &str = "http://schemas.openxmlformats.org/drawingml/2006/main";
const NS_REL: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

// Namespace declarations for PML slides
const NS_DECLS_SLIDE: &[(&str, &str)] = &[
    ("xmlns:a", NS_DRAWING),
    ("xmlns:r", NS_REL),
    ("xmlns:p", NS_PRES),
];

/// Build a minimal slide master XML string.
///
/// The master contains an empty shape tree and the required color map.  A
/// single slide layout (rId1 → slideLayout1.xml) is listed in the layout id
/// list so that slides can reference it.
fn build_slide_master_xml() -> &'static str {
    r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sldMaster xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:cSld>
    <p:spTree>
      <p:nvGrpSpPr>
        <p:cNvPr id="1" name=""/>
        <p:cNvGrpSpPr/>
        <p:nvPr/>
      </p:nvGrpSpPr>
      <p:grpSpPr>
        <a:xfrm>
          <a:off x="0" y="0"/>
          <a:ext cx="0" cy="0"/>
          <a:chOff x="0" y="0"/>
          <a:chExt cx="0" cy="0"/>
        </a:xfrm>
      </p:grpSpPr>
    </p:spTree>
  </p:cSld>
  <p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/>
  <p:sldLayoutIdLst>
    <p:sldLayoutId id="2147483648" r:id="rId1"/>
  </p:sldLayoutIdLst>
</p:sldMaster>"#
}

/// Build a minimal blank slide layout XML string.
fn build_slide_layout_xml() -> &'static str {
    r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sldLayout xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" type="blank" preserve="1">
  <p:cSld name="Blank">
    <p:spTree>
      <p:nvGrpSpPr>
        <p:cNvPr id="1" name=""/>
        <p:cNvGrpSpPr/>
        <p:nvPr/>
      </p:nvGrpSpPr>
      <p:grpSpPr>
        <a:xfrm>
          <a:off x="0" y="0"/>
          <a:ext cx="0" cy="0"/>
          <a:chOff x="0" y="0"/>
          <a:chExt cx="0" cy="0"/>
        </a:xfrm>
      </p:grpSpPr>
    </p:spTree>
  </p:cSld>
</p:sldLayout>"#
}

/// Serialize any `ToXml` type to XML bytes with a given tag name and PML namespace declarations.
fn serialize_pml_xml<T: ToXml>(value: &T, tag: &str) -> Result<Vec<u8>> {
    use quick_xml::Writer;
    use quick_xml::events::{BytesEnd, BytesStart, Event};

    let inner = Vec::new();
    let mut writer = Writer::new(inner);

    let start = BytesStart::new(tag);
    let start = value.write_attrs(start);
    let mut start = start;
    for &(key, val) in NS_DECLS_SLIDE {
        start.push_attribute((key, val));
    }

    if value.is_empty_element() {
        writer.write_event(Event::Empty(start))?;
    } else {
        writer.write_event(Event::Start(start))?;
        value.write_children(&mut writer)?;
        writer.write_event(Event::End(BytesEnd::new(tag)))?;
    }

    let inner = writer.into_inner();
    let mut buf = Vec::with_capacity(
        b"<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n".len() + inner.len(),
    );
    buf.extend_from_slice(b"<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n");
    buf.extend_from_slice(&inner);
    Ok(buf)
}

/// Serialize a `types::Slide` to XML bytes with PML namespace declarations.
fn serialize_slide_xml(slide: &types::Slide) -> Result<Vec<u8>> {
    serialize_pml_xml(slide, "p:sld")
}

/// Construct a `CTNonVisualDrawingProps` with required fields.
fn make_cnv_pr(id: u32, name: &str) -> Box<dml::CTNonVisualDrawingProps> {
    Box::new(dml::CTNonVisualDrawingProps {
        id,
        name: name.to_string(),
        descr: None,
        hidden: None,
        title: None,
        hlink_click: None,
        hlink_hover: None,
        ext_lst: None,
        extra_attrs: Default::default(),
        extra_children: Default::default(),
    })
}

/// Construct an empty `CTApplicationNonVisualDrawingProps`.
fn make_nv_pr() -> Box<types::CTApplicationNonVisualDrawingProps> {
    Box::new(types::CTApplicationNonVisualDrawingProps {
        is_photo: None,
        user_drawn: None,
        ph: None,
        cust_data_lst: None,
        ext_lst: None,
        extra_attrs: Default::default(),
        extra_children: Default::default(),
    })
}

/// Construct a `Transform2D` from position and size (all in EMUs / i64).
fn make_xfrm(x: i64, y: i64, cx: i64, cy: i64) -> Box<dml::Transform2D> {
    Box::new(dml::Transform2D {
        offset: Some(Box::new(dml::Point2D {
            x: x.to_string(),
            y: y.to_string(),
            extra_attrs: Default::default(),
        })),
        extents: Some(Box::new(dml::PositiveSize2D {
            cx,
            cy,
            extra_attrs: Default::default(),
        })),
        ..Default::default()
    })
}

/// Construct a preset rect geometry (prstGeom prst="rect").
fn make_rect_geom() -> Box<dml::EGGeometry> {
    Box::new(dml::EGGeometry::PrstGeom(Box::new(
        dml::CTPresetGeometry2D {
            preset: dml::STShapeType::Rect,
            av_lst: None,
            extra_attrs: Default::default(),
            extra_children: Default::default(),
        },
    )))
}

/// A text run in a paragraph, optionally with a hyperlink and formatting.
#[derive(Debug, Clone)]
pub struct TextRun {
    text: String,
    /// Optional hyperlink URL.
    hyperlink: Option<String>,
    /// Optional tooltip text for the hyperlink.
    tooltip: Option<String>,
    /// Font size in points (e.g. 24.0 for 24pt). Stored as hundredths of a point internally.
    font_size: Option<f32>,
    /// Text color as a 6-hex-digit RGB string (e.g. "FF0000" for red).
    color: Option<String>,
    /// Bold formatting.
    bold: Option<bool>,
    /// Italic formatting.
    italic: Option<bool>,
    /// Underline formatting.
    underline: bool,
}

impl TextRun {
    /// Create a plain text run.
    pub fn text(text: impl Into<String>) -> Self {
        Self {
            text: text.into(),
            hyperlink: None,
            tooltip: None,
            font_size: None,
            color: None,
            bold: None,
            italic: None,
            underline: false,
        }
    }

    /// Create a hyperlink run.
    pub fn hyperlink(text: impl Into<String>, url: impl Into<String>) -> Self {
        Self {
            text: text.into(),
            hyperlink: Some(url.into()),
            tooltip: None,
            font_size: None,
            color: None,
            bold: None,
            italic: None,
            underline: false,
        }
    }

    /// Set the font size in points.
    ///
    /// ECMA-376: `sz` attribute on `TextCharacterProperties` is in hundredths of a point
    /// (e.g. 24pt = 2400). This method accepts points and converts automatically.
    pub fn set_font_size(mut self, pt: f32) -> Self {
        self.font_size = Some(pt);
        self
    }

    /// Set the text color.
    ///
    /// `rgb` must be a 6-character hex string (e.g. `"FF0000"` for red).
    /// Sets `solidFill/srgbClr` on `TextCharacterProperties`.
    pub fn set_color(mut self, rgb: impl Into<String>) -> Self {
        self.color = Some(rgb.into());
        self
    }

    /// Set bold formatting.
    pub fn set_bold(mut self, bold: bool) -> Self {
        self.bold = Some(bold);
        self
    }

    /// Set italic formatting.
    pub fn set_italic(mut self, italic: bool) -> Self {
        self.italic = Some(italic);
        self
    }

    /// Set underline formatting.
    pub fn set_underline(mut self, underline: bool) -> Self {
        self.underline = underline;
        self
    }

    /// Set a hyperlink tooltip.
    ///
    /// Sets the `tooltip` attribute on `hlinkClick`.
    /// Only has effect when this run also has a hyperlink URL.
    pub fn set_tooltip(mut self, text: impl Into<String>) -> Self {
        self.tooltip = Some(text.into());
        self
    }
}

/// Convert a 6-character hex RGB string to bytes (e.g. "FF0000" → [0xFF, 0x00, 0x00]).
fn hex_to_bytes(hex: &str) -> dml::HexColorRgb {
    let hex = hex.trim_start_matches('#');
    (0..hex.len())
        .step_by(2)
        .filter_map(|i| u8::from_str_radix(&hex[i..i + 2], 16).ok())
        .collect()
}

/// Build a `SolidColorFill` from a 6-hex-digit RGB string.
fn make_solid_fill(rgb: &str) -> Box<dml::EGFillProperties> {
    Box::new(dml::EGFillProperties::SolidFill(Box::new(
        dml::SolidColorFill {
            color_choice: Some(Box::new(dml::EGColorChoice::SrgbClr(Box::new(
                dml::SrgbColor {
                    value: hex_to_bytes(rgb),
                    color_transform: Vec::new(),
                    #[cfg(feature = "extra-attrs")]
                    extra_attrs: Default::default(),
                    #[cfg(feature = "extra-children")]
                    extra_children: Default::default(),
                },
            )))),
            #[cfg(feature = "extra-children")]
            extra_children: Default::default(),
        },
    )))
}

/// Builder for adding a shape (text box) with full formatting control.
///
/// Create via [`SlideBuilder::shape`] and finalize via [`ShapeBuilder::add`].
#[derive(Debug)]
pub struct ShapeBuilder<'a> {
    slide: &'a mut SlideBuilder,
    runs: Vec<TextRun>,
    is_title: bool,
    x: i64,
    y: i64,
    width: i64,
    height: i64,
    fill_color: Option<String>,
    line_color: Option<String>,
    line_width: Option<i64>,
}

impl<'a> ShapeBuilder<'a> {
    /// Set the shape position and size (all in EMUs).
    ///
    /// Sets `spPr/xfrm` offsets and extents.
    pub fn set_position(mut self, x: i64, y: i64, cx: i64, cy: i64) -> Self {
        self.x = x;
        self.y = y;
        self.width = cx;
        self.height = cy;
        self
    }

    /// Set the shape fill color.
    ///
    /// `rgb` must be a 6-character hex string (e.g. `"FF0000"` for red).
    /// Sets `spPr/solidFill/srgbClr`.
    pub fn set_fill_color(mut self, rgb: impl Into<String>) -> Self {
        self.fill_color = Some(rgb.into());
        self
    }

    /// Set the shape border/line color.
    ///
    /// `rgb` must be a 6-character hex string.
    /// Sets `spPr/ln/solidFill/srgbClr`.
    pub fn set_line_color(mut self, rgb: impl Into<String>) -> Self {
        self.line_color = Some(rgb.into());
        self
    }

    /// Set the shape border/line width in EMUs.
    ///
    /// Sets the `w` attribute on `spPr/ln`.
    pub fn set_line_width(mut self, emu: i64) -> Self {
        self.line_width = Some(emu);
        self
    }

    /// Finalize the shape and add it to the slide.
    pub fn add(self) {
        let ShapeBuilder {
            slide,
            runs,
            is_title,
            x,
            y,
            width,
            height,
            fill_color,
            line_color,
            line_width,
        } = self;

        let element = TextElement {
            runs,
            is_title,
            x,
            y,
            width,
            height,
            fill_color,
            line_color,
            line_width,
        };
        slide.push_text_element(element);
    }
}

/// A text element (used for shapes with hyperlinks, deferred until write time).
#[derive(Debug, Clone)]
struct TextElement {
    runs: Vec<TextRun>,
    is_title: bool,
    x: i64,
    y: i64,
    width: i64,
    height: i64,
    /// Optional shape fill color (6-hex RGB).
    fill_color: Option<String>,
    /// Optional shape border color (6-hex RGB).
    line_color: Option<String>,
    /// Optional shape border width in EMUs.
    line_width: Option<i64>,
}

impl TextElement {
    fn simple(text: String, is_title: bool, x: i64, y: i64, width: i64, height: i64) -> Self {
        Self {
            runs: vec![TextRun {
                text,
                hyperlink: None,
                tooltip: None,
                font_size: None,
                color: None,
                bold: None,
                italic: None,
                underline: false,
            }],
            is_title,
            x,
            y,
            width,
            height,
            fill_color: None,
            line_color: None,
            line_width: None,
        }
    }

    fn has_hyperlink(&self) -> bool {
        self.runs.iter().any(|r| r.hyperlink.is_some())
    }
}

/// Image format for embedded images.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ImageFormat {
    /// JPEG image.
    Jpeg,
    /// PNG image.
    Png,
    /// GIF image.
    Gif,
}

impl ImageFormat {
    fn extension(self) -> &'static str {
        match self {
            ImageFormat::Jpeg => "jpeg",
            ImageFormat::Png => "png",
            ImageFormat::Gif => "gif",
        }
    }

    fn content_type(self) -> &'static str {
        match self {
            ImageFormat::Jpeg => CT_JPEG,
            ImageFormat::Png => CT_PNG,
            ImageFormat::Gif => CT_GIF,
        }
    }

    /// Detect image format from file extension.
    pub fn from_extension(ext: &str) -> Option<Self> {
        match ext.to_lowercase().as_str() {
            "jpg" | "jpeg" => Some(ImageFormat::Jpeg),
            "png" => Some(ImageFormat::Png),
            "gif" => Some(ImageFormat::Gif),
            _ => None,
        }
    }

    /// Detect image format from magic bytes.
    pub fn from_bytes(data: &[u8]) -> Option<Self> {
        if data.len() < 4 {
            return None;
        }
        if data.starts_with(&[0x89, 0x50, 0x4E, 0x47]) {
            return Some(ImageFormat::Png);
        }
        if data.starts_with(&[0xFF, 0xD8, 0xFF]) {
            return Some(ImageFormat::Jpeg);
        }
        if data.starts_with(b"GIF8") {
            return Some(ImageFormat::Gif);
        }
        None
    }
}

/// An image element to add to a slide.
#[derive(Debug, Clone)]
struct ImageElement {
    data: Vec<u8>,
    format: ImageFormat,
    x: i64,
    y: i64,
    width: i64,
    height: i64,
    description: Option<String>,
}

/// A table element to add to a slide.
#[derive(Debug, Clone)]
struct TableElement {
    name: Option<String>,
    rows: Vec<Vec<String>>,
    col_widths: Vec<i64>,
    row_heights: Vec<i64>,
    x: i64,
    y: i64,
    width: i64,
    height: i64,
}

/// Builder for creating tables in slides.
#[derive(Debug)]
pub struct TableBuilder {
    rows: Vec<Vec<String>>,
    col_widths: Option<Vec<i64>>,
    name: Option<String>,
}

impl TableBuilder {
    /// Create a new table builder.
    pub fn new() -> Self {
        Self {
            rows: Vec::new(),
            col_widths: None,
            name: None,
        }
    }

    /// Set the table name.
    pub fn name(mut self, name: impl Into<String>) -> Self {
        self.name = Some(name.into());
        self
    }

    /// Add a row to the table.
    pub fn add_row<I, S>(mut self, cells: I) -> Self
    where
        I: IntoIterator<Item = S>,
        S: Into<String>,
    {
        self.rows
            .push(cells.into_iter().map(|s| s.into()).collect());
        self
    }

    /// Set column widths in EMUs.
    ///
    /// If not set, columns will be evenly distributed.
    pub fn col_widths<I>(mut self, widths: I) -> Self
    where
        I: IntoIterator<Item = i64>,
    {
        self.col_widths = Some(widths.into_iter().collect());
        self
    }

    /// Get the number of rows.
    pub fn row_count(&self) -> usize {
        self.rows.len()
    }

    /// Get the number of columns (from first row, or 0 if empty).
    pub fn col_count(&self) -> usize {
        self.rows.first().map(|r| r.len()).unwrap_or(0)
    }
}

impl Default for TableBuilder {
    fn default() -> Self {
        Self::new()
    }
}

/// Slide transition settings for the writer.
#[derive(Debug, Clone)]
pub struct SlideTransition {
    /// Transition type.
    pub transition_type: crate::TransitionType,
    /// Transition speed.
    pub speed: crate::TransitionSpeed,
    /// Advance on mouse click.
    pub advance_on_click: bool,
    /// Auto-advance time in milliseconds.
    pub advance_after_ms: Option<u32>,
}

impl SlideTransition {
    /// Create a new transition with the given type.
    pub fn new(transition_type: crate::TransitionType) -> Self {
        Self {
            transition_type,
            speed: crate::TransitionSpeed::Medium,
            advance_on_click: true,
            advance_after_ms: None,
        }
    }

    /// Set the transition speed.
    pub fn with_speed(mut self, speed: crate::TransitionSpeed) -> Self {
        self.speed = speed;
        self
    }

    /// Set whether to advance on click.
    pub fn with_advance_on_click(mut self, advance: bool) -> Self {
        self.advance_on_click = advance;
        self
    }

    /// Set auto-advance time in milliseconds.
    pub fn with_advance_after(mut self, ms: u32) -> Self {
        self.advance_after_ms = Some(ms);
        self
    }
}

/// A slide being built.
#[derive(Debug)]
pub struct SlideBuilder {
    /// Images — raw bytes, need rId at write time.
    images: Vec<ImageElement>,
    /// Text elements containing hyperlinks — need rId at write time.
    hyperlink_elements: Vec<TextElement>,
    /// Speaker notes for this slide.
    notes: Option<String>,
    /// Shape ID counter: starts at 2 (1 is the group shape).
    next_shape_id: usize,
    /// The pre-built slide type.
    slide: types::Slide,
    /// Optional slide background color (6-hex RGB).
    background_color: Option<String>,
}

/// Create an empty `types::Slide` with the required boilerplate shape tree.
fn init_slide() -> types::Slide {
    // Group transform: all zeros for a flat slide.
    let grp_xfrm = dml::CTGroupTransform2D {
        offset: Some(Box::new(dml::Point2D {
            x: "0".to_string(),
            y: "0".to_string(),
            extra_attrs: Default::default(),
        })),
        extents: Some(Box::new(dml::PositiveSize2D {
            cx: 0,
            cy: 0,
            extra_attrs: Default::default(),
        })),
        child_offset: Some(Box::new(dml::Point2D {
            x: "0".to_string(),
            y: "0".to_string(),
            extra_attrs: Default::default(),
        })),
        child_extents: Some(Box::new(dml::PositiveSize2D {
            cx: 0,
            cy: 0,
            extra_attrs: Default::default(),
        })),
        ..Default::default()
    };

    let grp_sp_pr = dml::CTGroupShapeProperties {
        transform: Some(Box::new(grp_xfrm)),
        ..Default::default()
    };

    let nv_grp = Box::new(types::CTGroupShapeNonVisual {
        c_nv_pr: make_cnv_pr(1, ""),
        c_nv_grp_sp_pr: Box::new(dml::CTNonVisualGroupDrawingShapeProps {
            grp_sp_locks: None,
            ext_lst: None,
            extra_children: Default::default(),
        }),
        nv_pr: make_nv_pr(),
        extra_children: Default::default(),
    });

    let shape_tree = Box::new(types::GroupShape {
        non_visual_group_properties: nv_grp,
        grp_sp_pr: Box::new(grp_sp_pr),
        shape: Vec::new(),
        group_shape: Vec::new(),
        graphic_frame: Vec::new(),
        connector: Vec::new(),
        picture: Vec::new(),
        content_part: Vec::new(),
        ext_lst: None,
        extra_children: Default::default(),
    });

    let common_slide_data = Box::new(types::CommonSlideData {
        name: None,
        bg: None,
        shape_tree,
        cust_data_lst: None,
        controls: None,
        ext_lst: None,
        extra_attrs: Default::default(),
        extra_children: Default::default(),
    });

    types::Slide {
        show_master_sp: None,
        show_master_ph_anim: None,
        show: None,
        common_slide_data,
        clr_map_ovr: None,
        transition: None,
        timing: None,
        ext_lst: None,
        extra_attrs: Default::default(),
        extra_children: Default::default(),
    }
}

/// Build a `types::Shape` from a `TextElement`, resolving hyperlink rel IDs if provided.
fn build_shape_impl(
    element: &TextElement,
    shape_id: usize,
    hyperlink_rel_ids: Option<&std::collections::HashMap<&str, usize>>,
) -> types::Shape {
    let name = if element.is_title { "Title" } else { "Content" };
    let default_font_size: i32 = if element.is_title { 4400 } else { 2400 };

    let runs: Vec<dml::EGTextRun> = element
        .runs
        .iter()
        .map(|run| {
            // Build hyperlink if present and we have rel IDs
            let hlink_click = run.hyperlink.as_deref().and_then(|url| {
                hyperlink_rel_ids?.get(url).map(|&rel_id| {
                    Box::new(dml::CTHyperlink {
                        id: Some(format!("rId{}", rel_id)),
                        tooltip: run.tooltip.clone(),
                        ..Default::default()
                    })
                })
            });

            // Font size: per-run override, else element default.
            let sz = Some(
                run.font_size
                    .map(|pt| (pt * 100.0).round() as i32)
                    .unwrap_or(default_font_size),
            );

            // Color: per-run solidFill override.
            let fill_properties = run.color.as_deref().map(make_solid_fill);

            dml::EGTextRun::R(Box::new(dml::TextRun {
                r_pr: Some(Box::new(dml::TextCharacterProperties {
                    lang: Some("en-US".to_string()),
                    sz,
                    b: run.bold,
                    i: run.italic,
                    u: if run.underline {
                        Some(dml::STTextUnderlineType::Sng)
                    } else {
                        None
                    },
                    fill_properties,
                    hlink_click,
                    ..Default::default()
                })),
                t: run.text.clone(),
                extra_children: Default::default(),
            }))
        })
        .collect();

    let para = dml::TextParagraph {
        text_run: runs,
        ..Default::default()
    };

    let text_body = dml::TextBody {
        body_pr: Box::default(),
        lst_style: None,
        p: vec![para],
        extra_children: Default::default(),
    };

    // Shape fill color.
    let fill_properties = element.fill_color.as_deref().map(make_solid_fill);

    // Shape border/line (dml-lines fields are always available — DML defaults include all features).
    let line = {
        let has_line = element.line_color.is_some() || element.line_width.is_some();
        if has_line {
            let line_fill = element.line_color.as_deref().map(|rgb| {
                Box::new(dml::EGLineFillProperties::SolidFill(Box::new(
                    dml::SolidColorFill {
                        color_choice: Some(Box::new(dml::EGColorChoice::SrgbClr(Box::new(
                            dml::SrgbColor {
                                value: hex_to_bytes(rgb),
                                color_transform: Vec::new(),
                                #[cfg(feature = "extra-attrs")]
                                extra_attrs: Default::default(),
                                #[cfg(feature = "extra-children")]
                                extra_children: Default::default(),
                            },
                        )))),
                        #[cfg(feature = "extra-children")]
                        extra_children: Default::default(),
                    },
                )))
            });
            Some(Box::new(dml::LineProperties {
                width: element.line_width.map(|w| w as i32),
                line_fill_properties: line_fill,
                ..Default::default()
            }))
        } else {
            None
        }
    };

    let sp_pr = dml::CTShapeProperties {
        transform: Some(make_xfrm(
            element.x,
            element.y,
            element.width,
            element.height,
        )),
        geometry: Some(make_rect_geom()),
        fill_properties,
        line,
        ..Default::default()
    };

    let c_nv_sp_pr = dml::CTNonVisualDrawingShapeProps {
        tx_box: Some(true),
        ..Default::default()
    };

    types::Shape {
        use_bg_fill: None,
        non_visual_properties: Box::new(types::ShapeNonVisual {
            c_nv_pr: make_cnv_pr(shape_id as u32, name),
            c_nv_sp_pr: Box::new(c_nv_sp_pr),
            nv_pr: make_nv_pr(),
            extra_children: Default::default(),
        }),
        shape_properties: Box::new(sp_pr),
        style: None,
        text_body: Some(Box::new(text_body)),
        ext_lst: None,
        extra_attrs: Default::default(),
        extra_children: Default::default(),
    }
}

/// Build a `types::Picture` for an image element.
fn build_picture(image: &ImageElement, shape_id: usize, rel_id: usize) -> types::Picture {
    let pic_name = format!("Picture {}", shape_id);

    let blip = dml::Blip {
        embed: Some(format!("rId{}", rel_id)),
        ..Default::default()
    };

    let stretch = dml::CTStretchInfoProperties {
        fill_rect: Some(Box::default()),
        extra_children: Default::default(),
    };

    let blip_fill = dml::BlipFillProperties {
        blip: Some(Box::new(blip)),
        fill_mode_properties: Some(Box::new(dml::EGFillModeProperties::Stretch(Box::new(
            stretch,
        )))),
        ..Default::default()
    };

    let sp_pr = dml::CTShapeProperties {
        transform: Some(make_xfrm(image.x, image.y, image.width, image.height)),
        geometry: Some(make_rect_geom()),
        ..Default::default()
    };

    let pic_locks = dml::CTPictureLocking {
        no_change_aspect: Some(true),
        ..Default::default()
    };

    let c_nv_pic_pr = dml::CTNonVisualPictureProperties {
        pic_locks: Some(Box::new(pic_locks)),
        ..Default::default()
    };

    let mut c_nv_pr = *make_cnv_pr(shape_id as u32, &pic_name);
    c_nv_pr.descr = image.description.clone();

    types::Picture {
        non_visual_picture_properties: Box::new(types::CTPictureNonVisual {
            c_nv_pr: Box::new(c_nv_pr),
            c_nv_pic_pr: Box::new(c_nv_pic_pr),
            nv_pr: make_nv_pr(),
            extra_children: Default::default(),
        }),
        blip_fill: Box::new(blip_fill),
        shape_properties: Box::new(sp_pr),
        style: None,
        ext_lst: None,
        extra_children: Default::default(),
    }
}

/// Build the raw XML string for `a:graphic/a:graphicData/a:tbl` table content.
fn build_table_graphic_xml(table: &TableElement) -> String {
    let mut xml = String::new();
    xml.push_str(r#"<a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">"#);
    xml.push_str(r#"<a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table">"#);
    xml.push_str("<a:tbl>");
    xml.push_str(r#"<a:tblPr firstRow="1" bandRow="1"/>"#);
    xml.push_str("<a:tblGrid>");
    for col_width in &table.col_widths {
        xml.push_str(&format!(r#"<a:gridCol w="{}"/>"#, col_width));
    }
    xml.push_str("</a:tblGrid>");
    for (row_idx, row) in table.rows.iter().enumerate() {
        let row_height = table.row_heights.get(row_idx).copied().unwrap_or(370840);
        xml.push_str(&format!(r#"<a:tr h="{}">"#, row_height));
        for cell_text in row {
            xml.push_str("<a:tc>");
            xml.push_str("<a:txBody>");
            xml.push_str("<a:bodyPr/>");
            xml.push_str("<a:lstStyle/>");
            xml.push_str("<a:p>");
            xml.push_str("<a:r>");
            xml.push_str(r#"<a:rPr lang="en-US"/>"#);
            xml.push_str(&format!("<a:t>{}</a:t>", escape_xml(cell_text)));
            xml.push_str("</a:r>");
            xml.push_str("</a:p>");
            xml.push_str("</a:txBody>");
            xml.push_str("<a:tcPr/>");
            xml.push_str("</a:tc>");
        }
        xml.push_str("</a:tr>");
    }
    xml.push_str("</a:tbl>");
    xml.push_str("</a:graphicData>");
    xml.push_str("</a:graphic>");
    xml
}

/// Parse a raw XML element into a `PositionedNode` for use in `extra_children`.
///
/// The `a:graphic` element is the top-level content inside a graphicFrame.
/// It cannot be stored in a generated field because `GraphicalObjectFrame`
/// has no `graphic` field in the generated PML type — the element is captured
/// as an unknown child during parsing.
fn parse_extra_child_xml(xml: &str) -> Vec<ooxml_xml::PositionedNode> {
    use quick_xml::Reader;
    use quick_xml::events::Event;

    let mut reader = Reader::from_str(xml);
    reader.config_mut().trim_text(false);

    let mut buf = Vec::new();
    let mut result = Vec::new();
    let mut position = 0usize;

    loop {
        buf.clear();
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(ref e)) => {
                match ooxml_xml::RawXmlElement::from_reader(&mut reader, e) {
                    Ok(elem) => {
                        result.push(ooxml_xml::PositionedNode::new(
                            position,
                            ooxml_xml::RawXmlNode::Element(elem),
                        ));
                        position += 1;
                    }
                    Err(_) => break,
                }
            }
            Ok(Event::Empty(ref e)) => {
                // Self-closing element — shouldn't appear for a:graphic but handle anyway
                let name = String::from_utf8_lossy(e.name().as_ref()).to_string();
                let attrs = e
                    .attributes()
                    .filter_map(|a| a.ok())
                    .map(|a| {
                        (
                            String::from_utf8_lossy(a.key.as_ref()).to_string(),
                            String::from_utf8_lossy(&a.value).to_string(),
                        )
                    })
                    .collect();
                let elem = ooxml_xml::RawXmlElement {
                    name,
                    attributes: attrs,
                    children: Vec::new(),
                    self_closing: true,
                };
                result.push(ooxml_xml::PositionedNode::new(
                    position,
                    ooxml_xml::RawXmlNode::Element(elem),
                ));
                position += 1;
            }
            Ok(Event::Eof) => break,
            Err(_) => break,
            _ => {}
        }
    }

    result
}

/// Build a `types::GraphicalObjectFrame` for a table element.
fn build_graphic_frame(table: &TableElement, shape_id: usize) -> types::GraphicalObjectFrame {
    let name = table.name.as_deref().unwrap_or("Table");

    let graphic_frame_locks = dml::CTGraphicalObjectFrameLocking {
        no_grp: Some(true),
        ..Default::default()
    };

    let c_nv_graphic_frame_pr = dml::CTNonVisualGraphicFrameProperties {
        graphic_frame_locks: Some(Box::new(graphic_frame_locks)),
        ..Default::default()
    };

    let nv_graphic_frame_pr = Box::new(types::CTGraphicalObjectFrameNonVisual {
        c_nv_pr: make_cnv_pr(shape_id as u32, name),
        c_nv_graphic_frame_pr: Box::new(c_nv_graphic_frame_pr),
        nv_pr: make_nv_pr(),
        extra_children: Default::default(),
    });

    let xfrm = make_xfrm(table.x, table.y, table.width, table.height);

    // Build the a:graphic XML and store it in extra_children.
    let graphic_xml = build_table_graphic_xml(table);
    let extra_children = parse_extra_child_xml(&graphic_xml);

    types::GraphicalObjectFrame {
        bw_mode: None,
        nv_graphic_frame_pr,
        xfrm,
        ext_lst: None,
        extra_attrs: Default::default(),
        extra_children,
    }
}

/// Convert a `SlideTransition` builder into a generated `types::SlideTransition`.
#[cfg(feature = "pml-transitions")]
fn build_slide_transition(t: &SlideTransition) -> types::SlideTransition {
    use crate::TransitionSpeed;
    use crate::TransitionType;

    let spd = Some(match t.speed {
        TransitionSpeed::Slow => types::STTransitionSpeed::Slow,
        TransitionSpeed::Medium => types::STTransitionSpeed::Med,
        TransitionSpeed::Fast => types::STTransitionSpeed::Fast,
    });

    let mut st = types::SlideTransition {
        spd,
        adv_click: Some(t.advance_on_click),
        adv_tm: t.advance_after_ms,
        ..Default::default()
    };

    match &t.transition_type {
        TransitionType::Fade => st.fade = Some(Box::default()),
        TransitionType::Push => st.push = Some(Box::default()),
        TransitionType::Wipe => st.wipe = Some(Box::default()),
        TransitionType::Split => st.split = Some(Box::default()),
        TransitionType::Blinds => st.blinds = Some(Box::default()),
        TransitionType::Checker => st.checker = Some(Box::default()),
        TransitionType::Circle => st.circle = Some(Box::default()),
        TransitionType::Dissolve => st.dissolve = Some(Box::default()),
        TransitionType::Comb => st.comb = Some(Box::default()),
        TransitionType::Cover => st.cover = Some(Box::default()),
        TransitionType::Cut => st.cut = Some(Box::default()),
        TransitionType::Diamond => st.diamond = Some(Box::default()),
        TransitionType::Plus => st.plus = Some(Box::default()),
        TransitionType::Random => st.random = Some(Box::default()),
        TransitionType::Strips => st.strips = Some(Box::default()),
        TransitionType::Wedge => st.wedge = Some(Box::default()),
        TransitionType::Wheel => st.wheel = Some(Box::default()),
        TransitionType::Zoom => st.zoom = Some(Box::default()),
        TransitionType::Other(_) => {
            // Unknown transition — only timing attrs set, no child element.
        }
    }

    st
}

impl SlideBuilder {
    fn new() -> Self {
        Self {
            images: Vec::new(),
            hyperlink_elements: Vec::new(),
            notes: None,
            next_shape_id: 2,
            slide: init_slide(),
            background_color: None,
        }
    }

    /// Collect all hyperlinks from this slide's deferred text elements.
    fn hyperlinks(&self) -> Vec<&str> {
        let mut links = Vec::new();
        for element in &self.hyperlink_elements {
            for run in &element.runs {
                if let Some(ref url) = run.hyperlink
                    && !links.contains(&url.as_str())
                {
                    links.push(url.as_str());
                }
            }
        }
        links
    }

    /// Set speaker notes for this slide.
    pub fn set_notes(&mut self, notes: impl Into<String>) -> &mut Self {
        self.notes = Some(notes.into());
        self
    }

    /// Check if this slide has speaker notes.
    pub fn has_notes(&self) -> bool {
        self.notes.as_ref().is_some_and(|n| !n.is_empty())
    }

    /// Check if this slide has images.
    pub fn has_images(&self) -> bool {
        !self.images.is_empty()
    }

    /// Add a title to the slide.
    pub fn add_title(&mut self, text: impl Into<String>) -> &mut Self {
        let element = TextElement::simple(text.into(), true, 457200, 274638, 8229600, 1143000);
        self.push_text_element(element);
        self
    }

    /// Add text content to the slide.
    pub fn add_text(&mut self, text: impl Into<String>) -> &mut Self {
        let has_title = self
            .slide
            .common_slide_data
            .shape_tree
            .shape
            .iter()
            .any(|s| s.non_visual_properties.c_nv_pr.name == "Title")
            || self.hyperlink_elements.iter().any(|e| e.is_title);

        let y_offset = if has_title { 1600200 } else { 274638 };

        let element = TextElement::simple(text.into(), false, 457200, y_offset, 8229600, 4525963);
        self.push_text_element(element);
        self
    }

    /// Add a text box at a specific position.
    /// Position and size are in EMUs (English Metric Units, 914400 EMUs = 1 inch).
    pub fn add_text_at(
        &mut self,
        text: impl Into<String>,
        x: i64,
        y: i64,
        width: i64,
        height: i64,
    ) -> &mut Self {
        let element = TextElement::simple(text.into(), false, x, y, width, height);
        self.push_text_element(element);
        self
    }

    /// Add a hyperlink at a specific position.
    ///
    /// Position and size are in EMUs (914400 EMUs = 1 inch).
    pub fn add_hyperlink(
        &mut self,
        text: impl Into<String>,
        url: impl Into<String>,
        x: i64,
        y: i64,
        width: i64,
        height: i64,
    ) -> &mut Self {
        let element = TextElement {
            runs: vec![TextRun {
                text: text.into(),
                hyperlink: Some(url.into()),
                tooltip: None,
                font_size: None,
                color: None,
                bold: None,
                italic: None,
                underline: false,
            }],
            is_title: false,
            x,
            y,
            width,
            height,
            fill_color: None,
            line_color: None,
            line_width: None,
        };
        self.hyperlink_elements.push(element);
        self
    }

    /// Add text with mixed content (including hyperlinks) at a specific position.
    ///
    /// Use [`TextRun`] to create runs with or without hyperlinks and formatting.
    pub fn add_text_with_runs(
        &mut self,
        runs: Vec<TextRun>,
        x: i64,
        y: i64,
        width: i64,
        height: i64,
    ) -> &mut Self {
        let element = TextElement {
            runs,
            is_title: false,
            x,
            y,
            width,
            height,
            fill_color: None,
            line_color: None,
            line_width: None,
        };
        self.push_text_element(element);
        self
    }

    /// Begin building a shape with full formatting control.
    ///
    /// Returns a [`ShapeBuilder`] that allows setting position, fill color, and border
    /// before finalizing the shape with [`ShapeBuilder::add`].
    ///
    /// # Example
    ///
    /// ```no_run
    /// # use ooxml_pml::{PresentationBuilder, writer::TextRun};
    /// # let mut pres = PresentationBuilder::new();
    /// # let slide = pres.add_slide();
    /// slide.shape(vec![TextRun::text("Hello")], 914400, 914400, 3657600, 457200)
    ///     .set_fill_color("4472C4")
    ///     .set_line_color("2F5496")
    ///     .set_line_width(12700)
    ///     .add();
    /// ```
    pub fn shape(
        &mut self,
        runs: Vec<TextRun>,
        x: i64,
        y: i64,
        width: i64,
        height: i64,
    ) -> ShapeBuilder<'_> {
        ShapeBuilder {
            slide: self,
            runs,
            is_title: false,
            x,
            y,
            width,
            height,
            fill_color: None,
            line_color: None,
            line_width: None,
        }
    }

    /// Set the slide background color.
    ///
    /// `rgb` must be a 6-character hex string (e.g. `"FF0000"` for red).
    /// Sets `cSld/bg/bgPr/solidFill/srgbClr` on the slide XML.
    ///
    /// Requires the `pml-styling` feature (included in `full`/default).
    pub fn set_background_color(&mut self, rgb: impl Into<String>) -> &mut Self {
        self.background_color = Some(rgb.into());
        self
    }

    /// Route a text element: defer if it has hyperlinks, build immediately otherwise.
    fn push_text_element(&mut self, element: TextElement) {
        if element.has_hyperlink() {
            self.hyperlink_elements.push(element);
        } else {
            let shape_id = self.next_shape_id;
            self.next_shape_id += 1;
            let shape = build_shape_impl(&element, shape_id, None);
            self.slide.common_slide_data.shape_tree.shape.push(shape);
        }
    }

    /// Add an image to the slide.
    ///
    /// Position and size are in EMUs (914400 EMUs = 1 inch).
    /// The format will be auto-detected from the image bytes.
    ///
    /// Returns `&mut Self` for chaining if successful.
    pub fn add_image(
        &mut self,
        data: impl Into<Vec<u8>>,
        x: i64,
        y: i64,
        width: i64,
        height: i64,
    ) -> &mut Self {
        let data = data.into();
        if let Some(format) = ImageFormat::from_bytes(&data) {
            self.images.push(ImageElement {
                data,
                format,
                x,
                y,
                width,
                height,
                description: None,
            });
        }
        self
    }

    /// Add an image with explicit format.
    ///
    /// Position and size are in EMUs (914400 EMUs = 1 inch).
    pub fn add_image_with_format(
        &mut self,
        data: impl Into<Vec<u8>>,
        format: ImageFormat,
        x: i64,
        y: i64,
        width: i64,
        height: i64,
    ) -> &mut Self {
        self.images.push(ImageElement {
            data: data.into(),
            format,
            x,
            y,
            width,
            height,
            description: None,
        });
        self
    }

    /// Add an image with description/alt text.
    ///
    /// Position and size are in EMUs (914400 EMUs = 1 inch).
    pub fn add_image_with_description(
        &mut self,
        data: impl Into<Vec<u8>>,
        x: i64,
        y: i64,
        width: i64,
        height: i64,
        description: impl Into<String>,
    ) -> &mut Self {
        let data = data.into();
        if let Some(format) = ImageFormat::from_bytes(&data) {
            self.images.push(ImageElement {
                data,
                format,
                x,
                y,
                width,
                height,
                description: Some(description.into()),
            });
        }
        self
    }

    /// Add an image from a file.
    ///
    /// Position and size are in EMUs (914400 EMUs = 1 inch).
    /// Returns `Ok(&mut Self)` if successful, `Err` if the file can't be read.
    pub fn add_image_from_file<P: AsRef<std::path::Path>>(
        &mut self,
        path: P,
        x: i64,
        y: i64,
        width: i64,
        height: i64,
    ) -> std::io::Result<&mut Self> {
        let data = std::fs::read(path)?;
        Ok(self.add_image(data, x, y, width, height))
    }

    /// Add a table to the slide.
    ///
    /// Position and size are in EMUs (914400 EMUs = 1 inch).
    pub fn add_table(
        &mut self,
        table: TableBuilder,
        x: i64,
        y: i64,
        width: i64,
        height: i64,
    ) -> &mut Self {
        if table.rows.is_empty() {
            return self;
        }

        let num_cols = table.col_count();
        let num_rows = table.row_count();

        let col_widths = if let Some(widths) = table.col_widths {
            widths
        } else {
            let col_width = width / num_cols as i64;
            vec![col_width; num_cols]
        };

        let row_height = height / num_rows as i64;
        let row_heights = vec![row_height; num_rows];

        let element = TableElement {
            name: table.name,
            rows: table.rows,
            col_widths,
            row_heights,
            x,
            y,
            width,
            height,
        };

        let shape_id = self.next_shape_id;
        self.next_shape_id += 1;
        let frame = build_graphic_frame(&element, shape_id);
        self.slide
            .common_slide_data
            .shape_tree
            .graphic_frame
            .push(frame);
        self
    }

    /// Check if this slide has tables.
    pub fn has_tables(&self) -> bool {
        !self
            .slide
            .common_slide_data
            .shape_tree
            .graphic_frame
            .is_empty()
    }

    /// Set a slide transition.
    pub fn set_transition(&mut self, transition: SlideTransition) -> &mut Self {
        self.apply_transition(&transition);
        self
    }

    /// Set a simple fade transition.
    pub fn set_fade_transition(&mut self) -> &mut Self {
        let t = SlideTransition::new(crate::TransitionType::Fade);
        self.apply_transition(&t);
        self
    }

    fn apply_transition(&mut self, t: &SlideTransition) {
        #[cfg(feature = "pml-transitions")]
        {
            self.slide.transition = Some(Box::new(build_slide_transition(t)));
        }
        #[cfg(not(feature = "pml-transitions"))]
        let _ = t;
    }

    /// Check if this slide has a transition.
    pub fn has_transition(&self) -> bool {
        #[cfg(feature = "pml-transitions")]
        return self.slide.transition.is_some();
        #[cfg(not(feature = "pml-transitions"))]
        return false;
    }

    /// Serialize this slide to XML bytes, resolving deferred images and hyperlinks.
    fn serialize_slide(
        &self,
        image_start_rel_id: usize,
        hyperlink_rel_ids: &std::collections::HashMap<&str, usize>,
    ) -> Result<Vec<u8>> {
        let mut slide = self.slide.clone();
        let mut next_id = self.next_shape_id;

        // Apply background color if set (requires pml-styling feature).
        #[cfg(feature = "pml-styling")]
        if let Some(ref rgb) = self.background_color {
            use crate::types::CTBackground;
            use crate::types::CTBackgroundProperties;
            use crate::types::EGBackground;

            let bg_pr = CTBackgroundProperties {
                shade_to_title: None,
                ext_lst: None,
                #[cfg(feature = "extra-attrs")]
                extra_attrs: Default::default(),
                // Inject solidFill as extra child so it roundtrips through raw XML.
                #[cfg(feature = "extra-children")]
                extra_children: parse_extra_child_xml(&format!(
                    r#"<a:solidFill xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:srgbClr val="{}"/></a:solidFill>"#,
                    rgb
                )),
            };

            let bg = CTBackground {
                #[cfg(feature = "pml-styling")]
                bw_mode: None,
                background: Some(Box::new(EGBackground::BgPr(Box::new(bg_pr)))),
                #[cfg(feature = "extra-attrs")]
                extra_attrs: Default::default(),
                #[cfg(feature = "extra-children")]
                extra_children: Default::default(),
            };
            slide.common_slide_data.bg = Some(Box::new(bg));
        }

        // Build hyperlink shapes (deferred — need rIds).
        for element in &self.hyperlink_elements {
            let shape = build_shape_impl(element, next_id, Some(hyperlink_rel_ids));
            slide.common_slide_data.shape_tree.shape.push(shape);
            next_id += 1;
        }

        // Build image pictures (deferred — need rIds).
        for (i, image) in self.images.iter().enumerate() {
            let rel_id = image_start_rel_id + i;
            let picture = build_picture(image, next_id, rel_id);
            slide.common_slide_data.shape_tree.picture.push(picture);
            next_id += 1;
        }

        // Set clrMapOvr to masterClrMapping (always required for slides).
        slide.clr_map_ovr = Some(Box::new(dml::CTColorMappingOverride {
            master_clr_mapping: Some(Box::new(dml::CTEmptyElement)),
            override_clr_mapping: None,
            extra_children: Default::default(),
        }));

        serialize_slide_xml(&slide)
    }
}

/// Builder for creating PowerPoint presentations.
#[derive(Debug)]
pub struct PresentationBuilder {
    slides: Vec<SlideBuilder>,
    slide_width: i64,
    slide_height: i64,
}

impl Default for PresentationBuilder {
    fn default() -> Self {
        Self::new()
    }
}

impl PresentationBuilder {
    /// Create a new presentation builder.
    pub fn new() -> Self {
        Self {
            slides: Vec::new(),
            slide_width: 9144000,
            slide_height: 6858000,
        }
    }

    /// Set the slide size in EMUs (914400 EMUs = 1 inch).
    pub fn set_slide_size(&mut self, width: i64, height: i64) -> &mut Self {
        self.slide_width = width;
        self.slide_height = height;
        self
    }

    /// Set slide size to widescreen (16:9).
    pub fn set_widescreen(&mut self) -> &mut Self {
        self.slide_width = 12192000;
        self.slide_height = 6858000;
        self
    }

    /// Add a new slide to the presentation.
    pub fn add_slide(&mut self) -> &mut SlideBuilder {
        self.slides.push(SlideBuilder::new());
        self.slides.last_mut().unwrap()
    }

    /// Get the number of slides.
    pub fn slide_count(&self) -> usize {
        self.slides.len()
    }

    /// Save the presentation to a file.
    pub fn save<P: AsRef<Path>>(self, path: P) -> Result<()> {
        let file = File::create(path)?;
        let writer = BufWriter::new(file);
        self.write(writer)
    }

    /// Write the presentation to a writer.
    pub fn write<W: Write + Seek>(self, writer: W) -> Result<()> {
        let mut pkg = PackageWriter::new(writer);

        pkg.add_default_content_type("rels", CT_RELATIONSHIPS);
        pkg.add_default_content_type("xml", CT_XML);

        let has_images = self.slides.iter().any(|s| s.has_images());
        if has_images {
            pkg.add_default_content_type("jpeg", CT_JPEG);
            pkg.add_default_content_type("jpg", CT_JPEG);
            pkg.add_default_content_type("png", CT_PNG);
            pkg.add_default_content_type("gif", CT_GIF);
        }

        let root_rels = format!(
            r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="{}" Target="ppt/presentation.xml"/>
</Relationships>"#,
            REL_OFFICE_DOCUMENT
        );

        // presentation.xml.rels: slides (rId1..rIdN) + slide master (rId{N+1})
        let master_rel_id = self.slides.len() + 1;
        let mut pres_rels = String::new();
        pres_rels.push_str(r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>"#);
        pres_rels.push('\n');
        pres_rels.push_str(
            r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">"#,
        );
        pres_rels.push('\n');
        for i in 0..self.slides.len() {
            let rel_id = i + 1;
            pres_rels.push_str(&format!(
                r#"  <Relationship Id="rId{}" Type="{}" Target="slides/slide{}.xml"/>"#,
                rel_id, REL_SLIDE, rel_id
            ));
            pres_rels.push('\n');
        }
        pres_rels.push_str(&format!(
            r#"  <Relationship Id="rId{}" Type="{}" Target="slideMasters/slideMaster1.xml"/>"#,
            master_rel_id, REL_SLIDE_MASTER
        ));
        pres_rels.push('\n');
        pres_rels.push_str("</Relationships>");

        let presentation_xml =
            serialize_pml_xml(&self.build_presentation(master_rel_id), "p:presentation")?;

        pkg.add_part("_rels/.rels", CT_RELATIONSHIPS, root_rels.as_bytes())?;
        pkg.add_part(
            "ppt/_rels/presentation.xml.rels",
            CT_RELATIONSHIPS,
            pres_rels.as_bytes(),
        )?;
        pkg.add_part("ppt/presentation.xml", CT_PRESENTATION, &presentation_xml)?;

        // Write slide master and its single blank layout.
        const CT_SLIDE_MASTER: &str =
            "application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml";
        const CT_SLIDE_LAYOUT: &str =
            "application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml";

        let master_rels = format!(
            r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="{}" Target="../slideLayouts/slideLayout1.xml"/>
</Relationships>"#,
            REL_SLIDE_LAYOUT
        );
        let layout_rels = format!(
            r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="{}" Target="../slideMasters/slideMaster1.xml"/>
</Relationships>"#,
            REL_SLIDE_MASTER
        );

        pkg.add_part(
            "ppt/slideMasters/slideMaster1.xml",
            CT_SLIDE_MASTER,
            build_slide_master_xml().as_bytes(),
        )?;
        pkg.add_part(
            "ppt/slideMasters/_rels/slideMaster1.xml.rels",
            CT_RELATIONSHIPS,
            master_rels.as_bytes(),
        )?;
        pkg.add_part(
            "ppt/slideLayouts/slideLayout1.xml",
            CT_SLIDE_LAYOUT,
            build_slide_layout_xml().as_bytes(),
        )?;
        pkg.add_part(
            "ppt/slideLayouts/_rels/slideLayout1.xml.rels",
            CT_RELATIONSHIPS,
            layout_rels.as_bytes(),
        )?;

        let mut global_image_num = 1;

        for (i, slide) in self.slides.iter().enumerate() {
            let slide_num = i + 1;
            let hyperlinks = slide.hyperlinks();
            let mut hyperlink_rel_ids: std::collections::HashMap<&str, usize> =
                std::collections::HashMap::new();

            // Every slide always has a rels file (at minimum for the layout reference).
            {
                let mut slide_rels = String::new();
                slide_rels.push_str(r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>"#);
                slide_rels.push('\n');
                slide_rels.push_str(
                    r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">"#,
                );
                slide_rels.push('\n');

                // rId1: slide layout (always present)
                slide_rels.push_str(&format!(
                    r#"  <Relationship Id="rId1" Type="{}" Target="../slideLayouts/slideLayout1.xml"/>"#,
                    REL_SLIDE_LAYOUT
                ));
                slide_rels.push('\n');

                let mut rel_id = 2usize;

                if slide.has_notes() {
                    slide_rels.push_str(&format!(
                        r#"  <Relationship Id="rId{}" Type="{}" Target="../notesSlides/notesSlide{}.xml"/>"#,
                        rel_id, REL_NOTES_SLIDE, slide_num
                    ));
                    slide_rels.push('\n');
                    rel_id += 1;
                }

                let image_start_rel_id = rel_id;
                for (img_idx, img) in slide.images.iter().enumerate() {
                    let ext = img.format.extension();
                    let img_rel_id = image_start_rel_id + img_idx;
                    slide_rels.push_str(&format!(
                        r#"  <Relationship Id="rId{}" Type="{}" Target="../media/image{}.{}"/>"#,
                        img_rel_id,
                        REL_IMAGE,
                        global_image_num + img_idx,
                        ext
                    ));
                    slide_rels.push('\n');
                }
                rel_id += slide.images.len();

                for url in &hyperlinks {
                    hyperlink_rel_ids.insert(url, rel_id);
                    slide_rels.push_str(&format!(
                        r#"  <Relationship Id="rId{}" Type="{}" Target="{}" TargetMode="External"/>"#,
                        rel_id, REL_HYPERLINK, escape_xml(url)
                    ));
                    slide_rels.push('\n');
                    rel_id += 1;
                }

                slide_rels.push_str("</Relationships>");
                let rels_name = format!("ppt/slides/_rels/slide{}.xml.rels", slide_num);
                pkg.add_part(&rels_name, CT_RELATIONSHIPS, slide_rels.as_bytes())?;

                for (img_idx, img) in slide.images.iter().enumerate() {
                    let ext = img.format.extension();
                    let img_path = format!("ppt/media/image{}.{}", global_image_num + img_idx, ext);
                    pkg.add_part(&img_path, img.format.content_type(), &img.data)?;
                }
            }

            // rId1=layout, rId2=notes (if any), then images start.
            let image_start_rel_id = if slide.has_notes() { 3 } else { 2 };
            let slide_xml = slide.serialize_slide(image_start_rel_id, &hyperlink_rel_ids)?;
            let part_name = format!("ppt/slides/slide{}.xml", slide_num);
            pkg.add_part(&part_name, CT_SLIDE, &slide_xml)?;

            global_image_num += slide.images.len();

            if slide.has_notes() {
                let notes_xml =
                    serialize_pml_xml(&Self::build_notes_slide(slide, slide_num), "p:notes")?;
                let notes_name = format!("ppt/notesSlides/notesSlide{}.xml", slide_num);
                pkg.add_part(&notes_name, CT_NOTES_SLIDE, &notes_xml)?;
            }
        }

        pkg.finish()?;
        Ok(())
    }

    /// Serialize presentation.xml
    fn build_presentation(&self, master_rel_id: usize) -> types::Presentation {
        let sld_id_lst = types::SlideIdList {
            sld_id: self
                .slides
                .iter()
                .enumerate()
                .map(|(i, _)| {
                    let mut entry = types::CTSlideIdListEntry {
                        id: 256 + i as u32,
                        ext_lst: None,
                        extra_attrs: Default::default(),
                        extra_children: Default::default(),
                    };
                    entry
                        .extra_attrs
                        .insert("r:id".to_string(), format!("rId{}", i + 1));
                    entry
                })
                .collect(),
            extra_children: Default::default(),
        };

        // Build the slide master id list (one entry for our minimal master).
        let mut master_entry = types::CTSlideMasterIdListEntry {
            id: Some(2147483647),
            ext_lst: None,
            extra_attrs: Default::default(),
            extra_children: Default::default(),
        };
        master_entry
            .extra_attrs
            .insert("r:id".to_string(), format!("rId{}", master_rel_id));
        let sld_master_id_lst = types::CTSlideMasterIdList {
            sld_master_id: vec![master_entry],
            #[cfg(feature = "extra-children")]
            extra_children: Default::default(),
        };

        types::Presentation {
            first_slide_num: None,
            remove_personal_info_on_save: None,
            compat_mode: None,
            bookmark_id_seed: None,
            conformance: None,
            sld_master_id_lst: Some(Box::new(sld_master_id_lst)),
            sld_id_lst: Some(Box::new(sld_id_lst)),
            sld_sz: Some(Box::new(types::CTSlideSize {
                cx: self.slide_width as i32,
                cy: self.slide_height as i32,
                r#type: None,
                extra_attrs: Default::default(),
            })),
            #[cfg(feature = "pml-notes")]
            notes_master_id_lst: None,
            #[cfg(feature = "pml-notes")]
            notes_sz: Box::new(dml::PositiveSize2D {
                cx: self.slide_width,
                cy: self.slide_height,
                extra_attrs: Default::default(),
            }),
            #[cfg(feature = "pml-masters")]
            handout_master_id_lst: None,
            cust_show_lst: None,
            modify_verifier: None,
            #[cfg(feature = "pml-styling")]
            server_zoom: None,
            #[cfg(feature = "pml-styling")]
            show_special_pls_on_title_sld: None,
            #[cfg(feature = "pml-styling")]
            rtl: None,
            #[cfg(feature = "pml-styling")]
            strict_first_and_last_chars: None,
            #[cfg(feature = "pml-styling")]
            embed_true_type_fonts: None,
            #[cfg(feature = "pml-styling")]
            save_subset_fonts: None,
            #[cfg(feature = "pml-styling")]
            auto_compress_pictures: None,
            #[cfg(feature = "pml-styling")]
            embedded_font_lst: None,
            #[cfg(feature = "pml-styling")]
            kinsoku: None,
            #[cfg(feature = "pml-styling")]
            default_text_style: None,
            #[cfg(feature = "pml-external")]
            smart_tags: None,
            #[cfg(feature = "pml-external")]
            cust_data_lst: None,
            #[cfg(feature = "pml-media")]
            photo_album: None,
            #[cfg(feature = "pml-extensions")]
            ext_lst: None,
            extra_attrs: Default::default(),
            extra_children: Default::default(),
        }
    }

    /// Build a `types::NotesSlide` for a slide's speaker notes.
    fn build_notes_slide(slide: &SlideBuilder, slide_num: usize) -> types::NotesSlide {
        let notes_text = slide.notes.as_deref().unwrap_or("");

        // Slide image placeholder shape (id=2).
        let sld_img_shape = types::Shape {
            use_bg_fill: None,
            non_visual_properties: Box::new(types::ShapeNonVisual {
                c_nv_pr: make_cnv_pr(2, &format!("Slide Image Placeholder {}", slide_num)),
                c_nv_sp_pr: Box::new(dml::CTNonVisualDrawingShapeProps {
                    sp_locks: Some(Box::new(dml::CTShapeLocking {
                        no_grp: Some(true),
                        no_rot: Some(true),
                        no_change_aspect: Some(true),
                        ..Default::default()
                    })),
                    ..Default::default()
                }),
                nv_pr: Box::new(types::CTApplicationNonVisualDrawingProps {
                    ph: Some(Box::new(types::CTPlaceholder {
                        r#type: Some(types::STPlaceholderType::SldImg),
                        ..Default::default()
                    })),
                    ..Default::default()
                }),
                extra_children: Default::default(),
            }),
            shape_properties: Box::new(dml::CTShapeProperties::default()),
            style: None,
            text_body: None,
            ext_lst: None,
            extra_attrs: Default::default(),
            extra_children: Default::default(),
        };

        // Notes text paragraphs.
        let paragraphs: Vec<dml::TextParagraph> = if notes_text.is_empty() {
            vec![dml::TextParagraph::default()]
        } else {
            notes_text
                .lines()
                .map(|line| dml::TextParagraph {
                    text_run: vec![dml::EGTextRun::R(Box::new(dml::TextRun {
                        r_pr: Some(Box::new(dml::TextCharacterProperties {
                            lang: Some("en-US".to_string()),
                            ..Default::default()
                        })),
                        t: line.to_string(),
                        extra_children: Default::default(),
                    }))],
                    ..Default::default()
                })
                .collect()
        };

        // Notes body placeholder shape (id=3).
        let notes_shape = types::Shape {
            use_bg_fill: None,
            non_visual_properties: Box::new(types::ShapeNonVisual {
                c_nv_pr: make_cnv_pr(3, "Notes Placeholder"),
                c_nv_sp_pr: Box::new(dml::CTNonVisualDrawingShapeProps {
                    sp_locks: Some(Box::new(dml::CTShapeLocking {
                        no_grp: Some(true),
                        ..Default::default()
                    })),
                    ..Default::default()
                }),
                nv_pr: Box::new(types::CTApplicationNonVisualDrawingProps {
                    ph: Some(Box::new(types::CTPlaceholder {
                        r#type: Some(types::STPlaceholderType::Body),
                        idx: Some(1),
                        ..Default::default()
                    })),
                    ..Default::default()
                }),
                extra_children: Default::default(),
            }),
            shape_properties: Box::new(dml::CTShapeProperties::default()),
            style: None,
            text_body: Some(Box::new(dml::TextBody {
                body_pr: Box::default(),
                lst_style: Some(Box::new(dml::CTTextListStyle::default())),
                p: paragraphs,
                extra_children: Default::default(),
            })),
            ext_lst: None,
            extra_attrs: Default::default(),
            extra_children: Default::default(),
        };

        let grp_xfrm = dml::CTGroupTransform2D {
            offset: Some(Box::new(dml::Point2D {
                x: "0".to_string(),
                y: "0".to_string(),
                extra_attrs: Default::default(),
            })),
            extents: Some(Box::new(dml::PositiveSize2D {
                cx: 0,
                cy: 0,
                extra_attrs: Default::default(),
            })),
            child_offset: Some(Box::new(dml::Point2D {
                x: "0".to_string(),
                y: "0".to_string(),
                extra_attrs: Default::default(),
            })),
            child_extents: Some(Box::new(dml::PositiveSize2D {
                cx: 0,
                cy: 0,
                extra_attrs: Default::default(),
            })),
            ..Default::default()
        };

        let shape_tree = Box::new(types::GroupShape {
            non_visual_group_properties: Box::new(types::CTGroupShapeNonVisual {
                c_nv_pr: make_cnv_pr(1, ""),
                c_nv_grp_sp_pr: Box::new(dml::CTNonVisualGroupDrawingShapeProps {
                    grp_sp_locks: None,
                    ext_lst: None,
                    extra_children: Default::default(),
                }),
                nv_pr: make_nv_pr(),
                extra_children: Default::default(),
            }),
            grp_sp_pr: Box::new(dml::CTGroupShapeProperties {
                transform: Some(Box::new(grp_xfrm)),
                ..Default::default()
            }),
            shape: vec![sld_img_shape, notes_shape],
            picture: Vec::new(),
            connector: Vec::new(),
            group_shape: Vec::new(),
            graphic_frame: Vec::new(),
            content_part: Vec::new(),
            ext_lst: None,
            extra_children: Default::default(),
        });

        types::NotesSlide {
            common_slide_data: Box::new(types::CommonSlideData {
                name: None,
                bg: None,
                shape_tree,
                cust_data_lst: None,
                controls: None,
                ext_lst: None,
                extra_attrs: Default::default(),
                extra_children: Default::default(),
            }),
            #[cfg(feature = "pml-notes")]
            show_master_sp: None,
            #[cfg(feature = "pml-notes")]
            show_master_ph_anim: None,
            #[cfg(feature = "pml-styling")]
            clr_map_ovr: None,
            #[cfg(feature = "pml-extensions")]
            ext_lst: None,
            extra_attrs: Default::default(),
            extra_children: Default::default(),
        }
    }
}

/// Escape XML special characters.
fn escape_xml(s: &str) -> String {
    s.replace('&', "&amp;")
        .replace('<', "&lt;")
        .replace('>', "&gt;")
        .replace('"', "&quot;")
        .replace('\'', "&apos;")
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::ext::ShapeExt;
    use ooxml_dml::TextParagraphExt;

    #[test]
    fn test_presentation_builder() {
        let mut pres = PresentationBuilder::new();
        let slide = pres.add_slide();
        slide.add_title("Test Title");
        slide.add_text("Test content");

        assert_eq!(pres.slide_count(), 1);
    }

    #[test]
    fn test_roundtrip_simple() {
        use std::io::Cursor;

        let mut pres = PresentationBuilder::new();
        let slide = pres.add_slide();
        slide.add_title("Hello World");
        slide.add_text("This is a test presentation");

        let mut buffer = Cursor::new(Vec::new());
        pres.write(&mut buffer).unwrap();

        buffer.set_position(0);
        let mut presentation = crate::Presentation::from_reader(buffer).unwrap();
        assert_eq!(presentation.slide_count(), 1);

        let _read_slide = presentation.slide(0).unwrap();
    }

    #[test]
    fn test_table_builder() {
        let table = TableBuilder::new()
            .name("My Table")
            .add_row(["A", "B", "C"])
            .add_row(["1", "2", "3"]);

        assert_eq!(table.row_count(), 2);
        assert_eq!(table.col_count(), 3);
    }

    #[cfg(feature = "dml-tables")]
    #[test]
    fn test_roundtrip_with_table() {
        use ooxml_dml::TableCellExt;
        use std::io::Cursor;

        let mut pres = PresentationBuilder::new();
        let slide = pres.add_slide();
        slide.add_title("Table Test");

        let table = TableBuilder::new()
            .name("Data Table")
            .add_row(["Name", "Value"])
            .add_row(["Alpha", "100"])
            .add_row(["Beta", "200"]);

        slide.add_table(table, 914400, 1828800, 7315200, 1828800);

        let mut buffer = Cursor::new(Vec::new());
        pres.write(&mut buffer).unwrap();

        buffer.set_position(0);
        let mut presentation = crate::Presentation::from_reader(buffer).unwrap();
        assert_eq!(presentation.slide_count(), 1);

        let read_slide = presentation.slide(0).unwrap();
        assert!(read_slide.has_tables());
        assert_eq!(read_slide.table_count(), 1);

        let table = read_slide.table(0).unwrap();
        assert_eq!(table.row_count(), 3);
        assert_eq!(table.col_count(), 2);

        assert_eq!(table.cell(0, 0).unwrap().text(), "Name");
        assert_eq!(table.cell(0, 1).unwrap().text(), "Value");
        assert_eq!(table.cell(1, 0).unwrap().text(), "Alpha");
        assert_eq!(table.cell(2, 1).unwrap().text(), "200");
    }

    #[test]
    fn test_roundtrip_with_hyperlink() {
        use std::io::Cursor;

        let mut pres = PresentationBuilder::new();
        let slide = pres.add_slide();
        slide.add_title("Hyperlink Test");
        slide.add_hyperlink(
            "Click here to visit Rust",
            "https://www.rust-lang.org",
            457200,
            1600200,
            8229600,
            457200,
        );

        let mut buffer = Cursor::new(Vec::new());
        pres.write(&mut buffer).unwrap();

        buffer.set_position(0);
        let mut presentation = crate::Presentation::from_reader(buffer).unwrap();
        assert_eq!(presentation.slide_count(), 1);

        let read_slide = presentation.slide(0).unwrap();
        assert!(read_slide.has_hyperlinks());

        let links = read_slide.hyperlinks();
        assert_eq!(links.len(), 1);
        assert_eq!(links[0].text, "Click here to visit Rust");

        let url = presentation
            .resolve_hyperlink(&read_slide, &links[0].rel_id)
            .unwrap();
        assert_eq!(url, "https://www.rust-lang.org");
    }

    #[test]
    fn test_text_with_mixed_runs() {
        use std::io::Cursor;

        let mut pres = PresentationBuilder::new();
        let slide = pres.add_slide();
        slide.add_title("Mixed Content");
        slide.add_text_with_runs(
            vec![
                TextRun::text("Read the "),
                TextRun::hyperlink("documentation", "https://docs.rust-lang.org"),
                TextRun::text(" for more info."),
            ],
            457200,
            1600200,
            8229600,
            457200,
        );

        let mut buffer = Cursor::new(Vec::new());
        pres.write(&mut buffer).unwrap();

        buffer.set_position(0);
        let mut presentation = crate::Presentation::from_reader(buffer).unwrap();
        let read_slide = presentation.slide(0).unwrap();

        assert!(read_slide.has_hyperlinks());
        let links = read_slide.hyperlinks();
        assert_eq!(links.len(), 1);
        assert_eq!(links[0].text, "documentation");
    }

    #[test]
    fn test_roundtrip_slide_master_layout() {
        use std::io::Cursor;

        // Verify that written PPTX files contain a slide master and layout,
        // so they render correctly in PowerPoint instead of using fallbacks.
        let mut pres = PresentationBuilder::new();
        let slide = pres.add_slide();
        slide.add_title("Master Test");
        slide.add_text("Slide with layout");

        let mut buffer = Cursor::new(Vec::new());
        pres.write(&mut buffer).unwrap();

        buffer.set_position(0);
        let presentation = crate::Presentation::from_reader(buffer).unwrap();

        // Exactly one slide master should be present.
        assert_eq!(presentation.slide_masters().len(), 1);

        // The master should reference at least one layout.
        let master = &presentation.slide_masters()[0];
        assert_eq!(master.layout_count(), 1);

        // Layouts should be loaded too.
        assert_eq!(presentation.slide_layouts().len(), 1);
    }

    #[cfg(feature = "pml-transitions")]
    #[test]
    fn test_roundtrip_with_transition() {
        use std::io::Cursor;

        let mut pres = PresentationBuilder::new();
        let slide = pres.add_slide();
        slide.add_title("Transition Test");
        slide.set_transition(
            SlideTransition::new(crate::TransitionType::Fade)
                .with_speed(crate::TransitionSpeed::Slow)
                .with_advance_after(3000),
        );

        let mut buffer = Cursor::new(Vec::new());
        pres.write(&mut buffer).unwrap();

        buffer.set_position(0);
        let mut presentation = crate::Presentation::from_reader(buffer).unwrap();
        let read_slide = presentation.slide(0).unwrap();

        let transition = read_slide.transition();
        assert!(transition.is_some());
        let t = transition.unwrap();
        assert_eq!(t.transition_type, Some(crate::TransitionType::Fade));
        assert_eq!(t.speed, crate::TransitionSpeed::Slow);
        assert_eq!(t.advance_time_ms, Some(3000));
    }

    // -------------------------------------------------------------------------
    // Text run formatting tests (features 1–3, 7)
    // -------------------------------------------------------------------------

    #[test]
    fn test_text_run_font_size() {
        use ooxml_dml::TextRunExt;
        use std::io::Cursor;

        let mut pres = PresentationBuilder::new();
        let slide = pres.add_slide();
        slide.add_text_with_runs(
            vec![TextRun::text("Big text").set_font_size(36.0)],
            457200,
            914400,
            8229600,
            914400,
        );

        let mut buffer = Cursor::new(Vec::new());
        pres.write(&mut buffer).unwrap();

        buffer.set_position(0);
        let mut presentation = crate::Presentation::from_reader(buffer).unwrap();
        let read_slide = presentation.slide(0).unwrap();

        let shapes = read_slide.shapes();
        assert!(!shapes.is_empty());
        let shape = &shapes[0];
        let text_body = shape.text_body().expect("shape has text body");
        let para = &text_body.p[0];
        let runs = para.runs();
        assert!(!runs.is_empty());
        // 36pt = 3600 hundredths-of-a-point
        assert_eq!(runs[0].font_size(), Some(3600));
    }

    #[test]
    fn test_text_run_color() {
        use std::io::Cursor;

        let mut pres = PresentationBuilder::new();
        let slide = pres.add_slide();
        slide.add_text_with_runs(
            vec![TextRun::text("Red text").set_color("FF0000")],
            457200,
            914400,
            8229600,
            914400,
        );

        let mut buffer = Cursor::new(Vec::new());
        pres.write(&mut buffer).unwrap();

        buffer.set_position(0);
        let mut presentation = crate::Presentation::from_reader(buffer).unwrap();
        let read_slide = presentation.slide(0).unwrap();

        let shapes = read_slide.shapes();
        let shape = &shapes[0];
        let text_body = shape.text_body().expect("shape has text body");
        let para = &text_body.p[0];
        let runs = para.runs();
        assert!(!runs.is_empty());
        let r_pr = runs[0].r_pr.as_ref().expect("run has r_pr");
        // Verify solidFill/srgbClr is set.
        let fill = r_pr
            .fill_properties
            .as_ref()
            .expect("run has fill_properties");
        if let ooxml_dml::types::EGFillProperties::SolidFill(solid) = fill.as_ref()
            && let Some(color) = &solid.color_choice
            && let ooxml_dml::types::EGColorChoice::SrgbClr(srgb) = color.as_ref()
        {
            // FF0000 = [0xFF, 0x00, 0x00]
            assert_eq!(srgb.value, vec![0xFF, 0x00, 0x00]);
            return;
        }
        panic!("Expected solidFill/srgbClr with FF0000");
    }

    #[test]
    fn test_text_run_bold_italic_underline() {
        use ooxml_dml::ext::TextRunExt;
        use std::io::Cursor;

        let mut pres = PresentationBuilder::new();
        let slide = pres.add_slide();
        slide.add_text_with_runs(
            vec![
                TextRun::text("Formatted")
                    .set_bold(true)
                    .set_italic(true)
                    .set_underline(true),
            ],
            457200,
            914400,
            8229600,
            914400,
        );

        let mut buffer = Cursor::new(Vec::new());
        pres.write(&mut buffer).unwrap();

        buffer.set_position(0);
        let mut presentation = crate::Presentation::from_reader(buffer).unwrap();
        let read_slide = presentation.slide(0).unwrap();

        let shapes = read_slide.shapes();
        let shape = &shapes[0];
        let text_body = shape.text_body().expect("shape has text body");
        let para = &text_body.p[0];
        let runs = para.runs();
        assert!(!runs.is_empty());
        assert!(runs[0].is_bold(), "expected bold");
        assert!(runs[0].is_italic(), "expected italic");
        assert!(runs[0].is_underlined(), "expected underline");
    }

    #[test]
    fn test_hyperlink_tooltip() {
        use std::io::Cursor;

        let mut pres = PresentationBuilder::new();
        let slide = pres.add_slide();
        slide.add_text_with_runs(
            vec![
                TextRun::hyperlink("Hover me", "https://example.com")
                    .set_tooltip("This is a tooltip"),
            ],
            457200,
            914400,
            8229600,
            914400,
        );

        let mut buffer = Cursor::new(Vec::new());
        pres.write(&mut buffer).unwrap();

        buffer.set_position(0);
        let mut presentation = crate::Presentation::from_reader(buffer).unwrap();
        let read_slide = presentation.slide(0).unwrap();

        let shapes = read_slide.shapes();
        // hyperlink shapes are pushed via hyperlink_elements
        // After write/read they appear in the shape tree
        let mut found_tooltip = false;
        for shape in shapes {
            if let Some(tb) = shape.text_body() {
                for para in &tb.p {
                    for run in para.runs() {
                        if let Some(hlink) = &run.r_pr.as_ref().and_then(|p| p.hlink_click.as_ref())
                            && hlink.tooltip.as_deref() == Some("This is a tooltip")
                        {
                            found_tooltip = true;
                        }
                    }
                }
            }
        }
        assert!(found_tooltip, "Expected tooltip on hyperlink run");
    }

    // -------------------------------------------------------------------------
    // Shape position and fill/line tests (features 4–6)
    // -------------------------------------------------------------------------

    #[test]
    fn test_shape_position_and_size() {
        use std::io::Cursor;

        let mut pres = PresentationBuilder::new();
        let slide = pres.add_slide();
        slide
            .shape(
                vec![TextRun::text("Positioned")],
                914400,
                457200,
                3657600,
                914400,
            )
            .add();

        let mut buffer = Cursor::new(Vec::new());
        pres.write(&mut buffer).unwrap();

        buffer.set_position(0);
        let mut presentation = crate::Presentation::from_reader(buffer).unwrap();
        let read_slide = presentation.slide(0).unwrap();

        let shapes = read_slide.shapes();
        assert!(!shapes.is_empty());
        let sp_pr = &shapes[0].shape_properties;
        let xfrm = sp_pr.transform.as_ref().expect("shape has transform");
        let off = xfrm.offset.as_ref().expect("xfrm has offset");
        assert_eq!(off.x, "914400");
        assert_eq!(off.y, "457200");
        let ext = xfrm.extents.as_ref().expect("xfrm has extents");
        assert_eq!(ext.cx, 3657600);
        assert_eq!(ext.cy, 914400);
    }

    #[test]
    fn test_shape_fill_color() {
        use std::io::Cursor;

        let mut pres = PresentationBuilder::new();
        let slide = pres.add_slide();
        slide
            .shape(
                vec![TextRun::text("Filled")],
                457200,
                457200,
                3657600,
                914400,
            )
            .set_fill_color("4472C4")
            .add();

        let mut buffer = Cursor::new(Vec::new());
        pres.write(&mut buffer).unwrap();

        buffer.set_position(0);
        let mut presentation = crate::Presentation::from_reader(buffer).unwrap();
        let read_slide = presentation.slide(0).unwrap();

        let shapes = read_slide.shapes();
        assert!(!shapes.is_empty());
        let sp_pr = &shapes[0].shape_properties;
        let fill = sp_pr.fill_properties.as_ref().expect("shape has fill");
        if let ooxml_dml::types::EGFillProperties::SolidFill(solid) = fill.as_ref()
            && let Some(color) = &solid.color_choice
            && let ooxml_dml::types::EGColorChoice::SrgbClr(srgb) = color.as_ref()
        {
            // 4472C4 = [0x44, 0x72, 0xC4]
            assert_eq!(srgb.value, vec![0x44, 0x72, 0xC4]);
            return;
        }
        panic!("Expected shape solidFill/srgbClr with 4472C4");
    }

    #[test]
    fn test_shape_line_color_and_width() {
        use std::io::Cursor;

        let mut pres = PresentationBuilder::new();
        let slide = pres.add_slide();
        slide
            .shape(
                vec![TextRun::text("Bordered")],
                457200,
                457200,
                3657600,
                914400,
            )
            .set_line_color("FF0000")
            .set_line_width(38100) // 3pt line
            .add();

        let mut buffer = Cursor::new(Vec::new());
        pres.write(&mut buffer).unwrap();

        buffer.set_position(0);
        let mut presentation = crate::Presentation::from_reader(buffer).unwrap();
        let read_slide = presentation.slide(0).unwrap();

        let shapes = read_slide.shapes();
        assert!(!shapes.is_empty());
        let sp_pr = &shapes[0].shape_properties;
        let line = sp_pr.line.as_ref().expect("shape has line");
        assert_eq!(line.width, Some(38100));
        let line_fill = line.line_fill_properties.as_ref().expect("line has fill");
        if let ooxml_dml::types::EGLineFillProperties::SolidFill(solid) = line_fill.as_ref()
            && let Some(color) = &solid.color_choice
            && let ooxml_dml::types::EGColorChoice::SrgbClr(srgb) = color.as_ref()
        {
            assert_eq!(srgb.value, vec![0xFF, 0x00, 0x00]);
            return;
        }
        panic!("Expected line solidFill/srgbClr with FF0000");
    }

    // -------------------------------------------------------------------------
    // Slide background color test (feature 8)
    // -------------------------------------------------------------------------

    #[cfg(feature = "pml-styling")]
    #[test]
    fn test_slide_background_color() {
        use std::io::Cursor;

        let mut pres = PresentationBuilder::new();
        let slide = pres.add_slide();
        slide.add_title("BG Test");
        slide.set_background_color("1F497D");

        let mut buffer = Cursor::new(Vec::new());
        pres.write(&mut buffer).unwrap();

        // Verify the PPTX roundtrips (reads back without error).
        buffer.set_position(0);
        let mut presentation = crate::Presentation::from_reader(buffer.clone()).unwrap();
        assert_eq!(presentation.slide_count(), 1);
        let _slide = presentation.slide(0).unwrap();

        // Verify the raw PPTX bytes contain the background color and bgPr element.
        // We use ooxml-opc to read the slide XML part directly.
        buffer.set_position(0);
        let mut package = ooxml_opc::Package::open(buffer).unwrap();
        let slide_xml = package.read_part("ppt/slides/slide1.xml").unwrap();
        let slide_xml_str = String::from_utf8_lossy(&slide_xml);

        assert!(
            slide_xml_str.contains("1F497D"),
            "Slide XML should contain background color hex; got snippet: {}",
            &slide_xml_str[..slide_xml_str.len().min(800)]
        );
        assert!(
            slide_xml_str.contains("bgPr") || slide_xml_str.contains("solidFill"),
            "Slide XML should contain bgPr or solidFill"
        );
    }
}
