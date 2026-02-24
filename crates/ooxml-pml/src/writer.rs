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
const CT_NOTES_MASTER: &str =
    "application/vnd.openxmlformats-officedocument.presentationml.notesMaster+xml";
const CT_CHART: &str = "application/vnd.openxmlformats-officedocument.drawingml.chart+xml";
const CT_DIAGRAM_DATA: &str =
    "application/vnd.openxmlformats-officedocument.drawingml.diagramData+xml";
const CT_DIAGRAM_LAYOUT: &str =
    "application/vnd.openxmlformats-officedocument.drawingml.diagramLayout+xml";
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
const REL_NOTES_MASTER: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesMaster";
const REL_CHART: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart";
const REL_DIAGRAM_DATA: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramData";
const REL_DIAGRAM_LAYOUT: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramLayout";

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

/// An opaque handle for a slide master added to a `PresentationBuilder`.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct MasterId(usize);

/// An opaque handle for a slide layout within a master.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct LayoutId {
    master_idx: usize,
    layout_idx: usize,
}

/// Layout type for a slide layout.
///
/// Each variant maps to an OOXML `type` attribute on `<p:sldLayout>` and
/// controls which placeholder shapes appear in the layout XML.
/// ECMA-376 Part 1, Section 19.7.14 (ST_SlideLayoutType).
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum SlideLayoutType {
    /// Blank layout — no placeholders.
    Blank,
    /// Title slide — centred title + subtitle placeholders.
    TitleSlide,
    /// Title and content — title + content area.
    TitleContent,
    /// Two content areas — title + two side-by-side content areas.
    TwoContent,
    /// Section header — section divider with title + text.
    SectionHeader,
}

impl SlideLayoutType {
    fn xml_type(&self) -> &'static str {
        match self {
            SlideLayoutType::Blank => "blank",
            SlideLayoutType::TitleSlide => "title",
            SlideLayoutType::TitleContent => "obj",
            SlideLayoutType::TwoContent => "twoObj",
            SlideLayoutType::SectionHeader => "secHead",
        }
    }
}

/// Configuration for a slide layout added to a master.
#[derive(Debug, Clone)]
pub struct SlideLayoutConfig {
    /// Display name for the layout.
    pub name: String,
    /// Layout type controlling placeholder shapes.
    pub layout_type: SlideLayoutType,
}

/// Configuration for a slide master added to the presentation.
#[derive(Debug, Clone)]
pub struct SlideMasterConfig {
    /// Optional background fill color (6-hex RGB, e.g. `"1F497D"`).
    pub background_color: Option<String>,
    /// Optional theme name (stored in the master XML comment for identification).
    pub theme_name: Option<String>,
    /// Layouts associated with this master.
    pub layouts: Vec<SlideLayoutConfig>,
}

/// Internal representation of a configured slide master.
#[derive(Debug)]
struct SlideMasterEntry {
    config: SlideMasterConfig,
}

/// Build slide master XML for a given master entry.
///
/// Extends the minimal master with an optional background fill color.
fn build_master_xml(entry: &SlideMasterEntry, layout_count: usize) -> String {
    let mut xml = String::new();
    xml.push_str(r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>"#);
    xml.push('\n');
    xml.push_str(r#"<p:sldMaster xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">"#);
    xml.push_str("<p:cSld>");
    if let Some(ref rgb) = entry.config.background_color {
        xml.push_str(r#"<p:bg><p:bgPr>"#);
        xml.push_str(r#"<a:solidFill>"#);
        xml.push_str(&format!(r#"<a:srgbClr val="{}"/>"#, rgb));
        xml.push_str(r#"</a:solidFill>"#);
        xml.push_str(r#"<a:effectLst/>"#);
        xml.push_str(r#"</p:bgPr></p:bg>"#);
    }
    xml.push_str(r#"<p:spTree>"#);
    xml.push_str(r#"<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>"#);
    xml.push_str(r#"<p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>"#);
    xml.push_str(r#"</p:spTree></p:cSld>"#);
    xml.push_str(r#"<p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/>"#);
    xml.push_str("<p:sldLayoutIdLst>");
    for i in 0..layout_count {
        let id = 2147483648u64 + i as u64;
        xml.push_str(&format!(
            r#"<p:sldLayoutId id="{}" r:id="rId{}"/>"#,
            id,
            i + 1
        ));
    }
    xml.push_str("</p:sldLayoutIdLst>");
    xml.push_str("</p:sldMaster>");
    xml
}

/// Build slide layout XML for a given layout config.
fn build_layout_xml(layout: &SlideLayoutConfig) -> String {
    let layout_type = layout.layout_type.xml_type();
    let mut xml = String::new();
    xml.push_str(r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>"#);
    xml.push('\n');
    xml.push_str(&format!(
        r#"<p:sldLayout xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" type="{}" preserve="1">"#,
        layout_type
    ));
    xml.push_str(&format!(r#"<p:cSld name="{}">"#, escape_xml(&layout.name)));
    xml.push_str("<p:spTree>");
    xml.push_str(r#"<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>"#);
    xml.push_str(r#"<p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>"#);

    match layout.layout_type {
        SlideLayoutType::TitleSlide => {
            // Centre-title placeholder (id=2)
            xml.push_str(r#"<p:sp><p:nvSpPr><p:cNvPr id="2" name="Title 1"/>"#);
            xml.push_str(r#"<p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr>"#);
            xml.push_str(r#"<p:nvPr><p:ph type="ctrTitle"/></p:nvPr></p:nvSpPr>"#);
            xml.push_str(r#"<p:spPr/><p:txBody><a:bodyPr/><a:lstStyle/><a:p/></p:txBody></p:sp>"#);
            // Subtitle placeholder (id=3, idx=1)
            xml.push_str(r#"<p:sp><p:nvSpPr><p:cNvPr id="3" name="Subtitle 2"/>"#);
            xml.push_str(r#"<p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr>"#);
            xml.push_str(r#"<p:nvPr><p:ph type="subTitle" idx="1"/></p:nvPr></p:nvSpPr>"#);
            xml.push_str(r#"<p:spPr/><p:txBody><a:bodyPr/><a:lstStyle/><a:p/></p:txBody></p:sp>"#);
        }
        SlideLayoutType::TitleContent => {
            // Title placeholder (id=2)
            xml.push_str(r#"<p:sp><p:nvSpPr><p:cNvPr id="2" name="Title 1"/>"#);
            xml.push_str(r#"<p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr>"#);
            xml.push_str(r#"<p:nvPr><p:ph type="title"/></p:nvPr></p:nvSpPr>"#);
            xml.push_str(r#"<p:spPr/><p:txBody><a:bodyPr/><a:lstStyle/><a:p/></p:txBody></p:sp>"#);
            // Content placeholder (id=3, idx=1)
            xml.push_str(r#"<p:sp><p:nvSpPr><p:cNvPr id="3" name="Content Placeholder 2"/>"#);
            xml.push_str(r#"<p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr>"#);
            xml.push_str(r#"<p:nvPr><p:ph idx="1"/></p:nvPr></p:nvSpPr>"#);
            xml.push_str(r#"<p:spPr/><p:txBody><a:bodyPr/><a:lstStyle/><a:p/></p:txBody></p:sp>"#);
        }
        SlideLayoutType::TwoContent => {
            // Title placeholder (id=2)
            xml.push_str(r#"<p:sp><p:nvSpPr><p:cNvPr id="2" name="Title 1"/>"#);
            xml.push_str(r#"<p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr>"#);
            xml.push_str(r#"<p:nvPr><p:ph type="title"/></p:nvPr></p:nvSpPr>"#);
            xml.push_str(r#"<p:spPr/><p:txBody><a:bodyPr/><a:lstStyle/><a:p/></p:txBody></p:sp>"#);
            // Left content placeholder (id=3, idx=1)
            xml.push_str(r#"<p:sp><p:nvSpPr><p:cNvPr id="3" name="Content Placeholder 2"/>"#);
            xml.push_str(r#"<p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr>"#);
            xml.push_str(r#"<p:nvPr><p:ph idx="1"/></p:nvPr></p:nvSpPr>"#);
            xml.push_str(r#"<p:spPr/><p:txBody><a:bodyPr/><a:lstStyle/><a:p/></p:txBody></p:sp>"#);
            // Right content placeholder (id=4, idx=2)
            xml.push_str(r#"<p:sp><p:nvSpPr><p:cNvPr id="4" name="Content Placeholder 3"/>"#);
            xml.push_str(r#"<p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr>"#);
            xml.push_str(r#"<p:nvPr><p:ph idx="2"/></p:nvPr></p:nvSpPr>"#);
            xml.push_str(r#"<p:spPr/><p:txBody><a:bodyPr/><a:lstStyle/><a:p/></p:txBody></p:sp>"#);
        }
        SlideLayoutType::SectionHeader => {
            // Title placeholder (id=2)
            xml.push_str(r#"<p:sp><p:nvSpPr><p:cNvPr id="2" name="Title 1"/>"#);
            xml.push_str(r#"<p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr>"#);
            xml.push_str(r#"<p:nvPr><p:ph type="title"/></p:nvPr></p:nvSpPr>"#);
            xml.push_str(r#"<p:spPr/><p:txBody><a:bodyPr/><a:lstStyle/><a:p/></p:txBody></p:sp>"#);
            // Body placeholder (id=3, idx=1)
            xml.push_str(r#"<p:sp><p:nvSpPr><p:cNvPr id="3" name="Text Placeholder 2"/>"#);
            xml.push_str(r#"<p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr>"#);
            xml.push_str(r#"<p:nvPr><p:ph type="body" idx="1"/></p:nvPr></p:nvSpPr>"#);
            xml.push_str(r#"<p:spPr/><p:txBody><a:bodyPr/><a:lstStyle/><a:p/></p:txBody></p:sp>"#);
        }
        SlideLayoutType::Blank => {
            // No placeholders for blank layout.
        }
    }

    xml.push_str("</p:spTree>");
    xml.push_str("</p:cSld>");
    xml.push_str("</p:sldLayout>");
    xml
}

/// Build the minimal notes master XML.
fn build_notes_master_xml() -> &'static str {
    r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:notesMaster xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
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
      <p:sp>
        <p:nvSpPr>
          <p:cNvPr id="2" name=""/>
          <p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr>
          <p:nvPr><p:ph type="body" idx="1"/></p:nvPr>
        </p:nvSpPr>
        <p:spPr/>
        <p:txBody><a:bodyPr/><a:lstStyle/><a:p/></p:txBody>
      </p:sp>
    </p:spTree>
  </p:cSld>
  <p:txStyles><p:bodyStyle/><p:otherStyle/></p:txStyles>
</p:notesMaster>"#
}

/// Animation effect presets for slide shapes.
///
/// Maps to OOXML animation preset IDs and classes.
/// ECMA-376 Part 1, Section 19.5 (timing / animations).
#[cfg(feature = "pml-animations")]
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum AnimationEffect {
    /// Shape appears instantly (`presetID="1"` class `entr`).
    Appear,
    /// Fade in from transparent (`presetID="10"` class `entr`).
    FadeIn,
    /// Fly in from the left (`presetID="2"` class `entr` subtype `8`).
    FlyInFromLeft,
    /// Fly in from the right (`presetID="2"` class `entr` subtype `4`).
    FlyInFromRight,
    /// Fly in from the bottom (`presetID="2"` class `entr` subtype `16`).
    FlyInFromBottom,
    /// Fly in from the top (`presetID="2"` class `entr` subtype `2`).
    FlyInFromTop,
    /// Zoom in from small (`presetID="22"` class `entr`).
    ZoomIn,
    /// Spin clockwise (`presetID="3"` class `emph`).
    SpinClockwise,
    /// Blink/flash emphasis (`presetID="14"` class `emph`).
    Blink,
    /// Fade out (`presetID="10"` class `exit`).
    FadeOut,
    /// Fly out (`presetID="2"` class `exit`).
    FlyOut,
}

/// Trigger for when an animation begins.
///
/// ECMA-376 Part 1, §19.5.30 (cTn nodeType).
#[cfg(feature = "pml-animations")]
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum AnimationTrigger {
    /// Start on the next mouse click.
    OnClick,
    /// Start simultaneously with the previous animation.
    WithPrevious,
    /// Start after the previous animation finishes.
    AfterPrevious,
}

/// A pending animation to apply to a shape on this slide.
///
/// Stored on [`SlideBuilder`] and serialised into `<p:timing>` at write time.
/// ECMA-376 Part 1, §19.5 (timing).
#[cfg(feature = "pml-animations")]
#[derive(Debug, Clone)]
pub struct AnimationConfig {
    /// Shape to animate (1-based `id` attribute on the shape's `<p:cNvPr>`).
    pub shape_id: u32,
    /// Effect to apply.
    pub effect: AnimationEffect,
    /// What triggers the animation.
    pub trigger: AnimationTrigger,
    /// Duration of the effect in milliseconds.
    pub duration_ms: u32,
    /// Delay before the effect starts, in milliseconds.
    pub delay_ms: u32,
}

/// Internal storage for a chart to embed on a slide.
#[cfg(feature = "pml-charts")]
#[derive(Debug, Clone)]
struct ChartElement {
    data: Vec<u8>,
    x: i64,
    y: i64,
    cx: i64,
    cy: i64,
}

/// Internal storage for a SmartArt diagram to embed on a slide.
#[cfg(feature = "pml-charts")]
#[derive(Debug, Clone)]
struct SmartArtElement {
    data_xml: Vec<u8>,
    layout_xml: Option<Vec<u8>>,
    x: i64,
    y: i64,
    cx: i64,
    cy: i64,
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
    make_preset_geom(dml::STShapeType::Rect)
}

/// Construct a preset geometry from a given `STShapeType`.
fn make_preset_geom(preset: dml::STShapeType) -> Box<dml::EGGeometry> {
    Box::new(dml::EGGeometry::PrstGeom(Box::new(
        dml::CTPresetGeometry2D {
            preset,
            av_lst: None,
            extra_attrs: Default::default(),
            extra_children: Default::default(),
        },
    )))
}

/// Preset shape geometry for use with [`ShapeBuilder::set_geometry`].
///
/// Each variant maps to an OOXML preset geometry name (`prst` attribute on `<a:prstGeom>`).
/// ECMA-376 Part 1, Section 20.1.9.18 (prstGeom).
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum PresetGeometry {
    /// Rectangle (`rect`).
    Rect,
    /// Rounded rectangle (`roundRect`).
    RoundRect,
    /// Ellipse (`ellipse`).
    Ellipse,
    /// Triangle (`triangle`).
    Triangle,
    /// Right triangle (`rtTriangle`).
    RightTriangle,
    /// Diamond (`diamond`).
    Diamond,
    /// Pentagon (`pentagon`).
    Pentagon,
    /// Hexagon (`hexagon`).
    Hexagon,
    /// Heptagon (`heptagon`).
    Heptagon,
    /// Octagon (`octagon`).
    Octagon,
    /// 4-pointed star (`star4`).
    Star4,
    /// 5-pointed star (`star5`).
    Star5,
    /// 6-pointed star (`star6`).
    Star6,
    /// 7-pointed star (`star7`).
    Star7,
    /// 8-pointed star (`star8`).
    Star8,
    /// Right arrow (`rightArrow`).
    Arrow,
    /// Left arrow (`leftArrow`).
    LeftArrow,
    /// Up arrow (`upArrow`).
    UpArrow,
    /// Down arrow (`downArrow`).
    DownArrow,
    /// Rectangle callout (`callout1`).
    CalloutRect,
    /// Wedge rect callout (`wedgeRectCallout`).
    WedgeRectCallout,
    /// Cloud (`cloud`).
    Cloud,
    /// Heart (`heart`).
    Heart,
    /// Lightning bolt (`lightningBolt`).
    Lightning,
    /// Straight line — use as a connector geometry (`line`).
    Line,
    /// Arbitrary preset name for shapes not in this enum.
    Custom(String),
}

impl PresetGeometry {
    fn to_shape_type(&self) -> dml::STShapeType {
        match self {
            PresetGeometry::Rect => dml::STShapeType::Rect,
            PresetGeometry::RoundRect => dml::STShapeType::RoundRect,
            PresetGeometry::Ellipse => dml::STShapeType::Ellipse,
            PresetGeometry::Triangle => dml::STShapeType::Triangle,
            PresetGeometry::RightTriangle => dml::STShapeType::RtTriangle,
            PresetGeometry::Diamond => dml::STShapeType::Diamond,
            PresetGeometry::Pentagon => dml::STShapeType::Pentagon,
            PresetGeometry::Hexagon => dml::STShapeType::Hexagon,
            PresetGeometry::Heptagon => dml::STShapeType::Heptagon,
            PresetGeometry::Octagon => dml::STShapeType::Octagon,
            PresetGeometry::Star4 => dml::STShapeType::Star4,
            PresetGeometry::Star5 => dml::STShapeType::Star5,
            PresetGeometry::Star6 => dml::STShapeType::Star6,
            PresetGeometry::Star7 => dml::STShapeType::Star7,
            PresetGeometry::Star8 => dml::STShapeType::Star8,
            PresetGeometry::Arrow => dml::STShapeType::RightArrow,
            PresetGeometry::LeftArrow => dml::STShapeType::LeftArrow,
            PresetGeometry::UpArrow => dml::STShapeType::UpArrow,
            PresetGeometry::DownArrow => dml::STShapeType::DownArrow,
            PresetGeometry::CalloutRect => dml::STShapeType::Callout1,
            PresetGeometry::WedgeRectCallout => dml::STShapeType::WedgeRectCallout,
            PresetGeometry::Cloud => dml::STShapeType::Cloud,
            PresetGeometry::Heart => dml::STShapeType::Heart,
            PresetGeometry::Lightning => dml::STShapeType::LightningBolt,
            PresetGeometry::Line => dml::STShapeType::Line,
            PresetGeometry::Custom(s) => s.parse().unwrap_or(dml::STShapeType::Rect),
        }
    }
}

/// Paragraph alignment for text inside a shape.
///
/// Maps to the `algn` attribute on `<a:pPr>`.
/// ECMA-376 Part 1, Section 21.1.2.2.7 (pPr).
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum TextAlign {
    /// Left alignment (`l`).
    Left,
    /// Center alignment (`ctr`).
    Center,
    /// Right alignment (`r`).
    Right,
    /// Justify (`just`).
    Justify,
}

impl TextAlign {
    fn to_dml(self) -> dml::STTextAlignType {
        match self {
            TextAlign::Left => dml::STTextAlignType::L,
            TextAlign::Center => dml::STTextAlignType::Ctr,
            TextAlign::Right => dml::STTextAlignType::R,
            TextAlign::Justify => dml::STTextAlignType::Just,
        }
    }
}

/// A paragraph inside a shape, containing text runs and optional alignment.
///
/// Use [`Paragraph::new`] or [`Paragraph::with_runs`] to build paragraphs.
/// Add to a [`ShapeBuilder`] via [`ShapeBuilder::add_paragraph`].
#[derive(Debug, Clone)]
pub struct Paragraph {
    /// The text runs in this paragraph.
    pub runs: Vec<TextRun>,
    /// Paragraph-level alignment (maps to `<a:pPr algn="...">`).
    pub align: Option<TextAlign>,
}

impl Paragraph {
    /// Create a new paragraph with a single plain-text run.
    pub fn new(text: impl Into<String>) -> Self {
        Self {
            runs: vec![TextRun::text(text)],
            align: None,
        }
    }

    /// Create a paragraph with multiple runs.
    pub fn with_runs(runs: Vec<TextRun>) -> Self {
        Self { runs, align: None }
    }

    /// Set the paragraph alignment.
    pub fn set_align(mut self, align: TextAlign) -> Self {
        self.align = Some(align);
        self
    }
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
    /// Additional paragraphs beyond the first (built from `runs`).
    extra_paragraphs: Vec<Paragraph>,
    is_title: bool,
    x: i64,
    y: i64,
    width: i64,
    height: i64,
    fill_color: Option<String>,
    line_color: Option<String>,
    line_width: Option<i64>,
    /// Preset shape geometry (defaults to `Rect` if not set).
    geometry: Option<PresetGeometry>,
    /// Alignment for the initial paragraph (the one built from `runs`).
    align: Option<TextAlign>,
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

    /// Set the preset shape geometry.
    ///
    /// By default shapes use `Rect`. Use this to produce other shapes such as
    /// `Ellipse`, `RoundRect`, `Star5`, etc.
    ///
    /// Sets `spPr/prstGeom@prst`.
    /// ECMA-376 Part 1, Section 20.1.9.18 (prstGeom).
    pub fn set_geometry(mut self, geom: PresetGeometry) -> Self {
        self.geometry = Some(geom);
        self
    }

    /// Set paragraph-level text alignment for the initial paragraph.
    ///
    /// Sets `<a:pPr algn="..."/>` on the first paragraph of the text body.
    /// ECMA-376 Part 1, Section 21.1.2.2.7 (pPr).
    pub fn set_text_align(mut self, align: TextAlign) -> Self {
        self.align = Some(align);
        self
    }

    /// Add an additional paragraph to the shape's text body.
    ///
    /// This enables multiple paragraphs in a single shape, which can have
    /// different text runs, formatting, and alignment.
    pub fn add_paragraph(mut self, para: Paragraph) -> Self {
        self.extra_paragraphs.push(para);
        self
    }

    /// Finalize the shape and add it to the slide.
    pub fn add(self) {
        let ShapeBuilder {
            slide,
            runs,
            extra_paragraphs,
            is_title,
            x,
            y,
            width,
            height,
            fill_color,
            line_color,
            line_width,
            geometry,
            align,
        } = self;

        // Build the first paragraph from the initial `runs`.
        let first_para = Paragraph { runs, align };
        let mut paragraphs = vec![first_para];
        paragraphs.extend(extra_paragraphs);

        let element = TextElement {
            paragraphs,
            is_title,
            x,
            y,
            width,
            height,
            fill_color,
            line_color,
            line_width,
            geometry,
        };
        slide.push_text_element(element);
    }
}

/// A text element (used for shapes with hyperlinks, deferred until write time).
#[derive(Debug, Clone)]
struct TextElement {
    /// Paragraphs in the text body (at least one).
    paragraphs: Vec<Paragraph>,
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
    /// Optional preset shape geometry. Defaults to `Rect` if `None`.
    geometry: Option<PresetGeometry>,
}

impl TextElement {
    fn simple(text: String, is_title: bool, x: i64, y: i64, width: i64, height: i64) -> Self {
        Self {
            paragraphs: vec![Paragraph {
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
                align: None,
            }],
            is_title,
            x,
            y,
            width,
            height,
            fill_color: None,
            line_color: None,
            line_width: None,
            geometry: None,
        }
    }

    fn has_hyperlink(&self) -> bool {
        self.paragraphs
            .iter()
            .any(|p| p.runs.iter().any(|r| r.hyperlink.is_some()))
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
    /// Which layout this slide uses (index into the master's layout list).
    layout_id: Option<LayoutId>,
    /// Pending animations — serialised into `<p:timing>` at write time.
    #[cfg(feature = "pml-animations")]
    animations: Vec<AnimationConfig>,
    /// Charts to embed — serialised into chart parts at write time.
    #[cfg(feature = "pml-charts")]
    charts: Vec<ChartElement>,
    /// SmartArt diagrams to embed — serialised at write time.
    #[cfg(feature = "pml-charts")]
    smartarts: Vec<SmartArtElement>,
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

/// Build DML text runs from a slice of `TextRun`s, resolving hyperlink rel IDs.
fn build_dml_runs(
    runs: &[TextRun],
    default_font_size: i32,
    hyperlink_rel_ids: Option<&std::collections::HashMap<&str, usize>>,
) -> Vec<dml::EGTextRun> {
    runs.iter()
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
        .collect()
}

/// Build a `types::Shape` from a `TextElement`, resolving hyperlink rel IDs if provided.
fn build_shape_impl(
    element: &TextElement,
    shape_id: usize,
    hyperlink_rel_ids: Option<&std::collections::HashMap<&str, usize>>,
) -> types::Shape {
    let name = if element.is_title { "Title" } else { "Content" };
    let default_font_size: i32 = if element.is_title { 4400 } else { 2400 };

    // Build all paragraphs.
    let dml_paragraphs: Vec<dml::TextParagraph> = element
        .paragraphs
        .iter()
        .map(|para| {
            let runs = build_dml_runs(&para.runs, default_font_size, hyperlink_rel_ids);
            // Build paragraph properties if alignment is set.
            // dml-text is always on in DML's default feature set, so p_pr is always available.
            let p_pr = para.align.map(|align| {
                Box::new(dml::TextParagraphProperties {
                    algn: Some(align.to_dml()),
                    ..Default::default()
                })
            });
            dml::TextParagraph {
                text_run: runs,
                p_pr,
                ..Default::default()
            }
        })
        .collect();

    let text_body = dml::TextBody {
        body_pr: Box::default(),
        lst_style: None,
        p: dml_paragraphs,
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

    // Shape geometry: use explicit override or default to Rect.
    let geom = match &element.geometry {
        Some(g) => make_preset_geom(g.to_shape_type()),
        None => make_rect_geom(),
    };

    let sp_pr = dml::CTShapeProperties {
        transform: Some(make_xfrm(
            element.x,
            element.y,
            element.width,
            element.height,
        )),
        geometry: Some(geom),
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

/// A connector element stored on the slide.
#[derive(Debug, Clone)]
struct ConnectorElement {
    x1: i64,
    y1: i64,
    x2: i64,
    y2: i64,
    line_color: Option<String>,
}

/// Build a `types::Connector` from a `ConnectorElement`.
///
/// A connector spans from (x1, y1) to (x2, y2). The bounding box origin is the
/// top-left corner (min x, min y) and the extents are (|dx|, |dy|).  When the
/// line runs right-to-left or bottom-to-top we set `flipH`/`flipV` so that the
/// geometry is rendered in the correct direction.
///
/// ECMA-376 Part 1, Section 19.3.1.19 (cxnSp).
fn build_connector_impl(elem: &ConnectorElement, shape_id: usize) -> types::Connector {
    let dx = elem.x2 - elem.x1;
    let dy = elem.y2 - elem.y1;

    let bx = elem.x1.min(elem.x2);
    let by = elem.y1.min(elem.y2);
    let bw = dx.unsigned_abs() as i64;
    let bh = dy.unsigned_abs() as i64;

    let mut xfrm = *make_xfrm(bx, by, bw, bh);
    // flipH if the line goes right-to-left, flipV if it goes bottom-to-top.
    if dx < 0 {
        xfrm.flip_h = Some(true);
    }
    if dy < 0 {
        xfrm.flip_v = Some(true);
    }

    let line_fill = elem.line_color.as_deref().map(|rgb| {
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

    let line = if line_fill.is_some() {
        Some(Box::new(dml::LineProperties {
            line_fill_properties: line_fill,
            ..Default::default()
        }))
    } else {
        None
    };

    let sp_pr = dml::CTShapeProperties {
        transform: Some(Box::new(xfrm)),
        geometry: Some(make_preset_geom(dml::STShapeType::Line)),
        line,
        ..Default::default()
    };

    let name = format!("Connector {}", shape_id);
    types::Connector {
        non_visual_connector_properties: Box::new(types::CTConnectorNonVisual {
            c_nv_pr: make_cnv_pr(shape_id as u32, &name),
            c_nv_cxn_sp_pr: Box::new(dml::CTNonVisualConnectorProperties {
                ..Default::default()
            }),
            nv_pr: make_nv_pr(),
            extra_children: Default::default(),
        }),
        shape_properties: Box::new(sp_pr),
        #[cfg(feature = "pml-styling")]
        style: None,
        #[cfg(feature = "pml-extensions")]
        ext_lst: None,
        extra_children: Default::default(),
    }
}

/// A pending group of shapes to be wrapped in `<p:grpSp>`.
#[derive(Debug)]
pub struct GroupBuilder<'a> {
    slide: &'a mut SlideBuilder,
    shapes: Vec<types::Shape>,
    connectors: Vec<types::Connector>,
    name: Option<String>,
}

impl<'a> GroupBuilder<'a> {
    /// Set the group name (stored in `nvGrpSpPr/cNvPr@name`).
    pub fn set_name(mut self, name: impl Into<String>) -> Self {
        self.name = Some(name.into());
        self
    }

    /// Add a plain-text shape to the group.
    ///
    /// Position and size are in EMUs.
    pub fn add_text(
        mut self,
        text: impl Into<String>,
        x: i64,
        y: i64,
        width: i64,
        height: i64,
    ) -> Self {
        let element = TextElement::simple(text.into(), false, x, y, width, height);
        let shape_id = self.slide.next_shape_id;
        self.slide.next_shape_id += 1;
        self.shapes.push(build_shape_impl(&element, shape_id, None));
        self
    }

    /// Add a connector to the group.
    pub fn add_connector(mut self, x1: i64, y1: i64, x2: i64, y2: i64) -> Self {
        let elem = ConnectorElement {
            x1,
            y1,
            x2,
            y2,
            line_color: None,
        };
        let shape_id = self.slide.next_shape_id;
        self.slide.next_shape_id += 1;
        self.connectors.push(build_connector_impl(&elem, shape_id));
        self
    }

    /// Finalize the group and add it to the slide.
    ///
    /// Computes the bounding box from all child shapes and writes a `<p:grpSp>`
    /// element with `<p:grpSpPr>` defining the group transform.
    /// ECMA-376 Part 1, Section 19.3.1.22 (grpSp).
    pub fn finish(self) -> &'a mut SlideBuilder {
        let GroupBuilder {
            slide,
            shapes,
            connectors,
            name,
        } = self;

        if shapes.is_empty() && connectors.is_empty() {
            return slide;
        }

        let group_id = slide.next_shape_id;
        slide.next_shape_id += 1;

        let group_name = name.unwrap_or_else(|| format!("Group {}", group_id));

        // Compute bounding box from child shape transforms.
        let mut min_x = i64::MAX;
        let mut min_y = i64::MAX;
        let mut max_x = i64::MIN;
        let mut max_y = i64::MIN;

        for shape in &shapes {
            if let Some(xfrm) = &shape.shape_properties.transform
                && let Some(off) = &xfrm.offset
            {
                let x: i64 = off.x.parse().unwrap_or(0);
                let y: i64 = off.y.parse().unwrap_or(0);
                min_x = min_x.min(x);
                min_y = min_y.min(y);
                if let Some(ext) = &xfrm.extents {
                    max_x = max_x.max(x + ext.cx);
                    max_y = max_y.max(y + ext.cy);
                }
            }
        }

        for conn in &connectors {
            if let Some(xfrm) = &conn.shape_properties.transform
                && let Some(off) = &xfrm.offset
            {
                let x: i64 = off.x.parse().unwrap_or(0);
                let y: i64 = off.y.parse().unwrap_or(0);
                min_x = min_x.min(x);
                min_y = min_y.min(y);
                if let Some(ext) = &xfrm.extents {
                    max_x = max_x.max(x + ext.cx);
                    max_y = max_y.max(y + ext.cy);
                }
            }
        }

        // Fallback if no transforms were found.
        if min_x == i64::MAX {
            min_x = 0;
            min_y = 0;
            max_x = 0;
            max_y = 0;
        }

        let bw = (max_x - min_x).max(0);
        let bh = (max_y - min_y).max(0);

        let grp_xfrm = dml::CTGroupTransform2D {
            offset: Some(Box::new(dml::Point2D {
                x: min_x.to_string(),
                y: min_y.to_string(),
                extra_attrs: Default::default(),
            })),
            extents: Some(Box::new(dml::PositiveSize2D {
                cx: bw,
                cy: bh,
                extra_attrs: Default::default(),
            })),
            child_offset: Some(Box::new(dml::Point2D {
                x: min_x.to_string(),
                y: min_y.to_string(),
                extra_attrs: Default::default(),
            })),
            child_extents: Some(Box::new(dml::PositiveSize2D {
                cx: bw,
                cy: bh,
                extra_attrs: Default::default(),
            })),
            ..Default::default()
        };

        let grp_sp_pr = dml::CTGroupShapeProperties {
            transform: Some(Box::new(grp_xfrm)),
            ..Default::default()
        };

        let group = types::GroupShape {
            non_visual_group_properties: Box::new(types::CTGroupShapeNonVisual {
                c_nv_pr: make_cnv_pr(group_id as u32, &group_name),
                c_nv_grp_sp_pr: Box::new(dml::CTNonVisualGroupDrawingShapeProps {
                    grp_sp_locks: None,
                    ext_lst: None,
                    extra_children: Default::default(),
                }),
                nv_pr: make_nv_pr(),
                extra_children: Default::default(),
            }),
            grp_sp_pr: Box::new(grp_sp_pr),
            shape: shapes,
            group_shape: Vec::new(),
            graphic_frame: Vec::new(),
            connector: connectors,
            picture: Vec::new(),
            content_part: Vec::new(),
            ext_lst: None,
            extra_children: Default::default(),
        };

        slide
            .slide
            .common_slide_data
            .shape_tree
            .group_shape
            .push(group);
        slide
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

/// Build a `types::GraphicalObjectFrame` for an embedded chart.
///
/// The chart data is referenced via a relationship ID `rel_id`.
/// ECMA-376 Part 1, §14.2.1.
#[cfg(feature = "pml-charts")]
fn build_chart_frame(
    chart: &ChartElement,
    shape_id: usize,
    rel_id: usize,
) -> types::GraphicalObjectFrame {
    let chart_name = format!("Chart {}", shape_id);

    let graphic_frame_locks = dml::CTGraphicalObjectFrameLocking {
        no_grp: Some(true),
        ..Default::default()
    };
    let c_nv_graphic_frame_pr = dml::CTNonVisualGraphicFrameProperties {
        graphic_frame_locks: Some(Box::new(graphic_frame_locks)),
        ..Default::default()
    };
    let nv_graphic_frame_pr = Box::new(types::CTGraphicalObjectFrameNonVisual {
        c_nv_pr: make_cnv_pr(shape_id as u32, &chart_name),
        c_nv_graphic_frame_pr: Box::new(c_nv_graphic_frame_pr),
        nv_pr: make_nv_pr(),
        extra_children: Default::default(),
    });

    let xfrm = make_xfrm(chart.x, chart.y, chart.cx, chart.cy);

    let graphic_xml = format!(
        r#"<a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="rId{}"/></a:graphicData></a:graphic>"#,
        rel_id
    );
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

/// Build a `types::GraphicalObjectFrame` for an embedded SmartArt diagram.
///
/// The diagram data part is referenced via `data_rel_id`, and optionally a
/// layout part via `layout_rel_id`.
/// ECMA-376 Part 1, §14.2.4.
#[cfg(feature = "pml-charts")]
fn build_smartart_frame(
    smartart: &SmartArtElement,
    shape_id: usize,
    data_rel_id: usize,
    layout_rel_id: Option<usize>,
) -> types::GraphicalObjectFrame {
    let frame_name = format!("SmartArt {}", shape_id);

    let graphic_frame_locks = dml::CTGraphicalObjectFrameLocking {
        no_grp: Some(true),
        ..Default::default()
    };
    let c_nv_graphic_frame_pr = dml::CTNonVisualGraphicFrameProperties {
        graphic_frame_locks: Some(Box::new(graphic_frame_locks)),
        ..Default::default()
    };
    let nv_graphic_frame_pr = Box::new(types::CTGraphicalObjectFrameNonVisual {
        c_nv_pr: make_cnv_pr(shape_id as u32, &frame_name),
        c_nv_graphic_frame_pr: Box::new(c_nv_graphic_frame_pr),
        nv_pr: make_nv_pr(),
        extra_children: Default::default(),
    });

    let xfrm = make_xfrm(smartart.x, smartart.y, smartart.cx, smartart.cy);

    // Build the relIds attributes — dm= is the data rel, lo= is the layout rel.
    let layout_attr = if let Some(lo_id) = layout_rel_id {
        format!(r#" r:lo="rId{}""#, lo_id)
    } else {
        String::new()
    };
    let graphic_xml = format!(
        r#"<a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/diagram"><dgm:relIds xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:dm="rId{}"{}/></a:graphicData></a:graphic>"#,
        data_rel_id, layout_attr
    );
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

/// Build the raw XML string for `<p:timing>` from a list of animation configs.
///
/// Uses the minimal timing structure required by PowerPoint.
/// ECMA-376 Part 1, §19.5.
#[cfg(feature = "pml-animations")]
fn build_timing_xml(animations: &[AnimationConfig]) -> String {
    // cTn id assignment:
    //   1 = root tmRoot
    //   2 = mainSeq
    //   per animation (4 ids each): outer_par_ctn, effect_ctn, set_ctn, (reserved)
    let mut xml = String::new();
    xml.push_str("<p:timing>");
    xml.push_str("<p:tnLst><p:par>");
    xml.push_str(r#"<p:cTn id="1" dur="indefinite" restart="whenNotActive" nodeType="tmRoot">"#);
    xml.push_str("<p:childTnLst>");
    xml.push_str(r#"<p:seq concurrent="1" nextAc="seek">"#);
    xml.push_str(r#"<p:cTn id="2" dur="indefinite" nodeType="mainSeq"><p:childTnLst>"#);

    let mut ctn_id: u32 = 3;
    for (grp_idx, anim) in animations.iter().enumerate() {
        let (preset_id, preset_class, preset_subtype) = animation_preset(anim);
        let outer_delay = match anim.trigger {
            AnimationTrigger::OnClick => "indefinite",
            AnimationTrigger::WithPrevious | AnimationTrigger::AfterPrevious => "0",
        };
        let node_type = match anim.trigger {
            AnimationTrigger::OnClick => "clickEffect",
            AnimationTrigger::WithPrevious => "withEffect",
            AnimationTrigger::AfterPrevious => "afterEffect",
        };

        let outer_id = ctn_id;
        let effect_id = ctn_id + 1;
        let set_id = ctn_id + 2;
        ctn_id += 3;

        // Outer par — groups this click-level animation.
        xml.push_str("<p:par>");
        xml.push_str(&format!(
            r#"<p:cTn id="{}" fill="hold"><p:stCondLst><p:cond evt="onBegin" delay="{}"/></p:stCondLst><p:childTnLst><p:par>"#,
            outer_id, outer_delay
        ));
        xml.push_str(&format!(
            r#"<p:cTn id="{}" presetID="{}" presetClass="{}" presetSubtype="{}" fill="hold" grpId="{}" nodeType="{}" dur="{}">"#,
            effect_id,
            preset_id,
            preset_class,
            preset_subtype,
            grp_idx,
            node_type,
            anim.duration_ms
        ));
        xml.push_str(&format!(
            r#"<p:stCondLst><p:cond evt="begin" delay="{}"/></p:stCondLst>"#,
            anim.delay_ms
        ));
        xml.push_str("<p:childTnLst><p:set><p:cBhvr>");
        xml.push_str(&format!(r#"<p:cTn id="{}" dur="1" fill="hold"/>"#, set_id));
        xml.push_str(&format!(
            r#"<p:tgtEl><p:spTgt spid="{}"/></p:tgtEl>"#,
            anim.shape_id
        ));
        xml.push_str("<p:attrNameLst><p:attrName>style.visibility</p:attrName></p:attrNameLst>");
        xml.push_str(r#"</p:cBhvr><p:to><p:strVal val="visible"/></p:to></p:set>"#);
        // Close: childTnLst, effect cTn, inner par, childTnLst, outer cTn, outer par
        xml.push_str("</p:childTnLst></p:cTn></p:par></p:childTnLst></p:cTn></p:par>");
    }

    xml.push_str("</p:childTnLst></p:cTn>"); // mainSeq cTn
    xml.push_str(r#"<p:prevCondLst><p:cond evt="onBegin" delay="0"><p:tn><p:tgtEl><p:sldTgt/></p:tgtEl></p:tn></p:cond></p:prevCondLst>"#);
    xml.push_str("</p:seq>");
    xml.push_str("</p:childTnLst></p:cTn>"); // root cTn
    xml.push_str("</p:par></p:tnLst>");

    xml.push_str("<p:bldLst>");
    for anim in animations {
        xml.push_str(&format!(
            r#"<p:bldP spid="{}" grpId="0" uiExpand="1" build="p"/>"#,
            anim.shape_id
        ));
    }
    xml.push_str("</p:bldLst>");
    xml.push_str("</p:timing>");
    xml
}

/// Map an `AnimationConfig` to `(presetID, presetClass, presetSubtype)`.
#[cfg(feature = "pml-animations")]
fn animation_preset(anim: &AnimationConfig) -> (u32, &'static str, u32) {
    match anim.effect {
        AnimationEffect::Appear => (1, "entr", 0),
        AnimationEffect::FadeIn => (10, "entr", 0),
        AnimationEffect::FlyInFromLeft => (2, "entr", 8),
        AnimationEffect::FlyInFromRight => (2, "entr", 4),
        AnimationEffect::FlyInFromBottom => (2, "entr", 16),
        AnimationEffect::FlyInFromTop => (2, "entr", 2),
        AnimationEffect::ZoomIn => (22, "entr", 0),
        AnimationEffect::SpinClockwise => (3, "emph", 0),
        AnimationEffect::Blink => (14, "emph", 0),
        AnimationEffect::FadeOut => (10, "exit", 0),
        AnimationEffect::FlyOut => (2, "exit", 0),
    }
}

/// Parse the timing XML string into a `types::SlideTiming` element for the
/// generated slide type.
///
/// We store the whole `<p:timing>` as `extra_children` on the slide so it
/// roundtrips faithfully without needing to map every generated sub-type.
#[cfg(feature = "pml-animations")]
fn build_timing_element(timing_xml: &str) -> types::SlideTiming {
    // Strip the outer <p:timing> wrapper — its children go into extra_children
    // of the generated SlideTiming type.
    let inner = timing_xml
        .trim_start_matches("<p:timing>")
        .trim_end_matches("</p:timing>")
        .trim();
    let extra_children = parse_extra_child_xml(inner);
    types::SlideTiming {
        extra_children,
        ..Default::default()
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
            layout_id: None,
            #[cfg(feature = "pml-animations")]
            animations: Vec::new(),
            #[cfg(feature = "pml-charts")]
            charts: Vec::new(),
            #[cfg(feature = "pml-charts")]
            smartarts: Vec::new(),
        }
    }

    /// Collect all hyperlinks from this slide's deferred text elements.
    fn hyperlinks(&self) -> Vec<&str> {
        let mut links = Vec::new();
        for element in &self.hyperlink_elements {
            for para in &element.paragraphs {
                for run in &para.runs {
                    if let Some(ref url) = run.hyperlink
                        && !links.contains(&url.as_str())
                    {
                        links.push(url.as_str());
                    }
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
            paragraphs: vec![Paragraph {
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
                align: None,
            }],
            is_title: false,
            x,
            y,
            width,
            height,
            fill_color: None,
            line_color: None,
            line_width: None,
            geometry: None,
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
            paragraphs: vec![Paragraph { runs, align: None }],
            is_title: false,
            x,
            y,
            width,
            height,
            fill_color: None,
            line_color: None,
            line_width: None,
            geometry: None,
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
            extra_paragraphs: Vec::new(),
            is_title: false,
            x,
            y,
            width,
            height,
            fill_color: None,
            line_color: None,
            line_width: None,
            geometry: None,
            align: None,
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

    /// Add a connector (straight line) between two points.
    ///
    /// Position coordinates are in EMUs (914400 EMUs = 1 inch).
    /// `color` is an optional 6-character hex RGB string for the line color.
    ///
    /// Produces a `<p:cxnSp>` element with `<a:prstGeom prst="line"/>`.
    /// ECMA-376 Part 1, Section 19.3.1.19 (cxnSp).
    pub fn add_connector(
        &mut self,
        x1: i64,
        y1: i64,
        x2: i64,
        y2: i64,
        color: Option<&str>,
    ) -> &mut Self {
        let elem = ConnectorElement {
            x1,
            y1,
            x2,
            y2,
            line_color: color.map(|s| s.to_string()),
        };
        let shape_id = self.next_shape_id;
        self.next_shape_id += 1;
        let connector = build_connector_impl(&elem, shape_id);
        self.slide
            .common_slide_data
            .shape_tree
            .connector
            .push(connector);
        self
    }

    /// Begin building a group of shapes (`<p:grpSp>`).
    ///
    /// Use [`GroupBuilder`] to add shapes and connectors to the group, then
    /// call [`GroupBuilder::finish`] to finalize it.
    ///
    /// ECMA-376 Part 1, Section 19.3.1.22 (grpSp).
    pub fn begin_group(&mut self) -> GroupBuilder<'_> {
        GroupBuilder {
            slide: self,
            shapes: Vec::new(),
            connectors: Vec::new(),
            name: None,
        }
    }

    /// Check if this slide has connectors.
    pub fn has_connectors(&self) -> bool {
        !self.slide.common_slide_data.shape_tree.connector.is_empty()
    }

    /// Check if this slide has group shapes.
    pub fn has_groups(&self) -> bool {
        !self
            .slide
            .common_slide_data
            .shape_tree
            .group_shape
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

    /// Set which layout this slide uses.
    ///
    /// Pass the [`LayoutId`] returned by
    /// [`PresentationBuilder::add_slide_master`].  If not called, the slide
    /// uses the first layout of the first master (the default).
    pub fn set_layout(&mut self, id: LayoutId) -> &mut Self {
        self.layout_id = Some(id);
        self
    }

    /// Add an animation to a shape on this slide.
    ///
    /// The `shape_id` must match the shape's `id` attribute as written to the
    /// slide XML.  Shape IDs start at 2 (1 is the group shape) and increment
    /// in the order shapes are added.
    ///
    /// Requires the `pml-animations` feature.
    /// ECMA-376 Part 1, §19.5.
    #[cfg(feature = "pml-animations")]
    pub fn add_animation(&mut self, anim: AnimationConfig) -> &mut Self {
        self.animations.push(anim);
        self
    }

    /// Embed a chart in this slide.
    ///
    /// `chart_xml` must be valid DrawingML chart XML (a `<c:chartSpace>` root
    /// element).  At write time the bytes are written to
    /// `ppt/charts/chartN.xml` and a `<p:graphicFrame>` referencing it is
    /// added to the slide's shape tree.
    ///
    /// Requires the `pml-charts` feature.
    /// ECMA-376 Part 1, §14.2.1.
    #[cfg(feature = "pml-charts")]
    pub fn embed_chart(
        &mut self,
        chart_xml: impl Into<Vec<u8>>,
        x: i64,
        y: i64,
        cx: i64,
        cy: i64,
    ) -> &mut Self {
        self.charts.push(ChartElement {
            data: chart_xml.into(),
            x,
            y,
            cx,
            cy,
        });
        self
    }

    /// Embed a SmartArt diagram in this slide.
    ///
    /// `data_xml` must be a `<dgm:dataModel>` document.  `layout_xml`, if
    /// supplied, is written as a separate `diagramLayout` part.  At write time
    /// a `<p:graphicFrame>` with a `<dgm:relIds>` element is added to the
    /// slide's shape tree.
    ///
    /// Requires the `pml-charts` feature (which already covers diagrams).
    /// ECMA-376 Part 1, §14.2.4.
    #[cfg(feature = "pml-charts")]
    pub fn embed_smartart(
        &mut self,
        data_xml: impl Into<Vec<u8>>,
        layout_xml: Option<Vec<u8>>,
        x: i64,
        y: i64,
        cx: i64,
        cy: i64,
    ) -> &mut Self {
        self.smartarts.push(SmartArtElement {
            data_xml: data_xml.into(),
            layout_xml,
            x,
            y,
            cx,
            cy,
        });
        self
    }

    /// Check if this slide has animations.
    #[cfg(feature = "pml-animations")]
    pub fn has_animations(&self) -> bool {
        !self.animations.is_empty()
    }

    /// Check if this slide has embedded charts.
    #[cfg(feature = "pml-charts")]
    pub fn has_charts(&self) -> bool {
        !self.charts.is_empty()
    }

    /// Check if this slide has embedded SmartArt.
    #[cfg(feature = "pml-charts")]
    pub fn has_smartarts(&self) -> bool {
        !self.smartarts.is_empty()
    }

    /// Serialize this slide to XML bytes, resolving deferred images, hyperlinks,
    /// charts, and SmartArt.
    ///
    /// `chart_start_rel_id` and `smartart_start_rel_id` are only used when the
    /// `pml-charts` feature is enabled.  Pass `0` when that feature is off.
    fn serialize_slide(
        &self,
        image_start_rel_id: usize,
        hyperlink_rel_ids: &std::collections::HashMap<&str, usize>,
        chart_start_rel_id: usize,
        smartart_start_rel_id: usize,
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

        // Build graphic frames for embedded charts (deferred — need rIds).
        #[cfg(feature = "pml-charts")]
        for (i, chart) in self.charts.iter().enumerate() {
            let rel_id = chart_start_rel_id + i;
            let frame = build_chart_frame(chart, next_id, rel_id);
            slide.common_slide_data.shape_tree.graphic_frame.push(frame);
            next_id += 1;
        }

        // Build graphic frames for SmartArt diagrams (deferred — need rIds).
        #[cfg(feature = "pml-charts")]
        for (i, smartart) in self.smartarts.iter().enumerate() {
            let data_rel_id = smartart_start_rel_id + i * 2;
            let layout_rel_id = if smartart.layout_xml.is_some() {
                Some(data_rel_id + 1)
            } else {
                None
            };
            let frame = build_smartart_frame(smartart, next_id, data_rel_id, layout_rel_id);
            slide.common_slide_data.shape_tree.graphic_frame.push(frame);
            next_id += 1;
        }

        // Inject <p:timing> for slide animations.
        #[cfg(feature = "pml-animations")]
        if !self.animations.is_empty() {
            let timing_xml = build_timing_xml(&self.animations);
            slide.timing = Some(Box::new(build_timing_element(&timing_xml)));
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
    /// Additional slide masters (beyond the default minimal master).
    extra_masters: Vec<SlideMasterEntry>,
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
            extra_masters: Vec::new(),
        }
    }

    /// Add a custom slide master (with its layouts) to the presentation.
    ///
    /// Returns a [`MasterId`] that uniquely identifies this master and a
    /// `Vec<LayoutId>` with one handle per layout (in the same order as
    /// `config.layouts`).  Layouts can be assigned to slides via
    /// [`SlideBuilder::set_layout`].
    ///
    /// The default minimal master (master 1 / layout 1) is always written even
    /// when custom masters are added.  Custom masters start from index 2.
    pub fn add_slide_master(&mut self, config: SlideMasterConfig) -> (MasterId, Vec<LayoutId>) {
        // The default master occupies index 0 internally; extras start at 1.
        let master_idx = 1 + self.extra_masters.len(); // 1-based index (0 = default)
        let layout_ids: Vec<LayoutId> = (0..config.layouts.len())
            .map(|layout_idx| LayoutId {
                master_idx,
                layout_idx,
            })
            .collect();
        self.extra_masters.push(SlideMasterEntry { config });
        (MasterId(master_idx), layout_ids)
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

        let has_notes = self.slides.iter().any(|s| s.has_notes());

        let root_rels = format!(
            r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="{}" Target="ppt/presentation.xml"/>
</Relationships>"#,
            REL_OFFICE_DOCUMENT
        );

        // -----------------------------------------------------------------------
        // presentation.xml.rels
        //
        // IDs:   rId1..rIdN  = slides
        //        rId{N+1}    = default slide master (master 1)
        //        rId{N+2..}  = extra masters (if any)
        //        last entry  = notes master (if any slides have notes)
        // -----------------------------------------------------------------------
        let n_slides = self.slides.len();
        let first_master_rel_id = n_slides + 1;
        let n_masters = 1 + self.extra_masters.len(); // default + extras
        let notes_master_rel_id = first_master_rel_id + n_masters; // after all masters

        let mut pres_rels = String::new();
        pres_rels.push_str(r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>"#);
        pres_rels.push('\n');
        pres_rels.push_str(
            r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">"#,
        );
        pres_rels.push('\n');
        for i in 0..n_slides {
            let rel_id = i + 1;
            pres_rels.push_str(&format!(
                r#"  <Relationship Id="rId{}" Type="{}" Target="slides/slide{}.xml"/>"#,
                rel_id, REL_SLIDE, rel_id
            ));
            pres_rels.push('\n');
        }
        // Default master (slideMaster1.xml)
        pres_rels.push_str(&format!(
            r#"  <Relationship Id="rId{}" Type="{}" Target="slideMasters/slideMaster1.xml"/>"#,
            first_master_rel_id, REL_SLIDE_MASTER
        ));
        pres_rels.push('\n');
        // Extra masters
        for extra_idx in 0..self.extra_masters.len() {
            let master_num = 2 + extra_idx; // slideMaster2.xml, etc.
            let rel_id = first_master_rel_id + 1 + extra_idx;
            pres_rels.push_str(&format!(
                r#"  <Relationship Id="rId{}" Type="{}" Target="slideMasters/slideMaster{}.xml"/>"#,
                rel_id, REL_SLIDE_MASTER, master_num
            ));
            pres_rels.push('\n');
        }
        // Notes master (if any slide has notes)
        if has_notes {
            pres_rels.push_str(&format!(
                r#"  <Relationship Id="rId{}" Type="{}" Target="notesMasters/notesMaster1.xml"/>"#,
                notes_master_rel_id, REL_NOTES_MASTER
            ));
            pres_rels.push('\n');
        }
        pres_rels.push_str("</Relationships>");

        let presentation_xml = serialize_pml_xml(
            &self.build_presentation(first_master_rel_id),
            "p:presentation",
        )?;

        pkg.add_part("_rels/.rels", CT_RELATIONSHIPS, root_rels.as_bytes())?;
        pkg.add_part(
            "ppt/_rels/presentation.xml.rels",
            CT_RELATIONSHIPS,
            pres_rels.as_bytes(),
        )?;
        pkg.add_part("ppt/presentation.xml", CT_PRESENTATION, &presentation_xml)?;

        // -----------------------------------------------------------------------
        // Slide master / layout constants
        // -----------------------------------------------------------------------
        const CT_SLIDE_MASTER: &str =
            "application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml";
        const CT_SLIDE_LAYOUT: &str =
            "application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml";

        // -----------------------------------------------------------------------
        // Default slide master (slideMaster1.xml) + its single blank layout.
        // -----------------------------------------------------------------------
        let default_master_rels = format!(
            r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="{}" Target="../slideLayouts/slideLayout1.xml"/>
</Relationships>"#,
            REL_SLIDE_LAYOUT
        );
        let default_layout_rels = format!(
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
            default_master_rels.as_bytes(),
        )?;
        pkg.add_part(
            "ppt/slideLayouts/slideLayout1.xml",
            CT_SLIDE_LAYOUT,
            build_slide_layout_xml().as_bytes(),
        )?;
        pkg.add_part(
            "ppt/slideLayouts/_rels/slideLayout1.xml.rels",
            CT_RELATIONSHIPS,
            default_layout_rels.as_bytes(),
        )?;

        // -----------------------------------------------------------------------
        // Extra slide masters and their layouts.
        //
        // Layout numbering is global across all masters; we offset by 2 because
        // the default master already owns slideLayout1.xml.
        // -----------------------------------------------------------------------
        let mut global_layout_num = 2usize; // slideLayout2, slideLayout3, …

        for (extra_idx, entry) in self.extra_masters.iter().enumerate() {
            let master_num = 2 + extra_idx; // slideMaster2, …
            let n_layouts = entry.config.layouts.len();
            let master_xml = build_master_xml(entry, n_layouts);

            // Build master rels (points to each layout).
            let mut master_rels_str = String::new();
            master_rels_str.push_str(r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>"#);
            master_rels_str.push('\n');
            master_rels_str.push_str(r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">"#);
            master_rels_str.push('\n');
            for layout_local_idx in 0..n_layouts {
                let layout_num = global_layout_num + layout_local_idx;
                master_rels_str.push_str(&format!(
                    r#"  <Relationship Id="rId{}" Type="{}" Target="../slideLayouts/slideLayout{}.xml"/>"#,
                    layout_local_idx + 1, REL_SLIDE_LAYOUT, layout_num
                ));
                master_rels_str.push('\n');
            }
            master_rels_str.push_str("</Relationships>");

            let master_path = format!("ppt/slideMasters/slideMaster{}.xml", master_num);
            let master_rels_path =
                format!("ppt/slideMasters/_rels/slideMaster{}.xml.rels", master_num);
            pkg.add_part(&master_path, CT_SLIDE_MASTER, master_xml.as_bytes())?;
            pkg.add_part(
                &master_rels_path,
                CT_RELATIONSHIPS,
                master_rels_str.as_bytes(),
            )?;

            // Write each layout for this master.
            for (layout_local_idx, layout_config) in entry.config.layouts.iter().enumerate() {
                let layout_num = global_layout_num + layout_local_idx;
                let layout_xml = build_layout_xml(layout_config);
                let layout_rels_str = format!(
                    r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="{}" Target="../slideMasters/slideMaster{}.xml"/>
</Relationships>"#,
                    REL_SLIDE_MASTER, master_num
                );
                let layout_path = format!("ppt/slideLayouts/slideLayout{}.xml", layout_num);
                let layout_rels_path =
                    format!("ppt/slideLayouts/_rels/slideLayout{}.xml.rels", layout_num);
                pkg.add_part(&layout_path, CT_SLIDE_LAYOUT, layout_xml.as_bytes())?;
                pkg.add_part(
                    &layout_rels_path,
                    CT_RELATIONSHIPS,
                    layout_rels_str.as_bytes(),
                )?;
            }

            global_layout_num += n_layouts;
        }

        // -----------------------------------------------------------------------
        // Notes master (written once if any slide has notes).
        // -----------------------------------------------------------------------
        if has_notes {
            pkg.add_part(
                "ppt/notesMasters/notesMaster1.xml",
                CT_NOTES_MASTER,
                build_notes_master_xml().as_bytes(),
            )?;
            // Notes master has no rels beyond what the notes slides reference.
            let nm_rels = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
</Relationships>"#;
            pkg.add_part(
                "ppt/notesMasters/_rels/notesMaster1.xml.rels",
                CT_RELATIONSHIPS,
                nm_rels.as_bytes(),
            )?;
        }

        // -----------------------------------------------------------------------
        // Slides (with images, notes, hyperlinks, charts, SmartArt).
        // -----------------------------------------------------------------------
        let mut global_image_num = 1usize;
        #[cfg(feature = "pml-charts")]
        let mut global_chart_num = 1usize;
        #[cfg(feature = "pml-charts")]
        let mut global_diagram_num = 1usize;

        for (i, slide) in self.slides.iter().enumerate() {
            let slide_num = i + 1;
            let hyperlinks = slide.hyperlinks();
            let mut hyperlink_rel_ids: std::collections::HashMap<&str, usize> =
                std::collections::HashMap::new();

            // Determine which layout this slide references.
            let layout_path = if let Some(lid) = slide.layout_id {
                // Extra master layouts: compute the global layout number.
                // master_idx=0 is the default master (layout 1).
                // master_idx≥1 is an extra master; we need to count layouts.
                if lid.master_idx == 0 {
                    "slideLayouts/slideLayout1.xml".to_string()
                } else {
                    // Find the global offset of this master's first layout.
                    let mut offset = 2usize; // default master uses layout 1
                    for prev_extra in 0..(lid.master_idx - 1) {
                        offset += self.extra_masters[prev_extra].config.layouts.len();
                    }
                    let layout_num = offset + lid.layout_idx;
                    format!("slideLayouts/slideLayout{}.xml", layout_num)
                }
            } else {
                "slideLayouts/slideLayout1.xml".to_string()
            };

            // -----------------------------------------------------------------------
            // Slide rels
            // -----------------------------------------------------------------------
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
                    r#"  <Relationship Id="rId1" Type="{}" Target="../{}"/>"#,
                    REL_SLIDE_LAYOUT, layout_path
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

                // Chart rels
                #[cfg(feature = "pml-charts")]
                for (chart_idx, _) in slide.charts.iter().enumerate() {
                    let chart_num = global_chart_num + chart_idx;
                    slide_rels.push_str(&format!(
                        r#"  <Relationship Id="rId{}" Type="{}" Target="../charts/chart{}.xml"/>"#,
                        rel_id, REL_CHART, chart_num
                    ));
                    slide_rels.push('\n');
                    rel_id += 1;
                }

                // SmartArt rels
                #[cfg(feature = "pml-charts")]
                for (sa_idx, sa) in slide.smartarts.iter().enumerate() {
                    let diag_num = global_diagram_num + sa_idx * 2; // data part
                    slide_rels.push_str(&format!(
                        r#"  <Relationship Id="rId{}" Type="{}" Target="../diagrams/data{}.xml"/>"#,
                        rel_id, REL_DIAGRAM_DATA, diag_num
                    ));
                    slide_rels.push('\n');
                    rel_id += 1;
                    if sa.layout_xml.is_some() {
                        slide_rels.push_str(&format!(
                            r#"  <Relationship Id="rId{}" Type="{}" Target="../diagrams/layout{}.xml"/>"#,
                            rel_id, REL_DIAGRAM_LAYOUT, diag_num + 1
                        ));
                        slide_rels.push('\n');
                        rel_id += 1;
                    }
                }
                #[cfg(not(feature = "pml-charts"))]
                let _ = rel_id; // silence unused warning

                slide_rels.push_str("</Relationships>");
                let rels_name = format!("ppt/slides/_rels/slide{}.xml.rels", slide_num);
                pkg.add_part(&rels_name, CT_RELATIONSHIPS, slide_rels.as_bytes())?;

                // Write image data.
                for (img_idx, img) in slide.images.iter().enumerate() {
                    let ext = img.format.extension();
                    let img_path = format!("ppt/media/image{}.{}", global_image_num + img_idx, ext);
                    pkg.add_part(&img_path, img.format.content_type(), &img.data)?;
                }

                // Write chart data.
                #[cfg(feature = "pml-charts")]
                for (chart_idx, chart) in slide.charts.iter().enumerate() {
                    let chart_num = global_chart_num + chart_idx;
                    let chart_path = format!("ppt/charts/chart{}.xml", chart_num);
                    pkg.add_part(&chart_path, CT_CHART, &chart.data)?;
                }

                // Write SmartArt diagram data.
                #[cfg(feature = "pml-charts")]
                for (sa_idx, sa) in slide.smartarts.iter().enumerate() {
                    let diag_num = global_diagram_num + sa_idx * 2;
                    let data_path = format!("ppt/diagrams/data{}.xml", diag_num);
                    pkg.add_part(&data_path, CT_DIAGRAM_DATA, &sa.data_xml)?;
                    if let Some(ref layout_bytes) = sa.layout_xml {
                        let layout_path_diag = format!("ppt/diagrams/layout{}.xml", diag_num + 1);
                        pkg.add_part(&layout_path_diag, CT_DIAGRAM_LAYOUT, layout_bytes)?;
                    }
                }
            }

            // rId1=layout, rId2=notes (if any), then images start.
            let image_start_rel_id = if slide.has_notes() { 3 } else { 2 };
            let slide_xml = slide.serialize_slide(
                image_start_rel_id,
                &hyperlink_rel_ids,
                image_start_rel_id + slide.images.len() + hyperlinks.len(),
                image_start_rel_id + slide.images.len() + hyperlinks.len() + {
                    #[cfg(feature = "pml-charts")]
                    {
                        slide.charts.len()
                    }
                    #[cfg(not(feature = "pml-charts"))]
                    {
                        0
                    }
                },
            )?;
            let part_name = format!("ppt/slides/slide{}.xml", slide_num);
            pkg.add_part(&part_name, CT_SLIDE, &slide_xml)?;

            global_image_num += slide.images.len();
            #[cfg(feature = "pml-charts")]
            {
                global_chart_num += slide.charts.len();
                global_diagram_num += slide.smartarts.len() * 2;
            }

            if slide.has_notes() {
                // Write notes slide rels: rId1=notes master, rId2=back-ref to slide.
                let notes_rels = format!(
                    r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="{}" Target="../notesMasters/notesMaster1.xml"/>
  <Relationship Id="rId2" Type="{}" Target="../slides/slide{}.xml"/>
</Relationships>"#,
                    REL_NOTES_MASTER, REL_SLIDE, slide_num
                );
                let notes_rels_name =
                    format!("ppt/notesSlides/_rels/notesSlide{}.xml.rels", slide_num);
                pkg.add_part(&notes_rels_name, CT_RELATIONSHIPS, notes_rels.as_bytes())?;

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

    // -------------------------------------------------------------------------
    // Feature 1: Preset shape geometries
    // -------------------------------------------------------------------------

    #[test]
    fn test_preset_geometry_ellipse() {
        use std::io::Cursor;

        let mut pres = PresentationBuilder::new();
        let slide = pres.add_slide();
        slide
            .shape(vec![TextRun::text("Oval")], 457200, 457200, 1828800, 914400)
            .set_geometry(PresetGeometry::Ellipse)
            .add();

        let mut buffer = Cursor::new(Vec::new());
        pres.write(&mut buffer).unwrap();

        buffer.set_position(0);
        let mut presentation = crate::Presentation::from_reader(buffer).unwrap();
        let read_slide = presentation.slide(0).unwrap();

        let shapes = read_slide.shapes();
        assert!(!shapes.is_empty());
        let sp_pr = &shapes[0].shape_properties;
        let geom = sp_pr.geometry.as_ref().expect("shape has geometry");
        if let ooxml_dml::types::EGGeometry::PrstGeom(prst) = geom.as_ref() {
            assert_eq!(prst.preset, ooxml_dml::types::STShapeType::Ellipse);
        } else {
            panic!("Expected PrstGeom geometry");
        }
    }

    #[test]
    fn test_preset_geometry_roundrect() {
        use std::io::Cursor;

        let mut pres = PresentationBuilder::new();
        let slide = pres.add_slide();
        slide
            .shape(
                vec![TextRun::text("Rounded")],
                457200,
                457200,
                1828800,
                914400,
            )
            .set_geometry(PresetGeometry::RoundRect)
            .add();

        let mut buffer = Cursor::new(Vec::new());
        pres.write(&mut buffer).unwrap();

        buffer.set_position(0);
        let mut presentation = crate::Presentation::from_reader(buffer).unwrap();
        let read_slide = presentation.slide(0).unwrap();

        let shapes = read_slide.shapes();
        let sp_pr = &shapes[0].shape_properties;
        let geom = sp_pr.geometry.as_ref().expect("shape has geometry");
        if let ooxml_dml::types::EGGeometry::PrstGeom(prst) = geom.as_ref() {
            assert_eq!(prst.preset, ooxml_dml::types::STShapeType::RoundRect);
        } else {
            panic!("Expected PrstGeom geometry");
        }
    }

    #[test]
    fn test_preset_geometry_default_is_rect() {
        use std::io::Cursor;

        let mut pres = PresentationBuilder::new();
        let slide = pres.add_slide();
        slide
            .shape(
                vec![TextRun::text("Default")],
                457200,
                457200,
                1828800,
                914400,
            )
            .add();

        let mut buffer = Cursor::new(Vec::new());
        pres.write(&mut buffer).unwrap();

        buffer.set_position(0);
        let mut presentation = crate::Presentation::from_reader(buffer).unwrap();
        let read_slide = presentation.slide(0).unwrap();

        let shapes = read_slide.shapes();
        let sp_pr = &shapes[0].shape_properties;
        let geom = sp_pr.geometry.as_ref().expect("shape has geometry");
        if let ooxml_dml::types::EGGeometry::PrstGeom(prst) = geom.as_ref() {
            assert_eq!(prst.preset, ooxml_dml::types::STShapeType::Rect);
        } else {
            panic!("Expected PrstGeom geometry");
        }
    }

    // -------------------------------------------------------------------------
    // Feature 2: Connector shapes
    // -------------------------------------------------------------------------

    #[test]
    fn test_add_connector_basic() {
        use crate::ConnectorExt;
        use crate::ext::GroupShapeExt;
        use std::io::Cursor;

        let mut pres = PresentationBuilder::new();
        let slide = pres.add_slide();
        slide.add_connector(0, 0, 1828800, 914400, None);

        assert!(slide.has_connectors(), "slide should have a connector");

        let mut buffer = Cursor::new(Vec::new());
        pres.write(&mut buffer).unwrap();

        buffer.set_position(0);
        let mut presentation = crate::Presentation::from_reader(buffer).unwrap();
        let read_slide = presentation.slide(0).unwrap();

        // Connectors are read back from the shape tree via inner().
        let tree = &read_slide.inner().common_slide_data.shape_tree;
        let connectors = tree.connectors();
        assert_eq!(connectors.len(), 1, "expected 1 connector");
        assert!(connectors[0].name().starts_with("Connector"));
    }

    #[test]
    fn test_add_connector_with_color() {
        use crate::ext::GroupShapeExt;
        use std::io::Cursor;

        let mut pres = PresentationBuilder::new();
        let slide = pres.add_slide();
        slide.add_connector(914400, 0, 914400, 1828800, Some("FF0000"));

        let mut buffer = Cursor::new(Vec::new());
        pres.write(&mut buffer).unwrap();

        buffer.set_position(0);
        let mut presentation = crate::Presentation::from_reader(buffer).unwrap();
        let read_slide = presentation.slide(0).unwrap();

        let tree = &read_slide.inner().common_slide_data.shape_tree;
        let connectors = tree.connectors();
        assert_eq!(connectors.len(), 1);

        // Verify the line color is set on the connector.
        let sp_pr = &connectors[0].shape_properties;
        let line = sp_pr.line.as_ref().expect("connector has line");
        let line_fill = line.line_fill_properties.as_ref().expect("line has fill");
        if let ooxml_dml::types::EGLineFillProperties::SolidFill(solid) = line_fill.as_ref()
            && let Some(color) = &solid.color_choice
            && let ooxml_dml::types::EGColorChoice::SrgbClr(srgb) = color.as_ref()
        {
            assert_eq!(srgb.value, vec![0xFF, 0x00, 0x00]);
            return;
        }
        panic!("Expected connector line solidFill/srgbClr with FF0000");
    }

    #[test]
    fn test_connector_geometry_is_line() {
        use crate::ext::GroupShapeExt;
        use std::io::Cursor;

        let mut pres = PresentationBuilder::new();
        let slide = pres.add_slide();
        slide.add_connector(0, 0, 914400, 914400, None);

        let mut buffer = Cursor::new(Vec::new());
        pres.write(&mut buffer).unwrap();

        buffer.set_position(0);
        let mut presentation = crate::Presentation::from_reader(buffer).unwrap();
        let read_slide = presentation.slide(0).unwrap();

        let tree = &read_slide.inner().common_slide_data.shape_tree;
        let connectors = tree.connectors();
        let sp_pr = &connectors[0].shape_properties;
        let geom = sp_pr.geometry.as_ref().expect("connector has geometry");
        if let ooxml_dml::types::EGGeometry::PrstGeom(prst) = geom.as_ref() {
            assert_eq!(prst.preset, ooxml_dml::types::STShapeType::Line);
        } else {
            panic!("Expected PrstGeom line geometry for connector");
        }
    }

    // -------------------------------------------------------------------------
    // Feature 3: Group shapes
    // -------------------------------------------------------------------------

    #[test]
    fn test_group_shape_basic() {
        use crate::ext::GroupShapeExt;
        use std::io::Cursor;

        let mut pres = PresentationBuilder::new();
        let slide = pres.add_slide();
        slide
            .begin_group()
            .set_name("My Group")
            .add_text("Shape A", 0, 0, 914400, 457200)
            .add_text("Shape B", 0, 914400, 914400, 457200)
            .finish();

        assert!(slide.has_groups(), "slide should have a group");

        let mut buffer = Cursor::new(Vec::new());
        pres.write(&mut buffer).unwrap();

        buffer.set_position(0);
        let mut presentation = crate::Presentation::from_reader(buffer).unwrap();
        let read_slide = presentation.slide(0).unwrap();

        // The slide's shape tree should contain a group.
        let tree = &read_slide.inner().common_slide_data.shape_tree;
        let groups = tree.group_shapes();
        assert_eq!(groups.len(), 1, "expected 1 group shape");
        assert_eq!(groups[0].name(), "My Group");
        assert_eq!(groups[0].shapes().len(), 2, "expected 2 shapes in group");
    }

    #[test]
    fn test_group_shape_with_connector() {
        use crate::ext::GroupShapeExt;
        use std::io::Cursor;

        let mut pres = PresentationBuilder::new();
        let slide = pres.add_slide();
        slide
            .begin_group()
            .add_text("Box", 0, 0, 914400, 457200)
            .add_connector(457200, 457200, 914400, 914400)
            .finish();

        let mut buffer = Cursor::new(Vec::new());
        pres.write(&mut buffer).unwrap();

        buffer.set_position(0);
        let mut presentation = crate::Presentation::from_reader(buffer).unwrap();
        let read_slide = presentation.slide(0).unwrap();

        let tree = &read_slide.inner().common_slide_data.shape_tree;
        let groups = tree.group_shapes();
        assert_eq!(groups.len(), 1);
        assert_eq!(groups[0].shapes().len(), 1);
        assert_eq!(groups[0].connectors().len(), 1);
    }

    // -------------------------------------------------------------------------
    // Feature 5: Text paragraph alignment
    // -------------------------------------------------------------------------

    #[test]
    fn test_text_align_center() {
        use std::io::Cursor;

        let mut pres = PresentationBuilder::new();
        let slide = pres.add_slide();
        slide
            .shape(
                vec![TextRun::text("Centered text")],
                457200,
                457200,
                7315200,
                457200,
            )
            .set_text_align(TextAlign::Center)
            .add();

        let mut buffer = Cursor::new(Vec::new());
        pres.write(&mut buffer).unwrap();

        buffer.set_position(0);
        let mut presentation = crate::Presentation::from_reader(buffer).unwrap();
        let read_slide = presentation.slide(0).unwrap();

        let shapes = read_slide.shapes();
        assert!(!shapes.is_empty());
        let text_body = shapes[0].text_body().expect("shape has text body");
        let para = &text_body.p[0];
        let p_pr = para.p_pr.as_ref().expect("paragraph has pPr");
        assert_eq!(
            p_pr.algn,
            Some(ooxml_dml::types::STTextAlignType::Ctr),
            "expected center alignment"
        );
    }

    #[test]
    fn test_text_align_right() {
        use std::io::Cursor;

        let mut pres = PresentationBuilder::new();
        let slide = pres.add_slide();
        slide
            .shape(
                vec![TextRun::text("Right aligned")],
                457200,
                457200,
                7315200,
                457200,
            )
            .set_text_align(TextAlign::Right)
            .add();

        let mut buffer = Cursor::new(Vec::new());
        pres.write(&mut buffer).unwrap();

        buffer.set_position(0);
        let mut presentation = crate::Presentation::from_reader(buffer).unwrap();
        let read_slide = presentation.slide(0).unwrap();

        let shapes = read_slide.shapes();
        let text_body = shapes[0].text_body().expect("shape has text body");
        let p_pr = text_body.p[0].p_pr.as_ref().expect("paragraph has pPr");
        assert_eq!(p_pr.algn, Some(ooxml_dml::types::STTextAlignType::R));
    }

    // -------------------------------------------------------------------------
    // Feature 6: Multiple paragraphs in a shape
    // -------------------------------------------------------------------------

    #[test]
    fn test_multiple_paragraphs_in_shape() {
        use std::io::Cursor;

        let mut pres = PresentationBuilder::new();
        let slide = pres.add_slide();
        slide
            .shape(
                vec![TextRun::text("First paragraph")],
                457200,
                457200,
                7315200,
                1828800,
            )
            .add_paragraph(Paragraph::new("Second paragraph"))
            .add_paragraph(Paragraph::new("Third paragraph"))
            .add();

        let mut buffer = Cursor::new(Vec::new());
        pres.write(&mut buffer).unwrap();

        buffer.set_position(0);
        let mut presentation = crate::Presentation::from_reader(buffer).unwrap();
        let read_slide = presentation.slide(0).unwrap();

        let shapes = read_slide.shapes();
        assert!(!shapes.is_empty());
        let text_body = shapes[0].text_body().expect("shape has text body");
        assert_eq!(
            text_body.p.len(),
            3,
            "expected 3 paragraphs in shape, got {}",
            text_body.p.len()
        );
    }

    #[test]
    fn test_multiple_paragraphs_with_alignment() {
        use std::io::Cursor;

        let mut pres = PresentationBuilder::new();
        let slide = pres.add_slide();
        slide
            .shape(
                vec![TextRun::text("Left")],
                457200,
                457200,
                7315200,
                1828800,
            )
            .set_text_align(TextAlign::Left)
            .add_paragraph(Paragraph::new("Center").set_align(TextAlign::Center))
            .add_paragraph(Paragraph::new("Right").set_align(TextAlign::Right))
            .add();

        let mut buffer = Cursor::new(Vec::new());
        pres.write(&mut buffer).unwrap();

        buffer.set_position(0);
        let mut presentation = crate::Presentation::from_reader(buffer).unwrap();
        let read_slide = presentation.slide(0).unwrap();

        let shapes = read_slide.shapes();
        let text_body = shapes[0].text_body().expect("shape has text body");
        assert_eq!(text_body.p.len(), 3, "expected 3 paragraphs");

        let p0_algn = text_body.p[0].p_pr.as_ref().and_then(|p| p.algn);
        let p1_algn = text_body.p[1].p_pr.as_ref().and_then(|p| p.algn);
        let p2_algn = text_body.p[2].p_pr.as_ref().and_then(|p| p.algn);

        assert_eq!(p0_algn, Some(ooxml_dml::types::STTextAlignType::L));
        assert_eq!(p1_algn, Some(ooxml_dml::types::STTextAlignType::Ctr));
        assert_eq!(p2_algn, Some(ooxml_dml::types::STTextAlignType::R));
    }

    // =========================================================================
    // Feature: Multiple slide masters and layouts
    // =========================================================================

    #[test]
    fn test_add_slide_master_returns_ids() {
        let mut pres = PresentationBuilder::new();
        let (master_id, layout_ids) = pres.add_slide_master(SlideMasterConfig {
            background_color: None,
            theme_name: None,
            layouts: vec![
                SlideLayoutConfig {
                    name: "Title Slide".to_string(),
                    layout_type: SlideLayoutType::TitleSlide,
                },
                SlideLayoutConfig {
                    name: "Title and Content".to_string(),
                    layout_type: SlideLayoutType::TitleContent,
                },
            ],
        });
        assert_eq!(master_id, MasterId(1));
        assert_eq!(layout_ids.len(), 2);
        assert_eq!(layout_ids[0].master_idx, 1);
        assert_eq!(layout_ids[0].layout_idx, 0);
        assert_eq!(layout_ids[1].layout_idx, 1);
    }

    #[test]
    fn test_multiple_masters_roundtrip() {
        use std::io::Cursor;

        let mut pres = PresentationBuilder::new();
        let (_, layout_ids) = pres.add_slide_master(SlideMasterConfig {
            background_color: Some("1F497D".to_string()),
            theme_name: Some("Custom Theme".to_string()),
            layouts: vec![
                SlideLayoutConfig {
                    name: "Title Slide".to_string(),
                    layout_type: SlideLayoutType::TitleSlide,
                },
                SlideLayoutConfig {
                    name: "Title and Content".to_string(),
                    layout_type: SlideLayoutType::TitleContent,
                },
            ],
        });

        // Add a slide using the custom layout.
        let slide = pres.add_slide();
        slide.set_layout(layout_ids[0]);
        slide.add_title("Custom Master Slide");

        let mut buffer = Cursor::new(Vec::new());
        pres.write(&mut buffer).unwrap();

        // Verify it roundtrips.
        buffer.set_position(0);
        let presentation = crate::Presentation::from_reader(buffer).unwrap();
        // Default master + 1 extra = 2 masters total.
        assert_eq!(presentation.slide_masters().len(), 2);
    }

    #[test]
    fn test_slide_layout_xml_contains_placeholders() {
        // TitleSlide layout XML should contain the ctrTitle and subTitle placeholders.
        let config = SlideLayoutConfig {
            name: "My Title Layout".to_string(),
            layout_type: SlideLayoutType::TitleSlide,
        };
        let xml = build_layout_xml(&config);
        assert!(xml.contains(r#"type="ctrTitle""#), "Missing ctrTitle ph");
        assert!(xml.contains(r#"type="subTitle""#), "Missing subTitle ph");
        assert!(xml.contains("My Title Layout"), "Missing layout name");
    }

    #[test]
    fn test_slide_layout_xml_title_content() {
        let config = SlideLayoutConfig {
            name: "Title Content".to_string(),
            layout_type: SlideLayoutType::TitleContent,
        };
        let xml = build_layout_xml(&config);
        assert!(xml.contains(r#"type="title""#));
        // Content placeholder has no type attribute, just idx="1".
        assert!(xml.contains(r#"idx="1""#));
    }

    #[test]
    fn test_slide_layout_xml_two_content() {
        let config = SlideLayoutConfig {
            name: "Two Content".to_string(),
            layout_type: SlideLayoutType::TwoContent,
        };
        let xml = build_layout_xml(&config);
        assert!(xml.contains(r#"idx="1""#));
        assert!(xml.contains(r#"idx="2""#));
    }

    #[test]
    fn test_slide_layout_xml_blank() {
        let config = SlideLayoutConfig {
            name: "Blank".to_string(),
            layout_type: SlideLayoutType::Blank,
        };
        let xml = build_layout_xml(&config);
        // Blank layout should have no placeholder shapes.
        assert!(
            !xml.contains("<p:ph"),
            "Blank layout should not have placeholders"
        );
        assert!(xml.contains(r#"type="blank""#));
    }

    #[test]
    fn test_master_xml_background_color() {
        let entry = SlideMasterEntry {
            config: SlideMasterConfig {
                background_color: Some("FF0000".to_string()),
                theme_name: None,
                layouts: vec![],
            },
        };
        let xml = build_master_xml(&entry, 0);
        assert!(xml.contains("FF0000"), "Master XML should contain bg color");
        assert!(
            xml.contains("solidFill"),
            "Master XML should have solidFill"
        );
    }

    #[test]
    fn test_set_layout_writes_correct_rels() {
        use std::io::Cursor;

        let mut pres = PresentationBuilder::new();
        let (_, layout_ids) = pres.add_slide_master(SlideMasterConfig {
            background_color: None,
            theme_name: None,
            layouts: vec![SlideLayoutConfig {
                name: "Title Slide".to_string(),
                layout_type: SlideLayoutType::TitleSlide,
            }],
        });

        let slide = pres.add_slide();
        slide.set_layout(layout_ids[0]); // master 1, layout 0 → global layout 2

        let mut buffer = Cursor::new(Vec::new());
        pres.write(&mut buffer).unwrap();

        buffer.set_position(0);
        let mut package = ooxml_opc::Package::open(buffer).unwrap();
        let rels_xml = String::from_utf8(
            package
                .read_part("ppt/slides/_rels/slide1.xml.rels")
                .unwrap(),
        )
        .unwrap();
        // The slide should reference slideLayout2 (first layout of the extra master).
        assert!(
            rels_xml.contains("slideLayout2.xml"),
            "Slide rels should reference slideLayout2, got: {}",
            rels_xml
        );
    }

    // =========================================================================
    // Feature: Slide animations
    // =========================================================================

    #[cfg(feature = "pml-animations")]
    #[test]
    fn test_add_animation_stored() {
        let mut pres = PresentationBuilder::new();
        let slide = pres.add_slide();
        slide.add_title("Animated");
        slide.add_animation(AnimationConfig {
            shape_id: 2,
            effect: AnimationEffect::Appear,
            trigger: AnimationTrigger::OnClick,
            duration_ms: 500,
            delay_ms: 0,
        });
        assert!(slide.has_animations(), "slide should have animations");
    }

    #[cfg(feature = "pml-animations")]
    #[test]
    fn test_animation_timing_xml_structure() {
        let anims = vec![AnimationConfig {
            shape_id: 3,
            effect: AnimationEffect::FadeIn,
            trigger: AnimationTrigger::OnClick,
            duration_ms: 1000,
            delay_ms: 250,
        }];
        let xml = build_timing_xml(&anims);
        assert!(xml.contains("<p:timing>"), "missing <p:timing>");
        assert!(xml.contains("</p:timing>"), "missing </p:timing>");
        assert!(xml.contains("<p:tnLst>"), "missing <p:tnLst>");
        assert!(xml.contains("<p:bldLst>"), "missing <p:bldLst>");
        assert!(xml.contains(r#"spid="3""#), "bldP should reference shape 3");
        // FadeIn = presetID 10
        assert!(
            xml.contains(r#"presetID="10""#),
            "FadeIn should use presetID=10"
        );
    }

    #[cfg(feature = "pml-animations")]
    #[test]
    fn test_animation_roundtrip_writes_timing() {
        use std::io::Cursor;

        let mut pres = PresentationBuilder::new();
        let slide = pres.add_slide();
        slide.add_title("Animated Slide");
        slide.add_animation(AnimationConfig {
            shape_id: 2,
            effect: AnimationEffect::FlyInFromLeft,
            trigger: AnimationTrigger::OnClick,
            duration_ms: 750,
            delay_ms: 0,
        });

        let mut buffer = Cursor::new(Vec::new());
        pres.write(&mut buffer).unwrap();

        // Verify the timing element appears in the slide XML.
        buffer.set_position(0);
        let mut package = ooxml_opc::Package::open(buffer).unwrap();
        let slide_xml =
            String::from_utf8(package.read_part("ppt/slides/slide1.xml").unwrap()).unwrap();
        assert!(
            slide_xml.contains("timing") || slide_xml.contains("tnLst"),
            "Slide XML should contain timing element; got: {}",
            &slide_xml[..slide_xml.len().min(500)]
        );
    }

    #[cfg(feature = "pml-animations")]
    #[test]
    fn test_animation_multiple_effects() {
        let anims = vec![
            AnimationConfig {
                shape_id: 2,
                effect: AnimationEffect::Appear,
                trigger: AnimationTrigger::OnClick,
                duration_ms: 500,
                delay_ms: 0,
            },
            AnimationConfig {
                shape_id: 3,
                effect: AnimationEffect::FadeIn,
                trigger: AnimationTrigger::AfterPrevious,
                duration_ms: 1000,
                delay_ms: 500,
            },
        ];
        let xml = build_timing_xml(&anims);
        // Two bldP entries.
        assert_eq!(
            xml.matches("<p:bldP").count(),
            2,
            "Expected 2 bldP entries for 2 animations"
        );
        // Both shape refs present.
        assert!(xml.contains(r#"spid="2""#));
        assert!(xml.contains(r#"spid="3""#));
    }

    #[cfg(feature = "pml-animations")]
    #[test]
    fn test_animation_trigger_node_types() {
        let click_anim = vec![AnimationConfig {
            shape_id: 2,
            effect: AnimationEffect::Appear,
            trigger: AnimationTrigger::OnClick,
            duration_ms: 500,
            delay_ms: 0,
        }];
        let xml = build_timing_xml(&click_anim);
        assert!(xml.contains(r#"nodeType="clickEffect""#));

        let with_anim = vec![AnimationConfig {
            shape_id: 2,
            effect: AnimationEffect::Appear,
            trigger: AnimationTrigger::WithPrevious,
            duration_ms: 500,
            delay_ms: 0,
        }];
        let xml2 = build_timing_xml(&with_anim);
        assert!(xml2.contains(r#"nodeType="withEffect""#));

        let after_anim = vec![AnimationConfig {
            shape_id: 2,
            effect: AnimationEffect::Appear,
            trigger: AnimationTrigger::AfterPrevious,
            duration_ms: 500,
            delay_ms: 0,
        }];
        let xml3 = build_timing_xml(&after_anim);
        assert!(xml3.contains(r#"nodeType="afterEffect""#));
    }

    // =========================================================================
    // Feature: Notes master
    // =========================================================================

    #[test]
    fn test_notes_master_written_when_slide_has_notes() {
        use std::io::Cursor;

        let mut pres = PresentationBuilder::new();
        let slide = pres.add_slide();
        slide.add_title("Notes Test");
        slide.set_notes("These are speaker notes.");

        let mut buffer = Cursor::new(Vec::new());
        pres.write(&mut buffer).unwrap();

        buffer.set_position(0);
        let mut package = ooxml_opc::Package::open(buffer).unwrap();

        // Notes master XML should exist.
        let nm_xml = package.read_part("ppt/notesMasters/notesMaster1.xml");
        assert!(
            nm_xml.is_ok(),
            "notesMaster1.xml should exist when slide has notes"
        );
        let nm_str = String::from_utf8(nm_xml.unwrap()).unwrap();
        assert!(
            nm_str.contains("notesMaster"),
            "Wrong notes master root element"
        );

        // Notes master rels should exist.
        let nm_rels = package.read_part("ppt/notesMasters/_rels/notesMaster1.xml.rels");
        assert!(nm_rels.is_ok(), "notesMaster rels should exist");

        // presentation.xml.rels should reference the notes master.
        let pres_rels = String::from_utf8(
            package
                .read_part("ppt/_rels/presentation.xml.rels")
                .unwrap(),
        )
        .unwrap();
        assert!(
            pres_rels.contains(REL_NOTES_MASTER),
            "presentation.xml.rels should reference notes master"
        );
    }

    #[test]
    fn test_notes_slide_rels_reference_notes_master() {
        use std::io::Cursor;

        let mut pres = PresentationBuilder::new();
        let slide = pres.add_slide();
        slide.set_notes("Speaker notes here.");

        let mut buffer = Cursor::new(Vec::new());
        pres.write(&mut buffer).unwrap();

        buffer.set_position(0);
        let mut package = ooxml_opc::Package::open(buffer).unwrap();

        let notes_rels = String::from_utf8(
            package
                .read_part("ppt/notesSlides/_rels/notesSlide1.xml.rels")
                .unwrap(),
        )
        .unwrap();
        assert!(
            notes_rels.contains(REL_NOTES_MASTER),
            "notes slide rels should reference notes master"
        );
    }

    #[test]
    fn test_no_notes_master_when_no_notes() {
        use std::io::Cursor;

        let mut pres = PresentationBuilder::new();
        let slide = pres.add_slide();
        slide.add_title("No Notes");

        let mut buffer = Cursor::new(Vec::new());
        pres.write(&mut buffer).unwrap();

        buffer.set_position(0);
        let mut package = ooxml_opc::Package::open(buffer).unwrap();
        let nm = package.read_part("ppt/notesMasters/notesMaster1.xml");
        assert!(
            nm.is_err(),
            "Should not write notes master when no slide has notes"
        );
    }

    // =========================================================================
    // Feature: Chart embedding
    // =========================================================================

    #[cfg(feature = "pml-charts")]
    #[test]
    fn test_embed_chart_stored() {
        let mut pres = PresentationBuilder::new();
        let slide = pres.add_slide();
        slide.embed_chart(b"<c:chartSpace/>".to_vec(), 0, 0, 4572000, 3429000);
        assert!(slide.has_charts());
    }

    #[cfg(feature = "pml-charts")]
    #[test]
    fn test_embed_chart_writes_chart_part() {
        use std::io::Cursor;

        let chart_xml = b"<?xml version=\"1.0\"?><c:chartSpace xmlns:c=\"http://schemas.openxmlformats.org/drawingml/2006/chart\"/>";

        let mut pres = PresentationBuilder::new();
        let slide = pres.add_slide();
        slide.embed_chart(chart_xml.to_vec(), 457200, 457200, 4572000, 3429000);

        let mut buffer = Cursor::new(Vec::new());
        pres.write(&mut buffer).unwrap();

        buffer.set_position(0);
        let mut package = ooxml_opc::Package::open(buffer).unwrap();

        // chart1.xml should be written.
        let chart_data = package.read_part("ppt/charts/chart1.xml");
        assert!(chart_data.is_ok(), "ppt/charts/chart1.xml should exist");

        // The slide rels should reference the chart.
        let slide_rels = String::from_utf8(
            package
                .read_part("ppt/slides/_rels/slide1.xml.rels")
                .unwrap(),
        )
        .unwrap();
        assert!(
            slide_rels.contains(REL_CHART),
            "slide rels should contain chart relationship type"
        );
        assert!(
            slide_rels.contains("chart1.xml"),
            "slide rels should reference chart1.xml"
        );
    }

    #[cfg(feature = "pml-charts")]
    #[test]
    fn test_embed_chart_graphic_frame_in_slide_xml() {
        use std::io::Cursor;

        let chart_xml = b"<?xml version=\"1.0\"?><c:chartSpace/>";

        let mut pres = PresentationBuilder::new();
        let slide = pres.add_slide();
        slide.embed_chart(chart_xml.to_vec(), 0, 0, 4572000, 3429000);

        let mut buffer = Cursor::new(Vec::new());
        pres.write(&mut buffer).unwrap();

        buffer.set_position(0);
        let mut package = ooxml_opc::Package::open(buffer).unwrap();
        let slide_xml =
            String::from_utf8(package.read_part("ppt/slides/slide1.xml").unwrap()).unwrap();

        // Slide XML should contain graphicFrame and chart data URI.
        assert!(
            slide_xml.contains("graphicFrame") || slide_xml.contains("graphicData"),
            "Slide XML should contain graphicFrame for chart"
        );
        assert!(
            slide_xml.contains("drawingml/2006/chart"),
            "Slide XML should reference chart schema URI"
        );
    }

    #[cfg(feature = "pml-charts")]
    #[test]
    fn test_embed_multiple_charts() {
        use std::io::Cursor;

        let chart_xml = b"<?xml version=\"1.0\"?><c:chartSpace/>";

        let mut pres = PresentationBuilder::new();
        let slide = pres.add_slide();
        slide.embed_chart(chart_xml.to_vec(), 0, 0, 4572000, 3429000);
        slide.embed_chart(chart_xml.to_vec(), 5000000, 0, 4572000, 3429000);

        let mut buffer = Cursor::new(Vec::new());
        pres.write(&mut buffer).unwrap();

        buffer.set_position(0);
        let mut package = ooxml_opc::Package::open(buffer).unwrap();

        assert!(
            package.read_part("ppt/charts/chart1.xml").is_ok(),
            "chart1.xml should exist"
        );
        assert!(
            package.read_part("ppt/charts/chart2.xml").is_ok(),
            "chart2.xml should exist"
        );
    }

    // =========================================================================
    // Feature: SmartArt embedding
    // =========================================================================

    #[cfg(feature = "pml-charts")]
    #[test]
    fn test_embed_smartart_stored() {
        let mut pres = PresentationBuilder::new();
        let slide = pres.add_slide();
        slide.embed_smartart(b"<dgm:dataModel/>".to_vec(), None, 0, 0, 4572000, 3429000);
        assert!(slide.has_smartarts());
    }

    #[cfg(feature = "pml-charts")]
    #[test]
    fn test_embed_smartart_writes_data_part() {
        use std::io::Cursor;

        let data_xml = b"<?xml version=\"1.0\"?><dgm:dataModel xmlns:dgm=\"http://schemas.openxmlformats.org/drawingml/2006/diagram\"/>";

        let mut pres = PresentationBuilder::new();
        let slide = pres.add_slide();
        slide.embed_smartart(data_xml.to_vec(), None, 0, 0, 4572000, 3429000);

        let mut buffer = Cursor::new(Vec::new());
        pres.write(&mut buffer).unwrap();

        buffer.set_position(0);
        let mut package = ooxml_opc::Package::open(buffer).unwrap();

        // Diagram data part should exist.
        let data_part = package.read_part("ppt/diagrams/data1.xml");
        assert!(data_part.is_ok(), "ppt/diagrams/data1.xml should exist");

        // The slide rels should reference the diagram data.
        let slide_rels = String::from_utf8(
            package
                .read_part("ppt/slides/_rels/slide1.xml.rels")
                .unwrap(),
        )
        .unwrap();
        assert!(
            slide_rels.contains(REL_DIAGRAM_DATA),
            "slide rels should contain diagramData relationship type"
        );
    }

    #[cfg(feature = "pml-charts")]
    #[test]
    fn test_embed_smartart_with_layout() {
        use std::io::Cursor;

        let data_xml = b"<?xml version=\"1.0\"?><dgm:dataModel/>";
        let layout_xml = b"<?xml version=\"1.0\"?><dgm:layout/>";

        let mut pres = PresentationBuilder::new();
        let slide = pres.add_slide();
        slide.embed_smartart(
            data_xml.to_vec(),
            Some(layout_xml.to_vec()),
            0,
            0,
            4572000,
            3429000,
        );

        let mut buffer = Cursor::new(Vec::new());
        pres.write(&mut buffer).unwrap();

        buffer.set_position(0);
        let mut package = ooxml_opc::Package::open(buffer).unwrap();

        // Both data and layout parts should exist.
        assert!(package.read_part("ppt/diagrams/data1.xml").is_ok());
        assert!(package.read_part("ppt/diagrams/layout2.xml").is_ok());

        // The slide rels should reference both.
        let slide_rels = String::from_utf8(
            package
                .read_part("ppt/slides/_rels/slide1.xml.rels")
                .unwrap(),
        )
        .unwrap();
        assert!(slide_rels.contains(REL_DIAGRAM_DATA));
        assert!(slide_rels.contains(REL_DIAGRAM_LAYOUT));
    }

    #[cfg(feature = "pml-charts")]
    #[test]
    fn test_embed_smartart_graphic_frame_in_slide_xml() {
        use std::io::Cursor;

        let data_xml = b"<?xml version=\"1.0\"?><dgm:dataModel/>";

        let mut pres = PresentationBuilder::new();
        let slide = pres.add_slide();
        slide.embed_smartart(data_xml.to_vec(), None, 457200, 457200, 4572000, 3429000);

        let mut buffer = Cursor::new(Vec::new());
        pres.write(&mut buffer).unwrap();

        buffer.set_position(0);
        let mut package = ooxml_opc::Package::open(buffer).unwrap();
        let slide_xml =
            String::from_utf8(package.read_part("ppt/slides/slide1.xml").unwrap()).unwrap();

        // Slide XML should contain the diagram URI.
        assert!(
            slide_xml.contains("drawingml/2006/diagram"),
            "Slide XML should reference diagram schema URI"
        );
        assert!(
            slide_xml.contains("dgm:relIds") || slide_xml.contains("relIds"),
            "Slide XML should contain dgm:relIds"
        );
    }
}
