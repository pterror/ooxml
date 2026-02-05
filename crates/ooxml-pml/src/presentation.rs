//! Presentation API for reading and writing PowerPoint files.
//!
//! This module provides the main entry point for working with PPTX files.

use crate::error::{Error, Result};
use crate::ext::{CommonSlideDataExt, GroupShapeExt, PictureExt, ShapeExt};
use crate::parsers::FromXml;
use crate::types;
use ooxml_dml::ext::{TextBodyExt, TextParagraphExt, TextRunExt};
use ooxml_opc::{Package, Relationships};
use quick_xml::Reader;
use quick_xml::events::Event;
use std::fs::File;
use std::io::{BufReader, Cursor, Read, Seek};
use std::path::Path;

// Relationship types (ECMA-376 Part 1)
const REL_OFFICE_DOCUMENT: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument";
const REL_SLIDE: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide";
const REL_NOTES_SLIDE: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesSlide";
const REL_SLIDE_MASTER: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster";
const REL_SLIDE_LAYOUT: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout";

/// A PowerPoint presentation.
///
/// This is the main entry point for reading PPTX files.
pub struct Presentation<R: Read + Seek> {
    package: Package<R>,
    /// Path to the presentation part.
    #[allow(dead_code)]
    presentation_path: String,
    /// Presentation-level relationships.
    #[allow(dead_code)]
    pres_rels: Relationships,
    /// Slide metadata (relationship ID, path).
    slide_info: Vec<SlideInfo>,
    /// Slide masters in the presentation.
    slide_masters: Vec<SlideMaster>,
    /// Slide layouts in the presentation.
    slide_layouts: Vec<SlideLayout>,
}

/// Metadata about a slide.
#[derive(Debug, Clone)]
struct SlideInfo {
    #[allow(dead_code)]
    rel_id: String,
    path: String,
    index: usize,
    /// Relationship ID to the slide layout.
    layout_rel_id: Option<String>,
}

/// Image data loaded from the presentation.
#[derive(Debug, Clone)]
pub struct ImageData {
    /// The raw image data.
    pub data: Vec<u8>,
    /// The content type (MIME type) of the image.
    pub content_type: String,
}

/// A slide master in the presentation.
///
/// Slide masters define the overall theme and formatting for slides.
/// ECMA-376 Part 1, Section 19.3.1.42 (sldMaster).
#[derive(Debug, Clone)]
pub struct SlideMaster {
    /// Path to the slide master part.
    path: String,
    /// Name of the slide master (if specified).
    pub name: Option<String>,
    /// Relationship IDs of layouts using this master.
    layout_ids: Vec<String>,
    /// Color scheme name.
    pub color_scheme: Option<String>,
    /// Background color (ARGB).
    pub background_color: Option<String>,
}

impl SlideMaster {
    /// Get the path to this slide master.
    pub fn path(&self) -> &str {
        &self.path
    }

    /// Get the number of layouts using this master.
    pub fn layout_count(&self) -> usize {
        self.layout_ids.len()
    }
}

/// A slide layout in the presentation.
///
/// Slide layouts define the arrangement of content placeholders.
/// ECMA-376 Part 1, Section 19.3.1.39 (sldLayout).
#[derive(Debug, Clone)]
pub struct SlideLayout {
    /// Path to the slide layout part.
    path: String,
    /// Name of the layout (e.g., "Title Slide", "Title and Content").
    pub name: Option<String>,
    /// Layout type.
    pub layout_type: SlideLayoutType,
    /// Relationship ID to the slide master.
    #[allow(dead_code)]
    master_rel_id: Option<String>,
    /// Whether to match slide names.
    pub match_name: bool,
    /// Whether to show master shapes.
    pub show_master_shapes: bool,
}

impl SlideLayout {
    /// Get the path to this slide layout.
    pub fn path(&self) -> &str {
        &self.path
    }
}

/// Type of slide layout.
///
/// ECMA-376 Part 1, Section 19.7.15 (ST_SlideLayoutType).
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub enum SlideLayoutType {
    /// Blank slide.
    Blank,
    /// Title slide.
    #[default]
    Title,
    /// Title and content.
    TitleAndContent,
    /// Section header.
    SectionHeader,
    /// Two content.
    TwoContent,
    /// Two content and text.
    TwoContentAndText,
    /// Title only.
    TitleOnly,
    /// Content with caption.
    ContentWithCaption,
    /// Picture with caption.
    PictureWithCaption,
    /// Vertical title and text.
    VerticalTitleAndText,
    /// Vertical text.
    VerticalText,
    /// Custom layout.
    Custom,
    /// Unknown layout type.
    Unknown,
}

impl SlideLayoutType {
    /// Parse from the slideLayout type attribute.
    fn parse(s: &str) -> Self {
        match s {
            "blank" => Self::Blank,
            "title" | "tx" => Self::Title,
            "obj" | "objTx" | "twoObj" | "twoObjAndTx" => Self::TitleAndContent,
            "secHead" => Self::SectionHeader,
            "twoTxTwoObj" => Self::TwoContent,
            "objAndTx" => Self::TwoContentAndText,
            "titleOnly" => Self::TitleOnly,
            "objOnly" => Self::ContentWithCaption,
            "picTx" => Self::PictureWithCaption,
            "vertTx" => Self::VerticalText,
            "vertTitleAndTx" => Self::VerticalTitleAndText,
            "cust" => Self::Custom,
            _ => Self::Unknown,
        }
    }
}

impl Presentation<BufReader<File>> {
    /// Open a presentation from a file path.
    pub fn open<P: AsRef<Path>>(path: P) -> Result<Self> {
        let file = File::open(path)?;
        Self::from_reader(BufReader::new(file))
    }
}

impl<R: Read + Seek> Presentation<R> {
    /// Open a presentation from a reader.
    pub fn from_reader(reader: R) -> Result<Self> {
        let mut package = Package::open(reader)?;

        // Find the presentation part via root relationships
        let root_rels = package.read_relationships()?;
        let pres_rel = root_rels
            .get_by_type(REL_OFFICE_DOCUMENT)
            .ok_or_else(|| Error::Invalid("Missing presentation relationship".into()))?;
        let presentation_path = pres_rel.target.clone();

        // Load presentation relationships
        let pres_rels = package
            .read_part_relationships(&presentation_path)
            .unwrap_or_default();

        // Parse presentation.xml to get slide list
        let pres_xml = package.read_part(&presentation_path)?;
        let slide_order = parse_presentation_slides(&pres_xml)?;

        // Load slide masters
        let mut slide_masters: Vec<SlideMaster> = Vec::new();
        let mut slide_layouts: Vec<SlideLayout> = Vec::new();

        for rel in pres_rels.iter() {
            if rel.relationship_type == REL_SLIDE_MASTER {
                let path = resolve_path(&presentation_path, &rel.target);
                if let Ok(master_xml) = package.read_part(&path) {
                    let master = parse_slide_master(&master_xml, &path);
                    let master_path = path.clone();

                    // Load layouts for this master
                    if let Ok(master_rels) = package.read_part_relationships(&path) {
                        for layout_rel in master_rels.iter() {
                            if layout_rel.relationship_type == REL_SLIDE_LAYOUT {
                                let layout_path = resolve_path(&master_path, &layout_rel.target);
                                if let Ok(layout_xml) = package.read_part(&layout_path) {
                                    let layout = parse_slide_layout(
                                        &layout_xml,
                                        &layout_path,
                                        Some(layout_rel.id.clone()),
                                    );
                                    slide_layouts.push(layout);
                                }
                            }
                        }
                    }

                    slide_masters.push(master);
                }
            }
        }

        // Build slide info from relationships, getting layout references from slide XML
        let mut slide_info: Vec<SlideInfo> = Vec::new();
        for rel in pres_rels.iter() {
            if rel.relationship_type == REL_SLIDE {
                let path = resolve_path(&presentation_path, &rel.target);
                // Find index from slide order
                let index = slide_order
                    .iter()
                    .position(|id| id == &rel.id)
                    .unwrap_or(slide_info.len());

                // Get layout relationship from slide
                let layout_rel_id = if let Ok(slide_rels) = package.read_part_relationships(&path) {
                    slide_rels
                        .get_by_type(REL_SLIDE_LAYOUT)
                        .map(|r| r.id.clone())
                } else {
                    None
                };

                slide_info.push(SlideInfo {
                    rel_id: rel.id.clone(),
                    path,
                    index,
                    layout_rel_id,
                });
            }
        }

        // Sort by index
        slide_info.sort_by_key(|s| s.index);

        Ok(Self {
            package,
            presentation_path,
            pres_rels,
            slide_info,
            slide_masters,
            slide_layouts,
        })
    }

    /// Get the number of slides in the presentation.
    pub fn slide_count(&self) -> usize {
        self.slide_info.len()
    }

    /// Get all slide masters in the presentation.
    pub fn slide_masters(&self) -> &[SlideMaster] {
        &self.slide_masters
    }

    /// Get all slide layouts in the presentation.
    pub fn slide_layouts(&self) -> &[SlideLayout] {
        &self.slide_layouts
    }

    /// Get a slide layout by name.
    pub fn layout_by_name(&self, name: &str) -> Option<&SlideLayout> {
        self.slide_layouts
            .iter()
            .find(|l| l.name.as_deref() == Some(name))
    }

    /// Get a slide by index (0-based).
    pub fn slide(&mut self, index: usize) -> Result<Slide> {
        let info = self
            .slide_info
            .get(index)
            .ok_or_else(|| Error::Invalid(format!("Slide index {} out of range", index)))?
            .clone();

        self.load_slide(&info)
    }

    /// Load all slides.
    pub fn slides(&mut self) -> Result<Vec<Slide>> {
        let infos: Vec<_> = self.slide_info.clone();
        infos.iter().map(|info| self.load_slide(info)).collect()
    }

    /// Load a slide's data.
    fn load_slide(&mut self, info: &SlideInfo) -> Result<Slide> {
        let data = self.package.read_part(&info.path)?;

        // Parse slide using generated FromXml parser
        let inner = parse_slide_xml(&data)?;

        // Extract tables from graphic frames
        let tables = extract_tables_from_slide(&inner);

        // Build the slide wrapper
        let mut slide = Slide {
            inner,
            index: info.index,
            slide_path: info.path.clone(),
            notes: None,
            layout_rel_id: info.layout_rel_id.clone(),
            tables,
        };

        // Try to load speaker notes
        if let Ok(slide_rels) = self.package.read_part_relationships(&info.path)
            && let Some(notes_rel) = slide_rels.get_by_type(REL_NOTES_SLIDE)
        {
            let notes_path = resolve_path(&info.path, &notes_rel.target);
            if let Ok(notes_data) = self.package.read_part(&notes_path) {
                slide.notes = parse_notes_slide(&notes_data);
            }
        }

        Ok(slide)
    }

    /// Get image data for a picture from a specific slide.
    ///
    /// Loads the image data from the package using the picture's relationship ID.
    pub fn get_image_data(&mut self, slide: &Slide, picture: &types::Picture) -> Result<ImageData> {
        // Get the embed relationship ID using the extension trait
        let rel_id = picture
            .embed_rel_id()
            .ok_or_else(|| Error::Invalid("Picture has no embed relationship ID".into()))?;

        // Get slide relationships
        let slide_rels = self
            .package
            .read_part_relationships(slide.slide_path())
            .map_err(|_| Error::Invalid("Failed to read slide relationships".into()))?;

        // Find the image relationship
        let rel = slide_rels
            .get(rel_id)
            .ok_or_else(|| Error::Invalid(format!("Image relationship {} not found", rel_id)))?;

        // Resolve the image path
        let image_path = resolve_path(slide.slide_path(), &rel.target);

        // Read image data
        let data = self.package.read_part(&image_path)?;

        // Determine content type from extension
        let content_type = content_type_from_path(&image_path);

        Ok(ImageData { data, content_type })
    }

    /// Resolve a hyperlink relationship ID to its target URL.
    ///
    /// # Arguments
    /// * `slide` - The slide containing the hyperlink
    /// * `rel_id` - The relationship ID from the hyperlink
    ///
    /// # Returns
    /// The target URL/path of the hyperlink, or an error if not found.
    pub fn resolve_hyperlink(&mut self, slide: &Slide, rel_id: &str) -> Result<String> {
        // Get slide relationships
        let slide_rels = self
            .package
            .read_part_relationships(slide.slide_path())
            .map_err(|_| Error::Invalid("Failed to read slide relationships".into()))?;

        // Find the hyperlink relationship
        let rel = slide_rels.get(rel_id).ok_or_else(|| {
            Error::Invalid(format!("Hyperlink relationship {} not found", rel_id))
        })?;

        Ok(rel.target.clone())
    }

    /// Get all hyperlinks from a slide with their resolved URLs.
    ///
    /// Returns a list of (text, url) pairs for all hyperlinks on the slide.
    pub fn get_hyperlinks_with_urls(&mut self, slide: &Slide) -> Result<Vec<(String, String)>> {
        let hyperlinks = slide.hyperlinks();
        let mut results = Vec::new();

        for link in hyperlinks {
            if let Ok(url) = self.resolve_hyperlink(slide, &link.rel_id) {
                results.push((link.text, url));
            }
        }

        Ok(results)
    }
}

/// A slide in the presentation.
///
/// This wraps the generated `types::Slide` and provides additional context
/// (index, path, notes) that comes from the package structure rather than
/// the slide XML itself.
#[derive(Debug, Clone)]
pub struct Slide {
    /// The parsed slide content (generated from PML schema).
    inner: types::Slide,
    /// Slide index (0-based).
    index: usize,
    /// Path to this slide part (for resolving relationships).
    slide_path: String,
    /// Speaker notes for this slide (parsed from separate notes part).
    notes: Option<String>,
    /// Relationship ID to the slide layout.
    layout_rel_id: Option<String>,
    /// Tables extracted from graphic frames.
    tables: Vec<Table>,
}

/// Slide transition effect.
///
/// Represents the animation effect when advancing to this slide.
#[derive(Debug, Clone, Default)]
pub struct Transition {
    /// Transition type (fade, push, wipe, etc.)
    pub transition_type: Option<TransitionType>,
    /// Transition speed.
    pub speed: TransitionSpeed,
    /// Advance on mouse click.
    pub advance_on_click: bool,
    /// Auto-advance time in milliseconds (if set).
    pub advance_time_ms: Option<u32>,
}

/// Type of slide transition effect.
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum TransitionType {
    /// Fade transition.
    Fade,
    /// Push transition.
    Push,
    /// Wipe transition.
    Wipe,
    /// Split transition.
    Split,
    /// Blinds transition.
    Blinds,
    /// Checker transition.
    Checker,
    /// Circle transition.
    Circle,
    /// Dissolve transition.
    Dissolve,
    /// Comb transition.
    Comb,
    /// Cover transition.
    Cover,
    /// Cut transition.
    Cut,
    /// Diamond transition.
    Diamond,
    /// Plus transition.
    Plus,
    /// Random transition.
    Random,
    /// Strips transition.
    Strips,
    /// Wedge transition.
    Wedge,
    /// Wheel transition.
    Wheel,
    /// Zoom transition.
    Zoom,
    /// Unknown/unsupported transition type.
    Other(String),
}

impl TransitionType {
    /// Convert to XML element name.
    pub fn to_xml_value(&self) -> &str {
        match self {
            TransitionType::Fade => "fade",
            TransitionType::Push => "push",
            TransitionType::Wipe => "wipe",
            TransitionType::Split => "split",
            TransitionType::Blinds => "blinds",
            TransitionType::Checker => "checker",
            TransitionType::Circle => "circle",
            TransitionType::Dissolve => "dissolve",
            TransitionType::Comb => "comb",
            TransitionType::Cover => "cover",
            TransitionType::Cut => "cut",
            TransitionType::Diamond => "diamond",
            TransitionType::Plus => "plus",
            TransitionType::Random => "random",
            TransitionType::Strips => "strips",
            TransitionType::Wedge => "wedge",
            TransitionType::Wheel => "wheel",
            TransitionType::Zoom => "zoom",
            TransitionType::Other(name) => name.as_str(),
        }
    }
}

/// Speed of the slide transition.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub enum TransitionSpeed {
    /// Slow transition.
    Slow,
    /// Medium transition (default).
    #[default]
    Medium,
    /// Fast transition.
    Fast,
}

impl TransitionSpeed {
    /// Convert to XML attribute value.
    pub fn to_xml_value(self) -> &'static str {
        match self {
            TransitionSpeed::Slow => "slow",
            TransitionSpeed::Medium => "med",
            TransitionSpeed::Fast => "fast",
        }
    }
}

impl Slide {
    /// Get the slide index (0-based).
    pub fn index(&self) -> usize {
        self.index
    }

    /// Get all shapes on the slide.
    ///
    /// Returns the generated `types::Shape` structs. Use the `ShapeExt` trait
    /// for convenient accessor methods like `name()`, `text()`, `paragraphs()`.
    pub fn shapes(&self) -> &[types::Shape] {
        self.inner.common_slide_data.shape_tree.shapes()
    }

    /// Extract all text from the slide.
    pub fn text(&self) -> String {
        self.inner.common_slide_data.text()
    }

    /// Get the speaker notes for this slide.
    pub fn notes(&self) -> Option<&str> {
        self.notes.as_deref()
    }

    /// Check if this slide has speaker notes.
    pub fn has_notes(&self) -> bool {
        self.notes.as_ref().is_some_and(|n| !n.is_empty())
    }

    /// Get all pictures on the slide.
    ///
    /// Returns the generated `types::Picture` structs. Use the `PictureExt` trait
    /// for convenient accessor methods like `name()`, `description()`, `embed_rel_id()`.
    pub fn pictures(&self) -> &[types::Picture] {
        self.inner.common_slide_data.pictures()
    }

    /// Get the path to this slide part (for resolving image relationships).
    pub(crate) fn slide_path(&self) -> &str {
        &self.slide_path
    }

    /// Get the slide transition effect (if any).
    #[cfg(feature = "pml-transitions")]
    pub fn transition(&self) -> Option<Transition> {
        self.inner
            .transition
            .as_ref()
            .map(|t| convert_transition(t))
    }

    /// Get the slide transition effect (if any).
    #[cfg(not(feature = "pml-transitions"))]
    pub fn transition(&self) -> Option<Transition> {
        None
    }

    /// Check if this slide has a transition effect.
    #[cfg(feature = "pml-transitions")]
    pub fn has_transition(&self) -> bool {
        self.inner.transition.is_some()
    }

    /// Check if this slide has a transition effect.
    #[cfg(not(feature = "pml-transitions"))]
    pub fn has_transition(&self) -> bool {
        false
    }

    /// Get the relationship ID to the slide layout used by this slide.
    ///
    /// This can be used to look up the layout in the presentation's slide_layouts.
    pub fn layout_rel_id(&self) -> Option<&str> {
        self.layout_rel_id.as_deref()
    }

    /// Get all hyperlinks from all shapes on this slide.
    ///
    /// Returns hyperlinks with their text and relationship ID.
    pub fn hyperlinks(&self) -> Vec<Hyperlink> {
        let mut links = Vec::new();
        for shape in self.shapes() {
            if let Some(text_body) = shape.text_body() {
                for para in text_body.paragraphs() {
                    for run in para.runs() {
                        if let Some(rel_id) = run.hyperlink_rel_id() {
                            links.push(Hyperlink {
                                text: run.text().to_string(),
                                rel_id: rel_id.to_string(),
                            });
                        }
                    }
                }
            }
        }
        links
    }

    /// Check if this slide contains any hyperlinks.
    pub fn has_hyperlinks(&self) -> bool {
        self.shapes().iter().any(|s| {
            s.text_body().is_some_and(|tb| {
                tb.paragraphs()
                    .iter()
                    .any(|p| p.runs().iter().any(|r| r.has_hyperlink()))
            })
        })
    }

    /// Get all tables on the slide.
    pub fn tables(&self) -> &[Table] {
        &self.tables
    }

    /// Check if this slide contains any tables.
    pub fn has_tables(&self) -> bool {
        !self.tables.is_empty()
    }

    /// Get the number of tables on this slide.
    pub fn table_count(&self) -> usize {
        self.tables.len()
    }

    /// Get the underlying generated slide type.
    ///
    /// Use this for full access to all parsed fields.
    pub fn inner(&self) -> &types::Slide {
        &self.inner
    }

    /// Get a table by index (0-based).
    pub fn table(&self, index: usize) -> Option<&Table> {
        self.tables.get(index)
    }
}

/// A table on a slide.
///
/// Represents a table embedded via DrawingML `a:tbl` element inside a `p:graphicFrame`.
/// This is a thin wrapper around [`ooxml_dml::types::CTTable`] that adds the frame name.
///
/// Use the [`ooxml_dml::TableExt`], [`ooxml_dml::TableRowExt`], and [`ooxml_dml::TableCellExt`]
/// traits for convenient access to rows, cells, and text content.
#[derive(Debug, Clone)]
pub struct Table {
    /// Table name (from graphic frame's cNvPr).
    name: Option<String>,
    /// The underlying DrawingML table.
    inner: ooxml_dml::types::CTTable,
}

impl Table {
    /// Get the table name (from the containing graphic frame).
    pub fn name(&self) -> Option<&str> {
        self.name.as_deref()
    }

    /// Get a reference to the underlying DrawingML table.
    pub fn inner(&self) -> &ooxml_dml::types::CTTable {
        &self.inner
    }

    /// Get a mutable reference to the underlying DrawingML table.
    pub fn inner_mut(&mut self) -> &mut ooxml_dml::types::CTTable {
        &mut self.inner
    }

    /// Consume the wrapper and return the underlying DrawingML table.
    pub fn into_inner(self) -> ooxml_dml::types::CTTable {
        self.inner
    }

    /// Get all rows in the table.
    pub fn rows(&self) -> &[ooxml_dml::types::CTTableRow] {
        use ooxml_dml::TableExt;
        self.inner.rows()
    }

    /// Get the number of rows.
    pub fn row_count(&self) -> usize {
        use ooxml_dml::TableExt;
        self.inner.row_count()
    }

    /// Get the number of columns.
    pub fn col_count(&self) -> usize {
        use ooxml_dml::TableExt;
        self.inner.col_count()
    }

    /// Get a cell by row and column index (0-based).
    pub fn cell(&self, row: usize, col: usize) -> Option<&ooxml_dml::types::CTTableCell> {
        use ooxml_dml::TableExt;
        self.inner.cell(row, col)
    }

    /// Get all cell text as a 2D vector.
    pub fn to_text_grid(&self) -> Vec<Vec<String>> {
        use ooxml_dml::TableExt;
        self.inner.to_text_grid()
    }

    /// Get plain text representation (tab-separated values).
    pub fn text(&self) -> String {
        use ooxml_dml::TableExt;
        self.inner.text()
    }
}

/// A hyperlink extracted from a text run.
#[derive(Debug, Clone)]
pub struct Hyperlink {
    /// The text that is hyperlinked.
    pub text: String,
    /// The relationship ID (use with slide relationships to get the URL).
    pub rel_id: String,
}

// ============================================================================
// Parsing
// ============================================================================

/// Parse presentation.xml to get slide relationship IDs in order.
fn parse_presentation_slides(xml: &[u8]) -> Result<Vec<String>> {
    let mut reader = Reader::from_reader(Cursor::new(xml));
    let mut buf = Vec::new();
    let mut slide_ids = Vec::new();
    let mut in_sld_id_lst = false;

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) | Ok(Event::Empty(e)) => {
                let name = e.name();
                let name = name.as_ref();
                if name == b"p:sldIdLst" {
                    in_sld_id_lst = true;
                } else if in_sld_id_lst && name == b"p:sldId" {
                    // Get r:id attribute
                    for attr in e.attributes().filter_map(|a| a.ok()) {
                        if attr.key.as_ref() == b"r:id" {
                            slide_ids.push(String::from_utf8_lossy(&attr.value).into_owned());
                        }
                    }
                }
            }
            Ok(Event::End(e)) => {
                let name = e.name();
                if name.as_ref() == b"p:sldIdLst" {
                    in_sld_id_lst = false;
                }
            }
            Ok(Event::Eof) => break,
            Err(e) => return Err(Error::Xml(e)),
            _ => {}
        }
        buf.clear();
    }

    Ok(slide_ids)
}

/// Parse a notes slide XML file and extract the text content.
fn parse_notes_slide(xml: &[u8]) -> Option<String> {
    use ooxml_dml::parsers::FromXml as DmlFromXml;
    use ooxml_dml::types::TextBody;

    let mut reader = Reader::from_reader(Cursor::new(xml));
    let mut buf = Vec::new();
    let mut all_text = Vec::new();

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) => {
                let local = e.local_name();
                let local = local.as_ref();
                // Notes text is in p:txBody elements - use generated parser
                if local == b"txBody"
                    && let Ok(text_body) = TextBody::from_xml(&mut reader, &e, false)
                {
                    for para in &text_body.p {
                        let text = para.text();
                        if !text.is_empty() {
                            all_text.push(text);
                        }
                    }
                }
            }
            Ok(Event::Eof) => break,
            Err(_) => break,
            _ => {}
        }
        buf.clear();
    }

    if all_text.is_empty() {
        None
    } else {
        Some(all_text.join("\n"))
    }
}

// ============================================================================
// Utilities
// ============================================================================

/// Resolve a relative path against a base path.
fn resolve_path(base: &str, target: &str) -> String {
    if target.starts_with('/') {
        return target.to_string();
    }

    // Get the directory of the base path
    let base_dir = if let Some(idx) = base.rfind('/') {
        &base[..=idx]
    } else {
        "/"
    };

    format!("{}{}", base_dir, target)
}

/// Determine content type from file path extension.
fn content_type_from_path(path: &str) -> String {
    let ext = path.rsplit('.').next().unwrap_or("").to_ascii_lowercase();

    match ext.as_str() {
        "png" => "image/png",
        "jpg" | "jpeg" => "image/jpeg",
        "gif" => "image/gif",
        "bmp" => "image/bmp",
        "tiff" | "tif" => "image/tiff",
        "webp" => "image/webp",
        "svg" => "image/svg+xml",
        "emf" => "image/x-emf",
        "wmf" => "image/x-wmf",
        _ => "application/octet-stream",
    }
    .to_string()
}

/// Parse a slide master XML file.
fn parse_slide_master(xml: &[u8], path: &str) -> SlideMaster {
    let mut reader = Reader::from_reader(Cursor::new(xml));
    let mut buf = Vec::new();
    let mut name = None;
    let mut color_scheme = None;
    let mut background_color = None;
    let mut layout_ids = Vec::new();

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) | Ok(Event::Empty(e)) => {
                let tag = e.name();
                let tag = tag.as_ref();
                match tag {
                    b"p:cSld" => {
                        for attr in e.attributes().filter_map(|a| a.ok()) {
                            if attr.key.as_ref() == b"name" {
                                name = Some(String::from_utf8_lossy(&attr.value).into_owned());
                            }
                        }
                    }
                    b"p:sldLayoutId" => {
                        for attr in e.attributes().filter_map(|a| a.ok()) {
                            if attr.key.as_ref() == b"r:id" {
                                layout_ids.push(String::from_utf8_lossy(&attr.value).into_owned());
                            }
                        }
                    }
                    b"a:clrScheme" => {
                        for attr in e.attributes().filter_map(|a| a.ok()) {
                            if attr.key.as_ref() == b"name" {
                                color_scheme =
                                    Some(String::from_utf8_lossy(&attr.value).into_owned());
                            }
                        }
                    }
                    b"a:srgbClr" => {
                        if background_color.is_none() {
                            for attr in e.attributes().filter_map(|a| a.ok()) {
                                if attr.key.as_ref() == b"val" {
                                    background_color =
                                        Some(String::from_utf8_lossy(&attr.value).into_owned());
                                }
                            }
                        }
                    }
                    _ => {}
                }
            }
            Ok(Event::Eof) => break,
            _ => {}
        }
        buf.clear();
    }

    SlideMaster {
        path: path.to_string(),
        name,
        layout_ids,
        color_scheme,
        background_color,
    }
}

/// Parse a slide layout XML file.
fn parse_slide_layout(xml: &[u8], path: &str, master_rel_id: Option<String>) -> SlideLayout {
    let mut reader = Reader::from_reader(Cursor::new(xml));
    let mut buf = Vec::new();
    let mut name = None;
    let mut layout_type = SlideLayoutType::Unknown;
    let mut match_name = false;
    let mut show_master_shapes = true;

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) | Ok(Event::Empty(e)) => {
                let tag = e.name();
                let tag = tag.as_ref();
                match tag {
                    b"p:sldLayout" => {
                        for attr in e.attributes().filter_map(|a| a.ok()) {
                            let val = String::from_utf8_lossy(&attr.value);
                            match attr.key.as_ref() {
                                b"type" => layout_type = SlideLayoutType::parse(&val),
                                b"matchingName" => match_name = val == "1" || val == "true",
                                b"showMasterSp" => {
                                    show_master_shapes = val != "0" && val != "false"
                                }
                                _ => {}
                            }
                        }
                    }
                    b"p:cSld" => {
                        for attr in e.attributes().filter_map(|a| a.ok()) {
                            if attr.key.as_ref() == b"name" {
                                name = Some(String::from_utf8_lossy(&attr.value).into_owned());
                            }
                        }
                    }
                    _ => {}
                }
            }
            Ok(Event::Eof) => break,
            _ => {}
        }
        buf.clear();
    }

    SlideLayout {
        path: path.to_string(),
        name,
        layout_type,
        master_rel_id,
        match_name,
        show_master_shapes,
    }
}

// ============================================================================
// Generated parser helpers
// ============================================================================

/// Parse a slide using the generated FromXml parser.
fn parse_slide_xml(xml: &[u8]) -> Result<types::Slide> {
    let mut reader = Reader::from_reader(Cursor::new(xml));
    let mut buf = Vec::new();

    // Find the root p:sld element and parse it
    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) => {
                let local_name = e.local_name();
                if local_name.as_ref() == b"sld" {
                    return types::Slide::from_xml(&mut reader, &e, false)
                        .map_err(|e| Error::Invalid(format!("Failed to parse slide: {}", e)));
                }
            }
            Ok(Event::Eof) => break,
            Err(e) => return Err(Error::Xml(e)),
            _ => {}
        }
        buf.clear();
    }

    Err(Error::Invalid("No p:sld element found in slide XML".into()))
}

/// Convert a generated SlideTransition to the handwritten Transition type.
#[cfg(feature = "pml-transitions")]
fn convert_transition(trans: &types::SlideTransition) -> Transition {
    // Get speed
    let speed = trans
        .spd
        .as_ref()
        .map_or(TransitionSpeed::Medium, |s| match s {
            types::STTransitionSpeed::Slow => TransitionSpeed::Slow,
            types::STTransitionSpeed::Med => TransitionSpeed::Medium,
            types::STTransitionSpeed::Fast => TransitionSpeed::Fast,
        });

    // Determine transition type by checking which field is Some
    let transition_type = if trans.fade.is_some() {
        Some(TransitionType::Fade)
    } else if trans.push.is_some() {
        Some(TransitionType::Push)
    } else if trans.wipe.is_some() {
        Some(TransitionType::Wipe)
    } else if trans.split.is_some() {
        Some(TransitionType::Split)
    } else if trans.blinds.is_some() {
        Some(TransitionType::Blinds)
    } else if trans.checker.is_some() {
        Some(TransitionType::Checker)
    } else if trans.circle.is_some() {
        Some(TransitionType::Circle)
    } else if trans.dissolve.is_some() {
        Some(TransitionType::Dissolve)
    } else if trans.comb.is_some() {
        Some(TransitionType::Comb)
    } else if trans.cover.is_some() {
        Some(TransitionType::Cover)
    } else if trans.cut.is_some() {
        Some(TransitionType::Cut)
    } else if trans.diamond.is_some() {
        Some(TransitionType::Diamond)
    } else if trans.plus.is_some() {
        Some(TransitionType::Plus)
    } else if trans.random.is_some() {
        Some(TransitionType::Random)
    } else if trans.strips.is_some() {
        Some(TransitionType::Strips)
    } else if trans.wedge.is_some() {
        Some(TransitionType::Wedge)
    } else if trans.wheel.is_some() {
        Some(TransitionType::Wheel)
    } else if trans.zoom.is_some() {
        Some(TransitionType::Zoom)
    } else {
        None
    };

    Transition {
        transition_type,
        speed,
        advance_on_click: trans.adv_click.unwrap_or(true),
        advance_time_ms: trans.adv_tm,
    }
}

/// Extract tables from graphic frames in a slide.
///
/// Tables are embedded in `p:graphicFrame` elements containing DrawingML `a:tbl`.
/// The graphic frame structure is: p:graphicFrame/a:graphic/a:graphicData/a:tbl
#[cfg(feature = "extra-children")]
fn extract_tables_from_slide(slide: &types::Slide) -> Vec<Table> {
    use crate::ext::GraphicalObjectFrameExt;

    let mut tables = Vec::new();

    // Tables are in graphic frames in the shape tree
    for frame in &slide.common_slide_data.shape_tree.graphic_frame {
        // Get the frame name for the table
        let frame_name = Some(frame.name().to_string()).filter(|s| !s.is_empty());

        // Look through extra_children for the a:graphic element
        for node in &frame.extra_children {
            if let Some(table) = find_table_in_node(&node.node, frame_name.clone()) {
                tables.push(table);
            }
        }
    }

    tables
}

/// Stub for when extra-children feature is disabled.
#[cfg(not(feature = "extra-children"))]
fn extract_tables_from_slide(_slide: &types::Slide) -> Vec<Table> {
    Vec::new()
}

/// Recursively search for a table element in an XML node tree.
#[cfg(feature = "extra-children")]
fn find_table_in_node(node: &ooxml_xml::RawXmlNode, frame_name: Option<String>) -> Option<Table> {
    use ooxml_xml::RawXmlNode;

    match node {
        RawXmlNode::Element(elem) => {
            // Check if this is a table element (a:tbl or just tbl)
            let local_name = elem.name.split(':').next_back().unwrap_or(&elem.name);
            if local_name == "tbl" {
                return parse_table_element(elem, frame_name);
            }

            // Recursively search children
            for child in &elem.children {
                if let Some(table) = find_table_in_node(child, frame_name.clone()) {
                    return Some(table);
                }
            }
            None
        }
        _ => None,
    }
}

/// Parse a table from a RawXmlElement representing a:tbl.
#[cfg(feature = "extra-children")]
fn parse_table_element(
    elem: &ooxml_xml::RawXmlElement,
    frame_name: Option<String>,
) -> Option<Table> {
    use ooxml_dml::types::CTTable;

    // Use the RawXmlElement::parse_as helper to convert to typed struct
    elem.parse_as::<CTTable>()
        .ok()
        .map(|ct_table| wrap_ct_table(ct_table, frame_name))
}

/// Wrap a DML CTTable in our Table type with the frame name.
#[cfg(feature = "extra-children")]
fn wrap_ct_table(ct_table: ooxml_dml::types::CTTable, name: Option<String>) -> Table {
    Table {
        name,
        inner: ct_table,
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_resolve_path() {
        assert_eq!(
            resolve_path("ppt/presentation.xml", "slides/slide1.xml"),
            "ppt/slides/slide1.xml"
        );
        assert_eq!(
            resolve_path("ppt/presentation.xml", "/ppt/slides/slide1.xml"),
            "/ppt/slides/slide1.xml"
        );
    }
}
