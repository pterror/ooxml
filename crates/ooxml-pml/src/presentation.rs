//! Presentation API for reading and writing PowerPoint files.
//!
//! This module provides the main entry point for working with PPTX files.

// TODO: Migrate from ooxml_dml::text types to generated types with ext traits

use crate::error::{Error, Result};
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
        let mut slide = parse_slide(&data, info.index, &info.path)?;

        // Set layout relationship ID
        slide.layout_rel_id = info.layout_rel_id.clone();

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
    pub fn get_image_data(&mut self, slide: &Slide, picture: &Picture) -> Result<ImageData> {
        // Get slide relationships
        let slide_rels = self
            .package
            .read_part_relationships(slide.slide_path())
            .map_err(|_| Error::Invalid("Failed to read slide relationships".into()))?;

        // Find the image relationship
        let rel = slide_rels.get(picture.rel_id()).ok_or_else(|| {
            Error::Invalid(format!("Image relationship {} not found", picture.rel_id()))
        })?;

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
#[derive(Debug, Clone)]
pub struct Slide {
    index: usize,
    shapes: Vec<Shape>,
    /// Pictures on this slide.
    pictures: Vec<Picture>,
    /// Tables on this slide.
    tables: Vec<Table>,
    /// Path to this slide (for resolving image paths).
    slide_path: String,
    /// Speaker notes for this slide.
    notes: Option<String>,
    /// Slide transition effect.
    transition: Option<Transition>,
    /// Relationship ID to the slide layout (for linking to layout info).
    layout_rel_id: Option<String>,
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
    pub fn shapes(&self) -> &[Shape] {
        &self.shapes
    }

    /// Extract all text from the slide.
    pub fn text(&self) -> String {
        self.shapes
            .iter()
            .filter_map(|s| s.text())
            .collect::<Vec<_>>()
            .join("\n")
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
    pub fn pictures(&self) -> &[Picture] {
        &self.pictures
    }

    /// Get the path to this slide part (for resolving image relationships).
    pub(crate) fn slide_path(&self) -> &str {
        &self.slide_path
    }

    /// Get the slide transition effect (if any).
    pub fn transition(&self) -> Option<&Transition> {
        self.transition.as_ref()
    }

    /// Check if this slide has a transition effect.
    pub fn has_transition(&self) -> bool {
        self.transition.is_some()
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
        self.shapes.iter().flat_map(|s| s.hyperlinks()).collect()
    }

    /// Check if this slide contains any hyperlinks.
    pub fn has_hyperlinks(&self) -> bool {
        self.shapes.iter().any(|s| s.has_hyperlinks())
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

    /// Get a table by index (0-based).
    pub fn table(&self, index: usize) -> Option<&Table> {
        self.tables.get(index)
    }
}

/// A shape on a slide.
#[derive(Debug, Clone)]
pub struct Shape {
    /// Shape name (if any).
    name: Option<String>,
    /// Text paragraphs (DrawingML).
    paragraphs: Vec<ooxml_dml::text::Paragraph>,
}

impl Shape {
    /// Get the shape name.
    pub fn name(&self) -> Option<&str> {
        self.name.as_deref()
    }

    /// Get the text paragraphs.
    pub fn paragraphs(&self) -> &[ooxml_dml::text::Paragraph] {
        &self.paragraphs
    }

    /// Get the text content (paragraphs joined).
    pub fn text(&self) -> Option<String> {
        if self.paragraphs.is_empty() {
            None
        } else {
            Some(
                self.paragraphs
                    .iter()
                    .map(|p| p.text())
                    .collect::<Vec<_>>()
                    .join("\n"),
            )
        }
    }

    /// Check if the shape has text content.
    pub fn has_text(&self) -> bool {
        !self.paragraphs.is_empty()
    }

    /// Get all hyperlinks in this shape.
    ///
    /// Returns hyperlinks with their text and relationship ID.
    /// Use the slide relationships to resolve the rel_id to a URL.
    pub fn hyperlinks(&self) -> Vec<Hyperlink> {
        let mut links = Vec::new();
        for para in &self.paragraphs {
            for run in para.runs() {
                if let Some(rel_id) = run.hyperlink_rel_id() {
                    links.push(Hyperlink {
                        text: run.text().to_string(),
                        rel_id: rel_id.to_string(),
                    });
                }
            }
        }
        links
    }

    /// Check if the shape contains any hyperlinks.
    pub fn has_hyperlinks(&self) -> bool {
        self.paragraphs
            .iter()
            .any(|p| p.runs().iter().any(|r| r.has_hyperlink()))
    }
}

/// A picture element on a slide.
///
/// Represents an image embedded via `p:pic` element.
#[derive(Debug, Clone)]
pub struct Picture {
    /// Relationship ID referencing the image file.
    rel_id: String,
    /// Picture name (from cNvPr).
    name: Option<String>,
    /// Description/alt text.
    description: Option<String>,
}

impl Picture {
    /// Get the relationship ID for this picture's image.
    pub fn rel_id(&self) -> &str {
        &self.rel_id
    }

    /// Get the picture name.
    pub fn name(&self) -> Option<&str> {
        self.name.as_deref()
    }

    /// Get the description/alt text.
    pub fn description(&self) -> Option<&str> {
        self.description.as_deref()
    }
}

/// A table on a slide.
///
/// Represents a table embedded via DrawingML `a:tbl` element inside a `p:graphicFrame`.
#[derive(Debug, Clone)]
pub struct Table {
    /// Table name (from cNvPr).
    name: Option<String>,
    /// Table rows.
    rows: Vec<TableRow>,
}

impl Table {
    /// Get the table name.
    pub fn name(&self) -> Option<&str> {
        self.name.as_deref()
    }

    /// Get all rows in the table.
    pub fn rows(&self) -> &[TableRow] {
        &self.rows
    }

    /// Get the number of rows.
    pub fn row_count(&self) -> usize {
        self.rows.len()
    }

    /// Get the number of columns (from first row, or 0 if empty).
    pub fn col_count(&self) -> usize {
        self.rows.first().map(|r| r.cells.len()).unwrap_or(0)
    }

    /// Get a cell by row and column index (0-based).
    pub fn cell(&self, row: usize, col: usize) -> Option<&TableCell> {
        self.rows.get(row).and_then(|r| r.cells.get(col))
    }

    /// Get all cell text as a 2D vector.
    pub fn to_text_grid(&self) -> Vec<Vec<String>> {
        self.rows
            .iter()
            .map(|row| row.cells.iter().map(|c| c.text()).collect())
            .collect()
    }

    /// Get plain text representation (tab-separated values).
    pub fn text(&self) -> String {
        self.rows
            .iter()
            .map(|row| {
                row.cells
                    .iter()
                    .map(|c| c.text())
                    .collect::<Vec<_>>()
                    .join("\t")
            })
            .collect::<Vec<_>>()
            .join("\n")
    }
}

/// A row in a table.
#[derive(Debug, Clone)]
pub struct TableRow {
    /// Cells in this row.
    cells: Vec<TableCell>,
    /// Row height in EMUs (if specified).
    height: Option<i64>,
}

impl TableRow {
    /// Get all cells in this row.
    pub fn cells(&self) -> &[TableCell] {
        &self.cells
    }

    /// Get a cell by column index (0-based).
    pub fn cell(&self, col: usize) -> Option<&TableCell> {
        self.cells.get(col)
    }

    /// Get the row height in EMUs (if specified).
    pub fn height(&self) -> Option<i64> {
        self.height
    }
}

/// A cell in a table.
#[derive(Debug, Clone)]
pub struct TableCell {
    /// Text paragraphs in the cell.
    paragraphs: Vec<ooxml_dml::text::Paragraph>,
    /// Row span (number of rows this cell spans).
    row_span: u32,
    /// Column span (number of columns this cell spans).
    col_span: u32,
}

impl TableCell {
    /// Get the text paragraphs.
    pub fn paragraphs(&self) -> &[ooxml_dml::text::Paragraph] {
        &self.paragraphs
    }

    /// Get the cell text (paragraphs joined with newlines).
    pub fn text(&self) -> String {
        self.paragraphs
            .iter()
            .map(|p| p.text())
            .collect::<Vec<_>>()
            .join("\n")
    }

    /// Get the row span.
    pub fn row_span(&self) -> u32 {
        self.row_span
    }

    /// Get the column span.
    pub fn col_span(&self) -> u32 {
        self.col_span
    }

    /// Check if this cell spans multiple rows.
    pub fn has_row_span(&self) -> bool {
        self.row_span > 1
    }

    /// Check if this cell spans multiple columns.
    pub fn has_col_span(&self) -> bool {
        self.col_span > 1
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

/// Parse a slide XML file.
fn parse_slide(xml: &[u8], index: usize, slide_path: &str) -> Result<Slide> {
    let mut reader = Reader::from_reader(Cursor::new(xml));
    let mut buf = Vec::new();
    let mut shapes = Vec::new();
    let mut pictures = Vec::new();
    let mut tables = Vec::new();

    let mut current_shape_name: Option<String> = None;
    let mut current_paragraphs: Vec<ooxml_dml::text::Paragraph> = Vec::new();
    let mut in_sp = false; // Inside a shape

    // Picture parsing state
    let mut in_pic = false;
    let mut current_pic_name: Option<String> = None;
    let mut current_pic_descr: Option<String> = None;
    let mut current_pic_rel_id: Option<String> = None;

    // Table parsing state
    let mut in_graphic_frame = false;
    let mut current_table_name: Option<String> = None;
    let mut in_tbl = false;
    let mut current_table_rows: Vec<TableRow> = Vec::new();
    let mut in_tr = false;
    let mut current_row_cells: Vec<TableCell> = Vec::new();
    let mut current_row_height: Option<i64> = None;
    let mut in_tc = false;
    let mut current_cell_paragraphs: Vec<ooxml_dml::text::Paragraph> = Vec::new();
    let mut current_cell_row_span: u32 = 1;
    let mut current_cell_col_span: u32 = 1;

    // Transition state
    let mut transition: Option<Transition> = None;
    let mut in_transition = false;

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) => {
                let name = e.name();
                let name = name.as_ref();
                match name {
                    b"p:sp" => {
                        in_sp = true;
                        current_shape_name = None;
                        current_paragraphs.clear();
                    }
                    b"p:pic" => {
                        in_pic = true;
                        current_pic_name = None;
                        current_pic_descr = None;
                        current_pic_rel_id = None;
                    }
                    b"p:graphicFrame" => {
                        in_graphic_frame = true;
                        current_table_name = None;
                    }
                    b"a:tbl" if in_graphic_frame => {
                        in_tbl = true;
                        current_table_rows.clear();
                    }
                    b"a:tr" if in_tbl => {
                        in_tr = true;
                        current_row_cells.clear();
                        current_row_height = None;
                        for attr in e.attributes().filter_map(|a| a.ok()) {
                            if attr.key.as_ref() == b"h" {
                                current_row_height =
                                    String::from_utf8_lossy(&attr.value).parse().ok();
                            }
                        }
                    }
                    b"a:tc" if in_tr => {
                        in_tc = true;
                        current_cell_paragraphs.clear();
                        current_cell_row_span = 1;
                        current_cell_col_span = 1;
                        for attr in e.attributes().filter_map(|a| a.ok()) {
                            match attr.key.as_ref() {
                                b"rowSpan" => {
                                    current_cell_row_span =
                                        String::from_utf8_lossy(&attr.value).parse().unwrap_or(1);
                                }
                                b"gridSpan" => {
                                    current_cell_col_span =
                                        String::from_utf8_lossy(&attr.value).parse().unwrap_or(1);
                                }
                                _ => {}
                            }
                        }
                    }
                    b"a:txBody" if in_tc => {
                        // Parse text body for table cell
                        if let Ok(paras) = ooxml_dml::text::parse_text_body_from_reader(&mut reader)
                        {
                            current_cell_paragraphs = paras;
                        }
                    }
                    b"p:transition" => {
                        in_transition = true;
                        transition = Some(parse_transition_attrs(&e));
                    }
                    b"p:cNvPr" => {
                        // Non-visual properties - get name and description
                        if in_sp {
                            for attr in e.attributes().filter_map(|a| a.ok()) {
                                if attr.key.as_ref() == b"name" {
                                    current_shape_name =
                                        Some(String::from_utf8_lossy(&attr.value).into_owned());
                                }
                            }
                        } else if in_pic {
                            for attr in e.attributes().filter_map(|a| a.ok()) {
                                match attr.key.as_ref() {
                                    b"name" => {
                                        current_pic_name =
                                            Some(String::from_utf8_lossy(&attr.value).into_owned());
                                    }
                                    b"descr" => {
                                        current_pic_descr =
                                            Some(String::from_utf8_lossy(&attr.value).into_owned());
                                    }
                                    _ => {}
                                }
                            }
                        } else if in_graphic_frame && !in_tbl {
                            // Get table name from graphicFrame's cNvPr before we enter the table
                            for attr in e.attributes().filter_map(|a| a.ok()) {
                                if attr.key.as_ref() == b"name" {
                                    current_table_name =
                                        Some(String::from_utf8_lossy(&attr.value).into_owned());
                                }
                            }
                        }
                    }
                    b"p:txBody" => {
                        // Use DML parser for text body content
                        if in_sp
                            && let Ok(paras) =
                                ooxml_dml::text::parse_text_body_from_reader(&mut reader)
                        {
                            current_paragraphs = paras;
                        }
                    }
                    _ => {
                        // Check for transition type elements inside p:transition
                        if in_transition
                            && let Some(ref mut trans) = transition
                            && let Some(tt) = parse_transition_type_element(name)
                        {
                            trans.transition_type = Some(tt);
                        }
                    }
                }
            }
            Ok(Event::Empty(e)) => {
                let name = e.name();
                let name = name.as_ref();
                match name {
                    b"p:cNvPr" => {
                        // Handle self-closing cNvPr
                        if in_sp {
                            for attr in e.attributes().filter_map(|a| a.ok()) {
                                if attr.key.as_ref() == b"name" {
                                    current_shape_name =
                                        Some(String::from_utf8_lossy(&attr.value).into_owned());
                                }
                            }
                        } else if in_pic {
                            for attr in e.attributes().filter_map(|a| a.ok()) {
                                match attr.key.as_ref() {
                                    b"name" => {
                                        current_pic_name =
                                            Some(String::from_utf8_lossy(&attr.value).into_owned());
                                    }
                                    b"descr" => {
                                        current_pic_descr =
                                            Some(String::from_utf8_lossy(&attr.value).into_owned());
                                    }
                                    _ => {}
                                }
                            }
                        } else if in_graphic_frame && !in_tbl {
                            // Get table name from graphicFrame's cNvPr
                            for attr in e.attributes().filter_map(|a| a.ok()) {
                                if attr.key.as_ref() == b"name" {
                                    current_table_name =
                                        Some(String::from_utf8_lossy(&attr.value).into_owned());
                                }
                            }
                        }
                    }
                    b"a:blip" => {
                        // Image reference - get r:embed attribute
                        if in_pic {
                            for attr in e.attributes().filter_map(|a| a.ok()) {
                                if attr.key.as_ref() == b"r:embed" {
                                    current_pic_rel_id =
                                        Some(String::from_utf8_lossy(&attr.value).into_owned());
                                }
                            }
                        }
                    }
                    b"p:transition" => {
                        // Self-closing transition element
                        transition = Some(parse_transition_attrs(&e));
                    }
                    _ => {
                        // Check for self-closing transition type elements
                        if in_transition
                            && let Some(ref mut trans) = transition
                            && let Some(tt) = parse_transition_type_element(name)
                        {
                            trans.transition_type = Some(tt);
                        }
                    }
                }
            }
            Ok(Event::End(e)) => {
                let name = e.name();
                let name = name.as_ref();
                match name {
                    b"p:sp" => {
                        // End of shape
                        shapes.push(Shape {
                            name: current_shape_name.take(),
                            paragraphs: std::mem::take(&mut current_paragraphs),
                        });
                        in_sp = false;
                    }
                    b"p:pic" => {
                        // End of picture
                        if let Some(rel_id) = current_pic_rel_id.take() {
                            pictures.push(Picture {
                                rel_id,
                                name: current_pic_name.take(),
                                description: current_pic_descr.take(),
                            });
                        }
                        in_pic = false;
                    }
                    b"a:tc" if in_tc => {
                        // End of table cell
                        current_row_cells.push(TableCell {
                            paragraphs: std::mem::take(&mut current_cell_paragraphs),
                            row_span: current_cell_row_span,
                            col_span: current_cell_col_span,
                        });
                        in_tc = false;
                    }
                    b"a:tr" if in_tr => {
                        // End of table row
                        current_table_rows.push(TableRow {
                            cells: std::mem::take(&mut current_row_cells),
                            height: current_row_height.take(),
                        });
                        in_tr = false;
                    }
                    b"a:tbl" if in_tbl => {
                        // End of table
                        tables.push(Table {
                            name: current_table_name.take(),
                            rows: std::mem::take(&mut current_table_rows),
                        });
                        in_tbl = false;
                    }
                    b"p:graphicFrame" => {
                        in_graphic_frame = false;
                    }
                    b"p:transition" => {
                        in_transition = false;
                    }
                    _ => {}
                }
            }
            Ok(Event::Eof) => break,
            Err(e) => return Err(Error::Xml(e)),
            _ => {}
        }
        buf.clear();
    }

    Ok(Slide {
        index,
        shapes,
        pictures,
        tables,
        slide_path: slide_path.to_string(),
        notes: None,
        transition,
        layout_rel_id: None,
    })
}

/// Parse transition attributes from a p:transition element.
fn parse_transition_attrs(e: &quick_xml::events::BytesStart) -> Transition {
    let mut trans = Transition {
        advance_on_click: true, // Default is true
        ..Default::default()
    };

    for attr in e.attributes().filter_map(|a| a.ok()) {
        let value = String::from_utf8_lossy(&attr.value);
        match attr.key.as_ref() {
            b"spd" => {
                trans.speed = match value.as_ref() {
                    "slow" => TransitionSpeed::Slow,
                    "fast" => TransitionSpeed::Fast,
                    _ => TransitionSpeed::Medium,
                };
            }
            b"advClick" => {
                trans.advance_on_click = value != "0" && value != "false";
            }
            b"advTm" => {
                trans.advance_time_ms = value.parse().ok();
            }
            _ => {}
        }
    }

    trans
}

/// Parse a transition type from an element name.
fn parse_transition_type_element(name: &[u8]) -> Option<TransitionType> {
    // Remove namespace prefix if present
    let local = if let Some(pos) = name.iter().position(|&b| b == b':') {
        &name[pos + 1..]
    } else {
        name
    };

    match local {
        b"fade" => Some(TransitionType::Fade),
        b"push" => Some(TransitionType::Push),
        b"wipe" => Some(TransitionType::Wipe),
        b"split" => Some(TransitionType::Split),
        b"blinds" => Some(TransitionType::Blinds),
        b"checker" => Some(TransitionType::Checker),
        b"circle" => Some(TransitionType::Circle),
        b"dissolve" => Some(TransitionType::Dissolve),
        b"comb" => Some(TransitionType::Comb),
        b"cover" => Some(TransitionType::Cover),
        b"cut" => Some(TransitionType::Cut),
        b"diamond" => Some(TransitionType::Diamond),
        b"plus" => Some(TransitionType::Plus),
        b"random" => Some(TransitionType::Random),
        b"strips" => Some(TransitionType::Strips),
        b"wedge" => Some(TransitionType::Wedge),
        b"wheel" => Some(TransitionType::Wheel),
        b"zoom" => Some(TransitionType::Zoom),
        _ => {
            let name_str = String::from_utf8_lossy(local);
            // Only return Other for elements that look like transitions
            if name_str.chars().all(|c| c.is_ascii_alphabetic()) && name_str.len() > 2 {
                Some(TransitionType::Other(name_str.into_owned()))
            } else {
                None
            }
        }
    }
}

/// Parse a notes slide XML file and extract the text content.
fn parse_notes_slide(xml: &[u8]) -> Option<String> {
    let mut reader = Reader::from_reader(Cursor::new(xml));
    let mut buf = Vec::new();
    let mut all_text = Vec::new();

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) => {
                let name = e.name();
                let name = name.as_ref();
                // Notes text is in p:txBody elements
                if name == b"p:txBody"
                    && let Ok(paras) = ooxml_dml::text::parse_text_body_from_reader(&mut reader)
                {
                    for para in paras {
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
