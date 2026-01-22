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

// Namespaces
const NS_PRES: &str = "http://schemas.openxmlformats.org/presentationml/2006/main";
const NS_DRAWING: &str = "http://schemas.openxmlformats.org/drawingml/2006/main";
const NS_REL: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const NS_P: &str = "http://schemas.openxmlformats.org/presentationml/2006/main";

/// A text run in a paragraph, optionally with a hyperlink.
#[derive(Debug, Clone)]
pub struct TextRun {
    text: String,
    /// Optional hyperlink URL.
    hyperlink: Option<String>,
}

impl TextRun {
    /// Create a plain text run.
    pub fn text(text: impl Into<String>) -> Self {
        Self {
            text: text.into(),
            hyperlink: None,
        }
    }

    /// Create a hyperlink run.
    pub fn hyperlink(text: impl Into<String>, url: impl Into<String>) -> Self {
        Self {
            text: text.into(),
            hyperlink: Some(url.into()),
        }
    }
}

/// A text element to add to a slide.
#[derive(Debug, Clone)]
struct TextElement {
    /// Text runs (for supporting hyperlinks in part of text).
    runs: Vec<TextRun>,
    is_title: bool,
    x: i64,
    y: i64,
    width: i64,
    height: i64,
}

impl TextElement {
    /// Create a simple text element with a single run.
    fn simple(text: String, is_title: bool, x: i64, y: i64, width: i64, height: i64) -> Self {
        Self {
            runs: vec![TextRun {
                text,
                hyperlink: None,
            }],
            is_title,
            x,
            y,
            width,
            height,
        }
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
        // PNG: 89 50 4E 47
        if data.starts_with(&[0x89, 0x50, 0x4E, 0x47]) {
            return Some(ImageFormat::Png);
        }
        // JPEG: FF D8 FF
        if data.starts_with(&[0xFF, 0xD8, 0xFF]) {
            return Some(ImageFormat::Jpeg);
        }
        // GIF: 47 49 46 38 (GIF8)
        if data.starts_with(b"GIF8") {
            return Some(ImageFormat::Gif);
        }
        None
    }
}

/// An image element to add to a slide.
#[derive(Debug, Clone)]
struct ImageElement {
    /// Image data (raw bytes).
    data: Vec<u8>,
    /// Image format.
    format: ImageFormat,
    /// Position x in EMUs.
    x: i64,
    /// Position y in EMUs.
    y: i64,
    /// Width in EMUs.
    width: i64,
    /// Height in EMUs.
    height: i64,
    /// Description/alt text.
    description: Option<String>,
}

/// A table element to add to a slide.
#[derive(Debug, Clone)]
struct TableElement {
    /// Table name.
    name: Option<String>,
    /// Rows of cells.
    rows: Vec<Vec<String>>,
    /// Column widths in EMUs.
    col_widths: Vec<i64>,
    /// Row heights in EMUs.
    row_heights: Vec<i64>,
    /// Position x in EMUs.
    x: i64,
    /// Position y in EMUs.
    y: i64,
    /// Total width in EMUs.
    width: i64,
    /// Total height in EMUs.
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

/// A slide being built.
#[derive(Debug)]
pub struct SlideBuilder {
    elements: Vec<TextElement>,
    images: Vec<ImageElement>,
    tables: Vec<TableElement>,
    /// Speaker notes for this slide.
    notes: Option<String>,
}

impl SlideBuilder {
    fn new() -> Self {
        Self {
            elements: Vec::new(),
            images: Vec::new(),
            tables: Vec::new(),
            notes: None,
        }
    }

    /// Collect all hyperlinks from this slide's text elements.
    fn hyperlinks(&self) -> Vec<&str> {
        let mut links = Vec::new();
        for element in &self.elements {
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

    /// Check if this slide has hyperlinks.
    fn has_hyperlinks(&self) -> bool {
        self.elements
            .iter()
            .any(|e| e.runs.iter().any(|r| r.hyperlink.is_some()))
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
        self.elements.push(TextElement::simple(
            text.into(),
            true,    // is_title
            457200,  // ~0.5 inch from left
            274638,  // ~0.3 inch from top
            8229600, // ~9 inches wide
            1143000, // ~1.25 inches tall
        ));
        self
    }

    /// Add text content to the slide.
    pub fn add_text(&mut self, text: impl Into<String>) -> &mut Self {
        // Position below title area
        let y_offset = if self.elements.iter().any(|e| e.is_title) {
            1600200 // Below title
        } else {
            274638 // At title position if no title
        };

        self.elements.push(TextElement::simple(
            text.into(),
            false,
            457200,
            y_offset,
            8229600,
            4525963,
        ));
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
        self.elements
            .push(TextElement::simple(text.into(), false, x, y, width, height));
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
        self.elements.push(TextElement {
            runs: vec![TextRun {
                text: text.into(),
                hyperlink: Some(url.into()),
            }],
            is_title: false,
            x,
            y,
            width,
            height,
        });
        self
    }

    /// Add text with mixed content (including hyperlinks) at a specific position.
    ///
    /// Use `TextRunBuilder` to create runs with or without hyperlinks.
    pub fn add_text_with_runs(
        &mut self,
        runs: Vec<TextRun>,
        x: i64,
        y: i64,
        width: i64,
        height: i64,
    ) -> &mut Self {
        self.elements.push(TextElement {
            runs,
            is_title: false,
            x,
            y,
            width,
            height,
        });
        self
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

        // Calculate column widths
        let col_widths = if let Some(widths) = table.col_widths {
            widths
        } else {
            // Distribute evenly
            let col_width = width / num_cols as i64;
            vec![col_width; num_cols]
        };

        // Calculate row heights (distribute evenly)
        let row_height = height / num_rows as i64;
        let row_heights = vec![row_height; num_rows];

        self.tables.push(TableElement {
            name: table.name,
            rows: table.rows,
            col_widths,
            row_heights,
            x,
            y,
            width,
            height,
        });
        self
    }

    /// Check if this slide has tables.
    pub fn has_tables(&self) -> bool {
        !self.tables.is_empty()
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
            // Default slide size: 10" x 7.5" (standard 4:3)
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
        self.slide_width = 12192000; // 13.333 inches
        self.slide_height = 6858000; // 7.5 inches
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

        // Add default content types
        pkg.add_default_content_type("rels", CT_RELATIONSHIPS);
        pkg.add_default_content_type("xml", CT_XML);

        // Add image content types if needed
        let has_images = self.slides.iter().any(|s| s.has_images());
        if has_images {
            pkg.add_default_content_type("jpeg", CT_JPEG);
            pkg.add_default_content_type("jpg", CT_JPEG);
            pkg.add_default_content_type("png", CT_PNG);
            pkg.add_default_content_type("gif", CT_GIF);
        }

        // Build root relationships
        let root_rels = format!(
            r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="{}" Target="ppt/presentation.xml"/>
</Relationships>"#,
            REL_OFFICE_DOCUMENT
        );

        // Build presentation relationships
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
        pres_rels.push_str("</Relationships>");

        // Build presentation.xml
        let presentation_xml = self.serialize_presentation();

        // Write parts to package
        pkg.add_part("_rels/.rels", CT_RELATIONSHIPS, root_rels.as_bytes())?;
        pkg.add_part(
            "ppt/_rels/presentation.xml.rels",
            CT_RELATIONSHIPS,
            pres_rels.as_bytes(),
        )?;
        pkg.add_part(
            "ppt/presentation.xml",
            CT_PRESENTATION,
            presentation_xml.as_bytes(),
        )?;

        // Global image counter for unique image names
        let mut global_image_num = 1;

        // Write each slide, its images, hyperlinks, and notes
        for (i, slide) in self.slides.iter().enumerate() {
            let slide_num = i + 1;

            // Collect hyperlinks for this slide (deduped, in order of first occurrence)
            let hyperlinks = slide.hyperlinks();

            // Build slide relationships if we have images, notes, or hyperlinks
            let needs_rels = slide.has_images() || slide.has_notes() || slide.has_hyperlinks();
            let mut hyperlink_rel_ids: std::collections::HashMap<&str, usize> =
                std::collections::HashMap::new();

            if needs_rels {
                let mut slide_rels = String::new();
                slide_rels.push_str(r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>"#);
                slide_rels.push('\n');
                slide_rels.push_str(
                    r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">"#,
                );
                slide_rels.push('\n');

                let mut rel_id = 1;

                // Add notes relationship if present
                if slide.has_notes() {
                    slide_rels.push_str(&format!(
                        r#"  <Relationship Id="rId{}" Type="{}" Target="../notesSlides/notesSlide{}.xml"/>"#,
                        rel_id, REL_NOTES_SLIDE, slide_num
                    ));
                    slide_rels.push('\n');
                    rel_id += 1;
                }

                // Add image relationships
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

                // Add hyperlink relationships (external, with TargetMode="External")
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

                // Write images to media folder
                for (img_idx, img) in slide.images.iter().enumerate() {
                    let ext = img.format.extension();
                    let img_path = format!("ppt/media/image{}.{}", global_image_num + img_idx, ext);
                    pkg.add_part(&img_path, img.format.content_type(), &img.data)?;
                }
            }

            // Serialize slide with image and hyperlink relationship IDs
            let image_start_rel_id = if slide.has_notes() { 2 } else { 1 };
            let slide_xml =
                self.serialize_slide(slide, slide_num, image_start_rel_id, &hyperlink_rel_ids);
            let part_name = format!("ppt/slides/slide{}.xml", slide_num);
            pkg.add_part(&part_name, CT_SLIDE, slide_xml.as_bytes())?;

            // Update global image counter
            global_image_num += slide.images.len();

            // Write notes slide if present
            if slide.has_notes() {
                let notes_xml = self.serialize_notes_slide(slide, slide_num);
                let notes_name = format!("ppt/notesSlides/notesSlide{}.xml", slide_num);
                pkg.add_part(&notes_name, CT_NOTES_SLIDE, notes_xml.as_bytes())?;
            }
        }

        pkg.finish()?;
        Ok(())
    }

    /// Serialize presentation.xml
    fn serialize_presentation(&self) -> String {
        let mut xml = String::new();
        xml.push_str(r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>"#);
        xml.push('\n');
        xml.push_str(&format!(
            r#"<p:presentation xmlns:a="{}" xmlns:r="{}" xmlns:p="{}">"#,
            NS_DRAWING, NS_REL, NS_P
        ));
        xml.push('\n');

        // Slide size
        xml.push_str(&format!(
            r#"  <p:sldSz cx="{}" cy="{}"/>"#,
            self.slide_width, self.slide_height
        ));
        xml.push('\n');

        // Notes size (same as slide)
        xml.push_str(&format!(
            r#"  <p:notesSz cx="{}" cy="{}"/>"#,
            self.slide_width, self.slide_height
        ));
        xml.push('\n');

        // Slide list
        xml.push_str("  <p:sldIdLst>\n");
        for i in 0..self.slides.len() {
            let slide_id = 256 + i as u32; // IDs start at 256
            let rel_id = i + 1;
            xml.push_str(&format!(
                r#"    <p:sldId id="{}" r:id="rId{}"/>"#,
                slide_id, rel_id
            ));
            xml.push('\n');
        }
        xml.push_str("  </p:sldIdLst>\n");

        xml.push_str("</p:presentation>");
        xml
    }

    /// Serialize a slide to XML.
    fn serialize_slide(
        &self,
        slide: &SlideBuilder,
        _slide_num: usize,
        image_start_rel_id: usize,
        hyperlink_rel_ids: &std::collections::HashMap<&str, usize>,
    ) -> String {
        let mut xml = String::new();
        xml.push_str(r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>"#);
        xml.push('\n');
        xml.push_str(&format!(
            r#"<p:sld xmlns:a="{}" xmlns:r="{}" xmlns:p="{}">"#,
            NS_DRAWING, NS_REL, NS_PRES
        ));
        xml.push('\n');

        xml.push_str("  <p:cSld>\n");
        xml.push_str("    <p:spTree>\n");

        // Non-visual group shape properties
        xml.push_str("      <p:nvGrpSpPr>\n");
        xml.push_str(r#"        <p:cNvPr id="1" name=""/>"#);
        xml.push('\n');
        xml.push_str("        <p:cNvGrpSpPr/>\n");
        xml.push_str("        <p:nvPr/>\n");
        xml.push_str("      </p:nvGrpSpPr>\n");

        // Group shape properties
        xml.push_str("      <p:grpSpPr>\n");
        xml.push_str("        <a:xfrm>\n");
        xml.push_str(r#"          <a:off x="0" y="0"/>"#);
        xml.push('\n');
        xml.push_str(r#"          <a:ext cx="0" cy="0"/>"#);
        xml.push('\n');
        xml.push_str(r#"          <a:chOff x="0" y="0"/>"#);
        xml.push('\n');
        xml.push_str(r#"          <a:chExt cx="0" cy="0"/>"#);
        xml.push('\n');
        xml.push_str("        </a:xfrm>\n");
        xml.push_str("      </p:grpSpPr>\n");

        // Add shapes for each text element
        let mut next_shape_id = 2; // Start at 2 (1 is the group)
        for element in slide.elements.iter() {
            xml.push_str(&self.serialize_text_shape(element, next_shape_id, hyperlink_rel_ids));
            next_shape_id += 1;
        }

        // Add images
        for (i, image) in slide.images.iter().enumerate() {
            let rel_id = image_start_rel_id + i;
            xml.push_str(&self.serialize_picture(image, next_shape_id, rel_id));
            next_shape_id += 1;
        }

        // Add tables
        for table in slide.tables.iter() {
            xml.push_str(&self.serialize_table(table, next_shape_id));
            next_shape_id += 1;
        }

        xml.push_str("    </p:spTree>\n");
        xml.push_str("  </p:cSld>\n");
        xml.push_str("  <p:clrMapOvr>\n");
        xml.push_str("    <a:masterClrMapping/>\n");
        xml.push_str("  </p:clrMapOvr>\n");
        xml.push_str("</p:sld>");
        xml
    }

    /// Serialize a text shape.
    fn serialize_text_shape(
        &self,
        element: &TextElement,
        shape_id: usize,
        hyperlink_rel_ids: &std::collections::HashMap<&str, usize>,
    ) -> String {
        let name = if element.is_title { "Title" } else { "Content" };
        let font_size = if element.is_title { 4400 } else { 2400 }; // In hundredths of a point

        let mut xml = String::new();
        xml.push_str("      <p:sp>\n");

        // Non-visual properties
        xml.push_str("        <p:nvSpPr>\n");
        xml.push_str(&format!(
            r#"          <p:cNvPr id="{}" name="{}"/>"#,
            shape_id, name
        ));
        xml.push('\n');
        xml.push_str(r#"          <p:cNvSpPr txBox="1"/>"#);
        xml.push('\n');
        xml.push_str("          <p:nvPr/>\n");
        xml.push_str("        </p:nvSpPr>\n");

        // Shape properties
        xml.push_str("        <p:spPr>\n");
        xml.push_str("          <a:xfrm>\n");
        xml.push_str(&format!(
            r#"            <a:off x="{}" y="{}"/>"#,
            element.x, element.y
        ));
        xml.push('\n');
        xml.push_str(&format!(
            r#"            <a:ext cx="{}" cy="{}"/>"#,
            element.width, element.height
        ));
        xml.push('\n');
        xml.push_str("          </a:xfrm>\n");
        xml.push_str(r#"          <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>"#);
        xml.push('\n');
        xml.push_str("        </p:spPr>\n");

        // Text body
        xml.push_str("        <p:txBody>\n");
        xml.push_str(r#"          <a:bodyPr/>"#);
        xml.push('\n');
        xml.push_str(r#"          <a:lstStyle/>"#);
        xml.push('\n');
        xml.push_str("          <a:p>\n");

        // Serialize each text run
        for run in &element.runs {
            xml.push_str("            <a:r>\n");

            // Run properties with optional hyperlink
            if let Some(ref url) = run.hyperlink {
                if let Some(&rel_id) = hyperlink_rel_ids.get(url.as_str()) {
                    xml.push_str(&format!(
                        r#"              <a:rPr lang="en-US" sz="{}">"#,
                        font_size
                    ));
                    xml.push('\n');
                    xml.push_str(&format!(
                        r#"                <a:hlinkClick r:id="rId{}"/>"#,
                        rel_id
                    ));
                    xml.push('\n');
                    xml.push_str("              </a:rPr>\n");
                } else {
                    xml.push_str(&format!(
                        r#"              <a:rPr lang="en-US" sz="{}"/>"#,
                        font_size
                    ));
                    xml.push('\n');
                }
            } else {
                xml.push_str(&format!(
                    r#"              <a:rPr lang="en-US" sz="{}"/>"#,
                    font_size
                ));
                xml.push('\n');
            }

            xml.push_str(&format!(
                "              <a:t>{}</a:t>\n",
                escape_xml(&run.text)
            ));
            xml.push_str("            </a:r>\n");
        }

        xml.push_str("          </a:p>\n");
        xml.push_str("        </p:txBody>\n");

        xml.push_str("      </p:sp>\n");
        xml
    }

    /// Serialize a picture (image) shape.
    fn serialize_picture(&self, image: &ImageElement, shape_id: usize, rel_id: usize) -> String {
        let name = format!("Picture {}", shape_id);
        let desc = image.description.as_deref().unwrap_or("");

        let mut xml = String::new();
        xml.push_str("      <p:pic>\n");

        // Non-visual properties
        xml.push_str("        <p:nvPicPr>\n");
        xml.push_str(&format!(
            r#"          <p:cNvPr id="{}" name="{}" descr="{}"/>"#,
            shape_id,
            escape_xml(&name),
            escape_xml(desc)
        ));
        xml.push('\n');
        xml.push_str("          <p:cNvPicPr>\n");
        xml.push_str("            <a:picLocks noChangeAspect=\"1\"/>\n");
        xml.push_str("          </p:cNvPicPr>\n");
        xml.push_str("          <p:nvPr/>\n");
        xml.push_str("        </p:nvPicPr>\n");

        // Blip fill (reference to image)
        xml.push_str("        <p:blipFill>\n");
        xml.push_str(&format!(r#"          <a:blip r:embed="rId{}"/>"#, rel_id));
        xml.push('\n');
        xml.push_str("          <a:stretch>\n");
        xml.push_str("            <a:fillRect/>\n");
        xml.push_str("          </a:stretch>\n");
        xml.push_str("        </p:blipFill>\n");

        // Shape properties
        xml.push_str("        <p:spPr>\n");
        xml.push_str("          <a:xfrm>\n");
        xml.push_str(&format!(
            r#"            <a:off x="{}" y="{}"/>"#,
            image.x, image.y
        ));
        xml.push('\n');
        xml.push_str(&format!(
            r#"            <a:ext cx="{}" cy="{}"/>"#,
            image.width, image.height
        ));
        xml.push('\n');
        xml.push_str("          </a:xfrm>\n");
        xml.push_str(r#"          <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>"#);
        xml.push('\n');
        xml.push_str("        </p:spPr>\n");

        xml.push_str("      </p:pic>\n");
        xml
    }

    /// Serialize a table as a graphic frame.
    fn serialize_table(&self, table: &TableElement, shape_id: usize) -> String {
        let name = table.name.as_deref().unwrap_or("Table");

        let mut xml = String::new();
        xml.push_str("      <p:graphicFrame>\n");

        // Non-visual properties
        xml.push_str("        <p:nvGraphicFramePr>\n");
        xml.push_str(&format!(
            r#"          <p:cNvPr id="{}" name="{}"/>"#,
            shape_id,
            escape_xml(name)
        ));
        xml.push('\n');
        xml.push_str("          <p:cNvGraphicFramePr>\n");
        xml.push_str("            <a:graphicFrameLocks noGrp=\"1\"/>\n");
        xml.push_str("          </p:cNvGraphicFramePr>\n");
        xml.push_str("          <p:nvPr/>\n");
        xml.push_str("        </p:nvGraphicFramePr>\n");

        // Transform
        xml.push_str("        <p:xfrm>\n");
        xml.push_str(&format!(
            r#"          <a:off x="{}" y="{}"/>"#,
            table.x, table.y
        ));
        xml.push('\n');
        xml.push_str(&format!(
            r#"          <a:ext cx="{}" cy="{}"/>"#,
            table.width, table.height
        ));
        xml.push('\n');
        xml.push_str("        </p:xfrm>\n");

        // Graphic element containing the table
        xml.push_str("        <a:graphic>\n");
        xml.push_str(
            r#"          <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table">"#,
        );
        xml.push('\n');

        // The table itself
        xml.push_str("            <a:tbl>\n");

        // Table properties
        xml.push_str("              <a:tblPr firstRow=\"1\" bandRow=\"1\"/>\n");

        // Table grid (column definitions)
        xml.push_str("              <a:tblGrid>\n");
        for col_width in &table.col_widths {
            xml.push_str(&format!(
                r#"                <a:gridCol w="{}"/>"#,
                col_width
            ));
            xml.push('\n');
        }
        xml.push_str("              </a:tblGrid>\n");

        // Table rows
        for (row_idx, row) in table.rows.iter().enumerate() {
            let row_height = table.row_heights.get(row_idx).copied().unwrap_or(370840);
            xml.push_str(&format!(r#"              <a:tr h="{}">"#, row_height));
            xml.push('\n');

            for cell_text in row {
                xml.push_str("                <a:tc>\n");
                xml.push_str("                  <a:txBody>\n");
                xml.push_str("                    <a:bodyPr/>\n");
                xml.push_str("                    <a:lstStyle/>\n");
                xml.push_str("                    <a:p>\n");
                xml.push_str("                      <a:r>\n");
                xml.push_str(r#"                        <a:rPr lang="en-US"/>"#);
                xml.push('\n');
                xml.push_str(&format!(
                    "                        <a:t>{}</a:t>\n",
                    escape_xml(cell_text)
                ));
                xml.push_str("                      </a:r>\n");
                xml.push_str("                    </a:p>\n");
                xml.push_str("                  </a:txBody>\n");
                xml.push_str("                  <a:tcPr/>\n");
                xml.push_str("                </a:tc>\n");
            }

            xml.push_str("              </a:tr>\n");
        }

        xml.push_str("            </a:tbl>\n");
        xml.push_str("          </a:graphicData>\n");
        xml.push_str("        </a:graphic>\n");
        xml.push_str("      </p:graphicFrame>\n");
        xml
    }

    /// Serialize a notes slide to XML.
    fn serialize_notes_slide(&self, slide: &SlideBuilder, slide_num: usize) -> String {
        let notes_text = slide.notes.as_deref().unwrap_or("");

        let mut xml = String::new();
        xml.push_str(r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>"#);
        xml.push('\n');
        xml.push_str(&format!(
            r#"<p:notes xmlns:a="{}" xmlns:r="{}" xmlns:p="{}">"#,
            NS_DRAWING, NS_REL, NS_PRES
        ));
        xml.push('\n');

        xml.push_str("  <p:cSld>\n");
        xml.push_str("    <p:spTree>\n");

        // Non-visual group shape properties
        xml.push_str("      <p:nvGrpSpPr>\n");
        xml.push_str(r#"        <p:cNvPr id="1" name=""/>"#);
        xml.push('\n');
        xml.push_str("        <p:cNvGrpSpPr/>\n");
        xml.push_str("        <p:nvPr/>\n");
        xml.push_str("      </p:nvGrpSpPr>\n");

        // Group shape properties
        xml.push_str("      <p:grpSpPr>\n");
        xml.push_str("        <a:xfrm>\n");
        xml.push_str(r#"          <a:off x="0" y="0"/>"#);
        xml.push('\n');
        xml.push_str(r#"          <a:ext cx="0" cy="0"/>"#);
        xml.push('\n');
        xml.push_str(r#"          <a:chOff x="0" y="0"/>"#);
        xml.push('\n');
        xml.push_str(r#"          <a:chExt cx="0" cy="0"/>"#);
        xml.push('\n');
        xml.push_str("        </a:xfrm>\n");
        xml.push_str("      </p:grpSpPr>\n");

        // Slide image placeholder (shape 2)
        xml.push_str("      <p:sp>\n");
        xml.push_str("        <p:nvSpPr>\n");
        xml.push_str(&format!(
            r#"          <p:cNvPr id="2" name="Slide Image Placeholder {}"/>"#,
            slide_num
        ));
        xml.push('\n');
        xml.push_str("          <p:cNvSpPr>\n");
        xml.push_str(r#"            <a:spLocks noGrp="1" noRot="1" noChangeAspect="1"/>"#);
        xml.push('\n');
        xml.push_str("          </p:cNvSpPr>\n");
        xml.push_str(r#"          <p:nvPr><p:ph type="sldImg"/></p:nvPr>"#);
        xml.push('\n');
        xml.push_str("        </p:nvSpPr>\n");
        xml.push_str("        <p:spPr/>\n");
        xml.push_str("      </p:sp>\n");

        // Notes body placeholder (shape 3)
        xml.push_str("      <p:sp>\n");
        xml.push_str("        <p:nvSpPr>\n");
        xml.push_str(r#"          <p:cNvPr id="3" name="Notes Placeholder"/>"#);
        xml.push('\n');
        xml.push_str("          <p:cNvSpPr>\n");
        xml.push_str(r#"            <a:spLocks noGrp="1"/>"#);
        xml.push('\n');
        xml.push_str("          </p:cNvSpPr>\n");
        xml.push_str(r#"          <p:nvPr><p:ph type="body" idx="1"/></p:nvPr>"#);
        xml.push('\n');
        xml.push_str("        </p:nvSpPr>\n");
        xml.push_str("        <p:spPr/>\n");
        xml.push_str("        <p:txBody>\n");
        xml.push_str(r#"          <a:bodyPr/>"#);
        xml.push('\n');
        xml.push_str(r#"          <a:lstStyle/>"#);
        xml.push('\n');

        // Write notes text, each line as a paragraph
        for line in notes_text.lines() {
            xml.push_str("          <a:p>\n");
            xml.push_str("            <a:r>\n");
            xml.push_str(r#"              <a:rPr lang="en-US"/>"#);
            xml.push('\n');
            xml.push_str(&format!("              <a:t>{}</a:t>\n", escape_xml(line)));
            xml.push_str("            </a:r>\n");
            xml.push_str("          </a:p>\n");
        }

        // If no text, add empty paragraph
        if notes_text.is_empty() {
            xml.push_str("          <a:p/>\n");
        }

        xml.push_str("        </p:txBody>\n");
        xml.push_str("      </p:sp>\n");

        xml.push_str("    </p:spTree>\n");
        xml.push_str("  </p:cSld>\n");
        xml.push_str("</p:notes>");
        xml
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

        // Write to memory
        let mut buffer = Cursor::new(Vec::new());
        pres.write(&mut buffer).unwrap();

        // Read back - just verify structure, not content
        // (content verification needs XML namespace fixes)
        buffer.set_position(0);
        let mut presentation = crate::Presentation::from_reader(buffer).unwrap();
        assert_eq!(presentation.slide_count(), 1);

        // Verify slide can be loaded (even if shapes not parsed yet)
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

    #[test]
    fn test_roundtrip_with_table() {
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

        // Write to memory
        let mut buffer = Cursor::new(Vec::new());
        pres.write(&mut buffer).unwrap();

        // Read back
        buffer.set_position(0);
        let mut presentation = crate::Presentation::from_reader(buffer).unwrap();
        assert_eq!(presentation.slide_count(), 1);

        // Verify slide has the table
        let read_slide = presentation.slide(0).unwrap();
        assert!(read_slide.has_tables());
        assert_eq!(read_slide.table_count(), 1);

        let table = read_slide.table(0).unwrap();
        assert_eq!(table.row_count(), 3);
        assert_eq!(table.col_count(), 2);

        // Verify cell content
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

        // Write to memory
        let mut buffer = Cursor::new(Vec::new());
        pres.write(&mut buffer).unwrap();

        // Read back
        buffer.set_position(0);
        let mut presentation = crate::Presentation::from_reader(buffer).unwrap();
        assert_eq!(presentation.slide_count(), 1);

        // Verify slide has the hyperlink
        let read_slide = presentation.slide(0).unwrap();
        assert!(read_slide.has_hyperlinks());

        // Get the hyperlinks
        let links = read_slide.hyperlinks();
        assert_eq!(links.len(), 1);
        assert_eq!(links[0].text, "Click here to visit Rust");

        // Resolve the hyperlink to get URL
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

        // Write to memory
        let mut buffer = Cursor::new(Vec::new());
        pres.write(&mut buffer).unwrap();

        // Read back
        buffer.set_position(0);
        let mut presentation = crate::Presentation::from_reader(buffer).unwrap();
        let read_slide = presentation.slide(0).unwrap();

        // Should have one hyperlink
        assert!(read_slide.has_hyperlinks());
        let links = read_slide.hyperlinks();
        assert_eq!(links.len(), 1);
        assert_eq!(links[0].text, "documentation");
    }
}
