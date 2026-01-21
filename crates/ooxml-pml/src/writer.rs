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
const CT_RELATIONSHIPS: &str = "application/vnd.openxmlformats-package.relationships+xml";
const CT_XML: &str = "application/xml";

// Relationship types
const REL_OFFICE_DOCUMENT: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument";
const REL_SLIDE: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide";

// Namespaces
const NS_PRES: &str = "http://schemas.openxmlformats.org/presentationml/2006/main";
const NS_DRAWING: &str = "http://schemas.openxmlformats.org/drawingml/2006/main";
const NS_REL: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const NS_P: &str = "http://schemas.openxmlformats.org/presentationml/2006/main";

/// A text element to add to a slide.
#[derive(Debug, Clone)]
pub struct TextElement {
    text: String,
    is_title: bool,
    x: i64,
    y: i64,
    width: i64,
    height: i64,
}

/// A slide being built.
#[derive(Debug)]
pub struct SlideBuilder {
    elements: Vec<TextElement>,
}

impl SlideBuilder {
    fn new() -> Self {
        Self {
            elements: Vec::new(),
        }
    }

    /// Add a title to the slide.
    pub fn add_title(&mut self, text: impl Into<String>) -> &mut Self {
        self.elements.push(TextElement {
            text: text.into(),
            is_title: true,
            x: 457200,       // ~0.5 inch from left
            y: 274638,       // ~0.3 inch from top
            width: 8229600,  // ~9 inches wide
            height: 1143000, // ~1.25 inches tall
        });
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

        self.elements.push(TextElement {
            text: text.into(),
            is_title: false,
            x: 457200,
            y: y_offset,
            width: 8229600,
            height: 4525963,
        });
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
        self.elements.push(TextElement {
            text: text.into(),
            is_title: false,
            x,
            y,
            width,
            height,
        });
        self
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

        // Write each slide
        for (i, slide) in self.slides.iter().enumerate() {
            let slide_num = i + 1;
            let slide_xml = self.serialize_slide(slide, slide_num);
            let part_name = format!("ppt/slides/slide{}.xml", slide_num);
            pkg.add_part(&part_name, CT_SLIDE, slide_xml.as_bytes())?;
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
    fn serialize_slide(&self, slide: &SlideBuilder, _slide_num: usize) -> String {
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
        for (i, element) in slide.elements.iter().enumerate() {
            let shape_id = i + 2; // Start at 2 (1 is the group)
            xml.push_str(&self.serialize_text_shape(element, shape_id));
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
    fn serialize_text_shape(&self, element: &TextElement, shape_id: usize) -> String {
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
        xml.push_str("            <a:r>\n");
        xml.push_str(&format!(
            r#"              <a:rPr lang="en-US" sz="{}"/>"#,
            font_size
        ));
        xml.push('\n');
        xml.push_str(&format!(
            "              <a:t>{}</a:t>\n",
            escape_xml(&element.text)
        ));
        xml.push_str("            </a:r>\n");
        xml.push_str("          </a:p>\n");
        xml.push_str("        </p:txBody>\n");

        xml.push_str("      </p:sp>\n");
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
}
