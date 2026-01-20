//! Presentation API for reading and writing PowerPoint files.
//!
//! This module provides the main entry point for working with PPTX files.

use crate::error::{Error, Result};
use ooxml::{Package, Relationships};
use quick_xml::Reader;
use quick_xml::events::Event;
use std::fs::File;
use std::io::{BufReader, Cursor, Read, Seek};
use std::path::Path;

// Relationship types (ECMA-376 Part 1)
const REL_OFFICE_DOCUMENT: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument";
const REL_SLIDE: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide";

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
}

/// Metadata about a slide.
#[derive(Debug, Clone)]
struct SlideInfo {
    #[allow(dead_code)]
    rel_id: String,
    path: String,
    index: usize,
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

        // Build slide info from relationships
        let mut slide_info: Vec<SlideInfo> = Vec::new();
        for rel in pres_rels.iter() {
            if rel.relationship_type == REL_SLIDE {
                let path = resolve_path(&presentation_path, &rel.target);
                // Find index from slide order
                let index = slide_order
                    .iter()
                    .position(|id| id == &rel.id)
                    .unwrap_or(slide_info.len());
                slide_info.push(SlideInfo {
                    rel_id: rel.id.clone(),
                    path,
                    index,
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
        })
    }

    /// Get the number of slides in the presentation.
    pub fn slide_count(&self) -> usize {
        self.slide_info.len()
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
        parse_slide(&data, info.index)
    }
}

/// A slide in the presentation.
#[derive(Debug, Clone)]
pub struct Slide {
    index: usize,
    shapes: Vec<Shape>,
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
}

/// A shape on a slide.
#[derive(Debug, Clone)]
pub struct Shape {
    /// Shape name (if any).
    name: Option<String>,
    /// Text paragraphs (DrawingML).
    paragraphs: Vec<ooxml_dml::Paragraph>,
}

impl Shape {
    /// Get the shape name.
    pub fn name(&self) -> Option<&str> {
        self.name.as_deref()
    }

    /// Get the text paragraphs.
    pub fn paragraphs(&self) -> &[ooxml_dml::Paragraph] {
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
fn parse_slide(xml: &[u8], index: usize) -> Result<Slide> {
    let mut reader = Reader::from_reader(Cursor::new(xml));
    let mut buf = Vec::new();
    let mut shapes = Vec::new();

    let mut current_shape_name: Option<String> = None;
    let mut current_paragraphs: Vec<ooxml_dml::Paragraph> = Vec::new();
    let mut in_sp = false; // Inside a shape

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
                    b"p:cNvPr" => {
                        // Non-visual shape properties - get name
                        if in_sp {
                            for attr in e.attributes().filter_map(|a| a.ok()) {
                                if attr.key.as_ref() == b"name" {
                                    current_shape_name =
                                        Some(String::from_utf8_lossy(&attr.value).into_owned());
                                }
                            }
                        }
                    }
                    b"p:txBody" => {
                        // Use DML parser for text body content
                        if in_sp
                            && let Ok(paras) = ooxml_dml::parse_text_body_from_reader(&mut reader)
                        {
                            current_paragraphs = paras;
                        }
                    }
                    _ => {}
                }
            }
            Ok(Event::Empty(e)) => {
                let name = e.name();
                let name = name.as_ref();
                // Handle self-closing cNvPr
                if in_sp && name == b"p:cNvPr" {
                    for attr in e.attributes().filter_map(|a| a.ok()) {
                        if attr.key.as_ref() == b"name" {
                            current_shape_name =
                                Some(String::from_utf8_lossy(&attr.value).into_owned());
                        }
                    }
                }
            }
            Ok(Event::End(e)) => {
                let name = e.name();
                let name = name.as_ref();
                if name == b"p:sp" {
                    // End of shape
                    shapes.push(Shape {
                        name: current_shape_name.take(),
                        paragraphs: std::mem::take(&mut current_paragraphs),
                    });
                    in_sp = false;
                }
            }
            Ok(Event::Eof) => break,
            Err(e) => return Err(Error::Xml(e)),
            _ => {}
        }
        buf.clear();
    }

    Ok(Slide { index, shapes })
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
