//! Word document API.
//!
//! This module provides the main entry point for working with DOCX files.
//!
//! The `Document<R>` reader type wraps an OPC package and provides access to
//! the document body, styles, headers/footers, footnotes, endnotes, and comments
//! using generated types from the ECMA-376 schema.
//!
//! Metadata types (`CoreProperties`, `AppProperties`, `DocumentSettings`) are
//! handwritten because they come from OPC (not WML) and are not generated.

use crate::error::{Error, Result};
use crate::ext;
use crate::generated as types;
use crate::generated_serializers::ToXml;
use ooxml_opc::{Package, PackageWriter, Relationships, rel_type, rels_path_for};
use ooxml_xml::{PositionedNode, RawXmlElement, RawXmlNode};
use quick_xml::Reader;
use quick_xml::events::Event;
use std::fs::File;
use std::io::{BufReader, Read, Seek};
use std::path::Path;

/// A Word document (.docx file).
///
/// This is the main entry point for reading Word documents. The document stores
/// parsed generated types (`types::Document`, `types::Styles`) that can be
/// queried using the extension traits in `ext`.
///
/// For writing documents, use `DocumentBuilder`.
pub struct Document<R> {
    package: Package<R>,
    gen_doc: types::Document,
    gen_styles: types::Styles,
    /// Document part relationships (for images, hyperlinks, etc.)
    doc_rels: Relationships,
    /// Path to the document part (e.g., "word/document.xml")
    doc_path: String,
    /// Path to the styles part (e.g., "word/styles.xml"), if present.
    styles_path: Option<String>,
    /// Core document properties (title, author, etc.)
    core_properties: Option<CoreProperties>,
    /// Extended application properties (word count, etc.)
    app_properties: Option<AppProperties>,
}

impl Document<BufReader<File>> {
    /// Open a Word document from a file path.
    pub fn open<P: AsRef<Path>>(path: P) -> Result<Self> {
        let file = File::open(path)?;
        let reader = BufReader::new(file);
        Self::from_reader(reader)
    }
}

impl<R: Read + Seek> Document<R> {
    /// Open a Word document from a reader.
    pub fn from_reader(reader: R) -> Result<Self> {
        let mut package = Package::open(reader)?;

        // Find the main document part via relationships
        let rels = package.read_relationships()?;
        let doc_rel = rels
            .get_by_type(rel_type::OFFICE_DOCUMENT)
            .ok_or_else(|| Error::MissingPart("main document relationship".into()))?;

        let doc_path = doc_rel.target.clone();

        // Parse the document XML using the generated parser
        let doc_xml = package.read_part(&doc_path)?;
        let gen_doc = ext::parse_document(&doc_xml)?;

        // Load document-level relationships (for images, hyperlinks, etc.)
        let doc_rels_path = rels_path_for(&doc_path);
        let doc_rels = if package.has_part(&doc_rels_path) {
            let rels_xml = package.read_part(&doc_rels_path)?;
            Relationships::parse(&rels_xml[..])?
        } else {
            Relationships::new()
        };

        // Load styles using the generated parser.
        // Check package-level rels first, then document-level rels (where real
        // .docx files from Word typically reference styles).
        let (gen_styles, styles_path) = if let Some(styles_rel) = rels.get_by_type(rel_type::STYLES)
        {
            let path = styles_rel.target.clone();
            let styles_xml = package.read_part(&path)?;
            (ext::parse_styles(&styles_xml)?, Some(path))
        } else if let Some(styles_rel) = doc_rels.get_by_type(rel_type::STYLES) {
            let path = resolve_path(&doc_path, &styles_rel.target);
            let styles_xml = package.read_part(&path)?;
            (ext::parse_styles(&styles_xml)?, Some(path))
        } else {
            (types::Styles::default(), None)
        };

        // Load core properties if available
        let core_properties = if let Some(core_rel) = rels.get_by_type(rel_type::CORE_PROPERTIES) {
            let core_xml = package.read_part(&core_rel.target)?;
            Some(parse_core_properties(&core_xml)?)
        } else {
            None
        };

        // Load app properties if available
        let app_properties = if let Some(app_rel) = rels.get_by_type(rel_type::EXTENDED_PROPERTIES)
        {
            let app_xml = package.read_part(&app_rel.target)?;
            Some(parse_app_properties(&app_xml)?)
        } else {
            None
        };

        Ok(Self {
            package,
            gen_doc,
            gen_styles,
            doc_rels,
            doc_path,
            styles_path,
            core_properties,
            app_properties,
        })
    }

    /// Get the document body.
    ///
    /// Returns the generated `Body` type. Use extension traits from `ext` to
    /// access paragraphs, runs, and text content.
    pub fn body(&self) -> &types::Body {
        self.gen_doc
            .body
            .as_deref()
            .expect("document has no body element")
    }

    /// Get a mutable reference to the document body.
    pub fn body_mut(&mut self) -> &mut types::Body {
        self.gen_doc
            .body
            .as_deref_mut()
            .expect("document has no body element")
    }

    /// Get the generated document.
    pub fn gen_doc(&self) -> &types::Document {
        &self.gen_doc
    }

    /// Get the underlying package.
    pub fn package(&self) -> &Package<R> {
        &self.package
    }

    /// Get a mutable reference to the underlying package.
    pub fn package_mut(&mut self) -> &mut Package<R> {
        &mut self.package
    }

    /// Get the document styles (generated types).
    pub fn styles(&self) -> &types::Styles {
        &self.gen_styles
    }

    /// Get the core document properties (title, author, etc.).
    ///
    /// Returns `None` if the document doesn't have a core properties part.
    pub fn core_properties(&self) -> Option<&CoreProperties> {
        self.core_properties.as_ref()
    }

    /// Get the extended application properties (word count, page count, etc.).
    ///
    /// Returns `None` if the document doesn't have an app properties part.
    pub fn app_properties(&self) -> Option<&AppProperties> {
        self.app_properties.as_ref()
    }

    /// Extract all text from the document.
    ///
    /// Paragraphs are separated by newlines.
    pub fn text(&self) -> String {
        use crate::ext::BodyExt;
        self.gen_doc
            .body
            .as_deref()
            .map(|b| b.text())
            .unwrap_or_default()
    }

    /// Get image data by relationship ID.
    ///
    /// Looks up the relationship, reads the image file from the package,
    /// and returns the image data with its content type.
    pub fn get_image_data(&mut self, rel_id: &str) -> Result<ImageData> {
        // Look up the relationship
        let rel = self
            .doc_rels
            .get(rel_id)
            .ok_or_else(|| Error::MissingPart(format!("image relationship {}", rel_id)))?;

        // Resolve the target path relative to the document
        let image_path = resolve_path(&self.doc_path, &rel.target);

        // Read the image data from the package
        let data = self.package.read_part(&image_path)?;

        // Determine content type from extension
        let content_type = content_type_from_path(&image_path);

        Ok(ImageData { content_type, data })
    }

    /// Get the URL for a hyperlink by its relationship ID.
    ///
    /// Returns None if the relationship doesn't exist.
    pub fn get_hyperlink_url(&self, rel_id: &str) -> Option<&str> {
        self.doc_rels.get(rel_id).map(|rel| rel.target.as_str())
    }

    /// Get document relationships (for advanced use).
    pub fn doc_relationships(&self) -> &Relationships {
        &self.doc_rels
    }

    /// Load a header part by its relationship ID.
    ///
    /// Returns the parsed header as a generated `HeaderFooter` type.
    pub fn get_header(&mut self, rel_id: &str) -> Result<types::HeaderFooter> {
        let rel = self
            .doc_rels
            .get(rel_id)
            .ok_or_else(|| Error::MissingPart(format!("header relationship {}", rel_id)))?;

        let header_path = resolve_path(&self.doc_path, &rel.target);
        let header_xml = self.package.read_part(&header_path)?;
        Ok(ext::parse_hdr_ftr(&header_xml)?)
    }

    /// Load a footer part by its relationship ID.
    ///
    /// Returns the parsed footer as a generated `HeaderFooter` type.
    pub fn get_footer(&mut self, rel_id: &str) -> Result<types::HeaderFooter> {
        let rel = self
            .doc_rels
            .get(rel_id)
            .ok_or_else(|| Error::MissingPart(format!("footer relationship {}", rel_id)))?;

        let footer_path = resolve_path(&self.doc_path, &rel.target);
        let footer_xml = self.package.read_part(&footer_path)?;
        Ok(ext::parse_hdr_ftr(&footer_xml)?)
    }

    /// Load the footnotes part.
    ///
    /// Returns the parsed footnotes as a generated `Footnotes` type.
    ///
    /// Returns `Error::MissingPart` if the document has no footnotes.xml.
    pub fn get_footnotes(&mut self) -> Result<types::Footnotes> {
        let footnotes_rel = self
            .doc_rels
            .get_by_type(
                "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes",
            )
            .ok_or_else(|| Error::MissingPart("footnotes relationship".into()))?;

        let footnotes_path = resolve_path(&self.doc_path, &footnotes_rel.target);
        let footnotes_xml = self.package.read_part(&footnotes_path)?;
        Ok(ext::parse_footnotes(&footnotes_xml)?)
    }

    /// Load the endnotes part.
    ///
    /// Returns the parsed endnotes as a generated `Endnotes` type.
    ///
    /// Returns `Error::MissingPart` if the document has no endnotes.xml.
    pub fn get_endnotes(&mut self) -> Result<types::Endnotes> {
        let endnotes_rel = self
            .doc_rels
            .get_by_type(
                "http://schemas.openxmlformats.org/officeDocument/2006/relationships/endnotes",
            )
            .ok_or_else(|| Error::MissingPart("endnotes relationship".into()))?;

        let endnotes_path = resolve_path(&self.doc_path, &endnotes_rel.target);
        let endnotes_xml = self.package.read_part(&endnotes_path)?;
        Ok(ext::parse_endnotes(&endnotes_xml)?)
    }

    /// Load the comments part.
    ///
    /// Returns the parsed comments as a generated `Comments` type.
    ///
    /// Returns `Error::MissingPart` if the document has no comments.xml.
    pub fn get_comments(&mut self) -> Result<types::Comments> {
        let comments_rel = self
            .doc_rels
            .get_by_type(
                "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments",
            )
            .ok_or_else(|| Error::MissingPart("comments relationship".into()))?;

        let comments_path = resolve_path(&self.doc_path, &comments_rel.target);
        let comments_xml = self.package.read_part(&comments_path)?;
        Ok(ext::parse_comments(&comments_xml)?)
    }

    /// Load the document settings.
    ///
    /// Returns the parsed settings from word/settings.xml.
    ///
    /// Returns `Error::MissingPart` if the document has no settings.xml.
    pub fn get_settings(&mut self) -> Result<DocumentSettings> {
        let settings_rel = self
            .doc_rels
            .get_by_type(
                "http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings",
            )
            .ok_or_else(|| Error::MissingPart("settings relationship".into()))?;

        let settings_path = resolve_path(&self.doc_path, &settings_rel.target);
        let settings_xml = self.package.read_part(&settings_path)?;
        parse_settings(&settings_xml)
    }

    /// Save the document to a file.
    ///
    /// This serializes the current state of the generated types (`gen_doc`,
    /// `gen_styles`) back into the package, preserving all other parts verbatim.
    pub fn save<P: AsRef<Path>>(&mut self, path: P) -> Result<()> {
        let file = File::create(path)?;
        self.write(file)
    }

    /// Write the document to a writer.
    ///
    /// Serializes the document and styles using the generated `ToXml` serializers,
    /// then copies all package parts to the output, replacing only the modified
    /// parts.
    pub fn write<W: std::io::Write + Seek>(&mut self, writer: W) -> Result<()> {
        // Serialize document XML
        let doc_xml = serialize_xml(&self.gen_doc, "w:document")?;

        // Build replacements map
        let mut replacements = std::collections::HashMap::new();
        replacements.insert(self.doc_path.as_str(), doc_xml.as_slice());

        // Serialize styles if we have a styles path
        let styles_xml;
        if let Some(ref styles_path) = self.styles_path {
            styles_xml = serialize_xml(&self.gen_styles, "w:styles")?;
            replacements.insert(styles_path.as_str(), styles_xml.as_slice());
        }

        // Create package writer and copy all parts with replacements
        let mut pkg_writer = PackageWriter::new(writer);
        self.package
            .copy_to_writer(&mut pkg_writer, &replacements)?;
        pkg_writer.finish()?;

        Ok(())
    }
}

// =============================================================================
// Serialization helper
// =============================================================================

/// Serialize a ToXml value to bytes with an XML declaration prepended.
pub(crate) fn serialize_xml(value: &impl ToXml, tag: &str) -> Result<Vec<u8>> {
    let inner = Vec::new();
    let mut writer = quick_xml::Writer::new(inner);
    value.write_element(tag, &mut writer)?;
    let inner = writer.into_inner();

    let mut buf = Vec::with_capacity(
        b"<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n".len() + inner.len(),
    );
    buf.extend_from_slice(b"<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n");
    buf.extend_from_slice(&inner);
    Ok(buf)
}

// =============================================================================
// Path utilities
// =============================================================================

/// Resolve a relative path against a base path.
pub(crate) fn resolve_path(base: &str, relative: &str) -> String {
    // If the target is absolute (starts with /), use it directly (without the /)
    if let Some(stripped) = relative.strip_prefix('/') {
        return stripped.to_string();
    }

    // Otherwise, resolve relative to the base directory
    if let Some(slash_pos) = base.rfind('/') {
        format!("{}/{}", &base[..slash_pos], relative)
    } else {
        relative.to_string()
    }
}

/// Determine MIME content type from file extension.
pub(crate) fn content_type_from_path(path: &str) -> String {
    let ext = path.rsplit('.').next().unwrap_or("").to_lowercase();
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

// =============================================================================
// Metadata types (OPC, not WML — not generated)
// =============================================================================

/// Image data loaded from the package.
#[derive(Debug, Clone)]
pub struct ImageData {
    /// MIME content type (e.g., "image/png", "image/jpeg").
    pub content_type: String,
    /// Raw image bytes.
    pub data: Vec<u8>,
}

/// Document settings.
///
/// Corresponds to the `<w:settings>` element in word/settings.xml.
/// Contains document-wide settings like default tab stop, zoom level,
/// tracking changes settings, and various compatibility options.
///
/// ECMA-376 Part 1, Section 17.15.1 (Document Settings).
#[derive(Debug, Clone, Default)]
pub struct DocumentSettings {
    /// Default tab stop width in twentieths of a point (twips).
    /// Default is 720 twips (0.5 inch).
    pub default_tab_stop: Option<u32>,
    /// Document zoom percentage (e.g., 100 = 100%).
    pub zoom_percent: Option<u32>,
    /// Whether to display the document background shape.
    pub display_background_shape: bool,
    /// Whether track revisions (track changes) is enabled.
    pub track_revisions: bool,
    /// Whether to track moves separately in tracked changes.
    pub do_not_track_moves: bool,
    /// Whether to track formatting in tracked changes.
    pub do_not_track_formatting: bool,
    /// Spelling state - whether document has been fully checked.
    pub spelling_state: Option<ProofState>,
    /// Grammar state - whether document has been fully checked.
    pub grammar_state: Option<ProofState>,
    /// Character spacing control mode.
    pub character_spacing_control: Option<CharacterSpacingControl>,
    /// Compatibility mode (e.g., "15" for Word 2013+).
    pub compat_mode: Option<u32>,
    /// Unknown child elements preserved for round-trip fidelity.
    pub unknown_children: Vec<PositionedNode>,
}

/// Proof state for spelling/grammar checking.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ProofState {
    /// Document is clean (fully checked).
    Clean,
    /// Document is dirty (needs checking).
    Dirty,
}

/// Character spacing control mode.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum CharacterSpacingControl {
    /// Do not compress punctuation.
    DoNotCompress,
    /// Compress punctuation.
    CompressPunctuation,
    /// Compress punctuation and kana.
    CompressPunctuationAndJapaneseKana,
}

/// Core document properties (Dublin Core metadata).
///
/// Corresponds to `docProps/core.xml` in the OPC package.
/// Contains standard metadata like title, author, creation date, etc.
///
/// ECMA-376 Part 2, Section 11 (Core Properties).
#[derive(Debug, Clone, Default)]
pub struct CoreProperties {
    /// Document title (dc:title).
    pub title: Option<String>,
    /// Document creator/author (dc:creator).
    pub creator: Option<String>,
    /// Document subject (dc:subject).
    pub subject: Option<String>,
    /// Document description (dc:description).
    pub description: Option<String>,
    /// Keywords (cp:keywords).
    pub keywords: Option<String>,
    /// Category (cp:category).
    pub category: Option<String>,
    /// Last person to modify the document (cp:lastModifiedBy).
    pub last_modified_by: Option<String>,
    /// Revision number (cp:revision).
    pub revision: Option<String>,
    /// Creation date as ISO 8601 string (dcterms:created).
    pub created: Option<String>,
    /// Last modified date as ISO 8601 string (dcterms:modified).
    pub modified: Option<String>,
    /// Content status (cp:contentStatus).
    pub content_status: Option<String>,
}

/// Extended application properties.
///
/// Corresponds to `docProps/app.xml` in the OPC package.
/// Contains application-specific metadata like word count, page count, etc.
///
/// ECMA-376 Part 2, Section 11.1 (Extended Properties).
#[derive(Debug, Clone, Default)]
pub struct AppProperties {
    /// Application name that created the document.
    pub application: Option<String>,
    /// Application version.
    pub app_version: Option<String>,
    /// Company name.
    pub company: Option<String>,
    /// Document manager.
    pub manager: Option<String>,
    /// Total editing time in minutes.
    pub total_time: Option<u32>,
    /// Number of pages.
    pub pages: Option<u32>,
    /// Number of words.
    pub words: Option<u32>,
    /// Number of characters (excluding spaces).
    pub characters: Option<u32>,
    /// Number of characters (including spaces).
    pub characters_with_spaces: Option<u32>,
    /// Number of paragraphs.
    pub paragraphs: Option<u32>,
    /// Number of lines.
    pub lines: Option<u32>,
    /// Document template.
    pub template: Option<String>,
    /// Document security level (0 = none, 1 = password protected, etc.).
    pub doc_security: Option<u32>,
}

// =============================================================================
// Settings parser
// =============================================================================

// Element name constants for settings parser.
const EL_SETTINGS: &[u8] = b"settings";
const EL_DEFAULT_TAB_STOP: &[u8] = b"defaultTabStop";
const EL_ZOOM: &[u8] = b"zoom";
const EL_DISPLAY_BACKGROUND_SHAPE: &[u8] = b"displayBackgroundShape";
const EL_TRACK_REVISIONS: &[u8] = b"trackRevisions";
const EL_DO_NOT_TRACK_MOVES: &[u8] = b"doNotTrackMoves";
const EL_DO_NOT_TRACK_FORMATTING: &[u8] = b"doNotTrackFormatting";
const EL_PROOF_STATE: &[u8] = b"proofState";
const EL_CHAR_SPACE_CONTROL: &[u8] = b"characterSpacingControl";
const EL_COMPAT: &[u8] = b"compat";
const EL_COMPAT_SETTING: &[u8] = b"compatSetting";

/// Parse a settings.xml file into DocumentSettings.
///
/// The structure is `<w:settings>` containing various setting elements.
fn parse_settings(xml: &[u8]) -> Result<DocumentSettings> {
    let mut reader = Reader::from_reader(xml);
    reader.config_mut().trim_text(false);

    let mut buf = Vec::new();
    let mut settings = DocumentSettings::default();
    let mut in_settings = false;
    let mut in_compat = false;
    let mut child_idx: usize = 0;

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) => {
                let name = e.name();
                let local = local_name(name.as_ref());
                if local == EL_SETTINGS {
                    in_settings = true;
                } else if in_settings && local == EL_COMPAT {
                    in_compat = true;
                    child_idx += 1;
                } else if in_settings && !in_compat {
                    // Unknown element - preserve for roundtrip
                    let node = RawXmlElement::from_reader(&mut reader, &e)?;
                    settings
                        .unknown_children
                        .push(PositionedNode::new(child_idx, RawXmlNode::Element(node)));
                    child_idx += 1;
                }
            }
            Ok(Event::Empty(e)) => {
                let name = e.name();
                let local = local_name(name.as_ref());

                if !in_settings {
                    continue;
                }

                if local == EL_DEFAULT_TAB_STOP {
                    for attr in e.attributes().flatten() {
                        let key = local_name(attr.key.as_ref());
                        if key == b"val"
                            && let Ok(s) = std::str::from_utf8(&attr.value)
                        {
                            settings.default_tab_stop = s.parse().ok();
                        }
                    }
                    child_idx += 1;
                } else if local == EL_ZOOM {
                    for attr in e.attributes().flatten() {
                        let key = local_name(attr.key.as_ref());
                        if key == b"percent"
                            && let Ok(s) = std::str::from_utf8(&attr.value)
                        {
                            settings.zoom_percent = s.parse().ok();
                        }
                    }
                    child_idx += 1;
                } else if local == EL_DISPLAY_BACKGROUND_SHAPE {
                    settings.display_background_shape = parse_toggle_val(&e);
                    child_idx += 1;
                } else if local == EL_TRACK_REVISIONS {
                    settings.track_revisions = parse_toggle_val(&e);
                    child_idx += 1;
                } else if local == EL_DO_NOT_TRACK_MOVES {
                    settings.do_not_track_moves = parse_toggle_val(&e);
                    child_idx += 1;
                } else if local == EL_DO_NOT_TRACK_FORMATTING {
                    settings.do_not_track_formatting = parse_toggle_val(&e);
                    child_idx += 1;
                } else if local == EL_PROOF_STATE {
                    for attr in e.attributes().flatten() {
                        let key = local_name(attr.key.as_ref());
                        if let Ok(s) = std::str::from_utf8(&attr.value) {
                            if key == b"spelling" {
                                settings.spelling_state = match s {
                                    "clean" => Some(ProofState::Clean),
                                    "dirty" => Some(ProofState::Dirty),
                                    _ => None,
                                };
                            } else if key == b"grammar" {
                                settings.grammar_state = match s {
                                    "clean" => Some(ProofState::Clean),
                                    "dirty" => Some(ProofState::Dirty),
                                    _ => None,
                                };
                            }
                        }
                    }
                    child_idx += 1;
                } else if local == EL_CHAR_SPACE_CONTROL {
                    for attr in e.attributes().flatten() {
                        let key = local_name(attr.key.as_ref());
                        if key == b"val"
                            && let Ok(s) = std::str::from_utf8(&attr.value)
                        {
                            settings.character_spacing_control = match s {
                                "doNotCompress" => Some(CharacterSpacingControl::DoNotCompress),
                                "compressPunctuation" => {
                                    Some(CharacterSpacingControl::CompressPunctuation)
                                }
                                "compressPunctuationAndJapaneseKana" => Some(
                                    CharacterSpacingControl::CompressPunctuationAndJapaneseKana,
                                ),
                                _ => None,
                            };
                        }
                    }
                    child_idx += 1;
                } else if in_compat && local == EL_COMPAT_SETTING {
                    // Look for w:name="compatibilityMode"
                    for attr in e.attributes().flatten() {
                        let key = local_name(attr.key.as_ref());
                        if key == b"name" && &*attr.value == b"compatibilityMode" {
                            // Get the w:val attribute
                            for attr2 in e.attributes().flatten() {
                                let key2 = local_name(attr2.key.as_ref());
                                if key2 == b"val"
                                    && let Ok(s) = std::str::from_utf8(&attr2.value)
                                {
                                    settings.compat_mode = s.parse().ok();
                                }
                            }
                        }
                    }
                } else if !in_compat {
                    // Unknown empty element - preserve for roundtrip
                    let node = RawXmlElement::from_empty(&e);
                    settings
                        .unknown_children
                        .push(PositionedNode::new(child_idx, RawXmlNode::Element(node)));
                    child_idx += 1;
                }
            }
            Ok(Event::End(e)) => {
                let name = e.name();
                let local = local_name(name.as_ref());
                if local == EL_SETTINGS {
                    break;
                } else if local == EL_COMPAT {
                    in_compat = false;
                }
            }
            Ok(Event::Eof) => break,
            Err(e) => {
                return Err(Error::Xml(e)
                    .with_context("word/settings.xml")
                    .at_position(reader.error_position()));
            }
            _ => {}
        }
        buf.clear();
    }

    Ok(settings)
}

/// Parse a core.xml file into CoreProperties.
///
/// The structure is `<cp:coreProperties>` containing Dublin Core metadata.
/// Uses namespaces: dc (Dublin Core elements), dcterms (Dublin Core terms), cp (core properties).
fn parse_core_properties(xml: &[u8]) -> Result<CoreProperties> {
    let mut reader = Reader::from_reader(xml);
    reader.config_mut().trim_text(false);

    let mut buf = Vec::new();
    let mut props = CoreProperties::default();
    let mut in_core = false;
    let mut current_element: Option<&'static str> = None;

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) => {
                let name = e.name();
                let local = local_name(name.as_ref());
                if local == b"coreProperties" {
                    in_core = true;
                } else if in_core {
                    // Map element names to field
                    current_element = match local {
                        b"title" => Some("title"),
                        b"creator" => Some("creator"),
                        b"subject" => Some("subject"),
                        b"description" => Some("description"),
                        b"keywords" => Some("keywords"),
                        b"category" => Some("category"),
                        b"lastModifiedBy" => Some("lastModifiedBy"),
                        b"revision" => Some("revision"),
                        b"created" => Some("created"),
                        b"modified" => Some("modified"),
                        b"contentStatus" => Some("contentStatus"),
                        _ => None,
                    };
                }
            }
            Ok(Event::End(e)) => {
                let name = e.name();
                let local = local_name(name.as_ref());
                if local == b"coreProperties" {
                    in_core = false;
                } else if in_core {
                    current_element = None;
                }
            }
            Ok(Event::Text(e)) if current_element.is_some() => {
                let text = e.decode().ok().map(|s| s.into_owned());
                match current_element {
                    Some("title") => props.title = text,
                    Some("creator") => props.creator = text,
                    Some("subject") => props.subject = text,
                    Some("description") => props.description = text,
                    Some("keywords") => props.keywords = text,
                    Some("category") => props.category = text,
                    Some("lastModifiedBy") => props.last_modified_by = text,
                    Some("revision") => props.revision = text,
                    Some("created") => props.created = text,
                    Some("modified") => props.modified = text,
                    Some("contentStatus") => props.content_status = text,
                    _ => {}
                }
            }
            Ok(Event::Eof) => break,
            Err(e) => {
                return Err(Error::Xml(e)
                    .with_context("docProps/core.xml")
                    .at_position(reader.error_position()));
            }
            _ => {}
        }
        buf.clear();
    }

    Ok(props)
}

/// Parse an app.xml file into AppProperties.
///
/// The structure is `<Properties>` containing extended property elements.
fn parse_app_properties(xml: &[u8]) -> Result<AppProperties> {
    let mut reader = Reader::from_reader(xml);
    reader.config_mut().trim_text(false);

    let mut buf = Vec::new();
    let mut props = AppProperties::default();
    let mut in_props = false;
    let mut current_element: Option<&'static str> = None;

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) => {
                let name = e.name();
                let local = local_name(name.as_ref());
                if local == b"Properties" {
                    in_props = true;
                } else if in_props {
                    // Map element names to field
                    current_element = match local {
                        b"Application" => Some("Application"),
                        b"AppVersion" => Some("AppVersion"),
                        b"Company" => Some("Company"),
                        b"Manager" => Some("Manager"),
                        b"TotalTime" => Some("TotalTime"),
                        b"Pages" => Some("Pages"),
                        b"Words" => Some("Words"),
                        b"Characters" => Some("Characters"),
                        b"CharactersWithSpaces" => Some("CharactersWithSpaces"),
                        b"Paragraphs" => Some("Paragraphs"),
                        b"Lines" => Some("Lines"),
                        b"Template" => Some("Template"),
                        b"DocSecurity" => Some("DocSecurity"),
                        _ => None,
                    };
                }
            }
            Ok(Event::End(e)) => {
                let name = e.name();
                let local = local_name(name.as_ref());
                if local == b"Properties" {
                    in_props = false;
                } else if in_props {
                    current_element = None;
                }
            }
            Ok(Event::Text(e)) if current_element.is_some() => {
                let text = e.decode().ok().map(|s| s.into_owned());
                match current_element {
                    Some("Application") => props.application = text,
                    Some("AppVersion") => props.app_version = text,
                    Some("Company") => props.company = text,
                    Some("Manager") => props.manager = text,
                    Some("TotalTime") => {
                        props.total_time = text.as_deref().and_then(|s| s.parse().ok())
                    }
                    Some("Pages") => props.pages = text.as_deref().and_then(|s| s.parse().ok()),
                    Some("Words") => props.words = text.as_deref().and_then(|s| s.parse().ok()),
                    Some("Characters") => {
                        props.characters = text.as_deref().and_then(|s| s.parse().ok())
                    }
                    Some("CharactersWithSpaces") => {
                        props.characters_with_spaces = text.as_deref().and_then(|s| s.parse().ok())
                    }
                    Some("Paragraphs") => {
                        props.paragraphs = text.as_deref().and_then(|s| s.parse().ok())
                    }
                    Some("Lines") => props.lines = text.as_deref().and_then(|s| s.parse().ok()),
                    Some("Template") => props.template = text,
                    Some("DocSecurity") => {
                        props.doc_security = text.as_deref().and_then(|s| s.parse().ok())
                    }
                    _ => {}
                }
            }
            Ok(Event::Eof) => break,
            Err(e) => {
                return Err(Error::Xml(e)
                    .with_context("docProps/app.xml")
                    .at_position(reader.error_position()));
            }
            _ => {}
        }
        buf.clear();
    }

    Ok(props)
}

// =============================================================================
// Helpers
// =============================================================================

/// Extract the local name from a potentially namespaced element name.
fn local_name(name: &[u8]) -> &[u8] {
    // Handle both "w:p" and "p" formats
    if let Some(pos) = name.iter().position(|&b| b == b':') {
        &name[pos + 1..]
    } else {
        name
    }
}

/// Parse a toggle property value (like <w:b/> or <w:b w:val="true"/>).
///
/// Toggle properties are true if:
/// - Element is present with no val attribute
/// - Element has val="true", "1", or "on"
fn parse_toggle_val(e: &quick_xml::events::BytesStart) -> bool {
    for attr in e.attributes().filter_map(|a| a.ok()) {
        if attr.key.as_ref() == b"w:val" || attr.key.as_ref() == b"val" {
            let val: &[u8] = &attr.value;
            return matches!(val, b"true" | b"1" | b"on" | b"True" | b"On");
        }
    }
    // No val attribute means true
    true
}

// =============================================================================
// Tests
// =============================================================================

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_resolve_path() {
        // Relative path resolution
        assert_eq!(
            resolve_path("word/document.xml", "media/image1.png"),
            "word/media/image1.png"
        );
        assert_eq!(
            resolve_path("word/document.xml", "../media/image1.png"),
            "word/../media/image1.png"
        );

        // Absolute path
        assert_eq!(
            resolve_path("word/document.xml", "/word/media/image1.png"),
            "word/media/image1.png"
        );
    }

    #[test]
    fn test_content_type_from_path() {
        assert_eq!(content_type_from_path("word/media/image1.png"), "image/png");
        assert_eq!(
            content_type_from_path("word/media/image2.jpg"),
            "image/jpeg"
        );
        assert_eq!(
            content_type_from_path("word/media/image3.JPEG"),
            "image/jpeg"
        );
        assert_eq!(content_type_from_path("word/media/image4.gif"), "image/gif");
        assert_eq!(
            content_type_from_path("word/media/unknown.xyz"),
            "application/octet-stream"
        );
    }

    #[test]
    fn test_parse_core_properties() {
        let xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties"
    xmlns:dc="http://purl.org/dc/elements/1.1/"
    xmlns:dcterms="http://purl.org/dc/terms/"
    xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
  <dc:title>Test Document Title</dc:title>
  <dc:creator>John Doe</dc:creator>
  <dc:subject>Testing</dc:subject>
  <dc:description>A test document for unit testing.</dc:description>
  <cp:keywords>test, unit, document</cp:keywords>
  <cp:category>Testing</cp:category>
  <cp:lastModifiedBy>Jane Doe</cp:lastModifiedBy>
  <cp:revision>5</cp:revision>
  <dcterms:created xsi:type="dcterms:W3CDTF">2024-01-15T10:30:00Z</dcterms:created>
  <dcterms:modified xsi:type="dcterms:W3CDTF">2024-01-16T14:45:00Z</dcterms:modified>
  <cp:contentStatus>Draft</cp:contentStatus>
</cp:coreProperties>"#;

        let props = parse_core_properties(xml).unwrap();

        assert_eq!(props.title, Some("Test Document Title".to_string()));
        assert_eq!(props.creator, Some("John Doe".to_string()));
        assert_eq!(props.subject, Some("Testing".to_string()));
        assert_eq!(
            props.description,
            Some("A test document for unit testing.".to_string())
        );
        assert_eq!(props.keywords, Some("test, unit, document".to_string()));
        assert_eq!(props.category, Some("Testing".to_string()));
        assert_eq!(props.last_modified_by, Some("Jane Doe".to_string()));
        assert_eq!(props.revision, Some("5".to_string()));
        assert_eq!(props.created, Some("2024-01-15T10:30:00Z".to_string()));
        assert_eq!(props.modified, Some("2024-01-16T14:45:00Z".to_string()));
        assert_eq!(props.content_status, Some("Draft".to_string()));
    }

    #[test]
    fn test_parse_app_properties() {
        let xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"
    xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">
  <Application>Microsoft Office Word</Application>
  <AppVersion>16.0000</AppVersion>
  <Company>Test Corp</Company>
  <Manager>Project Lead</Manager>
  <TotalTime>120</TotalTime>
  <Pages>5</Pages>
  <Words>1234</Words>
  <Characters>6789</Characters>
  <CharactersWithSpaces>8000</CharactersWithSpaces>
  <Paragraphs>45</Paragraphs>
  <Lines>100</Lines>
  <Template>Normal.dotm</Template>
  <DocSecurity>0</DocSecurity>
</Properties>"#;

        let props = parse_app_properties(xml).unwrap();

        assert_eq!(props.application, Some("Microsoft Office Word".to_string()));
        assert_eq!(props.app_version, Some("16.0000".to_string()));
        assert_eq!(props.company, Some("Test Corp".to_string()));
        assert_eq!(props.manager, Some("Project Lead".to_string()));
        assert_eq!(props.total_time, Some(120));
        assert_eq!(props.pages, Some(5));
        assert_eq!(props.words, Some(1234));
        assert_eq!(props.characters, Some(6789));
        assert_eq!(props.characters_with_spaces, Some(8000));
        assert_eq!(props.paragraphs, Some(45));
        assert_eq!(props.lines, Some(100));
        assert_eq!(props.template, Some("Normal.dotm".to_string()));
        assert_eq!(props.doc_security, Some(0));
    }
}
