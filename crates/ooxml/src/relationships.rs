//! Relationships handling for OPC packages.
//!
//! Relationships (.rels files) define connections between package parts.
//! For example, a document has relationships to its styles, images, etc.
//!
//! # Structure
//!
//! Relationships are stored in `_rels/*.rels` files:
//! - `_rels/.rels` - Package-level relationships
//! - `word/_rels/document.xml.rels` - Relationships for word/document.xml
//!
//! Each relationship has:
//! - `Id` - Unique identifier within the .rels file (e.g., "rId1")
//! - `Type` - URI identifying the relationship type
//! - `Target` - Path to the target part (relative or absolute)
//! - `TargetMode` (optional) - "Internal" (default) or "External"

use crate::error::{Error, Result};
use std::collections::HashMap;
use std::io::Read;

/// A collection of relationships from a .rels file.
#[derive(Debug, Clone, Default)]
pub struct Relationships {
    /// Relationships indexed by ID.
    relationships: HashMap<String, Relationship>,
}

impl Relationships {
    /// Create an empty relationships collection.
    pub fn new() -> Self {
        Self::default()
    }

    /// Parse relationships from XML.
    pub fn parse<R: Read>(reader: R) -> Result<Self> {
        use quick_xml::Reader;
        use quick_xml::events::Event;

        let mut xml = Reader::from_reader(std::io::BufReader::new(reader));
        xml.config_mut().trim_text(true);

        let mut relationships = Self::new();
        let mut buf = Vec::new();

        loop {
            match xml.read_event_into(&mut buf) {
                Ok(Event::Empty(e)) | Ok(Event::Start(e)) => {
                    if e.name().as_ref() == b"Relationship" {
                        let mut id = None;
                        let mut rel_type = None;
                        let mut target = None;
                        let mut target_mode = TargetMode::Internal;

                        for attr in e.attributes().filter_map(|a| a.ok()) {
                            match attr.key.as_ref() {
                                b"Id" => {
                                    id = Some(String::from_utf8_lossy(&attr.value).into_owned());
                                }
                                b"Type" => {
                                    rel_type =
                                        Some(String::from_utf8_lossy(&attr.value).into_owned());
                                }
                                b"Target" => {
                                    target =
                                        Some(String::from_utf8_lossy(&attr.value).into_owned());
                                }
                                b"TargetMode" => {
                                    let mode = String::from_utf8_lossy(&attr.value);
                                    if mode == "External" {
                                        target_mode = TargetMode::External;
                                    }
                                }
                                _ => {}
                            }
                        }

                        if let (Some(id), Some(rel_type), Some(target)) = (id, rel_type, target) {
                            relationships.relationships.insert(
                                id.clone(),
                                Relationship {
                                    id,
                                    relationship_type: rel_type,
                                    target,
                                    target_mode,
                                },
                            );
                        }
                    }
                }
                Ok(Event::Eof) => break,
                Err(e) => return Err(Error::Xml(e)),
                _ => {}
            }
            buf.clear();
        }

        Ok(relationships)
    }

    /// Serialize relationships to XML.
    pub fn serialize(&self) -> String {
        let mut xml = String::from(r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>"#);
        xml.push('\n');
        xml.push_str(
            r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">"#,
        );

        for rel in self.relationships.values() {
            xml.push_str(&format!(
                r#"<Relationship Id="{}" Type="{}" Target="{}""#,
                rel.id, rel.relationship_type, rel.target
            ));
            if rel.target_mode == TargetMode::External {
                xml.push_str(r#" TargetMode="External""#);
            }
            xml.push_str("/>");
        }

        xml.push_str("</Relationships>");
        xml
    }

    /// Add a relationship.
    pub fn add(&mut self, rel: Relationship) {
        self.relationships.insert(rel.id.clone(), rel);
    }

    /// Get a relationship by ID.
    pub fn get(&self, id: &str) -> Option<&Relationship> {
        self.relationships.get(id)
    }

    /// Get the first relationship of a given type.
    pub fn get_by_type(&self, rel_type: &str) -> Option<&Relationship> {
        self.relationships
            .values()
            .find(|r| r.relationship_type == rel_type)
    }

    /// Get all relationships of a given type.
    pub fn get_all_by_type(&self, rel_type: &str) -> impl Iterator<Item = &Relationship> {
        self.relationships
            .values()
            .filter(move |r| r.relationship_type == rel_type)
    }

    /// Iterate over all relationships.
    pub fn iter(&self) -> impl Iterator<Item = &Relationship> {
        self.relationships.values()
    }

    /// Get the number of relationships.
    pub fn len(&self) -> usize {
        self.relationships.len()
    }

    /// Check if there are no relationships.
    pub fn is_empty(&self) -> bool {
        self.relationships.is_empty()
    }

    /// Generate a new unique relationship ID.
    pub fn next_id(&self) -> String {
        let mut n = 1;
        loop {
            let id = format!("rId{}", n);
            if !self.relationships.contains_key(&id) {
                return id;
            }
            n += 1;
        }
    }
}

/// A single relationship in an OPC package.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Relationship {
    /// Unique identifier (e.g., "rId1").
    pub id: String,
    /// Relationship type URI.
    pub relationship_type: String,
    /// Target path (relative to the source part, or absolute if external).
    pub target: String,
    /// Whether the target is internal or external.
    pub target_mode: TargetMode,
}

impl Relationship {
    /// Create a new internal relationship.
    pub fn new(
        id: impl Into<String>,
        rel_type: impl Into<String>,
        target: impl Into<String>,
    ) -> Self {
        Self {
            id: id.into(),
            relationship_type: rel_type.into(),
            target: target.into(),
            target_mode: TargetMode::Internal,
        }
    }

    /// Create a new external relationship.
    pub fn external(
        id: impl Into<String>,
        rel_type: impl Into<String>,
        target: impl Into<String>,
    ) -> Self {
        Self {
            id: id.into(),
            relationship_type: rel_type.into(),
            target: target.into(),
            target_mode: TargetMode::External,
        }
    }

    /// Check if this is an external relationship (e.g., hyperlink).
    pub fn is_external(&self) -> bool {
        self.target_mode == TargetMode::External
    }
}

/// Whether a relationship target is internal or external.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub enum TargetMode {
    /// Target is within the package.
    #[default]
    Internal,
    /// Target is external (e.g., a URL).
    External,
}

/// Common relationship type URIs.
pub mod rel_type {
    /// Office document (main document part).
    pub const OFFICE_DOCUMENT: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument";

    /// Styles part.
    pub const STYLES: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles";

    /// Numbering definitions.
    pub const NUMBERING: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering";

    /// Font table.
    pub const FONT_TABLE: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/fontTable";

    /// Settings.
    pub const SETTINGS: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings";

    /// Web settings.
    pub const WEB_SETTINGS: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/webSettings";

    /// Theme.
    pub const THEME: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme";

    /// Image.
    pub const IMAGE: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image";

    /// Hyperlink.
    pub const HYPERLINK: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink";

    /// Header.
    pub const HEADER: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/header";

    /// Footer.
    pub const FOOTER: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer";

    /// Footnotes.
    pub const FOOTNOTES: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes";

    /// Endnotes.
    pub const ENDNOTES: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/endnotes";

    /// Comments.
    pub const COMMENTS: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments";

    /// Core properties.
    pub const CORE_PROPERTIES: &str =
        "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties";

    /// Extended properties.
    pub const EXTENDED_PROPERTIES: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties";
}

/// Get the relationships file path for a given part.
///
/// For example:
/// - `word/document.xml` -> `word/_rels/document.xml.rels`
/// - Package root -> `_rels/.rels`
pub fn rels_path_for(part_path: &str) -> String {
    if part_path.is_empty() {
        return "_rels/.rels".to_string();
    }

    if let Some(slash_pos) = part_path.rfind('/') {
        let dir = &part_path[..slash_pos];
        let file = &part_path[slash_pos + 1..];
        format!("{}/_rels/{}.rels", dir, file)
    } else {
        format!("_rels/{}.rels", part_path)
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_parse_relationships() {
        let xml = r#"<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="word/styles.xml"/>
</Relationships>"#;

        let rels = Relationships::parse(xml.as_bytes()).unwrap();

        assert_eq!(rels.len(), 2);

        let r1 = rels.get("rId1").unwrap();
        assert_eq!(r1.target, "word/document.xml");
        assert_eq!(r1.relationship_type, rel_type::OFFICE_DOCUMENT);
        assert!(!r1.is_external());

        let doc = rels.get_by_type(rel_type::OFFICE_DOCUMENT).unwrap();
        assert_eq!(doc.id, "rId1");
    }

    #[test]
    fn test_external_relationship() {
        let xml = r#"<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://example.com" TargetMode="External"/>
</Relationships>"#;

        let rels = Relationships::parse(xml.as_bytes()).unwrap();
        let r1 = rels.get("rId1").unwrap();

        assert!(r1.is_external());
        assert_eq!(r1.target, "https://example.com");
    }

    #[test]
    fn test_relationships_roundtrip() {
        let mut rels = Relationships::new();
        rels.add(Relationship::new(
            "rId1",
            rel_type::OFFICE_DOCUMENT,
            "word/document.xml",
        ));
        rels.add(Relationship::external(
            "rId2",
            rel_type::HYPERLINK,
            "https://example.com",
        ));

        let xml = rels.serialize();
        let parsed = Relationships::parse(xml.as_bytes()).unwrap();

        assert_eq!(parsed.len(), 2);
        assert_eq!(parsed.get("rId1").unwrap().target, "word/document.xml");
        assert!(parsed.get("rId2").unwrap().is_external());
    }

    #[test]
    fn test_rels_path_for() {
        assert_eq!(rels_path_for(""), "_rels/.rels");
        assert_eq!(
            rels_path_for("word/document.xml"),
            "word/_rels/document.xml.rels"
        );
        assert_eq!(rels_path_for("document.xml"), "_rels/document.xml.rels");
    }

    #[test]
    fn test_next_id() {
        let mut rels = Relationships::new();
        assert_eq!(rels.next_id(), "rId1");

        rels.add(Relationship::new("rId1", "type", "target"));
        assert_eq!(rels.next_id(), "rId2");

        rels.add(Relationship::new("rId2", "type", "target"));
        rels.add(Relationship::new("rId3", "type", "target"));
        assert_eq!(rels.next_id(), "rId4");
    }
}
