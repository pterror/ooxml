//! OPC (Open Packaging Conventions) implementation.
//!
//! OOXML files are ZIP archives following the OPC specification (ECMA-376 Part 2).
//! This module handles reading and writing these packages.
//!
//! # Structure
//!
//! An OPC package contains:
//! - `[Content_Types].xml` - MIME type mappings for parts
//! - `_rels/.rels` - Package-level relationships
//! - Various parts (XML files, images, etc.)
//! - Part-specific relationships in `*/_rels/*.rels`

use crate::error::{Error, Result};
use std::collections::HashMap;
use std::io::{Read, Seek, Write};
use zip::read::ZipArchive;
use zip::write::ZipWriter;

/// An OPC package (ZIP-based container for OOXML files).
pub struct Package<R> {
    archive: ZipArchive<R>,
    content_types: ContentTypes,
}

impl<R: Read + Seek> Package<R> {
    /// Open an OPC package from a reader.
    pub fn open(reader: R) -> Result<Self> {
        let mut archive = ZipArchive::new(reader)?;

        // Parse [Content_Types].xml (required)
        let content_types = Self::read_content_types(&mut archive)?;

        Ok(Self {
            archive,
            content_types,
        })
    }

    /// Read [Content_Types].xml from the archive.
    fn read_content_types(archive: &mut ZipArchive<R>) -> Result<ContentTypes> {
        let file = archive
            .by_name("[Content_Types].xml")
            .map_err(|_| Error::MissingPart("[Content_Types].xml".into()))?;

        ContentTypes::parse(file)
    }

    /// Get the content types for this package.
    pub fn content_types(&self) -> &ContentTypes {
        &self.content_types
    }

    /// Check if a part exists in the package.
    pub fn has_part(&self, path: &str) -> bool {
        self.archive.file_names().any(|name| name == path)
    }

    /// List all parts in the package.
    pub fn parts(&self) -> impl Iterator<Item = &str> {
        self.archive.file_names()
    }

    /// Read a part's contents as bytes.
    pub fn read_part(&mut self, path: &str) -> Result<Vec<u8>> {
        let mut file = self
            .archive
            .by_name(path)
            .map_err(|_| Error::MissingPart(path.into()))?;

        let mut contents = Vec::new();
        file.read_to_end(&mut contents)?;
        Ok(contents)
    }

    /// Read a part's contents as a string.
    pub fn read_part_string(&mut self, path: &str) -> Result<String> {
        let bytes = self.read_part(path)?;
        String::from_utf8(bytes)
            .map_err(|e| Error::Invalid(format!("invalid UTF-8 in {}: {}", path, e)))
    }

    /// Get the content type for a part.
    pub fn content_type(&self, path: &str) -> Option<&str> {
        self.content_types.get(path)
    }

    /// Read package-level relationships (_rels/.rels).
    pub fn read_relationships(&mut self) -> Result<crate::relationships::Relationships> {
        self.read_part_relationships("")
    }

    /// Read relationships for a specific part.
    pub fn read_part_relationships(
        &mut self,
        part_path: &str,
    ) -> Result<crate::relationships::Relationships> {
        let rels_path = crate::relationships::rels_path_for(part_path);

        if !self.has_part(&rels_path) {
            return Ok(crate::relationships::Relationships::new());
        }

        let data = self.read_part(&rels_path)?;
        crate::relationships::Relationships::parse(&data[..])
    }
}

/// Builder for creating new OPC packages.
pub struct PackageWriter<W: Write + Seek> {
    writer: ZipWriter<W>,
    content_types: ContentTypes,
}

impl<W: Write + Seek> PackageWriter<W> {
    /// Create a new package writer.
    pub fn new(writer: W) -> Self {
        Self {
            writer: ZipWriter::new(writer),
            content_types: ContentTypes::new(),
        }
    }

    /// Add a part to the package.
    pub fn add_part(&mut self, path: &str, content_type: &str, data: &[u8]) -> Result<()> {
        // Register content type
        self.content_types.add_override(path, content_type);

        // Write to ZIP
        let options = zip::write::SimpleFileOptions::default()
            .compression_method(zip::CompressionMethod::Deflated);
        self.writer.start_file(path, options)?;
        self.writer.write_all(data)?;

        Ok(())
    }

    /// Add a default content type mapping for a file extension.
    pub fn add_default_content_type(&mut self, extension: &str, content_type: &str) {
        self.content_types.add_default(extension, content_type);
    }

    /// Finish writing the package.
    pub fn finish(mut self) -> Result<W> {
        // Write [Content_Types].xml
        let content_types_xml = self.content_types.serialize();
        let options = zip::write::SimpleFileOptions::default()
            .compression_method(zip::CompressionMethod::Deflated);
        self.writer.start_file("[Content_Types].xml", options)?;
        self.writer.write_all(content_types_xml.as_bytes())?;

        Ok(self.writer.finish()?)
    }
}

/// Content type mappings for package parts.
///
/// Maps file extensions and specific part names to MIME types.
#[derive(Debug, Clone, Default)]
pub struct ContentTypes {
    /// Default mappings by extension (e.g., "xml" -> "application/xml").
    defaults: HashMap<String, String>,
    /// Override mappings for specific parts (e.g., "/word/document.xml" -> "...").
    overrides: HashMap<String, String>,
}

impl ContentTypes {
    /// Create empty content types.
    pub fn new() -> Self {
        Self::default()
    }

    /// Parse content types from XML.
    pub fn parse<R: Read>(reader: R) -> Result<Self> {
        use quick_xml::Reader;
        use quick_xml::events::Event;

        let mut xml = Reader::from_reader(std::io::BufReader::new(reader));
        xml.config_mut().trim_text(true);

        let mut content_types = Self::new();
        let mut buf = Vec::new();

        loop {
            match xml.read_event_into(&mut buf) {
                Ok(Event::Empty(e)) => {
                    let name = e.name();
                    if name.as_ref() == b"Default" {
                        let mut extension = None;
                        let mut content_type = None;

                        for attr in e.attributes().filter_map(|a| a.ok()) {
                            match attr.key.as_ref() {
                                b"Extension" => {
                                    extension =
                                        Some(String::from_utf8_lossy(&attr.value).into_owned());
                                }
                                b"ContentType" => {
                                    content_type =
                                        Some(String::from_utf8_lossy(&attr.value).into_owned());
                                }
                                _ => {}
                            }
                        }

                        if let (Some(ext), Some(ct)) = (extension, content_type) {
                            content_types.defaults.insert(ext, ct);
                        }
                    } else if name.as_ref() == b"Override" {
                        let mut part_name = None;
                        let mut content_type = None;

                        for attr in e.attributes().filter_map(|a| a.ok()) {
                            match attr.key.as_ref() {
                                b"PartName" => {
                                    part_name =
                                        Some(String::from_utf8_lossy(&attr.value).into_owned());
                                }
                                b"ContentType" => {
                                    content_type =
                                        Some(String::from_utf8_lossy(&attr.value).into_owned());
                                }
                                _ => {}
                            }
                        }

                        if let (Some(pn), Some(ct)) = (part_name, content_type) {
                            // Normalize path (remove leading /)
                            let normalized = pn.strip_prefix('/').unwrap_or(&pn);
                            content_types.overrides.insert(normalized.to_string(), ct);
                        }
                    }
                }
                Ok(Event::Eof) => break,
                Err(e) => return Err(Error::Xml(e)),
                _ => {}
            }
            buf.clear();
        }

        Ok(content_types)
    }

    /// Serialize content types to XML.
    pub fn serialize(&self) -> String {
        let mut xml = String::from(r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>"#);
        xml.push('\n');
        xml.push_str(
            r#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">"#,
        );

        for (ext, ct) in &self.defaults {
            xml.push_str(&format!(
                r#"<Default Extension="{}" ContentType="{}"/>"#,
                ext, ct
            ));
        }

        for (part, ct) in &self.overrides {
            // Ensure leading /
            let part_name = if part.starts_with('/') {
                part.clone()
            } else {
                format!("/{}", part)
            };
            xml.push_str(&format!(
                r#"<Override PartName="{}" ContentType="{}"/>"#,
                part_name, ct
            ));
        }

        xml.push_str("</Types>");
        xml
    }

    /// Add a default content type mapping.
    pub fn add_default(&mut self, extension: &str, content_type: &str) {
        self.defaults
            .insert(extension.to_string(), content_type.to_string());
    }

    /// Add an override content type mapping.
    pub fn add_override(&mut self, part_name: &str, content_type: &str) {
        let normalized = part_name.strip_prefix('/').unwrap_or(part_name);
        self.overrides
            .insert(normalized.to_string(), content_type.to_string());
    }

    /// Get the content type for a part.
    pub fn get(&self, part_name: &str) -> Option<&str> {
        let normalized = part_name.strip_prefix('/').unwrap_or(part_name);

        // Check overrides first
        if let Some(ct) = self.overrides.get(normalized) {
            return Some(ct);
        }

        // Fall back to default by extension
        if let Some(ext) = normalized.rsplit('.').next()
            && let Some(ct) = self.defaults.get(ext)
        {
            return Some(ct);
        }

        None
    }

    /// Iterate over default mappings.
    pub fn defaults(&self) -> impl Iterator<Item = (&str, &str)> {
        self.defaults.iter().map(|(k, v)| (k.as_str(), v.as_str()))
    }

    /// Iterate over override mappings.
    pub fn overrides(&self) -> impl Iterator<Item = (&str, &str)> {
        self.overrides.iter().map(|(k, v)| (k.as_str(), v.as_str()))
    }
}

/// Common content types used in OOXML packages.
pub mod content_type {
    /// Relationships content type.
    pub const RELATIONSHIPS: &str = "application/vnd.openxmlformats-package.relationships+xml";

    /// XML content type.
    pub const XML: &str = "application/xml";

    /// WordprocessingML document.
    pub const WORDPROCESSING_DOCUMENT: &str =
        "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml";

    /// WordprocessingML styles.
    pub const WORDPROCESSING_STYLES: &str =
        "application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml";

    /// Core properties (Dublin Core metadata).
    pub const CORE_PROPERTIES: &str = "application/vnd.openxmlformats-package.core-properties+xml";

    /// Extended properties (app-specific metadata).
    pub const EXTENDED_PROPERTIES: &str =
        "application/vnd.openxmlformats-officedocument.extended-properties+xml";
}

#[cfg(test)]
mod tests {
    use super::*;
    use std::io::Cursor;

    fn create_test_package() -> Vec<u8> {
        let mut buf = Cursor::new(Vec::new());

        {
            let mut writer = PackageWriter::new(&mut buf);

            // Add default content types
            writer.add_default_content_type("rels", content_type::RELATIONSHIPS);
            writer.add_default_content_type("xml", content_type::XML);

            // Add main document
            let document = r#"<?xml version="1.0" encoding="UTF-8"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p><w:r><w:t>Hello!</w:t></w:r></w:p>
  </w:body>
</w:document>"#;
            writer
                .add_part(
                    "word/document.xml",
                    content_type::WORDPROCESSING_DOCUMENT,
                    document.as_bytes(),
                )
                .unwrap();

            // Add relationships
            let rels = r#"<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>"#;
            writer
                .add_part("_rels/.rels", content_type::RELATIONSHIPS, rels.as_bytes())
                .unwrap();

            writer.finish().unwrap();
        }

        buf.into_inner()
    }

    #[test]
    fn test_create_and_read_package() {
        let data = create_test_package();
        let cursor = Cursor::new(data);

        let mut pkg = Package::open(cursor).unwrap();

        // Check content types
        assert_eq!(
            pkg.content_type("word/document.xml"),
            Some(content_type::WORDPROCESSING_DOCUMENT)
        );
        assert_eq!(
            pkg.content_type("_rels/.rels"),
            Some(content_type::RELATIONSHIPS)
        );

        // Check parts exist
        assert!(pkg.has_part("word/document.xml"));
        assert!(pkg.has_part("_rels/.rels"));
        assert!(pkg.has_part("[Content_Types].xml"));

        // Read document
        let doc = pkg.read_part_string("word/document.xml").unwrap();
        assert!(doc.contains("Hello!"));

        // Read relationships
        let rels = pkg.read_relationships().unwrap();
        assert_eq!(rels.len(), 1);

        let doc_rel = rels
            .get_by_type(crate::relationships::rel_type::OFFICE_DOCUMENT)
            .unwrap();
        assert_eq!(doc_rel.target, "word/document.xml");
    }

    #[test]
    fn test_content_types_roundtrip() {
        let mut ct = ContentTypes::new();
        ct.add_default("xml", "application/xml");
        ct.add_default("rels", content_type::RELATIONSHIPS);
        ct.add_override("/word/document.xml", content_type::WORDPROCESSING_DOCUMENT);

        let xml = ct.serialize();
        let parsed = ContentTypes::parse(xml.as_bytes()).unwrap();

        assert_eq!(parsed.get("foo.xml"), Some("application/xml"));
        assert_eq!(parsed.get("_rels/.rels"), Some(content_type::RELATIONSHIPS));
        assert_eq!(
            parsed.get("word/document.xml"),
            Some(content_type::WORDPROCESSING_DOCUMENT)
        );
    }
}
