//! Raw XML preservation for round-trip fidelity.
//!
//! This module provides types for storing unparsed XML elements,
//! allowing documents to survive read→write cycles without losing
//! features we don't explicitly understand.

use quick_xml::events::{BytesCData, BytesEnd, BytesStart, BytesText, Event};
use quick_xml::{Reader, Writer};
use std::io::{BufRead, Write};

use crate::{Error, Result};

/// A raw XML node with its original position for correct round-trip ordering.
///
/// When unknown elements are captured during parsing, we store their position
/// among siblings so they can be interleaved correctly during serialization.
#[derive(Clone, Debug, PartialEq)]
pub struct PositionedNode {
    /// Original position among sibling elements (0-indexed).
    pub position: usize,
    /// The preserved XML node.
    pub node: RawXmlNode,
}

impl PositionedNode {
    /// Create a new positioned node.
    pub fn new(position: usize, node: RawXmlNode) -> Self {
        Self { position, node }
    }
}

/// An XML attribute with its original position for correct round-trip ordering.
///
/// When unknown attributes are captured during parsing, we store their position
/// among sibling attributes so they can be serialized in the original order.
#[derive(Clone, Debug, PartialEq)]
pub struct PositionedAttr {
    /// Original position among sibling attributes (0-indexed).
    pub position: usize,
    /// The attribute name (including namespace prefix if present).
    pub name: String,
    /// The attribute value.
    pub value: String,
}

impl PositionedAttr {
    /// Create a new positioned attribute.
    pub fn new(position: usize, name: impl Into<String>, value: impl Into<String>) -> Self {
        Self {
            position,
            name: name.into(),
            value: value.into(),
        }
    }
}

/// A raw XML node that can be preserved during round-trip.
#[derive(Clone, Debug, PartialEq)]
pub enum RawXmlNode {
    /// An XML element with name, attributes, and children.
    Element(RawXmlElement),
    /// Text content.
    Text(String),
    /// CDATA content.
    CData(String),
    /// A comment.
    Comment(String),
}

/// A raw XML element with its name, attributes, and children preserved.
#[derive(Clone, Debug, PartialEq)]
pub struct RawXmlElement {
    /// The full element name (including namespace prefix if present).
    pub name: String,
    /// Element attributes as (name, value) pairs.
    pub attributes: Vec<(String, String)>,
    /// Child nodes.
    pub children: Vec<RawXmlNode>,
    /// Whether this was a self-closing element.
    pub self_closing: bool,
}

impl RawXmlElement {
    /// Create a new empty element.
    pub fn new(name: impl Into<String>) -> Self {
        Self {
            name: name.into(),
            attributes: Vec::new(),
            children: Vec::new(),
            self_closing: false,
        }
    }

    /// Parse a raw XML element from a reader, starting after the opening tag.
    ///
    /// The `start` parameter should be the BytesStart event that opened this element.
    pub fn from_reader<R: BufRead>(reader: &mut Reader<R>, start: &BytesStart) -> Result<Self> {
        let name = String::from_utf8_lossy(start.name().as_ref()).to_string();

        let attributes = start
            .attributes()
            .filter_map(|a| a.ok())
            .map(|a| {
                (
                    String::from_utf8_lossy(a.key.as_ref()).to_string(),
                    String::from_utf8_lossy(&a.value).to_string(),
                )
            })
            .collect();

        let mut element = RawXmlElement {
            name: name.clone(),
            attributes,
            children: Vec::new(),
            self_closing: false,
        };

        let mut buf = Vec::new();
        let target_name = start.name().as_ref().to_vec();

        loop {
            match reader.read_event_into(&mut buf) {
                Ok(Event::Start(e)) => {
                    let child = RawXmlElement::from_reader(reader, &e)?;
                    element.children.push(RawXmlNode::Element(child));
                }
                Ok(Event::Empty(e)) => {
                    let child = RawXmlElement::from_empty(&e);
                    element.children.push(RawXmlNode::Element(child));
                }
                Ok(Event::Text(e)) => {
                    // Use decode() for quick-xml 0.38+
                    let text = e.decode().unwrap_or_default();
                    if !text.is_empty() {
                        element.children.push(RawXmlNode::Text(text.to_string()));
                    }
                }
                Ok(Event::CData(e)) => {
                    let text = String::from_utf8_lossy(&e).to_string();
                    element.children.push(RawXmlNode::CData(text));
                }
                Ok(Event::Comment(e)) => {
                    let text = String::from_utf8_lossy(&e).to_string();
                    element.children.push(RawXmlNode::Comment(text));
                }
                Ok(Event::End(e)) => {
                    if e.name().as_ref() == target_name {
                        break;
                    }
                }
                Ok(Event::Eof) => {
                    return Err(Error::Invalid(format!(
                        "Unexpected EOF while parsing element '{}'",
                        name
                    )));
                }
                Err(e) => return Err(Error::Xml(e)),
                _ => {}
            }
            buf.clear();
        }

        Ok(element)
    }

    /// Create from an empty/self-closing element.
    pub fn from_empty(start: &BytesStart) -> Self {
        let name = String::from_utf8_lossy(start.name().as_ref()).to_string();

        let attributes = start
            .attributes()
            .filter_map(|a| a.ok())
            .map(|a| {
                (
                    String::from_utf8_lossy(a.key.as_ref()).to_string(),
                    String::from_utf8_lossy(&a.value).to_string(),
                )
            })
            .collect();

        RawXmlElement {
            name,
            attributes,
            children: Vec::new(),
            self_closing: true,
        }
    }

    /// Write this element to an XML writer.
    pub fn write_to<W: Write>(&self, writer: &mut Writer<W>) -> Result<()> {
        let mut start = BytesStart::new(&self.name);
        for (key, value) in &self.attributes {
            start.push_attribute((key.as_str(), value.as_str()));
        }

        if self.self_closing && self.children.is_empty() {
            writer.write_event(Event::Empty(start))?;
        } else {
            writer.write_event(Event::Start(start))?;

            for child in &self.children {
                child.write_to(writer)?;
            }

            writer.write_event(Event::End(BytesEnd::new(&self.name)))?;
        }

        Ok(())
    }
}

impl RawXmlNode {
    /// Write this node to an XML writer.
    pub fn write_to<W: Write>(&self, writer: &mut Writer<W>) -> Result<()> {
        match self {
            RawXmlNode::Element(elem) => elem.write_to(writer),
            RawXmlNode::Text(text) => {
                writer.write_event(Event::Text(BytesText::new(text)))?;
                Ok(())
            }
            RawXmlNode::CData(text) => {
                writer.write_event(Event::CData(BytesCData::new(text)))?;
                Ok(())
            }
            RawXmlNode::Comment(text) => {
                writer.write_event(Event::Comment(BytesText::new(text)))?;
                Ok(())
            }
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use std::io::Cursor;

    #[test]
    fn test_parse_simple_element() {
        let xml = r#"<w:test attr="value">content</w:test>"#;
        let mut reader = Reader::from_str(xml);
        let mut buf = Vec::new();

        if let Ok(Event::Start(e)) = reader.read_event_into(&mut buf) {
            let elem = RawXmlElement::from_reader(&mut reader, &e).unwrap();
            assert_eq!(elem.name, "w:test");
            assert_eq!(
                elem.attributes,
                vec![("attr".to_string(), "value".to_string())]
            );
            assert_eq!(elem.children.len(), 1);
            if let RawXmlNode::Text(t) = &elem.children[0] {
                assert_eq!(t, "content");
            } else {
                panic!("Expected text node");
            }
        }
    }

    #[test]
    fn test_parse_nested_elements() {
        let xml = r#"<parent><child1/><child2>text</child2></parent>"#;
        let mut reader = Reader::from_str(xml);
        let mut buf = Vec::new();

        if let Ok(Event::Start(e)) = reader.read_event_into(&mut buf) {
            let elem = RawXmlElement::from_reader(&mut reader, &e).unwrap();
            assert_eq!(elem.name, "parent");
            assert_eq!(elem.children.len(), 2);
        }
    }

    #[test]
    fn test_roundtrip() {
        let xml = r#"<w:test attr="value"><w:child>text</w:child></w:test>"#;
        let mut reader = Reader::from_str(xml);
        let mut buf = Vec::new();

        if let Ok(Event::Start(e)) = reader.read_event_into(&mut buf) {
            let elem = RawXmlElement::from_reader(&mut reader, &e).unwrap();

            let mut output = Vec::new();
            let mut writer = Writer::new(Cursor::new(&mut output));
            elem.write_to(&mut writer).unwrap();

            let output_str = String::from_utf8(output).unwrap();
            assert_eq!(output_str, xml);
        }
    }
}
