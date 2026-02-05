//! Raw XML preservation for round-trip fidelity.
//!
//! This module provides types for storing unparsed XML elements,
//! allowing documents to survive read→write cycles without losing
//! features we don't explicitly understand.

use quick_xml::events::{BytesCData, BytesEnd, BytesStart, BytesText, Event};
use quick_xml::{Reader, Writer};
use std::io::{BufRead, Write};

use crate::{Error, FromXml, ParseError, Result};

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

    /// Parse this element as a typed struct using the FromXml trait.
    ///
    /// Uses a streaming approach that generates XML bytes lazily from the
    /// in-memory tree structure, avoiding full upfront serialization.
    ///
    /// # Example
    /// ```ignore
    /// use ooxml_dml::types::CTTable;
    /// if let Some(table) = raw_element.parse_as::<CTTable>() {
    ///     // Use the parsed table
    /// }
    /// ```
    pub fn parse_as<T: FromXml>(&self) -> std::result::Result<T, ParseError> {
        let streaming_reader = RawXmlStreamReader::new(self);
        let mut reader = Reader::from_reader(streaming_reader);
        let mut buf = Vec::new();

        loop {
            match reader.read_event_into(&mut buf) {
                Ok(Event::Start(e)) => {
                    return T::from_xml(&mut reader, &e, false);
                }
                Ok(Event::Empty(e)) => {
                    return T::from_xml(&mut reader, &e, true);
                }
                Ok(Event::Eof) => {
                    return Err(ParseError::UnexpectedElement(
                        "empty XML in parse_as".to_string(),
                    ));
                }
                Err(e) => return Err(ParseError::Xml(e)),
                _ => {}
            }
            buf.clear();
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

/// A streaming reader that produces XML bytes lazily from a RawXmlElement tree.
///
/// Implements `BufRead` so it can be used with `quick_xml::Reader::from_reader()`.
/// This avoids allocating the full XML string upfront - bytes are generated
/// on-demand as the parser reads.
pub struct RawXmlStreamReader<'a> {
    /// Stack of elements being processed (element, child_index, state)
    stack: Vec<(&'a RawXmlElement, usize, ElementState)>,
    /// Current buffered output
    buffer: Vec<u8>,
    /// Current read position in buffer
    pos: usize,
    /// Whether we've finished
    done: bool,
}

#[derive(Clone, Copy, PartialEq)]
enum ElementState {
    /// About to emit start tag
    Start,
    /// Emitting children
    Children,
    /// About to emit end tag
    End,
}

impl<'a> RawXmlStreamReader<'a> {
    /// Create a new streaming reader for the given element.
    pub fn new(elem: &'a RawXmlElement) -> Self {
        Self {
            stack: vec![(elem, 0, ElementState::Start)],
            buffer: Vec::with_capacity(256),
            pos: 0,
            done: false,
        }
    }

    /// Generate the next chunk of XML into the buffer.
    fn generate_next(&mut self) {
        self.buffer.clear();
        self.pos = 0;

        while self.buffer.is_empty() && !self.stack.is_empty() {
            let (elem, child_idx, state) = self.stack.pop().unwrap();

            match state {
                ElementState::Start => {
                    // Emit start tag or empty tag
                    self.buffer.push(b'<');
                    self.buffer.extend_from_slice(elem.name.as_bytes());

                    for (key, value) in &elem.attributes {
                        self.buffer.push(b' ');
                        self.buffer.extend_from_slice(key.as_bytes());
                        self.buffer.extend_from_slice(b"=\"");
                        // Escape attribute value
                        for &b in value.as_bytes() {
                            match b {
                                b'"' => self.buffer.extend_from_slice(b"&quot;"),
                                b'&' => self.buffer.extend_from_slice(b"&amp;"),
                                b'<' => self.buffer.extend_from_slice(b"&lt;"),
                                _ => self.buffer.push(b),
                            }
                        }
                        self.buffer.push(b'"');
                    }

                    if elem.self_closing && elem.children.is_empty() {
                        self.buffer.extend_from_slice(b"/>");
                        // Done with this element
                    } else {
                        self.buffer.push(b'>');
                        // Push back to process children
                        self.stack.push((elem, 0, ElementState::Children));
                    }
                }
                ElementState::Children => {
                    if child_idx < elem.children.len() {
                        // Push back with next child index
                        self.stack
                            .push((elem, child_idx + 1, ElementState::Children));

                        // Process current child
                        match &elem.children[child_idx] {
                            RawXmlNode::Element(child) => {
                                self.stack.push((child, 0, ElementState::Start));
                            }
                            RawXmlNode::Text(text) => {
                                // Escape text content
                                for &b in text.as_bytes() {
                                    match b {
                                        b'&' => self.buffer.extend_from_slice(b"&amp;"),
                                        b'<' => self.buffer.extend_from_slice(b"&lt;"),
                                        b'>' => self.buffer.extend_from_slice(b"&gt;"),
                                        _ => self.buffer.push(b),
                                    }
                                }
                            }
                            RawXmlNode::CData(text) => {
                                self.buffer.extend_from_slice(b"<![CDATA[");
                                self.buffer.extend_from_slice(text.as_bytes());
                                self.buffer.extend_from_slice(b"]]>");
                            }
                            RawXmlNode::Comment(text) => {
                                self.buffer.extend_from_slice(b"<!--");
                                self.buffer.extend_from_slice(text.as_bytes());
                                self.buffer.extend_from_slice(b"-->");
                            }
                        }
                    } else {
                        // Done with children, emit end tag
                        self.stack.push((elem, 0, ElementState::End));
                    }
                }
                ElementState::End => {
                    self.buffer.extend_from_slice(b"</");
                    self.buffer.extend_from_slice(elem.name.as_bytes());
                    self.buffer.push(b'>');
                }
            }
        }

        if self.stack.is_empty() && self.buffer.is_empty() {
            self.done = true;
        }
    }
}

impl<'a> std::io::Read for RawXmlStreamReader<'a> {
    fn read(&mut self, buf: &mut [u8]) -> std::io::Result<usize> {
        if self.pos >= self.buffer.len() {
            if self.done {
                return Ok(0);
            }
            self.generate_next();
            if self.done && self.buffer.is_empty() {
                return Ok(0);
            }
        }

        let remaining = &self.buffer[self.pos..];
        let to_copy = remaining.len().min(buf.len());
        buf[..to_copy].copy_from_slice(&remaining[..to_copy]);
        self.pos += to_copy;
        Ok(to_copy)
    }
}

impl<'a> BufRead for RawXmlStreamReader<'a> {
    fn fill_buf(&mut self) -> std::io::Result<&[u8]> {
        if self.pos >= self.buffer.len() {
            if self.done {
                return Ok(&[]);
            }
            self.generate_next();
        }
        Ok(&self.buffer[self.pos..])
    }

    fn consume(&mut self, amt: usize) {
        self.pos += amt;
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

    #[test]
    fn test_streaming_reader() {
        use std::io::Read;

        let xml = r#"<parent attr="val"><child>text</child></parent>"#;
        let mut reader = Reader::from_str(xml);
        let mut buf = Vec::new();

        if let Ok(Event::Start(e)) = reader.read_event_into(&mut buf) {
            let elem = RawXmlElement::from_reader(&mut reader, &e).unwrap();

            // Read from streaming reader
            let mut stream_reader = RawXmlStreamReader::new(&elem);
            let mut output = String::new();
            stream_reader.read_to_string(&mut output).unwrap();

            assert_eq!(output, xml);
        }
    }

    #[test]
    fn test_streaming_reader_escaping() {
        use std::io::Read;

        // Test that special characters are properly escaped
        let mut elem = RawXmlElement::new("test");
        elem.attributes
            .push(("attr".to_string(), "val\"ue".to_string()));
        elem.children
            .push(RawXmlNode::Text("a < b & c > d".to_string()));

        let mut stream_reader = RawXmlStreamReader::new(&elem);
        let mut output = String::new();
        stream_reader.read_to_string(&mut output).unwrap();

        assert_eq!(
            output,
            r#"<test attr="val&quot;ue">a &lt; b &amp; c &gt; d</test>"#
        );
    }
}
