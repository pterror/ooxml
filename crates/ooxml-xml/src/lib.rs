//! Shared XML utilities for the ooxml library.
//!
//! This crate provides common XML types and utilities used across OOXML format crates
//! for features like roundtrip fidelity (preserving unknown elements/attributes).

use quick_xml::events::{BytesEnd, BytesStart};
use quick_xml::{Reader, Writer};
use std::io::{BufRead, Write};

mod raw_xml;
pub mod serde_helpers;

pub use raw_xml::{PositionedAttr, PositionedNode, RawXmlElement, RawXmlNode};
pub use serde_helpers::{ooxml_bool, ooxml_bool_required};

/// Error type for XML operations.
#[derive(Debug, thiserror::Error)]
pub enum Error {
    #[error("XML error: {0}")]
    Xml(#[from] quick_xml::Error),
    #[error("IO error: {0}")]
    Io(#[from] std::io::Error),
    #[error("invalid XML: {0}")]
    Invalid(String),
}

/// Result type for XML operations.
pub type Result<T> = std::result::Result<T, Error>;

/// Error type for XML parsing (used by generated parsers).
#[derive(Debug, thiserror::Error)]
pub enum ParseError {
    #[error("XML error: {0}")]
    Xml(#[from] quick_xml::Error),
    #[error("raw XML error: {0}")]
    RawXml(#[from] Error),
    #[error("unexpected element: {0}")]
    UnexpectedElement(String),
    #[error("missing attribute: {0}")]
    MissingAttribute(String),
    #[error("invalid value: {0}")]
    InvalidValue(String),
}

/// Error type for XML serialization (used by generated serializers).
#[derive(Debug, thiserror::Error)]
pub enum SerializeError {
    #[error("XML error: {0}")]
    Xml(#[from] quick_xml::Error),
    #[error("IO error: {0}")]
    Io(#[from] std::io::Error),
    #[error("raw XML error: {0}")]
    RawXml(#[from] Error),
}

/// Trait for types that can be parsed from XML elements.
///
/// Implemented by generated OOXML types to enable event-based parsing.
pub trait FromXml: Sized {
    /// Parse from an XML reader positioned at the start tag.
    ///
    /// - `reader`: The XML reader
    /// - `start_tag`: The start tag that was just read
    /// - `is_empty`: True if this is an empty element (e.g., `<foo/>`)
    fn from_xml<R: BufRead>(
        reader: &mut Reader<R>,
        start_tag: &BytesStart,
        is_empty: bool,
    ) -> std::result::Result<Self, ParseError>;
}

/// Trait for types that can be serialized to XML elements.
///
/// Implemented by generated OOXML types to enable roundtrip serialization.
pub trait ToXml {
    /// Write attributes onto the start tag and return it.
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        start
    }

    /// Write child elements and text content inside the element.
    fn write_children<W: Write>(
        &self,
        _writer: &mut Writer<W>,
    ) -> std::result::Result<(), SerializeError> {
        Ok(())
    }

    /// Whether this element has no children (self-closing).
    fn is_empty_element(&self) -> bool {
        false
    }

    /// Write a complete element: `<tag attrs>children</tag>` or `<tag attrs/>`.
    fn write_element<W: Write>(
        &self,
        tag: &str,
        writer: &mut Writer<W>,
    ) -> std::result::Result<(), SerializeError> {
        use quick_xml::events::Event;
        let start = BytesStart::new(tag);
        let start = self.write_attrs(start);
        if self.is_empty_element() {
            writer.write_event(Event::Empty(start))?;
        } else {
            writer.write_event(Event::Start(start))?;
            self.write_children(writer)?;
            writer.write_event(Event::End(BytesEnd::new(tag)))?;
        }
        Ok(())
    }
}

// Blanket implementation for Box<T> where T: ToXml
impl<T: ToXml> ToXml for Box<T> {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        (**self).write_attrs(start)
    }

    fn write_children<W: Write>(
        &self,
        writer: &mut Writer<W>,
    ) -> std::result::Result<(), SerializeError> {
        (**self).write_children(writer)
    }

    fn is_empty_element(&self) -> bool {
        (**self).is_empty_element()
    }

    fn write_element<W: Write>(
        &self,
        tag: &str,
        writer: &mut Writer<W>,
    ) -> std::result::Result<(), SerializeError> {
        (**self).write_element(tag, writer)
    }
}
