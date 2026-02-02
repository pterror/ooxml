// ToXml serializers for generated types.
// Enables roundtrip XML serialization alongside FromXml parsers.

#![allow(unused_variables)]
#![allow(clippy::single_match)]
#![allow(clippy::match_single_binding)]

use super::generated::*;
use quick_xml::Writer;
use quick_xml::events::{BytesEnd, BytesStart, BytesText, Event};
use std::io::Write;

/// Error type for XML serialization.
#[derive(Debug)]
pub enum SerializeError {
    Xml(quick_xml::Error),
    Io(std::io::Error),
    #[cfg(feature = "extra-children")]
    RawXml(ooxml_xml::Error),
}

impl From<quick_xml::Error> for SerializeError {
    fn from(e: quick_xml::Error) -> Self {
        SerializeError::Xml(e)
    }
}

impl From<std::io::Error> for SerializeError {
    fn from(e: std::io::Error) -> Self {
        SerializeError::Io(e)
    }
}

#[cfg(feature = "extra-children")]
impl From<ooxml_xml::Error> for SerializeError {
    fn from(e: ooxml_xml::Error) -> Self {
        SerializeError::RawXml(e)
    }
}

impl std::fmt::Display for SerializeError {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::Xml(e) => write!(f, "XML error: {}", e),
            Self::Io(e) => write!(f, "IO error: {}", e),
            #[cfg(feature = "extra-children")]
            Self::RawXml(e) => write!(f, "RawXml error: {}", e),
        }
    }
}

impl std::error::Error for SerializeError {}

/// Trait for types that can be serialized to XML events.
pub trait ToXml {
    /// Write attributes onto the start tag and return it.
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        start
    }

    /// Write child elements and text content inside the element.
    fn write_children<W: Write>(&self, writer: &mut Writer<W>) -> Result<(), SerializeError> {
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
    ) -> Result<(), SerializeError> {
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

impl<T: ToXml> ToXml for Box<T> {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        (**self).write_attrs(start)
    }
    fn write_children<W: Write>(&self, writer: &mut Writer<W>) -> Result<(), SerializeError> {
        (**self).write_children(writer)
    }
    fn is_empty_element(&self) -> bool {
        (**self).is_empty_element()
    }
    fn write_element<W: Write>(
        &self,
        tag: &str,
        writer: &mut Writer<W>,
    ) -> Result<(), SerializeError> {
        (**self).write_element(tag, writer)
    }
}

#[allow(dead_code)]
/// Encode bytes as a hex string.
fn encode_hex(bytes: &[u8]) -> String {
    bytes.iter().map(|b| format!("{:02X}", b)).collect()
}

impl ToXml for CTEmpty {
    fn is_empty_element(&self) -> bool {
        true
    }
}

impl ToXml for CTOnOff {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        if let Some(ref val) = self.value {
            {
                let s = val.to_string();
                start.push_attribute(("w:val", s.as_str()));
            }
        }
        #[cfg(feature = "extra-attrs")]
        for (key, value) in &self.extra_attrs {
            start.push_attribute((key.as_str(), value.as_str()));
        }
        start
    }

    fn is_empty_element(&self) -> bool {
        true
    }
}

impl ToXml for CTLongHexNumber {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        {
            let val = &self.value;
            {
                let hex = encode_hex(val);
                start.push_attribute(("w:val", hex.as_str()));
            }
        }
        #[cfg(feature = "extra-attrs")]
        for (key, value) in &self.extra_attrs {
            start.push_attribute((key.as_str(), value.as_str()));
        }
        start
    }

    fn is_empty_element(&self) -> bool {
        true
    }
}

impl ToXml for CTCharset {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        if let Some(ref val) = self.value {
            {
                let hex = encode_hex(val);
                start.push_attribute(("w:val", hex.as_str()));
            }
        }
        if let Some(ref val) = self.character_set {
            start.push_attribute(("w:characterSet", val.as_str()));
        }
        #[cfg(feature = "extra-attrs")]
        for (key, value) in &self.extra_attrs {
            start.push_attribute((key.as_str(), value.as_str()));
        }
        start
    }

    fn is_empty_element(&self) -> bool {
        true
    }
}

impl ToXml for CTDecimalNumber {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        {
            let val = &self.value;
            {
                let s = val.to_string();
                start.push_attribute(("w:val", s.as_str()));
            }
        }
        #[cfg(feature = "extra-attrs")]
        for (key, value) in &self.extra_attrs {
            start.push_attribute((key.as_str(), value.as_str()));
        }
        start
    }

    fn is_empty_element(&self) -> bool {
        true
    }
}

impl ToXml for CTUnsignedDecimalNumber {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        {
            let val = &self.value;
            {
                let s = val.to_string();
                start.push_attribute(("w:val", s.as_str()));
            }
        }
        #[cfg(feature = "extra-attrs")]
        for (key, value) in &self.extra_attrs {
            start.push_attribute((key.as_str(), value.as_str()));
        }
        start
    }

    fn is_empty_element(&self) -> bool {
        true
    }
}

impl ToXml for CTDecimalNumberOrPrecent {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        {
            let val = &self.value;
            {
                let s = val.to_string();
                start.push_attribute(("w:val", s.as_str()));
            }
        }
        #[cfg(feature = "extra-attrs")]
        for (key, value) in &self.extra_attrs {
            start.push_attribute((key.as_str(), value.as_str()));
        }
        start
    }

    fn is_empty_element(&self) -> bool {
        true
    }
}

impl ToXml for CTTwipsMeasure {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        {
            let val = &self.value;
            {
                let s = val.to_string();
                start.push_attribute(("w:val", s.as_str()));
            }
        }
        #[cfg(feature = "extra-attrs")]
        for (key, value) in &self.extra_attrs {
            start.push_attribute((key.as_str(), value.as_str()));
        }
        start
    }

    fn is_empty_element(&self) -> bool {
        true
    }
}

impl ToXml for CTSignedTwipsMeasure {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        {
            let val = &self.value;
            {
                let s = val.to_string();
                start.push_attribute(("w:val", s.as_str()));
            }
        }
        #[cfg(feature = "extra-attrs")]
        for (key, value) in &self.extra_attrs {
            start.push_attribute((key.as_str(), value.as_str()));
        }
        start
    }

    fn is_empty_element(&self) -> bool {
        true
    }
}

impl ToXml for CTPixelsMeasure {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        {
            let val = &self.value;
            {
                let s = val.to_string();
                start.push_attribute(("w:val", s.as_str()));
            }
        }
        #[cfg(feature = "extra-attrs")]
        for (key, value) in &self.extra_attrs {
            start.push_attribute((key.as_str(), value.as_str()));
        }
        start
    }

    fn is_empty_element(&self) -> bool {
        true
    }
}

impl ToXml for CTHpsMeasure {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        {
            let val = &self.value;
            {
                let s = val.to_string();
                start.push_attribute(("w:val", s.as_str()));
            }
        }
        #[cfg(feature = "extra-attrs")]
        for (key, value) in &self.extra_attrs {
            start.push_attribute((key.as_str(), value.as_str()));
        }
        start
    }

    fn is_empty_element(&self) -> bool {
        true
    }
}

impl ToXml for CTSignedHpsMeasure {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        {
            let val = &self.value;
            {
                let s = val.to_string();
                start.push_attribute(("w:val", s.as_str()));
            }
        }
        #[cfg(feature = "extra-attrs")]
        for (key, value) in &self.extra_attrs {
            start.push_attribute((key.as_str(), value.as_str()));
        }
        start
    }

    fn is_empty_element(&self) -> bool {
        true
    }
}

impl ToXml for CTMacroName {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        {
            let val = &self.value;
            start.push_attribute(("w:val", val.as_str()));
        }
        #[cfg(feature = "extra-attrs")]
        for (key, value) in &self.extra_attrs {
            start.push_attribute((key.as_str(), value.as_str()));
        }
        start
    }

    fn is_empty_element(&self) -> bool {
        true
    }
}

impl ToXml for CTString {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        {
            let val = &self.value;
            start.push_attribute(("w:val", val.as_str()));
        }
        #[cfg(feature = "extra-attrs")]
        for (key, value) in &self.extra_attrs {
            start.push_attribute((key.as_str(), value.as_str()));
        }
        start
    }

    fn is_empty_element(&self) -> bool {
        true
    }
}

impl ToXml for CTTextScale {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        if let Some(ref val) = self.value {
            {
                let s = val.to_string();
                start.push_attribute(("w:val", s.as_str()));
            }
        }
        #[cfg(feature = "extra-attrs")]
        for (key, value) in &self.extra_attrs {
            start.push_attribute((key.as_str(), value.as_str()));
        }
        start
    }

    fn is_empty_element(&self) -> bool {
        true
    }
}

impl ToXml for CTHighlight {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        {
            let val = &self.value;
            {
                let s = val.to_string();
                start.push_attribute(("w:val", s.as_str()));
            }
        }
        #[cfg(feature = "extra-attrs")]
        for (key, value) in &self.extra_attrs {
            start.push_attribute((key.as_str(), value.as_str()));
        }
        start
    }

    fn is_empty_element(&self) -> bool {
        true
    }
}

impl ToXml for CTColor {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        {
            let val = &self.value;
            {
                let s = val.to_string();
                start.push_attribute(("w:val", s.as_str()));
            }
        }
        if let Some(ref val) = self.theme_color {
            {
                let s = val.to_string();
                start.push_attribute(("w:themeColor", s.as_str()));
            }
        }
        if let Some(ref val) = self.theme_tint {
            {
                let hex = encode_hex(val);
                start.push_attribute(("w:themeTint", hex.as_str()));
            }
        }
        if let Some(ref val) = self.theme_shade {
            {
                let hex = encode_hex(val);
                start.push_attribute(("w:themeShade", hex.as_str()));
            }
        }
        #[cfg(feature = "extra-attrs")]
        for (key, value) in &self.extra_attrs {
            start.push_attribute((key.as_str(), value.as_str()));
        }
        start
    }

    fn is_empty_element(&self) -> bool {
        true
    }
}

impl ToXml for CTLang {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        {
            let val = &self.value;
            start.push_attribute(("w:val", val.as_str()));
        }
        #[cfg(feature = "extra-attrs")]
        for (key, value) in &self.extra_attrs {
            start.push_attribute((key.as_str(), value.as_str()));
        }
        start
    }

    fn is_empty_element(&self) -> bool {
        true
    }
}

impl ToXml for CTGuid {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        if let Some(ref val) = self.value {
            start.push_attribute(("w:val", val.as_str()));
        }
        #[cfg(feature = "extra-attrs")]
        for (key, value) in &self.extra_attrs {
            start.push_attribute((key.as_str(), value.as_str()));
        }
        start
    }

    fn is_empty_element(&self) -> bool {
        true
    }
}

impl ToXml for CTUnderline {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        if let Some(ref val) = self.value {
            {
                let s = val.to_string();
                start.push_attribute(("w:val", s.as_str()));
            }
        }
        if let Some(ref val) = self.color {
            {
                let s = val.to_string();
                start.push_attribute(("w:color", s.as_str()));
            }
        }
        if let Some(ref val) = self.theme_color {
            {
                let s = val.to_string();
                start.push_attribute(("w:themeColor", s.as_str()));
            }
        }
        if let Some(ref val) = self.theme_tint {
            {
                let hex = encode_hex(val);
                start.push_attribute(("w:themeTint", hex.as_str()));
            }
        }
        if let Some(ref val) = self.theme_shade {
            {
                let hex = encode_hex(val);
                start.push_attribute(("w:themeShade", hex.as_str()));
            }
        }
        #[cfg(feature = "extra-attrs")]
        for (key, value) in &self.extra_attrs {
            start.push_attribute((key.as_str(), value.as_str()));
        }
        start
    }

    fn is_empty_element(&self) -> bool {
        true
    }
}

impl ToXml for CTTextEffect {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        {
            let val = &self.value;
            {
                let s = val.to_string();
                start.push_attribute(("w:val", s.as_str()));
            }
        }
        #[cfg(feature = "extra-attrs")]
        for (key, value) in &self.extra_attrs {
            start.push_attribute((key.as_str(), value.as_str()));
        }
        start
    }

    fn is_empty_element(&self) -> bool {
        true
    }
}

impl ToXml for CTBorder {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        {
            let val = &self.value;
            {
                let s = val.to_string();
                start.push_attribute(("w:val", s.as_str()));
            }
        }
        if let Some(ref val) = self.color {
            {
                let s = val.to_string();
                start.push_attribute(("w:color", s.as_str()));
            }
        }
        if let Some(ref val) = self.theme_color {
            {
                let s = val.to_string();
                start.push_attribute(("w:themeColor", s.as_str()));
            }
        }
        if let Some(ref val) = self.theme_tint {
            {
                let hex = encode_hex(val);
                start.push_attribute(("w:themeTint", hex.as_str()));
            }
        }
        if let Some(ref val) = self.theme_shade {
            {
                let hex = encode_hex(val);
                start.push_attribute(("w:themeShade", hex.as_str()));
            }
        }
        if let Some(ref val) = self.size {
            {
                let s = val.to_string();
                start.push_attribute(("w:sz", s.as_str()));
            }
        }
        if let Some(ref val) = self.space {
            {
                let s = val.to_string();
                start.push_attribute(("w:space", s.as_str()));
            }
        }
        if let Some(ref val) = self.shadow {
            {
                let s = val.to_string();
                start.push_attribute(("w:shadow", s.as_str()));
            }
        }
        if let Some(ref val) = self.frame {
            {
                let s = val.to_string();
                start.push_attribute(("w:frame", s.as_str()));
            }
        }
        #[cfg(feature = "extra-attrs")]
        for (key, value) in &self.extra_attrs {
            start.push_attribute((key.as_str(), value.as_str()));
        }
        start
    }

    fn is_empty_element(&self) -> bool {
        true
    }
}

impl ToXml for CTShd {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        {
            let val = &self.value;
            {
                let s = val.to_string();
                start.push_attribute(("w:val", s.as_str()));
            }
        }
        if let Some(ref val) = self.color {
            {
                let s = val.to_string();
                start.push_attribute(("w:color", s.as_str()));
            }
        }
        if let Some(ref val) = self.theme_color {
            {
                let s = val.to_string();
                start.push_attribute(("w:themeColor", s.as_str()));
            }
        }
        if let Some(ref val) = self.theme_tint {
            {
                let hex = encode_hex(val);
                start.push_attribute(("w:themeTint", hex.as_str()));
            }
        }
        if let Some(ref val) = self.theme_shade {
            {
                let hex = encode_hex(val);
                start.push_attribute(("w:themeShade", hex.as_str()));
            }
        }
        if let Some(ref val) = self.fill {
            {
                let s = val.to_string();
                start.push_attribute(("w:fill", s.as_str()));
            }
        }
        if let Some(ref val) = self.theme_fill {
            {
                let s = val.to_string();
                start.push_attribute(("w:themeFill", s.as_str()));
            }
        }
        if let Some(ref val) = self.theme_fill_tint {
            {
                let hex = encode_hex(val);
                start.push_attribute(("w:themeFillTint", hex.as_str()));
            }
        }
        if let Some(ref val) = self.theme_fill_shade {
            {
                let hex = encode_hex(val);
                start.push_attribute(("w:themeFillShade", hex.as_str()));
            }
        }
        #[cfg(feature = "extra-attrs")]
        for (key, value) in &self.extra_attrs {
            start.push_attribute((key.as_str(), value.as_str()));
        }
        start
    }

    fn is_empty_element(&self) -> bool {
        true
    }
}

impl ToXml for CTVerticalAlignRun {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        {
            let val = &self.value;
            {
                let s = val.to_string();
                start.push_attribute(("w:val", s.as_str()));
            }
        }
        #[cfg(feature = "extra-attrs")]
        for (key, value) in &self.extra_attrs {
            start.push_attribute((key.as_str(), value.as_str()));
        }
        start
    }

    fn is_empty_element(&self) -> bool {
        true
    }
}

impl ToXml for CTFitText {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        {
            let val = &self.value;
            {
                let s = val.to_string();
                start.push_attribute(("w:val", s.as_str()));
            }
        }
        if let Some(ref val) = self.id {
            {
                let s = val.to_string();
                start.push_attribute(("w:id", s.as_str()));
            }
        }
        #[cfg(feature = "extra-attrs")]
        for (key, value) in &self.extra_attrs {
            start.push_attribute((key.as_str(), value.as_str()));
        }
        start
    }

    fn is_empty_element(&self) -> bool {
        true
    }
}

impl ToXml for CTEm {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        {
            let val = &self.value;
            {
                let s = val.to_string();
                start.push_attribute(("w:val", s.as_str()));
            }
        }
        #[cfg(feature = "extra-attrs")]
        for (key, value) in &self.extra_attrs {
            start.push_attribute((key.as_str(), value.as_str()));
        }
        start
    }

    fn is_empty_element(&self) -> bool {
        true
    }
}

impl ToXml for CTLanguage {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        if let Some(ref val) = self.value {
            start.push_attribute(("w:val", val.as_str()));
        }
        if let Some(ref val) = self.east_asia {
            start.push_attribute(("w:eastAsia", val.as_str()));
        }
        if let Some(ref val) = self.bidi {
            start.push_attribute(("w:bidi", val.as_str()));
        }
        #[cfg(feature = "extra-attrs")]
        for (key, value) in &self.extra_attrs {
            start.push_attribute((key.as_str(), value.as_str()));
        }
        start
    }

    fn is_empty_element(&self) -> bool {
        true
    }
}

impl ToXml for CTEastAsianLayout {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        if let Some(ref val) = self.id {
            {
                let s = val.to_string();
                start.push_attribute(("w:id", s.as_str()));
            }
        }
        if let Some(ref val) = self.combine {
            {
                let s = val.to_string();
                start.push_attribute(("w:combine", s.as_str()));
            }
        }
        if let Some(ref val) = self.combine_brackets {
            {
                let s = val.to_string();
                start.push_attribute(("w:combineBrackets", s.as_str()));
            }
        }
        if let Some(ref val) = self.vert {
            {
                let s = val.to_string();
                start.push_attribute(("w:vert", s.as_str()));
            }
        }
        if let Some(ref val) = self.vert_compress {
            {
                let s = val.to_string();
                start.push_attribute(("w:vertCompress", s.as_str()));
            }
        }
        #[cfg(feature = "extra-attrs")]
        for (key, value) in &self.extra_attrs {
            start.push_attribute((key.as_str(), value.as_str()));
        }
        start
    }

    fn is_empty_element(&self) -> bool {
        true
    }
}

impl ToXml for CTFramePr {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        if let Some(ref val) = self.drop_cap {
            {
                let s = val.to_string();
                start.push_attribute(("w:dropCap", s.as_str()));
            }
        }
        if let Some(ref val) = self.lines {
            {
                let s = val.to_string();
                start.push_attribute(("w:lines", s.as_str()));
            }
        }
        if let Some(ref val) = self.width {
            {
                let s = val.to_string();
                start.push_attribute(("w:w", s.as_str()));
            }
        }
        if let Some(ref val) = self.height {
            {
                let s = val.to_string();
                start.push_attribute(("w:h", s.as_str()));
            }
        }
        if let Some(ref val) = self.v_space {
            {
                let s = val.to_string();
                start.push_attribute(("w:vSpace", s.as_str()));
            }
        }
        if let Some(ref val) = self.h_space {
            {
                let s = val.to_string();
                start.push_attribute(("w:hSpace", s.as_str()));
            }
        }
        if let Some(ref val) = self.wrap {
            {
                let s = val.to_string();
                start.push_attribute(("w:wrap", s.as_str()));
            }
        }
        if let Some(ref val) = self.h_anchor {
            {
                let s = val.to_string();
                start.push_attribute(("w:hAnchor", s.as_str()));
            }
        }
        if let Some(ref val) = self.v_anchor {
            {
                let s = val.to_string();
                start.push_attribute(("w:vAnchor", s.as_str()));
            }
        }
        if let Some(ref val) = self.x {
            {
                let s = val.to_string();
                start.push_attribute(("w:x", s.as_str()));
            }
        }
        if let Some(ref val) = self.x_align {
            {
                let s = val.to_string();
                start.push_attribute(("w:xAlign", s.as_str()));
            }
        }
        if let Some(ref val) = self.y {
            {
                let s = val.to_string();
                start.push_attribute(("w:y", s.as_str()));
            }
        }
        if let Some(ref val) = self.y_align {
            {
                let s = val.to_string();
                start.push_attribute(("w:yAlign", s.as_str()));
            }
        }
        if let Some(ref val) = self.h_rule {
            {
                let s = val.to_string();
                start.push_attribute(("w:hRule", s.as_str()));
            }
        }
        if let Some(ref val) = self.anchor_lock {
            {
                let s = val.to_string();
                start.push_attribute(("w:anchorLock", s.as_str()));
            }
        }
        #[cfg(feature = "extra-attrs")]
        for (key, value) in &self.extra_attrs {
            start.push_attribute((key.as_str(), value.as_str()));
        }
        start
    }

    fn is_empty_element(&self) -> bool {
        true
    }
}

impl ToXml for CTTabStop {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        {
            let val = &self.value;
            {
                let s = val.to_string();
                start.push_attribute(("w:val", s.as_str()));
            }
        }
        if let Some(ref val) = self.leader {
            {
                let s = val.to_string();
                start.push_attribute(("w:leader", s.as_str()));
            }
        }
        {
            let val = &self.pos;
            {
                let s = val.to_string();
                start.push_attribute(("w:pos", s.as_str()));
            }
        }
        #[cfg(feature = "extra-attrs")]
        for (key, value) in &self.extra_attrs {
            start.push_attribute((key.as_str(), value.as_str()));
        }
        start
    }

    fn is_empty_element(&self) -> bool {
        true
    }
}

impl ToXml for CTSpacing {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        if let Some(ref val) = self.before {
            {
                let s = val.to_string();
                start.push_attribute(("w:before", s.as_str()));
            }
        }
        if let Some(ref val) = self.before_lines {
            {
                let s = val.to_string();
                start.push_attribute(("w:beforeLines", s.as_str()));
            }
        }
        if let Some(ref val) = self.before_autospacing {
            {
                let s = val.to_string();
                start.push_attribute(("w:beforeAutospacing", s.as_str()));
            }
        }
        if let Some(ref val) = self.after {
            {
                let s = val.to_string();
                start.push_attribute(("w:after", s.as_str()));
            }
        }
        if let Some(ref val) = self.after_lines {
            {
                let s = val.to_string();
                start.push_attribute(("w:afterLines", s.as_str()));
            }
        }
        if let Some(ref val) = self.after_autospacing {
            {
                let s = val.to_string();
                start.push_attribute(("w:afterAutospacing", s.as_str()));
            }
        }
        if let Some(ref val) = self.line {
            {
                let s = val.to_string();
                start.push_attribute(("w:line", s.as_str()));
            }
        }
        if let Some(ref val) = self.line_rule {
            {
                let s = val.to_string();
                start.push_attribute(("w:lineRule", s.as_str()));
            }
        }
        #[cfg(feature = "extra-attrs")]
        for (key, value) in &self.extra_attrs {
            start.push_attribute((key.as_str(), value.as_str()));
        }
        start
    }

    fn is_empty_element(&self) -> bool {
        true
    }
}

impl ToXml for CTInd {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        if let Some(ref val) = self.start {
            {
                let s = val.to_string();
                start.push_attribute(("w:start", s.as_str()));
            }
        }
        if let Some(ref val) = self.start_chars {
            {
                let s = val.to_string();
                start.push_attribute(("w:startChars", s.as_str()));
            }
        }
        if let Some(ref val) = self.end {
            {
                let s = val.to_string();
                start.push_attribute(("w:end", s.as_str()));
            }
        }
        if let Some(ref val) = self.end_chars {
            {
                let s = val.to_string();
                start.push_attribute(("w:endChars", s.as_str()));
            }
        }
        if let Some(ref val) = self.left {
            {
                let s = val.to_string();
                start.push_attribute(("w:left", s.as_str()));
            }
        }
        if let Some(ref val) = self.left_chars {
            {
                let s = val.to_string();
                start.push_attribute(("w:leftChars", s.as_str()));
            }
        }
        if let Some(ref val) = self.right {
            {
                let s = val.to_string();
                start.push_attribute(("w:right", s.as_str()));
            }
        }
        if let Some(ref val) = self.right_chars {
            {
                let s = val.to_string();
                start.push_attribute(("w:rightChars", s.as_str()));
            }
        }
        if let Some(ref val) = self.hanging {
            {
                let s = val.to_string();
                start.push_attribute(("w:hanging", s.as_str()));
            }
        }
        if let Some(ref val) = self.hanging_chars {
            {
                let s = val.to_string();
                start.push_attribute(("w:hangingChars", s.as_str()));
            }
        }
        if let Some(ref val) = self.first_line {
            {
                let s = val.to_string();
                start.push_attribute(("w:firstLine", s.as_str()));
            }
        }
        if let Some(ref val) = self.first_line_chars {
            {
                let s = val.to_string();
                start.push_attribute(("w:firstLineChars", s.as_str()));
            }
        }
        #[cfg(feature = "extra-attrs")]
        for (key, value) in &self.extra_attrs {
            start.push_attribute((key.as_str(), value.as_str()));
        }
        start
    }

    fn is_empty_element(&self) -> bool {
        true
    }
}

impl ToXml for CTJc {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        {
            let val = &self.value;
            {
                let s = val.to_string();
                start.push_attribute(("w:val", s.as_str()));
            }
        }
        #[cfg(feature = "extra-attrs")]
        for (key, value) in &self.extra_attrs {
            start.push_attribute((key.as_str(), value.as_str()));
        }
        start
    }

    fn is_empty_element(&self) -> bool {
        true
    }
}

impl ToXml for CTJcTable {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        {
            let val = &self.value;
            {
                let s = val.to_string();
                start.push_attribute(("w:val", s.as_str()));
            }
        }
        #[cfg(feature = "extra-attrs")]
        for (key, value) in &self.extra_attrs {
            start.push_attribute((key.as_str(), value.as_str()));
        }
        start
    }

    fn is_empty_element(&self) -> bool {
        true
    }
}

impl ToXml for CTView {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        {
            let val = &self.value;
            {
                let s = val.to_string();
                start.push_attribute(("w:val", s.as_str()));
            }
        }
        #[cfg(feature = "extra-attrs")]
        for (key, value) in &self.extra_attrs {
            start.push_attribute((key.as_str(), value.as_str()));
        }
        start
    }

    fn is_empty_element(&self) -> bool {
        true
    }
}

impl ToXml for CTZoom {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        if let Some(ref val) = self.value {
            {
                let s = val.to_string();
                start.push_attribute(("w:val", s.as_str()));
            }
        }
        {
            let val = &self.percent;
            {
                let s = val.to_string();
                start.push_attribute(("w:percent", s.as_str()));
            }
        }
        #[cfg(feature = "extra-attrs")]
        for (key, value) in &self.extra_attrs {
            start.push_attribute((key.as_str(), value.as_str()));
        }
        start
    }

    fn is_empty_element(&self) -> bool {
        true
    }
}

impl ToXml for CTWritingStyle {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        {
            let val = &self.lang;
            start.push_attribute(("w:lang", val.as_str()));
        }
        {
            let val = &self.vendor_i_d;
            start.push_attribute(("w:vendorID", val.as_str()));
        }
        {
            let val = &self.dll_version;
            start.push_attribute(("w:dllVersion", val.as_str()));
        }
        if let Some(ref val) = self.nl_check {
            {
                let s = val.to_string();
                start.push_attribute(("w:nlCheck", s.as_str()));
            }
        }
        {
            let val = &self.check_style;
            {
                let s = val.to_string();
                start.push_attribute(("w:checkStyle", s.as_str()));
            }
        }
        {
            let val = &self.app_name;
            start.push_attribute(("w:appName", val.as_str()));
        }
        #[cfg(feature = "extra-attrs")]
        for (key, value) in &self.extra_attrs {
            start.push_attribute((key.as_str(), value.as_str()));
        }
        start
    }

    fn is_empty_element(&self) -> bool {
        true
    }
}

impl ToXml for CTProof {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        if let Some(ref val) = self.spelling {
            {
                let s = val.to_string();
                start.push_attribute(("w:spelling", s.as_str()));
            }
        }
        if let Some(ref val) = self.grammar {
            {
                let s = val.to_string();
                start.push_attribute(("w:grammar", s.as_str()));
            }
        }
        #[cfg(feature = "extra-attrs")]
        for (key, value) in &self.extra_attrs {
            start.push_attribute((key.as_str(), value.as_str()));
        }
        start
    }

    fn is_empty_element(&self) -> bool {
        true
    }
}

impl ToXml for CTDocType {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        {
            let val = &self.value;
            start.push_attribute(("w:val", val.as_str()));
        }
        #[cfg(feature = "extra-attrs")]
        for (key, value) in &self.extra_attrs {
            start.push_attribute((key.as_str(), value.as_str()));
        }
        start
    }

    fn is_empty_element(&self) -> bool {
        true
    }
}

impl ToXml for WAGPassword {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        if let Some(ref val) = self.algorithm_name {
            start.push_attribute(("w:algorithmName", val.as_str()));
        }
        if let Some(ref val) = self.hash_value {
            {
                let hex = encode_hex(val);
                start.push_attribute(("w:hashValue", hex.as_str()));
            }
        }
        if let Some(ref val) = self.salt_value {
            {
                let hex = encode_hex(val);
                start.push_attribute(("w:saltValue", hex.as_str()));
            }
        }
        if let Some(ref val) = self.spin_count {
            {
                let s = val.to_string();
                start.push_attribute(("w:spinCount", s.as_str()));
            }
        }
        #[cfg(feature = "extra-attrs")]
        for (key, value) in &self.extra_attrs {
            start.push_attribute((key.as_str(), value.as_str()));
        }
        start
    }

    fn is_empty_element(&self) -> bool {
        true
    }
}

impl ToXml for WAGTransitionalPassword {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        if let Some(ref val) = self.crypt_provider_type {
            {
                let s = val.to_string();
                start.push_attribute(("w:cryptProviderType", s.as_str()));
            }
        }
        if let Some(ref val) = self.crypt_algorithm_class {
            {
                let s = val.to_string();
                start.push_attribute(("w:cryptAlgorithmClass", s.as_str()));
            }
        }
        if let Some(ref val) = self.crypt_algorithm_type {
            {
                let s = val.to_string();
                start.push_attribute(("w:cryptAlgorithmType", s.as_str()));
            }
        }
        if let Some(ref val) = self.crypt_algorithm_sid {
            {
                let s = val.to_string();
                start.push_attribute(("w:cryptAlgorithmSid", s.as_str()));
            }
        }
        if let Some(ref val) = self.crypt_spin_count {
            {
                let s = val.to_string();
                start.push_attribute(("w:cryptSpinCount", s.as_str()));
            }
        }
        if let Some(ref val) = self.crypt_provider {
            start.push_attribute(("w:cryptProvider", val.as_str()));
        }
        if let Some(ref val) = self.alg_id_ext {
            {
                let hex = encode_hex(val);
                start.push_attribute(("w:algIdExt", hex.as_str()));
            }
        }
        if let Some(ref val) = self.alg_id_ext_source {
            start.push_attribute(("w:algIdExtSource", val.as_str()));
        }
        if let Some(ref val) = self.crypt_provider_type_ext {
            {
                let hex = encode_hex(val);
                start.push_attribute(("w:cryptProviderTypeExt", hex.as_str()));
            }
        }
        if let Some(ref val) = self.crypt_provider_type_ext_source {
            start.push_attribute(("w:cryptProviderTypeExtSource", val.as_str()));
        }
        if let Some(ref val) = self.hash {
            {
                let hex = encode_hex(val);
                start.push_attribute(("w:hash", hex.as_str()));
            }
        }
        if let Some(ref val) = self.salt {
            {
                let hex = encode_hex(val);
                start.push_attribute(("w:salt", hex.as_str()));
            }
        }
        #[cfg(feature = "extra-attrs")]
        for (key, value) in &self.extra_attrs {
            start.push_attribute((key.as_str(), value.as_str()));
        }
        start
    }

    fn is_empty_element(&self) -> bool {
        true
    }
}

impl ToXml for CTDocProtect {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        if let Some(ref val) = self.edit {
            {
                let s = val.to_string();
                start.push_attribute(("w:edit", s.as_str()));
            }
        }
        if let Some(ref val) = self.formatting {
            {
                let s = val.to_string();
                start.push_attribute(("w:formatting", s.as_str()));
            }
        }
        if let Some(ref val) = self.enforcement {
            {
                let s = val.to_string();
                start.push_attribute(("w:enforcement", s.as_str()));
            }
        }
        if let Some(ref val) = self.algorithm_name {
            start.push_attribute(("w:algorithmName", val.as_str()));
        }
        if let Some(ref val) = self.hash_value {
            {
                let hex = encode_hex(val);
                start.push_attribute(("w:hashValue", hex.as_str()));
            }
        }
        if let Some(ref val) = self.salt_value {
            {
                let hex = encode_hex(val);
                start.push_attribute(("w:saltValue", hex.as_str()));
            }
        }
        if let Some(ref val) = self.spin_count {
            {
                let s = val.to_string();
                start.push_attribute(("w:spinCount", s.as_str()));
            }
        }
        if let Some(ref val) = self.crypt_provider_type {
            {
                let s = val.to_string();
                start.push_attribute(("w:cryptProviderType", s.as_str()));
            }
        }
        if let Some(ref val) = self.crypt_algorithm_class {
            {
                let s = val.to_string();
                start.push_attribute(("w:cryptAlgorithmClass", s.as_str()));
            }
        }
        if let Some(ref val) = self.crypt_algorithm_type {
            {
                let s = val.to_string();
                start.push_attribute(("w:cryptAlgorithmType", s.as_str()));
            }
        }
        if let Some(ref val) = self.crypt_algorithm_sid {
            {
                let s = val.to_string();
                start.push_attribute(("w:cryptAlgorithmSid", s.as_str()));
            }
        }
        if let Some(ref val) = self.crypt_spin_count {
            {
                let s = val.to_string();
                start.push_attribute(("w:cryptSpinCount", s.as_str()));
            }
        }
        if let Some(ref val) = self.crypt_provider {
            start.push_attribute(("w:cryptProvider", val.as_str()));
        }
        if let Some(ref val) = self.alg_id_ext {
            {
                let hex = encode_hex(val);
                start.push_attribute(("w:algIdExt", hex.as_str()));
            }
        }
        if let Some(ref val) = self.alg_id_ext_source {
            start.push_attribute(("w:algIdExtSource", val.as_str()));
        }
        if let Some(ref val) = self.crypt_provider_type_ext {
            {
                let hex = encode_hex(val);
                start.push_attribute(("w:cryptProviderTypeExt", hex.as_str()));
            }
        }
        if let Some(ref val) = self.crypt_provider_type_ext_source {
            start.push_attribute(("w:cryptProviderTypeExtSource", val.as_str()));
        }
        if let Some(ref val) = self.hash {
            {
                let hex = encode_hex(val);
                start.push_attribute(("w:hash", hex.as_str()));
            }
        }
        if let Some(ref val) = self.salt {
            {
                let hex = encode_hex(val);
                start.push_attribute(("w:salt", hex.as_str()));
            }
        }
        #[cfg(feature = "extra-attrs")]
        for (key, value) in &self.extra_attrs {
            start.push_attribute((key.as_str(), value.as_str()));
        }
        start
    }

    fn is_empty_element(&self) -> bool {
        true
    }
}

impl ToXml for CTMailMergeDocType {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        {
            let val = &self.value;
            {
                let s = val.to_string();
                start.push_attribute(("w:val", s.as_str()));
            }
        }
        #[cfg(feature = "extra-attrs")]
        for (key, value) in &self.extra_attrs {
            start.push_attribute((key.as_str(), value.as_str()));
        }
        start
    }

    fn is_empty_element(&self) -> bool {
        true
    }
}

impl ToXml for CTMailMergeDataType {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        {
            let val = &self.value;
            start.push_attribute(("w:val", val.as_str()));
        }
        #[cfg(feature = "extra-attrs")]
        for (key, value) in &self.extra_attrs {
            start.push_attribute((key.as_str(), value.as_str()));
        }
        start
    }

    fn is_empty_element(&self) -> bool {
        true
    }
}

impl ToXml for CTMailMergeDest {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        {
            let val = &self.value;
            {
                let s = val.to_string();
                start.push_attribute(("w:val", s.as_str()));
            }
        }
        #[cfg(feature = "extra-attrs")]
        for (key, value) in &self.extra_attrs {
            start.push_attribute((key.as_str(), value.as_str()));
        }
        start
    }

    fn is_empty_element(&self) -> bool {
        true
    }
}

impl ToXml for CTMailMergeOdsoFMDFieldType {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        {
            let val = &self.value;
            {
                let s = val.to_string();
                start.push_attribute(("w:val", s.as_str()));
            }
        }
        #[cfg(feature = "extra-attrs")]
        for (key, value) in &self.extra_attrs {
            start.push_attribute((key.as_str(), value.as_str()));
        }
        start
    }

    fn is_empty_element(&self) -> bool {
        true
    }
}

impl ToXml for CTTrackChangesView {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        if let Some(ref val) = self.markup {
            {
                let s = val.to_string();
                start.push_attribute(("w:markup", s.as_str()));
            }
        }
        if let Some(ref val) = self.comments {
            {
                let s = val.to_string();
                start.push_attribute(("w:comments", s.as_str()));
            }
        }
        if let Some(ref val) = self.ins_del {
            {
                let s = val.to_string();
                start.push_attribute(("w:insDel", s.as_str()));
            }
        }
        if let Some(ref val) = self.formatting {
            {
                let s = val.to_string();
                start.push_attribute(("w:formatting", s.as_str()));
            }
        }
        if let Some(ref val) = self.ink_annotations {
            {
                let s = val.to_string();
                start.push_attribute(("w:inkAnnotations", s.as_str()));
            }
        }
        #[cfg(feature = "extra-attrs")]
        for (key, value) in &self.extra_attrs {
            start.push_attribute((key.as_str(), value.as_str()));
        }
        start
    }

    fn is_empty_element(&self) -> bool {
        true
    }
}

impl ToXml for CTKinsoku {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        {
            let val = &self.lang;
            start.push_attribute(("w:lang", val.as_str()));
        }
        {
            let val = &self.value;
            start.push_attribute(("w:val", val.as_str()));
        }
        #[cfg(feature = "extra-attrs")]
        for (key, value) in &self.extra_attrs {
            start.push_attribute((key.as_str(), value.as_str()));
        }
        start
    }

    fn is_empty_element(&self) -> bool {
        true
    }
}

impl ToXml for CTTextDirection {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        {
            let val = &self.value;
            {
                let s = val.to_string();
                start.push_attribute(("w:val", s.as_str()));
            }
        }
        #[cfg(feature = "extra-attrs")]
        for (key, value) in &self.extra_attrs {
            start.push_attribute((key.as_str(), value.as_str()));
        }
        start
    }

    fn is_empty_element(&self) -> bool {
        true
    }
}

impl ToXml for CTTextAlignment {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        {
            let val = &self.value;
            {
                let s = val.to_string();
                start.push_attribute(("w:val", s.as_str()));
            }
        }
        #[cfg(feature = "extra-attrs")]
        for (key, value) in &self.extra_attrs {
            start.push_attribute((key.as_str(), value.as_str()));
        }
        start
    }

    fn is_empty_element(&self) -> bool {
        true
    }
}

impl ToXml for CTMarkup {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        {
            let val = &self.id;
            {
                let s = val.to_string();
                start.push_attribute(("w:id", s.as_str()));
            }
        }
        #[cfg(feature = "extra-attrs")]
        for (key, value) in &self.extra_attrs {
            start.push_attribute((key.as_str(), value.as_str()));
        }
        start
    }

    fn is_empty_element(&self) -> bool {
        true
    }
}

impl ToXml for CTTrackChange {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        {
            let val = &self.id;
            {
                let s = val.to_string();
                start.push_attribute(("w:id", s.as_str()));
            }
        }
        {
            let val = &self.author;
            start.push_attribute(("w:author", val.as_str()));
        }
        if let Some(ref val) = self.date {
            start.push_attribute(("w:date", val.as_str()));
        }
        #[cfg(feature = "extra-attrs")]
        for (key, value) in &self.extra_attrs {
            start.push_attribute((key.as_str(), value.as_str()));
        }
        start
    }

    fn is_empty_element(&self) -> bool {
        true
    }
}

impl ToXml for CTCellMergeTrackChange {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        {
            let val = &self.id;
            {
                let s = val.to_string();
                start.push_attribute(("w:id", s.as_str()));
            }
        }
        {
            let val = &self.author;
            start.push_attribute(("w:author", val.as_str()));
        }
        if let Some(ref val) = self.date {
            start.push_attribute(("w:date", val.as_str()));
        }
        if let Some(ref val) = self.vertical_merge {
            {
                let s = val.to_string();
                start.push_attribute(("w:vMerge", s.as_str()));
            }
        }
        if let Some(ref val) = self.v_merge_orig {
            {
                let s = val.to_string();
                start.push_attribute(("w:vMergeOrig", s.as_str()));
            }
        }
        #[cfg(feature = "extra-attrs")]
        for (key, value) in &self.extra_attrs {
            start.push_attribute((key.as_str(), value.as_str()));
        }
        start
    }

    fn is_empty_element(&self) -> bool {
        true
    }
}

impl ToXml for CTTrackChangeRange {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        {
            let val = &self.id;
            {
                let s = val.to_string();
                start.push_attribute(("w:id", s.as_str()));
            }
        }
        {
            let val = &self.author;
            start.push_attribute(("w:author", val.as_str()));
        }
        if let Some(ref val) = self.date {
            start.push_attribute(("w:date", val.as_str()));
        }
        if let Some(ref val) = self.displaced_by_custom_xml {
            {
                let s = val.to_string();
                start.push_attribute(("w:displacedByCustomXml", s.as_str()));
            }
        }
        #[cfg(feature = "extra-attrs")]
        for (key, value) in &self.extra_attrs {
            start.push_attribute((key.as_str(), value.as_str()));
        }
        start
    }

    fn is_empty_element(&self) -> bool {
        true
    }
}

impl ToXml for CTMarkupRange {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        {
            let val = &self.id;
            {
                let s = val.to_string();
                start.push_attribute(("w:id", s.as_str()));
            }
        }
        if let Some(ref val) = self.displaced_by_custom_xml {
            {
                let s = val.to_string();
                start.push_attribute(("w:displacedByCustomXml", s.as_str()));
            }
        }
        #[cfg(feature = "extra-attrs")]
        for (key, value) in &self.extra_attrs {
            start.push_attribute((key.as_str(), value.as_str()));
        }
        start
    }

    fn is_empty_element(&self) -> bool {
        true
    }
}

impl ToXml for CTBookmarkRange {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        {
            let val = &self.id;
            {
                let s = val.to_string();
                start.push_attribute(("w:id", s.as_str()));
            }
        }
        if let Some(ref val) = self.displaced_by_custom_xml {
            {
                let s = val.to_string();
                start.push_attribute(("w:displacedByCustomXml", s.as_str()));
            }
        }
        if let Some(ref val) = self.col_first {
            {
                let s = val.to_string();
                start.push_attribute(("w:colFirst", s.as_str()));
            }
        }
        if let Some(ref val) = self.col_last {
            {
                let s = val.to_string();
                start.push_attribute(("w:colLast", s.as_str()));
            }
        }
        #[cfg(feature = "extra-attrs")]
        for (key, value) in &self.extra_attrs {
            start.push_attribute((key.as_str(), value.as_str()));
        }
        start
    }

    fn is_empty_element(&self) -> bool {
        true
    }
}

impl ToXml for Bookmark {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        {
            let val = &self.id;
            {
                let s = val.to_string();
                start.push_attribute(("w:id", s.as_str()));
            }
        }
        if let Some(ref val) = self.displaced_by_custom_xml {
            {
                let s = val.to_string();
                start.push_attribute(("w:displacedByCustomXml", s.as_str()));
            }
        }
        if let Some(ref val) = self.col_first {
            {
                let s = val.to_string();
                start.push_attribute(("w:colFirst", s.as_str()));
            }
        }
        if let Some(ref val) = self.col_last {
            {
                let s = val.to_string();
                start.push_attribute(("w:colLast", s.as_str()));
            }
        }
        {
            let val = &self.name;
            start.push_attribute(("w:name", val.as_str()));
        }
        #[cfg(feature = "extra-attrs")]
        for (key, value) in &self.extra_attrs {
            start.push_attribute((key.as_str(), value.as_str()));
        }
        start
    }

    fn is_empty_element(&self) -> bool {
        true
    }
}

impl ToXml for CTMoveBookmark {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        {
            let val = &self.id;
            {
                let s = val.to_string();
                start.push_attribute(("w:id", s.as_str()));
            }
        }
        if let Some(ref val) = self.displaced_by_custom_xml {
            {
                let s = val.to_string();
                start.push_attribute(("w:displacedByCustomXml", s.as_str()));
            }
        }
        if let Some(ref val) = self.col_first {
            {
                let s = val.to_string();
                start.push_attribute(("w:colFirst", s.as_str()));
            }
        }
        if let Some(ref val) = self.col_last {
            {
                let s = val.to_string();
                start.push_attribute(("w:colLast", s.as_str()));
            }
        }
        {
            let val = &self.name;
            start.push_attribute(("w:name", val.as_str()));
        }
        {
            let val = &self.author;
            start.push_attribute(("w:author", val.as_str()));
        }
        {
            let val = &self.date;
            start.push_attribute(("w:date", val.as_str()));
        }
        #[cfg(feature = "extra-attrs")]
        for (key, value) in &self.extra_attrs {
            start.push_attribute((key.as_str(), value.as_str()));
        }
        start
    }

    fn is_empty_element(&self) -> bool {
        true
    }
}

impl ToXml for Comment {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        {
            let val = &self.id;
            {
                let s = val.to_string();
                start.push_attribute(("w:id", s.as_str()));
            }
        }
        {
            let val = &self.author;
            start.push_attribute(("w:author", val.as_str()));
        }
        if let Some(ref val) = self.date {
            start.push_attribute(("w:date", val.as_str()));
        }
        if let Some(ref val) = self.initials {
            start.push_attribute(("w:initials", val.as_str()));
        }
        #[cfg(feature = "extra-attrs")]
        for (key, value) in &self.extra_attrs {
            start.push_attribute((key.as_str(), value.as_str()));
        }
        start
    }

    fn write_children<W: Write>(&self, writer: &mut Writer<W>) -> Result<(), SerializeError> {
        for item in &self.block_level_elts {
            item.write_element("", writer)?;
        }
        #[cfg(feature = "extra-children")]
        for child in &self.extra_children {
            child.write_to(writer).map_err(SerializeError::from)?;
        }
        Ok(())
    }

    fn is_empty_element(&self) -> bool {
        if !self.block_level_elts.is_empty() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for CTTrackChangeNumbering {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        {
            let val = &self.id;
            {
                let s = val.to_string();
                start.push_attribute(("w:id", s.as_str()));
            }
        }
        {
            let val = &self.author;
            start.push_attribute(("w:author", val.as_str()));
        }
        if let Some(ref val) = self.date {
            start.push_attribute(("w:date", val.as_str()));
        }
        if let Some(ref val) = self.original {
            start.push_attribute(("w:original", val.as_str()));
        }
        #[cfg(feature = "extra-attrs")]
        for (key, value) in &self.extra_attrs {
            start.push_attribute((key.as_str(), value.as_str()));
        }
        start
    }

    fn is_empty_element(&self) -> bool {
        true
    }
}

impl ToXml for CTTblPrExChange {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        {
            let val = &self.id;
            {
                let s = val.to_string();
                start.push_attribute(("w:id", s.as_str()));
            }
        }
        {
            let val = &self.author;
            start.push_attribute(("w:author", val.as_str()));
        }
        if let Some(ref val) = self.date {
            start.push_attribute(("w:date", val.as_str()));
        }
        #[cfg(feature = "extra-attrs")]
        for (key, value) in &self.extra_attrs {
            start.push_attribute((key.as_str(), value.as_str()));
        }
        start
    }

    fn write_children<W: Write>(&self, writer: &mut Writer<W>) -> Result<(), SerializeError> {
        {
            let val = &self.tbl_pr_ex;
            val.write_element("w:tblPrEx", writer)?;
        }
        #[cfg(feature = "extra-children")]
        for child in &self.extra_children {
            child.write_to(writer).map_err(SerializeError::from)?;
        }
        Ok(())
    }

    fn is_empty_element(&self) -> bool {
        false
    }
}

impl ToXml for CTTcPrChange {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        {
            let val = &self.id;
            {
                let s = val.to_string();
                start.push_attribute(("w:id", s.as_str()));
            }
        }
        {
            let val = &self.author;
            start.push_attribute(("w:author", val.as_str()));
        }
        if let Some(ref val) = self.date {
            start.push_attribute(("w:date", val.as_str()));
        }
        #[cfg(feature = "extra-attrs")]
        for (key, value) in &self.extra_attrs {
            start.push_attribute((key.as_str(), value.as_str()));
        }
        start
    }

    fn write_children<W: Write>(&self, writer: &mut Writer<W>) -> Result<(), SerializeError> {
        {
            let val = &self.cell_properties;
            val.write_element("w:tcPr", writer)?;
        }
        #[cfg(feature = "extra-children")]
        for child in &self.extra_children {
            child.write_to(writer).map_err(SerializeError::from)?;
        }
        Ok(())
    }

    fn is_empty_element(&self) -> bool {
        false
    }
}

impl ToXml for CTTrPrChange {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        {
            let val = &self.id;
            {
                let s = val.to_string();
                start.push_attribute(("w:id", s.as_str()));
            }
        }
        {
            let val = &self.author;
            start.push_attribute(("w:author", val.as_str()));
        }
        if let Some(ref val) = self.date {
            start.push_attribute(("w:date", val.as_str()));
        }
        #[cfg(feature = "extra-attrs")]
        for (key, value) in &self.extra_attrs {
            start.push_attribute((key.as_str(), value.as_str()));
        }
        start
    }

    fn write_children<W: Write>(&self, writer: &mut Writer<W>) -> Result<(), SerializeError> {
        {
            let val = &self.row_properties;
            val.write_element("w:trPr", writer)?;
        }
        #[cfg(feature = "extra-children")]
        for child in &self.extra_children {
            child.write_to(writer).map_err(SerializeError::from)?;
        }
        Ok(())
    }

    fn is_empty_element(&self) -> bool {
        false
    }
}

impl ToXml for CTTblGridChange {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        {
            let val = &self.id;
            {
                let s = val.to_string();
                start.push_attribute(("w:id", s.as_str()));
            }
        }
        #[cfg(feature = "extra-attrs")]
        for (key, value) in &self.extra_attrs {
            start.push_attribute((key.as_str(), value.as_str()));
        }
        start
    }

    fn write_children<W: Write>(&self, writer: &mut Writer<W>) -> Result<(), SerializeError> {
        {
            let val = &self.tbl_grid;
            val.write_element("w:tblGrid", writer)?;
        }
        #[cfg(feature = "extra-children")]
        for child in &self.extra_children {
            child.write_to(writer).map_err(SerializeError::from)?;
        }
        Ok(())
    }

    fn is_empty_element(&self) -> bool {
        false
    }
}

impl ToXml for CTTblPrChange {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        {
            let val = &self.id;
            {
                let s = val.to_string();
                start.push_attribute(("w:id", s.as_str()));
            }
        }
        {
            let val = &self.author;
            start.push_attribute(("w:author", val.as_str()));
        }
        if let Some(ref val) = self.date {
            start.push_attribute(("w:date", val.as_str()));
        }
        #[cfg(feature = "extra-attrs")]
        for (key, value) in &self.extra_attrs {
            start.push_attribute((key.as_str(), value.as_str()));
        }
        start
    }

    fn write_children<W: Write>(&self, writer: &mut Writer<W>) -> Result<(), SerializeError> {
        {
            let val = &self.table_properties;
            val.write_element("w:tblPr", writer)?;
        }
        #[cfg(feature = "extra-children")]
        for child in &self.extra_children {
            child.write_to(writer).map_err(SerializeError::from)?;
        }
        Ok(())
    }

    fn is_empty_element(&self) -> bool {
        false
    }
}

impl ToXml for CTSectPrChange {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        {
            let val = &self.id;
            {
                let s = val.to_string();
                start.push_attribute(("w:id", s.as_str()));
            }
        }
        {
            let val = &self.author;
            start.push_attribute(("w:author", val.as_str()));
        }
        if let Some(ref val) = self.date {
            start.push_attribute(("w:date", val.as_str()));
        }
        #[cfg(feature = "extra-attrs")]
        for (key, value) in &self.extra_attrs {
            start.push_attribute((key.as_str(), value.as_str()));
        }
        start
    }

    fn write_children<W: Write>(&self, writer: &mut Writer<W>) -> Result<(), SerializeError> {
        if let Some(ref val) = self.sect_pr {
            val.write_element("w:sectPr", writer)?;
        }
        #[cfg(feature = "extra-children")]
        for child in &self.extra_children {
            child.write_to(writer).map_err(SerializeError::from)?;
        }
        Ok(())
    }

    fn is_empty_element(&self) -> bool {
        if self.sect_pr.is_some() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for CTPPrChange {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        {
            let val = &self.id;
            {
                let s = val.to_string();
                start.push_attribute(("w:id", s.as_str()));
            }
        }
        {
            let val = &self.author;
            start.push_attribute(("w:author", val.as_str()));
        }
        if let Some(ref val) = self.date {
            start.push_attribute(("w:date", val.as_str()));
        }
        #[cfg(feature = "extra-attrs")]
        for (key, value) in &self.extra_attrs {
            start.push_attribute((key.as_str(), value.as_str()));
        }
        start
    }

    fn write_children<W: Write>(&self, writer: &mut Writer<W>) -> Result<(), SerializeError> {
        {
            let val = &self.p_pr;
            val.write_element("w:pPr", writer)?;
        }
        #[cfg(feature = "extra-children")]
        for child in &self.extra_children {
            child.write_to(writer).map_err(SerializeError::from)?;
        }
        Ok(())
    }

    fn is_empty_element(&self) -> bool {
        false
    }
}

impl ToXml for CTRPrChange {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        {
            let val = &self.id;
            {
                let s = val.to_string();
                start.push_attribute(("w:id", s.as_str()));
            }
        }
        {
            let val = &self.author;
            start.push_attribute(("w:author", val.as_str()));
        }
        if let Some(ref val) = self.date {
            start.push_attribute(("w:date", val.as_str()));
        }
        #[cfg(feature = "extra-attrs")]
        for (key, value) in &self.extra_attrs {
            start.push_attribute((key.as_str(), value.as_str()));
        }
        start
    }

    fn write_children<W: Write>(&self, writer: &mut Writer<W>) -> Result<(), SerializeError> {
        {
            let val = &self.r_pr;
            val.write_element("w:rPr", writer)?;
        }
        #[cfg(feature = "extra-children")]
        for child in &self.extra_children {
            child.write_to(writer).map_err(SerializeError::from)?;
        }
        Ok(())
    }

    fn is_empty_element(&self) -> bool {
        false
    }
}

impl ToXml for CTParaRPrChange {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        {
            let val = &self.id;
            {
                let s = val.to_string();
                start.push_attribute(("w:id", s.as_str()));
            }
        }
        {
            let val = &self.author;
            start.push_attribute(("w:author", val.as_str()));
        }
        if let Some(ref val) = self.date {
            start.push_attribute(("w:date", val.as_str()));
        }
        #[cfg(feature = "extra-attrs")]
        for (key, value) in &self.extra_attrs {
            start.push_attribute((key.as_str(), value.as_str()));
        }
        start
    }

    fn write_children<W: Write>(&self, writer: &mut Writer<W>) -> Result<(), SerializeError> {
        {
            let val = &self.r_pr;
            val.write_element("w:rPr", writer)?;
        }
        #[cfg(feature = "extra-children")]
        for child in &self.extra_children {
            child.write_to(writer).map_err(SerializeError::from)?;
        }
        Ok(())
    }

    fn is_empty_element(&self) -> bool {
        false
    }
}

impl ToXml for CTRunTrackChange {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        {
            let val = &self.id;
            {
                let s = val.to_string();
                start.push_attribute(("w:id", s.as_str()));
            }
        }
        {
            let val = &self.author;
            start.push_attribute(("w:author", val.as_str()));
        }
        if let Some(ref val) = self.date {
            start.push_attribute(("w:date", val.as_str()));
        }
        #[cfg(feature = "extra-attrs")]
        for (key, value) in &self.extra_attrs {
            start.push_attribute((key.as_str(), value.as_str()));
        }
        start
    }

    fn is_empty_element(&self) -> bool {
        true
    }
}

impl ToXml for EGPContentMath {
    fn write_children<W: Write>(&self, writer: &mut Writer<W>) -> Result<(), SerializeError> {
        #[cfg(feature = "extra-children")]
        for child in &self.extra_children {
            child.write_to(writer).map_err(SerializeError::from)?;
        }
        Ok(())
    }

    fn is_empty_element(&self) -> bool {
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for EGPContentBase {
    fn write_element<W: Write>(
        &self,
        _tag: &str,
        writer: &mut Writer<W>,
    ) -> Result<(), SerializeError> {
        match self {
            Self::CustomXml(inner) => inner.write_element("w:customXml", writer)?,
            Self::FldSimple(inner) => inner.write_element("w:fldSimple", writer)?,
            Self::Hyperlink(inner) => inner.write_element("w:hyperlink", writer)?,
        }
        Ok(())
    }
}

impl ToXml for EGContentRunContentBase {
    fn write_element<W: Write>(
        &self,
        _tag: &str,
        writer: &mut Writer<W>,
    ) -> Result<(), SerializeError> {
        match self {
            Self::SmartTag(inner) => inner.write_element("w:smartTag", writer)?,
            Self::Sdt(inner) => inner.write_element("w:sdt", writer)?,
            Self::ProofErr(inner) => inner.write_element("w:proofErr", writer)?,
            Self::PermStart(inner) => inner.write_element("w:permStart", writer)?,
            Self::PermEnd(inner) => inner.write_element("w:permEnd", writer)?,
            Self::BookmarkStart(inner) => inner.write_element("w:bookmarkStart", writer)?,
            Self::BookmarkEnd(inner) => inner.write_element("w:bookmarkEnd", writer)?,
            Self::MoveFromRangeStart(inner) => {
                inner.write_element("w:moveFromRangeStart", writer)?
            }
            Self::MoveFromRangeEnd(inner) => inner.write_element("w:moveFromRangeEnd", writer)?,
            Self::MoveToRangeStart(inner) => inner.write_element("w:moveToRangeStart", writer)?,
            Self::MoveToRangeEnd(inner) => inner.write_element("w:moveToRangeEnd", writer)?,
            Self::CommentRangeStart(inner) => inner.write_element("w:commentRangeStart", writer)?,
            Self::CommentRangeEnd(inner) => inner.write_element("w:commentRangeEnd", writer)?,
            Self::CustomXmlInsRangeStart(inner) => {
                inner.write_element("w:customXmlInsRangeStart", writer)?
            }
            Self::CustomXmlInsRangeEnd(inner) => {
                inner.write_element("w:customXmlInsRangeEnd", writer)?
            }
            Self::CustomXmlDelRangeStart(inner) => {
                inner.write_element("w:customXmlDelRangeStart", writer)?
            }
            Self::CustomXmlDelRangeEnd(inner) => {
                inner.write_element("w:customXmlDelRangeEnd", writer)?
            }
            Self::CustomXmlMoveFromRangeStart(inner) => {
                inner.write_element("w:customXmlMoveFromRangeStart", writer)?
            }
            Self::CustomXmlMoveFromRangeEnd(inner) => {
                inner.write_element("w:customXmlMoveFromRangeEnd", writer)?
            }
            Self::CustomXmlMoveToRangeStart(inner) => {
                inner.write_element("w:customXmlMoveToRangeStart", writer)?
            }
            Self::CustomXmlMoveToRangeEnd(inner) => {
                inner.write_element("w:customXmlMoveToRangeEnd", writer)?
            }
            Self::Ins(inner) => inner.write_element("w:ins", writer)?,
            Self::Del(inner) => inner.write_element("w:del", writer)?,
            Self::MoveFrom(inner) => inner.write_element("w:moveFrom", writer)?,
            Self::MoveTo(inner) => inner.write_element("w:moveTo", writer)?,
        }
        Ok(())
    }
}

impl ToXml for EGCellMarkupElements {
    fn write_element<W: Write>(
        &self,
        _tag: &str,
        writer: &mut Writer<W>,
    ) -> Result<(), SerializeError> {
        match self {
            Self::CellIns(inner) => inner.write_element("w:cellIns", writer)?,
            Self::CellDel(inner) => inner.write_element("w:cellDel", writer)?,
            Self::CellMerge(inner) => inner.write_element("w:cellMerge", writer)?,
        }
        Ok(())
    }
}

impl ToXml for EGRangeMarkupElements {
    fn write_element<W: Write>(
        &self,
        _tag: &str,
        writer: &mut Writer<W>,
    ) -> Result<(), SerializeError> {
        match self {
            Self::BookmarkStart(inner) => inner.write_element("w:bookmarkStart", writer)?,
            Self::BookmarkEnd(inner) => inner.write_element("w:bookmarkEnd", writer)?,
            Self::MoveFromRangeStart(inner) => {
                inner.write_element("w:moveFromRangeStart", writer)?
            }
            Self::MoveFromRangeEnd(inner) => inner.write_element("w:moveFromRangeEnd", writer)?,
            Self::MoveToRangeStart(inner) => inner.write_element("w:moveToRangeStart", writer)?,
            Self::MoveToRangeEnd(inner) => inner.write_element("w:moveToRangeEnd", writer)?,
            Self::CommentRangeStart(inner) => inner.write_element("w:commentRangeStart", writer)?,
            Self::CommentRangeEnd(inner) => inner.write_element("w:commentRangeEnd", writer)?,
            Self::CustomXmlInsRangeStart(inner) => {
                inner.write_element("w:customXmlInsRangeStart", writer)?
            }
            Self::CustomXmlInsRangeEnd(inner) => {
                inner.write_element("w:customXmlInsRangeEnd", writer)?
            }
            Self::CustomXmlDelRangeStart(inner) => {
                inner.write_element("w:customXmlDelRangeStart", writer)?
            }
            Self::CustomXmlDelRangeEnd(inner) => {
                inner.write_element("w:customXmlDelRangeEnd", writer)?
            }
            Self::CustomXmlMoveFromRangeStart(inner) => {
                inner.write_element("w:customXmlMoveFromRangeStart", writer)?
            }
            Self::CustomXmlMoveFromRangeEnd(inner) => {
                inner.write_element("w:customXmlMoveFromRangeEnd", writer)?
            }
            Self::CustomXmlMoveToRangeStart(inner) => {
                inner.write_element("w:customXmlMoveToRangeStart", writer)?
            }
            Self::CustomXmlMoveToRangeEnd(inner) => {
                inner.write_element("w:customXmlMoveToRangeEnd", writer)?
            }
        }
        Ok(())
    }
}

impl ToXml for CTNumPr {
    fn write_children<W: Write>(&self, writer: &mut Writer<W>) -> Result<(), SerializeError> {
        if let Some(ref val) = self.ilvl {
            val.write_element("w:ilvl", writer)?;
        }
        if let Some(ref val) = self.num_id {
            val.write_element("w:numId", writer)?;
        }
        if let Some(ref val) = self.numbering_change {
            val.write_element("w:numberingChange", writer)?;
        }
        if let Some(ref val) = self.ins {
            val.write_element("w:ins", writer)?;
        }
        #[cfg(feature = "extra-children")]
        for child in &self.extra_children {
            child.write_to(writer).map_err(SerializeError::from)?;
        }
        Ok(())
    }

    fn is_empty_element(&self) -> bool {
        if self.ilvl.is_some() {
            return false;
        }
        if self.num_id.is_some() {
            return false;
        }
        if self.numbering_change.is_some() {
            return false;
        }
        if self.ins.is_some() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for CTPBdr {
    fn write_children<W: Write>(&self, writer: &mut Writer<W>) -> Result<(), SerializeError> {
        if let Some(ref val) = self.top {
            val.write_element("w:top", writer)?;
        }
        if let Some(ref val) = self.left {
            val.write_element("w:left", writer)?;
        }
        if let Some(ref val) = self.bottom {
            val.write_element("w:bottom", writer)?;
        }
        if let Some(ref val) = self.right {
            val.write_element("w:right", writer)?;
        }
        if let Some(ref val) = self.between {
            val.write_element("w:between", writer)?;
        }
        if let Some(ref val) = self.bar {
            val.write_element("w:bar", writer)?;
        }
        #[cfg(feature = "extra-children")]
        for child in &self.extra_children {
            child.write_to(writer).map_err(SerializeError::from)?;
        }
        Ok(())
    }

    fn is_empty_element(&self) -> bool {
        if self.top.is_some() {
            return false;
        }
        if self.left.is_some() {
            return false;
        }
        if self.bottom.is_some() {
            return false;
        }
        if self.right.is_some() {
            return false;
        }
        if self.between.is_some() {
            return false;
        }
        if self.bar.is_some() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for CTTabs {
    fn write_children<W: Write>(&self, writer: &mut Writer<W>) -> Result<(), SerializeError> {
        for item in &self.tab {
            item.write_element("w:tab", writer)?;
        }
        #[cfg(feature = "extra-children")]
        for child in &self.extra_children {
            child.write_to(writer).map_err(SerializeError::from)?;
        }
        Ok(())
    }

    fn is_empty_element(&self) -> bool {
        if !self.tab.is_empty() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for CTTextboxTightWrap {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        {
            let val = &self.value;
            {
                let s = val.to_string();
                start.push_attribute(("w:val", s.as_str()));
            }
        }
        #[cfg(feature = "extra-attrs")]
        for (key, value) in &self.extra_attrs {
            start.push_attribute((key.as_str(), value.as_str()));
        }
        start
    }

    fn is_empty_element(&self) -> bool {
        true
    }
}

impl ToXml for ParagraphProperties {
    fn write_children<W: Write>(&self, writer: &mut Writer<W>) -> Result<(), SerializeError> {
        if let Some(ref val) = self.paragraph_style {
            val.write_element("w:pStyle", writer)?;
        }
        if let Some(ref val) = self.keep_next {
            val.write_element("w:keepNext", writer)?;
        }
        if let Some(ref val) = self.keep_lines {
            val.write_element("w:keepLines", writer)?;
        }
        if let Some(ref val) = self.page_break_before {
            val.write_element("w:pageBreakBefore", writer)?;
        }
        if let Some(ref val) = self.frame_pr {
            val.write_element("w:framePr", writer)?;
        }
        if let Some(ref val) = self.widow_control {
            val.write_element("w:widowControl", writer)?;
        }
        if let Some(ref val) = self.num_pr {
            val.write_element("w:numPr", writer)?;
        }
        if let Some(ref val) = self.suppress_line_numbers {
            val.write_element("w:suppressLineNumbers", writer)?;
        }
        if let Some(ref val) = self.paragraph_border {
            val.write_element("w:pBdr", writer)?;
        }
        if let Some(ref val) = self.shading {
            val.write_element("w:shd", writer)?;
        }
        if let Some(ref val) = self.tabs {
            val.write_element("w:tabs", writer)?;
        }
        if let Some(ref val) = self.suppress_auto_hyphens {
            val.write_element("w:suppressAutoHyphens", writer)?;
        }
        if let Some(ref val) = self.kinsoku {
            val.write_element("w:kinsoku", writer)?;
        }
        if let Some(ref val) = self.word_wrap {
            val.write_element("w:wordWrap", writer)?;
        }
        if let Some(ref val) = self.overflow_punct {
            val.write_element("w:overflowPunct", writer)?;
        }
        if let Some(ref val) = self.top_line_punct {
            val.write_element("w:topLinePunct", writer)?;
        }
        if let Some(ref val) = self.auto_space_d_e {
            val.write_element("w:autoSpaceDE", writer)?;
        }
        if let Some(ref val) = self.auto_space_d_n {
            val.write_element("w:autoSpaceDN", writer)?;
        }
        if let Some(ref val) = self.bidi {
            val.write_element("w:bidi", writer)?;
        }
        if let Some(ref val) = self.adjust_right_ind {
            val.write_element("w:adjustRightInd", writer)?;
        }
        if let Some(ref val) = self.snap_to_grid {
            val.write_element("w:snapToGrid", writer)?;
        }
        if let Some(ref val) = self.spacing {
            val.write_element("w:spacing", writer)?;
        }
        if let Some(ref val) = self.indentation {
            val.write_element("w:ind", writer)?;
        }
        if let Some(ref val) = self.contextual_spacing {
            val.write_element("w:contextualSpacing", writer)?;
        }
        if let Some(ref val) = self.mirror_indents {
            val.write_element("w:mirrorIndents", writer)?;
        }
        if let Some(ref val) = self.suppress_overlap {
            val.write_element("w:suppressOverlap", writer)?;
        }
        if let Some(ref val) = self.justification {
            val.write_element("w:jc", writer)?;
        }
        if let Some(ref val) = self.text_direction {
            val.write_element("w:textDirection", writer)?;
        }
        if let Some(ref val) = self.text_alignment {
            val.write_element("w:textAlignment", writer)?;
        }
        if let Some(ref val) = self.textbox_tight_wrap {
            val.write_element("w:textboxTightWrap", writer)?;
        }
        if let Some(ref val) = self.outline_lvl {
            val.write_element("w:outlineLvl", writer)?;
        }
        if let Some(ref val) = self.div_id {
            val.write_element("w:divId", writer)?;
        }
        if let Some(ref val) = self.cnf_style {
            val.write_element("w:cnfStyle", writer)?;
        }
        if let Some(ref val) = self.r_pr {
            val.write_element("w:rPr", writer)?;
        }
        if let Some(ref val) = self.sect_pr {
            val.write_element("w:sectPr", writer)?;
        }
        if let Some(ref val) = self.p_pr_change {
            val.write_element("w:pPrChange", writer)?;
        }
        #[cfg(feature = "extra-children")]
        for child in &self.extra_children {
            child.write_to(writer).map_err(SerializeError::from)?;
        }
        Ok(())
    }

    fn is_empty_element(&self) -> bool {
        if self.paragraph_style.is_some() {
            return false;
        }
        if self.keep_next.is_some() {
            return false;
        }
        if self.keep_lines.is_some() {
            return false;
        }
        if self.page_break_before.is_some() {
            return false;
        }
        if self.frame_pr.is_some() {
            return false;
        }
        if self.widow_control.is_some() {
            return false;
        }
        if self.num_pr.is_some() {
            return false;
        }
        if self.suppress_line_numbers.is_some() {
            return false;
        }
        if self.paragraph_border.is_some() {
            return false;
        }
        if self.shading.is_some() {
            return false;
        }
        if self.tabs.is_some() {
            return false;
        }
        if self.suppress_auto_hyphens.is_some() {
            return false;
        }
        if self.kinsoku.is_some() {
            return false;
        }
        if self.word_wrap.is_some() {
            return false;
        }
        if self.overflow_punct.is_some() {
            return false;
        }
        if self.top_line_punct.is_some() {
            return false;
        }
        if self.auto_space_d_e.is_some() {
            return false;
        }
        if self.auto_space_d_n.is_some() {
            return false;
        }
        if self.bidi.is_some() {
            return false;
        }
        if self.adjust_right_ind.is_some() {
            return false;
        }
        if self.snap_to_grid.is_some() {
            return false;
        }
        if self.spacing.is_some() {
            return false;
        }
        if self.indentation.is_some() {
            return false;
        }
        if self.contextual_spacing.is_some() {
            return false;
        }
        if self.mirror_indents.is_some() {
            return false;
        }
        if self.suppress_overlap.is_some() {
            return false;
        }
        if self.justification.is_some() {
            return false;
        }
        if self.text_direction.is_some() {
            return false;
        }
        if self.text_alignment.is_some() {
            return false;
        }
        if self.textbox_tight_wrap.is_some() {
            return false;
        }
        if self.outline_lvl.is_some() {
            return false;
        }
        if self.div_id.is_some() {
            return false;
        }
        if self.cnf_style.is_some() {
            return false;
        }
        if self.r_pr.is_some() {
            return false;
        }
        if self.sect_pr.is_some() {
            return false;
        }
        if self.p_pr_change.is_some() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for CTPPrBase {
    fn write_children<W: Write>(&self, writer: &mut Writer<W>) -> Result<(), SerializeError> {
        if let Some(ref val) = self.paragraph_style {
            val.write_element("w:pStyle", writer)?;
        }
        if let Some(ref val) = self.keep_next {
            val.write_element("w:keepNext", writer)?;
        }
        if let Some(ref val) = self.keep_lines {
            val.write_element("w:keepLines", writer)?;
        }
        if let Some(ref val) = self.page_break_before {
            val.write_element("w:pageBreakBefore", writer)?;
        }
        if let Some(ref val) = self.frame_pr {
            val.write_element("w:framePr", writer)?;
        }
        if let Some(ref val) = self.widow_control {
            val.write_element("w:widowControl", writer)?;
        }
        if let Some(ref val) = self.num_pr {
            val.write_element("w:numPr", writer)?;
        }
        if let Some(ref val) = self.suppress_line_numbers {
            val.write_element("w:suppressLineNumbers", writer)?;
        }
        if let Some(ref val) = self.paragraph_border {
            val.write_element("w:pBdr", writer)?;
        }
        if let Some(ref val) = self.shading {
            val.write_element("w:shd", writer)?;
        }
        if let Some(ref val) = self.tabs {
            val.write_element("w:tabs", writer)?;
        }
        if let Some(ref val) = self.suppress_auto_hyphens {
            val.write_element("w:suppressAutoHyphens", writer)?;
        }
        if let Some(ref val) = self.kinsoku {
            val.write_element("w:kinsoku", writer)?;
        }
        if let Some(ref val) = self.word_wrap {
            val.write_element("w:wordWrap", writer)?;
        }
        if let Some(ref val) = self.overflow_punct {
            val.write_element("w:overflowPunct", writer)?;
        }
        if let Some(ref val) = self.top_line_punct {
            val.write_element("w:topLinePunct", writer)?;
        }
        if let Some(ref val) = self.auto_space_d_e {
            val.write_element("w:autoSpaceDE", writer)?;
        }
        if let Some(ref val) = self.auto_space_d_n {
            val.write_element("w:autoSpaceDN", writer)?;
        }
        if let Some(ref val) = self.bidi {
            val.write_element("w:bidi", writer)?;
        }
        if let Some(ref val) = self.adjust_right_ind {
            val.write_element("w:adjustRightInd", writer)?;
        }
        if let Some(ref val) = self.snap_to_grid {
            val.write_element("w:snapToGrid", writer)?;
        }
        if let Some(ref val) = self.spacing {
            val.write_element("w:spacing", writer)?;
        }
        if let Some(ref val) = self.indentation {
            val.write_element("w:ind", writer)?;
        }
        if let Some(ref val) = self.contextual_spacing {
            val.write_element("w:contextualSpacing", writer)?;
        }
        if let Some(ref val) = self.mirror_indents {
            val.write_element("w:mirrorIndents", writer)?;
        }
        if let Some(ref val) = self.suppress_overlap {
            val.write_element("w:suppressOverlap", writer)?;
        }
        if let Some(ref val) = self.justification {
            val.write_element("w:jc", writer)?;
        }
        if let Some(ref val) = self.text_direction {
            val.write_element("w:textDirection", writer)?;
        }
        if let Some(ref val) = self.text_alignment {
            val.write_element("w:textAlignment", writer)?;
        }
        if let Some(ref val) = self.textbox_tight_wrap {
            val.write_element("w:textboxTightWrap", writer)?;
        }
        if let Some(ref val) = self.outline_lvl {
            val.write_element("w:outlineLvl", writer)?;
        }
        if let Some(ref val) = self.div_id {
            val.write_element("w:divId", writer)?;
        }
        if let Some(ref val) = self.cnf_style {
            val.write_element("w:cnfStyle", writer)?;
        }
        #[cfg(feature = "extra-children")]
        for child in &self.extra_children {
            child.write_to(writer).map_err(SerializeError::from)?;
        }
        Ok(())
    }

    fn is_empty_element(&self) -> bool {
        if self.paragraph_style.is_some() {
            return false;
        }
        if self.keep_next.is_some() {
            return false;
        }
        if self.keep_lines.is_some() {
            return false;
        }
        if self.page_break_before.is_some() {
            return false;
        }
        if self.frame_pr.is_some() {
            return false;
        }
        if self.widow_control.is_some() {
            return false;
        }
        if self.num_pr.is_some() {
            return false;
        }
        if self.suppress_line_numbers.is_some() {
            return false;
        }
        if self.paragraph_border.is_some() {
            return false;
        }
        if self.shading.is_some() {
            return false;
        }
        if self.tabs.is_some() {
            return false;
        }
        if self.suppress_auto_hyphens.is_some() {
            return false;
        }
        if self.kinsoku.is_some() {
            return false;
        }
        if self.word_wrap.is_some() {
            return false;
        }
        if self.overflow_punct.is_some() {
            return false;
        }
        if self.top_line_punct.is_some() {
            return false;
        }
        if self.auto_space_d_e.is_some() {
            return false;
        }
        if self.auto_space_d_n.is_some() {
            return false;
        }
        if self.bidi.is_some() {
            return false;
        }
        if self.adjust_right_ind.is_some() {
            return false;
        }
        if self.snap_to_grid.is_some() {
            return false;
        }
        if self.spacing.is_some() {
            return false;
        }
        if self.indentation.is_some() {
            return false;
        }
        if self.contextual_spacing.is_some() {
            return false;
        }
        if self.mirror_indents.is_some() {
            return false;
        }
        if self.suppress_overlap.is_some() {
            return false;
        }
        if self.justification.is_some() {
            return false;
        }
        if self.text_direction.is_some() {
            return false;
        }
        if self.text_alignment.is_some() {
            return false;
        }
        if self.textbox_tight_wrap.is_some() {
            return false;
        }
        if self.outline_lvl.is_some() {
            return false;
        }
        if self.div_id.is_some() {
            return false;
        }
        if self.cnf_style.is_some() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for CTPPrGeneral {
    fn write_children<W: Write>(&self, writer: &mut Writer<W>) -> Result<(), SerializeError> {
        if let Some(ref val) = self.paragraph_style {
            val.write_element("w:pStyle", writer)?;
        }
        if let Some(ref val) = self.keep_next {
            val.write_element("w:keepNext", writer)?;
        }
        if let Some(ref val) = self.keep_lines {
            val.write_element("w:keepLines", writer)?;
        }
        if let Some(ref val) = self.page_break_before {
            val.write_element("w:pageBreakBefore", writer)?;
        }
        if let Some(ref val) = self.frame_pr {
            val.write_element("w:framePr", writer)?;
        }
        if let Some(ref val) = self.widow_control {
            val.write_element("w:widowControl", writer)?;
        }
        if let Some(ref val) = self.num_pr {
            val.write_element("w:numPr", writer)?;
        }
        if let Some(ref val) = self.suppress_line_numbers {
            val.write_element("w:suppressLineNumbers", writer)?;
        }
        if let Some(ref val) = self.paragraph_border {
            val.write_element("w:pBdr", writer)?;
        }
        if let Some(ref val) = self.shading {
            val.write_element("w:shd", writer)?;
        }
        if let Some(ref val) = self.tabs {
            val.write_element("w:tabs", writer)?;
        }
        if let Some(ref val) = self.suppress_auto_hyphens {
            val.write_element("w:suppressAutoHyphens", writer)?;
        }
        if let Some(ref val) = self.kinsoku {
            val.write_element("w:kinsoku", writer)?;
        }
        if let Some(ref val) = self.word_wrap {
            val.write_element("w:wordWrap", writer)?;
        }
        if let Some(ref val) = self.overflow_punct {
            val.write_element("w:overflowPunct", writer)?;
        }
        if let Some(ref val) = self.top_line_punct {
            val.write_element("w:topLinePunct", writer)?;
        }
        if let Some(ref val) = self.auto_space_d_e {
            val.write_element("w:autoSpaceDE", writer)?;
        }
        if let Some(ref val) = self.auto_space_d_n {
            val.write_element("w:autoSpaceDN", writer)?;
        }
        if let Some(ref val) = self.bidi {
            val.write_element("w:bidi", writer)?;
        }
        if let Some(ref val) = self.adjust_right_ind {
            val.write_element("w:adjustRightInd", writer)?;
        }
        if let Some(ref val) = self.snap_to_grid {
            val.write_element("w:snapToGrid", writer)?;
        }
        if let Some(ref val) = self.spacing {
            val.write_element("w:spacing", writer)?;
        }
        if let Some(ref val) = self.indentation {
            val.write_element("w:ind", writer)?;
        }
        if let Some(ref val) = self.contextual_spacing {
            val.write_element("w:contextualSpacing", writer)?;
        }
        if let Some(ref val) = self.mirror_indents {
            val.write_element("w:mirrorIndents", writer)?;
        }
        if let Some(ref val) = self.suppress_overlap {
            val.write_element("w:suppressOverlap", writer)?;
        }
        if let Some(ref val) = self.justification {
            val.write_element("w:jc", writer)?;
        }
        if let Some(ref val) = self.text_direction {
            val.write_element("w:textDirection", writer)?;
        }
        if let Some(ref val) = self.text_alignment {
            val.write_element("w:textAlignment", writer)?;
        }
        if let Some(ref val) = self.textbox_tight_wrap {
            val.write_element("w:textboxTightWrap", writer)?;
        }
        if let Some(ref val) = self.outline_lvl {
            val.write_element("w:outlineLvl", writer)?;
        }
        if let Some(ref val) = self.div_id {
            val.write_element("w:divId", writer)?;
        }
        if let Some(ref val) = self.cnf_style {
            val.write_element("w:cnfStyle", writer)?;
        }
        if let Some(ref val) = self.p_pr_change {
            val.write_element("w:pPrChange", writer)?;
        }
        #[cfg(feature = "extra-children")]
        for child in &self.extra_children {
            child.write_to(writer).map_err(SerializeError::from)?;
        }
        Ok(())
    }

    fn is_empty_element(&self) -> bool {
        if self.paragraph_style.is_some() {
            return false;
        }
        if self.keep_next.is_some() {
            return false;
        }
        if self.keep_lines.is_some() {
            return false;
        }
        if self.page_break_before.is_some() {
            return false;
        }
        if self.frame_pr.is_some() {
            return false;
        }
        if self.widow_control.is_some() {
            return false;
        }
        if self.num_pr.is_some() {
            return false;
        }
        if self.suppress_line_numbers.is_some() {
            return false;
        }
        if self.paragraph_border.is_some() {
            return false;
        }
        if self.shading.is_some() {
            return false;
        }
        if self.tabs.is_some() {
            return false;
        }
        if self.suppress_auto_hyphens.is_some() {
            return false;
        }
        if self.kinsoku.is_some() {
            return false;
        }
        if self.word_wrap.is_some() {
            return false;
        }
        if self.overflow_punct.is_some() {
            return false;
        }
        if self.top_line_punct.is_some() {
            return false;
        }
        if self.auto_space_d_e.is_some() {
            return false;
        }
        if self.auto_space_d_n.is_some() {
            return false;
        }
        if self.bidi.is_some() {
            return false;
        }
        if self.adjust_right_ind.is_some() {
            return false;
        }
        if self.snap_to_grid.is_some() {
            return false;
        }
        if self.spacing.is_some() {
            return false;
        }
        if self.indentation.is_some() {
            return false;
        }
        if self.contextual_spacing.is_some() {
            return false;
        }
        if self.mirror_indents.is_some() {
            return false;
        }
        if self.suppress_overlap.is_some() {
            return false;
        }
        if self.justification.is_some() {
            return false;
        }
        if self.text_direction.is_some() {
            return false;
        }
        if self.text_alignment.is_some() {
            return false;
        }
        if self.textbox_tight_wrap.is_some() {
            return false;
        }
        if self.outline_lvl.is_some() {
            return false;
        }
        if self.div_id.is_some() {
            return false;
        }
        if self.cnf_style.is_some() {
            return false;
        }
        if self.p_pr_change.is_some() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for CTControl {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        if let Some(ref val) = self.name {
            start.push_attribute(("w:name", val.as_str()));
        }
        if let Some(ref val) = self.shapeid {
            start.push_attribute(("w:shapeid", val.as_str()));
        }
        #[cfg(feature = "extra-attrs")]
        for (key, value) in &self.extra_attrs {
            start.push_attribute((key.as_str(), value.as_str()));
        }
        start
    }

    fn is_empty_element(&self) -> bool {
        true
    }
}

impl ToXml for CTBackground {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        if let Some(ref val) = self.color {
            {
                let s = val.to_string();
                start.push_attribute(("w:color", s.as_str()));
            }
        }
        if let Some(ref val) = self.theme_color {
            {
                let s = val.to_string();
                start.push_attribute(("w:themeColor", s.as_str()));
            }
        }
        if let Some(ref val) = self.theme_tint {
            {
                let hex = encode_hex(val);
                start.push_attribute(("w:themeTint", hex.as_str()));
            }
        }
        if let Some(ref val) = self.theme_shade {
            {
                let hex = encode_hex(val);
                start.push_attribute(("w:themeShade", hex.as_str()));
            }
        }
        #[cfg(feature = "extra-attrs")]
        for (key, value) in &self.extra_attrs {
            start.push_attribute((key.as_str(), value.as_str()));
        }
        start
    }

    fn write_children<W: Write>(&self, writer: &mut Writer<W>) -> Result<(), SerializeError> {
        if let Some(ref val) = self.drawing {
            val.write_element("w:drawing", writer)?;
        }
        #[cfg(feature = "extra-children")]
        for child in &self.extra_children {
            child.write_to(writer).map_err(SerializeError::from)?;
        }
        Ok(())
    }

    fn is_empty_element(&self) -> bool {
        if self.drawing.is_some() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for CTRel {
    fn write_children<W: Write>(&self, writer: &mut Writer<W>) -> Result<(), SerializeError> {
        #[cfg(feature = "extra-children")]
        for child in &self.extra_children {
            child.write_to(writer).map_err(SerializeError::from)?;
        }
        Ok(())
    }

    fn is_empty_element(&self) -> bool {
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for CTObject {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        if let Some(ref val) = self.dxa_orig {
            {
                let s = val.to_string();
                start.push_attribute(("w:dxaOrig", s.as_str()));
            }
        }
        if let Some(ref val) = self.dya_orig {
            {
                let s = val.to_string();
                start.push_attribute(("w:dyaOrig", s.as_str()));
            }
        }
        #[cfg(feature = "extra-attrs")]
        for (key, value) in &self.extra_attrs {
            start.push_attribute((key.as_str(), value.as_str()));
        }
        start
    }

    fn write_children<W: Write>(&self, writer: &mut Writer<W>) -> Result<(), SerializeError> {
        if let Some(ref val) = self.drawing {
            val.write_element("w:drawing", writer)?;
        }
        #[cfg(feature = "extra-children")]
        for child in &self.extra_children {
            child.write_to(writer).map_err(SerializeError::from)?;
        }
        Ok(())
    }

    fn is_empty_element(&self) -> bool {
        if self.drawing.is_some() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for CTPicture {
    fn write_children<W: Write>(&self, writer: &mut Writer<W>) -> Result<(), SerializeError> {
        if let Some(ref val) = self.movie {
            val.write_element("w:movie", writer)?;
        }
        if let Some(ref val) = self.control {
            val.write_element("w:control", writer)?;
        }
        #[cfg(feature = "extra-children")]
        for child in &self.extra_children {
            child.write_to(writer).map_err(SerializeError::from)?;
        }
        Ok(())
    }

    fn is_empty_element(&self) -> bool {
        if self.movie.is_some() {
            return false;
        }
        if self.control.is_some() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for CTObjectEmbed {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        if let Some(ref val) = self.draw_aspect {
            {
                let s = val.to_string();
                start.push_attribute(("w:drawAspect", s.as_str()));
            }
        }
        if let Some(ref val) = self.prog_id {
            start.push_attribute(("w:progId", val.as_str()));
        }
        if let Some(ref val) = self.shape_id {
            start.push_attribute(("w:shapeId", val.as_str()));
        }
        if let Some(ref val) = self.field_codes {
            start.push_attribute(("w:fieldCodes", val.as_str()));
        }
        #[cfg(feature = "extra-attrs")]
        for (key, value) in &self.extra_attrs {
            start.push_attribute((key.as_str(), value.as_str()));
        }
        start
    }

    fn is_empty_element(&self) -> bool {
        true
    }
}

impl ToXml for CTObjectLink {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        if let Some(ref val) = self.draw_aspect {
            {
                let s = val.to_string();
                start.push_attribute(("w:drawAspect", s.as_str()));
            }
        }
        if let Some(ref val) = self.prog_id {
            start.push_attribute(("w:progId", val.as_str()));
        }
        if let Some(ref val) = self.shape_id {
            start.push_attribute(("w:shapeId", val.as_str()));
        }
        if let Some(ref val) = self.field_codes {
            start.push_attribute(("w:fieldCodes", val.as_str()));
        }
        {
            let val = &self.update_mode;
            {
                let s = val.to_string();
                start.push_attribute(("w:updateMode", s.as_str()));
            }
        }
        if let Some(ref val) = self.locked_field {
            {
                let s = val.to_string();
                start.push_attribute(("w:lockedField", s.as_str()));
            }
        }
        #[cfg(feature = "extra-attrs")]
        for (key, value) in &self.extra_attrs {
            start.push_attribute((key.as_str(), value.as_str()));
        }
        start
    }

    fn is_empty_element(&self) -> bool {
        true
    }
}

impl ToXml for CTDrawing {
    fn write_children<W: Write>(&self, writer: &mut Writer<W>) -> Result<(), SerializeError> {
        #[cfg(feature = "extra-children")]
        for child in &self.extra_children {
            child.write_to(writer).map_err(SerializeError::from)?;
        }
        Ok(())
    }

    fn is_empty_element(&self) -> bool {
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for CTSimpleField {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        {
            let val = &self.instr;
            start.push_attribute(("w:instr", val.as_str()));
        }
        if let Some(ref val) = self.fld_lock {
            {
                let s = val.to_string();
                start.push_attribute(("w:fldLock", s.as_str()));
            }
        }
        if let Some(ref val) = self.dirty {
            {
                let s = val.to_string();
                start.push_attribute(("w:dirty", s.as_str()));
            }
        }
        #[cfg(feature = "extra-attrs")]
        for (key, value) in &self.extra_attrs {
            start.push_attribute((key.as_str(), value.as_str()));
        }
        start
    }

    fn write_children<W: Write>(&self, writer: &mut Writer<W>) -> Result<(), SerializeError> {
        if let Some(ref val) = self.fld_data {
            val.write_element("w:fldData", writer)?;
        }
        for item in &self.p_content {
            item.write_element("", writer)?;
        }
        #[cfg(feature = "extra-children")]
        for child in &self.extra_children {
            child.write_to(writer).map_err(SerializeError::from)?;
        }
        Ok(())
    }

    fn is_empty_element(&self) -> bool {
        if self.fld_data.is_some() {
            return false;
        }
        if !self.p_content.is_empty() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for CTFFTextType {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        {
            let val = &self.value;
            {
                let s = val.to_string();
                start.push_attribute(("w:val", s.as_str()));
            }
        }
        #[cfg(feature = "extra-attrs")]
        for (key, value) in &self.extra_attrs {
            start.push_attribute((key.as_str(), value.as_str()));
        }
        start
    }

    fn is_empty_element(&self) -> bool {
        true
    }
}

impl ToXml for CTFFName {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        if let Some(ref val) = self.value {
            start.push_attribute(("w:val", val.as_str()));
        }
        #[cfg(feature = "extra-attrs")]
        for (key, value) in &self.extra_attrs {
            start.push_attribute((key.as_str(), value.as_str()));
        }
        start
    }

    fn is_empty_element(&self) -> bool {
        true
    }
}

impl ToXml for CTFldChar {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        {
            let val = &self.fld_char_type;
            {
                let s = val.to_string();
                start.push_attribute(("w:fldCharType", s.as_str()));
            }
        }
        if let Some(ref val) = self.fld_lock {
            {
                let s = val.to_string();
                start.push_attribute(("w:fldLock", s.as_str()));
            }
        }
        if let Some(ref val) = self.dirty {
            {
                let s = val.to_string();
                start.push_attribute(("w:dirty", s.as_str()));
            }
        }
        #[cfg(feature = "extra-attrs")]
        for (key, value) in &self.extra_attrs {
            start.push_attribute((key.as_str(), value.as_str()));
        }
        start
    }

    fn is_empty_element(&self) -> bool {
        true
    }
}

impl ToXml for Hyperlink {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        if let Some(ref val) = self.tgt_frame {
            start.push_attribute(("w:tgtFrame", val.as_str()));
        }
        if let Some(ref val) = self.tooltip {
            start.push_attribute(("w:tooltip", val.as_str()));
        }
        if let Some(ref val) = self.doc_location {
            start.push_attribute(("w:docLocation", val.as_str()));
        }
        if let Some(ref val) = self.history {
            {
                let s = val.to_string();
                start.push_attribute(("w:history", s.as_str()));
            }
        }
        if let Some(ref val) = self.anchor {
            start.push_attribute(("w:anchor", val.as_str()));
        }
        #[cfg(feature = "extra-attrs")]
        for (key, value) in &self.extra_attrs {
            start.push_attribute((key.as_str(), value.as_str()));
        }
        start
    }

    fn write_children<W: Write>(&self, writer: &mut Writer<W>) -> Result<(), SerializeError> {
        for item in &self.p_content {
            item.write_element("", writer)?;
        }
        #[cfg(feature = "extra-children")]
        for child in &self.extra_children {
            child.write_to(writer).map_err(SerializeError::from)?;
        }
        Ok(())
    }

    fn is_empty_element(&self) -> bool {
        if !self.p_content.is_empty() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for CTFFData {
    fn write_children<W: Write>(&self, writer: &mut Writer<W>) -> Result<(), SerializeError> {
        #[cfg(feature = "extra-children")]
        for child in &self.extra_children {
            child.write_to(writer).map_err(SerializeError::from)?;
        }
        Ok(())
    }

    fn is_empty_element(&self) -> bool {
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for CTFFHelpText {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        if let Some(ref val) = self.r#type {
            {
                let s = val.to_string();
                start.push_attribute(("w:type", s.as_str()));
            }
        }
        if let Some(ref val) = self.value {
            start.push_attribute(("w:val", val.as_str()));
        }
        #[cfg(feature = "extra-attrs")]
        for (key, value) in &self.extra_attrs {
            start.push_attribute((key.as_str(), value.as_str()));
        }
        start
    }

    fn is_empty_element(&self) -> bool {
        true
    }
}

impl ToXml for CTFFStatusText {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        if let Some(ref val) = self.r#type {
            {
                let s = val.to_string();
                start.push_attribute(("w:type", s.as_str()));
            }
        }
        if let Some(ref val) = self.value {
            start.push_attribute(("w:val", val.as_str()));
        }
        #[cfg(feature = "extra-attrs")]
        for (key, value) in &self.extra_attrs {
            start.push_attribute((key.as_str(), value.as_str()));
        }
        start
    }

    fn is_empty_element(&self) -> bool {
        true
    }
}

impl ToXml for CTFFCheckBox {
    fn write_children<W: Write>(&self, writer: &mut Writer<W>) -> Result<(), SerializeError> {
        if let Some(ref val) = self.default {
            val.write_element("w:default", writer)?;
        }
        if let Some(ref val) = self.checked {
            val.write_element("w:checked", writer)?;
        }
        #[cfg(feature = "extra-children")]
        for child in &self.extra_children {
            child.write_to(writer).map_err(SerializeError::from)?;
        }
        Ok(())
    }

    fn is_empty_element(&self) -> bool {
        if self.default.is_some() {
            return false;
        }
        if self.checked.is_some() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for CTFFDDList {
    fn write_children<W: Write>(&self, writer: &mut Writer<W>) -> Result<(), SerializeError> {
        if let Some(ref val) = self.result {
            val.write_element("w:result", writer)?;
        }
        if let Some(ref val) = self.default {
            val.write_element("w:default", writer)?;
        }
        for item in &self.list_entry {
            item.write_element("w:listEntry", writer)?;
        }
        #[cfg(feature = "extra-children")]
        for child in &self.extra_children {
            child.write_to(writer).map_err(SerializeError::from)?;
        }
        Ok(())
    }

    fn is_empty_element(&self) -> bool {
        if self.result.is_some() {
            return false;
        }
        if self.default.is_some() {
            return false;
        }
        if !self.list_entry.is_empty() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for CTFFTextInput {
    fn write_children<W: Write>(&self, writer: &mut Writer<W>) -> Result<(), SerializeError> {
        if let Some(ref val) = self.r#type {
            val.write_element("w:type", writer)?;
        }
        if let Some(ref val) = self.default {
            val.write_element("w:default", writer)?;
        }
        if let Some(ref val) = self.max_length {
            val.write_element("w:maxLength", writer)?;
        }
        if let Some(ref val) = self.format {
            val.write_element("w:format", writer)?;
        }
        #[cfg(feature = "extra-children")]
        for child in &self.extra_children {
            child.write_to(writer).map_err(SerializeError::from)?;
        }
        Ok(())
    }

    fn is_empty_element(&self) -> bool {
        if self.r#type.is_some() {
            return false;
        }
        if self.default.is_some() {
            return false;
        }
        if self.max_length.is_some() {
            return false;
        }
        if self.format.is_some() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for CTSectType {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        if let Some(ref val) = self.value {
            {
                let s = val.to_string();
                start.push_attribute(("w:val", s.as_str()));
            }
        }
        #[cfg(feature = "extra-attrs")]
        for (key, value) in &self.extra_attrs {
            start.push_attribute((key.as_str(), value.as_str()));
        }
        start
    }

    fn is_empty_element(&self) -> bool {
        true
    }
}

impl ToXml for CTPaperSource {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        if let Some(ref val) = self.first {
            {
                let s = val.to_string();
                start.push_attribute(("w:first", s.as_str()));
            }
        }
        if let Some(ref val) = self.other {
            {
                let s = val.to_string();
                start.push_attribute(("w:other", s.as_str()));
            }
        }
        #[cfg(feature = "extra-attrs")]
        for (key, value) in &self.extra_attrs {
            start.push_attribute((key.as_str(), value.as_str()));
        }
        start
    }

    fn is_empty_element(&self) -> bool {
        true
    }
}

impl ToXml for PageSize {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        if let Some(ref val) = self.width {
            {
                let s = val.to_string();
                start.push_attribute(("w:w", s.as_str()));
            }
        }
        if let Some(ref val) = self.height {
            {
                let s = val.to_string();
                start.push_attribute(("w:h", s.as_str()));
            }
        }
        if let Some(ref val) = self.orient {
            {
                let s = val.to_string();
                start.push_attribute(("w:orient", s.as_str()));
            }
        }
        if let Some(ref val) = self.code {
            {
                let s = val.to_string();
                start.push_attribute(("w:code", s.as_str()));
            }
        }
        #[cfg(feature = "extra-attrs")]
        for (key, value) in &self.extra_attrs {
            start.push_attribute((key.as_str(), value.as_str()));
        }
        start
    }

    fn is_empty_element(&self) -> bool {
        true
    }
}

impl ToXml for PageMargins {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        {
            let val = &self.top;
            {
                let s = val.to_string();
                start.push_attribute(("w:top", s.as_str()));
            }
        }
        {
            let val = &self.right;
            {
                let s = val.to_string();
                start.push_attribute(("w:right", s.as_str()));
            }
        }
        {
            let val = &self.bottom;
            {
                let s = val.to_string();
                start.push_attribute(("w:bottom", s.as_str()));
            }
        }
        {
            let val = &self.left;
            {
                let s = val.to_string();
                start.push_attribute(("w:left", s.as_str()));
            }
        }
        {
            let val = &self.header;
            {
                let s = val.to_string();
                start.push_attribute(("w:header", s.as_str()));
            }
        }
        {
            let val = &self.footer;
            {
                let s = val.to_string();
                start.push_attribute(("w:footer", s.as_str()));
            }
        }
        {
            let val = &self.gutter;
            {
                let s = val.to_string();
                start.push_attribute(("w:gutter", s.as_str()));
            }
        }
        #[cfg(feature = "extra-attrs")]
        for (key, value) in &self.extra_attrs {
            start.push_attribute((key.as_str(), value.as_str()));
        }
        start
    }

    fn is_empty_element(&self) -> bool {
        true
    }
}

impl ToXml for CTPageBorders {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        if let Some(ref val) = self.z_order {
            {
                let s = val.to_string();
                start.push_attribute(("w:zOrder", s.as_str()));
            }
        }
        if let Some(ref val) = self.display {
            {
                let s = val.to_string();
                start.push_attribute(("w:display", s.as_str()));
            }
        }
        if let Some(ref val) = self.offset_from {
            {
                let s = val.to_string();
                start.push_attribute(("w:offsetFrom", s.as_str()));
            }
        }
        #[cfg(feature = "extra-attrs")]
        for (key, value) in &self.extra_attrs {
            start.push_attribute((key.as_str(), value.as_str()));
        }
        start
    }

    fn write_children<W: Write>(&self, writer: &mut Writer<W>) -> Result<(), SerializeError> {
        if let Some(ref val) = self.top {
            val.write_element("w:top", writer)?;
        }
        if let Some(ref val) = self.left {
            val.write_element("w:left", writer)?;
        }
        if let Some(ref val) = self.bottom {
            val.write_element("w:bottom", writer)?;
        }
        if let Some(ref val) = self.right {
            val.write_element("w:right", writer)?;
        }
        #[cfg(feature = "extra-children")]
        for child in &self.extra_children {
            child.write_to(writer).map_err(SerializeError::from)?;
        }
        Ok(())
    }

    fn is_empty_element(&self) -> bool {
        if self.top.is_some() {
            return false;
        }
        if self.left.is_some() {
            return false;
        }
        if self.bottom.is_some() {
            return false;
        }
        if self.right.is_some() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for CTPageBorder {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        {
            let val = &self.value;
            {
                let s = val.to_string();
                start.push_attribute(("w:val", s.as_str()));
            }
        }
        if let Some(ref val) = self.color {
            {
                let s = val.to_string();
                start.push_attribute(("w:color", s.as_str()));
            }
        }
        if let Some(ref val) = self.theme_color {
            {
                let s = val.to_string();
                start.push_attribute(("w:themeColor", s.as_str()));
            }
        }
        if let Some(ref val) = self.theme_tint {
            {
                let hex = encode_hex(val);
                start.push_attribute(("w:themeTint", hex.as_str()));
            }
        }
        if let Some(ref val) = self.theme_shade {
            {
                let hex = encode_hex(val);
                start.push_attribute(("w:themeShade", hex.as_str()));
            }
        }
        if let Some(ref val) = self.size {
            {
                let s = val.to_string();
                start.push_attribute(("w:sz", s.as_str()));
            }
        }
        if let Some(ref val) = self.space {
            {
                let s = val.to_string();
                start.push_attribute(("w:space", s.as_str()));
            }
        }
        if let Some(ref val) = self.shadow {
            {
                let s = val.to_string();
                start.push_attribute(("w:shadow", s.as_str()));
            }
        }
        if let Some(ref val) = self.frame {
            {
                let s = val.to_string();
                start.push_attribute(("w:frame", s.as_str()));
            }
        }
        #[cfg(feature = "extra-attrs")]
        for (key, value) in &self.extra_attrs {
            start.push_attribute((key.as_str(), value.as_str()));
        }
        start
    }

    fn is_empty_element(&self) -> bool {
        true
    }
}

impl ToXml for CTBottomPageBorder {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        {
            let val = &self.value;
            {
                let s = val.to_string();
                start.push_attribute(("w:val", s.as_str()));
            }
        }
        if let Some(ref val) = self.color {
            {
                let s = val.to_string();
                start.push_attribute(("w:color", s.as_str()));
            }
        }
        if let Some(ref val) = self.theme_color {
            {
                let s = val.to_string();
                start.push_attribute(("w:themeColor", s.as_str()));
            }
        }
        if let Some(ref val) = self.theme_tint {
            {
                let hex = encode_hex(val);
                start.push_attribute(("w:themeTint", hex.as_str()));
            }
        }
        if let Some(ref val) = self.theme_shade {
            {
                let hex = encode_hex(val);
                start.push_attribute(("w:themeShade", hex.as_str()));
            }
        }
        if let Some(ref val) = self.size {
            {
                let s = val.to_string();
                start.push_attribute(("w:sz", s.as_str()));
            }
        }
        if let Some(ref val) = self.space {
            {
                let s = val.to_string();
                start.push_attribute(("w:space", s.as_str()));
            }
        }
        if let Some(ref val) = self.shadow {
            {
                let s = val.to_string();
                start.push_attribute(("w:shadow", s.as_str()));
            }
        }
        if let Some(ref val) = self.frame {
            {
                let s = val.to_string();
                start.push_attribute(("w:frame", s.as_str()));
            }
        }
        #[cfg(feature = "extra-attrs")]
        for (key, value) in &self.extra_attrs {
            start.push_attribute((key.as_str(), value.as_str()));
        }
        start
    }

    fn is_empty_element(&self) -> bool {
        true
    }
}

impl ToXml for CTTopPageBorder {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        {
            let val = &self.value;
            {
                let s = val.to_string();
                start.push_attribute(("w:val", s.as_str()));
            }
        }
        if let Some(ref val) = self.color {
            {
                let s = val.to_string();
                start.push_attribute(("w:color", s.as_str()));
            }
        }
        if let Some(ref val) = self.theme_color {
            {
                let s = val.to_string();
                start.push_attribute(("w:themeColor", s.as_str()));
            }
        }
        if let Some(ref val) = self.theme_tint {
            {
                let hex = encode_hex(val);
                start.push_attribute(("w:themeTint", hex.as_str()));
            }
        }
        if let Some(ref val) = self.theme_shade {
            {
                let hex = encode_hex(val);
                start.push_attribute(("w:themeShade", hex.as_str()));
            }
        }
        if let Some(ref val) = self.size {
            {
                let s = val.to_string();
                start.push_attribute(("w:sz", s.as_str()));
            }
        }
        if let Some(ref val) = self.space {
            {
                let s = val.to_string();
                start.push_attribute(("w:space", s.as_str()));
            }
        }
        if let Some(ref val) = self.shadow {
            {
                let s = val.to_string();
                start.push_attribute(("w:shadow", s.as_str()));
            }
        }
        if let Some(ref val) = self.frame {
            {
                let s = val.to_string();
                start.push_attribute(("w:frame", s.as_str()));
            }
        }
        #[cfg(feature = "extra-attrs")]
        for (key, value) in &self.extra_attrs {
            start.push_attribute((key.as_str(), value.as_str()));
        }
        start
    }

    fn is_empty_element(&self) -> bool {
        true
    }
}

impl ToXml for CTLineNumber {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        if let Some(ref val) = self.count_by {
            {
                let s = val.to_string();
                start.push_attribute(("w:countBy", s.as_str()));
            }
        }
        if let Some(ref val) = self.start {
            {
                let s = val.to_string();
                start.push_attribute(("w:start", s.as_str()));
            }
        }
        if let Some(ref val) = self.distance {
            {
                let s = val.to_string();
                start.push_attribute(("w:distance", s.as_str()));
            }
        }
        if let Some(ref val) = self.restart {
            {
                let s = val.to_string();
                start.push_attribute(("w:restart", s.as_str()));
            }
        }
        #[cfg(feature = "extra-attrs")]
        for (key, value) in &self.extra_attrs {
            start.push_attribute((key.as_str(), value.as_str()));
        }
        start
    }

    fn is_empty_element(&self) -> bool {
        true
    }
}

impl ToXml for CTPageNumber {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        if let Some(ref val) = self.fmt {
            {
                let s = val.to_string();
                start.push_attribute(("w:fmt", s.as_str()));
            }
        }
        if let Some(ref val) = self.start {
            {
                let s = val.to_string();
                start.push_attribute(("w:start", s.as_str()));
            }
        }
        if let Some(ref val) = self.chap_style {
            {
                let s = val.to_string();
                start.push_attribute(("w:chapStyle", s.as_str()));
            }
        }
        if let Some(ref val) = self.chap_sep {
            {
                let s = val.to_string();
                start.push_attribute(("w:chapSep", s.as_str()));
            }
        }
        #[cfg(feature = "extra-attrs")]
        for (key, value) in &self.extra_attrs {
            start.push_attribute((key.as_str(), value.as_str()));
        }
        start
    }

    fn is_empty_element(&self) -> bool {
        true
    }
}

impl ToXml for CTColumn {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        if let Some(ref val) = self.width {
            {
                let s = val.to_string();
                start.push_attribute(("w:w", s.as_str()));
            }
        }
        if let Some(ref val) = self.space {
            {
                let s = val.to_string();
                start.push_attribute(("w:space", s.as_str()));
            }
        }
        #[cfg(feature = "extra-attrs")]
        for (key, value) in &self.extra_attrs {
            start.push_attribute((key.as_str(), value.as_str()));
        }
        start
    }

    fn is_empty_element(&self) -> bool {
        true
    }
}

impl ToXml for Columns {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        if let Some(ref val) = self.equal_width {
            {
                let s = val.to_string();
                start.push_attribute(("w:equalWidth", s.as_str()));
            }
        }
        if let Some(ref val) = self.space {
            {
                let s = val.to_string();
                start.push_attribute(("w:space", s.as_str()));
            }
        }
        if let Some(ref val) = self.num {
            {
                let s = val.to_string();
                start.push_attribute(("w:num", s.as_str()));
            }
        }
        if let Some(ref val) = self.sep {
            {
                let s = val.to_string();
                start.push_attribute(("w:sep", s.as_str()));
            }
        }
        #[cfg(feature = "extra-attrs")]
        for (key, value) in &self.extra_attrs {
            start.push_attribute((key.as_str(), value.as_str()));
        }
        start
    }

    fn write_children<W: Write>(&self, writer: &mut Writer<W>) -> Result<(), SerializeError> {
        for item in &self.col {
            item.write_element("w:col", writer)?;
        }
        #[cfg(feature = "extra-children")]
        for child in &self.extra_children {
            child.write_to(writer).map_err(SerializeError::from)?;
        }
        Ok(())
    }

    fn is_empty_element(&self) -> bool {
        if !self.col.is_empty() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for CTVerticalJc {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        {
            let val = &self.value;
            {
                let s = val.to_string();
                start.push_attribute(("w:val", s.as_str()));
            }
        }
        #[cfg(feature = "extra-attrs")]
        for (key, value) in &self.extra_attrs {
            start.push_attribute((key.as_str(), value.as_str()));
        }
        start
    }

    fn is_empty_element(&self) -> bool {
        true
    }
}

impl ToXml for DocumentGrid {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        if let Some(ref val) = self.r#type {
            {
                let s = val.to_string();
                start.push_attribute(("w:type", s.as_str()));
            }
        }
        if let Some(ref val) = self.line_pitch {
            {
                let s = val.to_string();
                start.push_attribute(("w:linePitch", s.as_str()));
            }
        }
        if let Some(ref val) = self.char_space {
            {
                let s = val.to_string();
                start.push_attribute(("w:charSpace", s.as_str()));
            }
        }
        #[cfg(feature = "extra-attrs")]
        for (key, value) in &self.extra_attrs {
            start.push_attribute((key.as_str(), value.as_str()));
        }
        start
    }

    fn is_empty_element(&self) -> bool {
        true
    }
}

impl ToXml for CTHdrFtrRef {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        {
            let val = &self.r#type;
            {
                let s = val.to_string();
                start.push_attribute(("w:type", s.as_str()));
            }
        }
        #[cfg(feature = "extra-attrs")]
        for (key, value) in &self.extra_attrs {
            start.push_attribute((key.as_str(), value.as_str()));
        }
        start
    }

    fn is_empty_element(&self) -> bool {
        true
    }
}

impl ToXml for EGHdrFtrReferences {
    fn write_element<W: Write>(
        &self,
        _tag: &str,
        writer: &mut Writer<W>,
    ) -> Result<(), SerializeError> {
        match self {
            Self::HeaderReference(inner) => inner.write_element("w:headerReference", writer)?,
            Self::FooterReference(inner) => inner.write_element("w:footerReference", writer)?,
        }
        Ok(())
    }
}

impl ToXml for CTHdrFtr {
    fn write_children<W: Write>(&self, writer: &mut Writer<W>) -> Result<(), SerializeError> {
        for item in &self.block_level_elts {
            item.write_element("", writer)?;
        }
        #[cfg(feature = "extra-children")]
        for child in &self.extra_children {
            child.write_to(writer).map_err(SerializeError::from)?;
        }
        Ok(())
    }

    fn is_empty_element(&self) -> bool {
        if !self.block_level_elts.is_empty() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for EGSectPrContents {
    fn write_children<W: Write>(&self, writer: &mut Writer<W>) -> Result<(), SerializeError> {
        if let Some(ref val) = self.footnote_pr {
            val.write_element("w:footnotePr", writer)?;
        }
        if let Some(ref val) = self.endnote_pr {
            val.write_element("w:endnotePr", writer)?;
        }
        if let Some(ref val) = self.r#type {
            val.write_element("w:type", writer)?;
        }
        if let Some(ref val) = self.pg_sz {
            val.write_element("w:pgSz", writer)?;
        }
        if let Some(ref val) = self.pg_mar {
            val.write_element("w:pgMar", writer)?;
        }
        if let Some(ref val) = self.paper_src {
            val.write_element("w:paperSrc", writer)?;
        }
        if let Some(ref val) = self.pg_borders {
            val.write_element("w:pgBorders", writer)?;
        }
        if let Some(ref val) = self.ln_num_type {
            val.write_element("w:lnNumType", writer)?;
        }
        if let Some(ref val) = self.pg_num_type {
            val.write_element("w:pgNumType", writer)?;
        }
        if let Some(ref val) = self.cols {
            val.write_element("w:cols", writer)?;
        }
        if let Some(ref val) = self.form_prot {
            val.write_element("w:formProt", writer)?;
        }
        if let Some(ref val) = self.v_align {
            val.write_element("w:vAlign", writer)?;
        }
        if let Some(ref val) = self.no_endnote {
            val.write_element("w:noEndnote", writer)?;
        }
        if let Some(ref val) = self.title_pg {
            val.write_element("w:titlePg", writer)?;
        }
        if let Some(ref val) = self.text_direction {
            val.write_element("w:textDirection", writer)?;
        }
        if let Some(ref val) = self.bidi {
            val.write_element("w:bidi", writer)?;
        }
        if let Some(ref val) = self.rtl_gutter {
            val.write_element("w:rtlGutter", writer)?;
        }
        if let Some(ref val) = self.doc_grid {
            val.write_element("w:docGrid", writer)?;
        }
        if let Some(ref val) = self.printer_settings {
            val.write_element("w:printerSettings", writer)?;
        }
        #[cfg(feature = "extra-children")]
        for child in &self.extra_children {
            child.write_to(writer).map_err(SerializeError::from)?;
        }
        Ok(())
    }

    fn is_empty_element(&self) -> bool {
        if self.footnote_pr.is_some() {
            return false;
        }
        if self.endnote_pr.is_some() {
            return false;
        }
        if self.r#type.is_some() {
            return false;
        }
        if self.pg_sz.is_some() {
            return false;
        }
        if self.pg_mar.is_some() {
            return false;
        }
        if self.paper_src.is_some() {
            return false;
        }
        if self.pg_borders.is_some() {
            return false;
        }
        if self.ln_num_type.is_some() {
            return false;
        }
        if self.pg_num_type.is_some() {
            return false;
        }
        if self.cols.is_some() {
            return false;
        }
        if self.form_prot.is_some() {
            return false;
        }
        if self.v_align.is_some() {
            return false;
        }
        if self.no_endnote.is_some() {
            return false;
        }
        if self.title_pg.is_some() {
            return false;
        }
        if self.text_direction.is_some() {
            return false;
        }
        if self.bidi.is_some() {
            return false;
        }
        if self.rtl_gutter.is_some() {
            return false;
        }
        if self.doc_grid.is_some() {
            return false;
        }
        if self.printer_settings.is_some() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for WAGSectPrAttributes {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        if let Some(ref val) = self.rsid_r_pr {
            {
                let hex = encode_hex(val);
                start.push_attribute(("w:rsidRPr", hex.as_str()));
            }
        }
        if let Some(ref val) = self.rsid_del {
            {
                let hex = encode_hex(val);
                start.push_attribute(("w:rsidDel", hex.as_str()));
            }
        }
        if let Some(ref val) = self.rsid_r {
            {
                let hex = encode_hex(val);
                start.push_attribute(("w:rsidR", hex.as_str()));
            }
        }
        if let Some(ref val) = self.rsid_sect {
            {
                let hex = encode_hex(val);
                start.push_attribute(("w:rsidSect", hex.as_str()));
            }
        }
        #[cfg(feature = "extra-attrs")]
        for (key, value) in &self.extra_attrs {
            start.push_attribute((key.as_str(), value.as_str()));
        }
        start
    }

    fn is_empty_element(&self) -> bool {
        true
    }
}

impl ToXml for CTSectPrBase {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        if let Some(ref val) = self.rsid_r_pr {
            {
                let hex = encode_hex(val);
                start.push_attribute(("w:rsidRPr", hex.as_str()));
            }
        }
        if let Some(ref val) = self.rsid_del {
            {
                let hex = encode_hex(val);
                start.push_attribute(("w:rsidDel", hex.as_str()));
            }
        }
        if let Some(ref val) = self.rsid_r {
            {
                let hex = encode_hex(val);
                start.push_attribute(("w:rsidR", hex.as_str()));
            }
        }
        if let Some(ref val) = self.rsid_sect {
            {
                let hex = encode_hex(val);
                start.push_attribute(("w:rsidSect", hex.as_str()));
            }
        }
        #[cfg(feature = "extra-attrs")]
        for (key, value) in &self.extra_attrs {
            start.push_attribute((key.as_str(), value.as_str()));
        }
        start
    }

    fn write_children<W: Write>(&self, writer: &mut Writer<W>) -> Result<(), SerializeError> {
        if let Some(ref val) = self.footnote_pr {
            val.write_element("w:footnotePr", writer)?;
        }
        if let Some(ref val) = self.endnote_pr {
            val.write_element("w:endnotePr", writer)?;
        }
        if let Some(ref val) = self.r#type {
            val.write_element("w:type", writer)?;
        }
        if let Some(ref val) = self.pg_sz {
            val.write_element("w:pgSz", writer)?;
        }
        if let Some(ref val) = self.pg_mar {
            val.write_element("w:pgMar", writer)?;
        }
        if let Some(ref val) = self.paper_src {
            val.write_element("w:paperSrc", writer)?;
        }
        if let Some(ref val) = self.pg_borders {
            val.write_element("w:pgBorders", writer)?;
        }
        if let Some(ref val) = self.ln_num_type {
            val.write_element("w:lnNumType", writer)?;
        }
        if let Some(ref val) = self.pg_num_type {
            val.write_element("w:pgNumType", writer)?;
        }
        if let Some(ref val) = self.cols {
            val.write_element("w:cols", writer)?;
        }
        if let Some(ref val) = self.form_prot {
            val.write_element("w:formProt", writer)?;
        }
        if let Some(ref val) = self.v_align {
            val.write_element("w:vAlign", writer)?;
        }
        if let Some(ref val) = self.no_endnote {
            val.write_element("w:noEndnote", writer)?;
        }
        if let Some(ref val) = self.title_pg {
            val.write_element("w:titlePg", writer)?;
        }
        if let Some(ref val) = self.text_direction {
            val.write_element("w:textDirection", writer)?;
        }
        if let Some(ref val) = self.bidi {
            val.write_element("w:bidi", writer)?;
        }
        if let Some(ref val) = self.rtl_gutter {
            val.write_element("w:rtlGutter", writer)?;
        }
        if let Some(ref val) = self.doc_grid {
            val.write_element("w:docGrid", writer)?;
        }
        if let Some(ref val) = self.printer_settings {
            val.write_element("w:printerSettings", writer)?;
        }
        #[cfg(feature = "extra-children")]
        for child in &self.extra_children {
            child.write_to(writer).map_err(SerializeError::from)?;
        }
        Ok(())
    }

    fn is_empty_element(&self) -> bool {
        if self.footnote_pr.is_some() {
            return false;
        }
        if self.endnote_pr.is_some() {
            return false;
        }
        if self.r#type.is_some() {
            return false;
        }
        if self.pg_sz.is_some() {
            return false;
        }
        if self.pg_mar.is_some() {
            return false;
        }
        if self.paper_src.is_some() {
            return false;
        }
        if self.pg_borders.is_some() {
            return false;
        }
        if self.ln_num_type.is_some() {
            return false;
        }
        if self.pg_num_type.is_some() {
            return false;
        }
        if self.cols.is_some() {
            return false;
        }
        if self.form_prot.is_some() {
            return false;
        }
        if self.v_align.is_some() {
            return false;
        }
        if self.no_endnote.is_some() {
            return false;
        }
        if self.title_pg.is_some() {
            return false;
        }
        if self.text_direction.is_some() {
            return false;
        }
        if self.bidi.is_some() {
            return false;
        }
        if self.rtl_gutter.is_some() {
            return false;
        }
        if self.doc_grid.is_some() {
            return false;
        }
        if self.printer_settings.is_some() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for SectionProperties {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        if let Some(ref val) = self.rsid_r_pr {
            {
                let hex = encode_hex(val);
                start.push_attribute(("w:rsidRPr", hex.as_str()));
            }
        }
        if let Some(ref val) = self.rsid_del {
            {
                let hex = encode_hex(val);
                start.push_attribute(("w:rsidDel", hex.as_str()));
            }
        }
        if let Some(ref val) = self.rsid_r {
            {
                let hex = encode_hex(val);
                start.push_attribute(("w:rsidR", hex.as_str()));
            }
        }
        if let Some(ref val) = self.rsid_sect {
            {
                let hex = encode_hex(val);
                start.push_attribute(("w:rsidSect", hex.as_str()));
            }
        }
        #[cfg(feature = "extra-attrs")]
        for (key, value) in &self.extra_attrs {
            start.push_attribute((key.as_str(), value.as_str()));
        }
        start
    }

    fn write_children<W: Write>(&self, writer: &mut Writer<W>) -> Result<(), SerializeError> {
        for item in &self.hdr_ftr_references {
            item.write_element("", writer)?;
        }
        if let Some(ref val) = self.footnote_pr {
            val.write_element("w:footnotePr", writer)?;
        }
        if let Some(ref val) = self.endnote_pr {
            val.write_element("w:endnotePr", writer)?;
        }
        if let Some(ref val) = self.r#type {
            val.write_element("w:type", writer)?;
        }
        if let Some(ref val) = self.pg_sz {
            val.write_element("w:pgSz", writer)?;
        }
        if let Some(ref val) = self.pg_mar {
            val.write_element("w:pgMar", writer)?;
        }
        if let Some(ref val) = self.paper_src {
            val.write_element("w:paperSrc", writer)?;
        }
        if let Some(ref val) = self.pg_borders {
            val.write_element("w:pgBorders", writer)?;
        }
        if let Some(ref val) = self.ln_num_type {
            val.write_element("w:lnNumType", writer)?;
        }
        if let Some(ref val) = self.pg_num_type {
            val.write_element("w:pgNumType", writer)?;
        }
        if let Some(ref val) = self.cols {
            val.write_element("w:cols", writer)?;
        }
        if let Some(ref val) = self.form_prot {
            val.write_element("w:formProt", writer)?;
        }
        if let Some(ref val) = self.v_align {
            val.write_element("w:vAlign", writer)?;
        }
        if let Some(ref val) = self.no_endnote {
            val.write_element("w:noEndnote", writer)?;
        }
        if let Some(ref val) = self.title_pg {
            val.write_element("w:titlePg", writer)?;
        }
        if let Some(ref val) = self.text_direction {
            val.write_element("w:textDirection", writer)?;
        }
        if let Some(ref val) = self.bidi {
            val.write_element("w:bidi", writer)?;
        }
        if let Some(ref val) = self.rtl_gutter {
            val.write_element("w:rtlGutter", writer)?;
        }
        if let Some(ref val) = self.doc_grid {
            val.write_element("w:docGrid", writer)?;
        }
        if let Some(ref val) = self.printer_settings {
            val.write_element("w:printerSettings", writer)?;
        }
        if let Some(ref val) = self.sect_pr_change {
            val.write_element("w:sectPrChange", writer)?;
        }
        #[cfg(feature = "extra-children")]
        for child in &self.extra_children {
            child.write_to(writer).map_err(SerializeError::from)?;
        }
        Ok(())
    }

    fn is_empty_element(&self) -> bool {
        if !self.hdr_ftr_references.is_empty() {
            return false;
        }
        if self.footnote_pr.is_some() {
            return false;
        }
        if self.endnote_pr.is_some() {
            return false;
        }
        if self.r#type.is_some() {
            return false;
        }
        if self.pg_sz.is_some() {
            return false;
        }
        if self.pg_mar.is_some() {
            return false;
        }
        if self.paper_src.is_some() {
            return false;
        }
        if self.pg_borders.is_some() {
            return false;
        }
        if self.ln_num_type.is_some() {
            return false;
        }
        if self.pg_num_type.is_some() {
            return false;
        }
        if self.cols.is_some() {
            return false;
        }
        if self.form_prot.is_some() {
            return false;
        }
        if self.v_align.is_some() {
            return false;
        }
        if self.no_endnote.is_some() {
            return false;
        }
        if self.title_pg.is_some() {
            return false;
        }
        if self.text_direction.is_some() {
            return false;
        }
        if self.bidi.is_some() {
            return false;
        }
        if self.rtl_gutter.is_some() {
            return false;
        }
        if self.doc_grid.is_some() {
            return false;
        }
        if self.printer_settings.is_some() {
            return false;
        }
        if self.sect_pr_change.is_some() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for CTBr {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        if let Some(ref val) = self.r#type {
            {
                let s = val.to_string();
                start.push_attribute(("w:type", s.as_str()));
            }
        }
        if let Some(ref val) = self.clear {
            {
                let s = val.to_string();
                start.push_attribute(("w:clear", s.as_str()));
            }
        }
        #[cfg(feature = "extra-attrs")]
        for (key, value) in &self.extra_attrs {
            start.push_attribute((key.as_str(), value.as_str()));
        }
        start
    }

    fn is_empty_element(&self) -> bool {
        true
    }
}

impl ToXml for CTPTab {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        {
            let val = &self.alignment;
            {
                let s = val.to_string();
                start.push_attribute(("w:alignment", s.as_str()));
            }
        }
        {
            let val = &self.relative_to;
            {
                let s = val.to_string();
                start.push_attribute(("w:relativeTo", s.as_str()));
            }
        }
        {
            let val = &self.leader;
            {
                let s = val.to_string();
                start.push_attribute(("w:leader", s.as_str()));
            }
        }
        #[cfg(feature = "extra-attrs")]
        for (key, value) in &self.extra_attrs {
            start.push_attribute((key.as_str(), value.as_str()));
        }
        start
    }

    fn is_empty_element(&self) -> bool {
        true
    }
}

impl ToXml for CTSym {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        if let Some(ref val) = self.font {
            start.push_attribute(("w:font", val.as_str()));
        }
        if let Some(ref val) = self.char {
            {
                let hex = encode_hex(val);
                start.push_attribute(("w:char", hex.as_str()));
            }
        }
        #[cfg(feature = "extra-attrs")]
        for (key, value) in &self.extra_attrs {
            start.push_attribute((key.as_str(), value.as_str()));
        }
        start
    }

    fn is_empty_element(&self) -> bool {
        true
    }
}

impl ToXml for CTProofErr {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        {
            let val = &self.r#type;
            {
                let s = val.to_string();
                start.push_attribute(("w:type", s.as_str()));
            }
        }
        #[cfg(feature = "extra-attrs")]
        for (key, value) in &self.extra_attrs {
            start.push_attribute((key.as_str(), value.as_str()));
        }
        start
    }

    fn is_empty_element(&self) -> bool {
        true
    }
}

impl ToXml for CTPerm {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        {
            let val = &self.id;
            start.push_attribute(("w:id", val.as_str()));
        }
        if let Some(ref val) = self.displaced_by_custom_xml {
            {
                let s = val.to_string();
                start.push_attribute(("w:displacedByCustomXml", s.as_str()));
            }
        }
        #[cfg(feature = "extra-attrs")]
        for (key, value) in &self.extra_attrs {
            start.push_attribute((key.as_str(), value.as_str()));
        }
        start
    }

    fn is_empty_element(&self) -> bool {
        true
    }
}

impl ToXml for CTPermStart {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        {
            let val = &self.id;
            start.push_attribute(("w:id", val.as_str()));
        }
        if let Some(ref val) = self.displaced_by_custom_xml {
            {
                let s = val.to_string();
                start.push_attribute(("w:displacedByCustomXml", s.as_str()));
            }
        }
        if let Some(ref val) = self.ed_grp {
            {
                let s = val.to_string();
                start.push_attribute(("w:edGrp", s.as_str()));
            }
        }
        if let Some(ref val) = self.ed {
            start.push_attribute(("w:ed", val.as_str()));
        }
        if let Some(ref val) = self.col_first {
            {
                let s = val.to_string();
                start.push_attribute(("w:colFirst", s.as_str()));
            }
        }
        if let Some(ref val) = self.col_last {
            {
                let s = val.to_string();
                start.push_attribute(("w:colLast", s.as_str()));
            }
        }
        #[cfg(feature = "extra-attrs")]
        for (key, value) in &self.extra_attrs {
            start.push_attribute((key.as_str(), value.as_str()));
        }
        start
    }

    fn is_empty_element(&self) -> bool {
        true
    }
}

impl ToXml for Text {
    fn write_children<W: Write>(&self, writer: &mut Writer<W>) -> Result<(), SerializeError> {
        if let Some(ref text) = self.text {
            writer.write_event(Event::Text(BytesText::new(text)))?;
        }
        #[cfg(feature = "extra-children")]
        for child in &self.extra_children {
            child.write_to(writer).map_err(SerializeError::from)?;
        }
        Ok(())
    }

    fn is_empty_element(&self) -> bool {
        if self.text.is_some() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for EGRunInnerContent {
    fn write_element<W: Write>(
        &self,
        _tag: &str,
        writer: &mut Writer<W>,
    ) -> Result<(), SerializeError> {
        match self {
            Self::Br(inner) => inner.write_element("w:br", writer)?,
            Self::T(inner) => inner.write_element("w:t", writer)?,
            Self::ContentPart(inner) => inner.write_element("w:contentPart", writer)?,
            Self::DelText(inner) => inner.write_element("w:delText", writer)?,
            Self::InstrText(inner) => inner.write_element("w:instrText", writer)?,
            Self::DelInstrText(inner) => inner.write_element("w:delInstrText", writer)?,
            Self::NoBreakHyphen(inner) => inner.write_element("w:noBreakHyphen", writer)?,
            Self::SoftHyphen(inner) => inner.write_element("w:softHyphen", writer)?,
            Self::DayShort(inner) => inner.write_element("w:dayShort", writer)?,
            Self::MonthShort(inner) => inner.write_element("w:monthShort", writer)?,
            Self::YearShort(inner) => inner.write_element("w:yearShort", writer)?,
            Self::DayLong(inner) => inner.write_element("w:dayLong", writer)?,
            Self::MonthLong(inner) => inner.write_element("w:monthLong", writer)?,
            Self::YearLong(inner) => inner.write_element("w:yearLong", writer)?,
            Self::AnnotationRef(inner) => inner.write_element("w:annotationRef", writer)?,
            Self::FootnoteRef(inner) => inner.write_element("w:footnoteRef", writer)?,
            Self::EndnoteRef(inner) => inner.write_element("w:endnoteRef", writer)?,
            Self::Separator(inner) => inner.write_element("w:separator", writer)?,
            Self::ContinuationSeparator(inner) => {
                inner.write_element("w:continuationSeparator", writer)?
            }
            Self::Sym(inner) => inner.write_element("w:sym", writer)?,
            Self::PgNum(inner) => inner.write_element("w:pgNum", writer)?,
            Self::Cr(inner) => inner.write_element("w:cr", writer)?,
            Self::Tab(inner) => inner.write_element("w:tab", writer)?,
            Self::Object(inner) => inner.write_element("w:object", writer)?,
            Self::Pict(inner) => inner.write_element("w:pict", writer)?,
            Self::FldChar(inner) => inner.write_element("w:fldChar", writer)?,
            Self::Ruby(inner) => inner.write_element("w:ruby", writer)?,
            Self::FootnoteReference(inner) => inner.write_element("w:footnoteReference", writer)?,
            Self::EndnoteReference(inner) => inner.write_element("w:endnoteReference", writer)?,
            Self::CommentReference(inner) => inner.write_element("w:commentReference", writer)?,
            Self::Drawing(inner) => inner.write_element("w:drawing", writer)?,
            Self::Ptab(inner) => inner.write_element("w:ptab", writer)?,
            Self::LastRenderedPageBreak(inner) => {
                inner.write_element("w:lastRenderedPageBreak", writer)?
            }
        }
        Ok(())
    }
}

impl ToXml for Run {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        if let Some(ref val) = self.rsid_r_pr {
            {
                let hex = encode_hex(val);
                start.push_attribute(("w:rsidRPr", hex.as_str()));
            }
        }
        if let Some(ref val) = self.rsid_del {
            {
                let hex = encode_hex(val);
                start.push_attribute(("w:rsidDel", hex.as_str()));
            }
        }
        if let Some(ref val) = self.rsid_r {
            {
                let hex = encode_hex(val);
                start.push_attribute(("w:rsidR", hex.as_str()));
            }
        }
        #[cfg(feature = "extra-attrs")]
        for (key, value) in &self.extra_attrs {
            start.push_attribute((key.as_str(), value.as_str()));
        }
        start
    }

    fn write_children<W: Write>(&self, writer: &mut Writer<W>) -> Result<(), SerializeError> {
        #[cfg(feature = "wml-styling")]
        if let Some(ref val) = self.r_pr {
            val.write_element("w:rPr", writer)?;
        }
        for item in &self.run_inner_content {
            item.write_element("", writer)?;
        }
        #[cfg(feature = "extra-children")]
        for child in &self.extra_children {
            child.write_to(writer).map_err(SerializeError::from)?;
        }
        Ok(())
    }

    fn is_empty_element(&self) -> bool {
        #[cfg(feature = "wml-styling")]
        if self.r_pr.is_some() {
            return false;
        }
        if !self.run_inner_content.is_empty() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for Fonts {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        if let Some(ref val) = self.hint {
            {
                let s = val.to_string();
                start.push_attribute(("w:hint", s.as_str()));
            }
        }
        if let Some(ref val) = self.ascii {
            start.push_attribute(("w:ascii", val.as_str()));
        }
        if let Some(ref val) = self.h_ansi {
            start.push_attribute(("w:hAnsi", val.as_str()));
        }
        if let Some(ref val) = self.east_asia {
            start.push_attribute(("w:eastAsia", val.as_str()));
        }
        if let Some(ref val) = self.cs {
            start.push_attribute(("w:cs", val.as_str()));
        }
        if let Some(ref val) = self.ascii_theme {
            {
                let s = val.to_string();
                start.push_attribute(("w:asciiTheme", s.as_str()));
            }
        }
        if let Some(ref val) = self.h_ansi_theme {
            {
                let s = val.to_string();
                start.push_attribute(("w:hAnsiTheme", s.as_str()));
            }
        }
        if let Some(ref val) = self.east_asia_theme {
            {
                let s = val.to_string();
                start.push_attribute(("w:eastAsiaTheme", s.as_str()));
            }
        }
        if let Some(ref val) = self.cstheme {
            {
                let s = val.to_string();
                start.push_attribute(("w:cstheme", s.as_str()));
            }
        }
        #[cfg(feature = "extra-attrs")]
        for (key, value) in &self.extra_attrs {
            start.push_attribute((key.as_str(), value.as_str()));
        }
        start
    }

    fn is_empty_element(&self) -> bool {
        true
    }
}

impl ToXml for EGRPrBase {
    fn write_children<W: Write>(&self, writer: &mut Writer<W>) -> Result<(), SerializeError> {
        if let Some(ref val) = self.run_style {
            val.write_element("w:rStyle", writer)?;
        }
        if let Some(ref val) = self.fonts {
            val.write_element("w:rFonts", writer)?;
        }
        if let Some(ref val) = self.bold {
            val.write_element("w:b", writer)?;
        }
        if let Some(ref val) = self.b_cs {
            val.write_element("w:bCs", writer)?;
        }
        if let Some(ref val) = self.italic {
            val.write_element("w:i", writer)?;
        }
        if let Some(ref val) = self.i_cs {
            val.write_element("w:iCs", writer)?;
        }
        if let Some(ref val) = self.caps {
            val.write_element("w:caps", writer)?;
        }
        if let Some(ref val) = self.small_caps {
            val.write_element("w:smallCaps", writer)?;
        }
        if let Some(ref val) = self.strikethrough {
            val.write_element("w:strike", writer)?;
        }
        if let Some(ref val) = self.dstrike {
            val.write_element("w:dstrike", writer)?;
        }
        if let Some(ref val) = self.outline {
            val.write_element("w:outline", writer)?;
        }
        if let Some(ref val) = self.shadow {
            val.write_element("w:shadow", writer)?;
        }
        if let Some(ref val) = self.emboss {
            val.write_element("w:emboss", writer)?;
        }
        if let Some(ref val) = self.imprint {
            val.write_element("w:imprint", writer)?;
        }
        if let Some(ref val) = self.no_proof {
            val.write_element("w:noProof", writer)?;
        }
        if let Some(ref val) = self.snap_to_grid {
            val.write_element("w:snapToGrid", writer)?;
        }
        if let Some(ref val) = self.vanish {
            val.write_element("w:vanish", writer)?;
        }
        if let Some(ref val) = self.web_hidden {
            val.write_element("w:webHidden", writer)?;
        }
        if let Some(ref val) = self.color {
            val.write_element("w:color", writer)?;
        }
        if let Some(ref val) = self.spacing {
            val.write_element("w:spacing", writer)?;
        }
        if let Some(ref val) = self.width {
            val.write_element("w:w", writer)?;
        }
        if let Some(ref val) = self.kern {
            val.write_element("w:kern", writer)?;
        }
        if let Some(ref val) = self.position {
            val.write_element("w:position", writer)?;
        }
        if let Some(ref val) = self.size {
            val.write_element("w:sz", writer)?;
        }
        if let Some(ref val) = self.size_complex_script {
            val.write_element("w:szCs", writer)?;
        }
        if let Some(ref val) = self.highlight {
            val.write_element("w:highlight", writer)?;
        }
        if let Some(ref val) = self.underline {
            val.write_element("w:u", writer)?;
        }
        if let Some(ref val) = self.effect {
            val.write_element("w:effect", writer)?;
        }
        if let Some(ref val) = self.bdr {
            val.write_element("w:bdr", writer)?;
        }
        if let Some(ref val) = self.shading {
            val.write_element("w:shd", writer)?;
        }
        if let Some(ref val) = self.fit_text {
            val.write_element("w:fitText", writer)?;
        }
        if let Some(ref val) = self.vert_align {
            val.write_element("w:vertAlign", writer)?;
        }
        if let Some(ref val) = self.rtl {
            val.write_element("w:rtl", writer)?;
        }
        if let Some(ref val) = self.cs {
            val.write_element("w:cs", writer)?;
        }
        if let Some(ref val) = self.em {
            val.write_element("w:em", writer)?;
        }
        if let Some(ref val) = self.lang {
            val.write_element("w:lang", writer)?;
        }
        if let Some(ref val) = self.east_asian_layout {
            val.write_element("w:eastAsianLayout", writer)?;
        }
        if let Some(ref val) = self.spec_vanish {
            val.write_element("w:specVanish", writer)?;
        }
        if let Some(ref val) = self.o_math {
            val.write_element("w:oMath", writer)?;
        }
        #[cfg(feature = "extra-children")]
        for child in &self.extra_children {
            child.write_to(writer).map_err(SerializeError::from)?;
        }
        Ok(())
    }

    fn is_empty_element(&self) -> bool {
        if self.run_style.is_some() {
            return false;
        }
        if self.fonts.is_some() {
            return false;
        }
        if self.bold.is_some() {
            return false;
        }
        if self.b_cs.is_some() {
            return false;
        }
        if self.italic.is_some() {
            return false;
        }
        if self.i_cs.is_some() {
            return false;
        }
        if self.caps.is_some() {
            return false;
        }
        if self.small_caps.is_some() {
            return false;
        }
        if self.strikethrough.is_some() {
            return false;
        }
        if self.dstrike.is_some() {
            return false;
        }
        if self.outline.is_some() {
            return false;
        }
        if self.shadow.is_some() {
            return false;
        }
        if self.emboss.is_some() {
            return false;
        }
        if self.imprint.is_some() {
            return false;
        }
        if self.no_proof.is_some() {
            return false;
        }
        if self.snap_to_grid.is_some() {
            return false;
        }
        if self.vanish.is_some() {
            return false;
        }
        if self.web_hidden.is_some() {
            return false;
        }
        if self.color.is_some() {
            return false;
        }
        if self.spacing.is_some() {
            return false;
        }
        if self.width.is_some() {
            return false;
        }
        if self.kern.is_some() {
            return false;
        }
        if self.position.is_some() {
            return false;
        }
        if self.size.is_some() {
            return false;
        }
        if self.size_complex_script.is_some() {
            return false;
        }
        if self.highlight.is_some() {
            return false;
        }
        if self.underline.is_some() {
            return false;
        }
        if self.effect.is_some() {
            return false;
        }
        if self.bdr.is_some() {
            return false;
        }
        if self.shading.is_some() {
            return false;
        }
        if self.fit_text.is_some() {
            return false;
        }
        if self.vert_align.is_some() {
            return false;
        }
        if self.rtl.is_some() {
            return false;
        }
        if self.cs.is_some() {
            return false;
        }
        if self.em.is_some() {
            return false;
        }
        if self.lang.is_some() {
            return false;
        }
        if self.east_asian_layout.is_some() {
            return false;
        }
        if self.spec_vanish.is_some() {
            return false;
        }
        if self.o_math.is_some() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for EGRPrContent {
    fn write_children<W: Write>(&self, writer: &mut Writer<W>) -> Result<(), SerializeError> {
        if let Some(ref val) = self.run_style {
            val.write_element("w:rStyle", writer)?;
        }
        if let Some(ref val) = self.fonts {
            val.write_element("w:rFonts", writer)?;
        }
        if let Some(ref val) = self.bold {
            val.write_element("w:b", writer)?;
        }
        if let Some(ref val) = self.b_cs {
            val.write_element("w:bCs", writer)?;
        }
        if let Some(ref val) = self.italic {
            val.write_element("w:i", writer)?;
        }
        if let Some(ref val) = self.i_cs {
            val.write_element("w:iCs", writer)?;
        }
        if let Some(ref val) = self.caps {
            val.write_element("w:caps", writer)?;
        }
        if let Some(ref val) = self.small_caps {
            val.write_element("w:smallCaps", writer)?;
        }
        if let Some(ref val) = self.strikethrough {
            val.write_element("w:strike", writer)?;
        }
        if let Some(ref val) = self.dstrike {
            val.write_element("w:dstrike", writer)?;
        }
        if let Some(ref val) = self.outline {
            val.write_element("w:outline", writer)?;
        }
        if let Some(ref val) = self.shadow {
            val.write_element("w:shadow", writer)?;
        }
        if let Some(ref val) = self.emboss {
            val.write_element("w:emboss", writer)?;
        }
        if let Some(ref val) = self.imprint {
            val.write_element("w:imprint", writer)?;
        }
        if let Some(ref val) = self.no_proof {
            val.write_element("w:noProof", writer)?;
        }
        if let Some(ref val) = self.snap_to_grid {
            val.write_element("w:snapToGrid", writer)?;
        }
        if let Some(ref val) = self.vanish {
            val.write_element("w:vanish", writer)?;
        }
        if let Some(ref val) = self.web_hidden {
            val.write_element("w:webHidden", writer)?;
        }
        if let Some(ref val) = self.color {
            val.write_element("w:color", writer)?;
        }
        if let Some(ref val) = self.spacing {
            val.write_element("w:spacing", writer)?;
        }
        if let Some(ref val) = self.width {
            val.write_element("w:w", writer)?;
        }
        if let Some(ref val) = self.kern {
            val.write_element("w:kern", writer)?;
        }
        if let Some(ref val) = self.position {
            val.write_element("w:position", writer)?;
        }
        if let Some(ref val) = self.size {
            val.write_element("w:sz", writer)?;
        }
        if let Some(ref val) = self.size_complex_script {
            val.write_element("w:szCs", writer)?;
        }
        if let Some(ref val) = self.highlight {
            val.write_element("w:highlight", writer)?;
        }
        if let Some(ref val) = self.underline {
            val.write_element("w:u", writer)?;
        }
        if let Some(ref val) = self.effect {
            val.write_element("w:effect", writer)?;
        }
        if let Some(ref val) = self.bdr {
            val.write_element("w:bdr", writer)?;
        }
        if let Some(ref val) = self.shading {
            val.write_element("w:shd", writer)?;
        }
        if let Some(ref val) = self.fit_text {
            val.write_element("w:fitText", writer)?;
        }
        if let Some(ref val) = self.vert_align {
            val.write_element("w:vertAlign", writer)?;
        }
        if let Some(ref val) = self.rtl {
            val.write_element("w:rtl", writer)?;
        }
        if let Some(ref val) = self.cs {
            val.write_element("w:cs", writer)?;
        }
        if let Some(ref val) = self.em {
            val.write_element("w:em", writer)?;
        }
        if let Some(ref val) = self.lang {
            val.write_element("w:lang", writer)?;
        }
        if let Some(ref val) = self.east_asian_layout {
            val.write_element("w:eastAsianLayout", writer)?;
        }
        if let Some(ref val) = self.spec_vanish {
            val.write_element("w:specVanish", writer)?;
        }
        if let Some(ref val) = self.o_math {
            val.write_element("w:oMath", writer)?;
        }
        if let Some(ref val) = self.r_pr_change {
            val.write_element("w:rPrChange", writer)?;
        }
        #[cfg(feature = "extra-children")]
        for child in &self.extra_children {
            child.write_to(writer).map_err(SerializeError::from)?;
        }
        Ok(())
    }

    fn is_empty_element(&self) -> bool {
        if self.run_style.is_some() {
            return false;
        }
        if self.fonts.is_some() {
            return false;
        }
        if self.bold.is_some() {
            return false;
        }
        if self.b_cs.is_some() {
            return false;
        }
        if self.italic.is_some() {
            return false;
        }
        if self.i_cs.is_some() {
            return false;
        }
        if self.caps.is_some() {
            return false;
        }
        if self.small_caps.is_some() {
            return false;
        }
        if self.strikethrough.is_some() {
            return false;
        }
        if self.dstrike.is_some() {
            return false;
        }
        if self.outline.is_some() {
            return false;
        }
        if self.shadow.is_some() {
            return false;
        }
        if self.emboss.is_some() {
            return false;
        }
        if self.imprint.is_some() {
            return false;
        }
        if self.no_proof.is_some() {
            return false;
        }
        if self.snap_to_grid.is_some() {
            return false;
        }
        if self.vanish.is_some() {
            return false;
        }
        if self.web_hidden.is_some() {
            return false;
        }
        if self.color.is_some() {
            return false;
        }
        if self.spacing.is_some() {
            return false;
        }
        if self.width.is_some() {
            return false;
        }
        if self.kern.is_some() {
            return false;
        }
        if self.position.is_some() {
            return false;
        }
        if self.size.is_some() {
            return false;
        }
        if self.size_complex_script.is_some() {
            return false;
        }
        if self.highlight.is_some() {
            return false;
        }
        if self.underline.is_some() {
            return false;
        }
        if self.effect.is_some() {
            return false;
        }
        if self.bdr.is_some() {
            return false;
        }
        if self.shading.is_some() {
            return false;
        }
        if self.fit_text.is_some() {
            return false;
        }
        if self.vert_align.is_some() {
            return false;
        }
        if self.rtl.is_some() {
            return false;
        }
        if self.cs.is_some() {
            return false;
        }
        if self.em.is_some() {
            return false;
        }
        if self.lang.is_some() {
            return false;
        }
        if self.east_asian_layout.is_some() {
            return false;
        }
        if self.spec_vanish.is_some() {
            return false;
        }
        if self.o_math.is_some() {
            return false;
        }
        if self.r_pr_change.is_some() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for RunProperties {
    fn write_children<W: Write>(&self, writer: &mut Writer<W>) -> Result<(), SerializeError> {
        if let Some(ref val) = self.run_style {
            val.write_element("w:rStyle", writer)?;
        }
        if let Some(ref val) = self.fonts {
            val.write_element("w:rFonts", writer)?;
        }
        if let Some(ref val) = self.bold {
            val.write_element("w:b", writer)?;
        }
        if let Some(ref val) = self.b_cs {
            val.write_element("w:bCs", writer)?;
        }
        if let Some(ref val) = self.italic {
            val.write_element("w:i", writer)?;
        }
        if let Some(ref val) = self.i_cs {
            val.write_element("w:iCs", writer)?;
        }
        if let Some(ref val) = self.caps {
            val.write_element("w:caps", writer)?;
        }
        if let Some(ref val) = self.small_caps {
            val.write_element("w:smallCaps", writer)?;
        }
        if let Some(ref val) = self.strikethrough {
            val.write_element("w:strike", writer)?;
        }
        if let Some(ref val) = self.dstrike {
            val.write_element("w:dstrike", writer)?;
        }
        if let Some(ref val) = self.outline {
            val.write_element("w:outline", writer)?;
        }
        if let Some(ref val) = self.shadow {
            val.write_element("w:shadow", writer)?;
        }
        if let Some(ref val) = self.emboss {
            val.write_element("w:emboss", writer)?;
        }
        if let Some(ref val) = self.imprint {
            val.write_element("w:imprint", writer)?;
        }
        if let Some(ref val) = self.no_proof {
            val.write_element("w:noProof", writer)?;
        }
        if let Some(ref val) = self.snap_to_grid {
            val.write_element("w:snapToGrid", writer)?;
        }
        if let Some(ref val) = self.vanish {
            val.write_element("w:vanish", writer)?;
        }
        if let Some(ref val) = self.web_hidden {
            val.write_element("w:webHidden", writer)?;
        }
        if let Some(ref val) = self.color {
            val.write_element("w:color", writer)?;
        }
        if let Some(ref val) = self.spacing {
            val.write_element("w:spacing", writer)?;
        }
        if let Some(ref val) = self.width {
            val.write_element("w:w", writer)?;
        }
        if let Some(ref val) = self.kern {
            val.write_element("w:kern", writer)?;
        }
        if let Some(ref val) = self.position {
            val.write_element("w:position", writer)?;
        }
        if let Some(ref val) = self.size {
            val.write_element("w:sz", writer)?;
        }
        if let Some(ref val) = self.size_complex_script {
            val.write_element("w:szCs", writer)?;
        }
        if let Some(ref val) = self.highlight {
            val.write_element("w:highlight", writer)?;
        }
        if let Some(ref val) = self.underline {
            val.write_element("w:u", writer)?;
        }
        if let Some(ref val) = self.effect {
            val.write_element("w:effect", writer)?;
        }
        if let Some(ref val) = self.bdr {
            val.write_element("w:bdr", writer)?;
        }
        if let Some(ref val) = self.shading {
            val.write_element("w:shd", writer)?;
        }
        if let Some(ref val) = self.fit_text {
            val.write_element("w:fitText", writer)?;
        }
        if let Some(ref val) = self.vert_align {
            val.write_element("w:vertAlign", writer)?;
        }
        if let Some(ref val) = self.rtl {
            val.write_element("w:rtl", writer)?;
        }
        if let Some(ref val) = self.cs {
            val.write_element("w:cs", writer)?;
        }
        if let Some(ref val) = self.em {
            val.write_element("w:em", writer)?;
        }
        if let Some(ref val) = self.lang {
            val.write_element("w:lang", writer)?;
        }
        if let Some(ref val) = self.east_asian_layout {
            val.write_element("w:eastAsianLayout", writer)?;
        }
        if let Some(ref val) = self.spec_vanish {
            val.write_element("w:specVanish", writer)?;
        }
        if let Some(ref val) = self.o_math {
            val.write_element("w:oMath", writer)?;
        }
        if let Some(ref val) = self.r_pr_change {
            val.write_element("w:rPrChange", writer)?;
        }
        #[cfg(feature = "extra-children")]
        for child in &self.extra_children {
            child.write_to(writer).map_err(SerializeError::from)?;
        }
        Ok(())
    }

    fn is_empty_element(&self) -> bool {
        if self.run_style.is_some() {
            return false;
        }
        if self.fonts.is_some() {
            return false;
        }
        if self.bold.is_some() {
            return false;
        }
        if self.b_cs.is_some() {
            return false;
        }
        if self.italic.is_some() {
            return false;
        }
        if self.i_cs.is_some() {
            return false;
        }
        if self.caps.is_some() {
            return false;
        }
        if self.small_caps.is_some() {
            return false;
        }
        if self.strikethrough.is_some() {
            return false;
        }
        if self.dstrike.is_some() {
            return false;
        }
        if self.outline.is_some() {
            return false;
        }
        if self.shadow.is_some() {
            return false;
        }
        if self.emboss.is_some() {
            return false;
        }
        if self.imprint.is_some() {
            return false;
        }
        if self.no_proof.is_some() {
            return false;
        }
        if self.snap_to_grid.is_some() {
            return false;
        }
        if self.vanish.is_some() {
            return false;
        }
        if self.web_hidden.is_some() {
            return false;
        }
        if self.color.is_some() {
            return false;
        }
        if self.spacing.is_some() {
            return false;
        }
        if self.width.is_some() {
            return false;
        }
        if self.kern.is_some() {
            return false;
        }
        if self.position.is_some() {
            return false;
        }
        if self.size.is_some() {
            return false;
        }
        if self.size_complex_script.is_some() {
            return false;
        }
        if self.highlight.is_some() {
            return false;
        }
        if self.underline.is_some() {
            return false;
        }
        if self.effect.is_some() {
            return false;
        }
        if self.bdr.is_some() {
            return false;
        }
        if self.shading.is_some() {
            return false;
        }
        if self.fit_text.is_some() {
            return false;
        }
        if self.vert_align.is_some() {
            return false;
        }
        if self.rtl.is_some() {
            return false;
        }
        if self.cs.is_some() {
            return false;
        }
        if self.em.is_some() {
            return false;
        }
        if self.lang.is_some() {
            return false;
        }
        if self.east_asian_layout.is_some() {
            return false;
        }
        if self.spec_vanish.is_some() {
            return false;
        }
        if self.o_math.is_some() {
            return false;
        }
        if self.r_pr_change.is_some() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for EGRPr {
    fn write_children<W: Write>(&self, writer: &mut Writer<W>) -> Result<(), SerializeError> {
        if let Some(ref val) = self.r_pr {
            val.write_element("w:rPr", writer)?;
        }
        #[cfg(feature = "extra-children")]
        for child in &self.extra_children {
            child.write_to(writer).map_err(SerializeError::from)?;
        }
        Ok(())
    }

    fn is_empty_element(&self) -> bool {
        if self.r_pr.is_some() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for EGRPrMath {
    fn write_element<W: Write>(
        &self,
        _tag: &str,
        writer: &mut Writer<W>,
    ) -> Result<(), SerializeError> {
        match self {
            Self::RPr(inner) => inner.write_element("w:rPr", writer)?,
            Self::Ins(inner) => inner.write_element("w:ins", writer)?,
            Self::Del(inner) => inner.write_element("w:del", writer)?,
        }
        Ok(())
    }
}

impl ToXml for CTMathCtrlIns {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        {
            let val = &self.id;
            {
                let s = val.to_string();
                start.push_attribute(("w:id", s.as_str()));
            }
        }
        {
            let val = &self.author;
            start.push_attribute(("w:author", val.as_str()));
        }
        if let Some(ref val) = self.date {
            start.push_attribute(("w:date", val.as_str()));
        }
        #[cfg(feature = "extra-attrs")]
        for (key, value) in &self.extra_attrs {
            start.push_attribute((key.as_str(), value.as_str()));
        }
        start
    }

    fn is_empty_element(&self) -> bool {
        true
    }
}

impl ToXml for CTMathCtrlDel {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        {
            let val = &self.id;
            {
                let s = val.to_string();
                start.push_attribute(("w:id", s.as_str()));
            }
        }
        {
            let val = &self.author;
            start.push_attribute(("w:author", val.as_str()));
        }
        if let Some(ref val) = self.date {
            start.push_attribute(("w:date", val.as_str()));
        }
        #[cfg(feature = "extra-attrs")]
        for (key, value) in &self.extra_attrs {
            start.push_attribute((key.as_str(), value.as_str()));
        }
        start
    }

    fn write_children<W: Write>(&self, writer: &mut Writer<W>) -> Result<(), SerializeError> {
        if let Some(ref val) = self.r_pr {
            val.write_element("w:rPr", writer)?;
        }
        #[cfg(feature = "extra-children")]
        for child in &self.extra_children {
            child.write_to(writer).map_err(SerializeError::from)?;
        }
        Ok(())
    }

    fn is_empty_element(&self) -> bool {
        if self.r_pr.is_some() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for CTRPrOriginal {
    fn write_children<W: Write>(&self, writer: &mut Writer<W>) -> Result<(), SerializeError> {
        if let Some(ref val) = self.run_style {
            val.write_element("w:rStyle", writer)?;
        }
        if let Some(ref val) = self.fonts {
            val.write_element("w:rFonts", writer)?;
        }
        if let Some(ref val) = self.bold {
            val.write_element("w:b", writer)?;
        }
        if let Some(ref val) = self.b_cs {
            val.write_element("w:bCs", writer)?;
        }
        if let Some(ref val) = self.italic {
            val.write_element("w:i", writer)?;
        }
        if let Some(ref val) = self.i_cs {
            val.write_element("w:iCs", writer)?;
        }
        if let Some(ref val) = self.caps {
            val.write_element("w:caps", writer)?;
        }
        if let Some(ref val) = self.small_caps {
            val.write_element("w:smallCaps", writer)?;
        }
        if let Some(ref val) = self.strikethrough {
            val.write_element("w:strike", writer)?;
        }
        if let Some(ref val) = self.dstrike {
            val.write_element("w:dstrike", writer)?;
        }
        if let Some(ref val) = self.outline {
            val.write_element("w:outline", writer)?;
        }
        if let Some(ref val) = self.shadow {
            val.write_element("w:shadow", writer)?;
        }
        if let Some(ref val) = self.emboss {
            val.write_element("w:emboss", writer)?;
        }
        if let Some(ref val) = self.imprint {
            val.write_element("w:imprint", writer)?;
        }
        if let Some(ref val) = self.no_proof {
            val.write_element("w:noProof", writer)?;
        }
        if let Some(ref val) = self.snap_to_grid {
            val.write_element("w:snapToGrid", writer)?;
        }
        if let Some(ref val) = self.vanish {
            val.write_element("w:vanish", writer)?;
        }
        if let Some(ref val) = self.web_hidden {
            val.write_element("w:webHidden", writer)?;
        }
        if let Some(ref val) = self.color {
            val.write_element("w:color", writer)?;
        }
        if let Some(ref val) = self.spacing {
            val.write_element("w:spacing", writer)?;
        }
        if let Some(ref val) = self.width {
            val.write_element("w:w", writer)?;
        }
        if let Some(ref val) = self.kern {
            val.write_element("w:kern", writer)?;
        }
        if let Some(ref val) = self.position {
            val.write_element("w:position", writer)?;
        }
        if let Some(ref val) = self.size {
            val.write_element("w:sz", writer)?;
        }
        if let Some(ref val) = self.size_complex_script {
            val.write_element("w:szCs", writer)?;
        }
        if let Some(ref val) = self.highlight {
            val.write_element("w:highlight", writer)?;
        }
        if let Some(ref val) = self.underline {
            val.write_element("w:u", writer)?;
        }
        if let Some(ref val) = self.effect {
            val.write_element("w:effect", writer)?;
        }
        if let Some(ref val) = self.bdr {
            val.write_element("w:bdr", writer)?;
        }
        if let Some(ref val) = self.shading {
            val.write_element("w:shd", writer)?;
        }
        if let Some(ref val) = self.fit_text {
            val.write_element("w:fitText", writer)?;
        }
        if let Some(ref val) = self.vert_align {
            val.write_element("w:vertAlign", writer)?;
        }
        if let Some(ref val) = self.rtl {
            val.write_element("w:rtl", writer)?;
        }
        if let Some(ref val) = self.cs {
            val.write_element("w:cs", writer)?;
        }
        if let Some(ref val) = self.em {
            val.write_element("w:em", writer)?;
        }
        if let Some(ref val) = self.lang {
            val.write_element("w:lang", writer)?;
        }
        if let Some(ref val) = self.east_asian_layout {
            val.write_element("w:eastAsianLayout", writer)?;
        }
        if let Some(ref val) = self.spec_vanish {
            val.write_element("w:specVanish", writer)?;
        }
        if let Some(ref val) = self.o_math {
            val.write_element("w:oMath", writer)?;
        }
        #[cfg(feature = "extra-children")]
        for child in &self.extra_children {
            child.write_to(writer).map_err(SerializeError::from)?;
        }
        Ok(())
    }

    fn is_empty_element(&self) -> bool {
        if self.run_style.is_some() {
            return false;
        }
        if self.fonts.is_some() {
            return false;
        }
        if self.bold.is_some() {
            return false;
        }
        if self.b_cs.is_some() {
            return false;
        }
        if self.italic.is_some() {
            return false;
        }
        if self.i_cs.is_some() {
            return false;
        }
        if self.caps.is_some() {
            return false;
        }
        if self.small_caps.is_some() {
            return false;
        }
        if self.strikethrough.is_some() {
            return false;
        }
        if self.dstrike.is_some() {
            return false;
        }
        if self.outline.is_some() {
            return false;
        }
        if self.shadow.is_some() {
            return false;
        }
        if self.emboss.is_some() {
            return false;
        }
        if self.imprint.is_some() {
            return false;
        }
        if self.no_proof.is_some() {
            return false;
        }
        if self.snap_to_grid.is_some() {
            return false;
        }
        if self.vanish.is_some() {
            return false;
        }
        if self.web_hidden.is_some() {
            return false;
        }
        if self.color.is_some() {
            return false;
        }
        if self.spacing.is_some() {
            return false;
        }
        if self.width.is_some() {
            return false;
        }
        if self.kern.is_some() {
            return false;
        }
        if self.position.is_some() {
            return false;
        }
        if self.size.is_some() {
            return false;
        }
        if self.size_complex_script.is_some() {
            return false;
        }
        if self.highlight.is_some() {
            return false;
        }
        if self.underline.is_some() {
            return false;
        }
        if self.effect.is_some() {
            return false;
        }
        if self.bdr.is_some() {
            return false;
        }
        if self.shading.is_some() {
            return false;
        }
        if self.fit_text.is_some() {
            return false;
        }
        if self.vert_align.is_some() {
            return false;
        }
        if self.rtl.is_some() {
            return false;
        }
        if self.cs.is_some() {
            return false;
        }
        if self.em.is_some() {
            return false;
        }
        if self.lang.is_some() {
            return false;
        }
        if self.east_asian_layout.is_some() {
            return false;
        }
        if self.spec_vanish.is_some() {
            return false;
        }
        if self.o_math.is_some() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for CTParaRPrOriginal {
    fn write_children<W: Write>(&self, writer: &mut Writer<W>) -> Result<(), SerializeError> {
        if let Some(ref val) = self.ins {
            val.write_element("w:ins", writer)?;
        }
        if let Some(ref val) = self.del {
            val.write_element("w:del", writer)?;
        }
        if let Some(ref val) = self.move_from {
            val.write_element("w:moveFrom", writer)?;
        }
        if let Some(ref val) = self.move_to {
            val.write_element("w:moveTo", writer)?;
        }
        if let Some(ref val) = self.run_style {
            val.write_element("w:rStyle", writer)?;
        }
        if let Some(ref val) = self.fonts {
            val.write_element("w:rFonts", writer)?;
        }
        if let Some(ref val) = self.bold {
            val.write_element("w:b", writer)?;
        }
        if let Some(ref val) = self.b_cs {
            val.write_element("w:bCs", writer)?;
        }
        if let Some(ref val) = self.italic {
            val.write_element("w:i", writer)?;
        }
        if let Some(ref val) = self.i_cs {
            val.write_element("w:iCs", writer)?;
        }
        if let Some(ref val) = self.caps {
            val.write_element("w:caps", writer)?;
        }
        if let Some(ref val) = self.small_caps {
            val.write_element("w:smallCaps", writer)?;
        }
        if let Some(ref val) = self.strikethrough {
            val.write_element("w:strike", writer)?;
        }
        if let Some(ref val) = self.dstrike {
            val.write_element("w:dstrike", writer)?;
        }
        if let Some(ref val) = self.outline {
            val.write_element("w:outline", writer)?;
        }
        if let Some(ref val) = self.shadow {
            val.write_element("w:shadow", writer)?;
        }
        if let Some(ref val) = self.emboss {
            val.write_element("w:emboss", writer)?;
        }
        if let Some(ref val) = self.imprint {
            val.write_element("w:imprint", writer)?;
        }
        if let Some(ref val) = self.no_proof {
            val.write_element("w:noProof", writer)?;
        }
        if let Some(ref val) = self.snap_to_grid {
            val.write_element("w:snapToGrid", writer)?;
        }
        if let Some(ref val) = self.vanish {
            val.write_element("w:vanish", writer)?;
        }
        if let Some(ref val) = self.web_hidden {
            val.write_element("w:webHidden", writer)?;
        }
        if let Some(ref val) = self.color {
            val.write_element("w:color", writer)?;
        }
        if let Some(ref val) = self.spacing {
            val.write_element("w:spacing", writer)?;
        }
        if let Some(ref val) = self.width {
            val.write_element("w:w", writer)?;
        }
        if let Some(ref val) = self.kern {
            val.write_element("w:kern", writer)?;
        }
        if let Some(ref val) = self.position {
            val.write_element("w:position", writer)?;
        }
        if let Some(ref val) = self.size {
            val.write_element("w:sz", writer)?;
        }
        if let Some(ref val) = self.size_complex_script {
            val.write_element("w:szCs", writer)?;
        }
        if let Some(ref val) = self.highlight {
            val.write_element("w:highlight", writer)?;
        }
        if let Some(ref val) = self.underline {
            val.write_element("w:u", writer)?;
        }
        if let Some(ref val) = self.effect {
            val.write_element("w:effect", writer)?;
        }
        if let Some(ref val) = self.bdr {
            val.write_element("w:bdr", writer)?;
        }
        if let Some(ref val) = self.shading {
            val.write_element("w:shd", writer)?;
        }
        if let Some(ref val) = self.fit_text {
            val.write_element("w:fitText", writer)?;
        }
        if let Some(ref val) = self.vert_align {
            val.write_element("w:vertAlign", writer)?;
        }
        if let Some(ref val) = self.rtl {
            val.write_element("w:rtl", writer)?;
        }
        if let Some(ref val) = self.cs {
            val.write_element("w:cs", writer)?;
        }
        if let Some(ref val) = self.em {
            val.write_element("w:em", writer)?;
        }
        if let Some(ref val) = self.lang {
            val.write_element("w:lang", writer)?;
        }
        if let Some(ref val) = self.east_asian_layout {
            val.write_element("w:eastAsianLayout", writer)?;
        }
        if let Some(ref val) = self.spec_vanish {
            val.write_element("w:specVanish", writer)?;
        }
        if let Some(ref val) = self.o_math {
            val.write_element("w:oMath", writer)?;
        }
        #[cfg(feature = "extra-children")]
        for child in &self.extra_children {
            child.write_to(writer).map_err(SerializeError::from)?;
        }
        Ok(())
    }

    fn is_empty_element(&self) -> bool {
        if self.ins.is_some() {
            return false;
        }
        if self.del.is_some() {
            return false;
        }
        if self.move_from.is_some() {
            return false;
        }
        if self.move_to.is_some() {
            return false;
        }
        if self.run_style.is_some() {
            return false;
        }
        if self.fonts.is_some() {
            return false;
        }
        if self.bold.is_some() {
            return false;
        }
        if self.b_cs.is_some() {
            return false;
        }
        if self.italic.is_some() {
            return false;
        }
        if self.i_cs.is_some() {
            return false;
        }
        if self.caps.is_some() {
            return false;
        }
        if self.small_caps.is_some() {
            return false;
        }
        if self.strikethrough.is_some() {
            return false;
        }
        if self.dstrike.is_some() {
            return false;
        }
        if self.outline.is_some() {
            return false;
        }
        if self.shadow.is_some() {
            return false;
        }
        if self.emboss.is_some() {
            return false;
        }
        if self.imprint.is_some() {
            return false;
        }
        if self.no_proof.is_some() {
            return false;
        }
        if self.snap_to_grid.is_some() {
            return false;
        }
        if self.vanish.is_some() {
            return false;
        }
        if self.web_hidden.is_some() {
            return false;
        }
        if self.color.is_some() {
            return false;
        }
        if self.spacing.is_some() {
            return false;
        }
        if self.width.is_some() {
            return false;
        }
        if self.kern.is_some() {
            return false;
        }
        if self.position.is_some() {
            return false;
        }
        if self.size.is_some() {
            return false;
        }
        if self.size_complex_script.is_some() {
            return false;
        }
        if self.highlight.is_some() {
            return false;
        }
        if self.underline.is_some() {
            return false;
        }
        if self.effect.is_some() {
            return false;
        }
        if self.bdr.is_some() {
            return false;
        }
        if self.shading.is_some() {
            return false;
        }
        if self.fit_text.is_some() {
            return false;
        }
        if self.vert_align.is_some() {
            return false;
        }
        if self.rtl.is_some() {
            return false;
        }
        if self.cs.is_some() {
            return false;
        }
        if self.em.is_some() {
            return false;
        }
        if self.lang.is_some() {
            return false;
        }
        if self.east_asian_layout.is_some() {
            return false;
        }
        if self.spec_vanish.is_some() {
            return false;
        }
        if self.o_math.is_some() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for CTParaRPr {
    fn write_children<W: Write>(&self, writer: &mut Writer<W>) -> Result<(), SerializeError> {
        if let Some(ref val) = self.ins {
            val.write_element("w:ins", writer)?;
        }
        if let Some(ref val) = self.del {
            val.write_element("w:del", writer)?;
        }
        if let Some(ref val) = self.move_from {
            val.write_element("w:moveFrom", writer)?;
        }
        if let Some(ref val) = self.move_to {
            val.write_element("w:moveTo", writer)?;
        }
        if let Some(ref val) = self.run_style {
            val.write_element("w:rStyle", writer)?;
        }
        if let Some(ref val) = self.fonts {
            val.write_element("w:rFonts", writer)?;
        }
        if let Some(ref val) = self.bold {
            val.write_element("w:b", writer)?;
        }
        if let Some(ref val) = self.b_cs {
            val.write_element("w:bCs", writer)?;
        }
        if let Some(ref val) = self.italic {
            val.write_element("w:i", writer)?;
        }
        if let Some(ref val) = self.i_cs {
            val.write_element("w:iCs", writer)?;
        }
        if let Some(ref val) = self.caps {
            val.write_element("w:caps", writer)?;
        }
        if let Some(ref val) = self.small_caps {
            val.write_element("w:smallCaps", writer)?;
        }
        if let Some(ref val) = self.strikethrough {
            val.write_element("w:strike", writer)?;
        }
        if let Some(ref val) = self.dstrike {
            val.write_element("w:dstrike", writer)?;
        }
        if let Some(ref val) = self.outline {
            val.write_element("w:outline", writer)?;
        }
        if let Some(ref val) = self.shadow {
            val.write_element("w:shadow", writer)?;
        }
        if let Some(ref val) = self.emboss {
            val.write_element("w:emboss", writer)?;
        }
        if let Some(ref val) = self.imprint {
            val.write_element("w:imprint", writer)?;
        }
        if let Some(ref val) = self.no_proof {
            val.write_element("w:noProof", writer)?;
        }
        if let Some(ref val) = self.snap_to_grid {
            val.write_element("w:snapToGrid", writer)?;
        }
        if let Some(ref val) = self.vanish {
            val.write_element("w:vanish", writer)?;
        }
        if let Some(ref val) = self.web_hidden {
            val.write_element("w:webHidden", writer)?;
        }
        if let Some(ref val) = self.color {
            val.write_element("w:color", writer)?;
        }
        if let Some(ref val) = self.spacing {
            val.write_element("w:spacing", writer)?;
        }
        if let Some(ref val) = self.width {
            val.write_element("w:w", writer)?;
        }
        if let Some(ref val) = self.kern {
            val.write_element("w:kern", writer)?;
        }
        if let Some(ref val) = self.position {
            val.write_element("w:position", writer)?;
        }
        if let Some(ref val) = self.size {
            val.write_element("w:sz", writer)?;
        }
        if let Some(ref val) = self.size_complex_script {
            val.write_element("w:szCs", writer)?;
        }
        if let Some(ref val) = self.highlight {
            val.write_element("w:highlight", writer)?;
        }
        if let Some(ref val) = self.underline {
            val.write_element("w:u", writer)?;
        }
        if let Some(ref val) = self.effect {
            val.write_element("w:effect", writer)?;
        }
        if let Some(ref val) = self.bdr {
            val.write_element("w:bdr", writer)?;
        }
        if let Some(ref val) = self.shading {
            val.write_element("w:shd", writer)?;
        }
        if let Some(ref val) = self.fit_text {
            val.write_element("w:fitText", writer)?;
        }
        if let Some(ref val) = self.vert_align {
            val.write_element("w:vertAlign", writer)?;
        }
        if let Some(ref val) = self.rtl {
            val.write_element("w:rtl", writer)?;
        }
        if let Some(ref val) = self.cs {
            val.write_element("w:cs", writer)?;
        }
        if let Some(ref val) = self.em {
            val.write_element("w:em", writer)?;
        }
        if let Some(ref val) = self.lang {
            val.write_element("w:lang", writer)?;
        }
        if let Some(ref val) = self.east_asian_layout {
            val.write_element("w:eastAsianLayout", writer)?;
        }
        if let Some(ref val) = self.spec_vanish {
            val.write_element("w:specVanish", writer)?;
        }
        if let Some(ref val) = self.o_math {
            val.write_element("w:oMath", writer)?;
        }
        if let Some(ref val) = self.r_pr_change {
            val.write_element("w:rPrChange", writer)?;
        }
        #[cfg(feature = "extra-children")]
        for child in &self.extra_children {
            child.write_to(writer).map_err(SerializeError::from)?;
        }
        Ok(())
    }

    fn is_empty_element(&self) -> bool {
        if self.ins.is_some() {
            return false;
        }
        if self.del.is_some() {
            return false;
        }
        if self.move_from.is_some() {
            return false;
        }
        if self.move_to.is_some() {
            return false;
        }
        if self.run_style.is_some() {
            return false;
        }
        if self.fonts.is_some() {
            return false;
        }
        if self.bold.is_some() {
            return false;
        }
        if self.b_cs.is_some() {
            return false;
        }
        if self.italic.is_some() {
            return false;
        }
        if self.i_cs.is_some() {
            return false;
        }
        if self.caps.is_some() {
            return false;
        }
        if self.small_caps.is_some() {
            return false;
        }
        if self.strikethrough.is_some() {
            return false;
        }
        if self.dstrike.is_some() {
            return false;
        }
        if self.outline.is_some() {
            return false;
        }
        if self.shadow.is_some() {
            return false;
        }
        if self.emboss.is_some() {
            return false;
        }
        if self.imprint.is_some() {
            return false;
        }
        if self.no_proof.is_some() {
            return false;
        }
        if self.snap_to_grid.is_some() {
            return false;
        }
        if self.vanish.is_some() {
            return false;
        }
        if self.web_hidden.is_some() {
            return false;
        }
        if self.color.is_some() {
            return false;
        }
        if self.spacing.is_some() {
            return false;
        }
        if self.width.is_some() {
            return false;
        }
        if self.kern.is_some() {
            return false;
        }
        if self.position.is_some() {
            return false;
        }
        if self.size.is_some() {
            return false;
        }
        if self.size_complex_script.is_some() {
            return false;
        }
        if self.highlight.is_some() {
            return false;
        }
        if self.underline.is_some() {
            return false;
        }
        if self.effect.is_some() {
            return false;
        }
        if self.bdr.is_some() {
            return false;
        }
        if self.shading.is_some() {
            return false;
        }
        if self.fit_text.is_some() {
            return false;
        }
        if self.vert_align.is_some() {
            return false;
        }
        if self.rtl.is_some() {
            return false;
        }
        if self.cs.is_some() {
            return false;
        }
        if self.em.is_some() {
            return false;
        }
        if self.lang.is_some() {
            return false;
        }
        if self.east_asian_layout.is_some() {
            return false;
        }
        if self.spec_vanish.is_some() {
            return false;
        }
        if self.o_math.is_some() {
            return false;
        }
        if self.r_pr_change.is_some() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for EGParaRPrTrackChanges {
    fn write_children<W: Write>(&self, writer: &mut Writer<W>) -> Result<(), SerializeError> {
        if let Some(ref val) = self.ins {
            val.write_element("w:ins", writer)?;
        }
        if let Some(ref val) = self.del {
            val.write_element("w:del", writer)?;
        }
        if let Some(ref val) = self.move_from {
            val.write_element("w:moveFrom", writer)?;
        }
        if let Some(ref val) = self.move_to {
            val.write_element("w:moveTo", writer)?;
        }
        #[cfg(feature = "extra-children")]
        for child in &self.extra_children {
            child.write_to(writer).map_err(SerializeError::from)?;
        }
        Ok(())
    }

    fn is_empty_element(&self) -> bool {
        if self.ins.is_some() {
            return false;
        }
        if self.del.is_some() {
            return false;
        }
        if self.move_from.is_some() {
            return false;
        }
        if self.move_to.is_some() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for CTAltChunk {
    fn write_children<W: Write>(&self, writer: &mut Writer<W>) -> Result<(), SerializeError> {
        if let Some(ref val) = self.alt_chunk_pr {
            val.write_element("w:altChunkPr", writer)?;
        }
        #[cfg(feature = "extra-children")]
        for child in &self.extra_children {
            child.write_to(writer).map_err(SerializeError::from)?;
        }
        Ok(())
    }

    fn is_empty_element(&self) -> bool {
        if self.alt_chunk_pr.is_some() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for CTAltChunkPr {
    fn write_children<W: Write>(&self, writer: &mut Writer<W>) -> Result<(), SerializeError> {
        if let Some(ref val) = self.match_src {
            val.write_element("w:matchSrc", writer)?;
        }
        #[cfg(feature = "extra-children")]
        for child in &self.extra_children {
            child.write_to(writer).map_err(SerializeError::from)?;
        }
        Ok(())
    }

    fn is_empty_element(&self) -> bool {
        if self.match_src.is_some() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for CTRubyAlign {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        {
            let val = &self.value;
            {
                let s = val.to_string();
                start.push_attribute(("w:val", s.as_str()));
            }
        }
        #[cfg(feature = "extra-attrs")]
        for (key, value) in &self.extra_attrs {
            start.push_attribute((key.as_str(), value.as_str()));
        }
        start
    }

    fn is_empty_element(&self) -> bool {
        true
    }
}

impl ToXml for CTRubyPr {
    fn write_children<W: Write>(&self, writer: &mut Writer<W>) -> Result<(), SerializeError> {
        {
            let val = &self.ruby_align;
            val.write_element("w:rubyAlign", writer)?;
        }
        {
            let val = &self.hps;
            val.write_element("w:hps", writer)?;
        }
        {
            let val = &self.hps_raise;
            val.write_element("w:hpsRaise", writer)?;
        }
        {
            let val = &self.hps_base_text;
            val.write_element("w:hpsBaseText", writer)?;
        }
        {
            let val = &self.lid;
            val.write_element("w:lid", writer)?;
        }
        if let Some(ref val) = self.dirty {
            val.write_element("w:dirty", writer)?;
        }
        #[cfg(feature = "extra-children")]
        for child in &self.extra_children {
            child.write_to(writer).map_err(SerializeError::from)?;
        }
        Ok(())
    }

    fn is_empty_element(&self) -> bool {
        false
    }
}

impl ToXml for EGRubyContent {
    fn write_element<W: Write>(
        &self,
        _tag: &str,
        writer: &mut Writer<W>,
    ) -> Result<(), SerializeError> {
        match self {
            Self::R(inner) => inner.write_element("w:r", writer)?,
            Self::ProofErr(inner) => inner.write_element("w:proofErr", writer)?,
            Self::PermStart(inner) => inner.write_element("w:permStart", writer)?,
            Self::PermEnd(inner) => inner.write_element("w:permEnd", writer)?,
            Self::BookmarkStart(inner) => inner.write_element("w:bookmarkStart", writer)?,
            Self::BookmarkEnd(inner) => inner.write_element("w:bookmarkEnd", writer)?,
            Self::MoveFromRangeStart(inner) => {
                inner.write_element("w:moveFromRangeStart", writer)?
            }
            Self::MoveFromRangeEnd(inner) => inner.write_element("w:moveFromRangeEnd", writer)?,
            Self::MoveToRangeStart(inner) => inner.write_element("w:moveToRangeStart", writer)?,
            Self::MoveToRangeEnd(inner) => inner.write_element("w:moveToRangeEnd", writer)?,
            Self::CommentRangeStart(inner) => inner.write_element("w:commentRangeStart", writer)?,
            Self::CommentRangeEnd(inner) => inner.write_element("w:commentRangeEnd", writer)?,
            Self::CustomXmlInsRangeStart(inner) => {
                inner.write_element("w:customXmlInsRangeStart", writer)?
            }
            Self::CustomXmlInsRangeEnd(inner) => {
                inner.write_element("w:customXmlInsRangeEnd", writer)?
            }
            Self::CustomXmlDelRangeStart(inner) => {
                inner.write_element("w:customXmlDelRangeStart", writer)?
            }
            Self::CustomXmlDelRangeEnd(inner) => {
                inner.write_element("w:customXmlDelRangeEnd", writer)?
            }
            Self::CustomXmlMoveFromRangeStart(inner) => {
                inner.write_element("w:customXmlMoveFromRangeStart", writer)?
            }
            Self::CustomXmlMoveFromRangeEnd(inner) => {
                inner.write_element("w:customXmlMoveFromRangeEnd", writer)?
            }
            Self::CustomXmlMoveToRangeStart(inner) => {
                inner.write_element("w:customXmlMoveToRangeStart", writer)?
            }
            Self::CustomXmlMoveToRangeEnd(inner) => {
                inner.write_element("w:customXmlMoveToRangeEnd", writer)?
            }
            Self::Ins(inner) => inner.write_element("w:ins", writer)?,
            Self::Del(inner) => inner.write_element("w:del", writer)?,
            Self::MoveFrom(inner) => inner.write_element("w:moveFrom", writer)?,
            Self::MoveTo(inner) => inner.write_element("w:moveTo", writer)?,
        }
        Ok(())
    }
}

impl ToXml for CTRubyContent {
    fn write_children<W: Write>(&self, writer: &mut Writer<W>) -> Result<(), SerializeError> {
        for item in &self.ruby_content {
            item.write_element("", writer)?;
        }
        #[cfg(feature = "extra-children")]
        for child in &self.extra_children {
            child.write_to(writer).map_err(SerializeError::from)?;
        }
        Ok(())
    }

    fn is_empty_element(&self) -> bool {
        if !self.ruby_content.is_empty() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for CTRuby {
    fn write_children<W: Write>(&self, writer: &mut Writer<W>) -> Result<(), SerializeError> {
        {
            let val = &self.ruby_pr;
            val.write_element("w:rubyPr", writer)?;
        }
        {
            let val = &self.rt;
            val.write_element("w:rt", writer)?;
        }
        {
            let val = &self.ruby_base;
            val.write_element("w:rubyBase", writer)?;
        }
        #[cfg(feature = "extra-children")]
        for child in &self.extra_children {
            child.write_to(writer).map_err(SerializeError::from)?;
        }
        Ok(())
    }

    fn is_empty_element(&self) -> bool {
        false
    }
}

impl ToXml for CTLock {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        if let Some(ref val) = self.value {
            {
                let s = val.to_string();
                start.push_attribute(("w:val", s.as_str()));
            }
        }
        #[cfg(feature = "extra-attrs")]
        for (key, value) in &self.extra_attrs {
            start.push_attribute((key.as_str(), value.as_str()));
        }
        start
    }

    fn is_empty_element(&self) -> bool {
        true
    }
}

impl ToXml for CTSdtListItem {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        if let Some(ref val) = self.display_text {
            start.push_attribute(("w:displayText", val.as_str()));
        }
        if let Some(ref val) = self.value {
            start.push_attribute(("w:value", val.as_str()));
        }
        #[cfg(feature = "extra-attrs")]
        for (key, value) in &self.extra_attrs {
            start.push_attribute((key.as_str(), value.as_str()));
        }
        start
    }

    fn is_empty_element(&self) -> bool {
        true
    }
}

impl ToXml for CTSdtDateMappingType {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        if let Some(ref val) = self.value {
            {
                let s = val.to_string();
                start.push_attribute(("w:val", s.as_str()));
            }
        }
        #[cfg(feature = "extra-attrs")]
        for (key, value) in &self.extra_attrs {
            start.push_attribute((key.as_str(), value.as_str()));
        }
        start
    }

    fn is_empty_element(&self) -> bool {
        true
    }
}

impl ToXml for CTCalendarType {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        if let Some(ref val) = self.value {
            {
                let s = val.to_string();
                start.push_attribute(("w:val", s.as_str()));
            }
        }
        #[cfg(feature = "extra-attrs")]
        for (key, value) in &self.extra_attrs {
            start.push_attribute((key.as_str(), value.as_str()));
        }
        start
    }

    fn is_empty_element(&self) -> bool {
        true
    }
}

impl ToXml for CTSdtDate {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        if let Some(ref val) = self.full_date {
            start.push_attribute(("w:fullDate", val.as_str()));
        }
        #[cfg(feature = "extra-attrs")]
        for (key, value) in &self.extra_attrs {
            start.push_attribute((key.as_str(), value.as_str()));
        }
        start
    }

    fn write_children<W: Write>(&self, writer: &mut Writer<W>) -> Result<(), SerializeError> {
        if let Some(ref val) = self.date_format {
            val.write_element("w:dateFormat", writer)?;
        }
        if let Some(ref val) = self.lid {
            val.write_element("w:lid", writer)?;
        }
        if let Some(ref val) = self.store_mapped_data_as {
            val.write_element("w:storeMappedDataAs", writer)?;
        }
        if let Some(ref val) = self.calendar {
            val.write_element("w:calendar", writer)?;
        }
        #[cfg(feature = "extra-children")]
        for child in &self.extra_children {
            child.write_to(writer).map_err(SerializeError::from)?;
        }
        Ok(())
    }

    fn is_empty_element(&self) -> bool {
        if self.date_format.is_some() {
            return false;
        }
        if self.lid.is_some() {
            return false;
        }
        if self.store_mapped_data_as.is_some() {
            return false;
        }
        if self.calendar.is_some() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for CTSdtComboBox {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        if let Some(ref val) = self.last_value {
            start.push_attribute(("w:lastValue", val.as_str()));
        }
        #[cfg(feature = "extra-attrs")]
        for (key, value) in &self.extra_attrs {
            start.push_attribute((key.as_str(), value.as_str()));
        }
        start
    }

    fn write_children<W: Write>(&self, writer: &mut Writer<W>) -> Result<(), SerializeError> {
        for item in &self.list_item {
            item.write_element("w:listItem", writer)?;
        }
        #[cfg(feature = "extra-children")]
        for child in &self.extra_children {
            child.write_to(writer).map_err(SerializeError::from)?;
        }
        Ok(())
    }

    fn is_empty_element(&self) -> bool {
        if !self.list_item.is_empty() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for CTSdtDocPart {
    fn write_children<W: Write>(&self, writer: &mut Writer<W>) -> Result<(), SerializeError> {
        if let Some(ref val) = self.doc_part_gallery {
            val.write_element("w:docPartGallery", writer)?;
        }
        if let Some(ref val) = self.doc_part_category {
            val.write_element("w:docPartCategory", writer)?;
        }
        if let Some(ref val) = self.doc_part_unique {
            val.write_element("w:docPartUnique", writer)?;
        }
        #[cfg(feature = "extra-children")]
        for child in &self.extra_children {
            child.write_to(writer).map_err(SerializeError::from)?;
        }
        Ok(())
    }

    fn is_empty_element(&self) -> bool {
        if self.doc_part_gallery.is_some() {
            return false;
        }
        if self.doc_part_category.is_some() {
            return false;
        }
        if self.doc_part_unique.is_some() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for CTSdtDropDownList {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        if let Some(ref val) = self.last_value {
            start.push_attribute(("w:lastValue", val.as_str()));
        }
        #[cfg(feature = "extra-attrs")]
        for (key, value) in &self.extra_attrs {
            start.push_attribute((key.as_str(), value.as_str()));
        }
        start
    }

    fn write_children<W: Write>(&self, writer: &mut Writer<W>) -> Result<(), SerializeError> {
        for item in &self.list_item {
            item.write_element("w:listItem", writer)?;
        }
        #[cfg(feature = "extra-children")]
        for child in &self.extra_children {
            child.write_to(writer).map_err(SerializeError::from)?;
        }
        Ok(())
    }

    fn is_empty_element(&self) -> bool {
        if !self.list_item.is_empty() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for CTSdtText {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        if let Some(ref val) = self.multi_line {
            {
                let s = val.to_string();
                start.push_attribute(("w:multiLine", s.as_str()));
            }
        }
        #[cfg(feature = "extra-attrs")]
        for (key, value) in &self.extra_attrs {
            start.push_attribute((key.as_str(), value.as_str()));
        }
        start
    }

    fn is_empty_element(&self) -> bool {
        true
    }
}

impl ToXml for CTDataBinding {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        if let Some(ref val) = self.prefix_mappings {
            start.push_attribute(("w:prefixMappings", val.as_str()));
        }
        {
            let val = &self.xpath;
            start.push_attribute(("w:xpath", val.as_str()));
        }
        {
            let val = &self.store_item_i_d;
            start.push_attribute(("w:storeItemID", val.as_str()));
        }
        #[cfg(feature = "extra-attrs")]
        for (key, value) in &self.extra_attrs {
            start.push_attribute((key.as_str(), value.as_str()));
        }
        start
    }

    fn is_empty_element(&self) -> bool {
        true
    }
}

impl ToXml for CTSdtPr {
    fn write_children<W: Write>(&self, writer: &mut Writer<W>) -> Result<(), SerializeError> {
        if let Some(ref val) = self.r_pr {
            val.write_element("w:rPr", writer)?;
        }
        if let Some(ref val) = self.alias {
            val.write_element("w:alias", writer)?;
        }
        if let Some(ref val) = self.tag {
            val.write_element("w:tag", writer)?;
        }
        if let Some(ref val) = self.id {
            val.write_element("w:id", writer)?;
        }
        if let Some(ref val) = self.lock {
            val.write_element("w:lock", writer)?;
        }
        if let Some(ref val) = self.placeholder {
            val.write_element("w:placeholder", writer)?;
        }
        if let Some(ref val) = self.temporary {
            val.write_element("w:temporary", writer)?;
        }
        if let Some(ref val) = self.showing_plc_hdr {
            val.write_element("w:showingPlcHdr", writer)?;
        }
        if let Some(ref val) = self.data_binding {
            val.write_element("w:dataBinding", writer)?;
        }
        if let Some(ref val) = self.label {
            val.write_element("w:label", writer)?;
        }
        if let Some(ref val) = self.tab_index {
            val.write_element("w:tabIndex", writer)?;
        }
        #[cfg(feature = "extra-children")]
        for child in &self.extra_children {
            child.write_to(writer).map_err(SerializeError::from)?;
        }
        Ok(())
    }

    fn is_empty_element(&self) -> bool {
        if self.r_pr.is_some() {
            return false;
        }
        if self.alias.is_some() {
            return false;
        }
        if self.tag.is_some() {
            return false;
        }
        if self.id.is_some() {
            return false;
        }
        if self.lock.is_some() {
            return false;
        }
        if self.placeholder.is_some() {
            return false;
        }
        if self.temporary.is_some() {
            return false;
        }
        if self.showing_plc_hdr.is_some() {
            return false;
        }
        if self.data_binding.is_some() {
            return false;
        }
        if self.label.is_some() {
            return false;
        }
        if self.tab_index.is_some() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for CTSdtEndPr {
    fn write_children<W: Write>(&self, writer: &mut Writer<W>) -> Result<(), SerializeError> {
        #[cfg(feature = "extra-children")]
        for child in &self.extra_children {
            child.write_to(writer).map_err(SerializeError::from)?;
        }
        Ok(())
    }

    fn is_empty_element(&self) -> bool {
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for EGContentRunContent {
    fn write_element<W: Write>(
        &self,
        _tag: &str,
        writer: &mut Writer<W>,
    ) -> Result<(), SerializeError> {
        match self {
            Self::CustomXml(inner) => inner.write_element("w:customXml", writer)?,
            Self::SmartTag(inner) => inner.write_element("w:smartTag", writer)?,
            Self::Sdt(inner) => inner.write_element("w:sdt", writer)?,
            Self::Dir(inner) => inner.write_element("w:dir", writer)?,
            Self::Bdo(inner) => inner.write_element("w:bdo", writer)?,
            Self::R(inner) => inner.write_element("w:r", writer)?,
            Self::ProofErr(inner) => inner.write_element("w:proofErr", writer)?,
            Self::PermStart(inner) => inner.write_element("w:permStart", writer)?,
            Self::PermEnd(inner) => inner.write_element("w:permEnd", writer)?,
            Self::BookmarkStart(inner) => inner.write_element("w:bookmarkStart", writer)?,
            Self::BookmarkEnd(inner) => inner.write_element("w:bookmarkEnd", writer)?,
            Self::MoveFromRangeStart(inner) => {
                inner.write_element("w:moveFromRangeStart", writer)?
            }
            Self::MoveFromRangeEnd(inner) => inner.write_element("w:moveFromRangeEnd", writer)?,
            Self::MoveToRangeStart(inner) => inner.write_element("w:moveToRangeStart", writer)?,
            Self::MoveToRangeEnd(inner) => inner.write_element("w:moveToRangeEnd", writer)?,
            Self::CommentRangeStart(inner) => inner.write_element("w:commentRangeStart", writer)?,
            Self::CommentRangeEnd(inner) => inner.write_element("w:commentRangeEnd", writer)?,
            Self::CustomXmlInsRangeStart(inner) => {
                inner.write_element("w:customXmlInsRangeStart", writer)?
            }
            Self::CustomXmlInsRangeEnd(inner) => {
                inner.write_element("w:customXmlInsRangeEnd", writer)?
            }
            Self::CustomXmlDelRangeStart(inner) => {
                inner.write_element("w:customXmlDelRangeStart", writer)?
            }
            Self::CustomXmlDelRangeEnd(inner) => {
                inner.write_element("w:customXmlDelRangeEnd", writer)?
            }
            Self::CustomXmlMoveFromRangeStart(inner) => {
                inner.write_element("w:customXmlMoveFromRangeStart", writer)?
            }
            Self::CustomXmlMoveFromRangeEnd(inner) => {
                inner.write_element("w:customXmlMoveFromRangeEnd", writer)?
            }
            Self::CustomXmlMoveToRangeStart(inner) => {
                inner.write_element("w:customXmlMoveToRangeStart", writer)?
            }
            Self::CustomXmlMoveToRangeEnd(inner) => {
                inner.write_element("w:customXmlMoveToRangeEnd", writer)?
            }
            Self::Ins(inner) => inner.write_element("w:ins", writer)?,
            Self::Del(inner) => inner.write_element("w:del", writer)?,
            Self::MoveFrom(inner) => inner.write_element("w:moveFrom", writer)?,
            Self::MoveTo(inner) => inner.write_element("w:moveTo", writer)?,
        }
        Ok(())
    }
}

impl ToXml for CTDirContentRun {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        if let Some(ref val) = self.value {
            {
                let s = val.to_string();
                start.push_attribute(("w:val", s.as_str()));
            }
        }
        #[cfg(feature = "extra-attrs")]
        for (key, value) in &self.extra_attrs {
            start.push_attribute((key.as_str(), value.as_str()));
        }
        start
    }

    fn write_children<W: Write>(&self, writer: &mut Writer<W>) -> Result<(), SerializeError> {
        for item in &self.p_content {
            item.write_element("", writer)?;
        }
        #[cfg(feature = "extra-children")]
        for child in &self.extra_children {
            child.write_to(writer).map_err(SerializeError::from)?;
        }
        Ok(())
    }

    fn is_empty_element(&self) -> bool {
        if !self.p_content.is_empty() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for CTBdoContentRun {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        if let Some(ref val) = self.value {
            {
                let s = val.to_string();
                start.push_attribute(("w:val", s.as_str()));
            }
        }
        #[cfg(feature = "extra-attrs")]
        for (key, value) in &self.extra_attrs {
            start.push_attribute((key.as_str(), value.as_str()));
        }
        start
    }

    fn write_children<W: Write>(&self, writer: &mut Writer<W>) -> Result<(), SerializeError> {
        for item in &self.p_content {
            item.write_element("", writer)?;
        }
        #[cfg(feature = "extra-children")]
        for child in &self.extra_children {
            child.write_to(writer).map_err(SerializeError::from)?;
        }
        Ok(())
    }

    fn is_empty_element(&self) -> bool {
        if !self.p_content.is_empty() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for CTSdtContentRun {
    fn write_children<W: Write>(&self, writer: &mut Writer<W>) -> Result<(), SerializeError> {
        for item in &self.p_content {
            item.write_element("", writer)?;
        }
        #[cfg(feature = "extra-children")]
        for child in &self.extra_children {
            child.write_to(writer).map_err(SerializeError::from)?;
        }
        Ok(())
    }

    fn is_empty_element(&self) -> bool {
        if !self.p_content.is_empty() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for EGContentBlockContent {
    fn write_element<W: Write>(
        &self,
        _tag: &str,
        writer: &mut Writer<W>,
    ) -> Result<(), SerializeError> {
        match self {
            Self::CustomXml(inner) => inner.write_element("w:customXml", writer)?,
            Self::Sdt(inner) => inner.write_element("w:sdt", writer)?,
            Self::P(inner) => inner.write_element("w:p", writer)?,
            Self::Tbl(inner) => inner.write_element("w:tbl", writer)?,
            Self::ProofErr(inner) => inner.write_element("w:proofErr", writer)?,
            Self::PermStart(inner) => inner.write_element("w:permStart", writer)?,
            Self::PermEnd(inner) => inner.write_element("w:permEnd", writer)?,
            Self::BookmarkStart(inner) => inner.write_element("w:bookmarkStart", writer)?,
            Self::BookmarkEnd(inner) => inner.write_element("w:bookmarkEnd", writer)?,
            Self::MoveFromRangeStart(inner) => {
                inner.write_element("w:moveFromRangeStart", writer)?
            }
            Self::MoveFromRangeEnd(inner) => inner.write_element("w:moveFromRangeEnd", writer)?,
            Self::MoveToRangeStart(inner) => inner.write_element("w:moveToRangeStart", writer)?,
            Self::MoveToRangeEnd(inner) => inner.write_element("w:moveToRangeEnd", writer)?,
            Self::CommentRangeStart(inner) => inner.write_element("w:commentRangeStart", writer)?,
            Self::CommentRangeEnd(inner) => inner.write_element("w:commentRangeEnd", writer)?,
            Self::CustomXmlInsRangeStart(inner) => {
                inner.write_element("w:customXmlInsRangeStart", writer)?
            }
            Self::CustomXmlInsRangeEnd(inner) => {
                inner.write_element("w:customXmlInsRangeEnd", writer)?
            }
            Self::CustomXmlDelRangeStart(inner) => {
                inner.write_element("w:customXmlDelRangeStart", writer)?
            }
            Self::CustomXmlDelRangeEnd(inner) => {
                inner.write_element("w:customXmlDelRangeEnd", writer)?
            }
            Self::CustomXmlMoveFromRangeStart(inner) => {
                inner.write_element("w:customXmlMoveFromRangeStart", writer)?
            }
            Self::CustomXmlMoveFromRangeEnd(inner) => {
                inner.write_element("w:customXmlMoveFromRangeEnd", writer)?
            }
            Self::CustomXmlMoveToRangeStart(inner) => {
                inner.write_element("w:customXmlMoveToRangeStart", writer)?
            }
            Self::CustomXmlMoveToRangeEnd(inner) => {
                inner.write_element("w:customXmlMoveToRangeEnd", writer)?
            }
            Self::Ins(inner) => inner.write_element("w:ins", writer)?,
            Self::Del(inner) => inner.write_element("w:del", writer)?,
            Self::MoveFrom(inner) => inner.write_element("w:moveFrom", writer)?,
            Self::MoveTo(inner) => inner.write_element("w:moveTo", writer)?,
        }
        Ok(())
    }
}

impl ToXml for CTSdtContentBlock {
    fn write_children<W: Write>(&self, writer: &mut Writer<W>) -> Result<(), SerializeError> {
        for item in &self.content_block_content {
            item.write_element("", writer)?;
        }
        #[cfg(feature = "extra-children")]
        for child in &self.extra_children {
            child.write_to(writer).map_err(SerializeError::from)?;
        }
        Ok(())
    }

    fn is_empty_element(&self) -> bool {
        if !self.content_block_content.is_empty() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for EGContentRowContent {
    fn write_element<W: Write>(
        &self,
        _tag: &str,
        writer: &mut Writer<W>,
    ) -> Result<(), SerializeError> {
        match self {
            Self::Tr(inner) => inner.write_element("w:tr", writer)?,
            Self::CustomXml(inner) => inner.write_element("w:customXml", writer)?,
            Self::Sdt(inner) => inner.write_element("w:sdt", writer)?,
            Self::ProofErr(inner) => inner.write_element("w:proofErr", writer)?,
            Self::PermStart(inner) => inner.write_element("w:permStart", writer)?,
            Self::PermEnd(inner) => inner.write_element("w:permEnd", writer)?,
            Self::BookmarkStart(inner) => inner.write_element("w:bookmarkStart", writer)?,
            Self::BookmarkEnd(inner) => inner.write_element("w:bookmarkEnd", writer)?,
            Self::MoveFromRangeStart(inner) => {
                inner.write_element("w:moveFromRangeStart", writer)?
            }
            Self::MoveFromRangeEnd(inner) => inner.write_element("w:moveFromRangeEnd", writer)?,
            Self::MoveToRangeStart(inner) => inner.write_element("w:moveToRangeStart", writer)?,
            Self::MoveToRangeEnd(inner) => inner.write_element("w:moveToRangeEnd", writer)?,
            Self::CommentRangeStart(inner) => inner.write_element("w:commentRangeStart", writer)?,
            Self::CommentRangeEnd(inner) => inner.write_element("w:commentRangeEnd", writer)?,
            Self::CustomXmlInsRangeStart(inner) => {
                inner.write_element("w:customXmlInsRangeStart", writer)?
            }
            Self::CustomXmlInsRangeEnd(inner) => {
                inner.write_element("w:customXmlInsRangeEnd", writer)?
            }
            Self::CustomXmlDelRangeStart(inner) => {
                inner.write_element("w:customXmlDelRangeStart", writer)?
            }
            Self::CustomXmlDelRangeEnd(inner) => {
                inner.write_element("w:customXmlDelRangeEnd", writer)?
            }
            Self::CustomXmlMoveFromRangeStart(inner) => {
                inner.write_element("w:customXmlMoveFromRangeStart", writer)?
            }
            Self::CustomXmlMoveFromRangeEnd(inner) => {
                inner.write_element("w:customXmlMoveFromRangeEnd", writer)?
            }
            Self::CustomXmlMoveToRangeStart(inner) => {
                inner.write_element("w:customXmlMoveToRangeStart", writer)?
            }
            Self::CustomXmlMoveToRangeEnd(inner) => {
                inner.write_element("w:customXmlMoveToRangeEnd", writer)?
            }
            Self::Ins(inner) => inner.write_element("w:ins", writer)?,
            Self::Del(inner) => inner.write_element("w:del", writer)?,
            Self::MoveFrom(inner) => inner.write_element("w:moveFrom", writer)?,
            Self::MoveTo(inner) => inner.write_element("w:moveTo", writer)?,
        }
        Ok(())
    }
}

impl ToXml for CTSdtContentRow {
    fn write_children<W: Write>(&self, writer: &mut Writer<W>) -> Result<(), SerializeError> {
        for item in &self.content_row_content {
            item.write_element("", writer)?;
        }
        #[cfg(feature = "extra-children")]
        for child in &self.extra_children {
            child.write_to(writer).map_err(SerializeError::from)?;
        }
        Ok(())
    }

    fn is_empty_element(&self) -> bool {
        if !self.content_row_content.is_empty() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for EGContentCellContent {
    fn write_element<W: Write>(
        &self,
        _tag: &str,
        writer: &mut Writer<W>,
    ) -> Result<(), SerializeError> {
        match self {
            Self::Tc(inner) => inner.write_element("w:tc", writer)?,
            Self::CustomXml(inner) => inner.write_element("w:customXml", writer)?,
            Self::Sdt(inner) => inner.write_element("w:sdt", writer)?,
            Self::ProofErr(inner) => inner.write_element("w:proofErr", writer)?,
            Self::PermStart(inner) => inner.write_element("w:permStart", writer)?,
            Self::PermEnd(inner) => inner.write_element("w:permEnd", writer)?,
            Self::BookmarkStart(inner) => inner.write_element("w:bookmarkStart", writer)?,
            Self::BookmarkEnd(inner) => inner.write_element("w:bookmarkEnd", writer)?,
            Self::MoveFromRangeStart(inner) => {
                inner.write_element("w:moveFromRangeStart", writer)?
            }
            Self::MoveFromRangeEnd(inner) => inner.write_element("w:moveFromRangeEnd", writer)?,
            Self::MoveToRangeStart(inner) => inner.write_element("w:moveToRangeStart", writer)?,
            Self::MoveToRangeEnd(inner) => inner.write_element("w:moveToRangeEnd", writer)?,
            Self::CommentRangeStart(inner) => inner.write_element("w:commentRangeStart", writer)?,
            Self::CommentRangeEnd(inner) => inner.write_element("w:commentRangeEnd", writer)?,
            Self::CustomXmlInsRangeStart(inner) => {
                inner.write_element("w:customXmlInsRangeStart", writer)?
            }
            Self::CustomXmlInsRangeEnd(inner) => {
                inner.write_element("w:customXmlInsRangeEnd", writer)?
            }
            Self::CustomXmlDelRangeStart(inner) => {
                inner.write_element("w:customXmlDelRangeStart", writer)?
            }
            Self::CustomXmlDelRangeEnd(inner) => {
                inner.write_element("w:customXmlDelRangeEnd", writer)?
            }
            Self::CustomXmlMoveFromRangeStart(inner) => {
                inner.write_element("w:customXmlMoveFromRangeStart", writer)?
            }
            Self::CustomXmlMoveFromRangeEnd(inner) => {
                inner.write_element("w:customXmlMoveFromRangeEnd", writer)?
            }
            Self::CustomXmlMoveToRangeStart(inner) => {
                inner.write_element("w:customXmlMoveToRangeStart", writer)?
            }
            Self::CustomXmlMoveToRangeEnd(inner) => {
                inner.write_element("w:customXmlMoveToRangeEnd", writer)?
            }
            Self::Ins(inner) => inner.write_element("w:ins", writer)?,
            Self::Del(inner) => inner.write_element("w:del", writer)?,
            Self::MoveFrom(inner) => inner.write_element("w:moveFrom", writer)?,
            Self::MoveTo(inner) => inner.write_element("w:moveTo", writer)?,
        }
        Ok(())
    }
}

impl ToXml for CTSdtContentCell {
    fn write_children<W: Write>(&self, writer: &mut Writer<W>) -> Result<(), SerializeError> {
        for item in &self.content_cell_content {
            item.write_element("", writer)?;
        }
        #[cfg(feature = "extra-children")]
        for child in &self.extra_children {
            child.write_to(writer).map_err(SerializeError::from)?;
        }
        Ok(())
    }

    fn is_empty_element(&self) -> bool {
        if !self.content_cell_content.is_empty() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for CTSdtBlock {
    fn write_children<W: Write>(&self, writer: &mut Writer<W>) -> Result<(), SerializeError> {
        if let Some(ref val) = self.sdt_pr {
            val.write_element("w:sdtPr", writer)?;
        }
        if let Some(ref val) = self.sdt_end_pr {
            val.write_element("w:sdtEndPr", writer)?;
        }
        if let Some(ref val) = self.sdt_content {
            val.write_element("w:sdtContent", writer)?;
        }
        #[cfg(feature = "extra-children")]
        for child in &self.extra_children {
            child.write_to(writer).map_err(SerializeError::from)?;
        }
        Ok(())
    }

    fn is_empty_element(&self) -> bool {
        if self.sdt_pr.is_some() {
            return false;
        }
        if self.sdt_end_pr.is_some() {
            return false;
        }
        if self.sdt_content.is_some() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for CTSdtRun {
    fn write_children<W: Write>(&self, writer: &mut Writer<W>) -> Result<(), SerializeError> {
        if let Some(ref val) = self.sdt_pr {
            val.write_element("w:sdtPr", writer)?;
        }
        if let Some(ref val) = self.sdt_end_pr {
            val.write_element("w:sdtEndPr", writer)?;
        }
        if let Some(ref val) = self.sdt_content {
            val.write_element("w:sdtContent", writer)?;
        }
        #[cfg(feature = "extra-children")]
        for child in &self.extra_children {
            child.write_to(writer).map_err(SerializeError::from)?;
        }
        Ok(())
    }

    fn is_empty_element(&self) -> bool {
        if self.sdt_pr.is_some() {
            return false;
        }
        if self.sdt_end_pr.is_some() {
            return false;
        }
        if self.sdt_content.is_some() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for CTSdtCell {
    fn write_children<W: Write>(&self, writer: &mut Writer<W>) -> Result<(), SerializeError> {
        if let Some(ref val) = self.sdt_pr {
            val.write_element("w:sdtPr", writer)?;
        }
        if let Some(ref val) = self.sdt_end_pr {
            val.write_element("w:sdtEndPr", writer)?;
        }
        if let Some(ref val) = self.sdt_content {
            val.write_element("w:sdtContent", writer)?;
        }
        #[cfg(feature = "extra-children")]
        for child in &self.extra_children {
            child.write_to(writer).map_err(SerializeError::from)?;
        }
        Ok(())
    }

    fn is_empty_element(&self) -> bool {
        if self.sdt_pr.is_some() {
            return false;
        }
        if self.sdt_end_pr.is_some() {
            return false;
        }
        if self.sdt_content.is_some() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for CTSdtRow {
    fn write_children<W: Write>(&self, writer: &mut Writer<W>) -> Result<(), SerializeError> {
        if let Some(ref val) = self.sdt_pr {
            val.write_element("w:sdtPr", writer)?;
        }
        if let Some(ref val) = self.sdt_end_pr {
            val.write_element("w:sdtEndPr", writer)?;
        }
        if let Some(ref val) = self.sdt_content {
            val.write_element("w:sdtContent", writer)?;
        }
        #[cfg(feature = "extra-children")]
        for child in &self.extra_children {
            child.write_to(writer).map_err(SerializeError::from)?;
        }
        Ok(())
    }

    fn is_empty_element(&self) -> bool {
        if self.sdt_pr.is_some() {
            return false;
        }
        if self.sdt_end_pr.is_some() {
            return false;
        }
        if self.sdt_content.is_some() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for CTAttr {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        if let Some(ref val) = self.uri {
            start.push_attribute(("w:uri", val.as_str()));
        }
        {
            let val = &self.name;
            start.push_attribute(("w:name", val.as_str()));
        }
        {
            let val = &self.value;
            start.push_attribute(("w:val", val.as_str()));
        }
        #[cfg(feature = "extra-attrs")]
        for (key, value) in &self.extra_attrs {
            start.push_attribute((key.as_str(), value.as_str()));
        }
        start
    }

    fn is_empty_element(&self) -> bool {
        true
    }
}

impl ToXml for CTCustomXmlRun {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        if let Some(ref val) = self.uri {
            start.push_attribute(("w:uri", val.as_str()));
        }
        {
            let val = &self.element;
            start.push_attribute(("w:element", val.as_str()));
        }
        #[cfg(feature = "extra-attrs")]
        for (key, value) in &self.extra_attrs {
            start.push_attribute((key.as_str(), value.as_str()));
        }
        start
    }

    fn write_children<W: Write>(&self, writer: &mut Writer<W>) -> Result<(), SerializeError> {
        if let Some(ref val) = self.custom_xml_pr {
            val.write_element("w:customXmlPr", writer)?;
        }
        for item in &self.p_content {
            item.write_element("", writer)?;
        }
        #[cfg(feature = "extra-children")]
        for child in &self.extra_children {
            child.write_to(writer).map_err(SerializeError::from)?;
        }
        Ok(())
    }

    fn is_empty_element(&self) -> bool {
        if self.custom_xml_pr.is_some() {
            return false;
        }
        if !self.p_content.is_empty() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for CTSmartTagRun {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        if let Some(ref val) = self.uri {
            start.push_attribute(("w:uri", val.as_str()));
        }
        {
            let val = &self.element;
            start.push_attribute(("w:element", val.as_str()));
        }
        #[cfg(feature = "extra-attrs")]
        for (key, value) in &self.extra_attrs {
            start.push_attribute((key.as_str(), value.as_str()));
        }
        start
    }

    fn write_children<W: Write>(&self, writer: &mut Writer<W>) -> Result<(), SerializeError> {
        if let Some(ref val) = self.smart_tag_pr {
            val.write_element("w:smartTagPr", writer)?;
        }
        for item in &self.p_content {
            item.write_element("", writer)?;
        }
        #[cfg(feature = "extra-children")]
        for child in &self.extra_children {
            child.write_to(writer).map_err(SerializeError::from)?;
        }
        Ok(())
    }

    fn is_empty_element(&self) -> bool {
        if self.smart_tag_pr.is_some() {
            return false;
        }
        if !self.p_content.is_empty() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for CTCustomXmlBlock {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        if let Some(ref val) = self.uri {
            start.push_attribute(("w:uri", val.as_str()));
        }
        {
            let val = &self.element;
            start.push_attribute(("w:element", val.as_str()));
        }
        #[cfg(feature = "extra-attrs")]
        for (key, value) in &self.extra_attrs {
            start.push_attribute((key.as_str(), value.as_str()));
        }
        start
    }

    fn write_children<W: Write>(&self, writer: &mut Writer<W>) -> Result<(), SerializeError> {
        if let Some(ref val) = self.custom_xml_pr {
            val.write_element("w:customXmlPr", writer)?;
        }
        for item in &self.content_block_content {
            item.write_element("", writer)?;
        }
        #[cfg(feature = "extra-children")]
        for child in &self.extra_children {
            child.write_to(writer).map_err(SerializeError::from)?;
        }
        Ok(())
    }

    fn is_empty_element(&self) -> bool {
        if self.custom_xml_pr.is_some() {
            return false;
        }
        if !self.content_block_content.is_empty() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for CTCustomXmlPr {
    fn write_children<W: Write>(&self, writer: &mut Writer<W>) -> Result<(), SerializeError> {
        if let Some(ref val) = self.placeholder {
            val.write_element("w:placeholder", writer)?;
        }
        for item in &self.attr {
            item.write_element("w:attr", writer)?;
        }
        #[cfg(feature = "extra-children")]
        for child in &self.extra_children {
            child.write_to(writer).map_err(SerializeError::from)?;
        }
        Ok(())
    }

    fn is_empty_element(&self) -> bool {
        if self.placeholder.is_some() {
            return false;
        }
        if !self.attr.is_empty() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for CTCustomXmlRow {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        if let Some(ref val) = self.uri {
            start.push_attribute(("w:uri", val.as_str()));
        }
        {
            let val = &self.element;
            start.push_attribute(("w:element", val.as_str()));
        }
        #[cfg(feature = "extra-attrs")]
        for (key, value) in &self.extra_attrs {
            start.push_attribute((key.as_str(), value.as_str()));
        }
        start
    }

    fn write_children<W: Write>(&self, writer: &mut Writer<W>) -> Result<(), SerializeError> {
        if let Some(ref val) = self.custom_xml_pr {
            val.write_element("w:customXmlPr", writer)?;
        }
        for item in &self.content_row_content {
            item.write_element("", writer)?;
        }
        #[cfg(feature = "extra-children")]
        for child in &self.extra_children {
            child.write_to(writer).map_err(SerializeError::from)?;
        }
        Ok(())
    }

    fn is_empty_element(&self) -> bool {
        if self.custom_xml_pr.is_some() {
            return false;
        }
        if !self.content_row_content.is_empty() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for CTCustomXmlCell {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        if let Some(ref val) = self.uri {
            start.push_attribute(("w:uri", val.as_str()));
        }
        {
            let val = &self.element;
            start.push_attribute(("w:element", val.as_str()));
        }
        #[cfg(feature = "extra-attrs")]
        for (key, value) in &self.extra_attrs {
            start.push_attribute((key.as_str(), value.as_str()));
        }
        start
    }

    fn write_children<W: Write>(&self, writer: &mut Writer<W>) -> Result<(), SerializeError> {
        if let Some(ref val) = self.custom_xml_pr {
            val.write_element("w:customXmlPr", writer)?;
        }
        for item in &self.content_cell_content {
            item.write_element("", writer)?;
        }
        #[cfg(feature = "extra-children")]
        for child in &self.extra_children {
            child.write_to(writer).map_err(SerializeError::from)?;
        }
        Ok(())
    }

    fn is_empty_element(&self) -> bool {
        if self.custom_xml_pr.is_some() {
            return false;
        }
        if !self.content_cell_content.is_empty() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for CTSmartTagPr {
    fn write_children<W: Write>(&self, writer: &mut Writer<W>) -> Result<(), SerializeError> {
        for item in &self.attr {
            item.write_element("w:attr", writer)?;
        }
        #[cfg(feature = "extra-children")]
        for child in &self.extra_children {
            child.write_to(writer).map_err(SerializeError::from)?;
        }
        Ok(())
    }

    fn is_empty_element(&self) -> bool {
        if !self.attr.is_empty() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for EGPContent {
    fn write_element<W: Write>(
        &self,
        _tag: &str,
        writer: &mut Writer<W>,
    ) -> Result<(), SerializeError> {
        match self {
            Self::CustomXml(inner) => inner.write_element("w:customXml", writer)?,
            Self::SmartTag(inner) => inner.write_element("w:smartTag", writer)?,
            Self::Sdt(inner) => inner.write_element("w:sdt", writer)?,
            Self::Dir(inner) => inner.write_element("w:dir", writer)?,
            Self::Bdo(inner) => inner.write_element("w:bdo", writer)?,
            Self::R(inner) => inner.write_element("w:r", writer)?,
            Self::ProofErr(inner) => inner.write_element("w:proofErr", writer)?,
            Self::PermStart(inner) => inner.write_element("w:permStart", writer)?,
            Self::PermEnd(inner) => inner.write_element("w:permEnd", writer)?,
            Self::BookmarkStart(inner) => inner.write_element("w:bookmarkStart", writer)?,
            Self::BookmarkEnd(inner) => inner.write_element("w:bookmarkEnd", writer)?,
            Self::MoveFromRangeStart(inner) => {
                inner.write_element("w:moveFromRangeStart", writer)?
            }
            Self::MoveFromRangeEnd(inner) => inner.write_element("w:moveFromRangeEnd", writer)?,
            Self::MoveToRangeStart(inner) => inner.write_element("w:moveToRangeStart", writer)?,
            Self::MoveToRangeEnd(inner) => inner.write_element("w:moveToRangeEnd", writer)?,
            Self::CommentRangeStart(inner) => inner.write_element("w:commentRangeStart", writer)?,
            Self::CommentRangeEnd(inner) => inner.write_element("w:commentRangeEnd", writer)?,
            Self::CustomXmlInsRangeStart(inner) => {
                inner.write_element("w:customXmlInsRangeStart", writer)?
            }
            Self::CustomXmlInsRangeEnd(inner) => {
                inner.write_element("w:customXmlInsRangeEnd", writer)?
            }
            Self::CustomXmlDelRangeStart(inner) => {
                inner.write_element("w:customXmlDelRangeStart", writer)?
            }
            Self::CustomXmlDelRangeEnd(inner) => {
                inner.write_element("w:customXmlDelRangeEnd", writer)?
            }
            Self::CustomXmlMoveFromRangeStart(inner) => {
                inner.write_element("w:customXmlMoveFromRangeStart", writer)?
            }
            Self::CustomXmlMoveFromRangeEnd(inner) => {
                inner.write_element("w:customXmlMoveFromRangeEnd", writer)?
            }
            Self::CustomXmlMoveToRangeStart(inner) => {
                inner.write_element("w:customXmlMoveToRangeStart", writer)?
            }
            Self::CustomXmlMoveToRangeEnd(inner) => {
                inner.write_element("w:customXmlMoveToRangeEnd", writer)?
            }
            Self::Ins(inner) => inner.write_element("w:ins", writer)?,
            Self::Del(inner) => inner.write_element("w:del", writer)?,
            Self::MoveFrom(inner) => inner.write_element("w:moveFrom", writer)?,
            Self::MoveTo(inner) => inner.write_element("w:moveTo", writer)?,
            Self::FldSimple(inner) => inner.write_element("w:fldSimple", writer)?,
            Self::Hyperlink(inner) => inner.write_element("w:hyperlink", writer)?,
            Self::SubDoc(inner) => inner.write_element("w:subDoc", writer)?,
        }
        Ok(())
    }
}

impl ToXml for Paragraph {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        if let Some(ref val) = self.rsid_r_pr {
            {
                let hex = encode_hex(val);
                start.push_attribute(("w:rsidRPr", hex.as_str()));
            }
        }
        if let Some(ref val) = self.rsid_r {
            {
                let hex = encode_hex(val);
                start.push_attribute(("w:rsidR", hex.as_str()));
            }
        }
        if let Some(ref val) = self.rsid_del {
            {
                let hex = encode_hex(val);
                start.push_attribute(("w:rsidDel", hex.as_str()));
            }
        }
        if let Some(ref val) = self.rsid_p {
            {
                let hex = encode_hex(val);
                start.push_attribute(("w:rsidP", hex.as_str()));
            }
        }
        if let Some(ref val) = self.rsid_r_default {
            {
                let hex = encode_hex(val);
                start.push_attribute(("w:rsidRDefault", hex.as_str()));
            }
        }
        #[cfg(feature = "extra-attrs")]
        for (key, value) in &self.extra_attrs {
            start.push_attribute((key.as_str(), value.as_str()));
        }
        start
    }

    fn write_children<W: Write>(&self, writer: &mut Writer<W>) -> Result<(), SerializeError> {
        #[cfg(feature = "wml-styling")]
        if let Some(ref val) = self.p_pr {
            val.write_element("w:pPr", writer)?;
        }
        for item in &self.p_content {
            item.write_element("", writer)?;
        }
        #[cfg(feature = "extra-children")]
        for child in &self.extra_children {
            child.write_to(writer).map_err(SerializeError::from)?;
        }
        Ok(())
    }

    fn is_empty_element(&self) -> bool {
        #[cfg(feature = "wml-styling")]
        if self.p_pr.is_some() {
            return false;
        }
        if !self.p_content.is_empty() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for CTHeight {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        if let Some(ref val) = self.value {
            {
                let s = val.to_string();
                start.push_attribute(("w:val", s.as_str()));
            }
        }
        if let Some(ref val) = self.h_rule {
            {
                let s = val.to_string();
                start.push_attribute(("w:hRule", s.as_str()));
            }
        }
        #[cfg(feature = "extra-attrs")]
        for (key, value) in &self.extra_attrs {
            start.push_attribute((key.as_str(), value.as_str()));
        }
        start
    }

    fn is_empty_element(&self) -> bool {
        true
    }
}

impl ToXml for CTTblWidth {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        if let Some(ref val) = self.width {
            {
                let s = val.to_string();
                start.push_attribute(("w:w", s.as_str()));
            }
        }
        if let Some(ref val) = self.r#type {
            {
                let s = val.to_string();
                start.push_attribute(("w:type", s.as_str()));
            }
        }
        #[cfg(feature = "extra-attrs")]
        for (key, value) in &self.extra_attrs {
            start.push_attribute((key.as_str(), value.as_str()));
        }
        start
    }

    fn is_empty_element(&self) -> bool {
        true
    }
}

impl ToXml for TableGridColumn {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        if let Some(ref val) = self.width {
            {
                let s = val.to_string();
                start.push_attribute(("w:w", s.as_str()));
            }
        }
        #[cfg(feature = "extra-attrs")]
        for (key, value) in &self.extra_attrs {
            start.push_attribute((key.as_str(), value.as_str()));
        }
        start
    }

    fn is_empty_element(&self) -> bool {
        true
    }
}

impl ToXml for CTTblGridBase {
    fn write_children<W: Write>(&self, writer: &mut Writer<W>) -> Result<(), SerializeError> {
        for item in &self.grid_col {
            item.write_element("w:gridCol", writer)?;
        }
        #[cfg(feature = "extra-children")]
        for child in &self.extra_children {
            child.write_to(writer).map_err(SerializeError::from)?;
        }
        Ok(())
    }

    fn is_empty_element(&self) -> bool {
        if !self.grid_col.is_empty() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for TableGrid {
    fn write_children<W: Write>(&self, writer: &mut Writer<W>) -> Result<(), SerializeError> {
        for item in &self.grid_col {
            item.write_element("w:gridCol", writer)?;
        }
        if let Some(ref val) = self.tbl_grid_change {
            val.write_element("w:tblGridChange", writer)?;
        }
        #[cfg(feature = "extra-children")]
        for child in &self.extra_children {
            child.write_to(writer).map_err(SerializeError::from)?;
        }
        Ok(())
    }

    fn is_empty_element(&self) -> bool {
        if !self.grid_col.is_empty() {
            return false;
        }
        if self.tbl_grid_change.is_some() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for CTTcBorders {
    fn write_children<W: Write>(&self, writer: &mut Writer<W>) -> Result<(), SerializeError> {
        if let Some(ref val) = self.top {
            val.write_element("w:top", writer)?;
        }
        if let Some(ref val) = self.start {
            val.write_element("w:start", writer)?;
        }
        if let Some(ref val) = self.left {
            val.write_element("w:left", writer)?;
        }
        if let Some(ref val) = self.bottom {
            val.write_element("w:bottom", writer)?;
        }
        if let Some(ref val) = self.end {
            val.write_element("w:end", writer)?;
        }
        if let Some(ref val) = self.right {
            val.write_element("w:right", writer)?;
        }
        if let Some(ref val) = self.inside_h {
            val.write_element("w:insideH", writer)?;
        }
        if let Some(ref val) = self.inside_v {
            val.write_element("w:insideV", writer)?;
        }
        if let Some(ref val) = self.tl2br {
            val.write_element("w:tl2br", writer)?;
        }
        if let Some(ref val) = self.tr2bl {
            val.write_element("w:tr2bl", writer)?;
        }
        #[cfg(feature = "extra-children")]
        for child in &self.extra_children {
            child.write_to(writer).map_err(SerializeError::from)?;
        }
        Ok(())
    }

    fn is_empty_element(&self) -> bool {
        if self.top.is_some() {
            return false;
        }
        if self.start.is_some() {
            return false;
        }
        if self.left.is_some() {
            return false;
        }
        if self.bottom.is_some() {
            return false;
        }
        if self.end.is_some() {
            return false;
        }
        if self.right.is_some() {
            return false;
        }
        if self.inside_h.is_some() {
            return false;
        }
        if self.inside_v.is_some() {
            return false;
        }
        if self.tl2br.is_some() {
            return false;
        }
        if self.tr2bl.is_some() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for CTTcMar {
    fn write_children<W: Write>(&self, writer: &mut Writer<W>) -> Result<(), SerializeError> {
        if let Some(ref val) = self.top {
            val.write_element("w:top", writer)?;
        }
        if let Some(ref val) = self.start {
            val.write_element("w:start", writer)?;
        }
        if let Some(ref val) = self.left {
            val.write_element("w:left", writer)?;
        }
        if let Some(ref val) = self.bottom {
            val.write_element("w:bottom", writer)?;
        }
        if let Some(ref val) = self.end {
            val.write_element("w:end", writer)?;
        }
        if let Some(ref val) = self.right {
            val.write_element("w:right", writer)?;
        }
        #[cfg(feature = "extra-children")]
        for child in &self.extra_children {
            child.write_to(writer).map_err(SerializeError::from)?;
        }
        Ok(())
    }

    fn is_empty_element(&self) -> bool {
        if self.top.is_some() {
            return false;
        }
        if self.start.is_some() {
            return false;
        }
        if self.left.is_some() {
            return false;
        }
        if self.bottom.is_some() {
            return false;
        }
        if self.end.is_some() {
            return false;
        }
        if self.right.is_some() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for CTVMerge {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        if let Some(ref val) = self.value {
            {
                let s = val.to_string();
                start.push_attribute(("w:val", s.as_str()));
            }
        }
        #[cfg(feature = "extra-attrs")]
        for (key, value) in &self.extra_attrs {
            start.push_attribute((key.as_str(), value.as_str()));
        }
        start
    }

    fn is_empty_element(&self) -> bool {
        true
    }
}

impl ToXml for CTHMerge {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        if let Some(ref val) = self.value {
            {
                let s = val.to_string();
                start.push_attribute(("w:val", s.as_str()));
            }
        }
        #[cfg(feature = "extra-attrs")]
        for (key, value) in &self.extra_attrs {
            start.push_attribute((key.as_str(), value.as_str()));
        }
        start
    }

    fn is_empty_element(&self) -> bool {
        true
    }
}

impl ToXml for CTTcPrBase {
    fn write_children<W: Write>(&self, writer: &mut Writer<W>) -> Result<(), SerializeError> {
        if let Some(ref val) = self.cnf_style {
            val.write_element("w:cnfStyle", writer)?;
        }
        if let Some(ref val) = self.tc_w {
            val.write_element("w:tcW", writer)?;
        }
        if let Some(ref val) = self.grid_span {
            val.write_element("w:gridSpan", writer)?;
        }
        if let Some(ref val) = self.horizontal_merge {
            val.write_element("w:hMerge", writer)?;
        }
        if let Some(ref val) = self.vertical_merge {
            val.write_element("w:vMerge", writer)?;
        }
        if let Some(ref val) = self.tc_borders {
            val.write_element("w:tcBorders", writer)?;
        }
        if let Some(ref val) = self.shading {
            val.write_element("w:shd", writer)?;
        }
        if let Some(ref val) = self.no_wrap {
            val.write_element("w:noWrap", writer)?;
        }
        if let Some(ref val) = self.tc_mar {
            val.write_element("w:tcMar", writer)?;
        }
        if let Some(ref val) = self.text_direction {
            val.write_element("w:textDirection", writer)?;
        }
        if let Some(ref val) = self.tc_fit_text {
            val.write_element("w:tcFitText", writer)?;
        }
        if let Some(ref val) = self.v_align {
            val.write_element("w:vAlign", writer)?;
        }
        if let Some(ref val) = self.hide_mark {
            val.write_element("w:hideMark", writer)?;
        }
        if let Some(ref val) = self.headers {
            val.write_element("w:headers", writer)?;
        }
        #[cfg(feature = "extra-children")]
        for child in &self.extra_children {
            child.write_to(writer).map_err(SerializeError::from)?;
        }
        Ok(())
    }

    fn is_empty_element(&self) -> bool {
        if self.cnf_style.is_some() {
            return false;
        }
        if self.tc_w.is_some() {
            return false;
        }
        if self.grid_span.is_some() {
            return false;
        }
        if self.horizontal_merge.is_some() {
            return false;
        }
        if self.vertical_merge.is_some() {
            return false;
        }
        if self.tc_borders.is_some() {
            return false;
        }
        if self.shading.is_some() {
            return false;
        }
        if self.no_wrap.is_some() {
            return false;
        }
        if self.tc_mar.is_some() {
            return false;
        }
        if self.text_direction.is_some() {
            return false;
        }
        if self.tc_fit_text.is_some() {
            return false;
        }
        if self.v_align.is_some() {
            return false;
        }
        if self.hide_mark.is_some() {
            return false;
        }
        if self.headers.is_some() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for TableCellProperties {
    fn write_children<W: Write>(&self, writer: &mut Writer<W>) -> Result<(), SerializeError> {
        if let Some(ref val) = self.cnf_style {
            val.write_element("w:cnfStyle", writer)?;
        }
        if let Some(ref val) = self.tc_w {
            val.write_element("w:tcW", writer)?;
        }
        if let Some(ref val) = self.grid_span {
            val.write_element("w:gridSpan", writer)?;
        }
        if let Some(ref val) = self.horizontal_merge {
            val.write_element("w:hMerge", writer)?;
        }
        if let Some(ref val) = self.vertical_merge {
            val.write_element("w:vMerge", writer)?;
        }
        if let Some(ref val) = self.tc_borders {
            val.write_element("w:tcBorders", writer)?;
        }
        if let Some(ref val) = self.shading {
            val.write_element("w:shd", writer)?;
        }
        if let Some(ref val) = self.no_wrap {
            val.write_element("w:noWrap", writer)?;
        }
        if let Some(ref val) = self.tc_mar {
            val.write_element("w:tcMar", writer)?;
        }
        if let Some(ref val) = self.text_direction {
            val.write_element("w:textDirection", writer)?;
        }
        if let Some(ref val) = self.tc_fit_text {
            val.write_element("w:tcFitText", writer)?;
        }
        if let Some(ref val) = self.v_align {
            val.write_element("w:vAlign", writer)?;
        }
        if let Some(ref val) = self.hide_mark {
            val.write_element("w:hideMark", writer)?;
        }
        if let Some(ref val) = self.headers {
            val.write_element("w:headers", writer)?;
        }
        if let Some(ref val) = self.cell_markup_elements {
            val.write_element("", writer)?;
        }
        if let Some(ref val) = self.tc_pr_change {
            val.write_element("w:tcPrChange", writer)?;
        }
        #[cfg(feature = "extra-children")]
        for child in &self.extra_children {
            child.write_to(writer).map_err(SerializeError::from)?;
        }
        Ok(())
    }

    fn is_empty_element(&self) -> bool {
        if self.cnf_style.is_some() {
            return false;
        }
        if self.tc_w.is_some() {
            return false;
        }
        if self.grid_span.is_some() {
            return false;
        }
        if self.horizontal_merge.is_some() {
            return false;
        }
        if self.vertical_merge.is_some() {
            return false;
        }
        if self.tc_borders.is_some() {
            return false;
        }
        if self.shading.is_some() {
            return false;
        }
        if self.no_wrap.is_some() {
            return false;
        }
        if self.tc_mar.is_some() {
            return false;
        }
        if self.text_direction.is_some() {
            return false;
        }
        if self.tc_fit_text.is_some() {
            return false;
        }
        if self.v_align.is_some() {
            return false;
        }
        if self.hide_mark.is_some() {
            return false;
        }
        if self.headers.is_some() {
            return false;
        }
        if self.cell_markup_elements.is_some() {
            return false;
        }
        if self.tc_pr_change.is_some() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for CTTcPrInner {
    fn write_children<W: Write>(&self, writer: &mut Writer<W>) -> Result<(), SerializeError> {
        if let Some(ref val) = self.cnf_style {
            val.write_element("w:cnfStyle", writer)?;
        }
        if let Some(ref val) = self.tc_w {
            val.write_element("w:tcW", writer)?;
        }
        if let Some(ref val) = self.grid_span {
            val.write_element("w:gridSpan", writer)?;
        }
        if let Some(ref val) = self.horizontal_merge {
            val.write_element("w:hMerge", writer)?;
        }
        if let Some(ref val) = self.vertical_merge {
            val.write_element("w:vMerge", writer)?;
        }
        if let Some(ref val) = self.tc_borders {
            val.write_element("w:tcBorders", writer)?;
        }
        if let Some(ref val) = self.shading {
            val.write_element("w:shd", writer)?;
        }
        if let Some(ref val) = self.no_wrap {
            val.write_element("w:noWrap", writer)?;
        }
        if let Some(ref val) = self.tc_mar {
            val.write_element("w:tcMar", writer)?;
        }
        if let Some(ref val) = self.text_direction {
            val.write_element("w:textDirection", writer)?;
        }
        if let Some(ref val) = self.tc_fit_text {
            val.write_element("w:tcFitText", writer)?;
        }
        if let Some(ref val) = self.v_align {
            val.write_element("w:vAlign", writer)?;
        }
        if let Some(ref val) = self.hide_mark {
            val.write_element("w:hideMark", writer)?;
        }
        if let Some(ref val) = self.headers {
            val.write_element("w:headers", writer)?;
        }
        if let Some(ref val) = self.cell_markup_elements {
            val.write_element("", writer)?;
        }
        #[cfg(feature = "extra-children")]
        for child in &self.extra_children {
            child.write_to(writer).map_err(SerializeError::from)?;
        }
        Ok(())
    }

    fn is_empty_element(&self) -> bool {
        if self.cnf_style.is_some() {
            return false;
        }
        if self.tc_w.is_some() {
            return false;
        }
        if self.grid_span.is_some() {
            return false;
        }
        if self.horizontal_merge.is_some() {
            return false;
        }
        if self.vertical_merge.is_some() {
            return false;
        }
        if self.tc_borders.is_some() {
            return false;
        }
        if self.shading.is_some() {
            return false;
        }
        if self.no_wrap.is_some() {
            return false;
        }
        if self.tc_mar.is_some() {
            return false;
        }
        if self.text_direction.is_some() {
            return false;
        }
        if self.tc_fit_text.is_some() {
            return false;
        }
        if self.v_align.is_some() {
            return false;
        }
        if self.hide_mark.is_some() {
            return false;
        }
        if self.headers.is_some() {
            return false;
        }
        if self.cell_markup_elements.is_some() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for TableCell {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        if let Some(ref val) = self.id {
            start.push_attribute(("w:id", val.as_str()));
        }
        #[cfg(feature = "extra-attrs")]
        for (key, value) in &self.extra_attrs {
            start.push_attribute((key.as_str(), value.as_str()));
        }
        start
    }

    fn write_children<W: Write>(&self, writer: &mut Writer<W>) -> Result<(), SerializeError> {
        if let Some(ref val) = self.cell_properties {
            val.write_element("w:tcPr", writer)?;
        }
        for item in &self.block_level_elts {
            item.write_element("", writer)?;
        }
        #[cfg(feature = "extra-children")]
        for child in &self.extra_children {
            child.write_to(writer).map_err(SerializeError::from)?;
        }
        Ok(())
    }

    fn is_empty_element(&self) -> bool {
        if self.cell_properties.is_some() {
            return false;
        }
        if !self.block_level_elts.is_empty() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for CTCnf {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        if let Some(ref val) = self.value {
            start.push_attribute(("w:val", val.as_str()));
        }
        if let Some(ref val) = self.first_row {
            {
                let s = val.to_string();
                start.push_attribute(("w:firstRow", s.as_str()));
            }
        }
        if let Some(ref val) = self.last_row {
            {
                let s = val.to_string();
                start.push_attribute(("w:lastRow", s.as_str()));
            }
        }
        if let Some(ref val) = self.first_column {
            {
                let s = val.to_string();
                start.push_attribute(("w:firstColumn", s.as_str()));
            }
        }
        if let Some(ref val) = self.last_column {
            {
                let s = val.to_string();
                start.push_attribute(("w:lastColumn", s.as_str()));
            }
        }
        if let Some(ref val) = self.odd_v_band {
            {
                let s = val.to_string();
                start.push_attribute(("w:oddVBand", s.as_str()));
            }
        }
        if let Some(ref val) = self.even_v_band {
            {
                let s = val.to_string();
                start.push_attribute(("w:evenVBand", s.as_str()));
            }
        }
        if let Some(ref val) = self.odd_h_band {
            {
                let s = val.to_string();
                start.push_attribute(("w:oddHBand", s.as_str()));
            }
        }
        if let Some(ref val) = self.even_h_band {
            {
                let s = val.to_string();
                start.push_attribute(("w:evenHBand", s.as_str()));
            }
        }
        if let Some(ref val) = self.first_row_first_column {
            {
                let s = val.to_string();
                start.push_attribute(("w:firstRowFirstColumn", s.as_str()));
            }
        }
        if let Some(ref val) = self.first_row_last_column {
            {
                let s = val.to_string();
                start.push_attribute(("w:firstRowLastColumn", s.as_str()));
            }
        }
        if let Some(ref val) = self.last_row_first_column {
            {
                let s = val.to_string();
                start.push_attribute(("w:lastRowFirstColumn", s.as_str()));
            }
        }
        if let Some(ref val) = self.last_row_last_column {
            {
                let s = val.to_string();
                start.push_attribute(("w:lastRowLastColumn", s.as_str()));
            }
        }
        #[cfg(feature = "extra-attrs")]
        for (key, value) in &self.extra_attrs {
            start.push_attribute((key.as_str(), value.as_str()));
        }
        start
    }

    fn is_empty_element(&self) -> bool {
        true
    }
}

impl ToXml for CTHeaders {
    fn write_children<W: Write>(&self, writer: &mut Writer<W>) -> Result<(), SerializeError> {
        for item in &self.header {
            item.write_element("w:header", writer)?;
        }
        #[cfg(feature = "extra-children")]
        for child in &self.extra_children {
            child.write_to(writer).map_err(SerializeError::from)?;
        }
        Ok(())
    }

    fn is_empty_element(&self) -> bool {
        if !self.header.is_empty() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for CTTrPrBase {
    fn write_children<W: Write>(&self, writer: &mut Writer<W>) -> Result<(), SerializeError> {
        #[cfg(feature = "extra-children")]
        for child in &self.extra_children {
            child.write_to(writer).map_err(SerializeError::from)?;
        }
        Ok(())
    }

    fn is_empty_element(&self) -> bool {
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for TableRowProperties {
    fn write_children<W: Write>(&self, writer: &mut Writer<W>) -> Result<(), SerializeError> {
        if let Some(ref val) = self.ins {
            val.write_element("w:ins", writer)?;
        }
        if let Some(ref val) = self.del {
            val.write_element("w:del", writer)?;
        }
        if let Some(ref val) = self.tr_pr_change {
            val.write_element("w:trPrChange", writer)?;
        }
        #[cfg(feature = "extra-children")]
        for child in &self.extra_children {
            child.write_to(writer).map_err(SerializeError::from)?;
        }
        Ok(())
    }

    fn is_empty_element(&self) -> bool {
        if self.ins.is_some() {
            return false;
        }
        if self.del.is_some() {
            return false;
        }
        if self.tr_pr_change.is_some() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for CTRow {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        if let Some(ref val) = self.rsid_r_pr {
            {
                let hex = encode_hex(val);
                start.push_attribute(("w:rsidRPr", hex.as_str()));
            }
        }
        if let Some(ref val) = self.rsid_r {
            {
                let hex = encode_hex(val);
                start.push_attribute(("w:rsidR", hex.as_str()));
            }
        }
        if let Some(ref val) = self.rsid_del {
            {
                let hex = encode_hex(val);
                start.push_attribute(("w:rsidDel", hex.as_str()));
            }
        }
        if let Some(ref val) = self.rsid_tr {
            {
                let hex = encode_hex(val);
                start.push_attribute(("w:rsidTr", hex.as_str()));
            }
        }
        #[cfg(feature = "extra-attrs")]
        for (key, value) in &self.extra_attrs {
            start.push_attribute((key.as_str(), value.as_str()));
        }
        start
    }

    fn write_children<W: Write>(&self, writer: &mut Writer<W>) -> Result<(), SerializeError> {
        if let Some(ref val) = self.tbl_pr_ex {
            val.write_element("w:tblPrEx", writer)?;
        }
        if let Some(ref val) = self.row_properties {
            val.write_element("w:trPr", writer)?;
        }
        for item in &self.content_cell_content {
            item.write_element("", writer)?;
        }
        #[cfg(feature = "extra-children")]
        for child in &self.extra_children {
            child.write_to(writer).map_err(SerializeError::from)?;
        }
        Ok(())
    }

    fn is_empty_element(&self) -> bool {
        if self.tbl_pr_ex.is_some() {
            return false;
        }
        if self.row_properties.is_some() {
            return false;
        }
        if !self.content_cell_content.is_empty() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for CTTblLayoutType {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        if let Some(ref val) = self.r#type {
            {
                let s = val.to_string();
                start.push_attribute(("w:type", s.as_str()));
            }
        }
        #[cfg(feature = "extra-attrs")]
        for (key, value) in &self.extra_attrs {
            start.push_attribute((key.as_str(), value.as_str()));
        }
        start
    }

    fn is_empty_element(&self) -> bool {
        true
    }
}

impl ToXml for CTTblOverlap {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        {
            let val = &self.value;
            {
                let s = val.to_string();
                start.push_attribute(("w:val", s.as_str()));
            }
        }
        #[cfg(feature = "extra-attrs")]
        for (key, value) in &self.extra_attrs {
            start.push_attribute((key.as_str(), value.as_str()));
        }
        start
    }

    fn is_empty_element(&self) -> bool {
        true
    }
}

impl ToXml for CTTblPPr {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        if let Some(ref val) = self.left_from_text {
            {
                let s = val.to_string();
                start.push_attribute(("w:leftFromText", s.as_str()));
            }
        }
        if let Some(ref val) = self.right_from_text {
            {
                let s = val.to_string();
                start.push_attribute(("w:rightFromText", s.as_str()));
            }
        }
        if let Some(ref val) = self.top_from_text {
            {
                let s = val.to_string();
                start.push_attribute(("w:topFromText", s.as_str()));
            }
        }
        if let Some(ref val) = self.bottom_from_text {
            {
                let s = val.to_string();
                start.push_attribute(("w:bottomFromText", s.as_str()));
            }
        }
        if let Some(ref val) = self.vert_anchor {
            {
                let s = val.to_string();
                start.push_attribute(("w:vertAnchor", s.as_str()));
            }
        }
        if let Some(ref val) = self.horz_anchor {
            {
                let s = val.to_string();
                start.push_attribute(("w:horzAnchor", s.as_str()));
            }
        }
        if let Some(ref val) = self.tblp_x_spec {
            {
                let s = val.to_string();
                start.push_attribute(("w:tblpXSpec", s.as_str()));
            }
        }
        if let Some(ref val) = self.tblp_x {
            {
                let s = val.to_string();
                start.push_attribute(("w:tblpX", s.as_str()));
            }
        }
        if let Some(ref val) = self.tblp_y_spec {
            {
                let s = val.to_string();
                start.push_attribute(("w:tblpYSpec", s.as_str()));
            }
        }
        if let Some(ref val) = self.tblp_y {
            {
                let s = val.to_string();
                start.push_attribute(("w:tblpY", s.as_str()));
            }
        }
        #[cfg(feature = "extra-attrs")]
        for (key, value) in &self.extra_attrs {
            start.push_attribute((key.as_str(), value.as_str()));
        }
        start
    }

    fn is_empty_element(&self) -> bool {
        true
    }
}

impl ToXml for CTTblCellMar {
    fn write_children<W: Write>(&self, writer: &mut Writer<W>) -> Result<(), SerializeError> {
        if let Some(ref val) = self.top {
            val.write_element("w:top", writer)?;
        }
        if let Some(ref val) = self.start {
            val.write_element("w:start", writer)?;
        }
        if let Some(ref val) = self.left {
            val.write_element("w:left", writer)?;
        }
        if let Some(ref val) = self.bottom {
            val.write_element("w:bottom", writer)?;
        }
        if let Some(ref val) = self.end {
            val.write_element("w:end", writer)?;
        }
        if let Some(ref val) = self.right {
            val.write_element("w:right", writer)?;
        }
        #[cfg(feature = "extra-children")]
        for child in &self.extra_children {
            child.write_to(writer).map_err(SerializeError::from)?;
        }
        Ok(())
    }

    fn is_empty_element(&self) -> bool {
        if self.top.is_some() {
            return false;
        }
        if self.start.is_some() {
            return false;
        }
        if self.left.is_some() {
            return false;
        }
        if self.bottom.is_some() {
            return false;
        }
        if self.end.is_some() {
            return false;
        }
        if self.right.is_some() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for CTTblBorders {
    fn write_children<W: Write>(&self, writer: &mut Writer<W>) -> Result<(), SerializeError> {
        if let Some(ref val) = self.top {
            val.write_element("w:top", writer)?;
        }
        if let Some(ref val) = self.start {
            val.write_element("w:start", writer)?;
        }
        if let Some(ref val) = self.left {
            val.write_element("w:left", writer)?;
        }
        if let Some(ref val) = self.bottom {
            val.write_element("w:bottom", writer)?;
        }
        if let Some(ref val) = self.end {
            val.write_element("w:end", writer)?;
        }
        if let Some(ref val) = self.right {
            val.write_element("w:right", writer)?;
        }
        if let Some(ref val) = self.inside_h {
            val.write_element("w:insideH", writer)?;
        }
        if let Some(ref val) = self.inside_v {
            val.write_element("w:insideV", writer)?;
        }
        #[cfg(feature = "extra-children")]
        for child in &self.extra_children {
            child.write_to(writer).map_err(SerializeError::from)?;
        }
        Ok(())
    }

    fn is_empty_element(&self) -> bool {
        if self.top.is_some() {
            return false;
        }
        if self.start.is_some() {
            return false;
        }
        if self.left.is_some() {
            return false;
        }
        if self.bottom.is_some() {
            return false;
        }
        if self.end.is_some() {
            return false;
        }
        if self.right.is_some() {
            return false;
        }
        if self.inside_h.is_some() {
            return false;
        }
        if self.inside_v.is_some() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for CTTblPrBase {
    fn write_children<W: Write>(&self, writer: &mut Writer<W>) -> Result<(), SerializeError> {
        if let Some(ref val) = self.tbl_style {
            val.write_element("w:tblStyle", writer)?;
        }
        if let Some(ref val) = self.tblp_pr {
            val.write_element("w:tblpPr", writer)?;
        }
        if let Some(ref val) = self.tbl_overlap {
            val.write_element("w:tblOverlap", writer)?;
        }
        if let Some(ref val) = self.bidi_visual {
            val.write_element("w:bidiVisual", writer)?;
        }
        if let Some(ref val) = self.tbl_style_row_band_size {
            val.write_element("w:tblStyleRowBandSize", writer)?;
        }
        if let Some(ref val) = self.tbl_style_col_band_size {
            val.write_element("w:tblStyleColBandSize", writer)?;
        }
        if let Some(ref val) = self.tbl_w {
            val.write_element("w:tblW", writer)?;
        }
        if let Some(ref val) = self.justification {
            val.write_element("w:jc", writer)?;
        }
        if let Some(ref val) = self.tbl_cell_spacing {
            val.write_element("w:tblCellSpacing", writer)?;
        }
        if let Some(ref val) = self.tbl_ind {
            val.write_element("w:tblInd", writer)?;
        }
        if let Some(ref val) = self.tbl_borders {
            val.write_element("w:tblBorders", writer)?;
        }
        if let Some(ref val) = self.shading {
            val.write_element("w:shd", writer)?;
        }
        if let Some(ref val) = self.tbl_layout {
            val.write_element("w:tblLayout", writer)?;
        }
        if let Some(ref val) = self.tbl_cell_mar {
            val.write_element("w:tblCellMar", writer)?;
        }
        if let Some(ref val) = self.tbl_look {
            val.write_element("w:tblLook", writer)?;
        }
        if let Some(ref val) = self.tbl_caption {
            val.write_element("w:tblCaption", writer)?;
        }
        if let Some(ref val) = self.tbl_description {
            val.write_element("w:tblDescription", writer)?;
        }
        #[cfg(feature = "extra-children")]
        for child in &self.extra_children {
            child.write_to(writer).map_err(SerializeError::from)?;
        }
        Ok(())
    }

    fn is_empty_element(&self) -> bool {
        if self.tbl_style.is_some() {
            return false;
        }
        if self.tblp_pr.is_some() {
            return false;
        }
        if self.tbl_overlap.is_some() {
            return false;
        }
        if self.bidi_visual.is_some() {
            return false;
        }
        if self.tbl_style_row_band_size.is_some() {
            return false;
        }
        if self.tbl_style_col_band_size.is_some() {
            return false;
        }
        if self.tbl_w.is_some() {
            return false;
        }
        if self.justification.is_some() {
            return false;
        }
        if self.tbl_cell_spacing.is_some() {
            return false;
        }
        if self.tbl_ind.is_some() {
            return false;
        }
        if self.tbl_borders.is_some() {
            return false;
        }
        if self.shading.is_some() {
            return false;
        }
        if self.tbl_layout.is_some() {
            return false;
        }
        if self.tbl_cell_mar.is_some() {
            return false;
        }
        if self.tbl_look.is_some() {
            return false;
        }
        if self.tbl_caption.is_some() {
            return false;
        }
        if self.tbl_description.is_some() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for TableProperties {
    fn write_children<W: Write>(&self, writer: &mut Writer<W>) -> Result<(), SerializeError> {
        if let Some(ref val) = self.tbl_style {
            val.write_element("w:tblStyle", writer)?;
        }
        if let Some(ref val) = self.tblp_pr {
            val.write_element("w:tblpPr", writer)?;
        }
        if let Some(ref val) = self.tbl_overlap {
            val.write_element("w:tblOverlap", writer)?;
        }
        if let Some(ref val) = self.bidi_visual {
            val.write_element("w:bidiVisual", writer)?;
        }
        if let Some(ref val) = self.tbl_style_row_band_size {
            val.write_element("w:tblStyleRowBandSize", writer)?;
        }
        if let Some(ref val) = self.tbl_style_col_band_size {
            val.write_element("w:tblStyleColBandSize", writer)?;
        }
        if let Some(ref val) = self.tbl_w {
            val.write_element("w:tblW", writer)?;
        }
        if let Some(ref val) = self.justification {
            val.write_element("w:jc", writer)?;
        }
        if let Some(ref val) = self.tbl_cell_spacing {
            val.write_element("w:tblCellSpacing", writer)?;
        }
        if let Some(ref val) = self.tbl_ind {
            val.write_element("w:tblInd", writer)?;
        }
        if let Some(ref val) = self.tbl_borders {
            val.write_element("w:tblBorders", writer)?;
        }
        if let Some(ref val) = self.shading {
            val.write_element("w:shd", writer)?;
        }
        if let Some(ref val) = self.tbl_layout {
            val.write_element("w:tblLayout", writer)?;
        }
        if let Some(ref val) = self.tbl_cell_mar {
            val.write_element("w:tblCellMar", writer)?;
        }
        if let Some(ref val) = self.tbl_look {
            val.write_element("w:tblLook", writer)?;
        }
        if let Some(ref val) = self.tbl_caption {
            val.write_element("w:tblCaption", writer)?;
        }
        if let Some(ref val) = self.tbl_description {
            val.write_element("w:tblDescription", writer)?;
        }
        if let Some(ref val) = self.tbl_pr_change {
            val.write_element("w:tblPrChange", writer)?;
        }
        #[cfg(feature = "extra-children")]
        for child in &self.extra_children {
            child.write_to(writer).map_err(SerializeError::from)?;
        }
        Ok(())
    }

    fn is_empty_element(&self) -> bool {
        if self.tbl_style.is_some() {
            return false;
        }
        if self.tblp_pr.is_some() {
            return false;
        }
        if self.tbl_overlap.is_some() {
            return false;
        }
        if self.bidi_visual.is_some() {
            return false;
        }
        if self.tbl_style_row_band_size.is_some() {
            return false;
        }
        if self.tbl_style_col_band_size.is_some() {
            return false;
        }
        if self.tbl_w.is_some() {
            return false;
        }
        if self.justification.is_some() {
            return false;
        }
        if self.tbl_cell_spacing.is_some() {
            return false;
        }
        if self.tbl_ind.is_some() {
            return false;
        }
        if self.tbl_borders.is_some() {
            return false;
        }
        if self.shading.is_some() {
            return false;
        }
        if self.tbl_layout.is_some() {
            return false;
        }
        if self.tbl_cell_mar.is_some() {
            return false;
        }
        if self.tbl_look.is_some() {
            return false;
        }
        if self.tbl_caption.is_some() {
            return false;
        }
        if self.tbl_description.is_some() {
            return false;
        }
        if self.tbl_pr_change.is_some() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for CTTblPrExBase {
    fn write_children<W: Write>(&self, writer: &mut Writer<W>) -> Result<(), SerializeError> {
        if let Some(ref val) = self.tbl_w {
            val.write_element("w:tblW", writer)?;
        }
        if let Some(ref val) = self.justification {
            val.write_element("w:jc", writer)?;
        }
        if let Some(ref val) = self.tbl_cell_spacing {
            val.write_element("w:tblCellSpacing", writer)?;
        }
        if let Some(ref val) = self.tbl_ind {
            val.write_element("w:tblInd", writer)?;
        }
        if let Some(ref val) = self.tbl_borders {
            val.write_element("w:tblBorders", writer)?;
        }
        if let Some(ref val) = self.shading {
            val.write_element("w:shd", writer)?;
        }
        if let Some(ref val) = self.tbl_layout {
            val.write_element("w:tblLayout", writer)?;
        }
        if let Some(ref val) = self.tbl_cell_mar {
            val.write_element("w:tblCellMar", writer)?;
        }
        if let Some(ref val) = self.tbl_look {
            val.write_element("w:tblLook", writer)?;
        }
        #[cfg(feature = "extra-children")]
        for child in &self.extra_children {
            child.write_to(writer).map_err(SerializeError::from)?;
        }
        Ok(())
    }

    fn is_empty_element(&self) -> bool {
        if self.tbl_w.is_some() {
            return false;
        }
        if self.justification.is_some() {
            return false;
        }
        if self.tbl_cell_spacing.is_some() {
            return false;
        }
        if self.tbl_ind.is_some() {
            return false;
        }
        if self.tbl_borders.is_some() {
            return false;
        }
        if self.shading.is_some() {
            return false;
        }
        if self.tbl_layout.is_some() {
            return false;
        }
        if self.tbl_cell_mar.is_some() {
            return false;
        }
        if self.tbl_look.is_some() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for CTTblPrEx {
    fn write_children<W: Write>(&self, writer: &mut Writer<W>) -> Result<(), SerializeError> {
        if let Some(ref val) = self.tbl_w {
            val.write_element("w:tblW", writer)?;
        }
        if let Some(ref val) = self.justification {
            val.write_element("w:jc", writer)?;
        }
        if let Some(ref val) = self.tbl_cell_spacing {
            val.write_element("w:tblCellSpacing", writer)?;
        }
        if let Some(ref val) = self.tbl_ind {
            val.write_element("w:tblInd", writer)?;
        }
        if let Some(ref val) = self.tbl_borders {
            val.write_element("w:tblBorders", writer)?;
        }
        if let Some(ref val) = self.shading {
            val.write_element("w:shd", writer)?;
        }
        if let Some(ref val) = self.tbl_layout {
            val.write_element("w:tblLayout", writer)?;
        }
        if let Some(ref val) = self.tbl_cell_mar {
            val.write_element("w:tblCellMar", writer)?;
        }
        if let Some(ref val) = self.tbl_look {
            val.write_element("w:tblLook", writer)?;
        }
        if let Some(ref val) = self.tbl_pr_ex_change {
            val.write_element("w:tblPrExChange", writer)?;
        }
        #[cfg(feature = "extra-children")]
        for child in &self.extra_children {
            child.write_to(writer).map_err(SerializeError::from)?;
        }
        Ok(())
    }

    fn is_empty_element(&self) -> bool {
        if self.tbl_w.is_some() {
            return false;
        }
        if self.justification.is_some() {
            return false;
        }
        if self.tbl_cell_spacing.is_some() {
            return false;
        }
        if self.tbl_ind.is_some() {
            return false;
        }
        if self.tbl_borders.is_some() {
            return false;
        }
        if self.shading.is_some() {
            return false;
        }
        if self.tbl_layout.is_some() {
            return false;
        }
        if self.tbl_cell_mar.is_some() {
            return false;
        }
        if self.tbl_look.is_some() {
            return false;
        }
        if self.tbl_pr_ex_change.is_some() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for Table {
    fn write_children<W: Write>(&self, writer: &mut Writer<W>) -> Result<(), SerializeError> {
        for item in &self.range_markup_elements {
            item.write_element("", writer)?;
        }
        {
            let val = &self.table_properties;
            val.write_element("w:tblPr", writer)?;
        }
        {
            let val = &self.tbl_grid;
            val.write_element("w:tblGrid", writer)?;
        }
        for item in &self.content_row_content {
            item.write_element("", writer)?;
        }
        #[cfg(feature = "extra-children")]
        for child in &self.extra_children {
            child.write_to(writer).map_err(SerializeError::from)?;
        }
        Ok(())
    }

    fn is_empty_element(&self) -> bool {
        if !self.range_markup_elements.is_empty() {
            return false;
        }
        false
    }
}

impl ToXml for CTTblLook {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        if let Some(ref val) = self.first_row {
            {
                let s = val.to_string();
                start.push_attribute(("w:firstRow", s.as_str()));
            }
        }
        if let Some(ref val) = self.last_row {
            {
                let s = val.to_string();
                start.push_attribute(("w:lastRow", s.as_str()));
            }
        }
        if let Some(ref val) = self.first_column {
            {
                let s = val.to_string();
                start.push_attribute(("w:firstColumn", s.as_str()));
            }
        }
        if let Some(ref val) = self.last_column {
            {
                let s = val.to_string();
                start.push_attribute(("w:lastColumn", s.as_str()));
            }
        }
        if let Some(ref val) = self.no_h_band {
            {
                let s = val.to_string();
                start.push_attribute(("w:noHBand", s.as_str()));
            }
        }
        if let Some(ref val) = self.no_v_band {
            {
                let s = val.to_string();
                start.push_attribute(("w:noVBand", s.as_str()));
            }
        }
        if let Some(ref val) = self.value {
            {
                let hex = encode_hex(val);
                start.push_attribute(("w:val", hex.as_str()));
            }
        }
        #[cfg(feature = "extra-attrs")]
        for (key, value) in &self.extra_attrs {
            start.push_attribute((key.as_str(), value.as_str()));
        }
        start
    }

    fn is_empty_element(&self) -> bool {
        true
    }
}

impl ToXml for CTFtnPos {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        {
            let val = &self.value;
            {
                let s = val.to_string();
                start.push_attribute(("w:val", s.as_str()));
            }
        }
        #[cfg(feature = "extra-attrs")]
        for (key, value) in &self.extra_attrs {
            start.push_attribute((key.as_str(), value.as_str()));
        }
        start
    }

    fn is_empty_element(&self) -> bool {
        true
    }
}

impl ToXml for CTEdnPos {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        {
            let val = &self.value;
            {
                let s = val.to_string();
                start.push_attribute(("w:val", s.as_str()));
            }
        }
        #[cfg(feature = "extra-attrs")]
        for (key, value) in &self.extra_attrs {
            start.push_attribute((key.as_str(), value.as_str()));
        }
        start
    }

    fn is_empty_element(&self) -> bool {
        true
    }
}

impl ToXml for CTNumFmt {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        {
            let val = &self.value;
            {
                let s = val.to_string();
                start.push_attribute(("w:val", s.as_str()));
            }
        }
        if let Some(ref val) = self.format {
            start.push_attribute(("w:format", val.as_str()));
        }
        #[cfg(feature = "extra-attrs")]
        for (key, value) in &self.extra_attrs {
            start.push_attribute((key.as_str(), value.as_str()));
        }
        start
    }

    fn is_empty_element(&self) -> bool {
        true
    }
}

impl ToXml for CTNumRestart {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        {
            let val = &self.value;
            {
                let s = val.to_string();
                start.push_attribute(("w:val", s.as_str()));
            }
        }
        #[cfg(feature = "extra-attrs")]
        for (key, value) in &self.extra_attrs {
            start.push_attribute((key.as_str(), value.as_str()));
        }
        start
    }

    fn is_empty_element(&self) -> bool {
        true
    }
}

impl ToXml for CTFtnEdnRef {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        if let Some(ref val) = self.custom_mark_follows {
            {
                let s = val.to_string();
                start.push_attribute(("w:customMarkFollows", s.as_str()));
            }
        }
        {
            let val = &self.id;
            {
                let s = val.to_string();
                start.push_attribute(("w:id", s.as_str()));
            }
        }
        #[cfg(feature = "extra-attrs")]
        for (key, value) in &self.extra_attrs {
            start.push_attribute((key.as_str(), value.as_str()));
        }
        start
    }

    fn is_empty_element(&self) -> bool {
        true
    }
}

impl ToXml for CTFtnEdnSepRef {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        {
            let val = &self.id;
            {
                let s = val.to_string();
                start.push_attribute(("w:id", s.as_str()));
            }
        }
        #[cfg(feature = "extra-attrs")]
        for (key, value) in &self.extra_attrs {
            start.push_attribute((key.as_str(), value.as_str()));
        }
        start
    }

    fn is_empty_element(&self) -> bool {
        true
    }
}

impl ToXml for CTFtnEdn {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        if let Some(ref val) = self.r#type {
            {
                let s = val.to_string();
                start.push_attribute(("w:type", s.as_str()));
            }
        }
        {
            let val = &self.id;
            {
                let s = val.to_string();
                start.push_attribute(("w:id", s.as_str()));
            }
        }
        #[cfg(feature = "extra-attrs")]
        for (key, value) in &self.extra_attrs {
            start.push_attribute((key.as_str(), value.as_str()));
        }
        start
    }

    fn write_children<W: Write>(&self, writer: &mut Writer<W>) -> Result<(), SerializeError> {
        for item in &self.block_level_elts {
            item.write_element("", writer)?;
        }
        #[cfg(feature = "extra-children")]
        for child in &self.extra_children {
            child.write_to(writer).map_err(SerializeError::from)?;
        }
        Ok(())
    }

    fn is_empty_element(&self) -> bool {
        if !self.block_level_elts.is_empty() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for EGFtnEdnNumProps {
    fn write_children<W: Write>(&self, writer: &mut Writer<W>) -> Result<(), SerializeError> {
        if let Some(ref val) = self.num_start {
            val.write_element("w:numStart", writer)?;
        }
        if let Some(ref val) = self.num_restart {
            val.write_element("w:numRestart", writer)?;
        }
        #[cfg(feature = "extra-children")]
        for child in &self.extra_children {
            child.write_to(writer).map_err(SerializeError::from)?;
        }
        Ok(())
    }

    fn is_empty_element(&self) -> bool {
        if self.num_start.is_some() {
            return false;
        }
        if self.num_restart.is_some() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for CTFtnProps {
    fn write_children<W: Write>(&self, writer: &mut Writer<W>) -> Result<(), SerializeError> {
        if let Some(ref val) = self.pos {
            val.write_element("w:pos", writer)?;
        }
        if let Some(ref val) = self.num_fmt {
            val.write_element("w:numFmt", writer)?;
        }
        if let Some(ref val) = self.num_start {
            val.write_element("w:numStart", writer)?;
        }
        if let Some(ref val) = self.num_restart {
            val.write_element("w:numRestart", writer)?;
        }
        #[cfg(feature = "extra-children")]
        for child in &self.extra_children {
            child.write_to(writer).map_err(SerializeError::from)?;
        }
        Ok(())
    }

    fn is_empty_element(&self) -> bool {
        if self.pos.is_some() {
            return false;
        }
        if self.num_fmt.is_some() {
            return false;
        }
        if self.num_start.is_some() {
            return false;
        }
        if self.num_restart.is_some() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for CTEdnProps {
    fn write_children<W: Write>(&self, writer: &mut Writer<W>) -> Result<(), SerializeError> {
        if let Some(ref val) = self.pos {
            val.write_element("w:pos", writer)?;
        }
        if let Some(ref val) = self.num_fmt {
            val.write_element("w:numFmt", writer)?;
        }
        if let Some(ref val) = self.num_start {
            val.write_element("w:numStart", writer)?;
        }
        if let Some(ref val) = self.num_restart {
            val.write_element("w:numRestart", writer)?;
        }
        #[cfg(feature = "extra-children")]
        for child in &self.extra_children {
            child.write_to(writer).map_err(SerializeError::from)?;
        }
        Ok(())
    }

    fn is_empty_element(&self) -> bool {
        if self.pos.is_some() {
            return false;
        }
        if self.num_fmt.is_some() {
            return false;
        }
        if self.num_start.is_some() {
            return false;
        }
        if self.num_restart.is_some() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for CTFtnDocProps {
    fn write_children<W: Write>(&self, writer: &mut Writer<W>) -> Result<(), SerializeError> {
        if let Some(ref val) = self.pos {
            val.write_element("w:pos", writer)?;
        }
        if let Some(ref val) = self.num_fmt {
            val.write_element("w:numFmt", writer)?;
        }
        if let Some(ref val) = self.num_start {
            val.write_element("w:numStart", writer)?;
        }
        if let Some(ref val) = self.num_restart {
            val.write_element("w:numRestart", writer)?;
        }
        for item in &self.footnote {
            item.write_element("w:footnote", writer)?;
        }
        #[cfg(feature = "extra-children")]
        for child in &self.extra_children {
            child.write_to(writer).map_err(SerializeError::from)?;
        }
        Ok(())
    }

    fn is_empty_element(&self) -> bool {
        if self.pos.is_some() {
            return false;
        }
        if self.num_fmt.is_some() {
            return false;
        }
        if self.num_start.is_some() {
            return false;
        }
        if self.num_restart.is_some() {
            return false;
        }
        if !self.footnote.is_empty() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for CTEdnDocProps {
    fn write_children<W: Write>(&self, writer: &mut Writer<W>) -> Result<(), SerializeError> {
        if let Some(ref val) = self.pos {
            val.write_element("w:pos", writer)?;
        }
        if let Some(ref val) = self.num_fmt {
            val.write_element("w:numFmt", writer)?;
        }
        if let Some(ref val) = self.num_start {
            val.write_element("w:numStart", writer)?;
        }
        if let Some(ref val) = self.num_restart {
            val.write_element("w:numRestart", writer)?;
        }
        for item in &self.endnote {
            item.write_element("w:endnote", writer)?;
        }
        #[cfg(feature = "extra-children")]
        for child in &self.extra_children {
            child.write_to(writer).map_err(SerializeError::from)?;
        }
        Ok(())
    }

    fn is_empty_element(&self) -> bool {
        if self.pos.is_some() {
            return false;
        }
        if self.num_fmt.is_some() {
            return false;
        }
        if self.num_start.is_some() {
            return false;
        }
        if self.num_restart.is_some() {
            return false;
        }
        if !self.endnote.is_empty() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for CTRecipientData {
    fn write_children<W: Write>(&self, writer: &mut Writer<W>) -> Result<(), SerializeError> {
        if let Some(ref val) = self.active {
            val.write_element("w:active", writer)?;
        }
        {
            let val = &self.column;
            val.write_element("w:column", writer)?;
        }
        {
            let val = &self.unique_tag;
            val.write_element("w:uniqueTag", writer)?;
        }
        #[cfg(feature = "extra-children")]
        for child in &self.extra_children {
            child.write_to(writer).map_err(SerializeError::from)?;
        }
        Ok(())
    }

    fn is_empty_element(&self) -> bool {
        if self.active.is_some() {
            return false;
        }
        false
    }
}

impl ToXml for CTBase64Binary {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        {
            let val = &self.value;
            {
                let hex = encode_hex(val);
                start.push_attribute(("w:val", hex.as_str()));
            }
        }
        #[cfg(feature = "extra-attrs")]
        for (key, value) in &self.extra_attrs {
            start.push_attribute((key.as_str(), value.as_str()));
        }
        start
    }

    fn is_empty_element(&self) -> bool {
        true
    }
}

impl ToXml for CTRecipients {
    fn write_children<W: Write>(&self, writer: &mut Writer<W>) -> Result<(), SerializeError> {
        for item in &self.recipient_data {
            item.write_element("w:recipientData", writer)?;
        }
        #[cfg(feature = "extra-children")]
        for child in &self.extra_children {
            child.write_to(writer).map_err(SerializeError::from)?;
        }
        Ok(())
    }

    fn is_empty_element(&self) -> bool {
        if !self.recipient_data.is_empty() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for CTOdsoFieldMapData {
    fn write_children<W: Write>(&self, writer: &mut Writer<W>) -> Result<(), SerializeError> {
        if let Some(ref val) = self.r#type {
            val.write_element("w:type", writer)?;
        }
        if let Some(ref val) = self.name {
            val.write_element("w:name", writer)?;
        }
        if let Some(ref val) = self.mapped_name {
            val.write_element("w:mappedName", writer)?;
        }
        if let Some(ref val) = self.column {
            val.write_element("w:column", writer)?;
        }
        if let Some(ref val) = self.lid {
            val.write_element("w:lid", writer)?;
        }
        if let Some(ref val) = self.dynamic_address {
            val.write_element("w:dynamicAddress", writer)?;
        }
        #[cfg(feature = "extra-children")]
        for child in &self.extra_children {
            child.write_to(writer).map_err(SerializeError::from)?;
        }
        Ok(())
    }

    fn is_empty_element(&self) -> bool {
        if self.r#type.is_some() {
            return false;
        }
        if self.name.is_some() {
            return false;
        }
        if self.mapped_name.is_some() {
            return false;
        }
        if self.column.is_some() {
            return false;
        }
        if self.lid.is_some() {
            return false;
        }
        if self.dynamic_address.is_some() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for CTMailMergeSourceType {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        {
            let val = &self.value;
            {
                let s = val.to_string();
                start.push_attribute(("w:val", s.as_str()));
            }
        }
        #[cfg(feature = "extra-attrs")]
        for (key, value) in &self.extra_attrs {
            start.push_attribute((key.as_str(), value.as_str()));
        }
        start
    }

    fn is_empty_element(&self) -> bool {
        true
    }
}

impl ToXml for CTOdso {
    fn write_children<W: Write>(&self, writer: &mut Writer<W>) -> Result<(), SerializeError> {
        if let Some(ref val) = self.udl {
            val.write_element("w:udl", writer)?;
        }
        if let Some(ref val) = self.table {
            val.write_element("w:table", writer)?;
        }
        if let Some(ref val) = self.src {
            val.write_element("w:src", writer)?;
        }
        if let Some(ref val) = self.col_delim {
            val.write_element("w:colDelim", writer)?;
        }
        if let Some(ref val) = self.r#type {
            val.write_element("w:type", writer)?;
        }
        if let Some(ref val) = self.f_hdr {
            val.write_element("w:fHdr", writer)?;
        }
        for item in &self.field_map_data {
            item.write_element("w:fieldMapData", writer)?;
        }
        for item in &self.recipient_data {
            item.write_element("w:recipientData", writer)?;
        }
        #[cfg(feature = "extra-children")]
        for child in &self.extra_children {
            child.write_to(writer).map_err(SerializeError::from)?;
        }
        Ok(())
    }

    fn is_empty_element(&self) -> bool {
        if self.udl.is_some() {
            return false;
        }
        if self.table.is_some() {
            return false;
        }
        if self.src.is_some() {
            return false;
        }
        if self.col_delim.is_some() {
            return false;
        }
        if self.r#type.is_some() {
            return false;
        }
        if self.f_hdr.is_some() {
            return false;
        }
        if !self.field_map_data.is_empty() {
            return false;
        }
        if !self.recipient_data.is_empty() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for CTMailMerge {
    fn write_children<W: Write>(&self, writer: &mut Writer<W>) -> Result<(), SerializeError> {
        {
            let val = &self.main_document_type;
            val.write_element("w:mainDocumentType", writer)?;
        }
        if let Some(ref val) = self.link_to_query {
            val.write_element("w:linkToQuery", writer)?;
        }
        {
            let val = &self.data_type;
            val.write_element("w:dataType", writer)?;
        }
        if let Some(ref val) = self.connect_string {
            val.write_element("w:connectString", writer)?;
        }
        if let Some(ref val) = self.query {
            val.write_element("w:query", writer)?;
        }
        if let Some(ref val) = self.data_source {
            val.write_element("w:dataSource", writer)?;
        }
        if let Some(ref val) = self.header_source {
            val.write_element("w:headerSource", writer)?;
        }
        if let Some(ref val) = self.do_not_suppress_blank_lines {
            val.write_element("w:doNotSuppressBlankLines", writer)?;
        }
        if let Some(ref val) = self.destination {
            val.write_element("w:destination", writer)?;
        }
        if let Some(ref val) = self.address_field_name {
            val.write_element("w:addressFieldName", writer)?;
        }
        if let Some(ref val) = self.mail_subject {
            val.write_element("w:mailSubject", writer)?;
        }
        if let Some(ref val) = self.mail_as_attachment {
            val.write_element("w:mailAsAttachment", writer)?;
        }
        if let Some(ref val) = self.view_merged_data {
            val.write_element("w:viewMergedData", writer)?;
        }
        if let Some(ref val) = self.active_record {
            val.write_element("w:activeRecord", writer)?;
        }
        if let Some(ref val) = self.check_errors {
            val.write_element("w:checkErrors", writer)?;
        }
        if let Some(ref val) = self.odso {
            val.write_element("w:odso", writer)?;
        }
        #[cfg(feature = "extra-children")]
        for child in &self.extra_children {
            child.write_to(writer).map_err(SerializeError::from)?;
        }
        Ok(())
    }

    fn is_empty_element(&self) -> bool {
        false
    }
}

impl ToXml for CTTargetScreenSz {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        {
            let val = &self.value;
            {
                let s = val.to_string();
                start.push_attribute(("w:val", s.as_str()));
            }
        }
        #[cfg(feature = "extra-attrs")]
        for (key, value) in &self.extra_attrs {
            start.push_attribute((key.as_str(), value.as_str()));
        }
        start
    }

    fn is_empty_element(&self) -> bool {
        true
    }
}

impl ToXml for Compatibility {
    fn write_children<W: Write>(&self, writer: &mut Writer<W>) -> Result<(), SerializeError> {
        if let Some(ref val) = self.use_single_borderfor_contiguous_cells {
            val.write_element("w:useSingleBorderforContiguousCells", writer)?;
        }
        if let Some(ref val) = self.wp_justification {
            val.write_element("w:wpJustification", writer)?;
        }
        if let Some(ref val) = self.no_tab_hang_ind {
            val.write_element("w:noTabHangInd", writer)?;
        }
        if let Some(ref val) = self.no_leading {
            val.write_element("w:noLeading", writer)?;
        }
        if let Some(ref val) = self.space_for_u_l {
            val.write_element("w:spaceForUL", writer)?;
        }
        if let Some(ref val) = self.no_column_balance {
            val.write_element("w:noColumnBalance", writer)?;
        }
        if let Some(ref val) = self.balance_single_byte_double_byte_width {
            val.write_element("w:balanceSingleByteDoubleByteWidth", writer)?;
        }
        if let Some(ref val) = self.no_extra_line_spacing {
            val.write_element("w:noExtraLineSpacing", writer)?;
        }
        if let Some(ref val) = self.do_not_leave_backslash_alone {
            val.write_element("w:doNotLeaveBackslashAlone", writer)?;
        }
        if let Some(ref val) = self.ul_trail_space {
            val.write_element("w:ulTrailSpace", writer)?;
        }
        if let Some(ref val) = self.do_not_expand_shift_return {
            val.write_element("w:doNotExpandShiftReturn", writer)?;
        }
        if let Some(ref val) = self.spacing_in_whole_points {
            val.write_element("w:spacingInWholePoints", writer)?;
        }
        if let Some(ref val) = self.line_wrap_like_word6 {
            val.write_element("w:lineWrapLikeWord6", writer)?;
        }
        if let Some(ref val) = self.print_body_text_before_header {
            val.write_element("w:printBodyTextBeforeHeader", writer)?;
        }
        if let Some(ref val) = self.print_col_black {
            val.write_element("w:printColBlack", writer)?;
        }
        if let Some(ref val) = self.wp_space_width {
            val.write_element("w:wpSpaceWidth", writer)?;
        }
        if let Some(ref val) = self.show_breaks_in_frames {
            val.write_element("w:showBreaksInFrames", writer)?;
        }
        if let Some(ref val) = self.sub_font_by_size {
            val.write_element("w:subFontBySize", writer)?;
        }
        if let Some(ref val) = self.suppress_bottom_spacing {
            val.write_element("w:suppressBottomSpacing", writer)?;
        }
        if let Some(ref val) = self.suppress_top_spacing {
            val.write_element("w:suppressTopSpacing", writer)?;
        }
        if let Some(ref val) = self.suppress_spacing_at_top_of_page {
            val.write_element("w:suppressSpacingAtTopOfPage", writer)?;
        }
        if let Some(ref val) = self.suppress_top_spacing_w_p {
            val.write_element("w:suppressTopSpacingWP", writer)?;
        }
        if let Some(ref val) = self.suppress_sp_bf_after_pg_brk {
            val.write_element("w:suppressSpBfAfterPgBrk", writer)?;
        }
        if let Some(ref val) = self.swap_borders_facing_pages {
            val.write_element("w:swapBordersFacingPages", writer)?;
        }
        if let Some(ref val) = self.conv_mail_merge_esc {
            val.write_element("w:convMailMergeEsc", writer)?;
        }
        if let Some(ref val) = self.truncate_font_heights_like_w_p6 {
            val.write_element("w:truncateFontHeightsLikeWP6", writer)?;
        }
        if let Some(ref val) = self.mw_small_caps {
            val.write_element("w:mwSmallCaps", writer)?;
        }
        if let Some(ref val) = self.use_printer_metrics {
            val.write_element("w:usePrinterMetrics", writer)?;
        }
        if let Some(ref val) = self.do_not_suppress_paragraph_borders {
            val.write_element("w:doNotSuppressParagraphBorders", writer)?;
        }
        if let Some(ref val) = self.wrap_trail_spaces {
            val.write_element("w:wrapTrailSpaces", writer)?;
        }
        if let Some(ref val) = self.footnote_layout_like_w_w8 {
            val.write_element("w:footnoteLayoutLikeWW8", writer)?;
        }
        if let Some(ref val) = self.shape_layout_like_w_w8 {
            val.write_element("w:shapeLayoutLikeWW8", writer)?;
        }
        if let Some(ref val) = self.align_tables_row_by_row {
            val.write_element("w:alignTablesRowByRow", writer)?;
        }
        if let Some(ref val) = self.forget_last_tab_alignment {
            val.write_element("w:forgetLastTabAlignment", writer)?;
        }
        if let Some(ref val) = self.adjust_line_height_in_table {
            val.write_element("w:adjustLineHeightInTable", writer)?;
        }
        if let Some(ref val) = self.auto_space_like_word95 {
            val.write_element("w:autoSpaceLikeWord95", writer)?;
        }
        if let Some(ref val) = self.no_space_raise_lower {
            val.write_element("w:noSpaceRaiseLower", writer)?;
        }
        if let Some(ref val) = self.do_not_use_h_t_m_l_paragraph_auto_spacing {
            val.write_element("w:doNotUseHTMLParagraphAutoSpacing", writer)?;
        }
        if let Some(ref val) = self.layout_raw_table_width {
            val.write_element("w:layoutRawTableWidth", writer)?;
        }
        if let Some(ref val) = self.layout_table_rows_apart {
            val.write_element("w:layoutTableRowsApart", writer)?;
        }
        if let Some(ref val) = self.use_word97_line_break_rules {
            val.write_element("w:useWord97LineBreakRules", writer)?;
        }
        if let Some(ref val) = self.do_not_break_wrapped_tables {
            val.write_element("w:doNotBreakWrappedTables", writer)?;
        }
        if let Some(ref val) = self.do_not_snap_to_grid_in_cell {
            val.write_element("w:doNotSnapToGridInCell", writer)?;
        }
        if let Some(ref val) = self.select_fld_with_first_or_last_char {
            val.write_element("w:selectFldWithFirstOrLastChar", writer)?;
        }
        if let Some(ref val) = self.apply_breaking_rules {
            val.write_element("w:applyBreakingRules", writer)?;
        }
        if let Some(ref val) = self.do_not_wrap_text_with_punct {
            val.write_element("w:doNotWrapTextWithPunct", writer)?;
        }
        if let Some(ref val) = self.do_not_use_east_asian_break_rules {
            val.write_element("w:doNotUseEastAsianBreakRules", writer)?;
        }
        if let Some(ref val) = self.use_word2002_table_style_rules {
            val.write_element("w:useWord2002TableStyleRules", writer)?;
        }
        if let Some(ref val) = self.grow_autofit {
            val.write_element("w:growAutofit", writer)?;
        }
        if let Some(ref val) = self.use_f_e_layout {
            val.write_element("w:useFELayout", writer)?;
        }
        if let Some(ref val) = self.use_normal_style_for_list {
            val.write_element("w:useNormalStyleForList", writer)?;
        }
        if let Some(ref val) = self.do_not_use_indent_as_numbering_tab_stop {
            val.write_element("w:doNotUseIndentAsNumberingTabStop", writer)?;
        }
        if let Some(ref val) = self.use_alt_kinsoku_line_break_rules {
            val.write_element("w:useAltKinsokuLineBreakRules", writer)?;
        }
        if let Some(ref val) = self.allow_space_of_same_style_in_table {
            val.write_element("w:allowSpaceOfSameStyleInTable", writer)?;
        }
        if let Some(ref val) = self.do_not_suppress_indentation {
            val.write_element("w:doNotSuppressIndentation", writer)?;
        }
        if let Some(ref val) = self.do_not_autofit_constrained_tables {
            val.write_element("w:doNotAutofitConstrainedTables", writer)?;
        }
        if let Some(ref val) = self.autofit_to_first_fixed_width_cell {
            val.write_element("w:autofitToFirstFixedWidthCell", writer)?;
        }
        if let Some(ref val) = self.underline_tab_in_num_list {
            val.write_element("w:underlineTabInNumList", writer)?;
        }
        if let Some(ref val) = self.display_hangul_fixed_width {
            val.write_element("w:displayHangulFixedWidth", writer)?;
        }
        if let Some(ref val) = self.split_pg_break_and_para_mark {
            val.write_element("w:splitPgBreakAndParaMark", writer)?;
        }
        if let Some(ref val) = self.do_not_vert_align_cell_with_sp {
            val.write_element("w:doNotVertAlignCellWithSp", writer)?;
        }
        if let Some(ref val) = self.do_not_break_constrained_forced_table {
            val.write_element("w:doNotBreakConstrainedForcedTable", writer)?;
        }
        if let Some(ref val) = self.do_not_vert_align_in_txbx {
            val.write_element("w:doNotVertAlignInTxbx", writer)?;
        }
        if let Some(ref val) = self.use_ansi_kerning_pairs {
            val.write_element("w:useAnsiKerningPairs", writer)?;
        }
        if let Some(ref val) = self.cached_col_balance {
            val.write_element("w:cachedColBalance", writer)?;
        }
        for item in &self.compat_setting {
            item.write_element("w:compatSetting", writer)?;
        }
        #[cfg(feature = "extra-children")]
        for child in &self.extra_children {
            child.write_to(writer).map_err(SerializeError::from)?;
        }
        Ok(())
    }

    fn is_empty_element(&self) -> bool {
        if self.use_single_borderfor_contiguous_cells.is_some() {
            return false;
        }
        if self.wp_justification.is_some() {
            return false;
        }
        if self.no_tab_hang_ind.is_some() {
            return false;
        }
        if self.no_leading.is_some() {
            return false;
        }
        if self.space_for_u_l.is_some() {
            return false;
        }
        if self.no_column_balance.is_some() {
            return false;
        }
        if self.balance_single_byte_double_byte_width.is_some() {
            return false;
        }
        if self.no_extra_line_spacing.is_some() {
            return false;
        }
        if self.do_not_leave_backslash_alone.is_some() {
            return false;
        }
        if self.ul_trail_space.is_some() {
            return false;
        }
        if self.do_not_expand_shift_return.is_some() {
            return false;
        }
        if self.spacing_in_whole_points.is_some() {
            return false;
        }
        if self.line_wrap_like_word6.is_some() {
            return false;
        }
        if self.print_body_text_before_header.is_some() {
            return false;
        }
        if self.print_col_black.is_some() {
            return false;
        }
        if self.wp_space_width.is_some() {
            return false;
        }
        if self.show_breaks_in_frames.is_some() {
            return false;
        }
        if self.sub_font_by_size.is_some() {
            return false;
        }
        if self.suppress_bottom_spacing.is_some() {
            return false;
        }
        if self.suppress_top_spacing.is_some() {
            return false;
        }
        if self.suppress_spacing_at_top_of_page.is_some() {
            return false;
        }
        if self.suppress_top_spacing_w_p.is_some() {
            return false;
        }
        if self.suppress_sp_bf_after_pg_brk.is_some() {
            return false;
        }
        if self.swap_borders_facing_pages.is_some() {
            return false;
        }
        if self.conv_mail_merge_esc.is_some() {
            return false;
        }
        if self.truncate_font_heights_like_w_p6.is_some() {
            return false;
        }
        if self.mw_small_caps.is_some() {
            return false;
        }
        if self.use_printer_metrics.is_some() {
            return false;
        }
        if self.do_not_suppress_paragraph_borders.is_some() {
            return false;
        }
        if self.wrap_trail_spaces.is_some() {
            return false;
        }
        if self.footnote_layout_like_w_w8.is_some() {
            return false;
        }
        if self.shape_layout_like_w_w8.is_some() {
            return false;
        }
        if self.align_tables_row_by_row.is_some() {
            return false;
        }
        if self.forget_last_tab_alignment.is_some() {
            return false;
        }
        if self.adjust_line_height_in_table.is_some() {
            return false;
        }
        if self.auto_space_like_word95.is_some() {
            return false;
        }
        if self.no_space_raise_lower.is_some() {
            return false;
        }
        if self.do_not_use_h_t_m_l_paragraph_auto_spacing.is_some() {
            return false;
        }
        if self.layout_raw_table_width.is_some() {
            return false;
        }
        if self.layout_table_rows_apart.is_some() {
            return false;
        }
        if self.use_word97_line_break_rules.is_some() {
            return false;
        }
        if self.do_not_break_wrapped_tables.is_some() {
            return false;
        }
        if self.do_not_snap_to_grid_in_cell.is_some() {
            return false;
        }
        if self.select_fld_with_first_or_last_char.is_some() {
            return false;
        }
        if self.apply_breaking_rules.is_some() {
            return false;
        }
        if self.do_not_wrap_text_with_punct.is_some() {
            return false;
        }
        if self.do_not_use_east_asian_break_rules.is_some() {
            return false;
        }
        if self.use_word2002_table_style_rules.is_some() {
            return false;
        }
        if self.grow_autofit.is_some() {
            return false;
        }
        if self.use_f_e_layout.is_some() {
            return false;
        }
        if self.use_normal_style_for_list.is_some() {
            return false;
        }
        if self.do_not_use_indent_as_numbering_tab_stop.is_some() {
            return false;
        }
        if self.use_alt_kinsoku_line_break_rules.is_some() {
            return false;
        }
        if self.allow_space_of_same_style_in_table.is_some() {
            return false;
        }
        if self.do_not_suppress_indentation.is_some() {
            return false;
        }
        if self.do_not_autofit_constrained_tables.is_some() {
            return false;
        }
        if self.autofit_to_first_fixed_width_cell.is_some() {
            return false;
        }
        if self.underline_tab_in_num_list.is_some() {
            return false;
        }
        if self.display_hangul_fixed_width.is_some() {
            return false;
        }
        if self.split_pg_break_and_para_mark.is_some() {
            return false;
        }
        if self.do_not_vert_align_cell_with_sp.is_some() {
            return false;
        }
        if self.do_not_break_constrained_forced_table.is_some() {
            return false;
        }
        if self.do_not_vert_align_in_txbx.is_some() {
            return false;
        }
        if self.use_ansi_kerning_pairs.is_some() {
            return false;
        }
        if self.cached_col_balance.is_some() {
            return false;
        }
        if !self.compat_setting.is_empty() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for CTCompatSetting {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        if let Some(ref val) = self.name {
            start.push_attribute(("w:name", val.as_str()));
        }
        if let Some(ref val) = self.uri {
            start.push_attribute(("w:uri", val.as_str()));
        }
        if let Some(ref val) = self.value {
            start.push_attribute(("w:val", val.as_str()));
        }
        #[cfg(feature = "extra-attrs")]
        for (key, value) in &self.extra_attrs {
            start.push_attribute((key.as_str(), value.as_str()));
        }
        start
    }

    fn is_empty_element(&self) -> bool {
        true
    }
}

impl ToXml for CTDocVar {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        {
            let val = &self.name;
            start.push_attribute(("w:name", val.as_str()));
        }
        {
            let val = &self.value;
            start.push_attribute(("w:val", val.as_str()));
        }
        #[cfg(feature = "extra-attrs")]
        for (key, value) in &self.extra_attrs {
            start.push_attribute((key.as_str(), value.as_str()));
        }
        start
    }

    fn is_empty_element(&self) -> bool {
        true
    }
}

impl ToXml for CTDocVars {
    fn write_children<W: Write>(&self, writer: &mut Writer<W>) -> Result<(), SerializeError> {
        for item in &self.doc_var {
            item.write_element("w:docVar", writer)?;
        }
        #[cfg(feature = "extra-children")]
        for child in &self.extra_children {
            child.write_to(writer).map_err(SerializeError::from)?;
        }
        Ok(())
    }

    fn is_empty_element(&self) -> bool {
        if !self.doc_var.is_empty() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for CTDocRsids {
    fn write_children<W: Write>(&self, writer: &mut Writer<W>) -> Result<(), SerializeError> {
        if let Some(ref val) = self.rsid_root {
            val.write_element("w:rsidRoot", writer)?;
        }
        for item in &self.rsid {
            item.write_element("w:rsid", writer)?;
        }
        #[cfg(feature = "extra-children")]
        for child in &self.extra_children {
            child.write_to(writer).map_err(SerializeError::from)?;
        }
        Ok(())
    }

    fn is_empty_element(&self) -> bool {
        if self.rsid_root.is_some() {
            return false;
        }
        if !self.rsid.is_empty() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for CTCharacterSpacing {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        {
            let val = &self.value;
            {
                let s = val.to_string();
                start.push_attribute(("w:val", s.as_str()));
            }
        }
        #[cfg(feature = "extra-attrs")]
        for (key, value) in &self.extra_attrs {
            start.push_attribute((key.as_str(), value.as_str()));
        }
        start
    }

    fn is_empty_element(&self) -> bool {
        true
    }
}

impl ToXml for CTSaveThroughXslt {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        if let Some(ref val) = self.solution_i_d {
            start.push_attribute(("w:solutionID", val.as_str()));
        }
        #[cfg(feature = "extra-attrs")]
        for (key, value) in &self.extra_attrs {
            start.push_attribute((key.as_str(), value.as_str()));
        }
        start
    }

    fn is_empty_element(&self) -> bool {
        true
    }
}

impl ToXml for CTRPrDefault {
    fn write_children<W: Write>(&self, writer: &mut Writer<W>) -> Result<(), SerializeError> {
        if let Some(ref val) = self.r_pr {
            val.write_element("w:rPr", writer)?;
        }
        #[cfg(feature = "extra-children")]
        for child in &self.extra_children {
            child.write_to(writer).map_err(SerializeError::from)?;
        }
        Ok(())
    }

    fn is_empty_element(&self) -> bool {
        if self.r_pr.is_some() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for CTPPrDefault {
    fn write_children<W: Write>(&self, writer: &mut Writer<W>) -> Result<(), SerializeError> {
        if let Some(ref val) = self.p_pr {
            val.write_element("w:pPr", writer)?;
        }
        #[cfg(feature = "extra-children")]
        for child in &self.extra_children {
            child.write_to(writer).map_err(SerializeError::from)?;
        }
        Ok(())
    }

    fn is_empty_element(&self) -> bool {
        if self.p_pr.is_some() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for CTDocDefaults {
    fn write_children<W: Write>(&self, writer: &mut Writer<W>) -> Result<(), SerializeError> {
        if let Some(ref val) = self.r_pr_default {
            val.write_element("w:rPrDefault", writer)?;
        }
        if let Some(ref val) = self.p_pr_default {
            val.write_element("w:pPrDefault", writer)?;
        }
        #[cfg(feature = "extra-children")]
        for child in &self.extra_children {
            child.write_to(writer).map_err(SerializeError::from)?;
        }
        Ok(())
    }

    fn is_empty_element(&self) -> bool {
        if self.r_pr_default.is_some() {
            return false;
        }
        if self.p_pr_default.is_some() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for CTColorSchemeMapping {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        if let Some(ref val) = self.bg1 {
            {
                let s = val.to_string();
                start.push_attribute(("w:bg1", s.as_str()));
            }
        }
        if let Some(ref val) = self.t1 {
            {
                let s = val.to_string();
                start.push_attribute(("w:t1", s.as_str()));
            }
        }
        if let Some(ref val) = self.bg2 {
            {
                let s = val.to_string();
                start.push_attribute(("w:bg2", s.as_str()));
            }
        }
        if let Some(ref val) = self.t2 {
            {
                let s = val.to_string();
                start.push_attribute(("w:t2", s.as_str()));
            }
        }
        if let Some(ref val) = self.accent1 {
            {
                let s = val.to_string();
                start.push_attribute(("w:accent1", s.as_str()));
            }
        }
        if let Some(ref val) = self.accent2 {
            {
                let s = val.to_string();
                start.push_attribute(("w:accent2", s.as_str()));
            }
        }
        if let Some(ref val) = self.accent3 {
            {
                let s = val.to_string();
                start.push_attribute(("w:accent3", s.as_str()));
            }
        }
        if let Some(ref val) = self.accent4 {
            {
                let s = val.to_string();
                start.push_attribute(("w:accent4", s.as_str()));
            }
        }
        if let Some(ref val) = self.accent5 {
            {
                let s = val.to_string();
                start.push_attribute(("w:accent5", s.as_str()));
            }
        }
        if let Some(ref val) = self.accent6 {
            {
                let s = val.to_string();
                start.push_attribute(("w:accent6", s.as_str()));
            }
        }
        if let Some(ref val) = self.hyperlink {
            {
                let s = val.to_string();
                start.push_attribute(("w:hyperlink", s.as_str()));
            }
        }
        if let Some(ref val) = self.followed_hyperlink {
            {
                let s = val.to_string();
                start.push_attribute(("w:followedHyperlink", s.as_str()));
            }
        }
        #[cfg(feature = "extra-attrs")]
        for (key, value) in &self.extra_attrs {
            start.push_attribute((key.as_str(), value.as_str()));
        }
        start
    }

    fn is_empty_element(&self) -> bool {
        true
    }
}

impl ToXml for CTReadingModeInkLockDown {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        {
            let val = &self.actual_pg;
            {
                let s = val.to_string();
                start.push_attribute(("w:actualPg", s.as_str()));
            }
        }
        {
            let val = &self.width;
            {
                let s = val.to_string();
                start.push_attribute(("w:w", s.as_str()));
            }
        }
        {
            let val = &self.height;
            {
                let s = val.to_string();
                start.push_attribute(("w:h", s.as_str()));
            }
        }
        {
            let val = &self.font_sz;
            {
                let s = val.to_string();
                start.push_attribute(("w:fontSz", s.as_str()));
            }
        }
        #[cfg(feature = "extra-attrs")]
        for (key, value) in &self.extra_attrs {
            start.push_attribute((key.as_str(), value.as_str()));
        }
        start
    }

    fn is_empty_element(&self) -> bool {
        true
    }
}

impl ToXml for CTWriteProtection {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        if let Some(ref val) = self.recommended {
            {
                let s = val.to_string();
                start.push_attribute(("w:recommended", s.as_str()));
            }
        }
        if let Some(ref val) = self.algorithm_name {
            start.push_attribute(("w:algorithmName", val.as_str()));
        }
        if let Some(ref val) = self.hash_value {
            {
                let hex = encode_hex(val);
                start.push_attribute(("w:hashValue", hex.as_str()));
            }
        }
        if let Some(ref val) = self.salt_value {
            {
                let hex = encode_hex(val);
                start.push_attribute(("w:saltValue", hex.as_str()));
            }
        }
        if let Some(ref val) = self.spin_count {
            {
                let s = val.to_string();
                start.push_attribute(("w:spinCount", s.as_str()));
            }
        }
        if let Some(ref val) = self.crypt_provider_type {
            {
                let s = val.to_string();
                start.push_attribute(("w:cryptProviderType", s.as_str()));
            }
        }
        if let Some(ref val) = self.crypt_algorithm_class {
            {
                let s = val.to_string();
                start.push_attribute(("w:cryptAlgorithmClass", s.as_str()));
            }
        }
        if let Some(ref val) = self.crypt_algorithm_type {
            {
                let s = val.to_string();
                start.push_attribute(("w:cryptAlgorithmType", s.as_str()));
            }
        }
        if let Some(ref val) = self.crypt_algorithm_sid {
            {
                let s = val.to_string();
                start.push_attribute(("w:cryptAlgorithmSid", s.as_str()));
            }
        }
        if let Some(ref val) = self.crypt_spin_count {
            {
                let s = val.to_string();
                start.push_attribute(("w:cryptSpinCount", s.as_str()));
            }
        }
        if let Some(ref val) = self.crypt_provider {
            start.push_attribute(("w:cryptProvider", val.as_str()));
        }
        if let Some(ref val) = self.alg_id_ext {
            {
                let hex = encode_hex(val);
                start.push_attribute(("w:algIdExt", hex.as_str()));
            }
        }
        if let Some(ref val) = self.alg_id_ext_source {
            start.push_attribute(("w:algIdExtSource", val.as_str()));
        }
        if let Some(ref val) = self.crypt_provider_type_ext {
            {
                let hex = encode_hex(val);
                start.push_attribute(("w:cryptProviderTypeExt", hex.as_str()));
            }
        }
        if let Some(ref val) = self.crypt_provider_type_ext_source {
            start.push_attribute(("w:cryptProviderTypeExtSource", val.as_str()));
        }
        if let Some(ref val) = self.hash {
            {
                let hex = encode_hex(val);
                start.push_attribute(("w:hash", hex.as_str()));
            }
        }
        if let Some(ref val) = self.salt {
            {
                let hex = encode_hex(val);
                start.push_attribute(("w:salt", hex.as_str()));
            }
        }
        #[cfg(feature = "extra-attrs")]
        for (key, value) in &self.extra_attrs {
            start.push_attribute((key.as_str(), value.as_str()));
        }
        start
    }

    fn is_empty_element(&self) -> bool {
        true
    }
}

impl ToXml for Settings {
    fn write_children<W: Write>(&self, writer: &mut Writer<W>) -> Result<(), SerializeError> {
        if let Some(ref val) = self.write_protection {
            val.write_element("w:writeProtection", writer)?;
        }
        if let Some(ref val) = self.view {
            val.write_element("w:view", writer)?;
        }
        if let Some(ref val) = self.zoom {
            val.write_element("w:zoom", writer)?;
        }
        if let Some(ref val) = self.remove_personal_information {
            val.write_element("w:removePersonalInformation", writer)?;
        }
        if let Some(ref val) = self.remove_date_and_time {
            val.write_element("w:removeDateAndTime", writer)?;
        }
        if let Some(ref val) = self.do_not_display_page_boundaries {
            val.write_element("w:doNotDisplayPageBoundaries", writer)?;
        }
        if let Some(ref val) = self.display_background_shape {
            val.write_element("w:displayBackgroundShape", writer)?;
        }
        if let Some(ref val) = self.print_post_script_over_text {
            val.write_element("w:printPostScriptOverText", writer)?;
        }
        if let Some(ref val) = self.print_fractional_character_width {
            val.write_element("w:printFractionalCharacterWidth", writer)?;
        }
        if let Some(ref val) = self.print_forms_data {
            val.write_element("w:printFormsData", writer)?;
        }
        if let Some(ref val) = self.embed_true_type_fonts {
            val.write_element("w:embedTrueTypeFonts", writer)?;
        }
        if let Some(ref val) = self.embed_system_fonts {
            val.write_element("w:embedSystemFonts", writer)?;
        }
        if let Some(ref val) = self.save_subset_fonts {
            val.write_element("w:saveSubsetFonts", writer)?;
        }
        if let Some(ref val) = self.save_forms_data {
            val.write_element("w:saveFormsData", writer)?;
        }
        if let Some(ref val) = self.mirror_margins {
            val.write_element("w:mirrorMargins", writer)?;
        }
        if let Some(ref val) = self.align_borders_and_edges {
            val.write_element("w:alignBordersAndEdges", writer)?;
        }
        if let Some(ref val) = self.borders_do_not_surround_header {
            val.write_element("w:bordersDoNotSurroundHeader", writer)?;
        }
        if let Some(ref val) = self.borders_do_not_surround_footer {
            val.write_element("w:bordersDoNotSurroundFooter", writer)?;
        }
        if let Some(ref val) = self.gutter_at_top {
            val.write_element("w:gutterAtTop", writer)?;
        }
        if let Some(ref val) = self.hide_spelling_errors {
            val.write_element("w:hideSpellingErrors", writer)?;
        }
        if let Some(ref val) = self.hide_grammatical_errors {
            val.write_element("w:hideGrammaticalErrors", writer)?;
        }
        for item in &self.active_writing_style {
            item.write_element("w:activeWritingStyle", writer)?;
        }
        if let Some(ref val) = self.proof_state {
            val.write_element("w:proofState", writer)?;
        }
        if let Some(ref val) = self.forms_design {
            val.write_element("w:formsDesign", writer)?;
        }
        if let Some(ref val) = self.attached_template {
            val.write_element("w:attachedTemplate", writer)?;
        }
        if let Some(ref val) = self.link_styles {
            val.write_element("w:linkStyles", writer)?;
        }
        if let Some(ref val) = self.style_pane_format_filter {
            val.write_element("w:stylePaneFormatFilter", writer)?;
        }
        if let Some(ref val) = self.style_pane_sort_method {
            val.write_element("w:stylePaneSortMethod", writer)?;
        }
        if let Some(ref val) = self.document_type {
            val.write_element("w:documentType", writer)?;
        }
        if let Some(ref val) = self.mail_merge {
            val.write_element("w:mailMerge", writer)?;
        }
        if let Some(ref val) = self.revision_view {
            val.write_element("w:revisionView", writer)?;
        }
        if let Some(ref val) = self.track_revisions {
            val.write_element("w:trackRevisions", writer)?;
        }
        if let Some(ref val) = self.do_not_track_moves {
            val.write_element("w:doNotTrackMoves", writer)?;
        }
        if let Some(ref val) = self.do_not_track_formatting {
            val.write_element("w:doNotTrackFormatting", writer)?;
        }
        if let Some(ref val) = self.document_protection {
            val.write_element("w:documentProtection", writer)?;
        }
        if let Some(ref val) = self.auto_format_override {
            val.write_element("w:autoFormatOverride", writer)?;
        }
        if let Some(ref val) = self.style_lock_theme {
            val.write_element("w:styleLockTheme", writer)?;
        }
        if let Some(ref val) = self.style_lock_q_f_set {
            val.write_element("w:styleLockQFSet", writer)?;
        }
        if let Some(ref val) = self.default_tab_stop {
            val.write_element("w:defaultTabStop", writer)?;
        }
        if let Some(ref val) = self.auto_hyphenation {
            val.write_element("w:autoHyphenation", writer)?;
        }
        if let Some(ref val) = self.consecutive_hyphen_limit {
            val.write_element("w:consecutiveHyphenLimit", writer)?;
        }
        if let Some(ref val) = self.hyphenation_zone {
            val.write_element("w:hyphenationZone", writer)?;
        }
        if let Some(ref val) = self.do_not_hyphenate_caps {
            val.write_element("w:doNotHyphenateCaps", writer)?;
        }
        if let Some(ref val) = self.show_envelope {
            val.write_element("w:showEnvelope", writer)?;
        }
        if let Some(ref val) = self.summary_length {
            val.write_element("w:summaryLength", writer)?;
        }
        if let Some(ref val) = self.click_and_type_style {
            val.write_element("w:clickAndTypeStyle", writer)?;
        }
        if let Some(ref val) = self.default_table_style {
            val.write_element("w:defaultTableStyle", writer)?;
        }
        if let Some(ref val) = self.even_and_odd_headers {
            val.write_element("w:evenAndOddHeaders", writer)?;
        }
        if let Some(ref val) = self.book_fold_rev_printing {
            val.write_element("w:bookFoldRevPrinting", writer)?;
        }
        if let Some(ref val) = self.book_fold_printing {
            val.write_element("w:bookFoldPrinting", writer)?;
        }
        if let Some(ref val) = self.book_fold_printing_sheets {
            val.write_element("w:bookFoldPrintingSheets", writer)?;
        }
        if let Some(ref val) = self.drawing_grid_horizontal_spacing {
            val.write_element("w:drawingGridHorizontalSpacing", writer)?;
        }
        if let Some(ref val) = self.drawing_grid_vertical_spacing {
            val.write_element("w:drawingGridVerticalSpacing", writer)?;
        }
        if let Some(ref val) = self.display_horizontal_drawing_grid_every {
            val.write_element("w:displayHorizontalDrawingGridEvery", writer)?;
        }
        if let Some(ref val) = self.display_vertical_drawing_grid_every {
            val.write_element("w:displayVerticalDrawingGridEvery", writer)?;
        }
        if let Some(ref val) = self.do_not_use_margins_for_drawing_grid_origin {
            val.write_element("w:doNotUseMarginsForDrawingGridOrigin", writer)?;
        }
        if let Some(ref val) = self.drawing_grid_horizontal_origin {
            val.write_element("w:drawingGridHorizontalOrigin", writer)?;
        }
        if let Some(ref val) = self.drawing_grid_vertical_origin {
            val.write_element("w:drawingGridVerticalOrigin", writer)?;
        }
        if let Some(ref val) = self.do_not_shade_form_data {
            val.write_element("w:doNotShadeFormData", writer)?;
        }
        if let Some(ref val) = self.no_punctuation_kerning {
            val.write_element("w:noPunctuationKerning", writer)?;
        }
        if let Some(ref val) = self.character_spacing_control {
            val.write_element("w:characterSpacingControl", writer)?;
        }
        if let Some(ref val) = self.print_two_on_one {
            val.write_element("w:printTwoOnOne", writer)?;
        }
        if let Some(ref val) = self.strict_first_and_last_chars {
            val.write_element("w:strictFirstAndLastChars", writer)?;
        }
        if let Some(ref val) = self.no_line_breaks_after {
            val.write_element("w:noLineBreaksAfter", writer)?;
        }
        if let Some(ref val) = self.no_line_breaks_before {
            val.write_element("w:noLineBreaksBefore", writer)?;
        }
        if let Some(ref val) = self.save_preview_picture {
            val.write_element("w:savePreviewPicture", writer)?;
        }
        if let Some(ref val) = self.do_not_validate_against_schema {
            val.write_element("w:doNotValidateAgainstSchema", writer)?;
        }
        if let Some(ref val) = self.save_invalid_xml {
            val.write_element("w:saveInvalidXml", writer)?;
        }
        if let Some(ref val) = self.ignore_mixed_content {
            val.write_element("w:ignoreMixedContent", writer)?;
        }
        if let Some(ref val) = self.always_show_placeholder_text {
            val.write_element("w:alwaysShowPlaceholderText", writer)?;
        }
        if let Some(ref val) = self.do_not_demarcate_invalid_xml {
            val.write_element("w:doNotDemarcateInvalidXml", writer)?;
        }
        if let Some(ref val) = self.save_xml_data_only {
            val.write_element("w:saveXmlDataOnly", writer)?;
        }
        if let Some(ref val) = self.use_x_s_l_t_when_saving {
            val.write_element("w:useXSLTWhenSaving", writer)?;
        }
        if let Some(ref val) = self.save_through_xslt {
            val.write_element("w:saveThroughXslt", writer)?;
        }
        if let Some(ref val) = self.show_x_m_l_tags {
            val.write_element("w:showXMLTags", writer)?;
        }
        if let Some(ref val) = self.always_merge_empty_namespace {
            val.write_element("w:alwaysMergeEmptyNamespace", writer)?;
        }
        if let Some(ref val) = self.update_fields {
            val.write_element("w:updateFields", writer)?;
        }
        if let Some(ref val) = self.hdr_shape_defaults {
            val.write_element("w:hdrShapeDefaults", writer)?;
        }
        if let Some(ref val) = self.footnote_pr {
            val.write_element("w:footnotePr", writer)?;
        }
        if let Some(ref val) = self.endnote_pr {
            val.write_element("w:endnotePr", writer)?;
        }
        if let Some(ref val) = self.compat {
            val.write_element("w:compat", writer)?;
        }
        if let Some(ref val) = self.doc_vars {
            val.write_element("w:docVars", writer)?;
        }
        if let Some(ref val) = self.rsids {
            val.write_element("w:rsids", writer)?;
        }
        for item in &self.attached_schema {
            item.write_element("w:attachedSchema", writer)?;
        }
        if let Some(ref val) = self.theme_font_lang {
            val.write_element("w:themeFontLang", writer)?;
        }
        if let Some(ref val) = self.clr_scheme_mapping {
            val.write_element("w:clrSchemeMapping", writer)?;
        }
        if let Some(ref val) = self.do_not_include_subdocs_in_stats {
            val.write_element("w:doNotIncludeSubdocsInStats", writer)?;
        }
        if let Some(ref val) = self.do_not_auto_compress_pictures {
            val.write_element("w:doNotAutoCompressPictures", writer)?;
        }
        if let Some(ref val) = self.force_upgrade {
            val.write_element("w:forceUpgrade", writer)?;
        }
        if let Some(ref val) = self.captions {
            val.write_element("w:captions", writer)?;
        }
        if let Some(ref val) = self.read_mode_ink_lock_down {
            val.write_element("w:readModeInkLockDown", writer)?;
        }
        for item in &self.smart_tag_type {
            item.write_element("w:smartTagType", writer)?;
        }
        if let Some(ref val) = self.shape_defaults {
            val.write_element("w:shapeDefaults", writer)?;
        }
        if let Some(ref val) = self.do_not_embed_smart_tags {
            val.write_element("w:doNotEmbedSmartTags", writer)?;
        }
        if let Some(ref val) = self.decimal_symbol {
            val.write_element("w:decimalSymbol", writer)?;
        }
        if let Some(ref val) = self.list_separator {
            val.write_element("w:listSeparator", writer)?;
        }
        #[cfg(feature = "extra-children")]
        for child in &self.extra_children {
            child.write_to(writer).map_err(SerializeError::from)?;
        }
        Ok(())
    }

    fn is_empty_element(&self) -> bool {
        if self.write_protection.is_some() {
            return false;
        }
        if self.view.is_some() {
            return false;
        }
        if self.zoom.is_some() {
            return false;
        }
        if self.remove_personal_information.is_some() {
            return false;
        }
        if self.remove_date_and_time.is_some() {
            return false;
        }
        if self.do_not_display_page_boundaries.is_some() {
            return false;
        }
        if self.display_background_shape.is_some() {
            return false;
        }
        if self.print_post_script_over_text.is_some() {
            return false;
        }
        if self.print_fractional_character_width.is_some() {
            return false;
        }
        if self.print_forms_data.is_some() {
            return false;
        }
        if self.embed_true_type_fonts.is_some() {
            return false;
        }
        if self.embed_system_fonts.is_some() {
            return false;
        }
        if self.save_subset_fonts.is_some() {
            return false;
        }
        if self.save_forms_data.is_some() {
            return false;
        }
        if self.mirror_margins.is_some() {
            return false;
        }
        if self.align_borders_and_edges.is_some() {
            return false;
        }
        if self.borders_do_not_surround_header.is_some() {
            return false;
        }
        if self.borders_do_not_surround_footer.is_some() {
            return false;
        }
        if self.gutter_at_top.is_some() {
            return false;
        }
        if self.hide_spelling_errors.is_some() {
            return false;
        }
        if self.hide_grammatical_errors.is_some() {
            return false;
        }
        if !self.active_writing_style.is_empty() {
            return false;
        }
        if self.proof_state.is_some() {
            return false;
        }
        if self.forms_design.is_some() {
            return false;
        }
        if self.attached_template.is_some() {
            return false;
        }
        if self.link_styles.is_some() {
            return false;
        }
        if self.style_pane_format_filter.is_some() {
            return false;
        }
        if self.style_pane_sort_method.is_some() {
            return false;
        }
        if self.document_type.is_some() {
            return false;
        }
        if self.mail_merge.is_some() {
            return false;
        }
        if self.revision_view.is_some() {
            return false;
        }
        if self.track_revisions.is_some() {
            return false;
        }
        if self.do_not_track_moves.is_some() {
            return false;
        }
        if self.do_not_track_formatting.is_some() {
            return false;
        }
        if self.document_protection.is_some() {
            return false;
        }
        if self.auto_format_override.is_some() {
            return false;
        }
        if self.style_lock_theme.is_some() {
            return false;
        }
        if self.style_lock_q_f_set.is_some() {
            return false;
        }
        if self.default_tab_stop.is_some() {
            return false;
        }
        if self.auto_hyphenation.is_some() {
            return false;
        }
        if self.consecutive_hyphen_limit.is_some() {
            return false;
        }
        if self.hyphenation_zone.is_some() {
            return false;
        }
        if self.do_not_hyphenate_caps.is_some() {
            return false;
        }
        if self.show_envelope.is_some() {
            return false;
        }
        if self.summary_length.is_some() {
            return false;
        }
        if self.click_and_type_style.is_some() {
            return false;
        }
        if self.default_table_style.is_some() {
            return false;
        }
        if self.even_and_odd_headers.is_some() {
            return false;
        }
        if self.book_fold_rev_printing.is_some() {
            return false;
        }
        if self.book_fold_printing.is_some() {
            return false;
        }
        if self.book_fold_printing_sheets.is_some() {
            return false;
        }
        if self.drawing_grid_horizontal_spacing.is_some() {
            return false;
        }
        if self.drawing_grid_vertical_spacing.is_some() {
            return false;
        }
        if self.display_horizontal_drawing_grid_every.is_some() {
            return false;
        }
        if self.display_vertical_drawing_grid_every.is_some() {
            return false;
        }
        if self.do_not_use_margins_for_drawing_grid_origin.is_some() {
            return false;
        }
        if self.drawing_grid_horizontal_origin.is_some() {
            return false;
        }
        if self.drawing_grid_vertical_origin.is_some() {
            return false;
        }
        if self.do_not_shade_form_data.is_some() {
            return false;
        }
        if self.no_punctuation_kerning.is_some() {
            return false;
        }
        if self.character_spacing_control.is_some() {
            return false;
        }
        if self.print_two_on_one.is_some() {
            return false;
        }
        if self.strict_first_and_last_chars.is_some() {
            return false;
        }
        if self.no_line_breaks_after.is_some() {
            return false;
        }
        if self.no_line_breaks_before.is_some() {
            return false;
        }
        if self.save_preview_picture.is_some() {
            return false;
        }
        if self.do_not_validate_against_schema.is_some() {
            return false;
        }
        if self.save_invalid_xml.is_some() {
            return false;
        }
        if self.ignore_mixed_content.is_some() {
            return false;
        }
        if self.always_show_placeholder_text.is_some() {
            return false;
        }
        if self.do_not_demarcate_invalid_xml.is_some() {
            return false;
        }
        if self.save_xml_data_only.is_some() {
            return false;
        }
        if self.use_x_s_l_t_when_saving.is_some() {
            return false;
        }
        if self.save_through_xslt.is_some() {
            return false;
        }
        if self.show_x_m_l_tags.is_some() {
            return false;
        }
        if self.always_merge_empty_namespace.is_some() {
            return false;
        }
        if self.update_fields.is_some() {
            return false;
        }
        if self.hdr_shape_defaults.is_some() {
            return false;
        }
        if self.footnote_pr.is_some() {
            return false;
        }
        if self.endnote_pr.is_some() {
            return false;
        }
        if self.compat.is_some() {
            return false;
        }
        if self.doc_vars.is_some() {
            return false;
        }
        if self.rsids.is_some() {
            return false;
        }
        if !self.attached_schema.is_empty() {
            return false;
        }
        if self.theme_font_lang.is_some() {
            return false;
        }
        if self.clr_scheme_mapping.is_some() {
            return false;
        }
        if self.do_not_include_subdocs_in_stats.is_some() {
            return false;
        }
        if self.do_not_auto_compress_pictures.is_some() {
            return false;
        }
        if self.force_upgrade.is_some() {
            return false;
        }
        if self.captions.is_some() {
            return false;
        }
        if self.read_mode_ink_lock_down.is_some() {
            return false;
        }
        if !self.smart_tag_type.is_empty() {
            return false;
        }
        if self.shape_defaults.is_some() {
            return false;
        }
        if self.do_not_embed_smart_tags.is_some() {
            return false;
        }
        if self.decimal_symbol.is_some() {
            return false;
        }
        if self.list_separator.is_some() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for CTStyleSort {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        {
            let val = &self.value;
            {
                let s = val.to_string();
                start.push_attribute(("w:val", s.as_str()));
            }
        }
        #[cfg(feature = "extra-attrs")]
        for (key, value) in &self.extra_attrs {
            start.push_attribute((key.as_str(), value.as_str()));
        }
        start
    }

    fn is_empty_element(&self) -> bool {
        true
    }
}

impl ToXml for CTStylePaneFilter {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        if let Some(ref val) = self.all_styles {
            {
                let s = val.to_string();
                start.push_attribute(("w:allStyles", s.as_str()));
            }
        }
        if let Some(ref val) = self.custom_styles {
            {
                let s = val.to_string();
                start.push_attribute(("w:customStyles", s.as_str()));
            }
        }
        if let Some(ref val) = self.latent_styles {
            {
                let s = val.to_string();
                start.push_attribute(("w:latentStyles", s.as_str()));
            }
        }
        if let Some(ref val) = self.styles_in_use {
            {
                let s = val.to_string();
                start.push_attribute(("w:stylesInUse", s.as_str()));
            }
        }
        if let Some(ref val) = self.heading_styles {
            {
                let s = val.to_string();
                start.push_attribute(("w:headingStyles", s.as_str()));
            }
        }
        if let Some(ref val) = self.numbering_styles {
            {
                let s = val.to_string();
                start.push_attribute(("w:numberingStyles", s.as_str()));
            }
        }
        if let Some(ref val) = self.table_styles {
            {
                let s = val.to_string();
                start.push_attribute(("w:tableStyles", s.as_str()));
            }
        }
        if let Some(ref val) = self.direct_formatting_on_runs {
            {
                let s = val.to_string();
                start.push_attribute(("w:directFormattingOnRuns", s.as_str()));
            }
        }
        if let Some(ref val) = self.direct_formatting_on_paragraphs {
            {
                let s = val.to_string();
                start.push_attribute(("w:directFormattingOnParagraphs", s.as_str()));
            }
        }
        if let Some(ref val) = self.direct_formatting_on_numbering {
            {
                let s = val.to_string();
                start.push_attribute(("w:directFormattingOnNumbering", s.as_str()));
            }
        }
        if let Some(ref val) = self.direct_formatting_on_tables {
            {
                let s = val.to_string();
                start.push_attribute(("w:directFormattingOnTables", s.as_str()));
            }
        }
        if let Some(ref val) = self.clear_formatting {
            {
                let s = val.to_string();
                start.push_attribute(("w:clearFormatting", s.as_str()));
            }
        }
        if let Some(ref val) = self.top3_heading_styles {
            {
                let s = val.to_string();
                start.push_attribute(("w:top3HeadingStyles", s.as_str()));
            }
        }
        if let Some(ref val) = self.visible_styles {
            {
                let s = val.to_string();
                start.push_attribute(("w:visibleStyles", s.as_str()));
            }
        }
        if let Some(ref val) = self.alternate_style_names {
            {
                let s = val.to_string();
                start.push_attribute(("w:alternateStyleNames", s.as_str()));
            }
        }
        if let Some(ref val) = self.value {
            {
                let hex = encode_hex(val);
                start.push_attribute(("w:val", hex.as_str()));
            }
        }
        #[cfg(feature = "extra-attrs")]
        for (key, value) in &self.extra_attrs {
            start.push_attribute((key.as_str(), value.as_str()));
        }
        start
    }

    fn is_empty_element(&self) -> bool {
        true
    }
}

impl ToXml for CTWebSettings {
    fn write_children<W: Write>(&self, writer: &mut Writer<W>) -> Result<(), SerializeError> {
        if let Some(ref val) = self.frameset {
            val.write_element("w:frameset", writer)?;
        }
        if let Some(ref val) = self.divs {
            val.write_element("w:divs", writer)?;
        }
        if let Some(ref val) = self.encoding {
            val.write_element("w:encoding", writer)?;
        }
        if let Some(ref val) = self.optimize_for_browser {
            val.write_element("w:optimizeForBrowser", writer)?;
        }
        if let Some(ref val) = self.rely_on_v_m_l {
            val.write_element("w:relyOnVML", writer)?;
        }
        if let Some(ref val) = self.allow_p_n_g {
            val.write_element("w:allowPNG", writer)?;
        }
        if let Some(ref val) = self.do_not_rely_on_c_s_s {
            val.write_element("w:doNotRelyOnCSS", writer)?;
        }
        if let Some(ref val) = self.do_not_save_as_single_file {
            val.write_element("w:doNotSaveAsSingleFile", writer)?;
        }
        if let Some(ref val) = self.do_not_organize_in_folder {
            val.write_element("w:doNotOrganizeInFolder", writer)?;
        }
        if let Some(ref val) = self.do_not_use_long_file_names {
            val.write_element("w:doNotUseLongFileNames", writer)?;
        }
        if let Some(ref val) = self.pixels_per_inch {
            val.write_element("w:pixelsPerInch", writer)?;
        }
        if let Some(ref val) = self.target_screen_sz {
            val.write_element("w:targetScreenSz", writer)?;
        }
        if let Some(ref val) = self.save_smart_tags_as_xml {
            val.write_element("w:saveSmartTagsAsXml", writer)?;
        }
        #[cfg(feature = "extra-children")]
        for child in &self.extra_children {
            child.write_to(writer).map_err(SerializeError::from)?;
        }
        Ok(())
    }

    fn is_empty_element(&self) -> bool {
        if self.frameset.is_some() {
            return false;
        }
        if self.divs.is_some() {
            return false;
        }
        if self.encoding.is_some() {
            return false;
        }
        if self.optimize_for_browser.is_some() {
            return false;
        }
        if self.rely_on_v_m_l.is_some() {
            return false;
        }
        if self.allow_p_n_g.is_some() {
            return false;
        }
        if self.do_not_rely_on_c_s_s.is_some() {
            return false;
        }
        if self.do_not_save_as_single_file.is_some() {
            return false;
        }
        if self.do_not_organize_in_folder.is_some() {
            return false;
        }
        if self.do_not_use_long_file_names.is_some() {
            return false;
        }
        if self.pixels_per_inch.is_some() {
            return false;
        }
        if self.target_screen_sz.is_some() {
            return false;
        }
        if self.save_smart_tags_as_xml.is_some() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for CTFrameScrollbar {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        {
            let val = &self.value;
            {
                let s = val.to_string();
                start.push_attribute(("w:val", s.as_str()));
            }
        }
        #[cfg(feature = "extra-attrs")]
        for (key, value) in &self.extra_attrs {
            start.push_attribute((key.as_str(), value.as_str()));
        }
        start
    }

    fn is_empty_element(&self) -> bool {
        true
    }
}

impl ToXml for CTOptimizeForBrowser {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        if let Some(ref val) = self.value {
            {
                let s = val.to_string();
                start.push_attribute(("w:val", s.as_str()));
            }
        }
        if let Some(ref val) = self.target {
            start.push_attribute(("w:target", val.as_str()));
        }
        #[cfg(feature = "extra-attrs")]
        for (key, value) in &self.extra_attrs {
            start.push_attribute((key.as_str(), value.as_str()));
        }
        start
    }

    fn is_empty_element(&self) -> bool {
        true
    }
}

impl ToXml for CTFrame {
    fn write_children<W: Write>(&self, writer: &mut Writer<W>) -> Result<(), SerializeError> {
        if let Some(ref val) = self.size {
            val.write_element("w:sz", writer)?;
        }
        if let Some(ref val) = self.name {
            val.write_element("w:name", writer)?;
        }
        if let Some(ref val) = self.title {
            val.write_element("w:title", writer)?;
        }
        if let Some(ref val) = self.long_desc {
            val.write_element("w:longDesc", writer)?;
        }
        if let Some(ref val) = self.source_file_name {
            val.write_element("w:sourceFileName", writer)?;
        }
        if let Some(ref val) = self.mar_w {
            val.write_element("w:marW", writer)?;
        }
        if let Some(ref val) = self.mar_h {
            val.write_element("w:marH", writer)?;
        }
        if let Some(ref val) = self.scrollbar {
            val.write_element("w:scrollbar", writer)?;
        }
        if let Some(ref val) = self.no_resize_allowed {
            val.write_element("w:noResizeAllowed", writer)?;
        }
        if let Some(ref val) = self.linked_to_file {
            val.write_element("w:linkedToFile", writer)?;
        }
        #[cfg(feature = "extra-children")]
        for child in &self.extra_children {
            child.write_to(writer).map_err(SerializeError::from)?;
        }
        Ok(())
    }

    fn is_empty_element(&self) -> bool {
        if self.size.is_some() {
            return false;
        }
        if self.name.is_some() {
            return false;
        }
        if self.title.is_some() {
            return false;
        }
        if self.long_desc.is_some() {
            return false;
        }
        if self.source_file_name.is_some() {
            return false;
        }
        if self.mar_w.is_some() {
            return false;
        }
        if self.mar_h.is_some() {
            return false;
        }
        if self.scrollbar.is_some() {
            return false;
        }
        if self.no_resize_allowed.is_some() {
            return false;
        }
        if self.linked_to_file.is_some() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for CTFrameLayout {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        {
            let val = &self.value;
            {
                let s = val.to_string();
                start.push_attribute(("w:val", s.as_str()));
            }
        }
        #[cfg(feature = "extra-attrs")]
        for (key, value) in &self.extra_attrs {
            start.push_attribute((key.as_str(), value.as_str()));
        }
        start
    }

    fn is_empty_element(&self) -> bool {
        true
    }
}

impl ToXml for CTFramesetSplitbar {
    fn write_children<W: Write>(&self, writer: &mut Writer<W>) -> Result<(), SerializeError> {
        if let Some(ref val) = self.width {
            val.write_element("w:w", writer)?;
        }
        if let Some(ref val) = self.color {
            val.write_element("w:color", writer)?;
        }
        if let Some(ref val) = self.no_border {
            val.write_element("w:noBorder", writer)?;
        }
        if let Some(ref val) = self.flat_borders {
            val.write_element("w:flatBorders", writer)?;
        }
        #[cfg(feature = "extra-children")]
        for child in &self.extra_children {
            child.write_to(writer).map_err(SerializeError::from)?;
        }
        Ok(())
    }

    fn is_empty_element(&self) -> bool {
        if self.width.is_some() {
            return false;
        }
        if self.color.is_some() {
            return false;
        }
        if self.no_border.is_some() {
            return false;
        }
        if self.flat_borders.is_some() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for CTFrameset {
    fn write_children<W: Write>(&self, writer: &mut Writer<W>) -> Result<(), SerializeError> {
        if let Some(ref val) = self.size {
            val.write_element("w:sz", writer)?;
        }
        if let Some(ref val) = self.frameset_splitbar {
            val.write_element("w:framesetSplitbar", writer)?;
        }
        if let Some(ref val) = self.frame_layout {
            val.write_element("w:frameLayout", writer)?;
        }
        if let Some(ref val) = self.title {
            val.write_element("w:title", writer)?;
        }
        #[cfg(feature = "extra-children")]
        for child in &self.extra_children {
            child.write_to(writer).map_err(SerializeError::from)?;
        }
        Ok(())
    }

    fn is_empty_element(&self) -> bool {
        if self.size.is_some() {
            return false;
        }
        if self.frameset_splitbar.is_some() {
            return false;
        }
        if self.frame_layout.is_some() {
            return false;
        }
        if self.title.is_some() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for CTNumPicBullet {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        {
            let val = &self.num_pic_bullet_id;
            {
                let s = val.to_string();
                start.push_attribute(("w:numPicBulletId", s.as_str()));
            }
        }
        #[cfg(feature = "extra-attrs")]
        for (key, value) in &self.extra_attrs {
            start.push_attribute((key.as_str(), value.as_str()));
        }
        start
    }

    fn is_empty_element(&self) -> bool {
        true
    }
}

impl ToXml for CTLevelSuffix {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        {
            let val = &self.value;
            {
                let s = val.to_string();
                start.push_attribute(("w:val", s.as_str()));
            }
        }
        #[cfg(feature = "extra-attrs")]
        for (key, value) in &self.extra_attrs {
            start.push_attribute((key.as_str(), value.as_str()));
        }
        start
    }

    fn is_empty_element(&self) -> bool {
        true
    }
}

impl ToXml for CTLevelText {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        if let Some(ref val) = self.value {
            start.push_attribute(("w:val", val.as_str()));
        }
        if let Some(ref val) = self.null {
            {
                let s = val.to_string();
                start.push_attribute(("w:null", s.as_str()));
            }
        }
        #[cfg(feature = "extra-attrs")]
        for (key, value) in &self.extra_attrs {
            start.push_attribute((key.as_str(), value.as_str()));
        }
        start
    }

    fn is_empty_element(&self) -> bool {
        true
    }
}

impl ToXml for CTLvlLegacy {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        if let Some(ref val) = self.legacy {
            {
                let s = val.to_string();
                start.push_attribute(("w:legacy", s.as_str()));
            }
        }
        if let Some(ref val) = self.legacy_space {
            {
                let s = val.to_string();
                start.push_attribute(("w:legacySpace", s.as_str()));
            }
        }
        if let Some(ref val) = self.legacy_indent {
            {
                let s = val.to_string();
                start.push_attribute(("w:legacyIndent", s.as_str()));
            }
        }
        #[cfg(feature = "extra-attrs")]
        for (key, value) in &self.extra_attrs {
            start.push_attribute((key.as_str(), value.as_str()));
        }
        start
    }

    fn is_empty_element(&self) -> bool {
        true
    }
}

impl ToXml for Level {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        {
            let val = &self.ilvl;
            {
                let s = val.to_string();
                start.push_attribute(("w:ilvl", s.as_str()));
            }
        }
        if let Some(ref val) = self.tplc {
            {
                let hex = encode_hex(val);
                start.push_attribute(("w:tplc", hex.as_str()));
            }
        }
        if let Some(ref val) = self.tentative {
            {
                let s = val.to_string();
                start.push_attribute(("w:tentative", s.as_str()));
            }
        }
        #[cfg(feature = "extra-attrs")]
        for (key, value) in &self.extra_attrs {
            start.push_attribute((key.as_str(), value.as_str()));
        }
        start
    }

    fn write_children<W: Write>(&self, writer: &mut Writer<W>) -> Result<(), SerializeError> {
        if let Some(ref val) = self.start {
            val.write_element("w:start", writer)?;
        }
        if let Some(ref val) = self.num_fmt {
            val.write_element("w:numFmt", writer)?;
        }
        if let Some(ref val) = self.lvl_restart {
            val.write_element("w:lvlRestart", writer)?;
        }
        if let Some(ref val) = self.paragraph_style {
            val.write_element("w:pStyle", writer)?;
        }
        if let Some(ref val) = self.is_lgl {
            val.write_element("w:isLgl", writer)?;
        }
        if let Some(ref val) = self.suff {
            val.write_element("w:suff", writer)?;
        }
        if let Some(ref val) = self.lvl_text {
            val.write_element("w:lvlText", writer)?;
        }
        if let Some(ref val) = self.lvl_pic_bullet_id {
            val.write_element("w:lvlPicBulletId", writer)?;
        }
        if let Some(ref val) = self.legacy {
            val.write_element("w:legacy", writer)?;
        }
        if let Some(ref val) = self.lvl_jc {
            val.write_element("w:lvlJc", writer)?;
        }
        if let Some(ref val) = self.p_pr {
            val.write_element("w:pPr", writer)?;
        }
        if let Some(ref val) = self.r_pr {
            val.write_element("w:rPr", writer)?;
        }
        #[cfg(feature = "extra-children")]
        for child in &self.extra_children {
            child.write_to(writer).map_err(SerializeError::from)?;
        }
        Ok(())
    }

    fn is_empty_element(&self) -> bool {
        if self.start.is_some() {
            return false;
        }
        if self.num_fmt.is_some() {
            return false;
        }
        if self.lvl_restart.is_some() {
            return false;
        }
        if self.paragraph_style.is_some() {
            return false;
        }
        if self.is_lgl.is_some() {
            return false;
        }
        if self.suff.is_some() {
            return false;
        }
        if self.lvl_text.is_some() {
            return false;
        }
        if self.lvl_pic_bullet_id.is_some() {
            return false;
        }
        if self.legacy.is_some() {
            return false;
        }
        if self.lvl_jc.is_some() {
            return false;
        }
        if self.p_pr.is_some() {
            return false;
        }
        if self.r_pr.is_some() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for CTMultiLevelType {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        {
            let val = &self.value;
            {
                let s = val.to_string();
                start.push_attribute(("w:val", s.as_str()));
            }
        }
        #[cfg(feature = "extra-attrs")]
        for (key, value) in &self.extra_attrs {
            start.push_attribute((key.as_str(), value.as_str()));
        }
        start
    }

    fn is_empty_element(&self) -> bool {
        true
    }
}

impl ToXml for AbstractNumbering {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        {
            let val = &self.abstract_num_id;
            {
                let s = val.to_string();
                start.push_attribute(("w:abstractNumId", s.as_str()));
            }
        }
        #[cfg(feature = "extra-attrs")]
        for (key, value) in &self.extra_attrs {
            start.push_attribute((key.as_str(), value.as_str()));
        }
        start
    }

    fn write_children<W: Write>(&self, writer: &mut Writer<W>) -> Result<(), SerializeError> {
        if let Some(ref val) = self.nsid {
            val.write_element("w:nsid", writer)?;
        }
        if let Some(ref val) = self.multi_level_type {
            val.write_element("w:multiLevelType", writer)?;
        }
        if let Some(ref val) = self.tmpl {
            val.write_element("w:tmpl", writer)?;
        }
        if let Some(ref val) = self.name {
            val.write_element("w:name", writer)?;
        }
        if let Some(ref val) = self.style_link {
            val.write_element("w:styleLink", writer)?;
        }
        if let Some(ref val) = self.num_style_link {
            val.write_element("w:numStyleLink", writer)?;
        }
        for item in &self.lvl {
            item.write_element("w:lvl", writer)?;
        }
        #[cfg(feature = "extra-children")]
        for child in &self.extra_children {
            child.write_to(writer).map_err(SerializeError::from)?;
        }
        Ok(())
    }

    fn is_empty_element(&self) -> bool {
        if self.nsid.is_some() {
            return false;
        }
        if self.multi_level_type.is_some() {
            return false;
        }
        if self.tmpl.is_some() {
            return false;
        }
        if self.name.is_some() {
            return false;
        }
        if self.style_link.is_some() {
            return false;
        }
        if self.num_style_link.is_some() {
            return false;
        }
        if !self.lvl.is_empty() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for CTNumLvl {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        {
            let val = &self.ilvl;
            {
                let s = val.to_string();
                start.push_attribute(("w:ilvl", s.as_str()));
            }
        }
        #[cfg(feature = "extra-attrs")]
        for (key, value) in &self.extra_attrs {
            start.push_attribute((key.as_str(), value.as_str()));
        }
        start
    }

    fn write_children<W: Write>(&self, writer: &mut Writer<W>) -> Result<(), SerializeError> {
        if let Some(ref val) = self.start_override {
            val.write_element("w:startOverride", writer)?;
        }
        if let Some(ref val) = self.lvl {
            val.write_element("w:lvl", writer)?;
        }
        #[cfg(feature = "extra-children")]
        for child in &self.extra_children {
            child.write_to(writer).map_err(SerializeError::from)?;
        }
        Ok(())
    }

    fn is_empty_element(&self) -> bool {
        if self.start_override.is_some() {
            return false;
        }
        if self.lvl.is_some() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for NumberingInstance {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        {
            let val = &self.num_id;
            {
                let s = val.to_string();
                start.push_attribute(("w:numId", s.as_str()));
            }
        }
        #[cfg(feature = "extra-attrs")]
        for (key, value) in &self.extra_attrs {
            start.push_attribute((key.as_str(), value.as_str()));
        }
        start
    }

    fn write_children<W: Write>(&self, writer: &mut Writer<W>) -> Result<(), SerializeError> {
        {
            let val = &self.abstract_num_id;
            val.write_element("w:abstractNumId", writer)?;
        }
        for item in &self.lvl_override {
            item.write_element("w:lvlOverride", writer)?;
        }
        #[cfg(feature = "extra-children")]
        for child in &self.extra_children {
            child.write_to(writer).map_err(SerializeError::from)?;
        }
        Ok(())
    }

    fn is_empty_element(&self) -> bool {
        false
    }
}

impl ToXml for Numbering {
    fn write_children<W: Write>(&self, writer: &mut Writer<W>) -> Result<(), SerializeError> {
        for item in &self.num_pic_bullet {
            item.write_element("w:numPicBullet", writer)?;
        }
        for item in &self.abstract_num {
            item.write_element("w:abstractNum", writer)?;
        }
        for item in &self.num {
            item.write_element("w:num", writer)?;
        }
        if let Some(ref val) = self.num_id_mac_at_cleanup {
            val.write_element("w:numIdMacAtCleanup", writer)?;
        }
        #[cfg(feature = "extra-children")]
        for child in &self.extra_children {
            child.write_to(writer).map_err(SerializeError::from)?;
        }
        Ok(())
    }

    fn is_empty_element(&self) -> bool {
        if !self.num_pic_bullet.is_empty() {
            return false;
        }
        if !self.abstract_num.is_empty() {
            return false;
        }
        if !self.num.is_empty() {
            return false;
        }
        if self.num_id_mac_at_cleanup.is_some() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for CTTblStylePr {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        {
            let val = &self.r#type;
            {
                let s = val.to_string();
                start.push_attribute(("w:type", s.as_str()));
            }
        }
        #[cfg(feature = "extra-attrs")]
        for (key, value) in &self.extra_attrs {
            start.push_attribute((key.as_str(), value.as_str()));
        }
        start
    }

    fn write_children<W: Write>(&self, writer: &mut Writer<W>) -> Result<(), SerializeError> {
        if let Some(ref val) = self.p_pr {
            val.write_element("w:pPr", writer)?;
        }
        if let Some(ref val) = self.r_pr {
            val.write_element("w:rPr", writer)?;
        }
        if let Some(ref val) = self.table_properties {
            val.write_element("w:tblPr", writer)?;
        }
        if let Some(ref val) = self.row_properties {
            val.write_element("w:trPr", writer)?;
        }
        if let Some(ref val) = self.cell_properties {
            val.write_element("w:tcPr", writer)?;
        }
        #[cfg(feature = "extra-children")]
        for child in &self.extra_children {
            child.write_to(writer).map_err(SerializeError::from)?;
        }
        Ok(())
    }

    fn is_empty_element(&self) -> bool {
        if self.p_pr.is_some() {
            return false;
        }
        if self.r_pr.is_some() {
            return false;
        }
        if self.table_properties.is_some() {
            return false;
        }
        if self.row_properties.is_some() {
            return false;
        }
        if self.cell_properties.is_some() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for Style {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        if let Some(ref val) = self.r#type {
            {
                let s = val.to_string();
                start.push_attribute(("w:type", s.as_str()));
            }
        }
        if let Some(ref val) = self.style_id {
            start.push_attribute(("w:styleId", val.as_str()));
        }
        if let Some(ref val) = self.default {
            {
                let s = val.to_string();
                start.push_attribute(("w:default", s.as_str()));
            }
        }
        if let Some(ref val) = self.custom_style {
            {
                let s = val.to_string();
                start.push_attribute(("w:customStyle", s.as_str()));
            }
        }
        #[cfg(feature = "extra-attrs")]
        for (key, value) in &self.extra_attrs {
            start.push_attribute((key.as_str(), value.as_str()));
        }
        start
    }

    fn write_children<W: Write>(&self, writer: &mut Writer<W>) -> Result<(), SerializeError> {
        if let Some(ref val) = self.name {
            val.write_element("w:name", writer)?;
        }
        if let Some(ref val) = self.aliases {
            val.write_element("w:aliases", writer)?;
        }
        if let Some(ref val) = self.based_on {
            val.write_element("w:basedOn", writer)?;
        }
        if let Some(ref val) = self.next {
            val.write_element("w:next", writer)?;
        }
        if let Some(ref val) = self.link {
            val.write_element("w:link", writer)?;
        }
        if let Some(ref val) = self.auto_redefine {
            val.write_element("w:autoRedefine", writer)?;
        }
        if let Some(ref val) = self.hidden {
            val.write_element("w:hidden", writer)?;
        }
        if let Some(ref val) = self.ui_priority {
            val.write_element("w:uiPriority", writer)?;
        }
        if let Some(ref val) = self.semi_hidden {
            val.write_element("w:semiHidden", writer)?;
        }
        if let Some(ref val) = self.unhide_when_used {
            val.write_element("w:unhideWhenUsed", writer)?;
        }
        if let Some(ref val) = self.q_format {
            val.write_element("w:qFormat", writer)?;
        }
        if let Some(ref val) = self.locked {
            val.write_element("w:locked", writer)?;
        }
        if let Some(ref val) = self.personal {
            val.write_element("w:personal", writer)?;
        }
        if let Some(ref val) = self.personal_compose {
            val.write_element("w:personalCompose", writer)?;
        }
        if let Some(ref val) = self.personal_reply {
            val.write_element("w:personalReply", writer)?;
        }
        if let Some(ref val) = self.rsid {
            val.write_element("w:rsid", writer)?;
        }
        if let Some(ref val) = self.p_pr {
            val.write_element("w:pPr", writer)?;
        }
        if let Some(ref val) = self.r_pr {
            val.write_element("w:rPr", writer)?;
        }
        if let Some(ref val) = self.table_properties {
            val.write_element("w:tblPr", writer)?;
        }
        if let Some(ref val) = self.row_properties {
            val.write_element("w:trPr", writer)?;
        }
        if let Some(ref val) = self.cell_properties {
            val.write_element("w:tcPr", writer)?;
        }
        for item in &self.tbl_style_pr {
            item.write_element("w:tblStylePr", writer)?;
        }
        #[cfg(feature = "extra-children")]
        for child in &self.extra_children {
            child.write_to(writer).map_err(SerializeError::from)?;
        }
        Ok(())
    }

    fn is_empty_element(&self) -> bool {
        if self.name.is_some() {
            return false;
        }
        if self.aliases.is_some() {
            return false;
        }
        if self.based_on.is_some() {
            return false;
        }
        if self.next.is_some() {
            return false;
        }
        if self.link.is_some() {
            return false;
        }
        if self.auto_redefine.is_some() {
            return false;
        }
        if self.hidden.is_some() {
            return false;
        }
        if self.ui_priority.is_some() {
            return false;
        }
        if self.semi_hidden.is_some() {
            return false;
        }
        if self.unhide_when_used.is_some() {
            return false;
        }
        if self.q_format.is_some() {
            return false;
        }
        if self.locked.is_some() {
            return false;
        }
        if self.personal.is_some() {
            return false;
        }
        if self.personal_compose.is_some() {
            return false;
        }
        if self.personal_reply.is_some() {
            return false;
        }
        if self.rsid.is_some() {
            return false;
        }
        if self.p_pr.is_some() {
            return false;
        }
        if self.r_pr.is_some() {
            return false;
        }
        if self.table_properties.is_some() {
            return false;
        }
        if self.row_properties.is_some() {
            return false;
        }
        if self.cell_properties.is_some() {
            return false;
        }
        if !self.tbl_style_pr.is_empty() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for CTLsdException {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        {
            let val = &self.name;
            start.push_attribute(("w:name", val.as_str()));
        }
        if let Some(ref val) = self.locked {
            {
                let s = val.to_string();
                start.push_attribute(("w:locked", s.as_str()));
            }
        }
        if let Some(ref val) = self.ui_priority {
            {
                let s = val.to_string();
                start.push_attribute(("w:uiPriority", s.as_str()));
            }
        }
        if let Some(ref val) = self.semi_hidden {
            {
                let s = val.to_string();
                start.push_attribute(("w:semiHidden", s.as_str()));
            }
        }
        if let Some(ref val) = self.unhide_when_used {
            {
                let s = val.to_string();
                start.push_attribute(("w:unhideWhenUsed", s.as_str()));
            }
        }
        if let Some(ref val) = self.q_format {
            {
                let s = val.to_string();
                start.push_attribute(("w:qFormat", s.as_str()));
            }
        }
        #[cfg(feature = "extra-attrs")]
        for (key, value) in &self.extra_attrs {
            start.push_attribute((key.as_str(), value.as_str()));
        }
        start
    }

    fn is_empty_element(&self) -> bool {
        true
    }
}

impl ToXml for CTLatentStyles {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        if let Some(ref val) = self.def_locked_state {
            {
                let s = val.to_string();
                start.push_attribute(("w:defLockedState", s.as_str()));
            }
        }
        if let Some(ref val) = self.def_u_i_priority {
            {
                let s = val.to_string();
                start.push_attribute(("w:defUIPriority", s.as_str()));
            }
        }
        if let Some(ref val) = self.def_semi_hidden {
            {
                let s = val.to_string();
                start.push_attribute(("w:defSemiHidden", s.as_str()));
            }
        }
        if let Some(ref val) = self.def_unhide_when_used {
            {
                let s = val.to_string();
                start.push_attribute(("w:defUnhideWhenUsed", s.as_str()));
            }
        }
        if let Some(ref val) = self.def_q_format {
            {
                let s = val.to_string();
                start.push_attribute(("w:defQFormat", s.as_str()));
            }
        }
        if let Some(ref val) = self.count {
            {
                let s = val.to_string();
                start.push_attribute(("w:count", s.as_str()));
            }
        }
        #[cfg(feature = "extra-attrs")]
        for (key, value) in &self.extra_attrs {
            start.push_attribute((key.as_str(), value.as_str()));
        }
        start
    }

    fn write_children<W: Write>(&self, writer: &mut Writer<W>) -> Result<(), SerializeError> {
        for item in &self.lsd_exception {
            item.write_element("w:lsdException", writer)?;
        }
        #[cfg(feature = "extra-children")]
        for child in &self.extra_children {
            child.write_to(writer).map_err(SerializeError::from)?;
        }
        Ok(())
    }

    fn is_empty_element(&self) -> bool {
        if !self.lsd_exception.is_empty() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for Styles {
    fn write_children<W: Write>(&self, writer: &mut Writer<W>) -> Result<(), SerializeError> {
        if let Some(ref val) = self.doc_defaults {
            val.write_element("w:docDefaults", writer)?;
        }
        if let Some(ref val) = self.latent_styles {
            val.write_element("w:latentStyles", writer)?;
        }
        for item in &self.style {
            item.write_element("w:style", writer)?;
        }
        #[cfg(feature = "extra-children")]
        for child in &self.extra_children {
            child.write_to(writer).map_err(SerializeError::from)?;
        }
        Ok(())
    }

    fn is_empty_element(&self) -> bool {
        if self.doc_defaults.is_some() {
            return false;
        }
        if self.latent_styles.is_some() {
            return false;
        }
        if !self.style.is_empty() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for CTPanose {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        {
            let val = &self.value;
            {
                let hex = encode_hex(val);
                start.push_attribute(("w:val", hex.as_str()));
            }
        }
        #[cfg(feature = "extra-attrs")]
        for (key, value) in &self.extra_attrs {
            start.push_attribute((key.as_str(), value.as_str()));
        }
        start
    }

    fn is_empty_element(&self) -> bool {
        true
    }
}

impl ToXml for CTFontFamily {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        {
            let val = &self.value;
            {
                let s = val.to_string();
                start.push_attribute(("w:val", s.as_str()));
            }
        }
        #[cfg(feature = "extra-attrs")]
        for (key, value) in &self.extra_attrs {
            start.push_attribute((key.as_str(), value.as_str()));
        }
        start
    }

    fn is_empty_element(&self) -> bool {
        true
    }
}

impl ToXml for CTPitch {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        {
            let val = &self.value;
            {
                let s = val.to_string();
                start.push_attribute(("w:val", s.as_str()));
            }
        }
        #[cfg(feature = "extra-attrs")]
        for (key, value) in &self.extra_attrs {
            start.push_attribute((key.as_str(), value.as_str()));
        }
        start
    }

    fn is_empty_element(&self) -> bool {
        true
    }
}

impl ToXml for CTFontSig {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        {
            let val = &self.usb0;
            {
                let hex = encode_hex(val);
                start.push_attribute(("w:usb0", hex.as_str()));
            }
        }
        {
            let val = &self.usb1;
            {
                let hex = encode_hex(val);
                start.push_attribute(("w:usb1", hex.as_str()));
            }
        }
        {
            let val = &self.usb2;
            {
                let hex = encode_hex(val);
                start.push_attribute(("w:usb2", hex.as_str()));
            }
        }
        {
            let val = &self.usb3;
            {
                let hex = encode_hex(val);
                start.push_attribute(("w:usb3", hex.as_str()));
            }
        }
        {
            let val = &self.csb0;
            {
                let hex = encode_hex(val);
                start.push_attribute(("w:csb0", hex.as_str()));
            }
        }
        {
            let val = &self.csb1;
            {
                let hex = encode_hex(val);
                start.push_attribute(("w:csb1", hex.as_str()));
            }
        }
        #[cfg(feature = "extra-attrs")]
        for (key, value) in &self.extra_attrs {
            start.push_attribute((key.as_str(), value.as_str()));
        }
        start
    }

    fn is_empty_element(&self) -> bool {
        true
    }
}

impl ToXml for CTFontRel {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        if let Some(ref val) = self.font_key {
            start.push_attribute(("w:fontKey", val.as_str()));
        }
        if let Some(ref val) = self.subsetted {
            {
                let s = val.to_string();
                start.push_attribute(("w:subsetted", s.as_str()));
            }
        }
        #[cfg(feature = "extra-attrs")]
        for (key, value) in &self.extra_attrs {
            start.push_attribute((key.as_str(), value.as_str()));
        }
        start
    }

    fn is_empty_element(&self) -> bool {
        true
    }
}

impl ToXml for Font {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        {
            let val = &self.name;
            start.push_attribute(("w:name", val.as_str()));
        }
        #[cfg(feature = "extra-attrs")]
        for (key, value) in &self.extra_attrs {
            start.push_attribute((key.as_str(), value.as_str()));
        }
        start
    }

    fn write_children<W: Write>(&self, writer: &mut Writer<W>) -> Result<(), SerializeError> {
        if let Some(ref val) = self.alt_name {
            val.write_element("w:altName", writer)?;
        }
        if let Some(ref val) = self.panose1 {
            val.write_element("w:panose1", writer)?;
        }
        if let Some(ref val) = self.charset {
            val.write_element("w:charset", writer)?;
        }
        if let Some(ref val) = self.family {
            val.write_element("w:family", writer)?;
        }
        if let Some(ref val) = self.not_true_type {
            val.write_element("w:notTrueType", writer)?;
        }
        if let Some(ref val) = self.pitch {
            val.write_element("w:pitch", writer)?;
        }
        if let Some(ref val) = self.sig {
            val.write_element("w:sig", writer)?;
        }
        if let Some(ref val) = self.embed_regular {
            val.write_element("w:embedRegular", writer)?;
        }
        if let Some(ref val) = self.embed_bold {
            val.write_element("w:embedBold", writer)?;
        }
        if let Some(ref val) = self.embed_italic {
            val.write_element("w:embedItalic", writer)?;
        }
        if let Some(ref val) = self.embed_bold_italic {
            val.write_element("w:embedBoldItalic", writer)?;
        }
        #[cfg(feature = "extra-children")]
        for child in &self.extra_children {
            child.write_to(writer).map_err(SerializeError::from)?;
        }
        Ok(())
    }

    fn is_empty_element(&self) -> bool {
        if self.alt_name.is_some() {
            return false;
        }
        if self.panose1.is_some() {
            return false;
        }
        if self.charset.is_some() {
            return false;
        }
        if self.family.is_some() {
            return false;
        }
        if self.not_true_type.is_some() {
            return false;
        }
        if self.pitch.is_some() {
            return false;
        }
        if self.sig.is_some() {
            return false;
        }
        if self.embed_regular.is_some() {
            return false;
        }
        if self.embed_bold.is_some() {
            return false;
        }
        if self.embed_italic.is_some() {
            return false;
        }
        if self.embed_bold_italic.is_some() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for CTFontsList {
    fn write_children<W: Write>(&self, writer: &mut Writer<W>) -> Result<(), SerializeError> {
        for item in &self.font {
            item.write_element("w:font", writer)?;
        }
        #[cfg(feature = "extra-children")]
        for child in &self.extra_children {
            child.write_to(writer).map_err(SerializeError::from)?;
        }
        Ok(())
    }

    fn is_empty_element(&self) -> bool {
        if !self.font.is_empty() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for CTDivBdr {
    fn write_children<W: Write>(&self, writer: &mut Writer<W>) -> Result<(), SerializeError> {
        if let Some(ref val) = self.top {
            val.write_element("w:top", writer)?;
        }
        if let Some(ref val) = self.left {
            val.write_element("w:left", writer)?;
        }
        if let Some(ref val) = self.bottom {
            val.write_element("w:bottom", writer)?;
        }
        if let Some(ref val) = self.right {
            val.write_element("w:right", writer)?;
        }
        #[cfg(feature = "extra-children")]
        for child in &self.extra_children {
            child.write_to(writer).map_err(SerializeError::from)?;
        }
        Ok(())
    }

    fn is_empty_element(&self) -> bool {
        if self.top.is_some() {
            return false;
        }
        if self.left.is_some() {
            return false;
        }
        if self.bottom.is_some() {
            return false;
        }
        if self.right.is_some() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for CTDiv {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        {
            let val = &self.id;
            {
                let s = val.to_string();
                start.push_attribute(("w:id", s.as_str()));
            }
        }
        #[cfg(feature = "extra-attrs")]
        for (key, value) in &self.extra_attrs {
            start.push_attribute((key.as_str(), value.as_str()));
        }
        start
    }

    fn write_children<W: Write>(&self, writer: &mut Writer<W>) -> Result<(), SerializeError> {
        if let Some(ref val) = self.block_quote {
            val.write_element("w:blockQuote", writer)?;
        }
        if let Some(ref val) = self.body_div {
            val.write_element("w:bodyDiv", writer)?;
        }
        {
            let val = &self.mar_left;
            val.write_element("w:marLeft", writer)?;
        }
        {
            let val = &self.mar_right;
            val.write_element("w:marRight", writer)?;
        }
        {
            let val = &self.mar_top;
            val.write_element("w:marTop", writer)?;
        }
        {
            let val = &self.mar_bottom;
            val.write_element("w:marBottom", writer)?;
        }
        if let Some(ref val) = self.div_bdr {
            val.write_element("w:divBdr", writer)?;
        }
        for item in &self.divs_child {
            item.write_element("w:divsChild", writer)?;
        }
        #[cfg(feature = "extra-children")]
        for child in &self.extra_children {
            child.write_to(writer).map_err(SerializeError::from)?;
        }
        Ok(())
    }

    fn is_empty_element(&self) -> bool {
        if self.block_quote.is_some() {
            return false;
        }
        if self.body_div.is_some() {
            return false;
        }
        false
    }
}

impl ToXml for CTDivs {
    fn write_children<W: Write>(&self, writer: &mut Writer<W>) -> Result<(), SerializeError> {
        for item in &self.div {
            item.write_element("w:div", writer)?;
        }
        #[cfg(feature = "extra-children")]
        for child in &self.extra_children {
            child.write_to(writer).map_err(SerializeError::from)?;
        }
        Ok(())
    }

    fn is_empty_element(&self) -> bool {
        if !self.div.is_empty() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for CTTxbxContent {
    fn write_children<W: Write>(&self, writer: &mut Writer<W>) -> Result<(), SerializeError> {
        for item in &self.block_level_elts {
            item.write_element("", writer)?;
        }
        #[cfg(feature = "extra-children")]
        for child in &self.extra_children {
            child.write_to(writer).map_err(SerializeError::from)?;
        }
        Ok(())
    }

    fn is_empty_element(&self) -> bool {
        if !self.block_level_elts.is_empty() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for EGMathContent {
    fn write_children<W: Write>(&self, writer: &mut Writer<W>) -> Result<(), SerializeError> {
        #[cfg(feature = "extra-children")]
        for child in &self.extra_children {
            child.write_to(writer).map_err(SerializeError::from)?;
        }
        Ok(())
    }

    fn is_empty_element(&self) -> bool {
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for EGBlockLevelChunkElts {
    fn write_children<W: Write>(&self, writer: &mut Writer<W>) -> Result<(), SerializeError> {
        for item in &self.content_block_content {
            item.write_element("", writer)?;
        }
        #[cfg(feature = "extra-children")]
        for child in &self.extra_children {
            child.write_to(writer).map_err(SerializeError::from)?;
        }
        Ok(())
    }

    fn is_empty_element(&self) -> bool {
        if !self.content_block_content.is_empty() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for EGBlockLevelElts {
    fn write_element<W: Write>(
        &self,
        _tag: &str,
        writer: &mut Writer<W>,
    ) -> Result<(), SerializeError> {
        match self {
            Self::CustomXml(inner) => inner.write_element("w:customXml", writer)?,
            Self::Sdt(inner) => inner.write_element("w:sdt", writer)?,
            Self::P(inner) => inner.write_element("w:p", writer)?,
            Self::Tbl(inner) => inner.write_element("w:tbl", writer)?,
            Self::ProofErr(inner) => inner.write_element("w:proofErr", writer)?,
            Self::PermStart(inner) => inner.write_element("w:permStart", writer)?,
            Self::PermEnd(inner) => inner.write_element("w:permEnd", writer)?,
            Self::BookmarkStart(inner) => inner.write_element("w:bookmarkStart", writer)?,
            Self::BookmarkEnd(inner) => inner.write_element("w:bookmarkEnd", writer)?,
            Self::MoveFromRangeStart(inner) => {
                inner.write_element("w:moveFromRangeStart", writer)?
            }
            Self::MoveFromRangeEnd(inner) => inner.write_element("w:moveFromRangeEnd", writer)?,
            Self::MoveToRangeStart(inner) => inner.write_element("w:moveToRangeStart", writer)?,
            Self::MoveToRangeEnd(inner) => inner.write_element("w:moveToRangeEnd", writer)?,
            Self::CommentRangeStart(inner) => inner.write_element("w:commentRangeStart", writer)?,
            Self::CommentRangeEnd(inner) => inner.write_element("w:commentRangeEnd", writer)?,
            Self::CustomXmlInsRangeStart(inner) => {
                inner.write_element("w:customXmlInsRangeStart", writer)?
            }
            Self::CustomXmlInsRangeEnd(inner) => {
                inner.write_element("w:customXmlInsRangeEnd", writer)?
            }
            Self::CustomXmlDelRangeStart(inner) => {
                inner.write_element("w:customXmlDelRangeStart", writer)?
            }
            Self::CustomXmlDelRangeEnd(inner) => {
                inner.write_element("w:customXmlDelRangeEnd", writer)?
            }
            Self::CustomXmlMoveFromRangeStart(inner) => {
                inner.write_element("w:customXmlMoveFromRangeStart", writer)?
            }
            Self::CustomXmlMoveFromRangeEnd(inner) => {
                inner.write_element("w:customXmlMoveFromRangeEnd", writer)?
            }
            Self::CustomXmlMoveToRangeStart(inner) => {
                inner.write_element("w:customXmlMoveToRangeStart", writer)?
            }
            Self::CustomXmlMoveToRangeEnd(inner) => {
                inner.write_element("w:customXmlMoveToRangeEnd", writer)?
            }
            Self::Ins(inner) => inner.write_element("w:ins", writer)?,
            Self::Del(inner) => inner.write_element("w:del", writer)?,
            Self::MoveFrom(inner) => inner.write_element("w:moveFrom", writer)?,
            Self::MoveTo(inner) => inner.write_element("w:moveTo", writer)?,
            Self::AltChunk(inner) => inner.write_element("w:altChunk", writer)?,
        }
        Ok(())
    }
}

impl ToXml for EGRunLevelElts {
    fn write_element<W: Write>(
        &self,
        _tag: &str,
        writer: &mut Writer<W>,
    ) -> Result<(), SerializeError> {
        match self {
            Self::ProofErr(inner) => inner.write_element("w:proofErr", writer)?,
            Self::PermStart(inner) => inner.write_element("w:permStart", writer)?,
            Self::PermEnd(inner) => inner.write_element("w:permEnd", writer)?,
            Self::BookmarkStart(inner) => inner.write_element("w:bookmarkStart", writer)?,
            Self::BookmarkEnd(inner) => inner.write_element("w:bookmarkEnd", writer)?,
            Self::MoveFromRangeStart(inner) => {
                inner.write_element("w:moveFromRangeStart", writer)?
            }
            Self::MoveFromRangeEnd(inner) => inner.write_element("w:moveFromRangeEnd", writer)?,
            Self::MoveToRangeStart(inner) => inner.write_element("w:moveToRangeStart", writer)?,
            Self::MoveToRangeEnd(inner) => inner.write_element("w:moveToRangeEnd", writer)?,
            Self::CommentRangeStart(inner) => inner.write_element("w:commentRangeStart", writer)?,
            Self::CommentRangeEnd(inner) => inner.write_element("w:commentRangeEnd", writer)?,
            Self::CustomXmlInsRangeStart(inner) => {
                inner.write_element("w:customXmlInsRangeStart", writer)?
            }
            Self::CustomXmlInsRangeEnd(inner) => {
                inner.write_element("w:customXmlInsRangeEnd", writer)?
            }
            Self::CustomXmlDelRangeStart(inner) => {
                inner.write_element("w:customXmlDelRangeStart", writer)?
            }
            Self::CustomXmlDelRangeEnd(inner) => {
                inner.write_element("w:customXmlDelRangeEnd", writer)?
            }
            Self::CustomXmlMoveFromRangeStart(inner) => {
                inner.write_element("w:customXmlMoveFromRangeStart", writer)?
            }
            Self::CustomXmlMoveFromRangeEnd(inner) => {
                inner.write_element("w:customXmlMoveFromRangeEnd", writer)?
            }
            Self::CustomXmlMoveToRangeStart(inner) => {
                inner.write_element("w:customXmlMoveToRangeStart", writer)?
            }
            Self::CustomXmlMoveToRangeEnd(inner) => {
                inner.write_element("w:customXmlMoveToRangeEnd", writer)?
            }
            Self::Ins(inner) => inner.write_element("w:ins", writer)?,
            Self::Del(inner) => inner.write_element("w:del", writer)?,
            Self::MoveFrom(inner) => inner.write_element("w:moveFrom", writer)?,
            Self::MoveTo(inner) => inner.write_element("w:moveTo", writer)?,
        }
        Ok(())
    }
}

impl ToXml for Body {
    fn write_children<W: Write>(&self, writer: &mut Writer<W>) -> Result<(), SerializeError> {
        for item in &self.block_level_elts {
            item.write_element("", writer)?;
        }
        #[cfg(feature = "wml-layout")]
        if let Some(ref val) = self.sect_pr {
            val.write_element("w:sectPr", writer)?;
        }
        #[cfg(feature = "extra-children")]
        for child in &self.extra_children {
            child.write_to(writer).map_err(SerializeError::from)?;
        }
        Ok(())
    }

    fn is_empty_element(&self) -> bool {
        if !self.block_level_elts.is_empty() {
            return false;
        }
        #[cfg(feature = "wml-layout")]
        if self.sect_pr.is_some() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for CTShapeDefaults {
    fn write_children<W: Write>(&self, writer: &mut Writer<W>) -> Result<(), SerializeError> {
        #[cfg(feature = "extra-children")]
        for child in &self.extra_children {
            child.write_to(writer).map_err(SerializeError::from)?;
        }
        Ok(())
    }

    fn is_empty_element(&self) -> bool {
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for Comments {
    fn write_children<W: Write>(&self, writer: &mut Writer<W>) -> Result<(), SerializeError> {
        for item in &self.comment {
            item.write_element("w:comment", writer)?;
        }
        #[cfg(feature = "extra-children")]
        for child in &self.extra_children {
            child.write_to(writer).map_err(SerializeError::from)?;
        }
        Ok(())
    }

    fn is_empty_element(&self) -> bool {
        if !self.comment.is_empty() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for Footnotes {
    fn write_children<W: Write>(&self, writer: &mut Writer<W>) -> Result<(), SerializeError> {
        for item in &self.footnote {
            item.write_element("w:footnote", writer)?;
        }
        #[cfg(feature = "extra-children")]
        for child in &self.extra_children {
            child.write_to(writer).map_err(SerializeError::from)?;
        }
        Ok(())
    }

    fn is_empty_element(&self) -> bool {
        if !self.footnote.is_empty() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for Endnotes {
    fn write_children<W: Write>(&self, writer: &mut Writer<W>) -> Result<(), SerializeError> {
        for item in &self.endnote {
            item.write_element("w:endnote", writer)?;
        }
        #[cfg(feature = "extra-children")]
        for child in &self.extra_children {
            child.write_to(writer).map_err(SerializeError::from)?;
        }
        Ok(())
    }

    fn is_empty_element(&self) -> bool {
        if !self.endnote.is_empty() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for CTSmartTagType {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        if let Some(ref val) = self.namespaceuri {
            start.push_attribute(("w:namespaceuri", val.as_str()));
        }
        if let Some(ref val) = self.name {
            start.push_attribute(("w:name", val.as_str()));
        }
        if let Some(ref val) = self.url {
            start.push_attribute(("w:url", val.as_str()));
        }
        #[cfg(feature = "extra-attrs")]
        for (key, value) in &self.extra_attrs {
            start.push_attribute((key.as_str(), value.as_str()));
        }
        start
    }

    fn is_empty_element(&self) -> bool {
        true
    }
}

impl ToXml for CTDocPartBehavior {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        {
            let val = &self.value;
            {
                let s = val.to_string();
                start.push_attribute(("w:val", s.as_str()));
            }
        }
        #[cfg(feature = "extra-attrs")]
        for (key, value) in &self.extra_attrs {
            start.push_attribute((key.as_str(), value.as_str()));
        }
        start
    }

    fn is_empty_element(&self) -> bool {
        true
    }
}

impl ToXml for CTDocPartBehaviors {
    fn write_children<W: Write>(&self, writer: &mut Writer<W>) -> Result<(), SerializeError> {
        for item in &self.behavior {
            item.write_element("w:behavior", writer)?;
        }
        #[cfg(feature = "extra-children")]
        for child in &self.extra_children {
            child.write_to(writer).map_err(SerializeError::from)?;
        }
        Ok(())
    }

    fn is_empty_element(&self) -> bool {
        if !self.behavior.is_empty() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for CTDocPartType {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        {
            let val = &self.value;
            {
                let s = val.to_string();
                start.push_attribute(("w:val", s.as_str()));
            }
        }
        #[cfg(feature = "extra-attrs")]
        for (key, value) in &self.extra_attrs {
            start.push_attribute((key.as_str(), value.as_str()));
        }
        start
    }

    fn is_empty_element(&self) -> bool {
        true
    }
}

impl ToXml for CTDocPartTypes {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        if let Some(ref val) = self.all {
            {
                let s = val.to_string();
                start.push_attribute(("w:all", s.as_str()));
            }
        }
        #[cfg(feature = "extra-attrs")]
        for (key, value) in &self.extra_attrs {
            start.push_attribute((key.as_str(), value.as_str()));
        }
        start
    }

    fn write_children<W: Write>(&self, writer: &mut Writer<W>) -> Result<(), SerializeError> {
        for item in &self.r#type {
            item.write_element("w:type", writer)?;
        }
        #[cfg(feature = "extra-children")]
        for child in &self.extra_children {
            child.write_to(writer).map_err(SerializeError::from)?;
        }
        Ok(())
    }

    fn is_empty_element(&self) -> bool {
        if !self.r#type.is_empty() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for CTDocPartGallery {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        {
            let val = &self.value;
            {
                let s = val.to_string();
                start.push_attribute(("w:val", s.as_str()));
            }
        }
        #[cfg(feature = "extra-attrs")]
        for (key, value) in &self.extra_attrs {
            start.push_attribute((key.as_str(), value.as_str()));
        }
        start
    }

    fn is_empty_element(&self) -> bool {
        true
    }
}

impl ToXml for CTDocPartCategory {
    fn write_children<W: Write>(&self, writer: &mut Writer<W>) -> Result<(), SerializeError> {
        {
            let val = &self.name;
            val.write_element("w:name", writer)?;
        }
        {
            let val = &self.gallery;
            val.write_element("w:gallery", writer)?;
        }
        #[cfg(feature = "extra-children")]
        for child in &self.extra_children {
            child.write_to(writer).map_err(SerializeError::from)?;
        }
        Ok(())
    }

    fn is_empty_element(&self) -> bool {
        false
    }
}

impl ToXml for CTDocPartName {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        {
            let val = &self.value;
            start.push_attribute(("w:val", val.as_str()));
        }
        if let Some(ref val) = self.decorated {
            {
                let s = val.to_string();
                start.push_attribute(("w:decorated", s.as_str()));
            }
        }
        #[cfg(feature = "extra-attrs")]
        for (key, value) in &self.extra_attrs {
            start.push_attribute((key.as_str(), value.as_str()));
        }
        start
    }

    fn is_empty_element(&self) -> bool {
        true
    }
}

impl ToXml for CTDocPartPr {
    fn write_children<W: Write>(&self, writer: &mut Writer<W>) -> Result<(), SerializeError> {
        {
            let val = &self.name;
            val.write_element("w:name", writer)?;
        }
        if let Some(ref val) = self.style {
            val.write_element("w:style", writer)?;
        }
        if let Some(ref val) = self.category {
            val.write_element("w:category", writer)?;
        }
        if let Some(ref val) = self.types {
            val.write_element("w:types", writer)?;
        }
        if let Some(ref val) = self.behaviors {
            val.write_element("w:behaviors", writer)?;
        }
        if let Some(ref val) = self.description {
            val.write_element("w:description", writer)?;
        }
        if let Some(ref val) = self.guid {
            val.write_element("w:guid", writer)?;
        }
        #[cfg(feature = "extra-children")]
        for child in &self.extra_children {
            child.write_to(writer).map_err(SerializeError::from)?;
        }
        Ok(())
    }

    fn is_empty_element(&self) -> bool {
        false
    }
}

impl ToXml for CTDocPart {
    fn write_children<W: Write>(&self, writer: &mut Writer<W>) -> Result<(), SerializeError> {
        if let Some(ref val) = self.doc_part_pr {
            val.write_element("w:docPartPr", writer)?;
        }
        if let Some(ref val) = self.doc_part_body {
            val.write_element("w:docPartBody", writer)?;
        }
        #[cfg(feature = "extra-children")]
        for child in &self.extra_children {
            child.write_to(writer).map_err(SerializeError::from)?;
        }
        Ok(())
    }

    fn is_empty_element(&self) -> bool {
        if self.doc_part_pr.is_some() {
            return false;
        }
        if self.doc_part_body.is_some() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for CTDocParts {
    fn write_children<W: Write>(&self, writer: &mut Writer<W>) -> Result<(), SerializeError> {
        for item in &self.doc_part {
            item.write_element("w:docPart", writer)?;
        }
        #[cfg(feature = "extra-children")]
        for child in &self.extra_children {
            child.write_to(writer).map_err(SerializeError::from)?;
        }
        Ok(())
    }

    fn is_empty_element(&self) -> bool {
        if !self.doc_part.is_empty() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for CTCaption {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        {
            let val = &self.name;
            start.push_attribute(("w:name", val.as_str()));
        }
        if let Some(ref val) = self.pos {
            {
                let s = val.to_string();
                start.push_attribute(("w:pos", s.as_str()));
            }
        }
        if let Some(ref val) = self.chap_num {
            {
                let s = val.to_string();
                start.push_attribute(("w:chapNum", s.as_str()));
            }
        }
        if let Some(ref val) = self.heading {
            {
                let s = val.to_string();
                start.push_attribute(("w:heading", s.as_str()));
            }
        }
        if let Some(ref val) = self.no_label {
            {
                let s = val.to_string();
                start.push_attribute(("w:noLabel", s.as_str()));
            }
        }
        if let Some(ref val) = self.num_fmt {
            {
                let s = val.to_string();
                start.push_attribute(("w:numFmt", s.as_str()));
            }
        }
        if let Some(ref val) = self.sep {
            {
                let s = val.to_string();
                start.push_attribute(("w:sep", s.as_str()));
            }
        }
        #[cfg(feature = "extra-attrs")]
        for (key, value) in &self.extra_attrs {
            start.push_attribute((key.as_str(), value.as_str()));
        }
        start
    }

    fn is_empty_element(&self) -> bool {
        true
    }
}

impl ToXml for CTAutoCaption {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        {
            let val = &self.name;
            start.push_attribute(("w:name", val.as_str()));
        }
        {
            let val = &self.caption;
            start.push_attribute(("w:caption", val.as_str()));
        }
        #[cfg(feature = "extra-attrs")]
        for (key, value) in &self.extra_attrs {
            start.push_attribute((key.as_str(), value.as_str()));
        }
        start
    }

    fn is_empty_element(&self) -> bool {
        true
    }
}

impl ToXml for CTAutoCaptions {
    fn write_children<W: Write>(&self, writer: &mut Writer<W>) -> Result<(), SerializeError> {
        for item in &self.auto_caption {
            item.write_element("w:autoCaption", writer)?;
        }
        #[cfg(feature = "extra-children")]
        for child in &self.extra_children {
            child.write_to(writer).map_err(SerializeError::from)?;
        }
        Ok(())
    }

    fn is_empty_element(&self) -> bool {
        if !self.auto_caption.is_empty() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for CTCaptions {
    fn write_children<W: Write>(&self, writer: &mut Writer<W>) -> Result<(), SerializeError> {
        for item in &self.caption {
            item.write_element("w:caption", writer)?;
        }
        if let Some(ref val) = self.auto_captions {
            val.write_element("w:autoCaptions", writer)?;
        }
        #[cfg(feature = "extra-children")]
        for child in &self.extra_children {
            child.write_to(writer).map_err(SerializeError::from)?;
        }
        Ok(())
    }

    fn is_empty_element(&self) -> bool {
        if !self.caption.is_empty() {
            return false;
        }
        if self.auto_captions.is_some() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for CTDocumentBase {
    fn write_children<W: Write>(&self, writer: &mut Writer<W>) -> Result<(), SerializeError> {
        if let Some(ref val) = self.background {
            val.write_element("w:background", writer)?;
        }
        #[cfg(feature = "extra-children")]
        for child in &self.extra_children {
            child.write_to(writer).map_err(SerializeError::from)?;
        }
        Ok(())
    }

    fn is_empty_element(&self) -> bool {
        if self.background.is_some() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for Document {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        if let Some(ref val) = self.conformance {
            {
                let s = val.to_string();
                start.push_attribute(("w:conformance", s.as_str()));
            }
        }
        #[cfg(feature = "extra-attrs")]
        for (key, value) in &self.extra_attrs {
            start.push_attribute((key.as_str(), value.as_str()));
        }
        start
    }

    fn write_children<W: Write>(&self, writer: &mut Writer<W>) -> Result<(), SerializeError> {
        if let Some(ref val) = self.background {
            val.write_element("w:background", writer)?;
        }
        if let Some(ref val) = self.body {
            val.write_element("w:body", writer)?;
        }
        #[cfg(feature = "extra-children")]
        for child in &self.extra_children {
            child.write_to(writer).map_err(SerializeError::from)?;
        }
        Ok(())
    }

    fn is_empty_element(&self) -> bool {
        if self.background.is_some() {
            return false;
        }
        if self.body.is_some() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for CTGlossaryDocument {
    fn write_children<W: Write>(&self, writer: &mut Writer<W>) -> Result<(), SerializeError> {
        if let Some(ref val) = self.background {
            val.write_element("w:background", writer)?;
        }
        if let Some(ref val) = self.doc_parts {
            val.write_element("w:docParts", writer)?;
        }
        #[cfg(feature = "extra-children")]
        for child in &self.extra_children {
            child.write_to(writer).map_err(SerializeError::from)?;
        }
        Ok(())
    }

    fn is_empty_element(&self) -> bool {
        if self.background.is_some() {
            return false;
        }
        if self.doc_parts.is_some() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for WAnyVmlOffice {
    fn write_children<W: Write>(&self, writer: &mut Writer<W>) -> Result<(), SerializeError> {
        #[cfg(feature = "extra-children")]
        for child in &self.extra_children {
            child.write_to(writer).map_err(SerializeError::from)?;
        }
        Ok(())
    }

    fn is_empty_element(&self) -> bool {
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for WAnyVmlVml {
    fn write_children<W: Write>(&self, writer: &mut Writer<W>) -> Result<(), SerializeError> {
        #[cfg(feature = "extra-children")]
        for child in &self.extra_children {
            child.write_to(writer).map_err(SerializeError::from)?;
        }
        Ok(())
    }

    fn is_empty_element(&self) -> bool {
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}
