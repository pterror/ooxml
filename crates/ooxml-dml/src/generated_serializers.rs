// ToXml serializers for generated types.
// Enables roundtrip XML serialization alongside FromXml parsers.

#![allow(unused_variables, unused_assignments, unreachable_code, unused_imports)]
#![allow(clippy::single_match)]
#![allow(clippy::match_single_binding)]
#![allow(clippy::explicit_counter_loop)]

use super::generated::*;
pub use ooxml_xml::{SerializeError, ToXml};
use quick_xml::Writer;
use quick_xml::events::{BytesEnd, BytesStart, BytesText, Event};
use std::io::Write;

#[allow(dead_code)]
/// Encode bytes as a hex string.
fn encode_hex(bytes: &[u8]) -> String {
    bytes.iter().map(|b| format!("{:02X}", b)).collect()
}

#[allow(dead_code)]
/// Encode bytes as a base64 string.
fn encode_base64(bytes: &[u8]) -> String {
    use base64::Engine;
    base64::engine::general_purpose::STANDARD.encode(bytes)
}

impl ToXml for CTAudioFile {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        #[allow(unused_mut)]
        let mut start = start;
        #[cfg(feature = "dml-media")]
        {
            let val = &self.link;
            start.push_attribute(("r:link", val.as_str()));
        }
        #[cfg(feature = "dml-media")]
        if let Some(ref val) = self.content_type {
            start.push_attribute(("contentType", val.as_str()));
        }
        #[cfg(feature = "extra-attrs")]
        for (key, value) in &self.extra_attrs {
            start.push_attribute((key.as_str(), value.as_str()));
        }
        start
    }

    fn write_children<W: Write>(&self, writer: &mut Writer<W>) -> Result<(), SerializeError> {
        #[cfg(feature = "extra-children")]
        let mut extra_iter = self.extra_children.iter().peekable();
        #[cfg(feature = "extra-children")]
        let mut emit_idx: usize = 0;
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        if let Some(ref val) = self.ext_lst {
            val.write_element("a:extLst", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        for extra in extra_iter {
            extra.node.write_to(writer).map_err(SerializeError::from)?;
        }
        Ok(())
    }

    fn is_empty_element(&self) -> bool {
        if self.ext_lst.is_some() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for CTVideoFile {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        #[allow(unused_mut)]
        let mut start = start;
        #[cfg(feature = "dml-media")]
        {
            let val = &self.link;
            start.push_attribute(("r:link", val.as_str()));
        }
        #[cfg(feature = "dml-media")]
        if let Some(ref val) = self.content_type {
            start.push_attribute(("contentType", val.as_str()));
        }
        #[cfg(feature = "extra-attrs")]
        for (key, value) in &self.extra_attrs {
            start.push_attribute((key.as_str(), value.as_str()));
        }
        start
    }

    fn write_children<W: Write>(&self, writer: &mut Writer<W>) -> Result<(), SerializeError> {
        #[cfg(feature = "extra-children")]
        let mut extra_iter = self.extra_children.iter().peekable();
        #[cfg(feature = "extra-children")]
        let mut emit_idx: usize = 0;
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        if let Some(ref val) = self.ext_lst {
            val.write_element("a:extLst", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        for extra in extra_iter {
            extra.node.write_to(writer).map_err(SerializeError::from)?;
        }
        Ok(())
    }

    fn is_empty_element(&self) -> bool {
        if self.ext_lst.is_some() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for CTQuickTimeFile {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        #[allow(unused_mut)]
        let mut start = start;
        {
            let val = &self.link;
            start.push_attribute(("r:link", val.as_str()));
        }
        #[cfg(feature = "extra-attrs")]
        for (key, value) in &self.extra_attrs {
            start.push_attribute((key.as_str(), value.as_str()));
        }
        start
    }

    fn write_children<W: Write>(&self, writer: &mut Writer<W>) -> Result<(), SerializeError> {
        #[cfg(feature = "extra-children")]
        let mut extra_iter = self.extra_children.iter().peekable();
        #[cfg(feature = "extra-children")]
        let mut emit_idx: usize = 0;
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        if let Some(ref val) = self.ext_lst {
            val.write_element("a:extLst", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        for extra in extra_iter {
            extra.node.write_to(writer).map_err(SerializeError::from)?;
        }
        Ok(())
    }

    fn is_empty_element(&self) -> bool {
        if self.ext_lst.is_some() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for CTAudioCDTime {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        #[allow(unused_mut)]
        let mut start = start;
        {
            let val = &self.track;
            {
                let s = val.to_string();
                start.push_attribute(("track", s.as_str()));
            }
        }
        if let Some(ref val) = self.time {
            {
                let s = val.to_string();
                start.push_attribute(("time", s.as_str()));
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

impl ToXml for CTAudioCD {
    fn write_children<W: Write>(&self, writer: &mut Writer<W>) -> Result<(), SerializeError> {
        #[cfg(feature = "extra-children")]
        let mut extra_iter = self.extra_children.iter().peekable();
        #[cfg(feature = "extra-children")]
        let mut emit_idx: usize = 0;
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        {
            let val = &self.st;
            val.write_element("a:st", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        {
            let val = &self.end;
            val.write_element("a:end", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        if let Some(ref val) = self.ext_lst {
            val.write_element("a:extLst", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        for extra in extra_iter {
            extra.node.write_to(writer).map_err(SerializeError::from)?;
        }
        Ok(())
    }

    fn is_empty_element(&self) -> bool {
        false
    }
}

impl ToXml for EGMedia {
    fn write_element<W: Write>(
        &self,
        _tag: &str,
        writer: &mut Writer<W>,
    ) -> Result<(), SerializeError> {
        match self {
            Self::AudioCd(inner) => inner.write_element("a:audioCd", writer)?,
            Self::WavAudioFile(inner) => inner.write_element("a:wavAudioFile", writer)?,
            Self::AudioFile(inner) => inner.write_element("a:audioFile", writer)?,
            Self::VideoFile(inner) => inner.write_element("a:videoFile", writer)?,
            Self::QuickTimeFile(inner) => inner.write_element("a:quickTimeFile", writer)?,
        }
        Ok(())
    }
}

impl ToXml for ColorScheme {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        #[allow(unused_mut)]
        let mut start = start;
        #[cfg(feature = "dml-themes")]
        {
            let val = &self.name;
            start.push_attribute(("name", val.as_str()));
        }
        #[cfg(feature = "extra-attrs")]
        for (key, value) in &self.extra_attrs {
            start.push_attribute((key.as_str(), value.as_str()));
        }
        start
    }

    fn write_children<W: Write>(&self, writer: &mut Writer<W>) -> Result<(), SerializeError> {
        #[cfg(feature = "extra-children")]
        let mut extra_iter = self.extra_children.iter().peekable();
        #[cfg(feature = "extra-children")]
        let mut emit_idx: usize = 0;
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        #[cfg(feature = "dml-colors")]
        {
            let val = &self.dk1;
            val.write_element("a:dk1", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        #[cfg(feature = "dml-colors")]
        {
            let val = &self.lt1;
            val.write_element("a:lt1", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        #[cfg(feature = "dml-colors")]
        {
            let val = &self.dk2;
            val.write_element("a:dk2", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        #[cfg(feature = "dml-colors")]
        {
            let val = &self.lt2;
            val.write_element("a:lt2", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        #[cfg(feature = "dml-colors")]
        {
            let val = &self.accent1;
            val.write_element("a:accent1", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        #[cfg(feature = "dml-colors")]
        {
            let val = &self.accent2;
            val.write_element("a:accent2", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        #[cfg(feature = "dml-colors")]
        {
            let val = &self.accent3;
            val.write_element("a:accent3", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        #[cfg(feature = "dml-colors")]
        {
            let val = &self.accent4;
            val.write_element("a:accent4", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        #[cfg(feature = "dml-colors")]
        {
            let val = &self.accent5;
            val.write_element("a:accent5", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        #[cfg(feature = "dml-colors")]
        {
            let val = &self.accent6;
            val.write_element("a:accent6", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        #[cfg(feature = "dml-colors")]
        {
            let val = &self.hlink;
            val.write_element("a:hlink", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        #[cfg(feature = "dml-colors")]
        {
            let val = &self.fol_hlink;
            val.write_element("a:folHlink", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        #[cfg(feature = "dml-extensions")]
        if let Some(ref val) = self.ext_lst {
            val.write_element("a:extLst", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        for extra in extra_iter {
            extra.node.write_to(writer).map_err(SerializeError::from)?;
        }
        Ok(())
    }

    fn is_empty_element(&self) -> bool {
        #[cfg(feature = "dml-colors")]
        return false;
        #[cfg(feature = "dml-colors")]
        return false;
        #[cfg(feature = "dml-colors")]
        return false;
        #[cfg(feature = "dml-colors")]
        return false;
        #[cfg(feature = "dml-colors")]
        return false;
        #[cfg(feature = "dml-colors")]
        return false;
        #[cfg(feature = "dml-colors")]
        return false;
        #[cfg(feature = "dml-colors")]
        return false;
        #[cfg(feature = "dml-colors")]
        return false;
        #[cfg(feature = "dml-colors")]
        return false;
        #[cfg(feature = "dml-colors")]
        return false;
        #[cfg(feature = "dml-colors")]
        return false;
        #[cfg(feature = "dml-extensions")]
        if self.ext_lst.is_some() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for CTCustomColor {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        #[allow(unused_mut)]
        let mut start = start;
        if let Some(ref val) = self.name {
            start.push_attribute(("name", val.as_str()));
        }
        #[cfg(feature = "extra-attrs")]
        for (key, value) in &self.extra_attrs {
            start.push_attribute((key.as_str(), value.as_str()));
        }
        start
    }

    fn write_children<W: Write>(&self, writer: &mut Writer<W>) -> Result<(), SerializeError> {
        #[cfg(feature = "extra-children")]
        let mut extra_iter = self.extra_children.iter().peekable();
        #[cfg(feature = "extra-children")]
        let mut emit_idx: usize = 0;
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        if let Some(ref val) = self.color_choice {
            val.write_element("", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        for extra in extra_iter {
            extra.node.write_to(writer).map_err(SerializeError::from)?;
        }
        Ok(())
    }

    fn is_empty_element(&self) -> bool {
        false
    }
}

impl ToXml for CTSupplementalFont {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        #[allow(unused_mut)]
        let mut start = start;
        {
            let val = &self.script;
            start.push_attribute(("script", val.as_str()));
        }
        {
            let val = &self.typeface;
            start.push_attribute(("typeface", val.as_str()));
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

impl ToXml for CTCustomColorList {
    fn write_children<W: Write>(&self, writer: &mut Writer<W>) -> Result<(), SerializeError> {
        #[cfg(feature = "extra-children")]
        let mut extra_iter = self.extra_children.iter().peekable();
        #[cfg(feature = "extra-children")]
        let mut emit_idx: usize = 0;
        for item in &self.cust_clr {
            #[cfg(feature = "extra-children")]
            while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
                extra_iter
                    .next()
                    .unwrap()
                    .node
                    .write_to(writer)
                    .map_err(SerializeError::from)?;
            }
            item.write_element("a:custClr", writer)?;
            #[cfg(feature = "extra-children")]
            {
                emit_idx += 1;
            }
        }
        #[cfg(feature = "extra-children")]
        for extra in extra_iter {
            extra.node.write_to(writer).map_err(SerializeError::from)?;
        }
        Ok(())
    }

    fn is_empty_element(&self) -> bool {
        if !self.cust_clr.is_empty() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for CTFontCollection {
    fn write_children<W: Write>(&self, writer: &mut Writer<W>) -> Result<(), SerializeError> {
        #[cfg(feature = "extra-children")]
        let mut extra_iter = self.extra_children.iter().peekable();
        #[cfg(feature = "extra-children")]
        let mut emit_idx: usize = 0;
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        {
            let val = &self.latin;
            val.write_element("a:latin", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        {
            let val = &self.ea;
            val.write_element("a:ea", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        {
            let val = &self.cs;
            val.write_element("a:cs", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        for item in &self.font {
            #[cfg(feature = "extra-children")]
            while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
                extra_iter
                    .next()
                    .unwrap()
                    .node
                    .write_to(writer)
                    .map_err(SerializeError::from)?;
            }
            item.write_element("a:font", writer)?;
            #[cfg(feature = "extra-children")]
            {
                emit_idx += 1;
            }
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        if let Some(ref val) = self.ext_lst {
            val.write_element("a:extLst", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        for extra in extra_iter {
            extra.node.write_to(writer).map_err(SerializeError::from)?;
        }
        Ok(())
    }

    fn is_empty_element(&self) -> bool {
        false
    }
}

impl ToXml for CTEffectStyleItem {
    fn write_children<W: Write>(&self, writer: &mut Writer<W>) -> Result<(), SerializeError> {
        #[cfg(feature = "extra-children")]
        let mut extra_iter = self.extra_children.iter().peekable();
        #[cfg(feature = "extra-children")]
        let mut emit_idx: usize = 0;
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        if let Some(ref val) = self.effect_properties {
            val.write_element("", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        if let Some(ref val) = self.scene3d {
            val.write_element("a:scene3d", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        if let Some(ref val) = self.sp3d {
            val.write_element("a:sp3d", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        for extra in extra_iter {
            extra.node.write_to(writer).map_err(SerializeError::from)?;
        }
        Ok(())
    }

    fn is_empty_element(&self) -> bool {
        false
    }
}

impl ToXml for FontScheme {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        #[allow(unused_mut)]
        let mut start = start;
        #[cfg(feature = "dml-themes")]
        {
            let val = &self.name;
            start.push_attribute(("name", val.as_str()));
        }
        #[cfg(feature = "extra-attrs")]
        for (key, value) in &self.extra_attrs {
            start.push_attribute((key.as_str(), value.as_str()));
        }
        start
    }

    fn write_children<W: Write>(&self, writer: &mut Writer<W>) -> Result<(), SerializeError> {
        #[cfg(feature = "extra-children")]
        let mut extra_iter = self.extra_children.iter().peekable();
        #[cfg(feature = "extra-children")]
        let mut emit_idx: usize = 0;
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        #[cfg(feature = "dml-themes")]
        {
            let val = &self.major_font;
            val.write_element("a:majorFont", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        #[cfg(feature = "dml-themes")]
        {
            let val = &self.minor_font;
            val.write_element("a:minorFont", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        #[cfg(feature = "dml-extensions")]
        if let Some(ref val) = self.ext_lst {
            val.write_element("a:extLst", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        for extra in extra_iter {
            extra.node.write_to(writer).map_err(SerializeError::from)?;
        }
        Ok(())
    }

    fn is_empty_element(&self) -> bool {
        #[cfg(feature = "dml-themes")]
        return false;
        #[cfg(feature = "dml-themes")]
        return false;
        #[cfg(feature = "dml-extensions")]
        if self.ext_lst.is_some() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for CTFillStyleList {
    fn write_children<W: Write>(&self, writer: &mut Writer<W>) -> Result<(), SerializeError> {
        #[cfg(feature = "extra-children")]
        let mut extra_iter = self.extra_children.iter().peekable();
        #[cfg(feature = "extra-children")]
        let mut emit_idx: usize = 0;
        for item in &self.fill_properties {
            #[cfg(feature = "extra-children")]
            while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
                extra_iter
                    .next()
                    .unwrap()
                    .node
                    .write_to(writer)
                    .map_err(SerializeError::from)?;
            }
            item.write_element("", writer)?;
            #[cfg(feature = "extra-children")]
            {
                emit_idx += 1;
            }
        }
        #[cfg(feature = "extra-children")]
        for extra in extra_iter {
            extra.node.write_to(writer).map_err(SerializeError::from)?;
        }
        Ok(())
    }

    fn is_empty_element(&self) -> bool {
        if !self.fill_properties.is_empty() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for CTLineStyleList {
    fn write_children<W: Write>(&self, writer: &mut Writer<W>) -> Result<(), SerializeError> {
        #[cfg(feature = "extra-children")]
        let mut extra_iter = self.extra_children.iter().peekable();
        #[cfg(feature = "extra-children")]
        let mut emit_idx: usize = 0;
        for item in &self.line {
            #[cfg(feature = "extra-children")]
            while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
                extra_iter
                    .next()
                    .unwrap()
                    .node
                    .write_to(writer)
                    .map_err(SerializeError::from)?;
            }
            item.write_element("a:ln", writer)?;
            #[cfg(feature = "extra-children")]
            {
                emit_idx += 1;
            }
        }
        #[cfg(feature = "extra-children")]
        for extra in extra_iter {
            extra.node.write_to(writer).map_err(SerializeError::from)?;
        }
        Ok(())
    }

    fn is_empty_element(&self) -> bool {
        if !self.line.is_empty() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for CTEffectStyleList {
    fn write_children<W: Write>(&self, writer: &mut Writer<W>) -> Result<(), SerializeError> {
        #[cfg(feature = "extra-children")]
        let mut extra_iter = self.extra_children.iter().peekable();
        #[cfg(feature = "extra-children")]
        let mut emit_idx: usize = 0;
        for item in &self.effect_style {
            #[cfg(feature = "extra-children")]
            while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
                extra_iter
                    .next()
                    .unwrap()
                    .node
                    .write_to(writer)
                    .map_err(SerializeError::from)?;
            }
            item.write_element("a:effectStyle", writer)?;
            #[cfg(feature = "extra-children")]
            {
                emit_idx += 1;
            }
        }
        #[cfg(feature = "extra-children")]
        for extra in extra_iter {
            extra.node.write_to(writer).map_err(SerializeError::from)?;
        }
        Ok(())
    }

    fn is_empty_element(&self) -> bool {
        if !self.effect_style.is_empty() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for CTBackgroundFillStyleList {
    fn write_children<W: Write>(&self, writer: &mut Writer<W>) -> Result<(), SerializeError> {
        #[cfg(feature = "extra-children")]
        let mut extra_iter = self.extra_children.iter().peekable();
        #[cfg(feature = "extra-children")]
        let mut emit_idx: usize = 0;
        for item in &self.fill_properties {
            #[cfg(feature = "extra-children")]
            while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
                extra_iter
                    .next()
                    .unwrap()
                    .node
                    .write_to(writer)
                    .map_err(SerializeError::from)?;
            }
            item.write_element("", writer)?;
            #[cfg(feature = "extra-children")]
            {
                emit_idx += 1;
            }
        }
        #[cfg(feature = "extra-children")]
        for extra in extra_iter {
            extra.node.write_to(writer).map_err(SerializeError::from)?;
        }
        Ok(())
    }

    fn is_empty_element(&self) -> bool {
        if !self.fill_properties.is_empty() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for CTStyleMatrix {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        #[allow(unused_mut)]
        let mut start = start;
        #[cfg(feature = "dml-themes")]
        if let Some(ref val) = self.name {
            start.push_attribute(("name", val.as_str()));
        }
        #[cfg(feature = "extra-attrs")]
        for (key, value) in &self.extra_attrs {
            start.push_attribute((key.as_str(), value.as_str()));
        }
        start
    }

    fn write_children<W: Write>(&self, writer: &mut Writer<W>) -> Result<(), SerializeError> {
        #[cfg(feature = "extra-children")]
        let mut extra_iter = self.extra_children.iter().peekable();
        #[cfg(feature = "extra-children")]
        let mut emit_idx: usize = 0;
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        #[cfg(feature = "dml-themes")]
        {
            let val = &self.fill_style_lst;
            val.write_element("a:fillStyleLst", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        #[cfg(feature = "dml-themes")]
        {
            let val = &self.ln_style_lst;
            val.write_element("a:lnStyleLst", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        #[cfg(feature = "dml-themes")]
        {
            let val = &self.effect_style_lst;
            val.write_element("a:effectStyleLst", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        #[cfg(feature = "dml-themes")]
        {
            let val = &self.bg_fill_style_lst;
            val.write_element("a:bgFillStyleLst", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        for extra in extra_iter {
            extra.node.write_to(writer).map_err(SerializeError::from)?;
        }
        Ok(())
    }

    fn is_empty_element(&self) -> bool {
        #[cfg(feature = "dml-themes")]
        return false;
        #[cfg(feature = "dml-themes")]
        return false;
        #[cfg(feature = "dml-themes")]
        return false;
        #[cfg(feature = "dml-themes")]
        return false;
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for CTBaseStyles {
    fn write_children<W: Write>(&self, writer: &mut Writer<W>) -> Result<(), SerializeError> {
        #[cfg(feature = "extra-children")]
        let mut extra_iter = self.extra_children.iter().peekable();
        #[cfg(feature = "extra-children")]
        let mut emit_idx: usize = 0;
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        #[cfg(feature = "dml-colors")]
        {
            let val = &self.clr_scheme;
            val.write_element("a:clrScheme", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        #[cfg(feature = "dml-themes")]
        {
            let val = &self.font_scheme;
            val.write_element("a:fontScheme", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        #[cfg(feature = "dml-themes")]
        {
            let val = &self.fmt_scheme;
            val.write_element("a:fmtScheme", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        #[cfg(feature = "dml-extensions")]
        if let Some(ref val) = self.ext_lst {
            val.write_element("a:extLst", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        for extra in extra_iter {
            extra.node.write_to(writer).map_err(SerializeError::from)?;
        }
        Ok(())
    }

    fn is_empty_element(&self) -> bool {
        #[cfg(feature = "dml-colors")]
        return false;
        #[cfg(feature = "dml-themes")]
        return false;
        #[cfg(feature = "dml-themes")]
        return false;
        #[cfg(feature = "dml-extensions")]
        if self.ext_lst.is_some() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for CTOfficeArtExtension {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        #[allow(unused_mut)]
        let mut start = start;
        {
            let val = &self.uri;
            start.push_attribute(("uri", val.as_str()));
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

impl ToXml for CTAngle {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        #[allow(unused_mut)]
        let mut start = start;
        {
            let val = &self.value;
            {
                let s = val.to_string();
                start.push_attribute(("val", s.as_str()));
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

impl ToXml for CTPositiveFixedAngle {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        #[allow(unused_mut)]
        let mut start = start;
        {
            let val = &self.value;
            {
                let s = val.to_string();
                start.push_attribute(("val", s.as_str()));
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

impl ToXml for CTPercentage {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        #[allow(unused_mut)]
        let mut start = start;
        {
            let val = &self.value;
            {
                let s = val.to_string();
                start.push_attribute(("val", s.as_str()));
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

impl ToXml for CTPositivePercentage {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        #[allow(unused_mut)]
        let mut start = start;
        {
            let val = &self.value;
            {
                let s = val.to_string();
                start.push_attribute(("val", s.as_str()));
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

impl ToXml for CTFixedPercentage {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        #[allow(unused_mut)]
        let mut start = start;
        {
            let val = &self.value;
            {
                let s = val.to_string();
                start.push_attribute(("val", s.as_str()));
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

impl ToXml for CTPositiveFixedPercentage {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        #[allow(unused_mut)]
        let mut start = start;
        {
            let val = &self.value;
            {
                let s = val.to_string();
                start.push_attribute(("val", s.as_str()));
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

impl ToXml for CTRatio {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        #[allow(unused_mut)]
        let mut start = start;
        {
            let val = &self.n;
            {
                let s = val.to_string();
                start.push_attribute(("n", s.as_str()));
            }
        }
        {
            let val = &self.d;
            {
                let s = val.to_string();
                start.push_attribute(("d", s.as_str()));
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

impl ToXml for Point2D {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        #[allow(unused_mut)]
        let mut start = start;
        {
            let val = &self.x;
            {
                let s = val.to_string();
                start.push_attribute(("x", s.as_str()));
            }
        }
        {
            let val = &self.y;
            {
                let s = val.to_string();
                start.push_attribute(("y", s.as_str()));
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

impl ToXml for PositiveSize2D {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        #[allow(unused_mut)]
        let mut start = start;
        {
            let val = &self.cx;
            {
                let s = val.to_string();
                start.push_attribute(("cx", s.as_str()));
            }
        }
        {
            let val = &self.cy;
            {
                let s = val.to_string();
                start.push_attribute(("cy", s.as_str()));
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

impl ToXml for CTComplementTransform {
    fn is_empty_element(&self) -> bool {
        true
    }
}

impl ToXml for CTInverseTransform {
    fn is_empty_element(&self) -> bool {
        true
    }
}

impl ToXml for CTGrayscaleTransform {
    fn is_empty_element(&self) -> bool {
        true
    }
}

impl ToXml for CTGammaTransform {
    fn is_empty_element(&self) -> bool {
        true
    }
}

impl ToXml for CTInverseGammaTransform {
    fn is_empty_element(&self) -> bool {
        true
    }
}

impl ToXml for EGColorTransform {
    fn write_element<W: Write>(
        &self,
        _tag: &str,
        writer: &mut Writer<W>,
    ) -> Result<(), SerializeError> {
        match self {
            Self::Tint(inner) => inner.write_element("a:tint", writer)?,
            Self::Shade(inner) => inner.write_element("a:shade", writer)?,
            Self::Comp(inner) => inner.write_element("a:comp", writer)?,
            Self::Inv(inner) => inner.write_element("a:inv", writer)?,
            Self::Gray(inner) => inner.write_element("a:gray", writer)?,
            Self::Alpha(inner) => inner.write_element("a:alpha", writer)?,
            Self::AlphaOff(inner) => inner.write_element("a:alphaOff", writer)?,
            Self::AlphaMod(inner) => inner.write_element("a:alphaMod", writer)?,
            Self::Hue(inner) => inner.write_element("a:hue", writer)?,
            Self::HueOff(inner) => inner.write_element("a:hueOff", writer)?,
            Self::HueMod(inner) => inner.write_element("a:hueMod", writer)?,
            Self::Sat(inner) => inner.write_element("a:sat", writer)?,
            Self::SatOff(inner) => inner.write_element("a:satOff", writer)?,
            Self::SatMod(inner) => inner.write_element("a:satMod", writer)?,
            Self::Lum(inner) => inner.write_element("a:lum", writer)?,
            Self::LumOff(inner) => inner.write_element("a:lumOff", writer)?,
            Self::LumMod(inner) => inner.write_element("a:lumMod", writer)?,
            Self::Red(inner) => inner.write_element("a:red", writer)?,
            Self::RedOff(inner) => inner.write_element("a:redOff", writer)?,
            Self::RedMod(inner) => inner.write_element("a:redMod", writer)?,
            Self::Green(inner) => inner.write_element("a:green", writer)?,
            Self::GreenOff(inner) => inner.write_element("a:greenOff", writer)?,
            Self::GreenMod(inner) => inner.write_element("a:greenMod", writer)?,
            Self::Blue(inner) => inner.write_element("a:blue", writer)?,
            Self::BlueOff(inner) => inner.write_element("a:blueOff", writer)?,
            Self::BlueMod(inner) => inner.write_element("a:blueMod", writer)?,
            Self::Gamma(inner) => inner.write_element("a:gamma", writer)?,
            Self::InvGamma(inner) => inner.write_element("a:invGamma", writer)?,
        }
        Ok(())
    }
}

impl ToXml for CTScRgbColor {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        #[allow(unused_mut)]
        let mut start = start;
        {
            let val = &self.relationship_id;
            {
                let s = val.to_string();
                start.push_attribute(("r", s.as_str()));
            }
        }
        {
            let val = &self.g;
            {
                let s = val.to_string();
                start.push_attribute(("g", s.as_str()));
            }
        }
        {
            let val = &self.b;
            {
                let s = val.to_string();
                start.push_attribute(("b", s.as_str()));
            }
        }
        #[cfg(feature = "extra-attrs")]
        for (key, value) in &self.extra_attrs {
            start.push_attribute((key.as_str(), value.as_str()));
        }
        start
    }

    fn write_children<W: Write>(&self, writer: &mut Writer<W>) -> Result<(), SerializeError> {
        #[cfg(feature = "extra-children")]
        let mut extra_iter = self.extra_children.iter().peekable();
        #[cfg(feature = "extra-children")]
        let mut emit_idx: usize = 0;
        for item in &self.color_transform {
            #[cfg(feature = "extra-children")]
            while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
                extra_iter
                    .next()
                    .unwrap()
                    .node
                    .write_to(writer)
                    .map_err(SerializeError::from)?;
            }
            item.write_element("", writer)?;
            #[cfg(feature = "extra-children")]
            {
                emit_idx += 1;
            }
        }
        #[cfg(feature = "extra-children")]
        for extra in extra_iter {
            extra.node.write_to(writer).map_err(SerializeError::from)?;
        }
        Ok(())
    }

    fn is_empty_element(&self) -> bool {
        if !self.color_transform.is_empty() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for SrgbColor {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        #[allow(unused_mut)]
        let mut start = start;
        #[cfg(feature = "dml-colors")]
        {
            let val = &self.value;
            {
                let hex = encode_hex(val);
                start.push_attribute(("val", hex.as_str()));
            }
        }
        #[cfg(feature = "extra-attrs")]
        for (key, value) in &self.extra_attrs {
            start.push_attribute((key.as_str(), value.as_str()));
        }
        start
    }

    fn write_children<W: Write>(&self, writer: &mut Writer<W>) -> Result<(), SerializeError> {
        #[cfg(feature = "extra-children")]
        let mut extra_iter = self.extra_children.iter().peekable();
        #[cfg(feature = "extra-children")]
        let mut emit_idx: usize = 0;
        for item in &self.color_transform {
            #[cfg(feature = "extra-children")]
            while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
                extra_iter
                    .next()
                    .unwrap()
                    .node
                    .write_to(writer)
                    .map_err(SerializeError::from)?;
            }
            item.write_element("", writer)?;
            #[cfg(feature = "extra-children")]
            {
                emit_idx += 1;
            }
        }
        #[cfg(feature = "extra-children")]
        for extra in extra_iter {
            extra.node.write_to(writer).map_err(SerializeError::from)?;
        }
        Ok(())
    }

    fn is_empty_element(&self) -> bool {
        if !self.color_transform.is_empty() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for HslColor {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        #[allow(unused_mut)]
        let mut start = start;
        #[cfg(feature = "dml-colors")]
        {
            let val = &self.hue;
            {
                let s = val.to_string();
                start.push_attribute(("hue", s.as_str()));
            }
        }
        #[cfg(feature = "dml-colors")]
        {
            let val = &self.sat;
            {
                let s = val.to_string();
                start.push_attribute(("sat", s.as_str()));
            }
        }
        #[cfg(feature = "dml-colors")]
        {
            let val = &self.lum;
            {
                let s = val.to_string();
                start.push_attribute(("lum", s.as_str()));
            }
        }
        #[cfg(feature = "extra-attrs")]
        for (key, value) in &self.extra_attrs {
            start.push_attribute((key.as_str(), value.as_str()));
        }
        start
    }

    fn write_children<W: Write>(&self, writer: &mut Writer<W>) -> Result<(), SerializeError> {
        #[cfg(feature = "extra-children")]
        let mut extra_iter = self.extra_children.iter().peekable();
        #[cfg(feature = "extra-children")]
        let mut emit_idx: usize = 0;
        for item in &self.color_transform {
            #[cfg(feature = "extra-children")]
            while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
                extra_iter
                    .next()
                    .unwrap()
                    .node
                    .write_to(writer)
                    .map_err(SerializeError::from)?;
            }
            item.write_element("", writer)?;
            #[cfg(feature = "extra-children")]
            {
                emit_idx += 1;
            }
        }
        #[cfg(feature = "extra-children")]
        for extra in extra_iter {
            extra.node.write_to(writer).map_err(SerializeError::from)?;
        }
        Ok(())
    }

    fn is_empty_element(&self) -> bool {
        if !self.color_transform.is_empty() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for SystemColor {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        #[allow(unused_mut)]
        let mut start = start;
        {
            let val = &self.value;
            {
                let s = val.to_string();
                start.push_attribute(("val", s.as_str()));
            }
        }
        if let Some(ref val) = self.last_clr {
            {
                let hex = encode_hex(val);
                start.push_attribute(("lastClr", hex.as_str()));
            }
        }
        #[cfg(feature = "extra-attrs")]
        for (key, value) in &self.extra_attrs {
            start.push_attribute((key.as_str(), value.as_str()));
        }
        start
    }

    fn write_children<W: Write>(&self, writer: &mut Writer<W>) -> Result<(), SerializeError> {
        #[cfg(feature = "extra-children")]
        let mut extra_iter = self.extra_children.iter().peekable();
        #[cfg(feature = "extra-children")]
        let mut emit_idx: usize = 0;
        for item in &self.color_transform {
            #[cfg(feature = "extra-children")]
            while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
                extra_iter
                    .next()
                    .unwrap()
                    .node
                    .write_to(writer)
                    .map_err(SerializeError::from)?;
            }
            item.write_element("", writer)?;
            #[cfg(feature = "extra-children")]
            {
                emit_idx += 1;
            }
        }
        #[cfg(feature = "extra-children")]
        for extra in extra_iter {
            extra.node.write_to(writer).map_err(SerializeError::from)?;
        }
        Ok(())
    }

    fn is_empty_element(&self) -> bool {
        if !self.color_transform.is_empty() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for SchemeColor {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        #[allow(unused_mut)]
        let mut start = start;
        #[cfg(feature = "dml-colors")]
        {
            let val = &self.value;
            {
                let s = val.to_string();
                start.push_attribute(("val", s.as_str()));
            }
        }
        #[cfg(feature = "extra-attrs")]
        for (key, value) in &self.extra_attrs {
            start.push_attribute((key.as_str(), value.as_str()));
        }
        start
    }

    fn write_children<W: Write>(&self, writer: &mut Writer<W>) -> Result<(), SerializeError> {
        #[cfg(feature = "extra-children")]
        let mut extra_iter = self.extra_children.iter().peekable();
        #[cfg(feature = "extra-children")]
        let mut emit_idx: usize = 0;
        for item in &self.color_transform {
            #[cfg(feature = "extra-children")]
            while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
                extra_iter
                    .next()
                    .unwrap()
                    .node
                    .write_to(writer)
                    .map_err(SerializeError::from)?;
            }
            item.write_element("", writer)?;
            #[cfg(feature = "extra-children")]
            {
                emit_idx += 1;
            }
        }
        #[cfg(feature = "extra-children")]
        for extra in extra_iter {
            extra.node.write_to(writer).map_err(SerializeError::from)?;
        }
        Ok(())
    }

    fn is_empty_element(&self) -> bool {
        if !self.color_transform.is_empty() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for PresetColor {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        #[allow(unused_mut)]
        let mut start = start;
        {
            let val = &self.value;
            {
                let s = val.to_string();
                start.push_attribute(("val", s.as_str()));
            }
        }
        #[cfg(feature = "extra-attrs")]
        for (key, value) in &self.extra_attrs {
            start.push_attribute((key.as_str(), value.as_str()));
        }
        start
    }

    fn write_children<W: Write>(&self, writer: &mut Writer<W>) -> Result<(), SerializeError> {
        #[cfg(feature = "extra-children")]
        let mut extra_iter = self.extra_children.iter().peekable();
        #[cfg(feature = "extra-children")]
        let mut emit_idx: usize = 0;
        for item in &self.color_transform {
            #[cfg(feature = "extra-children")]
            while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
                extra_iter
                    .next()
                    .unwrap()
                    .node
                    .write_to(writer)
                    .map_err(SerializeError::from)?;
            }
            item.write_element("", writer)?;
            #[cfg(feature = "extra-children")]
            {
                emit_idx += 1;
            }
        }
        #[cfg(feature = "extra-children")]
        for extra in extra_iter {
            extra.node.write_to(writer).map_err(SerializeError::from)?;
        }
        Ok(())
    }

    fn is_empty_element(&self) -> bool {
        if !self.color_transform.is_empty() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for EGOfficeArtExtensionList {
    fn write_children<W: Write>(&self, writer: &mut Writer<W>) -> Result<(), SerializeError> {
        #[cfg(feature = "extra-children")]
        let mut extra_iter = self.extra_children.iter().peekable();
        #[cfg(feature = "extra-children")]
        let mut emit_idx: usize = 0;
        for item in &self.extents {
            #[cfg(feature = "extra-children")]
            while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
                extra_iter
                    .next()
                    .unwrap()
                    .node
                    .write_to(writer)
                    .map_err(SerializeError::from)?;
            }
            item.write_element("a:ext", writer)?;
            #[cfg(feature = "extra-children")]
            {
                emit_idx += 1;
            }
        }
        #[cfg(feature = "extra-children")]
        for extra in extra_iter {
            extra.node.write_to(writer).map_err(SerializeError::from)?;
        }
        Ok(())
    }

    fn is_empty_element(&self) -> bool {
        if !self.extents.is_empty() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for CTOfficeArtExtensionList {
    fn write_children<W: Write>(&self, writer: &mut Writer<W>) -> Result<(), SerializeError> {
        #[cfg(feature = "extra-children")]
        let mut extra_iter = self.extra_children.iter().peekable();
        #[cfg(feature = "extra-children")]
        let mut emit_idx: usize = 0;
        for item in &self.extents {
            #[cfg(feature = "extra-children")]
            while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
                extra_iter
                    .next()
                    .unwrap()
                    .node
                    .write_to(writer)
                    .map_err(SerializeError::from)?;
            }
            item.write_element("a:ext", writer)?;
            #[cfg(feature = "extra-children")]
            {
                emit_idx += 1;
            }
        }
        #[cfg(feature = "extra-children")]
        for extra in extra_iter {
            extra.node.write_to(writer).map_err(SerializeError::from)?;
        }
        Ok(())
    }

    fn is_empty_element(&self) -> bool {
        if !self.extents.is_empty() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for CTScale2D {
    fn write_children<W: Write>(&self, writer: &mut Writer<W>) -> Result<(), SerializeError> {
        #[cfg(feature = "extra-children")]
        let mut extra_iter = self.extra_children.iter().peekable();
        #[cfg(feature = "extra-children")]
        let mut emit_idx: usize = 0;
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        {
            let val = &self.sx;
            val.write_element("a:sx", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        {
            let val = &self.sy;
            val.write_element("a:sy", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        for extra in extra_iter {
            extra.node.write_to(writer).map_err(SerializeError::from)?;
        }
        Ok(())
    }

    fn is_empty_element(&self) -> bool {
        false
    }
}

impl ToXml for Transform2D {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        #[allow(unused_mut)]
        let mut start = start;
        if let Some(ref val) = self.rot {
            {
                let s = val.to_string();
                start.push_attribute(("rot", s.as_str()));
            }
        }
        if let Some(ref val) = self.flip_h {
            start.push_attribute(("flipH", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.flip_v {
            start.push_attribute(("flipV", if *val { "1" } else { "0" }));
        }
        #[cfg(feature = "extra-attrs")]
        for (key, value) in &self.extra_attrs {
            start.push_attribute((key.as_str(), value.as_str()));
        }
        start
    }

    fn write_children<W: Write>(&self, writer: &mut Writer<W>) -> Result<(), SerializeError> {
        #[cfg(feature = "extra-children")]
        let mut extra_iter = self.extra_children.iter().peekable();
        #[cfg(feature = "extra-children")]
        let mut emit_idx: usize = 0;
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        if let Some(ref val) = self.offset {
            val.write_element("a:off", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        if let Some(ref val) = self.extents {
            val.write_element("a:ext", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        for extra in extra_iter {
            extra.node.write_to(writer).map_err(SerializeError::from)?;
        }
        Ok(())
    }

    fn is_empty_element(&self) -> bool {
        if self.offset.is_some() {
            return false;
        }
        if self.extents.is_some() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for CTGroupTransform2D {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        #[allow(unused_mut)]
        let mut start = start;
        if let Some(ref val) = self.rot {
            {
                let s = val.to_string();
                start.push_attribute(("rot", s.as_str()));
            }
        }
        if let Some(ref val) = self.flip_h {
            start.push_attribute(("flipH", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.flip_v {
            start.push_attribute(("flipV", if *val { "1" } else { "0" }));
        }
        #[cfg(feature = "extra-attrs")]
        for (key, value) in &self.extra_attrs {
            start.push_attribute((key.as_str(), value.as_str()));
        }
        start
    }

    fn write_children<W: Write>(&self, writer: &mut Writer<W>) -> Result<(), SerializeError> {
        #[cfg(feature = "extra-children")]
        let mut extra_iter = self.extra_children.iter().peekable();
        #[cfg(feature = "extra-children")]
        let mut emit_idx: usize = 0;
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        if let Some(ref val) = self.offset {
            val.write_element("a:off", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        if let Some(ref val) = self.extents {
            val.write_element("a:ext", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        if let Some(ref val) = self.child_offset {
            val.write_element("a:chOff", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        if let Some(ref val) = self.child_extents {
            val.write_element("a:chExt", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        for extra in extra_iter {
            extra.node.write_to(writer).map_err(SerializeError::from)?;
        }
        Ok(())
    }

    fn is_empty_element(&self) -> bool {
        if self.offset.is_some() {
            return false;
        }
        if self.extents.is_some() {
            return false;
        }
        if self.child_offset.is_some() {
            return false;
        }
        if self.child_extents.is_some() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for CTPoint3D {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        #[allow(unused_mut)]
        let mut start = start;
        {
            let val = &self.x;
            {
                let s = val.to_string();
                start.push_attribute(("x", s.as_str()));
            }
        }
        {
            let val = &self.y;
            {
                let s = val.to_string();
                start.push_attribute(("y", s.as_str()));
            }
        }
        {
            let val = &self.z;
            {
                let s = val.to_string();
                start.push_attribute(("z", s.as_str()));
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

impl ToXml for CTVector3D {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        #[allow(unused_mut)]
        let mut start = start;
        {
            let val = &self.dx;
            {
                let s = val.to_string();
                start.push_attribute(("dx", s.as_str()));
            }
        }
        {
            let val = &self.dy;
            {
                let s = val.to_string();
                start.push_attribute(("dy", s.as_str()));
            }
        }
        {
            let val = &self.dz;
            {
                let s = val.to_string();
                start.push_attribute(("dz", s.as_str()));
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

impl ToXml for CTSphereCoords {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        #[allow(unused_mut)]
        let mut start = start;
        {
            let val = &self.lat;
            {
                let s = val.to_string();
                start.push_attribute(("lat", s.as_str()));
            }
        }
        {
            let val = &self.lon;
            {
                let s = val.to_string();
                start.push_attribute(("lon", s.as_str()));
            }
        }
        {
            let val = &self.rev;
            {
                let s = val.to_string();
                start.push_attribute(("rev", s.as_str()));
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

impl ToXml for CTRelativeRect {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        #[allow(unused_mut)]
        let mut start = start;
        if let Some(ref val) = self.l {
            {
                let s = val.to_string();
                start.push_attribute(("l", s.as_str()));
            }
        }
        if let Some(ref val) = self.t {
            {
                let s = val.to_string();
                start.push_attribute(("t", s.as_str()));
            }
        }
        if let Some(ref val) = self.relationship_id {
            {
                let s = val.to_string();
                start.push_attribute(("r", s.as_str()));
            }
        }
        if let Some(ref val) = self.b {
            {
                let s = val.to_string();
                start.push_attribute(("b", s.as_str()));
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

impl ToXml for EGColorChoice {
    fn write_element<W: Write>(
        &self,
        _tag: &str,
        writer: &mut Writer<W>,
    ) -> Result<(), SerializeError> {
        match self {
            Self::ScrgbClr(inner) => inner.write_element("a:scrgbClr", writer)?,
            Self::SrgbClr(inner) => inner.write_element("a:srgbClr", writer)?,
            Self::HslClr(inner) => inner.write_element("a:hslClr", writer)?,
            Self::SysClr(inner) => inner.write_element("a:sysClr", writer)?,
            Self::SchemeClr(inner) => inner.write_element("a:schemeClr", writer)?,
            Self::PrstClr(inner) => inner.write_element("a:prstClr", writer)?,
        }
        Ok(())
    }
}

impl ToXml for CTColor {
    fn write_children<W: Write>(&self, writer: &mut Writer<W>) -> Result<(), SerializeError> {
        #[cfg(feature = "extra-children")]
        let mut extra_iter = self.extra_children.iter().peekable();
        #[cfg(feature = "extra-children")]
        let mut emit_idx: usize = 0;
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        if let Some(ref val) = self.color_choice {
            val.write_element("", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        for extra in extra_iter {
            extra.node.write_to(writer).map_err(SerializeError::from)?;
        }
        Ok(())
    }

    fn is_empty_element(&self) -> bool {
        false
    }
}

impl ToXml for CTColorMRU {
    fn write_children<W: Write>(&self, writer: &mut Writer<W>) -> Result<(), SerializeError> {
        #[cfg(feature = "extra-children")]
        let mut extra_iter = self.extra_children.iter().peekable();
        #[cfg(feature = "extra-children")]
        let mut emit_idx: usize = 0;
        for item in &self.color_choice {
            #[cfg(feature = "extra-children")]
            while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
                extra_iter
                    .next()
                    .unwrap()
                    .node
                    .write_to(writer)
                    .map_err(SerializeError::from)?;
            }
            item.write_element("", writer)?;
            #[cfg(feature = "extra-children")]
            {
                emit_idx += 1;
            }
        }
        #[cfg(feature = "extra-children")]
        for extra in extra_iter {
            extra.node.write_to(writer).map_err(SerializeError::from)?;
        }
        Ok(())
    }

    fn is_empty_element(&self) -> bool {
        if !self.color_choice.is_empty() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for AAGBlob {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        #[allow(unused_mut)]
        let mut start = start;
        if let Some(ref val) = self.embed {
            start.push_attribute(("r:embed", val.as_str()));
        }
        if let Some(ref val) = self.link {
            start.push_attribute(("r:link", val.as_str()));
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

impl ToXml for CTEmbeddedWAVAudioFile {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        #[allow(unused_mut)]
        let mut start = start;
        {
            let val = &self.embed;
            start.push_attribute(("r:embed", val.as_str()));
        }
        if let Some(ref val) = self.name {
            start.push_attribute(("name", val.as_str()));
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

impl ToXml for CTHyperlink {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        #[allow(unused_mut)]
        let mut start = start;
        #[cfg(feature = "dml-text")]
        if let Some(ref val) = self.id {
            start.push_attribute(("r:id", val.as_str()));
        }
        #[cfg(feature = "dml-text")]
        if let Some(ref val) = self.invalid_url {
            start.push_attribute(("invalidUrl", val.as_str()));
        }
        #[cfg(feature = "dml-text")]
        if let Some(ref val) = self.action {
            start.push_attribute(("action", val.as_str()));
        }
        #[cfg(feature = "dml-text")]
        if let Some(ref val) = self.tgt_frame {
            start.push_attribute(("tgtFrame", val.as_str()));
        }
        #[cfg(feature = "dml-text")]
        if let Some(ref val) = self.tooltip {
            start.push_attribute(("tooltip", val.as_str()));
        }
        #[cfg(feature = "dml-text")]
        if let Some(ref val) = self.history {
            start.push_attribute(("history", if *val { "1" } else { "0" }));
        }
        #[cfg(feature = "dml-text")]
        if let Some(ref val) = self.highlight_click {
            start.push_attribute(("highlightClick", if *val { "1" } else { "0" }));
        }
        #[cfg(feature = "dml-text")]
        if let Some(ref val) = self.end_snd {
            start.push_attribute(("endSnd", if *val { "1" } else { "0" }));
        }
        #[cfg(feature = "extra-attrs")]
        for (key, value) in &self.extra_attrs {
            start.push_attribute((key.as_str(), value.as_str()));
        }
        start
    }

    fn write_children<W: Write>(&self, writer: &mut Writer<W>) -> Result<(), SerializeError> {
        #[cfg(feature = "extra-children")]
        let mut extra_iter = self.extra_children.iter().peekable();
        #[cfg(feature = "extra-children")]
        let mut emit_idx: usize = 0;
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        #[cfg(feature = "dml-text")]
        if let Some(ref val) = self.snd {
            val.write_element("a:snd", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        #[cfg(feature = "dml-extensions")]
        if let Some(ref val) = self.ext_lst {
            val.write_element("a:extLst", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        for extra in extra_iter {
            extra.node.write_to(writer).map_err(SerializeError::from)?;
        }
        Ok(())
    }

    fn is_empty_element(&self) -> bool {
        #[cfg(feature = "dml-text")]
        if self.snd.is_some() {
            return false;
        }
        #[cfg(feature = "dml-extensions")]
        if self.ext_lst.is_some() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for AAGLocking {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        #[allow(unused_mut)]
        let mut start = start;
        if let Some(ref val) = self.no_grp {
            start.push_attribute(("noGrp", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.no_select {
            start.push_attribute(("noSelect", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.no_rot {
            start.push_attribute(("noRot", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.no_change_aspect {
            start.push_attribute(("noChangeAspect", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.no_move {
            start.push_attribute(("noMove", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.no_resize {
            start.push_attribute(("noResize", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.no_edit_points {
            start.push_attribute(("noEditPoints", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.no_adjust_handles {
            start.push_attribute(("noAdjustHandles", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.no_change_arrowheads {
            start.push_attribute(("noChangeArrowheads", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.no_change_shape_type {
            start.push_attribute(("noChangeShapeType", if *val { "1" } else { "0" }));
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

impl ToXml for CTConnectorLocking {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        #[allow(unused_mut)]
        let mut start = start;
        if let Some(ref val) = self.no_grp {
            start.push_attribute(("noGrp", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.no_select {
            start.push_attribute(("noSelect", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.no_rot {
            start.push_attribute(("noRot", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.no_change_aspect {
            start.push_attribute(("noChangeAspect", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.no_move {
            start.push_attribute(("noMove", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.no_resize {
            start.push_attribute(("noResize", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.no_edit_points {
            start.push_attribute(("noEditPoints", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.no_adjust_handles {
            start.push_attribute(("noAdjustHandles", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.no_change_arrowheads {
            start.push_attribute(("noChangeArrowheads", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.no_change_shape_type {
            start.push_attribute(("noChangeShapeType", if *val { "1" } else { "0" }));
        }
        #[cfg(feature = "extra-attrs")]
        for (key, value) in &self.extra_attrs {
            start.push_attribute((key.as_str(), value.as_str()));
        }
        start
    }

    fn write_children<W: Write>(&self, writer: &mut Writer<W>) -> Result<(), SerializeError> {
        #[cfg(feature = "extra-children")]
        let mut extra_iter = self.extra_children.iter().peekable();
        #[cfg(feature = "extra-children")]
        let mut emit_idx: usize = 0;
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        if let Some(ref val) = self.ext_lst {
            val.write_element("a:extLst", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        for extra in extra_iter {
            extra.node.write_to(writer).map_err(SerializeError::from)?;
        }
        Ok(())
    }

    fn is_empty_element(&self) -> bool {
        if self.ext_lst.is_some() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for CTShapeLocking {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        #[allow(unused_mut)]
        let mut start = start;
        if let Some(ref val) = self.no_grp {
            start.push_attribute(("noGrp", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.no_select {
            start.push_attribute(("noSelect", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.no_rot {
            start.push_attribute(("noRot", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.no_change_aspect {
            start.push_attribute(("noChangeAspect", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.no_move {
            start.push_attribute(("noMove", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.no_resize {
            start.push_attribute(("noResize", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.no_edit_points {
            start.push_attribute(("noEditPoints", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.no_adjust_handles {
            start.push_attribute(("noAdjustHandles", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.no_change_arrowheads {
            start.push_attribute(("noChangeArrowheads", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.no_change_shape_type {
            start.push_attribute(("noChangeShapeType", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.no_text_edit {
            start.push_attribute(("noTextEdit", if *val { "1" } else { "0" }));
        }
        #[cfg(feature = "extra-attrs")]
        for (key, value) in &self.extra_attrs {
            start.push_attribute((key.as_str(), value.as_str()));
        }
        start
    }

    fn write_children<W: Write>(&self, writer: &mut Writer<W>) -> Result<(), SerializeError> {
        #[cfg(feature = "extra-children")]
        let mut extra_iter = self.extra_children.iter().peekable();
        #[cfg(feature = "extra-children")]
        let mut emit_idx: usize = 0;
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        if let Some(ref val) = self.ext_lst {
            val.write_element("a:extLst", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        for extra in extra_iter {
            extra.node.write_to(writer).map_err(SerializeError::from)?;
        }
        Ok(())
    }

    fn is_empty_element(&self) -> bool {
        if self.ext_lst.is_some() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for CTPictureLocking {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        #[allow(unused_mut)]
        let mut start = start;
        if let Some(ref val) = self.no_grp {
            start.push_attribute(("noGrp", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.no_select {
            start.push_attribute(("noSelect", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.no_rot {
            start.push_attribute(("noRot", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.no_change_aspect {
            start.push_attribute(("noChangeAspect", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.no_move {
            start.push_attribute(("noMove", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.no_resize {
            start.push_attribute(("noResize", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.no_edit_points {
            start.push_attribute(("noEditPoints", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.no_adjust_handles {
            start.push_attribute(("noAdjustHandles", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.no_change_arrowheads {
            start.push_attribute(("noChangeArrowheads", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.no_change_shape_type {
            start.push_attribute(("noChangeShapeType", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.no_crop {
            start.push_attribute(("noCrop", if *val { "1" } else { "0" }));
        }
        #[cfg(feature = "extra-attrs")]
        for (key, value) in &self.extra_attrs {
            start.push_attribute((key.as_str(), value.as_str()));
        }
        start
    }

    fn write_children<W: Write>(&self, writer: &mut Writer<W>) -> Result<(), SerializeError> {
        #[cfg(feature = "extra-children")]
        let mut extra_iter = self.extra_children.iter().peekable();
        #[cfg(feature = "extra-children")]
        let mut emit_idx: usize = 0;
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        if let Some(ref val) = self.ext_lst {
            val.write_element("a:extLst", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        for extra in extra_iter {
            extra.node.write_to(writer).map_err(SerializeError::from)?;
        }
        Ok(())
    }

    fn is_empty_element(&self) -> bool {
        if self.ext_lst.is_some() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for CTGroupLocking {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        #[allow(unused_mut)]
        let mut start = start;
        if let Some(ref val) = self.no_grp {
            start.push_attribute(("noGrp", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.no_ungrp {
            start.push_attribute(("noUngrp", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.no_select {
            start.push_attribute(("noSelect", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.no_rot {
            start.push_attribute(("noRot", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.no_change_aspect {
            start.push_attribute(("noChangeAspect", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.no_move {
            start.push_attribute(("noMove", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.no_resize {
            start.push_attribute(("noResize", if *val { "1" } else { "0" }));
        }
        #[cfg(feature = "extra-attrs")]
        for (key, value) in &self.extra_attrs {
            start.push_attribute((key.as_str(), value.as_str()));
        }
        start
    }

    fn write_children<W: Write>(&self, writer: &mut Writer<W>) -> Result<(), SerializeError> {
        #[cfg(feature = "extra-children")]
        let mut extra_iter = self.extra_children.iter().peekable();
        #[cfg(feature = "extra-children")]
        let mut emit_idx: usize = 0;
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        if let Some(ref val) = self.ext_lst {
            val.write_element("a:extLst", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        for extra in extra_iter {
            extra.node.write_to(writer).map_err(SerializeError::from)?;
        }
        Ok(())
    }

    fn is_empty_element(&self) -> bool {
        if self.ext_lst.is_some() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for CTGraphicalObjectFrameLocking {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        #[allow(unused_mut)]
        let mut start = start;
        if let Some(ref val) = self.no_grp {
            start.push_attribute(("noGrp", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.no_drilldown {
            start.push_attribute(("noDrilldown", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.no_select {
            start.push_attribute(("noSelect", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.no_change_aspect {
            start.push_attribute(("noChangeAspect", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.no_move {
            start.push_attribute(("noMove", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.no_resize {
            start.push_attribute(("noResize", if *val { "1" } else { "0" }));
        }
        #[cfg(feature = "extra-attrs")]
        for (key, value) in &self.extra_attrs {
            start.push_attribute((key.as_str(), value.as_str()));
        }
        start
    }

    fn write_children<W: Write>(&self, writer: &mut Writer<W>) -> Result<(), SerializeError> {
        #[cfg(feature = "extra-children")]
        let mut extra_iter = self.extra_children.iter().peekable();
        #[cfg(feature = "extra-children")]
        let mut emit_idx: usize = 0;
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        if let Some(ref val) = self.ext_lst {
            val.write_element("a:extLst", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        for extra in extra_iter {
            extra.node.write_to(writer).map_err(SerializeError::from)?;
        }
        Ok(())
    }

    fn is_empty_element(&self) -> bool {
        if self.ext_lst.is_some() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for CTContentPartLocking {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        #[allow(unused_mut)]
        let mut start = start;
        if let Some(ref val) = self.no_grp {
            start.push_attribute(("noGrp", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.no_select {
            start.push_attribute(("noSelect", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.no_rot {
            start.push_attribute(("noRot", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.no_change_aspect {
            start.push_attribute(("noChangeAspect", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.no_move {
            start.push_attribute(("noMove", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.no_resize {
            start.push_attribute(("noResize", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.no_edit_points {
            start.push_attribute(("noEditPoints", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.no_adjust_handles {
            start.push_attribute(("noAdjustHandles", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.no_change_arrowheads {
            start.push_attribute(("noChangeArrowheads", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.no_change_shape_type {
            start.push_attribute(("noChangeShapeType", if *val { "1" } else { "0" }));
        }
        #[cfg(feature = "extra-attrs")]
        for (key, value) in &self.extra_attrs {
            start.push_attribute((key.as_str(), value.as_str()));
        }
        start
    }

    fn write_children<W: Write>(&self, writer: &mut Writer<W>) -> Result<(), SerializeError> {
        #[cfg(feature = "extra-children")]
        let mut extra_iter = self.extra_children.iter().peekable();
        #[cfg(feature = "extra-children")]
        let mut emit_idx: usize = 0;
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        if let Some(ref val) = self.ext_lst {
            val.write_element("a:extLst", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        for extra in extra_iter {
            extra.node.write_to(writer).map_err(SerializeError::from)?;
        }
        Ok(())
    }

    fn is_empty_element(&self) -> bool {
        if self.ext_lst.is_some() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for CTNonVisualDrawingProps {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        #[allow(unused_mut)]
        let mut start = start;
        {
            let val = &self.id;
            {
                let s = val.to_string();
                start.push_attribute(("id", s.as_str()));
            }
        }
        {
            let val = &self.name;
            start.push_attribute(("name", val.as_str()));
        }
        if let Some(ref val) = self.descr {
            start.push_attribute(("descr", val.as_str()));
        }
        if let Some(ref val) = self.hidden {
            start.push_attribute(("hidden", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.title {
            start.push_attribute(("title", val.as_str()));
        }
        #[cfg(feature = "extra-attrs")]
        for (key, value) in &self.extra_attrs {
            start.push_attribute((key.as_str(), value.as_str()));
        }
        start
    }

    fn write_children<W: Write>(&self, writer: &mut Writer<W>) -> Result<(), SerializeError> {
        #[cfg(feature = "extra-children")]
        let mut extra_iter = self.extra_children.iter().peekable();
        #[cfg(feature = "extra-children")]
        let mut emit_idx: usize = 0;
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        #[cfg(feature = "dml-text")]
        if let Some(ref val) = self.hlink_click {
            val.write_element("a:hlinkClick", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        #[cfg(feature = "dml-text")]
        if let Some(ref val) = self.hlink_hover {
            val.write_element("a:hlinkHover", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        #[cfg(feature = "dml-extensions")]
        if let Some(ref val) = self.ext_lst {
            val.write_element("a:extLst", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        for extra in extra_iter {
            extra.node.write_to(writer).map_err(SerializeError::from)?;
        }
        Ok(())
    }

    fn is_empty_element(&self) -> bool {
        #[cfg(feature = "dml-text")]
        if self.hlink_click.is_some() {
            return false;
        }
        #[cfg(feature = "dml-text")]
        if self.hlink_hover.is_some() {
            return false;
        }
        #[cfg(feature = "dml-extensions")]
        if self.ext_lst.is_some() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for CTNonVisualDrawingShapeProps {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        #[allow(unused_mut)]
        let mut start = start;
        if let Some(ref val) = self.tx_box {
            start.push_attribute(("txBox", if *val { "1" } else { "0" }));
        }
        #[cfg(feature = "extra-attrs")]
        for (key, value) in &self.extra_attrs {
            start.push_attribute((key.as_str(), value.as_str()));
        }
        start
    }

    fn write_children<W: Write>(&self, writer: &mut Writer<W>) -> Result<(), SerializeError> {
        #[cfg(feature = "extra-children")]
        let mut extra_iter = self.extra_children.iter().peekable();
        #[cfg(feature = "extra-children")]
        let mut emit_idx: usize = 0;
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        if let Some(ref val) = self.sp_locks {
            val.write_element("a:spLocks", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        if let Some(ref val) = self.ext_lst {
            val.write_element("a:extLst", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        for extra in extra_iter {
            extra.node.write_to(writer).map_err(SerializeError::from)?;
        }
        Ok(())
    }

    fn is_empty_element(&self) -> bool {
        if self.sp_locks.is_some() {
            return false;
        }
        if self.ext_lst.is_some() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for CTNonVisualConnectorProperties {
    fn write_children<W: Write>(&self, writer: &mut Writer<W>) -> Result<(), SerializeError> {
        #[cfg(feature = "extra-children")]
        let mut extra_iter = self.extra_children.iter().peekable();
        #[cfg(feature = "extra-children")]
        let mut emit_idx: usize = 0;
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        #[cfg(feature = "dml-shapes")]
        if let Some(ref val) = self.cxn_sp_locks {
            val.write_element("a:cxnSpLocks", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        #[cfg(feature = "dml-shapes")]
        if let Some(ref val) = self.st_cxn {
            val.write_element("a:stCxn", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        #[cfg(feature = "dml-shapes")]
        if let Some(ref val) = self.end_cxn {
            val.write_element("a:endCxn", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        #[cfg(feature = "dml-extensions")]
        if let Some(ref val) = self.ext_lst {
            val.write_element("a:extLst", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        for extra in extra_iter {
            extra.node.write_to(writer).map_err(SerializeError::from)?;
        }
        Ok(())
    }

    fn is_empty_element(&self) -> bool {
        #[cfg(feature = "dml-shapes")]
        if self.cxn_sp_locks.is_some() {
            return false;
        }
        #[cfg(feature = "dml-shapes")]
        if self.st_cxn.is_some() {
            return false;
        }
        #[cfg(feature = "dml-shapes")]
        if self.end_cxn.is_some() {
            return false;
        }
        #[cfg(feature = "dml-extensions")]
        if self.ext_lst.is_some() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for CTNonVisualPictureProperties {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        #[allow(unused_mut)]
        let mut start = start;
        #[cfg(feature = "dml-shapes")]
        if let Some(ref val) = self.prefer_relative_resize {
            start.push_attribute(("preferRelativeResize", if *val { "1" } else { "0" }));
        }
        #[cfg(feature = "extra-attrs")]
        for (key, value) in &self.extra_attrs {
            start.push_attribute((key.as_str(), value.as_str()));
        }
        start
    }

    fn write_children<W: Write>(&self, writer: &mut Writer<W>) -> Result<(), SerializeError> {
        #[cfg(feature = "extra-children")]
        let mut extra_iter = self.extra_children.iter().peekable();
        #[cfg(feature = "extra-children")]
        let mut emit_idx: usize = 0;
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        #[cfg(feature = "dml-shapes")]
        if let Some(ref val) = self.pic_locks {
            val.write_element("a:picLocks", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        #[cfg(feature = "dml-extensions")]
        if let Some(ref val) = self.ext_lst {
            val.write_element("a:extLst", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        for extra in extra_iter {
            extra.node.write_to(writer).map_err(SerializeError::from)?;
        }
        Ok(())
    }

    fn is_empty_element(&self) -> bool {
        #[cfg(feature = "dml-shapes")]
        if self.pic_locks.is_some() {
            return false;
        }
        #[cfg(feature = "dml-extensions")]
        if self.ext_lst.is_some() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for CTNonVisualGroupDrawingShapeProps {
    fn write_children<W: Write>(&self, writer: &mut Writer<W>) -> Result<(), SerializeError> {
        #[cfg(feature = "extra-children")]
        let mut extra_iter = self.extra_children.iter().peekable();
        #[cfg(feature = "extra-children")]
        let mut emit_idx: usize = 0;
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        #[cfg(feature = "dml-shapes")]
        if let Some(ref val) = self.grp_sp_locks {
            val.write_element("a:grpSpLocks", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        #[cfg(feature = "dml-extensions")]
        if let Some(ref val) = self.ext_lst {
            val.write_element("a:extLst", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        for extra in extra_iter {
            extra.node.write_to(writer).map_err(SerializeError::from)?;
        }
        Ok(())
    }

    fn is_empty_element(&self) -> bool {
        #[cfg(feature = "dml-shapes")]
        if self.grp_sp_locks.is_some() {
            return false;
        }
        #[cfg(feature = "dml-extensions")]
        if self.ext_lst.is_some() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for CTNonVisualGraphicFrameProperties {
    fn write_children<W: Write>(&self, writer: &mut Writer<W>) -> Result<(), SerializeError> {
        #[cfg(feature = "extra-children")]
        let mut extra_iter = self.extra_children.iter().peekable();
        #[cfg(feature = "extra-children")]
        let mut emit_idx: usize = 0;
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        #[cfg(feature = "dml-shapes")]
        if let Some(ref val) = self.graphic_frame_locks {
            val.write_element("a:graphicFrameLocks", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        #[cfg(feature = "dml-extensions")]
        if let Some(ref val) = self.ext_lst {
            val.write_element("a:extLst", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        for extra in extra_iter {
            extra.node.write_to(writer).map_err(SerializeError::from)?;
        }
        Ok(())
    }

    fn is_empty_element(&self) -> bool {
        #[cfg(feature = "dml-shapes")]
        if self.graphic_frame_locks.is_some() {
            return false;
        }
        #[cfg(feature = "dml-extensions")]
        if self.ext_lst.is_some() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for CTNonVisualContentPartProperties {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        #[allow(unused_mut)]
        let mut start = start;
        if let Some(ref val) = self.is_comment {
            start.push_attribute(("isComment", if *val { "1" } else { "0" }));
        }
        #[cfg(feature = "extra-attrs")]
        for (key, value) in &self.extra_attrs {
            start.push_attribute((key.as_str(), value.as_str()));
        }
        start
    }

    fn write_children<W: Write>(&self, writer: &mut Writer<W>) -> Result<(), SerializeError> {
        #[cfg(feature = "extra-children")]
        let mut extra_iter = self.extra_children.iter().peekable();
        #[cfg(feature = "extra-children")]
        let mut emit_idx: usize = 0;
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        if let Some(ref val) = self.cp_locks {
            val.write_element("a:cpLocks", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        if let Some(ref val) = self.ext_lst {
            val.write_element("a:extLst", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        for extra in extra_iter {
            extra.node.write_to(writer).map_err(SerializeError::from)?;
        }
        Ok(())
    }

    fn is_empty_element(&self) -> bool {
        if self.cp_locks.is_some() {
            return false;
        }
        if self.ext_lst.is_some() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for CTGraphicalObjectData {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        #[allow(unused_mut)]
        let mut start = start;
        {
            let val = &self.uri;
            start.push_attribute(("uri", val.as_str()));
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

impl ToXml for CTAnimationDgmElement {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        #[allow(unused_mut)]
        let mut start = start;
        if let Some(ref val) = self.id {
            start.push_attribute(("id", val.as_str()));
        }
        if let Some(ref val) = self.bld_step {
            {
                let s = val.to_string();
                start.push_attribute(("bldStep", s.as_str()));
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

impl ToXml for CTAnimationChartElement {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        #[allow(unused_mut)]
        let mut start = start;
        if let Some(ref val) = self.series_idx {
            {
                let s = val.to_string();
                start.push_attribute(("seriesIdx", s.as_str()));
            }
        }
        if let Some(ref val) = self.category_idx {
            {
                let s = val.to_string();
                start.push_attribute(("categoryIdx", s.as_str()));
            }
        }
        {
            let val = &self.bld_step;
            {
                let s = val.to_string();
                start.push_attribute(("bldStep", s.as_str()));
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

impl ToXml for CTAnimationElementChoice {
    fn write_children<W: Write>(&self, writer: &mut Writer<W>) -> Result<(), SerializeError> {
        #[cfg(feature = "extra-children")]
        let mut extra_iter = self.extra_children.iter().peekable();
        #[cfg(feature = "extra-children")]
        let mut emit_idx: usize = 0;
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        if let Some(ref val) = self.dgm {
            val.write_element("a:dgm", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        if let Some(ref val) = self.chart {
            val.write_element("a:chart", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        for extra in extra_iter {
            extra.node.write_to(writer).map_err(SerializeError::from)?;
        }
        Ok(())
    }

    fn is_empty_element(&self) -> bool {
        if self.dgm.is_some() {
            return false;
        }
        if self.chart.is_some() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for CTAnimationDgmBuildProperties {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        #[allow(unused_mut)]
        let mut start = start;
        if let Some(ref val) = self.bld {
            {
                let s = val.to_string();
                start.push_attribute(("bld", s.as_str()));
            }
        }
        if let Some(ref val) = self.rev {
            start.push_attribute(("rev", if *val { "1" } else { "0" }));
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

impl ToXml for CTAnimationChartBuildProperties {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        #[allow(unused_mut)]
        let mut start = start;
        if let Some(ref val) = self.bld {
            {
                let s = val.to_string();
                start.push_attribute(("bld", s.as_str()));
            }
        }
        if let Some(ref val) = self.anim_bg {
            start.push_attribute(("animBg", if *val { "1" } else { "0" }));
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

impl ToXml for CTAnimationGraphicalObjectBuildProperties {
    fn write_children<W: Write>(&self, writer: &mut Writer<W>) -> Result<(), SerializeError> {
        #[cfg(feature = "extra-children")]
        let mut extra_iter = self.extra_children.iter().peekable();
        #[cfg(feature = "extra-children")]
        let mut emit_idx: usize = 0;
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        if let Some(ref val) = self.bld_dgm {
            val.write_element("a:bldDgm", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        if let Some(ref val) = self.bld_chart {
            val.write_element("a:bldChart", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        for extra in extra_iter {
            extra.node.write_to(writer).map_err(SerializeError::from)?;
        }
        Ok(())
    }

    fn is_empty_element(&self) -> bool {
        if self.bld_dgm.is_some() {
            return false;
        }
        if self.bld_chart.is_some() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for CTBackgroundFormatting {
    fn write_children<W: Write>(&self, writer: &mut Writer<W>) -> Result<(), SerializeError> {
        #[cfg(feature = "extra-children")]
        let mut extra_iter = self.extra_children.iter().peekable();
        #[cfg(feature = "extra-children")]
        let mut emit_idx: usize = 0;
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        if let Some(ref val) = self.fill_properties {
            val.write_element("", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        if let Some(ref val) = self.effect_properties {
            val.write_element("", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        for extra in extra_iter {
            extra.node.write_to(writer).map_err(SerializeError::from)?;
        }
        Ok(())
    }

    fn is_empty_element(&self) -> bool {
        if self.fill_properties.is_some() {
            return false;
        }
        if self.effect_properties.is_some() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for CTWholeE2oFormatting {
    fn write_children<W: Write>(&self, writer: &mut Writer<W>) -> Result<(), SerializeError> {
        #[cfg(feature = "extra-children")]
        let mut extra_iter = self.extra_children.iter().peekable();
        #[cfg(feature = "extra-children")]
        let mut emit_idx: usize = 0;
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        if let Some(ref val) = self.line {
            val.write_element("a:ln", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        if let Some(ref val) = self.effect_properties {
            val.write_element("", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        for extra in extra_iter {
            extra.node.write_to(writer).map_err(SerializeError::from)?;
        }
        Ok(())
    }

    fn is_empty_element(&self) -> bool {
        if self.line.is_some() {
            return false;
        }
        if self.effect_properties.is_some() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for CTGvmlUseShapeRectangle {
    fn is_empty_element(&self) -> bool {
        true
    }
}

impl ToXml for CTGvmlTextShape {
    fn write_children<W: Write>(&self, writer: &mut Writer<W>) -> Result<(), SerializeError> {
        #[cfg(feature = "extra-children")]
        let mut extra_iter = self.extra_children.iter().peekable();
        #[cfg(feature = "extra-children")]
        let mut emit_idx: usize = 0;
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        {
            let val = &self.tx_body;
            val.write_element("a:txBody", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        if let Some(ref val) = self.use_sp_rect {
            val.write_element("a:useSpRect", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        if let Some(ref val) = self.transform {
            val.write_element("a:xfrm", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        if let Some(ref val) = self.ext_lst {
            val.write_element("a:extLst", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        for extra in extra_iter {
            extra.node.write_to(writer).map_err(SerializeError::from)?;
        }
        Ok(())
    }

    fn is_empty_element(&self) -> bool {
        false
    }
}

impl ToXml for CTGvmlShapeNonVisual {
    fn write_children<W: Write>(&self, writer: &mut Writer<W>) -> Result<(), SerializeError> {
        #[cfg(feature = "extra-children")]
        let mut extra_iter = self.extra_children.iter().peekable();
        #[cfg(feature = "extra-children")]
        let mut emit_idx: usize = 0;
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        {
            let val = &self.common_non_visual_properties;
            val.write_element("a:cNvPr", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        {
            let val = &self.common_non_visual_shape_properties;
            val.write_element("a:cNvSpPr", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        for extra in extra_iter {
            extra.node.write_to(writer).map_err(SerializeError::from)?;
        }
        Ok(())
    }

    fn is_empty_element(&self) -> bool {
        false
    }
}

impl ToXml for CTGvmlShape {
    fn write_children<W: Write>(&self, writer: &mut Writer<W>) -> Result<(), SerializeError> {
        #[cfg(feature = "extra-children")]
        let mut extra_iter = self.extra_children.iter().peekable();
        #[cfg(feature = "extra-children")]
        let mut emit_idx: usize = 0;
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        {
            let val = &self.nv_sp_pr;
            val.write_element("a:nvSpPr", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        {
            let val = &self.sp_pr;
            val.write_element("a:spPr", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        if let Some(ref val) = self.tx_sp {
            val.write_element("a:txSp", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        if let Some(ref val) = self.style {
            val.write_element("a:style", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        if let Some(ref val) = self.ext_lst {
            val.write_element("a:extLst", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        for extra in extra_iter {
            extra.node.write_to(writer).map_err(SerializeError::from)?;
        }
        Ok(())
    }

    fn is_empty_element(&self) -> bool {
        false
    }
}

impl ToXml for CTGvmlConnectorNonVisual {
    fn write_children<W: Write>(&self, writer: &mut Writer<W>) -> Result<(), SerializeError> {
        #[cfg(feature = "extra-children")]
        let mut extra_iter = self.extra_children.iter().peekable();
        #[cfg(feature = "extra-children")]
        let mut emit_idx: usize = 0;
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        {
            let val = &self.common_non_visual_properties;
            val.write_element("a:cNvPr", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        {
            let val = &self.c_nv_cxn_sp_pr;
            val.write_element("a:cNvCxnSpPr", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        for extra in extra_iter {
            extra.node.write_to(writer).map_err(SerializeError::from)?;
        }
        Ok(())
    }

    fn is_empty_element(&self) -> bool {
        false
    }
}

impl ToXml for CTGvmlConnector {
    fn write_children<W: Write>(&self, writer: &mut Writer<W>) -> Result<(), SerializeError> {
        #[cfg(feature = "extra-children")]
        let mut extra_iter = self.extra_children.iter().peekable();
        #[cfg(feature = "extra-children")]
        let mut emit_idx: usize = 0;
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        {
            let val = &self.nv_cxn_sp_pr;
            val.write_element("a:nvCxnSpPr", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        {
            let val = &self.sp_pr;
            val.write_element("a:spPr", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        if let Some(ref val) = self.style {
            val.write_element("a:style", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        if let Some(ref val) = self.ext_lst {
            val.write_element("a:extLst", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        for extra in extra_iter {
            extra.node.write_to(writer).map_err(SerializeError::from)?;
        }
        Ok(())
    }

    fn is_empty_element(&self) -> bool {
        false
    }
}

impl ToXml for CTGvmlPictureNonVisual {
    fn write_children<W: Write>(&self, writer: &mut Writer<W>) -> Result<(), SerializeError> {
        #[cfg(feature = "extra-children")]
        let mut extra_iter = self.extra_children.iter().peekable();
        #[cfg(feature = "extra-children")]
        let mut emit_idx: usize = 0;
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        {
            let val = &self.common_non_visual_properties;
            val.write_element("a:cNvPr", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        {
            let val = &self.common_non_visual_picture_properties;
            val.write_element("a:cNvPicPr", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        for extra in extra_iter {
            extra.node.write_to(writer).map_err(SerializeError::from)?;
        }
        Ok(())
    }

    fn is_empty_element(&self) -> bool {
        false
    }
}

impl ToXml for CTGvmlPicture {
    fn write_children<W: Write>(&self, writer: &mut Writer<W>) -> Result<(), SerializeError> {
        #[cfg(feature = "extra-children")]
        let mut extra_iter = self.extra_children.iter().peekable();
        #[cfg(feature = "extra-children")]
        let mut emit_idx: usize = 0;
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        {
            let val = &self.nv_pic_pr;
            val.write_element("a:nvPicPr", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        {
            let val = &self.blip_fill;
            val.write_element("a:blipFill", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        {
            let val = &self.sp_pr;
            val.write_element("a:spPr", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        if let Some(ref val) = self.style {
            val.write_element("a:style", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        if let Some(ref val) = self.ext_lst {
            val.write_element("a:extLst", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        for extra in extra_iter {
            extra.node.write_to(writer).map_err(SerializeError::from)?;
        }
        Ok(())
    }

    fn is_empty_element(&self) -> bool {
        false
    }
}

impl ToXml for CTGvmlGraphicFrameNonVisual {
    fn write_children<W: Write>(&self, writer: &mut Writer<W>) -> Result<(), SerializeError> {
        #[cfg(feature = "extra-children")]
        let mut extra_iter = self.extra_children.iter().peekable();
        #[cfg(feature = "extra-children")]
        let mut emit_idx: usize = 0;
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        {
            let val = &self.common_non_visual_properties;
            val.write_element("a:cNvPr", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        {
            let val = &self.c_nv_graphic_frame_pr;
            val.write_element("a:cNvGraphicFramePr", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        for extra in extra_iter {
            extra.node.write_to(writer).map_err(SerializeError::from)?;
        }
        Ok(())
    }

    fn is_empty_element(&self) -> bool {
        false
    }
}

impl ToXml for CTGvmlGraphicalObjectFrame {
    fn write_children<W: Write>(&self, writer: &mut Writer<W>) -> Result<(), SerializeError> {
        #[cfg(feature = "extra-children")]
        let mut extra_iter = self.extra_children.iter().peekable();
        #[cfg(feature = "extra-children")]
        let mut emit_idx: usize = 0;
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        {
            let val = &self.nv_graphic_frame_pr;
            val.write_element("a:nvGraphicFramePr", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        {
            let val = &self.graphic;
            val.write_element("a:graphic", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        {
            let val = &self.transform;
            val.write_element("a:xfrm", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        if let Some(ref val) = self.ext_lst {
            val.write_element("a:extLst", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        for extra in extra_iter {
            extra.node.write_to(writer).map_err(SerializeError::from)?;
        }
        Ok(())
    }

    fn is_empty_element(&self) -> bool {
        false
    }
}

impl ToXml for CTGvmlGroupShapeNonVisual {
    fn write_children<W: Write>(&self, writer: &mut Writer<W>) -> Result<(), SerializeError> {
        #[cfg(feature = "extra-children")]
        let mut extra_iter = self.extra_children.iter().peekable();
        #[cfg(feature = "extra-children")]
        let mut emit_idx: usize = 0;
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        {
            let val = &self.common_non_visual_properties;
            val.write_element("a:cNvPr", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        {
            let val = &self.c_nv_grp_sp_pr;
            val.write_element("a:cNvGrpSpPr", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        for extra in extra_iter {
            extra.node.write_to(writer).map_err(SerializeError::from)?;
        }
        Ok(())
    }

    fn is_empty_element(&self) -> bool {
        false
    }
}

impl ToXml for CTGvmlGroupShape {
    fn write_children<W: Write>(&self, writer: &mut Writer<W>) -> Result<(), SerializeError> {
        #[cfg(feature = "extra-children")]
        let mut extra_iter = self.extra_children.iter().peekable();
        #[cfg(feature = "extra-children")]
        let mut emit_idx: usize = 0;
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        {
            let val = &self.nv_grp_sp_pr;
            val.write_element("a:nvGrpSpPr", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        {
            let val = &self.grp_sp_pr;
            val.write_element("a:grpSpPr", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        if let Some(ref val) = self.tx_sp {
            val.write_element("a:txSp", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        if let Some(ref val) = self.sp {
            val.write_element("a:sp", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        if let Some(ref val) = self.cxn_sp {
            val.write_element("a:cxnSp", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        if let Some(ref val) = self.pic {
            val.write_element("a:pic", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        if let Some(ref val) = self.graphic_frame {
            val.write_element("a:graphicFrame", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        if let Some(ref val) = self.grp_sp {
            val.write_element("a:grpSp", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        if let Some(ref val) = self.ext_lst {
            val.write_element("a:extLst", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        for extra in extra_iter {
            extra.node.write_to(writer).map_err(SerializeError::from)?;
        }
        Ok(())
    }

    fn is_empty_element(&self) -> bool {
        false
    }
}

impl ToXml for CTCamera {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        #[allow(unused_mut)]
        let mut start = start;
        #[cfg(feature = "dml-3d")]
        {
            let val = &self.preset;
            {
                let s = val.to_string();
                start.push_attribute(("prst", s.as_str()));
            }
        }
        #[cfg(feature = "dml-3d")]
        if let Some(ref val) = self.fov {
            {
                let s = val.to_string();
                start.push_attribute(("fov", s.as_str()));
            }
        }
        #[cfg(feature = "dml-3d")]
        if let Some(ref val) = self.zoom {
            {
                let s = val.to_string();
                start.push_attribute(("zoom", s.as_str()));
            }
        }
        #[cfg(feature = "extra-attrs")]
        for (key, value) in &self.extra_attrs {
            start.push_attribute((key.as_str(), value.as_str()));
        }
        start
    }

    fn write_children<W: Write>(&self, writer: &mut Writer<W>) -> Result<(), SerializeError> {
        #[cfg(feature = "extra-children")]
        let mut extra_iter = self.extra_children.iter().peekable();
        #[cfg(feature = "extra-children")]
        let mut emit_idx: usize = 0;
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        #[cfg(feature = "dml-3d")]
        if let Some(ref val) = self.rot {
            val.write_element("a:rot", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        for extra in extra_iter {
            extra.node.write_to(writer).map_err(SerializeError::from)?;
        }
        Ok(())
    }

    fn is_empty_element(&self) -> bool {
        #[cfg(feature = "dml-3d")]
        if self.rot.is_some() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for CTLightRig {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        #[allow(unused_mut)]
        let mut start = start;
        #[cfg(feature = "dml-3d")]
        {
            let val = &self.rig;
            {
                let s = val.to_string();
                start.push_attribute(("rig", s.as_str()));
            }
        }
        #[cfg(feature = "dml-3d")]
        {
            let val = &self.dir;
            {
                let s = val.to_string();
                start.push_attribute(("dir", s.as_str()));
            }
        }
        #[cfg(feature = "extra-attrs")]
        for (key, value) in &self.extra_attrs {
            start.push_attribute((key.as_str(), value.as_str()));
        }
        start
    }

    fn write_children<W: Write>(&self, writer: &mut Writer<W>) -> Result<(), SerializeError> {
        #[cfg(feature = "extra-children")]
        let mut extra_iter = self.extra_children.iter().peekable();
        #[cfg(feature = "extra-children")]
        let mut emit_idx: usize = 0;
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        #[cfg(feature = "dml-3d")]
        if let Some(ref val) = self.rot {
            val.write_element("a:rot", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        for extra in extra_iter {
            extra.node.write_to(writer).map_err(SerializeError::from)?;
        }
        Ok(())
    }

    fn is_empty_element(&self) -> bool {
        #[cfg(feature = "dml-3d")]
        if self.rot.is_some() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for CTScene3D {
    fn write_children<W: Write>(&self, writer: &mut Writer<W>) -> Result<(), SerializeError> {
        #[cfg(feature = "extra-children")]
        let mut extra_iter = self.extra_children.iter().peekable();
        #[cfg(feature = "extra-children")]
        let mut emit_idx: usize = 0;
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        #[cfg(feature = "dml-3d")]
        {
            let val = &self.camera;
            val.write_element("a:camera", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        #[cfg(feature = "dml-3d")]
        {
            let val = &self.light_rig;
            val.write_element("a:lightRig", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        #[cfg(feature = "dml-3d")]
        if let Some(ref val) = self.backdrop {
            val.write_element("a:backdrop", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        #[cfg(feature = "dml-extensions")]
        if let Some(ref val) = self.ext_lst {
            val.write_element("a:extLst", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        for extra in extra_iter {
            extra.node.write_to(writer).map_err(SerializeError::from)?;
        }
        Ok(())
    }

    fn is_empty_element(&self) -> bool {
        #[cfg(feature = "dml-3d")]
        return false;
        #[cfg(feature = "dml-3d")]
        return false;
        #[cfg(feature = "dml-3d")]
        if self.backdrop.is_some() {
            return false;
        }
        #[cfg(feature = "dml-extensions")]
        if self.ext_lst.is_some() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for CTBackdrop {
    fn write_children<W: Write>(&self, writer: &mut Writer<W>) -> Result<(), SerializeError> {
        #[cfg(feature = "extra-children")]
        let mut extra_iter = self.extra_children.iter().peekable();
        #[cfg(feature = "extra-children")]
        let mut emit_idx: usize = 0;
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        {
            let val = &self.anchor;
            val.write_element("a:anchor", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        {
            let val = &self.norm;
            val.write_element("a:norm", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        {
            let val = &self.up;
            val.write_element("a:up", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        if let Some(ref val) = self.ext_lst {
            val.write_element("a:extLst", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        for extra in extra_iter {
            extra.node.write_to(writer).map_err(SerializeError::from)?;
        }
        Ok(())
    }

    fn is_empty_element(&self) -> bool {
        false
    }
}

impl ToXml for CTBevel {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        #[allow(unused_mut)]
        let mut start = start;
        #[cfg(feature = "dml-3d")]
        if let Some(ref val) = self.width {
            {
                let s = val.to_string();
                start.push_attribute(("w", s.as_str()));
            }
        }
        #[cfg(feature = "dml-3d")]
        if let Some(ref val) = self.height {
            {
                let s = val.to_string();
                start.push_attribute(("h", s.as_str()));
            }
        }
        #[cfg(feature = "dml-3d")]
        if let Some(ref val) = self.preset {
            {
                let s = val.to_string();
                start.push_attribute(("prst", s.as_str()));
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

impl ToXml for CTShape3D {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        #[allow(unused_mut)]
        let mut start = start;
        #[cfg(feature = "dml-3d")]
        if let Some(ref val) = self.z {
            {
                let s = val.to_string();
                start.push_attribute(("z", s.as_str()));
            }
        }
        #[cfg(feature = "dml-3d")]
        if let Some(ref val) = self.extrusion_h {
            {
                let s = val.to_string();
                start.push_attribute(("extrusionH", s.as_str()));
            }
        }
        #[cfg(feature = "dml-3d")]
        if let Some(ref val) = self.contour_w {
            {
                let s = val.to_string();
                start.push_attribute(("contourW", s.as_str()));
            }
        }
        #[cfg(feature = "dml-3d")]
        if let Some(ref val) = self.prst_material {
            {
                let s = val.to_string();
                start.push_attribute(("prstMaterial", s.as_str()));
            }
        }
        #[cfg(feature = "extra-attrs")]
        for (key, value) in &self.extra_attrs {
            start.push_attribute((key.as_str(), value.as_str()));
        }
        start
    }

    fn write_children<W: Write>(&self, writer: &mut Writer<W>) -> Result<(), SerializeError> {
        #[cfg(feature = "extra-children")]
        let mut extra_iter = self.extra_children.iter().peekable();
        #[cfg(feature = "extra-children")]
        let mut emit_idx: usize = 0;
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        #[cfg(feature = "dml-3d")]
        if let Some(ref val) = self.bevel_t {
            val.write_element("a:bevelT", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        #[cfg(feature = "dml-3d")]
        if let Some(ref val) = self.bevel_b {
            val.write_element("a:bevelB", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        #[cfg(feature = "dml-3d")]
        if let Some(ref val) = self.extrusion_clr {
            val.write_element("a:extrusionClr", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        #[cfg(feature = "dml-3d")]
        if let Some(ref val) = self.contour_clr {
            val.write_element("a:contourClr", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        #[cfg(feature = "dml-extensions")]
        if let Some(ref val) = self.ext_lst {
            val.write_element("a:extLst", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        for extra in extra_iter {
            extra.node.write_to(writer).map_err(SerializeError::from)?;
        }
        Ok(())
    }

    fn is_empty_element(&self) -> bool {
        #[cfg(feature = "dml-3d")]
        if self.bevel_t.is_some() {
            return false;
        }
        #[cfg(feature = "dml-3d")]
        if self.bevel_b.is_some() {
            return false;
        }
        #[cfg(feature = "dml-3d")]
        if self.extrusion_clr.is_some() {
            return false;
        }
        #[cfg(feature = "dml-3d")]
        if self.contour_clr.is_some() {
            return false;
        }
        #[cfg(feature = "dml-extensions")]
        if self.ext_lst.is_some() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for CTFlatText {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        #[allow(unused_mut)]
        let mut start = start;
        if let Some(ref val) = self.z {
            {
                let s = val.to_string();
                start.push_attribute(("z", s.as_str()));
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

impl ToXml for EGText3D {
    fn write_element<W: Write>(
        &self,
        _tag: &str,
        writer: &mut Writer<W>,
    ) -> Result<(), SerializeError> {
        match self {
            Self::Sp3d(inner) => inner.write_element("a:sp3d", writer)?,
            Self::FlatTx(inner) => inner.write_element("a:flatTx", writer)?,
        }
        Ok(())
    }
}

impl ToXml for CTAlphaBiLevelEffect {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        #[allow(unused_mut)]
        let mut start = start;
        {
            let val = &self.thresh;
            {
                let s = val.to_string();
                start.push_attribute(("thresh", s.as_str()));
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

impl ToXml for CTAlphaCeilingEffect {
    fn is_empty_element(&self) -> bool {
        true
    }
}

impl ToXml for CTAlphaFloorEffect {
    fn is_empty_element(&self) -> bool {
        true
    }
}

impl ToXml for CTAlphaInverseEffect {
    fn write_children<W: Write>(&self, writer: &mut Writer<W>) -> Result<(), SerializeError> {
        #[cfg(feature = "extra-children")]
        let mut extra_iter = self.extra_children.iter().peekable();
        #[cfg(feature = "extra-children")]
        let mut emit_idx: usize = 0;
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        if let Some(ref val) = self.color_choice {
            val.write_element("", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        for extra in extra_iter {
            extra.node.write_to(writer).map_err(SerializeError::from)?;
        }
        Ok(())
    }

    fn is_empty_element(&self) -> bool {
        if self.color_choice.is_some() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for CTAlphaModulateFixedEffect {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        #[allow(unused_mut)]
        let mut start = start;
        if let Some(ref val) = self.amt {
            {
                let s = val.to_string();
                start.push_attribute(("amt", s.as_str()));
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

impl ToXml for CTAlphaOutsetEffect {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        #[allow(unused_mut)]
        let mut start = start;
        if let Some(ref val) = self.rad {
            {
                let s = val.to_string();
                start.push_attribute(("rad", s.as_str()));
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

impl ToXml for CTAlphaReplaceEffect {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        #[allow(unused_mut)]
        let mut start = start;
        {
            let val = &self.anchor;
            {
                let s = val.to_string();
                start.push_attribute(("a", s.as_str()));
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

impl ToXml for CTBiLevelEffect {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        #[allow(unused_mut)]
        let mut start = start;
        {
            let val = &self.thresh;
            {
                let s = val.to_string();
                start.push_attribute(("thresh", s.as_str()));
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

impl ToXml for CTBlurEffect {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        #[allow(unused_mut)]
        let mut start = start;
        #[cfg(feature = "dml-effects")]
        if let Some(ref val) = self.rad {
            {
                let s = val.to_string();
                start.push_attribute(("rad", s.as_str()));
            }
        }
        #[cfg(feature = "dml-effects")]
        if let Some(ref val) = self.grow {
            start.push_attribute(("grow", if *val { "1" } else { "0" }));
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

impl ToXml for CTColorChangeEffect {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        #[allow(unused_mut)]
        let mut start = start;
        if let Some(ref val) = self.use_a {
            start.push_attribute(("useA", if *val { "1" } else { "0" }));
        }
        #[cfg(feature = "extra-attrs")]
        for (key, value) in &self.extra_attrs {
            start.push_attribute((key.as_str(), value.as_str()));
        }
        start
    }

    fn write_children<W: Write>(&self, writer: &mut Writer<W>) -> Result<(), SerializeError> {
        #[cfg(feature = "extra-children")]
        let mut extra_iter = self.extra_children.iter().peekable();
        #[cfg(feature = "extra-children")]
        let mut emit_idx: usize = 0;
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        {
            let val = &self.clr_from;
            val.write_element("a:clrFrom", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        {
            let val = &self.clr_to;
            val.write_element("a:clrTo", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        for extra in extra_iter {
            extra.node.write_to(writer).map_err(SerializeError::from)?;
        }
        Ok(())
    }

    fn is_empty_element(&self) -> bool {
        false
    }
}

impl ToXml for CTColorReplaceEffect {
    fn write_children<W: Write>(&self, writer: &mut Writer<W>) -> Result<(), SerializeError> {
        #[cfg(feature = "extra-children")]
        let mut extra_iter = self.extra_children.iter().peekable();
        #[cfg(feature = "extra-children")]
        let mut emit_idx: usize = 0;
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        if let Some(ref val) = self.color_choice {
            val.write_element("", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        for extra in extra_iter {
            extra.node.write_to(writer).map_err(SerializeError::from)?;
        }
        Ok(())
    }

    fn is_empty_element(&self) -> bool {
        false
    }
}

impl ToXml for CTDuotoneEffect {
    fn write_children<W: Write>(&self, writer: &mut Writer<W>) -> Result<(), SerializeError> {
        #[cfg(feature = "extra-children")]
        let mut extra_iter = self.extra_children.iter().peekable();
        #[cfg(feature = "extra-children")]
        let mut emit_idx: usize = 0;
        for item in &self.color_choice {
            #[cfg(feature = "extra-children")]
            while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
                extra_iter
                    .next()
                    .unwrap()
                    .node
                    .write_to(writer)
                    .map_err(SerializeError::from)?;
            }
            item.write_element("", writer)?;
            #[cfg(feature = "extra-children")]
            {
                emit_idx += 1;
            }
        }
        #[cfg(feature = "extra-children")]
        for extra in extra_iter {
            extra.node.write_to(writer).map_err(SerializeError::from)?;
        }
        Ok(())
    }

    fn is_empty_element(&self) -> bool {
        if !self.color_choice.is_empty() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for CTGlowEffect {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        #[allow(unused_mut)]
        let mut start = start;
        #[cfg(feature = "dml-effects")]
        if let Some(ref val) = self.rad {
            {
                let s = val.to_string();
                start.push_attribute(("rad", s.as_str()));
            }
        }
        #[cfg(feature = "extra-attrs")]
        for (key, value) in &self.extra_attrs {
            start.push_attribute((key.as_str(), value.as_str()));
        }
        start
    }

    fn write_children<W: Write>(&self, writer: &mut Writer<W>) -> Result<(), SerializeError> {
        #[cfg(feature = "extra-children")]
        let mut extra_iter = self.extra_children.iter().peekable();
        #[cfg(feature = "extra-children")]
        let mut emit_idx: usize = 0;
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        if let Some(ref val) = self.color_choice {
            val.write_element("", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        for extra in extra_iter {
            extra.node.write_to(writer).map_err(SerializeError::from)?;
        }
        Ok(())
    }

    fn is_empty_element(&self) -> bool {
        false
    }
}

impl ToXml for CTGrayscaleEffect {
    fn is_empty_element(&self) -> bool {
        true
    }
}

impl ToXml for CTHSLEffect {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        #[allow(unused_mut)]
        let mut start = start;
        if let Some(ref val) = self.hue {
            {
                let s = val.to_string();
                start.push_attribute(("hue", s.as_str()));
            }
        }
        if let Some(ref val) = self.sat {
            {
                let s = val.to_string();
                start.push_attribute(("sat", s.as_str()));
            }
        }
        if let Some(ref val) = self.lum {
            {
                let s = val.to_string();
                start.push_attribute(("lum", s.as_str()));
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

impl ToXml for CTInnerShadowEffect {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        #[allow(unused_mut)]
        let mut start = start;
        #[cfg(feature = "dml-effects")]
        if let Some(ref val) = self.blur_rad {
            {
                let s = val.to_string();
                start.push_attribute(("blurRad", s.as_str()));
            }
        }
        #[cfg(feature = "dml-effects")]
        if let Some(ref val) = self.dist {
            {
                let s = val.to_string();
                start.push_attribute(("dist", s.as_str()));
            }
        }
        #[cfg(feature = "dml-effects")]
        if let Some(ref val) = self.dir {
            {
                let s = val.to_string();
                start.push_attribute(("dir", s.as_str()));
            }
        }
        #[cfg(feature = "extra-attrs")]
        for (key, value) in &self.extra_attrs {
            start.push_attribute((key.as_str(), value.as_str()));
        }
        start
    }

    fn write_children<W: Write>(&self, writer: &mut Writer<W>) -> Result<(), SerializeError> {
        #[cfg(feature = "extra-children")]
        let mut extra_iter = self.extra_children.iter().peekable();
        #[cfg(feature = "extra-children")]
        let mut emit_idx: usize = 0;
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        if let Some(ref val) = self.color_choice {
            val.write_element("", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        for extra in extra_iter {
            extra.node.write_to(writer).map_err(SerializeError::from)?;
        }
        Ok(())
    }

    fn is_empty_element(&self) -> bool {
        false
    }
}

impl ToXml for CTLuminanceEffect {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        #[allow(unused_mut)]
        let mut start = start;
        if let Some(ref val) = self.bright {
            {
                let s = val.to_string();
                start.push_attribute(("bright", s.as_str()));
            }
        }
        if let Some(ref val) = self.contrast {
            {
                let s = val.to_string();
                start.push_attribute(("contrast", s.as_str()));
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

impl ToXml for CTOuterShadowEffect {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        #[allow(unused_mut)]
        let mut start = start;
        #[cfg(feature = "dml-effects")]
        if let Some(ref val) = self.blur_rad {
            {
                let s = val.to_string();
                start.push_attribute(("blurRad", s.as_str()));
            }
        }
        #[cfg(feature = "dml-effects")]
        if let Some(ref val) = self.dist {
            {
                let s = val.to_string();
                start.push_attribute(("dist", s.as_str()));
            }
        }
        #[cfg(feature = "dml-effects")]
        if let Some(ref val) = self.dir {
            {
                let s = val.to_string();
                start.push_attribute(("dir", s.as_str()));
            }
        }
        #[cfg(feature = "dml-effects")]
        if let Some(ref val) = self.sx {
            {
                let s = val.to_string();
                start.push_attribute(("sx", s.as_str()));
            }
        }
        #[cfg(feature = "dml-effects")]
        if let Some(ref val) = self.sy {
            {
                let s = val.to_string();
                start.push_attribute(("sy", s.as_str()));
            }
        }
        #[cfg(feature = "dml-effects")]
        if let Some(ref val) = self.kx {
            {
                let s = val.to_string();
                start.push_attribute(("kx", s.as_str()));
            }
        }
        #[cfg(feature = "dml-effects")]
        if let Some(ref val) = self.ky {
            {
                let s = val.to_string();
                start.push_attribute(("ky", s.as_str()));
            }
        }
        #[cfg(feature = "dml-effects")]
        if let Some(ref val) = self.algn {
            {
                let s = val.to_string();
                start.push_attribute(("algn", s.as_str()));
            }
        }
        #[cfg(feature = "dml-effects")]
        if let Some(ref val) = self.rot_with_shape {
            start.push_attribute(("rotWithShape", if *val { "1" } else { "0" }));
        }
        #[cfg(feature = "extra-attrs")]
        for (key, value) in &self.extra_attrs {
            start.push_attribute((key.as_str(), value.as_str()));
        }
        start
    }

    fn write_children<W: Write>(&self, writer: &mut Writer<W>) -> Result<(), SerializeError> {
        #[cfg(feature = "extra-children")]
        let mut extra_iter = self.extra_children.iter().peekable();
        #[cfg(feature = "extra-children")]
        let mut emit_idx: usize = 0;
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        if let Some(ref val) = self.color_choice {
            val.write_element("", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        for extra in extra_iter {
            extra.node.write_to(writer).map_err(SerializeError::from)?;
        }
        Ok(())
    }

    fn is_empty_element(&self) -> bool {
        false
    }
}

impl ToXml for CTPresetShadowEffect {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        #[allow(unused_mut)]
        let mut start = start;
        {
            let val = &self.preset;
            {
                let s = val.to_string();
                start.push_attribute(("prst", s.as_str()));
            }
        }
        if let Some(ref val) = self.dist {
            {
                let s = val.to_string();
                start.push_attribute(("dist", s.as_str()));
            }
        }
        if let Some(ref val) = self.dir {
            {
                let s = val.to_string();
                start.push_attribute(("dir", s.as_str()));
            }
        }
        #[cfg(feature = "extra-attrs")]
        for (key, value) in &self.extra_attrs {
            start.push_attribute((key.as_str(), value.as_str()));
        }
        start
    }

    fn write_children<W: Write>(&self, writer: &mut Writer<W>) -> Result<(), SerializeError> {
        #[cfg(feature = "extra-children")]
        let mut extra_iter = self.extra_children.iter().peekable();
        #[cfg(feature = "extra-children")]
        let mut emit_idx: usize = 0;
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        if let Some(ref val) = self.color_choice {
            val.write_element("", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        for extra in extra_iter {
            extra.node.write_to(writer).map_err(SerializeError::from)?;
        }
        Ok(())
    }

    fn is_empty_element(&self) -> bool {
        false
    }
}

impl ToXml for CTReflectionEffect {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        #[allow(unused_mut)]
        let mut start = start;
        #[cfg(feature = "dml-effects")]
        if let Some(ref val) = self.blur_rad {
            {
                let s = val.to_string();
                start.push_attribute(("blurRad", s.as_str()));
            }
        }
        #[cfg(feature = "dml-effects")]
        if let Some(ref val) = self.st_a {
            {
                let s = val.to_string();
                start.push_attribute(("stA", s.as_str()));
            }
        }
        #[cfg(feature = "dml-effects")]
        if let Some(ref val) = self.st_pos {
            {
                let s = val.to_string();
                start.push_attribute(("stPos", s.as_str()));
            }
        }
        #[cfg(feature = "dml-effects")]
        if let Some(ref val) = self.end_a {
            {
                let s = val.to_string();
                start.push_attribute(("endA", s.as_str()));
            }
        }
        #[cfg(feature = "dml-effects")]
        if let Some(ref val) = self.end_pos {
            {
                let s = val.to_string();
                start.push_attribute(("endPos", s.as_str()));
            }
        }
        #[cfg(feature = "dml-effects")]
        if let Some(ref val) = self.dist {
            {
                let s = val.to_string();
                start.push_attribute(("dist", s.as_str()));
            }
        }
        #[cfg(feature = "dml-effects")]
        if let Some(ref val) = self.dir {
            {
                let s = val.to_string();
                start.push_attribute(("dir", s.as_str()));
            }
        }
        #[cfg(feature = "dml-effects")]
        if let Some(ref val) = self.fade_dir {
            {
                let s = val.to_string();
                start.push_attribute(("fadeDir", s.as_str()));
            }
        }
        #[cfg(feature = "dml-effects")]
        if let Some(ref val) = self.sx {
            {
                let s = val.to_string();
                start.push_attribute(("sx", s.as_str()));
            }
        }
        #[cfg(feature = "dml-effects")]
        if let Some(ref val) = self.sy {
            {
                let s = val.to_string();
                start.push_attribute(("sy", s.as_str()));
            }
        }
        #[cfg(feature = "dml-effects")]
        if let Some(ref val) = self.kx {
            {
                let s = val.to_string();
                start.push_attribute(("kx", s.as_str()));
            }
        }
        #[cfg(feature = "dml-effects")]
        if let Some(ref val) = self.ky {
            {
                let s = val.to_string();
                start.push_attribute(("ky", s.as_str()));
            }
        }
        #[cfg(feature = "dml-effects")]
        if let Some(ref val) = self.algn {
            {
                let s = val.to_string();
                start.push_attribute(("algn", s.as_str()));
            }
        }
        #[cfg(feature = "dml-effects")]
        if let Some(ref val) = self.rot_with_shape {
            start.push_attribute(("rotWithShape", if *val { "1" } else { "0" }));
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

impl ToXml for CTRelativeOffsetEffect {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        #[allow(unused_mut)]
        let mut start = start;
        if let Some(ref val) = self.tx {
            {
                let s = val.to_string();
                start.push_attribute(("tx", s.as_str()));
            }
        }
        if let Some(ref val) = self.ty {
            {
                let s = val.to_string();
                start.push_attribute(("ty", s.as_str()));
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

impl ToXml for CTSoftEdgesEffect {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        #[allow(unused_mut)]
        let mut start = start;
        #[cfg(feature = "dml-effects")]
        {
            let val = &self.rad;
            {
                let s = val.to_string();
                start.push_attribute(("rad", s.as_str()));
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

impl ToXml for CTTintEffect {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        #[allow(unused_mut)]
        let mut start = start;
        if let Some(ref val) = self.hue {
            {
                let s = val.to_string();
                start.push_attribute(("hue", s.as_str()));
            }
        }
        if let Some(ref val) = self.amt {
            {
                let s = val.to_string();
                start.push_attribute(("amt", s.as_str()));
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

impl ToXml for CTTransformEffect {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        #[allow(unused_mut)]
        let mut start = start;
        if let Some(ref val) = self.sx {
            {
                let s = val.to_string();
                start.push_attribute(("sx", s.as_str()));
            }
        }
        if let Some(ref val) = self.sy {
            {
                let s = val.to_string();
                start.push_attribute(("sy", s.as_str()));
            }
        }
        if let Some(ref val) = self.kx {
            {
                let s = val.to_string();
                start.push_attribute(("kx", s.as_str()));
            }
        }
        if let Some(ref val) = self.ky {
            {
                let s = val.to_string();
                start.push_attribute(("ky", s.as_str()));
            }
        }
        if let Some(ref val) = self.tx {
            {
                let s = val.to_string();
                start.push_attribute(("tx", s.as_str()));
            }
        }
        if let Some(ref val) = self.ty {
            {
                let s = val.to_string();
                start.push_attribute(("ty", s.as_str()));
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

impl ToXml for NoFill {
    fn is_empty_element(&self) -> bool {
        true
    }
}

impl ToXml for SolidColorFill {
    fn write_children<W: Write>(&self, writer: &mut Writer<W>) -> Result<(), SerializeError> {
        #[cfg(feature = "extra-children")]
        let mut extra_iter = self.extra_children.iter().peekable();
        #[cfg(feature = "extra-children")]
        let mut emit_idx: usize = 0;
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        if let Some(ref val) = self.color_choice {
            val.write_element("", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        for extra in extra_iter {
            extra.node.write_to(writer).map_err(SerializeError::from)?;
        }
        Ok(())
    }

    fn is_empty_element(&self) -> bool {
        if self.color_choice.is_some() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for CTLinearShadeProperties {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        #[allow(unused_mut)]
        let mut start = start;
        if let Some(ref val) = self.ang {
            {
                let s = val.to_string();
                start.push_attribute(("ang", s.as_str()));
            }
        }
        if let Some(ref val) = self.scaled {
            start.push_attribute(("scaled", if *val { "1" } else { "0" }));
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

impl ToXml for CTPathShadeProperties {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        #[allow(unused_mut)]
        let mut start = start;
        if let Some(ref val) = self.path {
            {
                let s = val.to_string();
                start.push_attribute(("path", s.as_str()));
            }
        }
        #[cfg(feature = "extra-attrs")]
        for (key, value) in &self.extra_attrs {
            start.push_attribute((key.as_str(), value.as_str()));
        }
        start
    }

    fn write_children<W: Write>(&self, writer: &mut Writer<W>) -> Result<(), SerializeError> {
        #[cfg(feature = "extra-children")]
        let mut extra_iter = self.extra_children.iter().peekable();
        #[cfg(feature = "extra-children")]
        let mut emit_idx: usize = 0;
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        if let Some(ref val) = self.fill_to_rect {
            val.write_element("a:fillToRect", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        for extra in extra_iter {
            extra.node.write_to(writer).map_err(SerializeError::from)?;
        }
        Ok(())
    }

    fn is_empty_element(&self) -> bool {
        if self.fill_to_rect.is_some() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for EGShadeProperties {
    fn write_element<W: Write>(
        &self,
        _tag: &str,
        writer: &mut Writer<W>,
    ) -> Result<(), SerializeError> {
        match self {
            Self::Lin(inner) => inner.write_element("a:lin", writer)?,
            Self::Path(inner) => inner.write_element("a:path", writer)?,
        }
        Ok(())
    }
}

impl ToXml for CTGradientStop {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        #[allow(unused_mut)]
        let mut start = start;
        {
            let val = &self.pos;
            {
                let s = val.to_string();
                start.push_attribute(("pos", s.as_str()));
            }
        }
        #[cfg(feature = "extra-attrs")]
        for (key, value) in &self.extra_attrs {
            start.push_attribute((key.as_str(), value.as_str()));
        }
        start
    }

    fn write_children<W: Write>(&self, writer: &mut Writer<W>) -> Result<(), SerializeError> {
        #[cfg(feature = "extra-children")]
        let mut extra_iter = self.extra_children.iter().peekable();
        #[cfg(feature = "extra-children")]
        let mut emit_idx: usize = 0;
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        if let Some(ref val) = self.color_choice {
            val.write_element("", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        for extra in extra_iter {
            extra.node.write_to(writer).map_err(SerializeError::from)?;
        }
        Ok(())
    }

    fn is_empty_element(&self) -> bool {
        false
    }
}

impl ToXml for CTGradientStopList {
    fn write_children<W: Write>(&self, writer: &mut Writer<W>) -> Result<(), SerializeError> {
        #[cfg(feature = "extra-children")]
        let mut extra_iter = self.extra_children.iter().peekable();
        #[cfg(feature = "extra-children")]
        let mut emit_idx: usize = 0;
        for item in &self.gs {
            #[cfg(feature = "extra-children")]
            while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
                extra_iter
                    .next()
                    .unwrap()
                    .node
                    .write_to(writer)
                    .map_err(SerializeError::from)?;
            }
            item.write_element("a:gs", writer)?;
            #[cfg(feature = "extra-children")]
            {
                emit_idx += 1;
            }
        }
        #[cfg(feature = "extra-children")]
        for extra in extra_iter {
            extra.node.write_to(writer).map_err(SerializeError::from)?;
        }
        Ok(())
    }

    fn is_empty_element(&self) -> bool {
        if !self.gs.is_empty() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for GradientFill {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        #[allow(unused_mut)]
        let mut start = start;
        #[cfg(feature = "dml-fills")]
        if let Some(ref val) = self.flip {
            {
                let s = val.to_string();
                start.push_attribute(("flip", s.as_str()));
            }
        }
        #[cfg(feature = "dml-fills")]
        if let Some(ref val) = self.rot_with_shape {
            start.push_attribute(("rotWithShape", if *val { "1" } else { "0" }));
        }
        #[cfg(feature = "extra-attrs")]
        for (key, value) in &self.extra_attrs {
            start.push_attribute((key.as_str(), value.as_str()));
        }
        start
    }

    fn write_children<W: Write>(&self, writer: &mut Writer<W>) -> Result<(), SerializeError> {
        #[cfg(feature = "extra-children")]
        let mut extra_iter = self.extra_children.iter().peekable();
        #[cfg(feature = "extra-children")]
        let mut emit_idx: usize = 0;
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        #[cfg(feature = "dml-fills")]
        if let Some(ref val) = self.gs_lst {
            val.write_element("a:gsLst", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        if let Some(ref val) = self.shade_properties {
            val.write_element("", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        #[cfg(feature = "dml-fills")]
        if let Some(ref val) = self.tile_rect {
            val.write_element("a:tileRect", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        for extra in extra_iter {
            extra.node.write_to(writer).map_err(SerializeError::from)?;
        }
        Ok(())
    }

    fn is_empty_element(&self) -> bool {
        #[cfg(feature = "dml-fills")]
        if self.gs_lst.is_some() {
            return false;
        }
        if self.shade_properties.is_some() {
            return false;
        }
        #[cfg(feature = "dml-fills")]
        if self.tile_rect.is_some() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for CTTileInfoProperties {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        #[allow(unused_mut)]
        let mut start = start;
        if let Some(ref val) = self.tx {
            {
                let s = val.to_string();
                start.push_attribute(("tx", s.as_str()));
            }
        }
        if let Some(ref val) = self.ty {
            {
                let s = val.to_string();
                start.push_attribute(("ty", s.as_str()));
            }
        }
        if let Some(ref val) = self.sx {
            {
                let s = val.to_string();
                start.push_attribute(("sx", s.as_str()));
            }
        }
        if let Some(ref val) = self.sy {
            {
                let s = val.to_string();
                start.push_attribute(("sy", s.as_str()));
            }
        }
        if let Some(ref val) = self.flip {
            {
                let s = val.to_string();
                start.push_attribute(("flip", s.as_str()));
            }
        }
        if let Some(ref val) = self.algn {
            {
                let s = val.to_string();
                start.push_attribute(("algn", s.as_str()));
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

impl ToXml for CTStretchInfoProperties {
    fn write_children<W: Write>(&self, writer: &mut Writer<W>) -> Result<(), SerializeError> {
        #[cfg(feature = "extra-children")]
        let mut extra_iter = self.extra_children.iter().peekable();
        #[cfg(feature = "extra-children")]
        let mut emit_idx: usize = 0;
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        if let Some(ref val) = self.fill_rect {
            val.write_element("a:fillRect", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        for extra in extra_iter {
            extra.node.write_to(writer).map_err(SerializeError::from)?;
        }
        Ok(())
    }

    fn is_empty_element(&self) -> bool {
        if self.fill_rect.is_some() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for EGFillModeProperties {
    fn write_element<W: Write>(
        &self,
        _tag: &str,
        writer: &mut Writer<W>,
    ) -> Result<(), SerializeError> {
        match self {
            Self::Tile(inner) => inner.write_element("a:tile", writer)?,
            Self::Stretch(inner) => inner.write_element("a:stretch", writer)?,
        }
        Ok(())
    }
}

impl ToXml for Blip {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        #[allow(unused_mut)]
        let mut start = start;
        #[cfg(feature = "dml-fills")]
        if let Some(ref val) = self.embed {
            start.push_attribute(("r:embed", val.as_str()));
        }
        #[cfg(feature = "dml-fills")]
        if let Some(ref val) = self.link {
            start.push_attribute(("r:link", val.as_str()));
        }
        #[cfg(feature = "dml-fills")]
        if let Some(ref val) = self.cstate {
            {
                let s = val.to_string();
                start.push_attribute(("cstate", s.as_str()));
            }
        }
        #[cfg(feature = "extra-attrs")]
        for (key, value) in &self.extra_attrs {
            start.push_attribute((key.as_str(), value.as_str()));
        }
        start
    }

    fn write_children<W: Write>(&self, writer: &mut Writer<W>) -> Result<(), SerializeError> {
        #[cfg(feature = "extra-children")]
        let mut extra_iter = self.extra_children.iter().peekable();
        #[cfg(feature = "extra-children")]
        let mut emit_idx: usize = 0;
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        #[cfg(feature = "dml-fills")]
        if let Some(ref val) = self.alpha_bi_level {
            val.write_element("a:alphaBiLevel", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        #[cfg(feature = "dml-fills")]
        if let Some(ref val) = self.alpha_ceiling {
            val.write_element("a:alphaCeiling", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        #[cfg(feature = "dml-fills")]
        if let Some(ref val) = self.alpha_floor {
            val.write_element("a:alphaFloor", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        #[cfg(feature = "dml-fills")]
        if let Some(ref val) = self.alpha_inv {
            val.write_element("a:alphaInv", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        #[cfg(feature = "dml-fills")]
        if let Some(ref val) = self.alpha_mod {
            val.write_element("a:alphaMod", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        #[cfg(feature = "dml-fills")]
        if let Some(ref val) = self.alpha_mod_fix {
            val.write_element("a:alphaModFix", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        #[cfg(feature = "dml-fills")]
        if let Some(ref val) = self.alpha_repl {
            val.write_element("a:alphaRepl", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        #[cfg(feature = "dml-fills")]
        if let Some(ref val) = self.bi_level {
            val.write_element("a:biLevel", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        #[cfg(feature = "dml-fills")]
        if let Some(ref val) = self.blur {
            val.write_element("a:blur", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        #[cfg(feature = "dml-fills")]
        if let Some(ref val) = self.clr_change {
            val.write_element("a:clrChange", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        #[cfg(feature = "dml-fills")]
        if let Some(ref val) = self.clr_repl {
            val.write_element("a:clrRepl", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        #[cfg(feature = "dml-fills")]
        if let Some(ref val) = self.duotone {
            val.write_element("a:duotone", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        #[cfg(feature = "dml-fills")]
        if let Some(ref val) = self.fill_overlay {
            val.write_element("a:fillOverlay", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        #[cfg(feature = "dml-fills")]
        if let Some(ref val) = self.grayscl {
            val.write_element("a:grayscl", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        #[cfg(feature = "dml-fills")]
        if let Some(ref val) = self.hsl {
            val.write_element("a:hsl", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        #[cfg(feature = "dml-fills")]
        if let Some(ref val) = self.lum {
            val.write_element("a:lum", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        #[cfg(feature = "dml-fills")]
        if let Some(ref val) = self.tint {
            val.write_element("a:tint", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        #[cfg(feature = "dml-extensions")]
        if let Some(ref val) = self.ext_lst {
            val.write_element("a:extLst", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        for extra in extra_iter {
            extra.node.write_to(writer).map_err(SerializeError::from)?;
        }
        Ok(())
    }

    fn is_empty_element(&self) -> bool {
        #[cfg(feature = "dml-fills")]
        if self.alpha_bi_level.is_some() {
            return false;
        }
        #[cfg(feature = "dml-fills")]
        if self.alpha_ceiling.is_some() {
            return false;
        }
        #[cfg(feature = "dml-fills")]
        if self.alpha_floor.is_some() {
            return false;
        }
        #[cfg(feature = "dml-fills")]
        if self.alpha_inv.is_some() {
            return false;
        }
        #[cfg(feature = "dml-fills")]
        if self.alpha_mod.is_some() {
            return false;
        }
        #[cfg(feature = "dml-fills")]
        if self.alpha_mod_fix.is_some() {
            return false;
        }
        #[cfg(feature = "dml-fills")]
        if self.alpha_repl.is_some() {
            return false;
        }
        #[cfg(feature = "dml-fills")]
        if self.bi_level.is_some() {
            return false;
        }
        #[cfg(feature = "dml-fills")]
        if self.blur.is_some() {
            return false;
        }
        #[cfg(feature = "dml-fills")]
        if self.clr_change.is_some() {
            return false;
        }
        #[cfg(feature = "dml-fills")]
        if self.clr_repl.is_some() {
            return false;
        }
        #[cfg(feature = "dml-fills")]
        if self.duotone.is_some() {
            return false;
        }
        #[cfg(feature = "dml-fills")]
        if self.fill_overlay.is_some() {
            return false;
        }
        #[cfg(feature = "dml-fills")]
        if self.grayscl.is_some() {
            return false;
        }
        #[cfg(feature = "dml-fills")]
        if self.hsl.is_some() {
            return false;
        }
        #[cfg(feature = "dml-fills")]
        if self.lum.is_some() {
            return false;
        }
        #[cfg(feature = "dml-fills")]
        if self.tint.is_some() {
            return false;
        }
        #[cfg(feature = "dml-extensions")]
        if self.ext_lst.is_some() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for BlipFillProperties {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        #[allow(unused_mut)]
        let mut start = start;
        #[cfg(feature = "dml-fills")]
        if let Some(ref val) = self.dpi {
            {
                let s = val.to_string();
                start.push_attribute(("dpi", s.as_str()));
            }
        }
        #[cfg(feature = "dml-fills")]
        if let Some(ref val) = self.rot_with_shape {
            start.push_attribute(("rotWithShape", if *val { "1" } else { "0" }));
        }
        #[cfg(feature = "extra-attrs")]
        for (key, value) in &self.extra_attrs {
            start.push_attribute((key.as_str(), value.as_str()));
        }
        start
    }

    fn write_children<W: Write>(&self, writer: &mut Writer<W>) -> Result<(), SerializeError> {
        #[cfg(feature = "extra-children")]
        let mut extra_iter = self.extra_children.iter().peekable();
        #[cfg(feature = "extra-children")]
        let mut emit_idx: usize = 0;
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        #[cfg(feature = "dml-fills")]
        if let Some(ref val) = self.blip {
            val.write_element("a:blip", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        #[cfg(feature = "dml-fills")]
        if let Some(ref val) = self.src_rect {
            val.write_element("a:srcRect", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        if let Some(ref val) = self.fill_mode_properties {
            val.write_element("", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        for extra in extra_iter {
            extra.node.write_to(writer).map_err(SerializeError::from)?;
        }
        Ok(())
    }

    fn is_empty_element(&self) -> bool {
        #[cfg(feature = "dml-fills")]
        if self.blip.is_some() {
            return false;
        }
        #[cfg(feature = "dml-fills")]
        if self.src_rect.is_some() {
            return false;
        }
        if self.fill_mode_properties.is_some() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for PatternFill {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        #[allow(unused_mut)]
        let mut start = start;
        #[cfg(feature = "dml-fills")]
        if let Some(ref val) = self.preset {
            {
                let s = val.to_string();
                start.push_attribute(("prst", s.as_str()));
            }
        }
        #[cfg(feature = "extra-attrs")]
        for (key, value) in &self.extra_attrs {
            start.push_attribute((key.as_str(), value.as_str()));
        }
        start
    }

    fn write_children<W: Write>(&self, writer: &mut Writer<W>) -> Result<(), SerializeError> {
        #[cfg(feature = "extra-children")]
        let mut extra_iter = self.extra_children.iter().peekable();
        #[cfg(feature = "extra-children")]
        let mut emit_idx: usize = 0;
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        #[cfg(feature = "dml-fills")]
        if let Some(ref val) = self.fg_clr {
            val.write_element("a:fgClr", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        #[cfg(feature = "dml-fills")]
        if let Some(ref val) = self.bg_clr {
            val.write_element("a:bgClr", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        for extra in extra_iter {
            extra.node.write_to(writer).map_err(SerializeError::from)?;
        }
        Ok(())
    }

    fn is_empty_element(&self) -> bool {
        #[cfg(feature = "dml-fills")]
        if self.fg_clr.is_some() {
            return false;
        }
        #[cfg(feature = "dml-fills")]
        if self.bg_clr.is_some() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for CTGroupFillProperties {
    fn is_empty_element(&self) -> bool {
        true
    }
}

impl ToXml for EGFillProperties {
    fn write_element<W: Write>(
        &self,
        _tag: &str,
        writer: &mut Writer<W>,
    ) -> Result<(), SerializeError> {
        match self {
            Self::NoFill(inner) => inner.write_element("a:noFill", writer)?,
            Self::SolidFill(inner) => inner.write_element("a:solidFill", writer)?,
            Self::GradFill(inner) => inner.write_element("a:gradFill", writer)?,
            Self::BlipFill(inner) => inner.write_element("a:blipFill", writer)?,
            Self::PattFill(inner) => inner.write_element("a:pattFill", writer)?,
            Self::GrpFill(inner) => inner.write_element("a:grpFill", writer)?,
        }
        Ok(())
    }
}

impl ToXml for CTFillProperties {
    fn write_children<W: Write>(&self, writer: &mut Writer<W>) -> Result<(), SerializeError> {
        #[cfg(feature = "extra-children")]
        let mut extra_iter = self.extra_children.iter().peekable();
        #[cfg(feature = "extra-children")]
        let mut emit_idx: usize = 0;
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        if let Some(ref val) = self.fill_properties {
            val.write_element("", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        for extra in extra_iter {
            extra.node.write_to(writer).map_err(SerializeError::from)?;
        }
        Ok(())
    }

    fn is_empty_element(&self) -> bool {
        false
    }
}

impl ToXml for CTFillEffect {
    fn write_children<W: Write>(&self, writer: &mut Writer<W>) -> Result<(), SerializeError> {
        #[cfg(feature = "extra-children")]
        let mut extra_iter = self.extra_children.iter().peekable();
        #[cfg(feature = "extra-children")]
        let mut emit_idx: usize = 0;
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        if let Some(ref val) = self.fill_properties {
            val.write_element("", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        for extra in extra_iter {
            extra.node.write_to(writer).map_err(SerializeError::from)?;
        }
        Ok(())
    }

    fn is_empty_element(&self) -> bool {
        false
    }
}

impl ToXml for CTFillOverlayEffect {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        #[allow(unused_mut)]
        let mut start = start;
        {
            let val = &self.blend;
            {
                let s = val.to_string();
                start.push_attribute(("blend", s.as_str()));
            }
        }
        #[cfg(feature = "extra-attrs")]
        for (key, value) in &self.extra_attrs {
            start.push_attribute((key.as_str(), value.as_str()));
        }
        start
    }

    fn write_children<W: Write>(&self, writer: &mut Writer<W>) -> Result<(), SerializeError> {
        #[cfg(feature = "extra-children")]
        let mut extra_iter = self.extra_children.iter().peekable();
        #[cfg(feature = "extra-children")]
        let mut emit_idx: usize = 0;
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        if let Some(ref val) = self.fill_properties {
            val.write_element("", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        for extra in extra_iter {
            extra.node.write_to(writer).map_err(SerializeError::from)?;
        }
        Ok(())
    }

    fn is_empty_element(&self) -> bool {
        false
    }
}

impl ToXml for CTEffectReference {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        #[allow(unused_mut)]
        let mut start = start;
        {
            let val = &self.r#ref;
            start.push_attribute(("ref", val.as_str()));
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

impl ToXml for EGEffect {
    fn write_element<W: Write>(
        &self,
        _tag: &str,
        writer: &mut Writer<W>,
    ) -> Result<(), SerializeError> {
        match self {
            Self::Cont(inner) => inner.write_element("a:cont", writer)?,
            Self::Effect(inner) => inner.write_element("a:effect", writer)?,
            Self::AlphaBiLevel(inner) => inner.write_element("a:alphaBiLevel", writer)?,
            Self::AlphaCeiling(inner) => inner.write_element("a:alphaCeiling", writer)?,
            Self::AlphaFloor(inner) => inner.write_element("a:alphaFloor", writer)?,
            Self::AlphaInv(inner) => inner.write_element("a:alphaInv", writer)?,
            Self::AlphaMod(inner) => inner.write_element("a:alphaMod", writer)?,
            Self::AlphaModFix(inner) => inner.write_element("a:alphaModFix", writer)?,
            Self::AlphaOutset(inner) => inner.write_element("a:alphaOutset", writer)?,
            Self::AlphaRepl(inner) => inner.write_element("a:alphaRepl", writer)?,
            Self::BiLevel(inner) => inner.write_element("a:biLevel", writer)?,
            Self::Blend(inner) => inner.write_element("a:blend", writer)?,
            Self::Blur(inner) => inner.write_element("a:blur", writer)?,
            Self::ClrChange(inner) => inner.write_element("a:clrChange", writer)?,
            Self::ClrRepl(inner) => inner.write_element("a:clrRepl", writer)?,
            Self::Duotone(inner) => inner.write_element("a:duotone", writer)?,
            Self::Fill(inner) => inner.write_element("a:fill", writer)?,
            Self::FillOverlay(inner) => inner.write_element("a:fillOverlay", writer)?,
            Self::Glow(inner) => inner.write_element("a:glow", writer)?,
            Self::Grayscl(inner) => inner.write_element("a:grayscl", writer)?,
            Self::Hsl(inner) => inner.write_element("a:hsl", writer)?,
            Self::InnerShdw(inner) => inner.write_element("a:innerShdw", writer)?,
            Self::Lum(inner) => inner.write_element("a:lum", writer)?,
            Self::OuterShdw(inner) => inner.write_element("a:outerShdw", writer)?,
            Self::PrstShdw(inner) => inner.write_element("a:prstShdw", writer)?,
            Self::Reflection(inner) => inner.write_element("a:reflection", writer)?,
            Self::RelOff(inner) => inner.write_element("a:relOff", writer)?,
            Self::SoftEdge(inner) => inner.write_element("a:softEdge", writer)?,
            Self::Tint(inner) => inner.write_element("a:tint", writer)?,
            Self::Xfrm(inner) => inner.write_element("a:xfrm", writer)?,
        }
        Ok(())
    }
}

impl ToXml for EffectContainer {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        #[allow(unused_mut)]
        let mut start = start;
        if let Some(ref val) = self.r#type {
            {
                let s = val.to_string();
                start.push_attribute(("type", s.as_str()));
            }
        }
        if let Some(ref val) = self.name {
            start.push_attribute(("name", val.as_str()));
        }
        #[cfg(feature = "extra-attrs")]
        for (key, value) in &self.extra_attrs {
            start.push_attribute((key.as_str(), value.as_str()));
        }
        start
    }

    fn write_children<W: Write>(&self, writer: &mut Writer<W>) -> Result<(), SerializeError> {
        #[cfg(feature = "extra-children")]
        let mut extra_iter = self.extra_children.iter().peekable();
        #[cfg(feature = "extra-children")]
        let mut emit_idx: usize = 0;
        for item in &self.effect {
            #[cfg(feature = "extra-children")]
            while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
                extra_iter
                    .next()
                    .unwrap()
                    .node
                    .write_to(writer)
                    .map_err(SerializeError::from)?;
            }
            item.write_element("", writer)?;
            #[cfg(feature = "extra-children")]
            {
                emit_idx += 1;
            }
        }
        #[cfg(feature = "extra-children")]
        for extra in extra_iter {
            extra.node.write_to(writer).map_err(SerializeError::from)?;
        }
        Ok(())
    }

    fn is_empty_element(&self) -> bool {
        if !self.effect.is_empty() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for CTBlendEffect {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        #[allow(unused_mut)]
        let mut start = start;
        {
            let val = &self.blend;
            {
                let s = val.to_string();
                start.push_attribute(("blend", s.as_str()));
            }
        }
        #[cfg(feature = "extra-attrs")]
        for (key, value) in &self.extra_attrs {
            start.push_attribute((key.as_str(), value.as_str()));
        }
        start
    }

    fn write_children<W: Write>(&self, writer: &mut Writer<W>) -> Result<(), SerializeError> {
        #[cfg(feature = "extra-children")]
        let mut extra_iter = self.extra_children.iter().peekable();
        #[cfg(feature = "extra-children")]
        let mut emit_idx: usize = 0;
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        {
            let val = &self.cont;
            val.write_element("a:cont", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        for extra in extra_iter {
            extra.node.write_to(writer).map_err(SerializeError::from)?;
        }
        Ok(())
    }

    fn is_empty_element(&self) -> bool {
        false
    }
}

impl ToXml for EffectList {
    fn write_children<W: Write>(&self, writer: &mut Writer<W>) -> Result<(), SerializeError> {
        #[cfg(feature = "extra-children")]
        let mut extra_iter = self.extra_children.iter().peekable();
        #[cfg(feature = "extra-children")]
        let mut emit_idx: usize = 0;
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        #[cfg(feature = "dml-effects")]
        if let Some(ref val) = self.blur {
            val.write_element("a:blur", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        #[cfg(feature = "dml-effects")]
        if let Some(ref val) = self.fill_overlay {
            val.write_element("a:fillOverlay", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        #[cfg(feature = "dml-effects")]
        if let Some(ref val) = self.glow {
            val.write_element("a:glow", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        #[cfg(feature = "dml-effects")]
        if let Some(ref val) = self.inner_shdw {
            val.write_element("a:innerShdw", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        #[cfg(feature = "dml-effects")]
        if let Some(ref val) = self.outer_shdw {
            val.write_element("a:outerShdw", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        #[cfg(feature = "dml-effects")]
        if let Some(ref val) = self.prst_shdw {
            val.write_element("a:prstShdw", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        #[cfg(feature = "dml-effects")]
        if let Some(ref val) = self.reflection {
            val.write_element("a:reflection", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        #[cfg(feature = "dml-effects")]
        if let Some(ref val) = self.soft_edge {
            val.write_element("a:softEdge", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        for extra in extra_iter {
            extra.node.write_to(writer).map_err(SerializeError::from)?;
        }
        Ok(())
    }

    fn is_empty_element(&self) -> bool {
        #[cfg(feature = "dml-effects")]
        if self.blur.is_some() {
            return false;
        }
        #[cfg(feature = "dml-effects")]
        if self.fill_overlay.is_some() {
            return false;
        }
        #[cfg(feature = "dml-effects")]
        if self.glow.is_some() {
            return false;
        }
        #[cfg(feature = "dml-effects")]
        if self.inner_shdw.is_some() {
            return false;
        }
        #[cfg(feature = "dml-effects")]
        if self.outer_shdw.is_some() {
            return false;
        }
        #[cfg(feature = "dml-effects")]
        if self.prst_shdw.is_some() {
            return false;
        }
        #[cfg(feature = "dml-effects")]
        if self.reflection.is_some() {
            return false;
        }
        #[cfg(feature = "dml-effects")]
        if self.soft_edge.is_some() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for EGEffectProperties {
    fn write_element<W: Write>(
        &self,
        _tag: &str,
        writer: &mut Writer<W>,
    ) -> Result<(), SerializeError> {
        match self {
            Self::EffectLst(inner) => inner.write_element("a:effectLst", writer)?,
            Self::EffectDag(inner) => inner.write_element("a:effectDag", writer)?,
        }
        Ok(())
    }
}

impl ToXml for CTEffectProperties {
    fn write_children<W: Write>(&self, writer: &mut Writer<W>) -> Result<(), SerializeError> {
        #[cfg(feature = "extra-children")]
        let mut extra_iter = self.extra_children.iter().peekable();
        #[cfg(feature = "extra-children")]
        let mut emit_idx: usize = 0;
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        if let Some(ref val) = self.effect_properties {
            val.write_element("", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        for extra in extra_iter {
            extra.node.write_to(writer).map_err(SerializeError::from)?;
        }
        Ok(())
    }

    fn is_empty_element(&self) -> bool {
        false
    }
}

impl ToXml for CTGeomGuide {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        #[allow(unused_mut)]
        let mut start = start;
        {
            let val = &self.name;
            start.push_attribute(("name", val.as_str()));
        }
        {
            let val = &self.fmla;
            start.push_attribute(("fmla", val.as_str()));
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

impl ToXml for CTGeomGuideList {
    fn write_children<W: Write>(&self, writer: &mut Writer<W>) -> Result<(), SerializeError> {
        #[cfg(feature = "extra-children")]
        let mut extra_iter = self.extra_children.iter().peekable();
        #[cfg(feature = "extra-children")]
        let mut emit_idx: usize = 0;
        for item in &self.gd {
            #[cfg(feature = "extra-children")]
            while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
                extra_iter
                    .next()
                    .unwrap()
                    .node
                    .write_to(writer)
                    .map_err(SerializeError::from)?;
            }
            item.write_element("a:gd", writer)?;
            #[cfg(feature = "extra-children")]
            {
                emit_idx += 1;
            }
        }
        #[cfg(feature = "extra-children")]
        for extra in extra_iter {
            extra.node.write_to(writer).map_err(SerializeError::from)?;
        }
        Ok(())
    }

    fn is_empty_element(&self) -> bool {
        if !self.gd.is_empty() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for CTAdjPoint2D {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        #[allow(unused_mut)]
        let mut start = start;
        {
            let val = &self.x;
            {
                let s = val.to_string();
                start.push_attribute(("x", s.as_str()));
            }
        }
        {
            let val = &self.y;
            {
                let s = val.to_string();
                start.push_attribute(("y", s.as_str()));
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

impl ToXml for CTGeomRect {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        #[allow(unused_mut)]
        let mut start = start;
        {
            let val = &self.l;
            {
                let s = val.to_string();
                start.push_attribute(("l", s.as_str()));
            }
        }
        {
            let val = &self.t;
            {
                let s = val.to_string();
                start.push_attribute(("t", s.as_str()));
            }
        }
        {
            let val = &self.relationship_id;
            {
                let s = val.to_string();
                start.push_attribute(("r", s.as_str()));
            }
        }
        {
            let val = &self.b;
            {
                let s = val.to_string();
                start.push_attribute(("b", s.as_str()));
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

impl ToXml for CTXYAdjustHandle {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        #[allow(unused_mut)]
        let mut start = start;
        if let Some(ref val) = self.gd_ref_x {
            start.push_attribute(("gdRefX", val.as_str()));
        }
        if let Some(ref val) = self.min_x {
            {
                let s = val.to_string();
                start.push_attribute(("minX", s.as_str()));
            }
        }
        if let Some(ref val) = self.max_x {
            {
                let s = val.to_string();
                start.push_attribute(("maxX", s.as_str()));
            }
        }
        if let Some(ref val) = self.gd_ref_y {
            start.push_attribute(("gdRefY", val.as_str()));
        }
        if let Some(ref val) = self.min_y {
            {
                let s = val.to_string();
                start.push_attribute(("minY", s.as_str()));
            }
        }
        if let Some(ref val) = self.max_y {
            {
                let s = val.to_string();
                start.push_attribute(("maxY", s.as_str()));
            }
        }
        #[cfg(feature = "extra-attrs")]
        for (key, value) in &self.extra_attrs {
            start.push_attribute((key.as_str(), value.as_str()));
        }
        start
    }

    fn write_children<W: Write>(&self, writer: &mut Writer<W>) -> Result<(), SerializeError> {
        #[cfg(feature = "extra-children")]
        let mut extra_iter = self.extra_children.iter().peekable();
        #[cfg(feature = "extra-children")]
        let mut emit_idx: usize = 0;
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        {
            let val = &self.pos;
            val.write_element("a:pos", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        for extra in extra_iter {
            extra.node.write_to(writer).map_err(SerializeError::from)?;
        }
        Ok(())
    }

    fn is_empty_element(&self) -> bool {
        false
    }
}

impl ToXml for CTPolarAdjustHandle {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        #[allow(unused_mut)]
        let mut start = start;
        if let Some(ref val) = self.gd_ref_r {
            start.push_attribute(("gdRefR", val.as_str()));
        }
        if let Some(ref val) = self.min_r {
            {
                let s = val.to_string();
                start.push_attribute(("minR", s.as_str()));
            }
        }
        if let Some(ref val) = self.max_r {
            {
                let s = val.to_string();
                start.push_attribute(("maxR", s.as_str()));
            }
        }
        if let Some(ref val) = self.gd_ref_ang {
            start.push_attribute(("gdRefAng", val.as_str()));
        }
        if let Some(ref val) = self.min_ang {
            {
                let s = val.to_string();
                start.push_attribute(("minAng", s.as_str()));
            }
        }
        if let Some(ref val) = self.max_ang {
            {
                let s = val.to_string();
                start.push_attribute(("maxAng", s.as_str()));
            }
        }
        #[cfg(feature = "extra-attrs")]
        for (key, value) in &self.extra_attrs {
            start.push_attribute((key.as_str(), value.as_str()));
        }
        start
    }

    fn write_children<W: Write>(&self, writer: &mut Writer<W>) -> Result<(), SerializeError> {
        #[cfg(feature = "extra-children")]
        let mut extra_iter = self.extra_children.iter().peekable();
        #[cfg(feature = "extra-children")]
        let mut emit_idx: usize = 0;
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        {
            let val = &self.pos;
            val.write_element("a:pos", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        for extra in extra_iter {
            extra.node.write_to(writer).map_err(SerializeError::from)?;
        }
        Ok(())
    }

    fn is_empty_element(&self) -> bool {
        false
    }
}

impl ToXml for CTConnectionSite {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        #[allow(unused_mut)]
        let mut start = start;
        {
            let val = &self.ang;
            {
                let s = val.to_string();
                start.push_attribute(("ang", s.as_str()));
            }
        }
        #[cfg(feature = "extra-attrs")]
        for (key, value) in &self.extra_attrs {
            start.push_attribute((key.as_str(), value.as_str()));
        }
        start
    }

    fn write_children<W: Write>(&self, writer: &mut Writer<W>) -> Result<(), SerializeError> {
        #[cfg(feature = "extra-children")]
        let mut extra_iter = self.extra_children.iter().peekable();
        #[cfg(feature = "extra-children")]
        let mut emit_idx: usize = 0;
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        {
            let val = &self.pos;
            val.write_element("a:pos", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        for extra in extra_iter {
            extra.node.write_to(writer).map_err(SerializeError::from)?;
        }
        Ok(())
    }

    fn is_empty_element(&self) -> bool {
        false
    }
}

impl ToXml for CTAdjustHandleList {
    fn write_children<W: Write>(&self, writer: &mut Writer<W>) -> Result<(), SerializeError> {
        #[cfg(feature = "extra-children")]
        let mut extra_iter = self.extra_children.iter().peekable();
        #[cfg(feature = "extra-children")]
        let mut emit_idx: usize = 0;
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        if let Some(ref val) = self.ah_x_y {
            val.write_element("a:ahXY", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        if let Some(ref val) = self.ah_polar {
            val.write_element("a:ahPolar", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        for extra in extra_iter {
            extra.node.write_to(writer).map_err(SerializeError::from)?;
        }
        Ok(())
    }

    fn is_empty_element(&self) -> bool {
        if self.ah_x_y.is_some() {
            return false;
        }
        if self.ah_polar.is_some() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for CTConnectionSiteList {
    fn write_children<W: Write>(&self, writer: &mut Writer<W>) -> Result<(), SerializeError> {
        #[cfg(feature = "extra-children")]
        let mut extra_iter = self.extra_children.iter().peekable();
        #[cfg(feature = "extra-children")]
        let mut emit_idx: usize = 0;
        for item in &self.cxn {
            #[cfg(feature = "extra-children")]
            while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
                extra_iter
                    .next()
                    .unwrap()
                    .node
                    .write_to(writer)
                    .map_err(SerializeError::from)?;
            }
            item.write_element("a:cxn", writer)?;
            #[cfg(feature = "extra-children")]
            {
                emit_idx += 1;
            }
        }
        #[cfg(feature = "extra-children")]
        for extra in extra_iter {
            extra.node.write_to(writer).map_err(SerializeError::from)?;
        }
        Ok(())
    }

    fn is_empty_element(&self) -> bool {
        if !self.cxn.is_empty() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for CTConnection {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        #[allow(unused_mut)]
        let mut start = start;
        {
            let val = &self.id;
            {
                let s = val.to_string();
                start.push_attribute(("id", s.as_str()));
            }
        }
        {
            let val = &self.idx;
            {
                let s = val.to_string();
                start.push_attribute(("idx", s.as_str()));
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

impl ToXml for CTPath2DArcTo {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        #[allow(unused_mut)]
        let mut start = start;
        {
            let val = &self.w_r;
            {
                let s = val.to_string();
                start.push_attribute(("wR", s.as_str()));
            }
        }
        {
            let val = &self.h_r;
            {
                let s = val.to_string();
                start.push_attribute(("hR", s.as_str()));
            }
        }
        {
            let val = &self.st_ang;
            {
                let s = val.to_string();
                start.push_attribute(("stAng", s.as_str()));
            }
        }
        {
            let val = &self.sw_ang;
            {
                let s = val.to_string();
                start.push_attribute(("swAng", s.as_str()));
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

impl ToXml for CTPath2DQuadBezierTo {
    fn write_children<W: Write>(&self, writer: &mut Writer<W>) -> Result<(), SerializeError> {
        #[cfg(feature = "extra-children")]
        let mut extra_iter = self.extra_children.iter().peekable();
        #[cfg(feature = "extra-children")]
        let mut emit_idx: usize = 0;
        for item in &self.pt {
            #[cfg(feature = "extra-children")]
            while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
                extra_iter
                    .next()
                    .unwrap()
                    .node
                    .write_to(writer)
                    .map_err(SerializeError::from)?;
            }
            item.write_element("a:pt", writer)?;
            #[cfg(feature = "extra-children")]
            {
                emit_idx += 1;
            }
        }
        #[cfg(feature = "extra-children")]
        for extra in extra_iter {
            extra.node.write_to(writer).map_err(SerializeError::from)?;
        }
        Ok(())
    }

    fn is_empty_element(&self) -> bool {
        if !self.pt.is_empty() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for CTPath2DCubicBezierTo {
    fn write_children<W: Write>(&self, writer: &mut Writer<W>) -> Result<(), SerializeError> {
        #[cfg(feature = "extra-children")]
        let mut extra_iter = self.extra_children.iter().peekable();
        #[cfg(feature = "extra-children")]
        let mut emit_idx: usize = 0;
        for item in &self.pt {
            #[cfg(feature = "extra-children")]
            while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
                extra_iter
                    .next()
                    .unwrap()
                    .node
                    .write_to(writer)
                    .map_err(SerializeError::from)?;
            }
            item.write_element("a:pt", writer)?;
            #[cfg(feature = "extra-children")]
            {
                emit_idx += 1;
            }
        }
        #[cfg(feature = "extra-children")]
        for extra in extra_iter {
            extra.node.write_to(writer).map_err(SerializeError::from)?;
        }
        Ok(())
    }

    fn is_empty_element(&self) -> bool {
        if !self.pt.is_empty() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for CTPath2DClose {
    fn is_empty_element(&self) -> bool {
        true
    }
}

impl ToXml for CTPath2D {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        #[allow(unused_mut)]
        let mut start = start;
        if let Some(ref val) = self.width {
            {
                let s = val.to_string();
                start.push_attribute(("w", s.as_str()));
            }
        }
        if let Some(ref val) = self.height {
            {
                let s = val.to_string();
                start.push_attribute(("h", s.as_str()));
            }
        }
        if let Some(ref val) = self.fill {
            {
                let s = val.to_string();
                start.push_attribute(("fill", s.as_str()));
            }
        }
        if let Some(ref val) = self.stroke {
            start.push_attribute(("stroke", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.extrusion_ok {
            start.push_attribute(("extrusionOk", if *val { "1" } else { "0" }));
        }
        #[cfg(feature = "extra-attrs")]
        for (key, value) in &self.extra_attrs {
            start.push_attribute((key.as_str(), value.as_str()));
        }
        start
    }

    fn write_children<W: Write>(&self, writer: &mut Writer<W>) -> Result<(), SerializeError> {
        #[cfg(feature = "extra-children")]
        let mut extra_iter = self.extra_children.iter().peekable();
        #[cfg(feature = "extra-children")]
        let mut emit_idx: usize = 0;
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        if let Some(ref val) = self.close {
            val.write_element("a:close", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        if let Some(ref val) = self.move_to {
            val.write_element("a:moveTo", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        if let Some(ref val) = self.ln_to {
            val.write_element("a:lnTo", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        if let Some(ref val) = self.arc_to {
            val.write_element("a:arcTo", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        if let Some(ref val) = self.quad_bez_to {
            val.write_element("a:quadBezTo", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        if let Some(ref val) = self.cubic_bez_to {
            val.write_element("a:cubicBezTo", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        for extra in extra_iter {
            extra.node.write_to(writer).map_err(SerializeError::from)?;
        }
        Ok(())
    }

    fn is_empty_element(&self) -> bool {
        if self.close.is_some() {
            return false;
        }
        if self.move_to.is_some() {
            return false;
        }
        if self.ln_to.is_some() {
            return false;
        }
        if self.arc_to.is_some() {
            return false;
        }
        if self.quad_bez_to.is_some() {
            return false;
        }
        if self.cubic_bez_to.is_some() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for CTPath2DList {
    fn write_children<W: Write>(&self, writer: &mut Writer<W>) -> Result<(), SerializeError> {
        #[cfg(feature = "extra-children")]
        let mut extra_iter = self.extra_children.iter().peekable();
        #[cfg(feature = "extra-children")]
        let mut emit_idx: usize = 0;
        for item in &self.path {
            #[cfg(feature = "extra-children")]
            while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
                extra_iter
                    .next()
                    .unwrap()
                    .node
                    .write_to(writer)
                    .map_err(SerializeError::from)?;
            }
            item.write_element("a:path", writer)?;
            #[cfg(feature = "extra-children")]
            {
                emit_idx += 1;
            }
        }
        #[cfg(feature = "extra-children")]
        for extra in extra_iter {
            extra.node.write_to(writer).map_err(SerializeError::from)?;
        }
        Ok(())
    }

    fn is_empty_element(&self) -> bool {
        if !self.path.is_empty() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for CTPresetGeometry2D {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        #[allow(unused_mut)]
        let mut start = start;
        {
            let val = &self.preset;
            {
                let s = val.to_string();
                start.push_attribute(("prst", s.as_str()));
            }
        }
        #[cfg(feature = "extra-attrs")]
        for (key, value) in &self.extra_attrs {
            start.push_attribute((key.as_str(), value.as_str()));
        }
        start
    }

    fn write_children<W: Write>(&self, writer: &mut Writer<W>) -> Result<(), SerializeError> {
        #[cfg(feature = "extra-children")]
        let mut extra_iter = self.extra_children.iter().peekable();
        #[cfg(feature = "extra-children")]
        let mut emit_idx: usize = 0;
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        if let Some(ref val) = self.av_lst {
            val.write_element("a:avLst", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        for extra in extra_iter {
            extra.node.write_to(writer).map_err(SerializeError::from)?;
        }
        Ok(())
    }

    fn is_empty_element(&self) -> bool {
        if self.av_lst.is_some() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for CTPresetTextShape {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        #[allow(unused_mut)]
        let mut start = start;
        {
            let val = &self.preset;
            {
                let s = val.to_string();
                start.push_attribute(("prst", s.as_str()));
            }
        }
        #[cfg(feature = "extra-attrs")]
        for (key, value) in &self.extra_attrs {
            start.push_attribute((key.as_str(), value.as_str()));
        }
        start
    }

    fn write_children<W: Write>(&self, writer: &mut Writer<W>) -> Result<(), SerializeError> {
        #[cfg(feature = "extra-children")]
        let mut extra_iter = self.extra_children.iter().peekable();
        #[cfg(feature = "extra-children")]
        let mut emit_idx: usize = 0;
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        if let Some(ref val) = self.av_lst {
            val.write_element("a:avLst", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        for extra in extra_iter {
            extra.node.write_to(writer).map_err(SerializeError::from)?;
        }
        Ok(())
    }

    fn is_empty_element(&self) -> bool {
        if self.av_lst.is_some() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for CTCustomGeometry2D {
    fn write_children<W: Write>(&self, writer: &mut Writer<W>) -> Result<(), SerializeError> {
        #[cfg(feature = "extra-children")]
        let mut extra_iter = self.extra_children.iter().peekable();
        #[cfg(feature = "extra-children")]
        let mut emit_idx: usize = 0;
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        if let Some(ref val) = self.av_lst {
            val.write_element("a:avLst", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        if let Some(ref val) = self.gd_lst {
            val.write_element("a:gdLst", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        if let Some(ref val) = self.ah_lst {
            val.write_element("a:ahLst", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        if let Some(ref val) = self.cxn_lst {
            val.write_element("a:cxnLst", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        if let Some(ref val) = self.rect {
            val.write_element("a:rect", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        {
            let val = &self.path_lst;
            val.write_element("a:pathLst", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        for extra in extra_iter {
            extra.node.write_to(writer).map_err(SerializeError::from)?;
        }
        Ok(())
    }

    fn is_empty_element(&self) -> bool {
        if self.av_lst.is_some() {
            return false;
        }
        if self.gd_lst.is_some() {
            return false;
        }
        if self.ah_lst.is_some() {
            return false;
        }
        if self.cxn_lst.is_some() {
            return false;
        }
        if self.rect.is_some() {
            return false;
        }
        false
    }
}

impl ToXml for EGGeometry {
    fn write_element<W: Write>(
        &self,
        _tag: &str,
        writer: &mut Writer<W>,
    ) -> Result<(), SerializeError> {
        match self {
            Self::CustGeom(inner) => inner.write_element("a:custGeom", writer)?,
            Self::PrstGeom(inner) => inner.write_element("a:prstGeom", writer)?,
        }
        Ok(())
    }
}

impl ToXml for EGTextGeometry {
    fn write_element<W: Write>(
        &self,
        _tag: &str,
        writer: &mut Writer<W>,
    ) -> Result<(), SerializeError> {
        match self {
            Self::CustGeom(inner) => inner.write_element("a:custGeom", writer)?,
            Self::PrstTxWarp(inner) => inner.write_element("a:prstTxWarp", writer)?,
        }
        Ok(())
    }
}

impl ToXml for CTLineEndProperties {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        #[allow(unused_mut)]
        let mut start = start;
        #[cfg(feature = "dml-lines")]
        if let Some(ref val) = self.r#type {
            {
                let s = val.to_string();
                start.push_attribute(("type", s.as_str()));
            }
        }
        #[cfg(feature = "dml-lines")]
        if let Some(ref val) = self.width {
            {
                let s = val.to_string();
                start.push_attribute(("w", s.as_str()));
            }
        }
        #[cfg(feature = "dml-lines")]
        if let Some(ref val) = self.len {
            {
                let s = val.to_string();
                start.push_attribute(("len", s.as_str()));
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

impl ToXml for EGLineFillProperties {
    fn write_element<W: Write>(
        &self,
        _tag: &str,
        writer: &mut Writer<W>,
    ) -> Result<(), SerializeError> {
        match self {
            Self::NoFill(inner) => inner.write_element("a:noFill", writer)?,
            Self::SolidFill(inner) => inner.write_element("a:solidFill", writer)?,
            Self::GradFill(inner) => inner.write_element("a:gradFill", writer)?,
            Self::PattFill(inner) => inner.write_element("a:pattFill", writer)?,
        }
        Ok(())
    }
}

impl ToXml for CTLineJoinBevel {
    fn is_empty_element(&self) -> bool {
        true
    }
}

impl ToXml for CTLineJoinRound {
    fn is_empty_element(&self) -> bool {
        true
    }
}

impl ToXml for CTLineJoinMiterProperties {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        #[allow(unused_mut)]
        let mut start = start;
        if let Some(ref val) = self.lim {
            {
                let s = val.to_string();
                start.push_attribute(("lim", s.as_str()));
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

impl ToXml for EGLineJoinProperties {
    fn write_element<W: Write>(
        &self,
        _tag: &str,
        writer: &mut Writer<W>,
    ) -> Result<(), SerializeError> {
        match self {
            Self::Round(inner) => inner.write_element("a:round", writer)?,
            Self::Bevel(inner) => inner.write_element("a:bevel", writer)?,
            Self::Miter(inner) => inner.write_element("a:miter", writer)?,
        }
        Ok(())
    }
}

impl ToXml for CTPresetLineDashProperties {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        #[allow(unused_mut)]
        let mut start = start;
        #[cfg(feature = "dml-lines")]
        if let Some(ref val) = self.value {
            {
                let s = val.to_string();
                start.push_attribute(("val", s.as_str()));
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

impl ToXml for CTDashStop {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        #[allow(unused_mut)]
        let mut start = start;
        #[cfg(feature = "dml-lines")]
        {
            let val = &self.d;
            {
                let s = val.to_string();
                start.push_attribute(("d", s.as_str()));
            }
        }
        #[cfg(feature = "dml-lines")]
        {
            let val = &self.sp;
            {
                let s = val.to_string();
                start.push_attribute(("sp", s.as_str()));
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

impl ToXml for CTDashStopList {
    fn write_children<W: Write>(&self, writer: &mut Writer<W>) -> Result<(), SerializeError> {
        #[cfg(feature = "extra-children")]
        let mut extra_iter = self.extra_children.iter().peekable();
        #[cfg(feature = "extra-children")]
        let mut emit_idx: usize = 0;
        for item in &self.ds {
            #[cfg(feature = "extra-children")]
            while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
                extra_iter
                    .next()
                    .unwrap()
                    .node
                    .write_to(writer)
                    .map_err(SerializeError::from)?;
            }
            item.write_element("a:ds", writer)?;
            #[cfg(feature = "extra-children")]
            {
                emit_idx += 1;
            }
        }
        #[cfg(feature = "extra-children")]
        for extra in extra_iter {
            extra.node.write_to(writer).map_err(SerializeError::from)?;
        }
        Ok(())
    }

    fn is_empty_element(&self) -> bool {
        if !self.ds.is_empty() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for EGLineDashProperties {
    fn write_element<W: Write>(
        &self,
        _tag: &str,
        writer: &mut Writer<W>,
    ) -> Result<(), SerializeError> {
        match self {
            Self::PrstDash(inner) => inner.write_element("a:prstDash", writer)?,
            Self::CustDash(inner) => inner.write_element("a:custDash", writer)?,
        }
        Ok(())
    }
}

impl ToXml for LineProperties {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        #[allow(unused_mut)]
        let mut start = start;
        #[cfg(feature = "dml-lines")]
        if let Some(ref val) = self.width {
            {
                let s = val.to_string();
                start.push_attribute(("w", s.as_str()));
            }
        }
        #[cfg(feature = "dml-lines")]
        if let Some(ref val) = self.cap {
            {
                let s = val.to_string();
                start.push_attribute(("cap", s.as_str()));
            }
        }
        #[cfg(feature = "dml-lines")]
        if let Some(ref val) = self.cmpd {
            {
                let s = val.to_string();
                start.push_attribute(("cmpd", s.as_str()));
            }
        }
        #[cfg(feature = "dml-lines")]
        if let Some(ref val) = self.algn {
            {
                let s = val.to_string();
                start.push_attribute(("algn", s.as_str()));
            }
        }
        #[cfg(feature = "extra-attrs")]
        for (key, value) in &self.extra_attrs {
            start.push_attribute((key.as_str(), value.as_str()));
        }
        start
    }

    fn write_children<W: Write>(&self, writer: &mut Writer<W>) -> Result<(), SerializeError> {
        #[cfg(feature = "extra-children")]
        let mut extra_iter = self.extra_children.iter().peekable();
        #[cfg(feature = "extra-children")]
        let mut emit_idx: usize = 0;
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        if let Some(ref val) = self.line_fill_properties {
            val.write_element("", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        if let Some(ref val) = self.line_dash_properties {
            val.write_element("", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        if let Some(ref val) = self.line_join_properties {
            val.write_element("", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        #[cfg(feature = "dml-lines")]
        if let Some(ref val) = self.head_end {
            val.write_element("a:headEnd", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        #[cfg(feature = "dml-lines")]
        if let Some(ref val) = self.tail_end {
            val.write_element("a:tailEnd", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        #[cfg(feature = "dml-extensions")]
        if let Some(ref val) = self.ext_lst {
            val.write_element("a:extLst", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        for extra in extra_iter {
            extra.node.write_to(writer).map_err(SerializeError::from)?;
        }
        Ok(())
    }

    fn is_empty_element(&self) -> bool {
        if self.line_fill_properties.is_some() {
            return false;
        }
        if self.line_dash_properties.is_some() {
            return false;
        }
        if self.line_join_properties.is_some() {
            return false;
        }
        #[cfg(feature = "dml-lines")]
        if self.head_end.is_some() {
            return false;
        }
        #[cfg(feature = "dml-lines")]
        if self.tail_end.is_some() {
            return false;
        }
        #[cfg(feature = "dml-extensions")]
        if self.ext_lst.is_some() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for CTShapeProperties {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        #[allow(unused_mut)]
        let mut start = start;
        #[cfg(feature = "dml-shapes")]
        if let Some(ref val) = self.bw_mode {
            {
                let s = val.to_string();
                start.push_attribute(("bwMode", s.as_str()));
            }
        }
        #[cfg(feature = "extra-attrs")]
        for (key, value) in &self.extra_attrs {
            start.push_attribute((key.as_str(), value.as_str()));
        }
        start
    }

    fn write_children<W: Write>(&self, writer: &mut Writer<W>) -> Result<(), SerializeError> {
        #[cfg(feature = "extra-children")]
        let mut extra_iter = self.extra_children.iter().peekable();
        #[cfg(feature = "extra-children")]
        let mut emit_idx: usize = 0;
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        if let Some(ref val) = self.transform {
            val.write_element("a:xfrm", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        if let Some(ref val) = self.geometry {
            val.write_element("", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        if let Some(ref val) = self.fill_properties {
            val.write_element("", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        #[cfg(feature = "dml-lines")]
        if let Some(ref val) = self.line {
            val.write_element("a:ln", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        if let Some(ref val) = self.effect_properties {
            val.write_element("", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        #[cfg(feature = "dml-3d")]
        if let Some(ref val) = self.scene3d {
            val.write_element("a:scene3d", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        #[cfg(feature = "dml-3d")]
        if let Some(ref val) = self.sp3d {
            val.write_element("a:sp3d", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        #[cfg(feature = "dml-extensions")]
        if let Some(ref val) = self.ext_lst {
            val.write_element("a:extLst", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        for extra in extra_iter {
            extra.node.write_to(writer).map_err(SerializeError::from)?;
        }
        Ok(())
    }

    fn is_empty_element(&self) -> bool {
        if self.transform.is_some() {
            return false;
        }
        if self.geometry.is_some() {
            return false;
        }
        if self.fill_properties.is_some() {
            return false;
        }
        #[cfg(feature = "dml-lines")]
        if self.line.is_some() {
            return false;
        }
        if self.effect_properties.is_some() {
            return false;
        }
        #[cfg(feature = "dml-3d")]
        if self.scene3d.is_some() {
            return false;
        }
        #[cfg(feature = "dml-3d")]
        if self.sp3d.is_some() {
            return false;
        }
        #[cfg(feature = "dml-extensions")]
        if self.ext_lst.is_some() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for CTGroupShapeProperties {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        #[allow(unused_mut)]
        let mut start = start;
        #[cfg(feature = "dml-shapes")]
        if let Some(ref val) = self.bw_mode {
            {
                let s = val.to_string();
                start.push_attribute(("bwMode", s.as_str()));
            }
        }
        #[cfg(feature = "extra-attrs")]
        for (key, value) in &self.extra_attrs {
            start.push_attribute((key.as_str(), value.as_str()));
        }
        start
    }

    fn write_children<W: Write>(&self, writer: &mut Writer<W>) -> Result<(), SerializeError> {
        #[cfg(feature = "extra-children")]
        let mut extra_iter = self.extra_children.iter().peekable();
        #[cfg(feature = "extra-children")]
        let mut emit_idx: usize = 0;
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        if let Some(ref val) = self.transform {
            val.write_element("a:xfrm", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        if let Some(ref val) = self.fill_properties {
            val.write_element("", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        if let Some(ref val) = self.effect_properties {
            val.write_element("", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        #[cfg(feature = "dml-3d")]
        if let Some(ref val) = self.scene3d {
            val.write_element("a:scene3d", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        #[cfg(feature = "dml-extensions")]
        if let Some(ref val) = self.ext_lst {
            val.write_element("a:extLst", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        for extra in extra_iter {
            extra.node.write_to(writer).map_err(SerializeError::from)?;
        }
        Ok(())
    }

    fn is_empty_element(&self) -> bool {
        if self.transform.is_some() {
            return false;
        }
        if self.fill_properties.is_some() {
            return false;
        }
        if self.effect_properties.is_some() {
            return false;
        }
        #[cfg(feature = "dml-3d")]
        if self.scene3d.is_some() {
            return false;
        }
        #[cfg(feature = "dml-extensions")]
        if self.ext_lst.is_some() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for CTStyleMatrixReference {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        #[allow(unused_mut)]
        let mut start = start;
        #[cfg(feature = "dml-themes")]
        {
            let val = &self.idx;
            {
                let s = val.to_string();
                start.push_attribute(("idx", s.as_str()));
            }
        }
        #[cfg(feature = "extra-attrs")]
        for (key, value) in &self.extra_attrs {
            start.push_attribute((key.as_str(), value.as_str()));
        }
        start
    }

    fn write_children<W: Write>(&self, writer: &mut Writer<W>) -> Result<(), SerializeError> {
        #[cfg(feature = "extra-children")]
        let mut extra_iter = self.extra_children.iter().peekable();
        #[cfg(feature = "extra-children")]
        let mut emit_idx: usize = 0;
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        if let Some(ref val) = self.color_choice {
            val.write_element("", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        for extra in extra_iter {
            extra.node.write_to(writer).map_err(SerializeError::from)?;
        }
        Ok(())
    }

    fn is_empty_element(&self) -> bool {
        if self.color_choice.is_some() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for CTFontReference {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        #[allow(unused_mut)]
        let mut start = start;
        #[cfg(feature = "dml-themes")]
        {
            let val = &self.idx;
            {
                let s = val.to_string();
                start.push_attribute(("idx", s.as_str()));
            }
        }
        #[cfg(feature = "extra-attrs")]
        for (key, value) in &self.extra_attrs {
            start.push_attribute((key.as_str(), value.as_str()));
        }
        start
    }

    fn write_children<W: Write>(&self, writer: &mut Writer<W>) -> Result<(), SerializeError> {
        #[cfg(feature = "extra-children")]
        let mut extra_iter = self.extra_children.iter().peekable();
        #[cfg(feature = "extra-children")]
        let mut emit_idx: usize = 0;
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        if let Some(ref val) = self.color_choice {
            val.write_element("", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        for extra in extra_iter {
            extra.node.write_to(writer).map_err(SerializeError::from)?;
        }
        Ok(())
    }

    fn is_empty_element(&self) -> bool {
        if self.color_choice.is_some() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for ShapeStyle {
    fn write_children<W: Write>(&self, writer: &mut Writer<W>) -> Result<(), SerializeError> {
        #[cfg(feature = "extra-children")]
        let mut extra_iter = self.extra_children.iter().peekable();
        #[cfg(feature = "extra-children")]
        let mut emit_idx: usize = 0;
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        #[cfg(feature = "dml-shapes")]
        {
            let val = &self.ln_ref;
            val.write_element("a:lnRef", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        #[cfg(feature = "dml-shapes")]
        {
            let val = &self.fill_ref;
            val.write_element("a:fillRef", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        #[cfg(feature = "dml-shapes")]
        {
            let val = &self.effect_ref;
            val.write_element("a:effectRef", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        #[cfg(feature = "dml-shapes")]
        {
            let val = &self.font_ref;
            val.write_element("a:fontRef", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        for extra in extra_iter {
            extra.node.write_to(writer).map_err(SerializeError::from)?;
        }
        Ok(())
    }

    fn is_empty_element(&self) -> bool {
        #[cfg(feature = "dml-shapes")]
        return false;
        #[cfg(feature = "dml-shapes")]
        return false;
        #[cfg(feature = "dml-shapes")]
        return false;
        #[cfg(feature = "dml-shapes")]
        return false;
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for CTDefaultShapeDefinition {
    fn write_children<W: Write>(&self, writer: &mut Writer<W>) -> Result<(), SerializeError> {
        #[cfg(feature = "extra-children")]
        let mut extra_iter = self.extra_children.iter().peekable();
        #[cfg(feature = "extra-children")]
        let mut emit_idx: usize = 0;
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        {
            let val = &self.sp_pr;
            val.write_element("a:spPr", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        {
            let val = &self.body_pr;
            val.write_element("a:bodyPr", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        {
            let val = &self.lst_style;
            val.write_element("a:lstStyle", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        if let Some(ref val) = self.style {
            val.write_element("a:style", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        if let Some(ref val) = self.ext_lst {
            val.write_element("a:extLst", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        for extra in extra_iter {
            extra.node.write_to(writer).map_err(SerializeError::from)?;
        }
        Ok(())
    }

    fn is_empty_element(&self) -> bool {
        false
    }
}

impl ToXml for CTObjectStyleDefaults {
    fn write_children<W: Write>(&self, writer: &mut Writer<W>) -> Result<(), SerializeError> {
        #[cfg(feature = "extra-children")]
        let mut extra_iter = self.extra_children.iter().peekable();
        #[cfg(feature = "extra-children")]
        let mut emit_idx: usize = 0;
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        if let Some(ref val) = self.sp_def {
            val.write_element("a:spDef", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        if let Some(ref val) = self.ln_def {
            val.write_element("a:lnDef", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        if let Some(ref val) = self.tx_def {
            val.write_element("a:txDef", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        if let Some(ref val) = self.ext_lst {
            val.write_element("a:extLst", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        for extra in extra_iter {
            extra.node.write_to(writer).map_err(SerializeError::from)?;
        }
        Ok(())
    }

    fn is_empty_element(&self) -> bool {
        if self.sp_def.is_some() {
            return false;
        }
        if self.ln_def.is_some() {
            return false;
        }
        if self.tx_def.is_some() {
            return false;
        }
        if self.ext_lst.is_some() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for CTEmptyElement {
    fn is_empty_element(&self) -> bool {
        true
    }
}

impl ToXml for CTColorMapping {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        #[allow(unused_mut)]
        let mut start = start;
        {
            let val = &self.bg1;
            {
                let s = val.to_string();
                start.push_attribute(("bg1", s.as_str()));
            }
        }
        {
            let val = &self.tx1;
            {
                let s = val.to_string();
                start.push_attribute(("tx1", s.as_str()));
            }
        }
        {
            let val = &self.bg2;
            {
                let s = val.to_string();
                start.push_attribute(("bg2", s.as_str()));
            }
        }
        {
            let val = &self.tx2;
            {
                let s = val.to_string();
                start.push_attribute(("tx2", s.as_str()));
            }
        }
        {
            let val = &self.accent1;
            {
                let s = val.to_string();
                start.push_attribute(("accent1", s.as_str()));
            }
        }
        {
            let val = &self.accent2;
            {
                let s = val.to_string();
                start.push_attribute(("accent2", s.as_str()));
            }
        }
        {
            let val = &self.accent3;
            {
                let s = val.to_string();
                start.push_attribute(("accent3", s.as_str()));
            }
        }
        {
            let val = &self.accent4;
            {
                let s = val.to_string();
                start.push_attribute(("accent4", s.as_str()));
            }
        }
        {
            let val = &self.accent5;
            {
                let s = val.to_string();
                start.push_attribute(("accent5", s.as_str()));
            }
        }
        {
            let val = &self.accent6;
            {
                let s = val.to_string();
                start.push_attribute(("accent6", s.as_str()));
            }
        }
        {
            let val = &self.hlink;
            {
                let s = val.to_string();
                start.push_attribute(("hlink", s.as_str()));
            }
        }
        {
            let val = &self.fol_hlink;
            {
                let s = val.to_string();
                start.push_attribute(("folHlink", s.as_str()));
            }
        }
        #[cfg(feature = "extra-attrs")]
        for (key, value) in &self.extra_attrs {
            start.push_attribute((key.as_str(), value.as_str()));
        }
        start
    }

    fn write_children<W: Write>(&self, writer: &mut Writer<W>) -> Result<(), SerializeError> {
        #[cfg(feature = "extra-children")]
        let mut extra_iter = self.extra_children.iter().peekable();
        #[cfg(feature = "extra-children")]
        let mut emit_idx: usize = 0;
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        if let Some(ref val) = self.ext_lst {
            val.write_element("a:extLst", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        for extra in extra_iter {
            extra.node.write_to(writer).map_err(SerializeError::from)?;
        }
        Ok(())
    }

    fn is_empty_element(&self) -> bool {
        if self.ext_lst.is_some() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for CTColorMappingOverride {
    fn write_children<W: Write>(&self, writer: &mut Writer<W>) -> Result<(), SerializeError> {
        #[cfg(feature = "extra-children")]
        let mut extra_iter = self.extra_children.iter().peekable();
        #[cfg(feature = "extra-children")]
        let mut emit_idx: usize = 0;
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        #[cfg(feature = "dml-colors")]
        if let Some(ref val) = self.master_clr_mapping {
            val.write_element("a:masterClrMapping", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        #[cfg(feature = "dml-colors")]
        if let Some(ref val) = self.override_clr_mapping {
            val.write_element("a:overrideClrMapping", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        for extra in extra_iter {
            extra.node.write_to(writer).map_err(SerializeError::from)?;
        }
        Ok(())
    }

    fn is_empty_element(&self) -> bool {
        #[cfg(feature = "dml-colors")]
        if self.master_clr_mapping.is_some() {
            return false;
        }
        #[cfg(feature = "dml-colors")]
        if self.override_clr_mapping.is_some() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for CTColorSchemeAndMapping {
    fn write_children<W: Write>(&self, writer: &mut Writer<W>) -> Result<(), SerializeError> {
        #[cfg(feature = "extra-children")]
        let mut extra_iter = self.extra_children.iter().peekable();
        #[cfg(feature = "extra-children")]
        let mut emit_idx: usize = 0;
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        {
            let val = &self.clr_scheme;
            val.write_element("a:clrScheme", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        if let Some(ref val) = self.clr_map {
            val.write_element("a:clrMap", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        for extra in extra_iter {
            extra.node.write_to(writer).map_err(SerializeError::from)?;
        }
        Ok(())
    }

    fn is_empty_element(&self) -> bool {
        false
    }
}

impl ToXml for CTColorSchemeList {
    fn write_children<W: Write>(&self, writer: &mut Writer<W>) -> Result<(), SerializeError> {
        #[cfg(feature = "extra-children")]
        let mut extra_iter = self.extra_children.iter().peekable();
        #[cfg(feature = "extra-children")]
        let mut emit_idx: usize = 0;
        for item in &self.extra_clr_scheme {
            #[cfg(feature = "extra-children")]
            while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
                extra_iter
                    .next()
                    .unwrap()
                    .node
                    .write_to(writer)
                    .map_err(SerializeError::from)?;
            }
            item.write_element("a:extraClrScheme", writer)?;
            #[cfg(feature = "extra-children")]
            {
                emit_idx += 1;
            }
        }
        #[cfg(feature = "extra-children")]
        for extra in extra_iter {
            extra.node.write_to(writer).map_err(SerializeError::from)?;
        }
        Ok(())
    }

    fn is_empty_element(&self) -> bool {
        if !self.extra_clr_scheme.is_empty() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for CTOfficeStyleSheet {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        #[allow(unused_mut)]
        let mut start = start;
        #[cfg(feature = "dml-themes")]
        if let Some(ref val) = self.name {
            start.push_attribute(("name", val.as_str()));
        }
        #[cfg(feature = "extra-attrs")]
        for (key, value) in &self.extra_attrs {
            start.push_attribute((key.as_str(), value.as_str()));
        }
        start
    }

    fn write_children<W: Write>(&self, writer: &mut Writer<W>) -> Result<(), SerializeError> {
        #[cfg(feature = "extra-children")]
        let mut extra_iter = self.extra_children.iter().peekable();
        #[cfg(feature = "extra-children")]
        let mut emit_idx: usize = 0;
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        #[cfg(feature = "dml-themes")]
        {
            let val = &self.theme_elements;
            val.write_element("a:themeElements", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        #[cfg(feature = "dml-themes")]
        if let Some(ref val) = self.object_defaults {
            val.write_element("a:objectDefaults", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        #[cfg(feature = "dml-themes")]
        if let Some(ref val) = self.extra_clr_scheme_lst {
            val.write_element("a:extraClrSchemeLst", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        #[cfg(feature = "dml-themes")]
        if let Some(ref val) = self.cust_clr_lst {
            val.write_element("a:custClrLst", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        #[cfg(feature = "dml-extensions")]
        if let Some(ref val) = self.ext_lst {
            val.write_element("a:extLst", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        for extra in extra_iter {
            extra.node.write_to(writer).map_err(SerializeError::from)?;
        }
        Ok(())
    }

    fn is_empty_element(&self) -> bool {
        #[cfg(feature = "dml-themes")]
        return false;
        #[cfg(feature = "dml-themes")]
        if self.object_defaults.is_some() {
            return false;
        }
        #[cfg(feature = "dml-themes")]
        if self.extra_clr_scheme_lst.is_some() {
            return false;
        }
        #[cfg(feature = "dml-themes")]
        if self.cust_clr_lst.is_some() {
            return false;
        }
        #[cfg(feature = "dml-extensions")]
        if self.ext_lst.is_some() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for CTBaseStylesOverride {
    fn write_children<W: Write>(&self, writer: &mut Writer<W>) -> Result<(), SerializeError> {
        #[cfg(feature = "extra-children")]
        let mut extra_iter = self.extra_children.iter().peekable();
        #[cfg(feature = "extra-children")]
        let mut emit_idx: usize = 0;
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        if let Some(ref val) = self.clr_scheme {
            val.write_element("a:clrScheme", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        if let Some(ref val) = self.font_scheme {
            val.write_element("a:fontScheme", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        if let Some(ref val) = self.fmt_scheme {
            val.write_element("a:fmtScheme", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        for extra in extra_iter {
            extra.node.write_to(writer).map_err(SerializeError::from)?;
        }
        Ok(())
    }

    fn is_empty_element(&self) -> bool {
        if self.clr_scheme.is_some() {
            return false;
        }
        if self.font_scheme.is_some() {
            return false;
        }
        if self.fmt_scheme.is_some() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for CTClipboardStyleSheet {
    fn write_children<W: Write>(&self, writer: &mut Writer<W>) -> Result<(), SerializeError> {
        #[cfg(feature = "extra-children")]
        let mut extra_iter = self.extra_children.iter().peekable();
        #[cfg(feature = "extra-children")]
        let mut emit_idx: usize = 0;
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        {
            let val = &self.theme_elements;
            val.write_element("a:themeElements", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        {
            let val = &self.clr_map;
            val.write_element("a:clrMap", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        for extra in extra_iter {
            extra.node.write_to(writer).map_err(SerializeError::from)?;
        }
        Ok(())
    }

    fn is_empty_element(&self) -> bool {
        false
    }
}

impl ToXml for CTTableCellProperties {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        #[allow(unused_mut)]
        let mut start = start;
        #[cfg(feature = "dml-tables")]
        if let Some(ref val) = self.mar_l {
            {
                let s = val.to_string();
                start.push_attribute(("marL", s.as_str()));
            }
        }
        #[cfg(feature = "dml-tables")]
        if let Some(ref val) = self.mar_r {
            {
                let s = val.to_string();
                start.push_attribute(("marR", s.as_str()));
            }
        }
        #[cfg(feature = "dml-tables")]
        if let Some(ref val) = self.mar_t {
            {
                let s = val.to_string();
                start.push_attribute(("marT", s.as_str()));
            }
        }
        #[cfg(feature = "dml-tables")]
        if let Some(ref val) = self.mar_b {
            {
                let s = val.to_string();
                start.push_attribute(("marB", s.as_str()));
            }
        }
        #[cfg(feature = "dml-tables")]
        if let Some(ref val) = self.vert {
            {
                let s = val.to_string();
                start.push_attribute(("vert", s.as_str()));
            }
        }
        #[cfg(feature = "dml-tables")]
        if let Some(ref val) = self.anchor {
            {
                let s = val.to_string();
                start.push_attribute(("anchor", s.as_str()));
            }
        }
        #[cfg(feature = "dml-tables")]
        if let Some(ref val) = self.anchor_ctr {
            start.push_attribute(("anchorCtr", if *val { "1" } else { "0" }));
        }
        #[cfg(feature = "dml-tables")]
        if let Some(ref val) = self.horz_overflow {
            {
                let s = val.to_string();
                start.push_attribute(("horzOverflow", s.as_str()));
            }
        }
        #[cfg(feature = "extra-attrs")]
        for (key, value) in &self.extra_attrs {
            start.push_attribute((key.as_str(), value.as_str()));
        }
        start
    }

    fn write_children<W: Write>(&self, writer: &mut Writer<W>) -> Result<(), SerializeError> {
        #[cfg(feature = "extra-children")]
        let mut extra_iter = self.extra_children.iter().peekable();
        #[cfg(feature = "extra-children")]
        let mut emit_idx: usize = 0;
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        #[cfg(feature = "dml-tables")]
        if let Some(ref val) = self.ln_l {
            val.write_element("a:lnL", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        #[cfg(feature = "dml-tables")]
        if let Some(ref val) = self.ln_r {
            val.write_element("a:lnR", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        #[cfg(feature = "dml-tables")]
        if let Some(ref val) = self.ln_t {
            val.write_element("a:lnT", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        #[cfg(feature = "dml-tables")]
        if let Some(ref val) = self.ln_b {
            val.write_element("a:lnB", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        #[cfg(feature = "dml-tables")]
        if let Some(ref val) = self.ln_tl_to_br {
            val.write_element("a:lnTlToBr", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        #[cfg(feature = "dml-tables")]
        if let Some(ref val) = self.ln_bl_to_tr {
            val.write_element("a:lnBlToTr", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        #[cfg(feature = "dml-tables")]
        if let Some(ref val) = self.cell3_d {
            val.write_element("a:cell3D", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        if let Some(ref val) = self.fill_properties {
            val.write_element("", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        if let Some(ref val) = self.headers {
            val.write_element("a:headers", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        #[cfg(feature = "dml-extensions")]
        if let Some(ref val) = self.ext_lst {
            val.write_element("a:extLst", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        for extra in extra_iter {
            extra.node.write_to(writer).map_err(SerializeError::from)?;
        }
        Ok(())
    }

    fn is_empty_element(&self) -> bool {
        #[cfg(feature = "dml-tables")]
        if self.ln_l.is_some() {
            return false;
        }
        #[cfg(feature = "dml-tables")]
        if self.ln_r.is_some() {
            return false;
        }
        #[cfg(feature = "dml-tables")]
        if self.ln_t.is_some() {
            return false;
        }
        #[cfg(feature = "dml-tables")]
        if self.ln_b.is_some() {
            return false;
        }
        #[cfg(feature = "dml-tables")]
        if self.ln_tl_to_br.is_some() {
            return false;
        }
        #[cfg(feature = "dml-tables")]
        if self.ln_bl_to_tr.is_some() {
            return false;
        }
        #[cfg(feature = "dml-tables")]
        if self.cell3_d.is_some() {
            return false;
        }
        if self.fill_properties.is_some() {
            return false;
        }
        if self.headers.is_some() {
            return false;
        }
        #[cfg(feature = "dml-extensions")]
        if self.ext_lst.is_some() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for CTHeaders {
    fn write_children<W: Write>(&self, writer: &mut Writer<W>) -> Result<(), SerializeError> {
        #[cfg(feature = "extra-children")]
        let mut extra_iter = self.extra_children.iter().peekable();
        #[cfg(feature = "extra-children")]
        let mut emit_idx: usize = 0;
        for item in &self.header {
            #[cfg(feature = "extra-children")]
            while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
                extra_iter
                    .next()
                    .unwrap()
                    .node
                    .write_to(writer)
                    .map_err(SerializeError::from)?;
            }
            {
                let start = BytesStart::new("a:header");
                writer.write_event(Event::Start(start))?;
                writer.write_event(Event::Text(BytesText::new(item.as_str())))?;
                writer.write_event(Event::End(BytesEnd::new("a:header")))?;
            }
            #[cfg(feature = "extra-children")]
            {
                emit_idx += 1;
            }
        }
        #[cfg(feature = "extra-children")]
        for extra in extra_iter {
            extra.node.write_to(writer).map_err(SerializeError::from)?;
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

impl ToXml for CTTableCol {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        #[allow(unused_mut)]
        let mut start = start;
        {
            let val = &self.width;
            {
                let s = val.to_string();
                start.push_attribute(("w", s.as_str()));
            }
        }
        #[cfg(feature = "extra-attrs")]
        for (key, value) in &self.extra_attrs {
            start.push_attribute((key.as_str(), value.as_str()));
        }
        start
    }

    fn write_children<W: Write>(&self, writer: &mut Writer<W>) -> Result<(), SerializeError> {
        #[cfg(feature = "extra-children")]
        let mut extra_iter = self.extra_children.iter().peekable();
        #[cfg(feature = "extra-children")]
        let mut emit_idx: usize = 0;
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        if let Some(ref val) = self.ext_lst {
            val.write_element("a:extLst", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        for extra in extra_iter {
            extra.node.write_to(writer).map_err(SerializeError::from)?;
        }
        Ok(())
    }

    fn is_empty_element(&self) -> bool {
        if self.ext_lst.is_some() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for CTTableGrid {
    fn write_children<W: Write>(&self, writer: &mut Writer<W>) -> Result<(), SerializeError> {
        #[cfg(feature = "extra-children")]
        let mut extra_iter = self.extra_children.iter().peekable();
        #[cfg(feature = "extra-children")]
        let mut emit_idx: usize = 0;
        #[cfg(feature = "dml-tables")]
        for item in &self.grid_col {
            #[cfg(feature = "extra-children")]
            while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
                extra_iter
                    .next()
                    .unwrap()
                    .node
                    .write_to(writer)
                    .map_err(SerializeError::from)?;
            }
            item.write_element("a:gridCol", writer)?;
            #[cfg(feature = "extra-children")]
            {
                emit_idx += 1;
            }
        }
        #[cfg(feature = "extra-children")]
        for extra in extra_iter {
            extra.node.write_to(writer).map_err(SerializeError::from)?;
        }
        Ok(())
    }

    fn is_empty_element(&self) -> bool {
        #[cfg(feature = "dml-tables")]
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

impl ToXml for CTTableCell {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        #[allow(unused_mut)]
        let mut start = start;
        #[cfg(feature = "dml-tables")]
        if let Some(ref val) = self.row_span {
            {
                let s = val.to_string();
                start.push_attribute(("rowSpan", s.as_str()));
            }
        }
        #[cfg(feature = "dml-tables")]
        if let Some(ref val) = self.grid_span {
            {
                let s = val.to_string();
                start.push_attribute(("gridSpan", s.as_str()));
            }
        }
        #[cfg(feature = "dml-tables")]
        if let Some(ref val) = self.h_merge {
            start.push_attribute(("hMerge", if *val { "1" } else { "0" }));
        }
        #[cfg(feature = "dml-tables")]
        if let Some(ref val) = self.v_merge {
            start.push_attribute(("vMerge", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.id {
            start.push_attribute(("id", val.as_str()));
        }
        #[cfg(feature = "extra-attrs")]
        for (key, value) in &self.extra_attrs {
            start.push_attribute((key.as_str(), value.as_str()));
        }
        start
    }

    fn write_children<W: Write>(&self, writer: &mut Writer<W>) -> Result<(), SerializeError> {
        #[cfg(feature = "extra-children")]
        let mut extra_iter = self.extra_children.iter().peekable();
        #[cfg(feature = "extra-children")]
        let mut emit_idx: usize = 0;
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        #[cfg(feature = "dml-tables")]
        if let Some(ref val) = self.tx_body {
            val.write_element("a:txBody", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        #[cfg(feature = "dml-tables")]
        if let Some(ref val) = self.tc_pr {
            val.write_element("a:tcPr", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        #[cfg(feature = "dml-extensions")]
        if let Some(ref val) = self.ext_lst {
            val.write_element("a:extLst", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        for extra in extra_iter {
            extra.node.write_to(writer).map_err(SerializeError::from)?;
        }
        Ok(())
    }

    fn is_empty_element(&self) -> bool {
        #[cfg(feature = "dml-tables")]
        if self.tx_body.is_some() {
            return false;
        }
        #[cfg(feature = "dml-tables")]
        if self.tc_pr.is_some() {
            return false;
        }
        #[cfg(feature = "dml-extensions")]
        if self.ext_lst.is_some() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for CTTableRow {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        #[allow(unused_mut)]
        let mut start = start;
        #[cfg(feature = "dml-tables")]
        {
            let val = &self.height;
            {
                let s = val.to_string();
                start.push_attribute(("h", s.as_str()));
            }
        }
        #[cfg(feature = "extra-attrs")]
        for (key, value) in &self.extra_attrs {
            start.push_attribute((key.as_str(), value.as_str()));
        }
        start
    }

    fn write_children<W: Write>(&self, writer: &mut Writer<W>) -> Result<(), SerializeError> {
        #[cfg(feature = "extra-children")]
        let mut extra_iter = self.extra_children.iter().peekable();
        #[cfg(feature = "extra-children")]
        let mut emit_idx: usize = 0;
        #[cfg(feature = "dml-tables")]
        for item in &self.tc {
            #[cfg(feature = "extra-children")]
            while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
                extra_iter
                    .next()
                    .unwrap()
                    .node
                    .write_to(writer)
                    .map_err(SerializeError::from)?;
            }
            item.write_element("a:tc", writer)?;
            #[cfg(feature = "extra-children")]
            {
                emit_idx += 1;
            }
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        #[cfg(feature = "dml-extensions")]
        if let Some(ref val) = self.ext_lst {
            val.write_element("a:extLst", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        for extra in extra_iter {
            extra.node.write_to(writer).map_err(SerializeError::from)?;
        }
        Ok(())
    }

    fn is_empty_element(&self) -> bool {
        #[cfg(feature = "dml-tables")]
        if !self.tc.is_empty() {
            return false;
        }
        #[cfg(feature = "dml-extensions")]
        if self.ext_lst.is_some() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for CTTableProperties {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        #[allow(unused_mut)]
        let mut start = start;
        #[cfg(feature = "dml-tables")]
        if let Some(ref val) = self.rtl {
            start.push_attribute(("rtl", if *val { "1" } else { "0" }));
        }
        #[cfg(feature = "dml-tables")]
        if let Some(ref val) = self.first_row {
            start.push_attribute(("firstRow", if *val { "1" } else { "0" }));
        }
        #[cfg(feature = "dml-tables")]
        if let Some(ref val) = self.first_col {
            start.push_attribute(("firstCol", if *val { "1" } else { "0" }));
        }
        #[cfg(feature = "dml-tables")]
        if let Some(ref val) = self.last_row {
            start.push_attribute(("lastRow", if *val { "1" } else { "0" }));
        }
        #[cfg(feature = "dml-tables")]
        if let Some(ref val) = self.last_col {
            start.push_attribute(("lastCol", if *val { "1" } else { "0" }));
        }
        #[cfg(feature = "dml-tables")]
        if let Some(ref val) = self.band_row {
            start.push_attribute(("bandRow", if *val { "1" } else { "0" }));
        }
        #[cfg(feature = "dml-tables")]
        if let Some(ref val) = self.band_col {
            start.push_attribute(("bandCol", if *val { "1" } else { "0" }));
        }
        #[cfg(feature = "extra-attrs")]
        for (key, value) in &self.extra_attrs {
            start.push_attribute((key.as_str(), value.as_str()));
        }
        start
    }

    fn write_children<W: Write>(&self, writer: &mut Writer<W>) -> Result<(), SerializeError> {
        #[cfg(feature = "extra-children")]
        let mut extra_iter = self.extra_children.iter().peekable();
        #[cfg(feature = "extra-children")]
        let mut emit_idx: usize = 0;
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        if let Some(ref val) = self.fill_properties {
            val.write_element("", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        if let Some(ref val) = self.effect_properties {
            val.write_element("", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        if let Some(ref val) = self.table_style {
            val.write_element("a:tableStyle", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        #[cfg(feature = "dml-tables")]
        if let Some(ref val) = self.table_style_id {
            {
                let start = BytesStart::new("a:tableStyleId");
                writer.write_event(Event::Start(start))?;
                writer.write_event(Event::Text(BytesText::new(val.as_str())))?;
                writer.write_event(Event::End(BytesEnd::new("a:tableStyleId")))?;
            }
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        #[cfg(feature = "dml-extensions")]
        if let Some(ref val) = self.ext_lst {
            val.write_element("a:extLst", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        for extra in extra_iter {
            extra.node.write_to(writer).map_err(SerializeError::from)?;
        }
        Ok(())
    }

    fn is_empty_element(&self) -> bool {
        if self.fill_properties.is_some() {
            return false;
        }
        if self.effect_properties.is_some() {
            return false;
        }
        if self.table_style.is_some() {
            return false;
        }
        #[cfg(feature = "dml-tables")]
        if self.table_style_id.is_some() {
            return false;
        }
        #[cfg(feature = "dml-extensions")]
        if self.ext_lst.is_some() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for CTTable {
    fn write_children<W: Write>(&self, writer: &mut Writer<W>) -> Result<(), SerializeError> {
        #[cfg(feature = "extra-children")]
        let mut extra_iter = self.extra_children.iter().peekable();
        #[cfg(feature = "extra-children")]
        let mut emit_idx: usize = 0;
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        #[cfg(feature = "dml-tables")]
        if let Some(ref val) = self.tbl_pr {
            val.write_element("a:tblPr", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        #[cfg(feature = "dml-tables")]
        {
            let val = &self.tbl_grid;
            val.write_element("a:tblGrid", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "dml-tables")]
        for item in &self.tr {
            #[cfg(feature = "extra-children")]
            while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
                extra_iter
                    .next()
                    .unwrap()
                    .node
                    .write_to(writer)
                    .map_err(SerializeError::from)?;
            }
            item.write_element("a:tr", writer)?;
            #[cfg(feature = "extra-children")]
            {
                emit_idx += 1;
            }
        }
        #[cfg(feature = "extra-children")]
        for extra in extra_iter {
            extra.node.write_to(writer).map_err(SerializeError::from)?;
        }
        Ok(())
    }

    fn is_empty_element(&self) -> bool {
        #[cfg(feature = "dml-tables")]
        if self.tbl_pr.is_some() {
            return false;
        }
        #[cfg(feature = "dml-tables")]
        return false;
        #[cfg(feature = "dml-tables")]
        if !self.tr.is_empty() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for CTCell3D {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        #[allow(unused_mut)]
        let mut start = start;
        if let Some(ref val) = self.prst_material {
            {
                let s = val.to_string();
                start.push_attribute(("prstMaterial", s.as_str()));
            }
        }
        #[cfg(feature = "extra-attrs")]
        for (key, value) in &self.extra_attrs {
            start.push_attribute((key.as_str(), value.as_str()));
        }
        start
    }

    fn write_children<W: Write>(&self, writer: &mut Writer<W>) -> Result<(), SerializeError> {
        #[cfg(feature = "extra-children")]
        let mut extra_iter = self.extra_children.iter().peekable();
        #[cfg(feature = "extra-children")]
        let mut emit_idx: usize = 0;
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        {
            let val = &self.bevel;
            val.write_element("a:bevel", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        if let Some(ref val) = self.light_rig {
            val.write_element("a:lightRig", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        if let Some(ref val) = self.ext_lst {
            val.write_element("a:extLst", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        for extra in extra_iter {
            extra.node.write_to(writer).map_err(SerializeError::from)?;
        }
        Ok(())
    }

    fn is_empty_element(&self) -> bool {
        false
    }
}

impl ToXml for EGThemeableFillStyle {
    fn write_element<W: Write>(
        &self,
        _tag: &str,
        writer: &mut Writer<W>,
    ) -> Result<(), SerializeError> {
        match self {
            Self::Fill(inner) => inner.write_element("a:fill", writer)?,
            Self::FillRef(inner) => inner.write_element("a:fillRef", writer)?,
        }
        Ok(())
    }
}

impl ToXml for CTThemeableLineStyle {
    fn write_children<W: Write>(&self, writer: &mut Writer<W>) -> Result<(), SerializeError> {
        #[cfg(feature = "extra-children")]
        let mut extra_iter = self.extra_children.iter().peekable();
        #[cfg(feature = "extra-children")]
        let mut emit_idx: usize = 0;
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        if let Some(ref val) = self.line {
            val.write_element("a:ln", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        if let Some(ref val) = self.ln_ref {
            val.write_element("a:lnRef", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        for extra in extra_iter {
            extra.node.write_to(writer).map_err(SerializeError::from)?;
        }
        Ok(())
    }

    fn is_empty_element(&self) -> bool {
        if self.line.is_some() {
            return false;
        }
        if self.ln_ref.is_some() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for EGThemeableEffectStyle {
    fn write_element<W: Write>(
        &self,
        _tag: &str,
        writer: &mut Writer<W>,
    ) -> Result<(), SerializeError> {
        match self {
            Self::Effect(inner) => inner.write_element("a:effect", writer)?,
            Self::EffectRef(inner) => inner.write_element("a:effectRef", writer)?,
        }
        Ok(())
    }
}

impl ToXml for EGThemeableFontStyles {
    fn write_element<W: Write>(
        &self,
        _tag: &str,
        writer: &mut Writer<W>,
    ) -> Result<(), SerializeError> {
        match self {
            Self::Font(inner) => inner.write_element("a:font", writer)?,
            Self::FontRef(inner) => inner.write_element("a:fontRef", writer)?,
        }
        Ok(())
    }
}

impl ToXml for CTTableStyleTextStyle {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        #[allow(unused_mut)]
        let mut start = start;
        if let Some(ref val) = self.b {
            {
                let s = val.to_string();
                start.push_attribute(("b", s.as_str()));
            }
        }
        if let Some(ref val) = self.i {
            {
                let s = val.to_string();
                start.push_attribute(("i", s.as_str()));
            }
        }
        #[cfg(feature = "extra-attrs")]
        for (key, value) in &self.extra_attrs {
            start.push_attribute((key.as_str(), value.as_str()));
        }
        start
    }

    fn write_children<W: Write>(&self, writer: &mut Writer<W>) -> Result<(), SerializeError> {
        #[cfg(feature = "extra-children")]
        let mut extra_iter = self.extra_children.iter().peekable();
        #[cfg(feature = "extra-children")]
        let mut emit_idx: usize = 0;
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        if let Some(ref val) = self.themeable_font_styles {
            val.write_element("", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        if let Some(ref val) = self.color_choice {
            val.write_element("", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        if let Some(ref val) = self.ext_lst {
            val.write_element("a:extLst", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        for extra in extra_iter {
            extra.node.write_to(writer).map_err(SerializeError::from)?;
        }
        Ok(())
    }

    fn is_empty_element(&self) -> bool {
        if self.themeable_font_styles.is_some() {
            return false;
        }
        if self.color_choice.is_some() {
            return false;
        }
        if self.ext_lst.is_some() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for CTTableCellBorderStyle {
    fn write_children<W: Write>(&self, writer: &mut Writer<W>) -> Result<(), SerializeError> {
        #[cfg(feature = "extra-children")]
        let mut extra_iter = self.extra_children.iter().peekable();
        #[cfg(feature = "extra-children")]
        let mut emit_idx: usize = 0;
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        if let Some(ref val) = self.left {
            val.write_element("a:left", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        if let Some(ref val) = self.right {
            val.write_element("a:right", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        if let Some(ref val) = self.top {
            val.write_element("a:top", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        if let Some(ref val) = self.bottom {
            val.write_element("a:bottom", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        if let Some(ref val) = self.inside_h {
            val.write_element("a:insideH", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        if let Some(ref val) = self.inside_v {
            val.write_element("a:insideV", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        if let Some(ref val) = self.tl2br {
            val.write_element("a:tl2br", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        if let Some(ref val) = self.tr2bl {
            val.write_element("a:tr2bl", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        if let Some(ref val) = self.ext_lst {
            val.write_element("a:extLst", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        for extra in extra_iter {
            extra.node.write_to(writer).map_err(SerializeError::from)?;
        }
        Ok(())
    }

    fn is_empty_element(&self) -> bool {
        if self.left.is_some() {
            return false;
        }
        if self.right.is_some() {
            return false;
        }
        if self.top.is_some() {
            return false;
        }
        if self.bottom.is_some() {
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
        if self.ext_lst.is_some() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for CTTableBackgroundStyle {
    fn write_children<W: Write>(&self, writer: &mut Writer<W>) -> Result<(), SerializeError> {
        #[cfg(feature = "extra-children")]
        let mut extra_iter = self.extra_children.iter().peekable();
        #[cfg(feature = "extra-children")]
        let mut emit_idx: usize = 0;
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        if let Some(ref val) = self.themeable_fill_style {
            val.write_element("", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        if let Some(ref val) = self.themeable_effect_style {
            val.write_element("", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        for extra in extra_iter {
            extra.node.write_to(writer).map_err(SerializeError::from)?;
        }
        Ok(())
    }

    fn is_empty_element(&self) -> bool {
        if self.themeable_fill_style.is_some() {
            return false;
        }
        if self.themeable_effect_style.is_some() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for CTTableStyleCellStyle {
    fn write_children<W: Write>(&self, writer: &mut Writer<W>) -> Result<(), SerializeError> {
        #[cfg(feature = "extra-children")]
        let mut extra_iter = self.extra_children.iter().peekable();
        #[cfg(feature = "extra-children")]
        let mut emit_idx: usize = 0;
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        if let Some(ref val) = self.tc_bdr {
            val.write_element("a:tcBdr", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        if let Some(ref val) = self.themeable_fill_style {
            val.write_element("", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        if let Some(ref val) = self.cell3_d {
            val.write_element("a:cell3D", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        for extra in extra_iter {
            extra.node.write_to(writer).map_err(SerializeError::from)?;
        }
        Ok(())
    }

    fn is_empty_element(&self) -> bool {
        if self.tc_bdr.is_some() {
            return false;
        }
        if self.themeable_fill_style.is_some() {
            return false;
        }
        if self.cell3_d.is_some() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for CTTablePartStyle {
    fn write_children<W: Write>(&self, writer: &mut Writer<W>) -> Result<(), SerializeError> {
        #[cfg(feature = "extra-children")]
        let mut extra_iter = self.extra_children.iter().peekable();
        #[cfg(feature = "extra-children")]
        let mut emit_idx: usize = 0;
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        if let Some(ref val) = self.tc_tx_style {
            val.write_element("a:tcTxStyle", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        if let Some(ref val) = self.tc_style {
            val.write_element("a:tcStyle", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        for extra in extra_iter {
            extra.node.write_to(writer).map_err(SerializeError::from)?;
        }
        Ok(())
    }

    fn is_empty_element(&self) -> bool {
        if self.tc_tx_style.is_some() {
            return false;
        }
        if self.tc_style.is_some() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for CTTableStyle {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        #[allow(unused_mut)]
        let mut start = start;
        {
            let val = &self.style_id;
            start.push_attribute(("styleId", val.as_str()));
        }
        {
            let val = &self.style_name;
            start.push_attribute(("styleName", val.as_str()));
        }
        #[cfg(feature = "extra-attrs")]
        for (key, value) in &self.extra_attrs {
            start.push_attribute((key.as_str(), value.as_str()));
        }
        start
    }

    fn write_children<W: Write>(&self, writer: &mut Writer<W>) -> Result<(), SerializeError> {
        #[cfg(feature = "extra-children")]
        let mut extra_iter = self.extra_children.iter().peekable();
        #[cfg(feature = "extra-children")]
        let mut emit_idx: usize = 0;
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        if let Some(ref val) = self.tbl_bg {
            val.write_element("a:tblBg", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        if let Some(ref val) = self.whole_tbl {
            val.write_element("a:wholeTbl", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        if let Some(ref val) = self.band1_h {
            val.write_element("a:band1H", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        if let Some(ref val) = self.band2_h {
            val.write_element("a:band2H", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        if let Some(ref val) = self.band1_v {
            val.write_element("a:band1V", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        if let Some(ref val) = self.band2_v {
            val.write_element("a:band2V", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        if let Some(ref val) = self.last_col {
            val.write_element("a:lastCol", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        if let Some(ref val) = self.first_col {
            val.write_element("a:firstCol", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        if let Some(ref val) = self.last_row {
            val.write_element("a:lastRow", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        if let Some(ref val) = self.se_cell {
            val.write_element("a:seCell", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        if let Some(ref val) = self.sw_cell {
            val.write_element("a:swCell", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        if let Some(ref val) = self.first_row {
            val.write_element("a:firstRow", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        if let Some(ref val) = self.ne_cell {
            val.write_element("a:neCell", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        if let Some(ref val) = self.nw_cell {
            val.write_element("a:nwCell", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        if let Some(ref val) = self.ext_lst {
            val.write_element("a:extLst", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        for extra in extra_iter {
            extra.node.write_to(writer).map_err(SerializeError::from)?;
        }
        Ok(())
    }

    fn is_empty_element(&self) -> bool {
        if self.tbl_bg.is_some() {
            return false;
        }
        if self.whole_tbl.is_some() {
            return false;
        }
        if self.band1_h.is_some() {
            return false;
        }
        if self.band2_h.is_some() {
            return false;
        }
        if self.band1_v.is_some() {
            return false;
        }
        if self.band2_v.is_some() {
            return false;
        }
        if self.last_col.is_some() {
            return false;
        }
        if self.first_col.is_some() {
            return false;
        }
        if self.last_row.is_some() {
            return false;
        }
        if self.se_cell.is_some() {
            return false;
        }
        if self.sw_cell.is_some() {
            return false;
        }
        if self.first_row.is_some() {
            return false;
        }
        if self.ne_cell.is_some() {
            return false;
        }
        if self.nw_cell.is_some() {
            return false;
        }
        if self.ext_lst.is_some() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for CTTableStyleList {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        #[allow(unused_mut)]
        let mut start = start;
        {
            let val = &self.def;
            start.push_attribute(("def", val.as_str()));
        }
        #[cfg(feature = "extra-attrs")]
        for (key, value) in &self.extra_attrs {
            start.push_attribute((key.as_str(), value.as_str()));
        }
        start
    }

    fn write_children<W: Write>(&self, writer: &mut Writer<W>) -> Result<(), SerializeError> {
        #[cfg(feature = "extra-children")]
        let mut extra_iter = self.extra_children.iter().peekable();
        #[cfg(feature = "extra-children")]
        let mut emit_idx: usize = 0;
        for item in &self.tbl_style {
            #[cfg(feature = "extra-children")]
            while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
                extra_iter
                    .next()
                    .unwrap()
                    .node
                    .write_to(writer)
                    .map_err(SerializeError::from)?;
            }
            item.write_element("a:tblStyle", writer)?;
            #[cfg(feature = "extra-children")]
            {
                emit_idx += 1;
            }
        }
        #[cfg(feature = "extra-children")]
        for extra in extra_iter {
            extra.node.write_to(writer).map_err(SerializeError::from)?;
        }
        Ok(())
    }

    fn is_empty_element(&self) -> bool {
        if !self.tbl_style.is_empty() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for TextParagraph {
    fn write_children<W: Write>(&self, writer: &mut Writer<W>) -> Result<(), SerializeError> {
        #[cfg(feature = "extra-children")]
        let mut extra_iter = self.extra_children.iter().peekable();
        #[cfg(feature = "extra-children")]
        let mut emit_idx: usize = 0;
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        #[cfg(feature = "dml-text")]
        if let Some(ref val) = self.p_pr {
            val.write_element("a:pPr", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        for item in &self.text_run {
            #[cfg(feature = "extra-children")]
            while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
                extra_iter
                    .next()
                    .unwrap()
                    .node
                    .write_to(writer)
                    .map_err(SerializeError::from)?;
            }
            item.write_element("", writer)?;
            #[cfg(feature = "extra-children")]
            {
                emit_idx += 1;
            }
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        #[cfg(feature = "dml-text")]
        if let Some(ref val) = self.end_para_r_pr {
            val.write_element("a:endParaRPr", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        for extra in extra_iter {
            extra.node.write_to(writer).map_err(SerializeError::from)?;
        }
        Ok(())
    }

    fn is_empty_element(&self) -> bool {
        #[cfg(feature = "dml-text")]
        if self.p_pr.is_some() {
            return false;
        }
        if !self.text_run.is_empty() {
            return false;
        }
        #[cfg(feature = "dml-text")]
        if self.end_para_r_pr.is_some() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for CTTextListStyle {
    fn write_children<W: Write>(&self, writer: &mut Writer<W>) -> Result<(), SerializeError> {
        #[cfg(feature = "extra-children")]
        let mut extra_iter = self.extra_children.iter().peekable();
        #[cfg(feature = "extra-children")]
        let mut emit_idx: usize = 0;
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        if let Some(ref val) = self.def_p_pr {
            val.write_element("a:defPPr", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        if let Some(ref val) = self.lvl1p_pr {
            val.write_element("a:lvl1pPr", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        if let Some(ref val) = self.lvl2p_pr {
            val.write_element("a:lvl2pPr", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        if let Some(ref val) = self.lvl3p_pr {
            val.write_element("a:lvl3pPr", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        if let Some(ref val) = self.lvl4p_pr {
            val.write_element("a:lvl4pPr", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        if let Some(ref val) = self.lvl5p_pr {
            val.write_element("a:lvl5pPr", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        if let Some(ref val) = self.lvl6p_pr {
            val.write_element("a:lvl6pPr", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        if let Some(ref val) = self.lvl7p_pr {
            val.write_element("a:lvl7pPr", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        if let Some(ref val) = self.lvl8p_pr {
            val.write_element("a:lvl8pPr", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        if let Some(ref val) = self.lvl9p_pr {
            val.write_element("a:lvl9pPr", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        if let Some(ref val) = self.ext_lst {
            val.write_element("a:extLst", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        for extra in extra_iter {
            extra.node.write_to(writer).map_err(SerializeError::from)?;
        }
        Ok(())
    }

    fn is_empty_element(&self) -> bool {
        if self.def_p_pr.is_some() {
            return false;
        }
        if self.lvl1p_pr.is_some() {
            return false;
        }
        if self.lvl2p_pr.is_some() {
            return false;
        }
        if self.lvl3p_pr.is_some() {
            return false;
        }
        if self.lvl4p_pr.is_some() {
            return false;
        }
        if self.lvl5p_pr.is_some() {
            return false;
        }
        if self.lvl6p_pr.is_some() {
            return false;
        }
        if self.lvl7p_pr.is_some() {
            return false;
        }
        if self.lvl8p_pr.is_some() {
            return false;
        }
        if self.lvl9p_pr.is_some() {
            return false;
        }
        if self.ext_lst.is_some() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for CTTextNormalAutofit {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        #[allow(unused_mut)]
        let mut start = start;
        if let Some(ref val) = self.font_scale {
            {
                let s = val.to_string();
                start.push_attribute(("fontScale", s.as_str()));
            }
        }
        if let Some(ref val) = self.ln_spc_reduction {
            {
                let s = val.to_string();
                start.push_attribute(("lnSpcReduction", s.as_str()));
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

impl ToXml for CTTextShapeAutofit {
    fn is_empty_element(&self) -> bool {
        true
    }
}

impl ToXml for CTTextNoAutofit {
    fn is_empty_element(&self) -> bool {
        true
    }
}

impl ToXml for EGTextAutofit {
    fn write_element<W: Write>(
        &self,
        _tag: &str,
        writer: &mut Writer<W>,
    ) -> Result<(), SerializeError> {
        match self {
            Self::NoAutofit(inner) => inner.write_element("a:noAutofit", writer)?,
            Self::NormAutofit(inner) => inner.write_element("a:normAutofit", writer)?,
            Self::SpAutoFit(inner) => inner.write_element("a:spAutoFit", writer)?,
        }
        Ok(())
    }
}

impl ToXml for CTTextBodyProperties {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        #[allow(unused_mut)]
        let mut start = start;
        #[cfg(feature = "dml-text")]
        if let Some(ref val) = self.rot {
            {
                let s = val.to_string();
                start.push_attribute(("rot", s.as_str()));
            }
        }
        #[cfg(feature = "dml-text")]
        if let Some(ref val) = self.spc_first_last_para {
            start.push_attribute(("spcFirstLastPara", if *val { "1" } else { "0" }));
        }
        #[cfg(feature = "dml-text")]
        if let Some(ref val) = self.vert_overflow {
            {
                let s = val.to_string();
                start.push_attribute(("vertOverflow", s.as_str()));
            }
        }
        #[cfg(feature = "dml-text")]
        if let Some(ref val) = self.horz_overflow {
            {
                let s = val.to_string();
                start.push_attribute(("horzOverflow", s.as_str()));
            }
        }
        #[cfg(feature = "dml-text")]
        if let Some(ref val) = self.vert {
            {
                let s = val.to_string();
                start.push_attribute(("vert", s.as_str()));
            }
        }
        #[cfg(feature = "dml-text")]
        if let Some(ref val) = self.wrap {
            {
                let s = val.to_string();
                start.push_attribute(("wrap", s.as_str()));
            }
        }
        #[cfg(feature = "dml-text")]
        if let Some(ref val) = self.l_ins {
            {
                let s = val.to_string();
                start.push_attribute(("lIns", s.as_str()));
            }
        }
        #[cfg(feature = "dml-text")]
        if let Some(ref val) = self.t_ins {
            {
                let s = val.to_string();
                start.push_attribute(("tIns", s.as_str()));
            }
        }
        #[cfg(feature = "dml-text")]
        if let Some(ref val) = self.r_ins {
            {
                let s = val.to_string();
                start.push_attribute(("rIns", s.as_str()));
            }
        }
        #[cfg(feature = "dml-text")]
        if let Some(ref val) = self.b_ins {
            {
                let s = val.to_string();
                start.push_attribute(("bIns", s.as_str()));
            }
        }
        #[cfg(feature = "dml-text")]
        if let Some(ref val) = self.num_col {
            {
                let s = val.to_string();
                start.push_attribute(("numCol", s.as_str()));
            }
        }
        #[cfg(feature = "dml-text")]
        if let Some(ref val) = self.spc_col {
            {
                let s = val.to_string();
                start.push_attribute(("spcCol", s.as_str()));
            }
        }
        #[cfg(feature = "dml-text")]
        if let Some(ref val) = self.rtl_col {
            start.push_attribute(("rtlCol", if *val { "1" } else { "0" }));
        }
        #[cfg(feature = "dml-text")]
        if let Some(ref val) = self.from_word_art {
            start.push_attribute(("fromWordArt", if *val { "1" } else { "0" }));
        }
        #[cfg(feature = "dml-text")]
        if let Some(ref val) = self.anchor {
            {
                let s = val.to_string();
                start.push_attribute(("anchor", s.as_str()));
            }
        }
        #[cfg(feature = "dml-text")]
        if let Some(ref val) = self.anchor_ctr {
            start.push_attribute(("anchorCtr", if *val { "1" } else { "0" }));
        }
        #[cfg(feature = "dml-text")]
        if let Some(ref val) = self.force_a_a {
            start.push_attribute(("forceAA", if *val { "1" } else { "0" }));
        }
        #[cfg(feature = "dml-text")]
        if let Some(ref val) = self.upright {
            start.push_attribute(("upright", if *val { "1" } else { "0" }));
        }
        #[cfg(feature = "dml-text")]
        if let Some(ref val) = self.compat_ln_spc {
            start.push_attribute(("compatLnSpc", if *val { "1" } else { "0" }));
        }
        #[cfg(feature = "extra-attrs")]
        for (key, value) in &self.extra_attrs {
            start.push_attribute((key.as_str(), value.as_str()));
        }
        start
    }

    fn write_children<W: Write>(&self, writer: &mut Writer<W>) -> Result<(), SerializeError> {
        #[cfg(feature = "extra-children")]
        let mut extra_iter = self.extra_children.iter().peekable();
        #[cfg(feature = "extra-children")]
        let mut emit_idx: usize = 0;
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        #[cfg(feature = "dml-text")]
        if let Some(ref val) = self.prst_tx_warp {
            val.write_element("a:prstTxWarp", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        if let Some(ref val) = self.text_autofit {
            val.write_element("", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        #[cfg(feature = "dml-3d")]
        if let Some(ref val) = self.scene3d {
            val.write_element("a:scene3d", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        if let Some(ref val) = self.text3_d {
            val.write_element("", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        #[cfg(feature = "dml-extensions")]
        if let Some(ref val) = self.ext_lst {
            val.write_element("a:extLst", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        for extra in extra_iter {
            extra.node.write_to(writer).map_err(SerializeError::from)?;
        }
        Ok(())
    }

    fn is_empty_element(&self) -> bool {
        #[cfg(feature = "dml-text")]
        if self.prst_tx_warp.is_some() {
            return false;
        }
        if self.text_autofit.is_some() {
            return false;
        }
        #[cfg(feature = "dml-3d")]
        if self.scene3d.is_some() {
            return false;
        }
        if self.text3_d.is_some() {
            return false;
        }
        #[cfg(feature = "dml-extensions")]
        if self.ext_lst.is_some() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for TextBody {
    fn write_children<W: Write>(&self, writer: &mut Writer<W>) -> Result<(), SerializeError> {
        #[cfg(feature = "extra-children")]
        let mut extra_iter = self.extra_children.iter().peekable();
        #[cfg(feature = "extra-children")]
        let mut emit_idx: usize = 0;
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        #[cfg(feature = "dml-text")]
        {
            let val = &self.body_pr;
            val.write_element("a:bodyPr", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        #[cfg(feature = "dml-text")]
        if let Some(ref val) = self.lst_style {
            val.write_element("a:lstStyle", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "dml-text")]
        for item in &self.p {
            #[cfg(feature = "extra-children")]
            while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
                extra_iter
                    .next()
                    .unwrap()
                    .node
                    .write_to(writer)
                    .map_err(SerializeError::from)?;
            }
            item.write_element("a:p", writer)?;
            #[cfg(feature = "extra-children")]
            {
                emit_idx += 1;
            }
        }
        #[cfg(feature = "extra-children")]
        for extra in extra_iter {
            extra.node.write_to(writer).map_err(SerializeError::from)?;
        }
        Ok(())
    }

    fn is_empty_element(&self) -> bool {
        #[cfg(feature = "dml-text")]
        return false;
        #[cfg(feature = "dml-text")]
        if self.lst_style.is_some() {
            return false;
        }
        #[cfg(feature = "dml-text")]
        if !self.p.is_empty() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for CTTextBulletColorFollowText {
    fn is_empty_element(&self) -> bool {
        true
    }
}

impl ToXml for EGTextBulletColor {
    fn write_element<W: Write>(
        &self,
        _tag: &str,
        writer: &mut Writer<W>,
    ) -> Result<(), SerializeError> {
        match self {
            Self::BuClrTx(inner) => inner.write_element("a:buClrTx", writer)?,
            Self::BuClr(inner) => inner.write_element("a:buClr", writer)?,
        }
        Ok(())
    }
}

impl ToXml for CTTextBulletSizeFollowText {
    fn is_empty_element(&self) -> bool {
        true
    }
}

impl ToXml for CTTextBulletSizePercent {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        #[allow(unused_mut)]
        let mut start = start;
        {
            let val = &self.value;
            start.push_attribute(("val", val.as_str()));
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

impl ToXml for CTTextBulletSizePoint {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        #[allow(unused_mut)]
        let mut start = start;
        {
            let val = &self.value;
            {
                let s = val.to_string();
                start.push_attribute(("val", s.as_str()));
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

impl ToXml for EGTextBulletSize {
    fn write_element<W: Write>(
        &self,
        _tag: &str,
        writer: &mut Writer<W>,
    ) -> Result<(), SerializeError> {
        match self {
            Self::BuSzTx(inner) => inner.write_element("a:buSzTx", writer)?,
            Self::BuSzPct(inner) => inner.write_element("a:buSzPct", writer)?,
            Self::BuSzPts(inner) => inner.write_element("a:buSzPts", writer)?,
        }
        Ok(())
    }
}

impl ToXml for CTTextBulletTypefaceFollowText {
    fn is_empty_element(&self) -> bool {
        true
    }
}

impl ToXml for EGTextBulletTypeface {
    fn write_element<W: Write>(
        &self,
        _tag: &str,
        writer: &mut Writer<W>,
    ) -> Result<(), SerializeError> {
        match self {
            Self::BuFontTx(inner) => inner.write_element("a:buFontTx", writer)?,
            Self::BuFont(inner) => inner.write_element("a:buFont", writer)?,
        }
        Ok(())
    }
}

impl ToXml for CTTextAutonumberBullet {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        #[allow(unused_mut)]
        let mut start = start;
        {
            let val = &self.r#type;
            {
                let s = val.to_string();
                start.push_attribute(("type", s.as_str()));
            }
        }
        if let Some(ref val) = self.start_at {
            {
                let s = val.to_string();
                start.push_attribute(("startAt", s.as_str()));
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

impl ToXml for CTTextCharBullet {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        #[allow(unused_mut)]
        let mut start = start;
        {
            let val = &self.char;
            start.push_attribute(("char", val.as_str()));
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

impl ToXml for CTTextNoBullet {
    fn is_empty_element(&self) -> bool {
        true
    }
}

impl ToXml for EGTextBullet {
    fn write_element<W: Write>(
        &self,
        _tag: &str,
        writer: &mut Writer<W>,
    ) -> Result<(), SerializeError> {
        match self {
            Self::BuNone(inner) => inner.write_element("a:buNone", writer)?,
            Self::BuAutoNum(inner) => inner.write_element("a:buAutoNum", writer)?,
            Self::BuChar(inner) => inner.write_element("a:buChar", writer)?,
            Self::BuBlip(inner) => inner.write_element("a:buBlip", writer)?,
        }
        Ok(())
    }
}

impl ToXml for TextFont {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        #[allow(unused_mut)]
        let mut start = start;
        #[cfg(feature = "dml-text")]
        {
            let val = &self.typeface;
            start.push_attribute(("typeface", val.as_str()));
        }
        #[cfg(feature = "dml-text")]
        if let Some(ref val) = self.panose {
            {
                let hex = encode_hex(val);
                start.push_attribute(("panose", hex.as_str()));
            }
        }
        #[cfg(feature = "dml-text")]
        if let Some(ref val) = self.pitch_family {
            {
                let s = val.to_string();
                start.push_attribute(("pitchFamily", s.as_str()));
            }
        }
        #[cfg(feature = "dml-text")]
        if let Some(ref val) = self.charset {
            {
                let s = val.to_string();
                start.push_attribute(("charset", s.as_str()));
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

impl ToXml for CTTextUnderlineLineFollowText {
    fn is_empty_element(&self) -> bool {
        true
    }
}

impl ToXml for CTTextUnderlineFillFollowText {
    fn is_empty_element(&self) -> bool {
        true
    }
}

impl ToXml for CTTextUnderlineFillGroupWrapper {
    fn write_children<W: Write>(&self, writer: &mut Writer<W>) -> Result<(), SerializeError> {
        #[cfg(feature = "extra-children")]
        let mut extra_iter = self.extra_children.iter().peekable();
        #[cfg(feature = "extra-children")]
        let mut emit_idx: usize = 0;
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        if let Some(ref val) = self.fill_properties {
            val.write_element("", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        for extra in extra_iter {
            extra.node.write_to(writer).map_err(SerializeError::from)?;
        }
        Ok(())
    }

    fn is_empty_element(&self) -> bool {
        false
    }
}

impl ToXml for EGTextUnderlineLine {
    fn write_element<W: Write>(
        &self,
        _tag: &str,
        writer: &mut Writer<W>,
    ) -> Result<(), SerializeError> {
        match self {
            Self::ULnTx(inner) => inner.write_element("a:uLnTx", writer)?,
            Self::ULn(inner) => inner.write_element("a:uLn", writer)?,
        }
        Ok(())
    }
}

impl ToXml for EGTextUnderlineFill {
    fn write_element<W: Write>(
        &self,
        _tag: &str,
        writer: &mut Writer<W>,
    ) -> Result<(), SerializeError> {
        match self {
            Self::UFillTx(inner) => inner.write_element("a:uFillTx", writer)?,
            Self::UFill(inner) => inner.write_element("a:uFill", writer)?,
        }
        Ok(())
    }
}

impl ToXml for TextCharacterProperties {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        #[allow(unused_mut)]
        let mut start = start;
        #[cfg(feature = "dml-text")]
        if let Some(ref val) = self.kumimoji {
            start.push_attribute(("kumimoji", if *val { "1" } else { "0" }));
        }
        #[cfg(feature = "dml-text")]
        if let Some(ref val) = self.lang {
            start.push_attribute(("lang", val.as_str()));
        }
        #[cfg(feature = "dml-text")]
        if let Some(ref val) = self.alt_lang {
            start.push_attribute(("altLang", val.as_str()));
        }
        #[cfg(feature = "dml-text")]
        if let Some(ref val) = self.sz {
            {
                let s = val.to_string();
                start.push_attribute(("sz", s.as_str()));
            }
        }
        #[cfg(feature = "dml-text")]
        if let Some(ref val) = self.b {
            start.push_attribute(("b", if *val { "1" } else { "0" }));
        }
        #[cfg(feature = "dml-text")]
        if let Some(ref val) = self.i {
            start.push_attribute(("i", if *val { "1" } else { "0" }));
        }
        #[cfg(feature = "dml-text")]
        if let Some(ref val) = self.u {
            {
                let s = val.to_string();
                start.push_attribute(("u", s.as_str()));
            }
        }
        #[cfg(feature = "dml-text")]
        if let Some(ref val) = self.strike {
            {
                let s = val.to_string();
                start.push_attribute(("strike", s.as_str()));
            }
        }
        #[cfg(feature = "dml-text")]
        if let Some(ref val) = self.kern {
            {
                let s = val.to_string();
                start.push_attribute(("kern", s.as_str()));
            }
        }
        #[cfg(feature = "dml-text")]
        if let Some(ref val) = self.cap {
            {
                let s = val.to_string();
                start.push_attribute(("cap", s.as_str()));
            }
        }
        #[cfg(feature = "dml-text")]
        if let Some(ref val) = self.spc {
            {
                let s = val.to_string();
                start.push_attribute(("spc", s.as_str()));
            }
        }
        #[cfg(feature = "dml-text")]
        if let Some(ref val) = self.normalize_h {
            start.push_attribute(("normalizeH", if *val { "1" } else { "0" }));
        }
        #[cfg(feature = "dml-text")]
        if let Some(ref val) = self.baseline {
            {
                let s = val.to_string();
                start.push_attribute(("baseline", s.as_str()));
            }
        }
        #[cfg(feature = "dml-text")]
        if let Some(ref val) = self.no_proof {
            start.push_attribute(("noProof", if *val { "1" } else { "0" }));
        }
        #[cfg(feature = "dml-text")]
        if let Some(ref val) = self.dirty {
            start.push_attribute(("dirty", if *val { "1" } else { "0" }));
        }
        #[cfg(feature = "dml-text")]
        if let Some(ref val) = self.err {
            start.push_attribute(("err", if *val { "1" } else { "0" }));
        }
        #[cfg(feature = "dml-text")]
        if let Some(ref val) = self.smt_clean {
            start.push_attribute(("smtClean", if *val { "1" } else { "0" }));
        }
        #[cfg(feature = "dml-text")]
        if let Some(ref val) = self.smt_id {
            {
                let s = val.to_string();
                start.push_attribute(("smtId", s.as_str()));
            }
        }
        #[cfg(feature = "dml-text")]
        if let Some(ref val) = self.bmk {
            start.push_attribute(("bmk", val.as_str()));
        }
        #[cfg(feature = "extra-attrs")]
        for (key, value) in &self.extra_attrs {
            start.push_attribute((key.as_str(), value.as_str()));
        }
        start
    }

    fn write_children<W: Write>(&self, writer: &mut Writer<W>) -> Result<(), SerializeError> {
        #[cfg(feature = "extra-children")]
        let mut extra_iter = self.extra_children.iter().peekable();
        #[cfg(feature = "extra-children")]
        let mut emit_idx: usize = 0;
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        #[cfg(feature = "dml-text")]
        if let Some(ref val) = self.line {
            val.write_element("a:ln", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        if let Some(ref val) = self.fill_properties {
            val.write_element("", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        if let Some(ref val) = self.effect_properties {
            val.write_element("", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        #[cfg(feature = "dml-text")]
        if let Some(ref val) = self.highlight {
            val.write_element("a:highlight", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        if let Some(ref val) = self.text_underline_line {
            val.write_element("", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        if let Some(ref val) = self.text_underline_fill {
            val.write_element("", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        #[cfg(feature = "dml-text")]
        if let Some(ref val) = self.latin {
            val.write_element("a:latin", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        #[cfg(feature = "dml-text")]
        if let Some(ref val) = self.ea {
            val.write_element("a:ea", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        #[cfg(feature = "dml-text")]
        if let Some(ref val) = self.cs {
            val.write_element("a:cs", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        #[cfg(feature = "dml-text")]
        if let Some(ref val) = self.sym {
            val.write_element("a:sym", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        #[cfg(feature = "dml-text")]
        if let Some(ref val) = self.hlink_click {
            val.write_element("a:hlinkClick", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        #[cfg(feature = "dml-text")]
        if let Some(ref val) = self.hlink_mouse_over {
            val.write_element("a:hlinkMouseOver", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        #[cfg(feature = "dml-text")]
        if let Some(ref val) = self.rtl {
            val.write_element("a:rtl", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        #[cfg(feature = "dml-extensions")]
        if let Some(ref val) = self.ext_lst {
            val.write_element("a:extLst", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        for extra in extra_iter {
            extra.node.write_to(writer).map_err(SerializeError::from)?;
        }
        Ok(())
    }

    fn is_empty_element(&self) -> bool {
        #[cfg(feature = "dml-text")]
        if self.line.is_some() {
            return false;
        }
        if self.fill_properties.is_some() {
            return false;
        }
        if self.effect_properties.is_some() {
            return false;
        }
        #[cfg(feature = "dml-text")]
        if self.highlight.is_some() {
            return false;
        }
        if self.text_underline_line.is_some() {
            return false;
        }
        if self.text_underline_fill.is_some() {
            return false;
        }
        #[cfg(feature = "dml-text")]
        if self.latin.is_some() {
            return false;
        }
        #[cfg(feature = "dml-text")]
        if self.ea.is_some() {
            return false;
        }
        #[cfg(feature = "dml-text")]
        if self.cs.is_some() {
            return false;
        }
        #[cfg(feature = "dml-text")]
        if self.sym.is_some() {
            return false;
        }
        #[cfg(feature = "dml-text")]
        if self.hlink_click.is_some() {
            return false;
        }
        #[cfg(feature = "dml-text")]
        if self.hlink_mouse_over.is_some() {
            return false;
        }
        #[cfg(feature = "dml-text")]
        if self.rtl.is_some() {
            return false;
        }
        #[cfg(feature = "dml-extensions")]
        if self.ext_lst.is_some() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for CTBoolean {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        #[allow(unused_mut)]
        let mut start = start;
        if let Some(ref val) = self.value {
            {
                let s = val.to_string();
                start.push_attribute(("val", s.as_str()));
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

impl ToXml for CTTextSpacingPercent {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        #[allow(unused_mut)]
        let mut start = start;
        {
            let val = &self.value;
            {
                let s = val.to_string();
                start.push_attribute(("val", s.as_str()));
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

impl ToXml for CTTextSpacingPoint {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        #[allow(unused_mut)]
        let mut start = start;
        {
            let val = &self.value;
            {
                let s = val.to_string();
                start.push_attribute(("val", s.as_str()));
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

impl ToXml for CTTextTabStop {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        #[allow(unused_mut)]
        let mut start = start;
        if let Some(ref val) = self.pos {
            {
                let s = val.to_string();
                start.push_attribute(("pos", s.as_str()));
            }
        }
        if let Some(ref val) = self.algn {
            {
                let s = val.to_string();
                start.push_attribute(("algn", s.as_str()));
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

impl ToXml for CTTextTabStopList {
    fn write_children<W: Write>(&self, writer: &mut Writer<W>) -> Result<(), SerializeError> {
        #[cfg(feature = "extra-children")]
        let mut extra_iter = self.extra_children.iter().peekable();
        #[cfg(feature = "extra-children")]
        let mut emit_idx: usize = 0;
        for item in &self.tab {
            #[cfg(feature = "extra-children")]
            while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
                extra_iter
                    .next()
                    .unwrap()
                    .node
                    .write_to(writer)
                    .map_err(SerializeError::from)?;
            }
            item.write_element("a:tab", writer)?;
            #[cfg(feature = "extra-children")]
            {
                emit_idx += 1;
            }
        }
        #[cfg(feature = "extra-children")]
        for extra in extra_iter {
            extra.node.write_to(writer).map_err(SerializeError::from)?;
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

impl ToXml for CTTextLineBreak {
    fn write_children<W: Write>(&self, writer: &mut Writer<W>) -> Result<(), SerializeError> {
        #[cfg(feature = "extra-children")]
        let mut extra_iter = self.extra_children.iter().peekable();
        #[cfg(feature = "extra-children")]
        let mut emit_idx: usize = 0;
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        if let Some(ref val) = self.r_pr {
            val.write_element("a:rPr", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        for extra in extra_iter {
            extra.node.write_to(writer).map_err(SerializeError::from)?;
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

impl ToXml for CTTextSpacing {
    fn write_children<W: Write>(&self, writer: &mut Writer<W>) -> Result<(), SerializeError> {
        #[cfg(feature = "extra-children")]
        let mut extra_iter = self.extra_children.iter().peekable();
        #[cfg(feature = "extra-children")]
        let mut emit_idx: usize = 0;
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        if let Some(ref val) = self.spc_pct {
            val.write_element("a:spcPct", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        if let Some(ref val) = self.spc_pts {
            val.write_element("a:spcPts", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        for extra in extra_iter {
            extra.node.write_to(writer).map_err(SerializeError::from)?;
        }
        Ok(())
    }

    fn is_empty_element(&self) -> bool {
        if self.spc_pct.is_some() {
            return false;
        }
        if self.spc_pts.is_some() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for TextParagraphProperties {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        #[allow(unused_mut)]
        let mut start = start;
        #[cfg(feature = "dml-text")]
        if let Some(ref val) = self.mar_l {
            {
                let s = val.to_string();
                start.push_attribute(("marL", s.as_str()));
            }
        }
        #[cfg(feature = "dml-text")]
        if let Some(ref val) = self.mar_r {
            {
                let s = val.to_string();
                start.push_attribute(("marR", s.as_str()));
            }
        }
        #[cfg(feature = "dml-text")]
        if let Some(ref val) = self.lvl {
            {
                let s = val.to_string();
                start.push_attribute(("lvl", s.as_str()));
            }
        }
        #[cfg(feature = "dml-text")]
        if let Some(ref val) = self.indent {
            {
                let s = val.to_string();
                start.push_attribute(("indent", s.as_str()));
            }
        }
        #[cfg(feature = "dml-text")]
        if let Some(ref val) = self.algn {
            {
                let s = val.to_string();
                start.push_attribute(("algn", s.as_str()));
            }
        }
        #[cfg(feature = "dml-text")]
        if let Some(ref val) = self.def_tab_sz {
            {
                let s = val.to_string();
                start.push_attribute(("defTabSz", s.as_str()));
            }
        }
        #[cfg(feature = "dml-text")]
        if let Some(ref val) = self.rtl {
            start.push_attribute(("rtl", if *val { "1" } else { "0" }));
        }
        #[cfg(feature = "dml-text")]
        if let Some(ref val) = self.ea_ln_brk {
            start.push_attribute(("eaLnBrk", if *val { "1" } else { "0" }));
        }
        #[cfg(feature = "dml-text")]
        if let Some(ref val) = self.font_algn {
            {
                let s = val.to_string();
                start.push_attribute(("fontAlgn", s.as_str()));
            }
        }
        #[cfg(feature = "dml-text")]
        if let Some(ref val) = self.latin_ln_brk {
            start.push_attribute(("latinLnBrk", if *val { "1" } else { "0" }));
        }
        #[cfg(feature = "dml-text")]
        if let Some(ref val) = self.hanging_punct {
            start.push_attribute(("hangingPunct", if *val { "1" } else { "0" }));
        }
        #[cfg(feature = "extra-attrs")]
        for (key, value) in &self.extra_attrs {
            start.push_attribute((key.as_str(), value.as_str()));
        }
        start
    }

    fn write_children<W: Write>(&self, writer: &mut Writer<W>) -> Result<(), SerializeError> {
        #[cfg(feature = "extra-children")]
        let mut extra_iter = self.extra_children.iter().peekable();
        #[cfg(feature = "extra-children")]
        let mut emit_idx: usize = 0;
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        #[cfg(feature = "dml-text")]
        if let Some(ref val) = self.ln_spc {
            val.write_element("a:lnSpc", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        #[cfg(feature = "dml-text")]
        if let Some(ref val) = self.spc_bef {
            val.write_element("a:spcBef", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        #[cfg(feature = "dml-text")]
        if let Some(ref val) = self.spc_aft {
            val.write_element("a:spcAft", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        if let Some(ref val) = self.text_bullet_color {
            val.write_element("", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        if let Some(ref val) = self.text_bullet_size {
            val.write_element("", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        if let Some(ref val) = self.text_bullet_typeface {
            val.write_element("", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        if let Some(ref val) = self.text_bullet {
            val.write_element("", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        #[cfg(feature = "dml-text")]
        if let Some(ref val) = self.tab_lst {
            val.write_element("a:tabLst", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        #[cfg(feature = "dml-text")]
        if let Some(ref val) = self.def_r_pr {
            val.write_element("a:defRPr", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        #[cfg(feature = "dml-extensions")]
        if let Some(ref val) = self.ext_lst {
            val.write_element("a:extLst", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        for extra in extra_iter {
            extra.node.write_to(writer).map_err(SerializeError::from)?;
        }
        Ok(())
    }

    fn is_empty_element(&self) -> bool {
        #[cfg(feature = "dml-text")]
        if self.ln_spc.is_some() {
            return false;
        }
        #[cfg(feature = "dml-text")]
        if self.spc_bef.is_some() {
            return false;
        }
        #[cfg(feature = "dml-text")]
        if self.spc_aft.is_some() {
            return false;
        }
        if self.text_bullet_color.is_some() {
            return false;
        }
        if self.text_bullet_size.is_some() {
            return false;
        }
        if self.text_bullet_typeface.is_some() {
            return false;
        }
        if self.text_bullet.is_some() {
            return false;
        }
        #[cfg(feature = "dml-text")]
        if self.tab_lst.is_some() {
            return false;
        }
        #[cfg(feature = "dml-text")]
        if self.def_r_pr.is_some() {
            return false;
        }
        #[cfg(feature = "dml-extensions")]
        if self.ext_lst.is_some() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for CTTextField {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        #[allow(unused_mut)]
        let mut start = start;
        {
            let val = &self.id;
            start.push_attribute(("id", val.as_str()));
        }
        if let Some(ref val) = self.r#type {
            start.push_attribute(("type", val.as_str()));
        }
        #[cfg(feature = "extra-attrs")]
        for (key, value) in &self.extra_attrs {
            start.push_attribute((key.as_str(), value.as_str()));
        }
        start
    }

    fn write_children<W: Write>(&self, writer: &mut Writer<W>) -> Result<(), SerializeError> {
        #[cfg(feature = "extra-children")]
        let mut extra_iter = self.extra_children.iter().peekable();
        #[cfg(feature = "extra-children")]
        let mut emit_idx: usize = 0;
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        if let Some(ref val) = self.r_pr {
            val.write_element("a:rPr", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        if let Some(ref val) = self.p_pr {
            val.write_element("a:pPr", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        if let Some(ref val) = self.t {
            {
                let start = BytesStart::new("a:t");
                writer.write_event(Event::Start(start))?;
                writer.write_event(Event::Text(BytesText::new(val.as_str())))?;
                writer.write_event(Event::End(BytesEnd::new("a:t")))?;
            }
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        for extra in extra_iter {
            extra.node.write_to(writer).map_err(SerializeError::from)?;
        }
        Ok(())
    }

    fn is_empty_element(&self) -> bool {
        if self.r_pr.is_some() {
            return false;
        }
        if self.p_pr.is_some() {
            return false;
        }
        if self.t.is_some() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for EGTextRun {
    fn write_element<W: Write>(
        &self,
        _tag: &str,
        writer: &mut Writer<W>,
    ) -> Result<(), SerializeError> {
        match self {
            Self::R(inner) => inner.write_element("a:r", writer)?,
            Self::Br(inner) => inner.write_element("a:br", writer)?,
            Self::Fld(inner) => inner.write_element("a:fld", writer)?,
        }
        Ok(())
    }
}

impl ToXml for TextRun {
    fn write_children<W: Write>(&self, writer: &mut Writer<W>) -> Result<(), SerializeError> {
        #[cfg(feature = "extra-children")]
        let mut extra_iter = self.extra_children.iter().peekable();
        #[cfg(feature = "extra-children")]
        let mut emit_idx: usize = 0;
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        #[cfg(feature = "dml-text")]
        if let Some(ref val) = self.r_pr {
            val.write_element("a:rPr", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        #[cfg(feature = "dml-text")]
        {
            let val = &self.t;
            {
                let start = BytesStart::new("a:t");
                writer.write_event(Event::Start(start))?;
                writer.write_event(Event::Text(BytesText::new(val.as_str())))?;
                writer.write_event(Event::End(BytesEnd::new("a:t")))?;
            }
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        for extra in extra_iter {
            extra.node.write_to(writer).map_err(SerializeError::from)?;
        }
        Ok(())
    }

    fn is_empty_element(&self) -> bool {
        #[cfg(feature = "dml-text")]
        if self.r_pr.is_some() {
            return false;
        }
        #[cfg(feature = "dml-text")]
        return false;
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}
