// ToXml serializers for generated types.
// Enables roundtrip XML serialization alongside FromXml parsers.

#![allow(unused_variables, unused_assignments, unreachable_code, unused_imports)]
#![allow(clippy::single_match)]
#![allow(clippy::match_single_binding)]
#![allow(clippy::explicit_counter_loop)]

use super::generated::*;
use ooxml_dml::types::*;
pub use ooxml_xml::{SerializeError, ToXml};
use quick_xml::Writer;
use quick_xml::events::{BytesEnd, BytesStart, BytesText, Event};
use std::io::Write;

#[allow(dead_code)]
/// Encode bytes as a hex string.
fn encode_hex(bytes: &[u8]) -> String {
    bytes.iter().map(|b| format!("{:02X}", b)).collect()
}

impl ToXml for CTSideDirectionTransition {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
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

    fn is_empty_element(&self) -> bool {
        true
    }
}

impl ToXml for CTCornerDirectionTransition {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
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

    fn is_empty_element(&self) -> bool {
        true
    }
}

impl ToXml for CTEightDirectionTransition {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
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

    fn is_empty_element(&self) -> bool {
        true
    }
}

impl ToXml for CTOrientationTransition {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
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

    fn is_empty_element(&self) -> bool {
        true
    }
}

impl ToXml for CTInOutTransition {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
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

    fn is_empty_element(&self) -> bool {
        true
    }
}

impl ToXml for CTOptionalBlackTransition {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        if let Some(ref val) = self.thru_blk {
            start.push_attribute(("thruBlk", if *val { "1" } else { "0" }));
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

impl ToXml for CTSplitTransition {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        if let Some(ref val) = self.orient {
            {
                let s = val.to_string();
                start.push_attribute(("orient", s.as_str()));
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

    fn is_empty_element(&self) -> bool {
        true
    }
}

impl ToXml for CTWheelTransition {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        if let Some(ref val) = self.spokes {
            {
                let s = val.to_string();
                start.push_attribute(("spokes", s.as_str()));
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

impl ToXml for CTTransitionStartSoundAction {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        if let Some(ref val) = self.r#loop {
            start.push_attribute(("loop", if *val { "1" } else { "0" }));
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
            let val = &self.snd;
            val.write_element("p:snd", writer)?;
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

impl ToXml for CTTransitionSoundAction {
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
        if let Some(ref val) = self.st_snd {
            val.write_element("p:stSnd", writer)?;
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
        if let Some(ref val) = self.end_snd {
            val.write_element("p:endSnd", writer)?;
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
        if self.st_snd.is_some() {
            return false;
        }
        if self.end_snd.is_some() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for SlideTransition {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        if let Some(ref val) = self.spd {
            {
                let s = val.to_string();
                start.push_attribute(("spd", s.as_str()));
            }
        }
        if let Some(ref val) = self.adv_click {
            start.push_attribute(("advClick", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.adv_tm {
            {
                let s = val.to_string();
                start.push_attribute(("advTm", s.as_str()));
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
        if let Some(ref val) = self.blinds {
            val.write_element("p:blinds", writer)?;
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
        if let Some(ref val) = self.checker {
            val.write_element("p:checker", writer)?;
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
        if let Some(ref val) = self.circle {
            val.write_element("p:circle", writer)?;
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
        if let Some(ref val) = self.dissolve {
            val.write_element("p:dissolve", writer)?;
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
        if let Some(ref val) = self.comb {
            val.write_element("p:comb", writer)?;
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
        if let Some(ref val) = self.cover {
            val.write_element("p:cover", writer)?;
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
        if let Some(ref val) = self.cut {
            val.write_element("p:cut", writer)?;
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
        if let Some(ref val) = self.diamond {
            val.write_element("p:diamond", writer)?;
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
        if let Some(ref val) = self.fade {
            val.write_element("p:fade", writer)?;
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
        if let Some(ref val) = self.newsflash {
            val.write_element("p:newsflash", writer)?;
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
        if let Some(ref val) = self.plus {
            val.write_element("p:plus", writer)?;
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
        if let Some(ref val) = self.pull {
            val.write_element("p:pull", writer)?;
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
        if let Some(ref val) = self.push {
            val.write_element("p:push", writer)?;
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
        if let Some(ref val) = self.random {
            val.write_element("p:random", writer)?;
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
        if let Some(ref val) = self.random_bar {
            val.write_element("p:randomBar", writer)?;
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
        if let Some(ref val) = self.split {
            val.write_element("p:split", writer)?;
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
        if let Some(ref val) = self.strips {
            val.write_element("p:strips", writer)?;
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
        if let Some(ref val) = self.wedge {
            val.write_element("p:wedge", writer)?;
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
        if let Some(ref val) = self.wheel {
            val.write_element("p:wheel", writer)?;
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
        if let Some(ref val) = self.wipe {
            val.write_element("p:wipe", writer)?;
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
        if let Some(ref val) = self.zoom {
            val.write_element("p:zoom", writer)?;
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
        if let Some(ref val) = self.snd_ac {
            val.write_element("p:sndAc", writer)?;
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
            val.write_element("p:extLst", writer)?;
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
        if self.blinds.is_some() {
            return false;
        }
        if self.checker.is_some() {
            return false;
        }
        if self.circle.is_some() {
            return false;
        }
        if self.dissolve.is_some() {
            return false;
        }
        if self.comb.is_some() {
            return false;
        }
        if self.cover.is_some() {
            return false;
        }
        if self.cut.is_some() {
            return false;
        }
        if self.diamond.is_some() {
            return false;
        }
        if self.fade.is_some() {
            return false;
        }
        if self.newsflash.is_some() {
            return false;
        }
        if self.plus.is_some() {
            return false;
        }
        if self.pull.is_some() {
            return false;
        }
        if self.push.is_some() {
            return false;
        }
        if self.random.is_some() {
            return false;
        }
        if self.random_bar.is_some() {
            return false;
        }
        if self.split.is_some() {
            return false;
        }
        if self.strips.is_some() {
            return false;
        }
        if self.wedge.is_some() {
            return false;
        }
        if self.wheel.is_some() {
            return false;
        }
        if self.wipe.is_some() {
            return false;
        }
        if self.zoom.is_some() {
            return false;
        }
        if self.snd_ac.is_some() {
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

impl ToXml for CTTLIterateIntervalTime {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
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

impl ToXml for CTTLIterateIntervalPercentage {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
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

impl ToXml for CTTLIterateData {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        if let Some(ref val) = self.r#type {
            {
                let s = val.to_string();
                start.push_attribute(("type", s.as_str()));
            }
        }
        if let Some(ref val) = self.backwards {
            start.push_attribute(("backwards", if *val { "1" } else { "0" }));
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
        if let Some(ref val) = self.tm_abs {
            val.write_element("p:tmAbs", writer)?;
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
        if let Some(ref val) = self.tm_pct {
            val.write_element("p:tmPct", writer)?;
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
        if self.tm_abs.is_some() {
            return false;
        }
        if self.tm_pct.is_some() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for CTTLSubShapeId {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        {
            let val = &self.spid;
            {
                let s = val.to_string();
                start.push_attribute(("spid", s.as_str()));
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

impl ToXml for CTTLTextTargetElement {
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
        if let Some(ref val) = self.char_rg {
            val.write_element("p:charRg", writer)?;
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
        if let Some(ref val) = self.p_rg {
            val.write_element("p:pRg", writer)?;
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
        if self.char_rg.is_some() {
            return false;
        }
        if self.p_rg.is_some() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for CTTLOleChartTargetElement {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        {
            let val = &self.r#type;
            {
                let s = val.to_string();
                start.push_attribute(("type", s.as_str()));
            }
        }
        if let Some(ref val) = self.lvl {
            {
                let s = val.to_string();
                start.push_attribute(("lvl", s.as_str()));
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

impl ToXml for CTTLShapeTargetElement {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        {
            let val = &self.spid;
            {
                let s = val.to_string();
                start.push_attribute(("spid", s.as_str()));
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
        if let Some(ref val) = self.bg {
            val.write_element("p:bg", writer)?;
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
        if let Some(ref val) = self.sub_sp {
            val.write_element("p:subSp", writer)?;
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
        if let Some(ref val) = self.ole_chart_el {
            val.write_element("p:oleChartEl", writer)?;
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
        if let Some(ref val) = self.tx_el {
            val.write_element("p:txEl", writer)?;
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
        if let Some(ref val) = self.graphic_el {
            val.write_element("p:graphicEl", writer)?;
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
        if self.bg.is_some() {
            return false;
        }
        if self.sub_sp.is_some() {
            return false;
        }
        if self.ole_chart_el.is_some() {
            return false;
        }
        if self.tx_el.is_some() {
            return false;
        }
        if self.graphic_el.is_some() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for CTTLTimeTargetElement {
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
        if let Some(ref val) = self.sld_tgt {
            val.write_element("p:sldTgt", writer)?;
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
        if let Some(ref val) = self.snd_tgt {
            val.write_element("p:sndTgt", writer)?;
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
        if let Some(ref val) = self.sp_tgt {
            val.write_element("p:spTgt", writer)?;
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
        if let Some(ref val) = self.ink_tgt {
            val.write_element("p:inkTgt", writer)?;
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
        if self.sld_tgt.is_some() {
            return false;
        }
        if self.snd_tgt.is_some() {
            return false;
        }
        if self.sp_tgt.is_some() {
            return false;
        }
        if self.ink_tgt.is_some() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for CTTLTriggerTimeNodeID {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
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

impl ToXml for CTTLTriggerRuntimeNode {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
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

impl ToXml for CTTLTimeCondition {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        if let Some(ref val) = self.evt {
            {
                let s = val.to_string();
                start.push_attribute(("evt", s.as_str()));
            }
        }
        if let Some(ref val) = self.delay {
            {
                let s = val.to_string();
                start.push_attribute(("delay", s.as_str()));
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
        if let Some(ref val) = self.tgt_el {
            val.write_element("p:tgtEl", writer)?;
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
        if let Some(ref val) = self.tn {
            val.write_element("p:tn", writer)?;
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
        if let Some(ref val) = self.rtn {
            val.write_element("p:rtn", writer)?;
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
        if self.tgt_el.is_some() {
            return false;
        }
        if self.tn.is_some() {
            return false;
        }
        if self.rtn.is_some() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for CTTLTimeConditionList {
    fn write_children<W: Write>(&self, writer: &mut Writer<W>) -> Result<(), SerializeError> {
        #[cfg(feature = "extra-children")]
        let mut extra_iter = self.extra_children.iter().peekable();
        #[cfg(feature = "extra-children")]
        let mut emit_idx: usize = 0;
        for item in &self.cond {
            #[cfg(feature = "extra-children")]
            while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
                extra_iter
                    .next()
                    .unwrap()
                    .node
                    .write_to(writer)
                    .map_err(SerializeError::from)?;
            }
            item.write_element("p:cond", writer)?;
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
        if !self.cond.is_empty() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for CTTimeNodeList {
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
        if let Some(ref val) = self.par {
            val.write_element("p:par", writer)?;
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
        if let Some(ref val) = self.seq {
            val.write_element("p:seq", writer)?;
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
        if let Some(ref val) = self.excl {
            val.write_element("p:excl", writer)?;
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
        if let Some(ref val) = self.anim {
            val.write_element("p:anim", writer)?;
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
        if let Some(ref val) = self.anim_clr {
            val.write_element("p:animClr", writer)?;
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
        if let Some(ref val) = self.anim_effect {
            val.write_element("p:animEffect", writer)?;
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
        if let Some(ref val) = self.anim_motion {
            val.write_element("p:animMotion", writer)?;
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
        if let Some(ref val) = self.anim_rot {
            val.write_element("p:animRot", writer)?;
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
        if let Some(ref val) = self.anim_scale {
            val.write_element("p:animScale", writer)?;
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
        if let Some(ref val) = self.cmd {
            val.write_element("p:cmd", writer)?;
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
        if let Some(ref val) = self.set {
            val.write_element("p:set", writer)?;
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
        if let Some(ref val) = self.audio {
            val.write_element("p:audio", writer)?;
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
        if let Some(ref val) = self.video {
            val.write_element("p:video", writer)?;
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
        if self.par.is_some() {
            return false;
        }
        if self.seq.is_some() {
            return false;
        }
        if self.excl.is_some() {
            return false;
        }
        if self.anim.is_some() {
            return false;
        }
        if self.anim_clr.is_some() {
            return false;
        }
        if self.anim_effect.is_some() {
            return false;
        }
        if self.anim_motion.is_some() {
            return false;
        }
        if self.anim_rot.is_some() {
            return false;
        }
        if self.anim_scale.is_some() {
            return false;
        }
        if self.cmd.is_some() {
            return false;
        }
        if self.set.is_some() {
            return false;
        }
        if self.audio.is_some() {
            return false;
        }
        if self.video.is_some() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for CTTLCommonTimeNodeData {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        if let Some(ref val) = self.id {
            {
                let s = val.to_string();
                start.push_attribute(("id", s.as_str()));
            }
        }
        if let Some(ref val) = self.preset_i_d {
            {
                let s = val.to_string();
                start.push_attribute(("presetID", s.as_str()));
            }
        }
        if let Some(ref val) = self.preset_class {
            {
                let s = val.to_string();
                start.push_attribute(("presetClass", s.as_str()));
            }
        }
        if let Some(ref val) = self.preset_subtype {
            {
                let s = val.to_string();
                start.push_attribute(("presetSubtype", s.as_str()));
            }
        }
        if let Some(ref val) = self.dur {
            {
                let s = val.to_string();
                start.push_attribute(("dur", s.as_str()));
            }
        }
        if let Some(ref val) = self.repeat_count {
            {
                let s = val.to_string();
                start.push_attribute(("repeatCount", s.as_str()));
            }
        }
        if let Some(ref val) = self.repeat_dur {
            {
                let s = val.to_string();
                start.push_attribute(("repeatDur", s.as_str()));
            }
        }
        if let Some(ref val) = self.spd {
            {
                let s = val.to_string();
                start.push_attribute(("spd", s.as_str()));
            }
        }
        if let Some(ref val) = self.accel {
            {
                let s = val.to_string();
                start.push_attribute(("accel", s.as_str()));
            }
        }
        if let Some(ref val) = self.decel {
            {
                let s = val.to_string();
                start.push_attribute(("decel", s.as_str()));
            }
        }
        if let Some(ref val) = self.auto_rev {
            start.push_attribute(("autoRev", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.restart {
            {
                let s = val.to_string();
                start.push_attribute(("restart", s.as_str()));
            }
        }
        if let Some(ref val) = self.fill {
            {
                let s = val.to_string();
                start.push_attribute(("fill", s.as_str()));
            }
        }
        if let Some(ref val) = self.sync_behavior {
            {
                let s = val.to_string();
                start.push_attribute(("syncBehavior", s.as_str()));
            }
        }
        if let Some(ref val) = self.tm_filter {
            start.push_attribute(("tmFilter", val.as_str()));
        }
        if let Some(ref val) = self.evt_filter {
            start.push_attribute(("evtFilter", val.as_str()));
        }
        if let Some(ref val) = self.display {
            start.push_attribute(("display", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.master_rel {
            {
                let s = val.to_string();
                start.push_attribute(("masterRel", s.as_str()));
            }
        }
        if let Some(ref val) = self.bld_lvl {
            {
                let s = val.to_string();
                start.push_attribute(("bldLvl", s.as_str()));
            }
        }
        if let Some(ref val) = self.grp_id {
            {
                let s = val.to_string();
                start.push_attribute(("grpId", s.as_str()));
            }
        }
        if let Some(ref val) = self.after_effect {
            start.push_attribute(("afterEffect", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.node_type {
            {
                let s = val.to_string();
                start.push_attribute(("nodeType", s.as_str()));
            }
        }
        if let Some(ref val) = self.node_ph {
            start.push_attribute(("nodePh", if *val { "1" } else { "0" }));
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
        if let Some(ref val) = self.st_cond_lst {
            val.write_element("p:stCondLst", writer)?;
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
        if let Some(ref val) = self.end_cond_lst {
            val.write_element("p:endCondLst", writer)?;
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
        if let Some(ref val) = self.end_sync {
            val.write_element("p:endSync", writer)?;
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
        if let Some(ref val) = self.iterate {
            val.write_element("p:iterate", writer)?;
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
        if let Some(ref val) = self.child_tn_lst {
            val.write_element("p:childTnLst", writer)?;
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
        if let Some(ref val) = self.sub_tn_lst {
            val.write_element("p:subTnLst", writer)?;
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
        if self.st_cond_lst.is_some() {
            return false;
        }
        if self.end_cond_lst.is_some() {
            return false;
        }
        if self.end_sync.is_some() {
            return false;
        }
        if self.iterate.is_some() {
            return false;
        }
        if self.child_tn_lst.is_some() {
            return false;
        }
        if self.sub_tn_lst.is_some() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for CTTLTimeNodeSequence {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        if let Some(ref val) = self.concurrent {
            start.push_attribute(("concurrent", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.prev_ac {
            {
                let s = val.to_string();
                start.push_attribute(("prevAc", s.as_str()));
            }
        }
        if let Some(ref val) = self.next_ac {
            {
                let s = val.to_string();
                start.push_attribute(("nextAc", s.as_str()));
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
            let val = &self.c_tn;
            val.write_element("p:cTn", writer)?;
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
        if let Some(ref val) = self.prev_cond_lst {
            val.write_element("p:prevCondLst", writer)?;
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
        if let Some(ref val) = self.next_cond_lst {
            val.write_element("p:nextCondLst", writer)?;
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

impl ToXml for CTTLBehaviorAttributeNameList {
    fn write_children<W: Write>(&self, writer: &mut Writer<W>) -> Result<(), SerializeError> {
        #[cfg(feature = "extra-children")]
        let mut extra_iter = self.extra_children.iter().peekable();
        #[cfg(feature = "extra-children")]
        let mut emit_idx: usize = 0;
        for item in &self.attr_name {
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
                let start = BytesStart::new("p:attrName");
                writer.write_event(Event::Start(start))?;
                writer.write_event(Event::Text(BytesText::new(item.as_str())))?;
                writer.write_event(Event::End(BytesEnd::new("p:attrName")))?;
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
        if !self.attr_name.is_empty() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for CTTLCommonBehaviorData {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        if let Some(ref val) = self.additive {
            {
                let s = val.to_string();
                start.push_attribute(("additive", s.as_str()));
            }
        }
        if let Some(ref val) = self.accumulate {
            {
                let s = val.to_string();
                start.push_attribute(("accumulate", s.as_str()));
            }
        }
        if let Some(ref val) = self.xfrm_type {
            {
                let s = val.to_string();
                start.push_attribute(("xfrmType", s.as_str()));
            }
        }
        if let Some(ref val) = self.from {
            start.push_attribute(("from", val.as_str()));
        }
        if let Some(ref val) = self.to {
            start.push_attribute(("to", val.as_str()));
        }
        if let Some(ref val) = self.by {
            start.push_attribute(("by", val.as_str()));
        }
        if let Some(ref val) = self.rctx {
            start.push_attribute(("rctx", val.as_str()));
        }
        if let Some(ref val) = self.r#override {
            {
                let s = val.to_string();
                start.push_attribute(("override", s.as_str()));
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
            let val = &self.c_tn;
            val.write_element("p:cTn", writer)?;
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
            let val = &self.tgt_el;
            val.write_element("p:tgtEl", writer)?;
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
        if let Some(ref val) = self.attr_name_lst {
            val.write_element("p:attrNameLst", writer)?;
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

impl ToXml for CTTLAnimVariantBooleanVal {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        {
            let val = &self.value;
            start.push_attribute(("val", if *val { "1" } else { "0" }));
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

impl ToXml for CTTLAnimVariantIntegerVal {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
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

impl ToXml for CTTLAnimVariantFloatVal {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
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

impl ToXml for CTTLAnimVariantStringVal {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
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

impl ToXml for CTTLAnimVariant {
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
        if let Some(ref val) = self.bool_val {
            val.write_element("p:boolVal", writer)?;
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
        if let Some(ref val) = self.int_val {
            val.write_element("p:intVal", writer)?;
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
        if let Some(ref val) = self.flt_val {
            val.write_element("p:fltVal", writer)?;
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
        if let Some(ref val) = self.str_val {
            val.write_element("p:strVal", writer)?;
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
        if let Some(ref val) = self.clr_val {
            val.write_element("p:clrVal", writer)?;
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
        if self.bool_val.is_some() {
            return false;
        }
        if self.int_val.is_some() {
            return false;
        }
        if self.flt_val.is_some() {
            return false;
        }
        if self.str_val.is_some() {
            return false;
        }
        if self.clr_val.is_some() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for CTTLTimeAnimateValue {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        if let Some(ref val) = self.tm {
            {
                let s = val.to_string();
                start.push_attribute(("tm", s.as_str()));
            }
        }
        if let Some(ref val) = self.fmla {
            start.push_attribute(("fmla", val.as_str()));
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
        if let Some(ref val) = self.value {
            val.write_element("p:val", writer)?;
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
        if self.value.is_some() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for CTTLTimeAnimateValueList {
    fn write_children<W: Write>(&self, writer: &mut Writer<W>) -> Result<(), SerializeError> {
        #[cfg(feature = "extra-children")]
        let mut extra_iter = self.extra_children.iter().peekable();
        #[cfg(feature = "extra-children")]
        let mut emit_idx: usize = 0;
        for item in &self.tav {
            #[cfg(feature = "extra-children")]
            while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
                extra_iter
                    .next()
                    .unwrap()
                    .node
                    .write_to(writer)
                    .map_err(SerializeError::from)?;
            }
            item.write_element("p:tav", writer)?;
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
        if !self.tav.is_empty() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for CTTLAnimateBehavior {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        if let Some(ref val) = self.by {
            start.push_attribute(("by", val.as_str()));
        }
        if let Some(ref val) = self.from {
            start.push_attribute(("from", val.as_str()));
        }
        if let Some(ref val) = self.to {
            start.push_attribute(("to", val.as_str()));
        }
        if let Some(ref val) = self.calcmode {
            {
                let s = val.to_string();
                start.push_attribute(("calcmode", s.as_str()));
            }
        }
        if let Some(ref val) = self.value_type {
            {
                let s = val.to_string();
                start.push_attribute(("valueType", s.as_str()));
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
            let val = &self.c_bhvr;
            val.write_element("p:cBhvr", writer)?;
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
        if let Some(ref val) = self.tav_lst {
            val.write_element("p:tavLst", writer)?;
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

impl ToXml for CTTLByRgbColorTransform {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        {
            let val = &self.reference;
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

    fn is_empty_element(&self) -> bool {
        true
    }
}

impl ToXml for CTTLByHslColorTransform {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        {
            let val = &self.height;
            {
                let s = val.to_string();
                start.push_attribute(("h", s.as_str()));
            }
        }
        {
            let val = &self.s;
            {
                let s = val.to_string();
                start.push_attribute(("s", s.as_str()));
            }
        }
        {
            let val = &self.l;
            {
                let s = val.to_string();
                start.push_attribute(("l", s.as_str()));
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

impl ToXml for CTTLByAnimateColorTransform {
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
        if let Some(ref val) = self.rgb {
            val.write_element("p:rgb", writer)?;
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
        if let Some(ref val) = self.hsl {
            val.write_element("p:hsl", writer)?;
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
        if self.rgb.is_some() {
            return false;
        }
        if self.hsl.is_some() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for CTTLAnimateColorBehavior {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        if let Some(ref val) = self.clr_spc {
            {
                let s = val.to_string();
                start.push_attribute(("clrSpc", s.as_str()));
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
        {
            let val = &self.c_bhvr;
            val.write_element("p:cBhvr", writer)?;
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
        if let Some(ref val) = self.by {
            val.write_element("p:by", writer)?;
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
        if let Some(ref val) = self.from {
            val.write_element("p:from", writer)?;
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
        if let Some(ref val) = self.to {
            val.write_element("p:to", writer)?;
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

impl ToXml for CTTLAnimateEffectBehavior {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        if let Some(ref val) = self.transition {
            {
                let s = val.to_string();
                start.push_attribute(("transition", s.as_str()));
            }
        }
        if let Some(ref val) = self.filter {
            start.push_attribute(("filter", val.as_str()));
        }
        if let Some(ref val) = self.pr_lst {
            start.push_attribute(("prLst", val.as_str()));
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
            let val = &self.c_bhvr;
            val.write_element("p:cBhvr", writer)?;
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
        if let Some(ref val) = self.progress {
            val.write_element("p:progress", writer)?;
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

impl ToXml for CTTLPoint {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
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

impl ToXml for CTTLAnimateMotionBehavior {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        if let Some(ref val) = self.origin {
            {
                let s = val.to_string();
                start.push_attribute(("origin", s.as_str()));
            }
        }
        if let Some(ref val) = self.path {
            start.push_attribute(("path", val.as_str()));
        }
        if let Some(ref val) = self.path_edit_mode {
            {
                let s = val.to_string();
                start.push_attribute(("pathEditMode", s.as_str()));
            }
        }
        if let Some(ref val) = self.r_ang {
            {
                let s = val.to_string();
                start.push_attribute(("rAng", s.as_str()));
            }
        }
        if let Some(ref val) = self.pts_types {
            start.push_attribute(("ptsTypes", val.as_str()));
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
            let val = &self.c_bhvr;
            val.write_element("p:cBhvr", writer)?;
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
        if let Some(ref val) = self.by {
            val.write_element("p:by", writer)?;
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
        if let Some(ref val) = self.from {
            val.write_element("p:from", writer)?;
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
        if let Some(ref val) = self.to {
            val.write_element("p:to", writer)?;
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
        if let Some(ref val) = self.r_ctr {
            val.write_element("p:rCtr", writer)?;
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

impl ToXml for CTTLAnimateRotationBehavior {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        if let Some(ref val) = self.by {
            {
                let s = val.to_string();
                start.push_attribute(("by", s.as_str()));
            }
        }
        if let Some(ref val) = self.from {
            {
                let s = val.to_string();
                start.push_attribute(("from", s.as_str()));
            }
        }
        if let Some(ref val) = self.to {
            {
                let s = val.to_string();
                start.push_attribute(("to", s.as_str()));
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
            let val = &self.c_bhvr;
            val.write_element("p:cBhvr", writer)?;
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

impl ToXml for CTTLAnimateScaleBehavior {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        if let Some(ref val) = self.zoom_contents {
            start.push_attribute(("zoomContents", if *val { "1" } else { "0" }));
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
            let val = &self.c_bhvr;
            val.write_element("p:cBhvr", writer)?;
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
        if let Some(ref val) = self.by {
            val.write_element("p:by", writer)?;
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
        if let Some(ref val) = self.from {
            val.write_element("p:from", writer)?;
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
        if let Some(ref val) = self.to {
            val.write_element("p:to", writer)?;
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

impl ToXml for CTTLCommandBehavior {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        if let Some(ref val) = self.r#type {
            {
                let s = val.to_string();
                start.push_attribute(("type", s.as_str()));
            }
        }
        if let Some(ref val) = self.cmd {
            start.push_attribute(("cmd", val.as_str()));
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
            let val = &self.c_bhvr;
            val.write_element("p:cBhvr", writer)?;
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

impl ToXml for CTTLSetBehavior {
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
            let val = &self.c_bhvr;
            val.write_element("p:cBhvr", writer)?;
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
        if let Some(ref val) = self.to {
            val.write_element("p:to", writer)?;
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

impl ToXml for CTTLCommonMediaNodeData {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        if let Some(ref val) = self.vol {
            {
                let s = val.to_string();
                start.push_attribute(("vol", s.as_str()));
            }
        }
        if let Some(ref val) = self.mute {
            start.push_attribute(("mute", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.num_sld {
            {
                let s = val.to_string();
                start.push_attribute(("numSld", s.as_str()));
            }
        }
        if let Some(ref val) = self.show_when_stopped {
            start.push_attribute(("showWhenStopped", if *val { "1" } else { "0" }));
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
            let val = &self.c_tn;
            val.write_element("p:cTn", writer)?;
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
            let val = &self.tgt_el;
            val.write_element("p:tgtEl", writer)?;
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

impl ToXml for CTTLMediaNodeAudio {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        if let Some(ref val) = self.is_narration {
            start.push_attribute(("isNarration", if *val { "1" } else { "0" }));
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
            let val = &self.c_media_node;
            val.write_element("p:cMediaNode", writer)?;
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

impl ToXml for CTTLMediaNodeVideo {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        if let Some(ref val) = self.full_scrn {
            start.push_attribute(("fullScrn", if *val { "1" } else { "0" }));
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
            let val = &self.c_media_node;
            val.write_element("p:cMediaNode", writer)?;
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

impl ToXml for PAGTLBuild {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        {
            let val = &self.spid;
            {
                let s = val.to_string();
                start.push_attribute(("spid", s.as_str()));
            }
        }
        {
            let val = &self.grp_id;
            {
                let s = val.to_string();
                start.push_attribute(("grpId", s.as_str()));
            }
        }
        if let Some(ref val) = self.ui_expand {
            start.push_attribute(("uiExpand", if *val { "1" } else { "0" }));
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

impl ToXml for CTTLTemplate {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        if let Some(ref val) = self.lvl {
            {
                let s = val.to_string();
                start.push_attribute(("lvl", s.as_str()));
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
            let val = &self.tn_lst;
            val.write_element("p:tnLst", writer)?;
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

impl ToXml for CTTLTemplateList {
    fn write_children<W: Write>(&self, writer: &mut Writer<W>) -> Result<(), SerializeError> {
        #[cfg(feature = "extra-children")]
        let mut extra_iter = self.extra_children.iter().peekable();
        #[cfg(feature = "extra-children")]
        let mut emit_idx: usize = 0;
        for item in &self.tmpl {
            #[cfg(feature = "extra-children")]
            while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
                extra_iter
                    .next()
                    .unwrap()
                    .node
                    .write_to(writer)
                    .map_err(SerializeError::from)?;
            }
            item.write_element("p:tmpl", writer)?;
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
        if !self.tmpl.is_empty() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for CTTLBuildParagraph {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        {
            let val = &self.spid;
            {
                let s = val.to_string();
                start.push_attribute(("spid", s.as_str()));
            }
        }
        {
            let val = &self.grp_id;
            {
                let s = val.to_string();
                start.push_attribute(("grpId", s.as_str()));
            }
        }
        if let Some(ref val) = self.ui_expand {
            start.push_attribute(("uiExpand", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.build {
            {
                let s = val.to_string();
                start.push_attribute(("build", s.as_str()));
            }
        }
        if let Some(ref val) = self.bld_lvl {
            {
                let s = val.to_string();
                start.push_attribute(("bldLvl", s.as_str()));
            }
        }
        if let Some(ref val) = self.anim_bg {
            start.push_attribute(("animBg", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.auto_update_anim_bg {
            start.push_attribute(("autoUpdateAnimBg", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.rev {
            start.push_attribute(("rev", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.adv_auto {
            {
                let s = val.to_string();
                start.push_attribute(("advAuto", s.as_str()));
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
        if let Some(ref val) = self.tmpl_lst {
            val.write_element("p:tmplLst", writer)?;
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
        if self.tmpl_lst.is_some() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for CTTLBuildDiagram {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        {
            let val = &self.spid;
            {
                let s = val.to_string();
                start.push_attribute(("spid", s.as_str()));
            }
        }
        {
            let val = &self.grp_id;
            {
                let s = val.to_string();
                start.push_attribute(("grpId", s.as_str()));
            }
        }
        if let Some(ref val) = self.ui_expand {
            start.push_attribute(("uiExpand", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.bld {
            {
                let s = val.to_string();
                start.push_attribute(("bld", s.as_str()));
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

impl ToXml for CTTLOleBuildChart {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        {
            let val = &self.spid;
            {
                let s = val.to_string();
                start.push_attribute(("spid", s.as_str()));
            }
        }
        {
            let val = &self.grp_id;
            {
                let s = val.to_string();
                start.push_attribute(("grpId", s.as_str()));
            }
        }
        if let Some(ref val) = self.ui_expand {
            start.push_attribute(("uiExpand", if *val { "1" } else { "0" }));
        }
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

impl ToXml for CTTLGraphicalObjectBuild {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        {
            let val = &self.spid;
            {
                let s = val.to_string();
                start.push_attribute(("spid", s.as_str()));
            }
        }
        {
            let val = &self.grp_id;
            {
                let s = val.to_string();
                start.push_attribute(("grpId", s.as_str()));
            }
        }
        if let Some(ref val) = self.ui_expand {
            start.push_attribute(("uiExpand", if *val { "1" } else { "0" }));
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
        if let Some(ref val) = self.bld_as_one {
            val.write_element("p:bldAsOne", writer)?;
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
        if let Some(ref val) = self.bld_sub {
            val.write_element("p:bldSub", writer)?;
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
        if self.bld_as_one.is_some() {
            return false;
        }
        if self.bld_sub.is_some() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for CTBuildList {
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
        if let Some(ref val) = self.bld_p {
            val.write_element("p:bldP", writer)?;
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
        if let Some(ref val) = self.bld_dgm {
            val.write_element("p:bldDgm", writer)?;
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
        if let Some(ref val) = self.bld_ole_chart {
            val.write_element("p:bldOleChart", writer)?;
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
        if let Some(ref val) = self.bld_graphic {
            val.write_element("p:bldGraphic", writer)?;
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
        if self.bld_p.is_some() {
            return false;
        }
        if self.bld_dgm.is_some() {
            return false;
        }
        if self.bld_ole_chart.is_some() {
            return false;
        }
        if self.bld_graphic.is_some() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for SlideTiming {
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
        if let Some(ref val) = self.tn_lst {
            val.write_element("p:tnLst", writer)?;
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
        if let Some(ref val) = self.bld_lst {
            val.write_element("p:bldLst", writer)?;
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
            val.write_element("p:extLst", writer)?;
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
        if self.tn_lst.is_some() {
            return false;
        }
        if self.bld_lst.is_some() {
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

impl ToXml for CTEmpty {
    fn is_empty_element(&self) -> bool {
        true
    }
}

impl ToXml for CTIndexRange {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        {
            let val = &self.st;
            {
                let s = val.to_string();
                start.push_attribute(("st", s.as_str()));
            }
        }
        {
            let val = &self.end;
            {
                let s = val.to_string();
                start.push_attribute(("end", s.as_str()));
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

impl ToXml for CTSlideRelationshipListEntry {
    fn write_children<W: Write>(&self, writer: &mut Writer<W>) -> Result<(), SerializeError> {
        #[cfg(feature = "extra-children")]
        for child in &self.extra_children {
            child.node.write_to(writer).map_err(SerializeError::from)?;
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

impl ToXml for CTSlideRelationshipList {
    fn write_children<W: Write>(&self, writer: &mut Writer<W>) -> Result<(), SerializeError> {
        #[cfg(feature = "extra-children")]
        let mut extra_iter = self.extra_children.iter().peekable();
        #[cfg(feature = "extra-children")]
        let mut emit_idx: usize = 0;
        for item in &self.sld {
            #[cfg(feature = "extra-children")]
            while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
                extra_iter
                    .next()
                    .unwrap()
                    .node
                    .write_to(writer)
                    .map_err(SerializeError::from)?;
            }
            item.write_element("p:sld", writer)?;
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
        if !self.sld.is_empty() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for CTCustomShowId {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        {
            let val = &self.id;
            {
                let s = val.to_string();
                start.push_attribute(("id", s.as_str()));
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

impl ToXml for EGSlideListChoice {
    fn write_element<W: Write>(
        &self,
        _tag: &str,
        writer: &mut Writer<W>,
    ) -> Result<(), SerializeError> {
        match self {
            Self::SldAll(inner) => inner.write_element("p:sldAll", writer)?,
            Self::SldRg(inner) => inner.write_element("p:sldRg", writer)?,
            Self::CustShow(inner) => inner.write_element("p:custShow", writer)?,
        }
        Ok(())
    }
}

impl ToXml for CTCustomerData {
    fn write_children<W: Write>(&self, writer: &mut Writer<W>) -> Result<(), SerializeError> {
        #[cfg(feature = "extra-children")]
        for child in &self.extra_children {
            child.node.write_to(writer).map_err(SerializeError::from)?;
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

impl ToXml for CTTagsData {
    fn write_children<W: Write>(&self, writer: &mut Writer<W>) -> Result<(), SerializeError> {
        #[cfg(feature = "extra-children")]
        for child in &self.extra_children {
            child.node.write_to(writer).map_err(SerializeError::from)?;
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

impl ToXml for CTCustomerDataList {
    fn write_children<W: Write>(&self, writer: &mut Writer<W>) -> Result<(), SerializeError> {
        #[cfg(feature = "extra-children")]
        let mut extra_iter = self.extra_children.iter().peekable();
        #[cfg(feature = "extra-children")]
        let mut emit_idx: usize = 0;
        for item in &self.cust_data {
            #[cfg(feature = "extra-children")]
            while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
                extra_iter
                    .next()
                    .unwrap()
                    .node
                    .write_to(writer)
                    .map_err(SerializeError::from)?;
            }
            item.write_element("p:custData", writer)?;
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
        if let Some(ref val) = self.tags {
            val.write_element("p:tags", writer)?;
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
        if !self.cust_data.is_empty() {
            return false;
        }
        if self.tags.is_some() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for CTExtension {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
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

impl ToXml for EGExtensionList {
    fn write_children<W: Write>(&self, writer: &mut Writer<W>) -> Result<(), SerializeError> {
        #[cfg(feature = "extra-children")]
        let mut extra_iter = self.extra_children.iter().peekable();
        #[cfg(feature = "extra-children")]
        let mut emit_idx: usize = 0;
        for item in &self.ext {
            #[cfg(feature = "extra-children")]
            while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
                extra_iter
                    .next()
                    .unwrap()
                    .node
                    .write_to(writer)
                    .map_err(SerializeError::from)?;
            }
            item.write_element("p:ext", writer)?;
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
        if !self.ext.is_empty() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for CTExtensionList {
    fn write_children<W: Write>(&self, writer: &mut Writer<W>) -> Result<(), SerializeError> {
        #[cfg(feature = "extra-children")]
        let mut extra_iter = self.extra_children.iter().peekable();
        #[cfg(feature = "extra-children")]
        let mut emit_idx: usize = 0;
        for item in &self.ext {
            #[cfg(feature = "extra-children")]
            while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
                extra_iter
                    .next()
                    .unwrap()
                    .node
                    .write_to(writer)
                    .map_err(SerializeError::from)?;
            }
            item.write_element("p:ext", writer)?;
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
        if !self.ext.is_empty() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for CTExtensionListModify {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        if let Some(ref val) = self.r#mod {
            start.push_attribute(("mod", if *val { "1" } else { "0" }));
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
        for item in &self.ext {
            #[cfg(feature = "extra-children")]
            while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
                extra_iter
                    .next()
                    .unwrap()
                    .node
                    .write_to(writer)
                    .map_err(SerializeError::from)?;
            }
            item.write_element("p:ext", writer)?;
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
        if !self.ext.is_empty() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for CTCommentAuthor {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
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
        {
            let val = &self.initials;
            start.push_attribute(("initials", val.as_str()));
        }
        {
            let val = &self.last_idx;
            {
                let s = val.to_string();
                start.push_attribute(("lastIdx", s.as_str()));
            }
        }
        {
            let val = &self.clr_idx;
            {
                let s = val.to_string();
                start.push_attribute(("clrIdx", s.as_str()));
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
            val.write_element("p:extLst", writer)?;
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

impl ToXml for CTCommentAuthorList {
    fn write_children<W: Write>(&self, writer: &mut Writer<W>) -> Result<(), SerializeError> {
        #[cfg(feature = "extra-children")]
        let mut extra_iter = self.extra_children.iter().peekable();
        #[cfg(feature = "extra-children")]
        let mut emit_idx: usize = 0;
        for item in &self.cm_author {
            #[cfg(feature = "extra-children")]
            while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
                extra_iter
                    .next()
                    .unwrap()
                    .node
                    .write_to(writer)
                    .map_err(SerializeError::from)?;
            }
            item.write_element("p:cmAuthor", writer)?;
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
        if !self.cm_author.is_empty() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for CTComment {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        {
            let val = &self.author_id;
            {
                let s = val.to_string();
                start.push_attribute(("authorId", s.as_str()));
            }
        }
        if let Some(ref val) = self.dt {
            start.push_attribute(("dt", val.as_str()));
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
            val.write_element("p:pos", writer)?;
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
            let val = &self.text;
            {
                let start = BytesStart::new("p:text");
                writer.write_event(Event::Start(start))?;
                writer.write_event(Event::Text(BytesText::new(val.as_str())))?;
                writer.write_event(Event::End(BytesEnd::new("p:text")))?;
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
        if let Some(ref val) = self.ext_lst {
            val.write_element("p:extLst", writer)?;
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

impl ToXml for CTCommentList {
    fn write_children<W: Write>(&self, writer: &mut Writer<W>) -> Result<(), SerializeError> {
        #[cfg(feature = "extra-children")]
        let mut extra_iter = self.extra_children.iter().peekable();
        #[cfg(feature = "extra-children")]
        let mut emit_idx: usize = 0;
        for item in &self.cm {
            #[cfg(feature = "extra-children")]
            while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
                extra_iter
                    .next()
                    .unwrap()
                    .node
                    .write_to(writer)
                    .map_err(SerializeError::from)?;
            }
            item.write_element("p:cm", writer)?;
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
        if !self.cm.is_empty() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for PAGOle {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        if let Some(ref val) = self.name {
            start.push_attribute(("name", val.as_str()));
        }
        if let Some(ref val) = self.show_as_icon {
            start.push_attribute(("showAsIcon", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.img_w {
            {
                let s = val.to_string();
                start.push_attribute(("imgW", s.as_str()));
            }
        }
        if let Some(ref val) = self.img_h {
            {
                let s = val.to_string();
                start.push_attribute(("imgH", s.as_str()));
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

impl ToXml for CTOleObjectEmbed {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        if let Some(ref val) = self.follow_color_scheme {
            {
                let s = val.to_string();
                start.push_attribute(("followColorScheme", s.as_str()));
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
            val.write_element("p:extLst", writer)?;
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

impl ToXml for CTOleObjectLink {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        if let Some(ref val) = self.update_automatic {
            start.push_attribute(("updateAutomatic", if *val { "1" } else { "0" }));
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
            val.write_element("p:extLst", writer)?;
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

impl ToXml for CTOleObject {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        if let Some(ref val) = self.name {
            start.push_attribute(("name", val.as_str()));
        }
        if let Some(ref val) = self.show_as_icon {
            start.push_attribute(("showAsIcon", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.img_w {
            {
                let s = val.to_string();
                start.push_attribute(("imgW", s.as_str()));
            }
        }
        if let Some(ref val) = self.img_h {
            {
                let s = val.to_string();
                start.push_attribute(("imgH", s.as_str()));
            }
        }
        if let Some(ref val) = self.prog_id {
            start.push_attribute(("progId", val.as_str()));
        }
        if let Some(ref val) = self.spid {
            {
                let s = val.to_string();
                start.push_attribute(("spid", s.as_str()));
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
        if let Some(ref val) = self.embed {
            val.write_element("p:embed", writer)?;
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
        if let Some(ref val) = self.link {
            val.write_element("p:link", writer)?;
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
        if let Some(ref val) = self.picture {
            val.write_element("p:pic", writer)?;
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
        if self.embed.is_some() {
            return false;
        }
        if self.link.is_some() {
            return false;
        }
        if self.picture.is_some() {
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
            start.push_attribute(("name", val.as_str()));
        }
        if let Some(ref val) = self.show_as_icon {
            start.push_attribute(("showAsIcon", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.img_w {
            {
                let s = val.to_string();
                start.push_attribute(("imgW", s.as_str()));
            }
        }
        if let Some(ref val) = self.img_h {
            {
                let s = val.to_string();
                start.push_attribute(("imgH", s.as_str()));
            }
        }
        if let Some(ref val) = self.spid {
            {
                let s = val.to_string();
                start.push_attribute(("spid", s.as_str()));
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
            val.write_element("p:extLst", writer)?;
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
        if let Some(ref val) = self.picture {
            val.write_element("p:pic", writer)?;
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
        if self.picture.is_some() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for CTControlList {
    fn write_children<W: Write>(&self, writer: &mut Writer<W>) -> Result<(), SerializeError> {
        #[cfg(feature = "extra-children")]
        let mut extra_iter = self.extra_children.iter().peekable();
        #[cfg(feature = "extra-children")]
        let mut emit_idx: usize = 0;
        for item in &self.control {
            #[cfg(feature = "extra-children")]
            while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
                extra_iter
                    .next()
                    .unwrap()
                    .node
                    .write_to(writer)
                    .map_err(SerializeError::from)?;
            }
            item.write_element("p:control", writer)?;
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
        if !self.control.is_empty() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for CTSlideIdListEntry {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        {
            let val = &self.id;
            {
                let s = val.to_string();
                start.push_attribute(("id", s.as_str()));
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
            val.write_element("p:extLst", writer)?;
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

impl ToXml for SlideIdList {
    fn write_children<W: Write>(&self, writer: &mut Writer<W>) -> Result<(), SerializeError> {
        #[cfg(feature = "extra-children")]
        let mut extra_iter = self.extra_children.iter().peekable();
        #[cfg(feature = "extra-children")]
        let mut emit_idx: usize = 0;
        for item in &self.sld_id {
            #[cfg(feature = "extra-children")]
            while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
                extra_iter
                    .next()
                    .unwrap()
                    .node
                    .write_to(writer)
                    .map_err(SerializeError::from)?;
            }
            item.write_element("p:sldId", writer)?;
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
        if !self.sld_id.is_empty() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for CTSlideMasterIdListEntry {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        if let Some(ref val) = self.id {
            {
                let s = val.to_string();
                start.push_attribute(("id", s.as_str()));
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
            val.write_element("p:extLst", writer)?;
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

impl ToXml for CTSlideMasterIdList {
    fn write_children<W: Write>(&self, writer: &mut Writer<W>) -> Result<(), SerializeError> {
        #[cfg(feature = "extra-children")]
        let mut extra_iter = self.extra_children.iter().peekable();
        #[cfg(feature = "extra-children")]
        let mut emit_idx: usize = 0;
        for item in &self.sld_master_id {
            #[cfg(feature = "extra-children")]
            while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
                extra_iter
                    .next()
                    .unwrap()
                    .node
                    .write_to(writer)
                    .map_err(SerializeError::from)?;
            }
            item.write_element("p:sldMasterId", writer)?;
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
        if !self.sld_master_id.is_empty() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for CTNotesMasterIdListEntry {
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
            val.write_element("p:extLst", writer)?;
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

impl ToXml for CTNotesMasterIdList {
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
        if let Some(ref val) = self.notes_master_id {
            val.write_element("p:notesMasterId", writer)?;
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
        if self.notes_master_id.is_some() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for CTHandoutMasterIdListEntry {
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
            val.write_element("p:extLst", writer)?;
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

impl ToXml for CTHandoutMasterIdList {
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
        if let Some(ref val) = self.handout_master_id {
            val.write_element("p:handoutMasterId", writer)?;
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
        if self.handout_master_id.is_some() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for CTEmbeddedFontDataId {
    fn write_children<W: Write>(&self, writer: &mut Writer<W>) -> Result<(), SerializeError> {
        #[cfg(feature = "extra-children")]
        for child in &self.extra_children {
            child.node.write_to(writer).map_err(SerializeError::from)?;
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

impl ToXml for CTEmbeddedFontListEntry {
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
            let val = &self.font;
            val.write_element("p:font", writer)?;
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
        if let Some(ref val) = self.regular {
            val.write_element("p:regular", writer)?;
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
        if let Some(ref val) = self.bold {
            val.write_element("p:bold", writer)?;
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
        if let Some(ref val) = self.italic {
            val.write_element("p:italic", writer)?;
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
        if let Some(ref val) = self.bold_italic {
            val.write_element("p:boldItalic", writer)?;
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

impl ToXml for CTEmbeddedFontList {
    fn write_children<W: Write>(&self, writer: &mut Writer<W>) -> Result<(), SerializeError> {
        #[cfg(feature = "extra-children")]
        let mut extra_iter = self.extra_children.iter().peekable();
        #[cfg(feature = "extra-children")]
        let mut emit_idx: usize = 0;
        for item in &self.embedded_font {
            #[cfg(feature = "extra-children")]
            while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
                extra_iter
                    .next()
                    .unwrap()
                    .node
                    .write_to(writer)
                    .map_err(SerializeError::from)?;
            }
            item.write_element("p:embeddedFont", writer)?;
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
        if !self.embedded_font.is_empty() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for CTSmartTags {
    fn write_children<W: Write>(&self, writer: &mut Writer<W>) -> Result<(), SerializeError> {
        #[cfg(feature = "extra-children")]
        for child in &self.extra_children {
            child.node.write_to(writer).map_err(SerializeError::from)?;
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

impl ToXml for CTCustomShow {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        {
            let val = &self.name;
            start.push_attribute(("name", val.as_str()));
        }
        {
            let val = &self.id;
            {
                let s = val.to_string();
                start.push_attribute(("id", s.as_str()));
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
            let val = &self.sld_lst;
            val.write_element("p:sldLst", writer)?;
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
            val.write_element("p:extLst", writer)?;
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

impl ToXml for CTCustomShowList {
    fn write_children<W: Write>(&self, writer: &mut Writer<W>) -> Result<(), SerializeError> {
        #[cfg(feature = "extra-children")]
        let mut extra_iter = self.extra_children.iter().peekable();
        #[cfg(feature = "extra-children")]
        let mut emit_idx: usize = 0;
        for item in &self.cust_show {
            #[cfg(feature = "extra-children")]
            while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
                extra_iter
                    .next()
                    .unwrap()
                    .node
                    .write_to(writer)
                    .map_err(SerializeError::from)?;
            }
            item.write_element("p:custShow", writer)?;
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
        if !self.cust_show.is_empty() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for CTPhotoAlbum {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        if let Some(ref val) = self.bw {
            start.push_attribute(("bw", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.show_captions {
            start.push_attribute(("showCaptions", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.layout {
            {
                let s = val.to_string();
                start.push_attribute(("layout", s.as_str()));
            }
        }
        if let Some(ref val) = self.frame {
            {
                let s = val.to_string();
                start.push_attribute(("frame", s.as_str()));
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
            val.write_element("p:extLst", writer)?;
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

impl ToXml for CTSlideSize {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
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
        if let Some(ref val) = self.r#type {
            {
                let s = val.to_string();
                start.push_attribute(("type", s.as_str()));
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
        if let Some(ref val) = self.lang {
            start.push_attribute(("lang", val.as_str()));
        }
        {
            let val = &self.inval_st_chars;
            start.push_attribute(("invalStChars", val.as_str()));
        }
        {
            let val = &self.inval_end_chars;
            start.push_attribute(("invalEndChars", val.as_str()));
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

impl ToXml for CTModifyVerifier {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        if let Some(ref val) = self.algorithm_name {
            start.push_attribute(("algorithmName", val.as_str()));
        }
        if let Some(ref val) = self.hash_value {
            {
                let hex = encode_hex(val);
                start.push_attribute(("hashValue", hex.as_str()));
            }
        }
        if let Some(ref val) = self.salt_value {
            {
                let hex = encode_hex(val);
                start.push_attribute(("saltValue", hex.as_str()));
            }
        }
        if let Some(ref val) = self.spin_value {
            {
                let s = val.to_string();
                start.push_attribute(("spinValue", s.as_str()));
            }
        }
        if let Some(ref val) = self.crypt_provider_type {
            {
                let s = val.to_string();
                start.push_attribute(("cryptProviderType", s.as_str()));
            }
        }
        if let Some(ref val) = self.crypt_algorithm_class {
            {
                let s = val.to_string();
                start.push_attribute(("cryptAlgorithmClass", s.as_str()));
            }
        }
        if let Some(ref val) = self.crypt_algorithm_type {
            {
                let s = val.to_string();
                start.push_attribute(("cryptAlgorithmType", s.as_str()));
            }
        }
        if let Some(ref val) = self.crypt_algorithm_sid {
            {
                let s = val.to_string();
                start.push_attribute(("cryptAlgorithmSid", s.as_str()));
            }
        }
        if let Some(ref val) = self.spin_count {
            {
                let s = val.to_string();
                start.push_attribute(("spinCount", s.as_str()));
            }
        }
        if let Some(ref val) = self.salt_data {
            {
                let hex = encode_hex(val);
                start.push_attribute(("saltData", hex.as_str()));
            }
        }
        if let Some(ref val) = self.hash_data {
            {
                let hex = encode_hex(val);
                start.push_attribute(("hashData", hex.as_str()));
            }
        }
        if let Some(ref val) = self.crypt_provider {
            start.push_attribute(("cryptProvider", val.as_str()));
        }
        if let Some(ref val) = self.alg_id_ext {
            {
                let s = val.to_string();
                start.push_attribute(("algIdExt", s.as_str()));
            }
        }
        if let Some(ref val) = self.alg_id_ext_source {
            start.push_attribute(("algIdExtSource", val.as_str()));
        }
        if let Some(ref val) = self.crypt_provider_type_ext {
            {
                let s = val.to_string();
                start.push_attribute(("cryptProviderTypeExt", s.as_str()));
            }
        }
        if let Some(ref val) = self.crypt_provider_type_ext_source {
            start.push_attribute(("cryptProviderTypeExtSource", val.as_str()));
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

impl ToXml for Presentation {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        if let Some(ref val) = self.server_zoom {
            {
                let s = val.to_string();
                start.push_attribute(("serverZoom", s.as_str()));
            }
        }
        if let Some(ref val) = self.first_slide_num {
            {
                let s = val.to_string();
                start.push_attribute(("firstSlideNum", s.as_str()));
            }
        }
        if let Some(ref val) = self.show_special_pls_on_title_sld {
            start.push_attribute(("showSpecialPlsOnTitleSld", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.rtl {
            start.push_attribute(("rtl", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.remove_personal_info_on_save {
            start.push_attribute(("removePersonalInfoOnSave", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.compat_mode {
            start.push_attribute(("compatMode", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.strict_first_and_last_chars {
            start.push_attribute(("strictFirstAndLastChars", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.embed_true_type_fonts {
            start.push_attribute(("embedTrueTypeFonts", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.save_subset_fonts {
            start.push_attribute(("saveSubsetFonts", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.auto_compress_pictures {
            start.push_attribute(("autoCompressPictures", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.bookmark_id_seed {
            {
                let s = val.to_string();
                start.push_attribute(("bookmarkIdSeed", s.as_str()));
            }
        }
        if let Some(ref val) = self.conformance {
            {
                let s = val.to_string();
                start.push_attribute(("conformance", s.as_str()));
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
        if let Some(ref val) = self.sld_master_id_lst {
            val.write_element("p:sldMasterIdLst", writer)?;
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
        if let Some(ref val) = self.notes_master_id_lst {
            val.write_element("p:notesMasterIdLst", writer)?;
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
        if let Some(ref val) = self.handout_master_id_lst {
            val.write_element("p:handoutMasterIdLst", writer)?;
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
        if let Some(ref val) = self.sld_id_lst {
            val.write_element("p:sldIdLst", writer)?;
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
        if let Some(ref val) = self.sld_sz {
            val.write_element("p:sldSz", writer)?;
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
            let val = &self.notes_sz;
            val.write_element("p:notesSz", writer)?;
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
        if let Some(ref val) = self.smart_tags {
            val.write_element("p:smartTags", writer)?;
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
        if let Some(ref val) = self.embedded_font_lst {
            val.write_element("p:embeddedFontLst", writer)?;
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
        if let Some(ref val) = self.cust_show_lst {
            val.write_element("p:custShowLst", writer)?;
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
        if let Some(ref val) = self.photo_album {
            val.write_element("p:photoAlbum", writer)?;
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
        if let Some(ref val) = self.cust_data_lst {
            val.write_element("p:custDataLst", writer)?;
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
        if let Some(ref val) = self.kinsoku {
            val.write_element("p:kinsoku", writer)?;
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
        if let Some(ref val) = self.default_text_style {
            val.write_element("p:defaultTextStyle", writer)?;
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
        if let Some(ref val) = self.modify_verifier {
            val.write_element("p:modifyVerifier", writer)?;
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
            val.write_element("p:extLst", writer)?;
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
        if self.sld_master_id_lst.is_some() {
            return false;
        }
        if self.notes_master_id_lst.is_some() {
            return false;
        }
        if self.handout_master_id_lst.is_some() {
            return false;
        }
        if self.sld_id_lst.is_some() {
            return false;
        }
        if self.sld_sz.is_some() {
            return false;
        }
        false
    }
}

impl ToXml for CTHtmlPublishProperties {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        if let Some(ref val) = self.show_speaker_notes {
            start.push_attribute(("showSpeakerNotes", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.target {
            start.push_attribute(("target", val.as_str()));
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
        if let Some(ref val) = self.slide_list_choice {
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
            val.write_element("p:extLst", writer)?;
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

impl ToXml for CTWebProperties {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        if let Some(ref val) = self.show_animation {
            start.push_attribute(("showAnimation", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.resize_graphics {
            start.push_attribute(("resizeGraphics", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.allow_png {
            start.push_attribute(("allowPng", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.rely_on_vml {
            start.push_attribute(("relyOnVml", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.organize_in_folders {
            start.push_attribute(("organizeInFolders", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.use_long_filenames {
            start.push_attribute(("useLongFilenames", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.img_sz {
            {
                let s = val.to_string();
                start.push_attribute(("imgSz", s.as_str()));
            }
        }
        if let Some(ref val) = self.encoding {
            start.push_attribute(("encoding", val.as_str()));
        }
        if let Some(ref val) = self.clr {
            {
                let s = val.to_string();
                start.push_attribute(("clr", s.as_str()));
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
            val.write_element("p:extLst", writer)?;
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

impl ToXml for CTPrintProperties {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        if let Some(ref val) = self.prn_what {
            {
                let s = val.to_string();
                start.push_attribute(("prnWhat", s.as_str()));
            }
        }
        if let Some(ref val) = self.clr_mode {
            {
                let s = val.to_string();
                start.push_attribute(("clrMode", s.as_str()));
            }
        }
        if let Some(ref val) = self.hidden_slides {
            start.push_attribute(("hiddenSlides", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.scale_to_fit_paper {
            start.push_attribute(("scaleToFitPaper", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.frame_slides {
            start.push_attribute(("frameSlides", if *val { "1" } else { "0" }));
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
            val.write_element("p:extLst", writer)?;
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

impl ToXml for CTShowInfoBrowse {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        if let Some(ref val) = self.show_scrollbar {
            start.push_attribute(("showScrollbar", if *val { "1" } else { "0" }));
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

impl ToXml for CTShowInfoKiosk {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        if let Some(ref val) = self.restart {
            {
                let s = val.to_string();
                start.push_attribute(("restart", s.as_str()));
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

impl ToXml for EGShowType {
    fn write_element<W: Write>(
        &self,
        _tag: &str,
        writer: &mut Writer<W>,
    ) -> Result<(), SerializeError> {
        match self {
            Self::Present(inner) => inner.write_element("p:present", writer)?,
            Self::Browse(inner) => inner.write_element("p:browse", writer)?,
            Self::Kiosk(inner) => inner.write_element("p:kiosk", writer)?,
        }
        Ok(())
    }
}

impl ToXml for CTShowProperties {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        if let Some(ref val) = self.r#loop {
            start.push_attribute(("loop", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.show_narration {
            start.push_attribute(("showNarration", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.show_animation {
            start.push_attribute(("showAnimation", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.use_timings {
            start.push_attribute(("useTimings", if *val { "1" } else { "0" }));
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
        if let Some(ref val) = self.show_type {
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
        if let Some(ref val) = self.slide_list_choice {
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
        if let Some(ref val) = self.pen_clr {
            val.write_element("p:penClr", writer)?;
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
            val.write_element("p:extLst", writer)?;
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
        if self.show_type.is_some() {
            return false;
        }
        if self.slide_list_choice.is_some() {
            return false;
        }
        if self.pen_clr.is_some() {
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

impl ToXml for CTPresentationProperties {
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
        if let Some(ref val) = self.html_pub_pr {
            val.write_element("p:htmlPubPr", writer)?;
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
        if let Some(ref val) = self.web_pr {
            val.write_element("p:webPr", writer)?;
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
        if let Some(ref val) = self.prn_pr {
            val.write_element("p:prnPr", writer)?;
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
        if let Some(ref val) = self.show_pr {
            val.write_element("p:showPr", writer)?;
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
        if let Some(ref val) = self.clr_mru {
            val.write_element("p:clrMru", writer)?;
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
            val.write_element("p:extLst", writer)?;
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
        if self.html_pub_pr.is_some() {
            return false;
        }
        if self.web_pr.is_some() {
            return false;
        }
        if self.prn_pr.is_some() {
            return false;
        }
        if self.show_pr.is_some() {
            return false;
        }
        if self.clr_mru.is_some() {
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

impl ToXml for CTHeaderFooter {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        if let Some(ref val) = self.sld_num {
            start.push_attribute(("sldNum", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.hdr {
            start.push_attribute(("hdr", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.ftr {
            start.push_attribute(("ftr", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.dt {
            start.push_attribute(("dt", if *val { "1" } else { "0" }));
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
            val.write_element("p:extLst", writer)?;
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

impl ToXml for CTPlaceholder {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        if let Some(ref val) = self.r#type {
            {
                let s = val.to_string();
                start.push_attribute(("type", s.as_str()));
            }
        }
        if let Some(ref val) = self.orient {
            {
                let s = val.to_string();
                start.push_attribute(("orient", s.as_str()));
            }
        }
        if let Some(ref val) = self.sz {
            {
                let s = val.to_string();
                start.push_attribute(("sz", s.as_str()));
            }
        }
        if let Some(ref val) = self.idx {
            {
                let s = val.to_string();
                start.push_attribute(("idx", s.as_str()));
            }
        }
        if let Some(ref val) = self.has_custom_prompt {
            start.push_attribute(("hasCustomPrompt", if *val { "1" } else { "0" }));
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
            val.write_element("p:extLst", writer)?;
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

impl ToXml for CTApplicationNonVisualDrawingProps {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        if let Some(ref val) = self.is_photo {
            start.push_attribute(("isPhoto", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.user_drawn {
            start.push_attribute(("userDrawn", if *val { "1" } else { "0" }));
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
        if let Some(ref val) = self.ph {
            val.write_element("p:ph", writer)?;
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
        if let Some(ref val) = self.cust_data_lst {
            val.write_element("p:custDataLst", writer)?;
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
            val.write_element("p:extLst", writer)?;
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
        if self.ph.is_some() {
            return false;
        }
        if self.cust_data_lst.is_some() {
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

impl ToXml for ShapeNonVisual {
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
            let val = &self.c_nv_pr;
            val.write_element("p:cNvPr", writer)?;
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
            let val = &self.c_nv_sp_pr;
            val.write_element("p:cNvSpPr", writer)?;
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
            let val = &self.nv_pr;
            val.write_element("p:nvPr", writer)?;
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

impl ToXml for Shape {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        if let Some(ref val) = self.use_bg_fill {
            start.push_attribute(("useBgFill", if *val { "1" } else { "0" }));
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
            let val = &self.non_visual_properties;
            val.write_element("p:nvSpPr", writer)?;
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
            let val = &self.shape_properties;
            val.write_element("p:spPr", writer)?;
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
            val.write_element("p:style", writer)?;
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
        if let Some(ref val) = self.text_body {
            val.write_element("p:txBody", writer)?;
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
            val.write_element("p:extLst", writer)?;
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

impl ToXml for CTConnectorNonVisual {
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
            let val = &self.c_nv_pr;
            val.write_element("p:cNvPr", writer)?;
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
            val.write_element("p:cNvCxnSpPr", writer)?;
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
            let val = &self.nv_pr;
            val.write_element("p:nvPr", writer)?;
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

impl ToXml for Connector {
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
            let val = &self.non_visual_connector_properties;
            val.write_element("p:nvCxnSpPr", writer)?;
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
            let val = &self.shape_properties;
            val.write_element("p:spPr", writer)?;
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
            val.write_element("p:style", writer)?;
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
            val.write_element("p:extLst", writer)?;
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

impl ToXml for CTPictureNonVisual {
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
            let val = &self.c_nv_pr;
            val.write_element("p:cNvPr", writer)?;
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
            let val = &self.c_nv_pic_pr;
            val.write_element("p:cNvPicPr", writer)?;
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
            let val = &self.nv_pr;
            val.write_element("p:nvPr", writer)?;
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

impl ToXml for Picture {
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
            let val = &self.non_visual_picture_properties;
            val.write_element("p:nvPicPr", writer)?;
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
            val.write_element("p:blipFill", writer)?;
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
            let val = &self.shape_properties;
            val.write_element("p:spPr", writer)?;
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
            val.write_element("p:style", writer)?;
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
            val.write_element("p:extLst", writer)?;
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

impl ToXml for CTGraphicalObjectFrameNonVisual {
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
            let val = &self.c_nv_pr;
            val.write_element("p:cNvPr", writer)?;
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
            val.write_element("p:cNvGraphicFramePr", writer)?;
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
            let val = &self.nv_pr;
            val.write_element("p:nvPr", writer)?;
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

impl ToXml for GraphicalObjectFrame {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
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
        {
            let val = &self.nv_graphic_frame_pr;
            val.write_element("p:nvGraphicFramePr", writer)?;
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
            let val = &self.xfrm;
            val.write_element("p:xfrm", writer)?;
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
            val.write_element("p:extLst", writer)?;
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

impl ToXml for CTGroupShapeNonVisual {
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
            let val = &self.c_nv_pr;
            val.write_element("p:cNvPr", writer)?;
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
            val.write_element("p:cNvGrpSpPr", writer)?;
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
            let val = &self.nv_pr;
            val.write_element("p:nvPr", writer)?;
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

impl ToXml for GroupShape {
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
            let val = &self.non_visual_group_properties;
            val.write_element("p:nvGrpSpPr", writer)?;
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
            val.write_element("p:grpSpPr", writer)?;
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
        if let Some(ref val) = self.shape {
            val.write_element("p:sp", writer)?;
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
        if let Some(ref val) = self.group_shape {
            val.write_element("p:grpSp", writer)?;
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
            val.write_element("p:graphicFrame", writer)?;
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
        if let Some(ref val) = self.connector {
            val.write_element("p:cxnSp", writer)?;
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
        if let Some(ref val) = self.picture {
            val.write_element("p:pic", writer)?;
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
        if let Some(ref val) = self.content_part {
            val.write_element("p:contentPart", writer)?;
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
            val.write_element("p:extLst", writer)?;
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

impl ToXml for CTRel {
    fn write_children<W: Write>(&self, writer: &mut Writer<W>) -> Result<(), SerializeError> {
        #[cfg(feature = "extra-children")]
        for child in &self.extra_children {
            child.node.write_to(writer).map_err(SerializeError::from)?;
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

impl ToXml for EGChildSlide {
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
        if let Some(ref val) = self.clr_map_ovr {
            val.write_element("p:clrMapOvr", writer)?;
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
        if self.clr_map_ovr.is_some() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for PAGChildSlide {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        if let Some(ref val) = self.show_master_sp {
            start.push_attribute(("showMasterSp", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.show_master_ph_anim {
            start.push_attribute(("showMasterPhAnim", if *val { "1" } else { "0" }));
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

impl ToXml for CTBackgroundProperties {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        if let Some(ref val) = self.shade_to_title {
            start.push_attribute(("shadeToTitle", if *val { "1" } else { "0" }));
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
            val.write_element("p:extLst", writer)?;
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

impl ToXml for EGBackground {
    fn write_element<W: Write>(
        &self,
        _tag: &str,
        writer: &mut Writer<W>,
    ) -> Result<(), SerializeError> {
        match self {
            Self::BgPr(inner) => inner.write_element("p:bgPr", writer)?,
            Self::BgRef(inner) => inner.write_element("p:bgRef", writer)?,
        }
        Ok(())
    }
}

impl ToXml for CTBackground {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
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
        if let Some(ref val) = self.background {
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

impl ToXml for CommonSlideData {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
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
        if let Some(ref val) = self.bg {
            val.write_element("p:bg", writer)?;
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
            let val = &self.shape_tree;
            val.write_element("p:spTree", writer)?;
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
        if let Some(ref val) = self.cust_data_lst {
            val.write_element("p:custDataLst", writer)?;
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
        if let Some(ref val) = self.controls {
            val.write_element("p:controls", writer)?;
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
            val.write_element("p:extLst", writer)?;
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
        if self.bg.is_some() {
            return false;
        }
        false
    }
}

impl ToXml for Slide {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        if let Some(ref val) = self.show_master_sp {
            start.push_attribute(("showMasterSp", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.show_master_ph_anim {
            start.push_attribute(("showMasterPhAnim", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.show {
            start.push_attribute(("show", if *val { "1" } else { "0" }));
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
            let val = &self.common_slide_data;
            val.write_element("p:cSld", writer)?;
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
        if let Some(ref val) = self.clr_map_ovr {
            val.write_element("p:clrMapOvr", writer)?;
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
        #[cfg(feature = "pml-transitions")]
        if let Some(ref val) = self.transition {
            val.write_element("p:transition", writer)?;
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
        #[cfg(feature = "pml-animations")]
        if let Some(ref val) = self.timing {
            val.write_element("p:timing", writer)?;
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
            val.write_element("p:extLst", writer)?;
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

impl ToXml for SlideLayout {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        if let Some(ref val) = self.show_master_sp {
            start.push_attribute(("showMasterSp", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.show_master_ph_anim {
            start.push_attribute(("showMasterPhAnim", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.matching_name {
            start.push_attribute(("matchingName", val.as_str()));
        }
        if let Some(ref val) = self.r#type {
            {
                let s = val.to_string();
                start.push_attribute(("type", s.as_str()));
            }
        }
        if let Some(ref val) = self.preserve {
            start.push_attribute(("preserve", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.user_drawn {
            start.push_attribute(("userDrawn", if *val { "1" } else { "0" }));
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
            let val = &self.common_slide_data;
            val.write_element("p:cSld", writer)?;
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
        if let Some(ref val) = self.clr_map_ovr {
            val.write_element("p:clrMapOvr", writer)?;
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
        if let Some(ref val) = self.transition {
            val.write_element("p:transition", writer)?;
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
        if let Some(ref val) = self.timing {
            val.write_element("p:timing", writer)?;
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
        if let Some(ref val) = self.hf {
            val.write_element("p:hf", writer)?;
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
            val.write_element("p:extLst", writer)?;
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

impl ToXml for CTSlideMasterTextStyles {
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
        if let Some(ref val) = self.title_style {
            val.write_element("p:titleStyle", writer)?;
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
        if let Some(ref val) = self.body_style {
            val.write_element("p:bodyStyle", writer)?;
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
        if let Some(ref val) = self.other_style {
            val.write_element("p:otherStyle", writer)?;
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
            val.write_element("p:extLst", writer)?;
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
        if self.title_style.is_some() {
            return false;
        }
        if self.body_style.is_some() {
            return false;
        }
        if self.other_style.is_some() {
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

impl ToXml for CTSlideLayoutIdListEntry {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        if let Some(ref val) = self.id {
            {
                let s = val.to_string();
                start.push_attribute(("id", s.as_str()));
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
            val.write_element("p:extLst", writer)?;
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

impl ToXml for CTSlideLayoutIdList {
    fn write_children<W: Write>(&self, writer: &mut Writer<W>) -> Result<(), SerializeError> {
        #[cfg(feature = "extra-children")]
        let mut extra_iter = self.extra_children.iter().peekable();
        #[cfg(feature = "extra-children")]
        let mut emit_idx: usize = 0;
        for item in &self.sld_layout_id {
            #[cfg(feature = "extra-children")]
            while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
                extra_iter
                    .next()
                    .unwrap()
                    .node
                    .write_to(writer)
                    .map_err(SerializeError::from)?;
            }
            item.write_element("p:sldLayoutId", writer)?;
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
        if !self.sld_layout_id.is_empty() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for SlideMaster {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        if let Some(ref val) = self.preserve {
            start.push_attribute(("preserve", if *val { "1" } else { "0" }));
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
            let val = &self.common_slide_data;
            val.write_element("p:cSld", writer)?;
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
            val.write_element("p:clrMap", writer)?;
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
        if let Some(ref val) = self.sld_layout_id_lst {
            val.write_element("p:sldLayoutIdLst", writer)?;
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
        if let Some(ref val) = self.transition {
            val.write_element("p:transition", writer)?;
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
        if let Some(ref val) = self.timing {
            val.write_element("p:timing", writer)?;
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
        if let Some(ref val) = self.hf {
            val.write_element("p:hf", writer)?;
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
        if let Some(ref val) = self.tx_styles {
            val.write_element("p:txStyles", writer)?;
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
            val.write_element("p:extLst", writer)?;
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

impl ToXml for HandoutMaster {
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
            let val = &self.common_slide_data;
            val.write_element("p:cSld", writer)?;
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
            val.write_element("p:clrMap", writer)?;
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
        if let Some(ref val) = self.hf {
            val.write_element("p:hf", writer)?;
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
            val.write_element("p:extLst", writer)?;
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

impl ToXml for NotesMaster {
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
            let val = &self.common_slide_data;
            val.write_element("p:cSld", writer)?;
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
            val.write_element("p:clrMap", writer)?;
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
        if let Some(ref val) = self.hf {
            val.write_element("p:hf", writer)?;
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
        if let Some(ref val) = self.notes_style {
            val.write_element("p:notesStyle", writer)?;
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
            val.write_element("p:extLst", writer)?;
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

impl ToXml for NotesSlide {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        if let Some(ref val) = self.show_master_sp {
            start.push_attribute(("showMasterSp", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.show_master_ph_anim {
            start.push_attribute(("showMasterPhAnim", if *val { "1" } else { "0" }));
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
            let val = &self.common_slide_data;
            val.write_element("p:cSld", writer)?;
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
        if let Some(ref val) = self.clr_map_ovr {
            val.write_element("p:clrMapOvr", writer)?;
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
            val.write_element("p:extLst", writer)?;
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

impl ToXml for CTSlideSyncProperties {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        {
            let val = &self.server_sld_id;
            start.push_attribute(("serverSldId", val.as_str()));
        }
        {
            let val = &self.server_sld_modified_time;
            start.push_attribute(("serverSldModifiedTime", val.as_str()));
        }
        {
            let val = &self.client_inserted_time;
            start.push_attribute(("clientInsertedTime", val.as_str()));
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
            val.write_element("p:extLst", writer)?;
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

impl ToXml for CTStringTag {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        {
            let val = &self.name;
            start.push_attribute(("name", val.as_str()));
        }
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

impl ToXml for CTTagList {
    fn write_children<W: Write>(&self, writer: &mut Writer<W>) -> Result<(), SerializeError> {
        #[cfg(feature = "extra-children")]
        let mut extra_iter = self.extra_children.iter().peekable();
        #[cfg(feature = "extra-children")]
        let mut emit_idx: usize = 0;
        for item in &self.tag {
            #[cfg(feature = "extra-children")]
            while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
                extra_iter
                    .next()
                    .unwrap()
                    .node
                    .write_to(writer)
                    .map_err(SerializeError::from)?;
            }
            item.write_element("p:tag", writer)?;
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
        if !self.tag.is_empty() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for CTNormalViewPortion {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        {
            let val = &self.sz;
            {
                let s = val.to_string();
                start.push_attribute(("sz", s.as_str()));
            }
        }
        if let Some(ref val) = self.auto_adjust {
            start.push_attribute(("autoAdjust", if *val { "1" } else { "0" }));
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

impl ToXml for CTNormalViewProperties {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        if let Some(ref val) = self.show_outline_icons {
            start.push_attribute(("showOutlineIcons", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.snap_vert_splitter {
            start.push_attribute(("snapVertSplitter", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.vert_bar_state {
            {
                let s = val.to_string();
                start.push_attribute(("vertBarState", s.as_str()));
            }
        }
        if let Some(ref val) = self.horz_bar_state {
            {
                let s = val.to_string();
                start.push_attribute(("horzBarState", s.as_str()));
            }
        }
        if let Some(ref val) = self.prefer_single_view {
            start.push_attribute(("preferSingleView", if *val { "1" } else { "0" }));
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
            let val = &self.restored_left;
            val.write_element("p:restoredLeft", writer)?;
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
            let val = &self.restored_top;
            val.write_element("p:restoredTop", writer)?;
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
            val.write_element("p:extLst", writer)?;
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

impl ToXml for CTCommonViewProperties {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        if let Some(ref val) = self.var_scale {
            start.push_attribute(("varScale", if *val { "1" } else { "0" }));
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
            let val = &self.scale;
            val.write_element("p:scale", writer)?;
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
            let val = &self.origin;
            val.write_element("p:origin", writer)?;
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

impl ToXml for CTNotesTextViewProperties {
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
            let val = &self.c_view_pr;
            val.write_element("p:cViewPr", writer)?;
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
            val.write_element("p:extLst", writer)?;
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

impl ToXml for CTOutlineViewSlideEntry {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        if let Some(ref val) = self.collapse {
            start.push_attribute(("collapse", if *val { "1" } else { "0" }));
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

impl ToXml for CTOutlineViewSlideList {
    fn write_children<W: Write>(&self, writer: &mut Writer<W>) -> Result<(), SerializeError> {
        #[cfg(feature = "extra-children")]
        let mut extra_iter = self.extra_children.iter().peekable();
        #[cfg(feature = "extra-children")]
        let mut emit_idx: usize = 0;
        for item in &self.sld {
            #[cfg(feature = "extra-children")]
            while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
                extra_iter
                    .next()
                    .unwrap()
                    .node
                    .write_to(writer)
                    .map_err(SerializeError::from)?;
            }
            item.write_element("p:sld", writer)?;
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
        if !self.sld.is_empty() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for CTOutlineViewProperties {
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
            let val = &self.c_view_pr;
            val.write_element("p:cViewPr", writer)?;
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
        if let Some(ref val) = self.sld_lst {
            val.write_element("p:sldLst", writer)?;
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
            val.write_element("p:extLst", writer)?;
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

impl ToXml for CTSlideSorterViewProperties {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        if let Some(ref val) = self.show_formatting {
            start.push_attribute(("showFormatting", if *val { "1" } else { "0" }));
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
            let val = &self.c_view_pr;
            val.write_element("p:cViewPr", writer)?;
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
            val.write_element("p:extLst", writer)?;
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

impl ToXml for CTGuide {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        if let Some(ref val) = self.orient {
            {
                let s = val.to_string();
                start.push_attribute(("orient", s.as_str()));
            }
        }
        if let Some(ref val) = self.pos {
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

    fn is_empty_element(&self) -> bool {
        true
    }
}

impl ToXml for CTGuideList {
    fn write_children<W: Write>(&self, writer: &mut Writer<W>) -> Result<(), SerializeError> {
        #[cfg(feature = "extra-children")]
        let mut extra_iter = self.extra_children.iter().peekable();
        #[cfg(feature = "extra-children")]
        let mut emit_idx: usize = 0;
        for item in &self.guide {
            #[cfg(feature = "extra-children")]
            while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
                extra_iter
                    .next()
                    .unwrap()
                    .node
                    .write_to(writer)
                    .map_err(SerializeError::from)?;
            }
            item.write_element("p:guide", writer)?;
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
        if !self.guide.is_empty() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for CTCommonSlideViewProperties {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        if let Some(ref val) = self.snap_to_grid {
            start.push_attribute(("snapToGrid", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.snap_to_objects {
            start.push_attribute(("snapToObjects", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.show_guides {
            start.push_attribute(("showGuides", if *val { "1" } else { "0" }));
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
            let val = &self.c_view_pr;
            val.write_element("p:cViewPr", writer)?;
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
        if let Some(ref val) = self.guide_lst {
            val.write_element("p:guideLst", writer)?;
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

impl ToXml for CTSlideViewProperties {
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
            let val = &self.c_sld_view_pr;
            val.write_element("p:cSldViewPr", writer)?;
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
            val.write_element("p:extLst", writer)?;
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

impl ToXml for CTNotesViewProperties {
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
            let val = &self.c_sld_view_pr;
            val.write_element("p:cSldViewPr", writer)?;
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
            val.write_element("p:extLst", writer)?;
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

impl ToXml for CTViewProperties {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        if let Some(ref val) = self.last_view {
            {
                let s = val.to_string();
                start.push_attribute(("lastView", s.as_str()));
            }
        }
        if let Some(ref val) = self.show_comments {
            start.push_attribute(("showComments", if *val { "1" } else { "0" }));
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
        if let Some(ref val) = self.normal_view_pr {
            val.write_element("p:normalViewPr", writer)?;
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
        if let Some(ref val) = self.slide_view_pr {
            val.write_element("p:slideViewPr", writer)?;
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
        if let Some(ref val) = self.outline_view_pr {
            val.write_element("p:outlineViewPr", writer)?;
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
        if let Some(ref val) = self.notes_text_view_pr {
            val.write_element("p:notesTextViewPr", writer)?;
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
        if let Some(ref val) = self.sorter_view_pr {
            val.write_element("p:sorterViewPr", writer)?;
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
        if let Some(ref val) = self.notes_view_pr {
            val.write_element("p:notesViewPr", writer)?;
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
        if let Some(ref val) = self.grid_spacing {
            val.write_element("p:gridSpacing", writer)?;
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
            val.write_element("p:extLst", writer)?;
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
        if self.normal_view_pr.is_some() {
            return false;
        }
        if self.slide_view_pr.is_some() {
            return false;
        }
        if self.outline_view_pr.is_some() {
            return false;
        }
        if self.notes_text_view_pr.is_some() {
            return false;
        }
        if self.sorter_view_pr.is_some() {
            return false;
        }
        if self.notes_view_pr.is_some() {
            return false;
        }
        if self.grid_spacing.is_some() {
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
