// Event-based parsers for generated types.
// ~3x faster than serde-based deserialization.

#![allow(unused_variables)]
#![allow(unused_imports)]
#![allow(clippy::single_match)]
#![allow(clippy::match_single_binding)]
#![allow(clippy::manual_is_multiple_of)]

use super::generated::*;
use ooxml_dml::types::*;
pub use ooxml_xml::{FromXml, ParseError};
#[cfg(feature = "extra-children")]
use ooxml_xml::{PositionedNode, RawXmlElement, RawXmlNode};
use quick_xml::Reader;
use quick_xml::events::{BytesStart, Event};
use std::io::BufRead;

#[allow(dead_code)]
/// Skip an element and all its children.
fn skip_element<R: BufRead>(reader: &mut Reader<R>) -> Result<(), ParseError> {
    let mut depth = 1u32;
    let mut buf = Vec::new();
    loop {
        match reader.read_event_into(&mut buf)? {
            Event::Start(_) => depth += 1,
            Event::End(_) => {
                depth -= 1;
                if depth == 0 {
                    break;
                }
            }
            Event::Eof => break,
            _ => {}
        }
        buf.clear();
    }
    Ok(())
}

#[allow(dead_code)]
/// Read the text content of an element until its end tag.
fn read_text_content<R: BufRead>(reader: &mut Reader<R>) -> Result<String, ParseError> {
    let mut text = String::new();
    let mut buf = Vec::new();
    loop {
        match reader.read_event_into(&mut buf)? {
            Event::Text(e) => text.push_str(&e.decode().unwrap_or_default()),
            Event::CData(e) => text.push_str(&e.decode().unwrap_or_default()),
            Event::End(_) => break,
            Event::Eof => break,
            _ => {}
        }
        buf.clear();
    }
    Ok(text)
}

#[allow(dead_code)]
/// Decode a hex string to bytes.
fn decode_hex(s: &str) -> Option<Vec<u8>> {
    let s = s.trim();
    if s.len() % 2 != 0 {
        return None;
    }
    (0..s.len())
        .step_by(2)
        .map(|i| u8::from_str_radix(&s[i..i + 2], 16).ok())
        .collect()
}

impl FromXml for CTSideDirectionTransition {
    fn from_xml<R: BufRead>(
        reader: &mut Reader<R>,
        start_tag: &BytesStart,
        is_empty: bool,
    ) -> Result<Self, ParseError> {
        let mut f_dir = None;
        #[cfg(feature = "extra-attrs")]
        let mut extra_attrs = std::collections::HashMap::new();

        // Parse attributes
        for attr in start_tag.attributes().filter_map(|a| a.ok()) {
            let val = String::from_utf8_lossy(&attr.value);
            match attr.key.local_name().as_ref() {
                b"dir" => {
                    f_dir = val.parse().ok();
                }
                #[cfg(feature = "extra-attrs")]
                unknown => {
                    let key = String::from_utf8_lossy(attr.key.as_ref()).into_owned();
                    extra_attrs.insert(key, val.into_owned());
                }
                #[cfg(not(feature = "extra-attrs"))]
                _ => {}
            }
        }

        if !is_empty {
            let mut buf = Vec::new();
            loop {
                match reader.read_event_into(&mut buf)? {
                    Event::End(_) => break,
                    Event::Eof => break,
                    _ => {}
                }
                buf.clear();
            }
        }

        Ok(Self {
            dir: f_dir,
            #[cfg(feature = "extra-attrs")]
            extra_attrs,
        })
    }
}

impl FromXml for CTCornerDirectionTransition {
    fn from_xml<R: BufRead>(
        reader: &mut Reader<R>,
        start_tag: &BytesStart,
        is_empty: bool,
    ) -> Result<Self, ParseError> {
        let mut f_dir = None;
        #[cfg(feature = "extra-attrs")]
        let mut extra_attrs = std::collections::HashMap::new();

        // Parse attributes
        for attr in start_tag.attributes().filter_map(|a| a.ok()) {
            let val = String::from_utf8_lossy(&attr.value);
            match attr.key.local_name().as_ref() {
                b"dir" => {
                    f_dir = val.parse().ok();
                }
                #[cfg(feature = "extra-attrs")]
                unknown => {
                    let key = String::from_utf8_lossy(attr.key.as_ref()).into_owned();
                    extra_attrs.insert(key, val.into_owned());
                }
                #[cfg(not(feature = "extra-attrs"))]
                _ => {}
            }
        }

        if !is_empty {
            let mut buf = Vec::new();
            loop {
                match reader.read_event_into(&mut buf)? {
                    Event::End(_) => break,
                    Event::Eof => break,
                    _ => {}
                }
                buf.clear();
            }
        }

        Ok(Self {
            dir: f_dir,
            #[cfg(feature = "extra-attrs")]
            extra_attrs,
        })
    }
}

impl FromXml for CTEightDirectionTransition {
    fn from_xml<R: BufRead>(
        reader: &mut Reader<R>,
        start_tag: &BytesStart,
        is_empty: bool,
    ) -> Result<Self, ParseError> {
        let mut f_dir = None;
        #[cfg(feature = "extra-attrs")]
        let mut extra_attrs = std::collections::HashMap::new();

        // Parse attributes
        for attr in start_tag.attributes().filter_map(|a| a.ok()) {
            let val = String::from_utf8_lossy(&attr.value);
            match attr.key.local_name().as_ref() {
                b"dir" => {
                    f_dir = Some(val.into_owned());
                }
                #[cfg(feature = "extra-attrs")]
                unknown => {
                    let key = String::from_utf8_lossy(attr.key.as_ref()).into_owned();
                    extra_attrs.insert(key, val.into_owned());
                }
                #[cfg(not(feature = "extra-attrs"))]
                _ => {}
            }
        }

        if !is_empty {
            let mut buf = Vec::new();
            loop {
                match reader.read_event_into(&mut buf)? {
                    Event::End(_) => break,
                    Event::Eof => break,
                    _ => {}
                }
                buf.clear();
            }
        }

        Ok(Self {
            dir: f_dir,
            #[cfg(feature = "extra-attrs")]
            extra_attrs,
        })
    }
}

impl FromXml for CTOrientationTransition {
    fn from_xml<R: BufRead>(
        reader: &mut Reader<R>,
        start_tag: &BytesStart,
        is_empty: bool,
    ) -> Result<Self, ParseError> {
        let mut f_dir = None;
        #[cfg(feature = "extra-attrs")]
        let mut extra_attrs = std::collections::HashMap::new();

        // Parse attributes
        for attr in start_tag.attributes().filter_map(|a| a.ok()) {
            let val = String::from_utf8_lossy(&attr.value);
            match attr.key.local_name().as_ref() {
                b"dir" => {
                    f_dir = val.parse().ok();
                }
                #[cfg(feature = "extra-attrs")]
                unknown => {
                    let key = String::from_utf8_lossy(attr.key.as_ref()).into_owned();
                    extra_attrs.insert(key, val.into_owned());
                }
                #[cfg(not(feature = "extra-attrs"))]
                _ => {}
            }
        }

        if !is_empty {
            let mut buf = Vec::new();
            loop {
                match reader.read_event_into(&mut buf)? {
                    Event::End(_) => break,
                    Event::Eof => break,
                    _ => {}
                }
                buf.clear();
            }
        }

        Ok(Self {
            dir: f_dir,
            #[cfg(feature = "extra-attrs")]
            extra_attrs,
        })
    }
}

impl FromXml for CTInOutTransition {
    fn from_xml<R: BufRead>(
        reader: &mut Reader<R>,
        start_tag: &BytesStart,
        is_empty: bool,
    ) -> Result<Self, ParseError> {
        let mut f_dir = None;
        #[cfg(feature = "extra-attrs")]
        let mut extra_attrs = std::collections::HashMap::new();

        // Parse attributes
        for attr in start_tag.attributes().filter_map(|a| a.ok()) {
            let val = String::from_utf8_lossy(&attr.value);
            match attr.key.local_name().as_ref() {
                b"dir" => {
                    f_dir = val.parse().ok();
                }
                #[cfg(feature = "extra-attrs")]
                unknown => {
                    let key = String::from_utf8_lossy(attr.key.as_ref()).into_owned();
                    extra_attrs.insert(key, val.into_owned());
                }
                #[cfg(not(feature = "extra-attrs"))]
                _ => {}
            }
        }

        if !is_empty {
            let mut buf = Vec::new();
            loop {
                match reader.read_event_into(&mut buf)? {
                    Event::End(_) => break,
                    Event::Eof => break,
                    _ => {}
                }
                buf.clear();
            }
        }

        Ok(Self {
            dir: f_dir,
            #[cfg(feature = "extra-attrs")]
            extra_attrs,
        })
    }
}

impl FromXml for CTOptionalBlackTransition {
    fn from_xml<R: BufRead>(
        reader: &mut Reader<R>,
        start_tag: &BytesStart,
        is_empty: bool,
    ) -> Result<Self, ParseError> {
        let mut f_thru_blk = None;
        #[cfg(feature = "extra-attrs")]
        let mut extra_attrs = std::collections::HashMap::new();

        // Parse attributes
        for attr in start_tag.attributes().filter_map(|a| a.ok()) {
            let val = String::from_utf8_lossy(&attr.value);
            match attr.key.local_name().as_ref() {
                b"thruBlk" => {
                    f_thru_blk = Some(val == "true" || val == "1");
                }
                #[cfg(feature = "extra-attrs")]
                unknown => {
                    let key = String::from_utf8_lossy(attr.key.as_ref()).into_owned();
                    extra_attrs.insert(key, val.into_owned());
                }
                #[cfg(not(feature = "extra-attrs"))]
                _ => {}
            }
        }

        if !is_empty {
            let mut buf = Vec::new();
            loop {
                match reader.read_event_into(&mut buf)? {
                    Event::End(_) => break,
                    Event::Eof => break,
                    _ => {}
                }
                buf.clear();
            }
        }

        Ok(Self {
            thru_blk: f_thru_blk,
            #[cfg(feature = "extra-attrs")]
            extra_attrs,
        })
    }
}

impl FromXml for CTSplitTransition {
    fn from_xml<R: BufRead>(
        reader: &mut Reader<R>,
        start_tag: &BytesStart,
        is_empty: bool,
    ) -> Result<Self, ParseError> {
        let mut f_orient = None;
        let mut f_dir = None;
        #[cfg(feature = "extra-attrs")]
        let mut extra_attrs = std::collections::HashMap::new();

        // Parse attributes
        for attr in start_tag.attributes().filter_map(|a| a.ok()) {
            let val = String::from_utf8_lossy(&attr.value);
            match attr.key.local_name().as_ref() {
                b"orient" => {
                    f_orient = val.parse().ok();
                }
                b"dir" => {
                    f_dir = val.parse().ok();
                }
                #[cfg(feature = "extra-attrs")]
                unknown => {
                    let key = String::from_utf8_lossy(attr.key.as_ref()).into_owned();
                    extra_attrs.insert(key, val.into_owned());
                }
                #[cfg(not(feature = "extra-attrs"))]
                _ => {}
            }
        }

        if !is_empty {
            let mut buf = Vec::new();
            loop {
                match reader.read_event_into(&mut buf)? {
                    Event::End(_) => break,
                    Event::Eof => break,
                    _ => {}
                }
                buf.clear();
            }
        }

        Ok(Self {
            orient: f_orient,
            dir: f_dir,
            #[cfg(feature = "extra-attrs")]
            extra_attrs,
        })
    }
}

impl FromXml for CTWheelTransition {
    fn from_xml<R: BufRead>(
        reader: &mut Reader<R>,
        start_tag: &BytesStart,
        is_empty: bool,
    ) -> Result<Self, ParseError> {
        let mut f_spokes = None;
        #[cfg(feature = "extra-attrs")]
        let mut extra_attrs = std::collections::HashMap::new();

        // Parse attributes
        for attr in start_tag.attributes().filter_map(|a| a.ok()) {
            let val = String::from_utf8_lossy(&attr.value);
            match attr.key.local_name().as_ref() {
                b"spokes" => {
                    f_spokes = val.parse().ok();
                }
                #[cfg(feature = "extra-attrs")]
                unknown => {
                    let key = String::from_utf8_lossy(attr.key.as_ref()).into_owned();
                    extra_attrs.insert(key, val.into_owned());
                }
                #[cfg(not(feature = "extra-attrs"))]
                _ => {}
            }
        }

        if !is_empty {
            let mut buf = Vec::new();
            loop {
                match reader.read_event_into(&mut buf)? {
                    Event::End(_) => break,
                    Event::Eof => break,
                    _ => {}
                }
                buf.clear();
            }
        }

        Ok(Self {
            spokes: f_spokes,
            #[cfg(feature = "extra-attrs")]
            extra_attrs,
        })
    }
}

impl FromXml for CTTransitionStartSoundAction {
    fn from_xml<R: BufRead>(
        reader: &mut Reader<R>,
        start_tag: &BytesStart,
        is_empty: bool,
    ) -> Result<Self, ParseError> {
        let mut f_loop = None;
        let mut f_snd: Option<Box<ooxml_dml::types::CTEmbeddedWAVAudioFile>> = None;
        #[cfg(feature = "extra-attrs")]
        let mut extra_attrs = std::collections::HashMap::new();
        #[cfg(feature = "extra-children")]
        let mut extra_children = Vec::new();
        #[cfg(feature = "extra-children")]
        let mut child_idx: usize = 0;

        // Parse attributes
        for attr in start_tag.attributes().filter_map(|a| a.ok()) {
            let val = String::from_utf8_lossy(&attr.value);
            match attr.key.local_name().as_ref() {
                b"loop" => {
                    f_loop = Some(val == "true" || val == "1");
                }
                #[cfg(feature = "extra-attrs")]
                unknown => {
                    let key = String::from_utf8_lossy(attr.key.as_ref()).into_owned();
                    extra_attrs.insert(key, val.into_owned());
                }
                #[cfg(not(feature = "extra-attrs"))]
                _ => {}
            }
        }

        // Parse child elements
        if !is_empty {
            let mut buf = Vec::new();
            loop {
                match reader.read_event_into(&mut buf)? {
                    Event::Start(e) => {
                        match e.local_name().as_ref() {
                            b"snd" => {
                                f_snd = Some(Box::new(
                                    ooxml_dml::types::CTEmbeddedWAVAudioFile::from_xml(
                                        reader, &e, false,
                                    )?,
                                ));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            #[cfg(feature = "extra-children")]
                            _ => {
                                // Capture unknown element for roundtrip
                                let elem = RawXmlElement::from_reader(reader, &e)?;
                                extra_children.push(PositionedNode::new(
                                    child_idx,
                                    RawXmlNode::Element(elem),
                                ));
                                child_idx += 1;
                            }
                            #[cfg(not(feature = "extra-children"))]
                            _ => {
                                // Skip unknown element
                                skip_element(reader)?;
                            }
                        }
                    }
                    Event::Empty(e) => {
                        match e.local_name().as_ref() {
                            b"snd" => {
                                f_snd = Some(Box::new(
                                    ooxml_dml::types::CTEmbeddedWAVAudioFile::from_xml(
                                        reader, &e, true,
                                    )?,
                                ));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            #[cfg(feature = "extra-children")]
                            _ => {
                                // Capture unknown empty element for roundtrip
                                let elem = RawXmlElement::from_empty(&e);
                                extra_children.push(PositionedNode::new(
                                    child_idx,
                                    RawXmlNode::Element(elem),
                                ));
                                child_idx += 1;
                            }
                            #[cfg(not(feature = "extra-children"))]
                            _ => {}
                        }
                    }
                    Event::End(_) => break,
                    Event::Eof => break,
                    _ => {}
                }
                buf.clear();
            }
        }

        Ok(Self {
            r#loop: f_loop,
            snd: f_snd.ok_or_else(|| ParseError::MissingAttribute("snd".to_string()))?,
            #[cfg(feature = "extra-attrs")]
            extra_attrs,
            #[cfg(feature = "extra-children")]
            extra_children,
        })
    }
}

impl FromXml for CTTransitionSoundAction {
    fn from_xml<R: BufRead>(
        reader: &mut Reader<R>,
        start_tag: &BytesStart,
        is_empty: bool,
    ) -> Result<Self, ParseError> {
        let mut f_st_snd = None;
        let mut f_end_snd = None;
        #[cfg(feature = "extra-children")]
        let mut extra_children = Vec::new();
        #[cfg(feature = "extra-children")]
        let mut child_idx: usize = 0;

        // Parse child elements
        if !is_empty {
            let mut buf = Vec::new();
            loop {
                match reader.read_event_into(&mut buf)? {
                    Event::Start(e) => {
                        match e.local_name().as_ref() {
                            b"stSnd" => {
                                f_st_snd = Some(Box::new(CTTransitionStartSoundAction::from_xml(
                                    reader, &e, false,
                                )?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            b"endSnd" => {
                                f_end_snd = Some(Box::new(CTEmpty::from_xml(reader, &e, false)?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            #[cfg(feature = "extra-children")]
                            _ => {
                                // Capture unknown element for roundtrip
                                let elem = RawXmlElement::from_reader(reader, &e)?;
                                extra_children.push(PositionedNode::new(
                                    child_idx,
                                    RawXmlNode::Element(elem),
                                ));
                                child_idx += 1;
                            }
                            #[cfg(not(feature = "extra-children"))]
                            _ => {
                                // Skip unknown element
                                skip_element(reader)?;
                            }
                        }
                    }
                    Event::Empty(e) => {
                        match e.local_name().as_ref() {
                            b"stSnd" => {
                                f_st_snd = Some(Box::new(CTTransitionStartSoundAction::from_xml(
                                    reader, &e, true,
                                )?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            b"endSnd" => {
                                f_end_snd = Some(Box::new(CTEmpty::from_xml(reader, &e, true)?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            #[cfg(feature = "extra-children")]
                            _ => {
                                // Capture unknown empty element for roundtrip
                                let elem = RawXmlElement::from_empty(&e);
                                extra_children.push(PositionedNode::new(
                                    child_idx,
                                    RawXmlNode::Element(elem),
                                ));
                                child_idx += 1;
                            }
                            #[cfg(not(feature = "extra-children"))]
                            _ => {}
                        }
                    }
                    Event::End(_) => break,
                    Event::Eof => break,
                    _ => {}
                }
                buf.clear();
            }
        }

        Ok(Self {
            st_snd: f_st_snd,
            end_snd: f_end_snd,
            #[cfg(feature = "extra-children")]
            extra_children,
        })
    }
}

impl FromXml for SlideTransition {
    fn from_xml<R: BufRead>(
        reader: &mut Reader<R>,
        start_tag: &BytesStart,
        is_empty: bool,
    ) -> Result<Self, ParseError> {
        let mut f_spd = None;
        let mut f_adv_click = None;
        let mut f_adv_tm = None;
        let mut f_blinds = None;
        let mut f_checker = None;
        let mut f_circle = None;
        let mut f_dissolve = None;
        let mut f_comb = None;
        let mut f_cover = None;
        let mut f_cut = None;
        let mut f_diamond = None;
        let mut f_fade = None;
        let mut f_newsflash = None;
        let mut f_plus = None;
        let mut f_pull = None;
        let mut f_push = None;
        let mut f_random = None;
        let mut f_random_bar = None;
        let mut f_split = None;
        let mut f_strips = None;
        let mut f_wedge = None;
        let mut f_wheel = None;
        let mut f_wipe = None;
        let mut f_zoom = None;
        let mut f_snd_ac = None;
        let mut f_ext_lst = None;
        #[cfg(feature = "extra-attrs")]
        let mut extra_attrs = std::collections::HashMap::new();
        #[cfg(feature = "extra-children")]
        let mut extra_children = Vec::new();
        #[cfg(feature = "extra-children")]
        let mut child_idx: usize = 0;

        // Parse attributes
        for attr in start_tag.attributes().filter_map(|a| a.ok()) {
            let val = String::from_utf8_lossy(&attr.value);
            match attr.key.local_name().as_ref() {
                b"spd" => {
                    f_spd = val.parse().ok();
                }
                b"advClick" => {
                    f_adv_click = Some(val == "true" || val == "1");
                }
                b"advTm" => {
                    f_adv_tm = val.parse().ok();
                }
                #[cfg(feature = "extra-attrs")]
                unknown => {
                    let key = String::from_utf8_lossy(attr.key.as_ref()).into_owned();
                    extra_attrs.insert(key, val.into_owned());
                }
                #[cfg(not(feature = "extra-attrs"))]
                _ => {}
            }
        }

        // Parse child elements
        if !is_empty {
            let mut buf = Vec::new();
            loop {
                match reader.read_event_into(&mut buf)? {
                    Event::Start(e) => {
                        match e.local_name().as_ref() {
                            b"blinds" => {
                                f_blinds = Some(Box::new(CTOrientationTransition::from_xml(
                                    reader, &e, false,
                                )?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            b"checker" => {
                                f_checker = Some(Box::new(CTOrientationTransition::from_xml(
                                    reader, &e, false,
                                )?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            b"circle" => {
                                f_circle = Some(Box::new(CTEmpty::from_xml(reader, &e, false)?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            b"dissolve" => {
                                f_dissolve = Some(Box::new(CTEmpty::from_xml(reader, &e, false)?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            b"comb" => {
                                f_comb = Some(Box::new(CTOrientationTransition::from_xml(
                                    reader, &e, false,
                                )?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            b"cover" => {
                                f_cover = Some(Box::new(CTEightDirectionTransition::from_xml(
                                    reader, &e, false,
                                )?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            b"cut" => {
                                f_cut = Some(Box::new(CTOptionalBlackTransition::from_xml(
                                    reader, &e, false,
                                )?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            b"diamond" => {
                                f_diamond = Some(Box::new(CTEmpty::from_xml(reader, &e, false)?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            b"fade" => {
                                f_fade = Some(Box::new(CTOptionalBlackTransition::from_xml(
                                    reader, &e, false,
                                )?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            b"newsflash" => {
                                f_newsflash = Some(Box::new(CTEmpty::from_xml(reader, &e, false)?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            b"plus" => {
                                f_plus = Some(Box::new(CTEmpty::from_xml(reader, &e, false)?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            b"pull" => {
                                f_pull = Some(Box::new(CTEightDirectionTransition::from_xml(
                                    reader, &e, false,
                                )?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            b"push" => {
                                f_push = Some(Box::new(CTSideDirectionTransition::from_xml(
                                    reader, &e, false,
                                )?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            b"random" => {
                                f_random = Some(Box::new(CTEmpty::from_xml(reader, &e, false)?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            b"randomBar" => {
                                f_random_bar = Some(Box::new(CTOrientationTransition::from_xml(
                                    reader, &e, false,
                                )?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            b"split" => {
                                f_split =
                                    Some(Box::new(CTSplitTransition::from_xml(reader, &e, false)?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            b"strips" => {
                                f_strips = Some(Box::new(CTCornerDirectionTransition::from_xml(
                                    reader, &e, false,
                                )?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            b"wedge" => {
                                f_wedge = Some(Box::new(CTEmpty::from_xml(reader, &e, false)?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            b"wheel" => {
                                f_wheel =
                                    Some(Box::new(CTWheelTransition::from_xml(reader, &e, false)?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            b"wipe" => {
                                f_wipe = Some(Box::new(CTSideDirectionTransition::from_xml(
                                    reader, &e, false,
                                )?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            b"zoom" => {
                                f_zoom =
                                    Some(Box::new(CTInOutTransition::from_xml(reader, &e, false)?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            b"sndAc" => {
                                f_snd_ac = Some(Box::new(CTTransitionSoundAction::from_xml(
                                    reader, &e, false,
                                )?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            b"extLst" => {
                                f_ext_lst = Some(Box::new(CTExtensionListModify::from_xml(
                                    reader, &e, false,
                                )?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            #[cfg(feature = "extra-children")]
                            _ => {
                                // Capture unknown element for roundtrip
                                let elem = RawXmlElement::from_reader(reader, &e)?;
                                extra_children.push(PositionedNode::new(
                                    child_idx,
                                    RawXmlNode::Element(elem),
                                ));
                                child_idx += 1;
                            }
                            #[cfg(not(feature = "extra-children"))]
                            _ => {
                                // Skip unknown element
                                skip_element(reader)?;
                            }
                        }
                    }
                    Event::Empty(e) => {
                        match e.local_name().as_ref() {
                            b"blinds" => {
                                f_blinds = Some(Box::new(CTOrientationTransition::from_xml(
                                    reader, &e, true,
                                )?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            b"checker" => {
                                f_checker = Some(Box::new(CTOrientationTransition::from_xml(
                                    reader, &e, true,
                                )?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            b"circle" => {
                                f_circle = Some(Box::new(CTEmpty::from_xml(reader, &e, true)?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            b"dissolve" => {
                                f_dissolve = Some(Box::new(CTEmpty::from_xml(reader, &e, true)?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            b"comb" => {
                                f_comb = Some(Box::new(CTOrientationTransition::from_xml(
                                    reader, &e, true,
                                )?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            b"cover" => {
                                f_cover = Some(Box::new(CTEightDirectionTransition::from_xml(
                                    reader, &e, true,
                                )?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            b"cut" => {
                                f_cut = Some(Box::new(CTOptionalBlackTransition::from_xml(
                                    reader, &e, true,
                                )?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            b"diamond" => {
                                f_diamond = Some(Box::new(CTEmpty::from_xml(reader, &e, true)?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            b"fade" => {
                                f_fade = Some(Box::new(CTOptionalBlackTransition::from_xml(
                                    reader, &e, true,
                                )?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            b"newsflash" => {
                                f_newsflash = Some(Box::new(CTEmpty::from_xml(reader, &e, true)?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            b"plus" => {
                                f_plus = Some(Box::new(CTEmpty::from_xml(reader, &e, true)?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            b"pull" => {
                                f_pull = Some(Box::new(CTEightDirectionTransition::from_xml(
                                    reader, &e, true,
                                )?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            b"push" => {
                                f_push = Some(Box::new(CTSideDirectionTransition::from_xml(
                                    reader, &e, true,
                                )?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            b"random" => {
                                f_random = Some(Box::new(CTEmpty::from_xml(reader, &e, true)?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            b"randomBar" => {
                                f_random_bar = Some(Box::new(CTOrientationTransition::from_xml(
                                    reader, &e, true,
                                )?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            b"split" => {
                                f_split =
                                    Some(Box::new(CTSplitTransition::from_xml(reader, &e, true)?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            b"strips" => {
                                f_strips = Some(Box::new(CTCornerDirectionTransition::from_xml(
                                    reader, &e, true,
                                )?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            b"wedge" => {
                                f_wedge = Some(Box::new(CTEmpty::from_xml(reader, &e, true)?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            b"wheel" => {
                                f_wheel =
                                    Some(Box::new(CTWheelTransition::from_xml(reader, &e, true)?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            b"wipe" => {
                                f_wipe = Some(Box::new(CTSideDirectionTransition::from_xml(
                                    reader, &e, true,
                                )?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            b"zoom" => {
                                f_zoom =
                                    Some(Box::new(CTInOutTransition::from_xml(reader, &e, true)?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            b"sndAc" => {
                                f_snd_ac = Some(Box::new(CTTransitionSoundAction::from_xml(
                                    reader, &e, true,
                                )?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            b"extLst" => {
                                f_ext_lst = Some(Box::new(CTExtensionListModify::from_xml(
                                    reader, &e, true,
                                )?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            #[cfg(feature = "extra-children")]
                            _ => {
                                // Capture unknown empty element for roundtrip
                                let elem = RawXmlElement::from_empty(&e);
                                extra_children.push(PositionedNode::new(
                                    child_idx,
                                    RawXmlNode::Element(elem),
                                ));
                                child_idx += 1;
                            }
                            #[cfg(not(feature = "extra-children"))]
                            _ => {}
                        }
                    }
                    Event::End(_) => break,
                    Event::Eof => break,
                    _ => {}
                }
                buf.clear();
            }
        }

        Ok(Self {
            spd: f_spd,
            adv_click: f_adv_click,
            adv_tm: f_adv_tm,
            blinds: f_blinds,
            checker: f_checker,
            circle: f_circle,
            dissolve: f_dissolve,
            comb: f_comb,
            cover: f_cover,
            cut: f_cut,
            diamond: f_diamond,
            fade: f_fade,
            newsflash: f_newsflash,
            plus: f_plus,
            pull: f_pull,
            push: f_push,
            random: f_random,
            random_bar: f_random_bar,
            split: f_split,
            strips: f_strips,
            wedge: f_wedge,
            wheel: f_wheel,
            wipe: f_wipe,
            zoom: f_zoom,
            snd_ac: f_snd_ac,
            ext_lst: f_ext_lst,
            #[cfg(feature = "extra-attrs")]
            extra_attrs,
            #[cfg(feature = "extra-children")]
            extra_children,
        })
    }
}

impl FromXml for CTTLIterateIntervalTime {
    fn from_xml<R: BufRead>(
        reader: &mut Reader<R>,
        start_tag: &BytesStart,
        is_empty: bool,
    ) -> Result<Self, ParseError> {
        let mut f_value: Option<STTLTime> = None;
        #[cfg(feature = "extra-attrs")]
        let mut extra_attrs = std::collections::HashMap::new();

        // Parse attributes
        for attr in start_tag.attributes().filter_map(|a| a.ok()) {
            let val = String::from_utf8_lossy(&attr.value);
            match attr.key.local_name().as_ref() {
                b"val" => {
                    f_value = Some(val.into_owned());
                }
                #[cfg(feature = "extra-attrs")]
                unknown => {
                    let key = String::from_utf8_lossy(attr.key.as_ref()).into_owned();
                    extra_attrs.insert(key, val.into_owned());
                }
                #[cfg(not(feature = "extra-attrs"))]
                _ => {}
            }
        }

        if !is_empty {
            let mut buf = Vec::new();
            loop {
                match reader.read_event_into(&mut buf)? {
                    Event::End(_) => break,
                    Event::Eof => break,
                    _ => {}
                }
                buf.clear();
            }
        }

        Ok(Self {
            value: f_value.ok_or_else(|| ParseError::MissingAttribute("val".to_string()))?,
            #[cfg(feature = "extra-attrs")]
            extra_attrs,
        })
    }
}

impl FromXml for CTTLIterateIntervalPercentage {
    fn from_xml<R: BufRead>(
        reader: &mut Reader<R>,
        start_tag: &BytesStart,
        is_empty: bool,
    ) -> Result<Self, ParseError> {
        let mut f_value: Option<ooxml_dml::types::STPositivePercentage> = None;
        #[cfg(feature = "extra-attrs")]
        let mut extra_attrs = std::collections::HashMap::new();

        // Parse attributes
        for attr in start_tag.attributes().filter_map(|a| a.ok()) {
            let val = String::from_utf8_lossy(&attr.value);
            match attr.key.local_name().as_ref() {
                b"val" => {
                    f_value = val.parse().ok();
                }
                #[cfg(feature = "extra-attrs")]
                unknown => {
                    let key = String::from_utf8_lossy(attr.key.as_ref()).into_owned();
                    extra_attrs.insert(key, val.into_owned());
                }
                #[cfg(not(feature = "extra-attrs"))]
                _ => {}
            }
        }

        if !is_empty {
            let mut buf = Vec::new();
            loop {
                match reader.read_event_into(&mut buf)? {
                    Event::End(_) => break,
                    Event::Eof => break,
                    _ => {}
                }
                buf.clear();
            }
        }

        Ok(Self {
            value: f_value.ok_or_else(|| ParseError::MissingAttribute("val".to_string()))?,
            #[cfg(feature = "extra-attrs")]
            extra_attrs,
        })
    }
}

impl FromXml for CTTLIterateData {
    fn from_xml<R: BufRead>(
        reader: &mut Reader<R>,
        start_tag: &BytesStart,
        is_empty: bool,
    ) -> Result<Self, ParseError> {
        let mut f_type = None;
        let mut f_backwards = None;
        let mut f_tm_abs = None;
        let mut f_tm_pct = None;
        #[cfg(feature = "extra-attrs")]
        let mut extra_attrs = std::collections::HashMap::new();
        #[cfg(feature = "extra-children")]
        let mut extra_children = Vec::new();
        #[cfg(feature = "extra-children")]
        let mut child_idx: usize = 0;

        // Parse attributes
        for attr in start_tag.attributes().filter_map(|a| a.ok()) {
            let val = String::from_utf8_lossy(&attr.value);
            match attr.key.local_name().as_ref() {
                b"type" => {
                    f_type = val.parse().ok();
                }
                b"backwards" => {
                    f_backwards = Some(val == "true" || val == "1");
                }
                #[cfg(feature = "extra-attrs")]
                unknown => {
                    let key = String::from_utf8_lossy(attr.key.as_ref()).into_owned();
                    extra_attrs.insert(key, val.into_owned());
                }
                #[cfg(not(feature = "extra-attrs"))]
                _ => {}
            }
        }

        // Parse child elements
        if !is_empty {
            let mut buf = Vec::new();
            loop {
                match reader.read_event_into(&mut buf)? {
                    Event::Start(e) => {
                        match e.local_name().as_ref() {
                            b"tmAbs" => {
                                f_tm_abs = Some(Box::new(CTTLIterateIntervalTime::from_xml(
                                    reader, &e, false,
                                )?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            b"tmPct" => {
                                f_tm_pct = Some(Box::new(CTTLIterateIntervalPercentage::from_xml(
                                    reader, &e, false,
                                )?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            #[cfg(feature = "extra-children")]
                            _ => {
                                // Capture unknown element for roundtrip
                                let elem = RawXmlElement::from_reader(reader, &e)?;
                                extra_children.push(PositionedNode::new(
                                    child_idx,
                                    RawXmlNode::Element(elem),
                                ));
                                child_idx += 1;
                            }
                            #[cfg(not(feature = "extra-children"))]
                            _ => {
                                // Skip unknown element
                                skip_element(reader)?;
                            }
                        }
                    }
                    Event::Empty(e) => {
                        match e.local_name().as_ref() {
                            b"tmAbs" => {
                                f_tm_abs = Some(Box::new(CTTLIterateIntervalTime::from_xml(
                                    reader, &e, true,
                                )?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            b"tmPct" => {
                                f_tm_pct = Some(Box::new(CTTLIterateIntervalPercentage::from_xml(
                                    reader, &e, true,
                                )?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            #[cfg(feature = "extra-children")]
                            _ => {
                                // Capture unknown empty element for roundtrip
                                let elem = RawXmlElement::from_empty(&e);
                                extra_children.push(PositionedNode::new(
                                    child_idx,
                                    RawXmlNode::Element(elem),
                                ));
                                child_idx += 1;
                            }
                            #[cfg(not(feature = "extra-children"))]
                            _ => {}
                        }
                    }
                    Event::End(_) => break,
                    Event::Eof => break,
                    _ => {}
                }
                buf.clear();
            }
        }

        Ok(Self {
            r#type: f_type,
            backwards: f_backwards,
            tm_abs: f_tm_abs,
            tm_pct: f_tm_pct,
            #[cfg(feature = "extra-attrs")]
            extra_attrs,
            #[cfg(feature = "extra-children")]
            extra_children,
        })
    }
}

impl FromXml for CTTLSubShapeId {
    fn from_xml<R: BufRead>(
        reader: &mut Reader<R>,
        start_tag: &BytesStart,
        is_empty: bool,
    ) -> Result<Self, ParseError> {
        let mut f_spid: Option<ooxml_dml::types::STShapeID> = None;
        #[cfg(feature = "extra-attrs")]
        let mut extra_attrs = std::collections::HashMap::new();

        // Parse attributes
        for attr in start_tag.attributes().filter_map(|a| a.ok()) {
            let val = String::from_utf8_lossy(&attr.value);
            match attr.key.local_name().as_ref() {
                b"spid" => {
                    f_spid = val.parse().ok();
                }
                #[cfg(feature = "extra-attrs")]
                unknown => {
                    let key = String::from_utf8_lossy(attr.key.as_ref()).into_owned();
                    extra_attrs.insert(key, val.into_owned());
                }
                #[cfg(not(feature = "extra-attrs"))]
                _ => {}
            }
        }

        if !is_empty {
            let mut buf = Vec::new();
            loop {
                match reader.read_event_into(&mut buf)? {
                    Event::End(_) => break,
                    Event::Eof => break,
                    _ => {}
                }
                buf.clear();
            }
        }

        Ok(Self {
            spid: f_spid.ok_or_else(|| ParseError::MissingAttribute("spid".to_string()))?,
            #[cfg(feature = "extra-attrs")]
            extra_attrs,
        })
    }
}

impl FromXml for CTTLTextTargetElement {
    fn from_xml<R: BufRead>(
        reader: &mut Reader<R>,
        start_tag: &BytesStart,
        is_empty: bool,
    ) -> Result<Self, ParseError> {
        let mut f_char_rg = None;
        let mut f_p_rg = None;
        #[cfg(feature = "extra-children")]
        let mut extra_children = Vec::new();
        #[cfg(feature = "extra-children")]
        let mut child_idx: usize = 0;

        // Parse child elements
        if !is_empty {
            let mut buf = Vec::new();
            loop {
                match reader.read_event_into(&mut buf)? {
                    Event::Start(e) => {
                        match e.local_name().as_ref() {
                            b"charRg" => {
                                f_char_rg =
                                    Some(Box::new(CTIndexRange::from_xml(reader, &e, false)?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            b"pRg" => {
                                f_p_rg = Some(Box::new(CTIndexRange::from_xml(reader, &e, false)?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            #[cfg(feature = "extra-children")]
                            _ => {
                                // Capture unknown element for roundtrip
                                let elem = RawXmlElement::from_reader(reader, &e)?;
                                extra_children.push(PositionedNode::new(
                                    child_idx,
                                    RawXmlNode::Element(elem),
                                ));
                                child_idx += 1;
                            }
                            #[cfg(not(feature = "extra-children"))]
                            _ => {
                                // Skip unknown element
                                skip_element(reader)?;
                            }
                        }
                    }
                    Event::Empty(e) => {
                        match e.local_name().as_ref() {
                            b"charRg" => {
                                f_char_rg =
                                    Some(Box::new(CTIndexRange::from_xml(reader, &e, true)?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            b"pRg" => {
                                f_p_rg = Some(Box::new(CTIndexRange::from_xml(reader, &e, true)?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            #[cfg(feature = "extra-children")]
                            _ => {
                                // Capture unknown empty element for roundtrip
                                let elem = RawXmlElement::from_empty(&e);
                                extra_children.push(PositionedNode::new(
                                    child_idx,
                                    RawXmlNode::Element(elem),
                                ));
                                child_idx += 1;
                            }
                            #[cfg(not(feature = "extra-children"))]
                            _ => {}
                        }
                    }
                    Event::End(_) => break,
                    Event::Eof => break,
                    _ => {}
                }
                buf.clear();
            }
        }

        Ok(Self {
            char_rg: f_char_rg,
            p_rg: f_p_rg,
            #[cfg(feature = "extra-children")]
            extra_children,
        })
    }
}

impl FromXml for CTTLOleChartTargetElement {
    fn from_xml<R: BufRead>(
        reader: &mut Reader<R>,
        start_tag: &BytesStart,
        is_empty: bool,
    ) -> Result<Self, ParseError> {
        let mut f_type: Option<STTLChartSubelementType> = None;
        let mut f_lvl = None;
        #[cfg(feature = "extra-attrs")]
        let mut extra_attrs = std::collections::HashMap::new();

        // Parse attributes
        for attr in start_tag.attributes().filter_map(|a| a.ok()) {
            let val = String::from_utf8_lossy(&attr.value);
            match attr.key.local_name().as_ref() {
                b"type" => {
                    f_type = val.parse().ok();
                }
                b"lvl" => {
                    f_lvl = val.parse().ok();
                }
                #[cfg(feature = "extra-attrs")]
                unknown => {
                    let key = String::from_utf8_lossy(attr.key.as_ref()).into_owned();
                    extra_attrs.insert(key, val.into_owned());
                }
                #[cfg(not(feature = "extra-attrs"))]
                _ => {}
            }
        }

        if !is_empty {
            let mut buf = Vec::new();
            loop {
                match reader.read_event_into(&mut buf)? {
                    Event::End(_) => break,
                    Event::Eof => break,
                    _ => {}
                }
                buf.clear();
            }
        }

        Ok(Self {
            r#type: f_type.ok_or_else(|| ParseError::MissingAttribute("type".to_string()))?,
            lvl: f_lvl,
            #[cfg(feature = "extra-attrs")]
            extra_attrs,
        })
    }
}

impl FromXml for CTTLShapeTargetElement {
    fn from_xml<R: BufRead>(
        reader: &mut Reader<R>,
        start_tag: &BytesStart,
        is_empty: bool,
    ) -> Result<Self, ParseError> {
        let mut f_spid: Option<ooxml_dml::types::STDrawingElementId> = None;
        let mut f_bg = None;
        let mut f_sub_sp = None;
        let mut f_ole_chart_el = None;
        let mut f_tx_el = None;
        let mut f_graphic_el = None;
        #[cfg(feature = "extra-attrs")]
        let mut extra_attrs = std::collections::HashMap::new();
        #[cfg(feature = "extra-children")]
        let mut extra_children = Vec::new();
        #[cfg(feature = "extra-children")]
        let mut child_idx: usize = 0;

        // Parse attributes
        for attr in start_tag.attributes().filter_map(|a| a.ok()) {
            let val = String::from_utf8_lossy(&attr.value);
            match attr.key.local_name().as_ref() {
                b"spid" => {
                    f_spid = val.parse().ok();
                }
                #[cfg(feature = "extra-attrs")]
                unknown => {
                    let key = String::from_utf8_lossy(attr.key.as_ref()).into_owned();
                    extra_attrs.insert(key, val.into_owned());
                }
                #[cfg(not(feature = "extra-attrs"))]
                _ => {}
            }
        }

        // Parse child elements
        if !is_empty {
            let mut buf = Vec::new();
            loop {
                match reader.read_event_into(&mut buf)? {
                    Event::Start(e) => {
                        match e.local_name().as_ref() {
                            b"bg" => {
                                f_bg = Some(Box::new(CTEmpty::from_xml(reader, &e, false)?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            b"subSp" => {
                                f_sub_sp =
                                    Some(Box::new(CTTLSubShapeId::from_xml(reader, &e, false)?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            b"oleChartEl" => {
                                f_ole_chart_el = Some(Box::new(
                                    CTTLOleChartTargetElement::from_xml(reader, &e, false)?,
                                ));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            b"txEl" => {
                                f_tx_el = Some(Box::new(CTTLTextTargetElement::from_xml(
                                    reader, &e, false,
                                )?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            b"graphicEl" => {
                                f_graphic_el = Some(Box::new(
                                    ooxml_dml::types::CTAnimationElementChoice::from_xml(
                                        reader, &e, false,
                                    )?,
                                ));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            #[cfg(feature = "extra-children")]
                            _ => {
                                // Capture unknown element for roundtrip
                                let elem = RawXmlElement::from_reader(reader, &e)?;
                                extra_children.push(PositionedNode::new(
                                    child_idx,
                                    RawXmlNode::Element(elem),
                                ));
                                child_idx += 1;
                            }
                            #[cfg(not(feature = "extra-children"))]
                            _ => {
                                // Skip unknown element
                                skip_element(reader)?;
                            }
                        }
                    }
                    Event::Empty(e) => {
                        match e.local_name().as_ref() {
                            b"bg" => {
                                f_bg = Some(Box::new(CTEmpty::from_xml(reader, &e, true)?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            b"subSp" => {
                                f_sub_sp =
                                    Some(Box::new(CTTLSubShapeId::from_xml(reader, &e, true)?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            b"oleChartEl" => {
                                f_ole_chart_el = Some(Box::new(
                                    CTTLOleChartTargetElement::from_xml(reader, &e, true)?,
                                ));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            b"txEl" => {
                                f_tx_el = Some(Box::new(CTTLTextTargetElement::from_xml(
                                    reader, &e, true,
                                )?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            b"graphicEl" => {
                                f_graphic_el = Some(Box::new(
                                    ooxml_dml::types::CTAnimationElementChoice::from_xml(
                                        reader, &e, true,
                                    )?,
                                ));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            #[cfg(feature = "extra-children")]
                            _ => {
                                // Capture unknown empty element for roundtrip
                                let elem = RawXmlElement::from_empty(&e);
                                extra_children.push(PositionedNode::new(
                                    child_idx,
                                    RawXmlNode::Element(elem),
                                ));
                                child_idx += 1;
                            }
                            #[cfg(not(feature = "extra-children"))]
                            _ => {}
                        }
                    }
                    Event::End(_) => break,
                    Event::Eof => break,
                    _ => {}
                }
                buf.clear();
            }
        }

        Ok(Self {
            spid: f_spid.ok_or_else(|| ParseError::MissingAttribute("spid".to_string()))?,
            bg: f_bg,
            sub_sp: f_sub_sp,
            ole_chart_el: f_ole_chart_el,
            tx_el: f_tx_el,
            graphic_el: f_graphic_el,
            #[cfg(feature = "extra-attrs")]
            extra_attrs,
            #[cfg(feature = "extra-children")]
            extra_children,
        })
    }
}

impl FromXml for CTTLTimeTargetElement {
    fn from_xml<R: BufRead>(
        reader: &mut Reader<R>,
        start_tag: &BytesStart,
        is_empty: bool,
    ) -> Result<Self, ParseError> {
        let mut f_sld_tgt = None;
        let mut f_snd_tgt = None;
        let mut f_sp_tgt = None;
        let mut f_ink_tgt = None;
        #[cfg(feature = "extra-children")]
        let mut extra_children = Vec::new();
        #[cfg(feature = "extra-children")]
        let mut child_idx: usize = 0;

        // Parse child elements
        if !is_empty {
            let mut buf = Vec::new();
            loop {
                match reader.read_event_into(&mut buf)? {
                    Event::Start(e) => {
                        match e.local_name().as_ref() {
                            b"sldTgt" => {
                                f_sld_tgt = Some(Box::new(CTEmpty::from_xml(reader, &e, false)?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            b"sndTgt" => {
                                f_snd_tgt = Some(Box::new(
                                    ooxml_dml::types::CTEmbeddedWAVAudioFile::from_xml(
                                        reader, &e, false,
                                    )?,
                                ));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            b"spTgt" => {
                                f_sp_tgt = Some(Box::new(CTTLShapeTargetElement::from_xml(
                                    reader, &e, false,
                                )?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            b"inkTgt" => {
                                f_ink_tgt =
                                    Some(Box::new(CTTLSubShapeId::from_xml(reader, &e, false)?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            #[cfg(feature = "extra-children")]
                            _ => {
                                // Capture unknown element for roundtrip
                                let elem = RawXmlElement::from_reader(reader, &e)?;
                                extra_children.push(PositionedNode::new(
                                    child_idx,
                                    RawXmlNode::Element(elem),
                                ));
                                child_idx += 1;
                            }
                            #[cfg(not(feature = "extra-children"))]
                            _ => {
                                // Skip unknown element
                                skip_element(reader)?;
                            }
                        }
                    }
                    Event::Empty(e) => {
                        match e.local_name().as_ref() {
                            b"sldTgt" => {
                                f_sld_tgt = Some(Box::new(CTEmpty::from_xml(reader, &e, true)?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            b"sndTgt" => {
                                f_snd_tgt = Some(Box::new(
                                    ooxml_dml::types::CTEmbeddedWAVAudioFile::from_xml(
                                        reader, &e, true,
                                    )?,
                                ));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            b"spTgt" => {
                                f_sp_tgt = Some(Box::new(CTTLShapeTargetElement::from_xml(
                                    reader, &e, true,
                                )?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            b"inkTgt" => {
                                f_ink_tgt =
                                    Some(Box::new(CTTLSubShapeId::from_xml(reader, &e, true)?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            #[cfg(feature = "extra-children")]
                            _ => {
                                // Capture unknown empty element for roundtrip
                                let elem = RawXmlElement::from_empty(&e);
                                extra_children.push(PositionedNode::new(
                                    child_idx,
                                    RawXmlNode::Element(elem),
                                ));
                                child_idx += 1;
                            }
                            #[cfg(not(feature = "extra-children"))]
                            _ => {}
                        }
                    }
                    Event::End(_) => break,
                    Event::Eof => break,
                    _ => {}
                }
                buf.clear();
            }
        }

        Ok(Self {
            sld_tgt: f_sld_tgt,
            snd_tgt: f_snd_tgt,
            sp_tgt: f_sp_tgt,
            ink_tgt: f_ink_tgt,
            #[cfg(feature = "extra-children")]
            extra_children,
        })
    }
}

impl FromXml for CTTLTriggerTimeNodeID {
    fn from_xml<R: BufRead>(
        reader: &mut Reader<R>,
        start_tag: &BytesStart,
        is_empty: bool,
    ) -> Result<Self, ParseError> {
        let mut f_value: Option<STTLTimeNodeID> = None;
        #[cfg(feature = "extra-attrs")]
        let mut extra_attrs = std::collections::HashMap::new();

        // Parse attributes
        for attr in start_tag.attributes().filter_map(|a| a.ok()) {
            let val = String::from_utf8_lossy(&attr.value);
            match attr.key.local_name().as_ref() {
                b"val" => {
                    f_value = val.parse().ok();
                }
                #[cfg(feature = "extra-attrs")]
                unknown => {
                    let key = String::from_utf8_lossy(attr.key.as_ref()).into_owned();
                    extra_attrs.insert(key, val.into_owned());
                }
                #[cfg(not(feature = "extra-attrs"))]
                _ => {}
            }
        }

        if !is_empty {
            let mut buf = Vec::new();
            loop {
                match reader.read_event_into(&mut buf)? {
                    Event::End(_) => break,
                    Event::Eof => break,
                    _ => {}
                }
                buf.clear();
            }
        }

        Ok(Self {
            value: f_value.ok_or_else(|| ParseError::MissingAttribute("val".to_string()))?,
            #[cfg(feature = "extra-attrs")]
            extra_attrs,
        })
    }
}

impl FromXml for CTTLTriggerRuntimeNode {
    fn from_xml<R: BufRead>(
        reader: &mut Reader<R>,
        start_tag: &BytesStart,
        is_empty: bool,
    ) -> Result<Self, ParseError> {
        let mut f_value: Option<STTLTriggerRuntimeNode> = None;
        #[cfg(feature = "extra-attrs")]
        let mut extra_attrs = std::collections::HashMap::new();

        // Parse attributes
        for attr in start_tag.attributes().filter_map(|a| a.ok()) {
            let val = String::from_utf8_lossy(&attr.value);
            match attr.key.local_name().as_ref() {
                b"val" => {
                    f_value = val.parse().ok();
                }
                #[cfg(feature = "extra-attrs")]
                unknown => {
                    let key = String::from_utf8_lossy(attr.key.as_ref()).into_owned();
                    extra_attrs.insert(key, val.into_owned());
                }
                #[cfg(not(feature = "extra-attrs"))]
                _ => {}
            }
        }

        if !is_empty {
            let mut buf = Vec::new();
            loop {
                match reader.read_event_into(&mut buf)? {
                    Event::End(_) => break,
                    Event::Eof => break,
                    _ => {}
                }
                buf.clear();
            }
        }

        Ok(Self {
            value: f_value.ok_or_else(|| ParseError::MissingAttribute("val".to_string()))?,
            #[cfg(feature = "extra-attrs")]
            extra_attrs,
        })
    }
}

impl FromXml for CTTLTimeCondition {
    fn from_xml<R: BufRead>(
        reader: &mut Reader<R>,
        start_tag: &BytesStart,
        is_empty: bool,
    ) -> Result<Self, ParseError> {
        let mut f_evt = None;
        let mut f_delay = None;
        let mut f_tgt_el = None;
        let mut f_tn = None;
        let mut f_rtn = None;
        #[cfg(feature = "extra-attrs")]
        let mut extra_attrs = std::collections::HashMap::new();
        #[cfg(feature = "extra-children")]
        let mut extra_children = Vec::new();
        #[cfg(feature = "extra-children")]
        let mut child_idx: usize = 0;

        // Parse attributes
        for attr in start_tag.attributes().filter_map(|a| a.ok()) {
            let val = String::from_utf8_lossy(&attr.value);
            match attr.key.local_name().as_ref() {
                b"evt" => {
                    f_evt = val.parse().ok();
                }
                b"delay" => {
                    f_delay = Some(val.into_owned());
                }
                #[cfg(feature = "extra-attrs")]
                unknown => {
                    let key = String::from_utf8_lossy(attr.key.as_ref()).into_owned();
                    extra_attrs.insert(key, val.into_owned());
                }
                #[cfg(not(feature = "extra-attrs"))]
                _ => {}
            }
        }

        // Parse child elements
        if !is_empty {
            let mut buf = Vec::new();
            loop {
                match reader.read_event_into(&mut buf)? {
                    Event::Start(e) => {
                        match e.local_name().as_ref() {
                            b"tgtEl" => {
                                f_tgt_el = Some(Box::new(CTTLTimeTargetElement::from_xml(
                                    reader, &e, false,
                                )?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            b"tn" => {
                                f_tn = Some(Box::new(CTTLTriggerTimeNodeID::from_xml(
                                    reader, &e, false,
                                )?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            b"rtn" => {
                                f_rtn = Some(Box::new(CTTLTriggerRuntimeNode::from_xml(
                                    reader, &e, false,
                                )?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            #[cfg(feature = "extra-children")]
                            _ => {
                                // Capture unknown element for roundtrip
                                let elem = RawXmlElement::from_reader(reader, &e)?;
                                extra_children.push(PositionedNode::new(
                                    child_idx,
                                    RawXmlNode::Element(elem),
                                ));
                                child_idx += 1;
                            }
                            #[cfg(not(feature = "extra-children"))]
                            _ => {
                                // Skip unknown element
                                skip_element(reader)?;
                            }
                        }
                    }
                    Event::Empty(e) => {
                        match e.local_name().as_ref() {
                            b"tgtEl" => {
                                f_tgt_el = Some(Box::new(CTTLTimeTargetElement::from_xml(
                                    reader, &e, true,
                                )?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            b"tn" => {
                                f_tn = Some(Box::new(CTTLTriggerTimeNodeID::from_xml(
                                    reader, &e, true,
                                )?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            b"rtn" => {
                                f_rtn = Some(Box::new(CTTLTriggerRuntimeNode::from_xml(
                                    reader, &e, true,
                                )?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            #[cfg(feature = "extra-children")]
                            _ => {
                                // Capture unknown empty element for roundtrip
                                let elem = RawXmlElement::from_empty(&e);
                                extra_children.push(PositionedNode::new(
                                    child_idx,
                                    RawXmlNode::Element(elem),
                                ));
                                child_idx += 1;
                            }
                            #[cfg(not(feature = "extra-children"))]
                            _ => {}
                        }
                    }
                    Event::End(_) => break,
                    Event::Eof => break,
                    _ => {}
                }
                buf.clear();
            }
        }

        Ok(Self {
            evt: f_evt,
            delay: f_delay,
            tgt_el: f_tgt_el,
            tn: f_tn,
            rtn: f_rtn,
            #[cfg(feature = "extra-attrs")]
            extra_attrs,
            #[cfg(feature = "extra-children")]
            extra_children,
        })
    }
}

impl FromXml for CTTLTimeConditionList {
    fn from_xml<R: BufRead>(
        reader: &mut Reader<R>,
        start_tag: &BytesStart,
        is_empty: bool,
    ) -> Result<Self, ParseError> {
        let mut f_cond = Vec::new();
        #[cfg(feature = "extra-children")]
        let mut extra_children = Vec::new();
        #[cfg(feature = "extra-children")]
        let mut child_idx: usize = 0;

        // Parse child elements
        if !is_empty {
            let mut buf = Vec::new();
            loop {
                match reader.read_event_into(&mut buf)? {
                    Event::Start(e) => {
                        match e.local_name().as_ref() {
                            b"cond" => {
                                f_cond.push(CTTLTimeCondition::from_xml(reader, &e, false)?);
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            #[cfg(feature = "extra-children")]
                            _ => {
                                // Capture unknown element for roundtrip
                                let elem = RawXmlElement::from_reader(reader, &e)?;
                                extra_children.push(PositionedNode::new(
                                    child_idx,
                                    RawXmlNode::Element(elem),
                                ));
                                child_idx += 1;
                            }
                            #[cfg(not(feature = "extra-children"))]
                            _ => {
                                // Skip unknown element
                                skip_element(reader)?;
                            }
                        }
                    }
                    Event::Empty(e) => {
                        match e.local_name().as_ref() {
                            b"cond" => {
                                f_cond.push(CTTLTimeCondition::from_xml(reader, &e, true)?);
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            #[cfg(feature = "extra-children")]
                            _ => {
                                // Capture unknown empty element for roundtrip
                                let elem = RawXmlElement::from_empty(&e);
                                extra_children.push(PositionedNode::new(
                                    child_idx,
                                    RawXmlNode::Element(elem),
                                ));
                                child_idx += 1;
                            }
                            #[cfg(not(feature = "extra-children"))]
                            _ => {}
                        }
                    }
                    Event::End(_) => break,
                    Event::Eof => break,
                    _ => {}
                }
                buf.clear();
            }
        }

        Ok(Self {
            cond: f_cond,
            #[cfg(feature = "extra-children")]
            extra_children,
        })
    }
}

impl FromXml for CTTimeNodeList {
    fn from_xml<R: BufRead>(
        reader: &mut Reader<R>,
        start_tag: &BytesStart,
        is_empty: bool,
    ) -> Result<Self, ParseError> {
        let mut f_par = None;
        let mut f_seq = None;
        let mut f_excl = None;
        let mut f_anim = None;
        let mut f_anim_clr = None;
        let mut f_anim_effect = None;
        let mut f_anim_motion = None;
        let mut f_anim_rot = None;
        let mut f_anim_scale = None;
        let mut f_cmd = None;
        let mut f_set = None;
        let mut f_audio = None;
        let mut f_video = None;
        #[cfg(feature = "extra-children")]
        let mut extra_children = Vec::new();
        #[cfg(feature = "extra-children")]
        let mut child_idx: usize = 0;

        // Parse child elements
        if !is_empty {
            let mut buf = Vec::new();
            loop {
                match reader.read_event_into(&mut buf)? {
                    Event::Start(e) => {
                        match e.local_name().as_ref() {
                            b"par" => {
                                f_par = Some(Box::new(CTTLCommonTimeNodeData::from_xml(
                                    reader, &e, false,
                                )?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            b"seq" => {
                                f_seq = Some(Box::new(CTTLTimeNodeSequence::from_xml(
                                    reader, &e, false,
                                )?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            b"excl" => {
                                f_excl = Some(Box::new(CTTLCommonTimeNodeData::from_xml(
                                    reader, &e, false,
                                )?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            b"anim" => {
                                f_anim = Some(Box::new(CTTLAnimateBehavior::from_xml(
                                    reader, &e, false,
                                )?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            b"animClr" => {
                                f_anim_clr = Some(Box::new(CTTLAnimateColorBehavior::from_xml(
                                    reader, &e, false,
                                )?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            b"animEffect" => {
                                f_anim_effect = Some(Box::new(
                                    CTTLAnimateEffectBehavior::from_xml(reader, &e, false)?,
                                ));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            b"animMotion" => {
                                f_anim_motion = Some(Box::new(
                                    CTTLAnimateMotionBehavior::from_xml(reader, &e, false)?,
                                ));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            b"animRot" => {
                                f_anim_rot = Some(Box::new(CTTLAnimateRotationBehavior::from_xml(
                                    reader, &e, false,
                                )?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            b"animScale" => {
                                f_anim_scale = Some(Box::new(CTTLAnimateScaleBehavior::from_xml(
                                    reader, &e, false,
                                )?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            b"cmd" => {
                                f_cmd = Some(Box::new(CTTLCommandBehavior::from_xml(
                                    reader, &e, false,
                                )?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            b"set" => {
                                f_set =
                                    Some(Box::new(CTTLSetBehavior::from_xml(reader, &e, false)?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            b"audio" => {
                                f_audio = Some(Box::new(CTTLMediaNodeAudio::from_xml(
                                    reader, &e, false,
                                )?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            b"video" => {
                                f_video = Some(Box::new(CTTLMediaNodeVideo::from_xml(
                                    reader, &e, false,
                                )?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            #[cfg(feature = "extra-children")]
                            _ => {
                                // Capture unknown element for roundtrip
                                let elem = RawXmlElement::from_reader(reader, &e)?;
                                extra_children.push(PositionedNode::new(
                                    child_idx,
                                    RawXmlNode::Element(elem),
                                ));
                                child_idx += 1;
                            }
                            #[cfg(not(feature = "extra-children"))]
                            _ => {
                                // Skip unknown element
                                skip_element(reader)?;
                            }
                        }
                    }
                    Event::Empty(e) => {
                        match e.local_name().as_ref() {
                            b"par" => {
                                f_par = Some(Box::new(CTTLCommonTimeNodeData::from_xml(
                                    reader, &e, true,
                                )?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            b"seq" => {
                                f_seq = Some(Box::new(CTTLTimeNodeSequence::from_xml(
                                    reader, &e, true,
                                )?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            b"excl" => {
                                f_excl = Some(Box::new(CTTLCommonTimeNodeData::from_xml(
                                    reader, &e, true,
                                )?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            b"anim" => {
                                f_anim = Some(Box::new(CTTLAnimateBehavior::from_xml(
                                    reader, &e, true,
                                )?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            b"animClr" => {
                                f_anim_clr = Some(Box::new(CTTLAnimateColorBehavior::from_xml(
                                    reader, &e, true,
                                )?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            b"animEffect" => {
                                f_anim_effect = Some(Box::new(
                                    CTTLAnimateEffectBehavior::from_xml(reader, &e, true)?,
                                ));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            b"animMotion" => {
                                f_anim_motion = Some(Box::new(
                                    CTTLAnimateMotionBehavior::from_xml(reader, &e, true)?,
                                ));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            b"animRot" => {
                                f_anim_rot = Some(Box::new(CTTLAnimateRotationBehavior::from_xml(
                                    reader, &e, true,
                                )?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            b"animScale" => {
                                f_anim_scale = Some(Box::new(CTTLAnimateScaleBehavior::from_xml(
                                    reader, &e, true,
                                )?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            b"cmd" => {
                                f_cmd = Some(Box::new(CTTLCommandBehavior::from_xml(
                                    reader, &e, true,
                                )?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            b"set" => {
                                f_set =
                                    Some(Box::new(CTTLSetBehavior::from_xml(reader, &e, true)?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            b"audio" => {
                                f_audio =
                                    Some(Box::new(CTTLMediaNodeAudio::from_xml(reader, &e, true)?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            b"video" => {
                                f_video =
                                    Some(Box::new(CTTLMediaNodeVideo::from_xml(reader, &e, true)?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            #[cfg(feature = "extra-children")]
                            _ => {
                                // Capture unknown empty element for roundtrip
                                let elem = RawXmlElement::from_empty(&e);
                                extra_children.push(PositionedNode::new(
                                    child_idx,
                                    RawXmlNode::Element(elem),
                                ));
                                child_idx += 1;
                            }
                            #[cfg(not(feature = "extra-children"))]
                            _ => {}
                        }
                    }
                    Event::End(_) => break,
                    Event::Eof => break,
                    _ => {}
                }
                buf.clear();
            }
        }

        Ok(Self {
            par: f_par,
            seq: f_seq,
            excl: f_excl,
            anim: f_anim,
            anim_clr: f_anim_clr,
            anim_effect: f_anim_effect,
            anim_motion: f_anim_motion,
            anim_rot: f_anim_rot,
            anim_scale: f_anim_scale,
            cmd: f_cmd,
            set: f_set,
            audio: f_audio,
            video: f_video,
            #[cfg(feature = "extra-children")]
            extra_children,
        })
    }
}

impl FromXml for CTTLCommonTimeNodeData {
    fn from_xml<R: BufRead>(
        reader: &mut Reader<R>,
        start_tag: &BytesStart,
        is_empty: bool,
    ) -> Result<Self, ParseError> {
        let mut f_id = None;
        let mut f_preset_i_d = None;
        let mut f_preset_class = None;
        let mut f_preset_subtype = None;
        let mut f_dur = None;
        let mut f_repeat_count = None;
        let mut f_repeat_dur = None;
        let mut f_spd = None;
        let mut f_accel = None;
        let mut f_decel = None;
        let mut f_auto_rev = None;
        let mut f_restart = None;
        let mut f_fill = None;
        let mut f_sync_behavior = None;
        let mut f_tm_filter = None;
        let mut f_evt_filter = None;
        let mut f_display = None;
        let mut f_master_rel = None;
        let mut f_bld_lvl = None;
        let mut f_grp_id = None;
        let mut f_after_effect = None;
        let mut f_node_type = None;
        let mut f_node_ph = None;
        let mut f_st_cond_lst = None;
        let mut f_end_cond_lst = None;
        let mut f_end_sync = None;
        let mut f_iterate = None;
        let mut f_child_tn_lst = None;
        let mut f_sub_tn_lst = None;
        #[cfg(feature = "extra-attrs")]
        let mut extra_attrs = std::collections::HashMap::new();
        #[cfg(feature = "extra-children")]
        let mut extra_children = Vec::new();
        #[cfg(feature = "extra-children")]
        let mut child_idx: usize = 0;

        // Parse attributes
        for attr in start_tag.attributes().filter_map(|a| a.ok()) {
            let val = String::from_utf8_lossy(&attr.value);
            match attr.key.local_name().as_ref() {
                b"id" => {
                    f_id = val.parse().ok();
                }
                b"presetID" => {
                    f_preset_i_d = val.parse().ok();
                }
                b"presetClass" => {
                    f_preset_class = val.parse().ok();
                }
                b"presetSubtype" => {
                    f_preset_subtype = val.parse().ok();
                }
                b"dur" => {
                    f_dur = Some(val.into_owned());
                }
                b"repeatCount" => {
                    f_repeat_count = Some(val.into_owned());
                }
                b"repeatDur" => {
                    f_repeat_dur = Some(val.into_owned());
                }
                b"spd" => {
                    f_spd = val.parse().ok();
                }
                b"accel" => {
                    f_accel = val.parse().ok();
                }
                b"decel" => {
                    f_decel = val.parse().ok();
                }
                b"autoRev" => {
                    f_auto_rev = Some(val == "true" || val == "1");
                }
                b"restart" => {
                    f_restart = val.parse().ok();
                }
                b"fill" => {
                    f_fill = val.parse().ok();
                }
                b"syncBehavior" => {
                    f_sync_behavior = val.parse().ok();
                }
                b"tmFilter" => {
                    f_tm_filter = Some(val.into_owned());
                }
                b"evtFilter" => {
                    f_evt_filter = Some(val.into_owned());
                }
                b"display" => {
                    f_display = Some(val == "true" || val == "1");
                }
                b"masterRel" => {
                    f_master_rel = val.parse().ok();
                }
                b"bldLvl" => {
                    f_bld_lvl = val.parse().ok();
                }
                b"grpId" => {
                    f_grp_id = val.parse().ok();
                }
                b"afterEffect" => {
                    f_after_effect = Some(val == "true" || val == "1");
                }
                b"nodeType" => {
                    f_node_type = val.parse().ok();
                }
                b"nodePh" => {
                    f_node_ph = Some(val == "true" || val == "1");
                }
                #[cfg(feature = "extra-attrs")]
                unknown => {
                    let key = String::from_utf8_lossy(attr.key.as_ref()).into_owned();
                    extra_attrs.insert(key, val.into_owned());
                }
                #[cfg(not(feature = "extra-attrs"))]
                _ => {}
            }
        }

        // Parse child elements
        if !is_empty {
            let mut buf = Vec::new();
            loop {
                match reader.read_event_into(&mut buf)? {
                    Event::Start(e) => {
                        match e.local_name().as_ref() {
                            b"stCondLst" => {
                                f_st_cond_lst = Some(Box::new(CTTLTimeConditionList::from_xml(
                                    reader, &e, false,
                                )?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            b"endCondLst" => {
                                f_end_cond_lst = Some(Box::new(CTTLTimeConditionList::from_xml(
                                    reader, &e, false,
                                )?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            b"endSync" => {
                                f_end_sync =
                                    Some(Box::new(CTTLTimeCondition::from_xml(reader, &e, false)?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            b"iterate" => {
                                f_iterate =
                                    Some(Box::new(CTTLIterateData::from_xml(reader, &e, false)?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            b"childTnLst" => {
                                f_child_tn_lst =
                                    Some(Box::new(CTTimeNodeList::from_xml(reader, &e, false)?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            b"subTnLst" => {
                                f_sub_tn_lst =
                                    Some(Box::new(CTTimeNodeList::from_xml(reader, &e, false)?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            #[cfg(feature = "extra-children")]
                            _ => {
                                // Capture unknown element for roundtrip
                                let elem = RawXmlElement::from_reader(reader, &e)?;
                                extra_children.push(PositionedNode::new(
                                    child_idx,
                                    RawXmlNode::Element(elem),
                                ));
                                child_idx += 1;
                            }
                            #[cfg(not(feature = "extra-children"))]
                            _ => {
                                // Skip unknown element
                                skip_element(reader)?;
                            }
                        }
                    }
                    Event::Empty(e) => {
                        match e.local_name().as_ref() {
                            b"stCondLst" => {
                                f_st_cond_lst = Some(Box::new(CTTLTimeConditionList::from_xml(
                                    reader, &e, true,
                                )?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            b"endCondLst" => {
                                f_end_cond_lst = Some(Box::new(CTTLTimeConditionList::from_xml(
                                    reader, &e, true,
                                )?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            b"endSync" => {
                                f_end_sync =
                                    Some(Box::new(CTTLTimeCondition::from_xml(reader, &e, true)?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            b"iterate" => {
                                f_iterate =
                                    Some(Box::new(CTTLIterateData::from_xml(reader, &e, true)?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            b"childTnLst" => {
                                f_child_tn_lst =
                                    Some(Box::new(CTTimeNodeList::from_xml(reader, &e, true)?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            b"subTnLst" => {
                                f_sub_tn_lst =
                                    Some(Box::new(CTTimeNodeList::from_xml(reader, &e, true)?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            #[cfg(feature = "extra-children")]
                            _ => {
                                // Capture unknown empty element for roundtrip
                                let elem = RawXmlElement::from_empty(&e);
                                extra_children.push(PositionedNode::new(
                                    child_idx,
                                    RawXmlNode::Element(elem),
                                ));
                                child_idx += 1;
                            }
                            #[cfg(not(feature = "extra-children"))]
                            _ => {}
                        }
                    }
                    Event::End(_) => break,
                    Event::Eof => break,
                    _ => {}
                }
                buf.clear();
            }
        }

        Ok(Self {
            id: f_id,
            preset_i_d: f_preset_i_d,
            preset_class: f_preset_class,
            preset_subtype: f_preset_subtype,
            dur: f_dur,
            repeat_count: f_repeat_count,
            repeat_dur: f_repeat_dur,
            spd: f_spd,
            accel: f_accel,
            decel: f_decel,
            auto_rev: f_auto_rev,
            restart: f_restart,
            fill: f_fill,
            sync_behavior: f_sync_behavior,
            tm_filter: f_tm_filter,
            evt_filter: f_evt_filter,
            display: f_display,
            master_rel: f_master_rel,
            bld_lvl: f_bld_lvl,
            grp_id: f_grp_id,
            after_effect: f_after_effect,
            node_type: f_node_type,
            node_ph: f_node_ph,
            st_cond_lst: f_st_cond_lst,
            end_cond_lst: f_end_cond_lst,
            end_sync: f_end_sync,
            iterate: f_iterate,
            child_tn_lst: f_child_tn_lst,
            sub_tn_lst: f_sub_tn_lst,
            #[cfg(feature = "extra-attrs")]
            extra_attrs,
            #[cfg(feature = "extra-children")]
            extra_children,
        })
    }
}

impl FromXml for CTTLTimeNodeSequence {
    fn from_xml<R: BufRead>(
        reader: &mut Reader<R>,
        start_tag: &BytesStart,
        is_empty: bool,
    ) -> Result<Self, ParseError> {
        let mut f_concurrent = None;
        let mut f_prev_ac = None;
        let mut f_next_ac = None;
        let mut f_c_tn: Option<Box<CTTLCommonTimeNodeData>> = None;
        let mut f_prev_cond_lst = None;
        let mut f_next_cond_lst = None;
        #[cfg(feature = "extra-attrs")]
        let mut extra_attrs = std::collections::HashMap::new();
        #[cfg(feature = "extra-children")]
        let mut extra_children = Vec::new();
        #[cfg(feature = "extra-children")]
        let mut child_idx: usize = 0;

        // Parse attributes
        for attr in start_tag.attributes().filter_map(|a| a.ok()) {
            let val = String::from_utf8_lossy(&attr.value);
            match attr.key.local_name().as_ref() {
                b"concurrent" => {
                    f_concurrent = Some(val == "true" || val == "1");
                }
                b"prevAc" => {
                    f_prev_ac = val.parse().ok();
                }
                b"nextAc" => {
                    f_next_ac = val.parse().ok();
                }
                #[cfg(feature = "extra-attrs")]
                unknown => {
                    let key = String::from_utf8_lossy(attr.key.as_ref()).into_owned();
                    extra_attrs.insert(key, val.into_owned());
                }
                #[cfg(not(feature = "extra-attrs"))]
                _ => {}
            }
        }

        // Parse child elements
        if !is_empty {
            let mut buf = Vec::new();
            loop {
                match reader.read_event_into(&mut buf)? {
                    Event::Start(e) => {
                        match e.local_name().as_ref() {
                            b"cTn" => {
                                f_c_tn = Some(Box::new(CTTLCommonTimeNodeData::from_xml(
                                    reader, &e, false,
                                )?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            b"prevCondLst" => {
                                f_prev_cond_lst = Some(Box::new(CTTLTimeConditionList::from_xml(
                                    reader, &e, false,
                                )?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            b"nextCondLst" => {
                                f_next_cond_lst = Some(Box::new(CTTLTimeConditionList::from_xml(
                                    reader, &e, false,
                                )?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            #[cfg(feature = "extra-children")]
                            _ => {
                                // Capture unknown element for roundtrip
                                let elem = RawXmlElement::from_reader(reader, &e)?;
                                extra_children.push(PositionedNode::new(
                                    child_idx,
                                    RawXmlNode::Element(elem),
                                ));
                                child_idx += 1;
                            }
                            #[cfg(not(feature = "extra-children"))]
                            _ => {
                                // Skip unknown element
                                skip_element(reader)?;
                            }
                        }
                    }
                    Event::Empty(e) => {
                        match e.local_name().as_ref() {
                            b"cTn" => {
                                f_c_tn = Some(Box::new(CTTLCommonTimeNodeData::from_xml(
                                    reader, &e, true,
                                )?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            b"prevCondLst" => {
                                f_prev_cond_lst = Some(Box::new(CTTLTimeConditionList::from_xml(
                                    reader, &e, true,
                                )?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            b"nextCondLst" => {
                                f_next_cond_lst = Some(Box::new(CTTLTimeConditionList::from_xml(
                                    reader, &e, true,
                                )?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            #[cfg(feature = "extra-children")]
                            _ => {
                                // Capture unknown empty element for roundtrip
                                let elem = RawXmlElement::from_empty(&e);
                                extra_children.push(PositionedNode::new(
                                    child_idx,
                                    RawXmlNode::Element(elem),
                                ));
                                child_idx += 1;
                            }
                            #[cfg(not(feature = "extra-children"))]
                            _ => {}
                        }
                    }
                    Event::End(_) => break,
                    Event::Eof => break,
                    _ => {}
                }
                buf.clear();
            }
        }

        Ok(Self {
            concurrent: f_concurrent,
            prev_ac: f_prev_ac,
            next_ac: f_next_ac,
            c_tn: f_c_tn.ok_or_else(|| ParseError::MissingAttribute("cTn".to_string()))?,
            prev_cond_lst: f_prev_cond_lst,
            next_cond_lst: f_next_cond_lst,
            #[cfg(feature = "extra-attrs")]
            extra_attrs,
            #[cfg(feature = "extra-children")]
            extra_children,
        })
    }
}

impl FromXml for CTTLBehaviorAttributeNameList {
    fn from_xml<R: BufRead>(
        reader: &mut Reader<R>,
        start_tag: &BytesStart,
        is_empty: bool,
    ) -> Result<Self, ParseError> {
        let mut f_attr_name = Vec::new();
        #[cfg(feature = "extra-children")]
        let mut extra_children = Vec::new();
        #[cfg(feature = "extra-children")]
        let mut child_idx: usize = 0;

        // Parse child elements
        if !is_empty {
            let mut buf = Vec::new();
            loop {
                match reader.read_event_into(&mut buf)? {
                    Event::Start(e) => {
                        match e.local_name().as_ref() {
                            b"attrName" => {
                                f_attr_name.push(read_text_content(reader)?);
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            #[cfg(feature = "extra-children")]
                            _ => {
                                // Capture unknown element for roundtrip
                                let elem = RawXmlElement::from_reader(reader, &e)?;
                                extra_children.push(PositionedNode::new(
                                    child_idx,
                                    RawXmlNode::Element(elem),
                                ));
                                child_idx += 1;
                            }
                            #[cfg(not(feature = "extra-children"))]
                            _ => {
                                // Skip unknown element
                                skip_element(reader)?;
                            }
                        }
                    }
                    Event::Empty(e) => {
                        match e.local_name().as_ref() {
                            b"attrName" => {
                                f_attr_name.push(String::new());
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            #[cfg(feature = "extra-children")]
                            _ => {
                                // Capture unknown empty element for roundtrip
                                let elem = RawXmlElement::from_empty(&e);
                                extra_children.push(PositionedNode::new(
                                    child_idx,
                                    RawXmlNode::Element(elem),
                                ));
                                child_idx += 1;
                            }
                            #[cfg(not(feature = "extra-children"))]
                            _ => {}
                        }
                    }
                    Event::End(_) => break,
                    Event::Eof => break,
                    _ => {}
                }
                buf.clear();
            }
        }

        Ok(Self {
            attr_name: f_attr_name,
            #[cfg(feature = "extra-children")]
            extra_children,
        })
    }
}

impl FromXml for CTTLCommonBehaviorData {
    fn from_xml<R: BufRead>(
        reader: &mut Reader<R>,
        start_tag: &BytesStart,
        is_empty: bool,
    ) -> Result<Self, ParseError> {
        let mut f_additive = None;
        let mut f_accumulate = None;
        let mut f_xfrm_type = None;
        let mut f_from = None;
        let mut f_to = None;
        let mut f_by = None;
        let mut f_rctx = None;
        let mut f_override = None;
        let mut f_c_tn: Option<Box<CTTLCommonTimeNodeData>> = None;
        let mut f_tgt_el: Option<Box<CTTLTimeTargetElement>> = None;
        let mut f_attr_name_lst = None;
        #[cfg(feature = "extra-attrs")]
        let mut extra_attrs = std::collections::HashMap::new();
        #[cfg(feature = "extra-children")]
        let mut extra_children = Vec::new();
        #[cfg(feature = "extra-children")]
        let mut child_idx: usize = 0;

        // Parse attributes
        for attr in start_tag.attributes().filter_map(|a| a.ok()) {
            let val = String::from_utf8_lossy(&attr.value);
            match attr.key.local_name().as_ref() {
                b"additive" => {
                    f_additive = val.parse().ok();
                }
                b"accumulate" => {
                    f_accumulate = val.parse().ok();
                }
                b"xfrmType" => {
                    f_xfrm_type = val.parse().ok();
                }
                b"from" => {
                    f_from = Some(val.into_owned());
                }
                b"to" => {
                    f_to = Some(val.into_owned());
                }
                b"by" => {
                    f_by = Some(val.into_owned());
                }
                b"rctx" => {
                    f_rctx = Some(val.into_owned());
                }
                b"override" => {
                    f_override = val.parse().ok();
                }
                #[cfg(feature = "extra-attrs")]
                unknown => {
                    let key = String::from_utf8_lossy(attr.key.as_ref()).into_owned();
                    extra_attrs.insert(key, val.into_owned());
                }
                #[cfg(not(feature = "extra-attrs"))]
                _ => {}
            }
        }

        // Parse child elements
        if !is_empty {
            let mut buf = Vec::new();
            loop {
                match reader.read_event_into(&mut buf)? {
                    Event::Start(e) => {
                        match e.local_name().as_ref() {
                            b"cTn" => {
                                f_c_tn = Some(Box::new(CTTLCommonTimeNodeData::from_xml(
                                    reader, &e, false,
                                )?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            b"tgtEl" => {
                                f_tgt_el = Some(Box::new(CTTLTimeTargetElement::from_xml(
                                    reader, &e, false,
                                )?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            b"attrNameLst" => {
                                f_attr_name_lst = Some(Box::new(
                                    CTTLBehaviorAttributeNameList::from_xml(reader, &e, false)?,
                                ));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            #[cfg(feature = "extra-children")]
                            _ => {
                                // Capture unknown element for roundtrip
                                let elem = RawXmlElement::from_reader(reader, &e)?;
                                extra_children.push(PositionedNode::new(
                                    child_idx,
                                    RawXmlNode::Element(elem),
                                ));
                                child_idx += 1;
                            }
                            #[cfg(not(feature = "extra-children"))]
                            _ => {
                                // Skip unknown element
                                skip_element(reader)?;
                            }
                        }
                    }
                    Event::Empty(e) => {
                        match e.local_name().as_ref() {
                            b"cTn" => {
                                f_c_tn = Some(Box::new(CTTLCommonTimeNodeData::from_xml(
                                    reader, &e, true,
                                )?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            b"tgtEl" => {
                                f_tgt_el = Some(Box::new(CTTLTimeTargetElement::from_xml(
                                    reader, &e, true,
                                )?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            b"attrNameLst" => {
                                f_attr_name_lst = Some(Box::new(
                                    CTTLBehaviorAttributeNameList::from_xml(reader, &e, true)?,
                                ));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            #[cfg(feature = "extra-children")]
                            _ => {
                                // Capture unknown empty element for roundtrip
                                let elem = RawXmlElement::from_empty(&e);
                                extra_children.push(PositionedNode::new(
                                    child_idx,
                                    RawXmlNode::Element(elem),
                                ));
                                child_idx += 1;
                            }
                            #[cfg(not(feature = "extra-children"))]
                            _ => {}
                        }
                    }
                    Event::End(_) => break,
                    Event::Eof => break,
                    _ => {}
                }
                buf.clear();
            }
        }

        Ok(Self {
            additive: f_additive,
            accumulate: f_accumulate,
            xfrm_type: f_xfrm_type,
            from: f_from,
            to: f_to,
            by: f_by,
            rctx: f_rctx,
            r#override: f_override,
            c_tn: f_c_tn.ok_or_else(|| ParseError::MissingAttribute("cTn".to_string()))?,
            tgt_el: f_tgt_el.ok_or_else(|| ParseError::MissingAttribute("tgtEl".to_string()))?,
            attr_name_lst: f_attr_name_lst,
            #[cfg(feature = "extra-attrs")]
            extra_attrs,
            #[cfg(feature = "extra-children")]
            extra_children,
        })
    }
}

impl FromXml for CTTLAnimVariantBooleanVal {
    fn from_xml<R: BufRead>(
        reader: &mut Reader<R>,
        start_tag: &BytesStart,
        is_empty: bool,
    ) -> Result<Self, ParseError> {
        let mut f_value: Option<bool> = None;
        #[cfg(feature = "extra-attrs")]
        let mut extra_attrs = std::collections::HashMap::new();

        // Parse attributes
        for attr in start_tag.attributes().filter_map(|a| a.ok()) {
            let val = String::from_utf8_lossy(&attr.value);
            match attr.key.local_name().as_ref() {
                b"val" => {
                    f_value = Some(val == "true" || val == "1");
                }
                #[cfg(feature = "extra-attrs")]
                unknown => {
                    let key = String::from_utf8_lossy(attr.key.as_ref()).into_owned();
                    extra_attrs.insert(key, val.into_owned());
                }
                #[cfg(not(feature = "extra-attrs"))]
                _ => {}
            }
        }

        if !is_empty {
            let mut buf = Vec::new();
            loop {
                match reader.read_event_into(&mut buf)? {
                    Event::End(_) => break,
                    Event::Eof => break,
                    _ => {}
                }
                buf.clear();
            }
        }

        Ok(Self {
            value: f_value.ok_or_else(|| ParseError::MissingAttribute("val".to_string()))?,
            #[cfg(feature = "extra-attrs")]
            extra_attrs,
        })
    }
}

impl FromXml for CTTLAnimVariantIntegerVal {
    fn from_xml<R: BufRead>(
        reader: &mut Reader<R>,
        start_tag: &BytesStart,
        is_empty: bool,
    ) -> Result<Self, ParseError> {
        let mut f_value: Option<i32> = None;
        #[cfg(feature = "extra-attrs")]
        let mut extra_attrs = std::collections::HashMap::new();

        // Parse attributes
        for attr in start_tag.attributes().filter_map(|a| a.ok()) {
            let val = String::from_utf8_lossy(&attr.value);
            match attr.key.local_name().as_ref() {
                b"val" => {
                    f_value = val.parse().ok();
                }
                #[cfg(feature = "extra-attrs")]
                unknown => {
                    let key = String::from_utf8_lossy(attr.key.as_ref()).into_owned();
                    extra_attrs.insert(key, val.into_owned());
                }
                #[cfg(not(feature = "extra-attrs"))]
                _ => {}
            }
        }

        if !is_empty {
            let mut buf = Vec::new();
            loop {
                match reader.read_event_into(&mut buf)? {
                    Event::End(_) => break,
                    Event::Eof => break,
                    _ => {}
                }
                buf.clear();
            }
        }

        Ok(Self {
            value: f_value.ok_or_else(|| ParseError::MissingAttribute("val".to_string()))?,
            #[cfg(feature = "extra-attrs")]
            extra_attrs,
        })
    }
}

impl FromXml for CTTLAnimVariantFloatVal {
    fn from_xml<R: BufRead>(
        reader: &mut Reader<R>,
        start_tag: &BytesStart,
        is_empty: bool,
    ) -> Result<Self, ParseError> {
        let mut f_value: Option<f32> = None;
        #[cfg(feature = "extra-attrs")]
        let mut extra_attrs = std::collections::HashMap::new();

        // Parse attributes
        for attr in start_tag.attributes().filter_map(|a| a.ok()) {
            let val = String::from_utf8_lossy(&attr.value);
            match attr.key.local_name().as_ref() {
                b"val" => {
                    f_value = val.parse().ok();
                }
                #[cfg(feature = "extra-attrs")]
                unknown => {
                    let key = String::from_utf8_lossy(attr.key.as_ref()).into_owned();
                    extra_attrs.insert(key, val.into_owned());
                }
                #[cfg(not(feature = "extra-attrs"))]
                _ => {}
            }
        }

        if !is_empty {
            let mut buf = Vec::new();
            loop {
                match reader.read_event_into(&mut buf)? {
                    Event::End(_) => break,
                    Event::Eof => break,
                    _ => {}
                }
                buf.clear();
            }
        }

        Ok(Self {
            value: f_value.ok_or_else(|| ParseError::MissingAttribute("val".to_string()))?,
            #[cfg(feature = "extra-attrs")]
            extra_attrs,
        })
    }
}

impl FromXml for CTTLAnimVariantStringVal {
    fn from_xml<R: BufRead>(
        reader: &mut Reader<R>,
        start_tag: &BytesStart,
        is_empty: bool,
    ) -> Result<Self, ParseError> {
        let mut f_value: Option<String> = None;
        #[cfg(feature = "extra-attrs")]
        let mut extra_attrs = std::collections::HashMap::new();

        // Parse attributes
        for attr in start_tag.attributes().filter_map(|a| a.ok()) {
            let val = String::from_utf8_lossy(&attr.value);
            match attr.key.local_name().as_ref() {
                b"val" => {
                    f_value = Some(val.into_owned());
                }
                #[cfg(feature = "extra-attrs")]
                unknown => {
                    let key = String::from_utf8_lossy(attr.key.as_ref()).into_owned();
                    extra_attrs.insert(key, val.into_owned());
                }
                #[cfg(not(feature = "extra-attrs"))]
                _ => {}
            }
        }

        if !is_empty {
            let mut buf = Vec::new();
            loop {
                match reader.read_event_into(&mut buf)? {
                    Event::End(_) => break,
                    Event::Eof => break,
                    _ => {}
                }
                buf.clear();
            }
        }

        Ok(Self {
            value: f_value.ok_or_else(|| ParseError::MissingAttribute("val".to_string()))?,
            #[cfg(feature = "extra-attrs")]
            extra_attrs,
        })
    }
}

impl FromXml for CTTLAnimVariant {
    fn from_xml<R: BufRead>(
        reader: &mut Reader<R>,
        start_tag: &BytesStart,
        is_empty: bool,
    ) -> Result<Self, ParseError> {
        let mut f_bool_val = None;
        let mut f_int_val = None;
        let mut f_flt_val = None;
        let mut f_str_val = None;
        let mut f_clr_val = None;
        #[cfg(feature = "extra-children")]
        let mut extra_children = Vec::new();
        #[cfg(feature = "extra-children")]
        let mut child_idx: usize = 0;

        // Parse child elements
        if !is_empty {
            let mut buf = Vec::new();
            loop {
                match reader.read_event_into(&mut buf)? {
                    Event::Start(e) => {
                        match e.local_name().as_ref() {
                            b"boolVal" => {
                                f_bool_val = Some(Box::new(CTTLAnimVariantBooleanVal::from_xml(
                                    reader, &e, false,
                                )?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            b"intVal" => {
                                f_int_val = Some(Box::new(CTTLAnimVariantIntegerVal::from_xml(
                                    reader, &e, false,
                                )?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            b"fltVal" => {
                                f_flt_val = Some(Box::new(CTTLAnimVariantFloatVal::from_xml(
                                    reader, &e, false,
                                )?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            b"strVal" => {
                                f_str_val = Some(Box::new(CTTLAnimVariantStringVal::from_xml(
                                    reader, &e, false,
                                )?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            b"clrVal" => {
                                f_clr_val = Some(Box::new(ooxml_dml::types::CTColor::from_xml(
                                    reader, &e, false,
                                )?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            #[cfg(feature = "extra-children")]
                            _ => {
                                // Capture unknown element for roundtrip
                                let elem = RawXmlElement::from_reader(reader, &e)?;
                                extra_children.push(PositionedNode::new(
                                    child_idx,
                                    RawXmlNode::Element(elem),
                                ));
                                child_idx += 1;
                            }
                            #[cfg(not(feature = "extra-children"))]
                            _ => {
                                // Skip unknown element
                                skip_element(reader)?;
                            }
                        }
                    }
                    Event::Empty(e) => {
                        match e.local_name().as_ref() {
                            b"boolVal" => {
                                f_bool_val = Some(Box::new(CTTLAnimVariantBooleanVal::from_xml(
                                    reader, &e, true,
                                )?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            b"intVal" => {
                                f_int_val = Some(Box::new(CTTLAnimVariantIntegerVal::from_xml(
                                    reader, &e, true,
                                )?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            b"fltVal" => {
                                f_flt_val = Some(Box::new(CTTLAnimVariantFloatVal::from_xml(
                                    reader, &e, true,
                                )?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            b"strVal" => {
                                f_str_val = Some(Box::new(CTTLAnimVariantStringVal::from_xml(
                                    reader, &e, true,
                                )?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            b"clrVal" => {
                                f_clr_val = Some(Box::new(ooxml_dml::types::CTColor::from_xml(
                                    reader, &e, true,
                                )?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            #[cfg(feature = "extra-children")]
                            _ => {
                                // Capture unknown empty element for roundtrip
                                let elem = RawXmlElement::from_empty(&e);
                                extra_children.push(PositionedNode::new(
                                    child_idx,
                                    RawXmlNode::Element(elem),
                                ));
                                child_idx += 1;
                            }
                            #[cfg(not(feature = "extra-children"))]
                            _ => {}
                        }
                    }
                    Event::End(_) => break,
                    Event::Eof => break,
                    _ => {}
                }
                buf.clear();
            }
        }

        Ok(Self {
            bool_val: f_bool_val,
            int_val: f_int_val,
            flt_val: f_flt_val,
            str_val: f_str_val,
            clr_val: f_clr_val,
            #[cfg(feature = "extra-children")]
            extra_children,
        })
    }
}

impl FromXml for CTTLTimeAnimateValue {
    fn from_xml<R: BufRead>(
        reader: &mut Reader<R>,
        start_tag: &BytesStart,
        is_empty: bool,
    ) -> Result<Self, ParseError> {
        let mut f_tm = None;
        let mut f_fmla = None;
        let mut f_value = None;
        #[cfg(feature = "extra-attrs")]
        let mut extra_attrs = std::collections::HashMap::new();
        #[cfg(feature = "extra-children")]
        let mut extra_children = Vec::new();
        #[cfg(feature = "extra-children")]
        let mut child_idx: usize = 0;

        // Parse attributes
        for attr in start_tag.attributes().filter_map(|a| a.ok()) {
            let val = String::from_utf8_lossy(&attr.value);
            match attr.key.local_name().as_ref() {
                b"tm" => {
                    f_tm = Some(val.into_owned());
                }
                b"fmla" => {
                    f_fmla = Some(val.into_owned());
                }
                #[cfg(feature = "extra-attrs")]
                unknown => {
                    let key = String::from_utf8_lossy(attr.key.as_ref()).into_owned();
                    extra_attrs.insert(key, val.into_owned());
                }
                #[cfg(not(feature = "extra-attrs"))]
                _ => {}
            }
        }

        // Parse child elements
        if !is_empty {
            let mut buf = Vec::new();
            loop {
                match reader.read_event_into(&mut buf)? {
                    Event::Start(e) => {
                        match e.local_name().as_ref() {
                            b"val" => {
                                f_value =
                                    Some(Box::new(CTTLAnimVariant::from_xml(reader, &e, false)?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            #[cfg(feature = "extra-children")]
                            _ => {
                                // Capture unknown element for roundtrip
                                let elem = RawXmlElement::from_reader(reader, &e)?;
                                extra_children.push(PositionedNode::new(
                                    child_idx,
                                    RawXmlNode::Element(elem),
                                ));
                                child_idx += 1;
                            }
                            #[cfg(not(feature = "extra-children"))]
                            _ => {
                                // Skip unknown element
                                skip_element(reader)?;
                            }
                        }
                    }
                    Event::Empty(e) => {
                        match e.local_name().as_ref() {
                            b"val" => {
                                f_value =
                                    Some(Box::new(CTTLAnimVariant::from_xml(reader, &e, true)?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            #[cfg(feature = "extra-children")]
                            _ => {
                                // Capture unknown empty element for roundtrip
                                let elem = RawXmlElement::from_empty(&e);
                                extra_children.push(PositionedNode::new(
                                    child_idx,
                                    RawXmlNode::Element(elem),
                                ));
                                child_idx += 1;
                            }
                            #[cfg(not(feature = "extra-children"))]
                            _ => {}
                        }
                    }
                    Event::End(_) => break,
                    Event::Eof => break,
                    _ => {}
                }
                buf.clear();
            }
        }

        Ok(Self {
            tm: f_tm,
            fmla: f_fmla,
            value: f_value,
            #[cfg(feature = "extra-attrs")]
            extra_attrs,
            #[cfg(feature = "extra-children")]
            extra_children,
        })
    }
}

impl FromXml for CTTLTimeAnimateValueList {
    fn from_xml<R: BufRead>(
        reader: &mut Reader<R>,
        start_tag: &BytesStart,
        is_empty: bool,
    ) -> Result<Self, ParseError> {
        let mut f_tav = Vec::new();
        #[cfg(feature = "extra-children")]
        let mut extra_children = Vec::new();
        #[cfg(feature = "extra-children")]
        let mut child_idx: usize = 0;

        // Parse child elements
        if !is_empty {
            let mut buf = Vec::new();
            loop {
                match reader.read_event_into(&mut buf)? {
                    Event::Start(e) => {
                        match e.local_name().as_ref() {
                            b"tav" => {
                                f_tav.push(CTTLTimeAnimateValue::from_xml(reader, &e, false)?);
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            #[cfg(feature = "extra-children")]
                            _ => {
                                // Capture unknown element for roundtrip
                                let elem = RawXmlElement::from_reader(reader, &e)?;
                                extra_children.push(PositionedNode::new(
                                    child_idx,
                                    RawXmlNode::Element(elem),
                                ));
                                child_idx += 1;
                            }
                            #[cfg(not(feature = "extra-children"))]
                            _ => {
                                // Skip unknown element
                                skip_element(reader)?;
                            }
                        }
                    }
                    Event::Empty(e) => {
                        match e.local_name().as_ref() {
                            b"tav" => {
                                f_tav.push(CTTLTimeAnimateValue::from_xml(reader, &e, true)?);
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            #[cfg(feature = "extra-children")]
                            _ => {
                                // Capture unknown empty element for roundtrip
                                let elem = RawXmlElement::from_empty(&e);
                                extra_children.push(PositionedNode::new(
                                    child_idx,
                                    RawXmlNode::Element(elem),
                                ));
                                child_idx += 1;
                            }
                            #[cfg(not(feature = "extra-children"))]
                            _ => {}
                        }
                    }
                    Event::End(_) => break,
                    Event::Eof => break,
                    _ => {}
                }
                buf.clear();
            }
        }

        Ok(Self {
            tav: f_tav,
            #[cfg(feature = "extra-children")]
            extra_children,
        })
    }
}

impl FromXml for CTTLAnimateBehavior {
    fn from_xml<R: BufRead>(
        reader: &mut Reader<R>,
        start_tag: &BytesStart,
        is_empty: bool,
    ) -> Result<Self, ParseError> {
        let mut f_by = None;
        let mut f_from = None;
        let mut f_to = None;
        let mut f_calcmode = None;
        let mut f_value_type = None;
        let mut f_c_bhvr: Option<Box<CTTLCommonBehaviorData>> = None;
        let mut f_tav_lst = None;
        #[cfg(feature = "extra-attrs")]
        let mut extra_attrs = std::collections::HashMap::new();
        #[cfg(feature = "extra-children")]
        let mut extra_children = Vec::new();
        #[cfg(feature = "extra-children")]
        let mut child_idx: usize = 0;

        // Parse attributes
        for attr in start_tag.attributes().filter_map(|a| a.ok()) {
            let val = String::from_utf8_lossy(&attr.value);
            match attr.key.local_name().as_ref() {
                b"by" => {
                    f_by = Some(val.into_owned());
                }
                b"from" => {
                    f_from = Some(val.into_owned());
                }
                b"to" => {
                    f_to = Some(val.into_owned());
                }
                b"calcmode" => {
                    f_calcmode = val.parse().ok();
                }
                b"valueType" => {
                    f_value_type = val.parse().ok();
                }
                #[cfg(feature = "extra-attrs")]
                unknown => {
                    let key = String::from_utf8_lossy(attr.key.as_ref()).into_owned();
                    extra_attrs.insert(key, val.into_owned());
                }
                #[cfg(not(feature = "extra-attrs"))]
                _ => {}
            }
        }

        // Parse child elements
        if !is_empty {
            let mut buf = Vec::new();
            loop {
                match reader.read_event_into(&mut buf)? {
                    Event::Start(e) => {
                        match e.local_name().as_ref() {
                            b"cBhvr" => {
                                f_c_bhvr = Some(Box::new(CTTLCommonBehaviorData::from_xml(
                                    reader, &e, false,
                                )?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            b"tavLst" => {
                                f_tav_lst = Some(Box::new(CTTLTimeAnimateValueList::from_xml(
                                    reader, &e, false,
                                )?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            #[cfg(feature = "extra-children")]
                            _ => {
                                // Capture unknown element for roundtrip
                                let elem = RawXmlElement::from_reader(reader, &e)?;
                                extra_children.push(PositionedNode::new(
                                    child_idx,
                                    RawXmlNode::Element(elem),
                                ));
                                child_idx += 1;
                            }
                            #[cfg(not(feature = "extra-children"))]
                            _ => {
                                // Skip unknown element
                                skip_element(reader)?;
                            }
                        }
                    }
                    Event::Empty(e) => {
                        match e.local_name().as_ref() {
                            b"cBhvr" => {
                                f_c_bhvr = Some(Box::new(CTTLCommonBehaviorData::from_xml(
                                    reader, &e, true,
                                )?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            b"tavLst" => {
                                f_tav_lst = Some(Box::new(CTTLTimeAnimateValueList::from_xml(
                                    reader, &e, true,
                                )?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            #[cfg(feature = "extra-children")]
                            _ => {
                                // Capture unknown empty element for roundtrip
                                let elem = RawXmlElement::from_empty(&e);
                                extra_children.push(PositionedNode::new(
                                    child_idx,
                                    RawXmlNode::Element(elem),
                                ));
                                child_idx += 1;
                            }
                            #[cfg(not(feature = "extra-children"))]
                            _ => {}
                        }
                    }
                    Event::End(_) => break,
                    Event::Eof => break,
                    _ => {}
                }
                buf.clear();
            }
        }

        Ok(Self {
            by: f_by,
            from: f_from,
            to: f_to,
            calcmode: f_calcmode,
            value_type: f_value_type,
            c_bhvr: f_c_bhvr.ok_or_else(|| ParseError::MissingAttribute("cBhvr".to_string()))?,
            tav_lst: f_tav_lst,
            #[cfg(feature = "extra-attrs")]
            extra_attrs,
            #[cfg(feature = "extra-children")]
            extra_children,
        })
    }
}

impl FromXml for CTTLByRgbColorTransform {
    fn from_xml<R: BufRead>(
        reader: &mut Reader<R>,
        start_tag: &BytesStart,
        is_empty: bool,
    ) -> Result<Self, ParseError> {
        let mut f_reference: Option<ooxml_dml::types::STFixedPercentage> = None;
        let mut f_g: Option<ooxml_dml::types::STFixedPercentage> = None;
        let mut f_b: Option<ooxml_dml::types::STFixedPercentage> = None;
        #[cfg(feature = "extra-attrs")]
        let mut extra_attrs = std::collections::HashMap::new();

        // Parse attributes
        for attr in start_tag.attributes().filter_map(|a| a.ok()) {
            let val = String::from_utf8_lossy(&attr.value);
            match attr.key.local_name().as_ref() {
                b"r" => {
                    f_reference = val.parse().ok();
                }
                b"g" => {
                    f_g = val.parse().ok();
                }
                b"b" => {
                    f_b = val.parse().ok();
                }
                #[cfg(feature = "extra-attrs")]
                unknown => {
                    let key = String::from_utf8_lossy(attr.key.as_ref()).into_owned();
                    extra_attrs.insert(key, val.into_owned());
                }
                #[cfg(not(feature = "extra-attrs"))]
                _ => {}
            }
        }

        if !is_empty {
            let mut buf = Vec::new();
            loop {
                match reader.read_event_into(&mut buf)? {
                    Event::End(_) => break,
                    Event::Eof => break,
                    _ => {}
                }
                buf.clear();
            }
        }

        Ok(Self {
            reference: f_reference.ok_or_else(|| ParseError::MissingAttribute("r".to_string()))?,
            g: f_g.ok_or_else(|| ParseError::MissingAttribute("g".to_string()))?,
            b: f_b.ok_or_else(|| ParseError::MissingAttribute("b".to_string()))?,
            #[cfg(feature = "extra-attrs")]
            extra_attrs,
        })
    }
}

impl FromXml for CTTLByHslColorTransform {
    fn from_xml<R: BufRead>(
        reader: &mut Reader<R>,
        start_tag: &BytesStart,
        is_empty: bool,
    ) -> Result<Self, ParseError> {
        let mut f_height: Option<ooxml_dml::types::STAngle> = None;
        let mut f_s: Option<ooxml_dml::types::STFixedPercentage> = None;
        let mut f_l: Option<ooxml_dml::types::STFixedPercentage> = None;
        #[cfg(feature = "extra-attrs")]
        let mut extra_attrs = std::collections::HashMap::new();

        // Parse attributes
        for attr in start_tag.attributes().filter_map(|a| a.ok()) {
            let val = String::from_utf8_lossy(&attr.value);
            match attr.key.local_name().as_ref() {
                b"h" => {
                    f_height = val.parse().ok();
                }
                b"s" => {
                    f_s = val.parse().ok();
                }
                b"l" => {
                    f_l = val.parse().ok();
                }
                #[cfg(feature = "extra-attrs")]
                unknown => {
                    let key = String::from_utf8_lossy(attr.key.as_ref()).into_owned();
                    extra_attrs.insert(key, val.into_owned());
                }
                #[cfg(not(feature = "extra-attrs"))]
                _ => {}
            }
        }

        if !is_empty {
            let mut buf = Vec::new();
            loop {
                match reader.read_event_into(&mut buf)? {
                    Event::End(_) => break,
                    Event::Eof => break,
                    _ => {}
                }
                buf.clear();
            }
        }

        Ok(Self {
            height: f_height.ok_or_else(|| ParseError::MissingAttribute("h".to_string()))?,
            s: f_s.ok_or_else(|| ParseError::MissingAttribute("s".to_string()))?,
            l: f_l.ok_or_else(|| ParseError::MissingAttribute("l".to_string()))?,
            #[cfg(feature = "extra-attrs")]
            extra_attrs,
        })
    }
}

impl FromXml for CTTLByAnimateColorTransform {
    fn from_xml<R: BufRead>(
        reader: &mut Reader<R>,
        start_tag: &BytesStart,
        is_empty: bool,
    ) -> Result<Self, ParseError> {
        let mut f_rgb = None;
        let mut f_hsl = None;
        #[cfg(feature = "extra-children")]
        let mut extra_children = Vec::new();
        #[cfg(feature = "extra-children")]
        let mut child_idx: usize = 0;

        // Parse child elements
        if !is_empty {
            let mut buf = Vec::new();
            loop {
                match reader.read_event_into(&mut buf)? {
                    Event::Start(e) => {
                        match e.local_name().as_ref() {
                            b"rgb" => {
                                f_rgb = Some(Box::new(CTTLByRgbColorTransform::from_xml(
                                    reader, &e, false,
                                )?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            b"hsl" => {
                                f_hsl = Some(Box::new(CTTLByHslColorTransform::from_xml(
                                    reader, &e, false,
                                )?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            #[cfg(feature = "extra-children")]
                            _ => {
                                // Capture unknown element for roundtrip
                                let elem = RawXmlElement::from_reader(reader, &e)?;
                                extra_children.push(PositionedNode::new(
                                    child_idx,
                                    RawXmlNode::Element(elem),
                                ));
                                child_idx += 1;
                            }
                            #[cfg(not(feature = "extra-children"))]
                            _ => {
                                // Skip unknown element
                                skip_element(reader)?;
                            }
                        }
                    }
                    Event::Empty(e) => {
                        match e.local_name().as_ref() {
                            b"rgb" => {
                                f_rgb = Some(Box::new(CTTLByRgbColorTransform::from_xml(
                                    reader, &e, true,
                                )?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            b"hsl" => {
                                f_hsl = Some(Box::new(CTTLByHslColorTransform::from_xml(
                                    reader, &e, true,
                                )?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            #[cfg(feature = "extra-children")]
                            _ => {
                                // Capture unknown empty element for roundtrip
                                let elem = RawXmlElement::from_empty(&e);
                                extra_children.push(PositionedNode::new(
                                    child_idx,
                                    RawXmlNode::Element(elem),
                                ));
                                child_idx += 1;
                            }
                            #[cfg(not(feature = "extra-children"))]
                            _ => {}
                        }
                    }
                    Event::End(_) => break,
                    Event::Eof => break,
                    _ => {}
                }
                buf.clear();
            }
        }

        Ok(Self {
            rgb: f_rgb,
            hsl: f_hsl,
            #[cfg(feature = "extra-children")]
            extra_children,
        })
    }
}

impl FromXml for CTTLAnimateColorBehavior {
    fn from_xml<R: BufRead>(
        reader: &mut Reader<R>,
        start_tag: &BytesStart,
        is_empty: bool,
    ) -> Result<Self, ParseError> {
        let mut f_clr_spc = None;
        let mut f_dir = None;
        let mut f_c_bhvr: Option<Box<CTTLCommonBehaviorData>> = None;
        let mut f_by = None;
        let mut f_from = None;
        let mut f_to = None;
        #[cfg(feature = "extra-attrs")]
        let mut extra_attrs = std::collections::HashMap::new();
        #[cfg(feature = "extra-children")]
        let mut extra_children = Vec::new();
        #[cfg(feature = "extra-children")]
        let mut child_idx: usize = 0;

        // Parse attributes
        for attr in start_tag.attributes().filter_map(|a| a.ok()) {
            let val = String::from_utf8_lossy(&attr.value);
            match attr.key.local_name().as_ref() {
                b"clrSpc" => {
                    f_clr_spc = val.parse().ok();
                }
                b"dir" => {
                    f_dir = val.parse().ok();
                }
                #[cfg(feature = "extra-attrs")]
                unknown => {
                    let key = String::from_utf8_lossy(attr.key.as_ref()).into_owned();
                    extra_attrs.insert(key, val.into_owned());
                }
                #[cfg(not(feature = "extra-attrs"))]
                _ => {}
            }
        }

        // Parse child elements
        if !is_empty {
            let mut buf = Vec::new();
            loop {
                match reader.read_event_into(&mut buf)? {
                    Event::Start(e) => {
                        match e.local_name().as_ref() {
                            b"cBhvr" => {
                                f_c_bhvr = Some(Box::new(CTTLCommonBehaviorData::from_xml(
                                    reader, &e, false,
                                )?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            b"by" => {
                                f_by = Some(Box::new(CTTLByAnimateColorTransform::from_xml(
                                    reader, &e, false,
                                )?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            b"from" => {
                                f_from = Some(Box::new(ooxml_dml::types::CTColor::from_xml(
                                    reader, &e, false,
                                )?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            b"to" => {
                                f_to = Some(Box::new(ooxml_dml::types::CTColor::from_xml(
                                    reader, &e, false,
                                )?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            #[cfg(feature = "extra-children")]
                            _ => {
                                // Capture unknown element for roundtrip
                                let elem = RawXmlElement::from_reader(reader, &e)?;
                                extra_children.push(PositionedNode::new(
                                    child_idx,
                                    RawXmlNode::Element(elem),
                                ));
                                child_idx += 1;
                            }
                            #[cfg(not(feature = "extra-children"))]
                            _ => {
                                // Skip unknown element
                                skip_element(reader)?;
                            }
                        }
                    }
                    Event::Empty(e) => {
                        match e.local_name().as_ref() {
                            b"cBhvr" => {
                                f_c_bhvr = Some(Box::new(CTTLCommonBehaviorData::from_xml(
                                    reader, &e, true,
                                )?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            b"by" => {
                                f_by = Some(Box::new(CTTLByAnimateColorTransform::from_xml(
                                    reader, &e, true,
                                )?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            b"from" => {
                                f_from = Some(Box::new(ooxml_dml::types::CTColor::from_xml(
                                    reader, &e, true,
                                )?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            b"to" => {
                                f_to = Some(Box::new(ooxml_dml::types::CTColor::from_xml(
                                    reader, &e, true,
                                )?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            #[cfg(feature = "extra-children")]
                            _ => {
                                // Capture unknown empty element for roundtrip
                                let elem = RawXmlElement::from_empty(&e);
                                extra_children.push(PositionedNode::new(
                                    child_idx,
                                    RawXmlNode::Element(elem),
                                ));
                                child_idx += 1;
                            }
                            #[cfg(not(feature = "extra-children"))]
                            _ => {}
                        }
                    }
                    Event::End(_) => break,
                    Event::Eof => break,
                    _ => {}
                }
                buf.clear();
            }
        }

        Ok(Self {
            clr_spc: f_clr_spc,
            dir: f_dir,
            c_bhvr: f_c_bhvr.ok_or_else(|| ParseError::MissingAttribute("cBhvr".to_string()))?,
            by: f_by,
            from: f_from,
            to: f_to,
            #[cfg(feature = "extra-attrs")]
            extra_attrs,
            #[cfg(feature = "extra-children")]
            extra_children,
        })
    }
}

impl FromXml for CTTLAnimateEffectBehavior {
    fn from_xml<R: BufRead>(
        reader: &mut Reader<R>,
        start_tag: &BytesStart,
        is_empty: bool,
    ) -> Result<Self, ParseError> {
        let mut f_transition = None;
        let mut f_filter = None;
        let mut f_pr_lst = None;
        let mut f_c_bhvr: Option<Box<CTTLCommonBehaviorData>> = None;
        let mut f_progress = None;
        #[cfg(feature = "extra-attrs")]
        let mut extra_attrs = std::collections::HashMap::new();
        #[cfg(feature = "extra-children")]
        let mut extra_children = Vec::new();
        #[cfg(feature = "extra-children")]
        let mut child_idx: usize = 0;

        // Parse attributes
        for attr in start_tag.attributes().filter_map(|a| a.ok()) {
            let val = String::from_utf8_lossy(&attr.value);
            match attr.key.local_name().as_ref() {
                b"transition" => {
                    f_transition = val.parse().ok();
                }
                b"filter" => {
                    f_filter = Some(val.into_owned());
                }
                b"prLst" => {
                    f_pr_lst = Some(val.into_owned());
                }
                #[cfg(feature = "extra-attrs")]
                unknown => {
                    let key = String::from_utf8_lossy(attr.key.as_ref()).into_owned();
                    extra_attrs.insert(key, val.into_owned());
                }
                #[cfg(not(feature = "extra-attrs"))]
                _ => {}
            }
        }

        // Parse child elements
        if !is_empty {
            let mut buf = Vec::new();
            loop {
                match reader.read_event_into(&mut buf)? {
                    Event::Start(e) => {
                        match e.local_name().as_ref() {
                            b"cBhvr" => {
                                f_c_bhvr = Some(Box::new(CTTLCommonBehaviorData::from_xml(
                                    reader, &e, false,
                                )?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            b"progress" => {
                                f_progress =
                                    Some(Box::new(CTTLAnimVariant::from_xml(reader, &e, false)?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            #[cfg(feature = "extra-children")]
                            _ => {
                                // Capture unknown element for roundtrip
                                let elem = RawXmlElement::from_reader(reader, &e)?;
                                extra_children.push(PositionedNode::new(
                                    child_idx,
                                    RawXmlNode::Element(elem),
                                ));
                                child_idx += 1;
                            }
                            #[cfg(not(feature = "extra-children"))]
                            _ => {
                                // Skip unknown element
                                skip_element(reader)?;
                            }
                        }
                    }
                    Event::Empty(e) => {
                        match e.local_name().as_ref() {
                            b"cBhvr" => {
                                f_c_bhvr = Some(Box::new(CTTLCommonBehaviorData::from_xml(
                                    reader, &e, true,
                                )?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            b"progress" => {
                                f_progress =
                                    Some(Box::new(CTTLAnimVariant::from_xml(reader, &e, true)?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            #[cfg(feature = "extra-children")]
                            _ => {
                                // Capture unknown empty element for roundtrip
                                let elem = RawXmlElement::from_empty(&e);
                                extra_children.push(PositionedNode::new(
                                    child_idx,
                                    RawXmlNode::Element(elem),
                                ));
                                child_idx += 1;
                            }
                            #[cfg(not(feature = "extra-children"))]
                            _ => {}
                        }
                    }
                    Event::End(_) => break,
                    Event::Eof => break,
                    _ => {}
                }
                buf.clear();
            }
        }

        Ok(Self {
            transition: f_transition,
            filter: f_filter,
            pr_lst: f_pr_lst,
            c_bhvr: f_c_bhvr.ok_or_else(|| ParseError::MissingAttribute("cBhvr".to_string()))?,
            progress: f_progress,
            #[cfg(feature = "extra-attrs")]
            extra_attrs,
            #[cfg(feature = "extra-children")]
            extra_children,
        })
    }
}

impl FromXml for CTTLPoint {
    fn from_xml<R: BufRead>(
        reader: &mut Reader<R>,
        start_tag: &BytesStart,
        is_empty: bool,
    ) -> Result<Self, ParseError> {
        let mut f_x: Option<ooxml_dml::types::STPercentage> = None;
        let mut f_y: Option<ooxml_dml::types::STPercentage> = None;
        #[cfg(feature = "extra-attrs")]
        let mut extra_attrs = std::collections::HashMap::new();

        // Parse attributes
        for attr in start_tag.attributes().filter_map(|a| a.ok()) {
            let val = String::from_utf8_lossy(&attr.value);
            match attr.key.local_name().as_ref() {
                b"x" => {
                    f_x = val.parse().ok();
                }
                b"y" => {
                    f_y = val.parse().ok();
                }
                #[cfg(feature = "extra-attrs")]
                unknown => {
                    let key = String::from_utf8_lossy(attr.key.as_ref()).into_owned();
                    extra_attrs.insert(key, val.into_owned());
                }
                #[cfg(not(feature = "extra-attrs"))]
                _ => {}
            }
        }

        if !is_empty {
            let mut buf = Vec::new();
            loop {
                match reader.read_event_into(&mut buf)? {
                    Event::End(_) => break,
                    Event::Eof => break,
                    _ => {}
                }
                buf.clear();
            }
        }

        Ok(Self {
            x: f_x.ok_or_else(|| ParseError::MissingAttribute("x".to_string()))?,
            y: f_y.ok_or_else(|| ParseError::MissingAttribute("y".to_string()))?,
            #[cfg(feature = "extra-attrs")]
            extra_attrs,
        })
    }
}

impl FromXml for CTTLAnimateMotionBehavior {
    fn from_xml<R: BufRead>(
        reader: &mut Reader<R>,
        start_tag: &BytesStart,
        is_empty: bool,
    ) -> Result<Self, ParseError> {
        let mut f_origin = None;
        let mut f_path = None;
        let mut f_path_edit_mode = None;
        let mut f_r_ang = None;
        let mut f_pts_types = None;
        let mut f_c_bhvr: Option<Box<CTTLCommonBehaviorData>> = None;
        let mut f_by = None;
        let mut f_from = None;
        let mut f_to = None;
        let mut f_r_ctr = None;
        #[cfg(feature = "extra-attrs")]
        let mut extra_attrs = std::collections::HashMap::new();
        #[cfg(feature = "extra-children")]
        let mut extra_children = Vec::new();
        #[cfg(feature = "extra-children")]
        let mut child_idx: usize = 0;

        // Parse attributes
        for attr in start_tag.attributes().filter_map(|a| a.ok()) {
            let val = String::from_utf8_lossy(&attr.value);
            match attr.key.local_name().as_ref() {
                b"origin" => {
                    f_origin = val.parse().ok();
                }
                b"path" => {
                    f_path = Some(val.into_owned());
                }
                b"pathEditMode" => {
                    f_path_edit_mode = val.parse().ok();
                }
                b"rAng" => {
                    f_r_ang = val.parse().ok();
                }
                b"ptsTypes" => {
                    f_pts_types = Some(val.into_owned());
                }
                #[cfg(feature = "extra-attrs")]
                unknown => {
                    let key = String::from_utf8_lossy(attr.key.as_ref()).into_owned();
                    extra_attrs.insert(key, val.into_owned());
                }
                #[cfg(not(feature = "extra-attrs"))]
                _ => {}
            }
        }

        // Parse child elements
        if !is_empty {
            let mut buf = Vec::new();
            loop {
                match reader.read_event_into(&mut buf)? {
                    Event::Start(e) => {
                        match e.local_name().as_ref() {
                            b"cBhvr" => {
                                f_c_bhvr = Some(Box::new(CTTLCommonBehaviorData::from_xml(
                                    reader, &e, false,
                                )?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            b"by" => {
                                f_by = Some(Box::new(CTTLPoint::from_xml(reader, &e, false)?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            b"from" => {
                                f_from = Some(Box::new(CTTLPoint::from_xml(reader, &e, false)?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            b"to" => {
                                f_to = Some(Box::new(CTTLPoint::from_xml(reader, &e, false)?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            b"rCtr" => {
                                f_r_ctr = Some(Box::new(CTTLPoint::from_xml(reader, &e, false)?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            #[cfg(feature = "extra-children")]
                            _ => {
                                // Capture unknown element for roundtrip
                                let elem = RawXmlElement::from_reader(reader, &e)?;
                                extra_children.push(PositionedNode::new(
                                    child_idx,
                                    RawXmlNode::Element(elem),
                                ));
                                child_idx += 1;
                            }
                            #[cfg(not(feature = "extra-children"))]
                            _ => {
                                // Skip unknown element
                                skip_element(reader)?;
                            }
                        }
                    }
                    Event::Empty(e) => {
                        match e.local_name().as_ref() {
                            b"cBhvr" => {
                                f_c_bhvr = Some(Box::new(CTTLCommonBehaviorData::from_xml(
                                    reader, &e, true,
                                )?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            b"by" => {
                                f_by = Some(Box::new(CTTLPoint::from_xml(reader, &e, true)?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            b"from" => {
                                f_from = Some(Box::new(CTTLPoint::from_xml(reader, &e, true)?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            b"to" => {
                                f_to = Some(Box::new(CTTLPoint::from_xml(reader, &e, true)?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            b"rCtr" => {
                                f_r_ctr = Some(Box::new(CTTLPoint::from_xml(reader, &e, true)?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            #[cfg(feature = "extra-children")]
                            _ => {
                                // Capture unknown empty element for roundtrip
                                let elem = RawXmlElement::from_empty(&e);
                                extra_children.push(PositionedNode::new(
                                    child_idx,
                                    RawXmlNode::Element(elem),
                                ));
                                child_idx += 1;
                            }
                            #[cfg(not(feature = "extra-children"))]
                            _ => {}
                        }
                    }
                    Event::End(_) => break,
                    Event::Eof => break,
                    _ => {}
                }
                buf.clear();
            }
        }

        Ok(Self {
            origin: f_origin,
            path: f_path,
            path_edit_mode: f_path_edit_mode,
            r_ang: f_r_ang,
            pts_types: f_pts_types,
            c_bhvr: f_c_bhvr.ok_or_else(|| ParseError::MissingAttribute("cBhvr".to_string()))?,
            by: f_by,
            from: f_from,
            to: f_to,
            r_ctr: f_r_ctr,
            #[cfg(feature = "extra-attrs")]
            extra_attrs,
            #[cfg(feature = "extra-children")]
            extra_children,
        })
    }
}

impl FromXml for CTTLAnimateRotationBehavior {
    fn from_xml<R: BufRead>(
        reader: &mut Reader<R>,
        start_tag: &BytesStart,
        is_empty: bool,
    ) -> Result<Self, ParseError> {
        let mut f_by = None;
        let mut f_from = None;
        let mut f_to = None;
        let mut f_c_bhvr: Option<Box<CTTLCommonBehaviorData>> = None;
        #[cfg(feature = "extra-attrs")]
        let mut extra_attrs = std::collections::HashMap::new();
        #[cfg(feature = "extra-children")]
        let mut extra_children = Vec::new();
        #[cfg(feature = "extra-children")]
        let mut child_idx: usize = 0;

        // Parse attributes
        for attr in start_tag.attributes().filter_map(|a| a.ok()) {
            let val = String::from_utf8_lossy(&attr.value);
            match attr.key.local_name().as_ref() {
                b"by" => {
                    f_by = val.parse().ok();
                }
                b"from" => {
                    f_from = val.parse().ok();
                }
                b"to" => {
                    f_to = val.parse().ok();
                }
                #[cfg(feature = "extra-attrs")]
                unknown => {
                    let key = String::from_utf8_lossy(attr.key.as_ref()).into_owned();
                    extra_attrs.insert(key, val.into_owned());
                }
                #[cfg(not(feature = "extra-attrs"))]
                _ => {}
            }
        }

        // Parse child elements
        if !is_empty {
            let mut buf = Vec::new();
            loop {
                match reader.read_event_into(&mut buf)? {
                    Event::Start(e) => {
                        match e.local_name().as_ref() {
                            b"cBhvr" => {
                                f_c_bhvr = Some(Box::new(CTTLCommonBehaviorData::from_xml(
                                    reader, &e, false,
                                )?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            #[cfg(feature = "extra-children")]
                            _ => {
                                // Capture unknown element for roundtrip
                                let elem = RawXmlElement::from_reader(reader, &e)?;
                                extra_children.push(PositionedNode::new(
                                    child_idx,
                                    RawXmlNode::Element(elem),
                                ));
                                child_idx += 1;
                            }
                            #[cfg(not(feature = "extra-children"))]
                            _ => {
                                // Skip unknown element
                                skip_element(reader)?;
                            }
                        }
                    }
                    Event::Empty(e) => {
                        match e.local_name().as_ref() {
                            b"cBhvr" => {
                                f_c_bhvr = Some(Box::new(CTTLCommonBehaviorData::from_xml(
                                    reader, &e, true,
                                )?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            #[cfg(feature = "extra-children")]
                            _ => {
                                // Capture unknown empty element for roundtrip
                                let elem = RawXmlElement::from_empty(&e);
                                extra_children.push(PositionedNode::new(
                                    child_idx,
                                    RawXmlNode::Element(elem),
                                ));
                                child_idx += 1;
                            }
                            #[cfg(not(feature = "extra-children"))]
                            _ => {}
                        }
                    }
                    Event::End(_) => break,
                    Event::Eof => break,
                    _ => {}
                }
                buf.clear();
            }
        }

        Ok(Self {
            by: f_by,
            from: f_from,
            to: f_to,
            c_bhvr: f_c_bhvr.ok_or_else(|| ParseError::MissingAttribute("cBhvr".to_string()))?,
            #[cfg(feature = "extra-attrs")]
            extra_attrs,
            #[cfg(feature = "extra-children")]
            extra_children,
        })
    }
}

impl FromXml for CTTLAnimateScaleBehavior {
    fn from_xml<R: BufRead>(
        reader: &mut Reader<R>,
        start_tag: &BytesStart,
        is_empty: bool,
    ) -> Result<Self, ParseError> {
        let mut f_zoom_contents = None;
        let mut f_c_bhvr: Option<Box<CTTLCommonBehaviorData>> = None;
        let mut f_by = None;
        let mut f_from = None;
        let mut f_to = None;
        #[cfg(feature = "extra-attrs")]
        let mut extra_attrs = std::collections::HashMap::new();
        #[cfg(feature = "extra-children")]
        let mut extra_children = Vec::new();
        #[cfg(feature = "extra-children")]
        let mut child_idx: usize = 0;

        // Parse attributes
        for attr in start_tag.attributes().filter_map(|a| a.ok()) {
            let val = String::from_utf8_lossy(&attr.value);
            match attr.key.local_name().as_ref() {
                b"zoomContents" => {
                    f_zoom_contents = Some(val == "true" || val == "1");
                }
                #[cfg(feature = "extra-attrs")]
                unknown => {
                    let key = String::from_utf8_lossy(attr.key.as_ref()).into_owned();
                    extra_attrs.insert(key, val.into_owned());
                }
                #[cfg(not(feature = "extra-attrs"))]
                _ => {}
            }
        }

        // Parse child elements
        if !is_empty {
            let mut buf = Vec::new();
            loop {
                match reader.read_event_into(&mut buf)? {
                    Event::Start(e) => {
                        match e.local_name().as_ref() {
                            b"cBhvr" => {
                                f_c_bhvr = Some(Box::new(CTTLCommonBehaviorData::from_xml(
                                    reader, &e, false,
                                )?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            b"by" => {
                                f_by = Some(Box::new(CTTLPoint::from_xml(reader, &e, false)?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            b"from" => {
                                f_from = Some(Box::new(CTTLPoint::from_xml(reader, &e, false)?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            b"to" => {
                                f_to = Some(Box::new(CTTLPoint::from_xml(reader, &e, false)?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            #[cfg(feature = "extra-children")]
                            _ => {
                                // Capture unknown element for roundtrip
                                let elem = RawXmlElement::from_reader(reader, &e)?;
                                extra_children.push(PositionedNode::new(
                                    child_idx,
                                    RawXmlNode::Element(elem),
                                ));
                                child_idx += 1;
                            }
                            #[cfg(not(feature = "extra-children"))]
                            _ => {
                                // Skip unknown element
                                skip_element(reader)?;
                            }
                        }
                    }
                    Event::Empty(e) => {
                        match e.local_name().as_ref() {
                            b"cBhvr" => {
                                f_c_bhvr = Some(Box::new(CTTLCommonBehaviorData::from_xml(
                                    reader, &e, true,
                                )?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            b"by" => {
                                f_by = Some(Box::new(CTTLPoint::from_xml(reader, &e, true)?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            b"from" => {
                                f_from = Some(Box::new(CTTLPoint::from_xml(reader, &e, true)?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            b"to" => {
                                f_to = Some(Box::new(CTTLPoint::from_xml(reader, &e, true)?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            #[cfg(feature = "extra-children")]
                            _ => {
                                // Capture unknown empty element for roundtrip
                                let elem = RawXmlElement::from_empty(&e);
                                extra_children.push(PositionedNode::new(
                                    child_idx,
                                    RawXmlNode::Element(elem),
                                ));
                                child_idx += 1;
                            }
                            #[cfg(not(feature = "extra-children"))]
                            _ => {}
                        }
                    }
                    Event::End(_) => break,
                    Event::Eof => break,
                    _ => {}
                }
                buf.clear();
            }
        }

        Ok(Self {
            zoom_contents: f_zoom_contents,
            c_bhvr: f_c_bhvr.ok_or_else(|| ParseError::MissingAttribute("cBhvr".to_string()))?,
            by: f_by,
            from: f_from,
            to: f_to,
            #[cfg(feature = "extra-attrs")]
            extra_attrs,
            #[cfg(feature = "extra-children")]
            extra_children,
        })
    }
}

impl FromXml for CTTLCommandBehavior {
    fn from_xml<R: BufRead>(
        reader: &mut Reader<R>,
        start_tag: &BytesStart,
        is_empty: bool,
    ) -> Result<Self, ParseError> {
        let mut f_type = None;
        let mut f_cmd = None;
        let mut f_c_bhvr: Option<Box<CTTLCommonBehaviorData>> = None;
        #[cfg(feature = "extra-attrs")]
        let mut extra_attrs = std::collections::HashMap::new();
        #[cfg(feature = "extra-children")]
        let mut extra_children = Vec::new();
        #[cfg(feature = "extra-children")]
        let mut child_idx: usize = 0;

        // Parse attributes
        for attr in start_tag.attributes().filter_map(|a| a.ok()) {
            let val = String::from_utf8_lossy(&attr.value);
            match attr.key.local_name().as_ref() {
                b"type" => {
                    f_type = val.parse().ok();
                }
                b"cmd" => {
                    f_cmd = Some(val.into_owned());
                }
                #[cfg(feature = "extra-attrs")]
                unknown => {
                    let key = String::from_utf8_lossy(attr.key.as_ref()).into_owned();
                    extra_attrs.insert(key, val.into_owned());
                }
                #[cfg(not(feature = "extra-attrs"))]
                _ => {}
            }
        }

        // Parse child elements
        if !is_empty {
            let mut buf = Vec::new();
            loop {
                match reader.read_event_into(&mut buf)? {
                    Event::Start(e) => {
                        match e.local_name().as_ref() {
                            b"cBhvr" => {
                                f_c_bhvr = Some(Box::new(CTTLCommonBehaviorData::from_xml(
                                    reader, &e, false,
                                )?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            #[cfg(feature = "extra-children")]
                            _ => {
                                // Capture unknown element for roundtrip
                                let elem = RawXmlElement::from_reader(reader, &e)?;
                                extra_children.push(PositionedNode::new(
                                    child_idx,
                                    RawXmlNode::Element(elem),
                                ));
                                child_idx += 1;
                            }
                            #[cfg(not(feature = "extra-children"))]
                            _ => {
                                // Skip unknown element
                                skip_element(reader)?;
                            }
                        }
                    }
                    Event::Empty(e) => {
                        match e.local_name().as_ref() {
                            b"cBhvr" => {
                                f_c_bhvr = Some(Box::new(CTTLCommonBehaviorData::from_xml(
                                    reader, &e, true,
                                )?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            #[cfg(feature = "extra-children")]
                            _ => {
                                // Capture unknown empty element for roundtrip
                                let elem = RawXmlElement::from_empty(&e);
                                extra_children.push(PositionedNode::new(
                                    child_idx,
                                    RawXmlNode::Element(elem),
                                ));
                                child_idx += 1;
                            }
                            #[cfg(not(feature = "extra-children"))]
                            _ => {}
                        }
                    }
                    Event::End(_) => break,
                    Event::Eof => break,
                    _ => {}
                }
                buf.clear();
            }
        }

        Ok(Self {
            r#type: f_type,
            cmd: f_cmd,
            c_bhvr: f_c_bhvr.ok_or_else(|| ParseError::MissingAttribute("cBhvr".to_string()))?,
            #[cfg(feature = "extra-attrs")]
            extra_attrs,
            #[cfg(feature = "extra-children")]
            extra_children,
        })
    }
}

impl FromXml for CTTLSetBehavior {
    fn from_xml<R: BufRead>(
        reader: &mut Reader<R>,
        start_tag: &BytesStart,
        is_empty: bool,
    ) -> Result<Self, ParseError> {
        let mut f_c_bhvr: Option<Box<CTTLCommonBehaviorData>> = None;
        let mut f_to = None;
        #[cfg(feature = "extra-children")]
        let mut extra_children = Vec::new();
        #[cfg(feature = "extra-children")]
        let mut child_idx: usize = 0;

        // Parse child elements
        if !is_empty {
            let mut buf = Vec::new();
            loop {
                match reader.read_event_into(&mut buf)? {
                    Event::Start(e) => {
                        match e.local_name().as_ref() {
                            b"cBhvr" => {
                                f_c_bhvr = Some(Box::new(CTTLCommonBehaviorData::from_xml(
                                    reader, &e, false,
                                )?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            b"to" => {
                                f_to =
                                    Some(Box::new(CTTLAnimVariant::from_xml(reader, &e, false)?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            #[cfg(feature = "extra-children")]
                            _ => {
                                // Capture unknown element for roundtrip
                                let elem = RawXmlElement::from_reader(reader, &e)?;
                                extra_children.push(PositionedNode::new(
                                    child_idx,
                                    RawXmlNode::Element(elem),
                                ));
                                child_idx += 1;
                            }
                            #[cfg(not(feature = "extra-children"))]
                            _ => {
                                // Skip unknown element
                                skip_element(reader)?;
                            }
                        }
                    }
                    Event::Empty(e) => {
                        match e.local_name().as_ref() {
                            b"cBhvr" => {
                                f_c_bhvr = Some(Box::new(CTTLCommonBehaviorData::from_xml(
                                    reader, &e, true,
                                )?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            b"to" => {
                                f_to = Some(Box::new(CTTLAnimVariant::from_xml(reader, &e, true)?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            #[cfg(feature = "extra-children")]
                            _ => {
                                // Capture unknown empty element for roundtrip
                                let elem = RawXmlElement::from_empty(&e);
                                extra_children.push(PositionedNode::new(
                                    child_idx,
                                    RawXmlNode::Element(elem),
                                ));
                                child_idx += 1;
                            }
                            #[cfg(not(feature = "extra-children"))]
                            _ => {}
                        }
                    }
                    Event::End(_) => break,
                    Event::Eof => break,
                    _ => {}
                }
                buf.clear();
            }
        }

        Ok(Self {
            c_bhvr: f_c_bhvr.ok_or_else(|| ParseError::MissingAttribute("cBhvr".to_string()))?,
            to: f_to,
            #[cfg(feature = "extra-children")]
            extra_children,
        })
    }
}

impl FromXml for CTTLCommonMediaNodeData {
    fn from_xml<R: BufRead>(
        reader: &mut Reader<R>,
        start_tag: &BytesStart,
        is_empty: bool,
    ) -> Result<Self, ParseError> {
        let mut f_vol = None;
        let mut f_mute = None;
        let mut f_num_sld = None;
        let mut f_show_when_stopped = None;
        let mut f_c_tn: Option<Box<CTTLCommonTimeNodeData>> = None;
        let mut f_tgt_el: Option<Box<CTTLTimeTargetElement>> = None;
        #[cfg(feature = "extra-attrs")]
        let mut extra_attrs = std::collections::HashMap::new();
        #[cfg(feature = "extra-children")]
        let mut extra_children = Vec::new();
        #[cfg(feature = "extra-children")]
        let mut child_idx: usize = 0;

        // Parse attributes
        for attr in start_tag.attributes().filter_map(|a| a.ok()) {
            let val = String::from_utf8_lossy(&attr.value);
            match attr.key.local_name().as_ref() {
                b"vol" => {
                    f_vol = val.parse().ok();
                }
                b"mute" => {
                    f_mute = Some(val == "true" || val == "1");
                }
                b"numSld" => {
                    f_num_sld = val.parse().ok();
                }
                b"showWhenStopped" => {
                    f_show_when_stopped = Some(val == "true" || val == "1");
                }
                #[cfg(feature = "extra-attrs")]
                unknown => {
                    let key = String::from_utf8_lossy(attr.key.as_ref()).into_owned();
                    extra_attrs.insert(key, val.into_owned());
                }
                #[cfg(not(feature = "extra-attrs"))]
                _ => {}
            }
        }

        // Parse child elements
        if !is_empty {
            let mut buf = Vec::new();
            loop {
                match reader.read_event_into(&mut buf)? {
                    Event::Start(e) => {
                        match e.local_name().as_ref() {
                            b"cTn" => {
                                f_c_tn = Some(Box::new(CTTLCommonTimeNodeData::from_xml(
                                    reader, &e, false,
                                )?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            b"tgtEl" => {
                                f_tgt_el = Some(Box::new(CTTLTimeTargetElement::from_xml(
                                    reader, &e, false,
                                )?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            #[cfg(feature = "extra-children")]
                            _ => {
                                // Capture unknown element for roundtrip
                                let elem = RawXmlElement::from_reader(reader, &e)?;
                                extra_children.push(PositionedNode::new(
                                    child_idx,
                                    RawXmlNode::Element(elem),
                                ));
                                child_idx += 1;
                            }
                            #[cfg(not(feature = "extra-children"))]
                            _ => {
                                // Skip unknown element
                                skip_element(reader)?;
                            }
                        }
                    }
                    Event::Empty(e) => {
                        match e.local_name().as_ref() {
                            b"cTn" => {
                                f_c_tn = Some(Box::new(CTTLCommonTimeNodeData::from_xml(
                                    reader, &e, true,
                                )?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            b"tgtEl" => {
                                f_tgt_el = Some(Box::new(CTTLTimeTargetElement::from_xml(
                                    reader, &e, true,
                                )?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            #[cfg(feature = "extra-children")]
                            _ => {
                                // Capture unknown empty element for roundtrip
                                let elem = RawXmlElement::from_empty(&e);
                                extra_children.push(PositionedNode::new(
                                    child_idx,
                                    RawXmlNode::Element(elem),
                                ));
                                child_idx += 1;
                            }
                            #[cfg(not(feature = "extra-children"))]
                            _ => {}
                        }
                    }
                    Event::End(_) => break,
                    Event::Eof => break,
                    _ => {}
                }
                buf.clear();
            }
        }

        Ok(Self {
            vol: f_vol,
            mute: f_mute,
            num_sld: f_num_sld,
            show_when_stopped: f_show_when_stopped,
            c_tn: f_c_tn.ok_or_else(|| ParseError::MissingAttribute("cTn".to_string()))?,
            tgt_el: f_tgt_el.ok_or_else(|| ParseError::MissingAttribute("tgtEl".to_string()))?,
            #[cfg(feature = "extra-attrs")]
            extra_attrs,
            #[cfg(feature = "extra-children")]
            extra_children,
        })
    }
}

impl FromXml for CTTLMediaNodeAudio {
    fn from_xml<R: BufRead>(
        reader: &mut Reader<R>,
        start_tag: &BytesStart,
        is_empty: bool,
    ) -> Result<Self, ParseError> {
        let mut f_is_narration = None;
        let mut f_c_media_node: Option<Box<CTTLCommonMediaNodeData>> = None;
        #[cfg(feature = "extra-attrs")]
        let mut extra_attrs = std::collections::HashMap::new();
        #[cfg(feature = "extra-children")]
        let mut extra_children = Vec::new();
        #[cfg(feature = "extra-children")]
        let mut child_idx: usize = 0;

        // Parse attributes
        for attr in start_tag.attributes().filter_map(|a| a.ok()) {
            let val = String::from_utf8_lossy(&attr.value);
            match attr.key.local_name().as_ref() {
                b"isNarration" => {
                    f_is_narration = Some(val == "true" || val == "1");
                }
                #[cfg(feature = "extra-attrs")]
                unknown => {
                    let key = String::from_utf8_lossy(attr.key.as_ref()).into_owned();
                    extra_attrs.insert(key, val.into_owned());
                }
                #[cfg(not(feature = "extra-attrs"))]
                _ => {}
            }
        }

        // Parse child elements
        if !is_empty {
            let mut buf = Vec::new();
            loop {
                match reader.read_event_into(&mut buf)? {
                    Event::Start(e) => {
                        match e.local_name().as_ref() {
                            b"cMediaNode" => {
                                f_c_media_node = Some(Box::new(CTTLCommonMediaNodeData::from_xml(
                                    reader, &e, false,
                                )?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            #[cfg(feature = "extra-children")]
                            _ => {
                                // Capture unknown element for roundtrip
                                let elem = RawXmlElement::from_reader(reader, &e)?;
                                extra_children.push(PositionedNode::new(
                                    child_idx,
                                    RawXmlNode::Element(elem),
                                ));
                                child_idx += 1;
                            }
                            #[cfg(not(feature = "extra-children"))]
                            _ => {
                                // Skip unknown element
                                skip_element(reader)?;
                            }
                        }
                    }
                    Event::Empty(e) => {
                        match e.local_name().as_ref() {
                            b"cMediaNode" => {
                                f_c_media_node = Some(Box::new(CTTLCommonMediaNodeData::from_xml(
                                    reader, &e, true,
                                )?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            #[cfg(feature = "extra-children")]
                            _ => {
                                // Capture unknown empty element for roundtrip
                                let elem = RawXmlElement::from_empty(&e);
                                extra_children.push(PositionedNode::new(
                                    child_idx,
                                    RawXmlNode::Element(elem),
                                ));
                                child_idx += 1;
                            }
                            #[cfg(not(feature = "extra-children"))]
                            _ => {}
                        }
                    }
                    Event::End(_) => break,
                    Event::Eof => break,
                    _ => {}
                }
                buf.clear();
            }
        }

        Ok(Self {
            is_narration: f_is_narration,
            c_media_node: f_c_media_node
                .ok_or_else(|| ParseError::MissingAttribute("cMediaNode".to_string()))?,
            #[cfg(feature = "extra-attrs")]
            extra_attrs,
            #[cfg(feature = "extra-children")]
            extra_children,
        })
    }
}

impl FromXml for CTTLMediaNodeVideo {
    fn from_xml<R: BufRead>(
        reader: &mut Reader<R>,
        start_tag: &BytesStart,
        is_empty: bool,
    ) -> Result<Self, ParseError> {
        let mut f_full_scrn = None;
        let mut f_c_media_node: Option<Box<CTTLCommonMediaNodeData>> = None;
        #[cfg(feature = "extra-attrs")]
        let mut extra_attrs = std::collections::HashMap::new();
        #[cfg(feature = "extra-children")]
        let mut extra_children = Vec::new();
        #[cfg(feature = "extra-children")]
        let mut child_idx: usize = 0;

        // Parse attributes
        for attr in start_tag.attributes().filter_map(|a| a.ok()) {
            let val = String::from_utf8_lossy(&attr.value);
            match attr.key.local_name().as_ref() {
                b"fullScrn" => {
                    f_full_scrn = Some(val == "true" || val == "1");
                }
                #[cfg(feature = "extra-attrs")]
                unknown => {
                    let key = String::from_utf8_lossy(attr.key.as_ref()).into_owned();
                    extra_attrs.insert(key, val.into_owned());
                }
                #[cfg(not(feature = "extra-attrs"))]
                _ => {}
            }
        }

        // Parse child elements
        if !is_empty {
            let mut buf = Vec::new();
            loop {
                match reader.read_event_into(&mut buf)? {
                    Event::Start(e) => {
                        match e.local_name().as_ref() {
                            b"cMediaNode" => {
                                f_c_media_node = Some(Box::new(CTTLCommonMediaNodeData::from_xml(
                                    reader, &e, false,
                                )?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            #[cfg(feature = "extra-children")]
                            _ => {
                                // Capture unknown element for roundtrip
                                let elem = RawXmlElement::from_reader(reader, &e)?;
                                extra_children.push(PositionedNode::new(
                                    child_idx,
                                    RawXmlNode::Element(elem),
                                ));
                                child_idx += 1;
                            }
                            #[cfg(not(feature = "extra-children"))]
                            _ => {
                                // Skip unknown element
                                skip_element(reader)?;
                            }
                        }
                    }
                    Event::Empty(e) => {
                        match e.local_name().as_ref() {
                            b"cMediaNode" => {
                                f_c_media_node = Some(Box::new(CTTLCommonMediaNodeData::from_xml(
                                    reader, &e, true,
                                )?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            #[cfg(feature = "extra-children")]
                            _ => {
                                // Capture unknown empty element for roundtrip
                                let elem = RawXmlElement::from_empty(&e);
                                extra_children.push(PositionedNode::new(
                                    child_idx,
                                    RawXmlNode::Element(elem),
                                ));
                                child_idx += 1;
                            }
                            #[cfg(not(feature = "extra-children"))]
                            _ => {}
                        }
                    }
                    Event::End(_) => break,
                    Event::Eof => break,
                    _ => {}
                }
                buf.clear();
            }
        }

        Ok(Self {
            full_scrn: f_full_scrn,
            c_media_node: f_c_media_node
                .ok_or_else(|| ParseError::MissingAttribute("cMediaNode".to_string()))?,
            #[cfg(feature = "extra-attrs")]
            extra_attrs,
            #[cfg(feature = "extra-children")]
            extra_children,
        })
    }
}

impl FromXml for PAGTLBuild {
    fn from_xml<R: BufRead>(
        reader: &mut Reader<R>,
        start_tag: &BytesStart,
        is_empty: bool,
    ) -> Result<Self, ParseError> {
        let mut f_spid: Option<ooxml_dml::types::STDrawingElementId> = None;
        let mut f_grp_id: Option<u32> = None;
        let mut f_ui_expand = None;
        #[cfg(feature = "extra-attrs")]
        let mut extra_attrs = std::collections::HashMap::new();

        // Parse attributes
        for attr in start_tag.attributes().filter_map(|a| a.ok()) {
            let val = String::from_utf8_lossy(&attr.value);
            match attr.key.local_name().as_ref() {
                b"spid" => {
                    f_spid = val.parse().ok();
                }
                b"grpId" => {
                    f_grp_id = val.parse().ok();
                }
                b"uiExpand" => {
                    f_ui_expand = Some(val == "true" || val == "1");
                }
                #[cfg(feature = "extra-attrs")]
                unknown => {
                    let key = String::from_utf8_lossy(attr.key.as_ref()).into_owned();
                    extra_attrs.insert(key, val.into_owned());
                }
                #[cfg(not(feature = "extra-attrs"))]
                _ => {}
            }
        }

        if !is_empty {
            let mut buf = Vec::new();
            loop {
                match reader.read_event_into(&mut buf)? {
                    Event::End(_) => break,
                    Event::Eof => break,
                    _ => {}
                }
                buf.clear();
            }
        }

        Ok(Self {
            spid: f_spid.ok_or_else(|| ParseError::MissingAttribute("spid".to_string()))?,
            grp_id: f_grp_id.ok_or_else(|| ParseError::MissingAttribute("grpId".to_string()))?,
            ui_expand: f_ui_expand,
            #[cfg(feature = "extra-attrs")]
            extra_attrs,
        })
    }
}

impl FromXml for CTTLTemplate {
    fn from_xml<R: BufRead>(
        reader: &mut Reader<R>,
        start_tag: &BytesStart,
        is_empty: bool,
    ) -> Result<Self, ParseError> {
        let mut f_lvl = None;
        let mut f_tn_lst: Option<Box<CTTimeNodeList>> = None;
        #[cfg(feature = "extra-attrs")]
        let mut extra_attrs = std::collections::HashMap::new();
        #[cfg(feature = "extra-children")]
        let mut extra_children = Vec::new();
        #[cfg(feature = "extra-children")]
        let mut child_idx: usize = 0;

        // Parse attributes
        for attr in start_tag.attributes().filter_map(|a| a.ok()) {
            let val = String::from_utf8_lossy(&attr.value);
            match attr.key.local_name().as_ref() {
                b"lvl" => {
                    f_lvl = val.parse().ok();
                }
                #[cfg(feature = "extra-attrs")]
                unknown => {
                    let key = String::from_utf8_lossy(attr.key.as_ref()).into_owned();
                    extra_attrs.insert(key, val.into_owned());
                }
                #[cfg(not(feature = "extra-attrs"))]
                _ => {}
            }
        }

        // Parse child elements
        if !is_empty {
            let mut buf = Vec::new();
            loop {
                match reader.read_event_into(&mut buf)? {
                    Event::Start(e) => {
                        match e.local_name().as_ref() {
                            b"tnLst" => {
                                f_tn_lst =
                                    Some(Box::new(CTTimeNodeList::from_xml(reader, &e, false)?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            #[cfg(feature = "extra-children")]
                            _ => {
                                // Capture unknown element for roundtrip
                                let elem = RawXmlElement::from_reader(reader, &e)?;
                                extra_children.push(PositionedNode::new(
                                    child_idx,
                                    RawXmlNode::Element(elem),
                                ));
                                child_idx += 1;
                            }
                            #[cfg(not(feature = "extra-children"))]
                            _ => {
                                // Skip unknown element
                                skip_element(reader)?;
                            }
                        }
                    }
                    Event::Empty(e) => {
                        match e.local_name().as_ref() {
                            b"tnLst" => {
                                f_tn_lst =
                                    Some(Box::new(CTTimeNodeList::from_xml(reader, &e, true)?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            #[cfg(feature = "extra-children")]
                            _ => {
                                // Capture unknown empty element for roundtrip
                                let elem = RawXmlElement::from_empty(&e);
                                extra_children.push(PositionedNode::new(
                                    child_idx,
                                    RawXmlNode::Element(elem),
                                ));
                                child_idx += 1;
                            }
                            #[cfg(not(feature = "extra-children"))]
                            _ => {}
                        }
                    }
                    Event::End(_) => break,
                    Event::Eof => break,
                    _ => {}
                }
                buf.clear();
            }
        }

        Ok(Self {
            lvl: f_lvl,
            tn_lst: f_tn_lst.ok_or_else(|| ParseError::MissingAttribute("tnLst".to_string()))?,
            #[cfg(feature = "extra-attrs")]
            extra_attrs,
            #[cfg(feature = "extra-children")]
            extra_children,
        })
    }
}

impl FromXml for CTTLTemplateList {
    fn from_xml<R: BufRead>(
        reader: &mut Reader<R>,
        start_tag: &BytesStart,
        is_empty: bool,
    ) -> Result<Self, ParseError> {
        let mut f_tmpl = Vec::new();
        #[cfg(feature = "extra-children")]
        let mut extra_children = Vec::new();
        #[cfg(feature = "extra-children")]
        let mut child_idx: usize = 0;

        // Parse child elements
        if !is_empty {
            let mut buf = Vec::new();
            loop {
                match reader.read_event_into(&mut buf)? {
                    Event::Start(e) => {
                        match e.local_name().as_ref() {
                            b"tmpl" => {
                                f_tmpl.push(CTTLTemplate::from_xml(reader, &e, false)?);
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            #[cfg(feature = "extra-children")]
                            _ => {
                                // Capture unknown element for roundtrip
                                let elem = RawXmlElement::from_reader(reader, &e)?;
                                extra_children.push(PositionedNode::new(
                                    child_idx,
                                    RawXmlNode::Element(elem),
                                ));
                                child_idx += 1;
                            }
                            #[cfg(not(feature = "extra-children"))]
                            _ => {
                                // Skip unknown element
                                skip_element(reader)?;
                            }
                        }
                    }
                    Event::Empty(e) => {
                        match e.local_name().as_ref() {
                            b"tmpl" => {
                                f_tmpl.push(CTTLTemplate::from_xml(reader, &e, true)?);
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            #[cfg(feature = "extra-children")]
                            _ => {
                                // Capture unknown empty element for roundtrip
                                let elem = RawXmlElement::from_empty(&e);
                                extra_children.push(PositionedNode::new(
                                    child_idx,
                                    RawXmlNode::Element(elem),
                                ));
                                child_idx += 1;
                            }
                            #[cfg(not(feature = "extra-children"))]
                            _ => {}
                        }
                    }
                    Event::End(_) => break,
                    Event::Eof => break,
                    _ => {}
                }
                buf.clear();
            }
        }

        Ok(Self {
            tmpl: f_tmpl,
            #[cfg(feature = "extra-children")]
            extra_children,
        })
    }
}

impl FromXml for CTTLBuildParagraph {
    fn from_xml<R: BufRead>(
        reader: &mut Reader<R>,
        start_tag: &BytesStart,
        is_empty: bool,
    ) -> Result<Self, ParseError> {
        let mut f_spid: Option<ooxml_dml::types::STDrawingElementId> = None;
        let mut f_grp_id: Option<u32> = None;
        let mut f_ui_expand = None;
        let mut f_build = None;
        let mut f_bld_lvl = None;
        let mut f_anim_bg = None;
        let mut f_auto_update_anim_bg = None;
        let mut f_rev = None;
        let mut f_adv_auto = None;
        let mut f_tmpl_lst = None;
        #[cfg(feature = "extra-attrs")]
        let mut extra_attrs = std::collections::HashMap::new();
        #[cfg(feature = "extra-children")]
        let mut extra_children = Vec::new();
        #[cfg(feature = "extra-children")]
        let mut child_idx: usize = 0;

        // Parse attributes
        for attr in start_tag.attributes().filter_map(|a| a.ok()) {
            let val = String::from_utf8_lossy(&attr.value);
            match attr.key.local_name().as_ref() {
                b"spid" => {
                    f_spid = val.parse().ok();
                }
                b"grpId" => {
                    f_grp_id = val.parse().ok();
                }
                b"uiExpand" => {
                    f_ui_expand = Some(val == "true" || val == "1");
                }
                b"build" => {
                    f_build = val.parse().ok();
                }
                b"bldLvl" => {
                    f_bld_lvl = val.parse().ok();
                }
                b"animBg" => {
                    f_anim_bg = Some(val == "true" || val == "1");
                }
                b"autoUpdateAnimBg" => {
                    f_auto_update_anim_bg = Some(val == "true" || val == "1");
                }
                b"rev" => {
                    f_rev = Some(val == "true" || val == "1");
                }
                b"advAuto" => {
                    f_adv_auto = Some(val.into_owned());
                }
                #[cfg(feature = "extra-attrs")]
                unknown => {
                    let key = String::from_utf8_lossy(attr.key.as_ref()).into_owned();
                    extra_attrs.insert(key, val.into_owned());
                }
                #[cfg(not(feature = "extra-attrs"))]
                _ => {}
            }
        }

        // Parse child elements
        if !is_empty {
            let mut buf = Vec::new();
            loop {
                match reader.read_event_into(&mut buf)? {
                    Event::Start(e) => {
                        match e.local_name().as_ref() {
                            b"tmplLst" => {
                                f_tmpl_lst =
                                    Some(Box::new(CTTLTemplateList::from_xml(reader, &e, false)?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            #[cfg(feature = "extra-children")]
                            _ => {
                                // Capture unknown element for roundtrip
                                let elem = RawXmlElement::from_reader(reader, &e)?;
                                extra_children.push(PositionedNode::new(
                                    child_idx,
                                    RawXmlNode::Element(elem),
                                ));
                                child_idx += 1;
                            }
                            #[cfg(not(feature = "extra-children"))]
                            _ => {
                                // Skip unknown element
                                skip_element(reader)?;
                            }
                        }
                    }
                    Event::Empty(e) => {
                        match e.local_name().as_ref() {
                            b"tmplLst" => {
                                f_tmpl_lst =
                                    Some(Box::new(CTTLTemplateList::from_xml(reader, &e, true)?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            #[cfg(feature = "extra-children")]
                            _ => {
                                // Capture unknown empty element for roundtrip
                                let elem = RawXmlElement::from_empty(&e);
                                extra_children.push(PositionedNode::new(
                                    child_idx,
                                    RawXmlNode::Element(elem),
                                ));
                                child_idx += 1;
                            }
                            #[cfg(not(feature = "extra-children"))]
                            _ => {}
                        }
                    }
                    Event::End(_) => break,
                    Event::Eof => break,
                    _ => {}
                }
                buf.clear();
            }
        }

        Ok(Self {
            spid: f_spid.ok_or_else(|| ParseError::MissingAttribute("spid".to_string()))?,
            grp_id: f_grp_id.ok_or_else(|| ParseError::MissingAttribute("grpId".to_string()))?,
            ui_expand: f_ui_expand,
            build: f_build,
            bld_lvl: f_bld_lvl,
            anim_bg: f_anim_bg,
            auto_update_anim_bg: f_auto_update_anim_bg,
            rev: f_rev,
            adv_auto: f_adv_auto,
            tmpl_lst: f_tmpl_lst,
            #[cfg(feature = "extra-attrs")]
            extra_attrs,
            #[cfg(feature = "extra-children")]
            extra_children,
        })
    }
}

impl FromXml for CTTLBuildDiagram {
    fn from_xml<R: BufRead>(
        reader: &mut Reader<R>,
        start_tag: &BytesStart,
        is_empty: bool,
    ) -> Result<Self, ParseError> {
        let mut f_spid: Option<ooxml_dml::types::STDrawingElementId> = None;
        let mut f_grp_id: Option<u32> = None;
        let mut f_ui_expand = None;
        let mut f_bld = None;
        #[cfg(feature = "extra-attrs")]
        let mut extra_attrs = std::collections::HashMap::new();

        // Parse attributes
        for attr in start_tag.attributes().filter_map(|a| a.ok()) {
            let val = String::from_utf8_lossy(&attr.value);
            match attr.key.local_name().as_ref() {
                b"spid" => {
                    f_spid = val.parse().ok();
                }
                b"grpId" => {
                    f_grp_id = val.parse().ok();
                }
                b"uiExpand" => {
                    f_ui_expand = Some(val == "true" || val == "1");
                }
                b"bld" => {
                    f_bld = val.parse().ok();
                }
                #[cfg(feature = "extra-attrs")]
                unknown => {
                    let key = String::from_utf8_lossy(attr.key.as_ref()).into_owned();
                    extra_attrs.insert(key, val.into_owned());
                }
                #[cfg(not(feature = "extra-attrs"))]
                _ => {}
            }
        }

        if !is_empty {
            let mut buf = Vec::new();
            loop {
                match reader.read_event_into(&mut buf)? {
                    Event::End(_) => break,
                    Event::Eof => break,
                    _ => {}
                }
                buf.clear();
            }
        }

        Ok(Self {
            spid: f_spid.ok_or_else(|| ParseError::MissingAttribute("spid".to_string()))?,
            grp_id: f_grp_id.ok_or_else(|| ParseError::MissingAttribute("grpId".to_string()))?,
            ui_expand: f_ui_expand,
            bld: f_bld,
            #[cfg(feature = "extra-attrs")]
            extra_attrs,
        })
    }
}

impl FromXml for CTTLOleBuildChart {
    fn from_xml<R: BufRead>(
        reader: &mut Reader<R>,
        start_tag: &BytesStart,
        is_empty: bool,
    ) -> Result<Self, ParseError> {
        let mut f_spid: Option<ooxml_dml::types::STDrawingElementId> = None;
        let mut f_grp_id: Option<u32> = None;
        let mut f_ui_expand = None;
        let mut f_bld = None;
        let mut f_anim_bg = None;
        #[cfg(feature = "extra-attrs")]
        let mut extra_attrs = std::collections::HashMap::new();

        // Parse attributes
        for attr in start_tag.attributes().filter_map(|a| a.ok()) {
            let val = String::from_utf8_lossy(&attr.value);
            match attr.key.local_name().as_ref() {
                b"spid" => {
                    f_spid = val.parse().ok();
                }
                b"grpId" => {
                    f_grp_id = val.parse().ok();
                }
                b"uiExpand" => {
                    f_ui_expand = Some(val == "true" || val == "1");
                }
                b"bld" => {
                    f_bld = val.parse().ok();
                }
                b"animBg" => {
                    f_anim_bg = Some(val == "true" || val == "1");
                }
                #[cfg(feature = "extra-attrs")]
                unknown => {
                    let key = String::from_utf8_lossy(attr.key.as_ref()).into_owned();
                    extra_attrs.insert(key, val.into_owned());
                }
                #[cfg(not(feature = "extra-attrs"))]
                _ => {}
            }
        }

        if !is_empty {
            let mut buf = Vec::new();
            loop {
                match reader.read_event_into(&mut buf)? {
                    Event::End(_) => break,
                    Event::Eof => break,
                    _ => {}
                }
                buf.clear();
            }
        }

        Ok(Self {
            spid: f_spid.ok_or_else(|| ParseError::MissingAttribute("spid".to_string()))?,
            grp_id: f_grp_id.ok_or_else(|| ParseError::MissingAttribute("grpId".to_string()))?,
            ui_expand: f_ui_expand,
            bld: f_bld,
            anim_bg: f_anim_bg,
            #[cfg(feature = "extra-attrs")]
            extra_attrs,
        })
    }
}

impl FromXml for CTTLGraphicalObjectBuild {
    fn from_xml<R: BufRead>(
        reader: &mut Reader<R>,
        start_tag: &BytesStart,
        is_empty: bool,
    ) -> Result<Self, ParseError> {
        let mut f_spid: Option<ooxml_dml::types::STDrawingElementId> = None;
        let mut f_grp_id: Option<u32> = None;
        let mut f_ui_expand = None;
        let mut f_bld_as_one = None;
        let mut f_bld_sub = None;
        #[cfg(feature = "extra-attrs")]
        let mut extra_attrs = std::collections::HashMap::new();
        #[cfg(feature = "extra-children")]
        let mut extra_children = Vec::new();
        #[cfg(feature = "extra-children")]
        let mut child_idx: usize = 0;

        // Parse attributes
        for attr in start_tag.attributes().filter_map(|a| a.ok()) {
            let val = String::from_utf8_lossy(&attr.value);
            match attr.key.local_name().as_ref() {
                b"spid" => {
                    f_spid = val.parse().ok();
                }
                b"grpId" => {
                    f_grp_id = val.parse().ok();
                }
                b"uiExpand" => {
                    f_ui_expand = Some(val == "true" || val == "1");
                }
                #[cfg(feature = "extra-attrs")]
                unknown => {
                    let key = String::from_utf8_lossy(attr.key.as_ref()).into_owned();
                    extra_attrs.insert(key, val.into_owned());
                }
                #[cfg(not(feature = "extra-attrs"))]
                _ => {}
            }
        }

        // Parse child elements
        if !is_empty {
            let mut buf = Vec::new();
            loop {
                match reader.read_event_into(&mut buf)? {
                    Event::Start(e) => {
                        match e.local_name().as_ref() {
                            b"bldAsOne" => {
                                f_bld_as_one =
                                    Some(Box::new(CTEmpty::from_xml(reader, &e, false)?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            b"bldSub" => {
                                f_bld_sub = Some(Box::new(ooxml_dml::types::CTAnimationGraphicalObjectBuildProperties::from_xml(reader, &e, false)?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            #[cfg(feature = "extra-children")]
                            _ => {
                                // Capture unknown element for roundtrip
                                let elem = RawXmlElement::from_reader(reader, &e)?;
                                extra_children.push(PositionedNode::new(
                                    child_idx,
                                    RawXmlNode::Element(elem),
                                ));
                                child_idx += 1;
                            }
                            #[cfg(not(feature = "extra-children"))]
                            _ => {
                                // Skip unknown element
                                skip_element(reader)?;
                            }
                        }
                    }
                    Event::Empty(e) => {
                        match e.local_name().as_ref() {
                            b"bldAsOne" => {
                                f_bld_as_one = Some(Box::new(CTEmpty::from_xml(reader, &e, true)?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            b"bldSub" => {
                                f_bld_sub = Some(Box::new(ooxml_dml::types::CTAnimationGraphicalObjectBuildProperties::from_xml(reader, &e, true)?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            #[cfg(feature = "extra-children")]
                            _ => {
                                // Capture unknown empty element for roundtrip
                                let elem = RawXmlElement::from_empty(&e);
                                extra_children.push(PositionedNode::new(
                                    child_idx,
                                    RawXmlNode::Element(elem),
                                ));
                                child_idx += 1;
                            }
                            #[cfg(not(feature = "extra-children"))]
                            _ => {}
                        }
                    }
                    Event::End(_) => break,
                    Event::Eof => break,
                    _ => {}
                }
                buf.clear();
            }
        }

        Ok(Self {
            spid: f_spid.ok_or_else(|| ParseError::MissingAttribute("spid".to_string()))?,
            grp_id: f_grp_id.ok_or_else(|| ParseError::MissingAttribute("grpId".to_string()))?,
            ui_expand: f_ui_expand,
            bld_as_one: f_bld_as_one,
            bld_sub: f_bld_sub,
            #[cfg(feature = "extra-attrs")]
            extra_attrs,
            #[cfg(feature = "extra-children")]
            extra_children,
        })
    }
}

impl FromXml for CTBuildList {
    fn from_xml<R: BufRead>(
        reader: &mut Reader<R>,
        start_tag: &BytesStart,
        is_empty: bool,
    ) -> Result<Self, ParseError> {
        let mut f_bld_p = None;
        let mut f_bld_dgm = None;
        let mut f_bld_ole_chart = None;
        let mut f_bld_graphic = None;
        #[cfg(feature = "extra-children")]
        let mut extra_children = Vec::new();
        #[cfg(feature = "extra-children")]
        let mut child_idx: usize = 0;

        // Parse child elements
        if !is_empty {
            let mut buf = Vec::new();
            loop {
                match reader.read_event_into(&mut buf)? {
                    Event::Start(e) => {
                        match e.local_name().as_ref() {
                            b"bldP" => {
                                f_bld_p = Some(Box::new(CTTLBuildParagraph::from_xml(
                                    reader, &e, false,
                                )?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            b"bldDgm" => {
                                f_bld_dgm =
                                    Some(Box::new(CTTLBuildDiagram::from_xml(reader, &e, false)?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            b"bldOleChart" => {
                                f_bld_ole_chart =
                                    Some(Box::new(CTTLOleBuildChart::from_xml(reader, &e, false)?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            b"bldGraphic" => {
                                f_bld_graphic = Some(Box::new(CTTLGraphicalObjectBuild::from_xml(
                                    reader, &e, false,
                                )?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            #[cfg(feature = "extra-children")]
                            _ => {
                                // Capture unknown element for roundtrip
                                let elem = RawXmlElement::from_reader(reader, &e)?;
                                extra_children.push(PositionedNode::new(
                                    child_idx,
                                    RawXmlNode::Element(elem),
                                ));
                                child_idx += 1;
                            }
                            #[cfg(not(feature = "extra-children"))]
                            _ => {
                                // Skip unknown element
                                skip_element(reader)?;
                            }
                        }
                    }
                    Event::Empty(e) => {
                        match e.local_name().as_ref() {
                            b"bldP" => {
                                f_bld_p =
                                    Some(Box::new(CTTLBuildParagraph::from_xml(reader, &e, true)?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            b"bldDgm" => {
                                f_bld_dgm =
                                    Some(Box::new(CTTLBuildDiagram::from_xml(reader, &e, true)?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            b"bldOleChart" => {
                                f_bld_ole_chart =
                                    Some(Box::new(CTTLOleBuildChart::from_xml(reader, &e, true)?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            b"bldGraphic" => {
                                f_bld_graphic = Some(Box::new(CTTLGraphicalObjectBuild::from_xml(
                                    reader, &e, true,
                                )?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            #[cfg(feature = "extra-children")]
                            _ => {
                                // Capture unknown empty element for roundtrip
                                let elem = RawXmlElement::from_empty(&e);
                                extra_children.push(PositionedNode::new(
                                    child_idx,
                                    RawXmlNode::Element(elem),
                                ));
                                child_idx += 1;
                            }
                            #[cfg(not(feature = "extra-children"))]
                            _ => {}
                        }
                    }
                    Event::End(_) => break,
                    Event::Eof => break,
                    _ => {}
                }
                buf.clear();
            }
        }

        Ok(Self {
            bld_p: f_bld_p,
            bld_dgm: f_bld_dgm,
            bld_ole_chart: f_bld_ole_chart,
            bld_graphic: f_bld_graphic,
            #[cfg(feature = "extra-children")]
            extra_children,
        })
    }
}

impl FromXml for SlideTiming {
    fn from_xml<R: BufRead>(
        reader: &mut Reader<R>,
        start_tag: &BytesStart,
        is_empty: bool,
    ) -> Result<Self, ParseError> {
        let mut f_tn_lst = None;
        let mut f_bld_lst = None;
        let mut f_ext_lst = None;
        #[cfg(feature = "extra-children")]
        let mut extra_children = Vec::new();
        #[cfg(feature = "extra-children")]
        let mut child_idx: usize = 0;

        // Parse child elements
        if !is_empty {
            let mut buf = Vec::new();
            loop {
                match reader.read_event_into(&mut buf)? {
                    Event::Start(e) => {
                        match e.local_name().as_ref() {
                            b"tnLst" => {
                                f_tn_lst =
                                    Some(Box::new(CTTimeNodeList::from_xml(reader, &e, false)?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            b"bldLst" => {
                                f_bld_lst =
                                    Some(Box::new(CTBuildList::from_xml(reader, &e, false)?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            b"extLst" => {
                                f_ext_lst = Some(Box::new(CTExtensionListModify::from_xml(
                                    reader, &e, false,
                                )?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            #[cfg(feature = "extra-children")]
                            _ => {
                                // Capture unknown element for roundtrip
                                let elem = RawXmlElement::from_reader(reader, &e)?;
                                extra_children.push(PositionedNode::new(
                                    child_idx,
                                    RawXmlNode::Element(elem),
                                ));
                                child_idx += 1;
                            }
                            #[cfg(not(feature = "extra-children"))]
                            _ => {
                                // Skip unknown element
                                skip_element(reader)?;
                            }
                        }
                    }
                    Event::Empty(e) => {
                        match e.local_name().as_ref() {
                            b"tnLst" => {
                                f_tn_lst =
                                    Some(Box::new(CTTimeNodeList::from_xml(reader, &e, true)?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            b"bldLst" => {
                                f_bld_lst =
                                    Some(Box::new(CTBuildList::from_xml(reader, &e, true)?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            b"extLst" => {
                                f_ext_lst = Some(Box::new(CTExtensionListModify::from_xml(
                                    reader, &e, true,
                                )?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            #[cfg(feature = "extra-children")]
                            _ => {
                                // Capture unknown empty element for roundtrip
                                let elem = RawXmlElement::from_empty(&e);
                                extra_children.push(PositionedNode::new(
                                    child_idx,
                                    RawXmlNode::Element(elem),
                                ));
                                child_idx += 1;
                            }
                            #[cfg(not(feature = "extra-children"))]
                            _ => {}
                        }
                    }
                    Event::End(_) => break,
                    Event::Eof => break,
                    _ => {}
                }
                buf.clear();
            }
        }

        Ok(Self {
            tn_lst: f_tn_lst,
            bld_lst: f_bld_lst,
            ext_lst: f_ext_lst,
            #[cfg(feature = "extra-children")]
            extra_children,
        })
    }
}

impl FromXml for CTEmpty {
    fn from_xml<R: BufRead>(
        reader: &mut Reader<R>,
        _start: &BytesStart,
        is_empty: bool,
    ) -> Result<Self, ParseError> {
        if !is_empty {
            let mut buf = Vec::new();
            let mut depth = 1u32;
            loop {
                match reader.read_event_into(&mut buf)? {
                    Event::Start(_) => depth += 1,
                    Event::End(_) => {
                        depth -= 1;
                        if depth == 0 {
                            break;
                        }
                    }
                    Event::Eof => break,
                    _ => {}
                }
                buf.clear();
            }
        }
        Ok(Self {})
    }
}

impl FromXml for CTIndexRange {
    fn from_xml<R: BufRead>(
        reader: &mut Reader<R>,
        start_tag: &BytesStart,
        is_empty: bool,
    ) -> Result<Self, ParseError> {
        let mut f_st: Option<STIndex> = None;
        let mut f_end: Option<STIndex> = None;
        #[cfg(feature = "extra-attrs")]
        let mut extra_attrs = std::collections::HashMap::new();

        // Parse attributes
        for attr in start_tag.attributes().filter_map(|a| a.ok()) {
            let val = String::from_utf8_lossy(&attr.value);
            match attr.key.local_name().as_ref() {
                b"st" => {
                    f_st = val.parse().ok();
                }
                b"end" => {
                    f_end = val.parse().ok();
                }
                #[cfg(feature = "extra-attrs")]
                unknown => {
                    let key = String::from_utf8_lossy(attr.key.as_ref()).into_owned();
                    extra_attrs.insert(key, val.into_owned());
                }
                #[cfg(not(feature = "extra-attrs"))]
                _ => {}
            }
        }

        if !is_empty {
            let mut buf = Vec::new();
            loop {
                match reader.read_event_into(&mut buf)? {
                    Event::End(_) => break,
                    Event::Eof => break,
                    _ => {}
                }
                buf.clear();
            }
        }

        Ok(Self {
            st: f_st.ok_or_else(|| ParseError::MissingAttribute("st".to_string()))?,
            end: f_end.ok_or_else(|| ParseError::MissingAttribute("end".to_string()))?,
            #[cfg(feature = "extra-attrs")]
            extra_attrs,
        })
    }
}

impl FromXml for CTSlideRelationshipListEntry {
    fn from_xml<R: BufRead>(
        reader: &mut Reader<R>,
        start_tag: &BytesStart,
        is_empty: bool,
    ) -> Result<Self, ParseError> {
        #[cfg(feature = "extra-children")]
        let mut extra_children = Vec::new();
        #[cfg(feature = "extra-children")]
        let mut child_idx: usize = 0;
        if !is_empty {
            let mut buf = Vec::new();
            loop {
                match reader.read_event_into(&mut buf)? {
                    #[cfg(feature = "extra-children")]
                    Event::Start(e) => {
                        let elem = RawXmlElement::from_reader(reader, &e)?;
                        extra_children
                            .push(PositionedNode::new(child_idx, RawXmlNode::Element(elem)));
                        child_idx += 1;
                    }
                    #[cfg(not(feature = "extra-children"))]
                    Event::Start(_) => {
                        skip_element(reader)?;
                    }
                    #[cfg(feature = "extra-children")]
                    Event::Empty(e) => {
                        let elem = RawXmlElement::from_empty(&e);
                        extra_children
                            .push(PositionedNode::new(child_idx, RawXmlNode::Element(elem)));
                        child_idx += 1;
                    }
                    #[cfg(not(feature = "extra-children"))]
                    Event::Empty(_) => {}
                    Event::End(_) => break,
                    Event::Eof => break,
                    _ => {}
                }
                buf.clear();
            }
        }
        Ok(Self {
            #[cfg(feature = "extra-children")]
            extra_children,
        })
    }
}

impl FromXml for CTSlideRelationshipList {
    fn from_xml<R: BufRead>(
        reader: &mut Reader<R>,
        start_tag: &BytesStart,
        is_empty: bool,
    ) -> Result<Self, ParseError> {
        let mut f_sld = Vec::new();
        #[cfg(feature = "extra-children")]
        let mut extra_children = Vec::new();
        #[cfg(feature = "extra-children")]
        let mut child_idx: usize = 0;

        // Parse child elements
        if !is_empty {
            let mut buf = Vec::new();
            loop {
                match reader.read_event_into(&mut buf)? {
                    Event::Start(e) => {
                        match e.local_name().as_ref() {
                            b"sld" => {
                                f_sld.push(CTSlideRelationshipListEntry::from_xml(
                                    reader, &e, false,
                                )?);
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            #[cfg(feature = "extra-children")]
                            _ => {
                                // Capture unknown element for roundtrip
                                let elem = RawXmlElement::from_reader(reader, &e)?;
                                extra_children.push(PositionedNode::new(
                                    child_idx,
                                    RawXmlNode::Element(elem),
                                ));
                                child_idx += 1;
                            }
                            #[cfg(not(feature = "extra-children"))]
                            _ => {
                                // Skip unknown element
                                skip_element(reader)?;
                            }
                        }
                    }
                    Event::Empty(e) => {
                        match e.local_name().as_ref() {
                            b"sld" => {
                                f_sld.push(CTSlideRelationshipListEntry::from_xml(
                                    reader, &e, true,
                                )?);
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            #[cfg(feature = "extra-children")]
                            _ => {
                                // Capture unknown empty element for roundtrip
                                let elem = RawXmlElement::from_empty(&e);
                                extra_children.push(PositionedNode::new(
                                    child_idx,
                                    RawXmlNode::Element(elem),
                                ));
                                child_idx += 1;
                            }
                            #[cfg(not(feature = "extra-children"))]
                            _ => {}
                        }
                    }
                    Event::End(_) => break,
                    Event::Eof => break,
                    _ => {}
                }
                buf.clear();
            }
        }

        Ok(Self {
            sld: f_sld,
            #[cfg(feature = "extra-children")]
            extra_children,
        })
    }
}

impl FromXml for CTCustomShowId {
    fn from_xml<R: BufRead>(
        reader: &mut Reader<R>,
        start_tag: &BytesStart,
        is_empty: bool,
    ) -> Result<Self, ParseError> {
        let mut f_id: Option<u32> = None;
        #[cfg(feature = "extra-attrs")]
        let mut extra_attrs = std::collections::HashMap::new();

        // Parse attributes
        for attr in start_tag.attributes().filter_map(|a| a.ok()) {
            let val = String::from_utf8_lossy(&attr.value);
            match attr.key.local_name().as_ref() {
                b"id" => {
                    f_id = val.parse().ok();
                }
                #[cfg(feature = "extra-attrs")]
                unknown => {
                    let key = String::from_utf8_lossy(attr.key.as_ref()).into_owned();
                    extra_attrs.insert(key, val.into_owned());
                }
                #[cfg(not(feature = "extra-attrs"))]
                _ => {}
            }
        }

        if !is_empty {
            let mut buf = Vec::new();
            loop {
                match reader.read_event_into(&mut buf)? {
                    Event::End(_) => break,
                    Event::Eof => break,
                    _ => {}
                }
                buf.clear();
            }
        }

        Ok(Self {
            id: f_id.ok_or_else(|| ParseError::MissingAttribute("id".to_string()))?,
            #[cfg(feature = "extra-attrs")]
            extra_attrs,
        })
    }
}

impl FromXml for EGSlideListChoice {
    fn from_xml<R: BufRead>(
        reader: &mut Reader<R>,
        start_tag: &BytesStart,
        is_empty: bool,
    ) -> Result<Self, ParseError> {
        let tag = start_tag.local_name();
        match tag.as_ref() {
            b"sldAll" => {
                let inner = CTEmpty::from_xml(reader, start_tag, is_empty)?;
                Ok(Self::SldAll(Box::new(inner)))
            }
            b"sldRg" => {
                let inner = CTIndexRange::from_xml(reader, start_tag, is_empty)?;
                Ok(Self::SldRg(Box::new(inner)))
            }
            b"custShow" => {
                let inner = CTCustomShowId::from_xml(reader, start_tag, is_empty)?;
                Ok(Self::CustShow(Box::new(inner)))
            }
            _ => Err(ParseError::UnexpectedElement(
                String::from_utf8_lossy(start_tag.name().as_ref()).into_owned(),
            )),
        }
    }
}

impl FromXml for CTCustomerData {
    fn from_xml<R: BufRead>(
        reader: &mut Reader<R>,
        start_tag: &BytesStart,
        is_empty: bool,
    ) -> Result<Self, ParseError> {
        #[cfg(feature = "extra-children")]
        let mut extra_children = Vec::new();
        #[cfg(feature = "extra-children")]
        let mut child_idx: usize = 0;
        if !is_empty {
            let mut buf = Vec::new();
            loop {
                match reader.read_event_into(&mut buf)? {
                    #[cfg(feature = "extra-children")]
                    Event::Start(e) => {
                        let elem = RawXmlElement::from_reader(reader, &e)?;
                        extra_children
                            .push(PositionedNode::new(child_idx, RawXmlNode::Element(elem)));
                        child_idx += 1;
                    }
                    #[cfg(not(feature = "extra-children"))]
                    Event::Start(_) => {
                        skip_element(reader)?;
                    }
                    #[cfg(feature = "extra-children")]
                    Event::Empty(e) => {
                        let elem = RawXmlElement::from_empty(&e);
                        extra_children
                            .push(PositionedNode::new(child_idx, RawXmlNode::Element(elem)));
                        child_idx += 1;
                    }
                    #[cfg(not(feature = "extra-children"))]
                    Event::Empty(_) => {}
                    Event::End(_) => break,
                    Event::Eof => break,
                    _ => {}
                }
                buf.clear();
            }
        }
        Ok(Self {
            #[cfg(feature = "extra-children")]
            extra_children,
        })
    }
}

impl FromXml for CTTagsData {
    fn from_xml<R: BufRead>(
        reader: &mut Reader<R>,
        start_tag: &BytesStart,
        is_empty: bool,
    ) -> Result<Self, ParseError> {
        #[cfg(feature = "extra-children")]
        let mut extra_children = Vec::new();
        #[cfg(feature = "extra-children")]
        let mut child_idx: usize = 0;
        if !is_empty {
            let mut buf = Vec::new();
            loop {
                match reader.read_event_into(&mut buf)? {
                    #[cfg(feature = "extra-children")]
                    Event::Start(e) => {
                        let elem = RawXmlElement::from_reader(reader, &e)?;
                        extra_children
                            .push(PositionedNode::new(child_idx, RawXmlNode::Element(elem)));
                        child_idx += 1;
                    }
                    #[cfg(not(feature = "extra-children"))]
                    Event::Start(_) => {
                        skip_element(reader)?;
                    }
                    #[cfg(feature = "extra-children")]
                    Event::Empty(e) => {
                        let elem = RawXmlElement::from_empty(&e);
                        extra_children
                            .push(PositionedNode::new(child_idx, RawXmlNode::Element(elem)));
                        child_idx += 1;
                    }
                    #[cfg(not(feature = "extra-children"))]
                    Event::Empty(_) => {}
                    Event::End(_) => break,
                    Event::Eof => break,
                    _ => {}
                }
                buf.clear();
            }
        }
        Ok(Self {
            #[cfg(feature = "extra-children")]
            extra_children,
        })
    }
}

impl FromXml for CTCustomerDataList {
    fn from_xml<R: BufRead>(
        reader: &mut Reader<R>,
        start_tag: &BytesStart,
        is_empty: bool,
    ) -> Result<Self, ParseError> {
        let mut f_cust_data = Vec::new();
        let mut f_tags = None;
        #[cfg(feature = "extra-children")]
        let mut extra_children = Vec::new();
        #[cfg(feature = "extra-children")]
        let mut child_idx: usize = 0;

        // Parse child elements
        if !is_empty {
            let mut buf = Vec::new();
            loop {
                match reader.read_event_into(&mut buf)? {
                    Event::Start(e) => {
                        match e.local_name().as_ref() {
                            b"custData" => {
                                f_cust_data.push(CTCustomerData::from_xml(reader, &e, false)?);
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            b"tags" => {
                                f_tags = Some(Box::new(CTTagsData::from_xml(reader, &e, false)?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            #[cfg(feature = "extra-children")]
                            _ => {
                                // Capture unknown element for roundtrip
                                let elem = RawXmlElement::from_reader(reader, &e)?;
                                extra_children.push(PositionedNode::new(
                                    child_idx,
                                    RawXmlNode::Element(elem),
                                ));
                                child_idx += 1;
                            }
                            #[cfg(not(feature = "extra-children"))]
                            _ => {
                                // Skip unknown element
                                skip_element(reader)?;
                            }
                        }
                    }
                    Event::Empty(e) => {
                        match e.local_name().as_ref() {
                            b"custData" => {
                                f_cust_data.push(CTCustomerData::from_xml(reader, &e, true)?);
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            b"tags" => {
                                f_tags = Some(Box::new(CTTagsData::from_xml(reader, &e, true)?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            #[cfg(feature = "extra-children")]
                            _ => {
                                // Capture unknown empty element for roundtrip
                                let elem = RawXmlElement::from_empty(&e);
                                extra_children.push(PositionedNode::new(
                                    child_idx,
                                    RawXmlNode::Element(elem),
                                ));
                                child_idx += 1;
                            }
                            #[cfg(not(feature = "extra-children"))]
                            _ => {}
                        }
                    }
                    Event::End(_) => break,
                    Event::Eof => break,
                    _ => {}
                }
                buf.clear();
            }
        }

        Ok(Self {
            cust_data: f_cust_data,
            tags: f_tags,
            #[cfg(feature = "extra-children")]
            extra_children,
        })
    }
}

impl FromXml for CTExtension {
    fn from_xml<R: BufRead>(
        reader: &mut Reader<R>,
        start_tag: &BytesStart,
        is_empty: bool,
    ) -> Result<Self, ParseError> {
        let mut f_uri: Option<String> = None;
        #[cfg(feature = "extra-attrs")]
        let mut extra_attrs = std::collections::HashMap::new();

        // Parse attributes
        for attr in start_tag.attributes().filter_map(|a| a.ok()) {
            let val = String::from_utf8_lossy(&attr.value);
            match attr.key.local_name().as_ref() {
                b"uri" => {
                    f_uri = Some(val.into_owned());
                }
                #[cfg(feature = "extra-attrs")]
                unknown => {
                    let key = String::from_utf8_lossy(attr.key.as_ref()).into_owned();
                    extra_attrs.insert(key, val.into_owned());
                }
                #[cfg(not(feature = "extra-attrs"))]
                _ => {}
            }
        }

        if !is_empty {
            let mut buf = Vec::new();
            loop {
                match reader.read_event_into(&mut buf)? {
                    Event::End(_) => break,
                    Event::Eof => break,
                    _ => {}
                }
                buf.clear();
            }
        }

        Ok(Self {
            uri: f_uri.ok_or_else(|| ParseError::MissingAttribute("uri".to_string()))?,
            #[cfg(feature = "extra-attrs")]
            extra_attrs,
        })
    }
}

impl FromXml for EGExtensionList {
    fn from_xml<R: BufRead>(
        reader: &mut Reader<R>,
        start_tag: &BytesStart,
        is_empty: bool,
    ) -> Result<Self, ParseError> {
        let mut f_ext = Vec::new();
        #[cfg(feature = "extra-children")]
        let mut extra_children = Vec::new();
        #[cfg(feature = "extra-children")]
        let mut child_idx: usize = 0;

        // Parse child elements
        if !is_empty {
            let mut buf = Vec::new();
            loop {
                match reader.read_event_into(&mut buf)? {
                    Event::Start(e) => {
                        match e.local_name().as_ref() {
                            b"ext" => {
                                f_ext.push(CTExtension::from_xml(reader, &e, false)?);
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            #[cfg(feature = "extra-children")]
                            _ => {
                                // Capture unknown element for roundtrip
                                let elem = RawXmlElement::from_reader(reader, &e)?;
                                extra_children.push(PositionedNode::new(
                                    child_idx,
                                    RawXmlNode::Element(elem),
                                ));
                                child_idx += 1;
                            }
                            #[cfg(not(feature = "extra-children"))]
                            _ => {
                                // Skip unknown element
                                skip_element(reader)?;
                            }
                        }
                    }
                    Event::Empty(e) => {
                        match e.local_name().as_ref() {
                            b"ext" => {
                                f_ext.push(CTExtension::from_xml(reader, &e, true)?);
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            #[cfg(feature = "extra-children")]
                            _ => {
                                // Capture unknown empty element for roundtrip
                                let elem = RawXmlElement::from_empty(&e);
                                extra_children.push(PositionedNode::new(
                                    child_idx,
                                    RawXmlNode::Element(elem),
                                ));
                                child_idx += 1;
                            }
                            #[cfg(not(feature = "extra-children"))]
                            _ => {}
                        }
                    }
                    Event::End(_) => break,
                    Event::Eof => break,
                    _ => {}
                }
                buf.clear();
            }
        }

        Ok(Self {
            ext: f_ext,
            #[cfg(feature = "extra-children")]
            extra_children,
        })
    }
}

impl FromXml for CTExtensionList {
    fn from_xml<R: BufRead>(
        reader: &mut Reader<R>,
        start_tag: &BytesStart,
        is_empty: bool,
    ) -> Result<Self, ParseError> {
        let mut f_ext = Vec::new();
        #[cfg(feature = "extra-children")]
        let mut extra_children = Vec::new();
        #[cfg(feature = "extra-children")]
        let mut child_idx: usize = 0;

        // Parse child elements
        if !is_empty {
            let mut buf = Vec::new();
            loop {
                match reader.read_event_into(&mut buf)? {
                    Event::Start(e) => {
                        match e.local_name().as_ref() {
                            b"ext" => {
                                f_ext.push(CTExtension::from_xml(reader, &e, false)?);
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            #[cfg(feature = "extra-children")]
                            _ => {
                                // Capture unknown element for roundtrip
                                let elem = RawXmlElement::from_reader(reader, &e)?;
                                extra_children.push(PositionedNode::new(
                                    child_idx,
                                    RawXmlNode::Element(elem),
                                ));
                                child_idx += 1;
                            }
                            #[cfg(not(feature = "extra-children"))]
                            _ => {
                                // Skip unknown element
                                skip_element(reader)?;
                            }
                        }
                    }
                    Event::Empty(e) => {
                        match e.local_name().as_ref() {
                            b"ext" => {
                                f_ext.push(CTExtension::from_xml(reader, &e, true)?);
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            #[cfg(feature = "extra-children")]
                            _ => {
                                // Capture unknown empty element for roundtrip
                                let elem = RawXmlElement::from_empty(&e);
                                extra_children.push(PositionedNode::new(
                                    child_idx,
                                    RawXmlNode::Element(elem),
                                ));
                                child_idx += 1;
                            }
                            #[cfg(not(feature = "extra-children"))]
                            _ => {}
                        }
                    }
                    Event::End(_) => break,
                    Event::Eof => break,
                    _ => {}
                }
                buf.clear();
            }
        }

        Ok(Self {
            ext: f_ext,
            #[cfg(feature = "extra-children")]
            extra_children,
        })
    }
}

impl FromXml for CTExtensionListModify {
    fn from_xml<R: BufRead>(
        reader: &mut Reader<R>,
        start_tag: &BytesStart,
        is_empty: bool,
    ) -> Result<Self, ParseError> {
        let mut f_mod = None;
        let mut f_ext = Vec::new();
        #[cfg(feature = "extra-attrs")]
        let mut extra_attrs = std::collections::HashMap::new();
        #[cfg(feature = "extra-children")]
        let mut extra_children = Vec::new();
        #[cfg(feature = "extra-children")]
        let mut child_idx: usize = 0;

        // Parse attributes
        for attr in start_tag.attributes().filter_map(|a| a.ok()) {
            let val = String::from_utf8_lossy(&attr.value);
            match attr.key.local_name().as_ref() {
                b"mod" => {
                    f_mod = Some(val == "true" || val == "1");
                }
                #[cfg(feature = "extra-attrs")]
                unknown => {
                    let key = String::from_utf8_lossy(attr.key.as_ref()).into_owned();
                    extra_attrs.insert(key, val.into_owned());
                }
                #[cfg(not(feature = "extra-attrs"))]
                _ => {}
            }
        }

        // Parse child elements
        if !is_empty {
            let mut buf = Vec::new();
            loop {
                match reader.read_event_into(&mut buf)? {
                    Event::Start(e) => {
                        match e.local_name().as_ref() {
                            b"ext" => {
                                f_ext.push(CTExtension::from_xml(reader, &e, false)?);
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            #[cfg(feature = "extra-children")]
                            _ => {
                                // Capture unknown element for roundtrip
                                let elem = RawXmlElement::from_reader(reader, &e)?;
                                extra_children.push(PositionedNode::new(
                                    child_idx,
                                    RawXmlNode::Element(elem),
                                ));
                                child_idx += 1;
                            }
                            #[cfg(not(feature = "extra-children"))]
                            _ => {
                                // Skip unknown element
                                skip_element(reader)?;
                            }
                        }
                    }
                    Event::Empty(e) => {
                        match e.local_name().as_ref() {
                            b"ext" => {
                                f_ext.push(CTExtension::from_xml(reader, &e, true)?);
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            #[cfg(feature = "extra-children")]
                            _ => {
                                // Capture unknown empty element for roundtrip
                                let elem = RawXmlElement::from_empty(&e);
                                extra_children.push(PositionedNode::new(
                                    child_idx,
                                    RawXmlNode::Element(elem),
                                ));
                                child_idx += 1;
                            }
                            #[cfg(not(feature = "extra-children"))]
                            _ => {}
                        }
                    }
                    Event::End(_) => break,
                    Event::Eof => break,
                    _ => {}
                }
                buf.clear();
            }
        }

        Ok(Self {
            r#mod: f_mod,
            ext: f_ext,
            #[cfg(feature = "extra-attrs")]
            extra_attrs,
            #[cfg(feature = "extra-children")]
            extra_children,
        })
    }
}

impl FromXml for CTCommentAuthor {
    fn from_xml<R: BufRead>(
        reader: &mut Reader<R>,
        start_tag: &BytesStart,
        is_empty: bool,
    ) -> Result<Self, ParseError> {
        let mut f_id: Option<u32> = None;
        let mut f_name: Option<STName> = None;
        let mut f_initials: Option<STName> = None;
        let mut f_last_idx: Option<u32> = None;
        let mut f_clr_idx: Option<u32> = None;
        let mut f_ext_lst = None;
        #[cfg(feature = "extra-attrs")]
        let mut extra_attrs = std::collections::HashMap::new();
        #[cfg(feature = "extra-children")]
        let mut extra_children = Vec::new();
        #[cfg(feature = "extra-children")]
        let mut child_idx: usize = 0;

        // Parse attributes
        for attr in start_tag.attributes().filter_map(|a| a.ok()) {
            let val = String::from_utf8_lossy(&attr.value);
            match attr.key.local_name().as_ref() {
                b"id" => {
                    f_id = val.parse().ok();
                }
                b"name" => {
                    f_name = Some(val.into_owned());
                }
                b"initials" => {
                    f_initials = Some(val.into_owned());
                }
                b"lastIdx" => {
                    f_last_idx = val.parse().ok();
                }
                b"clrIdx" => {
                    f_clr_idx = val.parse().ok();
                }
                #[cfg(feature = "extra-attrs")]
                unknown => {
                    let key = String::from_utf8_lossy(attr.key.as_ref()).into_owned();
                    extra_attrs.insert(key, val.into_owned());
                }
                #[cfg(not(feature = "extra-attrs"))]
                _ => {}
            }
        }

        // Parse child elements
        if !is_empty {
            let mut buf = Vec::new();
            loop {
                match reader.read_event_into(&mut buf)? {
                    Event::Start(e) => {
                        match e.local_name().as_ref() {
                            b"extLst" => {
                                f_ext_lst =
                                    Some(Box::new(CTExtensionList::from_xml(reader, &e, false)?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            #[cfg(feature = "extra-children")]
                            _ => {
                                // Capture unknown element for roundtrip
                                let elem = RawXmlElement::from_reader(reader, &e)?;
                                extra_children.push(PositionedNode::new(
                                    child_idx,
                                    RawXmlNode::Element(elem),
                                ));
                                child_idx += 1;
                            }
                            #[cfg(not(feature = "extra-children"))]
                            _ => {
                                // Skip unknown element
                                skip_element(reader)?;
                            }
                        }
                    }
                    Event::Empty(e) => {
                        match e.local_name().as_ref() {
                            b"extLst" => {
                                f_ext_lst =
                                    Some(Box::new(CTExtensionList::from_xml(reader, &e, true)?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            #[cfg(feature = "extra-children")]
                            _ => {
                                // Capture unknown empty element for roundtrip
                                let elem = RawXmlElement::from_empty(&e);
                                extra_children.push(PositionedNode::new(
                                    child_idx,
                                    RawXmlNode::Element(elem),
                                ));
                                child_idx += 1;
                            }
                            #[cfg(not(feature = "extra-children"))]
                            _ => {}
                        }
                    }
                    Event::End(_) => break,
                    Event::Eof => break,
                    _ => {}
                }
                buf.clear();
            }
        }

        Ok(Self {
            id: f_id.ok_or_else(|| ParseError::MissingAttribute("id".to_string()))?,
            name: f_name.ok_or_else(|| ParseError::MissingAttribute("name".to_string()))?,
            initials: f_initials
                .ok_or_else(|| ParseError::MissingAttribute("initials".to_string()))?,
            last_idx: f_last_idx
                .ok_or_else(|| ParseError::MissingAttribute("lastIdx".to_string()))?,
            clr_idx: f_clr_idx.ok_or_else(|| ParseError::MissingAttribute("clrIdx".to_string()))?,
            ext_lst: f_ext_lst,
            #[cfg(feature = "extra-attrs")]
            extra_attrs,
            #[cfg(feature = "extra-children")]
            extra_children,
        })
    }
}

impl FromXml for CTCommentAuthorList {
    fn from_xml<R: BufRead>(
        reader: &mut Reader<R>,
        start_tag: &BytesStart,
        is_empty: bool,
    ) -> Result<Self, ParseError> {
        let mut f_cm_author = Vec::new();
        #[cfg(feature = "extra-children")]
        let mut extra_children = Vec::new();
        #[cfg(feature = "extra-children")]
        let mut child_idx: usize = 0;

        // Parse child elements
        if !is_empty {
            let mut buf = Vec::new();
            loop {
                match reader.read_event_into(&mut buf)? {
                    Event::Start(e) => {
                        match e.local_name().as_ref() {
                            b"cmAuthor" => {
                                f_cm_author.push(CTCommentAuthor::from_xml(reader, &e, false)?);
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            #[cfg(feature = "extra-children")]
                            _ => {
                                // Capture unknown element for roundtrip
                                let elem = RawXmlElement::from_reader(reader, &e)?;
                                extra_children.push(PositionedNode::new(
                                    child_idx,
                                    RawXmlNode::Element(elem),
                                ));
                                child_idx += 1;
                            }
                            #[cfg(not(feature = "extra-children"))]
                            _ => {
                                // Skip unknown element
                                skip_element(reader)?;
                            }
                        }
                    }
                    Event::Empty(e) => {
                        match e.local_name().as_ref() {
                            b"cmAuthor" => {
                                f_cm_author.push(CTCommentAuthor::from_xml(reader, &e, true)?);
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            #[cfg(feature = "extra-children")]
                            _ => {
                                // Capture unknown empty element for roundtrip
                                let elem = RawXmlElement::from_empty(&e);
                                extra_children.push(PositionedNode::new(
                                    child_idx,
                                    RawXmlNode::Element(elem),
                                ));
                                child_idx += 1;
                            }
                            #[cfg(not(feature = "extra-children"))]
                            _ => {}
                        }
                    }
                    Event::End(_) => break,
                    Event::Eof => break,
                    _ => {}
                }
                buf.clear();
            }
        }

        Ok(Self {
            cm_author: f_cm_author,
            #[cfg(feature = "extra-children")]
            extra_children,
        })
    }
}

impl FromXml for CTComment {
    fn from_xml<R: BufRead>(
        reader: &mut Reader<R>,
        start_tag: &BytesStart,
        is_empty: bool,
    ) -> Result<Self, ParseError> {
        let mut f_author_id: Option<u32> = None;
        let mut f_dt = None;
        let mut f_idx: Option<STIndex> = None;
        let mut f_pos: Option<Box<ooxml_dml::types::Point2D>> = None;
        let mut f_text: Option<String> = None;
        let mut f_ext_lst = None;
        #[cfg(feature = "extra-attrs")]
        let mut extra_attrs = std::collections::HashMap::new();
        #[cfg(feature = "extra-children")]
        let mut extra_children = Vec::new();
        #[cfg(feature = "extra-children")]
        let mut child_idx: usize = 0;

        // Parse attributes
        for attr in start_tag.attributes().filter_map(|a| a.ok()) {
            let val = String::from_utf8_lossy(&attr.value);
            match attr.key.local_name().as_ref() {
                b"authorId" => {
                    f_author_id = val.parse().ok();
                }
                b"dt" => {
                    f_dt = Some(val.into_owned());
                }
                b"idx" => {
                    f_idx = val.parse().ok();
                }
                #[cfg(feature = "extra-attrs")]
                unknown => {
                    let key = String::from_utf8_lossy(attr.key.as_ref()).into_owned();
                    extra_attrs.insert(key, val.into_owned());
                }
                #[cfg(not(feature = "extra-attrs"))]
                _ => {}
            }
        }

        // Parse child elements
        if !is_empty {
            let mut buf = Vec::new();
            loop {
                match reader.read_event_into(&mut buf)? {
                    Event::Start(e) => {
                        match e.local_name().as_ref() {
                            b"pos" => {
                                f_pos = Some(Box::new(ooxml_dml::types::Point2D::from_xml(
                                    reader, &e, false,
                                )?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            b"text" => {
                                f_text = Some(read_text_content(reader)?);
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            b"extLst" => {
                                f_ext_lst = Some(Box::new(CTExtensionListModify::from_xml(
                                    reader, &e, false,
                                )?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            #[cfg(feature = "extra-children")]
                            _ => {
                                // Capture unknown element for roundtrip
                                let elem = RawXmlElement::from_reader(reader, &e)?;
                                extra_children.push(PositionedNode::new(
                                    child_idx,
                                    RawXmlNode::Element(elem),
                                ));
                                child_idx += 1;
                            }
                            #[cfg(not(feature = "extra-children"))]
                            _ => {
                                // Skip unknown element
                                skip_element(reader)?;
                            }
                        }
                    }
                    Event::Empty(e) => {
                        match e.local_name().as_ref() {
                            b"pos" => {
                                f_pos = Some(Box::new(ooxml_dml::types::Point2D::from_xml(
                                    reader, &e, true,
                                )?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            b"text" => {
                                f_text = Some(String::new());
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            b"extLst" => {
                                f_ext_lst = Some(Box::new(CTExtensionListModify::from_xml(
                                    reader, &e, true,
                                )?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            #[cfg(feature = "extra-children")]
                            _ => {
                                // Capture unknown empty element for roundtrip
                                let elem = RawXmlElement::from_empty(&e);
                                extra_children.push(PositionedNode::new(
                                    child_idx,
                                    RawXmlNode::Element(elem),
                                ));
                                child_idx += 1;
                            }
                            #[cfg(not(feature = "extra-children"))]
                            _ => {}
                        }
                    }
                    Event::End(_) => break,
                    Event::Eof => break,
                    _ => {}
                }
                buf.clear();
            }
        }

        Ok(Self {
            author_id: f_author_id
                .ok_or_else(|| ParseError::MissingAttribute("authorId".to_string()))?,
            dt: f_dt,
            idx: f_idx.ok_or_else(|| ParseError::MissingAttribute("idx".to_string()))?,
            pos: f_pos.ok_or_else(|| ParseError::MissingAttribute("pos".to_string()))?,
            text: f_text.ok_or_else(|| ParseError::MissingAttribute("text".to_string()))?,
            ext_lst: f_ext_lst,
            #[cfg(feature = "extra-attrs")]
            extra_attrs,
            #[cfg(feature = "extra-children")]
            extra_children,
        })
    }
}

impl FromXml for CTCommentList {
    fn from_xml<R: BufRead>(
        reader: &mut Reader<R>,
        start_tag: &BytesStart,
        is_empty: bool,
    ) -> Result<Self, ParseError> {
        let mut f_cm = Vec::new();
        #[cfg(feature = "extra-children")]
        let mut extra_children = Vec::new();
        #[cfg(feature = "extra-children")]
        let mut child_idx: usize = 0;

        // Parse child elements
        if !is_empty {
            let mut buf = Vec::new();
            loop {
                match reader.read_event_into(&mut buf)? {
                    Event::Start(e) => {
                        match e.local_name().as_ref() {
                            b"cm" => {
                                f_cm.push(CTComment::from_xml(reader, &e, false)?);
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            #[cfg(feature = "extra-children")]
                            _ => {
                                // Capture unknown element for roundtrip
                                let elem = RawXmlElement::from_reader(reader, &e)?;
                                extra_children.push(PositionedNode::new(
                                    child_idx,
                                    RawXmlNode::Element(elem),
                                ));
                                child_idx += 1;
                            }
                            #[cfg(not(feature = "extra-children"))]
                            _ => {
                                // Skip unknown element
                                skip_element(reader)?;
                            }
                        }
                    }
                    Event::Empty(e) => {
                        match e.local_name().as_ref() {
                            b"cm" => {
                                f_cm.push(CTComment::from_xml(reader, &e, true)?);
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            #[cfg(feature = "extra-children")]
                            _ => {
                                // Capture unknown empty element for roundtrip
                                let elem = RawXmlElement::from_empty(&e);
                                extra_children.push(PositionedNode::new(
                                    child_idx,
                                    RawXmlNode::Element(elem),
                                ));
                                child_idx += 1;
                            }
                            #[cfg(not(feature = "extra-children"))]
                            _ => {}
                        }
                    }
                    Event::End(_) => break,
                    Event::Eof => break,
                    _ => {}
                }
                buf.clear();
            }
        }

        Ok(Self {
            cm: f_cm,
            #[cfg(feature = "extra-children")]
            extra_children,
        })
    }
}

impl FromXml for PAGOle {
    fn from_xml<R: BufRead>(
        reader: &mut Reader<R>,
        start_tag: &BytesStart,
        is_empty: bool,
    ) -> Result<Self, ParseError> {
        let mut f_name = None;
        let mut f_show_as_icon = None;
        let mut f_img_w = None;
        let mut f_img_h = None;
        #[cfg(feature = "extra-attrs")]
        let mut extra_attrs = std::collections::HashMap::new();

        // Parse attributes
        for attr in start_tag.attributes().filter_map(|a| a.ok()) {
            let val = String::from_utf8_lossy(&attr.value);
            match attr.key.local_name().as_ref() {
                b"name" => {
                    f_name = Some(val.into_owned());
                }
                b"showAsIcon" => {
                    f_show_as_icon = Some(val == "true" || val == "1");
                }
                b"imgW" => {
                    f_img_w = val.parse().ok();
                }
                b"imgH" => {
                    f_img_h = val.parse().ok();
                }
                #[cfg(feature = "extra-attrs")]
                unknown => {
                    let key = String::from_utf8_lossy(attr.key.as_ref()).into_owned();
                    extra_attrs.insert(key, val.into_owned());
                }
                #[cfg(not(feature = "extra-attrs"))]
                _ => {}
            }
        }

        if !is_empty {
            let mut buf = Vec::new();
            loop {
                match reader.read_event_into(&mut buf)? {
                    Event::End(_) => break,
                    Event::Eof => break,
                    _ => {}
                }
                buf.clear();
            }
        }

        Ok(Self {
            name: f_name,
            show_as_icon: f_show_as_icon,
            img_w: f_img_w,
            img_h: f_img_h,
            #[cfg(feature = "extra-attrs")]
            extra_attrs,
        })
    }
}

impl FromXml for CTOleObjectEmbed {
    fn from_xml<R: BufRead>(
        reader: &mut Reader<R>,
        start_tag: &BytesStart,
        is_empty: bool,
    ) -> Result<Self, ParseError> {
        let mut f_follow_color_scheme = None;
        let mut f_ext_lst = None;
        #[cfg(feature = "extra-attrs")]
        let mut extra_attrs = std::collections::HashMap::new();
        #[cfg(feature = "extra-children")]
        let mut extra_children = Vec::new();
        #[cfg(feature = "extra-children")]
        let mut child_idx: usize = 0;

        // Parse attributes
        for attr in start_tag.attributes().filter_map(|a| a.ok()) {
            let val = String::from_utf8_lossy(&attr.value);
            match attr.key.local_name().as_ref() {
                b"followColorScheme" => {
                    f_follow_color_scheme = val.parse().ok();
                }
                #[cfg(feature = "extra-attrs")]
                unknown => {
                    let key = String::from_utf8_lossy(attr.key.as_ref()).into_owned();
                    extra_attrs.insert(key, val.into_owned());
                }
                #[cfg(not(feature = "extra-attrs"))]
                _ => {}
            }
        }

        // Parse child elements
        if !is_empty {
            let mut buf = Vec::new();
            loop {
                match reader.read_event_into(&mut buf)? {
                    Event::Start(e) => {
                        match e.local_name().as_ref() {
                            b"extLst" => {
                                f_ext_lst =
                                    Some(Box::new(CTExtensionList::from_xml(reader, &e, false)?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            #[cfg(feature = "extra-children")]
                            _ => {
                                // Capture unknown element for roundtrip
                                let elem = RawXmlElement::from_reader(reader, &e)?;
                                extra_children.push(PositionedNode::new(
                                    child_idx,
                                    RawXmlNode::Element(elem),
                                ));
                                child_idx += 1;
                            }
                            #[cfg(not(feature = "extra-children"))]
                            _ => {
                                // Skip unknown element
                                skip_element(reader)?;
                            }
                        }
                    }
                    Event::Empty(e) => {
                        match e.local_name().as_ref() {
                            b"extLst" => {
                                f_ext_lst =
                                    Some(Box::new(CTExtensionList::from_xml(reader, &e, true)?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            #[cfg(feature = "extra-children")]
                            _ => {
                                // Capture unknown empty element for roundtrip
                                let elem = RawXmlElement::from_empty(&e);
                                extra_children.push(PositionedNode::new(
                                    child_idx,
                                    RawXmlNode::Element(elem),
                                ));
                                child_idx += 1;
                            }
                            #[cfg(not(feature = "extra-children"))]
                            _ => {}
                        }
                    }
                    Event::End(_) => break,
                    Event::Eof => break,
                    _ => {}
                }
                buf.clear();
            }
        }

        Ok(Self {
            follow_color_scheme: f_follow_color_scheme,
            ext_lst: f_ext_lst,
            #[cfg(feature = "extra-attrs")]
            extra_attrs,
            #[cfg(feature = "extra-children")]
            extra_children,
        })
    }
}

impl FromXml for CTOleObjectLink {
    fn from_xml<R: BufRead>(
        reader: &mut Reader<R>,
        start_tag: &BytesStart,
        is_empty: bool,
    ) -> Result<Self, ParseError> {
        let mut f_update_automatic = None;
        let mut f_ext_lst = None;
        #[cfg(feature = "extra-attrs")]
        let mut extra_attrs = std::collections::HashMap::new();
        #[cfg(feature = "extra-children")]
        let mut extra_children = Vec::new();
        #[cfg(feature = "extra-children")]
        let mut child_idx: usize = 0;

        // Parse attributes
        for attr in start_tag.attributes().filter_map(|a| a.ok()) {
            let val = String::from_utf8_lossy(&attr.value);
            match attr.key.local_name().as_ref() {
                b"updateAutomatic" => {
                    f_update_automatic = Some(val == "true" || val == "1");
                }
                #[cfg(feature = "extra-attrs")]
                unknown => {
                    let key = String::from_utf8_lossy(attr.key.as_ref()).into_owned();
                    extra_attrs.insert(key, val.into_owned());
                }
                #[cfg(not(feature = "extra-attrs"))]
                _ => {}
            }
        }

        // Parse child elements
        if !is_empty {
            let mut buf = Vec::new();
            loop {
                match reader.read_event_into(&mut buf)? {
                    Event::Start(e) => {
                        match e.local_name().as_ref() {
                            b"extLst" => {
                                f_ext_lst =
                                    Some(Box::new(CTExtensionList::from_xml(reader, &e, false)?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            #[cfg(feature = "extra-children")]
                            _ => {
                                // Capture unknown element for roundtrip
                                let elem = RawXmlElement::from_reader(reader, &e)?;
                                extra_children.push(PositionedNode::new(
                                    child_idx,
                                    RawXmlNode::Element(elem),
                                ));
                                child_idx += 1;
                            }
                            #[cfg(not(feature = "extra-children"))]
                            _ => {
                                // Skip unknown element
                                skip_element(reader)?;
                            }
                        }
                    }
                    Event::Empty(e) => {
                        match e.local_name().as_ref() {
                            b"extLst" => {
                                f_ext_lst =
                                    Some(Box::new(CTExtensionList::from_xml(reader, &e, true)?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            #[cfg(feature = "extra-children")]
                            _ => {
                                // Capture unknown empty element for roundtrip
                                let elem = RawXmlElement::from_empty(&e);
                                extra_children.push(PositionedNode::new(
                                    child_idx,
                                    RawXmlNode::Element(elem),
                                ));
                                child_idx += 1;
                            }
                            #[cfg(not(feature = "extra-children"))]
                            _ => {}
                        }
                    }
                    Event::End(_) => break,
                    Event::Eof => break,
                    _ => {}
                }
                buf.clear();
            }
        }

        Ok(Self {
            update_automatic: f_update_automatic,
            ext_lst: f_ext_lst,
            #[cfg(feature = "extra-attrs")]
            extra_attrs,
            #[cfg(feature = "extra-children")]
            extra_children,
        })
    }
}

impl FromXml for CTOleObject {
    fn from_xml<R: BufRead>(
        reader: &mut Reader<R>,
        start_tag: &BytesStart,
        is_empty: bool,
    ) -> Result<Self, ParseError> {
        let mut f_name = None;
        let mut f_show_as_icon = None;
        let mut f_img_w = None;
        let mut f_img_h = None;
        let mut f_prog_id = None;
        let mut f_embed = None;
        let mut f_link = None;
        let mut f_spid = None;
        let mut f_picture = None;
        #[cfg(feature = "extra-attrs")]
        let mut extra_attrs = std::collections::HashMap::new();
        #[cfg(feature = "extra-children")]
        let mut extra_children = Vec::new();
        #[cfg(feature = "extra-children")]
        let mut child_idx: usize = 0;

        // Parse attributes
        for attr in start_tag.attributes().filter_map(|a| a.ok()) {
            let val = String::from_utf8_lossy(&attr.value);
            match attr.key.local_name().as_ref() {
                b"name" => {
                    f_name = Some(val.into_owned());
                }
                b"showAsIcon" => {
                    f_show_as_icon = Some(val == "true" || val == "1");
                }
                b"imgW" => {
                    f_img_w = val.parse().ok();
                }
                b"imgH" => {
                    f_img_h = val.parse().ok();
                }
                b"progId" => {
                    f_prog_id = Some(val.into_owned());
                }
                b"spid" => {
                    f_spid = val.parse().ok();
                }
                #[cfg(feature = "extra-attrs")]
                unknown => {
                    let key = String::from_utf8_lossy(attr.key.as_ref()).into_owned();
                    extra_attrs.insert(key, val.into_owned());
                }
                #[cfg(not(feature = "extra-attrs"))]
                _ => {}
            }
        }

        // Parse child elements
        if !is_empty {
            let mut buf = Vec::new();
            loop {
                match reader.read_event_into(&mut buf)? {
                    Event::Start(e) => {
                        match e.local_name().as_ref() {
                            b"embed" => {
                                f_embed =
                                    Some(Box::new(CTOleObjectEmbed::from_xml(reader, &e, false)?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            b"link" => {
                                f_link =
                                    Some(Box::new(CTOleObjectLink::from_xml(reader, &e, false)?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            b"pic" => {
                                f_picture = Some(Box::new(Picture::from_xml(reader, &e, false)?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            #[cfg(feature = "extra-children")]
                            _ => {
                                // Capture unknown element for roundtrip
                                let elem = RawXmlElement::from_reader(reader, &e)?;
                                extra_children.push(PositionedNode::new(
                                    child_idx,
                                    RawXmlNode::Element(elem),
                                ));
                                child_idx += 1;
                            }
                            #[cfg(not(feature = "extra-children"))]
                            _ => {
                                // Skip unknown element
                                skip_element(reader)?;
                            }
                        }
                    }
                    Event::Empty(e) => {
                        match e.local_name().as_ref() {
                            b"embed" => {
                                f_embed =
                                    Some(Box::new(CTOleObjectEmbed::from_xml(reader, &e, true)?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            b"link" => {
                                f_link =
                                    Some(Box::new(CTOleObjectLink::from_xml(reader, &e, true)?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            b"pic" => {
                                f_picture = Some(Box::new(Picture::from_xml(reader, &e, true)?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            #[cfg(feature = "extra-children")]
                            _ => {
                                // Capture unknown empty element for roundtrip
                                let elem = RawXmlElement::from_empty(&e);
                                extra_children.push(PositionedNode::new(
                                    child_idx,
                                    RawXmlNode::Element(elem),
                                ));
                                child_idx += 1;
                            }
                            #[cfg(not(feature = "extra-children"))]
                            _ => {}
                        }
                    }
                    Event::End(_) => break,
                    Event::Eof => break,
                    _ => {}
                }
                buf.clear();
            }
        }

        Ok(Self {
            name: f_name,
            show_as_icon: f_show_as_icon,
            img_w: f_img_w,
            img_h: f_img_h,
            prog_id: f_prog_id,
            embed: f_embed,
            link: f_link,
            spid: f_spid,
            picture: f_picture,
            #[cfg(feature = "extra-attrs")]
            extra_attrs,
            #[cfg(feature = "extra-children")]
            extra_children,
        })
    }
}

impl FromXml for CTControl {
    fn from_xml<R: BufRead>(
        reader: &mut Reader<R>,
        start_tag: &BytesStart,
        is_empty: bool,
    ) -> Result<Self, ParseError> {
        let mut f_name = None;
        let mut f_show_as_icon = None;
        let mut f_img_w = None;
        let mut f_img_h = None;
        let mut f_ext_lst = None;
        let mut f_spid = None;
        let mut f_picture = None;
        #[cfg(feature = "extra-attrs")]
        let mut extra_attrs = std::collections::HashMap::new();
        #[cfg(feature = "extra-children")]
        let mut extra_children = Vec::new();
        #[cfg(feature = "extra-children")]
        let mut child_idx: usize = 0;

        // Parse attributes
        for attr in start_tag.attributes().filter_map(|a| a.ok()) {
            let val = String::from_utf8_lossy(&attr.value);
            match attr.key.local_name().as_ref() {
                b"name" => {
                    f_name = Some(val.into_owned());
                }
                b"showAsIcon" => {
                    f_show_as_icon = Some(val == "true" || val == "1");
                }
                b"imgW" => {
                    f_img_w = val.parse().ok();
                }
                b"imgH" => {
                    f_img_h = val.parse().ok();
                }
                b"spid" => {
                    f_spid = val.parse().ok();
                }
                #[cfg(feature = "extra-attrs")]
                unknown => {
                    let key = String::from_utf8_lossy(attr.key.as_ref()).into_owned();
                    extra_attrs.insert(key, val.into_owned());
                }
                #[cfg(not(feature = "extra-attrs"))]
                _ => {}
            }
        }

        // Parse child elements
        if !is_empty {
            let mut buf = Vec::new();
            loop {
                match reader.read_event_into(&mut buf)? {
                    Event::Start(e) => {
                        match e.local_name().as_ref() {
                            b"extLst" => {
                                f_ext_lst =
                                    Some(Box::new(CTExtensionList::from_xml(reader, &e, false)?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            b"pic" => {
                                f_picture = Some(Box::new(Picture::from_xml(reader, &e, false)?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            #[cfg(feature = "extra-children")]
                            _ => {
                                // Capture unknown element for roundtrip
                                let elem = RawXmlElement::from_reader(reader, &e)?;
                                extra_children.push(PositionedNode::new(
                                    child_idx,
                                    RawXmlNode::Element(elem),
                                ));
                                child_idx += 1;
                            }
                            #[cfg(not(feature = "extra-children"))]
                            _ => {
                                // Skip unknown element
                                skip_element(reader)?;
                            }
                        }
                    }
                    Event::Empty(e) => {
                        match e.local_name().as_ref() {
                            b"extLst" => {
                                f_ext_lst =
                                    Some(Box::new(CTExtensionList::from_xml(reader, &e, true)?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            b"pic" => {
                                f_picture = Some(Box::new(Picture::from_xml(reader, &e, true)?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            #[cfg(feature = "extra-children")]
                            _ => {
                                // Capture unknown empty element for roundtrip
                                let elem = RawXmlElement::from_empty(&e);
                                extra_children.push(PositionedNode::new(
                                    child_idx,
                                    RawXmlNode::Element(elem),
                                ));
                                child_idx += 1;
                            }
                            #[cfg(not(feature = "extra-children"))]
                            _ => {}
                        }
                    }
                    Event::End(_) => break,
                    Event::Eof => break,
                    _ => {}
                }
                buf.clear();
            }
        }

        Ok(Self {
            name: f_name,
            show_as_icon: f_show_as_icon,
            img_w: f_img_w,
            img_h: f_img_h,
            ext_lst: f_ext_lst,
            spid: f_spid,
            picture: f_picture,
            #[cfg(feature = "extra-attrs")]
            extra_attrs,
            #[cfg(feature = "extra-children")]
            extra_children,
        })
    }
}

impl FromXml for CTControlList {
    fn from_xml<R: BufRead>(
        reader: &mut Reader<R>,
        start_tag: &BytesStart,
        is_empty: bool,
    ) -> Result<Self, ParseError> {
        let mut f_control = Vec::new();
        #[cfg(feature = "extra-children")]
        let mut extra_children = Vec::new();
        #[cfg(feature = "extra-children")]
        let mut child_idx: usize = 0;

        // Parse child elements
        if !is_empty {
            let mut buf = Vec::new();
            loop {
                match reader.read_event_into(&mut buf)? {
                    Event::Start(e) => {
                        match e.local_name().as_ref() {
                            b"control" => {
                                f_control.push(CTControl::from_xml(reader, &e, false)?);
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            #[cfg(feature = "extra-children")]
                            _ => {
                                // Capture unknown element for roundtrip
                                let elem = RawXmlElement::from_reader(reader, &e)?;
                                extra_children.push(PositionedNode::new(
                                    child_idx,
                                    RawXmlNode::Element(elem),
                                ));
                                child_idx += 1;
                            }
                            #[cfg(not(feature = "extra-children"))]
                            _ => {
                                // Skip unknown element
                                skip_element(reader)?;
                            }
                        }
                    }
                    Event::Empty(e) => {
                        match e.local_name().as_ref() {
                            b"control" => {
                                f_control.push(CTControl::from_xml(reader, &e, true)?);
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            #[cfg(feature = "extra-children")]
                            _ => {
                                // Capture unknown empty element for roundtrip
                                let elem = RawXmlElement::from_empty(&e);
                                extra_children.push(PositionedNode::new(
                                    child_idx,
                                    RawXmlNode::Element(elem),
                                ));
                                child_idx += 1;
                            }
                            #[cfg(not(feature = "extra-children"))]
                            _ => {}
                        }
                    }
                    Event::End(_) => break,
                    Event::Eof => break,
                    _ => {}
                }
                buf.clear();
            }
        }

        Ok(Self {
            control: f_control,
            #[cfg(feature = "extra-children")]
            extra_children,
        })
    }
}

impl FromXml for CTSlideIdListEntry {
    fn from_xml<R: BufRead>(
        reader: &mut Reader<R>,
        start_tag: &BytesStart,
        is_empty: bool,
    ) -> Result<Self, ParseError> {
        let mut f_id: Option<STSlideId> = None;
        let mut f_ext_lst = None;
        #[cfg(feature = "extra-attrs")]
        let mut extra_attrs = std::collections::HashMap::new();
        #[cfg(feature = "extra-children")]
        let mut extra_children = Vec::new();
        #[cfg(feature = "extra-children")]
        let mut child_idx: usize = 0;

        // Parse attributes
        for attr in start_tag.attributes().filter_map(|a| a.ok()) {
            let val = String::from_utf8_lossy(&attr.value);
            match attr.key.local_name().as_ref() {
                b"id" => {
                    f_id = val.parse().ok();
                }
                #[cfg(feature = "extra-attrs")]
                unknown => {
                    let key = String::from_utf8_lossy(attr.key.as_ref()).into_owned();
                    extra_attrs.insert(key, val.into_owned());
                }
                #[cfg(not(feature = "extra-attrs"))]
                _ => {}
            }
        }

        // Parse child elements
        if !is_empty {
            let mut buf = Vec::new();
            loop {
                match reader.read_event_into(&mut buf)? {
                    Event::Start(e) => {
                        match e.local_name().as_ref() {
                            b"extLst" => {
                                f_ext_lst =
                                    Some(Box::new(CTExtensionList::from_xml(reader, &e, false)?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            #[cfg(feature = "extra-children")]
                            _ => {
                                // Capture unknown element for roundtrip
                                let elem = RawXmlElement::from_reader(reader, &e)?;
                                extra_children.push(PositionedNode::new(
                                    child_idx,
                                    RawXmlNode::Element(elem),
                                ));
                                child_idx += 1;
                            }
                            #[cfg(not(feature = "extra-children"))]
                            _ => {
                                // Skip unknown element
                                skip_element(reader)?;
                            }
                        }
                    }
                    Event::Empty(e) => {
                        match e.local_name().as_ref() {
                            b"extLst" => {
                                f_ext_lst =
                                    Some(Box::new(CTExtensionList::from_xml(reader, &e, true)?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            #[cfg(feature = "extra-children")]
                            _ => {
                                // Capture unknown empty element for roundtrip
                                let elem = RawXmlElement::from_empty(&e);
                                extra_children.push(PositionedNode::new(
                                    child_idx,
                                    RawXmlNode::Element(elem),
                                ));
                                child_idx += 1;
                            }
                            #[cfg(not(feature = "extra-children"))]
                            _ => {}
                        }
                    }
                    Event::End(_) => break,
                    Event::Eof => break,
                    _ => {}
                }
                buf.clear();
            }
        }

        Ok(Self {
            id: f_id.ok_or_else(|| ParseError::MissingAttribute("id".to_string()))?,
            ext_lst: f_ext_lst,
            #[cfg(feature = "extra-attrs")]
            extra_attrs,
            #[cfg(feature = "extra-children")]
            extra_children,
        })
    }
}

impl FromXml for SlideIdList {
    fn from_xml<R: BufRead>(
        reader: &mut Reader<R>,
        start_tag: &BytesStart,
        is_empty: bool,
    ) -> Result<Self, ParseError> {
        let mut f_sld_id = Vec::new();
        #[cfg(feature = "extra-children")]
        let mut extra_children = Vec::new();
        #[cfg(feature = "extra-children")]
        let mut child_idx: usize = 0;

        // Parse child elements
        if !is_empty {
            let mut buf = Vec::new();
            loop {
                match reader.read_event_into(&mut buf)? {
                    Event::Start(e) => {
                        match e.local_name().as_ref() {
                            b"sldId" => {
                                f_sld_id.push(CTSlideIdListEntry::from_xml(reader, &e, false)?);
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            #[cfg(feature = "extra-children")]
                            _ => {
                                // Capture unknown element for roundtrip
                                let elem = RawXmlElement::from_reader(reader, &e)?;
                                extra_children.push(PositionedNode::new(
                                    child_idx,
                                    RawXmlNode::Element(elem),
                                ));
                                child_idx += 1;
                            }
                            #[cfg(not(feature = "extra-children"))]
                            _ => {
                                // Skip unknown element
                                skip_element(reader)?;
                            }
                        }
                    }
                    Event::Empty(e) => {
                        match e.local_name().as_ref() {
                            b"sldId" => {
                                f_sld_id.push(CTSlideIdListEntry::from_xml(reader, &e, true)?);
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            #[cfg(feature = "extra-children")]
                            _ => {
                                // Capture unknown empty element for roundtrip
                                let elem = RawXmlElement::from_empty(&e);
                                extra_children.push(PositionedNode::new(
                                    child_idx,
                                    RawXmlNode::Element(elem),
                                ));
                                child_idx += 1;
                            }
                            #[cfg(not(feature = "extra-children"))]
                            _ => {}
                        }
                    }
                    Event::End(_) => break,
                    Event::Eof => break,
                    _ => {}
                }
                buf.clear();
            }
        }

        Ok(Self {
            sld_id: f_sld_id,
            #[cfg(feature = "extra-children")]
            extra_children,
        })
    }
}

impl FromXml for CTSlideMasterIdListEntry {
    fn from_xml<R: BufRead>(
        reader: &mut Reader<R>,
        start_tag: &BytesStart,
        is_empty: bool,
    ) -> Result<Self, ParseError> {
        let mut f_id = None;
        let mut f_ext_lst = None;
        #[cfg(feature = "extra-attrs")]
        let mut extra_attrs = std::collections::HashMap::new();
        #[cfg(feature = "extra-children")]
        let mut extra_children = Vec::new();
        #[cfg(feature = "extra-children")]
        let mut child_idx: usize = 0;

        // Parse attributes
        for attr in start_tag.attributes().filter_map(|a| a.ok()) {
            let val = String::from_utf8_lossy(&attr.value);
            match attr.key.local_name().as_ref() {
                b"id" => {
                    f_id = val.parse().ok();
                }
                #[cfg(feature = "extra-attrs")]
                unknown => {
                    let key = String::from_utf8_lossy(attr.key.as_ref()).into_owned();
                    extra_attrs.insert(key, val.into_owned());
                }
                #[cfg(not(feature = "extra-attrs"))]
                _ => {}
            }
        }

        // Parse child elements
        if !is_empty {
            let mut buf = Vec::new();
            loop {
                match reader.read_event_into(&mut buf)? {
                    Event::Start(e) => {
                        match e.local_name().as_ref() {
                            b"extLst" => {
                                f_ext_lst =
                                    Some(Box::new(CTExtensionList::from_xml(reader, &e, false)?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            #[cfg(feature = "extra-children")]
                            _ => {
                                // Capture unknown element for roundtrip
                                let elem = RawXmlElement::from_reader(reader, &e)?;
                                extra_children.push(PositionedNode::new(
                                    child_idx,
                                    RawXmlNode::Element(elem),
                                ));
                                child_idx += 1;
                            }
                            #[cfg(not(feature = "extra-children"))]
                            _ => {
                                // Skip unknown element
                                skip_element(reader)?;
                            }
                        }
                    }
                    Event::Empty(e) => {
                        match e.local_name().as_ref() {
                            b"extLst" => {
                                f_ext_lst =
                                    Some(Box::new(CTExtensionList::from_xml(reader, &e, true)?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            #[cfg(feature = "extra-children")]
                            _ => {
                                // Capture unknown empty element for roundtrip
                                let elem = RawXmlElement::from_empty(&e);
                                extra_children.push(PositionedNode::new(
                                    child_idx,
                                    RawXmlNode::Element(elem),
                                ));
                                child_idx += 1;
                            }
                            #[cfg(not(feature = "extra-children"))]
                            _ => {}
                        }
                    }
                    Event::End(_) => break,
                    Event::Eof => break,
                    _ => {}
                }
                buf.clear();
            }
        }

        Ok(Self {
            id: f_id,
            ext_lst: f_ext_lst,
            #[cfg(feature = "extra-attrs")]
            extra_attrs,
            #[cfg(feature = "extra-children")]
            extra_children,
        })
    }
}

impl FromXml for CTSlideMasterIdList {
    fn from_xml<R: BufRead>(
        reader: &mut Reader<R>,
        start_tag: &BytesStart,
        is_empty: bool,
    ) -> Result<Self, ParseError> {
        let mut f_sld_master_id = Vec::new();
        #[cfg(feature = "extra-children")]
        let mut extra_children = Vec::new();
        #[cfg(feature = "extra-children")]
        let mut child_idx: usize = 0;

        // Parse child elements
        if !is_empty {
            let mut buf = Vec::new();
            loop {
                match reader.read_event_into(&mut buf)? {
                    Event::Start(e) => {
                        match e.local_name().as_ref() {
                            b"sldMasterId" => {
                                f_sld_master_id
                                    .push(CTSlideMasterIdListEntry::from_xml(reader, &e, false)?);
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            #[cfg(feature = "extra-children")]
                            _ => {
                                // Capture unknown element for roundtrip
                                let elem = RawXmlElement::from_reader(reader, &e)?;
                                extra_children.push(PositionedNode::new(
                                    child_idx,
                                    RawXmlNode::Element(elem),
                                ));
                                child_idx += 1;
                            }
                            #[cfg(not(feature = "extra-children"))]
                            _ => {
                                // Skip unknown element
                                skip_element(reader)?;
                            }
                        }
                    }
                    Event::Empty(e) => {
                        match e.local_name().as_ref() {
                            b"sldMasterId" => {
                                f_sld_master_id
                                    .push(CTSlideMasterIdListEntry::from_xml(reader, &e, true)?);
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            #[cfg(feature = "extra-children")]
                            _ => {
                                // Capture unknown empty element for roundtrip
                                let elem = RawXmlElement::from_empty(&e);
                                extra_children.push(PositionedNode::new(
                                    child_idx,
                                    RawXmlNode::Element(elem),
                                ));
                                child_idx += 1;
                            }
                            #[cfg(not(feature = "extra-children"))]
                            _ => {}
                        }
                    }
                    Event::End(_) => break,
                    Event::Eof => break,
                    _ => {}
                }
                buf.clear();
            }
        }

        Ok(Self {
            sld_master_id: f_sld_master_id,
            #[cfg(feature = "extra-children")]
            extra_children,
        })
    }
}

impl FromXml for CTNotesMasterIdListEntry {
    fn from_xml<R: BufRead>(
        reader: &mut Reader<R>,
        start_tag: &BytesStart,
        is_empty: bool,
    ) -> Result<Self, ParseError> {
        let mut f_ext_lst = None;
        #[cfg(feature = "extra-children")]
        let mut extra_children = Vec::new();
        #[cfg(feature = "extra-children")]
        let mut child_idx: usize = 0;

        // Parse child elements
        if !is_empty {
            let mut buf = Vec::new();
            loop {
                match reader.read_event_into(&mut buf)? {
                    Event::Start(e) => {
                        match e.local_name().as_ref() {
                            b"extLst" => {
                                f_ext_lst =
                                    Some(Box::new(CTExtensionList::from_xml(reader, &e, false)?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            #[cfg(feature = "extra-children")]
                            _ => {
                                // Capture unknown element for roundtrip
                                let elem = RawXmlElement::from_reader(reader, &e)?;
                                extra_children.push(PositionedNode::new(
                                    child_idx,
                                    RawXmlNode::Element(elem),
                                ));
                                child_idx += 1;
                            }
                            #[cfg(not(feature = "extra-children"))]
                            _ => {
                                // Skip unknown element
                                skip_element(reader)?;
                            }
                        }
                    }
                    Event::Empty(e) => {
                        match e.local_name().as_ref() {
                            b"extLst" => {
                                f_ext_lst =
                                    Some(Box::new(CTExtensionList::from_xml(reader, &e, true)?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            #[cfg(feature = "extra-children")]
                            _ => {
                                // Capture unknown empty element for roundtrip
                                let elem = RawXmlElement::from_empty(&e);
                                extra_children.push(PositionedNode::new(
                                    child_idx,
                                    RawXmlNode::Element(elem),
                                ));
                                child_idx += 1;
                            }
                            #[cfg(not(feature = "extra-children"))]
                            _ => {}
                        }
                    }
                    Event::End(_) => break,
                    Event::Eof => break,
                    _ => {}
                }
                buf.clear();
            }
        }

        Ok(Self {
            ext_lst: f_ext_lst,
            #[cfg(feature = "extra-children")]
            extra_children,
        })
    }
}

impl FromXml for CTNotesMasterIdList {
    fn from_xml<R: BufRead>(
        reader: &mut Reader<R>,
        start_tag: &BytesStart,
        is_empty: bool,
    ) -> Result<Self, ParseError> {
        let mut f_notes_master_id = None;
        #[cfg(feature = "extra-children")]
        let mut extra_children = Vec::new();
        #[cfg(feature = "extra-children")]
        let mut child_idx: usize = 0;

        // Parse child elements
        if !is_empty {
            let mut buf = Vec::new();
            loop {
                match reader.read_event_into(&mut buf)? {
                    Event::Start(e) => {
                        match e.local_name().as_ref() {
                            b"notesMasterId" => {
                                f_notes_master_id = Some(Box::new(
                                    CTNotesMasterIdListEntry::from_xml(reader, &e, false)?,
                                ));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            #[cfg(feature = "extra-children")]
                            _ => {
                                // Capture unknown element for roundtrip
                                let elem = RawXmlElement::from_reader(reader, &e)?;
                                extra_children.push(PositionedNode::new(
                                    child_idx,
                                    RawXmlNode::Element(elem),
                                ));
                                child_idx += 1;
                            }
                            #[cfg(not(feature = "extra-children"))]
                            _ => {
                                // Skip unknown element
                                skip_element(reader)?;
                            }
                        }
                    }
                    Event::Empty(e) => {
                        match e.local_name().as_ref() {
                            b"notesMasterId" => {
                                f_notes_master_id = Some(Box::new(
                                    CTNotesMasterIdListEntry::from_xml(reader, &e, true)?,
                                ));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            #[cfg(feature = "extra-children")]
                            _ => {
                                // Capture unknown empty element for roundtrip
                                let elem = RawXmlElement::from_empty(&e);
                                extra_children.push(PositionedNode::new(
                                    child_idx,
                                    RawXmlNode::Element(elem),
                                ));
                                child_idx += 1;
                            }
                            #[cfg(not(feature = "extra-children"))]
                            _ => {}
                        }
                    }
                    Event::End(_) => break,
                    Event::Eof => break,
                    _ => {}
                }
                buf.clear();
            }
        }

        Ok(Self {
            notes_master_id: f_notes_master_id,
            #[cfg(feature = "extra-children")]
            extra_children,
        })
    }
}

impl FromXml for CTHandoutMasterIdListEntry {
    fn from_xml<R: BufRead>(
        reader: &mut Reader<R>,
        start_tag: &BytesStart,
        is_empty: bool,
    ) -> Result<Self, ParseError> {
        let mut f_ext_lst = None;
        #[cfg(feature = "extra-children")]
        let mut extra_children = Vec::new();
        #[cfg(feature = "extra-children")]
        let mut child_idx: usize = 0;

        // Parse child elements
        if !is_empty {
            let mut buf = Vec::new();
            loop {
                match reader.read_event_into(&mut buf)? {
                    Event::Start(e) => {
                        match e.local_name().as_ref() {
                            b"extLst" => {
                                f_ext_lst =
                                    Some(Box::new(CTExtensionList::from_xml(reader, &e, false)?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            #[cfg(feature = "extra-children")]
                            _ => {
                                // Capture unknown element for roundtrip
                                let elem = RawXmlElement::from_reader(reader, &e)?;
                                extra_children.push(PositionedNode::new(
                                    child_idx,
                                    RawXmlNode::Element(elem),
                                ));
                                child_idx += 1;
                            }
                            #[cfg(not(feature = "extra-children"))]
                            _ => {
                                // Skip unknown element
                                skip_element(reader)?;
                            }
                        }
                    }
                    Event::Empty(e) => {
                        match e.local_name().as_ref() {
                            b"extLst" => {
                                f_ext_lst =
                                    Some(Box::new(CTExtensionList::from_xml(reader, &e, true)?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            #[cfg(feature = "extra-children")]
                            _ => {
                                // Capture unknown empty element for roundtrip
                                let elem = RawXmlElement::from_empty(&e);
                                extra_children.push(PositionedNode::new(
                                    child_idx,
                                    RawXmlNode::Element(elem),
                                ));
                                child_idx += 1;
                            }
                            #[cfg(not(feature = "extra-children"))]
                            _ => {}
                        }
                    }
                    Event::End(_) => break,
                    Event::Eof => break,
                    _ => {}
                }
                buf.clear();
            }
        }

        Ok(Self {
            ext_lst: f_ext_lst,
            #[cfg(feature = "extra-children")]
            extra_children,
        })
    }
}

impl FromXml for CTHandoutMasterIdList {
    fn from_xml<R: BufRead>(
        reader: &mut Reader<R>,
        start_tag: &BytesStart,
        is_empty: bool,
    ) -> Result<Self, ParseError> {
        let mut f_handout_master_id = None;
        #[cfg(feature = "extra-children")]
        let mut extra_children = Vec::new();
        #[cfg(feature = "extra-children")]
        let mut child_idx: usize = 0;

        // Parse child elements
        if !is_empty {
            let mut buf = Vec::new();
            loop {
                match reader.read_event_into(&mut buf)? {
                    Event::Start(e) => {
                        match e.local_name().as_ref() {
                            b"handoutMasterId" => {
                                f_handout_master_id = Some(Box::new(
                                    CTHandoutMasterIdListEntry::from_xml(reader, &e, false)?,
                                ));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            #[cfg(feature = "extra-children")]
                            _ => {
                                // Capture unknown element for roundtrip
                                let elem = RawXmlElement::from_reader(reader, &e)?;
                                extra_children.push(PositionedNode::new(
                                    child_idx,
                                    RawXmlNode::Element(elem),
                                ));
                                child_idx += 1;
                            }
                            #[cfg(not(feature = "extra-children"))]
                            _ => {
                                // Skip unknown element
                                skip_element(reader)?;
                            }
                        }
                    }
                    Event::Empty(e) => {
                        match e.local_name().as_ref() {
                            b"handoutMasterId" => {
                                f_handout_master_id = Some(Box::new(
                                    CTHandoutMasterIdListEntry::from_xml(reader, &e, true)?,
                                ));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            #[cfg(feature = "extra-children")]
                            _ => {
                                // Capture unknown empty element for roundtrip
                                let elem = RawXmlElement::from_empty(&e);
                                extra_children.push(PositionedNode::new(
                                    child_idx,
                                    RawXmlNode::Element(elem),
                                ));
                                child_idx += 1;
                            }
                            #[cfg(not(feature = "extra-children"))]
                            _ => {}
                        }
                    }
                    Event::End(_) => break,
                    Event::Eof => break,
                    _ => {}
                }
                buf.clear();
            }
        }

        Ok(Self {
            handout_master_id: f_handout_master_id,
            #[cfg(feature = "extra-children")]
            extra_children,
        })
    }
}

impl FromXml for CTEmbeddedFontDataId {
    fn from_xml<R: BufRead>(
        reader: &mut Reader<R>,
        start_tag: &BytesStart,
        is_empty: bool,
    ) -> Result<Self, ParseError> {
        #[cfg(feature = "extra-children")]
        let mut extra_children = Vec::new();
        #[cfg(feature = "extra-children")]
        let mut child_idx: usize = 0;
        if !is_empty {
            let mut buf = Vec::new();
            loop {
                match reader.read_event_into(&mut buf)? {
                    #[cfg(feature = "extra-children")]
                    Event::Start(e) => {
                        let elem = RawXmlElement::from_reader(reader, &e)?;
                        extra_children
                            .push(PositionedNode::new(child_idx, RawXmlNode::Element(elem)));
                        child_idx += 1;
                    }
                    #[cfg(not(feature = "extra-children"))]
                    Event::Start(_) => {
                        skip_element(reader)?;
                    }
                    #[cfg(feature = "extra-children")]
                    Event::Empty(e) => {
                        let elem = RawXmlElement::from_empty(&e);
                        extra_children
                            .push(PositionedNode::new(child_idx, RawXmlNode::Element(elem)));
                        child_idx += 1;
                    }
                    #[cfg(not(feature = "extra-children"))]
                    Event::Empty(_) => {}
                    Event::End(_) => break,
                    Event::Eof => break,
                    _ => {}
                }
                buf.clear();
            }
        }
        Ok(Self {
            #[cfg(feature = "extra-children")]
            extra_children,
        })
    }
}

impl FromXml for CTEmbeddedFontListEntry {
    fn from_xml<R: BufRead>(
        reader: &mut Reader<R>,
        start_tag: &BytesStart,
        is_empty: bool,
    ) -> Result<Self, ParseError> {
        let mut f_font: Option<Box<ooxml_dml::types::TextFont>> = None;
        let mut f_regular = None;
        let mut f_bold = None;
        let mut f_italic = None;
        let mut f_bold_italic = None;
        #[cfg(feature = "extra-children")]
        let mut extra_children = Vec::new();
        #[cfg(feature = "extra-children")]
        let mut child_idx: usize = 0;

        // Parse child elements
        if !is_empty {
            let mut buf = Vec::new();
            loop {
                match reader.read_event_into(&mut buf)? {
                    Event::Start(e) => {
                        match e.local_name().as_ref() {
                            b"font" => {
                                f_font = Some(Box::new(ooxml_dml::types::TextFont::from_xml(
                                    reader, &e, false,
                                )?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            b"regular" => {
                                f_regular = Some(Box::new(CTEmbeddedFontDataId::from_xml(
                                    reader, &e, false,
                                )?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            b"bold" => {
                                f_bold = Some(Box::new(CTEmbeddedFontDataId::from_xml(
                                    reader, &e, false,
                                )?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            b"italic" => {
                                f_italic = Some(Box::new(CTEmbeddedFontDataId::from_xml(
                                    reader, &e, false,
                                )?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            b"boldItalic" => {
                                f_bold_italic = Some(Box::new(CTEmbeddedFontDataId::from_xml(
                                    reader, &e, false,
                                )?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            #[cfg(feature = "extra-children")]
                            _ => {
                                // Capture unknown element for roundtrip
                                let elem = RawXmlElement::from_reader(reader, &e)?;
                                extra_children.push(PositionedNode::new(
                                    child_idx,
                                    RawXmlNode::Element(elem),
                                ));
                                child_idx += 1;
                            }
                            #[cfg(not(feature = "extra-children"))]
                            _ => {
                                // Skip unknown element
                                skip_element(reader)?;
                            }
                        }
                    }
                    Event::Empty(e) => {
                        match e.local_name().as_ref() {
                            b"font" => {
                                f_font = Some(Box::new(ooxml_dml::types::TextFont::from_xml(
                                    reader, &e, true,
                                )?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            b"regular" => {
                                f_regular = Some(Box::new(CTEmbeddedFontDataId::from_xml(
                                    reader, &e, true,
                                )?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            b"bold" => {
                                f_bold = Some(Box::new(CTEmbeddedFontDataId::from_xml(
                                    reader, &e, true,
                                )?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            b"italic" => {
                                f_italic = Some(Box::new(CTEmbeddedFontDataId::from_xml(
                                    reader, &e, true,
                                )?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            b"boldItalic" => {
                                f_bold_italic = Some(Box::new(CTEmbeddedFontDataId::from_xml(
                                    reader, &e, true,
                                )?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            #[cfg(feature = "extra-children")]
                            _ => {
                                // Capture unknown empty element for roundtrip
                                let elem = RawXmlElement::from_empty(&e);
                                extra_children.push(PositionedNode::new(
                                    child_idx,
                                    RawXmlNode::Element(elem),
                                ));
                                child_idx += 1;
                            }
                            #[cfg(not(feature = "extra-children"))]
                            _ => {}
                        }
                    }
                    Event::End(_) => break,
                    Event::Eof => break,
                    _ => {}
                }
                buf.clear();
            }
        }

        Ok(Self {
            font: f_font.ok_or_else(|| ParseError::MissingAttribute("font".to_string()))?,
            regular: f_regular,
            bold: f_bold,
            italic: f_italic,
            bold_italic: f_bold_italic,
            #[cfg(feature = "extra-children")]
            extra_children,
        })
    }
}

impl FromXml for CTEmbeddedFontList {
    fn from_xml<R: BufRead>(
        reader: &mut Reader<R>,
        start_tag: &BytesStart,
        is_empty: bool,
    ) -> Result<Self, ParseError> {
        let mut f_embedded_font = Vec::new();
        #[cfg(feature = "extra-children")]
        let mut extra_children = Vec::new();
        #[cfg(feature = "extra-children")]
        let mut child_idx: usize = 0;

        // Parse child elements
        if !is_empty {
            let mut buf = Vec::new();
            loop {
                match reader.read_event_into(&mut buf)? {
                    Event::Start(e) => {
                        match e.local_name().as_ref() {
                            b"embeddedFont" => {
                                f_embedded_font
                                    .push(CTEmbeddedFontListEntry::from_xml(reader, &e, false)?);
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            #[cfg(feature = "extra-children")]
                            _ => {
                                // Capture unknown element for roundtrip
                                let elem = RawXmlElement::from_reader(reader, &e)?;
                                extra_children.push(PositionedNode::new(
                                    child_idx,
                                    RawXmlNode::Element(elem),
                                ));
                                child_idx += 1;
                            }
                            #[cfg(not(feature = "extra-children"))]
                            _ => {
                                // Skip unknown element
                                skip_element(reader)?;
                            }
                        }
                    }
                    Event::Empty(e) => {
                        match e.local_name().as_ref() {
                            b"embeddedFont" => {
                                f_embedded_font
                                    .push(CTEmbeddedFontListEntry::from_xml(reader, &e, true)?);
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            #[cfg(feature = "extra-children")]
                            _ => {
                                // Capture unknown empty element for roundtrip
                                let elem = RawXmlElement::from_empty(&e);
                                extra_children.push(PositionedNode::new(
                                    child_idx,
                                    RawXmlNode::Element(elem),
                                ));
                                child_idx += 1;
                            }
                            #[cfg(not(feature = "extra-children"))]
                            _ => {}
                        }
                    }
                    Event::End(_) => break,
                    Event::Eof => break,
                    _ => {}
                }
                buf.clear();
            }
        }

        Ok(Self {
            embedded_font: f_embedded_font,
            #[cfg(feature = "extra-children")]
            extra_children,
        })
    }
}

impl FromXml for CTSmartTags {
    fn from_xml<R: BufRead>(
        reader: &mut Reader<R>,
        start_tag: &BytesStart,
        is_empty: bool,
    ) -> Result<Self, ParseError> {
        #[cfg(feature = "extra-children")]
        let mut extra_children = Vec::new();
        #[cfg(feature = "extra-children")]
        let mut child_idx: usize = 0;
        if !is_empty {
            let mut buf = Vec::new();
            loop {
                match reader.read_event_into(&mut buf)? {
                    #[cfg(feature = "extra-children")]
                    Event::Start(e) => {
                        let elem = RawXmlElement::from_reader(reader, &e)?;
                        extra_children
                            .push(PositionedNode::new(child_idx, RawXmlNode::Element(elem)));
                        child_idx += 1;
                    }
                    #[cfg(not(feature = "extra-children"))]
                    Event::Start(_) => {
                        skip_element(reader)?;
                    }
                    #[cfg(feature = "extra-children")]
                    Event::Empty(e) => {
                        let elem = RawXmlElement::from_empty(&e);
                        extra_children
                            .push(PositionedNode::new(child_idx, RawXmlNode::Element(elem)));
                        child_idx += 1;
                    }
                    #[cfg(not(feature = "extra-children"))]
                    Event::Empty(_) => {}
                    Event::End(_) => break,
                    Event::Eof => break,
                    _ => {}
                }
                buf.clear();
            }
        }
        Ok(Self {
            #[cfg(feature = "extra-children")]
            extra_children,
        })
    }
}

impl FromXml for CTCustomShow {
    fn from_xml<R: BufRead>(
        reader: &mut Reader<R>,
        start_tag: &BytesStart,
        is_empty: bool,
    ) -> Result<Self, ParseError> {
        let mut f_name: Option<STName> = None;
        let mut f_id: Option<u32> = None;
        let mut f_sld_lst: Option<Box<CTSlideRelationshipList>> = None;
        let mut f_ext_lst = None;
        #[cfg(feature = "extra-attrs")]
        let mut extra_attrs = std::collections::HashMap::new();
        #[cfg(feature = "extra-children")]
        let mut extra_children = Vec::new();
        #[cfg(feature = "extra-children")]
        let mut child_idx: usize = 0;

        // Parse attributes
        for attr in start_tag.attributes().filter_map(|a| a.ok()) {
            let val = String::from_utf8_lossy(&attr.value);
            match attr.key.local_name().as_ref() {
                b"name" => {
                    f_name = Some(val.into_owned());
                }
                b"id" => {
                    f_id = val.parse().ok();
                }
                #[cfg(feature = "extra-attrs")]
                unknown => {
                    let key = String::from_utf8_lossy(attr.key.as_ref()).into_owned();
                    extra_attrs.insert(key, val.into_owned());
                }
                #[cfg(not(feature = "extra-attrs"))]
                _ => {}
            }
        }

        // Parse child elements
        if !is_empty {
            let mut buf = Vec::new();
            loop {
                match reader.read_event_into(&mut buf)? {
                    Event::Start(e) => {
                        match e.local_name().as_ref() {
                            b"sldLst" => {
                                f_sld_lst = Some(Box::new(CTSlideRelationshipList::from_xml(
                                    reader, &e, false,
                                )?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            b"extLst" => {
                                f_ext_lst =
                                    Some(Box::new(CTExtensionList::from_xml(reader, &e, false)?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            #[cfg(feature = "extra-children")]
                            _ => {
                                // Capture unknown element for roundtrip
                                let elem = RawXmlElement::from_reader(reader, &e)?;
                                extra_children.push(PositionedNode::new(
                                    child_idx,
                                    RawXmlNode::Element(elem),
                                ));
                                child_idx += 1;
                            }
                            #[cfg(not(feature = "extra-children"))]
                            _ => {
                                // Skip unknown element
                                skip_element(reader)?;
                            }
                        }
                    }
                    Event::Empty(e) => {
                        match e.local_name().as_ref() {
                            b"sldLst" => {
                                f_sld_lst = Some(Box::new(CTSlideRelationshipList::from_xml(
                                    reader, &e, true,
                                )?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            b"extLst" => {
                                f_ext_lst =
                                    Some(Box::new(CTExtensionList::from_xml(reader, &e, true)?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            #[cfg(feature = "extra-children")]
                            _ => {
                                // Capture unknown empty element for roundtrip
                                let elem = RawXmlElement::from_empty(&e);
                                extra_children.push(PositionedNode::new(
                                    child_idx,
                                    RawXmlNode::Element(elem),
                                ));
                                child_idx += 1;
                            }
                            #[cfg(not(feature = "extra-children"))]
                            _ => {}
                        }
                    }
                    Event::End(_) => break,
                    Event::Eof => break,
                    _ => {}
                }
                buf.clear();
            }
        }

        Ok(Self {
            name: f_name.ok_or_else(|| ParseError::MissingAttribute("name".to_string()))?,
            id: f_id.ok_or_else(|| ParseError::MissingAttribute("id".to_string()))?,
            sld_lst: f_sld_lst.ok_or_else(|| ParseError::MissingAttribute("sldLst".to_string()))?,
            ext_lst: f_ext_lst,
            #[cfg(feature = "extra-attrs")]
            extra_attrs,
            #[cfg(feature = "extra-children")]
            extra_children,
        })
    }
}

impl FromXml for CTCustomShowList {
    fn from_xml<R: BufRead>(
        reader: &mut Reader<R>,
        start_tag: &BytesStart,
        is_empty: bool,
    ) -> Result<Self, ParseError> {
        let mut f_cust_show = Vec::new();
        #[cfg(feature = "extra-children")]
        let mut extra_children = Vec::new();
        #[cfg(feature = "extra-children")]
        let mut child_idx: usize = 0;

        // Parse child elements
        if !is_empty {
            let mut buf = Vec::new();
            loop {
                match reader.read_event_into(&mut buf)? {
                    Event::Start(e) => {
                        match e.local_name().as_ref() {
                            b"custShow" => {
                                f_cust_show.push(CTCustomShow::from_xml(reader, &e, false)?);
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            #[cfg(feature = "extra-children")]
                            _ => {
                                // Capture unknown element for roundtrip
                                let elem = RawXmlElement::from_reader(reader, &e)?;
                                extra_children.push(PositionedNode::new(
                                    child_idx,
                                    RawXmlNode::Element(elem),
                                ));
                                child_idx += 1;
                            }
                            #[cfg(not(feature = "extra-children"))]
                            _ => {
                                // Skip unknown element
                                skip_element(reader)?;
                            }
                        }
                    }
                    Event::Empty(e) => {
                        match e.local_name().as_ref() {
                            b"custShow" => {
                                f_cust_show.push(CTCustomShow::from_xml(reader, &e, true)?);
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            #[cfg(feature = "extra-children")]
                            _ => {
                                // Capture unknown empty element for roundtrip
                                let elem = RawXmlElement::from_empty(&e);
                                extra_children.push(PositionedNode::new(
                                    child_idx,
                                    RawXmlNode::Element(elem),
                                ));
                                child_idx += 1;
                            }
                            #[cfg(not(feature = "extra-children"))]
                            _ => {}
                        }
                    }
                    Event::End(_) => break,
                    Event::Eof => break,
                    _ => {}
                }
                buf.clear();
            }
        }

        Ok(Self {
            cust_show: f_cust_show,
            #[cfg(feature = "extra-children")]
            extra_children,
        })
    }
}

impl FromXml for CTPhotoAlbum {
    fn from_xml<R: BufRead>(
        reader: &mut Reader<R>,
        start_tag: &BytesStart,
        is_empty: bool,
    ) -> Result<Self, ParseError> {
        let mut f_bw = None;
        let mut f_show_captions = None;
        let mut f_layout = None;
        let mut f_frame = None;
        let mut f_ext_lst = None;
        #[cfg(feature = "extra-attrs")]
        let mut extra_attrs = std::collections::HashMap::new();
        #[cfg(feature = "extra-children")]
        let mut extra_children = Vec::new();
        #[cfg(feature = "extra-children")]
        let mut child_idx: usize = 0;

        // Parse attributes
        for attr in start_tag.attributes().filter_map(|a| a.ok()) {
            let val = String::from_utf8_lossy(&attr.value);
            match attr.key.local_name().as_ref() {
                b"bw" => {
                    f_bw = Some(val == "true" || val == "1");
                }
                b"showCaptions" => {
                    f_show_captions = Some(val == "true" || val == "1");
                }
                b"layout" => {
                    f_layout = val.parse().ok();
                }
                b"frame" => {
                    f_frame = val.parse().ok();
                }
                #[cfg(feature = "extra-attrs")]
                unknown => {
                    let key = String::from_utf8_lossy(attr.key.as_ref()).into_owned();
                    extra_attrs.insert(key, val.into_owned());
                }
                #[cfg(not(feature = "extra-attrs"))]
                _ => {}
            }
        }

        // Parse child elements
        if !is_empty {
            let mut buf = Vec::new();
            loop {
                match reader.read_event_into(&mut buf)? {
                    Event::Start(e) => {
                        match e.local_name().as_ref() {
                            b"extLst" => {
                                f_ext_lst =
                                    Some(Box::new(CTExtensionList::from_xml(reader, &e, false)?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            #[cfg(feature = "extra-children")]
                            _ => {
                                // Capture unknown element for roundtrip
                                let elem = RawXmlElement::from_reader(reader, &e)?;
                                extra_children.push(PositionedNode::new(
                                    child_idx,
                                    RawXmlNode::Element(elem),
                                ));
                                child_idx += 1;
                            }
                            #[cfg(not(feature = "extra-children"))]
                            _ => {
                                // Skip unknown element
                                skip_element(reader)?;
                            }
                        }
                    }
                    Event::Empty(e) => {
                        match e.local_name().as_ref() {
                            b"extLst" => {
                                f_ext_lst =
                                    Some(Box::new(CTExtensionList::from_xml(reader, &e, true)?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            #[cfg(feature = "extra-children")]
                            _ => {
                                // Capture unknown empty element for roundtrip
                                let elem = RawXmlElement::from_empty(&e);
                                extra_children.push(PositionedNode::new(
                                    child_idx,
                                    RawXmlNode::Element(elem),
                                ));
                                child_idx += 1;
                            }
                            #[cfg(not(feature = "extra-children"))]
                            _ => {}
                        }
                    }
                    Event::End(_) => break,
                    Event::Eof => break,
                    _ => {}
                }
                buf.clear();
            }
        }

        Ok(Self {
            bw: f_bw,
            show_captions: f_show_captions,
            layout: f_layout,
            frame: f_frame,
            ext_lst: f_ext_lst,
            #[cfg(feature = "extra-attrs")]
            extra_attrs,
            #[cfg(feature = "extra-children")]
            extra_children,
        })
    }
}

impl FromXml for CTSlideSize {
    fn from_xml<R: BufRead>(
        reader: &mut Reader<R>,
        start_tag: &BytesStart,
        is_empty: bool,
    ) -> Result<Self, ParseError> {
        let mut f_cx: Option<STSlideSizeCoordinate> = None;
        let mut f_cy: Option<STSlideSizeCoordinate> = None;
        let mut f_type = None;
        #[cfg(feature = "extra-attrs")]
        let mut extra_attrs = std::collections::HashMap::new();

        // Parse attributes
        for attr in start_tag.attributes().filter_map(|a| a.ok()) {
            let val = String::from_utf8_lossy(&attr.value);
            match attr.key.local_name().as_ref() {
                b"cx" => {
                    f_cx = val.parse().ok();
                }
                b"cy" => {
                    f_cy = val.parse().ok();
                }
                b"type" => {
                    f_type = val.parse().ok();
                }
                #[cfg(feature = "extra-attrs")]
                unknown => {
                    let key = String::from_utf8_lossy(attr.key.as_ref()).into_owned();
                    extra_attrs.insert(key, val.into_owned());
                }
                #[cfg(not(feature = "extra-attrs"))]
                _ => {}
            }
        }

        if !is_empty {
            let mut buf = Vec::new();
            loop {
                match reader.read_event_into(&mut buf)? {
                    Event::End(_) => break,
                    Event::Eof => break,
                    _ => {}
                }
                buf.clear();
            }
        }

        Ok(Self {
            cx: f_cx.ok_or_else(|| ParseError::MissingAttribute("cx".to_string()))?,
            cy: f_cy.ok_or_else(|| ParseError::MissingAttribute("cy".to_string()))?,
            r#type: f_type,
            #[cfg(feature = "extra-attrs")]
            extra_attrs,
        })
    }
}

impl FromXml for CTKinsoku {
    fn from_xml<R: BufRead>(
        reader: &mut Reader<R>,
        start_tag: &BytesStart,
        is_empty: bool,
    ) -> Result<Self, ParseError> {
        let mut f_lang = None;
        let mut f_inval_st_chars: Option<String> = None;
        let mut f_inval_end_chars: Option<String> = None;
        #[cfg(feature = "extra-attrs")]
        let mut extra_attrs = std::collections::HashMap::new();

        // Parse attributes
        for attr in start_tag.attributes().filter_map(|a| a.ok()) {
            let val = String::from_utf8_lossy(&attr.value);
            match attr.key.local_name().as_ref() {
                b"lang" => {
                    f_lang = Some(val.into_owned());
                }
                b"invalStChars" => {
                    f_inval_st_chars = Some(val.into_owned());
                }
                b"invalEndChars" => {
                    f_inval_end_chars = Some(val.into_owned());
                }
                #[cfg(feature = "extra-attrs")]
                unknown => {
                    let key = String::from_utf8_lossy(attr.key.as_ref()).into_owned();
                    extra_attrs.insert(key, val.into_owned());
                }
                #[cfg(not(feature = "extra-attrs"))]
                _ => {}
            }
        }

        if !is_empty {
            let mut buf = Vec::new();
            loop {
                match reader.read_event_into(&mut buf)? {
                    Event::End(_) => break,
                    Event::Eof => break,
                    _ => {}
                }
                buf.clear();
            }
        }

        Ok(Self {
            lang: f_lang,
            inval_st_chars: f_inval_st_chars
                .ok_or_else(|| ParseError::MissingAttribute("invalStChars".to_string()))?,
            inval_end_chars: f_inval_end_chars
                .ok_or_else(|| ParseError::MissingAttribute("invalEndChars".to_string()))?,
            #[cfg(feature = "extra-attrs")]
            extra_attrs,
        })
    }
}

impl FromXml for CTModifyVerifier {
    fn from_xml<R: BufRead>(
        reader: &mut Reader<R>,
        start_tag: &BytesStart,
        is_empty: bool,
    ) -> Result<Self, ParseError> {
        let mut f_algorithm_name = None;
        let mut f_hash_value = None;
        let mut f_salt_value = None;
        let mut f_spin_value = None;
        let mut f_crypt_provider_type = None;
        let mut f_crypt_algorithm_class = None;
        let mut f_crypt_algorithm_type = None;
        let mut f_crypt_algorithm_sid = None;
        let mut f_spin_count = None;
        let mut f_salt_data = None;
        let mut f_hash_data = None;
        let mut f_crypt_provider = None;
        let mut f_alg_id_ext = None;
        let mut f_alg_id_ext_source = None;
        let mut f_crypt_provider_type_ext = None;
        let mut f_crypt_provider_type_ext_source = None;
        #[cfg(feature = "extra-attrs")]
        let mut extra_attrs = std::collections::HashMap::new();

        // Parse attributes
        for attr in start_tag.attributes().filter_map(|a| a.ok()) {
            let val = String::from_utf8_lossy(&attr.value);
            match attr.key.local_name().as_ref() {
                b"algorithmName" => {
                    f_algorithm_name = Some(val.into_owned());
                }
                b"hashValue" => {
                    f_hash_value = decode_hex(&val);
                }
                b"saltValue" => {
                    f_salt_value = decode_hex(&val);
                }
                b"spinValue" => {
                    f_spin_value = val.parse().ok();
                }
                b"cryptProviderType" => {
                    f_crypt_provider_type = val.parse().ok();
                }
                b"cryptAlgorithmClass" => {
                    f_crypt_algorithm_class = val.parse().ok();
                }
                b"cryptAlgorithmType" => {
                    f_crypt_algorithm_type = val.parse().ok();
                }
                b"cryptAlgorithmSid" => {
                    f_crypt_algorithm_sid = val.parse().ok();
                }
                b"spinCount" => {
                    f_spin_count = val.parse().ok();
                }
                b"saltData" => {
                    f_salt_data = decode_hex(&val);
                }
                b"hashData" => {
                    f_hash_data = decode_hex(&val);
                }
                b"cryptProvider" => {
                    f_crypt_provider = Some(val.into_owned());
                }
                b"algIdExt" => {
                    f_alg_id_ext = val.parse().ok();
                }
                b"algIdExtSource" => {
                    f_alg_id_ext_source = Some(val.into_owned());
                }
                b"cryptProviderTypeExt" => {
                    f_crypt_provider_type_ext = val.parse().ok();
                }
                b"cryptProviderTypeExtSource" => {
                    f_crypt_provider_type_ext_source = Some(val.into_owned());
                }
                #[cfg(feature = "extra-attrs")]
                unknown => {
                    let key = String::from_utf8_lossy(attr.key.as_ref()).into_owned();
                    extra_attrs.insert(key, val.into_owned());
                }
                #[cfg(not(feature = "extra-attrs"))]
                _ => {}
            }
        }

        if !is_empty {
            let mut buf = Vec::new();
            loop {
                match reader.read_event_into(&mut buf)? {
                    Event::End(_) => break,
                    Event::Eof => break,
                    _ => {}
                }
                buf.clear();
            }
        }

        Ok(Self {
            algorithm_name: f_algorithm_name,
            hash_value: f_hash_value,
            salt_value: f_salt_value,
            spin_value: f_spin_value,
            crypt_provider_type: f_crypt_provider_type,
            crypt_algorithm_class: f_crypt_algorithm_class,
            crypt_algorithm_type: f_crypt_algorithm_type,
            crypt_algorithm_sid: f_crypt_algorithm_sid,
            spin_count: f_spin_count,
            salt_data: f_salt_data,
            hash_data: f_hash_data,
            crypt_provider: f_crypt_provider,
            alg_id_ext: f_alg_id_ext,
            alg_id_ext_source: f_alg_id_ext_source,
            crypt_provider_type_ext: f_crypt_provider_type_ext,
            crypt_provider_type_ext_source: f_crypt_provider_type_ext_source,
            #[cfg(feature = "extra-attrs")]
            extra_attrs,
        })
    }
}

impl FromXml for Presentation {
    fn from_xml<R: BufRead>(
        reader: &mut Reader<R>,
        start_tag: &BytesStart,
        is_empty: bool,
    ) -> Result<Self, ParseError> {
        let mut f_server_zoom = None;
        let mut f_first_slide_num = None;
        let mut f_show_special_pls_on_title_sld = None;
        let mut f_rtl = None;
        let mut f_remove_personal_info_on_save = None;
        let mut f_compat_mode = None;
        let mut f_strict_first_and_last_chars = None;
        let mut f_embed_true_type_fonts = None;
        let mut f_save_subset_fonts = None;
        let mut f_auto_compress_pictures = None;
        let mut f_bookmark_id_seed = None;
        let mut f_conformance = None;
        let mut f_sld_master_id_lst = None;
        let mut f_notes_master_id_lst = None;
        let mut f_handout_master_id_lst = None;
        let mut f_sld_id_lst = None;
        let mut f_sld_sz = None;
        let mut f_notes_sz: Option<Box<ooxml_dml::types::PositiveSize2D>> = None;
        let mut f_smart_tags = None;
        let mut f_embedded_font_lst = None;
        let mut f_cust_show_lst = None;
        let mut f_photo_album = None;
        let mut f_cust_data_lst = None;
        let mut f_kinsoku = None;
        let mut f_default_text_style = None;
        let mut f_modify_verifier = None;
        let mut f_ext_lst = None;
        #[cfg(feature = "extra-attrs")]
        let mut extra_attrs = std::collections::HashMap::new();
        #[cfg(feature = "extra-children")]
        let mut extra_children = Vec::new();
        #[cfg(feature = "extra-children")]
        let mut child_idx: usize = 0;

        // Parse attributes
        for attr in start_tag.attributes().filter_map(|a| a.ok()) {
            let val = String::from_utf8_lossy(&attr.value);
            match attr.key.local_name().as_ref() {
                b"serverZoom" => {
                    f_server_zoom = val.parse().ok();
                }
                b"firstSlideNum" => {
                    f_first_slide_num = val.parse().ok();
                }
                b"showSpecialPlsOnTitleSld" => {
                    f_show_special_pls_on_title_sld = Some(val == "true" || val == "1");
                }
                b"rtl" => {
                    f_rtl = Some(val == "true" || val == "1");
                }
                b"removePersonalInfoOnSave" => {
                    f_remove_personal_info_on_save = Some(val == "true" || val == "1");
                }
                b"compatMode" => {
                    f_compat_mode = Some(val == "true" || val == "1");
                }
                b"strictFirstAndLastChars" => {
                    f_strict_first_and_last_chars = Some(val == "true" || val == "1");
                }
                b"embedTrueTypeFonts" => {
                    f_embed_true_type_fonts = Some(val == "true" || val == "1");
                }
                b"saveSubsetFonts" => {
                    f_save_subset_fonts = Some(val == "true" || val == "1");
                }
                b"autoCompressPictures" => {
                    f_auto_compress_pictures = Some(val == "true" || val == "1");
                }
                b"bookmarkIdSeed" => {
                    f_bookmark_id_seed = val.parse().ok();
                }
                b"conformance" => {
                    f_conformance = val.parse().ok();
                }
                #[cfg(feature = "extra-attrs")]
                unknown => {
                    let key = String::from_utf8_lossy(attr.key.as_ref()).into_owned();
                    extra_attrs.insert(key, val.into_owned());
                }
                #[cfg(not(feature = "extra-attrs"))]
                _ => {}
            }
        }

        // Parse child elements
        if !is_empty {
            let mut buf = Vec::new();
            loop {
                match reader.read_event_into(&mut buf)? {
                    Event::Start(e) => {
                        match e.local_name().as_ref() {
                            b"sldMasterIdLst" => {
                                f_sld_master_id_lst = Some(Box::new(
                                    CTSlideMasterIdList::from_xml(reader, &e, false)?,
                                ));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            b"notesMasterIdLst" => {
                                f_notes_master_id_lst = Some(Box::new(
                                    CTNotesMasterIdList::from_xml(reader, &e, false)?,
                                ));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            b"handoutMasterIdLst" => {
                                f_handout_master_id_lst = Some(Box::new(
                                    CTHandoutMasterIdList::from_xml(reader, &e, false)?,
                                ));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            b"sldIdLst" => {
                                f_sld_id_lst =
                                    Some(Box::new(SlideIdList::from_xml(reader, &e, false)?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            b"sldSz" => {
                                f_sld_sz =
                                    Some(Box::new(CTSlideSize::from_xml(reader, &e, false)?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            b"notesSz" => {
                                f_notes_sz = Some(Box::new(
                                    ooxml_dml::types::PositiveSize2D::from_xml(reader, &e, false)?,
                                ));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            b"smartTags" => {
                                f_smart_tags =
                                    Some(Box::new(CTSmartTags::from_xml(reader, &e, false)?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            b"embeddedFontLst" => {
                                f_embedded_font_lst = Some(Box::new(CTEmbeddedFontList::from_xml(
                                    reader, &e, false,
                                )?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            b"custShowLst" => {
                                f_cust_show_lst =
                                    Some(Box::new(CTCustomShowList::from_xml(reader, &e, false)?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            b"photoAlbum" => {
                                f_photo_album =
                                    Some(Box::new(CTPhotoAlbum::from_xml(reader, &e, false)?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            b"custDataLst" => {
                                f_cust_data_lst = Some(Box::new(CTCustomerDataList::from_xml(
                                    reader, &e, false,
                                )?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            b"kinsoku" => {
                                f_kinsoku = Some(Box::new(CTKinsoku::from_xml(reader, &e, false)?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            b"defaultTextStyle" => {
                                f_default_text_style = Some(Box::new(
                                    ooxml_dml::types::CTTextListStyle::from_xml(reader, &e, false)?,
                                ));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            b"modifyVerifier" => {
                                f_modify_verifier =
                                    Some(Box::new(CTModifyVerifier::from_xml(reader, &e, false)?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            b"extLst" => {
                                f_ext_lst =
                                    Some(Box::new(CTExtensionList::from_xml(reader, &e, false)?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            #[cfg(feature = "extra-children")]
                            _ => {
                                // Capture unknown element for roundtrip
                                let elem = RawXmlElement::from_reader(reader, &e)?;
                                extra_children.push(PositionedNode::new(
                                    child_idx,
                                    RawXmlNode::Element(elem),
                                ));
                                child_idx += 1;
                            }
                            #[cfg(not(feature = "extra-children"))]
                            _ => {
                                // Skip unknown element
                                skip_element(reader)?;
                            }
                        }
                    }
                    Event::Empty(e) => {
                        match e.local_name().as_ref() {
                            b"sldMasterIdLst" => {
                                f_sld_master_id_lst = Some(Box::new(
                                    CTSlideMasterIdList::from_xml(reader, &e, true)?,
                                ));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            b"notesMasterIdLst" => {
                                f_notes_master_id_lst = Some(Box::new(
                                    CTNotesMasterIdList::from_xml(reader, &e, true)?,
                                ));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            b"handoutMasterIdLst" => {
                                f_handout_master_id_lst = Some(Box::new(
                                    CTHandoutMasterIdList::from_xml(reader, &e, true)?,
                                ));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            b"sldIdLst" => {
                                f_sld_id_lst =
                                    Some(Box::new(SlideIdList::from_xml(reader, &e, true)?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            b"sldSz" => {
                                f_sld_sz = Some(Box::new(CTSlideSize::from_xml(reader, &e, true)?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            b"notesSz" => {
                                f_notes_sz = Some(Box::new(
                                    ooxml_dml::types::PositiveSize2D::from_xml(reader, &e, true)?,
                                ));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            b"smartTags" => {
                                f_smart_tags =
                                    Some(Box::new(CTSmartTags::from_xml(reader, &e, true)?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            b"embeddedFontLst" => {
                                f_embedded_font_lst =
                                    Some(Box::new(CTEmbeddedFontList::from_xml(reader, &e, true)?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            b"custShowLst" => {
                                f_cust_show_lst =
                                    Some(Box::new(CTCustomShowList::from_xml(reader, &e, true)?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            b"photoAlbum" => {
                                f_photo_album =
                                    Some(Box::new(CTPhotoAlbum::from_xml(reader, &e, true)?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            b"custDataLst" => {
                                f_cust_data_lst =
                                    Some(Box::new(CTCustomerDataList::from_xml(reader, &e, true)?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            b"kinsoku" => {
                                f_kinsoku = Some(Box::new(CTKinsoku::from_xml(reader, &e, true)?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            b"defaultTextStyle" => {
                                f_default_text_style = Some(Box::new(
                                    ooxml_dml::types::CTTextListStyle::from_xml(reader, &e, true)?,
                                ));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            b"modifyVerifier" => {
                                f_modify_verifier =
                                    Some(Box::new(CTModifyVerifier::from_xml(reader, &e, true)?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            b"extLst" => {
                                f_ext_lst =
                                    Some(Box::new(CTExtensionList::from_xml(reader, &e, true)?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            #[cfg(feature = "extra-children")]
                            _ => {
                                // Capture unknown empty element for roundtrip
                                let elem = RawXmlElement::from_empty(&e);
                                extra_children.push(PositionedNode::new(
                                    child_idx,
                                    RawXmlNode::Element(elem),
                                ));
                                child_idx += 1;
                            }
                            #[cfg(not(feature = "extra-children"))]
                            _ => {}
                        }
                    }
                    Event::End(_) => break,
                    Event::Eof => break,
                    _ => {}
                }
                buf.clear();
            }
        }

        Ok(Self {
            server_zoom: f_server_zoom,
            first_slide_num: f_first_slide_num,
            show_special_pls_on_title_sld: f_show_special_pls_on_title_sld,
            rtl: f_rtl,
            remove_personal_info_on_save: f_remove_personal_info_on_save,
            compat_mode: f_compat_mode,
            strict_first_and_last_chars: f_strict_first_and_last_chars,
            embed_true_type_fonts: f_embed_true_type_fonts,
            save_subset_fonts: f_save_subset_fonts,
            auto_compress_pictures: f_auto_compress_pictures,
            bookmark_id_seed: f_bookmark_id_seed,
            conformance: f_conformance,
            sld_master_id_lst: f_sld_master_id_lst,
            notes_master_id_lst: f_notes_master_id_lst,
            handout_master_id_lst: f_handout_master_id_lst,
            sld_id_lst: f_sld_id_lst,
            sld_sz: f_sld_sz,
            notes_sz: f_notes_sz
                .ok_or_else(|| ParseError::MissingAttribute("notesSz".to_string()))?,
            smart_tags: f_smart_tags,
            embedded_font_lst: f_embedded_font_lst,
            cust_show_lst: f_cust_show_lst,
            photo_album: f_photo_album,
            cust_data_lst: f_cust_data_lst,
            kinsoku: f_kinsoku,
            default_text_style: f_default_text_style,
            modify_verifier: f_modify_verifier,
            ext_lst: f_ext_lst,
            #[cfg(feature = "extra-attrs")]
            extra_attrs,
            #[cfg(feature = "extra-children")]
            extra_children,
        })
    }
}

impl FromXml for CTHtmlPublishProperties {
    fn from_xml<R: BufRead>(
        reader: &mut Reader<R>,
        start_tag: &BytesStart,
        is_empty: bool,
    ) -> Result<Self, ParseError> {
        let mut f_show_speaker_notes = None;
        let mut f_target = None;
        let mut f_title = None;
        let mut f_slide_list_choice: Option<Box<EGSlideListChoice>> = None;
        let mut f_ext_lst = None;
        #[cfg(feature = "extra-attrs")]
        let mut extra_attrs = std::collections::HashMap::new();
        #[cfg(feature = "extra-children")]
        let mut extra_children = Vec::new();
        #[cfg(feature = "extra-children")]
        let mut child_idx: usize = 0;

        // Parse attributes
        for attr in start_tag.attributes().filter_map(|a| a.ok()) {
            let val = String::from_utf8_lossy(&attr.value);
            match attr.key.local_name().as_ref() {
                b"showSpeakerNotes" => {
                    f_show_speaker_notes = Some(val == "true" || val == "1");
                }
                b"target" => {
                    f_target = Some(val.into_owned());
                }
                b"title" => {
                    f_title = Some(val.into_owned());
                }
                #[cfg(feature = "extra-attrs")]
                unknown => {
                    let key = String::from_utf8_lossy(attr.key.as_ref()).into_owned();
                    extra_attrs.insert(key, val.into_owned());
                }
                #[cfg(not(feature = "extra-attrs"))]
                _ => {}
            }
        }

        // Parse child elements
        if !is_empty {
            let mut buf = Vec::new();
            loop {
                match reader.read_event_into(&mut buf)? {
                    Event::Start(e) => {
                        match e.local_name().as_ref() {
                            b"sldAll" | b"sldRg" | b"custShow" => {
                                f_slide_list_choice =
                                    Some(Box::new(EGSlideListChoice::from_xml(reader, &e, false)?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            b"extLst" => {
                                f_ext_lst =
                                    Some(Box::new(CTExtensionList::from_xml(reader, &e, false)?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            #[cfg(feature = "extra-children")]
                            _ => {
                                // Capture unknown element for roundtrip
                                let elem = RawXmlElement::from_reader(reader, &e)?;
                                extra_children.push(PositionedNode::new(
                                    child_idx,
                                    RawXmlNode::Element(elem),
                                ));
                                child_idx += 1;
                            }
                            #[cfg(not(feature = "extra-children"))]
                            _ => {
                                // Skip unknown element
                                skip_element(reader)?;
                            }
                        }
                    }
                    Event::Empty(e) => {
                        match e.local_name().as_ref() {
                            b"sldAll" | b"sldRg" | b"custShow" => {
                                f_slide_list_choice =
                                    Some(Box::new(EGSlideListChoice::from_xml(reader, &e, true)?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            b"extLst" => {
                                f_ext_lst =
                                    Some(Box::new(CTExtensionList::from_xml(reader, &e, true)?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            #[cfg(feature = "extra-children")]
                            _ => {
                                // Capture unknown empty element for roundtrip
                                let elem = RawXmlElement::from_empty(&e);
                                extra_children.push(PositionedNode::new(
                                    child_idx,
                                    RawXmlNode::Element(elem),
                                ));
                                child_idx += 1;
                            }
                            #[cfg(not(feature = "extra-children"))]
                            _ => {}
                        }
                    }
                    Event::End(_) => break,
                    Event::Eof => break,
                    _ => {}
                }
                buf.clear();
            }
        }

        Ok(Self {
            show_speaker_notes: f_show_speaker_notes,
            target: f_target,
            title: f_title,
            slide_list_choice: f_slide_list_choice,
            ext_lst: f_ext_lst,
            #[cfg(feature = "extra-attrs")]
            extra_attrs,
            #[cfg(feature = "extra-children")]
            extra_children,
        })
    }
}

impl FromXml for CTWebProperties {
    fn from_xml<R: BufRead>(
        reader: &mut Reader<R>,
        start_tag: &BytesStart,
        is_empty: bool,
    ) -> Result<Self, ParseError> {
        let mut f_show_animation = None;
        let mut f_resize_graphics = None;
        let mut f_allow_png = None;
        let mut f_rely_on_vml = None;
        let mut f_organize_in_folders = None;
        let mut f_use_long_filenames = None;
        let mut f_img_sz = None;
        let mut f_encoding = None;
        let mut f_clr = None;
        let mut f_ext_lst = None;
        #[cfg(feature = "extra-attrs")]
        let mut extra_attrs = std::collections::HashMap::new();
        #[cfg(feature = "extra-children")]
        let mut extra_children = Vec::new();
        #[cfg(feature = "extra-children")]
        let mut child_idx: usize = 0;

        // Parse attributes
        for attr in start_tag.attributes().filter_map(|a| a.ok()) {
            let val = String::from_utf8_lossy(&attr.value);
            match attr.key.local_name().as_ref() {
                b"showAnimation" => {
                    f_show_animation = Some(val == "true" || val == "1");
                }
                b"resizeGraphics" => {
                    f_resize_graphics = Some(val == "true" || val == "1");
                }
                b"allowPng" => {
                    f_allow_png = Some(val == "true" || val == "1");
                }
                b"relyOnVml" => {
                    f_rely_on_vml = Some(val == "true" || val == "1");
                }
                b"organizeInFolders" => {
                    f_organize_in_folders = Some(val == "true" || val == "1");
                }
                b"useLongFilenames" => {
                    f_use_long_filenames = Some(val == "true" || val == "1");
                }
                b"imgSz" => {
                    f_img_sz = val.parse().ok();
                }
                b"encoding" => {
                    f_encoding = Some(val.into_owned());
                }
                b"clr" => {
                    f_clr = val.parse().ok();
                }
                #[cfg(feature = "extra-attrs")]
                unknown => {
                    let key = String::from_utf8_lossy(attr.key.as_ref()).into_owned();
                    extra_attrs.insert(key, val.into_owned());
                }
                #[cfg(not(feature = "extra-attrs"))]
                _ => {}
            }
        }

        // Parse child elements
        if !is_empty {
            let mut buf = Vec::new();
            loop {
                match reader.read_event_into(&mut buf)? {
                    Event::Start(e) => {
                        match e.local_name().as_ref() {
                            b"extLst" => {
                                f_ext_lst =
                                    Some(Box::new(CTExtensionList::from_xml(reader, &e, false)?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            #[cfg(feature = "extra-children")]
                            _ => {
                                // Capture unknown element for roundtrip
                                let elem = RawXmlElement::from_reader(reader, &e)?;
                                extra_children.push(PositionedNode::new(
                                    child_idx,
                                    RawXmlNode::Element(elem),
                                ));
                                child_idx += 1;
                            }
                            #[cfg(not(feature = "extra-children"))]
                            _ => {
                                // Skip unknown element
                                skip_element(reader)?;
                            }
                        }
                    }
                    Event::Empty(e) => {
                        match e.local_name().as_ref() {
                            b"extLst" => {
                                f_ext_lst =
                                    Some(Box::new(CTExtensionList::from_xml(reader, &e, true)?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            #[cfg(feature = "extra-children")]
                            _ => {
                                // Capture unknown empty element for roundtrip
                                let elem = RawXmlElement::from_empty(&e);
                                extra_children.push(PositionedNode::new(
                                    child_idx,
                                    RawXmlNode::Element(elem),
                                ));
                                child_idx += 1;
                            }
                            #[cfg(not(feature = "extra-children"))]
                            _ => {}
                        }
                    }
                    Event::End(_) => break,
                    Event::Eof => break,
                    _ => {}
                }
                buf.clear();
            }
        }

        Ok(Self {
            show_animation: f_show_animation,
            resize_graphics: f_resize_graphics,
            allow_png: f_allow_png,
            rely_on_vml: f_rely_on_vml,
            organize_in_folders: f_organize_in_folders,
            use_long_filenames: f_use_long_filenames,
            img_sz: f_img_sz,
            encoding: f_encoding,
            clr: f_clr,
            ext_lst: f_ext_lst,
            #[cfg(feature = "extra-attrs")]
            extra_attrs,
            #[cfg(feature = "extra-children")]
            extra_children,
        })
    }
}

impl FromXml for CTPrintProperties {
    fn from_xml<R: BufRead>(
        reader: &mut Reader<R>,
        start_tag: &BytesStart,
        is_empty: bool,
    ) -> Result<Self, ParseError> {
        let mut f_prn_what = None;
        let mut f_clr_mode = None;
        let mut f_hidden_slides = None;
        let mut f_scale_to_fit_paper = None;
        let mut f_frame_slides = None;
        let mut f_ext_lst = None;
        #[cfg(feature = "extra-attrs")]
        let mut extra_attrs = std::collections::HashMap::new();
        #[cfg(feature = "extra-children")]
        let mut extra_children = Vec::new();
        #[cfg(feature = "extra-children")]
        let mut child_idx: usize = 0;

        // Parse attributes
        for attr in start_tag.attributes().filter_map(|a| a.ok()) {
            let val = String::from_utf8_lossy(&attr.value);
            match attr.key.local_name().as_ref() {
                b"prnWhat" => {
                    f_prn_what = val.parse().ok();
                }
                b"clrMode" => {
                    f_clr_mode = val.parse().ok();
                }
                b"hiddenSlides" => {
                    f_hidden_slides = Some(val == "true" || val == "1");
                }
                b"scaleToFitPaper" => {
                    f_scale_to_fit_paper = Some(val == "true" || val == "1");
                }
                b"frameSlides" => {
                    f_frame_slides = Some(val == "true" || val == "1");
                }
                #[cfg(feature = "extra-attrs")]
                unknown => {
                    let key = String::from_utf8_lossy(attr.key.as_ref()).into_owned();
                    extra_attrs.insert(key, val.into_owned());
                }
                #[cfg(not(feature = "extra-attrs"))]
                _ => {}
            }
        }

        // Parse child elements
        if !is_empty {
            let mut buf = Vec::new();
            loop {
                match reader.read_event_into(&mut buf)? {
                    Event::Start(e) => {
                        match e.local_name().as_ref() {
                            b"extLst" => {
                                f_ext_lst =
                                    Some(Box::new(CTExtensionList::from_xml(reader, &e, false)?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            #[cfg(feature = "extra-children")]
                            _ => {
                                // Capture unknown element for roundtrip
                                let elem = RawXmlElement::from_reader(reader, &e)?;
                                extra_children.push(PositionedNode::new(
                                    child_idx,
                                    RawXmlNode::Element(elem),
                                ));
                                child_idx += 1;
                            }
                            #[cfg(not(feature = "extra-children"))]
                            _ => {
                                // Skip unknown element
                                skip_element(reader)?;
                            }
                        }
                    }
                    Event::Empty(e) => {
                        match e.local_name().as_ref() {
                            b"extLst" => {
                                f_ext_lst =
                                    Some(Box::new(CTExtensionList::from_xml(reader, &e, true)?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            #[cfg(feature = "extra-children")]
                            _ => {
                                // Capture unknown empty element for roundtrip
                                let elem = RawXmlElement::from_empty(&e);
                                extra_children.push(PositionedNode::new(
                                    child_idx,
                                    RawXmlNode::Element(elem),
                                ));
                                child_idx += 1;
                            }
                            #[cfg(not(feature = "extra-children"))]
                            _ => {}
                        }
                    }
                    Event::End(_) => break,
                    Event::Eof => break,
                    _ => {}
                }
                buf.clear();
            }
        }

        Ok(Self {
            prn_what: f_prn_what,
            clr_mode: f_clr_mode,
            hidden_slides: f_hidden_slides,
            scale_to_fit_paper: f_scale_to_fit_paper,
            frame_slides: f_frame_slides,
            ext_lst: f_ext_lst,
            #[cfg(feature = "extra-attrs")]
            extra_attrs,
            #[cfg(feature = "extra-children")]
            extra_children,
        })
    }
}

impl FromXml for CTShowInfoBrowse {
    fn from_xml<R: BufRead>(
        reader: &mut Reader<R>,
        start_tag: &BytesStart,
        is_empty: bool,
    ) -> Result<Self, ParseError> {
        let mut f_show_scrollbar = None;
        #[cfg(feature = "extra-attrs")]
        let mut extra_attrs = std::collections::HashMap::new();

        // Parse attributes
        for attr in start_tag.attributes().filter_map(|a| a.ok()) {
            let val = String::from_utf8_lossy(&attr.value);
            match attr.key.local_name().as_ref() {
                b"showScrollbar" => {
                    f_show_scrollbar = Some(val == "true" || val == "1");
                }
                #[cfg(feature = "extra-attrs")]
                unknown => {
                    let key = String::from_utf8_lossy(attr.key.as_ref()).into_owned();
                    extra_attrs.insert(key, val.into_owned());
                }
                #[cfg(not(feature = "extra-attrs"))]
                _ => {}
            }
        }

        if !is_empty {
            let mut buf = Vec::new();
            loop {
                match reader.read_event_into(&mut buf)? {
                    Event::End(_) => break,
                    Event::Eof => break,
                    _ => {}
                }
                buf.clear();
            }
        }

        Ok(Self {
            show_scrollbar: f_show_scrollbar,
            #[cfg(feature = "extra-attrs")]
            extra_attrs,
        })
    }
}

impl FromXml for CTShowInfoKiosk {
    fn from_xml<R: BufRead>(
        reader: &mut Reader<R>,
        start_tag: &BytesStart,
        is_empty: bool,
    ) -> Result<Self, ParseError> {
        let mut f_restart = None;
        #[cfg(feature = "extra-attrs")]
        let mut extra_attrs = std::collections::HashMap::new();

        // Parse attributes
        for attr in start_tag.attributes().filter_map(|a| a.ok()) {
            let val = String::from_utf8_lossy(&attr.value);
            match attr.key.local_name().as_ref() {
                b"restart" => {
                    f_restart = val.parse().ok();
                }
                #[cfg(feature = "extra-attrs")]
                unknown => {
                    let key = String::from_utf8_lossy(attr.key.as_ref()).into_owned();
                    extra_attrs.insert(key, val.into_owned());
                }
                #[cfg(not(feature = "extra-attrs"))]
                _ => {}
            }
        }

        if !is_empty {
            let mut buf = Vec::new();
            loop {
                match reader.read_event_into(&mut buf)? {
                    Event::End(_) => break,
                    Event::Eof => break,
                    _ => {}
                }
                buf.clear();
            }
        }

        Ok(Self {
            restart: f_restart,
            #[cfg(feature = "extra-attrs")]
            extra_attrs,
        })
    }
}

impl FromXml for EGShowType {
    fn from_xml<R: BufRead>(
        reader: &mut Reader<R>,
        start_tag: &BytesStart,
        is_empty: bool,
    ) -> Result<Self, ParseError> {
        let tag = start_tag.local_name();
        match tag.as_ref() {
            b"present" => {
                let inner = CTEmpty::from_xml(reader, start_tag, is_empty)?;
                Ok(Self::Present(Box::new(inner)))
            }
            b"browse" => {
                let inner = CTShowInfoBrowse::from_xml(reader, start_tag, is_empty)?;
                Ok(Self::Browse(Box::new(inner)))
            }
            b"kiosk" => {
                let inner = CTShowInfoKiosk::from_xml(reader, start_tag, is_empty)?;
                Ok(Self::Kiosk(Box::new(inner)))
            }
            _ => Err(ParseError::UnexpectedElement(
                String::from_utf8_lossy(start_tag.name().as_ref()).into_owned(),
            )),
        }
    }
}

impl FromXml for CTShowProperties {
    fn from_xml<R: BufRead>(
        reader: &mut Reader<R>,
        start_tag: &BytesStart,
        is_empty: bool,
    ) -> Result<Self, ParseError> {
        let mut f_loop = None;
        let mut f_show_narration = None;
        let mut f_show_animation = None;
        let mut f_use_timings = None;
        let mut f_show_type = None;
        let mut f_slide_list_choice = None;
        let mut f_pen_clr = None;
        let mut f_ext_lst = None;
        #[cfg(feature = "extra-attrs")]
        let mut extra_attrs = std::collections::HashMap::new();
        #[cfg(feature = "extra-children")]
        let mut extra_children = Vec::new();
        #[cfg(feature = "extra-children")]
        let mut child_idx: usize = 0;

        // Parse attributes
        for attr in start_tag.attributes().filter_map(|a| a.ok()) {
            let val = String::from_utf8_lossy(&attr.value);
            match attr.key.local_name().as_ref() {
                b"loop" => {
                    f_loop = Some(val == "true" || val == "1");
                }
                b"showNarration" => {
                    f_show_narration = Some(val == "true" || val == "1");
                }
                b"showAnimation" => {
                    f_show_animation = Some(val == "true" || val == "1");
                }
                b"useTimings" => {
                    f_use_timings = Some(val == "true" || val == "1");
                }
                #[cfg(feature = "extra-attrs")]
                unknown => {
                    let key = String::from_utf8_lossy(attr.key.as_ref()).into_owned();
                    extra_attrs.insert(key, val.into_owned());
                }
                #[cfg(not(feature = "extra-attrs"))]
                _ => {}
            }
        }

        // Parse child elements
        if !is_empty {
            let mut buf = Vec::new();
            loop {
                match reader.read_event_into(&mut buf)? {
                    Event::Start(e) => {
                        match e.local_name().as_ref() {
                            b"present" | b"browse" | b"kiosk" => {
                                f_show_type =
                                    Some(Box::new(EGShowType::from_xml(reader, &e, false)?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            b"sldAll" | b"sldRg" | b"custShow" => {
                                f_slide_list_choice =
                                    Some(Box::new(EGSlideListChoice::from_xml(reader, &e, false)?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            b"penClr" => {
                                f_pen_clr = Some(Box::new(ooxml_dml::types::CTColor::from_xml(
                                    reader, &e, false,
                                )?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            b"extLst" => {
                                f_ext_lst =
                                    Some(Box::new(CTExtensionList::from_xml(reader, &e, false)?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            #[cfg(feature = "extra-children")]
                            _ => {
                                // Capture unknown element for roundtrip
                                let elem = RawXmlElement::from_reader(reader, &e)?;
                                extra_children.push(PositionedNode::new(
                                    child_idx,
                                    RawXmlNode::Element(elem),
                                ));
                                child_idx += 1;
                            }
                            #[cfg(not(feature = "extra-children"))]
                            _ => {
                                // Skip unknown element
                                skip_element(reader)?;
                            }
                        }
                    }
                    Event::Empty(e) => {
                        match e.local_name().as_ref() {
                            b"present" | b"browse" | b"kiosk" => {
                                f_show_type =
                                    Some(Box::new(EGShowType::from_xml(reader, &e, true)?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            b"sldAll" | b"sldRg" | b"custShow" => {
                                f_slide_list_choice =
                                    Some(Box::new(EGSlideListChoice::from_xml(reader, &e, true)?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            b"penClr" => {
                                f_pen_clr = Some(Box::new(ooxml_dml::types::CTColor::from_xml(
                                    reader, &e, true,
                                )?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            b"extLst" => {
                                f_ext_lst =
                                    Some(Box::new(CTExtensionList::from_xml(reader, &e, true)?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            #[cfg(feature = "extra-children")]
                            _ => {
                                // Capture unknown empty element for roundtrip
                                let elem = RawXmlElement::from_empty(&e);
                                extra_children.push(PositionedNode::new(
                                    child_idx,
                                    RawXmlNode::Element(elem),
                                ));
                                child_idx += 1;
                            }
                            #[cfg(not(feature = "extra-children"))]
                            _ => {}
                        }
                    }
                    Event::End(_) => break,
                    Event::Eof => break,
                    _ => {}
                }
                buf.clear();
            }
        }

        Ok(Self {
            r#loop: f_loop,
            show_narration: f_show_narration,
            show_animation: f_show_animation,
            use_timings: f_use_timings,
            show_type: f_show_type,
            slide_list_choice: f_slide_list_choice,
            pen_clr: f_pen_clr,
            ext_lst: f_ext_lst,
            #[cfg(feature = "extra-attrs")]
            extra_attrs,
            #[cfg(feature = "extra-children")]
            extra_children,
        })
    }
}

impl FromXml for CTPresentationProperties {
    fn from_xml<R: BufRead>(
        reader: &mut Reader<R>,
        start_tag: &BytesStart,
        is_empty: bool,
    ) -> Result<Self, ParseError> {
        let mut f_html_pub_pr = None;
        let mut f_web_pr = None;
        let mut f_prn_pr = None;
        let mut f_show_pr = None;
        let mut f_clr_mru = None;
        let mut f_ext_lst = None;
        #[cfg(feature = "extra-children")]
        let mut extra_children = Vec::new();
        #[cfg(feature = "extra-children")]
        let mut child_idx: usize = 0;

        // Parse child elements
        if !is_empty {
            let mut buf = Vec::new();
            loop {
                match reader.read_event_into(&mut buf)? {
                    Event::Start(e) => {
                        match e.local_name().as_ref() {
                            b"htmlPubPr" => {
                                f_html_pub_pr = Some(Box::new(CTHtmlPublishProperties::from_xml(
                                    reader, &e, false,
                                )?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            b"webPr" => {
                                f_web_pr =
                                    Some(Box::new(CTWebProperties::from_xml(reader, &e, false)?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            b"prnPr" => {
                                f_prn_pr =
                                    Some(Box::new(CTPrintProperties::from_xml(reader, &e, false)?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            b"showPr" => {
                                f_show_pr =
                                    Some(Box::new(CTShowProperties::from_xml(reader, &e, false)?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            b"clrMru" => {
                                f_clr_mru = Some(Box::new(ooxml_dml::types::CTColorMRU::from_xml(
                                    reader, &e, false,
                                )?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            b"extLst" => {
                                f_ext_lst =
                                    Some(Box::new(CTExtensionList::from_xml(reader, &e, false)?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            #[cfg(feature = "extra-children")]
                            _ => {
                                // Capture unknown element for roundtrip
                                let elem = RawXmlElement::from_reader(reader, &e)?;
                                extra_children.push(PositionedNode::new(
                                    child_idx,
                                    RawXmlNode::Element(elem),
                                ));
                                child_idx += 1;
                            }
                            #[cfg(not(feature = "extra-children"))]
                            _ => {
                                // Skip unknown element
                                skip_element(reader)?;
                            }
                        }
                    }
                    Event::Empty(e) => {
                        match e.local_name().as_ref() {
                            b"htmlPubPr" => {
                                f_html_pub_pr = Some(Box::new(CTHtmlPublishProperties::from_xml(
                                    reader, &e, true,
                                )?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            b"webPr" => {
                                f_web_pr =
                                    Some(Box::new(CTWebProperties::from_xml(reader, &e, true)?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            b"prnPr" => {
                                f_prn_pr =
                                    Some(Box::new(CTPrintProperties::from_xml(reader, &e, true)?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            b"showPr" => {
                                f_show_pr =
                                    Some(Box::new(CTShowProperties::from_xml(reader, &e, true)?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            b"clrMru" => {
                                f_clr_mru = Some(Box::new(ooxml_dml::types::CTColorMRU::from_xml(
                                    reader, &e, true,
                                )?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            b"extLst" => {
                                f_ext_lst =
                                    Some(Box::new(CTExtensionList::from_xml(reader, &e, true)?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            #[cfg(feature = "extra-children")]
                            _ => {
                                // Capture unknown empty element for roundtrip
                                let elem = RawXmlElement::from_empty(&e);
                                extra_children.push(PositionedNode::new(
                                    child_idx,
                                    RawXmlNode::Element(elem),
                                ));
                                child_idx += 1;
                            }
                            #[cfg(not(feature = "extra-children"))]
                            _ => {}
                        }
                    }
                    Event::End(_) => break,
                    Event::Eof => break,
                    _ => {}
                }
                buf.clear();
            }
        }

        Ok(Self {
            html_pub_pr: f_html_pub_pr,
            web_pr: f_web_pr,
            prn_pr: f_prn_pr,
            show_pr: f_show_pr,
            clr_mru: f_clr_mru,
            ext_lst: f_ext_lst,
            #[cfg(feature = "extra-children")]
            extra_children,
        })
    }
}

impl FromXml for CTHeaderFooter {
    fn from_xml<R: BufRead>(
        reader: &mut Reader<R>,
        start_tag: &BytesStart,
        is_empty: bool,
    ) -> Result<Self, ParseError> {
        let mut f_sld_num = None;
        let mut f_hdr = None;
        let mut f_ftr = None;
        let mut f_dt = None;
        let mut f_ext_lst = None;
        #[cfg(feature = "extra-attrs")]
        let mut extra_attrs = std::collections::HashMap::new();
        #[cfg(feature = "extra-children")]
        let mut extra_children = Vec::new();
        #[cfg(feature = "extra-children")]
        let mut child_idx: usize = 0;

        // Parse attributes
        for attr in start_tag.attributes().filter_map(|a| a.ok()) {
            let val = String::from_utf8_lossy(&attr.value);
            match attr.key.local_name().as_ref() {
                b"sldNum" => {
                    f_sld_num = Some(val == "true" || val == "1");
                }
                b"hdr" => {
                    f_hdr = Some(val == "true" || val == "1");
                }
                b"ftr" => {
                    f_ftr = Some(val == "true" || val == "1");
                }
                b"dt" => {
                    f_dt = Some(val == "true" || val == "1");
                }
                #[cfg(feature = "extra-attrs")]
                unknown => {
                    let key = String::from_utf8_lossy(attr.key.as_ref()).into_owned();
                    extra_attrs.insert(key, val.into_owned());
                }
                #[cfg(not(feature = "extra-attrs"))]
                _ => {}
            }
        }

        // Parse child elements
        if !is_empty {
            let mut buf = Vec::new();
            loop {
                match reader.read_event_into(&mut buf)? {
                    Event::Start(e) => {
                        match e.local_name().as_ref() {
                            b"extLst" => {
                                f_ext_lst = Some(Box::new(CTExtensionListModify::from_xml(
                                    reader, &e, false,
                                )?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            #[cfg(feature = "extra-children")]
                            _ => {
                                // Capture unknown element for roundtrip
                                let elem = RawXmlElement::from_reader(reader, &e)?;
                                extra_children.push(PositionedNode::new(
                                    child_idx,
                                    RawXmlNode::Element(elem),
                                ));
                                child_idx += 1;
                            }
                            #[cfg(not(feature = "extra-children"))]
                            _ => {
                                // Skip unknown element
                                skip_element(reader)?;
                            }
                        }
                    }
                    Event::Empty(e) => {
                        match e.local_name().as_ref() {
                            b"extLst" => {
                                f_ext_lst = Some(Box::new(CTExtensionListModify::from_xml(
                                    reader, &e, true,
                                )?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            #[cfg(feature = "extra-children")]
                            _ => {
                                // Capture unknown empty element for roundtrip
                                let elem = RawXmlElement::from_empty(&e);
                                extra_children.push(PositionedNode::new(
                                    child_idx,
                                    RawXmlNode::Element(elem),
                                ));
                                child_idx += 1;
                            }
                            #[cfg(not(feature = "extra-children"))]
                            _ => {}
                        }
                    }
                    Event::End(_) => break,
                    Event::Eof => break,
                    _ => {}
                }
                buf.clear();
            }
        }

        Ok(Self {
            sld_num: f_sld_num,
            hdr: f_hdr,
            ftr: f_ftr,
            dt: f_dt,
            ext_lst: f_ext_lst,
            #[cfg(feature = "extra-attrs")]
            extra_attrs,
            #[cfg(feature = "extra-children")]
            extra_children,
        })
    }
}

impl FromXml for CTPlaceholder {
    fn from_xml<R: BufRead>(
        reader: &mut Reader<R>,
        start_tag: &BytesStart,
        is_empty: bool,
    ) -> Result<Self, ParseError> {
        let mut f_type = None;
        let mut f_orient = None;
        let mut f_sz = None;
        let mut f_idx = None;
        let mut f_has_custom_prompt = None;
        let mut f_ext_lst = None;
        #[cfg(feature = "extra-attrs")]
        let mut extra_attrs = std::collections::HashMap::new();
        #[cfg(feature = "extra-children")]
        let mut extra_children = Vec::new();
        #[cfg(feature = "extra-children")]
        let mut child_idx: usize = 0;

        // Parse attributes
        for attr in start_tag.attributes().filter_map(|a| a.ok()) {
            let val = String::from_utf8_lossy(&attr.value);
            match attr.key.local_name().as_ref() {
                b"type" => {
                    f_type = val.parse().ok();
                }
                b"orient" => {
                    f_orient = val.parse().ok();
                }
                b"sz" => {
                    f_sz = val.parse().ok();
                }
                b"idx" => {
                    f_idx = val.parse().ok();
                }
                b"hasCustomPrompt" => {
                    f_has_custom_prompt = Some(val == "true" || val == "1");
                }
                #[cfg(feature = "extra-attrs")]
                unknown => {
                    let key = String::from_utf8_lossy(attr.key.as_ref()).into_owned();
                    extra_attrs.insert(key, val.into_owned());
                }
                #[cfg(not(feature = "extra-attrs"))]
                _ => {}
            }
        }

        // Parse child elements
        if !is_empty {
            let mut buf = Vec::new();
            loop {
                match reader.read_event_into(&mut buf)? {
                    Event::Start(e) => {
                        match e.local_name().as_ref() {
                            b"extLst" => {
                                f_ext_lst = Some(Box::new(CTExtensionListModify::from_xml(
                                    reader, &e, false,
                                )?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            #[cfg(feature = "extra-children")]
                            _ => {
                                // Capture unknown element for roundtrip
                                let elem = RawXmlElement::from_reader(reader, &e)?;
                                extra_children.push(PositionedNode::new(
                                    child_idx,
                                    RawXmlNode::Element(elem),
                                ));
                                child_idx += 1;
                            }
                            #[cfg(not(feature = "extra-children"))]
                            _ => {
                                // Skip unknown element
                                skip_element(reader)?;
                            }
                        }
                    }
                    Event::Empty(e) => {
                        match e.local_name().as_ref() {
                            b"extLst" => {
                                f_ext_lst = Some(Box::new(CTExtensionListModify::from_xml(
                                    reader, &e, true,
                                )?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            #[cfg(feature = "extra-children")]
                            _ => {
                                // Capture unknown empty element for roundtrip
                                let elem = RawXmlElement::from_empty(&e);
                                extra_children.push(PositionedNode::new(
                                    child_idx,
                                    RawXmlNode::Element(elem),
                                ));
                                child_idx += 1;
                            }
                            #[cfg(not(feature = "extra-children"))]
                            _ => {}
                        }
                    }
                    Event::End(_) => break,
                    Event::Eof => break,
                    _ => {}
                }
                buf.clear();
            }
        }

        Ok(Self {
            r#type: f_type,
            orient: f_orient,
            sz: f_sz,
            idx: f_idx,
            has_custom_prompt: f_has_custom_prompt,
            ext_lst: f_ext_lst,
            #[cfg(feature = "extra-attrs")]
            extra_attrs,
            #[cfg(feature = "extra-children")]
            extra_children,
        })
    }
}

impl FromXml for CTApplicationNonVisualDrawingProps {
    fn from_xml<R: BufRead>(
        reader: &mut Reader<R>,
        start_tag: &BytesStart,
        is_empty: bool,
    ) -> Result<Self, ParseError> {
        let mut f_is_photo = None;
        let mut f_user_drawn = None;
        let mut f_ph = None;
        let mut f_cust_data_lst = None;
        let mut f_ext_lst = None;
        #[cfg(feature = "extra-attrs")]
        let mut extra_attrs = std::collections::HashMap::new();
        #[cfg(feature = "extra-children")]
        let mut extra_children = Vec::new();
        #[cfg(feature = "extra-children")]
        let mut child_idx: usize = 0;

        // Parse attributes
        for attr in start_tag.attributes().filter_map(|a| a.ok()) {
            let val = String::from_utf8_lossy(&attr.value);
            match attr.key.local_name().as_ref() {
                b"isPhoto" => {
                    f_is_photo = Some(val == "true" || val == "1");
                }
                b"userDrawn" => {
                    f_user_drawn = Some(val == "true" || val == "1");
                }
                #[cfg(feature = "extra-attrs")]
                unknown => {
                    let key = String::from_utf8_lossy(attr.key.as_ref()).into_owned();
                    extra_attrs.insert(key, val.into_owned());
                }
                #[cfg(not(feature = "extra-attrs"))]
                _ => {}
            }
        }

        // Parse child elements
        if !is_empty {
            let mut buf = Vec::new();
            loop {
                match reader.read_event_into(&mut buf)? {
                    Event::Start(e) => {
                        match e.local_name().as_ref() {
                            b"ph" => {
                                f_ph = Some(Box::new(CTPlaceholder::from_xml(reader, &e, false)?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            b"custDataLst" => {
                                f_cust_data_lst = Some(Box::new(CTCustomerDataList::from_xml(
                                    reader, &e, false,
                                )?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            b"extLst" => {
                                f_ext_lst =
                                    Some(Box::new(CTExtensionList::from_xml(reader, &e, false)?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            #[cfg(feature = "extra-children")]
                            _ => {
                                // Capture unknown element for roundtrip
                                let elem = RawXmlElement::from_reader(reader, &e)?;
                                extra_children.push(PositionedNode::new(
                                    child_idx,
                                    RawXmlNode::Element(elem),
                                ));
                                child_idx += 1;
                            }
                            #[cfg(not(feature = "extra-children"))]
                            _ => {
                                // Skip unknown element
                                skip_element(reader)?;
                            }
                        }
                    }
                    Event::Empty(e) => {
                        match e.local_name().as_ref() {
                            b"ph" => {
                                f_ph = Some(Box::new(CTPlaceholder::from_xml(reader, &e, true)?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            b"custDataLst" => {
                                f_cust_data_lst =
                                    Some(Box::new(CTCustomerDataList::from_xml(reader, &e, true)?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            b"extLst" => {
                                f_ext_lst =
                                    Some(Box::new(CTExtensionList::from_xml(reader, &e, true)?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            #[cfg(feature = "extra-children")]
                            _ => {
                                // Capture unknown empty element for roundtrip
                                let elem = RawXmlElement::from_empty(&e);
                                extra_children.push(PositionedNode::new(
                                    child_idx,
                                    RawXmlNode::Element(elem),
                                ));
                                child_idx += 1;
                            }
                            #[cfg(not(feature = "extra-children"))]
                            _ => {}
                        }
                    }
                    Event::End(_) => break,
                    Event::Eof => break,
                    _ => {}
                }
                buf.clear();
            }
        }

        Ok(Self {
            is_photo: f_is_photo,
            user_drawn: f_user_drawn,
            ph: f_ph,
            cust_data_lst: f_cust_data_lst,
            ext_lst: f_ext_lst,
            #[cfg(feature = "extra-attrs")]
            extra_attrs,
            #[cfg(feature = "extra-children")]
            extra_children,
        })
    }
}

impl FromXml for ShapeNonVisual {
    fn from_xml<R: BufRead>(
        reader: &mut Reader<R>,
        start_tag: &BytesStart,
        is_empty: bool,
    ) -> Result<Self, ParseError> {
        let mut f_c_nv_pr: Option<Box<ooxml_dml::types::CTNonVisualDrawingProps>> = None;
        let mut f_c_nv_sp_pr: Option<Box<ooxml_dml::types::CTNonVisualDrawingShapeProps>> = None;
        let mut f_nv_pr: Option<Box<CTApplicationNonVisualDrawingProps>> = None;
        #[cfg(feature = "extra-children")]
        let mut extra_children = Vec::new();
        #[cfg(feature = "extra-children")]
        let mut child_idx: usize = 0;

        // Parse child elements
        if !is_empty {
            let mut buf = Vec::new();
            loop {
                match reader.read_event_into(&mut buf)? {
                    Event::Start(e) => {
                        match e.local_name().as_ref() {
                            b"cNvPr" => {
                                f_c_nv_pr = Some(Box::new(
                                    ooxml_dml::types::CTNonVisualDrawingProps::from_xml(
                                        reader, &e, false,
                                    )?,
                                ));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            b"cNvSpPr" => {
                                f_c_nv_sp_pr = Some(Box::new(
                                    ooxml_dml::types::CTNonVisualDrawingShapeProps::from_xml(
                                        reader, &e, false,
                                    )?,
                                ));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            b"nvPr" => {
                                f_nv_pr =
                                    Some(Box::new(CTApplicationNonVisualDrawingProps::from_xml(
                                        reader, &e, false,
                                    )?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            #[cfg(feature = "extra-children")]
                            _ => {
                                // Capture unknown element for roundtrip
                                let elem = RawXmlElement::from_reader(reader, &e)?;
                                extra_children.push(PositionedNode::new(
                                    child_idx,
                                    RawXmlNode::Element(elem),
                                ));
                                child_idx += 1;
                            }
                            #[cfg(not(feature = "extra-children"))]
                            _ => {
                                // Skip unknown element
                                skip_element(reader)?;
                            }
                        }
                    }
                    Event::Empty(e) => {
                        match e.local_name().as_ref() {
                            b"cNvPr" => {
                                f_c_nv_pr = Some(Box::new(
                                    ooxml_dml::types::CTNonVisualDrawingProps::from_xml(
                                        reader, &e, true,
                                    )?,
                                ));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            b"cNvSpPr" => {
                                f_c_nv_sp_pr = Some(Box::new(
                                    ooxml_dml::types::CTNonVisualDrawingShapeProps::from_xml(
                                        reader, &e, true,
                                    )?,
                                ));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            b"nvPr" => {
                                f_nv_pr = Some(Box::new(
                                    CTApplicationNonVisualDrawingProps::from_xml(reader, &e, true)?,
                                ));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            #[cfg(feature = "extra-children")]
                            _ => {
                                // Capture unknown empty element for roundtrip
                                let elem = RawXmlElement::from_empty(&e);
                                extra_children.push(PositionedNode::new(
                                    child_idx,
                                    RawXmlNode::Element(elem),
                                ));
                                child_idx += 1;
                            }
                            #[cfg(not(feature = "extra-children"))]
                            _ => {}
                        }
                    }
                    Event::End(_) => break,
                    Event::Eof => break,
                    _ => {}
                }
                buf.clear();
            }
        }

        Ok(Self {
            c_nv_pr: f_c_nv_pr.ok_or_else(|| ParseError::MissingAttribute("cNvPr".to_string()))?,
            c_nv_sp_pr: f_c_nv_sp_pr
                .ok_or_else(|| ParseError::MissingAttribute("cNvSpPr".to_string()))?,
            nv_pr: f_nv_pr.ok_or_else(|| ParseError::MissingAttribute("nvPr".to_string()))?,
            #[cfg(feature = "extra-children")]
            extra_children,
        })
    }
}

impl FromXml for Shape {
    fn from_xml<R: BufRead>(
        reader: &mut Reader<R>,
        start_tag: &BytesStart,
        is_empty: bool,
    ) -> Result<Self, ParseError> {
        let mut f_use_bg_fill = None;
        let mut f_non_visual_properties: Option<Box<ShapeNonVisual>> = None;
        let mut f_shape_properties: Option<Box<ooxml_dml::types::CTShapeProperties>> = None;
        let mut f_style = None;
        let mut f_text_body = None;
        let mut f_ext_lst = None;
        #[cfg(feature = "extra-attrs")]
        let mut extra_attrs = std::collections::HashMap::new();
        #[cfg(feature = "extra-children")]
        let mut extra_children = Vec::new();
        #[cfg(feature = "extra-children")]
        let mut child_idx: usize = 0;

        // Parse attributes
        for attr in start_tag.attributes().filter_map(|a| a.ok()) {
            let val = String::from_utf8_lossy(&attr.value);
            match attr.key.local_name().as_ref() {
                b"useBgFill" => {
                    f_use_bg_fill = Some(val == "true" || val == "1");
                }
                #[cfg(feature = "extra-attrs")]
                unknown => {
                    let key = String::from_utf8_lossy(attr.key.as_ref()).into_owned();
                    extra_attrs.insert(key, val.into_owned());
                }
                #[cfg(not(feature = "extra-attrs"))]
                _ => {}
            }
        }

        // Parse child elements
        if !is_empty {
            let mut buf = Vec::new();
            loop {
                match reader.read_event_into(&mut buf)? {
                    Event::Start(e) => {
                        match e.local_name().as_ref() {
                            b"nvSpPr" => {
                                f_non_visual_properties =
                                    Some(Box::new(ShapeNonVisual::from_xml(reader, &e, false)?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            b"spPr" => {
                                f_shape_properties =
                                    Some(Box::new(ooxml_dml::types::CTShapeProperties::from_xml(
                                        reader, &e, false,
                                    )?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            b"style" => {
                                f_style = Some(Box::new(ooxml_dml::types::ShapeStyle::from_xml(
                                    reader, &e, false,
                                )?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            b"txBody" => {
                                f_text_body = Some(Box::new(ooxml_dml::types::TextBody::from_xml(
                                    reader, &e, false,
                                )?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            b"extLst" => {
                                f_ext_lst = Some(Box::new(CTExtensionListModify::from_xml(
                                    reader, &e, false,
                                )?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            #[cfg(feature = "extra-children")]
                            _ => {
                                // Capture unknown element for roundtrip
                                let elem = RawXmlElement::from_reader(reader, &e)?;
                                extra_children.push(PositionedNode::new(
                                    child_idx,
                                    RawXmlNode::Element(elem),
                                ));
                                child_idx += 1;
                            }
                            #[cfg(not(feature = "extra-children"))]
                            _ => {
                                // Skip unknown element
                                skip_element(reader)?;
                            }
                        }
                    }
                    Event::Empty(e) => {
                        match e.local_name().as_ref() {
                            b"nvSpPr" => {
                                f_non_visual_properties =
                                    Some(Box::new(ShapeNonVisual::from_xml(reader, &e, true)?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            b"spPr" => {
                                f_shape_properties =
                                    Some(Box::new(ooxml_dml::types::CTShapeProperties::from_xml(
                                        reader, &e, true,
                                    )?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            b"style" => {
                                f_style = Some(Box::new(ooxml_dml::types::ShapeStyle::from_xml(
                                    reader, &e, true,
                                )?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            b"txBody" => {
                                f_text_body = Some(Box::new(ooxml_dml::types::TextBody::from_xml(
                                    reader, &e, true,
                                )?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            b"extLst" => {
                                f_ext_lst = Some(Box::new(CTExtensionListModify::from_xml(
                                    reader, &e, true,
                                )?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            #[cfg(feature = "extra-children")]
                            _ => {
                                // Capture unknown empty element for roundtrip
                                let elem = RawXmlElement::from_empty(&e);
                                extra_children.push(PositionedNode::new(
                                    child_idx,
                                    RawXmlNode::Element(elem),
                                ));
                                child_idx += 1;
                            }
                            #[cfg(not(feature = "extra-children"))]
                            _ => {}
                        }
                    }
                    Event::End(_) => break,
                    Event::Eof => break,
                    _ => {}
                }
                buf.clear();
            }
        }

        Ok(Self {
            use_bg_fill: f_use_bg_fill,
            non_visual_properties: f_non_visual_properties
                .ok_or_else(|| ParseError::MissingAttribute("nvSpPr".to_string()))?,
            shape_properties: f_shape_properties
                .ok_or_else(|| ParseError::MissingAttribute("spPr".to_string()))?,
            style: f_style,
            text_body: f_text_body,
            ext_lst: f_ext_lst,
            #[cfg(feature = "extra-attrs")]
            extra_attrs,
            #[cfg(feature = "extra-children")]
            extra_children,
        })
    }
}

impl FromXml for CTConnectorNonVisual {
    fn from_xml<R: BufRead>(
        reader: &mut Reader<R>,
        start_tag: &BytesStart,
        is_empty: bool,
    ) -> Result<Self, ParseError> {
        let mut f_c_nv_pr: Option<Box<ooxml_dml::types::CTNonVisualDrawingProps>> = None;
        let mut f_c_nv_cxn_sp_pr: Option<Box<ooxml_dml::types::CTNonVisualConnectorProperties>> =
            None;
        let mut f_nv_pr: Option<Box<CTApplicationNonVisualDrawingProps>> = None;
        #[cfg(feature = "extra-children")]
        let mut extra_children = Vec::new();
        #[cfg(feature = "extra-children")]
        let mut child_idx: usize = 0;

        // Parse child elements
        if !is_empty {
            let mut buf = Vec::new();
            loop {
                match reader.read_event_into(&mut buf)? {
                    Event::Start(e) => {
                        match e.local_name().as_ref() {
                            b"cNvPr" => {
                                f_c_nv_pr = Some(Box::new(
                                    ooxml_dml::types::CTNonVisualDrawingProps::from_xml(
                                        reader, &e, false,
                                    )?,
                                ));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            b"cNvCxnSpPr" => {
                                f_c_nv_cxn_sp_pr = Some(Box::new(
                                    ooxml_dml::types::CTNonVisualConnectorProperties::from_xml(
                                        reader, &e, false,
                                    )?,
                                ));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            b"nvPr" => {
                                f_nv_pr =
                                    Some(Box::new(CTApplicationNonVisualDrawingProps::from_xml(
                                        reader, &e, false,
                                    )?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            #[cfg(feature = "extra-children")]
                            _ => {
                                // Capture unknown element for roundtrip
                                let elem = RawXmlElement::from_reader(reader, &e)?;
                                extra_children.push(PositionedNode::new(
                                    child_idx,
                                    RawXmlNode::Element(elem),
                                ));
                                child_idx += 1;
                            }
                            #[cfg(not(feature = "extra-children"))]
                            _ => {
                                // Skip unknown element
                                skip_element(reader)?;
                            }
                        }
                    }
                    Event::Empty(e) => {
                        match e.local_name().as_ref() {
                            b"cNvPr" => {
                                f_c_nv_pr = Some(Box::new(
                                    ooxml_dml::types::CTNonVisualDrawingProps::from_xml(
                                        reader, &e, true,
                                    )?,
                                ));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            b"cNvCxnSpPr" => {
                                f_c_nv_cxn_sp_pr = Some(Box::new(
                                    ooxml_dml::types::CTNonVisualConnectorProperties::from_xml(
                                        reader, &e, true,
                                    )?,
                                ));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            b"nvPr" => {
                                f_nv_pr = Some(Box::new(
                                    CTApplicationNonVisualDrawingProps::from_xml(reader, &e, true)?,
                                ));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            #[cfg(feature = "extra-children")]
                            _ => {
                                // Capture unknown empty element for roundtrip
                                let elem = RawXmlElement::from_empty(&e);
                                extra_children.push(PositionedNode::new(
                                    child_idx,
                                    RawXmlNode::Element(elem),
                                ));
                                child_idx += 1;
                            }
                            #[cfg(not(feature = "extra-children"))]
                            _ => {}
                        }
                    }
                    Event::End(_) => break,
                    Event::Eof => break,
                    _ => {}
                }
                buf.clear();
            }
        }

        Ok(Self {
            c_nv_pr: f_c_nv_pr.ok_or_else(|| ParseError::MissingAttribute("cNvPr".to_string()))?,
            c_nv_cxn_sp_pr: f_c_nv_cxn_sp_pr
                .ok_or_else(|| ParseError::MissingAttribute("cNvCxnSpPr".to_string()))?,
            nv_pr: f_nv_pr.ok_or_else(|| ParseError::MissingAttribute("nvPr".to_string()))?,
            #[cfg(feature = "extra-children")]
            extra_children,
        })
    }
}

impl FromXml for Connector {
    fn from_xml<R: BufRead>(
        reader: &mut Reader<R>,
        start_tag: &BytesStart,
        is_empty: bool,
    ) -> Result<Self, ParseError> {
        let mut f_non_visual_connector_properties: Option<Box<CTConnectorNonVisual>> = None;
        let mut f_shape_properties: Option<Box<ooxml_dml::types::CTShapeProperties>> = None;
        let mut f_style = None;
        let mut f_ext_lst = None;
        #[cfg(feature = "extra-children")]
        let mut extra_children = Vec::new();
        #[cfg(feature = "extra-children")]
        let mut child_idx: usize = 0;

        // Parse child elements
        if !is_empty {
            let mut buf = Vec::new();
            loop {
                match reader.read_event_into(&mut buf)? {
                    Event::Start(e) => {
                        match e.local_name().as_ref() {
                            b"nvCxnSpPr" => {
                                f_non_visual_connector_properties = Some(Box::new(
                                    CTConnectorNonVisual::from_xml(reader, &e, false)?,
                                ));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            b"spPr" => {
                                f_shape_properties =
                                    Some(Box::new(ooxml_dml::types::CTShapeProperties::from_xml(
                                        reader, &e, false,
                                    )?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            b"style" => {
                                f_style = Some(Box::new(ooxml_dml::types::ShapeStyle::from_xml(
                                    reader, &e, false,
                                )?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            b"extLst" => {
                                f_ext_lst = Some(Box::new(CTExtensionListModify::from_xml(
                                    reader, &e, false,
                                )?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            #[cfg(feature = "extra-children")]
                            _ => {
                                // Capture unknown element for roundtrip
                                let elem = RawXmlElement::from_reader(reader, &e)?;
                                extra_children.push(PositionedNode::new(
                                    child_idx,
                                    RawXmlNode::Element(elem),
                                ));
                                child_idx += 1;
                            }
                            #[cfg(not(feature = "extra-children"))]
                            _ => {
                                // Skip unknown element
                                skip_element(reader)?;
                            }
                        }
                    }
                    Event::Empty(e) => {
                        match e.local_name().as_ref() {
                            b"nvCxnSpPr" => {
                                f_non_visual_connector_properties = Some(Box::new(
                                    CTConnectorNonVisual::from_xml(reader, &e, true)?,
                                ));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            b"spPr" => {
                                f_shape_properties =
                                    Some(Box::new(ooxml_dml::types::CTShapeProperties::from_xml(
                                        reader, &e, true,
                                    )?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            b"style" => {
                                f_style = Some(Box::new(ooxml_dml::types::ShapeStyle::from_xml(
                                    reader, &e, true,
                                )?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            b"extLst" => {
                                f_ext_lst = Some(Box::new(CTExtensionListModify::from_xml(
                                    reader, &e, true,
                                )?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            #[cfg(feature = "extra-children")]
                            _ => {
                                // Capture unknown empty element for roundtrip
                                let elem = RawXmlElement::from_empty(&e);
                                extra_children.push(PositionedNode::new(
                                    child_idx,
                                    RawXmlNode::Element(elem),
                                ));
                                child_idx += 1;
                            }
                            #[cfg(not(feature = "extra-children"))]
                            _ => {}
                        }
                    }
                    Event::End(_) => break,
                    Event::Eof => break,
                    _ => {}
                }
                buf.clear();
            }
        }

        Ok(Self {
            non_visual_connector_properties: f_non_visual_connector_properties
                .ok_or_else(|| ParseError::MissingAttribute("nvCxnSpPr".to_string()))?,
            shape_properties: f_shape_properties
                .ok_or_else(|| ParseError::MissingAttribute("spPr".to_string()))?,
            style: f_style,
            ext_lst: f_ext_lst,
            #[cfg(feature = "extra-children")]
            extra_children,
        })
    }
}

impl FromXml for CTPictureNonVisual {
    fn from_xml<R: BufRead>(
        reader: &mut Reader<R>,
        start_tag: &BytesStart,
        is_empty: bool,
    ) -> Result<Self, ParseError> {
        let mut f_c_nv_pr: Option<Box<ooxml_dml::types::CTNonVisualDrawingProps>> = None;
        let mut f_c_nv_pic_pr: Option<Box<ooxml_dml::types::CTNonVisualPictureProperties>> = None;
        let mut f_nv_pr: Option<Box<CTApplicationNonVisualDrawingProps>> = None;
        #[cfg(feature = "extra-children")]
        let mut extra_children = Vec::new();
        #[cfg(feature = "extra-children")]
        let mut child_idx: usize = 0;

        // Parse child elements
        if !is_empty {
            let mut buf = Vec::new();
            loop {
                match reader.read_event_into(&mut buf)? {
                    Event::Start(e) => {
                        match e.local_name().as_ref() {
                            b"cNvPr" => {
                                f_c_nv_pr = Some(Box::new(
                                    ooxml_dml::types::CTNonVisualDrawingProps::from_xml(
                                        reader, &e, false,
                                    )?,
                                ));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            b"cNvPicPr" => {
                                f_c_nv_pic_pr = Some(Box::new(
                                    ooxml_dml::types::CTNonVisualPictureProperties::from_xml(
                                        reader, &e, false,
                                    )?,
                                ));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            b"nvPr" => {
                                f_nv_pr =
                                    Some(Box::new(CTApplicationNonVisualDrawingProps::from_xml(
                                        reader, &e, false,
                                    )?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            #[cfg(feature = "extra-children")]
                            _ => {
                                // Capture unknown element for roundtrip
                                let elem = RawXmlElement::from_reader(reader, &e)?;
                                extra_children.push(PositionedNode::new(
                                    child_idx,
                                    RawXmlNode::Element(elem),
                                ));
                                child_idx += 1;
                            }
                            #[cfg(not(feature = "extra-children"))]
                            _ => {
                                // Skip unknown element
                                skip_element(reader)?;
                            }
                        }
                    }
                    Event::Empty(e) => {
                        match e.local_name().as_ref() {
                            b"cNvPr" => {
                                f_c_nv_pr = Some(Box::new(
                                    ooxml_dml::types::CTNonVisualDrawingProps::from_xml(
                                        reader, &e, true,
                                    )?,
                                ));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            b"cNvPicPr" => {
                                f_c_nv_pic_pr = Some(Box::new(
                                    ooxml_dml::types::CTNonVisualPictureProperties::from_xml(
                                        reader, &e, true,
                                    )?,
                                ));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            b"nvPr" => {
                                f_nv_pr = Some(Box::new(
                                    CTApplicationNonVisualDrawingProps::from_xml(reader, &e, true)?,
                                ));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            #[cfg(feature = "extra-children")]
                            _ => {
                                // Capture unknown empty element for roundtrip
                                let elem = RawXmlElement::from_empty(&e);
                                extra_children.push(PositionedNode::new(
                                    child_idx,
                                    RawXmlNode::Element(elem),
                                ));
                                child_idx += 1;
                            }
                            #[cfg(not(feature = "extra-children"))]
                            _ => {}
                        }
                    }
                    Event::End(_) => break,
                    Event::Eof => break,
                    _ => {}
                }
                buf.clear();
            }
        }

        Ok(Self {
            c_nv_pr: f_c_nv_pr.ok_or_else(|| ParseError::MissingAttribute("cNvPr".to_string()))?,
            c_nv_pic_pr: f_c_nv_pic_pr
                .ok_or_else(|| ParseError::MissingAttribute("cNvPicPr".to_string()))?,
            nv_pr: f_nv_pr.ok_or_else(|| ParseError::MissingAttribute("nvPr".to_string()))?,
            #[cfg(feature = "extra-children")]
            extra_children,
        })
    }
}

impl FromXml for Picture {
    fn from_xml<R: BufRead>(
        reader: &mut Reader<R>,
        start_tag: &BytesStart,
        is_empty: bool,
    ) -> Result<Self, ParseError> {
        let mut f_non_visual_picture_properties: Option<Box<CTPictureNonVisual>> = None;
        let mut f_blip_fill: Option<Box<ooxml_dml::types::BlipFillProperties>> = None;
        let mut f_shape_properties: Option<Box<ooxml_dml::types::CTShapeProperties>> = None;
        let mut f_style = None;
        let mut f_ext_lst = None;
        #[cfg(feature = "extra-children")]
        let mut extra_children = Vec::new();
        #[cfg(feature = "extra-children")]
        let mut child_idx: usize = 0;

        // Parse child elements
        if !is_empty {
            let mut buf = Vec::new();
            loop {
                match reader.read_event_into(&mut buf)? {
                    Event::Start(e) => {
                        match e.local_name().as_ref() {
                            b"nvPicPr" => {
                                f_non_visual_picture_properties = Some(Box::new(
                                    CTPictureNonVisual::from_xml(reader, &e, false)?,
                                ));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            b"blipFill" => {
                                f_blip_fill =
                                    Some(Box::new(ooxml_dml::types::BlipFillProperties::from_xml(
                                        reader, &e, false,
                                    )?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            b"spPr" => {
                                f_shape_properties =
                                    Some(Box::new(ooxml_dml::types::CTShapeProperties::from_xml(
                                        reader, &e, false,
                                    )?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            b"style" => {
                                f_style = Some(Box::new(ooxml_dml::types::ShapeStyle::from_xml(
                                    reader, &e, false,
                                )?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            b"extLst" => {
                                f_ext_lst = Some(Box::new(CTExtensionListModify::from_xml(
                                    reader, &e, false,
                                )?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            #[cfg(feature = "extra-children")]
                            _ => {
                                // Capture unknown element for roundtrip
                                let elem = RawXmlElement::from_reader(reader, &e)?;
                                extra_children.push(PositionedNode::new(
                                    child_idx,
                                    RawXmlNode::Element(elem),
                                ));
                                child_idx += 1;
                            }
                            #[cfg(not(feature = "extra-children"))]
                            _ => {
                                // Skip unknown element
                                skip_element(reader)?;
                            }
                        }
                    }
                    Event::Empty(e) => {
                        match e.local_name().as_ref() {
                            b"nvPicPr" => {
                                f_non_visual_picture_properties =
                                    Some(Box::new(CTPictureNonVisual::from_xml(reader, &e, true)?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            b"blipFill" => {
                                f_blip_fill =
                                    Some(Box::new(ooxml_dml::types::BlipFillProperties::from_xml(
                                        reader, &e, true,
                                    )?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            b"spPr" => {
                                f_shape_properties =
                                    Some(Box::new(ooxml_dml::types::CTShapeProperties::from_xml(
                                        reader, &e, true,
                                    )?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            b"style" => {
                                f_style = Some(Box::new(ooxml_dml::types::ShapeStyle::from_xml(
                                    reader, &e, true,
                                )?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            b"extLst" => {
                                f_ext_lst = Some(Box::new(CTExtensionListModify::from_xml(
                                    reader, &e, true,
                                )?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            #[cfg(feature = "extra-children")]
                            _ => {
                                // Capture unknown empty element for roundtrip
                                let elem = RawXmlElement::from_empty(&e);
                                extra_children.push(PositionedNode::new(
                                    child_idx,
                                    RawXmlNode::Element(elem),
                                ));
                                child_idx += 1;
                            }
                            #[cfg(not(feature = "extra-children"))]
                            _ => {}
                        }
                    }
                    Event::End(_) => break,
                    Event::Eof => break,
                    _ => {}
                }
                buf.clear();
            }
        }

        Ok(Self {
            non_visual_picture_properties: f_non_visual_picture_properties
                .ok_or_else(|| ParseError::MissingAttribute("nvPicPr".to_string()))?,
            blip_fill: f_blip_fill
                .ok_or_else(|| ParseError::MissingAttribute("blipFill".to_string()))?,
            shape_properties: f_shape_properties
                .ok_or_else(|| ParseError::MissingAttribute("spPr".to_string()))?,
            style: f_style,
            ext_lst: f_ext_lst,
            #[cfg(feature = "extra-children")]
            extra_children,
        })
    }
}

impl FromXml for CTGraphicalObjectFrameNonVisual {
    fn from_xml<R: BufRead>(
        reader: &mut Reader<R>,
        start_tag: &BytesStart,
        is_empty: bool,
    ) -> Result<Self, ParseError> {
        let mut f_c_nv_pr: Option<Box<ooxml_dml::types::CTNonVisualDrawingProps>> = None;
        let mut f_c_nv_graphic_frame_pr: Option<
            Box<ooxml_dml::types::CTNonVisualGraphicFrameProperties>,
        > = None;
        let mut f_nv_pr: Option<Box<CTApplicationNonVisualDrawingProps>> = None;
        #[cfg(feature = "extra-children")]
        let mut extra_children = Vec::new();
        #[cfg(feature = "extra-children")]
        let mut child_idx: usize = 0;

        // Parse child elements
        if !is_empty {
            let mut buf = Vec::new();
            loop {
                match reader.read_event_into(&mut buf)? {
                    Event::Start(e) => {
                        match e.local_name().as_ref() {
                            b"cNvPr" => {
                                f_c_nv_pr = Some(Box::new(
                                    ooxml_dml::types::CTNonVisualDrawingProps::from_xml(
                                        reader, &e, false,
                                    )?,
                                ));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            b"cNvGraphicFramePr" => {
                                f_c_nv_graphic_frame_pr = Some(Box::new(
                                    ooxml_dml::types::CTNonVisualGraphicFrameProperties::from_xml(
                                        reader, &e, false,
                                    )?,
                                ));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            b"nvPr" => {
                                f_nv_pr =
                                    Some(Box::new(CTApplicationNonVisualDrawingProps::from_xml(
                                        reader, &e, false,
                                    )?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            #[cfg(feature = "extra-children")]
                            _ => {
                                // Capture unknown element for roundtrip
                                let elem = RawXmlElement::from_reader(reader, &e)?;
                                extra_children.push(PositionedNode::new(
                                    child_idx,
                                    RawXmlNode::Element(elem),
                                ));
                                child_idx += 1;
                            }
                            #[cfg(not(feature = "extra-children"))]
                            _ => {
                                // Skip unknown element
                                skip_element(reader)?;
                            }
                        }
                    }
                    Event::Empty(e) => {
                        match e.local_name().as_ref() {
                            b"cNvPr" => {
                                f_c_nv_pr = Some(Box::new(
                                    ooxml_dml::types::CTNonVisualDrawingProps::from_xml(
                                        reader, &e, true,
                                    )?,
                                ));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            b"cNvGraphicFramePr" => {
                                f_c_nv_graphic_frame_pr = Some(Box::new(
                                    ooxml_dml::types::CTNonVisualGraphicFrameProperties::from_xml(
                                        reader, &e, true,
                                    )?,
                                ));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            b"nvPr" => {
                                f_nv_pr = Some(Box::new(
                                    CTApplicationNonVisualDrawingProps::from_xml(reader, &e, true)?,
                                ));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            #[cfg(feature = "extra-children")]
                            _ => {
                                // Capture unknown empty element for roundtrip
                                let elem = RawXmlElement::from_empty(&e);
                                extra_children.push(PositionedNode::new(
                                    child_idx,
                                    RawXmlNode::Element(elem),
                                ));
                                child_idx += 1;
                            }
                            #[cfg(not(feature = "extra-children"))]
                            _ => {}
                        }
                    }
                    Event::End(_) => break,
                    Event::Eof => break,
                    _ => {}
                }
                buf.clear();
            }
        }

        Ok(Self {
            c_nv_pr: f_c_nv_pr.ok_or_else(|| ParseError::MissingAttribute("cNvPr".to_string()))?,
            c_nv_graphic_frame_pr: f_c_nv_graphic_frame_pr
                .ok_or_else(|| ParseError::MissingAttribute("cNvGraphicFramePr".to_string()))?,
            nv_pr: f_nv_pr.ok_or_else(|| ParseError::MissingAttribute("nvPr".to_string()))?,
            #[cfg(feature = "extra-children")]
            extra_children,
        })
    }
}

impl FromXml for GraphicalObjectFrame {
    fn from_xml<R: BufRead>(
        reader: &mut Reader<R>,
        start_tag: &BytesStart,
        is_empty: bool,
    ) -> Result<Self, ParseError> {
        let mut f_bw_mode = None;
        let mut f_nv_graphic_frame_pr: Option<Box<CTGraphicalObjectFrameNonVisual>> = None;
        let mut f_xfrm: Option<Box<ooxml_dml::types::Transform2D>> = None;
        let mut f_ext_lst = None;
        #[cfg(feature = "extra-attrs")]
        let mut extra_attrs = std::collections::HashMap::new();
        #[cfg(feature = "extra-children")]
        let mut extra_children = Vec::new();
        #[cfg(feature = "extra-children")]
        let mut child_idx: usize = 0;

        // Parse attributes
        for attr in start_tag.attributes().filter_map(|a| a.ok()) {
            let val = String::from_utf8_lossy(&attr.value);
            match attr.key.local_name().as_ref() {
                b"bwMode" => {
                    f_bw_mode = val.parse().ok();
                }
                #[cfg(feature = "extra-attrs")]
                unknown => {
                    let key = String::from_utf8_lossy(attr.key.as_ref()).into_owned();
                    extra_attrs.insert(key, val.into_owned());
                }
                #[cfg(not(feature = "extra-attrs"))]
                _ => {}
            }
        }

        // Parse child elements
        if !is_empty {
            let mut buf = Vec::new();
            loop {
                match reader.read_event_into(&mut buf)? {
                    Event::Start(e) => {
                        match e.local_name().as_ref() {
                            b"nvGraphicFramePr" => {
                                f_nv_graphic_frame_pr = Some(Box::new(
                                    CTGraphicalObjectFrameNonVisual::from_xml(reader, &e, false)?,
                                ));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            b"xfrm" => {
                                f_xfrm = Some(Box::new(ooxml_dml::types::Transform2D::from_xml(
                                    reader, &e, false,
                                )?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            b"extLst" => {
                                f_ext_lst = Some(Box::new(CTExtensionListModify::from_xml(
                                    reader, &e, false,
                                )?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            #[cfg(feature = "extra-children")]
                            _ => {
                                // Capture unknown element for roundtrip
                                let elem = RawXmlElement::from_reader(reader, &e)?;
                                extra_children.push(PositionedNode::new(
                                    child_idx,
                                    RawXmlNode::Element(elem),
                                ));
                                child_idx += 1;
                            }
                            #[cfg(not(feature = "extra-children"))]
                            _ => {
                                // Skip unknown element
                                skip_element(reader)?;
                            }
                        }
                    }
                    Event::Empty(e) => {
                        match e.local_name().as_ref() {
                            b"nvGraphicFramePr" => {
                                f_nv_graphic_frame_pr = Some(Box::new(
                                    CTGraphicalObjectFrameNonVisual::from_xml(reader, &e, true)?,
                                ));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            b"xfrm" => {
                                f_xfrm = Some(Box::new(ooxml_dml::types::Transform2D::from_xml(
                                    reader, &e, true,
                                )?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            b"extLst" => {
                                f_ext_lst = Some(Box::new(CTExtensionListModify::from_xml(
                                    reader, &e, true,
                                )?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            #[cfg(feature = "extra-children")]
                            _ => {
                                // Capture unknown empty element for roundtrip
                                let elem = RawXmlElement::from_empty(&e);
                                extra_children.push(PositionedNode::new(
                                    child_idx,
                                    RawXmlNode::Element(elem),
                                ));
                                child_idx += 1;
                            }
                            #[cfg(not(feature = "extra-children"))]
                            _ => {}
                        }
                    }
                    Event::End(_) => break,
                    Event::Eof => break,
                    _ => {}
                }
                buf.clear();
            }
        }

        Ok(Self {
            bw_mode: f_bw_mode,
            nv_graphic_frame_pr: f_nv_graphic_frame_pr
                .ok_or_else(|| ParseError::MissingAttribute("nvGraphicFramePr".to_string()))?,
            xfrm: f_xfrm.ok_or_else(|| ParseError::MissingAttribute("xfrm".to_string()))?,
            ext_lst: f_ext_lst,
            #[cfg(feature = "extra-attrs")]
            extra_attrs,
            #[cfg(feature = "extra-children")]
            extra_children,
        })
    }
}

impl FromXml for CTGroupShapeNonVisual {
    fn from_xml<R: BufRead>(
        reader: &mut Reader<R>,
        start_tag: &BytesStart,
        is_empty: bool,
    ) -> Result<Self, ParseError> {
        let mut f_c_nv_pr: Option<Box<ooxml_dml::types::CTNonVisualDrawingProps>> = None;
        let mut f_c_nv_grp_sp_pr: Option<Box<ooxml_dml::types::CTNonVisualGroupDrawingShapeProps>> =
            None;
        let mut f_nv_pr: Option<Box<CTApplicationNonVisualDrawingProps>> = None;
        #[cfg(feature = "extra-children")]
        let mut extra_children = Vec::new();
        #[cfg(feature = "extra-children")]
        let mut child_idx: usize = 0;

        // Parse child elements
        if !is_empty {
            let mut buf = Vec::new();
            loop {
                match reader.read_event_into(&mut buf)? {
                    Event::Start(e) => {
                        match e.local_name().as_ref() {
                            b"cNvPr" => {
                                f_c_nv_pr = Some(Box::new(
                                    ooxml_dml::types::CTNonVisualDrawingProps::from_xml(
                                        reader, &e, false,
                                    )?,
                                ));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            b"cNvGrpSpPr" => {
                                f_c_nv_grp_sp_pr = Some(Box::new(
                                    ooxml_dml::types::CTNonVisualGroupDrawingShapeProps::from_xml(
                                        reader, &e, false,
                                    )?,
                                ));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            b"nvPr" => {
                                f_nv_pr =
                                    Some(Box::new(CTApplicationNonVisualDrawingProps::from_xml(
                                        reader, &e, false,
                                    )?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            #[cfg(feature = "extra-children")]
                            _ => {
                                // Capture unknown element for roundtrip
                                let elem = RawXmlElement::from_reader(reader, &e)?;
                                extra_children.push(PositionedNode::new(
                                    child_idx,
                                    RawXmlNode::Element(elem),
                                ));
                                child_idx += 1;
                            }
                            #[cfg(not(feature = "extra-children"))]
                            _ => {
                                // Skip unknown element
                                skip_element(reader)?;
                            }
                        }
                    }
                    Event::Empty(e) => {
                        match e.local_name().as_ref() {
                            b"cNvPr" => {
                                f_c_nv_pr = Some(Box::new(
                                    ooxml_dml::types::CTNonVisualDrawingProps::from_xml(
                                        reader, &e, true,
                                    )?,
                                ));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            b"cNvGrpSpPr" => {
                                f_c_nv_grp_sp_pr = Some(Box::new(
                                    ooxml_dml::types::CTNonVisualGroupDrawingShapeProps::from_xml(
                                        reader, &e, true,
                                    )?,
                                ));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            b"nvPr" => {
                                f_nv_pr = Some(Box::new(
                                    CTApplicationNonVisualDrawingProps::from_xml(reader, &e, true)?,
                                ));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            #[cfg(feature = "extra-children")]
                            _ => {
                                // Capture unknown empty element for roundtrip
                                let elem = RawXmlElement::from_empty(&e);
                                extra_children.push(PositionedNode::new(
                                    child_idx,
                                    RawXmlNode::Element(elem),
                                ));
                                child_idx += 1;
                            }
                            #[cfg(not(feature = "extra-children"))]
                            _ => {}
                        }
                    }
                    Event::End(_) => break,
                    Event::Eof => break,
                    _ => {}
                }
                buf.clear();
            }
        }

        Ok(Self {
            c_nv_pr: f_c_nv_pr.ok_or_else(|| ParseError::MissingAttribute("cNvPr".to_string()))?,
            c_nv_grp_sp_pr: f_c_nv_grp_sp_pr
                .ok_or_else(|| ParseError::MissingAttribute("cNvGrpSpPr".to_string()))?,
            nv_pr: f_nv_pr.ok_or_else(|| ParseError::MissingAttribute("nvPr".to_string()))?,
            #[cfg(feature = "extra-children")]
            extra_children,
        })
    }
}

impl FromXml for GroupShape {
    fn from_xml<R: BufRead>(
        reader: &mut Reader<R>,
        start_tag: &BytesStart,
        is_empty: bool,
    ) -> Result<Self, ParseError> {
        let mut f_non_visual_group_properties: Option<Box<CTGroupShapeNonVisual>> = None;
        let mut f_grp_sp_pr: Option<Box<ooxml_dml::types::CTGroupShapeProperties>> = None;
        let mut f_shape = None;
        let mut f_group_shape = None;
        let mut f_graphic_frame = None;
        let mut f_connector = None;
        let mut f_picture = None;
        let mut f_content_part = None;
        let mut f_ext_lst = None;
        #[cfg(feature = "extra-children")]
        let mut extra_children = Vec::new();
        #[cfg(feature = "extra-children")]
        let mut child_idx: usize = 0;

        // Parse child elements
        if !is_empty {
            let mut buf = Vec::new();
            loop {
                match reader.read_event_into(&mut buf)? {
                    Event::Start(e) => {
                        match e.local_name().as_ref() {
                            b"nvGrpSpPr" => {
                                f_non_visual_group_properties = Some(Box::new(
                                    CTGroupShapeNonVisual::from_xml(reader, &e, false)?,
                                ));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            b"grpSpPr" => {
                                f_grp_sp_pr = Some(Box::new(
                                    ooxml_dml::types::CTGroupShapeProperties::from_xml(
                                        reader, &e, false,
                                    )?,
                                ));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            b"sp" => {
                                f_shape = Some(Box::new(Shape::from_xml(reader, &e, false)?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            b"grpSp" => {
                                f_group_shape =
                                    Some(Box::new(GroupShape::from_xml(reader, &e, false)?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            b"graphicFrame" => {
                                f_graphic_frame = Some(Box::new(GraphicalObjectFrame::from_xml(
                                    reader, &e, false,
                                )?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            b"cxnSp" => {
                                f_connector =
                                    Some(Box::new(Connector::from_xml(reader, &e, false)?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            b"pic" => {
                                f_picture = Some(Box::new(Picture::from_xml(reader, &e, false)?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            b"contentPart" => {
                                f_content_part =
                                    Some(Box::new(CTRel::from_xml(reader, &e, false)?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            b"extLst" => {
                                f_ext_lst = Some(Box::new(CTExtensionListModify::from_xml(
                                    reader, &e, false,
                                )?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            #[cfg(feature = "extra-children")]
                            _ => {
                                // Capture unknown element for roundtrip
                                let elem = RawXmlElement::from_reader(reader, &e)?;
                                extra_children.push(PositionedNode::new(
                                    child_idx,
                                    RawXmlNode::Element(elem),
                                ));
                                child_idx += 1;
                            }
                            #[cfg(not(feature = "extra-children"))]
                            _ => {
                                // Skip unknown element
                                skip_element(reader)?;
                            }
                        }
                    }
                    Event::Empty(e) => {
                        match e.local_name().as_ref() {
                            b"nvGrpSpPr" => {
                                f_non_visual_group_properties = Some(Box::new(
                                    CTGroupShapeNonVisual::from_xml(reader, &e, true)?,
                                ));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            b"grpSpPr" => {
                                f_grp_sp_pr = Some(Box::new(
                                    ooxml_dml::types::CTGroupShapeProperties::from_xml(
                                        reader, &e, true,
                                    )?,
                                ));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            b"sp" => {
                                f_shape = Some(Box::new(Shape::from_xml(reader, &e, true)?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            b"grpSp" => {
                                f_group_shape =
                                    Some(Box::new(GroupShape::from_xml(reader, &e, true)?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            b"graphicFrame" => {
                                f_graphic_frame = Some(Box::new(GraphicalObjectFrame::from_xml(
                                    reader, &e, true,
                                )?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            b"cxnSp" => {
                                f_connector =
                                    Some(Box::new(Connector::from_xml(reader, &e, true)?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            b"pic" => {
                                f_picture = Some(Box::new(Picture::from_xml(reader, &e, true)?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            b"contentPart" => {
                                f_content_part = Some(Box::new(CTRel::from_xml(reader, &e, true)?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            b"extLst" => {
                                f_ext_lst = Some(Box::new(CTExtensionListModify::from_xml(
                                    reader, &e, true,
                                )?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            #[cfg(feature = "extra-children")]
                            _ => {
                                // Capture unknown empty element for roundtrip
                                let elem = RawXmlElement::from_empty(&e);
                                extra_children.push(PositionedNode::new(
                                    child_idx,
                                    RawXmlNode::Element(elem),
                                ));
                                child_idx += 1;
                            }
                            #[cfg(not(feature = "extra-children"))]
                            _ => {}
                        }
                    }
                    Event::End(_) => break,
                    Event::Eof => break,
                    _ => {}
                }
                buf.clear();
            }
        }

        Ok(Self {
            non_visual_group_properties: f_non_visual_group_properties
                .ok_or_else(|| ParseError::MissingAttribute("nvGrpSpPr".to_string()))?,
            grp_sp_pr: f_grp_sp_pr
                .ok_or_else(|| ParseError::MissingAttribute("grpSpPr".to_string()))?,
            shape: f_shape,
            group_shape: f_group_shape,
            graphic_frame: f_graphic_frame,
            connector: f_connector,
            picture: f_picture,
            content_part: f_content_part,
            ext_lst: f_ext_lst,
            #[cfg(feature = "extra-children")]
            extra_children,
        })
    }
}

impl FromXml for CTRel {
    fn from_xml<R: BufRead>(
        reader: &mut Reader<R>,
        start_tag: &BytesStart,
        is_empty: bool,
    ) -> Result<Self, ParseError> {
        #[cfg(feature = "extra-children")]
        let mut extra_children = Vec::new();
        #[cfg(feature = "extra-children")]
        let mut child_idx: usize = 0;
        if !is_empty {
            let mut buf = Vec::new();
            loop {
                match reader.read_event_into(&mut buf)? {
                    #[cfg(feature = "extra-children")]
                    Event::Start(e) => {
                        let elem = RawXmlElement::from_reader(reader, &e)?;
                        extra_children
                            .push(PositionedNode::new(child_idx, RawXmlNode::Element(elem)));
                        child_idx += 1;
                    }
                    #[cfg(not(feature = "extra-children"))]
                    Event::Start(_) => {
                        skip_element(reader)?;
                    }
                    #[cfg(feature = "extra-children")]
                    Event::Empty(e) => {
                        let elem = RawXmlElement::from_empty(&e);
                        extra_children
                            .push(PositionedNode::new(child_idx, RawXmlNode::Element(elem)));
                        child_idx += 1;
                    }
                    #[cfg(not(feature = "extra-children"))]
                    Event::Empty(_) => {}
                    Event::End(_) => break,
                    Event::Eof => break,
                    _ => {}
                }
                buf.clear();
            }
        }
        Ok(Self {
            #[cfg(feature = "extra-children")]
            extra_children,
        })
    }
}

impl FromXml for EGChildSlide {
    fn from_xml<R: BufRead>(
        reader: &mut Reader<R>,
        start_tag: &BytesStart,
        is_empty: bool,
    ) -> Result<Self, ParseError> {
        let mut f_clr_map_ovr = None;
        #[cfg(feature = "extra-children")]
        let mut extra_children = Vec::new();
        #[cfg(feature = "extra-children")]
        let mut child_idx: usize = 0;

        // Parse child elements
        if !is_empty {
            let mut buf = Vec::new();
            loop {
                match reader.read_event_into(&mut buf)? {
                    Event::Start(e) => {
                        match e.local_name().as_ref() {
                            b"clrMapOvr" => {
                                f_clr_map_ovr = Some(Box::new(
                                    ooxml_dml::types::CTColorMappingOverride::from_xml(
                                        reader, &e, false,
                                    )?,
                                ));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            #[cfg(feature = "extra-children")]
                            _ => {
                                // Capture unknown element for roundtrip
                                let elem = RawXmlElement::from_reader(reader, &e)?;
                                extra_children.push(PositionedNode::new(
                                    child_idx,
                                    RawXmlNode::Element(elem),
                                ));
                                child_idx += 1;
                            }
                            #[cfg(not(feature = "extra-children"))]
                            _ => {
                                // Skip unknown element
                                skip_element(reader)?;
                            }
                        }
                    }
                    Event::Empty(e) => {
                        match e.local_name().as_ref() {
                            b"clrMapOvr" => {
                                f_clr_map_ovr = Some(Box::new(
                                    ooxml_dml::types::CTColorMappingOverride::from_xml(
                                        reader, &e, true,
                                    )?,
                                ));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            #[cfg(feature = "extra-children")]
                            _ => {
                                // Capture unknown empty element for roundtrip
                                let elem = RawXmlElement::from_empty(&e);
                                extra_children.push(PositionedNode::new(
                                    child_idx,
                                    RawXmlNode::Element(elem),
                                ));
                                child_idx += 1;
                            }
                            #[cfg(not(feature = "extra-children"))]
                            _ => {}
                        }
                    }
                    Event::End(_) => break,
                    Event::Eof => break,
                    _ => {}
                }
                buf.clear();
            }
        }

        Ok(Self {
            clr_map_ovr: f_clr_map_ovr,
            #[cfg(feature = "extra-children")]
            extra_children,
        })
    }
}

impl FromXml for PAGChildSlide {
    fn from_xml<R: BufRead>(
        reader: &mut Reader<R>,
        start_tag: &BytesStart,
        is_empty: bool,
    ) -> Result<Self, ParseError> {
        let mut f_show_master_sp = None;
        let mut f_show_master_ph_anim = None;
        #[cfg(feature = "extra-attrs")]
        let mut extra_attrs = std::collections::HashMap::new();

        // Parse attributes
        for attr in start_tag.attributes().filter_map(|a| a.ok()) {
            let val = String::from_utf8_lossy(&attr.value);
            match attr.key.local_name().as_ref() {
                b"showMasterSp" => {
                    f_show_master_sp = Some(val == "true" || val == "1");
                }
                b"showMasterPhAnim" => {
                    f_show_master_ph_anim = Some(val == "true" || val == "1");
                }
                #[cfg(feature = "extra-attrs")]
                unknown => {
                    let key = String::from_utf8_lossy(attr.key.as_ref()).into_owned();
                    extra_attrs.insert(key, val.into_owned());
                }
                #[cfg(not(feature = "extra-attrs"))]
                _ => {}
            }
        }

        if !is_empty {
            let mut buf = Vec::new();
            loop {
                match reader.read_event_into(&mut buf)? {
                    Event::End(_) => break,
                    Event::Eof => break,
                    _ => {}
                }
                buf.clear();
            }
        }

        Ok(Self {
            show_master_sp: f_show_master_sp,
            show_master_ph_anim: f_show_master_ph_anim,
            #[cfg(feature = "extra-attrs")]
            extra_attrs,
        })
    }
}

impl FromXml for CTBackgroundProperties {
    fn from_xml<R: BufRead>(
        reader: &mut Reader<R>,
        start_tag: &BytesStart,
        is_empty: bool,
    ) -> Result<Self, ParseError> {
        let mut f_shade_to_title = None;
        let mut f_ext_lst = None;
        #[cfg(feature = "extra-attrs")]
        let mut extra_attrs = std::collections::HashMap::new();
        #[cfg(feature = "extra-children")]
        let mut extra_children = Vec::new();
        #[cfg(feature = "extra-children")]
        let mut child_idx: usize = 0;

        // Parse attributes
        for attr in start_tag.attributes().filter_map(|a| a.ok()) {
            let val = String::from_utf8_lossy(&attr.value);
            match attr.key.local_name().as_ref() {
                b"shadeToTitle" => {
                    f_shade_to_title = Some(val == "true" || val == "1");
                }
                #[cfg(feature = "extra-attrs")]
                unknown => {
                    let key = String::from_utf8_lossy(attr.key.as_ref()).into_owned();
                    extra_attrs.insert(key, val.into_owned());
                }
                #[cfg(not(feature = "extra-attrs"))]
                _ => {}
            }
        }

        // Parse child elements
        if !is_empty {
            let mut buf = Vec::new();
            loop {
                match reader.read_event_into(&mut buf)? {
                    Event::Start(e) => {
                        match e.local_name().as_ref() {
                            b"extLst" => {
                                f_ext_lst =
                                    Some(Box::new(CTExtensionList::from_xml(reader, &e, false)?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            #[cfg(feature = "extra-children")]
                            _ => {
                                // Capture unknown element for roundtrip
                                let elem = RawXmlElement::from_reader(reader, &e)?;
                                extra_children.push(PositionedNode::new(
                                    child_idx,
                                    RawXmlNode::Element(elem),
                                ));
                                child_idx += 1;
                            }
                            #[cfg(not(feature = "extra-children"))]
                            _ => {
                                // Skip unknown element
                                skip_element(reader)?;
                            }
                        }
                    }
                    Event::Empty(e) => {
                        match e.local_name().as_ref() {
                            b"extLst" => {
                                f_ext_lst =
                                    Some(Box::new(CTExtensionList::from_xml(reader, &e, true)?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            #[cfg(feature = "extra-children")]
                            _ => {
                                // Capture unknown empty element for roundtrip
                                let elem = RawXmlElement::from_empty(&e);
                                extra_children.push(PositionedNode::new(
                                    child_idx,
                                    RawXmlNode::Element(elem),
                                ));
                                child_idx += 1;
                            }
                            #[cfg(not(feature = "extra-children"))]
                            _ => {}
                        }
                    }
                    Event::End(_) => break,
                    Event::Eof => break,
                    _ => {}
                }
                buf.clear();
            }
        }

        Ok(Self {
            shade_to_title: f_shade_to_title,
            ext_lst: f_ext_lst,
            #[cfg(feature = "extra-attrs")]
            extra_attrs,
            #[cfg(feature = "extra-children")]
            extra_children,
        })
    }
}

impl FromXml for EGBackground {
    fn from_xml<R: BufRead>(
        reader: &mut Reader<R>,
        start_tag: &BytesStart,
        is_empty: bool,
    ) -> Result<Self, ParseError> {
        let tag = start_tag.local_name();
        match tag.as_ref() {
            b"bgPr" => {
                let inner = CTBackgroundProperties::from_xml(reader, start_tag, is_empty)?;
                Ok(Self::BgPr(Box::new(inner)))
            }
            b"bgRef" => {
                let inner = ooxml_dml::types::CTStyleMatrixReference::from_xml(
                    reader, start_tag, is_empty,
                )?;
                Ok(Self::BgRef(Box::new(inner)))
            }
            _ => Err(ParseError::UnexpectedElement(
                String::from_utf8_lossy(start_tag.name().as_ref()).into_owned(),
            )),
        }
    }
}

impl FromXml for CTBackground {
    fn from_xml<R: BufRead>(
        reader: &mut Reader<R>,
        start_tag: &BytesStart,
        is_empty: bool,
    ) -> Result<Self, ParseError> {
        let mut f_bw_mode = None;
        let mut f_background: Option<Box<EGBackground>> = None;
        #[cfg(feature = "extra-attrs")]
        let mut extra_attrs = std::collections::HashMap::new();
        #[cfg(feature = "extra-children")]
        let mut extra_children = Vec::new();
        #[cfg(feature = "extra-children")]
        let mut child_idx: usize = 0;

        // Parse attributes
        for attr in start_tag.attributes().filter_map(|a| a.ok()) {
            let val = String::from_utf8_lossy(&attr.value);
            match attr.key.local_name().as_ref() {
                b"bwMode" => {
                    f_bw_mode = val.parse().ok();
                }
                #[cfg(feature = "extra-attrs")]
                unknown => {
                    let key = String::from_utf8_lossy(attr.key.as_ref()).into_owned();
                    extra_attrs.insert(key, val.into_owned());
                }
                #[cfg(not(feature = "extra-attrs"))]
                _ => {}
            }
        }

        // Parse child elements
        if !is_empty {
            let mut buf = Vec::new();
            loop {
                match reader.read_event_into(&mut buf)? {
                    Event::Start(e) => {
                        match e.local_name().as_ref() {
                            b"bgPr" | b"bgRef" => {
                                f_background =
                                    Some(Box::new(EGBackground::from_xml(reader, &e, false)?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            #[cfg(feature = "extra-children")]
                            _ => {
                                // Capture unknown element for roundtrip
                                let elem = RawXmlElement::from_reader(reader, &e)?;
                                extra_children.push(PositionedNode::new(
                                    child_idx,
                                    RawXmlNode::Element(elem),
                                ));
                                child_idx += 1;
                            }
                            #[cfg(not(feature = "extra-children"))]
                            _ => {
                                // Skip unknown element
                                skip_element(reader)?;
                            }
                        }
                    }
                    Event::Empty(e) => {
                        match e.local_name().as_ref() {
                            b"bgPr" | b"bgRef" => {
                                f_background =
                                    Some(Box::new(EGBackground::from_xml(reader, &e, true)?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            #[cfg(feature = "extra-children")]
                            _ => {
                                // Capture unknown empty element for roundtrip
                                let elem = RawXmlElement::from_empty(&e);
                                extra_children.push(PositionedNode::new(
                                    child_idx,
                                    RawXmlNode::Element(elem),
                                ));
                                child_idx += 1;
                            }
                            #[cfg(not(feature = "extra-children"))]
                            _ => {}
                        }
                    }
                    Event::End(_) => break,
                    Event::Eof => break,
                    _ => {}
                }
                buf.clear();
            }
        }

        Ok(Self {
            bw_mode: f_bw_mode,
            background: f_background,
            #[cfg(feature = "extra-attrs")]
            extra_attrs,
            #[cfg(feature = "extra-children")]
            extra_children,
        })
    }
}

impl FromXml for CommonSlideData {
    fn from_xml<R: BufRead>(
        reader: &mut Reader<R>,
        start_tag: &BytesStart,
        is_empty: bool,
    ) -> Result<Self, ParseError> {
        let mut f_name = None;
        let mut f_bg = None;
        let mut f_shape_tree: Option<Box<GroupShape>> = None;
        let mut f_cust_data_lst = None;
        let mut f_controls = None;
        let mut f_ext_lst = None;
        #[cfg(feature = "extra-attrs")]
        let mut extra_attrs = std::collections::HashMap::new();
        #[cfg(feature = "extra-children")]
        let mut extra_children = Vec::new();
        #[cfg(feature = "extra-children")]
        let mut child_idx: usize = 0;

        // Parse attributes
        for attr in start_tag.attributes().filter_map(|a| a.ok()) {
            let val = String::from_utf8_lossy(&attr.value);
            match attr.key.local_name().as_ref() {
                b"name" => {
                    f_name = Some(val.into_owned());
                }
                #[cfg(feature = "extra-attrs")]
                unknown => {
                    let key = String::from_utf8_lossy(attr.key.as_ref()).into_owned();
                    extra_attrs.insert(key, val.into_owned());
                }
                #[cfg(not(feature = "extra-attrs"))]
                _ => {}
            }
        }

        // Parse child elements
        if !is_empty {
            let mut buf = Vec::new();
            loop {
                match reader.read_event_into(&mut buf)? {
                    Event::Start(e) => {
                        match e.local_name().as_ref() {
                            b"bg" => {
                                f_bg = Some(Box::new(CTBackground::from_xml(reader, &e, false)?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            b"spTree" => {
                                f_shape_tree =
                                    Some(Box::new(GroupShape::from_xml(reader, &e, false)?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            b"custDataLst" => {
                                f_cust_data_lst = Some(Box::new(CTCustomerDataList::from_xml(
                                    reader, &e, false,
                                )?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            b"controls" => {
                                f_controls =
                                    Some(Box::new(CTControlList::from_xml(reader, &e, false)?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            b"extLst" => {
                                f_ext_lst =
                                    Some(Box::new(CTExtensionList::from_xml(reader, &e, false)?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            #[cfg(feature = "extra-children")]
                            _ => {
                                // Capture unknown element for roundtrip
                                let elem = RawXmlElement::from_reader(reader, &e)?;
                                extra_children.push(PositionedNode::new(
                                    child_idx,
                                    RawXmlNode::Element(elem),
                                ));
                                child_idx += 1;
                            }
                            #[cfg(not(feature = "extra-children"))]
                            _ => {
                                // Skip unknown element
                                skip_element(reader)?;
                            }
                        }
                    }
                    Event::Empty(e) => {
                        match e.local_name().as_ref() {
                            b"bg" => {
                                f_bg = Some(Box::new(CTBackground::from_xml(reader, &e, true)?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            b"spTree" => {
                                f_shape_tree =
                                    Some(Box::new(GroupShape::from_xml(reader, &e, true)?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            b"custDataLst" => {
                                f_cust_data_lst =
                                    Some(Box::new(CTCustomerDataList::from_xml(reader, &e, true)?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            b"controls" => {
                                f_controls =
                                    Some(Box::new(CTControlList::from_xml(reader, &e, true)?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            b"extLst" => {
                                f_ext_lst =
                                    Some(Box::new(CTExtensionList::from_xml(reader, &e, true)?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            #[cfg(feature = "extra-children")]
                            _ => {
                                // Capture unknown empty element for roundtrip
                                let elem = RawXmlElement::from_empty(&e);
                                extra_children.push(PositionedNode::new(
                                    child_idx,
                                    RawXmlNode::Element(elem),
                                ));
                                child_idx += 1;
                            }
                            #[cfg(not(feature = "extra-children"))]
                            _ => {}
                        }
                    }
                    Event::End(_) => break,
                    Event::Eof => break,
                    _ => {}
                }
                buf.clear();
            }
        }

        Ok(Self {
            name: f_name,
            bg: f_bg,
            shape_tree: f_shape_tree
                .ok_or_else(|| ParseError::MissingAttribute("spTree".to_string()))?,
            cust_data_lst: f_cust_data_lst,
            controls: f_controls,
            ext_lst: f_ext_lst,
            #[cfg(feature = "extra-attrs")]
            extra_attrs,
            #[cfg(feature = "extra-children")]
            extra_children,
        })
    }
}

impl FromXml for Slide {
    fn from_xml<R: BufRead>(
        reader: &mut Reader<R>,
        start_tag: &BytesStart,
        is_empty: bool,
    ) -> Result<Self, ParseError> {
        let mut f_show_master_sp = None;
        let mut f_show_master_ph_anim = None;
        let mut f_show = None;
        let mut f_common_slide_data: Option<Box<CommonSlideData>> = None;
        let mut f_clr_map_ovr = None;
        #[cfg(feature = "pml-transitions")]
        let mut f_transition = None;
        #[cfg(feature = "pml-animations")]
        let mut f_timing = None;
        let mut f_ext_lst = None;
        #[cfg(feature = "extra-attrs")]
        let mut extra_attrs = std::collections::HashMap::new();
        #[cfg(feature = "extra-children")]
        let mut extra_children = Vec::new();
        #[cfg(feature = "extra-children")]
        let mut child_idx: usize = 0;

        // Parse attributes
        for attr in start_tag.attributes().filter_map(|a| a.ok()) {
            let val = String::from_utf8_lossy(&attr.value);
            match attr.key.local_name().as_ref() {
                b"showMasterSp" => {
                    f_show_master_sp = Some(val == "true" || val == "1");
                }
                b"showMasterPhAnim" => {
                    f_show_master_ph_anim = Some(val == "true" || val == "1");
                }
                b"show" => {
                    f_show = Some(val == "true" || val == "1");
                }
                #[cfg(feature = "extra-attrs")]
                unknown => {
                    let key = String::from_utf8_lossy(attr.key.as_ref()).into_owned();
                    extra_attrs.insert(key, val.into_owned());
                }
                #[cfg(not(feature = "extra-attrs"))]
                _ => {}
            }
        }

        // Parse child elements
        if !is_empty {
            let mut buf = Vec::new();
            loop {
                match reader.read_event_into(&mut buf)? {
                    Event::Start(e) => {
                        match e.local_name().as_ref() {
                            b"cSld" => {
                                f_common_slide_data =
                                    Some(Box::new(CommonSlideData::from_xml(reader, &e, false)?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            b"clrMapOvr" => {
                                f_clr_map_ovr = Some(Box::new(
                                    ooxml_dml::types::CTColorMappingOverride::from_xml(
                                        reader, &e, false,
                                    )?,
                                ));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            #[cfg(feature = "pml-transitions")]
                            b"transition" => {
                                f_transition =
                                    Some(Box::new(SlideTransition::from_xml(reader, &e, false)?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            #[cfg(feature = "pml-animations")]
                            b"timing" => {
                                f_timing =
                                    Some(Box::new(SlideTiming::from_xml(reader, &e, false)?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            b"extLst" => {
                                f_ext_lst = Some(Box::new(CTExtensionListModify::from_xml(
                                    reader, &e, false,
                                )?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            #[cfg(feature = "extra-children")]
                            _ => {
                                // Capture unknown element for roundtrip
                                let elem = RawXmlElement::from_reader(reader, &e)?;
                                extra_children.push(PositionedNode::new(
                                    child_idx,
                                    RawXmlNode::Element(elem),
                                ));
                                child_idx += 1;
                            }
                            #[cfg(not(feature = "extra-children"))]
                            _ => {
                                // Skip unknown element
                                skip_element(reader)?;
                            }
                        }
                    }
                    Event::Empty(e) => {
                        match e.local_name().as_ref() {
                            b"cSld" => {
                                f_common_slide_data =
                                    Some(Box::new(CommonSlideData::from_xml(reader, &e, true)?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            b"clrMapOvr" => {
                                f_clr_map_ovr = Some(Box::new(
                                    ooxml_dml::types::CTColorMappingOverride::from_xml(
                                        reader, &e, true,
                                    )?,
                                ));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            #[cfg(feature = "pml-transitions")]
                            b"transition" => {
                                f_transition =
                                    Some(Box::new(SlideTransition::from_xml(reader, &e, true)?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            #[cfg(feature = "pml-animations")]
                            b"timing" => {
                                f_timing = Some(Box::new(SlideTiming::from_xml(reader, &e, true)?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            b"extLst" => {
                                f_ext_lst = Some(Box::new(CTExtensionListModify::from_xml(
                                    reader, &e, true,
                                )?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            #[cfg(feature = "extra-children")]
                            _ => {
                                // Capture unknown empty element for roundtrip
                                let elem = RawXmlElement::from_empty(&e);
                                extra_children.push(PositionedNode::new(
                                    child_idx,
                                    RawXmlNode::Element(elem),
                                ));
                                child_idx += 1;
                            }
                            #[cfg(not(feature = "extra-children"))]
                            _ => {}
                        }
                    }
                    Event::End(_) => break,
                    Event::Eof => break,
                    _ => {}
                }
                buf.clear();
            }
        }

        Ok(Self {
            show_master_sp: f_show_master_sp,
            show_master_ph_anim: f_show_master_ph_anim,
            show: f_show,
            common_slide_data: f_common_slide_data
                .ok_or_else(|| ParseError::MissingAttribute("cSld".to_string()))?,
            clr_map_ovr: f_clr_map_ovr,
            #[cfg(feature = "pml-transitions")]
            transition: f_transition,
            #[cfg(feature = "pml-animations")]
            timing: f_timing,
            ext_lst: f_ext_lst,
            #[cfg(feature = "extra-attrs")]
            extra_attrs,
            #[cfg(feature = "extra-children")]
            extra_children,
        })
    }
}

impl FromXml for SlideLayout {
    fn from_xml<R: BufRead>(
        reader: &mut Reader<R>,
        start_tag: &BytesStart,
        is_empty: bool,
    ) -> Result<Self, ParseError> {
        let mut f_show_master_sp = None;
        let mut f_show_master_ph_anim = None;
        let mut f_matching_name = None;
        let mut f_type = None;
        let mut f_preserve = None;
        let mut f_user_drawn = None;
        let mut f_common_slide_data: Option<Box<CommonSlideData>> = None;
        let mut f_clr_map_ovr = None;
        let mut f_transition = None;
        let mut f_timing = None;
        let mut f_hf = None;
        let mut f_ext_lst = None;
        #[cfg(feature = "extra-attrs")]
        let mut extra_attrs = std::collections::HashMap::new();
        #[cfg(feature = "extra-children")]
        let mut extra_children = Vec::new();
        #[cfg(feature = "extra-children")]
        let mut child_idx: usize = 0;

        // Parse attributes
        for attr in start_tag.attributes().filter_map(|a| a.ok()) {
            let val = String::from_utf8_lossy(&attr.value);
            match attr.key.local_name().as_ref() {
                b"showMasterSp" => {
                    f_show_master_sp = Some(val == "true" || val == "1");
                }
                b"showMasterPhAnim" => {
                    f_show_master_ph_anim = Some(val == "true" || val == "1");
                }
                b"matchingName" => {
                    f_matching_name = Some(val.into_owned());
                }
                b"type" => {
                    f_type = val.parse().ok();
                }
                b"preserve" => {
                    f_preserve = Some(val == "true" || val == "1");
                }
                b"userDrawn" => {
                    f_user_drawn = Some(val == "true" || val == "1");
                }
                #[cfg(feature = "extra-attrs")]
                unknown => {
                    let key = String::from_utf8_lossy(attr.key.as_ref()).into_owned();
                    extra_attrs.insert(key, val.into_owned());
                }
                #[cfg(not(feature = "extra-attrs"))]
                _ => {}
            }
        }

        // Parse child elements
        if !is_empty {
            let mut buf = Vec::new();
            loop {
                match reader.read_event_into(&mut buf)? {
                    Event::Start(e) => {
                        match e.local_name().as_ref() {
                            b"cSld" => {
                                f_common_slide_data =
                                    Some(Box::new(CommonSlideData::from_xml(reader, &e, false)?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            b"clrMapOvr" => {
                                f_clr_map_ovr = Some(Box::new(
                                    ooxml_dml::types::CTColorMappingOverride::from_xml(
                                        reader, &e, false,
                                    )?,
                                ));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            b"transition" => {
                                f_transition =
                                    Some(Box::new(SlideTransition::from_xml(reader, &e, false)?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            b"timing" => {
                                f_timing =
                                    Some(Box::new(SlideTiming::from_xml(reader, &e, false)?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            b"hf" => {
                                f_hf = Some(Box::new(CTHeaderFooter::from_xml(reader, &e, false)?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            b"extLst" => {
                                f_ext_lst = Some(Box::new(CTExtensionListModify::from_xml(
                                    reader, &e, false,
                                )?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            #[cfg(feature = "extra-children")]
                            _ => {
                                // Capture unknown element for roundtrip
                                let elem = RawXmlElement::from_reader(reader, &e)?;
                                extra_children.push(PositionedNode::new(
                                    child_idx,
                                    RawXmlNode::Element(elem),
                                ));
                                child_idx += 1;
                            }
                            #[cfg(not(feature = "extra-children"))]
                            _ => {
                                // Skip unknown element
                                skip_element(reader)?;
                            }
                        }
                    }
                    Event::Empty(e) => {
                        match e.local_name().as_ref() {
                            b"cSld" => {
                                f_common_slide_data =
                                    Some(Box::new(CommonSlideData::from_xml(reader, &e, true)?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            b"clrMapOvr" => {
                                f_clr_map_ovr = Some(Box::new(
                                    ooxml_dml::types::CTColorMappingOverride::from_xml(
                                        reader, &e, true,
                                    )?,
                                ));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            b"transition" => {
                                f_transition =
                                    Some(Box::new(SlideTransition::from_xml(reader, &e, true)?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            b"timing" => {
                                f_timing = Some(Box::new(SlideTiming::from_xml(reader, &e, true)?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            b"hf" => {
                                f_hf = Some(Box::new(CTHeaderFooter::from_xml(reader, &e, true)?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            b"extLst" => {
                                f_ext_lst = Some(Box::new(CTExtensionListModify::from_xml(
                                    reader, &e, true,
                                )?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            #[cfg(feature = "extra-children")]
                            _ => {
                                // Capture unknown empty element for roundtrip
                                let elem = RawXmlElement::from_empty(&e);
                                extra_children.push(PositionedNode::new(
                                    child_idx,
                                    RawXmlNode::Element(elem),
                                ));
                                child_idx += 1;
                            }
                            #[cfg(not(feature = "extra-children"))]
                            _ => {}
                        }
                    }
                    Event::End(_) => break,
                    Event::Eof => break,
                    _ => {}
                }
                buf.clear();
            }
        }

        Ok(Self {
            show_master_sp: f_show_master_sp,
            show_master_ph_anim: f_show_master_ph_anim,
            matching_name: f_matching_name,
            r#type: f_type,
            preserve: f_preserve,
            user_drawn: f_user_drawn,
            common_slide_data: f_common_slide_data
                .ok_or_else(|| ParseError::MissingAttribute("cSld".to_string()))?,
            clr_map_ovr: f_clr_map_ovr,
            transition: f_transition,
            timing: f_timing,
            hf: f_hf,
            ext_lst: f_ext_lst,
            #[cfg(feature = "extra-attrs")]
            extra_attrs,
            #[cfg(feature = "extra-children")]
            extra_children,
        })
    }
}

impl FromXml for CTSlideMasterTextStyles {
    fn from_xml<R: BufRead>(
        reader: &mut Reader<R>,
        start_tag: &BytesStart,
        is_empty: bool,
    ) -> Result<Self, ParseError> {
        let mut f_title_style = None;
        let mut f_body_style = None;
        let mut f_other_style = None;
        let mut f_ext_lst = None;
        #[cfg(feature = "extra-children")]
        let mut extra_children = Vec::new();
        #[cfg(feature = "extra-children")]
        let mut child_idx: usize = 0;

        // Parse child elements
        if !is_empty {
            let mut buf = Vec::new();
            loop {
                match reader.read_event_into(&mut buf)? {
                    Event::Start(e) => {
                        match e.local_name().as_ref() {
                            b"titleStyle" => {
                                f_title_style = Some(Box::new(
                                    ooxml_dml::types::CTTextListStyle::from_xml(reader, &e, false)?,
                                ));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            b"bodyStyle" => {
                                f_body_style = Some(Box::new(
                                    ooxml_dml::types::CTTextListStyle::from_xml(reader, &e, false)?,
                                ));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            b"otherStyle" => {
                                f_other_style = Some(Box::new(
                                    ooxml_dml::types::CTTextListStyle::from_xml(reader, &e, false)?,
                                ));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            b"extLst" => {
                                f_ext_lst =
                                    Some(Box::new(CTExtensionList::from_xml(reader, &e, false)?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            #[cfg(feature = "extra-children")]
                            _ => {
                                // Capture unknown element for roundtrip
                                let elem = RawXmlElement::from_reader(reader, &e)?;
                                extra_children.push(PositionedNode::new(
                                    child_idx,
                                    RawXmlNode::Element(elem),
                                ));
                                child_idx += 1;
                            }
                            #[cfg(not(feature = "extra-children"))]
                            _ => {
                                // Skip unknown element
                                skip_element(reader)?;
                            }
                        }
                    }
                    Event::Empty(e) => {
                        match e.local_name().as_ref() {
                            b"titleStyle" => {
                                f_title_style = Some(Box::new(
                                    ooxml_dml::types::CTTextListStyle::from_xml(reader, &e, true)?,
                                ));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            b"bodyStyle" => {
                                f_body_style = Some(Box::new(
                                    ooxml_dml::types::CTTextListStyle::from_xml(reader, &e, true)?,
                                ));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            b"otherStyle" => {
                                f_other_style = Some(Box::new(
                                    ooxml_dml::types::CTTextListStyle::from_xml(reader, &e, true)?,
                                ));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            b"extLst" => {
                                f_ext_lst =
                                    Some(Box::new(CTExtensionList::from_xml(reader, &e, true)?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            #[cfg(feature = "extra-children")]
                            _ => {
                                // Capture unknown empty element for roundtrip
                                let elem = RawXmlElement::from_empty(&e);
                                extra_children.push(PositionedNode::new(
                                    child_idx,
                                    RawXmlNode::Element(elem),
                                ));
                                child_idx += 1;
                            }
                            #[cfg(not(feature = "extra-children"))]
                            _ => {}
                        }
                    }
                    Event::End(_) => break,
                    Event::Eof => break,
                    _ => {}
                }
                buf.clear();
            }
        }

        Ok(Self {
            title_style: f_title_style,
            body_style: f_body_style,
            other_style: f_other_style,
            ext_lst: f_ext_lst,
            #[cfg(feature = "extra-children")]
            extra_children,
        })
    }
}

impl FromXml for CTSlideLayoutIdListEntry {
    fn from_xml<R: BufRead>(
        reader: &mut Reader<R>,
        start_tag: &BytesStart,
        is_empty: bool,
    ) -> Result<Self, ParseError> {
        let mut f_id = None;
        let mut f_ext_lst = None;
        #[cfg(feature = "extra-attrs")]
        let mut extra_attrs = std::collections::HashMap::new();
        #[cfg(feature = "extra-children")]
        let mut extra_children = Vec::new();
        #[cfg(feature = "extra-children")]
        let mut child_idx: usize = 0;

        // Parse attributes
        for attr in start_tag.attributes().filter_map(|a| a.ok()) {
            let val = String::from_utf8_lossy(&attr.value);
            match attr.key.local_name().as_ref() {
                b"id" => {
                    f_id = val.parse().ok();
                }
                #[cfg(feature = "extra-attrs")]
                unknown => {
                    let key = String::from_utf8_lossy(attr.key.as_ref()).into_owned();
                    extra_attrs.insert(key, val.into_owned());
                }
                #[cfg(not(feature = "extra-attrs"))]
                _ => {}
            }
        }

        // Parse child elements
        if !is_empty {
            let mut buf = Vec::new();
            loop {
                match reader.read_event_into(&mut buf)? {
                    Event::Start(e) => {
                        match e.local_name().as_ref() {
                            b"extLst" => {
                                f_ext_lst =
                                    Some(Box::new(CTExtensionList::from_xml(reader, &e, false)?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            #[cfg(feature = "extra-children")]
                            _ => {
                                // Capture unknown element for roundtrip
                                let elem = RawXmlElement::from_reader(reader, &e)?;
                                extra_children.push(PositionedNode::new(
                                    child_idx,
                                    RawXmlNode::Element(elem),
                                ));
                                child_idx += 1;
                            }
                            #[cfg(not(feature = "extra-children"))]
                            _ => {
                                // Skip unknown element
                                skip_element(reader)?;
                            }
                        }
                    }
                    Event::Empty(e) => {
                        match e.local_name().as_ref() {
                            b"extLst" => {
                                f_ext_lst =
                                    Some(Box::new(CTExtensionList::from_xml(reader, &e, true)?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            #[cfg(feature = "extra-children")]
                            _ => {
                                // Capture unknown empty element for roundtrip
                                let elem = RawXmlElement::from_empty(&e);
                                extra_children.push(PositionedNode::new(
                                    child_idx,
                                    RawXmlNode::Element(elem),
                                ));
                                child_idx += 1;
                            }
                            #[cfg(not(feature = "extra-children"))]
                            _ => {}
                        }
                    }
                    Event::End(_) => break,
                    Event::Eof => break,
                    _ => {}
                }
                buf.clear();
            }
        }

        Ok(Self {
            id: f_id,
            ext_lst: f_ext_lst,
            #[cfg(feature = "extra-attrs")]
            extra_attrs,
            #[cfg(feature = "extra-children")]
            extra_children,
        })
    }
}

impl FromXml for CTSlideLayoutIdList {
    fn from_xml<R: BufRead>(
        reader: &mut Reader<R>,
        start_tag: &BytesStart,
        is_empty: bool,
    ) -> Result<Self, ParseError> {
        let mut f_sld_layout_id = Vec::new();
        #[cfg(feature = "extra-children")]
        let mut extra_children = Vec::new();
        #[cfg(feature = "extra-children")]
        let mut child_idx: usize = 0;

        // Parse child elements
        if !is_empty {
            let mut buf = Vec::new();
            loop {
                match reader.read_event_into(&mut buf)? {
                    Event::Start(e) => {
                        match e.local_name().as_ref() {
                            b"sldLayoutId" => {
                                f_sld_layout_id
                                    .push(CTSlideLayoutIdListEntry::from_xml(reader, &e, false)?);
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            #[cfg(feature = "extra-children")]
                            _ => {
                                // Capture unknown element for roundtrip
                                let elem = RawXmlElement::from_reader(reader, &e)?;
                                extra_children.push(PositionedNode::new(
                                    child_idx,
                                    RawXmlNode::Element(elem),
                                ));
                                child_idx += 1;
                            }
                            #[cfg(not(feature = "extra-children"))]
                            _ => {
                                // Skip unknown element
                                skip_element(reader)?;
                            }
                        }
                    }
                    Event::Empty(e) => {
                        match e.local_name().as_ref() {
                            b"sldLayoutId" => {
                                f_sld_layout_id
                                    .push(CTSlideLayoutIdListEntry::from_xml(reader, &e, true)?);
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            #[cfg(feature = "extra-children")]
                            _ => {
                                // Capture unknown empty element for roundtrip
                                let elem = RawXmlElement::from_empty(&e);
                                extra_children.push(PositionedNode::new(
                                    child_idx,
                                    RawXmlNode::Element(elem),
                                ));
                                child_idx += 1;
                            }
                            #[cfg(not(feature = "extra-children"))]
                            _ => {}
                        }
                    }
                    Event::End(_) => break,
                    Event::Eof => break,
                    _ => {}
                }
                buf.clear();
            }
        }

        Ok(Self {
            sld_layout_id: f_sld_layout_id,
            #[cfg(feature = "extra-children")]
            extra_children,
        })
    }
}

impl FromXml for SlideMaster {
    fn from_xml<R: BufRead>(
        reader: &mut Reader<R>,
        start_tag: &BytesStart,
        is_empty: bool,
    ) -> Result<Self, ParseError> {
        let mut f_preserve = None;
        let mut f_common_slide_data: Option<Box<CommonSlideData>> = None;
        let mut f_clr_map: Option<Box<ooxml_dml::types::CTColorMapping>> = None;
        let mut f_sld_layout_id_lst = None;
        let mut f_transition = None;
        let mut f_timing = None;
        let mut f_hf = None;
        let mut f_tx_styles = None;
        let mut f_ext_lst = None;
        #[cfg(feature = "extra-attrs")]
        let mut extra_attrs = std::collections::HashMap::new();
        #[cfg(feature = "extra-children")]
        let mut extra_children = Vec::new();
        #[cfg(feature = "extra-children")]
        let mut child_idx: usize = 0;

        // Parse attributes
        for attr in start_tag.attributes().filter_map(|a| a.ok()) {
            let val = String::from_utf8_lossy(&attr.value);
            match attr.key.local_name().as_ref() {
                b"preserve" => {
                    f_preserve = Some(val == "true" || val == "1");
                }
                #[cfg(feature = "extra-attrs")]
                unknown => {
                    let key = String::from_utf8_lossy(attr.key.as_ref()).into_owned();
                    extra_attrs.insert(key, val.into_owned());
                }
                #[cfg(not(feature = "extra-attrs"))]
                _ => {}
            }
        }

        // Parse child elements
        if !is_empty {
            let mut buf = Vec::new();
            loop {
                match reader.read_event_into(&mut buf)? {
                    Event::Start(e) => {
                        match e.local_name().as_ref() {
                            b"cSld" => {
                                f_common_slide_data =
                                    Some(Box::new(CommonSlideData::from_xml(reader, &e, false)?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            b"clrMap" => {
                                f_clr_map = Some(Box::new(
                                    ooxml_dml::types::CTColorMapping::from_xml(reader, &e, false)?,
                                ));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            b"sldLayoutIdLst" => {
                                f_sld_layout_id_lst = Some(Box::new(
                                    CTSlideLayoutIdList::from_xml(reader, &e, false)?,
                                ));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            b"transition" => {
                                f_transition =
                                    Some(Box::new(SlideTransition::from_xml(reader, &e, false)?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            b"timing" => {
                                f_timing =
                                    Some(Box::new(SlideTiming::from_xml(reader, &e, false)?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            b"hf" => {
                                f_hf = Some(Box::new(CTHeaderFooter::from_xml(reader, &e, false)?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            b"txStyles" => {
                                f_tx_styles = Some(Box::new(CTSlideMasterTextStyles::from_xml(
                                    reader, &e, false,
                                )?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            b"extLst" => {
                                f_ext_lst = Some(Box::new(CTExtensionListModify::from_xml(
                                    reader, &e, false,
                                )?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            #[cfg(feature = "extra-children")]
                            _ => {
                                // Capture unknown element for roundtrip
                                let elem = RawXmlElement::from_reader(reader, &e)?;
                                extra_children.push(PositionedNode::new(
                                    child_idx,
                                    RawXmlNode::Element(elem),
                                ));
                                child_idx += 1;
                            }
                            #[cfg(not(feature = "extra-children"))]
                            _ => {
                                // Skip unknown element
                                skip_element(reader)?;
                            }
                        }
                    }
                    Event::Empty(e) => {
                        match e.local_name().as_ref() {
                            b"cSld" => {
                                f_common_slide_data =
                                    Some(Box::new(CommonSlideData::from_xml(reader, &e, true)?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            b"clrMap" => {
                                f_clr_map = Some(Box::new(
                                    ooxml_dml::types::CTColorMapping::from_xml(reader, &e, true)?,
                                ));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            b"sldLayoutIdLst" => {
                                f_sld_layout_id_lst = Some(Box::new(
                                    CTSlideLayoutIdList::from_xml(reader, &e, true)?,
                                ));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            b"transition" => {
                                f_transition =
                                    Some(Box::new(SlideTransition::from_xml(reader, &e, true)?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            b"timing" => {
                                f_timing = Some(Box::new(SlideTiming::from_xml(reader, &e, true)?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            b"hf" => {
                                f_hf = Some(Box::new(CTHeaderFooter::from_xml(reader, &e, true)?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            b"txStyles" => {
                                f_tx_styles = Some(Box::new(CTSlideMasterTextStyles::from_xml(
                                    reader, &e, true,
                                )?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            b"extLst" => {
                                f_ext_lst = Some(Box::new(CTExtensionListModify::from_xml(
                                    reader, &e, true,
                                )?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            #[cfg(feature = "extra-children")]
                            _ => {
                                // Capture unknown empty element for roundtrip
                                let elem = RawXmlElement::from_empty(&e);
                                extra_children.push(PositionedNode::new(
                                    child_idx,
                                    RawXmlNode::Element(elem),
                                ));
                                child_idx += 1;
                            }
                            #[cfg(not(feature = "extra-children"))]
                            _ => {}
                        }
                    }
                    Event::End(_) => break,
                    Event::Eof => break,
                    _ => {}
                }
                buf.clear();
            }
        }

        Ok(Self {
            preserve: f_preserve,
            common_slide_data: f_common_slide_data
                .ok_or_else(|| ParseError::MissingAttribute("cSld".to_string()))?,
            clr_map: f_clr_map.ok_or_else(|| ParseError::MissingAttribute("clrMap".to_string()))?,
            sld_layout_id_lst: f_sld_layout_id_lst,
            transition: f_transition,
            timing: f_timing,
            hf: f_hf,
            tx_styles: f_tx_styles,
            ext_lst: f_ext_lst,
            #[cfg(feature = "extra-attrs")]
            extra_attrs,
            #[cfg(feature = "extra-children")]
            extra_children,
        })
    }
}

impl FromXml for HandoutMaster {
    fn from_xml<R: BufRead>(
        reader: &mut Reader<R>,
        start_tag: &BytesStart,
        is_empty: bool,
    ) -> Result<Self, ParseError> {
        let mut f_common_slide_data: Option<Box<CommonSlideData>> = None;
        let mut f_clr_map: Option<Box<ooxml_dml::types::CTColorMapping>> = None;
        let mut f_hf = None;
        let mut f_ext_lst = None;
        #[cfg(feature = "extra-children")]
        let mut extra_children = Vec::new();
        #[cfg(feature = "extra-children")]
        let mut child_idx: usize = 0;

        // Parse child elements
        if !is_empty {
            let mut buf = Vec::new();
            loop {
                match reader.read_event_into(&mut buf)? {
                    Event::Start(e) => {
                        match e.local_name().as_ref() {
                            b"cSld" => {
                                f_common_slide_data =
                                    Some(Box::new(CommonSlideData::from_xml(reader, &e, false)?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            b"clrMap" => {
                                f_clr_map = Some(Box::new(
                                    ooxml_dml::types::CTColorMapping::from_xml(reader, &e, false)?,
                                ));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            b"hf" => {
                                f_hf = Some(Box::new(CTHeaderFooter::from_xml(reader, &e, false)?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            b"extLst" => {
                                f_ext_lst = Some(Box::new(CTExtensionListModify::from_xml(
                                    reader, &e, false,
                                )?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            #[cfg(feature = "extra-children")]
                            _ => {
                                // Capture unknown element for roundtrip
                                let elem = RawXmlElement::from_reader(reader, &e)?;
                                extra_children.push(PositionedNode::new(
                                    child_idx,
                                    RawXmlNode::Element(elem),
                                ));
                                child_idx += 1;
                            }
                            #[cfg(not(feature = "extra-children"))]
                            _ => {
                                // Skip unknown element
                                skip_element(reader)?;
                            }
                        }
                    }
                    Event::Empty(e) => {
                        match e.local_name().as_ref() {
                            b"cSld" => {
                                f_common_slide_data =
                                    Some(Box::new(CommonSlideData::from_xml(reader, &e, true)?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            b"clrMap" => {
                                f_clr_map = Some(Box::new(
                                    ooxml_dml::types::CTColorMapping::from_xml(reader, &e, true)?,
                                ));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            b"hf" => {
                                f_hf = Some(Box::new(CTHeaderFooter::from_xml(reader, &e, true)?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            b"extLst" => {
                                f_ext_lst = Some(Box::new(CTExtensionListModify::from_xml(
                                    reader, &e, true,
                                )?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            #[cfg(feature = "extra-children")]
                            _ => {
                                // Capture unknown empty element for roundtrip
                                let elem = RawXmlElement::from_empty(&e);
                                extra_children.push(PositionedNode::new(
                                    child_idx,
                                    RawXmlNode::Element(elem),
                                ));
                                child_idx += 1;
                            }
                            #[cfg(not(feature = "extra-children"))]
                            _ => {}
                        }
                    }
                    Event::End(_) => break,
                    Event::Eof => break,
                    _ => {}
                }
                buf.clear();
            }
        }

        Ok(Self {
            common_slide_data: f_common_slide_data
                .ok_or_else(|| ParseError::MissingAttribute("cSld".to_string()))?,
            clr_map: f_clr_map.ok_or_else(|| ParseError::MissingAttribute("clrMap".to_string()))?,
            hf: f_hf,
            ext_lst: f_ext_lst,
            #[cfg(feature = "extra-children")]
            extra_children,
        })
    }
}

impl FromXml for NotesMaster {
    fn from_xml<R: BufRead>(
        reader: &mut Reader<R>,
        start_tag: &BytesStart,
        is_empty: bool,
    ) -> Result<Self, ParseError> {
        let mut f_common_slide_data: Option<Box<CommonSlideData>> = None;
        let mut f_clr_map: Option<Box<ooxml_dml::types::CTColorMapping>> = None;
        let mut f_hf = None;
        let mut f_notes_style = None;
        let mut f_ext_lst = None;
        #[cfg(feature = "extra-children")]
        let mut extra_children = Vec::new();
        #[cfg(feature = "extra-children")]
        let mut child_idx: usize = 0;

        // Parse child elements
        if !is_empty {
            let mut buf = Vec::new();
            loop {
                match reader.read_event_into(&mut buf)? {
                    Event::Start(e) => {
                        match e.local_name().as_ref() {
                            b"cSld" => {
                                f_common_slide_data =
                                    Some(Box::new(CommonSlideData::from_xml(reader, &e, false)?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            b"clrMap" => {
                                f_clr_map = Some(Box::new(
                                    ooxml_dml::types::CTColorMapping::from_xml(reader, &e, false)?,
                                ));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            b"hf" => {
                                f_hf = Some(Box::new(CTHeaderFooter::from_xml(reader, &e, false)?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            b"notesStyle" => {
                                f_notes_style = Some(Box::new(
                                    ooxml_dml::types::CTTextListStyle::from_xml(reader, &e, false)?,
                                ));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            b"extLst" => {
                                f_ext_lst = Some(Box::new(CTExtensionListModify::from_xml(
                                    reader, &e, false,
                                )?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            #[cfg(feature = "extra-children")]
                            _ => {
                                // Capture unknown element for roundtrip
                                let elem = RawXmlElement::from_reader(reader, &e)?;
                                extra_children.push(PositionedNode::new(
                                    child_idx,
                                    RawXmlNode::Element(elem),
                                ));
                                child_idx += 1;
                            }
                            #[cfg(not(feature = "extra-children"))]
                            _ => {
                                // Skip unknown element
                                skip_element(reader)?;
                            }
                        }
                    }
                    Event::Empty(e) => {
                        match e.local_name().as_ref() {
                            b"cSld" => {
                                f_common_slide_data =
                                    Some(Box::new(CommonSlideData::from_xml(reader, &e, true)?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            b"clrMap" => {
                                f_clr_map = Some(Box::new(
                                    ooxml_dml::types::CTColorMapping::from_xml(reader, &e, true)?,
                                ));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            b"hf" => {
                                f_hf = Some(Box::new(CTHeaderFooter::from_xml(reader, &e, true)?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            b"notesStyle" => {
                                f_notes_style = Some(Box::new(
                                    ooxml_dml::types::CTTextListStyle::from_xml(reader, &e, true)?,
                                ));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            b"extLst" => {
                                f_ext_lst = Some(Box::new(CTExtensionListModify::from_xml(
                                    reader, &e, true,
                                )?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            #[cfg(feature = "extra-children")]
                            _ => {
                                // Capture unknown empty element for roundtrip
                                let elem = RawXmlElement::from_empty(&e);
                                extra_children.push(PositionedNode::new(
                                    child_idx,
                                    RawXmlNode::Element(elem),
                                ));
                                child_idx += 1;
                            }
                            #[cfg(not(feature = "extra-children"))]
                            _ => {}
                        }
                    }
                    Event::End(_) => break,
                    Event::Eof => break,
                    _ => {}
                }
                buf.clear();
            }
        }

        Ok(Self {
            common_slide_data: f_common_slide_data
                .ok_or_else(|| ParseError::MissingAttribute("cSld".to_string()))?,
            clr_map: f_clr_map.ok_or_else(|| ParseError::MissingAttribute("clrMap".to_string()))?,
            hf: f_hf,
            notes_style: f_notes_style,
            ext_lst: f_ext_lst,
            #[cfg(feature = "extra-children")]
            extra_children,
        })
    }
}

impl FromXml for NotesSlide {
    fn from_xml<R: BufRead>(
        reader: &mut Reader<R>,
        start_tag: &BytesStart,
        is_empty: bool,
    ) -> Result<Self, ParseError> {
        let mut f_show_master_sp = None;
        let mut f_show_master_ph_anim = None;
        let mut f_common_slide_data: Option<Box<CommonSlideData>> = None;
        let mut f_clr_map_ovr = None;
        let mut f_ext_lst = None;
        #[cfg(feature = "extra-attrs")]
        let mut extra_attrs = std::collections::HashMap::new();
        #[cfg(feature = "extra-children")]
        let mut extra_children = Vec::new();
        #[cfg(feature = "extra-children")]
        let mut child_idx: usize = 0;

        // Parse attributes
        for attr in start_tag.attributes().filter_map(|a| a.ok()) {
            let val = String::from_utf8_lossy(&attr.value);
            match attr.key.local_name().as_ref() {
                b"showMasterSp" => {
                    f_show_master_sp = Some(val == "true" || val == "1");
                }
                b"showMasterPhAnim" => {
                    f_show_master_ph_anim = Some(val == "true" || val == "1");
                }
                #[cfg(feature = "extra-attrs")]
                unknown => {
                    let key = String::from_utf8_lossy(attr.key.as_ref()).into_owned();
                    extra_attrs.insert(key, val.into_owned());
                }
                #[cfg(not(feature = "extra-attrs"))]
                _ => {}
            }
        }

        // Parse child elements
        if !is_empty {
            let mut buf = Vec::new();
            loop {
                match reader.read_event_into(&mut buf)? {
                    Event::Start(e) => {
                        match e.local_name().as_ref() {
                            b"cSld" => {
                                f_common_slide_data =
                                    Some(Box::new(CommonSlideData::from_xml(reader, &e, false)?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            b"clrMapOvr" => {
                                f_clr_map_ovr = Some(Box::new(
                                    ooxml_dml::types::CTColorMappingOverride::from_xml(
                                        reader, &e, false,
                                    )?,
                                ));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            b"extLst" => {
                                f_ext_lst = Some(Box::new(CTExtensionListModify::from_xml(
                                    reader, &e, false,
                                )?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            #[cfg(feature = "extra-children")]
                            _ => {
                                // Capture unknown element for roundtrip
                                let elem = RawXmlElement::from_reader(reader, &e)?;
                                extra_children.push(PositionedNode::new(
                                    child_idx,
                                    RawXmlNode::Element(elem),
                                ));
                                child_idx += 1;
                            }
                            #[cfg(not(feature = "extra-children"))]
                            _ => {
                                // Skip unknown element
                                skip_element(reader)?;
                            }
                        }
                    }
                    Event::Empty(e) => {
                        match e.local_name().as_ref() {
                            b"cSld" => {
                                f_common_slide_data =
                                    Some(Box::new(CommonSlideData::from_xml(reader, &e, true)?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            b"clrMapOvr" => {
                                f_clr_map_ovr = Some(Box::new(
                                    ooxml_dml::types::CTColorMappingOverride::from_xml(
                                        reader, &e, true,
                                    )?,
                                ));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            b"extLst" => {
                                f_ext_lst = Some(Box::new(CTExtensionListModify::from_xml(
                                    reader, &e, true,
                                )?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            #[cfg(feature = "extra-children")]
                            _ => {
                                // Capture unknown empty element for roundtrip
                                let elem = RawXmlElement::from_empty(&e);
                                extra_children.push(PositionedNode::new(
                                    child_idx,
                                    RawXmlNode::Element(elem),
                                ));
                                child_idx += 1;
                            }
                            #[cfg(not(feature = "extra-children"))]
                            _ => {}
                        }
                    }
                    Event::End(_) => break,
                    Event::Eof => break,
                    _ => {}
                }
                buf.clear();
            }
        }

        Ok(Self {
            show_master_sp: f_show_master_sp,
            show_master_ph_anim: f_show_master_ph_anim,
            common_slide_data: f_common_slide_data
                .ok_or_else(|| ParseError::MissingAttribute("cSld".to_string()))?,
            clr_map_ovr: f_clr_map_ovr,
            ext_lst: f_ext_lst,
            #[cfg(feature = "extra-attrs")]
            extra_attrs,
            #[cfg(feature = "extra-children")]
            extra_children,
        })
    }
}

impl FromXml for CTSlideSyncProperties {
    fn from_xml<R: BufRead>(
        reader: &mut Reader<R>,
        start_tag: &BytesStart,
        is_empty: bool,
    ) -> Result<Self, ParseError> {
        let mut f_server_sld_id: Option<String> = None;
        let mut f_server_sld_modified_time: Option<String> = None;
        let mut f_client_inserted_time: Option<String> = None;
        let mut f_ext_lst = None;
        #[cfg(feature = "extra-attrs")]
        let mut extra_attrs = std::collections::HashMap::new();
        #[cfg(feature = "extra-children")]
        let mut extra_children = Vec::new();
        #[cfg(feature = "extra-children")]
        let mut child_idx: usize = 0;

        // Parse attributes
        for attr in start_tag.attributes().filter_map(|a| a.ok()) {
            let val = String::from_utf8_lossy(&attr.value);
            match attr.key.local_name().as_ref() {
                b"serverSldId" => {
                    f_server_sld_id = Some(val.into_owned());
                }
                b"serverSldModifiedTime" => {
                    f_server_sld_modified_time = Some(val.into_owned());
                }
                b"clientInsertedTime" => {
                    f_client_inserted_time = Some(val.into_owned());
                }
                #[cfg(feature = "extra-attrs")]
                unknown => {
                    let key = String::from_utf8_lossy(attr.key.as_ref()).into_owned();
                    extra_attrs.insert(key, val.into_owned());
                }
                #[cfg(not(feature = "extra-attrs"))]
                _ => {}
            }
        }

        // Parse child elements
        if !is_empty {
            let mut buf = Vec::new();
            loop {
                match reader.read_event_into(&mut buf)? {
                    Event::Start(e) => {
                        match e.local_name().as_ref() {
                            b"extLst" => {
                                f_ext_lst =
                                    Some(Box::new(CTExtensionList::from_xml(reader, &e, false)?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            #[cfg(feature = "extra-children")]
                            _ => {
                                // Capture unknown element for roundtrip
                                let elem = RawXmlElement::from_reader(reader, &e)?;
                                extra_children.push(PositionedNode::new(
                                    child_idx,
                                    RawXmlNode::Element(elem),
                                ));
                                child_idx += 1;
                            }
                            #[cfg(not(feature = "extra-children"))]
                            _ => {
                                // Skip unknown element
                                skip_element(reader)?;
                            }
                        }
                    }
                    Event::Empty(e) => {
                        match e.local_name().as_ref() {
                            b"extLst" => {
                                f_ext_lst =
                                    Some(Box::new(CTExtensionList::from_xml(reader, &e, true)?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            #[cfg(feature = "extra-children")]
                            _ => {
                                // Capture unknown empty element for roundtrip
                                let elem = RawXmlElement::from_empty(&e);
                                extra_children.push(PositionedNode::new(
                                    child_idx,
                                    RawXmlNode::Element(elem),
                                ));
                                child_idx += 1;
                            }
                            #[cfg(not(feature = "extra-children"))]
                            _ => {}
                        }
                    }
                    Event::End(_) => break,
                    Event::Eof => break,
                    _ => {}
                }
                buf.clear();
            }
        }

        Ok(Self {
            server_sld_id: f_server_sld_id
                .ok_or_else(|| ParseError::MissingAttribute("serverSldId".to_string()))?,
            server_sld_modified_time: f_server_sld_modified_time
                .ok_or_else(|| ParseError::MissingAttribute("serverSldModifiedTime".to_string()))?,
            client_inserted_time: f_client_inserted_time
                .ok_or_else(|| ParseError::MissingAttribute("clientInsertedTime".to_string()))?,
            ext_lst: f_ext_lst,
            #[cfg(feature = "extra-attrs")]
            extra_attrs,
            #[cfg(feature = "extra-children")]
            extra_children,
        })
    }
}

impl FromXml for CTStringTag {
    fn from_xml<R: BufRead>(
        reader: &mut Reader<R>,
        start_tag: &BytesStart,
        is_empty: bool,
    ) -> Result<Self, ParseError> {
        let mut f_name: Option<String> = None;
        let mut f_value: Option<String> = None;
        #[cfg(feature = "extra-attrs")]
        let mut extra_attrs = std::collections::HashMap::new();

        // Parse attributes
        for attr in start_tag.attributes().filter_map(|a| a.ok()) {
            let val = String::from_utf8_lossy(&attr.value);
            match attr.key.local_name().as_ref() {
                b"name" => {
                    f_name = Some(val.into_owned());
                }
                b"val" => {
                    f_value = Some(val.into_owned());
                }
                #[cfg(feature = "extra-attrs")]
                unknown => {
                    let key = String::from_utf8_lossy(attr.key.as_ref()).into_owned();
                    extra_attrs.insert(key, val.into_owned());
                }
                #[cfg(not(feature = "extra-attrs"))]
                _ => {}
            }
        }

        if !is_empty {
            let mut buf = Vec::new();
            loop {
                match reader.read_event_into(&mut buf)? {
                    Event::End(_) => break,
                    Event::Eof => break,
                    _ => {}
                }
                buf.clear();
            }
        }

        Ok(Self {
            name: f_name.ok_or_else(|| ParseError::MissingAttribute("name".to_string()))?,
            value: f_value.ok_or_else(|| ParseError::MissingAttribute("val".to_string()))?,
            #[cfg(feature = "extra-attrs")]
            extra_attrs,
        })
    }
}

impl FromXml for CTTagList {
    fn from_xml<R: BufRead>(
        reader: &mut Reader<R>,
        start_tag: &BytesStart,
        is_empty: bool,
    ) -> Result<Self, ParseError> {
        let mut f_tag = Vec::new();
        #[cfg(feature = "extra-children")]
        let mut extra_children = Vec::new();
        #[cfg(feature = "extra-children")]
        let mut child_idx: usize = 0;

        // Parse child elements
        if !is_empty {
            let mut buf = Vec::new();
            loop {
                match reader.read_event_into(&mut buf)? {
                    Event::Start(e) => {
                        match e.local_name().as_ref() {
                            b"tag" => {
                                f_tag.push(CTStringTag::from_xml(reader, &e, false)?);
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            #[cfg(feature = "extra-children")]
                            _ => {
                                // Capture unknown element for roundtrip
                                let elem = RawXmlElement::from_reader(reader, &e)?;
                                extra_children.push(PositionedNode::new(
                                    child_idx,
                                    RawXmlNode::Element(elem),
                                ));
                                child_idx += 1;
                            }
                            #[cfg(not(feature = "extra-children"))]
                            _ => {
                                // Skip unknown element
                                skip_element(reader)?;
                            }
                        }
                    }
                    Event::Empty(e) => {
                        match e.local_name().as_ref() {
                            b"tag" => {
                                f_tag.push(CTStringTag::from_xml(reader, &e, true)?);
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            #[cfg(feature = "extra-children")]
                            _ => {
                                // Capture unknown empty element for roundtrip
                                let elem = RawXmlElement::from_empty(&e);
                                extra_children.push(PositionedNode::new(
                                    child_idx,
                                    RawXmlNode::Element(elem),
                                ));
                                child_idx += 1;
                            }
                            #[cfg(not(feature = "extra-children"))]
                            _ => {}
                        }
                    }
                    Event::End(_) => break,
                    Event::Eof => break,
                    _ => {}
                }
                buf.clear();
            }
        }

        Ok(Self {
            tag: f_tag,
            #[cfg(feature = "extra-children")]
            extra_children,
        })
    }
}

impl FromXml for CTNormalViewPortion {
    fn from_xml<R: BufRead>(
        reader: &mut Reader<R>,
        start_tag: &BytesStart,
        is_empty: bool,
    ) -> Result<Self, ParseError> {
        let mut f_sz: Option<ooxml_dml::types::STPositiveFixedPercentage> = None;
        let mut f_auto_adjust = None;
        #[cfg(feature = "extra-attrs")]
        let mut extra_attrs = std::collections::HashMap::new();

        // Parse attributes
        for attr in start_tag.attributes().filter_map(|a| a.ok()) {
            let val = String::from_utf8_lossy(&attr.value);
            match attr.key.local_name().as_ref() {
                b"sz" => {
                    f_sz = val.parse().ok();
                }
                b"autoAdjust" => {
                    f_auto_adjust = Some(val == "true" || val == "1");
                }
                #[cfg(feature = "extra-attrs")]
                unknown => {
                    let key = String::from_utf8_lossy(attr.key.as_ref()).into_owned();
                    extra_attrs.insert(key, val.into_owned());
                }
                #[cfg(not(feature = "extra-attrs"))]
                _ => {}
            }
        }

        if !is_empty {
            let mut buf = Vec::new();
            loop {
                match reader.read_event_into(&mut buf)? {
                    Event::End(_) => break,
                    Event::Eof => break,
                    _ => {}
                }
                buf.clear();
            }
        }

        Ok(Self {
            sz: f_sz.ok_or_else(|| ParseError::MissingAttribute("sz".to_string()))?,
            auto_adjust: f_auto_adjust,
            #[cfg(feature = "extra-attrs")]
            extra_attrs,
        })
    }
}

impl FromXml for CTNormalViewProperties {
    fn from_xml<R: BufRead>(
        reader: &mut Reader<R>,
        start_tag: &BytesStart,
        is_empty: bool,
    ) -> Result<Self, ParseError> {
        let mut f_show_outline_icons = None;
        let mut f_snap_vert_splitter = None;
        let mut f_vert_bar_state = None;
        let mut f_horz_bar_state = None;
        let mut f_prefer_single_view = None;
        let mut f_restored_left: Option<Box<CTNormalViewPortion>> = None;
        let mut f_restored_top: Option<Box<CTNormalViewPortion>> = None;
        let mut f_ext_lst = None;
        #[cfg(feature = "extra-attrs")]
        let mut extra_attrs = std::collections::HashMap::new();
        #[cfg(feature = "extra-children")]
        let mut extra_children = Vec::new();
        #[cfg(feature = "extra-children")]
        let mut child_idx: usize = 0;

        // Parse attributes
        for attr in start_tag.attributes().filter_map(|a| a.ok()) {
            let val = String::from_utf8_lossy(&attr.value);
            match attr.key.local_name().as_ref() {
                b"showOutlineIcons" => {
                    f_show_outline_icons = Some(val == "true" || val == "1");
                }
                b"snapVertSplitter" => {
                    f_snap_vert_splitter = Some(val == "true" || val == "1");
                }
                b"vertBarState" => {
                    f_vert_bar_state = val.parse().ok();
                }
                b"horzBarState" => {
                    f_horz_bar_state = val.parse().ok();
                }
                b"preferSingleView" => {
                    f_prefer_single_view = Some(val == "true" || val == "1");
                }
                #[cfg(feature = "extra-attrs")]
                unknown => {
                    let key = String::from_utf8_lossy(attr.key.as_ref()).into_owned();
                    extra_attrs.insert(key, val.into_owned());
                }
                #[cfg(not(feature = "extra-attrs"))]
                _ => {}
            }
        }

        // Parse child elements
        if !is_empty {
            let mut buf = Vec::new();
            loop {
                match reader.read_event_into(&mut buf)? {
                    Event::Start(e) => {
                        match e.local_name().as_ref() {
                            b"restoredLeft" => {
                                f_restored_left = Some(Box::new(CTNormalViewPortion::from_xml(
                                    reader, &e, false,
                                )?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            b"restoredTop" => {
                                f_restored_top = Some(Box::new(CTNormalViewPortion::from_xml(
                                    reader, &e, false,
                                )?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            b"extLst" => {
                                f_ext_lst =
                                    Some(Box::new(CTExtensionList::from_xml(reader, &e, false)?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            #[cfg(feature = "extra-children")]
                            _ => {
                                // Capture unknown element for roundtrip
                                let elem = RawXmlElement::from_reader(reader, &e)?;
                                extra_children.push(PositionedNode::new(
                                    child_idx,
                                    RawXmlNode::Element(elem),
                                ));
                                child_idx += 1;
                            }
                            #[cfg(not(feature = "extra-children"))]
                            _ => {
                                // Skip unknown element
                                skip_element(reader)?;
                            }
                        }
                    }
                    Event::Empty(e) => {
                        match e.local_name().as_ref() {
                            b"restoredLeft" => {
                                f_restored_left = Some(Box::new(CTNormalViewPortion::from_xml(
                                    reader, &e, true,
                                )?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            b"restoredTop" => {
                                f_restored_top = Some(Box::new(CTNormalViewPortion::from_xml(
                                    reader, &e, true,
                                )?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            b"extLst" => {
                                f_ext_lst =
                                    Some(Box::new(CTExtensionList::from_xml(reader, &e, true)?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            #[cfg(feature = "extra-children")]
                            _ => {
                                // Capture unknown empty element for roundtrip
                                let elem = RawXmlElement::from_empty(&e);
                                extra_children.push(PositionedNode::new(
                                    child_idx,
                                    RawXmlNode::Element(elem),
                                ));
                                child_idx += 1;
                            }
                            #[cfg(not(feature = "extra-children"))]
                            _ => {}
                        }
                    }
                    Event::End(_) => break,
                    Event::Eof => break,
                    _ => {}
                }
                buf.clear();
            }
        }

        Ok(Self {
            show_outline_icons: f_show_outline_icons,
            snap_vert_splitter: f_snap_vert_splitter,
            vert_bar_state: f_vert_bar_state,
            horz_bar_state: f_horz_bar_state,
            prefer_single_view: f_prefer_single_view,
            restored_left: f_restored_left
                .ok_or_else(|| ParseError::MissingAttribute("restoredLeft".to_string()))?,
            restored_top: f_restored_top
                .ok_or_else(|| ParseError::MissingAttribute("restoredTop".to_string()))?,
            ext_lst: f_ext_lst,
            #[cfg(feature = "extra-attrs")]
            extra_attrs,
            #[cfg(feature = "extra-children")]
            extra_children,
        })
    }
}

impl FromXml for CTCommonViewProperties {
    fn from_xml<R: BufRead>(
        reader: &mut Reader<R>,
        start_tag: &BytesStart,
        is_empty: bool,
    ) -> Result<Self, ParseError> {
        let mut f_var_scale = None;
        let mut f_scale: Option<Box<ooxml_dml::types::CTScale2D>> = None;
        let mut f_origin: Option<Box<ooxml_dml::types::Point2D>> = None;
        #[cfg(feature = "extra-attrs")]
        let mut extra_attrs = std::collections::HashMap::new();
        #[cfg(feature = "extra-children")]
        let mut extra_children = Vec::new();
        #[cfg(feature = "extra-children")]
        let mut child_idx: usize = 0;

        // Parse attributes
        for attr in start_tag.attributes().filter_map(|a| a.ok()) {
            let val = String::from_utf8_lossy(&attr.value);
            match attr.key.local_name().as_ref() {
                b"varScale" => {
                    f_var_scale = Some(val == "true" || val == "1");
                }
                #[cfg(feature = "extra-attrs")]
                unknown => {
                    let key = String::from_utf8_lossy(attr.key.as_ref()).into_owned();
                    extra_attrs.insert(key, val.into_owned());
                }
                #[cfg(not(feature = "extra-attrs"))]
                _ => {}
            }
        }

        // Parse child elements
        if !is_empty {
            let mut buf = Vec::new();
            loop {
                match reader.read_event_into(&mut buf)? {
                    Event::Start(e) => {
                        match e.local_name().as_ref() {
                            b"scale" => {
                                f_scale = Some(Box::new(ooxml_dml::types::CTScale2D::from_xml(
                                    reader, &e, false,
                                )?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            b"origin" => {
                                f_origin = Some(Box::new(ooxml_dml::types::Point2D::from_xml(
                                    reader, &e, false,
                                )?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            #[cfg(feature = "extra-children")]
                            _ => {
                                // Capture unknown element for roundtrip
                                let elem = RawXmlElement::from_reader(reader, &e)?;
                                extra_children.push(PositionedNode::new(
                                    child_idx,
                                    RawXmlNode::Element(elem),
                                ));
                                child_idx += 1;
                            }
                            #[cfg(not(feature = "extra-children"))]
                            _ => {
                                // Skip unknown element
                                skip_element(reader)?;
                            }
                        }
                    }
                    Event::Empty(e) => {
                        match e.local_name().as_ref() {
                            b"scale" => {
                                f_scale = Some(Box::new(ooxml_dml::types::CTScale2D::from_xml(
                                    reader, &e, true,
                                )?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            b"origin" => {
                                f_origin = Some(Box::new(ooxml_dml::types::Point2D::from_xml(
                                    reader, &e, true,
                                )?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            #[cfg(feature = "extra-children")]
                            _ => {
                                // Capture unknown empty element for roundtrip
                                let elem = RawXmlElement::from_empty(&e);
                                extra_children.push(PositionedNode::new(
                                    child_idx,
                                    RawXmlNode::Element(elem),
                                ));
                                child_idx += 1;
                            }
                            #[cfg(not(feature = "extra-children"))]
                            _ => {}
                        }
                    }
                    Event::End(_) => break,
                    Event::Eof => break,
                    _ => {}
                }
                buf.clear();
            }
        }

        Ok(Self {
            var_scale: f_var_scale,
            scale: f_scale.ok_or_else(|| ParseError::MissingAttribute("scale".to_string()))?,
            origin: f_origin.ok_or_else(|| ParseError::MissingAttribute("origin".to_string()))?,
            #[cfg(feature = "extra-attrs")]
            extra_attrs,
            #[cfg(feature = "extra-children")]
            extra_children,
        })
    }
}

impl FromXml for CTNotesTextViewProperties {
    fn from_xml<R: BufRead>(
        reader: &mut Reader<R>,
        start_tag: &BytesStart,
        is_empty: bool,
    ) -> Result<Self, ParseError> {
        let mut f_c_view_pr: Option<Box<CTCommonViewProperties>> = None;
        let mut f_ext_lst = None;
        #[cfg(feature = "extra-children")]
        let mut extra_children = Vec::new();
        #[cfg(feature = "extra-children")]
        let mut child_idx: usize = 0;

        // Parse child elements
        if !is_empty {
            let mut buf = Vec::new();
            loop {
                match reader.read_event_into(&mut buf)? {
                    Event::Start(e) => {
                        match e.local_name().as_ref() {
                            b"cViewPr" => {
                                f_c_view_pr = Some(Box::new(CTCommonViewProperties::from_xml(
                                    reader, &e, false,
                                )?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            b"extLst" => {
                                f_ext_lst =
                                    Some(Box::new(CTExtensionList::from_xml(reader, &e, false)?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            #[cfg(feature = "extra-children")]
                            _ => {
                                // Capture unknown element for roundtrip
                                let elem = RawXmlElement::from_reader(reader, &e)?;
                                extra_children.push(PositionedNode::new(
                                    child_idx,
                                    RawXmlNode::Element(elem),
                                ));
                                child_idx += 1;
                            }
                            #[cfg(not(feature = "extra-children"))]
                            _ => {
                                // Skip unknown element
                                skip_element(reader)?;
                            }
                        }
                    }
                    Event::Empty(e) => {
                        match e.local_name().as_ref() {
                            b"cViewPr" => {
                                f_c_view_pr = Some(Box::new(CTCommonViewProperties::from_xml(
                                    reader, &e, true,
                                )?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            b"extLst" => {
                                f_ext_lst =
                                    Some(Box::new(CTExtensionList::from_xml(reader, &e, true)?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            #[cfg(feature = "extra-children")]
                            _ => {
                                // Capture unknown empty element for roundtrip
                                let elem = RawXmlElement::from_empty(&e);
                                extra_children.push(PositionedNode::new(
                                    child_idx,
                                    RawXmlNode::Element(elem),
                                ));
                                child_idx += 1;
                            }
                            #[cfg(not(feature = "extra-children"))]
                            _ => {}
                        }
                    }
                    Event::End(_) => break,
                    Event::Eof => break,
                    _ => {}
                }
                buf.clear();
            }
        }

        Ok(Self {
            c_view_pr: f_c_view_pr
                .ok_or_else(|| ParseError::MissingAttribute("cViewPr".to_string()))?,
            ext_lst: f_ext_lst,
            #[cfg(feature = "extra-children")]
            extra_children,
        })
    }
}

impl FromXml for CTOutlineViewSlideEntry {
    fn from_xml<R: BufRead>(
        reader: &mut Reader<R>,
        start_tag: &BytesStart,
        is_empty: bool,
    ) -> Result<Self, ParseError> {
        let mut f_collapse = None;
        #[cfg(feature = "extra-attrs")]
        let mut extra_attrs = std::collections::HashMap::new();

        // Parse attributes
        for attr in start_tag.attributes().filter_map(|a| a.ok()) {
            let val = String::from_utf8_lossy(&attr.value);
            match attr.key.local_name().as_ref() {
                b"collapse" => {
                    f_collapse = Some(val == "true" || val == "1");
                }
                #[cfg(feature = "extra-attrs")]
                unknown => {
                    let key = String::from_utf8_lossy(attr.key.as_ref()).into_owned();
                    extra_attrs.insert(key, val.into_owned());
                }
                #[cfg(not(feature = "extra-attrs"))]
                _ => {}
            }
        }

        if !is_empty {
            let mut buf = Vec::new();
            loop {
                match reader.read_event_into(&mut buf)? {
                    Event::End(_) => break,
                    Event::Eof => break,
                    _ => {}
                }
                buf.clear();
            }
        }

        Ok(Self {
            collapse: f_collapse,
            #[cfg(feature = "extra-attrs")]
            extra_attrs,
        })
    }
}

impl FromXml for CTOutlineViewSlideList {
    fn from_xml<R: BufRead>(
        reader: &mut Reader<R>,
        start_tag: &BytesStart,
        is_empty: bool,
    ) -> Result<Self, ParseError> {
        let mut f_sld = Vec::new();
        #[cfg(feature = "extra-children")]
        let mut extra_children = Vec::new();
        #[cfg(feature = "extra-children")]
        let mut child_idx: usize = 0;

        // Parse child elements
        if !is_empty {
            let mut buf = Vec::new();
            loop {
                match reader.read_event_into(&mut buf)? {
                    Event::Start(e) => {
                        match e.local_name().as_ref() {
                            b"sld" => {
                                f_sld.push(CTOutlineViewSlideEntry::from_xml(reader, &e, false)?);
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            #[cfg(feature = "extra-children")]
                            _ => {
                                // Capture unknown element for roundtrip
                                let elem = RawXmlElement::from_reader(reader, &e)?;
                                extra_children.push(PositionedNode::new(
                                    child_idx,
                                    RawXmlNode::Element(elem),
                                ));
                                child_idx += 1;
                            }
                            #[cfg(not(feature = "extra-children"))]
                            _ => {
                                // Skip unknown element
                                skip_element(reader)?;
                            }
                        }
                    }
                    Event::Empty(e) => {
                        match e.local_name().as_ref() {
                            b"sld" => {
                                f_sld.push(CTOutlineViewSlideEntry::from_xml(reader, &e, true)?);
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            #[cfg(feature = "extra-children")]
                            _ => {
                                // Capture unknown empty element for roundtrip
                                let elem = RawXmlElement::from_empty(&e);
                                extra_children.push(PositionedNode::new(
                                    child_idx,
                                    RawXmlNode::Element(elem),
                                ));
                                child_idx += 1;
                            }
                            #[cfg(not(feature = "extra-children"))]
                            _ => {}
                        }
                    }
                    Event::End(_) => break,
                    Event::Eof => break,
                    _ => {}
                }
                buf.clear();
            }
        }

        Ok(Self {
            sld: f_sld,
            #[cfg(feature = "extra-children")]
            extra_children,
        })
    }
}

impl FromXml for CTOutlineViewProperties {
    fn from_xml<R: BufRead>(
        reader: &mut Reader<R>,
        start_tag: &BytesStart,
        is_empty: bool,
    ) -> Result<Self, ParseError> {
        let mut f_c_view_pr: Option<Box<CTCommonViewProperties>> = None;
        let mut f_sld_lst = None;
        let mut f_ext_lst = None;
        #[cfg(feature = "extra-children")]
        let mut extra_children = Vec::new();
        #[cfg(feature = "extra-children")]
        let mut child_idx: usize = 0;

        // Parse child elements
        if !is_empty {
            let mut buf = Vec::new();
            loop {
                match reader.read_event_into(&mut buf)? {
                    Event::Start(e) => {
                        match e.local_name().as_ref() {
                            b"cViewPr" => {
                                f_c_view_pr = Some(Box::new(CTCommonViewProperties::from_xml(
                                    reader, &e, false,
                                )?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            b"sldLst" => {
                                f_sld_lst = Some(Box::new(CTOutlineViewSlideList::from_xml(
                                    reader, &e, false,
                                )?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            b"extLst" => {
                                f_ext_lst =
                                    Some(Box::new(CTExtensionList::from_xml(reader, &e, false)?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            #[cfg(feature = "extra-children")]
                            _ => {
                                // Capture unknown element for roundtrip
                                let elem = RawXmlElement::from_reader(reader, &e)?;
                                extra_children.push(PositionedNode::new(
                                    child_idx,
                                    RawXmlNode::Element(elem),
                                ));
                                child_idx += 1;
                            }
                            #[cfg(not(feature = "extra-children"))]
                            _ => {
                                // Skip unknown element
                                skip_element(reader)?;
                            }
                        }
                    }
                    Event::Empty(e) => {
                        match e.local_name().as_ref() {
                            b"cViewPr" => {
                                f_c_view_pr = Some(Box::new(CTCommonViewProperties::from_xml(
                                    reader, &e, true,
                                )?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            b"sldLst" => {
                                f_sld_lst = Some(Box::new(CTOutlineViewSlideList::from_xml(
                                    reader, &e, true,
                                )?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            b"extLst" => {
                                f_ext_lst =
                                    Some(Box::new(CTExtensionList::from_xml(reader, &e, true)?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            #[cfg(feature = "extra-children")]
                            _ => {
                                // Capture unknown empty element for roundtrip
                                let elem = RawXmlElement::from_empty(&e);
                                extra_children.push(PositionedNode::new(
                                    child_idx,
                                    RawXmlNode::Element(elem),
                                ));
                                child_idx += 1;
                            }
                            #[cfg(not(feature = "extra-children"))]
                            _ => {}
                        }
                    }
                    Event::End(_) => break,
                    Event::Eof => break,
                    _ => {}
                }
                buf.clear();
            }
        }

        Ok(Self {
            c_view_pr: f_c_view_pr
                .ok_or_else(|| ParseError::MissingAttribute("cViewPr".to_string()))?,
            sld_lst: f_sld_lst,
            ext_lst: f_ext_lst,
            #[cfg(feature = "extra-children")]
            extra_children,
        })
    }
}

impl FromXml for CTSlideSorterViewProperties {
    fn from_xml<R: BufRead>(
        reader: &mut Reader<R>,
        start_tag: &BytesStart,
        is_empty: bool,
    ) -> Result<Self, ParseError> {
        let mut f_show_formatting = None;
        let mut f_c_view_pr: Option<Box<CTCommonViewProperties>> = None;
        let mut f_ext_lst = None;
        #[cfg(feature = "extra-attrs")]
        let mut extra_attrs = std::collections::HashMap::new();
        #[cfg(feature = "extra-children")]
        let mut extra_children = Vec::new();
        #[cfg(feature = "extra-children")]
        let mut child_idx: usize = 0;

        // Parse attributes
        for attr in start_tag.attributes().filter_map(|a| a.ok()) {
            let val = String::from_utf8_lossy(&attr.value);
            match attr.key.local_name().as_ref() {
                b"showFormatting" => {
                    f_show_formatting = Some(val == "true" || val == "1");
                }
                #[cfg(feature = "extra-attrs")]
                unknown => {
                    let key = String::from_utf8_lossy(attr.key.as_ref()).into_owned();
                    extra_attrs.insert(key, val.into_owned());
                }
                #[cfg(not(feature = "extra-attrs"))]
                _ => {}
            }
        }

        // Parse child elements
        if !is_empty {
            let mut buf = Vec::new();
            loop {
                match reader.read_event_into(&mut buf)? {
                    Event::Start(e) => {
                        match e.local_name().as_ref() {
                            b"cViewPr" => {
                                f_c_view_pr = Some(Box::new(CTCommonViewProperties::from_xml(
                                    reader, &e, false,
                                )?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            b"extLst" => {
                                f_ext_lst =
                                    Some(Box::new(CTExtensionList::from_xml(reader, &e, false)?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            #[cfg(feature = "extra-children")]
                            _ => {
                                // Capture unknown element for roundtrip
                                let elem = RawXmlElement::from_reader(reader, &e)?;
                                extra_children.push(PositionedNode::new(
                                    child_idx,
                                    RawXmlNode::Element(elem),
                                ));
                                child_idx += 1;
                            }
                            #[cfg(not(feature = "extra-children"))]
                            _ => {
                                // Skip unknown element
                                skip_element(reader)?;
                            }
                        }
                    }
                    Event::Empty(e) => {
                        match e.local_name().as_ref() {
                            b"cViewPr" => {
                                f_c_view_pr = Some(Box::new(CTCommonViewProperties::from_xml(
                                    reader, &e, true,
                                )?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            b"extLst" => {
                                f_ext_lst =
                                    Some(Box::new(CTExtensionList::from_xml(reader, &e, true)?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            #[cfg(feature = "extra-children")]
                            _ => {
                                // Capture unknown empty element for roundtrip
                                let elem = RawXmlElement::from_empty(&e);
                                extra_children.push(PositionedNode::new(
                                    child_idx,
                                    RawXmlNode::Element(elem),
                                ));
                                child_idx += 1;
                            }
                            #[cfg(not(feature = "extra-children"))]
                            _ => {}
                        }
                    }
                    Event::End(_) => break,
                    Event::Eof => break,
                    _ => {}
                }
                buf.clear();
            }
        }

        Ok(Self {
            show_formatting: f_show_formatting,
            c_view_pr: f_c_view_pr
                .ok_or_else(|| ParseError::MissingAttribute("cViewPr".to_string()))?,
            ext_lst: f_ext_lst,
            #[cfg(feature = "extra-attrs")]
            extra_attrs,
            #[cfg(feature = "extra-children")]
            extra_children,
        })
    }
}

impl FromXml for CTGuide {
    fn from_xml<R: BufRead>(
        reader: &mut Reader<R>,
        start_tag: &BytesStart,
        is_empty: bool,
    ) -> Result<Self, ParseError> {
        let mut f_orient = None;
        let mut f_pos = None;
        #[cfg(feature = "extra-attrs")]
        let mut extra_attrs = std::collections::HashMap::new();

        // Parse attributes
        for attr in start_tag.attributes().filter_map(|a| a.ok()) {
            let val = String::from_utf8_lossy(&attr.value);
            match attr.key.local_name().as_ref() {
                b"orient" => {
                    f_orient = val.parse().ok();
                }
                b"pos" => {
                    f_pos = val.parse().ok();
                }
                #[cfg(feature = "extra-attrs")]
                unknown => {
                    let key = String::from_utf8_lossy(attr.key.as_ref()).into_owned();
                    extra_attrs.insert(key, val.into_owned());
                }
                #[cfg(not(feature = "extra-attrs"))]
                _ => {}
            }
        }

        if !is_empty {
            let mut buf = Vec::new();
            loop {
                match reader.read_event_into(&mut buf)? {
                    Event::End(_) => break,
                    Event::Eof => break,
                    _ => {}
                }
                buf.clear();
            }
        }

        Ok(Self {
            orient: f_orient,
            pos: f_pos,
            #[cfg(feature = "extra-attrs")]
            extra_attrs,
        })
    }
}

impl FromXml for CTGuideList {
    fn from_xml<R: BufRead>(
        reader: &mut Reader<R>,
        start_tag: &BytesStart,
        is_empty: bool,
    ) -> Result<Self, ParseError> {
        let mut f_guide = Vec::new();
        #[cfg(feature = "extra-children")]
        let mut extra_children = Vec::new();
        #[cfg(feature = "extra-children")]
        let mut child_idx: usize = 0;

        // Parse child elements
        if !is_empty {
            let mut buf = Vec::new();
            loop {
                match reader.read_event_into(&mut buf)? {
                    Event::Start(e) => {
                        match e.local_name().as_ref() {
                            b"guide" => {
                                f_guide.push(CTGuide::from_xml(reader, &e, false)?);
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            #[cfg(feature = "extra-children")]
                            _ => {
                                // Capture unknown element for roundtrip
                                let elem = RawXmlElement::from_reader(reader, &e)?;
                                extra_children.push(PositionedNode::new(
                                    child_idx,
                                    RawXmlNode::Element(elem),
                                ));
                                child_idx += 1;
                            }
                            #[cfg(not(feature = "extra-children"))]
                            _ => {
                                // Skip unknown element
                                skip_element(reader)?;
                            }
                        }
                    }
                    Event::Empty(e) => {
                        match e.local_name().as_ref() {
                            b"guide" => {
                                f_guide.push(CTGuide::from_xml(reader, &e, true)?);
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            #[cfg(feature = "extra-children")]
                            _ => {
                                // Capture unknown empty element for roundtrip
                                let elem = RawXmlElement::from_empty(&e);
                                extra_children.push(PositionedNode::new(
                                    child_idx,
                                    RawXmlNode::Element(elem),
                                ));
                                child_idx += 1;
                            }
                            #[cfg(not(feature = "extra-children"))]
                            _ => {}
                        }
                    }
                    Event::End(_) => break,
                    Event::Eof => break,
                    _ => {}
                }
                buf.clear();
            }
        }

        Ok(Self {
            guide: f_guide,
            #[cfg(feature = "extra-children")]
            extra_children,
        })
    }
}

impl FromXml for CTCommonSlideViewProperties {
    fn from_xml<R: BufRead>(
        reader: &mut Reader<R>,
        start_tag: &BytesStart,
        is_empty: bool,
    ) -> Result<Self, ParseError> {
        let mut f_snap_to_grid = None;
        let mut f_snap_to_objects = None;
        let mut f_show_guides = None;
        let mut f_c_view_pr: Option<Box<CTCommonViewProperties>> = None;
        let mut f_guide_lst = None;
        #[cfg(feature = "extra-attrs")]
        let mut extra_attrs = std::collections::HashMap::new();
        #[cfg(feature = "extra-children")]
        let mut extra_children = Vec::new();
        #[cfg(feature = "extra-children")]
        let mut child_idx: usize = 0;

        // Parse attributes
        for attr in start_tag.attributes().filter_map(|a| a.ok()) {
            let val = String::from_utf8_lossy(&attr.value);
            match attr.key.local_name().as_ref() {
                b"snapToGrid" => {
                    f_snap_to_grid = Some(val == "true" || val == "1");
                }
                b"snapToObjects" => {
                    f_snap_to_objects = Some(val == "true" || val == "1");
                }
                b"showGuides" => {
                    f_show_guides = Some(val == "true" || val == "1");
                }
                #[cfg(feature = "extra-attrs")]
                unknown => {
                    let key = String::from_utf8_lossy(attr.key.as_ref()).into_owned();
                    extra_attrs.insert(key, val.into_owned());
                }
                #[cfg(not(feature = "extra-attrs"))]
                _ => {}
            }
        }

        // Parse child elements
        if !is_empty {
            let mut buf = Vec::new();
            loop {
                match reader.read_event_into(&mut buf)? {
                    Event::Start(e) => {
                        match e.local_name().as_ref() {
                            b"cViewPr" => {
                                f_c_view_pr = Some(Box::new(CTCommonViewProperties::from_xml(
                                    reader, &e, false,
                                )?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            b"guideLst" => {
                                f_guide_lst =
                                    Some(Box::new(CTGuideList::from_xml(reader, &e, false)?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            #[cfg(feature = "extra-children")]
                            _ => {
                                // Capture unknown element for roundtrip
                                let elem = RawXmlElement::from_reader(reader, &e)?;
                                extra_children.push(PositionedNode::new(
                                    child_idx,
                                    RawXmlNode::Element(elem),
                                ));
                                child_idx += 1;
                            }
                            #[cfg(not(feature = "extra-children"))]
                            _ => {
                                // Skip unknown element
                                skip_element(reader)?;
                            }
                        }
                    }
                    Event::Empty(e) => {
                        match e.local_name().as_ref() {
                            b"cViewPr" => {
                                f_c_view_pr = Some(Box::new(CTCommonViewProperties::from_xml(
                                    reader, &e, true,
                                )?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            b"guideLst" => {
                                f_guide_lst =
                                    Some(Box::new(CTGuideList::from_xml(reader, &e, true)?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            #[cfg(feature = "extra-children")]
                            _ => {
                                // Capture unknown empty element for roundtrip
                                let elem = RawXmlElement::from_empty(&e);
                                extra_children.push(PositionedNode::new(
                                    child_idx,
                                    RawXmlNode::Element(elem),
                                ));
                                child_idx += 1;
                            }
                            #[cfg(not(feature = "extra-children"))]
                            _ => {}
                        }
                    }
                    Event::End(_) => break,
                    Event::Eof => break,
                    _ => {}
                }
                buf.clear();
            }
        }

        Ok(Self {
            snap_to_grid: f_snap_to_grid,
            snap_to_objects: f_snap_to_objects,
            show_guides: f_show_guides,
            c_view_pr: f_c_view_pr
                .ok_or_else(|| ParseError::MissingAttribute("cViewPr".to_string()))?,
            guide_lst: f_guide_lst,
            #[cfg(feature = "extra-attrs")]
            extra_attrs,
            #[cfg(feature = "extra-children")]
            extra_children,
        })
    }
}

impl FromXml for CTSlideViewProperties {
    fn from_xml<R: BufRead>(
        reader: &mut Reader<R>,
        start_tag: &BytesStart,
        is_empty: bool,
    ) -> Result<Self, ParseError> {
        let mut f_c_sld_view_pr: Option<Box<CTCommonSlideViewProperties>> = None;
        let mut f_ext_lst = None;
        #[cfg(feature = "extra-children")]
        let mut extra_children = Vec::new();
        #[cfg(feature = "extra-children")]
        let mut child_idx: usize = 0;

        // Parse child elements
        if !is_empty {
            let mut buf = Vec::new();
            loop {
                match reader.read_event_into(&mut buf)? {
                    Event::Start(e) => {
                        match e.local_name().as_ref() {
                            b"cSldViewPr" => {
                                f_c_sld_view_pr = Some(Box::new(
                                    CTCommonSlideViewProperties::from_xml(reader, &e, false)?,
                                ));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            b"extLst" => {
                                f_ext_lst =
                                    Some(Box::new(CTExtensionList::from_xml(reader, &e, false)?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            #[cfg(feature = "extra-children")]
                            _ => {
                                // Capture unknown element for roundtrip
                                let elem = RawXmlElement::from_reader(reader, &e)?;
                                extra_children.push(PositionedNode::new(
                                    child_idx,
                                    RawXmlNode::Element(elem),
                                ));
                                child_idx += 1;
                            }
                            #[cfg(not(feature = "extra-children"))]
                            _ => {
                                // Skip unknown element
                                skip_element(reader)?;
                            }
                        }
                    }
                    Event::Empty(e) => {
                        match e.local_name().as_ref() {
                            b"cSldViewPr" => {
                                f_c_sld_view_pr = Some(Box::new(
                                    CTCommonSlideViewProperties::from_xml(reader, &e, true)?,
                                ));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            b"extLst" => {
                                f_ext_lst =
                                    Some(Box::new(CTExtensionList::from_xml(reader, &e, true)?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            #[cfg(feature = "extra-children")]
                            _ => {
                                // Capture unknown empty element for roundtrip
                                let elem = RawXmlElement::from_empty(&e);
                                extra_children.push(PositionedNode::new(
                                    child_idx,
                                    RawXmlNode::Element(elem),
                                ));
                                child_idx += 1;
                            }
                            #[cfg(not(feature = "extra-children"))]
                            _ => {}
                        }
                    }
                    Event::End(_) => break,
                    Event::Eof => break,
                    _ => {}
                }
                buf.clear();
            }
        }

        Ok(Self {
            c_sld_view_pr: f_c_sld_view_pr
                .ok_or_else(|| ParseError::MissingAttribute("cSldViewPr".to_string()))?,
            ext_lst: f_ext_lst,
            #[cfg(feature = "extra-children")]
            extra_children,
        })
    }
}

impl FromXml for CTNotesViewProperties {
    fn from_xml<R: BufRead>(
        reader: &mut Reader<R>,
        start_tag: &BytesStart,
        is_empty: bool,
    ) -> Result<Self, ParseError> {
        let mut f_c_sld_view_pr: Option<Box<CTCommonSlideViewProperties>> = None;
        let mut f_ext_lst = None;
        #[cfg(feature = "extra-children")]
        let mut extra_children = Vec::new();
        #[cfg(feature = "extra-children")]
        let mut child_idx: usize = 0;

        // Parse child elements
        if !is_empty {
            let mut buf = Vec::new();
            loop {
                match reader.read_event_into(&mut buf)? {
                    Event::Start(e) => {
                        match e.local_name().as_ref() {
                            b"cSldViewPr" => {
                                f_c_sld_view_pr = Some(Box::new(
                                    CTCommonSlideViewProperties::from_xml(reader, &e, false)?,
                                ));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            b"extLst" => {
                                f_ext_lst =
                                    Some(Box::new(CTExtensionList::from_xml(reader, &e, false)?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            #[cfg(feature = "extra-children")]
                            _ => {
                                // Capture unknown element for roundtrip
                                let elem = RawXmlElement::from_reader(reader, &e)?;
                                extra_children.push(PositionedNode::new(
                                    child_idx,
                                    RawXmlNode::Element(elem),
                                ));
                                child_idx += 1;
                            }
                            #[cfg(not(feature = "extra-children"))]
                            _ => {
                                // Skip unknown element
                                skip_element(reader)?;
                            }
                        }
                    }
                    Event::Empty(e) => {
                        match e.local_name().as_ref() {
                            b"cSldViewPr" => {
                                f_c_sld_view_pr = Some(Box::new(
                                    CTCommonSlideViewProperties::from_xml(reader, &e, true)?,
                                ));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            b"extLst" => {
                                f_ext_lst =
                                    Some(Box::new(CTExtensionList::from_xml(reader, &e, true)?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            #[cfg(feature = "extra-children")]
                            _ => {
                                // Capture unknown empty element for roundtrip
                                let elem = RawXmlElement::from_empty(&e);
                                extra_children.push(PositionedNode::new(
                                    child_idx,
                                    RawXmlNode::Element(elem),
                                ));
                                child_idx += 1;
                            }
                            #[cfg(not(feature = "extra-children"))]
                            _ => {}
                        }
                    }
                    Event::End(_) => break,
                    Event::Eof => break,
                    _ => {}
                }
                buf.clear();
            }
        }

        Ok(Self {
            c_sld_view_pr: f_c_sld_view_pr
                .ok_or_else(|| ParseError::MissingAttribute("cSldViewPr".to_string()))?,
            ext_lst: f_ext_lst,
            #[cfg(feature = "extra-children")]
            extra_children,
        })
    }
}

impl FromXml for CTViewProperties {
    fn from_xml<R: BufRead>(
        reader: &mut Reader<R>,
        start_tag: &BytesStart,
        is_empty: bool,
    ) -> Result<Self, ParseError> {
        let mut f_last_view = None;
        let mut f_show_comments = None;
        let mut f_normal_view_pr = None;
        let mut f_slide_view_pr = None;
        let mut f_outline_view_pr = None;
        let mut f_notes_text_view_pr = None;
        let mut f_sorter_view_pr = None;
        let mut f_notes_view_pr = None;
        let mut f_grid_spacing = None;
        let mut f_ext_lst = None;
        #[cfg(feature = "extra-attrs")]
        let mut extra_attrs = std::collections::HashMap::new();
        #[cfg(feature = "extra-children")]
        let mut extra_children = Vec::new();
        #[cfg(feature = "extra-children")]
        let mut child_idx: usize = 0;

        // Parse attributes
        for attr in start_tag.attributes().filter_map(|a| a.ok()) {
            let val = String::from_utf8_lossy(&attr.value);
            match attr.key.local_name().as_ref() {
                b"lastView" => {
                    f_last_view = val.parse().ok();
                }
                b"showComments" => {
                    f_show_comments = Some(val == "true" || val == "1");
                }
                #[cfg(feature = "extra-attrs")]
                unknown => {
                    let key = String::from_utf8_lossy(attr.key.as_ref()).into_owned();
                    extra_attrs.insert(key, val.into_owned());
                }
                #[cfg(not(feature = "extra-attrs"))]
                _ => {}
            }
        }

        // Parse child elements
        if !is_empty {
            let mut buf = Vec::new();
            loop {
                match reader.read_event_into(&mut buf)? {
                    Event::Start(e) => {
                        match e.local_name().as_ref() {
                            b"normalViewPr" => {
                                f_normal_view_pr = Some(Box::new(
                                    CTNormalViewProperties::from_xml(reader, &e, false)?,
                                ));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            b"slideViewPr" => {
                                f_slide_view_pr = Some(Box::new(CTSlideViewProperties::from_xml(
                                    reader, &e, false,
                                )?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            b"outlineViewPr" => {
                                f_outline_view_pr = Some(Box::new(
                                    CTOutlineViewProperties::from_xml(reader, &e, false)?,
                                ));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            b"notesTextViewPr" => {
                                f_notes_text_view_pr = Some(Box::new(
                                    CTNotesTextViewProperties::from_xml(reader, &e, false)?,
                                ));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            b"sorterViewPr" => {
                                f_sorter_view_pr = Some(Box::new(
                                    CTSlideSorterViewProperties::from_xml(reader, &e, false)?,
                                ));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            b"notesViewPr" => {
                                f_notes_view_pr = Some(Box::new(CTNotesViewProperties::from_xml(
                                    reader, &e, false,
                                )?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            b"gridSpacing" => {
                                f_grid_spacing = Some(Box::new(
                                    ooxml_dml::types::PositiveSize2D::from_xml(reader, &e, false)?,
                                ));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            b"extLst" => {
                                f_ext_lst =
                                    Some(Box::new(CTExtensionList::from_xml(reader, &e, false)?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            #[cfg(feature = "extra-children")]
                            _ => {
                                // Capture unknown element for roundtrip
                                let elem = RawXmlElement::from_reader(reader, &e)?;
                                extra_children.push(PositionedNode::new(
                                    child_idx,
                                    RawXmlNode::Element(elem),
                                ));
                                child_idx += 1;
                            }
                            #[cfg(not(feature = "extra-children"))]
                            _ => {
                                // Skip unknown element
                                skip_element(reader)?;
                            }
                        }
                    }
                    Event::Empty(e) => {
                        match e.local_name().as_ref() {
                            b"normalViewPr" => {
                                f_normal_view_pr = Some(Box::new(
                                    CTNormalViewProperties::from_xml(reader, &e, true)?,
                                ));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            b"slideViewPr" => {
                                f_slide_view_pr = Some(Box::new(CTSlideViewProperties::from_xml(
                                    reader, &e, true,
                                )?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            b"outlineViewPr" => {
                                f_outline_view_pr = Some(Box::new(
                                    CTOutlineViewProperties::from_xml(reader, &e, true)?,
                                ));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            b"notesTextViewPr" => {
                                f_notes_text_view_pr = Some(Box::new(
                                    CTNotesTextViewProperties::from_xml(reader, &e, true)?,
                                ));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            b"sorterViewPr" => {
                                f_sorter_view_pr = Some(Box::new(
                                    CTSlideSorterViewProperties::from_xml(reader, &e, true)?,
                                ));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            b"notesViewPr" => {
                                f_notes_view_pr = Some(Box::new(CTNotesViewProperties::from_xml(
                                    reader, &e, true,
                                )?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            b"gridSpacing" => {
                                f_grid_spacing = Some(Box::new(
                                    ooxml_dml::types::PositiveSize2D::from_xml(reader, &e, true)?,
                                ));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            b"extLst" => {
                                f_ext_lst =
                                    Some(Box::new(CTExtensionList::from_xml(reader, &e, true)?));
                                #[cfg(feature = "extra-children")]
                                {
                                    child_idx += 1;
                                }
                            }
                            #[cfg(feature = "extra-children")]
                            _ => {
                                // Capture unknown empty element for roundtrip
                                let elem = RawXmlElement::from_empty(&e);
                                extra_children.push(PositionedNode::new(
                                    child_idx,
                                    RawXmlNode::Element(elem),
                                ));
                                child_idx += 1;
                            }
                            #[cfg(not(feature = "extra-children"))]
                            _ => {}
                        }
                    }
                    Event::End(_) => break,
                    Event::Eof => break,
                    _ => {}
                }
                buf.clear();
            }
        }

        Ok(Self {
            last_view: f_last_view,
            show_comments: f_show_comments,
            normal_view_pr: f_normal_view_pr,
            slide_view_pr: f_slide_view_pr,
            outline_view_pr: f_outline_view_pr,
            notes_text_view_pr: f_notes_text_view_pr,
            sorter_view_pr: f_sorter_view_pr,
            notes_view_pr: f_notes_view_pr,
            grid_spacing: f_grid_spacing,
            ext_lst: f_ext_lst,
            #[cfg(feature = "extra-attrs")]
            extra_attrs,
            #[cfg(feature = "extra-children")]
            extra_children,
        })
    }
}
