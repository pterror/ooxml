// ToXml serializers for generated types.
// Enables roundtrip XML serialization alongside FromXml parsers.

#![allow(unused_variables, unused_assignments, unreachable_code)]
#![allow(clippy::single_match)]
#![allow(clippy::match_single_binding)]
#![allow(clippy::explicit_counter_loop)]

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

impl ToXml for AutoFilter {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        #[cfg(feature = "sml-filtering")]
        if let Some(ref val) = self.reference {
            start.push_attribute(("ref", val.as_str()));
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
        #[cfg(feature = "sml-filtering")]
        for item in &self.filter_column {
            #[cfg(feature = "extra-children")]
            while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
                extra_iter
                    .next()
                    .unwrap()
                    .node
                    .write_to(writer)
                    .map_err(SerializeError::from)?;
            }
            item.write_element("filterColumn", writer)?;
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
        #[cfg(feature = "sml-filtering")]
        if let Some(ref val) = self.sort_state {
            val.write_element("sortState", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        #[cfg(feature = "sml-extensions")]
        if let Some(ref val) = self.extension_list {
            val.write_element("extLst", writer)?;
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
        #[cfg(feature = "sml-filtering")]
        if !self.filter_column.is_empty() {
            return false;
        }
        #[cfg(feature = "sml-filtering")]
        if self.sort_state.is_some() {
            return false;
        }
        #[cfg(feature = "sml-extensions")]
        if self.extension_list.is_some() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for FilterColumn {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        #[cfg(feature = "sml-filtering")]
        {
            let val = &self.column_id;
            {
                let s = val.to_string();
                start.push_attribute(("colId", s.as_str()));
            }
        }
        #[cfg(feature = "sml-filtering")]
        if let Some(ref val) = self.hidden_button {
            start.push_attribute(("hiddenButton", if *val { "1" } else { "0" }));
        }
        #[cfg(feature = "sml-filtering")]
        if let Some(ref val) = self.show_button {
            start.push_attribute(("showButton", if *val { "1" } else { "0" }));
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

impl ToXml for Filters {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        if let Some(ref val) = self.blank {
            start.push_attribute(("blank", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.calendar_type {
            {
                let s = val.to_string();
                start.push_attribute(("calendarType", s.as_str()));
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
        for item in &self.filter {
            #[cfg(feature = "extra-children")]
            while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
                extra_iter
                    .next()
                    .unwrap()
                    .node
                    .write_to(writer)
                    .map_err(SerializeError::from)?;
            }
            item.write_element("filter", writer)?;
            #[cfg(feature = "extra-children")]
            {
                emit_idx += 1;
            }
        }
        for item in &self.date_group_item {
            #[cfg(feature = "extra-children")]
            while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
                extra_iter
                    .next()
                    .unwrap()
                    .node
                    .write_to(writer)
                    .map_err(SerializeError::from)?;
            }
            item.write_element("dateGroupItem", writer)?;
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
        if !self.filter.is_empty() {
            return false;
        }
        if !self.date_group_item.is_empty() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for Filter {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        if let Some(ref val) = self.value {
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

impl ToXml for CustomFilters {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        if let Some(ref val) = self.and {
            start.push_attribute(("and", if *val { "1" } else { "0" }));
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
        for item in &self.custom_filter {
            #[cfg(feature = "extra-children")]
            while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
                extra_iter
                    .next()
                    .unwrap()
                    .node
                    .write_to(writer)
                    .map_err(SerializeError::from)?;
            }
            item.write_element("customFilter", writer)?;
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
        if !self.custom_filter.is_empty() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for CustomFilter {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        if let Some(ref val) = self.operator {
            {
                let s = val.to_string();
                start.push_attribute(("operator", s.as_str()));
            }
        }
        if let Some(ref val) = self.value {
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

impl ToXml for Top10Filter {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        if let Some(ref val) = self.top {
            start.push_attribute(("top", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.percent {
            start.push_attribute(("percent", if *val { "1" } else { "0" }));
        }
        {
            let val = &self.value;
            {
                let s = val.to_string();
                start.push_attribute(("val", s.as_str()));
            }
        }
        if let Some(ref val) = self.filter_val {
            {
                let s = val.to_string();
                start.push_attribute(("filterVal", s.as_str()));
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

impl ToXml for ColorFilter {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        if let Some(ref val) = self.dxf_id {
            {
                let s = val.to_string();
                start.push_attribute(("dxfId", s.as_str()));
            }
        }
        if let Some(ref val) = self.cell_color {
            start.push_attribute(("cellColor", if *val { "1" } else { "0" }));
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

impl ToXml for IconFilter {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        {
            let val = &self.icon_set;
            {
                let s = val.to_string();
                start.push_attribute(("iconSet", s.as_str()));
            }
        }
        if let Some(ref val) = self.icon_id {
            {
                let s = val.to_string();
                start.push_attribute(("iconId", s.as_str()));
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

impl ToXml for DynamicFilter {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        {
            let val = &self.r#type;
            {
                let s = val.to_string();
                start.push_attribute(("type", s.as_str()));
            }
        }
        if let Some(ref val) = self.value {
            {
                let s = val.to_string();
                start.push_attribute(("val", s.as_str()));
            }
        }
        if let Some(ref val) = self.val_iso {
            start.push_attribute(("valIso", val.as_str()));
        }
        if let Some(ref val) = self.max_val {
            {
                let s = val.to_string();
                start.push_attribute(("maxVal", s.as_str()));
            }
        }
        if let Some(ref val) = self.max_val_iso {
            start.push_attribute(("maxValIso", val.as_str()));
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

impl ToXml for SortState {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        if let Some(ref val) = self.column_sort {
            start.push_attribute(("columnSort", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.case_sensitive {
            start.push_attribute(("caseSensitive", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.sort_method {
            {
                let s = val.to_string();
                start.push_attribute(("sortMethod", s.as_str()));
            }
        }
        {
            let val = &self.reference;
            start.push_attribute(("ref", val.as_str()));
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
        for item in &self.sort_condition {
            #[cfg(feature = "extra-children")]
            while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
                extra_iter
                    .next()
                    .unwrap()
                    .node
                    .write_to(writer)
                    .map_err(SerializeError::from)?;
            }
            item.write_element("sortCondition", writer)?;
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
        if let Some(ref val) = self.extension_list {
            val.write_element("extLst", writer)?;
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
        if !self.sort_condition.is_empty() {
            return false;
        }
        if self.extension_list.is_some() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for SortCondition {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        if let Some(ref val) = self.descending {
            start.push_attribute(("descending", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.sort_by {
            {
                let s = val.to_string();
                start.push_attribute(("sortBy", s.as_str()));
            }
        }
        {
            let val = &self.reference;
            start.push_attribute(("ref", val.as_str()));
        }
        if let Some(ref val) = self.custom_list {
            start.push_attribute(("customList", val.as_str()));
        }
        if let Some(ref val) = self.dxf_id {
            {
                let s = val.to_string();
                start.push_attribute(("dxfId", s.as_str()));
            }
        }
        if let Some(ref val) = self.icon_set {
            {
                let s = val.to_string();
                start.push_attribute(("iconSet", s.as_str()));
            }
        }
        if let Some(ref val) = self.icon_id {
            {
                let s = val.to_string();
                start.push_attribute(("iconId", s.as_str()));
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

impl ToXml for DateGroupItem {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        {
            let val = &self.year;
            {
                let s = val.to_string();
                start.push_attribute(("year", s.as_str()));
            }
        }
        if let Some(ref val) = self.month {
            {
                let s = val.to_string();
                start.push_attribute(("month", s.as_str()));
            }
        }
        if let Some(ref val) = self.day {
            {
                let s = val.to_string();
                start.push_attribute(("day", s.as_str()));
            }
        }
        if let Some(ref val) = self.hour {
            {
                let s = val.to_string();
                start.push_attribute(("hour", s.as_str()));
            }
        }
        if let Some(ref val) = self.minute {
            {
                let s = val.to_string();
                start.push_attribute(("minute", s.as_str()));
            }
        }
        if let Some(ref val) = self.second {
            {
                let s = val.to_string();
                start.push_attribute(("second", s.as_str()));
            }
        }
        {
            let val = &self.date_time_grouping;
            {
                let s = val.to_string();
                start.push_attribute(("dateTimeGrouping", s.as_str()));
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

impl ToXml for CTXStringElement {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        {
            let val = &self.value;
            start.push_attribute(("v", val.as_str()));
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

impl ToXml for Extension {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        if let Some(ref val) = self.uri {
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

impl ToXml for ObjectAnchor {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        if let Some(ref val) = self.move_with_cells {
            start.push_attribute(("moveWithCells", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.size_with_cells {
            start.push_attribute(("sizeWithCells", if *val { "1" } else { "0" }));
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
            item.write_element("ext", writer)?;
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

impl ToXml for ExtensionList {
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
            item.write_element("ext", writer)?;
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

impl ToXml for CalcChain {
    fn write_children<W: Write>(&self, writer: &mut Writer<W>) -> Result<(), SerializeError> {
        #[cfg(feature = "extra-children")]
        let mut extra_iter = self.extra_children.iter().peekable();
        #[cfg(feature = "extra-children")]
        let mut emit_idx: usize = 0;
        for item in &self.cells {
            #[cfg(feature = "extra-children")]
            while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
                extra_iter
                    .next()
                    .unwrap()
                    .node
                    .write_to(writer)
                    .map_err(SerializeError::from)?;
            }
            item.write_element("c", writer)?;
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
        if let Some(ref val) = self.extension_list {
            val.write_element("extLst", writer)?;
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
        if !self.cells.is_empty() {
            return false;
        }
        if self.extension_list.is_some() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for CalcCell {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        {
            let val = &self._any;
            start.push_attribute(("_any", val.as_str()));
        }
        if let Some(ref val) = self.i {
            {
                let s = val.to_string();
                start.push_attribute(("i", s.as_str()));
            }
        }
        if let Some(ref val) = self.style_index {
            start.push_attribute(("s", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.l {
            start.push_attribute(("l", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.cell_type {
            start.push_attribute(("t", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.a {
            start.push_attribute(("a", if *val { "1" } else { "0" }));
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

impl ToXml for Comments {
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
            let val = &self.authors;
            val.write_element("authors", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
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
            let val = &self.comment_list;
            val.write_element("commentList", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        if let Some(ref val) = self.extension_list {
            val.write_element("extLst", writer)?;
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

impl ToXml for Authors {
    fn write_children<W: Write>(&self, writer: &mut Writer<W>) -> Result<(), SerializeError> {
        #[cfg(feature = "extra-children")]
        let mut extra_iter = self.extra_children.iter().peekable();
        #[cfg(feature = "extra-children")]
        let mut emit_idx: usize = 0;
        for item in &self.author {
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
                let start = BytesStart::new("author");
                writer.write_event(Event::Start(start))?;
                writer.write_event(Event::Text(BytesText::new(item.as_str())))?;
                writer.write_event(Event::End(BytesEnd::new("author")))?;
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
        if !self.author.is_empty() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for CommentList {
    fn write_children<W: Write>(&self, writer: &mut Writer<W>) -> Result<(), SerializeError> {
        #[cfg(feature = "extra-children")]
        let mut extra_iter = self.extra_children.iter().peekable();
        #[cfg(feature = "extra-children")]
        let mut emit_idx: usize = 0;
        for item in &self.comment {
            #[cfg(feature = "extra-children")]
            while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
                extra_iter
                    .next()
                    .unwrap()
                    .node
                    .write_to(writer)
                    .map_err(SerializeError::from)?;
            }
            item.write_element("comment", writer)?;
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

impl ToXml for Comment {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        #[cfg(feature = "sml-comments")]
        {
            let val = &self.reference;
            start.push_attribute(("ref", val.as_str()));
        }
        #[cfg(feature = "sml-comments")]
        {
            let val = &self.author_id;
            {
                let s = val.to_string();
                start.push_attribute(("authorId", s.as_str()));
            }
        }
        #[cfg(feature = "sml-comments")]
        if let Some(ref val) = self.guid {
            start.push_attribute(("guid", val.as_str()));
        }
        #[cfg(feature = "sml-comments")]
        if let Some(ref val) = self.shape_id {
            {
                let s = val.to_string();
                start.push_attribute(("shapeId", s.as_str()));
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
        #[cfg(feature = "sml-comments")]
        {
            let val = &self.text;
            val.write_element("text", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        if let Some(ref val) = self.comment_pr {
            val.write_element("commentPr", writer)?;
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
        #[cfg(feature = "sml-comments")]
        return false;
        if self.comment_pr.is_some() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for CTCommentPr {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        if let Some(ref val) = self.locked {
            start.push_attribute(("locked", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.default_size {
            start.push_attribute(("defaultSize", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.print {
            start.push_attribute(("print", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.disabled {
            start.push_attribute(("disabled", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.auto_fill {
            start.push_attribute(("autoFill", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.auto_line {
            start.push_attribute(("autoLine", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.alt_text {
            start.push_attribute(("altText", val.as_str()));
        }
        if let Some(ref val) = self.text_h_align {
            {
                let s = val.to_string();
                start.push_attribute(("textHAlign", s.as_str()));
            }
        }
        if let Some(ref val) = self.text_v_align {
            {
                let s = val.to_string();
                start.push_attribute(("textVAlign", s.as_str()));
            }
        }
        if let Some(ref val) = self.lock_text {
            start.push_attribute(("lockText", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.just_last_x {
            start.push_attribute(("justLastX", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.auto_scale {
            start.push_attribute(("autoScale", if *val { "1" } else { "0" }));
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
            let val = &self.anchor;
            val.write_element("anchor", writer)?;
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

impl ToXml for MapInfo {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        {
            let val = &self.selection_namespaces;
            start.push_attribute(("SelectionNamespaces", val.as_str()));
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
        for item in &self.schema {
            #[cfg(feature = "extra-children")]
            while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
                extra_iter
                    .next()
                    .unwrap()
                    .node
                    .write_to(writer)
                    .map_err(SerializeError::from)?;
            }
            item.write_element("Schema", writer)?;
            #[cfg(feature = "extra-children")]
            {
                emit_idx += 1;
            }
        }
        for item in &self.map {
            #[cfg(feature = "extra-children")]
            while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
                extra_iter
                    .next()
                    .unwrap()
                    .node
                    .write_to(writer)
                    .map_err(SerializeError::from)?;
            }
            item.write_element("Map", writer)?;
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
        if !self.schema.is_empty() {
            return false;
        }
        if !self.map.is_empty() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for XmlSchema {
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

impl ToXml for XmlMap {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        {
            let val = &self.i_d;
            {
                let s = val.to_string();
                start.push_attribute(("ID", s.as_str()));
            }
        }
        {
            let val = &self.name;
            start.push_attribute(("Name", val.as_str()));
        }
        {
            let val = &self.root_element;
            start.push_attribute(("RootElement", val.as_str()));
        }
        {
            let val = &self.schema_i_d;
            start.push_attribute(("SchemaID", val.as_str()));
        }
        {
            let val = &self.show_import_export_validation_errors;
            start.push_attribute((
                "ShowImportExportValidationErrors",
                if *val { "1" } else { "0" },
            ));
        }
        {
            let val = &self.auto_fit;
            start.push_attribute(("AutoFit", if *val { "1" } else { "0" }));
        }
        {
            let val = &self.append;
            start.push_attribute(("Append", if *val { "1" } else { "0" }));
        }
        {
            let val = &self.preserve_sort_a_f_layout;
            start.push_attribute(("PreserveSortAFLayout", if *val { "1" } else { "0" }));
        }
        {
            let val = &self.preserve_format;
            start.push_attribute(("PreserveFormat", if *val { "1" } else { "0" }));
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
        if let Some(ref val) = self.data_binding {
            val.write_element("DataBinding", writer)?;
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
        if self.data_binding.is_some() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for DataBinding {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        if let Some(ref val) = self.data_binding_name {
            start.push_attribute(("DataBindingName", val.as_str()));
        }
        if let Some(ref val) = self.file_binding {
            start.push_attribute(("FileBinding", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.connection_i_d {
            {
                let s = val.to_string();
                start.push_attribute(("ConnectionID", s.as_str()));
            }
        }
        if let Some(ref val) = self.file_binding_name {
            start.push_attribute(("FileBindingName", val.as_str()));
        }
        {
            let val = &self.data_binding_load_mode;
            {
                let s = val.to_string();
                start.push_attribute(("DataBindingLoadMode", s.as_str()));
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

impl ToXml for Connections {
    fn write_children<W: Write>(&self, writer: &mut Writer<W>) -> Result<(), SerializeError> {
        #[cfg(feature = "extra-children")]
        let mut extra_iter = self.extra_children.iter().peekable();
        #[cfg(feature = "extra-children")]
        let mut emit_idx: usize = 0;
        for item in &self.connection {
            #[cfg(feature = "extra-children")]
            while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
                extra_iter
                    .next()
                    .unwrap()
                    .node
                    .write_to(writer)
                    .map_err(SerializeError::from)?;
            }
            item.write_element("connection", writer)?;
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
        if !self.connection.is_empty() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for Connection {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        {
            let val = &self.id;
            {
                let s = val.to_string();
                start.push_attribute(("id", s.as_str()));
            }
        }
        if let Some(ref val) = self.source_file {
            start.push_attribute(("sourceFile", val.as_str()));
        }
        if let Some(ref val) = self.odc_file {
            start.push_attribute(("odcFile", val.as_str()));
        }
        if let Some(ref val) = self.keep_alive {
            start.push_attribute(("keepAlive", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.interval {
            {
                let s = val.to_string();
                start.push_attribute(("interval", s.as_str()));
            }
        }
        if let Some(ref val) = self.name {
            start.push_attribute(("name", val.as_str()));
        }
        if let Some(ref val) = self.description {
            start.push_attribute(("description", val.as_str()));
        }
        if let Some(ref val) = self.r#type {
            {
                let s = val.to_string();
                start.push_attribute(("type", s.as_str()));
            }
        }
        if let Some(ref val) = self.reconnection_method {
            {
                let s = val.to_string();
                start.push_attribute(("reconnectionMethod", s.as_str()));
            }
        }
        {
            let val = &self.refreshed_version;
            {
                let s = val.to_string();
                start.push_attribute(("refreshedVersion", s.as_str()));
            }
        }
        if let Some(ref val) = self.min_refreshable_version {
            {
                let s = val.to_string();
                start.push_attribute(("minRefreshableVersion", s.as_str()));
            }
        }
        if let Some(ref val) = self.save_password {
            start.push_attribute(("savePassword", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.new {
            start.push_attribute(("new", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.deleted {
            start.push_attribute(("deleted", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.only_use_connection_file {
            start.push_attribute(("onlyUseConnectionFile", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.background {
            start.push_attribute(("background", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.refresh_on_load {
            start.push_attribute(("refreshOnLoad", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.save_data {
            start.push_attribute(("saveData", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.credentials {
            {
                let s = val.to_string();
                start.push_attribute(("credentials", s.as_str()));
            }
        }
        if let Some(ref val) = self.single_sign_on_id {
            start.push_attribute(("singleSignOnId", val.as_str()));
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
        if let Some(ref val) = self.db_pr {
            val.write_element("dbPr", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        if let Some(ref val) = self.olap_pr {
            val.write_element("olapPr", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
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
            val.write_element("webPr", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        if let Some(ref val) = self.text_pr {
            val.write_element("textPr", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        if let Some(ref val) = self.parameters {
            val.write_element("parameters", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        if let Some(ref val) = self.extension_list {
            val.write_element("extLst", writer)?;
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
        if self.db_pr.is_some() {
            return false;
        }
        if self.olap_pr.is_some() {
            return false;
        }
        if self.web_pr.is_some() {
            return false;
        }
        if self.text_pr.is_some() {
            return false;
        }
        if self.parameters.is_some() {
            return false;
        }
        if self.extension_list.is_some() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for DatabaseProperties {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        {
            let val = &self.connection;
            start.push_attribute(("connection", val.as_str()));
        }
        if let Some(ref val) = self.command {
            start.push_attribute(("command", val.as_str()));
        }
        if let Some(ref val) = self.server_command {
            start.push_attribute(("serverCommand", val.as_str()));
        }
        if let Some(ref val) = self.command_type {
            {
                let s = val.to_string();
                start.push_attribute(("commandType", s.as_str()));
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

impl ToXml for OlapProperties {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        if let Some(ref val) = self.local {
            start.push_attribute(("local", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.local_connection {
            start.push_attribute(("localConnection", val.as_str()));
        }
        if let Some(ref val) = self.local_refresh {
            start.push_attribute(("localRefresh", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.send_locale {
            start.push_attribute(("sendLocale", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.row_drill_count {
            {
                let s = val.to_string();
                start.push_attribute(("rowDrillCount", s.as_str()));
            }
        }
        if let Some(ref val) = self.server_fill {
            start.push_attribute(("serverFill", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.server_number_format {
            start.push_attribute(("serverNumberFormat", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.server_font {
            start.push_attribute(("serverFont", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.server_font_color {
            start.push_attribute(("serverFontColor", if *val { "1" } else { "0" }));
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

impl ToXml for WebQueryProperties {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        if let Some(ref val) = self.xml {
            start.push_attribute(("xml", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.source_data {
            start.push_attribute(("sourceData", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.parse_pre {
            start.push_attribute(("parsePre", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.consecutive {
            start.push_attribute(("consecutive", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.first_row {
            start.push_attribute(("firstRow", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.xl97 {
            start.push_attribute(("xl97", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.text_dates {
            start.push_attribute(("textDates", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.xl2000 {
            start.push_attribute(("xl2000", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.url {
            start.push_attribute(("url", val.as_str()));
        }
        if let Some(ref val) = self.post {
            start.push_attribute(("post", val.as_str()));
        }
        if let Some(ref val) = self.html_tables {
            start.push_attribute(("htmlTables", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.html_format {
            {
                let s = val.to_string();
                start.push_attribute(("htmlFormat", s.as_str()));
            }
        }
        if let Some(ref val) = self.edit_page {
            start.push_attribute(("editPage", val.as_str()));
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
        if let Some(ref val) = self.tables {
            val.write_element("tables", writer)?;
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
        if self.tables.is_some() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for Parameters {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        if let Some(ref val) = self.count {
            {
                let s = val.to_string();
                start.push_attribute(("count", s.as_str()));
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
        for item in &self.parameter {
            #[cfg(feature = "extra-children")]
            while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
                extra_iter
                    .next()
                    .unwrap()
                    .node
                    .write_to(writer)
                    .map_err(SerializeError::from)?;
            }
            item.write_element("parameter", writer)?;
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
        if !self.parameter.is_empty() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for Parameter {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        if let Some(ref val) = self.name {
            start.push_attribute(("name", val.as_str()));
        }
        if let Some(ref val) = self.sql_type {
            {
                let s = val.to_string();
                start.push_attribute(("sqlType", s.as_str()));
            }
        }
        if let Some(ref val) = self.parameter_type {
            {
                let s = val.to_string();
                start.push_attribute(("parameterType", s.as_str()));
            }
        }
        if let Some(ref val) = self.refresh_on_change {
            start.push_attribute(("refreshOnChange", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.prompt {
            start.push_attribute(("prompt", val.as_str()));
        }
        if let Some(ref val) = self.boolean {
            start.push_attribute(("boolean", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.double {
            {
                let s = val.to_string();
                start.push_attribute(("double", s.as_str()));
            }
        }
        if let Some(ref val) = self.integer {
            {
                let s = val.to_string();
                start.push_attribute(("integer", s.as_str()));
            }
        }
        if let Some(ref val) = self.string {
            start.push_attribute(("string", val.as_str()));
        }
        if let Some(ref val) = self.cell {
            start.push_attribute(("cell", val.as_str()));
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

impl ToXml for DataTables {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        if let Some(ref val) = self.count {
            {
                let s = val.to_string();
                start.push_attribute(("count", s.as_str()));
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

impl ToXml for TableMissing {
    fn is_empty_element(&self) -> bool {
        true
    }
}

impl ToXml for TextImportProperties {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        if let Some(ref val) = self.prompt {
            start.push_attribute(("prompt", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.file_type {
            {
                let s = val.to_string();
                start.push_attribute(("fileType", s.as_str()));
            }
        }
        if let Some(ref val) = self.code_page {
            {
                let s = val.to_string();
                start.push_attribute(("codePage", s.as_str()));
            }
        }
        if let Some(ref val) = self.character_set {
            start.push_attribute(("characterSet", val.as_str()));
        }
        if let Some(ref val) = self.first_row {
            {
                let s = val.to_string();
                start.push_attribute(("firstRow", s.as_str()));
            }
        }
        if let Some(ref val) = self.source_file {
            start.push_attribute(("sourceFile", val.as_str()));
        }
        if let Some(ref val) = self.delimited {
            start.push_attribute(("delimited", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.decimal {
            start.push_attribute(("decimal", val.as_str()));
        }
        if let Some(ref val) = self.thousands {
            start.push_attribute(("thousands", val.as_str()));
        }
        if let Some(ref val) = self.tab {
            start.push_attribute(("tab", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.space {
            start.push_attribute(("space", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.comma {
            start.push_attribute(("comma", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.semicolon {
            start.push_attribute(("semicolon", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.consecutive {
            start.push_attribute(("consecutive", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.qualifier {
            {
                let s = val.to_string();
                start.push_attribute(("qualifier", s.as_str()));
            }
        }
        if let Some(ref val) = self.delimiter {
            start.push_attribute(("delimiter", val.as_str()));
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
        if let Some(ref val) = self.text_fields {
            val.write_element("textFields", writer)?;
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
        if self.text_fields.is_some() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for TextFields {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        if let Some(ref val) = self.count {
            {
                let s = val.to_string();
                start.push_attribute(("count", s.as_str()));
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
        for item in &self.text_field {
            #[cfg(feature = "extra-children")]
            while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
                extra_iter
                    .next()
                    .unwrap()
                    .node
                    .write_to(writer)
                    .map_err(SerializeError::from)?;
            }
            item.write_element("textField", writer)?;
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
        if !self.text_field.is_empty() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for TextField {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        if let Some(ref val) = self.r#type {
            {
                let s = val.to_string();
                start.push_attribute(("type", s.as_str()));
            }
        }
        if let Some(ref val) = self.position {
            {
                let s = val.to_string();
                start.push_attribute(("position", s.as_str()));
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

impl ToXml for PivotCacheDefinition {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        if let Some(ref val) = self.invalid {
            start.push_attribute(("invalid", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.save_data {
            start.push_attribute(("saveData", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.refresh_on_load {
            start.push_attribute(("refreshOnLoad", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.optimize_memory {
            start.push_attribute(("optimizeMemory", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.enable_refresh {
            start.push_attribute(("enableRefresh", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.refreshed_by {
            start.push_attribute(("refreshedBy", val.as_str()));
        }
        if let Some(ref val) = self.refreshed_date {
            {
                let s = val.to_string();
                start.push_attribute(("refreshedDate", s.as_str()));
            }
        }
        if let Some(ref val) = self.refreshed_date_iso {
            start.push_attribute(("refreshedDateIso", val.as_str()));
        }
        if let Some(ref val) = self.background_query {
            start.push_attribute(("backgroundQuery", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.missing_items_limit {
            {
                let s = val.to_string();
                start.push_attribute(("missingItemsLimit", s.as_str()));
            }
        }
        if let Some(ref val) = self.created_version {
            {
                let s = val.to_string();
                start.push_attribute(("createdVersion", s.as_str()));
            }
        }
        if let Some(ref val) = self.refreshed_version {
            {
                let s = val.to_string();
                start.push_attribute(("refreshedVersion", s.as_str()));
            }
        }
        if let Some(ref val) = self.min_refreshable_version {
            {
                let s = val.to_string();
                start.push_attribute(("minRefreshableVersion", s.as_str()));
            }
        }
        if let Some(ref val) = self.record_count {
            {
                let s = val.to_string();
                start.push_attribute(("recordCount", s.as_str()));
            }
        }
        if let Some(ref val) = self.upgrade_on_refresh {
            start.push_attribute(("upgradeOnRefresh", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.tuple_cache {
            start.push_attribute(("tupleCache", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.support_subquery {
            start.push_attribute(("supportSubquery", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.support_advanced_drill {
            start.push_attribute(("supportAdvancedDrill", if *val { "1" } else { "0" }));
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
            let val = &self.cache_source;
            val.write_element("cacheSource", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
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
            let val = &self.cache_fields;
            val.write_element("cacheFields", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        if let Some(ref val) = self.cache_hierarchies {
            val.write_element("cacheHierarchies", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        if let Some(ref val) = self.kpis {
            val.write_element("kpis", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        if let Some(ref val) = self.calculated_items {
            val.write_element("calculatedItems", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        if let Some(ref val) = self.calculated_members {
            val.write_element("calculatedMembers", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        if let Some(ref val) = self.dimensions {
            val.write_element("dimensions", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        if let Some(ref val) = self.measure_groups {
            val.write_element("measureGroups", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        if let Some(ref val) = self.maps {
            val.write_element("maps", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        if let Some(ref val) = self.extension_list {
            val.write_element("extLst", writer)?;
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

impl ToXml for CacheFields {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        if let Some(ref val) = self.count {
            {
                let s = val.to_string();
                start.push_attribute(("count", s.as_str()));
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
        for item in &self.cache_field {
            #[cfg(feature = "extra-children")]
            while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
                extra_iter
                    .next()
                    .unwrap()
                    .node
                    .write_to(writer)
                    .map_err(SerializeError::from)?;
            }
            item.write_element("cacheField", writer)?;
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
        if !self.cache_field.is_empty() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for CacheField {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        {
            let val = &self.name;
            start.push_attribute(("name", val.as_str()));
        }
        if let Some(ref val) = self.caption {
            start.push_attribute(("caption", val.as_str()));
        }
        if let Some(ref val) = self.property_name {
            start.push_attribute(("propertyName", val.as_str()));
        }
        if let Some(ref val) = self.server_field {
            start.push_attribute(("serverField", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.unique_list {
            start.push_attribute(("uniqueList", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.number_format_id {
            {
                let s = val.to_string();
                start.push_attribute(("numFmtId", s.as_str()));
            }
        }
        if let Some(ref val) = self.formula {
            start.push_attribute(("formula", val.as_str()));
        }
        if let Some(ref val) = self.sql_type {
            {
                let s = val.to_string();
                start.push_attribute(("sqlType", s.as_str()));
            }
        }
        if let Some(ref val) = self.hierarchy {
            {
                let s = val.to_string();
                start.push_attribute(("hierarchy", s.as_str()));
            }
        }
        if let Some(ref val) = self.level {
            {
                let s = val.to_string();
                start.push_attribute(("level", s.as_str()));
            }
        }
        if let Some(ref val) = self.database_field {
            start.push_attribute(("databaseField", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.mapping_count {
            {
                let s = val.to_string();
                start.push_attribute(("mappingCount", s.as_str()));
            }
        }
        if let Some(ref val) = self.member_property_field {
            start.push_attribute(("memberPropertyField", if *val { "1" } else { "0" }));
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
        if let Some(ref val) = self.shared_items {
            val.write_element("sharedItems", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        if let Some(ref val) = self.field_group {
            val.write_element("fieldGroup", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        for item in &self.mp_map {
            #[cfg(feature = "extra-children")]
            while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
                extra_iter
                    .next()
                    .unwrap()
                    .node
                    .write_to(writer)
                    .map_err(SerializeError::from)?;
            }
            item.write_element("mpMap", writer)?;
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
        if let Some(ref val) = self.extension_list {
            val.write_element("extLst", writer)?;
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
        if self.shared_items.is_some() {
            return false;
        }
        if self.field_group.is_some() {
            return false;
        }
        if !self.mp_map.is_empty() {
            return false;
        }
        if self.extension_list.is_some() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for CacheSource {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        {
            let val = &self.r#type;
            {
                let s = val.to_string();
                start.push_attribute(("type", s.as_str()));
            }
        }
        if let Some(ref val) = self.connection_id {
            {
                let s = val.to_string();
                start.push_attribute(("connectionId", s.as_str()));
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

impl ToXml for WorksheetSource {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        if let Some(ref val) = self.reference {
            start.push_attribute(("ref", val.as_str()));
        }
        if let Some(ref val) = self.name {
            start.push_attribute(("name", val.as_str()));
        }
        if let Some(ref val) = self.sheet {
            start.push_attribute(("sheet", val.as_str()));
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

impl ToXml for Consolidation {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        if let Some(ref val) = self.auto_page {
            start.push_attribute(("autoPage", if *val { "1" } else { "0" }));
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
        if let Some(ref val) = self.pages {
            val.write_element("pages", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
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
            let val = &self.range_sets;
            val.write_element("rangeSets", writer)?;
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
        if self.pages.is_some() {
            return false;
        }
        false
    }
}

impl ToXml for CTPages {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        if let Some(ref val) = self.count {
            {
                let s = val.to_string();
                start.push_attribute(("count", s.as_str()));
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
        for item in &self.page {
            #[cfg(feature = "extra-children")]
            while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
                extra_iter
                    .next()
                    .unwrap()
                    .node
                    .write_to(writer)
                    .map_err(SerializeError::from)?;
            }
            item.write_element("page", writer)?;
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
        if !self.page.is_empty() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for CTPCDSCPage {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        if let Some(ref val) = self.count {
            {
                let s = val.to_string();
                start.push_attribute(("count", s.as_str()));
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
        for item in &self.page_item {
            #[cfg(feature = "extra-children")]
            while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
                extra_iter
                    .next()
                    .unwrap()
                    .node
                    .write_to(writer)
                    .map_err(SerializeError::from)?;
            }
            item.write_element("pageItem", writer)?;
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
        if !self.page_item.is_empty() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for CTPageItem {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
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

    fn is_empty_element(&self) -> bool {
        true
    }
}

impl ToXml for CTRangeSets {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        if let Some(ref val) = self.count {
            {
                let s = val.to_string();
                start.push_attribute(("count", s.as_str()));
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
        for item in &self.range_set {
            #[cfg(feature = "extra-children")]
            while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
                extra_iter
                    .next()
                    .unwrap()
                    .node
                    .write_to(writer)
                    .map_err(SerializeError::from)?;
            }
            item.write_element("rangeSet", writer)?;
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
        if !self.range_set.is_empty() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for CTRangeSet {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        if let Some(ref val) = self.i1 {
            {
                let s = val.to_string();
                start.push_attribute(("i1", s.as_str()));
            }
        }
        if let Some(ref val) = self.i2 {
            {
                let s = val.to_string();
                start.push_attribute(("i2", s.as_str()));
            }
        }
        if let Some(ref val) = self.i3 {
            {
                let s = val.to_string();
                start.push_attribute(("i3", s.as_str()));
            }
        }
        if let Some(ref val) = self.i4 {
            {
                let s = val.to_string();
                start.push_attribute(("i4", s.as_str()));
            }
        }
        if let Some(ref val) = self.reference {
            start.push_attribute(("ref", val.as_str()));
        }
        if let Some(ref val) = self.name {
            start.push_attribute(("name", val.as_str()));
        }
        if let Some(ref val) = self.sheet {
            start.push_attribute(("sheet", val.as_str()));
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

impl ToXml for SharedItems {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        if let Some(ref val) = self.contains_semi_mixed_types {
            start.push_attribute(("containsSemiMixedTypes", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.contains_non_date {
            start.push_attribute(("containsNonDate", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.contains_date {
            start.push_attribute(("containsDate", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.contains_string {
            start.push_attribute(("containsString", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.contains_blank {
            start.push_attribute(("containsBlank", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.contains_mixed_types {
            start.push_attribute(("containsMixedTypes", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.contains_number {
            start.push_attribute(("containsNumber", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.contains_integer {
            start.push_attribute(("containsInteger", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.min_value {
            {
                let s = val.to_string();
                start.push_attribute(("minValue", s.as_str()));
            }
        }
        if let Some(ref val) = self.max_value {
            {
                let s = val.to_string();
                start.push_attribute(("maxValue", s.as_str()));
            }
        }
        if let Some(ref val) = self.min_date {
            start.push_attribute(("minDate", val.as_str()));
        }
        if let Some(ref val) = self.max_date {
            start.push_attribute(("maxDate", val.as_str()));
        }
        if let Some(ref val) = self.count {
            {
                let s = val.to_string();
                start.push_attribute(("count", s.as_str()));
            }
        }
        if let Some(ref val) = self.long_text {
            start.push_attribute(("longText", if *val { "1" } else { "0" }));
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

impl ToXml for CTMissing {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        if let Some(ref val) = self.u {
            start.push_attribute(("u", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.formula {
            start.push_attribute(("f", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.cells {
            start.push_attribute(("c", val.as_str()));
        }
        if let Some(ref val) = self.cp {
            {
                let s = val.to_string();
                start.push_attribute(("cp", s.as_str()));
            }
        }
        if let Some(ref val) = self.r#in {
            {
                let s = val.to_string();
                start.push_attribute(("in", s.as_str()));
            }
        }
        if let Some(ref val) = self.bc {
            {
                let hex = encode_hex(val);
                start.push_attribute(("bc", hex.as_str()));
            }
        }
        if let Some(ref val) = self.fc {
            {
                let hex = encode_hex(val);
                start.push_attribute(("fc", hex.as_str()));
            }
        }
        if let Some(ref val) = self.i {
            start.push_attribute(("i", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.un {
            start.push_attribute(("un", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.st {
            start.push_attribute(("st", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.b {
            start.push_attribute(("b", if *val { "1" } else { "0" }));
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
        for item in &self.tpls {
            #[cfg(feature = "extra-children")]
            while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
                extra_iter
                    .next()
                    .unwrap()
                    .node
                    .write_to(writer)
                    .map_err(SerializeError::from)?;
            }
            item.write_element("tpls", writer)?;
            #[cfg(feature = "extra-children")]
            {
                emit_idx += 1;
            }
        }
        for item in &self.x {
            #[cfg(feature = "extra-children")]
            while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
                extra_iter
                    .next()
                    .unwrap()
                    .node
                    .write_to(writer)
                    .map_err(SerializeError::from)?;
            }
            item.write_element("x", writer)?;
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
        if !self.tpls.is_empty() {
            return false;
        }
        if !self.x.is_empty() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for CTNumber {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        {
            let val = &self.value;
            {
                let s = val.to_string();
                start.push_attribute(("v", s.as_str()));
            }
        }
        if let Some(ref val) = self.u {
            start.push_attribute(("u", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.formula {
            start.push_attribute(("f", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.cells {
            start.push_attribute(("c", val.as_str()));
        }
        if let Some(ref val) = self.cp {
            {
                let s = val.to_string();
                start.push_attribute(("cp", s.as_str()));
            }
        }
        if let Some(ref val) = self.r#in {
            {
                let s = val.to_string();
                start.push_attribute(("in", s.as_str()));
            }
        }
        if let Some(ref val) = self.bc {
            {
                let hex = encode_hex(val);
                start.push_attribute(("bc", hex.as_str()));
            }
        }
        if let Some(ref val) = self.fc {
            {
                let hex = encode_hex(val);
                start.push_attribute(("fc", hex.as_str()));
            }
        }
        if let Some(ref val) = self.i {
            start.push_attribute(("i", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.un {
            start.push_attribute(("un", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.st {
            start.push_attribute(("st", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.b {
            start.push_attribute(("b", if *val { "1" } else { "0" }));
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
        for item in &self.tpls {
            #[cfg(feature = "extra-children")]
            while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
                extra_iter
                    .next()
                    .unwrap()
                    .node
                    .write_to(writer)
                    .map_err(SerializeError::from)?;
            }
            item.write_element("tpls", writer)?;
            #[cfg(feature = "extra-children")]
            {
                emit_idx += 1;
            }
        }
        for item in &self.x {
            #[cfg(feature = "extra-children")]
            while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
                extra_iter
                    .next()
                    .unwrap()
                    .node
                    .write_to(writer)
                    .map_err(SerializeError::from)?;
            }
            item.write_element("x", writer)?;
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
        if !self.tpls.is_empty() {
            return false;
        }
        if !self.x.is_empty() {
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
        let mut start = start;
        {
            let val = &self.value;
            start.push_attribute(("v", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.u {
            start.push_attribute(("u", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.formula {
            start.push_attribute(("f", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.cells {
            start.push_attribute(("c", val.as_str()));
        }
        if let Some(ref val) = self.cp {
            {
                let s = val.to_string();
                start.push_attribute(("cp", s.as_str()));
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
        for item in &self.x {
            #[cfg(feature = "extra-children")]
            while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
                extra_iter
                    .next()
                    .unwrap()
                    .node
                    .write_to(writer)
                    .map_err(SerializeError::from)?;
            }
            item.write_element("x", writer)?;
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
        if !self.x.is_empty() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for CTError {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        {
            let val = &self.value;
            start.push_attribute(("v", val.as_str()));
        }
        if let Some(ref val) = self.u {
            start.push_attribute(("u", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.formula {
            start.push_attribute(("f", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.cells {
            start.push_attribute(("c", val.as_str()));
        }
        if let Some(ref val) = self.cp {
            {
                let s = val.to_string();
                start.push_attribute(("cp", s.as_str()));
            }
        }
        if let Some(ref val) = self.r#in {
            {
                let s = val.to_string();
                start.push_attribute(("in", s.as_str()));
            }
        }
        if let Some(ref val) = self.bc {
            {
                let hex = encode_hex(val);
                start.push_attribute(("bc", hex.as_str()));
            }
        }
        if let Some(ref val) = self.fc {
            {
                let hex = encode_hex(val);
                start.push_attribute(("fc", hex.as_str()));
            }
        }
        if let Some(ref val) = self.i {
            start.push_attribute(("i", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.un {
            start.push_attribute(("un", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.st {
            start.push_attribute(("st", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.b {
            start.push_attribute(("b", if *val { "1" } else { "0" }));
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
        if let Some(ref val) = self.tpls {
            val.write_element("tpls", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        for item in &self.x {
            #[cfg(feature = "extra-children")]
            while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
                extra_iter
                    .next()
                    .unwrap()
                    .node
                    .write_to(writer)
                    .map_err(SerializeError::from)?;
            }
            item.write_element("x", writer)?;
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
        if self.tpls.is_some() {
            return false;
        }
        if !self.x.is_empty() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for CTString {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        {
            let val = &self.value;
            start.push_attribute(("v", val.as_str()));
        }
        if let Some(ref val) = self.u {
            start.push_attribute(("u", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.formula {
            start.push_attribute(("f", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.cells {
            start.push_attribute(("c", val.as_str()));
        }
        if let Some(ref val) = self.cp {
            {
                let s = val.to_string();
                start.push_attribute(("cp", s.as_str()));
            }
        }
        if let Some(ref val) = self.r#in {
            {
                let s = val.to_string();
                start.push_attribute(("in", s.as_str()));
            }
        }
        if let Some(ref val) = self.bc {
            {
                let hex = encode_hex(val);
                start.push_attribute(("bc", hex.as_str()));
            }
        }
        if let Some(ref val) = self.fc {
            {
                let hex = encode_hex(val);
                start.push_attribute(("fc", hex.as_str()));
            }
        }
        if let Some(ref val) = self.i {
            start.push_attribute(("i", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.un {
            start.push_attribute(("un", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.st {
            start.push_attribute(("st", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.b {
            start.push_attribute(("b", if *val { "1" } else { "0" }));
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
        for item in &self.tpls {
            #[cfg(feature = "extra-children")]
            while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
                extra_iter
                    .next()
                    .unwrap()
                    .node
                    .write_to(writer)
                    .map_err(SerializeError::from)?;
            }
            item.write_element("tpls", writer)?;
            #[cfg(feature = "extra-children")]
            {
                emit_idx += 1;
            }
        }
        for item in &self.x {
            #[cfg(feature = "extra-children")]
            while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
                extra_iter
                    .next()
                    .unwrap()
                    .node
                    .write_to(writer)
                    .map_err(SerializeError::from)?;
            }
            item.write_element("x", writer)?;
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
        if !self.tpls.is_empty() {
            return false;
        }
        if !self.x.is_empty() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for CTDateTime {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        {
            let val = &self.value;
            start.push_attribute(("v", val.as_str()));
        }
        if let Some(ref val) = self.u {
            start.push_attribute(("u", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.formula {
            start.push_attribute(("f", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.cells {
            start.push_attribute(("c", val.as_str()));
        }
        if let Some(ref val) = self.cp {
            {
                let s = val.to_string();
                start.push_attribute(("cp", s.as_str()));
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
        for item in &self.x {
            #[cfg(feature = "extra-children")]
            while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
                extra_iter
                    .next()
                    .unwrap()
                    .node
                    .write_to(writer)
                    .map_err(SerializeError::from)?;
            }
            item.write_element("x", writer)?;
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
        if !self.x.is_empty() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for FieldGroup {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        if let Some(ref val) = self.par {
            {
                let s = val.to_string();
                start.push_attribute(("par", s.as_str()));
            }
        }
        if let Some(ref val) = self.base {
            {
                let s = val.to_string();
                start.push_attribute(("base", s.as_str()));
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
        if let Some(ref val) = self.range_pr {
            val.write_element("rangePr", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        if let Some(ref val) = self.discrete_pr {
            val.write_element("discretePr", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        if let Some(ref val) = self.group_items {
            val.write_element("groupItems", writer)?;
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
        if self.range_pr.is_some() {
            return false;
        }
        if self.discrete_pr.is_some() {
            return false;
        }
        if self.group_items.is_some() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for CTRangePr {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        if let Some(ref val) = self.auto_start {
            start.push_attribute(("autoStart", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.auto_end {
            start.push_attribute(("autoEnd", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.group_by {
            {
                let s = val.to_string();
                start.push_attribute(("groupBy", s.as_str()));
            }
        }
        if let Some(ref val) = self.start_num {
            {
                let s = val.to_string();
                start.push_attribute(("startNum", s.as_str()));
            }
        }
        if let Some(ref val) = self.end_num {
            {
                let s = val.to_string();
                start.push_attribute(("endNum", s.as_str()));
            }
        }
        if let Some(ref val) = self.start_date {
            start.push_attribute(("startDate", val.as_str()));
        }
        if let Some(ref val) = self.end_date {
            start.push_attribute(("endDate", val.as_str()));
        }
        if let Some(ref val) = self.group_interval {
            {
                let s = val.to_string();
                start.push_attribute(("groupInterval", s.as_str()));
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

impl ToXml for CTDiscretePr {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        if let Some(ref val) = self.count {
            {
                let s = val.to_string();
                start.push_attribute(("count", s.as_str()));
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
        for item in &self.x {
            #[cfg(feature = "extra-children")]
            while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
                extra_iter
                    .next()
                    .unwrap()
                    .node
                    .write_to(writer)
                    .map_err(SerializeError::from)?;
            }
            item.write_element("x", writer)?;
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
        if !self.x.is_empty() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for GroupItems {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        if let Some(ref val) = self.count {
            {
                let s = val.to_string();
                start.push_attribute(("count", s.as_str()));
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

impl ToXml for PivotCacheRecords {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        if let Some(ref val) = self.count {
            {
                let s = val.to_string();
                start.push_attribute(("count", s.as_str()));
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
        for item in &self.reference {
            #[cfg(feature = "extra-children")]
            while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
                extra_iter
                    .next()
                    .unwrap()
                    .node
                    .write_to(writer)
                    .map_err(SerializeError::from)?;
            }
            item.write_element("r", writer)?;
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
        if let Some(ref val) = self.extension_list {
            val.write_element("extLst", writer)?;
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
        if !self.reference.is_empty() {
            return false;
        }
        if self.extension_list.is_some() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for CTRecord {
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

impl ToXml for CTPCDKPIs {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        if let Some(ref val) = self.count {
            {
                let s = val.to_string();
                start.push_attribute(("count", s.as_str()));
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
        for item in &self.kpi {
            #[cfg(feature = "extra-children")]
            while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
                extra_iter
                    .next()
                    .unwrap()
                    .node
                    .write_to(writer)
                    .map_err(SerializeError::from)?;
            }
            item.write_element("kpi", writer)?;
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
        if !self.kpi.is_empty() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for CTPCDKPI {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        {
            let val = &self.unique_name;
            start.push_attribute(("uniqueName", val.as_str()));
        }
        if let Some(ref val) = self.caption {
            start.push_attribute(("caption", val.as_str()));
        }
        if let Some(ref val) = self.display_folder {
            start.push_attribute(("displayFolder", val.as_str()));
        }
        if let Some(ref val) = self.measure_group {
            start.push_attribute(("measureGroup", val.as_str()));
        }
        if let Some(ref val) = self.parent {
            start.push_attribute(("parent", val.as_str()));
        }
        {
            let val = &self.value;
            start.push_attribute(("value", val.as_str()));
        }
        if let Some(ref val) = self.goal {
            start.push_attribute(("goal", val.as_str()));
        }
        if let Some(ref val) = self.status {
            start.push_attribute(("status", val.as_str()));
        }
        if let Some(ref val) = self.trend {
            start.push_attribute(("trend", val.as_str()));
        }
        if let Some(ref val) = self.weight {
            start.push_attribute(("weight", val.as_str()));
        }
        if let Some(ref val) = self.time {
            start.push_attribute(("time", val.as_str()));
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

impl ToXml for CTCacheHierarchies {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        if let Some(ref val) = self.count {
            {
                let s = val.to_string();
                start.push_attribute(("count", s.as_str()));
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
        for item in &self.cache_hierarchy {
            #[cfg(feature = "extra-children")]
            while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
                extra_iter
                    .next()
                    .unwrap()
                    .node
                    .write_to(writer)
                    .map_err(SerializeError::from)?;
            }
            item.write_element("cacheHierarchy", writer)?;
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
        if !self.cache_hierarchy.is_empty() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for CTCacheHierarchy {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        {
            let val = &self.unique_name;
            start.push_attribute(("uniqueName", val.as_str()));
        }
        if let Some(ref val) = self.caption {
            start.push_attribute(("caption", val.as_str()));
        }
        if let Some(ref val) = self.measure {
            start.push_attribute(("measure", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.set {
            start.push_attribute(("set", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.parent_set {
            {
                let s = val.to_string();
                start.push_attribute(("parentSet", s.as_str()));
            }
        }
        if let Some(ref val) = self.icon_set {
            {
                let s = val.to_string();
                start.push_attribute(("iconSet", s.as_str()));
            }
        }
        if let Some(ref val) = self.attribute {
            start.push_attribute(("attribute", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.time {
            start.push_attribute(("time", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.key_attribute {
            start.push_attribute(("keyAttribute", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.default_member_unique_name {
            start.push_attribute(("defaultMemberUniqueName", val.as_str()));
        }
        if let Some(ref val) = self.all_unique_name {
            start.push_attribute(("allUniqueName", val.as_str()));
        }
        if let Some(ref val) = self.all_caption {
            start.push_attribute(("allCaption", val.as_str()));
        }
        if let Some(ref val) = self.dimension_unique_name {
            start.push_attribute(("dimensionUniqueName", val.as_str()));
        }
        if let Some(ref val) = self.display_folder {
            start.push_attribute(("displayFolder", val.as_str()));
        }
        if let Some(ref val) = self.measure_group {
            start.push_attribute(("measureGroup", val.as_str()));
        }
        if let Some(ref val) = self.measures {
            start.push_attribute(("measures", if *val { "1" } else { "0" }));
        }
        {
            let val = &self.count;
            {
                let s = val.to_string();
                start.push_attribute(("count", s.as_str()));
            }
        }
        if let Some(ref val) = self.one_field {
            start.push_attribute(("oneField", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.member_value_datatype {
            {
                let s = val.to_string();
                start.push_attribute(("memberValueDatatype", s.as_str()));
            }
        }
        if let Some(ref val) = self.unbalanced {
            start.push_attribute(("unbalanced", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.unbalanced_group {
            start.push_attribute(("unbalancedGroup", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.hidden {
            start.push_attribute(("hidden", if *val { "1" } else { "0" }));
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
        if let Some(ref val) = self.fields_usage {
            val.write_element("fieldsUsage", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        if let Some(ref val) = self.group_levels {
            val.write_element("groupLevels", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        if let Some(ref val) = self.extension_list {
            val.write_element("extLst", writer)?;
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
        if self.fields_usage.is_some() {
            return false;
        }
        if self.group_levels.is_some() {
            return false;
        }
        if self.extension_list.is_some() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for CTFieldsUsage {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        if let Some(ref val) = self.count {
            {
                let s = val.to_string();
                start.push_attribute(("count", s.as_str()));
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
        for item in &self.field_usage {
            #[cfg(feature = "extra-children")]
            while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
                extra_iter
                    .next()
                    .unwrap()
                    .node
                    .write_to(writer)
                    .map_err(SerializeError::from)?;
            }
            item.write_element("fieldUsage", writer)?;
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
        if !self.field_usage.is_empty() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for CTFieldUsage {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        {
            let val = &self.x;
            {
                let s = val.to_string();
                start.push_attribute(("x", s.as_str()));
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

impl ToXml for CTGroupLevels {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        if let Some(ref val) = self.count {
            {
                let s = val.to_string();
                start.push_attribute(("count", s.as_str()));
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
        for item in &self.group_level {
            #[cfg(feature = "extra-children")]
            while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
                extra_iter
                    .next()
                    .unwrap()
                    .node
                    .write_to(writer)
                    .map_err(SerializeError::from)?;
            }
            item.write_element("groupLevel", writer)?;
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
        if !self.group_level.is_empty() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for CTGroupLevel {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        {
            let val = &self.unique_name;
            start.push_attribute(("uniqueName", val.as_str()));
        }
        {
            let val = &self.caption;
            start.push_attribute(("caption", val.as_str()));
        }
        if let Some(ref val) = self.user {
            start.push_attribute(("user", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.custom_roll_up {
            start.push_attribute(("customRollUp", if *val { "1" } else { "0" }));
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
        if let Some(ref val) = self.groups {
            val.write_element("groups", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        if let Some(ref val) = self.extension_list {
            val.write_element("extLst", writer)?;
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
        if self.groups.is_some() {
            return false;
        }
        if self.extension_list.is_some() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for CTGroups {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        if let Some(ref val) = self.count {
            {
                let s = val.to_string();
                start.push_attribute(("count", s.as_str()));
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
        for item in &self.group {
            #[cfg(feature = "extra-children")]
            while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
                extra_iter
                    .next()
                    .unwrap()
                    .node
                    .write_to(writer)
                    .map_err(SerializeError::from)?;
            }
            item.write_element("group", writer)?;
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
        if !self.group.is_empty() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for CTLevelGroup {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        {
            let val = &self.name;
            start.push_attribute(("name", val.as_str()));
        }
        {
            let val = &self.unique_name;
            start.push_attribute(("uniqueName", val.as_str()));
        }
        {
            let val = &self.caption;
            start.push_attribute(("caption", val.as_str()));
        }
        if let Some(ref val) = self.unique_parent {
            start.push_attribute(("uniqueParent", val.as_str()));
        }
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
        {
            let val = &self.group_members;
            val.write_element("groupMembers", writer)?;
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

impl ToXml for CTGroupMembers {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        if let Some(ref val) = self.count {
            {
                let s = val.to_string();
                start.push_attribute(("count", s.as_str()));
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
        for item in &self.group_member {
            #[cfg(feature = "extra-children")]
            while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
                extra_iter
                    .next()
                    .unwrap()
                    .node
                    .write_to(writer)
                    .map_err(SerializeError::from)?;
            }
            item.write_element("groupMember", writer)?;
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
        if !self.group_member.is_empty() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for CTGroupMember {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        {
            let val = &self.unique_name;
            start.push_attribute(("uniqueName", val.as_str()));
        }
        if let Some(ref val) = self.group {
            start.push_attribute(("group", if *val { "1" } else { "0" }));
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

impl ToXml for CTTupleCache {
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
        if let Some(ref val) = self.entries {
            val.write_element("entries", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        if let Some(ref val) = self.sets {
            val.write_element("sets", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        if let Some(ref val) = self.query_cache {
            val.write_element("queryCache", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        if let Some(ref val) = self.server_formats {
            val.write_element("serverFormats", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        if let Some(ref val) = self.extension_list {
            val.write_element("extLst", writer)?;
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
        if self.entries.is_some() {
            return false;
        }
        if self.sets.is_some() {
            return false;
        }
        if self.query_cache.is_some() {
            return false;
        }
        if self.server_formats.is_some() {
            return false;
        }
        if self.extension_list.is_some() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for CTServerFormat {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        if let Some(ref val) = self.culture {
            start.push_attribute(("culture", val.as_str()));
        }
        if let Some(ref val) = self.format {
            start.push_attribute(("format", val.as_str()));
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

impl ToXml for CTServerFormats {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        if let Some(ref val) = self.count {
            {
                let s = val.to_string();
                start.push_attribute(("count", s.as_str()));
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
        for item in &self.server_format {
            #[cfg(feature = "extra-children")]
            while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
                extra_iter
                    .next()
                    .unwrap()
                    .node
                    .write_to(writer)
                    .map_err(SerializeError::from)?;
            }
            item.write_element("serverFormat", writer)?;
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
        if !self.server_format.is_empty() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for CTPCDSDTCEntries {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        if let Some(ref val) = self.count {
            {
                let s = val.to_string();
                start.push_attribute(("count", s.as_str()));
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

impl ToXml for CTTuples {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        if let Some(ref val) = self.cells {
            {
                let s = val.to_string();
                start.push_attribute(("c", s.as_str()));
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
        for item in &self.tpl {
            #[cfg(feature = "extra-children")]
            while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
                extra_iter
                    .next()
                    .unwrap()
                    .node
                    .write_to(writer)
                    .map_err(SerializeError::from)?;
            }
            item.write_element("tpl", writer)?;
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
        if !self.tpl.is_empty() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for CTTuple {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        if let Some(ref val) = self.fld {
            {
                let s = val.to_string();
                start.push_attribute(("fld", s.as_str()));
            }
        }
        if let Some(ref val) = self.hier {
            {
                let s = val.to_string();
                start.push_attribute(("hier", s.as_str()));
            }
        }
        {
            let val = &self.item;
            {
                let s = val.to_string();
                start.push_attribute(("item", s.as_str()));
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

impl ToXml for CTSets {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        if let Some(ref val) = self.count {
            {
                let s = val.to_string();
                start.push_attribute(("count", s.as_str()));
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
        for item in &self.set {
            #[cfg(feature = "extra-children")]
            while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
                extra_iter
                    .next()
                    .unwrap()
                    .node
                    .write_to(writer)
                    .map_err(SerializeError::from)?;
            }
            item.write_element("set", writer)?;
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
        if !self.set.is_empty() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for CTSet {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        if let Some(ref val) = self.count {
            {
                let s = val.to_string();
                start.push_attribute(("count", s.as_str()));
            }
        }
        {
            let val = &self.max_rank;
            {
                let s = val.to_string();
                start.push_attribute(("maxRank", s.as_str()));
            }
        }
        {
            let val = &self.set_definition;
            start.push_attribute(("setDefinition", val.as_str()));
        }
        if let Some(ref val) = self.sort_type {
            {
                let s = val.to_string();
                start.push_attribute(("sortType", s.as_str()));
            }
        }
        if let Some(ref val) = self.query_failed {
            start.push_attribute(("queryFailed", if *val { "1" } else { "0" }));
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
        for item in &self.tpls {
            #[cfg(feature = "extra-children")]
            while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
                extra_iter
                    .next()
                    .unwrap()
                    .node
                    .write_to(writer)
                    .map_err(SerializeError::from)?;
            }
            item.write_element("tpls", writer)?;
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
        if let Some(ref val) = self.sort_by_tuple {
            val.write_element("sortByTuple", writer)?;
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
        if !self.tpls.is_empty() {
            return false;
        }
        if self.sort_by_tuple.is_some() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for CTQueryCache {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        if let Some(ref val) = self.count {
            {
                let s = val.to_string();
                start.push_attribute(("count", s.as_str()));
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
        for item in &self.query {
            #[cfg(feature = "extra-children")]
            while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
                extra_iter
                    .next()
                    .unwrap()
                    .node
                    .write_to(writer)
                    .map_err(SerializeError::from)?;
            }
            item.write_element("query", writer)?;
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
        if !self.query.is_empty() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for CTQuery {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        {
            let val = &self.mdx;
            start.push_attribute(("mdx", val.as_str()));
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
        if let Some(ref val) = self.tpls {
            val.write_element("tpls", writer)?;
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
        if self.tpls.is_some() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for CTCalculatedItems {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        if let Some(ref val) = self.count {
            {
                let s = val.to_string();
                start.push_attribute(("count", s.as_str()));
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
        for item in &self.calculated_item {
            #[cfg(feature = "extra-children")]
            while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
                extra_iter
                    .next()
                    .unwrap()
                    .node
                    .write_to(writer)
                    .map_err(SerializeError::from)?;
            }
            item.write_element("calculatedItem", writer)?;
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
        if !self.calculated_item.is_empty() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for CTCalculatedItem {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        if let Some(ref val) = self.field {
            {
                let s = val.to_string();
                start.push_attribute(("field", s.as_str()));
            }
        }
        if let Some(ref val) = self.formula {
            start.push_attribute(("formula", val.as_str()));
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
            let val = &self.pivot_area;
            val.write_element("pivotArea", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        if let Some(ref val) = self.extension_list {
            val.write_element("extLst", writer)?;
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

impl ToXml for CTCalculatedMembers {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        if let Some(ref val) = self.count {
            {
                let s = val.to_string();
                start.push_attribute(("count", s.as_str()));
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
        for item in &self.calculated_member {
            #[cfg(feature = "extra-children")]
            while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
                extra_iter
                    .next()
                    .unwrap()
                    .node
                    .write_to(writer)
                    .map_err(SerializeError::from)?;
            }
            item.write_element("calculatedMember", writer)?;
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
        if !self.calculated_member.is_empty() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for CTCalculatedMember {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        {
            let val = &self.name;
            start.push_attribute(("name", val.as_str()));
        }
        {
            let val = &self.mdx;
            start.push_attribute(("mdx", val.as_str()));
        }
        if let Some(ref val) = self.member_name {
            start.push_attribute(("memberName", val.as_str()));
        }
        if let Some(ref val) = self.hierarchy {
            start.push_attribute(("hierarchy", val.as_str()));
        }
        if let Some(ref val) = self.parent {
            start.push_attribute(("parent", val.as_str()));
        }
        if let Some(ref val) = self.solve_order {
            {
                let s = val.to_string();
                start.push_attribute(("solveOrder", s.as_str()));
            }
        }
        if let Some(ref val) = self.set {
            start.push_attribute(("set", if *val { "1" } else { "0" }));
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
        if let Some(ref val) = self.extension_list {
            val.write_element("extLst", writer)?;
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
        if self.extension_list.is_some() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for CTPivotTableDefinition {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        {
            let val = &self.name;
            start.push_attribute(("name", val.as_str()));
        }
        {
            let val = &self.cache_id;
            {
                let s = val.to_string();
                start.push_attribute(("cacheId", s.as_str()));
            }
        }
        if let Some(ref val) = self.data_on_rows {
            start.push_attribute(("dataOnRows", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.data_position {
            {
                let s = val.to_string();
                start.push_attribute(("dataPosition", s.as_str()));
            }
        }
        if let Some(ref val) = self.auto_format_id {
            {
                let s = val.to_string();
                start.push_attribute(("autoFormatId", s.as_str()));
            }
        }
        if let Some(ref val) = self.apply_number_formats {
            start.push_attribute(("applyNumberFormats", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.apply_border_formats {
            start.push_attribute(("applyBorderFormats", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.apply_font_formats {
            start.push_attribute(("applyFontFormats", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.apply_pattern_formats {
            start.push_attribute(("applyPatternFormats", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.apply_alignment_formats {
            start.push_attribute(("applyAlignmentFormats", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.apply_width_height_formats {
            start.push_attribute(("applyWidthHeightFormats", if *val { "1" } else { "0" }));
        }
        {
            let val = &self.data_caption;
            start.push_attribute(("dataCaption", val.as_str()));
        }
        if let Some(ref val) = self.grand_total_caption {
            start.push_attribute(("grandTotalCaption", val.as_str()));
        }
        if let Some(ref val) = self.error_caption {
            start.push_attribute(("errorCaption", val.as_str()));
        }
        if let Some(ref val) = self.show_error {
            start.push_attribute(("showError", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.missing_caption {
            start.push_attribute(("missingCaption", val.as_str()));
        }
        if let Some(ref val) = self.show_missing {
            start.push_attribute(("showMissing", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.page_style {
            start.push_attribute(("pageStyle", val.as_str()));
        }
        if let Some(ref val) = self.pivot_table_style {
            start.push_attribute(("pivotTableStyle", val.as_str()));
        }
        if let Some(ref val) = self.vacated_style {
            start.push_attribute(("vacatedStyle", val.as_str()));
        }
        if let Some(ref val) = self.tag {
            start.push_attribute(("tag", val.as_str()));
        }
        if let Some(ref val) = self.updated_version {
            {
                let s = val.to_string();
                start.push_attribute(("updatedVersion", s.as_str()));
            }
        }
        if let Some(ref val) = self.min_refreshable_version {
            {
                let s = val.to_string();
                start.push_attribute(("minRefreshableVersion", s.as_str()));
            }
        }
        if let Some(ref val) = self.asterisk_totals {
            start.push_attribute(("asteriskTotals", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.show_items {
            start.push_attribute(("showItems", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.edit_data {
            start.push_attribute(("editData", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.disable_field_list {
            start.push_attribute(("disableFieldList", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.show_calc_mbrs {
            start.push_attribute(("showCalcMbrs", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.visual_totals {
            start.push_attribute(("visualTotals", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.show_multiple_label {
            start.push_attribute(("showMultipleLabel", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.show_data_drop_down {
            start.push_attribute(("showDataDropDown", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.show_drill {
            start.push_attribute(("showDrill", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.print_drill {
            start.push_attribute(("printDrill", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.show_member_property_tips {
            start.push_attribute(("showMemberPropertyTips", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.show_data_tips {
            start.push_attribute(("showDataTips", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.enable_wizard {
            start.push_attribute(("enableWizard", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.enable_drill {
            start.push_attribute(("enableDrill", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.enable_field_properties {
            start.push_attribute(("enableFieldProperties", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.preserve_formatting {
            start.push_attribute(("preserveFormatting", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.use_auto_formatting {
            start.push_attribute(("useAutoFormatting", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.page_wrap {
            {
                let s = val.to_string();
                start.push_attribute(("pageWrap", s.as_str()));
            }
        }
        if let Some(ref val) = self.page_over_then_down {
            start.push_attribute(("pageOverThenDown", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.subtotal_hidden_items {
            start.push_attribute(("subtotalHiddenItems", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.row_grand_totals {
            start.push_attribute(("rowGrandTotals", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.col_grand_totals {
            start.push_attribute(("colGrandTotals", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.field_print_titles {
            start.push_attribute(("fieldPrintTitles", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.item_print_titles {
            start.push_attribute(("itemPrintTitles", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.merge_item {
            start.push_attribute(("mergeItem", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.show_drop_zones {
            start.push_attribute(("showDropZones", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.created_version {
            {
                let s = val.to_string();
                start.push_attribute(("createdVersion", s.as_str()));
            }
        }
        if let Some(ref val) = self.indent {
            {
                let s = val.to_string();
                start.push_attribute(("indent", s.as_str()));
            }
        }
        if let Some(ref val) = self.show_empty_row {
            start.push_attribute(("showEmptyRow", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.show_empty_col {
            start.push_attribute(("showEmptyCol", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.show_headers {
            start.push_attribute(("showHeaders", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.compact {
            start.push_attribute(("compact", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.outline {
            start.push_attribute(("outline", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.outline_data {
            start.push_attribute(("outlineData", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.compact_data {
            start.push_attribute(("compactData", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.published {
            start.push_attribute(("published", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.grid_drop_zones {
            start.push_attribute(("gridDropZones", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.immersive {
            start.push_attribute(("immersive", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.multiple_field_filters {
            start.push_attribute(("multipleFieldFilters", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.chart_format {
            {
                let s = val.to_string();
                start.push_attribute(("chartFormat", s.as_str()));
            }
        }
        if let Some(ref val) = self.row_header_caption {
            start.push_attribute(("rowHeaderCaption", val.as_str()));
        }
        if let Some(ref val) = self.col_header_caption {
            start.push_attribute(("colHeaderCaption", val.as_str()));
        }
        if let Some(ref val) = self.field_list_sort_ascending {
            start.push_attribute(("fieldListSortAscending", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.mdx_subqueries {
            start.push_attribute(("mdxSubqueries", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.custom_list_sort {
            start.push_attribute(("customListSort", if *val { "1" } else { "0" }));
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
            let val = &self.location;
            val.write_element("location", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        if let Some(ref val) = self.pivot_fields {
            val.write_element("pivotFields", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        if let Some(ref val) = self.row_fields {
            val.write_element("rowFields", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        if let Some(ref val) = self.row_items {
            val.write_element("rowItems", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        if let Some(ref val) = self.col_fields {
            val.write_element("colFields", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        if let Some(ref val) = self.col_items {
            val.write_element("colItems", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        if let Some(ref val) = self.page_fields {
            val.write_element("pageFields", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        if let Some(ref val) = self.data_fields {
            val.write_element("dataFields", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        if let Some(ref val) = self.formats {
            val.write_element("formats", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        if let Some(ref val) = self.conditional_formats {
            val.write_element("conditionalFormats", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        if let Some(ref val) = self.chart_formats {
            val.write_element("chartFormats", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        if let Some(ref val) = self.pivot_hierarchies {
            val.write_element("pivotHierarchies", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        if let Some(ref val) = self.pivot_table_style_info {
            val.write_element("pivotTableStyleInfo", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        if let Some(ref val) = self.filters {
            val.write_element("filters", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        if let Some(ref val) = self.row_hierarchies_usage {
            val.write_element("rowHierarchiesUsage", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        if let Some(ref val) = self.col_hierarchies_usage {
            val.write_element("colHierarchiesUsage", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        if let Some(ref val) = self.extension_list {
            val.write_element("extLst", writer)?;
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

impl ToXml for PivotLocation {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        {
            let val = &self.reference;
            start.push_attribute(("ref", val.as_str()));
        }
        {
            let val = &self.first_header_row;
            {
                let s = val.to_string();
                start.push_attribute(("firstHeaderRow", s.as_str()));
            }
        }
        {
            let val = &self.first_data_row;
            {
                let s = val.to_string();
                start.push_attribute(("firstDataRow", s.as_str()));
            }
        }
        {
            let val = &self.first_data_col;
            {
                let s = val.to_string();
                start.push_attribute(("firstDataCol", s.as_str()));
            }
        }
        if let Some(ref val) = self.row_page_count {
            {
                let s = val.to_string();
                start.push_attribute(("rowPageCount", s.as_str()));
            }
        }
        if let Some(ref val) = self.col_page_count {
            {
                let s = val.to_string();
                start.push_attribute(("colPageCount", s.as_str()));
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

impl ToXml for PivotFields {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        if let Some(ref val) = self.count {
            {
                let s = val.to_string();
                start.push_attribute(("count", s.as_str()));
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
        for item in &self.pivot_field {
            #[cfg(feature = "extra-children")]
            while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
                extra_iter
                    .next()
                    .unwrap()
                    .node
                    .write_to(writer)
                    .map_err(SerializeError::from)?;
            }
            item.write_element("pivotField", writer)?;
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
        if !self.pivot_field.is_empty() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for PivotField {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        if let Some(ref val) = self.name {
            start.push_attribute(("name", val.as_str()));
        }
        if let Some(ref val) = self.axis {
            {
                let s = val.to_string();
                start.push_attribute(("axis", s.as_str()));
            }
        }
        if let Some(ref val) = self.data_field {
            start.push_attribute(("dataField", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.subtotal_caption {
            start.push_attribute(("subtotalCaption", val.as_str()));
        }
        if let Some(ref val) = self.show_drop_downs {
            start.push_attribute(("showDropDowns", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.hidden_level {
            start.push_attribute(("hiddenLevel", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.unique_member_property {
            start.push_attribute(("uniqueMemberProperty", val.as_str()));
        }
        if let Some(ref val) = self.compact {
            start.push_attribute(("compact", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.all_drilled {
            start.push_attribute(("allDrilled", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.number_format_id {
            {
                let s = val.to_string();
                start.push_attribute(("numFmtId", s.as_str()));
            }
        }
        if let Some(ref val) = self.outline {
            start.push_attribute(("outline", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.subtotal_top {
            start.push_attribute(("subtotalTop", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.drag_to_row {
            start.push_attribute(("dragToRow", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.drag_to_col {
            start.push_attribute(("dragToCol", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.multiple_item_selection_allowed {
            start.push_attribute(("multipleItemSelectionAllowed", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.drag_to_page {
            start.push_attribute(("dragToPage", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.drag_to_data {
            start.push_attribute(("dragToData", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.drag_off {
            start.push_attribute(("dragOff", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.show_all {
            start.push_attribute(("showAll", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.insert_blank_row {
            start.push_attribute(("insertBlankRow", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.server_field {
            start.push_attribute(("serverField", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.insert_page_break {
            start.push_attribute(("insertPageBreak", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.auto_show {
            start.push_attribute(("autoShow", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.top_auto_show {
            start.push_attribute(("topAutoShow", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.hide_new_items {
            start.push_attribute(("hideNewItems", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.measure_filter {
            start.push_attribute(("measureFilter", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.include_new_items_in_filter {
            start.push_attribute(("includeNewItemsInFilter", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.item_page_count {
            {
                let s = val.to_string();
                start.push_attribute(("itemPageCount", s.as_str()));
            }
        }
        if let Some(ref val) = self.sort_type {
            {
                let s = val.to_string();
                start.push_attribute(("sortType", s.as_str()));
            }
        }
        if let Some(ref val) = self.data_source_sort {
            start.push_attribute(("dataSourceSort", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.non_auto_sort_default {
            start.push_attribute(("nonAutoSortDefault", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.rank_by {
            {
                let s = val.to_string();
                start.push_attribute(("rankBy", s.as_str()));
            }
        }
        if let Some(ref val) = self.default_subtotal {
            start.push_attribute(("defaultSubtotal", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.sum_subtotal {
            start.push_attribute(("sumSubtotal", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.count_a_subtotal {
            start.push_attribute(("countASubtotal", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.avg_subtotal {
            start.push_attribute(("avgSubtotal", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.max_subtotal {
            start.push_attribute(("maxSubtotal", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.min_subtotal {
            start.push_attribute(("minSubtotal", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.product_subtotal {
            start.push_attribute(("productSubtotal", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.count_subtotal {
            start.push_attribute(("countSubtotal", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.std_dev_subtotal {
            start.push_attribute(("stdDevSubtotal", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.std_dev_p_subtotal {
            start.push_attribute(("stdDevPSubtotal", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.var_subtotal {
            start.push_attribute(("varSubtotal", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.var_p_subtotal {
            start.push_attribute(("varPSubtotal", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.show_prop_cell {
            start.push_attribute(("showPropCell", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.show_prop_tip {
            start.push_attribute(("showPropTip", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.show_prop_as_caption {
            start.push_attribute(("showPropAsCaption", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.default_attribute_drill_state {
            start.push_attribute(("defaultAttributeDrillState", if *val { "1" } else { "0" }));
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
        if let Some(ref val) = self.items {
            val.write_element("items", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        if let Some(ref val) = self.auto_sort_scope {
            val.write_element("autoSortScope", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        if let Some(ref val) = self.extension_list {
            val.write_element("extLst", writer)?;
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
        if self.items.is_some() {
            return false;
        }
        if self.auto_sort_scope.is_some() {
            return false;
        }
        if self.extension_list.is_some() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for PivotItems {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        if let Some(ref val) = self.count {
            {
                let s = val.to_string();
                start.push_attribute(("count", s.as_str()));
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
        for item in &self.item {
            #[cfg(feature = "extra-children")]
            while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
                extra_iter
                    .next()
                    .unwrap()
                    .node
                    .write_to(writer)
                    .map_err(SerializeError::from)?;
            }
            item.write_element("item", writer)?;
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
        if !self.item.is_empty() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for PivotItem {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        if let Some(ref val) = self.n {
            start.push_attribute(("n", val.as_str()));
        }
        if let Some(ref val) = self.cell_type {
            {
                let s = val.to_string();
                start.push_attribute(("t", s.as_str()));
            }
        }
        if let Some(ref val) = self.height {
            start.push_attribute(("h", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.style_index {
            start.push_attribute(("s", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.sd {
            start.push_attribute(("sd", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.formula {
            start.push_attribute(("f", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.m {
            start.push_attribute(("m", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.cells {
            start.push_attribute(("c", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.x {
            {
                let s = val.to_string();
                start.push_attribute(("x", s.as_str()));
            }
        }
        if let Some(ref val) = self.d {
            start.push_attribute(("d", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.e {
            start.push_attribute(("e", if *val { "1" } else { "0" }));
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

impl ToXml for PageFields {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        if let Some(ref val) = self.count {
            {
                let s = val.to_string();
                start.push_attribute(("count", s.as_str()));
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
        for item in &self.page_field {
            #[cfg(feature = "extra-children")]
            while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
                extra_iter
                    .next()
                    .unwrap()
                    .node
                    .write_to(writer)
                    .map_err(SerializeError::from)?;
            }
            item.write_element("pageField", writer)?;
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
        if !self.page_field.is_empty() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for PageField {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        {
            let val = &self.fld;
            {
                let s = val.to_string();
                start.push_attribute(("fld", s.as_str()));
            }
        }
        if let Some(ref val) = self.item {
            {
                let s = val.to_string();
                start.push_attribute(("item", s.as_str()));
            }
        }
        if let Some(ref val) = self.hier {
            {
                let s = val.to_string();
                start.push_attribute(("hier", s.as_str()));
            }
        }
        if let Some(ref val) = self.name {
            start.push_attribute(("name", val.as_str()));
        }
        if let Some(ref val) = self.cap {
            start.push_attribute(("cap", val.as_str()));
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
        if let Some(ref val) = self.extension_list {
            val.write_element("extLst", writer)?;
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
        if self.extension_list.is_some() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for DataFields {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        if let Some(ref val) = self.count {
            {
                let s = val.to_string();
                start.push_attribute(("count", s.as_str()));
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
        for item in &self.data_field {
            #[cfg(feature = "extra-children")]
            while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
                extra_iter
                    .next()
                    .unwrap()
                    .node
                    .write_to(writer)
                    .map_err(SerializeError::from)?;
            }
            item.write_element("dataField", writer)?;
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
        if !self.data_field.is_empty() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for DataField {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        if let Some(ref val) = self.name {
            start.push_attribute(("name", val.as_str()));
        }
        {
            let val = &self.fld;
            {
                let s = val.to_string();
                start.push_attribute(("fld", s.as_str()));
            }
        }
        if let Some(ref val) = self.subtotal {
            {
                let s = val.to_string();
                start.push_attribute(("subtotal", s.as_str()));
            }
        }
        if let Some(ref val) = self.show_data_as {
            {
                let s = val.to_string();
                start.push_attribute(("showDataAs", s.as_str()));
            }
        }
        if let Some(ref val) = self.base_field {
            {
                let s = val.to_string();
                start.push_attribute(("baseField", s.as_str()));
            }
        }
        if let Some(ref val) = self.base_item {
            {
                let s = val.to_string();
                start.push_attribute(("baseItem", s.as_str()));
            }
        }
        if let Some(ref val) = self.number_format_id {
            {
                let s = val.to_string();
                start.push_attribute(("numFmtId", s.as_str()));
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
        if let Some(ref val) = self.extension_list {
            val.write_element("extLst", writer)?;
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
        if self.extension_list.is_some() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for CTRowItems {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        if let Some(ref val) = self.count {
            {
                let s = val.to_string();
                start.push_attribute(("count", s.as_str()));
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
        for item in &self.i {
            #[cfg(feature = "extra-children")]
            while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
                extra_iter
                    .next()
                    .unwrap()
                    .node
                    .write_to(writer)
                    .map_err(SerializeError::from)?;
            }
            item.write_element("i", writer)?;
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
        if !self.i.is_empty() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for CTColItems {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        if let Some(ref val) = self.count {
            {
                let s = val.to_string();
                start.push_attribute(("count", s.as_str()));
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
        for item in &self.i {
            #[cfg(feature = "extra-children")]
            while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
                extra_iter
                    .next()
                    .unwrap()
                    .node
                    .write_to(writer)
                    .map_err(SerializeError::from)?;
            }
            item.write_element("i", writer)?;
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
        if !self.i.is_empty() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for CTI {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        if let Some(ref val) = self.cell_type {
            {
                let s = val.to_string();
                start.push_attribute(("t", s.as_str()));
            }
        }
        if let Some(ref val) = self.reference {
            {
                let s = val.to_string();
                start.push_attribute(("r", s.as_str()));
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
        for item in &self.x {
            #[cfg(feature = "extra-children")]
            while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
                extra_iter
                    .next()
                    .unwrap()
                    .node
                    .write_to(writer)
                    .map_err(SerializeError::from)?;
            }
            item.write_element("x", writer)?;
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
        if !self.x.is_empty() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for CTX {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        if let Some(ref val) = self.value {
            {
                let s = val.to_string();
                start.push_attribute(("v", s.as_str()));
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

impl ToXml for RowFields {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        if let Some(ref val) = self.count {
            {
                let s = val.to_string();
                start.push_attribute(("count", s.as_str()));
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
        for item in &self.field {
            #[cfg(feature = "extra-children")]
            while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
                extra_iter
                    .next()
                    .unwrap()
                    .node
                    .write_to(writer)
                    .map_err(SerializeError::from)?;
            }
            item.write_element("field", writer)?;
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
        if !self.field.is_empty() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for ColFields {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        if let Some(ref val) = self.count {
            {
                let s = val.to_string();
                start.push_attribute(("count", s.as_str()));
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
        for item in &self.field {
            #[cfg(feature = "extra-children")]
            while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
                extra_iter
                    .next()
                    .unwrap()
                    .node
                    .write_to(writer)
                    .map_err(SerializeError::from)?;
            }
            item.write_element("field", writer)?;
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
        if !self.field.is_empty() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for CTField {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        {
            let val = &self.x;
            {
                let s = val.to_string();
                start.push_attribute(("x", s.as_str()));
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

impl ToXml for CTFormats {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        if let Some(ref val) = self.count {
            {
                let s = val.to_string();
                start.push_attribute(("count", s.as_str()));
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
        for item in &self.format {
            #[cfg(feature = "extra-children")]
            while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
                extra_iter
                    .next()
                    .unwrap()
                    .node
                    .write_to(writer)
                    .map_err(SerializeError::from)?;
            }
            item.write_element("format", writer)?;
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
        if !self.format.is_empty() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for CTFormat {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        if let Some(ref val) = self.action {
            {
                let s = val.to_string();
                start.push_attribute(("action", s.as_str()));
            }
        }
        if let Some(ref val) = self.dxf_id {
            {
                let s = val.to_string();
                start.push_attribute(("dxfId", s.as_str()));
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
            let val = &self.pivot_area;
            val.write_element("pivotArea", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        if let Some(ref val) = self.extension_list {
            val.write_element("extLst", writer)?;
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

impl ToXml for CTConditionalFormats {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        if let Some(ref val) = self.count {
            {
                let s = val.to_string();
                start.push_attribute(("count", s.as_str()));
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
        for item in &self.conditional_format {
            #[cfg(feature = "extra-children")]
            while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
                extra_iter
                    .next()
                    .unwrap()
                    .node
                    .write_to(writer)
                    .map_err(SerializeError::from)?;
            }
            item.write_element("conditionalFormat", writer)?;
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
        if !self.conditional_format.is_empty() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for CTConditionalFormat {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        if let Some(ref val) = self.scope {
            {
                let s = val.to_string();
                start.push_attribute(("scope", s.as_str()));
            }
        }
        if let Some(ref val) = self.r#type {
            {
                let s = val.to_string();
                start.push_attribute(("type", s.as_str()));
            }
        }
        {
            let val = &self.priority;
            {
                let s = val.to_string();
                start.push_attribute(("priority", s.as_str()));
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
            let val = &self.pivot_areas;
            val.write_element("pivotAreas", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        if let Some(ref val) = self.extension_list {
            val.write_element("extLst", writer)?;
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

impl ToXml for PivotAreas {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        if let Some(ref val) = self.count {
            {
                let s = val.to_string();
                start.push_attribute(("count", s.as_str()));
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
        for item in &self.pivot_area {
            #[cfg(feature = "extra-children")]
            while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
                extra_iter
                    .next()
                    .unwrap()
                    .node
                    .write_to(writer)
                    .map_err(SerializeError::from)?;
            }
            item.write_element("pivotArea", writer)?;
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
        if !self.pivot_area.is_empty() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for CTChartFormats {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        if let Some(ref val) = self.count {
            {
                let s = val.to_string();
                start.push_attribute(("count", s.as_str()));
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
        for item in &self.chart_format {
            #[cfg(feature = "extra-children")]
            while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
                extra_iter
                    .next()
                    .unwrap()
                    .node
                    .write_to(writer)
                    .map_err(SerializeError::from)?;
            }
            item.write_element("chartFormat", writer)?;
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
        if !self.chart_format.is_empty() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for CTChartFormat {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        {
            let val = &self.chart;
            {
                let s = val.to_string();
                start.push_attribute(("chart", s.as_str()));
            }
        }
        {
            let val = &self.format;
            {
                let s = val.to_string();
                start.push_attribute(("format", s.as_str()));
            }
        }
        if let Some(ref val) = self.series {
            start.push_attribute(("series", if *val { "1" } else { "0" }));
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
            let val = &self.pivot_area;
            val.write_element("pivotArea", writer)?;
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

impl ToXml for CTPivotHierarchies {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        if let Some(ref val) = self.count {
            {
                let s = val.to_string();
                start.push_attribute(("count", s.as_str()));
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
        for item in &self.pivot_hierarchy {
            #[cfg(feature = "extra-children")]
            while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
                extra_iter
                    .next()
                    .unwrap()
                    .node
                    .write_to(writer)
                    .map_err(SerializeError::from)?;
            }
            item.write_element("pivotHierarchy", writer)?;
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
        if !self.pivot_hierarchy.is_empty() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for CTPivotHierarchy {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        if let Some(ref val) = self.outline {
            start.push_attribute(("outline", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.multiple_item_selection_allowed {
            start.push_attribute(("multipleItemSelectionAllowed", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.subtotal_top {
            start.push_attribute(("subtotalTop", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.show_in_field_list {
            start.push_attribute(("showInFieldList", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.drag_to_row {
            start.push_attribute(("dragToRow", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.drag_to_col {
            start.push_attribute(("dragToCol", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.drag_to_page {
            start.push_attribute(("dragToPage", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.drag_to_data {
            start.push_attribute(("dragToData", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.drag_off {
            start.push_attribute(("dragOff", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.include_new_items_in_filter {
            start.push_attribute(("includeNewItemsInFilter", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.caption {
            start.push_attribute(("caption", val.as_str()));
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
        if let Some(ref val) = self.mps {
            val.write_element("mps", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        for item in &self.members {
            #[cfg(feature = "extra-children")]
            while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
                extra_iter
                    .next()
                    .unwrap()
                    .node
                    .write_to(writer)
                    .map_err(SerializeError::from)?;
            }
            item.write_element("members", writer)?;
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
        if let Some(ref val) = self.extension_list {
            val.write_element("extLst", writer)?;
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
        if self.mps.is_some() {
            return false;
        }
        if !self.members.is_empty() {
            return false;
        }
        if self.extension_list.is_some() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for CTRowHierarchiesUsage {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        if let Some(ref val) = self.count {
            {
                let s = val.to_string();
                start.push_attribute(("count", s.as_str()));
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
        for item in &self.row_hierarchy_usage {
            #[cfg(feature = "extra-children")]
            while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
                extra_iter
                    .next()
                    .unwrap()
                    .node
                    .write_to(writer)
                    .map_err(SerializeError::from)?;
            }
            item.write_element("rowHierarchyUsage", writer)?;
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
        if !self.row_hierarchy_usage.is_empty() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for CTColHierarchiesUsage {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        if let Some(ref val) = self.count {
            {
                let s = val.to_string();
                start.push_attribute(("count", s.as_str()));
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
        for item in &self.col_hierarchy_usage {
            #[cfg(feature = "extra-children")]
            while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
                extra_iter
                    .next()
                    .unwrap()
                    .node
                    .write_to(writer)
                    .map_err(SerializeError::from)?;
            }
            item.write_element("colHierarchyUsage", writer)?;
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
        if !self.col_hierarchy_usage.is_empty() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for CTHierarchyUsage {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        {
            let val = &self.hierarchy_usage;
            {
                let s = val.to_string();
                start.push_attribute(("hierarchyUsage", s.as_str()));
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

impl ToXml for CTMemberProperties {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        if let Some(ref val) = self.count {
            {
                let s = val.to_string();
                start.push_attribute(("count", s.as_str()));
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
        for item in &self.mp {
            #[cfg(feature = "extra-children")]
            while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
                extra_iter
                    .next()
                    .unwrap()
                    .node
                    .write_to(writer)
                    .map_err(SerializeError::from)?;
            }
            item.write_element("mp", writer)?;
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
        if !self.mp.is_empty() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for CTMemberProperty {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        if let Some(ref val) = self.name {
            start.push_attribute(("name", val.as_str()));
        }
        if let Some(ref val) = self.show_cell {
            start.push_attribute(("showCell", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.show_tip {
            start.push_attribute(("showTip", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.show_as_caption {
            start.push_attribute(("showAsCaption", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.name_len {
            {
                let s = val.to_string();
                start.push_attribute(("nameLen", s.as_str()));
            }
        }
        if let Some(ref val) = self.p_pos {
            {
                let s = val.to_string();
                start.push_attribute(("pPos", s.as_str()));
            }
        }
        if let Some(ref val) = self.p_len {
            {
                let s = val.to_string();
                start.push_attribute(("pLen", s.as_str()));
            }
        }
        if let Some(ref val) = self.level {
            {
                let s = val.to_string();
                start.push_attribute(("level", s.as_str()));
            }
        }
        {
            let val = &self.field;
            {
                let s = val.to_string();
                start.push_attribute(("field", s.as_str()));
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

impl ToXml for CTMembers {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        if let Some(ref val) = self.count {
            {
                let s = val.to_string();
                start.push_attribute(("count", s.as_str()));
            }
        }
        if let Some(ref val) = self.level {
            {
                let s = val.to_string();
                start.push_attribute(("level", s.as_str()));
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
        for item in &self.member {
            #[cfg(feature = "extra-children")]
            while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
                extra_iter
                    .next()
                    .unwrap()
                    .node
                    .write_to(writer)
                    .map_err(SerializeError::from)?;
            }
            item.write_element("member", writer)?;
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
        if !self.member.is_empty() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for CTMember {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
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

    fn is_empty_element(&self) -> bool {
        true
    }
}

impl ToXml for CTDimensions {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        if let Some(ref val) = self.count {
            {
                let s = val.to_string();
                start.push_attribute(("count", s.as_str()));
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
        for item in &self.dimension {
            #[cfg(feature = "extra-children")]
            while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
                extra_iter
                    .next()
                    .unwrap()
                    .node
                    .write_to(writer)
                    .map_err(SerializeError::from)?;
            }
            item.write_element("dimension", writer)?;
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
        if !self.dimension.is_empty() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for CTPivotDimension {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        if let Some(ref val) = self.measure {
            start.push_attribute(("measure", if *val { "1" } else { "0" }));
        }
        {
            let val = &self.name;
            start.push_attribute(("name", val.as_str()));
        }
        {
            let val = &self.unique_name;
            start.push_attribute(("uniqueName", val.as_str()));
        }
        {
            let val = &self.caption;
            start.push_attribute(("caption", val.as_str()));
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

impl ToXml for CTMeasureGroups {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        if let Some(ref val) = self.count {
            {
                let s = val.to_string();
                start.push_attribute(("count", s.as_str()));
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
        for item in &self.measure_group {
            #[cfg(feature = "extra-children")]
            while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
                extra_iter
                    .next()
                    .unwrap()
                    .node
                    .write_to(writer)
                    .map_err(SerializeError::from)?;
            }
            item.write_element("measureGroup", writer)?;
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
        if !self.measure_group.is_empty() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for CTMeasureDimensionMaps {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        if let Some(ref val) = self.count {
            {
                let s = val.to_string();
                start.push_attribute(("count", s.as_str()));
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
        for item in &self.map {
            #[cfg(feature = "extra-children")]
            while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
                extra_iter
                    .next()
                    .unwrap()
                    .node
                    .write_to(writer)
                    .map_err(SerializeError::from)?;
            }
            item.write_element("map", writer)?;
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
        if !self.map.is_empty() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for CTMeasureGroup {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        {
            let val = &self.name;
            start.push_attribute(("name", val.as_str()));
        }
        {
            let val = &self.caption;
            start.push_attribute(("caption", val.as_str()));
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

impl ToXml for CTMeasureDimensionMap {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        if let Some(ref val) = self.measure_group {
            {
                let s = val.to_string();
                start.push_attribute(("measureGroup", s.as_str()));
            }
        }
        if let Some(ref val) = self.dimension {
            {
                let s = val.to_string();
                start.push_attribute(("dimension", s.as_str()));
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

impl ToXml for CTPivotTableStyle {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        if let Some(ref val) = self.name {
            start.push_attribute(("name", val.as_str()));
        }
        if let Some(ref val) = self.show_row_headers {
            start.push_attribute(("showRowHeaders", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.show_col_headers {
            start.push_attribute(("showColHeaders", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.show_row_stripes {
            start.push_attribute(("showRowStripes", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.show_col_stripes {
            start.push_attribute(("showColStripes", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.show_last_column {
            start.push_attribute(("showLastColumn", if *val { "1" } else { "0" }));
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

impl ToXml for PivotFilters {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        if let Some(ref val) = self.count {
            {
                let s = val.to_string();
                start.push_attribute(("count", s.as_str()));
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
        for item in &self.filter {
            #[cfg(feature = "extra-children")]
            while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
                extra_iter
                    .next()
                    .unwrap()
                    .node
                    .write_to(writer)
                    .map_err(SerializeError::from)?;
            }
            item.write_element("filter", writer)?;
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
        if !self.filter.is_empty() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for PivotFilter {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        {
            let val = &self.fld;
            {
                let s = val.to_string();
                start.push_attribute(("fld", s.as_str()));
            }
        }
        if let Some(ref val) = self.mp_fld {
            {
                let s = val.to_string();
                start.push_attribute(("mpFld", s.as_str()));
            }
        }
        {
            let val = &self.r#type;
            {
                let s = val.to_string();
                start.push_attribute(("type", s.as_str()));
            }
        }
        if let Some(ref val) = self.eval_order {
            {
                let s = val.to_string();
                start.push_attribute(("evalOrder", s.as_str()));
            }
        }
        {
            let val = &self.id;
            {
                let s = val.to_string();
                start.push_attribute(("id", s.as_str()));
            }
        }
        if let Some(ref val) = self.i_measure_hier {
            {
                let s = val.to_string();
                start.push_attribute(("iMeasureHier", s.as_str()));
            }
        }
        if let Some(ref val) = self.i_measure_fld {
            {
                let s = val.to_string();
                start.push_attribute(("iMeasureFld", s.as_str()));
            }
        }
        if let Some(ref val) = self.name {
            start.push_attribute(("name", val.as_str()));
        }
        if let Some(ref val) = self.description {
            start.push_attribute(("description", val.as_str()));
        }
        if let Some(ref val) = self.string_value1 {
            start.push_attribute(("stringValue1", val.as_str()));
        }
        if let Some(ref val) = self.string_value2 {
            start.push_attribute(("stringValue2", val.as_str()));
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
            let val = &self.auto_filter;
            val.write_element("autoFilter", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        if let Some(ref val) = self.extension_list {
            val.write_element("extLst", writer)?;
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

impl ToXml for PivotArea {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        if let Some(ref val) = self.field {
            {
                let s = val.to_string();
                start.push_attribute(("field", s.as_str()));
            }
        }
        if let Some(ref val) = self.r#type {
            {
                let s = val.to_string();
                start.push_attribute(("type", s.as_str()));
            }
        }
        if let Some(ref val) = self.data_only {
            start.push_attribute(("dataOnly", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.label_only {
            start.push_attribute(("labelOnly", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.grand_row {
            start.push_attribute(("grandRow", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.grand_col {
            start.push_attribute(("grandCol", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.cache_index {
            start.push_attribute(("cacheIndex", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.outline {
            start.push_attribute(("outline", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.offset {
            start.push_attribute(("offset", val.as_str()));
        }
        if let Some(ref val) = self.collapsed_levels_are_subtotals {
            start.push_attribute(("collapsedLevelsAreSubtotals", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.axis {
            {
                let s = val.to_string();
                start.push_attribute(("axis", s.as_str()));
            }
        }
        if let Some(ref val) = self.field_position {
            {
                let s = val.to_string();
                start.push_attribute(("fieldPosition", s.as_str()));
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
        if let Some(ref val) = self.references {
            val.write_element("references", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        if let Some(ref val) = self.extension_list {
            val.write_element("extLst", writer)?;
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
        if self.references.is_some() {
            return false;
        }
        if self.extension_list.is_some() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for CTPivotAreaReferences {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        if let Some(ref val) = self.count {
            {
                let s = val.to_string();
                start.push_attribute(("count", s.as_str()));
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
        for item in &self.reference {
            #[cfg(feature = "extra-children")]
            while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
                extra_iter
                    .next()
                    .unwrap()
                    .node
                    .write_to(writer)
                    .map_err(SerializeError::from)?;
            }
            item.write_element("reference", writer)?;
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
        if !self.reference.is_empty() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for CTPivotAreaReference {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        if let Some(ref val) = self.field {
            {
                let s = val.to_string();
                start.push_attribute(("field", s.as_str()));
            }
        }
        if let Some(ref val) = self.count {
            {
                let s = val.to_string();
                start.push_attribute(("count", s.as_str()));
            }
        }
        if let Some(ref val) = self.selected {
            start.push_attribute(("selected", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.by_position {
            start.push_attribute(("byPosition", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.relative {
            start.push_attribute(("relative", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.default_subtotal {
            start.push_attribute(("defaultSubtotal", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.sum_subtotal {
            start.push_attribute(("sumSubtotal", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.count_a_subtotal {
            start.push_attribute(("countASubtotal", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.avg_subtotal {
            start.push_attribute(("avgSubtotal", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.max_subtotal {
            start.push_attribute(("maxSubtotal", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.min_subtotal {
            start.push_attribute(("minSubtotal", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.product_subtotal {
            start.push_attribute(("productSubtotal", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.count_subtotal {
            start.push_attribute(("countSubtotal", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.std_dev_subtotal {
            start.push_attribute(("stdDevSubtotal", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.std_dev_p_subtotal {
            start.push_attribute(("stdDevPSubtotal", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.var_subtotal {
            start.push_attribute(("varSubtotal", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.var_p_subtotal {
            start.push_attribute(("varPSubtotal", if *val { "1" } else { "0" }));
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
        for item in &self.x {
            #[cfg(feature = "extra-children")]
            while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
                extra_iter
                    .next()
                    .unwrap()
                    .node
                    .write_to(writer)
                    .map_err(SerializeError::from)?;
            }
            item.write_element("x", writer)?;
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
        if let Some(ref val) = self.extension_list {
            val.write_element("extLst", writer)?;
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
        if !self.x.is_empty() {
            return false;
        }
        if self.extension_list.is_some() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for CTIndex {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        {
            let val = &self.value;
            {
                let s = val.to_string();
                start.push_attribute(("v", s.as_str()));
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

impl ToXml for QueryTable {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        {
            let val = &self.name;
            start.push_attribute(("name", val.as_str()));
        }
        if let Some(ref val) = self.headers {
            start.push_attribute(("headers", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.row_numbers {
            start.push_attribute(("rowNumbers", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.disable_refresh {
            start.push_attribute(("disableRefresh", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.background_refresh {
            start.push_attribute(("backgroundRefresh", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.first_background_refresh {
            start.push_attribute(("firstBackgroundRefresh", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.refresh_on_load {
            start.push_attribute(("refreshOnLoad", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.grow_shrink_type {
            {
                let s = val.to_string();
                start.push_attribute(("growShrinkType", s.as_str()));
            }
        }
        if let Some(ref val) = self.fill_formulas {
            start.push_attribute(("fillFormulas", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.remove_data_on_save {
            start.push_attribute(("removeDataOnSave", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.disable_edit {
            start.push_attribute(("disableEdit", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.preserve_formatting {
            start.push_attribute(("preserveFormatting", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.adjust_column_width {
            start.push_attribute(("adjustColumnWidth", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.intermediate {
            start.push_attribute(("intermediate", if *val { "1" } else { "0" }));
        }
        {
            let val = &self.connection_id;
            {
                let s = val.to_string();
                start.push_attribute(("connectionId", s.as_str()));
            }
        }
        if let Some(ref val) = self.auto_format_id {
            {
                let s = val.to_string();
                start.push_attribute(("autoFormatId", s.as_str()));
            }
        }
        if let Some(ref val) = self.apply_number_formats {
            start.push_attribute(("applyNumberFormats", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.apply_border_formats {
            start.push_attribute(("applyBorderFormats", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.apply_font_formats {
            start.push_attribute(("applyFontFormats", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.apply_pattern_formats {
            start.push_attribute(("applyPatternFormats", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.apply_alignment_formats {
            start.push_attribute(("applyAlignmentFormats", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.apply_width_height_formats {
            start.push_attribute(("applyWidthHeightFormats", if *val { "1" } else { "0" }));
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
        if let Some(ref val) = self.query_table_refresh {
            val.write_element("queryTableRefresh", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        if let Some(ref val) = self.extension_list {
            val.write_element("extLst", writer)?;
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
        if self.query_table_refresh.is_some() {
            return false;
        }
        if self.extension_list.is_some() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for QueryTableRefresh {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        if let Some(ref val) = self.preserve_sort_filter_layout {
            start.push_attribute(("preserveSortFilterLayout", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.field_id_wrapped {
            start.push_attribute(("fieldIdWrapped", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.headers_in_last_refresh {
            start.push_attribute(("headersInLastRefresh", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.minimum_version {
            {
                let s = val.to_string();
                start.push_attribute(("minimumVersion", s.as_str()));
            }
        }
        if let Some(ref val) = self.next_id {
            {
                let s = val.to_string();
                start.push_attribute(("nextId", s.as_str()));
            }
        }
        if let Some(ref val) = self.unbound_columns_left {
            {
                let s = val.to_string();
                start.push_attribute(("unboundColumnsLeft", s.as_str()));
            }
        }
        if let Some(ref val) = self.unbound_columns_right {
            {
                let s = val.to_string();
                start.push_attribute(("unboundColumnsRight", s.as_str()));
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
            let val = &self.query_table_fields;
            val.write_element("queryTableFields", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        if let Some(ref val) = self.query_table_deleted_fields {
            val.write_element("queryTableDeletedFields", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        if let Some(ref val) = self.sort_state {
            val.write_element("sortState", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        if let Some(ref val) = self.extension_list {
            val.write_element("extLst", writer)?;
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

impl ToXml for QueryTableDeletedFields {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        if let Some(ref val) = self.count {
            {
                let s = val.to_string();
                start.push_attribute(("count", s.as_str()));
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
        for item in &self.deleted_field {
            #[cfg(feature = "extra-children")]
            while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
                extra_iter
                    .next()
                    .unwrap()
                    .node
                    .write_to(writer)
                    .map_err(SerializeError::from)?;
            }
            item.write_element("deletedField", writer)?;
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
        if !self.deleted_field.is_empty() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for CTDeletedField {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
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

    fn is_empty_element(&self) -> bool {
        true
    }
}

impl ToXml for QueryTableFields {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        if let Some(ref val) = self.count {
            {
                let s = val.to_string();
                start.push_attribute(("count", s.as_str()));
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
        for item in &self.query_table_field {
            #[cfg(feature = "extra-children")]
            while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
                extra_iter
                    .next()
                    .unwrap()
                    .node
                    .write_to(writer)
                    .map_err(SerializeError::from)?;
            }
            item.write_element("queryTableField", writer)?;
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
        if !self.query_table_field.is_empty() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for QueryTableField {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        {
            let val = &self.id;
            {
                let s = val.to_string();
                start.push_attribute(("id", s.as_str()));
            }
        }
        if let Some(ref val) = self.name {
            start.push_attribute(("name", val.as_str()));
        }
        if let Some(ref val) = self.data_bound {
            start.push_attribute(("dataBound", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.row_numbers {
            start.push_attribute(("rowNumbers", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.fill_formulas {
            start.push_attribute(("fillFormulas", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.clipped {
            start.push_attribute(("clipped", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.table_column_id {
            {
                let s = val.to_string();
                start.push_attribute(("tableColumnId", s.as_str()));
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
        if let Some(ref val) = self.extension_list {
            val.write_element("extLst", writer)?;
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
        if self.extension_list.is_some() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for SharedStrings {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        if let Some(ref val) = self.count {
            {
                let s = val.to_string();
                start.push_attribute(("count", s.as_str()));
            }
        }
        if let Some(ref val) = self.unique_count {
            {
                let s = val.to_string();
                start.push_attribute(("uniqueCount", s.as_str()));
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
        for item in &self.si {
            #[cfg(feature = "extra-children")]
            while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
                extra_iter
                    .next()
                    .unwrap()
                    .node
                    .write_to(writer)
                    .map_err(SerializeError::from)?;
            }
            item.write_element("si", writer)?;
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
        if let Some(ref val) = self.extension_list {
            val.write_element("extLst", writer)?;
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
        if !self.si.is_empty() {
            return false;
        }
        if self.extension_list.is_some() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for PhoneticRun {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        {
            let val = &self.sb;
            {
                let s = val.to_string();
                start.push_attribute(("sb", s.as_str()));
            }
        }
        {
            let val = &self.eb;
            {
                let s = val.to_string();
                start.push_attribute(("eb", s.as_str()));
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
            let val = &self.cell_type;
            {
                let start = BytesStart::new("t");
                writer.write_event(Event::Start(start))?;
                writer.write_event(Event::Text(BytesText::new(val.as_str())))?;
                writer.write_event(Event::End(BytesEnd::new("t")))?;
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
        false
    }
}

impl ToXml for RichTextElement {
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
            val.write_element("rPr", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
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
            let val = &self.cell_type;
            {
                let start = BytesStart::new("t");
                writer.write_event(Event::Start(start))?;
                writer.write_event(Event::Text(BytesText::new(val.as_str())))?;
                writer.write_event(Event::End(BytesEnd::new("t")))?;
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
        false
    }
}

impl ToXml for RichTextRunProperties {
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

impl ToXml for RichString {
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
        if let Some(ref val) = self.cell_type {
            {
                let start = BytesStart::new("t");
                writer.write_event(Event::Start(start))?;
                writer.write_event(Event::Text(BytesText::new(val.as_str())))?;
                writer.write_event(Event::End(BytesEnd::new("t")))?;
            }
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        for item in &self.reference {
            #[cfg(feature = "extra-children")]
            while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
                extra_iter
                    .next()
                    .unwrap()
                    .node
                    .write_to(writer)
                    .map_err(SerializeError::from)?;
            }
            item.write_element("r", writer)?;
            #[cfg(feature = "extra-children")]
            {
                emit_idx += 1;
            }
        }
        for item in &self.r_ph {
            #[cfg(feature = "extra-children")]
            while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
                extra_iter
                    .next()
                    .unwrap()
                    .node
                    .write_to(writer)
                    .map_err(SerializeError::from)?;
            }
            item.write_element("rPh", writer)?;
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
        if let Some(ref val) = self.phonetic_pr {
            val.write_element("phoneticPr", writer)?;
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
        if self.cell_type.is_some() {
            return false;
        }
        if !self.reference.is_empty() {
            return false;
        }
        if !self.r_ph.is_empty() {
            return false;
        }
        if self.phonetic_pr.is_some() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for PhoneticProperties {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        {
            let val = &self.font_id;
            {
                let s = val.to_string();
                start.push_attribute(("fontId", s.as_str()));
            }
        }
        if let Some(ref val) = self.r#type {
            {
                let s = val.to_string();
                start.push_attribute(("type", s.as_str()));
            }
        }
        if let Some(ref val) = self.alignment {
            {
                let s = val.to_string();
                start.push_attribute(("alignment", s.as_str()));
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

impl ToXml for RevisionHeaders {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        {
            let val = &self.guid;
            start.push_attribute(("guid", val.as_str()));
        }
        if let Some(ref val) = self.last_guid {
            start.push_attribute(("lastGuid", val.as_str()));
        }
        if let Some(ref val) = self.shared {
            start.push_attribute(("shared", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.disk_revisions {
            start.push_attribute(("diskRevisions", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.history {
            start.push_attribute(("history", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.track_revisions {
            start.push_attribute(("trackRevisions", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.exclusive {
            start.push_attribute(("exclusive", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.revision_id {
            {
                let s = val.to_string();
                start.push_attribute(("revisionId", s.as_str()));
            }
        }
        if let Some(ref val) = self.version {
            {
                let s = val.to_string();
                start.push_attribute(("version", s.as_str()));
            }
        }
        if let Some(ref val) = self.keep_change_history {
            start.push_attribute(("keepChangeHistory", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.protected {
            start.push_attribute(("protected", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.preserve_history {
            {
                let s = val.to_string();
                start.push_attribute(("preserveHistory", s.as_str()));
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
            item.write_element("header", writer)?;
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

impl ToXml for Revisions {
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

impl ToXml for SmlAGRevData {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        {
            let val = &self.r_id;
            {
                let s = val.to_string();
                start.push_attribute(("rId", s.as_str()));
            }
        }
        if let Some(ref val) = self.ua {
            start.push_attribute(("ua", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.ra {
            start.push_attribute(("ra", if *val { "1" } else { "0" }));
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

impl ToXml for RevisionHeader {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        {
            let val = &self.guid;
            start.push_attribute(("guid", val.as_str()));
        }
        {
            let val = &self.date_time;
            start.push_attribute(("dateTime", val.as_str()));
        }
        {
            let val = &self.max_sheet_id;
            {
                let s = val.to_string();
                start.push_attribute(("maxSheetId", s.as_str()));
            }
        }
        {
            let val = &self.user_name;
            start.push_attribute(("userName", val.as_str()));
        }
        if let Some(ref val) = self.min_r_id {
            {
                let s = val.to_string();
                start.push_attribute(("minRId", s.as_str()));
            }
        }
        if let Some(ref val) = self.max_r_id {
            {
                let s = val.to_string();
                start.push_attribute(("maxRId", s.as_str()));
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
            let val = &self.sheet_id_map;
            val.write_element("sheetIdMap", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        if let Some(ref val) = self.reviewed_list {
            val.write_element("reviewedList", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        if let Some(ref val) = self.extension_list {
            val.write_element("extLst", writer)?;
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

impl ToXml for CTSheetIdMap {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        if let Some(ref val) = self.count {
            {
                let s = val.to_string();
                start.push_attribute(("count", s.as_str()));
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
        for item in &self.sheet_id {
            #[cfg(feature = "extra-children")]
            while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
                extra_iter
                    .next()
                    .unwrap()
                    .node
                    .write_to(writer)
                    .map_err(SerializeError::from)?;
            }
            item.write_element("sheetId", writer)?;
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
        if !self.sheet_id.is_empty() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for CTSheetId {
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

impl ToXml for ReviewedRevisions {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        if let Some(ref val) = self.count {
            {
                let s = val.to_string();
                start.push_attribute(("count", s.as_str()));
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
        for item in &self.reviewed {
            #[cfg(feature = "extra-children")]
            while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
                extra_iter
                    .next()
                    .unwrap()
                    .node
                    .write_to(writer)
                    .map_err(SerializeError::from)?;
            }
            item.write_element("reviewed", writer)?;
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
        if !self.reviewed.is_empty() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for Reviewed {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        {
            let val = &self.r_id;
            {
                let s = val.to_string();
                start.push_attribute(("rId", s.as_str()));
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

impl ToXml for UndoInfo {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        {
            let val = &self.index;
            {
                let s = val.to_string();
                start.push_attribute(("index", s.as_str()));
            }
        }
        {
            let val = &self.exp;
            {
                let s = val.to_string();
                start.push_attribute(("exp", s.as_str()));
            }
        }
        if let Some(ref val) = self.ref3_d {
            start.push_attribute(("ref3D", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.array {
            start.push_attribute(("array", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.value {
            start.push_attribute(("v", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.nf {
            start.push_attribute(("nf", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.cs {
            start.push_attribute(("cs", if *val { "1" } else { "0" }));
        }
        {
            let val = &self.dr;
            start.push_attribute(("dr", val.as_str()));
        }
        if let Some(ref val) = self.dn {
            start.push_attribute(("dn", val.as_str()));
        }
        if let Some(ref val) = self.reference {
            start.push_attribute(("r", val.as_str()));
        }
        if let Some(ref val) = self.s_id {
            {
                let s = val.to_string();
                start.push_attribute(("sId", s.as_str()));
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

impl ToXml for RevisionRowColumn {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        {
            let val = &self.r_id;
            {
                let s = val.to_string();
                start.push_attribute(("rId", s.as_str()));
            }
        }
        if let Some(ref val) = self.ua {
            start.push_attribute(("ua", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.ra {
            start.push_attribute(("ra", if *val { "1" } else { "0" }));
        }
        {
            let val = &self.s_id;
            {
                let s = val.to_string();
                start.push_attribute(("sId", s.as_str()));
            }
        }
        if let Some(ref val) = self.eol {
            start.push_attribute(("eol", if *val { "1" } else { "0" }));
        }
        {
            let val = &self.reference;
            start.push_attribute(("ref", val.as_str()));
        }
        {
            let val = &self.action;
            {
                let s = val.to_string();
                start.push_attribute(("action", s.as_str()));
            }
        }
        if let Some(ref val) = self.edge {
            start.push_attribute(("edge", if *val { "1" } else { "0" }));
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

impl ToXml for RevisionMove {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        {
            let val = &self.r_id;
            {
                let s = val.to_string();
                start.push_attribute(("rId", s.as_str()));
            }
        }
        if let Some(ref val) = self.ua {
            start.push_attribute(("ua", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.ra {
            start.push_attribute(("ra", if *val { "1" } else { "0" }));
        }
        {
            let val = &self.sheet_id;
            {
                let s = val.to_string();
                start.push_attribute(("sheetId", s.as_str()));
            }
        }
        {
            let val = &self.source;
            start.push_attribute(("source", val.as_str()));
        }
        {
            let val = &self.destination;
            start.push_attribute(("destination", val.as_str()));
        }
        if let Some(ref val) = self.source_sheet_id {
            {
                let s = val.to_string();
                start.push_attribute(("sourceSheetId", s.as_str()));
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

impl ToXml for RevisionCustomView {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        {
            let val = &self.guid;
            start.push_attribute(("guid", val.as_str()));
        }
        {
            let val = &self.action;
            {
                let s = val.to_string();
                start.push_attribute(("action", s.as_str()));
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

impl ToXml for RevisionSheetRename {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        {
            let val = &self.r_id;
            {
                let s = val.to_string();
                start.push_attribute(("rId", s.as_str()));
            }
        }
        if let Some(ref val) = self.ua {
            start.push_attribute(("ua", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.ra {
            start.push_attribute(("ra", if *val { "1" } else { "0" }));
        }
        {
            let val = &self.sheet_id;
            {
                let s = val.to_string();
                start.push_attribute(("sheetId", s.as_str()));
            }
        }
        {
            let val = &self.old_name;
            start.push_attribute(("oldName", val.as_str()));
        }
        {
            let val = &self.new_name;
            start.push_attribute(("newName", val.as_str()));
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
        if let Some(ref val) = self.extension_list {
            val.write_element("extLst", writer)?;
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
        if self.extension_list.is_some() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for RevisionInsertSheet {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        {
            let val = &self.r_id;
            {
                let s = val.to_string();
                start.push_attribute(("rId", s.as_str()));
            }
        }
        if let Some(ref val) = self.ua {
            start.push_attribute(("ua", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.ra {
            start.push_attribute(("ra", if *val { "1" } else { "0" }));
        }
        {
            let val = &self.sheet_id;
            {
                let s = val.to_string();
                start.push_attribute(("sheetId", s.as_str()));
            }
        }
        {
            let val = &self.name;
            start.push_attribute(("name", val.as_str()));
        }
        {
            let val = &self.sheet_position;
            {
                let s = val.to_string();
                start.push_attribute(("sheetPosition", s.as_str()));
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

impl ToXml for RevisionCellChange {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        {
            let val = &self.r_id;
            {
                let s = val.to_string();
                start.push_attribute(("rId", s.as_str()));
            }
        }
        if let Some(ref val) = self.ua {
            start.push_attribute(("ua", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.ra {
            start.push_attribute(("ra", if *val { "1" } else { "0" }));
        }
        {
            let val = &self.s_id;
            {
                let s = val.to_string();
                start.push_attribute(("sId", s.as_str()));
            }
        }
        if let Some(ref val) = self.odxf {
            start.push_attribute(("odxf", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.xf_dxf {
            start.push_attribute(("xfDxf", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.style_index {
            start.push_attribute(("s", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.dxf {
            start.push_attribute(("dxf", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.number_format_id {
            {
                let s = val.to_string();
                start.push_attribute(("numFmtId", s.as_str()));
            }
        }
        if let Some(ref val) = self.quote_prefix {
            start.push_attribute(("quotePrefix", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.old_quote_prefix {
            start.push_attribute(("oldQuotePrefix", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.placeholder {
            start.push_attribute(("ph", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.old_ph {
            start.push_attribute(("oldPh", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.end_of_list_formula_update {
            start.push_attribute(("endOfListFormulaUpdate", if *val { "1" } else { "0" }));
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
        if let Some(ref val) = self.oc {
            val.write_element("oc", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
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
            let val = &self.nc;
            val.write_element("nc", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        if let Some(ref val) = self.ndxf {
            val.write_element("ndxf", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        if let Some(ref val) = self.extension_list {
            val.write_element("extLst", writer)?;
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
        if self.oc.is_some() {
            return false;
        }
        false
    }
}

impl ToXml for RevisionFormatting {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        {
            let val = &self.sheet_id;
            {
                let s = val.to_string();
                start.push_attribute(("sheetId", s.as_str()));
            }
        }
        if let Some(ref val) = self.xf_dxf {
            start.push_attribute(("xfDxf", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.style_index {
            start.push_attribute(("s", if *val { "1" } else { "0" }));
        }
        {
            let val = &self.square_reference;
            start.push_attribute(("sqref", val.as_str()));
        }
        if let Some(ref val) = self.start {
            {
                let s = val.to_string();
                start.push_attribute(("start", s.as_str()));
            }
        }
        if let Some(ref val) = self.length {
            {
                let s = val.to_string();
                start.push_attribute(("length", s.as_str()));
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
        if let Some(ref val) = self.dxf {
            val.write_element("dxf", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        if let Some(ref val) = self.extension_list {
            val.write_element("extLst", writer)?;
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
        if self.dxf.is_some() {
            return false;
        }
        if self.extension_list.is_some() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for RevisionAutoFormatting {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        {
            let val = &self.sheet_id;
            {
                let s = val.to_string();
                start.push_attribute(("sheetId", s.as_str()));
            }
        }
        if let Some(ref val) = self.auto_format_id {
            {
                let s = val.to_string();
                start.push_attribute(("autoFormatId", s.as_str()));
            }
        }
        if let Some(ref val) = self.apply_number_formats {
            start.push_attribute(("applyNumberFormats", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.apply_border_formats {
            start.push_attribute(("applyBorderFormats", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.apply_font_formats {
            start.push_attribute(("applyFontFormats", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.apply_pattern_formats {
            start.push_attribute(("applyPatternFormats", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.apply_alignment_formats {
            start.push_attribute(("applyAlignmentFormats", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.apply_width_height_formats {
            start.push_attribute(("applyWidthHeightFormats", if *val { "1" } else { "0" }));
        }
        {
            let val = &self.reference;
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

impl ToXml for RevisionComment {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        {
            let val = &self.sheet_id;
            {
                let s = val.to_string();
                start.push_attribute(("sheetId", s.as_str()));
            }
        }
        {
            let val = &self.cell;
            start.push_attribute(("cell", val.as_str()));
        }
        {
            let val = &self.guid;
            start.push_attribute(("guid", val.as_str()));
        }
        if let Some(ref val) = self.action {
            {
                let s = val.to_string();
                start.push_attribute(("action", s.as_str()));
            }
        }
        if let Some(ref val) = self.always_show {
            start.push_attribute(("alwaysShow", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.old {
            start.push_attribute(("old", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.hidden_row {
            start.push_attribute(("hiddenRow", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.hidden_column {
            start.push_attribute(("hiddenColumn", if *val { "1" } else { "0" }));
        }
        {
            let val = &self.author;
            start.push_attribute(("author", val.as_str()));
        }
        if let Some(ref val) = self.old_length {
            {
                let s = val.to_string();
                start.push_attribute(("oldLength", s.as_str()));
            }
        }
        if let Some(ref val) = self.new_length {
            {
                let s = val.to_string();
                start.push_attribute(("newLength", s.as_str()));
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

impl ToXml for RevisionDefinedName {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        {
            let val = &self.r_id;
            {
                let s = val.to_string();
                start.push_attribute(("rId", s.as_str()));
            }
        }
        if let Some(ref val) = self.ua {
            start.push_attribute(("ua", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.ra {
            start.push_attribute(("ra", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.local_sheet_id {
            {
                let s = val.to_string();
                start.push_attribute(("localSheetId", s.as_str()));
            }
        }
        if let Some(ref val) = self.custom_view {
            start.push_attribute(("customView", if *val { "1" } else { "0" }));
        }
        {
            let val = &self.name;
            start.push_attribute(("name", val.as_str()));
        }
        if let Some(ref val) = self.function {
            start.push_attribute(("function", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.old_function {
            start.push_attribute(("oldFunction", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.function_group_id {
            {
                let s = val.to_string();
                start.push_attribute(("functionGroupId", s.as_str()));
            }
        }
        if let Some(ref val) = self.old_function_group_id {
            {
                let s = val.to_string();
                start.push_attribute(("oldFunctionGroupId", s.as_str()));
            }
        }
        if let Some(ref val) = self.shortcut_key {
            {
                let s = val.to_string();
                start.push_attribute(("shortcutKey", s.as_str()));
            }
        }
        if let Some(ref val) = self.old_shortcut_key {
            {
                let s = val.to_string();
                start.push_attribute(("oldShortcutKey", s.as_str()));
            }
        }
        if let Some(ref val) = self.hidden {
            start.push_attribute(("hidden", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.old_hidden {
            start.push_attribute(("oldHidden", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.custom_menu {
            start.push_attribute(("customMenu", val.as_str()));
        }
        if let Some(ref val) = self.old_custom_menu {
            start.push_attribute(("oldCustomMenu", val.as_str()));
        }
        if let Some(ref val) = self.description {
            start.push_attribute(("description", val.as_str()));
        }
        if let Some(ref val) = self.old_description {
            start.push_attribute(("oldDescription", val.as_str()));
        }
        if let Some(ref val) = self.help {
            start.push_attribute(("help", val.as_str()));
        }
        if let Some(ref val) = self.old_help {
            start.push_attribute(("oldHelp", val.as_str()));
        }
        if let Some(ref val) = self.status_bar {
            start.push_attribute(("statusBar", val.as_str()));
        }
        if let Some(ref val) = self.old_status_bar {
            start.push_attribute(("oldStatusBar", val.as_str()));
        }
        if let Some(ref val) = self.comment {
            start.push_attribute(("comment", val.as_str()));
        }
        if let Some(ref val) = self.old_comment {
            start.push_attribute(("oldComment", val.as_str()));
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
        if let Some(ref val) = self.formula {
            {
                let start = BytesStart::new("formula");
                writer.write_event(Event::Start(start))?;
                writer.write_event(Event::Text(BytesText::new(val.as_str())))?;
                writer.write_event(Event::End(BytesEnd::new("formula")))?;
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
        if let Some(ref val) = self.old_formula {
            {
                let start = BytesStart::new("oldFormula");
                writer.write_event(Event::Start(start))?;
                writer.write_event(Event::Text(BytesText::new(val.as_str())))?;
                writer.write_event(Event::End(BytesEnd::new("oldFormula")))?;
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
        if let Some(ref val) = self.extension_list {
            val.write_element("extLst", writer)?;
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
        if self.formula.is_some() {
            return false;
        }
        if self.old_formula.is_some() {
            return false;
        }
        if self.extension_list.is_some() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for RevisionConflict {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        {
            let val = &self.r_id;
            {
                let s = val.to_string();
                start.push_attribute(("rId", s.as_str()));
            }
        }
        if let Some(ref val) = self.ua {
            start.push_attribute(("ua", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.ra {
            start.push_attribute(("ra", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.sheet_id {
            {
                let s = val.to_string();
                start.push_attribute(("sheetId", s.as_str()));
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

impl ToXml for RevisionQueryTableField {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        {
            let val = &self.sheet_id;
            {
                let s = val.to_string();
                start.push_attribute(("sheetId", s.as_str()));
            }
        }
        {
            let val = &self.reference;
            start.push_attribute(("ref", val.as_str()));
        }
        {
            let val = &self.field_id;
            {
                let s = val.to_string();
                start.push_attribute(("fieldId", s.as_str()));
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

impl ToXml for Users {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        if let Some(ref val) = self.count {
            {
                let s = val.to_string();
                start.push_attribute(("count", s.as_str()));
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
        for item in &self.user_info {
            #[cfg(feature = "extra-children")]
            while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
                extra_iter
                    .next()
                    .unwrap()
                    .node
                    .write_to(writer)
                    .map_err(SerializeError::from)?;
            }
            item.write_element("userInfo", writer)?;
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
        if !self.user_info.is_empty() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for SharedUser {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        {
            let val = &self.guid;
            start.push_attribute(("guid", val.as_str()));
        }
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
        {
            let val = &self.date_time;
            start.push_attribute(("dateTime", val.as_str()));
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
        if let Some(ref val) = self.extension_list {
            val.write_element("extLst", writer)?;
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
        if self.extension_list.is_some() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for CTMacrosheet {
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
        if let Some(ref val) = self.sheet_properties {
            val.write_element("sheetPr", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        if let Some(ref val) = self.dimension {
            val.write_element("dimension", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        if let Some(ref val) = self.sheet_views {
            val.write_element("sheetViews", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        if let Some(ref val) = self.sheet_format {
            val.write_element("sheetFormatPr", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        for item in &self.cols {
            #[cfg(feature = "extra-children")]
            while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
                extra_iter
                    .next()
                    .unwrap()
                    .node
                    .write_to(writer)
                    .map_err(SerializeError::from)?;
            }
            item.write_element("cols", writer)?;
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
        {
            let val = &self.sheet_data;
            val.write_element("sheetData", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        if let Some(ref val) = self.sheet_protection {
            val.write_element("sheetProtection", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        if let Some(ref val) = self.auto_filter {
            val.write_element("autoFilter", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        if let Some(ref val) = self.sort_state {
            val.write_element("sortState", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        if let Some(ref val) = self.data_consolidate {
            val.write_element("dataConsolidate", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        if let Some(ref val) = self.custom_sheet_views {
            val.write_element("customSheetViews", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        if let Some(ref val) = self.phonetic_pr {
            val.write_element("phoneticPr", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        for item in &self.conditional_formatting {
            #[cfg(feature = "extra-children")]
            while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
                extra_iter
                    .next()
                    .unwrap()
                    .node
                    .write_to(writer)
                    .map_err(SerializeError::from)?;
            }
            item.write_element("conditionalFormatting", writer)?;
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
        if let Some(ref val) = self.print_options {
            val.write_element("printOptions", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        if let Some(ref val) = self.page_margins {
            val.write_element("pageMargins", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        if let Some(ref val) = self.page_setup {
            val.write_element("pageSetup", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        if let Some(ref val) = self.header_footer {
            val.write_element("headerFooter", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        if let Some(ref val) = self.row_breaks {
            val.write_element("rowBreaks", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        if let Some(ref val) = self.col_breaks {
            val.write_element("colBreaks", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        if let Some(ref val) = self.custom_properties {
            val.write_element("customProperties", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        if let Some(ref val) = self.drawing {
            val.write_element("drawing", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        if let Some(ref val) = self.legacy_drawing {
            val.write_element("legacyDrawing", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        if let Some(ref val) = self.legacy_drawing_h_f {
            val.write_element("legacyDrawingHF", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        if let Some(ref val) = self.drawing_h_f {
            val.write_element("drawingHF", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
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
            val.write_element("picture", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        if let Some(ref val) = self.ole_objects {
            val.write_element("oleObjects", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        if let Some(ref val) = self.extension_list {
            val.write_element("extLst", writer)?;
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
        if self.sheet_properties.is_some() {
            return false;
        }
        if self.dimension.is_some() {
            return false;
        }
        if self.sheet_views.is_some() {
            return false;
        }
        if self.sheet_format.is_some() {
            return false;
        }
        if !self.cols.is_empty() {
            return false;
        }
        false
    }
}

impl ToXml for CTDialogsheet {
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
        if let Some(ref val) = self.sheet_properties {
            val.write_element("sheetPr", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        if let Some(ref val) = self.sheet_views {
            val.write_element("sheetViews", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        if let Some(ref val) = self.sheet_format {
            val.write_element("sheetFormatPr", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        if let Some(ref val) = self.sheet_protection {
            val.write_element("sheetProtection", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        if let Some(ref val) = self.custom_sheet_views {
            val.write_element("customSheetViews", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        if let Some(ref val) = self.print_options {
            val.write_element("printOptions", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        if let Some(ref val) = self.page_margins {
            val.write_element("pageMargins", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        if let Some(ref val) = self.page_setup {
            val.write_element("pageSetup", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        if let Some(ref val) = self.header_footer {
            val.write_element("headerFooter", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        if let Some(ref val) = self.drawing {
            val.write_element("drawing", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        if let Some(ref val) = self.legacy_drawing {
            val.write_element("legacyDrawing", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        if let Some(ref val) = self.legacy_drawing_h_f {
            val.write_element("legacyDrawingHF", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        if let Some(ref val) = self.drawing_h_f {
            val.write_element("drawingHF", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        if let Some(ref val) = self.ole_objects {
            val.write_element("oleObjects", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
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
            val.write_element("controls", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        if let Some(ref val) = self.extension_list {
            val.write_element("extLst", writer)?;
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
        if self.sheet_properties.is_some() {
            return false;
        }
        if self.sheet_views.is_some() {
            return false;
        }
        if self.sheet_format.is_some() {
            return false;
        }
        if self.sheet_protection.is_some() {
            return false;
        }
        if self.custom_sheet_views.is_some() {
            return false;
        }
        if self.print_options.is_some() {
            return false;
        }
        if self.page_margins.is_some() {
            return false;
        }
        if self.page_setup.is_some() {
            return false;
        }
        if self.header_footer.is_some() {
            return false;
        }
        if self.drawing.is_some() {
            return false;
        }
        if self.legacy_drawing.is_some() {
            return false;
        }
        if self.legacy_drawing_h_f.is_some() {
            return false;
        }
        if self.drawing_h_f.is_some() {
            return false;
        }
        if self.ole_objects.is_some() {
            return false;
        }
        if self.controls.is_some() {
            return false;
        }
        if self.extension_list.is_some() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for Worksheet {
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
        #[cfg(feature = "sml-styling")]
        if let Some(ref val) = self.sheet_properties {
            val.write_element("sheetPr", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        if let Some(ref val) = self.dimension {
            val.write_element("dimension", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        if let Some(ref val) = self.sheet_views {
            val.write_element("sheetViews", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        #[cfg(feature = "sml-styling")]
        if let Some(ref val) = self.sheet_format {
            val.write_element("sheetFormatPr", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "sml-styling")]
        for item in &self.cols {
            #[cfg(feature = "extra-children")]
            while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
                extra_iter
                    .next()
                    .unwrap()
                    .node
                    .write_to(writer)
                    .map_err(SerializeError::from)?;
            }
            item.write_element("cols", writer)?;
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
        {
            let val = &self.sheet_data;
            val.write_element("sheetData", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        #[cfg(feature = "sml-formulas")]
        if let Some(ref val) = self.sheet_calc_pr {
            val.write_element("sheetCalcPr", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        #[cfg(feature = "sml-protection")]
        if let Some(ref val) = self.sheet_protection {
            val.write_element("sheetProtection", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        #[cfg(feature = "sml-protection")]
        if let Some(ref val) = self.protected_ranges {
            val.write_element("protectedRanges", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        #[cfg(feature = "sml-formulas-advanced")]
        if let Some(ref val) = self.scenarios {
            val.write_element("scenarios", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        #[cfg(feature = "sml-filtering")]
        if let Some(ref val) = self.auto_filter {
            val.write_element("autoFilter", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        #[cfg(feature = "sml-filtering")]
        if let Some(ref val) = self.sort_state {
            val.write_element("sortState", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        #[cfg(feature = "sml-formulas-advanced")]
        if let Some(ref val) = self.data_consolidate {
            val.write_element("dataConsolidate", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        #[cfg(feature = "sml-structure")]
        if let Some(ref val) = self.custom_sheet_views {
            val.write_element("customSheetViews", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        if let Some(ref val) = self.merged_cells {
            val.write_element("mergeCells", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        #[cfg(feature = "sml-i18n")]
        if let Some(ref val) = self.phonetic_pr {
            val.write_element("phoneticPr", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "sml-styling")]
        for item in &self.conditional_formatting {
            #[cfg(feature = "extra-children")]
            while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
                extra_iter
                    .next()
                    .unwrap()
                    .node
                    .write_to(writer)
                    .map_err(SerializeError::from)?;
            }
            item.write_element("conditionalFormatting", writer)?;
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
        #[cfg(feature = "sml-validation")]
        if let Some(ref val) = self.data_validations {
            val.write_element("dataValidations", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        #[cfg(feature = "sml-hyperlinks")]
        if let Some(ref val) = self.hyperlinks {
            val.write_element("hyperlinks", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        #[cfg(feature = "sml-layout")]
        if let Some(ref val) = self.print_options {
            val.write_element("printOptions", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        #[cfg(feature = "sml-layout")]
        if let Some(ref val) = self.page_margins {
            val.write_element("pageMargins", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        #[cfg(feature = "sml-layout")]
        if let Some(ref val) = self.page_setup {
            val.write_element("pageSetup", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        #[cfg(feature = "sml-layout")]
        if let Some(ref val) = self.header_footer {
            val.write_element("headerFooter", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        #[cfg(feature = "sml-layout")]
        if let Some(ref val) = self.row_breaks {
            val.write_element("rowBreaks", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        #[cfg(feature = "sml-layout")]
        if let Some(ref val) = self.col_breaks {
            val.write_element("colBreaks", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        #[cfg(feature = "sml-metadata")]
        if let Some(ref val) = self.custom_properties {
            val.write_element("customProperties", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        #[cfg(feature = "sml-formulas-advanced")]
        if let Some(ref val) = self.cell_watches {
            val.write_element("cellWatches", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        #[cfg(feature = "sml-validation")]
        if let Some(ref val) = self.ignored_errors {
            val.write_element("ignoredErrors", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        #[cfg(feature = "sml-metadata")]
        if let Some(ref val) = self.smart_tags {
            val.write_element("smartTags", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        #[cfg(feature = "sml-drawings")]
        if let Some(ref val) = self.drawing {
            val.write_element("drawing", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        #[cfg(feature = "sml-comments")]
        if let Some(ref val) = self.legacy_drawing {
            val.write_element("legacyDrawing", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        #[cfg(feature = "sml-layout")]
        if let Some(ref val) = self.legacy_drawing_h_f {
            val.write_element("legacyDrawingHF", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        #[cfg(feature = "sml-drawings")]
        if let Some(ref val) = self.drawing_h_f {
            val.write_element("drawingHF", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        #[cfg(feature = "sml-drawings")]
        if let Some(ref val) = self.picture {
            val.write_element("picture", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        #[cfg(feature = "sml-external")]
        if let Some(ref val) = self.ole_objects {
            val.write_element("oleObjects", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        #[cfg(feature = "sml-external")]
        if let Some(ref val) = self.controls {
            val.write_element("controls", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        #[cfg(feature = "sml-external")]
        if let Some(ref val) = self.web_publish_items {
            val.write_element("webPublishItems", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        #[cfg(feature = "sml-tables")]
        if let Some(ref val) = self.table_parts {
            val.write_element("tableParts", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        #[cfg(feature = "sml-extensions")]
        if let Some(ref val) = self.extension_list {
            val.write_element("extLst", writer)?;
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
        #[cfg(feature = "sml-styling")]
        if self.sheet_properties.is_some() {
            return false;
        }
        if self.dimension.is_some() {
            return false;
        }
        if self.sheet_views.is_some() {
            return false;
        }
        #[cfg(feature = "sml-styling")]
        if self.sheet_format.is_some() {
            return false;
        }
        #[cfg(feature = "sml-styling")]
        if !self.cols.is_empty() {
            return false;
        }
        false
    }
}

impl ToXml for SheetData {
    fn write_children<W: Write>(&self, writer: &mut Writer<W>) -> Result<(), SerializeError> {
        #[cfg(feature = "extra-children")]
        let mut extra_iter = self.extra_children.iter().peekable();
        #[cfg(feature = "extra-children")]
        let mut emit_idx: usize = 0;
        for item in &self.row {
            #[cfg(feature = "extra-children")]
            while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
                extra_iter
                    .next()
                    .unwrap()
                    .node
                    .write_to(writer)
                    .map_err(SerializeError::from)?;
            }
            item.write_element("row", writer)?;
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
        if !self.row.is_empty() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for SheetCalcProperties {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        if let Some(ref val) = self.full_calc_on_load {
            start.push_attribute(("fullCalcOnLoad", if *val { "1" } else { "0" }));
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

impl ToXml for SheetFormat {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        if let Some(ref val) = self.base_col_width {
            {
                let s = val.to_string();
                start.push_attribute(("baseColWidth", s.as_str()));
            }
        }
        if let Some(ref val) = self.default_col_width {
            {
                let s = val.to_string();
                start.push_attribute(("defaultColWidth", s.as_str()));
            }
        }
        {
            let val = &self.default_row_height;
            {
                let s = val.to_string();
                start.push_attribute(("defaultRowHeight", s.as_str()));
            }
        }
        if let Some(ref val) = self.custom_height {
            start.push_attribute(("customHeight", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.zero_height {
            start.push_attribute(("zeroHeight", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.thick_top {
            start.push_attribute(("thickTop", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.thick_bottom {
            start.push_attribute(("thickBottom", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.outline_level_row {
            {
                let s = val.to_string();
                start.push_attribute(("outlineLevelRow", s.as_str()));
            }
        }
        if let Some(ref val) = self.outline_level_col {
            {
                let s = val.to_string();
                start.push_attribute(("outlineLevelCol", s.as_str()));
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
    fn write_children<W: Write>(&self, writer: &mut Writer<W>) -> Result<(), SerializeError> {
        #[cfg(feature = "extra-children")]
        let mut extra_iter = self.extra_children.iter().peekable();
        #[cfg(feature = "extra-children")]
        let mut emit_idx: usize = 0;
        for item in &self.col {
            #[cfg(feature = "extra-children")]
            while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
                extra_iter
                    .next()
                    .unwrap()
                    .node
                    .write_to(writer)
                    .map_err(SerializeError::from)?;
            }
            item.write_element("col", writer)?;
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

impl ToXml for Column {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        #[cfg(feature = "sml-styling")]
        {
            let val = &self.start_column;
            {
                let s = val.to_string();
                start.push_attribute(("min", s.as_str()));
            }
        }
        #[cfg(feature = "sml-styling")]
        {
            let val = &self.end_column;
            {
                let s = val.to_string();
                start.push_attribute(("max", s.as_str()));
            }
        }
        #[cfg(feature = "sml-styling")]
        if let Some(ref val) = self.width {
            {
                let s = val.to_string();
                start.push_attribute(("width", s.as_str()));
            }
        }
        #[cfg(feature = "sml-styling")]
        if let Some(ref val) = self.style {
            {
                let s = val.to_string();
                start.push_attribute(("style", s.as_str()));
            }
        }
        #[cfg(feature = "sml-structure")]
        if let Some(ref val) = self.hidden {
            start.push_attribute(("hidden", if *val { "1" } else { "0" }));
        }
        #[cfg(feature = "sml-styling")]
        if let Some(ref val) = self.best_fit {
            start.push_attribute(("bestFit", if *val { "1" } else { "0" }));
        }
        #[cfg(feature = "sml-styling")]
        if let Some(ref val) = self.custom_width {
            start.push_attribute(("customWidth", if *val { "1" } else { "0" }));
        }
        #[cfg(feature = "sml-i18n")]
        if let Some(ref val) = self.phonetic {
            start.push_attribute(("phonetic", if *val { "1" } else { "0" }));
        }
        #[cfg(feature = "sml-structure")]
        if let Some(ref val) = self.outline_level {
            {
                let s = val.to_string();
                start.push_attribute(("outlineLevel", s.as_str()));
            }
        }
        #[cfg(feature = "sml-structure")]
        if let Some(ref val) = self.collapsed {
            start.push_attribute(("collapsed", if *val { "1" } else { "0" }));
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

impl ToXml for Row {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        if let Some(ref val) = self.reference {
            {
                let s = val.to_string();
                start.push_attribute(("r", s.as_str()));
            }
        }
        if let Some(ref val) = self.cell_spans {
            start.push_attribute(("spans", val.as_str()));
        }
        if let Some(ref val) = self.style_index {
            {
                let s = val.to_string();
                start.push_attribute(("s", s.as_str()));
            }
        }
        #[cfg(feature = "sml-styling")]
        if let Some(ref val) = self.custom_format {
            start.push_attribute(("customFormat", if *val { "1" } else { "0" }));
        }
        #[cfg(feature = "sml-styling")]
        if let Some(ref val) = self.height {
            {
                let s = val.to_string();
                start.push_attribute(("ht", s.as_str()));
            }
        }
        #[cfg(feature = "sml-structure")]
        if let Some(ref val) = self.hidden {
            start.push_attribute(("hidden", if *val { "1" } else { "0" }));
        }
        #[cfg(feature = "sml-styling")]
        if let Some(ref val) = self.custom_height {
            start.push_attribute(("customHeight", if *val { "1" } else { "0" }));
        }
        #[cfg(feature = "sml-structure")]
        if let Some(ref val) = self.outline_level {
            {
                let s = val.to_string();
                start.push_attribute(("outlineLevel", s.as_str()));
            }
        }
        #[cfg(feature = "sml-structure")]
        if let Some(ref val) = self.collapsed {
            start.push_attribute(("collapsed", if *val { "1" } else { "0" }));
        }
        #[cfg(feature = "sml-styling")]
        if let Some(ref val) = self.thick_top {
            start.push_attribute(("thickTop", if *val { "1" } else { "0" }));
        }
        #[cfg(feature = "sml-styling")]
        if let Some(ref val) = self.thick_bot {
            start.push_attribute(("thickBot", if *val { "1" } else { "0" }));
        }
        #[cfg(feature = "sml-i18n")]
        if let Some(ref val) = self.placeholder {
            start.push_attribute(("ph", if *val { "1" } else { "0" }));
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
        for item in &self.cells {
            #[cfg(feature = "extra-children")]
            while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
                extra_iter
                    .next()
                    .unwrap()
                    .node
                    .write_to(writer)
                    .map_err(SerializeError::from)?;
            }
            item.write_element("c", writer)?;
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
        #[cfg(feature = "sml-extensions")]
        if let Some(ref val) = self.extension_list {
            val.write_element("extLst", writer)?;
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
        if !self.cells.is_empty() {
            return false;
        }
        #[cfg(feature = "sml-extensions")]
        if self.extension_list.is_some() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for Cell {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        if let Some(ref val) = self.reference {
            start.push_attribute(("r", val.as_str()));
        }
        if let Some(ref val) = self.style_index {
            {
                let s = val.to_string();
                start.push_attribute(("s", s.as_str()));
            }
        }
        if let Some(ref val) = self.cell_type {
            {
                let s = val.to_string();
                start.push_attribute(("t", s.as_str()));
            }
        }
        #[cfg(feature = "sml-metadata")]
        if let Some(ref val) = self.cm {
            {
                let s = val.to_string();
                start.push_attribute(("cm", s.as_str()));
            }
        }
        #[cfg(feature = "sml-metadata")]
        if let Some(ref val) = self.vm {
            {
                let s = val.to_string();
                start.push_attribute(("vm", s.as_str()));
            }
        }
        #[cfg(feature = "sml-i18n")]
        if let Some(ref val) = self.placeholder {
            start.push_attribute(("ph", if *val { "1" } else { "0" }));
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
        if let Some(ref val) = self.formula {
            val.write_element("f", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
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
            {
                let start = BytesStart::new("v");
                writer.write_event(Event::Start(start))?;
                writer.write_event(Event::Text(BytesText::new(val.as_str())))?;
                writer.write_event(Event::End(BytesEnd::new("v")))?;
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
        if let Some(ref val) = self.is {
            val.write_element("is", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        #[cfg(feature = "sml-extensions")]
        if let Some(ref val) = self.extension_list {
            val.write_element("extLst", writer)?;
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
        if self.formula.is_some() {
            return false;
        }
        if self.value.is_some() {
            return false;
        }
        if self.is.is_some() {
            return false;
        }
        #[cfg(feature = "sml-extensions")]
        if self.extension_list.is_some() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for SheetProperties {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        if let Some(ref val) = self.sync_horizontal {
            start.push_attribute(("syncHorizontal", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.sync_vertical {
            start.push_attribute(("syncVertical", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.sync_ref {
            start.push_attribute(("syncRef", val.as_str()));
        }
        if let Some(ref val) = self.transition_evaluation {
            start.push_attribute(("transitionEvaluation", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.transition_entry {
            start.push_attribute(("transitionEntry", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.published {
            start.push_attribute(("published", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.code_name {
            start.push_attribute(("codeName", val.as_str()));
        }
        if let Some(ref val) = self.filter_mode {
            start.push_attribute(("filterMode", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.enable_format_conditions_calculation {
            start.push_attribute((
                "enableFormatConditionsCalculation",
                if *val { "1" } else { "0" },
            ));
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
        if let Some(ref val) = self.tab_color {
            val.write_element("tabColor", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        if let Some(ref val) = self.outline_pr {
            val.write_element("outlinePr", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        if let Some(ref val) = self.page_set_up_pr {
            val.write_element("pageSetUpPr", writer)?;
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
        if self.tab_color.is_some() {
            return false;
        }
        if self.outline_pr.is_some() {
            return false;
        }
        if self.page_set_up_pr.is_some() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for SheetDimension {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        {
            let val = &self.reference;
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

impl ToXml for SheetViews {
    fn write_children<W: Write>(&self, writer: &mut Writer<W>) -> Result<(), SerializeError> {
        #[cfg(feature = "extra-children")]
        let mut extra_iter = self.extra_children.iter().peekable();
        #[cfg(feature = "extra-children")]
        let mut emit_idx: usize = 0;
        for item in &self.sheet_view {
            #[cfg(feature = "extra-children")]
            while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
                extra_iter
                    .next()
                    .unwrap()
                    .node
                    .write_to(writer)
                    .map_err(SerializeError::from)?;
            }
            item.write_element("sheetView", writer)?;
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
        if let Some(ref val) = self.extension_list {
            val.write_element("extLst", writer)?;
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
        if !self.sheet_view.is_empty() {
            return false;
        }
        if self.extension_list.is_some() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for SheetView {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        #[cfg(feature = "sml-protection")]
        if let Some(ref val) = self.window_protection {
            start.push_attribute(("windowProtection", if *val { "1" } else { "0" }));
        }
        #[cfg(feature = "sml-formulas")]
        if let Some(ref val) = self.show_formulas {
            start.push_attribute(("showFormulas", if *val { "1" } else { "0" }));
        }
        #[cfg(feature = "sml-styling")]
        if let Some(ref val) = self.show_grid_lines {
            start.push_attribute(("showGridLines", if *val { "1" } else { "0" }));
        }
        #[cfg(feature = "sml-styling")]
        if let Some(ref val) = self.show_row_col_headers {
            start.push_attribute(("showRowColHeaders", if *val { "1" } else { "0" }));
        }
        #[cfg(feature = "sml-styling")]
        if let Some(ref val) = self.show_zeros {
            start.push_attribute(("showZeros", if *val { "1" } else { "0" }));
        }
        #[cfg(feature = "sml-i18n")]
        if let Some(ref val) = self.right_to_left {
            start.push_attribute(("rightToLeft", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.tab_selected {
            start.push_attribute(("tabSelected", if *val { "1" } else { "0" }));
        }
        #[cfg(feature = "sml-layout")]
        if let Some(ref val) = self.show_ruler {
            start.push_attribute(("showRuler", if *val { "1" } else { "0" }));
        }
        #[cfg(feature = "sml-structure")]
        if let Some(ref val) = self.show_outline_symbols {
            start.push_attribute(("showOutlineSymbols", if *val { "1" } else { "0" }));
        }
        #[cfg(feature = "sml-styling")]
        if let Some(ref val) = self.default_grid_color {
            start.push_attribute(("defaultGridColor", if *val { "1" } else { "0" }));
        }
        #[cfg(feature = "sml-layout")]
        if let Some(ref val) = self.show_white_space {
            start.push_attribute(("showWhiteSpace", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.view {
            {
                let s = val.to_string();
                start.push_attribute(("view", s.as_str()));
            }
        }
        if let Some(ref val) = self.top_left_cell {
            start.push_attribute(("topLeftCell", val.as_str()));
        }
        #[cfg(feature = "sml-styling")]
        if let Some(ref val) = self.color_id {
            {
                let s = val.to_string();
                start.push_attribute(("colorId", s.as_str()));
            }
        }
        if let Some(ref val) = self.zoom_scale {
            {
                let s = val.to_string();
                start.push_attribute(("zoomScale", s.as_str()));
            }
        }
        if let Some(ref val) = self.zoom_scale_normal {
            {
                let s = val.to_string();
                start.push_attribute(("zoomScaleNormal", s.as_str()));
            }
        }
        #[cfg(feature = "sml-layout")]
        if let Some(ref val) = self.zoom_scale_sheet_layout_view {
            {
                let s = val.to_string();
                start.push_attribute(("zoomScaleSheetLayoutView", s.as_str()));
            }
        }
        #[cfg(feature = "sml-layout")]
        if let Some(ref val) = self.zoom_scale_page_layout_view {
            {
                let s = val.to_string();
                start.push_attribute(("zoomScalePageLayoutView", s.as_str()));
            }
        }
        {
            let val = &self.workbook_view_id;
            {
                let s = val.to_string();
                start.push_attribute(("workbookViewId", s.as_str()));
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
        #[cfg(feature = "sml-structure")]
        if let Some(ref val) = self.pane {
            val.write_element("pane", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        for item in &self.selection {
            #[cfg(feature = "extra-children")]
            while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
                extra_iter
                    .next()
                    .unwrap()
                    .node
                    .write_to(writer)
                    .map_err(SerializeError::from)?;
            }
            item.write_element("selection", writer)?;
            #[cfg(feature = "extra-children")]
            {
                emit_idx += 1;
            }
        }
        #[cfg(feature = "sml-pivot")]
        for item in &self.pivot_selection {
            #[cfg(feature = "extra-children")]
            while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
                extra_iter
                    .next()
                    .unwrap()
                    .node
                    .write_to(writer)
                    .map_err(SerializeError::from)?;
            }
            item.write_element("pivotSelection", writer)?;
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
        #[cfg(feature = "sml-extensions")]
        if let Some(ref val) = self.extension_list {
            val.write_element("extLst", writer)?;
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
        #[cfg(feature = "sml-structure")]
        if self.pane.is_some() {
            return false;
        }
        if !self.selection.is_empty() {
            return false;
        }
        #[cfg(feature = "sml-pivot")]
        if !self.pivot_selection.is_empty() {
            return false;
        }
        #[cfg(feature = "sml-extensions")]
        if self.extension_list.is_some() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for Pane {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        #[cfg(feature = "sml-structure")]
        if let Some(ref val) = self.x_split {
            {
                let s = val.to_string();
                start.push_attribute(("xSplit", s.as_str()));
            }
        }
        #[cfg(feature = "sml-structure")]
        if let Some(ref val) = self.y_split {
            {
                let s = val.to_string();
                start.push_attribute(("ySplit", s.as_str()));
            }
        }
        #[cfg(feature = "sml-structure")]
        if let Some(ref val) = self.top_left_cell {
            start.push_attribute(("topLeftCell", val.as_str()));
        }
        #[cfg(feature = "sml-structure")]
        if let Some(ref val) = self.active_pane {
            {
                let s = val.to_string();
                start.push_attribute(("activePane", s.as_str()));
            }
        }
        #[cfg(feature = "sml-structure")]
        if let Some(ref val) = self.state {
            {
                let s = val.to_string();
                start.push_attribute(("state", s.as_str()));
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

impl ToXml for CTPivotSelection {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        if let Some(ref val) = self.pane {
            {
                let s = val.to_string();
                start.push_attribute(("pane", s.as_str()));
            }
        }
        if let Some(ref val) = self.show_header {
            start.push_attribute(("showHeader", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.label {
            start.push_attribute(("label", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.data {
            start.push_attribute(("data", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.extendable {
            start.push_attribute(("extendable", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.count {
            {
                let s = val.to_string();
                start.push_attribute(("count", s.as_str()));
            }
        }
        if let Some(ref val) = self.axis {
            {
                let s = val.to_string();
                start.push_attribute(("axis", s.as_str()));
            }
        }
        if let Some(ref val) = self.dimension {
            {
                let s = val.to_string();
                start.push_attribute(("dimension", s.as_str()));
            }
        }
        if let Some(ref val) = self.start {
            {
                let s = val.to_string();
                start.push_attribute(("start", s.as_str()));
            }
        }
        if let Some(ref val) = self.start_column {
            {
                let s = val.to_string();
                start.push_attribute(("min", s.as_str()));
            }
        }
        if let Some(ref val) = self.end_column {
            {
                let s = val.to_string();
                start.push_attribute(("max", s.as_str()));
            }
        }
        if let Some(ref val) = self.active_row {
            {
                let s = val.to_string();
                start.push_attribute(("activeRow", s.as_str()));
            }
        }
        if let Some(ref val) = self.active_col {
            {
                let s = val.to_string();
                start.push_attribute(("activeCol", s.as_str()));
            }
        }
        if let Some(ref val) = self.previous_row {
            {
                let s = val.to_string();
                start.push_attribute(("previousRow", s.as_str()));
            }
        }
        if let Some(ref val) = self.previous_col {
            {
                let s = val.to_string();
                start.push_attribute(("previousCol", s.as_str()));
            }
        }
        if let Some(ref val) = self.click {
            {
                let s = val.to_string();
                start.push_attribute(("click", s.as_str()));
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
            let val = &self.pivot_area;
            val.write_element("pivotArea", writer)?;
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

impl ToXml for Selection {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        if let Some(ref val) = self.pane {
            {
                let s = val.to_string();
                start.push_attribute(("pane", s.as_str()));
            }
        }
        if let Some(ref val) = self.active_cell {
            start.push_attribute(("activeCell", val.as_str()));
        }
        if let Some(ref val) = self.active_cell_id {
            {
                let s = val.to_string();
                start.push_attribute(("activeCellId", s.as_str()));
            }
        }
        if let Some(ref val) = self.square_reference {
            start.push_attribute(("sqref", val.as_str()));
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

impl ToXml for PageBreaks {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        if let Some(ref val) = self.count {
            {
                let s = val.to_string();
                start.push_attribute(("count", s.as_str()));
            }
        }
        if let Some(ref val) = self.manual_break_count {
            {
                let s = val.to_string();
                start.push_attribute(("manualBreakCount", s.as_str()));
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
        for item in &self.brk {
            #[cfg(feature = "extra-children")]
            while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
                extra_iter
                    .next()
                    .unwrap()
                    .node
                    .write_to(writer)
                    .map_err(SerializeError::from)?;
            }
            item.write_element("brk", writer)?;
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
        if !self.brk.is_empty() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for PageBreak {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        if let Some(ref val) = self.id {
            {
                let s = val.to_string();
                start.push_attribute(("id", s.as_str()));
            }
        }
        if let Some(ref val) = self.start_column {
            {
                let s = val.to_string();
                start.push_attribute(("min", s.as_str()));
            }
        }
        if let Some(ref val) = self.end_column {
            {
                let s = val.to_string();
                start.push_attribute(("max", s.as_str()));
            }
        }
        if let Some(ref val) = self.man {
            start.push_attribute(("man", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.pt {
            start.push_attribute(("pt", if *val { "1" } else { "0" }));
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

impl ToXml for OutlineProperties {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        if let Some(ref val) = self.apply_styles {
            start.push_attribute(("applyStyles", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.summary_below {
            start.push_attribute(("summaryBelow", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.summary_right {
            start.push_attribute(("summaryRight", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.show_outline_symbols {
            start.push_attribute(("showOutlineSymbols", if *val { "1" } else { "0" }));
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

impl ToXml for PageSetupProperties {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        if let Some(ref val) = self.auto_page_breaks {
            start.push_attribute(("autoPageBreaks", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.fit_to_page {
            start.push_attribute(("fitToPage", if *val { "1" } else { "0" }));
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

impl ToXml for CTDataConsolidate {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        if let Some(ref val) = self.function {
            {
                let s = val.to_string();
                start.push_attribute(("function", s.as_str()));
            }
        }
        if let Some(ref val) = self.start_labels {
            start.push_attribute(("startLabels", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.left_labels {
            start.push_attribute(("leftLabels", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.top_labels {
            start.push_attribute(("topLabels", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.link {
            start.push_attribute(("link", if *val { "1" } else { "0" }));
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
        if let Some(ref val) = self.data_refs {
            val.write_element("dataRefs", writer)?;
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
        if self.data_refs.is_some() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for CTDataRefs {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        if let Some(ref val) = self.count {
            {
                let s = val.to_string();
                start.push_attribute(("count", s.as_str()));
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
        for item in &self.data_ref {
            #[cfg(feature = "extra-children")]
            while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
                extra_iter
                    .next()
                    .unwrap()
                    .node
                    .write_to(writer)
                    .map_err(SerializeError::from)?;
            }
            item.write_element("dataRef", writer)?;
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
        if !self.data_ref.is_empty() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for CTDataRef {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        if let Some(ref val) = self.reference {
            start.push_attribute(("ref", val.as_str()));
        }
        if let Some(ref val) = self.name {
            start.push_attribute(("name", val.as_str()));
        }
        if let Some(ref val) = self.sheet {
            start.push_attribute(("sheet", val.as_str()));
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

impl ToXml for MergedCells {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        if let Some(ref val) = self.count {
            {
                let s = val.to_string();
                start.push_attribute(("count", s.as_str()));
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
        for item in &self.merge_cell {
            #[cfg(feature = "extra-children")]
            while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
                extra_iter
                    .next()
                    .unwrap()
                    .node
                    .write_to(writer)
                    .map_err(SerializeError::from)?;
            }
            item.write_element("mergeCell", writer)?;
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
        if !self.merge_cell.is_empty() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for MergedCell {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        {
            let val = &self.reference;
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

impl ToXml for SmartTags {
    fn write_children<W: Write>(&self, writer: &mut Writer<W>) -> Result<(), SerializeError> {
        #[cfg(feature = "extra-children")]
        let mut extra_iter = self.extra_children.iter().peekable();
        #[cfg(feature = "extra-children")]
        let mut emit_idx: usize = 0;
        for item in &self.cell_smart_tags {
            #[cfg(feature = "extra-children")]
            while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
                extra_iter
                    .next()
                    .unwrap()
                    .node
                    .write_to(writer)
                    .map_err(SerializeError::from)?;
            }
            item.write_element("cellSmartTags", writer)?;
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
        if !self.cell_smart_tags.is_empty() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for CellSmartTags {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        {
            let val = &self.reference;
            start.push_attribute(("r", val.as_str()));
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
        for item in &self.cell_smart_tag {
            #[cfg(feature = "extra-children")]
            while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
                extra_iter
                    .next()
                    .unwrap()
                    .node
                    .write_to(writer)
                    .map_err(SerializeError::from)?;
            }
            item.write_element("cellSmartTag", writer)?;
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
        if !self.cell_smart_tag.is_empty() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for CellSmartTag {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        {
            let val = &self.r#type;
            {
                let s = val.to_string();
                start.push_attribute(("type", s.as_str()));
            }
        }
        if let Some(ref val) = self.deleted {
            start.push_attribute(("deleted", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.xml_based {
            start.push_attribute(("xmlBased", if *val { "1" } else { "0" }));
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
        for item in &self.cell_smart_tag_pr {
            #[cfg(feature = "extra-children")]
            while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
                extra_iter
                    .next()
                    .unwrap()
                    .node
                    .write_to(writer)
                    .map_err(SerializeError::from)?;
            }
            item.write_element("cellSmartTagPr", writer)?;
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
        if !self.cell_smart_tag_pr.is_empty() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for CTCellSmartTagPr {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        {
            let val = &self.key;
            start.push_attribute(("key", val.as_str()));
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

impl ToXml for Drawing {
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

impl ToXml for LegacyDrawing {
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

impl ToXml for DrawingHeaderFooter {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        if let Some(ref val) = self.lho {
            {
                let s = val.to_string();
                start.push_attribute(("lho", s.as_str()));
            }
        }
        if let Some(ref val) = self.lhe {
            {
                let s = val.to_string();
                start.push_attribute(("lhe", s.as_str()));
            }
        }
        if let Some(ref val) = self.lhf {
            {
                let s = val.to_string();
                start.push_attribute(("lhf", s.as_str()));
            }
        }
        if let Some(ref val) = self.cho {
            {
                let s = val.to_string();
                start.push_attribute(("cho", s.as_str()));
            }
        }
        if let Some(ref val) = self.che {
            {
                let s = val.to_string();
                start.push_attribute(("che", s.as_str()));
            }
        }
        if let Some(ref val) = self.chf {
            {
                let s = val.to_string();
                start.push_attribute(("chf", s.as_str()));
            }
        }
        if let Some(ref val) = self.rho {
            {
                let s = val.to_string();
                start.push_attribute(("rho", s.as_str()));
            }
        }
        if let Some(ref val) = self.rhe {
            {
                let s = val.to_string();
                start.push_attribute(("rhe", s.as_str()));
            }
        }
        if let Some(ref val) = self.rhf {
            {
                let s = val.to_string();
                start.push_attribute(("rhf", s.as_str()));
            }
        }
        if let Some(ref val) = self.lfo {
            {
                let s = val.to_string();
                start.push_attribute(("lfo", s.as_str()));
            }
        }
        if let Some(ref val) = self.lfe {
            {
                let s = val.to_string();
                start.push_attribute(("lfe", s.as_str()));
            }
        }
        if let Some(ref val) = self.lff {
            {
                let s = val.to_string();
                start.push_attribute(("lff", s.as_str()));
            }
        }
        if let Some(ref val) = self.cfo {
            {
                let s = val.to_string();
                start.push_attribute(("cfo", s.as_str()));
            }
        }
        if let Some(ref val) = self.cfe {
            {
                let s = val.to_string();
                start.push_attribute(("cfe", s.as_str()));
            }
        }
        if let Some(ref val) = self.cff {
            {
                let s = val.to_string();
                start.push_attribute(("cff", s.as_str()));
            }
        }
        if let Some(ref val) = self.rfo {
            {
                let s = val.to_string();
                start.push_attribute(("rfo", s.as_str()));
            }
        }
        if let Some(ref val) = self.rfe {
            {
                let s = val.to_string();
                start.push_attribute(("rfe", s.as_str()));
            }
        }
        if let Some(ref val) = self.rff {
            {
                let s = val.to_string();
                start.push_attribute(("rff", s.as_str()));
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

impl ToXml for CustomSheetViews {
    fn write_children<W: Write>(&self, writer: &mut Writer<W>) -> Result<(), SerializeError> {
        #[cfg(feature = "extra-children")]
        let mut extra_iter = self.extra_children.iter().peekable();
        #[cfg(feature = "extra-children")]
        let mut emit_idx: usize = 0;
        for item in &self.custom_sheet_view {
            #[cfg(feature = "extra-children")]
            while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
                extra_iter
                    .next()
                    .unwrap()
                    .node
                    .write_to(writer)
                    .map_err(SerializeError::from)?;
            }
            item.write_element("customSheetView", writer)?;
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
        if !self.custom_sheet_view.is_empty() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for CustomSheetView {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        {
            let val = &self.guid;
            start.push_attribute(("guid", val.as_str()));
        }
        if let Some(ref val) = self.scale {
            {
                let s = val.to_string();
                start.push_attribute(("scale", s.as_str()));
            }
        }
        if let Some(ref val) = self.color_id {
            {
                let s = val.to_string();
                start.push_attribute(("colorId", s.as_str()));
            }
        }
        if let Some(ref val) = self.show_page_breaks {
            start.push_attribute(("showPageBreaks", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.show_formulas {
            start.push_attribute(("showFormulas", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.show_grid_lines {
            start.push_attribute(("showGridLines", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.show_row_col {
            start.push_attribute(("showRowCol", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.outline_symbols {
            start.push_attribute(("outlineSymbols", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.zero_values {
            start.push_attribute(("zeroValues", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.fit_to_page {
            start.push_attribute(("fitToPage", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.print_area {
            start.push_attribute(("printArea", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.filter {
            start.push_attribute(("filter", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.show_auto_filter {
            start.push_attribute(("showAutoFilter", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.hidden_rows {
            start.push_attribute(("hiddenRows", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.hidden_columns {
            start.push_attribute(("hiddenColumns", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.state {
            {
                let s = val.to_string();
                start.push_attribute(("state", s.as_str()));
            }
        }
        if let Some(ref val) = self.filter_unique {
            start.push_attribute(("filterUnique", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.view {
            {
                let s = val.to_string();
                start.push_attribute(("view", s.as_str()));
            }
        }
        if let Some(ref val) = self.show_ruler {
            start.push_attribute(("showRuler", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.top_left_cell {
            start.push_attribute(("topLeftCell", val.as_str()));
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
        if let Some(ref val) = self.pane {
            val.write_element("pane", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        if let Some(ref val) = self.selection {
            val.write_element("selection", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        if let Some(ref val) = self.row_breaks {
            val.write_element("rowBreaks", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        if let Some(ref val) = self.col_breaks {
            val.write_element("colBreaks", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        if let Some(ref val) = self.page_margins {
            val.write_element("pageMargins", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        if let Some(ref val) = self.print_options {
            val.write_element("printOptions", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        if let Some(ref val) = self.page_setup {
            val.write_element("pageSetup", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        if let Some(ref val) = self.header_footer {
            val.write_element("headerFooter", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        if let Some(ref val) = self.auto_filter {
            val.write_element("autoFilter", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        if let Some(ref val) = self.extension_list {
            val.write_element("extLst", writer)?;
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
        if self.pane.is_some() {
            return false;
        }
        if self.selection.is_some() {
            return false;
        }
        if self.row_breaks.is_some() {
            return false;
        }
        if self.col_breaks.is_some() {
            return false;
        }
        if self.page_margins.is_some() {
            return false;
        }
        if self.print_options.is_some() {
            return false;
        }
        if self.page_setup.is_some() {
            return false;
        }
        if self.header_footer.is_some() {
            return false;
        }
        if self.auto_filter.is_some() {
            return false;
        }
        if self.extension_list.is_some() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for DataValidations {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        if let Some(ref val) = self.disable_prompts {
            start.push_attribute(("disablePrompts", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.x_window {
            {
                let s = val.to_string();
                start.push_attribute(("xWindow", s.as_str()));
            }
        }
        if let Some(ref val) = self.y_window {
            {
                let s = val.to_string();
                start.push_attribute(("yWindow", s.as_str()));
            }
        }
        if let Some(ref val) = self.count {
            {
                let s = val.to_string();
                start.push_attribute(("count", s.as_str()));
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
        for item in &self.data_validation {
            #[cfg(feature = "extra-children")]
            while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
                extra_iter
                    .next()
                    .unwrap()
                    .node
                    .write_to(writer)
                    .map_err(SerializeError::from)?;
            }
            item.write_element("dataValidation", writer)?;
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
        if !self.data_validation.is_empty() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for DataValidation {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        #[cfg(feature = "sml-validation")]
        if let Some(ref val) = self.r#type {
            {
                let s = val.to_string();
                start.push_attribute(("type", s.as_str()));
            }
        }
        #[cfg(feature = "sml-validation")]
        if let Some(ref val) = self.error_style {
            {
                let s = val.to_string();
                start.push_attribute(("errorStyle", s.as_str()));
            }
        }
        #[cfg(feature = "sml-validation")]
        if let Some(ref val) = self.ime_mode {
            {
                let s = val.to_string();
                start.push_attribute(("imeMode", s.as_str()));
            }
        }
        #[cfg(feature = "sml-validation")]
        if let Some(ref val) = self.operator {
            {
                let s = val.to_string();
                start.push_attribute(("operator", s.as_str()));
            }
        }
        #[cfg(feature = "sml-validation")]
        if let Some(ref val) = self.allow_blank {
            start.push_attribute(("allowBlank", if *val { "1" } else { "0" }));
        }
        #[cfg(feature = "sml-validation")]
        if let Some(ref val) = self.show_drop_down {
            start.push_attribute(("showDropDown", if *val { "1" } else { "0" }));
        }
        #[cfg(feature = "sml-validation")]
        if let Some(ref val) = self.show_input_message {
            start.push_attribute(("showInputMessage", if *val { "1" } else { "0" }));
        }
        #[cfg(feature = "sml-validation")]
        if let Some(ref val) = self.show_error_message {
            start.push_attribute(("showErrorMessage", if *val { "1" } else { "0" }));
        }
        #[cfg(feature = "sml-validation")]
        if let Some(ref val) = self.error_title {
            start.push_attribute(("errorTitle", val.as_str()));
        }
        #[cfg(feature = "sml-validation")]
        if let Some(ref val) = self.error {
            start.push_attribute(("error", val.as_str()));
        }
        #[cfg(feature = "sml-validation")]
        if let Some(ref val) = self.prompt_title {
            start.push_attribute(("promptTitle", val.as_str()));
        }
        #[cfg(feature = "sml-validation")]
        if let Some(ref val) = self.prompt {
            start.push_attribute(("prompt", val.as_str()));
        }
        #[cfg(feature = "sml-validation")]
        {
            let val = &self.square_reference;
            start.push_attribute(("sqref", val.as_str()));
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
        #[cfg(feature = "sml-validation")]
        if let Some(ref val) = self.formula1 {
            {
                let start = BytesStart::new("formula1");
                writer.write_event(Event::Start(start))?;
                writer.write_event(Event::Text(BytesText::new(val.as_str())))?;
                writer.write_event(Event::End(BytesEnd::new("formula1")))?;
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
        #[cfg(feature = "sml-validation")]
        if let Some(ref val) = self.formula2 {
            {
                let start = BytesStart::new("formula2");
                writer.write_event(Event::Start(start))?;
                writer.write_event(Event::Text(BytesText::new(val.as_str())))?;
                writer.write_event(Event::End(BytesEnd::new("formula2")))?;
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
        #[cfg(feature = "sml-validation")]
        if self.formula1.is_some() {
            return false;
        }
        #[cfg(feature = "sml-validation")]
        if self.formula2.is_some() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for ConditionalFormatting {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        #[cfg(feature = "sml-pivot")]
        if let Some(ref val) = self.pivot {
            start.push_attribute(("pivot", if *val { "1" } else { "0" }));
        }
        #[cfg(feature = "sml-styling")]
        if let Some(ref val) = self.square_reference {
            start.push_attribute(("sqref", val.as_str()));
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
        #[cfg(feature = "sml-styling")]
        for item in &self.cf_rule {
            #[cfg(feature = "extra-children")]
            while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
                extra_iter
                    .next()
                    .unwrap()
                    .node
                    .write_to(writer)
                    .map_err(SerializeError::from)?;
            }
            item.write_element("cfRule", writer)?;
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
        #[cfg(feature = "sml-extensions")]
        if let Some(ref val) = self.extension_list {
            val.write_element("extLst", writer)?;
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
        #[cfg(feature = "sml-styling")]
        if !self.cf_rule.is_empty() {
            return false;
        }
        #[cfg(feature = "sml-extensions")]
        if self.extension_list.is_some() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for ConditionalRule {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        #[cfg(feature = "sml-styling")]
        if let Some(ref val) = self.r#type {
            {
                let s = val.to_string();
                start.push_attribute(("type", s.as_str()));
            }
        }
        #[cfg(feature = "sml-styling")]
        if let Some(ref val) = self.dxf_id {
            {
                let s = val.to_string();
                start.push_attribute(("dxfId", s.as_str()));
            }
        }
        #[cfg(feature = "sml-styling")]
        {
            let val = &self.priority;
            {
                let s = val.to_string();
                start.push_attribute(("priority", s.as_str()));
            }
        }
        #[cfg(feature = "sml-styling")]
        if let Some(ref val) = self.stop_if_true {
            start.push_attribute(("stopIfTrue", if *val { "1" } else { "0" }));
        }
        #[cfg(feature = "sml-styling")]
        if let Some(ref val) = self.above_average {
            start.push_attribute(("aboveAverage", if *val { "1" } else { "0" }));
        }
        #[cfg(feature = "sml-styling")]
        if let Some(ref val) = self.percent {
            start.push_attribute(("percent", if *val { "1" } else { "0" }));
        }
        #[cfg(feature = "sml-styling")]
        if let Some(ref val) = self.bottom {
            start.push_attribute(("bottom", if *val { "1" } else { "0" }));
        }
        #[cfg(feature = "sml-styling")]
        if let Some(ref val) = self.operator {
            {
                let s = val.to_string();
                start.push_attribute(("operator", s.as_str()));
            }
        }
        #[cfg(feature = "sml-styling")]
        if let Some(ref val) = self.text {
            start.push_attribute(("text", val.as_str()));
        }
        #[cfg(feature = "sml-styling")]
        if let Some(ref val) = self.time_period {
            {
                let s = val.to_string();
                start.push_attribute(("timePeriod", s.as_str()));
            }
        }
        #[cfg(feature = "sml-styling")]
        if let Some(ref val) = self.rank {
            {
                let s = val.to_string();
                start.push_attribute(("rank", s.as_str()));
            }
        }
        #[cfg(feature = "sml-styling")]
        if let Some(ref val) = self.std_dev {
            {
                let s = val.to_string();
                start.push_attribute(("stdDev", s.as_str()));
            }
        }
        #[cfg(feature = "sml-styling")]
        if let Some(ref val) = self.equal_average {
            start.push_attribute(("equalAverage", if *val { "1" } else { "0" }));
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
        #[cfg(feature = "sml-styling")]
        for item in &self.formula {
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
                let start = BytesStart::new("formula");
                writer.write_event(Event::Start(start))?;
                writer.write_event(Event::Text(BytesText::new(item.as_str())))?;
                writer.write_event(Event::End(BytesEnd::new("formula")))?;
            }
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
        #[cfg(feature = "sml-styling")]
        if let Some(ref val) = self.color_scale {
            val.write_element("colorScale", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        #[cfg(feature = "sml-styling")]
        if let Some(ref val) = self.data_bar {
            val.write_element("dataBar", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        #[cfg(feature = "sml-styling")]
        if let Some(ref val) = self.icon_set {
            val.write_element("iconSet", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        #[cfg(feature = "sml-extensions")]
        if let Some(ref val) = self.extension_list {
            val.write_element("extLst", writer)?;
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
        #[cfg(feature = "sml-styling")]
        if !self.formula.is_empty() {
            return false;
        }
        #[cfg(feature = "sml-styling")]
        if self.color_scale.is_some() {
            return false;
        }
        #[cfg(feature = "sml-styling")]
        if self.data_bar.is_some() {
            return false;
        }
        #[cfg(feature = "sml-styling")]
        if self.icon_set.is_some() {
            return false;
        }
        #[cfg(feature = "sml-extensions")]
        if self.extension_list.is_some() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for Hyperlinks {
    fn write_children<W: Write>(&self, writer: &mut Writer<W>) -> Result<(), SerializeError> {
        #[cfg(feature = "extra-children")]
        let mut extra_iter = self.extra_children.iter().peekable();
        #[cfg(feature = "extra-children")]
        let mut emit_idx: usize = 0;
        for item in &self.hyperlink {
            #[cfg(feature = "extra-children")]
            while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
                extra_iter
                    .next()
                    .unwrap()
                    .node
                    .write_to(writer)
                    .map_err(SerializeError::from)?;
            }
            item.write_element("hyperlink", writer)?;
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
        if !self.hyperlink.is_empty() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for Hyperlink {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        #[cfg(feature = "sml-hyperlinks")]
        {
            let val = &self.reference;
            start.push_attribute(("ref", val.as_str()));
        }
        #[cfg(feature = "sml-hyperlinks")]
        if let Some(ref val) = self.location {
            start.push_attribute(("location", val.as_str()));
        }
        #[cfg(feature = "sml-hyperlinks")]
        if let Some(ref val) = self.tooltip {
            start.push_attribute(("tooltip", val.as_str()));
        }
        #[cfg(feature = "sml-hyperlinks")]
        if let Some(ref val) = self.display {
            start.push_attribute(("display", val.as_str()));
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

impl ToXml for CellFormula {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        if let Some(ref val) = self.cell_type {
            {
                let s = val.to_string();
                start.push_attribute(("t", s.as_str()));
            }
        }
        #[cfg(feature = "sml-formulas-advanced")]
        if let Some(ref val) = self.aca {
            start.push_attribute(("aca", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.reference {
            start.push_attribute(("ref", val.as_str()));
        }
        #[cfg(feature = "sml-formulas-advanced")]
        if let Some(ref val) = self.dt2_d {
            start.push_attribute(("dt2D", if *val { "1" } else { "0" }));
        }
        #[cfg(feature = "sml-formulas-advanced")]
        if let Some(ref val) = self.dtr {
            start.push_attribute(("dtr", if *val { "1" } else { "0" }));
        }
        #[cfg(feature = "sml-formulas-advanced")]
        if let Some(ref val) = self.del1 {
            start.push_attribute(("del1", if *val { "1" } else { "0" }));
        }
        #[cfg(feature = "sml-formulas-advanced")]
        if let Some(ref val) = self.del2 {
            start.push_attribute(("del2", if *val { "1" } else { "0" }));
        }
        #[cfg(feature = "sml-formulas-advanced")]
        if let Some(ref val) = self.r1 {
            start.push_attribute(("r1", val.as_str()));
        }
        #[cfg(feature = "sml-formulas-advanced")]
        if let Some(ref val) = self.r2 {
            start.push_attribute(("r2", val.as_str()));
        }
        #[cfg(feature = "sml-formulas-advanced")]
        if let Some(ref val) = self.ca {
            start.push_attribute(("ca", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.si {
            {
                let s = val.to_string();
                start.push_attribute(("si", s.as_str()));
            }
        }
        #[cfg(feature = "sml-formulas-advanced")]
        if let Some(ref val) = self.bx {
            start.push_attribute(("bx", if *val { "1" } else { "0" }));
        }
        #[cfg(feature = "extra-attrs")]
        for (key, value) in &self.extra_attrs {
            start.push_attribute((key.as_str(), value.as_str()));
        }
        start
    }

    fn write_children<W: Write>(&self, writer: &mut Writer<W>) -> Result<(), SerializeError> {
        if let Some(ref text) = self.text {
            writer.write_event(Event::Text(BytesText::new(text)))?;
        }
        #[cfg(feature = "extra-children")]
        for extra in &self.extra_children {
            extra.node.write_to(writer).map_err(SerializeError::from)?;
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

impl ToXml for ColorScale {
    fn write_children<W: Write>(&self, writer: &mut Writer<W>) -> Result<(), SerializeError> {
        #[cfg(feature = "extra-children")]
        let mut extra_iter = self.extra_children.iter().peekable();
        #[cfg(feature = "extra-children")]
        let mut emit_idx: usize = 0;
        for item in &self.cfvo {
            #[cfg(feature = "extra-children")]
            while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
                extra_iter
                    .next()
                    .unwrap()
                    .node
                    .write_to(writer)
                    .map_err(SerializeError::from)?;
            }
            item.write_element("cfvo", writer)?;
            #[cfg(feature = "extra-children")]
            {
                emit_idx += 1;
            }
        }
        for item in &self.color {
            #[cfg(feature = "extra-children")]
            while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
                extra_iter
                    .next()
                    .unwrap()
                    .node
                    .write_to(writer)
                    .map_err(SerializeError::from)?;
            }
            item.write_element("color", writer)?;
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
        if !self.cfvo.is_empty() {
            return false;
        }
        if !self.color.is_empty() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for DataBar {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        if let Some(ref val) = self.min_length {
            {
                let s = val.to_string();
                start.push_attribute(("minLength", s.as_str()));
            }
        }
        if let Some(ref val) = self.max_length {
            {
                let s = val.to_string();
                start.push_attribute(("maxLength", s.as_str()));
            }
        }
        if let Some(ref val) = self.show_value {
            start.push_attribute(("showValue", if *val { "1" } else { "0" }));
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
        for item in &self.cfvo {
            #[cfg(feature = "extra-children")]
            while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
                extra_iter
                    .next()
                    .unwrap()
                    .node
                    .write_to(writer)
                    .map_err(SerializeError::from)?;
            }
            item.write_element("cfvo", writer)?;
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
        {
            let val = &self.color;
            val.write_element("color", writer)?;
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
        if !self.cfvo.is_empty() {
            return false;
        }
        false
    }
}

impl ToXml for IconSet {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        if let Some(ref val) = self.icon_set {
            {
                let s = val.to_string();
                start.push_attribute(("iconSet", s.as_str()));
            }
        }
        if let Some(ref val) = self.show_value {
            start.push_attribute(("showValue", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.percent {
            start.push_attribute(("percent", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.reverse {
            start.push_attribute(("reverse", if *val { "1" } else { "0" }));
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
        for item in &self.cfvo {
            #[cfg(feature = "extra-children")]
            while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
                extra_iter
                    .next()
                    .unwrap()
                    .node
                    .write_to(writer)
                    .map_err(SerializeError::from)?;
            }
            item.write_element("cfvo", writer)?;
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
        if !self.cfvo.is_empty() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for ConditionalFormatValue {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        {
            let val = &self.r#type;
            {
                let s = val.to_string();
                start.push_attribute(("type", s.as_str()));
            }
        }
        if let Some(ref val) = self.value {
            start.push_attribute(("val", val.as_str()));
        }
        if let Some(ref val) = self.gte {
            start.push_attribute(("gte", if *val { "1" } else { "0" }));
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
        if let Some(ref val) = self.extension_list {
            val.write_element("extLst", writer)?;
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
        if self.extension_list.is_some() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for PageMargins {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        #[cfg(feature = "sml-layout")]
        {
            let val = &self.left;
            {
                let s = val.to_string();
                start.push_attribute(("left", s.as_str()));
            }
        }
        #[cfg(feature = "sml-layout")]
        {
            let val = &self.right;
            {
                let s = val.to_string();
                start.push_attribute(("right", s.as_str()));
            }
        }
        #[cfg(feature = "sml-layout")]
        {
            let val = &self.top;
            {
                let s = val.to_string();
                start.push_attribute(("top", s.as_str()));
            }
        }
        #[cfg(feature = "sml-layout")]
        {
            let val = &self.bottom;
            {
                let s = val.to_string();
                start.push_attribute(("bottom", s.as_str()));
            }
        }
        #[cfg(feature = "sml-layout")]
        {
            let val = &self.header;
            {
                let s = val.to_string();
                start.push_attribute(("header", s.as_str()));
            }
        }
        #[cfg(feature = "sml-layout")]
        {
            let val = &self.footer;
            {
                let s = val.to_string();
                start.push_attribute(("footer", s.as_str()));
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

impl ToXml for PrintOptions {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        if let Some(ref val) = self.horizontal_centered {
            start.push_attribute(("horizontalCentered", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.vertical_centered {
            start.push_attribute(("verticalCentered", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.headings {
            start.push_attribute(("headings", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.grid_lines {
            start.push_attribute(("gridLines", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.grid_lines_set {
            start.push_attribute(("gridLinesSet", if *val { "1" } else { "0" }));
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

impl ToXml for PageSetup {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        #[cfg(feature = "sml-layout")]
        if let Some(ref val) = self.paper_size {
            {
                let s = val.to_string();
                start.push_attribute(("paperSize", s.as_str()));
            }
        }
        #[cfg(feature = "sml-layout")]
        if let Some(ref val) = self.paper_height {
            start.push_attribute(("paperHeight", val.as_str()));
        }
        #[cfg(feature = "sml-layout")]
        if let Some(ref val) = self.paper_width {
            start.push_attribute(("paperWidth", val.as_str()));
        }
        #[cfg(feature = "sml-layout")]
        if let Some(ref val) = self.scale {
            {
                let s = val.to_string();
                start.push_attribute(("scale", s.as_str()));
            }
        }
        #[cfg(feature = "sml-layout")]
        if let Some(ref val) = self.first_page_number {
            {
                let s = val.to_string();
                start.push_attribute(("firstPageNumber", s.as_str()));
            }
        }
        #[cfg(feature = "sml-layout")]
        if let Some(ref val) = self.fit_to_width {
            {
                let s = val.to_string();
                start.push_attribute(("fitToWidth", s.as_str()));
            }
        }
        #[cfg(feature = "sml-layout")]
        if let Some(ref val) = self.fit_to_height {
            {
                let s = val.to_string();
                start.push_attribute(("fitToHeight", s.as_str()));
            }
        }
        #[cfg(feature = "sml-layout")]
        if let Some(ref val) = self.page_order {
            {
                let s = val.to_string();
                start.push_attribute(("pageOrder", s.as_str()));
            }
        }
        #[cfg(feature = "sml-layout")]
        if let Some(ref val) = self.orientation {
            {
                let s = val.to_string();
                start.push_attribute(("orientation", s.as_str()));
            }
        }
        #[cfg(feature = "sml-layout")]
        if let Some(ref val) = self.use_printer_defaults {
            start.push_attribute(("usePrinterDefaults", if *val { "1" } else { "0" }));
        }
        #[cfg(feature = "sml-layout")]
        if let Some(ref val) = self.black_and_white {
            start.push_attribute(("blackAndWhite", if *val { "1" } else { "0" }));
        }
        #[cfg(feature = "sml-layout")]
        if let Some(ref val) = self.draft {
            start.push_attribute(("draft", if *val { "1" } else { "0" }));
        }
        #[cfg(feature = "sml-layout")]
        if let Some(ref val) = self.cell_comments {
            {
                let s = val.to_string();
                start.push_attribute(("cellComments", s.as_str()));
            }
        }
        #[cfg(feature = "sml-layout")]
        if let Some(ref val) = self.use_first_page_number {
            start.push_attribute(("useFirstPageNumber", if *val { "1" } else { "0" }));
        }
        #[cfg(feature = "sml-layout")]
        if let Some(ref val) = self.errors {
            {
                let s = val.to_string();
                start.push_attribute(("errors", s.as_str()));
            }
        }
        #[cfg(feature = "sml-layout")]
        if let Some(ref val) = self.horizontal_dpi {
            {
                let s = val.to_string();
                start.push_attribute(("horizontalDpi", s.as_str()));
            }
        }
        #[cfg(feature = "sml-layout")]
        if let Some(ref val) = self.vertical_dpi {
            {
                let s = val.to_string();
                start.push_attribute(("verticalDpi", s.as_str()));
            }
        }
        #[cfg(feature = "sml-layout")]
        if let Some(ref val) = self.copies {
            {
                let s = val.to_string();
                start.push_attribute(("copies", s.as_str()));
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

impl ToXml for HeaderFooter {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        #[cfg(feature = "sml-layout")]
        if let Some(ref val) = self.different_odd_even {
            start.push_attribute(("differentOddEven", if *val { "1" } else { "0" }));
        }
        #[cfg(feature = "sml-layout")]
        if let Some(ref val) = self.different_first {
            start.push_attribute(("differentFirst", if *val { "1" } else { "0" }));
        }
        #[cfg(feature = "sml-layout")]
        if let Some(ref val) = self.scale_with_doc {
            start.push_attribute(("scaleWithDoc", if *val { "1" } else { "0" }));
        }
        #[cfg(feature = "sml-layout")]
        if let Some(ref val) = self.align_with_margins {
            start.push_attribute(("alignWithMargins", if *val { "1" } else { "0" }));
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
        #[cfg(feature = "sml-layout")]
        if let Some(ref val) = self.odd_header {
            {
                let start = BytesStart::new("oddHeader");
                writer.write_event(Event::Start(start))?;
                writer.write_event(Event::Text(BytesText::new(val.as_str())))?;
                writer.write_event(Event::End(BytesEnd::new("oddHeader")))?;
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
        #[cfg(feature = "sml-layout")]
        if let Some(ref val) = self.odd_footer {
            {
                let start = BytesStart::new("oddFooter");
                writer.write_event(Event::Start(start))?;
                writer.write_event(Event::Text(BytesText::new(val.as_str())))?;
                writer.write_event(Event::End(BytesEnd::new("oddFooter")))?;
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
        #[cfg(feature = "sml-layout")]
        if let Some(ref val) = self.even_header {
            {
                let start = BytesStart::new("evenHeader");
                writer.write_event(Event::Start(start))?;
                writer.write_event(Event::Text(BytesText::new(val.as_str())))?;
                writer.write_event(Event::End(BytesEnd::new("evenHeader")))?;
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
        #[cfg(feature = "sml-layout")]
        if let Some(ref val) = self.even_footer {
            {
                let start = BytesStart::new("evenFooter");
                writer.write_event(Event::Start(start))?;
                writer.write_event(Event::Text(BytesText::new(val.as_str())))?;
                writer.write_event(Event::End(BytesEnd::new("evenFooter")))?;
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
        #[cfg(feature = "sml-layout")]
        if let Some(ref val) = self.first_header {
            {
                let start = BytesStart::new("firstHeader");
                writer.write_event(Event::Start(start))?;
                writer.write_event(Event::Text(BytesText::new(val.as_str())))?;
                writer.write_event(Event::End(BytesEnd::new("firstHeader")))?;
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
        #[cfg(feature = "sml-layout")]
        if let Some(ref val) = self.first_footer {
            {
                let start = BytesStart::new("firstFooter");
                writer.write_event(Event::Start(start))?;
                writer.write_event(Event::Text(BytesText::new(val.as_str())))?;
                writer.write_event(Event::End(BytesEnd::new("firstFooter")))?;
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
        #[cfg(feature = "sml-layout")]
        if self.odd_header.is_some() {
            return false;
        }
        #[cfg(feature = "sml-layout")]
        if self.odd_footer.is_some() {
            return false;
        }
        #[cfg(feature = "sml-layout")]
        if self.even_header.is_some() {
            return false;
        }
        #[cfg(feature = "sml-layout")]
        if self.even_footer.is_some() {
            return false;
        }
        #[cfg(feature = "sml-layout")]
        if self.first_header.is_some() {
            return false;
        }
        #[cfg(feature = "sml-layout")]
        if self.first_footer.is_some() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for Scenarios {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        if let Some(ref val) = self.current {
            {
                let s = val.to_string();
                start.push_attribute(("current", s.as_str()));
            }
        }
        if let Some(ref val) = self.show {
            {
                let s = val.to_string();
                start.push_attribute(("show", s.as_str()));
            }
        }
        if let Some(ref val) = self.square_reference {
            start.push_attribute(("sqref", val.as_str()));
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
        for item in &self.scenario {
            #[cfg(feature = "extra-children")]
            while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
                extra_iter
                    .next()
                    .unwrap()
                    .node
                    .write_to(writer)
                    .map_err(SerializeError::from)?;
            }
            item.write_element("scenario", writer)?;
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
        if !self.scenario.is_empty() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for SheetProtection {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        if let Some(ref val) = self.password {
            {
                let hex = encode_hex(val);
                start.push_attribute(("password", hex.as_str()));
            }
        }
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
        if let Some(ref val) = self.spin_count {
            {
                let s = val.to_string();
                start.push_attribute(("spinCount", s.as_str()));
            }
        }
        if let Some(ref val) = self.sheet {
            start.push_attribute(("sheet", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.objects {
            start.push_attribute(("objects", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.scenarios {
            start.push_attribute(("scenarios", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.format_cells {
            start.push_attribute(("formatCells", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.format_columns {
            start.push_attribute(("formatColumns", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.format_rows {
            start.push_attribute(("formatRows", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.insert_columns {
            start.push_attribute(("insertColumns", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.insert_rows {
            start.push_attribute(("insertRows", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.insert_hyperlinks {
            start.push_attribute(("insertHyperlinks", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.delete_columns {
            start.push_attribute(("deleteColumns", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.delete_rows {
            start.push_attribute(("deleteRows", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.select_locked_cells {
            start.push_attribute(("selectLockedCells", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.sort {
            start.push_attribute(("sort", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.auto_filter {
            start.push_attribute(("autoFilter", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.pivot_tables {
            start.push_attribute(("pivotTables", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.select_unlocked_cells {
            start.push_attribute(("selectUnlockedCells", if *val { "1" } else { "0" }));
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

impl ToXml for ProtectedRanges {
    fn write_children<W: Write>(&self, writer: &mut Writer<W>) -> Result<(), SerializeError> {
        #[cfg(feature = "extra-children")]
        let mut extra_iter = self.extra_children.iter().peekable();
        #[cfg(feature = "extra-children")]
        let mut emit_idx: usize = 0;
        for item in &self.protected_range {
            #[cfg(feature = "extra-children")]
            while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
                extra_iter
                    .next()
                    .unwrap()
                    .node
                    .write_to(writer)
                    .map_err(SerializeError::from)?;
            }
            item.write_element("protectedRange", writer)?;
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
        if !self.protected_range.is_empty() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for ProtectedRange {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        if let Some(ref val) = self.password {
            {
                let hex = encode_hex(val);
                start.push_attribute(("password", hex.as_str()));
            }
        }
        {
            let val = &self.square_reference;
            start.push_attribute(("sqref", val.as_str()));
        }
        {
            let val = &self.name;
            start.push_attribute(("name", val.as_str()));
        }
        if let Some(ref val) = self.security_descriptor {
            start.push_attribute(("securityDescriptor", val.as_str()));
        }
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
        if let Some(ref val) = self.spin_count {
            {
                let s = val.to_string();
                start.push_attribute(("spinCount", s.as_str()));
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

impl ToXml for Scenario {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        {
            let val = &self.name;
            start.push_attribute(("name", val.as_str()));
        }
        if let Some(ref val) = self.locked {
            start.push_attribute(("locked", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.hidden {
            start.push_attribute(("hidden", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.count {
            {
                let s = val.to_string();
                start.push_attribute(("count", s.as_str()));
            }
        }
        if let Some(ref val) = self.user {
            start.push_attribute(("user", val.as_str()));
        }
        if let Some(ref val) = self.comment {
            start.push_attribute(("comment", val.as_str()));
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
        for item in &self.input_cells {
            #[cfg(feature = "extra-children")]
            while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
                extra_iter
                    .next()
                    .unwrap()
                    .node
                    .write_to(writer)
                    .map_err(SerializeError::from)?;
            }
            item.write_element("inputCells", writer)?;
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
        if !self.input_cells.is_empty() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for InputCells {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        {
            let val = &self.reference;
            start.push_attribute(("r", val.as_str()));
        }
        if let Some(ref val) = self.deleted {
            start.push_attribute(("deleted", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.undone {
            start.push_attribute(("undone", if *val { "1" } else { "0" }));
        }
        {
            let val = &self.value;
            start.push_attribute(("val", val.as_str()));
        }
        if let Some(ref val) = self.number_format_id {
            {
                let s = val.to_string();
                start.push_attribute(("numFmtId", s.as_str()));
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

impl ToXml for CellWatches {
    fn write_children<W: Write>(&self, writer: &mut Writer<W>) -> Result<(), SerializeError> {
        #[cfg(feature = "extra-children")]
        let mut extra_iter = self.extra_children.iter().peekable();
        #[cfg(feature = "extra-children")]
        let mut emit_idx: usize = 0;
        for item in &self.cell_watch {
            #[cfg(feature = "extra-children")]
            while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
                extra_iter
                    .next()
                    .unwrap()
                    .node
                    .write_to(writer)
                    .map_err(SerializeError::from)?;
            }
            item.write_element("cellWatch", writer)?;
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
        if !self.cell_watch.is_empty() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for CellWatch {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        {
            let val = &self.reference;
            start.push_attribute(("r", val.as_str()));
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

impl ToXml for Chartsheet {
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
        if let Some(ref val) = self.sheet_properties {
            val.write_element("sheetPr", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
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
            let val = &self.sheet_views;
            val.write_element("sheetViews", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        if let Some(ref val) = self.sheet_protection {
            val.write_element("sheetProtection", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        if let Some(ref val) = self.custom_sheet_views {
            val.write_element("customSheetViews", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        if let Some(ref val) = self.page_margins {
            val.write_element("pageMargins", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        if let Some(ref val) = self.page_setup {
            val.write_element("pageSetup", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        if let Some(ref val) = self.header_footer {
            val.write_element("headerFooter", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
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
            let val = &self.drawing;
            val.write_element("drawing", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        if let Some(ref val) = self.legacy_drawing {
            val.write_element("legacyDrawing", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        if let Some(ref val) = self.legacy_drawing_h_f {
            val.write_element("legacyDrawingHF", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        if let Some(ref val) = self.drawing_h_f {
            val.write_element("drawingHF", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
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
            val.write_element("picture", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        if let Some(ref val) = self.web_publish_items {
            val.write_element("webPublishItems", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        if let Some(ref val) = self.extension_list {
            val.write_element("extLst", writer)?;
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
        if self.sheet_properties.is_some() {
            return false;
        }
        false
    }
}

impl ToXml for ChartsheetProperties {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        if let Some(ref val) = self.published {
            start.push_attribute(("published", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.code_name {
            start.push_attribute(("codeName", val.as_str()));
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
        if let Some(ref val) = self.tab_color {
            val.write_element("tabColor", writer)?;
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
        if self.tab_color.is_some() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for ChartsheetViews {
    fn write_children<W: Write>(&self, writer: &mut Writer<W>) -> Result<(), SerializeError> {
        #[cfg(feature = "extra-children")]
        let mut extra_iter = self.extra_children.iter().peekable();
        #[cfg(feature = "extra-children")]
        let mut emit_idx: usize = 0;
        for item in &self.sheet_view {
            #[cfg(feature = "extra-children")]
            while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
                extra_iter
                    .next()
                    .unwrap()
                    .node
                    .write_to(writer)
                    .map_err(SerializeError::from)?;
            }
            item.write_element("sheetView", writer)?;
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
        if let Some(ref val) = self.extension_list {
            val.write_element("extLst", writer)?;
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
        if !self.sheet_view.is_empty() {
            return false;
        }
        if self.extension_list.is_some() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for ChartsheetView {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        if let Some(ref val) = self.tab_selected {
            start.push_attribute(("tabSelected", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.zoom_scale {
            {
                let s = val.to_string();
                start.push_attribute(("zoomScale", s.as_str()));
            }
        }
        {
            let val = &self.workbook_view_id;
            {
                let s = val.to_string();
                start.push_attribute(("workbookViewId", s.as_str()));
            }
        }
        if let Some(ref val) = self.zoom_to_fit {
            start.push_attribute(("zoomToFit", if *val { "1" } else { "0" }));
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
        if let Some(ref val) = self.extension_list {
            val.write_element("extLst", writer)?;
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
        if self.extension_list.is_some() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for ChartsheetProtection {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        if let Some(ref val) = self.password {
            {
                let hex = encode_hex(val);
                start.push_attribute(("password", hex.as_str()));
            }
        }
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
        if let Some(ref val) = self.spin_count {
            {
                let s = val.to_string();
                start.push_attribute(("spinCount", s.as_str()));
            }
        }
        if let Some(ref val) = self.content {
            start.push_attribute(("content", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.objects {
            start.push_attribute(("objects", if *val { "1" } else { "0" }));
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

impl ToXml for ChartsheetPageSetup {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        if let Some(ref val) = self.paper_size {
            {
                let s = val.to_string();
                start.push_attribute(("paperSize", s.as_str()));
            }
        }
        if let Some(ref val) = self.paper_height {
            start.push_attribute(("paperHeight", val.as_str()));
        }
        if let Some(ref val) = self.paper_width {
            start.push_attribute(("paperWidth", val.as_str()));
        }
        if let Some(ref val) = self.first_page_number {
            {
                let s = val.to_string();
                start.push_attribute(("firstPageNumber", s.as_str()));
            }
        }
        if let Some(ref val) = self.orientation {
            {
                let s = val.to_string();
                start.push_attribute(("orientation", s.as_str()));
            }
        }
        if let Some(ref val) = self.use_printer_defaults {
            start.push_attribute(("usePrinterDefaults", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.black_and_white {
            start.push_attribute(("blackAndWhite", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.draft {
            start.push_attribute(("draft", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.use_first_page_number {
            start.push_attribute(("useFirstPageNumber", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.horizontal_dpi {
            {
                let s = val.to_string();
                start.push_attribute(("horizontalDpi", s.as_str()));
            }
        }
        if let Some(ref val) = self.vertical_dpi {
            {
                let s = val.to_string();
                start.push_attribute(("verticalDpi", s.as_str()));
            }
        }
        if let Some(ref val) = self.copies {
            {
                let s = val.to_string();
                start.push_attribute(("copies", s.as_str()));
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

impl ToXml for CustomChartsheetViews {
    fn write_children<W: Write>(&self, writer: &mut Writer<W>) -> Result<(), SerializeError> {
        #[cfg(feature = "extra-children")]
        let mut extra_iter = self.extra_children.iter().peekable();
        #[cfg(feature = "extra-children")]
        let mut emit_idx: usize = 0;
        for item in &self.custom_sheet_view {
            #[cfg(feature = "extra-children")]
            while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
                extra_iter
                    .next()
                    .unwrap()
                    .node
                    .write_to(writer)
                    .map_err(SerializeError::from)?;
            }
            item.write_element("customSheetView", writer)?;
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
        if !self.custom_sheet_view.is_empty() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for CustomChartsheetView {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        {
            let val = &self.guid;
            start.push_attribute(("guid", val.as_str()));
        }
        if let Some(ref val) = self.scale {
            {
                let s = val.to_string();
                start.push_attribute(("scale", s.as_str()));
            }
        }
        if let Some(ref val) = self.state {
            {
                let s = val.to_string();
                start.push_attribute(("state", s.as_str()));
            }
        }
        if let Some(ref val) = self.zoom_to_fit {
            start.push_attribute(("zoomToFit", if *val { "1" } else { "0" }));
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
        if let Some(ref val) = self.page_margins {
            val.write_element("pageMargins", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        if let Some(ref val) = self.page_setup {
            val.write_element("pageSetup", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        if let Some(ref val) = self.header_footer {
            val.write_element("headerFooter", writer)?;
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
        if self.page_margins.is_some() {
            return false;
        }
        if self.page_setup.is_some() {
            return false;
        }
        if self.header_footer.is_some() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for CTCustomProperties {
    fn write_children<W: Write>(&self, writer: &mut Writer<W>) -> Result<(), SerializeError> {
        #[cfg(feature = "extra-children")]
        let mut extra_iter = self.extra_children.iter().peekable();
        #[cfg(feature = "extra-children")]
        let mut emit_idx: usize = 0;
        for item in &self.custom_pr {
            #[cfg(feature = "extra-children")]
            while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
                extra_iter
                    .next()
                    .unwrap()
                    .node
                    .write_to(writer)
                    .map_err(SerializeError::from)?;
            }
            item.write_element("customPr", writer)?;
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
        if !self.custom_pr.is_empty() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for CTCustomProperty {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
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

    fn is_empty_element(&self) -> bool {
        true
    }
}

impl ToXml for OleObjects {
    fn write_children<W: Write>(&self, writer: &mut Writer<W>) -> Result<(), SerializeError> {
        #[cfg(feature = "extra-children")]
        let mut extra_iter = self.extra_children.iter().peekable();
        #[cfg(feature = "extra-children")]
        let mut emit_idx: usize = 0;
        for item in &self.ole_object {
            #[cfg(feature = "extra-children")]
            while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
                extra_iter
                    .next()
                    .unwrap()
                    .node
                    .write_to(writer)
                    .map_err(SerializeError::from)?;
            }
            item.write_element("oleObject", writer)?;
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
        if !self.ole_object.is_empty() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for OleObject {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        if let Some(ref val) = self.prog_id {
            start.push_attribute(("progId", val.as_str()));
        }
        if let Some(ref val) = self.dv_aspect {
            {
                let s = val.to_string();
                start.push_attribute(("dvAspect", s.as_str()));
            }
        }
        if let Some(ref val) = self.link {
            start.push_attribute(("link", val.as_str()));
        }
        if let Some(ref val) = self.ole_update {
            {
                let s = val.to_string();
                start.push_attribute(("oleUpdate", s.as_str()));
            }
        }
        if let Some(ref val) = self.auto_load {
            start.push_attribute(("autoLoad", if *val { "1" } else { "0" }));
        }
        {
            let val = &self.shape_id;
            {
                let s = val.to_string();
                start.push_attribute(("shapeId", s.as_str()));
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
        if let Some(ref val) = self.object_pr {
            val.write_element("objectPr", writer)?;
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
        if self.object_pr.is_some() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for ObjectProperties {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        if let Some(ref val) = self.locked {
            start.push_attribute(("locked", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.default_size {
            start.push_attribute(("defaultSize", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.print {
            start.push_attribute(("print", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.disabled {
            start.push_attribute(("disabled", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.ui_object {
            start.push_attribute(("uiObject", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.auto_fill {
            start.push_attribute(("autoFill", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.auto_line {
            start.push_attribute(("autoLine", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.auto_pict {
            start.push_attribute(("autoPict", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.r#macro {
            start.push_attribute(("macro", val.as_str()));
        }
        if let Some(ref val) = self.alt_text {
            start.push_attribute(("altText", val.as_str()));
        }
        if let Some(ref val) = self.dde {
            start.push_attribute(("dde", if *val { "1" } else { "0" }));
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
            let val = &self.anchor;
            val.write_element("anchor", writer)?;
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

impl ToXml for WebPublishItems {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        if let Some(ref val) = self.count {
            {
                let s = val.to_string();
                start.push_attribute(("count", s.as_str()));
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
        for item in &self.web_publish_item {
            #[cfg(feature = "extra-children")]
            while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
                extra_iter
                    .next()
                    .unwrap()
                    .node
                    .write_to(writer)
                    .map_err(SerializeError::from)?;
            }
            item.write_element("webPublishItem", writer)?;
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
        if !self.web_publish_item.is_empty() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for WebPublishItem {
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
            let val = &self.div_id;
            start.push_attribute(("divId", val.as_str()));
        }
        {
            let val = &self.source_type;
            {
                let s = val.to_string();
                start.push_attribute(("sourceType", s.as_str()));
            }
        }
        if let Some(ref val) = self.source_ref {
            start.push_attribute(("sourceRef", val.as_str()));
        }
        if let Some(ref val) = self.source_object {
            start.push_attribute(("sourceObject", val.as_str()));
        }
        {
            let val = &self.destination_file;
            start.push_attribute(("destinationFile", val.as_str()));
        }
        if let Some(ref val) = self.title {
            start.push_attribute(("title", val.as_str()));
        }
        if let Some(ref val) = self.auto_republish {
            start.push_attribute(("autoRepublish", if *val { "1" } else { "0" }));
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

impl ToXml for Controls {
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
            item.write_element("control", writer)?;
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

impl ToXml for Control {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        {
            let val = &self.shape_id;
            {
                let s = val.to_string();
                start.push_attribute(("shapeId", s.as_str()));
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
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        if let Some(ref val) = self.control_pr {
            val.write_element("controlPr", writer)?;
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
        if self.control_pr.is_some() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for CTControlPr {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        if let Some(ref val) = self.locked {
            start.push_attribute(("locked", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.default_size {
            start.push_attribute(("defaultSize", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.print {
            start.push_attribute(("print", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.disabled {
            start.push_attribute(("disabled", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.recalc_always {
            start.push_attribute(("recalcAlways", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.ui_object {
            start.push_attribute(("uiObject", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.auto_fill {
            start.push_attribute(("autoFill", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.auto_line {
            start.push_attribute(("autoLine", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.auto_pict {
            start.push_attribute(("autoPict", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.r#macro {
            start.push_attribute(("macro", val.as_str()));
        }
        if let Some(ref val) = self.alt_text {
            start.push_attribute(("altText", val.as_str()));
        }
        if let Some(ref val) = self.linked_cell {
            start.push_attribute(("linkedCell", val.as_str()));
        }
        if let Some(ref val) = self.list_fill_range {
            start.push_attribute(("listFillRange", val.as_str()));
        }
        if let Some(ref val) = self.cf {
            start.push_attribute(("cf", val.as_str()));
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
            let val = &self.anchor;
            val.write_element("anchor", writer)?;
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

impl ToXml for IgnoredErrors {
    fn write_children<W: Write>(&self, writer: &mut Writer<W>) -> Result<(), SerializeError> {
        #[cfg(feature = "extra-children")]
        let mut extra_iter = self.extra_children.iter().peekable();
        #[cfg(feature = "extra-children")]
        let mut emit_idx: usize = 0;
        for item in &self.ignored_error {
            #[cfg(feature = "extra-children")]
            while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
                extra_iter
                    .next()
                    .unwrap()
                    .node
                    .write_to(writer)
                    .map_err(SerializeError::from)?;
            }
            item.write_element("ignoredError", writer)?;
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
        if let Some(ref val) = self.extension_list {
            val.write_element("extLst", writer)?;
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
        if !self.ignored_error.is_empty() {
            return false;
        }
        if self.extension_list.is_some() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for IgnoredError {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        {
            let val = &self.square_reference;
            start.push_attribute(("sqref", val.as_str()));
        }
        if let Some(ref val) = self.eval_error {
            start.push_attribute(("evalError", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.two_digit_text_year {
            start.push_attribute(("twoDigitTextYear", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.number_stored_as_text {
            start.push_attribute(("numberStoredAsText", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.formula {
            start.push_attribute(("formula", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.formula_range {
            start.push_attribute(("formulaRange", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.unlocked_formula {
            start.push_attribute(("unlockedFormula", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.empty_cell_reference {
            start.push_attribute(("emptyCellReference", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.list_data_validation {
            start.push_attribute(("listDataValidation", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.calculated_column {
            start.push_attribute(("calculatedColumn", if *val { "1" } else { "0" }));
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

impl ToXml for TableParts {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        if let Some(ref val) = self.count {
            {
                let s = val.to_string();
                start.push_attribute(("count", s.as_str()));
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
        for item in &self.table_part {
            #[cfg(feature = "extra-children")]
            while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
                extra_iter
                    .next()
                    .unwrap()
                    .node
                    .write_to(writer)
                    .map_err(SerializeError::from)?;
            }
            item.write_element("tablePart", writer)?;
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
        if !self.table_part.is_empty() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for TablePart {
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

impl ToXml for Metadata {
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
        if let Some(ref val) = self.metadata_types {
            val.write_element("metadataTypes", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        if let Some(ref val) = self.metadata_strings {
            val.write_element("metadataStrings", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        if let Some(ref val) = self.mdx_metadata {
            val.write_element("mdxMetadata", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        for item in &self.future_metadata {
            #[cfg(feature = "extra-children")]
            while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
                extra_iter
                    .next()
                    .unwrap()
                    .node
                    .write_to(writer)
                    .map_err(SerializeError::from)?;
            }
            item.write_element("futureMetadata", writer)?;
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
        if let Some(ref val) = self.cell_metadata {
            val.write_element("cellMetadata", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        if let Some(ref val) = self.value_metadata {
            val.write_element("valueMetadata", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        if let Some(ref val) = self.extension_list {
            val.write_element("extLst", writer)?;
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
        if self.metadata_types.is_some() {
            return false;
        }
        if self.metadata_strings.is_some() {
            return false;
        }
        if self.mdx_metadata.is_some() {
            return false;
        }
        if !self.future_metadata.is_empty() {
            return false;
        }
        if self.cell_metadata.is_some() {
            return false;
        }
        if self.value_metadata.is_some() {
            return false;
        }
        if self.extension_list.is_some() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for MetadataTypes {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        if let Some(ref val) = self.count {
            {
                let s = val.to_string();
                start.push_attribute(("count", s.as_str()));
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
        for item in &self.metadata_type {
            #[cfg(feature = "extra-children")]
            while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
                extra_iter
                    .next()
                    .unwrap()
                    .node
                    .write_to(writer)
                    .map_err(SerializeError::from)?;
            }
            item.write_element("metadataType", writer)?;
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
        if !self.metadata_type.is_empty() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for MetadataType {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        {
            let val = &self.name;
            start.push_attribute(("name", val.as_str()));
        }
        {
            let val = &self.min_supported_version;
            {
                let s = val.to_string();
                start.push_attribute(("minSupportedVersion", s.as_str()));
            }
        }
        if let Some(ref val) = self.ghost_row {
            start.push_attribute(("ghostRow", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.ghost_col {
            start.push_attribute(("ghostCol", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.edit {
            start.push_attribute(("edit", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.delete {
            start.push_attribute(("delete", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.copy {
            start.push_attribute(("copy", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.paste_all {
            start.push_attribute(("pasteAll", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.paste_formulas {
            start.push_attribute(("pasteFormulas", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.paste_values {
            start.push_attribute(("pasteValues", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.paste_formats {
            start.push_attribute(("pasteFormats", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.paste_comments {
            start.push_attribute(("pasteComments", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.paste_data_validation {
            start.push_attribute(("pasteDataValidation", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.paste_borders {
            start.push_attribute(("pasteBorders", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.paste_col_widths {
            start.push_attribute(("pasteColWidths", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.paste_number_formats {
            start.push_attribute(("pasteNumberFormats", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.merge {
            start.push_attribute(("merge", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.split_first {
            start.push_attribute(("splitFirst", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.split_all {
            start.push_attribute(("splitAll", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.row_col_shift {
            start.push_attribute(("rowColShift", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.clear_all {
            start.push_attribute(("clearAll", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.clear_formats {
            start.push_attribute(("clearFormats", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.clear_contents {
            start.push_attribute(("clearContents", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.clear_comments {
            start.push_attribute(("clearComments", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.assign {
            start.push_attribute(("assign", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.coerce {
            start.push_attribute(("coerce", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.adjust {
            start.push_attribute(("adjust", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.cell_meta {
            start.push_attribute(("cellMeta", if *val { "1" } else { "0" }));
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

impl ToXml for MetadataBlocks {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        if let Some(ref val) = self.count {
            {
                let s = val.to_string();
                start.push_attribute(("count", s.as_str()));
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
        for item in &self.bk {
            #[cfg(feature = "extra-children")]
            while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
                extra_iter
                    .next()
                    .unwrap()
                    .node
                    .write_to(writer)
                    .map_err(SerializeError::from)?;
            }
            item.write_element("bk", writer)?;
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
        if !self.bk.is_empty() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for MetadataBlock {
    fn write_children<W: Write>(&self, writer: &mut Writer<W>) -> Result<(), SerializeError> {
        #[cfg(feature = "extra-children")]
        let mut extra_iter = self.extra_children.iter().peekable();
        #[cfg(feature = "extra-children")]
        let mut emit_idx: usize = 0;
        for item in &self.rc {
            #[cfg(feature = "extra-children")]
            while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
                extra_iter
                    .next()
                    .unwrap()
                    .node
                    .write_to(writer)
                    .map_err(SerializeError::from)?;
            }
            item.write_element("rc", writer)?;
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
        if !self.rc.is_empty() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for MetadataRecord {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        {
            let val = &self.cell_type;
            {
                let s = val.to_string();
                start.push_attribute(("t", s.as_str()));
            }
        }
        {
            let val = &self.value;
            {
                let s = val.to_string();
                start.push_attribute(("v", s.as_str()));
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

impl ToXml for CTFutureMetadata {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        {
            let val = &self.name;
            start.push_attribute(("name", val.as_str()));
        }
        if let Some(ref val) = self.count {
            {
                let s = val.to_string();
                start.push_attribute(("count", s.as_str()));
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
        for item in &self.bk {
            #[cfg(feature = "extra-children")]
            while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
                extra_iter
                    .next()
                    .unwrap()
                    .node
                    .write_to(writer)
                    .map_err(SerializeError::from)?;
            }
            item.write_element("bk", writer)?;
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
        if let Some(ref val) = self.extension_list {
            val.write_element("extLst", writer)?;
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
        if !self.bk.is_empty() {
            return false;
        }
        if self.extension_list.is_some() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for CTFutureMetadataBlock {
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
        if let Some(ref val) = self.extension_list {
            val.write_element("extLst", writer)?;
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
        if self.extension_list.is_some() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for CTMdxMetadata {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        if let Some(ref val) = self.count {
            {
                let s = val.to_string();
                start.push_attribute(("count", s.as_str()));
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
        for item in &self.mdx {
            #[cfg(feature = "extra-children")]
            while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
                extra_iter
                    .next()
                    .unwrap()
                    .node
                    .write_to(writer)
                    .map_err(SerializeError::from)?;
            }
            item.write_element("mdx", writer)?;
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
        if !self.mdx.is_empty() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for CTMdx {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        {
            let val = &self.n;
            {
                let s = val.to_string();
                start.push_attribute(("n", s.as_str()));
            }
        }
        {
            let val = &self.formula;
            {
                let s = val.to_string();
                start.push_attribute(("f", s.as_str()));
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

impl ToXml for CTMdxTuple {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        if let Some(ref val) = self.cells {
            {
                let s = val.to_string();
                start.push_attribute(("c", s.as_str()));
            }
        }
        if let Some(ref val) = self.ct {
            start.push_attribute(("ct", val.as_str()));
        }
        if let Some(ref val) = self.si {
            {
                let s = val.to_string();
                start.push_attribute(("si", s.as_str()));
            }
        }
        if let Some(ref val) = self.fi {
            {
                let s = val.to_string();
                start.push_attribute(("fi", s.as_str()));
            }
        }
        if let Some(ref val) = self.bc {
            {
                let hex = encode_hex(val);
                start.push_attribute(("bc", hex.as_str()));
            }
        }
        if let Some(ref val) = self.fc {
            {
                let hex = encode_hex(val);
                start.push_attribute(("fc", hex.as_str()));
            }
        }
        if let Some(ref val) = self.i {
            start.push_attribute(("i", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.u {
            start.push_attribute(("u", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.st {
            start.push_attribute(("st", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.b {
            start.push_attribute(("b", if *val { "1" } else { "0" }));
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
        for item in &self.n {
            #[cfg(feature = "extra-children")]
            while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
                extra_iter
                    .next()
                    .unwrap()
                    .node
                    .write_to(writer)
                    .map_err(SerializeError::from)?;
            }
            item.write_element("n", writer)?;
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
        if !self.n.is_empty() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for CTMdxSet {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        {
            let val = &self.ns;
            {
                let s = val.to_string();
                start.push_attribute(("ns", s.as_str()));
            }
        }
        if let Some(ref val) = self.cells {
            {
                let s = val.to_string();
                start.push_attribute(("c", s.as_str()));
            }
        }
        if let Some(ref val) = self.o {
            {
                let s = val.to_string();
                start.push_attribute(("o", s.as_str()));
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
        for item in &self.n {
            #[cfg(feature = "extra-children")]
            while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
                extra_iter
                    .next()
                    .unwrap()
                    .node
                    .write_to(writer)
                    .map_err(SerializeError::from)?;
            }
            item.write_element("n", writer)?;
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
        if !self.n.is_empty() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for CTMdxMemeberProp {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        {
            let val = &self.n;
            {
                let s = val.to_string();
                start.push_attribute(("n", s.as_str()));
            }
        }
        {
            let val = &self.np;
            {
                let s = val.to_string();
                start.push_attribute(("np", s.as_str()));
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

impl ToXml for CTMdxKPI {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        {
            let val = &self.n;
            {
                let s = val.to_string();
                start.push_attribute(("n", s.as_str()));
            }
        }
        {
            let val = &self.np;
            {
                let s = val.to_string();
                start.push_attribute(("np", s.as_str()));
            }
        }
        {
            let val = &self.p;
            {
                let s = val.to_string();
                start.push_attribute(("p", s.as_str()));
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

impl ToXml for CTMetadataStringIndex {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        {
            let val = &self.x;
            {
                let s = val.to_string();
                start.push_attribute(("x", s.as_str()));
            }
        }
        if let Some(ref val) = self.style_index {
            start.push_attribute(("s", if *val { "1" } else { "0" }));
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

impl ToXml for MetadataStrings {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        if let Some(ref val) = self.count {
            {
                let s = val.to_string();
                start.push_attribute(("count", s.as_str()));
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
        for item in &self.style_index {
            #[cfg(feature = "extra-children")]
            while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
                extra_iter
                    .next()
                    .unwrap()
                    .node
                    .write_to(writer)
                    .map_err(SerializeError::from)?;
            }
            item.write_element("s", writer)?;
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
        if !self.style_index.is_empty() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for SingleXmlCells {
    fn write_children<W: Write>(&self, writer: &mut Writer<W>) -> Result<(), SerializeError> {
        #[cfg(feature = "extra-children")]
        let mut extra_iter = self.extra_children.iter().peekable();
        #[cfg(feature = "extra-children")]
        let mut emit_idx: usize = 0;
        for item in &self.single_xml_cell {
            #[cfg(feature = "extra-children")]
            while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
                extra_iter
                    .next()
                    .unwrap()
                    .node
                    .write_to(writer)
                    .map_err(SerializeError::from)?;
            }
            item.write_element("singleXmlCell", writer)?;
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
        if !self.single_xml_cell.is_empty() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for SingleXmlCell {
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
            let val = &self.reference;
            start.push_attribute(("r", val.as_str()));
        }
        {
            let val = &self.connection_id;
            {
                let s = val.to_string();
                start.push_attribute(("connectionId", s.as_str()));
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
            let val = &self.xml_cell_pr;
            val.write_element("xmlCellPr", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        if let Some(ref val) = self.extension_list {
            val.write_element("extLst", writer)?;
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

impl ToXml for XmlCellProperties {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        {
            let val = &self.id;
            {
                let s = val.to_string();
                start.push_attribute(("id", s.as_str()));
            }
        }
        if let Some(ref val) = self.unique_name {
            start.push_attribute(("uniqueName", val.as_str()));
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
            let val = &self.xml_pr;
            val.write_element("xmlPr", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        if let Some(ref val) = self.extension_list {
            val.write_element("extLst", writer)?;
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

impl ToXml for XmlProperties {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        {
            let val = &self.map_id;
            {
                let s = val.to_string();
                start.push_attribute(("mapId", s.as_str()));
            }
        }
        {
            let val = &self.xpath;
            start.push_attribute(("xpath", val.as_str()));
        }
        {
            let val = &self.xml_data_type;
            start.push_attribute(("xmlDataType", val.as_str()));
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
        if let Some(ref val) = self.extension_list {
            val.write_element("extLst", writer)?;
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
        if self.extension_list.is_some() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for Stylesheet {
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
        #[cfg(feature = "sml-styling")]
        if let Some(ref val) = self.num_fmts {
            val.write_element("numFmts", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        #[cfg(feature = "sml-styling")]
        if let Some(ref val) = self.fonts {
            val.write_element("fonts", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        #[cfg(feature = "sml-styling")]
        if let Some(ref val) = self.fills {
            val.write_element("fills", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        #[cfg(feature = "sml-styling")]
        if let Some(ref val) = self.borders {
            val.write_element("borders", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        #[cfg(feature = "sml-styling")]
        if let Some(ref val) = self.cell_style_xfs {
            val.write_element("cellStyleXfs", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        #[cfg(feature = "sml-styling")]
        if let Some(ref val) = self.cell_xfs {
            val.write_element("cellXfs", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        #[cfg(feature = "sml-styling")]
        if let Some(ref val) = self.cell_styles {
            val.write_element("cellStyles", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        #[cfg(feature = "sml-styling")]
        if let Some(ref val) = self.dxfs {
            val.write_element("dxfs", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        #[cfg(feature = "sml-styling")]
        if let Some(ref val) = self.table_styles {
            val.write_element("tableStyles", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        #[cfg(feature = "sml-styling")]
        if let Some(ref val) = self.colors {
            val.write_element("colors", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        #[cfg(feature = "sml-extensions")]
        if let Some(ref val) = self.extension_list {
            val.write_element("extLst", writer)?;
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
        #[cfg(feature = "sml-styling")]
        if self.num_fmts.is_some() {
            return false;
        }
        #[cfg(feature = "sml-styling")]
        if self.fonts.is_some() {
            return false;
        }
        #[cfg(feature = "sml-styling")]
        if self.fills.is_some() {
            return false;
        }
        #[cfg(feature = "sml-styling")]
        if self.borders.is_some() {
            return false;
        }
        #[cfg(feature = "sml-styling")]
        if self.cell_style_xfs.is_some() {
            return false;
        }
        #[cfg(feature = "sml-styling")]
        if self.cell_xfs.is_some() {
            return false;
        }
        #[cfg(feature = "sml-styling")]
        if self.cell_styles.is_some() {
            return false;
        }
        #[cfg(feature = "sml-styling")]
        if self.dxfs.is_some() {
            return false;
        }
        #[cfg(feature = "sml-styling")]
        if self.table_styles.is_some() {
            return false;
        }
        #[cfg(feature = "sml-styling")]
        if self.colors.is_some() {
            return false;
        }
        #[cfg(feature = "sml-extensions")]
        if self.extension_list.is_some() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for CellAlignment {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        if let Some(ref val) = self.horizontal {
            {
                let s = val.to_string();
                start.push_attribute(("horizontal", s.as_str()));
            }
        }
        if let Some(ref val) = self.vertical {
            {
                let s = val.to_string();
                start.push_attribute(("vertical", s.as_str()));
            }
        }
        if let Some(ref val) = self.text_rotation {
            {
                let s = val.to_string();
                start.push_attribute(("textRotation", s.as_str()));
            }
        }
        if let Some(ref val) = self.wrap_text {
            start.push_attribute(("wrapText", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.indent {
            {
                let s = val.to_string();
                start.push_attribute(("indent", s.as_str()));
            }
        }
        if let Some(ref val) = self.relative_indent {
            {
                let s = val.to_string();
                start.push_attribute(("relativeIndent", s.as_str()));
            }
        }
        if let Some(ref val) = self.justify_last_line {
            start.push_attribute(("justifyLastLine", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.shrink_to_fit {
            start.push_attribute(("shrinkToFit", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.reading_order {
            {
                let s = val.to_string();
                start.push_attribute(("readingOrder", s.as_str()));
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

impl ToXml for Borders {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        if let Some(ref val) = self.count {
            {
                let s = val.to_string();
                start.push_attribute(("count", s.as_str()));
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
        for item in &self.border {
            #[cfg(feature = "extra-children")]
            while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
                extra_iter
                    .next()
                    .unwrap()
                    .node
                    .write_to(writer)
                    .map_err(SerializeError::from)?;
            }
            item.write_element("border", writer)?;
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
        if !self.border.is_empty() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for Border {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        #[cfg(feature = "sml-styling")]
        if let Some(ref val) = self.diagonal_up {
            start.push_attribute(("diagonalUp", if *val { "1" } else { "0" }));
        }
        #[cfg(feature = "sml-styling")]
        if let Some(ref val) = self.diagonal_down {
            start.push_attribute(("diagonalDown", if *val { "1" } else { "0" }));
        }
        #[cfg(feature = "sml-styling")]
        if let Some(ref val) = self.outline {
            start.push_attribute(("outline", if *val { "1" } else { "0" }));
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
        if let Some(ref val) = self.start {
            val.write_element("start", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        if let Some(ref val) = self.end {
            val.write_element("end", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        #[cfg(feature = "sml-styling")]
        if let Some(ref val) = self.left {
            val.write_element("left", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        #[cfg(feature = "sml-styling")]
        if let Some(ref val) = self.right {
            val.write_element("right", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        #[cfg(feature = "sml-styling")]
        if let Some(ref val) = self.top {
            val.write_element("top", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        #[cfg(feature = "sml-styling")]
        if let Some(ref val) = self.bottom {
            val.write_element("bottom", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        #[cfg(feature = "sml-styling")]
        if let Some(ref val) = self.diagonal {
            val.write_element("diagonal", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        #[cfg(feature = "sml-styling")]
        if let Some(ref val) = self.vertical {
            val.write_element("vertical", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        #[cfg(feature = "sml-styling")]
        if let Some(ref val) = self.horizontal {
            val.write_element("horizontal", writer)?;
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
        if self.start.is_some() {
            return false;
        }
        if self.end.is_some() {
            return false;
        }
        #[cfg(feature = "sml-styling")]
        if self.left.is_some() {
            return false;
        }
        #[cfg(feature = "sml-styling")]
        if self.right.is_some() {
            return false;
        }
        #[cfg(feature = "sml-styling")]
        if self.top.is_some() {
            return false;
        }
        #[cfg(feature = "sml-styling")]
        if self.bottom.is_some() {
            return false;
        }
        #[cfg(feature = "sml-styling")]
        if self.diagonal.is_some() {
            return false;
        }
        #[cfg(feature = "sml-styling")]
        if self.vertical.is_some() {
            return false;
        }
        #[cfg(feature = "sml-styling")]
        if self.horizontal.is_some() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for BorderProperties {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        if let Some(ref val) = self.style {
            {
                let s = val.to_string();
                start.push_attribute(("style", s.as_str()));
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
        if let Some(ref val) = self.color {
            val.write_element("color", writer)?;
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
        if self.color.is_some() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for CellProtection {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        if let Some(ref val) = self.locked {
            start.push_attribute(("locked", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.hidden {
            start.push_attribute(("hidden", if *val { "1" } else { "0" }));
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

impl ToXml for Fonts {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        if let Some(ref val) = self.count {
            {
                let s = val.to_string();
                start.push_attribute(("count", s.as_str()));
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
            item.write_element("font", writer)?;
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

impl ToXml for Fills {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        if let Some(ref val) = self.count {
            {
                let s = val.to_string();
                start.push_attribute(("count", s.as_str()));
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
        for item in &self.fill {
            #[cfg(feature = "extra-children")]
            while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
                extra_iter
                    .next()
                    .unwrap()
                    .node
                    .write_to(writer)
                    .map_err(SerializeError::from)?;
            }
            item.write_element("fill", writer)?;
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
        if !self.fill.is_empty() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for Fill {
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

impl ToXml for PatternFill {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        if let Some(ref val) = self.pattern_type {
            {
                let s = val.to_string();
                start.push_attribute(("patternType", s.as_str()));
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
        if let Some(ref val) = self.fg_color {
            val.write_element("fgColor", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        if let Some(ref val) = self.bg_color {
            val.write_element("bgColor", writer)?;
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
        if self.fg_color.is_some() {
            return false;
        }
        if self.bg_color.is_some() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for Color {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        if let Some(ref val) = self.auto {
            start.push_attribute(("auto", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.indexed {
            {
                let s = val.to_string();
                start.push_attribute(("indexed", s.as_str()));
            }
        }
        if let Some(ref val) = self.rgb {
            {
                let hex = encode_hex(val);
                start.push_attribute(("rgb", hex.as_str()));
            }
        }
        if let Some(ref val) = self.theme {
            {
                let s = val.to_string();
                start.push_attribute(("theme", s.as_str()));
            }
        }
        if let Some(ref val) = self.tint {
            {
                let s = val.to_string();
                start.push_attribute(("tint", s.as_str()));
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

impl ToXml for GradientFill {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        if let Some(ref val) = self.r#type {
            {
                let s = val.to_string();
                start.push_attribute(("type", s.as_str()));
            }
        }
        if let Some(ref val) = self.degree {
            {
                let s = val.to_string();
                start.push_attribute(("degree", s.as_str()));
            }
        }
        if let Some(ref val) = self.left {
            {
                let s = val.to_string();
                start.push_attribute(("left", s.as_str()));
            }
        }
        if let Some(ref val) = self.right {
            {
                let s = val.to_string();
                start.push_attribute(("right", s.as_str()));
            }
        }
        if let Some(ref val) = self.top {
            {
                let s = val.to_string();
                start.push_attribute(("top", s.as_str()));
            }
        }
        if let Some(ref val) = self.bottom {
            {
                let s = val.to_string();
                start.push_attribute(("bottom", s.as_str()));
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
        for item in &self.stop {
            #[cfg(feature = "extra-children")]
            while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
                extra_iter
                    .next()
                    .unwrap()
                    .node
                    .write_to(writer)
                    .map_err(SerializeError::from)?;
            }
            item.write_element("stop", writer)?;
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
        if !self.stop.is_empty() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for GradientStop {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        {
            let val = &self.position;
            {
                let s = val.to_string();
                start.push_attribute(("position", s.as_str()));
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
            let val = &self.color;
            val.write_element("color", writer)?;
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

impl ToXml for NumberFormats {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        if let Some(ref val) = self.count {
            {
                let s = val.to_string();
                start.push_attribute(("count", s.as_str()));
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
        for item in &self.num_fmt {
            #[cfg(feature = "extra-children")]
            while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
                extra_iter
                    .next()
                    .unwrap()
                    .node
                    .write_to(writer)
                    .map_err(SerializeError::from)?;
            }
            item.write_element("numFmt", writer)?;
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
        if !self.num_fmt.is_empty() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for NumberFormat {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        {
            let val = &self.number_format_id;
            {
                let s = val.to_string();
                start.push_attribute(("numFmtId", s.as_str()));
            }
        }
        {
            let val = &self.format_code;
            start.push_attribute(("formatCode", val.as_str()));
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

impl ToXml for CellStyleFormats {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        if let Some(ref val) = self.count {
            {
                let s = val.to_string();
                start.push_attribute(("count", s.as_str()));
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
        for item in &self.xf {
            #[cfg(feature = "extra-children")]
            while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
                extra_iter
                    .next()
                    .unwrap()
                    .node
                    .write_to(writer)
                    .map_err(SerializeError::from)?;
            }
            item.write_element("xf", writer)?;
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
        if !self.xf.is_empty() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for CellFormats {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        if let Some(ref val) = self.count {
            {
                let s = val.to_string();
                start.push_attribute(("count", s.as_str()));
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
        for item in &self.xf {
            #[cfg(feature = "extra-children")]
            while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
                extra_iter
                    .next()
                    .unwrap()
                    .node
                    .write_to(writer)
                    .map_err(SerializeError::from)?;
            }
            item.write_element("xf", writer)?;
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
        if !self.xf.is_empty() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for Format {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        #[cfg(feature = "sml-styling")]
        if let Some(ref val) = self.number_format_id {
            {
                let s = val.to_string();
                start.push_attribute(("numFmtId", s.as_str()));
            }
        }
        #[cfg(feature = "sml-styling")]
        if let Some(ref val) = self.font_id {
            {
                let s = val.to_string();
                start.push_attribute(("fontId", s.as_str()));
            }
        }
        #[cfg(feature = "sml-styling")]
        if let Some(ref val) = self.fill_id {
            {
                let s = val.to_string();
                start.push_attribute(("fillId", s.as_str()));
            }
        }
        #[cfg(feature = "sml-styling")]
        if let Some(ref val) = self.border_id {
            {
                let s = val.to_string();
                start.push_attribute(("borderId", s.as_str()));
            }
        }
        #[cfg(feature = "sml-styling")]
        if let Some(ref val) = self.format_id {
            {
                let s = val.to_string();
                start.push_attribute(("xfId", s.as_str()));
            }
        }
        #[cfg(feature = "sml-styling")]
        if let Some(ref val) = self.quote_prefix {
            start.push_attribute(("quotePrefix", if *val { "1" } else { "0" }));
        }
        #[cfg(feature = "sml-pivot")]
        if let Some(ref val) = self.pivot_button {
            start.push_attribute(("pivotButton", if *val { "1" } else { "0" }));
        }
        #[cfg(feature = "sml-styling")]
        if let Some(ref val) = self.apply_number_format {
            start.push_attribute(("applyNumberFormat", if *val { "1" } else { "0" }));
        }
        #[cfg(feature = "sml-styling")]
        if let Some(ref val) = self.apply_font {
            start.push_attribute(("applyFont", if *val { "1" } else { "0" }));
        }
        #[cfg(feature = "sml-styling")]
        if let Some(ref val) = self.apply_fill {
            start.push_attribute(("applyFill", if *val { "1" } else { "0" }));
        }
        #[cfg(feature = "sml-styling")]
        if let Some(ref val) = self.apply_border {
            start.push_attribute(("applyBorder", if *val { "1" } else { "0" }));
        }
        #[cfg(feature = "sml-styling")]
        if let Some(ref val) = self.apply_alignment {
            start.push_attribute(("applyAlignment", if *val { "1" } else { "0" }));
        }
        #[cfg(feature = "sml-styling")]
        if let Some(ref val) = self.apply_protection {
            start.push_attribute(("applyProtection", if *val { "1" } else { "0" }));
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
        #[cfg(feature = "sml-styling")]
        if let Some(ref val) = self.alignment {
            val.write_element("alignment", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        #[cfg(feature = "sml-protection")]
        if let Some(ref val) = self.protection {
            val.write_element("protection", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        #[cfg(feature = "sml-extensions")]
        if let Some(ref val) = self.extension_list {
            val.write_element("extLst", writer)?;
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
        #[cfg(feature = "sml-styling")]
        if self.alignment.is_some() {
            return false;
        }
        #[cfg(feature = "sml-protection")]
        if self.protection.is_some() {
            return false;
        }
        #[cfg(feature = "sml-extensions")]
        if self.extension_list.is_some() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for CellStyles {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        if let Some(ref val) = self.count {
            {
                let s = val.to_string();
                start.push_attribute(("count", s.as_str()));
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
        for item in &self.cell_style {
            #[cfg(feature = "extra-children")]
            while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
                extra_iter
                    .next()
                    .unwrap()
                    .node
                    .write_to(writer)
                    .map_err(SerializeError::from)?;
            }
            item.write_element("cellStyle", writer)?;
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
        if !self.cell_style.is_empty() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for CellStyle {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        if let Some(ref val) = self.name {
            start.push_attribute(("name", val.as_str()));
        }
        {
            let val = &self.format_id;
            {
                let s = val.to_string();
                start.push_attribute(("xfId", s.as_str()));
            }
        }
        if let Some(ref val) = self.builtin_id {
            {
                let s = val.to_string();
                start.push_attribute(("builtinId", s.as_str()));
            }
        }
        if let Some(ref val) = self.i_level {
            {
                let s = val.to_string();
                start.push_attribute(("iLevel", s.as_str()));
            }
        }
        if let Some(ref val) = self.hidden {
            start.push_attribute(("hidden", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.custom_builtin {
            start.push_attribute(("customBuiltin", if *val { "1" } else { "0" }));
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
        if let Some(ref val) = self.extension_list {
            val.write_element("extLst", writer)?;
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
        if self.extension_list.is_some() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for DifferentialFormats {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        if let Some(ref val) = self.count {
            {
                let s = val.to_string();
                start.push_attribute(("count", s.as_str()));
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
        for item in &self.dxf {
            #[cfg(feature = "extra-children")]
            while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
                extra_iter
                    .next()
                    .unwrap()
                    .node
                    .write_to(writer)
                    .map_err(SerializeError::from)?;
            }
            item.write_element("dxf", writer)?;
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
        if !self.dxf.is_empty() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for DifferentialFormat {
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
        if let Some(ref val) = self.font {
            val.write_element("font", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        if let Some(ref val) = self.num_fmt {
            val.write_element("numFmt", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        if let Some(ref val) = self.fill {
            val.write_element("fill", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        if let Some(ref val) = self.alignment {
            val.write_element("alignment", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        if let Some(ref val) = self.border {
            val.write_element("border", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        if let Some(ref val) = self.protection {
            val.write_element("protection", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        if let Some(ref val) = self.extension_list {
            val.write_element("extLst", writer)?;
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
        if self.font.is_some() {
            return false;
        }
        if self.num_fmt.is_some() {
            return false;
        }
        if self.fill.is_some() {
            return false;
        }
        if self.alignment.is_some() {
            return false;
        }
        if self.border.is_some() {
            return false;
        }
        if self.protection.is_some() {
            return false;
        }
        if self.extension_list.is_some() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for Colors {
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
        if let Some(ref val) = self.indexed_colors {
            val.write_element("indexedColors", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        if let Some(ref val) = self.mru_colors {
            val.write_element("mruColors", writer)?;
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
        if self.indexed_colors.is_some() {
            return false;
        }
        if self.mru_colors.is_some() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for IndexedColors {
    fn write_children<W: Write>(&self, writer: &mut Writer<W>) -> Result<(), SerializeError> {
        #[cfg(feature = "extra-children")]
        let mut extra_iter = self.extra_children.iter().peekable();
        #[cfg(feature = "extra-children")]
        let mut emit_idx: usize = 0;
        for item in &self.rgb_color {
            #[cfg(feature = "extra-children")]
            while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
                extra_iter
                    .next()
                    .unwrap()
                    .node
                    .write_to(writer)
                    .map_err(SerializeError::from)?;
            }
            item.write_element("rgbColor", writer)?;
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
        if !self.rgb_color.is_empty() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for MostRecentColors {
    fn write_children<W: Write>(&self, writer: &mut Writer<W>) -> Result<(), SerializeError> {
        #[cfg(feature = "extra-children")]
        let mut extra_iter = self.extra_children.iter().peekable();
        #[cfg(feature = "extra-children")]
        let mut emit_idx: usize = 0;
        for item in &self.color {
            #[cfg(feature = "extra-children")]
            while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
                extra_iter
                    .next()
                    .unwrap()
                    .node
                    .write_to(writer)
                    .map_err(SerializeError::from)?;
            }
            item.write_element("color", writer)?;
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
        if !self.color.is_empty() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for RgbColor {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        if let Some(ref val) = self.rgb {
            {
                let hex = encode_hex(val);
                start.push_attribute(("rgb", hex.as_str()));
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

impl ToXml for TableStyles {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        if let Some(ref val) = self.count {
            {
                let s = val.to_string();
                start.push_attribute(("count", s.as_str()));
            }
        }
        if let Some(ref val) = self.default_table_style {
            start.push_attribute(("defaultTableStyle", val.as_str()));
        }
        if let Some(ref val) = self.default_pivot_style {
            start.push_attribute(("defaultPivotStyle", val.as_str()));
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
        for item in &self.table_style {
            #[cfg(feature = "extra-children")]
            while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
                extra_iter
                    .next()
                    .unwrap()
                    .node
                    .write_to(writer)
                    .map_err(SerializeError::from)?;
            }
            item.write_element("tableStyle", writer)?;
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
        if !self.table_style.is_empty() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for TableStyle {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        {
            let val = &self.name;
            start.push_attribute(("name", val.as_str()));
        }
        if let Some(ref val) = self.pivot {
            start.push_attribute(("pivot", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.table {
            start.push_attribute(("table", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.count {
            {
                let s = val.to_string();
                start.push_attribute(("count", s.as_str()));
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
        for item in &self.table_style_element {
            #[cfg(feature = "extra-children")]
            while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
                extra_iter
                    .next()
                    .unwrap()
                    .node
                    .write_to(writer)
                    .map_err(SerializeError::from)?;
            }
            item.write_element("tableStyleElement", writer)?;
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
        if !self.table_style_element.is_empty() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for TableStyleElement {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        {
            let val = &self.r#type;
            {
                let s = val.to_string();
                start.push_attribute(("type", s.as_str()));
            }
        }
        if let Some(ref val) = self.size {
            {
                let s = val.to_string();
                start.push_attribute(("size", s.as_str()));
            }
        }
        if let Some(ref val) = self.dxf_id {
            {
                let s = val.to_string();
                start.push_attribute(("dxfId", s.as_str()));
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

impl ToXml for BooleanProperty {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        if let Some(ref val) = self.value {
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

impl ToXml for FontSize {
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

impl ToXml for IntProperty {
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

impl ToXml for FontName {
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

impl ToXml for VerticalAlignFontProperty {
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

impl ToXml for FontSchemeProperty {
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

impl ToXml for UnderlineProperty {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
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

impl ToXml for Font {
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

impl ToXml for FontFamily {
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

impl ToXml for SmlAGAutoFormat {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        if let Some(ref val) = self.auto_format_id {
            {
                let s = val.to_string();
                start.push_attribute(("autoFormatId", s.as_str()));
            }
        }
        if let Some(ref val) = self.apply_number_formats {
            start.push_attribute(("applyNumberFormats", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.apply_border_formats {
            start.push_attribute(("applyBorderFormats", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.apply_font_formats {
            start.push_attribute(("applyFontFormats", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.apply_pattern_formats {
            start.push_attribute(("applyPatternFormats", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.apply_alignment_formats {
            start.push_attribute(("applyAlignmentFormats", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.apply_width_height_formats {
            start.push_attribute(("applyWidthHeightFormats", if *val { "1" } else { "0" }));
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

impl ToXml for ExternalLink {
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
        if let Some(ref val) = self.extension_list {
            val.write_element("extLst", writer)?;
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
        if self.extension_list.is_some() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for ExternalBook {
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
        if let Some(ref val) = self.sheet_names {
            val.write_element("sheetNames", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        if let Some(ref val) = self.defined_names {
            val.write_element("definedNames", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        if let Some(ref val) = self.sheet_data_set {
            val.write_element("sheetDataSet", writer)?;
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
        if self.sheet_names.is_some() {
            return false;
        }
        if self.defined_names.is_some() {
            return false;
        }
        if self.sheet_data_set.is_some() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for CTExternalSheetNames {
    fn write_children<W: Write>(&self, writer: &mut Writer<W>) -> Result<(), SerializeError> {
        #[cfg(feature = "extra-children")]
        let mut extra_iter = self.extra_children.iter().peekable();
        #[cfg(feature = "extra-children")]
        let mut emit_idx: usize = 0;
        for item in &self.sheet_name {
            #[cfg(feature = "extra-children")]
            while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
                extra_iter
                    .next()
                    .unwrap()
                    .node
                    .write_to(writer)
                    .map_err(SerializeError::from)?;
            }
            item.write_element("sheetName", writer)?;
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
        if !self.sheet_name.is_empty() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for CTExternalSheetName {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        if let Some(ref val) = self.value {
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

impl ToXml for CTExternalDefinedNames {
    fn write_children<W: Write>(&self, writer: &mut Writer<W>) -> Result<(), SerializeError> {
        #[cfg(feature = "extra-children")]
        let mut extra_iter = self.extra_children.iter().peekable();
        #[cfg(feature = "extra-children")]
        let mut emit_idx: usize = 0;
        for item in &self.defined_name {
            #[cfg(feature = "extra-children")]
            while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
                extra_iter
                    .next()
                    .unwrap()
                    .node
                    .write_to(writer)
                    .map_err(SerializeError::from)?;
            }
            item.write_element("definedName", writer)?;
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
        if !self.defined_name.is_empty() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for CTExternalDefinedName {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        {
            let val = &self.name;
            start.push_attribute(("name", val.as_str()));
        }
        if let Some(ref val) = self.refers_to {
            start.push_attribute(("refersTo", val.as_str()));
        }
        if let Some(ref val) = self.sheet_id {
            {
                let s = val.to_string();
                start.push_attribute(("sheetId", s.as_str()));
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

impl ToXml for ExternalSheetDataSet {
    fn write_children<W: Write>(&self, writer: &mut Writer<W>) -> Result<(), SerializeError> {
        #[cfg(feature = "extra-children")]
        let mut extra_iter = self.extra_children.iter().peekable();
        #[cfg(feature = "extra-children")]
        let mut emit_idx: usize = 0;
        for item in &self.sheet_data {
            #[cfg(feature = "extra-children")]
            while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
                extra_iter
                    .next()
                    .unwrap()
                    .node
                    .write_to(writer)
                    .map_err(SerializeError::from)?;
            }
            item.write_element("sheetData", writer)?;
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
        if !self.sheet_data.is_empty() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for ExternalSheetData {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        {
            let val = &self.sheet_id;
            {
                let s = val.to_string();
                start.push_attribute(("sheetId", s.as_str()));
            }
        }
        if let Some(ref val) = self.refresh_error {
            start.push_attribute(("refreshError", if *val { "1" } else { "0" }));
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
        for item in &self.row {
            #[cfg(feature = "extra-children")]
            while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
                extra_iter
                    .next()
                    .unwrap()
                    .node
                    .write_to(writer)
                    .map_err(SerializeError::from)?;
            }
            item.write_element("row", writer)?;
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
        if !self.row.is_empty() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for ExternalRow {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        {
            let val = &self.reference;
            {
                let s = val.to_string();
                start.push_attribute(("r", s.as_str()));
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
        for item in &self.cell {
            #[cfg(feature = "extra-children")]
            while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
                extra_iter
                    .next()
                    .unwrap()
                    .node
                    .write_to(writer)
                    .map_err(SerializeError::from)?;
            }
            item.write_element("cell", writer)?;
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
        if !self.cell.is_empty() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for ExternalCell {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        if let Some(ref val) = self.reference {
            start.push_attribute(("r", val.as_str()));
        }
        if let Some(ref val) = self.cell_type {
            {
                let s = val.to_string();
                start.push_attribute(("t", s.as_str()));
            }
        }
        if let Some(ref val) = self.vm {
            {
                let s = val.to_string();
                start.push_attribute(("vm", s.as_str()));
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
        if let Some(ref val) = self.value {
            {
                let start = BytesStart::new("v");
                writer.write_event(Event::Start(start))?;
                writer.write_event(Event::Text(BytesText::new(val.as_str())))?;
                writer.write_event(Event::End(BytesEnd::new("v")))?;
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

impl ToXml for DdeLink {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        {
            let val = &self.dde_service;
            start.push_attribute(("ddeService", val.as_str()));
        }
        {
            let val = &self.dde_topic;
            start.push_attribute(("ddeTopic", val.as_str()));
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
        if let Some(ref val) = self.dde_items {
            val.write_element("ddeItems", writer)?;
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
        if self.dde_items.is_some() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for DdeItems {
    fn write_children<W: Write>(&self, writer: &mut Writer<W>) -> Result<(), SerializeError> {
        #[cfg(feature = "extra-children")]
        let mut extra_iter = self.extra_children.iter().peekable();
        #[cfg(feature = "extra-children")]
        let mut emit_idx: usize = 0;
        for item in &self.dde_item {
            #[cfg(feature = "extra-children")]
            while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
                extra_iter
                    .next()
                    .unwrap()
                    .node
                    .write_to(writer)
                    .map_err(SerializeError::from)?;
            }
            item.write_element("ddeItem", writer)?;
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
        if !self.dde_item.is_empty() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for DdeItem {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        if let Some(ref val) = self.name {
            start.push_attribute(("name", val.as_str()));
        }
        if let Some(ref val) = self.ole {
            start.push_attribute(("ole", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.advise {
            start.push_attribute(("advise", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.prefer_pic {
            start.push_attribute(("preferPic", if *val { "1" } else { "0" }));
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
        if let Some(ref val) = self.values {
            val.write_element("values", writer)?;
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
        if self.values.is_some() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for CTDdeValues {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        if let Some(ref val) = self.rows {
            {
                let s = val.to_string();
                start.push_attribute(("rows", s.as_str()));
            }
        }
        if let Some(ref val) = self.cols {
            {
                let s = val.to_string();
                start.push_attribute(("cols", s.as_str()));
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
        for item in &self.value {
            #[cfg(feature = "extra-children")]
            while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
                extra_iter
                    .next()
                    .unwrap()
                    .node
                    .write_to(writer)
                    .map_err(SerializeError::from)?;
            }
            item.write_element("value", writer)?;
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
        if !self.value.is_empty() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for CTDdeValue {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        if let Some(ref val) = self.cell_type {
            {
                let s = val.to_string();
                start.push_attribute(("t", s.as_str()));
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
            let val = &self.value;
            {
                let start = BytesStart::new("val");
                writer.write_event(Event::Start(start))?;
                writer.write_event(Event::Text(BytesText::new(val.as_str())))?;
                writer.write_event(Event::End(BytesEnd::new("val")))?;
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
        false
    }
}

impl ToXml for OleLink {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        {
            let val = &self.prog_id;
            start.push_attribute(("progId", val.as_str()));
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
        if let Some(ref val) = self.ole_items {
            val.write_element("oleItems", writer)?;
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
        if self.ole_items.is_some() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for OleItems {
    fn write_children<W: Write>(&self, writer: &mut Writer<W>) -> Result<(), SerializeError> {
        #[cfg(feature = "extra-children")]
        let mut extra_iter = self.extra_children.iter().peekable();
        #[cfg(feature = "extra-children")]
        let mut emit_idx: usize = 0;
        for item in &self.ole_item {
            #[cfg(feature = "extra-children")]
            while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
                extra_iter
                    .next()
                    .unwrap()
                    .node
                    .write_to(writer)
                    .map_err(SerializeError::from)?;
            }
            item.write_element("oleItem", writer)?;
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
        if !self.ole_item.is_empty() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for OleItem {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        {
            let val = &self.name;
            start.push_attribute(("name", val.as_str()));
        }
        if let Some(ref val) = self.icon {
            start.push_attribute(("icon", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.advise {
            start.push_attribute(("advise", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.prefer_pic {
            start.push_attribute(("preferPic", if *val { "1" } else { "0" }));
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

impl ToXml for Table {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        #[cfg(feature = "sml-tables")]
        {
            let val = &self.id;
            {
                let s = val.to_string();
                start.push_attribute(("id", s.as_str()));
            }
        }
        #[cfg(feature = "sml-tables")]
        if let Some(ref val) = self.name {
            start.push_attribute(("name", val.as_str()));
        }
        #[cfg(feature = "sml-tables")]
        {
            let val = &self.display_name;
            start.push_attribute(("displayName", val.as_str()));
        }
        #[cfg(feature = "sml-tables")]
        if let Some(ref val) = self.comment {
            start.push_attribute(("comment", val.as_str()));
        }
        #[cfg(feature = "sml-tables")]
        {
            let val = &self.reference;
            start.push_attribute(("ref", val.as_str()));
        }
        #[cfg(feature = "sml-tables")]
        if let Some(ref val) = self.table_type {
            {
                let s = val.to_string();
                start.push_attribute(("tableType", s.as_str()));
            }
        }
        #[cfg(feature = "sml-tables")]
        if let Some(ref val) = self.header_row_count {
            {
                let s = val.to_string();
                start.push_attribute(("headerRowCount", s.as_str()));
            }
        }
        #[cfg(feature = "sml-tables")]
        if let Some(ref val) = self.insert_row {
            start.push_attribute(("insertRow", if *val { "1" } else { "0" }));
        }
        #[cfg(feature = "sml-tables")]
        if let Some(ref val) = self.insert_row_shift {
            start.push_attribute(("insertRowShift", if *val { "1" } else { "0" }));
        }
        #[cfg(feature = "sml-tables")]
        if let Some(ref val) = self.totals_row_count {
            {
                let s = val.to_string();
                start.push_attribute(("totalsRowCount", s.as_str()));
            }
        }
        #[cfg(feature = "sml-tables")]
        if let Some(ref val) = self.totals_row_shown {
            start.push_attribute(("totalsRowShown", if *val { "1" } else { "0" }));
        }
        #[cfg(feature = "sml-tables")]
        if let Some(ref val) = self.published {
            start.push_attribute(("published", if *val { "1" } else { "0" }));
        }
        #[cfg(feature = "sml-tables")]
        if let Some(ref val) = self.header_row_dxf_id {
            {
                let s = val.to_string();
                start.push_attribute(("headerRowDxfId", s.as_str()));
            }
        }
        #[cfg(feature = "sml-tables")]
        if let Some(ref val) = self.data_dxf_id {
            {
                let s = val.to_string();
                start.push_attribute(("dataDxfId", s.as_str()));
            }
        }
        #[cfg(feature = "sml-tables")]
        if let Some(ref val) = self.totals_row_dxf_id {
            {
                let s = val.to_string();
                start.push_attribute(("totalsRowDxfId", s.as_str()));
            }
        }
        #[cfg(feature = "sml-tables")]
        if let Some(ref val) = self.header_row_border_dxf_id {
            {
                let s = val.to_string();
                start.push_attribute(("headerRowBorderDxfId", s.as_str()));
            }
        }
        #[cfg(feature = "sml-tables")]
        if let Some(ref val) = self.table_border_dxf_id {
            {
                let s = val.to_string();
                start.push_attribute(("tableBorderDxfId", s.as_str()));
            }
        }
        #[cfg(feature = "sml-tables")]
        if let Some(ref val) = self.totals_row_border_dxf_id {
            {
                let s = val.to_string();
                start.push_attribute(("totalsRowBorderDxfId", s.as_str()));
            }
        }
        #[cfg(feature = "sml-tables")]
        if let Some(ref val) = self.header_row_cell_style {
            start.push_attribute(("headerRowCellStyle", val.as_str()));
        }
        #[cfg(feature = "sml-tables")]
        if let Some(ref val) = self.data_cell_style {
            start.push_attribute(("dataCellStyle", val.as_str()));
        }
        #[cfg(feature = "sml-tables")]
        if let Some(ref val) = self.totals_row_cell_style {
            start.push_attribute(("totalsRowCellStyle", val.as_str()));
        }
        #[cfg(feature = "sml-tables")]
        if let Some(ref val) = self.connection_id {
            {
                let s = val.to_string();
                start.push_attribute(("connectionId", s.as_str()));
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
        #[cfg(feature = "sml-tables")]
        if let Some(ref val) = self.auto_filter {
            val.write_element("autoFilter", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        #[cfg(feature = "sml-tables")]
        if let Some(ref val) = self.sort_state {
            val.write_element("sortState", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        #[cfg(feature = "sml-tables")]
        {
            let val = &self.table_columns;
            val.write_element("tableColumns", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        #[cfg(feature = "sml-tables")]
        if let Some(ref val) = self.table_style_info {
            val.write_element("tableStyleInfo", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        #[cfg(feature = "sml-extensions")]
        if let Some(ref val) = self.extension_list {
            val.write_element("extLst", writer)?;
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
        #[cfg(feature = "sml-tables")]
        if self.auto_filter.is_some() {
            return false;
        }
        #[cfg(feature = "sml-tables")]
        if self.sort_state.is_some() {
            return false;
        }
        #[cfg(feature = "sml-tables")]
        return false;
        #[cfg(feature = "sml-tables")]
        if self.table_style_info.is_some() {
            return false;
        }
        #[cfg(feature = "sml-extensions")]
        if self.extension_list.is_some() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for TableStyleInfo {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        if let Some(ref val) = self.name {
            start.push_attribute(("name", val.as_str()));
        }
        if let Some(ref val) = self.show_first_column {
            start.push_attribute(("showFirstColumn", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.show_last_column {
            start.push_attribute(("showLastColumn", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.show_row_stripes {
            start.push_attribute(("showRowStripes", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.show_column_stripes {
            start.push_attribute(("showColumnStripes", if *val { "1" } else { "0" }));
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

impl ToXml for TableColumns {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        if let Some(ref val) = self.count {
            {
                let s = val.to_string();
                start.push_attribute(("count", s.as_str()));
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
        for item in &self.table_column {
            #[cfg(feature = "extra-children")]
            while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
                extra_iter
                    .next()
                    .unwrap()
                    .node
                    .write_to(writer)
                    .map_err(SerializeError::from)?;
            }
            item.write_element("tableColumn", writer)?;
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
        if !self.table_column.is_empty() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for TableColumn {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        {
            let val = &self.id;
            {
                let s = val.to_string();
                start.push_attribute(("id", s.as_str()));
            }
        }
        if let Some(ref val) = self.unique_name {
            start.push_attribute(("uniqueName", val.as_str()));
        }
        {
            let val = &self.name;
            start.push_attribute(("name", val.as_str()));
        }
        if let Some(ref val) = self.totals_row_function {
            {
                let s = val.to_string();
                start.push_attribute(("totalsRowFunction", s.as_str()));
            }
        }
        if let Some(ref val) = self.totals_row_label {
            start.push_attribute(("totalsRowLabel", val.as_str()));
        }
        if let Some(ref val) = self.query_table_field_id {
            {
                let s = val.to_string();
                start.push_attribute(("queryTableFieldId", s.as_str()));
            }
        }
        if let Some(ref val) = self.header_row_dxf_id {
            {
                let s = val.to_string();
                start.push_attribute(("headerRowDxfId", s.as_str()));
            }
        }
        if let Some(ref val) = self.data_dxf_id {
            {
                let s = val.to_string();
                start.push_attribute(("dataDxfId", s.as_str()));
            }
        }
        if let Some(ref val) = self.totals_row_dxf_id {
            {
                let s = val.to_string();
                start.push_attribute(("totalsRowDxfId", s.as_str()));
            }
        }
        if let Some(ref val) = self.header_row_cell_style {
            start.push_attribute(("headerRowCellStyle", val.as_str()));
        }
        if let Some(ref val) = self.data_cell_style {
            start.push_attribute(("dataCellStyle", val.as_str()));
        }
        if let Some(ref val) = self.totals_row_cell_style {
            start.push_attribute(("totalsRowCellStyle", val.as_str()));
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
        if let Some(ref val) = self.calculated_column_formula {
            val.write_element("calculatedColumnFormula", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        if let Some(ref val) = self.totals_row_formula {
            val.write_element("totalsRowFormula", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        if let Some(ref val) = self.xml_column_pr {
            val.write_element("xmlColumnPr", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        if let Some(ref val) = self.extension_list {
            val.write_element("extLst", writer)?;
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
        if self.calculated_column_formula.is_some() {
            return false;
        }
        if self.totals_row_formula.is_some() {
            return false;
        }
        if self.xml_column_pr.is_some() {
            return false;
        }
        if self.extension_list.is_some() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for TableFormula {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        if let Some(ref val) = self.array {
            start.push_attribute(("array", if *val { "1" } else { "0" }));
        }
        #[cfg(feature = "extra-attrs")]
        for (key, value) in &self.extra_attrs {
            start.push_attribute((key.as_str(), value.as_str()));
        }
        start
    }

    fn write_children<W: Write>(&self, writer: &mut Writer<W>) -> Result<(), SerializeError> {
        if let Some(ref text) = self.text {
            writer.write_event(Event::Text(BytesText::new(text)))?;
        }
        #[cfg(feature = "extra-children")]
        for extra in &self.extra_children {
            extra.node.write_to(writer).map_err(SerializeError::from)?;
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

impl ToXml for XmlColumnProperties {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        {
            let val = &self.map_id;
            {
                let s = val.to_string();
                start.push_attribute(("mapId", s.as_str()));
            }
        }
        {
            let val = &self.xpath;
            start.push_attribute(("xpath", val.as_str()));
        }
        if let Some(ref val) = self.denormalized {
            start.push_attribute(("denormalized", if *val { "1" } else { "0" }));
        }
        {
            let val = &self.xml_data_type;
            start.push_attribute(("xmlDataType", val.as_str()));
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
        if let Some(ref val) = self.extension_list {
            val.write_element("extLst", writer)?;
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
        if self.extension_list.is_some() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for CTVolTypes {
    fn write_children<W: Write>(&self, writer: &mut Writer<W>) -> Result<(), SerializeError> {
        #[cfg(feature = "extra-children")]
        let mut extra_iter = self.extra_children.iter().peekable();
        #[cfg(feature = "extra-children")]
        let mut emit_idx: usize = 0;
        for item in &self.vol_type {
            #[cfg(feature = "extra-children")]
            while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
                extra_iter
                    .next()
                    .unwrap()
                    .node
                    .write_to(writer)
                    .map_err(SerializeError::from)?;
            }
            item.write_element("volType", writer)?;
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
        if let Some(ref val) = self.extension_list {
            val.write_element("extLst", writer)?;
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
        if !self.vol_type.is_empty() {
            return false;
        }
        if self.extension_list.is_some() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for CTVolType {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        {
            let val = &self.r#type;
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

    fn write_children<W: Write>(&self, writer: &mut Writer<W>) -> Result<(), SerializeError> {
        #[cfg(feature = "extra-children")]
        let mut extra_iter = self.extra_children.iter().peekable();
        #[cfg(feature = "extra-children")]
        let mut emit_idx: usize = 0;
        for item in &self.main {
            #[cfg(feature = "extra-children")]
            while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
                extra_iter
                    .next()
                    .unwrap()
                    .node
                    .write_to(writer)
                    .map_err(SerializeError::from)?;
            }
            item.write_element("main", writer)?;
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
        if !self.main.is_empty() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for CTVolMain {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        {
            let val = &self.first;
            start.push_attribute(("first", val.as_str()));
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
        for item in &self.tp {
            #[cfg(feature = "extra-children")]
            while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
                extra_iter
                    .next()
                    .unwrap()
                    .node
                    .write_to(writer)
                    .map_err(SerializeError::from)?;
            }
            item.write_element("tp", writer)?;
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
        if !self.tp.is_empty() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for CTVolTopic {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        if let Some(ref val) = self.cell_type {
            {
                let s = val.to_string();
                start.push_attribute(("t", s.as_str()));
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
            let val = &self.value;
            {
                let start = BytesStart::new("v");
                writer.write_event(Event::Start(start))?;
                writer.write_event(Event::Text(BytesText::new(val.as_str())))?;
                writer.write_event(Event::End(BytesEnd::new("v")))?;
            }
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        for item in &self.stp {
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
                let start = BytesStart::new("stp");
                writer.write_event(Event::Start(start))?;
                writer.write_event(Event::Text(BytesText::new(item.as_str())))?;
                writer.write_event(Event::End(BytesEnd::new("stp")))?;
            }
            #[cfg(feature = "extra-children")]
            {
                emit_idx += 1;
            }
        }
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
            item.write_element("tr", writer)?;
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
        false
    }
}

impl ToXml for CTVolTopicRef {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        {
            let val = &self.reference;
            start.push_attribute(("r", val.as_str()));
        }
        {
            let val = &self.style_index;
            {
                let s = val.to_string();
                start.push_attribute(("s", s.as_str()));
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

impl ToXml for Workbook {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
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
        if let Some(ref val) = self.file_version {
            val.write_element("fileVersion", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        #[cfg(feature = "sml-protection")]
        if let Some(ref val) = self.file_sharing {
            val.write_element("fileSharing", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        if let Some(ref val) = self.workbook_pr {
            val.write_element("workbookPr", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        #[cfg(feature = "sml-protection")]
        if let Some(ref val) = self.workbook_protection {
            val.write_element("workbookProtection", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        if let Some(ref val) = self.book_views {
            val.write_element("bookViews", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
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
            let val = &self.sheets;
            val.write_element("sheets", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        #[cfg(feature = "sml-formulas-advanced")]
        if let Some(ref val) = self.function_groups {
            val.write_element("functionGroups", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        #[cfg(feature = "sml-external")]
        if let Some(ref val) = self.external_references {
            val.write_element("externalReferences", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        if let Some(ref val) = self.defined_names {
            val.write_element("definedNames", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        #[cfg(feature = "sml-formulas")]
        if let Some(ref val) = self.calc_pr {
            val.write_element("calcPr", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        #[cfg(feature = "sml-external")]
        if let Some(ref val) = self.ole_size {
            val.write_element("oleSize", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        #[cfg(feature = "sml-structure")]
        if let Some(ref val) = self.custom_workbook_views {
            val.write_element("customWorkbookViews", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        #[cfg(feature = "sml-pivot")]
        if let Some(ref val) = self.pivot_caches {
            val.write_element("pivotCaches", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        #[cfg(feature = "sml-metadata")]
        if let Some(ref val) = self.smart_tag_pr {
            val.write_element("smartTagPr", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        #[cfg(feature = "sml-metadata")]
        if let Some(ref val) = self.smart_tag_types {
            val.write_element("smartTagTypes", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        #[cfg(feature = "sml-external")]
        if let Some(ref val) = self.web_publishing {
            val.write_element("webPublishing", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        for item in &self.file_recovery_pr {
            #[cfg(feature = "extra-children")]
            while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
                extra_iter
                    .next()
                    .unwrap()
                    .node
                    .write_to(writer)
                    .map_err(SerializeError::from)?;
            }
            item.write_element("fileRecoveryPr", writer)?;
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
        #[cfg(feature = "sml-external")]
        if let Some(ref val) = self.web_publish_objects {
            val.write_element("webPublishObjects", writer)?;
        }
        #[cfg(feature = "extra-children")]
        {
            emit_idx += 1;
        }
        #[cfg(feature = "extra-children")]
        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
            extra_iter
                .next()
                .unwrap()
                .node
                .write_to(writer)
                .map_err(SerializeError::from)?;
        }
        #[cfg(feature = "sml-extensions")]
        if let Some(ref val) = self.extension_list {
            val.write_element("extLst", writer)?;
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
        if self.file_version.is_some() {
            return false;
        }
        #[cfg(feature = "sml-protection")]
        if self.file_sharing.is_some() {
            return false;
        }
        if self.workbook_pr.is_some() {
            return false;
        }
        #[cfg(feature = "sml-protection")]
        if self.workbook_protection.is_some() {
            return false;
        }
        if self.book_views.is_some() {
            return false;
        }
        false
    }
}

impl ToXml for FileVersion {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        if let Some(ref val) = self.app_name {
            start.push_attribute(("appName", val.as_str()));
        }
        if let Some(ref val) = self.last_edited {
            start.push_attribute(("lastEdited", val.as_str()));
        }
        if let Some(ref val) = self.lowest_edited {
            start.push_attribute(("lowestEdited", val.as_str()));
        }
        if let Some(ref val) = self.rup_build {
            start.push_attribute(("rupBuild", val.as_str()));
        }
        if let Some(ref val) = self.code_name {
            start.push_attribute(("codeName", val.as_str()));
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

impl ToXml for BookViews {
    fn write_children<W: Write>(&self, writer: &mut Writer<W>) -> Result<(), SerializeError> {
        #[cfg(feature = "extra-children")]
        let mut extra_iter = self.extra_children.iter().peekable();
        #[cfg(feature = "extra-children")]
        let mut emit_idx: usize = 0;
        for item in &self.workbook_view {
            #[cfg(feature = "extra-children")]
            while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
                extra_iter
                    .next()
                    .unwrap()
                    .node
                    .write_to(writer)
                    .map_err(SerializeError::from)?;
            }
            item.write_element("workbookView", writer)?;
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
        if !self.workbook_view.is_empty() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for BookView {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        if let Some(ref val) = self.visibility {
            {
                let s = val.to_string();
                start.push_attribute(("visibility", s.as_str()));
            }
        }
        if let Some(ref val) = self.minimized {
            start.push_attribute(("minimized", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.show_horizontal_scroll {
            start.push_attribute(("showHorizontalScroll", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.show_vertical_scroll {
            start.push_attribute(("showVerticalScroll", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.show_sheet_tabs {
            start.push_attribute(("showSheetTabs", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.x_window {
            {
                let s = val.to_string();
                start.push_attribute(("xWindow", s.as_str()));
            }
        }
        if let Some(ref val) = self.y_window {
            {
                let s = val.to_string();
                start.push_attribute(("yWindow", s.as_str()));
            }
        }
        if let Some(ref val) = self.window_width {
            {
                let s = val.to_string();
                start.push_attribute(("windowWidth", s.as_str()));
            }
        }
        if let Some(ref val) = self.window_height {
            {
                let s = val.to_string();
                start.push_attribute(("windowHeight", s.as_str()));
            }
        }
        if let Some(ref val) = self.tab_ratio {
            {
                let s = val.to_string();
                start.push_attribute(("tabRatio", s.as_str()));
            }
        }
        if let Some(ref val) = self.first_sheet {
            {
                let s = val.to_string();
                start.push_attribute(("firstSheet", s.as_str()));
            }
        }
        if let Some(ref val) = self.active_tab {
            {
                let s = val.to_string();
                start.push_attribute(("activeTab", s.as_str()));
            }
        }
        if let Some(ref val) = self.auto_filter_date_grouping {
            start.push_attribute(("autoFilterDateGrouping", if *val { "1" } else { "0" }));
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
        if let Some(ref val) = self.extension_list {
            val.write_element("extLst", writer)?;
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
        if self.extension_list.is_some() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for CustomWorkbookViews {
    fn write_children<W: Write>(&self, writer: &mut Writer<W>) -> Result<(), SerializeError> {
        #[cfg(feature = "extra-children")]
        let mut extra_iter = self.extra_children.iter().peekable();
        #[cfg(feature = "extra-children")]
        let mut emit_idx: usize = 0;
        for item in &self.custom_workbook_view {
            #[cfg(feature = "extra-children")]
            while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
                extra_iter
                    .next()
                    .unwrap()
                    .node
                    .write_to(writer)
                    .map_err(SerializeError::from)?;
            }
            item.write_element("customWorkbookView", writer)?;
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
        if !self.custom_workbook_view.is_empty() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for CustomWorkbookView {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        {
            let val = &self.name;
            start.push_attribute(("name", val.as_str()));
        }
        {
            let val = &self.guid;
            start.push_attribute(("guid", val.as_str()));
        }
        if let Some(ref val) = self.auto_update {
            start.push_attribute(("autoUpdate", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.merge_interval {
            {
                let s = val.to_string();
                start.push_attribute(("mergeInterval", s.as_str()));
            }
        }
        if let Some(ref val) = self.changes_saved_win {
            start.push_attribute(("changesSavedWin", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.only_sync {
            start.push_attribute(("onlySync", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.personal_view {
            start.push_attribute(("personalView", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.include_print_settings {
            start.push_attribute(("includePrintSettings", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.include_hidden_row_col {
            start.push_attribute(("includeHiddenRowCol", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.maximized {
            start.push_attribute(("maximized", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.minimized {
            start.push_attribute(("minimized", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.show_horizontal_scroll {
            start.push_attribute(("showHorizontalScroll", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.show_vertical_scroll {
            start.push_attribute(("showVerticalScroll", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.show_sheet_tabs {
            start.push_attribute(("showSheetTabs", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.x_window {
            {
                let s = val.to_string();
                start.push_attribute(("xWindow", s.as_str()));
            }
        }
        if let Some(ref val) = self.y_window {
            {
                let s = val.to_string();
                start.push_attribute(("yWindow", s.as_str()));
            }
        }
        {
            let val = &self.window_width;
            {
                let s = val.to_string();
                start.push_attribute(("windowWidth", s.as_str()));
            }
        }
        {
            let val = &self.window_height;
            {
                let s = val.to_string();
                start.push_attribute(("windowHeight", s.as_str()));
            }
        }
        if let Some(ref val) = self.tab_ratio {
            {
                let s = val.to_string();
                start.push_attribute(("tabRatio", s.as_str()));
            }
        }
        {
            let val = &self.active_sheet_id;
            {
                let s = val.to_string();
                start.push_attribute(("activeSheetId", s.as_str()));
            }
        }
        if let Some(ref val) = self.show_formula_bar {
            start.push_attribute(("showFormulaBar", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.show_statusbar {
            start.push_attribute(("showStatusbar", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.show_comments {
            {
                let s = val.to_string();
                start.push_attribute(("showComments", s.as_str()));
            }
        }
        if let Some(ref val) = self.show_objects {
            {
                let s = val.to_string();
                start.push_attribute(("showObjects", s.as_str()));
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
        if let Some(ref val) = self.extension_list {
            val.write_element("extLst", writer)?;
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
        if self.extension_list.is_some() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for Sheets {
    fn write_children<W: Write>(&self, writer: &mut Writer<W>) -> Result<(), SerializeError> {
        #[cfg(feature = "extra-children")]
        let mut extra_iter = self.extra_children.iter().peekable();
        #[cfg(feature = "extra-children")]
        let mut emit_idx: usize = 0;
        for item in &self.sheet {
            #[cfg(feature = "extra-children")]
            while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
                extra_iter
                    .next()
                    .unwrap()
                    .node
                    .write_to(writer)
                    .map_err(SerializeError::from)?;
            }
            item.write_element("sheet", writer)?;
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
        if !self.sheet.is_empty() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for Sheet {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        {
            let val = &self.name;
            start.push_attribute(("name", val.as_str()));
        }
        {
            let val = &self.sheet_id;
            {
                let s = val.to_string();
                start.push_attribute(("sheetId", s.as_str()));
            }
        }
        #[cfg(feature = "sml-structure")]
        if let Some(ref val) = self.state {
            {
                let s = val.to_string();
                start.push_attribute(("state", s.as_str()));
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

impl ToXml for WorkbookProperties {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        if let Some(ref val) = self.date1904 {
            start.push_attribute(("date1904", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.show_objects {
            {
                let s = val.to_string();
                start.push_attribute(("showObjects", s.as_str()));
            }
        }
        if let Some(ref val) = self.show_border_unselected_tables {
            start.push_attribute(("showBorderUnselectedTables", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.filter_privacy {
            start.push_attribute(("filterPrivacy", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.prompted_solutions {
            start.push_attribute(("promptedSolutions", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.show_ink_annotation {
            start.push_attribute(("showInkAnnotation", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.backup_file {
            start.push_attribute(("backupFile", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.save_external_link_values {
            start.push_attribute(("saveExternalLinkValues", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.update_links {
            {
                let s = val.to_string();
                start.push_attribute(("updateLinks", s.as_str()));
            }
        }
        if let Some(ref val) = self.code_name {
            start.push_attribute(("codeName", val.as_str()));
        }
        if let Some(ref val) = self.hide_pivot_field_list {
            start.push_attribute(("hidePivotFieldList", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.show_pivot_chart_filter {
            start.push_attribute(("showPivotChartFilter", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.allow_refresh_query {
            start.push_attribute(("allowRefreshQuery", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.publish_items {
            start.push_attribute(("publishItems", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.check_compatibility {
            start.push_attribute(("checkCompatibility", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.auto_compress_pictures {
            start.push_attribute(("autoCompressPictures", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.refresh_all_connections {
            start.push_attribute(("refreshAllConnections", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.default_theme_version {
            {
                let s = val.to_string();
                start.push_attribute(("defaultThemeVersion", s.as_str()));
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

impl ToXml for CTSmartTagPr {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        if let Some(ref val) = self.embed {
            start.push_attribute(("embed", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.show {
            {
                let s = val.to_string();
                start.push_attribute(("show", s.as_str()));
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

impl ToXml for CTSmartTagTypes {
    fn write_children<W: Write>(&self, writer: &mut Writer<W>) -> Result<(), SerializeError> {
        #[cfg(feature = "extra-children")]
        let mut extra_iter = self.extra_children.iter().peekable();
        #[cfg(feature = "extra-children")]
        let mut emit_idx: usize = 0;
        for item in &self.smart_tag_type {
            #[cfg(feature = "extra-children")]
            while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
                extra_iter
                    .next()
                    .unwrap()
                    .node
                    .write_to(writer)
                    .map_err(SerializeError::from)?;
            }
            item.write_element("smartTagType", writer)?;
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
        if !self.smart_tag_type.is_empty() {
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
        if let Some(ref val) = self.namespace_uri {
            start.push_attribute(("namespaceUri", val.as_str()));
        }
        if let Some(ref val) = self.name {
            start.push_attribute(("name", val.as_str()));
        }
        if let Some(ref val) = self.url {
            start.push_attribute(("url", val.as_str()));
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

impl ToXml for FileRecoveryProperties {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        if let Some(ref val) = self.auto_recover {
            start.push_attribute(("autoRecover", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.crash_save {
            start.push_attribute(("crashSave", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.data_extract_load {
            start.push_attribute(("dataExtractLoad", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.repair_load {
            start.push_attribute(("repairLoad", if *val { "1" } else { "0" }));
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

impl ToXml for CalculationProperties {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        if let Some(ref val) = self.calc_id {
            {
                let s = val.to_string();
                start.push_attribute(("calcId", s.as_str()));
            }
        }
        if let Some(ref val) = self.calc_mode {
            {
                let s = val.to_string();
                start.push_attribute(("calcMode", s.as_str()));
            }
        }
        if let Some(ref val) = self.full_calc_on_load {
            start.push_attribute(("fullCalcOnLoad", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.ref_mode {
            {
                let s = val.to_string();
                start.push_attribute(("refMode", s.as_str()));
            }
        }
        if let Some(ref val) = self.iterate {
            start.push_attribute(("iterate", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.iterate_count {
            {
                let s = val.to_string();
                start.push_attribute(("iterateCount", s.as_str()));
            }
        }
        if let Some(ref val) = self.iterate_delta {
            {
                let s = val.to_string();
                start.push_attribute(("iterateDelta", s.as_str()));
            }
        }
        if let Some(ref val) = self.full_precision {
            start.push_attribute(("fullPrecision", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.calc_completed {
            start.push_attribute(("calcCompleted", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.calc_on_save {
            start.push_attribute(("calcOnSave", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.concurrent_calc {
            start.push_attribute(("concurrentCalc", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.concurrent_manual_count {
            {
                let s = val.to_string();
                start.push_attribute(("concurrentManualCount", s.as_str()));
            }
        }
        if let Some(ref val) = self.force_full_calc {
            start.push_attribute(("forceFullCalc", if *val { "1" } else { "0" }));
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

impl ToXml for DefinedNames {
    fn write_children<W: Write>(&self, writer: &mut Writer<W>) -> Result<(), SerializeError> {
        #[cfg(feature = "extra-children")]
        let mut extra_iter = self.extra_children.iter().peekable();
        #[cfg(feature = "extra-children")]
        let mut emit_idx: usize = 0;
        for item in &self.defined_name {
            #[cfg(feature = "extra-children")]
            while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
                extra_iter
                    .next()
                    .unwrap()
                    .node
                    .write_to(writer)
                    .map_err(SerializeError::from)?;
            }
            item.write_element("definedName", writer)?;
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
        if !self.defined_name.is_empty() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for DefinedName {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        {
            let val = &self.name;
            start.push_attribute(("name", val.as_str()));
        }
        if let Some(ref val) = self.comment {
            start.push_attribute(("comment", val.as_str()));
        }
        #[cfg(feature = "sml-formulas-advanced")]
        if let Some(ref val) = self.custom_menu {
            start.push_attribute(("customMenu", val.as_str()));
        }
        if let Some(ref val) = self.description {
            start.push_attribute(("description", val.as_str()));
        }
        #[cfg(feature = "sml-formulas-advanced")]
        if let Some(ref val) = self.help {
            start.push_attribute(("help", val.as_str()));
        }
        #[cfg(feature = "sml-formulas-advanced")]
        if let Some(ref val) = self.status_bar {
            start.push_attribute(("statusBar", val.as_str()));
        }
        if let Some(ref val) = self.local_sheet_id {
            {
                let s = val.to_string();
                start.push_attribute(("localSheetId", s.as_str()));
            }
        }
        #[cfg(feature = "sml-structure")]
        if let Some(ref val) = self.hidden {
            start.push_attribute(("hidden", if *val { "1" } else { "0" }));
        }
        #[cfg(feature = "sml-formulas-advanced")]
        if let Some(ref val) = self.function {
            start.push_attribute(("function", if *val { "1" } else { "0" }));
        }
        #[cfg(feature = "sml-formulas-advanced")]
        if let Some(ref val) = self.vb_procedure {
            start.push_attribute(("vbProcedure", if *val { "1" } else { "0" }));
        }
        #[cfg(feature = "sml-formulas-advanced")]
        if let Some(ref val) = self.xlm {
            start.push_attribute(("xlm", if *val { "1" } else { "0" }));
        }
        #[cfg(feature = "sml-formulas-advanced")]
        if let Some(ref val) = self.function_group_id {
            {
                let s = val.to_string();
                start.push_attribute(("functionGroupId", s.as_str()));
            }
        }
        #[cfg(feature = "sml-formulas-advanced")]
        if let Some(ref val) = self.shortcut_key {
            start.push_attribute(("shortcutKey", val.as_str()));
        }
        #[cfg(feature = "sml-external")]
        if let Some(ref val) = self.publish_to_server {
            start.push_attribute(("publishToServer", if *val { "1" } else { "0" }));
        }
        #[cfg(feature = "sml-formulas-advanced")]
        if let Some(ref val) = self.workbook_parameter {
            start.push_attribute(("workbookParameter", if *val { "1" } else { "0" }));
        }
        #[cfg(feature = "extra-attrs")]
        for (key, value) in &self.extra_attrs {
            start.push_attribute((key.as_str(), value.as_str()));
        }
        start
    }

    fn write_children<W: Write>(&self, writer: &mut Writer<W>) -> Result<(), SerializeError> {
        if let Some(ref text) = self.text {
            writer.write_event(Event::Text(BytesText::new(text)))?;
        }
        #[cfg(feature = "extra-children")]
        for extra in &self.extra_children {
            extra.node.write_to(writer).map_err(SerializeError::from)?;
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

impl ToXml for ExternalReferences {
    fn write_children<W: Write>(&self, writer: &mut Writer<W>) -> Result<(), SerializeError> {
        #[cfg(feature = "extra-children")]
        let mut extra_iter = self.extra_children.iter().peekable();
        #[cfg(feature = "extra-children")]
        let mut emit_idx: usize = 0;
        for item in &self.external_reference {
            #[cfg(feature = "extra-children")]
            while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
                extra_iter
                    .next()
                    .unwrap()
                    .node
                    .write_to(writer)
                    .map_err(SerializeError::from)?;
            }
            item.write_element("externalReference", writer)?;
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
        if !self.external_reference.is_empty() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for ExternalReference {
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

impl ToXml for SheetBackgroundPicture {
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

impl ToXml for PivotCaches {
    fn write_children<W: Write>(&self, writer: &mut Writer<W>) -> Result<(), SerializeError> {
        #[cfg(feature = "extra-children")]
        let mut extra_iter = self.extra_children.iter().peekable();
        #[cfg(feature = "extra-children")]
        let mut emit_idx: usize = 0;
        for item in &self.pivot_cache {
            #[cfg(feature = "extra-children")]
            while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
                extra_iter
                    .next()
                    .unwrap()
                    .node
                    .write_to(writer)
                    .map_err(SerializeError::from)?;
            }
            item.write_element("pivotCache", writer)?;
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
        if !self.pivot_cache.is_empty() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for CTPivotCache {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        {
            let val = &self.cache_id;
            {
                let s = val.to_string();
                start.push_attribute(("cacheId", s.as_str()));
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

impl ToXml for FileSharing {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        if let Some(ref val) = self.read_only_recommended {
            start.push_attribute(("readOnlyRecommended", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.user_name {
            start.push_attribute(("userName", val.as_str()));
        }
        if let Some(ref val) = self.reservation_password {
            {
                let hex = encode_hex(val);
                start.push_attribute(("reservationPassword", hex.as_str()));
            }
        }
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
        if let Some(ref val) = self.spin_count {
            {
                let s = val.to_string();
                start.push_attribute(("spinCount", s.as_str()));
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

impl ToXml for CTOleSize {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        {
            let val = &self.reference;
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

impl ToXml for WorkbookProtection {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        if let Some(ref val) = self.workbook_password {
            {
                let hex = encode_hex(val);
                start.push_attribute(("workbookPassword", hex.as_str()));
            }
        }
        if let Some(ref val) = self.workbook_password_character_set {
            start.push_attribute(("workbookPasswordCharacterSet", val.as_str()));
        }
        if let Some(ref val) = self.revisions_password {
            {
                let hex = encode_hex(val);
                start.push_attribute(("revisionsPassword", hex.as_str()));
            }
        }
        if let Some(ref val) = self.revisions_password_character_set {
            start.push_attribute(("revisionsPasswordCharacterSet", val.as_str()));
        }
        if let Some(ref val) = self.lock_structure {
            start.push_attribute(("lockStructure", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.lock_windows {
            start.push_attribute(("lockWindows", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.lock_revision {
            start.push_attribute(("lockRevision", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.revisions_algorithm_name {
            start.push_attribute(("revisionsAlgorithmName", val.as_str()));
        }
        if let Some(ref val) = self.revisions_hash_value {
            {
                let hex = encode_hex(val);
                start.push_attribute(("revisionsHashValue", hex.as_str()));
            }
        }
        if let Some(ref val) = self.revisions_salt_value {
            {
                let hex = encode_hex(val);
                start.push_attribute(("revisionsSaltValue", hex.as_str()));
            }
        }
        if let Some(ref val) = self.revisions_spin_count {
            {
                let s = val.to_string();
                start.push_attribute(("revisionsSpinCount", s.as_str()));
            }
        }
        if let Some(ref val) = self.workbook_algorithm_name {
            start.push_attribute(("workbookAlgorithmName", val.as_str()));
        }
        if let Some(ref val) = self.workbook_hash_value {
            {
                let hex = encode_hex(val);
                start.push_attribute(("workbookHashValue", hex.as_str()));
            }
        }
        if let Some(ref val) = self.workbook_salt_value {
            {
                let hex = encode_hex(val);
                start.push_attribute(("workbookSaltValue", hex.as_str()));
            }
        }
        if let Some(ref val) = self.workbook_spin_count {
            {
                let s = val.to_string();
                start.push_attribute(("workbookSpinCount", s.as_str()));
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

impl ToXml for WebPublishing {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        if let Some(ref val) = self.css {
            start.push_attribute(("css", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.thicket {
            start.push_attribute(("thicket", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.long_file_names {
            start.push_attribute(("longFileNames", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.vml {
            start.push_attribute(("vml", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.allow_png {
            start.push_attribute(("allowPng", if *val { "1" } else { "0" }));
        }
        if let Some(ref val) = self.target_screen_size {
            {
                let s = val.to_string();
                start.push_attribute(("targetScreenSize", s.as_str()));
            }
        }
        if let Some(ref val) = self.dpi {
            {
                let s = val.to_string();
                start.push_attribute(("dpi", s.as_str()));
            }
        }
        if let Some(ref val) = self.code_page {
            {
                let s = val.to_string();
                start.push_attribute(("codePage", s.as_str()));
            }
        }
        if let Some(ref val) = self.character_set {
            start.push_attribute(("characterSet", val.as_str()));
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

impl ToXml for CTFunctionGroups {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        if let Some(ref val) = self.built_in_group_count {
            {
                let s = val.to_string();
                start.push_attribute(("builtInGroupCount", s.as_str()));
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
        for item in &self.function_group {
            #[cfg(feature = "extra-children")]
            while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
                extra_iter
                    .next()
                    .unwrap()
                    .node
                    .write_to(writer)
                    .map_err(SerializeError::from)?;
            }
            item.write_element("functionGroup", writer)?;
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
        if !self.function_group.is_empty() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for CTFunctionGroup {
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

    fn is_empty_element(&self) -> bool {
        true
    }
}

impl ToXml for CTWebPublishObjects {
    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {
        let mut start = start;
        if let Some(ref val) = self.count {
            {
                let s = val.to_string();
                start.push_attribute(("count", s.as_str()));
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
        for item in &self.web_publish_object {
            #[cfg(feature = "extra-children")]
            while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {
                extra_iter
                    .next()
                    .unwrap()
                    .node
                    .write_to(writer)
                    .map_err(SerializeError::from)?;
            }
            item.write_element("webPublishObject", writer)?;
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
        if !self.web_publish_object.is_empty() {
            return false;
        }
        #[cfg(feature = "extra-children")]
        if !self.extra_children.is_empty() {
            return false;
        }
        true
    }
}

impl ToXml for CTWebPublishObject {
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
            let val = &self.div_id;
            start.push_attribute(("divId", val.as_str()));
        }
        if let Some(ref val) = self.source_object {
            start.push_attribute(("sourceObject", val.as_str()));
        }
        {
            let val = &self.destination_file;
            start.push_attribute(("destinationFile", val.as_str()));
        }
        if let Some(ref val) = self.title {
            start.push_attribute(("title", val.as_str()));
        }
        if let Some(ref val) = self.auto_republish {
            start.push_attribute(("autoRepublish", if *val { "1" } else { "0" }));
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
