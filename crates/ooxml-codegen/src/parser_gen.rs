//! Event-based XML parser generator.
//!
//! This module generates quick-xml event-based parsers for OOXML types,
//! which are ~3x faster than serde-based deserialization.

use crate::ast::{Pattern, QName, Schema};
use crate::codegen::CodegenConfig;
use std::collections::HashMap;
use std::fmt::Write;

/// Generate parser code for all types in the schema.
pub fn generate_parsers(schema: &Schema, config: &CodegenConfig) -> String {
    let mut g = ParserGenerator::new(schema, config);
    g.run()
}

/// How to parse a field value from XML.
#[derive(Debug, Clone, Copy, PartialEq)]
enum ParseStrategy {
    /// Call from_xml on a complex type (CT_* or EG_*)
    FromXml,
    /// Read text content and use FromStr (enums, numbers)
    TextFromStr,
    /// Read text content as String directly
    TextString,
    /// Read text content and decode as hex (Vec<u8>)
    TextHexBinary,
}

struct ParserGenerator<'a> {
    schema: &'a Schema,
    config: &'a CodegenConfig,
    output: String,
    definitions: HashMap<&'a str, &'a Pattern>,
}

impl<'a> ParserGenerator<'a> {
    fn new(schema: &'a Schema, config: &'a CodegenConfig) -> Self {
        let definitions = schema
            .definitions
            .iter()
            .map(|d| (d.name.as_str(), &d.pattern))
            .collect();

        Self {
            schema,
            config,
            output: String::new(),
            definitions,
        }
    }

    fn run(&mut self) -> String {
        self.write_header();

        // Generate parsers for complex types only
        for def in &self.schema.definitions {
            if !def.name.contains("_ST_") && !self.is_simple_type(&def.pattern) {
                if def.name.contains("_EG_") && self.is_element_choice(&def.pattern) {
                    // Element group - generate enum parser
                    if let Some(code) = self.gen_element_group_parser(def) {
                        self.output.push_str(&code);
                        self.output.push('\n');
                    }
                } else if !self.is_type_alias(&def.pattern) {
                    // Complex type struct - generate struct parser (skip type aliases)
                    if let Some(code) = self.gen_struct_parser(def) {
                        self.output.push_str(&code);
                        self.output.push('\n');
                    }
                }
            }
        }

        std::mem::take(&mut self.output)
    }

    /// Check if a pattern would generate a type alias rather than a struct.
    /// Type aliases are generated for:
    /// - Pattern::Element { pattern } (element-only definitions)
    /// - Pattern::Datatype (XSD type alias)
    /// - Pattern::Ref to a simple type, datatype, or unknown definition
    ///
    /// We DON'T skip Ref patterns that point to attribute definitions,
    /// because those generate structs that need FromXml impls.
    fn is_type_alias(&self, pattern: &Pattern) -> bool {
        match pattern {
            // Element wrappers are always type aliases
            Pattern::Element { .. } => true,
            // Direct datatype references are type aliases
            Pattern::Datatype { .. } => true,
            // Refs need to be checked - only skip if they resolve to simple types
            Pattern::Ref(name) => {
                if let Some(def_pattern) = self.definitions.get(name.as_str()) {
                    // If it resolves to a simple type (choice of strings, datatype, etc.)
                    self.is_simple_type(def_pattern)
                        // Or to another type alias
                        || self.is_type_alias(def_pattern)
                } else {
                    // Unknown ref (from another schema) - these generate empty structs
                    // that still need FromXml impls
                    false
                }
            }
            _ => false,
        }
    }

    fn write_header(&mut self) {
        writeln!(self.output, "// Event-based parsers for generated types.").unwrap();
        writeln!(
            self.output,
            "// ~3x faster than serde-based deserialization."
        )
        .unwrap();
        writeln!(self.output).unwrap();
        writeln!(self.output, "#![allow(unused_variables)]").unwrap();
        writeln!(self.output, "#![allow(clippy::single_match)]").unwrap();
        writeln!(self.output, "#![allow(clippy::match_single_binding)]").unwrap();
        writeln!(self.output, "#![allow(clippy::manual_is_multiple_of)]").unwrap();
        writeln!(self.output).unwrap();
        writeln!(self.output, "use super::generated::*;").unwrap();
        writeln!(self.output, "use quick_xml::Reader;").unwrap();
        writeln!(self.output, "use quick_xml::events::{{Event, BytesStart}};").unwrap();
        writeln!(self.output, "use std::io::BufRead;").unwrap();
        writeln!(self.output, "#[cfg(feature = \"extra-children\")]").unwrap();
        writeln!(self.output, "use ooxml_xml::{{RawXmlElement, RawXmlNode}};").unwrap();
        writeln!(self.output).unwrap();
        writeln!(self.output, "/// Error type for XML parsing.").unwrap();
        writeln!(self.output, "#[derive(Debug)]").unwrap();
        writeln!(self.output, "pub enum ParseError {{").unwrap();
        writeln!(self.output, "    Xml(quick_xml::Error),").unwrap();
        writeln!(self.output, "    #[cfg(feature = \"extra-children\")]").unwrap();
        writeln!(self.output, "    RawXml(ooxml_xml::Error),").unwrap();
        writeln!(self.output, "    UnexpectedElement(String),").unwrap();
        writeln!(self.output, "    MissingAttribute(String),").unwrap();
        writeln!(self.output, "    InvalidValue(String),").unwrap();
        writeln!(self.output, "}}").unwrap();
        writeln!(self.output).unwrap();
        writeln!(self.output, "impl From<quick_xml::Error> for ParseError {{").unwrap();
        writeln!(self.output, "    fn from(e: quick_xml::Error) -> Self {{").unwrap();
        writeln!(self.output, "        ParseError::Xml(e)").unwrap();
        writeln!(self.output, "    }}").unwrap();
        writeln!(self.output, "}}").unwrap();
        writeln!(self.output).unwrap();
        writeln!(self.output, "#[cfg(feature = \"extra-children\")]").unwrap();
        writeln!(self.output, "impl From<ooxml_xml::Error> for ParseError {{").unwrap();
        writeln!(self.output, "    fn from(e: ooxml_xml::Error) -> Self {{").unwrap();
        writeln!(self.output, "        ParseError::RawXml(e)").unwrap();
        writeln!(self.output, "    }}").unwrap();
        writeln!(self.output, "}}").unwrap();
        writeln!(self.output).unwrap();
        writeln!(
            self.output,
            "/// Trait for types that can be parsed from XML events."
        )
        .unwrap();
        writeln!(self.output, "pub trait FromXml: Sized {{").unwrap();
        writeln!(
            self.output,
            "    /// Parse from a reader, given the opening tag."
        )
        .unwrap();
        writeln!(
            self.output,
            "    /// If `is_empty` is true, the element was self-closing (no children to read)."
        )
        .unwrap();
        writeln!(
            self.output,
            "    fn from_xml<R: BufRead>(reader: &mut Reader<R>, start: &BytesStart, is_empty: bool) -> Result<Self, ParseError>;"
        )
        .unwrap();
        writeln!(self.output, "}}").unwrap();
        writeln!(self.output).unwrap();
        // Add skip_element helper (allow dead_code since extra-children feature captures instead)
        writeln!(self.output, "#[allow(dead_code)]").unwrap();
        writeln!(self.output, "/// Skip an element and all its children.").unwrap();
        writeln!(
            self.output,
            "fn skip_element<R: BufRead>(reader: &mut Reader<R>) -> Result<(), ParseError> {{"
        )
        .unwrap();
        writeln!(self.output, "    let mut depth = 1u32;").unwrap();
        writeln!(self.output, "    let mut buf = Vec::new();").unwrap();
        writeln!(self.output, "    loop {{").unwrap();
        writeln!(
            self.output,
            "        match reader.read_event_into(&mut buf)? {{"
        )
        .unwrap();
        writeln!(self.output, "            Event::Start(_) => depth += 1,").unwrap();
        writeln!(self.output, "            Event::End(_) => {{").unwrap();
        writeln!(self.output, "                depth -= 1;").unwrap();
        writeln!(self.output, "                if depth == 0 {{ break; }}").unwrap();
        writeln!(self.output, "            }}").unwrap();
        writeln!(self.output, "            Event::Eof => break,").unwrap();
        writeln!(self.output, "            _ => {{}}").unwrap();
        writeln!(self.output, "        }}").unwrap();
        writeln!(self.output, "        buf.clear();").unwrap();
        writeln!(self.output, "    }}").unwrap();
        writeln!(self.output, "    Ok(())").unwrap();
        writeln!(self.output, "}}").unwrap();
        writeln!(self.output).unwrap();
        // Add read_text_content helper for reading element text
        writeln!(
            self.output,
            "/// Read the text content of an element until its end tag."
        )
        .unwrap();
        writeln!(self.output, "fn read_text_content<R: BufRead>(reader: &mut Reader<R>) -> Result<String, ParseError> {{").unwrap();
        writeln!(self.output, "    let mut text = String::new();").unwrap();
        writeln!(self.output, "    let mut buf = Vec::new();").unwrap();
        writeln!(self.output, "    loop {{").unwrap();
        writeln!(
            self.output,
            "        match reader.read_event_into(&mut buf)? {{"
        )
        .unwrap();
        writeln!(
            self.output,
            "            Event::Text(e) => text.push_str(&e.decode().unwrap_or_default()),"
        )
        .unwrap();
        writeln!(
            self.output,
            "            Event::CData(e) => text.push_str(&e.decode().unwrap_or_default()),"
        )
        .unwrap();
        writeln!(self.output, "            Event::End(_) => break,").unwrap();
        writeln!(self.output, "            Event::Eof => break,").unwrap();
        writeln!(self.output, "            _ => {{}}").unwrap();
        writeln!(self.output, "        }}").unwrap();
        writeln!(self.output, "        buf.clear();").unwrap();
        writeln!(self.output, "    }}").unwrap();
        writeln!(self.output, "    Ok(text)").unwrap();
        writeln!(self.output, "}}").unwrap();
        writeln!(self.output).unwrap();
        // Add decode_hex helper for hexBinary types
        writeln!(self.output, "/// Decode a hex string to bytes.").unwrap();
        writeln!(self.output, "fn decode_hex(s: &str) -> Option<Vec<u8>> {{").unwrap();
        writeln!(self.output, "    let s = s.trim();").unwrap();
        writeln!(self.output, "    if s.len() % 2 != 0 {{ return None; }}").unwrap();
        writeln!(self.output, "    (0..s.len())").unwrap();
        writeln!(self.output, "        .step_by(2)").unwrap();
        writeln!(
            self.output,
            "        .map(|i| u8::from_str_radix(&s[i..i + 2], 16).ok())"
        )
        .unwrap();
        writeln!(self.output, "        .collect()").unwrap();
        writeln!(self.output, "}}").unwrap();
        writeln!(self.output).unwrap();
    }

    fn is_simple_type(&self, pattern: &Pattern) -> bool {
        match pattern {
            Pattern::Choice(variants) => variants
                .iter()
                .all(|v| matches!(v, Pattern::StringLiteral(_))),
            Pattern::StringLiteral(_) => true,
            Pattern::Datatype { .. } => true,
            Pattern::List(_) => true,
            Pattern::Ref(name) => self
                .definitions
                .get(name.as_str())
                .is_some_and(|p| self.is_simple_type(p)),
            _ => false,
        }
    }

    fn is_element_choice(&self, pattern: &Pattern) -> bool {
        match pattern {
            Pattern::Choice(variants) => variants.iter().any(Self::is_direct_element_variant),
            _ => false,
        }
    }

    fn is_direct_element_variant(pattern: &Pattern) -> bool {
        match pattern {
            Pattern::Element { .. } => true,
            Pattern::Optional(inner) => Self::is_direct_element_variant(inner),
            _ => false,
        }
    }

    fn gen_element_group_parser(&self, def: &crate::ast::Definition) -> Option<String> {
        let rust_name = self.to_rust_type_name(&def.name);

        let Pattern::Choice(variants) = &def.pattern else {
            return None;
        };

        // Collect element variants
        let element_variants: Vec<_> = variants
            .iter()
            .filter_map(|v| self.extract_element_variant(v))
            .collect();

        if element_variants.is_empty() {
            return None;
        }

        let mut code = String::new();
        writeln!(code, "impl FromXml for {} {{", rust_name).unwrap();
        writeln!(
            code,
            "    fn from_xml<R: BufRead>(reader: &mut Reader<R>, start_tag: &BytesStart, is_empty: bool) -> Result<Self, ParseError> {{"
        )
        .unwrap();
        writeln!(code, "        let tag = start_tag.name();").unwrap();
        writeln!(code, "        match tag.as_ref() {{").unwrap();

        for (xml_name, inner_type, needs_box) in &element_variants {
            let variant_name = self.to_rust_variant_name(xml_name);
            writeln!(code, "            b\"{}\" => {{", xml_name).unwrap();
            if *needs_box {
                writeln!(
                    code,
                    "                let inner = {}::from_xml(reader, start_tag, is_empty)?;",
                    inner_type
                )
                .unwrap();
                writeln!(
                    code,
                    "                Ok(Self::{}(Box::new(inner)))",
                    variant_name
                )
                .unwrap();
            } else {
                writeln!(
                    code,
                    "                let inner = {}::from_xml(reader, start_tag, is_empty)?;",
                    inner_type
                )
                .unwrap();
                writeln!(code, "                Ok(Self::{}(inner))", variant_name).unwrap();
            }
            writeln!(code, "            }}").unwrap();
        }

        writeln!(code, "            _ => Err(ParseError::UnexpectedElement(").unwrap();
        writeln!(
            code,
            "                String::from_utf8_lossy(tag.as_ref()).into_owned()"
        )
        .unwrap();
        writeln!(code, "            )),").unwrap();
        writeln!(code, "        }}").unwrap();
        writeln!(code, "    }}").unwrap();
        writeln!(code, "}}").unwrap();

        Some(code)
    }

    fn extract_element_variant(&self, pattern: &Pattern) -> Option<(String, String, bool)> {
        match pattern {
            Pattern::Element { name, pattern } => {
                let (inner_type, needs_box) = self.pattern_to_rust_type(pattern);
                Some((name.local.clone(), inner_type, needs_box))
            }
            Pattern::Optional(inner) => self.extract_element_variant(inner),
            Pattern::ZeroOrMore(inner) | Pattern::OneOrMore(inner) => {
                self.extract_element_variant(inner)
            }
            _ => None,
        }
    }

    fn gen_struct_parser(&self, def: &crate::ast::Definition) -> Option<String> {
        let rust_name = self.to_rust_type_name(&def.name);
        let fields = self.extract_fields(&def.pattern);

        if fields.is_empty() {
            // Empty struct - skip all children with depth tracking
            let mut code = String::new();
            writeln!(code, "impl FromXml for {} {{", rust_name).unwrap();
            writeln!(
                code,
                "    fn from_xml<R: BufRead>(reader: &mut Reader<R>, _start: &BytesStart, is_empty: bool) -> Result<Self, ParseError> {{"
            )
            .unwrap();
            writeln!(code, "        if !is_empty {{").unwrap();
            writeln!(
                code,
                "            // Skip to matching end tag with depth tracking"
            )
            .unwrap();
            writeln!(code, "            let mut buf = Vec::new();").unwrap();
            writeln!(code, "            let mut depth = 1u32;").unwrap();
            writeln!(code, "            loop {{").unwrap();
            writeln!(
                code,
                "                match reader.read_event_into(&mut buf)? {{"
            )
            .unwrap();
            writeln!(code, "                    Event::Start(_) => depth += 1,").unwrap();
            writeln!(
                code,
                "                    Event::End(_) => {{ depth -= 1; if depth == 0 {{ break; }} }}"
            )
            .unwrap();
            writeln!(code, "                    Event::Eof => break,").unwrap();
            writeln!(code, "                    _ => {{}}").unwrap();
            writeln!(code, "                }}").unwrap();
            writeln!(code, "                buf.clear();").unwrap();
            writeln!(code, "            }}").unwrap();
            writeln!(code, "        }}").unwrap();
            writeln!(code, "        Ok(Self {{}})").unwrap();
            writeln!(code, "    }}").unwrap();
            writeln!(code, "}}").unwrap();
            return Some(code);
        }

        let mut code = String::new();
        writeln!(code, "impl FromXml for {} {{", rust_name).unwrap();
        writeln!(
            code,
            "    fn from_xml<R: BufRead>(reader: &mut Reader<R>, start_tag: &BytesStart, is_empty: bool) -> Result<Self, ParseError> {{"
        )
        .unwrap();

        // Declare field variables with f_ prefix to avoid shadowing function parameters
        for field in &fields {
            // Strip r# from raw identifiers and leading underscores before prefixing
            let base_name = field.name.strip_prefix("r#").unwrap_or(&field.name);
            let base_name = base_name.trim_start_matches('_');
            let var_name = format!("f_{}", base_name);

            // Add cfg attribute for feature-gated fields
            if let Some(ref feature) = self.get_field_feature(&rust_name, &field.xml_name) {
                write!(code, "        #[cfg(feature = \"{}\")] ", feature).unwrap();
            } else {
                write!(code, "        ").unwrap();
            }

            if field.is_vec {
                writeln!(code, "let mut {} = Vec::new();", var_name).unwrap();
            } else if field.is_optional {
                writeln!(code, "let mut {} = None;", var_name).unwrap();
            } else {
                let (rust_type, needs_box) = self.pattern_to_rust_type(&field.pattern);
                let full_type = if needs_box {
                    format!("Box<{}>", rust_type)
                } else {
                    rust_type
                };
                writeln!(code, "let mut {}: Option<{}> = None;", var_name, full_type).unwrap();
            }
        }

        // Parse attributes
        let attr_fields: Vec<_> = fields.iter().filter(|f| f.is_attribute).collect();
        let has_attrs = !attr_fields.is_empty();
        let elem_fields: Vec<_> = fields
            .iter()
            .filter(|f| !f.is_attribute && !f.is_text_content)
            .collect();
        let text_fields: Vec<_> = fields.iter().filter(|f| f.is_text_content).collect();
        let has_children = !elem_fields.is_empty();
        let has_parsing_loop = has_children || !text_fields.is_empty();
        if has_attrs {
            // Declare extra_attrs for capturing unknown attributes (feature-gated)
            writeln!(code, "        #[cfg(feature = \"extra-attrs\")]").unwrap();
            writeln!(
                code,
                "        let mut extra_attrs = std::collections::HashMap::new();"
            )
            .unwrap();
        }
        if has_parsing_loop {
            // Declare extra_children for capturing unknown child elements (feature-gated)
            writeln!(code, "        #[cfg(feature = \"extra-children\")]").unwrap();
            writeln!(code, "        let mut extra_children = Vec::new();").unwrap();
        }
        if has_attrs || has_parsing_loop {
            writeln!(code).unwrap();
        }
        if has_attrs {
            writeln!(code, "        // Parse attributes").unwrap();
            writeln!(
                code,
                "        for attr in start_tag.attributes().filter_map(|a| a.ok()) {{"
            )
            .unwrap();
            writeln!(
                code,
                "            let val = String::from_utf8_lossy(&attr.value);"
            )
            .unwrap();
            writeln!(code, "            match attr.key.as_ref() {{").unwrap();
            for field in &attr_fields {
                let base_name = field.name.strip_prefix("r#").unwrap_or(&field.name);
                let base_name = base_name.trim_start_matches('_');
                let var_name = format!("f_{}", base_name);
                let parse_expr = self.gen_attr_parse_expr(&field.pattern);

                // Add cfg attribute for feature-gated fields
                if let Some(ref feature) = self.get_field_feature(&rust_name, &field.xml_name) {
                    writeln!(code, "                #[cfg(feature = \"{}\")]", feature).unwrap();
                }
                writeln!(code, "                b\"{}\" => {{", field.xml_name).unwrap();
                writeln!(code, "                    {} = {};", var_name, parse_expr).unwrap();
                writeln!(code, "                }}").unwrap();
            }
            // Capture unknown attributes for roundtrip fidelity (feature-gated)
            writeln!(code, "                #[cfg(feature = \"extra-attrs\")]").unwrap();
            writeln!(code, "                unknown => {{").unwrap();
            writeln!(
                code,
                "                    let key = String::from_utf8_lossy(unknown).into_owned();"
            )
            .unwrap();
            writeln!(
                code,
                "                    extra_attrs.insert(key, val.into_owned());"
            )
            .unwrap();
            writeln!(code, "                }}").unwrap();
            writeln!(
                code,
                "                #[cfg(not(feature = \"extra-attrs\"))]"
            )
            .unwrap();
            writeln!(code, "                _ => {{}}").unwrap();
            writeln!(code, "            }}").unwrap();
            writeln!(code, "        }}").unwrap();
        }

        // Parse child elements and text content (only if not empty element)
        if has_parsing_loop {
            writeln!(code).unwrap();
            writeln!(code, "        // Parse child elements").unwrap();
            writeln!(code, "        if !is_empty {{").unwrap();
            writeln!(code, "            let mut buf = Vec::new();").unwrap();
            writeln!(code, "            loop {{").unwrap();
            writeln!(
                code,
                "                match reader.read_event_into(&mut buf)? {{"
            )
            .unwrap();
            writeln!(code, "                    Event::Start(e) => {{").unwrap();
            writeln!(code, "                        match e.name().as_ref() {{").unwrap();

            for field in &elem_fields {
                let base_name = field.name.strip_prefix("r#").unwrap_or(&field.name);
                let base_name = base_name.trim_start_matches('_');
                let var_name = format!("f_{}", base_name);
                let parse_expr = self.gen_element_parse_code(field, false);

                // Add cfg attribute for feature-gated fields
                if let Some(ref feature) = self.get_field_feature(&rust_name, &field.xml_name) {
                    writeln!(
                        code,
                        "                            #[cfg(feature = \"{}\")]",
                        feature
                    )
                    .unwrap();
                }
                writeln!(
                    code,
                    "                            b\"{}\" => {{",
                    field.xml_name
                )
                .unwrap();
                if field.is_vec {
                    writeln!(
                        code,
                        "                                {}.push({});",
                        var_name, parse_expr
                    )
                    .unwrap();
                } else {
                    writeln!(
                        code,
                        "                                {} = Some({});",
                        var_name, parse_expr
                    )
                    .unwrap();
                }
                writeln!(code, "                            }}").unwrap();
            }

            // Capture or skip unknown elements (feature-gated)
            writeln!(
                code,
                "                            #[cfg(feature = \"extra-children\")]"
            )
            .unwrap();
            writeln!(code, "                            _ => {{").unwrap();
            writeln!(
                code,
                "                                // Capture unknown element for roundtrip"
            )
            .unwrap();
            writeln!(code, "                                let elem = RawXmlElement::from_reader(reader, &e)?;").unwrap();
            writeln!(
                code,
                "                                extra_children.push(RawXmlNode::Element(elem));"
            )
            .unwrap();
            writeln!(code, "                            }}").unwrap();
            writeln!(
                code,
                "                            #[cfg(not(feature = \"extra-children\"))]"
            )
            .unwrap();
            writeln!(code, "                            _ => {{").unwrap();
            writeln!(
                code,
                "                                // Skip unknown element"
            )
            .unwrap();
            writeln!(
                code,
                "                                skip_element(reader)?;"
            )
            .unwrap();
            writeln!(code, "                            }}").unwrap();
            writeln!(code, "                        }}").unwrap();
            writeln!(code, "                    }}").unwrap();
            writeln!(code, "                    Event::Empty(e) => {{").unwrap();
            writeln!(code, "                        match e.name().as_ref() {{").unwrap();

            for field in &elem_fields {
                let base_name = field.name.strip_prefix("r#").unwrap_or(&field.name);
                let base_name = base_name.trim_start_matches('_');
                let var_name = format!("f_{}", base_name);
                let parse_expr = self.gen_element_parse_code(field, true);

                // Add cfg attribute for feature-gated fields
                if let Some(ref feature) = self.get_field_feature(&rust_name, &field.xml_name) {
                    writeln!(
                        code,
                        "                            #[cfg(feature = \"{}\")]",
                        feature
                    )
                    .unwrap();
                }
                writeln!(
                    code,
                    "                            b\"{}\" => {{",
                    field.xml_name
                )
                .unwrap();
                if field.is_vec {
                    writeln!(
                        code,
                        "                                {}.push({});",
                        var_name, parse_expr
                    )
                    .unwrap();
                } else {
                    writeln!(
                        code,
                        "                                {} = Some({});",
                        var_name, parse_expr
                    )
                    .unwrap();
                }
                writeln!(code, "                            }}").unwrap();
            }

            // Capture or skip unknown empty elements (feature-gated)
            writeln!(
                code,
                "                            #[cfg(feature = \"extra-children\")]"
            )
            .unwrap();
            writeln!(code, "                            _ => {{").unwrap();
            writeln!(
                code,
                "                                // Capture unknown empty element for roundtrip"
            )
            .unwrap();
            writeln!(
                code,
                "                                let elem = RawXmlElement::from_empty(&e);"
            )
            .unwrap();
            writeln!(
                code,
                "                                extra_children.push(RawXmlNode::Element(elem));"
            )
            .unwrap();
            writeln!(code, "                            }}").unwrap();
            writeln!(
                code,
                "                            #[cfg(not(feature = \"extra-children\"))]"
            )
            .unwrap();
            writeln!(code, "                            _ => {{}}").unwrap();
            writeln!(code, "                        }}").unwrap();
            writeln!(code, "                    }}").unwrap();

            // Handle text content if any text fields
            if !text_fields.is_empty() {
                writeln!(code, "                    Event::Text(e) => {{").unwrap();
                for field in &text_fields {
                    let base_name = field.name.strip_prefix("r#").unwrap_or(&field.name);
                    let base_name = base_name.trim_start_matches('_');
                    let var_name = format!("f_{}", base_name);
                    writeln!(
                        code,
                        "                        {} = Some(e.decode().unwrap_or_default().into_owned());",
                        var_name
                    ).unwrap();
                }
                writeln!(code, "                    }}").unwrap();
            }

            writeln!(code, "                    Event::End(_) => break,").unwrap();
            writeln!(code, "                    Event::Eof => break,").unwrap();
            writeln!(code, "                    _ => {{}}").unwrap();
            writeln!(code, "                }}").unwrap();
            writeln!(code, "                buf.clear();").unwrap();
            writeln!(code, "            }}").unwrap();
            writeln!(code, "        }}").unwrap();
        } else {
            // No child elements, but still need to read to end tag if not empty
            writeln!(code).unwrap();
            writeln!(code, "        if !is_empty {{").unwrap();
            writeln!(code, "            let mut buf = Vec::new();").unwrap();
            writeln!(code, "            loop {{").unwrap();
            writeln!(
                code,
                "                match reader.read_event_into(&mut buf)? {{"
            )
            .unwrap();
            writeln!(code, "                    Event::End(_) => break,").unwrap();
            writeln!(code, "                    Event::Eof => break,").unwrap();
            writeln!(code, "                    _ => {{}}").unwrap();
            writeln!(code, "                }}").unwrap();
            writeln!(code, "                buf.clear();").unwrap();
            writeln!(code, "            }}").unwrap();
            writeln!(code, "        }}").unwrap();
        }

        // Build result struct - use original field names
        writeln!(code).unwrap();
        writeln!(code, "        Ok(Self {{").unwrap();
        for field in &fields {
            let base_name = field.name.strip_prefix("r#").unwrap_or(&field.name);
            let base_name = base_name.trim_start_matches('_');
            let var_name = format!("f_{}", base_name);

            // Add cfg attribute for feature-gated fields
            if let Some(ref feature) = self.get_field_feature(&rust_name, &field.xml_name) {
                writeln!(code, "            #[cfg(feature = \"{}\")]", feature).unwrap();
            }

            if field.is_optional || field.is_vec {
                writeln!(code, "            {}: {},", field.name, var_name).unwrap();
            } else {
                // Required field - unwrap with error
                writeln!(
                    code,
                    "            {}: {}.ok_or_else(|| ParseError::MissingAttribute(\"{}\".to_string()))?,",
                    field.name, var_name, field.xml_name
                )
                .unwrap();
            }
        }
        // Add extra_attrs if this struct has attributes (feature-gated)
        if has_attrs {
            writeln!(code, "            #[cfg(feature = \"extra-attrs\")]").unwrap();
            writeln!(code, "            extra_attrs,").unwrap();
        }
        // Add extra_children if this struct has a parsing loop (feature-gated)
        if has_parsing_loop {
            writeln!(code, "            #[cfg(feature = \"extra-children\")]").unwrap();
            writeln!(code, "            extra_children,").unwrap();
        }
        writeln!(code, "        }})").unwrap();
        writeln!(code, "    }}").unwrap();
        writeln!(code, "}}").unwrap();

        Some(code)
    }

    fn gen_attr_parse_expr(&self, pattern: &Pattern) -> String {
        match pattern {
            Pattern::Datatype { library, name, .. } if library == "xsd" => match name.as_str() {
                "boolean" => "Some(val == \"true\" || val == \"1\")".to_string(),
                "integer" | "int" | "long" | "short" | "byte" => "val.parse().ok()".to_string(),
                "unsignedInt" | "unsignedLong" | "unsignedShort" | "unsignedByte" => {
                    "val.parse().ok()".to_string()
                }
                "double" | "float" | "decimal" => "val.parse().ok()".to_string(),
                "hexBinary" | "base64Binary" => "decode_hex(&val)".to_string(),
                _ => "Some(val.into_owned())".to_string(),
            },
            Pattern::Ref(name) => {
                // First recurse to find what this ref resolves to
                if let Some(def_pattern) = self.definitions.get(name.as_str()) {
                    return self.gen_attr_parse_expr(def_pattern);
                }
                // Unknown ref - if it looks like an enum (ST_), use FromStr, otherwise string
                if name.contains("_ST_") {
                    "val.parse().ok()".to_string()
                } else {
                    "Some(val.into_owned())".to_string()
                }
            }
            // String choice enums
            Pattern::Choice(variants)
                if variants
                    .iter()
                    .all(|v| matches!(v, Pattern::StringLiteral(_))) =>
            {
                "val.parse().ok()".to_string()
            }
            _ => "Some(val.into_owned())".to_string(),
        }
    }

    fn extract_fields(&self, pattern: &Pattern) -> Vec<Field> {
        let mut fields = Vec::new();
        self.collect_fields(pattern, &mut fields, false);
        let mut seen = std::collections::HashSet::new();
        fields.retain(|f| seen.insert(f.name.clone()));
        fields
    }

    fn collect_fields(&self, pattern: &Pattern, fields: &mut Vec<Field>, is_optional: bool) {
        match pattern {
            Pattern::Attribute { name, pattern } => {
                fields.push(Field {
                    name: self.qname_to_field_name(name),
                    xml_name: name.local.clone(),
                    pattern: pattern.as_ref().clone(),
                    is_optional,
                    is_attribute: true,
                    is_vec: false,
                    is_text_content: false,
                });
            }
            Pattern::Element { name, pattern } => {
                fields.push(Field {
                    name: self.qname_to_field_name(name),
                    xml_name: name.local.clone(),
                    pattern: pattern.as_ref().clone(),
                    is_optional,
                    is_attribute: false,
                    is_vec: false,
                    is_text_content: false,
                });
            }
            Pattern::Sequence(items) | Pattern::Interleave(items) => {
                for item in items {
                    self.collect_fields(item, fields, is_optional);
                }
            }
            Pattern::Optional(inner) => {
                self.collect_fields(inner, fields, true);
            }
            Pattern::ZeroOrMore(inner) | Pattern::OneOrMore(inner) => match inner.as_ref() {
                Pattern::Element { name, pattern } => {
                    fields.push(Field {
                        name: self.qname_to_field_name(name),
                        xml_name: name.local.clone(),
                        pattern: pattern.as_ref().clone(),
                        is_optional: false,
                        is_attribute: false,
                        is_vec: true,
                        is_text_content: false,
                    });
                }
                Pattern::Ref(name) if name.contains("_EG_") => {
                    let _ = name;
                }
                Pattern::Choice(_) | Pattern::Ref(_) => {
                    self.collect_fields(inner, fields, false);
                }
                _ => {}
            },
            Pattern::Group(inner) => {
                self.collect_fields(inner, fields, is_optional);
            }
            Pattern::Ref(name) => {
                // Check if this is a reference to a simple type (text content)
                if let Some(pattern) = self.definitions.get(name.as_str())
                    && self.is_string_type(pattern)
                {
                    // This is text content - add as a "text" field
                    fields.push(Field {
                        name: "text".to_string(),
                        xml_name: "$text".to_string(),
                        pattern: Pattern::Datatype {
                            library: "xsd".to_string(),
                            name: "string".to_string(),
                            params: vec![],
                        },
                        is_optional,
                        is_attribute: false,
                        is_vec: false,
                        is_text_content: true,
                    });
                }
            }
            Pattern::Choice(_) => {}
            _ => {}
        }
    }

    /// Check if a pattern resolves to a string type (for text content detection).
    fn is_string_type(&self, pattern: &Pattern) -> bool {
        match pattern {
            Pattern::Datatype { library, name, .. } => {
                library == "xsd" && (name == "string" || name == "token" || name == "NCName")
            }
            Pattern::Ref(name) => {
                // Check if the referenced type is a string type
                self.definitions
                    .get(name.as_str())
                    .is_some_and(|p| self.is_string_type(p))
            }
            _ => false,
        }
    }

    fn to_rust_type_name(&self, name: &str) -> String {
        let spec_name = strip_namespace_prefix(name);
        if let Some(mappings) = &self.config.name_mappings
            && let Some(mapped) = mappings.resolve_type(&self.config.module_name, spec_name)
        {
            return mapped.to_string();
        }
        to_pascal_case(spec_name)
    }

    fn to_rust_variant_name(&self, name: &str) -> String {
        if name.is_empty() {
            return "Empty".to_string();
        }
        if let Some(mappings) = &self.config.name_mappings
            && let Some(mapped) = mappings.resolve_variant(&self.config.module_name, name)
        {
            return mapped.to_string();
        }
        let name = to_pascal_case(name);
        if name.chars().next().is_some_and(|c| c.is_ascii_digit()) {
            format!("_{}", name)
        } else {
            name
        }
    }

    fn qname_to_field_name(&self, qname: &QName) -> String {
        if let Some(mappings) = &self.config.name_mappings
            && let Some(mapped) = mappings.resolve_field(&self.config.module_name, &qname.local)
        {
            return mapped.to_string();
        }
        to_snake_case(&qname.local)
    }

    /// Get the feature name for a field if it requires feature gating.
    /// Returns None if the field is "core" (always included) or unmapped.
    fn get_field_feature(&self, struct_name: &str, xml_field_name: &str) -> Option<String> {
        self.config
            .feature_mappings
            .as_ref()
            .and_then(|fm| {
                fm.primary_feature(&self.config.module_name, struct_name, xml_field_name)
            })
            .map(|feature| format!("{}-{}", self.config.module_name, feature))
    }

    fn pattern_to_rust_type(&self, pattern: &Pattern) -> (String, bool) {
        match pattern {
            Pattern::Ref(name) => {
                if self.definitions.contains_key(name.as_str()) {
                    let type_name = self.to_rust_type_name(name);
                    let needs_box = name.contains("_CT_") || name.contains("_EG_");
                    (type_name, needs_box)
                } else {
                    ("String".to_string(), false)
                }
            }
            Pattern::Datatype { library, name, .. } => {
                (xsd_to_rust(library, name).to_string(), false)
            }
            Pattern::Empty => ("()".to_string(), false),
            Pattern::StringLiteral(_) => ("String".to_string(), false),
            Pattern::Choice(_) => ("String".to_string(), false),
            _ => ("String".to_string(), false),
        }
    }

    /// Generate code to parse a child element field.
    /// Returns the expression that produces the parsed value.
    fn gen_element_parse_code(&self, field: &Field, is_empty_element: bool) -> String {
        let (rust_type, needs_box) = self.pattern_to_rust_type(&field.pattern);
        let strategy = self.get_parse_strategy(&field.pattern);

        let parse_expr = match strategy {
            ParseStrategy::FromXml => {
                // For type aliases, we need to find the actual type that has FromXml impl
                // and check if the type alias already includes a Box
                let (actual_type, type_alias_has_box) =
                    self.resolve_from_xml_type_with_box(&field.pattern);
                let final_type = actual_type.unwrap_or(rust_type.clone());

                let mut expr = format!(
                    "{}::from_xml(reader, &e, {})?",
                    final_type, is_empty_element
                );

                // If the type alias is like `type X = Box<Y>`, we need a Box for that
                if type_alias_has_box {
                    expr = format!("Box::new({})", expr);
                }

                // If the field type is `Box<X>`, we need another Box for the field
                if needs_box {
                    expr = format!("Box::new({})", expr);
                }

                return expr;
            }
            ParseStrategy::TextFromStr => {
                if is_empty_element {
                    // Empty element with FromStr - try to parse empty string or use default
                    "Default::default()".to_string()
                } else {
                    "{ let text = read_text_content(reader)?; text.parse().map_err(|_| ParseError::InvalidValue(text))? }".to_string()
                }
            }
            ParseStrategy::TextString => {
                if is_empty_element {
                    "String::new()".to_string()
                } else {
                    "read_text_content(reader)?".to_string()
                }
            }
            ParseStrategy::TextHexBinary => {
                if is_empty_element {
                    "Vec::new()".to_string()
                } else {
                    "{ let text = read_text_content(reader)?; decode_hex(&text).unwrap_or_default() }".to_string()
                }
            }
        };

        // For non-FromXml strategies, handle boxing normally
        if strategy != ParseStrategy::FromXml && needs_box {
            format!("Box::new({})", parse_expr)
        } else {
            parse_expr
        }
    }

    /// Resolve a pattern to the actual type that has a FromXml impl.
    /// Returns (Option<type_name>, type_alias_already_has_box).
    /// For type aliases (definitions that are just Pattern::Element wrapping a ref),
    /// returns the underlying type's Rust name and indicates the alias includes a Box.
    fn resolve_from_xml_type_with_box(&self, pattern: &Pattern) -> (Option<String>, bool) {
        match pattern {
            Pattern::Ref(name) => {
                // Look up the definition
                if let Some(def_pattern) = self.definitions.get(name.as_str()) {
                    // If it's a type alias (Pattern::Element wrapping another ref),
                    // the generated type will be `type X = Box<InnerType>`
                    // So we resolve to the inner type and signal that boxing is already done
                    if let Pattern::Element { pattern: inner, .. } = def_pattern
                        && let Some(inner_type) = self.resolve_from_xml_type_simple(inner)
                    {
                        return (Some(inner_type), true);
                    }
                    // If it's a simple ref, just a type alias like `type X = Y`
                    // Check if the inner ref is external (not in our definitions)
                    if let Pattern::Ref(inner_name) = def_pattern
                        && !self.definitions.contains_key(inner_name.as_str())
                    {
                        // External ref like r_id - don't try to resolve, use original type
                        return (None, false);
                    } else if let Pattern::Ref(inner_name) = def_pattern {
                        return self
                            .resolve_from_xml_type_with_box(&Pattern::Ref(inner_name.clone()));
                    }
                }
                // Type in definitions - use its name
                (Some(self.to_rust_type_name(name)), false)
            }
            _ => (None, false),
        }
    }

    /// Simple resolver that just returns the type name without box tracking.
    fn resolve_from_xml_type_simple(&self, pattern: &Pattern) -> Option<String> {
        match pattern {
            Pattern::Ref(name) => {
                if let Some(def_pattern) = self.definitions.get(name.as_str()) {
                    if let Pattern::Element { pattern: inner, .. } = def_pattern {
                        return self.resolve_from_xml_type_simple(inner);
                    }
                    if let Pattern::Ref(inner_name) = def_pattern {
                        // Check if the inner ref is external
                        if !self.definitions.contains_key(inner_name.as_str()) {
                            // External ref - stop resolution
                            return None;
                        }
                        return self
                            .resolve_from_xml_type_simple(&Pattern::Ref(inner_name.clone()));
                    }
                }
                Some(self.to_rust_type_name(name))
            }
            _ => None,
        }
    }

    /// Determine how to parse a field's value from XML.
    fn get_parse_strategy(&self, pattern: &Pattern) -> ParseStrategy {
        match pattern {
            Pattern::Ref(name) => {
                // Check what this ref resolves to
                if let Some(def_pattern) = self.definitions.get(name.as_str()) {
                    // If it's a ref to another ref (like CT_Drawing = r_id),
                    // and that ref is external (not in our definitions), treat as empty struct
                    if let Pattern::Ref(inner_name) = def_pattern
                        && !self.definitions.contains_key(inner_name.as_str())
                    {
                        // External ref (e.g., r_id from another schema)
                        // These generate empty structs that need simple FromXml impls
                        return ParseStrategy::FromXml;
                    }
                    return self.get_parse_strategy(def_pattern);
                }

                // Complex types (CT_*) and element groups (EG_*) need from_xml
                if name.contains("_CT_") || name.contains("_EG_") {
                    return ParseStrategy::FromXml;
                }

                // Unknown ref - treat as string
                ParseStrategy::TextString
            }
            Pattern::Datatype { library, name, .. } => {
                if library == "xsd" {
                    match name.as_str() {
                        "string" | "token" | "NCName" | "ID" | "IDREF" | "anyURI" | "dateTime"
                        | "date" | "time" => ParseStrategy::TextString,
                        "hexBinary" => ParseStrategy::TextHexBinary,
                        "base64Binary" => ParseStrategy::TextHexBinary, // TODO: proper base64
                        // Numbers and booleans use FromStr
                        _ => ParseStrategy::TextFromStr,
                    }
                } else {
                    ParseStrategy::TextString
                }
            }
            // String enums use FromStr
            Pattern::Choice(variants)
                if variants
                    .iter()
                    .all(|v| matches!(v, Pattern::StringLiteral(_))) =>
            {
                ParseStrategy::TextFromStr
            }
            Pattern::StringLiteral(_) => ParseStrategy::TextString,
            Pattern::Empty => ParseStrategy::TextString,
            Pattern::List(_) => ParseStrategy::TextString,
            // For complex patterns (Sequence, etc.), assume it needs from_xml
            _ => ParseStrategy::FromXml,
        }
    }
}

struct Field {
    name: String,
    xml_name: String,
    pattern: Pattern,
    is_optional: bool,
    is_attribute: bool,
    is_vec: bool,
    is_text_content: bool,
}

fn strip_namespace_prefix(name: &str) -> &str {
    for kind in ["CT_", "ST_", "EG_"] {
        if let Some(pos) = name.find(kind)
            && pos > 0
        {
            return &name[pos..];
        }
    }
    name
}

fn to_pascal_case(s: &str) -> String {
    let mut result = String::new();
    let mut capitalize_next = true;
    for ch in s.chars() {
        if ch == '_' || ch == '-' {
            capitalize_next = true;
        } else if capitalize_next {
            result.extend(ch.to_uppercase());
            capitalize_next = false;
        } else {
            result.push(ch);
        }
    }
    result
}

fn to_snake_case(s: &str) -> String {
    let mut result = String::new();
    for (i, ch) in s.chars().enumerate() {
        if ch.is_uppercase() && i > 0 {
            result.push('_');
        }
        result.extend(ch.to_lowercase());
    }
    match result.as_str() {
        // Rust keywords
        "type" => "r#type".to_string(),
        "ref" => "r#ref".to_string(),
        "match" => "r#match".to_string(),
        "in" => "r#in".to_string(),
        "for" => "r#for".to_string(),
        "macro" => "r#macro".to_string(),
        "if" => "r#if".to_string(),
        "else" => "r#else".to_string(),
        "loop" => "r#loop".to_string(),
        "break" => "r#break".to_string(),
        "continue" => "r#continue".to_string(),
        "return" => "r#return".to_string(),
        "self" => "r#self".to_string(),
        "super" => "r#super".to_string(),
        "crate" => "r#crate".to_string(),
        "mod" => "r#mod".to_string(),
        "pub" => "r#pub".to_string(),
        "use" => "r#use".to_string(),
        "as" => "r#as".to_string(),
        "static" => "r#static".to_string(),
        "const" => "r#const".to_string(),
        "extern" => "r#extern".to_string(),
        "fn" => "r#fn".to_string(),
        "struct" => "r#struct".to_string(),
        "enum" => "r#enum".to_string(),
        "trait" => "r#trait".to_string(),
        "impl" => "r#impl".to_string(),
        "where" => "r#where".to_string(),
        "async" => "r#async".to_string(),
        "await" => "r#await".to_string(),
        "move" => "r#move".to_string(),
        "box" => "r#box".to_string(),
        "dyn" => "r#dyn".to_string(),
        "abstract" => "r#abstract".to_string(),
        "become" => "r#become".to_string(),
        "do" => "r#do".to_string(),
        "final" => "r#final".to_string(),
        "override" => "r#override".to_string(),
        "priv" => "r#priv".to_string(),
        "typeof" => "r#typeof".to_string(),
        "unsized" => "r#unsized".to_string(),
        "virtual" => "r#virtual".to_string(),
        "yield" => "r#yield".to_string(),
        "try" => "r#try".to_string(),
        _ => result,
    }
}

fn xsd_to_rust(library: &str, name: &str) -> &'static str {
    if library == "xsd" {
        match name {
            "string" => "String",
            "integer" => "i64",
            "int" => "i32",
            "long" => "i64",
            "short" => "i16",
            "byte" => "i8",
            "unsignedInt" => "u32",
            "unsignedLong" => "u64",
            "unsignedShort" => "u16",
            "unsignedByte" => "u8",
            "boolean" => "bool",
            "double" => "f64",
            "float" => "f32",
            "decimal" => "f64",
            _ => "String",
        }
    } else {
        "String"
    }
}
