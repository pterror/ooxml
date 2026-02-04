//! XML serializer code generator.
//!
//! This module generates `ToXml` trait implementations for OOXML types,
//! enabling roundtrip XML serialization to complement the `FromXml` parsers.

use crate::ast::{Pattern, QName, Schema};
use crate::codegen::CodegenConfig;
use crate::parser_gen::{
    Field, strip_namespace_prefix, to_pascal_case, to_snake_case, xsd_to_rust,
};
use std::collections::HashMap;
use std::fmt::Write;

/// Generate serializer code for all types in the schema.
pub fn generate_serializers(schema: &Schema, config: &CodegenConfig) -> String {
    let mut g = SerializerGenerator::new(schema, config);
    g.run()
}

/// How to write a field value to XML.
#[derive(Debug, Clone, Copy, PartialEq)]
enum WriteStrategy {
    /// Complex type — call write_element()
    ToXml,
    /// Enum or number — use .to_string()
    DisplayStr,
    /// String type — use value directly
    AsString,
    /// Vec<u8> — hex-encode
    HexBinary,
    /// bool — write as "1"/"0"
    OoxmlBool,
}

struct SerializerGenerator<'a> {
    schema: &'a Schema,
    config: &'a CodegenConfig,
    output: String,
    definitions: HashMap<&'a str, &'a Pattern>,
}

impl<'a> SerializerGenerator<'a> {
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

    /// Get the XML namespace prefix for serialization.
    /// Returns None for unprefixed (SML), Some("w") for WML, etc.
    fn xml_prefix(&self) -> Option<&str> {
        self.config.xml_serialize_prefix.as_deref()
    }

    fn run(&mut self) -> String {
        self.write_header();

        for def in &self.schema.definitions {
            if !def.name.contains("_ST_") && !self.is_simple_type(&def.pattern) {
                if def.name.contains("_EG_") && self.is_element_choice(&def.pattern) {
                    if let Some(code) = self.gen_element_group_serializer(def) {
                        self.output.push_str(&code);
                        self.output.push('\n');
                    }
                } else if !self.is_type_alias(&def.pattern)
                    && let Some(code) = self.gen_struct_serializer(def)
                {
                    self.output.push_str(&code);
                    self.output.push('\n');
                }
            }
        }

        std::mem::take(&mut self.output)
    }

    fn write_header(&mut self) {
        writeln!(self.output, "// ToXml serializers for generated types.").unwrap();
        writeln!(
            self.output,
            "// Enables roundtrip XML serialization alongside FromXml parsers."
        )
        .unwrap();
        writeln!(self.output).unwrap();
        writeln!(
            self.output,
            "#![allow(unused_variables, unused_assignments, unreachable_code)]"
        )
        .unwrap();
        writeln!(self.output, "#![allow(clippy::single_match)]").unwrap();
        writeln!(self.output, "#![allow(clippy::match_single_binding)]").unwrap();
        writeln!(self.output, "#![allow(clippy::explicit_counter_loop)]").unwrap();
        writeln!(self.output).unwrap();
        writeln!(self.output, "use super::generated::*;").unwrap();
        writeln!(self.output, "use quick_xml::Writer;").unwrap();
        writeln!(
            self.output,
            "use quick_xml::events::{{BytesEnd, BytesStart, BytesText, Event}};"
        )
        .unwrap();
        writeln!(self.output, "use std::io::Write;").unwrap();
        writeln!(self.output).unwrap();

        // SerializeError
        writeln!(self.output, "/// Error type for XML serialization.").unwrap();
        writeln!(self.output, "#[derive(Debug)]").unwrap();
        writeln!(self.output, "pub enum SerializeError {{").unwrap();
        writeln!(self.output, "    Xml(quick_xml::Error),").unwrap();
        writeln!(self.output, "    Io(std::io::Error),").unwrap();
        writeln!(self.output, "    #[cfg(feature = \"extra-children\")]").unwrap();
        writeln!(self.output, "    RawXml(ooxml_xml::Error),").unwrap();
        writeln!(self.output, "}}").unwrap();
        writeln!(self.output).unwrap();
        writeln!(
            self.output,
            "impl From<quick_xml::Error> for SerializeError {{"
        )
        .unwrap();
        writeln!(self.output, "    fn from(e: quick_xml::Error) -> Self {{").unwrap();
        writeln!(self.output, "        SerializeError::Xml(e)").unwrap();
        writeln!(self.output, "    }}").unwrap();
        writeln!(self.output, "}}").unwrap();
        writeln!(self.output).unwrap();
        writeln!(
            self.output,
            "impl From<std::io::Error> for SerializeError {{"
        )
        .unwrap();
        writeln!(self.output, "    fn from(e: std::io::Error) -> Self {{").unwrap();
        writeln!(self.output, "        SerializeError::Io(e)").unwrap();
        writeln!(self.output, "    }}").unwrap();
        writeln!(self.output, "}}").unwrap();
        writeln!(self.output).unwrap();
        writeln!(self.output, "#[cfg(feature = \"extra-children\")]").unwrap();
        writeln!(
            self.output,
            "impl From<ooxml_xml::Error> for SerializeError {{"
        )
        .unwrap();
        writeln!(self.output, "    fn from(e: ooxml_xml::Error) -> Self {{").unwrap();
        writeln!(self.output, "        SerializeError::RawXml(e)").unwrap();
        writeln!(self.output, "    }}").unwrap();
        writeln!(self.output, "}}").unwrap();
        writeln!(self.output).unwrap();
        writeln!(self.output, "impl std::fmt::Display for SerializeError {{").unwrap();
        writeln!(
            self.output,
            "    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {{"
        )
        .unwrap();
        writeln!(self.output, "        match self {{").unwrap();
        writeln!(
            self.output,
            "            Self::Xml(e) => write!(f, \"XML error: {{}}\", e),"
        )
        .unwrap();
        writeln!(
            self.output,
            "            Self::Io(e) => write!(f, \"IO error: {{}}\", e),"
        )
        .unwrap();
        writeln!(
            self.output,
            "            #[cfg(feature = \"extra-children\")]"
        )
        .unwrap();
        writeln!(
            self.output,
            "            Self::RawXml(e) => write!(f, \"RawXml error: {{}}\", e),"
        )
        .unwrap();
        writeln!(self.output, "        }}").unwrap();
        writeln!(self.output, "    }}").unwrap();
        writeln!(self.output, "}}").unwrap();
        writeln!(self.output).unwrap();
        writeln!(
            self.output,
            "impl std::error::Error for SerializeError {{}}"
        )
        .unwrap();
        writeln!(self.output).unwrap();

        // ToXml trait
        writeln!(
            self.output,
            "/// Trait for types that can be serialized to XML events."
        )
        .unwrap();
        writeln!(self.output, "pub trait ToXml {{").unwrap();
        writeln!(
            self.output,
            "    /// Write attributes onto the start tag and return it."
        )
        .unwrap();
        writeln!(
            self.output,
            "    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {{ start }}"
        )
        .unwrap();
        writeln!(self.output).unwrap();
        writeln!(
            self.output,
            "    /// Write child elements and text content inside the element."
        )
        .unwrap();
        writeln!(
            self.output,
            "    fn write_children<W: Write>(&self, writer: &mut Writer<W>) -> Result<(), SerializeError> {{ Ok(()) }}"
        )
        .unwrap();
        writeln!(self.output).unwrap();
        writeln!(
            self.output,
            "    /// Whether this element has no children (self-closing)."
        )
        .unwrap();
        writeln!(
            self.output,
            "    fn is_empty_element(&self) -> bool {{ false }}"
        )
        .unwrap();
        writeln!(self.output).unwrap();
        writeln!(
            self.output,
            "    /// Write a complete element: `<tag attrs>children</tag>` or `<tag attrs/>`."
        )
        .unwrap();
        writeln!(
            self.output,
            "    fn write_element<W: Write>(&self, tag: &str, writer: &mut Writer<W>) -> Result<(), SerializeError> {{"
        )
        .unwrap();
        writeln!(self.output, "        let start = BytesStart::new(tag);").unwrap();
        writeln!(self.output, "        let start = self.write_attrs(start);").unwrap();
        writeln!(self.output, "        if self.is_empty_element() {{").unwrap();
        writeln!(
            self.output,
            "            writer.write_event(Event::Empty(start))?;"
        )
        .unwrap();
        writeln!(self.output, "        }} else {{").unwrap();
        writeln!(
            self.output,
            "            writer.write_event(Event::Start(start))?;"
        )
        .unwrap();
        writeln!(self.output, "            self.write_children(writer)?;").unwrap();
        writeln!(
            self.output,
            "            writer.write_event(Event::End(BytesEnd::new(tag)))?;"
        )
        .unwrap();
        writeln!(self.output, "        }}").unwrap();
        writeln!(self.output, "        Ok(())").unwrap();
        writeln!(self.output, "    }}").unwrap();
        writeln!(self.output, "}}").unwrap();
        writeln!(self.output).unwrap();

        // Box<T> blanket impl
        writeln!(self.output, "impl<T: ToXml> ToXml for Box<T> {{").unwrap();
        writeln!(
            self.output,
            "    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {{"
        )
        .unwrap();
        writeln!(self.output, "        (**self).write_attrs(start)").unwrap();
        writeln!(self.output, "    }}").unwrap();
        writeln!(
            self.output,
            "    fn write_children<W: Write>(&self, writer: &mut Writer<W>) -> Result<(), SerializeError> {{"
        )
        .unwrap();
        writeln!(self.output, "        (**self).write_children(writer)").unwrap();
        writeln!(self.output, "    }}").unwrap();
        writeln!(self.output, "    fn is_empty_element(&self) -> bool {{").unwrap();
        writeln!(self.output, "        (**self).is_empty_element()").unwrap();
        writeln!(self.output, "    }}").unwrap();
        writeln!(
            self.output,
            "    fn write_element<W: Write>(&self, tag: &str, writer: &mut Writer<W>) -> Result<(), SerializeError> {{"
        )
        .unwrap();
        writeln!(self.output, "        (**self).write_element(tag, writer)").unwrap();
        writeln!(self.output, "    }}").unwrap();
        writeln!(self.output, "}}").unwrap();
        writeln!(self.output).unwrap();

        // encode_hex helper
        writeln!(self.output, "#[allow(dead_code)]").unwrap();
        writeln!(self.output, "/// Encode bytes as a hex string.").unwrap();
        writeln!(self.output, "fn encode_hex(bytes: &[u8]) -> String {{").unwrap();
        writeln!(
            self.output,
            "    bytes.iter().map(|b| format!(\"{{:02X}}\", b)).collect()"
        )
        .unwrap();
        writeln!(self.output, "}}").unwrap();
        writeln!(self.output).unwrap();
    }

    // =========================================================================
    // Struct serializer generation
    // =========================================================================

    fn gen_struct_serializer(&self, def: &crate::ast::Definition) -> Option<String> {
        let rust_name = self.to_rust_type_name(&def.name);
        let fields = self.extract_fields(&def.pattern);

        let attr_fields: Vec<_> = fields.iter().filter(|f| f.is_attribute).collect();
        let child_fields: Vec<_> = fields
            .iter()
            .filter(|f| !f.is_attribute && !f.is_text_content)
            .collect();
        let text_fields: Vec<_> = fields.iter().filter(|f| f.is_text_content).collect();
        let has_attrs = !attr_fields.is_empty();
        let has_children = !child_fields.is_empty();
        let has_text = !text_fields.is_empty();
        let has_parsing_content = has_children || has_text;

        if fields.is_empty() {
            let has_unresolved_children = self.has_xml_children_pattern(&def.pattern);
            let has_unresolved_attrs = self.has_xml_attr_pattern(&def.pattern);

            if !has_unresolved_children && !has_unresolved_attrs {
                // Truly empty struct — trivial impl
                let mut code = String::new();
                writeln!(code, "impl ToXml for {} {{", rust_name).unwrap();
                writeln!(code, "    fn is_empty_element(&self) -> bool {{ true }}").unwrap();
                writeln!(code, "}}").unwrap();
                return Some(code);
            }

            // Struct with only extra_children/extra_attrs (unresolved fields)
            let mut code = String::new();
            writeln!(code, "impl ToXml for {} {{", rust_name).unwrap();

            if has_unresolved_attrs {
                writeln!(
                    code,
                    "    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {{"
                )
                .unwrap();
                writeln!(code, "        #[allow(unused_mut)]").unwrap();
                writeln!(code, "        let mut start = start;").unwrap();
                writeln!(code, "        #[cfg(feature = \"extra-attrs\")]").unwrap();
                writeln!(code, "        for (key, value) in &self.extra_attrs {{").unwrap();
                writeln!(
                    code,
                    "            start.push_attribute((key.as_str(), value.as_str()));"
                )
                .unwrap();
                writeln!(code, "        }}").unwrap();
                writeln!(code, "        start").unwrap();
                writeln!(code, "    }}").unwrap();
            }

            if has_unresolved_children {
                writeln!(code).unwrap();
                writeln!(
                    code,
                    "    fn write_children<W: Write>(&self, writer: &mut Writer<W>) -> Result<(), SerializeError> {{"
                )
                .unwrap();
                writeln!(code, "        #[cfg(feature = \"extra-children\")]").unwrap();
                writeln!(code, "        for child in &self.extra_children {{").unwrap();
                writeln!(
                    code,
                    "            child.node.write_to(writer).map_err(SerializeError::from)?;"
                )
                .unwrap();
                writeln!(code, "        }}").unwrap();
                writeln!(code, "        Ok(())").unwrap();
                writeln!(code, "    }}").unwrap();
            }

            writeln!(code).unwrap();
            writeln!(code, "    fn is_empty_element(&self) -> bool {{").unwrap();
            if has_unresolved_children {
                writeln!(code, "        #[cfg(feature = \"extra-children\")]").unwrap();
                writeln!(
                    code,
                    "        if !self.extra_children.is_empty() {{ return false; }}"
                )
                .unwrap();
            }
            writeln!(code, "        true").unwrap();
            writeln!(code, "    }}").unwrap();
            writeln!(code, "}}").unwrap();
            return Some(code);
        }

        let mut code = String::new();
        writeln!(code, "impl ToXml for {} {{", rust_name).unwrap();

        // write_attrs
        if has_attrs {
            writeln!(
                code,
                "    fn write_attrs<'a>(&self, start: BytesStart<'a>) -> BytesStart<'a> {{"
            )
            .unwrap();
            writeln!(code, "        let mut start = start;").unwrap();

            for field in &attr_fields {
                let attr_name = self.qualified_attr_name(field);
                let feature = self.get_field_feature(&rust_name, &field.xml_name);
                if let Some(ref feat) = feature {
                    writeln!(code, "        #[cfg(feature = \"{}\")]", feat).unwrap();
                }

                let strategy = self.get_write_strategy(&field.pattern);
                if field.is_optional {
                    writeln!(
                        code,
                        "        if let Some(ref val) = self.{} {{",
                        field.name
                    )
                    .unwrap();
                    self.write_attr_push(&mut code, &attr_name, strategy, "val", "            ");
                    writeln!(code, "        }}").unwrap();
                } else {
                    writeln!(code, "        {{").unwrap();
                    writeln!(code, "            let val = &self.{};", field.name).unwrap();
                    self.write_attr_push(&mut code, &attr_name, strategy, "val", "            ");
                    writeln!(code, "        }}").unwrap();
                }
            }

            // extra_attrs
            if has_attrs {
                writeln!(code, "        #[cfg(feature = \"extra-attrs\")]").unwrap();
                writeln!(code, "        for (key, value) in &self.extra_attrs {{").unwrap();
                writeln!(
                    code,
                    "            start.push_attribute((key.as_str(), value.as_str()));"
                )
                .unwrap();
                writeln!(code, "        }}").unwrap();
            }

            writeln!(code, "        start").unwrap();
            writeln!(code, "    }}").unwrap();
        }

        // write_children
        if has_parsing_content {
            writeln!(code).unwrap();
            writeln!(
                code,
                "    fn write_children<W: Write>(&self, writer: &mut Writer<W>) -> Result<(), SerializeError> {{"
            )
            .unwrap();

            // Text content fields
            for field in &text_fields {
                let feature = self.get_field_feature(&rust_name, &field.xml_name);
                if let Some(ref feat) = feature {
                    writeln!(code, "        #[cfg(feature = \"{}\")]", feat).unwrap();
                }
                writeln!(
                    code,
                    "        if let Some(ref text) = self.{} {{",
                    field.name
                )
                .unwrap();
                writeln!(
                    code,
                    "            writer.write_event(Event::Text(BytesText::new(text)))?;"
                )
                .unwrap();
                writeln!(code, "        }}").unwrap();
            }

            // Position-interleaved extra_children (ADR-004)
            if !child_fields.is_empty() {
                // Declare iterator and position counter for interleaving unknown children
                // among known children by their original parse position.
                writeln!(code, "        #[cfg(feature = \"extra-children\")]").unwrap();
                writeln!(
                    code,
                    "        let mut extra_iter = self.extra_children.iter().peekable();"
                )
                .unwrap();
                writeln!(code, "        #[cfg(feature = \"extra-children\")]").unwrap();
                writeln!(code, "        let mut emit_idx: usize = 0;").unwrap();
            }

            // Child element fields (in schema order)
            for field in &child_fields {
                let feature = self.get_field_feature(&rust_name, &field.xml_name);

                if field.is_vec {
                    // Vec field — flush and increment inside loop (per item)
                    if let Some(ref feat) = feature {
                        writeln!(code, "        #[cfg(feature = \"{}\")]", feat).unwrap();
                    }
                    if self.is_eg_content_field(field) {
                        writeln!(code, "        for item in &self.{} {{", field.name).unwrap();
                        writeln!(code, "            #[cfg(feature = \"extra-children\")]").unwrap();
                        writeln!(code, "            while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {{").unwrap();
                        writeln!(code, "                extra_iter.next().unwrap().node.write_to(writer).map_err(SerializeError::from)?;").unwrap();
                        writeln!(code, "            }}").unwrap();
                        writeln!(code, "            item.write_element(\"\", writer)?;").unwrap();
                        writeln!(code, "            #[cfg(feature = \"extra-children\")]").unwrap();
                        writeln!(code, "            {{ emit_idx += 1; }}").unwrap();
                        writeln!(code, "        }}").unwrap();
                    } else {
                        let tag = self.qualified_element_name(field);
                        let strategy = self.get_write_strategy(&field.pattern);
                        writeln!(code, "        for item in &self.{} {{", field.name).unwrap();
                        writeln!(code, "            #[cfg(feature = \"extra-children\")]").unwrap();
                        writeln!(code, "            while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {{").unwrap();
                        writeln!(code, "                extra_iter.next().unwrap().node.write_to(writer).map_err(SerializeError::from)?;").unwrap();
                        writeln!(code, "            }}").unwrap();
                        self.write_child_element(
                            &mut code,
                            &tag,
                            strategy,
                            "item",
                            false,
                            "            ",
                        );
                        writeln!(code, "            #[cfg(feature = \"extra-children\")]").unwrap();
                        writeln!(code, "            {{ emit_idx += 1; }}").unwrap();
                        writeln!(code, "        }}").unwrap();
                    }
                } else {
                    // Scalar field — flush before write, increment after (both outside feature gate)
                    writeln!(code, "        #[cfg(feature = \"extra-children\")]").unwrap();
                    writeln!(
                        code,
                        "        while extra_iter.peek().is_some_and(|e| e.position <= emit_idx) {{"
                    )
                    .unwrap();
                    writeln!(code, "            extra_iter.next().unwrap().node.write_to(writer).map_err(SerializeError::from)?;").unwrap();
                    writeln!(code, "        }}").unwrap();

                    if let Some(ref feat) = feature {
                        writeln!(code, "        #[cfg(feature = \"{}\")]", feat).unwrap();
                    }

                    if self.is_eg_content_field(field) {
                        if field.is_optional {
                            writeln!(
                                code,
                                "        if let Some(ref val) = self.{} {{",
                                field.name
                            )
                            .unwrap();
                            writeln!(code, "            val.write_element(\"\", writer)?;")
                                .unwrap();
                            writeln!(code, "        }}").unwrap();
                        } else {
                            writeln!(
                                code,
                                "        self.{}.write_element(\"\", writer)?;",
                                field.name
                            )
                            .unwrap();
                        }
                    } else {
                        let tag = self.qualified_element_name(field);
                        let strategy = self.get_write_strategy(&field.pattern);

                        if field.is_optional {
                            writeln!(
                                code,
                                "        if let Some(ref val) = self.{} {{",
                                field.name
                            )
                            .unwrap();
                            self.write_child_element(
                                &mut code,
                                &tag,
                                strategy,
                                "val",
                                false,
                                "            ",
                            );
                            writeln!(code, "        }}").unwrap();
                        } else {
                            writeln!(code, "        {{").unwrap();
                            writeln!(code, "            let val = &self.{};", field.name).unwrap();
                            self.write_child_element(
                                &mut code,
                                &tag,
                                strategy,
                                "val",
                                true,
                                "            ",
                            );
                            writeln!(code, "        }}").unwrap();
                        }
                    }

                    // Always increment position counter (outside feature gate)
                    writeln!(code, "        #[cfg(feature = \"extra-children\")]").unwrap();
                    writeln!(code, "        {{ emit_idx += 1; }}").unwrap();
                }
            }

            // Flush remaining extra_children at end
            if !child_fields.is_empty() {
                writeln!(code, "        #[cfg(feature = \"extra-children\")]").unwrap();
                writeln!(code, "        for extra in extra_iter {{").unwrap();
                writeln!(
                    code,
                    "            extra.node.write_to(writer).map_err(SerializeError::from)?;"
                )
                .unwrap();
                writeln!(code, "        }}").unwrap();
            } else {
                // No child fields — just emit all extras in order (no interleaving needed)
                writeln!(code, "        #[cfg(feature = \"extra-children\")]").unwrap();
                writeln!(code, "        for extra in &self.extra_children {{").unwrap();
                writeln!(
                    code,
                    "            extra.node.write_to(writer).map_err(SerializeError::from)?;"
                )
                .unwrap();
                writeln!(code, "        }}").unwrap();
            }

            writeln!(code, "        Ok(())").unwrap();
            writeln!(code, "    }}").unwrap();
        }

        // is_empty_element
        writeln!(code).unwrap();
        writeln!(code, "    fn is_empty_element(&self) -> bool {{").unwrap();

        if !has_parsing_content {
            // Only attribute fields — always empty element
            writeln!(code, "        true").unwrap();
        } else {
            // Check each child/text field.
            // Track whether we hit an unconditional `return false` (required field)
            // so we don't emit unreachable code after it.
            let mut has_unconditional_return = false;
            for field in &text_fields {
                if has_unconditional_return {
                    break;
                }
                let feature = self.get_field_feature(&rust_name, &field.xml_name);
                if let Some(ref feat) = feature {
                    writeln!(code, "        #[cfg(feature = \"{}\")]", feat).unwrap();
                }
                writeln!(
                    code,
                    "        if self.{}.is_some() {{ return false; }}",
                    field.name
                )
                .unwrap();
            }
            for field in &child_fields {
                if has_unconditional_return {
                    break;
                }
                let feature = self.get_field_feature(&rust_name, &field.xml_name);
                if let Some(ref feat) = feature {
                    writeln!(code, "        #[cfg(feature = \"{}\")]", feat).unwrap();
                }
                if field.is_vec {
                    writeln!(
                        code,
                        "        if !self.{}.is_empty() {{ return false; }}",
                        field.name
                    )
                    .unwrap();
                } else if field.is_optional {
                    writeln!(
                        code,
                        "        if self.{}.is_some() {{ return false; }}",
                        field.name
                    )
                    .unwrap();
                } else if feature.is_some() {
                    // Feature-gated required field — only unconditional when feature enabled
                    writeln!(code, "        return false;").unwrap();
                } else {
                    // Required field — never empty. Use `false` (not `return false;`)
                    // since this will be the last expression in the function.
                    writeln!(code, "        false").unwrap();
                    has_unconditional_return = true;
                }
            }
            if !has_unconditional_return {
                writeln!(code, "        #[cfg(feature = \"extra-children\")]").unwrap();
                writeln!(
                    code,
                    "        if !self.extra_children.is_empty() {{ return false; }}"
                )
                .unwrap();
                writeln!(code, "        true").unwrap();
            }
        }
        writeln!(code, "    }}").unwrap();
        writeln!(code, "}}").unwrap();

        Some(code)
    }

    // =========================================================================
    // Element group (EG_*) serializer generation
    // =========================================================================

    fn gen_element_group_serializer(&self, def: &crate::ast::Definition) -> Option<String> {
        let rust_name = self.to_rust_type_name(&def.name);

        let Pattern::Choice(variants) = &def.pattern else {
            return None;
        };

        // Collect element variants with their prefix for qualified tag names
        let mut element_variants = Vec::new();
        let mut visited = std::collections::HashSet::new();
        visited.insert(def.name.clone());
        for v in variants {
            self.collect_eg_variants_with_prefix(v, &mut element_variants, &mut visited);
        }
        // Dedup by xml local name
        let mut seen = std::collections::HashSet::new();
        element_variants.retain(|(_, local, _, _)| seen.insert(local.clone()));

        if element_variants.is_empty() {
            return None;
        }

        let mut code = String::new();
        writeln!(code, "impl ToXml for {} {{", rust_name).unwrap();
        writeln!(
            code,
            "    fn write_element<W: Write>(&self, _tag: &str, writer: &mut Writer<W>) -> Result<(), SerializeError> {{"
        )
        .unwrap();
        writeln!(code, "        match self {{").unwrap();

        for (_prefix, xml_local, _inner_type, _needs_box) in &element_variants {
            let variant_name = self.to_rust_variant_name(xml_local);
            let qualified_name = match self.xml_prefix() {
                Some(p) => format!("{}:{}", p, xml_local),
                None => xml_local.clone(),
            };
            writeln!(
                code,
                "            Self::{}(inner) => inner.write_element(\"{}\", writer)?,",
                variant_name, qualified_name
            )
            .unwrap();
        }

        writeln!(code, "        }}").unwrap();
        writeln!(code, "        Ok(())").unwrap();
        writeln!(code, "    }}").unwrap();
        writeln!(code, "}}").unwrap();

        Some(code)
    }

    /// Collect element variants with their XML prefix for qualified name construction.
    /// Returns (prefix, local_name, rust_type, needs_box).
    fn collect_eg_variants_with_prefix(
        &self,
        pattern: &Pattern,
        variants: &mut Vec<(Option<String>, String, String, bool)>,
        visited: &mut std::collections::HashSet<String>,
    ) {
        match pattern {
            Pattern::Element { name, pattern } => {
                let (inner_type, needs_box) = self.pattern_to_rust_type(pattern);
                variants.push((
                    name.prefix.clone(),
                    name.local.clone(),
                    inner_type,
                    needs_box,
                ));
            }
            Pattern::Optional(inner)
            | Pattern::ZeroOrMore(inner)
            | Pattern::OneOrMore(inner)
            | Pattern::Group(inner) => {
                self.collect_eg_variants_with_prefix(inner, variants, visited);
            }
            Pattern::Ref(name) if name.contains("_EG_") && visited.insert(name.clone()) => {
                if let Some(def_pattern) = self.definitions.get(name.as_str()) {
                    self.collect_eg_variants_with_prefix(def_pattern, variants, visited);
                }
            }
            Pattern::Choice(items) | Pattern::Sequence(items) | Pattern::Interleave(items) => {
                for item in items {
                    self.collect_eg_variants_with_prefix(item, variants, visited);
                }
            }
            _ => {}
        }
    }

    // =========================================================================
    // Attribute writing helpers
    // =========================================================================

    /// Generate code to push an attribute onto the start tag.
    fn write_attr_push(
        &self,
        code: &mut String,
        attr_name: &str,
        strategy: WriteStrategy,
        val_expr: &str,
        indent: &str,
    ) {
        match strategy {
            WriteStrategy::OoxmlBool => {
                writeln!(
                    code,
                    "{}start.push_attribute((\"{}\", if *{} {{ \"1\" }} else {{ \"0\" }}));",
                    indent, attr_name, val_expr
                )
                .unwrap();
            }
            WriteStrategy::AsString => {
                writeln!(
                    code,
                    "{}start.push_attribute((\"{}\", {}.as_str()));",
                    indent, attr_name, val_expr
                )
                .unwrap();
            }
            WriteStrategy::HexBinary => {
                writeln!(code, "{}{{", indent).unwrap();
                writeln!(code, "{}    let hex = encode_hex({});", indent, val_expr).unwrap();
                writeln!(
                    code,
                    "{}    start.push_attribute((\"{}\", hex.as_str()));",
                    indent, attr_name
                )
                .unwrap();
                writeln!(code, "{}}}", indent).unwrap();
            }
            WriteStrategy::DisplayStr | WriteStrategy::ToXml => {
                // For enums/numbers, use to_string(). ToXml shouldn't appear for attrs
                // but handle it as to_string() as fallback.
                writeln!(code, "{}{{", indent).unwrap();
                writeln!(code, "{}    let s = {}.to_string();", indent, val_expr).unwrap();
                writeln!(
                    code,
                    "{}    start.push_attribute((\"{}\", s.as_str()));",
                    indent, attr_name
                )
                .unwrap();
                writeln!(code, "{}}}", indent).unwrap();
            }
        }
    }

    // =========================================================================
    // Child element writing helpers
    // =========================================================================

    /// Generate code to write a child element.
    fn write_child_element(
        &self,
        code: &mut String,
        tag: &str,
        strategy: WriteStrategy,
        val_expr: &str,
        _is_owned: bool,
        indent: &str,
    ) {
        match strategy {
            WriteStrategy::ToXml => {
                writeln!(
                    code,
                    "{}{}.write_element(\"{}\", writer)?;",
                    indent, val_expr, tag
                )
                .unwrap();
            }
            WriteStrategy::AsString => {
                writeln!(code, "{}{{", indent).unwrap();
                writeln!(
                    code,
                    "{}    let start = BytesStart::new(\"{}\");",
                    indent, tag
                )
                .unwrap();
                writeln!(
                    code,
                    "{}    writer.write_event(Event::Start(start))?;",
                    indent
                )
                .unwrap();
                writeln!(
                    code,
                    "{}    writer.write_event(Event::Text(BytesText::new({}.as_str())))?;",
                    indent, val_expr
                )
                .unwrap();
                writeln!(
                    code,
                    "{}    writer.write_event(Event::End(BytesEnd::new(\"{}\")))?;",
                    indent, tag
                )
                .unwrap();
                writeln!(code, "{}}}", indent).unwrap();
            }
            WriteStrategy::DisplayStr => {
                writeln!(code, "{}{{", indent).unwrap();
                writeln!(code, "{}    let s = {}.to_string();", indent, val_expr).unwrap();
                writeln!(
                    code,
                    "{}    let start = BytesStart::new(\"{}\");",
                    indent, tag
                )
                .unwrap();
                writeln!(
                    code,
                    "{}    writer.write_event(Event::Start(start))?;",
                    indent
                )
                .unwrap();
                writeln!(
                    code,
                    "{}    writer.write_event(Event::Text(BytesText::new(&s)))?;",
                    indent
                )
                .unwrap();
                writeln!(
                    code,
                    "{}    writer.write_event(Event::End(BytesEnd::new(\"{}\")))?;",
                    indent, tag
                )
                .unwrap();
                writeln!(code, "{}}}", indent).unwrap();
            }
            WriteStrategy::HexBinary => {
                writeln!(code, "{}{{", indent).unwrap();
                writeln!(code, "{}    let hex = encode_hex({});", indent, val_expr).unwrap();
                writeln!(
                    code,
                    "{}    let start = BytesStart::new(\"{}\");",
                    indent, tag
                )
                .unwrap();
                writeln!(
                    code,
                    "{}    writer.write_event(Event::Start(start))?;",
                    indent
                )
                .unwrap();
                writeln!(
                    code,
                    "{}    writer.write_event(Event::Text(BytesText::new(&hex)))?;",
                    indent
                )
                .unwrap();
                writeln!(
                    code,
                    "{}    writer.write_event(Event::End(BytesEnd::new(\"{}\")))?;",
                    indent, tag
                )
                .unwrap();
                writeln!(code, "{}}}", indent).unwrap();
            }
            WriteStrategy::OoxmlBool => {
                writeln!(code, "{}{{", indent).unwrap();
                writeln!(
                    code,
                    "{}    let s = if *{} {{ \"1\" }} else {{ \"0\" }};",
                    indent, val_expr
                )
                .unwrap();
                writeln!(
                    code,
                    "{}    let start = BytesStart::new(\"{}\");",
                    indent, tag
                )
                .unwrap();
                writeln!(
                    code,
                    "{}    writer.write_event(Event::Start(start))?;",
                    indent
                )
                .unwrap();
                writeln!(
                    code,
                    "{}    writer.write_event(Event::Text(BytesText::new(s)))?;",
                    indent
                )
                .unwrap();
                writeln!(
                    code,
                    "{}    writer.write_event(Event::End(BytesEnd::new(\"{}\")))?;",
                    indent, tag
                )
                .unwrap();
                writeln!(code, "{}}}", indent).unwrap();
            }
        }
    }

    // =========================================================================
    // Qualified name helpers
    // =========================================================================

    /// Build the qualified attribute name (e.g. "r:id" or "val").
    fn qualified_attr_name(&self, field: &Field) -> String {
        match &field.xml_prefix {
            Some(p) => format!("{}:{}", p, field.xml_name),
            None => field.xml_name.clone(),
        }
    }

    /// Build the qualified element name (e.g. "w:body" for WML, "body" for SML).
    /// Uses the config's xml_serialize_prefix (None = unprefixed, Some(p) = p:).
    fn qualified_element_name(&self, field: &Field) -> String {
        match self.xml_prefix() {
            Some(p) => format!("{}:{}", p, field.xml_name),
            None => field.xml_name.clone(),
        }
    }

    // =========================================================================
    // Strategy determination
    // =========================================================================

    fn get_write_strategy(&self, pattern: &Pattern) -> WriteStrategy {
        match pattern {
            Pattern::Ref(name) => {
                if name.contains("_CT_") || name.contains("_EG_") {
                    return WriteStrategy::ToXml;
                }
                if let Some(def_pattern) = self.definitions.get(name.as_str()) {
                    // Check for external ref
                    if let Pattern::Ref(inner_name) = def_pattern
                        && !self.definitions.contains_key(inner_name.as_str())
                    {
                        return WriteStrategy::ToXml;
                    }
                    return self.get_write_strategy(def_pattern);
                }
                // Unknown ref with ST_ — enum
                if name.contains("_ST_") {
                    WriteStrategy::DisplayStr
                } else {
                    WriteStrategy::AsString
                }
            }
            Pattern::Datatype { library, name, .. } if library == "xsd" => match name.as_str() {
                "boolean" => WriteStrategy::OoxmlBool,
                "string" | "token" | "NCName" | "ID" | "IDREF" | "anyURI" | "dateTime" | "date"
                | "time" => WriteStrategy::AsString,
                "hexBinary" | "base64Binary" => WriteStrategy::HexBinary,
                _ => WriteStrategy::DisplayStr, // numbers
            },
            Pattern::Choice(variants)
                if variants
                    .iter()
                    .all(|v| matches!(v, Pattern::StringLiteral(_))) =>
            {
                WriteStrategy::DisplayStr
            }
            Pattern::StringLiteral(_) => WriteStrategy::AsString,
            Pattern::Empty => WriteStrategy::ToXml,
            Pattern::List(_) => WriteStrategy::AsString,
            _ => WriteStrategy::ToXml,
        }
    }

    // =========================================================================
    // Shared helpers (mirrored from parser_gen)
    // =========================================================================

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
            Pattern::Optional(inner) | Pattern::ZeroOrMore(inner) | Pattern::OneOrMore(inner) => {
                Self::is_direct_element_variant(inner)
            }
            _ => false,
        }
    }

    fn is_type_alias(&self, pattern: &Pattern) -> bool {
        match pattern {
            Pattern::Element { .. } => true,
            Pattern::Datatype { .. } => true,
            Pattern::Ref(name) => {
                if let Some(def_pattern) = self.definitions.get(name.as_str()) {
                    self.is_simple_type(def_pattern) || self.is_type_alias(def_pattern)
                } else {
                    false
                }
            }
            _ => false,
        }
    }

    fn is_string_type(&self, pattern: &Pattern) -> bool {
        match pattern {
            Pattern::Datatype { library, name, .. } => {
                library == "xsd" && (name == "string" || name == "token" || name == "NCName")
            }
            Pattern::Ref(name) => self
                .definitions
                .get(name.as_str())
                .is_some_and(|p| self.is_string_type(p)),
            _ => false,
        }
    }

    fn has_xml_children_pattern(&self, pattern: &Pattern) -> bool {
        match pattern {
            Pattern::Empty => false,
            Pattern::Attribute { .. } => false,
            Pattern::Element { .. } => true,
            Pattern::Ref(name) => {
                if name.contains("_AG_") {
                    return false;
                }
                if let Some(def_pattern) = self.definitions.get(name.as_str()) {
                    self.has_xml_children_pattern(def_pattern)
                } else {
                    true
                }
            }
            Pattern::Sequence(items) | Pattern::Interleave(items) | Pattern::Choice(items) => {
                items.iter().any(|i| self.has_xml_children_pattern(i))
            }
            Pattern::Optional(inner)
            | Pattern::ZeroOrMore(inner)
            | Pattern::OneOrMore(inner)
            | Pattern::Group(inner)
            | Pattern::Mixed(inner) => self.has_xml_children_pattern(inner),
            Pattern::Text => true,
            _ => false,
        }
    }

    fn has_xml_attr_pattern(&self, pattern: &Pattern) -> bool {
        match pattern {
            Pattern::Attribute { .. } => true,
            Pattern::Ref(name) if name.contains("_AG_") => true,
            Pattern::Ref(name) => {
                if let Some(def_pattern) = self.definitions.get(name.as_str()) {
                    self.has_xml_attr_pattern(def_pattern)
                } else {
                    false
                }
            }
            Pattern::Sequence(items) | Pattern::Interleave(items) | Pattern::Choice(items) => {
                items.iter().any(|i| self.has_xml_attr_pattern(i))
            }
            Pattern::Optional(inner)
            | Pattern::ZeroOrMore(inner)
            | Pattern::OneOrMore(inner)
            | Pattern::Group(inner) => self.has_xml_attr_pattern(inner),
            _ => false,
        }
    }

    fn is_eg_content_field(&self, field: &Field) -> bool {
        if let Pattern::Ref(name) = &field.pattern
            && name.contains("_EG_")
            && let Some(pattern) = self.definitions.get(name.as_str())
        {
            return self.is_element_choice(pattern);
        }
        false
    }

    fn eg_ref_to_field_name(&self, name: &str) -> String {
        let spec_name = strip_namespace_prefix(name);
        let short = spec_name.strip_prefix("EG_").unwrap_or(spec_name);
        // Check names.yaml field mapping first
        if let Some(mappings) = &self.config.name_mappings
            && let Some(mapped) = mappings.resolve_field(&self.config.module_name, short)
        {
            return mapped.to_string();
        }
        to_snake_case(short)
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
                    xml_prefix: name.prefix.clone(),
                    pattern: pattern.as_ref().clone(),
                    is_optional,
                    is_attribute: true,
                    is_vec: false,
                    is_text_content: false,
                });
            }
            Pattern::Element { name, pattern } => {
                // Skip wildcard elements (element * { ... }) — handled by extra_children
                if name.local == "_any" {
                    return;
                }
                fields.push(Field {
                    name: self.qname_to_field_name(name),
                    xml_name: name.local.clone(),
                    xml_prefix: name.prefix.clone(),
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
                Pattern::Element { name, pattern } if name.local != "_any" => {
                    fields.push(Field {
                        name: self.qname_to_field_name(name),
                        xml_name: name.local.clone(),
                        xml_prefix: name.prefix.clone(),
                        pattern: pattern.as_ref().clone(),
                        is_optional: false,
                        is_attribute: false,
                        is_vec: true,
                        is_text_content: false,
                    });
                }
                Pattern::Ref(name) if name.contains("_EG_") => {
                    if let Some(def_pattern) = self.definitions.get(name.as_str()) {
                        if self.is_element_choice(def_pattern) {
                            fields.push(Field {
                                name: self.eg_ref_to_field_name(name),
                                xml_name: name.clone(),
                                xml_prefix: None,
                                pattern: Pattern::Ref(name.clone()),
                                is_optional: false,
                                is_attribute: false,
                                is_vec: true,
                                is_text_content: false,
                            });
                        } else {
                            self.collect_fields(def_pattern, fields, true);
                        }
                    }
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
                if let Some(def_pattern) = self.definitions.get(name.as_str()) {
                    if name.contains("_EG_") {
                        if self.is_element_choice(def_pattern) {
                            fields.push(Field {
                                name: self.eg_ref_to_field_name(name),
                                xml_name: name.clone(),
                                xml_prefix: None,
                                pattern: Pattern::Ref(name.clone()),
                                is_optional,
                                is_attribute: false,
                                is_vec: false,
                                is_text_content: false,
                            });
                        } else {
                            self.collect_fields(def_pattern, fields, is_optional);
                        }
                    } else if name.contains("_AG_") {
                        self.collect_fields(def_pattern, fields, is_optional);
                    } else if self.is_string_type(def_pattern) {
                        fields.push(Field {
                            name: "text".to_string(),
                            xml_name: "$text".to_string(),
                            xml_prefix: None,
                            pattern: Pattern::Datatype {
                                library: "xsd".to_string(),
                                name: "string".to_string(),
                                params: vec![],
                            },
                            is_optional: true,
                            is_attribute: false,
                            is_vec: false,
                            is_text_content: true,
                        });
                    } else {
                        // CT_* mixin or base type — inline its fields
                        self.collect_fields(def_pattern, fields, is_optional);
                    }
                }
            }
            Pattern::Choice(_) => {}
            _ => {}
        }
    }
}
