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
                } else if !matches!(&def.pattern, Pattern::Element { .. }) {
                    // Complex type struct - generate struct parser
                    if let Some(code) = self.gen_struct_parser(def) {
                        self.output.push_str(&code);
                        self.output.push('\n');
                    }
                }
            }
        }

        std::mem::take(&mut self.output)
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
        writeln!(self.output).unwrap();
        writeln!(self.output, "use super::generated::*;").unwrap();
        writeln!(self.output, "use quick_xml::Reader;").unwrap();
        writeln!(self.output, "use quick_xml::events::{{Event, BytesStart}};").unwrap();
        writeln!(self.output, "use std::io::BufRead;").unwrap();
        writeln!(self.output).unwrap();
        writeln!(self.output, "/// Error type for XML parsing.").unwrap();
        writeln!(self.output, "#[derive(Debug)]").unwrap();
        writeln!(self.output, "pub enum ParseError {{").unwrap();
        writeln!(self.output, "    Xml(quick_xml::Error),").unwrap();
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
        // Add skip_element helper
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
            "    fn from_xml<R: BufRead>(reader: &mut Reader<R>, start: &BytesStart, is_empty: bool) -> Result<Self, ParseError> {{"
        )
        .unwrap();
        writeln!(code, "        let tag = start.name();").unwrap();
        writeln!(code, "        match tag.as_ref() {{").unwrap();

        for (xml_name, inner_type, needs_box) in &element_variants {
            let variant_name = self.to_rust_variant_name(xml_name);
            writeln!(code, "            b\"{}\" => {{", xml_name).unwrap();
            if *needs_box {
                writeln!(
                    code,
                    "                let inner = {}::from_xml(reader, start, is_empty)?;",
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
                    "                let inner = {}::from_xml(reader, start, is_empty)?;",
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
            // Empty struct - simple impl
            let mut code = String::new();
            writeln!(code, "impl FromXml for {} {{", rust_name).unwrap();
            writeln!(
                code,
                "    fn from_xml<R: BufRead>(reader: &mut Reader<R>, _start: &BytesStart, is_empty: bool) -> Result<Self, ParseError> {{"
            )
            .unwrap();
            writeln!(code, "        if !is_empty {{").unwrap();
            writeln!(code, "            // Skip to end tag").unwrap();
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
            writeln!(code, "        Ok(Self {{}})").unwrap();
            writeln!(code, "    }}").unwrap();
            writeln!(code, "}}").unwrap();
            return Some(code);
        }

        let mut code = String::new();
        writeln!(code, "impl FromXml for {} {{", rust_name).unwrap();
        writeln!(
            code,
            "    fn from_xml<R: BufRead>(reader: &mut Reader<R>, start: &BytesStart, is_empty: bool) -> Result<Self, ParseError> {{"
        )
        .unwrap();

        // Declare field variables
        for field in &fields {
            if field.is_vec {
                writeln!(code, "        let mut {} = Vec::new();", field.name).unwrap();
            } else if field.is_optional {
                writeln!(code, "        let mut {} = None;", field.name).unwrap();
            } else {
                let (rust_type, needs_box) = self.pattern_to_rust_type(&field.pattern);
                let full_type = if needs_box {
                    format!("Box<{}>", rust_type)
                } else {
                    rust_type
                };
                writeln!(
                    code,
                    "        let mut {}: Option<{}> = None;",
                    field.name, full_type
                )
                .unwrap();
            }
        }

        // Parse attributes
        let attr_fields: Vec<_> = fields.iter().filter(|f| f.is_attribute).collect();
        if !attr_fields.is_empty() {
            writeln!(code).unwrap();
            writeln!(code, "        // Parse attributes").unwrap();
            writeln!(
                code,
                "        for attr in start.attributes().filter_map(|a| a.ok()) {{"
            )
            .unwrap();
            writeln!(code, "            match attr.key.as_ref() {{").unwrap();
            for field in &attr_fields {
                let parse_expr = self.gen_attr_parse_expr(&field.pattern);
                writeln!(code, "                b\"{}\" => {{", field.xml_name).unwrap();
                writeln!(
                    code,
                    "                    let val = String::from_utf8_lossy(&attr.value);"
                )
                .unwrap();
                writeln!(code, "                    {} = {};", field.name, parse_expr).unwrap();
                writeln!(code, "                }}").unwrap();
            }
            writeln!(code, "                _ => {{}}").unwrap();
            writeln!(code, "            }}").unwrap();
            writeln!(code, "        }}").unwrap();
        }

        // Parse child elements (only if not empty element)
        let elem_fields: Vec<_> = fields.iter().filter(|f| !f.is_attribute).collect();
        if !elem_fields.is_empty() {
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
                let (rust_type, needs_box) = self.pattern_to_rust_type(&field.pattern);
                writeln!(
                    code,
                    "                            b\"{}\" => {{",
                    field.xml_name
                )
                .unwrap();
                if field.is_vec {
                    if needs_box {
                        writeln!(
                            code,
                            "                                {}.push(Box::new({}::from_xml(reader, &e, false)?));",
                            field.name, rust_type
                        )
                        .unwrap();
                    } else {
                        writeln!(
                            code,
                            "                                {}.push({}::from_xml(reader, &e, false)?);",
                            field.name, rust_type
                        )
                        .unwrap();
                    }
                } else if needs_box {
                    writeln!(
                        code,
                        "                                {} = Some(Box::new({}::from_xml(reader, &e, false)?));",
                        field.name, rust_type
                    )
                    .unwrap();
                } else {
                    writeln!(
                        code,
                        "                                {} = Some({}::from_xml(reader, &e, false)?);",
                        field.name, rust_type
                    )
                    .unwrap();
                }
                writeln!(code, "                            }}").unwrap();
            }

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
                let (rust_type, needs_box) = self.pattern_to_rust_type(&field.pattern);
                writeln!(
                    code,
                    "                            b\"{}\" => {{",
                    field.xml_name
                )
                .unwrap();
                // For empty elements, pass is_empty=true
                if field.is_vec {
                    if needs_box {
                        writeln!(
                            code,
                            "                                {}.push(Box::new({}::from_xml(reader, &e, true)?));",
                            field.name, rust_type
                        )
                        .unwrap();
                    } else {
                        writeln!(
                            code,
                            "                                {}.push({}::from_xml(reader, &e, true)?);",
                            field.name, rust_type
                        )
                        .unwrap();
                    }
                } else if needs_box {
                    writeln!(
                        code,
                        "                                {} = Some(Box::new({}::from_xml(reader, &e, true)?));",
                        field.name, rust_type
                    )
                    .unwrap();
                } else {
                    writeln!(
                        code,
                        "                                {} = Some({}::from_xml(reader, &e, true)?);",
                        field.name, rust_type
                    )
                    .unwrap();
                }
                writeln!(code, "                            }}").unwrap();
            }

            writeln!(code, "                            _ => {{}}").unwrap();
            writeln!(code, "                        }}").unwrap();
            writeln!(code, "                    }}").unwrap();
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

        // Build result struct
        writeln!(code).unwrap();
        writeln!(code, "        Ok(Self {{").unwrap();
        for field in &fields {
            if field.is_optional || field.is_vec {
                writeln!(code, "            {},", field.name).unwrap();
            } else {
                // Required field - unwrap with error
                writeln!(
                    code,
                    "            {}: {}.ok_or_else(|| ParseError::MissingAttribute(\"{}\".to_string()))?,",
                    field.name, field.name, field.xml_name
                )
                .unwrap();
            }
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
                _ => "Some(val.into_owned())".to_string(),
            },
            Pattern::Ref(name) if name.contains("_ST_") => {
                // Enum type - use FromStr
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
            Pattern::Ref(_) | Pattern::Choice(_) => {}
            _ => {}
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
}

struct Field {
    name: String,
    xml_name: String,
    pattern: Pattern,
    is_optional: bool,
    is_attribute: bool,
    is_vec: bool,
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
