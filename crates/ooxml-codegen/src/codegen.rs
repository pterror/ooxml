//! Rust code generator from parsed RNC schemas.

use crate::ast::{Definition, Pattern, QName, Schema};
use std::collections::HashMap;
use std::fmt::Write;

/// Code generation configuration.
#[derive(Debug, Clone)]
pub struct CodegenConfig {
    /// Namespace prefix to strip from type names (e.g., "w_" for WordprocessingML).
    pub strip_prefix: Option<String>,
    /// Module name for the generated code.
    pub module_name: String,
}

impl Default for CodegenConfig {
    fn default() -> Self {
        Self {
            strip_prefix: None,
            module_name: "types".to_string(),
        }
    }
}

/// Generate Rust code from a parsed schema.
pub fn generate(schema: &Schema, config: &CodegenConfig) -> String {
    let mut g = Generator::new(schema, config);
    g.run()
}

struct Generator<'a> {
    schema: &'a Schema,
    config: &'a CodegenConfig,
    output: String,
    /// Map from definition name to its pattern for resolution.
    definitions: HashMap<&'a str, &'a Pattern>,
}

impl<'a> Generator<'a> {
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

        // Separate simple types (enums) from complex types (structs)
        let (simple_types, complex_types): (Vec<_>, Vec<_>) = self
            .schema
            .definitions
            .iter()
            .partition(|d| d.name.contains("_ST_") || self.is_simple_type(&d.pattern));

        // Generate enums for simple types
        for def in &simple_types {
            if let Some(code) = self.gen_simple_type(def) {
                self.output.push_str(&code);
                self.output.push('\n');
            }
        }

        // Generate structs for complex types
        for def in &complex_types {
            if let Some(code) = self.gen_complex_type(def) {
                self.output.push_str(&code);
                self.output.push('\n');
            }
        }

        std::mem::take(&mut self.output)
    }

    fn write_header(&mut self) {
        writeln!(self.output, "//! Generated from ECMA-376 RELAX NG schema.").unwrap();
        writeln!(self.output, "//! Do not edit manually.").unwrap();
        writeln!(self.output).unwrap();
        writeln!(self.output, "#![allow(dead_code)]").unwrap();
        writeln!(self.output).unwrap();
    }

    fn is_simple_type(&self, pattern: &Pattern) -> bool {
        match pattern {
            Pattern::Choice(variants) => variants
                .iter()
                .all(|v| matches!(v, Pattern::StringLiteral(_))),
            Pattern::StringLiteral(_) => true,
            Pattern::Datatype { .. } => true,
            Pattern::Ref(name) => {
                // Check if the referenced type is simple
                self.definitions
                    .get(name.as_str())
                    .is_some_and(|p| self.is_simple_type(p))
            }
            _ => false,
        }
    }

    fn gen_simple_type(&self, def: &Definition) -> Option<String> {
        let rust_name = self.to_rust_type_name(&def.name);

        match &def.pattern {
            Pattern::Choice(variants) => {
                let string_variants: Vec<_> = variants
                    .iter()
                    .filter_map(|v| match v {
                        Pattern::StringLiteral(s) => Some(s.as_str()),
                        _ => None,
                    })
                    .collect();

                if string_variants.is_empty() {
                    return None;
                }

                let mut code = String::new();
                writeln!(code, "#[derive(Debug, Clone, Copy, PartialEq, Eq)]").unwrap();
                writeln!(code, "pub enum {} {{", rust_name).unwrap();

                for variant in &string_variants {
                    let variant_name = self.to_rust_variant_name(variant);
                    writeln!(code, "    {},", variant_name).unwrap();
                }

                writeln!(code, "}}").unwrap();
                Some(code)
            }
            Pattern::Datatype { library, name, .. } => {
                // Type alias for XSD types
                let rust_type = self.xsd_to_rust(library, name);
                let mut code = String::new();
                writeln!(code, "pub type {} = {};", rust_name, rust_type).unwrap();
                Some(code)
            }
            Pattern::Ref(target) => {
                // Type alias
                let target_rust = self.to_rust_type_name(target);
                let mut code = String::new();
                writeln!(code, "pub type {} = {};", rust_name, target_rust).unwrap();
                Some(code)
            }
            _ => None,
        }
    }

    fn gen_complex_type(&self, def: &Definition) -> Option<String> {
        // Skip element definitions (e.g., w_document = element document { ... })
        if matches!(&def.pattern, Pattern::Element { .. }) {
            return None;
        }

        let rust_name = self.to_rust_type_name(&def.name);
        let mut code = String::new();

        // Collect fields from the pattern
        let fields = self.extract_fields(&def.pattern);

        if fields.is_empty() {
            // Empty struct
            writeln!(code, "#[derive(Debug, Clone, Default)]").unwrap();
            writeln!(code, "pub struct {};", rust_name).unwrap();
        } else {
            writeln!(code, "#[derive(Debug, Clone)]").unwrap();
            writeln!(code, "pub struct {} {{", rust_name).unwrap();

            for field in &fields {
                let field_type = self.pattern_to_rust_type(&field.pattern, field.is_optional);
                writeln!(code, "    pub {}: {},", field.name, field_type).unwrap();
            }

            writeln!(code, "}}").unwrap();
        }

        Some(code)
    }

    fn extract_fields(&self, pattern: &Pattern) -> Vec<Field> {
        let mut fields = Vec::new();
        self.collect_fields(pattern, &mut fields, false);
        fields
    }

    fn collect_fields(&self, pattern: &Pattern, fields: &mut Vec<Field>, is_optional: bool) {
        match pattern {
            Pattern::Attribute { name, pattern } => {
                fields.push(Field {
                    name: self.qname_to_field_name(name),
                    pattern: pattern.as_ref().clone(),
                    is_optional,
                    is_attribute: true,
                });
            }
            Pattern::Element { name, pattern } => {
                fields.push(Field {
                    name: self.qname_to_field_name(name),
                    pattern: pattern.as_ref().clone(),
                    is_optional,
                    is_attribute: false,
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
            Pattern::ZeroOrMore(inner) | Pattern::OneOrMore(inner) => {
                // These become Vec<T> fields
                self.collect_fields(inner, fields, is_optional);
            }
            Pattern::Group(inner) => {
                self.collect_fields(inner, fields, is_optional);
            }
            Pattern::Ref(name) => {
                // If referencing an element group (EG_*), inline its fields
                if name.contains("_EG_")
                    && let Some(target) = self.definitions.get(name.as_str())
                {
                    self.collect_fields(target, fields, is_optional);
                }
                // Otherwise it's a type reference, not a field
            }
            Pattern::Choice(_) => {
                // Choice patterns are complex - for now, skip them
                // TODO: Generate enum variants or handle differently
            }
            _ => {}
        }
    }

    fn to_rust_type_name(&self, name: &str) -> String {
        let name = if let Some(prefix) = &self.config.strip_prefix {
            name.strip_prefix(prefix).unwrap_or(name)
        } else {
            name
        };

        // Remove CT_, ST_, EG_ prefixes and convert to PascalCase
        let name = name
            .strip_prefix("CT_")
            .or_else(|| name.strip_prefix("ST_"))
            .or_else(|| name.strip_prefix("EG_"))
            .unwrap_or(name);

        to_pascal_case(name)
    }

    fn to_rust_variant_name(&self, name: &str) -> String {
        to_pascal_case(name)
    }

    fn qname_to_field_name(&self, qname: &QName) -> String {
        to_snake_case(&qname.local)
    }

    fn xsd_to_rust(&self, library: &str, name: &str) -> &'static str {
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
                "dateTime" => "String", // TODO: use chrono
                "date" => "String",
                "time" => "String",
                "hexBinary" => "Vec<u8>",
                "base64Binary" => "Vec<u8>",
                "anyURI" => "String",
                "token" => "String",
                "NCName" => "String",
                "ID" => "String",
                "IDREF" => "String",
                _ => "String",
            }
        } else {
            "String"
        }
    }

    fn pattern_to_rust_type(&self, pattern: &Pattern, is_optional: bool) -> String {
        let inner = match pattern {
            Pattern::Ref(name) => self.to_rust_type_name(name),
            Pattern::Datatype { library, name, .. } => self.xsd_to_rust(library, name).to_string(),
            Pattern::Empty => "()".to_string(),
            Pattern::StringLiteral(_) => "String".to_string(),
            Pattern::Choice(_) => "String".to_string(), // Simplified
            _ => "String".to_string(),
        };

        if is_optional {
            format!("Option<{}>", inner)
        } else {
            inner
        }
    }
}

struct Field {
    name: String,
    pattern: Pattern,
    is_optional: bool,
    #[allow(dead_code)] // Will be used for XML attribute annotations
    is_attribute: bool,
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

    // Handle reserved keywords
    match result.as_str() {
        "type" => "r#type".to_string(),
        "ref" => "r#ref".to_string(),
        "match" => "r#match".to_string(),
        "in" => "r#in".to_string(),
        "for" => "r#for".to_string(),
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
        "move" => "r#move".to_string(),
        "mut" => "r#mut".to_string(),
        "where" => "r#where".to_string(),
        "async" => "r#async".to_string(),
        "await" => "r#await".to_string(),
        "dyn" => "r#dyn".to_string(),
        "box" => "r#box".to_string(),
        "true" => "r#true".to_string(),
        "false" => "r#false".to_string(),
        _ => result,
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_to_pascal_case() {
        assert_eq!(to_pascal_case("foo_bar"), "FooBar");
        assert_eq!(to_pascal_case("fooBar"), "FooBar");
        assert_eq!(to_pascal_case("FOO"), "FOO");
    }

    #[test]
    fn test_to_snake_case() {
        assert_eq!(to_snake_case("fooBar"), "foo_bar");
        assert_eq!(to_snake_case("FooBar"), "foo_bar");
        assert_eq!(to_snake_case("type"), "r#type");
    }
}
