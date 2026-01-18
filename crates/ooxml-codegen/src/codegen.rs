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

        // Categorize definitions into simple types, element groups, and complex types
        let mut simple_types = Vec::new();
        let mut element_groups = Vec::new();
        let mut complex_types = Vec::new();

        for def in &self.schema.definitions {
            if def.name.contains("_ST_") || self.is_simple_type(&def.pattern) {
                simple_types.push(def);
            } else if def.name.contains("_EG_") && self.is_element_choice(&def.pattern) {
                element_groups.push(def);
            } else {
                complex_types.push(def);
            }
        }

        // Generate enums for simple types (string literal choices)
        for def in &simple_types {
            if let Some(code) = self.gen_simple_type(def) {
                self.output.push_str(&code);
                self.output.push('\n');
            }
        }

        // Generate enums for element groups (element choice patterns)
        for def in &element_groups {
            if let Some(code) = self.gen_element_group(def) {
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
        writeln!(self.output, "// Generated from ECMA-376 RELAX NG schema.").unwrap();
        writeln!(self.output, "// Do not edit manually.").unwrap();
        writeln!(self.output).unwrap();
        writeln!(self.output, "use serde::{{Deserialize, Serialize}};").unwrap();
        writeln!(self.output).unwrap();

        // Generate namespace constants
        if !self.schema.namespaces.is_empty() {
            writeln!(self.output, "/// XML namespace URIs used in this schema.").unwrap();
            writeln!(self.output, "pub mod ns {{").unwrap();

            for ns in &self.schema.namespaces {
                let const_name = ns.prefix.to_uppercase();
                if ns.is_default {
                    writeln!(
                        self.output,
                        "    /// Default namespace (prefix: {})",
                        ns.prefix
                    )
                    .unwrap();
                } else {
                    writeln!(self.output, "    /// Namespace prefix: {}", ns.prefix).unwrap();
                }
                writeln!(
                    self.output,
                    "    pub const {}: &str = \"{}\";",
                    const_name, ns.uri
                )
                .unwrap();
            }

            writeln!(self.output, "}}").unwrap();
            writeln!(self.output).unwrap();
        }
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

    /// Check if a pattern is a choice of elements (for element groups).
    fn is_element_choice(&self, pattern: &Pattern) -> bool {
        match pattern {
            Pattern::Choice(variants) => {
                // At least one variant must be an element (not just refs)
                // and we need to be able to extract at least some element variants
                variants.iter().any(Self::is_direct_element_variant)
            }
            _ => false,
        }
    }

    /// Check if a pattern is a direct element variant (not a ref to another EG_*).
    fn is_direct_element_variant(pattern: &Pattern) -> bool {
        match pattern {
            Pattern::Element { .. } => true,
            Pattern::Optional(inner) => Self::is_direct_element_variant(inner),
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

                if !string_variants.is_empty() {
                    // Deduplicate by Rust variant name (keep first occurrence)
                    let mut seen_variants = std::collections::HashSet::new();
                    let dedup_variants: Vec<_> = string_variants
                        .iter()
                        .filter(|v| {
                            let name = self.to_rust_variant_name(v);
                            seen_variants.insert(name)
                        })
                        .copied()
                        .collect();

                    // Enum of string literals
                    let mut code = String::new();
                    writeln!(
                        code,
                        "#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]"
                    )
                    .unwrap();
                    writeln!(code, "pub enum {} {{", rust_name).unwrap();

                    for variant in &dedup_variants {
                        let variant_name = self.to_rust_variant_name(variant);
                        // Add serde rename to preserve original XML value
                        writeln!(code, "    #[serde(rename = \"{}\")]", variant).unwrap();
                        writeln!(code, "    {},", variant_name).unwrap();
                    }

                    writeln!(code, "}}").unwrap();
                    writeln!(code).unwrap();

                    // Generate Display impl
                    writeln!(code, "impl std::fmt::Display for {} {{", rust_name).unwrap();
                    writeln!(
                        code,
                        "    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {{"
                    )
                    .unwrap();
                    writeln!(code, "        match self {{").unwrap();
                    for variant in &dedup_variants {
                        let variant_name = self.to_rust_variant_name(variant);
                        writeln!(
                            code,
                            "            Self::{} => write!(f, \"{}\"),",
                            variant_name, variant
                        )
                        .unwrap();
                    }
                    writeln!(code, "        }}").unwrap();
                    writeln!(code, "    }}").unwrap();
                    writeln!(code, "}}").unwrap();
                    writeln!(code).unwrap();

                    // Generate FromStr impl (include all string variants for parsing)
                    writeln!(code, "impl std::str::FromStr for {} {{", rust_name).unwrap();
                    writeln!(code, "    type Err = String;").unwrap();
                    writeln!(code).unwrap();
                    writeln!(
                        code,
                        "    fn from_str(s: &str) -> Result<Self, Self::Err> {{"
                    )
                    .unwrap();
                    writeln!(code, "        match s {{").unwrap();
                    // Use original string_variants for FromStr to handle aliases
                    for variant in &string_variants {
                        let variant_name = self.to_rust_variant_name(variant);
                        writeln!(
                            code,
                            "            \"{}\" => Ok(Self::{}),",
                            variant, variant_name
                        )
                        .unwrap();
                    }
                    writeln!(
                        code,
                        "            _ => Err(format!(\"unknown {} value: {{}}\", s)),",
                        rust_name
                    )
                    .unwrap();
                    writeln!(code, "        }}").unwrap();
                    writeln!(code, "    }}").unwrap();
                    writeln!(code, "}}").unwrap();

                    return Some(code);
                }

                // Choice of non-string types (e.g., xsd:integer | s_ST_Something)
                // Generate a type alias to String as fallback
                let mut code = String::new();
                writeln!(code, "pub type {} = String;", rust_name).unwrap();
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
                // Type alias - check if target exists in this schema
                let target_rust = if self.definitions.contains_key(target.as_str()) {
                    self.to_rust_type_name(target)
                } else {
                    // Unknown type from another schema - use String as fallback
                    "String".to_string()
                };
                let mut code = String::new();
                writeln!(code, "pub type {} = {};", rust_name, target_rust).unwrap();
                Some(code)
            }
            _ => None,
        }
    }

    fn gen_element_group(&self, def: &Definition) -> Option<String> {
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
            // Fallback to type alias
            let mut code = String::new();
            writeln!(code, "pub type {} = String;", rust_name).unwrap();
            return Some(code);
        }

        let mut code = String::new();
        writeln!(code, "#[derive(Debug, Clone, Serialize, Deserialize)]").unwrap();
        writeln!(code, "pub enum {} {{", rust_name).unwrap();

        for (xml_name, inner_type) in &element_variants {
            let variant_name = self.to_rust_variant_name(xml_name);
            writeln!(code, "    #[serde(rename = \"{}\")]", xml_name).unwrap();
            writeln!(code, "    {}({}),", variant_name, inner_type).unwrap();
        }

        writeln!(code, "}}").unwrap();

        Some(code)
    }

    /// Extract element info from a choice variant: (xml_name, rust_type)
    /// Only extracts direct Element patterns, not refs to other EG_* groups.
    fn extract_element_variant(&self, pattern: &Pattern) -> Option<(String, String)> {
        match pattern {
            Pattern::Element { name, pattern } => {
                let inner_type = self.pattern_to_rust_type(pattern, false);
                Some((name.local.clone(), inner_type))
            }
            Pattern::Optional(inner) => self.extract_element_variant(inner),
            Pattern::ZeroOrMore(inner) | Pattern::OneOrMore(inner) => {
                // For repeated elements in a choice, still extract but mark the type
                self.extract_element_variant(inner)
            }
            _ => None,
        }
    }

    fn gen_complex_type(&self, def: &Definition) -> Option<String> {
        // For element-only definitions, generate a type alias to the inner type
        if let Pattern::Element { pattern, .. } = &def.pattern {
            let rust_name = self.to_rust_type_name(&def.name);
            let inner_type = self.pattern_to_rust_type(pattern, false);
            let mut code = String::new();
            writeln!(code, "pub type {} = {};", rust_name, inner_type).unwrap();
            return Some(code);
        }

        let rust_name = self.to_rust_type_name(&def.name);
        let mut code = String::new();

        // Collect fields from the pattern
        let fields = self.extract_fields(&def.pattern);

        if fields.is_empty() {
            // Empty struct
            writeln!(
                code,
                "#[derive(Debug, Clone, Default, Serialize, Deserialize)]"
            )
            .unwrap();
            writeln!(code, "pub struct {};", rust_name).unwrap();
        } else {
            writeln!(code, "#[derive(Debug, Clone, Serialize, Deserialize)]").unwrap();
            writeln!(code, "pub struct {} {{", rust_name).unwrap();

            for field in &fields {
                let inner_type = self.pattern_to_rust_type(&field.pattern, false);
                let field_type = if field.is_vec {
                    format!("Vec<{}>", inner_type)
                } else if field.is_optional {
                    format!("Option<{}>", inner_type)
                } else {
                    inner_type
                };

                // Add serde attributes
                let xml_name = &field.xml_name;
                if field.is_attribute {
                    writeln!(code, "    #[serde(rename = \"@{}\")]", xml_name).unwrap();
                } else {
                    writeln!(code, "    #[serde(rename = \"{}\")]", xml_name).unwrap();
                }
                if field.is_optional || field.is_vec {
                    writeln!(code, "    #[serde(default)]").unwrap();
                }
                writeln!(code, "    pub {}: {},", field.name, field_type).unwrap();
            }

            writeln!(code, "}}").unwrap();
        }

        Some(code)
    }

    fn extract_fields(&self, pattern: &Pattern) -> Vec<Field> {
        let mut fields = Vec::new();
        self.collect_fields(pattern, &mut fields, false);
        // Deduplicate by name (keep first occurrence)
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
            Pattern::ZeroOrMore(inner) | Pattern::OneOrMore(inner) => {
                // These become Vec<T> fields
                match inner.as_ref() {
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
                        // EG_* element group references skipped - need mixed content handling
                        let _ = name;
                    }
                    Pattern::Choice(_) | Pattern::Ref(_) => {
                        // Complex repeated content - recurse but don't add directly
                        self.collect_fields(inner, fields, false);
                    }
                    _ => {}
                }
            }
            Pattern::Group(inner) => {
                self.collect_fields(inner, fields, is_optional);
            }
            Pattern::Ref(name) => {
                // EG_* element group references are skipped for now - they need
                // special mixed content handling that isn't implemented yet.
                // The EG_* enums are generated separately and can be used manually.
                // Non-EG_* refs are type references, not fields.
                let _ = name;
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

        // Convert to PascalCase, keeping CT/ST/EG prefix
        to_pascal_case(name)
    }

    fn to_rust_variant_name(&self, name: &str) -> String {
        // Handle empty string variant
        if name.is_empty() {
            return "Empty".to_string();
        }
        let name = to_pascal_case(name);
        // Prefix with underscore if starts with digit
        if name.chars().next().is_some_and(|c| c.is_ascii_digit()) {
            format!("_{}", name)
        } else {
            name
        }
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
        let (inner, needs_box) = match pattern {
            Pattern::Ref(name) => {
                // Check if this is a known definition
                if self.definitions.contains_key(name.as_str()) {
                    let type_name = self.to_rust_type_name(name);
                    // Box complex types (CT_*) and element groups (EG_*) to avoid infinite size
                    let needs_box = name.contains("_CT_") || name.contains("_EG_");
                    (type_name, needs_box)
                } else {
                    // Unknown reference (likely from another schema) - use String as fallback
                    ("String".to_string(), false)
                }
            }
            Pattern::Datatype { library, name, .. } => {
                (self.xsd_to_rust(library, name).to_string(), false)
            }
            Pattern::Empty => ("()".to_string(), false),
            Pattern::StringLiteral(_) => ("String".to_string(), false),
            Pattern::Choice(_) => ("String".to_string(), false),
            _ => ("String".to_string(), false),
        };

        let inner = if needs_box {
            format!("Box<{}>", inner)
        } else {
            inner
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
    xml_name: String,
    pattern: Pattern,
    is_optional: bool,
    is_attribute: bool,
    is_vec: bool,
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
