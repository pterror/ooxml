//! Static analysis for codegen configuration files.
//!
//! Analyzes schemas against ooxml-names.yaml and ooxml-features.yaml to find
//! unmapped types and fields.

use crate::ast::{Pattern, Schema};
use crate::codegen::CodegenConfig;
use std::collections::{HashMap, HashSet};

/// Analysis report for a single module.
#[derive(Debug, Default)]
pub struct ModuleReport {
    /// Types that exist in schema but have no name mapping.
    pub unmapped_types: Vec<String>,
    /// Fields (type.field) that exist in schema but have no feature mapping.
    pub unmapped_fields: Vec<String>,
    /// Total types analyzed.
    pub total_types: usize,
    /// Total fields analyzed.
    pub total_fields: usize,
}

impl ModuleReport {
    /// Check if the report has any unmapped items.
    pub fn has_unmapped(&self) -> bool {
        !self.unmapped_types.is_empty() || !self.unmapped_fields.is_empty()
    }

    /// Print the report to stderr.
    pub fn print(&self, module: &str) {
        if self.unmapped_types.is_empty() && self.unmapped_fields.is_empty() {
            eprintln!(
                "  {} types, {} fields - all mapped ✓",
                self.total_types, self.total_fields
            );
            return;
        }

        eprintln!(
            "  {} types ({} unmapped), {} fields ({} unmapped)",
            self.total_types,
            self.unmapped_types.len(),
            self.total_fields,
            self.unmapped_fields.len()
        );

        if !self.unmapped_types.is_empty() {
            eprintln!("  Unmapped types in ooxml-names.yaml [{}]:", module);
            for t in &self.unmapped_types {
                eprintln!("    - {}", t);
            }
        }

        if !self.unmapped_fields.is_empty() {
            eprintln!("  Unmapped fields in ooxml-features.yaml [{}]:", module);
            for f in &self.unmapped_fields {
                eprintln!("    - {}", f);
            }
        }
    }
}

/// Analyze a schema against configuration files.
pub fn analyze_schema(schema: &Schema, config: &CodegenConfig) -> ModuleReport {
    let mut report = ModuleReport::default();

    // Build definition map
    let definitions: HashMap<&str, &Pattern> = schema
        .definitions
        .iter()
        .map(|d| (d.name.as_str(), &d.pattern))
        .collect();

    // Analyze each definition
    for def in &schema.definitions {
        // Skip inline refs, simple types, and element groups
        if is_inline_attribute_ref(&def.name, &def.pattern)
            || is_simple_type(&def.pattern)
            || (def.name.contains("_EG_") && is_element_choice(&def.pattern, &definitions))
        {
            continue;
        }

        // This is a complex type that generates a struct
        let spec_name = strip_namespace_prefix(&def.name, &config.strip_prefix);

        // Check if type has name mapping
        report.total_types += 1;
        if !has_type_mapping(config, spec_name) {
            report.unmapped_types.push(spec_name.to_string());
        }

        // Collect and check fields
        let fields = collect_fields(&def.pattern, &definitions, &config.strip_prefix);
        for field in fields {
            report.total_fields += 1;
            if !has_field_mapping(config, spec_name, &field) {
                report
                    .unmapped_fields
                    .push(format!("{}.{}", spec_name, field));
            }
        }
    }

    report
}

/// Check if a type has a name mapping in the config.
fn has_type_mapping(config: &CodegenConfig, spec_name: &str) -> bool {
    if let Some(ref mappings) = config.name_mappings {
        let module_mappings = mappings.for_module(&config.module_name);
        // Check if there's a type mapping
        if module_mappings.types.contains_key(spec_name) {
            return true;
        }
        // Check shared mappings
        if mappings.shared.types.contains_key(spec_name) {
            return true;
        }
    }
    // If no mappings configured, consider it "mapped" (using default naming)
    config.name_mappings.is_none()
}

/// Check if a field has a feature mapping in the config.
fn has_field_mapping(config: &CodegenConfig, type_name: &str, field_name: &str) -> bool {
    if let Some(ref mappings) = config.feature_mappings {
        let module_features = mappings.for_module(&config.module_name);
        // Check if the type has any field mappings
        if let Some(type_fields) = module_features.get(type_name) {
            // If the type is listed, check if this specific field is mapped
            // (or if there's a wildcard)
            return type_fields.contains_key(field_name) || type_fields.contains_key("*");
        }
    }
    // If no mappings configured, consider it "mapped"
    config.feature_mappings.is_none()
}

/// Collect field names from a pattern.
fn collect_fields(
    pattern: &Pattern,
    definitions: &HashMap<&str, &Pattern>,
    strip_prefix: &Option<String>,
) -> Vec<String> {
    let mut fields = Vec::new();
    collect_fields_recursive(
        pattern,
        definitions,
        strip_prefix,
        &mut fields,
        &mut HashSet::new(),
    );
    fields
}

fn collect_fields_recursive(
    pattern: &Pattern,
    definitions: &HashMap<&str, &Pattern>,
    strip_prefix: &Option<String>,
    fields: &mut Vec<String>,
    visited: &mut HashSet<String>,
) {
    match pattern {
        Pattern::Group(inner)
        | Pattern::Optional(inner)
        | Pattern::ZeroOrMore(inner)
        | Pattern::OneOrMore(inner)
        | Pattern::Mixed(inner) => {
            collect_fields_recursive(inner, definitions, strip_prefix, fields, visited);
        }
        Pattern::Interleave(parts) | Pattern::Choice(parts) | Pattern::Sequence(parts) => {
            for part in parts {
                collect_fields_recursive(part, definitions, strip_prefix, fields, visited);
            }
        }
        Pattern::Attribute { name, .. } => {
            let field_name = to_snake_case(&name.local);
            if !fields.contains(&field_name) {
                fields.push(field_name);
            }
        }
        Pattern::Element { name, .. } => {
            let field_name = to_snake_case(&name.local);
            if !fields.contains(&field_name) {
                fields.push(field_name);
            }
        }
        Pattern::Ref(name) => {
            // Follow refs to inline attributes/groups
            if visited.insert(name.clone())
                && let Some(ref_pattern) = definitions.get(name.as_str())
            {
                // Only follow AG_* (attribute groups) and CT_* base types
                if name.contains("_AG_") || is_inline_attribute_ref(name, ref_pattern) {
                    collect_fields_recursive(
                        ref_pattern,
                        definitions,
                        strip_prefix,
                        fields,
                        visited,
                    );
                } else if name.contains("_EG_") {
                    // Element groups become their own field
                    let spec_name = strip_namespace_prefix(name, strip_prefix);
                    let short = spec_name.strip_prefix("EG_").unwrap_or(spec_name);
                    let field_name = to_snake_case(short);
                    if !fields.contains(&field_name) {
                        fields.push(field_name);
                    }
                }
            }
        }
        Pattern::Empty
        | Pattern::Text
        | Pattern::Any
        | Pattern::StringLiteral(_)
        | Pattern::Datatype { .. }
        | Pattern::List(_) => {}
    }
}

fn strip_namespace_prefix<'a>(name: &'a str, prefix: &Option<String>) -> &'a str {
    if let Some(p) = prefix {
        name.strip_prefix(p).unwrap_or(name)
    } else {
        name
    }
}

fn is_inline_attribute_ref(name: &str, pattern: &Pattern) -> bool {
    // Inline attribute refs like "r_id = attribute r:id {...}"
    matches!(pattern, Pattern::Attribute { .. }) && !name.contains("_CT_") && !name.contains("_AG_")
}

fn is_simple_type(pattern: &Pattern) -> bool {
    match pattern {
        Pattern::Choice(variants) => variants.iter().all(is_simple_type),
        Pattern::StringLiteral(_) | Pattern::Datatype { .. } | Pattern::Text => true,
        Pattern::Group(inner) => is_simple_type(inner),
        _ => false,
    }
}

fn is_element_choice(pattern: &Pattern, definitions: &HashMap<&str, &Pattern>) -> bool {
    match pattern {
        Pattern::Choice(variants) => variants
            .iter()
            .all(|v| is_element_or_ref_to_element(v, definitions)),
        Pattern::Group(inner) => is_element_choice(inner, definitions),
        _ => false,
    }
}

fn is_element_or_ref_to_element(pattern: &Pattern, definitions: &HashMap<&str, &Pattern>) -> bool {
    match pattern {
        Pattern::Element { .. } => true,
        Pattern::Ref(name) => {
            if let Some(p) = definitions.get(name.as_str()) {
                matches!(p, Pattern::Element { .. }) || is_element_choice(p, definitions)
            } else {
                false
            }
        }
        Pattern::Choice(variants) => variants
            .iter()
            .all(|v| is_element_or_ref_to_element(v, definitions)),
        Pattern::Group(inner) => is_element_or_ref_to_element(inner, definitions),
        _ => false,
    }
}

fn to_snake_case(s: &str) -> String {
    let mut result = String::new();
    let mut prev_lower = false;

    for c in s.chars() {
        if c.is_uppercase() {
            if prev_lower {
                result.push('_');
            }
            result.push(c.to_lowercase().next().unwrap());
            prev_lower = false;
        } else {
            result.push(c);
            prev_lower = c.is_lowercase();
        }
    }

    result
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_to_snake_case() {
        assert_eq!(to_snake_case("fooBar"), "foo_bar");
        assert_eq!(to_snake_case("FooBar"), "foo_bar");
        assert_eq!(to_snake_case("foo"), "foo");
        // All-caps at start stays lowercase (realistic for OOXML attr names)
        assert_eq!(to_snake_case("XMLParser"), "xmlparser");
        assert_eq!(to_snake_case("val"), "val");
        assert_eq!(to_snake_case("colId"), "col_id");
    }
}
