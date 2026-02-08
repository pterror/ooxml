//! Static analysis for codegen configuration files.
//!
//! Analyzes schemas against ooxml-names.yaml and ooxml-features.yaml to find
//! unmapped types and fields.

use crate::ast::{Pattern, Schema};
use crate::codegen::{CodegenConfig, to_pascal_case};
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
        // Skip inline refs, simple types, attribute groups, element groups,
        // and root element wrappers (e.g., `document = element w:document { CT_Document }`).
        // AG_* and EG_* are always inlined into parent types, so their fields
        // are already accounted for in the parent type's feature mapping.
        // Root element wrappers are deduplicated by the codegen (their PascalCase
        // name matches the mapped CT_* type), so they don't generate separate structs.
        if is_inline_attribute_ref(&def.name, &def.pattern)
            || is_simple_type(&def.pattern)
            || def.name.contains("_EG_")
            || def.name.contains("_AG_")
            || matches!(&def.pattern, Pattern::Element { .. })
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
        // Use mapped name for feature lookup (YAML uses Worksheet, not CT_Worksheet)
        let mapped_name = get_mapped_name(config, spec_name);
        let fields = collect_fields(&def.pattern, &definitions);
        for field in fields {
            report.total_fields += 1;
            if !has_field_mapping(config, &mapped_name, &field) {
                report
                    .unmapped_fields
                    .push(format!("{}.{}", mapped_name, field));
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

/// Get the mapped name for a type (e.g., CT_Worksheet -> Worksheet).
fn get_mapped_name(config: &CodegenConfig, spec_name: &str) -> String {
    if let Some(ref mappings) = config.name_mappings {
        let module_mappings = mappings.for_module(&config.module_name);
        // Check module-specific mappings first
        if let Some(mapped) = module_mappings.types.get(spec_name) {
            return mapped.clone();
        }
        // Check shared mappings
        if let Some(mapped) = mappings.shared.types.get(spec_name) {
            return mapped.clone();
        }
    }
    // No mapping - apply PascalCase like the codegen does
    to_pascal_case(spec_name)
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
fn collect_fields(pattern: &Pattern, definitions: &HashMap<&str, &Pattern>) -> Vec<String> {
    let mut fields = Vec::new();
    collect_fields_recursive(pattern, definitions, &mut fields, &mut HashSet::new());
    fields
}

fn collect_fields_recursive(
    pattern: &Pattern,
    definitions: &HashMap<&str, &Pattern>,
    fields: &mut Vec<String>,
    visited: &mut HashSet<String>,
) {
    match pattern {
        Pattern::Group(inner)
        | Pattern::Optional(inner)
        | Pattern::ZeroOrMore(inner)
        | Pattern::OneOrMore(inner)
        | Pattern::Mixed(inner) => {
            collect_fields_recursive(inner, definitions, fields, visited);
        }
        Pattern::Interleave(parts) | Pattern::Choice(parts) | Pattern::Sequence(parts) => {
            for part in parts {
                collect_fields_recursive(part, definitions, fields, visited);
            }
        }
        Pattern::Attribute { name, .. } => {
            // Use original XML name (camelCase) to match ooxml-features.yaml keys
            let field_name = name.local.clone();
            if !fields.contains(&field_name) {
                fields.push(field_name);
            }
        }
        Pattern::Element { name, .. } => {
            // Use original XML name (camelCase) to match ooxml-features.yaml keys
            let field_name = name.local.clone();
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
                    collect_fields_recursive(ref_pattern, definitions, fields, visited);
                } else if name.contains("_EG_") {
                    // Element groups are always included (no feature gating) — skip
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

#[cfg(test)]
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
