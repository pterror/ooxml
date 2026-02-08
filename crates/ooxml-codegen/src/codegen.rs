//! Rust code generator from parsed RNC schemas.

use crate::ast::{Definition, Pattern, QName, Schema};
use serde::Deserialize;
use std::collections::HashMap;
use std::fmt::Write;

/// Name mappings for a single module (sml, wml, pml, dml).
#[derive(Debug, Clone, Default, Deserialize)]
pub struct ModuleMappings {
    /// Type name mappings: `CT_AutoFilter` → `AutoFilter`
    #[serde(default)]
    pub types: HashMap<String, String>,
    /// Field name mappings: `r` → `reference`
    #[serde(default)]
    pub fields: HashMap<String, String>,
    /// Enum variant mappings: `customXml` → `CustomXmlContent`
    #[serde(default)]
    pub variants: HashMap<String, String>,
    /// XML element name mappings: `Worksheet` → `worksheet`
    /// Maps Rust type names to their XML element names for serde serialization.
    #[serde(default)]
    pub elements: HashMap<String, String>,
}

/// Complete name mappings file structure.
#[derive(Debug, Clone, Default, Deserialize)]
pub struct NameMappings {
    /// Shared mappings applied to all modules.
    #[serde(default)]
    pub shared: ModuleMappings,
    /// SpreadsheetML mappings.
    #[serde(default)]
    pub sml: ModuleMappings,
    /// WordprocessingML mappings.
    #[serde(default)]
    pub wml: ModuleMappings,
    /// PresentationML mappings.
    #[serde(default)]
    pub pml: ModuleMappings,
    /// DrawingML mappings.
    #[serde(default)]
    pub dml: ModuleMappings,
}

impl NameMappings {
    /// Load mappings from a YAML string.
    pub fn from_yaml(yaml: &str) -> Result<Self, serde_yaml::Error> {
        serde_yaml::from_str(yaml)
    }

    /// Load mappings from a YAML file.
    pub fn from_yaml_file(path: &std::path::Path) -> Result<Self, Box<dyn std::error::Error>> {
        let contents = std::fs::read_to_string(path)?;
        Ok(Self::from_yaml(&contents)?)
    }

    /// Get the module mappings for a given module name.
    pub fn for_module(&self, module: &str) -> &ModuleMappings {
        match module {
            "sml" => &self.sml,
            "wml" => &self.wml,
            "pml" => &self.pml,
            "dml" => &self.dml,
            _ => &self.shared,
        }
    }

    /// Resolve a type name, checking module-specific then shared mappings.
    pub fn resolve_type(&self, module: &str, spec_name: &str) -> Option<&str> {
        self.for_module(module)
            .types
            .get(spec_name)
            .or_else(|| self.shared.types.get(spec_name))
            .map(|s| s.as_str())
    }

    /// Resolve a field name, checking module-specific then shared mappings.
    pub fn resolve_field(&self, module: &str, spec_name: &str) -> Option<&str> {
        self.for_module(module)
            .fields
            .get(spec_name)
            .or_else(|| self.shared.fields.get(spec_name))
            .map(|s| s.as_str())
    }

    /// Resolve a variant name, checking module-specific then shared mappings.
    pub fn resolve_variant(&self, module: &str, spec_name: &str) -> Option<&str> {
        self.for_module(module)
            .variants
            .get(spec_name)
            .or_else(|| self.shared.variants.get(spec_name))
            .map(|s| s.as_str())
    }

    /// Resolve an XML element name for a Rust type name.
    /// Used for serde rename on structs.
    pub fn resolve_element(&self, module: &str, rust_type_name: &str) -> Option<&str> {
        self.for_module(module)
            .elements
            .get(rust_type_name)
            .or_else(|| self.shared.elements.get(rust_type_name))
            .map(|s| s.as_str())
    }
}

// =============================================================================
// Feature mappings for conditional compilation
// =============================================================================

/// Feature tags for a single element's attributes/children.
/// Maps attribute/child name -> list of feature tags.
pub type ElementFeatures = HashMap<String, Vec<String>>;

/// Feature mappings for a single module.
/// Maps element name (Rust name like "Row") -> attribute/child features.
pub type ModuleFeatures = HashMap<String, ElementFeatures>;

/// Complete feature mappings file structure.
#[derive(Debug, Clone, Default, Deserialize)]
pub struct FeatureMappings {
    /// SpreadsheetML feature mappings.
    #[serde(default)]
    pub sml: ModuleFeatures,
    /// WordprocessingML feature mappings.
    #[serde(default)]
    pub wml: ModuleFeatures,
    /// PresentationML feature mappings.
    #[serde(default)]
    pub pml: ModuleFeatures,
    /// DrawingML feature mappings.
    #[serde(default)]
    pub dml: ModuleFeatures,
}

impl FeatureMappings {
    /// Load mappings from a YAML string.
    pub fn from_yaml(yaml: &str) -> Result<Self, serde_yaml::Error> {
        serde_yaml::from_str(yaml)
    }

    /// Load mappings from a YAML file.
    pub fn from_yaml_file(path: &std::path::Path) -> Result<Self, Box<dyn std::error::Error>> {
        let contents = std::fs::read_to_string(path)?;
        Ok(Self::from_yaml(&contents)?)
    }

    /// Get the module features for a given module name.
    pub fn for_module(&self, module: &str) -> &ModuleFeatures {
        match module {
            "sml" => &self.sml,
            "wml" => &self.wml,
            "pml" => &self.pml,
            "dml" => &self.dml,
            _ => &self.sml, // Default to sml
        }
    }

    /// Get feature tags for a specific element's attribute/child.
    /// Returns None if no mapping exists (meaning it's always included).
    /// Supports `*` wildcard as a fallback for fields not explicitly listed.
    pub fn get_tags(&self, module: &str, element: &str, field: &str) -> Option<&[String]> {
        self.for_module(module)
            .get(element)
            .and_then(|elem| elem.get(field).or_else(|| elem.get("*")))
            .map(|v| v.as_slice())
    }

    /// Check if a field has the "core" tag (always included).
    pub fn is_core(&self, module: &str, element: &str, field: &str) -> bool {
        self.get_tags(module, element, field)
            .is_some_and(|tags| tags.iter().any(|t| t == "core"))
    }

    /// Get the primary feature name for a field (first non-core tag).
    /// Returns None if no feature gating needed (core or unmapped).
    pub fn primary_feature(&self, module: &str, element: &str, field: &str) -> Option<&str> {
        self.get_tags(module, element, field).and_then(|tags| {
            // If it's core, no feature gating needed
            if tags.iter().any(|t| t == "core") {
                return None;
            }
            // Return first tag as the primary feature
            tags.first().map(|s| s.as_str())
        })
    }
}

/// Code generation configuration.
#[derive(Debug, Clone, Default)]
pub struct CodegenConfig {
    /// Namespace prefix to strip from type names (e.g., "w_" for WordprocessingML).
    pub strip_prefix: Option<String>,
    /// Module name for the generated code (e.g., "sml", "wml").
    pub module_name: String,
    /// Optional name mappings for nicer Rust names.
    pub name_mappings: Option<NameMappings>,
    /// Optional feature mappings for conditional compilation.
    pub feature_mappings: Option<FeatureMappings>,
    /// Warn about types/fields without mappings (useful for finding unmapped items).
    pub warn_unmapped: bool,
    /// XML namespace prefix for serialized element/attribute names (e.g., "w" for WML).
    /// None = use unprefixed names (default namespace convention, used by SML/XLSX).
    /// Some("w") = use `w:` prefix (WML/DOCX), Some("p") = use `p:` prefix (PML/PPTX), etc.
    pub xml_serialize_prefix: Option<String>,
    /// Cross-crate imports for parser/serializer generation.
    /// Each entry is a full use path (e.g., "ooxml_dml::types::*" or "ooxml_dml::parsers::*").
    /// Used when types from another crate are referenced in this schema.
    pub cross_crate_imports: Vec<String>,
    /// Cross-crate type resolution for the type generator.
    /// Maps schema name prefixes (e.g., "a_") to (crate_path, module_name) tuples.
    /// Example: "a_" → ("ooxml_dml::types::", "dml")
    /// When a reference like "a_CT_Color" is not found locally, it's resolved using the
    /// module's name mappings (if available) or converted to PascalCase.
    pub cross_crate_type_prefix: HashMap<String, (String, String)>,
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
    /// Track generated Rust type names to avoid duplicates from merged schemas.
    generated_names: std::collections::HashSet<String>,
}

impl<'a> Generator<'a> {
    fn new(schema: &'a Schema, config: &'a CodegenConfig) -> Self {
        let definitions: HashMap<&str, &Pattern> = schema
            .definitions
            .iter()
            .map(|d| (d.name.as_str(), &d.pattern))
            .collect();

        Self {
            schema,
            config,
            output: String::new(),
            definitions,
            generated_names: std::collections::HashSet::new(),
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
            } else if self.is_inline_attribute_ref(&def.name, &def.pattern) {
                // Skip inline attribute references (like r_id = attribute r:id {...})
                // These are inlined into parent types via collect_fields
                continue;
            } else {
                complex_types.push(def);
            }
        }

        // Generate enums for simple types (string literal choices)
        for def in &simple_types {
            let rust_name = self.to_rust_type_name(&def.name);
            if !self.generated_names.insert(rust_name) {
                continue; // Skip duplicate Rust type names from merged schemas
            }
            if let Some(code) = self.gen_simple_type(def) {
                self.output.push_str(&code);
                self.output.push('\n');
            }
        }

        // Generate enums for element groups (element choice patterns)
        for def in &element_groups {
            let rust_name = self.to_rust_type_name(&def.name);
            if !self.generated_names.insert(rust_name) {
                continue;
            }
            if let Some(code) = self.gen_element_group(def) {
                self.output.push_str(&code);
                self.output.push('\n');
            }
        }

        // Generate structs for complex types
        for def in &complex_types {
            let rust_name = self.to_rust_type_name(&def.name);
            if !self.generated_names.insert(rust_name) {
                continue;
            }
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
                // Skip namespaces with empty prefix (default namespace without name)
                if ns.prefix.is_empty() {
                    continue;
                }
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
            Pattern::List(_) => true, // list { ... } is a simple type (space-separated string)
            Pattern::Ref(name) => {
                // Check if the referenced type is simple
                self.definitions
                    .get(name.as_str())
                    .is_some_and(|p| self.is_simple_type(p))
            }
            _ => false,
        }
    }

    /// Check if a definition is a pure attribute reference that should be inlined.
    /// These are attribute patterns (like r_id, r_embed) that don't have CT_ in their name.
    /// CT_* types with single attributes are element content types and need structs.
    fn is_inline_attribute_ref(&self, name: &str, pattern: &Pattern) -> bool {
        // Only skip non-CT types that are pure attribute patterns
        // CT_* types with single attributes are element content types
        !name.contains("_CT_") && matches!(pattern, Pattern::Attribute { .. })
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
            Pattern::Optional(inner) | Pattern::ZeroOrMore(inner) | Pattern::OneOrMore(inner) => {
                Self::is_direct_element_variant(inner)
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
                } else if let Some((cross_crate_type, _)) = self.resolve_cross_crate_type(target) {
                    // Resolved to a cross-crate type
                    cross_crate_type
                } else {
                    // Unknown type from another schema - use String as fallback
                    "String".to_string()
                };
                let mut code = String::new();
                writeln!(code, "pub type {} = {};", rust_name, target_rust).unwrap();
                Some(code)
            }
            Pattern::List(_) => {
                // List patterns (space-separated values) become String type aliases
                let mut code = String::new();
                writeln!(code, "pub type {} = String;", rust_name).unwrap();
                Some(code)
            }
            _ => None,
        }
    }

    fn gen_element_group(&self, def: &Definition) -> Option<String> {
        let rust_name = self.to_rust_type_name(&def.name);

        // Recursively collect all element variants, flattening nested EG_* refs
        let mut element_variants = Vec::new();
        let mut visited = std::collections::HashSet::new();
        visited.insert(def.name.clone()); // prevent self-recursion
        self.collect_element_variants(&def.pattern, &mut element_variants, &mut visited);

        // Deduplicate by xml_name (keep first occurrence)
        let mut seen = std::collections::HashSet::new();
        element_variants.retain(|(xml_name, _)| seen.insert(xml_name.clone()));

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

    /// Recursively collect all element variants from a pattern, following EG_* refs.
    /// Used for generating flattened element group enums.
    fn collect_element_variants(
        &self,
        pattern: &Pattern,
        variants: &mut Vec<(String, String)>,
        visited: &mut std::collections::HashSet<String>,
    ) {
        match pattern {
            Pattern::Element { name, pattern } => {
                // Enum variants are stored directly, not in Vec, so may need boxing
                let inner_type = self.pattern_to_rust_type(pattern, false, false);
                variants.push((name.local.clone(), inner_type));
            }
            Pattern::Optional(inner)
            | Pattern::ZeroOrMore(inner)
            | Pattern::OneOrMore(inner)
            | Pattern::Group(inner) => {
                self.collect_element_variants(inner, variants, visited);
            }
            Pattern::Ref(name) => {
                // Follow EG_* refs recursively to flatten nested element groups
                if name.contains("_EG_")
                    && visited.insert(name.clone())
                    && let Some(def_pattern) = self.definitions.get(name.as_str())
                {
                    self.collect_element_variants(def_pattern, variants, visited);
                }
            }
            Pattern::Choice(items) | Pattern::Sequence(items) | Pattern::Interleave(items) => {
                for item in items {
                    self.collect_element_variants(item, variants, visited);
                }
            }
            _ => {}
        }
    }

    /// Check if a pattern contains XML child elements (even from unresolved refs).
    /// Used to determine if empty-field structs should get extra_children.
    fn has_xml_children_pattern(&self, pattern: &Pattern) -> bool {
        match pattern {
            Pattern::Empty => false,
            Pattern::Attribute { .. } => false,
            Pattern::Element { .. } => true,
            Pattern::Ref(name) => {
                // Attribute groups don't produce children
                if name.contains("_AG_") {
                    return false;
                }
                // If it resolves to a known definition, check recursively
                if let Some(def_pattern) = self.definitions.get(name.as_str()) {
                    self.has_xml_children_pattern(def_pattern)
                } else {
                    // Unresolved ref — assume it produces children
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

    /// Check if a pattern contains XML attributes (even from unresolved refs).
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

    /// Check if a field holds element group content (needs serde skip).
    fn is_eg_content_field(&self, field: &Field) -> bool {
        if let Pattern::Ref(name) = &field.pattern
            && name.contains("_EG_")
            && let Some(pattern) = self.definitions.get(name.as_str())
        {
            return self.is_element_choice(pattern);
        }
        false
    }

    /// Convert an EG_* reference name to a snake_case field name.
    /// e.g., "w_EG_BlockLevelElts" → "block_content" (via names.yaml)
    fn eg_ref_to_field_name(&self, name: &str) -> String {
        let spec_name = strip_namespace_prefix(name);
        // Strip "EG_" prefix
        let short = spec_name.strip_prefix("EG_").unwrap_or(spec_name);
        // Check names.yaml field mapping first
        if let Some(mappings) = &self.config.name_mappings
            && let Some(mapped) = mappings.resolve_field(&self.config.module_name, short)
        {
            return mapped.to_string();
        }
        to_snake_case(short)
    }

    fn gen_complex_type(&self, def: &Definition) -> Option<String> {
        // For element-only definitions, generate a type alias to the inner type
        if let Pattern::Element { pattern, .. } = &def.pattern {
            let rust_name = self.to_rust_type_name(&def.name);
            // Type aliases may be used in various contexts, use boxed form for recursive types
            let inner_type = self.pattern_to_rust_type(pattern, false, false);
            let mut code = String::new();
            writeln!(code, "pub type {} = {};", rust_name, inner_type).unwrap();
            return Some(code);
        }

        let rust_name = self.to_rust_type_name(&def.name);
        let mut code = String::new();

        // Collect fields from the pattern
        let fields = self.extract_fields(&def.pattern);

        // Check for XML element name mapping
        let element_rename = self.get_element_name(&rust_name);

        if fields.is_empty() {
            // Check if the pattern has XML content that we couldn't resolve into fields
            // (e.g. refs to types from other schemas). If so, generate extra_children
            // to preserve the content for roundtrip fidelity.
            let has_unresolved_children = self.has_xml_children_pattern(&def.pattern);
            let has_unresolved_attrs = self.has_xml_attr_pattern(&def.pattern);

            writeln!(
                code,
                "#[derive(Debug, Clone, Default, Serialize, Deserialize)]"
            )
            .unwrap();
            if let Some(xml_name) = &element_rename {
                writeln!(code, "#[serde(rename = \"{}\")]", xml_name).unwrap();
            }

            if has_unresolved_children || has_unresolved_attrs {
                writeln!(code, "pub struct {} {{", rust_name).unwrap();
                if has_unresolved_attrs {
                    writeln!(
                        code,
                        "    /// Unknown attributes captured for roundtrip fidelity."
                    )
                    .unwrap();
                    writeln!(code, "    #[cfg(feature = \"extra-attrs\")]").unwrap();
                    writeln!(code, "    #[serde(skip)]").unwrap();
                    writeln!(code, "    #[cfg(feature = \"extra-attrs\")]").unwrap();
                    writeln!(code, "    #[serde(default)]").unwrap();
                    writeln!(code, "    #[cfg(feature = \"extra-attrs\")]").unwrap();
                    writeln!(
                        code,
                        "    pub extra_attrs: std::collections::HashMap<String, String>,"
                    )
                    .unwrap();
                }
                if has_unresolved_children {
                    writeln!(
                        code,
                        "    /// Unknown child elements captured for roundtrip fidelity."
                    )
                    .unwrap();
                    writeln!(code, "    #[cfg(feature = \"extra-children\")]").unwrap();
                    writeln!(code, "    #[serde(skip)]").unwrap();
                    writeln!(code, "    #[cfg(feature = \"extra-children\")]").unwrap();
                    writeln!(
                        code,
                        "    pub extra_children: Vec<ooxml_xml::PositionedNode>,"
                    )
                    .unwrap();
                }
                writeln!(code, "}}").unwrap();
            } else {
                writeln!(code, "pub struct {};", rust_name).unwrap();
            }
        } else {
            // Derive Default when all fields are optional, vec, or EG content (which we
            // always wrap in Option<> or Vec<>, making them defaultable).
            let all_defaultable = fields
                .iter()
                .all(|f| f.is_optional || f.is_vec || self.is_eg_content_field(f));
            if all_defaultable {
                writeln!(
                    code,
                    "#[derive(Debug, Clone, Default, Serialize, Deserialize)]"
                )
                .unwrap();
            } else {
                writeln!(code, "#[derive(Debug, Clone, Serialize, Deserialize)]").unwrap();
            }
            if let Some(xml_name) = &element_rename {
                writeln!(code, "#[serde(rename = \"{}\")]", xml_name).unwrap();
            }
            writeln!(code, "pub struct {} {{", rust_name).unwrap();

            for field in &fields {
                let is_eg_content = self.is_eg_content_field(field);
                // Pass is_vec to avoid boxing in Vec contexts (Vec provides heap indirection)
                let inner_type = self.pattern_to_rust_type(&field.pattern, false, field.is_vec);
                let is_bool = inner_type == "bool";
                // Required EG content fields are wrapped in Option<> for Default/serde
                // compatibility — EG enums don't impl Default, but Option<Box<EG>> does.
                let eg_needs_option = is_eg_content && !field.is_optional && !field.is_vec;
                let field_type = if field.is_vec {
                    format!("Vec<{}>", inner_type)
                } else if field.is_optional || eg_needs_option {
                    format!("Option<{}>", inner_type)
                } else {
                    inner_type
                };

                // Add feature cfg attribute if not core
                if let Some(ref feature) = self.get_field_feature(&rust_name, &field.xml_name) {
                    writeln!(code, "    #[cfg(feature = \"{}\")]", feature).unwrap();
                }

                if is_eg_content {
                    // EG content fields use serde skip — populated by FromXml parsers
                    writeln!(code, "    #[serde(skip)]").unwrap();
                    writeln!(code, "    #[serde(default)]").unwrap();
                } else {
                    // Add serde attributes
                    let xml_name = &field.xml_name;
                    if field.is_text_content {
                        writeln!(code, "    #[serde(rename = \"$text\")]").unwrap();
                    } else if field.is_attribute {
                        // Include namespace prefix for attributes (e.g., r:id → @r:id)
                        if let Some(prefix) = &field.xml_prefix {
                            writeln!(code, "    #[serde(rename = \"@{}:{}\")]", prefix, xml_name)
                                .unwrap();
                        } else {
                            writeln!(code, "    #[serde(rename = \"@{}\")]", xml_name).unwrap();
                        }
                    } else {
                        writeln!(code, "    #[serde(rename = \"{}\")]", xml_name).unwrap();
                    }
                }
                if field.is_optional {
                    if is_bool {
                        // OOXML booleans serialize as "1"/"0", not "true"/"false"
                        writeln!(
                            code,
                            "    #[serde(default, skip_serializing_if = \"Option::is_none\", with = \"ooxml_xml::ooxml_bool\")]"
                        )
                        .unwrap();
                    } else if !is_eg_content {
                        writeln!(
                            code,
                            "    #[serde(default, skip_serializing_if = \"Option::is_none\")]"
                        )
                        .unwrap();
                    }
                } else if field.is_vec && !is_eg_content {
                    writeln!(
                        code,
                        "    #[serde(default, skip_serializing_if = \"Vec::is_empty\")]"
                    )
                    .unwrap();
                } else if is_bool {
                    // Required booleans also need OOXML format
                    writeln!(
                        code,
                        "    #[serde(with = \"ooxml_xml::ooxml_bool_required\")]"
                    )
                    .unwrap();
                }
                writeln!(code, "    pub {}: {},", field.name, field_type).unwrap();
            }

            // Add extra_attrs field to capture unknown attributes for roundtrip fidelity
            let has_attrs = fields.iter().any(|f| f.is_attribute);
            if has_attrs {
                writeln!(
                    code,
                    "    /// Unknown attributes captured for roundtrip fidelity."
                )
                .unwrap();
                writeln!(code, "    #[cfg(feature = \"extra-attrs\")]").unwrap();
                // Use skip instead of flatten - flatten doesn't work well with quick-xml
                writeln!(code, "    #[serde(skip)]").unwrap();
                writeln!(code, "    #[cfg(feature = \"extra-attrs\")]").unwrap();
                writeln!(code, "    #[serde(default)]").unwrap();
                writeln!(code, "    #[cfg(feature = \"extra-attrs\")]").unwrap();
                writeln!(
                    code,
                    "    pub extra_attrs: std::collections::HashMap<String, String>,"
                )
                .unwrap();
            }

            // Add extra_children field to capture unknown child elements for roundtrip fidelity
            // Include types with text content, as they have a parsing loop that might encounter unknown children
            let has_parsing_content = fields.iter().any(|f| !f.is_attribute);
            if has_parsing_content {
                writeln!(
                    code,
                    "    /// Unknown child elements captured for roundtrip fidelity."
                )
                .unwrap();
                writeln!(code, "    #[cfg(feature = \"extra-children\")]").unwrap();
                writeln!(code, "    #[serde(skip)]").unwrap();
                writeln!(code, "    #[cfg(feature = \"extra-children\")]").unwrap();
                writeln!(
                    code,
                    "    pub extra_children: Vec<ooxml_xml::PositionedNode>,"
                )
                .unwrap();
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
            Pattern::ZeroOrMore(inner) | Pattern::OneOrMore(inner) => {
                // These become Vec<T> fields
                match inner.as_ref() {
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
                                // Element group choice → add Vec<Box<EGType>> field
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
                                // Struct-like property group → inline its fields
                                self.collect_fields(def_pattern, fields, true);
                            }
                        }
                    }
                    Pattern::Choice(alternatives) => {
                        // ZeroOrMore(Choice([elements])) - each element can appear multiple times
                        // Make each alternative a Vec field instead of Option
                        for alt in alternatives {
                            self.collect_fields_as_vec(alt, fields);
                        }
                    }
                    Pattern::Ref(_) => {
                        // Complex repeated content - recurse but don't add directly
                        self.collect_fields(inner, fields, false);
                    }
                    Pattern::Group(group_inner) => {
                        // Unwrap group and handle inner pattern
                        // For ZeroOrMore(Group(Choice([...]))) - treat each alternative as Vec
                        if let Pattern::Choice(alternatives) = group_inner.as_ref() {
                            for alt in alternatives {
                                self.collect_fields_as_vec(alt, fields);
                            }
                        } else {
                            self.collect_fields(group_inner, fields, false);
                        }
                    }
                    _ => {}
                }
            }
            Pattern::Group(inner) => {
                self.collect_fields(inner, fields, is_optional);
            }
            Pattern::Ref(name) => {
                if let Some(def_pattern) = self.definitions.get(name.as_str()) {
                    if self.is_string_type(def_pattern) {
                        // Text content - always optional since elements
                        // can be self-closing (e.g. shared formula refs: <f t="shared" si="0"/>)
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
                    } else if name.contains("_EG_") {
                        // Element group reference
                        if self.is_element_choice(def_pattern) {
                            // Element group choice → add as content field
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
                            // Struct-like property group → inline its fields
                            self.collect_fields(def_pattern, fields, is_optional);
                        }
                    } else if name.contains("_AG_") {
                        // Attribute group reference → inline its attribute fields
                        self.collect_fields(def_pattern, fields, is_optional);
                    } else {
                        // CT_* mixin or base type — inline its fields
                        self.collect_fields(def_pattern, fields, is_optional);
                    }
                }
            }
            Pattern::Choice(alternatives) => {
                // Flatten choice into optional fields.
                // In a choice, each alternative might not be selected, so all become optional.
                // This handles patterns like (element a? | element b? | ...)+ in Font/Fill.
                for alt in alternatives {
                    self.collect_fields(alt, fields, true);
                }
            }
            _ => {}
        }
    }

    /// Collect fields as Vec (for elements inside ZeroOrMore(Choice([...])))
    /// Only non-optional elements become Vec; optional elements stay as Option.
    fn collect_fields_as_vec(&self, pattern: &Pattern, fields: &mut Vec<Field>) {
        match pattern {
            Pattern::Element {
                name,
                pattern: inner_pattern,
            } if name.local != "_any" => {
                fields.push(Field {
                    name: self.qname_to_field_name(name),
                    xml_name: name.local.clone(),
                    xml_prefix: name.prefix.clone(),
                    pattern: inner_pattern.as_ref().clone(),
                    is_optional: false,
                    is_attribute: false,
                    is_vec: true,
                    is_text_content: false,
                });
            }
            Pattern::Optional(inner) => {
                // Optional inside a repeating choice means element can appear 0-1 times,
                // NOT multiple times. Delegate to collect_fields with is_optional=true.
                self.collect_fields(inner, fields, true);
            }
            Pattern::Group(inner) => {
                self.collect_fields_as_vec(inner, fields);
            }
            Pattern::Ref(name) => {
                // Ref inside a repeating choice - create Vec<RefType> field
                if let Some(def_pattern) = self.definitions.get(name.as_str()) {
                    if name.contains("_EG_") && self.is_element_choice(def_pattern) {
                        // EG_* element group → Vec<EGType>
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
                    } else if !name.contains("_AG_") && !name.contains("_CT_") {
                        // Other refs that wrap elements
                        self.collect_fields_as_vec(def_pattern, fields);
                    }
                }
            }
            _ => {}
        }
    }

    fn to_rust_type_name(&self, name: &str) -> String {
        // Strip namespace prefix to get spec name (e.g., "sml_CT_AutoFilter" → "CT_AutoFilter")
        let spec_name = strip_namespace_prefix(name);

        // Check name mappings first
        if let Some(mappings) = &self.config.name_mappings
            && let Some(mapped) = mappings.resolve_type(&self.config.module_name, spec_name)
        {
            return mapped.to_string();
        }

        // Warn about unmapped types if enabled
        if self.config.warn_unmapped && self.config.name_mappings.is_some() {
            eprintln!("warning: unmapped type '{}' (spec: {})", spec_name, name);
        }

        // Fall back to PascalCase conversion
        to_pascal_case(spec_name)
    }

    fn to_rust_variant_name(&self, name: &str) -> String {
        // Handle empty string variant
        if name.is_empty() {
            return "Empty".to_string();
        }

        // Check name mappings first
        if let Some(mappings) = &self.config.name_mappings
            && let Some(mapped) = mappings.resolve_variant(&self.config.module_name, name)
        {
            return mapped.to_string();
        }

        // Fall back to PascalCase conversion
        let name = to_pascal_case(name);
        // Prefix with underscore if starts with digit
        if name.chars().next().is_some_and(|c| c.is_ascii_digit()) {
            format!("_{}", name)
        } else {
            name
        }
    }

    fn qname_to_field_name(&self, qname: &QName) -> String {
        // Check name mappings first
        if let Some(mappings) = &self.config.name_mappings
            && let Some(mapped) = mappings.resolve_field(&self.config.module_name, &qname.local)
        {
            return mapped.to_string();
        }

        // Warn about unmapped fields if enabled
        if self.config.warn_unmapped && self.config.name_mappings.is_some() {
            eprintln!("warning: unmapped field '{}'", qname.local);
        }

        // Fall back to snake_case conversion
        to_snake_case(&qname.local)
    }

    /// Get the XML element name for a Rust type, if mapped.
    fn get_element_name(&self, rust_type_name: &str) -> Option<String> {
        self.config
            .name_mappings
            .as_ref()
            .and_then(|m| m.resolve_element(&self.config.module_name, rust_type_name))
            .map(|s| s.to_string())
    }

    /// Get the feature name for a field if it requires feature gating.
    /// Returns None if the field is "core" (always included) or unmapped.
    /// Returns the feature name prefixed with the module name (e.g., "sml-styling").
    fn get_field_feature(&self, struct_name: &str, xml_field_name: &str) -> Option<String> {
        self.config
            .feature_mappings
            .as_ref()
            .and_then(|fm| {
                fm.primary_feature(&self.config.module_name, struct_name, xml_field_name)
            })
            .map(|feature| format!("{}-{}", self.config.module_name, feature))
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

    /// Try to resolve a schema reference name to a cross-crate Rust type.
    /// Returns Some((full_path, needs_box)) if the name matches a configured prefix.
    fn resolve_cross_crate_type(&self, name: &str) -> Option<(String, bool)> {
        for (prefix, (crate_path, module_name)) in &self.config.cross_crate_type_prefix {
            if name.starts_with(prefix) {
                // Convert schema name (a_CT_Color) to Rust type name using the cross-crate module's mappings
                let spec_name = strip_namespace_prefix(name);
                let rust_type_name = if let Some(mappings) = &self.config.name_mappings
                    && let Some(mapped) = mappings.resolve_type(module_name, spec_name)
                {
                    mapped.to_string()
                } else {
                    to_pascal_case(spec_name)
                };
                let full_path = format!("{}{}", crate_path, rust_type_name);
                // Box complex types (CT_*) and element groups (EG_*) to avoid infinite size
                let needs_box = name.contains("_CT_") || name.contains("_EG_");
                return Some((full_path, needs_box));
            }
        }
        None
    }

    /// Check if a definition is an element-wrapper type alias (Element { pattern: Ref(...) }).
    /// These generate `pub type Foo = Box<T>;` and should not be double-boxed when used.
    fn is_element_wrapper_type_alias(&self, name: &str) -> bool {
        if let Some(def_pattern) = self.definitions.get(name) {
            matches!(def_pattern, Pattern::Element { pattern, .. } if matches!(pattern.as_ref(), Pattern::Ref(_)))
        } else {
            false
        }
    }

    /// Convert a pattern to a Rust type string.
    /// - `is_optional`: whether to wrap in Option<>
    /// - `is_vec`: if true, don't box even if recursive (Vec provides heap indirection)
    fn pattern_to_rust_type(&self, pattern: &Pattern, is_optional: bool, is_vec: bool) -> String {
        let (inner, needs_box) = match pattern {
            Pattern::Ref(name) => {
                // Check if this is a known definition
                if self.definitions.contains_key(name.as_str()) {
                    let type_name = self.to_rust_type_name(name);
                    // Box CT_* and EG_* types, but not in Vec context (Vec provides heap indirection)
                    // Also don't box element-wrapper type aliases - they're already Box<T>
                    let is_complex = name.contains("_CT_") || name.contains("_EG_");
                    let is_already_boxed = self.is_element_wrapper_type_alias(name);
                    let needs_box = is_complex && !is_vec && !is_already_boxed;
                    (type_name, needs_box)
                } else if let Some((cross_crate_type, cross_crate_needs_box)) =
                    self.resolve_cross_crate_type(name)
                {
                    // Resolved to a cross-crate type - apply same Vec logic
                    let needs_box = !is_vec && cross_crate_needs_box;
                    (cross_crate_type, needs_box)
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
    #[allow(dead_code)]
    xml_prefix: Option<String>,
    pattern: Pattern,
    is_optional: bool,
    is_attribute: bool,
    is_vec: bool,
    is_text_content: bool,
}

/// Strip namespace prefix from a definition name.
/// Examples:
/// - `sml_CT_AutoFilter` → `CT_AutoFilter`
/// - `s_ST_Lang` → `ST_Lang`
/// - `w_EG_ContentRunContent` → `EG_ContentRunContent`
/// - `CT_Foo` → `CT_Foo` (no prefix)
fn strip_namespace_prefix(name: &str) -> &str {
    // Find the type kind prefix (CT_, ST_, EG_)
    for kind in ["CT_", "ST_", "EG_"] {
        if let Some(pos) = name.find(kind)
            && pos > 0
        {
            // There's a namespace prefix before the kind
            return &name[pos..];
        }
    }
    // No known type kind found, return as-is
    name
}

pub(crate) fn to_pascal_case(s: &str) -> String {
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
        "macro" => "r#macro".to_string(),
        "try" => "r#try".to_string(),
        "abstract" => "r#abstract".to_string(),
        "become" => "r#become".to_string(),
        "final" => "r#final".to_string(),
        "override" => "r#override".to_string(),
        "priv" => "r#priv".to_string(),
        "typeof" => "r#typeof".to_string(),
        "unsized" => "r#unsized".to_string(),
        "virtual" => "r#virtual".to_string(),
        "yield" => "r#yield".to_string(),
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
