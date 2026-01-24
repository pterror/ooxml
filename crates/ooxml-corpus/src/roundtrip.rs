//! Roundtrip testing infrastructure for codegen migration.
//!
//! This module provides utilities to verify that parsing and re-serializing
//! OOXML documents produces semantically equivalent output. This is critical
//! for ensuring codegen doesn't introduce regressions when replacing handwritten types.
//!
//! # Approach
//!
//! For each document part (worksheet, document, slide), we:
//! 1. Parse the original XML using generated types
//! 2. Serialize back to XML
//! 3. Compare original vs roundtripped XML
//!
//! Comparison can be strict (byte-for-byte) or semantic (structure-aware).

use quick_xml::Reader;
use quick_xml::events::Event;
use std::collections::HashMap;
use std::io::Cursor;

/// Result of a roundtrip test.
#[derive(Debug, Clone)]
pub struct RoundtripResult {
    /// Whether the roundtrip was successful.
    pub success: bool,
    /// Part name being tested (e.g., "xl/worksheets/sheet1.xml").
    pub part_name: String,
    /// Original XML size in bytes.
    pub original_size: usize,
    /// Roundtripped XML size in bytes.
    pub roundtrip_size: usize,
    /// Differences found (if any).
    pub differences: Vec<XmlDifference>,
    /// Parse errors (if any).
    pub parse_error: Option<String>,
    /// Serialize errors (if any).
    pub serialize_error: Option<String>,
}

impl RoundtripResult {
    /// Create a successful result.
    pub fn ok(part_name: impl Into<String>, original_size: usize, roundtrip_size: usize) -> Self {
        Self {
            success: true,
            part_name: part_name.into(),
            original_size,
            roundtrip_size,
            differences: Vec::new(),
            parse_error: None,
            serialize_error: None,
        }
    }

    /// Create a failed result with differences.
    pub fn diff(
        part_name: impl Into<String>,
        original_size: usize,
        roundtrip_size: usize,
        differences: Vec<XmlDifference>,
    ) -> Self {
        Self {
            success: false,
            part_name: part_name.into(),
            original_size,
            roundtrip_size,
            differences,
            parse_error: None,
            serialize_error: None,
        }
    }

    /// Create a result for a parse error.
    pub fn parse_err(part_name: impl Into<String>, error: impl ToString) -> Self {
        Self {
            success: false,
            part_name: part_name.into(),
            original_size: 0,
            roundtrip_size: 0,
            differences: Vec::new(),
            parse_error: Some(error.to_string()),
            serialize_error: None,
        }
    }

    /// Create a result for a serialize error.
    pub fn serialize_err(
        part_name: impl Into<String>,
        original_size: usize,
        error: impl ToString,
    ) -> Self {
        Self {
            success: false,
            part_name: part_name.into(),
            original_size,
            roundtrip_size: 0,
            differences: Vec::new(),
            parse_error: None,
            serialize_error: Some(error.to_string()),
        }
    }
}

/// A difference found during XML comparison.
#[derive(Debug, Clone)]
pub struct XmlDifference {
    /// Path to the element (e.g., "/worksheet/sheetData/row[1]/c[2]").
    pub path: String,
    /// Type of difference.
    pub kind: DifferenceKind,
    /// Description of the difference.
    pub description: String,
}

/// Types of XML differences.
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum DifferenceKind {
    /// Element missing in roundtripped output.
    MissingElement,
    /// Extra element in roundtripped output.
    ExtraElement,
    /// Attribute value changed.
    AttributeChanged,
    /// Attribute missing in roundtripped output.
    MissingAttribute,
    /// Extra attribute in roundtripped output.
    ExtraAttribute,
    /// Text content changed.
    TextChanged,
    /// Element order changed.
    OrderChanged,
}

/// Options for XML comparison.
#[derive(Debug, Clone, Default)]
pub struct CompareOptions {
    /// Ignore whitespace-only text nodes.
    pub ignore_whitespace: bool,
    /// Ignore attribute order within elements.
    pub ignore_attribute_order: bool,
    /// Ignore element order (treat as set comparison).
    pub ignore_element_order: bool,
    /// Attributes to ignore (e.g., ["xmlns", "mc:Ignorable"]).
    pub ignore_attributes: Vec<String>,
    /// Elements to ignore (e.g., ["extLst"]).
    pub ignore_elements: Vec<String>,
}

impl CompareOptions {
    /// Strict comparison - byte-for-byte equivalent after normalization.
    pub fn strict() -> Self {
        Self::default()
    }

    /// Lenient comparison - ignores common variations.
    pub fn lenient() -> Self {
        Self {
            ignore_whitespace: true,
            ignore_attribute_order: true,
            ignore_element_order: false,
            ignore_attributes: vec!["xmlns".to_string()],
            ignore_elements: vec![],
        }
    }

    /// OOXML-specific comparison - ignores extension lists and namespaces.
    pub fn ooxml() -> Self {
        Self {
            ignore_whitespace: true,
            ignore_attribute_order: true,
            ignore_element_order: false,
            ignore_attributes: vec![
                "xmlns".to_string(),
                "xmlns:r".to_string(),
                "xmlns:mc".to_string(),
                "mc:Ignorable".to_string(),
            ],
            ignore_elements: vec!["extLst".to_string()],
        }
    }
}

/// Compare two XML documents and return differences.
pub fn compare_xml(
    original: &[u8],
    roundtripped: &[u8],
    options: &CompareOptions,
) -> Vec<XmlDifference> {
    let original_tree = match parse_xml_to_tree(original) {
        Ok(t) => t,
        Err(e) => {
            return vec![XmlDifference {
                path: "/".to_string(),
                kind: DifferenceKind::MissingElement,
                description: format!("Failed to parse original XML: {}", e),
            }];
        }
    };

    let roundtrip_tree = match parse_xml_to_tree(roundtripped) {
        Ok(t) => t,
        Err(e) => {
            return vec![XmlDifference {
                path: "/".to_string(),
                kind: DifferenceKind::MissingElement,
                description: format!("Failed to parse roundtripped XML: {}", e),
            }];
        }
    };

    let mut differences = Vec::new();
    compare_nodes(
        &original_tree,
        &roundtrip_tree,
        "",
        options,
        &mut differences,
    );
    differences
}

/// Simplified XML tree for comparison.
#[derive(Debug, Clone)]
enum XmlNode {
    Element {
        name: String,
        attributes: HashMap<String, String>,
        children: Vec<XmlNode>,
    },
    Text(String),
}

/// Info stored while parsing an element with children.
struct ElementBuilder {
    name: String,
    attributes: HashMap<String, String>,
}

fn parse_xml_to_tree(xml: &[u8]) -> Result<Vec<XmlNode>, String> {
    let mut reader = Reader::from_reader(Cursor::new(xml));
    let mut buf = Vec::new();
    // Stack of (element info, children being collected)
    let mut stack: Vec<(Option<ElementBuilder>, Vec<XmlNode>)> = vec![(None, Vec::new())];

    loop {
        buf.clear();
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) => {
                let name = String::from_utf8_lossy(e.name().as_ref()).into_owned();
                let attributes = e
                    .attributes()
                    .filter_map(|a| a.ok())
                    .map(|a| {
                        (
                            String::from_utf8_lossy(a.key.as_ref()).into_owned(),
                            String::from_utf8_lossy(&a.value).into_owned(),
                        )
                    })
                    .collect();
                // Push a new level for this element's children
                stack.push((Some(ElementBuilder { name, attributes }), Vec::new()));
            }
            Ok(Event::Empty(e)) => {
                let name = String::from_utf8_lossy(e.name().as_ref()).into_owned();
                let attributes = e
                    .attributes()
                    .filter_map(|a| a.ok())
                    .map(|a| {
                        (
                            String::from_utf8_lossy(a.key.as_ref()).into_owned(),
                            String::from_utf8_lossy(&a.value).into_owned(),
                        )
                    })
                    .collect();
                let elem = XmlNode::Element {
                    name,
                    attributes,
                    children: Vec::new(),
                };
                if let Some((_, children)) = stack.last_mut() {
                    children.push(elem);
                }
            }
            Ok(Event::End(_)) => {
                if stack.len() > 1 {
                    let (builder, children) = stack.pop().unwrap();
                    if let Some(ElementBuilder { name, attributes }) = builder {
                        let elem = XmlNode::Element {
                            name,
                            attributes,
                            children,
                        };
                        if let Some((_, parent_children)) = stack.last_mut() {
                            parent_children.push(elem);
                        }
                    }
                }
            }
            Ok(Event::Text(e)) => {
                let text = e.decode().unwrap_or_default().trim().to_string();
                if !text.is_empty()
                    && let Some((_, children)) = stack.last_mut()
                {
                    children.push(XmlNode::Text(text));
                }
            }
            Ok(Event::Eof) => break,
            Err(e) => return Err(format!("XML parse error: {}", e)),
            _ => {}
        }
    }

    Ok(stack.into_iter().next().map(|(_, c)| c).unwrap_or_default())
}

fn compare_nodes(
    original: &[XmlNode],
    roundtripped: &[XmlNode],
    path: &str,
    options: &CompareOptions,
    differences: &mut Vec<XmlDifference>,
) {
    // Filter out ignored elements
    let orig_filtered: Vec<_> = original
        .iter()
        .filter(|n| {
            if let XmlNode::Element { name, .. } = n {
                !options.ignore_elements.contains(name)
            } else if let XmlNode::Text(t) = n {
                !options.ignore_whitespace || !t.trim().is_empty()
            } else {
                true
            }
        })
        .collect();

    let rt_filtered: Vec<_> = roundtripped
        .iter()
        .filter(|n| {
            if let XmlNode::Element { name, .. } = n {
                !options.ignore_elements.contains(name)
            } else if let XmlNode::Text(t) = n {
                !options.ignore_whitespace || !t.trim().is_empty()
            } else {
                true
            }
        })
        .collect();

    // Compare by position (or as sets if ignore_element_order)
    if options.ignore_element_order {
        compare_as_sets(&orig_filtered, &rt_filtered, path, options, differences);
    } else {
        compare_by_position(&orig_filtered, &rt_filtered, path, options, differences);
    }
}

fn compare_by_position(
    original: &[&XmlNode],
    roundtripped: &[&XmlNode],
    path: &str,
    options: &CompareOptions,
    differences: &mut Vec<XmlDifference>,
) {
    let max_len = original.len().max(roundtripped.len());

    for i in 0..max_len {
        let orig = original.get(i);
        let rt = roundtripped.get(i);

        match (orig, rt) {
            (Some(o), Some(r)) => {
                compare_single_node(o, r, path, i, options, differences);
            }
            (Some(o), None) => {
                let name = node_name(o);
                differences.push(XmlDifference {
                    path: format!("{}/{}[{}]", path, name, i + 1),
                    kind: DifferenceKind::MissingElement,
                    description: format!("Element '{}' missing in roundtripped output", name),
                });
            }
            (None, Some(r)) => {
                let name = node_name(r);
                differences.push(XmlDifference {
                    path: format!("{}/{}[{}]", path, name, i + 1),
                    kind: DifferenceKind::ExtraElement,
                    description: format!("Extra element '{}' in roundtripped output", name),
                });
            }
            (None, None) => {}
        }
    }
}

fn compare_as_sets(
    original: &[&XmlNode],
    roundtripped: &[&XmlNode],
    path: &str,
    options: &CompareOptions,
    differences: &mut Vec<XmlDifference>,
) {
    // Simple set comparison - check each original element exists in roundtripped
    for (i, o) in original.iter().enumerate() {
        let name = node_name(o);
        let matching = roundtripped.iter().position(|r| nodes_match(o, r, options));

        if matching.is_none() {
            differences.push(XmlDifference {
                path: format!("{}/{}[{}]", path, name, i + 1),
                kind: DifferenceKind::MissingElement,
                description: format!("Element '{}' not found in roundtripped output", name),
            });
        }
    }

    // Check for extra elements
    for (i, r) in roundtripped.iter().enumerate() {
        let name = node_name(r);
        let matching = original.iter().position(|o| nodes_match(o, r, options));

        if matching.is_none() {
            differences.push(XmlDifference {
                path: format!("{}/{}[{}]", path, name, i + 1),
                kind: DifferenceKind::ExtraElement,
                description: format!("Extra element '{}' in roundtripped output", name),
            });
        }
    }
}

fn compare_single_node(
    original: &XmlNode,
    roundtripped: &XmlNode,
    path: &str,
    index: usize,
    options: &CompareOptions,
    differences: &mut Vec<XmlDifference>,
) {
    match (original, roundtripped) {
        (
            XmlNode::Element {
                name: n1,
                attributes: a1,
                children: c1,
            },
            XmlNode::Element {
                name: n2,
                attributes: a2,
                children: c2,
            },
        ) => {
            let elem_path = format!("{}/{}[{}]", path, n1, index + 1);

            if n1 != n2 {
                differences.push(XmlDifference {
                    path: elem_path.clone(),
                    kind: DifferenceKind::MissingElement,
                    description: format!("Element name mismatch: '{}' vs '{}'", n1, n2),
                });
                return;
            }

            // Compare attributes
            compare_attributes(a1, a2, &elem_path, options, differences);

            // Recursively compare children
            compare_nodes(c1, c2, &elem_path, options, differences);
        }
        (XmlNode::Text(t1), XmlNode::Text(t2)) => {
            if t1 != t2 {
                differences.push(XmlDifference {
                    path: format!("{}/#text[{}]", path, index + 1),
                    kind: DifferenceKind::TextChanged,
                    description: format!(
                        "Text changed: '{}' vs '{}'",
                        truncate(t1, 50),
                        truncate(t2, 50)
                    ),
                });
            }
        }
        (XmlNode::Element { name, .. }, XmlNode::Text(_)) => {
            differences.push(XmlDifference {
                path: format!("{}/[{}]", path, index + 1),
                kind: DifferenceKind::MissingElement,
                description: format!("Expected element '{}', got text", name),
            });
        }
        (XmlNode::Text(_), XmlNode::Element { name, .. }) => {
            differences.push(XmlDifference {
                path: format!("{}/[{}]", path, index + 1),
                kind: DifferenceKind::ExtraElement,
                description: format!("Expected text, got element '{}'", name),
            });
        }
    }
}

fn compare_attributes(
    original: &HashMap<String, String>,
    roundtripped: &HashMap<String, String>,
    path: &str,
    options: &CompareOptions,
    differences: &mut Vec<XmlDifference>,
) {
    // Check for missing/changed attributes
    for (k, v1) in original {
        // Skip ignored attributes
        if options.ignore_attributes.iter().any(|i| k.starts_with(i)) {
            continue;
        }

        match roundtripped.get(k) {
            Some(v2) if v1 != v2 => {
                differences.push(XmlDifference {
                    path: format!("{}/@{}", path, k),
                    kind: DifferenceKind::AttributeChanged,
                    description: format!("Attribute changed: '{}' vs '{}'", v1, v2),
                });
            }
            None => {
                differences.push(XmlDifference {
                    path: format!("{}/@{}", path, k),
                    kind: DifferenceKind::MissingAttribute,
                    description: format!("Attribute '{}' missing (was '{}')", k, v1),
                });
            }
            _ => {}
        }
    }

    // Check for extra attributes
    for (k, v) in roundtripped {
        // Skip ignored attributes
        if options.ignore_attributes.iter().any(|i| k.starts_with(i)) {
            continue;
        }

        if !original.contains_key(k) {
            differences.push(XmlDifference {
                path: format!("{}/@{}", path, k),
                kind: DifferenceKind::ExtraAttribute,
                description: format!("Extra attribute '{}' = '{}'", k, v),
            });
        }
    }
}

fn node_name(node: &XmlNode) -> &str {
    match node {
        XmlNode::Element { name, .. } => name,
        XmlNode::Text(_) => "#text",
    }
}

fn nodes_match(a: &XmlNode, b: &XmlNode, _options: &CompareOptions) -> bool {
    match (a, b) {
        (
            XmlNode::Element {
                name: n1,
                attributes: a1,
                ..
            },
            XmlNode::Element {
                name: n2,
                attributes: a2,
                ..
            },
        ) => n1 == n2 && a1 == a2,
        (XmlNode::Text(t1), XmlNode::Text(t2)) => t1 == t2,
        _ => false,
    }
}

fn truncate(s: &str, max_len: usize) -> String {
    if s.len() <= max_len {
        s.to_string()
    } else {
        format!("{}...", &s[..max_len])
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_compare_identical_xml() {
        let xml = br#"<root><child attr="value">text</child></root>"#;
        let diffs = compare_xml(xml, xml, &CompareOptions::strict());
        assert!(diffs.is_empty());
    }

    #[test]
    fn test_compare_missing_element() {
        let original = br#"<root><a/><b/></root>"#;
        let roundtrip = br#"<root><a/></root>"#;
        let diffs = compare_xml(original, roundtrip, &CompareOptions::strict());
        assert_eq!(diffs.len(), 1);
        assert_eq!(diffs[0].kind, DifferenceKind::MissingElement);
    }

    #[test]
    fn test_compare_extra_element() {
        let original = br#"<root><a/></root>"#;
        let roundtrip = br#"<root><a/><b/></root>"#;
        let diffs = compare_xml(original, roundtrip, &CompareOptions::strict());
        assert_eq!(diffs.len(), 1);
        assert_eq!(diffs[0].kind, DifferenceKind::ExtraElement);
    }

    #[test]
    fn test_compare_attribute_changed() {
        let original = br#"<root attr="old"/>"#;
        let roundtrip = br#"<root attr="new"/>"#;
        let diffs = compare_xml(original, roundtrip, &CompareOptions::strict());
        assert_eq!(diffs.len(), 1);
        assert_eq!(diffs[0].kind, DifferenceKind::AttributeChanged);
    }

    #[test]
    fn test_compare_ignore_whitespace() {
        let original = br#"<root>
            <child/>
        </root>"#;
        let roundtrip = br#"<root><child/></root>"#;
        let diffs = compare_xml(original, roundtrip, &CompareOptions::lenient());
        assert!(diffs.is_empty());
    }

    #[test]
    fn test_compare_ignore_xmlns() {
        let original = br#"<root xmlns="http://example.com"><child/></root>"#;
        let roundtrip = br#"<root><child/></root>"#;
        let diffs = compare_xml(original, roundtrip, &CompareOptions::lenient());
        assert!(diffs.is_empty());
    }
}
