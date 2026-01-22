//! AST types for RELAX NG Compact syntax (subset used by OOXML schemas).

/// A complete RNC schema file.
#[derive(Debug, Clone)]
pub struct Schema {
    pub namespaces: Vec<Namespace>,
    pub definitions: Vec<Definition>,
}

/// Namespace declaration: `namespace prefix = "uri"` or `default namespace prefix = "uri"`.
#[derive(Debug, Clone)]
pub struct Namespace {
    pub prefix: String,
    pub uri: String,
    pub is_default: bool,
}

/// A named definition: `name = pattern`.
#[derive(Debug, Clone)]
pub struct Definition {
    pub name: String,
    pub pattern: Pattern,
    pub doc_comment: Option<String>,
}

/// RELAX NG pattern (simplified for OOXML schemas).
#[derive(Debug, Clone)]
pub enum Pattern {
    /// Empty content.
    Empty,
    /// Reference to another definition.
    Ref(String),
    /// `element name { pattern }`.
    Element { name: QName, pattern: Box<Pattern> },
    /// `attribute name { pattern }`.
    Attribute { name: QName, pattern: Box<Pattern> },
    /// Sequence: `pattern, pattern`.
    Sequence(Vec<Pattern>),
    /// Choice: `pattern | pattern`.
    Choice(Vec<Pattern>),
    /// Interleave (unordered): `pattern & pattern`.
    Interleave(Vec<Pattern>),
    /// Optional: `pattern?`.
    Optional(Box<Pattern>),
    /// Zero or more: `pattern*`.
    ZeroOrMore(Box<Pattern>),
    /// One or more: `pattern+`.
    OneOrMore(Box<Pattern>),
    /// String literal: `string "value"`.
    StringLiteral(String),
    /// XSD datatype: `xsd:integer`, `xsd:string`, etc.
    Datatype {
        library: String,
        name: String,
        params: Vec<DatatypeParam>,
    },
    /// Parenthesized pattern.
    Group(Box<Pattern>),
    /// Mixed content: `mixed { pattern }`.
    Mixed(Box<Pattern>),
    /// List content: `list { pattern }` (space-separated).
    List(Box<Pattern>),
    /// Text content.
    Text,
    /// Any content (wildcard pattern).
    Any,
}

/// Qualified name with optional prefix.
#[derive(Debug, Clone)]
pub struct QName {
    pub prefix: Option<String>,
    pub local: String,
}

/// Name class for element/attribute names (supports wildcards and exclusions).
#[derive(Debug, Clone)]
pub enum NameClass {
    /// A specific qualified name.
    Name(QName),
    /// Wildcard: `*` (any name).
    AnyName,
    /// Namespace wildcard: `ns:*` (any name in namespace).
    NsName(String),
    /// Choice of name classes: `nc1 | nc2`.
    Choice(Vec<NameClass>),
    /// Subtraction: `nc1 - nc2` (nc1 except nc2).
    Except(Box<NameClass>, Box<NameClass>),
}

/// Datatype parameter: `{ length = "4" }`.
#[derive(Debug, Clone)]
pub struct DatatypeParam {
    pub name: String,
    pub value: String,
}

impl Pattern {
    /// Returns true if this pattern represents a simple type (enum of string literals).
    pub fn is_simple_type(&self) -> bool {
        match self {
            Pattern::Choice(variants) => variants
                .iter()
                .all(|v| matches!(v, Pattern::StringLiteral(_))),
            Pattern::StringLiteral(_) => true,
            Pattern::Datatype { .. } => true,
            _ => false,
        }
    }
}
