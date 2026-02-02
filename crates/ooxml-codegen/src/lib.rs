//! Code generator for OOXML types from RELAX NG schemas.
//!
//! This crate parses RELAX NG Compact (.rnc) schema files from the ECMA-376
//! specification and generates Rust structs for Office Open XML types.

pub mod ast;
pub mod codegen;
pub mod lexer;
pub mod parser;
pub mod parser_gen;
pub mod serializer_gen;

pub use ast::{DatatypeParam, Definition, Namespace, Pattern, QName, Schema};
pub use codegen::{CodegenConfig, FeatureMappings, ModuleMappings, NameMappings, generate};
pub use lexer::{LexError, Lexer};
pub use parser::{ParseError, Parser};
pub use parser_gen::generate_parsers;
pub use serializer_gen::generate_serializers;

/// Parse an RNC schema from a string.
pub fn parse_rnc(input: &str) -> Result<Schema, Error> {
    let tokens = Lexer::new(input).tokenize()?;
    let schema = Parser::new(tokens).parse()?;
    Ok(schema)
}

/// Lex an RNC schema into tokens (for debugging).
pub fn lex_rnc(input: &str) -> Result<Vec<lexer::Token>, LexError> {
    Lexer::new(input).tokenize()
}

#[derive(Debug, thiserror::Error)]
pub enum Error {
    #[error("lexer error: {0}")]
    Lex(#[from] LexError),
    #[error("parser error: {0}")]
    Parse(#[from] ParseError),
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_parse_simple_schema() {
        let input = r#"
namespace w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

w_CT_Empty = empty
w_CT_OnOff = attribute w:val { s_ST_OnOff }?
w_ST_HighlightColor =
  string "black"
  | string "blue"
  | string "cyan"
"#;
        let schema = parse_rnc(input).unwrap();
        assert_eq!(schema.namespaces.len(), 1);
        assert_eq!(schema.definitions.len(), 3);
    }
}
