//! Parser for RELAX NG Compact syntax.

use crate::ast::{DatatypeParam, Definition, NameClass, Namespace, Pattern, QName, Schema};
use crate::lexer::Token;

/// Parser state.
pub struct Parser {
    tokens: Vec<Token>,
    pos: usize,
}

impl Parser {
    pub fn new(tokens: Vec<Token>) -> Self {
        Self { tokens, pos: 0 }
    }

    pub fn parse(mut self) -> Result<Schema, ParseError> {
        let mut namespaces = Vec::new();
        let mut definitions = Vec::new();

        while !self.at_end() {
            if self.check(&Token::Namespace) || self.check(&Token::Default) {
                namespaces.push(self.parse_namespace()?);
            } else if let Some(Token::Ident(_)) = self.peek() {
                definitions.push(self.parse_definition()?);
            } else {
                return Err(self.error("expected namespace or definition"));
            }
        }

        Ok(Schema {
            namespaces,
            definitions,
        })
    }

    fn parse_namespace(&mut self) -> Result<Namespace, ParseError> {
        let is_default = self.check(&Token::Default);
        if is_default {
            self.advance();
        }
        self.expect(&Token::Namespace)?;

        // Handle both forms:
        // - `namespace prefix = "uri"`
        // - `default namespace = "uri"` (no prefix for default namespace)
        let prefix = if self.check(&Token::Equals) {
            // No prefix - use empty string to indicate default namespace
            String::new()
        } else {
            self.expect_ident()?
        };
        self.expect(&Token::Equals)?;
        let uri = self.expect_string()?;

        Ok(Namespace {
            prefix,
            uri,
            is_default,
        })
    }

    fn parse_definition(&mut self) -> Result<Definition, ParseError> {
        let name = self.expect_ident()?;
        self.expect(&Token::Equals)?;
        let pattern = self.parse_pattern()?;

        Ok(Definition {
            name,
            pattern,
            doc_comment: None,
        })
    }

    fn parse_pattern(&mut self) -> Result<Pattern, ParseError> {
        self.parse_interleave()
    }

    // Interleave has lowest precedence: a & b & c
    fn parse_interleave(&mut self) -> Result<Pattern, ParseError> {
        let mut left = self.parse_choice()?;

        while self.check(&Token::Ampersand) {
            self.advance();
            let right = self.parse_choice()?;
            left = match left {
                Pattern::Interleave(mut v) => {
                    v.push(right);
                    Pattern::Interleave(v)
                }
                _ => Pattern::Interleave(vec![left, right]),
            };
        }

        Ok(left)
    }

    // Choice: a | b | c
    fn parse_choice(&mut self) -> Result<Pattern, ParseError> {
        let mut left = self.parse_sequence()?;

        while self.check(&Token::Pipe) {
            self.advance();
            let right = self.parse_sequence()?;
            left = match left {
                Pattern::Choice(mut v) => {
                    v.push(right);
                    Pattern::Choice(v)
                }
                _ => Pattern::Choice(vec![left, right]),
            };
        }

        Ok(left)
    }

    // Sequence: a, b, c
    fn parse_sequence(&mut self) -> Result<Pattern, ParseError> {
        let mut left = self.parse_postfix()?;

        while self.check(&Token::Comma) {
            self.advance();
            let right = self.parse_postfix()?;
            left = match left {
                Pattern::Sequence(mut v) => {
                    v.push(right);
                    Pattern::Sequence(v)
                }
                _ => Pattern::Sequence(vec![left, right]),
            };
        }

        Ok(left)
    }

    // Postfix operators: ?, *, +
    fn parse_postfix(&mut self) -> Result<Pattern, ParseError> {
        let mut pattern = self.parse_primary()?;

        loop {
            if self.check(&Token::Question) {
                self.advance();
                pattern = Pattern::Optional(Box::new(pattern));
            } else if self.check(&Token::Star) {
                self.advance();
                pattern = Pattern::ZeroOrMore(Box::new(pattern));
            } else if self.check(&Token::Plus) {
                self.advance();
                pattern = Pattern::OneOrMore(Box::new(pattern));
            } else {
                break;
            }
        }

        Ok(pattern)
    }

    // Primary patterns
    fn parse_primary(&mut self) -> Result<Pattern, ParseError> {
        if self.check(&Token::Empty) {
            self.advance();
            return Ok(Pattern::Empty);
        }

        if self.check(&Token::String) {
            self.advance();
            let value = self.expect_string()?;
            return Ok(Pattern::StringLiteral(value));
        }

        if self.check(&Token::Element) {
            return self.parse_element();
        }

        if self.check(&Token::Attribute) {
            return self.parse_attribute();
        }

        if self.check(&Token::Mixed) {
            self.advance();
            self.expect(&Token::LBrace)?;
            let inner = self.parse_pattern()?;
            self.expect(&Token::RBrace)?;
            return Ok(Pattern::Mixed(Box::new(inner)));
        }

        if self.check(&Token::List) {
            self.advance();
            self.expect(&Token::LBrace)?;
            let inner = self.parse_pattern()?;
            self.expect(&Token::RBrace)?;
            return Ok(Pattern::List(Box::new(inner)));
        }

        if self.check(&Token::Text) {
            self.advance();
            return Ok(Pattern::Text);
        }

        // Bare string literal (RELAX NG value pattern): "literal"
        if let Some(Token::QuotedString(_)) = self.peek() {
            let value = self.expect_string()?;
            return Ok(Pattern::StringLiteral(value));
        }

        if self.check(&Token::LParen) {
            self.advance();
            let inner = self.parse_pattern()?;
            self.expect(&Token::RParen)?;
            return Ok(Pattern::Group(Box::new(inner)));
        }

        // Identifier - could be a reference or datatype
        if let Some(Token::Ident(_)) = self.peek() {
            let name = self.expect_ident()?;

            // Check for datatype with colon (e.g., xsd:integer, xsd:string)
            if self.check(&Token::Colon) {
                self.advance();
                let type_name = self.expect_ident_or_keyword()?;

                // Check for params { ... } or literal value "..."
                if self.check(&Token::LBrace) {
                    let params = self.parse_datatype_params()?;
                    return Ok(Pattern::Datatype {
                        library: name,
                        name: type_name,
                        params,
                    });
                } else if let Some(Token::QuotedString(_)) = self.peek() {
                    // Datatype with literal value pattern: xsd:int "255"
                    let value = self.expect_string()?;
                    return Ok(Pattern::Datatype {
                        library: name,
                        name: type_name,
                        params: vec![DatatypeParam {
                            name: "pattern".to_string(),
                            value,
                        }],
                    });
                } else {
                    return Ok(Pattern::Datatype {
                        library: name,
                        name: type_name,
                        params: vec![],
                    });
                }
            }

            return Ok(Pattern::Ref(name));
        }

        Err(self.error("expected pattern"))
    }

    fn parse_element(&mut self) -> Result<Pattern, ParseError> {
        self.expect(&Token::Element)?;
        let name_class = self.parse_name_class()?;
        self.expect(&Token::LBrace)?;
        let pattern = self.parse_pattern()?;
        self.expect(&Token::RBrace)?;

        // Convert NameClass to QName for simple cases, use placeholder for wildcards
        let name = match name_class {
            NameClass::Name(qn) => qn,
            _ => QName {
                prefix: None,
                local: "_any".to_string(),
            },
        };

        Ok(Pattern::Element {
            name,
            pattern: Box::new(pattern),
        })
    }

    fn parse_attribute(&mut self) -> Result<Pattern, ParseError> {
        self.expect(&Token::Attribute)?;
        let name_class = self.parse_name_class()?;
        self.expect(&Token::LBrace)?;
        let pattern = self.parse_pattern()?;
        self.expect(&Token::RBrace)?;

        // Convert NameClass to QName for simple cases, use placeholder for wildcards
        let name = match name_class {
            NameClass::Name(qn) => qn,
            _ => QName {
                prefix: None,
                local: "_any".to_string(),
            },
        };

        Ok(Pattern::Attribute {
            name,
            pattern: Box::new(pattern),
        })
    }

    /// Parse a name class (handles wildcards, namespace wildcards, and exclusions).
    fn parse_name_class(&mut self) -> Result<NameClass, ParseError> {
        let left = self.parse_name_class_primary()?;

        // Check for subtraction: `nc - nc`
        if self.check(&Token::Minus) {
            self.advance();
            let right = self.parse_name_class_primary()?;
            return Ok(NameClass::Except(Box::new(left), Box::new(right)));
        }

        Ok(left)
    }

    fn parse_name_class_primary(&mut self) -> Result<NameClass, ParseError> {
        // Check for wildcard `*`
        if self.check(&Token::Star) {
            self.advance();
            return Ok(NameClass::AnyName);
        }

        // Check for parenthesized name class (for choices/groups)
        if self.check(&Token::LParen) {
            self.advance();
            let mut choices = vec![self.parse_name_class()?];
            while self.check(&Token::Pipe) {
                self.advance();
                choices.push(self.parse_name_class()?);
            }
            self.expect(&Token::RParen)?;
            if choices.len() == 1 {
                return Ok(choices.pop().unwrap());
            }
            return Ok(NameClass::Choice(choices));
        }

        // Otherwise, parse as a QName (possibly with ns:* wildcard)
        let qname = self.parse_qname()?;

        // Check if it's a namespace wildcard: `ns:*`
        if qname.local == "*" {
            if let Some(prefix) = qname.prefix {
                return Ok(NameClass::NsName(prefix));
            }
            return Ok(NameClass::AnyName);
        }

        Ok(NameClass::Name(qname))
    }

    fn parse_qname(&mut self) -> Result<QName, ParseError> {
        // Element/attribute names can be keywords (e.g., "default", "string")
        let first = self.expect_name()?;

        if self.check(&Token::Colon) {
            self.advance();
            // After colon, could be a name or `*` for namespace wildcard
            let local = if self.check(&Token::Star) {
                self.advance();
                "*".to_string()
            } else {
                self.expect_name()?
            };
            Ok(QName {
                prefix: Some(first),
                local,
            })
        } else {
            Ok(QName {
                prefix: None,
                local: first,
            })
        }
    }

    /// Accept any token that can be a name (identifier or keyword).
    fn expect_name(&mut self) -> Result<String, ParseError> {
        match self.peek() {
            Some(Token::Ident(s)) => {
                let s = s.clone();
                self.advance();
                Ok(s)
            }
            Some(Token::String) => {
                self.advance();
                Ok("string".to_string())
            }
            Some(Token::Default) => {
                self.advance();
                Ok("default".to_string())
            }
            Some(Token::Element) => {
                self.advance();
                Ok("element".to_string())
            }
            Some(Token::Attribute) => {
                self.advance();
                Ok("attribute".to_string())
            }
            Some(Token::Namespace) => {
                self.advance();
                Ok("namespace".to_string())
            }
            Some(Token::Empty) => {
                self.advance();
                Ok("empty".to_string())
            }
            Some(Token::Mixed) => {
                self.advance();
                Ok("mixed".to_string())
            }
            Some(Token::List) => {
                self.advance();
                Ok("list".to_string())
            }
            Some(Token::Text) => {
                self.advance();
                Ok("text".to_string())
            }
            _ => Err(self.error("expected name")),
        }
    }

    fn parse_datatype_params(&mut self) -> Result<Vec<DatatypeParam>, ParseError> {
        if !self.check(&Token::LBrace) {
            return Ok(Vec::new());
        }
        self.advance();

        let mut params = Vec::new();
        while !self.check(&Token::RBrace) {
            let name = self.expect_ident()?;
            self.expect(&Token::Equals)?;
            let value = self.expect_string()?;
            params.push(DatatypeParam { name, value });
        }
        self.expect(&Token::RBrace)?;

        Ok(params)
    }

    // Helper methods

    fn peek(&self) -> Option<&Token> {
        self.tokens.get(self.pos)
    }

    fn check(&self, token: &Token) -> bool {
        self.peek()
            .is_some_and(|t| std::mem::discriminant(t) == std::mem::discriminant(token))
    }

    fn at_end(&self) -> bool {
        matches!(self.peek(), Some(Token::Eof) | None)
    }

    fn advance(&mut self) -> Option<&Token> {
        if !self.at_end() {
            self.pos += 1;
        }
        self.tokens.get(self.pos - 1)
    }

    fn expect(&mut self, expected: &Token) -> Result<(), ParseError> {
        if self.check(expected) {
            self.advance();
            Ok(())
        } else {
            Err(self.error(&format!("expected {:?}", expected)))
        }
    }

    fn expect_ident(&mut self) -> Result<String, ParseError> {
        match self.peek() {
            Some(Token::Ident(s)) => {
                let s = s.clone();
                self.advance();
                Ok(s)
            }
            _ => Err(self.error("expected identifier")),
        }
    }

    /// Like expect_ident but also accepts keywords that can be type names (e.g., "string" in xsd:string).
    fn expect_ident_or_keyword(&mut self) -> Result<String, ParseError> {
        match self.peek() {
            Some(Token::Ident(s)) => {
                let s = s.clone();
                self.advance();
                Ok(s)
            }
            Some(Token::String) => {
                self.advance();
                Ok("string".to_string())
            }
            _ => Err(self.error("expected identifier or type name")),
        }
    }

    fn expect_string(&mut self) -> Result<String, ParseError> {
        match self.peek() {
            Some(Token::QuotedString(s)) => {
                let s = s.clone();
                self.advance();
                Ok(s)
            }
            _ => Err(self.error("expected quoted string")),
        }
    }

    fn error(&self, msg: &str) -> ParseError {
        ParseError {
            message: msg.to_string(),
            position: self.pos,
            token: self.peek().cloned(),
        }
    }
}

#[derive(Debug, thiserror::Error)]
#[error("parse error at position {position}: {message} (found {:?})", token)]
pub struct ParseError {
    pub message: String,
    pub position: usize,
    pub token: Option<Token>,
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::lexer::Lexer;

    fn parse(input: &str) -> Schema {
        let tokens = Lexer::new(input).tokenize().unwrap();
        Parser::new(tokens).parse().unwrap()
    }

    #[test]
    fn test_empty_definition() {
        let schema = parse("w_CT_Empty = empty");
        assert_eq!(schema.definitions.len(), 1);
        assert_eq!(schema.definitions[0].name, "w_CT_Empty");
        assert!(matches!(schema.definitions[0].pattern, Pattern::Empty));
    }

    #[test]
    fn test_choice() {
        let schema = parse(r#"w_ST_Foo = string "a" | string "b" | string "c""#);
        assert_eq!(schema.definitions.len(), 1);
        match &schema.definitions[0].pattern {
            Pattern::Choice(v) => assert_eq!(v.len(), 3),
            _ => panic!("expected choice"),
        }
    }

    #[test]
    fn test_attribute() {
        let schema = parse("w_CT_OnOff = attribute w:val { s_ST_OnOff }?");
        assert_eq!(schema.definitions.len(), 1);
        match &schema.definitions[0].pattern {
            Pattern::Optional(inner) => match inner.as_ref() {
                Pattern::Attribute { name, .. } => {
                    assert_eq!(name.prefix, Some("w".into()));
                    assert_eq!(name.local, "val");
                }
                _ => panic!("expected attribute"),
            },
            _ => panic!("expected optional"),
        }
    }

    #[test]
    fn test_namespace() {
        let schema = parse(r#"default namespace w = "http://example.com""#);
        assert_eq!(schema.namespaces.len(), 1);
        assert!(schema.namespaces[0].is_default);
        assert_eq!(schema.namespaces[0].prefix, "w");
    }
}
