//! Lexer for RELAX NG Compact syntax.

use std::iter::Peekable;
use std::str::Chars;

/// Token types for RNC.
#[derive(Debug, Clone, PartialEq)]
pub enum Token {
    // Keywords
    Namespace,
    Default,
    Element,
    Attribute,
    Empty,
    String,
    // Identifiers and literals
    Ident(String),
    QuotedString(String),
    // Symbols
    Equals,
    Comma,
    Pipe,
    Ampersand,
    Question,
    Star,
    Plus,
    LBrace,
    RBrace,
    LParen,
    RParen,
    Colon,
    // Documentation
    DocComment(String),
    // End of file
    Eof,
}

/// Lexer state.
pub struct Lexer<'a> {
    input: Peekable<Chars<'a>>,
    current_line: usize,
}

impl<'a> Lexer<'a> {
    pub fn new(input: &'a str) -> Self {
        Self {
            input: input.chars().peekable(),
            current_line: 1,
        }
    }

    pub fn tokenize(mut self) -> Result<Vec<Token>, LexError> {
        let mut tokens = Vec::new();
        loop {
            let token = self.next_token()?;
            if token == Token::Eof {
                tokens.push(token);
                break;
            }
            tokens.push(token);
        }
        Ok(tokens)
    }

    fn next_token(&mut self) -> Result<Token, LexError> {
        self.skip_whitespace_and_comments();

        let Some(ch) = self.input.peek().copied() else {
            return Ok(Token::Eof);
        };

        match ch {
            '=' => {
                self.input.next();
                Ok(Token::Equals)
            }
            ',' => {
                self.input.next();
                Ok(Token::Comma)
            }
            '|' => {
                self.input.next();
                Ok(Token::Pipe)
            }
            '&' => {
                self.input.next();
                Ok(Token::Ampersand)
            }
            '?' => {
                self.input.next();
                Ok(Token::Question)
            }
            '*' => {
                self.input.next();
                Ok(Token::Star)
            }
            '+' => {
                self.input.next();
                Ok(Token::Plus)
            }
            '{' => {
                self.input.next();
                Ok(Token::LBrace)
            }
            '}' => {
                self.input.next();
                Ok(Token::RBrace)
            }
            '(' => {
                self.input.next();
                Ok(Token::LParen)
            }
            ')' => {
                self.input.next();
                Ok(Token::RParen)
            }
            ':' => {
                self.input.next();
                Ok(Token::Colon)
            }
            '"' => self.read_quoted_string(),
            _ if ch.is_alphabetic() || ch == '_' => self.read_ident(),
            _ => Err(LexError::UnexpectedChar(ch, self.current_line)),
        }
    }

    fn skip_whitespace_and_comments(&mut self) {
        loop {
            // Skip whitespace
            while let Some(&ch) = self.input.peek() {
                if ch == '\n' {
                    self.current_line += 1;
                    self.input.next();
                } else if ch.is_whitespace() {
                    self.input.next();
                } else {
                    break;
                }
            }

            // Check for comments (# or ##)
            if self.input.peek() == Some(&'#') {
                self.input.next(); // consume first #
                // Check for doc comment (##)
                let _is_doc = self.input.peek() == Some(&'#');
                if _is_doc {
                    self.input.next();
                }
                // Skip to end of line
                while let Some(&ch) = self.input.peek() {
                    if ch == '\n' {
                        self.current_line += 1;
                        self.input.next();
                        break;
                    }
                    self.input.next();
                }
            } else {
                break;
            }
        }
    }

    fn read_quoted_string(&mut self) -> Result<Token, LexError> {
        self.input.next(); // consume opening quote
        let mut s = String::new();
        loop {
            match self.input.next() {
                Some('"') => break,
                Some('\\') => {
                    // Handle escape sequences
                    match self.input.next() {
                        Some('n') => s.push('\n'),
                        Some('t') => s.push('\t'),
                        Some('\\') => s.push('\\'),
                        Some('"') => s.push('"'),
                        Some(ch) => s.push(ch),
                        None => return Err(LexError::UnterminatedString(self.current_line)),
                    }
                }
                Some('\n') => {
                    self.current_line += 1;
                    s.push('\n');
                }
                Some(ch) => s.push(ch),
                None => return Err(LexError::UnterminatedString(self.current_line)),
            }
        }
        Ok(Token::QuotedString(s))
    }

    fn read_ident(&mut self) -> Result<Token, LexError> {
        let mut s = String::new();
        while let Some(&ch) = self.input.peek() {
            if ch.is_alphanumeric() || ch == '_' || ch == '-' {
                s.push(ch);
                self.input.next();
            } else {
                break;
            }
        }
        let token = match s.as_str() {
            "namespace" => Token::Namespace,
            "default" => Token::Default,
            "element" => Token::Element,
            "attribute" => Token::Attribute,
            "empty" => Token::Empty,
            "string" => Token::String,
            _ => Token::Ident(s),
        };
        Ok(token)
    }
}

#[derive(Debug, thiserror::Error)]
pub enum LexError {
    #[error("unexpected character '{0}' at line {1}")]
    UnexpectedChar(char, usize),
    #[error("unterminated string at line {0}")]
    UnterminatedString(usize),
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_simple_definition() {
        let input = r#"w_CT_Empty = empty"#;
        let tokens = Lexer::new(input).tokenize().unwrap();
        assert_eq!(
            tokens,
            vec![
                Token::Ident("w_CT_Empty".into()),
                Token::Equals,
                Token::Empty,
                Token::Eof,
            ]
        );
    }

    #[test]
    fn test_attribute() {
        let input = r#"attribute w:val { s_ST_String }"#;
        let tokens = Lexer::new(input).tokenize().unwrap();
        assert_eq!(
            tokens,
            vec![
                Token::Attribute,
                Token::Ident("w".into()),
                Token::Colon,
                Token::Ident("val".into()),
                Token::LBrace,
                Token::Ident("s_ST_String".into()),
                Token::RBrace,
                Token::Eof,
            ]
        );
    }

    #[test]
    fn test_choice() {
        let input = r#"string "foo" | string "bar""#;
        let tokens = Lexer::new(input).tokenize().unwrap();
        assert_eq!(
            tokens,
            vec![
                Token::String,
                Token::QuotedString("foo".into()),
                Token::Pipe,
                Token::String,
                Token::QuotedString("bar".into()),
                Token::Eof,
            ]
        );
    }
}
