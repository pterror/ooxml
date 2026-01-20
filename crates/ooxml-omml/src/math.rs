//! Office Math Markup Language (OMML) types.
//!
//! OMML is used for mathematical formulas in Word, Excel, and PowerPoint.
//! Defined in ECMA-376 Part 4, Section 22.
//!
//! # Structure
//!
//! Math content is organized as:
//! - `MathZone` (`m:oMath`) - top-level container for inline math
//! - `MathParagraph` (`m:oMathPara`) - display math with alignment
//! - Various math constructs: fractions, radicals, scripts, etc.

use crate::error::{Error, Result};
use quick_xml::Reader;
use quick_xml::events::Event;
use std::io::{BufRead, Cursor};

/// A math zone (`<m:oMath>`).
///
/// Contains a sequence of math elements that form a mathematical expression.
#[derive(Debug, Clone, Default)]
pub struct MathZone {
    /// Math elements in this zone.
    pub elements: Vec<MathElement>,
}

impl MathZone {
    /// Create a new empty math zone.
    pub fn new() -> Self {
        Self::default()
    }

    /// Extract plain text representation of the math content.
    pub fn text(&self) -> String {
        self.elements.iter().map(|e| e.text()).collect()
    }
}

/// A math element - one of the possible OMML constructs.
#[derive(Debug, Clone)]
pub enum MathElement {
    /// Text run (`m:r`).
    Run(MathRun),
    /// Fraction (`m:f`).
    Fraction(Fraction),
    /// Radical/root (`m:rad`).
    Radical(Radical),
    /// N-ary operator like sum or integral (`m:nary`).
    Nary(Nary),
    /// Subscript (`m:sSub`).
    Subscript(Script),
    /// Superscript (`m:sSup`).
    Superscript(Script),
    /// Subscript and superscript (`m:sSubSup`).
    SubSuperscript(SubSuperscript),
    /// Pre-subscript/superscript (`m:sPre`).
    PreScript(PreScript),
    /// Delimiter/parentheses (`m:d`).
    Delimiter(Delimiter),
    /// Matrix (`m:m`).
    Matrix(Matrix),
    /// Function like sin, cos (`m:func`).
    Function(Function),
    /// Accent like hat, tilde (`m:acc`).
    Accent(Accent),
    /// Bar over/under (`m:bar`).
    Bar(Bar),
    /// Box (`m:box`).
    Box(MathBox),
    /// Border box (`m:borderBox`).
    BorderBox(BorderBox),
    /// Equation array (`m:eqArr`).
    EquationArray(EquationArray),
    /// Lower limit (`m:limLow`).
    LowerLimit(Limit),
    /// Upper limit (`m:limUpp`).
    UpperLimit(Limit),
    /// Group character/brace (`m:groupChr`).
    GroupChar(GroupChar),
    /// Phantom (`m:phant`).
    Phantom(Phantom),
}

impl MathElement {
    /// Get text representation.
    pub fn text(&self) -> String {
        match self {
            MathElement::Run(r) => r.text.clone(),
            MathElement::Fraction(f) => {
                format!("({})/({})", f.numerator.text(), f.denominator.text())
            }
            MathElement::Radical(r) => {
                if r.degree.elements.is_empty() {
                    format!("sqrt({})", r.base.text())
                } else {
                    format!("root[{}]({})", r.degree.text(), r.base.text())
                }
            }
            MathElement::Nary(n) => {
                let op = n.operator.as_deref().unwrap_or("∑");
                format!(
                    "{}[{},{}]({})",
                    op,
                    n.subscript.text(),
                    n.superscript.text(),
                    n.base.text()
                )
            }
            MathElement::Subscript(s) => format!("{}_{}", s.base.text(), s.script.text()),
            MathElement::Superscript(s) => format!("{}^{}", s.base.text(), s.script.text()),
            MathElement::SubSuperscript(s) => format!(
                "{}_{}^{}",
                s.base.text(),
                s.subscript.text(),
                s.superscript.text()
            ),
            MathElement::PreScript(p) => format!(
                "_{}^{}{}",
                p.subscript.text(),
                p.superscript.text(),
                p.base.text()
            ),
            MathElement::Delimiter(d) => {
                let begin = d.begin_char.as_deref().unwrap_or("(");
                let end = d.end_char.as_deref().unwrap_or(")");
                let inner: Vec<_> = d.elements.iter().map(|e| e.text()).collect();
                format!(
                    "{}{}{}",
                    begin,
                    inner.join(d.separator_char.as_deref().unwrap_or(",")),
                    end
                )
            }
            MathElement::Matrix(m) => {
                let rows: Vec<_> = m
                    .rows
                    .iter()
                    .map(|row| {
                        let cells: Vec<_> = row.iter().map(|c| c.text()).collect();
                        cells.join(", ")
                    })
                    .collect();
                format!("[{}]", rows.join("; "))
            }
            MathElement::Function(f) => format!("{}({})", f.name.text(), f.argument.text()),
            MathElement::Accent(a) => format!(
                "{}({})",
                a.character.as_deref().unwrap_or("^"),
                a.base.text()
            ),
            MathElement::Bar(b) => format!("bar({})", b.base.text()),
            MathElement::Box(b) => b.content.text(),
            MathElement::BorderBox(b) => format!("[{}]", b.content.text()),
            MathElement::EquationArray(e) => {
                let eqs: Vec<_> = e.equations.iter().map(|eq| eq.text()).collect();
                eqs.join("\n")
            }
            MathElement::LowerLimit(l) => format!("lim_{}({})", l.limit.text(), l.base.text()),
            MathElement::UpperLimit(l) => format!("lim^{}({})", l.limit.text(), l.base.text()),
            MathElement::GroupChar(g) => format!(
                "group[{}]({})",
                g.character.as_deref().unwrap_or("⏟"),
                g.base.text()
            ),
            MathElement::Phantom(p) => p.content.text(),
        }
    }
}

/// A math text run (`<m:r>`).
#[derive(Debug, Clone, Default)]
pub struct MathRun {
    /// Text content.
    pub text: String,
    /// Run properties.
    pub properties: Option<MathRunProperties>,
}

/// Math run properties (`<m:rPr>`).
#[derive(Debug, Clone, Default)]
pub struct MathRunProperties {
    /// Script type (roman, script, fraktur, etc.).
    pub script: Option<MathScript>,
    /// Style (plain, bold, italic, bold-italic).
    pub style: Option<MathStyle>,
    /// Literal text (no special formatting).
    pub literal: bool,
    /// Normal text (use document formatting).
    pub normal: bool,
}

/// Math script types.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum MathScript {
    Roman,
    Script,
    Fraktur,
    DoubleStruck,
    SansSerif,
    Monospace,
}

/// Math text styles.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum MathStyle {
    Plain,
    Bold,
    Italic,
    BoldItalic,
}

/// A fraction (`<m:f>`).
#[derive(Debug, Clone, Default)]
pub struct Fraction {
    /// Numerator.
    pub numerator: MathZone,
    /// Denominator.
    pub denominator: MathZone,
    /// Fraction type.
    pub fraction_type: Option<FractionType>,
}

/// Fraction display types.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum FractionType {
    /// Normal fraction bar.
    Bar,
    /// Skewed (diagonal).
    Skewed,
    /// Linear (inline).
    Linear,
    /// No bar.
    NoBar,
}

/// A radical/root (`<m:rad>`).
#[derive(Debug, Clone, Default)]
pub struct Radical {
    /// Base expression (under the radical).
    pub base: MathZone,
    /// Degree (for nth roots).
    pub degree: MathZone,
    /// Hide the degree.
    pub hide_degree: bool,
}

/// An n-ary operator (`<m:nary>`).
#[derive(Debug, Clone, Default)]
pub struct Nary {
    /// Operator character (∑, ∫, ∏, etc.).
    pub operator: Option<String>,
    /// Subscript (lower bound).
    pub subscript: MathZone,
    /// Superscript (upper bound).
    pub superscript: MathZone,
    /// Base expression.
    pub base: MathZone,
    /// Limit location (under/over or subscript/superscript).
    pub limit_location: Option<LimitLocation>,
    /// Whether operator grows with content.
    pub grow: bool,
}

/// Limit location for n-ary operators.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum LimitLocation {
    /// Under and over the operator.
    UnderOver,
    /// Subscript and superscript.
    SubSup,
}

/// A script (subscript or superscript).
#[derive(Debug, Clone, Default)]
pub struct Script {
    /// Base expression.
    pub base: MathZone,
    /// Script expression.
    pub script: MathZone,
}

/// Combined subscript and superscript (`<m:sSubSup>`).
#[derive(Debug, Clone, Default)]
pub struct SubSuperscript {
    /// Base expression.
    pub base: MathZone,
    /// Subscript.
    pub subscript: MathZone,
    /// Superscript.
    pub superscript: MathZone,
}

/// Pre-script (subscript/superscript before base) (`<m:sPre>`).
#[derive(Debug, Clone, Default)]
pub struct PreScript {
    /// Subscript.
    pub subscript: MathZone,
    /// Superscript.
    pub superscript: MathZone,
    /// Base expression.
    pub base: MathZone,
}

/// A delimiter/parentheses (`<m:d>`).
#[derive(Debug, Clone, Default)]
pub struct Delimiter {
    /// Beginning character (default: '(').
    pub begin_char: Option<String>,
    /// Separator character (default: ',').
    pub separator_char: Option<String>,
    /// Ending character (default: ')').
    pub end_char: Option<String>,
    /// Elements inside the delimiter.
    pub elements: Vec<MathZone>,
    /// Whether delimiters grow with content.
    pub grow: bool,
}

/// A matrix (`<m:m>`).
#[derive(Debug, Clone, Default)]
pub struct Matrix {
    /// Rows of the matrix.
    pub rows: Vec<Vec<MathZone>>,
}

/// A function application (`<m:func>`).
#[derive(Debug, Clone, Default)]
pub struct Function {
    /// Function name.
    pub name: MathZone,
    /// Function argument.
    pub argument: MathZone,
}

/// An accent (`<m:acc>`).
#[derive(Debug, Clone, Default)]
pub struct Accent {
    /// Accent character (default: combining circumflex).
    pub character: Option<String>,
    /// Base expression.
    pub base: MathZone,
}

/// A bar over or under expression (`<m:bar>`).
#[derive(Debug, Clone, Default)]
pub struct Bar {
    /// Base expression.
    pub base: MathZone,
    /// Position (top or bottom).
    pub position: Option<VerticalPosition>,
}

/// Vertical position for bars and limits.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum VerticalPosition {
    Top,
    Bottom,
}

/// A box around content (`<m:box>`).
#[derive(Debug, Clone, Default)]
pub struct MathBox {
    /// Box content.
    pub content: MathZone,
}

/// A bordered box (`<m:borderBox>`).
#[derive(Debug, Clone, Default)]
pub struct BorderBox {
    /// Box content.
    pub content: MathZone,
    /// Hide borders.
    pub hide_top: bool,
    pub hide_bottom: bool,
    pub hide_left: bool,
    pub hide_right: bool,
}

/// An equation array (`<m:eqArr>`).
#[derive(Debug, Clone, Default)]
pub struct EquationArray {
    /// Equations in the array.
    pub equations: Vec<MathZone>,
}

/// A limit expression (`<m:limLow>` or `<m:limUpp>`).
#[derive(Debug, Clone, Default)]
pub struct Limit {
    /// Base expression.
    pub base: MathZone,
    /// Limit expression.
    pub limit: MathZone,
}

/// A group character/brace (`<m:groupChr>`).
#[derive(Debug, Clone, Default)]
pub struct GroupChar {
    /// Grouping character (default: underbrace).
    pub character: Option<String>,
    /// Position of the character.
    pub position: Option<VerticalPosition>,
    /// Base expression.
    pub base: MathZone,
}

/// A phantom (invisible content for spacing) (`<m:phant>`).
#[derive(Debug, Clone, Default)]
pub struct Phantom {
    /// Content.
    pub content: MathZone,
    /// Show the content.
    pub show: bool,
}

// ============================================================================
// Parsing
// ============================================================================

/// Parse an OMML math zone from XML.
pub fn parse_math_zone(xml: &[u8]) -> Result<MathZone> {
    let mut reader = Reader::from_reader(Cursor::new(xml));
    parse_math_zone_from_reader(&mut reader)
}

/// Parse an OMML math zone from a reader.
pub fn parse_math_zone_from_reader<R: BufRead>(reader: &mut Reader<R>) -> Result<MathZone> {
    let mut buf = Vec::new();
    let mut zone = MathZone::new();

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) => {
                let name = e.name();
                let name = name.as_ref();
                if let Some(element) = parse_math_element_start(name, reader)? {
                    zone.elements.push(element);
                }
            }
            Ok(Event::Empty(e)) => {
                // Handle self-closing elements if needed
                let _ = e;
            }
            Ok(Event::End(e)) => {
                let name = e.name();
                if name.as_ref() == b"m:oMath" || name.as_ref() == b"oMath" {
                    break;
                }
            }
            Ok(Event::Eof) => break,
            Err(e) => return Err(Error::Xml(e)),
            _ => {}
        }
        buf.clear();
    }

    Ok(zone)
}

/// Parse a math element from its start tag.
fn parse_math_element_start<R: BufRead>(
    name: &[u8],
    reader: &mut Reader<R>,
) -> Result<Option<MathElement>> {
    match name {
        b"m:r" | b"r" => Ok(Some(MathElement::Run(parse_math_run(reader)?))),
        b"m:f" | b"f" => Ok(Some(MathElement::Fraction(parse_fraction(reader)?))),
        b"m:rad" | b"rad" => Ok(Some(MathElement::Radical(parse_radical(reader)?))),
        b"m:nary" | b"nary" => Ok(Some(MathElement::Nary(parse_nary(reader)?))),
        b"m:sSub" | b"sSub" => Ok(Some(MathElement::Subscript(parse_script(
            reader, b"m:sSub",
        )?))),
        b"m:sSup" | b"sSup" => Ok(Some(MathElement::Superscript(parse_script(
            reader, b"m:sSup",
        )?))),
        b"m:sSubSup" | b"sSubSup" => Ok(Some(MathElement::SubSuperscript(parse_sub_superscript(
            reader,
        )?))),
        b"m:sPre" | b"sPre" => Ok(Some(MathElement::PreScript(parse_pre_script(reader)?))),
        b"m:d" | b"d" => Ok(Some(MathElement::Delimiter(parse_delimiter(reader)?))),
        b"m:m" | b"m" => Ok(Some(MathElement::Matrix(parse_matrix(reader)?))),
        b"m:func" | b"func" => Ok(Some(MathElement::Function(parse_function(reader)?))),
        b"m:acc" | b"acc" => Ok(Some(MathElement::Accent(parse_accent(reader)?))),
        b"m:bar" | b"bar" => Ok(Some(MathElement::Bar(parse_bar(reader)?))),
        b"m:box" | b"box" => Ok(Some(MathElement::Box(parse_math_box(reader)?))),
        b"m:borderBox" | b"borderBox" => {
            Ok(Some(MathElement::BorderBox(parse_border_box(reader)?)))
        }
        b"m:eqArr" | b"eqArr" => Ok(Some(MathElement::EquationArray(parse_equation_array(
            reader,
        )?))),
        b"m:limLow" | b"limLow" => Ok(Some(MathElement::LowerLimit(parse_limit(
            reader,
            b"m:limLow",
        )?))),
        b"m:limUpp" | b"limUpp" => Ok(Some(MathElement::UpperLimit(parse_limit(
            reader,
            b"m:limUpp",
        )?))),
        b"m:groupChr" | b"groupChr" => Ok(Some(MathElement::GroupChar(parse_group_char(reader)?))),
        b"m:phant" | b"phant" => Ok(Some(MathElement::Phantom(parse_phantom(reader)?))),
        _ => Ok(None),
    }
}

/// Parse a math run.
fn parse_math_run<R: BufRead>(reader: &mut Reader<R>) -> Result<MathRun> {
    let mut buf = Vec::new();
    let mut run = MathRun::default();
    let mut in_text = false;

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) => {
                let name = e.name();
                if name.as_ref() == b"m:t" || name.as_ref() == b"t" {
                    in_text = true;
                }
            }
            Ok(Event::Text(e)) => {
                if in_text {
                    run.text.push_str(&e.decode().unwrap_or_default());
                }
            }
            Ok(Event::End(e)) => {
                let name = e.name();
                let name = name.as_ref();
                if name == b"m:t" || name == b"t" {
                    in_text = false;
                } else if name == b"m:r" || name == b"r" {
                    break;
                }
            }
            Ok(Event::Eof) => break,
            Err(e) => return Err(Error::Xml(e)),
            _ => {}
        }
        buf.clear();
    }

    Ok(run)
}

/// Parse math argument (content inside m:e, m:num, m:den, etc.).
fn parse_math_arg<R: BufRead>(reader: &mut Reader<R>, end_tag: &[u8]) -> Result<MathZone> {
    let mut buf = Vec::new();
    let mut zone = MathZone::new();

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) => {
                let name = e.name();
                let name = name.as_ref();
                if let Some(element) = parse_math_element_start(name, reader)? {
                    zone.elements.push(element);
                }
            }
            Ok(Event::End(e)) => {
                let name = e.name();
                if name.as_ref() == end_tag {
                    break;
                }
            }
            Ok(Event::Eof) => break,
            Err(e) => return Err(Error::Xml(e)),
            _ => {}
        }
        buf.clear();
    }

    Ok(zone)
}

/// Parse a fraction.
fn parse_fraction<R: BufRead>(reader: &mut Reader<R>) -> Result<Fraction> {
    let mut buf = Vec::new();
    let mut fraction = Fraction::default();

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) => {
                let name = e.name();
                let name = name.as_ref();
                match name {
                    b"m:num" | b"num" => fraction.numerator = parse_math_arg(reader, name)?,
                    b"m:den" | b"den" => fraction.denominator = parse_math_arg(reader, name)?,
                    _ => {}
                }
            }
            Ok(Event::End(e)) => {
                let name = e.name();
                if name.as_ref() == b"m:f" || name.as_ref() == b"f" {
                    break;
                }
            }
            Ok(Event::Eof) => break,
            Err(e) => return Err(Error::Xml(e)),
            _ => {}
        }
        buf.clear();
    }

    Ok(fraction)
}

/// Parse a radical.
fn parse_radical<R: BufRead>(reader: &mut Reader<R>) -> Result<Radical> {
    let mut buf = Vec::new();
    let mut radical = Radical::default();

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) => {
                let name = e.name();
                let name = name.as_ref();
                match name {
                    b"m:e" | b"e" => radical.base = parse_math_arg(reader, name)?,
                    b"m:deg" | b"deg" => radical.degree = parse_math_arg(reader, name)?,
                    _ => {}
                }
            }
            Ok(Event::End(e)) => {
                let name = e.name();
                if name.as_ref() == b"m:rad" || name.as_ref() == b"rad" {
                    break;
                }
            }
            Ok(Event::Eof) => break,
            Err(e) => return Err(Error::Xml(e)),
            _ => {}
        }
        buf.clear();
    }

    Ok(radical)
}

/// Parse an n-ary operator.
fn parse_nary<R: BufRead>(reader: &mut Reader<R>) -> Result<Nary> {
    let mut buf = Vec::new();
    let mut nary = Nary::default();

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) => {
                let name = e.name();
                let name = name.as_ref();
                match name {
                    b"m:e" | b"e" => nary.base = parse_math_arg(reader, name)?,
                    b"m:sub" | b"sub" => nary.subscript = parse_math_arg(reader, name)?,
                    b"m:sup" | b"sup" => nary.superscript = parse_math_arg(reader, name)?,
                    _ => {}
                }
            }
            Ok(Event::Empty(e)) => {
                let name = e.name();
                if name.as_ref() == b"m:chr" || name.as_ref() == b"chr" {
                    for attr in e.attributes().filter_map(|a| a.ok()) {
                        if attr.key.as_ref() == b"m:val" || attr.key.as_ref() == b"val" {
                            nary.operator = Some(String::from_utf8_lossy(&attr.value).into_owned());
                        }
                    }
                }
            }
            Ok(Event::End(e)) => {
                let name = e.name();
                if name.as_ref() == b"m:nary" || name.as_ref() == b"nary" {
                    break;
                }
            }
            Ok(Event::Eof) => break,
            Err(e) => return Err(Error::Xml(e)),
            _ => {}
        }
        buf.clear();
    }

    Ok(nary)
}

/// Parse a script (subscript or superscript).
fn parse_script<R: BufRead>(reader: &mut Reader<R>, end_tag: &[u8]) -> Result<Script> {
    let mut buf = Vec::new();
    let mut script = Script::default();
    let end_local = if end_tag.starts_with(b"m:") {
        &end_tag[2..]
    } else {
        end_tag
    };

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) => {
                let name = e.name();
                let name = name.as_ref();
                match name {
                    b"m:e" | b"e" => script.base = parse_math_arg(reader, name)?,
                    b"m:sub" | b"sub" | b"m:sup" | b"sup" => {
                        script.script = parse_math_arg(reader, name)?
                    }
                    _ => {}
                }
            }
            Ok(Event::End(e)) => {
                let name = e.name();
                let name = name.as_ref();
                if name == end_tag || name == end_local {
                    break;
                }
            }
            Ok(Event::Eof) => break,
            Err(e) => return Err(Error::Xml(e)),
            _ => {}
        }
        buf.clear();
    }

    Ok(script)
}

/// Parse combined subscript and superscript.
fn parse_sub_superscript<R: BufRead>(reader: &mut Reader<R>) -> Result<SubSuperscript> {
    let mut buf = Vec::new();
    let mut result = SubSuperscript::default();

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) => {
                let name = e.name();
                let name = name.as_ref();
                match name {
                    b"m:e" | b"e" => result.base = parse_math_arg(reader, name)?,
                    b"m:sub" | b"sub" => result.subscript = parse_math_arg(reader, name)?,
                    b"m:sup" | b"sup" => result.superscript = parse_math_arg(reader, name)?,
                    _ => {}
                }
            }
            Ok(Event::End(e)) => {
                let name = e.name();
                if name.as_ref() == b"m:sSubSup" || name.as_ref() == b"sSubSup" {
                    break;
                }
            }
            Ok(Event::Eof) => break,
            Err(e) => return Err(Error::Xml(e)),
            _ => {}
        }
        buf.clear();
    }

    Ok(result)
}

/// Parse pre-script.
fn parse_pre_script<R: BufRead>(reader: &mut Reader<R>) -> Result<PreScript> {
    let mut buf = Vec::new();
    let mut result = PreScript::default();

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) => {
                let name = e.name();
                let name = name.as_ref();
                match name {
                    b"m:e" | b"e" => result.base = parse_math_arg(reader, name)?,
                    b"m:sub" | b"sub" => result.subscript = parse_math_arg(reader, name)?,
                    b"m:sup" | b"sup" => result.superscript = parse_math_arg(reader, name)?,
                    _ => {}
                }
            }
            Ok(Event::End(e)) => {
                let name = e.name();
                if name.as_ref() == b"m:sPre" || name.as_ref() == b"sPre" {
                    break;
                }
            }
            Ok(Event::Eof) => break,
            Err(e) => return Err(Error::Xml(e)),
            _ => {}
        }
        buf.clear();
    }

    Ok(result)
}

/// Parse a delimiter.
fn parse_delimiter<R: BufRead>(reader: &mut Reader<R>) -> Result<Delimiter> {
    let mut buf = Vec::new();
    let mut delimiter = Delimiter::default();

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) => {
                let name = e.name();
                let name = name.as_ref();
                if name == b"m:e" || name == b"e" {
                    delimiter.elements.push(parse_math_arg(reader, name)?);
                }
            }
            Ok(Event::Empty(e)) => {
                let name = e.name();
                let name = name.as_ref();
                for attr in e.attributes().filter_map(|a| a.ok()) {
                    let val = String::from_utf8_lossy(&attr.value).into_owned();
                    if attr.key.as_ref() == b"m:val" || attr.key.as_ref() == b"val" {
                        match name {
                            b"m:begChr" | b"begChr" => delimiter.begin_char = Some(val),
                            b"m:sepChr" | b"sepChr" => delimiter.separator_char = Some(val),
                            b"m:endChr" | b"endChr" => delimiter.end_char = Some(val),
                            _ => {}
                        }
                    }
                }
            }
            Ok(Event::End(e)) => {
                let name = e.name();
                if name.as_ref() == b"m:d" || name.as_ref() == b"d" {
                    break;
                }
            }
            Ok(Event::Eof) => break,
            Err(e) => return Err(Error::Xml(e)),
            _ => {}
        }
        buf.clear();
    }

    Ok(delimiter)
}

/// Parse a matrix.
fn parse_matrix<R: BufRead>(reader: &mut Reader<R>) -> Result<Matrix> {
    let mut buf = Vec::new();
    let mut matrix = Matrix::default();
    let mut current_row: Vec<MathZone> = Vec::new();
    let mut in_row = false;

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) => {
                let name = e.name();
                let name = name.as_ref();
                match name {
                    b"m:mr" | b"mr" => {
                        in_row = true;
                        current_row = Vec::new();
                    }
                    b"m:e" | b"e" if in_row => {
                        current_row.push(parse_math_arg(reader, name)?);
                    }
                    _ => {}
                }
            }
            Ok(Event::End(e)) => {
                let name = e.name();
                let name = name.as_ref();
                match name {
                    b"m:mr" | b"mr" => {
                        matrix.rows.push(std::mem::take(&mut current_row));
                        in_row = false;
                    }
                    b"m:m" => break,
                    _ => {}
                }
            }
            Ok(Event::Eof) => break,
            Err(e) => return Err(Error::Xml(e)),
            _ => {}
        }
        buf.clear();
    }

    Ok(matrix)
}

/// Parse a function.
fn parse_function<R: BufRead>(reader: &mut Reader<R>) -> Result<Function> {
    let mut buf = Vec::new();
    let mut function = Function::default();

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) => {
                let name = e.name();
                let name = name.as_ref();
                match name {
                    b"m:fName" | b"fName" => function.name = parse_math_arg(reader, name)?,
                    b"m:e" | b"e" => function.argument = parse_math_arg(reader, name)?,
                    _ => {}
                }
            }
            Ok(Event::End(e)) => {
                let name = e.name();
                if name.as_ref() == b"m:func" || name.as_ref() == b"func" {
                    break;
                }
            }
            Ok(Event::Eof) => break,
            Err(e) => return Err(Error::Xml(e)),
            _ => {}
        }
        buf.clear();
    }

    Ok(function)
}

/// Parse an accent.
fn parse_accent<R: BufRead>(reader: &mut Reader<R>) -> Result<Accent> {
    let mut buf = Vec::new();
    let mut accent = Accent::default();

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) => {
                let name = e.name();
                let name = name.as_ref();
                if name == b"m:e" || name == b"e" {
                    accent.base = parse_math_arg(reader, name)?;
                }
            }
            Ok(Event::Empty(e)) => {
                let name = e.name();
                if name.as_ref() == b"m:chr" || name.as_ref() == b"chr" {
                    for attr in e.attributes().filter_map(|a| a.ok()) {
                        if attr.key.as_ref() == b"m:val" || attr.key.as_ref() == b"val" {
                            accent.character =
                                Some(String::from_utf8_lossy(&attr.value).into_owned());
                        }
                    }
                }
            }
            Ok(Event::End(e)) => {
                let name = e.name();
                if name.as_ref() == b"m:acc" || name.as_ref() == b"acc" {
                    break;
                }
            }
            Ok(Event::Eof) => break,
            Err(e) => return Err(Error::Xml(e)),
            _ => {}
        }
        buf.clear();
    }

    Ok(accent)
}

/// Parse a bar.
fn parse_bar<R: BufRead>(reader: &mut Reader<R>) -> Result<Bar> {
    let mut buf = Vec::new();
    let mut bar = Bar::default();

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) => {
                let name = e.name();
                let name = name.as_ref();
                if name == b"m:e" || name == b"e" {
                    bar.base = parse_math_arg(reader, name)?;
                }
            }
            Ok(Event::End(e)) => {
                let name = e.name();
                if name.as_ref() == b"m:bar" || name.as_ref() == b"bar" {
                    break;
                }
            }
            Ok(Event::Eof) => break,
            Err(e) => return Err(Error::Xml(e)),
            _ => {}
        }
        buf.clear();
    }

    Ok(bar)
}

/// Parse a math box.
fn parse_math_box<R: BufRead>(reader: &mut Reader<R>) -> Result<MathBox> {
    let mut buf = Vec::new();
    let mut result = MathBox::default();

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) => {
                let name = e.name();
                let name = name.as_ref();
                if name == b"m:e" || name == b"e" {
                    result.content = parse_math_arg(reader, name)?;
                }
            }
            Ok(Event::End(e)) => {
                let name = e.name();
                if name.as_ref() == b"m:box" || name.as_ref() == b"box" {
                    break;
                }
            }
            Ok(Event::Eof) => break,
            Err(e) => return Err(Error::Xml(e)),
            _ => {}
        }
        buf.clear();
    }

    Ok(result)
}

/// Parse a border box.
fn parse_border_box<R: BufRead>(reader: &mut Reader<R>) -> Result<BorderBox> {
    let mut buf = Vec::new();
    let mut result = BorderBox::default();

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) => {
                let name = e.name();
                let name = name.as_ref();
                if name == b"m:e" || name == b"e" {
                    result.content = parse_math_arg(reader, name)?;
                }
            }
            Ok(Event::End(e)) => {
                let name = e.name();
                if name.as_ref() == b"m:borderBox" || name.as_ref() == b"borderBox" {
                    break;
                }
            }
            Ok(Event::Eof) => break,
            Err(e) => return Err(Error::Xml(e)),
            _ => {}
        }
        buf.clear();
    }

    Ok(result)
}

/// Parse an equation array.
fn parse_equation_array<R: BufRead>(reader: &mut Reader<R>) -> Result<EquationArray> {
    let mut buf = Vec::new();
    let mut result = EquationArray::default();

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) => {
                let name = e.name();
                let name = name.as_ref();
                if name == b"m:e" || name == b"e" {
                    result.equations.push(parse_math_arg(reader, name)?);
                }
            }
            Ok(Event::End(e)) => {
                let name = e.name();
                if name.as_ref() == b"m:eqArr" || name.as_ref() == b"eqArr" {
                    break;
                }
            }
            Ok(Event::Eof) => break,
            Err(e) => return Err(Error::Xml(e)),
            _ => {}
        }
        buf.clear();
    }

    Ok(result)
}

/// Parse a limit.
fn parse_limit<R: BufRead>(reader: &mut Reader<R>, end_tag: &[u8]) -> Result<Limit> {
    let mut buf = Vec::new();
    let mut result = Limit::default();
    let end_local = if end_tag.starts_with(b"m:") {
        &end_tag[2..]
    } else {
        end_tag
    };

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) => {
                let name = e.name();
                let name = name.as_ref();
                match name {
                    b"m:e" | b"e" => result.base = parse_math_arg(reader, name)?,
                    b"m:lim" | b"lim" => result.limit = parse_math_arg(reader, name)?,
                    _ => {}
                }
            }
            Ok(Event::End(e)) => {
                let name = e.name();
                let name = name.as_ref();
                if name == end_tag || name == end_local {
                    break;
                }
            }
            Ok(Event::Eof) => break,
            Err(e) => return Err(Error::Xml(e)),
            _ => {}
        }
        buf.clear();
    }

    Ok(result)
}

/// Parse a group character.
fn parse_group_char<R: BufRead>(reader: &mut Reader<R>) -> Result<GroupChar> {
    let mut buf = Vec::new();
    let mut result = GroupChar::default();

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) => {
                let name = e.name();
                let name = name.as_ref();
                if name == b"m:e" || name == b"e" {
                    result.base = parse_math_arg(reader, name)?;
                }
            }
            Ok(Event::Empty(e)) => {
                let name = e.name();
                if name.as_ref() == b"m:chr" || name.as_ref() == b"chr" {
                    for attr in e.attributes().filter_map(|a| a.ok()) {
                        if attr.key.as_ref() == b"m:val" || attr.key.as_ref() == b"val" {
                            result.character =
                                Some(String::from_utf8_lossy(&attr.value).into_owned());
                        }
                    }
                }
            }
            Ok(Event::End(e)) => {
                let name = e.name();
                if name.as_ref() == b"m:groupChr" || name.as_ref() == b"groupChr" {
                    break;
                }
            }
            Ok(Event::Eof) => break,
            Err(e) => return Err(Error::Xml(e)),
            _ => {}
        }
        buf.clear();
    }

    Ok(result)
}

/// Parse a phantom.
fn parse_phantom<R: BufRead>(reader: &mut Reader<R>) -> Result<Phantom> {
    let mut buf = Vec::new();
    let mut result = Phantom::default();

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) => {
                let name = e.name();
                let name = name.as_ref();
                if name == b"m:e" || name == b"e" {
                    result.content = parse_math_arg(reader, name)?;
                }
            }
            Ok(Event::End(e)) => {
                let name = e.name();
                if name.as_ref() == b"m:phant" || name.as_ref() == b"phant" {
                    break;
                }
            }
            Ok(Event::Eof) => break,
            Err(e) => return Err(Error::Xml(e)),
            _ => {}
        }
        buf.clear();
    }

    Ok(result)
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_parse_simple_math() {
        let xml = r#"<m:oMath xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">
            <m:r><m:t>x+y</m:t></m:r>
        </m:oMath>"#;

        let zone = parse_math_zone(xml.as_bytes()).unwrap();
        assert_eq!(zone.text(), "x+y");
    }

    #[test]
    fn test_parse_fraction() {
        let xml = r#"<m:oMath xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">
            <m:f>
                <m:num><m:r><m:t>1</m:t></m:r></m:num>
                <m:den><m:r><m:t>2</m:t></m:r></m:den>
            </m:f>
        </m:oMath>"#;

        let zone = parse_math_zone(xml.as_bytes()).unwrap();
        assert_eq!(zone.text(), "(1)/(2)");
    }

    #[test]
    fn test_parse_nested_fraction() {
        let xml = r#"<m:oMath xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">
            <m:f>
                <m:num>
                    <m:r><m:t>a+</m:t></m:r>
                    <m:f>
                        <m:num><m:r><m:t>b</m:t></m:r></m:num>
                        <m:den><m:r><m:t>c</m:t></m:r></m:den>
                    </m:f>
                </m:num>
                <m:den><m:r><m:t>d</m:t></m:r></m:den>
            </m:f>
        </m:oMath>"#;

        let zone = parse_math_zone(xml.as_bytes()).unwrap();
        assert_eq!(zone.text(), "(a+(b)/(c))/(d)");
    }

    #[test]
    fn test_parse_radical() {
        let xml = r#"<m:oMath xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">
            <m:rad>
                <m:deg></m:deg>
                <m:e><m:r><m:t>x</m:t></m:r></m:e>
            </m:rad>
        </m:oMath>"#;

        let zone = parse_math_zone(xml.as_bytes()).unwrap();
        assert_eq!(zone.text(), "sqrt(x)");
    }

    #[test]
    fn test_parse_subscript() {
        let xml = r#"<m:oMath xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">
            <m:sSub>
                <m:e><m:r><m:t>x</m:t></m:r></m:e>
                <m:sub><m:r><m:t>i</m:t></m:r></m:sub>
            </m:sSub>
        </m:oMath>"#;

        let zone = parse_math_zone(xml.as_bytes()).unwrap();
        assert_eq!(zone.text(), "x_i");
    }

    #[test]
    fn test_parse_delimiter() {
        let xml = r#"<m:oMath xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">
            <m:d>
                <m:e><m:r><m:t>a+b</m:t></m:r></m:e>
            </m:d>
        </m:oMath>"#;

        let zone = parse_math_zone(xml.as_bytes()).unwrap();
        assert_eq!(zone.text(), "(a+b)");
    }

    #[test]
    fn test_parse_real_world_formula() {
        // Real formula from NapierOne corpus: ((30 + (90/2)) / 900) * 100 = 8%
        // Simplified version without WML formatting elements
        let xml = r#"<m:oMath xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">
            <m:r><m:t>       </m:t></m:r>
            <m:d>
                <m:e>
                    <m:f>
                        <m:num>
                            <m:r><m:t>30+</m:t></m:r>
                            <m:d>
                                <m:e>
                                    <m:f>
                                        <m:num><m:r><m:t>90</m:t></m:r></m:num>
                                        <m:den><m:r><m:t>2</m:t></m:r></m:den>
                                    </m:f>
                                </m:e>
                            </m:d>
                        </m:num>
                        <m:den><m:r><m:t>900</m:t></m:r></m:den>
                    </m:f>
                </m:e>
            </m:d>
            <m:r><m:t>*100=8%</m:t></m:r>
        </m:oMath>"#;

        let zone = parse_math_zone(xml.as_bytes()).unwrap();
        let text = zone.text();
        // Verify key parts are present
        assert!(text.contains("30+"));
        assert!(text.contains("90"));
        assert!(text.contains("900"));
        assert!(text.contains("*100=8%"));
    }
}
