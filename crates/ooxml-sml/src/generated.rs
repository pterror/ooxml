// Generated from ECMA-376 RELAX NG schema.
// Do not edit manually.

use serde::{Deserialize, Serialize};

/// XML namespace URIs used in this schema.
pub mod ns {
    /// Namespace prefix: o
    pub const O: &str = "urn:schemas-microsoft-com:office:office";
    /// Namespace prefix: s
    pub const S: &str = "http://schemas.openxmlformats.org/officeDocument/2006/sharedTypes";
    /// Namespace prefix: v
    pub const V: &str = "urn:schemas-microsoft-com:vml";
    /// Namespace prefix: w10
    pub const W10: &str = "urn:schemas-microsoft-com:office:word";
    /// Namespace prefix: x
    pub const X: &str = "urn:schemas-microsoft-com:office:excel";
    /// Namespace prefix: r
    pub const R: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    /// Default namespace (prefix: sml)
    pub const SML: &str = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
    /// Namespace prefix: xdr
    pub const XDR: &str = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing";
}

pub type Language = String;

pub type HexColorRgb = Vec<u8>;

pub type Panose = Vec<u8>;

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum CalendarType {
    #[serde(rename = "gregorian")]
    Gregorian,
    #[serde(rename = "gregorianUs")]
    GregorianUs,
    #[serde(rename = "gregorianMeFrench")]
    GregorianMeFrench,
    #[serde(rename = "gregorianArabic")]
    GregorianArabic,
    #[serde(rename = "hijri")]
    Hijri,
    #[serde(rename = "hebrew")]
    Hebrew,
    #[serde(rename = "taiwan")]
    Taiwan,
    #[serde(rename = "japan")]
    Japan,
    #[serde(rename = "thai")]
    Thai,
    #[serde(rename = "korea")]
    Korea,
    #[serde(rename = "saka")]
    Saka,
    #[serde(rename = "gregorianXlitEnglish")]
    GregorianXlitEnglish,
    #[serde(rename = "gregorianXlitFrench")]
    GregorianXlitFrench,
    #[serde(rename = "none")]
    None,
}

impl std::fmt::Display for CalendarType {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::Gregorian => write!(f, "gregorian"),
            Self::GregorianUs => write!(f, "gregorianUs"),
            Self::GregorianMeFrench => write!(f, "gregorianMeFrench"),
            Self::GregorianArabic => write!(f, "gregorianArabic"),
            Self::Hijri => write!(f, "hijri"),
            Self::Hebrew => write!(f, "hebrew"),
            Self::Taiwan => write!(f, "taiwan"),
            Self::Japan => write!(f, "japan"),
            Self::Thai => write!(f, "thai"),
            Self::Korea => write!(f, "korea"),
            Self::Saka => write!(f, "saka"),
            Self::GregorianXlitEnglish => write!(f, "gregorianXlitEnglish"),
            Self::GregorianXlitFrench => write!(f, "gregorianXlitFrench"),
            Self::None => write!(f, "none"),
        }
    }
}

impl std::str::FromStr for CalendarType {
    type Err = String;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        match s {
            "gregorian" => Ok(Self::Gregorian),
            "gregorianUs" => Ok(Self::GregorianUs),
            "gregorianMeFrench" => Ok(Self::GregorianMeFrench),
            "gregorianArabic" => Ok(Self::GregorianArabic),
            "hijri" => Ok(Self::Hijri),
            "hebrew" => Ok(Self::Hebrew),
            "taiwan" => Ok(Self::Taiwan),
            "japan" => Ok(Self::Japan),
            "thai" => Ok(Self::Thai),
            "korea" => Ok(Self::Korea),
            "saka" => Ok(Self::Saka),
            "gregorianXlitEnglish" => Ok(Self::GregorianXlitEnglish),
            "gregorianXlitFrench" => Ok(Self::GregorianXlitFrench),
            "none" => Ok(Self::None),
            _ => Err(format!("unknown CalendarType value: {}", s)),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum STAlgClass {
    #[serde(rename = "hash")]
    Hash,
    #[serde(rename = "custom")]
    Custom,
}

impl std::fmt::Display for STAlgClass {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::Hash => write!(f, "hash"),
            Self::Custom => write!(f, "custom"),
        }
    }
}

impl std::str::FromStr for STAlgClass {
    type Err = String;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        match s {
            "hash" => Ok(Self::Hash),
            "custom" => Ok(Self::Custom),
            _ => Err(format!("unknown STAlgClass value: {}", s)),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum STCryptProv {
    #[serde(rename = "rsaAES")]
    RsaAES,
    #[serde(rename = "rsaFull")]
    RsaFull,
    #[serde(rename = "custom")]
    Custom,
}

impl std::fmt::Display for STCryptProv {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::RsaAES => write!(f, "rsaAES"),
            Self::RsaFull => write!(f, "rsaFull"),
            Self::Custom => write!(f, "custom"),
        }
    }
}

impl std::str::FromStr for STCryptProv {
    type Err = String;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        match s {
            "rsaAES" => Ok(Self::RsaAES),
            "rsaFull" => Ok(Self::RsaFull),
            "custom" => Ok(Self::Custom),
            _ => Err(format!("unknown STCryptProv value: {}", s)),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum STAlgType {
    #[serde(rename = "typeAny")]
    TypeAny,
    #[serde(rename = "custom")]
    Custom,
}

impl std::fmt::Display for STAlgType {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::TypeAny => write!(f, "typeAny"),
            Self::Custom => write!(f, "custom"),
        }
    }
}

impl std::str::FromStr for STAlgType {
    type Err = String;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        match s {
            "typeAny" => Ok(Self::TypeAny),
            "custom" => Ok(Self::Custom),
            _ => Err(format!("unknown STAlgType value: {}", s)),
        }
    }
}

pub type STColorType = String;

pub type Guid = String;

pub type OnOff = String;

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum STOnOff1 {
    #[serde(rename = "on")]
    On,
    #[serde(rename = "off")]
    Off,
}

impl std::fmt::Display for STOnOff1 {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::On => write!(f, "on"),
            Self::Off => write!(f, "off"),
        }
    }
}

impl std::str::FromStr for STOnOff1 {
    type Err = String;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        match s {
            "on" => Ok(Self::On),
            "off" => Ok(Self::Off),
            _ => Err(format!("unknown STOnOff1 value: {}", s)),
        }
    }
}

pub type STString = String;

pub type STXmlName = String;

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum TrueFalse {
    #[serde(rename = "t")]
    T,
    #[serde(rename = "f")]
    F,
    #[serde(rename = "true")]
    True,
    #[serde(rename = "false")]
    False,
}

impl std::fmt::Display for TrueFalse {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::T => write!(f, "t"),
            Self::F => write!(f, "f"),
            Self::True => write!(f, "true"),
            Self::False => write!(f, "false"),
        }
    }
}

impl std::str::FromStr for TrueFalse {
    type Err = String;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        match s {
            "t" => Ok(Self::T),
            "f" => Ok(Self::F),
            "true" => Ok(Self::True),
            "false" => Ok(Self::False),
            _ => Err(format!("unknown TrueFalse value: {}", s)),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum STTrueFalseBlank {
    #[serde(rename = "t")]
    T,
    #[serde(rename = "f")]
    F,
    #[serde(rename = "true")]
    True,
    #[serde(rename = "false")]
    False,
    #[serde(rename = "")]
    Empty,
}

impl std::fmt::Display for STTrueFalseBlank {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::T => write!(f, "t"),
            Self::F => write!(f, "f"),
            Self::True => write!(f, "true"),
            Self::False => write!(f, "false"),
            Self::Empty => write!(f, ""),
        }
    }
}

impl std::str::FromStr for STTrueFalseBlank {
    type Err = String;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        match s {
            "t" => Ok(Self::T),
            "f" => Ok(Self::F),
            "true" => Ok(Self::True),
            "false" => Ok(Self::False),
            "" => Ok(Self::Empty),
            "True" => Ok(Self::True),
            "False" => Ok(Self::False),
            _ => Err(format!("unknown STTrueFalseBlank value: {}", s)),
        }
    }
}

pub type STUnsignedDecimalNumber = u64;

pub type STTwipsMeasure = String;

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum VerticalAlignRun {
    #[serde(rename = "baseline")]
    Baseline,
    #[serde(rename = "superscript")]
    Superscript,
    #[serde(rename = "subscript")]
    Subscript,
}

impl std::fmt::Display for VerticalAlignRun {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::Baseline => write!(f, "baseline"),
            Self::Superscript => write!(f, "superscript"),
            Self::Subscript => write!(f, "subscript"),
        }
    }
}

impl std::str::FromStr for VerticalAlignRun {
    type Err = String;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        match s {
            "baseline" => Ok(Self::Baseline),
            "superscript" => Ok(Self::Superscript),
            "subscript" => Ok(Self::Subscript),
            _ => Err(format!("unknown VerticalAlignRun value: {}", s)),
        }
    }
}

pub type XmlString = String;

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum STXAlign {
    #[serde(rename = "left")]
    Left,
    #[serde(rename = "center")]
    Center,
    #[serde(rename = "right")]
    Right,
    #[serde(rename = "inside")]
    Inside,
    #[serde(rename = "outside")]
    Outside,
}

impl std::fmt::Display for STXAlign {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::Left => write!(f, "left"),
            Self::Center => write!(f, "center"),
            Self::Right => write!(f, "right"),
            Self::Inside => write!(f, "inside"),
            Self::Outside => write!(f, "outside"),
        }
    }
}

impl std::str::FromStr for STXAlign {
    type Err = String;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        match s {
            "left" => Ok(Self::Left),
            "center" => Ok(Self::Center),
            "right" => Ok(Self::Right),
            "inside" => Ok(Self::Inside),
            "outside" => Ok(Self::Outside),
            _ => Err(format!("unknown STXAlign value: {}", s)),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum STYAlign {
    #[serde(rename = "inline")]
    Inline,
    #[serde(rename = "top")]
    Top,
    #[serde(rename = "center")]
    Center,
    #[serde(rename = "bottom")]
    Bottom,
    #[serde(rename = "inside")]
    Inside,
    #[serde(rename = "outside")]
    Outside,
}

impl std::fmt::Display for STYAlign {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::Inline => write!(f, "inline"),
            Self::Top => write!(f, "top"),
            Self::Center => write!(f, "center"),
            Self::Bottom => write!(f, "bottom"),
            Self::Inside => write!(f, "inside"),
            Self::Outside => write!(f, "outside"),
        }
    }
}

impl std::str::FromStr for STYAlign {
    type Err = String;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        match s {
            "inline" => Ok(Self::Inline),
            "top" => Ok(Self::Top),
            "center" => Ok(Self::Center),
            "bottom" => Ok(Self::Bottom),
            "inside" => Ok(Self::Inside),
            "outside" => Ok(Self::Outside),
            _ => Err(format!("unknown STYAlign value: {}", s)),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum STConformanceClass {
    #[serde(rename = "strict")]
    Strict,
    #[serde(rename = "transitional")]
    Transitional,
}

impl std::fmt::Display for STConformanceClass {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::Strict => write!(f, "strict"),
            Self::Transitional => write!(f, "transitional"),
        }
    }
}

impl std::str::FromStr for STConformanceClass {
    type Err = String;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        match s {
            "strict" => Ok(Self::Strict),
            "transitional" => Ok(Self::Transitional),
            _ => Err(format!("unknown STConformanceClass value: {}", s)),
        }
    }
}

pub type STUniversalMeasure = String;

pub type STPositiveUniversalMeasure = String;

pub type STPercentage = String;

pub type STFixedPercentage = String;

pub type STPositivePercentage = String;

pub type STPositiveFixedPercentage = String;

pub type STRelationshipId = String;

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum FilterOperator {
    #[serde(rename = "equal")]
    Equal,
    #[serde(rename = "lessThan")]
    LessThan,
    #[serde(rename = "lessThanOrEqual")]
    LessThanOrEqual,
    #[serde(rename = "notEqual")]
    NotEqual,
    #[serde(rename = "greaterThanOrEqual")]
    GreaterThanOrEqual,
    #[serde(rename = "greaterThan")]
    GreaterThan,
}

impl std::fmt::Display for FilterOperator {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::Equal => write!(f, "equal"),
            Self::LessThan => write!(f, "lessThan"),
            Self::LessThanOrEqual => write!(f, "lessThanOrEqual"),
            Self::NotEqual => write!(f, "notEqual"),
            Self::GreaterThanOrEqual => write!(f, "greaterThanOrEqual"),
            Self::GreaterThan => write!(f, "greaterThan"),
        }
    }
}

impl std::str::FromStr for FilterOperator {
    type Err = String;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        match s {
            "equal" => Ok(Self::Equal),
            "lessThan" => Ok(Self::LessThan),
            "lessThanOrEqual" => Ok(Self::LessThanOrEqual),
            "notEqual" => Ok(Self::NotEqual),
            "greaterThanOrEqual" => Ok(Self::GreaterThanOrEqual),
            "greaterThan" => Ok(Self::GreaterThan),
            _ => Err(format!("unknown FilterOperator value: {}", s)),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum DynamicFilterType {
    #[serde(rename = "null")]
    Null,
    #[serde(rename = "aboveAverage")]
    AboveAverage,
    #[serde(rename = "belowAverage")]
    BelowAverage,
    #[serde(rename = "tomorrow")]
    Tomorrow,
    #[serde(rename = "today")]
    Today,
    #[serde(rename = "yesterday")]
    Yesterday,
    #[serde(rename = "nextWeek")]
    NextWeek,
    #[serde(rename = "thisWeek")]
    ThisWeek,
    #[serde(rename = "lastWeek")]
    LastWeek,
    #[serde(rename = "nextMonth")]
    NextMonth,
    #[serde(rename = "thisMonth")]
    ThisMonth,
    #[serde(rename = "lastMonth")]
    LastMonth,
    #[serde(rename = "nextQuarter")]
    NextQuarter,
    #[serde(rename = "thisQuarter")]
    ThisQuarter,
    #[serde(rename = "lastQuarter")]
    LastQuarter,
    #[serde(rename = "nextYear")]
    NextYear,
    #[serde(rename = "thisYear")]
    ThisYear,
    #[serde(rename = "lastYear")]
    LastYear,
    #[serde(rename = "yearToDate")]
    YearToDate,
    #[serde(rename = "Q1")]
    Q1,
    #[serde(rename = "Q2")]
    Q2,
    #[serde(rename = "Q3")]
    Q3,
    #[serde(rename = "Q4")]
    Q4,
    #[serde(rename = "M1")]
    M1,
    #[serde(rename = "M2")]
    M2,
    #[serde(rename = "M3")]
    M3,
    #[serde(rename = "M4")]
    M4,
    #[serde(rename = "M5")]
    M5,
    #[serde(rename = "M6")]
    M6,
    #[serde(rename = "M7")]
    M7,
    #[serde(rename = "M8")]
    M8,
    #[serde(rename = "M9")]
    M9,
    #[serde(rename = "M10")]
    M10,
    #[serde(rename = "M11")]
    M11,
    #[serde(rename = "M12")]
    M12,
}

impl std::fmt::Display for DynamicFilterType {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::Null => write!(f, "null"),
            Self::AboveAverage => write!(f, "aboveAverage"),
            Self::BelowAverage => write!(f, "belowAverage"),
            Self::Tomorrow => write!(f, "tomorrow"),
            Self::Today => write!(f, "today"),
            Self::Yesterday => write!(f, "yesterday"),
            Self::NextWeek => write!(f, "nextWeek"),
            Self::ThisWeek => write!(f, "thisWeek"),
            Self::LastWeek => write!(f, "lastWeek"),
            Self::NextMonth => write!(f, "nextMonth"),
            Self::ThisMonth => write!(f, "thisMonth"),
            Self::LastMonth => write!(f, "lastMonth"),
            Self::NextQuarter => write!(f, "nextQuarter"),
            Self::ThisQuarter => write!(f, "thisQuarter"),
            Self::LastQuarter => write!(f, "lastQuarter"),
            Self::NextYear => write!(f, "nextYear"),
            Self::ThisYear => write!(f, "thisYear"),
            Self::LastYear => write!(f, "lastYear"),
            Self::YearToDate => write!(f, "yearToDate"),
            Self::Q1 => write!(f, "Q1"),
            Self::Q2 => write!(f, "Q2"),
            Self::Q3 => write!(f, "Q3"),
            Self::Q4 => write!(f, "Q4"),
            Self::M1 => write!(f, "M1"),
            Self::M2 => write!(f, "M2"),
            Self::M3 => write!(f, "M3"),
            Self::M4 => write!(f, "M4"),
            Self::M5 => write!(f, "M5"),
            Self::M6 => write!(f, "M6"),
            Self::M7 => write!(f, "M7"),
            Self::M8 => write!(f, "M8"),
            Self::M9 => write!(f, "M9"),
            Self::M10 => write!(f, "M10"),
            Self::M11 => write!(f, "M11"),
            Self::M12 => write!(f, "M12"),
        }
    }
}

impl std::str::FromStr for DynamicFilterType {
    type Err = String;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        match s {
            "null" => Ok(Self::Null),
            "aboveAverage" => Ok(Self::AboveAverage),
            "belowAverage" => Ok(Self::BelowAverage),
            "tomorrow" => Ok(Self::Tomorrow),
            "today" => Ok(Self::Today),
            "yesterday" => Ok(Self::Yesterday),
            "nextWeek" => Ok(Self::NextWeek),
            "thisWeek" => Ok(Self::ThisWeek),
            "lastWeek" => Ok(Self::LastWeek),
            "nextMonth" => Ok(Self::NextMonth),
            "thisMonth" => Ok(Self::ThisMonth),
            "lastMonth" => Ok(Self::LastMonth),
            "nextQuarter" => Ok(Self::NextQuarter),
            "thisQuarter" => Ok(Self::ThisQuarter),
            "lastQuarter" => Ok(Self::LastQuarter),
            "nextYear" => Ok(Self::NextYear),
            "thisYear" => Ok(Self::ThisYear),
            "lastYear" => Ok(Self::LastYear),
            "yearToDate" => Ok(Self::YearToDate),
            "Q1" => Ok(Self::Q1),
            "Q2" => Ok(Self::Q2),
            "Q3" => Ok(Self::Q3),
            "Q4" => Ok(Self::Q4),
            "M1" => Ok(Self::M1),
            "M2" => Ok(Self::M2),
            "M3" => Ok(Self::M3),
            "M4" => Ok(Self::M4),
            "M5" => Ok(Self::M5),
            "M6" => Ok(Self::M6),
            "M7" => Ok(Self::M7),
            "M8" => Ok(Self::M8),
            "M9" => Ok(Self::M9),
            "M10" => Ok(Self::M10),
            "M11" => Ok(Self::M11),
            "M12" => Ok(Self::M12),
            _ => Err(format!("unknown DynamicFilterType value: {}", s)),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum IconSetType {
    #[serde(rename = "3Arrows")]
    _3Arrows,
    #[serde(rename = "3ArrowsGray")]
    _3ArrowsGray,
    #[serde(rename = "3Flags")]
    _3Flags,
    #[serde(rename = "3TrafficLights1")]
    _3TrafficLights1,
    #[serde(rename = "3TrafficLights2")]
    _3TrafficLights2,
    #[serde(rename = "3Signs")]
    _3Signs,
    #[serde(rename = "3Symbols")]
    _3Symbols,
    #[serde(rename = "3Symbols2")]
    _3Symbols2,
    #[serde(rename = "4Arrows")]
    _4Arrows,
    #[serde(rename = "4ArrowsGray")]
    _4ArrowsGray,
    #[serde(rename = "4RedToBlack")]
    _4RedToBlack,
    #[serde(rename = "4Rating")]
    _4Rating,
    #[serde(rename = "4TrafficLights")]
    _4TrafficLights,
    #[serde(rename = "5Arrows")]
    _5Arrows,
    #[serde(rename = "5ArrowsGray")]
    _5ArrowsGray,
    #[serde(rename = "5Rating")]
    _5Rating,
    #[serde(rename = "5Quarters")]
    _5Quarters,
}

impl std::fmt::Display for IconSetType {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::_3Arrows => write!(f, "3Arrows"),
            Self::_3ArrowsGray => write!(f, "3ArrowsGray"),
            Self::_3Flags => write!(f, "3Flags"),
            Self::_3TrafficLights1 => write!(f, "3TrafficLights1"),
            Self::_3TrafficLights2 => write!(f, "3TrafficLights2"),
            Self::_3Signs => write!(f, "3Signs"),
            Self::_3Symbols => write!(f, "3Symbols"),
            Self::_3Symbols2 => write!(f, "3Symbols2"),
            Self::_4Arrows => write!(f, "4Arrows"),
            Self::_4ArrowsGray => write!(f, "4ArrowsGray"),
            Self::_4RedToBlack => write!(f, "4RedToBlack"),
            Self::_4Rating => write!(f, "4Rating"),
            Self::_4TrafficLights => write!(f, "4TrafficLights"),
            Self::_5Arrows => write!(f, "5Arrows"),
            Self::_5ArrowsGray => write!(f, "5ArrowsGray"),
            Self::_5Rating => write!(f, "5Rating"),
            Self::_5Quarters => write!(f, "5Quarters"),
        }
    }
}

impl std::str::FromStr for IconSetType {
    type Err = String;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        match s {
            "3Arrows" => Ok(Self::_3Arrows),
            "3ArrowsGray" => Ok(Self::_3ArrowsGray),
            "3Flags" => Ok(Self::_3Flags),
            "3TrafficLights1" => Ok(Self::_3TrafficLights1),
            "3TrafficLights2" => Ok(Self::_3TrafficLights2),
            "3Signs" => Ok(Self::_3Signs),
            "3Symbols" => Ok(Self::_3Symbols),
            "3Symbols2" => Ok(Self::_3Symbols2),
            "4Arrows" => Ok(Self::_4Arrows),
            "4ArrowsGray" => Ok(Self::_4ArrowsGray),
            "4RedToBlack" => Ok(Self::_4RedToBlack),
            "4Rating" => Ok(Self::_4Rating),
            "4TrafficLights" => Ok(Self::_4TrafficLights),
            "5Arrows" => Ok(Self::_5Arrows),
            "5ArrowsGray" => Ok(Self::_5ArrowsGray),
            "5Rating" => Ok(Self::_5Rating),
            "5Quarters" => Ok(Self::_5Quarters),
            _ => Err(format!("unknown IconSetType value: {}", s)),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum SortBy {
    #[serde(rename = "value")]
    Value,
    #[serde(rename = "cellColor")]
    CellColor,
    #[serde(rename = "fontColor")]
    FontColor,
    #[serde(rename = "icon")]
    Icon,
}

impl std::fmt::Display for SortBy {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::Value => write!(f, "value"),
            Self::CellColor => write!(f, "cellColor"),
            Self::FontColor => write!(f, "fontColor"),
            Self::Icon => write!(f, "icon"),
        }
    }
}

impl std::str::FromStr for SortBy {
    type Err = String;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        match s {
            "value" => Ok(Self::Value),
            "cellColor" => Ok(Self::CellColor),
            "fontColor" => Ok(Self::FontColor),
            "icon" => Ok(Self::Icon),
            _ => Err(format!("unknown SortBy value: {}", s)),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum SortMethod {
    #[serde(rename = "stroke")]
    Stroke,
    #[serde(rename = "pinYin")]
    PinYin,
    #[serde(rename = "none")]
    None,
}

impl std::fmt::Display for SortMethod {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::Stroke => write!(f, "stroke"),
            Self::PinYin => write!(f, "pinYin"),
            Self::None => write!(f, "none"),
        }
    }
}

impl std::str::FromStr for SortMethod {
    type Err = String;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        match s {
            "stroke" => Ok(Self::Stroke),
            "pinYin" => Ok(Self::PinYin),
            "none" => Ok(Self::None),
            _ => Err(format!("unknown SortMethod value: {}", s)),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum STDateTimeGrouping {
    #[serde(rename = "year")]
    Year,
    #[serde(rename = "month")]
    Month,
    #[serde(rename = "day")]
    Day,
    #[serde(rename = "hour")]
    Hour,
    #[serde(rename = "minute")]
    Minute,
    #[serde(rename = "second")]
    Second,
}

impl std::fmt::Display for STDateTimeGrouping {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::Year => write!(f, "year"),
            Self::Month => write!(f, "month"),
            Self::Day => write!(f, "day"),
            Self::Hour => write!(f, "hour"),
            Self::Minute => write!(f, "minute"),
            Self::Second => write!(f, "second"),
        }
    }
}

impl std::str::FromStr for STDateTimeGrouping {
    type Err = String;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        match s {
            "year" => Ok(Self::Year),
            "month" => Ok(Self::Month),
            "day" => Ok(Self::Day),
            "hour" => Ok(Self::Hour),
            "minute" => Ok(Self::Minute),
            "second" => Ok(Self::Second),
            _ => Err(format!("unknown STDateTimeGrouping value: {}", s)),
        }
    }
}

pub type CellRef = String;

pub type Reference = String;

pub type STRefA = String;

pub type SquareRef = String;

pub type STFormula = XmlString;

pub type STUnsignedIntHex = Vec<u8>;

pub type STUnsignedShortHex = Vec<u8>;

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum STTextHAlign {
    #[serde(rename = "left")]
    Left,
    #[serde(rename = "center")]
    Center,
    #[serde(rename = "right")]
    Right,
    #[serde(rename = "justify")]
    Justify,
    #[serde(rename = "distributed")]
    Distributed,
}

impl std::fmt::Display for STTextHAlign {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::Left => write!(f, "left"),
            Self::Center => write!(f, "center"),
            Self::Right => write!(f, "right"),
            Self::Justify => write!(f, "justify"),
            Self::Distributed => write!(f, "distributed"),
        }
    }
}

impl std::str::FromStr for STTextHAlign {
    type Err = String;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        match s {
            "left" => Ok(Self::Left),
            "center" => Ok(Self::Center),
            "right" => Ok(Self::Right),
            "justify" => Ok(Self::Justify),
            "distributed" => Ok(Self::Distributed),
            _ => Err(format!("unknown STTextHAlign value: {}", s)),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum STTextVAlign {
    #[serde(rename = "top")]
    Top,
    #[serde(rename = "center")]
    Center,
    #[serde(rename = "bottom")]
    Bottom,
    #[serde(rename = "justify")]
    Justify,
    #[serde(rename = "distributed")]
    Distributed,
}

impl std::fmt::Display for STTextVAlign {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::Top => write!(f, "top"),
            Self::Center => write!(f, "center"),
            Self::Bottom => write!(f, "bottom"),
            Self::Justify => write!(f, "justify"),
            Self::Distributed => write!(f, "distributed"),
        }
    }
}

impl std::str::FromStr for STTextVAlign {
    type Err = String;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        match s {
            "top" => Ok(Self::Top),
            "center" => Ok(Self::Center),
            "bottom" => Ok(Self::Bottom),
            "justify" => Ok(Self::Justify),
            "distributed" => Ok(Self::Distributed),
            _ => Err(format!("unknown STTextVAlign value: {}", s)),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum STCredMethod {
    #[serde(rename = "integrated")]
    Integrated,
    #[serde(rename = "none")]
    None,
    #[serde(rename = "stored")]
    Stored,
    #[serde(rename = "prompt")]
    Prompt,
}

impl std::fmt::Display for STCredMethod {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::Integrated => write!(f, "integrated"),
            Self::None => write!(f, "none"),
            Self::Stored => write!(f, "stored"),
            Self::Prompt => write!(f, "prompt"),
        }
    }
}

impl std::str::FromStr for STCredMethod {
    type Err = String;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        match s {
            "integrated" => Ok(Self::Integrated),
            "none" => Ok(Self::None),
            "stored" => Ok(Self::Stored),
            "prompt" => Ok(Self::Prompt),
            _ => Err(format!("unknown STCredMethod value: {}", s)),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum STHtmlFmt {
    #[serde(rename = "none")]
    None,
    #[serde(rename = "rtf")]
    Rtf,
    #[serde(rename = "all")]
    All,
}

impl std::fmt::Display for STHtmlFmt {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::None => write!(f, "none"),
            Self::Rtf => write!(f, "rtf"),
            Self::All => write!(f, "all"),
        }
    }
}

impl std::str::FromStr for STHtmlFmt {
    type Err = String;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        match s {
            "none" => Ok(Self::None),
            "rtf" => Ok(Self::Rtf),
            "all" => Ok(Self::All),
            _ => Err(format!("unknown STHtmlFmt value: {}", s)),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum STParameterType {
    #[serde(rename = "prompt")]
    Prompt,
    #[serde(rename = "value")]
    Value,
    #[serde(rename = "cell")]
    Cell,
}

impl std::fmt::Display for STParameterType {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::Prompt => write!(f, "prompt"),
            Self::Value => write!(f, "value"),
            Self::Cell => write!(f, "cell"),
        }
    }
}

impl std::str::FromStr for STParameterType {
    type Err = String;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        match s {
            "prompt" => Ok(Self::Prompt),
            "value" => Ok(Self::Value),
            "cell" => Ok(Self::Cell),
            _ => Err(format!("unknown STParameterType value: {}", s)),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum STFileType {
    #[serde(rename = "mac")]
    Mac,
    #[serde(rename = "win")]
    Win,
    #[serde(rename = "dos")]
    Dos,
    #[serde(rename = "lin")]
    Lin,
    #[serde(rename = "other")]
    Other,
}

impl std::fmt::Display for STFileType {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::Mac => write!(f, "mac"),
            Self::Win => write!(f, "win"),
            Self::Dos => write!(f, "dos"),
            Self::Lin => write!(f, "lin"),
            Self::Other => write!(f, "other"),
        }
    }
}

impl std::str::FromStr for STFileType {
    type Err = String;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        match s {
            "mac" => Ok(Self::Mac),
            "win" => Ok(Self::Win),
            "dos" => Ok(Self::Dos),
            "lin" => Ok(Self::Lin),
            "other" => Ok(Self::Other),
            _ => Err(format!("unknown STFileType value: {}", s)),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum STQualifier {
    #[serde(rename = "doubleQuote")]
    DoubleQuote,
    #[serde(rename = "singleQuote")]
    SingleQuote,
    #[serde(rename = "none")]
    None,
}

impl std::fmt::Display for STQualifier {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::DoubleQuote => write!(f, "doubleQuote"),
            Self::SingleQuote => write!(f, "singleQuote"),
            Self::None => write!(f, "none"),
        }
    }
}

impl std::str::FromStr for STQualifier {
    type Err = String;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        match s {
            "doubleQuote" => Ok(Self::DoubleQuote),
            "singleQuote" => Ok(Self::SingleQuote),
            "none" => Ok(Self::None),
            _ => Err(format!("unknown STQualifier value: {}", s)),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum STExternalConnectionType {
    #[serde(rename = "general")]
    General,
    #[serde(rename = "text")]
    Text,
    #[serde(rename = "MDY")]
    MDY,
    #[serde(rename = "DMY")]
    DMY,
    #[serde(rename = "YMD")]
    YMD,
    #[serde(rename = "MYD")]
    MYD,
    #[serde(rename = "DYM")]
    DYM,
    #[serde(rename = "YDM")]
    YDM,
    #[serde(rename = "skip")]
    Skip,
    #[serde(rename = "EMD")]
    EMD,
}

impl std::fmt::Display for STExternalConnectionType {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::General => write!(f, "general"),
            Self::Text => write!(f, "text"),
            Self::MDY => write!(f, "MDY"),
            Self::DMY => write!(f, "DMY"),
            Self::YMD => write!(f, "YMD"),
            Self::MYD => write!(f, "MYD"),
            Self::DYM => write!(f, "DYM"),
            Self::YDM => write!(f, "YDM"),
            Self::Skip => write!(f, "skip"),
            Self::EMD => write!(f, "EMD"),
        }
    }
}

impl std::str::FromStr for STExternalConnectionType {
    type Err = String;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        match s {
            "general" => Ok(Self::General),
            "text" => Ok(Self::Text),
            "MDY" => Ok(Self::MDY),
            "DMY" => Ok(Self::DMY),
            "YMD" => Ok(Self::YMD),
            "MYD" => Ok(Self::MYD),
            "DYM" => Ok(Self::DYM),
            "YDM" => Ok(Self::YDM),
            "skip" => Ok(Self::Skip),
            "EMD" => Ok(Self::EMD),
            _ => Err(format!("unknown STExternalConnectionType value: {}", s)),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum STSourceType {
    #[serde(rename = "worksheet")]
    Worksheet,
    #[serde(rename = "external")]
    External,
    #[serde(rename = "consolidation")]
    Consolidation,
    #[serde(rename = "scenario")]
    Scenario,
}

impl std::fmt::Display for STSourceType {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::Worksheet => write!(f, "worksheet"),
            Self::External => write!(f, "external"),
            Self::Consolidation => write!(f, "consolidation"),
            Self::Scenario => write!(f, "scenario"),
        }
    }
}

impl std::str::FromStr for STSourceType {
    type Err = String;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        match s {
            "worksheet" => Ok(Self::Worksheet),
            "external" => Ok(Self::External),
            "consolidation" => Ok(Self::Consolidation),
            "scenario" => Ok(Self::Scenario),
            _ => Err(format!("unknown STSourceType value: {}", s)),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum STGroupBy {
    #[serde(rename = "range")]
    Range,
    #[serde(rename = "seconds")]
    Seconds,
    #[serde(rename = "minutes")]
    Minutes,
    #[serde(rename = "hours")]
    Hours,
    #[serde(rename = "days")]
    Days,
    #[serde(rename = "months")]
    Months,
    #[serde(rename = "quarters")]
    Quarters,
    #[serde(rename = "years")]
    Years,
}

impl std::fmt::Display for STGroupBy {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::Range => write!(f, "range"),
            Self::Seconds => write!(f, "seconds"),
            Self::Minutes => write!(f, "minutes"),
            Self::Hours => write!(f, "hours"),
            Self::Days => write!(f, "days"),
            Self::Months => write!(f, "months"),
            Self::Quarters => write!(f, "quarters"),
            Self::Years => write!(f, "years"),
        }
    }
}

impl std::str::FromStr for STGroupBy {
    type Err = String;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        match s {
            "range" => Ok(Self::Range),
            "seconds" => Ok(Self::Seconds),
            "minutes" => Ok(Self::Minutes),
            "hours" => Ok(Self::Hours),
            "days" => Ok(Self::Days),
            "months" => Ok(Self::Months),
            "quarters" => Ok(Self::Quarters),
            "years" => Ok(Self::Years),
            _ => Err(format!("unknown STGroupBy value: {}", s)),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum STSortType {
    #[serde(rename = "none")]
    None,
    #[serde(rename = "ascending")]
    Ascending,
    #[serde(rename = "descending")]
    Descending,
    #[serde(rename = "ascendingAlpha")]
    AscendingAlpha,
    #[serde(rename = "descendingAlpha")]
    DescendingAlpha,
    #[serde(rename = "ascendingNatural")]
    AscendingNatural,
    #[serde(rename = "descendingNatural")]
    DescendingNatural,
}

impl std::fmt::Display for STSortType {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::None => write!(f, "none"),
            Self::Ascending => write!(f, "ascending"),
            Self::Descending => write!(f, "descending"),
            Self::AscendingAlpha => write!(f, "ascendingAlpha"),
            Self::DescendingAlpha => write!(f, "descendingAlpha"),
            Self::AscendingNatural => write!(f, "ascendingNatural"),
            Self::DescendingNatural => write!(f, "descendingNatural"),
        }
    }
}

impl std::str::FromStr for STSortType {
    type Err = String;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        match s {
            "none" => Ok(Self::None),
            "ascending" => Ok(Self::Ascending),
            "descending" => Ok(Self::Descending),
            "ascendingAlpha" => Ok(Self::AscendingAlpha),
            "descendingAlpha" => Ok(Self::DescendingAlpha),
            "ascendingNatural" => Ok(Self::AscendingNatural),
            "descendingNatural" => Ok(Self::DescendingNatural),
            _ => Err(format!("unknown STSortType value: {}", s)),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum STScope {
    #[serde(rename = "selection")]
    Selection,
    #[serde(rename = "data")]
    Data,
    #[serde(rename = "field")]
    Field,
}

impl std::fmt::Display for STScope {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::Selection => write!(f, "selection"),
            Self::Data => write!(f, "data"),
            Self::Field => write!(f, "field"),
        }
    }
}

impl std::str::FromStr for STScope {
    type Err = String;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        match s {
            "selection" => Ok(Self::Selection),
            "data" => Ok(Self::Data),
            "field" => Ok(Self::Field),
            _ => Err(format!("unknown STScope value: {}", s)),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum STType {
    #[serde(rename = "none")]
    None,
    #[serde(rename = "all")]
    All,
    #[serde(rename = "row")]
    Row,
    #[serde(rename = "column")]
    Column,
}

impl std::fmt::Display for STType {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::None => write!(f, "none"),
            Self::All => write!(f, "all"),
            Self::Row => write!(f, "row"),
            Self::Column => write!(f, "column"),
        }
    }
}

impl std::str::FromStr for STType {
    type Err = String;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        match s {
            "none" => Ok(Self::None),
            "all" => Ok(Self::All),
            "row" => Ok(Self::Row),
            "column" => Ok(Self::Column),
            _ => Err(format!("unknown STType value: {}", s)),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum STShowDataAs {
    #[serde(rename = "normal")]
    Normal,
    #[serde(rename = "difference")]
    Difference,
    #[serde(rename = "percent")]
    Percent,
    #[serde(rename = "percentDiff")]
    PercentDiff,
    #[serde(rename = "runTotal")]
    RunTotal,
    #[serde(rename = "percentOfRow")]
    PercentOfRow,
    #[serde(rename = "percentOfCol")]
    PercentOfCol,
    #[serde(rename = "percentOfTotal")]
    PercentOfTotal,
    #[serde(rename = "index")]
    Index,
}

impl std::fmt::Display for STShowDataAs {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::Normal => write!(f, "normal"),
            Self::Difference => write!(f, "difference"),
            Self::Percent => write!(f, "percent"),
            Self::PercentDiff => write!(f, "percentDiff"),
            Self::RunTotal => write!(f, "runTotal"),
            Self::PercentOfRow => write!(f, "percentOfRow"),
            Self::PercentOfCol => write!(f, "percentOfCol"),
            Self::PercentOfTotal => write!(f, "percentOfTotal"),
            Self::Index => write!(f, "index"),
        }
    }
}

impl std::str::FromStr for STShowDataAs {
    type Err = String;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        match s {
            "normal" => Ok(Self::Normal),
            "difference" => Ok(Self::Difference),
            "percent" => Ok(Self::Percent),
            "percentDiff" => Ok(Self::PercentDiff),
            "runTotal" => Ok(Self::RunTotal),
            "percentOfRow" => Ok(Self::PercentOfRow),
            "percentOfCol" => Ok(Self::PercentOfCol),
            "percentOfTotal" => Ok(Self::PercentOfTotal),
            "index" => Ok(Self::Index),
            _ => Err(format!("unknown STShowDataAs value: {}", s)),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum STItemType {
    #[serde(rename = "data")]
    Data,
    #[serde(rename = "default")]
    Default,
    #[serde(rename = "sum")]
    Sum,
    #[serde(rename = "countA")]
    CountA,
    #[serde(rename = "avg")]
    Avg,
    #[serde(rename = "max")]
    Max,
    #[serde(rename = "min")]
    Min,
    #[serde(rename = "product")]
    Product,
    #[serde(rename = "count")]
    Count,
    #[serde(rename = "stdDev")]
    StdDev,
    #[serde(rename = "stdDevP")]
    StdDevP,
    #[serde(rename = "var")]
    Var,
    #[serde(rename = "varP")]
    VarP,
    #[serde(rename = "grand")]
    Grand,
    #[serde(rename = "blank")]
    Blank,
}

impl std::fmt::Display for STItemType {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::Data => write!(f, "data"),
            Self::Default => write!(f, "default"),
            Self::Sum => write!(f, "sum"),
            Self::CountA => write!(f, "countA"),
            Self::Avg => write!(f, "avg"),
            Self::Max => write!(f, "max"),
            Self::Min => write!(f, "min"),
            Self::Product => write!(f, "product"),
            Self::Count => write!(f, "count"),
            Self::StdDev => write!(f, "stdDev"),
            Self::StdDevP => write!(f, "stdDevP"),
            Self::Var => write!(f, "var"),
            Self::VarP => write!(f, "varP"),
            Self::Grand => write!(f, "grand"),
            Self::Blank => write!(f, "blank"),
        }
    }
}

impl std::str::FromStr for STItemType {
    type Err = String;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        match s {
            "data" => Ok(Self::Data),
            "default" => Ok(Self::Default),
            "sum" => Ok(Self::Sum),
            "countA" => Ok(Self::CountA),
            "avg" => Ok(Self::Avg),
            "max" => Ok(Self::Max),
            "min" => Ok(Self::Min),
            "product" => Ok(Self::Product),
            "count" => Ok(Self::Count),
            "stdDev" => Ok(Self::StdDev),
            "stdDevP" => Ok(Self::StdDevP),
            "var" => Ok(Self::Var),
            "varP" => Ok(Self::VarP),
            "grand" => Ok(Self::Grand),
            "blank" => Ok(Self::Blank),
            _ => Err(format!("unknown STItemType value: {}", s)),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum STFormatAction {
    #[serde(rename = "blank")]
    Blank,
    #[serde(rename = "formatting")]
    Formatting,
    #[serde(rename = "drill")]
    Drill,
    #[serde(rename = "formula")]
    Formula,
}

impl std::fmt::Display for STFormatAction {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::Blank => write!(f, "blank"),
            Self::Formatting => write!(f, "formatting"),
            Self::Drill => write!(f, "drill"),
            Self::Formula => write!(f, "formula"),
        }
    }
}

impl std::str::FromStr for STFormatAction {
    type Err = String;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        match s {
            "blank" => Ok(Self::Blank),
            "formatting" => Ok(Self::Formatting),
            "drill" => Ok(Self::Drill),
            "formula" => Ok(Self::Formula),
            _ => Err(format!("unknown STFormatAction value: {}", s)),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum STFieldSortType {
    #[serde(rename = "manual")]
    Manual,
    #[serde(rename = "ascending")]
    Ascending,
    #[serde(rename = "descending")]
    Descending,
}

impl std::fmt::Display for STFieldSortType {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::Manual => write!(f, "manual"),
            Self::Ascending => write!(f, "ascending"),
            Self::Descending => write!(f, "descending"),
        }
    }
}

impl std::str::FromStr for STFieldSortType {
    type Err = String;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        match s {
            "manual" => Ok(Self::Manual),
            "ascending" => Ok(Self::Ascending),
            "descending" => Ok(Self::Descending),
            _ => Err(format!("unknown STFieldSortType value: {}", s)),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum STPivotFilterType {
    #[serde(rename = "unknown")]
    Unknown,
    #[serde(rename = "count")]
    Count,
    #[serde(rename = "percent")]
    Percent,
    #[serde(rename = "sum")]
    Sum,
    #[serde(rename = "captionEqual")]
    CaptionEqual,
    #[serde(rename = "captionNotEqual")]
    CaptionNotEqual,
    #[serde(rename = "captionBeginsWith")]
    CaptionBeginsWith,
    #[serde(rename = "captionNotBeginsWith")]
    CaptionNotBeginsWith,
    #[serde(rename = "captionEndsWith")]
    CaptionEndsWith,
    #[serde(rename = "captionNotEndsWith")]
    CaptionNotEndsWith,
    #[serde(rename = "captionContains")]
    CaptionContains,
    #[serde(rename = "captionNotContains")]
    CaptionNotContains,
    #[serde(rename = "captionGreaterThan")]
    CaptionGreaterThan,
    #[serde(rename = "captionGreaterThanOrEqual")]
    CaptionGreaterThanOrEqual,
    #[serde(rename = "captionLessThan")]
    CaptionLessThan,
    #[serde(rename = "captionLessThanOrEqual")]
    CaptionLessThanOrEqual,
    #[serde(rename = "captionBetween")]
    CaptionBetween,
    #[serde(rename = "captionNotBetween")]
    CaptionNotBetween,
    #[serde(rename = "valueEqual")]
    ValueEqual,
    #[serde(rename = "valueNotEqual")]
    ValueNotEqual,
    #[serde(rename = "valueGreaterThan")]
    ValueGreaterThan,
    #[serde(rename = "valueGreaterThanOrEqual")]
    ValueGreaterThanOrEqual,
    #[serde(rename = "valueLessThan")]
    ValueLessThan,
    #[serde(rename = "valueLessThanOrEqual")]
    ValueLessThanOrEqual,
    #[serde(rename = "valueBetween")]
    ValueBetween,
    #[serde(rename = "valueNotBetween")]
    ValueNotBetween,
    #[serde(rename = "dateEqual")]
    DateEqual,
    #[serde(rename = "dateNotEqual")]
    DateNotEqual,
    #[serde(rename = "dateOlderThan")]
    DateOlderThan,
    #[serde(rename = "dateOlderThanOrEqual")]
    DateOlderThanOrEqual,
    #[serde(rename = "dateNewerThan")]
    DateNewerThan,
    #[serde(rename = "dateNewerThanOrEqual")]
    DateNewerThanOrEqual,
    #[serde(rename = "dateBetween")]
    DateBetween,
    #[serde(rename = "dateNotBetween")]
    DateNotBetween,
    #[serde(rename = "tomorrow")]
    Tomorrow,
    #[serde(rename = "today")]
    Today,
    #[serde(rename = "yesterday")]
    Yesterday,
    #[serde(rename = "nextWeek")]
    NextWeek,
    #[serde(rename = "thisWeek")]
    ThisWeek,
    #[serde(rename = "lastWeek")]
    LastWeek,
    #[serde(rename = "nextMonth")]
    NextMonth,
    #[serde(rename = "thisMonth")]
    ThisMonth,
    #[serde(rename = "lastMonth")]
    LastMonth,
    #[serde(rename = "nextQuarter")]
    NextQuarter,
    #[serde(rename = "thisQuarter")]
    ThisQuarter,
    #[serde(rename = "lastQuarter")]
    LastQuarter,
    #[serde(rename = "nextYear")]
    NextYear,
    #[serde(rename = "thisYear")]
    ThisYear,
    #[serde(rename = "lastYear")]
    LastYear,
    #[serde(rename = "yearToDate")]
    YearToDate,
    #[serde(rename = "Q1")]
    Q1,
    #[serde(rename = "Q2")]
    Q2,
    #[serde(rename = "Q3")]
    Q3,
    #[serde(rename = "Q4")]
    Q4,
    #[serde(rename = "M1")]
    M1,
    #[serde(rename = "M2")]
    M2,
    #[serde(rename = "M3")]
    M3,
    #[serde(rename = "M4")]
    M4,
    #[serde(rename = "M5")]
    M5,
    #[serde(rename = "M6")]
    M6,
    #[serde(rename = "M7")]
    M7,
    #[serde(rename = "M8")]
    M8,
    #[serde(rename = "M9")]
    M9,
    #[serde(rename = "M10")]
    M10,
    #[serde(rename = "M11")]
    M11,
    #[serde(rename = "M12")]
    M12,
}

impl std::fmt::Display for STPivotFilterType {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::Unknown => write!(f, "unknown"),
            Self::Count => write!(f, "count"),
            Self::Percent => write!(f, "percent"),
            Self::Sum => write!(f, "sum"),
            Self::CaptionEqual => write!(f, "captionEqual"),
            Self::CaptionNotEqual => write!(f, "captionNotEqual"),
            Self::CaptionBeginsWith => write!(f, "captionBeginsWith"),
            Self::CaptionNotBeginsWith => write!(f, "captionNotBeginsWith"),
            Self::CaptionEndsWith => write!(f, "captionEndsWith"),
            Self::CaptionNotEndsWith => write!(f, "captionNotEndsWith"),
            Self::CaptionContains => write!(f, "captionContains"),
            Self::CaptionNotContains => write!(f, "captionNotContains"),
            Self::CaptionGreaterThan => write!(f, "captionGreaterThan"),
            Self::CaptionGreaterThanOrEqual => write!(f, "captionGreaterThanOrEqual"),
            Self::CaptionLessThan => write!(f, "captionLessThan"),
            Self::CaptionLessThanOrEqual => write!(f, "captionLessThanOrEqual"),
            Self::CaptionBetween => write!(f, "captionBetween"),
            Self::CaptionNotBetween => write!(f, "captionNotBetween"),
            Self::ValueEqual => write!(f, "valueEqual"),
            Self::ValueNotEqual => write!(f, "valueNotEqual"),
            Self::ValueGreaterThan => write!(f, "valueGreaterThan"),
            Self::ValueGreaterThanOrEqual => write!(f, "valueGreaterThanOrEqual"),
            Self::ValueLessThan => write!(f, "valueLessThan"),
            Self::ValueLessThanOrEqual => write!(f, "valueLessThanOrEqual"),
            Self::ValueBetween => write!(f, "valueBetween"),
            Self::ValueNotBetween => write!(f, "valueNotBetween"),
            Self::DateEqual => write!(f, "dateEqual"),
            Self::DateNotEqual => write!(f, "dateNotEqual"),
            Self::DateOlderThan => write!(f, "dateOlderThan"),
            Self::DateOlderThanOrEqual => write!(f, "dateOlderThanOrEqual"),
            Self::DateNewerThan => write!(f, "dateNewerThan"),
            Self::DateNewerThanOrEqual => write!(f, "dateNewerThanOrEqual"),
            Self::DateBetween => write!(f, "dateBetween"),
            Self::DateNotBetween => write!(f, "dateNotBetween"),
            Self::Tomorrow => write!(f, "tomorrow"),
            Self::Today => write!(f, "today"),
            Self::Yesterday => write!(f, "yesterday"),
            Self::NextWeek => write!(f, "nextWeek"),
            Self::ThisWeek => write!(f, "thisWeek"),
            Self::LastWeek => write!(f, "lastWeek"),
            Self::NextMonth => write!(f, "nextMonth"),
            Self::ThisMonth => write!(f, "thisMonth"),
            Self::LastMonth => write!(f, "lastMonth"),
            Self::NextQuarter => write!(f, "nextQuarter"),
            Self::ThisQuarter => write!(f, "thisQuarter"),
            Self::LastQuarter => write!(f, "lastQuarter"),
            Self::NextYear => write!(f, "nextYear"),
            Self::ThisYear => write!(f, "thisYear"),
            Self::LastYear => write!(f, "lastYear"),
            Self::YearToDate => write!(f, "yearToDate"),
            Self::Q1 => write!(f, "Q1"),
            Self::Q2 => write!(f, "Q2"),
            Self::Q3 => write!(f, "Q3"),
            Self::Q4 => write!(f, "Q4"),
            Self::M1 => write!(f, "M1"),
            Self::M2 => write!(f, "M2"),
            Self::M3 => write!(f, "M3"),
            Self::M4 => write!(f, "M4"),
            Self::M5 => write!(f, "M5"),
            Self::M6 => write!(f, "M6"),
            Self::M7 => write!(f, "M7"),
            Self::M8 => write!(f, "M8"),
            Self::M9 => write!(f, "M9"),
            Self::M10 => write!(f, "M10"),
            Self::M11 => write!(f, "M11"),
            Self::M12 => write!(f, "M12"),
        }
    }
}

impl std::str::FromStr for STPivotFilterType {
    type Err = String;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        match s {
            "unknown" => Ok(Self::Unknown),
            "count" => Ok(Self::Count),
            "percent" => Ok(Self::Percent),
            "sum" => Ok(Self::Sum),
            "captionEqual" => Ok(Self::CaptionEqual),
            "captionNotEqual" => Ok(Self::CaptionNotEqual),
            "captionBeginsWith" => Ok(Self::CaptionBeginsWith),
            "captionNotBeginsWith" => Ok(Self::CaptionNotBeginsWith),
            "captionEndsWith" => Ok(Self::CaptionEndsWith),
            "captionNotEndsWith" => Ok(Self::CaptionNotEndsWith),
            "captionContains" => Ok(Self::CaptionContains),
            "captionNotContains" => Ok(Self::CaptionNotContains),
            "captionGreaterThan" => Ok(Self::CaptionGreaterThan),
            "captionGreaterThanOrEqual" => Ok(Self::CaptionGreaterThanOrEqual),
            "captionLessThan" => Ok(Self::CaptionLessThan),
            "captionLessThanOrEqual" => Ok(Self::CaptionLessThanOrEqual),
            "captionBetween" => Ok(Self::CaptionBetween),
            "captionNotBetween" => Ok(Self::CaptionNotBetween),
            "valueEqual" => Ok(Self::ValueEqual),
            "valueNotEqual" => Ok(Self::ValueNotEqual),
            "valueGreaterThan" => Ok(Self::ValueGreaterThan),
            "valueGreaterThanOrEqual" => Ok(Self::ValueGreaterThanOrEqual),
            "valueLessThan" => Ok(Self::ValueLessThan),
            "valueLessThanOrEqual" => Ok(Self::ValueLessThanOrEqual),
            "valueBetween" => Ok(Self::ValueBetween),
            "valueNotBetween" => Ok(Self::ValueNotBetween),
            "dateEqual" => Ok(Self::DateEqual),
            "dateNotEqual" => Ok(Self::DateNotEqual),
            "dateOlderThan" => Ok(Self::DateOlderThan),
            "dateOlderThanOrEqual" => Ok(Self::DateOlderThanOrEqual),
            "dateNewerThan" => Ok(Self::DateNewerThan),
            "dateNewerThanOrEqual" => Ok(Self::DateNewerThanOrEqual),
            "dateBetween" => Ok(Self::DateBetween),
            "dateNotBetween" => Ok(Self::DateNotBetween),
            "tomorrow" => Ok(Self::Tomorrow),
            "today" => Ok(Self::Today),
            "yesterday" => Ok(Self::Yesterday),
            "nextWeek" => Ok(Self::NextWeek),
            "thisWeek" => Ok(Self::ThisWeek),
            "lastWeek" => Ok(Self::LastWeek),
            "nextMonth" => Ok(Self::NextMonth),
            "thisMonth" => Ok(Self::ThisMonth),
            "lastMonth" => Ok(Self::LastMonth),
            "nextQuarter" => Ok(Self::NextQuarter),
            "thisQuarter" => Ok(Self::ThisQuarter),
            "lastQuarter" => Ok(Self::LastQuarter),
            "nextYear" => Ok(Self::NextYear),
            "thisYear" => Ok(Self::ThisYear),
            "lastYear" => Ok(Self::LastYear),
            "yearToDate" => Ok(Self::YearToDate),
            "Q1" => Ok(Self::Q1),
            "Q2" => Ok(Self::Q2),
            "Q3" => Ok(Self::Q3),
            "Q4" => Ok(Self::Q4),
            "M1" => Ok(Self::M1),
            "M2" => Ok(Self::M2),
            "M3" => Ok(Self::M3),
            "M4" => Ok(Self::M4),
            "M5" => Ok(Self::M5),
            "M6" => Ok(Self::M6),
            "M7" => Ok(Self::M7),
            "M8" => Ok(Self::M8),
            "M9" => Ok(Self::M9),
            "M10" => Ok(Self::M10),
            "M11" => Ok(Self::M11),
            "M12" => Ok(Self::M12),
            _ => Err(format!("unknown STPivotFilterType value: {}", s)),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum STPivotAreaType {
    #[serde(rename = "none")]
    None,
    #[serde(rename = "normal")]
    Normal,
    #[serde(rename = "data")]
    Data,
    #[serde(rename = "all")]
    All,
    #[serde(rename = "origin")]
    Origin,
    #[serde(rename = "button")]
    Button,
    #[serde(rename = "topEnd")]
    TopEnd,
    #[serde(rename = "topRight")]
    TopRight,
}

impl std::fmt::Display for STPivotAreaType {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::None => write!(f, "none"),
            Self::Normal => write!(f, "normal"),
            Self::Data => write!(f, "data"),
            Self::All => write!(f, "all"),
            Self::Origin => write!(f, "origin"),
            Self::Button => write!(f, "button"),
            Self::TopEnd => write!(f, "topEnd"),
            Self::TopRight => write!(f, "topRight"),
        }
    }
}

impl std::str::FromStr for STPivotAreaType {
    type Err = String;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        match s {
            "none" => Ok(Self::None),
            "normal" => Ok(Self::Normal),
            "data" => Ok(Self::Data),
            "all" => Ok(Self::All),
            "origin" => Ok(Self::Origin),
            "button" => Ok(Self::Button),
            "topEnd" => Ok(Self::TopEnd),
            "topRight" => Ok(Self::TopRight),
            _ => Err(format!("unknown STPivotAreaType value: {}", s)),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum STAxis {
    #[serde(rename = "axisRow")]
    AxisRow,
    #[serde(rename = "axisCol")]
    AxisCol,
    #[serde(rename = "axisPage")]
    AxisPage,
    #[serde(rename = "axisValues")]
    AxisValues,
}

impl std::fmt::Display for STAxis {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::AxisRow => write!(f, "axisRow"),
            Self::AxisCol => write!(f, "axisCol"),
            Self::AxisPage => write!(f, "axisPage"),
            Self::AxisValues => write!(f, "axisValues"),
        }
    }
}

impl std::str::FromStr for STAxis {
    type Err = String;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        match s {
            "axisRow" => Ok(Self::AxisRow),
            "axisCol" => Ok(Self::AxisCol),
            "axisPage" => Ok(Self::AxisPage),
            "axisValues" => Ok(Self::AxisValues),
            _ => Err(format!("unknown STAxis value: {}", s)),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum STGrowShrinkType {
    #[serde(rename = "insertDelete")]
    InsertDelete,
    #[serde(rename = "insertClear")]
    InsertClear,
    #[serde(rename = "overwriteClear")]
    OverwriteClear,
}

impl std::fmt::Display for STGrowShrinkType {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::InsertDelete => write!(f, "insertDelete"),
            Self::InsertClear => write!(f, "insertClear"),
            Self::OverwriteClear => write!(f, "overwriteClear"),
        }
    }
}

impl std::str::FromStr for STGrowShrinkType {
    type Err = String;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        match s {
            "insertDelete" => Ok(Self::InsertDelete),
            "insertClear" => Ok(Self::InsertClear),
            "overwriteClear" => Ok(Self::OverwriteClear),
            _ => Err(format!("unknown STGrowShrinkType value: {}", s)),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum STPhoneticType {
    #[serde(rename = "halfwidthKatakana")]
    HalfwidthKatakana,
    #[serde(rename = "fullwidthKatakana")]
    FullwidthKatakana,
    #[serde(rename = "Hiragana")]
    Hiragana,
    #[serde(rename = "noConversion")]
    NoConversion,
}

impl std::fmt::Display for STPhoneticType {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::HalfwidthKatakana => write!(f, "halfwidthKatakana"),
            Self::FullwidthKatakana => write!(f, "fullwidthKatakana"),
            Self::Hiragana => write!(f, "Hiragana"),
            Self::NoConversion => write!(f, "noConversion"),
        }
    }
}

impl std::str::FromStr for STPhoneticType {
    type Err = String;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        match s {
            "halfwidthKatakana" => Ok(Self::HalfwidthKatakana),
            "fullwidthKatakana" => Ok(Self::FullwidthKatakana),
            "Hiragana" => Ok(Self::Hiragana),
            "noConversion" => Ok(Self::NoConversion),
            _ => Err(format!("unknown STPhoneticType value: {}", s)),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum STPhoneticAlignment {
    #[serde(rename = "noControl")]
    NoControl,
    #[serde(rename = "left")]
    Left,
    #[serde(rename = "center")]
    Center,
    #[serde(rename = "distributed")]
    Distributed,
}

impl std::fmt::Display for STPhoneticAlignment {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::NoControl => write!(f, "noControl"),
            Self::Left => write!(f, "left"),
            Self::Center => write!(f, "center"),
            Self::Distributed => write!(f, "distributed"),
        }
    }
}

impl std::str::FromStr for STPhoneticAlignment {
    type Err = String;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        match s {
            "noControl" => Ok(Self::NoControl),
            "left" => Ok(Self::Left),
            "center" => Ok(Self::Center),
            "distributed" => Ok(Self::Distributed),
            _ => Err(format!("unknown STPhoneticAlignment value: {}", s)),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum STRwColActionType {
    #[serde(rename = "insertRow")]
    InsertRow,
    #[serde(rename = "deleteRow")]
    DeleteRow,
    #[serde(rename = "insertCol")]
    InsertCol,
    #[serde(rename = "deleteCol")]
    DeleteCol,
}

impl std::fmt::Display for STRwColActionType {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::InsertRow => write!(f, "insertRow"),
            Self::DeleteRow => write!(f, "deleteRow"),
            Self::InsertCol => write!(f, "insertCol"),
            Self::DeleteCol => write!(f, "deleteCol"),
        }
    }
}

impl std::str::FromStr for STRwColActionType {
    type Err = String;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        match s {
            "insertRow" => Ok(Self::InsertRow),
            "deleteRow" => Ok(Self::DeleteRow),
            "insertCol" => Ok(Self::InsertCol),
            "deleteCol" => Ok(Self::DeleteCol),
            _ => Err(format!("unknown STRwColActionType value: {}", s)),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum STRevisionAction {
    #[serde(rename = "add")]
    Add,
    #[serde(rename = "delete")]
    Delete,
}

impl std::fmt::Display for STRevisionAction {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::Add => write!(f, "add"),
            Self::Delete => write!(f, "delete"),
        }
    }
}

impl std::str::FromStr for STRevisionAction {
    type Err = String;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        match s {
            "add" => Ok(Self::Add),
            "delete" => Ok(Self::Delete),
            _ => Err(format!("unknown STRevisionAction value: {}", s)),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum STFormulaExpression {
    #[serde(rename = "ref")]
    Ref,
    #[serde(rename = "refError")]
    RefError,
    #[serde(rename = "area")]
    Area,
    #[serde(rename = "areaError")]
    AreaError,
    #[serde(rename = "computedArea")]
    ComputedArea,
}

impl std::fmt::Display for STFormulaExpression {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::Ref => write!(f, "ref"),
            Self::RefError => write!(f, "refError"),
            Self::Area => write!(f, "area"),
            Self::AreaError => write!(f, "areaError"),
            Self::ComputedArea => write!(f, "computedArea"),
        }
    }
}

impl std::str::FromStr for STFormulaExpression {
    type Err = String;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        match s {
            "ref" => Ok(Self::Ref),
            "refError" => Ok(Self::RefError),
            "area" => Ok(Self::Area),
            "areaError" => Ok(Self::AreaError),
            "computedArea" => Ok(Self::ComputedArea),
            _ => Err(format!("unknown STFormulaExpression value: {}", s)),
        }
    }
}

pub type STCellSpan = String;

pub type CellSpans = String;

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum CellType {
    #[serde(rename = "b")]
    Boolean,
    #[serde(rename = "n")]
    Number,
    #[serde(rename = "e")]
    Error,
    #[serde(rename = "s")]
    SharedString,
    #[serde(rename = "str")]
    String,
    #[serde(rename = "inlineStr")]
    InlineString,
}

impl std::fmt::Display for CellType {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::Boolean => write!(f, "b"),
            Self::Number => write!(f, "n"),
            Self::Error => write!(f, "e"),
            Self::SharedString => write!(f, "s"),
            Self::String => write!(f, "str"),
            Self::InlineString => write!(f, "inlineStr"),
        }
    }
}

impl std::str::FromStr for CellType {
    type Err = String;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        match s {
            "b" => Ok(Self::Boolean),
            "n" => Ok(Self::Number),
            "e" => Ok(Self::Error),
            "s" => Ok(Self::SharedString),
            "str" => Ok(Self::String),
            "inlineStr" => Ok(Self::InlineString),
            _ => Err(format!("unknown CellType value: {}", s)),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum FormulaType {
    #[serde(rename = "normal")]
    Normal,
    #[serde(rename = "array")]
    Array,
    #[serde(rename = "dataTable")]
    DataTable,
    #[serde(rename = "shared")]
    Shared,
}

impl std::fmt::Display for FormulaType {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::Normal => write!(f, "normal"),
            Self::Array => write!(f, "array"),
            Self::DataTable => write!(f, "dataTable"),
            Self::Shared => write!(f, "shared"),
        }
    }
}

impl std::str::FromStr for FormulaType {
    type Err = String;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        match s {
            "normal" => Ok(Self::Normal),
            "array" => Ok(Self::Array),
            "dataTable" => Ok(Self::DataTable),
            "shared" => Ok(Self::Shared),
            _ => Err(format!("unknown FormulaType value: {}", s)),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum PaneType {
    #[serde(rename = "bottomRight")]
    BottomRight,
    #[serde(rename = "topRight")]
    TopRight,
    #[serde(rename = "bottomLeft")]
    BottomLeft,
    #[serde(rename = "topLeft")]
    TopLeft,
}

impl std::fmt::Display for PaneType {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::BottomRight => write!(f, "bottomRight"),
            Self::TopRight => write!(f, "topRight"),
            Self::BottomLeft => write!(f, "bottomLeft"),
            Self::TopLeft => write!(f, "topLeft"),
        }
    }
}

impl std::str::FromStr for PaneType {
    type Err = String;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        match s {
            "bottomRight" => Ok(Self::BottomRight),
            "topRight" => Ok(Self::TopRight),
            "bottomLeft" => Ok(Self::BottomLeft),
            "topLeft" => Ok(Self::TopLeft),
            _ => Err(format!("unknown PaneType value: {}", s)),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum SheetViewType {
    #[serde(rename = "normal")]
    Normal,
    #[serde(rename = "pageBreakPreview")]
    PageBreakPreview,
    #[serde(rename = "pageLayout")]
    PageLayout,
}

impl std::fmt::Display for SheetViewType {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::Normal => write!(f, "normal"),
            Self::PageBreakPreview => write!(f, "pageBreakPreview"),
            Self::PageLayout => write!(f, "pageLayout"),
        }
    }
}

impl std::str::FromStr for SheetViewType {
    type Err = String;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        match s {
            "normal" => Ok(Self::Normal),
            "pageBreakPreview" => Ok(Self::PageBreakPreview),
            "pageLayout" => Ok(Self::PageLayout),
            _ => Err(format!("unknown SheetViewType value: {}", s)),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum STDataConsolidateFunction {
    #[serde(rename = "average")]
    Average,
    #[serde(rename = "count")]
    Count,
    #[serde(rename = "countNums")]
    CountNums,
    #[serde(rename = "max")]
    Max,
    #[serde(rename = "min")]
    Min,
    #[serde(rename = "product")]
    Product,
    #[serde(rename = "stdDev")]
    StdDev,
    #[serde(rename = "stdDevp")]
    StdDevp,
    #[serde(rename = "sum")]
    Sum,
    #[serde(rename = "var")]
    Var,
    #[serde(rename = "varp")]
    Varp,
}

impl std::fmt::Display for STDataConsolidateFunction {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::Average => write!(f, "average"),
            Self::Count => write!(f, "count"),
            Self::CountNums => write!(f, "countNums"),
            Self::Max => write!(f, "max"),
            Self::Min => write!(f, "min"),
            Self::Product => write!(f, "product"),
            Self::StdDev => write!(f, "stdDev"),
            Self::StdDevp => write!(f, "stdDevp"),
            Self::Sum => write!(f, "sum"),
            Self::Var => write!(f, "var"),
            Self::Varp => write!(f, "varp"),
        }
    }
}

impl std::str::FromStr for STDataConsolidateFunction {
    type Err = String;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        match s {
            "average" => Ok(Self::Average),
            "count" => Ok(Self::Count),
            "countNums" => Ok(Self::CountNums),
            "max" => Ok(Self::Max),
            "min" => Ok(Self::Min),
            "product" => Ok(Self::Product),
            "stdDev" => Ok(Self::StdDev),
            "stdDevp" => Ok(Self::StdDevp),
            "sum" => Ok(Self::Sum),
            "var" => Ok(Self::Var),
            "varp" => Ok(Self::Varp),
            _ => Err(format!("unknown STDataConsolidateFunction value: {}", s)),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum ValidationType {
    #[serde(rename = "none")]
    None,
    #[serde(rename = "whole")]
    Whole,
    #[serde(rename = "decimal")]
    Decimal,
    #[serde(rename = "list")]
    List,
    #[serde(rename = "date")]
    Date,
    #[serde(rename = "time")]
    Time,
    #[serde(rename = "textLength")]
    TextLength,
    #[serde(rename = "custom")]
    Custom,
}

impl std::fmt::Display for ValidationType {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::None => write!(f, "none"),
            Self::Whole => write!(f, "whole"),
            Self::Decimal => write!(f, "decimal"),
            Self::List => write!(f, "list"),
            Self::Date => write!(f, "date"),
            Self::Time => write!(f, "time"),
            Self::TextLength => write!(f, "textLength"),
            Self::Custom => write!(f, "custom"),
        }
    }
}

impl std::str::FromStr for ValidationType {
    type Err = String;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        match s {
            "none" => Ok(Self::None),
            "whole" => Ok(Self::Whole),
            "decimal" => Ok(Self::Decimal),
            "list" => Ok(Self::List),
            "date" => Ok(Self::Date),
            "time" => Ok(Self::Time),
            "textLength" => Ok(Self::TextLength),
            "custom" => Ok(Self::Custom),
            _ => Err(format!("unknown ValidationType value: {}", s)),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum ValidationOperator {
    #[serde(rename = "between")]
    Between,
    #[serde(rename = "notBetween")]
    NotBetween,
    #[serde(rename = "equal")]
    Equal,
    #[serde(rename = "notEqual")]
    NotEqual,
    #[serde(rename = "lessThan")]
    LessThan,
    #[serde(rename = "lessThanOrEqual")]
    LessThanOrEqual,
    #[serde(rename = "greaterThan")]
    GreaterThan,
    #[serde(rename = "greaterThanOrEqual")]
    GreaterThanOrEqual,
}

impl std::fmt::Display for ValidationOperator {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::Between => write!(f, "between"),
            Self::NotBetween => write!(f, "notBetween"),
            Self::Equal => write!(f, "equal"),
            Self::NotEqual => write!(f, "notEqual"),
            Self::LessThan => write!(f, "lessThan"),
            Self::LessThanOrEqual => write!(f, "lessThanOrEqual"),
            Self::GreaterThan => write!(f, "greaterThan"),
            Self::GreaterThanOrEqual => write!(f, "greaterThanOrEqual"),
        }
    }
}

impl std::str::FromStr for ValidationOperator {
    type Err = String;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        match s {
            "between" => Ok(Self::Between),
            "notBetween" => Ok(Self::NotBetween),
            "equal" => Ok(Self::Equal),
            "notEqual" => Ok(Self::NotEqual),
            "lessThan" => Ok(Self::LessThan),
            "lessThanOrEqual" => Ok(Self::LessThanOrEqual),
            "greaterThan" => Ok(Self::GreaterThan),
            "greaterThanOrEqual" => Ok(Self::GreaterThanOrEqual),
            _ => Err(format!("unknown ValidationOperator value: {}", s)),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum ValidationErrorStyle {
    #[serde(rename = "stop")]
    Stop,
    #[serde(rename = "warning")]
    Warning,
    #[serde(rename = "information")]
    Information,
}

impl std::fmt::Display for ValidationErrorStyle {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::Stop => write!(f, "stop"),
            Self::Warning => write!(f, "warning"),
            Self::Information => write!(f, "information"),
        }
    }
}

impl std::str::FromStr for ValidationErrorStyle {
    type Err = String;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        match s {
            "stop" => Ok(Self::Stop),
            "warning" => Ok(Self::Warning),
            "information" => Ok(Self::Information),
            _ => Err(format!("unknown ValidationErrorStyle value: {}", s)),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum STDataValidationImeMode {
    #[serde(rename = "noControl")]
    NoControl,
    #[serde(rename = "off")]
    Off,
    #[serde(rename = "on")]
    On,
    #[serde(rename = "disabled")]
    Disabled,
    #[serde(rename = "hiragana")]
    Hiragana,
    #[serde(rename = "fullKatakana")]
    FullKatakana,
    #[serde(rename = "halfKatakana")]
    HalfKatakana,
    #[serde(rename = "fullAlpha")]
    FullAlpha,
    #[serde(rename = "halfAlpha")]
    HalfAlpha,
    #[serde(rename = "fullHangul")]
    FullHangul,
    #[serde(rename = "halfHangul")]
    HalfHangul,
}

impl std::fmt::Display for STDataValidationImeMode {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::NoControl => write!(f, "noControl"),
            Self::Off => write!(f, "off"),
            Self::On => write!(f, "on"),
            Self::Disabled => write!(f, "disabled"),
            Self::Hiragana => write!(f, "hiragana"),
            Self::FullKatakana => write!(f, "fullKatakana"),
            Self::HalfKatakana => write!(f, "halfKatakana"),
            Self::FullAlpha => write!(f, "fullAlpha"),
            Self::HalfAlpha => write!(f, "halfAlpha"),
            Self::FullHangul => write!(f, "fullHangul"),
            Self::HalfHangul => write!(f, "halfHangul"),
        }
    }
}

impl std::str::FromStr for STDataValidationImeMode {
    type Err = String;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        match s {
            "noControl" => Ok(Self::NoControl),
            "off" => Ok(Self::Off),
            "on" => Ok(Self::On),
            "disabled" => Ok(Self::Disabled),
            "hiragana" => Ok(Self::Hiragana),
            "fullKatakana" => Ok(Self::FullKatakana),
            "halfKatakana" => Ok(Self::HalfKatakana),
            "fullAlpha" => Ok(Self::FullAlpha),
            "halfAlpha" => Ok(Self::HalfAlpha),
            "fullHangul" => Ok(Self::FullHangul),
            "halfHangul" => Ok(Self::HalfHangul),
            _ => Err(format!("unknown STDataValidationImeMode value: {}", s)),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum ConditionalType {
    #[serde(rename = "expression")]
    Expression,
    #[serde(rename = "cellIs")]
    CellIs,
    #[serde(rename = "colorScale")]
    ColorScale,
    #[serde(rename = "dataBar")]
    DataBar,
    #[serde(rename = "iconSet")]
    IconSet,
    #[serde(rename = "top10")]
    Top10,
    #[serde(rename = "uniqueValues")]
    UniqueValues,
    #[serde(rename = "duplicateValues")]
    DuplicateValues,
    #[serde(rename = "containsText")]
    ContainsText,
    #[serde(rename = "notContainsText")]
    NotContainsText,
    #[serde(rename = "beginsWith")]
    BeginsWith,
    #[serde(rename = "endsWith")]
    EndsWith,
    #[serde(rename = "containsBlanks")]
    ContainsBlanks,
    #[serde(rename = "notContainsBlanks")]
    NotContainsBlanks,
    #[serde(rename = "containsErrors")]
    ContainsErrors,
    #[serde(rename = "notContainsErrors")]
    NotContainsErrors,
    #[serde(rename = "timePeriod")]
    TimePeriod,
    #[serde(rename = "aboveAverage")]
    AboveAverage,
}

impl std::fmt::Display for ConditionalType {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::Expression => write!(f, "expression"),
            Self::CellIs => write!(f, "cellIs"),
            Self::ColorScale => write!(f, "colorScale"),
            Self::DataBar => write!(f, "dataBar"),
            Self::IconSet => write!(f, "iconSet"),
            Self::Top10 => write!(f, "top10"),
            Self::UniqueValues => write!(f, "uniqueValues"),
            Self::DuplicateValues => write!(f, "duplicateValues"),
            Self::ContainsText => write!(f, "containsText"),
            Self::NotContainsText => write!(f, "notContainsText"),
            Self::BeginsWith => write!(f, "beginsWith"),
            Self::EndsWith => write!(f, "endsWith"),
            Self::ContainsBlanks => write!(f, "containsBlanks"),
            Self::NotContainsBlanks => write!(f, "notContainsBlanks"),
            Self::ContainsErrors => write!(f, "containsErrors"),
            Self::NotContainsErrors => write!(f, "notContainsErrors"),
            Self::TimePeriod => write!(f, "timePeriod"),
            Self::AboveAverage => write!(f, "aboveAverage"),
        }
    }
}

impl std::str::FromStr for ConditionalType {
    type Err = String;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        match s {
            "expression" => Ok(Self::Expression),
            "cellIs" => Ok(Self::CellIs),
            "colorScale" => Ok(Self::ColorScale),
            "dataBar" => Ok(Self::DataBar),
            "iconSet" => Ok(Self::IconSet),
            "top10" => Ok(Self::Top10),
            "uniqueValues" => Ok(Self::UniqueValues),
            "duplicateValues" => Ok(Self::DuplicateValues),
            "containsText" => Ok(Self::ContainsText),
            "notContainsText" => Ok(Self::NotContainsText),
            "beginsWith" => Ok(Self::BeginsWith),
            "endsWith" => Ok(Self::EndsWith),
            "containsBlanks" => Ok(Self::ContainsBlanks),
            "notContainsBlanks" => Ok(Self::NotContainsBlanks),
            "containsErrors" => Ok(Self::ContainsErrors),
            "notContainsErrors" => Ok(Self::NotContainsErrors),
            "timePeriod" => Ok(Self::TimePeriod),
            "aboveAverage" => Ok(Self::AboveAverage),
            _ => Err(format!("unknown ConditionalType value: {}", s)),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum STTimePeriod {
    #[serde(rename = "today")]
    Today,
    #[serde(rename = "yesterday")]
    Yesterday,
    #[serde(rename = "tomorrow")]
    Tomorrow,
    #[serde(rename = "last7Days")]
    Last7Days,
    #[serde(rename = "thisMonth")]
    ThisMonth,
    #[serde(rename = "lastMonth")]
    LastMonth,
    #[serde(rename = "nextMonth")]
    NextMonth,
    #[serde(rename = "thisWeek")]
    ThisWeek,
    #[serde(rename = "lastWeek")]
    LastWeek,
    #[serde(rename = "nextWeek")]
    NextWeek,
}

impl std::fmt::Display for STTimePeriod {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::Today => write!(f, "today"),
            Self::Yesterday => write!(f, "yesterday"),
            Self::Tomorrow => write!(f, "tomorrow"),
            Self::Last7Days => write!(f, "last7Days"),
            Self::ThisMonth => write!(f, "thisMonth"),
            Self::LastMonth => write!(f, "lastMonth"),
            Self::NextMonth => write!(f, "nextMonth"),
            Self::ThisWeek => write!(f, "thisWeek"),
            Self::LastWeek => write!(f, "lastWeek"),
            Self::NextWeek => write!(f, "nextWeek"),
        }
    }
}

impl std::str::FromStr for STTimePeriod {
    type Err = String;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        match s {
            "today" => Ok(Self::Today),
            "yesterday" => Ok(Self::Yesterday),
            "tomorrow" => Ok(Self::Tomorrow),
            "last7Days" => Ok(Self::Last7Days),
            "thisMonth" => Ok(Self::ThisMonth),
            "lastMonth" => Ok(Self::LastMonth),
            "nextMonth" => Ok(Self::NextMonth),
            "thisWeek" => Ok(Self::ThisWeek),
            "lastWeek" => Ok(Self::LastWeek),
            "nextWeek" => Ok(Self::NextWeek),
            _ => Err(format!("unknown STTimePeriod value: {}", s)),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum ConditionalOperator {
    #[serde(rename = "lessThan")]
    LessThan,
    #[serde(rename = "lessThanOrEqual")]
    LessThanOrEqual,
    #[serde(rename = "equal")]
    Equal,
    #[serde(rename = "notEqual")]
    NotEqual,
    #[serde(rename = "greaterThanOrEqual")]
    GreaterThanOrEqual,
    #[serde(rename = "greaterThan")]
    GreaterThan,
    #[serde(rename = "between")]
    Between,
    #[serde(rename = "notBetween")]
    NotBetween,
    #[serde(rename = "containsText")]
    ContainsText,
    #[serde(rename = "notContains")]
    NotContains,
    #[serde(rename = "beginsWith")]
    BeginsWith,
    #[serde(rename = "endsWith")]
    EndsWith,
}

impl std::fmt::Display for ConditionalOperator {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::LessThan => write!(f, "lessThan"),
            Self::LessThanOrEqual => write!(f, "lessThanOrEqual"),
            Self::Equal => write!(f, "equal"),
            Self::NotEqual => write!(f, "notEqual"),
            Self::GreaterThanOrEqual => write!(f, "greaterThanOrEqual"),
            Self::GreaterThan => write!(f, "greaterThan"),
            Self::Between => write!(f, "between"),
            Self::NotBetween => write!(f, "notBetween"),
            Self::ContainsText => write!(f, "containsText"),
            Self::NotContains => write!(f, "notContains"),
            Self::BeginsWith => write!(f, "beginsWith"),
            Self::EndsWith => write!(f, "endsWith"),
        }
    }
}

impl std::str::FromStr for ConditionalOperator {
    type Err = String;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        match s {
            "lessThan" => Ok(Self::LessThan),
            "lessThanOrEqual" => Ok(Self::LessThanOrEqual),
            "equal" => Ok(Self::Equal),
            "notEqual" => Ok(Self::NotEqual),
            "greaterThanOrEqual" => Ok(Self::GreaterThanOrEqual),
            "greaterThan" => Ok(Self::GreaterThan),
            "between" => Ok(Self::Between),
            "notBetween" => Ok(Self::NotBetween),
            "containsText" => Ok(Self::ContainsText),
            "notContains" => Ok(Self::NotContains),
            "beginsWith" => Ok(Self::BeginsWith),
            "endsWith" => Ok(Self::EndsWith),
            _ => Err(format!("unknown ConditionalOperator value: {}", s)),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum ConditionalValueType {
    #[serde(rename = "num")]
    Num,
    #[serde(rename = "percent")]
    Percent,
    #[serde(rename = "max")]
    Max,
    #[serde(rename = "min")]
    Min,
    #[serde(rename = "formula")]
    Formula,
    #[serde(rename = "percentile")]
    Percentile,
}

impl std::fmt::Display for ConditionalValueType {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::Num => write!(f, "num"),
            Self::Percent => write!(f, "percent"),
            Self::Max => write!(f, "max"),
            Self::Min => write!(f, "min"),
            Self::Formula => write!(f, "formula"),
            Self::Percentile => write!(f, "percentile"),
        }
    }
}

impl std::str::FromStr for ConditionalValueType {
    type Err = String;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        match s {
            "num" => Ok(Self::Num),
            "percent" => Ok(Self::Percent),
            "max" => Ok(Self::Max),
            "min" => Ok(Self::Min),
            "formula" => Ok(Self::Formula),
            "percentile" => Ok(Self::Percentile),
            _ => Err(format!("unknown ConditionalValueType value: {}", s)),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum STPageOrder {
    #[serde(rename = "downThenOver")]
    DownThenOver,
    #[serde(rename = "overThenDown")]
    OverThenDown,
}

impl std::fmt::Display for STPageOrder {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::DownThenOver => write!(f, "downThenOver"),
            Self::OverThenDown => write!(f, "overThenDown"),
        }
    }
}

impl std::str::FromStr for STPageOrder {
    type Err = String;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        match s {
            "downThenOver" => Ok(Self::DownThenOver),
            "overThenDown" => Ok(Self::OverThenDown),
            _ => Err(format!("unknown STPageOrder value: {}", s)),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum STOrientation {
    #[serde(rename = "default")]
    Default,
    #[serde(rename = "portrait")]
    Portrait,
    #[serde(rename = "landscape")]
    Landscape,
}

impl std::fmt::Display for STOrientation {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::Default => write!(f, "default"),
            Self::Portrait => write!(f, "portrait"),
            Self::Landscape => write!(f, "landscape"),
        }
    }
}

impl std::str::FromStr for STOrientation {
    type Err = String;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        match s {
            "default" => Ok(Self::Default),
            "portrait" => Ok(Self::Portrait),
            "landscape" => Ok(Self::Landscape),
            _ => Err(format!("unknown STOrientation value: {}", s)),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum STCellComments {
    #[serde(rename = "none")]
    None,
    #[serde(rename = "asDisplayed")]
    AsDisplayed,
    #[serde(rename = "atEnd")]
    AtEnd,
}

impl std::fmt::Display for STCellComments {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::None => write!(f, "none"),
            Self::AsDisplayed => write!(f, "asDisplayed"),
            Self::AtEnd => write!(f, "atEnd"),
        }
    }
}

impl std::str::FromStr for STCellComments {
    type Err = String;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        match s {
            "none" => Ok(Self::None),
            "asDisplayed" => Ok(Self::AsDisplayed),
            "atEnd" => Ok(Self::AtEnd),
            _ => Err(format!("unknown STCellComments value: {}", s)),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum STPrintError {
    #[serde(rename = "displayed")]
    Displayed,
    #[serde(rename = "blank")]
    Blank,
    #[serde(rename = "dash")]
    Dash,
    #[serde(rename = "NA")]
    NA,
}

impl std::fmt::Display for STPrintError {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::Displayed => write!(f, "displayed"),
            Self::Blank => write!(f, "blank"),
            Self::Dash => write!(f, "dash"),
            Self::NA => write!(f, "NA"),
        }
    }
}

impl std::str::FromStr for STPrintError {
    type Err = String;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        match s {
            "displayed" => Ok(Self::Displayed),
            "blank" => Ok(Self::Blank),
            "dash" => Ok(Self::Dash),
            "NA" => Ok(Self::NA),
            _ => Err(format!("unknown STPrintError value: {}", s)),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum STDvAspect {
    #[serde(rename = "DVASPECT_CONTENT")]
    DVASPECTCONTENT,
    #[serde(rename = "DVASPECT_ICON")]
    DVASPECTICON,
}

impl std::fmt::Display for STDvAspect {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::DVASPECTCONTENT => write!(f, "DVASPECT_CONTENT"),
            Self::DVASPECTICON => write!(f, "DVASPECT_ICON"),
        }
    }
}

impl std::str::FromStr for STDvAspect {
    type Err = String;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        match s {
            "DVASPECT_CONTENT" => Ok(Self::DVASPECTCONTENT),
            "DVASPECT_ICON" => Ok(Self::DVASPECTICON),
            _ => Err(format!("unknown STDvAspect value: {}", s)),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum STOleUpdate {
    #[serde(rename = "OLEUPDATE_ALWAYS")]
    OLEUPDATEALWAYS,
    #[serde(rename = "OLEUPDATE_ONCALL")]
    OLEUPDATEONCALL,
}

impl std::fmt::Display for STOleUpdate {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::OLEUPDATEALWAYS => write!(f, "OLEUPDATE_ALWAYS"),
            Self::OLEUPDATEONCALL => write!(f, "OLEUPDATE_ONCALL"),
        }
    }
}

impl std::str::FromStr for STOleUpdate {
    type Err = String;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        match s {
            "OLEUPDATE_ALWAYS" => Ok(Self::OLEUPDATEALWAYS),
            "OLEUPDATE_ONCALL" => Ok(Self::OLEUPDATEONCALL),
            _ => Err(format!("unknown STOleUpdate value: {}", s)),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum STWebSourceType {
    #[serde(rename = "sheet")]
    Sheet,
    #[serde(rename = "printArea")]
    PrintArea,
    #[serde(rename = "autoFilter")]
    AutoFilter,
    #[serde(rename = "range")]
    Range,
    #[serde(rename = "chart")]
    Chart,
    #[serde(rename = "pivotTable")]
    PivotTable,
    #[serde(rename = "query")]
    Query,
    #[serde(rename = "label")]
    Label,
}

impl std::fmt::Display for STWebSourceType {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::Sheet => write!(f, "sheet"),
            Self::PrintArea => write!(f, "printArea"),
            Self::AutoFilter => write!(f, "autoFilter"),
            Self::Range => write!(f, "range"),
            Self::Chart => write!(f, "chart"),
            Self::PivotTable => write!(f, "pivotTable"),
            Self::Query => write!(f, "query"),
            Self::Label => write!(f, "label"),
        }
    }
}

impl std::str::FromStr for STWebSourceType {
    type Err = String;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        match s {
            "sheet" => Ok(Self::Sheet),
            "printArea" => Ok(Self::PrintArea),
            "autoFilter" => Ok(Self::AutoFilter),
            "range" => Ok(Self::Range),
            "chart" => Ok(Self::Chart),
            "pivotTable" => Ok(Self::PivotTable),
            "query" => Ok(Self::Query),
            "label" => Ok(Self::Label),
            _ => Err(format!("unknown STWebSourceType value: {}", s)),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum PaneState {
    #[serde(rename = "split")]
    Split,
    #[serde(rename = "frozen")]
    Frozen,
    #[serde(rename = "frozenSplit")]
    FrozenSplit,
}

impl std::fmt::Display for PaneState {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::Split => write!(f, "split"),
            Self::Frozen => write!(f, "frozen"),
            Self::FrozenSplit => write!(f, "frozenSplit"),
        }
    }
}

impl std::str::FromStr for PaneState {
    type Err = String;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        match s {
            "split" => Ok(Self::Split),
            "frozen" => Ok(Self::Frozen),
            "frozenSplit" => Ok(Self::FrozenSplit),
            _ => Err(format!("unknown PaneState value: {}", s)),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum STMdxFunctionType {
    #[serde(rename = "m")]
    M,
    #[serde(rename = "v")]
    V,
    #[serde(rename = "s")]
    SharedString,
    #[serde(rename = "c")]
    C,
    #[serde(rename = "r")]
    R,
    #[serde(rename = "p")]
    P,
    #[serde(rename = "k")]
    K,
}

impl std::fmt::Display for STMdxFunctionType {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::M => write!(f, "m"),
            Self::V => write!(f, "v"),
            Self::SharedString => write!(f, "s"),
            Self::C => write!(f, "c"),
            Self::R => write!(f, "r"),
            Self::P => write!(f, "p"),
            Self::K => write!(f, "k"),
        }
    }
}

impl std::str::FromStr for STMdxFunctionType {
    type Err = String;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        match s {
            "m" => Ok(Self::M),
            "v" => Ok(Self::V),
            "s" => Ok(Self::SharedString),
            "c" => Ok(Self::C),
            "r" => Ok(Self::R),
            "p" => Ok(Self::P),
            "k" => Ok(Self::K),
            _ => Err(format!("unknown STMdxFunctionType value: {}", s)),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum STMdxSetOrder {
    #[serde(rename = "u")]
    U,
    #[serde(rename = "a")]
    A,
    #[serde(rename = "d")]
    D,
    #[serde(rename = "aa")]
    Aa,
    #[serde(rename = "ad")]
    Ad,
    #[serde(rename = "na")]
    Na,
    #[serde(rename = "nd")]
    Nd,
}

impl std::fmt::Display for STMdxSetOrder {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::U => write!(f, "u"),
            Self::A => write!(f, "a"),
            Self::D => write!(f, "d"),
            Self::Aa => write!(f, "aa"),
            Self::Ad => write!(f, "ad"),
            Self::Na => write!(f, "na"),
            Self::Nd => write!(f, "nd"),
        }
    }
}

impl std::str::FromStr for STMdxSetOrder {
    type Err = String;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        match s {
            "u" => Ok(Self::U),
            "a" => Ok(Self::A),
            "d" => Ok(Self::D),
            "aa" => Ok(Self::Aa),
            "ad" => Ok(Self::Ad),
            "na" => Ok(Self::Na),
            "nd" => Ok(Self::Nd),
            _ => Err(format!("unknown STMdxSetOrder value: {}", s)),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum STMdxKPIProperty {
    #[serde(rename = "v")]
    V,
    #[serde(rename = "g")]
    G,
    #[serde(rename = "s")]
    SharedString,
    #[serde(rename = "t")]
    T,
    #[serde(rename = "w")]
    W,
    #[serde(rename = "m")]
    M,
}

impl std::fmt::Display for STMdxKPIProperty {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::V => write!(f, "v"),
            Self::G => write!(f, "g"),
            Self::SharedString => write!(f, "s"),
            Self::T => write!(f, "t"),
            Self::W => write!(f, "w"),
            Self::M => write!(f, "m"),
        }
    }
}

impl std::str::FromStr for STMdxKPIProperty {
    type Err = String;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        match s {
            "v" => Ok(Self::V),
            "g" => Ok(Self::G),
            "s" => Ok(Self::SharedString),
            "t" => Ok(Self::T),
            "w" => Ok(Self::W),
            "m" => Ok(Self::M),
            _ => Err(format!("unknown STMdxKPIProperty value: {}", s)),
        }
    }
}

pub type STTextRotation = String;

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum BorderStyle {
    #[serde(rename = "none")]
    None,
    #[serde(rename = "thin")]
    Thin,
    #[serde(rename = "medium")]
    Medium,
    #[serde(rename = "dashed")]
    Dashed,
    #[serde(rename = "dotted")]
    Dotted,
    #[serde(rename = "thick")]
    Thick,
    #[serde(rename = "double")]
    Double,
    #[serde(rename = "hair")]
    Hair,
    #[serde(rename = "mediumDashed")]
    MediumDashed,
    #[serde(rename = "dashDot")]
    DashDot,
    #[serde(rename = "mediumDashDot")]
    MediumDashDot,
    #[serde(rename = "dashDotDot")]
    DashDotDot,
    #[serde(rename = "mediumDashDotDot")]
    MediumDashDotDot,
    #[serde(rename = "slantDashDot")]
    SlantDashDot,
}

impl std::fmt::Display for BorderStyle {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::None => write!(f, "none"),
            Self::Thin => write!(f, "thin"),
            Self::Medium => write!(f, "medium"),
            Self::Dashed => write!(f, "dashed"),
            Self::Dotted => write!(f, "dotted"),
            Self::Thick => write!(f, "thick"),
            Self::Double => write!(f, "double"),
            Self::Hair => write!(f, "hair"),
            Self::MediumDashed => write!(f, "mediumDashed"),
            Self::DashDot => write!(f, "dashDot"),
            Self::MediumDashDot => write!(f, "mediumDashDot"),
            Self::DashDotDot => write!(f, "dashDotDot"),
            Self::MediumDashDotDot => write!(f, "mediumDashDotDot"),
            Self::SlantDashDot => write!(f, "slantDashDot"),
        }
    }
}

impl std::str::FromStr for BorderStyle {
    type Err = String;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        match s {
            "none" => Ok(Self::None),
            "thin" => Ok(Self::Thin),
            "medium" => Ok(Self::Medium),
            "dashed" => Ok(Self::Dashed),
            "dotted" => Ok(Self::Dotted),
            "thick" => Ok(Self::Thick),
            "double" => Ok(Self::Double),
            "hair" => Ok(Self::Hair),
            "mediumDashed" => Ok(Self::MediumDashed),
            "dashDot" => Ok(Self::DashDot),
            "mediumDashDot" => Ok(Self::MediumDashDot),
            "dashDotDot" => Ok(Self::DashDotDot),
            "mediumDashDotDot" => Ok(Self::MediumDashDotDot),
            "slantDashDot" => Ok(Self::SlantDashDot),
            _ => Err(format!("unknown BorderStyle value: {}", s)),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum PatternType {
    #[serde(rename = "none")]
    None,
    #[serde(rename = "solid")]
    Solid,
    #[serde(rename = "mediumGray")]
    MediumGray,
    #[serde(rename = "darkGray")]
    DarkGray,
    #[serde(rename = "lightGray")]
    LightGray,
    #[serde(rename = "darkHorizontal")]
    DarkHorizontal,
    #[serde(rename = "darkVertical")]
    DarkVertical,
    #[serde(rename = "darkDown")]
    DarkDown,
    #[serde(rename = "darkUp")]
    DarkUp,
    #[serde(rename = "darkGrid")]
    DarkGrid,
    #[serde(rename = "darkTrellis")]
    DarkTrellis,
    #[serde(rename = "lightHorizontal")]
    LightHorizontal,
    #[serde(rename = "lightVertical")]
    LightVertical,
    #[serde(rename = "lightDown")]
    LightDown,
    #[serde(rename = "lightUp")]
    LightUp,
    #[serde(rename = "lightGrid")]
    LightGrid,
    #[serde(rename = "lightTrellis")]
    LightTrellis,
    #[serde(rename = "gray125")]
    Gray125,
    #[serde(rename = "gray0625")]
    Gray0625,
}

impl std::fmt::Display for PatternType {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::None => write!(f, "none"),
            Self::Solid => write!(f, "solid"),
            Self::MediumGray => write!(f, "mediumGray"),
            Self::DarkGray => write!(f, "darkGray"),
            Self::LightGray => write!(f, "lightGray"),
            Self::DarkHorizontal => write!(f, "darkHorizontal"),
            Self::DarkVertical => write!(f, "darkVertical"),
            Self::DarkDown => write!(f, "darkDown"),
            Self::DarkUp => write!(f, "darkUp"),
            Self::DarkGrid => write!(f, "darkGrid"),
            Self::DarkTrellis => write!(f, "darkTrellis"),
            Self::LightHorizontal => write!(f, "lightHorizontal"),
            Self::LightVertical => write!(f, "lightVertical"),
            Self::LightDown => write!(f, "lightDown"),
            Self::LightUp => write!(f, "lightUp"),
            Self::LightGrid => write!(f, "lightGrid"),
            Self::LightTrellis => write!(f, "lightTrellis"),
            Self::Gray125 => write!(f, "gray125"),
            Self::Gray0625 => write!(f, "gray0625"),
        }
    }
}

impl std::str::FromStr for PatternType {
    type Err = String;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        match s {
            "none" => Ok(Self::None),
            "solid" => Ok(Self::Solid),
            "mediumGray" => Ok(Self::MediumGray),
            "darkGray" => Ok(Self::DarkGray),
            "lightGray" => Ok(Self::LightGray),
            "darkHorizontal" => Ok(Self::DarkHorizontal),
            "darkVertical" => Ok(Self::DarkVertical),
            "darkDown" => Ok(Self::DarkDown),
            "darkUp" => Ok(Self::DarkUp),
            "darkGrid" => Ok(Self::DarkGrid),
            "darkTrellis" => Ok(Self::DarkTrellis),
            "lightHorizontal" => Ok(Self::LightHorizontal),
            "lightVertical" => Ok(Self::LightVertical),
            "lightDown" => Ok(Self::LightDown),
            "lightUp" => Ok(Self::LightUp),
            "lightGrid" => Ok(Self::LightGrid),
            "lightTrellis" => Ok(Self::LightTrellis),
            "gray125" => Ok(Self::Gray125),
            "gray0625" => Ok(Self::Gray0625),
            _ => Err(format!("unknown PatternType value: {}", s)),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum GradientType {
    #[serde(rename = "linear")]
    Linear,
    #[serde(rename = "path")]
    Path,
}

impl std::fmt::Display for GradientType {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::Linear => write!(f, "linear"),
            Self::Path => write!(f, "path"),
        }
    }
}

impl std::str::FromStr for GradientType {
    type Err = String;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        match s {
            "linear" => Ok(Self::Linear),
            "path" => Ok(Self::Path),
            _ => Err(format!("unknown GradientType value: {}", s)),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum HorizontalAlignment {
    #[serde(rename = "general")]
    General,
    #[serde(rename = "left")]
    Left,
    #[serde(rename = "center")]
    Center,
    #[serde(rename = "right")]
    Right,
    #[serde(rename = "fill")]
    Fill,
    #[serde(rename = "justify")]
    Justify,
    #[serde(rename = "centerContinuous")]
    CenterContinuous,
    #[serde(rename = "distributed")]
    Distributed,
}

impl std::fmt::Display for HorizontalAlignment {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::General => write!(f, "general"),
            Self::Left => write!(f, "left"),
            Self::Center => write!(f, "center"),
            Self::Right => write!(f, "right"),
            Self::Fill => write!(f, "fill"),
            Self::Justify => write!(f, "justify"),
            Self::CenterContinuous => write!(f, "centerContinuous"),
            Self::Distributed => write!(f, "distributed"),
        }
    }
}

impl std::str::FromStr for HorizontalAlignment {
    type Err = String;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        match s {
            "general" => Ok(Self::General),
            "left" => Ok(Self::Left),
            "center" => Ok(Self::Center),
            "right" => Ok(Self::Right),
            "fill" => Ok(Self::Fill),
            "justify" => Ok(Self::Justify),
            "centerContinuous" => Ok(Self::CenterContinuous),
            "distributed" => Ok(Self::Distributed),
            _ => Err(format!("unknown HorizontalAlignment value: {}", s)),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum VerticalAlignment {
    #[serde(rename = "top")]
    Top,
    #[serde(rename = "center")]
    Center,
    #[serde(rename = "bottom")]
    Bottom,
    #[serde(rename = "justify")]
    Justify,
    #[serde(rename = "distributed")]
    Distributed,
}

impl std::fmt::Display for VerticalAlignment {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::Top => write!(f, "top"),
            Self::Center => write!(f, "center"),
            Self::Bottom => write!(f, "bottom"),
            Self::Justify => write!(f, "justify"),
            Self::Distributed => write!(f, "distributed"),
        }
    }
}

impl std::str::FromStr for VerticalAlignment {
    type Err = String;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        match s {
            "top" => Ok(Self::Top),
            "center" => Ok(Self::Center),
            "bottom" => Ok(Self::Bottom),
            "justify" => Ok(Self::Justify),
            "distributed" => Ok(Self::Distributed),
            _ => Err(format!("unknown VerticalAlignment value: {}", s)),
        }
    }
}

pub type STNumFmtId = u32;

pub type STFontId = u32;

pub type STFillId = u32;

pub type STBorderId = u32;

pub type STCellStyleXfId = u32;

pub type STDxfId = u32;

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum STTableStyleType {
    #[serde(rename = "wholeTable")]
    WholeTable,
    #[serde(rename = "headerRow")]
    HeaderRow,
    #[serde(rename = "totalRow")]
    TotalRow,
    #[serde(rename = "firstColumn")]
    FirstColumn,
    #[serde(rename = "lastColumn")]
    LastColumn,
    #[serde(rename = "firstRowStripe")]
    FirstRowStripe,
    #[serde(rename = "secondRowStripe")]
    SecondRowStripe,
    #[serde(rename = "firstColumnStripe")]
    FirstColumnStripe,
    #[serde(rename = "secondColumnStripe")]
    SecondColumnStripe,
    #[serde(rename = "firstHeaderCell")]
    FirstHeaderCell,
    #[serde(rename = "lastHeaderCell")]
    LastHeaderCell,
    #[serde(rename = "firstTotalCell")]
    FirstTotalCell,
    #[serde(rename = "lastTotalCell")]
    LastTotalCell,
    #[serde(rename = "firstSubtotalColumn")]
    FirstSubtotalColumn,
    #[serde(rename = "secondSubtotalColumn")]
    SecondSubtotalColumn,
    #[serde(rename = "thirdSubtotalColumn")]
    ThirdSubtotalColumn,
    #[serde(rename = "firstSubtotalRow")]
    FirstSubtotalRow,
    #[serde(rename = "secondSubtotalRow")]
    SecondSubtotalRow,
    #[serde(rename = "thirdSubtotalRow")]
    ThirdSubtotalRow,
    #[serde(rename = "blankRow")]
    BlankRow,
    #[serde(rename = "firstColumnSubheading")]
    FirstColumnSubheading,
    #[serde(rename = "secondColumnSubheading")]
    SecondColumnSubheading,
    #[serde(rename = "thirdColumnSubheading")]
    ThirdColumnSubheading,
    #[serde(rename = "firstRowSubheading")]
    FirstRowSubheading,
    #[serde(rename = "secondRowSubheading")]
    SecondRowSubheading,
    #[serde(rename = "thirdRowSubheading")]
    ThirdRowSubheading,
    #[serde(rename = "pageFieldLabels")]
    PageFieldLabels,
    #[serde(rename = "pageFieldValues")]
    PageFieldValues,
}

impl std::fmt::Display for STTableStyleType {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::WholeTable => write!(f, "wholeTable"),
            Self::HeaderRow => write!(f, "headerRow"),
            Self::TotalRow => write!(f, "totalRow"),
            Self::FirstColumn => write!(f, "firstColumn"),
            Self::LastColumn => write!(f, "lastColumn"),
            Self::FirstRowStripe => write!(f, "firstRowStripe"),
            Self::SecondRowStripe => write!(f, "secondRowStripe"),
            Self::FirstColumnStripe => write!(f, "firstColumnStripe"),
            Self::SecondColumnStripe => write!(f, "secondColumnStripe"),
            Self::FirstHeaderCell => write!(f, "firstHeaderCell"),
            Self::LastHeaderCell => write!(f, "lastHeaderCell"),
            Self::FirstTotalCell => write!(f, "firstTotalCell"),
            Self::LastTotalCell => write!(f, "lastTotalCell"),
            Self::FirstSubtotalColumn => write!(f, "firstSubtotalColumn"),
            Self::SecondSubtotalColumn => write!(f, "secondSubtotalColumn"),
            Self::ThirdSubtotalColumn => write!(f, "thirdSubtotalColumn"),
            Self::FirstSubtotalRow => write!(f, "firstSubtotalRow"),
            Self::SecondSubtotalRow => write!(f, "secondSubtotalRow"),
            Self::ThirdSubtotalRow => write!(f, "thirdSubtotalRow"),
            Self::BlankRow => write!(f, "blankRow"),
            Self::FirstColumnSubheading => write!(f, "firstColumnSubheading"),
            Self::SecondColumnSubheading => write!(f, "secondColumnSubheading"),
            Self::ThirdColumnSubheading => write!(f, "thirdColumnSubheading"),
            Self::FirstRowSubheading => write!(f, "firstRowSubheading"),
            Self::SecondRowSubheading => write!(f, "secondRowSubheading"),
            Self::ThirdRowSubheading => write!(f, "thirdRowSubheading"),
            Self::PageFieldLabels => write!(f, "pageFieldLabels"),
            Self::PageFieldValues => write!(f, "pageFieldValues"),
        }
    }
}

impl std::str::FromStr for STTableStyleType {
    type Err = String;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        match s {
            "wholeTable" => Ok(Self::WholeTable),
            "headerRow" => Ok(Self::HeaderRow),
            "totalRow" => Ok(Self::TotalRow),
            "firstColumn" => Ok(Self::FirstColumn),
            "lastColumn" => Ok(Self::LastColumn),
            "firstRowStripe" => Ok(Self::FirstRowStripe),
            "secondRowStripe" => Ok(Self::SecondRowStripe),
            "firstColumnStripe" => Ok(Self::FirstColumnStripe),
            "secondColumnStripe" => Ok(Self::SecondColumnStripe),
            "firstHeaderCell" => Ok(Self::FirstHeaderCell),
            "lastHeaderCell" => Ok(Self::LastHeaderCell),
            "firstTotalCell" => Ok(Self::FirstTotalCell),
            "lastTotalCell" => Ok(Self::LastTotalCell),
            "firstSubtotalColumn" => Ok(Self::FirstSubtotalColumn),
            "secondSubtotalColumn" => Ok(Self::SecondSubtotalColumn),
            "thirdSubtotalColumn" => Ok(Self::ThirdSubtotalColumn),
            "firstSubtotalRow" => Ok(Self::FirstSubtotalRow),
            "secondSubtotalRow" => Ok(Self::SecondSubtotalRow),
            "thirdSubtotalRow" => Ok(Self::ThirdSubtotalRow),
            "blankRow" => Ok(Self::BlankRow),
            "firstColumnSubheading" => Ok(Self::FirstColumnSubheading),
            "secondColumnSubheading" => Ok(Self::SecondColumnSubheading),
            "thirdColumnSubheading" => Ok(Self::ThirdColumnSubheading),
            "firstRowSubheading" => Ok(Self::FirstRowSubheading),
            "secondRowSubheading" => Ok(Self::SecondRowSubheading),
            "thirdRowSubheading" => Ok(Self::ThirdRowSubheading),
            "pageFieldLabels" => Ok(Self::PageFieldLabels),
            "pageFieldValues" => Ok(Self::PageFieldValues),
            _ => Err(format!("unknown STTableStyleType value: {}", s)),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum FontScheme {
    #[serde(rename = "none")]
    None,
    #[serde(rename = "major")]
    Major,
    #[serde(rename = "minor")]
    Minor,
}

impl std::fmt::Display for FontScheme {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::None => write!(f, "none"),
            Self::Major => write!(f, "major"),
            Self::Minor => write!(f, "minor"),
        }
    }
}

impl std::str::FromStr for FontScheme {
    type Err = String;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        match s {
            "none" => Ok(Self::None),
            "major" => Ok(Self::Major),
            "minor" => Ok(Self::Minor),
            _ => Err(format!("unknown FontScheme value: {}", s)),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum UnderlineStyle {
    #[serde(rename = "single")]
    Single,
    #[serde(rename = "double")]
    Double,
    #[serde(rename = "singleAccounting")]
    SingleAccounting,
    #[serde(rename = "doubleAccounting")]
    DoubleAccounting,
    #[serde(rename = "none")]
    None,
}

impl std::fmt::Display for UnderlineStyle {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::Single => write!(f, "single"),
            Self::Double => write!(f, "double"),
            Self::SingleAccounting => write!(f, "singleAccounting"),
            Self::DoubleAccounting => write!(f, "doubleAccounting"),
            Self::None => write!(f, "none"),
        }
    }
}

impl std::str::FromStr for UnderlineStyle {
    type Err = String;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        match s {
            "single" => Ok(Self::Single),
            "double" => Ok(Self::Double),
            "singleAccounting" => Ok(Self::SingleAccounting),
            "doubleAccounting" => Ok(Self::DoubleAccounting),
            "none" => Ok(Self::None),
            _ => Err(format!("unknown UnderlineStyle value: {}", s)),
        }
    }
}

pub type STFontFamily = i64;

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum STDdeValueType {
    #[serde(rename = "nil")]
    Nil,
    #[serde(rename = "b")]
    Boolean,
    #[serde(rename = "n")]
    Number,
    #[serde(rename = "e")]
    Error,
    #[serde(rename = "str")]
    String,
}

impl std::fmt::Display for STDdeValueType {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::Nil => write!(f, "nil"),
            Self::Boolean => write!(f, "b"),
            Self::Number => write!(f, "n"),
            Self::Error => write!(f, "e"),
            Self::String => write!(f, "str"),
        }
    }
}

impl std::str::FromStr for STDdeValueType {
    type Err = String;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        match s {
            "nil" => Ok(Self::Nil),
            "b" => Ok(Self::Boolean),
            "n" => Ok(Self::Number),
            "e" => Ok(Self::Error),
            "str" => Ok(Self::String),
            _ => Err(format!("unknown STDdeValueType value: {}", s)),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum STTableType {
    #[serde(rename = "worksheet")]
    Worksheet,
    #[serde(rename = "xml")]
    Xml,
    #[serde(rename = "queryTable")]
    QueryTable,
}

impl std::fmt::Display for STTableType {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::Worksheet => write!(f, "worksheet"),
            Self::Xml => write!(f, "xml"),
            Self::QueryTable => write!(f, "queryTable"),
        }
    }
}

impl std::str::FromStr for STTableType {
    type Err = String;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        match s {
            "worksheet" => Ok(Self::Worksheet),
            "xml" => Ok(Self::Xml),
            "queryTable" => Ok(Self::QueryTable),
            _ => Err(format!("unknown STTableType value: {}", s)),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum STTotalsRowFunction {
    #[serde(rename = "none")]
    None,
    #[serde(rename = "sum")]
    Sum,
    #[serde(rename = "min")]
    Min,
    #[serde(rename = "max")]
    Max,
    #[serde(rename = "average")]
    Average,
    #[serde(rename = "count")]
    Count,
    #[serde(rename = "countNums")]
    CountNums,
    #[serde(rename = "stdDev")]
    StdDev,
    #[serde(rename = "var")]
    Var,
    #[serde(rename = "custom")]
    Custom,
}

impl std::fmt::Display for STTotalsRowFunction {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::None => write!(f, "none"),
            Self::Sum => write!(f, "sum"),
            Self::Min => write!(f, "min"),
            Self::Max => write!(f, "max"),
            Self::Average => write!(f, "average"),
            Self::Count => write!(f, "count"),
            Self::CountNums => write!(f, "countNums"),
            Self::StdDev => write!(f, "stdDev"),
            Self::Var => write!(f, "var"),
            Self::Custom => write!(f, "custom"),
        }
    }
}

impl std::str::FromStr for STTotalsRowFunction {
    type Err = String;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        match s {
            "none" => Ok(Self::None),
            "sum" => Ok(Self::Sum),
            "min" => Ok(Self::Min),
            "max" => Ok(Self::Max),
            "average" => Ok(Self::Average),
            "count" => Ok(Self::Count),
            "countNums" => Ok(Self::CountNums),
            "stdDev" => Ok(Self::StdDev),
            "var" => Ok(Self::Var),
            "custom" => Ok(Self::Custom),
            _ => Err(format!("unknown STTotalsRowFunction value: {}", s)),
        }
    }
}

pub type STXmlDataType = String;

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum STVolDepType {
    #[serde(rename = "realTimeData")]
    RealTimeData,
    #[serde(rename = "olapFunctions")]
    OlapFunctions,
}

impl std::fmt::Display for STVolDepType {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::RealTimeData => write!(f, "realTimeData"),
            Self::OlapFunctions => write!(f, "olapFunctions"),
        }
    }
}

impl std::str::FromStr for STVolDepType {
    type Err = String;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        match s {
            "realTimeData" => Ok(Self::RealTimeData),
            "olapFunctions" => Ok(Self::OlapFunctions),
            _ => Err(format!("unknown STVolDepType value: {}", s)),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum STVolValueType {
    #[serde(rename = "b")]
    Boolean,
    #[serde(rename = "n")]
    Number,
    #[serde(rename = "e")]
    Error,
    #[serde(rename = "s")]
    SharedString,
}

impl std::fmt::Display for STVolValueType {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::Boolean => write!(f, "b"),
            Self::Number => write!(f, "n"),
            Self::Error => write!(f, "e"),
            Self::SharedString => write!(f, "s"),
        }
    }
}

impl std::str::FromStr for STVolValueType {
    type Err = String;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        match s {
            "b" => Ok(Self::Boolean),
            "n" => Ok(Self::Number),
            "e" => Ok(Self::Error),
            "s" => Ok(Self::SharedString),
            _ => Err(format!("unknown STVolValueType value: {}", s)),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum Visibility {
    #[serde(rename = "visible")]
    Visible,
    #[serde(rename = "hidden")]
    Hidden,
    #[serde(rename = "veryHidden")]
    VeryHidden,
}

impl std::fmt::Display for Visibility {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::Visible => write!(f, "visible"),
            Self::Hidden => write!(f, "hidden"),
            Self::VeryHidden => write!(f, "veryHidden"),
        }
    }
}

impl std::str::FromStr for Visibility {
    type Err = String;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        match s {
            "visible" => Ok(Self::Visible),
            "hidden" => Ok(Self::Hidden),
            "veryHidden" => Ok(Self::VeryHidden),
            _ => Err(format!("unknown Visibility value: {}", s)),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum CommentVisibility {
    #[serde(rename = "commNone")]
    CommNone,
    #[serde(rename = "commIndicator")]
    CommIndicator,
    #[serde(rename = "commIndAndComment")]
    CommIndAndComment,
}

impl std::fmt::Display for CommentVisibility {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::CommNone => write!(f, "commNone"),
            Self::CommIndicator => write!(f, "commIndicator"),
            Self::CommIndAndComment => write!(f, "commIndAndComment"),
        }
    }
}

impl std::str::FromStr for CommentVisibility {
    type Err = String;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        match s {
            "commNone" => Ok(Self::CommNone),
            "commIndicator" => Ok(Self::CommIndicator),
            "commIndAndComment" => Ok(Self::CommIndAndComment),
            _ => Err(format!("unknown CommentVisibility value: {}", s)),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum ObjectVisibility {
    #[serde(rename = "all")]
    All,
    #[serde(rename = "placeholders")]
    Placeholders,
    #[serde(rename = "none")]
    None,
}

impl std::fmt::Display for ObjectVisibility {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::All => write!(f, "all"),
            Self::Placeholders => write!(f, "placeholders"),
            Self::None => write!(f, "none"),
        }
    }
}

impl std::str::FromStr for ObjectVisibility {
    type Err = String;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        match s {
            "all" => Ok(Self::All),
            "placeholders" => Ok(Self::Placeholders),
            "none" => Ok(Self::None),
            _ => Err(format!("unknown ObjectVisibility value: {}", s)),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum SheetState {
    #[serde(rename = "visible")]
    Visible,
    #[serde(rename = "hidden")]
    Hidden,
    #[serde(rename = "veryHidden")]
    VeryHidden,
}

impl std::fmt::Display for SheetState {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::Visible => write!(f, "visible"),
            Self::Hidden => write!(f, "hidden"),
            Self::VeryHidden => write!(f, "veryHidden"),
        }
    }
}

impl std::str::FromStr for SheetState {
    type Err = String;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        match s {
            "visible" => Ok(Self::Visible),
            "hidden" => Ok(Self::Hidden),
            "veryHidden" => Ok(Self::VeryHidden),
            _ => Err(format!("unknown SheetState value: {}", s)),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum UpdateLinks {
    #[serde(rename = "userSet")]
    UserSet,
    #[serde(rename = "never")]
    Never,
    #[serde(rename = "always")]
    Always,
}

impl std::fmt::Display for UpdateLinks {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::UserSet => write!(f, "userSet"),
            Self::Never => write!(f, "never"),
            Self::Always => write!(f, "always"),
        }
    }
}

impl std::str::FromStr for UpdateLinks {
    type Err = String;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        match s {
            "userSet" => Ok(Self::UserSet),
            "never" => Ok(Self::Never),
            "always" => Ok(Self::Always),
            _ => Err(format!("unknown UpdateLinks value: {}", s)),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum STSmartTagShow {
    #[serde(rename = "all")]
    All,
    #[serde(rename = "none")]
    None,
    #[serde(rename = "noIndicator")]
    NoIndicator,
}

impl std::fmt::Display for STSmartTagShow {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::All => write!(f, "all"),
            Self::None => write!(f, "none"),
            Self::NoIndicator => write!(f, "noIndicator"),
        }
    }
}

impl std::str::FromStr for STSmartTagShow {
    type Err = String;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        match s {
            "all" => Ok(Self::All),
            "none" => Ok(Self::None),
            "noIndicator" => Ok(Self::NoIndicator),
            _ => Err(format!("unknown STSmartTagShow value: {}", s)),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum CalculationMode {
    #[serde(rename = "manual")]
    Manual,
    #[serde(rename = "auto")]
    Auto,
    #[serde(rename = "autoNoTable")]
    AutoNoTable,
}

impl std::fmt::Display for CalculationMode {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::Manual => write!(f, "manual"),
            Self::Auto => write!(f, "auto"),
            Self::AutoNoTable => write!(f, "autoNoTable"),
        }
    }
}

impl std::str::FromStr for CalculationMode {
    type Err = String;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        match s {
            "manual" => Ok(Self::Manual),
            "auto" => Ok(Self::Auto),
            "autoNoTable" => Ok(Self::AutoNoTable),
            _ => Err(format!("unknown CalculationMode value: {}", s)),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum ReferenceMode {
    #[serde(rename = "A1")]
    A1,
    #[serde(rename = "R1C1")]
    R1C1,
}

impl std::fmt::Display for ReferenceMode {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::A1 => write!(f, "A1"),
            Self::R1C1 => write!(f, "R1C1"),
        }
    }
}

impl std::str::FromStr for ReferenceMode {
    type Err = String;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        match s {
            "A1" => Ok(Self::A1),
            "R1C1" => Ok(Self::R1C1),
            _ => Err(format!("unknown ReferenceMode value: {}", s)),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum STTargetScreenSize {
    #[serde(rename = "544x376")]
    _544x376,
    #[serde(rename = "640x480")]
    _640x480,
    #[serde(rename = "720x512")]
    _720x512,
    #[serde(rename = "800x600")]
    _800x600,
    #[serde(rename = "1024x768")]
    _1024x768,
    #[serde(rename = "1152x882")]
    _1152x882,
    #[serde(rename = "1152x900")]
    _1152x900,
    #[serde(rename = "1280x1024")]
    _1280x1024,
    #[serde(rename = "1600x1200")]
    _1600x1200,
    #[serde(rename = "1800x1440")]
    _1800x1440,
    #[serde(rename = "1920x1200")]
    _1920x1200,
}

impl std::fmt::Display for STTargetScreenSize {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::_544x376 => write!(f, "544x376"),
            Self::_640x480 => write!(f, "640x480"),
            Self::_720x512 => write!(f, "720x512"),
            Self::_800x600 => write!(f, "800x600"),
            Self::_1024x768 => write!(f, "1024x768"),
            Self::_1152x882 => write!(f, "1152x882"),
            Self::_1152x900 => write!(f, "1152x900"),
            Self::_1280x1024 => write!(f, "1280x1024"),
            Self::_1600x1200 => write!(f, "1600x1200"),
            Self::_1800x1440 => write!(f, "1800x1440"),
            Self::_1920x1200 => write!(f, "1920x1200"),
        }
    }
}

impl std::str::FromStr for STTargetScreenSize {
    type Err = String;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        match s {
            "544x376" => Ok(Self::_544x376),
            "640x480" => Ok(Self::_640x480),
            "720x512" => Ok(Self::_720x512),
            "800x600" => Ok(Self::_800x600),
            "1024x768" => Ok(Self::_1024x768),
            "1152x882" => Ok(Self::_1152x882),
            "1152x900" => Ok(Self::_1152x900),
            "1280x1024" => Ok(Self::_1280x1024),
            "1600x1200" => Ok(Self::_1600x1200),
            "1800x1440" => Ok(Self::_1800x1440),
            "1920x1200" => Ok(Self::_1920x1200),
            _ => Err(format!("unknown STTargetScreenSize value: {}", s)),
        }
    }
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
#[serde(rename = "autoFilter")]
pub struct AutoFilter {
    #[cfg(feature = "sml-filtering")]
    #[serde(rename = "@ref")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub reference: Option<Reference>,
    #[cfg(feature = "sml-filtering")]
    #[serde(rename = "filterColumn")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub filter_column: Vec<FilterColumn>,
    #[cfg(feature = "sml-filtering")]
    #[serde(rename = "sortState")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub sort_state: Option<Box<SortState>>,
    #[cfg(feature = "sml-extensions")]
    #[serde(rename = "extLst")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub extension_list: Option<Box<ExtensionList>>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(rename = "filterColumn")]
pub struct FilterColumn {
    #[cfg(feature = "sml-filtering")]
    #[serde(rename = "@colId")]
    pub column_id: u32,
    #[cfg(feature = "sml-filtering")]
    #[serde(rename = "@hiddenButton")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub hidden_button: Option<bool>,
    #[cfg(feature = "sml-filtering")]
    #[serde(rename = "@showButton")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub show_button: Option<bool>,
    #[cfg(feature = "sml-filtering")]
    #[serde(rename = "filters")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub filters: Option<Box<Filters>>,
    #[cfg(feature = "sml-filtering")]
    #[serde(rename = "top10")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub top10: Option<Box<Top10Filter>>,
    #[cfg(feature = "sml-filtering")]
    #[serde(rename = "customFilters")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub custom_filters: Option<Box<CustomFilters>>,
    #[cfg(feature = "sml-filtering")]
    #[serde(rename = "dynamicFilter")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub dynamic_filter: Option<Box<DynamicFilter>>,
    #[cfg(feature = "sml-filtering")]
    #[serde(rename = "colorFilter")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub color_filter: Option<Box<ColorFilter>>,
    #[cfg(feature = "sml-filtering")]
    #[serde(rename = "iconFilter")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub icon_filter: Option<Box<IconFilter>>,
    #[cfg(feature = "sml-extensions")]
    #[serde(rename = "extLst")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub extension_list: Option<Box<ExtensionList>>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
#[serde(rename = "filters")]
pub struct Filters {
    #[serde(rename = "@blank")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub blank: Option<bool>,
    #[serde(rename = "@calendarType")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub calendar_type: Option<CalendarType>,
    #[serde(rename = "filter")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub filter: Vec<Filter>,
    #[serde(rename = "dateGroupItem")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub date_group_item: Vec<DateGroupItem>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
#[serde(rename = "filter")]
pub struct Filter {
    #[serde(rename = "@val")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub value: Option<XmlString>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct CustomFilters {
    #[serde(rename = "@and")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub and: Option<bool>,
    #[serde(rename = "customFilter")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub custom_filter: Vec<CustomFilter>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct CustomFilter {
    #[serde(rename = "@operator")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub operator: Option<FilterOperator>,
    #[serde(rename = "@val")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub value: Option<XmlString>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct Top10Filter {
    #[serde(rename = "@top")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub top: Option<bool>,
    #[serde(rename = "@percent")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub percent: Option<bool>,
    #[serde(rename = "@val")]
    pub value: f64,
    #[serde(rename = "@filterVal")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub filter_val: Option<f64>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct ColorFilter {
    #[serde(rename = "@dxfId")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub dxf_id: Option<STDxfId>,
    #[serde(rename = "@cellColor")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub cell_color: Option<bool>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct IconFilter {
    #[serde(rename = "@iconSet")]
    pub icon_set: IconSetType,
    #[serde(rename = "@iconId")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub icon_id: Option<u32>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct DynamicFilter {
    #[serde(rename = "@type")]
    pub r#type: DynamicFilterType,
    #[serde(rename = "@val")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub value: Option<f64>,
    #[serde(rename = "@valIso")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub val_iso: Option<String>,
    #[serde(rename = "@maxVal")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub max_val: Option<f64>,
    #[serde(rename = "@maxValIso")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub max_val_iso: Option<String>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(rename = "sortState")]
pub struct SortState {
    #[serde(rename = "@columnSort")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub column_sort: Option<bool>,
    #[serde(rename = "@caseSensitive")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub case_sensitive: Option<bool>,
    #[serde(rename = "@sortMethod")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub sort_method: Option<SortMethod>,
    #[serde(rename = "@ref")]
    pub reference: Reference,
    #[serde(rename = "sortCondition")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub sort_condition: Vec<SortCondition>,
    #[serde(rename = "extLst")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub extension_list: Option<Box<ExtensionList>>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(rename = "sortCondition")]
pub struct SortCondition {
    #[serde(rename = "@descending")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub descending: Option<bool>,
    #[serde(rename = "@sortBy")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub sort_by: Option<SortBy>,
    #[serde(rename = "@ref")]
    pub reference: Reference,
    #[serde(rename = "@customList")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub custom_list: Option<XmlString>,
    #[serde(rename = "@dxfId")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub dxf_id: Option<STDxfId>,
    #[serde(rename = "@iconSet")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub icon_set: Option<IconSetType>,
    #[serde(rename = "@iconId")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub icon_id: Option<u32>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct DateGroupItem {
    #[serde(rename = "@year")]
    pub year: u16,
    #[serde(rename = "@month")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub month: Option<u16>,
    #[serde(rename = "@day")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub day: Option<u16>,
    #[serde(rename = "@hour")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub hour: Option<u16>,
    #[serde(rename = "@minute")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub minute: Option<u16>,
    #[serde(rename = "@second")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub second: Option<u16>,
    #[serde(rename = "@dateTimeGrouping")]
    pub date_time_grouping: STDateTimeGrouping,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTXStringElement {
    #[serde(rename = "@v")]
    pub value: XmlString,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
#[serde(rename = "ext")]
pub struct Extension {
    #[serde(rename = "@uri")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub uri: Option<String>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

pub type CTExtensionAny = String;

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct ObjectAnchor {
    #[serde(rename = "@moveWithCells")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub move_with_cells: Option<bool>,
    #[serde(rename = "@sizeWithCells")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub size_with_cells: Option<bool>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct EGExtensionList {
    #[serde(rename = "ext")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub ext: Vec<Extension>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
#[serde(rename = "extLst")]
pub struct ExtensionList {
    #[serde(rename = "ext")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub ext: Vec<Extension>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

pub type SmlCalcChain = Box<CalcChain>;

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
#[serde(rename = "calcChain")]
pub struct CalcChain {
    #[cfg(feature = "sml-formulas")]
    #[serde(rename = "c")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub cells: Vec<CalcCell>,
    #[cfg(feature = "sml-extensions")]
    #[serde(rename = "extLst")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub extension_list: Option<Box<ExtensionList>>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CalcCell {
    #[serde(rename = "@_any")]
    pub _any: CellRef,
    #[cfg(feature = "sml-formulas")]
    #[serde(rename = "@i")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub i: Option<i32>,
    #[cfg(feature = "sml-formulas")]
    #[serde(rename = "@s")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub style_index: Option<bool>,
    #[cfg(feature = "sml-formulas")]
    #[serde(rename = "@l")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub l: Option<bool>,
    #[cfg(feature = "sml-formulas")]
    #[serde(rename = "@t")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub cell_type: Option<bool>,
    #[cfg(feature = "sml-formulas")]
    #[serde(rename = "@a")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub a: Option<bool>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

pub type SmlComments = Box<Comments>;

#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(rename = "comments")]
pub struct Comments {
    #[serde(rename = "authors")]
    pub authors: Box<Authors>,
    #[serde(rename = "commentList")]
    pub comment_list: Box<CommentList>,
    #[serde(rename = "extLst")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub extension_list: Option<Box<ExtensionList>>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
#[serde(rename = "authors")]
pub struct Authors {
    #[serde(rename = "author")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub author: Vec<XmlString>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
#[serde(rename = "commentList")]
pub struct CommentList {
    #[serde(rename = "comment")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub comment: Vec<Comment>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(rename = "comment")]
pub struct Comment {
    #[cfg(feature = "sml-comments")]
    #[serde(rename = "@ref")]
    pub reference: Reference,
    #[cfg(feature = "sml-comments")]
    #[serde(rename = "@authorId")]
    pub author_id: u32,
    #[cfg(feature = "sml-comments")]
    #[serde(rename = "@guid")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub guid: Option<Guid>,
    #[cfg(feature = "sml-comments")]
    #[serde(rename = "@shapeId")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub shape_id: Option<u32>,
    #[cfg(feature = "sml-comments")]
    #[serde(rename = "text")]
    pub text: Box<RichString>,
    #[serde(rename = "commentPr")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub comment_pr: Option<Box<CTCommentPr>>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTCommentPr {
    #[serde(rename = "@locked")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub locked: Option<bool>,
    #[serde(rename = "@defaultSize")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub default_size: Option<bool>,
    #[serde(rename = "@print")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub print: Option<bool>,
    #[serde(rename = "@disabled")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub disabled: Option<bool>,
    #[serde(rename = "@autoFill")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub auto_fill: Option<bool>,
    #[serde(rename = "@autoLine")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub auto_line: Option<bool>,
    #[serde(rename = "@altText")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub alt_text: Option<XmlString>,
    #[serde(rename = "@textHAlign")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub text_h_align: Option<STTextHAlign>,
    #[serde(rename = "@textVAlign")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub text_v_align: Option<STTextVAlign>,
    #[serde(rename = "@lockText")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub lock_text: Option<bool>,
    #[serde(rename = "@justLastX")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub just_last_x: Option<bool>,
    #[serde(rename = "@autoScale")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub auto_scale: Option<bool>,
    #[serde(rename = "anchor")]
    pub anchor: Box<ObjectAnchor>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

pub type SmlMapInfo = Box<MapInfo>;

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct MapInfo {
    #[serde(rename = "@SelectionNamespaces")]
    pub selection_namespaces: String,
    #[serde(rename = "Schema")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub schema: Vec<XmlSchema>,
    #[serde(rename = "Map")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub map: Vec<XmlMap>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct XmlSchema {
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

pub type CTSchemaAny = String;

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct XmlMap {
    #[serde(rename = "@ID")]
    pub i_d: u32,
    #[serde(rename = "@Name")]
    pub name: String,
    #[serde(rename = "@RootElement")]
    pub root_element: String,
    #[serde(rename = "@SchemaID")]
    pub schema_i_d: String,
    #[serde(rename = "@ShowImportExportValidationErrors")]
    #[serde(with = "ooxml_xml::ooxml_bool_required")]
    pub show_import_export_validation_errors: bool,
    #[serde(rename = "@AutoFit")]
    #[serde(with = "ooxml_xml::ooxml_bool_required")]
    pub auto_fit: bool,
    #[serde(rename = "@Append")]
    #[serde(with = "ooxml_xml::ooxml_bool_required")]
    pub append: bool,
    #[serde(rename = "@PreserveSortAFLayout")]
    #[serde(with = "ooxml_xml::ooxml_bool_required")]
    pub preserve_sort_a_f_layout: bool,
    #[serde(rename = "@PreserveFormat")]
    #[serde(with = "ooxml_xml::ooxml_bool_required")]
    pub preserve_format: bool,
    #[serde(rename = "DataBinding")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub data_binding: Option<Box<DataBinding>>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct DataBinding {
    #[serde(rename = "@DataBindingName")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub data_binding_name: Option<String>,
    #[serde(rename = "@FileBinding")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub file_binding: Option<bool>,
    #[serde(rename = "@ConnectionID")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub connection_i_d: Option<u32>,
    #[serde(rename = "@FileBindingName")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub file_binding_name: Option<String>,
    #[serde(rename = "@DataBindingLoadMode")]
    pub data_binding_load_mode: u32,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

pub type CTDataBindingAny = String;

pub type SmlConnections = Box<Connections>;

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct Connections {
    #[serde(rename = "connection")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub connection: Vec<Connection>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct Connection {
    #[serde(rename = "@id")]
    pub id: u32,
    #[serde(rename = "@sourceFile")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub source_file: Option<XmlString>,
    #[serde(rename = "@odcFile")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub odc_file: Option<XmlString>,
    #[serde(rename = "@keepAlive")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub keep_alive: Option<bool>,
    #[serde(rename = "@interval")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub interval: Option<u32>,
    #[serde(rename = "@name")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub name: Option<XmlString>,
    #[serde(rename = "@description")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub description: Option<XmlString>,
    #[serde(rename = "@type")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub r#type: Option<u32>,
    #[serde(rename = "@reconnectionMethod")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub reconnection_method: Option<u32>,
    #[serde(rename = "@refreshedVersion")]
    pub refreshed_version: u8,
    #[serde(rename = "@minRefreshableVersion")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub min_refreshable_version: Option<u8>,
    #[serde(rename = "@savePassword")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub save_password: Option<bool>,
    #[serde(rename = "@new")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub new: Option<bool>,
    #[serde(rename = "@deleted")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub deleted: Option<bool>,
    #[serde(rename = "@onlyUseConnectionFile")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub only_use_connection_file: Option<bool>,
    #[serde(rename = "@background")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub background: Option<bool>,
    #[serde(rename = "@refreshOnLoad")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub refresh_on_load: Option<bool>,
    #[serde(rename = "@saveData")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub save_data: Option<bool>,
    #[serde(rename = "@credentials")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub credentials: Option<STCredMethod>,
    #[serde(rename = "@singleSignOnId")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub single_sign_on_id: Option<XmlString>,
    #[serde(rename = "dbPr")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub db_pr: Option<Box<DatabaseProperties>>,
    #[serde(rename = "olapPr")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub olap_pr: Option<Box<OlapProperties>>,
    #[serde(rename = "webPr")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub web_pr: Option<Box<WebQueryProperties>>,
    #[serde(rename = "textPr")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub text_pr: Option<Box<TextImportProperties>>,
    #[serde(rename = "parameters")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub parameters: Option<Box<Parameters>>,
    #[serde(rename = "extLst")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub extension_list: Option<Box<ExtensionList>>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct DatabaseProperties {
    #[serde(rename = "@connection")]
    pub connection: XmlString,
    #[serde(rename = "@command")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub command: Option<XmlString>,
    #[serde(rename = "@serverCommand")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub server_command: Option<XmlString>,
    #[serde(rename = "@commandType")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub command_type: Option<u32>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct OlapProperties {
    #[serde(rename = "@local")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub local: Option<bool>,
    #[serde(rename = "@localConnection")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub local_connection: Option<XmlString>,
    #[serde(rename = "@localRefresh")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub local_refresh: Option<bool>,
    #[serde(rename = "@sendLocale")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub send_locale: Option<bool>,
    #[serde(rename = "@rowDrillCount")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub row_drill_count: Option<u32>,
    #[serde(rename = "@serverFill")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub server_fill: Option<bool>,
    #[serde(rename = "@serverNumberFormat")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub server_number_format: Option<bool>,
    #[serde(rename = "@serverFont")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub server_font: Option<bool>,
    #[serde(rename = "@serverFontColor")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub server_font_color: Option<bool>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct WebQueryProperties {
    #[serde(rename = "@xml")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub xml: Option<bool>,
    #[serde(rename = "@sourceData")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub source_data: Option<bool>,
    #[serde(rename = "@parsePre")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub parse_pre: Option<bool>,
    #[serde(rename = "@consecutive")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub consecutive: Option<bool>,
    #[serde(rename = "@firstRow")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub first_row: Option<bool>,
    #[serde(rename = "@xl97")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub xl97: Option<bool>,
    #[serde(rename = "@textDates")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub text_dates: Option<bool>,
    #[serde(rename = "@xl2000")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub xl2000: Option<bool>,
    #[serde(rename = "@url")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub url: Option<XmlString>,
    #[serde(rename = "@post")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub post: Option<XmlString>,
    #[serde(rename = "@htmlTables")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub html_tables: Option<bool>,
    #[serde(rename = "@htmlFormat")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub html_format: Option<STHtmlFmt>,
    #[serde(rename = "@editPage")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub edit_page: Option<XmlString>,
    #[serde(rename = "tables")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub tables: Option<Box<DataTables>>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct Parameters {
    #[serde(rename = "@count")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub count: Option<u32>,
    #[serde(rename = "parameter")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub parameter: Vec<Parameter>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct Parameter {
    #[serde(rename = "@name")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub name: Option<XmlString>,
    #[serde(rename = "@sqlType")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub sql_type: Option<i32>,
    #[serde(rename = "@parameterType")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub parameter_type: Option<STParameterType>,
    #[serde(rename = "@refreshOnChange")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub refresh_on_change: Option<bool>,
    #[serde(rename = "@prompt")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub prompt: Option<XmlString>,
    #[serde(rename = "@boolean")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub boolean: Option<bool>,
    #[serde(rename = "@double")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub double: Option<f64>,
    #[serde(rename = "@integer")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub integer: Option<i32>,
    #[serde(rename = "@string")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub string: Option<XmlString>,
    #[serde(rename = "@cell")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub cell: Option<XmlString>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct DataTables {
    #[serde(rename = "@count")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub count: Option<u32>,
    #[serde(rename = "m")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub m: Vec<TableMissing>,
    #[serde(rename = "s")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub style_index: Vec<CTXStringElement>,
    #[serde(rename = "x")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub x: Vec<CTIndex>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct TableMissing;

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct TextImportProperties {
    #[serde(rename = "@prompt")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub prompt: Option<bool>,
    #[serde(rename = "@fileType")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub file_type: Option<STFileType>,
    #[serde(rename = "@codePage")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub code_page: Option<u32>,
    #[serde(rename = "@characterSet")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub character_set: Option<String>,
    #[serde(rename = "@firstRow")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub first_row: Option<u32>,
    #[serde(rename = "@sourceFile")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub source_file: Option<XmlString>,
    #[serde(rename = "@delimited")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub delimited: Option<bool>,
    #[serde(rename = "@decimal")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub decimal: Option<XmlString>,
    #[serde(rename = "@thousands")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub thousands: Option<XmlString>,
    #[serde(rename = "@tab")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub tab: Option<bool>,
    #[serde(rename = "@space")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub space: Option<bool>,
    #[serde(rename = "@comma")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub comma: Option<bool>,
    #[serde(rename = "@semicolon")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub semicolon: Option<bool>,
    #[serde(rename = "@consecutive")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub consecutive: Option<bool>,
    #[serde(rename = "@qualifier")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub qualifier: Option<STQualifier>,
    #[serde(rename = "@delimiter")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub delimiter: Option<XmlString>,
    #[serde(rename = "textFields")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub text_fields: Option<Box<TextFields>>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct TextFields {
    #[serde(rename = "@count")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub count: Option<u32>,
    #[serde(rename = "textField")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub text_field: Vec<TextField>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct TextField {
    #[serde(rename = "@type")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub r#type: Option<STExternalConnectionType>,
    #[serde(rename = "@position")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub position: Option<u32>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

pub type SmlPivotCacheDefinition = Box<PivotCacheDefinition>;

pub type SmlPivotCacheRecords = Box<PivotCacheRecords>;

pub type SmlPivotTableDefinition = Box<CTPivotTableDefinition>;

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct PivotCacheDefinition {
    #[serde(rename = "@r:id")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub id: Option<STRelationshipId>,
    #[serde(rename = "@invalid")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub invalid: Option<bool>,
    #[serde(rename = "@saveData")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub save_data: Option<bool>,
    #[serde(rename = "@refreshOnLoad")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub refresh_on_load: Option<bool>,
    #[serde(rename = "@optimizeMemory")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub optimize_memory: Option<bool>,
    #[serde(rename = "@enableRefresh")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub enable_refresh: Option<bool>,
    #[serde(rename = "@refreshedBy")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub refreshed_by: Option<XmlString>,
    #[serde(rename = "@refreshedDate")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub refreshed_date: Option<f64>,
    #[serde(rename = "@refreshedDateIso")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub refreshed_date_iso: Option<String>,
    #[serde(rename = "@backgroundQuery")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub background_query: Option<bool>,
    #[serde(rename = "@missingItemsLimit")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub missing_items_limit: Option<u32>,
    #[serde(rename = "@createdVersion")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub created_version: Option<u8>,
    #[serde(rename = "@refreshedVersion")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub refreshed_version: Option<u8>,
    #[serde(rename = "@minRefreshableVersion")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub min_refreshable_version: Option<u8>,
    #[serde(rename = "@recordCount")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub record_count: Option<u32>,
    #[serde(rename = "@upgradeOnRefresh")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub upgrade_on_refresh: Option<bool>,
    #[serde(rename = "@tupleCache")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub tuple_cache: Option<bool>,
    #[serde(rename = "@supportSubquery")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub support_subquery: Option<bool>,
    #[serde(rename = "@supportAdvancedDrill")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub support_advanced_drill: Option<bool>,
    #[serde(rename = "cacheSource")]
    pub cache_source: Box<CacheSource>,
    #[serde(rename = "cacheFields")]
    pub cache_fields: Box<CacheFields>,
    #[serde(rename = "cacheHierarchies")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub cache_hierarchies: Option<Box<CTCacheHierarchies>>,
    #[serde(rename = "kpis")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub kpis: Option<Box<CTPCDKPIs>>,
    #[serde(rename = "calculatedItems")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub calculated_items: Option<Box<CTCalculatedItems>>,
    #[serde(rename = "calculatedMembers")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub calculated_members: Option<Box<CTCalculatedMembers>>,
    #[serde(rename = "dimensions")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub dimensions: Option<Box<CTDimensions>>,
    #[serde(rename = "measureGroups")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub measure_groups: Option<Box<CTMeasureGroups>>,
    #[serde(rename = "maps")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub maps: Option<Box<CTMeasureDimensionMaps>>,
    #[serde(rename = "extLst")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub extension_list: Option<Box<ExtensionList>>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct CacheFields {
    #[serde(rename = "@count")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub count: Option<u32>,
    #[serde(rename = "cacheField")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub cache_field: Vec<CacheField>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CacheField {
    #[serde(rename = "@name")]
    pub name: XmlString,
    #[serde(rename = "@caption")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub caption: Option<XmlString>,
    #[serde(rename = "@propertyName")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub property_name: Option<XmlString>,
    #[serde(rename = "@serverField")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub server_field: Option<bool>,
    #[serde(rename = "@uniqueList")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub unique_list: Option<bool>,
    #[serde(rename = "@numFmtId")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub number_format_id: Option<STNumFmtId>,
    #[serde(rename = "@formula")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub formula: Option<XmlString>,
    #[serde(rename = "@sqlType")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub sql_type: Option<i32>,
    #[serde(rename = "@hierarchy")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub hierarchy: Option<i32>,
    #[serde(rename = "@level")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub level: Option<u32>,
    #[serde(rename = "@databaseField")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub database_field: Option<bool>,
    #[serde(rename = "@mappingCount")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub mapping_count: Option<u32>,
    #[serde(rename = "@memberPropertyField")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub member_property_field: Option<bool>,
    #[serde(rename = "sharedItems")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub shared_items: Option<Box<SharedItems>>,
    #[serde(rename = "fieldGroup")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub field_group: Option<Box<FieldGroup>>,
    #[serde(rename = "mpMap")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub mp_map: Vec<CTX>,
    #[serde(rename = "extLst")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub extension_list: Option<Box<ExtensionList>>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CacheSource {
    #[serde(rename = "@type")]
    pub r#type: STSourceType,
    #[serde(rename = "@connectionId")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub connection_id: Option<u32>,
    #[serde(rename = "worksheetSource")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub worksheet_source: Option<Box<WorksheetSource>>,
    #[serde(rename = "consolidation")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub consolidation: Option<Box<Consolidation>>,
    #[serde(rename = "extLst")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub extension_list: Option<Box<ExtensionList>>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct WorksheetSource {
    #[serde(rename = "@ref")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub reference: Option<Reference>,
    #[serde(rename = "@name")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub name: Option<XmlString>,
    #[serde(rename = "@sheet")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub sheet: Option<XmlString>,
    #[serde(rename = "@r:id")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub id: Option<STRelationshipId>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct Consolidation {
    #[serde(rename = "@autoPage")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub auto_page: Option<bool>,
    #[serde(rename = "pages")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub pages: Option<Box<CTPages>>,
    #[serde(rename = "rangeSets")]
    pub range_sets: Box<CTRangeSets>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct CTPages {
    #[serde(rename = "@count")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub count: Option<u32>,
    #[serde(rename = "page")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub page: Vec<CTPCDSCPage>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct CTPCDSCPage {
    #[serde(rename = "@count")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub count: Option<u32>,
    #[serde(rename = "pageItem")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub page_item: Vec<CTPageItem>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTPageItem {
    #[serde(rename = "@name")]
    pub name: XmlString,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct CTRangeSets {
    #[serde(rename = "@count")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub count: Option<u32>,
    #[serde(rename = "rangeSet")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub range_set: Vec<CTRangeSet>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct CTRangeSet {
    #[serde(rename = "@i1")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub i1: Option<u32>,
    #[serde(rename = "@i2")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub i2: Option<u32>,
    #[serde(rename = "@i3")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub i3: Option<u32>,
    #[serde(rename = "@i4")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub i4: Option<u32>,
    #[serde(rename = "@ref")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub reference: Option<Reference>,
    #[serde(rename = "@name")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub name: Option<XmlString>,
    #[serde(rename = "@sheet")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub sheet: Option<XmlString>,
    #[serde(rename = "@r:id")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub id: Option<STRelationshipId>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct SharedItems {
    #[serde(rename = "@containsSemiMixedTypes")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub contains_semi_mixed_types: Option<bool>,
    #[serde(rename = "@containsNonDate")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub contains_non_date: Option<bool>,
    #[serde(rename = "@containsDate")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub contains_date: Option<bool>,
    #[serde(rename = "@containsString")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub contains_string: Option<bool>,
    #[serde(rename = "@containsBlank")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub contains_blank: Option<bool>,
    #[serde(rename = "@containsMixedTypes")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub contains_mixed_types: Option<bool>,
    #[serde(rename = "@containsNumber")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub contains_number: Option<bool>,
    #[serde(rename = "@containsInteger")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub contains_integer: Option<bool>,
    #[serde(rename = "@minValue")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub min_value: Option<f64>,
    #[serde(rename = "@maxValue")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub max_value: Option<f64>,
    #[serde(rename = "@minDate")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub min_date: Option<String>,
    #[serde(rename = "@maxDate")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub max_date: Option<String>,
    #[serde(rename = "@count")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub count: Option<u32>,
    #[serde(rename = "@longText")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub long_text: Option<bool>,
    #[serde(rename = "m")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub m: Vec<CTMissing>,
    #[serde(rename = "n")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub n: Vec<CTNumber>,
    #[serde(rename = "b")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub b: Vec<CTBoolean>,
    #[serde(rename = "e")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub e: Vec<CTError>,
    #[serde(rename = "s")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub style_index: Vec<CTString>,
    #[serde(rename = "d")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub d: Vec<CTDateTime>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct CTMissing {
    #[serde(rename = "@u")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub u: Option<bool>,
    #[serde(rename = "@f")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub formula: Option<bool>,
    #[serde(rename = "@c")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub cells: Option<XmlString>,
    #[serde(rename = "@cp")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub cp: Option<u32>,
    #[serde(rename = "@in")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub r#in: Option<u32>,
    #[serde(rename = "@bc")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub bc: Option<STUnsignedIntHex>,
    #[serde(rename = "@fc")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub fc: Option<STUnsignedIntHex>,
    #[serde(rename = "@i")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub i: Option<bool>,
    #[serde(rename = "@un")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub un: Option<bool>,
    #[serde(rename = "@st")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub st: Option<bool>,
    #[serde(rename = "@b")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub b: Option<bool>,
    #[serde(rename = "tpls")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub tpls: Vec<CTTuples>,
    #[serde(rename = "x")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub x: Vec<CTX>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTNumber {
    #[serde(rename = "@v")]
    pub value: f64,
    #[serde(rename = "@u")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub u: Option<bool>,
    #[serde(rename = "@f")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub formula: Option<bool>,
    #[serde(rename = "@c")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub cells: Option<XmlString>,
    #[serde(rename = "@cp")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub cp: Option<u32>,
    #[serde(rename = "@in")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub r#in: Option<u32>,
    #[serde(rename = "@bc")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub bc: Option<STUnsignedIntHex>,
    #[serde(rename = "@fc")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub fc: Option<STUnsignedIntHex>,
    #[serde(rename = "@i")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub i: Option<bool>,
    #[serde(rename = "@un")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub un: Option<bool>,
    #[serde(rename = "@st")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub st: Option<bool>,
    #[serde(rename = "@b")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub b: Option<bool>,
    #[serde(rename = "tpls")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub tpls: Vec<CTTuples>,
    #[serde(rename = "x")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub x: Vec<CTX>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTBoolean {
    #[serde(rename = "@v")]
    #[serde(with = "ooxml_xml::ooxml_bool_required")]
    pub value: bool,
    #[serde(rename = "@u")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub u: Option<bool>,
    #[serde(rename = "@f")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub formula: Option<bool>,
    #[serde(rename = "@c")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub cells: Option<XmlString>,
    #[serde(rename = "@cp")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub cp: Option<u32>,
    #[serde(rename = "x")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub x: Vec<CTX>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTError {
    #[serde(rename = "@v")]
    pub value: XmlString,
    #[serde(rename = "@u")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub u: Option<bool>,
    #[serde(rename = "@f")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub formula: Option<bool>,
    #[serde(rename = "@c")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub cells: Option<XmlString>,
    #[serde(rename = "@cp")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub cp: Option<u32>,
    #[serde(rename = "@in")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub r#in: Option<u32>,
    #[serde(rename = "@bc")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub bc: Option<STUnsignedIntHex>,
    #[serde(rename = "@fc")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub fc: Option<STUnsignedIntHex>,
    #[serde(rename = "@i")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub i: Option<bool>,
    #[serde(rename = "@un")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub un: Option<bool>,
    #[serde(rename = "@st")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub st: Option<bool>,
    #[serde(rename = "@b")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub b: Option<bool>,
    #[serde(rename = "tpls")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub tpls: Option<Box<CTTuples>>,
    #[serde(rename = "x")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub x: Vec<CTX>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTString {
    #[serde(rename = "@v")]
    pub value: XmlString,
    #[serde(rename = "@u")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub u: Option<bool>,
    #[serde(rename = "@f")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub formula: Option<bool>,
    #[serde(rename = "@c")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub cells: Option<XmlString>,
    #[serde(rename = "@cp")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub cp: Option<u32>,
    #[serde(rename = "@in")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub r#in: Option<u32>,
    #[serde(rename = "@bc")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub bc: Option<STUnsignedIntHex>,
    #[serde(rename = "@fc")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub fc: Option<STUnsignedIntHex>,
    #[serde(rename = "@i")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub i: Option<bool>,
    #[serde(rename = "@un")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub un: Option<bool>,
    #[serde(rename = "@st")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub st: Option<bool>,
    #[serde(rename = "@b")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub b: Option<bool>,
    #[serde(rename = "tpls")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub tpls: Vec<CTTuples>,
    #[serde(rename = "x")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub x: Vec<CTX>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTDateTime {
    #[serde(rename = "@v")]
    pub value: String,
    #[serde(rename = "@u")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub u: Option<bool>,
    #[serde(rename = "@f")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub formula: Option<bool>,
    #[serde(rename = "@c")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub cells: Option<XmlString>,
    #[serde(rename = "@cp")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub cp: Option<u32>,
    #[serde(rename = "x")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub x: Vec<CTX>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct FieldGroup {
    #[serde(rename = "@par")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub par: Option<u32>,
    #[serde(rename = "@base")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub base: Option<u32>,
    #[serde(rename = "rangePr")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub range_pr: Option<Box<CTRangePr>>,
    #[serde(rename = "discretePr")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub discrete_pr: Option<Box<CTDiscretePr>>,
    #[serde(rename = "groupItems")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub group_items: Option<Box<GroupItems>>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct CTRangePr {
    #[serde(rename = "@autoStart")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub auto_start: Option<bool>,
    #[serde(rename = "@autoEnd")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub auto_end: Option<bool>,
    #[serde(rename = "@groupBy")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub group_by: Option<STGroupBy>,
    #[serde(rename = "@startNum")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub start_num: Option<f64>,
    #[serde(rename = "@endNum")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub end_num: Option<f64>,
    #[serde(rename = "@startDate")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub start_date: Option<String>,
    #[serde(rename = "@endDate")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub end_date: Option<String>,
    #[serde(rename = "@groupInterval")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub group_interval: Option<f64>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct CTDiscretePr {
    #[serde(rename = "@count")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub count: Option<u32>,
    #[serde(rename = "x")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub x: Vec<CTIndex>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct GroupItems {
    #[serde(rename = "@count")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub count: Option<u32>,
    #[serde(rename = "m")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub m: Vec<CTMissing>,
    #[serde(rename = "n")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub n: Vec<CTNumber>,
    #[serde(rename = "b")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub b: Vec<CTBoolean>,
    #[serde(rename = "e")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub e: Vec<CTError>,
    #[serde(rename = "s")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub style_index: Vec<CTString>,
    #[serde(rename = "d")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub d: Vec<CTDateTime>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct PivotCacheRecords {
    #[serde(rename = "@count")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub count: Option<u32>,
    #[serde(rename = "r")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub reference: Vec<CTRecord>,
    #[serde(rename = "extLst")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub extension_list: Option<Box<ExtensionList>>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct CTRecord {
    #[serde(rename = "m")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub m: Vec<CTMissing>,
    #[serde(rename = "n")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub n: Vec<CTNumber>,
    #[serde(rename = "b")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub b: Vec<CTBoolean>,
    #[serde(rename = "e")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub e: Vec<CTError>,
    #[serde(rename = "s")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub style_index: Vec<CTString>,
    #[serde(rename = "d")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub d: Vec<CTDateTime>,
    #[serde(rename = "x")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub x: Vec<CTIndex>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct CTPCDKPIs {
    #[serde(rename = "@count")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub count: Option<u32>,
    #[serde(rename = "kpi")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub kpi: Vec<CTPCDKPI>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTPCDKPI {
    #[serde(rename = "@uniqueName")]
    pub unique_name: XmlString,
    #[serde(rename = "@caption")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub caption: Option<XmlString>,
    #[serde(rename = "@displayFolder")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub display_folder: Option<XmlString>,
    #[serde(rename = "@measureGroup")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub measure_group: Option<XmlString>,
    #[serde(rename = "@parent")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub parent: Option<XmlString>,
    #[serde(rename = "@value")]
    pub value: XmlString,
    #[serde(rename = "@goal")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub goal: Option<XmlString>,
    #[serde(rename = "@status")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub status: Option<XmlString>,
    #[serde(rename = "@trend")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub trend: Option<XmlString>,
    #[serde(rename = "@weight")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub weight: Option<XmlString>,
    #[serde(rename = "@time")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub time: Option<XmlString>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct CTCacheHierarchies {
    #[serde(rename = "@count")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub count: Option<u32>,
    #[serde(rename = "cacheHierarchy")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub cache_hierarchy: Vec<CTCacheHierarchy>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTCacheHierarchy {
    #[serde(rename = "@uniqueName")]
    pub unique_name: XmlString,
    #[serde(rename = "@caption")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub caption: Option<XmlString>,
    #[serde(rename = "@measure")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub measure: Option<bool>,
    #[serde(rename = "@set")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub set: Option<bool>,
    #[serde(rename = "@parentSet")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub parent_set: Option<u32>,
    #[serde(rename = "@iconSet")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub icon_set: Option<i32>,
    #[serde(rename = "@attribute")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub attribute: Option<bool>,
    #[serde(rename = "@time")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub time: Option<bool>,
    #[serde(rename = "@keyAttribute")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub key_attribute: Option<bool>,
    #[serde(rename = "@defaultMemberUniqueName")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub default_member_unique_name: Option<XmlString>,
    #[serde(rename = "@allUniqueName")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub all_unique_name: Option<XmlString>,
    #[serde(rename = "@allCaption")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub all_caption: Option<XmlString>,
    #[serde(rename = "@dimensionUniqueName")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub dimension_unique_name: Option<XmlString>,
    #[serde(rename = "@displayFolder")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub display_folder: Option<XmlString>,
    #[serde(rename = "@measureGroup")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub measure_group: Option<XmlString>,
    #[serde(rename = "@measures")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub measures: Option<bool>,
    #[serde(rename = "@count")]
    pub count: u32,
    #[serde(rename = "@oneField")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub one_field: Option<bool>,
    #[serde(rename = "@memberValueDatatype")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub member_value_datatype: Option<u16>,
    #[serde(rename = "@unbalanced")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub unbalanced: Option<bool>,
    #[serde(rename = "@unbalancedGroup")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub unbalanced_group: Option<bool>,
    #[serde(rename = "@hidden")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub hidden: Option<bool>,
    #[serde(rename = "fieldsUsage")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub fields_usage: Option<Box<CTFieldsUsage>>,
    #[serde(rename = "groupLevels")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub group_levels: Option<Box<CTGroupLevels>>,
    #[serde(rename = "extLst")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub extension_list: Option<Box<ExtensionList>>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct CTFieldsUsage {
    #[serde(rename = "@count")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub count: Option<u32>,
    #[serde(rename = "fieldUsage")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub field_usage: Vec<CTFieldUsage>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTFieldUsage {
    #[serde(rename = "@x")]
    pub x: i32,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct CTGroupLevels {
    #[serde(rename = "@count")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub count: Option<u32>,
    #[serde(rename = "groupLevel")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub group_level: Vec<CTGroupLevel>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTGroupLevel {
    #[serde(rename = "@uniqueName")]
    pub unique_name: XmlString,
    #[serde(rename = "@caption")]
    pub caption: XmlString,
    #[serde(rename = "@user")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub user: Option<bool>,
    #[serde(rename = "@customRollUp")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub custom_roll_up: Option<bool>,
    #[serde(rename = "groups")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub groups: Option<Box<CTGroups>>,
    #[serde(rename = "extLst")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub extension_list: Option<Box<ExtensionList>>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct CTGroups {
    #[serde(rename = "@count")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub count: Option<u32>,
    #[serde(rename = "group")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub group: Vec<CTLevelGroup>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTLevelGroup {
    #[serde(rename = "@name")]
    pub name: XmlString,
    #[serde(rename = "@uniqueName")]
    pub unique_name: XmlString,
    #[serde(rename = "@caption")]
    pub caption: XmlString,
    #[serde(rename = "@uniqueParent")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub unique_parent: Option<XmlString>,
    #[serde(rename = "@id")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub id: Option<i32>,
    #[serde(rename = "groupMembers")]
    pub group_members: Box<CTGroupMembers>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct CTGroupMembers {
    #[serde(rename = "@count")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub count: Option<u32>,
    #[serde(rename = "groupMember")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub group_member: Vec<CTGroupMember>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTGroupMember {
    #[serde(rename = "@uniqueName")]
    pub unique_name: XmlString,
    #[serde(rename = "@group")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub group: Option<bool>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct CTTupleCache {
    #[serde(rename = "entries")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub entries: Option<Box<CTPCDSDTCEntries>>,
    #[serde(rename = "sets")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub sets: Option<Box<CTSets>>,
    #[serde(rename = "queryCache")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub query_cache: Option<Box<CTQueryCache>>,
    #[serde(rename = "serverFormats")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub server_formats: Option<Box<CTServerFormats>>,
    #[serde(rename = "extLst")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub extension_list: Option<Box<ExtensionList>>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct CTServerFormat {
    #[serde(rename = "@culture")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub culture: Option<XmlString>,
    #[serde(rename = "@format")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub format: Option<XmlString>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct CTServerFormats {
    #[serde(rename = "@count")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub count: Option<u32>,
    #[serde(rename = "serverFormat")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub server_format: Vec<CTServerFormat>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct CTPCDSDTCEntries {
    #[serde(rename = "@count")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub count: Option<u32>,
    #[serde(rename = "m")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub m: Vec<CTMissing>,
    #[serde(rename = "n")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub n: Vec<CTNumber>,
    #[serde(rename = "e")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub e: Vec<CTError>,
    #[serde(rename = "s")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub style_index: Vec<CTString>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct CTTuples {
    #[serde(rename = "@c")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub cells: Option<u32>,
    #[serde(rename = "tpl")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub tpl: Vec<CTTuple>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTTuple {
    #[serde(rename = "@fld")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub fld: Option<u32>,
    #[serde(rename = "@hier")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub hier: Option<u32>,
    #[serde(rename = "@item")]
    pub item: u32,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct CTSets {
    #[serde(rename = "@count")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub count: Option<u32>,
    #[serde(rename = "set")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub set: Vec<CTSet>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTSet {
    #[serde(rename = "@count")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub count: Option<u32>,
    #[serde(rename = "@maxRank")]
    pub max_rank: i32,
    #[serde(rename = "@setDefinition")]
    pub set_definition: XmlString,
    #[serde(rename = "@sortType")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub sort_type: Option<STSortType>,
    #[serde(rename = "@queryFailed")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub query_failed: Option<bool>,
    #[serde(rename = "tpls")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub tpls: Vec<CTTuples>,
    #[serde(rename = "sortByTuple")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub sort_by_tuple: Option<Box<CTTuples>>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct CTQueryCache {
    #[serde(rename = "@count")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub count: Option<u32>,
    #[serde(rename = "query")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub query: Vec<CTQuery>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTQuery {
    #[serde(rename = "@mdx")]
    pub mdx: XmlString,
    #[serde(rename = "tpls")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub tpls: Option<Box<CTTuples>>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct CTCalculatedItems {
    #[serde(rename = "@count")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub count: Option<u32>,
    #[serde(rename = "calculatedItem")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub calculated_item: Vec<CTCalculatedItem>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTCalculatedItem {
    #[serde(rename = "@field")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub field: Option<u32>,
    #[serde(rename = "@formula")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub formula: Option<XmlString>,
    #[serde(rename = "pivotArea")]
    pub pivot_area: Box<PivotArea>,
    #[serde(rename = "extLst")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub extension_list: Option<Box<ExtensionList>>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct CTCalculatedMembers {
    #[serde(rename = "@count")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub count: Option<u32>,
    #[serde(rename = "calculatedMember")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub calculated_member: Vec<CTCalculatedMember>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTCalculatedMember {
    #[serde(rename = "@name")]
    pub name: XmlString,
    #[serde(rename = "@mdx")]
    pub mdx: XmlString,
    #[serde(rename = "@memberName")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub member_name: Option<XmlString>,
    #[serde(rename = "@hierarchy")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub hierarchy: Option<XmlString>,
    #[serde(rename = "@parent")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub parent: Option<XmlString>,
    #[serde(rename = "@solveOrder")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub solve_order: Option<i32>,
    #[serde(rename = "@set")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub set: Option<bool>,
    #[serde(rename = "extLst")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub extension_list: Option<Box<ExtensionList>>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTPivotTableDefinition {
    #[serde(rename = "@name")]
    pub name: XmlString,
    #[serde(rename = "@cacheId")]
    pub cache_id: u32,
    #[serde(rename = "@dataOnRows")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub data_on_rows: Option<bool>,
    #[serde(rename = "@dataPosition")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub data_position: Option<u32>,
    #[serde(rename = "@autoFormatId")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub auto_format_id: Option<u32>,
    #[serde(rename = "@applyNumberFormats")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub apply_number_formats: Option<bool>,
    #[serde(rename = "@applyBorderFormats")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub apply_border_formats: Option<bool>,
    #[serde(rename = "@applyFontFormats")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub apply_font_formats: Option<bool>,
    #[serde(rename = "@applyPatternFormats")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub apply_pattern_formats: Option<bool>,
    #[serde(rename = "@applyAlignmentFormats")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub apply_alignment_formats: Option<bool>,
    #[serde(rename = "@applyWidthHeightFormats")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub apply_width_height_formats: Option<bool>,
    #[serde(rename = "@dataCaption")]
    pub data_caption: XmlString,
    #[serde(rename = "@grandTotalCaption")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub grand_total_caption: Option<XmlString>,
    #[serde(rename = "@errorCaption")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub error_caption: Option<XmlString>,
    #[serde(rename = "@showError")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub show_error: Option<bool>,
    #[serde(rename = "@missingCaption")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub missing_caption: Option<XmlString>,
    #[serde(rename = "@showMissing")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub show_missing: Option<bool>,
    #[serde(rename = "@pageStyle")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub page_style: Option<XmlString>,
    #[serde(rename = "@pivotTableStyle")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub pivot_table_style: Option<XmlString>,
    #[serde(rename = "@vacatedStyle")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub vacated_style: Option<XmlString>,
    #[serde(rename = "@tag")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub tag: Option<XmlString>,
    #[serde(rename = "@updatedVersion")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub updated_version: Option<u8>,
    #[serde(rename = "@minRefreshableVersion")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub min_refreshable_version: Option<u8>,
    #[serde(rename = "@asteriskTotals")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub asterisk_totals: Option<bool>,
    #[serde(rename = "@showItems")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub show_items: Option<bool>,
    #[serde(rename = "@editData")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub edit_data: Option<bool>,
    #[serde(rename = "@disableFieldList")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub disable_field_list: Option<bool>,
    #[serde(rename = "@showCalcMbrs")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub show_calc_mbrs: Option<bool>,
    #[serde(rename = "@visualTotals")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub visual_totals: Option<bool>,
    #[serde(rename = "@showMultipleLabel")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub show_multiple_label: Option<bool>,
    #[serde(rename = "@showDataDropDown")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub show_data_drop_down: Option<bool>,
    #[serde(rename = "@showDrill")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub show_drill: Option<bool>,
    #[serde(rename = "@printDrill")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub print_drill: Option<bool>,
    #[serde(rename = "@showMemberPropertyTips")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub show_member_property_tips: Option<bool>,
    #[serde(rename = "@showDataTips")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub show_data_tips: Option<bool>,
    #[serde(rename = "@enableWizard")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub enable_wizard: Option<bool>,
    #[serde(rename = "@enableDrill")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub enable_drill: Option<bool>,
    #[serde(rename = "@enableFieldProperties")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub enable_field_properties: Option<bool>,
    #[serde(rename = "@preserveFormatting")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub preserve_formatting: Option<bool>,
    #[serde(rename = "@useAutoFormatting")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub use_auto_formatting: Option<bool>,
    #[serde(rename = "@pageWrap")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub page_wrap: Option<u32>,
    #[serde(rename = "@pageOverThenDown")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub page_over_then_down: Option<bool>,
    #[serde(rename = "@subtotalHiddenItems")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub subtotal_hidden_items: Option<bool>,
    #[serde(rename = "@rowGrandTotals")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub row_grand_totals: Option<bool>,
    #[serde(rename = "@colGrandTotals")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub col_grand_totals: Option<bool>,
    #[serde(rename = "@fieldPrintTitles")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub field_print_titles: Option<bool>,
    #[serde(rename = "@itemPrintTitles")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub item_print_titles: Option<bool>,
    #[serde(rename = "@mergeItem")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub merge_item: Option<bool>,
    #[serde(rename = "@showDropZones")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub show_drop_zones: Option<bool>,
    #[serde(rename = "@createdVersion")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub created_version: Option<u8>,
    #[serde(rename = "@indent")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub indent: Option<u32>,
    #[serde(rename = "@showEmptyRow")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub show_empty_row: Option<bool>,
    #[serde(rename = "@showEmptyCol")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub show_empty_col: Option<bool>,
    #[serde(rename = "@showHeaders")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub show_headers: Option<bool>,
    #[serde(rename = "@compact")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub compact: Option<bool>,
    #[serde(rename = "@outline")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub outline: Option<bool>,
    #[serde(rename = "@outlineData")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub outline_data: Option<bool>,
    #[serde(rename = "@compactData")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub compact_data: Option<bool>,
    #[serde(rename = "@published")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub published: Option<bool>,
    #[serde(rename = "@gridDropZones")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub grid_drop_zones: Option<bool>,
    #[serde(rename = "@immersive")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub immersive: Option<bool>,
    #[serde(rename = "@multipleFieldFilters")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub multiple_field_filters: Option<bool>,
    #[serde(rename = "@chartFormat")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub chart_format: Option<u32>,
    #[serde(rename = "@rowHeaderCaption")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub row_header_caption: Option<XmlString>,
    #[serde(rename = "@colHeaderCaption")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub col_header_caption: Option<XmlString>,
    #[serde(rename = "@fieldListSortAscending")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub field_list_sort_ascending: Option<bool>,
    #[serde(rename = "@mdxSubqueries")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub mdx_subqueries: Option<bool>,
    #[serde(rename = "@customListSort")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub custom_list_sort: Option<bool>,
    #[serde(rename = "location")]
    pub location: Box<PivotLocation>,
    #[serde(rename = "pivotFields")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub pivot_fields: Option<Box<PivotFields>>,
    #[serde(rename = "rowFields")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub row_fields: Option<Box<RowFields>>,
    #[serde(rename = "rowItems")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub row_items: Option<Box<CTRowItems>>,
    #[serde(rename = "colFields")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub col_fields: Option<Box<ColFields>>,
    #[serde(rename = "colItems")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub col_items: Option<Box<CTColItems>>,
    #[serde(rename = "pageFields")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub page_fields: Option<Box<PageFields>>,
    #[serde(rename = "dataFields")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub data_fields: Option<Box<DataFields>>,
    #[serde(rename = "formats")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub formats: Option<Box<CTFormats>>,
    #[serde(rename = "conditionalFormats")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub conditional_formats: Option<Box<CTConditionalFormats>>,
    #[serde(rename = "chartFormats")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub chart_formats: Option<Box<CTChartFormats>>,
    #[serde(rename = "pivotHierarchies")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub pivot_hierarchies: Option<Box<CTPivotHierarchies>>,
    #[serde(rename = "pivotTableStyleInfo")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub pivot_table_style_info: Option<Box<CTPivotTableStyle>>,
    #[serde(rename = "filters")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub filters: Option<Box<PivotFilters>>,
    #[serde(rename = "rowHierarchiesUsage")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub row_hierarchies_usage: Option<Box<CTRowHierarchiesUsage>>,
    #[serde(rename = "colHierarchiesUsage")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub col_hierarchies_usage: Option<Box<CTColHierarchiesUsage>>,
    #[serde(rename = "extLst")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub extension_list: Option<Box<ExtensionList>>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct PivotLocation {
    #[serde(rename = "@ref")]
    pub reference: Reference,
    #[serde(rename = "@firstHeaderRow")]
    pub first_header_row: u32,
    #[serde(rename = "@firstDataRow")]
    pub first_data_row: u32,
    #[serde(rename = "@firstDataCol")]
    pub first_data_col: u32,
    #[serde(rename = "@rowPageCount")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub row_page_count: Option<u32>,
    #[serde(rename = "@colPageCount")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub col_page_count: Option<u32>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct PivotFields {
    #[serde(rename = "@count")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub count: Option<u32>,
    #[serde(rename = "pivotField")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub pivot_field: Vec<PivotField>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct PivotField {
    #[serde(rename = "@name")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub name: Option<XmlString>,
    #[serde(rename = "@axis")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub axis: Option<STAxis>,
    #[serde(rename = "@dataField")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub data_field: Option<bool>,
    #[serde(rename = "@subtotalCaption")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub subtotal_caption: Option<XmlString>,
    #[serde(rename = "@showDropDowns")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub show_drop_downs: Option<bool>,
    #[serde(rename = "@hiddenLevel")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub hidden_level: Option<bool>,
    #[serde(rename = "@uniqueMemberProperty")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub unique_member_property: Option<XmlString>,
    #[serde(rename = "@compact")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub compact: Option<bool>,
    #[serde(rename = "@allDrilled")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub all_drilled: Option<bool>,
    #[serde(rename = "@numFmtId")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub number_format_id: Option<STNumFmtId>,
    #[serde(rename = "@outline")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub outline: Option<bool>,
    #[serde(rename = "@subtotalTop")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub subtotal_top: Option<bool>,
    #[serde(rename = "@dragToRow")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub drag_to_row: Option<bool>,
    #[serde(rename = "@dragToCol")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub drag_to_col: Option<bool>,
    #[serde(rename = "@multipleItemSelectionAllowed")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub multiple_item_selection_allowed: Option<bool>,
    #[serde(rename = "@dragToPage")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub drag_to_page: Option<bool>,
    #[serde(rename = "@dragToData")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub drag_to_data: Option<bool>,
    #[serde(rename = "@dragOff")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub drag_off: Option<bool>,
    #[serde(rename = "@showAll")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub show_all: Option<bool>,
    #[serde(rename = "@insertBlankRow")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub insert_blank_row: Option<bool>,
    #[serde(rename = "@serverField")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub server_field: Option<bool>,
    #[serde(rename = "@insertPageBreak")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub insert_page_break: Option<bool>,
    #[serde(rename = "@autoShow")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub auto_show: Option<bool>,
    #[serde(rename = "@topAutoShow")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub top_auto_show: Option<bool>,
    #[serde(rename = "@hideNewItems")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub hide_new_items: Option<bool>,
    #[serde(rename = "@measureFilter")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub measure_filter: Option<bool>,
    #[serde(rename = "@includeNewItemsInFilter")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub include_new_items_in_filter: Option<bool>,
    #[serde(rename = "@itemPageCount")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub item_page_count: Option<u32>,
    #[serde(rename = "@sortType")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub sort_type: Option<STFieldSortType>,
    #[serde(rename = "@dataSourceSort")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub data_source_sort: Option<bool>,
    #[serde(rename = "@nonAutoSortDefault")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub non_auto_sort_default: Option<bool>,
    #[serde(rename = "@rankBy")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub rank_by: Option<u32>,
    #[serde(rename = "@defaultSubtotal")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub default_subtotal: Option<bool>,
    #[serde(rename = "@sumSubtotal")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub sum_subtotal: Option<bool>,
    #[serde(rename = "@countASubtotal")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub count_a_subtotal: Option<bool>,
    #[serde(rename = "@avgSubtotal")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub avg_subtotal: Option<bool>,
    #[serde(rename = "@maxSubtotal")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub max_subtotal: Option<bool>,
    #[serde(rename = "@minSubtotal")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub min_subtotal: Option<bool>,
    #[serde(rename = "@productSubtotal")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub product_subtotal: Option<bool>,
    #[serde(rename = "@countSubtotal")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub count_subtotal: Option<bool>,
    #[serde(rename = "@stdDevSubtotal")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub std_dev_subtotal: Option<bool>,
    #[serde(rename = "@stdDevPSubtotal")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub std_dev_p_subtotal: Option<bool>,
    #[serde(rename = "@varSubtotal")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub var_subtotal: Option<bool>,
    #[serde(rename = "@varPSubtotal")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub var_p_subtotal: Option<bool>,
    #[serde(rename = "@showPropCell")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub show_prop_cell: Option<bool>,
    #[serde(rename = "@showPropTip")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub show_prop_tip: Option<bool>,
    #[serde(rename = "@showPropAsCaption")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub show_prop_as_caption: Option<bool>,
    #[serde(rename = "@defaultAttributeDrillState")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub default_attribute_drill_state: Option<bool>,
    #[serde(rename = "items")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub items: Option<Box<PivotItems>>,
    #[serde(rename = "autoSortScope")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub auto_sort_scope: Option<CTAutoSortScope>,
    #[serde(rename = "extLst")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub extension_list: Option<Box<ExtensionList>>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

pub type CTAutoSortScope = Box<PivotArea>;

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct PivotItems {
    #[serde(rename = "@count")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub count: Option<u32>,
    #[serde(rename = "item")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub item: Vec<PivotItem>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct PivotItem {
    #[serde(rename = "@n")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub n: Option<XmlString>,
    #[serde(rename = "@t")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub cell_type: Option<STItemType>,
    #[serde(rename = "@h")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub height: Option<bool>,
    #[serde(rename = "@s")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub style_index: Option<bool>,
    #[serde(rename = "@sd")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub sd: Option<bool>,
    #[serde(rename = "@f")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub formula: Option<bool>,
    #[serde(rename = "@m")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub m: Option<bool>,
    #[serde(rename = "@c")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub cells: Option<bool>,
    #[serde(rename = "@x")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub x: Option<u32>,
    #[serde(rename = "@d")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub d: Option<bool>,
    #[serde(rename = "@e")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub e: Option<bool>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct PageFields {
    #[serde(rename = "@count")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub count: Option<u32>,
    #[serde(rename = "pageField")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub page_field: Vec<PageField>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct PageField {
    #[serde(rename = "@fld")]
    pub fld: i32,
    #[serde(rename = "@item")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub item: Option<u32>,
    #[serde(rename = "@hier")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub hier: Option<i32>,
    #[serde(rename = "@name")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub name: Option<XmlString>,
    #[serde(rename = "@cap")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub cap: Option<XmlString>,
    #[serde(rename = "extLst")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub extension_list: Option<Box<ExtensionList>>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct DataFields {
    #[serde(rename = "@count")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub count: Option<u32>,
    #[serde(rename = "dataField")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub data_field: Vec<DataField>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct DataField {
    #[serde(rename = "@name")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub name: Option<XmlString>,
    #[serde(rename = "@fld")]
    pub fld: u32,
    #[serde(rename = "@subtotal")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub subtotal: Option<STDataConsolidateFunction>,
    #[serde(rename = "@showDataAs")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub show_data_as: Option<STShowDataAs>,
    #[serde(rename = "@baseField")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub base_field: Option<i32>,
    #[serde(rename = "@baseItem")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub base_item: Option<u32>,
    #[serde(rename = "@numFmtId")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub number_format_id: Option<STNumFmtId>,
    #[serde(rename = "extLst")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub extension_list: Option<Box<ExtensionList>>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct CTRowItems {
    #[serde(rename = "@count")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub count: Option<u32>,
    #[serde(rename = "i")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub i: Vec<CTI>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct CTColItems {
    #[serde(rename = "@count")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub count: Option<u32>,
    #[serde(rename = "i")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub i: Vec<CTI>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct CTI {
    #[serde(rename = "@t")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub cell_type: Option<STItemType>,
    #[serde(rename = "@r")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub reference: Option<u32>,
    #[serde(rename = "@i")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub i: Option<u32>,
    #[serde(rename = "x")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub x: Vec<CTX>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct CTX {
    #[serde(rename = "@v")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub value: Option<i32>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct RowFields {
    #[serde(rename = "@count")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub count: Option<u32>,
    #[serde(rename = "field")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub field: Vec<CTField>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct ColFields {
    #[serde(rename = "@count")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub count: Option<u32>,
    #[serde(rename = "field")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub field: Vec<CTField>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTField {
    #[serde(rename = "@x")]
    pub x: i32,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct CTFormats {
    #[serde(rename = "@count")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub count: Option<u32>,
    #[serde(rename = "format")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub format: Vec<CTFormat>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTFormat {
    #[serde(rename = "@action")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub action: Option<STFormatAction>,
    #[serde(rename = "@dxfId")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub dxf_id: Option<STDxfId>,
    #[serde(rename = "pivotArea")]
    pub pivot_area: Box<PivotArea>,
    #[serde(rename = "extLst")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub extension_list: Option<Box<ExtensionList>>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct CTConditionalFormats {
    #[serde(rename = "@count")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub count: Option<u32>,
    #[serde(rename = "conditionalFormat")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub conditional_format: Vec<CTConditionalFormat>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTConditionalFormat {
    #[serde(rename = "@scope")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub scope: Option<STScope>,
    #[serde(rename = "@type")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub r#type: Option<STType>,
    #[serde(rename = "@priority")]
    pub priority: u32,
    #[serde(rename = "pivotAreas")]
    pub pivot_areas: Box<PivotAreas>,
    #[serde(rename = "extLst")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub extension_list: Option<Box<ExtensionList>>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct PivotAreas {
    #[serde(rename = "@count")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub count: Option<u32>,
    #[serde(rename = "pivotArea")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub pivot_area: Vec<PivotArea>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct CTChartFormats {
    #[serde(rename = "@count")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub count: Option<u32>,
    #[serde(rename = "chartFormat")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub chart_format: Vec<CTChartFormat>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTChartFormat {
    #[serde(rename = "@chart")]
    pub chart: u32,
    #[serde(rename = "@format")]
    pub format: u32,
    #[serde(rename = "@series")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub series: Option<bool>,
    #[serde(rename = "pivotArea")]
    pub pivot_area: Box<PivotArea>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct CTPivotHierarchies {
    #[serde(rename = "@count")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub count: Option<u32>,
    #[serde(rename = "pivotHierarchy")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub pivot_hierarchy: Vec<CTPivotHierarchy>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct CTPivotHierarchy {
    #[serde(rename = "@outline")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub outline: Option<bool>,
    #[serde(rename = "@multipleItemSelectionAllowed")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub multiple_item_selection_allowed: Option<bool>,
    #[serde(rename = "@subtotalTop")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub subtotal_top: Option<bool>,
    #[serde(rename = "@showInFieldList")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub show_in_field_list: Option<bool>,
    #[serde(rename = "@dragToRow")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub drag_to_row: Option<bool>,
    #[serde(rename = "@dragToCol")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub drag_to_col: Option<bool>,
    #[serde(rename = "@dragToPage")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub drag_to_page: Option<bool>,
    #[serde(rename = "@dragToData")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub drag_to_data: Option<bool>,
    #[serde(rename = "@dragOff")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub drag_off: Option<bool>,
    #[serde(rename = "@includeNewItemsInFilter")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub include_new_items_in_filter: Option<bool>,
    #[serde(rename = "@caption")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub caption: Option<XmlString>,
    #[serde(rename = "mps")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub mps: Option<Box<CTMemberProperties>>,
    #[serde(rename = "members")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub members: Vec<CTMembers>,
    #[serde(rename = "extLst")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub extension_list: Option<Box<ExtensionList>>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct CTRowHierarchiesUsage {
    #[serde(rename = "@count")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub count: Option<u32>,
    #[serde(rename = "rowHierarchyUsage")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub row_hierarchy_usage: Vec<CTHierarchyUsage>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct CTColHierarchiesUsage {
    #[serde(rename = "@count")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub count: Option<u32>,
    #[serde(rename = "colHierarchyUsage")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub col_hierarchy_usage: Vec<CTHierarchyUsage>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTHierarchyUsage {
    #[serde(rename = "@hierarchyUsage")]
    pub hierarchy_usage: i32,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct CTMemberProperties {
    #[serde(rename = "@count")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub count: Option<u32>,
    #[serde(rename = "mp")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub mp: Vec<CTMemberProperty>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTMemberProperty {
    #[serde(rename = "@name")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub name: Option<XmlString>,
    #[serde(rename = "@showCell")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub show_cell: Option<bool>,
    #[serde(rename = "@showTip")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub show_tip: Option<bool>,
    #[serde(rename = "@showAsCaption")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub show_as_caption: Option<bool>,
    #[serde(rename = "@nameLen")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub name_len: Option<u32>,
    #[serde(rename = "@pPos")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub p_pos: Option<u32>,
    #[serde(rename = "@pLen")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub p_len: Option<u32>,
    #[serde(rename = "@level")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub level: Option<u32>,
    #[serde(rename = "@field")]
    pub field: u32,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct CTMembers {
    #[serde(rename = "@count")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub count: Option<u32>,
    #[serde(rename = "@level")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub level: Option<u32>,
    #[serde(rename = "member")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub member: Vec<CTMember>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTMember {
    #[serde(rename = "@name")]
    pub name: XmlString,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct CTDimensions {
    #[serde(rename = "@count")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub count: Option<u32>,
    #[serde(rename = "dimension")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub dimension: Vec<CTPivotDimension>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTPivotDimension {
    #[serde(rename = "@measure")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub measure: Option<bool>,
    #[serde(rename = "@name")]
    pub name: XmlString,
    #[serde(rename = "@uniqueName")]
    pub unique_name: XmlString,
    #[serde(rename = "@caption")]
    pub caption: XmlString,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct CTMeasureGroups {
    #[serde(rename = "@count")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub count: Option<u32>,
    #[serde(rename = "measureGroup")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub measure_group: Vec<CTMeasureGroup>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct CTMeasureDimensionMaps {
    #[serde(rename = "@count")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub count: Option<u32>,
    #[serde(rename = "map")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub map: Vec<CTMeasureDimensionMap>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTMeasureGroup {
    #[serde(rename = "@name")]
    pub name: XmlString,
    #[serde(rename = "@caption")]
    pub caption: XmlString,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct CTMeasureDimensionMap {
    #[serde(rename = "@measureGroup")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub measure_group: Option<u32>,
    #[serde(rename = "@dimension")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub dimension: Option<u32>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct CTPivotTableStyle {
    #[serde(rename = "@name")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub name: Option<String>,
    #[serde(rename = "@showRowHeaders")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub show_row_headers: Option<bool>,
    #[serde(rename = "@showColHeaders")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub show_col_headers: Option<bool>,
    #[serde(rename = "@showRowStripes")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub show_row_stripes: Option<bool>,
    #[serde(rename = "@showColStripes")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub show_col_stripes: Option<bool>,
    #[serde(rename = "@showLastColumn")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub show_last_column: Option<bool>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct PivotFilters {
    #[serde(rename = "@count")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub count: Option<u32>,
    #[serde(rename = "filter")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub filter: Vec<PivotFilter>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct PivotFilter {
    #[serde(rename = "@fld")]
    pub fld: u32,
    #[serde(rename = "@mpFld")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub mp_fld: Option<u32>,
    #[serde(rename = "@type")]
    pub r#type: STPivotFilterType,
    #[serde(rename = "@evalOrder")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub eval_order: Option<i32>,
    #[serde(rename = "@id")]
    pub id: u32,
    #[serde(rename = "@iMeasureHier")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub i_measure_hier: Option<u32>,
    #[serde(rename = "@iMeasureFld")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub i_measure_fld: Option<u32>,
    #[serde(rename = "@name")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub name: Option<XmlString>,
    #[serde(rename = "@description")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub description: Option<XmlString>,
    #[serde(rename = "@stringValue1")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub string_value1: Option<XmlString>,
    #[serde(rename = "@stringValue2")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub string_value2: Option<XmlString>,
    #[serde(rename = "autoFilter")]
    pub auto_filter: Box<AutoFilter>,
    #[serde(rename = "extLst")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub extension_list: Option<Box<ExtensionList>>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct PivotArea {
    #[serde(rename = "@field")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub field: Option<i32>,
    #[serde(rename = "@type")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub r#type: Option<STPivotAreaType>,
    #[serde(rename = "@dataOnly")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub data_only: Option<bool>,
    #[serde(rename = "@labelOnly")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub label_only: Option<bool>,
    #[serde(rename = "@grandRow")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub grand_row: Option<bool>,
    #[serde(rename = "@grandCol")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub grand_col: Option<bool>,
    #[serde(rename = "@cacheIndex")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub cache_index: Option<bool>,
    #[serde(rename = "@outline")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub outline: Option<bool>,
    #[serde(rename = "@offset")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub offset: Option<Reference>,
    #[serde(rename = "@collapsedLevelsAreSubtotals")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub collapsed_levels_are_subtotals: Option<bool>,
    #[serde(rename = "@axis")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub axis: Option<STAxis>,
    #[serde(rename = "@fieldPosition")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub field_position: Option<u32>,
    #[serde(rename = "references")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub references: Option<Box<CTPivotAreaReferences>>,
    #[serde(rename = "extLst")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub extension_list: Option<Box<ExtensionList>>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct CTPivotAreaReferences {
    #[serde(rename = "@count")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub count: Option<u32>,
    #[serde(rename = "reference")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub reference: Vec<CTPivotAreaReference>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct CTPivotAreaReference {
    #[serde(rename = "@field")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub field: Option<u32>,
    #[serde(rename = "@count")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub count: Option<u32>,
    #[serde(rename = "@selected")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub selected: Option<bool>,
    #[serde(rename = "@byPosition")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub by_position: Option<bool>,
    #[serde(rename = "@relative")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub relative: Option<bool>,
    #[serde(rename = "@defaultSubtotal")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub default_subtotal: Option<bool>,
    #[serde(rename = "@sumSubtotal")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub sum_subtotal: Option<bool>,
    #[serde(rename = "@countASubtotal")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub count_a_subtotal: Option<bool>,
    #[serde(rename = "@avgSubtotal")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub avg_subtotal: Option<bool>,
    #[serde(rename = "@maxSubtotal")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub max_subtotal: Option<bool>,
    #[serde(rename = "@minSubtotal")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub min_subtotal: Option<bool>,
    #[serde(rename = "@productSubtotal")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub product_subtotal: Option<bool>,
    #[serde(rename = "@countSubtotal")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub count_subtotal: Option<bool>,
    #[serde(rename = "@stdDevSubtotal")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub std_dev_subtotal: Option<bool>,
    #[serde(rename = "@stdDevPSubtotal")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub std_dev_p_subtotal: Option<bool>,
    #[serde(rename = "@varSubtotal")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub var_subtotal: Option<bool>,
    #[serde(rename = "@varPSubtotal")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub var_p_subtotal: Option<bool>,
    #[serde(rename = "x")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub x: Vec<CTIndex>,
    #[serde(rename = "extLst")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub extension_list: Option<Box<ExtensionList>>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTIndex {
    #[serde(rename = "@v")]
    pub value: u32,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

pub type SmlQueryTable = Box<QueryTable>;

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct QueryTable {
    #[serde(rename = "@name")]
    pub name: XmlString,
    #[serde(rename = "@headers")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub headers: Option<bool>,
    #[serde(rename = "@rowNumbers")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub row_numbers: Option<bool>,
    #[serde(rename = "@disableRefresh")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub disable_refresh: Option<bool>,
    #[serde(rename = "@backgroundRefresh")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub background_refresh: Option<bool>,
    #[serde(rename = "@firstBackgroundRefresh")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub first_background_refresh: Option<bool>,
    #[serde(rename = "@refreshOnLoad")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub refresh_on_load: Option<bool>,
    #[serde(rename = "@growShrinkType")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub grow_shrink_type: Option<STGrowShrinkType>,
    #[serde(rename = "@fillFormulas")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub fill_formulas: Option<bool>,
    #[serde(rename = "@removeDataOnSave")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub remove_data_on_save: Option<bool>,
    #[serde(rename = "@disableEdit")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub disable_edit: Option<bool>,
    #[serde(rename = "@preserveFormatting")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub preserve_formatting: Option<bool>,
    #[serde(rename = "@adjustColumnWidth")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub adjust_column_width: Option<bool>,
    #[serde(rename = "@intermediate")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub intermediate: Option<bool>,
    #[serde(rename = "@connectionId")]
    pub connection_id: u32,
    #[serde(rename = "@autoFormatId")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub auto_format_id: Option<u32>,
    #[serde(rename = "@applyNumberFormats")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub apply_number_formats: Option<bool>,
    #[serde(rename = "@applyBorderFormats")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub apply_border_formats: Option<bool>,
    #[serde(rename = "@applyFontFormats")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub apply_font_formats: Option<bool>,
    #[serde(rename = "@applyPatternFormats")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub apply_pattern_formats: Option<bool>,
    #[serde(rename = "@applyAlignmentFormats")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub apply_alignment_formats: Option<bool>,
    #[serde(rename = "@applyWidthHeightFormats")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub apply_width_height_formats: Option<bool>,
    #[serde(rename = "queryTableRefresh")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub query_table_refresh: Option<Box<QueryTableRefresh>>,
    #[serde(rename = "extLst")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub extension_list: Option<Box<ExtensionList>>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct QueryTableRefresh {
    #[serde(rename = "@preserveSortFilterLayout")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub preserve_sort_filter_layout: Option<bool>,
    #[serde(rename = "@fieldIdWrapped")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub field_id_wrapped: Option<bool>,
    #[serde(rename = "@headersInLastRefresh")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub headers_in_last_refresh: Option<bool>,
    #[serde(rename = "@minimumVersion")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub minimum_version: Option<u8>,
    #[serde(rename = "@nextId")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub next_id: Option<u32>,
    #[serde(rename = "@unboundColumnsLeft")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub unbound_columns_left: Option<u32>,
    #[serde(rename = "@unboundColumnsRight")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub unbound_columns_right: Option<u32>,
    #[serde(rename = "queryTableFields")]
    pub query_table_fields: Box<QueryTableFields>,
    #[serde(rename = "queryTableDeletedFields")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub query_table_deleted_fields: Option<Box<QueryTableDeletedFields>>,
    #[serde(rename = "sortState")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub sort_state: Option<Box<SortState>>,
    #[serde(rename = "extLst")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub extension_list: Option<Box<ExtensionList>>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct QueryTableDeletedFields {
    #[serde(rename = "@count")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub count: Option<u32>,
    #[serde(rename = "deletedField")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub deleted_field: Vec<CTDeletedField>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTDeletedField {
    #[serde(rename = "@name")]
    pub name: XmlString,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct QueryTableFields {
    #[serde(rename = "@count")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub count: Option<u32>,
    #[serde(rename = "queryTableField")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub query_table_field: Vec<QueryTableField>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct QueryTableField {
    #[serde(rename = "@id")]
    pub id: u32,
    #[serde(rename = "@name")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub name: Option<XmlString>,
    #[serde(rename = "@dataBound")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub data_bound: Option<bool>,
    #[serde(rename = "@rowNumbers")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub row_numbers: Option<bool>,
    #[serde(rename = "@fillFormulas")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub fill_formulas: Option<bool>,
    #[serde(rename = "@clipped")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub clipped: Option<bool>,
    #[serde(rename = "@tableColumnId")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub table_column_id: Option<u32>,
    #[serde(rename = "extLst")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub extension_list: Option<Box<ExtensionList>>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

pub type SmlSst = Box<SharedStrings>;

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
#[serde(rename = "sst")]
pub struct SharedStrings {
    #[serde(rename = "@count")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub count: Option<u32>,
    #[serde(rename = "@uniqueCount")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub unique_count: Option<u32>,
    #[serde(rename = "si")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub si: Vec<RichString>,
    #[serde(rename = "extLst")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub extension_list: Option<Box<ExtensionList>>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct PhoneticRun {
    #[serde(rename = "@sb")]
    pub sb: u32,
    #[serde(rename = "@eb")]
    pub eb: u32,
    #[serde(rename = "t")]
    pub cell_type: XmlString,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(rename = "r")]
pub struct RichTextElement {
    #[serde(rename = "rPr")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub r_pr: Option<Box<RichTextRunProperties>>,
    #[serde(rename = "t")]
    pub cell_type: XmlString,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
#[serde(rename = "rPr")]
pub struct RichTextRunProperties {
    #[serde(rename = "rFont")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub r_font: Option<Box<FontName>>,
    #[serde(rename = "charset")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub charset: Option<Box<IntProperty>>,
    #[serde(rename = "family")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub family: Option<Box<IntProperty>>,
    #[serde(rename = "b")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub b: Option<Box<BooleanProperty>>,
    #[serde(rename = "i")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub i: Option<Box<BooleanProperty>>,
    #[serde(rename = "strike")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub strike: Option<Box<BooleanProperty>>,
    #[serde(rename = "outline")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub outline: Option<Box<BooleanProperty>>,
    #[serde(rename = "shadow")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub shadow: Option<Box<BooleanProperty>>,
    #[serde(rename = "condense")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub condense: Option<Box<BooleanProperty>>,
    #[serde(rename = "extend")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub extend: Option<Box<BooleanProperty>>,
    #[serde(rename = "color")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub color: Option<Box<Color>>,
    #[serde(rename = "sz")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub sz: Option<Box<FontSize>>,
    #[serde(rename = "u")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub u: Option<Box<UnderlineProperty>>,
    #[serde(rename = "vertAlign")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub vert_align: Option<Box<VerticalAlignFontProperty>>,
    #[serde(rename = "scheme")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub scheme: Option<Box<FontSchemeProperty>>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
#[serde(rename = "is")]
pub struct RichString {
    #[serde(rename = "t")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub cell_type: Option<XmlString>,
    #[serde(rename = "r")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub reference: Vec<RichTextElement>,
    #[serde(rename = "rPh")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub r_ph: Vec<PhoneticRun>,
    #[serde(rename = "phoneticPr")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub phonetic_pr: Option<Box<PhoneticProperties>>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct PhoneticProperties {
    #[serde(rename = "@fontId")]
    pub font_id: STFontId,
    #[serde(rename = "@type")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub r#type: Option<STPhoneticType>,
    #[serde(rename = "@alignment")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub alignment: Option<STPhoneticAlignment>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

pub type SmlHeaders = Box<RevisionHeaders>;

pub type SmlRevisions = Box<Revisions>;

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct RevisionHeaders {
    #[serde(rename = "@guid")]
    pub guid: Guid,
    #[serde(rename = "@lastGuid")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub last_guid: Option<Guid>,
    #[serde(rename = "@shared")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub shared: Option<bool>,
    #[serde(rename = "@diskRevisions")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub disk_revisions: Option<bool>,
    #[serde(rename = "@history")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub history: Option<bool>,
    #[serde(rename = "@trackRevisions")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub track_revisions: Option<bool>,
    #[serde(rename = "@exclusive")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub exclusive: Option<bool>,
    #[serde(rename = "@revisionId")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub revision_id: Option<u32>,
    #[serde(rename = "@version")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub version: Option<i32>,
    #[serde(rename = "@keepChangeHistory")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub keep_change_history: Option<bool>,
    #[serde(rename = "@protected")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub protected: Option<bool>,
    #[serde(rename = "@preserveHistory")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub preserve_history: Option<u32>,
    #[serde(rename = "header")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub header: Vec<RevisionHeader>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct Revisions {
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct SmlAGRevData {
    #[serde(rename = "@rId")]
    pub r_id: u32,
    #[serde(rename = "@ua")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub ua: Option<bool>,
    #[serde(rename = "@ra")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub ra: Option<bool>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct RevisionHeader {
    #[serde(rename = "@guid")]
    pub guid: Guid,
    #[serde(rename = "@dateTime")]
    pub date_time: String,
    #[serde(rename = "@maxSheetId")]
    pub max_sheet_id: u32,
    #[serde(rename = "@userName")]
    pub user_name: XmlString,
    #[serde(rename = "@r:id")]
    pub id: STRelationshipId,
    #[serde(rename = "@minRId")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub min_r_id: Option<u32>,
    #[serde(rename = "@maxRId")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub max_r_id: Option<u32>,
    #[serde(rename = "sheetIdMap")]
    pub sheet_id_map: Box<CTSheetIdMap>,
    #[serde(rename = "reviewedList")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub reviewed_list: Option<Box<ReviewedRevisions>>,
    #[serde(rename = "extLst")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub extension_list: Option<Box<ExtensionList>>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct CTSheetIdMap {
    #[serde(rename = "@count")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub count: Option<u32>,
    #[serde(rename = "sheetId")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub sheet_id: Vec<CTSheetId>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTSheetId {
    #[serde(rename = "@val")]
    pub value: u32,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct ReviewedRevisions {
    #[serde(rename = "@count")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub count: Option<u32>,
    #[serde(rename = "reviewed")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub reviewed: Vec<Reviewed>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct Reviewed {
    #[serde(rename = "@rId")]
    pub r_id: u32,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct UndoInfo {
    #[serde(rename = "@index")]
    pub index: u32,
    #[serde(rename = "@exp")]
    pub exp: STFormulaExpression,
    #[serde(rename = "@ref3D")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub ref3_d: Option<bool>,
    #[serde(rename = "@array")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub array: Option<bool>,
    #[serde(rename = "@v")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub value: Option<bool>,
    #[serde(rename = "@nf")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub nf: Option<bool>,
    #[serde(rename = "@cs")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub cs: Option<bool>,
    #[serde(rename = "@dr")]
    pub dr: STRefA,
    #[serde(rename = "@dn")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub dn: Option<XmlString>,
    #[serde(rename = "@r")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub reference: Option<CellRef>,
    #[serde(rename = "@sId")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub s_id: Option<u32>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct RevisionRowColumn {
    #[serde(rename = "@rId")]
    pub r_id: u32,
    #[serde(rename = "@ua")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub ua: Option<bool>,
    #[serde(rename = "@ra")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub ra: Option<bool>,
    #[serde(rename = "@sId")]
    pub s_id: u32,
    #[serde(rename = "@eol")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub eol: Option<bool>,
    #[serde(rename = "@ref")]
    pub reference: Reference,
    #[serde(rename = "@action")]
    pub action: STRwColActionType,
    #[serde(rename = "@edge")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub edge: Option<bool>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct RevisionMove {
    #[serde(rename = "@rId")]
    pub r_id: u32,
    #[serde(rename = "@ua")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub ua: Option<bool>,
    #[serde(rename = "@ra")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub ra: Option<bool>,
    #[serde(rename = "@sheetId")]
    pub sheet_id: u32,
    #[serde(rename = "@source")]
    pub source: Reference,
    #[serde(rename = "@destination")]
    pub destination: Reference,
    #[serde(rename = "@sourceSheetId")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub source_sheet_id: Option<u32>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct RevisionCustomView {
    #[serde(rename = "@guid")]
    pub guid: Guid,
    #[serde(rename = "@action")]
    pub action: STRevisionAction,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct RevisionSheetRename {
    #[serde(rename = "@rId")]
    pub r_id: u32,
    #[serde(rename = "@ua")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub ua: Option<bool>,
    #[serde(rename = "@ra")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub ra: Option<bool>,
    #[serde(rename = "@sheetId")]
    pub sheet_id: u32,
    #[serde(rename = "@oldName")]
    pub old_name: XmlString,
    #[serde(rename = "@newName")]
    pub new_name: XmlString,
    #[serde(rename = "extLst")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub extension_list: Option<Box<ExtensionList>>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct RevisionInsertSheet {
    #[serde(rename = "@rId")]
    pub r_id: u32,
    #[serde(rename = "@ua")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub ua: Option<bool>,
    #[serde(rename = "@ra")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub ra: Option<bool>,
    #[serde(rename = "@sheetId")]
    pub sheet_id: u32,
    #[serde(rename = "@name")]
    pub name: XmlString,
    #[serde(rename = "@sheetPosition")]
    pub sheet_position: u32,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct RevisionCellChange {
    #[serde(rename = "@rId")]
    pub r_id: u32,
    #[serde(rename = "@ua")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub ua: Option<bool>,
    #[serde(rename = "@ra")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub ra: Option<bool>,
    #[serde(rename = "@sId")]
    pub s_id: u32,
    #[serde(rename = "@odxf")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub odxf: Option<bool>,
    #[serde(rename = "@xfDxf")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub xf_dxf: Option<bool>,
    #[serde(rename = "@s")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub style_index: Option<bool>,
    #[serde(rename = "@dxf")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub dxf: Option<bool>,
    #[serde(rename = "@numFmtId")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub number_format_id: Option<STNumFmtId>,
    #[serde(rename = "@quotePrefix")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub quote_prefix: Option<bool>,
    #[serde(rename = "@oldQuotePrefix")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub old_quote_prefix: Option<bool>,
    #[serde(rename = "@ph")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub placeholder: Option<bool>,
    #[serde(rename = "@oldPh")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub old_ph: Option<bool>,
    #[serde(rename = "@endOfListFormulaUpdate")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub end_of_list_formula_update: Option<bool>,
    #[serde(rename = "oc")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub oc: Option<Box<Cell>>,
    #[serde(rename = "nc")]
    pub nc: Box<Cell>,
    #[serde(rename = "ndxf")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub ndxf: Option<Box<DifferentialFormat>>,
    #[serde(rename = "extLst")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub extension_list: Option<Box<ExtensionList>>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct RevisionFormatting {
    #[serde(rename = "@sheetId")]
    pub sheet_id: u32,
    #[serde(rename = "@xfDxf")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub xf_dxf: Option<bool>,
    #[serde(rename = "@s")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub style_index: Option<bool>,
    #[serde(rename = "@sqref")]
    pub square_reference: SquareRef,
    #[serde(rename = "@start")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub start: Option<u32>,
    #[serde(rename = "@length")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub length: Option<u32>,
    #[serde(rename = "dxf")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub dxf: Option<Box<DifferentialFormat>>,
    #[serde(rename = "extLst")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub extension_list: Option<Box<ExtensionList>>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct RevisionAutoFormatting {
    #[serde(rename = "@sheetId")]
    pub sheet_id: u32,
    #[serde(rename = "@autoFormatId")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub auto_format_id: Option<u32>,
    #[serde(rename = "@applyNumberFormats")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub apply_number_formats: Option<bool>,
    #[serde(rename = "@applyBorderFormats")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub apply_border_formats: Option<bool>,
    #[serde(rename = "@applyFontFormats")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub apply_font_formats: Option<bool>,
    #[serde(rename = "@applyPatternFormats")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub apply_pattern_formats: Option<bool>,
    #[serde(rename = "@applyAlignmentFormats")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub apply_alignment_formats: Option<bool>,
    #[serde(rename = "@applyWidthHeightFormats")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub apply_width_height_formats: Option<bool>,
    #[serde(rename = "@ref")]
    pub reference: Reference,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct RevisionComment {
    #[serde(rename = "@sheetId")]
    pub sheet_id: u32,
    #[serde(rename = "@cell")]
    pub cell: CellRef,
    #[serde(rename = "@guid")]
    pub guid: Guid,
    #[serde(rename = "@action")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub action: Option<STRevisionAction>,
    #[serde(rename = "@alwaysShow")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub always_show: Option<bool>,
    #[serde(rename = "@old")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub old: Option<bool>,
    #[serde(rename = "@hiddenRow")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub hidden_row: Option<bool>,
    #[serde(rename = "@hiddenColumn")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub hidden_column: Option<bool>,
    #[serde(rename = "@author")]
    pub author: XmlString,
    #[serde(rename = "@oldLength")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub old_length: Option<u32>,
    #[serde(rename = "@newLength")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub new_length: Option<u32>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct RevisionDefinedName {
    #[serde(rename = "@rId")]
    pub r_id: u32,
    #[serde(rename = "@ua")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub ua: Option<bool>,
    #[serde(rename = "@ra")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub ra: Option<bool>,
    #[serde(rename = "@localSheetId")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub local_sheet_id: Option<u32>,
    #[serde(rename = "@customView")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub custom_view: Option<bool>,
    #[serde(rename = "@name")]
    pub name: XmlString,
    #[serde(rename = "@function")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub function: Option<bool>,
    #[serde(rename = "@oldFunction")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub old_function: Option<bool>,
    #[serde(rename = "@functionGroupId")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub function_group_id: Option<u8>,
    #[serde(rename = "@oldFunctionGroupId")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub old_function_group_id: Option<u8>,
    #[serde(rename = "@shortcutKey")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub shortcut_key: Option<u8>,
    #[serde(rename = "@oldShortcutKey")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub old_shortcut_key: Option<u8>,
    #[serde(rename = "@hidden")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub hidden: Option<bool>,
    #[serde(rename = "@oldHidden")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub old_hidden: Option<bool>,
    #[serde(rename = "@customMenu")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub custom_menu: Option<XmlString>,
    #[serde(rename = "@oldCustomMenu")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub old_custom_menu: Option<XmlString>,
    #[serde(rename = "@description")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub description: Option<XmlString>,
    #[serde(rename = "@oldDescription")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub old_description: Option<XmlString>,
    #[serde(rename = "@help")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub help: Option<XmlString>,
    #[serde(rename = "@oldHelp")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub old_help: Option<XmlString>,
    #[serde(rename = "@statusBar")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub status_bar: Option<XmlString>,
    #[serde(rename = "@oldStatusBar")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub old_status_bar: Option<XmlString>,
    #[serde(rename = "@comment")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub comment: Option<XmlString>,
    #[serde(rename = "@oldComment")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub old_comment: Option<XmlString>,
    #[serde(rename = "formula")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub formula: Option<STFormula>,
    #[serde(rename = "oldFormula")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub old_formula: Option<STFormula>,
    #[serde(rename = "extLst")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub extension_list: Option<Box<ExtensionList>>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct RevisionConflict {
    #[serde(rename = "@rId")]
    pub r_id: u32,
    #[serde(rename = "@ua")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub ua: Option<bool>,
    #[serde(rename = "@ra")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub ra: Option<bool>,
    #[serde(rename = "@sheetId")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub sheet_id: Option<u32>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct RevisionQueryTableField {
    #[serde(rename = "@sheetId")]
    pub sheet_id: u32,
    #[serde(rename = "@ref")]
    pub reference: Reference,
    #[serde(rename = "@fieldId")]
    pub field_id: u32,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

pub type SmlUsers = Box<Users>;

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct Users {
    #[serde(rename = "@count")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub count: Option<u32>,
    #[serde(rename = "userInfo")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub user_info: Vec<SharedUser>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct SharedUser {
    #[serde(rename = "@guid")]
    pub guid: Guid,
    #[serde(rename = "@name")]
    pub name: XmlString,
    #[serde(rename = "@id")]
    pub id: i32,
    #[serde(rename = "@dateTime")]
    pub date_time: String,
    #[serde(rename = "extLst")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub extension_list: Option<Box<ExtensionList>>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

pub type SmlWorksheet = Box<Worksheet>;

pub type SmlChartsheet = Box<Chartsheet>;

pub type SmlDialogsheet = Box<CTDialogsheet>;

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTMacrosheet {
    #[serde(rename = "sheetPr")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub sheet_properties: Option<Box<SheetProperties>>,
    #[serde(rename = "dimension")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub dimension: Option<Box<SheetDimension>>,
    #[serde(rename = "sheetViews")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub sheet_views: Option<Box<SheetViews>>,
    #[serde(rename = "sheetFormatPr")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub sheet_format: Option<Box<SheetFormat>>,
    #[serde(rename = "cols")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub cols: Vec<Columns>,
    #[serde(rename = "sheetData")]
    pub sheet_data: Box<SheetData>,
    #[serde(rename = "sheetProtection")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub sheet_protection: Option<Box<SheetProtection>>,
    #[serde(rename = "autoFilter")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub auto_filter: Option<Box<AutoFilter>>,
    #[serde(rename = "sortState")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub sort_state: Option<Box<SortState>>,
    #[serde(rename = "dataConsolidate")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub data_consolidate: Option<Box<CTDataConsolidate>>,
    #[serde(rename = "customSheetViews")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub custom_sheet_views: Option<Box<CustomSheetViews>>,
    #[serde(rename = "phoneticPr")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub phonetic_pr: Option<Box<PhoneticProperties>>,
    #[serde(rename = "conditionalFormatting")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub conditional_formatting: Vec<ConditionalFormatting>,
    #[serde(rename = "printOptions")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub print_options: Option<Box<PrintOptions>>,
    #[serde(rename = "pageMargins")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub page_margins: Option<Box<PageMargins>>,
    #[serde(rename = "pageSetup")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub page_setup: Option<Box<PageSetup>>,
    #[serde(rename = "headerFooter")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub header_footer: Option<Box<HeaderFooter>>,
    #[serde(rename = "rowBreaks")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub row_breaks: Option<Box<PageBreaks>>,
    #[serde(rename = "colBreaks")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub col_breaks: Option<Box<PageBreaks>>,
    #[serde(rename = "customProperties")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub custom_properties: Option<Box<CTCustomProperties>>,
    #[serde(rename = "drawing")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub drawing: Option<Box<Drawing>>,
    #[serde(rename = "legacyDrawing")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub legacy_drawing: Option<Box<LegacyDrawing>>,
    #[serde(rename = "legacyDrawingHF")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub legacy_drawing_h_f: Option<Box<LegacyDrawing>>,
    #[serde(rename = "drawingHF")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub drawing_h_f: Option<Box<DrawingHeaderFooter>>,
    #[serde(rename = "picture")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub picture: Option<Box<SheetBackgroundPicture>>,
    #[serde(rename = "oleObjects")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub ole_objects: Option<Box<OleObjects>>,
    #[serde(rename = "extLst")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub extension_list: Option<Box<ExtensionList>>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct CTDialogsheet {
    #[serde(rename = "sheetPr")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub sheet_properties: Option<Box<SheetProperties>>,
    #[serde(rename = "sheetViews")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub sheet_views: Option<Box<SheetViews>>,
    #[serde(rename = "sheetFormatPr")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub sheet_format: Option<Box<SheetFormat>>,
    #[serde(rename = "sheetProtection")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub sheet_protection: Option<Box<SheetProtection>>,
    #[serde(rename = "customSheetViews")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub custom_sheet_views: Option<Box<CustomSheetViews>>,
    #[serde(rename = "printOptions")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub print_options: Option<Box<PrintOptions>>,
    #[serde(rename = "pageMargins")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub page_margins: Option<Box<PageMargins>>,
    #[serde(rename = "pageSetup")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub page_setup: Option<Box<PageSetup>>,
    #[serde(rename = "headerFooter")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub header_footer: Option<Box<HeaderFooter>>,
    #[serde(rename = "drawing")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub drawing: Option<Box<Drawing>>,
    #[serde(rename = "legacyDrawing")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub legacy_drawing: Option<Box<LegacyDrawing>>,
    #[serde(rename = "legacyDrawingHF")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub legacy_drawing_h_f: Option<Box<LegacyDrawing>>,
    #[serde(rename = "drawingHF")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub drawing_h_f: Option<Box<DrawingHeaderFooter>>,
    #[serde(rename = "oleObjects")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub ole_objects: Option<Box<OleObjects>>,
    #[serde(rename = "controls")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub controls: Option<Box<Controls>>,
    #[serde(rename = "extLst")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub extension_list: Option<Box<ExtensionList>>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(rename = "worksheet")]
pub struct Worksheet {
    #[cfg(feature = "sml-styling")]
    #[serde(rename = "sheetPr")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub sheet_properties: Option<Box<SheetProperties>>,
    #[serde(rename = "dimension")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub dimension: Option<Box<SheetDimension>>,
    #[serde(rename = "sheetViews")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub sheet_views: Option<Box<SheetViews>>,
    #[cfg(feature = "sml-styling")]
    #[serde(rename = "sheetFormatPr")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub sheet_format: Option<Box<SheetFormat>>,
    #[cfg(feature = "sml-styling")]
    #[serde(rename = "cols")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub cols: Vec<Columns>,
    #[serde(rename = "sheetData")]
    pub sheet_data: Box<SheetData>,
    #[cfg(feature = "sml-formulas")]
    #[serde(rename = "sheetCalcPr")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub sheet_calc_pr: Option<Box<SheetCalcProperties>>,
    #[cfg(feature = "sml-protection")]
    #[serde(rename = "sheetProtection")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub sheet_protection: Option<Box<SheetProtection>>,
    #[cfg(feature = "sml-protection")]
    #[serde(rename = "protectedRanges")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub protected_ranges: Option<Box<ProtectedRanges>>,
    #[cfg(feature = "sml-formulas-advanced")]
    #[serde(rename = "scenarios")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub scenarios: Option<Box<Scenarios>>,
    #[cfg(feature = "sml-filtering")]
    #[serde(rename = "autoFilter")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub auto_filter: Option<Box<AutoFilter>>,
    #[cfg(feature = "sml-filtering")]
    #[serde(rename = "sortState")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub sort_state: Option<Box<SortState>>,
    #[cfg(feature = "sml-formulas-advanced")]
    #[serde(rename = "dataConsolidate")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub data_consolidate: Option<Box<CTDataConsolidate>>,
    #[cfg(feature = "sml-structure")]
    #[serde(rename = "customSheetViews")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub custom_sheet_views: Option<Box<CustomSheetViews>>,
    #[serde(rename = "mergeCells")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub merged_cells: Option<Box<MergedCells>>,
    #[cfg(feature = "sml-i18n")]
    #[serde(rename = "phoneticPr")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub phonetic_pr: Option<Box<PhoneticProperties>>,
    #[cfg(feature = "sml-styling")]
    #[serde(rename = "conditionalFormatting")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub conditional_formatting: Vec<ConditionalFormatting>,
    #[cfg(feature = "sml-validation")]
    #[serde(rename = "dataValidations")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub data_validations: Option<Box<DataValidations>>,
    #[cfg(feature = "sml-hyperlinks")]
    #[serde(rename = "hyperlinks")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub hyperlinks: Option<Box<Hyperlinks>>,
    #[cfg(feature = "sml-layout")]
    #[serde(rename = "printOptions")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub print_options: Option<Box<PrintOptions>>,
    #[cfg(feature = "sml-layout")]
    #[serde(rename = "pageMargins")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub page_margins: Option<Box<PageMargins>>,
    #[cfg(feature = "sml-layout")]
    #[serde(rename = "pageSetup")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub page_setup: Option<Box<PageSetup>>,
    #[cfg(feature = "sml-layout")]
    #[serde(rename = "headerFooter")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub header_footer: Option<Box<HeaderFooter>>,
    #[cfg(feature = "sml-layout")]
    #[serde(rename = "rowBreaks")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub row_breaks: Option<Box<PageBreaks>>,
    #[cfg(feature = "sml-layout")]
    #[serde(rename = "colBreaks")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub col_breaks: Option<Box<PageBreaks>>,
    #[cfg(feature = "sml-metadata")]
    #[serde(rename = "customProperties")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub custom_properties: Option<Box<CTCustomProperties>>,
    #[cfg(feature = "sml-formulas-advanced")]
    #[serde(rename = "cellWatches")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub cell_watches: Option<Box<CellWatches>>,
    #[cfg(feature = "sml-validation")]
    #[serde(rename = "ignoredErrors")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub ignored_errors: Option<Box<IgnoredErrors>>,
    #[cfg(feature = "sml-metadata")]
    #[serde(rename = "smartTags")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub smart_tags: Option<Box<SmartTags>>,
    #[cfg(feature = "sml-drawings")]
    #[serde(rename = "drawing")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub drawing: Option<Box<Drawing>>,
    #[cfg(feature = "sml-comments")]
    #[serde(rename = "legacyDrawing")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub legacy_drawing: Option<Box<LegacyDrawing>>,
    #[cfg(feature = "sml-layout")]
    #[serde(rename = "legacyDrawingHF")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub legacy_drawing_h_f: Option<Box<LegacyDrawing>>,
    #[cfg(feature = "sml-drawings")]
    #[serde(rename = "drawingHF")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub drawing_h_f: Option<Box<DrawingHeaderFooter>>,
    #[cfg(feature = "sml-drawings")]
    #[serde(rename = "picture")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub picture: Option<Box<SheetBackgroundPicture>>,
    #[cfg(feature = "sml-external")]
    #[serde(rename = "oleObjects")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub ole_objects: Option<Box<OleObjects>>,
    #[cfg(feature = "sml-external")]
    #[serde(rename = "controls")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub controls: Option<Box<Controls>>,
    #[cfg(feature = "sml-external")]
    #[serde(rename = "webPublishItems")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub web_publish_items: Option<Box<WebPublishItems>>,
    #[cfg(feature = "sml-tables")]
    #[serde(rename = "tableParts")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub table_parts: Option<Box<TableParts>>,
    #[cfg(feature = "sml-extensions")]
    #[serde(rename = "extLst")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub extension_list: Option<Box<ExtensionList>>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
#[serde(rename = "sheetData")]
pub struct SheetData {
    #[serde(rename = "row")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub row: Vec<Row>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct SheetCalcProperties {
    #[serde(rename = "@fullCalcOnLoad")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub full_calc_on_load: Option<bool>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(rename = "sheetFormatPr")]
pub struct SheetFormat {
    #[serde(rename = "@baseColWidth")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub base_col_width: Option<u32>,
    #[serde(rename = "@defaultColWidth")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub default_col_width: Option<f64>,
    #[serde(rename = "@defaultRowHeight")]
    pub default_row_height: f64,
    #[serde(rename = "@customHeight")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub custom_height: Option<bool>,
    #[serde(rename = "@zeroHeight")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub zero_height: Option<bool>,
    #[serde(rename = "@thickTop")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub thick_top: Option<bool>,
    #[serde(rename = "@thickBottom")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub thick_bottom: Option<bool>,
    #[serde(rename = "@outlineLevelRow")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub outline_level_row: Option<u8>,
    #[serde(rename = "@outlineLevelCol")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub outline_level_col: Option<u8>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
#[serde(rename = "cols")]
pub struct Columns {
    #[serde(rename = "col")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub col: Vec<Column>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(rename = "col")]
pub struct Column {
    #[cfg(feature = "sml-styling")]
    #[serde(rename = "@min")]
    pub start_column: u32,
    #[cfg(feature = "sml-styling")]
    #[serde(rename = "@max")]
    pub end_column: u32,
    #[cfg(feature = "sml-styling")]
    #[serde(rename = "@width")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub width: Option<f64>,
    #[cfg(feature = "sml-styling")]
    #[serde(rename = "@style")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub style: Option<u32>,
    #[cfg(feature = "sml-structure")]
    #[serde(rename = "@hidden")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub hidden: Option<bool>,
    #[cfg(feature = "sml-styling")]
    #[serde(rename = "@bestFit")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub best_fit: Option<bool>,
    #[cfg(feature = "sml-styling")]
    #[serde(rename = "@customWidth")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub custom_width: Option<bool>,
    #[cfg(feature = "sml-i18n")]
    #[serde(rename = "@phonetic")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub phonetic: Option<bool>,
    #[cfg(feature = "sml-structure")]
    #[serde(rename = "@outlineLevel")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub outline_level: Option<u8>,
    #[cfg(feature = "sml-structure")]
    #[serde(rename = "@collapsed")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub collapsed: Option<bool>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
#[serde(rename = "row")]
pub struct Row {
    #[serde(rename = "@r")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub reference: Option<u32>,
    #[serde(rename = "@spans")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub cell_spans: Option<CellSpans>,
    #[serde(rename = "@s")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub style_index: Option<u32>,
    #[cfg(feature = "sml-styling")]
    #[serde(rename = "@customFormat")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub custom_format: Option<bool>,
    #[cfg(feature = "sml-styling")]
    #[serde(rename = "@ht")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub height: Option<f64>,
    #[cfg(feature = "sml-structure")]
    #[serde(rename = "@hidden")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub hidden: Option<bool>,
    #[cfg(feature = "sml-styling")]
    #[serde(rename = "@customHeight")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub custom_height: Option<bool>,
    #[cfg(feature = "sml-structure")]
    #[serde(rename = "@outlineLevel")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub outline_level: Option<u8>,
    #[cfg(feature = "sml-structure")]
    #[serde(rename = "@collapsed")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub collapsed: Option<bool>,
    #[cfg(feature = "sml-styling")]
    #[serde(rename = "@thickTop")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub thick_top: Option<bool>,
    #[cfg(feature = "sml-styling")]
    #[serde(rename = "@thickBot")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub thick_bot: Option<bool>,
    #[cfg(feature = "sml-i18n")]
    #[serde(rename = "@ph")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub placeholder: Option<bool>,
    #[serde(rename = "c")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub cells: Vec<Cell>,
    #[cfg(feature = "sml-extensions")]
    #[serde(rename = "extLst")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub extension_list: Option<Box<ExtensionList>>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
#[serde(rename = "c")]
pub struct Cell {
    #[serde(rename = "@r")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub reference: Option<CellRef>,
    #[serde(rename = "@s")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub style_index: Option<u32>,
    #[serde(rename = "@t")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub cell_type: Option<CellType>,
    #[cfg(feature = "sml-metadata")]
    #[serde(rename = "@cm")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub cm: Option<u32>,
    #[cfg(feature = "sml-metadata")]
    #[serde(rename = "@vm")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub vm: Option<u32>,
    #[cfg(feature = "sml-i18n")]
    #[serde(rename = "@ph")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub placeholder: Option<bool>,
    #[serde(rename = "f")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub formula: Option<Box<CellFormula>>,
    #[serde(rename = "v")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub value: Option<XmlString>,
    #[serde(rename = "is")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub is: Option<Box<RichString>>,
    #[cfg(feature = "sml-extensions")]
    #[serde(rename = "extLst")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub extension_list: Option<Box<ExtensionList>>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
#[serde(rename = "sheetPr")]
pub struct SheetProperties {
    #[cfg(feature = "sml-structure")]
    #[serde(rename = "@syncHorizontal")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub sync_horizontal: Option<bool>,
    #[cfg(feature = "sml-structure")]
    #[serde(rename = "@syncVertical")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub sync_vertical: Option<bool>,
    #[cfg(feature = "sml-structure")]
    #[serde(rename = "@syncRef")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub sync_ref: Option<Reference>,
    #[cfg(feature = "sml-formulas")]
    #[serde(rename = "@transitionEvaluation")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub transition_evaluation: Option<bool>,
    #[cfg(feature = "sml-formulas")]
    #[serde(rename = "@transitionEntry")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub transition_entry: Option<bool>,
    #[cfg(feature = "sml-external")]
    #[serde(rename = "@published")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub published: Option<bool>,
    #[cfg(feature = "sml-structure")]
    #[serde(rename = "@codeName")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub code_name: Option<String>,
    #[cfg(feature = "sml-filtering")]
    #[serde(rename = "@filterMode")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub filter_mode: Option<bool>,
    #[cfg(feature = "sml-styling")]
    #[serde(rename = "@enableFormatConditionsCalculation")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub enable_format_conditions_calculation: Option<bool>,
    #[cfg(feature = "sml-styling")]
    #[serde(rename = "tabColor")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub tab_color: Option<Box<Color>>,
    #[cfg(feature = "sml-structure")]
    #[serde(rename = "outlinePr")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub outline_pr: Option<Box<OutlineProperties>>,
    #[cfg(feature = "sml-layout")]
    #[serde(rename = "pageSetUpPr")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub page_set_up_pr: Option<Box<PageSetupProperties>>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(rename = "dimension")]
pub struct SheetDimension {
    #[serde(rename = "@ref")]
    pub reference: Reference,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
#[serde(rename = "sheetViews")]
pub struct SheetViews {
    #[serde(rename = "sheetView")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub sheet_view: Vec<SheetView>,
    #[serde(rename = "extLst")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub extension_list: Option<Box<ExtensionList>>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(rename = "sheetView")]
pub struct SheetView {
    #[cfg(feature = "sml-protection")]
    #[serde(rename = "@windowProtection")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub window_protection: Option<bool>,
    #[cfg(feature = "sml-formulas")]
    #[serde(rename = "@showFormulas")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub show_formulas: Option<bool>,
    #[cfg(feature = "sml-styling")]
    #[serde(rename = "@showGridLines")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub show_grid_lines: Option<bool>,
    #[cfg(feature = "sml-styling")]
    #[serde(rename = "@showRowColHeaders")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub show_row_col_headers: Option<bool>,
    #[cfg(feature = "sml-styling")]
    #[serde(rename = "@showZeros")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub show_zeros: Option<bool>,
    #[cfg(feature = "sml-i18n")]
    #[serde(rename = "@rightToLeft")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub right_to_left: Option<bool>,
    #[serde(rename = "@tabSelected")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub tab_selected: Option<bool>,
    #[cfg(feature = "sml-layout")]
    #[serde(rename = "@showRuler")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub show_ruler: Option<bool>,
    #[cfg(feature = "sml-structure")]
    #[serde(rename = "@showOutlineSymbols")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub show_outline_symbols: Option<bool>,
    #[cfg(feature = "sml-styling")]
    #[serde(rename = "@defaultGridColor")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub default_grid_color: Option<bool>,
    #[cfg(feature = "sml-layout")]
    #[serde(rename = "@showWhiteSpace")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub show_white_space: Option<bool>,
    #[serde(rename = "@view")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub view: Option<SheetViewType>,
    #[serde(rename = "@topLeftCell")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub top_left_cell: Option<CellRef>,
    #[cfg(feature = "sml-styling")]
    #[serde(rename = "@colorId")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub color_id: Option<u32>,
    #[serde(rename = "@zoomScale")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub zoom_scale: Option<u32>,
    #[serde(rename = "@zoomScaleNormal")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub zoom_scale_normal: Option<u32>,
    #[cfg(feature = "sml-layout")]
    #[serde(rename = "@zoomScaleSheetLayoutView")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub zoom_scale_sheet_layout_view: Option<u32>,
    #[cfg(feature = "sml-layout")]
    #[serde(rename = "@zoomScalePageLayoutView")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub zoom_scale_page_layout_view: Option<u32>,
    #[serde(rename = "@workbookViewId")]
    pub workbook_view_id: u32,
    #[cfg(feature = "sml-structure")]
    #[serde(rename = "pane")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub pane: Option<Box<Pane>>,
    #[serde(rename = "selection")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub selection: Vec<Selection>,
    #[cfg(feature = "sml-pivot")]
    #[serde(rename = "pivotSelection")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub pivot_selection: Vec<CTPivotSelection>,
    #[cfg(feature = "sml-extensions")]
    #[serde(rename = "extLst")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub extension_list: Option<Box<ExtensionList>>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
#[serde(rename = "pane")]
pub struct Pane {
    #[cfg(feature = "sml-structure")]
    #[serde(rename = "@xSplit")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub x_split: Option<f64>,
    #[cfg(feature = "sml-structure")]
    #[serde(rename = "@ySplit")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub y_split: Option<f64>,
    #[cfg(feature = "sml-structure")]
    #[serde(rename = "@topLeftCell")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub top_left_cell: Option<CellRef>,
    #[cfg(feature = "sml-structure")]
    #[serde(rename = "@activePane")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub active_pane: Option<PaneType>,
    #[cfg(feature = "sml-structure")]
    #[serde(rename = "@state")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub state: Option<PaneState>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTPivotSelection {
    #[serde(rename = "@pane")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub pane: Option<PaneType>,
    #[serde(rename = "@showHeader")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub show_header: Option<bool>,
    #[serde(rename = "@label")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub label: Option<bool>,
    #[serde(rename = "@data")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub data: Option<bool>,
    #[serde(rename = "@extendable")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub extendable: Option<bool>,
    #[serde(rename = "@count")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub count: Option<u32>,
    #[serde(rename = "@axis")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub axis: Option<STAxis>,
    #[serde(rename = "@dimension")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub dimension: Option<u32>,
    #[serde(rename = "@start")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub start: Option<u32>,
    #[serde(rename = "@min")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub start_column: Option<u32>,
    #[serde(rename = "@max")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub end_column: Option<u32>,
    #[serde(rename = "@activeRow")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub active_row: Option<u32>,
    #[serde(rename = "@activeCol")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub active_col: Option<u32>,
    #[serde(rename = "@previousRow")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub previous_row: Option<u32>,
    #[serde(rename = "@previousCol")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub previous_col: Option<u32>,
    #[serde(rename = "@click")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub click: Option<u32>,
    #[serde(rename = "@r:id")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub id: Option<STRelationshipId>,
    #[serde(rename = "pivotArea")]
    pub pivot_area: Box<PivotArea>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
#[serde(rename = "selection")]
pub struct Selection {
    #[serde(rename = "@pane")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub pane: Option<PaneType>,
    #[serde(rename = "@activeCell")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub active_cell: Option<CellRef>,
    #[serde(rename = "@activeCellId")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub active_cell_id: Option<u32>,
    #[serde(rename = "@sqref")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub square_reference: Option<SquareRef>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
#[serde(rename = "rowBreaks")]
pub struct PageBreaks {
    #[cfg(feature = "sml-layout")]
    #[serde(rename = "@count")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub count: Option<u32>,
    #[cfg(feature = "sml-layout")]
    #[serde(rename = "@manualBreakCount")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub manual_break_count: Option<u32>,
    #[cfg(feature = "sml-layout")]
    #[serde(rename = "brk")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub brk: Vec<PageBreak>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct PageBreak {
    #[serde(rename = "@id")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub id: Option<u32>,
    #[serde(rename = "@min")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub start_column: Option<u32>,
    #[serde(rename = "@max")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub end_column: Option<u32>,
    #[serde(rename = "@man")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub man: Option<bool>,
    #[serde(rename = "@pt")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub pt: Option<bool>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct OutlineProperties {
    #[serde(rename = "@applyStyles")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub apply_styles: Option<bool>,
    #[serde(rename = "@summaryBelow")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub summary_below: Option<bool>,
    #[serde(rename = "@summaryRight")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub summary_right: Option<bool>,
    #[serde(rename = "@showOutlineSymbols")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub show_outline_symbols: Option<bool>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct PageSetupProperties {
    #[serde(rename = "@autoPageBreaks")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub auto_page_breaks: Option<bool>,
    #[serde(rename = "@fitToPage")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub fit_to_page: Option<bool>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct CTDataConsolidate {
    #[serde(rename = "@function")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub function: Option<STDataConsolidateFunction>,
    #[serde(rename = "@startLabels")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub start_labels: Option<bool>,
    #[serde(rename = "@leftLabels")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub left_labels: Option<bool>,
    #[serde(rename = "@topLabels")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub top_labels: Option<bool>,
    #[serde(rename = "@link")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub link: Option<bool>,
    #[serde(rename = "dataRefs")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub data_refs: Option<Box<CTDataRefs>>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct CTDataRefs {
    #[serde(rename = "@count")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub count: Option<u32>,
    #[serde(rename = "dataRef")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub data_ref: Vec<CTDataRef>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct CTDataRef {
    #[serde(rename = "@ref")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub reference: Option<Reference>,
    #[serde(rename = "@name")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub name: Option<XmlString>,
    #[serde(rename = "@sheet")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub sheet: Option<XmlString>,
    #[serde(rename = "@r:id")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub id: Option<STRelationshipId>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
#[serde(rename = "mergeCells")]
pub struct MergedCells {
    #[serde(rename = "@count")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub count: Option<u32>,
    #[serde(rename = "mergeCell")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub merge_cell: Vec<MergedCell>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(rename = "mergeCell")]
pub struct MergedCell {
    #[serde(rename = "@ref")]
    pub reference: Reference,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct SmartTags {
    #[serde(rename = "cellSmartTags")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub cell_smart_tags: Vec<CellSmartTags>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CellSmartTags {
    #[serde(rename = "@r")]
    pub reference: CellRef,
    #[serde(rename = "cellSmartTag")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub cell_smart_tag: Vec<CellSmartTag>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CellSmartTag {
    #[serde(rename = "@type")]
    pub r#type: u32,
    #[serde(rename = "@deleted")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub deleted: Option<bool>,
    #[serde(rename = "@xmlBased")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub xml_based: Option<bool>,
    #[serde(rename = "cellSmartTagPr")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub cell_smart_tag_pr: Vec<CTCellSmartTagPr>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTCellSmartTagPr {
    #[serde(rename = "@key")]
    pub key: XmlString,
    #[serde(rename = "@val")]
    pub value: XmlString,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct Drawing {
    #[serde(rename = "@r:id")]
    pub id: STRelationshipId,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct LegacyDrawing {
    #[serde(rename = "@r:id")]
    pub id: STRelationshipId,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct DrawingHeaderFooter {
    #[serde(rename = "@r:id")]
    pub id: STRelationshipId,
    #[serde(rename = "@lho")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub lho: Option<u32>,
    #[serde(rename = "@lhe")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub lhe: Option<u32>,
    #[serde(rename = "@lhf")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub lhf: Option<u32>,
    #[serde(rename = "@cho")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub cho: Option<u32>,
    #[serde(rename = "@che")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub che: Option<u32>,
    #[serde(rename = "@chf")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub chf: Option<u32>,
    #[serde(rename = "@rho")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub rho: Option<u32>,
    #[serde(rename = "@rhe")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub rhe: Option<u32>,
    #[serde(rename = "@rhf")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub rhf: Option<u32>,
    #[serde(rename = "@lfo")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub lfo: Option<u32>,
    #[serde(rename = "@lfe")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub lfe: Option<u32>,
    #[serde(rename = "@lff")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub lff: Option<u32>,
    #[serde(rename = "@cfo")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub cfo: Option<u32>,
    #[serde(rename = "@cfe")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub cfe: Option<u32>,
    #[serde(rename = "@cff")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub cff: Option<u32>,
    #[serde(rename = "@rfo")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub rfo: Option<u32>,
    #[serde(rename = "@rfe")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub rfe: Option<u32>,
    #[serde(rename = "@rff")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub rff: Option<u32>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct CustomSheetViews {
    #[serde(rename = "customSheetView")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub custom_sheet_view: Vec<CustomSheetView>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CustomSheetView {
    #[serde(rename = "@guid")]
    pub guid: Guid,
    #[serde(rename = "@scale")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub scale: Option<u32>,
    #[serde(rename = "@colorId")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub color_id: Option<u32>,
    #[serde(rename = "@showPageBreaks")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub show_page_breaks: Option<bool>,
    #[serde(rename = "@showFormulas")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub show_formulas: Option<bool>,
    #[serde(rename = "@showGridLines")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub show_grid_lines: Option<bool>,
    #[serde(rename = "@showRowCol")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub show_row_col: Option<bool>,
    #[serde(rename = "@outlineSymbols")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub outline_symbols: Option<bool>,
    #[serde(rename = "@zeroValues")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub zero_values: Option<bool>,
    #[serde(rename = "@fitToPage")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub fit_to_page: Option<bool>,
    #[serde(rename = "@printArea")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub print_area: Option<bool>,
    #[serde(rename = "@filter")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub filter: Option<bool>,
    #[serde(rename = "@showAutoFilter")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub show_auto_filter: Option<bool>,
    #[serde(rename = "@hiddenRows")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub hidden_rows: Option<bool>,
    #[serde(rename = "@hiddenColumns")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub hidden_columns: Option<bool>,
    #[serde(rename = "@state")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub state: Option<SheetState>,
    #[serde(rename = "@filterUnique")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub filter_unique: Option<bool>,
    #[serde(rename = "@view")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub view: Option<SheetViewType>,
    #[serde(rename = "@showRuler")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub show_ruler: Option<bool>,
    #[serde(rename = "@topLeftCell")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub top_left_cell: Option<CellRef>,
    #[serde(rename = "pane")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub pane: Option<Box<Pane>>,
    #[serde(rename = "selection")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub selection: Option<Box<Selection>>,
    #[serde(rename = "rowBreaks")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub row_breaks: Option<Box<PageBreaks>>,
    #[serde(rename = "colBreaks")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub col_breaks: Option<Box<PageBreaks>>,
    #[serde(rename = "pageMargins")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub page_margins: Option<Box<PageMargins>>,
    #[serde(rename = "printOptions")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub print_options: Option<Box<PrintOptions>>,
    #[serde(rename = "pageSetup")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub page_setup: Option<Box<PageSetup>>,
    #[serde(rename = "headerFooter")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub header_footer: Option<Box<HeaderFooter>>,
    #[serde(rename = "autoFilter")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub auto_filter: Option<Box<AutoFilter>>,
    #[serde(rename = "extLst")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub extension_list: Option<Box<ExtensionList>>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
#[serde(rename = "dataValidations")]
pub struct DataValidations {
    #[serde(rename = "@disablePrompts")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub disable_prompts: Option<bool>,
    #[serde(rename = "@xWindow")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub x_window: Option<u32>,
    #[serde(rename = "@yWindow")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub y_window: Option<u32>,
    #[serde(rename = "@count")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub count: Option<u32>,
    #[serde(rename = "dataValidation")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub data_validation: Vec<DataValidation>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(rename = "dataValidation")]
pub struct DataValidation {
    #[cfg(feature = "sml-validation")]
    #[serde(rename = "@type")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub r#type: Option<ValidationType>,
    #[cfg(feature = "sml-validation")]
    #[serde(rename = "@errorStyle")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub error_style: Option<ValidationErrorStyle>,
    #[cfg(feature = "sml-validation")]
    #[serde(rename = "@imeMode")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub ime_mode: Option<STDataValidationImeMode>,
    #[cfg(feature = "sml-validation")]
    #[serde(rename = "@operator")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub operator: Option<ValidationOperator>,
    #[cfg(feature = "sml-validation")]
    #[serde(rename = "@allowBlank")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub allow_blank: Option<bool>,
    #[cfg(feature = "sml-validation")]
    #[serde(rename = "@showDropDown")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub show_drop_down: Option<bool>,
    #[cfg(feature = "sml-validation")]
    #[serde(rename = "@showInputMessage")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub show_input_message: Option<bool>,
    #[cfg(feature = "sml-validation")]
    #[serde(rename = "@showErrorMessage")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub show_error_message: Option<bool>,
    #[cfg(feature = "sml-validation")]
    #[serde(rename = "@errorTitle")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub error_title: Option<XmlString>,
    #[cfg(feature = "sml-validation")]
    #[serde(rename = "@error")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub error: Option<XmlString>,
    #[cfg(feature = "sml-validation")]
    #[serde(rename = "@promptTitle")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub prompt_title: Option<XmlString>,
    #[cfg(feature = "sml-validation")]
    #[serde(rename = "@prompt")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub prompt: Option<XmlString>,
    #[cfg(feature = "sml-validation")]
    #[serde(rename = "@sqref")]
    pub square_reference: SquareRef,
    #[cfg(feature = "sml-validation")]
    #[serde(rename = "formula1")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub formula1: Option<STFormula>,
    #[cfg(feature = "sml-validation")]
    #[serde(rename = "formula2")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub formula2: Option<STFormula>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
#[serde(rename = "conditionalFormatting")]
pub struct ConditionalFormatting {
    #[cfg(feature = "sml-pivot")]
    #[serde(rename = "@pivot")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub pivot: Option<bool>,
    #[cfg(feature = "sml-styling")]
    #[serde(rename = "@sqref")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub square_reference: Option<SquareRef>,
    #[cfg(feature = "sml-styling")]
    #[serde(rename = "cfRule")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub cf_rule: Vec<ConditionalRule>,
    #[cfg(feature = "sml-extensions")]
    #[serde(rename = "extLst")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub extension_list: Option<Box<ExtensionList>>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(rename = "cfRule")]
pub struct ConditionalRule {
    #[cfg(feature = "sml-styling")]
    #[serde(rename = "@type")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub r#type: Option<ConditionalType>,
    #[cfg(feature = "sml-styling")]
    #[serde(rename = "@dxfId")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub dxf_id: Option<STDxfId>,
    #[cfg(feature = "sml-styling")]
    #[serde(rename = "@priority")]
    pub priority: i32,
    #[cfg(feature = "sml-styling")]
    #[serde(rename = "@stopIfTrue")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub stop_if_true: Option<bool>,
    #[cfg(feature = "sml-styling")]
    #[serde(rename = "@aboveAverage")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub above_average: Option<bool>,
    #[cfg(feature = "sml-styling")]
    #[serde(rename = "@percent")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub percent: Option<bool>,
    #[cfg(feature = "sml-styling")]
    #[serde(rename = "@bottom")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub bottom: Option<bool>,
    #[cfg(feature = "sml-styling")]
    #[serde(rename = "@operator")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub operator: Option<ConditionalOperator>,
    #[cfg(feature = "sml-styling")]
    #[serde(rename = "@text")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub text: Option<String>,
    #[cfg(feature = "sml-styling")]
    #[serde(rename = "@timePeriod")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub time_period: Option<STTimePeriod>,
    #[cfg(feature = "sml-styling")]
    #[serde(rename = "@rank")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub rank: Option<u32>,
    #[cfg(feature = "sml-styling")]
    #[serde(rename = "@stdDev")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub std_dev: Option<i32>,
    #[cfg(feature = "sml-styling")]
    #[serde(rename = "@equalAverage")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub equal_average: Option<bool>,
    #[cfg(feature = "sml-styling")]
    #[serde(rename = "formula")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub formula: Vec<STFormula>,
    #[cfg(feature = "sml-styling")]
    #[serde(rename = "colorScale")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub color_scale: Option<Box<ColorScale>>,
    #[cfg(feature = "sml-styling")]
    #[serde(rename = "dataBar")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub data_bar: Option<Box<DataBar>>,
    #[cfg(feature = "sml-styling")]
    #[serde(rename = "iconSet")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub icon_set: Option<Box<IconSet>>,
    #[cfg(feature = "sml-extensions")]
    #[serde(rename = "extLst")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub extension_list: Option<Box<ExtensionList>>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
#[serde(rename = "hyperlinks")]
pub struct Hyperlinks {
    #[serde(rename = "hyperlink")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub hyperlink: Vec<Hyperlink>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(rename = "hyperlink")]
pub struct Hyperlink {
    #[cfg(feature = "sml-hyperlinks")]
    #[serde(rename = "@ref")]
    pub reference: Reference,
    #[serde(rename = "@r:id")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub id: Option<STRelationshipId>,
    #[cfg(feature = "sml-hyperlinks")]
    #[serde(rename = "@location")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub location: Option<XmlString>,
    #[cfg(feature = "sml-hyperlinks")]
    #[serde(rename = "@tooltip")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub tooltip: Option<XmlString>,
    #[cfg(feature = "sml-hyperlinks")]
    #[serde(rename = "@display")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub display: Option<XmlString>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
#[serde(rename = "f")]
pub struct CellFormula {
    #[serde(rename = "$text")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub text: Option<String>,
    #[serde(rename = "@t")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub cell_type: Option<FormulaType>,
    #[cfg(feature = "sml-formulas-advanced")]
    #[serde(rename = "@aca")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub aca: Option<bool>,
    #[serde(rename = "@ref")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub reference: Option<Reference>,
    #[cfg(feature = "sml-formulas-advanced")]
    #[serde(rename = "@dt2D")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub dt2_d: Option<bool>,
    #[cfg(feature = "sml-formulas-advanced")]
    #[serde(rename = "@dtr")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub dtr: Option<bool>,
    #[cfg(feature = "sml-formulas-advanced")]
    #[serde(rename = "@del1")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub del1: Option<bool>,
    #[cfg(feature = "sml-formulas-advanced")]
    #[serde(rename = "@del2")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub del2: Option<bool>,
    #[cfg(feature = "sml-formulas-advanced")]
    #[serde(rename = "@r1")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub r1: Option<CellRef>,
    #[cfg(feature = "sml-formulas-advanced")]
    #[serde(rename = "@r2")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub r2: Option<CellRef>,
    #[cfg(feature = "sml-formulas-advanced")]
    #[serde(rename = "@ca")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub ca: Option<bool>,
    #[serde(rename = "@si")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub si: Option<u32>,
    #[cfg(feature = "sml-formulas-advanced")]
    #[serde(rename = "@bx")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub bx: Option<bool>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
#[serde(rename = "colorScale")]
pub struct ColorScale {
    #[cfg(feature = "sml-styling")]
    #[serde(rename = "cfvo")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub cfvo: Vec<ConditionalFormatValue>,
    #[cfg(feature = "sml-styling")]
    #[serde(rename = "color")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub color: Vec<Color>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(rename = "dataBar")]
pub struct DataBar {
    #[cfg(feature = "sml-styling")]
    #[serde(rename = "@minLength")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub min_length: Option<u32>,
    #[cfg(feature = "sml-styling")]
    #[serde(rename = "@maxLength")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub max_length: Option<u32>,
    #[cfg(feature = "sml-styling")]
    #[serde(rename = "@showValue")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub show_value: Option<bool>,
    #[cfg(feature = "sml-styling")]
    #[serde(rename = "cfvo")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub cfvo: Vec<ConditionalFormatValue>,
    #[cfg(feature = "sml-styling")]
    #[serde(rename = "color")]
    pub color: Box<Color>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
#[serde(rename = "iconSet")]
pub struct IconSet {
    #[cfg(feature = "sml-styling")]
    #[serde(rename = "@iconSet")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub icon_set: Option<IconSetType>,
    #[cfg(feature = "sml-styling")]
    #[serde(rename = "@showValue")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub show_value: Option<bool>,
    #[cfg(feature = "sml-styling")]
    #[serde(rename = "@percent")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub percent: Option<bool>,
    #[cfg(feature = "sml-styling")]
    #[serde(rename = "@reverse")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub reverse: Option<bool>,
    #[cfg(feature = "sml-styling")]
    #[serde(rename = "cfvo")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub cfvo: Vec<ConditionalFormatValue>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(rename = "cfvo")]
pub struct ConditionalFormatValue {
    #[cfg(feature = "sml-styling")]
    #[serde(rename = "@type")]
    pub r#type: ConditionalValueType,
    #[cfg(feature = "sml-styling")]
    #[serde(rename = "@val")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub value: Option<XmlString>,
    #[cfg(feature = "sml-styling")]
    #[serde(rename = "@gte")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub gte: Option<bool>,
    #[serde(rename = "extLst")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub extension_list: Option<Box<ExtensionList>>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(rename = "pageMargins")]
pub struct PageMargins {
    #[cfg(feature = "sml-layout")]
    #[serde(rename = "@left")]
    pub left: f64,
    #[cfg(feature = "sml-layout")]
    #[serde(rename = "@right")]
    pub right: f64,
    #[cfg(feature = "sml-layout")]
    #[serde(rename = "@top")]
    pub top: f64,
    #[cfg(feature = "sml-layout")]
    #[serde(rename = "@bottom")]
    pub bottom: f64,
    #[cfg(feature = "sml-layout")]
    #[serde(rename = "@header")]
    pub header: f64,
    #[cfg(feature = "sml-layout")]
    #[serde(rename = "@footer")]
    pub footer: f64,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
#[serde(rename = "printOptions")]
pub struct PrintOptions {
    #[serde(rename = "@horizontalCentered")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub horizontal_centered: Option<bool>,
    #[serde(rename = "@verticalCentered")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub vertical_centered: Option<bool>,
    #[serde(rename = "@headings")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub headings: Option<bool>,
    #[serde(rename = "@gridLines")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub grid_lines: Option<bool>,
    #[serde(rename = "@gridLinesSet")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub grid_lines_set: Option<bool>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
#[serde(rename = "pageSetup")]
pub struct PageSetup {
    #[cfg(feature = "sml-layout")]
    #[serde(rename = "@paperSize")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub paper_size: Option<u32>,
    #[cfg(feature = "sml-layout")]
    #[serde(rename = "@paperHeight")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub paper_height: Option<STPositiveUniversalMeasure>,
    #[cfg(feature = "sml-layout")]
    #[serde(rename = "@paperWidth")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub paper_width: Option<STPositiveUniversalMeasure>,
    #[cfg(feature = "sml-layout")]
    #[serde(rename = "@scale")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub scale: Option<u32>,
    #[cfg(feature = "sml-layout")]
    #[serde(rename = "@firstPageNumber")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub first_page_number: Option<u32>,
    #[cfg(feature = "sml-layout")]
    #[serde(rename = "@fitToWidth")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub fit_to_width: Option<u32>,
    #[cfg(feature = "sml-layout")]
    #[serde(rename = "@fitToHeight")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub fit_to_height: Option<u32>,
    #[cfg(feature = "sml-layout")]
    #[serde(rename = "@pageOrder")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub page_order: Option<STPageOrder>,
    #[cfg(feature = "sml-layout")]
    #[serde(rename = "@orientation")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub orientation: Option<STOrientation>,
    #[cfg(feature = "sml-layout")]
    #[serde(rename = "@usePrinterDefaults")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub use_printer_defaults: Option<bool>,
    #[cfg(feature = "sml-layout")]
    #[serde(rename = "@blackAndWhite")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub black_and_white: Option<bool>,
    #[cfg(feature = "sml-layout")]
    #[serde(rename = "@draft")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub draft: Option<bool>,
    #[cfg(feature = "sml-layout")]
    #[serde(rename = "@cellComments")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub cell_comments: Option<STCellComments>,
    #[cfg(feature = "sml-layout")]
    #[serde(rename = "@useFirstPageNumber")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub use_first_page_number: Option<bool>,
    #[cfg(feature = "sml-layout")]
    #[serde(rename = "@errors")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub errors: Option<STPrintError>,
    #[cfg(feature = "sml-layout")]
    #[serde(rename = "@horizontalDpi")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub horizontal_dpi: Option<u32>,
    #[cfg(feature = "sml-layout")]
    #[serde(rename = "@verticalDpi")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub vertical_dpi: Option<u32>,
    #[cfg(feature = "sml-layout")]
    #[serde(rename = "@copies")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub copies: Option<u32>,
    #[serde(rename = "@r:id")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub id: Option<STRelationshipId>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
#[serde(rename = "headerFooter")]
pub struct HeaderFooter {
    #[cfg(feature = "sml-layout")]
    #[serde(rename = "@differentOddEven")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub different_odd_even: Option<bool>,
    #[cfg(feature = "sml-layout")]
    #[serde(rename = "@differentFirst")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub different_first: Option<bool>,
    #[cfg(feature = "sml-layout")]
    #[serde(rename = "@scaleWithDoc")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub scale_with_doc: Option<bool>,
    #[cfg(feature = "sml-layout")]
    #[serde(rename = "@alignWithMargins")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub align_with_margins: Option<bool>,
    #[cfg(feature = "sml-layout")]
    #[serde(rename = "oddHeader")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub odd_header: Option<XmlString>,
    #[cfg(feature = "sml-layout")]
    #[serde(rename = "oddFooter")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub odd_footer: Option<XmlString>,
    #[cfg(feature = "sml-layout")]
    #[serde(rename = "evenHeader")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub even_header: Option<XmlString>,
    #[cfg(feature = "sml-layout")]
    #[serde(rename = "evenFooter")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub even_footer: Option<XmlString>,
    #[cfg(feature = "sml-layout")]
    #[serde(rename = "firstHeader")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub first_header: Option<XmlString>,
    #[cfg(feature = "sml-layout")]
    #[serde(rename = "firstFooter")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub first_footer: Option<XmlString>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct Scenarios {
    #[serde(rename = "@current")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub current: Option<u32>,
    #[serde(rename = "@show")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub show: Option<u32>,
    #[serde(rename = "@sqref")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub square_reference: Option<SquareRef>,
    #[serde(rename = "scenario")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub scenario: Vec<Scenario>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
#[serde(rename = "sheetProtection")]
pub struct SheetProtection {
    #[serde(rename = "@password")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub password: Option<STUnsignedShortHex>,
    #[serde(rename = "@algorithmName")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub algorithm_name: Option<XmlString>,
    #[serde(rename = "@hashValue")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub hash_value: Option<Vec<u8>>,
    #[serde(rename = "@saltValue")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub salt_value: Option<Vec<u8>>,
    #[serde(rename = "@spinCount")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub spin_count: Option<u32>,
    #[serde(rename = "@sheet")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub sheet: Option<bool>,
    #[serde(rename = "@objects")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub objects: Option<bool>,
    #[serde(rename = "@scenarios")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub scenarios: Option<bool>,
    #[serde(rename = "@formatCells")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub format_cells: Option<bool>,
    #[serde(rename = "@formatColumns")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub format_columns: Option<bool>,
    #[serde(rename = "@formatRows")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub format_rows: Option<bool>,
    #[serde(rename = "@insertColumns")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub insert_columns: Option<bool>,
    #[serde(rename = "@insertRows")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub insert_rows: Option<bool>,
    #[serde(rename = "@insertHyperlinks")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub insert_hyperlinks: Option<bool>,
    #[serde(rename = "@deleteColumns")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub delete_columns: Option<bool>,
    #[serde(rename = "@deleteRows")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub delete_rows: Option<bool>,
    #[serde(rename = "@selectLockedCells")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub select_locked_cells: Option<bool>,
    #[serde(rename = "@sort")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub sort: Option<bool>,
    #[serde(rename = "@autoFilter")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub auto_filter: Option<bool>,
    #[serde(rename = "@pivotTables")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub pivot_tables: Option<bool>,
    #[serde(rename = "@selectUnlockedCells")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub select_unlocked_cells: Option<bool>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
#[serde(rename = "protectedRanges")]
pub struct ProtectedRanges {
    #[serde(rename = "protectedRange")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub protected_range: Vec<ProtectedRange>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(rename = "protectedRange")]
pub struct ProtectedRange {
    #[serde(rename = "@password")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub password: Option<STUnsignedShortHex>,
    #[serde(rename = "@sqref")]
    pub square_reference: SquareRef,
    #[serde(rename = "@name")]
    pub name: XmlString,
    #[serde(rename = "@securityDescriptor")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub security_descriptor: Option<String>,
    #[serde(rename = "@algorithmName")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub algorithm_name: Option<XmlString>,
    #[serde(rename = "@hashValue")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub hash_value: Option<Vec<u8>>,
    #[serde(rename = "@saltValue")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub salt_value: Option<Vec<u8>>,
    #[serde(rename = "@spinCount")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub spin_count: Option<u32>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct Scenario {
    #[serde(rename = "@name")]
    pub name: XmlString,
    #[serde(rename = "@locked")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub locked: Option<bool>,
    #[serde(rename = "@hidden")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub hidden: Option<bool>,
    #[serde(rename = "@count")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub count: Option<u32>,
    #[serde(rename = "@user")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub user: Option<XmlString>,
    #[serde(rename = "@comment")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub comment: Option<XmlString>,
    #[serde(rename = "inputCells")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub input_cells: Vec<InputCells>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct InputCells {
    #[serde(rename = "@r")]
    pub reference: CellRef,
    #[serde(rename = "@deleted")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub deleted: Option<bool>,
    #[serde(rename = "@undone")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub undone: Option<bool>,
    #[serde(rename = "@val")]
    pub value: XmlString,
    #[serde(rename = "@numFmtId")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub number_format_id: Option<STNumFmtId>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct CellWatches {
    #[serde(rename = "cellWatch")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub cell_watch: Vec<CellWatch>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CellWatch {
    #[serde(rename = "@r")]
    pub reference: CellRef,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct Chartsheet {
    #[serde(rename = "sheetPr")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub sheet_properties: Option<Box<ChartsheetProperties>>,
    #[serde(rename = "sheetViews")]
    pub sheet_views: Box<ChartsheetViews>,
    #[serde(rename = "sheetProtection")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub sheet_protection: Option<Box<ChartsheetProtection>>,
    #[serde(rename = "customSheetViews")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub custom_sheet_views: Option<Box<CustomChartsheetViews>>,
    #[serde(rename = "pageMargins")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub page_margins: Option<Box<PageMargins>>,
    #[serde(rename = "pageSetup")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub page_setup: Option<Box<ChartsheetPageSetup>>,
    #[serde(rename = "headerFooter")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub header_footer: Option<Box<HeaderFooter>>,
    #[serde(rename = "drawing")]
    pub drawing: Box<Drawing>,
    #[serde(rename = "legacyDrawing")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub legacy_drawing: Option<Box<LegacyDrawing>>,
    #[serde(rename = "legacyDrawingHF")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub legacy_drawing_h_f: Option<Box<LegacyDrawing>>,
    #[serde(rename = "drawingHF")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub drawing_h_f: Option<Box<DrawingHeaderFooter>>,
    #[serde(rename = "picture")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub picture: Option<Box<SheetBackgroundPicture>>,
    #[serde(rename = "webPublishItems")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub web_publish_items: Option<Box<WebPublishItems>>,
    #[serde(rename = "extLst")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub extension_list: Option<Box<ExtensionList>>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct ChartsheetProperties {
    #[serde(rename = "@published")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub published: Option<bool>,
    #[serde(rename = "@codeName")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub code_name: Option<String>,
    #[serde(rename = "tabColor")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub tab_color: Option<Box<Color>>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct ChartsheetViews {
    #[serde(rename = "sheetView")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub sheet_view: Vec<ChartsheetView>,
    #[serde(rename = "extLst")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub extension_list: Option<Box<ExtensionList>>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct ChartsheetView {
    #[serde(rename = "@tabSelected")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub tab_selected: Option<bool>,
    #[serde(rename = "@zoomScale")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub zoom_scale: Option<u32>,
    #[serde(rename = "@workbookViewId")]
    pub workbook_view_id: u32,
    #[serde(rename = "@zoomToFit")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub zoom_to_fit: Option<bool>,
    #[serde(rename = "extLst")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub extension_list: Option<Box<ExtensionList>>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct ChartsheetProtection {
    #[serde(rename = "@password")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub password: Option<STUnsignedShortHex>,
    #[serde(rename = "@algorithmName")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub algorithm_name: Option<XmlString>,
    #[serde(rename = "@hashValue")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub hash_value: Option<Vec<u8>>,
    #[serde(rename = "@saltValue")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub salt_value: Option<Vec<u8>>,
    #[serde(rename = "@spinCount")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub spin_count: Option<u32>,
    #[serde(rename = "@content")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub content: Option<bool>,
    #[serde(rename = "@objects")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub objects: Option<bool>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct ChartsheetPageSetup {
    #[serde(rename = "@paperSize")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub paper_size: Option<u32>,
    #[serde(rename = "@paperHeight")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub paper_height: Option<STPositiveUniversalMeasure>,
    #[serde(rename = "@paperWidth")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub paper_width: Option<STPositiveUniversalMeasure>,
    #[serde(rename = "@firstPageNumber")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub first_page_number: Option<u32>,
    #[serde(rename = "@orientation")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub orientation: Option<STOrientation>,
    #[serde(rename = "@usePrinterDefaults")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub use_printer_defaults: Option<bool>,
    #[serde(rename = "@blackAndWhite")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub black_and_white: Option<bool>,
    #[serde(rename = "@draft")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub draft: Option<bool>,
    #[serde(rename = "@useFirstPageNumber")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub use_first_page_number: Option<bool>,
    #[serde(rename = "@horizontalDpi")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub horizontal_dpi: Option<u32>,
    #[serde(rename = "@verticalDpi")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub vertical_dpi: Option<u32>,
    #[serde(rename = "@copies")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub copies: Option<u32>,
    #[serde(rename = "@r:id")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub id: Option<STRelationshipId>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct CustomChartsheetViews {
    #[serde(rename = "customSheetView")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub custom_sheet_view: Vec<CustomChartsheetView>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CustomChartsheetView {
    #[serde(rename = "@guid")]
    pub guid: Guid,
    #[serde(rename = "@scale")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub scale: Option<u32>,
    #[serde(rename = "@state")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub state: Option<SheetState>,
    #[serde(rename = "@zoomToFit")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub zoom_to_fit: Option<bool>,
    #[serde(rename = "pageMargins")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub page_margins: Option<Box<PageMargins>>,
    #[serde(rename = "pageSetup")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub page_setup: Option<Box<ChartsheetPageSetup>>,
    #[serde(rename = "headerFooter")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub header_footer: Option<Box<HeaderFooter>>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct CTCustomProperties {
    #[serde(rename = "customPr")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub custom_pr: Vec<CTCustomProperty>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTCustomProperty {
    #[serde(rename = "@name")]
    pub name: XmlString,
    #[serde(rename = "@r:id")]
    pub id: STRelationshipId,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct OleObjects {
    #[serde(rename = "oleObject")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub ole_object: Vec<OleObject>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct OleObject {
    #[serde(rename = "@progId")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub prog_id: Option<String>,
    #[serde(rename = "@dvAspect")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub dv_aspect: Option<STDvAspect>,
    #[serde(rename = "@link")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub link: Option<XmlString>,
    #[serde(rename = "@oleUpdate")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub ole_update: Option<STOleUpdate>,
    #[serde(rename = "@autoLoad")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub auto_load: Option<bool>,
    #[serde(rename = "@shapeId")]
    pub shape_id: u32,
    #[serde(rename = "@r:id")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub id: Option<STRelationshipId>,
    #[serde(rename = "objectPr")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub object_pr: Option<Box<ObjectProperties>>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct ObjectProperties {
    #[serde(rename = "@locked")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub locked: Option<bool>,
    #[serde(rename = "@defaultSize")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub default_size: Option<bool>,
    #[serde(rename = "@print")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub print: Option<bool>,
    #[serde(rename = "@disabled")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub disabled: Option<bool>,
    #[serde(rename = "@uiObject")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub ui_object: Option<bool>,
    #[serde(rename = "@autoFill")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub auto_fill: Option<bool>,
    #[serde(rename = "@autoLine")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub auto_line: Option<bool>,
    #[serde(rename = "@autoPict")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub auto_pict: Option<bool>,
    #[serde(rename = "@macro")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub r#macro: Option<STFormula>,
    #[serde(rename = "@altText")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub alt_text: Option<XmlString>,
    #[serde(rename = "@dde")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub dde: Option<bool>,
    #[serde(rename = "@r:id")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub id: Option<STRelationshipId>,
    #[serde(rename = "anchor")]
    pub anchor: Box<ObjectAnchor>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct WebPublishItems {
    #[serde(rename = "@count")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub count: Option<u32>,
    #[serde(rename = "webPublishItem")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub web_publish_item: Vec<WebPublishItem>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct WebPublishItem {
    #[serde(rename = "@id")]
    pub id: u32,
    #[serde(rename = "@divId")]
    pub div_id: XmlString,
    #[serde(rename = "@sourceType")]
    pub source_type: STWebSourceType,
    #[serde(rename = "@sourceRef")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub source_ref: Option<Reference>,
    #[serde(rename = "@sourceObject")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub source_object: Option<XmlString>,
    #[serde(rename = "@destinationFile")]
    pub destination_file: XmlString,
    #[serde(rename = "@title")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub title: Option<XmlString>,
    #[serde(rename = "@autoRepublish")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub auto_republish: Option<bool>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct Controls {
    #[serde(rename = "control")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub control: Vec<Control>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct Control {
    #[serde(rename = "@shapeId")]
    pub shape_id: u32,
    #[serde(rename = "@r:id")]
    pub id: STRelationshipId,
    #[serde(rename = "@name")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub name: Option<String>,
    #[serde(rename = "controlPr")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub control_pr: Option<Box<CTControlPr>>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTControlPr {
    #[serde(rename = "@locked")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub locked: Option<bool>,
    #[serde(rename = "@defaultSize")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub default_size: Option<bool>,
    #[serde(rename = "@print")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub print: Option<bool>,
    #[serde(rename = "@disabled")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub disabled: Option<bool>,
    #[serde(rename = "@recalcAlways")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub recalc_always: Option<bool>,
    #[serde(rename = "@uiObject")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub ui_object: Option<bool>,
    #[serde(rename = "@autoFill")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub auto_fill: Option<bool>,
    #[serde(rename = "@autoLine")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub auto_line: Option<bool>,
    #[serde(rename = "@autoPict")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub auto_pict: Option<bool>,
    #[serde(rename = "@macro")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub r#macro: Option<STFormula>,
    #[serde(rename = "@altText")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub alt_text: Option<XmlString>,
    #[serde(rename = "@linkedCell")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub linked_cell: Option<STFormula>,
    #[serde(rename = "@listFillRange")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub list_fill_range: Option<STFormula>,
    #[serde(rename = "@cf")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub cf: Option<XmlString>,
    #[serde(rename = "@r:id")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub id: Option<STRelationshipId>,
    #[serde(rename = "anchor")]
    pub anchor: Box<ObjectAnchor>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct IgnoredErrors {
    #[serde(rename = "ignoredError")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub ignored_error: Vec<IgnoredError>,
    #[serde(rename = "extLst")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub extension_list: Option<Box<ExtensionList>>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct IgnoredError {
    #[serde(rename = "@sqref")]
    pub square_reference: SquareRef,
    #[serde(rename = "@evalError")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub eval_error: Option<bool>,
    #[serde(rename = "@twoDigitTextYear")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub two_digit_text_year: Option<bool>,
    #[serde(rename = "@numberStoredAsText")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub number_stored_as_text: Option<bool>,
    #[serde(rename = "@formula")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub formula: Option<bool>,
    #[serde(rename = "@formulaRange")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub formula_range: Option<bool>,
    #[serde(rename = "@unlockedFormula")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub unlocked_formula: Option<bool>,
    #[serde(rename = "@emptyCellReference")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub empty_cell_reference: Option<bool>,
    #[serde(rename = "@listDataValidation")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub list_data_validation: Option<bool>,
    #[serde(rename = "@calculatedColumn")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub calculated_column: Option<bool>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
#[serde(rename = "tableParts")]
pub struct TableParts {
    #[serde(rename = "@count")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub count: Option<u32>,
    #[serde(rename = "tablePart")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub table_part: Vec<TablePart>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(rename = "tablePart")]
pub struct TablePart {
    #[serde(rename = "@r:id")]
    pub id: STRelationshipId,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

pub type SmlMetadata = Box<Metadata>;

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct Metadata {
    #[serde(rename = "metadataTypes")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub metadata_types: Option<Box<MetadataTypes>>,
    #[serde(rename = "metadataStrings")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub metadata_strings: Option<Box<MetadataStrings>>,
    #[serde(rename = "mdxMetadata")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub mdx_metadata: Option<Box<CTMdxMetadata>>,
    #[serde(rename = "futureMetadata")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub future_metadata: Vec<CTFutureMetadata>,
    #[serde(rename = "cellMetadata")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub cell_metadata: Option<Box<MetadataBlocks>>,
    #[serde(rename = "valueMetadata")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub value_metadata: Option<Box<MetadataBlocks>>,
    #[serde(rename = "extLst")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub extension_list: Option<Box<ExtensionList>>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct MetadataTypes {
    #[serde(rename = "@count")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub count: Option<u32>,
    #[serde(rename = "metadataType")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub metadata_type: Vec<MetadataType>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct MetadataType {
    #[serde(rename = "@name")]
    pub name: XmlString,
    #[serde(rename = "@minSupportedVersion")]
    pub min_supported_version: u32,
    #[serde(rename = "@ghostRow")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub ghost_row: Option<bool>,
    #[serde(rename = "@ghostCol")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub ghost_col: Option<bool>,
    #[serde(rename = "@edit")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub edit: Option<bool>,
    #[serde(rename = "@delete")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub delete: Option<bool>,
    #[serde(rename = "@copy")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub copy: Option<bool>,
    #[serde(rename = "@pasteAll")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub paste_all: Option<bool>,
    #[serde(rename = "@pasteFormulas")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub paste_formulas: Option<bool>,
    #[serde(rename = "@pasteValues")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub paste_values: Option<bool>,
    #[serde(rename = "@pasteFormats")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub paste_formats: Option<bool>,
    #[serde(rename = "@pasteComments")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub paste_comments: Option<bool>,
    #[serde(rename = "@pasteDataValidation")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub paste_data_validation: Option<bool>,
    #[serde(rename = "@pasteBorders")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub paste_borders: Option<bool>,
    #[serde(rename = "@pasteColWidths")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub paste_col_widths: Option<bool>,
    #[serde(rename = "@pasteNumberFormats")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub paste_number_formats: Option<bool>,
    #[serde(rename = "@merge")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub merge: Option<bool>,
    #[serde(rename = "@splitFirst")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub split_first: Option<bool>,
    #[serde(rename = "@splitAll")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub split_all: Option<bool>,
    #[serde(rename = "@rowColShift")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub row_col_shift: Option<bool>,
    #[serde(rename = "@clearAll")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub clear_all: Option<bool>,
    #[serde(rename = "@clearFormats")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub clear_formats: Option<bool>,
    #[serde(rename = "@clearContents")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub clear_contents: Option<bool>,
    #[serde(rename = "@clearComments")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub clear_comments: Option<bool>,
    #[serde(rename = "@assign")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub assign: Option<bool>,
    #[serde(rename = "@coerce")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub coerce: Option<bool>,
    #[serde(rename = "@adjust")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub adjust: Option<bool>,
    #[serde(rename = "@cellMeta")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub cell_meta: Option<bool>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct MetadataBlocks {
    #[serde(rename = "@count")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub count: Option<u32>,
    #[serde(rename = "bk")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub bk: Vec<MetadataBlock>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct MetadataBlock {
    #[serde(rename = "rc")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub rc: Vec<MetadataRecord>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct MetadataRecord {
    #[serde(rename = "@t")]
    pub cell_type: u32,
    #[serde(rename = "@v")]
    pub value: u32,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTFutureMetadata {
    #[serde(rename = "@name")]
    pub name: XmlString,
    #[serde(rename = "@count")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub count: Option<u32>,
    #[serde(rename = "bk")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub bk: Vec<CTFutureMetadataBlock>,
    #[serde(rename = "extLst")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub extension_list: Option<Box<ExtensionList>>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct CTFutureMetadataBlock {
    #[serde(rename = "extLst")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub extension_list: Option<Box<ExtensionList>>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct CTMdxMetadata {
    #[serde(rename = "@count")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub count: Option<u32>,
    #[serde(rename = "mdx")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub mdx: Vec<CTMdx>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTMdx {
    #[serde(rename = "@n")]
    pub n: u32,
    #[serde(rename = "@f")]
    pub formula: STMdxFunctionType,
    #[serde(rename = "t")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub cell_type: Option<Box<CTMdxTuple>>,
    #[serde(rename = "ms")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub ms: Option<Box<CTMdxSet>>,
    #[serde(rename = "p")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub p: Option<Box<CTMdxMemeberProp>>,
    #[serde(rename = "k")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub k: Option<Box<CTMdxKPI>>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct CTMdxTuple {
    #[serde(rename = "@c")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub cells: Option<u32>,
    #[serde(rename = "@ct")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub ct: Option<XmlString>,
    #[serde(rename = "@si")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub si: Option<u32>,
    #[serde(rename = "@fi")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub fi: Option<u32>,
    #[serde(rename = "@bc")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub bc: Option<STUnsignedIntHex>,
    #[serde(rename = "@fc")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub fc: Option<STUnsignedIntHex>,
    #[serde(rename = "@i")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub i: Option<bool>,
    #[serde(rename = "@u")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub u: Option<bool>,
    #[serde(rename = "@st")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub st: Option<bool>,
    #[serde(rename = "@b")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub b: Option<bool>,
    #[serde(rename = "n")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub n: Vec<CTMetadataStringIndex>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTMdxSet {
    #[serde(rename = "@ns")]
    pub ns: u32,
    #[serde(rename = "@c")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub cells: Option<u32>,
    #[serde(rename = "@o")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub o: Option<STMdxSetOrder>,
    #[serde(rename = "n")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub n: Vec<CTMetadataStringIndex>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTMdxMemeberProp {
    #[serde(rename = "@n")]
    pub n: u32,
    #[serde(rename = "@np")]
    pub np: u32,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTMdxKPI {
    #[serde(rename = "@n")]
    pub n: u32,
    #[serde(rename = "@np")]
    pub np: u32,
    #[serde(rename = "@p")]
    pub p: STMdxKPIProperty,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTMetadataStringIndex {
    #[serde(rename = "@x")]
    pub x: u32,
    #[serde(rename = "@s")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub style_index: Option<bool>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct MetadataStrings {
    #[serde(rename = "@count")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub count: Option<u32>,
    #[serde(rename = "s")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub style_index: Vec<CTXStringElement>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

pub type SmlSingleXmlCells = Box<SingleXmlCells>;

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct SingleXmlCells {
    #[serde(rename = "singleXmlCell")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub single_xml_cell: Vec<SingleXmlCell>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct SingleXmlCell {
    #[serde(rename = "@id")]
    pub id: u32,
    #[serde(rename = "@r")]
    pub reference: CellRef,
    #[serde(rename = "@connectionId")]
    pub connection_id: u32,
    #[serde(rename = "xmlCellPr")]
    pub xml_cell_pr: Box<XmlCellProperties>,
    #[serde(rename = "extLst")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub extension_list: Option<Box<ExtensionList>>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct XmlCellProperties {
    #[serde(rename = "@id")]
    pub id: u32,
    #[serde(rename = "@uniqueName")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub unique_name: Option<XmlString>,
    #[serde(rename = "xmlPr")]
    pub xml_pr: Box<XmlProperties>,
    #[serde(rename = "extLst")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub extension_list: Option<Box<ExtensionList>>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct XmlProperties {
    #[serde(rename = "@mapId")]
    pub map_id: u32,
    #[serde(rename = "@xpath")]
    pub xpath: XmlString,
    #[serde(rename = "@xmlDataType")]
    pub xml_data_type: STXmlDataType,
    #[serde(rename = "extLst")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub extension_list: Option<Box<ExtensionList>>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

pub type SmlStyleSheet = Box<Stylesheet>;

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
#[serde(rename = "styleSheet")]
pub struct Stylesheet {
    #[cfg(feature = "sml-styling")]
    #[serde(rename = "numFmts")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub num_fmts: Option<Box<NumberFormats>>,
    #[cfg(feature = "sml-styling")]
    #[serde(rename = "fonts")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub fonts: Option<Box<Fonts>>,
    #[cfg(feature = "sml-styling")]
    #[serde(rename = "fills")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub fills: Option<Box<Fills>>,
    #[cfg(feature = "sml-styling")]
    #[serde(rename = "borders")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub borders: Option<Box<Borders>>,
    #[cfg(feature = "sml-styling")]
    #[serde(rename = "cellStyleXfs")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub cell_style_xfs: Option<Box<CellStyleFormats>>,
    #[cfg(feature = "sml-styling")]
    #[serde(rename = "cellXfs")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub cell_xfs: Option<Box<CellFormats>>,
    #[cfg(feature = "sml-styling")]
    #[serde(rename = "cellStyles")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub cell_styles: Option<Box<CellStyles>>,
    #[cfg(feature = "sml-styling")]
    #[serde(rename = "dxfs")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub dxfs: Option<Box<DifferentialFormats>>,
    #[cfg(feature = "sml-styling")]
    #[serde(rename = "tableStyles")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub table_styles: Option<Box<TableStyles>>,
    #[cfg(feature = "sml-styling")]
    #[serde(rename = "colors")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub colors: Option<Box<Colors>>,
    #[cfg(feature = "sml-extensions")]
    #[serde(rename = "extLst")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub extension_list: Option<Box<ExtensionList>>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
#[serde(rename = "alignment")]
pub struct CellAlignment {
    #[cfg(feature = "sml-styling")]
    #[serde(rename = "@horizontal")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub horizontal: Option<HorizontalAlignment>,
    #[cfg(feature = "sml-styling")]
    #[serde(rename = "@vertical")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub vertical: Option<VerticalAlignment>,
    #[cfg(feature = "sml-styling")]
    #[serde(rename = "@textRotation")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub text_rotation: Option<STTextRotation>,
    #[cfg(feature = "sml-styling")]
    #[serde(rename = "@wrapText")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub wrap_text: Option<bool>,
    #[cfg(feature = "sml-styling")]
    #[serde(rename = "@indent")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub indent: Option<u32>,
    #[cfg(feature = "sml-styling")]
    #[serde(rename = "@relativeIndent")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub relative_indent: Option<i32>,
    #[cfg(feature = "sml-styling")]
    #[serde(rename = "@justifyLastLine")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub justify_last_line: Option<bool>,
    #[cfg(feature = "sml-styling")]
    #[serde(rename = "@shrinkToFit")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub shrink_to_fit: Option<bool>,
    #[cfg(feature = "sml-styling")]
    #[serde(rename = "@readingOrder")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub reading_order: Option<u32>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
#[serde(rename = "borders")]
pub struct Borders {
    #[serde(rename = "@count")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub count: Option<u32>,
    #[serde(rename = "border")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub border: Vec<Border>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
#[serde(rename = "border")]
pub struct Border {
    #[cfg(feature = "sml-styling")]
    #[serde(rename = "@diagonalUp")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub diagonal_up: Option<bool>,
    #[cfg(feature = "sml-styling")]
    #[serde(rename = "@diagonalDown")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub diagonal_down: Option<bool>,
    #[cfg(feature = "sml-styling")]
    #[serde(rename = "@outline")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub outline: Option<bool>,
    #[serde(rename = "start")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub start: Option<Box<BorderProperties>>,
    #[serde(rename = "end")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub end: Option<Box<BorderProperties>>,
    #[cfg(feature = "sml-styling")]
    #[serde(rename = "left")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub left: Option<Box<BorderProperties>>,
    #[cfg(feature = "sml-styling")]
    #[serde(rename = "right")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub right: Option<Box<BorderProperties>>,
    #[cfg(feature = "sml-styling")]
    #[serde(rename = "top")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub top: Option<Box<BorderProperties>>,
    #[cfg(feature = "sml-styling")]
    #[serde(rename = "bottom")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub bottom: Option<Box<BorderProperties>>,
    #[cfg(feature = "sml-styling")]
    #[serde(rename = "diagonal")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub diagonal: Option<Box<BorderProperties>>,
    #[cfg(feature = "sml-styling")]
    #[serde(rename = "vertical")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub vertical: Option<Box<BorderProperties>>,
    #[cfg(feature = "sml-styling")]
    #[serde(rename = "horizontal")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub horizontal: Option<Box<BorderProperties>>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
#[serde(rename = "left")]
pub struct BorderProperties {
    #[serde(rename = "@style")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub style: Option<BorderStyle>,
    #[serde(rename = "color")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub color: Option<Box<Color>>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
#[serde(rename = "protection")]
pub struct CellProtection {
    #[cfg(feature = "sml-protection")]
    #[serde(rename = "@locked")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub locked: Option<bool>,
    #[cfg(feature = "sml-protection")]
    #[serde(rename = "@hidden")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub hidden: Option<bool>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
#[serde(rename = "fonts")]
pub struct Fonts {
    #[serde(rename = "@count")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub count: Option<u32>,
    #[serde(rename = "font")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub font: Vec<Font>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
#[serde(rename = "fills")]
pub struct Fills {
    #[serde(rename = "@count")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub count: Option<u32>,
    #[serde(rename = "fill")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub fill: Vec<Fill>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
#[serde(rename = "fill")]
pub struct Fill {
    #[cfg(feature = "sml-styling")]
    #[serde(rename = "patternFill")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub pattern_fill: Option<Box<PatternFill>>,
    #[cfg(feature = "sml-styling")]
    #[serde(rename = "gradientFill")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub gradient_fill: Option<Box<GradientFill>>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
#[serde(rename = "patternFill")]
pub struct PatternFill {
    #[serde(rename = "@patternType")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub pattern_type: Option<PatternType>,
    #[serde(rename = "fgColor")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub fg_color: Option<Box<Color>>,
    #[serde(rename = "bgColor")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub bg_color: Option<Box<Color>>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
#[serde(rename = "color")]
pub struct Color {
    #[cfg(feature = "sml-styling")]
    #[serde(rename = "@auto")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub auto: Option<bool>,
    #[cfg(feature = "sml-styling")]
    #[serde(rename = "@indexed")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub indexed: Option<u32>,
    #[cfg(feature = "sml-styling")]
    #[serde(rename = "@rgb")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub rgb: Option<STUnsignedIntHex>,
    #[cfg(feature = "sml-styling")]
    #[serde(rename = "@theme")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub theme: Option<u32>,
    #[cfg(feature = "sml-styling")]
    #[serde(rename = "@tint")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub tint: Option<f64>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
#[serde(rename = "gradientFill")]
pub struct GradientFill {
    #[serde(rename = "@type")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub r#type: Option<GradientType>,
    #[serde(rename = "@degree")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub degree: Option<f64>,
    #[serde(rename = "@left")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub left: Option<f64>,
    #[serde(rename = "@right")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub right: Option<f64>,
    #[serde(rename = "@top")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub top: Option<f64>,
    #[serde(rename = "@bottom")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub bottom: Option<f64>,
    #[serde(rename = "stop")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub stop: Vec<GradientStop>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(rename = "stop")]
pub struct GradientStop {
    #[serde(rename = "@position")]
    pub position: f64,
    #[serde(rename = "color")]
    pub color: Box<Color>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
#[serde(rename = "numFmts")]
pub struct NumberFormats {
    #[serde(rename = "@count")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub count: Option<u32>,
    #[serde(rename = "numFmt")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub num_fmt: Vec<NumberFormat>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(rename = "numFmt")]
pub struct NumberFormat {
    #[cfg(feature = "sml-styling")]
    #[serde(rename = "@numFmtId")]
    pub number_format_id: STNumFmtId,
    #[cfg(feature = "sml-styling")]
    #[serde(rename = "@formatCode")]
    pub format_code: XmlString,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
#[serde(rename = "cellStyleXfs")]
pub struct CellStyleFormats {
    #[serde(rename = "@count")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub count: Option<u32>,
    #[serde(rename = "xf")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub xf: Vec<Format>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
#[serde(rename = "cellXfs")]
pub struct CellFormats {
    #[serde(rename = "@count")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub count: Option<u32>,
    #[serde(rename = "xf")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub xf: Vec<Format>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
#[serde(rename = "xf")]
pub struct Format {
    #[cfg(feature = "sml-styling")]
    #[serde(rename = "@numFmtId")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub number_format_id: Option<STNumFmtId>,
    #[cfg(feature = "sml-styling")]
    #[serde(rename = "@fontId")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub font_id: Option<STFontId>,
    #[cfg(feature = "sml-styling")]
    #[serde(rename = "@fillId")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub fill_id: Option<STFillId>,
    #[cfg(feature = "sml-styling")]
    #[serde(rename = "@borderId")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub border_id: Option<STBorderId>,
    #[cfg(feature = "sml-styling")]
    #[serde(rename = "@xfId")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub format_id: Option<STCellStyleXfId>,
    #[cfg(feature = "sml-styling")]
    #[serde(rename = "@quotePrefix")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub quote_prefix: Option<bool>,
    #[cfg(feature = "sml-pivot")]
    #[serde(rename = "@pivotButton")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub pivot_button: Option<bool>,
    #[cfg(feature = "sml-styling")]
    #[serde(rename = "@applyNumberFormat")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub apply_number_format: Option<bool>,
    #[cfg(feature = "sml-styling")]
    #[serde(rename = "@applyFont")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub apply_font: Option<bool>,
    #[cfg(feature = "sml-styling")]
    #[serde(rename = "@applyFill")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub apply_fill: Option<bool>,
    #[cfg(feature = "sml-styling")]
    #[serde(rename = "@applyBorder")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub apply_border: Option<bool>,
    #[cfg(feature = "sml-styling")]
    #[serde(rename = "@applyAlignment")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub apply_alignment: Option<bool>,
    #[cfg(feature = "sml-styling")]
    #[serde(rename = "@applyProtection")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub apply_protection: Option<bool>,
    #[cfg(feature = "sml-styling")]
    #[serde(rename = "alignment")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub alignment: Option<Box<CellAlignment>>,
    #[cfg(feature = "sml-protection")]
    #[serde(rename = "protection")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub protection: Option<Box<CellProtection>>,
    #[cfg(feature = "sml-extensions")]
    #[serde(rename = "extLst")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub extension_list: Option<Box<ExtensionList>>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
#[serde(rename = "cellStyles")]
pub struct CellStyles {
    #[serde(rename = "@count")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub count: Option<u32>,
    #[serde(rename = "cellStyle")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub cell_style: Vec<CellStyle>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(rename = "cellStyle")]
pub struct CellStyle {
    #[serde(rename = "@name")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub name: Option<XmlString>,
    #[serde(rename = "@xfId")]
    pub format_id: STCellStyleXfId,
    #[serde(rename = "@builtinId")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub builtin_id: Option<u32>,
    #[serde(rename = "@iLevel")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub i_level: Option<u32>,
    #[serde(rename = "@hidden")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub hidden: Option<bool>,
    #[serde(rename = "@customBuiltin")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub custom_builtin: Option<bool>,
    #[serde(rename = "extLst")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub extension_list: Option<Box<ExtensionList>>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
#[serde(rename = "dxfs")]
pub struct DifferentialFormats {
    #[serde(rename = "@count")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub count: Option<u32>,
    #[serde(rename = "dxf")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub dxf: Vec<DifferentialFormat>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
#[serde(rename = "dxf")]
pub struct DifferentialFormat {
    #[serde(rename = "font")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub font: Option<Box<Font>>,
    #[serde(rename = "numFmt")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub num_fmt: Option<Box<NumberFormat>>,
    #[serde(rename = "fill")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub fill: Option<Box<Fill>>,
    #[serde(rename = "alignment")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub alignment: Option<Box<CellAlignment>>,
    #[serde(rename = "border")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub border: Option<Box<Border>>,
    #[serde(rename = "protection")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub protection: Option<Box<CellProtection>>,
    #[serde(rename = "extLst")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub extension_list: Option<Box<ExtensionList>>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
#[serde(rename = "colors")]
pub struct Colors {
    #[cfg(feature = "sml-styling")]
    #[serde(rename = "indexedColors")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub indexed_colors: Option<Box<IndexedColors>>,
    #[cfg(feature = "sml-styling")]
    #[serde(rename = "mruColors")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub mru_colors: Option<Box<MostRecentColors>>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct IndexedColors {
    #[serde(rename = "rgbColor")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub rgb_color: Vec<RgbColor>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct MostRecentColors {
    #[serde(rename = "color")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub color: Vec<Color>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct RgbColor {
    #[serde(rename = "@rgb")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub rgb: Option<STUnsignedIntHex>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct TableStyles {
    #[serde(rename = "@count")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub count: Option<u32>,
    #[serde(rename = "@defaultTableStyle")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub default_table_style: Option<String>,
    #[serde(rename = "@defaultPivotStyle")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub default_pivot_style: Option<String>,
    #[serde(rename = "tableStyle")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub table_style: Vec<TableStyle>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct TableStyle {
    #[serde(rename = "@name")]
    pub name: String,
    #[serde(rename = "@pivot")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub pivot: Option<bool>,
    #[serde(rename = "@table")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub table: Option<bool>,
    #[serde(rename = "@count")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub count: Option<u32>,
    #[serde(rename = "tableStyleElement")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub table_style_element: Vec<TableStyleElement>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct TableStyleElement {
    #[serde(rename = "@type")]
    pub r#type: STTableStyleType,
    #[serde(rename = "@size")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub size: Option<u32>,
    #[serde(rename = "@dxfId")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub dxf_id: Option<STDxfId>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct BooleanProperty {
    #[serde(rename = "@val")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub value: Option<bool>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct FontSize {
    #[serde(rename = "@val")]
    pub value: f64,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct IntProperty {
    #[serde(rename = "@val")]
    pub value: i32,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct FontName {
    #[serde(rename = "@val")]
    pub value: XmlString,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct VerticalAlignFontProperty {
    #[serde(rename = "@val")]
    pub value: VerticalAlignRun,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct FontSchemeProperty {
    #[serde(rename = "@val")]
    pub value: FontScheme,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct UnderlineProperty {
    #[serde(rename = "@val")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub value: Option<UnderlineStyle>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
#[serde(rename = "font")]
pub struct Font {
    #[cfg(feature = "sml-styling")]
    #[serde(rename = "name")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub name: Option<Box<FontName>>,
    #[cfg(feature = "sml-styling")]
    #[serde(rename = "charset")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub charset: Option<Box<IntProperty>>,
    #[cfg(feature = "sml-styling")]
    #[serde(rename = "family")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub family: Option<Box<FontFamily>>,
    #[cfg(feature = "sml-styling")]
    #[serde(rename = "b")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub b: Option<Box<BooleanProperty>>,
    #[cfg(feature = "sml-styling")]
    #[serde(rename = "i")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub i: Option<Box<BooleanProperty>>,
    #[cfg(feature = "sml-styling")]
    #[serde(rename = "strike")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub strike: Option<Box<BooleanProperty>>,
    #[cfg(feature = "sml-styling")]
    #[serde(rename = "outline")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub outline: Option<Box<BooleanProperty>>,
    #[cfg(feature = "sml-styling")]
    #[serde(rename = "shadow")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub shadow: Option<Box<BooleanProperty>>,
    #[cfg(feature = "sml-styling")]
    #[serde(rename = "condense")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub condense: Option<Box<BooleanProperty>>,
    #[cfg(feature = "sml-styling")]
    #[serde(rename = "extend")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub extend: Option<Box<BooleanProperty>>,
    #[cfg(feature = "sml-styling")]
    #[serde(rename = "color")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub color: Option<Box<Color>>,
    #[cfg(feature = "sml-styling")]
    #[serde(rename = "sz")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub sz: Option<Box<FontSize>>,
    #[cfg(feature = "sml-styling")]
    #[serde(rename = "u")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub u: Option<Box<UnderlineProperty>>,
    #[cfg(feature = "sml-styling")]
    #[serde(rename = "vertAlign")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub vert_align: Option<Box<VerticalAlignFontProperty>>,
    #[cfg(feature = "sml-styling")]
    #[serde(rename = "scheme")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub scheme: Option<Box<FontSchemeProperty>>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct FontFamily {
    #[serde(rename = "@val")]
    pub value: STFontFamily,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct SmlAGAutoFormat {
    #[serde(rename = "@autoFormatId")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub auto_format_id: Option<u32>,
    #[serde(rename = "@applyNumberFormats")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub apply_number_formats: Option<bool>,
    #[serde(rename = "@applyBorderFormats")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub apply_border_formats: Option<bool>,
    #[serde(rename = "@applyFontFormats")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub apply_font_formats: Option<bool>,
    #[serde(rename = "@applyPatternFormats")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub apply_pattern_formats: Option<bool>,
    #[serde(rename = "@applyAlignmentFormats")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub apply_alignment_formats: Option<bool>,
    #[serde(rename = "@applyWidthHeightFormats")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub apply_width_height_formats: Option<bool>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

pub type SmlExternalLink = Box<ExternalLink>;

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct ExternalLink {
    #[serde(rename = "externalBook")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub external_book: Option<Box<ExternalBook>>,
    #[serde(rename = "ddeLink")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub dde_link: Option<Box<DdeLink>>,
    #[serde(rename = "oleLink")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub ole_link: Option<Box<OleLink>>,
    #[serde(rename = "extLst")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub extension_list: Option<Box<ExtensionList>>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct ExternalBook {
    #[serde(rename = "@r:id")]
    pub id: STRelationshipId,
    #[serde(rename = "sheetNames")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub sheet_names: Option<Box<CTExternalSheetNames>>,
    #[serde(rename = "definedNames")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub defined_names: Option<Box<CTExternalDefinedNames>>,
    #[serde(rename = "sheetDataSet")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub sheet_data_set: Option<Box<ExternalSheetDataSet>>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct CTExternalSheetNames {
    #[serde(rename = "sheetName")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub sheet_name: Vec<CTExternalSheetName>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct CTExternalSheetName {
    #[serde(rename = "@val")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub value: Option<XmlString>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct CTExternalDefinedNames {
    #[serde(rename = "definedName")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub defined_name: Vec<CTExternalDefinedName>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTExternalDefinedName {
    #[serde(rename = "@name")]
    pub name: XmlString,
    #[serde(rename = "@refersTo")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub refers_to: Option<XmlString>,
    #[serde(rename = "@sheetId")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub sheet_id: Option<u32>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct ExternalSheetDataSet {
    #[serde(rename = "sheetData")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub sheet_data: Vec<ExternalSheetData>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct ExternalSheetData {
    #[serde(rename = "@sheetId")]
    pub sheet_id: u32,
    #[serde(rename = "@refreshError")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub refresh_error: Option<bool>,
    #[serde(rename = "row")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub row: Vec<ExternalRow>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct ExternalRow {
    #[serde(rename = "@r")]
    pub reference: u32,
    #[serde(rename = "cell")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub cell: Vec<ExternalCell>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct ExternalCell {
    #[serde(rename = "@r")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub reference: Option<CellRef>,
    #[serde(rename = "@t")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub cell_type: Option<CellType>,
    #[serde(rename = "@vm")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub vm: Option<u32>,
    #[serde(rename = "v")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub value: Option<XmlString>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct DdeLink {
    #[serde(rename = "@ddeService")]
    pub dde_service: XmlString,
    #[serde(rename = "@ddeTopic")]
    pub dde_topic: XmlString,
    #[serde(rename = "ddeItems")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub dde_items: Option<Box<DdeItems>>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct DdeItems {
    #[serde(rename = "ddeItem")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub dde_item: Vec<DdeItem>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct DdeItem {
    #[serde(rename = "@name")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub name: Option<XmlString>,
    #[serde(rename = "@ole")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub ole: Option<bool>,
    #[serde(rename = "@advise")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub advise: Option<bool>,
    #[serde(rename = "@preferPic")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub prefer_pic: Option<bool>,
    #[serde(rename = "values")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub values: Option<Box<CTDdeValues>>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct CTDdeValues {
    #[serde(rename = "@rows")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub rows: Option<u32>,
    #[serde(rename = "@cols")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub cols: Option<u32>,
    #[serde(rename = "value")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub value: Vec<CTDdeValue>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTDdeValue {
    #[serde(rename = "@t")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub cell_type: Option<STDdeValueType>,
    #[serde(rename = "val")]
    pub value: XmlString,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct OleLink {
    #[serde(rename = "@r:id")]
    pub id: STRelationshipId,
    #[serde(rename = "@progId")]
    pub prog_id: XmlString,
    #[serde(rename = "oleItems")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub ole_items: Option<Box<OleItems>>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct OleItems {
    #[serde(rename = "oleItem")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub ole_item: Vec<OleItem>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct OleItem {
    #[serde(rename = "@name")]
    pub name: XmlString,
    #[serde(rename = "@icon")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub icon: Option<bool>,
    #[serde(rename = "@advise")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub advise: Option<bool>,
    #[serde(rename = "@preferPic")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub prefer_pic: Option<bool>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

pub type SmlTable = Box<Table>;

#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(rename = "table")]
pub struct Table {
    #[cfg(feature = "sml-tables")]
    #[serde(rename = "@id")]
    pub id: u32,
    #[cfg(feature = "sml-tables")]
    #[serde(rename = "@name")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub name: Option<XmlString>,
    #[cfg(feature = "sml-tables")]
    #[serde(rename = "@displayName")]
    pub display_name: XmlString,
    #[cfg(feature = "sml-tables")]
    #[serde(rename = "@comment")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub comment: Option<XmlString>,
    #[cfg(feature = "sml-tables")]
    #[serde(rename = "@ref")]
    pub reference: Reference,
    #[cfg(feature = "sml-tables")]
    #[serde(rename = "@tableType")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub table_type: Option<STTableType>,
    #[cfg(feature = "sml-tables")]
    #[serde(rename = "@headerRowCount")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub header_row_count: Option<u32>,
    #[cfg(feature = "sml-tables")]
    #[serde(rename = "@insertRow")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub insert_row: Option<bool>,
    #[cfg(feature = "sml-tables")]
    #[serde(rename = "@insertRowShift")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub insert_row_shift: Option<bool>,
    #[cfg(feature = "sml-tables")]
    #[serde(rename = "@totalsRowCount")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub totals_row_count: Option<u32>,
    #[cfg(feature = "sml-tables")]
    #[serde(rename = "@totalsRowShown")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub totals_row_shown: Option<bool>,
    #[cfg(feature = "sml-tables")]
    #[serde(rename = "@published")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub published: Option<bool>,
    #[cfg(feature = "sml-tables")]
    #[serde(rename = "@headerRowDxfId")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub header_row_dxf_id: Option<STDxfId>,
    #[cfg(feature = "sml-tables")]
    #[serde(rename = "@dataDxfId")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub data_dxf_id: Option<STDxfId>,
    #[cfg(feature = "sml-tables")]
    #[serde(rename = "@totalsRowDxfId")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub totals_row_dxf_id: Option<STDxfId>,
    #[cfg(feature = "sml-tables")]
    #[serde(rename = "@headerRowBorderDxfId")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub header_row_border_dxf_id: Option<STDxfId>,
    #[cfg(feature = "sml-tables")]
    #[serde(rename = "@tableBorderDxfId")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub table_border_dxf_id: Option<STDxfId>,
    #[cfg(feature = "sml-tables")]
    #[serde(rename = "@totalsRowBorderDxfId")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub totals_row_border_dxf_id: Option<STDxfId>,
    #[cfg(feature = "sml-tables")]
    #[serde(rename = "@headerRowCellStyle")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub header_row_cell_style: Option<XmlString>,
    #[cfg(feature = "sml-tables")]
    #[serde(rename = "@dataCellStyle")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub data_cell_style: Option<XmlString>,
    #[cfg(feature = "sml-tables")]
    #[serde(rename = "@totalsRowCellStyle")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub totals_row_cell_style: Option<XmlString>,
    #[cfg(feature = "sml-tables")]
    #[serde(rename = "@connectionId")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub connection_id: Option<u32>,
    #[cfg(feature = "sml-tables")]
    #[serde(rename = "autoFilter")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub auto_filter: Option<Box<AutoFilter>>,
    #[cfg(feature = "sml-tables")]
    #[serde(rename = "sortState")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub sort_state: Option<Box<SortState>>,
    #[cfg(feature = "sml-tables")]
    #[serde(rename = "tableColumns")]
    pub table_columns: Box<TableColumns>,
    #[cfg(feature = "sml-tables")]
    #[serde(rename = "tableStyleInfo")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub table_style_info: Option<Box<TableStyleInfo>>,
    #[cfg(feature = "sml-extensions")]
    #[serde(rename = "extLst")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub extension_list: Option<Box<ExtensionList>>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct TableStyleInfo {
    #[serde(rename = "@name")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub name: Option<XmlString>,
    #[serde(rename = "@showFirstColumn")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub show_first_column: Option<bool>,
    #[serde(rename = "@showLastColumn")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub show_last_column: Option<bool>,
    #[serde(rename = "@showRowStripes")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub show_row_stripes: Option<bool>,
    #[serde(rename = "@showColumnStripes")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub show_column_stripes: Option<bool>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
#[serde(rename = "tableColumns")]
pub struct TableColumns {
    #[serde(rename = "@count")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub count: Option<u32>,
    #[serde(rename = "tableColumn")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub table_column: Vec<TableColumn>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(rename = "tableColumn")]
pub struct TableColumn {
    #[serde(rename = "@id")]
    pub id: u32,
    #[serde(rename = "@uniqueName")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub unique_name: Option<XmlString>,
    #[serde(rename = "@name")]
    pub name: XmlString,
    #[serde(rename = "@totalsRowFunction")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub totals_row_function: Option<STTotalsRowFunction>,
    #[serde(rename = "@totalsRowLabel")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub totals_row_label: Option<XmlString>,
    #[serde(rename = "@queryTableFieldId")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub query_table_field_id: Option<u32>,
    #[serde(rename = "@headerRowDxfId")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub header_row_dxf_id: Option<STDxfId>,
    #[serde(rename = "@dataDxfId")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub data_dxf_id: Option<STDxfId>,
    #[serde(rename = "@totalsRowDxfId")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub totals_row_dxf_id: Option<STDxfId>,
    #[serde(rename = "@headerRowCellStyle")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub header_row_cell_style: Option<XmlString>,
    #[serde(rename = "@dataCellStyle")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub data_cell_style: Option<XmlString>,
    #[serde(rename = "@totalsRowCellStyle")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub totals_row_cell_style: Option<XmlString>,
    #[serde(rename = "calculatedColumnFormula")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub calculated_column_formula: Option<Box<TableFormula>>,
    #[serde(rename = "totalsRowFormula")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub totals_row_formula: Option<Box<TableFormula>>,
    #[serde(rename = "xmlColumnPr")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub xml_column_pr: Option<Box<XmlColumnProperties>>,
    #[serde(rename = "extLst")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub extension_list: Option<Box<ExtensionList>>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct TableFormula {
    #[serde(rename = "$text")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub text: Option<String>,
    #[serde(rename = "@array")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub array: Option<bool>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct XmlColumnProperties {
    #[serde(rename = "@mapId")]
    pub map_id: u32,
    #[serde(rename = "@xpath")]
    pub xpath: XmlString,
    #[serde(rename = "@denormalized")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub denormalized: Option<bool>,
    #[serde(rename = "@xmlDataType")]
    pub xml_data_type: STXmlDataType,
    #[serde(rename = "extLst")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub extension_list: Option<Box<ExtensionList>>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

pub type SmlVolTypes = Box<CTVolTypes>;

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct CTVolTypes {
    #[serde(rename = "volType")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub vol_type: Vec<CTVolType>,
    #[serde(rename = "extLst")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub extension_list: Option<Box<ExtensionList>>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTVolType {
    #[serde(rename = "@type")]
    pub r#type: STVolDepType,
    #[serde(rename = "main")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub main: Vec<CTVolMain>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTVolMain {
    #[serde(rename = "@first")]
    pub first: XmlString,
    #[serde(rename = "tp")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub tp: Vec<CTVolTopic>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTVolTopic {
    #[serde(rename = "@t")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub cell_type: Option<STVolValueType>,
    #[serde(rename = "v")]
    pub value: XmlString,
    #[serde(rename = "stp")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub stp: Vec<XmlString>,
    #[serde(rename = "tr")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub tr: Vec<CTVolTopicRef>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTVolTopicRef {
    #[serde(rename = "@r")]
    pub reference: CellRef,
    #[serde(rename = "@s")]
    pub style_index: u32,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

pub type SmlWorkbook = Box<Workbook>;

#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(rename = "workbook")]
pub struct Workbook {
    #[serde(rename = "@conformance")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub conformance: Option<STConformanceClass>,
    #[serde(rename = "fileVersion")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub file_version: Option<Box<FileVersion>>,
    #[cfg(feature = "sml-protection")]
    #[serde(rename = "fileSharing")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub file_sharing: Option<Box<FileSharing>>,
    #[serde(rename = "workbookPr")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub workbook_pr: Option<Box<WorkbookProperties>>,
    #[cfg(feature = "sml-protection")]
    #[serde(rename = "workbookProtection")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub workbook_protection: Option<Box<WorkbookProtection>>,
    #[serde(rename = "bookViews")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub book_views: Option<Box<BookViews>>,
    #[serde(rename = "sheets")]
    pub sheets: Box<Sheets>,
    #[cfg(feature = "sml-formulas-advanced")]
    #[serde(rename = "functionGroups")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub function_groups: Option<Box<CTFunctionGroups>>,
    #[cfg(feature = "sml-external")]
    #[serde(rename = "externalReferences")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub external_references: Option<Box<ExternalReferences>>,
    #[serde(rename = "definedNames")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub defined_names: Option<Box<DefinedNames>>,
    #[cfg(feature = "sml-formulas")]
    #[serde(rename = "calcPr")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub calc_pr: Option<Box<CalculationProperties>>,
    #[cfg(feature = "sml-external")]
    #[serde(rename = "oleSize")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub ole_size: Option<Box<CTOleSize>>,
    #[cfg(feature = "sml-structure")]
    #[serde(rename = "customWorkbookViews")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub custom_workbook_views: Option<Box<CustomWorkbookViews>>,
    #[cfg(feature = "sml-pivot")]
    #[serde(rename = "pivotCaches")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub pivot_caches: Option<Box<PivotCaches>>,
    #[cfg(feature = "sml-metadata")]
    #[serde(rename = "smartTagPr")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub smart_tag_pr: Option<Box<CTSmartTagPr>>,
    #[cfg(feature = "sml-metadata")]
    #[serde(rename = "smartTagTypes")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub smart_tag_types: Option<Box<CTSmartTagTypes>>,
    #[cfg(feature = "sml-external")]
    #[serde(rename = "webPublishing")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub web_publishing: Option<Box<WebPublishing>>,
    #[serde(rename = "fileRecoveryPr")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub file_recovery_pr: Vec<FileRecoveryProperties>,
    #[cfg(feature = "sml-external")]
    #[serde(rename = "webPublishObjects")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub web_publish_objects: Option<Box<CTWebPublishObjects>>,
    #[cfg(feature = "sml-extensions")]
    #[serde(rename = "extLst")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub extension_list: Option<Box<ExtensionList>>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct FileVersion {
    #[serde(rename = "@appName")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub app_name: Option<String>,
    #[serde(rename = "@lastEdited")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub last_edited: Option<String>,
    #[serde(rename = "@lowestEdited")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub lowest_edited: Option<String>,
    #[serde(rename = "@rupBuild")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub rup_build: Option<String>,
    #[serde(rename = "@codeName")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub code_name: Option<Guid>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
#[serde(rename = "bookViews")]
pub struct BookViews {
    #[serde(rename = "workbookView")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub workbook_view: Vec<BookView>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
#[serde(rename = "workbookView")]
pub struct BookView {
    #[cfg(feature = "sml-structure")]
    #[serde(rename = "@visibility")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub visibility: Option<Visibility>,
    #[cfg(feature = "sml-structure")]
    #[serde(rename = "@minimized")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub minimized: Option<bool>,
    #[cfg(feature = "sml-structure")]
    #[serde(rename = "@showHorizontalScroll")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub show_horizontal_scroll: Option<bool>,
    #[cfg(feature = "sml-structure")]
    #[serde(rename = "@showVerticalScroll")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub show_vertical_scroll: Option<bool>,
    #[cfg(feature = "sml-structure")]
    #[serde(rename = "@showSheetTabs")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub show_sheet_tabs: Option<bool>,
    #[cfg(feature = "sml-structure")]
    #[serde(rename = "@xWindow")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub x_window: Option<i32>,
    #[cfg(feature = "sml-structure")]
    #[serde(rename = "@yWindow")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub y_window: Option<i32>,
    #[cfg(feature = "sml-structure")]
    #[serde(rename = "@windowWidth")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub window_width: Option<u32>,
    #[cfg(feature = "sml-structure")]
    #[serde(rename = "@windowHeight")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub window_height: Option<u32>,
    #[cfg(feature = "sml-structure")]
    #[serde(rename = "@tabRatio")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub tab_ratio: Option<u32>,
    #[cfg(feature = "sml-structure")]
    #[serde(rename = "@firstSheet")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub first_sheet: Option<u32>,
    #[serde(rename = "@activeTab")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub active_tab: Option<u32>,
    #[cfg(feature = "sml-filtering")]
    #[serde(rename = "@autoFilterDateGrouping")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub auto_filter_date_grouping: Option<bool>,
    #[cfg(feature = "sml-extensions")]
    #[serde(rename = "extLst")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub extension_list: Option<Box<ExtensionList>>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct CustomWorkbookViews {
    #[serde(rename = "customWorkbookView")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub custom_workbook_view: Vec<CustomWorkbookView>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CustomWorkbookView {
    #[serde(rename = "@name")]
    pub name: XmlString,
    #[serde(rename = "@guid")]
    pub guid: Guid,
    #[serde(rename = "@autoUpdate")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub auto_update: Option<bool>,
    #[serde(rename = "@mergeInterval")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub merge_interval: Option<u32>,
    #[serde(rename = "@changesSavedWin")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub changes_saved_win: Option<bool>,
    #[serde(rename = "@onlySync")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub only_sync: Option<bool>,
    #[serde(rename = "@personalView")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub personal_view: Option<bool>,
    #[serde(rename = "@includePrintSettings")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub include_print_settings: Option<bool>,
    #[serde(rename = "@includeHiddenRowCol")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub include_hidden_row_col: Option<bool>,
    #[serde(rename = "@maximized")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub maximized: Option<bool>,
    #[serde(rename = "@minimized")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub minimized: Option<bool>,
    #[serde(rename = "@showHorizontalScroll")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub show_horizontal_scroll: Option<bool>,
    #[serde(rename = "@showVerticalScroll")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub show_vertical_scroll: Option<bool>,
    #[serde(rename = "@showSheetTabs")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub show_sheet_tabs: Option<bool>,
    #[serde(rename = "@xWindow")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub x_window: Option<i32>,
    #[serde(rename = "@yWindow")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub y_window: Option<i32>,
    #[serde(rename = "@windowWidth")]
    pub window_width: u32,
    #[serde(rename = "@windowHeight")]
    pub window_height: u32,
    #[serde(rename = "@tabRatio")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub tab_ratio: Option<u32>,
    #[serde(rename = "@activeSheetId")]
    pub active_sheet_id: u32,
    #[serde(rename = "@showFormulaBar")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub show_formula_bar: Option<bool>,
    #[serde(rename = "@showStatusbar")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub show_statusbar: Option<bool>,
    #[serde(rename = "@showComments")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub show_comments: Option<CommentVisibility>,
    #[serde(rename = "@showObjects")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub show_objects: Option<ObjectVisibility>,
    #[serde(rename = "extLst")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub extension_list: Option<Box<ExtensionList>>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
#[serde(rename = "sheets")]
pub struct Sheets {
    #[serde(rename = "sheet")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub sheet: Vec<Sheet>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(rename = "sheet")]
pub struct Sheet {
    #[serde(rename = "@name")]
    pub name: XmlString,
    #[serde(rename = "@sheetId")]
    pub sheet_id: u32,
    #[cfg(feature = "sml-structure")]
    #[serde(rename = "@state")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub state: Option<SheetState>,
    #[serde(rename = "@r:id")]
    pub id: STRelationshipId,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
#[serde(rename = "workbookPr")]
pub struct WorkbookProperties {
    #[serde(rename = "@date1904")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub date1904: Option<bool>,
    #[serde(rename = "@showObjects")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub show_objects: Option<ObjectVisibility>,
    #[serde(rename = "@showBorderUnselectedTables")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub show_border_unselected_tables: Option<bool>,
    #[serde(rename = "@filterPrivacy")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub filter_privacy: Option<bool>,
    #[serde(rename = "@promptedSolutions")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub prompted_solutions: Option<bool>,
    #[serde(rename = "@showInkAnnotation")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub show_ink_annotation: Option<bool>,
    #[serde(rename = "@backupFile")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub backup_file: Option<bool>,
    #[serde(rename = "@saveExternalLinkValues")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub save_external_link_values: Option<bool>,
    #[serde(rename = "@updateLinks")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub update_links: Option<UpdateLinks>,
    #[serde(rename = "@codeName")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub code_name: Option<String>,
    #[serde(rename = "@hidePivotFieldList")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub hide_pivot_field_list: Option<bool>,
    #[serde(rename = "@showPivotChartFilter")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub show_pivot_chart_filter: Option<bool>,
    #[serde(rename = "@allowRefreshQuery")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub allow_refresh_query: Option<bool>,
    #[serde(rename = "@publishItems")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub publish_items: Option<bool>,
    #[serde(rename = "@checkCompatibility")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub check_compatibility: Option<bool>,
    #[serde(rename = "@autoCompressPictures")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub auto_compress_pictures: Option<bool>,
    #[serde(rename = "@refreshAllConnections")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub refresh_all_connections: Option<bool>,
    #[serde(rename = "@defaultThemeVersion")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub default_theme_version: Option<u32>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct CTSmartTagPr {
    #[serde(rename = "@embed")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub embed: Option<bool>,
    #[serde(rename = "@show")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub show: Option<STSmartTagShow>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct CTSmartTagTypes {
    #[serde(rename = "smartTagType")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub smart_tag_type: Vec<CTSmartTagType>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct CTSmartTagType {
    #[serde(rename = "@namespaceUri")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub namespace_uri: Option<XmlString>,
    #[serde(rename = "@name")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub name: Option<XmlString>,
    #[serde(rename = "@url")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub url: Option<XmlString>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct FileRecoveryProperties {
    #[serde(rename = "@autoRecover")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub auto_recover: Option<bool>,
    #[serde(rename = "@crashSave")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub crash_save: Option<bool>,
    #[serde(rename = "@dataExtractLoad")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub data_extract_load: Option<bool>,
    #[serde(rename = "@repairLoad")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub repair_load: Option<bool>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
#[serde(rename = "calcPr")]
pub struct CalculationProperties {
    #[cfg(feature = "sml-formulas")]
    #[serde(rename = "@calcId")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub calc_id: Option<u32>,
    #[cfg(feature = "sml-formulas")]
    #[serde(rename = "@calcMode")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub calc_mode: Option<CalculationMode>,
    #[cfg(feature = "sml-formulas")]
    #[serde(rename = "@fullCalcOnLoad")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub full_calc_on_load: Option<bool>,
    #[cfg(feature = "sml-formulas")]
    #[serde(rename = "@refMode")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub ref_mode: Option<ReferenceMode>,
    #[cfg(feature = "sml-formulas-advanced")]
    #[serde(rename = "@iterate")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub iterate: Option<bool>,
    #[cfg(feature = "sml-formulas-advanced")]
    #[serde(rename = "@iterateCount")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub iterate_count: Option<u32>,
    #[cfg(feature = "sml-formulas-advanced")]
    #[serde(rename = "@iterateDelta")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub iterate_delta: Option<f64>,
    #[cfg(feature = "sml-formulas")]
    #[serde(rename = "@fullPrecision")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub full_precision: Option<bool>,
    #[cfg(feature = "sml-formulas")]
    #[serde(rename = "@calcCompleted")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub calc_completed: Option<bool>,
    #[cfg(feature = "sml-formulas")]
    #[serde(rename = "@calcOnSave")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub calc_on_save: Option<bool>,
    #[cfg(feature = "sml-formulas-advanced")]
    #[serde(rename = "@concurrentCalc")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub concurrent_calc: Option<bool>,
    #[cfg(feature = "sml-formulas-advanced")]
    #[serde(rename = "@concurrentManualCount")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub concurrent_manual_count: Option<u32>,
    #[cfg(feature = "sml-formulas")]
    #[serde(rename = "@forceFullCalc")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub force_full_calc: Option<bool>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
#[serde(rename = "definedNames")]
pub struct DefinedNames {
    #[serde(rename = "definedName")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub defined_name: Vec<DefinedName>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(rename = "definedName")]
pub struct DefinedName {
    #[serde(rename = "$text")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub text: Option<String>,
    #[serde(rename = "@name")]
    pub name: XmlString,
    #[serde(rename = "@comment")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub comment: Option<XmlString>,
    #[cfg(feature = "sml-formulas-advanced")]
    #[serde(rename = "@customMenu")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub custom_menu: Option<XmlString>,
    #[serde(rename = "@description")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub description: Option<XmlString>,
    #[cfg(feature = "sml-formulas-advanced")]
    #[serde(rename = "@help")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub help: Option<XmlString>,
    #[cfg(feature = "sml-formulas-advanced")]
    #[serde(rename = "@statusBar")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub status_bar: Option<XmlString>,
    #[serde(rename = "@localSheetId")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub local_sheet_id: Option<u32>,
    #[cfg(feature = "sml-structure")]
    #[serde(rename = "@hidden")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub hidden: Option<bool>,
    #[cfg(feature = "sml-formulas-advanced")]
    #[serde(rename = "@function")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub function: Option<bool>,
    #[cfg(feature = "sml-formulas-advanced")]
    #[serde(rename = "@vbProcedure")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub vb_procedure: Option<bool>,
    #[cfg(feature = "sml-formulas-advanced")]
    #[serde(rename = "@xlm")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub xlm: Option<bool>,
    #[cfg(feature = "sml-formulas-advanced")]
    #[serde(rename = "@functionGroupId")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub function_group_id: Option<u32>,
    #[cfg(feature = "sml-formulas-advanced")]
    #[serde(rename = "@shortcutKey")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub shortcut_key: Option<XmlString>,
    #[cfg(feature = "sml-external")]
    #[serde(rename = "@publishToServer")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub publish_to_server: Option<bool>,
    #[cfg(feature = "sml-formulas-advanced")]
    #[serde(rename = "@workbookParameter")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub workbook_parameter: Option<bool>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct ExternalReferences {
    #[serde(rename = "externalReference")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub external_reference: Vec<ExternalReference>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct ExternalReference {
    #[serde(rename = "@r:id")]
    pub id: STRelationshipId,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct SheetBackgroundPicture {
    #[serde(rename = "@r:id")]
    pub id: STRelationshipId,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct PivotCaches {
    #[serde(rename = "pivotCache")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub pivot_cache: Vec<CTPivotCache>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTPivotCache {
    #[serde(rename = "@cacheId")]
    pub cache_id: u32,
    #[serde(rename = "@r:id")]
    pub id: STRelationshipId,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct FileSharing {
    #[serde(rename = "@readOnlyRecommended")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub read_only_recommended: Option<bool>,
    #[serde(rename = "@userName")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub user_name: Option<XmlString>,
    #[serde(rename = "@reservationPassword")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub reservation_password: Option<STUnsignedShortHex>,
    #[serde(rename = "@algorithmName")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub algorithm_name: Option<XmlString>,
    #[serde(rename = "@hashValue")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub hash_value: Option<Vec<u8>>,
    #[serde(rename = "@saltValue")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub salt_value: Option<Vec<u8>>,
    #[serde(rename = "@spinCount")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub spin_count: Option<u32>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTOleSize {
    #[serde(rename = "@ref")]
    pub reference: Reference,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
#[serde(rename = "workbookProtection")]
pub struct WorkbookProtection {
    #[serde(rename = "@workbookPassword")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub workbook_password: Option<STUnsignedShortHex>,
    #[serde(rename = "@workbookPasswordCharacterSet")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub workbook_password_character_set: Option<String>,
    #[serde(rename = "@revisionsPassword")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub revisions_password: Option<STUnsignedShortHex>,
    #[serde(rename = "@revisionsPasswordCharacterSet")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub revisions_password_character_set: Option<String>,
    #[serde(rename = "@lockStructure")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub lock_structure: Option<bool>,
    #[serde(rename = "@lockWindows")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub lock_windows: Option<bool>,
    #[serde(rename = "@lockRevision")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub lock_revision: Option<bool>,
    #[serde(rename = "@revisionsAlgorithmName")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub revisions_algorithm_name: Option<XmlString>,
    #[serde(rename = "@revisionsHashValue")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub revisions_hash_value: Option<Vec<u8>>,
    #[serde(rename = "@revisionsSaltValue")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub revisions_salt_value: Option<Vec<u8>>,
    #[serde(rename = "@revisionsSpinCount")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub revisions_spin_count: Option<u32>,
    #[serde(rename = "@workbookAlgorithmName")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub workbook_algorithm_name: Option<XmlString>,
    #[serde(rename = "@workbookHashValue")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub workbook_hash_value: Option<Vec<u8>>,
    #[serde(rename = "@workbookSaltValue")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub workbook_salt_value: Option<Vec<u8>>,
    #[serde(rename = "@workbookSpinCount")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub workbook_spin_count: Option<u32>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct WebPublishing {
    #[serde(rename = "@css")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub css: Option<bool>,
    #[serde(rename = "@thicket")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub thicket: Option<bool>,
    #[serde(rename = "@longFileNames")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub long_file_names: Option<bool>,
    #[serde(rename = "@vml")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub vml: Option<bool>,
    #[serde(rename = "@allowPng")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub allow_png: Option<bool>,
    #[serde(rename = "@targetScreenSize")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub target_screen_size: Option<STTargetScreenSize>,
    #[serde(rename = "@dpi")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub dpi: Option<u32>,
    #[serde(rename = "@codePage")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub code_page: Option<u32>,
    #[serde(rename = "@characterSet")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub character_set: Option<String>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct CTFunctionGroups {
    #[serde(rename = "@builtInGroupCount")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub built_in_group_count: Option<u32>,
    #[serde(rename = "functionGroup")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub function_group: Vec<CTFunctionGroup>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct CTFunctionGroup {
    #[serde(rename = "@name")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub name: Option<XmlString>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct CTWebPublishObjects {
    #[serde(rename = "@count")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub count: Option<u32>,
    #[serde(rename = "webPublishObject")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub web_publish_object: Vec<CTWebPublishObject>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTWebPublishObject {
    #[serde(rename = "@id")]
    pub id: u32,
    #[serde(rename = "@divId")]
    pub div_id: XmlString,
    #[serde(rename = "@sourceObject")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub source_object: Option<XmlString>,
    #[serde(rename = "@destinationFile")]
    pub destination_file: XmlString,
    #[serde(rename = "@title")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub title: Option<XmlString>,
    #[serde(rename = "@autoRepublish")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub auto_republish: Option<bool>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}
