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
    B,
    #[serde(rename = "n")]
    N,
    #[serde(rename = "e")]
    E,
    #[serde(rename = "s")]
    S,
    #[serde(rename = "str")]
    Str,
    #[serde(rename = "inlineStr")]
    InlineStr,
}

impl std::fmt::Display for CellType {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::B => write!(f, "b"),
            Self::N => write!(f, "n"),
            Self::E => write!(f, "e"),
            Self::S => write!(f, "s"),
            Self::Str => write!(f, "str"),
            Self::InlineStr => write!(f, "inlineStr"),
        }
    }
}

impl std::str::FromStr for CellType {
    type Err = String;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        match s {
            "b" => Ok(Self::B),
            "n" => Ok(Self::N),
            "e" => Ok(Self::E),
            "s" => Ok(Self::S),
            "str" => Ok(Self::Str),
            "inlineStr" => Ok(Self::InlineStr),
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
    S,
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
            Self::S => write!(f, "s"),
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
            "s" => Ok(Self::S),
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
    S,
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
            Self::S => write!(f, "s"),
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
            "s" => Ok(Self::S),
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
    B,
    #[serde(rename = "n")]
    N,
    #[serde(rename = "e")]
    E,
    #[serde(rename = "str")]
    Str,
}

impl std::fmt::Display for STDdeValueType {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::Nil => write!(f, "nil"),
            Self::B => write!(f, "b"),
            Self::N => write!(f, "n"),
            Self::E => write!(f, "e"),
            Self::Str => write!(f, "str"),
        }
    }
}

impl std::str::FromStr for STDdeValueType {
    type Err = String;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        match s {
            "nil" => Ok(Self::Nil),
            "b" => Ok(Self::B),
            "n" => Ok(Self::N),
            "e" => Ok(Self::E),
            "str" => Ok(Self::Str),
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
    B,
    #[serde(rename = "n")]
    N,
    #[serde(rename = "e")]
    E,
    #[serde(rename = "s")]
    S,
}

impl std::fmt::Display for STVolValueType {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::B => write!(f, "b"),
            Self::N => write!(f, "n"),
            Self::E => write!(f, "e"),
            Self::S => write!(f, "s"),
        }
    }
}

impl std::str::FromStr for STVolValueType {
    type Err = String;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        match s {
            "b" => Ok(Self::B),
            "n" => Ok(Self::N),
            "e" => Ok(Self::E),
            "s" => Ok(Self::S),
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

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct AutoFilter {
    #[serde(rename = "@ref")]
    #[serde(default)]
    pub reference: Option<Reference>,
    #[serde(rename = "filterColumn")]
    #[serde(default)]
    pub filter_column: Vec<Box<FilterColumn>>,
    #[serde(rename = "sortState")]
    #[serde(default)]
    pub sort_state: Option<Box<SortState>>,
    #[serde(rename = "extLst")]
    #[serde(default)]
    pub extension_list: Option<Box<ExtensionList>>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct FilterColumn {
    #[serde(rename = "@colId")]
    pub column_id: u32,
    #[serde(rename = "@hiddenButton")]
    #[serde(default)]
    pub hidden_button: Option<bool>,
    #[serde(rename = "@showButton")]
    #[serde(default)]
    pub show_button: Option<bool>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct Filters {
    #[serde(rename = "@blank")]
    #[serde(default)]
    pub blank: Option<bool>,
    #[serde(rename = "@calendarType")]
    #[serde(default)]
    pub calendar_type: Option<CalendarType>,
    #[serde(rename = "filter")]
    #[serde(default)]
    pub filter: Vec<Box<Filter>>,
    #[serde(rename = "dateGroupItem")]
    #[serde(default)]
    pub date_group_item: Vec<Box<DateGroupItem>>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct Filter {
    #[serde(rename = "@val")]
    #[serde(default)]
    pub value: Option<XmlString>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CustomFilters {
    #[serde(rename = "@and")]
    #[serde(default)]
    pub and: Option<bool>,
    #[serde(rename = "customFilter")]
    #[serde(default)]
    pub custom_filter: Vec<Box<CustomFilter>>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CustomFilter {
    #[serde(rename = "@operator")]
    #[serde(default)]
    pub operator: Option<FilterOperator>,
    #[serde(rename = "@val")]
    #[serde(default)]
    pub value: Option<XmlString>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct Top10Filter {
    #[serde(rename = "@top")]
    #[serde(default)]
    pub top: Option<bool>,
    #[serde(rename = "@percent")]
    #[serde(default)]
    pub percent: Option<bool>,
    #[serde(rename = "@val")]
    pub value: f64,
    #[serde(rename = "@filterVal")]
    #[serde(default)]
    pub filter_val: Option<f64>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct ColorFilter {
    #[serde(rename = "@dxfId")]
    #[serde(default)]
    pub dxf_id: Option<STDxfId>,
    #[serde(rename = "@cellColor")]
    #[serde(default)]
    pub cell_color: Option<bool>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct IconFilter {
    #[serde(rename = "@iconSet")]
    pub icon_set: IconSetType,
    #[serde(rename = "@iconId")]
    #[serde(default)]
    pub icon_id: Option<u32>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct DynamicFilter {
    #[serde(rename = "@type")]
    pub r#type: DynamicFilterType,
    #[serde(rename = "@val")]
    #[serde(default)]
    pub value: Option<f64>,
    #[serde(rename = "@valIso")]
    #[serde(default)]
    pub val_iso: Option<String>,
    #[serde(rename = "@maxVal")]
    #[serde(default)]
    pub max_val: Option<f64>,
    #[serde(rename = "@maxValIso")]
    #[serde(default)]
    pub max_val_iso: Option<String>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct SortState {
    #[serde(rename = "@columnSort")]
    #[serde(default)]
    pub column_sort: Option<bool>,
    #[serde(rename = "@caseSensitive")]
    #[serde(default)]
    pub case_sensitive: Option<bool>,
    #[serde(rename = "@sortMethod")]
    #[serde(default)]
    pub sort_method: Option<SortMethod>,
    #[serde(rename = "@ref")]
    pub reference: Reference,
    #[serde(rename = "sortCondition")]
    #[serde(default)]
    pub sort_condition: Vec<Box<SortCondition>>,
    #[serde(rename = "extLst")]
    #[serde(default)]
    pub extension_list: Option<Box<ExtensionList>>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct SortCondition {
    #[serde(rename = "@descending")]
    #[serde(default)]
    pub descending: Option<bool>,
    #[serde(rename = "@sortBy")]
    #[serde(default)]
    pub sort_by: Option<SortBy>,
    #[serde(rename = "@ref")]
    pub reference: Reference,
    #[serde(rename = "@customList")]
    #[serde(default)]
    pub custom_list: Option<XmlString>,
    #[serde(rename = "@dxfId")]
    #[serde(default)]
    pub dxf_id: Option<STDxfId>,
    #[serde(rename = "@iconSet")]
    #[serde(default)]
    pub icon_set: Option<IconSetType>,
    #[serde(rename = "@iconId")]
    #[serde(default)]
    pub icon_id: Option<u32>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct DateGroupItem {
    #[serde(rename = "@year")]
    pub year: u16,
    #[serde(rename = "@month")]
    #[serde(default)]
    pub month: Option<u16>,
    #[serde(rename = "@day")]
    #[serde(default)]
    pub day: Option<u16>,
    #[serde(rename = "@hour")]
    #[serde(default)]
    pub hour: Option<u16>,
    #[serde(rename = "@minute")]
    #[serde(default)]
    pub minute: Option<u16>,
    #[serde(rename = "@second")]
    #[serde(default)]
    pub second: Option<u16>,
    #[serde(rename = "@dateTimeGrouping")]
    pub date_time_grouping: STDateTimeGrouping,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTXStringElement {
    #[serde(rename = "@v")]
    pub value: XmlString,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct Extension {
    #[serde(rename = "@uri")]
    #[serde(default)]
    pub uri: Option<String>,
}

pub type CTExtensionAny = String;

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct ObjectAnchor {
    #[serde(rename = "@moveWithCells")]
    #[serde(default)]
    pub move_with_cells: Option<bool>,
    #[serde(rename = "@sizeWithCells")]
    #[serde(default)]
    pub size_with_cells: Option<bool>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct EGExtensionList {
    #[serde(rename = "ext")]
    #[serde(default)]
    pub ext: Vec<Box<Extension>>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct ExtensionList;

pub type SmlCalcChain = Box<CalcChain>;

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CalcChain {
    #[serde(rename = "c")]
    #[serde(default)]
    pub cells: Vec<Box<CalcCell>>,
    #[serde(rename = "extLst")]
    #[serde(default)]
    pub extension_list: Option<Box<ExtensionList>>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CalcCell {
    #[serde(rename = "@_any")]
    pub _any: CellRef,
    #[serde(rename = "@i")]
    #[serde(default)]
    pub i: Option<i32>,
    #[serde(rename = "@s")]
    #[serde(default)]
    pub style_index: Option<bool>,
    #[serde(rename = "@l")]
    #[serde(default)]
    pub l: Option<bool>,
    #[serde(rename = "@t")]
    #[serde(default)]
    pub cell_type: Option<bool>,
    #[serde(rename = "@a")]
    #[serde(default)]
    pub a: Option<bool>,
}

pub type SmlComments = Box<Comments>;

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct Comments {
    #[serde(rename = "authors")]
    pub authors: Box<Authors>,
    #[serde(rename = "commentList")]
    pub comment_list: Box<CommentList>,
    #[serde(rename = "extLst")]
    #[serde(default)]
    pub extension_list: Option<Box<ExtensionList>>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct Authors {
    #[serde(rename = "author")]
    #[serde(default)]
    pub author: Vec<XmlString>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CommentList {
    #[serde(rename = "comment")]
    #[serde(default)]
    pub comment: Vec<Box<Comment>>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct Comment {
    #[serde(rename = "@ref")]
    pub reference: Reference,
    #[serde(rename = "@authorId")]
    pub author_id: u32,
    #[serde(rename = "@guid")]
    #[serde(default)]
    pub guid: Option<Guid>,
    #[serde(rename = "@shapeId")]
    #[serde(default)]
    pub shape_id: Option<u32>,
    #[serde(rename = "text")]
    pub text: Box<RichString>,
    #[serde(rename = "commentPr")]
    #[serde(default)]
    pub comment_pr: Option<Box<CTCommentPr>>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTCommentPr {
    #[serde(rename = "@locked")]
    #[serde(default)]
    pub locked: Option<bool>,
    #[serde(rename = "@defaultSize")]
    #[serde(default)]
    pub default_size: Option<bool>,
    #[serde(rename = "@print")]
    #[serde(default)]
    pub print: Option<bool>,
    #[serde(rename = "@disabled")]
    #[serde(default)]
    pub disabled: Option<bool>,
    #[serde(rename = "@autoFill")]
    #[serde(default)]
    pub auto_fill: Option<bool>,
    #[serde(rename = "@autoLine")]
    #[serde(default)]
    pub auto_line: Option<bool>,
    #[serde(rename = "@altText")]
    #[serde(default)]
    pub alt_text: Option<XmlString>,
    #[serde(rename = "@textHAlign")]
    #[serde(default)]
    pub text_h_align: Option<STTextHAlign>,
    #[serde(rename = "@textVAlign")]
    #[serde(default)]
    pub text_v_align: Option<STTextVAlign>,
    #[serde(rename = "@lockText")]
    #[serde(default)]
    pub lock_text: Option<bool>,
    #[serde(rename = "@justLastX")]
    #[serde(default)]
    pub just_last_x: Option<bool>,
    #[serde(rename = "@autoScale")]
    #[serde(default)]
    pub auto_scale: Option<bool>,
    #[serde(rename = "anchor")]
    pub anchor: Box<ObjectAnchor>,
}

pub type SmlMapInfo = Box<MapInfo>;

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct MapInfo {
    #[serde(rename = "@SelectionNamespaces")]
    pub selection_namespaces: String,
    #[serde(rename = "Schema")]
    #[serde(default)]
    pub schema: Vec<Box<XmlSchema>>,
    #[serde(rename = "Map")]
    #[serde(default)]
    pub map: Vec<Box<XmlMap>>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct XmlSchema;

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
    pub show_import_export_validation_errors: bool,
    #[serde(rename = "@AutoFit")]
    pub auto_fit: bool,
    #[serde(rename = "@Append")]
    pub append: bool,
    #[serde(rename = "@PreserveSortAFLayout")]
    pub preserve_sort_a_f_layout: bool,
    #[serde(rename = "@PreserveFormat")]
    pub preserve_format: bool,
    #[serde(rename = "DataBinding")]
    #[serde(default)]
    pub data_binding: Option<Box<DataBinding>>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct DataBinding {
    #[serde(rename = "@DataBindingName")]
    #[serde(default)]
    pub data_binding_name: Option<String>,
    #[serde(rename = "@FileBinding")]
    #[serde(default)]
    pub file_binding: Option<bool>,
    #[serde(rename = "@ConnectionID")]
    #[serde(default)]
    pub connection_i_d: Option<u32>,
    #[serde(rename = "@FileBindingName")]
    #[serde(default)]
    pub file_binding_name: Option<String>,
    #[serde(rename = "@DataBindingLoadMode")]
    pub data_binding_load_mode: u32,
}

pub type CTDataBindingAny = String;

pub type SmlConnections = Box<Connections>;

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct Connections {
    #[serde(rename = "connection")]
    #[serde(default)]
    pub connection: Vec<Box<Connection>>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct Connection {
    #[serde(rename = "@id")]
    pub id: u32,
    #[serde(rename = "@sourceFile")]
    #[serde(default)]
    pub source_file: Option<XmlString>,
    #[serde(rename = "@odcFile")]
    #[serde(default)]
    pub odc_file: Option<XmlString>,
    #[serde(rename = "@keepAlive")]
    #[serde(default)]
    pub keep_alive: Option<bool>,
    #[serde(rename = "@interval")]
    #[serde(default)]
    pub interval: Option<u32>,
    #[serde(rename = "@name")]
    #[serde(default)]
    pub name: Option<XmlString>,
    #[serde(rename = "@description")]
    #[serde(default)]
    pub description: Option<XmlString>,
    #[serde(rename = "@type")]
    #[serde(default)]
    pub r#type: Option<u32>,
    #[serde(rename = "@reconnectionMethod")]
    #[serde(default)]
    pub reconnection_method: Option<u32>,
    #[serde(rename = "@refreshedVersion")]
    pub refreshed_version: u8,
    #[serde(rename = "@minRefreshableVersion")]
    #[serde(default)]
    pub min_refreshable_version: Option<u8>,
    #[serde(rename = "@savePassword")]
    #[serde(default)]
    pub save_password: Option<bool>,
    #[serde(rename = "@new")]
    #[serde(default)]
    pub new: Option<bool>,
    #[serde(rename = "@deleted")]
    #[serde(default)]
    pub deleted: Option<bool>,
    #[serde(rename = "@onlyUseConnectionFile")]
    #[serde(default)]
    pub only_use_connection_file: Option<bool>,
    #[serde(rename = "@background")]
    #[serde(default)]
    pub background: Option<bool>,
    #[serde(rename = "@refreshOnLoad")]
    #[serde(default)]
    pub refresh_on_load: Option<bool>,
    #[serde(rename = "@saveData")]
    #[serde(default)]
    pub save_data: Option<bool>,
    #[serde(rename = "@credentials")]
    #[serde(default)]
    pub credentials: Option<STCredMethod>,
    #[serde(rename = "@singleSignOnId")]
    #[serde(default)]
    pub single_sign_on_id: Option<XmlString>,
    #[serde(rename = "dbPr")]
    #[serde(default)]
    pub db_pr: Option<Box<DatabaseProperties>>,
    #[serde(rename = "olapPr")]
    #[serde(default)]
    pub olap_pr: Option<Box<OlapProperties>>,
    #[serde(rename = "webPr")]
    #[serde(default)]
    pub web_pr: Option<Box<WebQueryProperties>>,
    #[serde(rename = "textPr")]
    #[serde(default)]
    pub text_pr: Option<Box<TextImportProperties>>,
    #[serde(rename = "parameters")]
    #[serde(default)]
    pub parameters: Option<Box<Parameters>>,
    #[serde(rename = "extLst")]
    #[serde(default)]
    pub extension_list: Option<Box<ExtensionList>>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct DatabaseProperties {
    #[serde(rename = "@connection")]
    pub connection: XmlString,
    #[serde(rename = "@command")]
    #[serde(default)]
    pub command: Option<XmlString>,
    #[serde(rename = "@serverCommand")]
    #[serde(default)]
    pub server_command: Option<XmlString>,
    #[serde(rename = "@commandType")]
    #[serde(default)]
    pub command_type: Option<u32>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct OlapProperties {
    #[serde(rename = "@local")]
    #[serde(default)]
    pub local: Option<bool>,
    #[serde(rename = "@localConnection")]
    #[serde(default)]
    pub local_connection: Option<XmlString>,
    #[serde(rename = "@localRefresh")]
    #[serde(default)]
    pub local_refresh: Option<bool>,
    #[serde(rename = "@sendLocale")]
    #[serde(default)]
    pub send_locale: Option<bool>,
    #[serde(rename = "@rowDrillCount")]
    #[serde(default)]
    pub row_drill_count: Option<u32>,
    #[serde(rename = "@serverFill")]
    #[serde(default)]
    pub server_fill: Option<bool>,
    #[serde(rename = "@serverNumberFormat")]
    #[serde(default)]
    pub server_number_format: Option<bool>,
    #[serde(rename = "@serverFont")]
    #[serde(default)]
    pub server_font: Option<bool>,
    #[serde(rename = "@serverFontColor")]
    #[serde(default)]
    pub server_font_color: Option<bool>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct WebQueryProperties {
    #[serde(rename = "@xml")]
    #[serde(default)]
    pub xml: Option<bool>,
    #[serde(rename = "@sourceData")]
    #[serde(default)]
    pub source_data: Option<bool>,
    #[serde(rename = "@parsePre")]
    #[serde(default)]
    pub parse_pre: Option<bool>,
    #[serde(rename = "@consecutive")]
    #[serde(default)]
    pub consecutive: Option<bool>,
    #[serde(rename = "@firstRow")]
    #[serde(default)]
    pub first_row: Option<bool>,
    #[serde(rename = "@xl97")]
    #[serde(default)]
    pub xl97: Option<bool>,
    #[serde(rename = "@textDates")]
    #[serde(default)]
    pub text_dates: Option<bool>,
    #[serde(rename = "@xl2000")]
    #[serde(default)]
    pub xl2000: Option<bool>,
    #[serde(rename = "@url")]
    #[serde(default)]
    pub url: Option<XmlString>,
    #[serde(rename = "@post")]
    #[serde(default)]
    pub post: Option<XmlString>,
    #[serde(rename = "@htmlTables")]
    #[serde(default)]
    pub html_tables: Option<bool>,
    #[serde(rename = "@htmlFormat")]
    #[serde(default)]
    pub html_format: Option<STHtmlFmt>,
    #[serde(rename = "@editPage")]
    #[serde(default)]
    pub edit_page: Option<XmlString>,
    #[serde(rename = "tables")]
    #[serde(default)]
    pub tables: Option<Box<DataTables>>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct Parameters {
    #[serde(rename = "@count")]
    #[serde(default)]
    pub count: Option<u32>,
    #[serde(rename = "parameter")]
    #[serde(default)]
    pub parameter: Vec<Box<Parameter>>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct Parameter {
    #[serde(rename = "@name")]
    #[serde(default)]
    pub name: Option<XmlString>,
    #[serde(rename = "@sqlType")]
    #[serde(default)]
    pub sql_type: Option<i32>,
    #[serde(rename = "@parameterType")]
    #[serde(default)]
    pub parameter_type: Option<STParameterType>,
    #[serde(rename = "@refreshOnChange")]
    #[serde(default)]
    pub refresh_on_change: Option<bool>,
    #[serde(rename = "@prompt")]
    #[serde(default)]
    pub prompt: Option<XmlString>,
    #[serde(rename = "@boolean")]
    #[serde(default)]
    pub boolean: Option<bool>,
    #[serde(rename = "@double")]
    #[serde(default)]
    pub double: Option<f64>,
    #[serde(rename = "@integer")]
    #[serde(default)]
    pub integer: Option<i32>,
    #[serde(rename = "@string")]
    #[serde(default)]
    pub string: Option<XmlString>,
    #[serde(rename = "@cell")]
    #[serde(default)]
    pub cell: Option<XmlString>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct DataTables {
    #[serde(rename = "@count")]
    #[serde(default)]
    pub count: Option<u32>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct TableMissing;

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct TextImportProperties {
    #[serde(rename = "@prompt")]
    #[serde(default)]
    pub prompt: Option<bool>,
    #[serde(rename = "@fileType")]
    #[serde(default)]
    pub file_type: Option<STFileType>,
    #[serde(rename = "@codePage")]
    #[serde(default)]
    pub code_page: Option<u32>,
    #[serde(rename = "@characterSet")]
    #[serde(default)]
    pub character_set: Option<String>,
    #[serde(rename = "@firstRow")]
    #[serde(default)]
    pub first_row: Option<u32>,
    #[serde(rename = "@sourceFile")]
    #[serde(default)]
    pub source_file: Option<XmlString>,
    #[serde(rename = "@delimited")]
    #[serde(default)]
    pub delimited: Option<bool>,
    #[serde(rename = "@decimal")]
    #[serde(default)]
    pub decimal: Option<XmlString>,
    #[serde(rename = "@thousands")]
    #[serde(default)]
    pub thousands: Option<XmlString>,
    #[serde(rename = "@tab")]
    #[serde(default)]
    pub tab: Option<bool>,
    #[serde(rename = "@space")]
    #[serde(default)]
    pub space: Option<bool>,
    #[serde(rename = "@comma")]
    #[serde(default)]
    pub comma: Option<bool>,
    #[serde(rename = "@semicolon")]
    #[serde(default)]
    pub semicolon: Option<bool>,
    #[serde(rename = "@consecutive")]
    #[serde(default)]
    pub consecutive: Option<bool>,
    #[serde(rename = "@qualifier")]
    #[serde(default)]
    pub qualifier: Option<STQualifier>,
    #[serde(rename = "@delimiter")]
    #[serde(default)]
    pub delimiter: Option<XmlString>,
    #[serde(rename = "textFields")]
    #[serde(default)]
    pub text_fields: Option<Box<TextFields>>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct TextFields {
    #[serde(rename = "@count")]
    #[serde(default)]
    pub count: Option<u32>,
    #[serde(rename = "textField")]
    #[serde(default)]
    pub text_field: Vec<Box<TextField>>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct TextField {
    #[serde(rename = "@type")]
    #[serde(default)]
    pub r#type: Option<STExternalConnectionType>,
    #[serde(rename = "@position")]
    #[serde(default)]
    pub position: Option<u32>,
}

pub type SmlPivotCacheDefinition = Box<PivotCacheDefinition>;

pub type SmlPivotCacheRecords = Box<PivotCacheRecords>;

pub type SmlPivotTableDefinition = Box<CTPivotTableDefinition>;

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct PivotCacheDefinition {
    #[serde(rename = "@invalid")]
    #[serde(default)]
    pub invalid: Option<bool>,
    #[serde(rename = "@saveData")]
    #[serde(default)]
    pub save_data: Option<bool>,
    #[serde(rename = "@refreshOnLoad")]
    #[serde(default)]
    pub refresh_on_load: Option<bool>,
    #[serde(rename = "@optimizeMemory")]
    #[serde(default)]
    pub optimize_memory: Option<bool>,
    #[serde(rename = "@enableRefresh")]
    #[serde(default)]
    pub enable_refresh: Option<bool>,
    #[serde(rename = "@refreshedBy")]
    #[serde(default)]
    pub refreshed_by: Option<XmlString>,
    #[serde(rename = "@refreshedDate")]
    #[serde(default)]
    pub refreshed_date: Option<f64>,
    #[serde(rename = "@refreshedDateIso")]
    #[serde(default)]
    pub refreshed_date_iso: Option<String>,
    #[serde(rename = "@backgroundQuery")]
    #[serde(default)]
    pub background_query: Option<bool>,
    #[serde(rename = "@missingItemsLimit")]
    #[serde(default)]
    pub missing_items_limit: Option<u32>,
    #[serde(rename = "@createdVersion")]
    #[serde(default)]
    pub created_version: Option<u8>,
    #[serde(rename = "@refreshedVersion")]
    #[serde(default)]
    pub refreshed_version: Option<u8>,
    #[serde(rename = "@minRefreshableVersion")]
    #[serde(default)]
    pub min_refreshable_version: Option<u8>,
    #[serde(rename = "@recordCount")]
    #[serde(default)]
    pub record_count: Option<u32>,
    #[serde(rename = "@upgradeOnRefresh")]
    #[serde(default)]
    pub upgrade_on_refresh: Option<bool>,
    #[serde(rename = "@tupleCache")]
    #[serde(default)]
    pub tuple_cache: Option<bool>,
    #[serde(rename = "@supportSubquery")]
    #[serde(default)]
    pub support_subquery: Option<bool>,
    #[serde(rename = "@supportAdvancedDrill")]
    #[serde(default)]
    pub support_advanced_drill: Option<bool>,
    #[serde(rename = "cacheSource")]
    pub cache_source: Box<CacheSource>,
    #[serde(rename = "cacheFields")]
    pub cache_fields: Box<CacheFields>,
    #[serde(rename = "cacheHierarchies")]
    #[serde(default)]
    pub cache_hierarchies: Option<Box<CTCacheHierarchies>>,
    #[serde(rename = "kpis")]
    #[serde(default)]
    pub kpis: Option<Box<CTPCDKPIs>>,
    #[serde(rename = "calculatedItems")]
    #[serde(default)]
    pub calculated_items: Option<Box<CTCalculatedItems>>,
    #[serde(rename = "calculatedMembers")]
    #[serde(default)]
    pub calculated_members: Option<Box<CTCalculatedMembers>>,
    #[serde(rename = "dimensions")]
    #[serde(default)]
    pub dimensions: Option<Box<CTDimensions>>,
    #[serde(rename = "measureGroups")]
    #[serde(default)]
    pub measure_groups: Option<Box<CTMeasureGroups>>,
    #[serde(rename = "maps")]
    #[serde(default)]
    pub maps: Option<Box<CTMeasureDimensionMaps>>,
    #[serde(rename = "extLst")]
    #[serde(default)]
    pub extension_list: Option<Box<ExtensionList>>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CacheFields {
    #[serde(rename = "@count")]
    #[serde(default)]
    pub count: Option<u32>,
    #[serde(rename = "cacheField")]
    #[serde(default)]
    pub cache_field: Vec<Box<CacheField>>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CacheField {
    #[serde(rename = "@name")]
    pub name: XmlString,
    #[serde(rename = "@caption")]
    #[serde(default)]
    pub caption: Option<XmlString>,
    #[serde(rename = "@propertyName")]
    #[serde(default)]
    pub property_name: Option<XmlString>,
    #[serde(rename = "@serverField")]
    #[serde(default)]
    pub server_field: Option<bool>,
    #[serde(rename = "@uniqueList")]
    #[serde(default)]
    pub unique_list: Option<bool>,
    #[serde(rename = "@numFmtId")]
    #[serde(default)]
    pub number_format_id: Option<STNumFmtId>,
    #[serde(rename = "@formula")]
    #[serde(default)]
    pub formula: Option<XmlString>,
    #[serde(rename = "@sqlType")]
    #[serde(default)]
    pub sql_type: Option<i32>,
    #[serde(rename = "@hierarchy")]
    #[serde(default)]
    pub hierarchy: Option<i32>,
    #[serde(rename = "@level")]
    #[serde(default)]
    pub level: Option<u32>,
    #[serde(rename = "@databaseField")]
    #[serde(default)]
    pub database_field: Option<bool>,
    #[serde(rename = "@mappingCount")]
    #[serde(default)]
    pub mapping_count: Option<u32>,
    #[serde(rename = "@memberPropertyField")]
    #[serde(default)]
    pub member_property_field: Option<bool>,
    #[serde(rename = "sharedItems")]
    #[serde(default)]
    pub shared_items: Option<Box<SharedItems>>,
    #[serde(rename = "fieldGroup")]
    #[serde(default)]
    pub field_group: Option<Box<FieldGroup>>,
    #[serde(rename = "mpMap")]
    #[serde(default)]
    pub mp_map: Vec<Box<CTX>>,
    #[serde(rename = "extLst")]
    #[serde(default)]
    pub extension_list: Option<Box<ExtensionList>>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CacheSource {
    #[serde(rename = "@type")]
    pub r#type: STSourceType,
    #[serde(rename = "@connectionId")]
    #[serde(default)]
    pub connection_id: Option<u32>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct WorksheetSource {
    #[serde(rename = "@ref")]
    #[serde(default)]
    pub reference: Option<Reference>,
    #[serde(rename = "@name")]
    #[serde(default)]
    pub name: Option<XmlString>,
    #[serde(rename = "@sheet")]
    #[serde(default)]
    pub sheet: Option<XmlString>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct Consolidation {
    #[serde(rename = "@autoPage")]
    #[serde(default)]
    pub auto_page: Option<bool>,
    #[serde(rename = "pages")]
    #[serde(default)]
    pub pages: Option<Box<CTPages>>,
    #[serde(rename = "rangeSets")]
    pub range_sets: Box<CTRangeSets>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTPages {
    #[serde(rename = "@count")]
    #[serde(default)]
    pub count: Option<u32>,
    #[serde(rename = "page")]
    #[serde(default)]
    pub page: Vec<Box<CTPCDSCPage>>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTPCDSCPage {
    #[serde(rename = "@count")]
    #[serde(default)]
    pub count: Option<u32>,
    #[serde(rename = "pageItem")]
    #[serde(default)]
    pub page_item: Vec<Box<CTPageItem>>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTPageItem {
    #[serde(rename = "@name")]
    pub name: XmlString,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTRangeSets {
    #[serde(rename = "@count")]
    #[serde(default)]
    pub count: Option<u32>,
    #[serde(rename = "rangeSet")]
    #[serde(default)]
    pub range_set: Vec<Box<CTRangeSet>>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTRangeSet {
    #[serde(rename = "@i1")]
    #[serde(default)]
    pub i1: Option<u32>,
    #[serde(rename = "@i2")]
    #[serde(default)]
    pub i2: Option<u32>,
    #[serde(rename = "@i3")]
    #[serde(default)]
    pub i3: Option<u32>,
    #[serde(rename = "@i4")]
    #[serde(default)]
    pub i4: Option<u32>,
    #[serde(rename = "@ref")]
    #[serde(default)]
    pub reference: Option<Reference>,
    #[serde(rename = "@name")]
    #[serde(default)]
    pub name: Option<XmlString>,
    #[serde(rename = "@sheet")]
    #[serde(default)]
    pub sheet: Option<XmlString>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct SharedItems {
    #[serde(rename = "@containsSemiMixedTypes")]
    #[serde(default)]
    pub contains_semi_mixed_types: Option<bool>,
    #[serde(rename = "@containsNonDate")]
    #[serde(default)]
    pub contains_non_date: Option<bool>,
    #[serde(rename = "@containsDate")]
    #[serde(default)]
    pub contains_date: Option<bool>,
    #[serde(rename = "@containsString")]
    #[serde(default)]
    pub contains_string: Option<bool>,
    #[serde(rename = "@containsBlank")]
    #[serde(default)]
    pub contains_blank: Option<bool>,
    #[serde(rename = "@containsMixedTypes")]
    #[serde(default)]
    pub contains_mixed_types: Option<bool>,
    #[serde(rename = "@containsNumber")]
    #[serde(default)]
    pub contains_number: Option<bool>,
    #[serde(rename = "@containsInteger")]
    #[serde(default)]
    pub contains_integer: Option<bool>,
    #[serde(rename = "@minValue")]
    #[serde(default)]
    pub min_value: Option<f64>,
    #[serde(rename = "@maxValue")]
    #[serde(default)]
    pub max_value: Option<f64>,
    #[serde(rename = "@minDate")]
    #[serde(default)]
    pub min_date: Option<String>,
    #[serde(rename = "@maxDate")]
    #[serde(default)]
    pub max_date: Option<String>,
    #[serde(rename = "@count")]
    #[serde(default)]
    pub count: Option<u32>,
    #[serde(rename = "@longText")]
    #[serde(default)]
    pub long_text: Option<bool>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTMissing {
    #[serde(rename = "@u")]
    #[serde(default)]
    pub u: Option<bool>,
    #[serde(rename = "@f")]
    #[serde(default)]
    pub formula: Option<bool>,
    #[serde(rename = "@c")]
    #[serde(default)]
    pub cells: Option<XmlString>,
    #[serde(rename = "@cp")]
    #[serde(default)]
    pub cp: Option<u32>,
    #[serde(rename = "@in")]
    #[serde(default)]
    pub r#in: Option<u32>,
    #[serde(rename = "@bc")]
    #[serde(default)]
    pub bc: Option<STUnsignedIntHex>,
    #[serde(rename = "@fc")]
    #[serde(default)]
    pub fc: Option<STUnsignedIntHex>,
    #[serde(rename = "@i")]
    #[serde(default)]
    pub i: Option<bool>,
    #[serde(rename = "@un")]
    #[serde(default)]
    pub un: Option<bool>,
    #[serde(rename = "@st")]
    #[serde(default)]
    pub st: Option<bool>,
    #[serde(rename = "@b")]
    #[serde(default)]
    pub b: Option<bool>,
    #[serde(rename = "tpls")]
    #[serde(default)]
    pub tpls: Vec<Box<CTTuples>>,
    #[serde(rename = "x")]
    #[serde(default)]
    pub x: Vec<Box<CTX>>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTNumber {
    #[serde(rename = "@v")]
    pub value: f64,
    #[serde(rename = "@u")]
    #[serde(default)]
    pub u: Option<bool>,
    #[serde(rename = "@f")]
    #[serde(default)]
    pub formula: Option<bool>,
    #[serde(rename = "@c")]
    #[serde(default)]
    pub cells: Option<XmlString>,
    #[serde(rename = "@cp")]
    #[serde(default)]
    pub cp: Option<u32>,
    #[serde(rename = "@in")]
    #[serde(default)]
    pub r#in: Option<u32>,
    #[serde(rename = "@bc")]
    #[serde(default)]
    pub bc: Option<STUnsignedIntHex>,
    #[serde(rename = "@fc")]
    #[serde(default)]
    pub fc: Option<STUnsignedIntHex>,
    #[serde(rename = "@i")]
    #[serde(default)]
    pub i: Option<bool>,
    #[serde(rename = "@un")]
    #[serde(default)]
    pub un: Option<bool>,
    #[serde(rename = "@st")]
    #[serde(default)]
    pub st: Option<bool>,
    #[serde(rename = "@b")]
    #[serde(default)]
    pub b: Option<bool>,
    #[serde(rename = "tpls")]
    #[serde(default)]
    pub tpls: Vec<Box<CTTuples>>,
    #[serde(rename = "x")]
    #[serde(default)]
    pub x: Vec<Box<CTX>>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTBoolean {
    #[serde(rename = "@v")]
    pub value: bool,
    #[serde(rename = "@u")]
    #[serde(default)]
    pub u: Option<bool>,
    #[serde(rename = "@f")]
    #[serde(default)]
    pub formula: Option<bool>,
    #[serde(rename = "@c")]
    #[serde(default)]
    pub cells: Option<XmlString>,
    #[serde(rename = "@cp")]
    #[serde(default)]
    pub cp: Option<u32>,
    #[serde(rename = "x")]
    #[serde(default)]
    pub x: Vec<Box<CTX>>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTError {
    #[serde(rename = "@v")]
    pub value: XmlString,
    #[serde(rename = "@u")]
    #[serde(default)]
    pub u: Option<bool>,
    #[serde(rename = "@f")]
    #[serde(default)]
    pub formula: Option<bool>,
    #[serde(rename = "@c")]
    #[serde(default)]
    pub cells: Option<XmlString>,
    #[serde(rename = "@cp")]
    #[serde(default)]
    pub cp: Option<u32>,
    #[serde(rename = "@in")]
    #[serde(default)]
    pub r#in: Option<u32>,
    #[serde(rename = "@bc")]
    #[serde(default)]
    pub bc: Option<STUnsignedIntHex>,
    #[serde(rename = "@fc")]
    #[serde(default)]
    pub fc: Option<STUnsignedIntHex>,
    #[serde(rename = "@i")]
    #[serde(default)]
    pub i: Option<bool>,
    #[serde(rename = "@un")]
    #[serde(default)]
    pub un: Option<bool>,
    #[serde(rename = "@st")]
    #[serde(default)]
    pub st: Option<bool>,
    #[serde(rename = "@b")]
    #[serde(default)]
    pub b: Option<bool>,
    #[serde(rename = "tpls")]
    #[serde(default)]
    pub tpls: Option<Box<CTTuples>>,
    #[serde(rename = "x")]
    #[serde(default)]
    pub x: Vec<Box<CTX>>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTString {
    #[serde(rename = "@v")]
    pub value: XmlString,
    #[serde(rename = "@u")]
    #[serde(default)]
    pub u: Option<bool>,
    #[serde(rename = "@f")]
    #[serde(default)]
    pub formula: Option<bool>,
    #[serde(rename = "@c")]
    #[serde(default)]
    pub cells: Option<XmlString>,
    #[serde(rename = "@cp")]
    #[serde(default)]
    pub cp: Option<u32>,
    #[serde(rename = "@in")]
    #[serde(default)]
    pub r#in: Option<u32>,
    #[serde(rename = "@bc")]
    #[serde(default)]
    pub bc: Option<STUnsignedIntHex>,
    #[serde(rename = "@fc")]
    #[serde(default)]
    pub fc: Option<STUnsignedIntHex>,
    #[serde(rename = "@i")]
    #[serde(default)]
    pub i: Option<bool>,
    #[serde(rename = "@un")]
    #[serde(default)]
    pub un: Option<bool>,
    #[serde(rename = "@st")]
    #[serde(default)]
    pub st: Option<bool>,
    #[serde(rename = "@b")]
    #[serde(default)]
    pub b: Option<bool>,
    #[serde(rename = "tpls")]
    #[serde(default)]
    pub tpls: Vec<Box<CTTuples>>,
    #[serde(rename = "x")]
    #[serde(default)]
    pub x: Vec<Box<CTX>>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTDateTime {
    #[serde(rename = "@v")]
    pub value: String,
    #[serde(rename = "@u")]
    #[serde(default)]
    pub u: Option<bool>,
    #[serde(rename = "@f")]
    #[serde(default)]
    pub formula: Option<bool>,
    #[serde(rename = "@c")]
    #[serde(default)]
    pub cells: Option<XmlString>,
    #[serde(rename = "@cp")]
    #[serde(default)]
    pub cp: Option<u32>,
    #[serde(rename = "x")]
    #[serde(default)]
    pub x: Vec<Box<CTX>>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct FieldGroup {
    #[serde(rename = "@par")]
    #[serde(default)]
    pub par: Option<u32>,
    #[serde(rename = "@base")]
    #[serde(default)]
    pub base: Option<u32>,
    #[serde(rename = "rangePr")]
    #[serde(default)]
    pub range_pr: Option<Box<CTRangePr>>,
    #[serde(rename = "discretePr")]
    #[serde(default)]
    pub discrete_pr: Option<Box<CTDiscretePr>>,
    #[serde(rename = "groupItems")]
    #[serde(default)]
    pub group_items: Option<Box<GroupItems>>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTRangePr {
    #[serde(rename = "@autoStart")]
    #[serde(default)]
    pub auto_start: Option<bool>,
    #[serde(rename = "@autoEnd")]
    #[serde(default)]
    pub auto_end: Option<bool>,
    #[serde(rename = "@groupBy")]
    #[serde(default)]
    pub group_by: Option<STGroupBy>,
    #[serde(rename = "@startNum")]
    #[serde(default)]
    pub start_num: Option<f64>,
    #[serde(rename = "@endNum")]
    #[serde(default)]
    pub end_num: Option<f64>,
    #[serde(rename = "@startDate")]
    #[serde(default)]
    pub start_date: Option<String>,
    #[serde(rename = "@endDate")]
    #[serde(default)]
    pub end_date: Option<String>,
    #[serde(rename = "@groupInterval")]
    #[serde(default)]
    pub group_interval: Option<f64>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTDiscretePr {
    #[serde(rename = "@count")]
    #[serde(default)]
    pub count: Option<u32>,
    #[serde(rename = "x")]
    #[serde(default)]
    pub x: Vec<Box<CTIndex>>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct GroupItems {
    #[serde(rename = "@count")]
    #[serde(default)]
    pub count: Option<u32>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct PivotCacheRecords {
    #[serde(rename = "@count")]
    #[serde(default)]
    pub count: Option<u32>,
    #[serde(rename = "r")]
    #[serde(default)]
    pub reference: Vec<Box<CTRecord>>,
    #[serde(rename = "extLst")]
    #[serde(default)]
    pub extension_list: Option<Box<ExtensionList>>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct CTRecord;

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTPCDKPIs {
    #[serde(rename = "@count")]
    #[serde(default)]
    pub count: Option<u32>,
    #[serde(rename = "kpi")]
    #[serde(default)]
    pub kpi: Vec<Box<CTPCDKPI>>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTPCDKPI {
    #[serde(rename = "@uniqueName")]
    pub unique_name: XmlString,
    #[serde(rename = "@caption")]
    #[serde(default)]
    pub caption: Option<XmlString>,
    #[serde(rename = "@displayFolder")]
    #[serde(default)]
    pub display_folder: Option<XmlString>,
    #[serde(rename = "@measureGroup")]
    #[serde(default)]
    pub measure_group: Option<XmlString>,
    #[serde(rename = "@parent")]
    #[serde(default)]
    pub parent: Option<XmlString>,
    #[serde(rename = "@value")]
    pub value: XmlString,
    #[serde(rename = "@goal")]
    #[serde(default)]
    pub goal: Option<XmlString>,
    #[serde(rename = "@status")]
    #[serde(default)]
    pub status: Option<XmlString>,
    #[serde(rename = "@trend")]
    #[serde(default)]
    pub trend: Option<XmlString>,
    #[serde(rename = "@weight")]
    #[serde(default)]
    pub weight: Option<XmlString>,
    #[serde(rename = "@time")]
    #[serde(default)]
    pub time: Option<XmlString>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTCacheHierarchies {
    #[serde(rename = "@count")]
    #[serde(default)]
    pub count: Option<u32>,
    #[serde(rename = "cacheHierarchy")]
    #[serde(default)]
    pub cache_hierarchy: Vec<Box<CTCacheHierarchy>>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTCacheHierarchy {
    #[serde(rename = "@uniqueName")]
    pub unique_name: XmlString,
    #[serde(rename = "@caption")]
    #[serde(default)]
    pub caption: Option<XmlString>,
    #[serde(rename = "@measure")]
    #[serde(default)]
    pub measure: Option<bool>,
    #[serde(rename = "@set")]
    #[serde(default)]
    pub set: Option<bool>,
    #[serde(rename = "@parentSet")]
    #[serde(default)]
    pub parent_set: Option<u32>,
    #[serde(rename = "@iconSet")]
    #[serde(default)]
    pub icon_set: Option<i32>,
    #[serde(rename = "@attribute")]
    #[serde(default)]
    pub attribute: Option<bool>,
    #[serde(rename = "@time")]
    #[serde(default)]
    pub time: Option<bool>,
    #[serde(rename = "@keyAttribute")]
    #[serde(default)]
    pub key_attribute: Option<bool>,
    #[serde(rename = "@defaultMemberUniqueName")]
    #[serde(default)]
    pub default_member_unique_name: Option<XmlString>,
    #[serde(rename = "@allUniqueName")]
    #[serde(default)]
    pub all_unique_name: Option<XmlString>,
    #[serde(rename = "@allCaption")]
    #[serde(default)]
    pub all_caption: Option<XmlString>,
    #[serde(rename = "@dimensionUniqueName")]
    #[serde(default)]
    pub dimension_unique_name: Option<XmlString>,
    #[serde(rename = "@displayFolder")]
    #[serde(default)]
    pub display_folder: Option<XmlString>,
    #[serde(rename = "@measureGroup")]
    #[serde(default)]
    pub measure_group: Option<XmlString>,
    #[serde(rename = "@measures")]
    #[serde(default)]
    pub measures: Option<bool>,
    #[serde(rename = "@count")]
    pub count: u32,
    #[serde(rename = "@oneField")]
    #[serde(default)]
    pub one_field: Option<bool>,
    #[serde(rename = "@memberValueDatatype")]
    #[serde(default)]
    pub member_value_datatype: Option<u16>,
    #[serde(rename = "@unbalanced")]
    #[serde(default)]
    pub unbalanced: Option<bool>,
    #[serde(rename = "@unbalancedGroup")]
    #[serde(default)]
    pub unbalanced_group: Option<bool>,
    #[serde(rename = "@hidden")]
    #[serde(default)]
    pub hidden: Option<bool>,
    #[serde(rename = "fieldsUsage")]
    #[serde(default)]
    pub fields_usage: Option<Box<CTFieldsUsage>>,
    #[serde(rename = "groupLevels")]
    #[serde(default)]
    pub group_levels: Option<Box<CTGroupLevels>>,
    #[serde(rename = "extLst")]
    #[serde(default)]
    pub extension_list: Option<Box<ExtensionList>>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTFieldsUsage {
    #[serde(rename = "@count")]
    #[serde(default)]
    pub count: Option<u32>,
    #[serde(rename = "fieldUsage")]
    #[serde(default)]
    pub field_usage: Vec<Box<CTFieldUsage>>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTFieldUsage {
    #[serde(rename = "@x")]
    pub x: i32,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTGroupLevels {
    #[serde(rename = "@count")]
    #[serde(default)]
    pub count: Option<u32>,
    #[serde(rename = "groupLevel")]
    #[serde(default)]
    pub group_level: Vec<Box<CTGroupLevel>>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTGroupLevel {
    #[serde(rename = "@uniqueName")]
    pub unique_name: XmlString,
    #[serde(rename = "@caption")]
    pub caption: XmlString,
    #[serde(rename = "@user")]
    #[serde(default)]
    pub user: Option<bool>,
    #[serde(rename = "@customRollUp")]
    #[serde(default)]
    pub custom_roll_up: Option<bool>,
    #[serde(rename = "groups")]
    #[serde(default)]
    pub groups: Option<Box<CTGroups>>,
    #[serde(rename = "extLst")]
    #[serde(default)]
    pub extension_list: Option<Box<ExtensionList>>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTGroups {
    #[serde(rename = "@count")]
    #[serde(default)]
    pub count: Option<u32>,
    #[serde(rename = "group")]
    #[serde(default)]
    pub group: Vec<Box<CTLevelGroup>>,
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
    #[serde(default)]
    pub unique_parent: Option<XmlString>,
    #[serde(rename = "@id")]
    #[serde(default)]
    pub id: Option<i32>,
    #[serde(rename = "groupMembers")]
    pub group_members: Box<CTGroupMembers>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTGroupMembers {
    #[serde(rename = "@count")]
    #[serde(default)]
    pub count: Option<u32>,
    #[serde(rename = "groupMember")]
    #[serde(default)]
    pub group_member: Vec<Box<CTGroupMember>>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTGroupMember {
    #[serde(rename = "@uniqueName")]
    pub unique_name: XmlString,
    #[serde(rename = "@group")]
    #[serde(default)]
    pub group: Option<bool>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTTupleCache {
    #[serde(rename = "entries")]
    #[serde(default)]
    pub entries: Option<Box<CTPCDSDTCEntries>>,
    #[serde(rename = "sets")]
    #[serde(default)]
    pub sets: Option<Box<CTSets>>,
    #[serde(rename = "queryCache")]
    #[serde(default)]
    pub query_cache: Option<Box<CTQueryCache>>,
    #[serde(rename = "serverFormats")]
    #[serde(default)]
    pub server_formats: Option<Box<CTServerFormats>>,
    #[serde(rename = "extLst")]
    #[serde(default)]
    pub extension_list: Option<Box<ExtensionList>>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTServerFormat {
    #[serde(rename = "@culture")]
    #[serde(default)]
    pub culture: Option<XmlString>,
    #[serde(rename = "@format")]
    #[serde(default)]
    pub format: Option<XmlString>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTServerFormats {
    #[serde(rename = "@count")]
    #[serde(default)]
    pub count: Option<u32>,
    #[serde(rename = "serverFormat")]
    #[serde(default)]
    pub server_format: Vec<Box<CTServerFormat>>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTPCDSDTCEntries {
    #[serde(rename = "@count")]
    #[serde(default)]
    pub count: Option<u32>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTTuples {
    #[serde(rename = "@c")]
    #[serde(default)]
    pub cells: Option<u32>,
    #[serde(rename = "tpl")]
    #[serde(default)]
    pub tpl: Vec<Box<CTTuple>>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTTuple {
    #[serde(rename = "@fld")]
    #[serde(default)]
    pub fld: Option<u32>,
    #[serde(rename = "@hier")]
    #[serde(default)]
    pub hier: Option<u32>,
    #[serde(rename = "@item")]
    pub item: u32,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTSets {
    #[serde(rename = "@count")]
    #[serde(default)]
    pub count: Option<u32>,
    #[serde(rename = "set")]
    #[serde(default)]
    pub set: Vec<Box<CTSet>>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTSet {
    #[serde(rename = "@count")]
    #[serde(default)]
    pub count: Option<u32>,
    #[serde(rename = "@maxRank")]
    pub max_rank: i32,
    #[serde(rename = "@setDefinition")]
    pub set_definition: XmlString,
    #[serde(rename = "@sortType")]
    #[serde(default)]
    pub sort_type: Option<STSortType>,
    #[serde(rename = "@queryFailed")]
    #[serde(default)]
    pub query_failed: Option<bool>,
    #[serde(rename = "tpls")]
    #[serde(default)]
    pub tpls: Vec<Box<CTTuples>>,
    #[serde(rename = "sortByTuple")]
    #[serde(default)]
    pub sort_by_tuple: Option<Box<CTTuples>>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTQueryCache {
    #[serde(rename = "@count")]
    #[serde(default)]
    pub count: Option<u32>,
    #[serde(rename = "query")]
    #[serde(default)]
    pub query: Vec<Box<CTQuery>>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTQuery {
    #[serde(rename = "@mdx")]
    pub mdx: XmlString,
    #[serde(rename = "tpls")]
    #[serde(default)]
    pub tpls: Option<Box<CTTuples>>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTCalculatedItems {
    #[serde(rename = "@count")]
    #[serde(default)]
    pub count: Option<u32>,
    #[serde(rename = "calculatedItem")]
    #[serde(default)]
    pub calculated_item: Vec<Box<CTCalculatedItem>>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTCalculatedItem {
    #[serde(rename = "@field")]
    #[serde(default)]
    pub field: Option<u32>,
    #[serde(rename = "@formula")]
    #[serde(default)]
    pub formula: Option<XmlString>,
    #[serde(rename = "pivotArea")]
    pub pivot_area: Box<PivotArea>,
    #[serde(rename = "extLst")]
    #[serde(default)]
    pub extension_list: Option<Box<ExtensionList>>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTCalculatedMembers {
    #[serde(rename = "@count")]
    #[serde(default)]
    pub count: Option<u32>,
    #[serde(rename = "calculatedMember")]
    #[serde(default)]
    pub calculated_member: Vec<Box<CTCalculatedMember>>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTCalculatedMember {
    #[serde(rename = "@name")]
    pub name: XmlString,
    #[serde(rename = "@mdx")]
    pub mdx: XmlString,
    #[serde(rename = "@memberName")]
    #[serde(default)]
    pub member_name: Option<XmlString>,
    #[serde(rename = "@hierarchy")]
    #[serde(default)]
    pub hierarchy: Option<XmlString>,
    #[serde(rename = "@parent")]
    #[serde(default)]
    pub parent: Option<XmlString>,
    #[serde(rename = "@solveOrder")]
    #[serde(default)]
    pub solve_order: Option<i32>,
    #[serde(rename = "@set")]
    #[serde(default)]
    pub set: Option<bool>,
    #[serde(rename = "extLst")]
    #[serde(default)]
    pub extension_list: Option<Box<ExtensionList>>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTPivotTableDefinition {
    #[serde(rename = "@name")]
    pub name: XmlString,
    #[serde(rename = "@cacheId")]
    pub cache_id: u32,
    #[serde(rename = "@dataOnRows")]
    #[serde(default)]
    pub data_on_rows: Option<bool>,
    #[serde(rename = "@dataPosition")]
    #[serde(default)]
    pub data_position: Option<u32>,
    #[serde(rename = "@dataCaption")]
    pub data_caption: XmlString,
    #[serde(rename = "@grandTotalCaption")]
    #[serde(default)]
    pub grand_total_caption: Option<XmlString>,
    #[serde(rename = "@errorCaption")]
    #[serde(default)]
    pub error_caption: Option<XmlString>,
    #[serde(rename = "@showError")]
    #[serde(default)]
    pub show_error: Option<bool>,
    #[serde(rename = "@missingCaption")]
    #[serde(default)]
    pub missing_caption: Option<XmlString>,
    #[serde(rename = "@showMissing")]
    #[serde(default)]
    pub show_missing: Option<bool>,
    #[serde(rename = "@pageStyle")]
    #[serde(default)]
    pub page_style: Option<XmlString>,
    #[serde(rename = "@pivotTableStyle")]
    #[serde(default)]
    pub pivot_table_style: Option<XmlString>,
    #[serde(rename = "@vacatedStyle")]
    #[serde(default)]
    pub vacated_style: Option<XmlString>,
    #[serde(rename = "@tag")]
    #[serde(default)]
    pub tag: Option<XmlString>,
    #[serde(rename = "@updatedVersion")]
    #[serde(default)]
    pub updated_version: Option<u8>,
    #[serde(rename = "@minRefreshableVersion")]
    #[serde(default)]
    pub min_refreshable_version: Option<u8>,
    #[serde(rename = "@asteriskTotals")]
    #[serde(default)]
    pub asterisk_totals: Option<bool>,
    #[serde(rename = "@showItems")]
    #[serde(default)]
    pub show_items: Option<bool>,
    #[serde(rename = "@editData")]
    #[serde(default)]
    pub edit_data: Option<bool>,
    #[serde(rename = "@disableFieldList")]
    #[serde(default)]
    pub disable_field_list: Option<bool>,
    #[serde(rename = "@showCalcMbrs")]
    #[serde(default)]
    pub show_calc_mbrs: Option<bool>,
    #[serde(rename = "@visualTotals")]
    #[serde(default)]
    pub visual_totals: Option<bool>,
    #[serde(rename = "@showMultipleLabel")]
    #[serde(default)]
    pub show_multiple_label: Option<bool>,
    #[serde(rename = "@showDataDropDown")]
    #[serde(default)]
    pub show_data_drop_down: Option<bool>,
    #[serde(rename = "@showDrill")]
    #[serde(default)]
    pub show_drill: Option<bool>,
    #[serde(rename = "@printDrill")]
    #[serde(default)]
    pub print_drill: Option<bool>,
    #[serde(rename = "@showMemberPropertyTips")]
    #[serde(default)]
    pub show_member_property_tips: Option<bool>,
    #[serde(rename = "@showDataTips")]
    #[serde(default)]
    pub show_data_tips: Option<bool>,
    #[serde(rename = "@enableWizard")]
    #[serde(default)]
    pub enable_wizard: Option<bool>,
    #[serde(rename = "@enableDrill")]
    #[serde(default)]
    pub enable_drill: Option<bool>,
    #[serde(rename = "@enableFieldProperties")]
    #[serde(default)]
    pub enable_field_properties: Option<bool>,
    #[serde(rename = "@preserveFormatting")]
    #[serde(default)]
    pub preserve_formatting: Option<bool>,
    #[serde(rename = "@useAutoFormatting")]
    #[serde(default)]
    pub use_auto_formatting: Option<bool>,
    #[serde(rename = "@pageWrap")]
    #[serde(default)]
    pub page_wrap: Option<u32>,
    #[serde(rename = "@pageOverThenDown")]
    #[serde(default)]
    pub page_over_then_down: Option<bool>,
    #[serde(rename = "@subtotalHiddenItems")]
    #[serde(default)]
    pub subtotal_hidden_items: Option<bool>,
    #[serde(rename = "@rowGrandTotals")]
    #[serde(default)]
    pub row_grand_totals: Option<bool>,
    #[serde(rename = "@colGrandTotals")]
    #[serde(default)]
    pub col_grand_totals: Option<bool>,
    #[serde(rename = "@fieldPrintTitles")]
    #[serde(default)]
    pub field_print_titles: Option<bool>,
    #[serde(rename = "@itemPrintTitles")]
    #[serde(default)]
    pub item_print_titles: Option<bool>,
    #[serde(rename = "@mergeItem")]
    #[serde(default)]
    pub merge_item: Option<bool>,
    #[serde(rename = "@showDropZones")]
    #[serde(default)]
    pub show_drop_zones: Option<bool>,
    #[serde(rename = "@createdVersion")]
    #[serde(default)]
    pub created_version: Option<u8>,
    #[serde(rename = "@indent")]
    #[serde(default)]
    pub indent: Option<u32>,
    #[serde(rename = "@showEmptyRow")]
    #[serde(default)]
    pub show_empty_row: Option<bool>,
    #[serde(rename = "@showEmptyCol")]
    #[serde(default)]
    pub show_empty_col: Option<bool>,
    #[serde(rename = "@showHeaders")]
    #[serde(default)]
    pub show_headers: Option<bool>,
    #[serde(rename = "@compact")]
    #[serde(default)]
    pub compact: Option<bool>,
    #[serde(rename = "@outline")]
    #[serde(default)]
    pub outline: Option<bool>,
    #[serde(rename = "@outlineData")]
    #[serde(default)]
    pub outline_data: Option<bool>,
    #[serde(rename = "@compactData")]
    #[serde(default)]
    pub compact_data: Option<bool>,
    #[serde(rename = "@published")]
    #[serde(default)]
    pub published: Option<bool>,
    #[serde(rename = "@gridDropZones")]
    #[serde(default)]
    pub grid_drop_zones: Option<bool>,
    #[serde(rename = "@immersive")]
    #[serde(default)]
    pub immersive: Option<bool>,
    #[serde(rename = "@multipleFieldFilters")]
    #[serde(default)]
    pub multiple_field_filters: Option<bool>,
    #[serde(rename = "@chartFormat")]
    #[serde(default)]
    pub chart_format: Option<u32>,
    #[serde(rename = "@rowHeaderCaption")]
    #[serde(default)]
    pub row_header_caption: Option<XmlString>,
    #[serde(rename = "@colHeaderCaption")]
    #[serde(default)]
    pub col_header_caption: Option<XmlString>,
    #[serde(rename = "@fieldListSortAscending")]
    #[serde(default)]
    pub field_list_sort_ascending: Option<bool>,
    #[serde(rename = "@mdxSubqueries")]
    #[serde(default)]
    pub mdx_subqueries: Option<bool>,
    #[serde(rename = "@customListSort")]
    #[serde(default)]
    pub custom_list_sort: Option<bool>,
    #[serde(rename = "location")]
    pub location: Box<PivotLocation>,
    #[serde(rename = "pivotFields")]
    #[serde(default)]
    pub pivot_fields: Option<Box<PivotFields>>,
    #[serde(rename = "rowFields")]
    #[serde(default)]
    pub row_fields: Option<Box<RowFields>>,
    #[serde(rename = "rowItems")]
    #[serde(default)]
    pub row_items: Option<Box<CTRowItems>>,
    #[serde(rename = "colFields")]
    #[serde(default)]
    pub col_fields: Option<Box<ColFields>>,
    #[serde(rename = "colItems")]
    #[serde(default)]
    pub col_items: Option<Box<CTColItems>>,
    #[serde(rename = "pageFields")]
    #[serde(default)]
    pub page_fields: Option<Box<PageFields>>,
    #[serde(rename = "dataFields")]
    #[serde(default)]
    pub data_fields: Option<Box<DataFields>>,
    #[serde(rename = "formats")]
    #[serde(default)]
    pub formats: Option<Box<CTFormats>>,
    #[serde(rename = "conditionalFormats")]
    #[serde(default)]
    pub conditional_formats: Option<Box<CTConditionalFormats>>,
    #[serde(rename = "chartFormats")]
    #[serde(default)]
    pub chart_formats: Option<Box<CTChartFormats>>,
    #[serde(rename = "pivotHierarchies")]
    #[serde(default)]
    pub pivot_hierarchies: Option<Box<CTPivotHierarchies>>,
    #[serde(rename = "pivotTableStyleInfo")]
    #[serde(default)]
    pub pivot_table_style_info: Option<Box<CTPivotTableStyle>>,
    #[serde(rename = "filters")]
    #[serde(default)]
    pub filters: Option<Box<PivotFilters>>,
    #[serde(rename = "rowHierarchiesUsage")]
    #[serde(default)]
    pub row_hierarchies_usage: Option<Box<CTRowHierarchiesUsage>>,
    #[serde(rename = "colHierarchiesUsage")]
    #[serde(default)]
    pub col_hierarchies_usage: Option<Box<CTColHierarchiesUsage>>,
    #[serde(rename = "extLst")]
    #[serde(default)]
    pub extension_list: Option<Box<ExtensionList>>,
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
    #[serde(default)]
    pub row_page_count: Option<u32>,
    #[serde(rename = "@colPageCount")]
    #[serde(default)]
    pub col_page_count: Option<u32>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct PivotFields {
    #[serde(rename = "@count")]
    #[serde(default)]
    pub count: Option<u32>,
    #[serde(rename = "pivotField")]
    #[serde(default)]
    pub pivot_field: Vec<Box<PivotField>>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct PivotField {
    #[serde(rename = "@name")]
    #[serde(default)]
    pub name: Option<XmlString>,
    #[serde(rename = "@axis")]
    #[serde(default)]
    pub axis: Option<STAxis>,
    #[serde(rename = "@dataField")]
    #[serde(default)]
    pub data_field: Option<bool>,
    #[serde(rename = "@subtotalCaption")]
    #[serde(default)]
    pub subtotal_caption: Option<XmlString>,
    #[serde(rename = "@showDropDowns")]
    #[serde(default)]
    pub show_drop_downs: Option<bool>,
    #[serde(rename = "@hiddenLevel")]
    #[serde(default)]
    pub hidden_level: Option<bool>,
    #[serde(rename = "@uniqueMemberProperty")]
    #[serde(default)]
    pub unique_member_property: Option<XmlString>,
    #[serde(rename = "@compact")]
    #[serde(default)]
    pub compact: Option<bool>,
    #[serde(rename = "@allDrilled")]
    #[serde(default)]
    pub all_drilled: Option<bool>,
    #[serde(rename = "@numFmtId")]
    #[serde(default)]
    pub number_format_id: Option<STNumFmtId>,
    #[serde(rename = "@outline")]
    #[serde(default)]
    pub outline: Option<bool>,
    #[serde(rename = "@subtotalTop")]
    #[serde(default)]
    pub subtotal_top: Option<bool>,
    #[serde(rename = "@dragToRow")]
    #[serde(default)]
    pub drag_to_row: Option<bool>,
    #[serde(rename = "@dragToCol")]
    #[serde(default)]
    pub drag_to_col: Option<bool>,
    #[serde(rename = "@multipleItemSelectionAllowed")]
    #[serde(default)]
    pub multiple_item_selection_allowed: Option<bool>,
    #[serde(rename = "@dragToPage")]
    #[serde(default)]
    pub drag_to_page: Option<bool>,
    #[serde(rename = "@dragToData")]
    #[serde(default)]
    pub drag_to_data: Option<bool>,
    #[serde(rename = "@dragOff")]
    #[serde(default)]
    pub drag_off: Option<bool>,
    #[serde(rename = "@showAll")]
    #[serde(default)]
    pub show_all: Option<bool>,
    #[serde(rename = "@insertBlankRow")]
    #[serde(default)]
    pub insert_blank_row: Option<bool>,
    #[serde(rename = "@serverField")]
    #[serde(default)]
    pub server_field: Option<bool>,
    #[serde(rename = "@insertPageBreak")]
    #[serde(default)]
    pub insert_page_break: Option<bool>,
    #[serde(rename = "@autoShow")]
    #[serde(default)]
    pub auto_show: Option<bool>,
    #[serde(rename = "@topAutoShow")]
    #[serde(default)]
    pub top_auto_show: Option<bool>,
    #[serde(rename = "@hideNewItems")]
    #[serde(default)]
    pub hide_new_items: Option<bool>,
    #[serde(rename = "@measureFilter")]
    #[serde(default)]
    pub measure_filter: Option<bool>,
    #[serde(rename = "@includeNewItemsInFilter")]
    #[serde(default)]
    pub include_new_items_in_filter: Option<bool>,
    #[serde(rename = "@itemPageCount")]
    #[serde(default)]
    pub item_page_count: Option<u32>,
    #[serde(rename = "@sortType")]
    #[serde(default)]
    pub sort_type: Option<STFieldSortType>,
    #[serde(rename = "@dataSourceSort")]
    #[serde(default)]
    pub data_source_sort: Option<bool>,
    #[serde(rename = "@nonAutoSortDefault")]
    #[serde(default)]
    pub non_auto_sort_default: Option<bool>,
    #[serde(rename = "@rankBy")]
    #[serde(default)]
    pub rank_by: Option<u32>,
    #[serde(rename = "@defaultSubtotal")]
    #[serde(default)]
    pub default_subtotal: Option<bool>,
    #[serde(rename = "@sumSubtotal")]
    #[serde(default)]
    pub sum_subtotal: Option<bool>,
    #[serde(rename = "@countASubtotal")]
    #[serde(default)]
    pub count_a_subtotal: Option<bool>,
    #[serde(rename = "@avgSubtotal")]
    #[serde(default)]
    pub avg_subtotal: Option<bool>,
    #[serde(rename = "@maxSubtotal")]
    #[serde(default)]
    pub max_subtotal: Option<bool>,
    #[serde(rename = "@minSubtotal")]
    #[serde(default)]
    pub min_subtotal: Option<bool>,
    #[serde(rename = "@productSubtotal")]
    #[serde(default)]
    pub product_subtotal: Option<bool>,
    #[serde(rename = "@countSubtotal")]
    #[serde(default)]
    pub count_subtotal: Option<bool>,
    #[serde(rename = "@stdDevSubtotal")]
    #[serde(default)]
    pub std_dev_subtotal: Option<bool>,
    #[serde(rename = "@stdDevPSubtotal")]
    #[serde(default)]
    pub std_dev_p_subtotal: Option<bool>,
    #[serde(rename = "@varSubtotal")]
    #[serde(default)]
    pub var_subtotal: Option<bool>,
    #[serde(rename = "@varPSubtotal")]
    #[serde(default)]
    pub var_p_subtotal: Option<bool>,
    #[serde(rename = "@showPropCell")]
    #[serde(default)]
    pub show_prop_cell: Option<bool>,
    #[serde(rename = "@showPropTip")]
    #[serde(default)]
    pub show_prop_tip: Option<bool>,
    #[serde(rename = "@showPropAsCaption")]
    #[serde(default)]
    pub show_prop_as_caption: Option<bool>,
    #[serde(rename = "@defaultAttributeDrillState")]
    #[serde(default)]
    pub default_attribute_drill_state: Option<bool>,
    #[serde(rename = "items")]
    #[serde(default)]
    pub items: Option<Box<PivotItems>>,
    #[serde(rename = "autoSortScope")]
    #[serde(default)]
    pub auto_sort_scope: Option<Box<CTAutoSortScope>>,
    #[serde(rename = "extLst")]
    #[serde(default)]
    pub extension_list: Option<Box<ExtensionList>>,
}

pub type CTAutoSortScope = Box<PivotArea>;

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct PivotItems {
    #[serde(rename = "@count")]
    #[serde(default)]
    pub count: Option<u32>,
    #[serde(rename = "item")]
    #[serde(default)]
    pub item: Vec<Box<PivotItem>>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct PivotItem {
    #[serde(rename = "@n")]
    #[serde(default)]
    pub n: Option<XmlString>,
    #[serde(rename = "@t")]
    #[serde(default)]
    pub cell_type: Option<STItemType>,
    #[serde(rename = "@h")]
    #[serde(default)]
    pub height: Option<bool>,
    #[serde(rename = "@s")]
    #[serde(default)]
    pub style_index: Option<bool>,
    #[serde(rename = "@sd")]
    #[serde(default)]
    pub sd: Option<bool>,
    #[serde(rename = "@f")]
    #[serde(default)]
    pub formula: Option<bool>,
    #[serde(rename = "@m")]
    #[serde(default)]
    pub m: Option<bool>,
    #[serde(rename = "@c")]
    #[serde(default)]
    pub cells: Option<bool>,
    #[serde(rename = "@x")]
    #[serde(default)]
    pub x: Option<u32>,
    #[serde(rename = "@d")]
    #[serde(default)]
    pub d: Option<bool>,
    #[serde(rename = "@e")]
    #[serde(default)]
    pub e: Option<bool>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct PageFields {
    #[serde(rename = "@count")]
    #[serde(default)]
    pub count: Option<u32>,
    #[serde(rename = "pageField")]
    #[serde(default)]
    pub page_field: Vec<Box<PageField>>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct PageField {
    #[serde(rename = "@fld")]
    pub fld: i32,
    #[serde(rename = "@item")]
    #[serde(default)]
    pub item: Option<u32>,
    #[serde(rename = "@hier")]
    #[serde(default)]
    pub hier: Option<i32>,
    #[serde(rename = "@name")]
    #[serde(default)]
    pub name: Option<XmlString>,
    #[serde(rename = "@cap")]
    #[serde(default)]
    pub cap: Option<XmlString>,
    #[serde(rename = "extLst")]
    #[serde(default)]
    pub extension_list: Option<Box<ExtensionList>>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct DataFields {
    #[serde(rename = "@count")]
    #[serde(default)]
    pub count: Option<u32>,
    #[serde(rename = "dataField")]
    #[serde(default)]
    pub data_field: Vec<Box<DataField>>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct DataField {
    #[serde(rename = "@name")]
    #[serde(default)]
    pub name: Option<XmlString>,
    #[serde(rename = "@fld")]
    pub fld: u32,
    #[serde(rename = "@subtotal")]
    #[serde(default)]
    pub subtotal: Option<STDataConsolidateFunction>,
    #[serde(rename = "@showDataAs")]
    #[serde(default)]
    pub show_data_as: Option<STShowDataAs>,
    #[serde(rename = "@baseField")]
    #[serde(default)]
    pub base_field: Option<i32>,
    #[serde(rename = "@baseItem")]
    #[serde(default)]
    pub base_item: Option<u32>,
    #[serde(rename = "@numFmtId")]
    #[serde(default)]
    pub number_format_id: Option<STNumFmtId>,
    #[serde(rename = "extLst")]
    #[serde(default)]
    pub extension_list: Option<Box<ExtensionList>>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTRowItems {
    #[serde(rename = "@count")]
    #[serde(default)]
    pub count: Option<u32>,
    #[serde(rename = "i")]
    #[serde(default)]
    pub i: Vec<Box<CTI>>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTColItems {
    #[serde(rename = "@count")]
    #[serde(default)]
    pub count: Option<u32>,
    #[serde(rename = "i")]
    #[serde(default)]
    pub i: Vec<Box<CTI>>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTI {
    #[serde(rename = "@t")]
    #[serde(default)]
    pub cell_type: Option<STItemType>,
    #[serde(rename = "@r")]
    #[serde(default)]
    pub reference: Option<u32>,
    #[serde(rename = "@i")]
    #[serde(default)]
    pub i: Option<u32>,
    #[serde(rename = "x")]
    #[serde(default)]
    pub x: Vec<Box<CTX>>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTX {
    #[serde(rename = "@v")]
    #[serde(default)]
    pub value: Option<i32>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct RowFields {
    #[serde(rename = "@count")]
    #[serde(default)]
    pub count: Option<u32>,
    #[serde(rename = "field")]
    #[serde(default)]
    pub field: Vec<Box<CTField>>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct ColFields {
    #[serde(rename = "@count")]
    #[serde(default)]
    pub count: Option<u32>,
    #[serde(rename = "field")]
    #[serde(default)]
    pub field: Vec<Box<CTField>>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTField {
    #[serde(rename = "@x")]
    pub x: i32,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTFormats {
    #[serde(rename = "@count")]
    #[serde(default)]
    pub count: Option<u32>,
    #[serde(rename = "format")]
    #[serde(default)]
    pub format: Vec<Box<CTFormat>>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTFormat {
    #[serde(rename = "@action")]
    #[serde(default)]
    pub action: Option<STFormatAction>,
    #[serde(rename = "@dxfId")]
    #[serde(default)]
    pub dxf_id: Option<STDxfId>,
    #[serde(rename = "pivotArea")]
    pub pivot_area: Box<PivotArea>,
    #[serde(rename = "extLst")]
    #[serde(default)]
    pub extension_list: Option<Box<ExtensionList>>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTConditionalFormats {
    #[serde(rename = "@count")]
    #[serde(default)]
    pub count: Option<u32>,
    #[serde(rename = "conditionalFormat")]
    #[serde(default)]
    pub conditional_format: Vec<Box<CTConditionalFormat>>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTConditionalFormat {
    #[serde(rename = "@scope")]
    #[serde(default)]
    pub scope: Option<STScope>,
    #[serde(rename = "@type")]
    #[serde(default)]
    pub r#type: Option<STType>,
    #[serde(rename = "@priority")]
    pub priority: u32,
    #[serde(rename = "pivotAreas")]
    pub pivot_areas: Box<PivotAreas>,
    #[serde(rename = "extLst")]
    #[serde(default)]
    pub extension_list: Option<Box<ExtensionList>>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct PivotAreas {
    #[serde(rename = "@count")]
    #[serde(default)]
    pub count: Option<u32>,
    #[serde(rename = "pivotArea")]
    #[serde(default)]
    pub pivot_area: Vec<Box<PivotArea>>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTChartFormats {
    #[serde(rename = "@count")]
    #[serde(default)]
    pub count: Option<u32>,
    #[serde(rename = "chartFormat")]
    #[serde(default)]
    pub chart_format: Vec<Box<CTChartFormat>>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTChartFormat {
    #[serde(rename = "@chart")]
    pub chart: u32,
    #[serde(rename = "@format")]
    pub format: u32,
    #[serde(rename = "@series")]
    #[serde(default)]
    pub series: Option<bool>,
    #[serde(rename = "pivotArea")]
    pub pivot_area: Box<PivotArea>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTPivotHierarchies {
    #[serde(rename = "@count")]
    #[serde(default)]
    pub count: Option<u32>,
    #[serde(rename = "pivotHierarchy")]
    #[serde(default)]
    pub pivot_hierarchy: Vec<Box<CTPivotHierarchy>>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTPivotHierarchy {
    #[serde(rename = "@outline")]
    #[serde(default)]
    pub outline: Option<bool>,
    #[serde(rename = "@multipleItemSelectionAllowed")]
    #[serde(default)]
    pub multiple_item_selection_allowed: Option<bool>,
    #[serde(rename = "@subtotalTop")]
    #[serde(default)]
    pub subtotal_top: Option<bool>,
    #[serde(rename = "@showInFieldList")]
    #[serde(default)]
    pub show_in_field_list: Option<bool>,
    #[serde(rename = "@dragToRow")]
    #[serde(default)]
    pub drag_to_row: Option<bool>,
    #[serde(rename = "@dragToCol")]
    #[serde(default)]
    pub drag_to_col: Option<bool>,
    #[serde(rename = "@dragToPage")]
    #[serde(default)]
    pub drag_to_page: Option<bool>,
    #[serde(rename = "@dragToData")]
    #[serde(default)]
    pub drag_to_data: Option<bool>,
    #[serde(rename = "@dragOff")]
    #[serde(default)]
    pub drag_off: Option<bool>,
    #[serde(rename = "@includeNewItemsInFilter")]
    #[serde(default)]
    pub include_new_items_in_filter: Option<bool>,
    #[serde(rename = "@caption")]
    #[serde(default)]
    pub caption: Option<XmlString>,
    #[serde(rename = "mps")]
    #[serde(default)]
    pub mps: Option<Box<CTMemberProperties>>,
    #[serde(rename = "members")]
    #[serde(default)]
    pub members: Vec<Box<CTMembers>>,
    #[serde(rename = "extLst")]
    #[serde(default)]
    pub extension_list: Option<Box<ExtensionList>>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTRowHierarchiesUsage {
    #[serde(rename = "@count")]
    #[serde(default)]
    pub count: Option<u32>,
    #[serde(rename = "rowHierarchyUsage")]
    #[serde(default)]
    pub row_hierarchy_usage: Vec<Box<CTHierarchyUsage>>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTColHierarchiesUsage {
    #[serde(rename = "@count")]
    #[serde(default)]
    pub count: Option<u32>,
    #[serde(rename = "colHierarchyUsage")]
    #[serde(default)]
    pub col_hierarchy_usage: Vec<Box<CTHierarchyUsage>>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTHierarchyUsage {
    #[serde(rename = "@hierarchyUsage")]
    pub hierarchy_usage: i32,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTMemberProperties {
    #[serde(rename = "@count")]
    #[serde(default)]
    pub count: Option<u32>,
    #[serde(rename = "mp")]
    #[serde(default)]
    pub mp: Vec<Box<CTMemberProperty>>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTMemberProperty {
    #[serde(rename = "@name")]
    #[serde(default)]
    pub name: Option<XmlString>,
    #[serde(rename = "@showCell")]
    #[serde(default)]
    pub show_cell: Option<bool>,
    #[serde(rename = "@showTip")]
    #[serde(default)]
    pub show_tip: Option<bool>,
    #[serde(rename = "@showAsCaption")]
    #[serde(default)]
    pub show_as_caption: Option<bool>,
    #[serde(rename = "@nameLen")]
    #[serde(default)]
    pub name_len: Option<u32>,
    #[serde(rename = "@pPos")]
    #[serde(default)]
    pub p_pos: Option<u32>,
    #[serde(rename = "@pLen")]
    #[serde(default)]
    pub p_len: Option<u32>,
    #[serde(rename = "@level")]
    #[serde(default)]
    pub level: Option<u32>,
    #[serde(rename = "@field")]
    pub field: u32,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTMembers {
    #[serde(rename = "@count")]
    #[serde(default)]
    pub count: Option<u32>,
    #[serde(rename = "@level")]
    #[serde(default)]
    pub level: Option<u32>,
    #[serde(rename = "member")]
    #[serde(default)]
    pub member: Vec<Box<CTMember>>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTMember {
    #[serde(rename = "@name")]
    pub name: XmlString,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTDimensions {
    #[serde(rename = "@count")]
    #[serde(default)]
    pub count: Option<u32>,
    #[serde(rename = "dimension")]
    #[serde(default)]
    pub dimension: Vec<Box<CTPivotDimension>>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTPivotDimension {
    #[serde(rename = "@measure")]
    #[serde(default)]
    pub measure: Option<bool>,
    #[serde(rename = "@name")]
    pub name: XmlString,
    #[serde(rename = "@uniqueName")]
    pub unique_name: XmlString,
    #[serde(rename = "@caption")]
    pub caption: XmlString,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTMeasureGroups {
    #[serde(rename = "@count")]
    #[serde(default)]
    pub count: Option<u32>,
    #[serde(rename = "measureGroup")]
    #[serde(default)]
    pub measure_group: Vec<Box<CTMeasureGroup>>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTMeasureDimensionMaps {
    #[serde(rename = "@count")]
    #[serde(default)]
    pub count: Option<u32>,
    #[serde(rename = "map")]
    #[serde(default)]
    pub map: Vec<Box<CTMeasureDimensionMap>>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTMeasureGroup {
    #[serde(rename = "@name")]
    pub name: XmlString,
    #[serde(rename = "@caption")]
    pub caption: XmlString,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTMeasureDimensionMap {
    #[serde(rename = "@measureGroup")]
    #[serde(default)]
    pub measure_group: Option<u32>,
    #[serde(rename = "@dimension")]
    #[serde(default)]
    pub dimension: Option<u32>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTPivotTableStyle {
    #[serde(rename = "@name")]
    #[serde(default)]
    pub name: Option<String>,
    #[serde(rename = "@showRowHeaders")]
    #[serde(default)]
    pub show_row_headers: Option<bool>,
    #[serde(rename = "@showColHeaders")]
    #[serde(default)]
    pub show_col_headers: Option<bool>,
    #[serde(rename = "@showRowStripes")]
    #[serde(default)]
    pub show_row_stripes: Option<bool>,
    #[serde(rename = "@showColStripes")]
    #[serde(default)]
    pub show_col_stripes: Option<bool>,
    #[serde(rename = "@showLastColumn")]
    #[serde(default)]
    pub show_last_column: Option<bool>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct PivotFilters {
    #[serde(rename = "@count")]
    #[serde(default)]
    pub count: Option<u32>,
    #[serde(rename = "filter")]
    #[serde(default)]
    pub filter: Vec<Box<PivotFilter>>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct PivotFilter {
    #[serde(rename = "@fld")]
    pub fld: u32,
    #[serde(rename = "@mpFld")]
    #[serde(default)]
    pub mp_fld: Option<u32>,
    #[serde(rename = "@type")]
    pub r#type: STPivotFilterType,
    #[serde(rename = "@evalOrder")]
    #[serde(default)]
    pub eval_order: Option<i32>,
    #[serde(rename = "@id")]
    pub id: u32,
    #[serde(rename = "@iMeasureHier")]
    #[serde(default)]
    pub i_measure_hier: Option<u32>,
    #[serde(rename = "@iMeasureFld")]
    #[serde(default)]
    pub i_measure_fld: Option<u32>,
    #[serde(rename = "@name")]
    #[serde(default)]
    pub name: Option<XmlString>,
    #[serde(rename = "@description")]
    #[serde(default)]
    pub description: Option<XmlString>,
    #[serde(rename = "@stringValue1")]
    #[serde(default)]
    pub string_value1: Option<XmlString>,
    #[serde(rename = "@stringValue2")]
    #[serde(default)]
    pub string_value2: Option<XmlString>,
    #[serde(rename = "autoFilter")]
    pub auto_filter: Box<AutoFilter>,
    #[serde(rename = "extLst")]
    #[serde(default)]
    pub extension_list: Option<Box<ExtensionList>>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct PivotArea {
    #[serde(rename = "@field")]
    #[serde(default)]
    pub field: Option<i32>,
    #[serde(rename = "@type")]
    #[serde(default)]
    pub r#type: Option<STPivotAreaType>,
    #[serde(rename = "@dataOnly")]
    #[serde(default)]
    pub data_only: Option<bool>,
    #[serde(rename = "@labelOnly")]
    #[serde(default)]
    pub label_only: Option<bool>,
    #[serde(rename = "@grandRow")]
    #[serde(default)]
    pub grand_row: Option<bool>,
    #[serde(rename = "@grandCol")]
    #[serde(default)]
    pub grand_col: Option<bool>,
    #[serde(rename = "@cacheIndex")]
    #[serde(default)]
    pub cache_index: Option<bool>,
    #[serde(rename = "@outline")]
    #[serde(default)]
    pub outline: Option<bool>,
    #[serde(rename = "@offset")]
    #[serde(default)]
    pub offset: Option<Reference>,
    #[serde(rename = "@collapsedLevelsAreSubtotals")]
    #[serde(default)]
    pub collapsed_levels_are_subtotals: Option<bool>,
    #[serde(rename = "@axis")]
    #[serde(default)]
    pub axis: Option<STAxis>,
    #[serde(rename = "@fieldPosition")]
    #[serde(default)]
    pub field_position: Option<u32>,
    #[serde(rename = "references")]
    #[serde(default)]
    pub references: Option<Box<CTPivotAreaReferences>>,
    #[serde(rename = "extLst")]
    #[serde(default)]
    pub extension_list: Option<Box<ExtensionList>>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTPivotAreaReferences {
    #[serde(rename = "@count")]
    #[serde(default)]
    pub count: Option<u32>,
    #[serde(rename = "reference")]
    #[serde(default)]
    pub reference: Vec<Box<CTPivotAreaReference>>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTPivotAreaReference {
    #[serde(rename = "@field")]
    #[serde(default)]
    pub field: Option<u32>,
    #[serde(rename = "@count")]
    #[serde(default)]
    pub count: Option<u32>,
    #[serde(rename = "@selected")]
    #[serde(default)]
    pub selected: Option<bool>,
    #[serde(rename = "@byPosition")]
    #[serde(default)]
    pub by_position: Option<bool>,
    #[serde(rename = "@relative")]
    #[serde(default)]
    pub relative: Option<bool>,
    #[serde(rename = "@defaultSubtotal")]
    #[serde(default)]
    pub default_subtotal: Option<bool>,
    #[serde(rename = "@sumSubtotal")]
    #[serde(default)]
    pub sum_subtotal: Option<bool>,
    #[serde(rename = "@countASubtotal")]
    #[serde(default)]
    pub count_a_subtotal: Option<bool>,
    #[serde(rename = "@avgSubtotal")]
    #[serde(default)]
    pub avg_subtotal: Option<bool>,
    #[serde(rename = "@maxSubtotal")]
    #[serde(default)]
    pub max_subtotal: Option<bool>,
    #[serde(rename = "@minSubtotal")]
    #[serde(default)]
    pub min_subtotal: Option<bool>,
    #[serde(rename = "@productSubtotal")]
    #[serde(default)]
    pub product_subtotal: Option<bool>,
    #[serde(rename = "@countSubtotal")]
    #[serde(default)]
    pub count_subtotal: Option<bool>,
    #[serde(rename = "@stdDevSubtotal")]
    #[serde(default)]
    pub std_dev_subtotal: Option<bool>,
    #[serde(rename = "@stdDevPSubtotal")]
    #[serde(default)]
    pub std_dev_p_subtotal: Option<bool>,
    #[serde(rename = "@varSubtotal")]
    #[serde(default)]
    pub var_subtotal: Option<bool>,
    #[serde(rename = "@varPSubtotal")]
    #[serde(default)]
    pub var_p_subtotal: Option<bool>,
    #[serde(rename = "x")]
    #[serde(default)]
    pub x: Vec<Box<CTIndex>>,
    #[serde(rename = "extLst")]
    #[serde(default)]
    pub extension_list: Option<Box<ExtensionList>>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTIndex {
    #[serde(rename = "@v")]
    pub value: u32,
}

pub type SmlQueryTable = Box<QueryTable>;

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct QueryTable {
    #[serde(rename = "@name")]
    pub name: XmlString,
    #[serde(rename = "@headers")]
    #[serde(default)]
    pub headers: Option<bool>,
    #[serde(rename = "@rowNumbers")]
    #[serde(default)]
    pub row_numbers: Option<bool>,
    #[serde(rename = "@disableRefresh")]
    #[serde(default)]
    pub disable_refresh: Option<bool>,
    #[serde(rename = "@backgroundRefresh")]
    #[serde(default)]
    pub background_refresh: Option<bool>,
    #[serde(rename = "@firstBackgroundRefresh")]
    #[serde(default)]
    pub first_background_refresh: Option<bool>,
    #[serde(rename = "@refreshOnLoad")]
    #[serde(default)]
    pub refresh_on_load: Option<bool>,
    #[serde(rename = "@growShrinkType")]
    #[serde(default)]
    pub grow_shrink_type: Option<STGrowShrinkType>,
    #[serde(rename = "@fillFormulas")]
    #[serde(default)]
    pub fill_formulas: Option<bool>,
    #[serde(rename = "@removeDataOnSave")]
    #[serde(default)]
    pub remove_data_on_save: Option<bool>,
    #[serde(rename = "@disableEdit")]
    #[serde(default)]
    pub disable_edit: Option<bool>,
    #[serde(rename = "@preserveFormatting")]
    #[serde(default)]
    pub preserve_formatting: Option<bool>,
    #[serde(rename = "@adjustColumnWidth")]
    #[serde(default)]
    pub adjust_column_width: Option<bool>,
    #[serde(rename = "@intermediate")]
    #[serde(default)]
    pub intermediate: Option<bool>,
    #[serde(rename = "@connectionId")]
    pub connection_id: u32,
    #[serde(rename = "queryTableRefresh")]
    #[serde(default)]
    pub query_table_refresh: Option<Box<QueryTableRefresh>>,
    #[serde(rename = "extLst")]
    #[serde(default)]
    pub extension_list: Option<Box<ExtensionList>>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct QueryTableRefresh {
    #[serde(rename = "@preserveSortFilterLayout")]
    #[serde(default)]
    pub preserve_sort_filter_layout: Option<bool>,
    #[serde(rename = "@fieldIdWrapped")]
    #[serde(default)]
    pub field_id_wrapped: Option<bool>,
    #[serde(rename = "@headersInLastRefresh")]
    #[serde(default)]
    pub headers_in_last_refresh: Option<bool>,
    #[serde(rename = "@minimumVersion")]
    #[serde(default)]
    pub minimum_version: Option<u8>,
    #[serde(rename = "@nextId")]
    #[serde(default)]
    pub next_id: Option<u32>,
    #[serde(rename = "@unboundColumnsLeft")]
    #[serde(default)]
    pub unbound_columns_left: Option<u32>,
    #[serde(rename = "@unboundColumnsRight")]
    #[serde(default)]
    pub unbound_columns_right: Option<u32>,
    #[serde(rename = "queryTableFields")]
    pub query_table_fields: Box<QueryTableFields>,
    #[serde(rename = "queryTableDeletedFields")]
    #[serde(default)]
    pub query_table_deleted_fields: Option<Box<QueryTableDeletedFields>>,
    #[serde(rename = "sortState")]
    #[serde(default)]
    pub sort_state: Option<Box<SortState>>,
    #[serde(rename = "extLst")]
    #[serde(default)]
    pub extension_list: Option<Box<ExtensionList>>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct QueryTableDeletedFields {
    #[serde(rename = "@count")]
    #[serde(default)]
    pub count: Option<u32>,
    #[serde(rename = "deletedField")]
    #[serde(default)]
    pub deleted_field: Vec<Box<CTDeletedField>>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTDeletedField {
    #[serde(rename = "@name")]
    pub name: XmlString,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct QueryTableFields {
    #[serde(rename = "@count")]
    #[serde(default)]
    pub count: Option<u32>,
    #[serde(rename = "queryTableField")]
    #[serde(default)]
    pub query_table_field: Vec<Box<QueryTableField>>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct QueryTableField {
    #[serde(rename = "@id")]
    pub id: u32,
    #[serde(rename = "@name")]
    #[serde(default)]
    pub name: Option<XmlString>,
    #[serde(rename = "@dataBound")]
    #[serde(default)]
    pub data_bound: Option<bool>,
    #[serde(rename = "@rowNumbers")]
    #[serde(default)]
    pub row_numbers: Option<bool>,
    #[serde(rename = "@fillFormulas")]
    #[serde(default)]
    pub fill_formulas: Option<bool>,
    #[serde(rename = "@clipped")]
    #[serde(default)]
    pub clipped: Option<bool>,
    #[serde(rename = "@tableColumnId")]
    #[serde(default)]
    pub table_column_id: Option<u32>,
    #[serde(rename = "extLst")]
    #[serde(default)]
    pub extension_list: Option<Box<ExtensionList>>,
}

pub type SmlSst = Box<SharedStrings>;

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct SharedStrings {
    #[serde(rename = "@count")]
    #[serde(default)]
    pub count: Option<u32>,
    #[serde(rename = "@uniqueCount")]
    #[serde(default)]
    pub unique_count: Option<u32>,
    #[serde(rename = "si")]
    #[serde(default)]
    pub si: Vec<Box<RichString>>,
    #[serde(rename = "extLst")]
    #[serde(default)]
    pub extension_list: Option<Box<ExtensionList>>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct PhoneticRun {
    #[serde(rename = "@sb")]
    pub sb: u32,
    #[serde(rename = "@eb")]
    pub eb: u32,
    #[serde(rename = "t")]
    pub cell_type: XmlString,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct RichTextElement {
    #[serde(rename = "rPr")]
    #[serde(default)]
    pub r_pr: Option<Box<RichTextRunProperties>>,
    #[serde(rename = "t")]
    pub cell_type: XmlString,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct RichTextRunProperties;

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct RichString {
    #[serde(rename = "t")]
    #[serde(default)]
    pub cell_type: Option<XmlString>,
    #[serde(rename = "r")]
    #[serde(default)]
    pub reference: Vec<Box<RichTextElement>>,
    #[serde(rename = "rPh")]
    #[serde(default)]
    pub r_ph: Vec<Box<PhoneticRun>>,
    #[serde(rename = "phoneticPr")]
    #[serde(default)]
    pub phonetic_pr: Option<Box<PhoneticProperties>>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct PhoneticProperties {
    #[serde(rename = "@fontId")]
    pub font_id: STFontId,
    #[serde(rename = "@type")]
    #[serde(default)]
    pub r#type: Option<STPhoneticType>,
    #[serde(rename = "@alignment")]
    #[serde(default)]
    pub alignment: Option<STPhoneticAlignment>,
}

pub type SmlHeaders = Box<RevisionHeaders>;

pub type SmlRevisions = Box<Revisions>;

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct RevisionHeaders {
    #[serde(rename = "@guid")]
    pub guid: Guid,
    #[serde(rename = "@lastGuid")]
    #[serde(default)]
    pub last_guid: Option<Guid>,
    #[serde(rename = "@shared")]
    #[serde(default)]
    pub shared: Option<bool>,
    #[serde(rename = "@diskRevisions")]
    #[serde(default)]
    pub disk_revisions: Option<bool>,
    #[serde(rename = "@history")]
    #[serde(default)]
    pub history: Option<bool>,
    #[serde(rename = "@trackRevisions")]
    #[serde(default)]
    pub track_revisions: Option<bool>,
    #[serde(rename = "@exclusive")]
    #[serde(default)]
    pub exclusive: Option<bool>,
    #[serde(rename = "@revisionId")]
    #[serde(default)]
    pub revision_id: Option<u32>,
    #[serde(rename = "@version")]
    #[serde(default)]
    pub version: Option<i32>,
    #[serde(rename = "@keepChangeHistory")]
    #[serde(default)]
    pub keep_change_history: Option<bool>,
    #[serde(rename = "@protected")]
    #[serde(default)]
    pub protected: Option<bool>,
    #[serde(rename = "@preserveHistory")]
    #[serde(default)]
    pub preserve_history: Option<u32>,
    #[serde(rename = "header")]
    #[serde(default)]
    pub header: Vec<Box<RevisionHeader>>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct Revisions;

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct SmlAGRevData {
    #[serde(rename = "@rId")]
    pub r_id: u32,
    #[serde(rename = "@ua")]
    #[serde(default)]
    pub ua: Option<bool>,
    #[serde(rename = "@ra")]
    #[serde(default)]
    pub ra: Option<bool>,
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
    #[serde(rename = "@minRId")]
    #[serde(default)]
    pub min_r_id: Option<u32>,
    #[serde(rename = "@maxRId")]
    #[serde(default)]
    pub max_r_id: Option<u32>,
    #[serde(rename = "sheetIdMap")]
    pub sheet_id_map: Box<CTSheetIdMap>,
    #[serde(rename = "reviewedList")]
    #[serde(default)]
    pub reviewed_list: Option<Box<ReviewedRevisions>>,
    #[serde(rename = "extLst")]
    #[serde(default)]
    pub extension_list: Option<Box<ExtensionList>>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTSheetIdMap {
    #[serde(rename = "@count")]
    #[serde(default)]
    pub count: Option<u32>,
    #[serde(rename = "sheetId")]
    #[serde(default)]
    pub sheet_id: Vec<Box<CTSheetId>>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTSheetId {
    #[serde(rename = "@val")]
    pub value: u32,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct ReviewedRevisions {
    #[serde(rename = "@count")]
    #[serde(default)]
    pub count: Option<u32>,
    #[serde(rename = "reviewed")]
    #[serde(default)]
    pub reviewed: Vec<Box<Reviewed>>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct Reviewed {
    #[serde(rename = "@rId")]
    pub r_id: u32,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct UndoInfo {
    #[serde(rename = "@index")]
    pub index: u32,
    #[serde(rename = "@exp")]
    pub exp: STFormulaExpression,
    #[serde(rename = "@ref3D")]
    #[serde(default)]
    pub ref3_d: Option<bool>,
    #[serde(rename = "@array")]
    #[serde(default)]
    pub array: Option<bool>,
    #[serde(rename = "@v")]
    #[serde(default)]
    pub value: Option<bool>,
    #[serde(rename = "@nf")]
    #[serde(default)]
    pub nf: Option<bool>,
    #[serde(rename = "@cs")]
    #[serde(default)]
    pub cs: Option<bool>,
    #[serde(rename = "@dr")]
    pub dr: STRefA,
    #[serde(rename = "@dn")]
    #[serde(default)]
    pub dn: Option<XmlString>,
    #[serde(rename = "@r")]
    #[serde(default)]
    pub reference: Option<CellRef>,
    #[serde(rename = "@sId")]
    #[serde(default)]
    pub s_id: Option<u32>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct RevisionRowColumn {
    #[serde(rename = "@sId")]
    pub s_id: u32,
    #[serde(rename = "@eol")]
    #[serde(default)]
    pub eol: Option<bool>,
    #[serde(rename = "@ref")]
    pub reference: Reference,
    #[serde(rename = "@action")]
    pub action: STRwColActionType,
    #[serde(rename = "@edge")]
    #[serde(default)]
    pub edge: Option<bool>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct RevisionMove {
    #[serde(rename = "@sheetId")]
    pub sheet_id: u32,
    #[serde(rename = "@source")]
    pub source: Reference,
    #[serde(rename = "@destination")]
    pub destination: Reference,
    #[serde(rename = "@sourceSheetId")]
    #[serde(default)]
    pub source_sheet_id: Option<u32>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct RevisionCustomView {
    #[serde(rename = "@guid")]
    pub guid: Guid,
    #[serde(rename = "@action")]
    pub action: STRevisionAction,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct RevisionSheetRename {
    #[serde(rename = "@sheetId")]
    pub sheet_id: u32,
    #[serde(rename = "@oldName")]
    pub old_name: XmlString,
    #[serde(rename = "@newName")]
    pub new_name: XmlString,
    #[serde(rename = "extLst")]
    #[serde(default)]
    pub extension_list: Option<Box<ExtensionList>>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct RevisionInsertSheet {
    #[serde(rename = "@sheetId")]
    pub sheet_id: u32,
    #[serde(rename = "@name")]
    pub name: XmlString,
    #[serde(rename = "@sheetPosition")]
    pub sheet_position: u32,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct RevisionCellChange {
    #[serde(rename = "@sId")]
    pub s_id: u32,
    #[serde(rename = "@odxf")]
    #[serde(default)]
    pub odxf: Option<bool>,
    #[serde(rename = "@xfDxf")]
    #[serde(default)]
    pub xf_dxf: Option<bool>,
    #[serde(rename = "@s")]
    #[serde(default)]
    pub style_index: Option<bool>,
    #[serde(rename = "@dxf")]
    #[serde(default)]
    pub dxf: Option<bool>,
    #[serde(rename = "@numFmtId")]
    #[serde(default)]
    pub number_format_id: Option<STNumFmtId>,
    #[serde(rename = "@quotePrefix")]
    #[serde(default)]
    pub quote_prefix: Option<bool>,
    #[serde(rename = "@oldQuotePrefix")]
    #[serde(default)]
    pub old_quote_prefix: Option<bool>,
    #[serde(rename = "@ph")]
    #[serde(default)]
    pub placeholder: Option<bool>,
    #[serde(rename = "@oldPh")]
    #[serde(default)]
    pub old_ph: Option<bool>,
    #[serde(rename = "@endOfListFormulaUpdate")]
    #[serde(default)]
    pub end_of_list_formula_update: Option<bool>,
    #[serde(rename = "oc")]
    #[serde(default)]
    pub oc: Option<Box<Cell>>,
    #[serde(rename = "nc")]
    pub nc: Box<Cell>,
    #[serde(rename = "ndxf")]
    #[serde(default)]
    pub ndxf: Option<Box<DifferentialFormat>>,
    #[serde(rename = "extLst")]
    #[serde(default)]
    pub extension_list: Option<Box<ExtensionList>>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct RevisionFormatting {
    #[serde(rename = "@sheetId")]
    pub sheet_id: u32,
    #[serde(rename = "@xfDxf")]
    #[serde(default)]
    pub xf_dxf: Option<bool>,
    #[serde(rename = "@s")]
    #[serde(default)]
    pub style_index: Option<bool>,
    #[serde(rename = "@sqref")]
    pub square_reference: SquareRef,
    #[serde(rename = "@start")]
    #[serde(default)]
    pub start: Option<u32>,
    #[serde(rename = "@length")]
    #[serde(default)]
    pub length: Option<u32>,
    #[serde(rename = "dxf")]
    #[serde(default)]
    pub dxf: Option<Box<DifferentialFormat>>,
    #[serde(rename = "extLst")]
    #[serde(default)]
    pub extension_list: Option<Box<ExtensionList>>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct RevisionAutoFormatting {
    #[serde(rename = "@sheetId")]
    pub sheet_id: u32,
    #[serde(rename = "@ref")]
    pub reference: Reference,
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
    #[serde(default)]
    pub action: Option<STRevisionAction>,
    #[serde(rename = "@alwaysShow")]
    #[serde(default)]
    pub always_show: Option<bool>,
    #[serde(rename = "@old")]
    #[serde(default)]
    pub old: Option<bool>,
    #[serde(rename = "@hiddenRow")]
    #[serde(default)]
    pub hidden_row: Option<bool>,
    #[serde(rename = "@hiddenColumn")]
    #[serde(default)]
    pub hidden_column: Option<bool>,
    #[serde(rename = "@author")]
    pub author: XmlString,
    #[serde(rename = "@oldLength")]
    #[serde(default)]
    pub old_length: Option<u32>,
    #[serde(rename = "@newLength")]
    #[serde(default)]
    pub new_length: Option<u32>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct RevisionDefinedName {
    #[serde(rename = "@localSheetId")]
    #[serde(default)]
    pub local_sheet_id: Option<u32>,
    #[serde(rename = "@customView")]
    #[serde(default)]
    pub custom_view: Option<bool>,
    #[serde(rename = "@name")]
    pub name: XmlString,
    #[serde(rename = "@function")]
    #[serde(default)]
    pub function: Option<bool>,
    #[serde(rename = "@oldFunction")]
    #[serde(default)]
    pub old_function: Option<bool>,
    #[serde(rename = "@functionGroupId")]
    #[serde(default)]
    pub function_group_id: Option<u8>,
    #[serde(rename = "@oldFunctionGroupId")]
    #[serde(default)]
    pub old_function_group_id: Option<u8>,
    #[serde(rename = "@shortcutKey")]
    #[serde(default)]
    pub shortcut_key: Option<u8>,
    #[serde(rename = "@oldShortcutKey")]
    #[serde(default)]
    pub old_shortcut_key: Option<u8>,
    #[serde(rename = "@hidden")]
    #[serde(default)]
    pub hidden: Option<bool>,
    #[serde(rename = "@oldHidden")]
    #[serde(default)]
    pub old_hidden: Option<bool>,
    #[serde(rename = "@customMenu")]
    #[serde(default)]
    pub custom_menu: Option<XmlString>,
    #[serde(rename = "@oldCustomMenu")]
    #[serde(default)]
    pub old_custom_menu: Option<XmlString>,
    #[serde(rename = "@description")]
    #[serde(default)]
    pub description: Option<XmlString>,
    #[serde(rename = "@oldDescription")]
    #[serde(default)]
    pub old_description: Option<XmlString>,
    #[serde(rename = "@help")]
    #[serde(default)]
    pub help: Option<XmlString>,
    #[serde(rename = "@oldHelp")]
    #[serde(default)]
    pub old_help: Option<XmlString>,
    #[serde(rename = "@statusBar")]
    #[serde(default)]
    pub status_bar: Option<XmlString>,
    #[serde(rename = "@oldStatusBar")]
    #[serde(default)]
    pub old_status_bar: Option<XmlString>,
    #[serde(rename = "@comment")]
    #[serde(default)]
    pub comment: Option<XmlString>,
    #[serde(rename = "@oldComment")]
    #[serde(default)]
    pub old_comment: Option<XmlString>,
    #[serde(rename = "formula")]
    #[serde(default)]
    pub formula: Option<STFormula>,
    #[serde(rename = "oldFormula")]
    #[serde(default)]
    pub old_formula: Option<STFormula>,
    #[serde(rename = "extLst")]
    #[serde(default)]
    pub extension_list: Option<Box<ExtensionList>>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct RevisionConflict {
    #[serde(rename = "@sheetId")]
    #[serde(default)]
    pub sheet_id: Option<u32>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct RevisionQueryTableField {
    #[serde(rename = "@sheetId")]
    pub sheet_id: u32,
    #[serde(rename = "@ref")]
    pub reference: Reference,
    #[serde(rename = "@fieldId")]
    pub field_id: u32,
}

pub type SmlUsers = Box<Users>;

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct Users {
    #[serde(rename = "@count")]
    #[serde(default)]
    pub count: Option<u32>,
    #[serde(rename = "userInfo")]
    #[serde(default)]
    pub user_info: Vec<Box<SharedUser>>,
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
    #[serde(default)]
    pub extension_list: Option<Box<ExtensionList>>,
}

pub type SmlWorksheet = Box<Worksheet>;

pub type SmlChartsheet = Box<Chartsheet>;

pub type SmlDialogsheet = Box<CTDialogsheet>;

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTMacrosheet {
    #[serde(rename = "sheetPr")]
    #[serde(default)]
    pub sheet_properties: Option<Box<SheetProperties>>,
    #[serde(rename = "dimension")]
    #[serde(default)]
    pub dimension: Option<Box<SheetDimension>>,
    #[serde(rename = "sheetViews")]
    #[serde(default)]
    pub sheet_views: Option<Box<SheetViews>>,
    #[serde(rename = "sheetFormatPr")]
    #[serde(default)]
    pub sheet_format: Option<Box<SheetFormat>>,
    #[serde(rename = "cols")]
    #[serde(default)]
    pub cols: Vec<Box<Columns>>,
    #[serde(rename = "sheetData")]
    pub sheet_data: Box<SheetData>,
    #[serde(rename = "sheetProtection")]
    #[serde(default)]
    pub sheet_protection: Option<Box<SheetProtection>>,
    #[serde(rename = "autoFilter")]
    #[serde(default)]
    pub auto_filter: Option<Box<AutoFilter>>,
    #[serde(rename = "sortState")]
    #[serde(default)]
    pub sort_state: Option<Box<SortState>>,
    #[serde(rename = "dataConsolidate")]
    #[serde(default)]
    pub data_consolidate: Option<Box<CTDataConsolidate>>,
    #[serde(rename = "customSheetViews")]
    #[serde(default)]
    pub custom_sheet_views: Option<Box<CustomSheetViews>>,
    #[serde(rename = "phoneticPr")]
    #[serde(default)]
    pub phonetic_pr: Option<Box<PhoneticProperties>>,
    #[serde(rename = "conditionalFormatting")]
    #[serde(default)]
    pub conditional_formatting: Vec<Box<ConditionalFormatting>>,
    #[serde(rename = "printOptions")]
    #[serde(default)]
    pub print_options: Option<Box<PrintOptions>>,
    #[serde(rename = "pageMargins")]
    #[serde(default)]
    pub page_margins: Option<Box<PageMargins>>,
    #[serde(rename = "pageSetup")]
    #[serde(default)]
    pub page_setup: Option<Box<PageSetup>>,
    #[serde(rename = "headerFooter")]
    #[serde(default)]
    pub header_footer: Option<Box<HeaderFooter>>,
    #[serde(rename = "rowBreaks")]
    #[serde(default)]
    pub row_breaks: Option<Box<PageBreaks>>,
    #[serde(rename = "colBreaks")]
    #[serde(default)]
    pub col_breaks: Option<Box<PageBreaks>>,
    #[serde(rename = "customProperties")]
    #[serde(default)]
    pub custom_properties: Option<Box<CTCustomProperties>>,
    #[serde(rename = "drawing")]
    #[serde(default)]
    pub drawing: Option<Box<Drawing>>,
    #[serde(rename = "legacyDrawing")]
    #[serde(default)]
    pub legacy_drawing: Option<Box<LegacyDrawing>>,
    #[serde(rename = "legacyDrawingHF")]
    #[serde(default)]
    pub legacy_drawing_h_f: Option<Box<LegacyDrawing>>,
    #[serde(rename = "drawingHF")]
    #[serde(default)]
    pub drawing_h_f: Option<Box<DrawingHeaderFooter>>,
    #[serde(rename = "picture")]
    #[serde(default)]
    pub picture: Option<Box<SheetBackgroundPicture>>,
    #[serde(rename = "oleObjects")]
    #[serde(default)]
    pub ole_objects: Option<Box<OleObjects>>,
    #[serde(rename = "extLst")]
    #[serde(default)]
    pub extension_list: Option<Box<ExtensionList>>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTDialogsheet {
    #[serde(rename = "sheetPr")]
    #[serde(default)]
    pub sheet_properties: Option<Box<SheetProperties>>,
    #[serde(rename = "sheetViews")]
    #[serde(default)]
    pub sheet_views: Option<Box<SheetViews>>,
    #[serde(rename = "sheetFormatPr")]
    #[serde(default)]
    pub sheet_format: Option<Box<SheetFormat>>,
    #[serde(rename = "sheetProtection")]
    #[serde(default)]
    pub sheet_protection: Option<Box<SheetProtection>>,
    #[serde(rename = "customSheetViews")]
    #[serde(default)]
    pub custom_sheet_views: Option<Box<CustomSheetViews>>,
    #[serde(rename = "printOptions")]
    #[serde(default)]
    pub print_options: Option<Box<PrintOptions>>,
    #[serde(rename = "pageMargins")]
    #[serde(default)]
    pub page_margins: Option<Box<PageMargins>>,
    #[serde(rename = "pageSetup")]
    #[serde(default)]
    pub page_setup: Option<Box<PageSetup>>,
    #[serde(rename = "headerFooter")]
    #[serde(default)]
    pub header_footer: Option<Box<HeaderFooter>>,
    #[serde(rename = "drawing")]
    #[serde(default)]
    pub drawing: Option<Box<Drawing>>,
    #[serde(rename = "legacyDrawing")]
    #[serde(default)]
    pub legacy_drawing: Option<Box<LegacyDrawing>>,
    #[serde(rename = "legacyDrawingHF")]
    #[serde(default)]
    pub legacy_drawing_h_f: Option<Box<LegacyDrawing>>,
    #[serde(rename = "drawingHF")]
    #[serde(default)]
    pub drawing_h_f: Option<Box<DrawingHeaderFooter>>,
    #[serde(rename = "oleObjects")]
    #[serde(default)]
    pub ole_objects: Option<Box<OleObjects>>,
    #[serde(rename = "controls")]
    #[serde(default)]
    pub controls: Option<Box<Controls>>,
    #[serde(rename = "extLst")]
    #[serde(default)]
    pub extension_list: Option<Box<ExtensionList>>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct Worksheet {
    #[serde(rename = "sheetPr")]
    #[serde(default)]
    pub sheet_properties: Option<Box<SheetProperties>>,
    #[serde(rename = "dimension")]
    #[serde(default)]
    pub dimension: Option<Box<SheetDimension>>,
    #[serde(rename = "sheetViews")]
    #[serde(default)]
    pub sheet_views: Option<Box<SheetViews>>,
    #[serde(rename = "sheetFormatPr")]
    #[serde(default)]
    pub sheet_format: Option<Box<SheetFormat>>,
    #[serde(rename = "cols")]
    #[serde(default)]
    pub cols: Vec<Box<Columns>>,
    #[serde(rename = "sheetData")]
    pub sheet_data: Box<SheetData>,
    #[serde(rename = "sheetCalcPr")]
    #[serde(default)]
    pub sheet_calc_pr: Option<Box<SheetCalcProperties>>,
    #[serde(rename = "sheetProtection")]
    #[serde(default)]
    pub sheet_protection: Option<Box<SheetProtection>>,
    #[serde(rename = "protectedRanges")]
    #[serde(default)]
    pub protected_ranges: Option<Box<ProtectedRanges>>,
    #[serde(rename = "scenarios")]
    #[serde(default)]
    pub scenarios: Option<Box<Scenarios>>,
    #[serde(rename = "autoFilter")]
    #[serde(default)]
    pub auto_filter: Option<Box<AutoFilter>>,
    #[serde(rename = "sortState")]
    #[serde(default)]
    pub sort_state: Option<Box<SortState>>,
    #[serde(rename = "dataConsolidate")]
    #[serde(default)]
    pub data_consolidate: Option<Box<CTDataConsolidate>>,
    #[serde(rename = "customSheetViews")]
    #[serde(default)]
    pub custom_sheet_views: Option<Box<CustomSheetViews>>,
    #[serde(rename = "mergeCells")]
    #[serde(default)]
    pub merged_cells: Option<Box<MergedCells>>,
    #[serde(rename = "phoneticPr")]
    #[serde(default)]
    pub phonetic_pr: Option<Box<PhoneticProperties>>,
    #[serde(rename = "conditionalFormatting")]
    #[serde(default)]
    pub conditional_formatting: Vec<Box<ConditionalFormatting>>,
    #[serde(rename = "dataValidations")]
    #[serde(default)]
    pub data_validations: Option<Box<DataValidations>>,
    #[serde(rename = "hyperlinks")]
    #[serde(default)]
    pub hyperlinks: Option<Box<Hyperlinks>>,
    #[serde(rename = "printOptions")]
    #[serde(default)]
    pub print_options: Option<Box<PrintOptions>>,
    #[serde(rename = "pageMargins")]
    #[serde(default)]
    pub page_margins: Option<Box<PageMargins>>,
    #[serde(rename = "pageSetup")]
    #[serde(default)]
    pub page_setup: Option<Box<PageSetup>>,
    #[serde(rename = "headerFooter")]
    #[serde(default)]
    pub header_footer: Option<Box<HeaderFooter>>,
    #[serde(rename = "rowBreaks")]
    #[serde(default)]
    pub row_breaks: Option<Box<PageBreaks>>,
    #[serde(rename = "colBreaks")]
    #[serde(default)]
    pub col_breaks: Option<Box<PageBreaks>>,
    #[serde(rename = "customProperties")]
    #[serde(default)]
    pub custom_properties: Option<Box<CTCustomProperties>>,
    #[serde(rename = "cellWatches")]
    #[serde(default)]
    pub cell_watches: Option<Box<CellWatches>>,
    #[serde(rename = "ignoredErrors")]
    #[serde(default)]
    pub ignored_errors: Option<Box<IgnoredErrors>>,
    #[serde(rename = "smartTags")]
    #[serde(default)]
    pub smart_tags: Option<Box<SmartTags>>,
    #[serde(rename = "drawing")]
    #[serde(default)]
    pub drawing: Option<Box<Drawing>>,
    #[serde(rename = "legacyDrawing")]
    #[serde(default)]
    pub legacy_drawing: Option<Box<LegacyDrawing>>,
    #[serde(rename = "legacyDrawingHF")]
    #[serde(default)]
    pub legacy_drawing_h_f: Option<Box<LegacyDrawing>>,
    #[serde(rename = "drawingHF")]
    #[serde(default)]
    pub drawing_h_f: Option<Box<DrawingHeaderFooter>>,
    #[serde(rename = "picture")]
    #[serde(default)]
    pub picture: Option<Box<SheetBackgroundPicture>>,
    #[serde(rename = "oleObjects")]
    #[serde(default)]
    pub ole_objects: Option<Box<OleObjects>>,
    #[serde(rename = "controls")]
    #[serde(default)]
    pub controls: Option<Box<Controls>>,
    #[serde(rename = "webPublishItems")]
    #[serde(default)]
    pub web_publish_items: Option<Box<WebPublishItems>>,
    #[serde(rename = "tableParts")]
    #[serde(default)]
    pub table_parts: Option<Box<TableParts>>,
    #[serde(rename = "extLst")]
    #[serde(default)]
    pub extension_list: Option<Box<ExtensionList>>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct SheetData {
    #[serde(rename = "row")]
    #[serde(default)]
    pub row: Vec<Box<Row>>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct SheetCalcProperties {
    #[serde(rename = "@fullCalcOnLoad")]
    #[serde(default)]
    pub full_calc_on_load: Option<bool>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct SheetFormat {
    #[serde(rename = "@baseColWidth")]
    #[serde(default)]
    pub base_col_width: Option<u32>,
    #[serde(rename = "@defaultColWidth")]
    #[serde(default)]
    pub default_col_width: Option<f64>,
    #[serde(rename = "@defaultRowHeight")]
    pub default_row_height: f64,
    #[serde(rename = "@customHeight")]
    #[serde(default)]
    pub custom_height: Option<bool>,
    #[serde(rename = "@zeroHeight")]
    #[serde(default)]
    pub zero_height: Option<bool>,
    #[serde(rename = "@thickTop")]
    #[serde(default)]
    pub thick_top: Option<bool>,
    #[serde(rename = "@thickBottom")]
    #[serde(default)]
    pub thick_bottom: Option<bool>,
    #[serde(rename = "@outlineLevelRow")]
    #[serde(default)]
    pub outline_level_row: Option<u8>,
    #[serde(rename = "@outlineLevelCol")]
    #[serde(default)]
    pub outline_level_col: Option<u8>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct Columns {
    #[serde(rename = "col")]
    #[serde(default)]
    pub col: Vec<Box<Column>>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct Column {
    #[serde(rename = "@min")]
    pub start_column: u32,
    #[serde(rename = "@max")]
    pub end_column: u32,
    #[serde(rename = "@width")]
    #[serde(default)]
    pub width: Option<f64>,
    #[serde(rename = "@style")]
    #[serde(default)]
    pub style: Option<u32>,
    #[serde(rename = "@hidden")]
    #[serde(default)]
    pub hidden: Option<bool>,
    #[serde(rename = "@bestFit")]
    #[serde(default)]
    pub best_fit: Option<bool>,
    #[serde(rename = "@customWidth")]
    #[serde(default)]
    pub custom_width: Option<bool>,
    #[serde(rename = "@phonetic")]
    #[serde(default)]
    pub phonetic: Option<bool>,
    #[serde(rename = "@outlineLevel")]
    #[serde(default)]
    pub outline_level: Option<u8>,
    #[serde(rename = "@collapsed")]
    #[serde(default)]
    pub collapsed: Option<bool>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct Row {
    #[serde(rename = "@r")]
    #[serde(default)]
    pub reference: Option<u32>,
    #[serde(rename = "@spans")]
    #[serde(default)]
    pub cell_spans: Option<CellSpans>,
    #[serde(rename = "@s")]
    #[serde(default)]
    pub style_index: Option<u32>,
    #[serde(rename = "@customFormat")]
    #[serde(default)]
    pub custom_format: Option<bool>,
    #[serde(rename = "@ht")]
    #[serde(default)]
    pub height: Option<f64>,
    #[serde(rename = "@hidden")]
    #[serde(default)]
    pub hidden: Option<bool>,
    #[serde(rename = "@customHeight")]
    #[serde(default)]
    pub custom_height: Option<bool>,
    #[serde(rename = "@outlineLevel")]
    #[serde(default)]
    pub outline_level: Option<u8>,
    #[serde(rename = "@collapsed")]
    #[serde(default)]
    pub collapsed: Option<bool>,
    #[serde(rename = "@thickTop")]
    #[serde(default)]
    pub thick_top: Option<bool>,
    #[serde(rename = "@thickBot")]
    #[serde(default)]
    pub thick_bot: Option<bool>,
    #[serde(rename = "@ph")]
    #[serde(default)]
    pub placeholder: Option<bool>,
    #[serde(rename = "c")]
    #[serde(default)]
    pub cells: Vec<Box<Cell>>,
    #[serde(rename = "extLst")]
    #[serde(default)]
    pub extension_list: Option<Box<ExtensionList>>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct Cell {
    #[serde(rename = "@r")]
    #[serde(default)]
    pub reference: Option<CellRef>,
    #[serde(rename = "@s")]
    #[serde(default)]
    pub style_index: Option<u32>,
    #[serde(rename = "@t")]
    #[serde(default)]
    pub cell_type: Option<CellType>,
    #[serde(rename = "@cm")]
    #[serde(default)]
    pub cm: Option<u32>,
    #[serde(rename = "@vm")]
    #[serde(default)]
    pub vm: Option<u32>,
    #[serde(rename = "@ph")]
    #[serde(default)]
    pub placeholder: Option<bool>,
    #[serde(rename = "f")]
    #[serde(default)]
    pub formula: Option<Box<CellFormula>>,
    #[serde(rename = "v")]
    #[serde(default)]
    pub value: Option<XmlString>,
    #[serde(rename = "is")]
    #[serde(default)]
    pub is: Option<Box<RichString>>,
    #[serde(rename = "extLst")]
    #[serde(default)]
    pub extension_list: Option<Box<ExtensionList>>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct SheetProperties {
    #[serde(rename = "@syncHorizontal")]
    #[serde(default)]
    pub sync_horizontal: Option<bool>,
    #[serde(rename = "@syncVertical")]
    #[serde(default)]
    pub sync_vertical: Option<bool>,
    #[serde(rename = "@syncRef")]
    #[serde(default)]
    pub sync_ref: Option<Reference>,
    #[serde(rename = "@transitionEvaluation")]
    #[serde(default)]
    pub transition_evaluation: Option<bool>,
    #[serde(rename = "@transitionEntry")]
    #[serde(default)]
    pub transition_entry: Option<bool>,
    #[serde(rename = "@published")]
    #[serde(default)]
    pub published: Option<bool>,
    #[serde(rename = "@codeName")]
    #[serde(default)]
    pub code_name: Option<String>,
    #[serde(rename = "@filterMode")]
    #[serde(default)]
    pub filter_mode: Option<bool>,
    #[serde(rename = "@enableFormatConditionsCalculation")]
    #[serde(default)]
    pub enable_format_conditions_calculation: Option<bool>,
    #[serde(rename = "tabColor")]
    #[serde(default)]
    pub tab_color: Option<Box<Color>>,
    #[serde(rename = "outlinePr")]
    #[serde(default)]
    pub outline_pr: Option<Box<OutlineProperties>>,
    #[serde(rename = "pageSetUpPr")]
    #[serde(default)]
    pub page_set_up_pr: Option<Box<PageSetupProperties>>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct SheetDimension {
    #[serde(rename = "@ref")]
    pub reference: Reference,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct SheetViews {
    #[serde(rename = "sheetView")]
    #[serde(default)]
    pub sheet_view: Vec<Box<SheetView>>,
    #[serde(rename = "extLst")]
    #[serde(default)]
    pub extension_list: Option<Box<ExtensionList>>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct SheetView {
    #[serde(rename = "@windowProtection")]
    #[serde(default)]
    pub window_protection: Option<bool>,
    #[serde(rename = "@showFormulas")]
    #[serde(default)]
    pub show_formulas: Option<bool>,
    #[serde(rename = "@showGridLines")]
    #[serde(default)]
    pub show_grid_lines: Option<bool>,
    #[serde(rename = "@showRowColHeaders")]
    #[serde(default)]
    pub show_row_col_headers: Option<bool>,
    #[serde(rename = "@showZeros")]
    #[serde(default)]
    pub show_zeros: Option<bool>,
    #[serde(rename = "@rightToLeft")]
    #[serde(default)]
    pub right_to_left: Option<bool>,
    #[serde(rename = "@tabSelected")]
    #[serde(default)]
    pub tab_selected: Option<bool>,
    #[serde(rename = "@showRuler")]
    #[serde(default)]
    pub show_ruler: Option<bool>,
    #[serde(rename = "@showOutlineSymbols")]
    #[serde(default)]
    pub show_outline_symbols: Option<bool>,
    #[serde(rename = "@defaultGridColor")]
    #[serde(default)]
    pub default_grid_color: Option<bool>,
    #[serde(rename = "@showWhiteSpace")]
    #[serde(default)]
    pub show_white_space: Option<bool>,
    #[serde(rename = "@view")]
    #[serde(default)]
    pub view: Option<SheetViewType>,
    #[serde(rename = "@topLeftCell")]
    #[serde(default)]
    pub top_left_cell: Option<CellRef>,
    #[serde(rename = "@colorId")]
    #[serde(default)]
    pub color_id: Option<u32>,
    #[serde(rename = "@zoomScale")]
    #[serde(default)]
    pub zoom_scale: Option<u32>,
    #[serde(rename = "@zoomScaleNormal")]
    #[serde(default)]
    pub zoom_scale_normal: Option<u32>,
    #[serde(rename = "@zoomScaleSheetLayoutView")]
    #[serde(default)]
    pub zoom_scale_sheet_layout_view: Option<u32>,
    #[serde(rename = "@zoomScalePageLayoutView")]
    #[serde(default)]
    pub zoom_scale_page_layout_view: Option<u32>,
    #[serde(rename = "@workbookViewId")]
    pub workbook_view_id: u32,
    #[serde(rename = "pane")]
    #[serde(default)]
    pub pane: Option<Box<Pane>>,
    #[serde(rename = "selection")]
    #[serde(default)]
    pub selection: Vec<Box<Selection>>,
    #[serde(rename = "pivotSelection")]
    #[serde(default)]
    pub pivot_selection: Vec<Box<CTPivotSelection>>,
    #[serde(rename = "extLst")]
    #[serde(default)]
    pub extension_list: Option<Box<ExtensionList>>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct Pane {
    #[serde(rename = "@xSplit")]
    #[serde(default)]
    pub x_split: Option<f64>,
    #[serde(rename = "@ySplit")]
    #[serde(default)]
    pub y_split: Option<f64>,
    #[serde(rename = "@topLeftCell")]
    #[serde(default)]
    pub top_left_cell: Option<CellRef>,
    #[serde(rename = "@activePane")]
    #[serde(default)]
    pub active_pane: Option<PaneType>,
    #[serde(rename = "@state")]
    #[serde(default)]
    pub state: Option<PaneState>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTPivotSelection {
    #[serde(rename = "@pane")]
    #[serde(default)]
    pub pane: Option<PaneType>,
    #[serde(rename = "@showHeader")]
    #[serde(default)]
    pub show_header: Option<bool>,
    #[serde(rename = "@label")]
    #[serde(default)]
    pub label: Option<bool>,
    #[serde(rename = "@data")]
    #[serde(default)]
    pub data: Option<bool>,
    #[serde(rename = "@extendable")]
    #[serde(default)]
    pub extendable: Option<bool>,
    #[serde(rename = "@count")]
    #[serde(default)]
    pub count: Option<u32>,
    #[serde(rename = "@axis")]
    #[serde(default)]
    pub axis: Option<STAxis>,
    #[serde(rename = "@dimension")]
    #[serde(default)]
    pub dimension: Option<u32>,
    #[serde(rename = "@start")]
    #[serde(default)]
    pub start: Option<u32>,
    #[serde(rename = "@min")]
    #[serde(default)]
    pub start_column: Option<u32>,
    #[serde(rename = "@max")]
    #[serde(default)]
    pub end_column: Option<u32>,
    #[serde(rename = "@activeRow")]
    #[serde(default)]
    pub active_row: Option<u32>,
    #[serde(rename = "@activeCol")]
    #[serde(default)]
    pub active_col: Option<u32>,
    #[serde(rename = "@previousRow")]
    #[serde(default)]
    pub previous_row: Option<u32>,
    #[serde(rename = "@previousCol")]
    #[serde(default)]
    pub previous_col: Option<u32>,
    #[serde(rename = "@click")]
    #[serde(default)]
    pub click: Option<u32>,
    #[serde(rename = "pivotArea")]
    pub pivot_area: Box<PivotArea>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct Selection {
    #[serde(rename = "@pane")]
    #[serde(default)]
    pub pane: Option<PaneType>,
    #[serde(rename = "@activeCell")]
    #[serde(default)]
    pub active_cell: Option<CellRef>,
    #[serde(rename = "@activeCellId")]
    #[serde(default)]
    pub active_cell_id: Option<u32>,
    #[serde(rename = "@sqref")]
    #[serde(default)]
    pub square_reference: Option<SquareRef>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct PageBreaks {
    #[serde(rename = "@count")]
    #[serde(default)]
    pub count: Option<u32>,
    #[serde(rename = "@manualBreakCount")]
    #[serde(default)]
    pub manual_break_count: Option<u32>,
    #[serde(rename = "brk")]
    #[serde(default)]
    pub brk: Vec<Box<PageBreak>>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct PageBreak {
    #[serde(rename = "@id")]
    #[serde(default)]
    pub id: Option<u32>,
    #[serde(rename = "@min")]
    #[serde(default)]
    pub start_column: Option<u32>,
    #[serde(rename = "@max")]
    #[serde(default)]
    pub end_column: Option<u32>,
    #[serde(rename = "@man")]
    #[serde(default)]
    pub man: Option<bool>,
    #[serde(rename = "@pt")]
    #[serde(default)]
    pub pt: Option<bool>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct OutlineProperties {
    #[serde(rename = "@applyStyles")]
    #[serde(default)]
    pub apply_styles: Option<bool>,
    #[serde(rename = "@summaryBelow")]
    #[serde(default)]
    pub summary_below: Option<bool>,
    #[serde(rename = "@summaryRight")]
    #[serde(default)]
    pub summary_right: Option<bool>,
    #[serde(rename = "@showOutlineSymbols")]
    #[serde(default)]
    pub show_outline_symbols: Option<bool>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct PageSetupProperties {
    #[serde(rename = "@autoPageBreaks")]
    #[serde(default)]
    pub auto_page_breaks: Option<bool>,
    #[serde(rename = "@fitToPage")]
    #[serde(default)]
    pub fit_to_page: Option<bool>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTDataConsolidate {
    #[serde(rename = "@function")]
    #[serde(default)]
    pub function: Option<STDataConsolidateFunction>,
    #[serde(rename = "@startLabels")]
    #[serde(default)]
    pub start_labels: Option<bool>,
    #[serde(rename = "@leftLabels")]
    #[serde(default)]
    pub left_labels: Option<bool>,
    #[serde(rename = "@topLabels")]
    #[serde(default)]
    pub top_labels: Option<bool>,
    #[serde(rename = "@link")]
    #[serde(default)]
    pub link: Option<bool>,
    #[serde(rename = "dataRefs")]
    #[serde(default)]
    pub data_refs: Option<Box<CTDataRefs>>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTDataRefs {
    #[serde(rename = "@count")]
    #[serde(default)]
    pub count: Option<u32>,
    #[serde(rename = "dataRef")]
    #[serde(default)]
    pub data_ref: Vec<Box<CTDataRef>>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTDataRef {
    #[serde(rename = "@ref")]
    #[serde(default)]
    pub reference: Option<Reference>,
    #[serde(rename = "@name")]
    #[serde(default)]
    pub name: Option<XmlString>,
    #[serde(rename = "@sheet")]
    #[serde(default)]
    pub sheet: Option<XmlString>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct MergedCells {
    #[serde(rename = "@count")]
    #[serde(default)]
    pub count: Option<u32>,
    #[serde(rename = "mergeCell")]
    #[serde(default)]
    pub merge_cell: Vec<Box<MergedCell>>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct MergedCell {
    #[serde(rename = "@ref")]
    pub reference: Reference,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct SmartTags {
    #[serde(rename = "cellSmartTags")]
    #[serde(default)]
    pub cell_smart_tags: Vec<Box<CellSmartTags>>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CellSmartTags {
    #[serde(rename = "@r")]
    pub reference: CellRef,
    #[serde(rename = "cellSmartTag")]
    #[serde(default)]
    pub cell_smart_tag: Vec<Box<CellSmartTag>>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CellSmartTag {
    #[serde(rename = "@type")]
    pub r#type: u32,
    #[serde(rename = "@deleted")]
    #[serde(default)]
    pub deleted: Option<bool>,
    #[serde(rename = "@xmlBased")]
    #[serde(default)]
    pub xml_based: Option<bool>,
    #[serde(rename = "cellSmartTagPr")]
    #[serde(default)]
    pub cell_smart_tag_pr: Vec<Box<CTCellSmartTagPr>>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTCellSmartTagPr {
    #[serde(rename = "@key")]
    pub key: XmlString,
    #[serde(rename = "@val")]
    pub value: XmlString,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct Drawing;

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct LegacyDrawing;

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct DrawingHeaderFooter {
    #[serde(rename = "@lho")]
    #[serde(default)]
    pub lho: Option<u32>,
    #[serde(rename = "@lhe")]
    #[serde(default)]
    pub lhe: Option<u32>,
    #[serde(rename = "@lhf")]
    #[serde(default)]
    pub lhf: Option<u32>,
    #[serde(rename = "@cho")]
    #[serde(default)]
    pub cho: Option<u32>,
    #[serde(rename = "@che")]
    #[serde(default)]
    pub che: Option<u32>,
    #[serde(rename = "@chf")]
    #[serde(default)]
    pub chf: Option<u32>,
    #[serde(rename = "@rho")]
    #[serde(default)]
    pub rho: Option<u32>,
    #[serde(rename = "@rhe")]
    #[serde(default)]
    pub rhe: Option<u32>,
    #[serde(rename = "@rhf")]
    #[serde(default)]
    pub rhf: Option<u32>,
    #[serde(rename = "@lfo")]
    #[serde(default)]
    pub lfo: Option<u32>,
    #[serde(rename = "@lfe")]
    #[serde(default)]
    pub lfe: Option<u32>,
    #[serde(rename = "@lff")]
    #[serde(default)]
    pub lff: Option<u32>,
    #[serde(rename = "@cfo")]
    #[serde(default)]
    pub cfo: Option<u32>,
    #[serde(rename = "@cfe")]
    #[serde(default)]
    pub cfe: Option<u32>,
    #[serde(rename = "@cff")]
    #[serde(default)]
    pub cff: Option<u32>,
    #[serde(rename = "@rfo")]
    #[serde(default)]
    pub rfo: Option<u32>,
    #[serde(rename = "@rfe")]
    #[serde(default)]
    pub rfe: Option<u32>,
    #[serde(rename = "@rff")]
    #[serde(default)]
    pub rff: Option<u32>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CustomSheetViews {
    #[serde(rename = "customSheetView")]
    #[serde(default)]
    pub custom_sheet_view: Vec<Box<CustomSheetView>>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CustomSheetView {
    #[serde(rename = "@guid")]
    pub guid: Guid,
    #[serde(rename = "@scale")]
    #[serde(default)]
    pub scale: Option<u32>,
    #[serde(rename = "@colorId")]
    #[serde(default)]
    pub color_id: Option<u32>,
    #[serde(rename = "@showPageBreaks")]
    #[serde(default)]
    pub show_page_breaks: Option<bool>,
    #[serde(rename = "@showFormulas")]
    #[serde(default)]
    pub show_formulas: Option<bool>,
    #[serde(rename = "@showGridLines")]
    #[serde(default)]
    pub show_grid_lines: Option<bool>,
    #[serde(rename = "@showRowCol")]
    #[serde(default)]
    pub show_row_col: Option<bool>,
    #[serde(rename = "@outlineSymbols")]
    #[serde(default)]
    pub outline_symbols: Option<bool>,
    #[serde(rename = "@zeroValues")]
    #[serde(default)]
    pub zero_values: Option<bool>,
    #[serde(rename = "@fitToPage")]
    #[serde(default)]
    pub fit_to_page: Option<bool>,
    #[serde(rename = "@printArea")]
    #[serde(default)]
    pub print_area: Option<bool>,
    #[serde(rename = "@filter")]
    #[serde(default)]
    pub filter: Option<bool>,
    #[serde(rename = "@showAutoFilter")]
    #[serde(default)]
    pub show_auto_filter: Option<bool>,
    #[serde(rename = "@hiddenRows")]
    #[serde(default)]
    pub hidden_rows: Option<bool>,
    #[serde(rename = "@hiddenColumns")]
    #[serde(default)]
    pub hidden_columns: Option<bool>,
    #[serde(rename = "@state")]
    #[serde(default)]
    pub state: Option<SheetState>,
    #[serde(rename = "@filterUnique")]
    #[serde(default)]
    pub filter_unique: Option<bool>,
    #[serde(rename = "@view")]
    #[serde(default)]
    pub view: Option<SheetViewType>,
    #[serde(rename = "@showRuler")]
    #[serde(default)]
    pub show_ruler: Option<bool>,
    #[serde(rename = "@topLeftCell")]
    #[serde(default)]
    pub top_left_cell: Option<CellRef>,
    #[serde(rename = "pane")]
    #[serde(default)]
    pub pane: Option<Box<Pane>>,
    #[serde(rename = "selection")]
    #[serde(default)]
    pub selection: Option<Box<Selection>>,
    #[serde(rename = "rowBreaks")]
    #[serde(default)]
    pub row_breaks: Option<Box<PageBreaks>>,
    #[serde(rename = "colBreaks")]
    #[serde(default)]
    pub col_breaks: Option<Box<PageBreaks>>,
    #[serde(rename = "pageMargins")]
    #[serde(default)]
    pub page_margins: Option<Box<PageMargins>>,
    #[serde(rename = "printOptions")]
    #[serde(default)]
    pub print_options: Option<Box<PrintOptions>>,
    #[serde(rename = "pageSetup")]
    #[serde(default)]
    pub page_setup: Option<Box<PageSetup>>,
    #[serde(rename = "headerFooter")]
    #[serde(default)]
    pub header_footer: Option<Box<HeaderFooter>>,
    #[serde(rename = "autoFilter")]
    #[serde(default)]
    pub auto_filter: Option<Box<AutoFilter>>,
    #[serde(rename = "extLst")]
    #[serde(default)]
    pub extension_list: Option<Box<ExtensionList>>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct DataValidations {
    #[serde(rename = "@disablePrompts")]
    #[serde(default)]
    pub disable_prompts: Option<bool>,
    #[serde(rename = "@xWindow")]
    #[serde(default)]
    pub x_window: Option<u32>,
    #[serde(rename = "@yWindow")]
    #[serde(default)]
    pub y_window: Option<u32>,
    #[serde(rename = "@count")]
    #[serde(default)]
    pub count: Option<u32>,
    #[serde(rename = "dataValidation")]
    #[serde(default)]
    pub data_validation: Vec<Box<DataValidation>>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct DataValidation {
    #[serde(rename = "@type")]
    #[serde(default)]
    pub r#type: Option<ValidationType>,
    #[serde(rename = "@errorStyle")]
    #[serde(default)]
    pub error_style: Option<ValidationErrorStyle>,
    #[serde(rename = "@imeMode")]
    #[serde(default)]
    pub ime_mode: Option<STDataValidationImeMode>,
    #[serde(rename = "@operator")]
    #[serde(default)]
    pub operator: Option<ValidationOperator>,
    #[serde(rename = "@allowBlank")]
    #[serde(default)]
    pub allow_blank: Option<bool>,
    #[serde(rename = "@showDropDown")]
    #[serde(default)]
    pub show_drop_down: Option<bool>,
    #[serde(rename = "@showInputMessage")]
    #[serde(default)]
    pub show_input_message: Option<bool>,
    #[serde(rename = "@showErrorMessage")]
    #[serde(default)]
    pub show_error_message: Option<bool>,
    #[serde(rename = "@errorTitle")]
    #[serde(default)]
    pub error_title: Option<XmlString>,
    #[serde(rename = "@error")]
    #[serde(default)]
    pub error: Option<XmlString>,
    #[serde(rename = "@promptTitle")]
    #[serde(default)]
    pub prompt_title: Option<XmlString>,
    #[serde(rename = "@prompt")]
    #[serde(default)]
    pub prompt: Option<XmlString>,
    #[serde(rename = "@sqref")]
    pub square_reference: SquareRef,
    #[serde(rename = "formula1")]
    #[serde(default)]
    pub formula1: Option<STFormula>,
    #[serde(rename = "formula2")]
    #[serde(default)]
    pub formula2: Option<STFormula>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct ConditionalFormatting {
    #[serde(rename = "@pivot")]
    #[serde(default)]
    pub pivot: Option<bool>,
    #[serde(rename = "@sqref")]
    #[serde(default)]
    pub square_reference: Option<SquareRef>,
    #[serde(rename = "cfRule")]
    #[serde(default)]
    pub cf_rule: Vec<Box<ConditionalRule>>,
    #[serde(rename = "extLst")]
    #[serde(default)]
    pub extension_list: Option<Box<ExtensionList>>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct ConditionalRule {
    #[serde(rename = "@type")]
    #[serde(default)]
    pub r#type: Option<ConditionalType>,
    #[serde(rename = "@dxfId")]
    #[serde(default)]
    pub dxf_id: Option<STDxfId>,
    #[serde(rename = "@priority")]
    pub priority: i32,
    #[serde(rename = "@stopIfTrue")]
    #[serde(default)]
    pub stop_if_true: Option<bool>,
    #[serde(rename = "@aboveAverage")]
    #[serde(default)]
    pub above_average: Option<bool>,
    #[serde(rename = "@percent")]
    #[serde(default)]
    pub percent: Option<bool>,
    #[serde(rename = "@bottom")]
    #[serde(default)]
    pub bottom: Option<bool>,
    #[serde(rename = "@operator")]
    #[serde(default)]
    pub operator: Option<ConditionalOperator>,
    #[serde(rename = "@text")]
    #[serde(default)]
    pub text: Option<String>,
    #[serde(rename = "@timePeriod")]
    #[serde(default)]
    pub time_period: Option<STTimePeriod>,
    #[serde(rename = "@rank")]
    #[serde(default)]
    pub rank: Option<u32>,
    #[serde(rename = "@stdDev")]
    #[serde(default)]
    pub std_dev: Option<i32>,
    #[serde(rename = "@equalAverage")]
    #[serde(default)]
    pub equal_average: Option<bool>,
    #[serde(rename = "formula")]
    #[serde(default)]
    pub formula: Vec<STFormula>,
    #[serde(rename = "colorScale")]
    #[serde(default)]
    pub color_scale: Option<Box<ColorScale>>,
    #[serde(rename = "dataBar")]
    #[serde(default)]
    pub data_bar: Option<Box<DataBar>>,
    #[serde(rename = "iconSet")]
    #[serde(default)]
    pub icon_set: Option<Box<IconSet>>,
    #[serde(rename = "extLst")]
    #[serde(default)]
    pub extension_list: Option<Box<ExtensionList>>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct Hyperlinks {
    #[serde(rename = "hyperlink")]
    #[serde(default)]
    pub hyperlink: Vec<Box<Hyperlink>>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct Hyperlink {
    #[serde(rename = "@ref")]
    pub reference: Reference,
    #[serde(rename = "@location")]
    #[serde(default)]
    pub location: Option<XmlString>,
    #[serde(rename = "@tooltip")]
    #[serde(default)]
    pub tooltip: Option<XmlString>,
    #[serde(rename = "@display")]
    #[serde(default)]
    pub display: Option<XmlString>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CellFormula {
    #[serde(rename = "$text")]
    pub text: String,
    #[serde(rename = "@t")]
    #[serde(default)]
    pub cell_type: Option<FormulaType>,
    #[serde(rename = "@aca")]
    #[serde(default)]
    pub aca: Option<bool>,
    #[serde(rename = "@ref")]
    #[serde(default)]
    pub reference: Option<Reference>,
    #[serde(rename = "@dt2D")]
    #[serde(default)]
    pub dt2_d: Option<bool>,
    #[serde(rename = "@dtr")]
    #[serde(default)]
    pub dtr: Option<bool>,
    #[serde(rename = "@del1")]
    #[serde(default)]
    pub del1: Option<bool>,
    #[serde(rename = "@del2")]
    #[serde(default)]
    pub del2: Option<bool>,
    #[serde(rename = "@r1")]
    #[serde(default)]
    pub r1: Option<CellRef>,
    #[serde(rename = "@r2")]
    #[serde(default)]
    pub r2: Option<CellRef>,
    #[serde(rename = "@ca")]
    #[serde(default)]
    pub ca: Option<bool>,
    #[serde(rename = "@si")]
    #[serde(default)]
    pub si: Option<u32>,
    #[serde(rename = "@bx")]
    #[serde(default)]
    pub bx: Option<bool>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct ColorScale {
    #[serde(rename = "cfvo")]
    #[serde(default)]
    pub cfvo: Vec<Box<ConditionalFormatValue>>,
    #[serde(rename = "color")]
    #[serde(default)]
    pub color: Vec<Box<Color>>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct DataBar {
    #[serde(rename = "@minLength")]
    #[serde(default)]
    pub min_length: Option<u32>,
    #[serde(rename = "@maxLength")]
    #[serde(default)]
    pub max_length: Option<u32>,
    #[serde(rename = "@showValue")]
    #[serde(default)]
    pub show_value: Option<bool>,
    #[serde(rename = "cfvo")]
    #[serde(default)]
    pub cfvo: Vec<Box<ConditionalFormatValue>>,
    #[serde(rename = "color")]
    pub color: Box<Color>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct IconSet {
    #[serde(rename = "@iconSet")]
    #[serde(default)]
    pub icon_set: Option<IconSetType>,
    #[serde(rename = "@showValue")]
    #[serde(default)]
    pub show_value: Option<bool>,
    #[serde(rename = "@percent")]
    #[serde(default)]
    pub percent: Option<bool>,
    #[serde(rename = "@reverse")]
    #[serde(default)]
    pub reverse: Option<bool>,
    #[serde(rename = "cfvo")]
    #[serde(default)]
    pub cfvo: Vec<Box<ConditionalFormatValue>>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct ConditionalFormatValue {
    #[serde(rename = "@type")]
    pub r#type: ConditionalValueType,
    #[serde(rename = "@val")]
    #[serde(default)]
    pub value: Option<XmlString>,
    #[serde(rename = "@gte")]
    #[serde(default)]
    pub gte: Option<bool>,
    #[serde(rename = "extLst")]
    #[serde(default)]
    pub extension_list: Option<Box<ExtensionList>>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct PageMargins {
    #[serde(rename = "@left")]
    pub left: f64,
    #[serde(rename = "@right")]
    pub right: f64,
    #[serde(rename = "@top")]
    pub top: f64,
    #[serde(rename = "@bottom")]
    pub bottom: f64,
    #[serde(rename = "@header")]
    pub header: f64,
    #[serde(rename = "@footer")]
    pub footer: f64,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct PrintOptions {
    #[serde(rename = "@horizontalCentered")]
    #[serde(default)]
    pub horizontal_centered: Option<bool>,
    #[serde(rename = "@verticalCentered")]
    #[serde(default)]
    pub vertical_centered: Option<bool>,
    #[serde(rename = "@headings")]
    #[serde(default)]
    pub headings: Option<bool>,
    #[serde(rename = "@gridLines")]
    #[serde(default)]
    pub grid_lines: Option<bool>,
    #[serde(rename = "@gridLinesSet")]
    #[serde(default)]
    pub grid_lines_set: Option<bool>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct PageSetup {
    #[serde(rename = "@paperSize")]
    #[serde(default)]
    pub paper_size: Option<u32>,
    #[serde(rename = "@paperHeight")]
    #[serde(default)]
    pub paper_height: Option<STPositiveUniversalMeasure>,
    #[serde(rename = "@paperWidth")]
    #[serde(default)]
    pub paper_width: Option<STPositiveUniversalMeasure>,
    #[serde(rename = "@scale")]
    #[serde(default)]
    pub scale: Option<u32>,
    #[serde(rename = "@firstPageNumber")]
    #[serde(default)]
    pub first_page_number: Option<u32>,
    #[serde(rename = "@fitToWidth")]
    #[serde(default)]
    pub fit_to_width: Option<u32>,
    #[serde(rename = "@fitToHeight")]
    #[serde(default)]
    pub fit_to_height: Option<u32>,
    #[serde(rename = "@pageOrder")]
    #[serde(default)]
    pub page_order: Option<STPageOrder>,
    #[serde(rename = "@orientation")]
    #[serde(default)]
    pub orientation: Option<STOrientation>,
    #[serde(rename = "@usePrinterDefaults")]
    #[serde(default)]
    pub use_printer_defaults: Option<bool>,
    #[serde(rename = "@blackAndWhite")]
    #[serde(default)]
    pub black_and_white: Option<bool>,
    #[serde(rename = "@draft")]
    #[serde(default)]
    pub draft: Option<bool>,
    #[serde(rename = "@cellComments")]
    #[serde(default)]
    pub cell_comments: Option<STCellComments>,
    #[serde(rename = "@useFirstPageNumber")]
    #[serde(default)]
    pub use_first_page_number: Option<bool>,
    #[serde(rename = "@errors")]
    #[serde(default)]
    pub errors: Option<STPrintError>,
    #[serde(rename = "@horizontalDpi")]
    #[serde(default)]
    pub horizontal_dpi: Option<u32>,
    #[serde(rename = "@verticalDpi")]
    #[serde(default)]
    pub vertical_dpi: Option<u32>,
    #[serde(rename = "@copies")]
    #[serde(default)]
    pub copies: Option<u32>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct HeaderFooter {
    #[serde(rename = "@differentOddEven")]
    #[serde(default)]
    pub different_odd_even: Option<bool>,
    #[serde(rename = "@differentFirst")]
    #[serde(default)]
    pub different_first: Option<bool>,
    #[serde(rename = "@scaleWithDoc")]
    #[serde(default)]
    pub scale_with_doc: Option<bool>,
    #[serde(rename = "@alignWithMargins")]
    #[serde(default)]
    pub align_with_margins: Option<bool>,
    #[serde(rename = "oddHeader")]
    #[serde(default)]
    pub odd_header: Option<XmlString>,
    #[serde(rename = "oddFooter")]
    #[serde(default)]
    pub odd_footer: Option<XmlString>,
    #[serde(rename = "evenHeader")]
    #[serde(default)]
    pub even_header: Option<XmlString>,
    #[serde(rename = "evenFooter")]
    #[serde(default)]
    pub even_footer: Option<XmlString>,
    #[serde(rename = "firstHeader")]
    #[serde(default)]
    pub first_header: Option<XmlString>,
    #[serde(rename = "firstFooter")]
    #[serde(default)]
    pub first_footer: Option<XmlString>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct Scenarios {
    #[serde(rename = "@current")]
    #[serde(default)]
    pub current: Option<u32>,
    #[serde(rename = "@show")]
    #[serde(default)]
    pub show: Option<u32>,
    #[serde(rename = "@sqref")]
    #[serde(default)]
    pub square_reference: Option<SquareRef>,
    #[serde(rename = "scenario")]
    #[serde(default)]
    pub scenario: Vec<Box<Scenario>>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct SheetProtection {
    #[serde(rename = "@password")]
    #[serde(default)]
    pub password: Option<STUnsignedShortHex>,
    #[serde(rename = "@algorithmName")]
    #[serde(default)]
    pub algorithm_name: Option<XmlString>,
    #[serde(rename = "@hashValue")]
    #[serde(default)]
    pub hash_value: Option<Vec<u8>>,
    #[serde(rename = "@saltValue")]
    #[serde(default)]
    pub salt_value: Option<Vec<u8>>,
    #[serde(rename = "@spinCount")]
    #[serde(default)]
    pub spin_count: Option<u32>,
    #[serde(rename = "@sheet")]
    #[serde(default)]
    pub sheet: Option<bool>,
    #[serde(rename = "@objects")]
    #[serde(default)]
    pub objects: Option<bool>,
    #[serde(rename = "@scenarios")]
    #[serde(default)]
    pub scenarios: Option<bool>,
    #[serde(rename = "@formatCells")]
    #[serde(default)]
    pub format_cells: Option<bool>,
    #[serde(rename = "@formatColumns")]
    #[serde(default)]
    pub format_columns: Option<bool>,
    #[serde(rename = "@formatRows")]
    #[serde(default)]
    pub format_rows: Option<bool>,
    #[serde(rename = "@insertColumns")]
    #[serde(default)]
    pub insert_columns: Option<bool>,
    #[serde(rename = "@insertRows")]
    #[serde(default)]
    pub insert_rows: Option<bool>,
    #[serde(rename = "@insertHyperlinks")]
    #[serde(default)]
    pub insert_hyperlinks: Option<bool>,
    #[serde(rename = "@deleteColumns")]
    #[serde(default)]
    pub delete_columns: Option<bool>,
    #[serde(rename = "@deleteRows")]
    #[serde(default)]
    pub delete_rows: Option<bool>,
    #[serde(rename = "@selectLockedCells")]
    #[serde(default)]
    pub select_locked_cells: Option<bool>,
    #[serde(rename = "@sort")]
    #[serde(default)]
    pub sort: Option<bool>,
    #[serde(rename = "@autoFilter")]
    #[serde(default)]
    pub auto_filter: Option<bool>,
    #[serde(rename = "@pivotTables")]
    #[serde(default)]
    pub pivot_tables: Option<bool>,
    #[serde(rename = "@selectUnlockedCells")]
    #[serde(default)]
    pub select_unlocked_cells: Option<bool>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct ProtectedRanges {
    #[serde(rename = "protectedRange")]
    #[serde(default)]
    pub protected_range: Vec<Box<ProtectedRange>>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct ProtectedRange {
    #[serde(rename = "@password")]
    #[serde(default)]
    pub password: Option<STUnsignedShortHex>,
    #[serde(rename = "@sqref")]
    pub square_reference: SquareRef,
    #[serde(rename = "@name")]
    pub name: XmlString,
    #[serde(rename = "@securityDescriptor")]
    #[serde(default)]
    pub security_descriptor: Option<String>,
    #[serde(rename = "@algorithmName")]
    #[serde(default)]
    pub algorithm_name: Option<XmlString>,
    #[serde(rename = "@hashValue")]
    #[serde(default)]
    pub hash_value: Option<Vec<u8>>,
    #[serde(rename = "@saltValue")]
    #[serde(default)]
    pub salt_value: Option<Vec<u8>>,
    #[serde(rename = "@spinCount")]
    #[serde(default)]
    pub spin_count: Option<u32>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct Scenario {
    #[serde(rename = "@name")]
    pub name: XmlString,
    #[serde(rename = "@locked")]
    #[serde(default)]
    pub locked: Option<bool>,
    #[serde(rename = "@hidden")]
    #[serde(default)]
    pub hidden: Option<bool>,
    #[serde(rename = "@count")]
    #[serde(default)]
    pub count: Option<u32>,
    #[serde(rename = "@user")]
    #[serde(default)]
    pub user: Option<XmlString>,
    #[serde(rename = "@comment")]
    #[serde(default)]
    pub comment: Option<XmlString>,
    #[serde(rename = "inputCells")]
    #[serde(default)]
    pub input_cells: Vec<Box<InputCells>>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct InputCells {
    #[serde(rename = "@r")]
    pub reference: CellRef,
    #[serde(rename = "@deleted")]
    #[serde(default)]
    pub deleted: Option<bool>,
    #[serde(rename = "@undone")]
    #[serde(default)]
    pub undone: Option<bool>,
    #[serde(rename = "@val")]
    pub value: XmlString,
    #[serde(rename = "@numFmtId")]
    #[serde(default)]
    pub number_format_id: Option<STNumFmtId>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CellWatches {
    #[serde(rename = "cellWatch")]
    #[serde(default)]
    pub cell_watch: Vec<Box<CellWatch>>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CellWatch {
    #[serde(rename = "@r")]
    pub reference: CellRef,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct Chartsheet {
    #[serde(rename = "sheetPr")]
    #[serde(default)]
    pub sheet_properties: Option<Box<ChartsheetProperties>>,
    #[serde(rename = "sheetViews")]
    pub sheet_views: Box<ChartsheetViews>,
    #[serde(rename = "sheetProtection")]
    #[serde(default)]
    pub sheet_protection: Option<Box<ChartsheetProtection>>,
    #[serde(rename = "customSheetViews")]
    #[serde(default)]
    pub custom_sheet_views: Option<Box<CustomChartsheetViews>>,
    #[serde(rename = "pageMargins")]
    #[serde(default)]
    pub page_margins: Option<Box<PageMargins>>,
    #[serde(rename = "pageSetup")]
    #[serde(default)]
    pub page_setup: Option<Box<ChartsheetPageSetup>>,
    #[serde(rename = "headerFooter")]
    #[serde(default)]
    pub header_footer: Option<Box<HeaderFooter>>,
    #[serde(rename = "drawing")]
    pub drawing: Box<Drawing>,
    #[serde(rename = "legacyDrawing")]
    #[serde(default)]
    pub legacy_drawing: Option<Box<LegacyDrawing>>,
    #[serde(rename = "legacyDrawingHF")]
    #[serde(default)]
    pub legacy_drawing_h_f: Option<Box<LegacyDrawing>>,
    #[serde(rename = "drawingHF")]
    #[serde(default)]
    pub drawing_h_f: Option<Box<DrawingHeaderFooter>>,
    #[serde(rename = "picture")]
    #[serde(default)]
    pub picture: Option<Box<SheetBackgroundPicture>>,
    #[serde(rename = "webPublishItems")]
    #[serde(default)]
    pub web_publish_items: Option<Box<WebPublishItems>>,
    #[serde(rename = "extLst")]
    #[serde(default)]
    pub extension_list: Option<Box<ExtensionList>>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct ChartsheetProperties {
    #[serde(rename = "@published")]
    #[serde(default)]
    pub published: Option<bool>,
    #[serde(rename = "@codeName")]
    #[serde(default)]
    pub code_name: Option<String>,
    #[serde(rename = "tabColor")]
    #[serde(default)]
    pub tab_color: Option<Box<Color>>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct ChartsheetViews {
    #[serde(rename = "sheetView")]
    #[serde(default)]
    pub sheet_view: Vec<Box<ChartsheetView>>,
    #[serde(rename = "extLst")]
    #[serde(default)]
    pub extension_list: Option<Box<ExtensionList>>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct ChartsheetView {
    #[serde(rename = "@tabSelected")]
    #[serde(default)]
    pub tab_selected: Option<bool>,
    #[serde(rename = "@zoomScale")]
    #[serde(default)]
    pub zoom_scale: Option<u32>,
    #[serde(rename = "@workbookViewId")]
    pub workbook_view_id: u32,
    #[serde(rename = "@zoomToFit")]
    #[serde(default)]
    pub zoom_to_fit: Option<bool>,
    #[serde(rename = "extLst")]
    #[serde(default)]
    pub extension_list: Option<Box<ExtensionList>>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct ChartsheetProtection {
    #[serde(rename = "@password")]
    #[serde(default)]
    pub password: Option<STUnsignedShortHex>,
    #[serde(rename = "@algorithmName")]
    #[serde(default)]
    pub algorithm_name: Option<XmlString>,
    #[serde(rename = "@hashValue")]
    #[serde(default)]
    pub hash_value: Option<Vec<u8>>,
    #[serde(rename = "@saltValue")]
    #[serde(default)]
    pub salt_value: Option<Vec<u8>>,
    #[serde(rename = "@spinCount")]
    #[serde(default)]
    pub spin_count: Option<u32>,
    #[serde(rename = "@content")]
    #[serde(default)]
    pub content: Option<bool>,
    #[serde(rename = "@objects")]
    #[serde(default)]
    pub objects: Option<bool>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct ChartsheetPageSetup {
    #[serde(rename = "@paperSize")]
    #[serde(default)]
    pub paper_size: Option<u32>,
    #[serde(rename = "@paperHeight")]
    #[serde(default)]
    pub paper_height: Option<STPositiveUniversalMeasure>,
    #[serde(rename = "@paperWidth")]
    #[serde(default)]
    pub paper_width: Option<STPositiveUniversalMeasure>,
    #[serde(rename = "@firstPageNumber")]
    #[serde(default)]
    pub first_page_number: Option<u32>,
    #[serde(rename = "@orientation")]
    #[serde(default)]
    pub orientation: Option<STOrientation>,
    #[serde(rename = "@usePrinterDefaults")]
    #[serde(default)]
    pub use_printer_defaults: Option<bool>,
    #[serde(rename = "@blackAndWhite")]
    #[serde(default)]
    pub black_and_white: Option<bool>,
    #[serde(rename = "@draft")]
    #[serde(default)]
    pub draft: Option<bool>,
    #[serde(rename = "@useFirstPageNumber")]
    #[serde(default)]
    pub use_first_page_number: Option<bool>,
    #[serde(rename = "@horizontalDpi")]
    #[serde(default)]
    pub horizontal_dpi: Option<u32>,
    #[serde(rename = "@verticalDpi")]
    #[serde(default)]
    pub vertical_dpi: Option<u32>,
    #[serde(rename = "@copies")]
    #[serde(default)]
    pub copies: Option<u32>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CustomChartsheetViews {
    #[serde(rename = "customSheetView")]
    #[serde(default)]
    pub custom_sheet_view: Vec<Box<CustomChartsheetView>>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CustomChartsheetView {
    #[serde(rename = "@guid")]
    pub guid: Guid,
    #[serde(rename = "@scale")]
    #[serde(default)]
    pub scale: Option<u32>,
    #[serde(rename = "@state")]
    #[serde(default)]
    pub state: Option<SheetState>,
    #[serde(rename = "@zoomToFit")]
    #[serde(default)]
    pub zoom_to_fit: Option<bool>,
    #[serde(rename = "pageMargins")]
    #[serde(default)]
    pub page_margins: Option<Box<PageMargins>>,
    #[serde(rename = "pageSetup")]
    #[serde(default)]
    pub page_setup: Option<Box<ChartsheetPageSetup>>,
    #[serde(rename = "headerFooter")]
    #[serde(default)]
    pub header_footer: Option<Box<HeaderFooter>>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTCustomProperties {
    #[serde(rename = "customPr")]
    #[serde(default)]
    pub custom_pr: Vec<Box<CTCustomProperty>>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTCustomProperty {
    #[serde(rename = "@name")]
    pub name: XmlString,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct OleObjects {
    #[serde(rename = "oleObject")]
    #[serde(default)]
    pub ole_object: Vec<Box<OleObject>>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct OleObject {
    #[serde(rename = "@progId")]
    #[serde(default)]
    pub prog_id: Option<String>,
    #[serde(rename = "@dvAspect")]
    #[serde(default)]
    pub dv_aspect: Option<STDvAspect>,
    #[serde(rename = "@link")]
    #[serde(default)]
    pub link: Option<XmlString>,
    #[serde(rename = "@oleUpdate")]
    #[serde(default)]
    pub ole_update: Option<STOleUpdate>,
    #[serde(rename = "@autoLoad")]
    #[serde(default)]
    pub auto_load: Option<bool>,
    #[serde(rename = "@shapeId")]
    pub shape_id: u32,
    #[serde(rename = "objectPr")]
    #[serde(default)]
    pub object_pr: Option<Box<ObjectProperties>>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct ObjectProperties {
    #[serde(rename = "@locked")]
    #[serde(default)]
    pub locked: Option<bool>,
    #[serde(rename = "@defaultSize")]
    #[serde(default)]
    pub default_size: Option<bool>,
    #[serde(rename = "@print")]
    #[serde(default)]
    pub print: Option<bool>,
    #[serde(rename = "@disabled")]
    #[serde(default)]
    pub disabled: Option<bool>,
    #[serde(rename = "@uiObject")]
    #[serde(default)]
    pub ui_object: Option<bool>,
    #[serde(rename = "@autoFill")]
    #[serde(default)]
    pub auto_fill: Option<bool>,
    #[serde(rename = "@autoLine")]
    #[serde(default)]
    pub auto_line: Option<bool>,
    #[serde(rename = "@autoPict")]
    #[serde(default)]
    pub auto_pict: Option<bool>,
    #[serde(rename = "@macro")]
    #[serde(default)]
    pub r#macro: Option<STFormula>,
    #[serde(rename = "@altText")]
    #[serde(default)]
    pub alt_text: Option<XmlString>,
    #[serde(rename = "@dde")]
    #[serde(default)]
    pub dde: Option<bool>,
    #[serde(rename = "anchor")]
    pub anchor: Box<ObjectAnchor>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct WebPublishItems {
    #[serde(rename = "@count")]
    #[serde(default)]
    pub count: Option<u32>,
    #[serde(rename = "webPublishItem")]
    #[serde(default)]
    pub web_publish_item: Vec<Box<WebPublishItem>>,
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
    #[serde(default)]
    pub source_ref: Option<Reference>,
    #[serde(rename = "@sourceObject")]
    #[serde(default)]
    pub source_object: Option<XmlString>,
    #[serde(rename = "@destinationFile")]
    pub destination_file: XmlString,
    #[serde(rename = "@title")]
    #[serde(default)]
    pub title: Option<XmlString>,
    #[serde(rename = "@autoRepublish")]
    #[serde(default)]
    pub auto_republish: Option<bool>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct Controls {
    #[serde(rename = "control")]
    #[serde(default)]
    pub control: Vec<Box<Control>>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct Control {
    #[serde(rename = "@shapeId")]
    pub shape_id: u32,
    #[serde(rename = "@name")]
    #[serde(default)]
    pub name: Option<String>,
    #[serde(rename = "controlPr")]
    #[serde(default)]
    pub control_pr: Option<Box<CTControlPr>>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTControlPr {
    #[serde(rename = "@locked")]
    #[serde(default)]
    pub locked: Option<bool>,
    #[serde(rename = "@defaultSize")]
    #[serde(default)]
    pub default_size: Option<bool>,
    #[serde(rename = "@print")]
    #[serde(default)]
    pub print: Option<bool>,
    #[serde(rename = "@disabled")]
    #[serde(default)]
    pub disabled: Option<bool>,
    #[serde(rename = "@recalcAlways")]
    #[serde(default)]
    pub recalc_always: Option<bool>,
    #[serde(rename = "@uiObject")]
    #[serde(default)]
    pub ui_object: Option<bool>,
    #[serde(rename = "@autoFill")]
    #[serde(default)]
    pub auto_fill: Option<bool>,
    #[serde(rename = "@autoLine")]
    #[serde(default)]
    pub auto_line: Option<bool>,
    #[serde(rename = "@autoPict")]
    #[serde(default)]
    pub auto_pict: Option<bool>,
    #[serde(rename = "@macro")]
    #[serde(default)]
    pub r#macro: Option<STFormula>,
    #[serde(rename = "@altText")]
    #[serde(default)]
    pub alt_text: Option<XmlString>,
    #[serde(rename = "@linkedCell")]
    #[serde(default)]
    pub linked_cell: Option<STFormula>,
    #[serde(rename = "@listFillRange")]
    #[serde(default)]
    pub list_fill_range: Option<STFormula>,
    #[serde(rename = "@cf")]
    #[serde(default)]
    pub cf: Option<XmlString>,
    #[serde(rename = "anchor")]
    pub anchor: Box<ObjectAnchor>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct IgnoredErrors {
    #[serde(rename = "ignoredError")]
    #[serde(default)]
    pub ignored_error: Vec<Box<IgnoredError>>,
    #[serde(rename = "extLst")]
    #[serde(default)]
    pub extension_list: Option<Box<ExtensionList>>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct IgnoredError {
    #[serde(rename = "@sqref")]
    pub square_reference: SquareRef,
    #[serde(rename = "@evalError")]
    #[serde(default)]
    pub eval_error: Option<bool>,
    #[serde(rename = "@twoDigitTextYear")]
    #[serde(default)]
    pub two_digit_text_year: Option<bool>,
    #[serde(rename = "@numberStoredAsText")]
    #[serde(default)]
    pub number_stored_as_text: Option<bool>,
    #[serde(rename = "@formula")]
    #[serde(default)]
    pub formula: Option<bool>,
    #[serde(rename = "@formulaRange")]
    #[serde(default)]
    pub formula_range: Option<bool>,
    #[serde(rename = "@unlockedFormula")]
    #[serde(default)]
    pub unlocked_formula: Option<bool>,
    #[serde(rename = "@emptyCellReference")]
    #[serde(default)]
    pub empty_cell_reference: Option<bool>,
    #[serde(rename = "@listDataValidation")]
    #[serde(default)]
    pub list_data_validation: Option<bool>,
    #[serde(rename = "@calculatedColumn")]
    #[serde(default)]
    pub calculated_column: Option<bool>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct TableParts {
    #[serde(rename = "@count")]
    #[serde(default)]
    pub count: Option<u32>,
    #[serde(rename = "tablePart")]
    #[serde(default)]
    pub table_part: Vec<Box<TablePart>>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct TablePart;

pub type SmlMetadata = Box<Metadata>;

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct Metadata {
    #[serde(rename = "metadataTypes")]
    #[serde(default)]
    pub metadata_types: Option<Box<MetadataTypes>>,
    #[serde(rename = "metadataStrings")]
    #[serde(default)]
    pub metadata_strings: Option<Box<MetadataStrings>>,
    #[serde(rename = "mdxMetadata")]
    #[serde(default)]
    pub mdx_metadata: Option<Box<CTMdxMetadata>>,
    #[serde(rename = "futureMetadata")]
    #[serde(default)]
    pub future_metadata: Vec<Box<CTFutureMetadata>>,
    #[serde(rename = "cellMetadata")]
    #[serde(default)]
    pub cell_metadata: Option<Box<MetadataBlocks>>,
    #[serde(rename = "valueMetadata")]
    #[serde(default)]
    pub value_metadata: Option<Box<MetadataBlocks>>,
    #[serde(rename = "extLst")]
    #[serde(default)]
    pub extension_list: Option<Box<ExtensionList>>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct MetadataTypes {
    #[serde(rename = "@count")]
    #[serde(default)]
    pub count: Option<u32>,
    #[serde(rename = "metadataType")]
    #[serde(default)]
    pub metadata_type: Vec<Box<MetadataType>>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct MetadataType {
    #[serde(rename = "@name")]
    pub name: XmlString,
    #[serde(rename = "@minSupportedVersion")]
    pub min_supported_version: u32,
    #[serde(rename = "@ghostRow")]
    #[serde(default)]
    pub ghost_row: Option<bool>,
    #[serde(rename = "@ghostCol")]
    #[serde(default)]
    pub ghost_col: Option<bool>,
    #[serde(rename = "@edit")]
    #[serde(default)]
    pub edit: Option<bool>,
    #[serde(rename = "@delete")]
    #[serde(default)]
    pub delete: Option<bool>,
    #[serde(rename = "@copy")]
    #[serde(default)]
    pub copy: Option<bool>,
    #[serde(rename = "@pasteAll")]
    #[serde(default)]
    pub paste_all: Option<bool>,
    #[serde(rename = "@pasteFormulas")]
    #[serde(default)]
    pub paste_formulas: Option<bool>,
    #[serde(rename = "@pasteValues")]
    #[serde(default)]
    pub paste_values: Option<bool>,
    #[serde(rename = "@pasteFormats")]
    #[serde(default)]
    pub paste_formats: Option<bool>,
    #[serde(rename = "@pasteComments")]
    #[serde(default)]
    pub paste_comments: Option<bool>,
    #[serde(rename = "@pasteDataValidation")]
    #[serde(default)]
    pub paste_data_validation: Option<bool>,
    #[serde(rename = "@pasteBorders")]
    #[serde(default)]
    pub paste_borders: Option<bool>,
    #[serde(rename = "@pasteColWidths")]
    #[serde(default)]
    pub paste_col_widths: Option<bool>,
    #[serde(rename = "@pasteNumberFormats")]
    #[serde(default)]
    pub paste_number_formats: Option<bool>,
    #[serde(rename = "@merge")]
    #[serde(default)]
    pub merge: Option<bool>,
    #[serde(rename = "@splitFirst")]
    #[serde(default)]
    pub split_first: Option<bool>,
    #[serde(rename = "@splitAll")]
    #[serde(default)]
    pub split_all: Option<bool>,
    #[serde(rename = "@rowColShift")]
    #[serde(default)]
    pub row_col_shift: Option<bool>,
    #[serde(rename = "@clearAll")]
    #[serde(default)]
    pub clear_all: Option<bool>,
    #[serde(rename = "@clearFormats")]
    #[serde(default)]
    pub clear_formats: Option<bool>,
    #[serde(rename = "@clearContents")]
    #[serde(default)]
    pub clear_contents: Option<bool>,
    #[serde(rename = "@clearComments")]
    #[serde(default)]
    pub clear_comments: Option<bool>,
    #[serde(rename = "@assign")]
    #[serde(default)]
    pub assign: Option<bool>,
    #[serde(rename = "@coerce")]
    #[serde(default)]
    pub coerce: Option<bool>,
    #[serde(rename = "@adjust")]
    #[serde(default)]
    pub adjust: Option<bool>,
    #[serde(rename = "@cellMeta")]
    #[serde(default)]
    pub cell_meta: Option<bool>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct MetadataBlocks {
    #[serde(rename = "@count")]
    #[serde(default)]
    pub count: Option<u32>,
    #[serde(rename = "bk")]
    #[serde(default)]
    pub bk: Vec<Box<MetadataBlock>>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct MetadataBlock {
    #[serde(rename = "rc")]
    #[serde(default)]
    pub rc: Vec<Box<MetadataRecord>>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct MetadataRecord {
    #[serde(rename = "@t")]
    pub cell_type: u32,
    #[serde(rename = "@v")]
    pub value: u32,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTFutureMetadata {
    #[serde(rename = "@name")]
    pub name: XmlString,
    #[serde(rename = "@count")]
    #[serde(default)]
    pub count: Option<u32>,
    #[serde(rename = "bk")]
    #[serde(default)]
    pub bk: Vec<Box<CTFutureMetadataBlock>>,
    #[serde(rename = "extLst")]
    #[serde(default)]
    pub extension_list: Option<Box<ExtensionList>>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTFutureMetadataBlock {
    #[serde(rename = "extLst")]
    #[serde(default)]
    pub extension_list: Option<Box<ExtensionList>>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTMdxMetadata {
    #[serde(rename = "@count")]
    #[serde(default)]
    pub count: Option<u32>,
    #[serde(rename = "mdx")]
    #[serde(default)]
    pub mdx: Vec<Box<CTMdx>>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTMdx {
    #[serde(rename = "@n")]
    pub n: u32,
    #[serde(rename = "@f")]
    pub formula: STMdxFunctionType,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTMdxTuple {
    #[serde(rename = "@c")]
    #[serde(default)]
    pub cells: Option<u32>,
    #[serde(rename = "@ct")]
    #[serde(default)]
    pub ct: Option<XmlString>,
    #[serde(rename = "@si")]
    #[serde(default)]
    pub si: Option<u32>,
    #[serde(rename = "@fi")]
    #[serde(default)]
    pub fi: Option<u32>,
    #[serde(rename = "@bc")]
    #[serde(default)]
    pub bc: Option<STUnsignedIntHex>,
    #[serde(rename = "@fc")]
    #[serde(default)]
    pub fc: Option<STUnsignedIntHex>,
    #[serde(rename = "@i")]
    #[serde(default)]
    pub i: Option<bool>,
    #[serde(rename = "@u")]
    #[serde(default)]
    pub u: Option<bool>,
    #[serde(rename = "@st")]
    #[serde(default)]
    pub st: Option<bool>,
    #[serde(rename = "@b")]
    #[serde(default)]
    pub b: Option<bool>,
    #[serde(rename = "n")]
    #[serde(default)]
    pub n: Vec<Box<CTMetadataStringIndex>>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTMdxSet {
    #[serde(rename = "@ns")]
    pub ns: u32,
    #[serde(rename = "@c")]
    #[serde(default)]
    pub cells: Option<u32>,
    #[serde(rename = "@o")]
    #[serde(default)]
    pub o: Option<STMdxSetOrder>,
    #[serde(rename = "n")]
    #[serde(default)]
    pub n: Vec<Box<CTMetadataStringIndex>>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTMdxMemeberProp {
    #[serde(rename = "@n")]
    pub n: u32,
    #[serde(rename = "@np")]
    pub np: u32,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTMdxKPI {
    #[serde(rename = "@n")]
    pub n: u32,
    #[serde(rename = "@np")]
    pub np: u32,
    #[serde(rename = "@p")]
    pub p: STMdxKPIProperty,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTMetadataStringIndex {
    #[serde(rename = "@x")]
    pub x: u32,
    #[serde(rename = "@s")]
    #[serde(default)]
    pub style_index: Option<bool>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct MetadataStrings {
    #[serde(rename = "@count")]
    #[serde(default)]
    pub count: Option<u32>,
    #[serde(rename = "s")]
    #[serde(default)]
    pub style_index: Vec<Box<CTXStringElement>>,
}

pub type SmlSingleXmlCells = Box<SingleXmlCells>;

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct SingleXmlCells {
    #[serde(rename = "singleXmlCell")]
    #[serde(default)]
    pub single_xml_cell: Vec<Box<SingleXmlCell>>,
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
    #[serde(default)]
    pub extension_list: Option<Box<ExtensionList>>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct XmlCellProperties {
    #[serde(rename = "@id")]
    pub id: u32,
    #[serde(rename = "@uniqueName")]
    #[serde(default)]
    pub unique_name: Option<XmlString>,
    #[serde(rename = "xmlPr")]
    pub xml_pr: Box<XmlProperties>,
    #[serde(rename = "extLst")]
    #[serde(default)]
    pub extension_list: Option<Box<ExtensionList>>,
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
    #[serde(default)]
    pub extension_list: Option<Box<ExtensionList>>,
}

pub type SmlStyleSheet = Box<Stylesheet>;

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct Stylesheet {
    #[serde(rename = "numFmts")]
    #[serde(default)]
    pub num_fmts: Option<Box<NumberFormats>>,
    #[serde(rename = "fonts")]
    #[serde(default)]
    pub fonts: Option<Box<Fonts>>,
    #[serde(rename = "fills")]
    #[serde(default)]
    pub fills: Option<Box<Fills>>,
    #[serde(rename = "borders")]
    #[serde(default)]
    pub borders: Option<Box<Borders>>,
    #[serde(rename = "cellStyleXfs")]
    #[serde(default)]
    pub cell_style_xfs: Option<Box<CellStyleFormats>>,
    #[serde(rename = "cellXfs")]
    #[serde(default)]
    pub cell_xfs: Option<Box<CellFormats>>,
    #[serde(rename = "cellStyles")]
    #[serde(default)]
    pub cell_styles: Option<Box<CellStyles>>,
    #[serde(rename = "dxfs")]
    #[serde(default)]
    pub dxfs: Option<Box<DifferentialFormats>>,
    #[serde(rename = "tableStyles")]
    #[serde(default)]
    pub table_styles: Option<Box<TableStyles>>,
    #[serde(rename = "colors")]
    #[serde(default)]
    pub colors: Option<Box<Colors>>,
    #[serde(rename = "extLst")]
    #[serde(default)]
    pub extension_list: Option<Box<ExtensionList>>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CellAlignment {
    #[serde(rename = "@horizontal")]
    #[serde(default)]
    pub horizontal: Option<HorizontalAlignment>,
    #[serde(rename = "@vertical")]
    #[serde(default)]
    pub vertical: Option<VerticalAlignment>,
    #[serde(rename = "@textRotation")]
    #[serde(default)]
    pub text_rotation: Option<STTextRotation>,
    #[serde(rename = "@wrapText")]
    #[serde(default)]
    pub wrap_text: Option<bool>,
    #[serde(rename = "@indent")]
    #[serde(default)]
    pub indent: Option<u32>,
    #[serde(rename = "@relativeIndent")]
    #[serde(default)]
    pub relative_indent: Option<i32>,
    #[serde(rename = "@justifyLastLine")]
    #[serde(default)]
    pub justify_last_line: Option<bool>,
    #[serde(rename = "@shrinkToFit")]
    #[serde(default)]
    pub shrink_to_fit: Option<bool>,
    #[serde(rename = "@readingOrder")]
    #[serde(default)]
    pub reading_order: Option<u32>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct Borders {
    #[serde(rename = "@count")]
    #[serde(default)]
    pub count: Option<u32>,
    #[serde(rename = "border")]
    #[serde(default)]
    pub border: Vec<Box<Border>>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct Border {
    #[serde(rename = "@diagonalUp")]
    #[serde(default)]
    pub diagonal_up: Option<bool>,
    #[serde(rename = "@diagonalDown")]
    #[serde(default)]
    pub diagonal_down: Option<bool>,
    #[serde(rename = "@outline")]
    #[serde(default)]
    pub outline: Option<bool>,
    #[serde(rename = "start")]
    #[serde(default)]
    pub start: Option<Box<BorderProperties>>,
    #[serde(rename = "end")]
    #[serde(default)]
    pub end: Option<Box<BorderProperties>>,
    #[serde(rename = "left")]
    #[serde(default)]
    pub left: Option<Box<BorderProperties>>,
    #[serde(rename = "right")]
    #[serde(default)]
    pub right: Option<Box<BorderProperties>>,
    #[serde(rename = "top")]
    #[serde(default)]
    pub top: Option<Box<BorderProperties>>,
    #[serde(rename = "bottom")]
    #[serde(default)]
    pub bottom: Option<Box<BorderProperties>>,
    #[serde(rename = "diagonal")]
    #[serde(default)]
    pub diagonal: Option<Box<BorderProperties>>,
    #[serde(rename = "vertical")]
    #[serde(default)]
    pub vertical: Option<Box<BorderProperties>>,
    #[serde(rename = "horizontal")]
    #[serde(default)]
    pub horizontal: Option<Box<BorderProperties>>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct BorderProperties {
    #[serde(rename = "@style")]
    #[serde(default)]
    pub style: Option<BorderStyle>,
    #[serde(rename = "color")]
    #[serde(default)]
    pub color: Option<Box<Color>>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CellProtection {
    #[serde(rename = "@locked")]
    #[serde(default)]
    pub locked: Option<bool>,
    #[serde(rename = "@hidden")]
    #[serde(default)]
    pub hidden: Option<bool>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct Fonts {
    #[serde(rename = "@count")]
    #[serde(default)]
    pub count: Option<u32>,
    #[serde(rename = "font")]
    #[serde(default)]
    pub font: Vec<Box<Font>>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct Fills {
    #[serde(rename = "@count")]
    #[serde(default)]
    pub count: Option<u32>,
    #[serde(rename = "fill")]
    #[serde(default)]
    pub fill: Vec<Box<Fill>>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct Fill;

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct PatternFill {
    #[serde(rename = "@patternType")]
    #[serde(default)]
    pub pattern_type: Option<PatternType>,
    #[serde(rename = "fgColor")]
    #[serde(default)]
    pub fg_color: Option<Box<Color>>,
    #[serde(rename = "bgColor")]
    #[serde(default)]
    pub bg_color: Option<Box<Color>>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct Color {
    #[serde(rename = "@auto")]
    #[serde(default)]
    pub auto: Option<bool>,
    #[serde(rename = "@indexed")]
    #[serde(default)]
    pub indexed: Option<u32>,
    #[serde(rename = "@rgb")]
    #[serde(default)]
    pub rgb: Option<STUnsignedIntHex>,
    #[serde(rename = "@theme")]
    #[serde(default)]
    pub theme: Option<u32>,
    #[serde(rename = "@tint")]
    #[serde(default)]
    pub tint: Option<f64>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct GradientFill {
    #[serde(rename = "@type")]
    #[serde(default)]
    pub r#type: Option<GradientType>,
    #[serde(rename = "@degree")]
    #[serde(default)]
    pub degree: Option<f64>,
    #[serde(rename = "@left")]
    #[serde(default)]
    pub left: Option<f64>,
    #[serde(rename = "@right")]
    #[serde(default)]
    pub right: Option<f64>,
    #[serde(rename = "@top")]
    #[serde(default)]
    pub top: Option<f64>,
    #[serde(rename = "@bottom")]
    #[serde(default)]
    pub bottom: Option<f64>,
    #[serde(rename = "stop")]
    #[serde(default)]
    pub stop: Vec<Box<GradientStop>>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct GradientStop {
    #[serde(rename = "@position")]
    pub position: f64,
    #[serde(rename = "color")]
    pub color: Box<Color>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct NumberFormats {
    #[serde(rename = "@count")]
    #[serde(default)]
    pub count: Option<u32>,
    #[serde(rename = "numFmt")]
    #[serde(default)]
    pub num_fmt: Vec<Box<NumberFormat>>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct NumberFormat {
    #[serde(rename = "@numFmtId")]
    pub number_format_id: STNumFmtId,
    #[serde(rename = "@formatCode")]
    pub format_code: XmlString,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CellStyleFormats {
    #[serde(rename = "@count")]
    #[serde(default)]
    pub count: Option<u32>,
    #[serde(rename = "xf")]
    #[serde(default)]
    pub xf: Vec<Box<Format>>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CellFormats {
    #[serde(rename = "@count")]
    #[serde(default)]
    pub count: Option<u32>,
    #[serde(rename = "xf")]
    #[serde(default)]
    pub xf: Vec<Box<Format>>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct Format {
    #[serde(rename = "@numFmtId")]
    #[serde(default)]
    pub number_format_id: Option<STNumFmtId>,
    #[serde(rename = "@fontId")]
    #[serde(default)]
    pub font_id: Option<STFontId>,
    #[serde(rename = "@fillId")]
    #[serde(default)]
    pub fill_id: Option<STFillId>,
    #[serde(rename = "@borderId")]
    #[serde(default)]
    pub border_id: Option<STBorderId>,
    #[serde(rename = "@xfId")]
    #[serde(default)]
    pub format_id: Option<STCellStyleXfId>,
    #[serde(rename = "@quotePrefix")]
    #[serde(default)]
    pub quote_prefix: Option<bool>,
    #[serde(rename = "@pivotButton")]
    #[serde(default)]
    pub pivot_button: Option<bool>,
    #[serde(rename = "@applyNumberFormat")]
    #[serde(default)]
    pub apply_number_format: Option<bool>,
    #[serde(rename = "@applyFont")]
    #[serde(default)]
    pub apply_font: Option<bool>,
    #[serde(rename = "@applyFill")]
    #[serde(default)]
    pub apply_fill: Option<bool>,
    #[serde(rename = "@applyBorder")]
    #[serde(default)]
    pub apply_border: Option<bool>,
    #[serde(rename = "@applyAlignment")]
    #[serde(default)]
    pub apply_alignment: Option<bool>,
    #[serde(rename = "@applyProtection")]
    #[serde(default)]
    pub apply_protection: Option<bool>,
    #[serde(rename = "alignment")]
    #[serde(default)]
    pub alignment: Option<Box<CellAlignment>>,
    #[serde(rename = "protection")]
    #[serde(default)]
    pub protection: Option<Box<CellProtection>>,
    #[serde(rename = "extLst")]
    #[serde(default)]
    pub extension_list: Option<Box<ExtensionList>>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CellStyles {
    #[serde(rename = "@count")]
    #[serde(default)]
    pub count: Option<u32>,
    #[serde(rename = "cellStyle")]
    #[serde(default)]
    pub cell_style: Vec<Box<CellStyle>>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CellStyle {
    #[serde(rename = "@name")]
    #[serde(default)]
    pub name: Option<XmlString>,
    #[serde(rename = "@xfId")]
    pub format_id: STCellStyleXfId,
    #[serde(rename = "@builtinId")]
    #[serde(default)]
    pub builtin_id: Option<u32>,
    #[serde(rename = "@iLevel")]
    #[serde(default)]
    pub i_level: Option<u32>,
    #[serde(rename = "@hidden")]
    #[serde(default)]
    pub hidden: Option<bool>,
    #[serde(rename = "@customBuiltin")]
    #[serde(default)]
    pub custom_builtin: Option<bool>,
    #[serde(rename = "extLst")]
    #[serde(default)]
    pub extension_list: Option<Box<ExtensionList>>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct DifferentialFormats {
    #[serde(rename = "@count")]
    #[serde(default)]
    pub count: Option<u32>,
    #[serde(rename = "dxf")]
    #[serde(default)]
    pub dxf: Vec<Box<DifferentialFormat>>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct DifferentialFormat {
    #[serde(rename = "font")]
    #[serde(default)]
    pub font: Option<Box<Font>>,
    #[serde(rename = "numFmt")]
    #[serde(default)]
    pub num_fmt: Option<Box<NumberFormat>>,
    #[serde(rename = "fill")]
    #[serde(default)]
    pub fill: Option<Box<Fill>>,
    #[serde(rename = "alignment")]
    #[serde(default)]
    pub alignment: Option<Box<CellAlignment>>,
    #[serde(rename = "border")]
    #[serde(default)]
    pub border: Option<Box<Border>>,
    #[serde(rename = "protection")]
    #[serde(default)]
    pub protection: Option<Box<CellProtection>>,
    #[serde(rename = "extLst")]
    #[serde(default)]
    pub extension_list: Option<Box<ExtensionList>>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct Colors {
    #[serde(rename = "indexedColors")]
    #[serde(default)]
    pub indexed_colors: Option<Box<IndexedColors>>,
    #[serde(rename = "mruColors")]
    #[serde(default)]
    pub mru_colors: Option<Box<MostRecentColors>>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct IndexedColors {
    #[serde(rename = "rgbColor")]
    #[serde(default)]
    pub rgb_color: Vec<Box<RgbColor>>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct MostRecentColors {
    #[serde(rename = "color")]
    #[serde(default)]
    pub color: Vec<Box<Color>>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct RgbColor {
    #[serde(rename = "@rgb")]
    #[serde(default)]
    pub rgb: Option<STUnsignedIntHex>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct TableStyles {
    #[serde(rename = "@count")]
    #[serde(default)]
    pub count: Option<u32>,
    #[serde(rename = "@defaultTableStyle")]
    #[serde(default)]
    pub default_table_style: Option<String>,
    #[serde(rename = "@defaultPivotStyle")]
    #[serde(default)]
    pub default_pivot_style: Option<String>,
    #[serde(rename = "tableStyle")]
    #[serde(default)]
    pub table_style: Vec<Box<TableStyle>>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct TableStyle {
    #[serde(rename = "@name")]
    pub name: String,
    #[serde(rename = "@pivot")]
    #[serde(default)]
    pub pivot: Option<bool>,
    #[serde(rename = "@table")]
    #[serde(default)]
    pub table: Option<bool>,
    #[serde(rename = "@count")]
    #[serde(default)]
    pub count: Option<u32>,
    #[serde(rename = "tableStyleElement")]
    #[serde(default)]
    pub table_style_element: Vec<Box<TableStyleElement>>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct TableStyleElement {
    #[serde(rename = "@type")]
    pub r#type: STTableStyleType,
    #[serde(rename = "@size")]
    #[serde(default)]
    pub size: Option<u32>,
    #[serde(rename = "@dxfId")]
    #[serde(default)]
    pub dxf_id: Option<STDxfId>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct BooleanProperty {
    #[serde(rename = "@val")]
    #[serde(default)]
    pub value: Option<bool>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct FontSize {
    #[serde(rename = "@val")]
    pub value: f64,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct IntProperty {
    #[serde(rename = "@val")]
    pub value: i32,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct FontName {
    #[serde(rename = "@val")]
    pub value: XmlString,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct VerticalAlignFontProperty {
    #[serde(rename = "@val")]
    pub value: VerticalAlignRun,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct FontSchemeProperty {
    #[serde(rename = "@val")]
    pub value: FontScheme,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct UnderlineProperty {
    #[serde(rename = "@val")]
    #[serde(default)]
    pub value: Option<UnderlineStyle>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct Font;

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct FontFamily {
    #[serde(rename = "@val")]
    pub value: STFontFamily,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct SmlAGAutoFormat {
    #[serde(rename = "@autoFormatId")]
    #[serde(default)]
    pub auto_format_id: Option<u32>,
    #[serde(rename = "@applyNumberFormats")]
    #[serde(default)]
    pub apply_number_formats: Option<bool>,
    #[serde(rename = "@applyBorderFormats")]
    #[serde(default)]
    pub apply_border_formats: Option<bool>,
    #[serde(rename = "@applyFontFormats")]
    #[serde(default)]
    pub apply_font_formats: Option<bool>,
    #[serde(rename = "@applyPatternFormats")]
    #[serde(default)]
    pub apply_pattern_formats: Option<bool>,
    #[serde(rename = "@applyAlignmentFormats")]
    #[serde(default)]
    pub apply_alignment_formats: Option<bool>,
    #[serde(rename = "@applyWidthHeightFormats")]
    #[serde(default)]
    pub apply_width_height_formats: Option<bool>,
}

pub type SmlExternalLink = Box<ExternalLink>;

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct ExternalLink {
    #[serde(rename = "extLst")]
    #[serde(default)]
    pub extension_list: Option<Box<ExtensionList>>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct ExternalBook {
    #[serde(rename = "sheetNames")]
    #[serde(default)]
    pub sheet_names: Option<Box<CTExternalSheetNames>>,
    #[serde(rename = "definedNames")]
    #[serde(default)]
    pub defined_names: Option<Box<CTExternalDefinedNames>>,
    #[serde(rename = "sheetDataSet")]
    #[serde(default)]
    pub sheet_data_set: Option<Box<ExternalSheetDataSet>>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTExternalSheetNames {
    #[serde(rename = "sheetName")]
    #[serde(default)]
    pub sheet_name: Vec<Box<CTExternalSheetName>>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTExternalSheetName {
    #[serde(rename = "@val")]
    #[serde(default)]
    pub value: Option<XmlString>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTExternalDefinedNames {
    #[serde(rename = "definedName")]
    #[serde(default)]
    pub defined_name: Vec<Box<CTExternalDefinedName>>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTExternalDefinedName {
    #[serde(rename = "@name")]
    pub name: XmlString,
    #[serde(rename = "@refersTo")]
    #[serde(default)]
    pub refers_to: Option<XmlString>,
    #[serde(rename = "@sheetId")]
    #[serde(default)]
    pub sheet_id: Option<u32>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct ExternalSheetDataSet {
    #[serde(rename = "sheetData")]
    #[serde(default)]
    pub sheet_data: Vec<Box<ExternalSheetData>>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct ExternalSheetData {
    #[serde(rename = "@sheetId")]
    pub sheet_id: u32,
    #[serde(rename = "@refreshError")]
    #[serde(default)]
    pub refresh_error: Option<bool>,
    #[serde(rename = "row")]
    #[serde(default)]
    pub row: Vec<Box<ExternalRow>>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct ExternalRow {
    #[serde(rename = "@r")]
    pub reference: u32,
    #[serde(rename = "cell")]
    #[serde(default)]
    pub cell: Vec<Box<ExternalCell>>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct ExternalCell {
    #[serde(rename = "@r")]
    #[serde(default)]
    pub reference: Option<CellRef>,
    #[serde(rename = "@t")]
    #[serde(default)]
    pub cell_type: Option<CellType>,
    #[serde(rename = "@vm")]
    #[serde(default)]
    pub vm: Option<u32>,
    #[serde(rename = "v")]
    #[serde(default)]
    pub value: Option<XmlString>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct DdeLink {
    #[serde(rename = "@ddeService")]
    pub dde_service: XmlString,
    #[serde(rename = "@ddeTopic")]
    pub dde_topic: XmlString,
    #[serde(rename = "ddeItems")]
    #[serde(default)]
    pub dde_items: Option<Box<DdeItems>>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct DdeItems {
    #[serde(rename = "ddeItem")]
    #[serde(default)]
    pub dde_item: Vec<Box<DdeItem>>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct DdeItem {
    #[serde(rename = "@name")]
    #[serde(default)]
    pub name: Option<XmlString>,
    #[serde(rename = "@ole")]
    #[serde(default)]
    pub ole: Option<bool>,
    #[serde(rename = "@advise")]
    #[serde(default)]
    pub advise: Option<bool>,
    #[serde(rename = "@preferPic")]
    #[serde(default)]
    pub prefer_pic: Option<bool>,
    #[serde(rename = "values")]
    #[serde(default)]
    pub values: Option<Box<CTDdeValues>>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTDdeValues {
    #[serde(rename = "@rows")]
    #[serde(default)]
    pub rows: Option<u32>,
    #[serde(rename = "@cols")]
    #[serde(default)]
    pub cols: Option<u32>,
    #[serde(rename = "value")]
    #[serde(default)]
    pub value: Vec<Box<CTDdeValue>>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTDdeValue {
    #[serde(rename = "@t")]
    #[serde(default)]
    pub cell_type: Option<STDdeValueType>,
    #[serde(rename = "val")]
    pub value: XmlString,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct OleLink {
    #[serde(rename = "@progId")]
    pub prog_id: XmlString,
    #[serde(rename = "oleItems")]
    #[serde(default)]
    pub ole_items: Option<Box<OleItems>>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct OleItems {
    #[serde(rename = "oleItem")]
    #[serde(default)]
    pub ole_item: Vec<Box<OleItem>>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct OleItem {
    #[serde(rename = "@name")]
    pub name: XmlString,
    #[serde(rename = "@icon")]
    #[serde(default)]
    pub icon: Option<bool>,
    #[serde(rename = "@advise")]
    #[serde(default)]
    pub advise: Option<bool>,
    #[serde(rename = "@preferPic")]
    #[serde(default)]
    pub prefer_pic: Option<bool>,
}

pub type SmlTable = Box<Table>;

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct Table {
    #[serde(rename = "@id")]
    pub id: u32,
    #[serde(rename = "@name")]
    #[serde(default)]
    pub name: Option<XmlString>,
    #[serde(rename = "@displayName")]
    pub display_name: XmlString,
    #[serde(rename = "@comment")]
    #[serde(default)]
    pub comment: Option<XmlString>,
    #[serde(rename = "@ref")]
    pub reference: Reference,
    #[serde(rename = "@tableType")]
    #[serde(default)]
    pub table_type: Option<STTableType>,
    #[serde(rename = "@headerRowCount")]
    #[serde(default)]
    pub header_row_count: Option<u32>,
    #[serde(rename = "@insertRow")]
    #[serde(default)]
    pub insert_row: Option<bool>,
    #[serde(rename = "@insertRowShift")]
    #[serde(default)]
    pub insert_row_shift: Option<bool>,
    #[serde(rename = "@totalsRowCount")]
    #[serde(default)]
    pub totals_row_count: Option<u32>,
    #[serde(rename = "@totalsRowShown")]
    #[serde(default)]
    pub totals_row_shown: Option<bool>,
    #[serde(rename = "@published")]
    #[serde(default)]
    pub published: Option<bool>,
    #[serde(rename = "@headerRowDxfId")]
    #[serde(default)]
    pub header_row_dxf_id: Option<STDxfId>,
    #[serde(rename = "@dataDxfId")]
    #[serde(default)]
    pub data_dxf_id: Option<STDxfId>,
    #[serde(rename = "@totalsRowDxfId")]
    #[serde(default)]
    pub totals_row_dxf_id: Option<STDxfId>,
    #[serde(rename = "@headerRowBorderDxfId")]
    #[serde(default)]
    pub header_row_border_dxf_id: Option<STDxfId>,
    #[serde(rename = "@tableBorderDxfId")]
    #[serde(default)]
    pub table_border_dxf_id: Option<STDxfId>,
    #[serde(rename = "@totalsRowBorderDxfId")]
    #[serde(default)]
    pub totals_row_border_dxf_id: Option<STDxfId>,
    #[serde(rename = "@headerRowCellStyle")]
    #[serde(default)]
    pub header_row_cell_style: Option<XmlString>,
    #[serde(rename = "@dataCellStyle")]
    #[serde(default)]
    pub data_cell_style: Option<XmlString>,
    #[serde(rename = "@totalsRowCellStyle")]
    #[serde(default)]
    pub totals_row_cell_style: Option<XmlString>,
    #[serde(rename = "@connectionId")]
    #[serde(default)]
    pub connection_id: Option<u32>,
    #[serde(rename = "autoFilter")]
    #[serde(default)]
    pub auto_filter: Option<Box<AutoFilter>>,
    #[serde(rename = "sortState")]
    #[serde(default)]
    pub sort_state: Option<Box<SortState>>,
    #[serde(rename = "tableColumns")]
    pub table_columns: Box<TableColumns>,
    #[serde(rename = "tableStyleInfo")]
    #[serde(default)]
    pub table_style_info: Option<Box<TableStyleInfo>>,
    #[serde(rename = "extLst")]
    #[serde(default)]
    pub extension_list: Option<Box<ExtensionList>>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct TableStyleInfo {
    #[serde(rename = "@name")]
    #[serde(default)]
    pub name: Option<XmlString>,
    #[serde(rename = "@showFirstColumn")]
    #[serde(default)]
    pub show_first_column: Option<bool>,
    #[serde(rename = "@showLastColumn")]
    #[serde(default)]
    pub show_last_column: Option<bool>,
    #[serde(rename = "@showRowStripes")]
    #[serde(default)]
    pub show_row_stripes: Option<bool>,
    #[serde(rename = "@showColumnStripes")]
    #[serde(default)]
    pub show_column_stripes: Option<bool>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct TableColumns {
    #[serde(rename = "@count")]
    #[serde(default)]
    pub count: Option<u32>,
    #[serde(rename = "tableColumn")]
    #[serde(default)]
    pub table_column: Vec<Box<TableColumn>>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct TableColumn {
    #[serde(rename = "@id")]
    pub id: u32,
    #[serde(rename = "@uniqueName")]
    #[serde(default)]
    pub unique_name: Option<XmlString>,
    #[serde(rename = "@name")]
    pub name: XmlString,
    #[serde(rename = "@totalsRowFunction")]
    #[serde(default)]
    pub totals_row_function: Option<STTotalsRowFunction>,
    #[serde(rename = "@totalsRowLabel")]
    #[serde(default)]
    pub totals_row_label: Option<XmlString>,
    #[serde(rename = "@queryTableFieldId")]
    #[serde(default)]
    pub query_table_field_id: Option<u32>,
    #[serde(rename = "@headerRowDxfId")]
    #[serde(default)]
    pub header_row_dxf_id: Option<STDxfId>,
    #[serde(rename = "@dataDxfId")]
    #[serde(default)]
    pub data_dxf_id: Option<STDxfId>,
    #[serde(rename = "@totalsRowDxfId")]
    #[serde(default)]
    pub totals_row_dxf_id: Option<STDxfId>,
    #[serde(rename = "@headerRowCellStyle")]
    #[serde(default)]
    pub header_row_cell_style: Option<XmlString>,
    #[serde(rename = "@dataCellStyle")]
    #[serde(default)]
    pub data_cell_style: Option<XmlString>,
    #[serde(rename = "@totalsRowCellStyle")]
    #[serde(default)]
    pub totals_row_cell_style: Option<XmlString>,
    #[serde(rename = "calculatedColumnFormula")]
    #[serde(default)]
    pub calculated_column_formula: Option<Box<TableFormula>>,
    #[serde(rename = "totalsRowFormula")]
    #[serde(default)]
    pub totals_row_formula: Option<Box<TableFormula>>,
    #[serde(rename = "xmlColumnPr")]
    #[serde(default)]
    pub xml_column_pr: Option<Box<XmlColumnProperties>>,
    #[serde(rename = "extLst")]
    #[serde(default)]
    pub extension_list: Option<Box<ExtensionList>>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct TableFormula {
    #[serde(rename = "$text")]
    pub text: String,
    #[serde(rename = "@array")]
    #[serde(default)]
    pub array: Option<bool>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct XmlColumnProperties {
    #[serde(rename = "@mapId")]
    pub map_id: u32,
    #[serde(rename = "@xpath")]
    pub xpath: XmlString,
    #[serde(rename = "@denormalized")]
    #[serde(default)]
    pub denormalized: Option<bool>,
    #[serde(rename = "@xmlDataType")]
    pub xml_data_type: STXmlDataType,
    #[serde(rename = "extLst")]
    #[serde(default)]
    pub extension_list: Option<Box<ExtensionList>>,
}

pub type SmlVolTypes = Box<CTVolTypes>;

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTVolTypes {
    #[serde(rename = "volType")]
    #[serde(default)]
    pub vol_type: Vec<Box<CTVolType>>,
    #[serde(rename = "extLst")]
    #[serde(default)]
    pub extension_list: Option<Box<ExtensionList>>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTVolType {
    #[serde(rename = "@type")]
    pub r#type: STVolDepType,
    #[serde(rename = "main")]
    #[serde(default)]
    pub main: Vec<Box<CTVolMain>>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTVolMain {
    #[serde(rename = "@first")]
    pub first: XmlString,
    #[serde(rename = "tp")]
    #[serde(default)]
    pub tp: Vec<Box<CTVolTopic>>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTVolTopic {
    #[serde(rename = "@t")]
    #[serde(default)]
    pub cell_type: Option<STVolValueType>,
    #[serde(rename = "v")]
    pub value: XmlString,
    #[serde(rename = "stp")]
    #[serde(default)]
    pub stp: Vec<XmlString>,
    #[serde(rename = "tr")]
    #[serde(default)]
    pub tr: Vec<Box<CTVolTopicRef>>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTVolTopicRef {
    #[serde(rename = "@r")]
    pub reference: CellRef,
    #[serde(rename = "@s")]
    pub style_index: u32,
}

pub type SmlWorkbook = Box<Workbook>;

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct Workbook {
    #[serde(rename = "@conformance")]
    #[serde(default)]
    pub conformance: Option<STConformanceClass>,
    #[serde(rename = "fileVersion")]
    #[serde(default)]
    pub file_version: Option<Box<FileVersion>>,
    #[serde(rename = "fileSharing")]
    #[serde(default)]
    pub file_sharing: Option<Box<FileSharing>>,
    #[serde(rename = "workbookPr")]
    #[serde(default)]
    pub workbook_pr: Option<Box<WorkbookProperties>>,
    #[serde(rename = "workbookProtection")]
    #[serde(default)]
    pub workbook_protection: Option<Box<WorkbookProtection>>,
    #[serde(rename = "bookViews")]
    #[serde(default)]
    pub book_views: Option<Box<BookViews>>,
    #[serde(rename = "sheets")]
    pub sheets: Box<Sheets>,
    #[serde(rename = "functionGroups")]
    #[serde(default)]
    pub function_groups: Option<Box<CTFunctionGroups>>,
    #[serde(rename = "externalReferences")]
    #[serde(default)]
    pub external_references: Option<Box<ExternalReferences>>,
    #[serde(rename = "definedNames")]
    #[serde(default)]
    pub defined_names: Option<Box<DefinedNames>>,
    #[serde(rename = "calcPr")]
    #[serde(default)]
    pub calc_pr: Option<Box<CalculationProperties>>,
    #[serde(rename = "oleSize")]
    #[serde(default)]
    pub ole_size: Option<Box<CTOleSize>>,
    #[serde(rename = "customWorkbookViews")]
    #[serde(default)]
    pub custom_workbook_views: Option<Box<CustomWorkbookViews>>,
    #[serde(rename = "pivotCaches")]
    #[serde(default)]
    pub pivot_caches: Option<Box<PivotCaches>>,
    #[serde(rename = "smartTagPr")]
    #[serde(default)]
    pub smart_tag_pr: Option<Box<CTSmartTagPr>>,
    #[serde(rename = "smartTagTypes")]
    #[serde(default)]
    pub smart_tag_types: Option<Box<CTSmartTagTypes>>,
    #[serde(rename = "webPublishing")]
    #[serde(default)]
    pub web_publishing: Option<Box<WebPublishing>>,
    #[serde(rename = "fileRecoveryPr")]
    #[serde(default)]
    pub file_recovery_pr: Vec<Box<FileRecoveryProperties>>,
    #[serde(rename = "webPublishObjects")]
    #[serde(default)]
    pub web_publish_objects: Option<Box<CTWebPublishObjects>>,
    #[serde(rename = "extLst")]
    #[serde(default)]
    pub extension_list: Option<Box<ExtensionList>>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct FileVersion {
    #[serde(rename = "@appName")]
    #[serde(default)]
    pub app_name: Option<String>,
    #[serde(rename = "@lastEdited")]
    #[serde(default)]
    pub last_edited: Option<String>,
    #[serde(rename = "@lowestEdited")]
    #[serde(default)]
    pub lowest_edited: Option<String>,
    #[serde(rename = "@rupBuild")]
    #[serde(default)]
    pub rup_build: Option<String>,
    #[serde(rename = "@codeName")]
    #[serde(default)]
    pub code_name: Option<Guid>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct BookViews {
    #[serde(rename = "workbookView")]
    #[serde(default)]
    pub workbook_view: Vec<Box<BookView>>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct BookView {
    #[serde(rename = "@visibility")]
    #[serde(default)]
    pub visibility: Option<Visibility>,
    #[serde(rename = "@minimized")]
    #[serde(default)]
    pub minimized: Option<bool>,
    #[serde(rename = "@showHorizontalScroll")]
    #[serde(default)]
    pub show_horizontal_scroll: Option<bool>,
    #[serde(rename = "@showVerticalScroll")]
    #[serde(default)]
    pub show_vertical_scroll: Option<bool>,
    #[serde(rename = "@showSheetTabs")]
    #[serde(default)]
    pub show_sheet_tabs: Option<bool>,
    #[serde(rename = "@xWindow")]
    #[serde(default)]
    pub x_window: Option<i32>,
    #[serde(rename = "@yWindow")]
    #[serde(default)]
    pub y_window: Option<i32>,
    #[serde(rename = "@windowWidth")]
    #[serde(default)]
    pub window_width: Option<u32>,
    #[serde(rename = "@windowHeight")]
    #[serde(default)]
    pub window_height: Option<u32>,
    #[serde(rename = "@tabRatio")]
    #[serde(default)]
    pub tab_ratio: Option<u32>,
    #[serde(rename = "@firstSheet")]
    #[serde(default)]
    pub first_sheet: Option<u32>,
    #[serde(rename = "@activeTab")]
    #[serde(default)]
    pub active_tab: Option<u32>,
    #[serde(rename = "@autoFilterDateGrouping")]
    #[serde(default)]
    pub auto_filter_date_grouping: Option<bool>,
    #[serde(rename = "extLst")]
    #[serde(default)]
    pub extension_list: Option<Box<ExtensionList>>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CustomWorkbookViews {
    #[serde(rename = "customWorkbookView")]
    #[serde(default)]
    pub custom_workbook_view: Vec<Box<CustomWorkbookView>>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CustomWorkbookView {
    #[serde(rename = "@name")]
    pub name: XmlString,
    #[serde(rename = "@guid")]
    pub guid: Guid,
    #[serde(rename = "@autoUpdate")]
    #[serde(default)]
    pub auto_update: Option<bool>,
    #[serde(rename = "@mergeInterval")]
    #[serde(default)]
    pub merge_interval: Option<u32>,
    #[serde(rename = "@changesSavedWin")]
    #[serde(default)]
    pub changes_saved_win: Option<bool>,
    #[serde(rename = "@onlySync")]
    #[serde(default)]
    pub only_sync: Option<bool>,
    #[serde(rename = "@personalView")]
    #[serde(default)]
    pub personal_view: Option<bool>,
    #[serde(rename = "@includePrintSettings")]
    #[serde(default)]
    pub include_print_settings: Option<bool>,
    #[serde(rename = "@includeHiddenRowCol")]
    #[serde(default)]
    pub include_hidden_row_col: Option<bool>,
    #[serde(rename = "@maximized")]
    #[serde(default)]
    pub maximized: Option<bool>,
    #[serde(rename = "@minimized")]
    #[serde(default)]
    pub minimized: Option<bool>,
    #[serde(rename = "@showHorizontalScroll")]
    #[serde(default)]
    pub show_horizontal_scroll: Option<bool>,
    #[serde(rename = "@showVerticalScroll")]
    #[serde(default)]
    pub show_vertical_scroll: Option<bool>,
    #[serde(rename = "@showSheetTabs")]
    #[serde(default)]
    pub show_sheet_tabs: Option<bool>,
    #[serde(rename = "@xWindow")]
    #[serde(default)]
    pub x_window: Option<i32>,
    #[serde(rename = "@yWindow")]
    #[serde(default)]
    pub y_window: Option<i32>,
    #[serde(rename = "@windowWidth")]
    pub window_width: u32,
    #[serde(rename = "@windowHeight")]
    pub window_height: u32,
    #[serde(rename = "@tabRatio")]
    #[serde(default)]
    pub tab_ratio: Option<u32>,
    #[serde(rename = "@activeSheetId")]
    pub active_sheet_id: u32,
    #[serde(rename = "@showFormulaBar")]
    #[serde(default)]
    pub show_formula_bar: Option<bool>,
    #[serde(rename = "@showStatusbar")]
    #[serde(default)]
    pub show_statusbar: Option<bool>,
    #[serde(rename = "@showComments")]
    #[serde(default)]
    pub show_comments: Option<CommentVisibility>,
    #[serde(rename = "@showObjects")]
    #[serde(default)]
    pub show_objects: Option<ObjectVisibility>,
    #[serde(rename = "extLst")]
    #[serde(default)]
    pub extension_list: Option<Box<ExtensionList>>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct Sheets {
    #[serde(rename = "sheet")]
    #[serde(default)]
    pub sheet: Vec<Box<Sheet>>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct Sheet {
    #[serde(rename = "@name")]
    pub name: XmlString,
    #[serde(rename = "@sheetId")]
    pub sheet_id: u32,
    #[serde(rename = "@state")]
    #[serde(default)]
    pub state: Option<SheetState>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct WorkbookProperties {
    #[serde(rename = "@date1904")]
    #[serde(default)]
    pub date1904: Option<bool>,
    #[serde(rename = "@showObjects")]
    #[serde(default)]
    pub show_objects: Option<ObjectVisibility>,
    #[serde(rename = "@showBorderUnselectedTables")]
    #[serde(default)]
    pub show_border_unselected_tables: Option<bool>,
    #[serde(rename = "@filterPrivacy")]
    #[serde(default)]
    pub filter_privacy: Option<bool>,
    #[serde(rename = "@promptedSolutions")]
    #[serde(default)]
    pub prompted_solutions: Option<bool>,
    #[serde(rename = "@showInkAnnotation")]
    #[serde(default)]
    pub show_ink_annotation: Option<bool>,
    #[serde(rename = "@backupFile")]
    #[serde(default)]
    pub backup_file: Option<bool>,
    #[serde(rename = "@saveExternalLinkValues")]
    #[serde(default)]
    pub save_external_link_values: Option<bool>,
    #[serde(rename = "@updateLinks")]
    #[serde(default)]
    pub update_links: Option<UpdateLinks>,
    #[serde(rename = "@codeName")]
    #[serde(default)]
    pub code_name: Option<String>,
    #[serde(rename = "@hidePivotFieldList")]
    #[serde(default)]
    pub hide_pivot_field_list: Option<bool>,
    #[serde(rename = "@showPivotChartFilter")]
    #[serde(default)]
    pub show_pivot_chart_filter: Option<bool>,
    #[serde(rename = "@allowRefreshQuery")]
    #[serde(default)]
    pub allow_refresh_query: Option<bool>,
    #[serde(rename = "@publishItems")]
    #[serde(default)]
    pub publish_items: Option<bool>,
    #[serde(rename = "@checkCompatibility")]
    #[serde(default)]
    pub check_compatibility: Option<bool>,
    #[serde(rename = "@autoCompressPictures")]
    #[serde(default)]
    pub auto_compress_pictures: Option<bool>,
    #[serde(rename = "@refreshAllConnections")]
    #[serde(default)]
    pub refresh_all_connections: Option<bool>,
    #[serde(rename = "@defaultThemeVersion")]
    #[serde(default)]
    pub default_theme_version: Option<u32>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTSmartTagPr {
    #[serde(rename = "@embed")]
    #[serde(default)]
    pub embed: Option<bool>,
    #[serde(rename = "@show")]
    #[serde(default)]
    pub show: Option<STSmartTagShow>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTSmartTagTypes {
    #[serde(rename = "smartTagType")]
    #[serde(default)]
    pub smart_tag_type: Vec<Box<CTSmartTagType>>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTSmartTagType {
    #[serde(rename = "@namespaceUri")]
    #[serde(default)]
    pub namespace_uri: Option<XmlString>,
    #[serde(rename = "@name")]
    #[serde(default)]
    pub name: Option<XmlString>,
    #[serde(rename = "@url")]
    #[serde(default)]
    pub url: Option<XmlString>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct FileRecoveryProperties {
    #[serde(rename = "@autoRecover")]
    #[serde(default)]
    pub auto_recover: Option<bool>,
    #[serde(rename = "@crashSave")]
    #[serde(default)]
    pub crash_save: Option<bool>,
    #[serde(rename = "@dataExtractLoad")]
    #[serde(default)]
    pub data_extract_load: Option<bool>,
    #[serde(rename = "@repairLoad")]
    #[serde(default)]
    pub repair_load: Option<bool>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CalculationProperties {
    #[serde(rename = "@calcId")]
    #[serde(default)]
    pub calc_id: Option<u32>,
    #[serde(rename = "@calcMode")]
    #[serde(default)]
    pub calc_mode: Option<CalculationMode>,
    #[serde(rename = "@fullCalcOnLoad")]
    #[serde(default)]
    pub full_calc_on_load: Option<bool>,
    #[serde(rename = "@refMode")]
    #[serde(default)]
    pub ref_mode: Option<ReferenceMode>,
    #[serde(rename = "@iterate")]
    #[serde(default)]
    pub iterate: Option<bool>,
    #[serde(rename = "@iterateCount")]
    #[serde(default)]
    pub iterate_count: Option<u32>,
    #[serde(rename = "@iterateDelta")]
    #[serde(default)]
    pub iterate_delta: Option<f64>,
    #[serde(rename = "@fullPrecision")]
    #[serde(default)]
    pub full_precision: Option<bool>,
    #[serde(rename = "@calcCompleted")]
    #[serde(default)]
    pub calc_completed: Option<bool>,
    #[serde(rename = "@calcOnSave")]
    #[serde(default)]
    pub calc_on_save: Option<bool>,
    #[serde(rename = "@concurrentCalc")]
    #[serde(default)]
    pub concurrent_calc: Option<bool>,
    #[serde(rename = "@concurrentManualCount")]
    #[serde(default)]
    pub concurrent_manual_count: Option<u32>,
    #[serde(rename = "@forceFullCalc")]
    #[serde(default)]
    pub force_full_calc: Option<bool>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct DefinedNames {
    #[serde(rename = "definedName")]
    #[serde(default)]
    pub defined_name: Vec<Box<DefinedName>>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct DefinedName {
    #[serde(rename = "$text")]
    pub text: String,
    #[serde(rename = "@name")]
    pub name: XmlString,
    #[serde(rename = "@comment")]
    #[serde(default)]
    pub comment: Option<XmlString>,
    #[serde(rename = "@customMenu")]
    #[serde(default)]
    pub custom_menu: Option<XmlString>,
    #[serde(rename = "@description")]
    #[serde(default)]
    pub description: Option<XmlString>,
    #[serde(rename = "@help")]
    #[serde(default)]
    pub help: Option<XmlString>,
    #[serde(rename = "@statusBar")]
    #[serde(default)]
    pub status_bar: Option<XmlString>,
    #[serde(rename = "@localSheetId")]
    #[serde(default)]
    pub local_sheet_id: Option<u32>,
    #[serde(rename = "@hidden")]
    #[serde(default)]
    pub hidden: Option<bool>,
    #[serde(rename = "@function")]
    #[serde(default)]
    pub function: Option<bool>,
    #[serde(rename = "@vbProcedure")]
    #[serde(default)]
    pub vb_procedure: Option<bool>,
    #[serde(rename = "@xlm")]
    #[serde(default)]
    pub xlm: Option<bool>,
    #[serde(rename = "@functionGroupId")]
    #[serde(default)]
    pub function_group_id: Option<u32>,
    #[serde(rename = "@shortcutKey")]
    #[serde(default)]
    pub shortcut_key: Option<XmlString>,
    #[serde(rename = "@publishToServer")]
    #[serde(default)]
    pub publish_to_server: Option<bool>,
    #[serde(rename = "@workbookParameter")]
    #[serde(default)]
    pub workbook_parameter: Option<bool>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct ExternalReferences {
    #[serde(rename = "externalReference")]
    #[serde(default)]
    pub external_reference: Vec<Box<ExternalReference>>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct ExternalReference;

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct SheetBackgroundPicture;

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct PivotCaches {
    #[serde(rename = "pivotCache")]
    #[serde(default)]
    pub pivot_cache: Vec<Box<CTPivotCache>>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTPivotCache {
    #[serde(rename = "@cacheId")]
    pub cache_id: u32,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct FileSharing {
    #[serde(rename = "@readOnlyRecommended")]
    #[serde(default)]
    pub read_only_recommended: Option<bool>,
    #[serde(rename = "@userName")]
    #[serde(default)]
    pub user_name: Option<XmlString>,
    #[serde(rename = "@reservationPassword")]
    #[serde(default)]
    pub reservation_password: Option<STUnsignedShortHex>,
    #[serde(rename = "@algorithmName")]
    #[serde(default)]
    pub algorithm_name: Option<XmlString>,
    #[serde(rename = "@hashValue")]
    #[serde(default)]
    pub hash_value: Option<Vec<u8>>,
    #[serde(rename = "@saltValue")]
    #[serde(default)]
    pub salt_value: Option<Vec<u8>>,
    #[serde(rename = "@spinCount")]
    #[serde(default)]
    pub spin_count: Option<u32>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTOleSize {
    #[serde(rename = "@ref")]
    pub reference: Reference,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct WorkbookProtection {
    #[serde(rename = "@workbookPassword")]
    #[serde(default)]
    pub workbook_password: Option<STUnsignedShortHex>,
    #[serde(rename = "@workbookPasswordCharacterSet")]
    #[serde(default)]
    pub workbook_password_character_set: Option<String>,
    #[serde(rename = "@revisionsPassword")]
    #[serde(default)]
    pub revisions_password: Option<STUnsignedShortHex>,
    #[serde(rename = "@revisionsPasswordCharacterSet")]
    #[serde(default)]
    pub revisions_password_character_set: Option<String>,
    #[serde(rename = "@lockStructure")]
    #[serde(default)]
    pub lock_structure: Option<bool>,
    #[serde(rename = "@lockWindows")]
    #[serde(default)]
    pub lock_windows: Option<bool>,
    #[serde(rename = "@lockRevision")]
    #[serde(default)]
    pub lock_revision: Option<bool>,
    #[serde(rename = "@revisionsAlgorithmName")]
    #[serde(default)]
    pub revisions_algorithm_name: Option<XmlString>,
    #[serde(rename = "@revisionsHashValue")]
    #[serde(default)]
    pub revisions_hash_value: Option<Vec<u8>>,
    #[serde(rename = "@revisionsSaltValue")]
    #[serde(default)]
    pub revisions_salt_value: Option<Vec<u8>>,
    #[serde(rename = "@revisionsSpinCount")]
    #[serde(default)]
    pub revisions_spin_count: Option<u32>,
    #[serde(rename = "@workbookAlgorithmName")]
    #[serde(default)]
    pub workbook_algorithm_name: Option<XmlString>,
    #[serde(rename = "@workbookHashValue")]
    #[serde(default)]
    pub workbook_hash_value: Option<Vec<u8>>,
    #[serde(rename = "@workbookSaltValue")]
    #[serde(default)]
    pub workbook_salt_value: Option<Vec<u8>>,
    #[serde(rename = "@workbookSpinCount")]
    #[serde(default)]
    pub workbook_spin_count: Option<u32>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct WebPublishing {
    #[serde(rename = "@css")]
    #[serde(default)]
    pub css: Option<bool>,
    #[serde(rename = "@thicket")]
    #[serde(default)]
    pub thicket: Option<bool>,
    #[serde(rename = "@longFileNames")]
    #[serde(default)]
    pub long_file_names: Option<bool>,
    #[serde(rename = "@vml")]
    #[serde(default)]
    pub vml: Option<bool>,
    #[serde(rename = "@allowPng")]
    #[serde(default)]
    pub allow_png: Option<bool>,
    #[serde(rename = "@targetScreenSize")]
    #[serde(default)]
    pub target_screen_size: Option<STTargetScreenSize>,
    #[serde(rename = "@dpi")]
    #[serde(default)]
    pub dpi: Option<u32>,
    #[serde(rename = "@codePage")]
    #[serde(default)]
    pub code_page: Option<u32>,
    #[serde(rename = "@characterSet")]
    #[serde(default)]
    pub character_set: Option<String>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTFunctionGroups {
    #[serde(rename = "@builtInGroupCount")]
    #[serde(default)]
    pub built_in_group_count: Option<u32>,
    #[serde(rename = "functionGroup")]
    #[serde(default)]
    pub function_group: Vec<Box<CTFunctionGroup>>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTFunctionGroup {
    #[serde(rename = "@name")]
    #[serde(default)]
    pub name: Option<XmlString>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTWebPublishObjects {
    #[serde(rename = "@count")]
    #[serde(default)]
    pub count: Option<u32>,
    #[serde(rename = "webPublishObject")]
    #[serde(default)]
    pub web_publish_object: Vec<Box<CTWebPublishObject>>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTWebPublishObject {
    #[serde(rename = "@id")]
    pub id: u32,
    #[serde(rename = "@divId")]
    pub div_id: XmlString,
    #[serde(rename = "@sourceObject")]
    #[serde(default)]
    pub source_object: Option<XmlString>,
    #[serde(rename = "@destinationFile")]
    pub destination_file: XmlString,
    #[serde(rename = "@title")]
    #[serde(default)]
    pub title: Option<XmlString>,
    #[serde(rename = "@autoRepublish")]
    #[serde(default)]
    pub auto_republish: Option<bool>,
}
