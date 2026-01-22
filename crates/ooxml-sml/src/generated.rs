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

pub type SSTLang = String;

pub type SSTHexColorRGB = Vec<u8>;

pub type SSTPanose = Vec<u8>;

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum SSTCalendarType {
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

impl std::fmt::Display for SSTCalendarType {
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

impl std::str::FromStr for SSTCalendarType {
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
            _ => Err(format!("unknown SSTCalendarType value: {}", s)),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum SSTAlgClass {
    #[serde(rename = "hash")]
    Hash,
    #[serde(rename = "custom")]
    Custom,
}

impl std::fmt::Display for SSTAlgClass {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::Hash => write!(f, "hash"),
            Self::Custom => write!(f, "custom"),
        }
    }
}

impl std::str::FromStr for SSTAlgClass {
    type Err = String;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        match s {
            "hash" => Ok(Self::Hash),
            "custom" => Ok(Self::Custom),
            _ => Err(format!("unknown SSTAlgClass value: {}", s)),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum SSTCryptProv {
    #[serde(rename = "rsaAES")]
    RsaAES,
    #[serde(rename = "rsaFull")]
    RsaFull,
    #[serde(rename = "custom")]
    Custom,
}

impl std::fmt::Display for SSTCryptProv {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::RsaAES => write!(f, "rsaAES"),
            Self::RsaFull => write!(f, "rsaFull"),
            Self::Custom => write!(f, "custom"),
        }
    }
}

impl std::str::FromStr for SSTCryptProv {
    type Err = String;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        match s {
            "rsaAES" => Ok(Self::RsaAES),
            "rsaFull" => Ok(Self::RsaFull),
            "custom" => Ok(Self::Custom),
            _ => Err(format!("unknown SSTCryptProv value: {}", s)),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum SSTAlgType {
    #[serde(rename = "typeAny")]
    TypeAny,
    #[serde(rename = "custom")]
    Custom,
}

impl std::fmt::Display for SSTAlgType {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::TypeAny => write!(f, "typeAny"),
            Self::Custom => write!(f, "custom"),
        }
    }
}

impl std::str::FromStr for SSTAlgType {
    type Err = String;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        match s {
            "typeAny" => Ok(Self::TypeAny),
            "custom" => Ok(Self::Custom),
            _ => Err(format!("unknown SSTAlgType value: {}", s)),
        }
    }
}

pub type SSTColorType = String;

pub type SSTGuid = String;

pub type SSTOnOff = String;

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum SSTOnOff1 {
    #[serde(rename = "on")]
    On,
    #[serde(rename = "off")]
    Off,
}

impl std::fmt::Display for SSTOnOff1 {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::On => write!(f, "on"),
            Self::Off => write!(f, "off"),
        }
    }
}

impl std::str::FromStr for SSTOnOff1 {
    type Err = String;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        match s {
            "on" => Ok(Self::On),
            "off" => Ok(Self::Off),
            _ => Err(format!("unknown SSTOnOff1 value: {}", s)),
        }
    }
}

pub type SSTString = String;

pub type SSTXmlName = String;

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum SSTTrueFalse {
    #[serde(rename = "t")]
    T,
    #[serde(rename = "f")]
    F,
    #[serde(rename = "true")]
    True,
    #[serde(rename = "false")]
    False,
}

impl std::fmt::Display for SSTTrueFalse {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::T => write!(f, "t"),
            Self::F => write!(f, "f"),
            Self::True => write!(f, "true"),
            Self::False => write!(f, "false"),
        }
    }
}

impl std::str::FromStr for SSTTrueFalse {
    type Err = String;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        match s {
            "t" => Ok(Self::T),
            "f" => Ok(Self::F),
            "true" => Ok(Self::True),
            "false" => Ok(Self::False),
            _ => Err(format!("unknown SSTTrueFalse value: {}", s)),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum SSTTrueFalseBlank {
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

impl std::fmt::Display for SSTTrueFalseBlank {
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

impl std::str::FromStr for SSTTrueFalseBlank {
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
            _ => Err(format!("unknown SSTTrueFalseBlank value: {}", s)),
        }
    }
}

pub type SSTUnsignedDecimalNumber = u64;

pub type SSTTwipsMeasure = String;

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum SSTVerticalAlignRun {
    #[serde(rename = "baseline")]
    Baseline,
    #[serde(rename = "superscript")]
    Superscript,
    #[serde(rename = "subscript")]
    Subscript,
}

impl std::fmt::Display for SSTVerticalAlignRun {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::Baseline => write!(f, "baseline"),
            Self::Superscript => write!(f, "superscript"),
            Self::Subscript => write!(f, "subscript"),
        }
    }
}

impl std::str::FromStr for SSTVerticalAlignRun {
    type Err = String;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        match s {
            "baseline" => Ok(Self::Baseline),
            "superscript" => Ok(Self::Superscript),
            "subscript" => Ok(Self::Subscript),
            _ => Err(format!("unknown SSTVerticalAlignRun value: {}", s)),
        }
    }
}

pub type SSTXstring = String;

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum SSTXAlign {
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

impl std::fmt::Display for SSTXAlign {
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

impl std::str::FromStr for SSTXAlign {
    type Err = String;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        match s {
            "left" => Ok(Self::Left),
            "center" => Ok(Self::Center),
            "right" => Ok(Self::Right),
            "inside" => Ok(Self::Inside),
            "outside" => Ok(Self::Outside),
            _ => Err(format!("unknown SSTXAlign value: {}", s)),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum SSTYAlign {
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

impl std::fmt::Display for SSTYAlign {
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

impl std::str::FromStr for SSTYAlign {
    type Err = String;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        match s {
            "inline" => Ok(Self::Inline),
            "top" => Ok(Self::Top),
            "center" => Ok(Self::Center),
            "bottom" => Ok(Self::Bottom),
            "inside" => Ok(Self::Inside),
            "outside" => Ok(Self::Outside),
            _ => Err(format!("unknown SSTYAlign value: {}", s)),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum SSTConformanceClass {
    #[serde(rename = "strict")]
    Strict,
    #[serde(rename = "transitional")]
    Transitional,
}

impl std::fmt::Display for SSTConformanceClass {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::Strict => write!(f, "strict"),
            Self::Transitional => write!(f, "transitional"),
        }
    }
}

impl std::str::FromStr for SSTConformanceClass {
    type Err = String;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        match s {
            "strict" => Ok(Self::Strict),
            "transitional" => Ok(Self::Transitional),
            _ => Err(format!("unknown SSTConformanceClass value: {}", s)),
        }
    }
}

pub type SSTUniversalMeasure = String;

pub type SSTPositiveUniversalMeasure = String;

pub type SSTPercentage = String;

pub type SSTFixedPercentage = String;

pub type SSTPositivePercentage = String;

pub type SSTPositiveFixedPercentage = String;

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum SmlSTFilterOperator {
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

impl std::fmt::Display for SmlSTFilterOperator {
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

impl std::str::FromStr for SmlSTFilterOperator {
    type Err = String;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        match s {
            "equal" => Ok(Self::Equal),
            "lessThan" => Ok(Self::LessThan),
            "lessThanOrEqual" => Ok(Self::LessThanOrEqual),
            "notEqual" => Ok(Self::NotEqual),
            "greaterThanOrEqual" => Ok(Self::GreaterThanOrEqual),
            "greaterThan" => Ok(Self::GreaterThan),
            _ => Err(format!("unknown SmlSTFilterOperator value: {}", s)),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum SmlSTDynamicFilterType {
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

impl std::fmt::Display for SmlSTDynamicFilterType {
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

impl std::str::FromStr for SmlSTDynamicFilterType {
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
            _ => Err(format!("unknown SmlSTDynamicFilterType value: {}", s)),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum SmlSTIconSetType {
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

impl std::fmt::Display for SmlSTIconSetType {
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

impl std::str::FromStr for SmlSTIconSetType {
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
            _ => Err(format!("unknown SmlSTIconSetType value: {}", s)),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum SmlSTSortBy {
    #[serde(rename = "value")]
    Value,
    #[serde(rename = "cellColor")]
    CellColor,
    #[serde(rename = "fontColor")]
    FontColor,
    #[serde(rename = "icon")]
    Icon,
}

impl std::fmt::Display for SmlSTSortBy {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::Value => write!(f, "value"),
            Self::CellColor => write!(f, "cellColor"),
            Self::FontColor => write!(f, "fontColor"),
            Self::Icon => write!(f, "icon"),
        }
    }
}

impl std::str::FromStr for SmlSTSortBy {
    type Err = String;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        match s {
            "value" => Ok(Self::Value),
            "cellColor" => Ok(Self::CellColor),
            "fontColor" => Ok(Self::FontColor),
            "icon" => Ok(Self::Icon),
            _ => Err(format!("unknown SmlSTSortBy value: {}", s)),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum SmlSTSortMethod {
    #[serde(rename = "stroke")]
    Stroke,
    #[serde(rename = "pinYin")]
    PinYin,
    #[serde(rename = "none")]
    None,
}

impl std::fmt::Display for SmlSTSortMethod {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::Stroke => write!(f, "stroke"),
            Self::PinYin => write!(f, "pinYin"),
            Self::None => write!(f, "none"),
        }
    }
}

impl std::str::FromStr for SmlSTSortMethod {
    type Err = String;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        match s {
            "stroke" => Ok(Self::Stroke),
            "pinYin" => Ok(Self::PinYin),
            "none" => Ok(Self::None),
            _ => Err(format!("unknown SmlSTSortMethod value: {}", s)),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum SmlSTDateTimeGrouping {
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

impl std::fmt::Display for SmlSTDateTimeGrouping {
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

impl std::str::FromStr for SmlSTDateTimeGrouping {
    type Err = String;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        match s {
            "year" => Ok(Self::Year),
            "month" => Ok(Self::Month),
            "day" => Ok(Self::Day),
            "hour" => Ok(Self::Hour),
            "minute" => Ok(Self::Minute),
            "second" => Ok(Self::Second),
            _ => Err(format!("unknown SmlSTDateTimeGrouping value: {}", s)),
        }
    }
}

pub type SmlSTCellRef = String;

pub type SmlSTRef = String;

pub type SmlSTRefA = String;

pub type SmlSTSqref = String;

pub type SmlSTFormula = SSTXstring;

pub type SmlSTUnsignedIntHex = Vec<u8>;

pub type SmlSTUnsignedShortHex = Vec<u8>;

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum SmlSTTextHAlign {
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

impl std::fmt::Display for SmlSTTextHAlign {
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

impl std::str::FromStr for SmlSTTextHAlign {
    type Err = String;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        match s {
            "left" => Ok(Self::Left),
            "center" => Ok(Self::Center),
            "right" => Ok(Self::Right),
            "justify" => Ok(Self::Justify),
            "distributed" => Ok(Self::Distributed),
            _ => Err(format!("unknown SmlSTTextHAlign value: {}", s)),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum SmlSTTextVAlign {
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

impl std::fmt::Display for SmlSTTextVAlign {
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

impl std::str::FromStr for SmlSTTextVAlign {
    type Err = String;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        match s {
            "top" => Ok(Self::Top),
            "center" => Ok(Self::Center),
            "bottom" => Ok(Self::Bottom),
            "justify" => Ok(Self::Justify),
            "distributed" => Ok(Self::Distributed),
            _ => Err(format!("unknown SmlSTTextVAlign value: {}", s)),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum SmlSTCredMethod {
    #[serde(rename = "integrated")]
    Integrated,
    #[serde(rename = "none")]
    None,
    #[serde(rename = "stored")]
    Stored,
    #[serde(rename = "prompt")]
    Prompt,
}

impl std::fmt::Display for SmlSTCredMethod {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::Integrated => write!(f, "integrated"),
            Self::None => write!(f, "none"),
            Self::Stored => write!(f, "stored"),
            Self::Prompt => write!(f, "prompt"),
        }
    }
}

impl std::str::FromStr for SmlSTCredMethod {
    type Err = String;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        match s {
            "integrated" => Ok(Self::Integrated),
            "none" => Ok(Self::None),
            "stored" => Ok(Self::Stored),
            "prompt" => Ok(Self::Prompt),
            _ => Err(format!("unknown SmlSTCredMethod value: {}", s)),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum SmlSTHtmlFmt {
    #[serde(rename = "none")]
    None,
    #[serde(rename = "rtf")]
    Rtf,
    #[serde(rename = "all")]
    All,
}

impl std::fmt::Display for SmlSTHtmlFmt {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::None => write!(f, "none"),
            Self::Rtf => write!(f, "rtf"),
            Self::All => write!(f, "all"),
        }
    }
}

impl std::str::FromStr for SmlSTHtmlFmt {
    type Err = String;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        match s {
            "none" => Ok(Self::None),
            "rtf" => Ok(Self::Rtf),
            "all" => Ok(Self::All),
            _ => Err(format!("unknown SmlSTHtmlFmt value: {}", s)),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum SmlSTParameterType {
    #[serde(rename = "prompt")]
    Prompt,
    #[serde(rename = "value")]
    Value,
    #[serde(rename = "cell")]
    Cell,
}

impl std::fmt::Display for SmlSTParameterType {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::Prompt => write!(f, "prompt"),
            Self::Value => write!(f, "value"),
            Self::Cell => write!(f, "cell"),
        }
    }
}

impl std::str::FromStr for SmlSTParameterType {
    type Err = String;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        match s {
            "prompt" => Ok(Self::Prompt),
            "value" => Ok(Self::Value),
            "cell" => Ok(Self::Cell),
            _ => Err(format!("unknown SmlSTParameterType value: {}", s)),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum SmlSTFileType {
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

impl std::fmt::Display for SmlSTFileType {
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

impl std::str::FromStr for SmlSTFileType {
    type Err = String;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        match s {
            "mac" => Ok(Self::Mac),
            "win" => Ok(Self::Win),
            "dos" => Ok(Self::Dos),
            "lin" => Ok(Self::Lin),
            "other" => Ok(Self::Other),
            _ => Err(format!("unknown SmlSTFileType value: {}", s)),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum SmlSTQualifier {
    #[serde(rename = "doubleQuote")]
    DoubleQuote,
    #[serde(rename = "singleQuote")]
    SingleQuote,
    #[serde(rename = "none")]
    None,
}

impl std::fmt::Display for SmlSTQualifier {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::DoubleQuote => write!(f, "doubleQuote"),
            Self::SingleQuote => write!(f, "singleQuote"),
            Self::None => write!(f, "none"),
        }
    }
}

impl std::str::FromStr for SmlSTQualifier {
    type Err = String;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        match s {
            "doubleQuote" => Ok(Self::DoubleQuote),
            "singleQuote" => Ok(Self::SingleQuote),
            "none" => Ok(Self::None),
            _ => Err(format!("unknown SmlSTQualifier value: {}", s)),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum SmlSTExternalConnectionType {
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

impl std::fmt::Display for SmlSTExternalConnectionType {
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

impl std::str::FromStr for SmlSTExternalConnectionType {
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
            _ => Err(format!("unknown SmlSTExternalConnectionType value: {}", s)),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum SmlSTSourceType {
    #[serde(rename = "worksheet")]
    Worksheet,
    #[serde(rename = "external")]
    External,
    #[serde(rename = "consolidation")]
    Consolidation,
    #[serde(rename = "scenario")]
    Scenario,
}

impl std::fmt::Display for SmlSTSourceType {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::Worksheet => write!(f, "worksheet"),
            Self::External => write!(f, "external"),
            Self::Consolidation => write!(f, "consolidation"),
            Self::Scenario => write!(f, "scenario"),
        }
    }
}

impl std::str::FromStr for SmlSTSourceType {
    type Err = String;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        match s {
            "worksheet" => Ok(Self::Worksheet),
            "external" => Ok(Self::External),
            "consolidation" => Ok(Self::Consolidation),
            "scenario" => Ok(Self::Scenario),
            _ => Err(format!("unknown SmlSTSourceType value: {}", s)),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum SmlSTGroupBy {
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

impl std::fmt::Display for SmlSTGroupBy {
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

impl std::str::FromStr for SmlSTGroupBy {
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
            _ => Err(format!("unknown SmlSTGroupBy value: {}", s)),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum SmlSTSortType {
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

impl std::fmt::Display for SmlSTSortType {
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

impl std::str::FromStr for SmlSTSortType {
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
            _ => Err(format!("unknown SmlSTSortType value: {}", s)),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum SmlSTScope {
    #[serde(rename = "selection")]
    Selection,
    #[serde(rename = "data")]
    Data,
    #[serde(rename = "field")]
    Field,
}

impl std::fmt::Display for SmlSTScope {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::Selection => write!(f, "selection"),
            Self::Data => write!(f, "data"),
            Self::Field => write!(f, "field"),
        }
    }
}

impl std::str::FromStr for SmlSTScope {
    type Err = String;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        match s {
            "selection" => Ok(Self::Selection),
            "data" => Ok(Self::Data),
            "field" => Ok(Self::Field),
            _ => Err(format!("unknown SmlSTScope value: {}", s)),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum SmlSTType {
    #[serde(rename = "none")]
    None,
    #[serde(rename = "all")]
    All,
    #[serde(rename = "row")]
    Row,
    #[serde(rename = "column")]
    Column,
}

impl std::fmt::Display for SmlSTType {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::None => write!(f, "none"),
            Self::All => write!(f, "all"),
            Self::Row => write!(f, "row"),
            Self::Column => write!(f, "column"),
        }
    }
}

impl std::str::FromStr for SmlSTType {
    type Err = String;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        match s {
            "none" => Ok(Self::None),
            "all" => Ok(Self::All),
            "row" => Ok(Self::Row),
            "column" => Ok(Self::Column),
            _ => Err(format!("unknown SmlSTType value: {}", s)),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum SmlSTShowDataAs {
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

impl std::fmt::Display for SmlSTShowDataAs {
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

impl std::str::FromStr for SmlSTShowDataAs {
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
            _ => Err(format!("unknown SmlSTShowDataAs value: {}", s)),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum SmlSTItemType {
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

impl std::fmt::Display for SmlSTItemType {
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

impl std::str::FromStr for SmlSTItemType {
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
            _ => Err(format!("unknown SmlSTItemType value: {}", s)),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum SmlSTFormatAction {
    #[serde(rename = "blank")]
    Blank,
    #[serde(rename = "formatting")]
    Formatting,
    #[serde(rename = "drill")]
    Drill,
    #[serde(rename = "formula")]
    Formula,
}

impl std::fmt::Display for SmlSTFormatAction {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::Blank => write!(f, "blank"),
            Self::Formatting => write!(f, "formatting"),
            Self::Drill => write!(f, "drill"),
            Self::Formula => write!(f, "formula"),
        }
    }
}

impl std::str::FromStr for SmlSTFormatAction {
    type Err = String;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        match s {
            "blank" => Ok(Self::Blank),
            "formatting" => Ok(Self::Formatting),
            "drill" => Ok(Self::Drill),
            "formula" => Ok(Self::Formula),
            _ => Err(format!("unknown SmlSTFormatAction value: {}", s)),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum SmlSTFieldSortType {
    #[serde(rename = "manual")]
    Manual,
    #[serde(rename = "ascending")]
    Ascending,
    #[serde(rename = "descending")]
    Descending,
}

impl std::fmt::Display for SmlSTFieldSortType {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::Manual => write!(f, "manual"),
            Self::Ascending => write!(f, "ascending"),
            Self::Descending => write!(f, "descending"),
        }
    }
}

impl std::str::FromStr for SmlSTFieldSortType {
    type Err = String;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        match s {
            "manual" => Ok(Self::Manual),
            "ascending" => Ok(Self::Ascending),
            "descending" => Ok(Self::Descending),
            _ => Err(format!("unknown SmlSTFieldSortType value: {}", s)),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum SmlSTPivotFilterType {
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

impl std::fmt::Display for SmlSTPivotFilterType {
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

impl std::str::FromStr for SmlSTPivotFilterType {
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
            _ => Err(format!("unknown SmlSTPivotFilterType value: {}", s)),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum SmlSTPivotAreaType {
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

impl std::fmt::Display for SmlSTPivotAreaType {
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

impl std::str::FromStr for SmlSTPivotAreaType {
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
            _ => Err(format!("unknown SmlSTPivotAreaType value: {}", s)),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum SmlSTAxis {
    #[serde(rename = "axisRow")]
    AxisRow,
    #[serde(rename = "axisCol")]
    AxisCol,
    #[serde(rename = "axisPage")]
    AxisPage,
    #[serde(rename = "axisValues")]
    AxisValues,
}

impl std::fmt::Display for SmlSTAxis {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::AxisRow => write!(f, "axisRow"),
            Self::AxisCol => write!(f, "axisCol"),
            Self::AxisPage => write!(f, "axisPage"),
            Self::AxisValues => write!(f, "axisValues"),
        }
    }
}

impl std::str::FromStr for SmlSTAxis {
    type Err = String;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        match s {
            "axisRow" => Ok(Self::AxisRow),
            "axisCol" => Ok(Self::AxisCol),
            "axisPage" => Ok(Self::AxisPage),
            "axisValues" => Ok(Self::AxisValues),
            _ => Err(format!("unknown SmlSTAxis value: {}", s)),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum SmlSTGrowShrinkType {
    #[serde(rename = "insertDelete")]
    InsertDelete,
    #[serde(rename = "insertClear")]
    InsertClear,
    #[serde(rename = "overwriteClear")]
    OverwriteClear,
}

impl std::fmt::Display for SmlSTGrowShrinkType {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::InsertDelete => write!(f, "insertDelete"),
            Self::InsertClear => write!(f, "insertClear"),
            Self::OverwriteClear => write!(f, "overwriteClear"),
        }
    }
}

impl std::str::FromStr for SmlSTGrowShrinkType {
    type Err = String;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        match s {
            "insertDelete" => Ok(Self::InsertDelete),
            "insertClear" => Ok(Self::InsertClear),
            "overwriteClear" => Ok(Self::OverwriteClear),
            _ => Err(format!("unknown SmlSTGrowShrinkType value: {}", s)),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum SmlSTPhoneticType {
    #[serde(rename = "halfwidthKatakana")]
    HalfwidthKatakana,
    #[serde(rename = "fullwidthKatakana")]
    FullwidthKatakana,
    #[serde(rename = "Hiragana")]
    Hiragana,
    #[serde(rename = "noConversion")]
    NoConversion,
}

impl std::fmt::Display for SmlSTPhoneticType {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::HalfwidthKatakana => write!(f, "halfwidthKatakana"),
            Self::FullwidthKatakana => write!(f, "fullwidthKatakana"),
            Self::Hiragana => write!(f, "Hiragana"),
            Self::NoConversion => write!(f, "noConversion"),
        }
    }
}

impl std::str::FromStr for SmlSTPhoneticType {
    type Err = String;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        match s {
            "halfwidthKatakana" => Ok(Self::HalfwidthKatakana),
            "fullwidthKatakana" => Ok(Self::FullwidthKatakana),
            "Hiragana" => Ok(Self::Hiragana),
            "noConversion" => Ok(Self::NoConversion),
            _ => Err(format!("unknown SmlSTPhoneticType value: {}", s)),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum SmlSTPhoneticAlignment {
    #[serde(rename = "noControl")]
    NoControl,
    #[serde(rename = "left")]
    Left,
    #[serde(rename = "center")]
    Center,
    #[serde(rename = "distributed")]
    Distributed,
}

impl std::fmt::Display for SmlSTPhoneticAlignment {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::NoControl => write!(f, "noControl"),
            Self::Left => write!(f, "left"),
            Self::Center => write!(f, "center"),
            Self::Distributed => write!(f, "distributed"),
        }
    }
}

impl std::str::FromStr for SmlSTPhoneticAlignment {
    type Err = String;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        match s {
            "noControl" => Ok(Self::NoControl),
            "left" => Ok(Self::Left),
            "center" => Ok(Self::Center),
            "distributed" => Ok(Self::Distributed),
            _ => Err(format!("unknown SmlSTPhoneticAlignment value: {}", s)),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum SmlSTRwColActionType {
    #[serde(rename = "insertRow")]
    InsertRow,
    #[serde(rename = "deleteRow")]
    DeleteRow,
    #[serde(rename = "insertCol")]
    InsertCol,
    #[serde(rename = "deleteCol")]
    DeleteCol,
}

impl std::fmt::Display for SmlSTRwColActionType {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::InsertRow => write!(f, "insertRow"),
            Self::DeleteRow => write!(f, "deleteRow"),
            Self::InsertCol => write!(f, "insertCol"),
            Self::DeleteCol => write!(f, "deleteCol"),
        }
    }
}

impl std::str::FromStr for SmlSTRwColActionType {
    type Err = String;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        match s {
            "insertRow" => Ok(Self::InsertRow),
            "deleteRow" => Ok(Self::DeleteRow),
            "insertCol" => Ok(Self::InsertCol),
            "deleteCol" => Ok(Self::DeleteCol),
            _ => Err(format!("unknown SmlSTRwColActionType value: {}", s)),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum SmlSTRevisionAction {
    #[serde(rename = "add")]
    Add,
    #[serde(rename = "delete")]
    Delete,
}

impl std::fmt::Display for SmlSTRevisionAction {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::Add => write!(f, "add"),
            Self::Delete => write!(f, "delete"),
        }
    }
}

impl std::str::FromStr for SmlSTRevisionAction {
    type Err = String;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        match s {
            "add" => Ok(Self::Add),
            "delete" => Ok(Self::Delete),
            _ => Err(format!("unknown SmlSTRevisionAction value: {}", s)),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum SmlSTFormulaExpression {
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

impl std::fmt::Display for SmlSTFormulaExpression {
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

impl std::str::FromStr for SmlSTFormulaExpression {
    type Err = String;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        match s {
            "ref" => Ok(Self::Ref),
            "refError" => Ok(Self::RefError),
            "area" => Ok(Self::Area),
            "areaError" => Ok(Self::AreaError),
            "computedArea" => Ok(Self::ComputedArea),
            _ => Err(format!("unknown SmlSTFormulaExpression value: {}", s)),
        }
    }
}

pub type SmlSTCellSpan = String;

pub type SmlSTCellSpans = String;

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum SmlSTCellType {
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

impl std::fmt::Display for SmlSTCellType {
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

impl std::str::FromStr for SmlSTCellType {
    type Err = String;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        match s {
            "b" => Ok(Self::B),
            "n" => Ok(Self::N),
            "e" => Ok(Self::E),
            "s" => Ok(Self::S),
            "str" => Ok(Self::Str),
            "inlineStr" => Ok(Self::InlineStr),
            _ => Err(format!("unknown SmlSTCellType value: {}", s)),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum SmlSTCellFormulaType {
    #[serde(rename = "normal")]
    Normal,
    #[serde(rename = "array")]
    Array,
    #[serde(rename = "dataTable")]
    DataTable,
    #[serde(rename = "shared")]
    Shared,
}

impl std::fmt::Display for SmlSTCellFormulaType {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::Normal => write!(f, "normal"),
            Self::Array => write!(f, "array"),
            Self::DataTable => write!(f, "dataTable"),
            Self::Shared => write!(f, "shared"),
        }
    }
}

impl std::str::FromStr for SmlSTCellFormulaType {
    type Err = String;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        match s {
            "normal" => Ok(Self::Normal),
            "array" => Ok(Self::Array),
            "dataTable" => Ok(Self::DataTable),
            "shared" => Ok(Self::Shared),
            _ => Err(format!("unknown SmlSTCellFormulaType value: {}", s)),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum SmlSTPane {
    #[serde(rename = "bottomRight")]
    BottomRight,
    #[serde(rename = "topRight")]
    TopRight,
    #[serde(rename = "bottomLeft")]
    BottomLeft,
    #[serde(rename = "topLeft")]
    TopLeft,
}

impl std::fmt::Display for SmlSTPane {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::BottomRight => write!(f, "bottomRight"),
            Self::TopRight => write!(f, "topRight"),
            Self::BottomLeft => write!(f, "bottomLeft"),
            Self::TopLeft => write!(f, "topLeft"),
        }
    }
}

impl std::str::FromStr for SmlSTPane {
    type Err = String;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        match s {
            "bottomRight" => Ok(Self::BottomRight),
            "topRight" => Ok(Self::TopRight),
            "bottomLeft" => Ok(Self::BottomLeft),
            "topLeft" => Ok(Self::TopLeft),
            _ => Err(format!("unknown SmlSTPane value: {}", s)),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum SmlSTSheetViewType {
    #[serde(rename = "normal")]
    Normal,
    #[serde(rename = "pageBreakPreview")]
    PageBreakPreview,
    #[serde(rename = "pageLayout")]
    PageLayout,
}

impl std::fmt::Display for SmlSTSheetViewType {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::Normal => write!(f, "normal"),
            Self::PageBreakPreview => write!(f, "pageBreakPreview"),
            Self::PageLayout => write!(f, "pageLayout"),
        }
    }
}

impl std::str::FromStr for SmlSTSheetViewType {
    type Err = String;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        match s {
            "normal" => Ok(Self::Normal),
            "pageBreakPreview" => Ok(Self::PageBreakPreview),
            "pageLayout" => Ok(Self::PageLayout),
            _ => Err(format!("unknown SmlSTSheetViewType value: {}", s)),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum SmlSTDataConsolidateFunction {
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

impl std::fmt::Display for SmlSTDataConsolidateFunction {
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

impl std::str::FromStr for SmlSTDataConsolidateFunction {
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
            _ => Err(format!("unknown SmlSTDataConsolidateFunction value: {}", s)),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum SmlSTDataValidationType {
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

impl std::fmt::Display for SmlSTDataValidationType {
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

impl std::str::FromStr for SmlSTDataValidationType {
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
            _ => Err(format!("unknown SmlSTDataValidationType value: {}", s)),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum SmlSTDataValidationOperator {
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

impl std::fmt::Display for SmlSTDataValidationOperator {
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

impl std::str::FromStr for SmlSTDataValidationOperator {
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
            _ => Err(format!("unknown SmlSTDataValidationOperator value: {}", s)),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum SmlSTDataValidationErrorStyle {
    #[serde(rename = "stop")]
    Stop,
    #[serde(rename = "warning")]
    Warning,
    #[serde(rename = "information")]
    Information,
}

impl std::fmt::Display for SmlSTDataValidationErrorStyle {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::Stop => write!(f, "stop"),
            Self::Warning => write!(f, "warning"),
            Self::Information => write!(f, "information"),
        }
    }
}

impl std::str::FromStr for SmlSTDataValidationErrorStyle {
    type Err = String;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        match s {
            "stop" => Ok(Self::Stop),
            "warning" => Ok(Self::Warning),
            "information" => Ok(Self::Information),
            _ => Err(format!(
                "unknown SmlSTDataValidationErrorStyle value: {}",
                s
            )),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum SmlSTDataValidationImeMode {
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

impl std::fmt::Display for SmlSTDataValidationImeMode {
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

impl std::str::FromStr for SmlSTDataValidationImeMode {
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
            _ => Err(format!("unknown SmlSTDataValidationImeMode value: {}", s)),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum SmlSTCfType {
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

impl std::fmt::Display for SmlSTCfType {
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

impl std::str::FromStr for SmlSTCfType {
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
            _ => Err(format!("unknown SmlSTCfType value: {}", s)),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum SmlSTTimePeriod {
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

impl std::fmt::Display for SmlSTTimePeriod {
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

impl std::str::FromStr for SmlSTTimePeriod {
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
            _ => Err(format!("unknown SmlSTTimePeriod value: {}", s)),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum SmlSTConditionalFormattingOperator {
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

impl std::fmt::Display for SmlSTConditionalFormattingOperator {
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

impl std::str::FromStr for SmlSTConditionalFormattingOperator {
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
            _ => Err(format!(
                "unknown SmlSTConditionalFormattingOperator value: {}",
                s
            )),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum SmlSTCfvoType {
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

impl std::fmt::Display for SmlSTCfvoType {
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

impl std::str::FromStr for SmlSTCfvoType {
    type Err = String;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        match s {
            "num" => Ok(Self::Num),
            "percent" => Ok(Self::Percent),
            "max" => Ok(Self::Max),
            "min" => Ok(Self::Min),
            "formula" => Ok(Self::Formula),
            "percentile" => Ok(Self::Percentile),
            _ => Err(format!("unknown SmlSTCfvoType value: {}", s)),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum SmlSTPageOrder {
    #[serde(rename = "downThenOver")]
    DownThenOver,
    #[serde(rename = "overThenDown")]
    OverThenDown,
}

impl std::fmt::Display for SmlSTPageOrder {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::DownThenOver => write!(f, "downThenOver"),
            Self::OverThenDown => write!(f, "overThenDown"),
        }
    }
}

impl std::str::FromStr for SmlSTPageOrder {
    type Err = String;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        match s {
            "downThenOver" => Ok(Self::DownThenOver),
            "overThenDown" => Ok(Self::OverThenDown),
            _ => Err(format!("unknown SmlSTPageOrder value: {}", s)),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum SmlSTOrientation {
    #[serde(rename = "default")]
    Default,
    #[serde(rename = "portrait")]
    Portrait,
    #[serde(rename = "landscape")]
    Landscape,
}

impl std::fmt::Display for SmlSTOrientation {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::Default => write!(f, "default"),
            Self::Portrait => write!(f, "portrait"),
            Self::Landscape => write!(f, "landscape"),
        }
    }
}

impl std::str::FromStr for SmlSTOrientation {
    type Err = String;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        match s {
            "default" => Ok(Self::Default),
            "portrait" => Ok(Self::Portrait),
            "landscape" => Ok(Self::Landscape),
            _ => Err(format!("unknown SmlSTOrientation value: {}", s)),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum SmlSTCellComments {
    #[serde(rename = "none")]
    None,
    #[serde(rename = "asDisplayed")]
    AsDisplayed,
    #[serde(rename = "atEnd")]
    AtEnd,
}

impl std::fmt::Display for SmlSTCellComments {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::None => write!(f, "none"),
            Self::AsDisplayed => write!(f, "asDisplayed"),
            Self::AtEnd => write!(f, "atEnd"),
        }
    }
}

impl std::str::FromStr for SmlSTCellComments {
    type Err = String;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        match s {
            "none" => Ok(Self::None),
            "asDisplayed" => Ok(Self::AsDisplayed),
            "atEnd" => Ok(Self::AtEnd),
            _ => Err(format!("unknown SmlSTCellComments value: {}", s)),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum SmlSTPrintError {
    #[serde(rename = "displayed")]
    Displayed,
    #[serde(rename = "blank")]
    Blank,
    #[serde(rename = "dash")]
    Dash,
    #[serde(rename = "NA")]
    NA,
}

impl std::fmt::Display for SmlSTPrintError {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::Displayed => write!(f, "displayed"),
            Self::Blank => write!(f, "blank"),
            Self::Dash => write!(f, "dash"),
            Self::NA => write!(f, "NA"),
        }
    }
}

impl std::str::FromStr for SmlSTPrintError {
    type Err = String;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        match s {
            "displayed" => Ok(Self::Displayed),
            "blank" => Ok(Self::Blank),
            "dash" => Ok(Self::Dash),
            "NA" => Ok(Self::NA),
            _ => Err(format!("unknown SmlSTPrintError value: {}", s)),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum SmlSTDvAspect {
    #[serde(rename = "DVASPECT_CONTENT")]
    DVASPECTCONTENT,
    #[serde(rename = "DVASPECT_ICON")]
    DVASPECTICON,
}

impl std::fmt::Display for SmlSTDvAspect {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::DVASPECTCONTENT => write!(f, "DVASPECT_CONTENT"),
            Self::DVASPECTICON => write!(f, "DVASPECT_ICON"),
        }
    }
}

impl std::str::FromStr for SmlSTDvAspect {
    type Err = String;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        match s {
            "DVASPECT_CONTENT" => Ok(Self::DVASPECTCONTENT),
            "DVASPECT_ICON" => Ok(Self::DVASPECTICON),
            _ => Err(format!("unknown SmlSTDvAspect value: {}", s)),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum SmlSTOleUpdate {
    #[serde(rename = "OLEUPDATE_ALWAYS")]
    OLEUPDATEALWAYS,
    #[serde(rename = "OLEUPDATE_ONCALL")]
    OLEUPDATEONCALL,
}

impl std::fmt::Display for SmlSTOleUpdate {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::OLEUPDATEALWAYS => write!(f, "OLEUPDATE_ALWAYS"),
            Self::OLEUPDATEONCALL => write!(f, "OLEUPDATE_ONCALL"),
        }
    }
}

impl std::str::FromStr for SmlSTOleUpdate {
    type Err = String;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        match s {
            "OLEUPDATE_ALWAYS" => Ok(Self::OLEUPDATEALWAYS),
            "OLEUPDATE_ONCALL" => Ok(Self::OLEUPDATEONCALL),
            _ => Err(format!("unknown SmlSTOleUpdate value: {}", s)),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum SmlSTWebSourceType {
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

impl std::fmt::Display for SmlSTWebSourceType {
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

impl std::str::FromStr for SmlSTWebSourceType {
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
            _ => Err(format!("unknown SmlSTWebSourceType value: {}", s)),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum SmlSTPaneState {
    #[serde(rename = "split")]
    Split,
    #[serde(rename = "frozen")]
    Frozen,
    #[serde(rename = "frozenSplit")]
    FrozenSplit,
}

impl std::fmt::Display for SmlSTPaneState {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::Split => write!(f, "split"),
            Self::Frozen => write!(f, "frozen"),
            Self::FrozenSplit => write!(f, "frozenSplit"),
        }
    }
}

impl std::str::FromStr for SmlSTPaneState {
    type Err = String;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        match s {
            "split" => Ok(Self::Split),
            "frozen" => Ok(Self::Frozen),
            "frozenSplit" => Ok(Self::FrozenSplit),
            _ => Err(format!("unknown SmlSTPaneState value: {}", s)),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum SmlSTMdxFunctionType {
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

impl std::fmt::Display for SmlSTMdxFunctionType {
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

impl std::str::FromStr for SmlSTMdxFunctionType {
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
            _ => Err(format!("unknown SmlSTMdxFunctionType value: {}", s)),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum SmlSTMdxSetOrder {
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

impl std::fmt::Display for SmlSTMdxSetOrder {
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

impl std::str::FromStr for SmlSTMdxSetOrder {
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
            _ => Err(format!("unknown SmlSTMdxSetOrder value: {}", s)),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum SmlSTMdxKPIProperty {
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

impl std::fmt::Display for SmlSTMdxKPIProperty {
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

impl std::str::FromStr for SmlSTMdxKPIProperty {
    type Err = String;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        match s {
            "v" => Ok(Self::V),
            "g" => Ok(Self::G),
            "s" => Ok(Self::S),
            "t" => Ok(Self::T),
            "w" => Ok(Self::W),
            "m" => Ok(Self::M),
            _ => Err(format!("unknown SmlSTMdxKPIProperty value: {}", s)),
        }
    }
}

pub type SmlSTTextRotation = String;

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum SmlSTBorderStyle {
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

impl std::fmt::Display for SmlSTBorderStyle {
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

impl std::str::FromStr for SmlSTBorderStyle {
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
            _ => Err(format!("unknown SmlSTBorderStyle value: {}", s)),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum SmlSTPatternType {
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

impl std::fmt::Display for SmlSTPatternType {
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

impl std::str::FromStr for SmlSTPatternType {
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
            _ => Err(format!("unknown SmlSTPatternType value: {}", s)),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum SmlSTGradientType {
    #[serde(rename = "linear")]
    Linear,
    #[serde(rename = "path")]
    Path,
}

impl std::fmt::Display for SmlSTGradientType {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::Linear => write!(f, "linear"),
            Self::Path => write!(f, "path"),
        }
    }
}

impl std::str::FromStr for SmlSTGradientType {
    type Err = String;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        match s {
            "linear" => Ok(Self::Linear),
            "path" => Ok(Self::Path),
            _ => Err(format!("unknown SmlSTGradientType value: {}", s)),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum SmlSTHorizontalAlignment {
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

impl std::fmt::Display for SmlSTHorizontalAlignment {
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

impl std::str::FromStr for SmlSTHorizontalAlignment {
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
            _ => Err(format!("unknown SmlSTHorizontalAlignment value: {}", s)),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum SmlSTVerticalAlignment {
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

impl std::fmt::Display for SmlSTVerticalAlignment {
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

impl std::str::FromStr for SmlSTVerticalAlignment {
    type Err = String;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        match s {
            "top" => Ok(Self::Top),
            "center" => Ok(Self::Center),
            "bottom" => Ok(Self::Bottom),
            "justify" => Ok(Self::Justify),
            "distributed" => Ok(Self::Distributed),
            _ => Err(format!("unknown SmlSTVerticalAlignment value: {}", s)),
        }
    }
}

pub type SmlSTNumFmtId = u32;

pub type SmlSTFontId = u32;

pub type SmlSTFillId = u32;

pub type SmlSTBorderId = u32;

pub type SmlSTCellStyleXfId = u32;

pub type SmlSTDxfId = u32;

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum SmlSTTableStyleType {
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

impl std::fmt::Display for SmlSTTableStyleType {
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

impl std::str::FromStr for SmlSTTableStyleType {
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
            _ => Err(format!("unknown SmlSTTableStyleType value: {}", s)),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum SmlSTFontScheme {
    #[serde(rename = "none")]
    None,
    #[serde(rename = "major")]
    Major,
    #[serde(rename = "minor")]
    Minor,
}

impl std::fmt::Display for SmlSTFontScheme {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::None => write!(f, "none"),
            Self::Major => write!(f, "major"),
            Self::Minor => write!(f, "minor"),
        }
    }
}

impl std::str::FromStr for SmlSTFontScheme {
    type Err = String;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        match s {
            "none" => Ok(Self::None),
            "major" => Ok(Self::Major),
            "minor" => Ok(Self::Minor),
            _ => Err(format!("unknown SmlSTFontScheme value: {}", s)),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum SmlSTUnderlineValues {
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

impl std::fmt::Display for SmlSTUnderlineValues {
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

impl std::str::FromStr for SmlSTUnderlineValues {
    type Err = String;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        match s {
            "single" => Ok(Self::Single),
            "double" => Ok(Self::Double),
            "singleAccounting" => Ok(Self::SingleAccounting),
            "doubleAccounting" => Ok(Self::DoubleAccounting),
            "none" => Ok(Self::None),
            _ => Err(format!("unknown SmlSTUnderlineValues value: {}", s)),
        }
    }
}

pub type SmlSTFontFamily = i64;

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum SmlSTDdeValueType {
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

impl std::fmt::Display for SmlSTDdeValueType {
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

impl std::str::FromStr for SmlSTDdeValueType {
    type Err = String;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        match s {
            "nil" => Ok(Self::Nil),
            "b" => Ok(Self::B),
            "n" => Ok(Self::N),
            "e" => Ok(Self::E),
            "str" => Ok(Self::Str),
            _ => Err(format!("unknown SmlSTDdeValueType value: {}", s)),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum SmlSTTableType {
    #[serde(rename = "worksheet")]
    Worksheet,
    #[serde(rename = "xml")]
    Xml,
    #[serde(rename = "queryTable")]
    QueryTable,
}

impl std::fmt::Display for SmlSTTableType {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::Worksheet => write!(f, "worksheet"),
            Self::Xml => write!(f, "xml"),
            Self::QueryTable => write!(f, "queryTable"),
        }
    }
}

impl std::str::FromStr for SmlSTTableType {
    type Err = String;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        match s {
            "worksheet" => Ok(Self::Worksheet),
            "xml" => Ok(Self::Xml),
            "queryTable" => Ok(Self::QueryTable),
            _ => Err(format!("unknown SmlSTTableType value: {}", s)),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum SmlSTTotalsRowFunction {
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

impl std::fmt::Display for SmlSTTotalsRowFunction {
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

impl std::str::FromStr for SmlSTTotalsRowFunction {
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
            _ => Err(format!("unknown SmlSTTotalsRowFunction value: {}", s)),
        }
    }
}

pub type SmlSTXmlDataType = String;

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum SmlSTVolDepType {
    #[serde(rename = "realTimeData")]
    RealTimeData,
    #[serde(rename = "olapFunctions")]
    OlapFunctions,
}

impl std::fmt::Display for SmlSTVolDepType {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::RealTimeData => write!(f, "realTimeData"),
            Self::OlapFunctions => write!(f, "olapFunctions"),
        }
    }
}

impl std::str::FromStr for SmlSTVolDepType {
    type Err = String;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        match s {
            "realTimeData" => Ok(Self::RealTimeData),
            "olapFunctions" => Ok(Self::OlapFunctions),
            _ => Err(format!("unknown SmlSTVolDepType value: {}", s)),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum SmlSTVolValueType {
    #[serde(rename = "b")]
    B,
    #[serde(rename = "n")]
    N,
    #[serde(rename = "e")]
    E,
    #[serde(rename = "s")]
    S,
}

impl std::fmt::Display for SmlSTVolValueType {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::B => write!(f, "b"),
            Self::N => write!(f, "n"),
            Self::E => write!(f, "e"),
            Self::S => write!(f, "s"),
        }
    }
}

impl std::str::FromStr for SmlSTVolValueType {
    type Err = String;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        match s {
            "b" => Ok(Self::B),
            "n" => Ok(Self::N),
            "e" => Ok(Self::E),
            "s" => Ok(Self::S),
            _ => Err(format!("unknown SmlSTVolValueType value: {}", s)),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum SmlSTVisibility {
    #[serde(rename = "visible")]
    Visible,
    #[serde(rename = "hidden")]
    Hidden,
    #[serde(rename = "veryHidden")]
    VeryHidden,
}

impl std::fmt::Display for SmlSTVisibility {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::Visible => write!(f, "visible"),
            Self::Hidden => write!(f, "hidden"),
            Self::VeryHidden => write!(f, "veryHidden"),
        }
    }
}

impl std::str::FromStr for SmlSTVisibility {
    type Err = String;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        match s {
            "visible" => Ok(Self::Visible),
            "hidden" => Ok(Self::Hidden),
            "veryHidden" => Ok(Self::VeryHidden),
            _ => Err(format!("unknown SmlSTVisibility value: {}", s)),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum SmlSTComments {
    #[serde(rename = "commNone")]
    CommNone,
    #[serde(rename = "commIndicator")]
    CommIndicator,
    #[serde(rename = "commIndAndComment")]
    CommIndAndComment,
}

impl std::fmt::Display for SmlSTComments {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::CommNone => write!(f, "commNone"),
            Self::CommIndicator => write!(f, "commIndicator"),
            Self::CommIndAndComment => write!(f, "commIndAndComment"),
        }
    }
}

impl std::str::FromStr for SmlSTComments {
    type Err = String;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        match s {
            "commNone" => Ok(Self::CommNone),
            "commIndicator" => Ok(Self::CommIndicator),
            "commIndAndComment" => Ok(Self::CommIndAndComment),
            _ => Err(format!("unknown SmlSTComments value: {}", s)),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum SmlSTObjects {
    #[serde(rename = "all")]
    All,
    #[serde(rename = "placeholders")]
    Placeholders,
    #[serde(rename = "none")]
    None,
}

impl std::fmt::Display for SmlSTObjects {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::All => write!(f, "all"),
            Self::Placeholders => write!(f, "placeholders"),
            Self::None => write!(f, "none"),
        }
    }
}

impl std::str::FromStr for SmlSTObjects {
    type Err = String;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        match s {
            "all" => Ok(Self::All),
            "placeholders" => Ok(Self::Placeholders),
            "none" => Ok(Self::None),
            _ => Err(format!("unknown SmlSTObjects value: {}", s)),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum SmlSTSheetState {
    #[serde(rename = "visible")]
    Visible,
    #[serde(rename = "hidden")]
    Hidden,
    #[serde(rename = "veryHidden")]
    VeryHidden,
}

impl std::fmt::Display for SmlSTSheetState {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::Visible => write!(f, "visible"),
            Self::Hidden => write!(f, "hidden"),
            Self::VeryHidden => write!(f, "veryHidden"),
        }
    }
}

impl std::str::FromStr for SmlSTSheetState {
    type Err = String;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        match s {
            "visible" => Ok(Self::Visible),
            "hidden" => Ok(Self::Hidden),
            "veryHidden" => Ok(Self::VeryHidden),
            _ => Err(format!("unknown SmlSTSheetState value: {}", s)),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum SmlSTUpdateLinks {
    #[serde(rename = "userSet")]
    UserSet,
    #[serde(rename = "never")]
    Never,
    #[serde(rename = "always")]
    Always,
}

impl std::fmt::Display for SmlSTUpdateLinks {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::UserSet => write!(f, "userSet"),
            Self::Never => write!(f, "never"),
            Self::Always => write!(f, "always"),
        }
    }
}

impl std::str::FromStr for SmlSTUpdateLinks {
    type Err = String;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        match s {
            "userSet" => Ok(Self::UserSet),
            "never" => Ok(Self::Never),
            "always" => Ok(Self::Always),
            _ => Err(format!("unknown SmlSTUpdateLinks value: {}", s)),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum SmlSTSmartTagShow {
    #[serde(rename = "all")]
    All,
    #[serde(rename = "none")]
    None,
    #[serde(rename = "noIndicator")]
    NoIndicator,
}

impl std::fmt::Display for SmlSTSmartTagShow {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::All => write!(f, "all"),
            Self::None => write!(f, "none"),
            Self::NoIndicator => write!(f, "noIndicator"),
        }
    }
}

impl std::str::FromStr for SmlSTSmartTagShow {
    type Err = String;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        match s {
            "all" => Ok(Self::All),
            "none" => Ok(Self::None),
            "noIndicator" => Ok(Self::NoIndicator),
            _ => Err(format!("unknown SmlSTSmartTagShow value: {}", s)),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum SmlSTCalcMode {
    #[serde(rename = "manual")]
    Manual,
    #[serde(rename = "auto")]
    Auto,
    #[serde(rename = "autoNoTable")]
    AutoNoTable,
}

impl std::fmt::Display for SmlSTCalcMode {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::Manual => write!(f, "manual"),
            Self::Auto => write!(f, "auto"),
            Self::AutoNoTable => write!(f, "autoNoTable"),
        }
    }
}

impl std::str::FromStr for SmlSTCalcMode {
    type Err = String;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        match s {
            "manual" => Ok(Self::Manual),
            "auto" => Ok(Self::Auto),
            "autoNoTable" => Ok(Self::AutoNoTable),
            _ => Err(format!("unknown SmlSTCalcMode value: {}", s)),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum SmlSTRefMode {
    #[serde(rename = "A1")]
    A1,
    #[serde(rename = "R1C1")]
    R1C1,
}

impl std::fmt::Display for SmlSTRefMode {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::A1 => write!(f, "A1"),
            Self::R1C1 => write!(f, "R1C1"),
        }
    }
}

impl std::str::FromStr for SmlSTRefMode {
    type Err = String;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        match s {
            "A1" => Ok(Self::A1),
            "R1C1" => Ok(Self::R1C1),
            _ => Err(format!("unknown SmlSTRefMode value: {}", s)),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum SmlSTTargetScreenSize {
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

impl std::fmt::Display for SmlSTTargetScreenSize {
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

impl std::str::FromStr for SmlSTTargetScreenSize {
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
            _ => Err(format!("unknown SmlSTTargetScreenSize value: {}", s)),
        }
    }
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct SmlCTAutoFilter {
    #[serde(rename = "@ref")]
    #[serde(default)]
    pub r#ref: Option<SmlSTRef>,
    #[serde(rename = "filterColumn")]
    #[serde(default)]
    pub filter_column: Vec<Box<SmlCTFilterColumn>>,
    #[serde(rename = "sortState")]
    #[serde(default)]
    pub sort_state: Option<Box<SmlCTSortState>>,
    #[serde(rename = "extLst")]
    #[serde(default)]
    pub ext_lst: Option<Box<SmlCTExtensionList>>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct SmlCTFilterColumn {
    #[serde(rename = "@colId")]
    pub col_id: u32,
    #[serde(rename = "@hiddenButton")]
    #[serde(default)]
    pub hidden_button: Option<bool>,
    #[serde(rename = "@showButton")]
    #[serde(default)]
    pub show_button: Option<bool>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct SmlCTFilters {
    #[serde(rename = "@blank")]
    #[serde(default)]
    pub blank: Option<bool>,
    #[serde(rename = "@calendarType")]
    #[serde(default)]
    pub calendar_type: Option<SSTCalendarType>,
    #[serde(rename = "filter")]
    #[serde(default)]
    pub filter: Vec<Box<SmlCTFilter>>,
    #[serde(rename = "dateGroupItem")]
    #[serde(default)]
    pub date_group_item: Vec<Box<SmlCTDateGroupItem>>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct SmlCTFilter {
    #[serde(rename = "@val")]
    #[serde(default)]
    pub val: Option<SSTXstring>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct SmlCTCustomFilters {
    #[serde(rename = "@and")]
    #[serde(default)]
    pub and: Option<bool>,
    #[serde(rename = "customFilter")]
    #[serde(default)]
    pub custom_filter: Vec<Box<SmlCTCustomFilter>>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct SmlCTCustomFilter {
    #[serde(rename = "@operator")]
    #[serde(default)]
    pub operator: Option<SmlSTFilterOperator>,
    #[serde(rename = "@val")]
    #[serde(default)]
    pub val: Option<SSTXstring>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct SmlCTTop10 {
    #[serde(rename = "@top")]
    #[serde(default)]
    pub top: Option<bool>,
    #[serde(rename = "@percent")]
    #[serde(default)]
    pub percent: Option<bool>,
    #[serde(rename = "@val")]
    pub val: f64,
    #[serde(rename = "@filterVal")]
    #[serde(default)]
    pub filter_val: Option<f64>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct SmlCTColorFilter {
    #[serde(rename = "@dxfId")]
    #[serde(default)]
    pub dxf_id: Option<SmlSTDxfId>,
    #[serde(rename = "@cellColor")]
    #[serde(default)]
    pub cell_color: Option<bool>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct SmlCTIconFilter {
    #[serde(rename = "@iconSet")]
    pub icon_set: SmlSTIconSetType,
    #[serde(rename = "@iconId")]
    #[serde(default)]
    pub icon_id: Option<u32>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct SmlCTDynamicFilter {
    #[serde(rename = "@type")]
    pub r#type: SmlSTDynamicFilterType,
    #[serde(rename = "@val")]
    #[serde(default)]
    pub val: Option<f64>,
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
pub struct SmlCTSortState {
    #[serde(rename = "@columnSort")]
    #[serde(default)]
    pub column_sort: Option<bool>,
    #[serde(rename = "@caseSensitive")]
    #[serde(default)]
    pub case_sensitive: Option<bool>,
    #[serde(rename = "@sortMethod")]
    #[serde(default)]
    pub sort_method: Option<SmlSTSortMethod>,
    #[serde(rename = "@ref")]
    pub r#ref: SmlSTRef,
    #[serde(rename = "sortCondition")]
    #[serde(default)]
    pub sort_condition: Vec<Box<SmlCTSortCondition>>,
    #[serde(rename = "extLst")]
    #[serde(default)]
    pub ext_lst: Option<Box<SmlCTExtensionList>>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct SmlCTSortCondition {
    #[serde(rename = "@descending")]
    #[serde(default)]
    pub descending: Option<bool>,
    #[serde(rename = "@sortBy")]
    #[serde(default)]
    pub sort_by: Option<SmlSTSortBy>,
    #[serde(rename = "@ref")]
    pub r#ref: SmlSTRef,
    #[serde(rename = "@customList")]
    #[serde(default)]
    pub custom_list: Option<SSTXstring>,
    #[serde(rename = "@dxfId")]
    #[serde(default)]
    pub dxf_id: Option<SmlSTDxfId>,
    #[serde(rename = "@iconSet")]
    #[serde(default)]
    pub icon_set: Option<SmlSTIconSetType>,
    #[serde(rename = "@iconId")]
    #[serde(default)]
    pub icon_id: Option<u32>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct SmlCTDateGroupItem {
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
    pub date_time_grouping: SmlSTDateTimeGrouping,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct SmlCTXStringElement {
    #[serde(rename = "@v")]
    pub v: SSTXstring,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct SmlCTExtension {
    #[serde(rename = "@uri")]
    #[serde(default)]
    pub uri: Option<String>,
}

pub type SmlCTExtensionAny = String;

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct SmlCTObjectAnchor {
    #[serde(rename = "@moveWithCells")]
    #[serde(default)]
    pub move_with_cells: Option<bool>,
    #[serde(rename = "@sizeWithCells")]
    #[serde(default)]
    pub size_with_cells: Option<bool>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct SmlEGExtensionList {
    #[serde(rename = "ext")]
    #[serde(default)]
    pub ext: Vec<Box<SmlCTExtension>>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct SmlCTExtensionList;

pub type SmlCalcChain = Box<SmlCTCalcChain>;

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct SmlCTCalcChain {
    #[serde(rename = "c")]
    #[serde(default)]
    pub c: Vec<Box<SmlCTCalcCell>>,
    #[serde(rename = "extLst")]
    #[serde(default)]
    pub ext_lst: Option<Box<SmlCTExtensionList>>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct SmlCTCalcCell {
    #[serde(rename = "@_any")]
    pub _any: SmlSTCellRef,
    #[serde(rename = "@i")]
    #[serde(default)]
    pub i: Option<i32>,
    #[serde(rename = "@s")]
    #[serde(default)]
    pub s: Option<bool>,
    #[serde(rename = "@l")]
    #[serde(default)]
    pub l: Option<bool>,
    #[serde(rename = "@t")]
    #[serde(default)]
    pub t: Option<bool>,
    #[serde(rename = "@a")]
    #[serde(default)]
    pub a: Option<bool>,
}

pub type SmlComments = Box<SmlCTComments>;

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct SmlCTComments {
    #[serde(rename = "authors")]
    pub authors: Box<SmlCTAuthors>,
    #[serde(rename = "commentList")]
    pub comment_list: Box<SmlCTCommentList>,
    #[serde(rename = "extLst")]
    #[serde(default)]
    pub ext_lst: Option<Box<SmlCTExtensionList>>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct SmlCTAuthors {
    #[serde(rename = "author")]
    #[serde(default)]
    pub author: Vec<SSTXstring>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct SmlCTCommentList {
    #[serde(rename = "comment")]
    #[serde(default)]
    pub comment: Vec<Box<SmlCTComment>>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct SmlCTComment {
    #[serde(rename = "@ref")]
    pub r#ref: SmlSTRef,
    #[serde(rename = "@authorId")]
    pub author_id: u32,
    #[serde(rename = "@guid")]
    #[serde(default)]
    pub guid: Option<SSTGuid>,
    #[serde(rename = "@shapeId")]
    #[serde(default)]
    pub shape_id: Option<u32>,
    #[serde(rename = "text")]
    pub text: Box<SmlCTRst>,
    #[serde(rename = "commentPr")]
    #[serde(default)]
    pub comment_pr: Option<Box<SmlCTCommentPr>>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct SmlCTCommentPr {
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
    pub alt_text: Option<SSTXstring>,
    #[serde(rename = "@textHAlign")]
    #[serde(default)]
    pub text_h_align: Option<SmlSTTextHAlign>,
    #[serde(rename = "@textVAlign")]
    #[serde(default)]
    pub text_v_align: Option<SmlSTTextVAlign>,
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
    pub anchor: Box<SmlCTObjectAnchor>,
}

pub type SmlMapInfo = Box<SmlCTMapInfo>;

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct SmlCTMapInfo {
    #[serde(rename = "@SelectionNamespaces")]
    pub selection_namespaces: String,
    #[serde(rename = "Schema")]
    #[serde(default)]
    pub schema: Vec<Box<SmlCTSchema>>,
    #[serde(rename = "Map")]
    #[serde(default)]
    pub map: Vec<Box<SmlCTMap>>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct SmlCTSchema;

pub type SmlCTSchemaAny = String;

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct SmlCTMap {
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
    pub data_binding: Option<Box<SmlCTDataBinding>>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct SmlCTDataBinding {
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

pub type SmlCTDataBindingAny = String;

pub type SmlConnections = Box<SmlCTConnections>;

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct SmlCTConnections {
    #[serde(rename = "connection")]
    #[serde(default)]
    pub connection: Vec<Box<SmlCTConnection>>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct SmlCTConnection {
    #[serde(rename = "@id")]
    pub id: u32,
    #[serde(rename = "@sourceFile")]
    #[serde(default)]
    pub source_file: Option<SSTXstring>,
    #[serde(rename = "@odcFile")]
    #[serde(default)]
    pub odc_file: Option<SSTXstring>,
    #[serde(rename = "@keepAlive")]
    #[serde(default)]
    pub keep_alive: Option<bool>,
    #[serde(rename = "@interval")]
    #[serde(default)]
    pub interval: Option<u32>,
    #[serde(rename = "@name")]
    #[serde(default)]
    pub name: Option<SSTXstring>,
    #[serde(rename = "@description")]
    #[serde(default)]
    pub description: Option<SSTXstring>,
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
    pub credentials: Option<SmlSTCredMethod>,
    #[serde(rename = "@singleSignOnId")]
    #[serde(default)]
    pub single_sign_on_id: Option<SSTXstring>,
    #[serde(rename = "dbPr")]
    #[serde(default)]
    pub db_pr: Option<Box<SmlCTDbPr>>,
    #[serde(rename = "olapPr")]
    #[serde(default)]
    pub olap_pr: Option<Box<SmlCTOlapPr>>,
    #[serde(rename = "webPr")]
    #[serde(default)]
    pub web_pr: Option<Box<SmlCTWebPr>>,
    #[serde(rename = "textPr")]
    #[serde(default)]
    pub text_pr: Option<Box<SmlCTTextPr>>,
    #[serde(rename = "parameters")]
    #[serde(default)]
    pub parameters: Option<Box<SmlCTParameters>>,
    #[serde(rename = "extLst")]
    #[serde(default)]
    pub ext_lst: Option<Box<SmlCTExtensionList>>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct SmlCTDbPr {
    #[serde(rename = "@connection")]
    pub connection: SSTXstring,
    #[serde(rename = "@command")]
    #[serde(default)]
    pub command: Option<SSTXstring>,
    #[serde(rename = "@serverCommand")]
    #[serde(default)]
    pub server_command: Option<SSTXstring>,
    #[serde(rename = "@commandType")]
    #[serde(default)]
    pub command_type: Option<u32>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct SmlCTOlapPr {
    #[serde(rename = "@local")]
    #[serde(default)]
    pub local: Option<bool>,
    #[serde(rename = "@localConnection")]
    #[serde(default)]
    pub local_connection: Option<SSTXstring>,
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
pub struct SmlCTWebPr {
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
    pub url: Option<SSTXstring>,
    #[serde(rename = "@post")]
    #[serde(default)]
    pub post: Option<SSTXstring>,
    #[serde(rename = "@htmlTables")]
    #[serde(default)]
    pub html_tables: Option<bool>,
    #[serde(rename = "@htmlFormat")]
    #[serde(default)]
    pub html_format: Option<SmlSTHtmlFmt>,
    #[serde(rename = "@editPage")]
    #[serde(default)]
    pub edit_page: Option<SSTXstring>,
    #[serde(rename = "tables")]
    #[serde(default)]
    pub tables: Option<Box<SmlCTTables>>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct SmlCTParameters {
    #[serde(rename = "@count")]
    #[serde(default)]
    pub count: Option<u32>,
    #[serde(rename = "parameter")]
    #[serde(default)]
    pub parameter: Vec<Box<SmlCTParameter>>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct SmlCTParameter {
    #[serde(rename = "@name")]
    #[serde(default)]
    pub name: Option<SSTXstring>,
    #[serde(rename = "@sqlType")]
    #[serde(default)]
    pub sql_type: Option<i32>,
    #[serde(rename = "@parameterType")]
    #[serde(default)]
    pub parameter_type: Option<SmlSTParameterType>,
    #[serde(rename = "@refreshOnChange")]
    #[serde(default)]
    pub refresh_on_change: Option<bool>,
    #[serde(rename = "@prompt")]
    #[serde(default)]
    pub prompt: Option<SSTXstring>,
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
    pub string: Option<SSTXstring>,
    #[serde(rename = "@cell")]
    #[serde(default)]
    pub cell: Option<SSTXstring>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct SmlCTTables {
    #[serde(rename = "@count")]
    #[serde(default)]
    pub count: Option<u32>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct SmlCTTableMissing;

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct SmlCTTextPr {
    #[serde(rename = "@prompt")]
    #[serde(default)]
    pub prompt: Option<bool>,
    #[serde(rename = "@fileType")]
    #[serde(default)]
    pub file_type: Option<SmlSTFileType>,
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
    pub source_file: Option<SSTXstring>,
    #[serde(rename = "@delimited")]
    #[serde(default)]
    pub delimited: Option<bool>,
    #[serde(rename = "@decimal")]
    #[serde(default)]
    pub decimal: Option<SSTXstring>,
    #[serde(rename = "@thousands")]
    #[serde(default)]
    pub thousands: Option<SSTXstring>,
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
    pub qualifier: Option<SmlSTQualifier>,
    #[serde(rename = "@delimiter")]
    #[serde(default)]
    pub delimiter: Option<SSTXstring>,
    #[serde(rename = "textFields")]
    #[serde(default)]
    pub text_fields: Option<Box<SmlCTTextFields>>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct SmlCTTextFields {
    #[serde(rename = "@count")]
    #[serde(default)]
    pub count: Option<u32>,
    #[serde(rename = "textField")]
    #[serde(default)]
    pub text_field: Vec<Box<SmlCTTextField>>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct SmlCTTextField {
    #[serde(rename = "@type")]
    #[serde(default)]
    pub r#type: Option<SmlSTExternalConnectionType>,
    #[serde(rename = "@position")]
    #[serde(default)]
    pub position: Option<u32>,
}

pub type SmlPivotCacheDefinition = Box<SmlCTPivotCacheDefinition>;

pub type SmlPivotCacheRecords = Box<SmlCTPivotCacheRecords>;

pub type SmlPivotTableDefinition = Box<SmlCTPivotTableDefinition>;

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct SmlCTPivotCacheDefinition {
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
    pub refreshed_by: Option<SSTXstring>,
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
    pub cache_source: Box<SmlCTCacheSource>,
    #[serde(rename = "cacheFields")]
    pub cache_fields: Box<SmlCTCacheFields>,
    #[serde(rename = "cacheHierarchies")]
    #[serde(default)]
    pub cache_hierarchies: Option<Box<SmlCTCacheHierarchies>>,
    #[serde(rename = "kpis")]
    #[serde(default)]
    pub kpis: Option<Box<SmlCTPCDKPIs>>,
    #[serde(rename = "calculatedItems")]
    #[serde(default)]
    pub calculated_items: Option<Box<SmlCTCalculatedItems>>,
    #[serde(rename = "calculatedMembers")]
    #[serde(default)]
    pub calculated_members: Option<Box<SmlCTCalculatedMembers>>,
    #[serde(rename = "dimensions")]
    #[serde(default)]
    pub dimensions: Option<Box<SmlCTDimensions>>,
    #[serde(rename = "measureGroups")]
    #[serde(default)]
    pub measure_groups: Option<Box<SmlCTMeasureGroups>>,
    #[serde(rename = "maps")]
    #[serde(default)]
    pub maps: Option<Box<SmlCTMeasureDimensionMaps>>,
    #[serde(rename = "extLst")]
    #[serde(default)]
    pub ext_lst: Option<Box<SmlCTExtensionList>>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct SmlCTCacheFields {
    #[serde(rename = "@count")]
    #[serde(default)]
    pub count: Option<u32>,
    #[serde(rename = "cacheField")]
    #[serde(default)]
    pub cache_field: Vec<Box<SmlCTCacheField>>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct SmlCTCacheField {
    #[serde(rename = "@name")]
    pub name: SSTXstring,
    #[serde(rename = "@caption")]
    #[serde(default)]
    pub caption: Option<SSTXstring>,
    #[serde(rename = "@propertyName")]
    #[serde(default)]
    pub property_name: Option<SSTXstring>,
    #[serde(rename = "@serverField")]
    #[serde(default)]
    pub server_field: Option<bool>,
    #[serde(rename = "@uniqueList")]
    #[serde(default)]
    pub unique_list: Option<bool>,
    #[serde(rename = "@numFmtId")]
    #[serde(default)]
    pub num_fmt_id: Option<SmlSTNumFmtId>,
    #[serde(rename = "@formula")]
    #[serde(default)]
    pub formula: Option<SSTXstring>,
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
    pub shared_items: Option<Box<SmlCTSharedItems>>,
    #[serde(rename = "fieldGroup")]
    #[serde(default)]
    pub field_group: Option<Box<SmlCTFieldGroup>>,
    #[serde(rename = "mpMap")]
    #[serde(default)]
    pub mp_map: Vec<Box<SmlCTX>>,
    #[serde(rename = "extLst")]
    #[serde(default)]
    pub ext_lst: Option<Box<SmlCTExtensionList>>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct SmlCTCacheSource {
    #[serde(rename = "@type")]
    pub r#type: SmlSTSourceType,
    #[serde(rename = "@connectionId")]
    #[serde(default)]
    pub connection_id: Option<u32>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct SmlCTWorksheetSource {
    #[serde(rename = "@ref")]
    #[serde(default)]
    pub r#ref: Option<SmlSTRef>,
    #[serde(rename = "@name")]
    #[serde(default)]
    pub name: Option<SSTXstring>,
    #[serde(rename = "@sheet")]
    #[serde(default)]
    pub sheet: Option<SSTXstring>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct SmlCTConsolidation {
    #[serde(rename = "@autoPage")]
    #[serde(default)]
    pub auto_page: Option<bool>,
    #[serde(rename = "pages")]
    #[serde(default)]
    pub pages: Option<Box<SmlCTPages>>,
    #[serde(rename = "rangeSets")]
    pub range_sets: Box<SmlCTRangeSets>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct SmlCTPages {
    #[serde(rename = "@count")]
    #[serde(default)]
    pub count: Option<u32>,
    #[serde(rename = "page")]
    #[serde(default)]
    pub page: Vec<Box<SmlCTPCDSCPage>>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct SmlCTPCDSCPage {
    #[serde(rename = "@count")]
    #[serde(default)]
    pub count: Option<u32>,
    #[serde(rename = "pageItem")]
    #[serde(default)]
    pub page_item: Vec<Box<SmlCTPageItem>>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct SmlCTPageItem {
    #[serde(rename = "@name")]
    pub name: SSTXstring,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct SmlCTRangeSets {
    #[serde(rename = "@count")]
    #[serde(default)]
    pub count: Option<u32>,
    #[serde(rename = "rangeSet")]
    #[serde(default)]
    pub range_set: Vec<Box<SmlCTRangeSet>>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct SmlCTRangeSet {
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
    pub r#ref: Option<SmlSTRef>,
    #[serde(rename = "@name")]
    #[serde(default)]
    pub name: Option<SSTXstring>,
    #[serde(rename = "@sheet")]
    #[serde(default)]
    pub sheet: Option<SSTXstring>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct SmlCTSharedItems {
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
pub struct SmlCTMissing {
    #[serde(rename = "@u")]
    #[serde(default)]
    pub u: Option<bool>,
    #[serde(rename = "@f")]
    #[serde(default)]
    pub f: Option<bool>,
    #[serde(rename = "@c")]
    #[serde(default)]
    pub c: Option<SSTXstring>,
    #[serde(rename = "@cp")]
    #[serde(default)]
    pub cp: Option<u32>,
    #[serde(rename = "@in")]
    #[serde(default)]
    pub r#in: Option<u32>,
    #[serde(rename = "@bc")]
    #[serde(default)]
    pub bc: Option<SmlSTUnsignedIntHex>,
    #[serde(rename = "@fc")]
    #[serde(default)]
    pub fc: Option<SmlSTUnsignedIntHex>,
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
    pub tpls: Vec<Box<SmlCTTuples>>,
    #[serde(rename = "x")]
    #[serde(default)]
    pub x: Vec<Box<SmlCTX>>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct SmlCTNumber {
    #[serde(rename = "@v")]
    pub v: f64,
    #[serde(rename = "@u")]
    #[serde(default)]
    pub u: Option<bool>,
    #[serde(rename = "@f")]
    #[serde(default)]
    pub f: Option<bool>,
    #[serde(rename = "@c")]
    #[serde(default)]
    pub c: Option<SSTXstring>,
    #[serde(rename = "@cp")]
    #[serde(default)]
    pub cp: Option<u32>,
    #[serde(rename = "@in")]
    #[serde(default)]
    pub r#in: Option<u32>,
    #[serde(rename = "@bc")]
    #[serde(default)]
    pub bc: Option<SmlSTUnsignedIntHex>,
    #[serde(rename = "@fc")]
    #[serde(default)]
    pub fc: Option<SmlSTUnsignedIntHex>,
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
    pub tpls: Vec<Box<SmlCTTuples>>,
    #[serde(rename = "x")]
    #[serde(default)]
    pub x: Vec<Box<SmlCTX>>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct SmlCTBoolean {
    #[serde(rename = "@v")]
    pub v: bool,
    #[serde(rename = "@u")]
    #[serde(default)]
    pub u: Option<bool>,
    #[serde(rename = "@f")]
    #[serde(default)]
    pub f: Option<bool>,
    #[serde(rename = "@c")]
    #[serde(default)]
    pub c: Option<SSTXstring>,
    #[serde(rename = "@cp")]
    #[serde(default)]
    pub cp: Option<u32>,
    #[serde(rename = "x")]
    #[serde(default)]
    pub x: Vec<Box<SmlCTX>>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct SmlCTError {
    #[serde(rename = "@v")]
    pub v: SSTXstring,
    #[serde(rename = "@u")]
    #[serde(default)]
    pub u: Option<bool>,
    #[serde(rename = "@f")]
    #[serde(default)]
    pub f: Option<bool>,
    #[serde(rename = "@c")]
    #[serde(default)]
    pub c: Option<SSTXstring>,
    #[serde(rename = "@cp")]
    #[serde(default)]
    pub cp: Option<u32>,
    #[serde(rename = "@in")]
    #[serde(default)]
    pub r#in: Option<u32>,
    #[serde(rename = "@bc")]
    #[serde(default)]
    pub bc: Option<SmlSTUnsignedIntHex>,
    #[serde(rename = "@fc")]
    #[serde(default)]
    pub fc: Option<SmlSTUnsignedIntHex>,
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
    pub tpls: Option<Box<SmlCTTuples>>,
    #[serde(rename = "x")]
    #[serde(default)]
    pub x: Vec<Box<SmlCTX>>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct SmlCTString {
    #[serde(rename = "@v")]
    pub v: SSTXstring,
    #[serde(rename = "@u")]
    #[serde(default)]
    pub u: Option<bool>,
    #[serde(rename = "@f")]
    #[serde(default)]
    pub f: Option<bool>,
    #[serde(rename = "@c")]
    #[serde(default)]
    pub c: Option<SSTXstring>,
    #[serde(rename = "@cp")]
    #[serde(default)]
    pub cp: Option<u32>,
    #[serde(rename = "@in")]
    #[serde(default)]
    pub r#in: Option<u32>,
    #[serde(rename = "@bc")]
    #[serde(default)]
    pub bc: Option<SmlSTUnsignedIntHex>,
    #[serde(rename = "@fc")]
    #[serde(default)]
    pub fc: Option<SmlSTUnsignedIntHex>,
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
    pub tpls: Vec<Box<SmlCTTuples>>,
    #[serde(rename = "x")]
    #[serde(default)]
    pub x: Vec<Box<SmlCTX>>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct SmlCTDateTime {
    #[serde(rename = "@v")]
    pub v: String,
    #[serde(rename = "@u")]
    #[serde(default)]
    pub u: Option<bool>,
    #[serde(rename = "@f")]
    #[serde(default)]
    pub f: Option<bool>,
    #[serde(rename = "@c")]
    #[serde(default)]
    pub c: Option<SSTXstring>,
    #[serde(rename = "@cp")]
    #[serde(default)]
    pub cp: Option<u32>,
    #[serde(rename = "x")]
    #[serde(default)]
    pub x: Vec<Box<SmlCTX>>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct SmlCTFieldGroup {
    #[serde(rename = "@par")]
    #[serde(default)]
    pub par: Option<u32>,
    #[serde(rename = "@base")]
    #[serde(default)]
    pub base: Option<u32>,
    #[serde(rename = "rangePr")]
    #[serde(default)]
    pub range_pr: Option<Box<SmlCTRangePr>>,
    #[serde(rename = "discretePr")]
    #[serde(default)]
    pub discrete_pr: Option<Box<SmlCTDiscretePr>>,
    #[serde(rename = "groupItems")]
    #[serde(default)]
    pub group_items: Option<Box<SmlCTGroupItems>>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct SmlCTRangePr {
    #[serde(rename = "@autoStart")]
    #[serde(default)]
    pub auto_start: Option<bool>,
    #[serde(rename = "@autoEnd")]
    #[serde(default)]
    pub auto_end: Option<bool>,
    #[serde(rename = "@groupBy")]
    #[serde(default)]
    pub group_by: Option<SmlSTGroupBy>,
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
pub struct SmlCTDiscretePr {
    #[serde(rename = "@count")]
    #[serde(default)]
    pub count: Option<u32>,
    #[serde(rename = "x")]
    #[serde(default)]
    pub x: Vec<Box<SmlCTIndex>>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct SmlCTGroupItems {
    #[serde(rename = "@count")]
    #[serde(default)]
    pub count: Option<u32>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct SmlCTPivotCacheRecords {
    #[serde(rename = "@count")]
    #[serde(default)]
    pub count: Option<u32>,
    #[serde(rename = "r")]
    #[serde(default)]
    pub r: Vec<Box<SmlCTRecord>>,
    #[serde(rename = "extLst")]
    #[serde(default)]
    pub ext_lst: Option<Box<SmlCTExtensionList>>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct SmlCTRecord;

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct SmlCTPCDKPIs {
    #[serde(rename = "@count")]
    #[serde(default)]
    pub count: Option<u32>,
    #[serde(rename = "kpi")]
    #[serde(default)]
    pub kpi: Vec<Box<SmlCTPCDKPI>>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct SmlCTPCDKPI {
    #[serde(rename = "@uniqueName")]
    pub unique_name: SSTXstring,
    #[serde(rename = "@caption")]
    #[serde(default)]
    pub caption: Option<SSTXstring>,
    #[serde(rename = "@displayFolder")]
    #[serde(default)]
    pub display_folder: Option<SSTXstring>,
    #[serde(rename = "@measureGroup")]
    #[serde(default)]
    pub measure_group: Option<SSTXstring>,
    #[serde(rename = "@parent")]
    #[serde(default)]
    pub parent: Option<SSTXstring>,
    #[serde(rename = "@value")]
    pub value: SSTXstring,
    #[serde(rename = "@goal")]
    #[serde(default)]
    pub goal: Option<SSTXstring>,
    #[serde(rename = "@status")]
    #[serde(default)]
    pub status: Option<SSTXstring>,
    #[serde(rename = "@trend")]
    #[serde(default)]
    pub trend: Option<SSTXstring>,
    #[serde(rename = "@weight")]
    #[serde(default)]
    pub weight: Option<SSTXstring>,
    #[serde(rename = "@time")]
    #[serde(default)]
    pub time: Option<SSTXstring>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct SmlCTCacheHierarchies {
    #[serde(rename = "@count")]
    #[serde(default)]
    pub count: Option<u32>,
    #[serde(rename = "cacheHierarchy")]
    #[serde(default)]
    pub cache_hierarchy: Vec<Box<SmlCTCacheHierarchy>>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct SmlCTCacheHierarchy {
    #[serde(rename = "@uniqueName")]
    pub unique_name: SSTXstring,
    #[serde(rename = "@caption")]
    #[serde(default)]
    pub caption: Option<SSTXstring>,
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
    pub default_member_unique_name: Option<SSTXstring>,
    #[serde(rename = "@allUniqueName")]
    #[serde(default)]
    pub all_unique_name: Option<SSTXstring>,
    #[serde(rename = "@allCaption")]
    #[serde(default)]
    pub all_caption: Option<SSTXstring>,
    #[serde(rename = "@dimensionUniqueName")]
    #[serde(default)]
    pub dimension_unique_name: Option<SSTXstring>,
    #[serde(rename = "@displayFolder")]
    #[serde(default)]
    pub display_folder: Option<SSTXstring>,
    #[serde(rename = "@measureGroup")]
    #[serde(default)]
    pub measure_group: Option<SSTXstring>,
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
    pub fields_usage: Option<Box<SmlCTFieldsUsage>>,
    #[serde(rename = "groupLevels")]
    #[serde(default)]
    pub group_levels: Option<Box<SmlCTGroupLevels>>,
    #[serde(rename = "extLst")]
    #[serde(default)]
    pub ext_lst: Option<Box<SmlCTExtensionList>>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct SmlCTFieldsUsage {
    #[serde(rename = "@count")]
    #[serde(default)]
    pub count: Option<u32>,
    #[serde(rename = "fieldUsage")]
    #[serde(default)]
    pub field_usage: Vec<Box<SmlCTFieldUsage>>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct SmlCTFieldUsage {
    #[serde(rename = "@x")]
    pub x: i32,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct SmlCTGroupLevels {
    #[serde(rename = "@count")]
    #[serde(default)]
    pub count: Option<u32>,
    #[serde(rename = "groupLevel")]
    #[serde(default)]
    pub group_level: Vec<Box<SmlCTGroupLevel>>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct SmlCTGroupLevel {
    #[serde(rename = "@uniqueName")]
    pub unique_name: SSTXstring,
    #[serde(rename = "@caption")]
    pub caption: SSTXstring,
    #[serde(rename = "@user")]
    #[serde(default)]
    pub user: Option<bool>,
    #[serde(rename = "@customRollUp")]
    #[serde(default)]
    pub custom_roll_up: Option<bool>,
    #[serde(rename = "groups")]
    #[serde(default)]
    pub groups: Option<Box<SmlCTGroups>>,
    #[serde(rename = "extLst")]
    #[serde(default)]
    pub ext_lst: Option<Box<SmlCTExtensionList>>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct SmlCTGroups {
    #[serde(rename = "@count")]
    #[serde(default)]
    pub count: Option<u32>,
    #[serde(rename = "group")]
    #[serde(default)]
    pub group: Vec<Box<SmlCTLevelGroup>>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct SmlCTLevelGroup {
    #[serde(rename = "@name")]
    pub name: SSTXstring,
    #[serde(rename = "@uniqueName")]
    pub unique_name: SSTXstring,
    #[serde(rename = "@caption")]
    pub caption: SSTXstring,
    #[serde(rename = "@uniqueParent")]
    #[serde(default)]
    pub unique_parent: Option<SSTXstring>,
    #[serde(rename = "@id")]
    #[serde(default)]
    pub id: Option<i32>,
    #[serde(rename = "groupMembers")]
    pub group_members: Box<SmlCTGroupMembers>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct SmlCTGroupMembers {
    #[serde(rename = "@count")]
    #[serde(default)]
    pub count: Option<u32>,
    #[serde(rename = "groupMember")]
    #[serde(default)]
    pub group_member: Vec<Box<SmlCTGroupMember>>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct SmlCTGroupMember {
    #[serde(rename = "@uniqueName")]
    pub unique_name: SSTXstring,
    #[serde(rename = "@group")]
    #[serde(default)]
    pub group: Option<bool>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct SmlCTTupleCache {
    #[serde(rename = "entries")]
    #[serde(default)]
    pub entries: Option<Box<SmlCTPCDSDTCEntries>>,
    #[serde(rename = "sets")]
    #[serde(default)]
    pub sets: Option<Box<SmlCTSets>>,
    #[serde(rename = "queryCache")]
    #[serde(default)]
    pub query_cache: Option<Box<SmlCTQueryCache>>,
    #[serde(rename = "serverFormats")]
    #[serde(default)]
    pub server_formats: Option<Box<SmlCTServerFormats>>,
    #[serde(rename = "extLst")]
    #[serde(default)]
    pub ext_lst: Option<Box<SmlCTExtensionList>>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct SmlCTServerFormat {
    #[serde(rename = "@culture")]
    #[serde(default)]
    pub culture: Option<SSTXstring>,
    #[serde(rename = "@format")]
    #[serde(default)]
    pub format: Option<SSTXstring>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct SmlCTServerFormats {
    #[serde(rename = "@count")]
    #[serde(default)]
    pub count: Option<u32>,
    #[serde(rename = "serverFormat")]
    #[serde(default)]
    pub server_format: Vec<Box<SmlCTServerFormat>>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct SmlCTPCDSDTCEntries {
    #[serde(rename = "@count")]
    #[serde(default)]
    pub count: Option<u32>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct SmlCTTuples {
    #[serde(rename = "@c")]
    #[serde(default)]
    pub c: Option<u32>,
    #[serde(rename = "tpl")]
    #[serde(default)]
    pub tpl: Vec<Box<SmlCTTuple>>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct SmlCTTuple {
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
pub struct SmlCTSets {
    #[serde(rename = "@count")]
    #[serde(default)]
    pub count: Option<u32>,
    #[serde(rename = "set")]
    #[serde(default)]
    pub set: Vec<Box<SmlCTSet>>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct SmlCTSet {
    #[serde(rename = "@count")]
    #[serde(default)]
    pub count: Option<u32>,
    #[serde(rename = "@maxRank")]
    pub max_rank: i32,
    #[serde(rename = "@setDefinition")]
    pub set_definition: SSTXstring,
    #[serde(rename = "@sortType")]
    #[serde(default)]
    pub sort_type: Option<SmlSTSortType>,
    #[serde(rename = "@queryFailed")]
    #[serde(default)]
    pub query_failed: Option<bool>,
    #[serde(rename = "tpls")]
    #[serde(default)]
    pub tpls: Vec<Box<SmlCTTuples>>,
    #[serde(rename = "sortByTuple")]
    #[serde(default)]
    pub sort_by_tuple: Option<Box<SmlCTTuples>>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct SmlCTQueryCache {
    #[serde(rename = "@count")]
    #[serde(default)]
    pub count: Option<u32>,
    #[serde(rename = "query")]
    #[serde(default)]
    pub query: Vec<Box<SmlCTQuery>>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct SmlCTQuery {
    #[serde(rename = "@mdx")]
    pub mdx: SSTXstring,
    #[serde(rename = "tpls")]
    #[serde(default)]
    pub tpls: Option<Box<SmlCTTuples>>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct SmlCTCalculatedItems {
    #[serde(rename = "@count")]
    #[serde(default)]
    pub count: Option<u32>,
    #[serde(rename = "calculatedItem")]
    #[serde(default)]
    pub calculated_item: Vec<Box<SmlCTCalculatedItem>>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct SmlCTCalculatedItem {
    #[serde(rename = "@field")]
    #[serde(default)]
    pub field: Option<u32>,
    #[serde(rename = "@formula")]
    #[serde(default)]
    pub formula: Option<SSTXstring>,
    #[serde(rename = "pivotArea")]
    pub pivot_area: Box<SmlCTPivotArea>,
    #[serde(rename = "extLst")]
    #[serde(default)]
    pub ext_lst: Option<Box<SmlCTExtensionList>>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct SmlCTCalculatedMembers {
    #[serde(rename = "@count")]
    #[serde(default)]
    pub count: Option<u32>,
    #[serde(rename = "calculatedMember")]
    #[serde(default)]
    pub calculated_member: Vec<Box<SmlCTCalculatedMember>>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct SmlCTCalculatedMember {
    #[serde(rename = "@name")]
    pub name: SSTXstring,
    #[serde(rename = "@mdx")]
    pub mdx: SSTXstring,
    #[serde(rename = "@memberName")]
    #[serde(default)]
    pub member_name: Option<SSTXstring>,
    #[serde(rename = "@hierarchy")]
    #[serde(default)]
    pub hierarchy: Option<SSTXstring>,
    #[serde(rename = "@parent")]
    #[serde(default)]
    pub parent: Option<SSTXstring>,
    #[serde(rename = "@solveOrder")]
    #[serde(default)]
    pub solve_order: Option<i32>,
    #[serde(rename = "@set")]
    #[serde(default)]
    pub set: Option<bool>,
    #[serde(rename = "extLst")]
    #[serde(default)]
    pub ext_lst: Option<Box<SmlCTExtensionList>>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct SmlCTPivotTableDefinition {
    #[serde(rename = "@name")]
    pub name: SSTXstring,
    #[serde(rename = "@cacheId")]
    pub cache_id: u32,
    #[serde(rename = "@dataOnRows")]
    #[serde(default)]
    pub data_on_rows: Option<bool>,
    #[serde(rename = "@dataPosition")]
    #[serde(default)]
    pub data_position: Option<u32>,
    #[serde(rename = "@dataCaption")]
    pub data_caption: SSTXstring,
    #[serde(rename = "@grandTotalCaption")]
    #[serde(default)]
    pub grand_total_caption: Option<SSTXstring>,
    #[serde(rename = "@errorCaption")]
    #[serde(default)]
    pub error_caption: Option<SSTXstring>,
    #[serde(rename = "@showError")]
    #[serde(default)]
    pub show_error: Option<bool>,
    #[serde(rename = "@missingCaption")]
    #[serde(default)]
    pub missing_caption: Option<SSTXstring>,
    #[serde(rename = "@showMissing")]
    #[serde(default)]
    pub show_missing: Option<bool>,
    #[serde(rename = "@pageStyle")]
    #[serde(default)]
    pub page_style: Option<SSTXstring>,
    #[serde(rename = "@pivotTableStyle")]
    #[serde(default)]
    pub pivot_table_style: Option<SSTXstring>,
    #[serde(rename = "@vacatedStyle")]
    #[serde(default)]
    pub vacated_style: Option<SSTXstring>,
    #[serde(rename = "@tag")]
    #[serde(default)]
    pub tag: Option<SSTXstring>,
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
    pub row_header_caption: Option<SSTXstring>,
    #[serde(rename = "@colHeaderCaption")]
    #[serde(default)]
    pub col_header_caption: Option<SSTXstring>,
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
    pub location: Box<SmlCTLocation>,
    #[serde(rename = "pivotFields")]
    #[serde(default)]
    pub pivot_fields: Option<Box<SmlCTPivotFields>>,
    #[serde(rename = "rowFields")]
    #[serde(default)]
    pub row_fields: Option<Box<SmlCTRowFields>>,
    #[serde(rename = "rowItems")]
    #[serde(default)]
    pub row_items: Option<Box<SmlCTRowItems>>,
    #[serde(rename = "colFields")]
    #[serde(default)]
    pub col_fields: Option<Box<SmlCTColFields>>,
    #[serde(rename = "colItems")]
    #[serde(default)]
    pub col_items: Option<Box<SmlCTColItems>>,
    #[serde(rename = "pageFields")]
    #[serde(default)]
    pub page_fields: Option<Box<SmlCTPageFields>>,
    #[serde(rename = "dataFields")]
    #[serde(default)]
    pub data_fields: Option<Box<SmlCTDataFields>>,
    #[serde(rename = "formats")]
    #[serde(default)]
    pub formats: Option<Box<SmlCTFormats>>,
    #[serde(rename = "conditionalFormats")]
    #[serde(default)]
    pub conditional_formats: Option<Box<SmlCTConditionalFormats>>,
    #[serde(rename = "chartFormats")]
    #[serde(default)]
    pub chart_formats: Option<Box<SmlCTChartFormats>>,
    #[serde(rename = "pivotHierarchies")]
    #[serde(default)]
    pub pivot_hierarchies: Option<Box<SmlCTPivotHierarchies>>,
    #[serde(rename = "pivotTableStyleInfo")]
    #[serde(default)]
    pub pivot_table_style_info: Option<Box<SmlCTPivotTableStyle>>,
    #[serde(rename = "filters")]
    #[serde(default)]
    pub filters: Option<Box<SmlCTPivotFilters>>,
    #[serde(rename = "rowHierarchiesUsage")]
    #[serde(default)]
    pub row_hierarchies_usage: Option<Box<SmlCTRowHierarchiesUsage>>,
    #[serde(rename = "colHierarchiesUsage")]
    #[serde(default)]
    pub col_hierarchies_usage: Option<Box<SmlCTColHierarchiesUsage>>,
    #[serde(rename = "extLst")]
    #[serde(default)]
    pub ext_lst: Option<Box<SmlCTExtensionList>>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct SmlCTLocation {
    #[serde(rename = "@ref")]
    pub r#ref: SmlSTRef,
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
pub struct SmlCTPivotFields {
    #[serde(rename = "@count")]
    #[serde(default)]
    pub count: Option<u32>,
    #[serde(rename = "pivotField")]
    #[serde(default)]
    pub pivot_field: Vec<Box<SmlCTPivotField>>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct SmlCTPivotField {
    #[serde(rename = "@name")]
    #[serde(default)]
    pub name: Option<SSTXstring>,
    #[serde(rename = "@axis")]
    #[serde(default)]
    pub axis: Option<SmlSTAxis>,
    #[serde(rename = "@dataField")]
    #[serde(default)]
    pub data_field: Option<bool>,
    #[serde(rename = "@subtotalCaption")]
    #[serde(default)]
    pub subtotal_caption: Option<SSTXstring>,
    #[serde(rename = "@showDropDowns")]
    #[serde(default)]
    pub show_drop_downs: Option<bool>,
    #[serde(rename = "@hiddenLevel")]
    #[serde(default)]
    pub hidden_level: Option<bool>,
    #[serde(rename = "@uniqueMemberProperty")]
    #[serde(default)]
    pub unique_member_property: Option<SSTXstring>,
    #[serde(rename = "@compact")]
    #[serde(default)]
    pub compact: Option<bool>,
    #[serde(rename = "@allDrilled")]
    #[serde(default)]
    pub all_drilled: Option<bool>,
    #[serde(rename = "@numFmtId")]
    #[serde(default)]
    pub num_fmt_id: Option<SmlSTNumFmtId>,
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
    pub sort_type: Option<SmlSTFieldSortType>,
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
    pub items: Option<Box<SmlCTItems>>,
    #[serde(rename = "autoSortScope")]
    #[serde(default)]
    pub auto_sort_scope: Option<Box<SmlCTAutoSortScope>>,
    #[serde(rename = "extLst")]
    #[serde(default)]
    pub ext_lst: Option<Box<SmlCTExtensionList>>,
}

pub type SmlCTAutoSortScope = Box<SmlCTPivotArea>;

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct SmlCTItems {
    #[serde(rename = "@count")]
    #[serde(default)]
    pub count: Option<u32>,
    #[serde(rename = "item")]
    #[serde(default)]
    pub item: Vec<Box<SmlCTItem>>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct SmlCTItem {
    #[serde(rename = "@n")]
    #[serde(default)]
    pub n: Option<SSTXstring>,
    #[serde(rename = "@t")]
    #[serde(default)]
    pub t: Option<SmlSTItemType>,
    #[serde(rename = "@h")]
    #[serde(default)]
    pub h: Option<bool>,
    #[serde(rename = "@s")]
    #[serde(default)]
    pub s: Option<bool>,
    #[serde(rename = "@sd")]
    #[serde(default)]
    pub sd: Option<bool>,
    #[serde(rename = "@f")]
    #[serde(default)]
    pub f: Option<bool>,
    #[serde(rename = "@m")]
    #[serde(default)]
    pub m: Option<bool>,
    #[serde(rename = "@c")]
    #[serde(default)]
    pub c: Option<bool>,
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
pub struct SmlCTPageFields {
    #[serde(rename = "@count")]
    #[serde(default)]
    pub count: Option<u32>,
    #[serde(rename = "pageField")]
    #[serde(default)]
    pub page_field: Vec<Box<SmlCTPageField>>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct SmlCTPageField {
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
    pub name: Option<SSTXstring>,
    #[serde(rename = "@cap")]
    #[serde(default)]
    pub cap: Option<SSTXstring>,
    #[serde(rename = "extLst")]
    #[serde(default)]
    pub ext_lst: Option<Box<SmlCTExtensionList>>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct SmlCTDataFields {
    #[serde(rename = "@count")]
    #[serde(default)]
    pub count: Option<u32>,
    #[serde(rename = "dataField")]
    #[serde(default)]
    pub data_field: Vec<Box<SmlCTDataField>>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct SmlCTDataField {
    #[serde(rename = "@name")]
    #[serde(default)]
    pub name: Option<SSTXstring>,
    #[serde(rename = "@fld")]
    pub fld: u32,
    #[serde(rename = "@subtotal")]
    #[serde(default)]
    pub subtotal: Option<SmlSTDataConsolidateFunction>,
    #[serde(rename = "@showDataAs")]
    #[serde(default)]
    pub show_data_as: Option<SmlSTShowDataAs>,
    #[serde(rename = "@baseField")]
    #[serde(default)]
    pub base_field: Option<i32>,
    #[serde(rename = "@baseItem")]
    #[serde(default)]
    pub base_item: Option<u32>,
    #[serde(rename = "@numFmtId")]
    #[serde(default)]
    pub num_fmt_id: Option<SmlSTNumFmtId>,
    #[serde(rename = "extLst")]
    #[serde(default)]
    pub ext_lst: Option<Box<SmlCTExtensionList>>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct SmlCTRowItems {
    #[serde(rename = "@count")]
    #[serde(default)]
    pub count: Option<u32>,
    #[serde(rename = "i")]
    #[serde(default)]
    pub i: Vec<Box<SmlCTI>>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct SmlCTColItems {
    #[serde(rename = "@count")]
    #[serde(default)]
    pub count: Option<u32>,
    #[serde(rename = "i")]
    #[serde(default)]
    pub i: Vec<Box<SmlCTI>>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct SmlCTI {
    #[serde(rename = "@t")]
    #[serde(default)]
    pub t: Option<SmlSTItemType>,
    #[serde(rename = "@r")]
    #[serde(default)]
    pub r: Option<u32>,
    #[serde(rename = "@i")]
    #[serde(default)]
    pub i: Option<u32>,
    #[serde(rename = "x")]
    #[serde(default)]
    pub x: Vec<Box<SmlCTX>>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct SmlCTX {
    #[serde(rename = "@v")]
    #[serde(default)]
    pub v: Option<i32>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct SmlCTRowFields {
    #[serde(rename = "@count")]
    #[serde(default)]
    pub count: Option<u32>,
    #[serde(rename = "field")]
    #[serde(default)]
    pub field: Vec<Box<SmlCTField>>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct SmlCTColFields {
    #[serde(rename = "@count")]
    #[serde(default)]
    pub count: Option<u32>,
    #[serde(rename = "field")]
    #[serde(default)]
    pub field: Vec<Box<SmlCTField>>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct SmlCTField {
    #[serde(rename = "@x")]
    pub x: i32,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct SmlCTFormats {
    #[serde(rename = "@count")]
    #[serde(default)]
    pub count: Option<u32>,
    #[serde(rename = "format")]
    #[serde(default)]
    pub format: Vec<Box<SmlCTFormat>>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct SmlCTFormat {
    #[serde(rename = "@action")]
    #[serde(default)]
    pub action: Option<SmlSTFormatAction>,
    #[serde(rename = "@dxfId")]
    #[serde(default)]
    pub dxf_id: Option<SmlSTDxfId>,
    #[serde(rename = "pivotArea")]
    pub pivot_area: Box<SmlCTPivotArea>,
    #[serde(rename = "extLst")]
    #[serde(default)]
    pub ext_lst: Option<Box<SmlCTExtensionList>>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct SmlCTConditionalFormats {
    #[serde(rename = "@count")]
    #[serde(default)]
    pub count: Option<u32>,
    #[serde(rename = "conditionalFormat")]
    #[serde(default)]
    pub conditional_format: Vec<Box<SmlCTConditionalFormat>>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct SmlCTConditionalFormat {
    #[serde(rename = "@scope")]
    #[serde(default)]
    pub scope: Option<SmlSTScope>,
    #[serde(rename = "@type")]
    #[serde(default)]
    pub r#type: Option<SmlSTType>,
    #[serde(rename = "@priority")]
    pub priority: u32,
    #[serde(rename = "pivotAreas")]
    pub pivot_areas: Box<SmlCTPivotAreas>,
    #[serde(rename = "extLst")]
    #[serde(default)]
    pub ext_lst: Option<Box<SmlCTExtensionList>>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct SmlCTPivotAreas {
    #[serde(rename = "@count")]
    #[serde(default)]
    pub count: Option<u32>,
    #[serde(rename = "pivotArea")]
    #[serde(default)]
    pub pivot_area: Vec<Box<SmlCTPivotArea>>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct SmlCTChartFormats {
    #[serde(rename = "@count")]
    #[serde(default)]
    pub count: Option<u32>,
    #[serde(rename = "chartFormat")]
    #[serde(default)]
    pub chart_format: Vec<Box<SmlCTChartFormat>>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct SmlCTChartFormat {
    #[serde(rename = "@chart")]
    pub chart: u32,
    #[serde(rename = "@format")]
    pub format: u32,
    #[serde(rename = "@series")]
    #[serde(default)]
    pub series: Option<bool>,
    #[serde(rename = "pivotArea")]
    pub pivot_area: Box<SmlCTPivotArea>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct SmlCTPivotHierarchies {
    #[serde(rename = "@count")]
    #[serde(default)]
    pub count: Option<u32>,
    #[serde(rename = "pivotHierarchy")]
    #[serde(default)]
    pub pivot_hierarchy: Vec<Box<SmlCTPivotHierarchy>>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct SmlCTPivotHierarchy {
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
    pub caption: Option<SSTXstring>,
    #[serde(rename = "mps")]
    #[serde(default)]
    pub mps: Option<Box<SmlCTMemberProperties>>,
    #[serde(rename = "members")]
    #[serde(default)]
    pub members: Vec<Box<SmlCTMembers>>,
    #[serde(rename = "extLst")]
    #[serde(default)]
    pub ext_lst: Option<Box<SmlCTExtensionList>>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct SmlCTRowHierarchiesUsage {
    #[serde(rename = "@count")]
    #[serde(default)]
    pub count: Option<u32>,
    #[serde(rename = "rowHierarchyUsage")]
    #[serde(default)]
    pub row_hierarchy_usage: Vec<Box<SmlCTHierarchyUsage>>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct SmlCTColHierarchiesUsage {
    #[serde(rename = "@count")]
    #[serde(default)]
    pub count: Option<u32>,
    #[serde(rename = "colHierarchyUsage")]
    #[serde(default)]
    pub col_hierarchy_usage: Vec<Box<SmlCTHierarchyUsage>>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct SmlCTHierarchyUsage {
    #[serde(rename = "@hierarchyUsage")]
    pub hierarchy_usage: i32,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct SmlCTMemberProperties {
    #[serde(rename = "@count")]
    #[serde(default)]
    pub count: Option<u32>,
    #[serde(rename = "mp")]
    #[serde(default)]
    pub mp: Vec<Box<SmlCTMemberProperty>>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct SmlCTMemberProperty {
    #[serde(rename = "@name")]
    #[serde(default)]
    pub name: Option<SSTXstring>,
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
pub struct SmlCTMembers {
    #[serde(rename = "@count")]
    #[serde(default)]
    pub count: Option<u32>,
    #[serde(rename = "@level")]
    #[serde(default)]
    pub level: Option<u32>,
    #[serde(rename = "member")]
    #[serde(default)]
    pub member: Vec<Box<SmlCTMember>>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct SmlCTMember {
    #[serde(rename = "@name")]
    pub name: SSTXstring,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct SmlCTDimensions {
    #[serde(rename = "@count")]
    #[serde(default)]
    pub count: Option<u32>,
    #[serde(rename = "dimension")]
    #[serde(default)]
    pub dimension: Vec<Box<SmlCTPivotDimension>>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct SmlCTPivotDimension {
    #[serde(rename = "@measure")]
    #[serde(default)]
    pub measure: Option<bool>,
    #[serde(rename = "@name")]
    pub name: SSTXstring,
    #[serde(rename = "@uniqueName")]
    pub unique_name: SSTXstring,
    #[serde(rename = "@caption")]
    pub caption: SSTXstring,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct SmlCTMeasureGroups {
    #[serde(rename = "@count")]
    #[serde(default)]
    pub count: Option<u32>,
    #[serde(rename = "measureGroup")]
    #[serde(default)]
    pub measure_group: Vec<Box<SmlCTMeasureGroup>>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct SmlCTMeasureDimensionMaps {
    #[serde(rename = "@count")]
    #[serde(default)]
    pub count: Option<u32>,
    #[serde(rename = "map")]
    #[serde(default)]
    pub map: Vec<Box<SmlCTMeasureDimensionMap>>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct SmlCTMeasureGroup {
    #[serde(rename = "@name")]
    pub name: SSTXstring,
    #[serde(rename = "@caption")]
    pub caption: SSTXstring,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct SmlCTMeasureDimensionMap {
    #[serde(rename = "@measureGroup")]
    #[serde(default)]
    pub measure_group: Option<u32>,
    #[serde(rename = "@dimension")]
    #[serde(default)]
    pub dimension: Option<u32>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct SmlCTPivotTableStyle {
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
pub struct SmlCTPivotFilters {
    #[serde(rename = "@count")]
    #[serde(default)]
    pub count: Option<u32>,
    #[serde(rename = "filter")]
    #[serde(default)]
    pub filter: Vec<Box<SmlCTPivotFilter>>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct SmlCTPivotFilter {
    #[serde(rename = "@fld")]
    pub fld: u32,
    #[serde(rename = "@mpFld")]
    #[serde(default)]
    pub mp_fld: Option<u32>,
    #[serde(rename = "@type")]
    pub r#type: SmlSTPivotFilterType,
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
    pub name: Option<SSTXstring>,
    #[serde(rename = "@description")]
    #[serde(default)]
    pub description: Option<SSTXstring>,
    #[serde(rename = "@stringValue1")]
    #[serde(default)]
    pub string_value1: Option<SSTXstring>,
    #[serde(rename = "@stringValue2")]
    #[serde(default)]
    pub string_value2: Option<SSTXstring>,
    #[serde(rename = "autoFilter")]
    pub auto_filter: Box<SmlCTAutoFilter>,
    #[serde(rename = "extLst")]
    #[serde(default)]
    pub ext_lst: Option<Box<SmlCTExtensionList>>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct SmlCTPivotArea {
    #[serde(rename = "@field")]
    #[serde(default)]
    pub field: Option<i32>,
    #[serde(rename = "@type")]
    #[serde(default)]
    pub r#type: Option<SmlSTPivotAreaType>,
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
    pub offset: Option<SmlSTRef>,
    #[serde(rename = "@collapsedLevelsAreSubtotals")]
    #[serde(default)]
    pub collapsed_levels_are_subtotals: Option<bool>,
    #[serde(rename = "@axis")]
    #[serde(default)]
    pub axis: Option<SmlSTAxis>,
    #[serde(rename = "@fieldPosition")]
    #[serde(default)]
    pub field_position: Option<u32>,
    #[serde(rename = "references")]
    #[serde(default)]
    pub references: Option<Box<SmlCTPivotAreaReferences>>,
    #[serde(rename = "extLst")]
    #[serde(default)]
    pub ext_lst: Option<Box<SmlCTExtensionList>>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct SmlCTPivotAreaReferences {
    #[serde(rename = "@count")]
    #[serde(default)]
    pub count: Option<u32>,
    #[serde(rename = "reference")]
    #[serde(default)]
    pub reference: Vec<Box<SmlCTPivotAreaReference>>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct SmlCTPivotAreaReference {
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
    pub x: Vec<Box<SmlCTIndex>>,
    #[serde(rename = "extLst")]
    #[serde(default)]
    pub ext_lst: Option<Box<SmlCTExtensionList>>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct SmlCTIndex {
    #[serde(rename = "@v")]
    pub v: u32,
}

pub type SmlQueryTable = Box<SmlCTQueryTable>;

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct SmlCTQueryTable {
    #[serde(rename = "@name")]
    pub name: SSTXstring,
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
    pub grow_shrink_type: Option<SmlSTGrowShrinkType>,
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
    pub query_table_refresh: Option<Box<SmlCTQueryTableRefresh>>,
    #[serde(rename = "extLst")]
    #[serde(default)]
    pub ext_lst: Option<Box<SmlCTExtensionList>>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct SmlCTQueryTableRefresh {
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
    pub query_table_fields: Box<SmlCTQueryTableFields>,
    #[serde(rename = "queryTableDeletedFields")]
    #[serde(default)]
    pub query_table_deleted_fields: Option<Box<SmlCTQueryTableDeletedFields>>,
    #[serde(rename = "sortState")]
    #[serde(default)]
    pub sort_state: Option<Box<SmlCTSortState>>,
    #[serde(rename = "extLst")]
    #[serde(default)]
    pub ext_lst: Option<Box<SmlCTExtensionList>>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct SmlCTQueryTableDeletedFields {
    #[serde(rename = "@count")]
    #[serde(default)]
    pub count: Option<u32>,
    #[serde(rename = "deletedField")]
    #[serde(default)]
    pub deleted_field: Vec<Box<SmlCTDeletedField>>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct SmlCTDeletedField {
    #[serde(rename = "@name")]
    pub name: SSTXstring,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct SmlCTQueryTableFields {
    #[serde(rename = "@count")]
    #[serde(default)]
    pub count: Option<u32>,
    #[serde(rename = "queryTableField")]
    #[serde(default)]
    pub query_table_field: Vec<Box<SmlCTQueryTableField>>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct SmlCTQueryTableField {
    #[serde(rename = "@id")]
    pub id: u32,
    #[serde(rename = "@name")]
    #[serde(default)]
    pub name: Option<SSTXstring>,
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
    pub ext_lst: Option<Box<SmlCTExtensionList>>,
}

pub type SmlSst = Box<SmlCTSst>;

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct SmlCTSst {
    #[serde(rename = "@count")]
    #[serde(default)]
    pub count: Option<u32>,
    #[serde(rename = "@uniqueCount")]
    #[serde(default)]
    pub unique_count: Option<u32>,
    #[serde(rename = "si")]
    #[serde(default)]
    pub si: Vec<Box<SmlCTRst>>,
    #[serde(rename = "extLst")]
    #[serde(default)]
    pub ext_lst: Option<Box<SmlCTExtensionList>>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct SmlCTPhoneticRun {
    #[serde(rename = "@sb")]
    pub sb: u32,
    #[serde(rename = "@eb")]
    pub eb: u32,
    #[serde(rename = "t")]
    pub t: SSTXstring,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct SmlCTRElt {
    #[serde(rename = "rPr")]
    #[serde(default)]
    pub r_pr: Option<Box<SmlCTRPrElt>>,
    #[serde(rename = "t")]
    pub t: SSTXstring,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct SmlCTRPrElt;

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct SmlCTRst {
    #[serde(rename = "t")]
    #[serde(default)]
    pub t: Option<SSTXstring>,
    #[serde(rename = "r")]
    #[serde(default)]
    pub r: Vec<Box<SmlCTRElt>>,
    #[serde(rename = "rPh")]
    #[serde(default)]
    pub r_ph: Vec<Box<SmlCTPhoneticRun>>,
    #[serde(rename = "phoneticPr")]
    #[serde(default)]
    pub phonetic_pr: Option<Box<SmlCTPhoneticPr>>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct SmlCTPhoneticPr {
    #[serde(rename = "@fontId")]
    pub font_id: SmlSTFontId,
    #[serde(rename = "@type")]
    #[serde(default)]
    pub r#type: Option<SmlSTPhoneticType>,
    #[serde(rename = "@alignment")]
    #[serde(default)]
    pub alignment: Option<SmlSTPhoneticAlignment>,
}

pub type SmlHeaders = Box<SmlCTRevisionHeaders>;

pub type SmlRevisions = Box<SmlCTRevisions>;

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct SmlCTRevisionHeaders {
    #[serde(rename = "@guid")]
    pub guid: SSTGuid,
    #[serde(rename = "@lastGuid")]
    #[serde(default)]
    pub last_guid: Option<SSTGuid>,
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
    pub header: Vec<Box<SmlCTRevisionHeader>>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct SmlCTRevisions;

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
pub struct SmlCTRevisionHeader {
    #[serde(rename = "@guid")]
    pub guid: SSTGuid,
    #[serde(rename = "@dateTime")]
    pub date_time: String,
    #[serde(rename = "@maxSheetId")]
    pub max_sheet_id: u32,
    #[serde(rename = "@userName")]
    pub user_name: SSTXstring,
    #[serde(rename = "@minRId")]
    #[serde(default)]
    pub min_r_id: Option<u32>,
    #[serde(rename = "@maxRId")]
    #[serde(default)]
    pub max_r_id: Option<u32>,
    #[serde(rename = "sheetIdMap")]
    pub sheet_id_map: Box<SmlCTSheetIdMap>,
    #[serde(rename = "reviewedList")]
    #[serde(default)]
    pub reviewed_list: Option<Box<SmlCTReviewedRevisions>>,
    #[serde(rename = "extLst")]
    #[serde(default)]
    pub ext_lst: Option<Box<SmlCTExtensionList>>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct SmlCTSheetIdMap {
    #[serde(rename = "@count")]
    #[serde(default)]
    pub count: Option<u32>,
    #[serde(rename = "sheetId")]
    #[serde(default)]
    pub sheet_id: Vec<Box<SmlCTSheetId>>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct SmlCTSheetId {
    #[serde(rename = "@val")]
    pub val: u32,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct SmlCTReviewedRevisions {
    #[serde(rename = "@count")]
    #[serde(default)]
    pub count: Option<u32>,
    #[serde(rename = "reviewed")]
    #[serde(default)]
    pub reviewed: Vec<Box<SmlCTReviewed>>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct SmlCTReviewed {
    #[serde(rename = "@rId")]
    pub r_id: u32,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct SmlCTUndoInfo {
    #[serde(rename = "@index")]
    pub index: u32,
    #[serde(rename = "@exp")]
    pub exp: SmlSTFormulaExpression,
    #[serde(rename = "@ref3D")]
    #[serde(default)]
    pub ref3_d: Option<bool>,
    #[serde(rename = "@array")]
    #[serde(default)]
    pub array: Option<bool>,
    #[serde(rename = "@v")]
    #[serde(default)]
    pub v: Option<bool>,
    #[serde(rename = "@nf")]
    #[serde(default)]
    pub nf: Option<bool>,
    #[serde(rename = "@cs")]
    #[serde(default)]
    pub cs: Option<bool>,
    #[serde(rename = "@dr")]
    pub dr: SmlSTRefA,
    #[serde(rename = "@dn")]
    #[serde(default)]
    pub dn: Option<SSTXstring>,
    #[serde(rename = "@r")]
    #[serde(default)]
    pub r: Option<SmlSTCellRef>,
    #[serde(rename = "@sId")]
    #[serde(default)]
    pub s_id: Option<u32>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct SmlCTRevisionRowColumn {
    #[serde(rename = "@sId")]
    pub s_id: u32,
    #[serde(rename = "@eol")]
    #[serde(default)]
    pub eol: Option<bool>,
    #[serde(rename = "@ref")]
    pub r#ref: SmlSTRef,
    #[serde(rename = "@action")]
    pub action: SmlSTRwColActionType,
    #[serde(rename = "@edge")]
    #[serde(default)]
    pub edge: Option<bool>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct SmlCTRevisionMove {
    #[serde(rename = "@sheetId")]
    pub sheet_id: u32,
    #[serde(rename = "@source")]
    pub source: SmlSTRef,
    #[serde(rename = "@destination")]
    pub destination: SmlSTRef,
    #[serde(rename = "@sourceSheetId")]
    #[serde(default)]
    pub source_sheet_id: Option<u32>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct SmlCTRevisionCustomView {
    #[serde(rename = "@guid")]
    pub guid: SSTGuid,
    #[serde(rename = "@action")]
    pub action: SmlSTRevisionAction,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct SmlCTRevisionSheetRename {
    #[serde(rename = "@sheetId")]
    pub sheet_id: u32,
    #[serde(rename = "@oldName")]
    pub old_name: SSTXstring,
    #[serde(rename = "@newName")]
    pub new_name: SSTXstring,
    #[serde(rename = "extLst")]
    #[serde(default)]
    pub ext_lst: Option<Box<SmlCTExtensionList>>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct SmlCTRevisionInsertSheet {
    #[serde(rename = "@sheetId")]
    pub sheet_id: u32,
    #[serde(rename = "@name")]
    pub name: SSTXstring,
    #[serde(rename = "@sheetPosition")]
    pub sheet_position: u32,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct SmlCTRevisionCellChange {
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
    pub s: Option<bool>,
    #[serde(rename = "@dxf")]
    #[serde(default)]
    pub dxf: Option<bool>,
    #[serde(rename = "@numFmtId")]
    #[serde(default)]
    pub num_fmt_id: Option<SmlSTNumFmtId>,
    #[serde(rename = "@quotePrefix")]
    #[serde(default)]
    pub quote_prefix: Option<bool>,
    #[serde(rename = "@oldQuotePrefix")]
    #[serde(default)]
    pub old_quote_prefix: Option<bool>,
    #[serde(rename = "@ph")]
    #[serde(default)]
    pub ph: Option<bool>,
    #[serde(rename = "@oldPh")]
    #[serde(default)]
    pub old_ph: Option<bool>,
    #[serde(rename = "@endOfListFormulaUpdate")]
    #[serde(default)]
    pub end_of_list_formula_update: Option<bool>,
    #[serde(rename = "oc")]
    #[serde(default)]
    pub oc: Option<Box<SmlCTCell>>,
    #[serde(rename = "nc")]
    pub nc: Box<SmlCTCell>,
    #[serde(rename = "ndxf")]
    #[serde(default)]
    pub ndxf: Option<Box<SmlCTDxf>>,
    #[serde(rename = "extLst")]
    #[serde(default)]
    pub ext_lst: Option<Box<SmlCTExtensionList>>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct SmlCTRevisionFormatting {
    #[serde(rename = "@sheetId")]
    pub sheet_id: u32,
    #[serde(rename = "@xfDxf")]
    #[serde(default)]
    pub xf_dxf: Option<bool>,
    #[serde(rename = "@s")]
    #[serde(default)]
    pub s: Option<bool>,
    #[serde(rename = "@sqref")]
    pub sqref: SmlSTSqref,
    #[serde(rename = "@start")]
    #[serde(default)]
    pub start: Option<u32>,
    #[serde(rename = "@length")]
    #[serde(default)]
    pub length: Option<u32>,
    #[serde(rename = "dxf")]
    #[serde(default)]
    pub dxf: Option<Box<SmlCTDxf>>,
    #[serde(rename = "extLst")]
    #[serde(default)]
    pub ext_lst: Option<Box<SmlCTExtensionList>>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct SmlCTRevisionAutoFormatting {
    #[serde(rename = "@sheetId")]
    pub sheet_id: u32,
    #[serde(rename = "@ref")]
    pub r#ref: SmlSTRef,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct SmlCTRevisionComment {
    #[serde(rename = "@sheetId")]
    pub sheet_id: u32,
    #[serde(rename = "@cell")]
    pub cell: SmlSTCellRef,
    #[serde(rename = "@guid")]
    pub guid: SSTGuid,
    #[serde(rename = "@action")]
    #[serde(default)]
    pub action: Option<SmlSTRevisionAction>,
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
    pub author: SSTXstring,
    #[serde(rename = "@oldLength")]
    #[serde(default)]
    pub old_length: Option<u32>,
    #[serde(rename = "@newLength")]
    #[serde(default)]
    pub new_length: Option<u32>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct SmlCTRevisionDefinedName {
    #[serde(rename = "@localSheetId")]
    #[serde(default)]
    pub local_sheet_id: Option<u32>,
    #[serde(rename = "@customView")]
    #[serde(default)]
    pub custom_view: Option<bool>,
    #[serde(rename = "@name")]
    pub name: SSTXstring,
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
    pub custom_menu: Option<SSTXstring>,
    #[serde(rename = "@oldCustomMenu")]
    #[serde(default)]
    pub old_custom_menu: Option<SSTXstring>,
    #[serde(rename = "@description")]
    #[serde(default)]
    pub description: Option<SSTXstring>,
    #[serde(rename = "@oldDescription")]
    #[serde(default)]
    pub old_description: Option<SSTXstring>,
    #[serde(rename = "@help")]
    #[serde(default)]
    pub help: Option<SSTXstring>,
    #[serde(rename = "@oldHelp")]
    #[serde(default)]
    pub old_help: Option<SSTXstring>,
    #[serde(rename = "@statusBar")]
    #[serde(default)]
    pub status_bar: Option<SSTXstring>,
    #[serde(rename = "@oldStatusBar")]
    #[serde(default)]
    pub old_status_bar: Option<SSTXstring>,
    #[serde(rename = "@comment")]
    #[serde(default)]
    pub comment: Option<SSTXstring>,
    #[serde(rename = "@oldComment")]
    #[serde(default)]
    pub old_comment: Option<SSTXstring>,
    #[serde(rename = "formula")]
    #[serde(default)]
    pub formula: Option<SmlSTFormula>,
    #[serde(rename = "oldFormula")]
    #[serde(default)]
    pub old_formula: Option<SmlSTFormula>,
    #[serde(rename = "extLst")]
    #[serde(default)]
    pub ext_lst: Option<Box<SmlCTExtensionList>>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct SmlCTRevisionConflict {
    #[serde(rename = "@sheetId")]
    #[serde(default)]
    pub sheet_id: Option<u32>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct SmlCTRevisionQueryTableField {
    #[serde(rename = "@sheetId")]
    pub sheet_id: u32,
    #[serde(rename = "@ref")]
    pub r#ref: SmlSTRef,
    #[serde(rename = "@fieldId")]
    pub field_id: u32,
}

pub type SmlUsers = Box<SmlCTUsers>;

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct SmlCTUsers {
    #[serde(rename = "@count")]
    #[serde(default)]
    pub count: Option<u32>,
    #[serde(rename = "userInfo")]
    #[serde(default)]
    pub user_info: Vec<Box<SmlCTSharedUser>>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct SmlCTSharedUser {
    #[serde(rename = "@guid")]
    pub guid: SSTGuid,
    #[serde(rename = "@name")]
    pub name: SSTXstring,
    #[serde(rename = "@id")]
    pub id: i32,
    #[serde(rename = "@dateTime")]
    pub date_time: String,
    #[serde(rename = "extLst")]
    #[serde(default)]
    pub ext_lst: Option<Box<SmlCTExtensionList>>,
}

pub type SmlWorksheet = Box<SmlCTWorksheet>;

pub type SmlChartsheet = Box<SmlCTChartsheet>;

pub type SmlDialogsheet = Box<SmlCTDialogsheet>;

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct SmlCTMacrosheet {
    #[serde(rename = "sheetPr")]
    #[serde(default)]
    pub sheet_pr: Option<Box<SmlCTSheetPr>>,
    #[serde(rename = "dimension")]
    #[serde(default)]
    pub dimension: Option<Box<SmlCTSheetDimension>>,
    #[serde(rename = "sheetViews")]
    #[serde(default)]
    pub sheet_views: Option<Box<SmlCTSheetViews>>,
    #[serde(rename = "sheetFormatPr")]
    #[serde(default)]
    pub sheet_format_pr: Option<Box<SmlCTSheetFormatPr>>,
    #[serde(rename = "cols")]
    #[serde(default)]
    pub cols: Vec<Box<SmlCTCols>>,
    #[serde(rename = "sheetData")]
    pub sheet_data: Box<SmlCTSheetData>,
    #[serde(rename = "sheetProtection")]
    #[serde(default)]
    pub sheet_protection: Option<Box<SmlCTSheetProtection>>,
    #[serde(rename = "autoFilter")]
    #[serde(default)]
    pub auto_filter: Option<Box<SmlCTAutoFilter>>,
    #[serde(rename = "sortState")]
    #[serde(default)]
    pub sort_state: Option<Box<SmlCTSortState>>,
    #[serde(rename = "dataConsolidate")]
    #[serde(default)]
    pub data_consolidate: Option<Box<SmlCTDataConsolidate>>,
    #[serde(rename = "customSheetViews")]
    #[serde(default)]
    pub custom_sheet_views: Option<Box<SmlCTCustomSheetViews>>,
    #[serde(rename = "phoneticPr")]
    #[serde(default)]
    pub phonetic_pr: Option<Box<SmlCTPhoneticPr>>,
    #[serde(rename = "conditionalFormatting")]
    #[serde(default)]
    pub conditional_formatting: Vec<Box<SmlCTConditionalFormatting>>,
    #[serde(rename = "printOptions")]
    #[serde(default)]
    pub print_options: Option<Box<SmlCTPrintOptions>>,
    #[serde(rename = "pageMargins")]
    #[serde(default)]
    pub page_margins: Option<Box<SmlCTPageMargins>>,
    #[serde(rename = "pageSetup")]
    #[serde(default)]
    pub page_setup: Option<Box<SmlCTPageSetup>>,
    #[serde(rename = "headerFooter")]
    #[serde(default)]
    pub header_footer: Option<Box<SmlCTHeaderFooter>>,
    #[serde(rename = "rowBreaks")]
    #[serde(default)]
    pub row_breaks: Option<Box<SmlCTPageBreak>>,
    #[serde(rename = "colBreaks")]
    #[serde(default)]
    pub col_breaks: Option<Box<SmlCTPageBreak>>,
    #[serde(rename = "customProperties")]
    #[serde(default)]
    pub custom_properties: Option<Box<SmlCTCustomProperties>>,
    #[serde(rename = "drawing")]
    #[serde(default)]
    pub drawing: Option<Box<SmlCTDrawing>>,
    #[serde(rename = "legacyDrawing")]
    #[serde(default)]
    pub legacy_drawing: Option<Box<SmlCTLegacyDrawing>>,
    #[serde(rename = "legacyDrawingHF")]
    #[serde(default)]
    pub legacy_drawing_h_f: Option<Box<SmlCTLegacyDrawing>>,
    #[serde(rename = "drawingHF")]
    #[serde(default)]
    pub drawing_h_f: Option<Box<SmlCTDrawingHF>>,
    #[serde(rename = "picture")]
    #[serde(default)]
    pub picture: Option<Box<SmlCTSheetBackgroundPicture>>,
    #[serde(rename = "oleObjects")]
    #[serde(default)]
    pub ole_objects: Option<Box<SmlCTOleObjects>>,
    #[serde(rename = "extLst")]
    #[serde(default)]
    pub ext_lst: Option<Box<SmlCTExtensionList>>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct SmlCTDialogsheet {
    #[serde(rename = "sheetPr")]
    #[serde(default)]
    pub sheet_pr: Option<Box<SmlCTSheetPr>>,
    #[serde(rename = "sheetViews")]
    #[serde(default)]
    pub sheet_views: Option<Box<SmlCTSheetViews>>,
    #[serde(rename = "sheetFormatPr")]
    #[serde(default)]
    pub sheet_format_pr: Option<Box<SmlCTSheetFormatPr>>,
    #[serde(rename = "sheetProtection")]
    #[serde(default)]
    pub sheet_protection: Option<Box<SmlCTSheetProtection>>,
    #[serde(rename = "customSheetViews")]
    #[serde(default)]
    pub custom_sheet_views: Option<Box<SmlCTCustomSheetViews>>,
    #[serde(rename = "printOptions")]
    #[serde(default)]
    pub print_options: Option<Box<SmlCTPrintOptions>>,
    #[serde(rename = "pageMargins")]
    #[serde(default)]
    pub page_margins: Option<Box<SmlCTPageMargins>>,
    #[serde(rename = "pageSetup")]
    #[serde(default)]
    pub page_setup: Option<Box<SmlCTPageSetup>>,
    #[serde(rename = "headerFooter")]
    #[serde(default)]
    pub header_footer: Option<Box<SmlCTHeaderFooter>>,
    #[serde(rename = "drawing")]
    #[serde(default)]
    pub drawing: Option<Box<SmlCTDrawing>>,
    #[serde(rename = "legacyDrawing")]
    #[serde(default)]
    pub legacy_drawing: Option<Box<SmlCTLegacyDrawing>>,
    #[serde(rename = "legacyDrawingHF")]
    #[serde(default)]
    pub legacy_drawing_h_f: Option<Box<SmlCTLegacyDrawing>>,
    #[serde(rename = "drawingHF")]
    #[serde(default)]
    pub drawing_h_f: Option<Box<SmlCTDrawingHF>>,
    #[serde(rename = "oleObjects")]
    #[serde(default)]
    pub ole_objects: Option<Box<SmlCTOleObjects>>,
    #[serde(rename = "controls")]
    #[serde(default)]
    pub controls: Option<Box<SmlCTControls>>,
    #[serde(rename = "extLst")]
    #[serde(default)]
    pub ext_lst: Option<Box<SmlCTExtensionList>>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct SmlCTWorksheet {
    #[serde(rename = "sheetPr")]
    #[serde(default)]
    pub sheet_pr: Option<Box<SmlCTSheetPr>>,
    #[serde(rename = "dimension")]
    #[serde(default)]
    pub dimension: Option<Box<SmlCTSheetDimension>>,
    #[serde(rename = "sheetViews")]
    #[serde(default)]
    pub sheet_views: Option<Box<SmlCTSheetViews>>,
    #[serde(rename = "sheetFormatPr")]
    #[serde(default)]
    pub sheet_format_pr: Option<Box<SmlCTSheetFormatPr>>,
    #[serde(rename = "cols")]
    #[serde(default)]
    pub cols: Vec<Box<SmlCTCols>>,
    #[serde(rename = "sheetData")]
    pub sheet_data: Box<SmlCTSheetData>,
    #[serde(rename = "sheetCalcPr")]
    #[serde(default)]
    pub sheet_calc_pr: Option<Box<SmlCTSheetCalcPr>>,
    #[serde(rename = "sheetProtection")]
    #[serde(default)]
    pub sheet_protection: Option<Box<SmlCTSheetProtection>>,
    #[serde(rename = "protectedRanges")]
    #[serde(default)]
    pub protected_ranges: Option<Box<SmlCTProtectedRanges>>,
    #[serde(rename = "scenarios")]
    #[serde(default)]
    pub scenarios: Option<Box<SmlCTScenarios>>,
    #[serde(rename = "autoFilter")]
    #[serde(default)]
    pub auto_filter: Option<Box<SmlCTAutoFilter>>,
    #[serde(rename = "sortState")]
    #[serde(default)]
    pub sort_state: Option<Box<SmlCTSortState>>,
    #[serde(rename = "dataConsolidate")]
    #[serde(default)]
    pub data_consolidate: Option<Box<SmlCTDataConsolidate>>,
    #[serde(rename = "customSheetViews")]
    #[serde(default)]
    pub custom_sheet_views: Option<Box<SmlCTCustomSheetViews>>,
    #[serde(rename = "mergeCells")]
    #[serde(default)]
    pub merge_cells: Option<Box<SmlCTMergeCells>>,
    #[serde(rename = "phoneticPr")]
    #[serde(default)]
    pub phonetic_pr: Option<Box<SmlCTPhoneticPr>>,
    #[serde(rename = "conditionalFormatting")]
    #[serde(default)]
    pub conditional_formatting: Vec<Box<SmlCTConditionalFormatting>>,
    #[serde(rename = "dataValidations")]
    #[serde(default)]
    pub data_validations: Option<Box<SmlCTDataValidations>>,
    #[serde(rename = "hyperlinks")]
    #[serde(default)]
    pub hyperlinks: Option<Box<SmlCTHyperlinks>>,
    #[serde(rename = "printOptions")]
    #[serde(default)]
    pub print_options: Option<Box<SmlCTPrintOptions>>,
    #[serde(rename = "pageMargins")]
    #[serde(default)]
    pub page_margins: Option<Box<SmlCTPageMargins>>,
    #[serde(rename = "pageSetup")]
    #[serde(default)]
    pub page_setup: Option<Box<SmlCTPageSetup>>,
    #[serde(rename = "headerFooter")]
    #[serde(default)]
    pub header_footer: Option<Box<SmlCTHeaderFooter>>,
    #[serde(rename = "rowBreaks")]
    #[serde(default)]
    pub row_breaks: Option<Box<SmlCTPageBreak>>,
    #[serde(rename = "colBreaks")]
    #[serde(default)]
    pub col_breaks: Option<Box<SmlCTPageBreak>>,
    #[serde(rename = "customProperties")]
    #[serde(default)]
    pub custom_properties: Option<Box<SmlCTCustomProperties>>,
    #[serde(rename = "cellWatches")]
    #[serde(default)]
    pub cell_watches: Option<Box<SmlCTCellWatches>>,
    #[serde(rename = "ignoredErrors")]
    #[serde(default)]
    pub ignored_errors: Option<Box<SmlCTIgnoredErrors>>,
    #[serde(rename = "smartTags")]
    #[serde(default)]
    pub smart_tags: Option<Box<SmlCTSmartTags>>,
    #[serde(rename = "drawing")]
    #[serde(default)]
    pub drawing: Option<Box<SmlCTDrawing>>,
    #[serde(rename = "legacyDrawing")]
    #[serde(default)]
    pub legacy_drawing: Option<Box<SmlCTLegacyDrawing>>,
    #[serde(rename = "legacyDrawingHF")]
    #[serde(default)]
    pub legacy_drawing_h_f: Option<Box<SmlCTLegacyDrawing>>,
    #[serde(rename = "drawingHF")]
    #[serde(default)]
    pub drawing_h_f: Option<Box<SmlCTDrawingHF>>,
    #[serde(rename = "picture")]
    #[serde(default)]
    pub picture: Option<Box<SmlCTSheetBackgroundPicture>>,
    #[serde(rename = "oleObjects")]
    #[serde(default)]
    pub ole_objects: Option<Box<SmlCTOleObjects>>,
    #[serde(rename = "controls")]
    #[serde(default)]
    pub controls: Option<Box<SmlCTControls>>,
    #[serde(rename = "webPublishItems")]
    #[serde(default)]
    pub web_publish_items: Option<Box<SmlCTWebPublishItems>>,
    #[serde(rename = "tableParts")]
    #[serde(default)]
    pub table_parts: Option<Box<SmlCTTableParts>>,
    #[serde(rename = "extLst")]
    #[serde(default)]
    pub ext_lst: Option<Box<SmlCTExtensionList>>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct SmlCTSheetData {
    #[serde(rename = "row")]
    #[serde(default)]
    pub row: Vec<Box<SmlCTRow>>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct SmlCTSheetCalcPr {
    #[serde(rename = "@fullCalcOnLoad")]
    #[serde(default)]
    pub full_calc_on_load: Option<bool>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct SmlCTSheetFormatPr {
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
pub struct SmlCTCols {
    #[serde(rename = "col")]
    #[serde(default)]
    pub col: Vec<Box<SmlCTCol>>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct SmlCTCol {
    #[serde(rename = "@min")]
    pub min: u32,
    #[serde(rename = "@max")]
    pub max: u32,
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
pub struct SmlCTRow {
    #[serde(rename = "@r")]
    #[serde(default)]
    pub r: Option<u32>,
    #[serde(rename = "@spans")]
    #[serde(default)]
    pub spans: Option<SmlSTCellSpans>,
    #[serde(rename = "@s")]
    #[serde(default)]
    pub s: Option<u32>,
    #[serde(rename = "@customFormat")]
    #[serde(default)]
    pub custom_format: Option<bool>,
    #[serde(rename = "@ht")]
    #[serde(default)]
    pub ht: Option<f64>,
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
    pub ph: Option<bool>,
    #[serde(rename = "c")]
    #[serde(default)]
    pub c: Vec<Box<SmlCTCell>>,
    #[serde(rename = "extLst")]
    #[serde(default)]
    pub ext_lst: Option<Box<SmlCTExtensionList>>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct SmlCTCell {
    #[serde(rename = "@r")]
    #[serde(default)]
    pub r: Option<SmlSTCellRef>,
    #[serde(rename = "@s")]
    #[serde(default)]
    pub s: Option<u32>,
    #[serde(rename = "@t")]
    #[serde(default)]
    pub t: Option<SmlSTCellType>,
    #[serde(rename = "@cm")]
    #[serde(default)]
    pub cm: Option<u32>,
    #[serde(rename = "@vm")]
    #[serde(default)]
    pub vm: Option<u32>,
    #[serde(rename = "@ph")]
    #[serde(default)]
    pub ph: Option<bool>,
    #[serde(rename = "f")]
    #[serde(default)]
    pub f: Option<Box<SmlCTCellFormula>>,
    #[serde(rename = "v")]
    #[serde(default)]
    pub v: Option<SSTXstring>,
    #[serde(rename = "is")]
    #[serde(default)]
    pub is: Option<Box<SmlCTRst>>,
    #[serde(rename = "extLst")]
    #[serde(default)]
    pub ext_lst: Option<Box<SmlCTExtensionList>>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct SmlCTSheetPr {
    #[serde(rename = "@syncHorizontal")]
    #[serde(default)]
    pub sync_horizontal: Option<bool>,
    #[serde(rename = "@syncVertical")]
    #[serde(default)]
    pub sync_vertical: Option<bool>,
    #[serde(rename = "@syncRef")]
    #[serde(default)]
    pub sync_ref: Option<SmlSTRef>,
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
    pub tab_color: Option<Box<SmlCTColor>>,
    #[serde(rename = "outlinePr")]
    #[serde(default)]
    pub outline_pr: Option<Box<SmlCTOutlinePr>>,
    #[serde(rename = "pageSetUpPr")]
    #[serde(default)]
    pub page_set_up_pr: Option<Box<SmlCTPageSetUpPr>>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct SmlCTSheetDimension {
    #[serde(rename = "@ref")]
    pub r#ref: SmlSTRef,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct SmlCTSheetViews {
    #[serde(rename = "sheetView")]
    #[serde(default)]
    pub sheet_view: Vec<Box<SmlCTSheetView>>,
    #[serde(rename = "extLst")]
    #[serde(default)]
    pub ext_lst: Option<Box<SmlCTExtensionList>>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct SmlCTSheetView {
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
    pub view: Option<SmlSTSheetViewType>,
    #[serde(rename = "@topLeftCell")]
    #[serde(default)]
    pub top_left_cell: Option<SmlSTCellRef>,
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
    pub pane: Option<Box<SmlCTPane>>,
    #[serde(rename = "selection")]
    #[serde(default)]
    pub selection: Vec<Box<SmlCTSelection>>,
    #[serde(rename = "pivotSelection")]
    #[serde(default)]
    pub pivot_selection: Vec<Box<SmlCTPivotSelection>>,
    #[serde(rename = "extLst")]
    #[serde(default)]
    pub ext_lst: Option<Box<SmlCTExtensionList>>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct SmlCTPane {
    #[serde(rename = "@xSplit")]
    #[serde(default)]
    pub x_split: Option<f64>,
    #[serde(rename = "@ySplit")]
    #[serde(default)]
    pub y_split: Option<f64>,
    #[serde(rename = "@topLeftCell")]
    #[serde(default)]
    pub top_left_cell: Option<SmlSTCellRef>,
    #[serde(rename = "@activePane")]
    #[serde(default)]
    pub active_pane: Option<SmlSTPane>,
    #[serde(rename = "@state")]
    #[serde(default)]
    pub state: Option<SmlSTPaneState>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct SmlCTPivotSelection {
    #[serde(rename = "@pane")]
    #[serde(default)]
    pub pane: Option<SmlSTPane>,
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
    pub axis: Option<SmlSTAxis>,
    #[serde(rename = "@dimension")]
    #[serde(default)]
    pub dimension: Option<u32>,
    #[serde(rename = "@start")]
    #[serde(default)]
    pub start: Option<u32>,
    #[serde(rename = "@min")]
    #[serde(default)]
    pub min: Option<u32>,
    #[serde(rename = "@max")]
    #[serde(default)]
    pub max: Option<u32>,
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
    pub pivot_area: Box<SmlCTPivotArea>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct SmlCTSelection {
    #[serde(rename = "@pane")]
    #[serde(default)]
    pub pane: Option<SmlSTPane>,
    #[serde(rename = "@activeCell")]
    #[serde(default)]
    pub active_cell: Option<SmlSTCellRef>,
    #[serde(rename = "@activeCellId")]
    #[serde(default)]
    pub active_cell_id: Option<u32>,
    #[serde(rename = "@sqref")]
    #[serde(default)]
    pub sqref: Option<SmlSTSqref>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct SmlCTPageBreak {
    #[serde(rename = "@count")]
    #[serde(default)]
    pub count: Option<u32>,
    #[serde(rename = "@manualBreakCount")]
    #[serde(default)]
    pub manual_break_count: Option<u32>,
    #[serde(rename = "brk")]
    #[serde(default)]
    pub brk: Vec<Box<SmlCTBreak>>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct SmlCTBreak {
    #[serde(rename = "@id")]
    #[serde(default)]
    pub id: Option<u32>,
    #[serde(rename = "@min")]
    #[serde(default)]
    pub min: Option<u32>,
    #[serde(rename = "@max")]
    #[serde(default)]
    pub max: Option<u32>,
    #[serde(rename = "@man")]
    #[serde(default)]
    pub man: Option<bool>,
    #[serde(rename = "@pt")]
    #[serde(default)]
    pub pt: Option<bool>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct SmlCTOutlinePr {
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
pub struct SmlCTPageSetUpPr {
    #[serde(rename = "@autoPageBreaks")]
    #[serde(default)]
    pub auto_page_breaks: Option<bool>,
    #[serde(rename = "@fitToPage")]
    #[serde(default)]
    pub fit_to_page: Option<bool>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct SmlCTDataConsolidate {
    #[serde(rename = "@function")]
    #[serde(default)]
    pub function: Option<SmlSTDataConsolidateFunction>,
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
    pub data_refs: Option<Box<SmlCTDataRefs>>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct SmlCTDataRefs {
    #[serde(rename = "@count")]
    #[serde(default)]
    pub count: Option<u32>,
    #[serde(rename = "dataRef")]
    #[serde(default)]
    pub data_ref: Vec<Box<SmlCTDataRef>>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct SmlCTDataRef {
    #[serde(rename = "@ref")]
    #[serde(default)]
    pub r#ref: Option<SmlSTRef>,
    #[serde(rename = "@name")]
    #[serde(default)]
    pub name: Option<SSTXstring>,
    #[serde(rename = "@sheet")]
    #[serde(default)]
    pub sheet: Option<SSTXstring>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct SmlCTMergeCells {
    #[serde(rename = "@count")]
    #[serde(default)]
    pub count: Option<u32>,
    #[serde(rename = "mergeCell")]
    #[serde(default)]
    pub merge_cell: Vec<Box<SmlCTMergeCell>>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct SmlCTMergeCell {
    #[serde(rename = "@ref")]
    pub r#ref: SmlSTRef,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct SmlCTSmartTags {
    #[serde(rename = "cellSmartTags")]
    #[serde(default)]
    pub cell_smart_tags: Vec<Box<SmlCTCellSmartTags>>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct SmlCTCellSmartTags {
    #[serde(rename = "@r")]
    pub r: SmlSTCellRef,
    #[serde(rename = "cellSmartTag")]
    #[serde(default)]
    pub cell_smart_tag: Vec<Box<SmlCTCellSmartTag>>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct SmlCTCellSmartTag {
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
    pub cell_smart_tag_pr: Vec<Box<SmlCTCellSmartTagPr>>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct SmlCTCellSmartTagPr {
    #[serde(rename = "@key")]
    pub key: SSTXstring,
    #[serde(rename = "@val")]
    pub val: SSTXstring,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct SmlCTDrawing;

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct SmlCTLegacyDrawing;

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct SmlCTDrawingHF {
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
pub struct SmlCTCustomSheetViews {
    #[serde(rename = "customSheetView")]
    #[serde(default)]
    pub custom_sheet_view: Vec<Box<SmlCTCustomSheetView>>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct SmlCTCustomSheetView {
    #[serde(rename = "@guid")]
    pub guid: SSTGuid,
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
    pub state: Option<SmlSTSheetState>,
    #[serde(rename = "@filterUnique")]
    #[serde(default)]
    pub filter_unique: Option<bool>,
    #[serde(rename = "@view")]
    #[serde(default)]
    pub view: Option<SmlSTSheetViewType>,
    #[serde(rename = "@showRuler")]
    #[serde(default)]
    pub show_ruler: Option<bool>,
    #[serde(rename = "@topLeftCell")]
    #[serde(default)]
    pub top_left_cell: Option<SmlSTCellRef>,
    #[serde(rename = "pane")]
    #[serde(default)]
    pub pane: Option<Box<SmlCTPane>>,
    #[serde(rename = "selection")]
    #[serde(default)]
    pub selection: Option<Box<SmlCTSelection>>,
    #[serde(rename = "rowBreaks")]
    #[serde(default)]
    pub row_breaks: Option<Box<SmlCTPageBreak>>,
    #[serde(rename = "colBreaks")]
    #[serde(default)]
    pub col_breaks: Option<Box<SmlCTPageBreak>>,
    #[serde(rename = "pageMargins")]
    #[serde(default)]
    pub page_margins: Option<Box<SmlCTPageMargins>>,
    #[serde(rename = "printOptions")]
    #[serde(default)]
    pub print_options: Option<Box<SmlCTPrintOptions>>,
    #[serde(rename = "pageSetup")]
    #[serde(default)]
    pub page_setup: Option<Box<SmlCTPageSetup>>,
    #[serde(rename = "headerFooter")]
    #[serde(default)]
    pub header_footer: Option<Box<SmlCTHeaderFooter>>,
    #[serde(rename = "autoFilter")]
    #[serde(default)]
    pub auto_filter: Option<Box<SmlCTAutoFilter>>,
    #[serde(rename = "extLst")]
    #[serde(default)]
    pub ext_lst: Option<Box<SmlCTExtensionList>>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct SmlCTDataValidations {
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
    pub data_validation: Vec<Box<SmlCTDataValidation>>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct SmlCTDataValidation {
    #[serde(rename = "@type")]
    #[serde(default)]
    pub r#type: Option<SmlSTDataValidationType>,
    #[serde(rename = "@errorStyle")]
    #[serde(default)]
    pub error_style: Option<SmlSTDataValidationErrorStyle>,
    #[serde(rename = "@imeMode")]
    #[serde(default)]
    pub ime_mode: Option<SmlSTDataValidationImeMode>,
    #[serde(rename = "@operator")]
    #[serde(default)]
    pub operator: Option<SmlSTDataValidationOperator>,
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
    pub error_title: Option<SSTXstring>,
    #[serde(rename = "@error")]
    #[serde(default)]
    pub error: Option<SSTXstring>,
    #[serde(rename = "@promptTitle")]
    #[serde(default)]
    pub prompt_title: Option<SSTXstring>,
    #[serde(rename = "@prompt")]
    #[serde(default)]
    pub prompt: Option<SSTXstring>,
    #[serde(rename = "@sqref")]
    pub sqref: SmlSTSqref,
    #[serde(rename = "formula1")]
    #[serde(default)]
    pub formula1: Option<SmlSTFormula>,
    #[serde(rename = "formula2")]
    #[serde(default)]
    pub formula2: Option<SmlSTFormula>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct SmlCTConditionalFormatting {
    #[serde(rename = "@pivot")]
    #[serde(default)]
    pub pivot: Option<bool>,
    #[serde(rename = "@sqref")]
    #[serde(default)]
    pub sqref: Option<SmlSTSqref>,
    #[serde(rename = "cfRule")]
    #[serde(default)]
    pub cf_rule: Vec<Box<SmlCTCfRule>>,
    #[serde(rename = "extLst")]
    #[serde(default)]
    pub ext_lst: Option<Box<SmlCTExtensionList>>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct SmlCTCfRule {
    #[serde(rename = "@type")]
    #[serde(default)]
    pub r#type: Option<SmlSTCfType>,
    #[serde(rename = "@dxfId")]
    #[serde(default)]
    pub dxf_id: Option<SmlSTDxfId>,
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
    pub operator: Option<SmlSTConditionalFormattingOperator>,
    #[serde(rename = "@text")]
    #[serde(default)]
    pub text: Option<String>,
    #[serde(rename = "@timePeriod")]
    #[serde(default)]
    pub time_period: Option<SmlSTTimePeriod>,
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
    pub formula: Vec<SmlSTFormula>,
    #[serde(rename = "colorScale")]
    #[serde(default)]
    pub color_scale: Option<Box<SmlCTColorScale>>,
    #[serde(rename = "dataBar")]
    #[serde(default)]
    pub data_bar: Option<Box<SmlCTDataBar>>,
    #[serde(rename = "iconSet")]
    #[serde(default)]
    pub icon_set: Option<Box<SmlCTIconSet>>,
    #[serde(rename = "extLst")]
    #[serde(default)]
    pub ext_lst: Option<Box<SmlCTExtensionList>>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct SmlCTHyperlinks {
    #[serde(rename = "hyperlink")]
    #[serde(default)]
    pub hyperlink: Vec<Box<SmlCTHyperlink>>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct SmlCTHyperlink {
    #[serde(rename = "@ref")]
    pub r#ref: SmlSTRef,
    #[serde(rename = "@location")]
    #[serde(default)]
    pub location: Option<SSTXstring>,
    #[serde(rename = "@tooltip")]
    #[serde(default)]
    pub tooltip: Option<SSTXstring>,
    #[serde(rename = "@display")]
    #[serde(default)]
    pub display: Option<SSTXstring>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct SmlCTCellFormula {
    #[serde(rename = "@t")]
    #[serde(default)]
    pub t: Option<SmlSTCellFormulaType>,
    #[serde(rename = "@aca")]
    #[serde(default)]
    pub aca: Option<bool>,
    #[serde(rename = "@ref")]
    #[serde(default)]
    pub r#ref: Option<SmlSTRef>,
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
    pub r1: Option<SmlSTCellRef>,
    #[serde(rename = "@r2")]
    #[serde(default)]
    pub r2: Option<SmlSTCellRef>,
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
pub struct SmlCTColorScale {
    #[serde(rename = "cfvo")]
    #[serde(default)]
    pub cfvo: Vec<Box<SmlCTCfvo>>,
    #[serde(rename = "color")]
    #[serde(default)]
    pub color: Vec<Box<SmlCTColor>>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct SmlCTDataBar {
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
    pub cfvo: Vec<Box<SmlCTCfvo>>,
    #[serde(rename = "color")]
    pub color: Box<SmlCTColor>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct SmlCTIconSet {
    #[serde(rename = "@iconSet")]
    #[serde(default)]
    pub icon_set: Option<SmlSTIconSetType>,
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
    pub cfvo: Vec<Box<SmlCTCfvo>>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct SmlCTCfvo {
    #[serde(rename = "@type")]
    pub r#type: SmlSTCfvoType,
    #[serde(rename = "@val")]
    #[serde(default)]
    pub val: Option<SSTXstring>,
    #[serde(rename = "@gte")]
    #[serde(default)]
    pub gte: Option<bool>,
    #[serde(rename = "extLst")]
    #[serde(default)]
    pub ext_lst: Option<Box<SmlCTExtensionList>>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct SmlCTPageMargins {
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
pub struct SmlCTPrintOptions {
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
pub struct SmlCTPageSetup {
    #[serde(rename = "@paperSize")]
    #[serde(default)]
    pub paper_size: Option<u32>,
    #[serde(rename = "@paperHeight")]
    #[serde(default)]
    pub paper_height: Option<SSTPositiveUniversalMeasure>,
    #[serde(rename = "@paperWidth")]
    #[serde(default)]
    pub paper_width: Option<SSTPositiveUniversalMeasure>,
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
    pub page_order: Option<SmlSTPageOrder>,
    #[serde(rename = "@orientation")]
    #[serde(default)]
    pub orientation: Option<SmlSTOrientation>,
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
    pub cell_comments: Option<SmlSTCellComments>,
    #[serde(rename = "@useFirstPageNumber")]
    #[serde(default)]
    pub use_first_page_number: Option<bool>,
    #[serde(rename = "@errors")]
    #[serde(default)]
    pub errors: Option<SmlSTPrintError>,
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
pub struct SmlCTHeaderFooter {
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
    pub odd_header: Option<SSTXstring>,
    #[serde(rename = "oddFooter")]
    #[serde(default)]
    pub odd_footer: Option<SSTXstring>,
    #[serde(rename = "evenHeader")]
    #[serde(default)]
    pub even_header: Option<SSTXstring>,
    #[serde(rename = "evenFooter")]
    #[serde(default)]
    pub even_footer: Option<SSTXstring>,
    #[serde(rename = "firstHeader")]
    #[serde(default)]
    pub first_header: Option<SSTXstring>,
    #[serde(rename = "firstFooter")]
    #[serde(default)]
    pub first_footer: Option<SSTXstring>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct SmlCTScenarios {
    #[serde(rename = "@current")]
    #[serde(default)]
    pub current: Option<u32>,
    #[serde(rename = "@show")]
    #[serde(default)]
    pub show: Option<u32>,
    #[serde(rename = "@sqref")]
    #[serde(default)]
    pub sqref: Option<SmlSTSqref>,
    #[serde(rename = "scenario")]
    #[serde(default)]
    pub scenario: Vec<Box<SmlCTScenario>>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct SmlCTSheetProtection {
    #[serde(rename = "@password")]
    #[serde(default)]
    pub password: Option<SmlSTUnsignedShortHex>,
    #[serde(rename = "@algorithmName")]
    #[serde(default)]
    pub algorithm_name: Option<SSTXstring>,
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
pub struct SmlCTProtectedRanges {
    #[serde(rename = "protectedRange")]
    #[serde(default)]
    pub protected_range: Vec<Box<SmlCTProtectedRange>>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct SmlCTProtectedRange {
    #[serde(rename = "@password")]
    #[serde(default)]
    pub password: Option<SmlSTUnsignedShortHex>,
    #[serde(rename = "@sqref")]
    pub sqref: SmlSTSqref,
    #[serde(rename = "@name")]
    pub name: SSTXstring,
    #[serde(rename = "@securityDescriptor")]
    #[serde(default)]
    pub security_descriptor: Option<String>,
    #[serde(rename = "@algorithmName")]
    #[serde(default)]
    pub algorithm_name: Option<SSTXstring>,
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
pub struct SmlCTScenario {
    #[serde(rename = "@name")]
    pub name: SSTXstring,
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
    pub user: Option<SSTXstring>,
    #[serde(rename = "@comment")]
    #[serde(default)]
    pub comment: Option<SSTXstring>,
    #[serde(rename = "inputCells")]
    #[serde(default)]
    pub input_cells: Vec<Box<SmlCTInputCells>>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct SmlCTInputCells {
    #[serde(rename = "@r")]
    pub r: SmlSTCellRef,
    #[serde(rename = "@deleted")]
    #[serde(default)]
    pub deleted: Option<bool>,
    #[serde(rename = "@undone")]
    #[serde(default)]
    pub undone: Option<bool>,
    #[serde(rename = "@val")]
    pub val: SSTXstring,
    #[serde(rename = "@numFmtId")]
    #[serde(default)]
    pub num_fmt_id: Option<SmlSTNumFmtId>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct SmlCTCellWatches {
    #[serde(rename = "cellWatch")]
    #[serde(default)]
    pub cell_watch: Vec<Box<SmlCTCellWatch>>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct SmlCTCellWatch {
    #[serde(rename = "@r")]
    pub r: SmlSTCellRef,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct SmlCTChartsheet {
    #[serde(rename = "sheetPr")]
    #[serde(default)]
    pub sheet_pr: Option<Box<SmlCTChartsheetPr>>,
    #[serde(rename = "sheetViews")]
    pub sheet_views: Box<SmlCTChartsheetViews>,
    #[serde(rename = "sheetProtection")]
    #[serde(default)]
    pub sheet_protection: Option<Box<SmlCTChartsheetProtection>>,
    #[serde(rename = "customSheetViews")]
    #[serde(default)]
    pub custom_sheet_views: Option<Box<SmlCTCustomChartsheetViews>>,
    #[serde(rename = "pageMargins")]
    #[serde(default)]
    pub page_margins: Option<Box<SmlCTPageMargins>>,
    #[serde(rename = "pageSetup")]
    #[serde(default)]
    pub page_setup: Option<Box<SmlCTCsPageSetup>>,
    #[serde(rename = "headerFooter")]
    #[serde(default)]
    pub header_footer: Option<Box<SmlCTHeaderFooter>>,
    #[serde(rename = "drawing")]
    pub drawing: Box<SmlCTDrawing>,
    #[serde(rename = "legacyDrawing")]
    #[serde(default)]
    pub legacy_drawing: Option<Box<SmlCTLegacyDrawing>>,
    #[serde(rename = "legacyDrawingHF")]
    #[serde(default)]
    pub legacy_drawing_h_f: Option<Box<SmlCTLegacyDrawing>>,
    #[serde(rename = "drawingHF")]
    #[serde(default)]
    pub drawing_h_f: Option<Box<SmlCTDrawingHF>>,
    #[serde(rename = "picture")]
    #[serde(default)]
    pub picture: Option<Box<SmlCTSheetBackgroundPicture>>,
    #[serde(rename = "webPublishItems")]
    #[serde(default)]
    pub web_publish_items: Option<Box<SmlCTWebPublishItems>>,
    #[serde(rename = "extLst")]
    #[serde(default)]
    pub ext_lst: Option<Box<SmlCTExtensionList>>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct SmlCTChartsheetPr {
    #[serde(rename = "@published")]
    #[serde(default)]
    pub published: Option<bool>,
    #[serde(rename = "@codeName")]
    #[serde(default)]
    pub code_name: Option<String>,
    #[serde(rename = "tabColor")]
    #[serde(default)]
    pub tab_color: Option<Box<SmlCTColor>>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct SmlCTChartsheetViews {
    #[serde(rename = "sheetView")]
    #[serde(default)]
    pub sheet_view: Vec<Box<SmlCTChartsheetView>>,
    #[serde(rename = "extLst")]
    #[serde(default)]
    pub ext_lst: Option<Box<SmlCTExtensionList>>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct SmlCTChartsheetView {
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
    pub ext_lst: Option<Box<SmlCTExtensionList>>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct SmlCTChartsheetProtection {
    #[serde(rename = "@password")]
    #[serde(default)]
    pub password: Option<SmlSTUnsignedShortHex>,
    #[serde(rename = "@algorithmName")]
    #[serde(default)]
    pub algorithm_name: Option<SSTXstring>,
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
pub struct SmlCTCsPageSetup {
    #[serde(rename = "@paperSize")]
    #[serde(default)]
    pub paper_size: Option<u32>,
    #[serde(rename = "@paperHeight")]
    #[serde(default)]
    pub paper_height: Option<SSTPositiveUniversalMeasure>,
    #[serde(rename = "@paperWidth")]
    #[serde(default)]
    pub paper_width: Option<SSTPositiveUniversalMeasure>,
    #[serde(rename = "@firstPageNumber")]
    #[serde(default)]
    pub first_page_number: Option<u32>,
    #[serde(rename = "@orientation")]
    #[serde(default)]
    pub orientation: Option<SmlSTOrientation>,
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
pub struct SmlCTCustomChartsheetViews {
    #[serde(rename = "customSheetView")]
    #[serde(default)]
    pub custom_sheet_view: Vec<Box<SmlCTCustomChartsheetView>>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct SmlCTCustomChartsheetView {
    #[serde(rename = "@guid")]
    pub guid: SSTGuid,
    #[serde(rename = "@scale")]
    #[serde(default)]
    pub scale: Option<u32>,
    #[serde(rename = "@state")]
    #[serde(default)]
    pub state: Option<SmlSTSheetState>,
    #[serde(rename = "@zoomToFit")]
    #[serde(default)]
    pub zoom_to_fit: Option<bool>,
    #[serde(rename = "pageMargins")]
    #[serde(default)]
    pub page_margins: Option<Box<SmlCTPageMargins>>,
    #[serde(rename = "pageSetup")]
    #[serde(default)]
    pub page_setup: Option<Box<SmlCTCsPageSetup>>,
    #[serde(rename = "headerFooter")]
    #[serde(default)]
    pub header_footer: Option<Box<SmlCTHeaderFooter>>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct SmlCTCustomProperties {
    #[serde(rename = "customPr")]
    #[serde(default)]
    pub custom_pr: Vec<Box<SmlCTCustomProperty>>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct SmlCTCustomProperty {
    #[serde(rename = "@name")]
    pub name: SSTXstring,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct SmlCTOleObjects {
    #[serde(rename = "oleObject")]
    #[serde(default)]
    pub ole_object: Vec<Box<SmlCTOleObject>>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct SmlCTOleObject {
    #[serde(rename = "@progId")]
    #[serde(default)]
    pub prog_id: Option<String>,
    #[serde(rename = "@dvAspect")]
    #[serde(default)]
    pub dv_aspect: Option<SmlSTDvAspect>,
    #[serde(rename = "@link")]
    #[serde(default)]
    pub link: Option<SSTXstring>,
    #[serde(rename = "@oleUpdate")]
    #[serde(default)]
    pub ole_update: Option<SmlSTOleUpdate>,
    #[serde(rename = "@autoLoad")]
    #[serde(default)]
    pub auto_load: Option<bool>,
    #[serde(rename = "@shapeId")]
    pub shape_id: u32,
    #[serde(rename = "objectPr")]
    #[serde(default)]
    pub object_pr: Option<Box<SmlCTObjectPr>>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct SmlCTObjectPr {
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
    pub r#macro: Option<SmlSTFormula>,
    #[serde(rename = "@altText")]
    #[serde(default)]
    pub alt_text: Option<SSTXstring>,
    #[serde(rename = "@dde")]
    #[serde(default)]
    pub dde: Option<bool>,
    #[serde(rename = "anchor")]
    pub anchor: Box<SmlCTObjectAnchor>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct SmlCTWebPublishItems {
    #[serde(rename = "@count")]
    #[serde(default)]
    pub count: Option<u32>,
    #[serde(rename = "webPublishItem")]
    #[serde(default)]
    pub web_publish_item: Vec<Box<SmlCTWebPublishItem>>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct SmlCTWebPublishItem {
    #[serde(rename = "@id")]
    pub id: u32,
    #[serde(rename = "@divId")]
    pub div_id: SSTXstring,
    #[serde(rename = "@sourceType")]
    pub source_type: SmlSTWebSourceType,
    #[serde(rename = "@sourceRef")]
    #[serde(default)]
    pub source_ref: Option<SmlSTRef>,
    #[serde(rename = "@sourceObject")]
    #[serde(default)]
    pub source_object: Option<SSTXstring>,
    #[serde(rename = "@destinationFile")]
    pub destination_file: SSTXstring,
    #[serde(rename = "@title")]
    #[serde(default)]
    pub title: Option<SSTXstring>,
    #[serde(rename = "@autoRepublish")]
    #[serde(default)]
    pub auto_republish: Option<bool>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct SmlCTControls {
    #[serde(rename = "control")]
    #[serde(default)]
    pub control: Vec<Box<SmlCTControl>>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct SmlCTControl {
    #[serde(rename = "@shapeId")]
    pub shape_id: u32,
    #[serde(rename = "@name")]
    #[serde(default)]
    pub name: Option<String>,
    #[serde(rename = "controlPr")]
    #[serde(default)]
    pub control_pr: Option<Box<SmlCTControlPr>>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct SmlCTControlPr {
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
    pub r#macro: Option<SmlSTFormula>,
    #[serde(rename = "@altText")]
    #[serde(default)]
    pub alt_text: Option<SSTXstring>,
    #[serde(rename = "@linkedCell")]
    #[serde(default)]
    pub linked_cell: Option<SmlSTFormula>,
    #[serde(rename = "@listFillRange")]
    #[serde(default)]
    pub list_fill_range: Option<SmlSTFormula>,
    #[serde(rename = "@cf")]
    #[serde(default)]
    pub cf: Option<SSTXstring>,
    #[serde(rename = "anchor")]
    pub anchor: Box<SmlCTObjectAnchor>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct SmlCTIgnoredErrors {
    #[serde(rename = "ignoredError")]
    #[serde(default)]
    pub ignored_error: Vec<Box<SmlCTIgnoredError>>,
    #[serde(rename = "extLst")]
    #[serde(default)]
    pub ext_lst: Option<Box<SmlCTExtensionList>>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct SmlCTIgnoredError {
    #[serde(rename = "@sqref")]
    pub sqref: SmlSTSqref,
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
pub struct SmlCTTableParts {
    #[serde(rename = "@count")]
    #[serde(default)]
    pub count: Option<u32>,
    #[serde(rename = "tablePart")]
    #[serde(default)]
    pub table_part: Vec<Box<SmlCTTablePart>>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct SmlCTTablePart;

pub type SmlMetadata = Box<SmlCTMetadata>;

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct SmlCTMetadata {
    #[serde(rename = "metadataTypes")]
    #[serde(default)]
    pub metadata_types: Option<Box<SmlCTMetadataTypes>>,
    #[serde(rename = "metadataStrings")]
    #[serde(default)]
    pub metadata_strings: Option<Box<SmlCTMetadataStrings>>,
    #[serde(rename = "mdxMetadata")]
    #[serde(default)]
    pub mdx_metadata: Option<Box<SmlCTMdxMetadata>>,
    #[serde(rename = "futureMetadata")]
    #[serde(default)]
    pub future_metadata: Vec<Box<SmlCTFutureMetadata>>,
    #[serde(rename = "cellMetadata")]
    #[serde(default)]
    pub cell_metadata: Option<Box<SmlCTMetadataBlocks>>,
    #[serde(rename = "valueMetadata")]
    #[serde(default)]
    pub value_metadata: Option<Box<SmlCTMetadataBlocks>>,
    #[serde(rename = "extLst")]
    #[serde(default)]
    pub ext_lst: Option<Box<SmlCTExtensionList>>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct SmlCTMetadataTypes {
    #[serde(rename = "@count")]
    #[serde(default)]
    pub count: Option<u32>,
    #[serde(rename = "metadataType")]
    #[serde(default)]
    pub metadata_type: Vec<Box<SmlCTMetadataType>>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct SmlCTMetadataType {
    #[serde(rename = "@name")]
    pub name: SSTXstring,
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
pub struct SmlCTMetadataBlocks {
    #[serde(rename = "@count")]
    #[serde(default)]
    pub count: Option<u32>,
    #[serde(rename = "bk")]
    #[serde(default)]
    pub bk: Vec<Box<SmlCTMetadataBlock>>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct SmlCTMetadataBlock {
    #[serde(rename = "rc")]
    #[serde(default)]
    pub rc: Vec<Box<SmlCTMetadataRecord>>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct SmlCTMetadataRecord {
    #[serde(rename = "@t")]
    pub t: u32,
    #[serde(rename = "@v")]
    pub v: u32,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct SmlCTFutureMetadata {
    #[serde(rename = "@name")]
    pub name: SSTXstring,
    #[serde(rename = "@count")]
    #[serde(default)]
    pub count: Option<u32>,
    #[serde(rename = "bk")]
    #[serde(default)]
    pub bk: Vec<Box<SmlCTFutureMetadataBlock>>,
    #[serde(rename = "extLst")]
    #[serde(default)]
    pub ext_lst: Option<Box<SmlCTExtensionList>>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct SmlCTFutureMetadataBlock {
    #[serde(rename = "extLst")]
    #[serde(default)]
    pub ext_lst: Option<Box<SmlCTExtensionList>>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct SmlCTMdxMetadata {
    #[serde(rename = "@count")]
    #[serde(default)]
    pub count: Option<u32>,
    #[serde(rename = "mdx")]
    #[serde(default)]
    pub mdx: Vec<Box<SmlCTMdx>>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct SmlCTMdx {
    #[serde(rename = "@n")]
    pub n: u32,
    #[serde(rename = "@f")]
    pub f: SmlSTMdxFunctionType,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct SmlCTMdxTuple {
    #[serde(rename = "@c")]
    #[serde(default)]
    pub c: Option<u32>,
    #[serde(rename = "@ct")]
    #[serde(default)]
    pub ct: Option<SSTXstring>,
    #[serde(rename = "@si")]
    #[serde(default)]
    pub si: Option<u32>,
    #[serde(rename = "@fi")]
    #[serde(default)]
    pub fi: Option<u32>,
    #[serde(rename = "@bc")]
    #[serde(default)]
    pub bc: Option<SmlSTUnsignedIntHex>,
    #[serde(rename = "@fc")]
    #[serde(default)]
    pub fc: Option<SmlSTUnsignedIntHex>,
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
    pub n: Vec<Box<SmlCTMetadataStringIndex>>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct SmlCTMdxSet {
    #[serde(rename = "@ns")]
    pub ns: u32,
    #[serde(rename = "@c")]
    #[serde(default)]
    pub c: Option<u32>,
    #[serde(rename = "@o")]
    #[serde(default)]
    pub o: Option<SmlSTMdxSetOrder>,
    #[serde(rename = "n")]
    #[serde(default)]
    pub n: Vec<Box<SmlCTMetadataStringIndex>>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct SmlCTMdxMemeberProp {
    #[serde(rename = "@n")]
    pub n: u32,
    #[serde(rename = "@np")]
    pub np: u32,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct SmlCTMdxKPI {
    #[serde(rename = "@n")]
    pub n: u32,
    #[serde(rename = "@np")]
    pub np: u32,
    #[serde(rename = "@p")]
    pub p: SmlSTMdxKPIProperty,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct SmlCTMetadataStringIndex {
    #[serde(rename = "@x")]
    pub x: u32,
    #[serde(rename = "@s")]
    #[serde(default)]
    pub s: Option<bool>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct SmlCTMetadataStrings {
    #[serde(rename = "@count")]
    #[serde(default)]
    pub count: Option<u32>,
    #[serde(rename = "s")]
    #[serde(default)]
    pub s: Vec<Box<SmlCTXStringElement>>,
}

pub type SmlSingleXmlCells = Box<SmlCTSingleXmlCells>;

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct SmlCTSingleXmlCells {
    #[serde(rename = "singleXmlCell")]
    #[serde(default)]
    pub single_xml_cell: Vec<Box<SmlCTSingleXmlCell>>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct SmlCTSingleXmlCell {
    #[serde(rename = "@id")]
    pub id: u32,
    #[serde(rename = "@r")]
    pub r: SmlSTCellRef,
    #[serde(rename = "@connectionId")]
    pub connection_id: u32,
    #[serde(rename = "xmlCellPr")]
    pub xml_cell_pr: Box<SmlCTXmlCellPr>,
    #[serde(rename = "extLst")]
    #[serde(default)]
    pub ext_lst: Option<Box<SmlCTExtensionList>>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct SmlCTXmlCellPr {
    #[serde(rename = "@id")]
    pub id: u32,
    #[serde(rename = "@uniqueName")]
    #[serde(default)]
    pub unique_name: Option<SSTXstring>,
    #[serde(rename = "xmlPr")]
    pub xml_pr: Box<SmlCTXmlPr>,
    #[serde(rename = "extLst")]
    #[serde(default)]
    pub ext_lst: Option<Box<SmlCTExtensionList>>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct SmlCTXmlPr {
    #[serde(rename = "@mapId")]
    pub map_id: u32,
    #[serde(rename = "@xpath")]
    pub xpath: SSTXstring,
    #[serde(rename = "@xmlDataType")]
    pub xml_data_type: SmlSTXmlDataType,
    #[serde(rename = "extLst")]
    #[serde(default)]
    pub ext_lst: Option<Box<SmlCTExtensionList>>,
}

pub type SmlStyleSheet = Box<SmlCTStylesheet>;

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct SmlCTStylesheet {
    #[serde(rename = "numFmts")]
    #[serde(default)]
    pub num_fmts: Option<Box<SmlCTNumFmts>>,
    #[serde(rename = "fonts")]
    #[serde(default)]
    pub fonts: Option<Box<SmlCTFonts>>,
    #[serde(rename = "fills")]
    #[serde(default)]
    pub fills: Option<Box<SmlCTFills>>,
    #[serde(rename = "borders")]
    #[serde(default)]
    pub borders: Option<Box<SmlCTBorders>>,
    #[serde(rename = "cellStyleXfs")]
    #[serde(default)]
    pub cell_style_xfs: Option<Box<SmlCTCellStyleXfs>>,
    #[serde(rename = "cellXfs")]
    #[serde(default)]
    pub cell_xfs: Option<Box<SmlCTCellXfs>>,
    #[serde(rename = "cellStyles")]
    #[serde(default)]
    pub cell_styles: Option<Box<SmlCTCellStyles>>,
    #[serde(rename = "dxfs")]
    #[serde(default)]
    pub dxfs: Option<Box<SmlCTDxfs>>,
    #[serde(rename = "tableStyles")]
    #[serde(default)]
    pub table_styles: Option<Box<SmlCTTableStyles>>,
    #[serde(rename = "colors")]
    #[serde(default)]
    pub colors: Option<Box<SmlCTColors>>,
    #[serde(rename = "extLst")]
    #[serde(default)]
    pub ext_lst: Option<Box<SmlCTExtensionList>>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct SmlCTCellAlignment {
    #[serde(rename = "@horizontal")]
    #[serde(default)]
    pub horizontal: Option<SmlSTHorizontalAlignment>,
    #[serde(rename = "@vertical")]
    #[serde(default)]
    pub vertical: Option<SmlSTVerticalAlignment>,
    #[serde(rename = "@textRotation")]
    #[serde(default)]
    pub text_rotation: Option<SmlSTTextRotation>,
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
pub struct SmlCTBorders {
    #[serde(rename = "@count")]
    #[serde(default)]
    pub count: Option<u32>,
    #[serde(rename = "border")]
    #[serde(default)]
    pub border: Vec<Box<SmlCTBorder>>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct SmlCTBorder {
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
    pub start: Option<Box<SmlCTBorderPr>>,
    #[serde(rename = "end")]
    #[serde(default)]
    pub end: Option<Box<SmlCTBorderPr>>,
    #[serde(rename = "left")]
    #[serde(default)]
    pub left: Option<Box<SmlCTBorderPr>>,
    #[serde(rename = "right")]
    #[serde(default)]
    pub right: Option<Box<SmlCTBorderPr>>,
    #[serde(rename = "top")]
    #[serde(default)]
    pub top: Option<Box<SmlCTBorderPr>>,
    #[serde(rename = "bottom")]
    #[serde(default)]
    pub bottom: Option<Box<SmlCTBorderPr>>,
    #[serde(rename = "diagonal")]
    #[serde(default)]
    pub diagonal: Option<Box<SmlCTBorderPr>>,
    #[serde(rename = "vertical")]
    #[serde(default)]
    pub vertical: Option<Box<SmlCTBorderPr>>,
    #[serde(rename = "horizontal")]
    #[serde(default)]
    pub horizontal: Option<Box<SmlCTBorderPr>>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct SmlCTBorderPr {
    #[serde(rename = "@style")]
    #[serde(default)]
    pub style: Option<SmlSTBorderStyle>,
    #[serde(rename = "color")]
    #[serde(default)]
    pub color: Option<Box<SmlCTColor>>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct SmlCTCellProtection {
    #[serde(rename = "@locked")]
    #[serde(default)]
    pub locked: Option<bool>,
    #[serde(rename = "@hidden")]
    #[serde(default)]
    pub hidden: Option<bool>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct SmlCTFonts {
    #[serde(rename = "@count")]
    #[serde(default)]
    pub count: Option<u32>,
    #[serde(rename = "font")]
    #[serde(default)]
    pub font: Vec<Box<SmlCTFont>>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct SmlCTFills {
    #[serde(rename = "@count")]
    #[serde(default)]
    pub count: Option<u32>,
    #[serde(rename = "fill")]
    #[serde(default)]
    pub fill: Vec<Box<SmlCTFill>>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct SmlCTFill;

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct SmlCTPatternFill {
    #[serde(rename = "@patternType")]
    #[serde(default)]
    pub pattern_type: Option<SmlSTPatternType>,
    #[serde(rename = "fgColor")]
    #[serde(default)]
    pub fg_color: Option<Box<SmlCTColor>>,
    #[serde(rename = "bgColor")]
    #[serde(default)]
    pub bg_color: Option<Box<SmlCTColor>>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct SmlCTColor {
    #[serde(rename = "@auto")]
    #[serde(default)]
    pub auto: Option<bool>,
    #[serde(rename = "@indexed")]
    #[serde(default)]
    pub indexed: Option<u32>,
    #[serde(rename = "@rgb")]
    #[serde(default)]
    pub rgb: Option<SmlSTUnsignedIntHex>,
    #[serde(rename = "@theme")]
    #[serde(default)]
    pub theme: Option<u32>,
    #[serde(rename = "@tint")]
    #[serde(default)]
    pub tint: Option<f64>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct SmlCTGradientFill {
    #[serde(rename = "@type")]
    #[serde(default)]
    pub r#type: Option<SmlSTGradientType>,
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
    pub stop: Vec<Box<SmlCTGradientStop>>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct SmlCTGradientStop {
    #[serde(rename = "@position")]
    pub position: f64,
    #[serde(rename = "color")]
    pub color: Box<SmlCTColor>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct SmlCTNumFmts {
    #[serde(rename = "@count")]
    #[serde(default)]
    pub count: Option<u32>,
    #[serde(rename = "numFmt")]
    #[serde(default)]
    pub num_fmt: Vec<Box<SmlCTNumFmt>>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct SmlCTNumFmt {
    #[serde(rename = "@numFmtId")]
    pub num_fmt_id: SmlSTNumFmtId,
    #[serde(rename = "@formatCode")]
    pub format_code: SSTXstring,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct SmlCTCellStyleXfs {
    #[serde(rename = "@count")]
    #[serde(default)]
    pub count: Option<u32>,
    #[serde(rename = "xf")]
    #[serde(default)]
    pub xf: Vec<Box<SmlCTXf>>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct SmlCTCellXfs {
    #[serde(rename = "@count")]
    #[serde(default)]
    pub count: Option<u32>,
    #[serde(rename = "xf")]
    #[serde(default)]
    pub xf: Vec<Box<SmlCTXf>>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct SmlCTXf {
    #[serde(rename = "@numFmtId")]
    #[serde(default)]
    pub num_fmt_id: Option<SmlSTNumFmtId>,
    #[serde(rename = "@fontId")]
    #[serde(default)]
    pub font_id: Option<SmlSTFontId>,
    #[serde(rename = "@fillId")]
    #[serde(default)]
    pub fill_id: Option<SmlSTFillId>,
    #[serde(rename = "@borderId")]
    #[serde(default)]
    pub border_id: Option<SmlSTBorderId>,
    #[serde(rename = "@xfId")]
    #[serde(default)]
    pub xf_id: Option<SmlSTCellStyleXfId>,
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
    pub alignment: Option<Box<SmlCTCellAlignment>>,
    #[serde(rename = "protection")]
    #[serde(default)]
    pub protection: Option<Box<SmlCTCellProtection>>,
    #[serde(rename = "extLst")]
    #[serde(default)]
    pub ext_lst: Option<Box<SmlCTExtensionList>>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct SmlCTCellStyles {
    #[serde(rename = "@count")]
    #[serde(default)]
    pub count: Option<u32>,
    #[serde(rename = "cellStyle")]
    #[serde(default)]
    pub cell_style: Vec<Box<SmlCTCellStyle>>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct SmlCTCellStyle {
    #[serde(rename = "@name")]
    #[serde(default)]
    pub name: Option<SSTXstring>,
    #[serde(rename = "@xfId")]
    pub xf_id: SmlSTCellStyleXfId,
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
    pub ext_lst: Option<Box<SmlCTExtensionList>>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct SmlCTDxfs {
    #[serde(rename = "@count")]
    #[serde(default)]
    pub count: Option<u32>,
    #[serde(rename = "dxf")]
    #[serde(default)]
    pub dxf: Vec<Box<SmlCTDxf>>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct SmlCTDxf {
    #[serde(rename = "font")]
    #[serde(default)]
    pub font: Option<Box<SmlCTFont>>,
    #[serde(rename = "numFmt")]
    #[serde(default)]
    pub num_fmt: Option<Box<SmlCTNumFmt>>,
    #[serde(rename = "fill")]
    #[serde(default)]
    pub fill: Option<Box<SmlCTFill>>,
    #[serde(rename = "alignment")]
    #[serde(default)]
    pub alignment: Option<Box<SmlCTCellAlignment>>,
    #[serde(rename = "border")]
    #[serde(default)]
    pub border: Option<Box<SmlCTBorder>>,
    #[serde(rename = "protection")]
    #[serde(default)]
    pub protection: Option<Box<SmlCTCellProtection>>,
    #[serde(rename = "extLst")]
    #[serde(default)]
    pub ext_lst: Option<Box<SmlCTExtensionList>>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct SmlCTColors {
    #[serde(rename = "indexedColors")]
    #[serde(default)]
    pub indexed_colors: Option<Box<SmlCTIndexedColors>>,
    #[serde(rename = "mruColors")]
    #[serde(default)]
    pub mru_colors: Option<Box<SmlCTMRUColors>>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct SmlCTIndexedColors {
    #[serde(rename = "rgbColor")]
    #[serde(default)]
    pub rgb_color: Vec<Box<SmlCTRgbColor>>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct SmlCTMRUColors {
    #[serde(rename = "color")]
    #[serde(default)]
    pub color: Vec<Box<SmlCTColor>>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct SmlCTRgbColor {
    #[serde(rename = "@rgb")]
    #[serde(default)]
    pub rgb: Option<SmlSTUnsignedIntHex>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct SmlCTTableStyles {
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
    pub table_style: Vec<Box<SmlCTTableStyle>>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct SmlCTTableStyle {
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
    pub table_style_element: Vec<Box<SmlCTTableStyleElement>>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct SmlCTTableStyleElement {
    #[serde(rename = "@type")]
    pub r#type: SmlSTTableStyleType,
    #[serde(rename = "@size")]
    #[serde(default)]
    pub size: Option<u32>,
    #[serde(rename = "@dxfId")]
    #[serde(default)]
    pub dxf_id: Option<SmlSTDxfId>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct SmlCTBooleanProperty {
    #[serde(rename = "@val")]
    #[serde(default)]
    pub val: Option<bool>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct SmlCTFontSize {
    #[serde(rename = "@val")]
    pub val: f64,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct SmlCTIntProperty {
    #[serde(rename = "@val")]
    pub val: i32,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct SmlCTFontName {
    #[serde(rename = "@val")]
    pub val: SSTXstring,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct SmlCTVerticalAlignFontProperty {
    #[serde(rename = "@val")]
    pub val: SSTVerticalAlignRun,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct SmlCTFontScheme {
    #[serde(rename = "@val")]
    pub val: SmlSTFontScheme,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct SmlCTUnderlineProperty {
    #[serde(rename = "@val")]
    #[serde(default)]
    pub val: Option<SmlSTUnderlineValues>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct SmlCTFont;

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct SmlCTFontFamily {
    #[serde(rename = "@val")]
    pub val: SmlSTFontFamily,
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

pub type SmlExternalLink = Box<SmlCTExternalLink>;

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct SmlCTExternalLink {
    #[serde(rename = "extLst")]
    #[serde(default)]
    pub ext_lst: Option<Box<SmlCTExtensionList>>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct SmlCTExternalBook {
    #[serde(rename = "sheetNames")]
    #[serde(default)]
    pub sheet_names: Option<Box<SmlCTExternalSheetNames>>,
    #[serde(rename = "definedNames")]
    #[serde(default)]
    pub defined_names: Option<Box<SmlCTExternalDefinedNames>>,
    #[serde(rename = "sheetDataSet")]
    #[serde(default)]
    pub sheet_data_set: Option<Box<SmlCTExternalSheetDataSet>>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct SmlCTExternalSheetNames {
    #[serde(rename = "sheetName")]
    #[serde(default)]
    pub sheet_name: Vec<Box<SmlCTExternalSheetName>>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct SmlCTExternalSheetName {
    #[serde(rename = "@val")]
    #[serde(default)]
    pub val: Option<SSTXstring>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct SmlCTExternalDefinedNames {
    #[serde(rename = "definedName")]
    #[serde(default)]
    pub defined_name: Vec<Box<SmlCTExternalDefinedName>>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct SmlCTExternalDefinedName {
    #[serde(rename = "@name")]
    pub name: SSTXstring,
    #[serde(rename = "@refersTo")]
    #[serde(default)]
    pub refers_to: Option<SSTXstring>,
    #[serde(rename = "@sheetId")]
    #[serde(default)]
    pub sheet_id: Option<u32>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct SmlCTExternalSheetDataSet {
    #[serde(rename = "sheetData")]
    #[serde(default)]
    pub sheet_data: Vec<Box<SmlCTExternalSheetData>>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct SmlCTExternalSheetData {
    #[serde(rename = "@sheetId")]
    pub sheet_id: u32,
    #[serde(rename = "@refreshError")]
    #[serde(default)]
    pub refresh_error: Option<bool>,
    #[serde(rename = "row")]
    #[serde(default)]
    pub row: Vec<Box<SmlCTExternalRow>>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct SmlCTExternalRow {
    #[serde(rename = "@r")]
    pub r: u32,
    #[serde(rename = "cell")]
    #[serde(default)]
    pub cell: Vec<Box<SmlCTExternalCell>>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct SmlCTExternalCell {
    #[serde(rename = "@r")]
    #[serde(default)]
    pub r: Option<SmlSTCellRef>,
    #[serde(rename = "@t")]
    #[serde(default)]
    pub t: Option<SmlSTCellType>,
    #[serde(rename = "@vm")]
    #[serde(default)]
    pub vm: Option<u32>,
    #[serde(rename = "v")]
    #[serde(default)]
    pub v: Option<SSTXstring>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct SmlCTDdeLink {
    #[serde(rename = "@ddeService")]
    pub dde_service: SSTXstring,
    #[serde(rename = "@ddeTopic")]
    pub dde_topic: SSTXstring,
    #[serde(rename = "ddeItems")]
    #[serde(default)]
    pub dde_items: Option<Box<SmlCTDdeItems>>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct SmlCTDdeItems {
    #[serde(rename = "ddeItem")]
    #[serde(default)]
    pub dde_item: Vec<Box<SmlCTDdeItem>>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct SmlCTDdeItem {
    #[serde(rename = "@name")]
    #[serde(default)]
    pub name: Option<SSTXstring>,
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
    pub values: Option<Box<SmlCTDdeValues>>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct SmlCTDdeValues {
    #[serde(rename = "@rows")]
    #[serde(default)]
    pub rows: Option<u32>,
    #[serde(rename = "@cols")]
    #[serde(default)]
    pub cols: Option<u32>,
    #[serde(rename = "value")]
    #[serde(default)]
    pub value: Vec<Box<SmlCTDdeValue>>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct SmlCTDdeValue {
    #[serde(rename = "@t")]
    #[serde(default)]
    pub t: Option<SmlSTDdeValueType>,
    #[serde(rename = "val")]
    pub val: SSTXstring,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct SmlCTOleLink {
    #[serde(rename = "@progId")]
    pub prog_id: SSTXstring,
    #[serde(rename = "oleItems")]
    #[serde(default)]
    pub ole_items: Option<Box<SmlCTOleItems>>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct SmlCTOleItems {
    #[serde(rename = "oleItem")]
    #[serde(default)]
    pub ole_item: Vec<Box<SmlCTOleItem>>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct SmlCTOleItem {
    #[serde(rename = "@name")]
    pub name: SSTXstring,
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

pub type SmlTable = Box<SmlCTTable>;

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct SmlCTTable {
    #[serde(rename = "@id")]
    pub id: u32,
    #[serde(rename = "@name")]
    #[serde(default)]
    pub name: Option<SSTXstring>,
    #[serde(rename = "@displayName")]
    pub display_name: SSTXstring,
    #[serde(rename = "@comment")]
    #[serde(default)]
    pub comment: Option<SSTXstring>,
    #[serde(rename = "@ref")]
    pub r#ref: SmlSTRef,
    #[serde(rename = "@tableType")]
    #[serde(default)]
    pub table_type: Option<SmlSTTableType>,
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
    pub header_row_dxf_id: Option<SmlSTDxfId>,
    #[serde(rename = "@dataDxfId")]
    #[serde(default)]
    pub data_dxf_id: Option<SmlSTDxfId>,
    #[serde(rename = "@totalsRowDxfId")]
    #[serde(default)]
    pub totals_row_dxf_id: Option<SmlSTDxfId>,
    #[serde(rename = "@headerRowBorderDxfId")]
    #[serde(default)]
    pub header_row_border_dxf_id: Option<SmlSTDxfId>,
    #[serde(rename = "@tableBorderDxfId")]
    #[serde(default)]
    pub table_border_dxf_id: Option<SmlSTDxfId>,
    #[serde(rename = "@totalsRowBorderDxfId")]
    #[serde(default)]
    pub totals_row_border_dxf_id: Option<SmlSTDxfId>,
    #[serde(rename = "@headerRowCellStyle")]
    #[serde(default)]
    pub header_row_cell_style: Option<SSTXstring>,
    #[serde(rename = "@dataCellStyle")]
    #[serde(default)]
    pub data_cell_style: Option<SSTXstring>,
    #[serde(rename = "@totalsRowCellStyle")]
    #[serde(default)]
    pub totals_row_cell_style: Option<SSTXstring>,
    #[serde(rename = "@connectionId")]
    #[serde(default)]
    pub connection_id: Option<u32>,
    #[serde(rename = "autoFilter")]
    #[serde(default)]
    pub auto_filter: Option<Box<SmlCTAutoFilter>>,
    #[serde(rename = "sortState")]
    #[serde(default)]
    pub sort_state: Option<Box<SmlCTSortState>>,
    #[serde(rename = "tableColumns")]
    pub table_columns: Box<SmlCTTableColumns>,
    #[serde(rename = "tableStyleInfo")]
    #[serde(default)]
    pub table_style_info: Option<Box<SmlCTTableStyleInfo>>,
    #[serde(rename = "extLst")]
    #[serde(default)]
    pub ext_lst: Option<Box<SmlCTExtensionList>>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct SmlCTTableStyleInfo {
    #[serde(rename = "@name")]
    #[serde(default)]
    pub name: Option<SSTXstring>,
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
pub struct SmlCTTableColumns {
    #[serde(rename = "@count")]
    #[serde(default)]
    pub count: Option<u32>,
    #[serde(rename = "tableColumn")]
    #[serde(default)]
    pub table_column: Vec<Box<SmlCTTableColumn>>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct SmlCTTableColumn {
    #[serde(rename = "@id")]
    pub id: u32,
    #[serde(rename = "@uniqueName")]
    #[serde(default)]
    pub unique_name: Option<SSTXstring>,
    #[serde(rename = "@name")]
    pub name: SSTXstring,
    #[serde(rename = "@totalsRowFunction")]
    #[serde(default)]
    pub totals_row_function: Option<SmlSTTotalsRowFunction>,
    #[serde(rename = "@totalsRowLabel")]
    #[serde(default)]
    pub totals_row_label: Option<SSTXstring>,
    #[serde(rename = "@queryTableFieldId")]
    #[serde(default)]
    pub query_table_field_id: Option<u32>,
    #[serde(rename = "@headerRowDxfId")]
    #[serde(default)]
    pub header_row_dxf_id: Option<SmlSTDxfId>,
    #[serde(rename = "@dataDxfId")]
    #[serde(default)]
    pub data_dxf_id: Option<SmlSTDxfId>,
    #[serde(rename = "@totalsRowDxfId")]
    #[serde(default)]
    pub totals_row_dxf_id: Option<SmlSTDxfId>,
    #[serde(rename = "@headerRowCellStyle")]
    #[serde(default)]
    pub header_row_cell_style: Option<SSTXstring>,
    #[serde(rename = "@dataCellStyle")]
    #[serde(default)]
    pub data_cell_style: Option<SSTXstring>,
    #[serde(rename = "@totalsRowCellStyle")]
    #[serde(default)]
    pub totals_row_cell_style: Option<SSTXstring>,
    #[serde(rename = "calculatedColumnFormula")]
    #[serde(default)]
    pub calculated_column_formula: Option<Box<SmlCTTableFormula>>,
    #[serde(rename = "totalsRowFormula")]
    #[serde(default)]
    pub totals_row_formula: Option<Box<SmlCTTableFormula>>,
    #[serde(rename = "xmlColumnPr")]
    #[serde(default)]
    pub xml_column_pr: Option<Box<SmlCTXmlColumnPr>>,
    #[serde(rename = "extLst")]
    #[serde(default)]
    pub ext_lst: Option<Box<SmlCTExtensionList>>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct SmlCTTableFormula {
    #[serde(rename = "@array")]
    #[serde(default)]
    pub array: Option<bool>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct SmlCTXmlColumnPr {
    #[serde(rename = "@mapId")]
    pub map_id: u32,
    #[serde(rename = "@xpath")]
    pub xpath: SSTXstring,
    #[serde(rename = "@denormalized")]
    #[serde(default)]
    pub denormalized: Option<bool>,
    #[serde(rename = "@xmlDataType")]
    pub xml_data_type: SmlSTXmlDataType,
    #[serde(rename = "extLst")]
    #[serde(default)]
    pub ext_lst: Option<Box<SmlCTExtensionList>>,
}

pub type SmlVolTypes = Box<SmlCTVolTypes>;

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct SmlCTVolTypes {
    #[serde(rename = "volType")]
    #[serde(default)]
    pub vol_type: Vec<Box<SmlCTVolType>>,
    #[serde(rename = "extLst")]
    #[serde(default)]
    pub ext_lst: Option<Box<SmlCTExtensionList>>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct SmlCTVolType {
    #[serde(rename = "@type")]
    pub r#type: SmlSTVolDepType,
    #[serde(rename = "main")]
    #[serde(default)]
    pub main: Vec<Box<SmlCTVolMain>>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct SmlCTVolMain {
    #[serde(rename = "@first")]
    pub first: SSTXstring,
    #[serde(rename = "tp")]
    #[serde(default)]
    pub tp: Vec<Box<SmlCTVolTopic>>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct SmlCTVolTopic {
    #[serde(rename = "@t")]
    #[serde(default)]
    pub t: Option<SmlSTVolValueType>,
    #[serde(rename = "v")]
    pub v: SSTXstring,
    #[serde(rename = "stp")]
    #[serde(default)]
    pub stp: Vec<SSTXstring>,
    #[serde(rename = "tr")]
    #[serde(default)]
    pub tr: Vec<Box<SmlCTVolTopicRef>>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct SmlCTVolTopicRef {
    #[serde(rename = "@r")]
    pub r: SmlSTCellRef,
    #[serde(rename = "@s")]
    pub s: u32,
}

pub type SmlWorkbook = Box<SmlCTWorkbook>;

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct SmlCTWorkbook {
    #[serde(rename = "@conformance")]
    #[serde(default)]
    pub conformance: Option<SSTConformanceClass>,
    #[serde(rename = "fileVersion")]
    #[serde(default)]
    pub file_version: Option<Box<SmlCTFileVersion>>,
    #[serde(rename = "fileSharing")]
    #[serde(default)]
    pub file_sharing: Option<Box<SmlCTFileSharing>>,
    #[serde(rename = "workbookPr")]
    #[serde(default)]
    pub workbook_pr: Option<Box<SmlCTWorkbookPr>>,
    #[serde(rename = "workbookProtection")]
    #[serde(default)]
    pub workbook_protection: Option<Box<SmlCTWorkbookProtection>>,
    #[serde(rename = "bookViews")]
    #[serde(default)]
    pub book_views: Option<Box<SmlCTBookViews>>,
    #[serde(rename = "sheets")]
    pub sheets: Box<SmlCTSheets>,
    #[serde(rename = "functionGroups")]
    #[serde(default)]
    pub function_groups: Option<Box<SmlCTFunctionGroups>>,
    #[serde(rename = "externalReferences")]
    #[serde(default)]
    pub external_references: Option<Box<SmlCTExternalReferences>>,
    #[serde(rename = "definedNames")]
    #[serde(default)]
    pub defined_names: Option<Box<SmlCTDefinedNames>>,
    #[serde(rename = "calcPr")]
    #[serde(default)]
    pub calc_pr: Option<Box<SmlCTCalcPr>>,
    #[serde(rename = "oleSize")]
    #[serde(default)]
    pub ole_size: Option<Box<SmlCTOleSize>>,
    #[serde(rename = "customWorkbookViews")]
    #[serde(default)]
    pub custom_workbook_views: Option<Box<SmlCTCustomWorkbookViews>>,
    #[serde(rename = "pivotCaches")]
    #[serde(default)]
    pub pivot_caches: Option<Box<SmlCTPivotCaches>>,
    #[serde(rename = "smartTagPr")]
    #[serde(default)]
    pub smart_tag_pr: Option<Box<SmlCTSmartTagPr>>,
    #[serde(rename = "smartTagTypes")]
    #[serde(default)]
    pub smart_tag_types: Option<Box<SmlCTSmartTagTypes>>,
    #[serde(rename = "webPublishing")]
    #[serde(default)]
    pub web_publishing: Option<Box<SmlCTWebPublishing>>,
    #[serde(rename = "fileRecoveryPr")]
    #[serde(default)]
    pub file_recovery_pr: Vec<Box<SmlCTFileRecoveryPr>>,
    #[serde(rename = "webPublishObjects")]
    #[serde(default)]
    pub web_publish_objects: Option<Box<SmlCTWebPublishObjects>>,
    #[serde(rename = "extLst")]
    #[serde(default)]
    pub ext_lst: Option<Box<SmlCTExtensionList>>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct SmlCTFileVersion {
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
    pub code_name: Option<SSTGuid>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct SmlCTBookViews {
    #[serde(rename = "workbookView")]
    #[serde(default)]
    pub workbook_view: Vec<Box<SmlCTBookView>>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct SmlCTBookView {
    #[serde(rename = "@visibility")]
    #[serde(default)]
    pub visibility: Option<SmlSTVisibility>,
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
    pub ext_lst: Option<Box<SmlCTExtensionList>>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct SmlCTCustomWorkbookViews {
    #[serde(rename = "customWorkbookView")]
    #[serde(default)]
    pub custom_workbook_view: Vec<Box<SmlCTCustomWorkbookView>>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct SmlCTCustomWorkbookView {
    #[serde(rename = "@name")]
    pub name: SSTXstring,
    #[serde(rename = "@guid")]
    pub guid: SSTGuid,
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
    pub show_comments: Option<SmlSTComments>,
    #[serde(rename = "@showObjects")]
    #[serde(default)]
    pub show_objects: Option<SmlSTObjects>,
    #[serde(rename = "extLst")]
    #[serde(default)]
    pub ext_lst: Option<Box<SmlCTExtensionList>>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct SmlCTSheets {
    #[serde(rename = "sheet")]
    #[serde(default)]
    pub sheet: Vec<Box<SmlCTSheet>>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct SmlCTSheet {
    #[serde(rename = "@name")]
    pub name: SSTXstring,
    #[serde(rename = "@sheetId")]
    pub sheet_id: u32,
    #[serde(rename = "@state")]
    #[serde(default)]
    pub state: Option<SmlSTSheetState>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct SmlCTWorkbookPr {
    #[serde(rename = "@date1904")]
    #[serde(default)]
    pub date1904: Option<bool>,
    #[serde(rename = "@showObjects")]
    #[serde(default)]
    pub show_objects: Option<SmlSTObjects>,
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
    pub update_links: Option<SmlSTUpdateLinks>,
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
pub struct SmlCTSmartTagPr {
    #[serde(rename = "@embed")]
    #[serde(default)]
    pub embed: Option<bool>,
    #[serde(rename = "@show")]
    #[serde(default)]
    pub show: Option<SmlSTSmartTagShow>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct SmlCTSmartTagTypes {
    #[serde(rename = "smartTagType")]
    #[serde(default)]
    pub smart_tag_type: Vec<Box<SmlCTSmartTagType>>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct SmlCTSmartTagType {
    #[serde(rename = "@namespaceUri")]
    #[serde(default)]
    pub namespace_uri: Option<SSTXstring>,
    #[serde(rename = "@name")]
    #[serde(default)]
    pub name: Option<SSTXstring>,
    #[serde(rename = "@url")]
    #[serde(default)]
    pub url: Option<SSTXstring>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct SmlCTFileRecoveryPr {
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
pub struct SmlCTCalcPr {
    #[serde(rename = "@calcId")]
    #[serde(default)]
    pub calc_id: Option<u32>,
    #[serde(rename = "@calcMode")]
    #[serde(default)]
    pub calc_mode: Option<SmlSTCalcMode>,
    #[serde(rename = "@fullCalcOnLoad")]
    #[serde(default)]
    pub full_calc_on_load: Option<bool>,
    #[serde(rename = "@refMode")]
    #[serde(default)]
    pub ref_mode: Option<SmlSTRefMode>,
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
pub struct SmlCTDefinedNames {
    #[serde(rename = "definedName")]
    #[serde(default)]
    pub defined_name: Vec<Box<SmlCTDefinedName>>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct SmlCTDefinedName {
    #[serde(rename = "@name")]
    pub name: SSTXstring,
    #[serde(rename = "@comment")]
    #[serde(default)]
    pub comment: Option<SSTXstring>,
    #[serde(rename = "@customMenu")]
    #[serde(default)]
    pub custom_menu: Option<SSTXstring>,
    #[serde(rename = "@description")]
    #[serde(default)]
    pub description: Option<SSTXstring>,
    #[serde(rename = "@help")]
    #[serde(default)]
    pub help: Option<SSTXstring>,
    #[serde(rename = "@statusBar")]
    #[serde(default)]
    pub status_bar: Option<SSTXstring>,
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
    pub shortcut_key: Option<SSTXstring>,
    #[serde(rename = "@publishToServer")]
    #[serde(default)]
    pub publish_to_server: Option<bool>,
    #[serde(rename = "@workbookParameter")]
    #[serde(default)]
    pub workbook_parameter: Option<bool>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct SmlCTExternalReferences {
    #[serde(rename = "externalReference")]
    #[serde(default)]
    pub external_reference: Vec<Box<SmlCTExternalReference>>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct SmlCTExternalReference;

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct SmlCTSheetBackgroundPicture;

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct SmlCTPivotCaches {
    #[serde(rename = "pivotCache")]
    #[serde(default)]
    pub pivot_cache: Vec<Box<SmlCTPivotCache>>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct SmlCTPivotCache {
    #[serde(rename = "@cacheId")]
    pub cache_id: u32,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct SmlCTFileSharing {
    #[serde(rename = "@readOnlyRecommended")]
    #[serde(default)]
    pub read_only_recommended: Option<bool>,
    #[serde(rename = "@userName")]
    #[serde(default)]
    pub user_name: Option<SSTXstring>,
    #[serde(rename = "@reservationPassword")]
    #[serde(default)]
    pub reservation_password: Option<SmlSTUnsignedShortHex>,
    #[serde(rename = "@algorithmName")]
    #[serde(default)]
    pub algorithm_name: Option<SSTXstring>,
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
pub struct SmlCTOleSize {
    #[serde(rename = "@ref")]
    pub r#ref: SmlSTRef,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct SmlCTWorkbookProtection {
    #[serde(rename = "@workbookPassword")]
    #[serde(default)]
    pub workbook_password: Option<SmlSTUnsignedShortHex>,
    #[serde(rename = "@workbookPasswordCharacterSet")]
    #[serde(default)]
    pub workbook_password_character_set: Option<String>,
    #[serde(rename = "@revisionsPassword")]
    #[serde(default)]
    pub revisions_password: Option<SmlSTUnsignedShortHex>,
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
    pub revisions_algorithm_name: Option<SSTXstring>,
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
    pub workbook_algorithm_name: Option<SSTXstring>,
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
pub struct SmlCTWebPublishing {
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
    pub target_screen_size: Option<SmlSTTargetScreenSize>,
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
pub struct SmlCTFunctionGroups {
    #[serde(rename = "@builtInGroupCount")]
    #[serde(default)]
    pub built_in_group_count: Option<u32>,
    #[serde(rename = "functionGroup")]
    #[serde(default)]
    pub function_group: Vec<Box<SmlCTFunctionGroup>>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct SmlCTFunctionGroup {
    #[serde(rename = "@name")]
    #[serde(default)]
    pub name: Option<SSTXstring>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct SmlCTWebPublishObjects {
    #[serde(rename = "@count")]
    #[serde(default)]
    pub count: Option<u32>,
    #[serde(rename = "webPublishObject")]
    #[serde(default)]
    pub web_publish_object: Vec<Box<SmlCTWebPublishObject>>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct SmlCTWebPublishObject {
    #[serde(rename = "@id")]
    pub id: u32,
    #[serde(rename = "@divId")]
    pub div_id: SSTXstring,
    #[serde(rename = "@sourceObject")]
    #[serde(default)]
    pub source_object: Option<SSTXstring>,
    #[serde(rename = "@destinationFile")]
    pub destination_file: SSTXstring,
    #[serde(rename = "@title")]
    #[serde(default)]
    pub title: Option<SSTXstring>,
    #[serde(rename = "@autoRepublish")]
    #[serde(default)]
    pub auto_republish: Option<bool>,
}
