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
    /// Namespace prefix: a
    pub const A: &str = "http://schemas.openxmlformats.org/drawingml/2006/main";
    /// Namespace prefix: p
    pub const P: &str = "http://schemas.openxmlformats.org/presentationml/2006/main";
    /// Namespace prefix: r
    pub const R: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
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
pub enum STVerticalAlignRun {
    #[serde(rename = "baseline")]
    Baseline,
    #[serde(rename = "superscript")]
    Superscript,
    #[serde(rename = "subscript")]
    Subscript,
}

impl std::fmt::Display for STVerticalAlignRun {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::Baseline => write!(f, "baseline"),
            Self::Superscript => write!(f, "superscript"),
            Self::Subscript => write!(f, "subscript"),
        }
    }
}

impl std::str::FromStr for STVerticalAlignRun {
    type Err = String;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        match s {
            "baseline" => Ok(Self::Baseline),
            "superscript" => Ok(Self::Superscript),
            "subscript" => Ok(Self::Subscript),
            _ => Err(format!("unknown STVerticalAlignRun value: {}", s)),
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
pub enum STTransitionSideDirectionType {
    #[serde(rename = "l")]
    L,
    #[serde(rename = "u")]
    U,
    #[serde(rename = "r")]
    R,
    #[serde(rename = "d")]
    D,
}

impl std::fmt::Display for STTransitionSideDirectionType {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::L => write!(f, "l"),
            Self::U => write!(f, "u"),
            Self::R => write!(f, "r"),
            Self::D => write!(f, "d"),
        }
    }
}

impl std::str::FromStr for STTransitionSideDirectionType {
    type Err = String;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        match s {
            "l" => Ok(Self::L),
            "u" => Ok(Self::U),
            "r" => Ok(Self::R),
            "d" => Ok(Self::D),
            _ => Err(format!(
                "unknown STTransitionSideDirectionType value: {}",
                s
            )),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum STTransitionCornerDirectionType {
    #[serde(rename = "lu")]
    Lu,
    #[serde(rename = "ru")]
    Ru,
    #[serde(rename = "ld")]
    Ld,
    #[serde(rename = "rd")]
    Rd,
}

impl std::fmt::Display for STTransitionCornerDirectionType {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::Lu => write!(f, "lu"),
            Self::Ru => write!(f, "ru"),
            Self::Ld => write!(f, "ld"),
            Self::Rd => write!(f, "rd"),
        }
    }
}

impl std::str::FromStr for STTransitionCornerDirectionType {
    type Err = String;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        match s {
            "lu" => Ok(Self::Lu),
            "ru" => Ok(Self::Ru),
            "ld" => Ok(Self::Ld),
            "rd" => Ok(Self::Rd),
            _ => Err(format!(
                "unknown STTransitionCornerDirectionType value: {}",
                s
            )),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum STTransitionInOutDirectionType {
    #[serde(rename = "out")]
    Out,
    #[serde(rename = "in")]
    In,
}

impl std::fmt::Display for STTransitionInOutDirectionType {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::Out => write!(f, "out"),
            Self::In => write!(f, "in"),
        }
    }
}

impl std::str::FromStr for STTransitionInOutDirectionType {
    type Err = String;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        match s {
            "out" => Ok(Self::Out),
            "in" => Ok(Self::In),
            _ => Err(format!(
                "unknown STTransitionInOutDirectionType value: {}",
                s
            )),
        }
    }
}

pub type STTransitionEightDirectionType = String;

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum STTransitionSpeed {
    #[serde(rename = "slow")]
    Slow,
    #[serde(rename = "med")]
    Med,
    #[serde(rename = "fast")]
    Fast,
}

impl std::fmt::Display for STTransitionSpeed {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::Slow => write!(f, "slow"),
            Self::Med => write!(f, "med"),
            Self::Fast => write!(f, "fast"),
        }
    }
}

impl std::str::FromStr for STTransitionSpeed {
    type Err = String;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        match s {
            "slow" => Ok(Self::Slow),
            "med" => Ok(Self::Med),
            "fast" => Ok(Self::Fast),
            _ => Err(format!("unknown STTransitionSpeed value: {}", s)),
        }
    }
}

pub type STTLTime = String;

pub type STTLTimeNodeID = u32;

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum STIterateType {
    #[serde(rename = "el")]
    El,
    #[serde(rename = "wd")]
    Wd,
    #[serde(rename = "lt")]
    Lt,
}

impl std::fmt::Display for STIterateType {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::El => write!(f, "el"),
            Self::Wd => write!(f, "wd"),
            Self::Lt => write!(f, "lt"),
        }
    }
}

impl std::str::FromStr for STIterateType {
    type Err = String;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        match s {
            "el" => Ok(Self::El),
            "wd" => Ok(Self::Wd),
            "lt" => Ok(Self::Lt),
            _ => Err(format!("unknown STIterateType value: {}", s)),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum STTLChartSubelementType {
    #[serde(rename = "gridLegend")]
    GridLegend,
    #[serde(rename = "series")]
    Series,
    #[serde(rename = "category")]
    Category,
    #[serde(rename = "ptInSeries")]
    PtInSeries,
    #[serde(rename = "ptInCategory")]
    PtInCategory,
}

impl std::fmt::Display for STTLChartSubelementType {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::GridLegend => write!(f, "gridLegend"),
            Self::Series => write!(f, "series"),
            Self::Category => write!(f, "category"),
            Self::PtInSeries => write!(f, "ptInSeries"),
            Self::PtInCategory => write!(f, "ptInCategory"),
        }
    }
}

impl std::str::FromStr for STTLChartSubelementType {
    type Err = String;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        match s {
            "gridLegend" => Ok(Self::GridLegend),
            "series" => Ok(Self::Series),
            "category" => Ok(Self::Category),
            "ptInSeries" => Ok(Self::PtInSeries),
            "ptInCategory" => Ok(Self::PtInCategory),
            _ => Err(format!("unknown STTLChartSubelementType value: {}", s)),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum STTLTriggerRuntimeNode {
    #[serde(rename = "first")]
    First,
    #[serde(rename = "last")]
    Last,
    #[serde(rename = "all")]
    All,
}

impl std::fmt::Display for STTLTriggerRuntimeNode {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::First => write!(f, "first"),
            Self::Last => write!(f, "last"),
            Self::All => write!(f, "all"),
        }
    }
}

impl std::str::FromStr for STTLTriggerRuntimeNode {
    type Err = String;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        match s {
            "first" => Ok(Self::First),
            "last" => Ok(Self::Last),
            "all" => Ok(Self::All),
            _ => Err(format!("unknown STTLTriggerRuntimeNode value: {}", s)),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum STTLTriggerEvent {
    #[serde(rename = "onBegin")]
    OnBegin,
    #[serde(rename = "onEnd")]
    OnEnd,
    #[serde(rename = "begin")]
    Begin,
    #[serde(rename = "end")]
    End,
    #[serde(rename = "onClick")]
    OnClick,
    #[serde(rename = "onDblClick")]
    OnDblClick,
    #[serde(rename = "onMouseOver")]
    OnMouseOver,
    #[serde(rename = "onMouseOut")]
    OnMouseOut,
    #[serde(rename = "onNext")]
    OnNext,
    #[serde(rename = "onPrev")]
    OnPrev,
    #[serde(rename = "onStopAudio")]
    OnStopAudio,
}

impl std::fmt::Display for STTLTriggerEvent {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::OnBegin => write!(f, "onBegin"),
            Self::OnEnd => write!(f, "onEnd"),
            Self::Begin => write!(f, "begin"),
            Self::End => write!(f, "end"),
            Self::OnClick => write!(f, "onClick"),
            Self::OnDblClick => write!(f, "onDblClick"),
            Self::OnMouseOver => write!(f, "onMouseOver"),
            Self::OnMouseOut => write!(f, "onMouseOut"),
            Self::OnNext => write!(f, "onNext"),
            Self::OnPrev => write!(f, "onPrev"),
            Self::OnStopAudio => write!(f, "onStopAudio"),
        }
    }
}

impl std::str::FromStr for STTLTriggerEvent {
    type Err = String;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        match s {
            "onBegin" => Ok(Self::OnBegin),
            "onEnd" => Ok(Self::OnEnd),
            "begin" => Ok(Self::Begin),
            "end" => Ok(Self::End),
            "onClick" => Ok(Self::OnClick),
            "onDblClick" => Ok(Self::OnDblClick),
            "onMouseOver" => Ok(Self::OnMouseOver),
            "onMouseOut" => Ok(Self::OnMouseOut),
            "onNext" => Ok(Self::OnNext),
            "onPrev" => Ok(Self::OnPrev),
            "onStopAudio" => Ok(Self::OnStopAudio),
            _ => Err(format!("unknown STTLTriggerEvent value: {}", s)),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum STTLTimeNodePresetClassType {
    #[serde(rename = "entr")]
    Entr,
    #[serde(rename = "exit")]
    Exit,
    #[serde(rename = "emph")]
    Emph,
    #[serde(rename = "path")]
    Path,
    #[serde(rename = "verb")]
    Verb,
    #[serde(rename = "mediacall")]
    Mediacall,
}

impl std::fmt::Display for STTLTimeNodePresetClassType {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::Entr => write!(f, "entr"),
            Self::Exit => write!(f, "exit"),
            Self::Emph => write!(f, "emph"),
            Self::Path => write!(f, "path"),
            Self::Verb => write!(f, "verb"),
            Self::Mediacall => write!(f, "mediacall"),
        }
    }
}

impl std::str::FromStr for STTLTimeNodePresetClassType {
    type Err = String;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        match s {
            "entr" => Ok(Self::Entr),
            "exit" => Ok(Self::Exit),
            "emph" => Ok(Self::Emph),
            "path" => Ok(Self::Path),
            "verb" => Ok(Self::Verb),
            "mediacall" => Ok(Self::Mediacall),
            _ => Err(format!("unknown STTLTimeNodePresetClassType value: {}", s)),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum STTLTimeNodeRestartType {
    #[serde(rename = "always")]
    Always,
    #[serde(rename = "whenNotActive")]
    WhenNotActive,
    #[serde(rename = "never")]
    Never,
}

impl std::fmt::Display for STTLTimeNodeRestartType {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::Always => write!(f, "always"),
            Self::WhenNotActive => write!(f, "whenNotActive"),
            Self::Never => write!(f, "never"),
        }
    }
}

impl std::str::FromStr for STTLTimeNodeRestartType {
    type Err = String;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        match s {
            "always" => Ok(Self::Always),
            "whenNotActive" => Ok(Self::WhenNotActive),
            "never" => Ok(Self::Never),
            _ => Err(format!("unknown STTLTimeNodeRestartType value: {}", s)),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum STTLTimeNodeFillType {
    #[serde(rename = "remove")]
    Remove,
    #[serde(rename = "freeze")]
    Freeze,
    #[serde(rename = "hold")]
    Hold,
    #[serde(rename = "transition")]
    Transition,
}

impl std::fmt::Display for STTLTimeNodeFillType {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::Remove => write!(f, "remove"),
            Self::Freeze => write!(f, "freeze"),
            Self::Hold => write!(f, "hold"),
            Self::Transition => write!(f, "transition"),
        }
    }
}

impl std::str::FromStr for STTLTimeNodeFillType {
    type Err = String;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        match s {
            "remove" => Ok(Self::Remove),
            "freeze" => Ok(Self::Freeze),
            "hold" => Ok(Self::Hold),
            "transition" => Ok(Self::Transition),
            _ => Err(format!("unknown STTLTimeNodeFillType value: {}", s)),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum STTLTimeNodeSyncType {
    #[serde(rename = "canSlip")]
    CanSlip,
    #[serde(rename = "locked")]
    Locked,
}

impl std::fmt::Display for STTLTimeNodeSyncType {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::CanSlip => write!(f, "canSlip"),
            Self::Locked => write!(f, "locked"),
        }
    }
}

impl std::str::FromStr for STTLTimeNodeSyncType {
    type Err = String;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        match s {
            "canSlip" => Ok(Self::CanSlip),
            "locked" => Ok(Self::Locked),
            _ => Err(format!("unknown STTLTimeNodeSyncType value: {}", s)),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum STTLTimeNodeMasterRelation {
    #[serde(rename = "sameClick")]
    SameClick,
    #[serde(rename = "lastClick")]
    LastClick,
    #[serde(rename = "nextClick")]
    NextClick,
}

impl std::fmt::Display for STTLTimeNodeMasterRelation {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::SameClick => write!(f, "sameClick"),
            Self::LastClick => write!(f, "lastClick"),
            Self::NextClick => write!(f, "nextClick"),
        }
    }
}

impl std::str::FromStr for STTLTimeNodeMasterRelation {
    type Err = String;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        match s {
            "sameClick" => Ok(Self::SameClick),
            "lastClick" => Ok(Self::LastClick),
            "nextClick" => Ok(Self::NextClick),
            _ => Err(format!("unknown STTLTimeNodeMasterRelation value: {}", s)),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum STTLTimeNodeType {
    #[serde(rename = "clickEffect")]
    ClickEffect,
    #[serde(rename = "withEffect")]
    WithEffect,
    #[serde(rename = "afterEffect")]
    AfterEffect,
    #[serde(rename = "mainSeq")]
    MainSeq,
    #[serde(rename = "interactiveSeq")]
    InteractiveSeq,
    #[serde(rename = "clickPar")]
    ClickPar,
    #[serde(rename = "withGroup")]
    WithGroup,
    #[serde(rename = "afterGroup")]
    AfterGroup,
    #[serde(rename = "tmRoot")]
    TmRoot,
}

impl std::fmt::Display for STTLTimeNodeType {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::ClickEffect => write!(f, "clickEffect"),
            Self::WithEffect => write!(f, "withEffect"),
            Self::AfterEffect => write!(f, "afterEffect"),
            Self::MainSeq => write!(f, "mainSeq"),
            Self::InteractiveSeq => write!(f, "interactiveSeq"),
            Self::ClickPar => write!(f, "clickPar"),
            Self::WithGroup => write!(f, "withGroup"),
            Self::AfterGroup => write!(f, "afterGroup"),
            Self::TmRoot => write!(f, "tmRoot"),
        }
    }
}

impl std::str::FromStr for STTLTimeNodeType {
    type Err = String;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        match s {
            "clickEffect" => Ok(Self::ClickEffect),
            "withEffect" => Ok(Self::WithEffect),
            "afterEffect" => Ok(Self::AfterEffect),
            "mainSeq" => Ok(Self::MainSeq),
            "interactiveSeq" => Ok(Self::InteractiveSeq),
            "clickPar" => Ok(Self::ClickPar),
            "withGroup" => Ok(Self::WithGroup),
            "afterGroup" => Ok(Self::AfterGroup),
            "tmRoot" => Ok(Self::TmRoot),
            _ => Err(format!("unknown STTLTimeNodeType value: {}", s)),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum STTLNextActionType {
    #[serde(rename = "none")]
    None,
    #[serde(rename = "seek")]
    Seek,
}

impl std::fmt::Display for STTLNextActionType {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::None => write!(f, "none"),
            Self::Seek => write!(f, "seek"),
        }
    }
}

impl std::str::FromStr for STTLNextActionType {
    type Err = String;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        match s {
            "none" => Ok(Self::None),
            "seek" => Ok(Self::Seek),
            _ => Err(format!("unknown STTLNextActionType value: {}", s)),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum STTLPreviousActionType {
    #[serde(rename = "none")]
    None,
    #[serde(rename = "skipTimed")]
    SkipTimed,
}

impl std::fmt::Display for STTLPreviousActionType {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::None => write!(f, "none"),
            Self::SkipTimed => write!(f, "skipTimed"),
        }
    }
}

impl std::str::FromStr for STTLPreviousActionType {
    type Err = String;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        match s {
            "none" => Ok(Self::None),
            "skipTimed" => Ok(Self::SkipTimed),
            _ => Err(format!("unknown STTLPreviousActionType value: {}", s)),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum STTLBehaviorAdditiveType {
    #[serde(rename = "base")]
    Base,
    #[serde(rename = "sum")]
    Sum,
    #[serde(rename = "repl")]
    Repl,
    #[serde(rename = "mult")]
    Mult,
    #[serde(rename = "none")]
    None,
}

impl std::fmt::Display for STTLBehaviorAdditiveType {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::Base => write!(f, "base"),
            Self::Sum => write!(f, "sum"),
            Self::Repl => write!(f, "repl"),
            Self::Mult => write!(f, "mult"),
            Self::None => write!(f, "none"),
        }
    }
}

impl std::str::FromStr for STTLBehaviorAdditiveType {
    type Err = String;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        match s {
            "base" => Ok(Self::Base),
            "sum" => Ok(Self::Sum),
            "repl" => Ok(Self::Repl),
            "mult" => Ok(Self::Mult),
            "none" => Ok(Self::None),
            _ => Err(format!("unknown STTLBehaviorAdditiveType value: {}", s)),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum STTLBehaviorAccumulateType {
    #[serde(rename = "none")]
    None,
    #[serde(rename = "always")]
    Always,
}

impl std::fmt::Display for STTLBehaviorAccumulateType {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::None => write!(f, "none"),
            Self::Always => write!(f, "always"),
        }
    }
}

impl std::str::FromStr for STTLBehaviorAccumulateType {
    type Err = String;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        match s {
            "none" => Ok(Self::None),
            "always" => Ok(Self::Always),
            _ => Err(format!("unknown STTLBehaviorAccumulateType value: {}", s)),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum STTLBehaviorTransformType {
    #[serde(rename = "pt")]
    Pt,
    #[serde(rename = "img")]
    Img,
}

impl std::fmt::Display for STTLBehaviorTransformType {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::Pt => write!(f, "pt"),
            Self::Img => write!(f, "img"),
        }
    }
}

impl std::str::FromStr for STTLBehaviorTransformType {
    type Err = String;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        match s {
            "pt" => Ok(Self::Pt),
            "img" => Ok(Self::Img),
            _ => Err(format!("unknown STTLBehaviorTransformType value: {}", s)),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum STTLBehaviorOverrideType {
    #[serde(rename = "normal")]
    Normal,
    #[serde(rename = "childStyle")]
    ChildStyle,
}

impl std::fmt::Display for STTLBehaviorOverrideType {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::Normal => write!(f, "normal"),
            Self::ChildStyle => write!(f, "childStyle"),
        }
    }
}

impl std::str::FromStr for STTLBehaviorOverrideType {
    type Err = String;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        match s {
            "normal" => Ok(Self::Normal),
            "childStyle" => Ok(Self::ChildStyle),
            _ => Err(format!("unknown STTLBehaviorOverrideType value: {}", s)),
        }
    }
}

pub type STTLTimeAnimateValueTime = String;

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum STTLAnimateBehaviorCalcMode {
    #[serde(rename = "discrete")]
    Discrete,
    #[serde(rename = "lin")]
    Lin,
    #[serde(rename = "fmla")]
    Fmla,
}

impl std::fmt::Display for STTLAnimateBehaviorCalcMode {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::Discrete => write!(f, "discrete"),
            Self::Lin => write!(f, "lin"),
            Self::Fmla => write!(f, "fmla"),
        }
    }
}

impl std::str::FromStr for STTLAnimateBehaviorCalcMode {
    type Err = String;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        match s {
            "discrete" => Ok(Self::Discrete),
            "lin" => Ok(Self::Lin),
            "fmla" => Ok(Self::Fmla),
            _ => Err(format!("unknown STTLAnimateBehaviorCalcMode value: {}", s)),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum STTLAnimateBehaviorValueType {
    #[serde(rename = "str")]
    Str,
    #[serde(rename = "num")]
    Num,
    #[serde(rename = "clr")]
    Clr,
}

impl std::fmt::Display for STTLAnimateBehaviorValueType {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::Str => write!(f, "str"),
            Self::Num => write!(f, "num"),
            Self::Clr => write!(f, "clr"),
        }
    }
}

impl std::str::FromStr for STTLAnimateBehaviorValueType {
    type Err = String;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        match s {
            "str" => Ok(Self::Str),
            "num" => Ok(Self::Num),
            "clr" => Ok(Self::Clr),
            _ => Err(format!("unknown STTLAnimateBehaviorValueType value: {}", s)),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum STTLAnimateColorSpace {
    #[serde(rename = "rgb")]
    Rgb,
    #[serde(rename = "hsl")]
    Hsl,
}

impl std::fmt::Display for STTLAnimateColorSpace {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::Rgb => write!(f, "rgb"),
            Self::Hsl => write!(f, "hsl"),
        }
    }
}

impl std::str::FromStr for STTLAnimateColorSpace {
    type Err = String;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        match s {
            "rgb" => Ok(Self::Rgb),
            "hsl" => Ok(Self::Hsl),
            _ => Err(format!("unknown STTLAnimateColorSpace value: {}", s)),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum STTLAnimateColorDirection {
    #[serde(rename = "cw")]
    Cw,
    #[serde(rename = "ccw")]
    Ccw,
}

impl std::fmt::Display for STTLAnimateColorDirection {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::Cw => write!(f, "cw"),
            Self::Ccw => write!(f, "ccw"),
        }
    }
}

impl std::str::FromStr for STTLAnimateColorDirection {
    type Err = String;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        match s {
            "cw" => Ok(Self::Cw),
            "ccw" => Ok(Self::Ccw),
            _ => Err(format!("unknown STTLAnimateColorDirection value: {}", s)),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum STTLAnimateEffectTransition {
    #[serde(rename = "in")]
    In,
    #[serde(rename = "out")]
    Out,
    #[serde(rename = "none")]
    None,
}

impl std::fmt::Display for STTLAnimateEffectTransition {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::In => write!(f, "in"),
            Self::Out => write!(f, "out"),
            Self::None => write!(f, "none"),
        }
    }
}

impl std::str::FromStr for STTLAnimateEffectTransition {
    type Err = String;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        match s {
            "in" => Ok(Self::In),
            "out" => Ok(Self::Out),
            "none" => Ok(Self::None),
            _ => Err(format!("unknown STTLAnimateEffectTransition value: {}", s)),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum STTLAnimateMotionBehaviorOrigin {
    #[serde(rename = "parent")]
    Parent,
    #[serde(rename = "layout")]
    Layout,
}

impl std::fmt::Display for STTLAnimateMotionBehaviorOrigin {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::Parent => write!(f, "parent"),
            Self::Layout => write!(f, "layout"),
        }
    }
}

impl std::str::FromStr for STTLAnimateMotionBehaviorOrigin {
    type Err = String;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        match s {
            "parent" => Ok(Self::Parent),
            "layout" => Ok(Self::Layout),
            _ => Err(format!(
                "unknown STTLAnimateMotionBehaviorOrigin value: {}",
                s
            )),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum STTLAnimateMotionPathEditMode {
    #[serde(rename = "relative")]
    Relative,
    #[serde(rename = "fixed")]
    Fixed,
}

impl std::fmt::Display for STTLAnimateMotionPathEditMode {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::Relative => write!(f, "relative"),
            Self::Fixed => write!(f, "fixed"),
        }
    }
}

impl std::str::FromStr for STTLAnimateMotionPathEditMode {
    type Err = String;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        match s {
            "relative" => Ok(Self::Relative),
            "fixed" => Ok(Self::Fixed),
            _ => Err(format!(
                "unknown STTLAnimateMotionPathEditMode value: {}",
                s
            )),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum STTLCommandType {
    #[serde(rename = "evt")]
    Evt,
    #[serde(rename = "call")]
    Call,
    #[serde(rename = "verb")]
    Verb,
}

impl std::fmt::Display for STTLCommandType {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::Evt => write!(f, "evt"),
            Self::Call => write!(f, "call"),
            Self::Verb => write!(f, "verb"),
        }
    }
}

impl std::str::FromStr for STTLCommandType {
    type Err = String;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        match s {
            "evt" => Ok(Self::Evt),
            "call" => Ok(Self::Call),
            "verb" => Ok(Self::Verb),
            _ => Err(format!("unknown STTLCommandType value: {}", s)),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum STTLParaBuildType {
    #[serde(rename = "allAtOnce")]
    AllAtOnce,
    #[serde(rename = "p")]
    P,
    #[serde(rename = "cust")]
    Cust,
    #[serde(rename = "whole")]
    Whole,
}

impl std::fmt::Display for STTLParaBuildType {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::AllAtOnce => write!(f, "allAtOnce"),
            Self::P => write!(f, "p"),
            Self::Cust => write!(f, "cust"),
            Self::Whole => write!(f, "whole"),
        }
    }
}

impl std::str::FromStr for STTLParaBuildType {
    type Err = String;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        match s {
            "allAtOnce" => Ok(Self::AllAtOnce),
            "p" => Ok(Self::P),
            "cust" => Ok(Self::Cust),
            "whole" => Ok(Self::Whole),
            _ => Err(format!("unknown STTLParaBuildType value: {}", s)),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum STTLDiagramBuildType {
    #[serde(rename = "whole")]
    Whole,
    #[serde(rename = "depthByNode")]
    DepthByNode,
    #[serde(rename = "depthByBranch")]
    DepthByBranch,
    #[serde(rename = "breadthByNode")]
    BreadthByNode,
    #[serde(rename = "breadthByLvl")]
    BreadthByLvl,
    #[serde(rename = "cw")]
    Cw,
    #[serde(rename = "cwIn")]
    CwIn,
    #[serde(rename = "cwOut")]
    CwOut,
    #[serde(rename = "ccw")]
    Ccw,
    #[serde(rename = "ccwIn")]
    CcwIn,
    #[serde(rename = "ccwOut")]
    CcwOut,
    #[serde(rename = "inByRing")]
    InByRing,
    #[serde(rename = "outByRing")]
    OutByRing,
    #[serde(rename = "up")]
    Up,
    #[serde(rename = "down")]
    Down,
    #[serde(rename = "allAtOnce")]
    AllAtOnce,
    #[serde(rename = "cust")]
    Cust,
}

impl std::fmt::Display for STTLDiagramBuildType {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::Whole => write!(f, "whole"),
            Self::DepthByNode => write!(f, "depthByNode"),
            Self::DepthByBranch => write!(f, "depthByBranch"),
            Self::BreadthByNode => write!(f, "breadthByNode"),
            Self::BreadthByLvl => write!(f, "breadthByLvl"),
            Self::Cw => write!(f, "cw"),
            Self::CwIn => write!(f, "cwIn"),
            Self::CwOut => write!(f, "cwOut"),
            Self::Ccw => write!(f, "ccw"),
            Self::CcwIn => write!(f, "ccwIn"),
            Self::CcwOut => write!(f, "ccwOut"),
            Self::InByRing => write!(f, "inByRing"),
            Self::OutByRing => write!(f, "outByRing"),
            Self::Up => write!(f, "up"),
            Self::Down => write!(f, "down"),
            Self::AllAtOnce => write!(f, "allAtOnce"),
            Self::Cust => write!(f, "cust"),
        }
    }
}

impl std::str::FromStr for STTLDiagramBuildType {
    type Err = String;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        match s {
            "whole" => Ok(Self::Whole),
            "depthByNode" => Ok(Self::DepthByNode),
            "depthByBranch" => Ok(Self::DepthByBranch),
            "breadthByNode" => Ok(Self::BreadthByNode),
            "breadthByLvl" => Ok(Self::BreadthByLvl),
            "cw" => Ok(Self::Cw),
            "cwIn" => Ok(Self::CwIn),
            "cwOut" => Ok(Self::CwOut),
            "ccw" => Ok(Self::Ccw),
            "ccwIn" => Ok(Self::CcwIn),
            "ccwOut" => Ok(Self::CcwOut),
            "inByRing" => Ok(Self::InByRing),
            "outByRing" => Ok(Self::OutByRing),
            "up" => Ok(Self::Up),
            "down" => Ok(Self::Down),
            "allAtOnce" => Ok(Self::AllAtOnce),
            "cust" => Ok(Self::Cust),
            _ => Err(format!("unknown STTLDiagramBuildType value: {}", s)),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum STTLOleChartBuildType {
    #[serde(rename = "allAtOnce")]
    AllAtOnce,
    #[serde(rename = "series")]
    Series,
    #[serde(rename = "category")]
    Category,
    #[serde(rename = "seriesEl")]
    SeriesEl,
    #[serde(rename = "categoryEl")]
    CategoryEl,
}

impl std::fmt::Display for STTLOleChartBuildType {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::AllAtOnce => write!(f, "allAtOnce"),
            Self::Series => write!(f, "series"),
            Self::Category => write!(f, "category"),
            Self::SeriesEl => write!(f, "seriesEl"),
            Self::CategoryEl => write!(f, "categoryEl"),
        }
    }
}

impl std::str::FromStr for STTLOleChartBuildType {
    type Err = String;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        match s {
            "allAtOnce" => Ok(Self::AllAtOnce),
            "series" => Ok(Self::Series),
            "category" => Ok(Self::Category),
            "seriesEl" => Ok(Self::SeriesEl),
            "categoryEl" => Ok(Self::CategoryEl),
            _ => Err(format!("unknown STTLOleChartBuildType value: {}", s)),
        }
    }
}

pub type STName = String;

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum STDirection {
    #[serde(rename = "horz")]
    Horz,
    #[serde(rename = "vert")]
    Vert,
}

impl std::fmt::Display for STDirection {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::Horz => write!(f, "horz"),
            Self::Vert => write!(f, "vert"),
        }
    }
}

impl std::str::FromStr for STDirection {
    type Err = String;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        match s {
            "horz" => Ok(Self::Horz),
            "vert" => Ok(Self::Vert),
            _ => Err(format!("unknown STDirection value: {}", s)),
        }
    }
}

pub type STIndex = u32;

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum STOleObjectFollowColorScheme {
    #[serde(rename = "none")]
    None,
    #[serde(rename = "full")]
    Full,
    #[serde(rename = "textAndBackground")]
    TextAndBackground,
}

impl std::fmt::Display for STOleObjectFollowColorScheme {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::None => write!(f, "none"),
            Self::Full => write!(f, "full"),
            Self::TextAndBackground => write!(f, "textAndBackground"),
        }
    }
}

impl std::str::FromStr for STOleObjectFollowColorScheme {
    type Err = String;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        match s {
            "none" => Ok(Self::None),
            "full" => Ok(Self::Full),
            "textAndBackground" => Ok(Self::TextAndBackground),
            _ => Err(format!("unknown STOleObjectFollowColorScheme value: {}", s)),
        }
    }
}

pub type STSlideId = u32;

pub type STSlideMasterId = u32;

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum STPhotoAlbumLayout {
    #[serde(rename = "fitToSlide")]
    FitToSlide,
    #[serde(rename = "1pic")]
    _1pic,
    #[serde(rename = "2pic")]
    _2pic,
    #[serde(rename = "4pic")]
    _4pic,
    #[serde(rename = "1picTitle")]
    _1picTitle,
    #[serde(rename = "2picTitle")]
    _2picTitle,
    #[serde(rename = "4picTitle")]
    _4picTitle,
}

impl std::fmt::Display for STPhotoAlbumLayout {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::FitToSlide => write!(f, "fitToSlide"),
            Self::_1pic => write!(f, "1pic"),
            Self::_2pic => write!(f, "2pic"),
            Self::_4pic => write!(f, "4pic"),
            Self::_1picTitle => write!(f, "1picTitle"),
            Self::_2picTitle => write!(f, "2picTitle"),
            Self::_4picTitle => write!(f, "4picTitle"),
        }
    }
}

impl std::str::FromStr for STPhotoAlbumLayout {
    type Err = String;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        match s {
            "fitToSlide" => Ok(Self::FitToSlide),
            "1pic" => Ok(Self::_1pic),
            "2pic" => Ok(Self::_2pic),
            "4pic" => Ok(Self::_4pic),
            "1picTitle" => Ok(Self::_1picTitle),
            "2picTitle" => Ok(Self::_2picTitle),
            "4picTitle" => Ok(Self::_4picTitle),
            _ => Err(format!("unknown STPhotoAlbumLayout value: {}", s)),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum STPhotoAlbumFrameShape {
    #[serde(rename = "frameStyle1")]
    FrameStyle1,
    #[serde(rename = "frameStyle2")]
    FrameStyle2,
    #[serde(rename = "frameStyle3")]
    FrameStyle3,
    #[serde(rename = "frameStyle4")]
    FrameStyle4,
    #[serde(rename = "frameStyle5")]
    FrameStyle5,
    #[serde(rename = "frameStyle6")]
    FrameStyle6,
    #[serde(rename = "frameStyle7")]
    FrameStyle7,
}

impl std::fmt::Display for STPhotoAlbumFrameShape {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::FrameStyle1 => write!(f, "frameStyle1"),
            Self::FrameStyle2 => write!(f, "frameStyle2"),
            Self::FrameStyle3 => write!(f, "frameStyle3"),
            Self::FrameStyle4 => write!(f, "frameStyle4"),
            Self::FrameStyle5 => write!(f, "frameStyle5"),
            Self::FrameStyle6 => write!(f, "frameStyle6"),
            Self::FrameStyle7 => write!(f, "frameStyle7"),
        }
    }
}

impl std::str::FromStr for STPhotoAlbumFrameShape {
    type Err = String;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        match s {
            "frameStyle1" => Ok(Self::FrameStyle1),
            "frameStyle2" => Ok(Self::FrameStyle2),
            "frameStyle3" => Ok(Self::FrameStyle3),
            "frameStyle4" => Ok(Self::FrameStyle4),
            "frameStyle5" => Ok(Self::FrameStyle5),
            "frameStyle6" => Ok(Self::FrameStyle6),
            "frameStyle7" => Ok(Self::FrameStyle7),
            _ => Err(format!("unknown STPhotoAlbumFrameShape value: {}", s)),
        }
    }
}

pub type STSlideSizeCoordinate = i32;

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum STSlideSizeType {
    #[serde(rename = "screen4x3")]
    Screen4x3,
    #[serde(rename = "letter")]
    Letter,
    #[serde(rename = "A4")]
    A4,
    #[serde(rename = "35mm")]
    _35mm,
    #[serde(rename = "overhead")]
    Overhead,
    #[serde(rename = "banner")]
    Banner,
    #[serde(rename = "custom")]
    Custom,
    #[serde(rename = "ledger")]
    Ledger,
    #[serde(rename = "A3")]
    A3,
    #[serde(rename = "B4ISO")]
    B4ISO,
    #[serde(rename = "B5ISO")]
    B5ISO,
    #[serde(rename = "B4JIS")]
    B4JIS,
    #[serde(rename = "B5JIS")]
    B5JIS,
    #[serde(rename = "hagakiCard")]
    HagakiCard,
    #[serde(rename = "screen16x9")]
    Screen16x9,
    #[serde(rename = "screen16x10")]
    Screen16x10,
}

impl std::fmt::Display for STSlideSizeType {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::Screen4x3 => write!(f, "screen4x3"),
            Self::Letter => write!(f, "letter"),
            Self::A4 => write!(f, "A4"),
            Self::_35mm => write!(f, "35mm"),
            Self::Overhead => write!(f, "overhead"),
            Self::Banner => write!(f, "banner"),
            Self::Custom => write!(f, "custom"),
            Self::Ledger => write!(f, "ledger"),
            Self::A3 => write!(f, "A3"),
            Self::B4ISO => write!(f, "B4ISO"),
            Self::B5ISO => write!(f, "B5ISO"),
            Self::B4JIS => write!(f, "B4JIS"),
            Self::B5JIS => write!(f, "B5JIS"),
            Self::HagakiCard => write!(f, "hagakiCard"),
            Self::Screen16x9 => write!(f, "screen16x9"),
            Self::Screen16x10 => write!(f, "screen16x10"),
        }
    }
}

impl std::str::FromStr for STSlideSizeType {
    type Err = String;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        match s {
            "screen4x3" => Ok(Self::Screen4x3),
            "letter" => Ok(Self::Letter),
            "A4" => Ok(Self::A4),
            "35mm" => Ok(Self::_35mm),
            "overhead" => Ok(Self::Overhead),
            "banner" => Ok(Self::Banner),
            "custom" => Ok(Self::Custom),
            "ledger" => Ok(Self::Ledger),
            "A3" => Ok(Self::A3),
            "B4ISO" => Ok(Self::B4ISO),
            "B5ISO" => Ok(Self::B5ISO),
            "B4JIS" => Ok(Self::B4JIS),
            "B5JIS" => Ok(Self::B5JIS),
            "hagakiCard" => Ok(Self::HagakiCard),
            "screen16x9" => Ok(Self::Screen16x9),
            "screen16x10" => Ok(Self::Screen16x10),
            _ => Err(format!("unknown STSlideSizeType value: {}", s)),
        }
    }
}

pub type STBookmarkIdSeed = u32;

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum STWebColorType {
    #[serde(rename = "none")]
    None,
    #[serde(rename = "browser")]
    Browser,
    #[serde(rename = "presentationText")]
    PresentationText,
    #[serde(rename = "presentationAccent")]
    PresentationAccent,
    #[serde(rename = "whiteTextOnBlack")]
    WhiteTextOnBlack,
    #[serde(rename = "blackTextOnWhite")]
    BlackTextOnWhite,
}

impl std::fmt::Display for STWebColorType {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::None => write!(f, "none"),
            Self::Browser => write!(f, "browser"),
            Self::PresentationText => write!(f, "presentationText"),
            Self::PresentationAccent => write!(f, "presentationAccent"),
            Self::WhiteTextOnBlack => write!(f, "whiteTextOnBlack"),
            Self::BlackTextOnWhite => write!(f, "blackTextOnWhite"),
        }
    }
}

impl std::str::FromStr for STWebColorType {
    type Err = String;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        match s {
            "none" => Ok(Self::None),
            "browser" => Ok(Self::Browser),
            "presentationText" => Ok(Self::PresentationText),
            "presentationAccent" => Ok(Self::PresentationAccent),
            "whiteTextOnBlack" => Ok(Self::WhiteTextOnBlack),
            "blackTextOnWhite" => Ok(Self::BlackTextOnWhite),
            _ => Err(format!("unknown STWebColorType value: {}", s)),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum STWebScreenSize {
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
    #[serde(rename = "1800x1400")]
    _1800x1400,
    #[serde(rename = "1920x1200")]
    _1920x1200,
}

impl std::fmt::Display for STWebScreenSize {
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
            Self::_1800x1400 => write!(f, "1800x1400"),
            Self::_1920x1200 => write!(f, "1920x1200"),
        }
    }
}

impl std::str::FromStr for STWebScreenSize {
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
            "1800x1400" => Ok(Self::_1800x1400),
            "1920x1200" => Ok(Self::_1920x1200),
            _ => Err(format!("unknown STWebScreenSize value: {}", s)),
        }
    }
}

pub type STWebEncoding = String;

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum STPrintWhat {
    #[serde(rename = "slides")]
    Slides,
    #[serde(rename = "handouts1")]
    Handouts1,
    #[serde(rename = "handouts2")]
    Handouts2,
    #[serde(rename = "handouts3")]
    Handouts3,
    #[serde(rename = "handouts4")]
    Handouts4,
    #[serde(rename = "handouts6")]
    Handouts6,
    #[serde(rename = "handouts9")]
    Handouts9,
    #[serde(rename = "notes")]
    Notes,
    #[serde(rename = "outline")]
    Outline,
}

impl std::fmt::Display for STPrintWhat {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::Slides => write!(f, "slides"),
            Self::Handouts1 => write!(f, "handouts1"),
            Self::Handouts2 => write!(f, "handouts2"),
            Self::Handouts3 => write!(f, "handouts3"),
            Self::Handouts4 => write!(f, "handouts4"),
            Self::Handouts6 => write!(f, "handouts6"),
            Self::Handouts9 => write!(f, "handouts9"),
            Self::Notes => write!(f, "notes"),
            Self::Outline => write!(f, "outline"),
        }
    }
}

impl std::str::FromStr for STPrintWhat {
    type Err = String;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        match s {
            "slides" => Ok(Self::Slides),
            "handouts1" => Ok(Self::Handouts1),
            "handouts2" => Ok(Self::Handouts2),
            "handouts3" => Ok(Self::Handouts3),
            "handouts4" => Ok(Self::Handouts4),
            "handouts6" => Ok(Self::Handouts6),
            "handouts9" => Ok(Self::Handouts9),
            "notes" => Ok(Self::Notes),
            "outline" => Ok(Self::Outline),
            _ => Err(format!("unknown STPrintWhat value: {}", s)),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum STPrintColorMode {
    #[serde(rename = "bw")]
    Bw,
    #[serde(rename = "gray")]
    Gray,
    #[serde(rename = "clr")]
    Clr,
}

impl std::fmt::Display for STPrintColorMode {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::Bw => write!(f, "bw"),
            Self::Gray => write!(f, "gray"),
            Self::Clr => write!(f, "clr"),
        }
    }
}

impl std::str::FromStr for STPrintColorMode {
    type Err = String;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        match s {
            "bw" => Ok(Self::Bw),
            "gray" => Ok(Self::Gray),
            "clr" => Ok(Self::Clr),
            _ => Err(format!("unknown STPrintColorMode value: {}", s)),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum STPlaceholderType {
    #[serde(rename = "title")]
    Title,
    #[serde(rename = "body")]
    Body,
    #[serde(rename = "ctrTitle")]
    CtrTitle,
    #[serde(rename = "subTitle")]
    SubTitle,
    #[serde(rename = "dt")]
    Dt,
    #[serde(rename = "sldNum")]
    SldNum,
    #[serde(rename = "ftr")]
    Ftr,
    #[serde(rename = "hdr")]
    Hdr,
    #[serde(rename = "obj")]
    Obj,
    #[serde(rename = "chart")]
    Chart,
    #[serde(rename = "tbl")]
    Tbl,
    #[serde(rename = "clipArt")]
    ClipArt,
    #[serde(rename = "dgm")]
    Dgm,
    #[serde(rename = "media")]
    Media,
    #[serde(rename = "sldImg")]
    SldImg,
    #[serde(rename = "pic")]
    Pic,
}

impl std::fmt::Display for STPlaceholderType {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::Title => write!(f, "title"),
            Self::Body => write!(f, "body"),
            Self::CtrTitle => write!(f, "ctrTitle"),
            Self::SubTitle => write!(f, "subTitle"),
            Self::Dt => write!(f, "dt"),
            Self::SldNum => write!(f, "sldNum"),
            Self::Ftr => write!(f, "ftr"),
            Self::Hdr => write!(f, "hdr"),
            Self::Obj => write!(f, "obj"),
            Self::Chart => write!(f, "chart"),
            Self::Tbl => write!(f, "tbl"),
            Self::ClipArt => write!(f, "clipArt"),
            Self::Dgm => write!(f, "dgm"),
            Self::Media => write!(f, "media"),
            Self::SldImg => write!(f, "sldImg"),
            Self::Pic => write!(f, "pic"),
        }
    }
}

impl std::str::FromStr for STPlaceholderType {
    type Err = String;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        match s {
            "title" => Ok(Self::Title),
            "body" => Ok(Self::Body),
            "ctrTitle" => Ok(Self::CtrTitle),
            "subTitle" => Ok(Self::SubTitle),
            "dt" => Ok(Self::Dt),
            "sldNum" => Ok(Self::SldNum),
            "ftr" => Ok(Self::Ftr),
            "hdr" => Ok(Self::Hdr),
            "obj" => Ok(Self::Obj),
            "chart" => Ok(Self::Chart),
            "tbl" => Ok(Self::Tbl),
            "clipArt" => Ok(Self::ClipArt),
            "dgm" => Ok(Self::Dgm),
            "media" => Ok(Self::Media),
            "sldImg" => Ok(Self::SldImg),
            "pic" => Ok(Self::Pic),
            _ => Err(format!("unknown STPlaceholderType value: {}", s)),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum STPlaceholderSize {
    #[serde(rename = "full")]
    Full,
    #[serde(rename = "half")]
    Half,
    #[serde(rename = "quarter")]
    Quarter,
}

impl std::fmt::Display for STPlaceholderSize {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::Full => write!(f, "full"),
            Self::Half => write!(f, "half"),
            Self::Quarter => write!(f, "quarter"),
        }
    }
}

impl std::str::FromStr for STPlaceholderSize {
    type Err = String;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        match s {
            "full" => Ok(Self::Full),
            "half" => Ok(Self::Half),
            "quarter" => Ok(Self::Quarter),
            _ => Err(format!("unknown STPlaceholderSize value: {}", s)),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum STSlideLayoutType {
    #[serde(rename = "title")]
    Title,
    #[serde(rename = "tx")]
    Tx,
    #[serde(rename = "twoColTx")]
    TwoColTx,
    #[serde(rename = "tbl")]
    Tbl,
    #[serde(rename = "txAndChart")]
    TxAndChart,
    #[serde(rename = "chartAndTx")]
    ChartAndTx,
    #[serde(rename = "dgm")]
    Dgm,
    #[serde(rename = "chart")]
    Chart,
    #[serde(rename = "txAndClipArt")]
    TxAndClipArt,
    #[serde(rename = "clipArtAndTx")]
    ClipArtAndTx,
    #[serde(rename = "titleOnly")]
    TitleOnly,
    #[serde(rename = "blank")]
    Blank,
    #[serde(rename = "txAndObj")]
    TxAndObj,
    #[serde(rename = "objAndTx")]
    ObjAndTx,
    #[serde(rename = "objOnly")]
    ObjOnly,
    #[serde(rename = "obj")]
    Obj,
    #[serde(rename = "txAndMedia")]
    TxAndMedia,
    #[serde(rename = "mediaAndTx")]
    MediaAndTx,
    #[serde(rename = "objOverTx")]
    ObjOverTx,
    #[serde(rename = "txOverObj")]
    TxOverObj,
    #[serde(rename = "txAndTwoObj")]
    TxAndTwoObj,
    #[serde(rename = "twoObjAndTx")]
    TwoObjAndTx,
    #[serde(rename = "twoObjOverTx")]
    TwoObjOverTx,
    #[serde(rename = "fourObj")]
    FourObj,
    #[serde(rename = "vertTx")]
    VertTx,
    #[serde(rename = "clipArtAndVertTx")]
    ClipArtAndVertTx,
    #[serde(rename = "vertTitleAndTx")]
    VertTitleAndTx,
    #[serde(rename = "vertTitleAndTxOverChart")]
    VertTitleAndTxOverChart,
    #[serde(rename = "twoObj")]
    TwoObj,
    #[serde(rename = "objAndTwoObj")]
    ObjAndTwoObj,
    #[serde(rename = "twoObjAndObj")]
    TwoObjAndObj,
    #[serde(rename = "cust")]
    Cust,
    #[serde(rename = "secHead")]
    SecHead,
    #[serde(rename = "twoTxTwoObj")]
    TwoTxTwoObj,
    #[serde(rename = "objTx")]
    ObjTx,
    #[serde(rename = "picTx")]
    PicTx,
}

impl std::fmt::Display for STSlideLayoutType {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::Title => write!(f, "title"),
            Self::Tx => write!(f, "tx"),
            Self::TwoColTx => write!(f, "twoColTx"),
            Self::Tbl => write!(f, "tbl"),
            Self::TxAndChart => write!(f, "txAndChart"),
            Self::ChartAndTx => write!(f, "chartAndTx"),
            Self::Dgm => write!(f, "dgm"),
            Self::Chart => write!(f, "chart"),
            Self::TxAndClipArt => write!(f, "txAndClipArt"),
            Self::ClipArtAndTx => write!(f, "clipArtAndTx"),
            Self::TitleOnly => write!(f, "titleOnly"),
            Self::Blank => write!(f, "blank"),
            Self::TxAndObj => write!(f, "txAndObj"),
            Self::ObjAndTx => write!(f, "objAndTx"),
            Self::ObjOnly => write!(f, "objOnly"),
            Self::Obj => write!(f, "obj"),
            Self::TxAndMedia => write!(f, "txAndMedia"),
            Self::MediaAndTx => write!(f, "mediaAndTx"),
            Self::ObjOverTx => write!(f, "objOverTx"),
            Self::TxOverObj => write!(f, "txOverObj"),
            Self::TxAndTwoObj => write!(f, "txAndTwoObj"),
            Self::TwoObjAndTx => write!(f, "twoObjAndTx"),
            Self::TwoObjOverTx => write!(f, "twoObjOverTx"),
            Self::FourObj => write!(f, "fourObj"),
            Self::VertTx => write!(f, "vertTx"),
            Self::ClipArtAndVertTx => write!(f, "clipArtAndVertTx"),
            Self::VertTitleAndTx => write!(f, "vertTitleAndTx"),
            Self::VertTitleAndTxOverChart => write!(f, "vertTitleAndTxOverChart"),
            Self::TwoObj => write!(f, "twoObj"),
            Self::ObjAndTwoObj => write!(f, "objAndTwoObj"),
            Self::TwoObjAndObj => write!(f, "twoObjAndObj"),
            Self::Cust => write!(f, "cust"),
            Self::SecHead => write!(f, "secHead"),
            Self::TwoTxTwoObj => write!(f, "twoTxTwoObj"),
            Self::ObjTx => write!(f, "objTx"),
            Self::PicTx => write!(f, "picTx"),
        }
    }
}

impl std::str::FromStr for STSlideLayoutType {
    type Err = String;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        match s {
            "title" => Ok(Self::Title),
            "tx" => Ok(Self::Tx),
            "twoColTx" => Ok(Self::TwoColTx),
            "tbl" => Ok(Self::Tbl),
            "txAndChart" => Ok(Self::TxAndChart),
            "chartAndTx" => Ok(Self::ChartAndTx),
            "dgm" => Ok(Self::Dgm),
            "chart" => Ok(Self::Chart),
            "txAndClipArt" => Ok(Self::TxAndClipArt),
            "clipArtAndTx" => Ok(Self::ClipArtAndTx),
            "titleOnly" => Ok(Self::TitleOnly),
            "blank" => Ok(Self::Blank),
            "txAndObj" => Ok(Self::TxAndObj),
            "objAndTx" => Ok(Self::ObjAndTx),
            "objOnly" => Ok(Self::ObjOnly),
            "obj" => Ok(Self::Obj),
            "txAndMedia" => Ok(Self::TxAndMedia),
            "mediaAndTx" => Ok(Self::MediaAndTx),
            "objOverTx" => Ok(Self::ObjOverTx),
            "txOverObj" => Ok(Self::TxOverObj),
            "txAndTwoObj" => Ok(Self::TxAndTwoObj),
            "twoObjAndTx" => Ok(Self::TwoObjAndTx),
            "twoObjOverTx" => Ok(Self::TwoObjOverTx),
            "fourObj" => Ok(Self::FourObj),
            "vertTx" => Ok(Self::VertTx),
            "clipArtAndVertTx" => Ok(Self::ClipArtAndVertTx),
            "vertTitleAndTx" => Ok(Self::VertTitleAndTx),
            "vertTitleAndTxOverChart" => Ok(Self::VertTitleAndTxOverChart),
            "twoObj" => Ok(Self::TwoObj),
            "objAndTwoObj" => Ok(Self::ObjAndTwoObj),
            "twoObjAndObj" => Ok(Self::TwoObjAndObj),
            "cust" => Ok(Self::Cust),
            "secHead" => Ok(Self::SecHead),
            "twoTxTwoObj" => Ok(Self::TwoTxTwoObj),
            "objTx" => Ok(Self::ObjTx),
            "picTx" => Ok(Self::PicTx),
            _ => Err(format!("unknown STSlideLayoutType value: {}", s)),
        }
    }
}

pub type STSlideLayoutId = u32;

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum STSplitterBarState {
    #[serde(rename = "minimized")]
    Minimized,
    #[serde(rename = "restored")]
    Restored,
    #[serde(rename = "maximized")]
    Maximized,
}

impl std::fmt::Display for STSplitterBarState {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::Minimized => write!(f, "minimized"),
            Self::Restored => write!(f, "restored"),
            Self::Maximized => write!(f, "maximized"),
        }
    }
}

impl std::str::FromStr for STSplitterBarState {
    type Err = String;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        match s {
            "minimized" => Ok(Self::Minimized),
            "restored" => Ok(Self::Restored),
            "maximized" => Ok(Self::Maximized),
            _ => Err(format!("unknown STSplitterBarState value: {}", s)),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum STViewType {
    #[serde(rename = "sldView")]
    SldView,
    #[serde(rename = "sldMasterView")]
    SldMasterView,
    #[serde(rename = "notesView")]
    NotesView,
    #[serde(rename = "handoutView")]
    HandoutView,
    #[serde(rename = "notesMasterView")]
    NotesMasterView,
    #[serde(rename = "outlineView")]
    OutlineView,
    #[serde(rename = "sldSorterView")]
    SldSorterView,
    #[serde(rename = "sldThumbnailView")]
    SldThumbnailView,
}

impl std::fmt::Display for STViewType {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::SldView => write!(f, "sldView"),
            Self::SldMasterView => write!(f, "sldMasterView"),
            Self::NotesView => write!(f, "notesView"),
            Self::HandoutView => write!(f, "handoutView"),
            Self::NotesMasterView => write!(f, "notesMasterView"),
            Self::OutlineView => write!(f, "outlineView"),
            Self::SldSorterView => write!(f, "sldSorterView"),
            Self::SldThumbnailView => write!(f, "sldThumbnailView"),
        }
    }
}

impl std::str::FromStr for STViewType {
    type Err = String;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        match s {
            "sldView" => Ok(Self::SldView),
            "sldMasterView" => Ok(Self::SldMasterView),
            "notesView" => Ok(Self::NotesView),
            "handoutView" => Ok(Self::HandoutView),
            "notesMasterView" => Ok(Self::NotesMasterView),
            "outlineView" => Ok(Self::OutlineView),
            "sldSorterView" => Ok(Self::SldSorterView),
            "sldThumbnailView" => Ok(Self::SldThumbnailView),
            _ => Err(format!("unknown STViewType value: {}", s)),
        }
    }
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub enum EGSlideListChoice {
    #[serde(rename = "sldAll")]
    SldAll(Box<CTEmpty>),
    #[serde(rename = "sldRg")]
    SldRg(Box<CTIndexRange>),
    #[serde(rename = "custShow")]
    CustShow(Box<CTCustomShowId>),
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub enum EGShowType {
    #[serde(rename = "present")]
    Present(Box<CTEmpty>),
    #[serde(rename = "browse")]
    Browse(Box<CTShowInfoBrowse>),
    #[serde(rename = "kiosk")]
    Kiosk(Box<CTShowInfoKiosk>),
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub enum EGBackground {
    #[serde(rename = "bgPr")]
    BgPr(Box<CTBackgroundProperties>),
    #[serde(rename = "bgRef")]
    BgRef(String),
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTSideDirectionTransition {
    #[serde(rename = "@dir")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub dir: Option<STTransitionSideDirectionType>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTCornerDirectionTransition {
    #[serde(rename = "@dir")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub dir: Option<STTransitionCornerDirectionType>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTEightDirectionTransition {
    #[serde(rename = "@dir")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub dir: Option<STTransitionEightDirectionType>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTOrientationTransition {
    #[serde(rename = "@dir")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub dir: Option<STDirection>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTInOutTransition {
    #[serde(rename = "@dir")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub dir: Option<STTransitionInOutDirectionType>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTOptionalBlackTransition {
    #[serde(rename = "@thruBlk")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub thru_blk: Option<bool>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTSplitTransition {
    #[serde(rename = "@orient")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub orient: Option<STDirection>,
    #[serde(rename = "@dir")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub dir: Option<STTransitionInOutDirectionType>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTWheelTransition {
    #[serde(rename = "@spokes")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub spokes: Option<u32>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTTransitionStartSoundAction {
    #[serde(rename = "@loop")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub r#loop: Option<bool>,
    #[serde(rename = "snd")]
    pub snd: String,
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
    pub extra_children: Vec<ooxml_xml::RawXmlNode>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct CTTransitionSoundAction;

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct SlideTransition {
    #[serde(rename = "@spd")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub spd: Option<STTransitionSpeed>,
    #[serde(rename = "@advClick")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub adv_click: Option<bool>,
    #[serde(rename = "@advTm")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub adv_tm: Option<u32>,
    #[serde(rename = "sndAc")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub snd_ac: Option<Box<CTTransitionSoundAction>>,
    #[serde(rename = "extLst")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub ext_lst: Option<Box<CTExtensionListModify>>,
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
    pub extra_children: Vec<ooxml_xml::RawXmlNode>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTTLIterateIntervalTime {
    #[serde(rename = "@val")]
    pub value: STTLTime,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTTLIterateIntervalPercentage {
    #[serde(rename = "@val")]
    pub value: String,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTTLIterateData {
    #[serde(rename = "@type")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub r#type: Option<STIterateType>,
    #[serde(rename = "@backwards")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub backwards: Option<bool>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTTLSubShapeId {
    #[serde(rename = "@spid")]
    pub spid: String,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct CTTLTextTargetElement;

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTTLOleChartTargetElement {
    #[serde(rename = "@type")]
    pub r#type: STTLChartSubelementType,
    #[serde(rename = "@lvl")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub lvl: Option<u32>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTTLShapeTargetElement {
    #[serde(rename = "@spid")]
    pub spid: String,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct CTTLTimeTargetElement;

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTTLTriggerTimeNodeID {
    #[serde(rename = "@val")]
    pub value: STTLTimeNodeID,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTTLTriggerRuntimeNode {
    #[serde(rename = "@val")]
    pub value: STTLTriggerRuntimeNode,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTTLTimeCondition {
    #[serde(rename = "@evt")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub evt: Option<STTLTriggerEvent>,
    #[serde(rename = "@delay")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub delay: Option<STTLTime>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTTLTimeConditionList {
    #[serde(rename = "cond")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub cond: Vec<Box<CTTLTimeCondition>>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::RawXmlNode>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct CTTimeNodeList;

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTTLCommonTimeNodeData {
    #[serde(rename = "@id")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub id: Option<STTLTimeNodeID>,
    #[serde(rename = "@presetID")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub preset_i_d: Option<i32>,
    #[serde(rename = "@presetClass")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub preset_class: Option<STTLTimeNodePresetClassType>,
    #[serde(rename = "@presetSubtype")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub preset_subtype: Option<i32>,
    #[serde(rename = "@dur")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub dur: Option<STTLTime>,
    #[serde(rename = "@repeatCount")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub repeat_count: Option<STTLTime>,
    #[serde(rename = "@repeatDur")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub repeat_dur: Option<STTLTime>,
    #[serde(rename = "@spd")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub spd: Option<String>,
    #[serde(rename = "@accel")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub accel: Option<String>,
    #[serde(rename = "@decel")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub decel: Option<String>,
    #[serde(rename = "@autoRev")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub auto_rev: Option<bool>,
    #[serde(rename = "@restart")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub restart: Option<STTLTimeNodeRestartType>,
    #[serde(rename = "@fill")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub fill: Option<STTLTimeNodeFillType>,
    #[serde(rename = "@syncBehavior")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub sync_behavior: Option<STTLTimeNodeSyncType>,
    #[serde(rename = "@tmFilter")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub tm_filter: Option<String>,
    #[serde(rename = "@evtFilter")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub evt_filter: Option<String>,
    #[serde(rename = "@display")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub display: Option<bool>,
    #[serde(rename = "@masterRel")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub master_rel: Option<STTLTimeNodeMasterRelation>,
    #[serde(rename = "@bldLvl")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub bld_lvl: Option<i32>,
    #[serde(rename = "@grpId")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub grp_id: Option<u32>,
    #[serde(rename = "@afterEffect")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub after_effect: Option<bool>,
    #[serde(rename = "@nodeType")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub node_type: Option<STTLTimeNodeType>,
    #[serde(rename = "@nodePh")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub node_ph: Option<bool>,
    #[serde(rename = "stCondLst")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub st_cond_lst: Option<Box<CTTLTimeConditionList>>,
    #[serde(rename = "endCondLst")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub end_cond_lst: Option<Box<CTTLTimeConditionList>>,
    #[serde(rename = "endSync")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub end_sync: Option<Box<CTTLTimeCondition>>,
    #[serde(rename = "iterate")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub iterate: Option<Box<CTTLIterateData>>,
    #[serde(rename = "childTnLst")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub child_tn_lst: Option<Box<CTTimeNodeList>>,
    #[serde(rename = "subTnLst")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub sub_tn_lst: Option<Box<CTTimeNodeList>>,
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
    pub extra_children: Vec<ooxml_xml::RawXmlNode>,
}

pub type CTTLTimeNodeParallel = Box<CTTLCommonTimeNodeData>;

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTTLTimeNodeSequence {
    #[serde(rename = "@concurrent")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub concurrent: Option<bool>,
    #[serde(rename = "@prevAc")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub prev_ac: Option<STTLPreviousActionType>,
    #[serde(rename = "@nextAc")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub next_ac: Option<STTLNextActionType>,
    #[serde(rename = "cTn")]
    pub c_tn: Box<CTTLCommonTimeNodeData>,
    #[serde(rename = "prevCondLst")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub prev_cond_lst: Option<Box<CTTLTimeConditionList>>,
    #[serde(rename = "nextCondLst")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub next_cond_lst: Option<Box<CTTLTimeConditionList>>,
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
    pub extra_children: Vec<ooxml_xml::RawXmlNode>,
}

pub type CTTLTimeNodeExclusive = Box<CTTLCommonTimeNodeData>;

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTTLBehaviorAttributeNameList {
    #[serde(rename = "attrName")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub attr_name: Vec<String>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::RawXmlNode>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTTLCommonBehaviorData {
    #[serde(rename = "@additive")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub additive: Option<STTLBehaviorAdditiveType>,
    #[serde(rename = "@accumulate")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub accumulate: Option<STTLBehaviorAccumulateType>,
    #[serde(rename = "@xfrmType")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub xfrm_type: Option<STTLBehaviorTransformType>,
    #[serde(rename = "@from")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub from: Option<String>,
    #[serde(rename = "@to")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub to: Option<String>,
    #[serde(rename = "@by")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub by: Option<String>,
    #[serde(rename = "@rctx")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub rctx: Option<String>,
    #[serde(rename = "@override")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub r#override: Option<STTLBehaviorOverrideType>,
    #[serde(rename = "cTn")]
    pub c_tn: Box<CTTLCommonTimeNodeData>,
    #[serde(rename = "tgtEl")]
    pub tgt_el: Box<CTTLTimeTargetElement>,
    #[serde(rename = "attrNameLst")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub attr_name_lst: Option<Box<CTTLBehaviorAttributeNameList>>,
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
    pub extra_children: Vec<ooxml_xml::RawXmlNode>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTTLAnimVariantBooleanVal {
    #[serde(rename = "@val")]
    #[serde(with = "ooxml_xml::ooxml_bool_required")]
    pub value: bool,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTTLAnimVariantIntegerVal {
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
pub struct CTTLAnimVariantFloatVal {
    #[serde(rename = "@val")]
    pub value: f32,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTTLAnimVariantStringVal {
    #[serde(rename = "@val")]
    pub value: String,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct CTTLAnimVariant;

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTTLTimeAnimateValue {
    #[serde(rename = "@tm")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub tm: Option<STTLTimeAnimateValueTime>,
    #[serde(rename = "@fmla")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub fmla: Option<String>,
    #[serde(rename = "val")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub value: Option<Box<CTTLAnimVariant>>,
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
    pub extra_children: Vec<ooxml_xml::RawXmlNode>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTTLTimeAnimateValueList {
    #[serde(rename = "tav")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub tav: Vec<Box<CTTLTimeAnimateValue>>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::RawXmlNode>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTTLAnimateBehavior {
    #[serde(rename = "@by")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub by: Option<String>,
    #[serde(rename = "@from")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub from: Option<String>,
    #[serde(rename = "@to")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub to: Option<String>,
    #[serde(rename = "@calcmode")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub calcmode: Option<STTLAnimateBehaviorCalcMode>,
    #[serde(rename = "@valueType")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub value_type: Option<STTLAnimateBehaviorValueType>,
    #[serde(rename = "cBhvr")]
    pub c_bhvr: Box<CTTLCommonBehaviorData>,
    #[serde(rename = "tavLst")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub tav_lst: Option<Box<CTTLTimeAnimateValueList>>,
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
    pub extra_children: Vec<ooxml_xml::RawXmlNode>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTTLByRgbColorTransform {
    #[serde(rename = "@r")]
    pub reference: String,
    #[serde(rename = "@g")]
    pub g: String,
    #[serde(rename = "@b")]
    pub b: String,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTTLByHslColorTransform {
    #[serde(rename = "@h")]
    pub height: String,
    #[serde(rename = "@s")]
    pub s: String,
    #[serde(rename = "@l")]
    pub l: String,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct CTTLByAnimateColorTransform;

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTTLAnimateColorBehavior {
    #[serde(rename = "@clrSpc")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub clr_spc: Option<STTLAnimateColorSpace>,
    #[serde(rename = "@dir")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub dir: Option<STTLAnimateColorDirection>,
    #[serde(rename = "cBhvr")]
    pub c_bhvr: Box<CTTLCommonBehaviorData>,
    #[serde(rename = "by")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub by: Option<Box<CTTLByAnimateColorTransform>>,
    #[serde(rename = "from")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub from: Option<String>,
    #[serde(rename = "to")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub to: Option<String>,
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
    pub extra_children: Vec<ooxml_xml::RawXmlNode>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTTLAnimateEffectBehavior {
    #[serde(rename = "@transition")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub transition: Option<STTLAnimateEffectTransition>,
    #[serde(rename = "@filter")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub filter: Option<String>,
    #[serde(rename = "@prLst")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub pr_lst: Option<String>,
    #[serde(rename = "cBhvr")]
    pub c_bhvr: Box<CTTLCommonBehaviorData>,
    #[serde(rename = "progress")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub progress: Option<Box<CTTLAnimVariant>>,
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
    pub extra_children: Vec<ooxml_xml::RawXmlNode>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTTLPoint {
    #[serde(rename = "@x")]
    pub x: String,
    #[serde(rename = "@y")]
    pub y: String,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTTLAnimateMotionBehavior {
    #[serde(rename = "@origin")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub origin: Option<STTLAnimateMotionBehaviorOrigin>,
    #[serde(rename = "@path")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub path: Option<String>,
    #[serde(rename = "@pathEditMode")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub path_edit_mode: Option<STTLAnimateMotionPathEditMode>,
    #[serde(rename = "@rAng")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub r_ang: Option<String>,
    #[serde(rename = "@ptsTypes")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub pts_types: Option<String>,
    #[serde(rename = "cBhvr")]
    pub c_bhvr: Box<CTTLCommonBehaviorData>,
    #[serde(rename = "by")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub by: Option<Box<CTTLPoint>>,
    #[serde(rename = "from")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub from: Option<Box<CTTLPoint>>,
    #[serde(rename = "to")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub to: Option<Box<CTTLPoint>>,
    #[serde(rename = "rCtr")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub r_ctr: Option<Box<CTTLPoint>>,
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
    pub extra_children: Vec<ooxml_xml::RawXmlNode>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTTLAnimateRotationBehavior {
    #[serde(rename = "@by")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub by: Option<String>,
    #[serde(rename = "@from")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub from: Option<String>,
    #[serde(rename = "@to")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub to: Option<String>,
    #[serde(rename = "cBhvr")]
    pub c_bhvr: Box<CTTLCommonBehaviorData>,
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
    pub extra_children: Vec<ooxml_xml::RawXmlNode>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTTLAnimateScaleBehavior {
    #[serde(rename = "@zoomContents")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub zoom_contents: Option<bool>,
    #[serde(rename = "cBhvr")]
    pub c_bhvr: Box<CTTLCommonBehaviorData>,
    #[serde(rename = "by")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub by: Option<Box<CTTLPoint>>,
    #[serde(rename = "from")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub from: Option<Box<CTTLPoint>>,
    #[serde(rename = "to")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub to: Option<Box<CTTLPoint>>,
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
    pub extra_children: Vec<ooxml_xml::RawXmlNode>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTTLCommandBehavior {
    #[serde(rename = "@type")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub r#type: Option<STTLCommandType>,
    #[serde(rename = "@cmd")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub cmd: Option<String>,
    #[serde(rename = "cBhvr")]
    pub c_bhvr: Box<CTTLCommonBehaviorData>,
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
    pub extra_children: Vec<ooxml_xml::RawXmlNode>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTTLSetBehavior {
    #[serde(rename = "cBhvr")]
    pub c_bhvr: Box<CTTLCommonBehaviorData>,
    #[serde(rename = "to")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub to: Option<Box<CTTLAnimVariant>>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::RawXmlNode>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTTLCommonMediaNodeData {
    #[serde(rename = "@vol")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub vol: Option<String>,
    #[serde(rename = "@mute")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub mute: Option<bool>,
    #[serde(rename = "@numSld")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub num_sld: Option<u32>,
    #[serde(rename = "@showWhenStopped")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub show_when_stopped: Option<bool>,
    #[serde(rename = "cTn")]
    pub c_tn: Box<CTTLCommonTimeNodeData>,
    #[serde(rename = "tgtEl")]
    pub tgt_el: Box<CTTLTimeTargetElement>,
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
    pub extra_children: Vec<ooxml_xml::RawXmlNode>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTTLMediaNodeAudio {
    #[serde(rename = "@isNarration")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub is_narration: Option<bool>,
    #[serde(rename = "cMediaNode")]
    pub c_media_node: Box<CTTLCommonMediaNodeData>,
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
    pub extra_children: Vec<ooxml_xml::RawXmlNode>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTTLMediaNodeVideo {
    #[serde(rename = "@fullScrn")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub full_scrn: Option<bool>,
    #[serde(rename = "cMediaNode")]
    pub c_media_node: Box<CTTLCommonMediaNodeData>,
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
    pub extra_children: Vec<ooxml_xml::RawXmlNode>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct PAGTLBuild {
    #[serde(rename = "@spid")]
    pub spid: String,
    #[serde(rename = "@grpId")]
    pub grp_id: u32,
    #[serde(rename = "@uiExpand")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub ui_expand: Option<bool>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTTLTemplate {
    #[serde(rename = "@lvl")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub lvl: Option<u32>,
    #[serde(rename = "tnLst")]
    pub tn_lst: Box<CTTimeNodeList>,
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
    pub extra_children: Vec<ooxml_xml::RawXmlNode>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTTLTemplateList {
    #[serde(rename = "tmpl")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub tmpl: Vec<Box<CTTLTemplate>>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::RawXmlNode>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTTLBuildParagraph {
    #[serde(rename = "@build")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub build: Option<STTLParaBuildType>,
    #[serde(rename = "@bldLvl")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub bld_lvl: Option<u32>,
    #[serde(rename = "@animBg")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub anim_bg: Option<bool>,
    #[serde(rename = "@autoUpdateAnimBg")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub auto_update_anim_bg: Option<bool>,
    #[serde(rename = "@rev")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub rev: Option<bool>,
    #[serde(rename = "@advAuto")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub adv_auto: Option<STTLTime>,
    #[serde(rename = "tmplLst")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub tmpl_lst: Option<Box<CTTLTemplateList>>,
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
    pub extra_children: Vec<ooxml_xml::RawXmlNode>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTTLBuildDiagram {
    #[serde(rename = "@bld")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub bld: Option<STTLDiagramBuildType>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTTLOleBuildChart {
    #[serde(rename = "@bld")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub bld: Option<STTLOleChartBuildType>,
    #[serde(rename = "@animBg")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub anim_bg: Option<bool>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct CTTLGraphicalObjectBuild;

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct CTBuildList;

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct SlideTiming {
    #[serde(rename = "tnLst")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub tn_lst: Option<Box<CTTimeNodeList>>,
    #[serde(rename = "bldLst")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub bld_lst: Option<Box<CTBuildList>>,
    #[serde(rename = "extLst")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub ext_lst: Option<Box<CTExtensionListModify>>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::RawXmlNode>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct CTEmpty;

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTIndexRange {
    #[serde(rename = "@st")]
    pub st: STIndex,
    #[serde(rename = "@end")]
    pub end: STIndex,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct CTSlideRelationshipListEntry;

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTSlideRelationshipList {
    #[serde(rename = "sld")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub sld: Vec<Box<CTSlideRelationshipListEntry>>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::RawXmlNode>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTCustomShowId {
    #[serde(rename = "@id")]
    pub id: u32,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct CTCustomerData;

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct CTTagsData;

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTCustomerDataList {
    #[serde(rename = "custData")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub cust_data: Vec<Box<CTCustomerData>>,
    #[serde(rename = "tags")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub tags: Option<Box<CTTagsData>>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::RawXmlNode>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTExtension {
    #[serde(rename = "@uri")]
    pub uri: String,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

pub type CTExtensionAny = String;

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct EGExtensionList {
    #[serde(rename = "ext")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub ext: Vec<Box<CTExtension>>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::RawXmlNode>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct CTExtensionList;

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTExtensionListModify {
    #[serde(rename = "@mod")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub r#mod: Option<bool>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTCommentAuthor {
    #[serde(rename = "@id")]
    pub id: u32,
    #[serde(rename = "@name")]
    pub name: STName,
    #[serde(rename = "@initials")]
    pub initials: STName,
    #[serde(rename = "@lastIdx")]
    pub last_idx: u32,
    #[serde(rename = "@clrIdx")]
    pub clr_idx: u32,
    #[serde(rename = "extLst")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub ext_lst: Option<Box<CTExtensionList>>,
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
    pub extra_children: Vec<ooxml_xml::RawXmlNode>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTCommentAuthorList {
    #[serde(rename = "cmAuthor")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub cm_author: Vec<Box<CTCommentAuthor>>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::RawXmlNode>,
}

pub type PCmAuthorLst = Box<CTCommentAuthorList>;

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTComment {
    #[serde(rename = "@authorId")]
    pub author_id: u32,
    #[serde(rename = "@dt")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub dt: Option<String>,
    #[serde(rename = "@idx")]
    pub idx: STIndex,
    #[serde(rename = "pos")]
    pub pos: String,
    #[serde(rename = "text")]
    pub text: String,
    #[serde(rename = "extLst")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub ext_lst: Option<Box<CTExtensionListModify>>,
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
    pub extra_children: Vec<ooxml_xml::RawXmlNode>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTCommentList {
    #[serde(rename = "cm")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub cm: Vec<Box<CTComment>>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::RawXmlNode>,
}

pub type PCmLst = Box<CTCommentList>;

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct PAGOle {
    #[serde(rename = "@name")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub name: Option<String>,
    #[serde(rename = "@showAsIcon")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub show_as_icon: Option<bool>,
    #[serde(rename = "@imgW")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub img_w: Option<String>,
    #[serde(rename = "@imgH")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub img_h: Option<String>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTOleObjectEmbed {
    #[serde(rename = "@followColorScheme")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub follow_color_scheme: Option<STOleObjectFollowColorScheme>,
    #[serde(rename = "extLst")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub ext_lst: Option<Box<CTExtensionList>>,
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
    pub extra_children: Vec<ooxml_xml::RawXmlNode>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTOleObjectLink {
    #[serde(rename = "@updateAutomatic")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub update_automatic: Option<bool>,
    #[serde(rename = "extLst")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub ext_lst: Option<Box<CTExtensionList>>,
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
    pub extra_children: Vec<ooxml_xml::RawXmlNode>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTOleObject {
    #[serde(rename = "@progId")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub prog_id: Option<String>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

pub type POleObj = Box<CTOleObject>;

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTControl {
    #[serde(rename = "extLst")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub ext_lst: Option<Box<CTExtensionList>>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::RawXmlNode>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTControlList {
    #[serde(rename = "control")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub control: Vec<Box<CTControl>>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::RawXmlNode>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTSlideIdListEntry {
    #[serde(rename = "@id")]
    pub id: STSlideId,
    #[serde(rename = "extLst")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub ext_lst: Option<Box<CTExtensionList>>,
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
    pub extra_children: Vec<ooxml_xml::RawXmlNode>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct SlideIdList {
    #[serde(rename = "sldId")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub sld_id: Vec<Box<CTSlideIdListEntry>>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::RawXmlNode>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTSlideMasterIdListEntry {
    #[serde(rename = "@id")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub id: Option<STSlideMasterId>,
    #[serde(rename = "extLst")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub ext_lst: Option<Box<CTExtensionList>>,
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
    pub extra_children: Vec<ooxml_xml::RawXmlNode>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTSlideMasterIdList {
    #[serde(rename = "sldMasterId")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub sld_master_id: Vec<Box<CTSlideMasterIdListEntry>>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::RawXmlNode>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTNotesMasterIdListEntry {
    #[serde(rename = "extLst")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub ext_lst: Option<Box<CTExtensionList>>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::RawXmlNode>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTNotesMasterIdList {
    #[serde(rename = "notesMasterId")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub notes_master_id: Option<Box<CTNotesMasterIdListEntry>>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::RawXmlNode>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTHandoutMasterIdListEntry {
    #[serde(rename = "extLst")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub ext_lst: Option<Box<CTExtensionList>>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::RawXmlNode>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTHandoutMasterIdList {
    #[serde(rename = "handoutMasterId")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub handout_master_id: Option<Box<CTHandoutMasterIdListEntry>>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::RawXmlNode>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct CTEmbeddedFontDataId;

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTEmbeddedFontListEntry {
    #[serde(rename = "font")]
    pub font: String,
    #[serde(rename = "regular")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub regular: Option<Box<CTEmbeddedFontDataId>>,
    #[serde(rename = "bold")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub bold: Option<Box<CTEmbeddedFontDataId>>,
    #[serde(rename = "italic")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub italic: Option<Box<CTEmbeddedFontDataId>>,
    #[serde(rename = "boldItalic")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub bold_italic: Option<Box<CTEmbeddedFontDataId>>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::RawXmlNode>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTEmbeddedFontList {
    #[serde(rename = "embeddedFont")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub embedded_font: Vec<Box<CTEmbeddedFontListEntry>>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::RawXmlNode>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct CTSmartTags;

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTCustomShow {
    #[serde(rename = "@name")]
    pub name: STName,
    #[serde(rename = "@id")]
    pub id: u32,
    #[serde(rename = "sldLst")]
    pub sld_lst: Box<CTSlideRelationshipList>,
    #[serde(rename = "extLst")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub ext_lst: Option<Box<CTExtensionList>>,
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
    pub extra_children: Vec<ooxml_xml::RawXmlNode>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTCustomShowList {
    #[serde(rename = "custShow")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub cust_show: Vec<Box<CTCustomShow>>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::RawXmlNode>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTPhotoAlbum {
    #[serde(rename = "@bw")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub bw: Option<bool>,
    #[serde(rename = "@showCaptions")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub show_captions: Option<bool>,
    #[serde(rename = "@layout")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub layout: Option<STPhotoAlbumLayout>,
    #[serde(rename = "@frame")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub frame: Option<STPhotoAlbumFrameShape>,
    #[serde(rename = "extLst")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub ext_lst: Option<Box<CTExtensionList>>,
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
    pub extra_children: Vec<ooxml_xml::RawXmlNode>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTSlideSize {
    #[serde(rename = "@cx")]
    pub cx: STSlideSizeCoordinate,
    #[serde(rename = "@cy")]
    pub cy: STSlideSizeCoordinate,
    #[serde(rename = "@type")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub r#type: Option<STSlideSizeType>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTKinsoku {
    #[serde(rename = "@lang")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub lang: Option<String>,
    #[serde(rename = "@invalStChars")]
    pub inval_st_chars: String,
    #[serde(rename = "@invalEndChars")]
    pub inval_end_chars: String,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTModifyVerifier {
    #[serde(rename = "@algorithmName")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub algorithm_name: Option<String>,
    #[serde(rename = "@hashValue")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub hash_value: Option<Vec<u8>>,
    #[serde(rename = "@saltValue")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub salt_value: Option<Vec<u8>>,
    #[serde(rename = "@spinValue")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub spin_value: Option<u32>,
    #[serde(rename = "@cryptProviderType")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub crypt_provider_type: Option<STCryptProv>,
    #[serde(rename = "@cryptAlgorithmClass")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub crypt_algorithm_class: Option<STAlgClass>,
    #[serde(rename = "@cryptAlgorithmType")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub crypt_algorithm_type: Option<STAlgType>,
    #[serde(rename = "@cryptAlgorithmSid")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub crypt_algorithm_sid: Option<u32>,
    #[serde(rename = "@spinCount")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub spin_count: Option<u32>,
    #[serde(rename = "@saltData")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub salt_data: Option<Vec<u8>>,
    #[serde(rename = "@hashData")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub hash_data: Option<Vec<u8>>,
    #[serde(rename = "@cryptProvider")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub crypt_provider: Option<String>,
    #[serde(rename = "@algIdExt")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub alg_id_ext: Option<u32>,
    #[serde(rename = "@algIdExtSource")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub alg_id_ext_source: Option<String>,
    #[serde(rename = "@cryptProviderTypeExt")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub crypt_provider_type_ext: Option<u32>,
    #[serde(rename = "@cryptProviderTypeExtSource")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub crypt_provider_type_ext_source: Option<String>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct Presentation {
    #[serde(rename = "@serverZoom")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub server_zoom: Option<String>,
    #[serde(rename = "@firstSlideNum")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub first_slide_num: Option<i32>,
    #[serde(rename = "@showSpecialPlsOnTitleSld")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub show_special_pls_on_title_sld: Option<bool>,
    #[serde(rename = "@rtl")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub rtl: Option<bool>,
    #[serde(rename = "@removePersonalInfoOnSave")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub remove_personal_info_on_save: Option<bool>,
    #[serde(rename = "@compatMode")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub compat_mode: Option<bool>,
    #[serde(rename = "@strictFirstAndLastChars")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub strict_first_and_last_chars: Option<bool>,
    #[serde(rename = "@embedTrueTypeFonts")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub embed_true_type_fonts: Option<bool>,
    #[serde(rename = "@saveSubsetFonts")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub save_subset_fonts: Option<bool>,
    #[serde(rename = "@autoCompressPictures")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub auto_compress_pictures: Option<bool>,
    #[serde(rename = "@bookmarkIdSeed")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub bookmark_id_seed: Option<STBookmarkIdSeed>,
    #[serde(rename = "@conformance")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub conformance: Option<STConformanceClass>,
    #[serde(rename = "sldMasterIdLst")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub sld_master_id_lst: Option<Box<CTSlideMasterIdList>>,
    #[serde(rename = "notesMasterIdLst")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub notes_master_id_lst: Option<Box<CTNotesMasterIdList>>,
    #[serde(rename = "handoutMasterIdLst")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub handout_master_id_lst: Option<Box<CTHandoutMasterIdList>>,
    #[serde(rename = "sldIdLst")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub sld_id_lst: Option<Box<SlideIdList>>,
    #[serde(rename = "sldSz")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub sld_sz: Option<Box<CTSlideSize>>,
    #[serde(rename = "notesSz")]
    pub notes_sz: String,
    #[serde(rename = "smartTags")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub smart_tags: Option<Box<CTSmartTags>>,
    #[serde(rename = "embeddedFontLst")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub embedded_font_lst: Option<Box<CTEmbeddedFontList>>,
    #[serde(rename = "custShowLst")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub cust_show_lst: Option<Box<CTCustomShowList>>,
    #[serde(rename = "photoAlbum")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub photo_album: Option<Box<CTPhotoAlbum>>,
    #[serde(rename = "custDataLst")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub cust_data_lst: Option<Box<CTCustomerDataList>>,
    #[serde(rename = "kinsoku")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub kinsoku: Option<Box<CTKinsoku>>,
    #[serde(rename = "defaultTextStyle")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub default_text_style: Option<String>,
    #[serde(rename = "modifyVerifier")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub modify_verifier: Option<Box<CTModifyVerifier>>,
    #[serde(rename = "extLst")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub ext_lst: Option<Box<CTExtensionList>>,
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
    pub extra_children: Vec<ooxml_xml::RawXmlNode>,
}

pub type PPresentation = Box<Presentation>;

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTHtmlPublishProperties {
    #[serde(rename = "@showSpeakerNotes")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub show_speaker_notes: Option<bool>,
    #[serde(rename = "@target")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub target: Option<String>,
    #[serde(rename = "@title")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub title: Option<String>,
    #[serde(rename = "extLst")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub ext_lst: Option<Box<CTExtensionList>>,
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
    pub extra_children: Vec<ooxml_xml::RawXmlNode>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTWebProperties {
    #[serde(rename = "@showAnimation")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub show_animation: Option<bool>,
    #[serde(rename = "@resizeGraphics")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub resize_graphics: Option<bool>,
    #[serde(rename = "@allowPng")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub allow_png: Option<bool>,
    #[serde(rename = "@relyOnVml")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub rely_on_vml: Option<bool>,
    #[serde(rename = "@organizeInFolders")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub organize_in_folders: Option<bool>,
    #[serde(rename = "@useLongFilenames")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub use_long_filenames: Option<bool>,
    #[serde(rename = "@imgSz")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub img_sz: Option<STWebScreenSize>,
    #[serde(rename = "@encoding")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub encoding: Option<STWebEncoding>,
    #[serde(rename = "@clr")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub clr: Option<STWebColorType>,
    #[serde(rename = "extLst")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub ext_lst: Option<Box<CTExtensionList>>,
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
    pub extra_children: Vec<ooxml_xml::RawXmlNode>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTPrintProperties {
    #[serde(rename = "@prnWhat")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub prn_what: Option<STPrintWhat>,
    #[serde(rename = "@clrMode")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub clr_mode: Option<STPrintColorMode>,
    #[serde(rename = "@hiddenSlides")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub hidden_slides: Option<bool>,
    #[serde(rename = "@scaleToFitPaper")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub scale_to_fit_paper: Option<bool>,
    #[serde(rename = "@frameSlides")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub frame_slides: Option<bool>,
    #[serde(rename = "extLst")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub ext_lst: Option<Box<CTExtensionList>>,
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
    pub extra_children: Vec<ooxml_xml::RawXmlNode>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTShowInfoBrowse {
    #[serde(rename = "@showScrollbar")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub show_scrollbar: Option<bool>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTShowInfoKiosk {
    #[serde(rename = "@restart")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub restart: Option<u32>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTShowProperties {
    #[serde(rename = "@loop")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub r#loop: Option<bool>,
    #[serde(rename = "@showNarration")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub show_narration: Option<bool>,
    #[serde(rename = "@showAnimation")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub show_animation: Option<bool>,
    #[serde(rename = "@useTimings")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub use_timings: Option<bool>,
    #[serde(rename = "penClr")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub pen_clr: Option<String>,
    #[serde(rename = "extLst")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub ext_lst: Option<Box<CTExtensionList>>,
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
    pub extra_children: Vec<ooxml_xml::RawXmlNode>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTPresentationProperties {
    #[serde(rename = "htmlPubPr")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub html_pub_pr: Option<Box<CTHtmlPublishProperties>>,
    #[serde(rename = "webPr")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub web_pr: Option<Box<CTWebProperties>>,
    #[serde(rename = "prnPr")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub prn_pr: Option<Box<CTPrintProperties>>,
    #[serde(rename = "showPr")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub show_pr: Option<Box<CTShowProperties>>,
    #[serde(rename = "clrMru")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub clr_mru: Option<String>,
    #[serde(rename = "extLst")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub ext_lst: Option<Box<CTExtensionList>>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::RawXmlNode>,
}

pub type PPresentationPr = Box<CTPresentationProperties>;

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTHeaderFooter {
    #[serde(rename = "@sldNum")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub sld_num: Option<bool>,
    #[serde(rename = "@hdr")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub hdr: Option<bool>,
    #[serde(rename = "@ftr")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub ftr: Option<bool>,
    #[serde(rename = "@dt")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub dt: Option<bool>,
    #[serde(rename = "extLst")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub ext_lst: Option<Box<CTExtensionListModify>>,
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
    pub extra_children: Vec<ooxml_xml::RawXmlNode>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTPlaceholder {
    #[serde(rename = "@type")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub r#type: Option<STPlaceholderType>,
    #[serde(rename = "@orient")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub orient: Option<STDirection>,
    #[serde(rename = "@sz")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub sz: Option<STPlaceholderSize>,
    #[serde(rename = "@idx")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub idx: Option<u32>,
    #[serde(rename = "@hasCustomPrompt")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub has_custom_prompt: Option<bool>,
    #[serde(rename = "extLst")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub ext_lst: Option<Box<CTExtensionListModify>>,
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
    pub extra_children: Vec<ooxml_xml::RawXmlNode>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTApplicationNonVisualDrawingProps {
    #[serde(rename = "@isPhoto")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub is_photo: Option<bool>,
    #[serde(rename = "@userDrawn")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub user_drawn: Option<bool>,
    #[serde(rename = "ph")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub ph: Option<Box<CTPlaceholder>>,
    #[serde(rename = "custDataLst")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub cust_data_lst: Option<Box<CTCustomerDataList>>,
    #[serde(rename = "extLst")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub ext_lst: Option<Box<CTExtensionList>>,
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
    pub extra_children: Vec<ooxml_xml::RawXmlNode>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct ShapeNonVisual {
    #[serde(rename = "cNvPr")]
    pub c_nv_pr: String,
    #[serde(rename = "cNvSpPr")]
    pub c_nv_sp_pr: String,
    #[serde(rename = "nvPr")]
    pub nv_pr: Box<CTApplicationNonVisualDrawingProps>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::RawXmlNode>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct Shape {
    #[serde(rename = "@useBgFill")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub use_bg_fill: Option<bool>,
    #[serde(rename = "nvSpPr")]
    pub non_visual_properties: Box<ShapeNonVisual>,
    #[serde(rename = "spPr")]
    pub shape_properties: String,
    #[serde(rename = "style")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub style: Option<String>,
    #[serde(rename = "txBody")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub text_body: Option<String>,
    #[serde(rename = "extLst")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub ext_lst: Option<Box<CTExtensionListModify>>,
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
    pub extra_children: Vec<ooxml_xml::RawXmlNode>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTConnectorNonVisual {
    #[serde(rename = "cNvPr")]
    pub c_nv_pr: String,
    #[serde(rename = "cNvCxnSpPr")]
    pub c_nv_cxn_sp_pr: String,
    #[serde(rename = "nvPr")]
    pub nv_pr: Box<CTApplicationNonVisualDrawingProps>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::RawXmlNode>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct Connector {
    #[serde(rename = "nvCxnSpPr")]
    pub non_visual_connector_properties: Box<CTConnectorNonVisual>,
    #[serde(rename = "spPr")]
    pub shape_properties: String,
    #[serde(rename = "style")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub style: Option<String>,
    #[serde(rename = "extLst")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub ext_lst: Option<Box<CTExtensionListModify>>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::RawXmlNode>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTPictureNonVisual {
    #[serde(rename = "cNvPr")]
    pub c_nv_pr: String,
    #[serde(rename = "cNvPicPr")]
    pub c_nv_pic_pr: String,
    #[serde(rename = "nvPr")]
    pub nv_pr: Box<CTApplicationNonVisualDrawingProps>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::RawXmlNode>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct Picture {
    #[serde(rename = "nvPicPr")]
    pub non_visual_picture_properties: Box<CTPictureNonVisual>,
    #[serde(rename = "blipFill")]
    pub blip_fill: String,
    #[serde(rename = "spPr")]
    pub shape_properties: String,
    #[serde(rename = "style")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub style: Option<String>,
    #[serde(rename = "extLst")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub ext_lst: Option<Box<CTExtensionListModify>>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::RawXmlNode>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTGraphicalObjectFrameNonVisual {
    #[serde(rename = "cNvPr")]
    pub c_nv_pr: String,
    #[serde(rename = "cNvGraphicFramePr")]
    pub c_nv_graphic_frame_pr: String,
    #[serde(rename = "nvPr")]
    pub nv_pr: Box<CTApplicationNonVisualDrawingProps>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::RawXmlNode>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct GraphicalObjectFrame {
    #[serde(rename = "@bwMode")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub bw_mode: Option<String>,
    #[serde(rename = "nvGraphicFramePr")]
    pub nv_graphic_frame_pr: Box<CTGraphicalObjectFrameNonVisual>,
    #[serde(rename = "xfrm")]
    pub xfrm: String,
    #[serde(rename = "extLst")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub ext_lst: Option<Box<CTExtensionListModify>>,
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
    pub extra_children: Vec<ooxml_xml::RawXmlNode>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTGroupShapeNonVisual {
    #[serde(rename = "cNvPr")]
    pub c_nv_pr: String,
    #[serde(rename = "cNvGrpSpPr")]
    pub c_nv_grp_sp_pr: String,
    #[serde(rename = "nvPr")]
    pub nv_pr: Box<CTApplicationNonVisualDrawingProps>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::RawXmlNode>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct GroupShape {
    #[serde(rename = "nvGrpSpPr")]
    pub non_visual_group_properties: Box<CTGroupShapeNonVisual>,
    #[serde(rename = "grpSpPr")]
    pub grp_sp_pr: String,
    #[serde(rename = "extLst")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub ext_lst: Option<Box<CTExtensionListModify>>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::RawXmlNode>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct CTRel;

pub type EGTopLevelSlide = String;

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct EGChildSlide {
    #[serde(rename = "clrMapOvr")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub clr_map_ovr: Option<String>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::RawXmlNode>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct PAGChildSlide {
    #[serde(rename = "@showMasterSp")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub show_master_sp: Option<bool>,
    #[serde(rename = "@showMasterPhAnim")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub show_master_ph_anim: Option<bool>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTBackgroundProperties {
    #[serde(rename = "@shadeToTitle")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub shade_to_title: Option<bool>,
    #[serde(rename = "extLst")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub ext_lst: Option<Box<CTExtensionList>>,
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
    pub extra_children: Vec<ooxml_xml::RawXmlNode>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTBackground {
    #[serde(rename = "@bwMode")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub bw_mode: Option<String>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CommonSlideData {
    #[serde(rename = "@name")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub name: Option<String>,
    #[serde(rename = "bg")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub bg: Option<Box<CTBackground>>,
    #[serde(rename = "spTree")]
    pub shape_tree: Box<GroupShape>,
    #[serde(rename = "custDataLst")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub cust_data_lst: Option<Box<CTCustomerDataList>>,
    #[serde(rename = "controls")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub controls: Option<Box<CTControlList>>,
    #[serde(rename = "extLst")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub ext_lst: Option<Box<CTExtensionList>>,
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
    pub extra_children: Vec<ooxml_xml::RawXmlNode>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct Slide {
    #[serde(rename = "@show")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub show: Option<bool>,
    #[serde(rename = "cSld")]
    pub common_slide_data: Box<CommonSlideData>,
    #[cfg(feature = "pml-transitions")]
    #[serde(rename = "transition")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub transition: Option<Box<SlideTransition>>,
    #[cfg(feature = "pml-animations")]
    #[serde(rename = "timing")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub timing: Option<Box<SlideTiming>>,
    #[serde(rename = "extLst")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub ext_lst: Option<Box<CTExtensionListModify>>,
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
    pub extra_children: Vec<ooxml_xml::RawXmlNode>,
}

pub type PSld = Box<Slide>;

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct SlideLayout {
    #[serde(rename = "@matchingName")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub matching_name: Option<String>,
    #[serde(rename = "@type")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub r#type: Option<STSlideLayoutType>,
    #[serde(rename = "@preserve")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub preserve: Option<bool>,
    #[serde(rename = "@userDrawn")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub user_drawn: Option<bool>,
    #[serde(rename = "cSld")]
    pub common_slide_data: Box<CommonSlideData>,
    #[serde(rename = "transition")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub transition: Option<Box<SlideTransition>>,
    #[serde(rename = "timing")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub timing: Option<Box<SlideTiming>>,
    #[serde(rename = "hf")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub hf: Option<Box<CTHeaderFooter>>,
    #[serde(rename = "extLst")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub ext_lst: Option<Box<CTExtensionListModify>>,
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
    pub extra_children: Vec<ooxml_xml::RawXmlNode>,
}

pub type PSldLayout = Box<SlideLayout>;

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTSlideMasterTextStyles {
    #[serde(rename = "titleStyle")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub title_style: Option<String>,
    #[serde(rename = "bodyStyle")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub body_style: Option<String>,
    #[serde(rename = "otherStyle")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub other_style: Option<String>,
    #[serde(rename = "extLst")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub ext_lst: Option<Box<CTExtensionList>>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::RawXmlNode>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTSlideLayoutIdListEntry {
    #[serde(rename = "@id")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub id: Option<STSlideLayoutId>,
    #[serde(rename = "extLst")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub ext_lst: Option<Box<CTExtensionList>>,
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
    pub extra_children: Vec<ooxml_xml::RawXmlNode>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTSlideLayoutIdList {
    #[serde(rename = "sldLayoutId")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub sld_layout_id: Vec<Box<CTSlideLayoutIdListEntry>>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::RawXmlNode>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct SlideMaster {
    #[serde(rename = "@preserve")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub preserve: Option<bool>,
    #[serde(rename = "cSld")]
    pub common_slide_data: Box<CommonSlideData>,
    #[serde(rename = "sldLayoutIdLst")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub sld_layout_id_lst: Option<Box<CTSlideLayoutIdList>>,
    #[serde(rename = "transition")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub transition: Option<Box<SlideTransition>>,
    #[serde(rename = "timing")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub timing: Option<Box<SlideTiming>>,
    #[serde(rename = "hf")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub hf: Option<Box<CTHeaderFooter>>,
    #[serde(rename = "txStyles")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub tx_styles: Option<Box<CTSlideMasterTextStyles>>,
    #[serde(rename = "extLst")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub ext_lst: Option<Box<CTExtensionListModify>>,
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
    pub extra_children: Vec<ooxml_xml::RawXmlNode>,
}

pub type PSldMaster = Box<SlideMaster>;

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct HandoutMaster {
    #[serde(rename = "cSld")]
    pub common_slide_data: Box<CommonSlideData>,
    #[serde(rename = "hf")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub hf: Option<Box<CTHeaderFooter>>,
    #[serde(rename = "extLst")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub ext_lst: Option<Box<CTExtensionListModify>>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::RawXmlNode>,
}

pub type PHandoutMaster = Box<HandoutMaster>;

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct NotesMaster {
    #[serde(rename = "cSld")]
    pub common_slide_data: Box<CommonSlideData>,
    #[serde(rename = "hf")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub hf: Option<Box<CTHeaderFooter>>,
    #[serde(rename = "notesStyle")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub notes_style: Option<String>,
    #[serde(rename = "extLst")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub ext_lst: Option<Box<CTExtensionListModify>>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::RawXmlNode>,
}

pub type PNotesMaster = Box<NotesMaster>;

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct NotesSlide {
    #[serde(rename = "cSld")]
    pub common_slide_data: Box<CommonSlideData>,
    #[serde(rename = "extLst")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub ext_lst: Option<Box<CTExtensionListModify>>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::RawXmlNode>,
}

pub type PNotes = Box<NotesSlide>;

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTSlideSyncProperties {
    #[serde(rename = "@serverSldId")]
    pub server_sld_id: String,
    #[serde(rename = "@serverSldModifiedTime")]
    pub server_sld_modified_time: String,
    #[serde(rename = "@clientInsertedTime")]
    pub client_inserted_time: String,
    #[serde(rename = "extLst")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub ext_lst: Option<Box<CTExtensionList>>,
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
    pub extra_children: Vec<ooxml_xml::RawXmlNode>,
}

pub type PSldSyncPr = Box<CTSlideSyncProperties>;

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTStringTag {
    #[serde(rename = "@name")]
    pub name: String,
    #[serde(rename = "@val")]
    pub value: String,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTTagList {
    #[serde(rename = "tag")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub tag: Vec<Box<CTStringTag>>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::RawXmlNode>,
}

pub type PTagLst = Box<CTTagList>;

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTNormalViewPortion {
    #[serde(rename = "@sz")]
    pub sz: String,
    #[serde(rename = "@autoAdjust")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub auto_adjust: Option<bool>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTNormalViewProperties {
    #[serde(rename = "@showOutlineIcons")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub show_outline_icons: Option<bool>,
    #[serde(rename = "@snapVertSplitter")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub snap_vert_splitter: Option<bool>,
    #[serde(rename = "@vertBarState")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub vert_bar_state: Option<STSplitterBarState>,
    #[serde(rename = "@horzBarState")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub horz_bar_state: Option<STSplitterBarState>,
    #[serde(rename = "@preferSingleView")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub prefer_single_view: Option<bool>,
    #[serde(rename = "restoredLeft")]
    pub restored_left: Box<CTNormalViewPortion>,
    #[serde(rename = "restoredTop")]
    pub restored_top: Box<CTNormalViewPortion>,
    #[serde(rename = "extLst")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub ext_lst: Option<Box<CTExtensionList>>,
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
    pub extra_children: Vec<ooxml_xml::RawXmlNode>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTCommonViewProperties {
    #[serde(rename = "@varScale")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub var_scale: Option<bool>,
    #[serde(rename = "scale")]
    pub scale: String,
    #[serde(rename = "origin")]
    pub origin: String,
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
    pub extra_children: Vec<ooxml_xml::RawXmlNode>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTNotesTextViewProperties {
    #[serde(rename = "cViewPr")]
    pub c_view_pr: Box<CTCommonViewProperties>,
    #[serde(rename = "extLst")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub ext_lst: Option<Box<CTExtensionList>>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::RawXmlNode>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTOutlineViewSlideEntry {
    #[serde(rename = "@collapse")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub collapse: Option<bool>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTOutlineViewSlideList {
    #[serde(rename = "sld")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub sld: Vec<Box<CTOutlineViewSlideEntry>>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::RawXmlNode>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTOutlineViewProperties {
    #[serde(rename = "cViewPr")]
    pub c_view_pr: Box<CTCommonViewProperties>,
    #[serde(rename = "sldLst")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub sld_lst: Option<Box<CTOutlineViewSlideList>>,
    #[serde(rename = "extLst")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub ext_lst: Option<Box<CTExtensionList>>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::RawXmlNode>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTSlideSorterViewProperties {
    #[serde(rename = "@showFormatting")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub show_formatting: Option<bool>,
    #[serde(rename = "cViewPr")]
    pub c_view_pr: Box<CTCommonViewProperties>,
    #[serde(rename = "extLst")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub ext_lst: Option<Box<CTExtensionList>>,
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
    pub extra_children: Vec<ooxml_xml::RawXmlNode>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTGuide {
    #[serde(rename = "@orient")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub orient: Option<STDirection>,
    #[serde(rename = "@pos")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub pos: Option<String>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTGuideList {
    #[serde(rename = "guide")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub guide: Vec<Box<CTGuide>>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::RawXmlNode>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTCommonSlideViewProperties {
    #[serde(rename = "@snapToGrid")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub snap_to_grid: Option<bool>,
    #[serde(rename = "@snapToObjects")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub snap_to_objects: Option<bool>,
    #[serde(rename = "@showGuides")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub show_guides: Option<bool>,
    #[serde(rename = "cViewPr")]
    pub c_view_pr: Box<CTCommonViewProperties>,
    #[serde(rename = "guideLst")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub guide_lst: Option<Box<CTGuideList>>,
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
    pub extra_children: Vec<ooxml_xml::RawXmlNode>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTSlideViewProperties {
    #[serde(rename = "cSldViewPr")]
    pub c_sld_view_pr: Box<CTCommonSlideViewProperties>,
    #[serde(rename = "extLst")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub ext_lst: Option<Box<CTExtensionList>>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::RawXmlNode>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTNotesViewProperties {
    #[serde(rename = "cSldViewPr")]
    pub c_sld_view_pr: Box<CTCommonSlideViewProperties>,
    #[serde(rename = "extLst")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub ext_lst: Option<Box<CTExtensionList>>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::RawXmlNode>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTViewProperties {
    #[serde(rename = "@lastView")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub last_view: Option<STViewType>,
    #[serde(rename = "@showComments")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub show_comments: Option<bool>,
    #[serde(rename = "normalViewPr")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub normal_view_pr: Option<Box<CTNormalViewProperties>>,
    #[serde(rename = "slideViewPr")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub slide_view_pr: Option<Box<CTSlideViewProperties>>,
    #[serde(rename = "outlineViewPr")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub outline_view_pr: Option<Box<CTOutlineViewProperties>>,
    #[serde(rename = "notesTextViewPr")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub notes_text_view_pr: Option<Box<CTNotesTextViewProperties>>,
    #[serde(rename = "sorterViewPr")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub sorter_view_pr: Option<Box<CTSlideSorterViewProperties>>,
    #[serde(rename = "notesViewPr")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub notes_view_pr: Option<Box<CTNotesViewProperties>>,
    #[serde(rename = "gridSpacing")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub grid_spacing: Option<String>,
    #[serde(rename = "extLst")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub ext_lst: Option<Box<CTExtensionList>>,
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
    pub extra_children: Vec<ooxml_xml::RawXmlNode>,
}

pub type PViewPr = Box<CTViewProperties>;
