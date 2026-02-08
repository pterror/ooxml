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
    /// Namespace prefix: m
    pub const M: &str = "http://schemas.openxmlformats.org/officeDocument/2006/math";
    /// Namespace prefix: sl
    pub const SL: &str = "http://schemas.openxmlformats.org/schemaLibrary/2006/main";
    /// Default namespace (prefix: w)
    pub const W: &str = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    /// Namespace prefix: wp
    pub const WP: &str = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing";
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

pub type STRelationshipId = String;

pub type STLongHexNumber = Vec<u8>;

pub type STShortHexNumber = Vec<u8>;

pub type STUcharHexNumber = Vec<u8>;

pub type STDecimalNumberOrPercent = String;

pub type STUnqualifiedPercentage = i64;

pub type STDecimalNumber = i64;

pub type STSignedTwipsMeasure = String;

pub type STPixelsMeasure = STUnsignedDecimalNumber;

pub type STHpsMeasure = String;

pub type STSignedHpsMeasure = String;

pub type STDateTime = String;

pub type STMacroName = String;

pub type STEighthPointMeasure = STUnsignedDecimalNumber;

pub type STPointMeasure = STUnsignedDecimalNumber;

pub type STTextScale = String;

pub type STTextScalePercent = String;

pub type STTextScaleDecimal = i64;

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum STHighlightColor {
    #[serde(rename = "black")]
    Black,
    #[serde(rename = "blue")]
    Blue,
    #[serde(rename = "cyan")]
    Cyan,
    #[serde(rename = "green")]
    Green,
    #[serde(rename = "magenta")]
    Magenta,
    #[serde(rename = "red")]
    Red,
    #[serde(rename = "yellow")]
    Yellow,
    #[serde(rename = "white")]
    White,
    #[serde(rename = "darkBlue")]
    DarkBlue,
    #[serde(rename = "darkCyan")]
    DarkCyan,
    #[serde(rename = "darkGreen")]
    DarkGreen,
    #[serde(rename = "darkMagenta")]
    DarkMagenta,
    #[serde(rename = "darkRed")]
    DarkRed,
    #[serde(rename = "darkYellow")]
    DarkYellow,
    #[serde(rename = "darkGray")]
    DarkGray,
    #[serde(rename = "lightGray")]
    LightGray,
    #[serde(rename = "none")]
    None,
}

impl std::fmt::Display for STHighlightColor {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::Black => write!(f, "black"),
            Self::Blue => write!(f, "blue"),
            Self::Cyan => write!(f, "cyan"),
            Self::Green => write!(f, "green"),
            Self::Magenta => write!(f, "magenta"),
            Self::Red => write!(f, "red"),
            Self::Yellow => write!(f, "yellow"),
            Self::White => write!(f, "white"),
            Self::DarkBlue => write!(f, "darkBlue"),
            Self::DarkCyan => write!(f, "darkCyan"),
            Self::DarkGreen => write!(f, "darkGreen"),
            Self::DarkMagenta => write!(f, "darkMagenta"),
            Self::DarkRed => write!(f, "darkRed"),
            Self::DarkYellow => write!(f, "darkYellow"),
            Self::DarkGray => write!(f, "darkGray"),
            Self::LightGray => write!(f, "lightGray"),
            Self::None => write!(f, "none"),
        }
    }
}

impl std::str::FromStr for STHighlightColor {
    type Err = String;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        match s {
            "black" => Ok(Self::Black),
            "blue" => Ok(Self::Blue),
            "cyan" => Ok(Self::Cyan),
            "green" => Ok(Self::Green),
            "magenta" => Ok(Self::Magenta),
            "red" => Ok(Self::Red),
            "yellow" => Ok(Self::Yellow),
            "white" => Ok(Self::White),
            "darkBlue" => Ok(Self::DarkBlue),
            "darkCyan" => Ok(Self::DarkCyan),
            "darkGreen" => Ok(Self::DarkGreen),
            "darkMagenta" => Ok(Self::DarkMagenta),
            "darkRed" => Ok(Self::DarkRed),
            "darkYellow" => Ok(Self::DarkYellow),
            "darkGray" => Ok(Self::DarkGray),
            "lightGray" => Ok(Self::LightGray),
            "none" => Ok(Self::None),
            _ => Err(format!("unknown STHighlightColor value: {}", s)),
        }
    }
}

pub type STHexColor = String;

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum STUnderline {
    #[serde(rename = "single")]
    Single,
    #[serde(rename = "words")]
    Words,
    #[serde(rename = "double")]
    Double,
    #[serde(rename = "thick")]
    Thick,
    #[serde(rename = "dotted")]
    Dotted,
    #[serde(rename = "dottedHeavy")]
    DottedHeavy,
    #[serde(rename = "dash")]
    Dash,
    #[serde(rename = "dashedHeavy")]
    DashedHeavy,
    #[serde(rename = "dashLong")]
    DashLong,
    #[serde(rename = "dashLongHeavy")]
    DashLongHeavy,
    #[serde(rename = "dotDash")]
    DotDash,
    #[serde(rename = "dashDotHeavy")]
    DashDotHeavy,
    #[serde(rename = "dotDotDash")]
    DotDotDash,
    #[serde(rename = "dashDotDotHeavy")]
    DashDotDotHeavy,
    #[serde(rename = "wave")]
    Wave,
    #[serde(rename = "wavyHeavy")]
    WavyHeavy,
    #[serde(rename = "wavyDouble")]
    WavyDouble,
    #[serde(rename = "none")]
    None,
}

impl std::fmt::Display for STUnderline {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::Single => write!(f, "single"),
            Self::Words => write!(f, "words"),
            Self::Double => write!(f, "double"),
            Self::Thick => write!(f, "thick"),
            Self::Dotted => write!(f, "dotted"),
            Self::DottedHeavy => write!(f, "dottedHeavy"),
            Self::Dash => write!(f, "dash"),
            Self::DashedHeavy => write!(f, "dashedHeavy"),
            Self::DashLong => write!(f, "dashLong"),
            Self::DashLongHeavy => write!(f, "dashLongHeavy"),
            Self::DotDash => write!(f, "dotDash"),
            Self::DashDotHeavy => write!(f, "dashDotHeavy"),
            Self::DotDotDash => write!(f, "dotDotDash"),
            Self::DashDotDotHeavy => write!(f, "dashDotDotHeavy"),
            Self::Wave => write!(f, "wave"),
            Self::WavyHeavy => write!(f, "wavyHeavy"),
            Self::WavyDouble => write!(f, "wavyDouble"),
            Self::None => write!(f, "none"),
        }
    }
}

impl std::str::FromStr for STUnderline {
    type Err = String;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        match s {
            "single" => Ok(Self::Single),
            "words" => Ok(Self::Words),
            "double" => Ok(Self::Double),
            "thick" => Ok(Self::Thick),
            "dotted" => Ok(Self::Dotted),
            "dottedHeavy" => Ok(Self::DottedHeavy),
            "dash" => Ok(Self::Dash),
            "dashedHeavy" => Ok(Self::DashedHeavy),
            "dashLong" => Ok(Self::DashLong),
            "dashLongHeavy" => Ok(Self::DashLongHeavy),
            "dotDash" => Ok(Self::DotDash),
            "dashDotHeavy" => Ok(Self::DashDotHeavy),
            "dotDotDash" => Ok(Self::DotDotDash),
            "dashDotDotHeavy" => Ok(Self::DashDotDotHeavy),
            "wave" => Ok(Self::Wave),
            "wavyHeavy" => Ok(Self::WavyHeavy),
            "wavyDouble" => Ok(Self::WavyDouble),
            "none" => Ok(Self::None),
            _ => Err(format!("unknown STUnderline value: {}", s)),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum STTextEffect {
    #[serde(rename = "blinkBackground")]
    BlinkBackground,
    #[serde(rename = "lights")]
    Lights,
    #[serde(rename = "antsBlack")]
    AntsBlack,
    #[serde(rename = "antsRed")]
    AntsRed,
    #[serde(rename = "shimmer")]
    Shimmer,
    #[serde(rename = "sparkle")]
    Sparkle,
    #[serde(rename = "none")]
    None,
}

impl std::fmt::Display for STTextEffect {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::BlinkBackground => write!(f, "blinkBackground"),
            Self::Lights => write!(f, "lights"),
            Self::AntsBlack => write!(f, "antsBlack"),
            Self::AntsRed => write!(f, "antsRed"),
            Self::Shimmer => write!(f, "shimmer"),
            Self::Sparkle => write!(f, "sparkle"),
            Self::None => write!(f, "none"),
        }
    }
}

impl std::str::FromStr for STTextEffect {
    type Err = String;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        match s {
            "blinkBackground" => Ok(Self::BlinkBackground),
            "lights" => Ok(Self::Lights),
            "antsBlack" => Ok(Self::AntsBlack),
            "antsRed" => Ok(Self::AntsRed),
            "shimmer" => Ok(Self::Shimmer),
            "sparkle" => Ok(Self::Sparkle),
            "none" => Ok(Self::None),
            _ => Err(format!("unknown STTextEffect value: {}", s)),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum STBorder {
    #[serde(rename = "nil")]
    Nil,
    #[serde(rename = "none")]
    None,
    #[serde(rename = "single")]
    Single,
    #[serde(rename = "thick")]
    Thick,
    #[serde(rename = "double")]
    Double,
    #[serde(rename = "dotted")]
    Dotted,
    #[serde(rename = "dashed")]
    Dashed,
    #[serde(rename = "dotDash")]
    DotDash,
    #[serde(rename = "dotDotDash")]
    DotDotDash,
    #[serde(rename = "triple")]
    Triple,
    #[serde(rename = "thinThickSmallGap")]
    ThinThickSmallGap,
    #[serde(rename = "thickThinSmallGap")]
    ThickThinSmallGap,
    #[serde(rename = "thinThickThinSmallGap")]
    ThinThickThinSmallGap,
    #[serde(rename = "thinThickMediumGap")]
    ThinThickMediumGap,
    #[serde(rename = "thickThinMediumGap")]
    ThickThinMediumGap,
    #[serde(rename = "thinThickThinMediumGap")]
    ThinThickThinMediumGap,
    #[serde(rename = "thinThickLargeGap")]
    ThinThickLargeGap,
    #[serde(rename = "thickThinLargeGap")]
    ThickThinLargeGap,
    #[serde(rename = "thinThickThinLargeGap")]
    ThinThickThinLargeGap,
    #[serde(rename = "wave")]
    Wave,
    #[serde(rename = "doubleWave")]
    DoubleWave,
    #[serde(rename = "dashSmallGap")]
    DashSmallGap,
    #[serde(rename = "dashDotStroked")]
    DashDotStroked,
    #[serde(rename = "threeDEmboss")]
    ThreeDEmboss,
    #[serde(rename = "threeDEngrave")]
    ThreeDEngrave,
    #[serde(rename = "outset")]
    Outset,
    #[serde(rename = "inset")]
    Inset,
    #[serde(rename = "apples")]
    Apples,
    #[serde(rename = "archedScallops")]
    ArchedScallops,
    #[serde(rename = "babyPacifier")]
    BabyPacifier,
    #[serde(rename = "babyRattle")]
    BabyRattle,
    #[serde(rename = "balloons3Colors")]
    Balloons3Colors,
    #[serde(rename = "balloonsHotAir")]
    BalloonsHotAir,
    #[serde(rename = "basicBlackDashes")]
    BasicBlackDashes,
    #[serde(rename = "basicBlackDots")]
    BasicBlackDots,
    #[serde(rename = "basicBlackSquares")]
    BasicBlackSquares,
    #[serde(rename = "basicThinLines")]
    BasicThinLines,
    #[serde(rename = "basicWhiteDashes")]
    BasicWhiteDashes,
    #[serde(rename = "basicWhiteDots")]
    BasicWhiteDots,
    #[serde(rename = "basicWhiteSquares")]
    BasicWhiteSquares,
    #[serde(rename = "basicWideInline")]
    BasicWideInline,
    #[serde(rename = "basicWideMidline")]
    BasicWideMidline,
    #[serde(rename = "basicWideOutline")]
    BasicWideOutline,
    #[serde(rename = "bats")]
    Bats,
    #[serde(rename = "birds")]
    Birds,
    #[serde(rename = "birdsFlight")]
    BirdsFlight,
    #[serde(rename = "cabins")]
    Cabins,
    #[serde(rename = "cakeSlice")]
    CakeSlice,
    #[serde(rename = "candyCorn")]
    CandyCorn,
    #[serde(rename = "celticKnotwork")]
    CelticKnotwork,
    #[serde(rename = "certificateBanner")]
    CertificateBanner,
    #[serde(rename = "chainLink")]
    ChainLink,
    #[serde(rename = "champagneBottle")]
    ChampagneBottle,
    #[serde(rename = "checkedBarBlack")]
    CheckedBarBlack,
    #[serde(rename = "checkedBarColor")]
    CheckedBarColor,
    #[serde(rename = "checkered")]
    Checkered,
    #[serde(rename = "christmasTree")]
    ChristmasTree,
    #[serde(rename = "circlesLines")]
    CirclesLines,
    #[serde(rename = "circlesRectangles")]
    CirclesRectangles,
    #[serde(rename = "classicalWave")]
    ClassicalWave,
    #[serde(rename = "clocks")]
    Clocks,
    #[serde(rename = "compass")]
    Compass,
    #[serde(rename = "confetti")]
    Confetti,
    #[serde(rename = "confettiGrays")]
    ConfettiGrays,
    #[serde(rename = "confettiOutline")]
    ConfettiOutline,
    #[serde(rename = "confettiStreamers")]
    ConfettiStreamers,
    #[serde(rename = "confettiWhite")]
    ConfettiWhite,
    #[serde(rename = "cornerTriangles")]
    CornerTriangles,
    #[serde(rename = "couponCutoutDashes")]
    CouponCutoutDashes,
    #[serde(rename = "couponCutoutDots")]
    CouponCutoutDots,
    #[serde(rename = "crazyMaze")]
    CrazyMaze,
    #[serde(rename = "creaturesButterfly")]
    CreaturesButterfly,
    #[serde(rename = "creaturesFish")]
    CreaturesFish,
    #[serde(rename = "creaturesInsects")]
    CreaturesInsects,
    #[serde(rename = "creaturesLadyBug")]
    CreaturesLadyBug,
    #[serde(rename = "crossStitch")]
    CrossStitch,
    #[serde(rename = "cup")]
    Cup,
    #[serde(rename = "decoArch")]
    DecoArch,
    #[serde(rename = "decoArchColor")]
    DecoArchColor,
    #[serde(rename = "decoBlocks")]
    DecoBlocks,
    #[serde(rename = "diamondsGray")]
    DiamondsGray,
    #[serde(rename = "doubleD")]
    DoubleD,
    #[serde(rename = "doubleDiamonds")]
    DoubleDiamonds,
    #[serde(rename = "earth1")]
    Earth1,
    #[serde(rename = "earth2")]
    Earth2,
    #[serde(rename = "earth3")]
    Earth3,
    #[serde(rename = "eclipsingSquares1")]
    EclipsingSquares1,
    #[serde(rename = "eclipsingSquares2")]
    EclipsingSquares2,
    #[serde(rename = "eggsBlack")]
    EggsBlack,
    #[serde(rename = "fans")]
    Fans,
    #[serde(rename = "film")]
    Film,
    #[serde(rename = "firecrackers")]
    Firecrackers,
    #[serde(rename = "flowersBlockPrint")]
    FlowersBlockPrint,
    #[serde(rename = "flowersDaisies")]
    FlowersDaisies,
    #[serde(rename = "flowersModern1")]
    FlowersModern1,
    #[serde(rename = "flowersModern2")]
    FlowersModern2,
    #[serde(rename = "flowersPansy")]
    FlowersPansy,
    #[serde(rename = "flowersRedRose")]
    FlowersRedRose,
    #[serde(rename = "flowersRoses")]
    FlowersRoses,
    #[serde(rename = "flowersTeacup")]
    FlowersTeacup,
    #[serde(rename = "flowersTiny")]
    FlowersTiny,
    #[serde(rename = "gems")]
    Gems,
    #[serde(rename = "gingerbreadMan")]
    GingerbreadMan,
    #[serde(rename = "gradient")]
    Gradient,
    #[serde(rename = "handmade1")]
    Handmade1,
    #[serde(rename = "handmade2")]
    Handmade2,
    #[serde(rename = "heartBalloon")]
    HeartBalloon,
    #[serde(rename = "heartGray")]
    HeartGray,
    #[serde(rename = "hearts")]
    Hearts,
    #[serde(rename = "heebieJeebies")]
    HeebieJeebies,
    #[serde(rename = "holly")]
    Holly,
    #[serde(rename = "houseFunky")]
    HouseFunky,
    #[serde(rename = "hypnotic")]
    Hypnotic,
    #[serde(rename = "iceCreamCones")]
    IceCreamCones,
    #[serde(rename = "lightBulb")]
    LightBulb,
    #[serde(rename = "lightning1")]
    Lightning1,
    #[serde(rename = "lightning2")]
    Lightning2,
    #[serde(rename = "mapPins")]
    MapPins,
    #[serde(rename = "mapleLeaf")]
    MapleLeaf,
    #[serde(rename = "mapleMuffins")]
    MapleMuffins,
    #[serde(rename = "marquee")]
    Marquee,
    #[serde(rename = "marqueeToothed")]
    MarqueeToothed,
    #[serde(rename = "moons")]
    Moons,
    #[serde(rename = "mosaic")]
    Mosaic,
    #[serde(rename = "musicNotes")]
    MusicNotes,
    #[serde(rename = "northwest")]
    Northwest,
    #[serde(rename = "ovals")]
    Ovals,
    #[serde(rename = "packages")]
    Packages,
    #[serde(rename = "palmsBlack")]
    PalmsBlack,
    #[serde(rename = "palmsColor")]
    PalmsColor,
    #[serde(rename = "paperClips")]
    PaperClips,
    #[serde(rename = "papyrus")]
    Papyrus,
    #[serde(rename = "partyFavor")]
    PartyFavor,
    #[serde(rename = "partyGlass")]
    PartyGlass,
    #[serde(rename = "pencils")]
    Pencils,
    #[serde(rename = "people")]
    People,
    #[serde(rename = "peopleWaving")]
    PeopleWaving,
    #[serde(rename = "peopleHats")]
    PeopleHats,
    #[serde(rename = "poinsettias")]
    Poinsettias,
    #[serde(rename = "postageStamp")]
    PostageStamp,
    #[serde(rename = "pumpkin1")]
    Pumpkin1,
    #[serde(rename = "pushPinNote2")]
    PushPinNote2,
    #[serde(rename = "pushPinNote1")]
    PushPinNote1,
    #[serde(rename = "pyramids")]
    Pyramids,
    #[serde(rename = "pyramidsAbove")]
    PyramidsAbove,
    #[serde(rename = "quadrants")]
    Quadrants,
    #[serde(rename = "rings")]
    Rings,
    #[serde(rename = "safari")]
    Safari,
    #[serde(rename = "sawtooth")]
    Sawtooth,
    #[serde(rename = "sawtoothGray")]
    SawtoothGray,
    #[serde(rename = "scaredCat")]
    ScaredCat,
    #[serde(rename = "seattle")]
    Seattle,
    #[serde(rename = "shadowedSquares")]
    ShadowedSquares,
    #[serde(rename = "sharksTeeth")]
    SharksTeeth,
    #[serde(rename = "shorebirdTracks")]
    ShorebirdTracks,
    #[serde(rename = "skyrocket")]
    Skyrocket,
    #[serde(rename = "snowflakeFancy")]
    SnowflakeFancy,
    #[serde(rename = "snowflakes")]
    Snowflakes,
    #[serde(rename = "sombrero")]
    Sombrero,
    #[serde(rename = "southwest")]
    Southwest,
    #[serde(rename = "stars")]
    Stars,
    #[serde(rename = "starsTop")]
    StarsTop,
    #[serde(rename = "stars3d")]
    Stars3d,
    #[serde(rename = "starsBlack")]
    StarsBlack,
    #[serde(rename = "starsShadowed")]
    StarsShadowed,
    #[serde(rename = "sun")]
    Sun,
    #[serde(rename = "swirligig")]
    Swirligig,
    #[serde(rename = "tornPaper")]
    TornPaper,
    #[serde(rename = "tornPaperBlack")]
    TornPaperBlack,
    #[serde(rename = "trees")]
    Trees,
    #[serde(rename = "triangleParty")]
    TriangleParty,
    #[serde(rename = "triangles")]
    Triangles,
    #[serde(rename = "triangle1")]
    Triangle1,
    #[serde(rename = "triangle2")]
    Triangle2,
    #[serde(rename = "triangleCircle1")]
    TriangleCircle1,
    #[serde(rename = "triangleCircle2")]
    TriangleCircle2,
    #[serde(rename = "shapes1")]
    Shapes1,
    #[serde(rename = "shapes2")]
    Shapes2,
    #[serde(rename = "twistedLines1")]
    TwistedLines1,
    #[serde(rename = "twistedLines2")]
    TwistedLines2,
    #[serde(rename = "vine")]
    Vine,
    #[serde(rename = "waveline")]
    Waveline,
    #[serde(rename = "weavingAngles")]
    WeavingAngles,
    #[serde(rename = "weavingBraid")]
    WeavingBraid,
    #[serde(rename = "weavingRibbon")]
    WeavingRibbon,
    #[serde(rename = "weavingStrips")]
    WeavingStrips,
    #[serde(rename = "whiteFlowers")]
    WhiteFlowers,
    #[serde(rename = "woodwork")]
    Woodwork,
    #[serde(rename = "xIllusions")]
    XIllusions,
    #[serde(rename = "zanyTriangles")]
    ZanyTriangles,
    #[serde(rename = "zigZag")]
    ZigZag,
    #[serde(rename = "zigZagStitch")]
    ZigZagStitch,
    #[serde(rename = "custom")]
    Custom,
}

impl std::fmt::Display for STBorder {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::Nil => write!(f, "nil"),
            Self::None => write!(f, "none"),
            Self::Single => write!(f, "single"),
            Self::Thick => write!(f, "thick"),
            Self::Double => write!(f, "double"),
            Self::Dotted => write!(f, "dotted"),
            Self::Dashed => write!(f, "dashed"),
            Self::DotDash => write!(f, "dotDash"),
            Self::DotDotDash => write!(f, "dotDotDash"),
            Self::Triple => write!(f, "triple"),
            Self::ThinThickSmallGap => write!(f, "thinThickSmallGap"),
            Self::ThickThinSmallGap => write!(f, "thickThinSmallGap"),
            Self::ThinThickThinSmallGap => write!(f, "thinThickThinSmallGap"),
            Self::ThinThickMediumGap => write!(f, "thinThickMediumGap"),
            Self::ThickThinMediumGap => write!(f, "thickThinMediumGap"),
            Self::ThinThickThinMediumGap => write!(f, "thinThickThinMediumGap"),
            Self::ThinThickLargeGap => write!(f, "thinThickLargeGap"),
            Self::ThickThinLargeGap => write!(f, "thickThinLargeGap"),
            Self::ThinThickThinLargeGap => write!(f, "thinThickThinLargeGap"),
            Self::Wave => write!(f, "wave"),
            Self::DoubleWave => write!(f, "doubleWave"),
            Self::DashSmallGap => write!(f, "dashSmallGap"),
            Self::DashDotStroked => write!(f, "dashDotStroked"),
            Self::ThreeDEmboss => write!(f, "threeDEmboss"),
            Self::ThreeDEngrave => write!(f, "threeDEngrave"),
            Self::Outset => write!(f, "outset"),
            Self::Inset => write!(f, "inset"),
            Self::Apples => write!(f, "apples"),
            Self::ArchedScallops => write!(f, "archedScallops"),
            Self::BabyPacifier => write!(f, "babyPacifier"),
            Self::BabyRattle => write!(f, "babyRattle"),
            Self::Balloons3Colors => write!(f, "balloons3Colors"),
            Self::BalloonsHotAir => write!(f, "balloonsHotAir"),
            Self::BasicBlackDashes => write!(f, "basicBlackDashes"),
            Self::BasicBlackDots => write!(f, "basicBlackDots"),
            Self::BasicBlackSquares => write!(f, "basicBlackSquares"),
            Self::BasicThinLines => write!(f, "basicThinLines"),
            Self::BasicWhiteDashes => write!(f, "basicWhiteDashes"),
            Self::BasicWhiteDots => write!(f, "basicWhiteDots"),
            Self::BasicWhiteSquares => write!(f, "basicWhiteSquares"),
            Self::BasicWideInline => write!(f, "basicWideInline"),
            Self::BasicWideMidline => write!(f, "basicWideMidline"),
            Self::BasicWideOutline => write!(f, "basicWideOutline"),
            Self::Bats => write!(f, "bats"),
            Self::Birds => write!(f, "birds"),
            Self::BirdsFlight => write!(f, "birdsFlight"),
            Self::Cabins => write!(f, "cabins"),
            Self::CakeSlice => write!(f, "cakeSlice"),
            Self::CandyCorn => write!(f, "candyCorn"),
            Self::CelticKnotwork => write!(f, "celticKnotwork"),
            Self::CertificateBanner => write!(f, "certificateBanner"),
            Self::ChainLink => write!(f, "chainLink"),
            Self::ChampagneBottle => write!(f, "champagneBottle"),
            Self::CheckedBarBlack => write!(f, "checkedBarBlack"),
            Self::CheckedBarColor => write!(f, "checkedBarColor"),
            Self::Checkered => write!(f, "checkered"),
            Self::ChristmasTree => write!(f, "christmasTree"),
            Self::CirclesLines => write!(f, "circlesLines"),
            Self::CirclesRectangles => write!(f, "circlesRectangles"),
            Self::ClassicalWave => write!(f, "classicalWave"),
            Self::Clocks => write!(f, "clocks"),
            Self::Compass => write!(f, "compass"),
            Self::Confetti => write!(f, "confetti"),
            Self::ConfettiGrays => write!(f, "confettiGrays"),
            Self::ConfettiOutline => write!(f, "confettiOutline"),
            Self::ConfettiStreamers => write!(f, "confettiStreamers"),
            Self::ConfettiWhite => write!(f, "confettiWhite"),
            Self::CornerTriangles => write!(f, "cornerTriangles"),
            Self::CouponCutoutDashes => write!(f, "couponCutoutDashes"),
            Self::CouponCutoutDots => write!(f, "couponCutoutDots"),
            Self::CrazyMaze => write!(f, "crazyMaze"),
            Self::CreaturesButterfly => write!(f, "creaturesButterfly"),
            Self::CreaturesFish => write!(f, "creaturesFish"),
            Self::CreaturesInsects => write!(f, "creaturesInsects"),
            Self::CreaturesLadyBug => write!(f, "creaturesLadyBug"),
            Self::CrossStitch => write!(f, "crossStitch"),
            Self::Cup => write!(f, "cup"),
            Self::DecoArch => write!(f, "decoArch"),
            Self::DecoArchColor => write!(f, "decoArchColor"),
            Self::DecoBlocks => write!(f, "decoBlocks"),
            Self::DiamondsGray => write!(f, "diamondsGray"),
            Self::DoubleD => write!(f, "doubleD"),
            Self::DoubleDiamonds => write!(f, "doubleDiamonds"),
            Self::Earth1 => write!(f, "earth1"),
            Self::Earth2 => write!(f, "earth2"),
            Self::Earth3 => write!(f, "earth3"),
            Self::EclipsingSquares1 => write!(f, "eclipsingSquares1"),
            Self::EclipsingSquares2 => write!(f, "eclipsingSquares2"),
            Self::EggsBlack => write!(f, "eggsBlack"),
            Self::Fans => write!(f, "fans"),
            Self::Film => write!(f, "film"),
            Self::Firecrackers => write!(f, "firecrackers"),
            Self::FlowersBlockPrint => write!(f, "flowersBlockPrint"),
            Self::FlowersDaisies => write!(f, "flowersDaisies"),
            Self::FlowersModern1 => write!(f, "flowersModern1"),
            Self::FlowersModern2 => write!(f, "flowersModern2"),
            Self::FlowersPansy => write!(f, "flowersPansy"),
            Self::FlowersRedRose => write!(f, "flowersRedRose"),
            Self::FlowersRoses => write!(f, "flowersRoses"),
            Self::FlowersTeacup => write!(f, "flowersTeacup"),
            Self::FlowersTiny => write!(f, "flowersTiny"),
            Self::Gems => write!(f, "gems"),
            Self::GingerbreadMan => write!(f, "gingerbreadMan"),
            Self::Gradient => write!(f, "gradient"),
            Self::Handmade1 => write!(f, "handmade1"),
            Self::Handmade2 => write!(f, "handmade2"),
            Self::HeartBalloon => write!(f, "heartBalloon"),
            Self::HeartGray => write!(f, "heartGray"),
            Self::Hearts => write!(f, "hearts"),
            Self::HeebieJeebies => write!(f, "heebieJeebies"),
            Self::Holly => write!(f, "holly"),
            Self::HouseFunky => write!(f, "houseFunky"),
            Self::Hypnotic => write!(f, "hypnotic"),
            Self::IceCreamCones => write!(f, "iceCreamCones"),
            Self::LightBulb => write!(f, "lightBulb"),
            Self::Lightning1 => write!(f, "lightning1"),
            Self::Lightning2 => write!(f, "lightning2"),
            Self::MapPins => write!(f, "mapPins"),
            Self::MapleLeaf => write!(f, "mapleLeaf"),
            Self::MapleMuffins => write!(f, "mapleMuffins"),
            Self::Marquee => write!(f, "marquee"),
            Self::MarqueeToothed => write!(f, "marqueeToothed"),
            Self::Moons => write!(f, "moons"),
            Self::Mosaic => write!(f, "mosaic"),
            Self::MusicNotes => write!(f, "musicNotes"),
            Self::Northwest => write!(f, "northwest"),
            Self::Ovals => write!(f, "ovals"),
            Self::Packages => write!(f, "packages"),
            Self::PalmsBlack => write!(f, "palmsBlack"),
            Self::PalmsColor => write!(f, "palmsColor"),
            Self::PaperClips => write!(f, "paperClips"),
            Self::Papyrus => write!(f, "papyrus"),
            Self::PartyFavor => write!(f, "partyFavor"),
            Self::PartyGlass => write!(f, "partyGlass"),
            Self::Pencils => write!(f, "pencils"),
            Self::People => write!(f, "people"),
            Self::PeopleWaving => write!(f, "peopleWaving"),
            Self::PeopleHats => write!(f, "peopleHats"),
            Self::Poinsettias => write!(f, "poinsettias"),
            Self::PostageStamp => write!(f, "postageStamp"),
            Self::Pumpkin1 => write!(f, "pumpkin1"),
            Self::PushPinNote2 => write!(f, "pushPinNote2"),
            Self::PushPinNote1 => write!(f, "pushPinNote1"),
            Self::Pyramids => write!(f, "pyramids"),
            Self::PyramidsAbove => write!(f, "pyramidsAbove"),
            Self::Quadrants => write!(f, "quadrants"),
            Self::Rings => write!(f, "rings"),
            Self::Safari => write!(f, "safari"),
            Self::Sawtooth => write!(f, "sawtooth"),
            Self::SawtoothGray => write!(f, "sawtoothGray"),
            Self::ScaredCat => write!(f, "scaredCat"),
            Self::Seattle => write!(f, "seattle"),
            Self::ShadowedSquares => write!(f, "shadowedSquares"),
            Self::SharksTeeth => write!(f, "sharksTeeth"),
            Self::ShorebirdTracks => write!(f, "shorebirdTracks"),
            Self::Skyrocket => write!(f, "skyrocket"),
            Self::SnowflakeFancy => write!(f, "snowflakeFancy"),
            Self::Snowflakes => write!(f, "snowflakes"),
            Self::Sombrero => write!(f, "sombrero"),
            Self::Southwest => write!(f, "southwest"),
            Self::Stars => write!(f, "stars"),
            Self::StarsTop => write!(f, "starsTop"),
            Self::Stars3d => write!(f, "stars3d"),
            Self::StarsBlack => write!(f, "starsBlack"),
            Self::StarsShadowed => write!(f, "starsShadowed"),
            Self::Sun => write!(f, "sun"),
            Self::Swirligig => write!(f, "swirligig"),
            Self::TornPaper => write!(f, "tornPaper"),
            Self::TornPaperBlack => write!(f, "tornPaperBlack"),
            Self::Trees => write!(f, "trees"),
            Self::TriangleParty => write!(f, "triangleParty"),
            Self::Triangles => write!(f, "triangles"),
            Self::Triangle1 => write!(f, "triangle1"),
            Self::Triangle2 => write!(f, "triangle2"),
            Self::TriangleCircle1 => write!(f, "triangleCircle1"),
            Self::TriangleCircle2 => write!(f, "triangleCircle2"),
            Self::Shapes1 => write!(f, "shapes1"),
            Self::Shapes2 => write!(f, "shapes2"),
            Self::TwistedLines1 => write!(f, "twistedLines1"),
            Self::TwistedLines2 => write!(f, "twistedLines2"),
            Self::Vine => write!(f, "vine"),
            Self::Waveline => write!(f, "waveline"),
            Self::WeavingAngles => write!(f, "weavingAngles"),
            Self::WeavingBraid => write!(f, "weavingBraid"),
            Self::WeavingRibbon => write!(f, "weavingRibbon"),
            Self::WeavingStrips => write!(f, "weavingStrips"),
            Self::WhiteFlowers => write!(f, "whiteFlowers"),
            Self::Woodwork => write!(f, "woodwork"),
            Self::XIllusions => write!(f, "xIllusions"),
            Self::ZanyTriangles => write!(f, "zanyTriangles"),
            Self::ZigZag => write!(f, "zigZag"),
            Self::ZigZagStitch => write!(f, "zigZagStitch"),
            Self::Custom => write!(f, "custom"),
        }
    }
}

impl std::str::FromStr for STBorder {
    type Err = String;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        match s {
            "nil" => Ok(Self::Nil),
            "none" => Ok(Self::None),
            "single" => Ok(Self::Single),
            "thick" => Ok(Self::Thick),
            "double" => Ok(Self::Double),
            "dotted" => Ok(Self::Dotted),
            "dashed" => Ok(Self::Dashed),
            "dotDash" => Ok(Self::DotDash),
            "dotDotDash" => Ok(Self::DotDotDash),
            "triple" => Ok(Self::Triple),
            "thinThickSmallGap" => Ok(Self::ThinThickSmallGap),
            "thickThinSmallGap" => Ok(Self::ThickThinSmallGap),
            "thinThickThinSmallGap" => Ok(Self::ThinThickThinSmallGap),
            "thinThickMediumGap" => Ok(Self::ThinThickMediumGap),
            "thickThinMediumGap" => Ok(Self::ThickThinMediumGap),
            "thinThickThinMediumGap" => Ok(Self::ThinThickThinMediumGap),
            "thinThickLargeGap" => Ok(Self::ThinThickLargeGap),
            "thickThinLargeGap" => Ok(Self::ThickThinLargeGap),
            "thinThickThinLargeGap" => Ok(Self::ThinThickThinLargeGap),
            "wave" => Ok(Self::Wave),
            "doubleWave" => Ok(Self::DoubleWave),
            "dashSmallGap" => Ok(Self::DashSmallGap),
            "dashDotStroked" => Ok(Self::DashDotStroked),
            "threeDEmboss" => Ok(Self::ThreeDEmboss),
            "threeDEngrave" => Ok(Self::ThreeDEngrave),
            "outset" => Ok(Self::Outset),
            "inset" => Ok(Self::Inset),
            "apples" => Ok(Self::Apples),
            "archedScallops" => Ok(Self::ArchedScallops),
            "babyPacifier" => Ok(Self::BabyPacifier),
            "babyRattle" => Ok(Self::BabyRattle),
            "balloons3Colors" => Ok(Self::Balloons3Colors),
            "balloonsHotAir" => Ok(Self::BalloonsHotAir),
            "basicBlackDashes" => Ok(Self::BasicBlackDashes),
            "basicBlackDots" => Ok(Self::BasicBlackDots),
            "basicBlackSquares" => Ok(Self::BasicBlackSquares),
            "basicThinLines" => Ok(Self::BasicThinLines),
            "basicWhiteDashes" => Ok(Self::BasicWhiteDashes),
            "basicWhiteDots" => Ok(Self::BasicWhiteDots),
            "basicWhiteSquares" => Ok(Self::BasicWhiteSquares),
            "basicWideInline" => Ok(Self::BasicWideInline),
            "basicWideMidline" => Ok(Self::BasicWideMidline),
            "basicWideOutline" => Ok(Self::BasicWideOutline),
            "bats" => Ok(Self::Bats),
            "birds" => Ok(Self::Birds),
            "birdsFlight" => Ok(Self::BirdsFlight),
            "cabins" => Ok(Self::Cabins),
            "cakeSlice" => Ok(Self::CakeSlice),
            "candyCorn" => Ok(Self::CandyCorn),
            "celticKnotwork" => Ok(Self::CelticKnotwork),
            "certificateBanner" => Ok(Self::CertificateBanner),
            "chainLink" => Ok(Self::ChainLink),
            "champagneBottle" => Ok(Self::ChampagneBottle),
            "checkedBarBlack" => Ok(Self::CheckedBarBlack),
            "checkedBarColor" => Ok(Self::CheckedBarColor),
            "checkered" => Ok(Self::Checkered),
            "christmasTree" => Ok(Self::ChristmasTree),
            "circlesLines" => Ok(Self::CirclesLines),
            "circlesRectangles" => Ok(Self::CirclesRectangles),
            "classicalWave" => Ok(Self::ClassicalWave),
            "clocks" => Ok(Self::Clocks),
            "compass" => Ok(Self::Compass),
            "confetti" => Ok(Self::Confetti),
            "confettiGrays" => Ok(Self::ConfettiGrays),
            "confettiOutline" => Ok(Self::ConfettiOutline),
            "confettiStreamers" => Ok(Self::ConfettiStreamers),
            "confettiWhite" => Ok(Self::ConfettiWhite),
            "cornerTriangles" => Ok(Self::CornerTriangles),
            "couponCutoutDashes" => Ok(Self::CouponCutoutDashes),
            "couponCutoutDots" => Ok(Self::CouponCutoutDots),
            "crazyMaze" => Ok(Self::CrazyMaze),
            "creaturesButterfly" => Ok(Self::CreaturesButterfly),
            "creaturesFish" => Ok(Self::CreaturesFish),
            "creaturesInsects" => Ok(Self::CreaturesInsects),
            "creaturesLadyBug" => Ok(Self::CreaturesLadyBug),
            "crossStitch" => Ok(Self::CrossStitch),
            "cup" => Ok(Self::Cup),
            "decoArch" => Ok(Self::DecoArch),
            "decoArchColor" => Ok(Self::DecoArchColor),
            "decoBlocks" => Ok(Self::DecoBlocks),
            "diamondsGray" => Ok(Self::DiamondsGray),
            "doubleD" => Ok(Self::DoubleD),
            "doubleDiamonds" => Ok(Self::DoubleDiamonds),
            "earth1" => Ok(Self::Earth1),
            "earth2" => Ok(Self::Earth2),
            "earth3" => Ok(Self::Earth3),
            "eclipsingSquares1" => Ok(Self::EclipsingSquares1),
            "eclipsingSquares2" => Ok(Self::EclipsingSquares2),
            "eggsBlack" => Ok(Self::EggsBlack),
            "fans" => Ok(Self::Fans),
            "film" => Ok(Self::Film),
            "firecrackers" => Ok(Self::Firecrackers),
            "flowersBlockPrint" => Ok(Self::FlowersBlockPrint),
            "flowersDaisies" => Ok(Self::FlowersDaisies),
            "flowersModern1" => Ok(Self::FlowersModern1),
            "flowersModern2" => Ok(Self::FlowersModern2),
            "flowersPansy" => Ok(Self::FlowersPansy),
            "flowersRedRose" => Ok(Self::FlowersRedRose),
            "flowersRoses" => Ok(Self::FlowersRoses),
            "flowersTeacup" => Ok(Self::FlowersTeacup),
            "flowersTiny" => Ok(Self::FlowersTiny),
            "gems" => Ok(Self::Gems),
            "gingerbreadMan" => Ok(Self::GingerbreadMan),
            "gradient" => Ok(Self::Gradient),
            "handmade1" => Ok(Self::Handmade1),
            "handmade2" => Ok(Self::Handmade2),
            "heartBalloon" => Ok(Self::HeartBalloon),
            "heartGray" => Ok(Self::HeartGray),
            "hearts" => Ok(Self::Hearts),
            "heebieJeebies" => Ok(Self::HeebieJeebies),
            "holly" => Ok(Self::Holly),
            "houseFunky" => Ok(Self::HouseFunky),
            "hypnotic" => Ok(Self::Hypnotic),
            "iceCreamCones" => Ok(Self::IceCreamCones),
            "lightBulb" => Ok(Self::LightBulb),
            "lightning1" => Ok(Self::Lightning1),
            "lightning2" => Ok(Self::Lightning2),
            "mapPins" => Ok(Self::MapPins),
            "mapleLeaf" => Ok(Self::MapleLeaf),
            "mapleMuffins" => Ok(Self::MapleMuffins),
            "marquee" => Ok(Self::Marquee),
            "marqueeToothed" => Ok(Self::MarqueeToothed),
            "moons" => Ok(Self::Moons),
            "mosaic" => Ok(Self::Mosaic),
            "musicNotes" => Ok(Self::MusicNotes),
            "northwest" => Ok(Self::Northwest),
            "ovals" => Ok(Self::Ovals),
            "packages" => Ok(Self::Packages),
            "palmsBlack" => Ok(Self::PalmsBlack),
            "palmsColor" => Ok(Self::PalmsColor),
            "paperClips" => Ok(Self::PaperClips),
            "papyrus" => Ok(Self::Papyrus),
            "partyFavor" => Ok(Self::PartyFavor),
            "partyGlass" => Ok(Self::PartyGlass),
            "pencils" => Ok(Self::Pencils),
            "people" => Ok(Self::People),
            "peopleWaving" => Ok(Self::PeopleWaving),
            "peopleHats" => Ok(Self::PeopleHats),
            "poinsettias" => Ok(Self::Poinsettias),
            "postageStamp" => Ok(Self::PostageStamp),
            "pumpkin1" => Ok(Self::Pumpkin1),
            "pushPinNote2" => Ok(Self::PushPinNote2),
            "pushPinNote1" => Ok(Self::PushPinNote1),
            "pyramids" => Ok(Self::Pyramids),
            "pyramidsAbove" => Ok(Self::PyramidsAbove),
            "quadrants" => Ok(Self::Quadrants),
            "rings" => Ok(Self::Rings),
            "safari" => Ok(Self::Safari),
            "sawtooth" => Ok(Self::Sawtooth),
            "sawtoothGray" => Ok(Self::SawtoothGray),
            "scaredCat" => Ok(Self::ScaredCat),
            "seattle" => Ok(Self::Seattle),
            "shadowedSquares" => Ok(Self::ShadowedSquares),
            "sharksTeeth" => Ok(Self::SharksTeeth),
            "shorebirdTracks" => Ok(Self::ShorebirdTracks),
            "skyrocket" => Ok(Self::Skyrocket),
            "snowflakeFancy" => Ok(Self::SnowflakeFancy),
            "snowflakes" => Ok(Self::Snowflakes),
            "sombrero" => Ok(Self::Sombrero),
            "southwest" => Ok(Self::Southwest),
            "stars" => Ok(Self::Stars),
            "starsTop" => Ok(Self::StarsTop),
            "stars3d" => Ok(Self::Stars3d),
            "starsBlack" => Ok(Self::StarsBlack),
            "starsShadowed" => Ok(Self::StarsShadowed),
            "sun" => Ok(Self::Sun),
            "swirligig" => Ok(Self::Swirligig),
            "tornPaper" => Ok(Self::TornPaper),
            "tornPaperBlack" => Ok(Self::TornPaperBlack),
            "trees" => Ok(Self::Trees),
            "triangleParty" => Ok(Self::TriangleParty),
            "triangles" => Ok(Self::Triangles),
            "triangle1" => Ok(Self::Triangle1),
            "triangle2" => Ok(Self::Triangle2),
            "triangleCircle1" => Ok(Self::TriangleCircle1),
            "triangleCircle2" => Ok(Self::TriangleCircle2),
            "shapes1" => Ok(Self::Shapes1),
            "shapes2" => Ok(Self::Shapes2),
            "twistedLines1" => Ok(Self::TwistedLines1),
            "twistedLines2" => Ok(Self::TwistedLines2),
            "vine" => Ok(Self::Vine),
            "waveline" => Ok(Self::Waveline),
            "weavingAngles" => Ok(Self::WeavingAngles),
            "weavingBraid" => Ok(Self::WeavingBraid),
            "weavingRibbon" => Ok(Self::WeavingRibbon),
            "weavingStrips" => Ok(Self::WeavingStrips),
            "whiteFlowers" => Ok(Self::WhiteFlowers),
            "woodwork" => Ok(Self::Woodwork),
            "xIllusions" => Ok(Self::XIllusions),
            "zanyTriangles" => Ok(Self::ZanyTriangles),
            "zigZag" => Ok(Self::ZigZag),
            "zigZagStitch" => Ok(Self::ZigZagStitch),
            "custom" => Ok(Self::Custom),
            _ => Err(format!("unknown STBorder value: {}", s)),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum STShd {
    #[serde(rename = "nil")]
    Nil,
    #[serde(rename = "clear")]
    Clear,
    #[serde(rename = "solid")]
    Solid,
    #[serde(rename = "horzStripe")]
    HorzStripe,
    #[serde(rename = "vertStripe")]
    VertStripe,
    #[serde(rename = "reverseDiagStripe")]
    ReverseDiagStripe,
    #[serde(rename = "diagStripe")]
    DiagStripe,
    #[serde(rename = "horzCross")]
    HorzCross,
    #[serde(rename = "diagCross")]
    DiagCross,
    #[serde(rename = "thinHorzStripe")]
    ThinHorzStripe,
    #[serde(rename = "thinVertStripe")]
    ThinVertStripe,
    #[serde(rename = "thinReverseDiagStripe")]
    ThinReverseDiagStripe,
    #[serde(rename = "thinDiagStripe")]
    ThinDiagStripe,
    #[serde(rename = "thinHorzCross")]
    ThinHorzCross,
    #[serde(rename = "thinDiagCross")]
    ThinDiagCross,
    #[serde(rename = "pct5")]
    Pct5,
    #[serde(rename = "pct10")]
    Pct10,
    #[serde(rename = "pct12")]
    Pct12,
    #[serde(rename = "pct15")]
    Pct15,
    #[serde(rename = "pct20")]
    Pct20,
    #[serde(rename = "pct25")]
    Pct25,
    #[serde(rename = "pct30")]
    Pct30,
    #[serde(rename = "pct35")]
    Pct35,
    #[serde(rename = "pct37")]
    Pct37,
    #[serde(rename = "pct40")]
    Pct40,
    #[serde(rename = "pct45")]
    Pct45,
    #[serde(rename = "pct50")]
    Pct50,
    #[serde(rename = "pct55")]
    Pct55,
    #[serde(rename = "pct60")]
    Pct60,
    #[serde(rename = "pct62")]
    Pct62,
    #[serde(rename = "pct65")]
    Pct65,
    #[serde(rename = "pct70")]
    Pct70,
    #[serde(rename = "pct75")]
    Pct75,
    #[serde(rename = "pct80")]
    Pct80,
    #[serde(rename = "pct85")]
    Pct85,
    #[serde(rename = "pct87")]
    Pct87,
    #[serde(rename = "pct90")]
    Pct90,
    #[serde(rename = "pct95")]
    Pct95,
}

impl std::fmt::Display for STShd {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::Nil => write!(f, "nil"),
            Self::Clear => write!(f, "clear"),
            Self::Solid => write!(f, "solid"),
            Self::HorzStripe => write!(f, "horzStripe"),
            Self::VertStripe => write!(f, "vertStripe"),
            Self::ReverseDiagStripe => write!(f, "reverseDiagStripe"),
            Self::DiagStripe => write!(f, "diagStripe"),
            Self::HorzCross => write!(f, "horzCross"),
            Self::DiagCross => write!(f, "diagCross"),
            Self::ThinHorzStripe => write!(f, "thinHorzStripe"),
            Self::ThinVertStripe => write!(f, "thinVertStripe"),
            Self::ThinReverseDiagStripe => write!(f, "thinReverseDiagStripe"),
            Self::ThinDiagStripe => write!(f, "thinDiagStripe"),
            Self::ThinHorzCross => write!(f, "thinHorzCross"),
            Self::ThinDiagCross => write!(f, "thinDiagCross"),
            Self::Pct5 => write!(f, "pct5"),
            Self::Pct10 => write!(f, "pct10"),
            Self::Pct12 => write!(f, "pct12"),
            Self::Pct15 => write!(f, "pct15"),
            Self::Pct20 => write!(f, "pct20"),
            Self::Pct25 => write!(f, "pct25"),
            Self::Pct30 => write!(f, "pct30"),
            Self::Pct35 => write!(f, "pct35"),
            Self::Pct37 => write!(f, "pct37"),
            Self::Pct40 => write!(f, "pct40"),
            Self::Pct45 => write!(f, "pct45"),
            Self::Pct50 => write!(f, "pct50"),
            Self::Pct55 => write!(f, "pct55"),
            Self::Pct60 => write!(f, "pct60"),
            Self::Pct62 => write!(f, "pct62"),
            Self::Pct65 => write!(f, "pct65"),
            Self::Pct70 => write!(f, "pct70"),
            Self::Pct75 => write!(f, "pct75"),
            Self::Pct80 => write!(f, "pct80"),
            Self::Pct85 => write!(f, "pct85"),
            Self::Pct87 => write!(f, "pct87"),
            Self::Pct90 => write!(f, "pct90"),
            Self::Pct95 => write!(f, "pct95"),
        }
    }
}

impl std::str::FromStr for STShd {
    type Err = String;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        match s {
            "nil" => Ok(Self::Nil),
            "clear" => Ok(Self::Clear),
            "solid" => Ok(Self::Solid),
            "horzStripe" => Ok(Self::HorzStripe),
            "vertStripe" => Ok(Self::VertStripe),
            "reverseDiagStripe" => Ok(Self::ReverseDiagStripe),
            "diagStripe" => Ok(Self::DiagStripe),
            "horzCross" => Ok(Self::HorzCross),
            "diagCross" => Ok(Self::DiagCross),
            "thinHorzStripe" => Ok(Self::ThinHorzStripe),
            "thinVertStripe" => Ok(Self::ThinVertStripe),
            "thinReverseDiagStripe" => Ok(Self::ThinReverseDiagStripe),
            "thinDiagStripe" => Ok(Self::ThinDiagStripe),
            "thinHorzCross" => Ok(Self::ThinHorzCross),
            "thinDiagCross" => Ok(Self::ThinDiagCross),
            "pct5" => Ok(Self::Pct5),
            "pct10" => Ok(Self::Pct10),
            "pct12" => Ok(Self::Pct12),
            "pct15" => Ok(Self::Pct15),
            "pct20" => Ok(Self::Pct20),
            "pct25" => Ok(Self::Pct25),
            "pct30" => Ok(Self::Pct30),
            "pct35" => Ok(Self::Pct35),
            "pct37" => Ok(Self::Pct37),
            "pct40" => Ok(Self::Pct40),
            "pct45" => Ok(Self::Pct45),
            "pct50" => Ok(Self::Pct50),
            "pct55" => Ok(Self::Pct55),
            "pct60" => Ok(Self::Pct60),
            "pct62" => Ok(Self::Pct62),
            "pct65" => Ok(Self::Pct65),
            "pct70" => Ok(Self::Pct70),
            "pct75" => Ok(Self::Pct75),
            "pct80" => Ok(Self::Pct80),
            "pct85" => Ok(Self::Pct85),
            "pct87" => Ok(Self::Pct87),
            "pct90" => Ok(Self::Pct90),
            "pct95" => Ok(Self::Pct95),
            _ => Err(format!("unknown STShd value: {}", s)),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum STEm {
    #[serde(rename = "none")]
    None,
    #[serde(rename = "dot")]
    Dot,
    #[serde(rename = "comma")]
    Comma,
    #[serde(rename = "circle")]
    Circle,
    #[serde(rename = "underDot")]
    UnderDot,
}

impl std::fmt::Display for STEm {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::None => write!(f, "none"),
            Self::Dot => write!(f, "dot"),
            Self::Comma => write!(f, "comma"),
            Self::Circle => write!(f, "circle"),
            Self::UnderDot => write!(f, "underDot"),
        }
    }
}

impl std::str::FromStr for STEm {
    type Err = String;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        match s {
            "none" => Ok(Self::None),
            "dot" => Ok(Self::Dot),
            "comma" => Ok(Self::Comma),
            "circle" => Ok(Self::Circle),
            "underDot" => Ok(Self::UnderDot),
            _ => Err(format!("unknown STEm value: {}", s)),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum STCombineBrackets {
    #[serde(rename = "none")]
    None,
    #[serde(rename = "round")]
    Round,
    #[serde(rename = "square")]
    Square,
    #[serde(rename = "angle")]
    Angle,
    #[serde(rename = "curly")]
    Curly,
}

impl std::fmt::Display for STCombineBrackets {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::None => write!(f, "none"),
            Self::Round => write!(f, "round"),
            Self::Square => write!(f, "square"),
            Self::Angle => write!(f, "angle"),
            Self::Curly => write!(f, "curly"),
        }
    }
}

impl std::str::FromStr for STCombineBrackets {
    type Err = String;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        match s {
            "none" => Ok(Self::None),
            "round" => Ok(Self::Round),
            "square" => Ok(Self::Square),
            "angle" => Ok(Self::Angle),
            "curly" => Ok(Self::Curly),
            _ => Err(format!("unknown STCombineBrackets value: {}", s)),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum STHeightRule {
    #[serde(rename = "auto")]
    Auto,
    #[serde(rename = "exact")]
    Exact,
    #[serde(rename = "atLeast")]
    AtLeast,
}

impl std::fmt::Display for STHeightRule {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::Auto => write!(f, "auto"),
            Self::Exact => write!(f, "exact"),
            Self::AtLeast => write!(f, "atLeast"),
        }
    }
}

impl std::str::FromStr for STHeightRule {
    type Err = String;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        match s {
            "auto" => Ok(Self::Auto),
            "exact" => Ok(Self::Exact),
            "atLeast" => Ok(Self::AtLeast),
            _ => Err(format!("unknown STHeightRule value: {}", s)),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum STWrap {
    #[serde(rename = "auto")]
    Auto,
    #[serde(rename = "notBeside")]
    NotBeside,
    #[serde(rename = "around")]
    Around,
    #[serde(rename = "tight")]
    Tight,
    #[serde(rename = "through")]
    Through,
    #[serde(rename = "none")]
    None,
}

impl std::fmt::Display for STWrap {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::Auto => write!(f, "auto"),
            Self::NotBeside => write!(f, "notBeside"),
            Self::Around => write!(f, "around"),
            Self::Tight => write!(f, "tight"),
            Self::Through => write!(f, "through"),
            Self::None => write!(f, "none"),
        }
    }
}

impl std::str::FromStr for STWrap {
    type Err = String;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        match s {
            "auto" => Ok(Self::Auto),
            "notBeside" => Ok(Self::NotBeside),
            "around" => Ok(Self::Around),
            "tight" => Ok(Self::Tight),
            "through" => Ok(Self::Through),
            "none" => Ok(Self::None),
            _ => Err(format!("unknown STWrap value: {}", s)),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum STVAnchor {
    #[serde(rename = "text")]
    Text,
    #[serde(rename = "margin")]
    Margin,
    #[serde(rename = "page")]
    Page,
}

impl std::fmt::Display for STVAnchor {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::Text => write!(f, "text"),
            Self::Margin => write!(f, "margin"),
            Self::Page => write!(f, "page"),
        }
    }
}

impl std::str::FromStr for STVAnchor {
    type Err = String;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        match s {
            "text" => Ok(Self::Text),
            "margin" => Ok(Self::Margin),
            "page" => Ok(Self::Page),
            _ => Err(format!("unknown STVAnchor value: {}", s)),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum STHAnchor {
    #[serde(rename = "text")]
    Text,
    #[serde(rename = "margin")]
    Margin,
    #[serde(rename = "page")]
    Page,
}

impl std::fmt::Display for STHAnchor {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::Text => write!(f, "text"),
            Self::Margin => write!(f, "margin"),
            Self::Page => write!(f, "page"),
        }
    }
}

impl std::str::FromStr for STHAnchor {
    type Err = String;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        match s {
            "text" => Ok(Self::Text),
            "margin" => Ok(Self::Margin),
            "page" => Ok(Self::Page),
            _ => Err(format!("unknown STHAnchor value: {}", s)),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum STDropCap {
    #[serde(rename = "none")]
    None,
    #[serde(rename = "drop")]
    Drop,
    #[serde(rename = "margin")]
    Margin,
}

impl std::fmt::Display for STDropCap {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::None => write!(f, "none"),
            Self::Drop => write!(f, "drop"),
            Self::Margin => write!(f, "margin"),
        }
    }
}

impl std::str::FromStr for STDropCap {
    type Err = String;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        match s {
            "none" => Ok(Self::None),
            "drop" => Ok(Self::Drop),
            "margin" => Ok(Self::Margin),
            _ => Err(format!("unknown STDropCap value: {}", s)),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum STTabJc {
    #[serde(rename = "clear")]
    Clear,
    #[serde(rename = "start")]
    Start,
    #[serde(rename = "center")]
    Center,
    #[serde(rename = "end")]
    End,
    #[serde(rename = "decimal")]
    Decimal,
    #[serde(rename = "bar")]
    Bar,
    #[serde(rename = "num")]
    Num,
    #[serde(rename = "left")]
    Left,
    #[serde(rename = "right")]
    Right,
}

impl std::fmt::Display for STTabJc {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::Clear => write!(f, "clear"),
            Self::Start => write!(f, "start"),
            Self::Center => write!(f, "center"),
            Self::End => write!(f, "end"),
            Self::Decimal => write!(f, "decimal"),
            Self::Bar => write!(f, "bar"),
            Self::Num => write!(f, "num"),
            Self::Left => write!(f, "left"),
            Self::Right => write!(f, "right"),
        }
    }
}

impl std::str::FromStr for STTabJc {
    type Err = String;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        match s {
            "clear" => Ok(Self::Clear),
            "start" => Ok(Self::Start),
            "center" => Ok(Self::Center),
            "end" => Ok(Self::End),
            "decimal" => Ok(Self::Decimal),
            "bar" => Ok(Self::Bar),
            "num" => Ok(Self::Num),
            "left" => Ok(Self::Left),
            "right" => Ok(Self::Right),
            _ => Err(format!("unknown STTabJc value: {}", s)),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum STTabTlc {
    #[serde(rename = "none")]
    None,
    #[serde(rename = "dot")]
    Dot,
    #[serde(rename = "hyphen")]
    Hyphen,
    #[serde(rename = "underscore")]
    Underscore,
    #[serde(rename = "heavy")]
    Heavy,
    #[serde(rename = "middleDot")]
    MiddleDot,
}

impl std::fmt::Display for STTabTlc {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::None => write!(f, "none"),
            Self::Dot => write!(f, "dot"),
            Self::Hyphen => write!(f, "hyphen"),
            Self::Underscore => write!(f, "underscore"),
            Self::Heavy => write!(f, "heavy"),
            Self::MiddleDot => write!(f, "middleDot"),
        }
    }
}

impl std::str::FromStr for STTabTlc {
    type Err = String;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        match s {
            "none" => Ok(Self::None),
            "dot" => Ok(Self::Dot),
            "hyphen" => Ok(Self::Hyphen),
            "underscore" => Ok(Self::Underscore),
            "heavy" => Ok(Self::Heavy),
            "middleDot" => Ok(Self::MiddleDot),
            _ => Err(format!("unknown STTabTlc value: {}", s)),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum STLineSpacingRule {
    #[serde(rename = "auto")]
    Auto,
    #[serde(rename = "exact")]
    Exact,
    #[serde(rename = "atLeast")]
    AtLeast,
}

impl std::fmt::Display for STLineSpacingRule {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::Auto => write!(f, "auto"),
            Self::Exact => write!(f, "exact"),
            Self::AtLeast => write!(f, "atLeast"),
        }
    }
}

impl std::str::FromStr for STLineSpacingRule {
    type Err = String;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        match s {
            "auto" => Ok(Self::Auto),
            "exact" => Ok(Self::Exact),
            "atLeast" => Ok(Self::AtLeast),
            _ => Err(format!("unknown STLineSpacingRule value: {}", s)),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum STJc {
    #[serde(rename = "start")]
    Start,
    #[serde(rename = "center")]
    Center,
    #[serde(rename = "end")]
    End,
    #[serde(rename = "both")]
    Both,
    #[serde(rename = "mediumKashida")]
    MediumKashida,
    #[serde(rename = "distribute")]
    Distribute,
    #[serde(rename = "numTab")]
    NumTab,
    #[serde(rename = "highKashida")]
    HighKashida,
    #[serde(rename = "lowKashida")]
    LowKashida,
    #[serde(rename = "thaiDistribute")]
    ThaiDistribute,
    #[serde(rename = "left")]
    Left,
    #[serde(rename = "right")]
    Right,
}

impl std::fmt::Display for STJc {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::Start => write!(f, "start"),
            Self::Center => write!(f, "center"),
            Self::End => write!(f, "end"),
            Self::Both => write!(f, "both"),
            Self::MediumKashida => write!(f, "mediumKashida"),
            Self::Distribute => write!(f, "distribute"),
            Self::NumTab => write!(f, "numTab"),
            Self::HighKashida => write!(f, "highKashida"),
            Self::LowKashida => write!(f, "lowKashida"),
            Self::ThaiDistribute => write!(f, "thaiDistribute"),
            Self::Left => write!(f, "left"),
            Self::Right => write!(f, "right"),
        }
    }
}

impl std::str::FromStr for STJc {
    type Err = String;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        match s {
            "start" => Ok(Self::Start),
            "center" => Ok(Self::Center),
            "end" => Ok(Self::End),
            "both" => Ok(Self::Both),
            "mediumKashida" => Ok(Self::MediumKashida),
            "distribute" => Ok(Self::Distribute),
            "numTab" => Ok(Self::NumTab),
            "highKashida" => Ok(Self::HighKashida),
            "lowKashida" => Ok(Self::LowKashida),
            "thaiDistribute" => Ok(Self::ThaiDistribute),
            "left" => Ok(Self::Left),
            "right" => Ok(Self::Right),
            _ => Err(format!("unknown STJc value: {}", s)),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum STJcTable {
    #[serde(rename = "center")]
    Center,
    #[serde(rename = "end")]
    End,
    #[serde(rename = "left")]
    Left,
    #[serde(rename = "right")]
    Right,
    #[serde(rename = "start")]
    Start,
}

impl std::fmt::Display for STJcTable {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::Center => write!(f, "center"),
            Self::End => write!(f, "end"),
            Self::Left => write!(f, "left"),
            Self::Right => write!(f, "right"),
            Self::Start => write!(f, "start"),
        }
    }
}

impl std::str::FromStr for STJcTable {
    type Err = String;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        match s {
            "center" => Ok(Self::Center),
            "end" => Ok(Self::End),
            "left" => Ok(Self::Left),
            "right" => Ok(Self::Right),
            "start" => Ok(Self::Start),
            _ => Err(format!("unknown STJcTable value: {}", s)),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum STView {
    #[serde(rename = "none")]
    None,
    #[serde(rename = "print")]
    Print,
    #[serde(rename = "outline")]
    Outline,
    #[serde(rename = "masterPages")]
    MasterPages,
    #[serde(rename = "normal")]
    Normal,
    #[serde(rename = "web")]
    Web,
}

impl std::fmt::Display for STView {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::None => write!(f, "none"),
            Self::Print => write!(f, "print"),
            Self::Outline => write!(f, "outline"),
            Self::MasterPages => write!(f, "masterPages"),
            Self::Normal => write!(f, "normal"),
            Self::Web => write!(f, "web"),
        }
    }
}

impl std::str::FromStr for STView {
    type Err = String;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        match s {
            "none" => Ok(Self::None),
            "print" => Ok(Self::Print),
            "outline" => Ok(Self::Outline),
            "masterPages" => Ok(Self::MasterPages),
            "normal" => Ok(Self::Normal),
            "web" => Ok(Self::Web),
            _ => Err(format!("unknown STView value: {}", s)),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum STZoom {
    #[serde(rename = "none")]
    None,
    #[serde(rename = "fullPage")]
    FullPage,
    #[serde(rename = "bestFit")]
    BestFit,
    #[serde(rename = "textFit")]
    TextFit,
}

impl std::fmt::Display for STZoom {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::None => write!(f, "none"),
            Self::FullPage => write!(f, "fullPage"),
            Self::BestFit => write!(f, "bestFit"),
            Self::TextFit => write!(f, "textFit"),
        }
    }
}

impl std::str::FromStr for STZoom {
    type Err = String;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        match s {
            "none" => Ok(Self::None),
            "fullPage" => Ok(Self::FullPage),
            "bestFit" => Ok(Self::BestFit),
            "textFit" => Ok(Self::TextFit),
            _ => Err(format!("unknown STZoom value: {}", s)),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum STProof {
    #[serde(rename = "clean")]
    Clean,
    #[serde(rename = "dirty")]
    Dirty,
}

impl std::fmt::Display for STProof {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::Clean => write!(f, "clean"),
            Self::Dirty => write!(f, "dirty"),
        }
    }
}

impl std::str::FromStr for STProof {
    type Err = String;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        match s {
            "clean" => Ok(Self::Clean),
            "dirty" => Ok(Self::Dirty),
            _ => Err(format!("unknown STProof value: {}", s)),
        }
    }
}

pub type STDocType = String;

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum STDocProtect {
    #[serde(rename = "none")]
    None,
    #[serde(rename = "readOnly")]
    ReadOnly,
    #[serde(rename = "comments")]
    Comments,
    #[serde(rename = "trackedChanges")]
    TrackedChanges,
    #[serde(rename = "forms")]
    Forms,
}

impl std::fmt::Display for STDocProtect {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::None => write!(f, "none"),
            Self::ReadOnly => write!(f, "readOnly"),
            Self::Comments => write!(f, "comments"),
            Self::TrackedChanges => write!(f, "trackedChanges"),
            Self::Forms => write!(f, "forms"),
        }
    }
}

impl std::str::FromStr for STDocProtect {
    type Err = String;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        match s {
            "none" => Ok(Self::None),
            "readOnly" => Ok(Self::ReadOnly),
            "comments" => Ok(Self::Comments),
            "trackedChanges" => Ok(Self::TrackedChanges),
            "forms" => Ok(Self::Forms),
            _ => Err(format!("unknown STDocProtect value: {}", s)),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum STMailMergeDocType {
    #[serde(rename = "catalog")]
    Catalog,
    #[serde(rename = "envelopes")]
    Envelopes,
    #[serde(rename = "mailingLabels")]
    MailingLabels,
    #[serde(rename = "formLetters")]
    FormLetters,
    #[serde(rename = "email")]
    Email,
    #[serde(rename = "fax")]
    Fax,
}

impl std::fmt::Display for STMailMergeDocType {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::Catalog => write!(f, "catalog"),
            Self::Envelopes => write!(f, "envelopes"),
            Self::MailingLabels => write!(f, "mailingLabels"),
            Self::FormLetters => write!(f, "formLetters"),
            Self::Email => write!(f, "email"),
            Self::Fax => write!(f, "fax"),
        }
    }
}

impl std::str::FromStr for STMailMergeDocType {
    type Err = String;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        match s {
            "catalog" => Ok(Self::Catalog),
            "envelopes" => Ok(Self::Envelopes),
            "mailingLabels" => Ok(Self::MailingLabels),
            "formLetters" => Ok(Self::FormLetters),
            "email" => Ok(Self::Email),
            "fax" => Ok(Self::Fax),
            _ => Err(format!("unknown STMailMergeDocType value: {}", s)),
        }
    }
}

pub type STMailMergeDataType = String;

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum STMailMergeDest {
    #[serde(rename = "newDocument")]
    NewDocument,
    #[serde(rename = "printer")]
    Printer,
    #[serde(rename = "email")]
    Email,
    #[serde(rename = "fax")]
    Fax,
}

impl std::fmt::Display for STMailMergeDest {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::NewDocument => write!(f, "newDocument"),
            Self::Printer => write!(f, "printer"),
            Self::Email => write!(f, "email"),
            Self::Fax => write!(f, "fax"),
        }
    }
}

impl std::str::FromStr for STMailMergeDest {
    type Err = String;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        match s {
            "newDocument" => Ok(Self::NewDocument),
            "printer" => Ok(Self::Printer),
            "email" => Ok(Self::Email),
            "fax" => Ok(Self::Fax),
            _ => Err(format!("unknown STMailMergeDest value: {}", s)),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum STMailMergeOdsoFMDFieldType {
    #[serde(rename = "null")]
    Null,
    #[serde(rename = "dbColumn")]
    DbColumn,
}

impl std::fmt::Display for STMailMergeOdsoFMDFieldType {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::Null => write!(f, "null"),
            Self::DbColumn => write!(f, "dbColumn"),
        }
    }
}

impl std::str::FromStr for STMailMergeOdsoFMDFieldType {
    type Err = String;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        match s {
            "null" => Ok(Self::Null),
            "dbColumn" => Ok(Self::DbColumn),
            _ => Err(format!("unknown STMailMergeOdsoFMDFieldType value: {}", s)),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum STTextDirection {
    #[serde(rename = "tb")]
    Tb,
    #[serde(rename = "rl")]
    Rl,
    #[serde(rename = "lr")]
    Lr,
    #[serde(rename = "tbV")]
    TbV,
    #[serde(rename = "rlV")]
    RlV,
    #[serde(rename = "lrV")]
    LrV,
    #[serde(rename = "btLr")]
    BtLr,
    #[serde(rename = "lrTb")]
    LrTb,
    #[serde(rename = "lrTbV")]
    LrTbV,
    #[serde(rename = "tbLrV")]
    TbLrV,
    #[serde(rename = "tbRl")]
    TbRl,
    #[serde(rename = "tbRlV")]
    TbRlV,
}

impl std::fmt::Display for STTextDirection {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::Tb => write!(f, "tb"),
            Self::Rl => write!(f, "rl"),
            Self::Lr => write!(f, "lr"),
            Self::TbV => write!(f, "tbV"),
            Self::RlV => write!(f, "rlV"),
            Self::LrV => write!(f, "lrV"),
            Self::BtLr => write!(f, "btLr"),
            Self::LrTb => write!(f, "lrTb"),
            Self::LrTbV => write!(f, "lrTbV"),
            Self::TbLrV => write!(f, "tbLrV"),
            Self::TbRl => write!(f, "tbRl"),
            Self::TbRlV => write!(f, "tbRlV"),
        }
    }
}

impl std::str::FromStr for STTextDirection {
    type Err = String;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        match s {
            "tb" => Ok(Self::Tb),
            "rl" => Ok(Self::Rl),
            "lr" => Ok(Self::Lr),
            "tbV" => Ok(Self::TbV),
            "rlV" => Ok(Self::RlV),
            "lrV" => Ok(Self::LrV),
            "btLr" => Ok(Self::BtLr),
            "lrTb" => Ok(Self::LrTb),
            "lrTbV" => Ok(Self::LrTbV),
            "tbLrV" => Ok(Self::TbLrV),
            "tbRl" => Ok(Self::TbRl),
            "tbRlV" => Ok(Self::TbRlV),
            _ => Err(format!("unknown STTextDirection value: {}", s)),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum STTextAlignment {
    #[serde(rename = "top")]
    Top,
    #[serde(rename = "center")]
    Center,
    #[serde(rename = "baseline")]
    Baseline,
    #[serde(rename = "bottom")]
    Bottom,
    #[serde(rename = "auto")]
    Auto,
}

impl std::fmt::Display for STTextAlignment {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::Top => write!(f, "top"),
            Self::Center => write!(f, "center"),
            Self::Baseline => write!(f, "baseline"),
            Self::Bottom => write!(f, "bottom"),
            Self::Auto => write!(f, "auto"),
        }
    }
}

impl std::str::FromStr for STTextAlignment {
    type Err = String;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        match s {
            "top" => Ok(Self::Top),
            "center" => Ok(Self::Center),
            "baseline" => Ok(Self::Baseline),
            "bottom" => Ok(Self::Bottom),
            "auto" => Ok(Self::Auto),
            _ => Err(format!("unknown STTextAlignment value: {}", s)),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum STDisplacedByCustomXml {
    #[serde(rename = "next")]
    Next,
    #[serde(rename = "prev")]
    Prev,
}

impl std::fmt::Display for STDisplacedByCustomXml {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::Next => write!(f, "next"),
            Self::Prev => write!(f, "prev"),
        }
    }
}

impl std::str::FromStr for STDisplacedByCustomXml {
    type Err = String;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        match s {
            "next" => Ok(Self::Next),
            "prev" => Ok(Self::Prev),
            _ => Err(format!("unknown STDisplacedByCustomXml value: {}", s)),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum STAnnotationVMerge {
    #[serde(rename = "cont")]
    Cont,
    #[serde(rename = "rest")]
    Rest,
}

impl std::fmt::Display for STAnnotationVMerge {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::Cont => write!(f, "cont"),
            Self::Rest => write!(f, "rest"),
        }
    }
}

impl std::str::FromStr for STAnnotationVMerge {
    type Err = String;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        match s {
            "cont" => Ok(Self::Cont),
            "rest" => Ok(Self::Rest),
            _ => Err(format!("unknown STAnnotationVMerge value: {}", s)),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum STTextboxTightWrap {
    #[serde(rename = "none")]
    None,
    #[serde(rename = "allLines")]
    AllLines,
    #[serde(rename = "firstAndLastLine")]
    FirstAndLastLine,
    #[serde(rename = "firstLineOnly")]
    FirstLineOnly,
    #[serde(rename = "lastLineOnly")]
    LastLineOnly,
}

impl std::fmt::Display for STTextboxTightWrap {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::None => write!(f, "none"),
            Self::AllLines => write!(f, "allLines"),
            Self::FirstAndLastLine => write!(f, "firstAndLastLine"),
            Self::FirstLineOnly => write!(f, "firstLineOnly"),
            Self::LastLineOnly => write!(f, "lastLineOnly"),
        }
    }
}

impl std::str::FromStr for STTextboxTightWrap {
    type Err = String;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        match s {
            "none" => Ok(Self::None),
            "allLines" => Ok(Self::AllLines),
            "firstAndLastLine" => Ok(Self::FirstAndLastLine),
            "firstLineOnly" => Ok(Self::FirstLineOnly),
            "lastLineOnly" => Ok(Self::LastLineOnly),
            _ => Err(format!("unknown STTextboxTightWrap value: {}", s)),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum STObjectDrawAspect {
    #[serde(rename = "content")]
    Content,
    #[serde(rename = "icon")]
    Icon,
}

impl std::fmt::Display for STObjectDrawAspect {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::Content => write!(f, "content"),
            Self::Icon => write!(f, "icon"),
        }
    }
}

impl std::str::FromStr for STObjectDrawAspect {
    type Err = String;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        match s {
            "content" => Ok(Self::Content),
            "icon" => Ok(Self::Icon),
            _ => Err(format!("unknown STObjectDrawAspect value: {}", s)),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum STObjectUpdateMode {
    #[serde(rename = "always")]
    Always,
    #[serde(rename = "onCall")]
    OnCall,
}

impl std::fmt::Display for STObjectUpdateMode {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::Always => write!(f, "always"),
            Self::OnCall => write!(f, "onCall"),
        }
    }
}

impl std::str::FromStr for STObjectUpdateMode {
    type Err = String;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        match s {
            "always" => Ok(Self::Always),
            "onCall" => Ok(Self::OnCall),
            _ => Err(format!("unknown STObjectUpdateMode value: {}", s)),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum STFldCharType {
    #[serde(rename = "begin")]
    Begin,
    #[serde(rename = "separate")]
    Separate,
    #[serde(rename = "end")]
    End,
}

impl std::fmt::Display for STFldCharType {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::Begin => write!(f, "begin"),
            Self::Separate => write!(f, "separate"),
            Self::End => write!(f, "end"),
        }
    }
}

impl std::str::FromStr for STFldCharType {
    type Err = String;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        match s {
            "begin" => Ok(Self::Begin),
            "separate" => Ok(Self::Separate),
            "end" => Ok(Self::End),
            _ => Err(format!("unknown STFldCharType value: {}", s)),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum STInfoTextType {
    #[serde(rename = "text")]
    Text,
    #[serde(rename = "autoText")]
    AutoText,
}

impl std::fmt::Display for STInfoTextType {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::Text => write!(f, "text"),
            Self::AutoText => write!(f, "autoText"),
        }
    }
}

impl std::str::FromStr for STInfoTextType {
    type Err = String;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        match s {
            "text" => Ok(Self::Text),
            "autoText" => Ok(Self::AutoText),
            _ => Err(format!("unknown STInfoTextType value: {}", s)),
        }
    }
}

pub type STFFHelpTextVal = String;

pub type STFFStatusTextVal = String;

pub type STFFName = String;

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum STFFTextType {
    #[serde(rename = "regular")]
    Regular,
    #[serde(rename = "number")]
    Number,
    #[serde(rename = "date")]
    Date,
    #[serde(rename = "currentTime")]
    CurrentTime,
    #[serde(rename = "currentDate")]
    CurrentDate,
    #[serde(rename = "calculated")]
    Calculated,
}

impl std::fmt::Display for STFFTextType {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::Regular => write!(f, "regular"),
            Self::Number => write!(f, "number"),
            Self::Date => write!(f, "date"),
            Self::CurrentTime => write!(f, "currentTime"),
            Self::CurrentDate => write!(f, "currentDate"),
            Self::Calculated => write!(f, "calculated"),
        }
    }
}

impl std::str::FromStr for STFFTextType {
    type Err = String;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        match s {
            "regular" => Ok(Self::Regular),
            "number" => Ok(Self::Number),
            "date" => Ok(Self::Date),
            "currentTime" => Ok(Self::CurrentTime),
            "currentDate" => Ok(Self::CurrentDate),
            "calculated" => Ok(Self::Calculated),
            _ => Err(format!("unknown STFFTextType value: {}", s)),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum STSectionMark {
    #[serde(rename = "nextPage")]
    NextPage,
    #[serde(rename = "nextColumn")]
    NextColumn,
    #[serde(rename = "continuous")]
    Continuous,
    #[serde(rename = "evenPage")]
    EvenPage,
    #[serde(rename = "oddPage")]
    OddPage,
}

impl std::fmt::Display for STSectionMark {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::NextPage => write!(f, "nextPage"),
            Self::NextColumn => write!(f, "nextColumn"),
            Self::Continuous => write!(f, "continuous"),
            Self::EvenPage => write!(f, "evenPage"),
            Self::OddPage => write!(f, "oddPage"),
        }
    }
}

impl std::str::FromStr for STSectionMark {
    type Err = String;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        match s {
            "nextPage" => Ok(Self::NextPage),
            "nextColumn" => Ok(Self::NextColumn),
            "continuous" => Ok(Self::Continuous),
            "evenPage" => Ok(Self::EvenPage),
            "oddPage" => Ok(Self::OddPage),
            _ => Err(format!("unknown STSectionMark value: {}", s)),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum STNumberFormat {
    #[serde(rename = "decimal")]
    Decimal,
    #[serde(rename = "upperRoman")]
    UpperRoman,
    #[serde(rename = "lowerRoman")]
    LowerRoman,
    #[serde(rename = "upperLetter")]
    UpperLetter,
    #[serde(rename = "lowerLetter")]
    LowerLetter,
    #[serde(rename = "ordinal")]
    Ordinal,
    #[serde(rename = "cardinalText")]
    CardinalText,
    #[serde(rename = "ordinalText")]
    OrdinalText,
    #[serde(rename = "hex")]
    Hex,
    #[serde(rename = "chicago")]
    Chicago,
    #[serde(rename = "ideographDigital")]
    IdeographDigital,
    #[serde(rename = "japaneseCounting")]
    JapaneseCounting,
    #[serde(rename = "aiueo")]
    Aiueo,
    #[serde(rename = "iroha")]
    Iroha,
    #[serde(rename = "decimalFullWidth")]
    DecimalFullWidth,
    #[serde(rename = "decimalHalfWidth")]
    DecimalHalfWidth,
    #[serde(rename = "japaneseLegal")]
    JapaneseLegal,
    #[serde(rename = "japaneseDigitalTenThousand")]
    JapaneseDigitalTenThousand,
    #[serde(rename = "decimalEnclosedCircle")]
    DecimalEnclosedCircle,
    #[serde(rename = "decimalFullWidth2")]
    DecimalFullWidth2,
    #[serde(rename = "aiueoFullWidth")]
    AiueoFullWidth,
    #[serde(rename = "irohaFullWidth")]
    IrohaFullWidth,
    #[serde(rename = "decimalZero")]
    DecimalZero,
    #[serde(rename = "bullet")]
    Bullet,
    #[serde(rename = "ganada")]
    Ganada,
    #[serde(rename = "chosung")]
    Chosung,
    #[serde(rename = "decimalEnclosedFullstop")]
    DecimalEnclosedFullstop,
    #[serde(rename = "decimalEnclosedParen")]
    DecimalEnclosedParen,
    #[serde(rename = "decimalEnclosedCircleChinese")]
    DecimalEnclosedCircleChinese,
    #[serde(rename = "ideographEnclosedCircle")]
    IdeographEnclosedCircle,
    #[serde(rename = "ideographTraditional")]
    IdeographTraditional,
    #[serde(rename = "ideographZodiac")]
    IdeographZodiac,
    #[serde(rename = "ideographZodiacTraditional")]
    IdeographZodiacTraditional,
    #[serde(rename = "taiwaneseCounting")]
    TaiwaneseCounting,
    #[serde(rename = "ideographLegalTraditional")]
    IdeographLegalTraditional,
    #[serde(rename = "taiwaneseCountingThousand")]
    TaiwaneseCountingThousand,
    #[serde(rename = "taiwaneseDigital")]
    TaiwaneseDigital,
    #[serde(rename = "chineseCounting")]
    ChineseCounting,
    #[serde(rename = "chineseLegalSimplified")]
    ChineseLegalSimplified,
    #[serde(rename = "chineseCountingThousand")]
    ChineseCountingThousand,
    #[serde(rename = "koreanDigital")]
    KoreanDigital,
    #[serde(rename = "koreanCounting")]
    KoreanCounting,
    #[serde(rename = "koreanLegal")]
    KoreanLegal,
    #[serde(rename = "koreanDigital2")]
    KoreanDigital2,
    #[serde(rename = "vietnameseCounting")]
    VietnameseCounting,
    #[serde(rename = "russianLower")]
    RussianLower,
    #[serde(rename = "russianUpper")]
    RussianUpper,
    #[serde(rename = "none")]
    None,
    #[serde(rename = "numberInDash")]
    NumberInDash,
    #[serde(rename = "hebrew1")]
    Hebrew1,
    #[serde(rename = "hebrew2")]
    Hebrew2,
    #[serde(rename = "arabicAlpha")]
    ArabicAlpha,
    #[serde(rename = "arabicAbjad")]
    ArabicAbjad,
    #[serde(rename = "hindiVowels")]
    HindiVowels,
    #[serde(rename = "hindiConsonants")]
    HindiConsonants,
    #[serde(rename = "hindiNumbers")]
    HindiNumbers,
    #[serde(rename = "hindiCounting")]
    HindiCounting,
    #[serde(rename = "thaiLetters")]
    ThaiLetters,
    #[serde(rename = "thaiNumbers")]
    ThaiNumbers,
    #[serde(rename = "thaiCounting")]
    ThaiCounting,
    #[serde(rename = "bahtText")]
    BahtText,
    #[serde(rename = "dollarText")]
    DollarText,
    #[serde(rename = "custom")]
    Custom,
}

impl std::fmt::Display for STNumberFormat {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::Decimal => write!(f, "decimal"),
            Self::UpperRoman => write!(f, "upperRoman"),
            Self::LowerRoman => write!(f, "lowerRoman"),
            Self::UpperLetter => write!(f, "upperLetter"),
            Self::LowerLetter => write!(f, "lowerLetter"),
            Self::Ordinal => write!(f, "ordinal"),
            Self::CardinalText => write!(f, "cardinalText"),
            Self::OrdinalText => write!(f, "ordinalText"),
            Self::Hex => write!(f, "hex"),
            Self::Chicago => write!(f, "chicago"),
            Self::IdeographDigital => write!(f, "ideographDigital"),
            Self::JapaneseCounting => write!(f, "japaneseCounting"),
            Self::Aiueo => write!(f, "aiueo"),
            Self::Iroha => write!(f, "iroha"),
            Self::DecimalFullWidth => write!(f, "decimalFullWidth"),
            Self::DecimalHalfWidth => write!(f, "decimalHalfWidth"),
            Self::JapaneseLegal => write!(f, "japaneseLegal"),
            Self::JapaneseDigitalTenThousand => write!(f, "japaneseDigitalTenThousand"),
            Self::DecimalEnclosedCircle => write!(f, "decimalEnclosedCircle"),
            Self::DecimalFullWidth2 => write!(f, "decimalFullWidth2"),
            Self::AiueoFullWidth => write!(f, "aiueoFullWidth"),
            Self::IrohaFullWidth => write!(f, "irohaFullWidth"),
            Self::DecimalZero => write!(f, "decimalZero"),
            Self::Bullet => write!(f, "bullet"),
            Self::Ganada => write!(f, "ganada"),
            Self::Chosung => write!(f, "chosung"),
            Self::DecimalEnclosedFullstop => write!(f, "decimalEnclosedFullstop"),
            Self::DecimalEnclosedParen => write!(f, "decimalEnclosedParen"),
            Self::DecimalEnclosedCircleChinese => write!(f, "decimalEnclosedCircleChinese"),
            Self::IdeographEnclosedCircle => write!(f, "ideographEnclosedCircle"),
            Self::IdeographTraditional => write!(f, "ideographTraditional"),
            Self::IdeographZodiac => write!(f, "ideographZodiac"),
            Self::IdeographZodiacTraditional => write!(f, "ideographZodiacTraditional"),
            Self::TaiwaneseCounting => write!(f, "taiwaneseCounting"),
            Self::IdeographLegalTraditional => write!(f, "ideographLegalTraditional"),
            Self::TaiwaneseCountingThousand => write!(f, "taiwaneseCountingThousand"),
            Self::TaiwaneseDigital => write!(f, "taiwaneseDigital"),
            Self::ChineseCounting => write!(f, "chineseCounting"),
            Self::ChineseLegalSimplified => write!(f, "chineseLegalSimplified"),
            Self::ChineseCountingThousand => write!(f, "chineseCountingThousand"),
            Self::KoreanDigital => write!(f, "koreanDigital"),
            Self::KoreanCounting => write!(f, "koreanCounting"),
            Self::KoreanLegal => write!(f, "koreanLegal"),
            Self::KoreanDigital2 => write!(f, "koreanDigital2"),
            Self::VietnameseCounting => write!(f, "vietnameseCounting"),
            Self::RussianLower => write!(f, "russianLower"),
            Self::RussianUpper => write!(f, "russianUpper"),
            Self::None => write!(f, "none"),
            Self::NumberInDash => write!(f, "numberInDash"),
            Self::Hebrew1 => write!(f, "hebrew1"),
            Self::Hebrew2 => write!(f, "hebrew2"),
            Self::ArabicAlpha => write!(f, "arabicAlpha"),
            Self::ArabicAbjad => write!(f, "arabicAbjad"),
            Self::HindiVowels => write!(f, "hindiVowels"),
            Self::HindiConsonants => write!(f, "hindiConsonants"),
            Self::HindiNumbers => write!(f, "hindiNumbers"),
            Self::HindiCounting => write!(f, "hindiCounting"),
            Self::ThaiLetters => write!(f, "thaiLetters"),
            Self::ThaiNumbers => write!(f, "thaiNumbers"),
            Self::ThaiCounting => write!(f, "thaiCounting"),
            Self::BahtText => write!(f, "bahtText"),
            Self::DollarText => write!(f, "dollarText"),
            Self::Custom => write!(f, "custom"),
        }
    }
}

impl std::str::FromStr for STNumberFormat {
    type Err = String;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        match s {
            "decimal" => Ok(Self::Decimal),
            "upperRoman" => Ok(Self::UpperRoman),
            "lowerRoman" => Ok(Self::LowerRoman),
            "upperLetter" => Ok(Self::UpperLetter),
            "lowerLetter" => Ok(Self::LowerLetter),
            "ordinal" => Ok(Self::Ordinal),
            "cardinalText" => Ok(Self::CardinalText),
            "ordinalText" => Ok(Self::OrdinalText),
            "hex" => Ok(Self::Hex),
            "chicago" => Ok(Self::Chicago),
            "ideographDigital" => Ok(Self::IdeographDigital),
            "japaneseCounting" => Ok(Self::JapaneseCounting),
            "aiueo" => Ok(Self::Aiueo),
            "iroha" => Ok(Self::Iroha),
            "decimalFullWidth" => Ok(Self::DecimalFullWidth),
            "decimalHalfWidth" => Ok(Self::DecimalHalfWidth),
            "japaneseLegal" => Ok(Self::JapaneseLegal),
            "japaneseDigitalTenThousand" => Ok(Self::JapaneseDigitalTenThousand),
            "decimalEnclosedCircle" => Ok(Self::DecimalEnclosedCircle),
            "decimalFullWidth2" => Ok(Self::DecimalFullWidth2),
            "aiueoFullWidth" => Ok(Self::AiueoFullWidth),
            "irohaFullWidth" => Ok(Self::IrohaFullWidth),
            "decimalZero" => Ok(Self::DecimalZero),
            "bullet" => Ok(Self::Bullet),
            "ganada" => Ok(Self::Ganada),
            "chosung" => Ok(Self::Chosung),
            "decimalEnclosedFullstop" => Ok(Self::DecimalEnclosedFullstop),
            "decimalEnclosedParen" => Ok(Self::DecimalEnclosedParen),
            "decimalEnclosedCircleChinese" => Ok(Self::DecimalEnclosedCircleChinese),
            "ideographEnclosedCircle" => Ok(Self::IdeographEnclosedCircle),
            "ideographTraditional" => Ok(Self::IdeographTraditional),
            "ideographZodiac" => Ok(Self::IdeographZodiac),
            "ideographZodiacTraditional" => Ok(Self::IdeographZodiacTraditional),
            "taiwaneseCounting" => Ok(Self::TaiwaneseCounting),
            "ideographLegalTraditional" => Ok(Self::IdeographLegalTraditional),
            "taiwaneseCountingThousand" => Ok(Self::TaiwaneseCountingThousand),
            "taiwaneseDigital" => Ok(Self::TaiwaneseDigital),
            "chineseCounting" => Ok(Self::ChineseCounting),
            "chineseLegalSimplified" => Ok(Self::ChineseLegalSimplified),
            "chineseCountingThousand" => Ok(Self::ChineseCountingThousand),
            "koreanDigital" => Ok(Self::KoreanDigital),
            "koreanCounting" => Ok(Self::KoreanCounting),
            "koreanLegal" => Ok(Self::KoreanLegal),
            "koreanDigital2" => Ok(Self::KoreanDigital2),
            "vietnameseCounting" => Ok(Self::VietnameseCounting),
            "russianLower" => Ok(Self::RussianLower),
            "russianUpper" => Ok(Self::RussianUpper),
            "none" => Ok(Self::None),
            "numberInDash" => Ok(Self::NumberInDash),
            "hebrew1" => Ok(Self::Hebrew1),
            "hebrew2" => Ok(Self::Hebrew2),
            "arabicAlpha" => Ok(Self::ArabicAlpha),
            "arabicAbjad" => Ok(Self::ArabicAbjad),
            "hindiVowels" => Ok(Self::HindiVowels),
            "hindiConsonants" => Ok(Self::HindiConsonants),
            "hindiNumbers" => Ok(Self::HindiNumbers),
            "hindiCounting" => Ok(Self::HindiCounting),
            "thaiLetters" => Ok(Self::ThaiLetters),
            "thaiNumbers" => Ok(Self::ThaiNumbers),
            "thaiCounting" => Ok(Self::ThaiCounting),
            "bahtText" => Ok(Self::BahtText),
            "dollarText" => Ok(Self::DollarText),
            "custom" => Ok(Self::Custom),
            _ => Err(format!("unknown STNumberFormat value: {}", s)),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum STPageOrientation {
    #[serde(rename = "portrait")]
    Portrait,
    #[serde(rename = "landscape")]
    Landscape,
}

impl std::fmt::Display for STPageOrientation {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::Portrait => write!(f, "portrait"),
            Self::Landscape => write!(f, "landscape"),
        }
    }
}

impl std::str::FromStr for STPageOrientation {
    type Err = String;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        match s {
            "portrait" => Ok(Self::Portrait),
            "landscape" => Ok(Self::Landscape),
            _ => Err(format!("unknown STPageOrientation value: {}", s)),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum STPageBorderZOrder {
    #[serde(rename = "front")]
    Front,
    #[serde(rename = "back")]
    Back,
}

impl std::fmt::Display for STPageBorderZOrder {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::Front => write!(f, "front"),
            Self::Back => write!(f, "back"),
        }
    }
}

impl std::str::FromStr for STPageBorderZOrder {
    type Err = String;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        match s {
            "front" => Ok(Self::Front),
            "back" => Ok(Self::Back),
            _ => Err(format!("unknown STPageBorderZOrder value: {}", s)),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum STPageBorderDisplay {
    #[serde(rename = "allPages")]
    AllPages,
    #[serde(rename = "firstPage")]
    FirstPage,
    #[serde(rename = "notFirstPage")]
    NotFirstPage,
}

impl std::fmt::Display for STPageBorderDisplay {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::AllPages => write!(f, "allPages"),
            Self::FirstPage => write!(f, "firstPage"),
            Self::NotFirstPage => write!(f, "notFirstPage"),
        }
    }
}

impl std::str::FromStr for STPageBorderDisplay {
    type Err = String;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        match s {
            "allPages" => Ok(Self::AllPages),
            "firstPage" => Ok(Self::FirstPage),
            "notFirstPage" => Ok(Self::NotFirstPage),
            _ => Err(format!("unknown STPageBorderDisplay value: {}", s)),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum STPageBorderOffset {
    #[serde(rename = "page")]
    Page,
    #[serde(rename = "text")]
    Text,
}

impl std::fmt::Display for STPageBorderOffset {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::Page => write!(f, "page"),
            Self::Text => write!(f, "text"),
        }
    }
}

impl std::str::FromStr for STPageBorderOffset {
    type Err = String;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        match s {
            "page" => Ok(Self::Page),
            "text" => Ok(Self::Text),
            _ => Err(format!("unknown STPageBorderOffset value: {}", s)),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum STChapterSep {
    #[serde(rename = "hyphen")]
    Hyphen,
    #[serde(rename = "period")]
    Period,
    #[serde(rename = "colon")]
    Colon,
    #[serde(rename = "emDash")]
    EmDash,
    #[serde(rename = "enDash")]
    EnDash,
}

impl std::fmt::Display for STChapterSep {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::Hyphen => write!(f, "hyphen"),
            Self::Period => write!(f, "period"),
            Self::Colon => write!(f, "colon"),
            Self::EmDash => write!(f, "emDash"),
            Self::EnDash => write!(f, "enDash"),
        }
    }
}

impl std::str::FromStr for STChapterSep {
    type Err = String;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        match s {
            "hyphen" => Ok(Self::Hyphen),
            "period" => Ok(Self::Period),
            "colon" => Ok(Self::Colon),
            "emDash" => Ok(Self::EmDash),
            "enDash" => Ok(Self::EnDash),
            _ => Err(format!("unknown STChapterSep value: {}", s)),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum STLineNumberRestart {
    #[serde(rename = "newPage")]
    NewPage,
    #[serde(rename = "newSection")]
    NewSection,
    #[serde(rename = "continuous")]
    Continuous,
}

impl std::fmt::Display for STLineNumberRestart {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::NewPage => write!(f, "newPage"),
            Self::NewSection => write!(f, "newSection"),
            Self::Continuous => write!(f, "continuous"),
        }
    }
}

impl std::str::FromStr for STLineNumberRestart {
    type Err = String;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        match s {
            "newPage" => Ok(Self::NewPage),
            "newSection" => Ok(Self::NewSection),
            "continuous" => Ok(Self::Continuous),
            _ => Err(format!("unknown STLineNumberRestart value: {}", s)),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum STVerticalJc {
    #[serde(rename = "top")]
    Top,
    #[serde(rename = "center")]
    Center,
    #[serde(rename = "both")]
    Both,
    #[serde(rename = "bottom")]
    Bottom,
}

impl std::fmt::Display for STVerticalJc {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::Top => write!(f, "top"),
            Self::Center => write!(f, "center"),
            Self::Both => write!(f, "both"),
            Self::Bottom => write!(f, "bottom"),
        }
    }
}

impl std::str::FromStr for STVerticalJc {
    type Err = String;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        match s {
            "top" => Ok(Self::Top),
            "center" => Ok(Self::Center),
            "both" => Ok(Self::Both),
            "bottom" => Ok(Self::Bottom),
            _ => Err(format!("unknown STVerticalJc value: {}", s)),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum STDocGrid {
    #[serde(rename = "default")]
    Default,
    #[serde(rename = "lines")]
    Lines,
    #[serde(rename = "linesAndChars")]
    LinesAndChars,
    #[serde(rename = "snapToChars")]
    SnapToChars,
}

impl std::fmt::Display for STDocGrid {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::Default => write!(f, "default"),
            Self::Lines => write!(f, "lines"),
            Self::LinesAndChars => write!(f, "linesAndChars"),
            Self::SnapToChars => write!(f, "snapToChars"),
        }
    }
}

impl std::str::FromStr for STDocGrid {
    type Err = String;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        match s {
            "default" => Ok(Self::Default),
            "lines" => Ok(Self::Lines),
            "linesAndChars" => Ok(Self::LinesAndChars),
            "snapToChars" => Ok(Self::SnapToChars),
            _ => Err(format!("unknown STDocGrid value: {}", s)),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum STHdrFtr {
    #[serde(rename = "even")]
    Even,
    #[serde(rename = "default")]
    Default,
    #[serde(rename = "first")]
    First,
}

impl std::fmt::Display for STHdrFtr {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::Even => write!(f, "even"),
            Self::Default => write!(f, "default"),
            Self::First => write!(f, "first"),
        }
    }
}

impl std::str::FromStr for STHdrFtr {
    type Err = String;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        match s {
            "even" => Ok(Self::Even),
            "default" => Ok(Self::Default),
            "first" => Ok(Self::First),
            _ => Err(format!("unknown STHdrFtr value: {}", s)),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum STFtnEdn {
    #[serde(rename = "normal")]
    Normal,
    #[serde(rename = "separator")]
    Separator,
    #[serde(rename = "continuationSeparator")]
    ContinuationSeparator,
    #[serde(rename = "continuationNotice")]
    ContinuationNotice,
}

impl std::fmt::Display for STFtnEdn {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::Normal => write!(f, "normal"),
            Self::Separator => write!(f, "separator"),
            Self::ContinuationSeparator => write!(f, "continuationSeparator"),
            Self::ContinuationNotice => write!(f, "continuationNotice"),
        }
    }
}

impl std::str::FromStr for STFtnEdn {
    type Err = String;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        match s {
            "normal" => Ok(Self::Normal),
            "separator" => Ok(Self::Separator),
            "continuationSeparator" => Ok(Self::ContinuationSeparator),
            "continuationNotice" => Ok(Self::ContinuationNotice),
            _ => Err(format!("unknown STFtnEdn value: {}", s)),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum STBrType {
    #[serde(rename = "page")]
    Page,
    #[serde(rename = "column")]
    Column,
    #[serde(rename = "textWrapping")]
    TextWrapping,
}

impl std::fmt::Display for STBrType {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::Page => write!(f, "page"),
            Self::Column => write!(f, "column"),
            Self::TextWrapping => write!(f, "textWrapping"),
        }
    }
}

impl std::str::FromStr for STBrType {
    type Err = String;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        match s {
            "page" => Ok(Self::Page),
            "column" => Ok(Self::Column),
            "textWrapping" => Ok(Self::TextWrapping),
            _ => Err(format!("unknown STBrType value: {}", s)),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum STBrClear {
    #[serde(rename = "none")]
    None,
    #[serde(rename = "left")]
    Left,
    #[serde(rename = "right")]
    Right,
    #[serde(rename = "all")]
    All,
}

impl std::fmt::Display for STBrClear {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::None => write!(f, "none"),
            Self::Left => write!(f, "left"),
            Self::Right => write!(f, "right"),
            Self::All => write!(f, "all"),
        }
    }
}

impl std::str::FromStr for STBrClear {
    type Err = String;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        match s {
            "none" => Ok(Self::None),
            "left" => Ok(Self::Left),
            "right" => Ok(Self::Right),
            "all" => Ok(Self::All),
            _ => Err(format!("unknown STBrClear value: {}", s)),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum STPTabAlignment {
    #[serde(rename = "left")]
    Left,
    #[serde(rename = "center")]
    Center,
    #[serde(rename = "right")]
    Right,
}

impl std::fmt::Display for STPTabAlignment {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::Left => write!(f, "left"),
            Self::Center => write!(f, "center"),
            Self::Right => write!(f, "right"),
        }
    }
}

impl std::str::FromStr for STPTabAlignment {
    type Err = String;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        match s {
            "left" => Ok(Self::Left),
            "center" => Ok(Self::Center),
            "right" => Ok(Self::Right),
            _ => Err(format!("unknown STPTabAlignment value: {}", s)),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum STPTabRelativeTo {
    #[serde(rename = "margin")]
    Margin,
    #[serde(rename = "indent")]
    Indent,
}

impl std::fmt::Display for STPTabRelativeTo {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::Margin => write!(f, "margin"),
            Self::Indent => write!(f, "indent"),
        }
    }
}

impl std::str::FromStr for STPTabRelativeTo {
    type Err = String;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        match s {
            "margin" => Ok(Self::Margin),
            "indent" => Ok(Self::Indent),
            _ => Err(format!("unknown STPTabRelativeTo value: {}", s)),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum STPTabLeader {
    #[serde(rename = "none")]
    None,
    #[serde(rename = "dot")]
    Dot,
    #[serde(rename = "hyphen")]
    Hyphen,
    #[serde(rename = "underscore")]
    Underscore,
    #[serde(rename = "middleDot")]
    MiddleDot,
}

impl std::fmt::Display for STPTabLeader {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::None => write!(f, "none"),
            Self::Dot => write!(f, "dot"),
            Self::Hyphen => write!(f, "hyphen"),
            Self::Underscore => write!(f, "underscore"),
            Self::MiddleDot => write!(f, "middleDot"),
        }
    }
}

impl std::str::FromStr for STPTabLeader {
    type Err = String;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        match s {
            "none" => Ok(Self::None),
            "dot" => Ok(Self::Dot),
            "hyphen" => Ok(Self::Hyphen),
            "underscore" => Ok(Self::Underscore),
            "middleDot" => Ok(Self::MiddleDot),
            _ => Err(format!("unknown STPTabLeader value: {}", s)),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum STProofErr {
    #[serde(rename = "spellStart")]
    SpellStart,
    #[serde(rename = "spellEnd")]
    SpellEnd,
    #[serde(rename = "gramStart")]
    GramStart,
    #[serde(rename = "gramEnd")]
    GramEnd,
}

impl std::fmt::Display for STProofErr {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::SpellStart => write!(f, "spellStart"),
            Self::SpellEnd => write!(f, "spellEnd"),
            Self::GramStart => write!(f, "gramStart"),
            Self::GramEnd => write!(f, "gramEnd"),
        }
    }
}

impl std::str::FromStr for STProofErr {
    type Err = String;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        match s {
            "spellStart" => Ok(Self::SpellStart),
            "spellEnd" => Ok(Self::SpellEnd),
            "gramStart" => Ok(Self::GramStart),
            "gramEnd" => Ok(Self::GramEnd),
            _ => Err(format!("unknown STProofErr value: {}", s)),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum STEdGrp {
    #[serde(rename = "none")]
    None,
    #[serde(rename = "everyone")]
    Everyone,
    #[serde(rename = "administrators")]
    Administrators,
    #[serde(rename = "contributors")]
    Contributors,
    #[serde(rename = "editors")]
    Editors,
    #[serde(rename = "owners")]
    Owners,
    #[serde(rename = "current")]
    Current,
}

impl std::fmt::Display for STEdGrp {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::None => write!(f, "none"),
            Self::Everyone => write!(f, "everyone"),
            Self::Administrators => write!(f, "administrators"),
            Self::Contributors => write!(f, "contributors"),
            Self::Editors => write!(f, "editors"),
            Self::Owners => write!(f, "owners"),
            Self::Current => write!(f, "current"),
        }
    }
}

impl std::str::FromStr for STEdGrp {
    type Err = String;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        match s {
            "none" => Ok(Self::None),
            "everyone" => Ok(Self::Everyone),
            "administrators" => Ok(Self::Administrators),
            "contributors" => Ok(Self::Contributors),
            "editors" => Ok(Self::Editors),
            "owners" => Ok(Self::Owners),
            "current" => Ok(Self::Current),
            _ => Err(format!("unknown STEdGrp value: {}", s)),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum STHint {
    #[serde(rename = "default")]
    Default,
    #[serde(rename = "eastAsia")]
    EastAsia,
}

impl std::fmt::Display for STHint {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::Default => write!(f, "default"),
            Self::EastAsia => write!(f, "eastAsia"),
        }
    }
}

impl std::str::FromStr for STHint {
    type Err = String;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        match s {
            "default" => Ok(Self::Default),
            "eastAsia" => Ok(Self::EastAsia),
            _ => Err(format!("unknown STHint value: {}", s)),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum STTheme {
    #[serde(rename = "majorEastAsia")]
    MajorEastAsia,
    #[serde(rename = "majorBidi")]
    MajorBidi,
    #[serde(rename = "majorAscii")]
    MajorAscii,
    #[serde(rename = "majorHAnsi")]
    MajorHAnsi,
    #[serde(rename = "minorEastAsia")]
    MinorEastAsia,
    #[serde(rename = "minorBidi")]
    MinorBidi,
    #[serde(rename = "minorAscii")]
    MinorAscii,
    #[serde(rename = "minorHAnsi")]
    MinorHAnsi,
}

impl std::fmt::Display for STTheme {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::MajorEastAsia => write!(f, "majorEastAsia"),
            Self::MajorBidi => write!(f, "majorBidi"),
            Self::MajorAscii => write!(f, "majorAscii"),
            Self::MajorHAnsi => write!(f, "majorHAnsi"),
            Self::MinorEastAsia => write!(f, "minorEastAsia"),
            Self::MinorBidi => write!(f, "minorBidi"),
            Self::MinorAscii => write!(f, "minorAscii"),
            Self::MinorHAnsi => write!(f, "minorHAnsi"),
        }
    }
}

impl std::str::FromStr for STTheme {
    type Err = String;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        match s {
            "majorEastAsia" => Ok(Self::MajorEastAsia),
            "majorBidi" => Ok(Self::MajorBidi),
            "majorAscii" => Ok(Self::MajorAscii),
            "majorHAnsi" => Ok(Self::MajorHAnsi),
            "minorEastAsia" => Ok(Self::MinorEastAsia),
            "minorBidi" => Ok(Self::MinorBidi),
            "minorAscii" => Ok(Self::MinorAscii),
            "minorHAnsi" => Ok(Self::MinorHAnsi),
            _ => Err(format!("unknown STTheme value: {}", s)),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum STRubyAlign {
    #[serde(rename = "center")]
    Center,
    #[serde(rename = "distributeLetter")]
    DistributeLetter,
    #[serde(rename = "distributeSpace")]
    DistributeSpace,
    #[serde(rename = "left")]
    Left,
    #[serde(rename = "right")]
    Right,
    #[serde(rename = "rightVertical")]
    RightVertical,
}

impl std::fmt::Display for STRubyAlign {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::Center => write!(f, "center"),
            Self::DistributeLetter => write!(f, "distributeLetter"),
            Self::DistributeSpace => write!(f, "distributeSpace"),
            Self::Left => write!(f, "left"),
            Self::Right => write!(f, "right"),
            Self::RightVertical => write!(f, "rightVertical"),
        }
    }
}

impl std::str::FromStr for STRubyAlign {
    type Err = String;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        match s {
            "center" => Ok(Self::Center),
            "distributeLetter" => Ok(Self::DistributeLetter),
            "distributeSpace" => Ok(Self::DistributeSpace),
            "left" => Ok(Self::Left),
            "right" => Ok(Self::Right),
            "rightVertical" => Ok(Self::RightVertical),
            _ => Err(format!("unknown STRubyAlign value: {}", s)),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum STLock {
    #[serde(rename = "sdtLocked")]
    SdtLocked,
    #[serde(rename = "contentLocked")]
    ContentLocked,
    #[serde(rename = "unlocked")]
    Unlocked,
    #[serde(rename = "sdtContentLocked")]
    SdtContentLocked,
}

impl std::fmt::Display for STLock {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::SdtLocked => write!(f, "sdtLocked"),
            Self::ContentLocked => write!(f, "contentLocked"),
            Self::Unlocked => write!(f, "unlocked"),
            Self::SdtContentLocked => write!(f, "sdtContentLocked"),
        }
    }
}

impl std::str::FromStr for STLock {
    type Err = String;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        match s {
            "sdtLocked" => Ok(Self::SdtLocked),
            "contentLocked" => Ok(Self::ContentLocked),
            "unlocked" => Ok(Self::Unlocked),
            "sdtContentLocked" => Ok(Self::SdtContentLocked),
            _ => Err(format!("unknown STLock value: {}", s)),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum STSdtDateMappingType {
    #[serde(rename = "text")]
    Text,
    #[serde(rename = "date")]
    Date,
    #[serde(rename = "dateTime")]
    DateTime,
}

impl std::fmt::Display for STSdtDateMappingType {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::Text => write!(f, "text"),
            Self::Date => write!(f, "date"),
            Self::DateTime => write!(f, "dateTime"),
        }
    }
}

impl std::str::FromStr for STSdtDateMappingType {
    type Err = String;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        match s {
            "text" => Ok(Self::Text),
            "date" => Ok(Self::Date),
            "dateTime" => Ok(Self::DateTime),
            _ => Err(format!("unknown STSdtDateMappingType value: {}", s)),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum STDirection {
    #[serde(rename = "ltr")]
    Ltr,
    #[serde(rename = "rtl")]
    Rtl,
}

impl std::fmt::Display for STDirection {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::Ltr => write!(f, "ltr"),
            Self::Rtl => write!(f, "rtl"),
        }
    }
}

impl std::str::FromStr for STDirection {
    type Err = String;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        match s {
            "ltr" => Ok(Self::Ltr),
            "rtl" => Ok(Self::Rtl),
            _ => Err(format!("unknown STDirection value: {}", s)),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum STTblWidth {
    #[serde(rename = "nil")]
    Nil,
    #[serde(rename = "pct")]
    Pct,
    #[serde(rename = "dxa")]
    Dxa,
    #[serde(rename = "auto")]
    Auto,
}

impl std::fmt::Display for STTblWidth {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::Nil => write!(f, "nil"),
            Self::Pct => write!(f, "pct"),
            Self::Dxa => write!(f, "dxa"),
            Self::Auto => write!(f, "auto"),
        }
    }
}

impl std::str::FromStr for STTblWidth {
    type Err = String;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        match s {
            "nil" => Ok(Self::Nil),
            "pct" => Ok(Self::Pct),
            "dxa" => Ok(Self::Dxa),
            "auto" => Ok(Self::Auto),
            _ => Err(format!("unknown STTblWidth value: {}", s)),
        }
    }
}

pub type STMeasurementOrPercent = String;

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum STMerge {
    #[serde(rename = "continue")]
    Continue,
    #[serde(rename = "restart")]
    Restart,
}

impl std::fmt::Display for STMerge {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::Continue => write!(f, "continue"),
            Self::Restart => write!(f, "restart"),
        }
    }
}

impl std::str::FromStr for STMerge {
    type Err = String;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        match s {
            "continue" => Ok(Self::Continue),
            "restart" => Ok(Self::Restart),
            _ => Err(format!("unknown STMerge value: {}", s)),
        }
    }
}

pub type STCnf = String;

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum STTblLayoutType {
    #[serde(rename = "fixed")]
    Fixed,
    #[serde(rename = "autofit")]
    Autofit,
}

impl std::fmt::Display for STTblLayoutType {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::Fixed => write!(f, "fixed"),
            Self::Autofit => write!(f, "autofit"),
        }
    }
}

impl std::str::FromStr for STTblLayoutType {
    type Err = String;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        match s {
            "fixed" => Ok(Self::Fixed),
            "autofit" => Ok(Self::Autofit),
            _ => Err(format!("unknown STTblLayoutType value: {}", s)),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum STTblOverlap {
    #[serde(rename = "never")]
    Never,
    #[serde(rename = "overlap")]
    Overlap,
}

impl std::fmt::Display for STTblOverlap {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::Never => write!(f, "never"),
            Self::Overlap => write!(f, "overlap"),
        }
    }
}

impl std::str::FromStr for STTblOverlap {
    type Err = String;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        match s {
            "never" => Ok(Self::Never),
            "overlap" => Ok(Self::Overlap),
            _ => Err(format!("unknown STTblOverlap value: {}", s)),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum STFtnPos {
    #[serde(rename = "pageBottom")]
    PageBottom,
    #[serde(rename = "beneathText")]
    BeneathText,
    #[serde(rename = "sectEnd")]
    SectEnd,
    #[serde(rename = "docEnd")]
    DocEnd,
}

impl std::fmt::Display for STFtnPos {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::PageBottom => write!(f, "pageBottom"),
            Self::BeneathText => write!(f, "beneathText"),
            Self::SectEnd => write!(f, "sectEnd"),
            Self::DocEnd => write!(f, "docEnd"),
        }
    }
}

impl std::str::FromStr for STFtnPos {
    type Err = String;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        match s {
            "pageBottom" => Ok(Self::PageBottom),
            "beneathText" => Ok(Self::BeneathText),
            "sectEnd" => Ok(Self::SectEnd),
            "docEnd" => Ok(Self::DocEnd),
            _ => Err(format!("unknown STFtnPos value: {}", s)),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum STEdnPos {
    #[serde(rename = "sectEnd")]
    SectEnd,
    #[serde(rename = "docEnd")]
    DocEnd,
}

impl std::fmt::Display for STEdnPos {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::SectEnd => write!(f, "sectEnd"),
            Self::DocEnd => write!(f, "docEnd"),
        }
    }
}

impl std::str::FromStr for STEdnPos {
    type Err = String;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        match s {
            "sectEnd" => Ok(Self::SectEnd),
            "docEnd" => Ok(Self::DocEnd),
            _ => Err(format!("unknown STEdnPos value: {}", s)),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum STRestartNumber {
    #[serde(rename = "continuous")]
    Continuous,
    #[serde(rename = "eachSect")]
    EachSect,
    #[serde(rename = "eachPage")]
    EachPage,
}

impl std::fmt::Display for STRestartNumber {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::Continuous => write!(f, "continuous"),
            Self::EachSect => write!(f, "eachSect"),
            Self::EachPage => write!(f, "eachPage"),
        }
    }
}

impl std::str::FromStr for STRestartNumber {
    type Err = String;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        match s {
            "continuous" => Ok(Self::Continuous),
            "eachSect" => Ok(Self::EachSect),
            "eachPage" => Ok(Self::EachPage),
            _ => Err(format!("unknown STRestartNumber value: {}", s)),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum STMailMergeSourceType {
    #[serde(rename = "database")]
    Database,
    #[serde(rename = "addressBook")]
    AddressBook,
    #[serde(rename = "document1")]
    Document1,
    #[serde(rename = "document2")]
    Document2,
    #[serde(rename = "text")]
    Text,
    #[serde(rename = "email")]
    Email,
    #[serde(rename = "native")]
    Native,
    #[serde(rename = "legacy")]
    Legacy,
    #[serde(rename = "master")]
    Master,
}

impl std::fmt::Display for STMailMergeSourceType {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::Database => write!(f, "database"),
            Self::AddressBook => write!(f, "addressBook"),
            Self::Document1 => write!(f, "document1"),
            Self::Document2 => write!(f, "document2"),
            Self::Text => write!(f, "text"),
            Self::Email => write!(f, "email"),
            Self::Native => write!(f, "native"),
            Self::Legacy => write!(f, "legacy"),
            Self::Master => write!(f, "master"),
        }
    }
}

impl std::str::FromStr for STMailMergeSourceType {
    type Err = String;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        match s {
            "database" => Ok(Self::Database),
            "addressBook" => Ok(Self::AddressBook),
            "document1" => Ok(Self::Document1),
            "document2" => Ok(Self::Document2),
            "text" => Ok(Self::Text),
            "email" => Ok(Self::Email),
            "native" => Ok(Self::Native),
            "legacy" => Ok(Self::Legacy),
            "master" => Ok(Self::Master),
            _ => Err(format!("unknown STMailMergeSourceType value: {}", s)),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum STTargetScreenSz {
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

impl std::fmt::Display for STTargetScreenSz {
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

impl std::str::FromStr for STTargetScreenSz {
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
            _ => Err(format!("unknown STTargetScreenSz value: {}", s)),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum STCharacterSpacing {
    #[serde(rename = "doNotCompress")]
    DoNotCompress,
    #[serde(rename = "compressPunctuation")]
    CompressPunctuation,
    #[serde(rename = "compressPunctuationAndJapaneseKana")]
    CompressPunctuationAndJapaneseKana,
}

impl std::fmt::Display for STCharacterSpacing {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::DoNotCompress => write!(f, "doNotCompress"),
            Self::CompressPunctuation => write!(f, "compressPunctuation"),
            Self::CompressPunctuationAndJapaneseKana => {
                write!(f, "compressPunctuationAndJapaneseKana")
            }
        }
    }
}

impl std::str::FromStr for STCharacterSpacing {
    type Err = String;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        match s {
            "doNotCompress" => Ok(Self::DoNotCompress),
            "compressPunctuation" => Ok(Self::CompressPunctuation),
            "compressPunctuationAndJapaneseKana" => Ok(Self::CompressPunctuationAndJapaneseKana),
            _ => Err(format!("unknown STCharacterSpacing value: {}", s)),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum STWmlColorSchemeIndex {
    #[serde(rename = "dark1")]
    Dark1,
    #[serde(rename = "light1")]
    Light1,
    #[serde(rename = "dark2")]
    Dark2,
    #[serde(rename = "light2")]
    Light2,
    #[serde(rename = "accent1")]
    Accent1,
    #[serde(rename = "accent2")]
    Accent2,
    #[serde(rename = "accent3")]
    Accent3,
    #[serde(rename = "accent4")]
    Accent4,
    #[serde(rename = "accent5")]
    Accent5,
    #[serde(rename = "accent6")]
    Accent6,
    #[serde(rename = "hyperlink")]
    Hyperlink,
    #[serde(rename = "followedHyperlink")]
    FollowedHyperlink,
}

impl std::fmt::Display for STWmlColorSchemeIndex {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::Dark1 => write!(f, "dark1"),
            Self::Light1 => write!(f, "light1"),
            Self::Dark2 => write!(f, "dark2"),
            Self::Light2 => write!(f, "light2"),
            Self::Accent1 => write!(f, "accent1"),
            Self::Accent2 => write!(f, "accent2"),
            Self::Accent3 => write!(f, "accent3"),
            Self::Accent4 => write!(f, "accent4"),
            Self::Accent5 => write!(f, "accent5"),
            Self::Accent6 => write!(f, "accent6"),
            Self::Hyperlink => write!(f, "hyperlink"),
            Self::FollowedHyperlink => write!(f, "followedHyperlink"),
        }
    }
}

impl std::str::FromStr for STWmlColorSchemeIndex {
    type Err = String;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        match s {
            "dark1" => Ok(Self::Dark1),
            "light1" => Ok(Self::Light1),
            "dark2" => Ok(Self::Dark2),
            "light2" => Ok(Self::Light2),
            "accent1" => Ok(Self::Accent1),
            "accent2" => Ok(Self::Accent2),
            "accent3" => Ok(Self::Accent3),
            "accent4" => Ok(Self::Accent4),
            "accent5" => Ok(Self::Accent5),
            "accent6" => Ok(Self::Accent6),
            "hyperlink" => Ok(Self::Hyperlink),
            "followedHyperlink" => Ok(Self::FollowedHyperlink),
            _ => Err(format!("unknown STWmlColorSchemeIndex value: {}", s)),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum STStyleSort {
    #[serde(rename = "name")]
    Name,
    #[serde(rename = "priority")]
    Priority,
    #[serde(rename = "default")]
    Default,
    #[serde(rename = "font")]
    Font,
    #[serde(rename = "basedOn")]
    BasedOn,
    #[serde(rename = "type")]
    Type,
    #[serde(rename = "0000")]
    _0000,
    #[serde(rename = "0001")]
    _0001,
    #[serde(rename = "0002")]
    _0002,
    #[serde(rename = "0003")]
    _0003,
    #[serde(rename = "0004")]
    _0004,
    #[serde(rename = "0005")]
    _0005,
}

impl std::fmt::Display for STStyleSort {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::Name => write!(f, "name"),
            Self::Priority => write!(f, "priority"),
            Self::Default => write!(f, "default"),
            Self::Font => write!(f, "font"),
            Self::BasedOn => write!(f, "basedOn"),
            Self::Type => write!(f, "type"),
            Self::_0000 => write!(f, "0000"),
            Self::_0001 => write!(f, "0001"),
            Self::_0002 => write!(f, "0002"),
            Self::_0003 => write!(f, "0003"),
            Self::_0004 => write!(f, "0004"),
            Self::_0005 => write!(f, "0005"),
        }
    }
}

impl std::str::FromStr for STStyleSort {
    type Err = String;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        match s {
            "name" => Ok(Self::Name),
            "priority" => Ok(Self::Priority),
            "default" => Ok(Self::Default),
            "font" => Ok(Self::Font),
            "basedOn" => Ok(Self::BasedOn),
            "type" => Ok(Self::Type),
            "0000" => Ok(Self::_0000),
            "0001" => Ok(Self::_0001),
            "0002" => Ok(Self::_0002),
            "0003" => Ok(Self::_0003),
            "0004" => Ok(Self::_0004),
            "0005" => Ok(Self::_0005),
            _ => Err(format!("unknown STStyleSort value: {}", s)),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum STFrameScrollbar {
    #[serde(rename = "on")]
    On,
    #[serde(rename = "off")]
    Off,
    #[serde(rename = "auto")]
    Auto,
}

impl std::fmt::Display for STFrameScrollbar {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::On => write!(f, "on"),
            Self::Off => write!(f, "off"),
            Self::Auto => write!(f, "auto"),
        }
    }
}

impl std::str::FromStr for STFrameScrollbar {
    type Err = String;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        match s {
            "on" => Ok(Self::On),
            "off" => Ok(Self::Off),
            "auto" => Ok(Self::Auto),
            _ => Err(format!("unknown STFrameScrollbar value: {}", s)),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum STFrameLayout {
    #[serde(rename = "rows")]
    Rows,
    #[serde(rename = "cols")]
    Cols,
    #[serde(rename = "none")]
    None,
}

impl std::fmt::Display for STFrameLayout {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::Rows => write!(f, "rows"),
            Self::Cols => write!(f, "cols"),
            Self::None => write!(f, "none"),
        }
    }
}

impl std::str::FromStr for STFrameLayout {
    type Err = String;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        match s {
            "rows" => Ok(Self::Rows),
            "cols" => Ok(Self::Cols),
            "none" => Ok(Self::None),
            _ => Err(format!("unknown STFrameLayout value: {}", s)),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum STLevelSuffix {
    #[serde(rename = "tab")]
    Tab,
    #[serde(rename = "space")]
    Space,
    #[serde(rename = "nothing")]
    Nothing,
}

impl std::fmt::Display for STLevelSuffix {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::Tab => write!(f, "tab"),
            Self::Space => write!(f, "space"),
            Self::Nothing => write!(f, "nothing"),
        }
    }
}

impl std::str::FromStr for STLevelSuffix {
    type Err = String;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        match s {
            "tab" => Ok(Self::Tab),
            "space" => Ok(Self::Space),
            "nothing" => Ok(Self::Nothing),
            _ => Err(format!("unknown STLevelSuffix value: {}", s)),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum STMultiLevelType {
    #[serde(rename = "singleLevel")]
    SingleLevel,
    #[serde(rename = "multilevel")]
    Multilevel,
    #[serde(rename = "hybridMultilevel")]
    HybridMultilevel,
}

impl std::fmt::Display for STMultiLevelType {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::SingleLevel => write!(f, "singleLevel"),
            Self::Multilevel => write!(f, "multilevel"),
            Self::HybridMultilevel => write!(f, "hybridMultilevel"),
        }
    }
}

impl std::str::FromStr for STMultiLevelType {
    type Err = String;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        match s {
            "singleLevel" => Ok(Self::SingleLevel),
            "multilevel" => Ok(Self::Multilevel),
            "hybridMultilevel" => Ok(Self::HybridMultilevel),
            _ => Err(format!("unknown STMultiLevelType value: {}", s)),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum STTblStyleOverrideType {
    #[serde(rename = "wholeTable")]
    WholeTable,
    #[serde(rename = "firstRow")]
    FirstRow,
    #[serde(rename = "lastRow")]
    LastRow,
    #[serde(rename = "firstCol")]
    FirstCol,
    #[serde(rename = "lastCol")]
    LastCol,
    #[serde(rename = "band1Vert")]
    Band1Vert,
    #[serde(rename = "band2Vert")]
    Band2Vert,
    #[serde(rename = "band1Horz")]
    Band1Horz,
    #[serde(rename = "band2Horz")]
    Band2Horz,
    #[serde(rename = "neCell")]
    NeCell,
    #[serde(rename = "nwCell")]
    NwCell,
    #[serde(rename = "seCell")]
    SeCell,
    #[serde(rename = "swCell")]
    SwCell,
}

impl std::fmt::Display for STTblStyleOverrideType {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::WholeTable => write!(f, "wholeTable"),
            Self::FirstRow => write!(f, "firstRow"),
            Self::LastRow => write!(f, "lastRow"),
            Self::FirstCol => write!(f, "firstCol"),
            Self::LastCol => write!(f, "lastCol"),
            Self::Band1Vert => write!(f, "band1Vert"),
            Self::Band2Vert => write!(f, "band2Vert"),
            Self::Band1Horz => write!(f, "band1Horz"),
            Self::Band2Horz => write!(f, "band2Horz"),
            Self::NeCell => write!(f, "neCell"),
            Self::NwCell => write!(f, "nwCell"),
            Self::SeCell => write!(f, "seCell"),
            Self::SwCell => write!(f, "swCell"),
        }
    }
}

impl std::str::FromStr for STTblStyleOverrideType {
    type Err = String;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        match s {
            "wholeTable" => Ok(Self::WholeTable),
            "firstRow" => Ok(Self::FirstRow),
            "lastRow" => Ok(Self::LastRow),
            "firstCol" => Ok(Self::FirstCol),
            "lastCol" => Ok(Self::LastCol),
            "band1Vert" => Ok(Self::Band1Vert),
            "band2Vert" => Ok(Self::Band2Vert),
            "band1Horz" => Ok(Self::Band1Horz),
            "band2Horz" => Ok(Self::Band2Horz),
            "neCell" => Ok(Self::NeCell),
            "nwCell" => Ok(Self::NwCell),
            "seCell" => Ok(Self::SeCell),
            "swCell" => Ok(Self::SwCell),
            _ => Err(format!("unknown STTblStyleOverrideType value: {}", s)),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum STStyleType {
    #[serde(rename = "paragraph")]
    Paragraph,
    #[serde(rename = "character")]
    Character,
    #[serde(rename = "table")]
    Table,
    #[serde(rename = "numbering")]
    Numbering,
}

impl std::fmt::Display for STStyleType {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::Paragraph => write!(f, "paragraph"),
            Self::Character => write!(f, "character"),
            Self::Table => write!(f, "table"),
            Self::Numbering => write!(f, "numbering"),
        }
    }
}

impl std::str::FromStr for STStyleType {
    type Err = String;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        match s {
            "paragraph" => Ok(Self::Paragraph),
            "character" => Ok(Self::Character),
            "table" => Ok(Self::Table),
            "numbering" => Ok(Self::Numbering),
            _ => Err(format!("unknown STStyleType value: {}", s)),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum STFontFamily {
    #[serde(rename = "decorative")]
    Decorative,
    #[serde(rename = "modern")]
    Modern,
    #[serde(rename = "roman")]
    Roman,
    #[serde(rename = "script")]
    Script,
    #[serde(rename = "swiss")]
    Swiss,
    #[serde(rename = "auto")]
    Auto,
}

impl std::fmt::Display for STFontFamily {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::Decorative => write!(f, "decorative"),
            Self::Modern => write!(f, "modern"),
            Self::Roman => write!(f, "roman"),
            Self::Script => write!(f, "script"),
            Self::Swiss => write!(f, "swiss"),
            Self::Auto => write!(f, "auto"),
        }
    }
}

impl std::str::FromStr for STFontFamily {
    type Err = String;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        match s {
            "decorative" => Ok(Self::Decorative),
            "modern" => Ok(Self::Modern),
            "roman" => Ok(Self::Roman),
            "script" => Ok(Self::Script),
            "swiss" => Ok(Self::Swiss),
            "auto" => Ok(Self::Auto),
            _ => Err(format!("unknown STFontFamily value: {}", s)),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum STPitch {
    #[serde(rename = "fixed")]
    Fixed,
    #[serde(rename = "variable")]
    Variable,
    #[serde(rename = "default")]
    Default,
}

impl std::fmt::Display for STPitch {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::Fixed => write!(f, "fixed"),
            Self::Variable => write!(f, "variable"),
            Self::Default => write!(f, "default"),
        }
    }
}

impl std::str::FromStr for STPitch {
    type Err = String;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        match s {
            "fixed" => Ok(Self::Fixed),
            "variable" => Ok(Self::Variable),
            "default" => Ok(Self::Default),
            _ => Err(format!("unknown STPitch value: {}", s)),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum STThemeColor {
    #[serde(rename = "dark1")]
    Dark1,
    #[serde(rename = "light1")]
    Light1,
    #[serde(rename = "dark2")]
    Dark2,
    #[serde(rename = "light2")]
    Light2,
    #[serde(rename = "accent1")]
    Accent1,
    #[serde(rename = "accent2")]
    Accent2,
    #[serde(rename = "accent3")]
    Accent3,
    #[serde(rename = "accent4")]
    Accent4,
    #[serde(rename = "accent5")]
    Accent5,
    #[serde(rename = "accent6")]
    Accent6,
    #[serde(rename = "hyperlink")]
    Hyperlink,
    #[serde(rename = "followedHyperlink")]
    FollowedHyperlink,
    #[serde(rename = "none")]
    None,
    #[serde(rename = "background1")]
    Background1,
    #[serde(rename = "text1")]
    Text1,
    #[serde(rename = "background2")]
    Background2,
    #[serde(rename = "text2")]
    Text2,
}

impl std::fmt::Display for STThemeColor {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::Dark1 => write!(f, "dark1"),
            Self::Light1 => write!(f, "light1"),
            Self::Dark2 => write!(f, "dark2"),
            Self::Light2 => write!(f, "light2"),
            Self::Accent1 => write!(f, "accent1"),
            Self::Accent2 => write!(f, "accent2"),
            Self::Accent3 => write!(f, "accent3"),
            Self::Accent4 => write!(f, "accent4"),
            Self::Accent5 => write!(f, "accent5"),
            Self::Accent6 => write!(f, "accent6"),
            Self::Hyperlink => write!(f, "hyperlink"),
            Self::FollowedHyperlink => write!(f, "followedHyperlink"),
            Self::None => write!(f, "none"),
            Self::Background1 => write!(f, "background1"),
            Self::Text1 => write!(f, "text1"),
            Self::Background2 => write!(f, "background2"),
            Self::Text2 => write!(f, "text2"),
        }
    }
}

impl std::str::FromStr for STThemeColor {
    type Err = String;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        match s {
            "dark1" => Ok(Self::Dark1),
            "light1" => Ok(Self::Light1),
            "dark2" => Ok(Self::Dark2),
            "light2" => Ok(Self::Light2),
            "accent1" => Ok(Self::Accent1),
            "accent2" => Ok(Self::Accent2),
            "accent3" => Ok(Self::Accent3),
            "accent4" => Ok(Self::Accent4),
            "accent5" => Ok(Self::Accent5),
            "accent6" => Ok(Self::Accent6),
            "hyperlink" => Ok(Self::Hyperlink),
            "followedHyperlink" => Ok(Self::FollowedHyperlink),
            "none" => Ok(Self::None),
            "background1" => Ok(Self::Background1),
            "text1" => Ok(Self::Text1),
            "background2" => Ok(Self::Background2),
            "text2" => Ok(Self::Text2),
            _ => Err(format!("unknown STThemeColor value: {}", s)),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum STDocPartBehavior {
    #[serde(rename = "content")]
    Content,
    #[serde(rename = "p")]
    P,
    #[serde(rename = "pg")]
    Pg,
}

impl std::fmt::Display for STDocPartBehavior {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::Content => write!(f, "content"),
            Self::P => write!(f, "p"),
            Self::Pg => write!(f, "pg"),
        }
    }
}

impl std::str::FromStr for STDocPartBehavior {
    type Err = String;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        match s {
            "content" => Ok(Self::Content),
            "p" => Ok(Self::P),
            "pg" => Ok(Self::Pg),
            _ => Err(format!("unknown STDocPartBehavior value: {}", s)),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum STDocPartType {
    #[serde(rename = "none")]
    None,
    #[serde(rename = "normal")]
    Normal,
    #[serde(rename = "autoExp")]
    AutoExp,
    #[serde(rename = "toolbar")]
    Toolbar,
    #[serde(rename = "speller")]
    Speller,
    #[serde(rename = "formFld")]
    FormFld,
    #[serde(rename = "bbPlcHdr")]
    BbPlcHdr,
}

impl std::fmt::Display for STDocPartType {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::None => write!(f, "none"),
            Self::Normal => write!(f, "normal"),
            Self::AutoExp => write!(f, "autoExp"),
            Self::Toolbar => write!(f, "toolbar"),
            Self::Speller => write!(f, "speller"),
            Self::FormFld => write!(f, "formFld"),
            Self::BbPlcHdr => write!(f, "bbPlcHdr"),
        }
    }
}

impl std::str::FromStr for STDocPartType {
    type Err = String;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        match s {
            "none" => Ok(Self::None),
            "normal" => Ok(Self::Normal),
            "autoExp" => Ok(Self::AutoExp),
            "toolbar" => Ok(Self::Toolbar),
            "speller" => Ok(Self::Speller),
            "formFld" => Ok(Self::FormFld),
            "bbPlcHdr" => Ok(Self::BbPlcHdr),
            _ => Err(format!("unknown STDocPartType value: {}", s)),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum STDocPartGallery {
    #[serde(rename = "placeholder")]
    Placeholder,
    #[serde(rename = "any")]
    Any,
    #[serde(rename = "default")]
    Default,
    #[serde(rename = "docParts")]
    DocParts,
    #[serde(rename = "coverPg")]
    CoverPg,
    #[serde(rename = "eq")]
    Eq,
    #[serde(rename = "ftrs")]
    Ftrs,
    #[serde(rename = "hdrs")]
    Hdrs,
    #[serde(rename = "pgNum")]
    PgNum,
    #[serde(rename = "tbls")]
    Tbls,
    #[serde(rename = "watermarks")]
    Watermarks,
    #[serde(rename = "autoTxt")]
    AutoTxt,
    #[serde(rename = "txtBox")]
    TxtBox,
    #[serde(rename = "pgNumT")]
    PgNumT,
    #[serde(rename = "pgNumB")]
    PgNumB,
    #[serde(rename = "pgNumMargins")]
    PgNumMargins,
    #[serde(rename = "tblOfContents")]
    TblOfContents,
    #[serde(rename = "bib")]
    Bib,
    #[serde(rename = "custQuickParts")]
    CustQuickParts,
    #[serde(rename = "custCoverPg")]
    CustCoverPg,
    #[serde(rename = "custEq")]
    CustEq,
    #[serde(rename = "custFtrs")]
    CustFtrs,
    #[serde(rename = "custHdrs")]
    CustHdrs,
    #[serde(rename = "custPgNum")]
    CustPgNum,
    #[serde(rename = "custTbls")]
    CustTbls,
    #[serde(rename = "custWatermarks")]
    CustWatermarks,
    #[serde(rename = "custAutoTxt")]
    CustAutoTxt,
    #[serde(rename = "custTxtBox")]
    CustTxtBox,
    #[serde(rename = "custPgNumT")]
    CustPgNumT,
    #[serde(rename = "custPgNumB")]
    CustPgNumB,
    #[serde(rename = "custPgNumMargins")]
    CustPgNumMargins,
    #[serde(rename = "custTblOfContents")]
    CustTblOfContents,
    #[serde(rename = "custBib")]
    CustBib,
    #[serde(rename = "custom1")]
    Custom1,
    #[serde(rename = "custom2")]
    Custom2,
    #[serde(rename = "custom3")]
    Custom3,
    #[serde(rename = "custom4")]
    Custom4,
    #[serde(rename = "custom5")]
    Custom5,
}

impl std::fmt::Display for STDocPartGallery {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::Placeholder => write!(f, "placeholder"),
            Self::Any => write!(f, "any"),
            Self::Default => write!(f, "default"),
            Self::DocParts => write!(f, "docParts"),
            Self::CoverPg => write!(f, "coverPg"),
            Self::Eq => write!(f, "eq"),
            Self::Ftrs => write!(f, "ftrs"),
            Self::Hdrs => write!(f, "hdrs"),
            Self::PgNum => write!(f, "pgNum"),
            Self::Tbls => write!(f, "tbls"),
            Self::Watermarks => write!(f, "watermarks"),
            Self::AutoTxt => write!(f, "autoTxt"),
            Self::TxtBox => write!(f, "txtBox"),
            Self::PgNumT => write!(f, "pgNumT"),
            Self::PgNumB => write!(f, "pgNumB"),
            Self::PgNumMargins => write!(f, "pgNumMargins"),
            Self::TblOfContents => write!(f, "tblOfContents"),
            Self::Bib => write!(f, "bib"),
            Self::CustQuickParts => write!(f, "custQuickParts"),
            Self::CustCoverPg => write!(f, "custCoverPg"),
            Self::CustEq => write!(f, "custEq"),
            Self::CustFtrs => write!(f, "custFtrs"),
            Self::CustHdrs => write!(f, "custHdrs"),
            Self::CustPgNum => write!(f, "custPgNum"),
            Self::CustTbls => write!(f, "custTbls"),
            Self::CustWatermarks => write!(f, "custWatermarks"),
            Self::CustAutoTxt => write!(f, "custAutoTxt"),
            Self::CustTxtBox => write!(f, "custTxtBox"),
            Self::CustPgNumT => write!(f, "custPgNumT"),
            Self::CustPgNumB => write!(f, "custPgNumB"),
            Self::CustPgNumMargins => write!(f, "custPgNumMargins"),
            Self::CustTblOfContents => write!(f, "custTblOfContents"),
            Self::CustBib => write!(f, "custBib"),
            Self::Custom1 => write!(f, "custom1"),
            Self::Custom2 => write!(f, "custom2"),
            Self::Custom3 => write!(f, "custom3"),
            Self::Custom4 => write!(f, "custom4"),
            Self::Custom5 => write!(f, "custom5"),
        }
    }
}

impl std::str::FromStr for STDocPartGallery {
    type Err = String;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        match s {
            "placeholder" => Ok(Self::Placeholder),
            "any" => Ok(Self::Any),
            "default" => Ok(Self::Default),
            "docParts" => Ok(Self::DocParts),
            "coverPg" => Ok(Self::CoverPg),
            "eq" => Ok(Self::Eq),
            "ftrs" => Ok(Self::Ftrs),
            "hdrs" => Ok(Self::Hdrs),
            "pgNum" => Ok(Self::PgNum),
            "tbls" => Ok(Self::Tbls),
            "watermarks" => Ok(Self::Watermarks),
            "autoTxt" => Ok(Self::AutoTxt),
            "txtBox" => Ok(Self::TxtBox),
            "pgNumT" => Ok(Self::PgNumT),
            "pgNumB" => Ok(Self::PgNumB),
            "pgNumMargins" => Ok(Self::PgNumMargins),
            "tblOfContents" => Ok(Self::TblOfContents),
            "bib" => Ok(Self::Bib),
            "custQuickParts" => Ok(Self::CustQuickParts),
            "custCoverPg" => Ok(Self::CustCoverPg),
            "custEq" => Ok(Self::CustEq),
            "custFtrs" => Ok(Self::CustFtrs),
            "custHdrs" => Ok(Self::CustHdrs),
            "custPgNum" => Ok(Self::CustPgNum),
            "custTbls" => Ok(Self::CustTbls),
            "custWatermarks" => Ok(Self::CustWatermarks),
            "custAutoTxt" => Ok(Self::CustAutoTxt),
            "custTxtBox" => Ok(Self::CustTxtBox),
            "custPgNumT" => Ok(Self::CustPgNumT),
            "custPgNumB" => Ok(Self::CustPgNumB),
            "custPgNumMargins" => Ok(Self::CustPgNumMargins),
            "custTblOfContents" => Ok(Self::CustTblOfContents),
            "custBib" => Ok(Self::CustBib),
            "custom1" => Ok(Self::Custom1),
            "custom2" => Ok(Self::Custom2),
            "custom3" => Ok(Self::Custom3),
            "custom4" => Ok(Self::Custom4),
            "custom5" => Ok(Self::Custom5),
            _ => Err(format!("unknown STDocPartGallery value: {}", s)),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum STCaptionPos {
    #[serde(rename = "above")]
    Above,
    #[serde(rename = "below")]
    Below,
    #[serde(rename = "left")]
    Left,
    #[serde(rename = "right")]
    Right,
}

impl std::fmt::Display for STCaptionPos {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::Above => write!(f, "above"),
            Self::Below => write!(f, "below"),
            Self::Left => write!(f, "left"),
            Self::Right => write!(f, "right"),
        }
    }
}

impl std::str::FromStr for STCaptionPos {
    type Err = String;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        match s {
            "above" => Ok(Self::Above),
            "below" => Ok(Self::Below),
            "left" => Ok(Self::Left),
            "right" => Ok(Self::Right),
            _ => Err(format!("unknown STCaptionPos value: {}", s)),
        }
    }
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub enum ParagraphContentBase {
    #[serde(rename = "customXml")]
    CustomXml(Box<CTCustomXmlRun>),
    #[serde(rename = "fldSimple")]
    FldSimple(Box<CTSimpleField>),
    #[serde(rename = "hyperlink")]
    Hyperlink(Box<Hyperlink>),
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub enum RunContentBase {
    #[serde(rename = "smartTag")]
    SmartTag(Box<CTSmartTagRun>),
    #[serde(rename = "sdt")]
    Sdt(Box<CTSdtRun>),
    #[serde(rename = "proofErr")]
    ProofErr(Box<CTProofErr>),
    #[serde(rename = "permStart")]
    PermStart(Box<CTPermStart>),
    #[serde(rename = "permEnd")]
    PermEnd(Box<CTPerm>),
    #[serde(rename = "bookmarkStart")]
    BookmarkStart(Box<Bookmark>),
    #[serde(rename = "bookmarkEnd")]
    BookmarkEnd(Box<CTMarkupRange>),
    #[serde(rename = "moveFromRangeStart")]
    MoveFromRangeStart(Box<CTMoveBookmark>),
    #[serde(rename = "moveFromRangeEnd")]
    MoveFromRangeEnd(Box<CTMarkupRange>),
    #[serde(rename = "moveToRangeStart")]
    MoveToRangeStart(Box<CTMoveBookmark>),
    #[serde(rename = "moveToRangeEnd")]
    MoveToRangeEnd(Box<CTMarkupRange>),
    #[serde(rename = "commentRangeStart")]
    CommentRangeStart(Box<CTMarkupRange>),
    #[serde(rename = "commentRangeEnd")]
    CommentRangeEnd(Box<CTMarkupRange>),
    #[serde(rename = "customXmlInsRangeStart")]
    CustomXmlInsRangeStart(Box<CTTrackChange>),
    #[serde(rename = "customXmlInsRangeEnd")]
    CustomXmlInsRangeEnd(Box<CTMarkup>),
    #[serde(rename = "customXmlDelRangeStart")]
    CustomXmlDelRangeStart(Box<CTTrackChange>),
    #[serde(rename = "customXmlDelRangeEnd")]
    CustomXmlDelRangeEnd(Box<CTMarkup>),
    #[serde(rename = "customXmlMoveFromRangeStart")]
    CustomXmlMoveFromRangeStart(Box<CTTrackChange>),
    #[serde(rename = "customXmlMoveFromRangeEnd")]
    CustomXmlMoveFromRangeEnd(Box<CTMarkup>),
    #[serde(rename = "customXmlMoveToRangeStart")]
    CustomXmlMoveToRangeStart(Box<CTTrackChange>),
    #[serde(rename = "customXmlMoveToRangeEnd")]
    CustomXmlMoveToRangeEnd(Box<CTMarkup>),
    #[serde(rename = "ins")]
    Ins(Box<CTRunTrackChange>),
    #[serde(rename = "del")]
    Del(Box<CTRunTrackChange>),
    #[serde(rename = "moveFrom")]
    MoveFrom(Box<CTRunTrackChange>),
    #[serde(rename = "moveTo")]
    MoveTo(Box<CTRunTrackChange>),
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub enum CellMarkup {
    #[serde(rename = "cellIns")]
    CellIns(Box<CTTrackChange>),
    #[serde(rename = "cellDel")]
    CellDel(Box<CTTrackChange>),
    #[serde(rename = "cellMerge")]
    CellMerge(Box<CTCellMergeTrackChange>),
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub enum RangeMarkup {
    #[serde(rename = "bookmarkStart")]
    BookmarkStart(Box<Bookmark>),
    #[serde(rename = "bookmarkEnd")]
    BookmarkEnd(Box<CTMarkupRange>),
    #[serde(rename = "moveFromRangeStart")]
    MoveFromRangeStart(Box<CTMoveBookmark>),
    #[serde(rename = "moveFromRangeEnd")]
    MoveFromRangeEnd(Box<CTMarkupRange>),
    #[serde(rename = "moveToRangeStart")]
    MoveToRangeStart(Box<CTMoveBookmark>),
    #[serde(rename = "moveToRangeEnd")]
    MoveToRangeEnd(Box<CTMarkupRange>),
    #[serde(rename = "commentRangeStart")]
    CommentRangeStart(Box<CTMarkupRange>),
    #[serde(rename = "commentRangeEnd")]
    CommentRangeEnd(Box<CTMarkupRange>),
    #[serde(rename = "customXmlInsRangeStart")]
    CustomXmlInsRangeStart(Box<CTTrackChange>),
    #[serde(rename = "customXmlInsRangeEnd")]
    CustomXmlInsRangeEnd(Box<CTMarkup>),
    #[serde(rename = "customXmlDelRangeStart")]
    CustomXmlDelRangeStart(Box<CTTrackChange>),
    #[serde(rename = "customXmlDelRangeEnd")]
    CustomXmlDelRangeEnd(Box<CTMarkup>),
    #[serde(rename = "customXmlMoveFromRangeStart")]
    CustomXmlMoveFromRangeStart(Box<CTTrackChange>),
    #[serde(rename = "customXmlMoveFromRangeEnd")]
    CustomXmlMoveFromRangeEnd(Box<CTMarkup>),
    #[serde(rename = "customXmlMoveToRangeStart")]
    CustomXmlMoveToRangeStart(Box<CTTrackChange>),
    #[serde(rename = "customXmlMoveToRangeEnd")]
    CustomXmlMoveToRangeEnd(Box<CTMarkup>),
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub enum HeaderFooterRef {
    #[serde(rename = "headerReference")]
    HeaderReference(Box<HeaderFooterReference>),
    #[serde(rename = "footerReference")]
    FooterReference(Box<HeaderFooterReference>),
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub enum RunContent {
    #[serde(rename = "br")]
    Br(Box<CTBr>),
    #[serde(rename = "t")]
    T(Box<Text>),
    #[serde(rename = "contentPart")]
    ContentPart(Box<CTRel>),
    #[serde(rename = "delText")]
    DelText(Box<Text>),
    #[serde(rename = "instrText")]
    InstrText(Box<Text>),
    #[serde(rename = "delInstrText")]
    DelInstrText(Box<Text>),
    #[serde(rename = "noBreakHyphen")]
    NoBreakHyphen(Box<CTEmpty>),
    #[serde(rename = "softHyphen")]
    SoftHyphen(Box<CTEmpty>),
    #[serde(rename = "dayShort")]
    DayShort(Box<CTEmpty>),
    #[serde(rename = "monthShort")]
    MonthShort(Box<CTEmpty>),
    #[serde(rename = "yearShort")]
    YearShort(Box<CTEmpty>),
    #[serde(rename = "dayLong")]
    DayLong(Box<CTEmpty>),
    #[serde(rename = "monthLong")]
    MonthLong(Box<CTEmpty>),
    #[serde(rename = "yearLong")]
    YearLong(Box<CTEmpty>),
    #[serde(rename = "annotationRef")]
    AnnotationRef(Box<CTEmpty>),
    #[serde(rename = "footnoteRef")]
    FootnoteRef(Box<CTEmpty>),
    #[serde(rename = "endnoteRef")]
    EndnoteRef(Box<CTEmpty>),
    #[serde(rename = "separator")]
    Separator(Box<CTEmpty>),
    #[serde(rename = "continuationSeparator")]
    ContinuationSeparator(Box<CTEmpty>),
    #[serde(rename = "sym")]
    Sym(Box<CTSym>),
    #[serde(rename = "pgNum")]
    PgNum(Box<CTEmpty>),
    #[serde(rename = "cr")]
    Cr(Box<CTEmpty>),
    #[serde(rename = "tab")]
    Tab(Box<CTEmpty>),
    #[serde(rename = "object")]
    Object(Box<CTObject>),
    #[serde(rename = "pict")]
    Pict(Box<CTPicture>),
    #[serde(rename = "fldChar")]
    FldChar(Box<CTFldChar>),
    #[serde(rename = "ruby")]
    Ruby(Box<CTRuby>),
    #[serde(rename = "footnoteReference")]
    FootnoteReference(Box<FootnoteEndnoteRef>),
    #[serde(rename = "endnoteReference")]
    EndnoteReference(Box<FootnoteEndnoteRef>),
    #[serde(rename = "commentReference")]
    CommentReference(Box<CTMarkup>),
    #[serde(rename = "drawing")]
    Drawing(Box<CTDrawing>),
    #[serde(rename = "ptab")]
    Ptab(Box<CTPTab>),
    #[serde(rename = "lastRenderedPageBreak")]
    LastRenderedPageBreak(Box<CTEmpty>),
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub enum MathRunProperties {
    #[serde(rename = "rPr")]
    RPr(Box<RunProperties>),
    #[serde(rename = "ins")]
    Ins(Box<CTMathCtrlIns>),
    #[serde(rename = "del")]
    Del(Box<CTMathCtrlDel>),
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub enum RubyContent {
    #[serde(rename = "r")]
    R(Box<Run>),
    #[serde(rename = "proofErr")]
    ProofErr(Box<CTProofErr>),
    #[serde(rename = "permStart")]
    PermStart(Box<CTPermStart>),
    #[serde(rename = "permEnd")]
    PermEnd(Box<CTPerm>),
    #[serde(rename = "bookmarkStart")]
    BookmarkStart(Box<Bookmark>),
    #[serde(rename = "bookmarkEnd")]
    BookmarkEnd(Box<CTMarkupRange>),
    #[serde(rename = "moveFromRangeStart")]
    MoveFromRangeStart(Box<CTMoveBookmark>),
    #[serde(rename = "moveFromRangeEnd")]
    MoveFromRangeEnd(Box<CTMarkupRange>),
    #[serde(rename = "moveToRangeStart")]
    MoveToRangeStart(Box<CTMoveBookmark>),
    #[serde(rename = "moveToRangeEnd")]
    MoveToRangeEnd(Box<CTMarkupRange>),
    #[serde(rename = "commentRangeStart")]
    CommentRangeStart(Box<CTMarkupRange>),
    #[serde(rename = "commentRangeEnd")]
    CommentRangeEnd(Box<CTMarkupRange>),
    #[serde(rename = "customXmlInsRangeStart")]
    CustomXmlInsRangeStart(Box<CTTrackChange>),
    #[serde(rename = "customXmlInsRangeEnd")]
    CustomXmlInsRangeEnd(Box<CTMarkup>),
    #[serde(rename = "customXmlDelRangeStart")]
    CustomXmlDelRangeStart(Box<CTTrackChange>),
    #[serde(rename = "customXmlDelRangeEnd")]
    CustomXmlDelRangeEnd(Box<CTMarkup>),
    #[serde(rename = "customXmlMoveFromRangeStart")]
    CustomXmlMoveFromRangeStart(Box<CTTrackChange>),
    #[serde(rename = "customXmlMoveFromRangeEnd")]
    CustomXmlMoveFromRangeEnd(Box<CTMarkup>),
    #[serde(rename = "customXmlMoveToRangeStart")]
    CustomXmlMoveToRangeStart(Box<CTTrackChange>),
    #[serde(rename = "customXmlMoveToRangeEnd")]
    CustomXmlMoveToRangeEnd(Box<CTMarkup>),
    #[serde(rename = "ins")]
    Ins(Box<CTRunTrackChange>),
    #[serde(rename = "del")]
    Del(Box<CTRunTrackChange>),
    #[serde(rename = "moveFrom")]
    MoveFrom(Box<CTRunTrackChange>),
    #[serde(rename = "moveTo")]
    MoveTo(Box<CTRunTrackChange>),
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub enum RunContentChoice {
    #[serde(rename = "customXml")]
    CustomXml(Box<CTCustomXmlRun>),
    #[serde(rename = "smartTag")]
    SmartTag(Box<CTSmartTagRun>),
    #[serde(rename = "sdt")]
    Sdt(Box<CTSdtRun>),
    #[serde(rename = "dir")]
    Dir(Box<CTDirContentRun>),
    #[serde(rename = "bdo")]
    Bdo(Box<CTBdoContentRun>),
    #[serde(rename = "r")]
    R(Box<Run>),
    #[serde(rename = "proofErr")]
    ProofErr(Box<CTProofErr>),
    #[serde(rename = "permStart")]
    PermStart(Box<CTPermStart>),
    #[serde(rename = "permEnd")]
    PermEnd(Box<CTPerm>),
    #[serde(rename = "bookmarkStart")]
    BookmarkStart(Box<Bookmark>),
    #[serde(rename = "bookmarkEnd")]
    BookmarkEnd(Box<CTMarkupRange>),
    #[serde(rename = "moveFromRangeStart")]
    MoveFromRangeStart(Box<CTMoveBookmark>),
    #[serde(rename = "moveFromRangeEnd")]
    MoveFromRangeEnd(Box<CTMarkupRange>),
    #[serde(rename = "moveToRangeStart")]
    MoveToRangeStart(Box<CTMoveBookmark>),
    #[serde(rename = "moveToRangeEnd")]
    MoveToRangeEnd(Box<CTMarkupRange>),
    #[serde(rename = "commentRangeStart")]
    CommentRangeStart(Box<CTMarkupRange>),
    #[serde(rename = "commentRangeEnd")]
    CommentRangeEnd(Box<CTMarkupRange>),
    #[serde(rename = "customXmlInsRangeStart")]
    CustomXmlInsRangeStart(Box<CTTrackChange>),
    #[serde(rename = "customXmlInsRangeEnd")]
    CustomXmlInsRangeEnd(Box<CTMarkup>),
    #[serde(rename = "customXmlDelRangeStart")]
    CustomXmlDelRangeStart(Box<CTTrackChange>),
    #[serde(rename = "customXmlDelRangeEnd")]
    CustomXmlDelRangeEnd(Box<CTMarkup>),
    #[serde(rename = "customXmlMoveFromRangeStart")]
    CustomXmlMoveFromRangeStart(Box<CTTrackChange>),
    #[serde(rename = "customXmlMoveFromRangeEnd")]
    CustomXmlMoveFromRangeEnd(Box<CTMarkup>),
    #[serde(rename = "customXmlMoveToRangeStart")]
    CustomXmlMoveToRangeStart(Box<CTTrackChange>),
    #[serde(rename = "customXmlMoveToRangeEnd")]
    CustomXmlMoveToRangeEnd(Box<CTMarkup>),
    #[serde(rename = "ins")]
    Ins(Box<CTRunTrackChange>),
    #[serde(rename = "del")]
    Del(Box<CTRunTrackChange>),
    #[serde(rename = "moveFrom")]
    MoveFrom(Box<CTRunTrackChange>),
    #[serde(rename = "moveTo")]
    MoveTo(Box<CTRunTrackChange>),
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub enum BlockContentChoice {
    #[serde(rename = "customXml")]
    CustomXml(Box<CTCustomXmlBlock>),
    #[serde(rename = "sdt")]
    Sdt(Box<CTSdtBlock>),
    #[serde(rename = "p")]
    P(Box<Paragraph>),
    #[serde(rename = "tbl")]
    Tbl(Box<Table>),
    #[serde(rename = "proofErr")]
    ProofErr(Box<CTProofErr>),
    #[serde(rename = "permStart")]
    PermStart(Box<CTPermStart>),
    #[serde(rename = "permEnd")]
    PermEnd(Box<CTPerm>),
    #[serde(rename = "bookmarkStart")]
    BookmarkStart(Box<Bookmark>),
    #[serde(rename = "bookmarkEnd")]
    BookmarkEnd(Box<CTMarkupRange>),
    #[serde(rename = "moveFromRangeStart")]
    MoveFromRangeStart(Box<CTMoveBookmark>),
    #[serde(rename = "moveFromRangeEnd")]
    MoveFromRangeEnd(Box<CTMarkupRange>),
    #[serde(rename = "moveToRangeStart")]
    MoveToRangeStart(Box<CTMoveBookmark>),
    #[serde(rename = "moveToRangeEnd")]
    MoveToRangeEnd(Box<CTMarkupRange>),
    #[serde(rename = "commentRangeStart")]
    CommentRangeStart(Box<CTMarkupRange>),
    #[serde(rename = "commentRangeEnd")]
    CommentRangeEnd(Box<CTMarkupRange>),
    #[serde(rename = "customXmlInsRangeStart")]
    CustomXmlInsRangeStart(Box<CTTrackChange>),
    #[serde(rename = "customXmlInsRangeEnd")]
    CustomXmlInsRangeEnd(Box<CTMarkup>),
    #[serde(rename = "customXmlDelRangeStart")]
    CustomXmlDelRangeStart(Box<CTTrackChange>),
    #[serde(rename = "customXmlDelRangeEnd")]
    CustomXmlDelRangeEnd(Box<CTMarkup>),
    #[serde(rename = "customXmlMoveFromRangeStart")]
    CustomXmlMoveFromRangeStart(Box<CTTrackChange>),
    #[serde(rename = "customXmlMoveFromRangeEnd")]
    CustomXmlMoveFromRangeEnd(Box<CTMarkup>),
    #[serde(rename = "customXmlMoveToRangeStart")]
    CustomXmlMoveToRangeStart(Box<CTTrackChange>),
    #[serde(rename = "customXmlMoveToRangeEnd")]
    CustomXmlMoveToRangeEnd(Box<CTMarkup>),
    #[serde(rename = "ins")]
    Ins(Box<CTRunTrackChange>),
    #[serde(rename = "del")]
    Del(Box<CTRunTrackChange>),
    #[serde(rename = "moveFrom")]
    MoveFrom(Box<CTRunTrackChange>),
    #[serde(rename = "moveTo")]
    MoveTo(Box<CTRunTrackChange>),
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub enum RowContent {
    #[serde(rename = "tr")]
    Tr(Box<CTRow>),
    #[serde(rename = "customXml")]
    CustomXml(Box<CTCustomXmlRow>),
    #[serde(rename = "sdt")]
    Sdt(Box<CTSdtRow>),
    #[serde(rename = "proofErr")]
    ProofErr(Box<CTProofErr>),
    #[serde(rename = "permStart")]
    PermStart(Box<CTPermStart>),
    #[serde(rename = "permEnd")]
    PermEnd(Box<CTPerm>),
    #[serde(rename = "bookmarkStart")]
    BookmarkStart(Box<Bookmark>),
    #[serde(rename = "bookmarkEnd")]
    BookmarkEnd(Box<CTMarkupRange>),
    #[serde(rename = "moveFromRangeStart")]
    MoveFromRangeStart(Box<CTMoveBookmark>),
    #[serde(rename = "moveFromRangeEnd")]
    MoveFromRangeEnd(Box<CTMarkupRange>),
    #[serde(rename = "moveToRangeStart")]
    MoveToRangeStart(Box<CTMoveBookmark>),
    #[serde(rename = "moveToRangeEnd")]
    MoveToRangeEnd(Box<CTMarkupRange>),
    #[serde(rename = "commentRangeStart")]
    CommentRangeStart(Box<CTMarkupRange>),
    #[serde(rename = "commentRangeEnd")]
    CommentRangeEnd(Box<CTMarkupRange>),
    #[serde(rename = "customXmlInsRangeStart")]
    CustomXmlInsRangeStart(Box<CTTrackChange>),
    #[serde(rename = "customXmlInsRangeEnd")]
    CustomXmlInsRangeEnd(Box<CTMarkup>),
    #[serde(rename = "customXmlDelRangeStart")]
    CustomXmlDelRangeStart(Box<CTTrackChange>),
    #[serde(rename = "customXmlDelRangeEnd")]
    CustomXmlDelRangeEnd(Box<CTMarkup>),
    #[serde(rename = "customXmlMoveFromRangeStart")]
    CustomXmlMoveFromRangeStart(Box<CTTrackChange>),
    #[serde(rename = "customXmlMoveFromRangeEnd")]
    CustomXmlMoveFromRangeEnd(Box<CTMarkup>),
    #[serde(rename = "customXmlMoveToRangeStart")]
    CustomXmlMoveToRangeStart(Box<CTTrackChange>),
    #[serde(rename = "customXmlMoveToRangeEnd")]
    CustomXmlMoveToRangeEnd(Box<CTMarkup>),
    #[serde(rename = "ins")]
    Ins(Box<CTRunTrackChange>),
    #[serde(rename = "del")]
    Del(Box<CTRunTrackChange>),
    #[serde(rename = "moveFrom")]
    MoveFrom(Box<CTRunTrackChange>),
    #[serde(rename = "moveTo")]
    MoveTo(Box<CTRunTrackChange>),
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub enum CellContent {
    #[serde(rename = "tc")]
    Tc(Box<TableCell>),
    #[serde(rename = "customXml")]
    CustomXml(Box<CTCustomXmlCell>),
    #[serde(rename = "sdt")]
    Sdt(Box<CTSdtCell>),
    #[serde(rename = "proofErr")]
    ProofErr(Box<CTProofErr>),
    #[serde(rename = "permStart")]
    PermStart(Box<CTPermStart>),
    #[serde(rename = "permEnd")]
    PermEnd(Box<CTPerm>),
    #[serde(rename = "bookmarkStart")]
    BookmarkStart(Box<Bookmark>),
    #[serde(rename = "bookmarkEnd")]
    BookmarkEnd(Box<CTMarkupRange>),
    #[serde(rename = "moveFromRangeStart")]
    MoveFromRangeStart(Box<CTMoveBookmark>),
    #[serde(rename = "moveFromRangeEnd")]
    MoveFromRangeEnd(Box<CTMarkupRange>),
    #[serde(rename = "moveToRangeStart")]
    MoveToRangeStart(Box<CTMoveBookmark>),
    #[serde(rename = "moveToRangeEnd")]
    MoveToRangeEnd(Box<CTMarkupRange>),
    #[serde(rename = "commentRangeStart")]
    CommentRangeStart(Box<CTMarkupRange>),
    #[serde(rename = "commentRangeEnd")]
    CommentRangeEnd(Box<CTMarkupRange>),
    #[serde(rename = "customXmlInsRangeStart")]
    CustomXmlInsRangeStart(Box<CTTrackChange>),
    #[serde(rename = "customXmlInsRangeEnd")]
    CustomXmlInsRangeEnd(Box<CTMarkup>),
    #[serde(rename = "customXmlDelRangeStart")]
    CustomXmlDelRangeStart(Box<CTTrackChange>),
    #[serde(rename = "customXmlDelRangeEnd")]
    CustomXmlDelRangeEnd(Box<CTMarkup>),
    #[serde(rename = "customXmlMoveFromRangeStart")]
    CustomXmlMoveFromRangeStart(Box<CTTrackChange>),
    #[serde(rename = "customXmlMoveFromRangeEnd")]
    CustomXmlMoveFromRangeEnd(Box<CTMarkup>),
    #[serde(rename = "customXmlMoveToRangeStart")]
    CustomXmlMoveToRangeStart(Box<CTTrackChange>),
    #[serde(rename = "customXmlMoveToRangeEnd")]
    CustomXmlMoveToRangeEnd(Box<CTMarkup>),
    #[serde(rename = "ins")]
    Ins(Box<CTRunTrackChange>),
    #[serde(rename = "del")]
    Del(Box<CTRunTrackChange>),
    #[serde(rename = "moveFrom")]
    MoveFrom(Box<CTRunTrackChange>),
    #[serde(rename = "moveTo")]
    MoveTo(Box<CTRunTrackChange>),
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub enum ParagraphContent {
    #[serde(rename = "customXml")]
    CustomXml(Box<CTCustomXmlRun>),
    #[serde(rename = "smartTag")]
    SmartTag(Box<CTSmartTagRun>),
    #[serde(rename = "sdt")]
    Sdt(Box<CTSdtRun>),
    #[serde(rename = "dir")]
    Dir(Box<CTDirContentRun>),
    #[serde(rename = "bdo")]
    Bdo(Box<CTBdoContentRun>),
    #[serde(rename = "r")]
    R(Box<Run>),
    #[serde(rename = "proofErr")]
    ProofErr(Box<CTProofErr>),
    #[serde(rename = "permStart")]
    PermStart(Box<CTPermStart>),
    #[serde(rename = "permEnd")]
    PermEnd(Box<CTPerm>),
    #[serde(rename = "bookmarkStart")]
    BookmarkStart(Box<Bookmark>),
    #[serde(rename = "bookmarkEnd")]
    BookmarkEnd(Box<CTMarkupRange>),
    #[serde(rename = "moveFromRangeStart")]
    MoveFromRangeStart(Box<CTMoveBookmark>),
    #[serde(rename = "moveFromRangeEnd")]
    MoveFromRangeEnd(Box<CTMarkupRange>),
    #[serde(rename = "moveToRangeStart")]
    MoveToRangeStart(Box<CTMoveBookmark>),
    #[serde(rename = "moveToRangeEnd")]
    MoveToRangeEnd(Box<CTMarkupRange>),
    #[serde(rename = "commentRangeStart")]
    CommentRangeStart(Box<CTMarkupRange>),
    #[serde(rename = "commentRangeEnd")]
    CommentRangeEnd(Box<CTMarkupRange>),
    #[serde(rename = "customXmlInsRangeStart")]
    CustomXmlInsRangeStart(Box<CTTrackChange>),
    #[serde(rename = "customXmlInsRangeEnd")]
    CustomXmlInsRangeEnd(Box<CTMarkup>),
    #[serde(rename = "customXmlDelRangeStart")]
    CustomXmlDelRangeStart(Box<CTTrackChange>),
    #[serde(rename = "customXmlDelRangeEnd")]
    CustomXmlDelRangeEnd(Box<CTMarkup>),
    #[serde(rename = "customXmlMoveFromRangeStart")]
    CustomXmlMoveFromRangeStart(Box<CTTrackChange>),
    #[serde(rename = "customXmlMoveFromRangeEnd")]
    CustomXmlMoveFromRangeEnd(Box<CTMarkup>),
    #[serde(rename = "customXmlMoveToRangeStart")]
    CustomXmlMoveToRangeStart(Box<CTTrackChange>),
    #[serde(rename = "customXmlMoveToRangeEnd")]
    CustomXmlMoveToRangeEnd(Box<CTMarkup>),
    #[serde(rename = "ins")]
    Ins(Box<CTRunTrackChange>),
    #[serde(rename = "del")]
    Del(Box<CTRunTrackChange>),
    #[serde(rename = "moveFrom")]
    MoveFrom(Box<CTRunTrackChange>),
    #[serde(rename = "moveTo")]
    MoveTo(Box<CTRunTrackChange>),
    #[serde(rename = "fldSimple")]
    FldSimple(Box<CTSimpleField>),
    #[serde(rename = "hyperlink")]
    Hyperlink(Box<Hyperlink>),
    #[serde(rename = "subDoc")]
    SubDoc(Box<CTRel>),
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub enum BlockContent {
    #[serde(rename = "customXml")]
    CustomXml(Box<CTCustomXmlBlock>),
    #[serde(rename = "sdt")]
    Sdt(Box<CTSdtBlock>),
    #[serde(rename = "p")]
    P(Box<Paragraph>),
    #[serde(rename = "tbl")]
    Tbl(Box<Table>),
    #[serde(rename = "proofErr")]
    ProofErr(Box<CTProofErr>),
    #[serde(rename = "permStart")]
    PermStart(Box<CTPermStart>),
    #[serde(rename = "permEnd")]
    PermEnd(Box<CTPerm>),
    #[serde(rename = "bookmarkStart")]
    BookmarkStart(Box<Bookmark>),
    #[serde(rename = "bookmarkEnd")]
    BookmarkEnd(Box<CTMarkupRange>),
    #[serde(rename = "moveFromRangeStart")]
    MoveFromRangeStart(Box<CTMoveBookmark>),
    #[serde(rename = "moveFromRangeEnd")]
    MoveFromRangeEnd(Box<CTMarkupRange>),
    #[serde(rename = "moveToRangeStart")]
    MoveToRangeStart(Box<CTMoveBookmark>),
    #[serde(rename = "moveToRangeEnd")]
    MoveToRangeEnd(Box<CTMarkupRange>),
    #[serde(rename = "commentRangeStart")]
    CommentRangeStart(Box<CTMarkupRange>),
    #[serde(rename = "commentRangeEnd")]
    CommentRangeEnd(Box<CTMarkupRange>),
    #[serde(rename = "customXmlInsRangeStart")]
    CustomXmlInsRangeStart(Box<CTTrackChange>),
    #[serde(rename = "customXmlInsRangeEnd")]
    CustomXmlInsRangeEnd(Box<CTMarkup>),
    #[serde(rename = "customXmlDelRangeStart")]
    CustomXmlDelRangeStart(Box<CTTrackChange>),
    #[serde(rename = "customXmlDelRangeEnd")]
    CustomXmlDelRangeEnd(Box<CTMarkup>),
    #[serde(rename = "customXmlMoveFromRangeStart")]
    CustomXmlMoveFromRangeStart(Box<CTTrackChange>),
    #[serde(rename = "customXmlMoveFromRangeEnd")]
    CustomXmlMoveFromRangeEnd(Box<CTMarkup>),
    #[serde(rename = "customXmlMoveToRangeStart")]
    CustomXmlMoveToRangeStart(Box<CTTrackChange>),
    #[serde(rename = "customXmlMoveToRangeEnd")]
    CustomXmlMoveToRangeEnd(Box<CTMarkup>),
    #[serde(rename = "ins")]
    Ins(Box<CTRunTrackChange>),
    #[serde(rename = "del")]
    Del(Box<CTRunTrackChange>),
    #[serde(rename = "moveFrom")]
    MoveFrom(Box<CTRunTrackChange>),
    #[serde(rename = "moveTo")]
    MoveTo(Box<CTRunTrackChange>),
    #[serde(rename = "altChunk")]
    AltChunk(Box<CTAltChunk>),
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub enum RunLevelContent {
    #[serde(rename = "proofErr")]
    ProofErr(Box<CTProofErr>),
    #[serde(rename = "permStart")]
    PermStart(Box<CTPermStart>),
    #[serde(rename = "permEnd")]
    PermEnd(Box<CTPerm>),
    #[serde(rename = "bookmarkStart")]
    BookmarkStart(Box<Bookmark>),
    #[serde(rename = "bookmarkEnd")]
    BookmarkEnd(Box<CTMarkupRange>),
    #[serde(rename = "moveFromRangeStart")]
    MoveFromRangeStart(Box<CTMoveBookmark>),
    #[serde(rename = "moveFromRangeEnd")]
    MoveFromRangeEnd(Box<CTMarkupRange>),
    #[serde(rename = "moveToRangeStart")]
    MoveToRangeStart(Box<CTMoveBookmark>),
    #[serde(rename = "moveToRangeEnd")]
    MoveToRangeEnd(Box<CTMarkupRange>),
    #[serde(rename = "commentRangeStart")]
    CommentRangeStart(Box<CTMarkupRange>),
    #[serde(rename = "commentRangeEnd")]
    CommentRangeEnd(Box<CTMarkupRange>),
    #[serde(rename = "customXmlInsRangeStart")]
    CustomXmlInsRangeStart(Box<CTTrackChange>),
    #[serde(rename = "customXmlInsRangeEnd")]
    CustomXmlInsRangeEnd(Box<CTMarkup>),
    #[serde(rename = "customXmlDelRangeStart")]
    CustomXmlDelRangeStart(Box<CTTrackChange>),
    #[serde(rename = "customXmlDelRangeEnd")]
    CustomXmlDelRangeEnd(Box<CTMarkup>),
    #[serde(rename = "customXmlMoveFromRangeStart")]
    CustomXmlMoveFromRangeStart(Box<CTTrackChange>),
    #[serde(rename = "customXmlMoveFromRangeEnd")]
    CustomXmlMoveFromRangeEnd(Box<CTMarkup>),
    #[serde(rename = "customXmlMoveToRangeStart")]
    CustomXmlMoveToRangeStart(Box<CTTrackChange>),
    #[serde(rename = "customXmlMoveToRangeEnd")]
    CustomXmlMoveToRangeEnd(Box<CTMarkup>),
    #[serde(rename = "ins")]
    Ins(Box<CTRunTrackChange>),
    #[serde(rename = "del")]
    Del(Box<CTRunTrackChange>),
    #[serde(rename = "moveFrom")]
    MoveFrom(Box<CTRunTrackChange>),
    #[serde(rename = "moveTo")]
    MoveTo(Box<CTRunTrackChange>),
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct CTEmpty;

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct CTOnOff {
    #[serde(rename = "@w:val")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub value: Option<OnOff>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTLongHexNumber {
    #[serde(rename = "@w:val")]
    pub value: STLongHexNumber,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct CTCharset {
    #[serde(rename = "@w:val")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub value: Option<STUcharHexNumber>,
    #[cfg(feature = "wml-styling")]
    #[serde(rename = "@w:characterSet")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub character_set: Option<STString>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTDecimalNumber {
    #[serde(rename = "@w:val")]
    pub value: STDecimalNumber,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTUnsignedDecimalNumber {
    #[serde(rename = "@w:val")]
    pub value: STUnsignedDecimalNumber,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTDecimalNumberOrPrecent {
    #[serde(rename = "@w:val")]
    pub value: STDecimalNumberOrPercent,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTTwipsMeasure {
    #[serde(rename = "@w:val")]
    pub value: STTwipsMeasure,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTSignedTwipsMeasure {
    #[serde(rename = "@w:val")]
    pub value: STSignedTwipsMeasure,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTPixelsMeasure {
    #[serde(rename = "@w:val")]
    pub value: STPixelsMeasure,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTHpsMeasure {
    #[serde(rename = "@w:val")]
    pub value: STHpsMeasure,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTSignedHpsMeasure {
    #[serde(rename = "@w:val")]
    pub value: STSignedHpsMeasure,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTMacroName {
    #[serde(rename = "@w:val")]
    pub value: STMacroName,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTString {
    #[serde(rename = "@w:val")]
    pub value: STString,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct CTTextScale {
    #[serde(rename = "@w:val")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub value: Option<STTextScale>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTHighlight {
    #[serde(rename = "@w:val")]
    pub value: STHighlightColor,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTColor {
    #[serde(rename = "@w:val")]
    pub value: STHexColor,
    #[cfg(feature = "wml-styling")]
    #[serde(rename = "@w:themeColor")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub theme_color: Option<STThemeColor>,
    #[cfg(feature = "wml-styling")]
    #[serde(rename = "@w:themeTint")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub theme_tint: Option<STUcharHexNumber>,
    #[cfg(feature = "wml-styling")]
    #[serde(rename = "@w:themeShade")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub theme_shade: Option<STUcharHexNumber>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTLang {
    #[serde(rename = "@w:val")]
    pub value: Language,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct CTGuid {
    #[serde(rename = "@w:val")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub value: Option<Guid>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct CTUnderline {
    #[serde(rename = "@w:val")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub value: Option<STUnderline>,
    #[cfg(feature = "wml-styling")]
    #[serde(rename = "@w:color")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub color: Option<STHexColor>,
    #[cfg(feature = "wml-styling")]
    #[serde(rename = "@w:themeColor")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub theme_color: Option<STThemeColor>,
    #[cfg(feature = "wml-styling")]
    #[serde(rename = "@w:themeTint")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub theme_tint: Option<STUcharHexNumber>,
    #[cfg(feature = "wml-styling")]
    #[serde(rename = "@w:themeShade")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub theme_shade: Option<STUcharHexNumber>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTTextEffect {
    #[serde(rename = "@w:val")]
    pub value: STTextEffect,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTBorder {
    #[serde(rename = "@w:val")]
    pub value: STBorder,
    #[cfg(feature = "wml-styling")]
    #[serde(rename = "@w:color")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub color: Option<STHexColor>,
    #[cfg(feature = "wml-styling")]
    #[serde(rename = "@w:themeColor")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub theme_color: Option<STThemeColor>,
    #[cfg(feature = "wml-styling")]
    #[serde(rename = "@w:themeTint")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub theme_tint: Option<STUcharHexNumber>,
    #[cfg(feature = "wml-styling")]
    #[serde(rename = "@w:themeShade")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub theme_shade: Option<STUcharHexNumber>,
    #[cfg(feature = "wml-styling")]
    #[serde(rename = "@w:sz")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub size: Option<STEighthPointMeasure>,
    #[cfg(feature = "wml-styling")]
    #[serde(rename = "@w:space")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub space: Option<STPointMeasure>,
    #[cfg(feature = "wml-styling")]
    #[serde(rename = "@w:shadow")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub shadow: Option<OnOff>,
    #[cfg(feature = "wml-styling")]
    #[serde(rename = "@w:frame")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub frame: Option<OnOff>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTShd {
    #[serde(rename = "@w:val")]
    pub value: STShd,
    #[cfg(feature = "wml-styling")]
    #[serde(rename = "@w:color")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub color: Option<STHexColor>,
    #[cfg(feature = "wml-styling")]
    #[serde(rename = "@w:themeColor")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub theme_color: Option<STThemeColor>,
    #[cfg(feature = "wml-styling")]
    #[serde(rename = "@w:themeTint")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub theme_tint: Option<STUcharHexNumber>,
    #[cfg(feature = "wml-styling")]
    #[serde(rename = "@w:themeShade")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub theme_shade: Option<STUcharHexNumber>,
    #[cfg(feature = "wml-styling")]
    #[serde(rename = "@w:fill")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub fill: Option<STHexColor>,
    #[cfg(feature = "wml-styling")]
    #[serde(rename = "@w:themeFill")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub theme_fill: Option<STThemeColor>,
    #[cfg(feature = "wml-styling")]
    #[serde(rename = "@w:themeFillTint")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub theme_fill_tint: Option<STUcharHexNumber>,
    #[cfg(feature = "wml-styling")]
    #[serde(rename = "@w:themeFillShade")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub theme_fill_shade: Option<STUcharHexNumber>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTVerticalAlignRun {
    #[serde(rename = "@w:val")]
    pub value: STVerticalAlignRun,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTFitText {
    #[cfg(feature = "wml-styling")]
    #[serde(rename = "@w:val")]
    pub value: STTwipsMeasure,
    #[cfg(feature = "wml-styling")]
    #[serde(rename = "@w:id")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub id: Option<STDecimalNumber>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTEm {
    #[serde(rename = "@w:val")]
    pub value: STEm,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct CTLanguage {
    #[cfg(feature = "wml-styling")]
    #[serde(rename = "@w:val")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub value: Option<Language>,
    #[cfg(feature = "wml-styling")]
    #[serde(rename = "@w:eastAsia")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub east_asia: Option<Language>,
    #[cfg(feature = "wml-styling")]
    #[serde(rename = "@w:bidi")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub bidi: Option<Language>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct CTEastAsianLayout {
    #[cfg(feature = "wml-styling")]
    #[serde(rename = "@w:id")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub id: Option<STDecimalNumber>,
    #[cfg(feature = "wml-styling")]
    #[serde(rename = "@w:combine")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub combine: Option<OnOff>,
    #[cfg(feature = "wml-styling")]
    #[serde(rename = "@w:combineBrackets")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub combine_brackets: Option<STCombineBrackets>,
    #[cfg(feature = "wml-styling")]
    #[serde(rename = "@w:vert")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub vert: Option<OnOff>,
    #[cfg(feature = "wml-styling")]
    #[serde(rename = "@w:vertCompress")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub vert_compress: Option<OnOff>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct CTFramePr {
    #[cfg(feature = "wml-layout")]
    #[serde(rename = "@w:dropCap")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub drop_cap: Option<STDropCap>,
    #[cfg(feature = "wml-layout")]
    #[serde(rename = "@w:lines")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub lines: Option<STDecimalNumber>,
    #[cfg(feature = "wml-layout")]
    #[serde(rename = "@w:w")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub width: Option<STTwipsMeasure>,
    #[cfg(feature = "wml-layout")]
    #[serde(rename = "@w:h")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub height: Option<STTwipsMeasure>,
    #[cfg(feature = "wml-layout")]
    #[serde(rename = "@w:vSpace")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub v_space: Option<STTwipsMeasure>,
    #[cfg(feature = "wml-layout")]
    #[serde(rename = "@w:hSpace")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub h_space: Option<STTwipsMeasure>,
    #[cfg(feature = "wml-layout")]
    #[serde(rename = "@w:wrap")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub wrap: Option<STWrap>,
    #[cfg(feature = "wml-layout")]
    #[serde(rename = "@w:hAnchor")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub h_anchor: Option<STHAnchor>,
    #[cfg(feature = "wml-layout")]
    #[serde(rename = "@w:vAnchor")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub v_anchor: Option<STVAnchor>,
    #[cfg(feature = "wml-layout")]
    #[serde(rename = "@w:x")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub x: Option<STSignedTwipsMeasure>,
    #[cfg(feature = "wml-layout")]
    #[serde(rename = "@w:xAlign")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub x_align: Option<STXAlign>,
    #[cfg(feature = "wml-layout")]
    #[serde(rename = "@w:y")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub y: Option<STSignedTwipsMeasure>,
    #[cfg(feature = "wml-layout")]
    #[serde(rename = "@w:yAlign")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub y_align: Option<STYAlign>,
    #[cfg(feature = "wml-layout")]
    #[serde(rename = "@w:hRule")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub h_rule: Option<STHeightRule>,
    #[cfg(feature = "wml-layout")]
    #[serde(rename = "@w:anchorLock")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub anchor_lock: Option<OnOff>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTTabStop {
    #[serde(rename = "@w:val")]
    pub value: STTabJc,
    #[cfg(feature = "wml-styling")]
    #[serde(rename = "@w:leader")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub leader: Option<STTabTlc>,
    #[serde(rename = "@w:pos")]
    pub pos: STSignedTwipsMeasure,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct CTSpacing {
    #[cfg(feature = "wml-styling")]
    #[serde(rename = "@w:before")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub before: Option<STTwipsMeasure>,
    #[cfg(feature = "wml-styling")]
    #[serde(rename = "@w:beforeLines")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub before_lines: Option<STDecimalNumber>,
    #[cfg(feature = "wml-styling")]
    #[serde(rename = "@w:beforeAutospacing")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub before_autospacing: Option<OnOff>,
    #[cfg(feature = "wml-styling")]
    #[serde(rename = "@w:after")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub after: Option<STTwipsMeasure>,
    #[cfg(feature = "wml-styling")]
    #[serde(rename = "@w:afterLines")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub after_lines: Option<STDecimalNumber>,
    #[cfg(feature = "wml-styling")]
    #[serde(rename = "@w:afterAutospacing")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub after_autospacing: Option<OnOff>,
    #[cfg(feature = "wml-styling")]
    #[serde(rename = "@w:line")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub line: Option<STSignedTwipsMeasure>,
    #[cfg(feature = "wml-styling")]
    #[serde(rename = "@w:lineRule")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub line_rule: Option<STLineSpacingRule>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct CTInd {
    #[cfg(feature = "wml-styling")]
    #[serde(rename = "@w:start")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub start: Option<STSignedTwipsMeasure>,
    #[cfg(feature = "wml-styling")]
    #[serde(rename = "@w:startChars")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub start_chars: Option<STDecimalNumber>,
    #[cfg(feature = "wml-styling")]
    #[serde(rename = "@w:end")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub end: Option<STSignedTwipsMeasure>,
    #[cfg(feature = "wml-styling")]
    #[serde(rename = "@w:endChars")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub end_chars: Option<STDecimalNumber>,
    #[cfg(feature = "wml-styling")]
    #[serde(rename = "@w:left")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub left: Option<STSignedTwipsMeasure>,
    #[cfg(feature = "wml-styling")]
    #[serde(rename = "@w:leftChars")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub left_chars: Option<STDecimalNumber>,
    #[cfg(feature = "wml-styling")]
    #[serde(rename = "@w:right")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub right: Option<STSignedTwipsMeasure>,
    #[cfg(feature = "wml-styling")]
    #[serde(rename = "@w:rightChars")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub right_chars: Option<STDecimalNumber>,
    #[cfg(feature = "wml-styling")]
    #[serde(rename = "@w:hanging")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub hanging: Option<STTwipsMeasure>,
    #[cfg(feature = "wml-styling")]
    #[serde(rename = "@w:hangingChars")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub hanging_chars: Option<STDecimalNumber>,
    #[cfg(feature = "wml-styling")]
    #[serde(rename = "@w:firstLine")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub first_line: Option<STTwipsMeasure>,
    #[cfg(feature = "wml-styling")]
    #[serde(rename = "@w:firstLineChars")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub first_line_chars: Option<STDecimalNumber>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTJc {
    #[serde(rename = "@w:val")]
    pub value: STJc,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTJcTable {
    #[serde(rename = "@w:val")]
    pub value: STJcTable,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTView {
    #[serde(rename = "@w:val")]
    pub value: STView,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTZoom {
    #[cfg(feature = "wml-settings")]
    #[serde(rename = "@w:val")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub value: Option<STZoom>,
    #[serde(rename = "@w:percent")]
    pub percent: STDecimalNumberOrPercent,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTWritingStyle {
    #[cfg(feature = "wml-settings")]
    #[serde(rename = "@w:lang")]
    pub lang: Language,
    #[cfg(feature = "wml-settings")]
    #[serde(rename = "@w:vendorID")]
    pub vendor_i_d: STString,
    #[cfg(feature = "wml-settings")]
    #[serde(rename = "@w:dllVersion")]
    pub dll_version: STString,
    #[cfg(feature = "wml-settings")]
    #[serde(rename = "@w:nlCheck")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub nl_check: Option<OnOff>,
    #[cfg(feature = "wml-settings")]
    #[serde(rename = "@w:checkStyle")]
    pub check_style: OnOff,
    #[cfg(feature = "wml-settings")]
    #[serde(rename = "@w:appName")]
    pub app_name: STString,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct CTProof {
    #[cfg(feature = "wml-settings")]
    #[serde(rename = "@w:spelling")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub spelling: Option<STProof>,
    #[cfg(feature = "wml-settings")]
    #[serde(rename = "@w:grammar")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub grammar: Option<STProof>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTDocType {
    #[serde(rename = "@w:val")]
    pub value: STDocType,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct WAGPassword {
    #[serde(rename = "@w:algorithmName")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub algorithm_name: Option<STString>,
    #[serde(rename = "@w:hashValue")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub hash_value: Option<Vec<u8>>,
    #[serde(rename = "@w:saltValue")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub salt_value: Option<Vec<u8>>,
    #[serde(rename = "@w:spinCount")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub spin_count: Option<STDecimalNumber>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct WAGTransitionalPassword {
    #[serde(rename = "@w:cryptProviderType")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub crypt_provider_type: Option<STCryptProv>,
    #[serde(rename = "@w:cryptAlgorithmClass")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub crypt_algorithm_class: Option<STAlgClass>,
    #[serde(rename = "@w:cryptAlgorithmType")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub crypt_algorithm_type: Option<STAlgType>,
    #[serde(rename = "@w:cryptAlgorithmSid")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub crypt_algorithm_sid: Option<STDecimalNumber>,
    #[serde(rename = "@w:cryptSpinCount")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub crypt_spin_count: Option<STDecimalNumber>,
    #[serde(rename = "@w:cryptProvider")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub crypt_provider: Option<STString>,
    #[serde(rename = "@w:algIdExt")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub alg_id_ext: Option<STLongHexNumber>,
    #[serde(rename = "@w:algIdExtSource")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub alg_id_ext_source: Option<STString>,
    #[serde(rename = "@w:cryptProviderTypeExt")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub crypt_provider_type_ext: Option<STLongHexNumber>,
    #[serde(rename = "@w:cryptProviderTypeExtSource")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub crypt_provider_type_ext_source: Option<STString>,
    #[serde(rename = "@w:hash")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub hash: Option<Vec<u8>>,
    #[serde(rename = "@w:salt")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub salt: Option<Vec<u8>>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct CTDocProtect {
    #[cfg(feature = "wml-settings")]
    #[serde(rename = "@w:edit")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub edit: Option<STDocProtect>,
    #[cfg(feature = "wml-settings")]
    #[serde(rename = "@w:formatting")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub formatting: Option<OnOff>,
    #[cfg(feature = "wml-settings")]
    #[serde(rename = "@w:enforcement")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub enforcement: Option<OnOff>,
    #[cfg(feature = "wml-settings")]
    #[serde(rename = "@w:algorithmName")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub algorithm_name: Option<STString>,
    #[cfg(feature = "wml-settings")]
    #[serde(rename = "@w:hashValue")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub hash_value: Option<Vec<u8>>,
    #[cfg(feature = "wml-settings")]
    #[serde(rename = "@w:saltValue")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub salt_value: Option<Vec<u8>>,
    #[cfg(feature = "wml-settings")]
    #[serde(rename = "@w:spinCount")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub spin_count: Option<STDecimalNumber>,
    #[cfg(feature = "wml-settings")]
    #[serde(rename = "@w:cryptProviderType")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub crypt_provider_type: Option<STCryptProv>,
    #[cfg(feature = "wml-settings")]
    #[serde(rename = "@w:cryptAlgorithmClass")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub crypt_algorithm_class: Option<STAlgClass>,
    #[cfg(feature = "wml-settings")]
    #[serde(rename = "@w:cryptAlgorithmType")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub crypt_algorithm_type: Option<STAlgType>,
    #[cfg(feature = "wml-settings")]
    #[serde(rename = "@w:cryptAlgorithmSid")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub crypt_algorithm_sid: Option<STDecimalNumber>,
    #[cfg(feature = "wml-settings")]
    #[serde(rename = "@w:cryptSpinCount")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub crypt_spin_count: Option<STDecimalNumber>,
    #[cfg(feature = "wml-settings")]
    #[serde(rename = "@w:cryptProvider")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub crypt_provider: Option<STString>,
    #[cfg(feature = "wml-settings")]
    #[serde(rename = "@w:algIdExt")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub alg_id_ext: Option<STLongHexNumber>,
    #[cfg(feature = "wml-settings")]
    #[serde(rename = "@w:algIdExtSource")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub alg_id_ext_source: Option<STString>,
    #[cfg(feature = "wml-settings")]
    #[serde(rename = "@w:cryptProviderTypeExt")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub crypt_provider_type_ext: Option<STLongHexNumber>,
    #[cfg(feature = "wml-settings")]
    #[serde(rename = "@w:cryptProviderTypeExtSource")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub crypt_provider_type_ext_source: Option<STString>,
    #[cfg(feature = "wml-settings")]
    #[serde(rename = "@w:hash")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub hash: Option<Vec<u8>>,
    #[cfg(feature = "wml-settings")]
    #[serde(rename = "@w:salt")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub salt: Option<Vec<u8>>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTMailMergeDocType {
    #[serde(rename = "@w:val")]
    pub value: STMailMergeDocType,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTMailMergeDataType {
    #[serde(rename = "@w:val")]
    pub value: STMailMergeDataType,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTMailMergeDest {
    #[serde(rename = "@w:val")]
    pub value: STMailMergeDest,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTMailMergeOdsoFMDFieldType {
    #[serde(rename = "@w:val")]
    pub value: STMailMergeOdsoFMDFieldType,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct CTTrackChangesView {
    #[cfg(feature = "wml-settings")]
    #[serde(rename = "@w:markup")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub markup: Option<OnOff>,
    #[cfg(feature = "wml-settings")]
    #[serde(rename = "@w:comments")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub comments: Option<OnOff>,
    #[cfg(feature = "wml-settings")]
    #[serde(rename = "@w:insDel")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub ins_del: Option<OnOff>,
    #[cfg(feature = "wml-settings")]
    #[serde(rename = "@w:formatting")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub formatting: Option<OnOff>,
    #[cfg(feature = "wml-settings")]
    #[serde(rename = "@w:inkAnnotations")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub ink_annotations: Option<OnOff>,
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
    #[cfg(feature = "wml-styling")]
    #[serde(rename = "@w:lang")]
    pub lang: Language,
    #[cfg(feature = "wml-styling")]
    #[serde(rename = "@w:val")]
    pub value: STString,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTTextDirection {
    #[serde(rename = "@w:val")]
    pub value: STTextDirection,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTTextAlignment {
    #[serde(rename = "@w:val")]
    pub value: STTextAlignment,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTMarkup {
    #[serde(rename = "@w:id")]
    pub id: STDecimalNumber,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTTrackChange {
    #[serde(rename = "@w:id")]
    pub id: STDecimalNumber,
    #[serde(rename = "@w:author")]
    pub author: STString,
    #[cfg(feature = "wml-track-changes")]
    #[serde(rename = "@w:date")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub date: Option<STDateTime>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTCellMergeTrackChange {
    #[serde(rename = "@w:id")]
    pub id: STDecimalNumber,
    #[serde(rename = "@w:author")]
    pub author: STString,
    #[serde(rename = "@w:date")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub date: Option<STDateTime>,
    #[cfg(feature = "wml-track-changes")]
    #[serde(rename = "@w:vMerge")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub vertical_merge: Option<STAnnotationVMerge>,
    #[cfg(feature = "wml-track-changes")]
    #[serde(rename = "@w:vMergeOrig")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub v_merge_orig: Option<STAnnotationVMerge>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTTrackChangeRange {
    #[serde(rename = "@w:id")]
    pub id: STDecimalNumber,
    #[serde(rename = "@w:author")]
    pub author: STString,
    #[serde(rename = "@w:date")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub date: Option<STDateTime>,
    #[cfg(feature = "wml-settings")]
    #[serde(rename = "@w:displacedByCustomXml")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub displaced_by_custom_xml: Option<STDisplacedByCustomXml>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTMarkupRange {
    #[serde(rename = "@w:id")]
    pub id: STDecimalNumber,
    #[cfg(feature = "wml-settings")]
    #[serde(rename = "@w:displacedByCustomXml")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub displaced_by_custom_xml: Option<STDisplacedByCustomXml>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTBookmarkRange {
    #[serde(rename = "@w:id")]
    pub id: STDecimalNumber,
    #[serde(rename = "@w:displacedByCustomXml")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub displaced_by_custom_xml: Option<STDisplacedByCustomXml>,
    #[cfg(feature = "wml-tables")]
    #[serde(rename = "@w:colFirst")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub col_first: Option<STDecimalNumber>,
    #[cfg(feature = "wml-tables")]
    #[serde(rename = "@w:colLast")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub col_last: Option<STDecimalNumber>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct Bookmark {
    #[serde(rename = "@w:id")]
    pub id: STDecimalNumber,
    #[cfg(feature = "wml-settings")]
    #[serde(rename = "@w:displacedByCustomXml")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub displaced_by_custom_xml: Option<STDisplacedByCustomXml>,
    #[cfg(feature = "wml-tables")]
    #[serde(rename = "@w:colFirst")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub col_first: Option<STDecimalNumber>,
    #[cfg(feature = "wml-tables")]
    #[serde(rename = "@w:colLast")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub col_last: Option<STDecimalNumber>,
    #[serde(rename = "@w:name")]
    pub name: STString,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTMoveBookmark {
    #[serde(rename = "@w:id")]
    pub id: STDecimalNumber,
    #[serde(rename = "@w:displacedByCustomXml")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub displaced_by_custom_xml: Option<STDisplacedByCustomXml>,
    #[serde(rename = "@w:colFirst")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub col_first: Option<STDecimalNumber>,
    #[serde(rename = "@w:colLast")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub col_last: Option<STDecimalNumber>,
    #[serde(rename = "@w:name")]
    pub name: STString,
    #[cfg(feature = "wml-track-changes")]
    #[serde(rename = "@w:author")]
    pub author: STString,
    #[cfg(feature = "wml-track-changes")]
    #[serde(rename = "@w:date")]
    pub date: STDateTime,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct Comment {
    #[serde(rename = "@w:id")]
    pub id: STDecimalNumber,
    #[serde(rename = "@w:author")]
    pub author: STString,
    #[cfg(feature = "wml-comments")]
    #[serde(rename = "@w:date")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub date: Option<STDateTime>,
    #[serde(skip)]
    #[serde(default)]
    pub block_content: Vec<BlockContent>,
    #[cfg(feature = "wml-comments")]
    #[serde(rename = "@w:initials")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub initials: Option<STString>,
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
pub struct CTTrackChangeNumbering {
    #[serde(rename = "@w:id")]
    pub id: STDecimalNumber,
    #[serde(rename = "@w:author")]
    pub author: STString,
    #[serde(rename = "@w:date")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub date: Option<STDateTime>,
    #[cfg(feature = "wml-track-changes")]
    #[serde(rename = "@w:original")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub original: Option<STString>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTTblPrExChange {
    #[serde(rename = "@w:id")]
    pub id: STDecimalNumber,
    #[serde(rename = "@w:author")]
    pub author: STString,
    #[serde(rename = "@w:date")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub date: Option<STDateTime>,
    #[cfg(feature = "wml-track-changes")]
    #[serde(rename = "tblPrEx")]
    pub tbl_pr_ex: Box<CTTblPrExBase>,
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
pub struct CTTcPrChange {
    #[serde(rename = "@w:id")]
    pub id: STDecimalNumber,
    #[serde(rename = "@w:author")]
    pub author: STString,
    #[serde(rename = "@w:date")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub date: Option<STDateTime>,
    #[cfg(feature = "wml-track-changes")]
    #[serde(rename = "tcPr")]
    pub cell_properties: Box<CTTcPrInner>,
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
pub struct CTTrPrChange {
    #[serde(rename = "@w:id")]
    pub id: STDecimalNumber,
    #[serde(rename = "@w:author")]
    pub author: STString,
    #[serde(rename = "@w:date")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub date: Option<STDateTime>,
    #[cfg(feature = "wml-track-changes")]
    #[serde(rename = "trPr")]
    pub row_properties: Box<CTTrPrBase>,
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
pub struct CTTblGridChange {
    #[serde(rename = "@w:id")]
    pub id: STDecimalNumber,
    #[cfg(feature = "wml-track-changes")]
    #[serde(rename = "tblGrid")]
    pub tbl_grid: Box<CTTblGridBase>,
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
pub struct CTTblPrChange {
    #[serde(rename = "@w:id")]
    pub id: STDecimalNumber,
    #[serde(rename = "@w:author")]
    pub author: STString,
    #[serde(rename = "@w:date")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub date: Option<STDateTime>,
    #[cfg(feature = "wml-track-changes")]
    #[serde(rename = "tblPr")]
    pub table_properties: Box<CTTblPrBase>,
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
pub struct CTSectPrChange {
    #[serde(rename = "@w:id")]
    pub id: STDecimalNumber,
    #[serde(rename = "@w:author")]
    pub author: STString,
    #[serde(rename = "@w:date")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub date: Option<STDateTime>,
    #[cfg(feature = "wml-track-changes")]
    #[serde(rename = "sectPr")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub sect_pr: Option<Box<CTSectPrBase>>,
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
pub struct CTPPrChange {
    #[serde(rename = "@w:id")]
    pub id: STDecimalNumber,
    #[serde(rename = "@w:author")]
    pub author: STString,
    #[serde(rename = "@w:date")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub date: Option<STDateTime>,
    #[cfg(feature = "wml-track-changes")]
    #[serde(rename = "pPr")]
    pub p_pr: Box<CTPPrBase>,
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
pub struct CTRPrChange {
    #[serde(rename = "@w:id")]
    pub id: STDecimalNumber,
    #[serde(rename = "@w:author")]
    pub author: STString,
    #[serde(rename = "@w:date")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub date: Option<STDateTime>,
    #[cfg(feature = "wml-track-changes")]
    #[serde(rename = "rPr")]
    pub r_pr: Box<CTRPrOriginal>,
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
pub struct CTParaRPrChange {
    #[serde(rename = "@w:id")]
    pub id: STDecimalNumber,
    #[serde(rename = "@w:author")]
    pub author: STString,
    #[serde(rename = "@w:date")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub date: Option<STDateTime>,
    #[cfg(feature = "wml-track-changes")]
    #[serde(rename = "rPr")]
    pub r_pr: Box<CTParaRPrOriginal>,
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
pub struct CTRunTrackChange {
    #[cfg(feature = "wml-track-changes")]
    #[serde(rename = "@w:id")]
    pub id: STDecimalNumber,
    #[cfg(feature = "wml-track-changes")]
    #[serde(rename = "@w:author")]
    pub author: STString,
    #[cfg(feature = "wml-track-changes")]
    #[serde(rename = "@w:date")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub date: Option<STDateTime>,
    #[cfg(feature = "wml-track-changes")]
    #[serde(skip)]
    #[serde(default)]
    pub run_content: Vec<RunContentChoice>,
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
pub struct MathContent {
    #[serde(skip)]
    #[serde(default)]
    pub p_content_base: Vec<ParagraphContentBase>,
    #[serde(skip)]
    #[serde(default)]
    pub content_run_content_base: Vec<RunContentBase>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct NumberingProperties {
    #[cfg(feature = "wml-numbering")]
    #[serde(rename = "ilvl")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub ilvl: Option<Box<CTDecimalNumber>>,
    #[cfg(feature = "wml-numbering")]
    #[serde(rename = "numId")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub num_id: Option<Box<CTDecimalNumber>>,
    #[cfg(feature = "wml-track-changes")]
    #[serde(rename = "numberingChange")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub numbering_change: Option<Box<CTTrackChangeNumbering>>,
    #[cfg(feature = "wml-track-changes")]
    #[serde(rename = "ins")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub ins: Option<Box<CTTrackChange>>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct CTPBdr {
    #[cfg(feature = "wml-styling")]
    #[serde(rename = "top")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub top: Option<Box<CTBorder>>,
    #[cfg(feature = "wml-styling")]
    #[serde(rename = "left")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub left: Option<Box<CTBorder>>,
    #[cfg(feature = "wml-styling")]
    #[serde(rename = "bottom")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub bottom: Option<Box<CTBorder>>,
    #[cfg(feature = "wml-styling")]
    #[serde(rename = "right")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub right: Option<Box<CTBorder>>,
    #[cfg(feature = "wml-styling")]
    #[serde(rename = "between")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub between: Option<Box<CTBorder>>,
    #[cfg(feature = "wml-styling")]
    #[serde(rename = "bar")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub bar: Option<Box<CTBorder>>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct CTTabs {
    #[serde(rename = "tab")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub tab: Vec<CTTabStop>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTTextboxTightWrap {
    #[serde(rename = "@w:val")]
    pub value: STTextboxTightWrap,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct ParagraphProperties {
    #[serde(rename = "pStyle")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub paragraph_style: Option<Box<CTString>>,
    #[cfg(feature = "wml-layout")]
    #[serde(rename = "keepNext")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub keep_next: Option<Box<CTOnOff>>,
    #[cfg(feature = "wml-layout")]
    #[serde(rename = "keepLines")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub keep_lines: Option<Box<CTOnOff>>,
    #[cfg(feature = "wml-layout")]
    #[serde(rename = "pageBreakBefore")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub page_break_before: Option<Box<CTOnOff>>,
    #[cfg(feature = "wml-layout")]
    #[serde(rename = "framePr")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub frame_pr: Option<Box<CTFramePr>>,
    #[cfg(feature = "wml-layout")]
    #[serde(rename = "widowControl")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub widow_control: Option<Box<CTOnOff>>,
    #[cfg(feature = "wml-numbering")]
    #[serde(rename = "numPr")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub num_pr: Option<Box<NumberingProperties>>,
    #[cfg(feature = "wml-layout")]
    #[serde(rename = "suppressLineNumbers")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub suppress_line_numbers: Option<Box<CTOnOff>>,
    #[cfg(feature = "wml-styling")]
    #[serde(rename = "pBdr")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub paragraph_border: Option<Box<CTPBdr>>,
    #[cfg(feature = "wml-styling")]
    #[serde(rename = "shd")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub shading: Option<Box<CTShd>>,
    #[cfg(feature = "wml-styling")]
    #[serde(rename = "tabs")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub tabs: Option<Box<CTTabs>>,
    #[cfg(feature = "wml-styling")]
    #[serde(rename = "suppressAutoHyphens")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub suppress_auto_hyphens: Option<Box<CTOnOff>>,
    #[cfg(feature = "wml-styling")]
    #[serde(rename = "kinsoku")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub kinsoku: Option<Box<CTOnOff>>,
    #[cfg(feature = "wml-styling")]
    #[serde(rename = "wordWrap")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub word_wrap: Option<Box<CTOnOff>>,
    #[cfg(feature = "wml-styling")]
    #[serde(rename = "overflowPunct")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub overflow_punct: Option<Box<CTOnOff>>,
    #[cfg(feature = "wml-styling")]
    #[serde(rename = "topLinePunct")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub top_line_punct: Option<Box<CTOnOff>>,
    #[cfg(feature = "wml-styling")]
    #[serde(rename = "autoSpaceDE")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub auto_space_d_e: Option<Box<CTOnOff>>,
    #[cfg(feature = "wml-styling")]
    #[serde(rename = "autoSpaceDN")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub auto_space_d_n: Option<Box<CTOnOff>>,
    #[cfg(feature = "wml-styling")]
    #[serde(rename = "bidi")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub bidi: Option<Box<CTOnOff>>,
    #[cfg(feature = "wml-styling")]
    #[serde(rename = "adjustRightInd")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub adjust_right_ind: Option<Box<CTOnOff>>,
    #[cfg(feature = "wml-layout")]
    #[serde(rename = "snapToGrid")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub snap_to_grid: Option<Box<CTOnOff>>,
    #[cfg(feature = "wml-styling")]
    #[serde(rename = "spacing")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub spacing: Option<Box<CTSpacing>>,
    #[cfg(feature = "wml-styling")]
    #[serde(rename = "ind")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub indentation: Option<Box<CTInd>>,
    #[cfg(feature = "wml-styling")]
    #[serde(rename = "contextualSpacing")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub contextual_spacing: Option<Box<CTOnOff>>,
    #[cfg(feature = "wml-styling")]
    #[serde(rename = "mirrorIndents")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub mirror_indents: Option<Box<CTOnOff>>,
    #[cfg(feature = "wml-layout")]
    #[serde(rename = "suppressOverlap")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub suppress_overlap: Option<Box<CTOnOff>>,
    #[cfg(feature = "wml-styling")]
    #[serde(rename = "jc")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub justification: Option<Box<CTJc>>,
    #[cfg(feature = "wml-styling")]
    #[serde(rename = "textDirection")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub text_direction: Option<Box<CTTextDirection>>,
    #[cfg(feature = "wml-styling")]
    #[serde(rename = "textAlignment")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub text_alignment: Option<Box<CTTextAlignment>>,
    #[cfg(feature = "wml-styling")]
    #[serde(rename = "textboxTightWrap")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub textbox_tight_wrap: Option<Box<CTTextboxTightWrap>>,
    #[cfg(feature = "wml-styling")]
    #[serde(rename = "outlineLvl")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub outline_lvl: Option<Box<CTDecimalNumber>>,
    #[cfg(feature = "wml-styling")]
    #[serde(rename = "divId")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub div_id: Option<Box<CTDecimalNumber>>,
    #[cfg(feature = "wml-styling")]
    #[serde(rename = "cnfStyle")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub cnf_style: Option<Box<CTCnf>>,
    #[cfg(feature = "wml-styling")]
    #[serde(rename = "rPr")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub r_pr: Option<Box<CTParaRPr>>,
    #[cfg(feature = "wml-layout")]
    #[serde(rename = "sectPr")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub sect_pr: Option<Box<SectionProperties>>,
    #[cfg(feature = "wml-track-changes")]
    #[serde(rename = "pPrChange")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub p_pr_change: Option<Box<CTPPrChange>>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct CTPPrBase {
    #[serde(rename = "pStyle")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub paragraph_style: Option<Box<CTString>>,
    #[cfg(feature = "wml-layout")]
    #[serde(rename = "keepNext")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub keep_next: Option<Box<CTOnOff>>,
    #[cfg(feature = "wml-layout")]
    #[serde(rename = "keepLines")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub keep_lines: Option<Box<CTOnOff>>,
    #[cfg(feature = "wml-layout")]
    #[serde(rename = "pageBreakBefore")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub page_break_before: Option<Box<CTOnOff>>,
    #[cfg(feature = "wml-layout")]
    #[serde(rename = "framePr")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub frame_pr: Option<Box<CTFramePr>>,
    #[cfg(feature = "wml-layout")]
    #[serde(rename = "widowControl")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub widow_control: Option<Box<CTOnOff>>,
    #[cfg(feature = "wml-numbering")]
    #[serde(rename = "numPr")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub num_pr: Option<Box<NumberingProperties>>,
    #[cfg(feature = "wml-layout")]
    #[serde(rename = "suppressLineNumbers")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub suppress_line_numbers: Option<Box<CTOnOff>>,
    #[cfg(feature = "wml-styling")]
    #[serde(rename = "pBdr")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub paragraph_border: Option<Box<CTPBdr>>,
    #[cfg(feature = "wml-styling")]
    #[serde(rename = "shd")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub shading: Option<Box<CTShd>>,
    #[cfg(feature = "wml-styling")]
    #[serde(rename = "tabs")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub tabs: Option<Box<CTTabs>>,
    #[cfg(feature = "wml-styling")]
    #[serde(rename = "suppressAutoHyphens")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub suppress_auto_hyphens: Option<Box<CTOnOff>>,
    #[cfg(feature = "wml-styling")]
    #[serde(rename = "kinsoku")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub kinsoku: Option<Box<CTOnOff>>,
    #[cfg(feature = "wml-styling")]
    #[serde(rename = "wordWrap")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub word_wrap: Option<Box<CTOnOff>>,
    #[cfg(feature = "wml-styling")]
    #[serde(rename = "overflowPunct")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub overflow_punct: Option<Box<CTOnOff>>,
    #[cfg(feature = "wml-styling")]
    #[serde(rename = "topLinePunct")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub top_line_punct: Option<Box<CTOnOff>>,
    #[cfg(feature = "wml-styling")]
    #[serde(rename = "autoSpaceDE")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub auto_space_d_e: Option<Box<CTOnOff>>,
    #[cfg(feature = "wml-styling")]
    #[serde(rename = "autoSpaceDN")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub auto_space_d_n: Option<Box<CTOnOff>>,
    #[cfg(feature = "wml-styling")]
    #[serde(rename = "bidi")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub bidi: Option<Box<CTOnOff>>,
    #[cfg(feature = "wml-styling")]
    #[serde(rename = "adjustRightInd")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub adjust_right_ind: Option<Box<CTOnOff>>,
    #[cfg(feature = "wml-layout")]
    #[serde(rename = "snapToGrid")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub snap_to_grid: Option<Box<CTOnOff>>,
    #[cfg(feature = "wml-styling")]
    #[serde(rename = "spacing")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub spacing: Option<Box<CTSpacing>>,
    #[cfg(feature = "wml-styling")]
    #[serde(rename = "ind")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub indentation: Option<Box<CTInd>>,
    #[cfg(feature = "wml-styling")]
    #[serde(rename = "contextualSpacing")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub contextual_spacing: Option<Box<CTOnOff>>,
    #[cfg(feature = "wml-styling")]
    #[serde(rename = "mirrorIndents")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub mirror_indents: Option<Box<CTOnOff>>,
    #[cfg(feature = "wml-layout")]
    #[serde(rename = "suppressOverlap")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub suppress_overlap: Option<Box<CTOnOff>>,
    #[cfg(feature = "wml-styling")]
    #[serde(rename = "jc")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub justification: Option<Box<CTJc>>,
    #[cfg(feature = "wml-styling")]
    #[serde(rename = "textDirection")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub text_direction: Option<Box<CTTextDirection>>,
    #[cfg(feature = "wml-styling")]
    #[serde(rename = "textAlignment")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub text_alignment: Option<Box<CTTextAlignment>>,
    #[cfg(feature = "wml-styling")]
    #[serde(rename = "textboxTightWrap")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub textbox_tight_wrap: Option<Box<CTTextboxTightWrap>>,
    #[cfg(feature = "wml-styling")]
    #[serde(rename = "outlineLvl")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub outline_lvl: Option<Box<CTDecimalNumber>>,
    #[cfg(feature = "wml-styling")]
    #[serde(rename = "divId")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub div_id: Option<Box<CTDecimalNumber>>,
    #[cfg(feature = "wml-styling")]
    #[serde(rename = "cnfStyle")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub cnf_style: Option<Box<CTCnf>>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct CTPPrGeneral {
    #[serde(rename = "pStyle")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub paragraph_style: Option<Box<CTString>>,
    #[serde(rename = "keepNext")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub keep_next: Option<Box<CTOnOff>>,
    #[serde(rename = "keepLines")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub keep_lines: Option<Box<CTOnOff>>,
    #[serde(rename = "pageBreakBefore")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub page_break_before: Option<Box<CTOnOff>>,
    #[serde(rename = "framePr")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub frame_pr: Option<Box<CTFramePr>>,
    #[serde(rename = "widowControl")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub widow_control: Option<Box<CTOnOff>>,
    #[serde(rename = "numPr")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub num_pr: Option<Box<NumberingProperties>>,
    #[serde(rename = "suppressLineNumbers")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub suppress_line_numbers: Option<Box<CTOnOff>>,
    #[serde(rename = "pBdr")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub paragraph_border: Option<Box<CTPBdr>>,
    #[serde(rename = "shd")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub shading: Option<Box<CTShd>>,
    #[serde(rename = "tabs")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub tabs: Option<Box<CTTabs>>,
    #[serde(rename = "suppressAutoHyphens")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub suppress_auto_hyphens: Option<Box<CTOnOff>>,
    #[serde(rename = "kinsoku")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub kinsoku: Option<Box<CTOnOff>>,
    #[serde(rename = "wordWrap")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub word_wrap: Option<Box<CTOnOff>>,
    #[serde(rename = "overflowPunct")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub overflow_punct: Option<Box<CTOnOff>>,
    #[serde(rename = "topLinePunct")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub top_line_punct: Option<Box<CTOnOff>>,
    #[serde(rename = "autoSpaceDE")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub auto_space_d_e: Option<Box<CTOnOff>>,
    #[serde(rename = "autoSpaceDN")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub auto_space_d_n: Option<Box<CTOnOff>>,
    #[serde(rename = "bidi")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub bidi: Option<Box<CTOnOff>>,
    #[serde(rename = "adjustRightInd")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub adjust_right_ind: Option<Box<CTOnOff>>,
    #[serde(rename = "snapToGrid")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub snap_to_grid: Option<Box<CTOnOff>>,
    #[serde(rename = "spacing")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub spacing: Option<Box<CTSpacing>>,
    #[serde(rename = "ind")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub indentation: Option<Box<CTInd>>,
    #[serde(rename = "contextualSpacing")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub contextual_spacing: Option<Box<CTOnOff>>,
    #[serde(rename = "mirrorIndents")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub mirror_indents: Option<Box<CTOnOff>>,
    #[serde(rename = "suppressOverlap")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub suppress_overlap: Option<Box<CTOnOff>>,
    #[serde(rename = "jc")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub justification: Option<Box<CTJc>>,
    #[serde(rename = "textDirection")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub text_direction: Option<Box<CTTextDirection>>,
    #[serde(rename = "textAlignment")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub text_alignment: Option<Box<CTTextAlignment>>,
    #[serde(rename = "textboxTightWrap")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub textbox_tight_wrap: Option<Box<CTTextboxTightWrap>>,
    #[serde(rename = "outlineLvl")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub outline_lvl: Option<Box<CTDecimalNumber>>,
    #[serde(rename = "divId")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub div_id: Option<Box<CTDecimalNumber>>,
    #[serde(rename = "cnfStyle")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub cnf_style: Option<Box<CTCnf>>,
    #[cfg(feature = "wml-track-changes")]
    #[serde(rename = "pPrChange")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub p_pr_change: Option<Box<CTPPrChange>>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct CTControl {
    #[cfg(feature = "wml-drawings")]
    #[serde(rename = "@w:name")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub name: Option<STString>,
    #[cfg(feature = "wml-drawings")]
    #[serde(rename = "@w:shapeid")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub shapeid: Option<STString>,
    #[cfg(feature = "wml-drawings")]
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
pub struct CTBackground {
    #[cfg(feature = "wml-styling")]
    #[serde(rename = "@w:color")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub color: Option<STHexColor>,
    #[cfg(feature = "wml-styling")]
    #[serde(rename = "@w:themeColor")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub theme_color: Option<STThemeColor>,
    #[cfg(feature = "wml-styling")]
    #[serde(rename = "@w:themeTint")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub theme_tint: Option<STUcharHexNumber>,
    #[cfg(feature = "wml-styling")]
    #[serde(rename = "@w:themeShade")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub theme_shade: Option<STUcharHexNumber>,
    #[cfg(feature = "wml-drawings")]
    #[serde(rename = "drawing")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub drawing: Option<Box<CTDrawing>>,
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
pub struct CTRel {
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
pub struct CTObject {
    #[cfg(feature = "wml-drawings")]
    #[serde(rename = "@w:dxaOrig")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub dxa_orig: Option<STTwipsMeasure>,
    #[cfg(feature = "wml-drawings")]
    #[serde(rename = "@w:dyaOrig")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub dya_orig: Option<STTwipsMeasure>,
    #[cfg(feature = "wml-drawings")]
    #[serde(rename = "drawing")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub drawing: Option<Box<CTDrawing>>,
    #[cfg(feature = "wml-drawings")]
    #[serde(rename = "control")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub control: Option<Box<CTControl>>,
    #[cfg(feature = "wml-drawings")]
    #[serde(rename = "objectLink")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub object_link: Option<Box<CTObjectLink>>,
    #[cfg(feature = "wml-drawings")]
    #[serde(rename = "objectEmbed")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub object_embed: Option<Box<CTObjectEmbed>>,
    #[cfg(feature = "wml-drawings")]
    #[serde(rename = "movie")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub movie: Option<Box<CTRel>>,
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
pub struct CTPicture {
    #[cfg(feature = "wml-drawings")]
    #[serde(rename = "movie")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub movie: Option<Box<CTRel>>,
    #[cfg(feature = "wml-drawings")]
    #[serde(rename = "control")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub control: Option<Box<CTControl>>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTObjectEmbed {
    #[cfg(feature = "wml-drawings")]
    #[serde(rename = "@w:drawAspect")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub draw_aspect: Option<STObjectDrawAspect>,
    #[cfg(feature = "wml-drawings")]
    #[serde(rename = "@r:id")]
    pub id: STRelationshipId,
    #[cfg(feature = "wml-drawings")]
    #[serde(rename = "@w:progId")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub prog_id: Option<STString>,
    #[cfg(feature = "wml-drawings")]
    #[serde(rename = "@w:shapeId")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub shape_id: Option<STString>,
    #[cfg(feature = "wml-drawings")]
    #[serde(rename = "@w:fieldCodes")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub field_codes: Option<STString>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTObjectLink {
    #[cfg(feature = "wml-drawings")]
    #[serde(rename = "@w:drawAspect")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub draw_aspect: Option<STObjectDrawAspect>,
    #[cfg(feature = "wml-drawings")]
    #[serde(rename = "@r:id")]
    pub id: STRelationshipId,
    #[cfg(feature = "wml-drawings")]
    #[serde(rename = "@w:progId")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub prog_id: Option<STString>,
    #[cfg(feature = "wml-drawings")]
    #[serde(rename = "@w:shapeId")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub shape_id: Option<STString>,
    #[cfg(feature = "wml-drawings")]
    #[serde(rename = "@w:fieldCodes")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub field_codes: Option<STString>,
    #[cfg(feature = "wml-drawings")]
    #[serde(rename = "@w:updateMode")]
    pub update_mode: STObjectUpdateMode,
    #[cfg(feature = "wml-drawings")]
    #[serde(rename = "@w:lockedField")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub locked_field: Option<OnOff>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct CTDrawing {
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTSimpleField {
    #[serde(rename = "@w:instr")]
    pub instr: STString,
    #[cfg(feature = "wml-fields")]
    #[serde(rename = "@w:fldLock")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub fld_lock: Option<OnOff>,
    #[cfg(feature = "wml-fields")]
    #[serde(rename = "@w:dirty")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub dirty: Option<OnOff>,
    #[cfg(feature = "wml-fields")]
    #[serde(rename = "fldData")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub fld_data: Option<Box<Text>>,
    #[serde(skip)]
    #[serde(default)]
    pub paragraph_content: Vec<ParagraphContent>,
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
pub struct CTFFTextType {
    #[serde(rename = "@w:val")]
    pub value: STFFTextType,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct CTFFName {
    #[serde(rename = "@w:val")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub value: Option<STFFName>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTFldChar {
    #[serde(rename = "@w:fldCharType")]
    pub fld_char_type: STFldCharType,
    #[cfg(feature = "wml-fields")]
    #[serde(rename = "@w:fldLock")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub fld_lock: Option<OnOff>,
    #[cfg(feature = "wml-fields")]
    #[serde(rename = "@w:dirty")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub dirty: Option<OnOff>,
    #[cfg(feature = "wml-fields")]
    #[serde(rename = "fldData")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub fld_data: Option<Box<Text>>,
    #[cfg(feature = "wml-fields")]
    #[serde(rename = "ffData")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub ff_data: Option<Box<CTFFData>>,
    #[cfg(feature = "wml-track-changes")]
    #[serde(rename = "numberingChange")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub numbering_change: Option<Box<CTTrackChangeNumbering>>,
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
pub struct Hyperlink {
    #[cfg(feature = "wml-hyperlinks")]
    #[serde(rename = "@w:tgtFrame")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub tgt_frame: Option<STString>,
    #[cfg(feature = "wml-hyperlinks")]
    #[serde(rename = "@w:tooltip")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub tooltip: Option<STString>,
    #[cfg(feature = "wml-hyperlinks")]
    #[serde(rename = "@w:docLocation")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub doc_location: Option<STString>,
    #[cfg(feature = "wml-hyperlinks")]
    #[serde(rename = "@w:history")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub history: Option<OnOff>,
    #[cfg(feature = "wml-hyperlinks")]
    #[serde(rename = "@w:anchor")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub anchor: Option<STString>,
    #[cfg(feature = "wml-hyperlinks")]
    #[serde(rename = "@r:id")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub id: Option<STRelationshipId>,
    #[serde(skip)]
    #[serde(default)]
    pub paragraph_content: Vec<ParagraphContent>,
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
pub struct CTFFData {
    #[cfg(feature = "wml-fields")]
    #[serde(rename = "name")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub name: Vec<CTFFName>,
    #[cfg(feature = "wml-fields")]
    #[serde(rename = "label")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub label: Option<Box<CTDecimalNumber>>,
    #[cfg(feature = "wml-fields")]
    #[serde(rename = "tabIndex")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub tab_index: Option<Box<CTUnsignedDecimalNumber>>,
    #[cfg(feature = "wml-fields")]
    #[serde(rename = "enabled")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub enabled: Vec<CTOnOff>,
    #[cfg(feature = "wml-fields")]
    #[serde(rename = "calcOnExit")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub calc_on_exit: Vec<CTOnOff>,
    #[cfg(feature = "wml-fields")]
    #[serde(rename = "entryMacro")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub entry_macro: Option<Box<CTMacroName>>,
    #[cfg(feature = "wml-fields")]
    #[serde(rename = "exitMacro")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub exit_macro: Option<Box<CTMacroName>>,
    #[cfg(feature = "wml-fields")]
    #[serde(rename = "helpText")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub help_text: Option<Box<CTFFHelpText>>,
    #[cfg(feature = "wml-fields")]
    #[serde(rename = "statusText")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub status_text: Option<Box<CTFFStatusText>>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct CTFFHelpText {
    #[serde(rename = "@w:type")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub r#type: Option<STInfoTextType>,
    #[serde(rename = "@w:val")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub value: Option<STFFHelpTextVal>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct CTFFStatusText {
    #[serde(rename = "@w:type")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub r#type: Option<STInfoTextType>,
    #[serde(rename = "@w:val")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub value: Option<STFFStatusTextVal>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct CTFFCheckBox {
    #[cfg(feature = "wml-fields")]
    #[serde(rename = "size")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub size: Option<Box<CTHpsMeasure>>,
    #[cfg(feature = "wml-fields")]
    #[serde(rename = "sizeAuto")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub size_auto: Option<Box<CTOnOff>>,
    #[cfg(feature = "wml-fields")]
    #[serde(rename = "default")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub default: Option<Box<CTOnOff>>,
    #[cfg(feature = "wml-fields")]
    #[serde(rename = "checked")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub checked: Option<Box<CTOnOff>>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct CTFFDDList {
    #[cfg(feature = "wml-fields")]
    #[serde(rename = "result")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub result: Option<Box<CTDecimalNumber>>,
    #[cfg(feature = "wml-fields")]
    #[serde(rename = "default")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub default: Option<Box<CTDecimalNumber>>,
    #[cfg(feature = "wml-fields")]
    #[serde(rename = "listEntry")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub list_entry: Vec<CTString>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct CTFFTextInput {
    #[cfg(feature = "wml-fields")]
    #[serde(rename = "type")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub r#type: Option<Box<CTFFTextType>>,
    #[cfg(feature = "wml-fields")]
    #[serde(rename = "default")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub default: Option<Box<CTString>>,
    #[cfg(feature = "wml-fields")]
    #[serde(rename = "maxLength")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub max_length: Option<Box<CTDecimalNumber>>,
    #[cfg(feature = "wml-fields")]
    #[serde(rename = "format")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub format: Option<Box<CTString>>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct CTSectType {
    #[serde(rename = "@w:val")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub value: Option<STSectionMark>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct CTPaperSource {
    #[cfg(feature = "wml-layout")]
    #[serde(rename = "@w:first")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub first: Option<STDecimalNumber>,
    #[cfg(feature = "wml-layout")]
    #[serde(rename = "@w:other")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub other: Option<STDecimalNumber>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct PageSize {
    #[cfg(feature = "wml-layout")]
    #[serde(rename = "@w:w")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub width: Option<STTwipsMeasure>,
    #[cfg(feature = "wml-layout")]
    #[serde(rename = "@w:h")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub height: Option<STTwipsMeasure>,
    #[cfg(feature = "wml-layout")]
    #[serde(rename = "@w:orient")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub orient: Option<STPageOrientation>,
    #[cfg(feature = "wml-layout")]
    #[serde(rename = "@w:code")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub code: Option<STDecimalNumber>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct PageMargins {
    #[cfg(feature = "wml-layout")]
    #[serde(rename = "@w:top")]
    pub top: STSignedTwipsMeasure,
    #[cfg(feature = "wml-layout")]
    #[serde(rename = "@w:right")]
    pub right: STTwipsMeasure,
    #[cfg(feature = "wml-layout")]
    #[serde(rename = "@w:bottom")]
    pub bottom: STSignedTwipsMeasure,
    #[cfg(feature = "wml-layout")]
    #[serde(rename = "@w:left")]
    pub left: STTwipsMeasure,
    #[cfg(feature = "wml-layout")]
    #[serde(rename = "@w:header")]
    pub header: STTwipsMeasure,
    #[cfg(feature = "wml-layout")]
    #[serde(rename = "@w:footer")]
    pub footer: STTwipsMeasure,
    #[cfg(feature = "wml-layout")]
    #[serde(rename = "@w:gutter")]
    pub gutter: STTwipsMeasure,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct CTPageBorders {
    #[cfg(feature = "wml-layout")]
    #[serde(rename = "@w:zOrder")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub z_order: Option<STPageBorderZOrder>,
    #[cfg(feature = "wml-layout")]
    #[serde(rename = "@w:display")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub display: Option<STPageBorderDisplay>,
    #[cfg(feature = "wml-layout")]
    #[serde(rename = "@w:offsetFrom")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub offset_from: Option<STPageBorderOffset>,
    #[cfg(feature = "wml-layout")]
    #[serde(rename = "top")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub top: Option<Box<CTTopPageBorder>>,
    #[cfg(feature = "wml-layout")]
    #[serde(rename = "left")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub left: Option<Box<CTPageBorder>>,
    #[cfg(feature = "wml-layout")]
    #[serde(rename = "bottom")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub bottom: Option<Box<CTBottomPageBorder>>,
    #[cfg(feature = "wml-layout")]
    #[serde(rename = "right")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub right: Option<Box<CTPageBorder>>,
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
pub struct CTPageBorder {
    #[serde(rename = "@w:val")]
    pub value: STBorder,
    #[serde(rename = "@w:color")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub color: Option<STHexColor>,
    #[serde(rename = "@w:themeColor")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub theme_color: Option<STThemeColor>,
    #[serde(rename = "@w:themeTint")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub theme_tint: Option<STUcharHexNumber>,
    #[serde(rename = "@w:themeShade")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub theme_shade: Option<STUcharHexNumber>,
    #[serde(rename = "@w:sz")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub size: Option<STEighthPointMeasure>,
    #[serde(rename = "@w:space")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub space: Option<STPointMeasure>,
    #[serde(rename = "@w:shadow")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub shadow: Option<OnOff>,
    #[serde(rename = "@w:frame")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub frame: Option<OnOff>,
    #[cfg(feature = "wml-layout")]
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
pub struct CTBottomPageBorder {
    #[serde(rename = "@w:val")]
    pub value: STBorder,
    #[serde(rename = "@w:color")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub color: Option<STHexColor>,
    #[serde(rename = "@w:themeColor")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub theme_color: Option<STThemeColor>,
    #[serde(rename = "@w:themeTint")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub theme_tint: Option<STUcharHexNumber>,
    #[serde(rename = "@w:themeShade")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub theme_shade: Option<STUcharHexNumber>,
    #[serde(rename = "@w:sz")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub size: Option<STEighthPointMeasure>,
    #[serde(rename = "@w:space")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub space: Option<STPointMeasure>,
    #[serde(rename = "@w:shadow")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub shadow: Option<OnOff>,
    #[serde(rename = "@w:frame")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub frame: Option<OnOff>,
    #[serde(rename = "@r:id")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub id: Option<STRelationshipId>,
    #[cfg(feature = "wml-layout")]
    #[serde(rename = "@r:bottomLeft")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub bottom_left: Option<STRelationshipId>,
    #[cfg(feature = "wml-layout")]
    #[serde(rename = "@r:bottomRight")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub bottom_right: Option<STRelationshipId>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTTopPageBorder {
    #[serde(rename = "@w:val")]
    pub value: STBorder,
    #[serde(rename = "@w:color")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub color: Option<STHexColor>,
    #[serde(rename = "@w:themeColor")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub theme_color: Option<STThemeColor>,
    #[serde(rename = "@w:themeTint")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub theme_tint: Option<STUcharHexNumber>,
    #[serde(rename = "@w:themeShade")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub theme_shade: Option<STUcharHexNumber>,
    #[serde(rename = "@w:sz")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub size: Option<STEighthPointMeasure>,
    #[serde(rename = "@w:space")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub space: Option<STPointMeasure>,
    #[serde(rename = "@w:shadow")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub shadow: Option<OnOff>,
    #[serde(rename = "@w:frame")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub frame: Option<OnOff>,
    #[serde(rename = "@r:id")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub id: Option<STRelationshipId>,
    #[cfg(feature = "wml-layout")]
    #[serde(rename = "@r:topLeft")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub top_left: Option<STRelationshipId>,
    #[cfg(feature = "wml-layout")]
    #[serde(rename = "@r:topRight")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub top_right: Option<STRelationshipId>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct CTLineNumber {
    #[cfg(feature = "wml-layout")]
    #[serde(rename = "@w:countBy")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub count_by: Option<STDecimalNumber>,
    #[cfg(feature = "wml-layout")]
    #[serde(rename = "@w:start")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub start: Option<STDecimalNumber>,
    #[cfg(feature = "wml-layout")]
    #[serde(rename = "@w:distance")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub distance: Option<STTwipsMeasure>,
    #[cfg(feature = "wml-layout")]
    #[serde(rename = "@w:restart")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub restart: Option<STLineNumberRestart>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct CTPageNumber {
    #[cfg(feature = "wml-layout")]
    #[serde(rename = "@w:fmt")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub fmt: Option<STNumberFormat>,
    #[cfg(feature = "wml-layout")]
    #[serde(rename = "@w:start")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub start: Option<STDecimalNumber>,
    #[cfg(feature = "wml-layout")]
    #[serde(rename = "@w:chapStyle")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub chap_style: Option<STDecimalNumber>,
    #[cfg(feature = "wml-layout")]
    #[serde(rename = "@w:chapSep")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub chap_sep: Option<STChapterSep>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct CTColumn {
    #[cfg(feature = "wml-layout")]
    #[serde(rename = "@w:w")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub width: Option<STTwipsMeasure>,
    #[cfg(feature = "wml-layout")]
    #[serde(rename = "@w:space")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub space: Option<STTwipsMeasure>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct Columns {
    #[cfg(feature = "wml-layout")]
    #[serde(rename = "@w:equalWidth")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub equal_width: Option<OnOff>,
    #[cfg(feature = "wml-layout")]
    #[serde(rename = "@w:space")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub space: Option<STTwipsMeasure>,
    #[cfg(feature = "wml-layout")]
    #[serde(rename = "@w:num")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub num: Option<STDecimalNumber>,
    #[cfg(feature = "wml-layout")]
    #[serde(rename = "@w:sep")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub sep: Option<OnOff>,
    #[cfg(feature = "wml-layout")]
    #[serde(rename = "col")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub col: Vec<CTColumn>,
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
pub struct CTVerticalJc {
    #[serde(rename = "@w:val")]
    pub value: STVerticalJc,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct DocumentGrid {
    #[cfg(feature = "wml-layout")]
    #[serde(rename = "@w:type")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub r#type: Option<STDocGrid>,
    #[cfg(feature = "wml-layout")]
    #[serde(rename = "@w:linePitch")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub line_pitch: Option<STDecimalNumber>,
    #[cfg(feature = "wml-layout")]
    #[serde(rename = "@w:charSpace")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub char_space: Option<STDecimalNumber>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct HeaderFooterReference {
    #[serde(rename = "@r:id")]
    pub id: STRelationshipId,
    #[serde(rename = "@w:type")]
    pub r#type: STHdrFtr,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct HeaderFooter {
    #[serde(skip)]
    #[serde(default)]
    pub block_content: Vec<BlockContent>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct EGSectPrContents {
    #[serde(rename = "footnotePr")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub footnote_pr: Option<Box<CTFtnProps>>,
    #[serde(rename = "endnotePr")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub endnote_pr: Option<Box<CTEdnProps>>,
    #[serde(rename = "type")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub r#type: Option<Box<CTSectType>>,
    #[serde(rename = "pgSz")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub pg_sz: Option<Box<PageSize>>,
    #[serde(rename = "pgMar")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub pg_mar: Option<Box<PageMargins>>,
    #[serde(rename = "paperSrc")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub paper_src: Option<Box<CTPaperSource>>,
    #[serde(rename = "pgBorders")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub pg_borders: Option<Box<CTPageBorders>>,
    #[serde(rename = "lnNumType")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub ln_num_type: Option<Box<CTLineNumber>>,
    #[serde(rename = "pgNumType")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub pg_num_type: Option<Box<CTPageNumber>>,
    #[serde(rename = "cols")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub cols: Option<Box<Columns>>,
    #[serde(rename = "formProt")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub form_prot: Option<Box<CTOnOff>>,
    #[serde(rename = "vAlign")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub v_align: Option<Box<CTVerticalJc>>,
    #[serde(rename = "noEndnote")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub no_endnote: Option<Box<CTOnOff>>,
    #[serde(rename = "titlePg")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub title_pg: Option<Box<CTOnOff>>,
    #[serde(rename = "textDirection")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub text_direction: Option<Box<CTTextDirection>>,
    #[serde(rename = "bidi")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub bidi: Option<Box<CTOnOff>>,
    #[serde(rename = "rtlGutter")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub rtl_gutter: Option<Box<CTOnOff>>,
    #[serde(rename = "docGrid")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub doc_grid: Option<Box<DocumentGrid>>,
    #[serde(rename = "printerSettings")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub printer_settings: Option<Box<CTRel>>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct WAGSectPrAttributes {
    #[serde(rename = "@w:rsidRPr")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub rsid_r_pr: Option<STLongHexNumber>,
    #[serde(rename = "@w:rsidDel")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub rsid_del: Option<STLongHexNumber>,
    #[serde(rename = "@w:rsidR")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub rsid_r: Option<STLongHexNumber>,
    #[serde(rename = "@w:rsidSect")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub rsid_sect: Option<STLongHexNumber>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct CTSectPrBase {
    #[cfg(feature = "wml-track-changes")]
    #[serde(rename = "@w:rsidRPr")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub rsid_r_pr: Option<STLongHexNumber>,
    #[cfg(feature = "wml-track-changes")]
    #[serde(rename = "@w:rsidDel")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub rsid_del: Option<STLongHexNumber>,
    #[cfg(feature = "wml-track-changes")]
    #[serde(rename = "@w:rsidR")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub rsid_r: Option<STLongHexNumber>,
    #[cfg(feature = "wml-track-changes")]
    #[serde(rename = "@w:rsidSect")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub rsid_sect: Option<STLongHexNumber>,
    #[serde(rename = "footnotePr")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub footnote_pr: Option<Box<CTFtnProps>>,
    #[serde(rename = "endnotePr")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub endnote_pr: Option<Box<CTEdnProps>>,
    #[serde(rename = "type")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub r#type: Option<Box<CTSectType>>,
    #[serde(rename = "pgSz")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub pg_sz: Option<Box<PageSize>>,
    #[serde(rename = "pgMar")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub pg_mar: Option<Box<PageMargins>>,
    #[serde(rename = "paperSrc")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub paper_src: Option<Box<CTPaperSource>>,
    #[serde(rename = "pgBorders")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub pg_borders: Option<Box<CTPageBorders>>,
    #[serde(rename = "lnNumType")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub ln_num_type: Option<Box<CTLineNumber>>,
    #[serde(rename = "pgNumType")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub pg_num_type: Option<Box<CTPageNumber>>,
    #[serde(rename = "cols")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub cols: Option<Box<Columns>>,
    #[serde(rename = "formProt")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub form_prot: Option<Box<CTOnOff>>,
    #[serde(rename = "vAlign")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub v_align: Option<Box<CTVerticalJc>>,
    #[serde(rename = "noEndnote")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub no_endnote: Option<Box<CTOnOff>>,
    #[serde(rename = "titlePg")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub title_pg: Option<Box<CTOnOff>>,
    #[serde(rename = "textDirection")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub text_direction: Option<Box<CTTextDirection>>,
    #[serde(rename = "bidi")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub bidi: Option<Box<CTOnOff>>,
    #[serde(rename = "rtlGutter")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub rtl_gutter: Option<Box<CTOnOff>>,
    #[serde(rename = "docGrid")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub doc_grid: Option<Box<DocumentGrid>>,
    #[serde(rename = "printerSettings")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub printer_settings: Option<Box<CTRel>>,
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
pub struct SectionProperties {
    #[cfg(feature = "wml-track-changes")]
    #[serde(rename = "@w:rsidRPr")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub rsid_r_pr: Option<STLongHexNumber>,
    #[cfg(feature = "wml-track-changes")]
    #[serde(rename = "@w:rsidDel")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub rsid_del: Option<STLongHexNumber>,
    #[cfg(feature = "wml-track-changes")]
    #[serde(rename = "@w:rsidR")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub rsid_r: Option<STLongHexNumber>,
    #[cfg(feature = "wml-track-changes")]
    #[serde(rename = "@w:rsidSect")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub rsid_sect: Option<STLongHexNumber>,
    #[serde(skip)]
    #[serde(default)]
    pub header_footer_refs: Vec<HeaderFooterRef>,
    #[cfg(feature = "wml-comments")]
    #[serde(rename = "footnotePr")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub footnote_pr: Option<Box<CTFtnProps>>,
    #[cfg(feature = "wml-comments")]
    #[serde(rename = "endnotePr")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub endnote_pr: Option<Box<CTEdnProps>>,
    #[cfg(feature = "wml-layout")]
    #[serde(rename = "type")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub r#type: Option<Box<CTSectType>>,
    #[cfg(feature = "wml-layout")]
    #[serde(rename = "pgSz")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub pg_sz: Option<Box<PageSize>>,
    #[cfg(feature = "wml-layout")]
    #[serde(rename = "pgMar")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub pg_mar: Option<Box<PageMargins>>,
    #[cfg(feature = "wml-layout")]
    #[serde(rename = "paperSrc")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub paper_src: Option<Box<CTPaperSource>>,
    #[cfg(feature = "wml-layout")]
    #[serde(rename = "pgBorders")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub pg_borders: Option<Box<CTPageBorders>>,
    #[cfg(feature = "wml-layout")]
    #[serde(rename = "lnNumType")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub ln_num_type: Option<Box<CTLineNumber>>,
    #[cfg(feature = "wml-layout")]
    #[serde(rename = "pgNumType")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub pg_num_type: Option<Box<CTPageNumber>>,
    #[cfg(feature = "wml-layout")]
    #[serde(rename = "cols")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub cols: Option<Box<Columns>>,
    #[cfg(feature = "wml-layout")]
    #[serde(rename = "formProt")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub form_prot: Option<Box<CTOnOff>>,
    #[cfg(feature = "wml-layout")]
    #[serde(rename = "vAlign")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub v_align: Option<Box<CTVerticalJc>>,
    #[cfg(feature = "wml-comments")]
    #[serde(rename = "noEndnote")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub no_endnote: Option<Box<CTOnOff>>,
    #[cfg(feature = "wml-layout")]
    #[serde(rename = "titlePg")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub title_pg: Option<Box<CTOnOff>>,
    #[cfg(feature = "wml-layout")]
    #[serde(rename = "textDirection")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub text_direction: Option<Box<CTTextDirection>>,
    #[cfg(feature = "wml-layout")]
    #[serde(rename = "bidi")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub bidi: Option<Box<CTOnOff>>,
    #[cfg(feature = "wml-layout")]
    #[serde(rename = "rtlGutter")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub rtl_gutter: Option<Box<CTOnOff>>,
    #[cfg(feature = "wml-layout")]
    #[serde(rename = "docGrid")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub doc_grid: Option<Box<DocumentGrid>>,
    #[cfg(feature = "wml-layout")]
    #[serde(rename = "printerSettings")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub printer_settings: Option<Box<CTRel>>,
    #[cfg(feature = "wml-track-changes")]
    #[serde(rename = "sectPrChange")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub sect_pr_change: Option<Box<CTSectPrChange>>,
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
pub struct CTBr {
    #[serde(rename = "@w:type")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub r#type: Option<STBrType>,
    #[serde(rename = "@w:clear")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub clear: Option<STBrClear>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTPTab {
    #[serde(rename = "@w:alignment")]
    pub alignment: STPTabAlignment,
    #[serde(rename = "@w:relativeTo")]
    pub relative_to: STPTabRelativeTo,
    #[cfg(feature = "wml-styling")]
    #[serde(rename = "@w:leader")]
    pub leader: STPTabLeader,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct CTSym {
    #[cfg(feature = "wml-styling")]
    #[serde(rename = "@w:font")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub font: Option<STString>,
    #[serde(rename = "@w:char")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub char: Option<STShortHexNumber>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTProofErr {
    #[serde(rename = "@w:type")]
    pub r#type: STProofErr,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTPerm {
    #[serde(rename = "@w:id")]
    pub id: STString,
    #[cfg(feature = "wml-settings")]
    #[serde(rename = "@w:displacedByCustomXml")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub displaced_by_custom_xml: Option<STDisplacedByCustomXml>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTPermStart {
    #[serde(rename = "@w:id")]
    pub id: STString,
    #[serde(rename = "@w:displacedByCustomXml")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub displaced_by_custom_xml: Option<STDisplacedByCustomXml>,
    #[cfg(feature = "wml-settings")]
    #[serde(rename = "@w:edGrp")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub ed_grp: Option<STEdGrp>,
    #[cfg(feature = "wml-settings")]
    #[serde(rename = "@w:ed")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub ed: Option<STString>,
    #[cfg(feature = "wml-tables")]
    #[serde(rename = "@w:colFirst")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub col_first: Option<STDecimalNumber>,
    #[cfg(feature = "wml-tables")]
    #[serde(rename = "@w:colLast")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub col_last: Option<STDecimalNumber>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct Text {
    #[serde(rename = "$text")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub text: Option<String>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct Run {
    #[cfg(feature = "wml-track-changes")]
    #[serde(rename = "@w:rsidRPr")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub rsid_r_pr: Option<STLongHexNumber>,
    #[cfg(feature = "wml-track-changes")]
    #[serde(rename = "@w:rsidDel")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub rsid_del: Option<STLongHexNumber>,
    #[cfg(feature = "wml-track-changes")]
    #[serde(rename = "@w:rsidR")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub rsid_r: Option<STLongHexNumber>,
    #[cfg(feature = "wml-styling")]
    #[serde(rename = "rPr")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub r_pr: Option<Box<RunProperties>>,
    #[serde(skip)]
    #[serde(default)]
    pub run_content: Vec<RunContent>,
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
pub struct Fonts {
    #[cfg(feature = "wml-styling")]
    #[serde(rename = "@w:hint")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub hint: Option<STHint>,
    #[cfg(feature = "wml-styling")]
    #[serde(rename = "@w:ascii")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub ascii: Option<STString>,
    #[cfg(feature = "wml-styling")]
    #[serde(rename = "@w:hAnsi")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub h_ansi: Option<STString>,
    #[cfg(feature = "wml-styling")]
    #[serde(rename = "@w:eastAsia")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub east_asia: Option<STString>,
    #[cfg(feature = "wml-styling")]
    #[serde(rename = "@w:cs")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub cs: Option<STString>,
    #[cfg(feature = "wml-styling")]
    #[serde(rename = "@w:asciiTheme")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub ascii_theme: Option<STTheme>,
    #[cfg(feature = "wml-styling")]
    #[serde(rename = "@w:hAnsiTheme")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub h_ansi_theme: Option<STTheme>,
    #[cfg(feature = "wml-styling")]
    #[serde(rename = "@w:eastAsiaTheme")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub east_asia_theme: Option<STTheme>,
    #[cfg(feature = "wml-styling")]
    #[serde(rename = "@w:cstheme")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub cstheme: Option<STTheme>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct EGRPrBase {
    #[serde(rename = "rStyle")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub run_style: Option<Box<CTString>>,
    #[serde(rename = "rFonts")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub fonts: Option<Box<Fonts>>,
    #[serde(rename = "b")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub bold: Option<Box<CTOnOff>>,
    #[serde(rename = "bCs")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub b_cs: Option<Box<CTOnOff>>,
    #[serde(rename = "i")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub italic: Option<Box<CTOnOff>>,
    #[serde(rename = "iCs")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub i_cs: Option<Box<CTOnOff>>,
    #[serde(rename = "caps")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub caps: Option<Box<CTOnOff>>,
    #[serde(rename = "smallCaps")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub small_caps: Option<Box<CTOnOff>>,
    #[serde(rename = "strike")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub strikethrough: Option<Box<CTOnOff>>,
    #[serde(rename = "dstrike")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub dstrike: Option<Box<CTOnOff>>,
    #[serde(rename = "outline")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub outline: Option<Box<CTOnOff>>,
    #[serde(rename = "shadow")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub shadow: Option<Box<CTOnOff>>,
    #[serde(rename = "emboss")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub emboss: Option<Box<CTOnOff>>,
    #[serde(rename = "imprint")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub imprint: Option<Box<CTOnOff>>,
    #[serde(rename = "noProof")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub no_proof: Option<Box<CTOnOff>>,
    #[serde(rename = "snapToGrid")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub snap_to_grid: Option<Box<CTOnOff>>,
    #[serde(rename = "vanish")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub vanish: Option<Box<CTOnOff>>,
    #[serde(rename = "webHidden")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub web_hidden: Option<Box<CTOnOff>>,
    #[serde(rename = "color")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub color: Option<Box<CTColor>>,
    #[serde(rename = "spacing")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub spacing: Option<Box<CTSignedTwipsMeasure>>,
    #[serde(rename = "w")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub width: Option<Box<CTTextScale>>,
    #[serde(rename = "kern")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub kern: Option<Box<CTHpsMeasure>>,
    #[serde(rename = "position")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub position: Option<Box<CTSignedHpsMeasure>>,
    #[serde(rename = "sz")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub size: Option<Box<CTHpsMeasure>>,
    #[serde(rename = "szCs")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub size_complex_script: Option<Box<CTHpsMeasure>>,
    #[serde(rename = "highlight")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub highlight: Option<Box<CTHighlight>>,
    #[serde(rename = "u")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub underline: Option<Box<CTUnderline>>,
    #[serde(rename = "effect")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub effect: Option<Box<CTTextEffect>>,
    #[serde(rename = "bdr")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub bdr: Option<Box<CTBorder>>,
    #[serde(rename = "shd")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub shading: Option<Box<CTShd>>,
    #[serde(rename = "fitText")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub fit_text: Option<Box<CTFitText>>,
    #[serde(rename = "vertAlign")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub vert_align: Option<Box<CTVerticalAlignRun>>,
    #[serde(rename = "rtl")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub rtl: Option<Box<CTOnOff>>,
    #[serde(rename = "cs")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub cs: Option<Box<CTOnOff>>,
    #[serde(rename = "em")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub em: Option<Box<CTEm>>,
    #[serde(rename = "lang")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub lang: Option<Box<CTLanguage>>,
    #[serde(rename = "eastAsianLayout")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub east_asian_layout: Option<Box<CTEastAsianLayout>>,
    #[serde(rename = "specVanish")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub spec_vanish: Option<Box<CTOnOff>>,
    #[serde(rename = "oMath")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub o_math: Option<Box<CTOnOff>>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct EGRPrContent {
    #[serde(rename = "rStyle")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub run_style: Option<Box<CTString>>,
    #[serde(rename = "rFonts")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub fonts: Option<Box<Fonts>>,
    #[serde(rename = "b")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub bold: Option<Box<CTOnOff>>,
    #[serde(rename = "bCs")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub b_cs: Option<Box<CTOnOff>>,
    #[serde(rename = "i")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub italic: Option<Box<CTOnOff>>,
    #[serde(rename = "iCs")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub i_cs: Option<Box<CTOnOff>>,
    #[serde(rename = "caps")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub caps: Option<Box<CTOnOff>>,
    #[serde(rename = "smallCaps")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub small_caps: Option<Box<CTOnOff>>,
    #[serde(rename = "strike")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub strikethrough: Option<Box<CTOnOff>>,
    #[serde(rename = "dstrike")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub dstrike: Option<Box<CTOnOff>>,
    #[serde(rename = "outline")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub outline: Option<Box<CTOnOff>>,
    #[serde(rename = "shadow")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub shadow: Option<Box<CTOnOff>>,
    #[serde(rename = "emboss")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub emboss: Option<Box<CTOnOff>>,
    #[serde(rename = "imprint")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub imprint: Option<Box<CTOnOff>>,
    #[serde(rename = "noProof")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub no_proof: Option<Box<CTOnOff>>,
    #[serde(rename = "snapToGrid")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub snap_to_grid: Option<Box<CTOnOff>>,
    #[serde(rename = "vanish")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub vanish: Option<Box<CTOnOff>>,
    #[serde(rename = "webHidden")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub web_hidden: Option<Box<CTOnOff>>,
    #[serde(rename = "color")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub color: Option<Box<CTColor>>,
    #[serde(rename = "spacing")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub spacing: Option<Box<CTSignedTwipsMeasure>>,
    #[serde(rename = "w")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub width: Option<Box<CTTextScale>>,
    #[serde(rename = "kern")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub kern: Option<Box<CTHpsMeasure>>,
    #[serde(rename = "position")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub position: Option<Box<CTSignedHpsMeasure>>,
    #[serde(rename = "sz")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub size: Option<Box<CTHpsMeasure>>,
    #[serde(rename = "szCs")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub size_complex_script: Option<Box<CTHpsMeasure>>,
    #[serde(rename = "highlight")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub highlight: Option<Box<CTHighlight>>,
    #[serde(rename = "u")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub underline: Option<Box<CTUnderline>>,
    #[serde(rename = "effect")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub effect: Option<Box<CTTextEffect>>,
    #[serde(rename = "bdr")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub bdr: Option<Box<CTBorder>>,
    #[serde(rename = "shd")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub shading: Option<Box<CTShd>>,
    #[serde(rename = "fitText")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub fit_text: Option<Box<CTFitText>>,
    #[serde(rename = "vertAlign")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub vert_align: Option<Box<CTVerticalAlignRun>>,
    #[serde(rename = "rtl")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub rtl: Option<Box<CTOnOff>>,
    #[serde(rename = "cs")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub cs: Option<Box<CTOnOff>>,
    #[serde(rename = "em")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub em: Option<Box<CTEm>>,
    #[serde(rename = "lang")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub lang: Option<Box<CTLanguage>>,
    #[serde(rename = "eastAsianLayout")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub east_asian_layout: Option<Box<CTEastAsianLayout>>,
    #[serde(rename = "specVanish")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub spec_vanish: Option<Box<CTOnOff>>,
    #[serde(rename = "oMath")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub o_math: Option<Box<CTOnOff>>,
    #[serde(rename = "rPrChange")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub r_pr_change: Option<Box<CTRPrChange>>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct RunProperties {
    #[serde(rename = "rStyle")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub run_style: Option<Box<CTString>>,
    #[cfg(feature = "wml-styling")]
    #[serde(rename = "rFonts")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub fonts: Option<Box<Fonts>>,
    #[cfg(feature = "wml-styling")]
    #[serde(rename = "b")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub bold: Option<Box<CTOnOff>>,
    #[cfg(feature = "wml-styling")]
    #[serde(rename = "bCs")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub b_cs: Option<Box<CTOnOff>>,
    #[cfg(feature = "wml-styling")]
    #[serde(rename = "i")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub italic: Option<Box<CTOnOff>>,
    #[cfg(feature = "wml-styling")]
    #[serde(rename = "iCs")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub i_cs: Option<Box<CTOnOff>>,
    #[cfg(feature = "wml-styling")]
    #[serde(rename = "caps")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub caps: Option<Box<CTOnOff>>,
    #[cfg(feature = "wml-styling")]
    #[serde(rename = "smallCaps")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub small_caps: Option<Box<CTOnOff>>,
    #[cfg(feature = "wml-styling")]
    #[serde(rename = "strike")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub strikethrough: Option<Box<CTOnOff>>,
    #[cfg(feature = "wml-styling")]
    #[serde(rename = "dstrike")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub dstrike: Option<Box<CTOnOff>>,
    #[cfg(feature = "wml-styling")]
    #[serde(rename = "outline")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub outline: Option<Box<CTOnOff>>,
    #[cfg(feature = "wml-styling")]
    #[serde(rename = "shadow")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub shadow: Option<Box<CTOnOff>>,
    #[cfg(feature = "wml-styling")]
    #[serde(rename = "emboss")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub emboss: Option<Box<CTOnOff>>,
    #[cfg(feature = "wml-styling")]
    #[serde(rename = "imprint")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub imprint: Option<Box<CTOnOff>>,
    #[cfg(feature = "wml-styling")]
    #[serde(rename = "noProof")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub no_proof: Option<Box<CTOnOff>>,
    #[cfg(feature = "wml-styling")]
    #[serde(rename = "snapToGrid")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub snap_to_grid: Option<Box<CTOnOff>>,
    #[cfg(feature = "wml-styling")]
    #[serde(rename = "vanish")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub vanish: Option<Box<CTOnOff>>,
    #[cfg(feature = "wml-styling")]
    #[serde(rename = "webHidden")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub web_hidden: Option<Box<CTOnOff>>,
    #[cfg(feature = "wml-styling")]
    #[serde(rename = "color")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub color: Option<Box<CTColor>>,
    #[cfg(feature = "wml-styling")]
    #[serde(rename = "spacing")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub spacing: Option<Box<CTSignedTwipsMeasure>>,
    #[cfg(feature = "wml-styling")]
    #[serde(rename = "w")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub width: Option<Box<CTTextScale>>,
    #[cfg(feature = "wml-styling")]
    #[serde(rename = "kern")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub kern: Option<Box<CTHpsMeasure>>,
    #[cfg(feature = "wml-styling")]
    #[serde(rename = "position")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub position: Option<Box<CTSignedHpsMeasure>>,
    #[cfg(feature = "wml-styling")]
    #[serde(rename = "sz")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub size: Option<Box<CTHpsMeasure>>,
    #[cfg(feature = "wml-styling")]
    #[serde(rename = "szCs")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub size_complex_script: Option<Box<CTHpsMeasure>>,
    #[cfg(feature = "wml-styling")]
    #[serde(rename = "highlight")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub highlight: Option<Box<CTHighlight>>,
    #[cfg(feature = "wml-styling")]
    #[serde(rename = "u")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub underline: Option<Box<CTUnderline>>,
    #[cfg(feature = "wml-styling")]
    #[serde(rename = "effect")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub effect: Option<Box<CTTextEffect>>,
    #[cfg(feature = "wml-styling")]
    #[serde(rename = "bdr")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub bdr: Option<Box<CTBorder>>,
    #[cfg(feature = "wml-styling")]
    #[serde(rename = "shd")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub shading: Option<Box<CTShd>>,
    #[cfg(feature = "wml-styling")]
    #[serde(rename = "fitText")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub fit_text: Option<Box<CTFitText>>,
    #[cfg(feature = "wml-styling")]
    #[serde(rename = "vertAlign")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub vert_align: Option<Box<CTVerticalAlignRun>>,
    #[cfg(feature = "wml-styling")]
    #[serde(rename = "rtl")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub rtl: Option<Box<CTOnOff>>,
    #[cfg(feature = "wml-styling")]
    #[serde(rename = "cs")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub cs: Option<Box<CTOnOff>>,
    #[cfg(feature = "wml-styling")]
    #[serde(rename = "em")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub em: Option<Box<CTEm>>,
    #[cfg(feature = "wml-styling")]
    #[serde(rename = "lang")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub lang: Option<Box<CTLanguage>>,
    #[cfg(feature = "wml-styling")]
    #[serde(rename = "eastAsianLayout")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub east_asian_layout: Option<Box<CTEastAsianLayout>>,
    #[cfg(feature = "wml-styling")]
    #[serde(rename = "specVanish")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub spec_vanish: Option<Box<CTOnOff>>,
    #[cfg(feature = "wml-styling")]
    #[serde(rename = "oMath")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub o_math: Option<Box<CTOnOff>>,
    #[cfg(feature = "wml-track-changes")]
    #[serde(rename = "rPrChange")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub r_pr_change: Option<Box<CTRPrChange>>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct EGRPr {
    #[serde(rename = "rPr")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub r_pr: Option<Box<RunProperties>>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTMathCtrlIns {
    #[serde(rename = "@w:id")]
    pub id: STDecimalNumber,
    #[serde(rename = "@w:author")]
    pub author: STString,
    #[serde(rename = "@w:date")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub date: Option<STDateTime>,
    #[cfg(feature = "wml-track-changes")]
    #[serde(rename = "del")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub del: Option<Box<CTRPrChange>>,
    #[cfg(feature = "wml-styling")]
    #[serde(rename = "rPr")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub r_pr: Option<Box<RunProperties>>,
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
pub struct CTMathCtrlDel {
    #[serde(rename = "@w:id")]
    pub id: STDecimalNumber,
    #[serde(rename = "@w:author")]
    pub author: STString,
    #[serde(rename = "@w:date")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub date: Option<STDateTime>,
    #[cfg(feature = "wml-styling")]
    #[serde(rename = "rPr")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub r_pr: Option<Box<RunProperties>>,
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
pub struct CTRPrOriginal {
    #[serde(rename = "rStyle")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub run_style: Option<Box<CTString>>,
    #[serde(rename = "rFonts")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub fonts: Option<Box<Fonts>>,
    #[serde(rename = "b")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub bold: Option<Box<CTOnOff>>,
    #[serde(rename = "bCs")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub b_cs: Option<Box<CTOnOff>>,
    #[serde(rename = "i")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub italic: Option<Box<CTOnOff>>,
    #[serde(rename = "iCs")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub i_cs: Option<Box<CTOnOff>>,
    #[serde(rename = "caps")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub caps: Option<Box<CTOnOff>>,
    #[serde(rename = "smallCaps")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub small_caps: Option<Box<CTOnOff>>,
    #[serde(rename = "strike")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub strikethrough: Option<Box<CTOnOff>>,
    #[serde(rename = "dstrike")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub dstrike: Option<Box<CTOnOff>>,
    #[serde(rename = "outline")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub outline: Option<Box<CTOnOff>>,
    #[serde(rename = "shadow")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub shadow: Option<Box<CTOnOff>>,
    #[serde(rename = "emboss")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub emboss: Option<Box<CTOnOff>>,
    #[serde(rename = "imprint")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub imprint: Option<Box<CTOnOff>>,
    #[serde(rename = "noProof")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub no_proof: Option<Box<CTOnOff>>,
    #[serde(rename = "snapToGrid")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub snap_to_grid: Option<Box<CTOnOff>>,
    #[serde(rename = "vanish")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub vanish: Option<Box<CTOnOff>>,
    #[serde(rename = "webHidden")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub web_hidden: Option<Box<CTOnOff>>,
    #[serde(rename = "color")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub color: Option<Box<CTColor>>,
    #[serde(rename = "spacing")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub spacing: Option<Box<CTSignedTwipsMeasure>>,
    #[serde(rename = "w")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub width: Option<Box<CTTextScale>>,
    #[serde(rename = "kern")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub kern: Option<Box<CTHpsMeasure>>,
    #[serde(rename = "position")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub position: Option<Box<CTSignedHpsMeasure>>,
    #[serde(rename = "sz")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub size: Option<Box<CTHpsMeasure>>,
    #[serde(rename = "szCs")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub size_complex_script: Option<Box<CTHpsMeasure>>,
    #[serde(rename = "highlight")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub highlight: Option<Box<CTHighlight>>,
    #[serde(rename = "u")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub underline: Option<Box<CTUnderline>>,
    #[serde(rename = "effect")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub effect: Option<Box<CTTextEffect>>,
    #[serde(rename = "bdr")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub bdr: Option<Box<CTBorder>>,
    #[serde(rename = "shd")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub shading: Option<Box<CTShd>>,
    #[serde(rename = "fitText")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub fit_text: Option<Box<CTFitText>>,
    #[serde(rename = "vertAlign")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub vert_align: Option<Box<CTVerticalAlignRun>>,
    #[serde(rename = "rtl")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub rtl: Option<Box<CTOnOff>>,
    #[serde(rename = "cs")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub cs: Option<Box<CTOnOff>>,
    #[serde(rename = "em")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub em: Option<Box<CTEm>>,
    #[serde(rename = "lang")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub lang: Option<Box<CTLanguage>>,
    #[serde(rename = "eastAsianLayout")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub east_asian_layout: Option<Box<CTEastAsianLayout>>,
    #[serde(rename = "specVanish")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub spec_vanish: Option<Box<CTOnOff>>,
    #[serde(rename = "oMath")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub o_math: Option<Box<CTOnOff>>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct CTParaRPrOriginal {
    #[serde(rename = "ins")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub ins: Option<Box<CTTrackChange>>,
    #[serde(rename = "del")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub del: Option<Box<CTTrackChange>>,
    #[serde(rename = "moveFrom")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub move_from: Option<Box<CTTrackChange>>,
    #[serde(rename = "moveTo")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub move_to: Option<Box<CTTrackChange>>,
    #[serde(rename = "rStyle")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub run_style: Option<Box<CTString>>,
    #[serde(rename = "rFonts")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub fonts: Option<Box<Fonts>>,
    #[serde(rename = "b")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub bold: Option<Box<CTOnOff>>,
    #[serde(rename = "bCs")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub b_cs: Option<Box<CTOnOff>>,
    #[serde(rename = "i")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub italic: Option<Box<CTOnOff>>,
    #[serde(rename = "iCs")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub i_cs: Option<Box<CTOnOff>>,
    #[serde(rename = "caps")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub caps: Option<Box<CTOnOff>>,
    #[serde(rename = "smallCaps")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub small_caps: Option<Box<CTOnOff>>,
    #[serde(rename = "strike")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub strikethrough: Option<Box<CTOnOff>>,
    #[serde(rename = "dstrike")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub dstrike: Option<Box<CTOnOff>>,
    #[serde(rename = "outline")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub outline: Option<Box<CTOnOff>>,
    #[serde(rename = "shadow")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub shadow: Option<Box<CTOnOff>>,
    #[serde(rename = "emboss")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub emboss: Option<Box<CTOnOff>>,
    #[serde(rename = "imprint")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub imprint: Option<Box<CTOnOff>>,
    #[serde(rename = "noProof")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub no_proof: Option<Box<CTOnOff>>,
    #[serde(rename = "snapToGrid")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub snap_to_grid: Option<Box<CTOnOff>>,
    #[serde(rename = "vanish")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub vanish: Option<Box<CTOnOff>>,
    #[serde(rename = "webHidden")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub web_hidden: Option<Box<CTOnOff>>,
    #[serde(rename = "color")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub color: Option<Box<CTColor>>,
    #[serde(rename = "spacing")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub spacing: Option<Box<CTSignedTwipsMeasure>>,
    #[serde(rename = "w")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub width: Option<Box<CTTextScale>>,
    #[serde(rename = "kern")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub kern: Option<Box<CTHpsMeasure>>,
    #[serde(rename = "position")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub position: Option<Box<CTSignedHpsMeasure>>,
    #[serde(rename = "sz")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub size: Option<Box<CTHpsMeasure>>,
    #[serde(rename = "szCs")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub size_complex_script: Option<Box<CTHpsMeasure>>,
    #[serde(rename = "highlight")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub highlight: Option<Box<CTHighlight>>,
    #[serde(rename = "u")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub underline: Option<Box<CTUnderline>>,
    #[serde(rename = "effect")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub effect: Option<Box<CTTextEffect>>,
    #[serde(rename = "bdr")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub bdr: Option<Box<CTBorder>>,
    #[serde(rename = "shd")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub shading: Option<Box<CTShd>>,
    #[serde(rename = "fitText")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub fit_text: Option<Box<CTFitText>>,
    #[serde(rename = "vertAlign")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub vert_align: Option<Box<CTVerticalAlignRun>>,
    #[serde(rename = "rtl")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub rtl: Option<Box<CTOnOff>>,
    #[serde(rename = "cs")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub cs: Option<Box<CTOnOff>>,
    #[serde(rename = "em")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub em: Option<Box<CTEm>>,
    #[serde(rename = "lang")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub lang: Option<Box<CTLanguage>>,
    #[serde(rename = "eastAsianLayout")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub east_asian_layout: Option<Box<CTEastAsianLayout>>,
    #[serde(rename = "specVanish")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub spec_vanish: Option<Box<CTOnOff>>,
    #[serde(rename = "oMath")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub o_math: Option<Box<CTOnOff>>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct CTParaRPr {
    #[serde(rename = "ins")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub ins: Option<Box<CTTrackChange>>,
    #[serde(rename = "del")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub del: Option<Box<CTTrackChange>>,
    #[serde(rename = "moveFrom")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub move_from: Option<Box<CTTrackChange>>,
    #[serde(rename = "moveTo")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub move_to: Option<Box<CTTrackChange>>,
    #[serde(rename = "rStyle")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub run_style: Option<Box<CTString>>,
    #[serde(rename = "rFonts")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub fonts: Option<Box<Fonts>>,
    #[serde(rename = "b")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub bold: Option<Box<CTOnOff>>,
    #[serde(rename = "bCs")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub b_cs: Option<Box<CTOnOff>>,
    #[serde(rename = "i")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub italic: Option<Box<CTOnOff>>,
    #[serde(rename = "iCs")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub i_cs: Option<Box<CTOnOff>>,
    #[serde(rename = "caps")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub caps: Option<Box<CTOnOff>>,
    #[serde(rename = "smallCaps")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub small_caps: Option<Box<CTOnOff>>,
    #[serde(rename = "strike")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub strikethrough: Option<Box<CTOnOff>>,
    #[serde(rename = "dstrike")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub dstrike: Option<Box<CTOnOff>>,
    #[serde(rename = "outline")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub outline: Option<Box<CTOnOff>>,
    #[serde(rename = "shadow")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub shadow: Option<Box<CTOnOff>>,
    #[serde(rename = "emboss")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub emboss: Option<Box<CTOnOff>>,
    #[serde(rename = "imprint")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub imprint: Option<Box<CTOnOff>>,
    #[serde(rename = "noProof")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub no_proof: Option<Box<CTOnOff>>,
    #[serde(rename = "snapToGrid")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub snap_to_grid: Option<Box<CTOnOff>>,
    #[serde(rename = "vanish")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub vanish: Option<Box<CTOnOff>>,
    #[serde(rename = "webHidden")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub web_hidden: Option<Box<CTOnOff>>,
    #[serde(rename = "color")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub color: Option<Box<CTColor>>,
    #[serde(rename = "spacing")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub spacing: Option<Box<CTSignedTwipsMeasure>>,
    #[serde(rename = "w")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub width: Option<Box<CTTextScale>>,
    #[serde(rename = "kern")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub kern: Option<Box<CTHpsMeasure>>,
    #[serde(rename = "position")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub position: Option<Box<CTSignedHpsMeasure>>,
    #[serde(rename = "sz")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub size: Option<Box<CTHpsMeasure>>,
    #[serde(rename = "szCs")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub size_complex_script: Option<Box<CTHpsMeasure>>,
    #[serde(rename = "highlight")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub highlight: Option<Box<CTHighlight>>,
    #[serde(rename = "u")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub underline: Option<Box<CTUnderline>>,
    #[serde(rename = "effect")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub effect: Option<Box<CTTextEffect>>,
    #[serde(rename = "bdr")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub bdr: Option<Box<CTBorder>>,
    #[serde(rename = "shd")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub shading: Option<Box<CTShd>>,
    #[serde(rename = "fitText")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub fit_text: Option<Box<CTFitText>>,
    #[serde(rename = "vertAlign")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub vert_align: Option<Box<CTVerticalAlignRun>>,
    #[serde(rename = "rtl")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub rtl: Option<Box<CTOnOff>>,
    #[serde(rename = "cs")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub cs: Option<Box<CTOnOff>>,
    #[serde(rename = "em")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub em: Option<Box<CTEm>>,
    #[serde(rename = "lang")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub lang: Option<Box<CTLanguage>>,
    #[serde(rename = "eastAsianLayout")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub east_asian_layout: Option<Box<CTEastAsianLayout>>,
    #[serde(rename = "specVanish")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub spec_vanish: Option<Box<CTOnOff>>,
    #[serde(rename = "oMath")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub o_math: Option<Box<CTOnOff>>,
    #[cfg(feature = "wml-track-changes")]
    #[serde(rename = "rPrChange")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub r_pr_change: Option<Box<CTParaRPrChange>>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct EGParaRPrTrackChanges {
    #[serde(rename = "ins")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub ins: Option<Box<CTTrackChange>>,
    #[serde(rename = "del")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub del: Option<Box<CTTrackChange>>,
    #[serde(rename = "moveFrom")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub move_from: Option<Box<CTTrackChange>>,
    #[serde(rename = "moveTo")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub move_to: Option<Box<CTTrackChange>>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct CTAltChunk {
    #[serde(rename = "@r:id")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub id: Option<STRelationshipId>,
    #[cfg(feature = "wml-settings")]
    #[serde(rename = "altChunkPr")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub alt_chunk_pr: Option<Box<CTAltChunkPr>>,
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
pub struct CTAltChunkPr {
    #[cfg(feature = "wml-settings")]
    #[serde(rename = "matchSrc")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub match_src: Option<Box<CTOnOff>>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTRubyAlign {
    #[serde(rename = "@w:val")]
    pub value: STRubyAlign,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTRubyPr {
    #[cfg(feature = "wml-styling")]
    #[serde(rename = "rubyAlign")]
    pub ruby_align: Box<CTRubyAlign>,
    #[cfg(feature = "wml-styling")]
    #[serde(rename = "hps")]
    pub hps: Box<CTHpsMeasure>,
    #[cfg(feature = "wml-styling")]
    #[serde(rename = "hpsRaise")]
    pub hps_raise: Box<CTHpsMeasure>,
    #[cfg(feature = "wml-styling")]
    #[serde(rename = "hpsBaseText")]
    pub hps_base_text: Box<CTHpsMeasure>,
    #[cfg(feature = "wml-styling")]
    #[serde(rename = "lid")]
    pub lid: Box<CTLang>,
    #[cfg(feature = "wml-styling")]
    #[serde(rename = "dirty")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub dirty: Option<Box<CTOnOff>>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct CTRubyContent {
    #[serde(skip)]
    #[serde(default)]
    pub ruby_content: Vec<RubyContent>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTRuby {
    #[cfg(feature = "wml-styling")]
    #[serde(rename = "rubyPr")]
    pub ruby_pr: Box<CTRubyPr>,
    #[serde(rename = "rt")]
    pub rt: Box<CTRubyContent>,
    #[serde(rename = "rubyBase")]
    pub ruby_base: Box<CTRubyContent>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct CTLock {
    #[serde(rename = "@w:val")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub value: Option<STLock>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct CTSdtListItem {
    #[serde(rename = "@w:displayText")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub display_text: Option<STString>,
    #[serde(rename = "@w:value")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub value: Option<STString>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct CTSdtDateMappingType {
    #[serde(rename = "@w:val")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub value: Option<STSdtDateMappingType>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct CTCalendarType {
    #[serde(rename = "@w:val")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub value: Option<CalendarType>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct CTSdtDate {
    #[cfg(feature = "wml-settings")]
    #[serde(rename = "@w:fullDate")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub full_date: Option<STDateTime>,
    #[cfg(feature = "wml-settings")]
    #[serde(rename = "dateFormat")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub date_format: Option<Box<CTString>>,
    #[cfg(feature = "wml-settings")]
    #[serde(rename = "lid")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub lid: Option<Box<CTLang>>,
    #[cfg(feature = "wml-settings")]
    #[serde(rename = "storeMappedDataAs")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub store_mapped_data_as: Option<Box<CTSdtDateMappingType>>,
    #[cfg(feature = "wml-settings")]
    #[serde(rename = "calendar")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub calendar: Option<Box<CTCalendarType>>,
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
pub struct CTSdtComboBox {
    #[cfg(feature = "wml-settings")]
    #[serde(rename = "@w:lastValue")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub last_value: Option<STString>,
    #[cfg(feature = "wml-settings")]
    #[serde(rename = "listItem")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub list_item: Vec<CTSdtListItem>,
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
pub struct CTSdtDocPart {
    #[cfg(feature = "wml-settings")]
    #[serde(rename = "docPartGallery")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub doc_part_gallery: Option<Box<CTString>>,
    #[cfg(feature = "wml-settings")]
    #[serde(rename = "docPartCategory")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub doc_part_category: Option<Box<CTString>>,
    #[cfg(feature = "wml-settings")]
    #[serde(rename = "docPartUnique")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub doc_part_unique: Option<Box<CTOnOff>>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct CTSdtDropDownList {
    #[cfg(feature = "wml-settings")]
    #[serde(rename = "@w:lastValue")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub last_value: Option<STString>,
    #[cfg(feature = "wml-settings")]
    #[serde(rename = "listItem")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub list_item: Vec<CTSdtListItem>,
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

pub type CTPlaceholder = Box<CTString>;

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct CTSdtText {
    #[cfg(feature = "wml-settings")]
    #[serde(rename = "@w:multiLine")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub multi_line: Option<OnOff>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTDataBinding {
    #[cfg(feature = "wml-settings")]
    #[serde(rename = "@w:prefixMappings")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub prefix_mappings: Option<STString>,
    #[cfg(feature = "wml-settings")]
    #[serde(rename = "@w:xpath")]
    pub xpath: STString,
    #[cfg(feature = "wml-settings")]
    #[serde(rename = "@w:storeItemID")]
    pub store_item_i_d: STString,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct CTSdtPr {
    #[cfg(feature = "wml-settings")]
    #[serde(rename = "rPr")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub r_pr: Option<Box<RunProperties>>,
    #[cfg(feature = "wml-settings")]
    #[serde(rename = "alias")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub alias: Option<Box<CTString>>,
    #[cfg(feature = "wml-settings")]
    #[serde(rename = "tag")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub tag: Option<Box<CTString>>,
    #[cfg(feature = "wml-settings")]
    #[serde(rename = "id")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub id: Option<Box<CTDecimalNumber>>,
    #[cfg(feature = "wml-settings")]
    #[serde(rename = "lock")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub lock: Option<Box<CTLock>>,
    #[cfg(feature = "wml-settings")]
    #[serde(rename = "placeholder")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub placeholder: Option<CTPlaceholder>,
    #[cfg(feature = "wml-settings")]
    #[serde(rename = "temporary")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub temporary: Option<Box<CTOnOff>>,
    #[cfg(feature = "wml-settings")]
    #[serde(rename = "showingPlcHdr")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub showing_plc_hdr: Option<Box<CTOnOff>>,
    #[cfg(feature = "wml-settings")]
    #[serde(rename = "dataBinding")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub data_binding: Option<Box<CTDataBinding>>,
    #[cfg(feature = "wml-settings")]
    #[serde(rename = "label")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub label: Option<Box<CTDecimalNumber>>,
    #[cfg(feature = "wml-settings")]
    #[serde(rename = "tabIndex")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub tab_index: Option<Box<CTUnsignedDecimalNumber>>,
    #[cfg(feature = "wml-settings")]
    #[serde(rename = "equation")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub equation: Option<Box<CTEmpty>>,
    #[cfg(feature = "wml-settings")]
    #[serde(rename = "comboBox")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub combo_box: Option<Box<CTSdtComboBox>>,
    #[cfg(feature = "wml-settings")]
    #[serde(rename = "date")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub date: Option<Box<CTSdtDate>>,
    #[cfg(feature = "wml-settings")]
    #[serde(rename = "docPartObj")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub doc_part_obj: Option<Box<CTSdtDocPart>>,
    #[cfg(feature = "wml-settings")]
    #[serde(rename = "docPartList")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub doc_part_list: Option<Box<CTSdtDocPart>>,
    #[cfg(feature = "wml-settings")]
    #[serde(rename = "dropDownList")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub drop_down_list: Option<Box<CTSdtDropDownList>>,
    #[cfg(feature = "wml-settings")]
    #[serde(rename = "picture")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub picture: Option<Box<CTEmpty>>,
    #[cfg(feature = "wml-settings")]
    #[serde(rename = "richText")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub rich_text: Option<Box<CTEmpty>>,
    #[cfg(feature = "wml-settings")]
    #[serde(rename = "text")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub text: Option<Box<CTSdtText>>,
    #[cfg(feature = "wml-settings")]
    #[serde(rename = "citation")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub citation: Option<Box<CTEmpty>>,
    #[cfg(feature = "wml-settings")]
    #[serde(rename = "group")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub group: Option<Box<CTEmpty>>,
    #[cfg(feature = "wml-settings")]
    #[serde(rename = "bibliography")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub bibliography: Option<Box<CTEmpty>>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct CTSdtEndPr {
    #[cfg(feature = "wml-styling")]
    #[serde(rename = "rPr")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub r_pr: Option<Box<RunProperties>>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct CTDirContentRun {
    #[serde(rename = "@w:val")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub value: Option<STDirection>,
    #[serde(skip)]
    #[serde(default)]
    pub paragraph_content: Vec<ParagraphContent>,
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
pub struct CTBdoContentRun {
    #[serde(rename = "@w:val")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub value: Option<STDirection>,
    #[serde(skip)]
    #[serde(default)]
    pub paragraph_content: Vec<ParagraphContent>,
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
pub struct CTSdtContentRun {
    #[serde(skip)]
    #[serde(default)]
    pub paragraph_content: Vec<ParagraphContent>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct CTSdtContentBlock {
    #[serde(skip)]
    #[serde(default)]
    pub block_content: Vec<BlockContentChoice>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct CTSdtContentRow {
    #[serde(skip)]
    #[serde(default)]
    pub rows: Vec<RowContent>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct CTSdtContentCell {
    #[serde(skip)]
    #[serde(default)]
    pub cells: Vec<CellContent>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct CTSdtBlock {
    #[cfg(feature = "wml-settings")]
    #[serde(rename = "sdtPr")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub sdt_pr: Option<Box<CTSdtPr>>,
    #[cfg(feature = "wml-settings")]
    #[serde(rename = "sdtEndPr")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub sdt_end_pr: Option<Box<CTSdtEndPr>>,
    #[serde(rename = "sdtContent")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub sdt_content: Option<Box<CTSdtContentBlock>>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct CTSdtRun {
    #[cfg(feature = "wml-settings")]
    #[serde(rename = "sdtPr")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub sdt_pr: Option<Box<CTSdtPr>>,
    #[cfg(feature = "wml-settings")]
    #[serde(rename = "sdtEndPr")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub sdt_end_pr: Option<Box<CTSdtEndPr>>,
    #[serde(rename = "sdtContent")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub sdt_content: Option<Box<CTSdtContentRun>>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct CTSdtCell {
    #[cfg(feature = "wml-settings")]
    #[serde(rename = "sdtPr")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub sdt_pr: Option<Box<CTSdtPr>>,
    #[cfg(feature = "wml-settings")]
    #[serde(rename = "sdtEndPr")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub sdt_end_pr: Option<Box<CTSdtEndPr>>,
    #[serde(rename = "sdtContent")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub sdt_content: Option<Box<CTSdtContentCell>>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct CTSdtRow {
    #[cfg(feature = "wml-settings")]
    #[serde(rename = "sdtPr")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub sdt_pr: Option<Box<CTSdtPr>>,
    #[cfg(feature = "wml-settings")]
    #[serde(rename = "sdtEndPr")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub sdt_end_pr: Option<Box<CTSdtEndPr>>,
    #[serde(rename = "sdtContent")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub sdt_content: Option<Box<CTSdtContentRow>>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTAttr {
    #[serde(rename = "@w:uri")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub uri: Option<STString>,
    #[serde(rename = "@w:name")]
    pub name: STString,
    #[serde(rename = "@w:val")]
    pub value: STString,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTCustomXmlRun {
    #[cfg(feature = "wml-settings")]
    #[serde(rename = "@w:uri")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub uri: Option<STString>,
    #[cfg(feature = "wml-settings")]
    #[serde(rename = "@w:element")]
    pub element: STXmlName,
    #[cfg(feature = "wml-settings")]
    #[serde(rename = "customXmlPr")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub custom_xml_pr: Option<Box<CTCustomXmlPr>>,
    #[serde(skip)]
    #[serde(default)]
    pub paragraph_content: Vec<ParagraphContent>,
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
pub struct CTSmartTagRun {
    #[cfg(feature = "wml-settings")]
    #[serde(rename = "@w:uri")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub uri: Option<STString>,
    #[cfg(feature = "wml-settings")]
    #[serde(rename = "@w:element")]
    pub element: STXmlName,
    #[cfg(feature = "wml-settings")]
    #[serde(rename = "smartTagPr")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub smart_tag_pr: Option<Box<CTSmartTagPr>>,
    #[serde(skip)]
    #[serde(default)]
    pub paragraph_content: Vec<ParagraphContent>,
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
pub struct CTCustomXmlBlock {
    #[cfg(feature = "wml-settings")]
    #[serde(rename = "@w:uri")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub uri: Option<STString>,
    #[cfg(feature = "wml-settings")]
    #[serde(rename = "@w:element")]
    pub element: STXmlName,
    #[cfg(feature = "wml-settings")]
    #[serde(rename = "customXmlPr")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub custom_xml_pr: Option<Box<CTCustomXmlPr>>,
    #[serde(skip)]
    #[serde(default)]
    pub block_content: Vec<BlockContentChoice>,
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
pub struct CTCustomXmlPr {
    #[cfg(feature = "wml-settings")]
    #[serde(rename = "placeholder")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub placeholder: Option<Box<CTString>>,
    #[cfg(feature = "wml-settings")]
    #[serde(rename = "attr")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub attr: Vec<CTAttr>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTCustomXmlRow {
    #[cfg(feature = "wml-settings")]
    #[serde(rename = "@w:uri")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub uri: Option<STString>,
    #[cfg(feature = "wml-settings")]
    #[serde(rename = "@w:element")]
    pub element: STXmlName,
    #[cfg(feature = "wml-settings")]
    #[serde(rename = "customXmlPr")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub custom_xml_pr: Option<Box<CTCustomXmlPr>>,
    #[serde(skip)]
    #[serde(default)]
    pub rows: Vec<RowContent>,
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
pub struct CTCustomXmlCell {
    #[cfg(feature = "wml-settings")]
    #[serde(rename = "@w:uri")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub uri: Option<STString>,
    #[cfg(feature = "wml-settings")]
    #[serde(rename = "@w:element")]
    pub element: STXmlName,
    #[cfg(feature = "wml-settings")]
    #[serde(rename = "customXmlPr")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub custom_xml_pr: Option<Box<CTCustomXmlPr>>,
    #[serde(skip)]
    #[serde(default)]
    pub cells: Vec<CellContent>,
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
pub struct CTSmartTagPr {
    #[cfg(feature = "wml-settings")]
    #[serde(rename = "attr")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub attr: Vec<CTAttr>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct Paragraph {
    #[cfg(feature = "wml-track-changes")]
    #[serde(rename = "@w:rsidRPr")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub rsid_r_pr: Option<STLongHexNumber>,
    #[cfg(feature = "wml-track-changes")]
    #[serde(rename = "@w:rsidR")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub rsid_r: Option<STLongHexNumber>,
    #[cfg(feature = "wml-track-changes")]
    #[serde(rename = "@w:rsidDel")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub rsid_del: Option<STLongHexNumber>,
    #[cfg(feature = "wml-track-changes")]
    #[serde(rename = "@w:rsidP")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub rsid_p: Option<STLongHexNumber>,
    #[cfg(feature = "wml-track-changes")]
    #[serde(rename = "@w:rsidRDefault")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub rsid_r_default: Option<STLongHexNumber>,
    #[cfg(feature = "wml-styling")]
    #[serde(rename = "pPr")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub p_pr: Option<Box<ParagraphProperties>>,
    #[serde(skip)]
    #[serde(default)]
    pub paragraph_content: Vec<ParagraphContent>,
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
pub struct CTHeight {
    #[serde(rename = "@w:val")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub value: Option<STTwipsMeasure>,
    #[cfg(feature = "wml-tables")]
    #[serde(rename = "@w:hRule")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub h_rule: Option<STHeightRule>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct CTTblWidth {
    #[serde(rename = "@w:w")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub width: Option<STMeasurementOrPercent>,
    #[serde(rename = "@w:type")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub r#type: Option<STTblWidth>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct TableGridColumn {
    #[serde(rename = "@w:w")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub width: Option<STTwipsMeasure>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct CTTblGridBase {
    #[serde(rename = "gridCol")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub grid_col: Vec<TableGridColumn>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct TableGrid {
    #[serde(rename = "gridCol")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub grid_col: Vec<TableGridColumn>,
    #[cfg(feature = "wml-track-changes")]
    #[serde(rename = "tblGridChange")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub tbl_grid_change: Option<Box<CTTblGridChange>>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct CTTcBorders {
    #[cfg(feature = "wml-tables")]
    #[serde(rename = "top")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub top: Option<Box<CTBorder>>,
    #[cfg(feature = "wml-tables")]
    #[serde(rename = "start")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub start: Option<Box<CTBorder>>,
    #[cfg(feature = "wml-tables")]
    #[serde(rename = "left")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub left: Option<Box<CTBorder>>,
    #[cfg(feature = "wml-tables")]
    #[serde(rename = "bottom")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub bottom: Option<Box<CTBorder>>,
    #[cfg(feature = "wml-tables")]
    #[serde(rename = "end")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub end: Option<Box<CTBorder>>,
    #[cfg(feature = "wml-tables")]
    #[serde(rename = "right")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub right: Option<Box<CTBorder>>,
    #[cfg(feature = "wml-tables")]
    #[serde(rename = "insideH")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub inside_h: Option<Box<CTBorder>>,
    #[cfg(feature = "wml-tables")]
    #[serde(rename = "insideV")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub inside_v: Option<Box<CTBorder>>,
    #[cfg(feature = "wml-tables")]
    #[serde(rename = "tl2br")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub tl2br: Option<Box<CTBorder>>,
    #[cfg(feature = "wml-tables")]
    #[serde(rename = "tr2bl")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub tr2bl: Option<Box<CTBorder>>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct CTTcMar {
    #[cfg(feature = "wml-tables")]
    #[serde(rename = "top")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub top: Option<Box<CTTblWidth>>,
    #[cfg(feature = "wml-tables")]
    #[serde(rename = "start")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub start: Option<Box<CTTblWidth>>,
    #[cfg(feature = "wml-tables")]
    #[serde(rename = "left")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub left: Option<Box<CTTblWidth>>,
    #[cfg(feature = "wml-tables")]
    #[serde(rename = "bottom")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub bottom: Option<Box<CTTblWidth>>,
    #[cfg(feature = "wml-tables")]
    #[serde(rename = "end")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub end: Option<Box<CTTblWidth>>,
    #[cfg(feature = "wml-tables")]
    #[serde(rename = "right")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub right: Option<Box<CTTblWidth>>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct CTVMerge {
    #[serde(rename = "@w:val")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub value: Option<STMerge>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct CTHMerge {
    #[serde(rename = "@w:val")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub value: Option<STMerge>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct CTTcPrBase {
    #[cfg(feature = "wml-styling")]
    #[serde(rename = "cnfStyle")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub cnf_style: Option<Box<CTCnf>>,
    #[cfg(feature = "wml-tables")]
    #[serde(rename = "tcW")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub tc_w: Option<Box<CTTblWidth>>,
    #[cfg(feature = "wml-tables")]
    #[serde(rename = "gridSpan")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub grid_span: Option<Box<CTDecimalNumber>>,
    #[cfg(feature = "wml-tables")]
    #[serde(rename = "hMerge")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub horizontal_merge: Option<Box<CTHMerge>>,
    #[cfg(feature = "wml-tables")]
    #[serde(rename = "vMerge")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub vertical_merge: Option<Box<CTVMerge>>,
    #[cfg(feature = "wml-tables")]
    #[serde(rename = "tcBorders")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub tc_borders: Option<Box<CTTcBorders>>,
    #[cfg(feature = "wml-tables")]
    #[serde(rename = "shd")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub shading: Option<Box<CTShd>>,
    #[cfg(feature = "wml-tables")]
    #[serde(rename = "noWrap")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub no_wrap: Option<Box<CTOnOff>>,
    #[cfg(feature = "wml-tables")]
    #[serde(rename = "tcMar")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub tc_mar: Option<Box<CTTcMar>>,
    #[cfg(feature = "wml-tables")]
    #[serde(rename = "textDirection")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub text_direction: Option<Box<CTTextDirection>>,
    #[cfg(feature = "wml-tables")]
    #[serde(rename = "tcFitText")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub tc_fit_text: Option<Box<CTOnOff>>,
    #[cfg(feature = "wml-tables")]
    #[serde(rename = "vAlign")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub v_align: Option<Box<CTVerticalJc>>,
    #[cfg(feature = "wml-tables")]
    #[serde(rename = "hideMark")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub hide_mark: Option<Box<CTOnOff>>,
    #[cfg(feature = "wml-tables")]
    #[serde(rename = "headers")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub headers: Option<Box<CTHeaders>>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct TableCellProperties {
    #[cfg(feature = "wml-styling")]
    #[serde(rename = "cnfStyle")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub cnf_style: Option<Box<CTCnf>>,
    #[cfg(feature = "wml-tables")]
    #[serde(rename = "tcW")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub tc_w: Option<Box<CTTblWidth>>,
    #[cfg(feature = "wml-tables")]
    #[serde(rename = "gridSpan")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub grid_span: Option<Box<CTDecimalNumber>>,
    #[cfg(feature = "wml-tables")]
    #[serde(rename = "hMerge")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub horizontal_merge: Option<Box<CTHMerge>>,
    #[cfg(feature = "wml-tables")]
    #[serde(rename = "vMerge")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub vertical_merge: Option<Box<CTVMerge>>,
    #[cfg(feature = "wml-tables")]
    #[serde(rename = "tcBorders")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub tc_borders: Option<Box<CTTcBorders>>,
    #[cfg(feature = "wml-tables")]
    #[serde(rename = "shd")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub shading: Option<Box<CTShd>>,
    #[cfg(feature = "wml-tables")]
    #[serde(rename = "noWrap")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub no_wrap: Option<Box<CTOnOff>>,
    #[cfg(feature = "wml-tables")]
    #[serde(rename = "tcMar")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub tc_mar: Option<Box<CTTcMar>>,
    #[cfg(feature = "wml-tables")]
    #[serde(rename = "textDirection")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub text_direction: Option<Box<CTTextDirection>>,
    #[cfg(feature = "wml-tables")]
    #[serde(rename = "tcFitText")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub tc_fit_text: Option<Box<CTOnOff>>,
    #[cfg(feature = "wml-tables")]
    #[serde(rename = "vAlign")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub v_align: Option<Box<CTVerticalJc>>,
    #[cfg(feature = "wml-tables")]
    #[serde(rename = "hideMark")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub hide_mark: Option<Box<CTOnOff>>,
    #[cfg(feature = "wml-tables")]
    #[serde(rename = "headers")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub headers: Option<Box<CTHeaders>>,
    #[serde(skip)]
    #[serde(default)]
    pub cell_markup: Option<Box<CellMarkup>>,
    #[cfg(feature = "wml-track-changes")]
    #[serde(rename = "tcPrChange")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub tc_pr_change: Option<Box<CTTcPrChange>>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct CTTcPrInner {
    #[serde(rename = "cnfStyle")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub cnf_style: Option<Box<CTCnf>>,
    #[serde(rename = "tcW")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub tc_w: Option<Box<CTTblWidth>>,
    #[serde(rename = "gridSpan")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub grid_span: Option<Box<CTDecimalNumber>>,
    #[serde(rename = "hMerge")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub horizontal_merge: Option<Box<CTHMerge>>,
    #[serde(rename = "vMerge")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub vertical_merge: Option<Box<CTVMerge>>,
    #[serde(rename = "tcBorders")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub tc_borders: Option<Box<CTTcBorders>>,
    #[serde(rename = "shd")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub shading: Option<Box<CTShd>>,
    #[serde(rename = "noWrap")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub no_wrap: Option<Box<CTOnOff>>,
    #[serde(rename = "tcMar")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub tc_mar: Option<Box<CTTcMar>>,
    #[serde(rename = "textDirection")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub text_direction: Option<Box<CTTextDirection>>,
    #[serde(rename = "tcFitText")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub tc_fit_text: Option<Box<CTOnOff>>,
    #[serde(rename = "vAlign")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub v_align: Option<Box<CTVerticalJc>>,
    #[serde(rename = "hideMark")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub hide_mark: Option<Box<CTOnOff>>,
    #[serde(rename = "headers")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub headers: Option<Box<CTHeaders>>,
    #[serde(skip)]
    #[serde(default)]
    pub cell_markup: Option<Box<CellMarkup>>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct TableCell {
    #[cfg(feature = "wml-tables")]
    #[serde(rename = "@w:id")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub id: Option<STString>,
    #[cfg(feature = "wml-tables")]
    #[serde(rename = "tcPr")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub cell_properties: Option<Box<TableCellProperties>>,
    #[serde(skip)]
    #[serde(default)]
    pub block_content: Vec<BlockContent>,
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
pub struct CTCnf {
    #[cfg(feature = "wml-styling")]
    #[serde(rename = "@w:val")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub value: Option<STCnf>,
    #[cfg(feature = "wml-styling")]
    #[serde(rename = "@w:firstRow")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub first_row: Option<OnOff>,
    #[cfg(feature = "wml-styling")]
    #[serde(rename = "@w:lastRow")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub last_row: Option<OnOff>,
    #[cfg(feature = "wml-styling")]
    #[serde(rename = "@w:firstColumn")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub first_column: Option<OnOff>,
    #[cfg(feature = "wml-styling")]
    #[serde(rename = "@w:lastColumn")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub last_column: Option<OnOff>,
    #[cfg(feature = "wml-styling")]
    #[serde(rename = "@w:oddVBand")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub odd_v_band: Option<OnOff>,
    #[cfg(feature = "wml-styling")]
    #[serde(rename = "@w:evenVBand")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub even_v_band: Option<OnOff>,
    #[cfg(feature = "wml-styling")]
    #[serde(rename = "@w:oddHBand")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub odd_h_band: Option<OnOff>,
    #[cfg(feature = "wml-styling")]
    #[serde(rename = "@w:evenHBand")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub even_h_band: Option<OnOff>,
    #[cfg(feature = "wml-styling")]
    #[serde(rename = "@w:firstRowFirstColumn")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub first_row_first_column: Option<OnOff>,
    #[cfg(feature = "wml-styling")]
    #[serde(rename = "@w:firstRowLastColumn")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub first_row_last_column: Option<OnOff>,
    #[cfg(feature = "wml-styling")]
    #[serde(rename = "@w:lastRowFirstColumn")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub last_row_first_column: Option<OnOff>,
    #[cfg(feature = "wml-styling")]
    #[serde(rename = "@w:lastRowLastColumn")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub last_row_last_column: Option<OnOff>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct CTHeaders {
    #[serde(rename = "header")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub header: Vec<CTString>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct CTTrPrBase {
    #[cfg(feature = "wml-styling")]
    #[serde(rename = "cnfStyle")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub cnf_style: Option<Box<CTCnf>>,
    #[cfg(feature = "wml-settings")]
    #[serde(rename = "divId")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub div_id: Option<Box<CTDecimalNumber>>,
    #[cfg(feature = "wml-tables")]
    #[serde(rename = "gridBefore")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub grid_before: Option<Box<CTDecimalNumber>>,
    #[cfg(feature = "wml-tables")]
    #[serde(rename = "gridAfter")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub grid_after: Option<Box<CTDecimalNumber>>,
    #[cfg(feature = "wml-tables")]
    #[serde(rename = "wBefore")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub w_before: Option<Box<CTTblWidth>>,
    #[cfg(feature = "wml-tables")]
    #[serde(rename = "wAfter")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub w_after: Option<Box<CTTblWidth>>,
    #[cfg(feature = "wml-tables")]
    #[serde(rename = "cantSplit")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub cant_split: Option<Box<CTOnOff>>,
    #[cfg(feature = "wml-tables")]
    #[serde(rename = "trHeight")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub tr_height: Option<Box<CTHeight>>,
    #[cfg(feature = "wml-tables")]
    #[serde(rename = "tblHeader")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub tbl_header: Option<Box<CTOnOff>>,
    #[cfg(feature = "wml-tables")]
    #[serde(rename = "tblCellSpacing")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub tbl_cell_spacing: Option<Box<CTTblWidth>>,
    #[cfg(feature = "wml-tables")]
    #[serde(rename = "jc")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub justification: Option<Box<CTJcTable>>,
    #[cfg(feature = "wml-tables")]
    #[serde(rename = "hidden")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub hidden: Option<Box<CTOnOff>>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct TableRowProperties {
    #[cfg(feature = "wml-styling")]
    #[serde(rename = "cnfStyle")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub cnf_style: Option<Box<CTCnf>>,
    #[cfg(feature = "wml-settings")]
    #[serde(rename = "divId")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub div_id: Option<Box<CTDecimalNumber>>,
    #[cfg(feature = "wml-tables")]
    #[serde(rename = "gridBefore")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub grid_before: Option<Box<CTDecimalNumber>>,
    #[cfg(feature = "wml-tables")]
    #[serde(rename = "gridAfter")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub grid_after: Option<Box<CTDecimalNumber>>,
    #[cfg(feature = "wml-tables")]
    #[serde(rename = "wBefore")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub w_before: Option<Box<CTTblWidth>>,
    #[cfg(feature = "wml-tables")]
    #[serde(rename = "wAfter")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub w_after: Option<Box<CTTblWidth>>,
    #[cfg(feature = "wml-tables")]
    #[serde(rename = "cantSplit")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub cant_split: Option<Box<CTOnOff>>,
    #[cfg(feature = "wml-tables")]
    #[serde(rename = "trHeight")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub tr_height: Option<Box<CTHeight>>,
    #[cfg(feature = "wml-tables")]
    #[serde(rename = "tblHeader")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub tbl_header: Option<Box<CTOnOff>>,
    #[cfg(feature = "wml-tables")]
    #[serde(rename = "tblCellSpacing")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub tbl_cell_spacing: Option<Box<CTTblWidth>>,
    #[cfg(feature = "wml-tables")]
    #[serde(rename = "jc")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub justification: Option<Box<CTJcTable>>,
    #[cfg(feature = "wml-tables")]
    #[serde(rename = "hidden")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub hidden: Option<Box<CTOnOff>>,
    #[cfg(feature = "wml-track-changes")]
    #[serde(rename = "ins")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub ins: Option<Box<CTTrackChange>>,
    #[cfg(feature = "wml-track-changes")]
    #[serde(rename = "del")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub del: Option<Box<CTTrackChange>>,
    #[cfg(feature = "wml-track-changes")]
    #[serde(rename = "trPrChange")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub tr_pr_change: Option<Box<CTTrPrChange>>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct CTRow {
    #[cfg(feature = "wml-track-changes")]
    #[serde(rename = "@w:rsidRPr")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub rsid_r_pr: Option<STLongHexNumber>,
    #[cfg(feature = "wml-track-changes")]
    #[serde(rename = "@w:rsidR")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub rsid_r: Option<STLongHexNumber>,
    #[cfg(feature = "wml-track-changes")]
    #[serde(rename = "@w:rsidDel")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub rsid_del: Option<STLongHexNumber>,
    #[cfg(feature = "wml-track-changes")]
    #[serde(rename = "@w:rsidTr")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub rsid_tr: Option<STLongHexNumber>,
    #[cfg(feature = "wml-tables")]
    #[serde(rename = "tblPrEx")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub tbl_pr_ex: Option<Box<CTTblPrEx>>,
    #[cfg(feature = "wml-tables")]
    #[serde(rename = "trPr")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub row_properties: Option<Box<TableRowProperties>>,
    #[serde(skip)]
    #[serde(default)]
    pub cells: Vec<CellContent>,
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
pub struct CTTblLayoutType {
    #[serde(rename = "@w:type")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub r#type: Option<STTblLayoutType>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTTblOverlap {
    #[serde(rename = "@w:val")]
    pub value: STTblOverlap,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct CTTblPPr {
    #[cfg(feature = "wml-tables")]
    #[serde(rename = "@w:leftFromText")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub left_from_text: Option<STTwipsMeasure>,
    #[cfg(feature = "wml-tables")]
    #[serde(rename = "@w:rightFromText")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub right_from_text: Option<STTwipsMeasure>,
    #[cfg(feature = "wml-tables")]
    #[serde(rename = "@w:topFromText")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub top_from_text: Option<STTwipsMeasure>,
    #[cfg(feature = "wml-tables")]
    #[serde(rename = "@w:bottomFromText")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub bottom_from_text: Option<STTwipsMeasure>,
    #[cfg(feature = "wml-tables")]
    #[serde(rename = "@w:vertAnchor")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub vert_anchor: Option<STVAnchor>,
    #[cfg(feature = "wml-tables")]
    #[serde(rename = "@w:horzAnchor")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub horz_anchor: Option<STHAnchor>,
    #[cfg(feature = "wml-tables")]
    #[serde(rename = "@w:tblpXSpec")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub tblp_x_spec: Option<STXAlign>,
    #[cfg(feature = "wml-tables")]
    #[serde(rename = "@w:tblpX")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub tblp_x: Option<STSignedTwipsMeasure>,
    #[cfg(feature = "wml-tables")]
    #[serde(rename = "@w:tblpYSpec")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub tblp_y_spec: Option<STYAlign>,
    #[cfg(feature = "wml-tables")]
    #[serde(rename = "@w:tblpY")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub tblp_y: Option<STSignedTwipsMeasure>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct CTTblCellMar {
    #[cfg(feature = "wml-tables")]
    #[serde(rename = "top")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub top: Option<Box<CTTblWidth>>,
    #[cfg(feature = "wml-tables")]
    #[serde(rename = "start")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub start: Option<Box<CTTblWidth>>,
    #[cfg(feature = "wml-tables")]
    #[serde(rename = "left")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub left: Option<Box<CTTblWidth>>,
    #[cfg(feature = "wml-tables")]
    #[serde(rename = "bottom")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub bottom: Option<Box<CTTblWidth>>,
    #[cfg(feature = "wml-tables")]
    #[serde(rename = "end")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub end: Option<Box<CTTblWidth>>,
    #[cfg(feature = "wml-tables")]
    #[serde(rename = "right")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub right: Option<Box<CTTblWidth>>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct CTTblBorders {
    #[cfg(feature = "wml-tables")]
    #[serde(rename = "top")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub top: Option<Box<CTBorder>>,
    #[cfg(feature = "wml-tables")]
    #[serde(rename = "start")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub start: Option<Box<CTBorder>>,
    #[cfg(feature = "wml-tables")]
    #[serde(rename = "left")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub left: Option<Box<CTBorder>>,
    #[cfg(feature = "wml-tables")]
    #[serde(rename = "bottom")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub bottom: Option<Box<CTBorder>>,
    #[cfg(feature = "wml-tables")]
    #[serde(rename = "end")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub end: Option<Box<CTBorder>>,
    #[cfg(feature = "wml-tables")]
    #[serde(rename = "right")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub right: Option<Box<CTBorder>>,
    #[cfg(feature = "wml-tables")]
    #[serde(rename = "insideH")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub inside_h: Option<Box<CTBorder>>,
    #[cfg(feature = "wml-tables")]
    #[serde(rename = "insideV")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub inside_v: Option<Box<CTBorder>>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct CTTblPrBase {
    #[cfg(feature = "wml-styling")]
    #[serde(rename = "tblStyle")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub tbl_style: Option<Box<CTString>>,
    #[cfg(feature = "wml-tables")]
    #[serde(rename = "tblpPr")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub tblp_pr: Option<Box<CTTblPPr>>,
    #[cfg(feature = "wml-tables")]
    #[serde(rename = "tblOverlap")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub tbl_overlap: Option<Box<CTTblOverlap>>,
    #[cfg(feature = "wml-tables")]
    #[serde(rename = "bidiVisual")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub bidi_visual: Option<Box<CTOnOff>>,
    #[cfg(feature = "wml-styling")]
    #[serde(rename = "tblStyleRowBandSize")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub tbl_style_row_band_size: Option<Box<CTDecimalNumber>>,
    #[cfg(feature = "wml-styling")]
    #[serde(rename = "tblStyleColBandSize")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub tbl_style_col_band_size: Option<Box<CTDecimalNumber>>,
    #[cfg(feature = "wml-tables")]
    #[serde(rename = "tblW")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub tbl_w: Option<Box<CTTblWidth>>,
    #[cfg(feature = "wml-tables")]
    #[serde(rename = "jc")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub justification: Option<Box<CTJcTable>>,
    #[cfg(feature = "wml-tables")]
    #[serde(rename = "tblCellSpacing")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub tbl_cell_spacing: Option<Box<CTTblWidth>>,
    #[cfg(feature = "wml-tables")]
    #[serde(rename = "tblInd")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub tbl_ind: Option<Box<CTTblWidth>>,
    #[cfg(feature = "wml-tables")]
    #[serde(rename = "tblBorders")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub tbl_borders: Option<Box<CTTblBorders>>,
    #[cfg(feature = "wml-tables")]
    #[serde(rename = "shd")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub shading: Option<Box<CTShd>>,
    #[cfg(feature = "wml-tables")]
    #[serde(rename = "tblLayout")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub tbl_layout: Option<Box<CTTblLayoutType>>,
    #[cfg(feature = "wml-tables")]
    #[serde(rename = "tblCellMar")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub tbl_cell_mar: Option<Box<CTTblCellMar>>,
    #[cfg(feature = "wml-tables")]
    #[serde(rename = "tblLook")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub tbl_look: Option<Box<CTTblLook>>,
    #[cfg(feature = "wml-tables")]
    #[serde(rename = "tblCaption")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub tbl_caption: Option<Box<CTString>>,
    #[cfg(feature = "wml-tables")]
    #[serde(rename = "tblDescription")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub tbl_description: Option<Box<CTString>>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct TableProperties {
    #[cfg(feature = "wml-styling")]
    #[serde(rename = "tblStyle")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub tbl_style: Option<Box<CTString>>,
    #[cfg(feature = "wml-tables")]
    #[serde(rename = "tblpPr")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub tblp_pr: Option<Box<CTTblPPr>>,
    #[cfg(feature = "wml-tables")]
    #[serde(rename = "tblOverlap")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub tbl_overlap: Option<Box<CTTblOverlap>>,
    #[cfg(feature = "wml-tables")]
    #[serde(rename = "bidiVisual")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub bidi_visual: Option<Box<CTOnOff>>,
    #[cfg(feature = "wml-styling")]
    #[serde(rename = "tblStyleRowBandSize")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub tbl_style_row_band_size: Option<Box<CTDecimalNumber>>,
    #[cfg(feature = "wml-styling")]
    #[serde(rename = "tblStyleColBandSize")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub tbl_style_col_band_size: Option<Box<CTDecimalNumber>>,
    #[cfg(feature = "wml-tables")]
    #[serde(rename = "tblW")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub tbl_w: Option<Box<CTTblWidth>>,
    #[cfg(feature = "wml-tables")]
    #[serde(rename = "jc")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub justification: Option<Box<CTJcTable>>,
    #[cfg(feature = "wml-tables")]
    #[serde(rename = "tblCellSpacing")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub tbl_cell_spacing: Option<Box<CTTblWidth>>,
    #[cfg(feature = "wml-tables")]
    #[serde(rename = "tblInd")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub tbl_ind: Option<Box<CTTblWidth>>,
    #[cfg(feature = "wml-tables")]
    #[serde(rename = "tblBorders")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub tbl_borders: Option<Box<CTTblBorders>>,
    #[cfg(feature = "wml-tables")]
    #[serde(rename = "shd")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub shading: Option<Box<CTShd>>,
    #[cfg(feature = "wml-tables")]
    #[serde(rename = "tblLayout")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub tbl_layout: Option<Box<CTTblLayoutType>>,
    #[cfg(feature = "wml-tables")]
    #[serde(rename = "tblCellMar")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub tbl_cell_mar: Option<Box<CTTblCellMar>>,
    #[cfg(feature = "wml-tables")]
    #[serde(rename = "tblLook")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub tbl_look: Option<Box<CTTblLook>>,
    #[cfg(feature = "wml-tables")]
    #[serde(rename = "tblCaption")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub tbl_caption: Option<Box<CTString>>,
    #[cfg(feature = "wml-tables")]
    #[serde(rename = "tblDescription")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub tbl_description: Option<Box<CTString>>,
    #[cfg(feature = "wml-track-changes")]
    #[serde(rename = "tblPrChange")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub tbl_pr_change: Option<Box<CTTblPrChange>>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct CTTblPrExBase {
    #[cfg(feature = "wml-tables")]
    #[serde(rename = "tblW")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub tbl_w: Option<Box<CTTblWidth>>,
    #[cfg(feature = "wml-tables")]
    #[serde(rename = "jc")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub justification: Option<Box<CTJcTable>>,
    #[cfg(feature = "wml-tables")]
    #[serde(rename = "tblCellSpacing")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub tbl_cell_spacing: Option<Box<CTTblWidth>>,
    #[cfg(feature = "wml-tables")]
    #[serde(rename = "tblInd")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub tbl_ind: Option<Box<CTTblWidth>>,
    #[cfg(feature = "wml-tables")]
    #[serde(rename = "tblBorders")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub tbl_borders: Option<Box<CTTblBorders>>,
    #[cfg(feature = "wml-tables")]
    #[serde(rename = "shd")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub shading: Option<Box<CTShd>>,
    #[cfg(feature = "wml-tables")]
    #[serde(rename = "tblLayout")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub tbl_layout: Option<Box<CTTblLayoutType>>,
    #[cfg(feature = "wml-tables")]
    #[serde(rename = "tblCellMar")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub tbl_cell_mar: Option<Box<CTTblCellMar>>,
    #[cfg(feature = "wml-tables")]
    #[serde(rename = "tblLook")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub tbl_look: Option<Box<CTTblLook>>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct CTTblPrEx {
    #[serde(rename = "tblW")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub tbl_w: Option<Box<CTTblWidth>>,
    #[serde(rename = "jc")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub justification: Option<Box<CTJcTable>>,
    #[serde(rename = "tblCellSpacing")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub tbl_cell_spacing: Option<Box<CTTblWidth>>,
    #[serde(rename = "tblInd")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub tbl_ind: Option<Box<CTTblWidth>>,
    #[serde(rename = "tblBorders")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub tbl_borders: Option<Box<CTTblBorders>>,
    #[serde(rename = "shd")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub shading: Option<Box<CTShd>>,
    #[serde(rename = "tblLayout")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub tbl_layout: Option<Box<CTTblLayoutType>>,
    #[serde(rename = "tblCellMar")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub tbl_cell_mar: Option<Box<CTTblCellMar>>,
    #[serde(rename = "tblLook")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub tbl_look: Option<Box<CTTblLook>>,
    #[cfg(feature = "wml-track-changes")]
    #[serde(rename = "tblPrExChange")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub tbl_pr_ex_change: Option<Box<CTTblPrExChange>>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct Table {
    #[serde(skip)]
    #[serde(default)]
    pub range_markup: Vec<RangeMarkup>,
    #[serde(rename = "tblPr")]
    pub table_properties: Box<TableProperties>,
    #[serde(rename = "tblGrid")]
    pub tbl_grid: Box<TableGrid>,
    #[serde(skip)]
    #[serde(default)]
    pub rows: Vec<RowContent>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct CTTblLook {
    #[cfg(feature = "wml-tables")]
    #[serde(rename = "@w:firstRow")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub first_row: Option<OnOff>,
    #[cfg(feature = "wml-tables")]
    #[serde(rename = "@w:lastRow")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub last_row: Option<OnOff>,
    #[cfg(feature = "wml-tables")]
    #[serde(rename = "@w:firstColumn")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub first_column: Option<OnOff>,
    #[cfg(feature = "wml-tables")]
    #[serde(rename = "@w:lastColumn")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub last_column: Option<OnOff>,
    #[cfg(feature = "wml-tables")]
    #[serde(rename = "@w:noHBand")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub no_h_band: Option<OnOff>,
    #[cfg(feature = "wml-tables")]
    #[serde(rename = "@w:noVBand")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub no_v_band: Option<OnOff>,
    #[cfg(feature = "wml-tables")]
    #[serde(rename = "@w:val")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub value: Option<STShortHexNumber>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTFtnPos {
    #[serde(rename = "@w:val")]
    pub value: STFtnPos,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTEdnPos {
    #[serde(rename = "@w:val")]
    pub value: STEdnPos,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTNumFmt {
    #[serde(rename = "@w:val")]
    pub value: STNumberFormat,
    #[cfg(feature = "wml-numbering")]
    #[serde(rename = "@w:format")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub format: Option<STString>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTNumRestart {
    #[serde(rename = "@w:val")]
    pub value: STRestartNumber,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct FootnoteEndnoteRef {
    #[cfg(feature = "wml-comments")]
    #[serde(rename = "@w:customMarkFollows")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub custom_mark_follows: Option<OnOff>,
    #[serde(rename = "@w:id")]
    pub id: STDecimalNumber,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTFtnEdnSepRef {
    #[serde(rename = "@w:id")]
    pub id: STDecimalNumber,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct FootnoteEndnote {
    #[cfg(feature = "wml-comments")]
    #[serde(rename = "@w:type")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub r#type: Option<STFtnEdn>,
    #[serde(rename = "@w:id")]
    pub id: STDecimalNumber,
    #[serde(skip)]
    #[serde(default)]
    pub block_content: Vec<BlockContent>,
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
pub struct EGFtnEdnNumProps {
    #[serde(rename = "numStart")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub num_start: Option<Box<CTDecimalNumber>>,
    #[serde(rename = "numRestart")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub num_restart: Option<Box<CTNumRestart>>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct CTFtnProps {
    #[cfg(feature = "wml-comments")]
    #[serde(rename = "pos")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub pos: Option<Box<CTFtnPos>>,
    #[cfg(feature = "wml-comments")]
    #[serde(rename = "numFmt")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub num_fmt: Option<Box<CTNumFmt>>,
    #[serde(rename = "numStart")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub num_start: Option<Box<CTDecimalNumber>>,
    #[serde(rename = "numRestart")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub num_restart: Option<Box<CTNumRestart>>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct CTEdnProps {
    #[cfg(feature = "wml-comments")]
    #[serde(rename = "pos")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub pos: Option<Box<CTEdnPos>>,
    #[cfg(feature = "wml-comments")]
    #[serde(rename = "numFmt")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub num_fmt: Option<Box<CTNumFmt>>,
    #[serde(rename = "numStart")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub num_start: Option<Box<CTDecimalNumber>>,
    #[serde(rename = "numRestart")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub num_restart: Option<Box<CTNumRestart>>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct CTFtnDocProps {
    #[serde(rename = "pos")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub pos: Option<Box<CTFtnPos>>,
    #[serde(rename = "numFmt")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub num_fmt: Option<Box<CTNumFmt>>,
    #[serde(rename = "numStart")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub num_start: Option<Box<CTDecimalNumber>>,
    #[serde(rename = "numRestart")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub num_restart: Option<Box<CTNumRestart>>,
    #[cfg(feature = "wml-comments")]
    #[serde(rename = "footnote")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub footnote: Vec<CTFtnEdnSepRef>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct CTEdnDocProps {
    #[serde(rename = "pos")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub pos: Option<Box<CTEdnPos>>,
    #[serde(rename = "numFmt")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub num_fmt: Option<Box<CTNumFmt>>,
    #[serde(rename = "numStart")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub num_start: Option<Box<CTDecimalNumber>>,
    #[serde(rename = "numRestart")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub num_restart: Option<Box<CTNumRestart>>,
    #[cfg(feature = "wml-comments")]
    #[serde(rename = "endnote")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub endnote: Vec<CTFtnEdnSepRef>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTRecipientData {
    #[cfg(feature = "wml-settings")]
    #[serde(rename = "active")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub active: Option<Box<CTOnOff>>,
    #[cfg(feature = "wml-settings")]
    #[serde(rename = "column")]
    pub column: Box<CTDecimalNumber>,
    #[cfg(feature = "wml-settings")]
    #[serde(rename = "uniqueTag")]
    pub unique_tag: Box<CTBase64Binary>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTBase64Binary {
    #[serde(rename = "@w:val")]
    pub value: Vec<u8>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct CTRecipients {
    #[cfg(feature = "wml-settings")]
    #[serde(rename = "recipientData")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub recipient_data: Vec<CTRecipientData>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

pub type WRecipients = Box<CTRecipients>;

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct CTOdsoFieldMapData {
    #[cfg(feature = "wml-settings")]
    #[serde(rename = "type")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub r#type: Option<Box<CTMailMergeOdsoFMDFieldType>>,
    #[cfg(feature = "wml-settings")]
    #[serde(rename = "name")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub name: Option<Box<CTString>>,
    #[cfg(feature = "wml-settings")]
    #[serde(rename = "mappedName")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub mapped_name: Option<Box<CTString>>,
    #[cfg(feature = "wml-settings")]
    #[serde(rename = "column")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub column: Option<Box<CTDecimalNumber>>,
    #[cfg(feature = "wml-settings")]
    #[serde(rename = "lid")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub lid: Option<Box<CTLang>>,
    #[cfg(feature = "wml-settings")]
    #[serde(rename = "dynamicAddress")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub dynamic_address: Option<Box<CTOnOff>>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTMailMergeSourceType {
    #[serde(rename = "@w:val")]
    pub value: STMailMergeSourceType,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct CTOdso {
    #[cfg(feature = "wml-settings")]
    #[serde(rename = "udl")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub udl: Option<Box<CTString>>,
    #[cfg(feature = "wml-settings")]
    #[serde(rename = "table")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub table: Option<Box<CTString>>,
    #[cfg(feature = "wml-settings")]
    #[serde(rename = "src")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub src: Option<Box<CTRel>>,
    #[cfg(feature = "wml-settings")]
    #[serde(rename = "colDelim")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub col_delim: Option<Box<CTDecimalNumber>>,
    #[cfg(feature = "wml-settings")]
    #[serde(rename = "type")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub r#type: Option<Box<CTMailMergeSourceType>>,
    #[cfg(feature = "wml-settings")]
    #[serde(rename = "fHdr")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub f_hdr: Option<Box<CTOnOff>>,
    #[cfg(feature = "wml-settings")]
    #[serde(rename = "fieldMapData")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub field_map_data: Vec<CTOdsoFieldMapData>,
    #[cfg(feature = "wml-settings")]
    #[serde(rename = "recipientData")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub recipient_data: Vec<CTRel>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTMailMerge {
    #[cfg(feature = "wml-settings")]
    #[serde(rename = "mainDocumentType")]
    pub main_document_type: Box<CTMailMergeDocType>,
    #[cfg(feature = "wml-settings")]
    #[serde(rename = "linkToQuery")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub link_to_query: Option<Box<CTOnOff>>,
    #[cfg(feature = "wml-settings")]
    #[serde(rename = "dataType")]
    pub data_type: Box<CTMailMergeDataType>,
    #[cfg(feature = "wml-settings")]
    #[serde(rename = "connectString")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub connect_string: Option<Box<CTString>>,
    #[cfg(feature = "wml-settings")]
    #[serde(rename = "query")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub query: Option<Box<CTString>>,
    #[cfg(feature = "wml-settings")]
    #[serde(rename = "dataSource")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub data_source: Option<Box<CTRel>>,
    #[cfg(feature = "wml-settings")]
    #[serde(rename = "headerSource")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub header_source: Option<Box<CTRel>>,
    #[cfg(feature = "wml-settings")]
    #[serde(rename = "doNotSuppressBlankLines")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub do_not_suppress_blank_lines: Option<Box<CTOnOff>>,
    #[cfg(feature = "wml-settings")]
    #[serde(rename = "destination")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub destination: Option<Box<CTMailMergeDest>>,
    #[cfg(feature = "wml-settings")]
    #[serde(rename = "addressFieldName")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub address_field_name: Option<Box<CTString>>,
    #[cfg(feature = "wml-settings")]
    #[serde(rename = "mailSubject")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub mail_subject: Option<Box<CTString>>,
    #[cfg(feature = "wml-settings")]
    #[serde(rename = "mailAsAttachment")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub mail_as_attachment: Option<Box<CTOnOff>>,
    #[cfg(feature = "wml-settings")]
    #[serde(rename = "viewMergedData")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub view_merged_data: Option<Box<CTOnOff>>,
    #[cfg(feature = "wml-settings")]
    #[serde(rename = "activeRecord")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub active_record: Option<Box<CTDecimalNumber>>,
    #[cfg(feature = "wml-settings")]
    #[serde(rename = "checkErrors")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub check_errors: Option<Box<CTDecimalNumber>>,
    #[cfg(feature = "wml-settings")]
    #[serde(rename = "odso")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub odso: Option<Box<CTOdso>>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTTargetScreenSz {
    #[serde(rename = "@w:val")]
    pub value: STTargetScreenSz,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct Compatibility {
    #[cfg(feature = "wml-settings")]
    #[serde(rename = "useSingleBorderforContiguousCells")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub use_single_borderfor_contiguous_cells: Option<Box<CTOnOff>>,
    #[cfg(feature = "wml-settings")]
    #[serde(rename = "wpJustification")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub wp_justification: Option<Box<CTOnOff>>,
    #[cfg(feature = "wml-settings")]
    #[serde(rename = "noTabHangInd")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub no_tab_hang_ind: Option<Box<CTOnOff>>,
    #[cfg(feature = "wml-settings")]
    #[serde(rename = "noLeading")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub no_leading: Option<Box<CTOnOff>>,
    #[cfg(feature = "wml-settings")]
    #[serde(rename = "spaceForUL")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub space_for_u_l: Option<Box<CTOnOff>>,
    #[cfg(feature = "wml-settings")]
    #[serde(rename = "noColumnBalance")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub no_column_balance: Option<Box<CTOnOff>>,
    #[cfg(feature = "wml-settings")]
    #[serde(rename = "balanceSingleByteDoubleByteWidth")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub balance_single_byte_double_byte_width: Option<Box<CTOnOff>>,
    #[cfg(feature = "wml-settings")]
    #[serde(rename = "noExtraLineSpacing")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub no_extra_line_spacing: Option<Box<CTOnOff>>,
    #[cfg(feature = "wml-settings")]
    #[serde(rename = "doNotLeaveBackslashAlone")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub do_not_leave_backslash_alone: Option<Box<CTOnOff>>,
    #[cfg(feature = "wml-settings")]
    #[serde(rename = "ulTrailSpace")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub ul_trail_space: Option<Box<CTOnOff>>,
    #[cfg(feature = "wml-settings")]
    #[serde(rename = "doNotExpandShiftReturn")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub do_not_expand_shift_return: Option<Box<CTOnOff>>,
    #[cfg(feature = "wml-settings")]
    #[serde(rename = "spacingInWholePoints")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub spacing_in_whole_points: Option<Box<CTOnOff>>,
    #[cfg(feature = "wml-settings")]
    #[serde(rename = "lineWrapLikeWord6")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub line_wrap_like_word6: Option<Box<CTOnOff>>,
    #[cfg(feature = "wml-settings")]
    #[serde(rename = "printBodyTextBeforeHeader")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub print_body_text_before_header: Option<Box<CTOnOff>>,
    #[cfg(feature = "wml-settings")]
    #[serde(rename = "printColBlack")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub print_col_black: Option<Box<CTOnOff>>,
    #[cfg(feature = "wml-settings")]
    #[serde(rename = "wpSpaceWidth")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub wp_space_width: Option<Box<CTOnOff>>,
    #[cfg(feature = "wml-settings")]
    #[serde(rename = "showBreaksInFrames")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub show_breaks_in_frames: Option<Box<CTOnOff>>,
    #[cfg(feature = "wml-settings")]
    #[serde(rename = "subFontBySize")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub sub_font_by_size: Option<Box<CTOnOff>>,
    #[cfg(feature = "wml-settings")]
    #[serde(rename = "suppressBottomSpacing")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub suppress_bottom_spacing: Option<Box<CTOnOff>>,
    #[cfg(feature = "wml-settings")]
    #[serde(rename = "suppressTopSpacing")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub suppress_top_spacing: Option<Box<CTOnOff>>,
    #[cfg(feature = "wml-settings")]
    #[serde(rename = "suppressSpacingAtTopOfPage")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub suppress_spacing_at_top_of_page: Option<Box<CTOnOff>>,
    #[cfg(feature = "wml-settings")]
    #[serde(rename = "suppressTopSpacingWP")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub suppress_top_spacing_w_p: Option<Box<CTOnOff>>,
    #[cfg(feature = "wml-settings")]
    #[serde(rename = "suppressSpBfAfterPgBrk")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub suppress_sp_bf_after_pg_brk: Option<Box<CTOnOff>>,
    #[cfg(feature = "wml-settings")]
    #[serde(rename = "swapBordersFacingPages")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub swap_borders_facing_pages: Option<Box<CTOnOff>>,
    #[cfg(feature = "wml-settings")]
    #[serde(rename = "convMailMergeEsc")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub conv_mail_merge_esc: Option<Box<CTOnOff>>,
    #[cfg(feature = "wml-settings")]
    #[serde(rename = "truncateFontHeightsLikeWP6")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub truncate_font_heights_like_w_p6: Option<Box<CTOnOff>>,
    #[cfg(feature = "wml-settings")]
    #[serde(rename = "mwSmallCaps")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub mw_small_caps: Option<Box<CTOnOff>>,
    #[cfg(feature = "wml-settings")]
    #[serde(rename = "usePrinterMetrics")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub use_printer_metrics: Option<Box<CTOnOff>>,
    #[cfg(feature = "wml-settings")]
    #[serde(rename = "doNotSuppressParagraphBorders")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub do_not_suppress_paragraph_borders: Option<Box<CTOnOff>>,
    #[cfg(feature = "wml-settings")]
    #[serde(rename = "wrapTrailSpaces")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub wrap_trail_spaces: Option<Box<CTOnOff>>,
    #[cfg(feature = "wml-settings")]
    #[serde(rename = "footnoteLayoutLikeWW8")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub footnote_layout_like_w_w8: Option<Box<CTOnOff>>,
    #[cfg(feature = "wml-settings")]
    #[serde(rename = "shapeLayoutLikeWW8")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub shape_layout_like_w_w8: Option<Box<CTOnOff>>,
    #[cfg(feature = "wml-settings")]
    #[serde(rename = "alignTablesRowByRow")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub align_tables_row_by_row: Option<Box<CTOnOff>>,
    #[cfg(feature = "wml-settings")]
    #[serde(rename = "forgetLastTabAlignment")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub forget_last_tab_alignment: Option<Box<CTOnOff>>,
    #[cfg(feature = "wml-settings")]
    #[serde(rename = "adjustLineHeightInTable")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub adjust_line_height_in_table: Option<Box<CTOnOff>>,
    #[cfg(feature = "wml-settings")]
    #[serde(rename = "autoSpaceLikeWord95")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub auto_space_like_word95: Option<Box<CTOnOff>>,
    #[cfg(feature = "wml-settings")]
    #[serde(rename = "noSpaceRaiseLower")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub no_space_raise_lower: Option<Box<CTOnOff>>,
    #[cfg(feature = "wml-settings")]
    #[serde(rename = "doNotUseHTMLParagraphAutoSpacing")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub do_not_use_h_t_m_l_paragraph_auto_spacing: Option<Box<CTOnOff>>,
    #[cfg(feature = "wml-settings")]
    #[serde(rename = "layoutRawTableWidth")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub layout_raw_table_width: Option<Box<CTOnOff>>,
    #[cfg(feature = "wml-settings")]
    #[serde(rename = "layoutTableRowsApart")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub layout_table_rows_apart: Option<Box<CTOnOff>>,
    #[cfg(feature = "wml-settings")]
    #[serde(rename = "useWord97LineBreakRules")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub use_word97_line_break_rules: Option<Box<CTOnOff>>,
    #[cfg(feature = "wml-settings")]
    #[serde(rename = "doNotBreakWrappedTables")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub do_not_break_wrapped_tables: Option<Box<CTOnOff>>,
    #[cfg(feature = "wml-settings")]
    #[serde(rename = "doNotSnapToGridInCell")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub do_not_snap_to_grid_in_cell: Option<Box<CTOnOff>>,
    #[cfg(feature = "wml-settings")]
    #[serde(rename = "selectFldWithFirstOrLastChar")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub select_fld_with_first_or_last_char: Option<Box<CTOnOff>>,
    #[cfg(feature = "wml-settings")]
    #[serde(rename = "applyBreakingRules")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub apply_breaking_rules: Option<Box<CTOnOff>>,
    #[cfg(feature = "wml-settings")]
    #[serde(rename = "doNotWrapTextWithPunct")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub do_not_wrap_text_with_punct: Option<Box<CTOnOff>>,
    #[cfg(feature = "wml-settings")]
    #[serde(rename = "doNotUseEastAsianBreakRules")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub do_not_use_east_asian_break_rules: Option<Box<CTOnOff>>,
    #[cfg(feature = "wml-settings")]
    #[serde(rename = "useWord2002TableStyleRules")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub use_word2002_table_style_rules: Option<Box<CTOnOff>>,
    #[cfg(feature = "wml-settings")]
    #[serde(rename = "growAutofit")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub grow_autofit: Option<Box<CTOnOff>>,
    #[cfg(feature = "wml-settings")]
    #[serde(rename = "useFELayout")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub use_f_e_layout: Option<Box<CTOnOff>>,
    #[cfg(feature = "wml-settings")]
    #[serde(rename = "useNormalStyleForList")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub use_normal_style_for_list: Option<Box<CTOnOff>>,
    #[cfg(feature = "wml-settings")]
    #[serde(rename = "doNotUseIndentAsNumberingTabStop")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub do_not_use_indent_as_numbering_tab_stop: Option<Box<CTOnOff>>,
    #[cfg(feature = "wml-settings")]
    #[serde(rename = "useAltKinsokuLineBreakRules")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub use_alt_kinsoku_line_break_rules: Option<Box<CTOnOff>>,
    #[cfg(feature = "wml-settings")]
    #[serde(rename = "allowSpaceOfSameStyleInTable")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub allow_space_of_same_style_in_table: Option<Box<CTOnOff>>,
    #[cfg(feature = "wml-settings")]
    #[serde(rename = "doNotSuppressIndentation")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub do_not_suppress_indentation: Option<Box<CTOnOff>>,
    #[cfg(feature = "wml-settings")]
    #[serde(rename = "doNotAutofitConstrainedTables")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub do_not_autofit_constrained_tables: Option<Box<CTOnOff>>,
    #[cfg(feature = "wml-settings")]
    #[serde(rename = "autofitToFirstFixedWidthCell")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub autofit_to_first_fixed_width_cell: Option<Box<CTOnOff>>,
    #[cfg(feature = "wml-settings")]
    #[serde(rename = "underlineTabInNumList")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub underline_tab_in_num_list: Option<Box<CTOnOff>>,
    #[cfg(feature = "wml-settings")]
    #[serde(rename = "displayHangulFixedWidth")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub display_hangul_fixed_width: Option<Box<CTOnOff>>,
    #[cfg(feature = "wml-settings")]
    #[serde(rename = "splitPgBreakAndParaMark")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub split_pg_break_and_para_mark: Option<Box<CTOnOff>>,
    #[cfg(feature = "wml-settings")]
    #[serde(rename = "doNotVertAlignCellWithSp")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub do_not_vert_align_cell_with_sp: Option<Box<CTOnOff>>,
    #[cfg(feature = "wml-settings")]
    #[serde(rename = "doNotBreakConstrainedForcedTable")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub do_not_break_constrained_forced_table: Option<Box<CTOnOff>>,
    #[cfg(feature = "wml-settings")]
    #[serde(rename = "doNotVertAlignInTxbx")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub do_not_vert_align_in_txbx: Option<Box<CTOnOff>>,
    #[cfg(feature = "wml-settings")]
    #[serde(rename = "useAnsiKerningPairs")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub use_ansi_kerning_pairs: Option<Box<CTOnOff>>,
    #[cfg(feature = "wml-settings")]
    #[serde(rename = "cachedColBalance")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub cached_col_balance: Option<Box<CTOnOff>>,
    #[cfg(feature = "wml-settings")]
    #[serde(rename = "compatSetting")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub compat_setting: Vec<CTCompatSetting>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct CTCompatSetting {
    #[serde(rename = "@w:name")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub name: Option<STString>,
    #[serde(rename = "@w:uri")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub uri: Option<STString>,
    #[serde(rename = "@w:val")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub value: Option<STString>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTDocVar {
    #[serde(rename = "@w:name")]
    pub name: STString,
    #[serde(rename = "@w:val")]
    pub value: STString,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct CTDocVars {
    #[cfg(feature = "wml-settings")]
    #[serde(rename = "docVar")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub doc_var: Vec<CTDocVar>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct CTDocRsids {
    #[cfg(feature = "wml-track-changes")]
    #[serde(rename = "rsidRoot")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub rsid_root: Option<Box<CTLongHexNumber>>,
    #[cfg(feature = "wml-track-changes")]
    #[serde(rename = "rsid")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub rsid: Vec<CTLongHexNumber>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTCharacterSpacing {
    #[serde(rename = "@w:val")]
    pub value: STCharacterSpacing,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct CTSaveThroughXslt {
    #[cfg(feature = "wml-settings")]
    #[serde(rename = "@r:id")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub id: Option<STRelationshipId>,
    #[cfg(feature = "wml-settings")]
    #[serde(rename = "@w:solutionID")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub solution_i_d: Option<STString>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct RunPropertiesDefault {
    #[serde(rename = "rPr")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub r_pr: Option<Box<RunProperties>>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct ParagraphPropertiesDefault {
    #[serde(rename = "pPr")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub p_pr: Option<Box<CTPPrGeneral>>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct DocumentDefaults {
    #[serde(rename = "rPrDefault")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub r_pr_default: Option<Box<RunPropertiesDefault>>,
    #[serde(rename = "pPrDefault")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub p_pr_default: Option<Box<ParagraphPropertiesDefault>>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct CTColorSchemeMapping {
    #[cfg(feature = "wml-settings")]
    #[serde(rename = "@w:bg1")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub bg1: Option<STWmlColorSchemeIndex>,
    #[cfg(feature = "wml-settings")]
    #[serde(rename = "@w:t1")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub t1: Option<STWmlColorSchemeIndex>,
    #[cfg(feature = "wml-settings")]
    #[serde(rename = "@w:bg2")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub bg2: Option<STWmlColorSchemeIndex>,
    #[cfg(feature = "wml-settings")]
    #[serde(rename = "@w:t2")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub t2: Option<STWmlColorSchemeIndex>,
    #[cfg(feature = "wml-settings")]
    #[serde(rename = "@w:accent1")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub accent1: Option<STWmlColorSchemeIndex>,
    #[cfg(feature = "wml-settings")]
    #[serde(rename = "@w:accent2")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub accent2: Option<STWmlColorSchemeIndex>,
    #[cfg(feature = "wml-settings")]
    #[serde(rename = "@w:accent3")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub accent3: Option<STWmlColorSchemeIndex>,
    #[cfg(feature = "wml-settings")]
    #[serde(rename = "@w:accent4")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub accent4: Option<STWmlColorSchemeIndex>,
    #[cfg(feature = "wml-settings")]
    #[serde(rename = "@w:accent5")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub accent5: Option<STWmlColorSchemeIndex>,
    #[cfg(feature = "wml-settings")]
    #[serde(rename = "@w:accent6")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub accent6: Option<STWmlColorSchemeIndex>,
    #[cfg(feature = "wml-settings")]
    #[serde(rename = "@w:hyperlink")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub hyperlink: Option<STWmlColorSchemeIndex>,
    #[cfg(feature = "wml-settings")]
    #[serde(rename = "@w:followedHyperlink")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub followed_hyperlink: Option<STWmlColorSchemeIndex>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTReadingModeInkLockDown {
    #[cfg(feature = "wml-settings")]
    #[serde(rename = "@w:actualPg")]
    pub actual_pg: OnOff,
    #[cfg(feature = "wml-settings")]
    #[serde(rename = "@w:w")]
    pub width: STPixelsMeasure,
    #[cfg(feature = "wml-settings")]
    #[serde(rename = "@w:h")]
    pub height: STPixelsMeasure,
    #[cfg(feature = "wml-settings")]
    #[serde(rename = "@w:fontSz")]
    pub font_sz: STDecimalNumberOrPercent,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct CTWriteProtection {
    #[cfg(feature = "wml-settings")]
    #[serde(rename = "@w:recommended")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub recommended: Option<OnOff>,
    #[cfg(feature = "wml-settings")]
    #[serde(rename = "@w:algorithmName")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub algorithm_name: Option<STString>,
    #[cfg(feature = "wml-settings")]
    #[serde(rename = "@w:hashValue")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub hash_value: Option<Vec<u8>>,
    #[cfg(feature = "wml-settings")]
    #[serde(rename = "@w:saltValue")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub salt_value: Option<Vec<u8>>,
    #[cfg(feature = "wml-settings")]
    #[serde(rename = "@w:spinCount")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub spin_count: Option<STDecimalNumber>,
    #[cfg(feature = "wml-settings")]
    #[serde(rename = "@w:cryptProviderType")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub crypt_provider_type: Option<STCryptProv>,
    #[cfg(feature = "wml-settings")]
    #[serde(rename = "@w:cryptAlgorithmClass")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub crypt_algorithm_class: Option<STAlgClass>,
    #[cfg(feature = "wml-settings")]
    #[serde(rename = "@w:cryptAlgorithmType")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub crypt_algorithm_type: Option<STAlgType>,
    #[cfg(feature = "wml-settings")]
    #[serde(rename = "@w:cryptAlgorithmSid")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub crypt_algorithm_sid: Option<STDecimalNumber>,
    #[cfg(feature = "wml-settings")]
    #[serde(rename = "@w:cryptSpinCount")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub crypt_spin_count: Option<STDecimalNumber>,
    #[cfg(feature = "wml-settings")]
    #[serde(rename = "@w:cryptProvider")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub crypt_provider: Option<STString>,
    #[cfg(feature = "wml-settings")]
    #[serde(rename = "@w:algIdExt")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub alg_id_ext: Option<STLongHexNumber>,
    #[cfg(feature = "wml-settings")]
    #[serde(rename = "@w:algIdExtSource")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub alg_id_ext_source: Option<STString>,
    #[cfg(feature = "wml-settings")]
    #[serde(rename = "@w:cryptProviderTypeExt")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub crypt_provider_type_ext: Option<STLongHexNumber>,
    #[cfg(feature = "wml-settings")]
    #[serde(rename = "@w:cryptProviderTypeExtSource")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub crypt_provider_type_ext_source: Option<STString>,
    #[cfg(feature = "wml-settings")]
    #[serde(rename = "@w:hash")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub hash: Option<Vec<u8>>,
    #[cfg(feature = "wml-settings")]
    #[serde(rename = "@w:salt")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub salt: Option<Vec<u8>>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct Settings {
    #[cfg(feature = "wml-settings")]
    #[serde(rename = "writeProtection")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub write_protection: Option<Box<CTWriteProtection>>,
    #[cfg(feature = "wml-settings")]
    #[serde(rename = "view")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub view: Option<Box<CTView>>,
    #[cfg(feature = "wml-settings")]
    #[serde(rename = "zoom")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub zoom: Option<Box<CTZoom>>,
    #[cfg(feature = "wml-settings")]
    #[serde(rename = "removePersonalInformation")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub remove_personal_information: Option<Box<CTOnOff>>,
    #[cfg(feature = "wml-settings")]
    #[serde(rename = "removeDateAndTime")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub remove_date_and_time: Option<Box<CTOnOff>>,
    #[cfg(feature = "wml-settings")]
    #[serde(rename = "doNotDisplayPageBoundaries")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub do_not_display_page_boundaries: Option<Box<CTOnOff>>,
    #[cfg(feature = "wml-settings")]
    #[serde(rename = "displayBackgroundShape")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub display_background_shape: Option<Box<CTOnOff>>,
    #[cfg(feature = "wml-settings")]
    #[serde(rename = "printPostScriptOverText")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub print_post_script_over_text: Option<Box<CTOnOff>>,
    #[cfg(feature = "wml-settings")]
    #[serde(rename = "printFractionalCharacterWidth")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub print_fractional_character_width: Option<Box<CTOnOff>>,
    #[cfg(feature = "wml-settings")]
    #[serde(rename = "printFormsData")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub print_forms_data: Option<Box<CTOnOff>>,
    #[cfg(feature = "wml-settings")]
    #[serde(rename = "embedTrueTypeFonts")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub embed_true_type_fonts: Option<Box<CTOnOff>>,
    #[cfg(feature = "wml-settings")]
    #[serde(rename = "embedSystemFonts")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub embed_system_fonts: Option<Box<CTOnOff>>,
    #[cfg(feature = "wml-settings")]
    #[serde(rename = "saveSubsetFonts")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub save_subset_fonts: Option<Box<CTOnOff>>,
    #[cfg(feature = "wml-settings")]
    #[serde(rename = "saveFormsData")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub save_forms_data: Option<Box<CTOnOff>>,
    #[cfg(feature = "wml-settings")]
    #[serde(rename = "mirrorMargins")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub mirror_margins: Option<Box<CTOnOff>>,
    #[cfg(feature = "wml-settings")]
    #[serde(rename = "alignBordersAndEdges")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub align_borders_and_edges: Option<Box<CTOnOff>>,
    #[cfg(feature = "wml-settings")]
    #[serde(rename = "bordersDoNotSurroundHeader")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub borders_do_not_surround_header: Option<Box<CTOnOff>>,
    #[cfg(feature = "wml-settings")]
    #[serde(rename = "bordersDoNotSurroundFooter")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub borders_do_not_surround_footer: Option<Box<CTOnOff>>,
    #[cfg(feature = "wml-settings")]
    #[serde(rename = "gutterAtTop")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub gutter_at_top: Option<Box<CTOnOff>>,
    #[cfg(feature = "wml-settings")]
    #[serde(rename = "hideSpellingErrors")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub hide_spelling_errors: Option<Box<CTOnOff>>,
    #[cfg(feature = "wml-settings")]
    #[serde(rename = "hideGrammaticalErrors")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub hide_grammatical_errors: Option<Box<CTOnOff>>,
    #[cfg(feature = "wml-settings")]
    #[serde(rename = "activeWritingStyle")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub active_writing_style: Vec<CTWritingStyle>,
    #[cfg(feature = "wml-settings")]
    #[serde(rename = "proofState")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub proof_state: Option<Box<CTProof>>,
    #[cfg(feature = "wml-settings")]
    #[serde(rename = "formsDesign")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub forms_design: Option<Box<CTOnOff>>,
    #[cfg(feature = "wml-settings")]
    #[serde(rename = "attachedTemplate")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub attached_template: Option<Box<CTRel>>,
    #[cfg(feature = "wml-settings")]
    #[serde(rename = "linkStyles")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub link_styles: Option<Box<CTOnOff>>,
    #[cfg(feature = "wml-settings")]
    #[serde(rename = "stylePaneFormatFilter")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub style_pane_format_filter: Option<Box<CTStylePaneFilter>>,
    #[cfg(feature = "wml-settings")]
    #[serde(rename = "stylePaneSortMethod")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub style_pane_sort_method: Option<Box<CTStyleSort>>,
    #[cfg(feature = "wml-settings")]
    #[serde(rename = "documentType")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub document_type: Option<Box<CTDocType>>,
    #[cfg(feature = "wml-settings")]
    #[serde(rename = "mailMerge")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub mail_merge: Option<Box<CTMailMerge>>,
    #[cfg(feature = "wml-settings")]
    #[serde(rename = "revisionView")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub revision_view: Option<Box<CTTrackChangesView>>,
    #[cfg(feature = "wml-settings")]
    #[serde(rename = "trackRevisions")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub track_revisions: Option<Box<CTOnOff>>,
    #[cfg(feature = "wml-track-changes")]
    #[serde(rename = "doNotTrackMoves")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub do_not_track_moves: Option<Box<CTOnOff>>,
    #[cfg(feature = "wml-track-changes")]
    #[serde(rename = "doNotTrackFormatting")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub do_not_track_formatting: Option<Box<CTOnOff>>,
    #[cfg(feature = "wml-settings")]
    #[serde(rename = "documentProtection")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub document_protection: Option<Box<CTDocProtect>>,
    #[cfg(feature = "wml-settings")]
    #[serde(rename = "autoFormatOverride")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub auto_format_override: Option<Box<CTOnOff>>,
    #[cfg(feature = "wml-settings")]
    #[serde(rename = "styleLockTheme")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub style_lock_theme: Option<Box<CTOnOff>>,
    #[cfg(feature = "wml-settings")]
    #[serde(rename = "styleLockQFSet")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub style_lock_q_f_set: Option<Box<CTOnOff>>,
    #[cfg(feature = "wml-settings")]
    #[serde(rename = "defaultTabStop")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub default_tab_stop: Option<Box<CTTwipsMeasure>>,
    #[cfg(feature = "wml-settings")]
    #[serde(rename = "autoHyphenation")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub auto_hyphenation: Option<Box<CTOnOff>>,
    #[cfg(feature = "wml-settings")]
    #[serde(rename = "consecutiveHyphenLimit")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub consecutive_hyphen_limit: Option<Box<CTDecimalNumber>>,
    #[cfg(feature = "wml-settings")]
    #[serde(rename = "hyphenationZone")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub hyphenation_zone: Option<Box<CTTwipsMeasure>>,
    #[cfg(feature = "wml-settings")]
    #[serde(rename = "doNotHyphenateCaps")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub do_not_hyphenate_caps: Option<Box<CTOnOff>>,
    #[cfg(feature = "wml-settings")]
    #[serde(rename = "showEnvelope")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub show_envelope: Option<Box<CTOnOff>>,
    #[cfg(feature = "wml-settings")]
    #[serde(rename = "summaryLength")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub summary_length: Option<Box<CTDecimalNumberOrPrecent>>,
    #[cfg(feature = "wml-settings")]
    #[serde(rename = "clickAndTypeStyle")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub click_and_type_style: Option<Box<CTString>>,
    #[cfg(feature = "wml-settings")]
    #[serde(rename = "defaultTableStyle")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub default_table_style: Option<Box<CTString>>,
    #[cfg(feature = "wml-settings")]
    #[serde(rename = "evenAndOddHeaders")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub even_and_odd_headers: Option<Box<CTOnOff>>,
    #[cfg(feature = "wml-settings")]
    #[serde(rename = "bookFoldRevPrinting")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub book_fold_rev_printing: Option<Box<CTOnOff>>,
    #[cfg(feature = "wml-settings")]
    #[serde(rename = "bookFoldPrinting")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub book_fold_printing: Option<Box<CTOnOff>>,
    #[cfg(feature = "wml-settings")]
    #[serde(rename = "bookFoldPrintingSheets")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub book_fold_printing_sheets: Option<Box<CTDecimalNumber>>,
    #[cfg(feature = "wml-settings")]
    #[serde(rename = "drawingGridHorizontalSpacing")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub drawing_grid_horizontal_spacing: Option<Box<CTTwipsMeasure>>,
    #[cfg(feature = "wml-settings")]
    #[serde(rename = "drawingGridVerticalSpacing")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub drawing_grid_vertical_spacing: Option<Box<CTTwipsMeasure>>,
    #[cfg(feature = "wml-settings")]
    #[serde(rename = "displayHorizontalDrawingGridEvery")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub display_horizontal_drawing_grid_every: Option<Box<CTDecimalNumber>>,
    #[cfg(feature = "wml-settings")]
    #[serde(rename = "displayVerticalDrawingGridEvery")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub display_vertical_drawing_grid_every: Option<Box<CTDecimalNumber>>,
    #[cfg(feature = "wml-settings")]
    #[serde(rename = "doNotUseMarginsForDrawingGridOrigin")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub do_not_use_margins_for_drawing_grid_origin: Option<Box<CTOnOff>>,
    #[cfg(feature = "wml-settings")]
    #[serde(rename = "drawingGridHorizontalOrigin")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub drawing_grid_horizontal_origin: Option<Box<CTTwipsMeasure>>,
    #[cfg(feature = "wml-settings")]
    #[serde(rename = "drawingGridVerticalOrigin")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub drawing_grid_vertical_origin: Option<Box<CTTwipsMeasure>>,
    #[cfg(feature = "wml-settings")]
    #[serde(rename = "doNotShadeFormData")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub do_not_shade_form_data: Option<Box<CTOnOff>>,
    #[cfg(feature = "wml-settings")]
    #[serde(rename = "noPunctuationKerning")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub no_punctuation_kerning: Option<Box<CTOnOff>>,
    #[cfg(feature = "wml-settings")]
    #[serde(rename = "characterSpacingControl")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub character_spacing_control: Option<Box<CTCharacterSpacing>>,
    #[cfg(feature = "wml-settings")]
    #[serde(rename = "printTwoOnOne")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub print_two_on_one: Option<Box<CTOnOff>>,
    #[cfg(feature = "wml-settings")]
    #[serde(rename = "strictFirstAndLastChars")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub strict_first_and_last_chars: Option<Box<CTOnOff>>,
    #[cfg(feature = "wml-settings")]
    #[serde(rename = "noLineBreaksAfter")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub no_line_breaks_after: Option<Box<CTKinsoku>>,
    #[cfg(feature = "wml-settings")]
    #[serde(rename = "noLineBreaksBefore")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub no_line_breaks_before: Option<Box<CTKinsoku>>,
    #[cfg(feature = "wml-settings")]
    #[serde(rename = "savePreviewPicture")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub save_preview_picture: Option<Box<CTOnOff>>,
    #[cfg(feature = "wml-settings")]
    #[serde(rename = "doNotValidateAgainstSchema")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub do_not_validate_against_schema: Option<Box<CTOnOff>>,
    #[cfg(feature = "wml-settings")]
    #[serde(rename = "saveInvalidXml")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub save_invalid_xml: Option<Box<CTOnOff>>,
    #[cfg(feature = "wml-settings")]
    #[serde(rename = "ignoreMixedContent")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub ignore_mixed_content: Option<Box<CTOnOff>>,
    #[cfg(feature = "wml-settings")]
    #[serde(rename = "alwaysShowPlaceholderText")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub always_show_placeholder_text: Option<Box<CTOnOff>>,
    #[cfg(feature = "wml-settings")]
    #[serde(rename = "doNotDemarcateInvalidXml")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub do_not_demarcate_invalid_xml: Option<Box<CTOnOff>>,
    #[cfg(feature = "wml-settings")]
    #[serde(rename = "saveXmlDataOnly")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub save_xml_data_only: Option<Box<CTOnOff>>,
    #[cfg(feature = "wml-settings")]
    #[serde(rename = "useXSLTWhenSaving")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub use_x_s_l_t_when_saving: Option<Box<CTOnOff>>,
    #[cfg(feature = "wml-settings")]
    #[serde(rename = "saveThroughXslt")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub save_through_xslt: Option<Box<CTSaveThroughXslt>>,
    #[cfg(feature = "wml-settings")]
    #[serde(rename = "showXMLTags")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub show_x_m_l_tags: Option<Box<CTOnOff>>,
    #[cfg(feature = "wml-settings")]
    #[serde(rename = "alwaysMergeEmptyNamespace")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub always_merge_empty_namespace: Option<Box<CTOnOff>>,
    #[cfg(feature = "wml-settings")]
    #[serde(rename = "updateFields")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub update_fields: Option<Box<CTOnOff>>,
    #[cfg(feature = "wml-drawings")]
    #[serde(rename = "hdrShapeDefaults")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub hdr_shape_defaults: Option<Box<CTShapeDefaults>>,
    #[cfg(feature = "wml-comments")]
    #[serde(rename = "footnotePr")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub footnote_pr: Option<Box<CTFtnDocProps>>,
    #[cfg(feature = "wml-comments")]
    #[serde(rename = "endnotePr")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub endnote_pr: Option<Box<CTEdnDocProps>>,
    #[cfg(feature = "wml-settings")]
    #[serde(rename = "compat")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub compat: Option<Box<Compatibility>>,
    #[cfg(feature = "wml-settings")]
    #[serde(rename = "docVars")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub doc_vars: Option<Box<CTDocVars>>,
    #[cfg(feature = "wml-settings")]
    #[serde(rename = "rsids")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub rsids: Option<Box<CTDocRsids>>,
    #[cfg(feature = "wml-settings")]
    #[serde(rename = "attachedSchema")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub attached_schema: Vec<CTString>,
    #[cfg(feature = "wml-settings")]
    #[serde(rename = "themeFontLang")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub theme_font_lang: Option<Box<CTLanguage>>,
    #[cfg(feature = "wml-settings")]
    #[serde(rename = "clrSchemeMapping")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub clr_scheme_mapping: Option<Box<CTColorSchemeMapping>>,
    #[cfg(feature = "wml-settings")]
    #[serde(rename = "doNotIncludeSubdocsInStats")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub do_not_include_subdocs_in_stats: Option<Box<CTOnOff>>,
    #[cfg(feature = "wml-settings")]
    #[serde(rename = "doNotAutoCompressPictures")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub do_not_auto_compress_pictures: Option<Box<CTOnOff>>,
    #[cfg(feature = "wml-settings")]
    #[serde(rename = "forceUpgrade")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub force_upgrade: Option<Box<CTEmpty>>,
    #[cfg(feature = "wml-settings")]
    #[serde(rename = "captions")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub captions: Option<Box<CTCaptions>>,
    #[cfg(feature = "wml-settings")]
    #[serde(rename = "readModeInkLockDown")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub read_mode_ink_lock_down: Option<Box<CTReadingModeInkLockDown>>,
    #[cfg(feature = "wml-settings")]
    #[serde(rename = "smartTagType")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub smart_tag_type: Vec<CTSmartTagType>,
    #[cfg(feature = "wml-drawings")]
    #[serde(rename = "shapeDefaults")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub shape_defaults: Option<Box<CTShapeDefaults>>,
    #[cfg(feature = "wml-settings")]
    #[serde(rename = "doNotEmbedSmartTags")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub do_not_embed_smart_tags: Option<Box<CTOnOff>>,
    #[cfg(feature = "wml-settings")]
    #[serde(rename = "decimalSymbol")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub decimal_symbol: Option<Box<CTString>>,
    #[cfg(feature = "wml-settings")]
    #[serde(rename = "listSeparator")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub list_separator: Option<Box<CTString>>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTStyleSort {
    #[serde(rename = "@w:val")]
    pub value: STStyleSort,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct CTStylePaneFilter {
    #[cfg(feature = "wml-settings")]
    #[serde(rename = "@w:allStyles")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub all_styles: Option<OnOff>,
    #[cfg(feature = "wml-settings")]
    #[serde(rename = "@w:customStyles")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub custom_styles: Option<OnOff>,
    #[cfg(feature = "wml-settings")]
    #[serde(rename = "@w:latentStyles")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub latent_styles: Option<OnOff>,
    #[cfg(feature = "wml-settings")]
    #[serde(rename = "@w:stylesInUse")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub styles_in_use: Option<OnOff>,
    #[cfg(feature = "wml-settings")]
    #[serde(rename = "@w:headingStyles")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub heading_styles: Option<OnOff>,
    #[cfg(feature = "wml-settings")]
    #[serde(rename = "@w:numberingStyles")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub numbering_styles: Option<OnOff>,
    #[cfg(feature = "wml-settings")]
    #[serde(rename = "@w:tableStyles")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub table_styles: Option<OnOff>,
    #[cfg(feature = "wml-settings")]
    #[serde(rename = "@w:directFormattingOnRuns")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub direct_formatting_on_runs: Option<OnOff>,
    #[cfg(feature = "wml-settings")]
    #[serde(rename = "@w:directFormattingOnParagraphs")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub direct_formatting_on_paragraphs: Option<OnOff>,
    #[cfg(feature = "wml-settings")]
    #[serde(rename = "@w:directFormattingOnNumbering")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub direct_formatting_on_numbering: Option<OnOff>,
    #[cfg(feature = "wml-settings")]
    #[serde(rename = "@w:directFormattingOnTables")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub direct_formatting_on_tables: Option<OnOff>,
    #[cfg(feature = "wml-settings")]
    #[serde(rename = "@w:clearFormatting")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub clear_formatting: Option<OnOff>,
    #[cfg(feature = "wml-settings")]
    #[serde(rename = "@w:top3HeadingStyles")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub top3_heading_styles: Option<OnOff>,
    #[cfg(feature = "wml-settings")]
    #[serde(rename = "@w:visibleStyles")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub visible_styles: Option<OnOff>,
    #[cfg(feature = "wml-settings")]
    #[serde(rename = "@w:alternateStyleNames")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub alternate_style_names: Option<OnOff>,
    #[cfg(feature = "wml-settings")]
    #[serde(rename = "@w:val")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub value: Option<STShortHexNumber>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct CTWebSettings {
    #[cfg(feature = "wml-settings")]
    #[serde(rename = "frameset")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub frameset: Option<Box<CTFrameset>>,
    #[cfg(feature = "wml-settings")]
    #[serde(rename = "divs")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub divs: Option<Box<CTDivs>>,
    #[cfg(feature = "wml-settings")]
    #[serde(rename = "encoding")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub encoding: Option<Box<CTString>>,
    #[cfg(feature = "wml-settings")]
    #[serde(rename = "optimizeForBrowser")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub optimize_for_browser: Option<Box<CTOptimizeForBrowser>>,
    #[cfg(feature = "wml-settings")]
    #[serde(rename = "relyOnVML")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub rely_on_v_m_l: Option<Box<CTOnOff>>,
    #[cfg(feature = "wml-settings")]
    #[serde(rename = "allowPNG")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub allow_p_n_g: Option<Box<CTOnOff>>,
    #[cfg(feature = "wml-settings")]
    #[serde(rename = "doNotRelyOnCSS")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub do_not_rely_on_c_s_s: Option<Box<CTOnOff>>,
    #[cfg(feature = "wml-settings")]
    #[serde(rename = "doNotSaveAsSingleFile")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub do_not_save_as_single_file: Option<Box<CTOnOff>>,
    #[cfg(feature = "wml-settings")]
    #[serde(rename = "doNotOrganizeInFolder")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub do_not_organize_in_folder: Option<Box<CTOnOff>>,
    #[cfg(feature = "wml-settings")]
    #[serde(rename = "doNotUseLongFileNames")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub do_not_use_long_file_names: Option<Box<CTOnOff>>,
    #[cfg(feature = "wml-settings")]
    #[serde(rename = "pixelsPerInch")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub pixels_per_inch: Option<Box<CTDecimalNumber>>,
    #[cfg(feature = "wml-settings")]
    #[serde(rename = "targetScreenSz")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub target_screen_sz: Option<Box<CTTargetScreenSz>>,
    #[cfg(feature = "wml-settings")]
    #[serde(rename = "saveSmartTagsAsXml")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub save_smart_tags_as_xml: Option<Box<CTOnOff>>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTFrameScrollbar {
    #[serde(rename = "@w:val")]
    pub value: STFrameScrollbar,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct CTOptimizeForBrowser {
    #[serde(rename = "@w:val")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub value: Option<OnOff>,
    #[cfg(feature = "wml-settings")]
    #[serde(rename = "@w:target")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub target: Option<STString>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct CTFrame {
    #[cfg(feature = "wml-settings")]
    #[serde(rename = "sz")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub size: Option<Box<CTString>>,
    #[cfg(feature = "wml-settings")]
    #[serde(rename = "name")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub name: Option<Box<CTString>>,
    #[cfg(feature = "wml-settings")]
    #[serde(rename = "title")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub title: Option<Box<CTString>>,
    #[cfg(feature = "wml-settings")]
    #[serde(rename = "longDesc")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub long_desc: Option<Box<CTRel>>,
    #[cfg(feature = "wml-settings")]
    #[serde(rename = "sourceFileName")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub source_file_name: Option<Box<CTRel>>,
    #[cfg(feature = "wml-settings")]
    #[serde(rename = "marW")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub mar_w: Option<Box<CTPixelsMeasure>>,
    #[cfg(feature = "wml-settings")]
    #[serde(rename = "marH")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub mar_h: Option<Box<CTPixelsMeasure>>,
    #[cfg(feature = "wml-settings")]
    #[serde(rename = "scrollbar")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub scrollbar: Option<Box<CTFrameScrollbar>>,
    #[cfg(feature = "wml-settings")]
    #[serde(rename = "noResizeAllowed")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub no_resize_allowed: Option<Box<CTOnOff>>,
    #[cfg(feature = "wml-settings")]
    #[serde(rename = "linkedToFile")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub linked_to_file: Option<Box<CTOnOff>>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTFrameLayout {
    #[serde(rename = "@w:val")]
    pub value: STFrameLayout,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct CTFramesetSplitbar {
    #[cfg(feature = "wml-settings")]
    #[serde(rename = "w")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub width: Option<Box<CTTwipsMeasure>>,
    #[cfg(feature = "wml-settings")]
    #[serde(rename = "color")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub color: Option<Box<CTColor>>,
    #[cfg(feature = "wml-settings")]
    #[serde(rename = "noBorder")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub no_border: Option<Box<CTOnOff>>,
    #[cfg(feature = "wml-settings")]
    #[serde(rename = "flatBorders")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub flat_borders: Option<Box<CTOnOff>>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct CTFrameset {
    #[cfg(feature = "wml-settings")]
    #[serde(rename = "sz")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub size: Option<Box<CTString>>,
    #[cfg(feature = "wml-settings")]
    #[serde(rename = "framesetSplitbar")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub frameset_splitbar: Option<Box<CTFramesetSplitbar>>,
    #[cfg(feature = "wml-settings")]
    #[serde(rename = "frameLayout")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub frame_layout: Option<Box<CTFrameLayout>>,
    #[cfg(feature = "wml-settings")]
    #[serde(rename = "title")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub title: Option<Box<CTString>>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTNumPicBullet {
    #[serde(rename = "@w:numPicBulletId")]
    pub num_pic_bullet_id: STDecimalNumber,
    #[cfg(feature = "wml-numbering")]
    #[serde(rename = "pict")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub pict: Option<Box<CTPicture>>,
    #[cfg(feature = "wml-numbering")]
    #[serde(rename = "drawing")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub drawing: Option<Box<CTDrawing>>,
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
pub struct CTLevelSuffix {
    #[serde(rename = "@w:val")]
    pub value: STLevelSuffix,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct CTLevelText {
    #[serde(rename = "@w:val")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub value: Option<STString>,
    #[cfg(feature = "wml-numbering")]
    #[serde(rename = "@w:null")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub null: Option<OnOff>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct CTLvlLegacy {
    #[cfg(feature = "wml-numbering")]
    #[serde(rename = "@w:legacy")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub legacy: Option<OnOff>,
    #[cfg(feature = "wml-numbering")]
    #[serde(rename = "@w:legacySpace")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub legacy_space: Option<STTwipsMeasure>,
    #[cfg(feature = "wml-numbering")]
    #[serde(rename = "@w:legacyIndent")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub legacy_indent: Option<STSignedTwipsMeasure>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct Level {
    #[serde(rename = "@w:ilvl")]
    pub ilvl: STDecimalNumber,
    #[cfg(feature = "wml-numbering")]
    #[serde(rename = "@w:tplc")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub tplc: Option<STLongHexNumber>,
    #[cfg(feature = "wml-numbering")]
    #[serde(rename = "@w:tentative")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub tentative: Option<OnOff>,
    #[cfg(feature = "wml-numbering")]
    #[serde(rename = "start")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub start: Option<Box<CTDecimalNumber>>,
    #[cfg(feature = "wml-numbering")]
    #[serde(rename = "numFmt")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub num_fmt: Option<Box<CTNumFmt>>,
    #[cfg(feature = "wml-numbering")]
    #[serde(rename = "lvlRestart")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub lvl_restart: Option<Box<CTDecimalNumber>>,
    #[cfg(feature = "wml-numbering")]
    #[serde(rename = "pStyle")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub paragraph_style: Option<Box<CTString>>,
    #[cfg(feature = "wml-numbering")]
    #[serde(rename = "isLgl")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub is_lgl: Option<Box<CTOnOff>>,
    #[cfg(feature = "wml-numbering")]
    #[serde(rename = "suff")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub suff: Option<Box<CTLevelSuffix>>,
    #[cfg(feature = "wml-numbering")]
    #[serde(rename = "lvlText")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub lvl_text: Option<Box<CTLevelText>>,
    #[cfg(feature = "wml-numbering")]
    #[serde(rename = "lvlPicBulletId")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub lvl_pic_bullet_id: Option<Box<CTDecimalNumber>>,
    #[cfg(feature = "wml-numbering")]
    #[serde(rename = "legacy")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub legacy: Option<Box<CTLvlLegacy>>,
    #[cfg(feature = "wml-numbering")]
    #[serde(rename = "lvlJc")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub lvl_jc: Option<Box<CTJc>>,
    #[cfg(feature = "wml-numbering")]
    #[serde(rename = "pPr")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub p_pr: Option<Box<CTPPrGeneral>>,
    #[cfg(feature = "wml-numbering")]
    #[serde(rename = "rPr")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub r_pr: Option<Box<RunProperties>>,
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
pub struct CTMultiLevelType {
    #[serde(rename = "@w:val")]
    pub value: STMultiLevelType,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct AbstractNumbering {
    #[serde(rename = "@w:abstractNumId")]
    pub abstract_num_id: STDecimalNumber,
    #[cfg(feature = "wml-numbering")]
    #[serde(rename = "nsid")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub nsid: Option<Box<CTLongHexNumber>>,
    #[cfg(feature = "wml-numbering")]
    #[serde(rename = "multiLevelType")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub multi_level_type: Option<Box<CTMultiLevelType>>,
    #[cfg(feature = "wml-numbering")]
    #[serde(rename = "tmpl")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub tmpl: Option<Box<CTLongHexNumber>>,
    #[cfg(feature = "wml-numbering")]
    #[serde(rename = "name")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub name: Option<Box<CTString>>,
    #[cfg(feature = "wml-numbering")]
    #[serde(rename = "styleLink")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub style_link: Option<Box<CTString>>,
    #[cfg(feature = "wml-numbering")]
    #[serde(rename = "numStyleLink")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub num_style_link: Option<Box<CTString>>,
    #[serde(rename = "lvl")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub lvl: Vec<Level>,
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
pub struct CTNumLvl {
    #[serde(rename = "@w:ilvl")]
    pub ilvl: STDecimalNumber,
    #[cfg(feature = "wml-numbering")]
    #[serde(rename = "startOverride")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub start_override: Option<Box<CTDecimalNumber>>,
    #[cfg(feature = "wml-numbering")]
    #[serde(rename = "lvl")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub lvl: Option<Box<Level>>,
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
pub struct NumberingInstance {
    #[serde(rename = "@w:numId")]
    pub num_id: STDecimalNumber,
    #[serde(rename = "abstractNumId")]
    pub abstract_num_id: Box<CTDecimalNumber>,
    #[cfg(feature = "wml-numbering")]
    #[serde(rename = "lvlOverride")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub lvl_override: Vec<CTNumLvl>,
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
pub struct Numbering {
    #[cfg(feature = "wml-numbering")]
    #[serde(rename = "numPicBullet")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub num_pic_bullet: Vec<CTNumPicBullet>,
    #[serde(rename = "abstractNum")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub abstract_num: Vec<AbstractNumbering>,
    #[serde(rename = "num")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub num: Vec<NumberingInstance>,
    #[cfg(feature = "wml-numbering")]
    #[serde(rename = "numIdMacAtCleanup")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub num_id_mac_at_cleanup: Option<Box<CTDecimalNumber>>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct TableStyleProperties {
    #[serde(rename = "@w:type")]
    pub r#type: STTblStyleOverrideType,
    #[cfg(feature = "wml-styling")]
    #[serde(rename = "pPr")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub p_pr: Option<Box<CTPPrGeneral>>,
    #[cfg(feature = "wml-styling")]
    #[serde(rename = "rPr")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub r_pr: Option<Box<RunProperties>>,
    #[cfg(feature = "wml-styling")]
    #[serde(rename = "tblPr")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub table_properties: Option<Box<CTTblPrBase>>,
    #[cfg(feature = "wml-styling")]
    #[serde(rename = "trPr")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub row_properties: Option<Box<TableRowProperties>>,
    #[cfg(feature = "wml-styling")]
    #[serde(rename = "tcPr")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub cell_properties: Option<Box<TableCellProperties>>,
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
pub struct Style {
    #[serde(rename = "@w:type")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub r#type: Option<STStyleType>,
    #[serde(rename = "@w:styleId")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub style_id: Option<STString>,
    #[serde(rename = "@w:default")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub default: Option<OnOff>,
    #[serde(rename = "@w:customStyle")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub custom_style: Option<OnOff>,
    #[serde(rename = "name")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub name: Option<Box<CTString>>,
    #[serde(rename = "aliases")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub aliases: Option<Box<CTString>>,
    #[serde(rename = "basedOn")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub based_on: Option<Box<CTString>>,
    #[cfg(feature = "wml-styling")]
    #[serde(rename = "next")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub next: Option<Box<CTString>>,
    #[cfg(feature = "wml-styling")]
    #[serde(rename = "link")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub link: Option<Box<CTString>>,
    #[cfg(feature = "wml-styling")]
    #[serde(rename = "autoRedefine")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub auto_redefine: Option<Box<CTOnOff>>,
    #[cfg(feature = "wml-styling")]
    #[serde(rename = "hidden")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub hidden: Option<Box<CTOnOff>>,
    #[cfg(feature = "wml-styling")]
    #[serde(rename = "uiPriority")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub ui_priority: Option<Box<CTDecimalNumber>>,
    #[cfg(feature = "wml-styling")]
    #[serde(rename = "semiHidden")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub semi_hidden: Option<Box<CTOnOff>>,
    #[cfg(feature = "wml-styling")]
    #[serde(rename = "unhideWhenUsed")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub unhide_when_used: Option<Box<CTOnOff>>,
    #[cfg(feature = "wml-styling")]
    #[serde(rename = "qFormat")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub q_format: Option<Box<CTOnOff>>,
    #[cfg(feature = "wml-settings")]
    #[serde(rename = "locked")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub locked: Option<Box<CTOnOff>>,
    #[cfg(feature = "wml-settings")]
    #[serde(rename = "personal")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub personal: Option<Box<CTOnOff>>,
    #[cfg(feature = "wml-settings")]
    #[serde(rename = "personalCompose")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub personal_compose: Option<Box<CTOnOff>>,
    #[cfg(feature = "wml-settings")]
    #[serde(rename = "personalReply")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub personal_reply: Option<Box<CTOnOff>>,
    #[cfg(feature = "wml-track-changes")]
    #[serde(rename = "rsid")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub rsid: Option<Box<CTLongHexNumber>>,
    #[cfg(feature = "wml-styling")]
    #[serde(rename = "pPr")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub p_pr: Option<Box<CTPPrGeneral>>,
    #[cfg(feature = "wml-styling")]
    #[serde(rename = "rPr")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub r_pr: Option<Box<RunProperties>>,
    #[cfg(feature = "wml-styling")]
    #[serde(rename = "tblPr")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub table_properties: Option<Box<CTTblPrBase>>,
    #[cfg(feature = "wml-styling")]
    #[serde(rename = "trPr")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub row_properties: Option<Box<TableRowProperties>>,
    #[cfg(feature = "wml-styling")]
    #[serde(rename = "tcPr")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub cell_properties: Option<Box<TableCellProperties>>,
    #[cfg(feature = "wml-styling")]
    #[serde(rename = "tblStylePr")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub tbl_style_pr: Vec<TableStyleProperties>,
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
pub struct LatentStyleException {
    #[serde(rename = "@w:name")]
    pub name: STString,
    #[cfg(feature = "wml-styling")]
    #[serde(rename = "@w:locked")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub locked: Option<OnOff>,
    #[cfg(feature = "wml-styling")]
    #[serde(rename = "@w:uiPriority")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub ui_priority: Option<STDecimalNumber>,
    #[cfg(feature = "wml-styling")]
    #[serde(rename = "@w:semiHidden")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub semi_hidden: Option<OnOff>,
    #[cfg(feature = "wml-styling")]
    #[serde(rename = "@w:unhideWhenUsed")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub unhide_when_used: Option<OnOff>,
    #[cfg(feature = "wml-styling")]
    #[serde(rename = "@w:qFormat")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub q_format: Option<OnOff>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct LatentStyles {
    #[cfg(feature = "wml-styling")]
    #[serde(rename = "@w:defLockedState")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub def_locked_state: Option<OnOff>,
    #[cfg(feature = "wml-styling")]
    #[serde(rename = "@w:defUIPriority")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub def_u_i_priority: Option<STDecimalNumber>,
    #[cfg(feature = "wml-styling")]
    #[serde(rename = "@w:defSemiHidden")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub def_semi_hidden: Option<OnOff>,
    #[cfg(feature = "wml-styling")]
    #[serde(rename = "@w:defUnhideWhenUsed")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub def_unhide_when_used: Option<OnOff>,
    #[cfg(feature = "wml-styling")]
    #[serde(rename = "@w:defQFormat")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub def_q_format: Option<OnOff>,
    #[cfg(feature = "wml-styling")]
    #[serde(rename = "@w:count")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub count: Option<STDecimalNumber>,
    #[serde(rename = "lsdException")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub lsd_exception: Vec<LatentStyleException>,
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
pub struct Styles {
    #[serde(rename = "docDefaults")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub doc_defaults: Option<Box<DocumentDefaults>>,
    #[cfg(feature = "wml-styling")]
    #[serde(rename = "latentStyles")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub latent_styles: Option<Box<LatentStyles>>,
    #[serde(rename = "style")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub style: Vec<Style>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTPanose {
    #[serde(rename = "@w:val")]
    pub value: Panose,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTFontFamily {
    #[serde(rename = "@w:val")]
    pub value: STFontFamily,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTPitch {
    #[serde(rename = "@w:val")]
    pub value: STPitch,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTFontSig {
    #[cfg(feature = "wml-styling")]
    #[serde(rename = "@w:usb0")]
    pub usb0: STLongHexNumber,
    #[cfg(feature = "wml-styling")]
    #[serde(rename = "@w:usb1")]
    pub usb1: STLongHexNumber,
    #[cfg(feature = "wml-styling")]
    #[serde(rename = "@w:usb2")]
    pub usb2: STLongHexNumber,
    #[cfg(feature = "wml-styling")]
    #[serde(rename = "@w:usb3")]
    pub usb3: STLongHexNumber,
    #[cfg(feature = "wml-styling")]
    #[serde(rename = "@w:csb0")]
    pub csb0: STLongHexNumber,
    #[cfg(feature = "wml-styling")]
    #[serde(rename = "@w:csb1")]
    pub csb1: STLongHexNumber,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTFontRel {
    #[cfg(feature = "wml-styling")]
    #[serde(rename = "@r:id")]
    pub id: STRelationshipId,
    #[cfg(feature = "wml-styling")]
    #[serde(rename = "@w:fontKey")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub font_key: Option<Guid>,
    #[cfg(feature = "wml-styling")]
    #[serde(rename = "@w:subsetted")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub subsetted: Option<OnOff>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct Font {
    #[serde(rename = "@w:name")]
    pub name: STString,
    #[cfg(feature = "wml-styling")]
    #[serde(rename = "altName")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub alt_name: Option<Box<CTString>>,
    #[cfg(feature = "wml-styling")]
    #[serde(rename = "panose1")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub panose1: Option<Box<CTPanose>>,
    #[cfg(feature = "wml-styling")]
    #[serde(rename = "charset")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub charset: Option<Box<CTCharset>>,
    #[cfg(feature = "wml-styling")]
    #[serde(rename = "family")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub family: Option<Box<CTFontFamily>>,
    #[cfg(feature = "wml-styling")]
    #[serde(rename = "notTrueType")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub not_true_type: Option<Box<CTOnOff>>,
    #[cfg(feature = "wml-styling")]
    #[serde(rename = "pitch")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub pitch: Option<Box<CTPitch>>,
    #[cfg(feature = "wml-styling")]
    #[serde(rename = "sig")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub sig: Option<Box<CTFontSig>>,
    #[cfg(feature = "wml-styling")]
    #[serde(rename = "embedRegular")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub embed_regular: Option<Box<CTFontRel>>,
    #[cfg(feature = "wml-styling")]
    #[serde(rename = "embedBold")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub embed_bold: Option<Box<CTFontRel>>,
    #[cfg(feature = "wml-styling")]
    #[serde(rename = "embedItalic")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub embed_italic: Option<Box<CTFontRel>>,
    #[cfg(feature = "wml-styling")]
    #[serde(rename = "embedBoldItalic")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub embed_bold_italic: Option<Box<CTFontRel>>,
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
pub struct CTFontsList {
    #[serde(rename = "font")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub font: Vec<Font>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct CTDivBdr {
    #[cfg(feature = "wml-settings")]
    #[serde(rename = "top")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub top: Option<Box<CTBorder>>,
    #[cfg(feature = "wml-settings")]
    #[serde(rename = "left")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub left: Option<Box<CTBorder>>,
    #[cfg(feature = "wml-settings")]
    #[serde(rename = "bottom")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub bottom: Option<Box<CTBorder>>,
    #[cfg(feature = "wml-settings")]
    #[serde(rename = "right")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub right: Option<Box<CTBorder>>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTDiv {
    #[cfg(feature = "wml-settings")]
    #[serde(rename = "@w:id")]
    pub id: STDecimalNumber,
    #[cfg(feature = "wml-settings")]
    #[serde(rename = "blockQuote")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub block_quote: Option<Box<CTOnOff>>,
    #[cfg(feature = "wml-settings")]
    #[serde(rename = "bodyDiv")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub body_div: Option<Box<CTOnOff>>,
    #[cfg(feature = "wml-settings")]
    #[serde(rename = "marLeft")]
    pub mar_left: Box<CTSignedTwipsMeasure>,
    #[cfg(feature = "wml-settings")]
    #[serde(rename = "marRight")]
    pub mar_right: Box<CTSignedTwipsMeasure>,
    #[cfg(feature = "wml-settings")]
    #[serde(rename = "marTop")]
    pub mar_top: Box<CTSignedTwipsMeasure>,
    #[cfg(feature = "wml-settings")]
    #[serde(rename = "marBottom")]
    pub mar_bottom: Box<CTSignedTwipsMeasure>,
    #[cfg(feature = "wml-settings")]
    #[serde(rename = "divBdr")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub div_bdr: Option<Box<CTDivBdr>>,
    #[cfg(feature = "wml-settings")]
    #[serde(rename = "divsChild")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub divs_child: Vec<CTDivs>,
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
pub struct CTDivs {
    #[cfg(feature = "wml-settings")]
    #[serde(rename = "div")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub div: Vec<CTDiv>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct CTTxbxContent {
    #[serde(skip)]
    #[serde(default)]
    pub block_content: Vec<BlockContent>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

pub type WTxbxContent = Box<CTTxbxContent>;

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct EGMathContent {
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct EGBlockLevelChunkElts {
    #[serde(skip)]
    #[serde(default)]
    pub block_content: Vec<BlockContentChoice>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct Body {
    #[serde(skip)]
    #[serde(default)]
    pub block_content: Vec<BlockContent>,
    #[cfg(feature = "wml-layout")]
    #[serde(rename = "sectPr")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub sect_pr: Option<Box<SectionProperties>>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct CTShapeDefaults {
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct Comments {
    #[serde(rename = "comment")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub comment: Vec<Comment>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

pub type WComments = Box<Comments>;

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct Footnotes {
    #[serde(rename = "footnote")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub footnote: Vec<FootnoteEndnote>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

pub type WFootnotes = Box<Footnotes>;

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct Endnotes {
    #[serde(rename = "endnote")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub endnote: Vec<FootnoteEndnote>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

pub type WEndnotes = Box<Endnotes>;

pub type WHdr = Box<HeaderFooter>;

pub type WFtr = Box<HeaderFooter>;

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct CTSmartTagType {
    #[cfg(feature = "wml-settings")]
    #[serde(rename = "@w:namespaceuri")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub namespaceuri: Option<STString>,
    #[cfg(feature = "wml-settings")]
    #[serde(rename = "@w:name")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub name: Option<STString>,
    #[cfg(feature = "wml-settings")]
    #[serde(rename = "@w:url")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub url: Option<STString>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTDocPartBehavior {
    #[serde(rename = "@w:val")]
    pub value: STDocPartBehavior,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct CTDocPartBehaviors {
    #[cfg(feature = "wml-settings")]
    #[serde(rename = "behavior")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub behavior: Vec<CTDocPartBehavior>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTDocPartType {
    #[serde(rename = "@w:val")]
    pub value: STDocPartType,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct CTDocPartTypes {
    #[cfg(feature = "wml-settings")]
    #[serde(rename = "@w:all")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub all: Option<OnOff>,
    #[cfg(feature = "wml-settings")]
    #[serde(rename = "type")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub r#type: Vec<CTDocPartType>,
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
pub struct CTDocPartGallery {
    #[serde(rename = "@w:val")]
    pub value: STDocPartGallery,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTDocPartCategory {
    #[serde(rename = "name")]
    pub name: Box<CTString>,
    #[cfg(feature = "wml-settings")]
    #[serde(rename = "gallery")]
    pub gallery: Box<CTDocPartGallery>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTDocPartName {
    #[serde(rename = "@w:val")]
    pub value: STString,
    #[cfg(feature = "wml-settings")]
    #[serde(rename = "@w:decorated")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub decorated: Option<OnOff>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTDocPartPr {
    #[cfg(feature = "wml-settings")]
    #[serde(rename = "name")]
    pub name: Box<CTDocPartName>,
    #[cfg(feature = "wml-settings")]
    #[serde(rename = "style")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub style: Option<Box<CTString>>,
    #[cfg(feature = "wml-settings")]
    #[serde(rename = "category")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub category: Option<Box<CTDocPartCategory>>,
    #[cfg(feature = "wml-settings")]
    #[serde(rename = "types")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub types: Option<Box<CTDocPartTypes>>,
    #[cfg(feature = "wml-settings")]
    #[serde(rename = "behaviors")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub behaviors: Option<Box<CTDocPartBehaviors>>,
    #[cfg(feature = "wml-settings")]
    #[serde(rename = "description")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub description: Option<Box<CTString>>,
    #[cfg(feature = "wml-settings")]
    #[serde(rename = "guid")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub guid: Option<Box<CTGuid>>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct CTDocPart {
    #[cfg(feature = "wml-settings")]
    #[serde(rename = "docPartPr")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub doc_part_pr: Option<Box<CTDocPartPr>>,
    #[cfg(feature = "wml-settings")]
    #[serde(rename = "docPartBody")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub doc_part_body: Option<Box<Body>>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct CTDocParts {
    #[cfg(feature = "wml-settings")]
    #[serde(rename = "docPart")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub doc_part: Vec<CTDocPart>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

pub type WSettings = Box<Settings>;

pub type WWebSettings = Box<CTWebSettings>;

pub type WFonts = Box<CTFontsList>;

pub type WNumbering = Box<Numbering>;

pub type WStyles = Box<Styles>;

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTCaption {
    #[cfg(feature = "wml-settings")]
    #[serde(rename = "@w:name")]
    pub name: STString,
    #[cfg(feature = "wml-settings")]
    #[serde(rename = "@w:pos")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub pos: Option<STCaptionPos>,
    #[cfg(feature = "wml-settings")]
    #[serde(rename = "@w:chapNum")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub chap_num: Option<OnOff>,
    #[cfg(feature = "wml-settings")]
    #[serde(rename = "@w:heading")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub heading: Option<STDecimalNumber>,
    #[cfg(feature = "wml-settings")]
    #[serde(rename = "@w:noLabel")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub no_label: Option<OnOff>,
    #[cfg(feature = "wml-settings")]
    #[serde(rename = "@w:numFmt")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub num_fmt: Option<STNumberFormat>,
    #[cfg(feature = "wml-settings")]
    #[serde(rename = "@w:sep")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub sep: Option<STChapterSep>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTAutoCaption {
    #[cfg(feature = "wml-settings")]
    #[serde(rename = "@w:name")]
    pub name: STString,
    #[cfg(feature = "wml-settings")]
    #[serde(rename = "@w:caption")]
    pub caption: STString,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct CTAutoCaptions {
    #[cfg(feature = "wml-settings")]
    #[serde(rename = "autoCaption")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub auto_caption: Vec<CTAutoCaption>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct CTCaptions {
    #[cfg(feature = "wml-settings")]
    #[serde(rename = "caption")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub caption: Vec<CTCaption>,
    #[cfg(feature = "wml-settings")]
    #[serde(rename = "autoCaptions")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub auto_captions: Option<Box<CTAutoCaptions>>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct CTDocumentBase {
    #[cfg(feature = "wml-styling")]
    #[serde(rename = "background")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub background: Option<Box<CTBackground>>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct Document {
    #[cfg(feature = "wml-styling")]
    #[serde(rename = "background")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub background: Option<Box<CTBackground>>,
    #[serde(rename = "body")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub body: Option<Box<Body>>,
    #[serde(rename = "@w:conformance")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub conformance: Option<STConformanceClass>,
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
pub struct CTGlossaryDocument {
    #[serde(rename = "background")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub background: Option<Box<CTBackground>>,
    #[cfg(feature = "wml-settings")]
    #[serde(rename = "docParts")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub doc_parts: Option<Box<CTDocParts>>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

pub type WDocument = Box<Document>;

pub type WGlossaryDocument = Box<CTGlossaryDocument>;

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct WAnyVmlOffice {
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct WAnyVmlVml {
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::PositionedNode>,
}
