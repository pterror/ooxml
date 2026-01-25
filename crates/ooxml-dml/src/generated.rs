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

pub type STStyleMatrixColumnIndex = u32;

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum STFontCollectionIndex {
    #[serde(rename = "major")]
    Major,
    #[serde(rename = "minor")]
    Minor,
    #[serde(rename = "none")]
    None,
}

impl std::fmt::Display for STFontCollectionIndex {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::Major => write!(f, "major"),
            Self::Minor => write!(f, "minor"),
            Self::None => write!(f, "none"),
        }
    }
}

impl std::str::FromStr for STFontCollectionIndex {
    type Err = String;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        match s {
            "major" => Ok(Self::Major),
            "minor" => Ok(Self::Minor),
            "none" => Ok(Self::None),
            _ => Err(format!("unknown STFontCollectionIndex value: {}", s)),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum STColorSchemeIndex {
    #[serde(rename = "dk1")]
    Dk1,
    #[serde(rename = "lt1")]
    Lt1,
    #[serde(rename = "dk2")]
    Dk2,
    #[serde(rename = "lt2")]
    Lt2,
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
    #[serde(rename = "hlink")]
    Hlink,
    #[serde(rename = "folHlink")]
    FolHlink,
}

impl std::fmt::Display for STColorSchemeIndex {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::Dk1 => write!(f, "dk1"),
            Self::Lt1 => write!(f, "lt1"),
            Self::Dk2 => write!(f, "dk2"),
            Self::Lt2 => write!(f, "lt2"),
            Self::Accent1 => write!(f, "accent1"),
            Self::Accent2 => write!(f, "accent2"),
            Self::Accent3 => write!(f, "accent3"),
            Self::Accent4 => write!(f, "accent4"),
            Self::Accent5 => write!(f, "accent5"),
            Self::Accent6 => write!(f, "accent6"),
            Self::Hlink => write!(f, "hlink"),
            Self::FolHlink => write!(f, "folHlink"),
        }
    }
}

impl std::str::FromStr for STColorSchemeIndex {
    type Err = String;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        match s {
            "dk1" => Ok(Self::Dk1),
            "lt1" => Ok(Self::Lt1),
            "dk2" => Ok(Self::Dk2),
            "lt2" => Ok(Self::Lt2),
            "accent1" => Ok(Self::Accent1),
            "accent2" => Ok(Self::Accent2),
            "accent3" => Ok(Self::Accent3),
            "accent4" => Ok(Self::Accent4),
            "accent5" => Ok(Self::Accent5),
            "accent6" => Ok(Self::Accent6),
            "hlink" => Ok(Self::Hlink),
            "folHlink" => Ok(Self::FolHlink),
            _ => Err(format!("unknown STColorSchemeIndex value: {}", s)),
        }
    }
}

pub type STCoordinate = String;

pub type STCoordinateUnqualified = i64;

pub type STCoordinate32 = String;

pub type STCoordinate32Unqualified = i32;

pub type STPositiveCoordinate = i64;

pub type STPositiveCoordinate32 = i32;

pub type STAngle = i32;

pub type STFixedAngle = i32;

pub type STPositiveFixedAngle = i32;

pub type STPercentageDecimal = i32;

pub type STPositivePercentageDecimal = i32;

pub type STFixedPercentageDecimal = i32;

pub type STPositiveFixedPercentageDecimal = i32;

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum STSystemColorVal {
    #[serde(rename = "scrollBar")]
    ScrollBar,
    #[serde(rename = "background")]
    Background,
    #[serde(rename = "activeCaption")]
    ActiveCaption,
    #[serde(rename = "inactiveCaption")]
    InactiveCaption,
    #[serde(rename = "menu")]
    Menu,
    #[serde(rename = "window")]
    Window,
    #[serde(rename = "windowFrame")]
    WindowFrame,
    #[serde(rename = "menuText")]
    MenuText,
    #[serde(rename = "windowText")]
    WindowText,
    #[serde(rename = "captionText")]
    CaptionText,
    #[serde(rename = "activeBorder")]
    ActiveBorder,
    #[serde(rename = "inactiveBorder")]
    InactiveBorder,
    #[serde(rename = "appWorkspace")]
    AppWorkspace,
    #[serde(rename = "highlight")]
    Highlight,
    #[serde(rename = "highlightText")]
    HighlightText,
    #[serde(rename = "btnFace")]
    BtnFace,
    #[serde(rename = "btnShadow")]
    BtnShadow,
    #[serde(rename = "grayText")]
    GrayText,
    #[serde(rename = "btnText")]
    BtnText,
    #[serde(rename = "inactiveCaptionText")]
    InactiveCaptionText,
    #[serde(rename = "btnHighlight")]
    BtnHighlight,
    #[serde(rename = "3dDkShadow")]
    _3dDkShadow,
    #[serde(rename = "3dLight")]
    _3dLight,
    #[serde(rename = "infoText")]
    InfoText,
    #[serde(rename = "infoBk")]
    InfoBk,
    #[serde(rename = "hotLight")]
    HotLight,
    #[serde(rename = "gradientActiveCaption")]
    GradientActiveCaption,
    #[serde(rename = "gradientInactiveCaption")]
    GradientInactiveCaption,
    #[serde(rename = "menuHighlight")]
    MenuHighlight,
    #[serde(rename = "menuBar")]
    MenuBar,
}

impl std::fmt::Display for STSystemColorVal {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::ScrollBar => write!(f, "scrollBar"),
            Self::Background => write!(f, "background"),
            Self::ActiveCaption => write!(f, "activeCaption"),
            Self::InactiveCaption => write!(f, "inactiveCaption"),
            Self::Menu => write!(f, "menu"),
            Self::Window => write!(f, "window"),
            Self::WindowFrame => write!(f, "windowFrame"),
            Self::MenuText => write!(f, "menuText"),
            Self::WindowText => write!(f, "windowText"),
            Self::CaptionText => write!(f, "captionText"),
            Self::ActiveBorder => write!(f, "activeBorder"),
            Self::InactiveBorder => write!(f, "inactiveBorder"),
            Self::AppWorkspace => write!(f, "appWorkspace"),
            Self::Highlight => write!(f, "highlight"),
            Self::HighlightText => write!(f, "highlightText"),
            Self::BtnFace => write!(f, "btnFace"),
            Self::BtnShadow => write!(f, "btnShadow"),
            Self::GrayText => write!(f, "grayText"),
            Self::BtnText => write!(f, "btnText"),
            Self::InactiveCaptionText => write!(f, "inactiveCaptionText"),
            Self::BtnHighlight => write!(f, "btnHighlight"),
            Self::_3dDkShadow => write!(f, "3dDkShadow"),
            Self::_3dLight => write!(f, "3dLight"),
            Self::InfoText => write!(f, "infoText"),
            Self::InfoBk => write!(f, "infoBk"),
            Self::HotLight => write!(f, "hotLight"),
            Self::GradientActiveCaption => write!(f, "gradientActiveCaption"),
            Self::GradientInactiveCaption => write!(f, "gradientInactiveCaption"),
            Self::MenuHighlight => write!(f, "menuHighlight"),
            Self::MenuBar => write!(f, "menuBar"),
        }
    }
}

impl std::str::FromStr for STSystemColorVal {
    type Err = String;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        match s {
            "scrollBar" => Ok(Self::ScrollBar),
            "background" => Ok(Self::Background),
            "activeCaption" => Ok(Self::ActiveCaption),
            "inactiveCaption" => Ok(Self::InactiveCaption),
            "menu" => Ok(Self::Menu),
            "window" => Ok(Self::Window),
            "windowFrame" => Ok(Self::WindowFrame),
            "menuText" => Ok(Self::MenuText),
            "windowText" => Ok(Self::WindowText),
            "captionText" => Ok(Self::CaptionText),
            "activeBorder" => Ok(Self::ActiveBorder),
            "inactiveBorder" => Ok(Self::InactiveBorder),
            "appWorkspace" => Ok(Self::AppWorkspace),
            "highlight" => Ok(Self::Highlight),
            "highlightText" => Ok(Self::HighlightText),
            "btnFace" => Ok(Self::BtnFace),
            "btnShadow" => Ok(Self::BtnShadow),
            "grayText" => Ok(Self::GrayText),
            "btnText" => Ok(Self::BtnText),
            "inactiveCaptionText" => Ok(Self::InactiveCaptionText),
            "btnHighlight" => Ok(Self::BtnHighlight),
            "3dDkShadow" => Ok(Self::_3dDkShadow),
            "3dLight" => Ok(Self::_3dLight),
            "infoText" => Ok(Self::InfoText),
            "infoBk" => Ok(Self::InfoBk),
            "hotLight" => Ok(Self::HotLight),
            "gradientActiveCaption" => Ok(Self::GradientActiveCaption),
            "gradientInactiveCaption" => Ok(Self::GradientInactiveCaption),
            "menuHighlight" => Ok(Self::MenuHighlight),
            "menuBar" => Ok(Self::MenuBar),
            _ => Err(format!("unknown STSystemColorVal value: {}", s)),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum STSchemeColorVal {
    #[serde(rename = "bg1")]
    Bg1,
    #[serde(rename = "tx1")]
    Tx1,
    #[serde(rename = "bg2")]
    Bg2,
    #[serde(rename = "tx2")]
    Tx2,
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
    #[serde(rename = "hlink")]
    Hlink,
    #[serde(rename = "folHlink")]
    FolHlink,
    #[serde(rename = "phClr")]
    PhClr,
    #[serde(rename = "dk1")]
    Dk1,
    #[serde(rename = "lt1")]
    Lt1,
    #[serde(rename = "dk2")]
    Dk2,
    #[serde(rename = "lt2")]
    Lt2,
}

impl std::fmt::Display for STSchemeColorVal {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::Bg1 => write!(f, "bg1"),
            Self::Tx1 => write!(f, "tx1"),
            Self::Bg2 => write!(f, "bg2"),
            Self::Tx2 => write!(f, "tx2"),
            Self::Accent1 => write!(f, "accent1"),
            Self::Accent2 => write!(f, "accent2"),
            Self::Accent3 => write!(f, "accent3"),
            Self::Accent4 => write!(f, "accent4"),
            Self::Accent5 => write!(f, "accent5"),
            Self::Accent6 => write!(f, "accent6"),
            Self::Hlink => write!(f, "hlink"),
            Self::FolHlink => write!(f, "folHlink"),
            Self::PhClr => write!(f, "phClr"),
            Self::Dk1 => write!(f, "dk1"),
            Self::Lt1 => write!(f, "lt1"),
            Self::Dk2 => write!(f, "dk2"),
            Self::Lt2 => write!(f, "lt2"),
        }
    }
}

impl std::str::FromStr for STSchemeColorVal {
    type Err = String;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        match s {
            "bg1" => Ok(Self::Bg1),
            "tx1" => Ok(Self::Tx1),
            "bg2" => Ok(Self::Bg2),
            "tx2" => Ok(Self::Tx2),
            "accent1" => Ok(Self::Accent1),
            "accent2" => Ok(Self::Accent2),
            "accent3" => Ok(Self::Accent3),
            "accent4" => Ok(Self::Accent4),
            "accent5" => Ok(Self::Accent5),
            "accent6" => Ok(Self::Accent6),
            "hlink" => Ok(Self::Hlink),
            "folHlink" => Ok(Self::FolHlink),
            "phClr" => Ok(Self::PhClr),
            "dk1" => Ok(Self::Dk1),
            "lt1" => Ok(Self::Lt1),
            "dk2" => Ok(Self::Dk2),
            "lt2" => Ok(Self::Lt2),
            _ => Err(format!("unknown STSchemeColorVal value: {}", s)),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum STPresetColorVal {
    #[serde(rename = "aliceBlue")]
    AliceBlue,
    #[serde(rename = "antiqueWhite")]
    AntiqueWhite,
    #[serde(rename = "aqua")]
    Aqua,
    #[serde(rename = "aquamarine")]
    Aquamarine,
    #[serde(rename = "azure")]
    Azure,
    #[serde(rename = "beige")]
    Beige,
    #[serde(rename = "bisque")]
    Bisque,
    #[serde(rename = "black")]
    Black,
    #[serde(rename = "blanchedAlmond")]
    BlanchedAlmond,
    #[serde(rename = "blue")]
    Blue,
    #[serde(rename = "blueViolet")]
    BlueViolet,
    #[serde(rename = "brown")]
    Brown,
    #[serde(rename = "burlyWood")]
    BurlyWood,
    #[serde(rename = "cadetBlue")]
    CadetBlue,
    #[serde(rename = "chartreuse")]
    Chartreuse,
    #[serde(rename = "chocolate")]
    Chocolate,
    #[serde(rename = "coral")]
    Coral,
    #[serde(rename = "cornflowerBlue")]
    CornflowerBlue,
    #[serde(rename = "cornsilk")]
    Cornsilk,
    #[serde(rename = "crimson")]
    Crimson,
    #[serde(rename = "cyan")]
    Cyan,
    #[serde(rename = "darkBlue")]
    DarkBlue,
    #[serde(rename = "darkCyan")]
    DarkCyan,
    #[serde(rename = "darkGoldenrod")]
    DarkGoldenrod,
    #[serde(rename = "darkGray")]
    DarkGray,
    #[serde(rename = "darkGrey")]
    DarkGrey,
    #[serde(rename = "darkGreen")]
    DarkGreen,
    #[serde(rename = "darkKhaki")]
    DarkKhaki,
    #[serde(rename = "darkMagenta")]
    DarkMagenta,
    #[serde(rename = "darkOliveGreen")]
    DarkOliveGreen,
    #[serde(rename = "darkOrange")]
    DarkOrange,
    #[serde(rename = "darkOrchid")]
    DarkOrchid,
    #[serde(rename = "darkRed")]
    DarkRed,
    #[serde(rename = "darkSalmon")]
    DarkSalmon,
    #[serde(rename = "darkSeaGreen")]
    DarkSeaGreen,
    #[serde(rename = "darkSlateBlue")]
    DarkSlateBlue,
    #[serde(rename = "darkSlateGray")]
    DarkSlateGray,
    #[serde(rename = "darkSlateGrey")]
    DarkSlateGrey,
    #[serde(rename = "darkTurquoise")]
    DarkTurquoise,
    #[serde(rename = "darkViolet")]
    DarkViolet,
    #[serde(rename = "dkBlue")]
    DkBlue,
    #[serde(rename = "dkCyan")]
    DkCyan,
    #[serde(rename = "dkGoldenrod")]
    DkGoldenrod,
    #[serde(rename = "dkGray")]
    DkGray,
    #[serde(rename = "dkGrey")]
    DkGrey,
    #[serde(rename = "dkGreen")]
    DkGreen,
    #[serde(rename = "dkKhaki")]
    DkKhaki,
    #[serde(rename = "dkMagenta")]
    DkMagenta,
    #[serde(rename = "dkOliveGreen")]
    DkOliveGreen,
    #[serde(rename = "dkOrange")]
    DkOrange,
    #[serde(rename = "dkOrchid")]
    DkOrchid,
    #[serde(rename = "dkRed")]
    DkRed,
    #[serde(rename = "dkSalmon")]
    DkSalmon,
    #[serde(rename = "dkSeaGreen")]
    DkSeaGreen,
    #[serde(rename = "dkSlateBlue")]
    DkSlateBlue,
    #[serde(rename = "dkSlateGray")]
    DkSlateGray,
    #[serde(rename = "dkSlateGrey")]
    DkSlateGrey,
    #[serde(rename = "dkTurquoise")]
    DkTurquoise,
    #[serde(rename = "dkViolet")]
    DkViolet,
    #[serde(rename = "deepPink")]
    DeepPink,
    #[serde(rename = "deepSkyBlue")]
    DeepSkyBlue,
    #[serde(rename = "dimGray")]
    DimGray,
    #[serde(rename = "dimGrey")]
    DimGrey,
    #[serde(rename = "dodgerBlue")]
    DodgerBlue,
    #[serde(rename = "firebrick")]
    Firebrick,
    #[serde(rename = "floralWhite")]
    FloralWhite,
    #[serde(rename = "forestGreen")]
    ForestGreen,
    #[serde(rename = "fuchsia")]
    Fuchsia,
    #[serde(rename = "gainsboro")]
    Gainsboro,
    #[serde(rename = "ghostWhite")]
    GhostWhite,
    #[serde(rename = "gold")]
    Gold,
    #[serde(rename = "goldenrod")]
    Goldenrod,
    #[serde(rename = "gray")]
    Gray,
    #[serde(rename = "grey")]
    Grey,
    #[serde(rename = "green")]
    Green,
    #[serde(rename = "greenYellow")]
    GreenYellow,
    #[serde(rename = "honeydew")]
    Honeydew,
    #[serde(rename = "hotPink")]
    HotPink,
    #[serde(rename = "indianRed")]
    IndianRed,
    #[serde(rename = "indigo")]
    Indigo,
    #[serde(rename = "ivory")]
    Ivory,
    #[serde(rename = "khaki")]
    Khaki,
    #[serde(rename = "lavender")]
    Lavender,
    #[serde(rename = "lavenderBlush")]
    LavenderBlush,
    #[serde(rename = "lawnGreen")]
    LawnGreen,
    #[serde(rename = "lemonChiffon")]
    LemonChiffon,
    #[serde(rename = "lightBlue")]
    LightBlue,
    #[serde(rename = "lightCoral")]
    LightCoral,
    #[serde(rename = "lightCyan")]
    LightCyan,
    #[serde(rename = "lightGoldenrodYellow")]
    LightGoldenrodYellow,
    #[serde(rename = "lightGray")]
    LightGray,
    #[serde(rename = "lightGrey")]
    LightGrey,
    #[serde(rename = "lightGreen")]
    LightGreen,
    #[serde(rename = "lightPink")]
    LightPink,
    #[serde(rename = "lightSalmon")]
    LightSalmon,
    #[serde(rename = "lightSeaGreen")]
    LightSeaGreen,
    #[serde(rename = "lightSkyBlue")]
    LightSkyBlue,
    #[serde(rename = "lightSlateGray")]
    LightSlateGray,
    #[serde(rename = "lightSlateGrey")]
    LightSlateGrey,
    #[serde(rename = "lightSteelBlue")]
    LightSteelBlue,
    #[serde(rename = "lightYellow")]
    LightYellow,
    #[serde(rename = "ltBlue")]
    LtBlue,
    #[serde(rename = "ltCoral")]
    LtCoral,
    #[serde(rename = "ltCyan")]
    LtCyan,
    #[serde(rename = "ltGoldenrodYellow")]
    LtGoldenrodYellow,
    #[serde(rename = "ltGray")]
    LtGray,
    #[serde(rename = "ltGrey")]
    LtGrey,
    #[serde(rename = "ltGreen")]
    LtGreen,
    #[serde(rename = "ltPink")]
    LtPink,
    #[serde(rename = "ltSalmon")]
    LtSalmon,
    #[serde(rename = "ltSeaGreen")]
    LtSeaGreen,
    #[serde(rename = "ltSkyBlue")]
    LtSkyBlue,
    #[serde(rename = "ltSlateGray")]
    LtSlateGray,
    #[serde(rename = "ltSlateGrey")]
    LtSlateGrey,
    #[serde(rename = "ltSteelBlue")]
    LtSteelBlue,
    #[serde(rename = "ltYellow")]
    LtYellow,
    #[serde(rename = "lime")]
    Lime,
    #[serde(rename = "limeGreen")]
    LimeGreen,
    #[serde(rename = "linen")]
    Linen,
    #[serde(rename = "magenta")]
    Magenta,
    #[serde(rename = "maroon")]
    Maroon,
    #[serde(rename = "medAquamarine")]
    MedAquamarine,
    #[serde(rename = "medBlue")]
    MedBlue,
    #[serde(rename = "medOrchid")]
    MedOrchid,
    #[serde(rename = "medPurple")]
    MedPurple,
    #[serde(rename = "medSeaGreen")]
    MedSeaGreen,
    #[serde(rename = "medSlateBlue")]
    MedSlateBlue,
    #[serde(rename = "medSpringGreen")]
    MedSpringGreen,
    #[serde(rename = "medTurquoise")]
    MedTurquoise,
    #[serde(rename = "medVioletRed")]
    MedVioletRed,
    #[serde(rename = "mediumAquamarine")]
    MediumAquamarine,
    #[serde(rename = "mediumBlue")]
    MediumBlue,
    #[serde(rename = "mediumOrchid")]
    MediumOrchid,
    #[serde(rename = "mediumPurple")]
    MediumPurple,
    #[serde(rename = "mediumSeaGreen")]
    MediumSeaGreen,
    #[serde(rename = "mediumSlateBlue")]
    MediumSlateBlue,
    #[serde(rename = "mediumSpringGreen")]
    MediumSpringGreen,
    #[serde(rename = "mediumTurquoise")]
    MediumTurquoise,
    #[serde(rename = "mediumVioletRed")]
    MediumVioletRed,
    #[serde(rename = "midnightBlue")]
    MidnightBlue,
    #[serde(rename = "mintCream")]
    MintCream,
    #[serde(rename = "mistyRose")]
    MistyRose,
    #[serde(rename = "moccasin")]
    Moccasin,
    #[serde(rename = "navajoWhite")]
    NavajoWhite,
    #[serde(rename = "navy")]
    Navy,
    #[serde(rename = "oldLace")]
    OldLace,
    #[serde(rename = "olive")]
    Olive,
    #[serde(rename = "oliveDrab")]
    OliveDrab,
    #[serde(rename = "orange")]
    Orange,
    #[serde(rename = "orangeRed")]
    OrangeRed,
    #[serde(rename = "orchid")]
    Orchid,
    #[serde(rename = "paleGoldenrod")]
    PaleGoldenrod,
    #[serde(rename = "paleGreen")]
    PaleGreen,
    #[serde(rename = "paleTurquoise")]
    PaleTurquoise,
    #[serde(rename = "paleVioletRed")]
    PaleVioletRed,
    #[serde(rename = "papayaWhip")]
    PapayaWhip,
    #[serde(rename = "peachPuff")]
    PeachPuff,
    #[serde(rename = "peru")]
    Peru,
    #[serde(rename = "pink")]
    Pink,
    #[serde(rename = "plum")]
    Plum,
    #[serde(rename = "powderBlue")]
    PowderBlue,
    #[serde(rename = "purple")]
    Purple,
    #[serde(rename = "red")]
    Red,
    #[serde(rename = "rosyBrown")]
    RosyBrown,
    #[serde(rename = "royalBlue")]
    RoyalBlue,
    #[serde(rename = "saddleBrown")]
    SaddleBrown,
    #[serde(rename = "salmon")]
    Salmon,
    #[serde(rename = "sandyBrown")]
    SandyBrown,
    #[serde(rename = "seaGreen")]
    SeaGreen,
    #[serde(rename = "seaShell")]
    SeaShell,
    #[serde(rename = "sienna")]
    Sienna,
    #[serde(rename = "silver")]
    Silver,
    #[serde(rename = "skyBlue")]
    SkyBlue,
    #[serde(rename = "slateBlue")]
    SlateBlue,
    #[serde(rename = "slateGray")]
    SlateGray,
    #[serde(rename = "slateGrey")]
    SlateGrey,
    #[serde(rename = "snow")]
    Snow,
    #[serde(rename = "springGreen")]
    SpringGreen,
    #[serde(rename = "steelBlue")]
    SteelBlue,
    #[serde(rename = "tan")]
    Tan,
    #[serde(rename = "teal")]
    Teal,
    #[serde(rename = "thistle")]
    Thistle,
    #[serde(rename = "tomato")]
    Tomato,
    #[serde(rename = "turquoise")]
    Turquoise,
    #[serde(rename = "violet")]
    Violet,
    #[serde(rename = "wheat")]
    Wheat,
    #[serde(rename = "white")]
    White,
    #[serde(rename = "whiteSmoke")]
    WhiteSmoke,
    #[serde(rename = "yellow")]
    Yellow,
    #[serde(rename = "yellowGreen")]
    YellowGreen,
}

impl std::fmt::Display for STPresetColorVal {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::AliceBlue => write!(f, "aliceBlue"),
            Self::AntiqueWhite => write!(f, "antiqueWhite"),
            Self::Aqua => write!(f, "aqua"),
            Self::Aquamarine => write!(f, "aquamarine"),
            Self::Azure => write!(f, "azure"),
            Self::Beige => write!(f, "beige"),
            Self::Bisque => write!(f, "bisque"),
            Self::Black => write!(f, "black"),
            Self::BlanchedAlmond => write!(f, "blanchedAlmond"),
            Self::Blue => write!(f, "blue"),
            Self::BlueViolet => write!(f, "blueViolet"),
            Self::Brown => write!(f, "brown"),
            Self::BurlyWood => write!(f, "burlyWood"),
            Self::CadetBlue => write!(f, "cadetBlue"),
            Self::Chartreuse => write!(f, "chartreuse"),
            Self::Chocolate => write!(f, "chocolate"),
            Self::Coral => write!(f, "coral"),
            Self::CornflowerBlue => write!(f, "cornflowerBlue"),
            Self::Cornsilk => write!(f, "cornsilk"),
            Self::Crimson => write!(f, "crimson"),
            Self::Cyan => write!(f, "cyan"),
            Self::DarkBlue => write!(f, "darkBlue"),
            Self::DarkCyan => write!(f, "darkCyan"),
            Self::DarkGoldenrod => write!(f, "darkGoldenrod"),
            Self::DarkGray => write!(f, "darkGray"),
            Self::DarkGrey => write!(f, "darkGrey"),
            Self::DarkGreen => write!(f, "darkGreen"),
            Self::DarkKhaki => write!(f, "darkKhaki"),
            Self::DarkMagenta => write!(f, "darkMagenta"),
            Self::DarkOliveGreen => write!(f, "darkOliveGreen"),
            Self::DarkOrange => write!(f, "darkOrange"),
            Self::DarkOrchid => write!(f, "darkOrchid"),
            Self::DarkRed => write!(f, "darkRed"),
            Self::DarkSalmon => write!(f, "darkSalmon"),
            Self::DarkSeaGreen => write!(f, "darkSeaGreen"),
            Self::DarkSlateBlue => write!(f, "darkSlateBlue"),
            Self::DarkSlateGray => write!(f, "darkSlateGray"),
            Self::DarkSlateGrey => write!(f, "darkSlateGrey"),
            Self::DarkTurquoise => write!(f, "darkTurquoise"),
            Self::DarkViolet => write!(f, "darkViolet"),
            Self::DkBlue => write!(f, "dkBlue"),
            Self::DkCyan => write!(f, "dkCyan"),
            Self::DkGoldenrod => write!(f, "dkGoldenrod"),
            Self::DkGray => write!(f, "dkGray"),
            Self::DkGrey => write!(f, "dkGrey"),
            Self::DkGreen => write!(f, "dkGreen"),
            Self::DkKhaki => write!(f, "dkKhaki"),
            Self::DkMagenta => write!(f, "dkMagenta"),
            Self::DkOliveGreen => write!(f, "dkOliveGreen"),
            Self::DkOrange => write!(f, "dkOrange"),
            Self::DkOrchid => write!(f, "dkOrchid"),
            Self::DkRed => write!(f, "dkRed"),
            Self::DkSalmon => write!(f, "dkSalmon"),
            Self::DkSeaGreen => write!(f, "dkSeaGreen"),
            Self::DkSlateBlue => write!(f, "dkSlateBlue"),
            Self::DkSlateGray => write!(f, "dkSlateGray"),
            Self::DkSlateGrey => write!(f, "dkSlateGrey"),
            Self::DkTurquoise => write!(f, "dkTurquoise"),
            Self::DkViolet => write!(f, "dkViolet"),
            Self::DeepPink => write!(f, "deepPink"),
            Self::DeepSkyBlue => write!(f, "deepSkyBlue"),
            Self::DimGray => write!(f, "dimGray"),
            Self::DimGrey => write!(f, "dimGrey"),
            Self::DodgerBlue => write!(f, "dodgerBlue"),
            Self::Firebrick => write!(f, "firebrick"),
            Self::FloralWhite => write!(f, "floralWhite"),
            Self::ForestGreen => write!(f, "forestGreen"),
            Self::Fuchsia => write!(f, "fuchsia"),
            Self::Gainsboro => write!(f, "gainsboro"),
            Self::GhostWhite => write!(f, "ghostWhite"),
            Self::Gold => write!(f, "gold"),
            Self::Goldenrod => write!(f, "goldenrod"),
            Self::Gray => write!(f, "gray"),
            Self::Grey => write!(f, "grey"),
            Self::Green => write!(f, "green"),
            Self::GreenYellow => write!(f, "greenYellow"),
            Self::Honeydew => write!(f, "honeydew"),
            Self::HotPink => write!(f, "hotPink"),
            Self::IndianRed => write!(f, "indianRed"),
            Self::Indigo => write!(f, "indigo"),
            Self::Ivory => write!(f, "ivory"),
            Self::Khaki => write!(f, "khaki"),
            Self::Lavender => write!(f, "lavender"),
            Self::LavenderBlush => write!(f, "lavenderBlush"),
            Self::LawnGreen => write!(f, "lawnGreen"),
            Self::LemonChiffon => write!(f, "lemonChiffon"),
            Self::LightBlue => write!(f, "lightBlue"),
            Self::LightCoral => write!(f, "lightCoral"),
            Self::LightCyan => write!(f, "lightCyan"),
            Self::LightGoldenrodYellow => write!(f, "lightGoldenrodYellow"),
            Self::LightGray => write!(f, "lightGray"),
            Self::LightGrey => write!(f, "lightGrey"),
            Self::LightGreen => write!(f, "lightGreen"),
            Self::LightPink => write!(f, "lightPink"),
            Self::LightSalmon => write!(f, "lightSalmon"),
            Self::LightSeaGreen => write!(f, "lightSeaGreen"),
            Self::LightSkyBlue => write!(f, "lightSkyBlue"),
            Self::LightSlateGray => write!(f, "lightSlateGray"),
            Self::LightSlateGrey => write!(f, "lightSlateGrey"),
            Self::LightSteelBlue => write!(f, "lightSteelBlue"),
            Self::LightYellow => write!(f, "lightYellow"),
            Self::LtBlue => write!(f, "ltBlue"),
            Self::LtCoral => write!(f, "ltCoral"),
            Self::LtCyan => write!(f, "ltCyan"),
            Self::LtGoldenrodYellow => write!(f, "ltGoldenrodYellow"),
            Self::LtGray => write!(f, "ltGray"),
            Self::LtGrey => write!(f, "ltGrey"),
            Self::LtGreen => write!(f, "ltGreen"),
            Self::LtPink => write!(f, "ltPink"),
            Self::LtSalmon => write!(f, "ltSalmon"),
            Self::LtSeaGreen => write!(f, "ltSeaGreen"),
            Self::LtSkyBlue => write!(f, "ltSkyBlue"),
            Self::LtSlateGray => write!(f, "ltSlateGray"),
            Self::LtSlateGrey => write!(f, "ltSlateGrey"),
            Self::LtSteelBlue => write!(f, "ltSteelBlue"),
            Self::LtYellow => write!(f, "ltYellow"),
            Self::Lime => write!(f, "lime"),
            Self::LimeGreen => write!(f, "limeGreen"),
            Self::Linen => write!(f, "linen"),
            Self::Magenta => write!(f, "magenta"),
            Self::Maroon => write!(f, "maroon"),
            Self::MedAquamarine => write!(f, "medAquamarine"),
            Self::MedBlue => write!(f, "medBlue"),
            Self::MedOrchid => write!(f, "medOrchid"),
            Self::MedPurple => write!(f, "medPurple"),
            Self::MedSeaGreen => write!(f, "medSeaGreen"),
            Self::MedSlateBlue => write!(f, "medSlateBlue"),
            Self::MedSpringGreen => write!(f, "medSpringGreen"),
            Self::MedTurquoise => write!(f, "medTurquoise"),
            Self::MedVioletRed => write!(f, "medVioletRed"),
            Self::MediumAquamarine => write!(f, "mediumAquamarine"),
            Self::MediumBlue => write!(f, "mediumBlue"),
            Self::MediumOrchid => write!(f, "mediumOrchid"),
            Self::MediumPurple => write!(f, "mediumPurple"),
            Self::MediumSeaGreen => write!(f, "mediumSeaGreen"),
            Self::MediumSlateBlue => write!(f, "mediumSlateBlue"),
            Self::MediumSpringGreen => write!(f, "mediumSpringGreen"),
            Self::MediumTurquoise => write!(f, "mediumTurquoise"),
            Self::MediumVioletRed => write!(f, "mediumVioletRed"),
            Self::MidnightBlue => write!(f, "midnightBlue"),
            Self::MintCream => write!(f, "mintCream"),
            Self::MistyRose => write!(f, "mistyRose"),
            Self::Moccasin => write!(f, "moccasin"),
            Self::NavajoWhite => write!(f, "navajoWhite"),
            Self::Navy => write!(f, "navy"),
            Self::OldLace => write!(f, "oldLace"),
            Self::Olive => write!(f, "olive"),
            Self::OliveDrab => write!(f, "oliveDrab"),
            Self::Orange => write!(f, "orange"),
            Self::OrangeRed => write!(f, "orangeRed"),
            Self::Orchid => write!(f, "orchid"),
            Self::PaleGoldenrod => write!(f, "paleGoldenrod"),
            Self::PaleGreen => write!(f, "paleGreen"),
            Self::PaleTurquoise => write!(f, "paleTurquoise"),
            Self::PaleVioletRed => write!(f, "paleVioletRed"),
            Self::PapayaWhip => write!(f, "papayaWhip"),
            Self::PeachPuff => write!(f, "peachPuff"),
            Self::Peru => write!(f, "peru"),
            Self::Pink => write!(f, "pink"),
            Self::Plum => write!(f, "plum"),
            Self::PowderBlue => write!(f, "powderBlue"),
            Self::Purple => write!(f, "purple"),
            Self::Red => write!(f, "red"),
            Self::RosyBrown => write!(f, "rosyBrown"),
            Self::RoyalBlue => write!(f, "royalBlue"),
            Self::SaddleBrown => write!(f, "saddleBrown"),
            Self::Salmon => write!(f, "salmon"),
            Self::SandyBrown => write!(f, "sandyBrown"),
            Self::SeaGreen => write!(f, "seaGreen"),
            Self::SeaShell => write!(f, "seaShell"),
            Self::Sienna => write!(f, "sienna"),
            Self::Silver => write!(f, "silver"),
            Self::SkyBlue => write!(f, "skyBlue"),
            Self::SlateBlue => write!(f, "slateBlue"),
            Self::SlateGray => write!(f, "slateGray"),
            Self::SlateGrey => write!(f, "slateGrey"),
            Self::Snow => write!(f, "snow"),
            Self::SpringGreen => write!(f, "springGreen"),
            Self::SteelBlue => write!(f, "steelBlue"),
            Self::Tan => write!(f, "tan"),
            Self::Teal => write!(f, "teal"),
            Self::Thistle => write!(f, "thistle"),
            Self::Tomato => write!(f, "tomato"),
            Self::Turquoise => write!(f, "turquoise"),
            Self::Violet => write!(f, "violet"),
            Self::Wheat => write!(f, "wheat"),
            Self::White => write!(f, "white"),
            Self::WhiteSmoke => write!(f, "whiteSmoke"),
            Self::Yellow => write!(f, "yellow"),
            Self::YellowGreen => write!(f, "yellowGreen"),
        }
    }
}

impl std::str::FromStr for STPresetColorVal {
    type Err = String;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        match s {
            "aliceBlue" => Ok(Self::AliceBlue),
            "antiqueWhite" => Ok(Self::AntiqueWhite),
            "aqua" => Ok(Self::Aqua),
            "aquamarine" => Ok(Self::Aquamarine),
            "azure" => Ok(Self::Azure),
            "beige" => Ok(Self::Beige),
            "bisque" => Ok(Self::Bisque),
            "black" => Ok(Self::Black),
            "blanchedAlmond" => Ok(Self::BlanchedAlmond),
            "blue" => Ok(Self::Blue),
            "blueViolet" => Ok(Self::BlueViolet),
            "brown" => Ok(Self::Brown),
            "burlyWood" => Ok(Self::BurlyWood),
            "cadetBlue" => Ok(Self::CadetBlue),
            "chartreuse" => Ok(Self::Chartreuse),
            "chocolate" => Ok(Self::Chocolate),
            "coral" => Ok(Self::Coral),
            "cornflowerBlue" => Ok(Self::CornflowerBlue),
            "cornsilk" => Ok(Self::Cornsilk),
            "crimson" => Ok(Self::Crimson),
            "cyan" => Ok(Self::Cyan),
            "darkBlue" => Ok(Self::DarkBlue),
            "darkCyan" => Ok(Self::DarkCyan),
            "darkGoldenrod" => Ok(Self::DarkGoldenrod),
            "darkGray" => Ok(Self::DarkGray),
            "darkGrey" => Ok(Self::DarkGrey),
            "darkGreen" => Ok(Self::DarkGreen),
            "darkKhaki" => Ok(Self::DarkKhaki),
            "darkMagenta" => Ok(Self::DarkMagenta),
            "darkOliveGreen" => Ok(Self::DarkOliveGreen),
            "darkOrange" => Ok(Self::DarkOrange),
            "darkOrchid" => Ok(Self::DarkOrchid),
            "darkRed" => Ok(Self::DarkRed),
            "darkSalmon" => Ok(Self::DarkSalmon),
            "darkSeaGreen" => Ok(Self::DarkSeaGreen),
            "darkSlateBlue" => Ok(Self::DarkSlateBlue),
            "darkSlateGray" => Ok(Self::DarkSlateGray),
            "darkSlateGrey" => Ok(Self::DarkSlateGrey),
            "darkTurquoise" => Ok(Self::DarkTurquoise),
            "darkViolet" => Ok(Self::DarkViolet),
            "dkBlue" => Ok(Self::DkBlue),
            "dkCyan" => Ok(Self::DkCyan),
            "dkGoldenrod" => Ok(Self::DkGoldenrod),
            "dkGray" => Ok(Self::DkGray),
            "dkGrey" => Ok(Self::DkGrey),
            "dkGreen" => Ok(Self::DkGreen),
            "dkKhaki" => Ok(Self::DkKhaki),
            "dkMagenta" => Ok(Self::DkMagenta),
            "dkOliveGreen" => Ok(Self::DkOliveGreen),
            "dkOrange" => Ok(Self::DkOrange),
            "dkOrchid" => Ok(Self::DkOrchid),
            "dkRed" => Ok(Self::DkRed),
            "dkSalmon" => Ok(Self::DkSalmon),
            "dkSeaGreen" => Ok(Self::DkSeaGreen),
            "dkSlateBlue" => Ok(Self::DkSlateBlue),
            "dkSlateGray" => Ok(Self::DkSlateGray),
            "dkSlateGrey" => Ok(Self::DkSlateGrey),
            "dkTurquoise" => Ok(Self::DkTurquoise),
            "dkViolet" => Ok(Self::DkViolet),
            "deepPink" => Ok(Self::DeepPink),
            "deepSkyBlue" => Ok(Self::DeepSkyBlue),
            "dimGray" => Ok(Self::DimGray),
            "dimGrey" => Ok(Self::DimGrey),
            "dodgerBlue" => Ok(Self::DodgerBlue),
            "firebrick" => Ok(Self::Firebrick),
            "floralWhite" => Ok(Self::FloralWhite),
            "forestGreen" => Ok(Self::ForestGreen),
            "fuchsia" => Ok(Self::Fuchsia),
            "gainsboro" => Ok(Self::Gainsboro),
            "ghostWhite" => Ok(Self::GhostWhite),
            "gold" => Ok(Self::Gold),
            "goldenrod" => Ok(Self::Goldenrod),
            "gray" => Ok(Self::Gray),
            "grey" => Ok(Self::Grey),
            "green" => Ok(Self::Green),
            "greenYellow" => Ok(Self::GreenYellow),
            "honeydew" => Ok(Self::Honeydew),
            "hotPink" => Ok(Self::HotPink),
            "indianRed" => Ok(Self::IndianRed),
            "indigo" => Ok(Self::Indigo),
            "ivory" => Ok(Self::Ivory),
            "khaki" => Ok(Self::Khaki),
            "lavender" => Ok(Self::Lavender),
            "lavenderBlush" => Ok(Self::LavenderBlush),
            "lawnGreen" => Ok(Self::LawnGreen),
            "lemonChiffon" => Ok(Self::LemonChiffon),
            "lightBlue" => Ok(Self::LightBlue),
            "lightCoral" => Ok(Self::LightCoral),
            "lightCyan" => Ok(Self::LightCyan),
            "lightGoldenrodYellow" => Ok(Self::LightGoldenrodYellow),
            "lightGray" => Ok(Self::LightGray),
            "lightGrey" => Ok(Self::LightGrey),
            "lightGreen" => Ok(Self::LightGreen),
            "lightPink" => Ok(Self::LightPink),
            "lightSalmon" => Ok(Self::LightSalmon),
            "lightSeaGreen" => Ok(Self::LightSeaGreen),
            "lightSkyBlue" => Ok(Self::LightSkyBlue),
            "lightSlateGray" => Ok(Self::LightSlateGray),
            "lightSlateGrey" => Ok(Self::LightSlateGrey),
            "lightSteelBlue" => Ok(Self::LightSteelBlue),
            "lightYellow" => Ok(Self::LightYellow),
            "ltBlue" => Ok(Self::LtBlue),
            "ltCoral" => Ok(Self::LtCoral),
            "ltCyan" => Ok(Self::LtCyan),
            "ltGoldenrodYellow" => Ok(Self::LtGoldenrodYellow),
            "ltGray" => Ok(Self::LtGray),
            "ltGrey" => Ok(Self::LtGrey),
            "ltGreen" => Ok(Self::LtGreen),
            "ltPink" => Ok(Self::LtPink),
            "ltSalmon" => Ok(Self::LtSalmon),
            "ltSeaGreen" => Ok(Self::LtSeaGreen),
            "ltSkyBlue" => Ok(Self::LtSkyBlue),
            "ltSlateGray" => Ok(Self::LtSlateGray),
            "ltSlateGrey" => Ok(Self::LtSlateGrey),
            "ltSteelBlue" => Ok(Self::LtSteelBlue),
            "ltYellow" => Ok(Self::LtYellow),
            "lime" => Ok(Self::Lime),
            "limeGreen" => Ok(Self::LimeGreen),
            "linen" => Ok(Self::Linen),
            "magenta" => Ok(Self::Magenta),
            "maroon" => Ok(Self::Maroon),
            "medAquamarine" => Ok(Self::MedAquamarine),
            "medBlue" => Ok(Self::MedBlue),
            "medOrchid" => Ok(Self::MedOrchid),
            "medPurple" => Ok(Self::MedPurple),
            "medSeaGreen" => Ok(Self::MedSeaGreen),
            "medSlateBlue" => Ok(Self::MedSlateBlue),
            "medSpringGreen" => Ok(Self::MedSpringGreen),
            "medTurquoise" => Ok(Self::MedTurquoise),
            "medVioletRed" => Ok(Self::MedVioletRed),
            "mediumAquamarine" => Ok(Self::MediumAquamarine),
            "mediumBlue" => Ok(Self::MediumBlue),
            "mediumOrchid" => Ok(Self::MediumOrchid),
            "mediumPurple" => Ok(Self::MediumPurple),
            "mediumSeaGreen" => Ok(Self::MediumSeaGreen),
            "mediumSlateBlue" => Ok(Self::MediumSlateBlue),
            "mediumSpringGreen" => Ok(Self::MediumSpringGreen),
            "mediumTurquoise" => Ok(Self::MediumTurquoise),
            "mediumVioletRed" => Ok(Self::MediumVioletRed),
            "midnightBlue" => Ok(Self::MidnightBlue),
            "mintCream" => Ok(Self::MintCream),
            "mistyRose" => Ok(Self::MistyRose),
            "moccasin" => Ok(Self::Moccasin),
            "navajoWhite" => Ok(Self::NavajoWhite),
            "navy" => Ok(Self::Navy),
            "oldLace" => Ok(Self::OldLace),
            "olive" => Ok(Self::Olive),
            "oliveDrab" => Ok(Self::OliveDrab),
            "orange" => Ok(Self::Orange),
            "orangeRed" => Ok(Self::OrangeRed),
            "orchid" => Ok(Self::Orchid),
            "paleGoldenrod" => Ok(Self::PaleGoldenrod),
            "paleGreen" => Ok(Self::PaleGreen),
            "paleTurquoise" => Ok(Self::PaleTurquoise),
            "paleVioletRed" => Ok(Self::PaleVioletRed),
            "papayaWhip" => Ok(Self::PapayaWhip),
            "peachPuff" => Ok(Self::PeachPuff),
            "peru" => Ok(Self::Peru),
            "pink" => Ok(Self::Pink),
            "plum" => Ok(Self::Plum),
            "powderBlue" => Ok(Self::PowderBlue),
            "purple" => Ok(Self::Purple),
            "red" => Ok(Self::Red),
            "rosyBrown" => Ok(Self::RosyBrown),
            "royalBlue" => Ok(Self::RoyalBlue),
            "saddleBrown" => Ok(Self::SaddleBrown),
            "salmon" => Ok(Self::Salmon),
            "sandyBrown" => Ok(Self::SandyBrown),
            "seaGreen" => Ok(Self::SeaGreen),
            "seaShell" => Ok(Self::SeaShell),
            "sienna" => Ok(Self::Sienna),
            "silver" => Ok(Self::Silver),
            "skyBlue" => Ok(Self::SkyBlue),
            "slateBlue" => Ok(Self::SlateBlue),
            "slateGray" => Ok(Self::SlateGray),
            "slateGrey" => Ok(Self::SlateGrey),
            "snow" => Ok(Self::Snow),
            "springGreen" => Ok(Self::SpringGreen),
            "steelBlue" => Ok(Self::SteelBlue),
            "tan" => Ok(Self::Tan),
            "teal" => Ok(Self::Teal),
            "thistle" => Ok(Self::Thistle),
            "tomato" => Ok(Self::Tomato),
            "turquoise" => Ok(Self::Turquoise),
            "violet" => Ok(Self::Violet),
            "wheat" => Ok(Self::Wheat),
            "white" => Ok(Self::White),
            "whiteSmoke" => Ok(Self::WhiteSmoke),
            "yellow" => Ok(Self::Yellow),
            "yellowGreen" => Ok(Self::YellowGreen),
            _ => Err(format!("unknown STPresetColorVal value: {}", s)),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum STRectAlignment {
    #[serde(rename = "tl")]
    Tl,
    #[serde(rename = "t")]
    T,
    #[serde(rename = "tr")]
    Tr,
    #[serde(rename = "l")]
    L,
    #[serde(rename = "ctr")]
    Ctr,
    #[serde(rename = "r")]
    R,
    #[serde(rename = "bl")]
    Bl,
    #[serde(rename = "b")]
    B,
    #[serde(rename = "br")]
    Br,
}

impl std::fmt::Display for STRectAlignment {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::Tl => write!(f, "tl"),
            Self::T => write!(f, "t"),
            Self::Tr => write!(f, "tr"),
            Self::L => write!(f, "l"),
            Self::Ctr => write!(f, "ctr"),
            Self::R => write!(f, "r"),
            Self::Bl => write!(f, "bl"),
            Self::B => write!(f, "b"),
            Self::Br => write!(f, "br"),
        }
    }
}

impl std::str::FromStr for STRectAlignment {
    type Err = String;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        match s {
            "tl" => Ok(Self::Tl),
            "t" => Ok(Self::T),
            "tr" => Ok(Self::Tr),
            "l" => Ok(Self::L),
            "ctr" => Ok(Self::Ctr),
            "r" => Ok(Self::R),
            "bl" => Ok(Self::Bl),
            "b" => Ok(Self::B),
            "br" => Ok(Self::Br),
            _ => Err(format!("unknown STRectAlignment value: {}", s)),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum STBlackWhiteMode {
    #[serde(rename = "clr")]
    Clr,
    #[serde(rename = "auto")]
    Auto,
    #[serde(rename = "gray")]
    Gray,
    #[serde(rename = "ltGray")]
    LtGray,
    #[serde(rename = "invGray")]
    InvGray,
    #[serde(rename = "grayWhite")]
    GrayWhite,
    #[serde(rename = "blackGray")]
    BlackGray,
    #[serde(rename = "blackWhite")]
    BlackWhite,
    #[serde(rename = "black")]
    Black,
    #[serde(rename = "white")]
    White,
    #[serde(rename = "hidden")]
    Hidden,
}

impl std::fmt::Display for STBlackWhiteMode {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::Clr => write!(f, "clr"),
            Self::Auto => write!(f, "auto"),
            Self::Gray => write!(f, "gray"),
            Self::LtGray => write!(f, "ltGray"),
            Self::InvGray => write!(f, "invGray"),
            Self::GrayWhite => write!(f, "grayWhite"),
            Self::BlackGray => write!(f, "blackGray"),
            Self::BlackWhite => write!(f, "blackWhite"),
            Self::Black => write!(f, "black"),
            Self::White => write!(f, "white"),
            Self::Hidden => write!(f, "hidden"),
        }
    }
}

impl std::str::FromStr for STBlackWhiteMode {
    type Err = String;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        match s {
            "clr" => Ok(Self::Clr),
            "auto" => Ok(Self::Auto),
            "gray" => Ok(Self::Gray),
            "ltGray" => Ok(Self::LtGray),
            "invGray" => Ok(Self::InvGray),
            "grayWhite" => Ok(Self::GrayWhite),
            "blackGray" => Ok(Self::BlackGray),
            "blackWhite" => Ok(Self::BlackWhite),
            "black" => Ok(Self::Black),
            "white" => Ok(Self::White),
            "hidden" => Ok(Self::Hidden),
            _ => Err(format!("unknown STBlackWhiteMode value: {}", s)),
        }
    }
}

pub type STDrawingElementId = u32;

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum STChartBuildStep {
    #[serde(rename = "category")]
    Category,
    #[serde(rename = "ptInCategory")]
    PtInCategory,
    #[serde(rename = "series")]
    Series,
    #[serde(rename = "ptInSeries")]
    PtInSeries,
    #[serde(rename = "allPts")]
    AllPts,
    #[serde(rename = "gridLegend")]
    GridLegend,
}

impl std::fmt::Display for STChartBuildStep {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::Category => write!(f, "category"),
            Self::PtInCategory => write!(f, "ptInCategory"),
            Self::Series => write!(f, "series"),
            Self::PtInSeries => write!(f, "ptInSeries"),
            Self::AllPts => write!(f, "allPts"),
            Self::GridLegend => write!(f, "gridLegend"),
        }
    }
}

impl std::str::FromStr for STChartBuildStep {
    type Err = String;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        match s {
            "category" => Ok(Self::Category),
            "ptInCategory" => Ok(Self::PtInCategory),
            "series" => Ok(Self::Series),
            "ptInSeries" => Ok(Self::PtInSeries),
            "allPts" => Ok(Self::AllPts),
            "gridLegend" => Ok(Self::GridLegend),
            _ => Err(format!("unknown STChartBuildStep value: {}", s)),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum STDgmBuildStep {
    #[serde(rename = "sp")]
    Sp,
    #[serde(rename = "bg")]
    Bg,
}

impl std::fmt::Display for STDgmBuildStep {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::Sp => write!(f, "sp"),
            Self::Bg => write!(f, "bg"),
        }
    }
}

impl std::str::FromStr for STDgmBuildStep {
    type Err = String;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        match s {
            "sp" => Ok(Self::Sp),
            "bg" => Ok(Self::Bg),
            _ => Err(format!("unknown STDgmBuildStep value: {}", s)),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum STAnimationDgmOnlyBuildType {
    #[serde(rename = "one")]
    One,
    #[serde(rename = "lvlOne")]
    LvlOne,
    #[serde(rename = "lvlAtOnce")]
    LvlAtOnce,
}

impl std::fmt::Display for STAnimationDgmOnlyBuildType {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::One => write!(f, "one"),
            Self::LvlOne => write!(f, "lvlOne"),
            Self::LvlAtOnce => write!(f, "lvlAtOnce"),
        }
    }
}

impl std::str::FromStr for STAnimationDgmOnlyBuildType {
    type Err = String;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        match s {
            "one" => Ok(Self::One),
            "lvlOne" => Ok(Self::LvlOne),
            "lvlAtOnce" => Ok(Self::LvlAtOnce),
            _ => Err(format!("unknown STAnimationDgmOnlyBuildType value: {}", s)),
        }
    }
}

pub type STAnimationDgmBuildType = String;

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum STAnimationChartOnlyBuildType {
    #[serde(rename = "series")]
    Series,
    #[serde(rename = "category")]
    Category,
    #[serde(rename = "seriesEl")]
    SeriesEl,
    #[serde(rename = "categoryEl")]
    CategoryEl,
}

impl std::fmt::Display for STAnimationChartOnlyBuildType {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::Series => write!(f, "series"),
            Self::Category => write!(f, "category"),
            Self::SeriesEl => write!(f, "seriesEl"),
            Self::CategoryEl => write!(f, "categoryEl"),
        }
    }
}

impl std::str::FromStr for STAnimationChartOnlyBuildType {
    type Err = String;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        match s {
            "series" => Ok(Self::Series),
            "category" => Ok(Self::Category),
            "seriesEl" => Ok(Self::SeriesEl),
            "categoryEl" => Ok(Self::CategoryEl),
            _ => Err(format!(
                "unknown STAnimationChartOnlyBuildType value: {}",
                s
            )),
        }
    }
}

pub type STAnimationChartBuildType = String;

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum STPresetCameraType {
    #[serde(rename = "legacyObliqueTopLeft")]
    LegacyObliqueTopLeft,
    #[serde(rename = "legacyObliqueTop")]
    LegacyObliqueTop,
    #[serde(rename = "legacyObliqueTopRight")]
    LegacyObliqueTopRight,
    #[serde(rename = "legacyObliqueLeft")]
    LegacyObliqueLeft,
    #[serde(rename = "legacyObliqueFront")]
    LegacyObliqueFront,
    #[serde(rename = "legacyObliqueRight")]
    LegacyObliqueRight,
    #[serde(rename = "legacyObliqueBottomLeft")]
    LegacyObliqueBottomLeft,
    #[serde(rename = "legacyObliqueBottom")]
    LegacyObliqueBottom,
    #[serde(rename = "legacyObliqueBottomRight")]
    LegacyObliqueBottomRight,
    #[serde(rename = "legacyPerspectiveTopLeft")]
    LegacyPerspectiveTopLeft,
    #[serde(rename = "legacyPerspectiveTop")]
    LegacyPerspectiveTop,
    #[serde(rename = "legacyPerspectiveTopRight")]
    LegacyPerspectiveTopRight,
    #[serde(rename = "legacyPerspectiveLeft")]
    LegacyPerspectiveLeft,
    #[serde(rename = "legacyPerspectiveFront")]
    LegacyPerspectiveFront,
    #[serde(rename = "legacyPerspectiveRight")]
    LegacyPerspectiveRight,
    #[serde(rename = "legacyPerspectiveBottomLeft")]
    LegacyPerspectiveBottomLeft,
    #[serde(rename = "legacyPerspectiveBottom")]
    LegacyPerspectiveBottom,
    #[serde(rename = "legacyPerspectiveBottomRight")]
    LegacyPerspectiveBottomRight,
    #[serde(rename = "orthographicFront")]
    OrthographicFront,
    #[serde(rename = "isometricTopUp")]
    IsometricTopUp,
    #[serde(rename = "isometricTopDown")]
    IsometricTopDown,
    #[serde(rename = "isometricBottomUp")]
    IsometricBottomUp,
    #[serde(rename = "isometricBottomDown")]
    IsometricBottomDown,
    #[serde(rename = "isometricLeftUp")]
    IsometricLeftUp,
    #[serde(rename = "isometricLeftDown")]
    IsometricLeftDown,
    #[serde(rename = "isometricRightUp")]
    IsometricRightUp,
    #[serde(rename = "isometricRightDown")]
    IsometricRightDown,
    #[serde(rename = "isometricOffAxis1Left")]
    IsometricOffAxis1Left,
    #[serde(rename = "isometricOffAxis1Right")]
    IsometricOffAxis1Right,
    #[serde(rename = "isometricOffAxis1Top")]
    IsometricOffAxis1Top,
    #[serde(rename = "isometricOffAxis2Left")]
    IsometricOffAxis2Left,
    #[serde(rename = "isometricOffAxis2Right")]
    IsometricOffAxis2Right,
    #[serde(rename = "isometricOffAxis2Top")]
    IsometricOffAxis2Top,
    #[serde(rename = "isometricOffAxis3Left")]
    IsometricOffAxis3Left,
    #[serde(rename = "isometricOffAxis3Right")]
    IsometricOffAxis3Right,
    #[serde(rename = "isometricOffAxis3Bottom")]
    IsometricOffAxis3Bottom,
    #[serde(rename = "isometricOffAxis4Left")]
    IsometricOffAxis4Left,
    #[serde(rename = "isometricOffAxis4Right")]
    IsometricOffAxis4Right,
    #[serde(rename = "isometricOffAxis4Bottom")]
    IsometricOffAxis4Bottom,
    #[serde(rename = "obliqueTopLeft")]
    ObliqueTopLeft,
    #[serde(rename = "obliqueTop")]
    ObliqueTop,
    #[serde(rename = "obliqueTopRight")]
    ObliqueTopRight,
    #[serde(rename = "obliqueLeft")]
    ObliqueLeft,
    #[serde(rename = "obliqueRight")]
    ObliqueRight,
    #[serde(rename = "obliqueBottomLeft")]
    ObliqueBottomLeft,
    #[serde(rename = "obliqueBottom")]
    ObliqueBottom,
    #[serde(rename = "obliqueBottomRight")]
    ObliqueBottomRight,
    #[serde(rename = "perspectiveFront")]
    PerspectiveFront,
    #[serde(rename = "perspectiveLeft")]
    PerspectiveLeft,
    #[serde(rename = "perspectiveRight")]
    PerspectiveRight,
    #[serde(rename = "perspectiveAbove")]
    PerspectiveAbove,
    #[serde(rename = "perspectiveBelow")]
    PerspectiveBelow,
    #[serde(rename = "perspectiveAboveLeftFacing")]
    PerspectiveAboveLeftFacing,
    #[serde(rename = "perspectiveAboveRightFacing")]
    PerspectiveAboveRightFacing,
    #[serde(rename = "perspectiveContrastingLeftFacing")]
    PerspectiveContrastingLeftFacing,
    #[serde(rename = "perspectiveContrastingRightFacing")]
    PerspectiveContrastingRightFacing,
    #[serde(rename = "perspectiveHeroicLeftFacing")]
    PerspectiveHeroicLeftFacing,
    #[serde(rename = "perspectiveHeroicRightFacing")]
    PerspectiveHeroicRightFacing,
    #[serde(rename = "perspectiveHeroicExtremeLeftFacing")]
    PerspectiveHeroicExtremeLeftFacing,
    #[serde(rename = "perspectiveHeroicExtremeRightFacing")]
    PerspectiveHeroicExtremeRightFacing,
    #[serde(rename = "perspectiveRelaxed")]
    PerspectiveRelaxed,
    #[serde(rename = "perspectiveRelaxedModerately")]
    PerspectiveRelaxedModerately,
}

impl std::fmt::Display for STPresetCameraType {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::LegacyObliqueTopLeft => write!(f, "legacyObliqueTopLeft"),
            Self::LegacyObliqueTop => write!(f, "legacyObliqueTop"),
            Self::LegacyObliqueTopRight => write!(f, "legacyObliqueTopRight"),
            Self::LegacyObliqueLeft => write!(f, "legacyObliqueLeft"),
            Self::LegacyObliqueFront => write!(f, "legacyObliqueFront"),
            Self::LegacyObliqueRight => write!(f, "legacyObliqueRight"),
            Self::LegacyObliqueBottomLeft => write!(f, "legacyObliqueBottomLeft"),
            Self::LegacyObliqueBottom => write!(f, "legacyObliqueBottom"),
            Self::LegacyObliqueBottomRight => write!(f, "legacyObliqueBottomRight"),
            Self::LegacyPerspectiveTopLeft => write!(f, "legacyPerspectiveTopLeft"),
            Self::LegacyPerspectiveTop => write!(f, "legacyPerspectiveTop"),
            Self::LegacyPerspectiveTopRight => write!(f, "legacyPerspectiveTopRight"),
            Self::LegacyPerspectiveLeft => write!(f, "legacyPerspectiveLeft"),
            Self::LegacyPerspectiveFront => write!(f, "legacyPerspectiveFront"),
            Self::LegacyPerspectiveRight => write!(f, "legacyPerspectiveRight"),
            Self::LegacyPerspectiveBottomLeft => write!(f, "legacyPerspectiveBottomLeft"),
            Self::LegacyPerspectiveBottom => write!(f, "legacyPerspectiveBottom"),
            Self::LegacyPerspectiveBottomRight => write!(f, "legacyPerspectiveBottomRight"),
            Self::OrthographicFront => write!(f, "orthographicFront"),
            Self::IsometricTopUp => write!(f, "isometricTopUp"),
            Self::IsometricTopDown => write!(f, "isometricTopDown"),
            Self::IsometricBottomUp => write!(f, "isometricBottomUp"),
            Self::IsometricBottomDown => write!(f, "isometricBottomDown"),
            Self::IsometricLeftUp => write!(f, "isometricLeftUp"),
            Self::IsometricLeftDown => write!(f, "isometricLeftDown"),
            Self::IsometricRightUp => write!(f, "isometricRightUp"),
            Self::IsometricRightDown => write!(f, "isometricRightDown"),
            Self::IsometricOffAxis1Left => write!(f, "isometricOffAxis1Left"),
            Self::IsometricOffAxis1Right => write!(f, "isometricOffAxis1Right"),
            Self::IsometricOffAxis1Top => write!(f, "isometricOffAxis1Top"),
            Self::IsometricOffAxis2Left => write!(f, "isometricOffAxis2Left"),
            Self::IsometricOffAxis2Right => write!(f, "isometricOffAxis2Right"),
            Self::IsometricOffAxis2Top => write!(f, "isometricOffAxis2Top"),
            Self::IsometricOffAxis3Left => write!(f, "isometricOffAxis3Left"),
            Self::IsometricOffAxis3Right => write!(f, "isometricOffAxis3Right"),
            Self::IsometricOffAxis3Bottom => write!(f, "isometricOffAxis3Bottom"),
            Self::IsometricOffAxis4Left => write!(f, "isometricOffAxis4Left"),
            Self::IsometricOffAxis4Right => write!(f, "isometricOffAxis4Right"),
            Self::IsometricOffAxis4Bottom => write!(f, "isometricOffAxis4Bottom"),
            Self::ObliqueTopLeft => write!(f, "obliqueTopLeft"),
            Self::ObliqueTop => write!(f, "obliqueTop"),
            Self::ObliqueTopRight => write!(f, "obliqueTopRight"),
            Self::ObliqueLeft => write!(f, "obliqueLeft"),
            Self::ObliqueRight => write!(f, "obliqueRight"),
            Self::ObliqueBottomLeft => write!(f, "obliqueBottomLeft"),
            Self::ObliqueBottom => write!(f, "obliqueBottom"),
            Self::ObliqueBottomRight => write!(f, "obliqueBottomRight"),
            Self::PerspectiveFront => write!(f, "perspectiveFront"),
            Self::PerspectiveLeft => write!(f, "perspectiveLeft"),
            Self::PerspectiveRight => write!(f, "perspectiveRight"),
            Self::PerspectiveAbove => write!(f, "perspectiveAbove"),
            Self::PerspectiveBelow => write!(f, "perspectiveBelow"),
            Self::PerspectiveAboveLeftFacing => write!(f, "perspectiveAboveLeftFacing"),
            Self::PerspectiveAboveRightFacing => write!(f, "perspectiveAboveRightFacing"),
            Self::PerspectiveContrastingLeftFacing => write!(f, "perspectiveContrastingLeftFacing"),
            Self::PerspectiveContrastingRightFacing => {
                write!(f, "perspectiveContrastingRightFacing")
            }
            Self::PerspectiveHeroicLeftFacing => write!(f, "perspectiveHeroicLeftFacing"),
            Self::PerspectiveHeroicRightFacing => write!(f, "perspectiveHeroicRightFacing"),
            Self::PerspectiveHeroicExtremeLeftFacing => {
                write!(f, "perspectiveHeroicExtremeLeftFacing")
            }
            Self::PerspectiveHeroicExtremeRightFacing => {
                write!(f, "perspectiveHeroicExtremeRightFacing")
            }
            Self::PerspectiveRelaxed => write!(f, "perspectiveRelaxed"),
            Self::PerspectiveRelaxedModerately => write!(f, "perspectiveRelaxedModerately"),
        }
    }
}

impl std::str::FromStr for STPresetCameraType {
    type Err = String;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        match s {
            "legacyObliqueTopLeft" => Ok(Self::LegacyObliqueTopLeft),
            "legacyObliqueTop" => Ok(Self::LegacyObliqueTop),
            "legacyObliqueTopRight" => Ok(Self::LegacyObliqueTopRight),
            "legacyObliqueLeft" => Ok(Self::LegacyObliqueLeft),
            "legacyObliqueFront" => Ok(Self::LegacyObliqueFront),
            "legacyObliqueRight" => Ok(Self::LegacyObliqueRight),
            "legacyObliqueBottomLeft" => Ok(Self::LegacyObliqueBottomLeft),
            "legacyObliqueBottom" => Ok(Self::LegacyObliqueBottom),
            "legacyObliqueBottomRight" => Ok(Self::LegacyObliqueBottomRight),
            "legacyPerspectiveTopLeft" => Ok(Self::LegacyPerspectiveTopLeft),
            "legacyPerspectiveTop" => Ok(Self::LegacyPerspectiveTop),
            "legacyPerspectiveTopRight" => Ok(Self::LegacyPerspectiveTopRight),
            "legacyPerspectiveLeft" => Ok(Self::LegacyPerspectiveLeft),
            "legacyPerspectiveFront" => Ok(Self::LegacyPerspectiveFront),
            "legacyPerspectiveRight" => Ok(Self::LegacyPerspectiveRight),
            "legacyPerspectiveBottomLeft" => Ok(Self::LegacyPerspectiveBottomLeft),
            "legacyPerspectiveBottom" => Ok(Self::LegacyPerspectiveBottom),
            "legacyPerspectiveBottomRight" => Ok(Self::LegacyPerspectiveBottomRight),
            "orthographicFront" => Ok(Self::OrthographicFront),
            "isometricTopUp" => Ok(Self::IsometricTopUp),
            "isometricTopDown" => Ok(Self::IsometricTopDown),
            "isometricBottomUp" => Ok(Self::IsometricBottomUp),
            "isometricBottomDown" => Ok(Self::IsometricBottomDown),
            "isometricLeftUp" => Ok(Self::IsometricLeftUp),
            "isometricLeftDown" => Ok(Self::IsometricLeftDown),
            "isometricRightUp" => Ok(Self::IsometricRightUp),
            "isometricRightDown" => Ok(Self::IsometricRightDown),
            "isometricOffAxis1Left" => Ok(Self::IsometricOffAxis1Left),
            "isometricOffAxis1Right" => Ok(Self::IsometricOffAxis1Right),
            "isometricOffAxis1Top" => Ok(Self::IsometricOffAxis1Top),
            "isometricOffAxis2Left" => Ok(Self::IsometricOffAxis2Left),
            "isometricOffAxis2Right" => Ok(Self::IsometricOffAxis2Right),
            "isometricOffAxis2Top" => Ok(Self::IsometricOffAxis2Top),
            "isometricOffAxis3Left" => Ok(Self::IsometricOffAxis3Left),
            "isometricOffAxis3Right" => Ok(Self::IsometricOffAxis3Right),
            "isometricOffAxis3Bottom" => Ok(Self::IsometricOffAxis3Bottom),
            "isometricOffAxis4Left" => Ok(Self::IsometricOffAxis4Left),
            "isometricOffAxis4Right" => Ok(Self::IsometricOffAxis4Right),
            "isometricOffAxis4Bottom" => Ok(Self::IsometricOffAxis4Bottom),
            "obliqueTopLeft" => Ok(Self::ObliqueTopLeft),
            "obliqueTop" => Ok(Self::ObliqueTop),
            "obliqueTopRight" => Ok(Self::ObliqueTopRight),
            "obliqueLeft" => Ok(Self::ObliqueLeft),
            "obliqueRight" => Ok(Self::ObliqueRight),
            "obliqueBottomLeft" => Ok(Self::ObliqueBottomLeft),
            "obliqueBottom" => Ok(Self::ObliqueBottom),
            "obliqueBottomRight" => Ok(Self::ObliqueBottomRight),
            "perspectiveFront" => Ok(Self::PerspectiveFront),
            "perspectiveLeft" => Ok(Self::PerspectiveLeft),
            "perspectiveRight" => Ok(Self::PerspectiveRight),
            "perspectiveAbove" => Ok(Self::PerspectiveAbove),
            "perspectiveBelow" => Ok(Self::PerspectiveBelow),
            "perspectiveAboveLeftFacing" => Ok(Self::PerspectiveAboveLeftFacing),
            "perspectiveAboveRightFacing" => Ok(Self::PerspectiveAboveRightFacing),
            "perspectiveContrastingLeftFacing" => Ok(Self::PerspectiveContrastingLeftFacing),
            "perspectiveContrastingRightFacing" => Ok(Self::PerspectiveContrastingRightFacing),
            "perspectiveHeroicLeftFacing" => Ok(Self::PerspectiveHeroicLeftFacing),
            "perspectiveHeroicRightFacing" => Ok(Self::PerspectiveHeroicRightFacing),
            "perspectiveHeroicExtremeLeftFacing" => Ok(Self::PerspectiveHeroicExtremeLeftFacing),
            "perspectiveHeroicExtremeRightFacing" => Ok(Self::PerspectiveHeroicExtremeRightFacing),
            "perspectiveRelaxed" => Ok(Self::PerspectiveRelaxed),
            "perspectiveRelaxedModerately" => Ok(Self::PerspectiveRelaxedModerately),
            _ => Err(format!("unknown STPresetCameraType value: {}", s)),
        }
    }
}

pub type STFOVAngle = i32;

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum STLightRigDirection {
    #[serde(rename = "tl")]
    Tl,
    #[serde(rename = "t")]
    T,
    #[serde(rename = "tr")]
    Tr,
    #[serde(rename = "l")]
    L,
    #[serde(rename = "r")]
    R,
    #[serde(rename = "bl")]
    Bl,
    #[serde(rename = "b")]
    B,
    #[serde(rename = "br")]
    Br,
}

impl std::fmt::Display for STLightRigDirection {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::Tl => write!(f, "tl"),
            Self::T => write!(f, "t"),
            Self::Tr => write!(f, "tr"),
            Self::L => write!(f, "l"),
            Self::R => write!(f, "r"),
            Self::Bl => write!(f, "bl"),
            Self::B => write!(f, "b"),
            Self::Br => write!(f, "br"),
        }
    }
}

impl std::str::FromStr for STLightRigDirection {
    type Err = String;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        match s {
            "tl" => Ok(Self::Tl),
            "t" => Ok(Self::T),
            "tr" => Ok(Self::Tr),
            "l" => Ok(Self::L),
            "r" => Ok(Self::R),
            "bl" => Ok(Self::Bl),
            "b" => Ok(Self::B),
            "br" => Ok(Self::Br),
            _ => Err(format!("unknown STLightRigDirection value: {}", s)),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum STLightRigType {
    #[serde(rename = "legacyFlat1")]
    LegacyFlat1,
    #[serde(rename = "legacyFlat2")]
    LegacyFlat2,
    #[serde(rename = "legacyFlat3")]
    LegacyFlat3,
    #[serde(rename = "legacyFlat4")]
    LegacyFlat4,
    #[serde(rename = "legacyNormal1")]
    LegacyNormal1,
    #[serde(rename = "legacyNormal2")]
    LegacyNormal2,
    #[serde(rename = "legacyNormal3")]
    LegacyNormal3,
    #[serde(rename = "legacyNormal4")]
    LegacyNormal4,
    #[serde(rename = "legacyHarsh1")]
    LegacyHarsh1,
    #[serde(rename = "legacyHarsh2")]
    LegacyHarsh2,
    #[serde(rename = "legacyHarsh3")]
    LegacyHarsh3,
    #[serde(rename = "legacyHarsh4")]
    LegacyHarsh4,
    #[serde(rename = "threePt")]
    ThreePt,
    #[serde(rename = "balanced")]
    Balanced,
    #[serde(rename = "soft")]
    Soft,
    #[serde(rename = "harsh")]
    Harsh,
    #[serde(rename = "flood")]
    Flood,
    #[serde(rename = "contrasting")]
    Contrasting,
    #[serde(rename = "morning")]
    Morning,
    #[serde(rename = "sunrise")]
    Sunrise,
    #[serde(rename = "sunset")]
    Sunset,
    #[serde(rename = "chilly")]
    Chilly,
    #[serde(rename = "freezing")]
    Freezing,
    #[serde(rename = "flat")]
    Flat,
    #[serde(rename = "twoPt")]
    TwoPt,
    #[serde(rename = "glow")]
    Glow,
    #[serde(rename = "brightRoom")]
    BrightRoom,
}

impl std::fmt::Display for STLightRigType {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::LegacyFlat1 => write!(f, "legacyFlat1"),
            Self::LegacyFlat2 => write!(f, "legacyFlat2"),
            Self::LegacyFlat3 => write!(f, "legacyFlat3"),
            Self::LegacyFlat4 => write!(f, "legacyFlat4"),
            Self::LegacyNormal1 => write!(f, "legacyNormal1"),
            Self::LegacyNormal2 => write!(f, "legacyNormal2"),
            Self::LegacyNormal3 => write!(f, "legacyNormal3"),
            Self::LegacyNormal4 => write!(f, "legacyNormal4"),
            Self::LegacyHarsh1 => write!(f, "legacyHarsh1"),
            Self::LegacyHarsh2 => write!(f, "legacyHarsh2"),
            Self::LegacyHarsh3 => write!(f, "legacyHarsh3"),
            Self::LegacyHarsh4 => write!(f, "legacyHarsh4"),
            Self::ThreePt => write!(f, "threePt"),
            Self::Balanced => write!(f, "balanced"),
            Self::Soft => write!(f, "soft"),
            Self::Harsh => write!(f, "harsh"),
            Self::Flood => write!(f, "flood"),
            Self::Contrasting => write!(f, "contrasting"),
            Self::Morning => write!(f, "morning"),
            Self::Sunrise => write!(f, "sunrise"),
            Self::Sunset => write!(f, "sunset"),
            Self::Chilly => write!(f, "chilly"),
            Self::Freezing => write!(f, "freezing"),
            Self::Flat => write!(f, "flat"),
            Self::TwoPt => write!(f, "twoPt"),
            Self::Glow => write!(f, "glow"),
            Self::BrightRoom => write!(f, "brightRoom"),
        }
    }
}

impl std::str::FromStr for STLightRigType {
    type Err = String;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        match s {
            "legacyFlat1" => Ok(Self::LegacyFlat1),
            "legacyFlat2" => Ok(Self::LegacyFlat2),
            "legacyFlat3" => Ok(Self::LegacyFlat3),
            "legacyFlat4" => Ok(Self::LegacyFlat4),
            "legacyNormal1" => Ok(Self::LegacyNormal1),
            "legacyNormal2" => Ok(Self::LegacyNormal2),
            "legacyNormal3" => Ok(Self::LegacyNormal3),
            "legacyNormal4" => Ok(Self::LegacyNormal4),
            "legacyHarsh1" => Ok(Self::LegacyHarsh1),
            "legacyHarsh2" => Ok(Self::LegacyHarsh2),
            "legacyHarsh3" => Ok(Self::LegacyHarsh3),
            "legacyHarsh4" => Ok(Self::LegacyHarsh4),
            "threePt" => Ok(Self::ThreePt),
            "balanced" => Ok(Self::Balanced),
            "soft" => Ok(Self::Soft),
            "harsh" => Ok(Self::Harsh),
            "flood" => Ok(Self::Flood),
            "contrasting" => Ok(Self::Contrasting),
            "morning" => Ok(Self::Morning),
            "sunrise" => Ok(Self::Sunrise),
            "sunset" => Ok(Self::Sunset),
            "chilly" => Ok(Self::Chilly),
            "freezing" => Ok(Self::Freezing),
            "flat" => Ok(Self::Flat),
            "twoPt" => Ok(Self::TwoPt),
            "glow" => Ok(Self::Glow),
            "brightRoom" => Ok(Self::BrightRoom),
            _ => Err(format!("unknown STLightRigType value: {}", s)),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum STBevelPresetType {
    #[serde(rename = "relaxedInset")]
    RelaxedInset,
    #[serde(rename = "circle")]
    Circle,
    #[serde(rename = "slope")]
    Slope,
    #[serde(rename = "cross")]
    Cross,
    #[serde(rename = "angle")]
    Angle,
    #[serde(rename = "softRound")]
    SoftRound,
    #[serde(rename = "convex")]
    Convex,
    #[serde(rename = "coolSlant")]
    CoolSlant,
    #[serde(rename = "divot")]
    Divot,
    #[serde(rename = "riblet")]
    Riblet,
    #[serde(rename = "hardEdge")]
    HardEdge,
    #[serde(rename = "artDeco")]
    ArtDeco,
}

impl std::fmt::Display for STBevelPresetType {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::RelaxedInset => write!(f, "relaxedInset"),
            Self::Circle => write!(f, "circle"),
            Self::Slope => write!(f, "slope"),
            Self::Cross => write!(f, "cross"),
            Self::Angle => write!(f, "angle"),
            Self::SoftRound => write!(f, "softRound"),
            Self::Convex => write!(f, "convex"),
            Self::CoolSlant => write!(f, "coolSlant"),
            Self::Divot => write!(f, "divot"),
            Self::Riblet => write!(f, "riblet"),
            Self::HardEdge => write!(f, "hardEdge"),
            Self::ArtDeco => write!(f, "artDeco"),
        }
    }
}

impl std::str::FromStr for STBevelPresetType {
    type Err = String;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        match s {
            "relaxedInset" => Ok(Self::RelaxedInset),
            "circle" => Ok(Self::Circle),
            "slope" => Ok(Self::Slope),
            "cross" => Ok(Self::Cross),
            "angle" => Ok(Self::Angle),
            "softRound" => Ok(Self::SoftRound),
            "convex" => Ok(Self::Convex),
            "coolSlant" => Ok(Self::CoolSlant),
            "divot" => Ok(Self::Divot),
            "riblet" => Ok(Self::Riblet),
            "hardEdge" => Ok(Self::HardEdge),
            "artDeco" => Ok(Self::ArtDeco),
            _ => Err(format!("unknown STBevelPresetType value: {}", s)),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum STPresetMaterialType {
    #[serde(rename = "legacyMatte")]
    LegacyMatte,
    #[serde(rename = "legacyPlastic")]
    LegacyPlastic,
    #[serde(rename = "legacyMetal")]
    LegacyMetal,
    #[serde(rename = "legacyWireframe")]
    LegacyWireframe,
    #[serde(rename = "matte")]
    Matte,
    #[serde(rename = "plastic")]
    Plastic,
    #[serde(rename = "metal")]
    Metal,
    #[serde(rename = "warmMatte")]
    WarmMatte,
    #[serde(rename = "translucentPowder")]
    TranslucentPowder,
    #[serde(rename = "powder")]
    Powder,
    #[serde(rename = "dkEdge")]
    DkEdge,
    #[serde(rename = "softEdge")]
    SoftEdge,
    #[serde(rename = "clear")]
    Clear,
    #[serde(rename = "flat")]
    Flat,
    #[serde(rename = "softmetal")]
    Softmetal,
}

impl std::fmt::Display for STPresetMaterialType {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::LegacyMatte => write!(f, "legacyMatte"),
            Self::LegacyPlastic => write!(f, "legacyPlastic"),
            Self::LegacyMetal => write!(f, "legacyMetal"),
            Self::LegacyWireframe => write!(f, "legacyWireframe"),
            Self::Matte => write!(f, "matte"),
            Self::Plastic => write!(f, "plastic"),
            Self::Metal => write!(f, "metal"),
            Self::WarmMatte => write!(f, "warmMatte"),
            Self::TranslucentPowder => write!(f, "translucentPowder"),
            Self::Powder => write!(f, "powder"),
            Self::DkEdge => write!(f, "dkEdge"),
            Self::SoftEdge => write!(f, "softEdge"),
            Self::Clear => write!(f, "clear"),
            Self::Flat => write!(f, "flat"),
            Self::Softmetal => write!(f, "softmetal"),
        }
    }
}

impl std::str::FromStr for STPresetMaterialType {
    type Err = String;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        match s {
            "legacyMatte" => Ok(Self::LegacyMatte),
            "legacyPlastic" => Ok(Self::LegacyPlastic),
            "legacyMetal" => Ok(Self::LegacyMetal),
            "legacyWireframe" => Ok(Self::LegacyWireframe),
            "matte" => Ok(Self::Matte),
            "plastic" => Ok(Self::Plastic),
            "metal" => Ok(Self::Metal),
            "warmMatte" => Ok(Self::WarmMatte),
            "translucentPowder" => Ok(Self::TranslucentPowder),
            "powder" => Ok(Self::Powder),
            "dkEdge" => Ok(Self::DkEdge),
            "softEdge" => Ok(Self::SoftEdge),
            "clear" => Ok(Self::Clear),
            "flat" => Ok(Self::Flat),
            "softmetal" => Ok(Self::Softmetal),
            _ => Err(format!("unknown STPresetMaterialType value: {}", s)),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum STPresetShadowVal {
    #[serde(rename = "shdw1")]
    Shdw1,
    #[serde(rename = "shdw2")]
    Shdw2,
    #[serde(rename = "shdw3")]
    Shdw3,
    #[serde(rename = "shdw4")]
    Shdw4,
    #[serde(rename = "shdw5")]
    Shdw5,
    #[serde(rename = "shdw6")]
    Shdw6,
    #[serde(rename = "shdw7")]
    Shdw7,
    #[serde(rename = "shdw8")]
    Shdw8,
    #[serde(rename = "shdw9")]
    Shdw9,
    #[serde(rename = "shdw10")]
    Shdw10,
    #[serde(rename = "shdw11")]
    Shdw11,
    #[serde(rename = "shdw12")]
    Shdw12,
    #[serde(rename = "shdw13")]
    Shdw13,
    #[serde(rename = "shdw14")]
    Shdw14,
    #[serde(rename = "shdw15")]
    Shdw15,
    #[serde(rename = "shdw16")]
    Shdw16,
    #[serde(rename = "shdw17")]
    Shdw17,
    #[serde(rename = "shdw18")]
    Shdw18,
    #[serde(rename = "shdw19")]
    Shdw19,
    #[serde(rename = "shdw20")]
    Shdw20,
}

impl std::fmt::Display for STPresetShadowVal {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::Shdw1 => write!(f, "shdw1"),
            Self::Shdw2 => write!(f, "shdw2"),
            Self::Shdw3 => write!(f, "shdw3"),
            Self::Shdw4 => write!(f, "shdw4"),
            Self::Shdw5 => write!(f, "shdw5"),
            Self::Shdw6 => write!(f, "shdw6"),
            Self::Shdw7 => write!(f, "shdw7"),
            Self::Shdw8 => write!(f, "shdw8"),
            Self::Shdw9 => write!(f, "shdw9"),
            Self::Shdw10 => write!(f, "shdw10"),
            Self::Shdw11 => write!(f, "shdw11"),
            Self::Shdw12 => write!(f, "shdw12"),
            Self::Shdw13 => write!(f, "shdw13"),
            Self::Shdw14 => write!(f, "shdw14"),
            Self::Shdw15 => write!(f, "shdw15"),
            Self::Shdw16 => write!(f, "shdw16"),
            Self::Shdw17 => write!(f, "shdw17"),
            Self::Shdw18 => write!(f, "shdw18"),
            Self::Shdw19 => write!(f, "shdw19"),
            Self::Shdw20 => write!(f, "shdw20"),
        }
    }
}

impl std::str::FromStr for STPresetShadowVal {
    type Err = String;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        match s {
            "shdw1" => Ok(Self::Shdw1),
            "shdw2" => Ok(Self::Shdw2),
            "shdw3" => Ok(Self::Shdw3),
            "shdw4" => Ok(Self::Shdw4),
            "shdw5" => Ok(Self::Shdw5),
            "shdw6" => Ok(Self::Shdw6),
            "shdw7" => Ok(Self::Shdw7),
            "shdw8" => Ok(Self::Shdw8),
            "shdw9" => Ok(Self::Shdw9),
            "shdw10" => Ok(Self::Shdw10),
            "shdw11" => Ok(Self::Shdw11),
            "shdw12" => Ok(Self::Shdw12),
            "shdw13" => Ok(Self::Shdw13),
            "shdw14" => Ok(Self::Shdw14),
            "shdw15" => Ok(Self::Shdw15),
            "shdw16" => Ok(Self::Shdw16),
            "shdw17" => Ok(Self::Shdw17),
            "shdw18" => Ok(Self::Shdw18),
            "shdw19" => Ok(Self::Shdw19),
            "shdw20" => Ok(Self::Shdw20),
            _ => Err(format!("unknown STPresetShadowVal value: {}", s)),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum STPathShadeType {
    #[serde(rename = "shape")]
    Shape,
    #[serde(rename = "circle")]
    Circle,
    #[serde(rename = "rect")]
    Rect,
}

impl std::fmt::Display for STPathShadeType {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::Shape => write!(f, "shape"),
            Self::Circle => write!(f, "circle"),
            Self::Rect => write!(f, "rect"),
        }
    }
}

impl std::str::FromStr for STPathShadeType {
    type Err = String;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        match s {
            "shape" => Ok(Self::Shape),
            "circle" => Ok(Self::Circle),
            "rect" => Ok(Self::Rect),
            _ => Err(format!("unknown STPathShadeType value: {}", s)),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum STTileFlipMode {
    #[serde(rename = "none")]
    None,
    #[serde(rename = "x")]
    X,
    #[serde(rename = "y")]
    Y,
    #[serde(rename = "xy")]
    Xy,
}

impl std::fmt::Display for STTileFlipMode {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::None => write!(f, "none"),
            Self::X => write!(f, "x"),
            Self::Y => write!(f, "y"),
            Self::Xy => write!(f, "xy"),
        }
    }
}

impl std::str::FromStr for STTileFlipMode {
    type Err = String;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        match s {
            "none" => Ok(Self::None),
            "x" => Ok(Self::X),
            "y" => Ok(Self::Y),
            "xy" => Ok(Self::Xy),
            _ => Err(format!("unknown STTileFlipMode value: {}", s)),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum STBlipCompression {
    #[serde(rename = "email")]
    Email,
    #[serde(rename = "screen")]
    Screen,
    #[serde(rename = "print")]
    Print,
    #[serde(rename = "hqprint")]
    Hqprint,
    #[serde(rename = "none")]
    None,
}

impl std::fmt::Display for STBlipCompression {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::Email => write!(f, "email"),
            Self::Screen => write!(f, "screen"),
            Self::Print => write!(f, "print"),
            Self::Hqprint => write!(f, "hqprint"),
            Self::None => write!(f, "none"),
        }
    }
}

impl std::str::FromStr for STBlipCompression {
    type Err = String;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        match s {
            "email" => Ok(Self::Email),
            "screen" => Ok(Self::Screen),
            "print" => Ok(Self::Print),
            "hqprint" => Ok(Self::Hqprint),
            "none" => Ok(Self::None),
            _ => Err(format!("unknown STBlipCompression value: {}", s)),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum STPresetPatternVal {
    #[serde(rename = "pct5")]
    Pct5,
    #[serde(rename = "pct10")]
    Pct10,
    #[serde(rename = "pct20")]
    Pct20,
    #[serde(rename = "pct25")]
    Pct25,
    #[serde(rename = "pct30")]
    Pct30,
    #[serde(rename = "pct40")]
    Pct40,
    #[serde(rename = "pct50")]
    Pct50,
    #[serde(rename = "pct60")]
    Pct60,
    #[serde(rename = "pct70")]
    Pct70,
    #[serde(rename = "pct75")]
    Pct75,
    #[serde(rename = "pct80")]
    Pct80,
    #[serde(rename = "pct90")]
    Pct90,
    #[serde(rename = "horz")]
    Horz,
    #[serde(rename = "vert")]
    Vert,
    #[serde(rename = "ltHorz")]
    LtHorz,
    #[serde(rename = "ltVert")]
    LtVert,
    #[serde(rename = "dkHorz")]
    DkHorz,
    #[serde(rename = "dkVert")]
    DkVert,
    #[serde(rename = "narHorz")]
    NarHorz,
    #[serde(rename = "narVert")]
    NarVert,
    #[serde(rename = "dashHorz")]
    DashHorz,
    #[serde(rename = "dashVert")]
    DashVert,
    #[serde(rename = "cross")]
    Cross,
    #[serde(rename = "dnDiag")]
    DnDiag,
    #[serde(rename = "upDiag")]
    UpDiag,
    #[serde(rename = "ltDnDiag")]
    LtDnDiag,
    #[serde(rename = "ltUpDiag")]
    LtUpDiag,
    #[serde(rename = "dkDnDiag")]
    DkDnDiag,
    #[serde(rename = "dkUpDiag")]
    DkUpDiag,
    #[serde(rename = "wdDnDiag")]
    WdDnDiag,
    #[serde(rename = "wdUpDiag")]
    WdUpDiag,
    #[serde(rename = "dashDnDiag")]
    DashDnDiag,
    #[serde(rename = "dashUpDiag")]
    DashUpDiag,
    #[serde(rename = "diagCross")]
    DiagCross,
    #[serde(rename = "smCheck")]
    SmCheck,
    #[serde(rename = "lgCheck")]
    LgCheck,
    #[serde(rename = "smGrid")]
    SmGrid,
    #[serde(rename = "lgGrid")]
    LgGrid,
    #[serde(rename = "dotGrid")]
    DotGrid,
    #[serde(rename = "smConfetti")]
    SmConfetti,
    #[serde(rename = "lgConfetti")]
    LgConfetti,
    #[serde(rename = "horzBrick")]
    HorzBrick,
    #[serde(rename = "diagBrick")]
    DiagBrick,
    #[serde(rename = "solidDmnd")]
    SolidDmnd,
    #[serde(rename = "openDmnd")]
    OpenDmnd,
    #[serde(rename = "dotDmnd")]
    DotDmnd,
    #[serde(rename = "plaid")]
    Plaid,
    #[serde(rename = "sphere")]
    Sphere,
    #[serde(rename = "weave")]
    Weave,
    #[serde(rename = "divot")]
    Divot,
    #[serde(rename = "shingle")]
    Shingle,
    #[serde(rename = "wave")]
    Wave,
    #[serde(rename = "trellis")]
    Trellis,
    #[serde(rename = "zigZag")]
    ZigZag,
}

impl std::fmt::Display for STPresetPatternVal {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::Pct5 => write!(f, "pct5"),
            Self::Pct10 => write!(f, "pct10"),
            Self::Pct20 => write!(f, "pct20"),
            Self::Pct25 => write!(f, "pct25"),
            Self::Pct30 => write!(f, "pct30"),
            Self::Pct40 => write!(f, "pct40"),
            Self::Pct50 => write!(f, "pct50"),
            Self::Pct60 => write!(f, "pct60"),
            Self::Pct70 => write!(f, "pct70"),
            Self::Pct75 => write!(f, "pct75"),
            Self::Pct80 => write!(f, "pct80"),
            Self::Pct90 => write!(f, "pct90"),
            Self::Horz => write!(f, "horz"),
            Self::Vert => write!(f, "vert"),
            Self::LtHorz => write!(f, "ltHorz"),
            Self::LtVert => write!(f, "ltVert"),
            Self::DkHorz => write!(f, "dkHorz"),
            Self::DkVert => write!(f, "dkVert"),
            Self::NarHorz => write!(f, "narHorz"),
            Self::NarVert => write!(f, "narVert"),
            Self::DashHorz => write!(f, "dashHorz"),
            Self::DashVert => write!(f, "dashVert"),
            Self::Cross => write!(f, "cross"),
            Self::DnDiag => write!(f, "dnDiag"),
            Self::UpDiag => write!(f, "upDiag"),
            Self::LtDnDiag => write!(f, "ltDnDiag"),
            Self::LtUpDiag => write!(f, "ltUpDiag"),
            Self::DkDnDiag => write!(f, "dkDnDiag"),
            Self::DkUpDiag => write!(f, "dkUpDiag"),
            Self::WdDnDiag => write!(f, "wdDnDiag"),
            Self::WdUpDiag => write!(f, "wdUpDiag"),
            Self::DashDnDiag => write!(f, "dashDnDiag"),
            Self::DashUpDiag => write!(f, "dashUpDiag"),
            Self::DiagCross => write!(f, "diagCross"),
            Self::SmCheck => write!(f, "smCheck"),
            Self::LgCheck => write!(f, "lgCheck"),
            Self::SmGrid => write!(f, "smGrid"),
            Self::LgGrid => write!(f, "lgGrid"),
            Self::DotGrid => write!(f, "dotGrid"),
            Self::SmConfetti => write!(f, "smConfetti"),
            Self::LgConfetti => write!(f, "lgConfetti"),
            Self::HorzBrick => write!(f, "horzBrick"),
            Self::DiagBrick => write!(f, "diagBrick"),
            Self::SolidDmnd => write!(f, "solidDmnd"),
            Self::OpenDmnd => write!(f, "openDmnd"),
            Self::DotDmnd => write!(f, "dotDmnd"),
            Self::Plaid => write!(f, "plaid"),
            Self::Sphere => write!(f, "sphere"),
            Self::Weave => write!(f, "weave"),
            Self::Divot => write!(f, "divot"),
            Self::Shingle => write!(f, "shingle"),
            Self::Wave => write!(f, "wave"),
            Self::Trellis => write!(f, "trellis"),
            Self::ZigZag => write!(f, "zigZag"),
        }
    }
}

impl std::str::FromStr for STPresetPatternVal {
    type Err = String;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        match s {
            "pct5" => Ok(Self::Pct5),
            "pct10" => Ok(Self::Pct10),
            "pct20" => Ok(Self::Pct20),
            "pct25" => Ok(Self::Pct25),
            "pct30" => Ok(Self::Pct30),
            "pct40" => Ok(Self::Pct40),
            "pct50" => Ok(Self::Pct50),
            "pct60" => Ok(Self::Pct60),
            "pct70" => Ok(Self::Pct70),
            "pct75" => Ok(Self::Pct75),
            "pct80" => Ok(Self::Pct80),
            "pct90" => Ok(Self::Pct90),
            "horz" => Ok(Self::Horz),
            "vert" => Ok(Self::Vert),
            "ltHorz" => Ok(Self::LtHorz),
            "ltVert" => Ok(Self::LtVert),
            "dkHorz" => Ok(Self::DkHorz),
            "dkVert" => Ok(Self::DkVert),
            "narHorz" => Ok(Self::NarHorz),
            "narVert" => Ok(Self::NarVert),
            "dashHorz" => Ok(Self::DashHorz),
            "dashVert" => Ok(Self::DashVert),
            "cross" => Ok(Self::Cross),
            "dnDiag" => Ok(Self::DnDiag),
            "upDiag" => Ok(Self::UpDiag),
            "ltDnDiag" => Ok(Self::LtDnDiag),
            "ltUpDiag" => Ok(Self::LtUpDiag),
            "dkDnDiag" => Ok(Self::DkDnDiag),
            "dkUpDiag" => Ok(Self::DkUpDiag),
            "wdDnDiag" => Ok(Self::WdDnDiag),
            "wdUpDiag" => Ok(Self::WdUpDiag),
            "dashDnDiag" => Ok(Self::DashDnDiag),
            "dashUpDiag" => Ok(Self::DashUpDiag),
            "diagCross" => Ok(Self::DiagCross),
            "smCheck" => Ok(Self::SmCheck),
            "lgCheck" => Ok(Self::LgCheck),
            "smGrid" => Ok(Self::SmGrid),
            "lgGrid" => Ok(Self::LgGrid),
            "dotGrid" => Ok(Self::DotGrid),
            "smConfetti" => Ok(Self::SmConfetti),
            "lgConfetti" => Ok(Self::LgConfetti),
            "horzBrick" => Ok(Self::HorzBrick),
            "diagBrick" => Ok(Self::DiagBrick),
            "solidDmnd" => Ok(Self::SolidDmnd),
            "openDmnd" => Ok(Self::OpenDmnd),
            "dotDmnd" => Ok(Self::DotDmnd),
            "plaid" => Ok(Self::Plaid),
            "sphere" => Ok(Self::Sphere),
            "weave" => Ok(Self::Weave),
            "divot" => Ok(Self::Divot),
            "shingle" => Ok(Self::Shingle),
            "wave" => Ok(Self::Wave),
            "trellis" => Ok(Self::Trellis),
            "zigZag" => Ok(Self::ZigZag),
            _ => Err(format!("unknown STPresetPatternVal value: {}", s)),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum STBlendMode {
    #[serde(rename = "over")]
    Over,
    #[serde(rename = "mult")]
    Mult,
    #[serde(rename = "screen")]
    Screen,
    #[serde(rename = "darken")]
    Darken,
    #[serde(rename = "lighten")]
    Lighten,
}

impl std::fmt::Display for STBlendMode {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::Over => write!(f, "over"),
            Self::Mult => write!(f, "mult"),
            Self::Screen => write!(f, "screen"),
            Self::Darken => write!(f, "darken"),
            Self::Lighten => write!(f, "lighten"),
        }
    }
}

impl std::str::FromStr for STBlendMode {
    type Err = String;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        match s {
            "over" => Ok(Self::Over),
            "mult" => Ok(Self::Mult),
            "screen" => Ok(Self::Screen),
            "darken" => Ok(Self::Darken),
            "lighten" => Ok(Self::Lighten),
            _ => Err(format!("unknown STBlendMode value: {}", s)),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum STEffectContainerType {
    #[serde(rename = "sib")]
    Sib,
    #[serde(rename = "tree")]
    Tree,
}

impl std::fmt::Display for STEffectContainerType {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::Sib => write!(f, "sib"),
            Self::Tree => write!(f, "tree"),
        }
    }
}

impl std::str::FromStr for STEffectContainerType {
    type Err = String;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        match s {
            "sib" => Ok(Self::Sib),
            "tree" => Ok(Self::Tree),
            _ => Err(format!("unknown STEffectContainerType value: {}", s)),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum STShapeType {
    #[serde(rename = "line")]
    Line,
    #[serde(rename = "lineInv")]
    LineInv,
    #[serde(rename = "triangle")]
    Triangle,
    #[serde(rename = "rtTriangle")]
    RtTriangle,
    #[serde(rename = "rect")]
    Rect,
    #[serde(rename = "diamond")]
    Diamond,
    #[serde(rename = "parallelogram")]
    Parallelogram,
    #[serde(rename = "trapezoid")]
    Trapezoid,
    #[serde(rename = "nonIsoscelesTrapezoid")]
    NonIsoscelesTrapezoid,
    #[serde(rename = "pentagon")]
    Pentagon,
    #[serde(rename = "hexagon")]
    Hexagon,
    #[serde(rename = "heptagon")]
    Heptagon,
    #[serde(rename = "octagon")]
    Octagon,
    #[serde(rename = "decagon")]
    Decagon,
    #[serde(rename = "dodecagon")]
    Dodecagon,
    #[serde(rename = "star4")]
    Star4,
    #[serde(rename = "star5")]
    Star5,
    #[serde(rename = "star6")]
    Star6,
    #[serde(rename = "star7")]
    Star7,
    #[serde(rename = "star8")]
    Star8,
    #[serde(rename = "star10")]
    Star10,
    #[serde(rename = "star12")]
    Star12,
    #[serde(rename = "star16")]
    Star16,
    #[serde(rename = "star24")]
    Star24,
    #[serde(rename = "star32")]
    Star32,
    #[serde(rename = "roundRect")]
    RoundRect,
    #[serde(rename = "round1Rect")]
    Round1Rect,
    #[serde(rename = "round2SameRect")]
    Round2SameRect,
    #[serde(rename = "round2DiagRect")]
    Round2DiagRect,
    #[serde(rename = "snipRoundRect")]
    SnipRoundRect,
    #[serde(rename = "snip1Rect")]
    Snip1Rect,
    #[serde(rename = "snip2SameRect")]
    Snip2SameRect,
    #[serde(rename = "snip2DiagRect")]
    Snip2DiagRect,
    #[serde(rename = "plaque")]
    Plaque,
    #[serde(rename = "ellipse")]
    Ellipse,
    #[serde(rename = "teardrop")]
    Teardrop,
    #[serde(rename = "homePlate")]
    HomePlate,
    #[serde(rename = "chevron")]
    Chevron,
    #[serde(rename = "pieWedge")]
    PieWedge,
    #[serde(rename = "pie")]
    Pie,
    #[serde(rename = "blockArc")]
    BlockArc,
    #[serde(rename = "donut")]
    Donut,
    #[serde(rename = "noSmoking")]
    NoSmoking,
    #[serde(rename = "rightArrow")]
    RightArrow,
    #[serde(rename = "leftArrow")]
    LeftArrow,
    #[serde(rename = "upArrow")]
    UpArrow,
    #[serde(rename = "downArrow")]
    DownArrow,
    #[serde(rename = "stripedRightArrow")]
    StripedRightArrow,
    #[serde(rename = "notchedRightArrow")]
    NotchedRightArrow,
    #[serde(rename = "bentUpArrow")]
    BentUpArrow,
    #[serde(rename = "leftRightArrow")]
    LeftRightArrow,
    #[serde(rename = "upDownArrow")]
    UpDownArrow,
    #[serde(rename = "leftUpArrow")]
    LeftUpArrow,
    #[serde(rename = "leftRightUpArrow")]
    LeftRightUpArrow,
    #[serde(rename = "quadArrow")]
    QuadArrow,
    #[serde(rename = "leftArrowCallout")]
    LeftArrowCallout,
    #[serde(rename = "rightArrowCallout")]
    RightArrowCallout,
    #[serde(rename = "upArrowCallout")]
    UpArrowCallout,
    #[serde(rename = "downArrowCallout")]
    DownArrowCallout,
    #[serde(rename = "leftRightArrowCallout")]
    LeftRightArrowCallout,
    #[serde(rename = "upDownArrowCallout")]
    UpDownArrowCallout,
    #[serde(rename = "quadArrowCallout")]
    QuadArrowCallout,
    #[serde(rename = "bentArrow")]
    BentArrow,
    #[serde(rename = "uturnArrow")]
    UturnArrow,
    #[serde(rename = "circularArrow")]
    CircularArrow,
    #[serde(rename = "leftCircularArrow")]
    LeftCircularArrow,
    #[serde(rename = "leftRightCircularArrow")]
    LeftRightCircularArrow,
    #[serde(rename = "curvedRightArrow")]
    CurvedRightArrow,
    #[serde(rename = "curvedLeftArrow")]
    CurvedLeftArrow,
    #[serde(rename = "curvedUpArrow")]
    CurvedUpArrow,
    #[serde(rename = "curvedDownArrow")]
    CurvedDownArrow,
    #[serde(rename = "swooshArrow")]
    SwooshArrow,
    #[serde(rename = "cube")]
    Cube,
    #[serde(rename = "can")]
    Can,
    #[serde(rename = "lightningBolt")]
    LightningBolt,
    #[serde(rename = "heart")]
    Heart,
    #[serde(rename = "sun")]
    Sun,
    #[serde(rename = "moon")]
    Moon,
    #[serde(rename = "smileyFace")]
    SmileyFace,
    #[serde(rename = "irregularSeal1")]
    IrregularSeal1,
    #[serde(rename = "irregularSeal2")]
    IrregularSeal2,
    #[serde(rename = "foldedCorner")]
    FoldedCorner,
    #[serde(rename = "bevel")]
    Bevel,
    #[serde(rename = "frame")]
    Frame,
    #[serde(rename = "halfFrame")]
    HalfFrame,
    #[serde(rename = "corner")]
    Corner,
    #[serde(rename = "diagStripe")]
    DiagStripe,
    #[serde(rename = "chord")]
    Chord,
    #[serde(rename = "arc")]
    Arc,
    #[serde(rename = "leftBracket")]
    LeftBracket,
    #[serde(rename = "rightBracket")]
    RightBracket,
    #[serde(rename = "leftBrace")]
    LeftBrace,
    #[serde(rename = "rightBrace")]
    RightBrace,
    #[serde(rename = "bracketPair")]
    BracketPair,
    #[serde(rename = "bracePair")]
    BracePair,
    #[serde(rename = "straightConnector1")]
    StraightConnector1,
    #[serde(rename = "bentConnector2")]
    BentConnector2,
    #[serde(rename = "bentConnector3")]
    BentConnector3,
    #[serde(rename = "bentConnector4")]
    BentConnector4,
    #[serde(rename = "bentConnector5")]
    BentConnector5,
    #[serde(rename = "curvedConnector2")]
    CurvedConnector2,
    #[serde(rename = "curvedConnector3")]
    CurvedConnector3,
    #[serde(rename = "curvedConnector4")]
    CurvedConnector4,
    #[serde(rename = "curvedConnector5")]
    CurvedConnector5,
    #[serde(rename = "callout1")]
    Callout1,
    #[serde(rename = "callout2")]
    Callout2,
    #[serde(rename = "callout3")]
    Callout3,
    #[serde(rename = "accentCallout1")]
    AccentCallout1,
    #[serde(rename = "accentCallout2")]
    AccentCallout2,
    #[serde(rename = "accentCallout3")]
    AccentCallout3,
    #[serde(rename = "borderCallout1")]
    BorderCallout1,
    #[serde(rename = "borderCallout2")]
    BorderCallout2,
    #[serde(rename = "borderCallout3")]
    BorderCallout3,
    #[serde(rename = "accentBorderCallout1")]
    AccentBorderCallout1,
    #[serde(rename = "accentBorderCallout2")]
    AccentBorderCallout2,
    #[serde(rename = "accentBorderCallout3")]
    AccentBorderCallout3,
    #[serde(rename = "wedgeRectCallout")]
    WedgeRectCallout,
    #[serde(rename = "wedgeRoundRectCallout")]
    WedgeRoundRectCallout,
    #[serde(rename = "wedgeEllipseCallout")]
    WedgeEllipseCallout,
    #[serde(rename = "cloudCallout")]
    CloudCallout,
    #[serde(rename = "cloud")]
    Cloud,
    #[serde(rename = "ribbon")]
    Ribbon,
    #[serde(rename = "ribbon2")]
    Ribbon2,
    #[serde(rename = "ellipseRibbon")]
    EllipseRibbon,
    #[serde(rename = "ellipseRibbon2")]
    EllipseRibbon2,
    #[serde(rename = "leftRightRibbon")]
    LeftRightRibbon,
    #[serde(rename = "verticalScroll")]
    VerticalScroll,
    #[serde(rename = "horizontalScroll")]
    HorizontalScroll,
    #[serde(rename = "wave")]
    Wave,
    #[serde(rename = "doubleWave")]
    DoubleWave,
    #[serde(rename = "plus")]
    Plus,
    #[serde(rename = "flowChartProcess")]
    FlowChartProcess,
    #[serde(rename = "flowChartDecision")]
    FlowChartDecision,
    #[serde(rename = "flowChartInputOutput")]
    FlowChartInputOutput,
    #[serde(rename = "flowChartPredefinedProcess")]
    FlowChartPredefinedProcess,
    #[serde(rename = "flowChartInternalStorage")]
    FlowChartInternalStorage,
    #[serde(rename = "flowChartDocument")]
    FlowChartDocument,
    #[serde(rename = "flowChartMultidocument")]
    FlowChartMultidocument,
    #[serde(rename = "flowChartTerminator")]
    FlowChartTerminator,
    #[serde(rename = "flowChartPreparation")]
    FlowChartPreparation,
    #[serde(rename = "flowChartManualInput")]
    FlowChartManualInput,
    #[serde(rename = "flowChartManualOperation")]
    FlowChartManualOperation,
    #[serde(rename = "flowChartConnector")]
    FlowChartConnector,
    #[serde(rename = "flowChartPunchedCard")]
    FlowChartPunchedCard,
    #[serde(rename = "flowChartPunchedTape")]
    FlowChartPunchedTape,
    #[serde(rename = "flowChartSummingJunction")]
    FlowChartSummingJunction,
    #[serde(rename = "flowChartOr")]
    FlowChartOr,
    #[serde(rename = "flowChartCollate")]
    FlowChartCollate,
    #[serde(rename = "flowChartSort")]
    FlowChartSort,
    #[serde(rename = "flowChartExtract")]
    FlowChartExtract,
    #[serde(rename = "flowChartMerge")]
    FlowChartMerge,
    #[serde(rename = "flowChartOfflineStorage")]
    FlowChartOfflineStorage,
    #[serde(rename = "flowChartOnlineStorage")]
    FlowChartOnlineStorage,
    #[serde(rename = "flowChartMagneticTape")]
    FlowChartMagneticTape,
    #[serde(rename = "flowChartMagneticDisk")]
    FlowChartMagneticDisk,
    #[serde(rename = "flowChartMagneticDrum")]
    FlowChartMagneticDrum,
    #[serde(rename = "flowChartDisplay")]
    FlowChartDisplay,
    #[serde(rename = "flowChartDelay")]
    FlowChartDelay,
    #[serde(rename = "flowChartAlternateProcess")]
    FlowChartAlternateProcess,
    #[serde(rename = "flowChartOffpageConnector")]
    FlowChartOffpageConnector,
    #[serde(rename = "actionButtonBlank")]
    ActionButtonBlank,
    #[serde(rename = "actionButtonHome")]
    ActionButtonHome,
    #[serde(rename = "actionButtonHelp")]
    ActionButtonHelp,
    #[serde(rename = "actionButtonInformation")]
    ActionButtonInformation,
    #[serde(rename = "actionButtonForwardNext")]
    ActionButtonForwardNext,
    #[serde(rename = "actionButtonBackPrevious")]
    ActionButtonBackPrevious,
    #[serde(rename = "actionButtonEnd")]
    ActionButtonEnd,
    #[serde(rename = "actionButtonBeginning")]
    ActionButtonBeginning,
    #[serde(rename = "actionButtonReturn")]
    ActionButtonReturn,
    #[serde(rename = "actionButtonDocument")]
    ActionButtonDocument,
    #[serde(rename = "actionButtonSound")]
    ActionButtonSound,
    #[serde(rename = "actionButtonMovie")]
    ActionButtonMovie,
    #[serde(rename = "gear6")]
    Gear6,
    #[serde(rename = "gear9")]
    Gear9,
    #[serde(rename = "funnel")]
    Funnel,
    #[serde(rename = "mathPlus")]
    MathPlus,
    #[serde(rename = "mathMinus")]
    MathMinus,
    #[serde(rename = "mathMultiply")]
    MathMultiply,
    #[serde(rename = "mathDivide")]
    MathDivide,
    #[serde(rename = "mathEqual")]
    MathEqual,
    #[serde(rename = "mathNotEqual")]
    MathNotEqual,
    #[serde(rename = "cornerTabs")]
    CornerTabs,
    #[serde(rename = "squareTabs")]
    SquareTabs,
    #[serde(rename = "plaqueTabs")]
    PlaqueTabs,
    #[serde(rename = "chartX")]
    ChartX,
    #[serde(rename = "chartStar")]
    ChartStar,
    #[serde(rename = "chartPlus")]
    ChartPlus,
}

impl std::fmt::Display for STShapeType {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::Line => write!(f, "line"),
            Self::LineInv => write!(f, "lineInv"),
            Self::Triangle => write!(f, "triangle"),
            Self::RtTriangle => write!(f, "rtTriangle"),
            Self::Rect => write!(f, "rect"),
            Self::Diamond => write!(f, "diamond"),
            Self::Parallelogram => write!(f, "parallelogram"),
            Self::Trapezoid => write!(f, "trapezoid"),
            Self::NonIsoscelesTrapezoid => write!(f, "nonIsoscelesTrapezoid"),
            Self::Pentagon => write!(f, "pentagon"),
            Self::Hexagon => write!(f, "hexagon"),
            Self::Heptagon => write!(f, "heptagon"),
            Self::Octagon => write!(f, "octagon"),
            Self::Decagon => write!(f, "decagon"),
            Self::Dodecagon => write!(f, "dodecagon"),
            Self::Star4 => write!(f, "star4"),
            Self::Star5 => write!(f, "star5"),
            Self::Star6 => write!(f, "star6"),
            Self::Star7 => write!(f, "star7"),
            Self::Star8 => write!(f, "star8"),
            Self::Star10 => write!(f, "star10"),
            Self::Star12 => write!(f, "star12"),
            Self::Star16 => write!(f, "star16"),
            Self::Star24 => write!(f, "star24"),
            Self::Star32 => write!(f, "star32"),
            Self::RoundRect => write!(f, "roundRect"),
            Self::Round1Rect => write!(f, "round1Rect"),
            Self::Round2SameRect => write!(f, "round2SameRect"),
            Self::Round2DiagRect => write!(f, "round2DiagRect"),
            Self::SnipRoundRect => write!(f, "snipRoundRect"),
            Self::Snip1Rect => write!(f, "snip1Rect"),
            Self::Snip2SameRect => write!(f, "snip2SameRect"),
            Self::Snip2DiagRect => write!(f, "snip2DiagRect"),
            Self::Plaque => write!(f, "plaque"),
            Self::Ellipse => write!(f, "ellipse"),
            Self::Teardrop => write!(f, "teardrop"),
            Self::HomePlate => write!(f, "homePlate"),
            Self::Chevron => write!(f, "chevron"),
            Self::PieWedge => write!(f, "pieWedge"),
            Self::Pie => write!(f, "pie"),
            Self::BlockArc => write!(f, "blockArc"),
            Self::Donut => write!(f, "donut"),
            Self::NoSmoking => write!(f, "noSmoking"),
            Self::RightArrow => write!(f, "rightArrow"),
            Self::LeftArrow => write!(f, "leftArrow"),
            Self::UpArrow => write!(f, "upArrow"),
            Self::DownArrow => write!(f, "downArrow"),
            Self::StripedRightArrow => write!(f, "stripedRightArrow"),
            Self::NotchedRightArrow => write!(f, "notchedRightArrow"),
            Self::BentUpArrow => write!(f, "bentUpArrow"),
            Self::LeftRightArrow => write!(f, "leftRightArrow"),
            Self::UpDownArrow => write!(f, "upDownArrow"),
            Self::LeftUpArrow => write!(f, "leftUpArrow"),
            Self::LeftRightUpArrow => write!(f, "leftRightUpArrow"),
            Self::QuadArrow => write!(f, "quadArrow"),
            Self::LeftArrowCallout => write!(f, "leftArrowCallout"),
            Self::RightArrowCallout => write!(f, "rightArrowCallout"),
            Self::UpArrowCallout => write!(f, "upArrowCallout"),
            Self::DownArrowCallout => write!(f, "downArrowCallout"),
            Self::LeftRightArrowCallout => write!(f, "leftRightArrowCallout"),
            Self::UpDownArrowCallout => write!(f, "upDownArrowCallout"),
            Self::QuadArrowCallout => write!(f, "quadArrowCallout"),
            Self::BentArrow => write!(f, "bentArrow"),
            Self::UturnArrow => write!(f, "uturnArrow"),
            Self::CircularArrow => write!(f, "circularArrow"),
            Self::LeftCircularArrow => write!(f, "leftCircularArrow"),
            Self::LeftRightCircularArrow => write!(f, "leftRightCircularArrow"),
            Self::CurvedRightArrow => write!(f, "curvedRightArrow"),
            Self::CurvedLeftArrow => write!(f, "curvedLeftArrow"),
            Self::CurvedUpArrow => write!(f, "curvedUpArrow"),
            Self::CurvedDownArrow => write!(f, "curvedDownArrow"),
            Self::SwooshArrow => write!(f, "swooshArrow"),
            Self::Cube => write!(f, "cube"),
            Self::Can => write!(f, "can"),
            Self::LightningBolt => write!(f, "lightningBolt"),
            Self::Heart => write!(f, "heart"),
            Self::Sun => write!(f, "sun"),
            Self::Moon => write!(f, "moon"),
            Self::SmileyFace => write!(f, "smileyFace"),
            Self::IrregularSeal1 => write!(f, "irregularSeal1"),
            Self::IrregularSeal2 => write!(f, "irregularSeal2"),
            Self::FoldedCorner => write!(f, "foldedCorner"),
            Self::Bevel => write!(f, "bevel"),
            Self::Frame => write!(f, "frame"),
            Self::HalfFrame => write!(f, "halfFrame"),
            Self::Corner => write!(f, "corner"),
            Self::DiagStripe => write!(f, "diagStripe"),
            Self::Chord => write!(f, "chord"),
            Self::Arc => write!(f, "arc"),
            Self::LeftBracket => write!(f, "leftBracket"),
            Self::RightBracket => write!(f, "rightBracket"),
            Self::LeftBrace => write!(f, "leftBrace"),
            Self::RightBrace => write!(f, "rightBrace"),
            Self::BracketPair => write!(f, "bracketPair"),
            Self::BracePair => write!(f, "bracePair"),
            Self::StraightConnector1 => write!(f, "straightConnector1"),
            Self::BentConnector2 => write!(f, "bentConnector2"),
            Self::BentConnector3 => write!(f, "bentConnector3"),
            Self::BentConnector4 => write!(f, "bentConnector4"),
            Self::BentConnector5 => write!(f, "bentConnector5"),
            Self::CurvedConnector2 => write!(f, "curvedConnector2"),
            Self::CurvedConnector3 => write!(f, "curvedConnector3"),
            Self::CurvedConnector4 => write!(f, "curvedConnector4"),
            Self::CurvedConnector5 => write!(f, "curvedConnector5"),
            Self::Callout1 => write!(f, "callout1"),
            Self::Callout2 => write!(f, "callout2"),
            Self::Callout3 => write!(f, "callout3"),
            Self::AccentCallout1 => write!(f, "accentCallout1"),
            Self::AccentCallout2 => write!(f, "accentCallout2"),
            Self::AccentCallout3 => write!(f, "accentCallout3"),
            Self::BorderCallout1 => write!(f, "borderCallout1"),
            Self::BorderCallout2 => write!(f, "borderCallout2"),
            Self::BorderCallout3 => write!(f, "borderCallout3"),
            Self::AccentBorderCallout1 => write!(f, "accentBorderCallout1"),
            Self::AccentBorderCallout2 => write!(f, "accentBorderCallout2"),
            Self::AccentBorderCallout3 => write!(f, "accentBorderCallout3"),
            Self::WedgeRectCallout => write!(f, "wedgeRectCallout"),
            Self::WedgeRoundRectCallout => write!(f, "wedgeRoundRectCallout"),
            Self::WedgeEllipseCallout => write!(f, "wedgeEllipseCallout"),
            Self::CloudCallout => write!(f, "cloudCallout"),
            Self::Cloud => write!(f, "cloud"),
            Self::Ribbon => write!(f, "ribbon"),
            Self::Ribbon2 => write!(f, "ribbon2"),
            Self::EllipseRibbon => write!(f, "ellipseRibbon"),
            Self::EllipseRibbon2 => write!(f, "ellipseRibbon2"),
            Self::LeftRightRibbon => write!(f, "leftRightRibbon"),
            Self::VerticalScroll => write!(f, "verticalScroll"),
            Self::HorizontalScroll => write!(f, "horizontalScroll"),
            Self::Wave => write!(f, "wave"),
            Self::DoubleWave => write!(f, "doubleWave"),
            Self::Plus => write!(f, "plus"),
            Self::FlowChartProcess => write!(f, "flowChartProcess"),
            Self::FlowChartDecision => write!(f, "flowChartDecision"),
            Self::FlowChartInputOutput => write!(f, "flowChartInputOutput"),
            Self::FlowChartPredefinedProcess => write!(f, "flowChartPredefinedProcess"),
            Self::FlowChartInternalStorage => write!(f, "flowChartInternalStorage"),
            Self::FlowChartDocument => write!(f, "flowChartDocument"),
            Self::FlowChartMultidocument => write!(f, "flowChartMultidocument"),
            Self::FlowChartTerminator => write!(f, "flowChartTerminator"),
            Self::FlowChartPreparation => write!(f, "flowChartPreparation"),
            Self::FlowChartManualInput => write!(f, "flowChartManualInput"),
            Self::FlowChartManualOperation => write!(f, "flowChartManualOperation"),
            Self::FlowChartConnector => write!(f, "flowChartConnector"),
            Self::FlowChartPunchedCard => write!(f, "flowChartPunchedCard"),
            Self::FlowChartPunchedTape => write!(f, "flowChartPunchedTape"),
            Self::FlowChartSummingJunction => write!(f, "flowChartSummingJunction"),
            Self::FlowChartOr => write!(f, "flowChartOr"),
            Self::FlowChartCollate => write!(f, "flowChartCollate"),
            Self::FlowChartSort => write!(f, "flowChartSort"),
            Self::FlowChartExtract => write!(f, "flowChartExtract"),
            Self::FlowChartMerge => write!(f, "flowChartMerge"),
            Self::FlowChartOfflineStorage => write!(f, "flowChartOfflineStorage"),
            Self::FlowChartOnlineStorage => write!(f, "flowChartOnlineStorage"),
            Self::FlowChartMagneticTape => write!(f, "flowChartMagneticTape"),
            Self::FlowChartMagneticDisk => write!(f, "flowChartMagneticDisk"),
            Self::FlowChartMagneticDrum => write!(f, "flowChartMagneticDrum"),
            Self::FlowChartDisplay => write!(f, "flowChartDisplay"),
            Self::FlowChartDelay => write!(f, "flowChartDelay"),
            Self::FlowChartAlternateProcess => write!(f, "flowChartAlternateProcess"),
            Self::FlowChartOffpageConnector => write!(f, "flowChartOffpageConnector"),
            Self::ActionButtonBlank => write!(f, "actionButtonBlank"),
            Self::ActionButtonHome => write!(f, "actionButtonHome"),
            Self::ActionButtonHelp => write!(f, "actionButtonHelp"),
            Self::ActionButtonInformation => write!(f, "actionButtonInformation"),
            Self::ActionButtonForwardNext => write!(f, "actionButtonForwardNext"),
            Self::ActionButtonBackPrevious => write!(f, "actionButtonBackPrevious"),
            Self::ActionButtonEnd => write!(f, "actionButtonEnd"),
            Self::ActionButtonBeginning => write!(f, "actionButtonBeginning"),
            Self::ActionButtonReturn => write!(f, "actionButtonReturn"),
            Self::ActionButtonDocument => write!(f, "actionButtonDocument"),
            Self::ActionButtonSound => write!(f, "actionButtonSound"),
            Self::ActionButtonMovie => write!(f, "actionButtonMovie"),
            Self::Gear6 => write!(f, "gear6"),
            Self::Gear9 => write!(f, "gear9"),
            Self::Funnel => write!(f, "funnel"),
            Self::MathPlus => write!(f, "mathPlus"),
            Self::MathMinus => write!(f, "mathMinus"),
            Self::MathMultiply => write!(f, "mathMultiply"),
            Self::MathDivide => write!(f, "mathDivide"),
            Self::MathEqual => write!(f, "mathEqual"),
            Self::MathNotEqual => write!(f, "mathNotEqual"),
            Self::CornerTabs => write!(f, "cornerTabs"),
            Self::SquareTabs => write!(f, "squareTabs"),
            Self::PlaqueTabs => write!(f, "plaqueTabs"),
            Self::ChartX => write!(f, "chartX"),
            Self::ChartStar => write!(f, "chartStar"),
            Self::ChartPlus => write!(f, "chartPlus"),
        }
    }
}

impl std::str::FromStr for STShapeType {
    type Err = String;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        match s {
            "line" => Ok(Self::Line),
            "lineInv" => Ok(Self::LineInv),
            "triangle" => Ok(Self::Triangle),
            "rtTriangle" => Ok(Self::RtTriangle),
            "rect" => Ok(Self::Rect),
            "diamond" => Ok(Self::Diamond),
            "parallelogram" => Ok(Self::Parallelogram),
            "trapezoid" => Ok(Self::Trapezoid),
            "nonIsoscelesTrapezoid" => Ok(Self::NonIsoscelesTrapezoid),
            "pentagon" => Ok(Self::Pentagon),
            "hexagon" => Ok(Self::Hexagon),
            "heptagon" => Ok(Self::Heptagon),
            "octagon" => Ok(Self::Octagon),
            "decagon" => Ok(Self::Decagon),
            "dodecagon" => Ok(Self::Dodecagon),
            "star4" => Ok(Self::Star4),
            "star5" => Ok(Self::Star5),
            "star6" => Ok(Self::Star6),
            "star7" => Ok(Self::Star7),
            "star8" => Ok(Self::Star8),
            "star10" => Ok(Self::Star10),
            "star12" => Ok(Self::Star12),
            "star16" => Ok(Self::Star16),
            "star24" => Ok(Self::Star24),
            "star32" => Ok(Self::Star32),
            "roundRect" => Ok(Self::RoundRect),
            "round1Rect" => Ok(Self::Round1Rect),
            "round2SameRect" => Ok(Self::Round2SameRect),
            "round2DiagRect" => Ok(Self::Round2DiagRect),
            "snipRoundRect" => Ok(Self::SnipRoundRect),
            "snip1Rect" => Ok(Self::Snip1Rect),
            "snip2SameRect" => Ok(Self::Snip2SameRect),
            "snip2DiagRect" => Ok(Self::Snip2DiagRect),
            "plaque" => Ok(Self::Plaque),
            "ellipse" => Ok(Self::Ellipse),
            "teardrop" => Ok(Self::Teardrop),
            "homePlate" => Ok(Self::HomePlate),
            "chevron" => Ok(Self::Chevron),
            "pieWedge" => Ok(Self::PieWedge),
            "pie" => Ok(Self::Pie),
            "blockArc" => Ok(Self::BlockArc),
            "donut" => Ok(Self::Donut),
            "noSmoking" => Ok(Self::NoSmoking),
            "rightArrow" => Ok(Self::RightArrow),
            "leftArrow" => Ok(Self::LeftArrow),
            "upArrow" => Ok(Self::UpArrow),
            "downArrow" => Ok(Self::DownArrow),
            "stripedRightArrow" => Ok(Self::StripedRightArrow),
            "notchedRightArrow" => Ok(Self::NotchedRightArrow),
            "bentUpArrow" => Ok(Self::BentUpArrow),
            "leftRightArrow" => Ok(Self::LeftRightArrow),
            "upDownArrow" => Ok(Self::UpDownArrow),
            "leftUpArrow" => Ok(Self::LeftUpArrow),
            "leftRightUpArrow" => Ok(Self::LeftRightUpArrow),
            "quadArrow" => Ok(Self::QuadArrow),
            "leftArrowCallout" => Ok(Self::LeftArrowCallout),
            "rightArrowCallout" => Ok(Self::RightArrowCallout),
            "upArrowCallout" => Ok(Self::UpArrowCallout),
            "downArrowCallout" => Ok(Self::DownArrowCallout),
            "leftRightArrowCallout" => Ok(Self::LeftRightArrowCallout),
            "upDownArrowCallout" => Ok(Self::UpDownArrowCallout),
            "quadArrowCallout" => Ok(Self::QuadArrowCallout),
            "bentArrow" => Ok(Self::BentArrow),
            "uturnArrow" => Ok(Self::UturnArrow),
            "circularArrow" => Ok(Self::CircularArrow),
            "leftCircularArrow" => Ok(Self::LeftCircularArrow),
            "leftRightCircularArrow" => Ok(Self::LeftRightCircularArrow),
            "curvedRightArrow" => Ok(Self::CurvedRightArrow),
            "curvedLeftArrow" => Ok(Self::CurvedLeftArrow),
            "curvedUpArrow" => Ok(Self::CurvedUpArrow),
            "curvedDownArrow" => Ok(Self::CurvedDownArrow),
            "swooshArrow" => Ok(Self::SwooshArrow),
            "cube" => Ok(Self::Cube),
            "can" => Ok(Self::Can),
            "lightningBolt" => Ok(Self::LightningBolt),
            "heart" => Ok(Self::Heart),
            "sun" => Ok(Self::Sun),
            "moon" => Ok(Self::Moon),
            "smileyFace" => Ok(Self::SmileyFace),
            "irregularSeal1" => Ok(Self::IrregularSeal1),
            "irregularSeal2" => Ok(Self::IrregularSeal2),
            "foldedCorner" => Ok(Self::FoldedCorner),
            "bevel" => Ok(Self::Bevel),
            "frame" => Ok(Self::Frame),
            "halfFrame" => Ok(Self::HalfFrame),
            "corner" => Ok(Self::Corner),
            "diagStripe" => Ok(Self::DiagStripe),
            "chord" => Ok(Self::Chord),
            "arc" => Ok(Self::Arc),
            "leftBracket" => Ok(Self::LeftBracket),
            "rightBracket" => Ok(Self::RightBracket),
            "leftBrace" => Ok(Self::LeftBrace),
            "rightBrace" => Ok(Self::RightBrace),
            "bracketPair" => Ok(Self::BracketPair),
            "bracePair" => Ok(Self::BracePair),
            "straightConnector1" => Ok(Self::StraightConnector1),
            "bentConnector2" => Ok(Self::BentConnector2),
            "bentConnector3" => Ok(Self::BentConnector3),
            "bentConnector4" => Ok(Self::BentConnector4),
            "bentConnector5" => Ok(Self::BentConnector5),
            "curvedConnector2" => Ok(Self::CurvedConnector2),
            "curvedConnector3" => Ok(Self::CurvedConnector3),
            "curvedConnector4" => Ok(Self::CurvedConnector4),
            "curvedConnector5" => Ok(Self::CurvedConnector5),
            "callout1" => Ok(Self::Callout1),
            "callout2" => Ok(Self::Callout2),
            "callout3" => Ok(Self::Callout3),
            "accentCallout1" => Ok(Self::AccentCallout1),
            "accentCallout2" => Ok(Self::AccentCallout2),
            "accentCallout3" => Ok(Self::AccentCallout3),
            "borderCallout1" => Ok(Self::BorderCallout1),
            "borderCallout2" => Ok(Self::BorderCallout2),
            "borderCallout3" => Ok(Self::BorderCallout3),
            "accentBorderCallout1" => Ok(Self::AccentBorderCallout1),
            "accentBorderCallout2" => Ok(Self::AccentBorderCallout2),
            "accentBorderCallout3" => Ok(Self::AccentBorderCallout3),
            "wedgeRectCallout" => Ok(Self::WedgeRectCallout),
            "wedgeRoundRectCallout" => Ok(Self::WedgeRoundRectCallout),
            "wedgeEllipseCallout" => Ok(Self::WedgeEllipseCallout),
            "cloudCallout" => Ok(Self::CloudCallout),
            "cloud" => Ok(Self::Cloud),
            "ribbon" => Ok(Self::Ribbon),
            "ribbon2" => Ok(Self::Ribbon2),
            "ellipseRibbon" => Ok(Self::EllipseRibbon),
            "ellipseRibbon2" => Ok(Self::EllipseRibbon2),
            "leftRightRibbon" => Ok(Self::LeftRightRibbon),
            "verticalScroll" => Ok(Self::VerticalScroll),
            "horizontalScroll" => Ok(Self::HorizontalScroll),
            "wave" => Ok(Self::Wave),
            "doubleWave" => Ok(Self::DoubleWave),
            "plus" => Ok(Self::Plus),
            "flowChartProcess" => Ok(Self::FlowChartProcess),
            "flowChartDecision" => Ok(Self::FlowChartDecision),
            "flowChartInputOutput" => Ok(Self::FlowChartInputOutput),
            "flowChartPredefinedProcess" => Ok(Self::FlowChartPredefinedProcess),
            "flowChartInternalStorage" => Ok(Self::FlowChartInternalStorage),
            "flowChartDocument" => Ok(Self::FlowChartDocument),
            "flowChartMultidocument" => Ok(Self::FlowChartMultidocument),
            "flowChartTerminator" => Ok(Self::FlowChartTerminator),
            "flowChartPreparation" => Ok(Self::FlowChartPreparation),
            "flowChartManualInput" => Ok(Self::FlowChartManualInput),
            "flowChartManualOperation" => Ok(Self::FlowChartManualOperation),
            "flowChartConnector" => Ok(Self::FlowChartConnector),
            "flowChartPunchedCard" => Ok(Self::FlowChartPunchedCard),
            "flowChartPunchedTape" => Ok(Self::FlowChartPunchedTape),
            "flowChartSummingJunction" => Ok(Self::FlowChartSummingJunction),
            "flowChartOr" => Ok(Self::FlowChartOr),
            "flowChartCollate" => Ok(Self::FlowChartCollate),
            "flowChartSort" => Ok(Self::FlowChartSort),
            "flowChartExtract" => Ok(Self::FlowChartExtract),
            "flowChartMerge" => Ok(Self::FlowChartMerge),
            "flowChartOfflineStorage" => Ok(Self::FlowChartOfflineStorage),
            "flowChartOnlineStorage" => Ok(Self::FlowChartOnlineStorage),
            "flowChartMagneticTape" => Ok(Self::FlowChartMagneticTape),
            "flowChartMagneticDisk" => Ok(Self::FlowChartMagneticDisk),
            "flowChartMagneticDrum" => Ok(Self::FlowChartMagneticDrum),
            "flowChartDisplay" => Ok(Self::FlowChartDisplay),
            "flowChartDelay" => Ok(Self::FlowChartDelay),
            "flowChartAlternateProcess" => Ok(Self::FlowChartAlternateProcess),
            "flowChartOffpageConnector" => Ok(Self::FlowChartOffpageConnector),
            "actionButtonBlank" => Ok(Self::ActionButtonBlank),
            "actionButtonHome" => Ok(Self::ActionButtonHome),
            "actionButtonHelp" => Ok(Self::ActionButtonHelp),
            "actionButtonInformation" => Ok(Self::ActionButtonInformation),
            "actionButtonForwardNext" => Ok(Self::ActionButtonForwardNext),
            "actionButtonBackPrevious" => Ok(Self::ActionButtonBackPrevious),
            "actionButtonEnd" => Ok(Self::ActionButtonEnd),
            "actionButtonBeginning" => Ok(Self::ActionButtonBeginning),
            "actionButtonReturn" => Ok(Self::ActionButtonReturn),
            "actionButtonDocument" => Ok(Self::ActionButtonDocument),
            "actionButtonSound" => Ok(Self::ActionButtonSound),
            "actionButtonMovie" => Ok(Self::ActionButtonMovie),
            "gear6" => Ok(Self::Gear6),
            "gear9" => Ok(Self::Gear9),
            "funnel" => Ok(Self::Funnel),
            "mathPlus" => Ok(Self::MathPlus),
            "mathMinus" => Ok(Self::MathMinus),
            "mathMultiply" => Ok(Self::MathMultiply),
            "mathDivide" => Ok(Self::MathDivide),
            "mathEqual" => Ok(Self::MathEqual),
            "mathNotEqual" => Ok(Self::MathNotEqual),
            "cornerTabs" => Ok(Self::CornerTabs),
            "squareTabs" => Ok(Self::SquareTabs),
            "plaqueTabs" => Ok(Self::PlaqueTabs),
            "chartX" => Ok(Self::ChartX),
            "chartStar" => Ok(Self::ChartStar),
            "chartPlus" => Ok(Self::ChartPlus),
            _ => Err(format!("unknown STShapeType value: {}", s)),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum STTextShapeType {
    #[serde(rename = "textNoShape")]
    TextNoShape,
    #[serde(rename = "textPlain")]
    TextPlain,
    #[serde(rename = "textStop")]
    TextStop,
    #[serde(rename = "textTriangle")]
    TextTriangle,
    #[serde(rename = "textTriangleInverted")]
    TextTriangleInverted,
    #[serde(rename = "textChevron")]
    TextChevron,
    #[serde(rename = "textChevronInverted")]
    TextChevronInverted,
    #[serde(rename = "textRingInside")]
    TextRingInside,
    #[serde(rename = "textRingOutside")]
    TextRingOutside,
    #[serde(rename = "textArchUp")]
    TextArchUp,
    #[serde(rename = "textArchDown")]
    TextArchDown,
    #[serde(rename = "textCircle")]
    TextCircle,
    #[serde(rename = "textButton")]
    TextButton,
    #[serde(rename = "textArchUpPour")]
    TextArchUpPour,
    #[serde(rename = "textArchDownPour")]
    TextArchDownPour,
    #[serde(rename = "textCirclePour")]
    TextCirclePour,
    #[serde(rename = "textButtonPour")]
    TextButtonPour,
    #[serde(rename = "textCurveUp")]
    TextCurveUp,
    #[serde(rename = "textCurveDown")]
    TextCurveDown,
    #[serde(rename = "textCanUp")]
    TextCanUp,
    #[serde(rename = "textCanDown")]
    TextCanDown,
    #[serde(rename = "textWave1")]
    TextWave1,
    #[serde(rename = "textWave2")]
    TextWave2,
    #[serde(rename = "textDoubleWave1")]
    TextDoubleWave1,
    #[serde(rename = "textWave4")]
    TextWave4,
    #[serde(rename = "textInflate")]
    TextInflate,
    #[serde(rename = "textDeflate")]
    TextDeflate,
    #[serde(rename = "textInflateBottom")]
    TextInflateBottom,
    #[serde(rename = "textDeflateBottom")]
    TextDeflateBottom,
    #[serde(rename = "textInflateTop")]
    TextInflateTop,
    #[serde(rename = "textDeflateTop")]
    TextDeflateTop,
    #[serde(rename = "textDeflateInflate")]
    TextDeflateInflate,
    #[serde(rename = "textDeflateInflateDeflate")]
    TextDeflateInflateDeflate,
    #[serde(rename = "textFadeRight")]
    TextFadeRight,
    #[serde(rename = "textFadeLeft")]
    TextFadeLeft,
    #[serde(rename = "textFadeUp")]
    TextFadeUp,
    #[serde(rename = "textFadeDown")]
    TextFadeDown,
    #[serde(rename = "textSlantUp")]
    TextSlantUp,
    #[serde(rename = "textSlantDown")]
    TextSlantDown,
    #[serde(rename = "textCascadeUp")]
    TextCascadeUp,
    #[serde(rename = "textCascadeDown")]
    TextCascadeDown,
}

impl std::fmt::Display for STTextShapeType {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::TextNoShape => write!(f, "textNoShape"),
            Self::TextPlain => write!(f, "textPlain"),
            Self::TextStop => write!(f, "textStop"),
            Self::TextTriangle => write!(f, "textTriangle"),
            Self::TextTriangleInverted => write!(f, "textTriangleInverted"),
            Self::TextChevron => write!(f, "textChevron"),
            Self::TextChevronInverted => write!(f, "textChevronInverted"),
            Self::TextRingInside => write!(f, "textRingInside"),
            Self::TextRingOutside => write!(f, "textRingOutside"),
            Self::TextArchUp => write!(f, "textArchUp"),
            Self::TextArchDown => write!(f, "textArchDown"),
            Self::TextCircle => write!(f, "textCircle"),
            Self::TextButton => write!(f, "textButton"),
            Self::TextArchUpPour => write!(f, "textArchUpPour"),
            Self::TextArchDownPour => write!(f, "textArchDownPour"),
            Self::TextCirclePour => write!(f, "textCirclePour"),
            Self::TextButtonPour => write!(f, "textButtonPour"),
            Self::TextCurveUp => write!(f, "textCurveUp"),
            Self::TextCurveDown => write!(f, "textCurveDown"),
            Self::TextCanUp => write!(f, "textCanUp"),
            Self::TextCanDown => write!(f, "textCanDown"),
            Self::TextWave1 => write!(f, "textWave1"),
            Self::TextWave2 => write!(f, "textWave2"),
            Self::TextDoubleWave1 => write!(f, "textDoubleWave1"),
            Self::TextWave4 => write!(f, "textWave4"),
            Self::TextInflate => write!(f, "textInflate"),
            Self::TextDeflate => write!(f, "textDeflate"),
            Self::TextInflateBottom => write!(f, "textInflateBottom"),
            Self::TextDeflateBottom => write!(f, "textDeflateBottom"),
            Self::TextInflateTop => write!(f, "textInflateTop"),
            Self::TextDeflateTop => write!(f, "textDeflateTop"),
            Self::TextDeflateInflate => write!(f, "textDeflateInflate"),
            Self::TextDeflateInflateDeflate => write!(f, "textDeflateInflateDeflate"),
            Self::TextFadeRight => write!(f, "textFadeRight"),
            Self::TextFadeLeft => write!(f, "textFadeLeft"),
            Self::TextFadeUp => write!(f, "textFadeUp"),
            Self::TextFadeDown => write!(f, "textFadeDown"),
            Self::TextSlantUp => write!(f, "textSlantUp"),
            Self::TextSlantDown => write!(f, "textSlantDown"),
            Self::TextCascadeUp => write!(f, "textCascadeUp"),
            Self::TextCascadeDown => write!(f, "textCascadeDown"),
        }
    }
}

impl std::str::FromStr for STTextShapeType {
    type Err = String;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        match s {
            "textNoShape" => Ok(Self::TextNoShape),
            "textPlain" => Ok(Self::TextPlain),
            "textStop" => Ok(Self::TextStop),
            "textTriangle" => Ok(Self::TextTriangle),
            "textTriangleInverted" => Ok(Self::TextTriangleInverted),
            "textChevron" => Ok(Self::TextChevron),
            "textChevronInverted" => Ok(Self::TextChevronInverted),
            "textRingInside" => Ok(Self::TextRingInside),
            "textRingOutside" => Ok(Self::TextRingOutside),
            "textArchUp" => Ok(Self::TextArchUp),
            "textArchDown" => Ok(Self::TextArchDown),
            "textCircle" => Ok(Self::TextCircle),
            "textButton" => Ok(Self::TextButton),
            "textArchUpPour" => Ok(Self::TextArchUpPour),
            "textArchDownPour" => Ok(Self::TextArchDownPour),
            "textCirclePour" => Ok(Self::TextCirclePour),
            "textButtonPour" => Ok(Self::TextButtonPour),
            "textCurveUp" => Ok(Self::TextCurveUp),
            "textCurveDown" => Ok(Self::TextCurveDown),
            "textCanUp" => Ok(Self::TextCanUp),
            "textCanDown" => Ok(Self::TextCanDown),
            "textWave1" => Ok(Self::TextWave1),
            "textWave2" => Ok(Self::TextWave2),
            "textDoubleWave1" => Ok(Self::TextDoubleWave1),
            "textWave4" => Ok(Self::TextWave4),
            "textInflate" => Ok(Self::TextInflate),
            "textDeflate" => Ok(Self::TextDeflate),
            "textInflateBottom" => Ok(Self::TextInflateBottom),
            "textDeflateBottom" => Ok(Self::TextDeflateBottom),
            "textInflateTop" => Ok(Self::TextInflateTop),
            "textDeflateTop" => Ok(Self::TextDeflateTop),
            "textDeflateInflate" => Ok(Self::TextDeflateInflate),
            "textDeflateInflateDeflate" => Ok(Self::TextDeflateInflateDeflate),
            "textFadeRight" => Ok(Self::TextFadeRight),
            "textFadeLeft" => Ok(Self::TextFadeLeft),
            "textFadeUp" => Ok(Self::TextFadeUp),
            "textFadeDown" => Ok(Self::TextFadeDown),
            "textSlantUp" => Ok(Self::TextSlantUp),
            "textSlantDown" => Ok(Self::TextSlantDown),
            "textCascadeUp" => Ok(Self::TextCascadeUp),
            "textCascadeDown" => Ok(Self::TextCascadeDown),
            _ => Err(format!("unknown STTextShapeType value: {}", s)),
        }
    }
}

pub type STGeomGuideName = String;

pub type STGeomGuideFormula = String;

pub type STAdjCoordinate = String;

pub type STAdjAngle = String;

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum STPathFillMode {
    #[serde(rename = "none")]
    None,
    #[serde(rename = "norm")]
    Norm,
    #[serde(rename = "lighten")]
    Lighten,
    #[serde(rename = "lightenLess")]
    LightenLess,
    #[serde(rename = "darken")]
    Darken,
    #[serde(rename = "darkenLess")]
    DarkenLess,
}

impl std::fmt::Display for STPathFillMode {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::None => write!(f, "none"),
            Self::Norm => write!(f, "norm"),
            Self::Lighten => write!(f, "lighten"),
            Self::LightenLess => write!(f, "lightenLess"),
            Self::Darken => write!(f, "darken"),
            Self::DarkenLess => write!(f, "darkenLess"),
        }
    }
}

impl std::str::FromStr for STPathFillMode {
    type Err = String;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        match s {
            "none" => Ok(Self::None),
            "norm" => Ok(Self::Norm),
            "lighten" => Ok(Self::Lighten),
            "lightenLess" => Ok(Self::LightenLess),
            "darken" => Ok(Self::Darken),
            "darkenLess" => Ok(Self::DarkenLess),
            _ => Err(format!("unknown STPathFillMode value: {}", s)),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum STLineEndType {
    #[serde(rename = "none")]
    None,
    #[serde(rename = "triangle")]
    Triangle,
    #[serde(rename = "stealth")]
    Stealth,
    #[serde(rename = "diamond")]
    Diamond,
    #[serde(rename = "oval")]
    Oval,
    #[serde(rename = "arrow")]
    Arrow,
}

impl std::fmt::Display for STLineEndType {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::None => write!(f, "none"),
            Self::Triangle => write!(f, "triangle"),
            Self::Stealth => write!(f, "stealth"),
            Self::Diamond => write!(f, "diamond"),
            Self::Oval => write!(f, "oval"),
            Self::Arrow => write!(f, "arrow"),
        }
    }
}

impl std::str::FromStr for STLineEndType {
    type Err = String;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        match s {
            "none" => Ok(Self::None),
            "triangle" => Ok(Self::Triangle),
            "stealth" => Ok(Self::Stealth),
            "diamond" => Ok(Self::Diamond),
            "oval" => Ok(Self::Oval),
            "arrow" => Ok(Self::Arrow),
            _ => Err(format!("unknown STLineEndType value: {}", s)),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum STLineEndWidth {
    #[serde(rename = "sm")]
    Sm,
    #[serde(rename = "med")]
    Med,
    #[serde(rename = "lg")]
    Lg,
}

impl std::fmt::Display for STLineEndWidth {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::Sm => write!(f, "sm"),
            Self::Med => write!(f, "med"),
            Self::Lg => write!(f, "lg"),
        }
    }
}

impl std::str::FromStr for STLineEndWidth {
    type Err = String;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        match s {
            "sm" => Ok(Self::Sm),
            "med" => Ok(Self::Med),
            "lg" => Ok(Self::Lg),
            _ => Err(format!("unknown STLineEndWidth value: {}", s)),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum STLineEndLength {
    #[serde(rename = "sm")]
    Sm,
    #[serde(rename = "med")]
    Med,
    #[serde(rename = "lg")]
    Lg,
}

impl std::fmt::Display for STLineEndLength {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::Sm => write!(f, "sm"),
            Self::Med => write!(f, "med"),
            Self::Lg => write!(f, "lg"),
        }
    }
}

impl std::str::FromStr for STLineEndLength {
    type Err = String;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        match s {
            "sm" => Ok(Self::Sm),
            "med" => Ok(Self::Med),
            "lg" => Ok(Self::Lg),
            _ => Err(format!("unknown STLineEndLength value: {}", s)),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum STPresetLineDashVal {
    #[serde(rename = "solid")]
    Solid,
    #[serde(rename = "dot")]
    Dot,
    #[serde(rename = "dash")]
    Dash,
    #[serde(rename = "lgDash")]
    LgDash,
    #[serde(rename = "dashDot")]
    DashDot,
    #[serde(rename = "lgDashDot")]
    LgDashDot,
    #[serde(rename = "lgDashDotDot")]
    LgDashDotDot,
    #[serde(rename = "sysDash")]
    SysDash,
    #[serde(rename = "sysDot")]
    SysDot,
    #[serde(rename = "sysDashDot")]
    SysDashDot,
    #[serde(rename = "sysDashDotDot")]
    SysDashDotDot,
}

impl std::fmt::Display for STPresetLineDashVal {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::Solid => write!(f, "solid"),
            Self::Dot => write!(f, "dot"),
            Self::Dash => write!(f, "dash"),
            Self::LgDash => write!(f, "lgDash"),
            Self::DashDot => write!(f, "dashDot"),
            Self::LgDashDot => write!(f, "lgDashDot"),
            Self::LgDashDotDot => write!(f, "lgDashDotDot"),
            Self::SysDash => write!(f, "sysDash"),
            Self::SysDot => write!(f, "sysDot"),
            Self::SysDashDot => write!(f, "sysDashDot"),
            Self::SysDashDotDot => write!(f, "sysDashDotDot"),
        }
    }
}

impl std::str::FromStr for STPresetLineDashVal {
    type Err = String;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        match s {
            "solid" => Ok(Self::Solid),
            "dot" => Ok(Self::Dot),
            "dash" => Ok(Self::Dash),
            "lgDash" => Ok(Self::LgDash),
            "dashDot" => Ok(Self::DashDot),
            "lgDashDot" => Ok(Self::LgDashDot),
            "lgDashDotDot" => Ok(Self::LgDashDotDot),
            "sysDash" => Ok(Self::SysDash),
            "sysDot" => Ok(Self::SysDot),
            "sysDashDot" => Ok(Self::SysDashDot),
            "sysDashDotDot" => Ok(Self::SysDashDotDot),
            _ => Err(format!("unknown STPresetLineDashVal value: {}", s)),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum STLineCap {
    #[serde(rename = "rnd")]
    Rnd,
    #[serde(rename = "sq")]
    Sq,
    #[serde(rename = "flat")]
    Flat,
}

impl std::fmt::Display for STLineCap {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::Rnd => write!(f, "rnd"),
            Self::Sq => write!(f, "sq"),
            Self::Flat => write!(f, "flat"),
        }
    }
}

impl std::str::FromStr for STLineCap {
    type Err = String;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        match s {
            "rnd" => Ok(Self::Rnd),
            "sq" => Ok(Self::Sq),
            "flat" => Ok(Self::Flat),
            _ => Err(format!("unknown STLineCap value: {}", s)),
        }
    }
}

pub type STLineWidth = i32;

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum STPenAlignment {
    #[serde(rename = "ctr")]
    Ctr,
    #[serde(rename = "in")]
    In,
}

impl std::fmt::Display for STPenAlignment {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::Ctr => write!(f, "ctr"),
            Self::In => write!(f, "in"),
        }
    }
}

impl std::str::FromStr for STPenAlignment {
    type Err = String;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        match s {
            "ctr" => Ok(Self::Ctr),
            "in" => Ok(Self::In),
            _ => Err(format!("unknown STPenAlignment value: {}", s)),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum STCompoundLine {
    #[serde(rename = "sng")]
    Sng,
    #[serde(rename = "dbl")]
    Dbl,
    #[serde(rename = "thickThin")]
    ThickThin,
    #[serde(rename = "thinThick")]
    ThinThick,
    #[serde(rename = "tri")]
    Tri,
}

impl std::fmt::Display for STCompoundLine {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::Sng => write!(f, "sng"),
            Self::Dbl => write!(f, "dbl"),
            Self::ThickThin => write!(f, "thickThin"),
            Self::ThinThick => write!(f, "thinThick"),
            Self::Tri => write!(f, "tri"),
        }
    }
}

impl std::str::FromStr for STCompoundLine {
    type Err = String;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        match s {
            "sng" => Ok(Self::Sng),
            "dbl" => Ok(Self::Dbl),
            "thickThin" => Ok(Self::ThickThin),
            "thinThick" => Ok(Self::ThinThick),
            "tri" => Ok(Self::Tri),
            _ => Err(format!("unknown STCompoundLine value: {}", s)),
        }
    }
}

pub type STShapeID = String;

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum STOnOffStyleType {
    #[serde(rename = "on")]
    On,
    #[serde(rename = "off")]
    Off,
    #[serde(rename = "def")]
    Def,
}

impl std::fmt::Display for STOnOffStyleType {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::On => write!(f, "on"),
            Self::Off => write!(f, "off"),
            Self::Def => write!(f, "def"),
        }
    }
}

impl std::str::FromStr for STOnOffStyleType {
    type Err = String;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        match s {
            "on" => Ok(Self::On),
            "off" => Ok(Self::Off),
            "def" => Ok(Self::Def),
            _ => Err(format!("unknown STOnOffStyleType value: {}", s)),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum STTextAnchoringType {
    #[serde(rename = "t")]
    T,
    #[serde(rename = "ctr")]
    Ctr,
    #[serde(rename = "b")]
    B,
    #[serde(rename = "just")]
    Just,
    #[serde(rename = "dist")]
    Dist,
}

impl std::fmt::Display for STTextAnchoringType {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::T => write!(f, "t"),
            Self::Ctr => write!(f, "ctr"),
            Self::B => write!(f, "b"),
            Self::Just => write!(f, "just"),
            Self::Dist => write!(f, "dist"),
        }
    }
}

impl std::str::FromStr for STTextAnchoringType {
    type Err = String;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        match s {
            "t" => Ok(Self::T),
            "ctr" => Ok(Self::Ctr),
            "b" => Ok(Self::B),
            "just" => Ok(Self::Just),
            "dist" => Ok(Self::Dist),
            _ => Err(format!("unknown STTextAnchoringType value: {}", s)),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum STTextVertOverflowType {
    #[serde(rename = "overflow")]
    Overflow,
    #[serde(rename = "ellipsis")]
    Ellipsis,
    #[serde(rename = "clip")]
    Clip,
}

impl std::fmt::Display for STTextVertOverflowType {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::Overflow => write!(f, "overflow"),
            Self::Ellipsis => write!(f, "ellipsis"),
            Self::Clip => write!(f, "clip"),
        }
    }
}

impl std::str::FromStr for STTextVertOverflowType {
    type Err = String;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        match s {
            "overflow" => Ok(Self::Overflow),
            "ellipsis" => Ok(Self::Ellipsis),
            "clip" => Ok(Self::Clip),
            _ => Err(format!("unknown STTextVertOverflowType value: {}", s)),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum STTextHorzOverflowType {
    #[serde(rename = "overflow")]
    Overflow,
    #[serde(rename = "clip")]
    Clip,
}

impl std::fmt::Display for STTextHorzOverflowType {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::Overflow => write!(f, "overflow"),
            Self::Clip => write!(f, "clip"),
        }
    }
}

impl std::str::FromStr for STTextHorzOverflowType {
    type Err = String;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        match s {
            "overflow" => Ok(Self::Overflow),
            "clip" => Ok(Self::Clip),
            _ => Err(format!("unknown STTextHorzOverflowType value: {}", s)),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum STTextVerticalType {
    #[serde(rename = "horz")]
    Horz,
    #[serde(rename = "vert")]
    Vert,
    #[serde(rename = "vert270")]
    Vert270,
    #[serde(rename = "wordArtVert")]
    WordArtVert,
    #[serde(rename = "eaVert")]
    EaVert,
    #[serde(rename = "mongolianVert")]
    MongolianVert,
    #[serde(rename = "wordArtVertRtl")]
    WordArtVertRtl,
}

impl std::fmt::Display for STTextVerticalType {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::Horz => write!(f, "horz"),
            Self::Vert => write!(f, "vert"),
            Self::Vert270 => write!(f, "vert270"),
            Self::WordArtVert => write!(f, "wordArtVert"),
            Self::EaVert => write!(f, "eaVert"),
            Self::MongolianVert => write!(f, "mongolianVert"),
            Self::WordArtVertRtl => write!(f, "wordArtVertRtl"),
        }
    }
}

impl std::str::FromStr for STTextVerticalType {
    type Err = String;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        match s {
            "horz" => Ok(Self::Horz),
            "vert" => Ok(Self::Vert),
            "vert270" => Ok(Self::Vert270),
            "wordArtVert" => Ok(Self::WordArtVert),
            "eaVert" => Ok(Self::EaVert),
            "mongolianVert" => Ok(Self::MongolianVert),
            "wordArtVertRtl" => Ok(Self::WordArtVertRtl),
            _ => Err(format!("unknown STTextVerticalType value: {}", s)),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum STTextWrappingType {
    #[serde(rename = "none")]
    None,
    #[serde(rename = "square")]
    Square,
}

impl std::fmt::Display for STTextWrappingType {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::None => write!(f, "none"),
            Self::Square => write!(f, "square"),
        }
    }
}

impl std::str::FromStr for STTextWrappingType {
    type Err = String;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        match s {
            "none" => Ok(Self::None),
            "square" => Ok(Self::Square),
            _ => Err(format!("unknown STTextWrappingType value: {}", s)),
        }
    }
}

pub type STTextColumnCount = i32;

pub type STTextFontScalePercentOrPercentString = String;

pub type STTextFontScalePercent = i32;

pub type STTextBulletStartAtNum = i32;

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum STTextAutonumberScheme {
    #[serde(rename = "alphaLcParenBoth")]
    AlphaLcParenBoth,
    #[serde(rename = "alphaUcParenBoth")]
    AlphaUcParenBoth,
    #[serde(rename = "alphaLcParenR")]
    AlphaLcParenR,
    #[serde(rename = "alphaUcParenR")]
    AlphaUcParenR,
    #[serde(rename = "alphaLcPeriod")]
    AlphaLcPeriod,
    #[serde(rename = "alphaUcPeriod")]
    AlphaUcPeriod,
    #[serde(rename = "arabicParenBoth")]
    ArabicParenBoth,
    #[serde(rename = "arabicParenR")]
    ArabicParenR,
    #[serde(rename = "arabicPeriod")]
    ArabicPeriod,
    #[serde(rename = "arabicPlain")]
    ArabicPlain,
    #[serde(rename = "romanLcParenBoth")]
    RomanLcParenBoth,
    #[serde(rename = "romanUcParenBoth")]
    RomanUcParenBoth,
    #[serde(rename = "romanLcParenR")]
    RomanLcParenR,
    #[serde(rename = "romanUcParenR")]
    RomanUcParenR,
    #[serde(rename = "romanLcPeriod")]
    RomanLcPeriod,
    #[serde(rename = "romanUcPeriod")]
    RomanUcPeriod,
    #[serde(rename = "circleNumDbPlain")]
    CircleNumDbPlain,
    #[serde(rename = "circleNumWdBlackPlain")]
    CircleNumWdBlackPlain,
    #[serde(rename = "circleNumWdWhitePlain")]
    CircleNumWdWhitePlain,
    #[serde(rename = "arabicDbPeriod")]
    ArabicDbPeriod,
    #[serde(rename = "arabicDbPlain")]
    ArabicDbPlain,
    #[serde(rename = "ea1ChsPeriod")]
    Ea1ChsPeriod,
    #[serde(rename = "ea1ChsPlain")]
    Ea1ChsPlain,
    #[serde(rename = "ea1ChtPeriod")]
    Ea1ChtPeriod,
    #[serde(rename = "ea1ChtPlain")]
    Ea1ChtPlain,
    #[serde(rename = "ea1JpnChsDbPeriod")]
    Ea1JpnChsDbPeriod,
    #[serde(rename = "ea1JpnKorPlain")]
    Ea1JpnKorPlain,
    #[serde(rename = "ea1JpnKorPeriod")]
    Ea1JpnKorPeriod,
    #[serde(rename = "arabic1Minus")]
    Arabic1Minus,
    #[serde(rename = "arabic2Minus")]
    Arabic2Minus,
    #[serde(rename = "hebrew2Minus")]
    Hebrew2Minus,
    #[serde(rename = "thaiAlphaPeriod")]
    ThaiAlphaPeriod,
    #[serde(rename = "thaiAlphaParenR")]
    ThaiAlphaParenR,
    #[serde(rename = "thaiAlphaParenBoth")]
    ThaiAlphaParenBoth,
    #[serde(rename = "thaiNumPeriod")]
    ThaiNumPeriod,
    #[serde(rename = "thaiNumParenR")]
    ThaiNumParenR,
    #[serde(rename = "thaiNumParenBoth")]
    ThaiNumParenBoth,
    #[serde(rename = "hindiAlphaPeriod")]
    HindiAlphaPeriod,
    #[serde(rename = "hindiNumPeriod")]
    HindiNumPeriod,
    #[serde(rename = "hindiNumParenR")]
    HindiNumParenR,
    #[serde(rename = "hindiAlpha1Period")]
    HindiAlpha1Period,
}

impl std::fmt::Display for STTextAutonumberScheme {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::AlphaLcParenBoth => write!(f, "alphaLcParenBoth"),
            Self::AlphaUcParenBoth => write!(f, "alphaUcParenBoth"),
            Self::AlphaLcParenR => write!(f, "alphaLcParenR"),
            Self::AlphaUcParenR => write!(f, "alphaUcParenR"),
            Self::AlphaLcPeriod => write!(f, "alphaLcPeriod"),
            Self::AlphaUcPeriod => write!(f, "alphaUcPeriod"),
            Self::ArabicParenBoth => write!(f, "arabicParenBoth"),
            Self::ArabicParenR => write!(f, "arabicParenR"),
            Self::ArabicPeriod => write!(f, "arabicPeriod"),
            Self::ArabicPlain => write!(f, "arabicPlain"),
            Self::RomanLcParenBoth => write!(f, "romanLcParenBoth"),
            Self::RomanUcParenBoth => write!(f, "romanUcParenBoth"),
            Self::RomanLcParenR => write!(f, "romanLcParenR"),
            Self::RomanUcParenR => write!(f, "romanUcParenR"),
            Self::RomanLcPeriod => write!(f, "romanLcPeriod"),
            Self::RomanUcPeriod => write!(f, "romanUcPeriod"),
            Self::CircleNumDbPlain => write!(f, "circleNumDbPlain"),
            Self::CircleNumWdBlackPlain => write!(f, "circleNumWdBlackPlain"),
            Self::CircleNumWdWhitePlain => write!(f, "circleNumWdWhitePlain"),
            Self::ArabicDbPeriod => write!(f, "arabicDbPeriod"),
            Self::ArabicDbPlain => write!(f, "arabicDbPlain"),
            Self::Ea1ChsPeriod => write!(f, "ea1ChsPeriod"),
            Self::Ea1ChsPlain => write!(f, "ea1ChsPlain"),
            Self::Ea1ChtPeriod => write!(f, "ea1ChtPeriod"),
            Self::Ea1ChtPlain => write!(f, "ea1ChtPlain"),
            Self::Ea1JpnChsDbPeriod => write!(f, "ea1JpnChsDbPeriod"),
            Self::Ea1JpnKorPlain => write!(f, "ea1JpnKorPlain"),
            Self::Ea1JpnKorPeriod => write!(f, "ea1JpnKorPeriod"),
            Self::Arabic1Minus => write!(f, "arabic1Minus"),
            Self::Arabic2Minus => write!(f, "arabic2Minus"),
            Self::Hebrew2Minus => write!(f, "hebrew2Minus"),
            Self::ThaiAlphaPeriod => write!(f, "thaiAlphaPeriod"),
            Self::ThaiAlphaParenR => write!(f, "thaiAlphaParenR"),
            Self::ThaiAlphaParenBoth => write!(f, "thaiAlphaParenBoth"),
            Self::ThaiNumPeriod => write!(f, "thaiNumPeriod"),
            Self::ThaiNumParenR => write!(f, "thaiNumParenR"),
            Self::ThaiNumParenBoth => write!(f, "thaiNumParenBoth"),
            Self::HindiAlphaPeriod => write!(f, "hindiAlphaPeriod"),
            Self::HindiNumPeriod => write!(f, "hindiNumPeriod"),
            Self::HindiNumParenR => write!(f, "hindiNumParenR"),
            Self::HindiAlpha1Period => write!(f, "hindiAlpha1Period"),
        }
    }
}

impl std::str::FromStr for STTextAutonumberScheme {
    type Err = String;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        match s {
            "alphaLcParenBoth" => Ok(Self::AlphaLcParenBoth),
            "alphaUcParenBoth" => Ok(Self::AlphaUcParenBoth),
            "alphaLcParenR" => Ok(Self::AlphaLcParenR),
            "alphaUcParenR" => Ok(Self::AlphaUcParenR),
            "alphaLcPeriod" => Ok(Self::AlphaLcPeriod),
            "alphaUcPeriod" => Ok(Self::AlphaUcPeriod),
            "arabicParenBoth" => Ok(Self::ArabicParenBoth),
            "arabicParenR" => Ok(Self::ArabicParenR),
            "arabicPeriod" => Ok(Self::ArabicPeriod),
            "arabicPlain" => Ok(Self::ArabicPlain),
            "romanLcParenBoth" => Ok(Self::RomanLcParenBoth),
            "romanUcParenBoth" => Ok(Self::RomanUcParenBoth),
            "romanLcParenR" => Ok(Self::RomanLcParenR),
            "romanUcParenR" => Ok(Self::RomanUcParenR),
            "romanLcPeriod" => Ok(Self::RomanLcPeriod),
            "romanUcPeriod" => Ok(Self::RomanUcPeriod),
            "circleNumDbPlain" => Ok(Self::CircleNumDbPlain),
            "circleNumWdBlackPlain" => Ok(Self::CircleNumWdBlackPlain),
            "circleNumWdWhitePlain" => Ok(Self::CircleNumWdWhitePlain),
            "arabicDbPeriod" => Ok(Self::ArabicDbPeriod),
            "arabicDbPlain" => Ok(Self::ArabicDbPlain),
            "ea1ChsPeriod" => Ok(Self::Ea1ChsPeriod),
            "ea1ChsPlain" => Ok(Self::Ea1ChsPlain),
            "ea1ChtPeriod" => Ok(Self::Ea1ChtPeriod),
            "ea1ChtPlain" => Ok(Self::Ea1ChtPlain),
            "ea1JpnChsDbPeriod" => Ok(Self::Ea1JpnChsDbPeriod),
            "ea1JpnKorPlain" => Ok(Self::Ea1JpnKorPlain),
            "ea1JpnKorPeriod" => Ok(Self::Ea1JpnKorPeriod),
            "arabic1Minus" => Ok(Self::Arabic1Minus),
            "arabic2Minus" => Ok(Self::Arabic2Minus),
            "hebrew2Minus" => Ok(Self::Hebrew2Minus),
            "thaiAlphaPeriod" => Ok(Self::ThaiAlphaPeriod),
            "thaiAlphaParenR" => Ok(Self::ThaiAlphaParenR),
            "thaiAlphaParenBoth" => Ok(Self::ThaiAlphaParenBoth),
            "thaiNumPeriod" => Ok(Self::ThaiNumPeriod),
            "thaiNumParenR" => Ok(Self::ThaiNumParenR),
            "thaiNumParenBoth" => Ok(Self::ThaiNumParenBoth),
            "hindiAlphaPeriod" => Ok(Self::HindiAlphaPeriod),
            "hindiNumPeriod" => Ok(Self::HindiNumPeriod),
            "hindiNumParenR" => Ok(Self::HindiNumParenR),
            "hindiAlpha1Period" => Ok(Self::HindiAlpha1Period),
            _ => Err(format!("unknown STTextAutonumberScheme value: {}", s)),
        }
    }
}

pub type STTextBulletSize = String;

pub type STTextBulletSizePercent = String;

pub type STTextBulletSizeDecimal = i32;

pub type STTextPoint = String;

pub type STTextPointUnqualified = i32;

pub type STTextNonNegativePoint = i32;

pub type STTextFontSize = i32;

pub type STTextTypeface = String;

pub type STPitchFamily = String;

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum STTextUnderlineType {
    #[serde(rename = "none")]
    None,
    #[serde(rename = "words")]
    Words,
    #[serde(rename = "sng")]
    Sng,
    #[serde(rename = "dbl")]
    Dbl,
    #[serde(rename = "heavy")]
    Heavy,
    #[serde(rename = "dotted")]
    Dotted,
    #[serde(rename = "dottedHeavy")]
    DottedHeavy,
    #[serde(rename = "dash")]
    Dash,
    #[serde(rename = "dashHeavy")]
    DashHeavy,
    #[serde(rename = "dashLong")]
    DashLong,
    #[serde(rename = "dashLongHeavy")]
    DashLongHeavy,
    #[serde(rename = "dotDash")]
    DotDash,
    #[serde(rename = "dotDashHeavy")]
    DotDashHeavy,
    #[serde(rename = "dotDotDash")]
    DotDotDash,
    #[serde(rename = "dotDotDashHeavy")]
    DotDotDashHeavy,
    #[serde(rename = "wavy")]
    Wavy,
    #[serde(rename = "wavyHeavy")]
    WavyHeavy,
    #[serde(rename = "wavyDbl")]
    WavyDbl,
}

impl std::fmt::Display for STTextUnderlineType {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::None => write!(f, "none"),
            Self::Words => write!(f, "words"),
            Self::Sng => write!(f, "sng"),
            Self::Dbl => write!(f, "dbl"),
            Self::Heavy => write!(f, "heavy"),
            Self::Dotted => write!(f, "dotted"),
            Self::DottedHeavy => write!(f, "dottedHeavy"),
            Self::Dash => write!(f, "dash"),
            Self::DashHeavy => write!(f, "dashHeavy"),
            Self::DashLong => write!(f, "dashLong"),
            Self::DashLongHeavy => write!(f, "dashLongHeavy"),
            Self::DotDash => write!(f, "dotDash"),
            Self::DotDashHeavy => write!(f, "dotDashHeavy"),
            Self::DotDotDash => write!(f, "dotDotDash"),
            Self::DotDotDashHeavy => write!(f, "dotDotDashHeavy"),
            Self::Wavy => write!(f, "wavy"),
            Self::WavyHeavy => write!(f, "wavyHeavy"),
            Self::WavyDbl => write!(f, "wavyDbl"),
        }
    }
}

impl std::str::FromStr for STTextUnderlineType {
    type Err = String;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        match s {
            "none" => Ok(Self::None),
            "words" => Ok(Self::Words),
            "sng" => Ok(Self::Sng),
            "dbl" => Ok(Self::Dbl),
            "heavy" => Ok(Self::Heavy),
            "dotted" => Ok(Self::Dotted),
            "dottedHeavy" => Ok(Self::DottedHeavy),
            "dash" => Ok(Self::Dash),
            "dashHeavy" => Ok(Self::DashHeavy),
            "dashLong" => Ok(Self::DashLong),
            "dashLongHeavy" => Ok(Self::DashLongHeavy),
            "dotDash" => Ok(Self::DotDash),
            "dotDashHeavy" => Ok(Self::DotDashHeavy),
            "dotDotDash" => Ok(Self::DotDotDash),
            "dotDotDashHeavy" => Ok(Self::DotDotDashHeavy),
            "wavy" => Ok(Self::Wavy),
            "wavyHeavy" => Ok(Self::WavyHeavy),
            "wavyDbl" => Ok(Self::WavyDbl),
            _ => Err(format!("unknown STTextUnderlineType value: {}", s)),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum STTextStrikeType {
    #[serde(rename = "noStrike")]
    NoStrike,
    #[serde(rename = "sngStrike")]
    SngStrike,
    #[serde(rename = "dblStrike")]
    DblStrike,
}

impl std::fmt::Display for STTextStrikeType {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::NoStrike => write!(f, "noStrike"),
            Self::SngStrike => write!(f, "sngStrike"),
            Self::DblStrike => write!(f, "dblStrike"),
        }
    }
}

impl std::str::FromStr for STTextStrikeType {
    type Err = String;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        match s {
            "noStrike" => Ok(Self::NoStrike),
            "sngStrike" => Ok(Self::SngStrike),
            "dblStrike" => Ok(Self::DblStrike),
            _ => Err(format!("unknown STTextStrikeType value: {}", s)),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum STTextCapsType {
    #[serde(rename = "none")]
    None,
    #[serde(rename = "small")]
    Small,
    #[serde(rename = "all")]
    All,
}

impl std::fmt::Display for STTextCapsType {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::None => write!(f, "none"),
            Self::Small => write!(f, "small"),
            Self::All => write!(f, "all"),
        }
    }
}

impl std::str::FromStr for STTextCapsType {
    type Err = String;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        match s {
            "none" => Ok(Self::None),
            "small" => Ok(Self::Small),
            "all" => Ok(Self::All),
            _ => Err(format!("unknown STTextCapsType value: {}", s)),
        }
    }
}

pub type STTextSpacingPoint = i32;

pub type STTextSpacingPercentOrPercentString = String;

pub type STTextSpacingPercent = i32;

pub type STTextMargin = i32;

pub type STTextIndent = i32;

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum STTextTabAlignType {
    #[serde(rename = "l")]
    L,
    #[serde(rename = "ctr")]
    Ctr,
    #[serde(rename = "r")]
    R,
    #[serde(rename = "dec")]
    Dec,
}

impl std::fmt::Display for STTextTabAlignType {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::L => write!(f, "l"),
            Self::Ctr => write!(f, "ctr"),
            Self::R => write!(f, "r"),
            Self::Dec => write!(f, "dec"),
        }
    }
}

impl std::str::FromStr for STTextTabAlignType {
    type Err = String;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        match s {
            "l" => Ok(Self::L),
            "ctr" => Ok(Self::Ctr),
            "r" => Ok(Self::R),
            "dec" => Ok(Self::Dec),
            _ => Err(format!("unknown STTextTabAlignType value: {}", s)),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum STTextAlignType {
    #[serde(rename = "l")]
    L,
    #[serde(rename = "ctr")]
    Ctr,
    #[serde(rename = "r")]
    R,
    #[serde(rename = "just")]
    Just,
    #[serde(rename = "justLow")]
    JustLow,
    #[serde(rename = "dist")]
    Dist,
    #[serde(rename = "thaiDist")]
    ThaiDist,
}

impl std::fmt::Display for STTextAlignType {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::L => write!(f, "l"),
            Self::Ctr => write!(f, "ctr"),
            Self::R => write!(f, "r"),
            Self::Just => write!(f, "just"),
            Self::JustLow => write!(f, "justLow"),
            Self::Dist => write!(f, "dist"),
            Self::ThaiDist => write!(f, "thaiDist"),
        }
    }
}

impl std::str::FromStr for STTextAlignType {
    type Err = String;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        match s {
            "l" => Ok(Self::L),
            "ctr" => Ok(Self::Ctr),
            "r" => Ok(Self::R),
            "just" => Ok(Self::Just),
            "justLow" => Ok(Self::JustLow),
            "dist" => Ok(Self::Dist),
            "thaiDist" => Ok(Self::ThaiDist),
            _ => Err(format!("unknown STTextAlignType value: {}", s)),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum STTextFontAlignType {
    #[serde(rename = "auto")]
    Auto,
    #[serde(rename = "t")]
    T,
    #[serde(rename = "ctr")]
    Ctr,
    #[serde(rename = "base")]
    Base,
    #[serde(rename = "b")]
    B,
}

impl std::fmt::Display for STTextFontAlignType {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::Auto => write!(f, "auto"),
            Self::T => write!(f, "t"),
            Self::Ctr => write!(f, "ctr"),
            Self::Base => write!(f, "base"),
            Self::B => write!(f, "b"),
        }
    }
}

impl std::str::FromStr for STTextFontAlignType {
    type Err = String;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        match s {
            "auto" => Ok(Self::Auto),
            "t" => Ok(Self::T),
            "ctr" => Ok(Self::Ctr),
            "base" => Ok(Self::Base),
            "b" => Ok(Self::B),
            _ => Err(format!("unknown STTextFontAlignType value: {}", s)),
        }
    }
}

pub type STTextIndentLevelType = i32;

#[derive(Debug, Clone, Serialize, Deserialize)]
pub enum EGMedia {
    #[serde(rename = "audioCd")]
    AudioCd(Box<CTAudioCD>),
    #[serde(rename = "wavAudioFile")]
    WavAudioFile(Box<CTEmbeddedWAVAudioFile>),
    #[serde(rename = "audioFile")]
    AudioFile(Box<CTAudioFile>),
    #[serde(rename = "videoFile")]
    VideoFile(Box<CTVideoFile>),
    #[serde(rename = "quickTimeFile")]
    QuickTimeFile(Box<CTQuickTimeFile>),
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub enum EGColorTransform {
    #[serde(rename = "tint")]
    Tint(Box<CTPositiveFixedPercentage>),
    #[serde(rename = "shade")]
    Shade(Box<CTPositiveFixedPercentage>),
    #[serde(rename = "comp")]
    Comp(Box<CTComplementTransform>),
    #[serde(rename = "inv")]
    Inv(Box<CTInverseTransform>),
    #[serde(rename = "gray")]
    Gray(Box<CTGrayscaleTransform>),
    #[serde(rename = "alpha")]
    Alpha(Box<CTPositiveFixedPercentage>),
    #[serde(rename = "alphaOff")]
    AlphaOff(Box<CTFixedPercentage>),
    #[serde(rename = "alphaMod")]
    AlphaMod(Box<CTPositivePercentage>),
    #[serde(rename = "hue")]
    Hue(Box<CTPositiveFixedAngle>),
    #[serde(rename = "hueOff")]
    HueOff(Box<CTAngle>),
    #[serde(rename = "hueMod")]
    HueMod(Box<CTPositivePercentage>),
    #[serde(rename = "sat")]
    Sat(Box<CTPercentage>),
    #[serde(rename = "satOff")]
    SatOff(Box<CTPercentage>),
    #[serde(rename = "satMod")]
    SatMod(Box<CTPercentage>),
    #[serde(rename = "lum")]
    Lum(Box<CTPercentage>),
    #[serde(rename = "lumOff")]
    LumOff(Box<CTPercentage>),
    #[serde(rename = "lumMod")]
    LumMod(Box<CTPercentage>),
    #[serde(rename = "red")]
    Red(Box<CTPercentage>),
    #[serde(rename = "redOff")]
    RedOff(Box<CTPercentage>),
    #[serde(rename = "redMod")]
    RedMod(Box<CTPercentage>),
    #[serde(rename = "green")]
    Green(Box<CTPercentage>),
    #[serde(rename = "greenOff")]
    GreenOff(Box<CTPercentage>),
    #[serde(rename = "greenMod")]
    GreenMod(Box<CTPercentage>),
    #[serde(rename = "blue")]
    Blue(Box<CTPercentage>),
    #[serde(rename = "blueOff")]
    BlueOff(Box<CTPercentage>),
    #[serde(rename = "blueMod")]
    BlueMod(Box<CTPercentage>),
    #[serde(rename = "gamma")]
    Gamma(Box<CTGammaTransform>),
    #[serde(rename = "invGamma")]
    InvGamma(Box<CTInverseGammaTransform>),
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub enum EGColorChoice {
    #[serde(rename = "scrgbClr")]
    ScrgbClr(Box<CTScRgbColor>),
    #[serde(rename = "srgbClr")]
    SrgbClr(Box<SrgbColor>),
    #[serde(rename = "hslClr")]
    HslClr(Box<HslColor>),
    #[serde(rename = "sysClr")]
    SysClr(Box<SystemColor>),
    #[serde(rename = "schemeClr")]
    SchemeClr(Box<SchemeColor>),
    #[serde(rename = "prstClr")]
    PrstClr(Box<PresetColor>),
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub enum EGText3D {
    #[serde(rename = "sp3d")]
    Sp3d(Box<CTShape3D>),
    #[serde(rename = "flatTx")]
    FlatTx(Box<CTFlatText>),
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub enum EGShadeProperties {
    #[serde(rename = "lin")]
    Lin(Box<CTLinearShadeProperties>),
    #[serde(rename = "path")]
    Path(Box<CTPathShadeProperties>),
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub enum EGFillModeProperties {
    #[serde(rename = "tile")]
    Tile(Box<CTTileInfoProperties>),
    #[serde(rename = "stretch")]
    Stretch(Box<CTStretchInfoProperties>),
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub enum EGFillProperties {
    #[serde(rename = "noFill")]
    NoFill(Box<NoFill>),
    #[serde(rename = "solidFill")]
    SolidFill(Box<SolidColorFill>),
    #[serde(rename = "gradFill")]
    GradFill(Box<GradientFill>),
    #[serde(rename = "blipFill")]
    BlipFill(Box<BlipFillProperties>),
    #[serde(rename = "pattFill")]
    PattFill(Box<PatternFill>),
    #[serde(rename = "grpFill")]
    GrpFill(Box<CTGroupFillProperties>),
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub enum EGEffect {
    #[serde(rename = "cont")]
    Cont(Box<EffectContainer>),
    #[serde(rename = "effect")]
    Effect(Box<CTEffectReference>),
    #[serde(rename = "alphaBiLevel")]
    AlphaBiLevel(Box<CTAlphaBiLevelEffect>),
    #[serde(rename = "alphaCeiling")]
    AlphaCeiling(Box<CTAlphaCeilingEffect>),
    #[serde(rename = "alphaFloor")]
    AlphaFloor(Box<CTAlphaFloorEffect>),
    #[serde(rename = "alphaInv")]
    AlphaInv(Box<CTAlphaInverseEffect>),
    #[serde(rename = "alphaMod")]
    AlphaMod(Box<CTAlphaModulateEffect>),
    #[serde(rename = "alphaModFix")]
    AlphaModFix(Box<CTAlphaModulateFixedEffect>),
    #[serde(rename = "alphaOutset")]
    AlphaOutset(Box<CTAlphaOutsetEffect>),
    #[serde(rename = "alphaRepl")]
    AlphaRepl(Box<CTAlphaReplaceEffect>),
    #[serde(rename = "biLevel")]
    BiLevel(Box<CTBiLevelEffect>),
    #[serde(rename = "blend")]
    Blend(Box<CTBlendEffect>),
    #[serde(rename = "blur")]
    Blur(Box<CTBlurEffect>),
    #[serde(rename = "clrChange")]
    ClrChange(Box<CTColorChangeEffect>),
    #[serde(rename = "clrRepl")]
    ClrRepl(Box<CTColorReplaceEffect>),
    #[serde(rename = "duotone")]
    Duotone(Box<CTDuotoneEffect>),
    #[serde(rename = "fill")]
    Fill(Box<CTFillEffect>),
    #[serde(rename = "fillOverlay")]
    FillOverlay(Box<CTFillOverlayEffect>),
    #[serde(rename = "glow")]
    Glow(Box<CTGlowEffect>),
    #[serde(rename = "grayscl")]
    Grayscl(Box<CTGrayscaleEffect>),
    #[serde(rename = "hsl")]
    Hsl(Box<CTHSLEffect>),
    #[serde(rename = "innerShdw")]
    InnerShdw(Box<CTInnerShadowEffect>),
    #[serde(rename = "lum")]
    Lum(Box<CTLuminanceEffect>),
    #[serde(rename = "outerShdw")]
    OuterShdw(Box<CTOuterShadowEffect>),
    #[serde(rename = "prstShdw")]
    PrstShdw(Box<CTPresetShadowEffect>),
    #[serde(rename = "reflection")]
    Reflection(Box<CTReflectionEffect>),
    #[serde(rename = "relOff")]
    RelOff(Box<CTRelativeOffsetEffect>),
    #[serde(rename = "softEdge")]
    SoftEdge(Box<CTSoftEdgesEffect>),
    #[serde(rename = "tint")]
    Tint(Box<CTTintEffect>),
    #[serde(rename = "xfrm")]
    Xfrm(Box<CTTransformEffect>),
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub enum EGEffectProperties {
    #[serde(rename = "effectLst")]
    EffectLst(Box<EffectList>),
    #[serde(rename = "effectDag")]
    EffectDag(Box<EffectContainer>),
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub enum EGGeometry {
    #[serde(rename = "custGeom")]
    CustGeom(Box<CTCustomGeometry2D>),
    #[serde(rename = "prstGeom")]
    PrstGeom(Box<CTPresetGeometry2D>),
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub enum EGTextGeometry {
    #[serde(rename = "custGeom")]
    CustGeom(Box<CTCustomGeometry2D>),
    #[serde(rename = "prstTxWarp")]
    PrstTxWarp(Box<CTPresetTextShape>),
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub enum EGLineFillProperties {
    #[serde(rename = "noFill")]
    NoFill(Box<NoFill>),
    #[serde(rename = "solidFill")]
    SolidFill(Box<SolidColorFill>),
    #[serde(rename = "gradFill")]
    GradFill(Box<GradientFill>),
    #[serde(rename = "pattFill")]
    PattFill(Box<PatternFill>),
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub enum EGLineJoinProperties {
    #[serde(rename = "round")]
    Round(Box<CTLineJoinRound>),
    #[serde(rename = "bevel")]
    Bevel(Box<CTLineJoinBevel>),
    #[serde(rename = "miter")]
    Miter(Box<CTLineJoinMiterProperties>),
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub enum EGLineDashProperties {
    #[serde(rename = "prstDash")]
    PrstDash(Box<CTPresetLineDashProperties>),
    #[serde(rename = "custDash")]
    CustDash(Box<CTDashStopList>),
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub enum EGThemeableFillStyle {
    #[serde(rename = "fill")]
    Fill(Box<CTFillProperties>),
    #[serde(rename = "fillRef")]
    FillRef(Box<CTStyleMatrixReference>),
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub enum EGThemeableEffectStyle {
    #[serde(rename = "effect")]
    Effect(Box<CTEffectProperties>),
    #[serde(rename = "effectRef")]
    EffectRef(Box<CTStyleMatrixReference>),
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub enum EGThemeableFontStyles {
    #[serde(rename = "font")]
    Font(Box<CTFontCollection>),
    #[serde(rename = "fontRef")]
    FontRef(Box<CTFontReference>),
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub enum EGTextAutofit {
    #[serde(rename = "noAutofit")]
    NoAutofit(Box<CTTextNoAutofit>),
    #[serde(rename = "normAutofit")]
    NormAutofit(Box<CTTextNormalAutofit>),
    #[serde(rename = "spAutoFit")]
    SpAutoFit(Box<CTTextShapeAutofit>),
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub enum EGTextBulletColor {
    #[serde(rename = "buClrTx")]
    BuClrTx(Box<CTTextBulletColorFollowText>),
    #[serde(rename = "buClr")]
    BuClr(Box<CTColor>),
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub enum EGTextBulletSize {
    #[serde(rename = "buSzTx")]
    BuSzTx(Box<CTTextBulletSizeFollowText>),
    #[serde(rename = "buSzPct")]
    BuSzPct(Box<CTTextBulletSizePercent>),
    #[serde(rename = "buSzPts")]
    BuSzPts(Box<CTTextBulletSizePoint>),
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub enum EGTextBulletTypeface {
    #[serde(rename = "buFontTx")]
    BuFontTx(Box<CTTextBulletTypefaceFollowText>),
    #[serde(rename = "buFont")]
    BuFont(Box<TextFont>),
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub enum EGTextBullet {
    #[serde(rename = "buNone")]
    BuNone(Box<CTTextNoBullet>),
    #[serde(rename = "buAutoNum")]
    BuAutoNum(Box<CTTextAutonumberBullet>),
    #[serde(rename = "buChar")]
    BuChar(Box<CTTextCharBullet>),
    #[serde(rename = "buBlip")]
    BuBlip(Box<CTTextBlipBullet>),
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub enum EGTextUnderlineLine {
    #[serde(rename = "uLnTx")]
    ULnTx(Box<CTTextUnderlineLineFollowText>),
    #[serde(rename = "uLn")]
    ULn(Box<LineProperties>),
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub enum EGTextUnderlineFill {
    #[serde(rename = "uFillTx")]
    UFillTx(Box<CTTextUnderlineFillFollowText>),
    #[serde(rename = "uFill")]
    UFill(Box<CTTextUnderlineFillGroupWrapper>),
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub enum EGTextRun {
    #[serde(rename = "r")]
    R(Box<TextRun>),
    #[serde(rename = "br")]
    Br(Box<CTTextLineBreak>),
    #[serde(rename = "fld")]
    Fld(Box<CTTextField>),
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTAudioFile {
    #[serde(rename = "@contentType")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub content_type: Option<String>,
    #[serde(rename = "extLst")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub ext_lst: Option<Box<CTOfficeArtExtensionList>>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
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
pub struct CTVideoFile {
    #[serde(rename = "@contentType")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub content_type: Option<String>,
    #[serde(rename = "extLst")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub ext_lst: Option<Box<CTOfficeArtExtensionList>>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
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
pub struct CTQuickTimeFile {
    #[serde(rename = "extLst")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub ext_lst: Option<Box<CTOfficeArtExtensionList>>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::RawXmlNode>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTAudioCDTime {
    #[serde(rename = "@track")]
    pub track: u8,
    #[serde(rename = "@time")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub time: Option<u32>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTAudioCD {
    #[serde(rename = "st")]
    pub st: Box<CTAudioCDTime>,
    #[serde(rename = "end")]
    pub end: Box<CTAudioCDTime>,
    #[serde(rename = "extLst")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub ext_lst: Option<Box<CTOfficeArtExtensionList>>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::RawXmlNode>,
}

pub type AVideoFile = Box<CTVideoFile>;

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct ColorScheme {
    #[serde(rename = "@name")]
    pub name: String,
    #[serde(rename = "dk1")]
    pub dk1: Box<CTColor>,
    #[serde(rename = "lt1")]
    pub lt1: Box<CTColor>,
    #[serde(rename = "dk2")]
    pub dk2: Box<CTColor>,
    #[serde(rename = "lt2")]
    pub lt2: Box<CTColor>,
    #[serde(rename = "accent1")]
    pub accent1: Box<CTColor>,
    #[serde(rename = "accent2")]
    pub accent2: Box<CTColor>,
    #[serde(rename = "accent3")]
    pub accent3: Box<CTColor>,
    #[serde(rename = "accent4")]
    pub accent4: Box<CTColor>,
    #[serde(rename = "accent5")]
    pub accent5: Box<CTColor>,
    #[serde(rename = "accent6")]
    pub accent6: Box<CTColor>,
    #[serde(rename = "hlink")]
    pub hlink: Box<CTColor>,
    #[serde(rename = "folHlink")]
    pub fol_hlink: Box<CTColor>,
    #[serde(rename = "extLst")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub ext_lst: Option<Box<CTOfficeArtExtensionList>>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
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
pub struct CTCustomColor {
    #[serde(rename = "@name")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub name: Option<String>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTSupplementalFont {
    #[serde(rename = "@script")]
    pub script: String,
    #[serde(rename = "@typeface")]
    pub typeface: STTextTypeface,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTCustomColorList {
    #[serde(rename = "custClr")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub cust_clr: Vec<Box<CTCustomColor>>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::RawXmlNode>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTFontCollection {
    #[serde(rename = "latin")]
    pub latin: Box<TextFont>,
    #[serde(rename = "ea")]
    pub ea: Box<TextFont>,
    #[serde(rename = "cs")]
    pub cs: Box<TextFont>,
    #[serde(rename = "font")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub font: Vec<Box<CTSupplementalFont>>,
    #[serde(rename = "extLst")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub ext_lst: Option<Box<CTOfficeArtExtensionList>>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::RawXmlNode>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTEffectStyleItem {
    #[serde(rename = "scene3d")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub scene3d: Option<Box<CTScene3D>>,
    #[serde(rename = "sp3d")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub sp3d: Option<Box<CTShape3D>>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::RawXmlNode>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct FontScheme {
    #[serde(rename = "@name")]
    pub name: String,
    #[serde(rename = "majorFont")]
    pub major_font: Box<CTFontCollection>,
    #[serde(rename = "minorFont")]
    pub minor_font: Box<CTFontCollection>,
    #[serde(rename = "extLst")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub ext_lst: Option<Box<CTOfficeArtExtensionList>>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
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
pub struct CTFillStyleList;

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTLineStyleList {
    #[serde(rename = "ln")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub line: Vec<Box<LineProperties>>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::RawXmlNode>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTEffectStyleList {
    #[serde(rename = "effectStyle")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub effect_style: Vec<Box<CTEffectStyleItem>>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::RawXmlNode>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct CTBackgroundFillStyleList;

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTStyleMatrix {
    #[serde(rename = "@name")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub name: Option<String>,
    #[serde(rename = "fillStyleLst")]
    pub fill_style_lst: Box<CTFillStyleList>,
    #[serde(rename = "lnStyleLst")]
    pub ln_style_lst: Box<CTLineStyleList>,
    #[serde(rename = "effectStyleLst")]
    pub effect_style_lst: Box<CTEffectStyleList>,
    #[serde(rename = "bgFillStyleLst")]
    pub bg_fill_style_lst: Box<CTBackgroundFillStyleList>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
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
pub struct CTBaseStyles {
    #[serde(rename = "clrScheme")]
    pub clr_scheme: Box<ColorScheme>,
    #[serde(rename = "fontScheme")]
    pub font_scheme: Box<FontScheme>,
    #[serde(rename = "fmtScheme")]
    pub fmt_scheme: Box<CTStyleMatrix>,
    #[serde(rename = "extLst")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub ext_lst: Option<Box<CTOfficeArtExtensionList>>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::RawXmlNode>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTOfficeArtExtension {
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

pub type CTOfficeArtExtensionAny = String;

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTAngle {
    #[serde(rename = "@val")]
    pub value: STAngle,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTPositiveFixedAngle {
    #[serde(rename = "@val")]
    pub value: STPositiveFixedAngle,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTPercentage {
    #[serde(rename = "@val")]
    pub value: STPercentage,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTPositivePercentage {
    #[serde(rename = "@val")]
    pub value: STPositivePercentage,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTFixedPercentage {
    #[serde(rename = "@val")]
    pub value: STFixedPercentage,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTPositiveFixedPercentage {
    #[serde(rename = "@val")]
    pub value: STPositiveFixedPercentage,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTRatio {
    #[serde(rename = "@n")]
    pub n: i64,
    #[serde(rename = "@d")]
    pub d: i64,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct Point2D {
    #[serde(rename = "@x")]
    pub x: STCoordinate,
    #[serde(rename = "@y")]
    pub y: STCoordinate,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct PositiveSize2D {
    #[serde(rename = "@cx")]
    pub cx: STPositiveCoordinate,
    #[serde(rename = "@cy")]
    pub cy: STPositiveCoordinate,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct CTComplementTransform;

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct CTInverseTransform;

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct CTGrayscaleTransform;

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct CTGammaTransform;

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct CTInverseGammaTransform;

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTScRgbColor {
    #[serde(rename = "@r")]
    pub relationship_id: STPercentage,
    #[serde(rename = "@g")]
    pub g: STPercentage,
    #[serde(rename = "@b")]
    pub b: STPercentage,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct SrgbColor {
    #[serde(rename = "@val")]
    pub value: HexColorRgb,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct HslColor {
    #[serde(rename = "@hue")]
    pub hue: STPositiveFixedAngle,
    #[serde(rename = "@sat")]
    pub sat: STPercentage,
    #[serde(rename = "@lum")]
    pub lum: STPercentage,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct SystemColor {
    #[serde(rename = "@val")]
    pub value: STSystemColorVal,
    #[serde(rename = "@lastClr")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub last_clr: Option<HexColorRgb>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct SchemeColor {
    #[serde(rename = "@val")]
    pub value: STSchemeColorVal,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct PresetColor {
    #[serde(rename = "@val")]
    pub value: STPresetColorVal,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct EGOfficeArtExtensionList {
    #[serde(rename = "ext")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub extents: Vec<Box<CTOfficeArtExtension>>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::RawXmlNode>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct CTOfficeArtExtensionList;

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTScale2D {
    #[serde(rename = "sx")]
    pub sx: Box<CTRatio>,
    #[serde(rename = "sy")]
    pub sy: Box<CTRatio>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::RawXmlNode>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct Transform2D {
    #[serde(rename = "@rot")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub rot: Option<STAngle>,
    #[serde(rename = "@flipH")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub flip_h: Option<bool>,
    #[serde(rename = "@flipV")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub flip_v: Option<bool>,
    #[serde(rename = "off")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub offset: Option<Box<Point2D>>,
    #[serde(rename = "ext")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub extents: Option<Box<PositiveSize2D>>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
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
pub struct CTGroupTransform2D {
    #[serde(rename = "@rot")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub rot: Option<STAngle>,
    #[serde(rename = "@flipH")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub flip_h: Option<bool>,
    #[serde(rename = "@flipV")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub flip_v: Option<bool>,
    #[serde(rename = "off")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub offset: Option<Box<Point2D>>,
    #[serde(rename = "ext")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub extents: Option<Box<PositiveSize2D>>,
    #[serde(rename = "chOff")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub child_offset: Option<Box<Point2D>>,
    #[serde(rename = "chExt")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub child_extents: Option<Box<PositiveSize2D>>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
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
pub struct CTPoint3D {
    #[serde(rename = "@x")]
    pub x: STCoordinate,
    #[serde(rename = "@y")]
    pub y: STCoordinate,
    #[serde(rename = "@z")]
    pub z: STCoordinate,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTVector3D {
    #[serde(rename = "@dx")]
    pub dx: STCoordinate,
    #[serde(rename = "@dy")]
    pub dy: STCoordinate,
    #[serde(rename = "@dz")]
    pub dz: STCoordinate,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTSphereCoords {
    #[serde(rename = "@lat")]
    pub lat: STPositiveFixedAngle,
    #[serde(rename = "@lon")]
    pub lon: STPositiveFixedAngle,
    #[serde(rename = "@rev")]
    pub rev: STPositiveFixedAngle,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTRelativeRect {
    #[serde(rename = "@l")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub l: Option<STPercentage>,
    #[serde(rename = "@t")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub t: Option<STPercentage>,
    #[serde(rename = "@r")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub relationship_id: Option<STPercentage>,
    #[serde(rename = "@b")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub b: Option<STPercentage>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct CTColor;

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct CTColorMRU;

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct AAGBlob;

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTEmbeddedWAVAudioFile {
    #[serde(rename = "@name")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub name: Option<String>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTHyperlink {
    #[serde(rename = "@invalidUrl")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub invalid_url: Option<String>,
    #[serde(rename = "@action")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub action: Option<String>,
    #[serde(rename = "@tgtFrame")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub tgt_frame: Option<String>,
    #[serde(rename = "@tooltip")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub tooltip: Option<String>,
    #[serde(rename = "@history")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub history: Option<bool>,
    #[serde(rename = "@highlightClick")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub highlight_click: Option<bool>,
    #[serde(rename = "@endSnd")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub end_snd: Option<bool>,
    #[serde(rename = "snd")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub snd: Option<Box<CTEmbeddedWAVAudioFile>>,
    #[serde(rename = "extLst")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub ext_lst: Option<Box<CTOfficeArtExtensionList>>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
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
pub struct AAGLocking {
    #[serde(rename = "@noGrp")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub no_grp: Option<bool>,
    #[serde(rename = "@noSelect")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub no_select: Option<bool>,
    #[serde(rename = "@noRot")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub no_rot: Option<bool>,
    #[serde(rename = "@noChangeAspect")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub no_change_aspect: Option<bool>,
    #[serde(rename = "@noMove")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub no_move: Option<bool>,
    #[serde(rename = "@noResize")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub no_resize: Option<bool>,
    #[serde(rename = "@noEditPoints")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub no_edit_points: Option<bool>,
    #[serde(rename = "@noAdjustHandles")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub no_adjust_handles: Option<bool>,
    #[serde(rename = "@noChangeArrowheads")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub no_change_arrowheads: Option<bool>,
    #[serde(rename = "@noChangeShapeType")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub no_change_shape_type: Option<bool>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTConnectorLocking {
    #[serde(rename = "extLst")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub ext_lst: Option<Box<CTOfficeArtExtensionList>>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::RawXmlNode>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTShapeLocking {
    #[serde(rename = "@noTextEdit")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub no_text_edit: Option<bool>,
    #[serde(rename = "extLst")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub ext_lst: Option<Box<CTOfficeArtExtensionList>>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
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
pub struct CTPictureLocking {
    #[serde(rename = "@noCrop")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub no_crop: Option<bool>,
    #[serde(rename = "extLst")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub ext_lst: Option<Box<CTOfficeArtExtensionList>>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
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
pub struct CTGroupLocking {
    #[serde(rename = "@noGrp")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub no_grp: Option<bool>,
    #[serde(rename = "@noUngrp")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub no_ungrp: Option<bool>,
    #[serde(rename = "@noSelect")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub no_select: Option<bool>,
    #[serde(rename = "@noRot")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub no_rot: Option<bool>,
    #[serde(rename = "@noChangeAspect")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub no_change_aspect: Option<bool>,
    #[serde(rename = "@noMove")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub no_move: Option<bool>,
    #[serde(rename = "@noResize")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub no_resize: Option<bool>,
    #[serde(rename = "extLst")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub ext_lst: Option<Box<CTOfficeArtExtensionList>>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
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
pub struct CTGraphicalObjectFrameLocking {
    #[serde(rename = "@noGrp")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub no_grp: Option<bool>,
    #[serde(rename = "@noDrilldown")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub no_drilldown: Option<bool>,
    #[serde(rename = "@noSelect")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub no_select: Option<bool>,
    #[serde(rename = "@noChangeAspect")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub no_change_aspect: Option<bool>,
    #[serde(rename = "@noMove")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub no_move: Option<bool>,
    #[serde(rename = "@noResize")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub no_resize: Option<bool>,
    #[serde(rename = "extLst")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub ext_lst: Option<Box<CTOfficeArtExtensionList>>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
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
pub struct CTContentPartLocking {
    #[serde(rename = "extLst")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub ext_lst: Option<Box<CTOfficeArtExtensionList>>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::RawXmlNode>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTNonVisualDrawingProps {
    #[serde(rename = "@id")]
    pub id: STDrawingElementId,
    #[serde(rename = "@name")]
    pub name: String,
    #[serde(rename = "@descr")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub descr: Option<String>,
    #[serde(rename = "@hidden")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub hidden: Option<bool>,
    #[serde(rename = "@title")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub title: Option<String>,
    #[serde(rename = "hlinkClick")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub hlink_click: Option<Box<CTHyperlink>>,
    #[serde(rename = "hlinkHover")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub hlink_hover: Option<Box<CTHyperlink>>,
    #[serde(rename = "extLst")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub ext_lst: Option<Box<CTOfficeArtExtensionList>>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
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
pub struct CTNonVisualDrawingShapeProps {
    #[serde(rename = "@txBox")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub tx_box: Option<bool>,
    #[serde(rename = "spLocks")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub sp_locks: Option<Box<CTShapeLocking>>,
    #[serde(rename = "extLst")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub ext_lst: Option<Box<CTOfficeArtExtensionList>>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
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
pub struct CTNonVisualConnectorProperties {
    #[serde(rename = "cxnSpLocks")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub cxn_sp_locks: Option<Box<CTConnectorLocking>>,
    #[serde(rename = "stCxn")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub st_cxn: Option<Box<CTConnection>>,
    #[serde(rename = "endCxn")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub end_cxn: Option<Box<CTConnection>>,
    #[serde(rename = "extLst")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub ext_lst: Option<Box<CTOfficeArtExtensionList>>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::RawXmlNode>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTNonVisualPictureProperties {
    #[serde(rename = "@preferRelativeResize")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub prefer_relative_resize: Option<bool>,
    #[serde(rename = "picLocks")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub pic_locks: Option<Box<CTPictureLocking>>,
    #[serde(rename = "extLst")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub ext_lst: Option<Box<CTOfficeArtExtensionList>>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
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
pub struct CTNonVisualGroupDrawingShapeProps {
    #[serde(rename = "grpSpLocks")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub grp_sp_locks: Option<Box<CTGroupLocking>>,
    #[serde(rename = "extLst")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub ext_lst: Option<Box<CTOfficeArtExtensionList>>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::RawXmlNode>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTNonVisualGraphicFrameProperties {
    #[serde(rename = "graphicFrameLocks")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub graphic_frame_locks: Option<Box<CTGraphicalObjectFrameLocking>>,
    #[serde(rename = "extLst")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub ext_lst: Option<Box<CTOfficeArtExtensionList>>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::RawXmlNode>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTNonVisualContentPartProperties {
    #[serde(rename = "@isComment")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub is_comment: Option<bool>,
    #[serde(rename = "cpLocks")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub cp_locks: Option<Box<CTContentPartLocking>>,
    #[serde(rename = "extLst")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub ext_lst: Option<Box<CTOfficeArtExtensionList>>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
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
pub struct CTGraphicalObjectData {
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

pub type CTGraphicalObjectDataAny = String;

pub type CTGraphicalObject = Box<CTGraphicalObjectData>;

pub type AGraphic = Box<CTGraphicalObject>;

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTAnimationDgmElement {
    #[serde(rename = "@id")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub id: Option<Guid>,
    #[serde(rename = "@bldStep")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub bld_step: Option<STDgmBuildStep>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTAnimationChartElement {
    #[serde(rename = "@seriesIdx")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub series_idx: Option<i32>,
    #[serde(rename = "@categoryIdx")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub category_idx: Option<i32>,
    #[serde(rename = "@bldStep")]
    pub bld_step: STChartBuildStep,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct CTAnimationElementChoice;

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTAnimationDgmBuildProperties {
    #[serde(rename = "@bld")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub bld: Option<STAnimationDgmBuildType>,
    #[serde(rename = "@rev")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub rev: Option<bool>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTAnimationChartBuildProperties {
    #[serde(rename = "@bld")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub bld: Option<STAnimationChartBuildType>,
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
pub struct CTAnimationGraphicalObjectBuildProperties;

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct CTBackgroundFormatting;

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTWholeE2oFormatting {
    #[serde(rename = "ln")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub line: Option<Box<LineProperties>>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::RawXmlNode>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct CTGvmlUseShapeRectangle;

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTGvmlTextShape {
    #[serde(rename = "txBody")]
    pub tx_body: Box<TextBody>,
    #[serde(rename = "extLst")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub ext_lst: Option<Box<CTOfficeArtExtensionList>>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::RawXmlNode>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTGvmlShapeNonVisual {
    #[serde(rename = "cNvPr")]
    pub common_non_visual_properties: Box<CTNonVisualDrawingProps>,
    #[serde(rename = "cNvSpPr")]
    pub common_non_visual_shape_properties: Box<CTNonVisualDrawingShapeProps>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::RawXmlNode>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTGvmlShape {
    #[serde(rename = "nvSpPr")]
    pub nv_sp_pr: Box<CTGvmlShapeNonVisual>,
    #[serde(rename = "spPr")]
    pub sp_pr: Box<CTShapeProperties>,
    #[serde(rename = "txSp")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub tx_sp: Option<Box<CTGvmlTextShape>>,
    #[serde(rename = "style")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub style: Option<Box<ShapeStyle>>,
    #[serde(rename = "extLst")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub ext_lst: Option<Box<CTOfficeArtExtensionList>>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::RawXmlNode>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTGvmlConnectorNonVisual {
    #[serde(rename = "cNvPr")]
    pub common_non_visual_properties: Box<CTNonVisualDrawingProps>,
    #[serde(rename = "cNvCxnSpPr")]
    pub c_nv_cxn_sp_pr: Box<CTNonVisualConnectorProperties>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::RawXmlNode>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTGvmlConnector {
    #[serde(rename = "nvCxnSpPr")]
    pub nv_cxn_sp_pr: Box<CTGvmlConnectorNonVisual>,
    #[serde(rename = "spPr")]
    pub sp_pr: Box<CTShapeProperties>,
    #[serde(rename = "style")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub style: Option<Box<ShapeStyle>>,
    #[serde(rename = "extLst")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub ext_lst: Option<Box<CTOfficeArtExtensionList>>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::RawXmlNode>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTGvmlPictureNonVisual {
    #[serde(rename = "cNvPr")]
    pub common_non_visual_properties: Box<CTNonVisualDrawingProps>,
    #[serde(rename = "cNvPicPr")]
    pub common_non_visual_picture_properties: Box<CTNonVisualPictureProperties>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::RawXmlNode>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTGvmlPicture {
    #[serde(rename = "nvPicPr")]
    pub nv_pic_pr: Box<CTGvmlPictureNonVisual>,
    #[serde(rename = "blipFill")]
    pub blip_fill: Box<BlipFillProperties>,
    #[serde(rename = "spPr")]
    pub sp_pr: Box<CTShapeProperties>,
    #[serde(rename = "style")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub style: Option<Box<ShapeStyle>>,
    #[serde(rename = "extLst")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub ext_lst: Option<Box<CTOfficeArtExtensionList>>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::RawXmlNode>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTGvmlGraphicFrameNonVisual {
    #[serde(rename = "cNvPr")]
    pub common_non_visual_properties: Box<CTNonVisualDrawingProps>,
    #[serde(rename = "cNvGraphicFramePr")]
    pub c_nv_graphic_frame_pr: Box<CTNonVisualGraphicFrameProperties>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::RawXmlNode>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTGvmlGraphicalObjectFrame {
    #[serde(rename = "nvGraphicFramePr")]
    pub nv_graphic_frame_pr: Box<CTGvmlGraphicFrameNonVisual>,
    #[serde(rename = "xfrm")]
    pub transform: Box<Transform2D>,
    #[serde(rename = "extLst")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub ext_lst: Option<Box<CTOfficeArtExtensionList>>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::RawXmlNode>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTGvmlGroupShapeNonVisual {
    #[serde(rename = "cNvPr")]
    pub common_non_visual_properties: Box<CTNonVisualDrawingProps>,
    #[serde(rename = "cNvGrpSpPr")]
    pub c_nv_grp_sp_pr: Box<CTNonVisualGroupDrawingShapeProps>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::RawXmlNode>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTGvmlGroupShape {
    #[serde(rename = "nvGrpSpPr")]
    pub nv_grp_sp_pr: Box<CTGvmlGroupShapeNonVisual>,
    #[serde(rename = "grpSpPr")]
    pub grp_sp_pr: Box<CTGroupShapeProperties>,
    #[serde(rename = "extLst")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub ext_lst: Option<Box<CTOfficeArtExtensionList>>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::RawXmlNode>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTCamera {
    #[serde(rename = "@prst")]
    pub preset: STPresetCameraType,
    #[serde(rename = "@fov")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub fov: Option<STFOVAngle>,
    #[serde(rename = "@zoom")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub zoom: Option<STPositivePercentage>,
    #[serde(rename = "rot")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub rot: Option<Box<CTSphereCoords>>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
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
pub struct CTLightRig {
    #[serde(rename = "@rig")]
    pub rig: STLightRigType,
    #[serde(rename = "@dir")]
    pub dir: STLightRigDirection,
    #[serde(rename = "rot")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub rot: Option<Box<CTSphereCoords>>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
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
pub struct CTScene3D {
    #[serde(rename = "camera")]
    pub camera: Box<CTCamera>,
    #[serde(rename = "lightRig")]
    pub light_rig: Box<CTLightRig>,
    #[serde(rename = "backdrop")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub backdrop: Option<Box<CTBackdrop>>,
    #[serde(rename = "extLst")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub ext_lst: Option<Box<CTOfficeArtExtensionList>>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::RawXmlNode>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTBackdrop {
    #[serde(rename = "anchor")]
    pub anchor: Box<CTPoint3D>,
    #[serde(rename = "norm")]
    pub norm: Box<CTVector3D>,
    #[serde(rename = "up")]
    pub up: Box<CTVector3D>,
    #[serde(rename = "extLst")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub ext_lst: Option<Box<CTOfficeArtExtensionList>>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::RawXmlNode>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTBevel {
    #[serde(rename = "@w")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub width: Option<STPositiveCoordinate>,
    #[serde(rename = "@h")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub height: Option<STPositiveCoordinate>,
    #[serde(rename = "@prst")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub preset: Option<STBevelPresetType>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTShape3D {
    #[serde(rename = "@z")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub z: Option<STCoordinate>,
    #[serde(rename = "@extrusionH")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub extrusion_h: Option<STPositiveCoordinate>,
    #[serde(rename = "@contourW")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub contour_w: Option<STPositiveCoordinate>,
    #[serde(rename = "@prstMaterial")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub prst_material: Option<STPresetMaterialType>,
    #[serde(rename = "bevelT")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub bevel_t: Option<Box<CTBevel>>,
    #[serde(rename = "bevelB")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub bevel_b: Option<Box<CTBevel>>,
    #[serde(rename = "extrusionClr")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub extrusion_clr: Option<Box<CTColor>>,
    #[serde(rename = "contourClr")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub contour_clr: Option<Box<CTColor>>,
    #[serde(rename = "extLst")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub ext_lst: Option<Box<CTOfficeArtExtensionList>>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
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
pub struct CTFlatText {
    #[serde(rename = "@z")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub z: Option<STCoordinate>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTAlphaBiLevelEffect {
    #[serde(rename = "@thresh")]
    pub thresh: STPositiveFixedPercentage,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct CTAlphaCeilingEffect;

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct CTAlphaFloorEffect;

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct CTAlphaInverseEffect;

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTAlphaModulateFixedEffect {
    #[serde(rename = "@amt")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub amt: Option<STPositivePercentage>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTAlphaOutsetEffect {
    #[serde(rename = "@rad")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub rad: Option<STCoordinate>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTAlphaReplaceEffect {
    #[serde(rename = "@a")]
    pub anchor: STPositiveFixedPercentage,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTBiLevelEffect {
    #[serde(rename = "@thresh")]
    pub thresh: STPositiveFixedPercentage,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTBlurEffect {
    #[serde(rename = "@rad")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub rad: Option<STPositiveCoordinate>,
    #[serde(rename = "@grow")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub grow: Option<bool>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTColorChangeEffect {
    #[serde(rename = "@useA")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub use_a: Option<bool>,
    #[serde(rename = "clrFrom")]
    pub clr_from: Box<CTColor>,
    #[serde(rename = "clrTo")]
    pub clr_to: Box<CTColor>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
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
pub struct CTColorReplaceEffect;

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct CTDuotoneEffect;

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTGlowEffect {
    #[serde(rename = "@rad")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub rad: Option<STPositiveCoordinate>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct CTGrayscaleEffect;

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTHSLEffect {
    #[serde(rename = "@hue")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub hue: Option<STPositiveFixedAngle>,
    #[serde(rename = "@sat")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub sat: Option<STFixedPercentage>,
    #[serde(rename = "@lum")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub lum: Option<STFixedPercentage>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTInnerShadowEffect {
    #[serde(rename = "@blurRad")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub blur_rad: Option<STPositiveCoordinate>,
    #[serde(rename = "@dist")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub dist: Option<STPositiveCoordinate>,
    #[serde(rename = "@dir")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub dir: Option<STPositiveFixedAngle>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTLuminanceEffect {
    #[serde(rename = "@bright")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub bright: Option<STFixedPercentage>,
    #[serde(rename = "@contrast")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub contrast: Option<STFixedPercentage>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTOuterShadowEffect {
    #[serde(rename = "@blurRad")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub blur_rad: Option<STPositiveCoordinate>,
    #[serde(rename = "@dist")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub dist: Option<STPositiveCoordinate>,
    #[serde(rename = "@dir")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub dir: Option<STPositiveFixedAngle>,
    #[serde(rename = "@sx")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub sx: Option<STPercentage>,
    #[serde(rename = "@sy")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub sy: Option<STPercentage>,
    #[serde(rename = "@kx")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub kx: Option<STFixedAngle>,
    #[serde(rename = "@ky")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub ky: Option<STFixedAngle>,
    #[serde(rename = "@algn")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub algn: Option<STRectAlignment>,
    #[serde(rename = "@rotWithShape")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub rot_with_shape: Option<bool>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTPresetShadowEffect {
    #[serde(rename = "@prst")]
    pub preset: STPresetShadowVal,
    #[serde(rename = "@dist")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub dist: Option<STPositiveCoordinate>,
    #[serde(rename = "@dir")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub dir: Option<STPositiveFixedAngle>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTReflectionEffect {
    #[serde(rename = "@blurRad")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub blur_rad: Option<STPositiveCoordinate>,
    #[serde(rename = "@stA")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub st_a: Option<STPositiveFixedPercentage>,
    #[serde(rename = "@stPos")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub st_pos: Option<STPositiveFixedPercentage>,
    #[serde(rename = "@endA")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub end_a: Option<STPositiveFixedPercentage>,
    #[serde(rename = "@endPos")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub end_pos: Option<STPositiveFixedPercentage>,
    #[serde(rename = "@dist")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub dist: Option<STPositiveCoordinate>,
    #[serde(rename = "@dir")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub dir: Option<STPositiveFixedAngle>,
    #[serde(rename = "@fadeDir")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub fade_dir: Option<STPositiveFixedAngle>,
    #[serde(rename = "@sx")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub sx: Option<STPercentage>,
    #[serde(rename = "@sy")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub sy: Option<STPercentage>,
    #[serde(rename = "@kx")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub kx: Option<STFixedAngle>,
    #[serde(rename = "@ky")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub ky: Option<STFixedAngle>,
    #[serde(rename = "@algn")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub algn: Option<STRectAlignment>,
    #[serde(rename = "@rotWithShape")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub rot_with_shape: Option<bool>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTRelativeOffsetEffect {
    #[serde(rename = "@tx")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub tx: Option<STPercentage>,
    #[serde(rename = "@ty")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub ty: Option<STPercentage>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTSoftEdgesEffect {
    #[serde(rename = "@rad")]
    pub rad: STPositiveCoordinate,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTTintEffect {
    #[serde(rename = "@hue")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub hue: Option<STPositiveFixedAngle>,
    #[serde(rename = "@amt")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub amt: Option<STFixedPercentage>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTTransformEffect {
    #[serde(rename = "@sx")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub sx: Option<STPercentage>,
    #[serde(rename = "@sy")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub sy: Option<STPercentage>,
    #[serde(rename = "@kx")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub kx: Option<STFixedAngle>,
    #[serde(rename = "@ky")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub ky: Option<STFixedAngle>,
    #[serde(rename = "@tx")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub tx: Option<STCoordinate>,
    #[serde(rename = "@ty")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub ty: Option<STCoordinate>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct NoFill;

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct SolidColorFill;

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTLinearShadeProperties {
    #[serde(rename = "@ang")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub ang: Option<STPositiveFixedAngle>,
    #[serde(rename = "@scaled")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub scaled: Option<bool>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTPathShadeProperties {
    #[serde(rename = "@path")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub path: Option<STPathShadeType>,
    #[serde(rename = "fillToRect")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub fill_to_rect: Option<Box<CTRelativeRect>>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
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
pub struct CTGradientStop {
    #[serde(rename = "@pos")]
    pub pos: STPositiveFixedPercentage,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTGradientStopList {
    #[serde(rename = "gs")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub gs: Vec<Box<CTGradientStop>>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::RawXmlNode>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct GradientFill {
    #[serde(rename = "@flip")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub flip: Option<STTileFlipMode>,
    #[serde(rename = "@rotWithShape")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub rot_with_shape: Option<bool>,
    #[serde(rename = "gsLst")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub gs_lst: Option<Box<CTGradientStopList>>,
    #[serde(rename = "tileRect")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub tile_rect: Option<Box<CTRelativeRect>>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
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
pub struct CTTileInfoProperties {
    #[serde(rename = "@tx")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub tx: Option<STCoordinate>,
    #[serde(rename = "@ty")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub ty: Option<STCoordinate>,
    #[serde(rename = "@sx")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub sx: Option<STPercentage>,
    #[serde(rename = "@sy")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub sy: Option<STPercentage>,
    #[serde(rename = "@flip")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub flip: Option<STTileFlipMode>,
    #[serde(rename = "@algn")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub algn: Option<STRectAlignment>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTStretchInfoProperties {
    #[serde(rename = "fillRect")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub fill_rect: Option<Box<CTRelativeRect>>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::RawXmlNode>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct Blip {
    #[serde(rename = "@cstate")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub cstate: Option<STBlipCompression>,
    #[serde(rename = "extLst")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub ext_lst: Option<Box<CTOfficeArtExtensionList>>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
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
pub struct BlipFillProperties {
    #[serde(rename = "@dpi")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub dpi: Option<u32>,
    #[serde(rename = "@rotWithShape")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub rot_with_shape: Option<bool>,
    #[serde(rename = "blip")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub blip: Option<Box<Blip>>,
    #[serde(rename = "srcRect")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub src_rect: Option<Box<CTRelativeRect>>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
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
pub struct PatternFill {
    #[serde(rename = "@prst")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub preset: Option<STPresetPatternVal>,
    #[serde(rename = "fgClr")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub fg_clr: Option<Box<CTColor>>,
    #[serde(rename = "bgClr")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub bg_clr: Option<Box<CTColor>>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
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
pub struct CTGroupFillProperties;

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct CTFillProperties;

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct CTFillEffect;

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTFillOverlayEffect {
    #[serde(rename = "@blend")]
    pub blend: STBlendMode,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTEffectReference {
    #[serde(rename = "@ref")]
    pub r#ref: String,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct EffectContainer {
    #[serde(rename = "@type")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub r#type: Option<STEffectContainerType>,
    #[serde(rename = "@name")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub name: Option<String>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

pub type CTAlphaModulateEffect = Box<EffectContainer>;

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTBlendEffect {
    #[serde(rename = "@blend")]
    pub blend: STBlendMode,
    #[serde(rename = "cont")]
    pub cont: Box<EffectContainer>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
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
pub struct EffectList {
    #[serde(rename = "blur")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub blur: Option<Box<CTBlurEffect>>,
    #[serde(rename = "fillOverlay")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub fill_overlay: Option<Box<CTFillOverlayEffect>>,
    #[serde(rename = "glow")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub glow: Option<Box<CTGlowEffect>>,
    #[serde(rename = "innerShdw")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub inner_shdw: Option<Box<CTInnerShadowEffect>>,
    #[serde(rename = "outerShdw")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub outer_shdw: Option<Box<CTOuterShadowEffect>>,
    #[serde(rename = "prstShdw")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub prst_shdw: Option<Box<CTPresetShadowEffect>>,
    #[serde(rename = "reflection")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub reflection: Option<Box<CTReflectionEffect>>,
    #[serde(rename = "softEdge")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub soft_edge: Option<Box<CTSoftEdgesEffect>>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::RawXmlNode>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct CTEffectProperties;

pub type ABlip = Box<Blip>;

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTGeomGuide {
    #[serde(rename = "@name")]
    pub name: STGeomGuideName,
    #[serde(rename = "@fmla")]
    pub fmla: STGeomGuideFormula,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTGeomGuideList {
    #[serde(rename = "gd")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub gd: Vec<Box<CTGeomGuide>>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::RawXmlNode>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTAdjPoint2D {
    #[serde(rename = "@x")]
    pub x: STAdjCoordinate,
    #[serde(rename = "@y")]
    pub y: STAdjCoordinate,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTGeomRect {
    #[serde(rename = "@l")]
    pub l: STAdjCoordinate,
    #[serde(rename = "@t")]
    pub t: STAdjCoordinate,
    #[serde(rename = "@r")]
    pub relationship_id: STAdjCoordinate,
    #[serde(rename = "@b")]
    pub b: STAdjCoordinate,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTXYAdjustHandle {
    #[serde(rename = "@gdRefX")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub gd_ref_x: Option<STGeomGuideName>,
    #[serde(rename = "@minX")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub min_x: Option<STAdjCoordinate>,
    #[serde(rename = "@maxX")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub max_x: Option<STAdjCoordinate>,
    #[serde(rename = "@gdRefY")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub gd_ref_y: Option<STGeomGuideName>,
    #[serde(rename = "@minY")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub min_y: Option<STAdjCoordinate>,
    #[serde(rename = "@maxY")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub max_y: Option<STAdjCoordinate>,
    #[serde(rename = "pos")]
    pub pos: Box<CTAdjPoint2D>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
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
pub struct CTPolarAdjustHandle {
    #[serde(rename = "@gdRefR")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub gd_ref_r: Option<STGeomGuideName>,
    #[serde(rename = "@minR")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub min_r: Option<STAdjCoordinate>,
    #[serde(rename = "@maxR")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub max_r: Option<STAdjCoordinate>,
    #[serde(rename = "@gdRefAng")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub gd_ref_ang: Option<STGeomGuideName>,
    #[serde(rename = "@minAng")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub min_ang: Option<STAdjAngle>,
    #[serde(rename = "@maxAng")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub max_ang: Option<STAdjAngle>,
    #[serde(rename = "pos")]
    pub pos: Box<CTAdjPoint2D>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
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
pub struct CTConnectionSite {
    #[serde(rename = "@ang")]
    pub ang: STAdjAngle,
    #[serde(rename = "pos")]
    pub pos: Box<CTAdjPoint2D>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
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
pub struct CTAdjustHandleList;

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTConnectionSiteList {
    #[serde(rename = "cxn")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub cxn: Vec<Box<CTConnectionSite>>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::RawXmlNode>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTConnection {
    #[serde(rename = "@id")]
    pub id: STDrawingElementId,
    #[serde(rename = "@idx")]
    pub idx: u32,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

pub type CTPath2DMoveTo = Box<CTAdjPoint2D>;

pub type CTPath2DLineTo = Box<CTAdjPoint2D>;

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTPath2DArcTo {
    #[serde(rename = "@wR")]
    pub w_r: STAdjCoordinate,
    #[serde(rename = "@hR")]
    pub h_r: STAdjCoordinate,
    #[serde(rename = "@stAng")]
    pub st_ang: STAdjAngle,
    #[serde(rename = "@swAng")]
    pub sw_ang: STAdjAngle,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTPath2DQuadBezierTo {
    #[serde(rename = "pt")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub pt: Vec<Box<CTAdjPoint2D>>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::RawXmlNode>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTPath2DCubicBezierTo {
    #[serde(rename = "pt")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub pt: Vec<Box<CTAdjPoint2D>>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::RawXmlNode>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct CTPath2DClose;

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTPath2D {
    #[serde(rename = "@w")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub width: Option<STPositiveCoordinate>,
    #[serde(rename = "@h")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub height: Option<STPositiveCoordinate>,
    #[serde(rename = "@fill")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub fill: Option<STPathFillMode>,
    #[serde(rename = "@stroke")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub stroke: Option<bool>,
    #[serde(rename = "@extrusionOk")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub extrusion_ok: Option<bool>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTPath2DList {
    #[serde(rename = "path")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub path: Vec<Box<CTPath2D>>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::RawXmlNode>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTPresetGeometry2D {
    #[serde(rename = "@prst")]
    pub preset: STShapeType,
    #[serde(rename = "avLst")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub av_lst: Option<Box<CTGeomGuideList>>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
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
pub struct CTPresetTextShape {
    #[serde(rename = "@prst")]
    pub preset: STTextShapeType,
    #[serde(rename = "avLst")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub av_lst: Option<Box<CTGeomGuideList>>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
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
pub struct CTCustomGeometry2D {
    #[serde(rename = "avLst")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub av_lst: Option<Box<CTGeomGuideList>>,
    #[serde(rename = "gdLst")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub gd_lst: Option<Box<CTGeomGuideList>>,
    #[serde(rename = "ahLst")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub ah_lst: Option<Box<CTAdjustHandleList>>,
    #[serde(rename = "cxnLst")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub cxn_lst: Option<Box<CTConnectionSiteList>>,
    #[serde(rename = "rect")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub rect: Option<Box<CTGeomRect>>,
    #[serde(rename = "pathLst")]
    pub path_lst: Box<CTPath2DList>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::RawXmlNode>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTLineEndProperties {
    #[serde(rename = "@type")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub r#type: Option<STLineEndType>,
    #[serde(rename = "@w")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub width: Option<STLineEndWidth>,
    #[serde(rename = "@len")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub len: Option<STLineEndLength>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct CTLineJoinBevel;

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct CTLineJoinRound;

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTLineJoinMiterProperties {
    #[serde(rename = "@lim")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub lim: Option<STPositivePercentage>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTPresetLineDashProperties {
    #[serde(rename = "@val")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub value: Option<STPresetLineDashVal>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTDashStop {
    #[serde(rename = "@d")]
    pub d: STPositivePercentage,
    #[serde(rename = "@sp")]
    pub sp: STPositivePercentage,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTDashStopList {
    #[serde(rename = "ds")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub ds: Vec<Box<CTDashStop>>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::RawXmlNode>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct LineProperties {
    #[serde(rename = "@w")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub width: Option<STLineWidth>,
    #[serde(rename = "@cap")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub cap: Option<STLineCap>,
    #[serde(rename = "@cmpd")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub cmpd: Option<STCompoundLine>,
    #[serde(rename = "@algn")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub algn: Option<STPenAlignment>,
    #[serde(rename = "headEnd")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub head_end: Option<Box<CTLineEndProperties>>,
    #[serde(rename = "tailEnd")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub tail_end: Option<Box<CTLineEndProperties>>,
    #[serde(rename = "extLst")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub ext_lst: Option<Box<CTOfficeArtExtensionList>>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
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
pub struct CTShapeProperties {
    #[serde(rename = "@bwMode")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub bw_mode: Option<STBlackWhiteMode>,
    #[serde(rename = "xfrm")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub transform: Option<Box<Transform2D>>,
    #[serde(rename = "ln")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub line: Option<Box<LineProperties>>,
    #[serde(rename = "scene3d")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub scene3d: Option<Box<CTScene3D>>,
    #[serde(rename = "sp3d")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub sp3d: Option<Box<CTShape3D>>,
    #[serde(rename = "extLst")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub ext_lst: Option<Box<CTOfficeArtExtensionList>>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
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
pub struct CTGroupShapeProperties {
    #[serde(rename = "@bwMode")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub bw_mode: Option<STBlackWhiteMode>,
    #[serde(rename = "xfrm")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub transform: Option<Box<CTGroupTransform2D>>,
    #[serde(rename = "scene3d")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub scene3d: Option<Box<CTScene3D>>,
    #[serde(rename = "extLst")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub ext_lst: Option<Box<CTOfficeArtExtensionList>>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
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
pub struct CTStyleMatrixReference {
    #[serde(rename = "@idx")]
    pub idx: STStyleMatrixColumnIndex,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTFontReference {
    #[serde(rename = "@idx")]
    pub idx: STFontCollectionIndex,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct ShapeStyle {
    #[serde(rename = "lnRef")]
    pub ln_ref: Box<CTStyleMatrixReference>,
    #[serde(rename = "fillRef")]
    pub fill_ref: Box<CTStyleMatrixReference>,
    #[serde(rename = "effectRef")]
    pub effect_ref: Box<CTStyleMatrixReference>,
    #[serde(rename = "fontRef")]
    pub font_ref: Box<CTFontReference>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::RawXmlNode>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTDefaultShapeDefinition {
    #[serde(rename = "spPr")]
    pub sp_pr: Box<CTShapeProperties>,
    #[serde(rename = "bodyPr")]
    pub body_pr: Box<CTTextBodyProperties>,
    #[serde(rename = "lstStyle")]
    pub lst_style: Box<CTTextListStyle>,
    #[serde(rename = "style")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub style: Option<Box<ShapeStyle>>,
    #[serde(rename = "extLst")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub ext_lst: Option<Box<CTOfficeArtExtensionList>>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::RawXmlNode>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTObjectStyleDefaults {
    #[serde(rename = "spDef")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub sp_def: Option<Box<CTDefaultShapeDefinition>>,
    #[serde(rename = "lnDef")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub ln_def: Option<Box<CTDefaultShapeDefinition>>,
    #[serde(rename = "txDef")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub tx_def: Option<Box<CTDefaultShapeDefinition>>,
    #[serde(rename = "extLst")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub ext_lst: Option<Box<CTOfficeArtExtensionList>>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::RawXmlNode>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct CTEmptyElement;

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTColorMapping {
    #[serde(rename = "@bg1")]
    pub bg1: STColorSchemeIndex,
    #[serde(rename = "@tx1")]
    pub tx1: STColorSchemeIndex,
    #[serde(rename = "@bg2")]
    pub bg2: STColorSchemeIndex,
    #[serde(rename = "@tx2")]
    pub tx2: STColorSchemeIndex,
    #[serde(rename = "@accent1")]
    pub accent1: STColorSchemeIndex,
    #[serde(rename = "@accent2")]
    pub accent2: STColorSchemeIndex,
    #[serde(rename = "@accent3")]
    pub accent3: STColorSchemeIndex,
    #[serde(rename = "@accent4")]
    pub accent4: STColorSchemeIndex,
    #[serde(rename = "@accent5")]
    pub accent5: STColorSchemeIndex,
    #[serde(rename = "@accent6")]
    pub accent6: STColorSchemeIndex,
    #[serde(rename = "@hlink")]
    pub hlink: STColorSchemeIndex,
    #[serde(rename = "@folHlink")]
    pub fol_hlink: STColorSchemeIndex,
    #[serde(rename = "extLst")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub ext_lst: Option<Box<CTOfficeArtExtensionList>>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
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
pub struct CTColorMappingOverride;

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTColorSchemeAndMapping {
    #[serde(rename = "clrScheme")]
    pub clr_scheme: Box<ColorScheme>,
    #[serde(rename = "clrMap")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub clr_map: Option<Box<CTColorMapping>>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::RawXmlNode>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTColorSchemeList {
    #[serde(rename = "extraClrScheme")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub extra_clr_scheme: Vec<Box<CTColorSchemeAndMapping>>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::RawXmlNode>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTOfficeStyleSheet {
    #[serde(rename = "@name")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub name: Option<String>,
    #[serde(rename = "themeElements")]
    pub theme_elements: Box<CTBaseStyles>,
    #[serde(rename = "objectDefaults")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub object_defaults: Option<Box<CTObjectStyleDefaults>>,
    #[serde(rename = "extraClrSchemeLst")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub extra_clr_scheme_lst: Option<Box<CTColorSchemeList>>,
    #[serde(rename = "custClrLst")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub cust_clr_lst: Option<Box<CTCustomColorList>>,
    #[serde(rename = "extLst")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub ext_lst: Option<Box<CTOfficeArtExtensionList>>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
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
pub struct CTBaseStylesOverride {
    #[serde(rename = "clrScheme")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub clr_scheme: Option<Box<ColorScheme>>,
    #[serde(rename = "fontScheme")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub font_scheme: Option<Box<FontScheme>>,
    #[serde(rename = "fmtScheme")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub fmt_scheme: Option<Box<CTStyleMatrix>>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::RawXmlNode>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTClipboardStyleSheet {
    #[serde(rename = "themeElements")]
    pub theme_elements: Box<CTBaseStyles>,
    #[serde(rename = "clrMap")]
    pub clr_map: Box<CTColorMapping>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::RawXmlNode>,
}

pub type ATheme = Box<CTOfficeStyleSheet>;

pub type AThemeOverride = Box<CTBaseStylesOverride>;

pub type AThemeManager = Box<CTEmptyElement>;

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTTableCellProperties {
    #[serde(rename = "@marL")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub mar_l: Option<STCoordinate32>,
    #[serde(rename = "@marR")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub mar_r: Option<STCoordinate32>,
    #[serde(rename = "@marT")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub mar_t: Option<STCoordinate32>,
    #[serde(rename = "@marB")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub mar_b: Option<STCoordinate32>,
    #[serde(rename = "@vert")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub vert: Option<STTextVerticalType>,
    #[serde(rename = "@anchor")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub anchor: Option<STTextAnchoringType>,
    #[serde(rename = "@anchorCtr")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub anchor_ctr: Option<bool>,
    #[serde(rename = "@horzOverflow")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub horz_overflow: Option<STTextHorzOverflowType>,
    #[serde(rename = "lnL")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub ln_l: Option<Box<LineProperties>>,
    #[serde(rename = "lnR")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub ln_r: Option<Box<LineProperties>>,
    #[serde(rename = "lnT")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub ln_t: Option<Box<LineProperties>>,
    #[serde(rename = "lnB")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub ln_b: Option<Box<LineProperties>>,
    #[serde(rename = "lnTlToBr")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub ln_tl_to_br: Option<Box<LineProperties>>,
    #[serde(rename = "lnBlToTr")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub ln_bl_to_tr: Option<Box<LineProperties>>,
    #[serde(rename = "cell3D")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub cell3_d: Option<Box<CTCell3D>>,
    #[serde(rename = "headers")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub headers: Option<Box<CTHeaders>>,
    #[serde(rename = "extLst")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub ext_lst: Option<Box<CTOfficeArtExtensionList>>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
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
pub struct CTHeaders {
    #[serde(rename = "header")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub header: Vec<String>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::RawXmlNode>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTTableCol {
    #[serde(rename = "@w")]
    pub width: STCoordinate,
    #[serde(rename = "extLst")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub ext_lst: Option<Box<CTOfficeArtExtensionList>>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
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
pub struct CTTableGrid {
    #[serde(rename = "gridCol")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub grid_col: Vec<Box<CTTableCol>>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::RawXmlNode>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTTableCell {
    #[serde(rename = "@rowSpan")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub row_span: Option<i32>,
    #[serde(rename = "@gridSpan")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub grid_span: Option<i32>,
    #[serde(rename = "@hMerge")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub h_merge: Option<bool>,
    #[serde(rename = "@vMerge")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub v_merge: Option<bool>,
    #[serde(rename = "@id")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub id: Option<String>,
    #[serde(rename = "txBody")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub tx_body: Option<Box<TextBody>>,
    #[serde(rename = "tcPr")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub tc_pr: Option<Box<CTTableCellProperties>>,
    #[serde(rename = "extLst")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub ext_lst: Option<Box<CTOfficeArtExtensionList>>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
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
pub struct CTTableRow {
    #[serde(rename = "@h")]
    pub height: STCoordinate,
    #[serde(rename = "tc")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub tc: Vec<Box<CTTableCell>>,
    #[serde(rename = "extLst")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub ext_lst: Option<Box<CTOfficeArtExtensionList>>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
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
pub struct CTTableProperties {
    #[serde(rename = "@rtl")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub rtl: Option<bool>,
    #[serde(rename = "@firstRow")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub first_row: Option<bool>,
    #[serde(rename = "@firstCol")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub first_col: Option<bool>,
    #[serde(rename = "@lastRow")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub last_row: Option<bool>,
    #[serde(rename = "@lastCol")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub last_col: Option<bool>,
    #[serde(rename = "@bandRow")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub band_row: Option<bool>,
    #[serde(rename = "@bandCol")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub band_col: Option<bool>,
    #[serde(rename = "extLst")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub ext_lst: Option<Box<CTOfficeArtExtensionList>>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
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
pub struct CTTable {
    #[serde(rename = "tblPr")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub tbl_pr: Option<Box<CTTableProperties>>,
    #[serde(rename = "tblGrid")]
    pub tbl_grid: Box<CTTableGrid>,
    #[serde(rename = "tr")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub tr: Vec<Box<CTTableRow>>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::RawXmlNode>,
}

pub type ATbl = Box<CTTable>;

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTCell3D {
    #[serde(rename = "@prstMaterial")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub prst_material: Option<STPresetMaterialType>,
    #[serde(rename = "bevel")]
    pub bevel: Box<CTBevel>,
    #[serde(rename = "lightRig")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub light_rig: Option<Box<CTLightRig>>,
    #[serde(rename = "extLst")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub ext_lst: Option<Box<CTOfficeArtExtensionList>>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
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
pub struct CTThemeableLineStyle;

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTTableStyleTextStyle {
    #[serde(rename = "@b")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub b: Option<STOnOffStyleType>,
    #[serde(rename = "@i")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub i: Option<STOnOffStyleType>,
    #[serde(rename = "extLst")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub ext_lst: Option<Box<CTOfficeArtExtensionList>>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
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
pub struct CTTableCellBorderStyle {
    #[serde(rename = "left")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub left: Option<Box<CTThemeableLineStyle>>,
    #[serde(rename = "right")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub right: Option<Box<CTThemeableLineStyle>>,
    #[serde(rename = "top")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub top: Option<Box<CTThemeableLineStyle>>,
    #[serde(rename = "bottom")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub bottom: Option<Box<CTThemeableLineStyle>>,
    #[serde(rename = "insideH")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub inside_h: Option<Box<CTThemeableLineStyle>>,
    #[serde(rename = "insideV")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub inside_v: Option<Box<CTThemeableLineStyle>>,
    #[serde(rename = "tl2br")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub tl2br: Option<Box<CTThemeableLineStyle>>,
    #[serde(rename = "tr2bl")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub tr2bl: Option<Box<CTThemeableLineStyle>>,
    #[serde(rename = "extLst")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub ext_lst: Option<Box<CTOfficeArtExtensionList>>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::RawXmlNode>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct CTTableBackgroundStyle;

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTTableStyleCellStyle {
    #[serde(rename = "tcBdr")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub tc_bdr: Option<Box<CTTableCellBorderStyle>>,
    #[serde(rename = "cell3D")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub cell3_d: Option<Box<CTCell3D>>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::RawXmlNode>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTTablePartStyle {
    #[serde(rename = "tcTxStyle")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub tc_tx_style: Option<Box<CTTableStyleTextStyle>>,
    #[serde(rename = "tcStyle")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub tc_style: Option<Box<CTTableStyleCellStyle>>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::RawXmlNode>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTTableStyle {
    #[serde(rename = "@styleId")]
    pub style_id: Guid,
    #[serde(rename = "@styleName")]
    pub style_name: String,
    #[serde(rename = "tblBg")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub tbl_bg: Option<Box<CTTableBackgroundStyle>>,
    #[serde(rename = "wholeTbl")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub whole_tbl: Option<Box<CTTablePartStyle>>,
    #[serde(rename = "band1H")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub band1_h: Option<Box<CTTablePartStyle>>,
    #[serde(rename = "band2H")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub band2_h: Option<Box<CTTablePartStyle>>,
    #[serde(rename = "band1V")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub band1_v: Option<Box<CTTablePartStyle>>,
    #[serde(rename = "band2V")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub band2_v: Option<Box<CTTablePartStyle>>,
    #[serde(rename = "lastCol")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub last_col: Option<Box<CTTablePartStyle>>,
    #[serde(rename = "firstCol")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub first_col: Option<Box<CTTablePartStyle>>,
    #[serde(rename = "lastRow")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub last_row: Option<Box<CTTablePartStyle>>,
    #[serde(rename = "seCell")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub se_cell: Option<Box<CTTablePartStyle>>,
    #[serde(rename = "swCell")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub sw_cell: Option<Box<CTTablePartStyle>>,
    #[serde(rename = "firstRow")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub first_row: Option<Box<CTTablePartStyle>>,
    #[serde(rename = "neCell")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub ne_cell: Option<Box<CTTablePartStyle>>,
    #[serde(rename = "nwCell")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub nw_cell: Option<Box<CTTablePartStyle>>,
    #[serde(rename = "extLst")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub ext_lst: Option<Box<CTOfficeArtExtensionList>>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
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
pub struct CTTableStyleList {
    #[serde(rename = "@def")]
    pub def: Guid,
    #[serde(rename = "tblStyle")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub tbl_style: Vec<Box<CTTableStyle>>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
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

pub type ATblStyleLst = Box<CTTableStyleList>;

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct TextParagraph {
    #[serde(rename = "pPr")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub p_pr: Option<Box<TextParagraphProperties>>,
    #[serde(rename = "endParaRPr")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub end_para_r_pr: Option<Box<TextCharacterProperties>>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::RawXmlNode>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTTextListStyle {
    #[serde(rename = "defPPr")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub def_p_pr: Option<Box<TextParagraphProperties>>,
    #[serde(rename = "lvl1pPr")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub lvl1p_pr: Option<Box<TextParagraphProperties>>,
    #[serde(rename = "lvl2pPr")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub lvl2p_pr: Option<Box<TextParagraphProperties>>,
    #[serde(rename = "lvl3pPr")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub lvl3p_pr: Option<Box<TextParagraphProperties>>,
    #[serde(rename = "lvl4pPr")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub lvl4p_pr: Option<Box<TextParagraphProperties>>,
    #[serde(rename = "lvl5pPr")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub lvl5p_pr: Option<Box<TextParagraphProperties>>,
    #[serde(rename = "lvl6pPr")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub lvl6p_pr: Option<Box<TextParagraphProperties>>,
    #[serde(rename = "lvl7pPr")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub lvl7p_pr: Option<Box<TextParagraphProperties>>,
    #[serde(rename = "lvl8pPr")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub lvl8p_pr: Option<Box<TextParagraphProperties>>,
    #[serde(rename = "lvl9pPr")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub lvl9p_pr: Option<Box<TextParagraphProperties>>,
    #[serde(rename = "extLst")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub ext_lst: Option<Box<CTOfficeArtExtensionList>>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::RawXmlNode>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTTextNormalAutofit {
    #[serde(rename = "@fontScale")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub font_scale: Option<STTextFontScalePercentOrPercentString>,
    #[serde(rename = "@lnSpcReduction")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub ln_spc_reduction: Option<STTextSpacingPercentOrPercentString>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct CTTextShapeAutofit;

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct CTTextNoAutofit;

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTTextBodyProperties {
    #[serde(rename = "@rot")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub rot: Option<STAngle>,
    #[serde(rename = "@spcFirstLastPara")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub spc_first_last_para: Option<bool>,
    #[serde(rename = "@vertOverflow")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub vert_overflow: Option<STTextVertOverflowType>,
    #[serde(rename = "@horzOverflow")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub horz_overflow: Option<STTextHorzOverflowType>,
    #[serde(rename = "@vert")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub vert: Option<STTextVerticalType>,
    #[serde(rename = "@wrap")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub wrap: Option<STTextWrappingType>,
    #[serde(rename = "@lIns")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub l_ins: Option<STCoordinate32>,
    #[serde(rename = "@tIns")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub t_ins: Option<STCoordinate32>,
    #[serde(rename = "@rIns")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub r_ins: Option<STCoordinate32>,
    #[serde(rename = "@bIns")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub b_ins: Option<STCoordinate32>,
    #[serde(rename = "@numCol")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub num_col: Option<STTextColumnCount>,
    #[serde(rename = "@spcCol")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub spc_col: Option<STPositiveCoordinate32>,
    #[serde(rename = "@rtlCol")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub rtl_col: Option<bool>,
    #[serde(rename = "@fromWordArt")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub from_word_art: Option<bool>,
    #[serde(rename = "@anchor")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub anchor: Option<STTextAnchoringType>,
    #[serde(rename = "@anchorCtr")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub anchor_ctr: Option<bool>,
    #[serde(rename = "@forceAA")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub force_a_a: Option<bool>,
    #[serde(rename = "@upright")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub upright: Option<bool>,
    #[serde(rename = "@compatLnSpc")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub compat_ln_spc: Option<bool>,
    #[serde(rename = "prstTxWarp")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub prst_tx_warp: Option<Box<CTPresetTextShape>>,
    #[serde(rename = "scene3d")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub scene3d: Option<Box<CTScene3D>>,
    #[serde(rename = "extLst")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub ext_lst: Option<Box<CTOfficeArtExtensionList>>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
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
pub struct TextBody {
    #[serde(rename = "bodyPr")]
    pub body_pr: Box<CTTextBodyProperties>,
    #[serde(rename = "lstStyle")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub lst_style: Option<Box<CTTextListStyle>>,
    #[serde(rename = "p")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub p: Vec<Box<TextParagraph>>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::RawXmlNode>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct CTTextBulletColorFollowText;

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct CTTextBulletSizeFollowText;

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTTextBulletSizePercent {
    #[serde(rename = "@val")]
    pub value: STTextBulletSizePercent,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTTextBulletSizePoint {
    #[serde(rename = "@val")]
    pub value: STTextFontSize,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct CTTextBulletTypefaceFollowText;

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTTextAutonumberBullet {
    #[serde(rename = "@type")]
    pub r#type: STTextAutonumberScheme,
    #[serde(rename = "@startAt")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub start_at: Option<STTextBulletStartAtNum>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTTextCharBullet {
    #[serde(rename = "@char")]
    pub char: String,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

pub type CTTextBlipBullet = Box<Blip>;

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct CTTextNoBullet;

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct TextFont {
    #[serde(rename = "@typeface")]
    pub typeface: STTextTypeface,
    #[serde(rename = "@panose")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub panose: Option<Panose>,
    #[serde(rename = "@pitchFamily")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub pitch_family: Option<STPitchFamily>,
    #[serde(rename = "@charset")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub charset: Option<i8>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct CTTextUnderlineLineFollowText;

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct CTTextUnderlineFillFollowText;

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct CTTextUnderlineFillGroupWrapper;

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct TextCharacterProperties {
    #[serde(rename = "@kumimoji")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub kumimoji: Option<bool>,
    #[serde(rename = "@lang")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub lang: Option<Language>,
    #[serde(rename = "@altLang")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub alt_lang: Option<Language>,
    #[serde(rename = "@sz")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub sz: Option<STTextFontSize>,
    #[serde(rename = "@b")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub b: Option<bool>,
    #[serde(rename = "@i")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub i: Option<bool>,
    #[serde(rename = "@u")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub u: Option<STTextUnderlineType>,
    #[serde(rename = "@strike")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub strike: Option<STTextStrikeType>,
    #[serde(rename = "@kern")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub kern: Option<STTextNonNegativePoint>,
    #[serde(rename = "@cap")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub cap: Option<STTextCapsType>,
    #[serde(rename = "@spc")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub spc: Option<STTextPoint>,
    #[serde(rename = "@normalizeH")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub normalize_h: Option<bool>,
    #[serde(rename = "@baseline")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub baseline: Option<STPercentage>,
    #[serde(rename = "@noProof")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub no_proof: Option<bool>,
    #[serde(rename = "@dirty")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub dirty: Option<bool>,
    #[serde(rename = "@err")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub err: Option<bool>,
    #[serde(rename = "@smtClean")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub smt_clean: Option<bool>,
    #[serde(rename = "@smtId")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub smt_id: Option<u32>,
    #[serde(rename = "@bmk")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub bmk: Option<String>,
    #[serde(rename = "ln")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub line: Option<Box<LineProperties>>,
    #[serde(rename = "highlight")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub highlight: Option<Box<CTColor>>,
    #[serde(rename = "latin")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub latin: Option<Box<TextFont>>,
    #[serde(rename = "ea")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub ea: Option<Box<TextFont>>,
    #[serde(rename = "cs")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub cs: Option<Box<TextFont>>,
    #[serde(rename = "sym")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub sym: Option<Box<TextFont>>,
    #[serde(rename = "hlinkClick")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub hlink_click: Option<Box<CTHyperlink>>,
    #[serde(rename = "hlinkMouseOver")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub hlink_mouse_over: Option<Box<CTHyperlink>>,
    #[serde(rename = "rtl")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub rtl: Option<Box<CTBoolean>>,
    #[serde(rename = "extLst")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub ext_lst: Option<Box<CTOfficeArtExtensionList>>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
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
pub struct CTBoolean {
    #[serde(rename = "@val")]
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
pub struct CTTextSpacingPercent {
    #[serde(rename = "@val")]
    pub value: STTextSpacingPercentOrPercentString,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTTextSpacingPoint {
    #[serde(rename = "@val")]
    pub value: STTextSpacingPoint,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTTextTabStop {
    #[serde(rename = "@pos")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub pos: Option<STCoordinate32>,
    #[serde(rename = "@algn")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub algn: Option<STTextTabAlignType>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
    #[serde(skip)]
    #[cfg(feature = "extra-attrs")]
    #[serde(default)]
    #[cfg(feature = "extra-attrs")]
    pub extra_attrs: std::collections::HashMap<String, String>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTTextTabStopList {
    #[serde(rename = "tab")]
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub tab: Vec<Box<CTTextTabStop>>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::RawXmlNode>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CTTextLineBreak {
    #[serde(rename = "rPr")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub r_pr: Option<Box<TextCharacterProperties>>,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::RawXmlNode>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct CTTextSpacing;

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct TextParagraphProperties {
    #[serde(rename = "@marL")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub mar_l: Option<STTextMargin>,
    #[serde(rename = "@marR")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub mar_r: Option<STTextMargin>,
    #[serde(rename = "@lvl")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub lvl: Option<STTextIndentLevelType>,
    #[serde(rename = "@indent")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub indent: Option<STTextIndent>,
    #[serde(rename = "@algn")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub algn: Option<STTextAlignType>,
    #[serde(rename = "@defTabSz")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub def_tab_sz: Option<STCoordinate32>,
    #[serde(rename = "@rtl")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub rtl: Option<bool>,
    #[serde(rename = "@eaLnBrk")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub ea_ln_brk: Option<bool>,
    #[serde(rename = "@fontAlgn")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub font_algn: Option<STTextFontAlignType>,
    #[serde(rename = "@latinLnBrk")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub latin_ln_brk: Option<bool>,
    #[serde(rename = "@hangingPunct")]
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        with = "ooxml_xml::ooxml_bool"
    )]
    pub hanging_punct: Option<bool>,
    #[serde(rename = "lnSpc")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub ln_spc: Option<Box<CTTextSpacing>>,
    #[serde(rename = "spcBef")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub spc_bef: Option<Box<CTTextSpacing>>,
    #[serde(rename = "spcAft")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub spc_aft: Option<Box<CTTextSpacing>>,
    #[serde(rename = "tabLst")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub tab_lst: Option<Box<CTTextTabStopList>>,
    #[serde(rename = "defRPr")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub def_r_pr: Option<Box<TextCharacterProperties>>,
    #[serde(rename = "extLst")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub ext_lst: Option<Box<CTOfficeArtExtensionList>>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
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
pub struct CTTextField {
    #[serde(rename = "@id")]
    pub id: Guid,
    #[serde(rename = "@type")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub r#type: Option<String>,
    #[serde(rename = "rPr")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub r_pr: Option<Box<TextCharacterProperties>>,
    #[serde(rename = "pPr")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub p_pr: Option<Box<TextParagraphProperties>>,
    #[serde(rename = "t")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub t: Option<String>,
    /// Unknown attributes captured for roundtrip fidelity.
    #[cfg(feature = "extra-attrs")]
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
pub struct TextRun {
    #[serde(rename = "rPr")]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub r_pr: Option<Box<TextCharacterProperties>>,
    #[serde(rename = "t")]
    pub t: String,
    /// Unknown child elements captured for roundtrip fidelity.
    #[cfg(feature = "extra-children")]
    #[serde(skip)]
    #[cfg(feature = "extra-children")]
    pub extra_children: Vec<ooxml_xml::RawXmlNode>,
}
